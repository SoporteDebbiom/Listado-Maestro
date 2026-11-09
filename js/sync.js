
var SYNC = {
  initialized:  false,
  connected:    false,
  onlineCount:  0,
  onlineUsers:  [],
  sessionId:    Date.now().toString(36) + Math.random().toString(36).substr(2, 6),
  _version:     0,
  _isSaving:    false,
  _saveTimer:   null,
  _saveQueue:   false,
  _client:      null,
  _channel:     null,
  _presence:    null,
  _lastSaveAt:  0,
  _initialLoadDone: false,
  _heartbeatTimer: null,
  _vigenciaTimer:  null,
  _backupTimer:    null,
  _lastBackupAt:   0,
  _lastSyncTime:   0,
  _lastKnownRecordCount: 0
};

async function _deriveKey(password, salt) {
  var enc = new TextEncoder();
  var keyMaterial = await crypto.subtle.importKey('raw', enc.encode(password), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({name:'PBKDF2', salt:salt, iterations:100000, hash:'SHA-256'}, keyMaterial, {name:'AES-GCM', length:256}, false, ['encrypt','decrypt']);
}
async function _encryptBackup(jsonStr, password) {
  var salt = crypto.getRandomValues(new Uint8Array(16));
  var iv = crypto.getRandomValues(new Uint8Array(12));
  var key = await _deriveKey(password, salt);
  var enc = new TextEncoder();
  var ct = await crypto.subtle.encrypt({name:'AES-GCM', iv:iv}, key, enc.encode(jsonStr));
  var result = new Uint8Array(16 + 12 + ct.byteLength);
  result.set(salt, 0); result.set(iv, 16); result.set(new Uint8Array(ct), 28);
  return result;
}
async function _decryptBackup(buffer, password) {
  var data = new Uint8Array(buffer);
  var salt = data.slice(0, 16);
  var iv = data.slice(16, 28);
  var ct = data.slice(28);
  var key = await _deriveKey(password, salt);
  var pt = await crypto.subtle.decrypt({name:'AES-GCM', iv:iv}, key, ct);
  return new TextDecoder().decode(pt);
}

var _MAX_RECORDS = 50000;

function _validateStateData(data) {
  if (!data || typeof data !== 'object') return false;
  var arrays = ['records','obsoletos','papelera','salidas','logs'];
  for (var i = 0; i < arrays.length; i++) {
    var k = arrays[i];
    if (data[k] && (!Array.isArray(data[k]) || data[k].length > _MAX_RECORDS)) return false;
  }
  if (data.records) {
    for (var j = 0; j < data.records.length; j++) {
      var r = data.records[j];
      if (!r || typeof r !== 'object' || !r.id) return false;
    }
  }
  if (typeof data.nextId !== 'undefined' && (typeof data.nextId !== 'number' || data.nextId < 0)) return false;
  return true;
}

function _sanitizeIncoming(stateObj) {
  if (!stateObj || typeof stateObj !== 'object') return null;
  var clean = {};
  var arrays = ['records','obsoletos','papelera','salidas','logs'];
  for (var i = 0; i < arrays.length; i++) {
    var k = arrays[i];
    if (Array.isArray(stateObj[k])) {
      clean[k] = stateObj[k].slice(0, _MAX_RECORDS).filter(function(item) {
        return item && typeof item === 'object';
      });
    } else {
      clean[k] = [];
    }
  }
  clean.elaboros = Array.isArray(stateObj.elaboros) ? stateObj.elaboros.filter(function(e) { return typeof e === 'string'; }).slice(0, 1000) : [];
  clean.nextId = (typeof stateObj.nextId === 'number' && stateObj.nextId >= 0) ? stateObj.nextId : 1000;
  return clean;
}

var SAVE_DEBOUNCE = 1500;
var HEARTBEAT_INTERVAL = 30000;
var VIGENCIA_CHECK_INTERVAL = 60000;
var AUTO_BACKUP_INTERVAL = 3600000;
async function initSync() {
  var client = typeof getDB === 'function' ? getDB() : null;
  if (!client) {
    return false;
  }

  try {
    SYNC._client = client;

    var result = await SYNC._client
      .from('app_state')
      .select('version')
      .eq('id', 1)
      .single();

    if (result.error) {
      return false;
    }

    SYNC._version    = result.data.version || 0;
    SYNC.initialized = true;
    SYNC.connected   = true;

    _startHeartbeat();
    _startVigenciaCheck();
    _startAutoBackup();

    updateSyncIndicator();
    return true;

  } catch (e) {
    return false;
  }
}
function syncSaveState() {
  if (!SYNC.initialized || !SYNC._client) return;
  if (!SYNC._initialLoadDone) {
    return;
  }

  clearTimeout(SYNC._saveTimer);
  SYNC._saveTimer = setTimeout(_doSave, SAVE_DEBOUNCE);
}

var _retryCount = 0;
var _maxRetries = 5;
var _retryTimer = null;

function _retryDelay(attempt) {
  return Math.min(1000 * Math.pow(2, attempt), 30000) + Math.random() * 500;
}

async function _doSave() {
  if (SYNC._isSaving) {
    SYNC._saveQueue = true;
    return;
  }
  if (typeof getCsrfToken === 'function' && !getCsrfToken()) {
    return;
  }

  SYNC._isSaving = true;
  _showSyncAnimation(true);

  try {
    var newVersion = Date.now();
    var stateData = {
      records:   STATE.records || [],
      obsoletos: STATE.obsoletos || [],
      papelera:  STATE.papelera || [],
      salidas:   STATE.salidas || [],
      logs:      STATE.logs || [],
      elaboros:  STATE.elaboros || [],
      nextId:    STATE.nextId || 1000
    };

    if (!_validateStateData(stateData)) {
      showToast('Error interno: datos inválidos, no se guardó', 'error');
      SYNC._isSaving = false;
      _showSyncAnimation(false);
      return;
    }

    // Prevent accidental data wipe: don't save empty state if we previously had records
    if (stateData.records.length === 0 && SYNC._lastKnownRecordCount > 0) {
      showToast('Protección: no se puede guardar un estado vacío. Usa "Sincronizar" para recargar.', 'error');
      SYNC._isSaving = false;
      _showSyncAnimation(false);
      return;
    }
    SYNC._lastKnownRecordCount = stateData.records.length;

    if (typeof _dataHash === 'function') stateData._hash = _dataHash(stateData);

    var modifiedBy = AUTH.currentUser ? AUTH.currentUser.username : 'Sistema';

    // Optimistic concurrency: only update if version matches what we last saw
    var result = await SYNC._client
      .from('app_state')
      .update({
        state:       stateData,
        version:     newVersion,
        modified_by: modifiedBy,
        session_id:  SYNC.sessionId,
        updated_at:  new Date().toISOString()
      })
      .eq('id', 1)
      .eq('version', SYNC._version)
      .select('version')
      .single();

    if (result.error) {
      // Version conflict or connection error — pull remote and merge
      var pullResult = await SYNC._client
        .from('app_state')
        .select('*')
        .eq('id', 1)
        .single();

      if (pullResult.error || !pullResult.data) {
        SYNC.connected = false;
        _scheduleRetry();
      } else if (pullResult.data.version !== SYNC._version) {
        // Conflict: merge remote into local, then retry save
        if (pullResult.data.state) {
          STATE.records   = _mergeById(STATE.records,   pullResult.data.state.records   || []);
          STATE.obsoletos = _mergeById(STATE.obsoletos, pullResult.data.state.obsoletos || []);
          STATE.papelera  = _mergeById(STATE.papelera,  pullResult.data.state.papelera  || []);
          STATE.salidas   = _mergeById(STATE.salidas,   pullResult.data.state.salidas   || []);
          STATE.logs      = (pullResult.data.state.logs || []).length >= (STATE.logs || []).length ? pullResult.data.state.logs : STATE.logs;
          STATE.elaboros  = _mergeArrays(STATE.elaboros, pullResult.data.state.elaboros || []);
          STATE.nextId    = Math.max(STATE.nextId || 1000, pullResult.data.state.nextId || 1000);
        }
        SYNC._version = pullResult.data.version;
        if (typeof render === 'function') render();
        showToast('Datos fusionados con cambios de otro usuario, reintentando...', 'info');
        SYNC._isSaving = false;
        setTimeout(_doSave, 500);
        _showSyncAnimation(false);
        return;
      } else {
        // Same version but error — connection issue
        SYNC.connected = false;
        _scheduleRetry();
      }
    } else {
      SYNC._version    = newVersion;
      SYNC._lastSaveAt = Date.now();
      SYNC._lastSyncTime = Date.now();
      SYNC.connected   = true;
      _retryCount = 0;
      clearTimeout(_retryTimer);
      if (typeof _unsavedChanges !== 'undefined') _unsavedChanges = false;
      _updateLastSyncUI();
      _checkStateSize(stateData);
    }
  } catch (e) {
    SYNC.connected = false;
    _scheduleRetry();
  }

  updateSyncIndicator();
  _showSyncAnimation(false);
  SYNC._isSaving = false;

  if (SYNC._saveQueue) {
    SYNC._saveQueue = false;
    setTimeout(_doSave, 300);
  }
}

function _scheduleRetry() {
  if (_retryCount >= _maxRetries) {
    showToast('No se pudieron guardar los cambios tras ' + _maxRetries + ' intentos. Verifica tu conexion.', 'error');
    _retryCount = 0;
    return;
  }
  _retryCount++;
  var delay = _retryDelay(_retryCount);
  clearTimeout(_retryTimer);
  _retryTimer = setTimeout(function() {
    if (SYNC.initialized && SYNC._client && AUTH.currentUser) {
      _doSave();
    }
  }, delay);
}

var _STATE_WARN_MB = 3;
var _STATE_CRITICAL_MB = 4.5;
var _lastSizeWarnAt = 0;
function _checkStateSize(data) {
  try {
    var json = JSON.stringify(data);
    var sizeMB = json.length / (1024 * 1024);
    var now = Date.now();
    if (now - _lastSizeWarnAt < 300000) return;
    if (sizeMB >= _STATE_CRITICAL_MB) {
      _lastSizeWarnAt = now;
      if (typeof isAdmin === 'function' && isAdmin()) {
        showToast('Estado critico: ' + sizeMB.toFixed(1) + ' MB. Considera archivar registros obsoletos.', 'error');
      }
      if (typeof addLog === 'function') addLog('sizeAlert', 'Estado: ' + sizeMB.toFixed(2) + ' MB (critico)');
    } else if (sizeMB >= _STATE_WARN_MB) {
      _lastSizeWarnAt = now;
      if (typeof isAdmin === 'function' && isAdmin()) {
        showToast('Tamano del estado: ' + sizeMB.toFixed(1) + ' MB. Monitorear crecimiento.', 'info');
      }
      if (typeof addLog === 'function') addLog('sizeWarn', 'Estado: ' + sizeMB.toFixed(2) + ' MB');
    }
  } catch (e) {}
}

window.addEventListener('online', function() {
  if (SYNC.initialized && !SYNC.connected && AUTH.currentUser) {
    showToast('Conexion restaurada, sincronizando...', 'info');
    setTimeout(_doSave, 1000);
  }
});

window.addEventListener('offline', function() {
  if (SYNC.initialized) {
    SYNC.connected = false;
    updateSyncIndicator();
    showToast('Sin conexion. Los cambios se guardaran al reconectar.', 'error');
  }
});

async function forceSync() {
  if (!SYNC.initialized) {
    showToast('Sincronización no disponible (modo local)', 'error');
    return;
  }
  _showSyncAnimation(true);

  await _doSave();
  await _pullRemote(true);

  _showSyncAnimation(false);
  showToast('✅ Sincronizado', 'success');
}
function listenStateChanges() {
  if (!SYNC.initialized || !SYNC._client) return;

  if (SYNC._channel) {
    SYNC._client.removeChannel(SYNC._channel);
    SYNC._channel = null;
  }

  SYNC._channel = SYNC._client
    .channel('db-changes')
    .on('postgres_changes',
      { event: 'UPDATE', schema: 'public', table: 'app_state' },
      function(payload) {
        _handleRealtimeUpdate(payload.new, false);
      }
    )
    .subscribe(function(status) {
      if (status === 'SUBSCRIBED') {
        SYNC.connected = true;
        updateSyncIndicator();
      } else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') {
        SYNC.connected = false;
        updateSyncIndicator();
        setTimeout(function() {
          if (SYNC.initialized && !SYNC.connected) {
            listenStateChanges();
          }
        }, 5000);
      }
    });
}

function _mergeById(local, remote) {
  if (!local || local.length === 0) return remote || [];
  if (!remote || remote.length === 0) return local || [];
  var localMap = {};
  local.forEach(function(r) { if (r && r.id) localMap[r.id] = r; });
  var remoteMap = {};
  remote.forEach(function(r) { if (r && r.id) remoteMap[r.id] = r; });
  var merged = {};
  var conflicts = [];
  remote.forEach(function(r) { if (r && r.id) merged[r.id] = r; });
  local.forEach(function(r) {
    if (!r || !r.id) return;
    if (!merged[r.id]) { merged[r.id] = r; return; }
    var rem = merged[r.id];
    var localTs = r._history && r._history.length > 0 ? r._history[r._history.length-1].ts : null;
    var remoteTs = rem._history && rem._history.length > 0 ? rem._history[rem._history.length-1].ts : null;
    if (localTs && remoteTs && localTs !== remoteTs) {
      var result = _mergeFields(r, rem);
      merged[r.id] = result.record;
      if (result.conflictFields.length > 0) {
        conflicts.push({id: r.id, nombre: r.nombreDoc || rem.nombreDoc || r.id, fields: result.conflictFields});
      }
    } else if (!remoteTs && localTs) {
      merged[r.id] = r;
    }
  });
  if (conflicts.length > 0) {
    var msg = conflicts.length + ' registro(s) con cambios simultaneos. Se conservo la version mas reciente de cada campo: ';
    msg += conflicts.map(function(c) { return '"' + c.nombre + '" (' + c.fields.join(', ') + ')'; }).join('; ');
    if (typeof showToast === 'function') showToast(msg, 'info');
  }
  return Object.values(merged);
}

function _mergeFields(local, remote) {
  var result = {};
  var conflictFields = [];
  var allKeys = {};
  Object.keys(local).forEach(function(k) { allKeys[k] = true; });
  Object.keys(remote).forEach(function(k) { allKeys[k] = true; });
  var localTs = local._history && local._history.length > 0 ? new Date(local._history[local._history.length-1].ts).getTime() : 0;
  var remoteTs = remote._history && remote._history.length > 0 ? new Date(remote._history[remote._history.length-1].ts).getTime() : 0;
  for (var k in allKeys) {
    if (k === '_history') {
      var hist = (local._history || []).concat(remote._history || []);
      hist.sort(function(a, b) { return new Date(a.ts) - new Date(b.ts); });
      var seen = {};
      result._history = hist.filter(function(h) {
        var key = h.ts + h.user + h.action;
        if (seen[key]) return false;
        seen[key] = true;
        return true;
      });
      if (result._history.length > 50) result._history = result._history.slice(-50);
      continue;
    }
    var lv = JSON.stringify(local[k]);
    var rv = JSON.stringify(remote[k]);
    if (lv === rv) {
      result[k] = local[k];
    } else {
      if (localTs >= remoteTs) { result[k] = local[k]; } else { result[k] = remote[k]; }
      if (k !== 'id' && k !== '_hash') conflictFields.push(k);
    }
  }
  return { record: result, conflictFields: conflictFields };
}

function _mergeArrays(local, remote) {
  var set = {};
  (local || []).forEach(function(v) { set[v] = true; });
  (remote || []).forEach(function(v) { set[v] = true; });
  return Object.keys(set);
}

function _handleRealtimeUpdate(row, forceApply) {
  if (!row) return;

  var remoteVersion = row.version || 0;

  if (!forceApply) {
    var remoteSessionId = row.session_id || '';
    if (remoteSessionId && remoteSessionId === SYNC.sessionId) {
      return;
    }
    if (!remoteSessionId) {
      var myUser = AUTH.currentUser ? AUTH.currentUser.username : '';
      var remoteUser = row.modified_by || '';
      if (myUser && remoteUser === myUser && (Date.now() - SYNC._lastSaveAt < 5000)) {
        return;
      }
    }
  }

  if (!forceApply && remoteVersion <= SYNC._version) return;

  var by = row.modified_by || '?';

  if (row.state) {
    var incoming = _sanitizeIncoming(row.state);
    if (!incoming) return;
    STATE.records   = _mergeById(STATE.records,   incoming.records);
    STATE.obsoletos = _mergeById(STATE.obsoletos, incoming.obsoletos);
    STATE.papelera  = _mergeById(STATE.papelera,  incoming.papelera);
    STATE.salidas   = _mergeById(STATE.salidas,   incoming.salidas);
    STATE.logs      = incoming.logs.length >= (STATE.logs || []).length ? incoming.logs : STATE.logs;
    STATE.elaboros  = _mergeArrays(STATE.elaboros, incoming.elaboros);
    STATE.nextId    = Math.max(STATE.nextId || 1000, incoming.nextId);
  }

  SYNC._version = remoteVersion;

  if (AUTH.currentUser) {
    render();
    if (by && by !== (AUTH.currentUser ? AUTH.currentUser.username : '')) {
      showToast('📥 ' + by + ' actualizó los datos', 'info');
      if (typeof _playNotificationSound === 'function') _playNotificationSound();
    }
    if (typeof _unsavedChanges !== 'undefined') _unsavedChanges = false;
    SYNC._lastSyncTime = Date.now();
    _updateLastSyncUI();
  }
}

async function _pullRemote(forceApply) {
  if (!SYNC._client) return;

  try {
    var result = await SYNC._client
      .from('app_state')
      .select('*')
      .eq('id', 1)
      .single();

    if (result.error || !result.data) return;

    var shouldApply = forceApply || (result.data.version > SYNC._version);
    if (shouldApply && result.data.state) {
      _handleRealtimeUpdate(result.data, !!forceApply);
    }
  } catch (e) {
  }
}
function _startHeartbeat() {
  clearInterval(SYNC._heartbeatTimer);
  SYNC._heartbeatTimer = setInterval(async function() {
    if (!SYNC.initialized || !SYNC._client) return;

    try {
      var result = await SYNC._client
        .from('app_state')
        .select('version')
        .eq('id', 1)
        .single();

      if (result.error) {
        if (SYNC.connected) {
          SYNC.connected = false;
          updateSyncIndicator();
        }
        return;
      }

      if (!SYNC.connected) {
        SYNC.connected = true;
        updateSyncIndicator();
        listenStateChanges();
        if (AUTH.currentUser) registerPresence();
        showToast('🔄 Reconectado a la nube', 'success');
      }

      var remoteVersion = result.data.version || 0;
      if (remoteVersion > SYNC._version) {
        await _pullRemote(false);
      }

    } catch (e) {
      if (SYNC.connected) {
        SYNC.connected = false;
        updateSyncIndicator();
      }
    }
  }, HEARTBEAT_INTERVAL);
}
function _startVigenciaCheck() {
  clearInterval(SYNC._vigenciaTimer);

  var prevVencidosCount = -1;

  SYNC._vigenciaTimer = setInterval(function() {
    if (!AUTH.currentUser) return;
    if (typeof calcEstado !== 'function') return;

    var currentVencidos = 0;
    (STATE.records || []).forEach(function(r) {
      if (calcEstado(r.fechaEmision) === 'Vencido') currentVencidos++;
    });

    if (prevVencidosCount >= 0 && currentVencidos !== prevVencidosCount) {
      render();
      if (currentVencidos > prevVencidosCount) {
        var nuevos = currentVencidos - prevVencidosCount;
        showToast('⚠️ ' + nuevos + ' documento(s) acaba(n) de vencer', 'error');
      }
    }
    prevVencidosCount = currentVencidos;
  }, VIGENCIA_CHECK_INTERVAL);
}
async function generateBackup() {
  if (!STATE.records || STATE.records.length === 0) {
    showToast('No hay datos para respaldar', 'error');
    return;
  }

  var now = new Date();
  var timestamp = now.getFullYear() + ''
    + String(now.getMonth()+1).padStart(2,'0')
    + String(now.getDate()).padStart(2,'0') + '_'
    + String(now.getHours()).padStart(2,'0')
    + String(now.getMinutes()).padStart(2,'0')
    + String(now.getSeconds()).padStart(2,'0');

  var backupData = {
    _meta: {
      tipo:      'DEBBIOM_BACKUP',
      version:   '2.0',
      fecha:     now.toISOString(),
      timestamp: timestamp,
      usuario:   AUTH.currentUser ? AUTH.currentUser.username : 'Sistema',
      syncVersion: SYNC._version,
      totalRegistros: (STATE.records || []).length,
      totalObsoletos: (STATE.obsoletos || []).length,
      totalPapelera:  (STATE.papelera || []).length,
      totalSalidas:   (STATE.salidas || []).length,
      totalLogs:      (STATE.logs || []).length
    },
    records:   STATE.records   || [],
    obsoletos: STATE.obsoletos || [],
    papelera:  STATE.papelera  || [],
    salidas:   STATE.salidas   || [],
    logs:      STATE.logs || [],
    elaboros:  STATE.elaboros  || [],
    nextId:    STATE.nextId    || 1000,
    users:     (AUTH.users || []).map(function(u) {
      return { username: u.username, role: u.role, full_name: u.full_name || u.fullName || '' };
    })
  };

  var json = JSON.stringify(backupData, null, 2);
  var useEncrypt = confirm('¿Deseas cifrar el backup con contraseña?\n(Recomendado para mayor seguridad)');
  var blob, ext;
  if (useEncrypt) {
    var pwd = prompt('Ingresa una contraseña para cifrar el backup:');
    if (!pwd || pwd.length < 4) { showToast('Contraseña muy corta (mín. 4 caracteres)', 'error'); return; }
    try {
      var encrypted = await _encryptBackup(json, pwd);
      blob = new Blob([encrypted], { type: 'application/octet-stream' });
      ext = '.dbk';
    } catch(e) { showToast('Error al cifrar backup', 'error'); return; }
  } else {
    blob = new Blob([json], { type: 'application/json' });
    ext = '.json';
  }
  var url  = URL.createObjectURL(blob);

  var a = document.createElement('a');
  a.href = url;
  a.download = 'DEBBIOM_BACKUP_' + timestamp + ext;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  SYNC._lastBackupAt = Date.now();
  addLog('backup', 'Backup ' + (useEncrypt ? 'cifrado' : 'JSON') + ' generado: DEBBIOM_BACKUP_' + timestamp + ext);
  showToast('Backup descargado' + (useEncrypt ? ' (cifrado)' : '') + ' + guardado en la nube', 'success');

  if (SYNC.initialized && SYNC._client) {
    _saveCloudBackup('json', 'Backup JSON descargado en ' + (AUTH.currentUser ? AUTH.currentUser.username : 'desconocido'));
  }
}

function generateBackupExcel() {
  if (typeof XLSX === 'undefined') {
    showToast('Error: Librería XLSX no disponible', 'error');
    return;
  }
  if (!STATE.records || STATE.records.length === 0) {
    showToast('No hay datos para respaldar', 'error');
    return;
  }

  var now = new Date();
  var timestamp = now.getFullYear() + ''
    + String(now.getMonth()+1).padStart(2,'0')
    + String(now.getDate()).padStart(2,'0') + '_'
    + String(now.getHours()).padStart(2,'0')
    + String(now.getMinutes()).padStart(2,'0');

  var wb = XLSX.utils.book_new();

  var recHeaders = ['ID','No.','Área','Cód. Manual','Cód. Instructivo','Cód. Listado','Cód. Formato','Nombre Documento','Versión','Fecha Emisión','Elaboró','Vigencia','Estado','Días Vigencia','Tipo Resguardo','Ubicación','Copias','Tipo Resguardo Copia','Usuarios','Observaciones','URLs Archivos','URLs Archivos Copia'];
  var recRows = (STATE.records || []).map(function(r) {
    var urls = (r.archivoURLs || []).join(' | ');
    var urlsCopia = (r.archivoURLsCopia || []).join(' | ');
    if (r.archivos && r.archivos.length > 0) {
      var nombres = r.archivos.map(function(f) { return f.name || ''; }).join(' | ');
      urls = urls ? urls + ' | [Adjuntos: ' + nombres + ']' : '[Adjuntos: ' + nombres + ']';
    }
    if (r.archivosCopia && r.archivosCopia.length > 0) {
      var nombresCopia = r.archivosCopia.map(function(f) { return f.name || ''; }).join(' | ');
      urlsCopia = urlsCopia ? urlsCopia + ' | [Adjuntos: ' + nombresCopia + ']' : '[Adjuntos: ' + nombresCopia + ']';
    }
    return [r.id, r.no, r.area, r.codManual, r.codInstructivo, r.codListado, r.codFormato, r.nombreDoc, r.version, r.fechaEmision, r.elaboro,
      typeof renderVigenciaDate === 'function' ? renderVigenciaDate(r.fechaEmision) : '',
      typeof calcEstado === 'function' ? calcEstado(r.fechaEmision) : '',
      typeof calcDiasVigencia === 'function' ? calcDiasVigencia(r.fechaEmision) : '',
      r.tipoResguardo, r.ubicacion, r.copias, r.tipoResguardoCopia, r.usuarios, r.observaciones,
      urls, urlsCopia
    ];
  });
  var ws1 = XLSX.utils.aoa_to_sheet([recHeaders].concat(recRows));
  XLSX.utils.book_append_sheet(wb, ws1, 'Registros');

  if ((STATE.obsoletos || []).length > 0) {
    var obsHeaders = ['ID','Fecha Obsoleto','No.','Área','Cód. Manual','Cód. Instructivo','Cód. Listado','Cód. Formato','Nombre Documento','Versión','Fecha Emisión','Elaboró','Motivo','Observaciones','URLs Archivos'];
    var obsRows = STATE.obsoletos.map(function(r) {
      var urls = (r.archivoURLs || []).join(' | ');
      return [r.id, r.fechaObsoleto || '', r.no, r.area, r.codManual, r.codInstructivo, r.codListado, r.codFormato, r.nombreDoc, r.version, r.fechaEmision, r.elaboro, r.motivo, r.observaciones, urls];
    });
    var ws2 = XLSX.utils.aoa_to_sheet([obsHeaders].concat(obsRows));
    XLSX.utils.book_append_sheet(wb, ws2, 'Obsoletos');
  }

  if ((STATE.salidas || []).length > 0) {
    var salHeaders = ['Nombre Documento','Versión','Fecha Emisión','Área','Motivo','Destinatario','Correo','Fecha Salida'];
    var salRows = STATE.salidas.map(function(r) {
      return [r.nombreDoc, r.version, r.fechaEmision, r.area, r.motivo, r.destinatario, r.correo, r.fechaSalida];
    });
    var ws3 = XLSX.utils.aoa_to_sheet([salHeaders].concat(salRows));
    XLSX.utils.book_append_sheet(wb, ws3, 'Salidas');
  }

  if ((STATE.papelera || []).length > 0) {
    var papHeaders = ['ID','No.','Nombre Documento','Área','Cód. Manual','Cód. Instructivo','Cód. Listado','Cód. Formato','Versión','Fecha Emisión','Elaboró','Observaciones','Fecha Eliminación'];
    var papRows = STATE.papelera.map(function(r) {
      return [r.id, r.no, r.nombreDoc, r.area, r.codManual, r.codInstructivo, r.codListado, r.codFormato, r.version, r.fechaEmision, r.elaboro, r.observaciones, r.deletedAt || ''];
    });
    var ws4 = XLSX.utils.aoa_to_sheet([papHeaders].concat(papRows));
    XLSX.utils.book_append_sheet(wb, ws4, 'Papelera');
  }

  if ((STATE.logs || []).length > 0) {
    var logHeaders = ['Fecha','Hora','Usuario','Tipo','Detalle','Documento'];
    var logRows = STATE.logs.map(function(l) {
      var fecha = '', hora = '';
      if (l.ts) {
        try {
          var d = new Date(l.ts);
          fecha = d.toLocaleDateString('es-MX');
          hora = d.toLocaleTimeString('es-MX');
        } catch(e) { fecha = l.ts; }
      }
      return [fecha, hora, l.user || '', l.type || '', l.details || '', l.doc || ''];
    });
    var ws5 = XLSX.utils.aoa_to_sheet([logHeaders].concat(logRows));
    XLSX.utils.book_append_sheet(wb, ws5, 'Logs');
  }

  if (AUTH.users && AUTH.users.length > 0) {
    var userHeaders = ['Usuario','Nombre Completo','Rol'];
    var userRows = AUTH.users.map(function(u) {
      return [u.username, u.full_name || u.fullName || '', u.role || 'user'];
    });
    var ws6 = XLSX.utils.aoa_to_sheet([userHeaders].concat(userRows));
    XLSX.utils.book_append_sheet(wb, ws6, 'Usuarios');
  }

  if (STATE.elaboros && STATE.elaboros.length > 0) {
    var elaHeaders = ['Nombre'];
    var elaRows = STATE.elaboros.map(function(e) { return [e]; });
    var ws7 = XLSX.utils.aoa_to_sheet([elaHeaders].concat(elaRows));
    XLSX.utils.book_append_sheet(wb, ws7, 'Elaboradores');
  }

  var resumen = [
    ['DEBBIOM - Backup COMPLETO'],
    [''],
    ['Fecha de Backup', now.toLocaleString('es-MX')],
    ['Generado por', AUTH.currentUser ? AUTH.currentUser.username : 'Sistema'],
    ['Versión de Sync', String(SYNC._version)],
    ['Next ID', STATE.nextId || 1000],
    [''],
    ['Sección', 'Total'],
    ['Registros activos', (STATE.records || []).length],
    ['Documentos obsoletos', (STATE.obsoletos || []).length],
    ['Registro de salidas', (STATE.salidas || []).length],
    ['Papelera', (STATE.papelera || []).length],
    ['Logs de actividad', (STATE.logs || []).length],
    ['Usuarios', (AUTH.users || []).length],
    ['Elaboradores', (STATE.elaboros || []).length]
  ];
  if (typeof getAlertasResumen === 'function') {
    var alertas = getAlertasResumen();
    resumen.push(['']);
    resumen.push(['Estado de Vigencia', 'Total']);
    resumen.push(['Vencidos', alertas.vencidos]);
    resumen.push(['Críticos (≤15 días)', alertas.criticos]);
    resumen.push(['Aviso (≤30 días)', alertas.avisos]);
    resumen.push(['Próximos (≤60 días)', alertas.proximos]);
  }
  var ws8 = XLSX.utils.aoa_to_sheet(resumen);
  XLSX.utils.book_append_sheet(wb, ws8, 'Resumen');

  var filename = 'DEBBIOM_BACKUP_COMPLETO_' + timestamp + '.xlsx';
  XLSX.writeFile(wb, filename);

  SYNC._lastBackupAt = Date.now();
  addLog('backup', 'Backup Excel COMPLETO generado: ' + filename);
  showToast('✅ Backup Excel descargado + guardado en la nube (' + (STATE.logs||[]).length + ' logs)', 'success');

  if (SYNC.initialized && SYNC._client) {
    _saveCloudBackup('excel', 'Backup Excel descargado en ' + (AUTH.currentUser ? AUTH.currentUser.username : 'desconocido'));
  }
}

function restoreBackup(file) {
  if (!file) return;

  var MAX_BACKUP_SIZE = 50 * 1024 * 1024;
  var MAX_RECORDS = 50000;

  if (file.size > MAX_BACKUP_SIZE) {
    showToast('Archivo demasiado grande (máx 50MB)', 'error');
    return;
  }

  var isEncrypted = file.name.endsWith('.dbk');

  if (isEncrypted) {
    var binReader = new FileReader();
    binReader.onload = async function(ev) {
      var pwd = prompt('Ingresa la contraseña del backup cifrado:');
      if (!pwd) { showToast('Se requiere contraseña', 'error'); return; }
      try {
        var jsonStr = await _decryptBackup(ev.target.result, pwd);
        _processRestoreData(jsonStr, MAX_BACKUP_SIZE, MAX_RECORDS);
      } catch(e) {
        showToast('Contraseña incorrecta o archivo corrupto', 'error');
      }
    };
    binReader.readAsArrayBuffer(file);
    return;
  }

  var reader = new FileReader();
  reader.onload = function(e) {
    _processRestoreData(e.target.result, MAX_BACKUP_SIZE, MAX_RECORDS);
  };
  reader.readAsText(file);
}

function _processRestoreData(raw, MAX_BACKUP_SIZE, MAX_RECORDS) {
    try {
      if (typeof raw !== 'string' || raw.length > MAX_BACKUP_SIZE) {
        showToast('Contenido de backup inválido o demasiado grande', 'error');
        return;
      }

      var data = JSON.parse(raw);

      if (!data || typeof data !== 'object' || Array.isArray(data)) {
        showToast('Archivo no es un backup válido', 'error');
        return;
      }

      if (!data._meta || data._meta.tipo !== 'DEBBIOM_BACKUP') {
        showToast('Archivo no es un backup válido de DEBBIOM', 'error');
        return;
      }

      if (!data.records || !Array.isArray(data.records)) {
        showToast('Backup corrupto: no contiene registros', 'error');
        return;
      }

      if (data.records.length > MAX_RECORDS) {
        showToast('Backup excede el límite de ' + MAX_RECORDS + ' registros', 'error');
        return;
      }

      var arrays = ['obsoletos','papelera','salidas','logs','elaboros'];
      for (var ai = 0; ai < arrays.length; ai++) {
        var k = arrays[ai];
        if (data[k] && (!Array.isArray(data[k]) || data[k].length > MAX_RECORDS)) {
          showToast('Backup contiene datos inválidos en ' + k, 'error');
          return;
        }
      }

      var msg = '¿Restaurar backup del ' + new Date(data._meta.fecha).toLocaleString('es-MX') + '?\n\n'
        + 'Contenido:\n'
        + '• ' + data.records.length + ' registros\n'
        + '• ' + (data.obsoletos || []).length + ' obsoletos\n'
        + '• ' + (data.salidas || []).length + ' salidas\n\n'
        + '⚠️ Esto REEMPLAZARÁ todos los datos actuales.';

      if (!confirm(msg)) return;

      STATE.records   = data.records;
      STATE.obsoletos = data.obsoletos || [];
      STATE.papelera  = data.papelera  || [];
      STATE.salidas   = data.salidas   || [];
      STATE.logs      = data.logs      || [];
      STATE.elaboros  = data.elaboros  || [];
      STATE.nextId    = data.nextId    || 1000;
      saveState();
      addLog('restore', 'Backup restaurado: ' + data._meta.fecha + ' (' + data.records.length + ' registros)');
      render();
      showToast('✅ Backup restaurado: ' + data.records.length + ' registros', 'success');

    } catch (err) {
      showToast('Error al leer el archivo de backup', 'error');
    }
}

function _startAutoBackup() {
  clearInterval(SYNC._backupTimer);
  SYNC._backupTimer = setInterval(function() {
    if (SYNC.initialized && SYNC._client && (STATE.records||[]).length > 0) {
      _saveCloudBackup('auto', 'Auto-backup programado');
    }
  }, AUTO_BACKUP_INTERVAL);
}
async function _saveCloudBackup(tipo, nota) {
  if (!SYNC.initialized || !SYNC._client) {
    return null;
  }
  if (!STATE.records || STATE.records.length === 0) {
    return null;
  }

  var now = new Date();
  var timestamp = now.getFullYear() + ''
    + String(now.getMonth()+1).padStart(2,'0')
    + String(now.getDate()).padStart(2,'0') + '_'
    + String(now.getHours()).padStart(2,'0')
    + String(now.getMinutes()).padStart(2,'0')
    + String(now.getSeconds()).padStart(2,'0');

  var nombre = 'BACKUP_' + tipo.toUpperCase() + '_' + timestamp;

  var stateData = {
    records:   STATE.records   || [],
    obsoletos: STATE.obsoletos || [],
    papelera:  STATE.papelera  || [],
    salidas:   STATE.salidas   || [],
    logs:      STATE.logs || [],
    elaboros:  STATE.elaboros  || [],
    nextId:    STATE.nextId    || 1000
  };

  var usersData = (AUTH.users || []).map(function(u) {
    return { username: u.username, role: u.role, full_name: u.full_name || u.fullName || '' };
  });

  try {
    var result = await SYNC._client
      .from('backups')
      .insert({
        nombre:          nombre,
        tipo:            tipo,
        created_by:      AUTH.currentUser ? AUTH.currentUser.username : 'Sistema',
        created_at:      now.toISOString(),
        sync_version:    SYNC._version || 0,
        total_registros: (STATE.records || []).length,
        total_obsoletos: (STATE.obsoletos || []).length,
        total_salidas:   (STATE.salidas || []).length,
        total_papelera:  (STATE.papelera || []).length,
        total_logs:      (STATE.logs || []).length,
        state:           stateData,
        users:           usersData,
        nota:            nota || ''
      })
      .select('id, nombre')
      .single();

    if (result.error) {
      return null;
    }

    return result.data;
  } catch (e) {
    return null;
  }
}

async function listCloudBackups() {
  if (!SYNC.initialized || !SYNC._client) return [];

  try {
    var result = await SYNC._client
      .from('backups')
      .select('id, nombre, tipo, created_by, created_at, sync_version, total_registros, total_obsoletos, total_salidas, total_papelera, total_logs, nota')
      .order('created_at', { ascending: false })
      .limit(50);

    if (result.error) {
      return [];
    }

    return result.data || [];
  } catch (e) {
    return [];
  }
}

async function restoreCloudBackup(backupId) {
  if (!SYNC.initialized || !SYNC._client) {
    showToast('No hay conexión a Supabase', 'error');
    return false;
  }

  try {
    var result = await SYNC._client
      .from('backups')
      .select('*')
      .eq('id', backupId)
      .single();

    if (result.error || !result.data) {
      showToast('Backup no encontrado', 'error');
      return false;
    }

    var bk = result.data;

    if (!bk.state || typeof bk.state !== 'object' || !bk.state.records || !Array.isArray(bk.state.records)) {
      showToast('Backup corrupto: sin registros válidos', 'error');
      return false;
    }

    var MAX_RECORDS = 50000;
    if (bk.state.records.length > MAX_RECORDS) {
      showToast('Backup excede el límite de ' + MAX_RECORDS + ' registros', 'error');
      return false;
    }

    var arrays = ['obsoletos','papelera','salidas','logs','elaboros'];
    for (var ai = 0; ai < arrays.length; ai++) {
      var k = arrays[ai];
      if (bk.state[k] && (!Array.isArray(bk.state[k]) || bk.state[k].length > MAX_RECORDS)) {
        showToast('Backup contiene datos inválidos en ' + k, 'error');
        return false;
      }
    }

    var msg = '¿Restaurar backup "' + bk.nombre + '"?\n\n'
      + 'Fecha: ' + new Date(bk.created_at).toLocaleString('es-MX') + '\n'
      + 'Creado por: ' + bk.created_by + '\n'
      + 'Registros: ' + bk.total_registros + '\n'
      + 'Obsoletos: ' + bk.total_obsoletos + '\n'
      + 'Salidas: ' + bk.total_salidas + '\n\n'
      + '⚠️ Esto REEMPLAZARÁ todos los datos actuales en TODOS los dispositivos.';

    if (!confirm(msg)) return false;

    STATE.records   = bk.state.records   || [];
    STATE.obsoletos = bk.state.obsoletos || [];
    STATE.papelera  = bk.state.papelera  || [];
    STATE.salidas   = bk.state.salidas   || [];
    STATE.logs      = bk.state.logs      || [];
    STATE.elaboros  = bk.state.elaboros  || [];
    STATE.nextId    = bk.state.nextId    || 1000;
    saveState();
    addLog('restore-cloud', 'Backup restaurado desde nube: ' + bk.nombre + ' (' + bk.total_registros + ' registros)');
    render();
    showToast('✅ Backup restaurado: ' + bk.total_registros + ' registros (sincronizando a todos los dispositivos...)', 'success');
    return true;

  } catch (e) {
    showToast('Error al restaurar backup', 'error');
    return false;
  }
}

async function deleteCloudBackup() {
  showToast('Los backups no se pueden eliminar', 'error');
  return false;
}

async function saveManualCloudBackup(nota) {
  var n = nota || 'Backup manual';
  showToast('☁️ Guardando backup en la nube...', 'info');
  var result = await _saveCloudBackup('manual', n);
  if (result) {
    addLog('backup-cloud', 'Backup guardado en la nube: ' + result.nombre);
    showToast('✅ Backup guardado en la nube', 'success');
  } else {
    showToast('❌ Error al guardar backup en la nube', 'error');
  }
  return result;
}

async function downloadCloudBackup(backupId) {
  if (!SYNC.initialized || !SYNC._client) return;

  try {
    var result = await SYNC._client
      .from('backups')
      .select('*')
      .eq('id', backupId)
      .single();

    if (result.error || !result.data) {
      showToast('Backup no encontrado', 'error');
      return;
    }

    var bk = result.data;
    var backupData = {
      _meta: {
        tipo:           'DEBBIOM_BACKUP',
        version:        '2.0',
        fecha:          bk.created_at,
        usuario:        bk.created_by,
        syncVersion:    bk.sync_version,
        totalRegistros: bk.total_registros,
        totalObsoletos: bk.total_obsoletos,
        totalSalidas:   bk.total_salidas,
        totalPapelera:  bk.total_papelera,
        totalLogs:      bk.total_logs,
        origen:         'supabase_cloud',
        nota:           bk.nota
      },
      records:   bk.state.records   || [],
      obsoletos: bk.state.obsoletos || [],
      papelera:  bk.state.papelera  || [],
      salidas:   bk.state.salidas   || [],
      logs:      bk.state.logs      || [],
      elaboros:  bk.state.elaboros  || [],
      nextId:    bk.state.nextId    || 1000,
      users:     bk.users           || []
    };

    var json = JSON.stringify(backupData, null, 2);
    var blob = new Blob([json], { type: 'application/json' });
    var url  = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = bk.nombre + '.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast('✅ Backup descargado: ' + bk.nombre + '.json', 'success');
  } catch (e) {
    showToast('Error descargando backup', 'error');
  }
}

async function _cleanOldCloudBackups() { return; }
function registerPresence() {
  if (!SYNC.initialized || !SYNC._client || !AUTH.currentUser) return;

  if (SYNC._presence) {
    try {
      SYNC._presence.untrack();
      SYNC._client.removeChannel(SYNC._presence);
    } catch (e) {}
    SYNC._presence = null;
  }

  SYNC._presence = SYNC._client.channel('online-users', {
    config: { presence: { key: SYNC.sessionId } }
  });

  SYNC._presence.on('presence', { event: 'sync' }, function() {
    var state = SYNC._presence.presenceState();
    var users = [];
    for (var key in state) {
      var entries = state[key];
      if (entries && entries.length > 0) {
        users.push(entries[0].username);
      }
    }
    SYNC.onlineUsers = users.filter(function(v, i, a) { return a.indexOf(v) === i; });
    SYNC.onlineCount = SYNC.onlineUsers.length;
    updateOnlineCountUI();
  });

  SYNC._presence.subscribe(async function(status) {
    if (status === 'SUBSCRIBED') {
      await SYNC._presence.track({
        username:  AUTH.currentUser.username,
        fullName:  AUTH.currentUser.fullName || AUTH.currentUser.username,
        online_at: new Date().toISOString()
      });
    }
  });
}

function unregisterPresence() {
  if (SYNC._presence) {
    try {
      SYNC._presence.untrack();
      SYNC._client.removeChannel(SYNC._presence);
    } catch (e) {}
    SYNC._presence = null;
  }
}

function updateOnlineCountUI() {
  var el1 = document.querySelector('.online-count');
  if (el1) el1.textContent = String(SYNC.onlineCount);
  var el2 = document.getElementById('sidebar-online-count');
  if (el2) el2.textContent = String(SYNC.onlineCount);
}
function _updateLastSyncUI() {
  var el = document.getElementById('last-sync-time');
  if (!el) return;
  if (!SYNC._lastSyncTime) { el.textContent = '—'; return; }
  var diff = Math.floor((Date.now() - SYNC._lastSyncTime) / 1000);
  if (diff < 5) el.textContent = 'Ahora';
  else if (diff < 60) el.textContent = 'Hace ' + diff + 's';
  else if (diff < 3600) el.textContent = 'Hace ' + Math.floor(diff/60) + ' min';
  else el.textContent = 'Hace ' + Math.floor(diff/3600) + 'h';
}
setInterval(_updateLastSyncUI, 10000);

function updateSyncIndicator() {
  var el = document.getElementById('sync-indicator');
  if (!el) return;
  el.textContent = '';
  var icon = document.createElement('i');
  var txt;
  if (!SYNC.initialized) {
    el.className = 'sync-indicator sync-offline';
    icon.className = 'fas fa-database';
    txt = ' Local';
  } else if (SYNC.connected) {
    el.className = 'sync-indicator sync-online';
    icon.className = 'fas fa-cloud';
    txt = ' En línea';
  } else {
    el.className = 'sync-indicator sync-reconnecting';
    icon.className = 'fas fa-cloud-arrow-up';
    txt = ' Reconectando...';
  }
  el.appendChild(icon);
  el.appendChild(document.createTextNode(txt));
}

function _showSyncAnimation(active) {
  var btn = document.getElementById('btn-sync');
  if (!btn) return;
  btn.textContent = '';
  var icon = document.createElement('i');
  if (active) {
    btn.classList.add('syncing');
    icon.className = 'fas fa-spinner fa-spin';
    btn.appendChild(icon);
    btn.appendChild(document.createTextNode(' Sincronizando...'));
  } else {
    btn.classList.remove('syncing');
    icon.className = 'fas fa-arrows-rotate';
    btn.appendChild(icon);
    btn.appendChild(document.createTextNode(' Sincronizar'));
  }
}
async function syncLoadState() {
  var status = await _syncLoadWithStatus();
  return status === 'loaded';
}

async function _syncLoadWithStatus() {
  if (!SYNC.initialized || !SYNC._client) return 'error';

  try {
    var result = await SYNC._client
      .from('app_state')
      .select('*')
      .eq('id', 1)
      .single();

    if (result.error) return 'error';
    if (!result.data || !result.data.state) return 'empty';

    if (result.data.state.records && result.data.state.records.length > 0) {
      var clean = _sanitizeIncoming(result.data.state);
      if (!clean) return 'error';
      STATE.records   = clean.records;
      STATE.obsoletos = clean.obsoletos;
      STATE.papelera  = clean.papelera;
      STATE.salidas   = clean.salidas;
      STATE.logs      = clean.logs;
      STATE.elaboros  = clean.elaboros;
      STATE.nextId    = clean.nextId;
      SYNC._version   = result.data.version || 0;
      SYNC._lastKnownRecordCount = clean.records.length;
      return 'loaded';
    }

    return 'empty';
  } catch (e) {
    return 'error';
  }
}
function syncSaveUsers() { /* Usuarios gestionados por Supabase Auth */ }
async function syncLoadUsers() { return null; }
function listenUserChanges() { /* Usuarios gestionados por Supabase Auth */ }
function cleanupSync() {
  clearTimeout(SYNC._saveTimer);
  clearInterval(SYNC._heartbeatTimer);
  clearInterval(SYNC._vigenciaTimer);
  clearInterval(SYNC._backupTimer);
  if (SYNC._channel) {
    SYNC._client.removeChannel(SYNC._channel);
  }
  unregisterPresence();
}

window.addEventListener('beforeunload', cleanupSync);
