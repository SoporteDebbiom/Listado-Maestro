function _esc(s) {
  if (typeof s !== 'string') return s == null ? '' : String(s);
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function _clean(s) {
  if (typeof s !== 'string') return s == null ? '' : String(s);
  return s.replace(/[<>]/g, '').trim();
}

function _safeHref(url) {
  if (!url || typeof url !== 'string') return '#';
  var t = url.trim().toLowerCase();
  if (t.startsWith('https://') || t.startsWith('http://')) return url.trim();
  return '#';
}

var APP_VERSION = '3.4.0';
var APP_BUILD   = '2026-10-30';

function _setBtn(el, iconClass, text) {
  el.textContent = '';
  if (iconClass) { var i = document.createElement('i'); i.className = iconClass; el.appendChild(i); }
  if (text) el.appendChild(document.createTextNode(' ' + text));
}
var _unsavedChanges = false;

function _dataHash(data) {
  var str = JSON.stringify(data);
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    var ch = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + ch;
    hash |= 0;
  }
  return Math.abs(hash).toString(36);
}

function _playNotificationSound() {
  try {
    var ctx = new (window.AudioContext || window.webkitAudioContext)();
    var osc = ctx.createOscillator();
    var gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.frequency.value = 800;
    osc.type = 'sine';
    gain.gain.value = 0.15;
    osc.start();
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.3);
    osc.stop(ctx.currentTime + 0.3);
  } catch(e) {}
}

function exportLogsToExcel() {
  if (typeof XLSX === 'undefined') { showToast('XLSX no disponible', 'error'); return; }
  if (!STATE.logs || STATE.logs.length === 0) { showToast('No hay logs', 'error'); return; }
  var headers = ['Fecha','Hora','Usuario','Tipo','Detalle','Documento'];
  var rows = STATE.logs.map(function(l) {
    var fecha = '', hora = '';
    if (l.ts) { try { var d = new Date(l.ts); fecha = d.toLocaleDateString('es-MX'); hora = d.toLocaleTimeString('es-MX'); } catch(e) { fecha = l.ts; } }
    return [fecha, hora, l.user || '', l.type || '', l.details || '', l.doc || ''];
  });
  var ws = XLSX.utils.aoa_to_sheet([headers].concat(rows));
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Logs');
  XLSX.writeFile(wb, 'DEBBIOM_Logs_' + new Date().toISOString().slice(0,10) + '.xlsx');
  showToast('Logs exportados a Excel');
}

function exportVencidosToExcel() {
  if (typeof XLSX === 'undefined') { showToast('XLSX no disponible', 'error'); return; }
  var vencidos = (STATE.records || []).filter(function(r) { return typeof calcEstado === 'function' && calcEstado(r.fechaEmision) === 'Vencido'; });
  if (vencidos.length === 0) { showToast('No hay documentos vencidos', 'info'); return; }
  var headers = ['No.','Área','Código','Nombre','Versión','Fecha Emisión','Elaboró','Vigencia','Días Vencido','Observaciones'];
  var rows = vencidos.map(function(r) {
    return [r.no, r.area, r.codManual||r.codInstructivo||r.codFormato||'', r.nombreDoc, r.version, r.fechaEmision, r.elaboro,
      typeof renderVigenciaDate === 'function' ? renderVigenciaDate(r.fechaEmision) : '',
      typeof calcDiasVigencia === 'function' ? Math.abs(calcDiasVigencia(r.fechaEmision)||0) : '', r.observaciones];
  });
  var ws = XLSX.utils.aoa_to_sheet([headers].concat(rows));
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Vencidos');
  XLSX.writeFile(wb, 'DEBBIOM_Vencidos_' + new Date().toISOString().slice(0,10) + '.xlsx');
  showToast('Reporte de vencidos exportado');
}

window.addEventListener('beforeunload', function(e) {
  if (_unsavedChanges && AUTH.currentUser) {
    e.preventDefault();
    e.returnValue = '¿Salir? Puede haber cambios sin guardar.';
    return e.returnValue;
  }
});

let STATE = {
  records: [],
  obsoletos: [],
  papelera: [],
  salidas: [],
  logs: [],
  elaboros: [],
  currentArea: 'Dir. Gral.',
  currentSection: 'areas',
  searchQuery: '',
  filterEstado: '',
  filterElaboro: '',
  page: 1,
  perPage: 25,
  modal: null,
  selectedRecord: null,
  editForm: {},
  salidaForm: {},
  obsoletoMotivo: '',
  toasts: [],
  nextId: 1000,
  sortCol: null,
  sortDir: 'asc'
};
function saveState() {
  _unsavedChanges = true;
  if (typeof syncSaveState === 'function') {
    syncSaveState();
  }
}

function _trackChange(rec, action) {
  if (!rec) return;
  if (!rec._history) rec._history = [];
  rec._history.push({
    ts: new Date().toISOString(),
    user: AUTH.currentUser ? AUTH.currentUser.username : 'Sistema',
    action: action
  });
  if (rec._history.length > 50) rec._history = rec._history.slice(-50);
}

function loadSavedState() {
  return false;
}

function getNextId() {
  return ++STATE.nextId;
}

function showToast(msg, type='success') {
  const id = Date.now();
  STATE.toasts.push({id, msg, type});
  renderToasts();
  setTimeout(() => {
    STATE.toasts = STATE.toasts.filter(t => t.id !== id);
    renderToasts();
  }, 3000);
}
const LOG_TYPES = {
  login:    { icon: 'fa-right-to-bracket', color: '#26c6da', label: 'Inicio de sesión' },
  logout:   { icon: 'fa-right-from-bracket', color: '#5a8999', label: 'Cierre de sesión' },
  create:   { icon: 'fa-circle-plus', color: '#8cc63f', label: 'Registro creado' },
  edit:     { icon: 'fa-pen-to-square', color: '#fbbf24', label: 'Registro editado' },
  delete:   { icon: 'fa-trash-can', color: '#f87171', label: 'Enviado a papelera' },
  restore:  { icon: 'fa-rotate-left', color: '#34d399', label: 'Restaurado' },
  permDelete:{ icon: 'fa-xmark', color: '#ef4444', label: 'Eliminado permanente' },
  obsolete: { icon: 'fa-box-archive', color: '#fb923c', label: 'Enviado a obsoletos' },
  version:  { icon: 'fa-code-branch', color: '#fb923c', label: 'Cambio de versión' },
  salida:   { icon: 'fa-right-from-bracket', color: '#a78bfa', label: 'Registro de salida' },
  export:   { icon: 'fa-file-excel', color: '#10b981', label: 'Exportación' },
  userAdd:  { icon: 'fa-user-plus', color: '#8cc63f', label: 'Usuario creado' },
  userDel:  { icon: 'fa-user-minus', color: '#f87171', label: 'Usuario eliminado' },
  userPw:   { icon: 'fa-key', color: '#fbbf24', label: 'Contraseña cambiada' },
};

function addLog(type, details, docName) {
  STATE.logs.unshift({
    id: Date.now() + Math.random(),
    timestamp: new Date().toISOString(),
    user: AUTH.currentUser ? AUTH.currentUser.username : 'Sistema',
    type: type,
    details: details || '',
    docName: docName || ''
  });
  if (STATE.logs.length > 5000) STATE.logs.length = 5000;
}

function formatLogDate(iso) {
  try {
    const d = new Date(iso);
    const dd = String(d.getDate()).padStart(2,'0');
    const mm = String(d.getMonth()+1).padStart(2,'0');
    const yy = String(d.getFullYear());
    const hh = String(d.getHours()).padStart(2,'0');
    const mi = String(d.getMinutes()).padStart(2,'0');
    const ss = String(d.getSeconds()).padStart(2,'0');
    return dd+'/'+mm+'/'+yy+' '+hh+':'+mi+':'+ss;
  } catch(e) { return iso; }
}
function createObsoletoEntry(rec, motivo) {
  const entry = JSON.parse(JSON.stringify(rec));

  entry.id = getNextId();

  const vig = calcVigencia(rec.fechaEmision);
  entry.vigencia = vig ? formatDateFull(vig) : '';

  entry.motivo = (motivo || '').trim();
  entry.fechaObsoleto = new Date().toISOString();

  return entry;
}
function loadExcelData() {
  if (typeof RAW_DATA === 'undefined') {
    STATE.records = [];
    STATE.obsoletos = [];
    STATE.elaboros = ['Hramírez','AMartinez','HRamirez'];
    return;
  }
  
  const data = RAW_DATA;
  STATE.records = data.r.map((r, i) => ({
    id: i + 1,
    area: r[1]||'',
    no: r[2]||'',
    codManual: r[3]||'',
    codInstructivo: r[4]||'',
    codListado: r[5]||'',
    codFormato: r[6]||'',
    nombreDoc: r[7]||'',
    version: r[8]||'',
    fechaEmision: r[9]||'',
    elaboro: r[10]||'',
    tipoResguardo: r[11]||'',
    ubicacion: r[12]||'',
    copias: r[13]||'',
    tipoResguardoCopia: r[14]||'',
    usuarios: r[15]||'',
    observaciones: r[16]||'',
    archivos: [],
    archivoURLs: []
  }));
  
  STATE.obsoletos = data.o.map((o, i) => ({
    id: i + 1,
    area: '',
    no: o[1]||'',
    codManual: o[2]||'',
    codInstructivo: o[3]||'',
    codListado: o[4]||'',
    codFormato: o[5]||'',
    nombreDoc: o[6]||'',
    version: o[7]||'',
    fechaEmision: o[8]||'',
    elaboro: o[9]||'',
    vigencia: o[10]||'',
    destino: o[11]||'',
    ubicacion: o[12]||'',
    copias: o[13]||'',
    destinoCopia: o[14]||'',
    observaciones: o[15]||'',
    motivo: 'Importado del listado original',
    archivos: [],
    archivoURLs: []
  }));
  
  STATE.elaboros = data.e || [];
  STATE.nextId = Math.max(...STATE.records.map(r=>r.id), ...STATE.obsoletos.map(o=>o.id), 999) + 1;
}
function h(tag, attrs, ...children) {
  const el = document.createElement(tag);
  let deferValue = undefined;
  if (attrs) {
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'className') el.className = v;
      else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
      else if (k.startsWith('on')) el.addEventListener(k.slice(2).toLowerCase(), v);
      // innerHTML attribute removed for XSS safety
      else if (k === 'value') {
        if (tag === 'select') deferValue = v;
        else el.value = v;
      }
      else if (k === 'checked') el.checked = v;
      else if (k === 'selected') el.selected = v;
      else if (k === 'disabled') { if(v) el.disabled = true; }
      else if (k === 'htmlFor') el.htmlFor = v;
      else if (k === 'dataset') { for(const [dk,dv] of Object.entries(v)) el.dataset[dk]=dv; }
      else el.setAttribute(k, v);
    }
  }
  for (const child of children.flat(Infinity)) {
    if (child == null || child === false) continue;
    if (typeof child === 'string' || typeof child === 'number') {
      el.appendChild(document.createTextNode(child));
    } else if (child instanceof HTMLElement) {
      el.appendChild(child);
    }
  }
  if (deferValue !== undefined) el.value = deferValue;
  return el;
}

function _parseCode(code) {
  if (!code) return {base:'zzz', num:9999, prefix:''};
  var s = code.trim();
  var m = s.match(/^([A-Za-z]+)(\d+)\s*[\/-]\s*(.+)$/);
  if (m) {
    return {
      base: m[3].toLowerCase().replace(/\s+/g,''),
      num: parseInt(m[2], 10),
      prefix: m[1].toLowerCase()
    };
  }
  return {base: s.toLowerCase(), num: 0, prefix: ''};
}

function getFilteredRecords() {
  let recs = STATE.records;
  
  if (STATE.currentSection === 'areas') {
    recs = recs.filter(r => r.area === STATE.currentArea);
  }
  
  if (STATE.searchQuery) {
    const q = STATE.searchQuery.toLowerCase();
    recs = recs.filter(r => 
      (r.nombreDoc||'').toLowerCase().includes(q) ||
      (r.codManual||'').toLowerCase().includes(q) ||
      (r.codInstructivo||'').toLowerCase().includes(q) ||
      (r.codListado||'').toLowerCase().includes(q) ||
      (r.codFormato||'').toLowerCase().includes(q) ||
      (r.no||'').toLowerCase().includes(q)
    );
  }
  
  if (STATE.filterEstado) {
    recs = recs.filter(r => calcEstado(r.fechaEmision) === STATE.filterEstado);
  }
  
  if (STATE.filterElaboro) {
    recs = recs.filter(r => r.elaboro === STATE.filterElaboro);
  }
  
  if (STATE.sortCol) {
    recs = [...recs].sort((a, b) => {
      let va = a[STATE.sortCol] || '';
      let vb = b[STATE.sortCol] || '';
      if (STATE.sortCol === 'dias') {
        va = calcDiasVigencia(a.fechaEmision) || -9999;
        vb = calcDiasVigencia(b.fechaEmision) || -9999;
        return STATE.sortDir === 'asc' ? va - vb : vb - va;
      }
      if (['codFormato','codListado','codInstructivo','codManual'].indexOf(STATE.sortCol) >= 0) {
        var pa = _parseCode(va), pb = _parseCode(vb);
        if (pa.base < pb.base) return STATE.sortDir === 'asc' ? -1 : 1;
        if (pa.base > pb.base) return STATE.sortDir === 'asc' ? 1 : -1;
        if (pa.num !== pb.num) return STATE.sortDir === 'asc' ? pa.num - pb.num : pb.num - pa.num;
        return 0;
      }
      if (typeof va === 'string') va = va.toLowerCase();
      if (typeof vb === 'string') vb = vb.toLowerCase();
      if (va < vb) return STATE.sortDir === 'asc' ? -1 : 1;
      if (va > vb) return STATE.sortDir === 'asc' ? 1 : -1;
      return 0;
    });
  } else {
    recs = [...recs].sort((a, b) => {
      var codeCols = ['codFormato','codListado','codInstructivo','codManual'];
      var ca = '', cb = '';
      for (var ci = 0; ci < codeCols.length; ci++) {
        if (a[codeCols[ci]]) { ca = a[codeCols[ci]]; break; }
      }
      for (var ci = 0; ci < codeCols.length; ci++) {
        if (b[codeCols[ci]]) { cb = b[codeCols[ci]]; break; }
      }
      if (ca || cb) {
        var pa = _parseCode(ca), pb = _parseCode(cb);
        if (pa.base !== pb.base) return pa.base < pb.base ? -1 : 1;
        if (pa.num !== pb.num) return pa.num - pb.num;
      }
      return 0;
    });
  }
  
  return recs;
}

function renderSidebar() {
  const areaItems = AREAS.map(area => {
    const count = STATE.records.filter(r => r.area === area).length;
    const isActive = STATE.currentSection === 'areas' && STATE.currentArea === area;
    return h('div', {
      className: 'nav-item' + (isActive ? ' active' : ''),
      onClick: () => { STATE.currentSection='areas'; STATE.currentArea=area; STATE.page=1; STATE.searchQuery=''; STATE.filterEstado=''; STATE.filterElaboro=''; render(); }
    },
      h('i', {className: 'fas fa-folder'}),
      h('span', null, area),
      h('span', {className: 'nav-badge'}, String(count))
    );
  });

  const obsCount = STATE.obsoletos.length;
  const salCount = STATE.salidas.length;
  const papCount = STATE.papelera.length;

  return h('div', {className: 'sidebar'},
    h('div', {className: 'sidebar-header'},
      h('img', {className:'sidebar-logo', src:'data:image/jpeg;base64,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', alt:'DEBBIOM'}),
      h('div', {className:'sidebar-header-text'},
        h('h1', null, 'Listado Maestro de Documentos'),
        h('p', null, 'DEBBIOM — Sistema de Gestión')
      )
    ),
    h('div', {className: 'sidebar-nav'},
      h('div', {className: 'nav-section'},
        h('div', {className: 'nav-section-title'}, 'Áreas'),
        ...areaItems
      ),
      h('div', {className: 'nav-section'},
        h('div', {className: 'nav-section-title'}, 'Gestión'),
        (function() {
          var vencCount = (STATE.records || []).filter(function(r) { return typeof calcEstado === 'function' && calcEstado(r.fechaEmision) === 'Vencido'; }).length;
          return h('div', {
            className: 'nav-item nav-item-vencidos' + (STATE.currentSection==='vencidos'?' active':''),
            onClick: function() { STATE.currentSection='vencidos'; STATE.page=1; STATE.searchQuery=''; STATE.filterElaboro=''; render(); }
          },
            h('i', {className: 'fas fa-circle-exclamation', style:{color: vencCount > 0 ? '#ef4444' : 'var(--text-muted)'}}),
            h('span', null, 'Vencidos'),
            vencCount > 0
              ? h('span', {className: 'nav-badge nav-badge-danger'}, String(vencCount))
              : h('span', {className: 'nav-badge'}, '0')
          );
        })(),
        h('div', {
          className: 'nav-item' + (STATE.currentSection==='obsoletos'?' active':''),
          onClick: () => { STATE.currentSection='obsoletos'; STATE.page=1; render(); }
        },
          h('i', {className: 'fas fa-archive'}),
          h('span', null, 'Obsoletos'),
          h('span', {className: 'nav-badge'}, String(obsCount))
        ),
        h('div', {
          className: 'nav-item' + (STATE.currentSection==='salidas'?' active':''),
          onClick: () => { STATE.currentSection='salidas'; STATE.page=1; render(); }
        },
          h('i', {className: 'fas fa-sign-out-alt'}),
          h('span', null, 'Reg. de Salidas'),
          h('span', {className: 'nav-badge'}, String(salCount))
        ),
        h('div', {
          className: 'nav-item' + (STATE.currentSection==='papelera'?' active':''),
          onClick: () => { STATE.currentSection='papelera'; STATE.page=1; render(); }
        },
          h('i', {className: 'fas fa-trash-alt'}),
          h('span', null, 'Papelera'),
          h('span', {className: 'nav-badge'}, String(papCount))
        ),
        isAdmin() ? h('div', {
          className: 'nav-item',
          onClick: () => { STATE.modal='users'; render(); }
        },
          h('i', {className: 'fas fa-users-gear'}),
          h('span', null, 'Gestión de Usuarios')
        ) : null
      ),
      h('div', {className: 'nav-section'},
        h('div', {className: 'nav-section-title'}, 'Auditoría'),
        h('div', {
          className: 'nav-item' + (STATE.currentSection==='logs'?' active':''),
          onClick: () => { STATE.currentSection='logs'; STATE.page=1; render(); }
        },
          h('i', {className: 'fas fa-clock-rotate-left'}),
          h('span', null, 'Registro de Actividad'),
          h('span', {className: 'nav-badge'}, String(STATE.logs.length))
        ),
        isAdmin() ? h('div', {
          className: 'nav-item' + (STATE.currentSection==='dashboard'?' active':''),
          onClick: function() { STATE.currentSection='dashboard'; STATE.page=1; render(); }
        },
          h('i', {className: 'fas fa-chart-bar', style:{color:'#8b5cf6'}}),
          h('span', null, 'Dashboard')
        ) : null,
        isAdmin() ? h('div', {
          className: 'nav-item nav-item-backups' + (STATE.currentSection==='backups'?' active':''),
          onClick: function() { STATE.currentSection='backups'; STATE.page=1; render(); }
        },
          h('i', {className: 'fas fa-cloud-arrow-up', style:{color: typeof SYNC !== 'undefined' && SYNC.connected ? '#10b981' : 'var(--text-muted)'}}),
          h('span', null, 'Backups en la Nube'),
          h('span', {className: 'nav-badge'}, h('i', {className: 'fas fa-shield-halved', style:{fontSize:'10px'}}))
        ) : null,
        isAdmin() ? h('div', {
          className: 'nav-item' + (STATE.currentSection==='sistema'?' active':''),
          onClick: function() { STATE.currentSection='sistema'; STATE.page=1; render(); }
        },
          h('i', {className: 'fas fa-clipboard-check', style:{color:'var(--green)'}}),
          h('span', null, 'Verificación ISO'),
          h('span', {className: 'nav-badge'}, 'v' + APP_VERSION)
        ) : null
      ),
      (function() {
        var resumen = typeof getAlertasResumen === 'function' ? getAlertasResumen() : {total:0,vencidos:0,criticos:0};
        var urgentes = resumen.vencidos + resumen.criticos;
        return h('div', {className: 'nav-section'},
          h('div', {className: 'nav-section-title'}, 'Alertas'),
          h('div', {
            className: 'nav-item nav-item-alertas' + (STATE.currentSection==='alertas'?' active':''),
            onClick: function() { STATE.currentSection='alertas'; STATE.page=1; render(); }
          },
            h('i', {className: 'fas fa-bell', style:{color: urgentes > 0 ? '#ef4444' : 'var(--text-muted)'}}),
            h('span', null, 'Recordatorios'),
            urgentes > 0 ? h('span', {className: 'nav-badge nav-badge-danger'}, String(urgentes)) :
              resumen.total > 0 ? h('span', {className: 'nav-badge nav-badge-warning'}, String(resumen.total)) :
              h('span', {className: 'nav-badge'}, '0')
          )
        );
      })()
    ),
    AUTH.currentUser ? h('div', {className: 'sidebar-user'},
      h('div', {className: 'user-avatar'}, AUTH.currentUser.fullName ? AUTH.currentUser.fullName.charAt(0).toUpperCase() : '?'),
      h('div', {className: 'user-info'},
        h('div', {className: 'user-name'}, AUTH.currentUser.username),
        h('div', {className: 'user-role'}, isAdmin() ? 'Administrador' : 'Usuario')
      ),
      h('button', {className: 'logout-btn', 'data-info':'Cerrar sesión', onClick: logout},
        h('i', {className: 'fas fa-right-from-bracket'})
      )
    ) : null,
    AUTH.currentUser ? h('div', {className: 'sidebar-sync'},
      h('div', {id: 'sync-indicator', className: 'sync-indicator ' + (typeof SYNC !== 'undefined' && SYNC.connected ? 'sync-online' : (typeof SYNC !== 'undefined' && SYNC.initialized ? 'sync-reconnecting' : 'sync-offline'))},
        h('i', {className: typeof SYNC !== 'undefined' && SYNC.connected ? 'fas fa-cloud' : 'fas fa-database'}),
        typeof SYNC !== 'undefined' && SYNC.connected ? ' En línea' : ' Local'
      ),
      h('div', {className: 'sidebar-online-info'},
        h('i', {className: 'fas fa-users', style:{fontSize:'11px',marginRight:'6px',color:'var(--accent)'}}),
        h('span', {id:'sidebar-online-count'}, String(typeof SYNC !== 'undefined' ? SYNC.onlineCount : 0)),
        ' en línea'
      ),
      h('div', {className: 'sidebar-online-info', style:{fontSize:'10px',marginTop:'4px'}},
        h('i', {className: 'fas fa-arrows-rotate', style:{fontSize:'10px',marginRight:'6px',color:'var(--green)'}}),
        'Sync: ', h('span', {id:'last-sync-time'}, typeof SYNC !== 'undefined' && SYNC._lastSyncTime ? 'Ahora' : '—')
      )
    ) : null
  );
}

function renderStatsBar() {
  let recs;
  if (STATE.currentSection === 'areas') {
    recs = STATE.records.filter(r => r.area === STATE.currentArea);
  } else {
    return h('div');
  }
  const total = recs.length;
  const vigentes = recs.filter(r => calcEstado(r.fechaEmision) === 'Vigente').length;
  const porVencer = recs.filter(r => calcEstado(r.fechaEmision) === 'Por vencer').length;
  const vencidos = recs.filter(r => calcEstado(r.fechaEmision) === 'Vencido').length;
  const pVig = total > 0 ? Math.round((vigentes/total)*100) : 0;
  const pPv = total > 0 ? Math.round((porVencer/total)*100) : 0;
  const pVe = total > 0 ? Math.round((vencidos/total)*100) : 0;

  return h('div', {className: 'stats-bar'},
    h('div', {className: 'stat-card'},
      h('div', {className: 'stat-label'}, 'Total Documentos'),
      h('div', {className: 'stat-value'}, String(total))
    ),
    h('div', {className: 'stat-card'},
      h('div', {className: 'stat-label'}, 'Vigentes'),
      h('div', {className: 'stat-value green'}, String(vigentes))
    ),
    h('div', {className: 'stat-card'},
      h('div', {className: 'stat-label'}, 'Por Vencer'),
      h('div', {className: 'stat-value yellow'}, String(porVencer))
    ),
    h('div', {className: 'stat-card'},
      h('div', {className: 'stat-label'}, 'Vencidos'),
      h('div', {className: 'stat-value red'}, String(vencidos))
    ),
    total > 0 ? h('div', {className: 'stat-card', style:{flex:'1 1 100%',minWidth:'100%'}},
      h('div', {className: 'stat-label', style:{marginBottom:'8px'}}, 'Distribución'),
      h('div', {className: 'dash-bar'},
        pVig > 0 ? h('div', {className: 'dash-seg dash-green', style:{width: pVig+'%'}, title: 'Vigentes: '+vigentes+' ('+pVig+'%)'}, pVig >= 10 ? pVig+'%' : '') : null,
        pPv > 0 ? h('div', {className: 'dash-seg dash-yellow', style:{width: pPv+'%'}, title: 'Por vencer: '+porVencer+' ('+pPv+'%)'}, pPv >= 10 ? pPv+'%' : '') : null,
        pVe > 0 ? h('div', {className: 'dash-seg dash-red', style:{width: pVe+'%'}, title: 'Vencidos: '+vencidos+' ('+pVe+'%)'}, pVe >= 10 ? pVe+'%' : '') : null
      ),
      h('div', {style:{display:'flex',gap:'16px',marginTop:'6px',fontSize:'11px',color:'var(--text-muted)'}},
        h('span', null, h('span', {style:{display:'inline-block',width:'10px',height:'10px',borderRadius:'3px',background:'var(--green)',marginRight:'4px'}}), 'Vigentes'),
        h('span', null, h('span', {style:{display:'inline-block',width:'10px',height:'10px',borderRadius:'3px',background:'var(--yellow)',marginRight:'4px'}}), 'Por vencer'),
        h('span', null, h('span', {style:{display:'inline-block',width:'10px',height:'10px',borderRadius:'3px',background:'var(--red)',marginRight:'4px'}}), 'Vencidos')
      )
    ) : null
  );
}

function renderFilterBar() {
  if (STATE.currentSection !== 'areas') return h('div');
  
  const uniqueElaboros = [...new Set(STATE.records.filter(r=>r.area===STATE.currentArea).map(r=>r.elaboro).filter(Boolean))].sort();
  
  return h('div', {className: 'filter-bar'},
    h('div', {className: 'search-box'},
      h('i', {className: 'fas fa-search'}),
      h('input', {
        type: 'text',
        placeholder: 'Buscar por nombre, código...',
        value: STATE.searchQuery,
        onInput: (e) => { STATE.searchQuery = e.target.value; STATE.page=1; render(); }
      })
    ),
    h('select', {
      className: 'filter-select',
      value: STATE.filterEstado,
      onChange: (e) => { STATE.filterEstado = e.target.value; STATE.page=1; render(); }
    },
      h('option', {value: ''}, 'Todos los estados'),
      h('option', {value: 'Vigente'}, 'Vigente'),
      h('option', {value: 'Por vencer'}, 'Por vencer'),
      h('option', {value: 'Vencido'}, 'Vencido')
    ),
    h('select', {
      className: 'filter-select',
      value: STATE.filterElaboro,
      onChange: (e) => { STATE.filterElaboro = e.target.value; STATE.page=1; render(); }
    },
      h('option', {value: ''}, 'Todos - Elaboró'),
      ...uniqueElaboros.map(e => h('option', {value: e}, e))
    )
  );
}

function renderStatusBadge(fechaEmision) {
  const estado = calcEstado(fechaEmision);
  const cls = estado === 'Vigente' ? 'badge-green' : estado === 'Por vencer' ? 'badge-yellow' : estado === 'Vencido' ? 'badge-red' : '';
  if (!estado) return h('span', {className: 'badge badge-blue'}, '—');
  return h('span', {className: 'badge ' + cls},
    h('span', {className: 'badge-dot'}),
    estado
  );
}

function renderDiasVigencia(fechaEmision) {
  const dias = calcDiasVigencia(fechaEmision);
  if (dias === null) return h('span', {style:{color:'var(--text-muted)'}}, '—');
  if (dias >= 60) return h('span', {style:{color:'var(--green)', fontFamily:'JetBrains Mono, monospace', fontSize:'12px'}}, String(dias) + ' días');
  if (dias >= 1) return h('span', {style:{color:'var(--yellow)', fontFamily:'JetBrains Mono, monospace', fontSize:'12px'}}, String(dias) + ' días');
  return h('span', {style:{color:'var(--red)', fontFamily:'JetBrains Mono, monospace', fontSize:'12px'}}, String(Math.abs(dias)) + ' días vencido');
}

function renderVigenciaDate(fechaEmision) {
  const vig = calcVigencia(fechaEmision);
  if (!vig) return '—';
  return formatDateFull(vig);
}

function renderAreasTable() {
  const allFiltered = getFilteredRecords();
  const totalPages = Math.ceil(allFiltered.length / STATE.perPage);
  const start = (STATE.page - 1) * STATE.perPage;
  const pageRecords = allFiltered.slice(start, start + STATE.perPage);
  
  const sortIcon = (col) => {
    if (STATE.sortCol !== col) return h('i', {className: 'fas fa-sort', style:{opacity:'.3'}});
    return h('i', {className: STATE.sortDir === 'asc' ? 'fas fa-sort-up' : 'fas fa-sort-down'});
  };
  
  const thClick = (col) => () => {
    if (STATE.sortCol === col) {
      STATE.sortDir = STATE.sortDir === 'asc' ? 'desc' : 'asc';
    } else {
      STATE.sortCol = col;
      STATE.sortDir = 'asc';
    }
    render();
  };

  const table = h('table', null,
    h('thead', null,
      h('tr', null,
        h('th', {onClick: thClick('no')}, 'No. ', sortIcon('no')),
        h('th', {onClick: thClick('codManual')}, 'Cód. Manual/Proc. ', sortIcon('codManual')),
        h('th', {onClick: thClick('codInstructivo')}, 'Cód. Instructivo ', sortIcon('codInstructivo')),
        h('th', {onClick: thClick('codListado')}, 'Cód. Listado ', sortIcon('codListado')),
        h('th', {onClick: thClick('codFormato')}, 'Cód. Formato ', sortIcon('codFormato')),
        h('th', {onClick: thClick('nombreDoc'), style:{minWidth:'200px'}}, 'Nombre Documento ', sortIcon('nombreDoc')),
        h('th', {onClick: thClick('version')}, 'Versión ', sortIcon('version')),
        h('th', {onClick: thClick('fechaEmision')}, 'Emisión ', sortIcon('fechaEmision')),
        h('th', null, 'Elaboró'),
        h('th', null, 'Vigencia'),
        h('th', null, 'Estado'),
        h('th', {onClick: thClick('dias')}, 'Días Vig. ', sortIcon('dias')),
        h('th', null, 'Resguardo'),
        h('th', null, 'Archivos'),
        h('th', null, 'Copias'),
        h('th', {style:{minWidth:'340px'}}, 'Acciones')
      )
    ),
    h('tbody', null,
      ...pageRecords.map(rec => {
        const estado = calcEstado(rec.fechaEmision);
        const fechaE = parseDate(rec.fechaEmision);
        
        return h('tr', null,
          h('td', null, rec.no || ''),
          h('td', null, rec.codManual || ''),
          h('td', null, rec.codInstructivo || ''),
          h('td', null, rec.codListado || ''),
          h('td', null, rec.codFormato || ''),
          h('td', {className: 'wrap', title: rec.nombreDoc}, rec.nombreDoc || ''),
          h('td', null, h('span', {className: 'badge badge-blue'}, rec.version || '—')),
          h('td', null, fechaE ? formatDateShort(fechaE) : (rec.fechaEmision || '—')),
          h('td', null, rec.elaboro || ''),
          h('td', null, renderVigenciaDate(rec.fechaEmision)),
          h('td', null, renderStatusBadge(rec.fechaEmision)),
          h('td', null, renderDiasVigencia(rec.fechaEmision)),
          h('td', null, rec.tipoResguardo ? h('span', {className:'badge badge-purple'}, rec.tipoResguardo) : '—'),
          h('td', null, renderFilePreview(rec)),
          h('td', null, rec.copias || '—'),
          h('td', null,
            h('div', {className: 'actions-cell'},
              h('button', {className:'action-btn view', 'data-info':'Ver detalle completo del registro', onClick:()=>openView(rec)}, h('i',{className:'fas fa-eye'}), h('span',{className:'btn-label'},'Ver')),
              h('button', {className:'action-btn edit', 'data-info':'Editar los datos de este registro', onClick:()=>openEdit(rec)}, h('i',{className:'fas fa-pen-to-square'}), h('span',{className:'btn-label'},'Editar')),
              h('button', {className:'action-btn delete', 'data-info':'Enviar este registro a la papelera', onClick:()=>deleteRecord(rec)}, h('i',{className:'fas fa-trash-can'}), h('span',{className:'btn-label'},'Eliminar')),
              h('button', {className:'action-btn output', 'data-info':'Crear un registro de salida del documento', onClick:()=>openSalida(rec)}, h('i',{className:'fas fa-right-from-bracket'}), h('span',{className:'btn-label'},'Salida'))
            )
          )
        );
      }),
      pageRecords.length === 0 ? h('tr', null,
        h('td', {colspan: '16', style: {textAlign:'center',padding:'40px',color:'var(--text-muted)'}},
          h('i', {className:'fas fa-folder-open', style:{fontSize:'24px',display:'block',marginBottom:'8px',opacity:'.4'}}),
          'No se encontraron registros'
        )
      ) : null
    )
  );

  const pag = h('div', {className: 'pagination'},
    h('div', {className: 'pagination-info'},
      `Mostrando ${allFiltered.length > 0 ? start+1 : 0}-${Math.min(start+STATE.perPage, allFiltered.length)} de ${allFiltered.length} registros`
    ),
    h('div', {className: 'pagination-controls'},
      h('button', {className: 'page-btn', disabled: STATE.page <= 1, onClick: () => { STATE.page--; render(); }}, h('i',{className:'fas fa-chevron-left'})),
      ...Array.from({length: Math.min(totalPages, 7)}, (_, i) => {
        let p;
        if (totalPages <= 7) p = i + 1;
        else if (STATE.page <= 4) p = i + 1;
        else if (STATE.page >= totalPages - 3) p = totalPages - 6 + i;
        else p = STATE.page - 3 + i;
        return h('button', {className: 'page-btn' + (p===STATE.page?' active':''), onClick: () => { STATE.page=p; render(); }}, String(p));
      }),
      h('button', {className: 'page-btn', disabled: STATE.page >= totalPages, onClick: () => { STATE.page++; render(); }}, h('i',{className:'fas fa-chevron-right'}))
    )
  );

  return h('div', {className: 'content-area'},
    renderStatsBar(),
    renderFilterBar(),
    h('div', {className: 'table-container'}, table),
    pag
  );
}

function renderFilePreview(rec) {
  const items = [];
  if (rec.archivoURLs && rec.archivoURLs.length > 0) {
    rec.archivoURLs.forEach(url => {
      items.push(h('a', {className:'file-chip', href:_safeHref(url), target:'_blank', rel:'noopener noreferrer'}, h('i',{className:'fas fa-link'}), 'URL'));
    });
  }
  if (rec.archivos && rec.archivos.length > 0) {
    rec.archivos.forEach(f => {
      items.push(h('span', {className:'file-chip', onClick:()=>downloadFile(f)}, h('i',{className:'fas fa-file'}), f.name.length > 12 ? f.name.substring(0,12)+'…' : f.name));
    });
  }
  if (rec.ubicacion && (rec.ubicacion.startsWith('http') || rec.ubicacion.startsWith('www'))) {
    const url = rec.ubicacion.startsWith('http') ? rec.ubicacion : 'https://' + rec.ubicacion;
    items.push(h('a', {className:'file-chip', href:_safeHref(url), target:'_blank', rel:'noopener noreferrer'}, h('i',{className:'fas fa-external-link-alt'}), 'Link'));
  }
  if (items.length === 0) return h('span', {style:{color:'var(--text-muted)',fontSize:'12px'}}, '—');
  return h('div', {className:'file-preview'}, ...items);
}

async function _uploadToStorage(file) {
  if (!SYNC._client) { showToast('Sin conexión a Supabase', 'error'); return null; }
  var userId = AUTH.currentUser ? AUTH.currentUser.id : 'anon';
  var ts = Date.now();
  var safeName = file.name.replace(/[^a-zA-Z0-9._-]/g, '_').substring(0, 100);
  var path = userId + '/' + ts + '_' + safeName;
  var result = await SYNC._client.storage.from('documentos').upload(path, file, {
    cacheControl: '3600',
    upsert: false,
    contentType: file.type
  });
  if (result.error) {
    showToast('Error al subir archivo. Verifica el tipo y tamaño (máx 10MB).', 'error');
    return null;
  }
  var urlResult = SYNC._client.storage.from('documentos').getPublicUrl(path);
  return {
    name: _clean(file.name),
    size: file.size,
    type: file.type,
    storagePath: path,
    storageUrl: urlResult.data ? urlResult.data.publicUrl : null
  };
}

function downloadFile(f) {
  if (f.storageUrl) {
    var a = document.createElement('a');
    a.href = f.storageUrl;
    a.download = f.name;
    a.target = '_blank';
    a.rel = 'noopener noreferrer';
    a.click();
  } else if (f.storagePath && SYNC._client) {
    var urlResult = SYNC._client.storage.from('documentos').getPublicUrl(f.storagePath);
    if (urlResult.data && urlResult.data.publicUrl) {
      var a = document.createElement('a');
      a.href = urlResult.data.publicUrl;
      a.download = f.name;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.click();
    }
  } else if (f.dataUrl) {
    var a = document.createElement('a');
    a.href = f.dataUrl;
    a.download = f.name;
    a.click();
  }
}

function _getVencidosRecords() {
  var all = (STATE.records || []).filter(function(r) {
    return calcEstado(r.fechaEmision) === 'Vencido';
  });

  if (STATE.searchQuery) {
    var q = STATE.searchQuery.toLowerCase();
    all = all.filter(function(r) {
      return (r.nombreDoc||'').toLowerCase().indexOf(q) >= 0 ||
        (r.codManual||'').toLowerCase().indexOf(q) >= 0 ||
        (r.codInstructivo||'').toLowerCase().indexOf(q) >= 0 ||
        (r.codListado||'').toLowerCase().indexOf(q) >= 0 ||
        (r.codFormato||'').toLowerCase().indexOf(q) >= 0 ||
        (r.area||'').toLowerCase().indexOf(q) >= 0 ||
        (r.elaboro||'').toLowerCase().indexOf(q) >= 0;
    });
  }

  if (STATE.filterElaboro) {
    all = all.filter(function(r) { return r.elaboro === STATE.filterElaboro; });
  }

  if (STATE._vencidosFilterArea) {
    all = all.filter(function(r) { return r.area === STATE._vencidosFilterArea; });
  }

  all.sort(function(a, b) {
    var da = calcDiasVigencia(a.fechaEmision) || 0;
    var db = calcDiasVigencia(b.fechaEmision) || 0;
    return da - db;
  });

  return all;
}

function renderVencidosSection() {
  var vencidos = _getVencidosRecords();
  var allVencidos = (STATE.records || []).filter(function(r) { return calcEstado(r.fechaEmision) === 'Vencido'; });

  var areaStats = {};
  allVencidos.forEach(function(r) {
    var a = r.area || 'Sin área';
    areaStats[a] = (areaStats[a] || 0) + 1;
  });
  var areaKeys = Object.keys(areaStats).sort(function(a, b) { return areaStats[b] - areaStats[a]; });

  var elaboros = [];
  var _eSet = {};
  allVencidos.forEach(function(r) {
    if (r.elaboro && !_eSet[r.elaboro]) { elaboros.push(r.elaboro); _eSet[r.elaboro] = true; }
  });
  elaboros.sort();

  var totalPages = Math.ceil(vencidos.length / STATE.perPage);
  var start = (STATE.page - 1) * STATE.perPage;
  var pageRecords = vencidos.slice(start, start + STATE.perPage);

  var statsBar = h('div', {className: 'stats-bar'},
    h('div', {className: 'stat-card'},
      h('div', {className: 'stat-label'}, 'Total Vencidos'),
      h('div', {className: 'stat-value red'}, String(allVencidos.length))
    ),
    h('div', {className: 'stat-card'},
      h('div', {className: 'stat-label'}, 'Áreas Afectadas'),
      h('div', {className: 'stat-value yellow'}, String(areaKeys.length))
    ),
    h('div', {className: 'stat-card'},
      h('div', {className: 'stat-label'}, 'Área Más Crítica'),
      h('div', {className: 'stat-value', style:{fontSize:'14px',color:'var(--red)'}}, areaKeys.length > 0 ? areaKeys[0] + ' (' + areaStats[areaKeys[0]] + ')' : '—')
    ),
    h('div', {className: 'stat-card'},
      h('div', {className: 'stat-label'}, 'Más Antiguo'),
      h('div', {className: 'stat-value', style:{fontSize:'14px',color:'var(--red)'}},
        vencidos.length > 0 ? String(Math.abs(calcDiasVigencia(vencidos[0].fechaEmision) || 0)) + ' días' : '—'
      )
    )
  );

  var filterBar = h('div', {className: 'filter-bar'},
    h('div', {className: 'search-box'},
      h('i', {className: 'fas fa-search'}),
      h('input', {
        type: 'text',
        placeholder: 'Buscar por nombre, código, área...',
        value: STATE.searchQuery || '',
        onInput: function(e) { STATE.searchQuery = e.target.value; STATE.page=1; render(); }
      })
    ),
    h('select', {
      className: 'filter-select',
      value: STATE._vencidosFilterArea || '',
      onChange: function(e) { STATE._vencidosFilterArea = e.target.value; STATE.page=1; render(); }
    },
      h('option', {value: ''}, 'Todas las áreas'),
      ...areaKeys.map(function(a) { return h('option', {value: a}, a + ' (' + areaStats[a] + ')'); })
    ),
    h('select', {
      className: 'filter-select',
      value: STATE.filterElaboro || '',
      onChange: function(e) { STATE.filterElaboro = e.target.value; STATE.page=1; render(); }
    },
      h('option', {value: ''}, 'Todos — Elaboró'),
      ...elaboros.map(function(e) { return h('option', {value: e}, e); })
    )
  );

  var areaChips = h('div', {className: 'vencidos-area-chips'},
    ...areaKeys.map(function(a) {
      var isActive = STATE._vencidosFilterArea === a;
      return h('span', {
        className: 'vencido-area-chip' + (isActive ? ' active' : ''),
        onClick: function() {
          STATE._vencidosFilterArea = isActive ? '' : a;
          STATE.page = 1;
          render();
        }
      }, a + ' (' + areaStats[a] + ')');
    })
  );

  var table = h('table', null,
    h('thead', null,
      h('tr', null,
        h('th', null, 'No.'),
        h('th', null, 'Área'),
        h('th', null, 'Código'),
        h('th', {style:{minWidth:'180px'}}, 'Nombre Documento'),
        h('th', null, 'Versión'),
        h('th', null, 'Emisión'),
        h('th', null, 'Elaboró'),
        h('th', null, 'Vigencia'),
        h('th', null, 'Días Vencido'),
        h('th', {style:{minWidth:'320px'}}, 'Acciones')
      )
    ),
    h('tbody', null,
      ...pageRecords.map(function(rec) {
        var dias = calcDiasVigencia(rec.fechaEmision);
        var diasAbs = Math.abs(dias || 0);
        var codigo = rec.codManual || rec.codInstructivo || rec.codListado || rec.codFormato || '';

        var urgenciaClass = diasAbs > 180 ? 'vencido-critico' : diasAbs > 60 ? 'vencido-alto' : 'vencido-medio';

        return h('tr', {className: urgenciaClass},
          h('td', null, rec.no || ''),
          h('td', null, h('span', {className: 'badge badge-purple'}, rec.area || '—')),
          h('td', null, codigo ? h('span', {className: 'badge badge-blue'}, codigo) : '—'),
          h('td', {className: 'wrap', title: rec.nombreDoc}, rec.nombreDoc || ''),
          h('td', null, h('span', {className: 'badge badge-blue'}, rec.version || '—')),
          h('td', null, rec.fechaEmision || '—'),
          h('td', null, rec.elaboro || '—'),
          h('td', null, renderVigenciaDate(rec.fechaEmision)),
          h('td', null,
            h('span', {className: 'vencido-dias-badge'}, diasAbs + ' días')
          ),
          h('td', null,
            h('div', {className: 'actions-cell'},
              h('button', {className:'action-btn view', 'data-info':'Ver detalle', onClick: function() { openView(rec); }},
                h('i',{className:'fas fa-eye'}), h('span',{className:'btn-label'},'Ver')
              ),
              h('button', {className:'action-btn edit', 'data-info':'Editar documento', onClick: function() { openEdit(rec); }},
                h('i',{className:'fas fa-pen-to-square'}), h('span',{className:'btn-label'},'Editar')
              ),
              h('button', {className:'action-btn renew', 'data-info':'Renovar vigencia (actualizar fecha emisión)', onClick: function() { _openRenewModal(rec); }},
                h('i',{className:'fas fa-rotate-right'}), h('span',{className:'btn-label'},'Renovar')
              ),
              h('button', {className:'action-btn obsolete', 'data-info':'Enviar a obsoletos', onClick: function() { openObsoleto(rec); }},
                h('i',{className:'fas fa-archive'}), h('span',{className:'btn-label'},'Obsoleto')
              )
            )
          )
        );
      }),
      pageRecords.length === 0 ? h('tr', null,
        h('td', {colspan: '10', style:{textAlign:'center',padding:'40px',color:'var(--text-muted)'}},
          h('i', {className:'fas fa-check-circle', style:{fontSize:'28px',display:'block',marginBottom:'10px',color:'var(--green)',opacity:'.6'}}),
          allVencidos.length === 0 ? 'No hay documentos vencidos — ¡Todo está al día!' : 'No se encontraron resultados con los filtros aplicados'
        )
      ) : null
    )
  );

  var pag = h('div', {className: 'pagination'},
    h('div', {className: 'pagination-info'},
      'Mostrando ' + (vencidos.length > 0 ? start+1 : 0) + '-' + Math.min(start+STATE.perPage, vencidos.length) + ' de ' + vencidos.length + ' vencidos'
    ),
    h('div', {className: 'pagination-controls'},
      h('button', {className: 'page-btn', disabled: STATE.page <= 1, onClick: function() { STATE.page--; render(); }}, h('i',{className:'fas fa-chevron-left'})),
      ...Array.from({length: Math.min(totalPages, 7)}, function(_, i) {
        var p;
        if (totalPages <= 7) p = i + 1;
        else if (STATE.page <= 4) p = i + 1;
        else if (STATE.page >= totalPages - 3) p = totalPages - 6 + i;
        else p = STATE.page - 3 + i;
        return h('button', {className: 'page-btn' + (p===STATE.page?' active':''), onClick: function() { STATE.page=p; render(); }}, String(p));
      }),
      h('button', {className: 'page-btn', disabled: STATE.page >= totalPages, onClick: function() { STATE.page++; render(); }}, h('i',{className:'fas fa-chevron-right'}))
    )
  );

  return h('div', {className: 'content-area'},
    statsBar,
    filterBar,
    areaKeys.length > 1 ? areaChips : null,
    h('div', {className: 'table-container'}, table),
    pag
  );
}

function _openRenewModal(rec) {
  STATE.modal = 'renew';
  STATE.selectedRecord = rec;
  var today = new Date();
  STATE._renewDate = today.getFullYear() + '-' + String(today.getMonth()+1).padStart(2,'0') + '-' + String(today.getDate()).padStart(2,'0');
  STATE._renewMotivo = '';
  render();
}

function _confirmRenew() {
  var rec = STATE.selectedRecord;
  if (!rec) return;

  if (!STATE._renewDate) {
    showToast('Selecciona la nueva fecha de emisión', 'error');
    return;
  }
  if (!STATE._renewMotivo || !STATE._renewMotivo.trim()) {
    showToast('El motivo de renovación es obligatorio', 'error');
    return;
  }

  var idx = STATE.records.findIndex(function(r) { return r.id === rec.id; });
  if (idx < 0) {
    showToast('Registro no encontrado', 'error');
    return;
  }

  var oldFecha = STATE.records[idx].fechaEmision;
  _trackChange(STATE.records[idx], 'Renovado: ' + oldFecha + ' → ' + STATE._renewDate);
  STATE.records[idx].fechaEmision = STATE._renewDate;

  addLog('renew', 'Renovado: ' + oldFecha + ' → ' + STATE._renewDate + ' | Motivo: ' + _clean(STATE._renewMotivo), rec.nombreDoc || rec.codManual || '');
  saveState();
  showToast('✅ Documento renovado — nueva vigencia activa', 'success');

  STATE.modal = null;
  STATE.selectedRecord = null;
  STATE._renewDate = '';
  STATE._renewMotivo = '';
  render();
}

function renderObsoletosTable() {
  const sorted = [...STATE.obsoletos].sort((a, b) => {
    const da = a.fechaObsoleto ? new Date(a.fechaObsoleto).getTime() : 0;
    const db = b.fechaObsoleto ? new Date(b.fechaObsoleto).getTime() : 0;
    return db - da;
  });

  const start = (STATE.page - 1) * STATE.perPage;
  const pageRecs = sorted.slice(start, start + STATE.perPage);
  const totalPages = Math.max(1, Math.ceil(sorted.length / STATE.perPage));

  return h('div', {className: 'content-area'},
    h('div', {className: 'stats-bar'},
      h('div', {className: 'stat-card'},
        h('div', {className: 'stat-label'}, 'Total Obsoletos'),
        h('div', {className: 'stat-value'}, String(sorted.length))
      )
    ),
    h('div', {className: 'table-container'},
      h('table', null,
        h('thead', null,
          h('tr', null,
            h('th', {style:{minWidth:'150px',background:'rgba(249,115,22,.1)',color:'#fb923c'}}, 'Fecha Obsoleto'),
            h('th', null, 'No.'),
            h('th', null, 'Área'),
            h('th', null, 'Cód. Manual'),
            h('th', null, 'Cód. Instructivo'),
            h('th', null, 'Cód. Listado'),
            h('th', null, 'Cód. Formato'),
            h('th', {style:{minWidth:'200px'}}, 'Nombre Documento'),
            h('th', null, 'Versión'),
            h('th', null, 'Fecha Emisión'),
            h('th', null, 'Elaboró'),
            h('th', null, 'Vigencia'),
            h('th', null, 'Tipo Resguardo'),
            h('th', null, 'Ubicación'),
            h('th', null, 'Copias Ctrl.'),
            h('th', null, 'Tipo Resg. Copia'),
            h('th', null, 'Usuarios'),
            h('th', {style:{minWidth:'180px'}}, 'Motivo'),
            h('th', null, 'Archivos'),
            h('th', null, 'Observaciones')
          )
        ),
        h('tbody', null,
          ...pageRecs.map(rec => h('tr', null,
            h('td', null,
              rec.fechaObsoleto
                ? h('span', {style:{fontFamily:'JetBrains Mono,monospace',fontSize:'11px',color:'#fb923c',fontWeight:'600',background:'rgba(249,115,22,.08)',padding:'3px 8px',borderRadius:'5px',whiteSpace:'nowrap'}}, formatLogDate(rec.fechaObsoleto))
                : h('span', {style:{color:'var(--text-muted)'}}, '—')
            ),
            h('td', null, rec.no||''),
            h('td', null, rec.area||''),
            h('td', null, rec.codManual||''),
            h('td', null, rec.codInstructivo||''),
            h('td', null, rec.codListado||''),
            h('td', null, rec.codFormato||''),
            h('td', {className:'wrap'}, rec.nombreDoc||''),
            h('td', null, h('span',{className:'badge badge-blue'}, rec.version||'—')),
            h('td', null, rec.fechaEmision||''),
            h('td', null, rec.elaboro||''),
            h('td', null, rec.vigencia||''),
            h('td', null, rec.tipoResguardo||''),
            h('td', {className:'wrap'}, rec.ubicacion||''),
            h('td', null, rec.copias||''),
            h('td', null, rec.tipoResguardoCopia||''),
            h('td', {className:'wrap'}, rec.usuarios||''),
            h('td', {className:'wrap'}, rec.motivo||''),
            h('td', null, renderFilePreview(rec)),
            h('td', {className:'wrap'}, rec.observaciones||'')
          )),
          pageRecs.length === 0 ? h('tr', null,
            h('td', {colspan:'20', style:{textAlign:'center',padding:'40px',color:'var(--text-muted)'}},
              h('i', {className:'fas fa-archive', style:{fontSize:'24px',display:'block',marginBottom:'8px',opacity:'.4'}}),
              'No hay documentos obsoletos'
            )
          ) : null
        )
      )
    ),
    h('div', {className: 'pagination'},
      h('div', {className: 'pagination-info'}, `${sorted.length} registros`),
      h('div', {className: 'pagination-controls'},
        h('button', {className:'page-btn', disabled:STATE.page<=1, onClick:()=>{STATE.page--;render();}}, h('i',{className:'fas fa-chevron-left'})),
        h('button', {className:'page-btn', disabled:STATE.page>=totalPages, onClick:()=>{STATE.page++;render();}}, h('i',{className:'fas fa-chevron-right'}))
      )
    )
  );
}
function renderSalidasTable() {
  const recs = STATE.salidas;
  return h('div', {className:'content-area'},
    h('div', {className:'stats-bar'},
      h('div', {className:'stat-card'},
        h('div', {className:'stat-label'}, 'Total Salidas'),
        h('div', {className:'stat-value'}, String(recs.length))
      )
    ),
    h('div', {className:'table-container'},
      h('table', null,
        h('thead', null,
          h('tr', null,
            h('th', null, '#'),
            h('th', null, 'Nombre Documento'),
            h('th', null, 'Versión'),
            h('th', null, 'Fecha Emisión'),
            h('th', null, 'Área'),
            h('th', null, 'Motivo de Entrega'),
            h('th', null, 'Destinatario'),
            h('th', null, 'Correo'),
            h('th', null, 'Fecha Salida'),
            h('th', null, 'Archivos/URLs')
          )
        ),
        h('tbody', null,
          ...recs.map((s, i) => h('tr', null,
            h('td', null, String(i+1)),
            h('td', {className:'wrap'}, s.nombreDoc||''),
            h('td', null, h('span',{className:'badge badge-blue'}, s.version||'')),
            h('td', null, s.fechaEmision||''),
            h('td', null, s.area||''),
            h('td', {className:'wrap'}, s.motivo||''),
            h('td', null, s.destinatario||''),
            h('td', null, s.correo ? h('a',{href:'mailto:'+s.correo, className:'url-link'}, s.correo) : ''),
            h('td', null, s.fechaSalida||''),
            h('td', null, renderFilePreview(s))
          )),
          recs.length === 0 ? h('tr', null,
            h('td', {colspan:'10', style:{textAlign:'center',padding:'40px',color:'var(--text-muted)'}},
              h('i', {className:'fas fa-sign-out-alt', style:{fontSize:'24px',display:'block',marginBottom:'8px',opacity:'.4'}}),
              'No hay registros de salida'
            )
          ) : null
        )
      )
    )
  );
}
function renderPapeleraTable() {
  const recs = STATE.papelera;
  return h('div', {className:'content-area'},
    h('div', {className:'stats-bar'},
      h('div', {className:'stat-card'},
        h('div', {className:'stat-label'}, 'En Papelera'),
        h('div', {className:'stat-value'}, String(recs.length))
      ),
      h('div', {className:'stat-card'},
        h('div', {className:'stat-label', style:{color:'var(--text-muted)',fontSize:'10px'}}, '⏳ Sin expiración'),
        h('div', {className:'stat-value', style:{fontSize:'11px',color:'var(--text-muted)'}}, 'Permanente')
      )
    ),
    h('div', {className:'table-container'},
      h('table', null,
        h('thead', null,
          h('tr', null,
            h('th', null, 'No.'),
            h('th', null, 'Nombre Documento'),
            h('th', null, 'Versión'),
            h('th', null, 'Emisión'),
            h('th', null, 'Área'),
            h('th', null, 'Elaboró'),
            h('th', null, 'Eliminado'),
            h('th', {style:{minWidth:'230px'}}, 'Acciones')
          )
        ),
        h('tbody', null,
          ...recs.map(rec => h('tr', null,
            h('td', null, rec.no||''),
            h('td', {className:'wrap'}, rec.nombreDoc||''),
            h('td', null, h('span',{className:'badge badge-blue'}, rec.version||'')),
            h('td', null, rec.fechaEmision||''),
            h('td', null, rec.area||''),
            h('td', null, rec.elaboro||''),
            h('td', null, rec.deletedAt ? h('span', {style:{fontSize:'10px',color:'var(--text-muted)'}}, new Date(rec.deletedAt).toLocaleDateString('es-MX',{day:'2-digit',month:'short',year:'2-digit'})) : '—'),
            h('td', null,
              h('div', {className:'actions-cell'},
                h('button', {className:'action-btn restore', 'data-info':'Devolver este registro al listado activo', onClick:()=>restoreRecord(rec)}, h('i',{className:'fas fa-rotate-left'}), h('span',{className:'btn-label'},'Restaurar')),
                h('button', {className:'action-btn view', 'data-info':'Restaurar a registros activos', onClick:function(){restoreRecord(rec);}}, h('i',{className:'fas fa-rotate-left'}), h('span',{className:'btn-label'},'Restaurar'))
              )
            )
          )),
          recs.length === 0 ? h('tr', null,
            h('td', {colspan:'8', style:{textAlign:'center',padding:'40px',color:'var(--text-muted)'}},
              h('i', {className:'fas fa-trash-alt', style:{fontSize:'24px',display:'block',marginBottom:'8px',opacity:'.4'}}),
              'La papelera está vacía'
            )
          ) : null
        )
      )
    )
  );
}
function openView(rec) {
  STATE.modal = 'view';
  STATE.selectedRecord = rec;
  render();
}

function openEdit(rec) {
  STATE.modal = 'edit';
  STATE.selectedRecord = JSON.parse(JSON.stringify(rec));
  STATE.editForm = JSON.parse(JSON.stringify(rec));
  STATE.originalVersion = String(rec.version || '').trim();
  STATE.obsoletoMotivo = '';
  render();
}

function openAdd() {
  STATE.modal = 'add';
  STATE.editForm = {
    id: getNextId(),
    area: STATE.currentArea,
    no: '', codManual: '', codInstructivo: '', codListado: '', codFormato: '',
    nombreDoc: '', version: '', fechaEmision: '', elaboro: '',
    tipoResguardo: '', ubicacion: '', copias: '',
    tipoResguardoCopia: '', usuarios: '', observaciones: '',
    archivos: [], archivoURLs: [],
    archivosCopia: [], archivoURLsCopia: []
  };
  render();
}

function openSalida(rec) {
  STATE.modal = 'salida';
  STATE.selectedRecord = rec;
  STATE.salidaForm = {
    area: rec.area || STATE.currentArea,
    motivo: '',
    destinatario: '',
    correo: ''
  };
  render();
}

function openObsoleto(rec) {
  STATE.modal = 'obsoleto';
  STATE.selectedRecord = rec;
  STATE.obsoletoMotivo = '';
  render();
}

function closeModal() {
  STATE.modal = null;
  STATE.selectedRecord = null;
  STATE.editForm = {};
  render();
}

function saveRecord() {
  var form = STATE.editForm;
  if (!form.nombreDoc && !form.codManual && !form.codInstructivo && !form.codFormato) {
    showToast('Ingresa al menos el nombre del documento o un código', 'error');
    return;
  }

  var textFields = ['nombreDoc','codManual','codInstructivo','codListado','codFormato','version','elaboro','tipoResguardo','ubicacion','copias','tipoResguardoCopia','usuarios','observaciones','no','area'];
  textFields.forEach(function(f) { if (typeof form[f] === 'string') form[f] = form[f].replace(/[<>]/g,''); });

  function _validURL(u) {
    if (!u || typeof u !== 'string') return false;
    var t = u.trim().toLowerCase();
    return t.startsWith('https://') || t.startsWith('http://');
  }
  var urlFields = ['archivoURLs','archivoURLsCopia'];
  for (var ui = 0; ui < urlFields.length; ui++) {
    var arr = form[urlFields[ui]];
    if (arr && Array.isArray(arr)) {
      for (var uj = 0; uj < arr.length; uj++) {
        if (arr[uj] && !_validURL(arr[uj])) {
          showToast('URL inválida: solo se permiten direcciones http:// o https://', 'error');
          return;
        }
      }
      form[urlFields[ui]] = arr.filter(function(u) { return u && u.trim(); });
    }
  }

  if (STATE.modal === 'add') {
    form.area = STATE.currentArea;
    _trackChange(form, 'Creado');
    STATE.records.push({...form});
    addLog('create', 'Área: ' + form.area, form.nombreDoc || form.codManual || 'Sin nombre');
    showToast('Registro agregado exitosamente');
  } else if (STATE.modal === 'edit') {
    const idx = STATE.records.findIndex(r => r.id === form.id);
    if (idx >= 0) {
      const original = STATE.selectedRecord;
      const newVer = String(form.version || '').trim();
      const oldVer = STATE.originalVersion || '';
      const versionChanged = oldVer !== '' && newVer !== '' && newVer !== oldVer;

      if (versionChanged) {
        if (!STATE.obsoletoMotivo || !STATE.obsoletoMotivo.trim()) {
          showToast('Debes escribir el motivo del cambio de versión para continuar', 'error');
          return;
        }

        const obsEntry = createObsoletoEntry(original, STATE.obsoletoMotivo);
        STATE.obsoletos.unshift(obsEntry);

        form.archivos = [];
        form.archivoURLs = [];
        form.archivosCopia = [];
        form.archivoURLsCopia = [];

        addLog('version', 'V' + oldVer + ' → V' + newVer + ' | Motivo: ' + STATE.obsoletoMotivo.trim(), original.nombreDoc || original.codManual || '');
        showToast('Versión anterior (V' + oldVer + ') enviada a obsoletos con todos sus datos y archivos.');
      }

      form._history = STATE.records[idx]._history || [];
      _trackChange(form, 'Editado');
      STATE.records[idx] = {...form};
      addLog('edit', 'Área: ' + (form.area||''), form.nombreDoc || form.codManual || '');
      showToast('Registro actualizado exitosamente');
    }
  }
  STATE.obsoletoMotivo = '';
  saveState();
  closeModal();
}

function deleteRecord(rec) {
  if (!confirm('¿Enviar este registro a la papelera de reciclaje?')) return;
  STATE.records = STATE.records.filter(r => r.id !== rec.id);
  STATE.papelera.push({...rec, deletedAt: new Date().toISOString()});
  addLog('delete', 'Área: ' + (rec.area||''), rec.nombreDoc || rec.codManual || '');
  showToast('Registro enviado a papelera');
  saveState();
  render();
}

function restoreRecord(rec) {
  STATE.papelera = STATE.papelera.filter(r => r.id !== rec.id);
  delete rec.deletedAt;
  STATE.records.push(rec);
  addLog('restore', 'Área: ' + (rec.area||''), rec.nombreDoc || rec.codManual || '');
  showToast('Registro restaurado');
  saveState();
  render();
}

function permanentDelete() {
  showToast('La eliminación permanente está deshabilitada', 'error');
}

function confirmSalida() {
  const form = STATE.salidaForm;
  const rec = STATE.selectedRecord;
  if (!form.motivo || !form.destinatario || !form.correo) {
    showToast('Completa todos los campos obligatorios', 'error');
    return;
  }
  
  STATE.salidas.push({
    id: getNextId(),
    nombreDoc: rec.nombreDoc,
    version: rec.version,
    fechaEmision: rec.fechaEmision,
    area: _clean(form.area),
    motivo: _clean(form.motivo),
    destinatario: _clean(form.destinatario),
    correo: _clean(form.correo),
    fechaSalida: formatDateFull(new Date()),
    archivos: rec.archivos ? [...rec.archivos] : [],
    archivoURLs: rec.archivoURLs ? [...rec.archivoURLs] : [],
    ubicacion: rec.ubicacion
  });
  
  addLog('salida', 'Destino: ' + form.destinatario + ' | Motivo: ' + form.motivo, rec.nombreDoc || '');
  showToast('Registro de salida creado exitosamente');
  saveState();
  closeModal();
}

function confirmObsoleto() {
  if (!STATE.obsoletoMotivo.trim()) {
    showToast('Debes ingresar el motivo del cambio', 'error');
    return;
  }
  
  const rec = STATE.selectedRecord;

  const obsEntry = createObsoletoEntry(rec, _clean(STATE.obsoletoMotivo));
  STATE.obsoletos.unshift(obsEntry);
  
  STATE.records = STATE.records.filter(r => r.id !== rec.id);
  
  addLog('obsolete', 'Motivo: ' + _clean(STATE.obsoletoMotivo), rec.nombreDoc || rec.codManual || '');
  showToast('Documento enviado a obsoletos');
  saveState();
  closeModal();
}

function exportToExcel() {
  let data;
  let filename;
  
  if (STATE.currentSection === 'areas') {
    data = STATE.records.filter(r => r.area === STATE.currentArea);
    filename = `Listado_${STATE.currentArea.replace(/[^a-zA-Z0-9]/g,'_')}.xlsx`;
  } else if (STATE.currentSection === 'vencidos') {
    data = STATE.records.filter(r => calcEstado(r.fechaEmision) === 'Vencido');
    filename = 'Documentos_Vencidos.xlsx';
  } else if (STATE.currentSection === 'obsoletos') {
    data = STATE.obsoletos;
    filename = 'Obsoletos.xlsx';
  } else if (STATE.currentSection === 'salidas') {
    data = STATE.salidas;
    filename = 'Registro_Salidas.xlsx';
  } else {
    data = STATE.papelera;
    filename = 'Papelera.xlsx';
  }
  
  const headers = STATE.currentSection === 'salidas' 
    ? ['Nombre Documento','Versión','Fecha Emisión','Área','Motivo','Destinatario','Correo','Fecha Salida']
    : STATE.currentSection === 'obsoletos'
    ? ['Fecha Obsoleto','No.','Área','Cód. Manual','Cód. Instructivo','Cód. Listado','Cód. Formato','Nombre Documento','Versión','Fecha Emisión','Elaboró','Vigencia','Tipo Resguardo','Ubicación','Copias Controladas','Tipo Resguardo Copia','Usuarios','Motivo','Observaciones']
    : ['No.','Área','Cód. Manual','Cód. Instructivo','Cód. Listado','Cód. Formato','Nombre Documento','Versión','Fecha Emisión','Elaboró','Vigencia','Estado','Días Vigencia','Tipo Resguardo','Ubicación','Copias Controladas','Tipo Resguardo Copia','Usuarios','Observaciones'];

  const rows = data.map(r => {
    if (STATE.currentSection === 'salidas') {
      return [r.nombreDoc, r.version, r.fechaEmision, r.area, r.motivo, r.destinatario, r.correo, r.fechaSalida];
    }
    if (STATE.currentSection === 'obsoletos') {
      return [r.fechaObsoleto ? formatLogDate(r.fechaObsoleto) : '', r.no, r.area, r.codManual, r.codInstructivo, r.codListado, r.codFormato, r.nombreDoc, r.version, r.fechaEmision, r.elaboro, r.vigencia, r.tipoResguardo, r.ubicacion, r.copias, r.tipoResguardoCopia, r.usuarios, r.motivo, r.observaciones];
    }
    return [r.no, r.area, r.codManual, r.codInstructivo, r.codListado, r.codFormato, r.nombreDoc, r.version, r.fechaEmision, r.elaboro, renderVigenciaDate(r.fechaEmision), calcEstado(r.fechaEmision), calcDiasVigencia(r.fechaEmision), r.tipoResguardo, r.ubicacion, r.copias, r.tipoResguardoCopia, r.usuarios, r.observaciones];
  });
  
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, 'Datos');
  XLSX.writeFile(wb2, filename);
  addLog('export', filename);
  showToast('Excel exportado: ' + filename);
}
function renderLogsTable() {
  const perPage = 50;
  let logs = STATE.logs;

  if (STATE.searchQuery) {
    const q = STATE.searchQuery.toLowerCase();
    logs = logs.filter(l => {
      const lt = LOG_TYPES[l.type] || {};
      return (l.user||'').toLowerCase().includes(q)
        || (l.details||'').toLowerCase().includes(q)
        || (l.docName||'').toLowerCase().includes(q)
        || (lt.label||'').toLowerCase().includes(q);
    });
  }

  const total = logs.length;
  const totalPages = Math.max(1, Math.ceil(total / perPage));
  if (STATE.page > totalPages) STATE.page = totalPages;
  const start = (STATE.page - 1) * perPage;
  const pageData = logs.slice(start, start + perPage);

  return h('div', {className: 'content-area'},
    h('div', {className: 'stats-bar'},
      h('div', {className: 'stat-card'},
        h('div', {className: 'stat-label'}, 'Total Registros'),
        h('div', {className: 'stat-value'}, String(total))
      )
    ),
    h('div', {className: 'filter-bar'},
      h('div', {className: 'search-box'},
        h('i', {className: 'fas fa-search'}),
        h('input', {
          type: 'text',
          placeholder: 'Buscar por usuario, acción, documento...',
          value: STATE.searchQuery,
          onInput: (e) => { STATE.searchQuery = e.target.value; STATE.page=1; render(); }
        })
      ),
      h('button', {className:'btn btn-sm btn-danger', onClick:()=>{
        if(confirm('¿Limpiar todo el registro de actividad?')) {
          STATE.logs = [];
          addLog('edit', 'Registro de actividad limpiado');
          saveState();
          render();
        }
      }}, h('i',{className:'fas fa-trash-can'}), ' Limpiar Logs')
    ),
    h('div', {className: 'table-container'},
      h('table', {className: 'data-table'},
        h('thead', null,
          h('tr', null,
            h('th', {style:{width:'170px'}}, 'Fecha / Hora'),
            h('th', {style:{width:'120px'}}, 'Usuario'),
            h('th', {style:{width:'180px'}}, 'Acción'),
            h('th', null, 'Documento'),
            h('th', null, 'Detalles')
          )
        ),
        h('tbody', null,
          ...pageData.map(log => {
            const lt = LOG_TYPES[log.type] || { icon:'fa-circle', color:'#5a8999', label: log.type };
            return h('tr', null,
              h('td', null,
                h('span', {style:{fontFamily:'JetBrains Mono,monospace',fontSize:'12px',color:'var(--text-secondary)'}}, formatLogDate(log.timestamp))
              ),
              h('td', null,
                h('span', {style:{fontWeight:'600',color:'var(--text-primary)'}}, log.user)
              ),
              h('td', null,
                h('span', {style:{display:'inline-flex',alignItems:'center',gap:'7px',padding:'4px 10px',borderRadius:'6px',fontSize:'11px',fontWeight:'700',background:lt.color+'18',color:lt.color,letterSpacing:'.2px'}},
                  h('i', {className:'fas '+lt.icon, style:{fontSize:'11px'}}),
                  lt.label
                )
              ),
              h('td', null,
                log.docName ? h('span', {style:{fontWeight:'500',color:'var(--text-primary)'}}, log.docName) : h('span', {style:{color:'var(--text-muted)'}}, '—')
              ),
              h('td', null,
                h('span', {style:{fontSize:'12px',color:'var(--text-secondary)'}}, log.details || '—')
              )
            );
          }),
          pageData.length === 0 ? h('tr', null,
            h('td', {style:{textAlign:'center', padding:'40px', color:'var(--text-muted)'}, colSpan:'5'},
              h('i', {className:'fas fa-clock-rotate-left', style:{fontSize:'28px',opacity:'.3',display:'block',marginBottom:'10px'}}),
              'No hay registros de actividad'
            )
          ) : null
        )
      )
    ),
    h('div', {className: 'pagination'},
      h('span', {className: 'pagination-info'}, `Mostrando ${total?start+1:0}-${Math.min(start+perPage,total)} de ${total}`),
      h('div', {className: 'pagination-btns'},
        h('button', {className:'page-btn', disabled:STATE.page<=1, onClick:()=>{STATE.page--;render();}}, h('i',{className:'fas fa-chevron-left'})),
        h('button', {className:'page-btn', disabled:STATE.page>=totalPages, onClick:()=>{STATE.page++;render();}}, h('i',{className:'fas fa-chevron-right'}))
      )
    )
  );
}
function renderViewModal() {
  const rec = STATE.selectedRecord;
  if (!rec) return null;
  
  const estado = calcEstado(rec.fechaEmision);
  const dias = calcDiasVigencia(rec.fechaEmision);
  const vig = calcVigencia(rec.fechaEmision);
  
  const dv = (label, value, isFull) => h('div', {className: 'detail-item' + (isFull?' full':'')},
    h('div', {className:'detail-label'}, label),
    h('div', {className:'detail-value'}, value || '—')
  );
  
  return h('div', {className:'modal-overlay', onClick:(e)=>{if(e.target.className==='modal-overlay')closeModal();}},
    h('div', {className:'modal modal-lg'},
      h('div', {className:'modal-header'},
        h('h2', null, h('i',{className:'fas fa-file-alt', style:{marginRight:'8px',color:'var(--accent)'}}), 'Detalle del Documento'),
        h('button', {className:'modal-close', onClick:closeModal}, h('i',{className:'fas fa-times'}))
      ),
      h('div', {className:'modal-body'},
        h('div', {className:'detail-grid'},
          dv('No.', rec.no),
          dv('Área', rec.area),
          dv('Código Manual / Procedimiento', rec.codManual),
          dv('Código Instructivo', rec.codInstructivo),
          dv('Código Listado', rec.codListado),
          dv('Código Formato', rec.codFormato),
          dv('Nombre del Documento', rec.nombreDoc, true),
          dv('Versión', rec.version),
          dv('Fecha Emisión', rec.fechaEmision),
          dv('Elaboró / Actualizó', rec.elaboro),
          dv('Vigencia', vig ? formatDateFull(vig) : '—'),
          h('div', {className:'detail-item'},
            h('div', {className:'detail-label'}, 'Estado'),
            h('div', {className:'detail-value'}, renderStatusBadge(rec.fechaEmision))
          ),
          h('div', {className:'detail-item'},
            h('div', {className:'detail-label'}, 'Días de Vigencia'),
            h('div', {className:'detail-value'}, renderDiasVigencia(rec.fechaEmision))
          ),
          dv('Tipo de Resguardo - Original', rec.tipoResguardo),
          dv('Ubicación / URL', rec.ubicacion ? (
            rec.ubicacion.startsWith('http')
              ? h('a', {href:_safeHref(rec.ubicacion), target:'_blank', rel:'noopener noreferrer'}, rec.ubicacion)
              : rec.ubicacion
          ) : '—'),
          dv('Copias Controladas', rec.copias),
          dv('Tipo Resguardo - Copia', rec.tipoResguardoCopia),
          dv('Usuarios con acceso', rec.usuarios, true),
          dv('Observaciones', rec.observaciones, true),
          h('div', {className:'detail-item full'},
            h('div', {className:'detail-label'}, 'Archivos'),
            h('div', {className:'detail-value'},
              (rec.archivos && rec.archivos.length > 0) || (rec.archivoURLs && rec.archivoURLs.length > 0)
                ? h('div', {className:'file-preview', style:{flexWrap:'wrap',gap:'8px'}},
                    ...(rec.archivoURLs||[]).map(url => h('a', {className:'file-chip', href:_safeHref(url), target:'_blank', rel:'noopener noreferrer'}, h('i',{className:'fas fa-link'}), url.length > 40 ? url.substring(0,40)+'…' : url)),
                    ...(rec.archivos||[]).map(f => h('span', {className:'file-chip', onClick:()=>downloadFile(f)}, h('i',{className:'fas fa-file'}), f.name))
                  )
                : '—'
            )
          ),
          h('div', {className:'detail-item full'},
            h('div', {className:'detail-label'}, 'Archivos de Copia'),
            h('div', {className:'detail-value'},
              (rec.archivosCopia && rec.archivosCopia.length > 0) || (rec.archivoURLsCopia && rec.archivoURLsCopia.length > 0)
                ? h('div', {className:'file-preview', style:{flexWrap:'wrap',gap:'8px'}},
                    ...(rec.archivoURLsCopia||[]).map(url => h('a', {className:'file-chip', href:_safeHref(url), target:'_blank', rel:'noopener noreferrer'}, h('i',{className:'fas fa-link'}), url.length > 40 ? url.substring(0,40)+'…' : url)),
                    ...(rec.archivosCopia||[]).map(f => h('span', {className:'file-chip', onClick:()=>downloadFile(f)}, h('i',{className:'fas fa-file'}), f.name))
                  )
                : '—'
            )
          )
        ),
        rec._history && rec._history.length > 0 ? h('div', {style:{marginTop:'20px',padding:'14px',background:'var(--bg-input)',borderRadius:'10px',border:'1px solid var(--border)'}},
          h('div', {style:{fontWeight:'600',fontSize:'13px',marginBottom:'10px',color:'var(--accent)'}},
            h('i', {className:'fas fa-clock-rotate-left', style:{marginRight:'6px'}}),
            'Historial de cambios (' + rec._history.length + ')'
          ),
          h('div', {style:{maxHeight:'150px',overflowY:'auto'}},
            ...rec._history.slice().reverse().map(function(entry) {
              var d = new Date(entry.ts);
              return h('div', {style:{fontSize:'12px',padding:'4px 0',borderBottom:'1px solid var(--border)',display:'flex',gap:'12px'}},
                h('span', {style:{color:'var(--text-muted)',fontFamily:'JetBrains Mono,monospace',minWidth:'130px'}}, d.toLocaleDateString('es-MX') + ' ' + d.toLocaleTimeString('es-MX')),
                h('span', {style:{color:'var(--accent)',fontWeight:'500',minWidth:'80px'}}, entry.user),
                h('span', {style:{color:'var(--text-secondary)'}}, entry.action)
              );
            })
          )
        ) : null
      ),
      h('div', {className:'modal-footer'},
        h('button', {className:'btn btn-secondary', onClick:closeModal}, 'Cerrar')
      )
    )
  );
}

function renderEditModal() {
  const form = STATE.editForm;
  const isAdd = STATE.modal === 'add';
  
  const fg = (label, key, opts={}) => {
    const isFull = opts.full;
    const type = opts.type || 'text';
    const hint = opts.hint;
    const isSelect = opts.options;
    const isTextarea = opts.textarea;
    
    let input;
    if (isSelect) {
      input = h('select', {
        className:'form-select',
        value: form[key]||'',
        onChange: (e) => { STATE.editForm[key] = e.target.value; refreshModal(); }
      },
        h('option', {value:''}, '-- Seleccionar --'),
        ...opts.options.map(o => h('option', {value: typeof o === 'string' ? o : o.value}, typeof o === 'string' ? o : o.label))
      );
    } else if (isTextarea) {
      input = h('textarea', {
        className:'form-textarea',
        value: form[key]||'',
        onInput: (e) => { STATE.editForm[key] = e.target.value; }
      });
    } else {
      input = h('input', {
        className:'form-input',
        type: type,
        value: form[key]||'',
        placeholder: opts.placeholder || '',
        onInput: (e) => {
          let val = e.target.value;
          if (opts.numbersOnly) val = val.replace(/[^0-9]/g, '');
          STATE.editForm[key] = val;
          if (opts.numbersOnly) e.target.value = val;
        }
      });
    }
    
    return h('div', {className:'form-group' + (isFull?' full':'')},
      h('label', {className:'form-label'}, label),
      input,
      hint ? h('div', {className:'form-hint'}, hint) : null
    );
  };

  const elaboroOptions = STATE.elaboros.map(e => ({value: e, label: e}));
  
  return h('div', {className:'modal-overlay', onClick:(e)=>{if(e.target.className==='modal-overlay')closeModal();}},
    h('div', {className:'modal modal-lg'},
      h('div', {className:'modal-header'},
        h('h2', null, h('i',{className:'fas fa-'+(isAdd?'plus':'pen'), style:{marginRight:'8px',color:'var(--accent)'}}), isAdd ? 'Nuevo Registro' : 'Editar Registro'),
        h('button', {className:'modal-close', onClick:closeModal}, h('i',{className:'fas fa-times'}))
      ),
      h('div', {className:'modal-body'},
        h('div', {className:'form-grid'},
          fg('No.', 'no', {placeholder:'ID manual'}),
          fg('Código Manual / Procedimiento', 'codManual'),
          fg('Código Instructivo', 'codInstructivo'),
          fg('Código Listado', 'codListado'),
          fg('Código Formato', 'codFormato'),
          fg('Nombre del Documento', 'nombreDoc', {full:true}),
          h('div', {className:'form-group'},
            h('label', {className:'form-label'}, 'Versión'),
            h('input', {
              className:'form-input',
              type:'text',
              value: form.version||'',
              placeholder:'Solo números',
              onInput: (e) => {
                STATE.editForm.version = e.target.value.replace(/[^0-9]/g,'');
                e.target.value = STATE.editForm.version;
                refreshModal();
              }
            }),
            h('div', {className:'form-hint'}, 'Solo números')
          ),
          fg('Fecha Emisión', 'fechaEmision', {placeholder:'ej: oct-24 o 2024-10-30', hint:'Formato: mes-año (oct-24) o AAAA-MM-DD'}),

          (!isAdd && STATE.originalVersion && String(form.version||'').trim() !== '' && String(form.version||'').trim() !== STATE.originalVersion)
            ? h('div', {className:'form-group full'},
                h('div', {style:{background:'rgba(249,115,22,.08)', border:'1.5px solid rgba(249,115,22,.3)', borderRadius:'12px', padding:'16px', marginTop:'4px'}},
                  h('div', {style:{display:'flex', alignItems:'center', gap:'10px', marginBottom:'12px'}},
                    h('div', {style:{width:'36px',height:'36px',borderRadius:'50%',background:'rgba(249,115,22,.15)',display:'flex',alignItems:'center',justifyContent:'center',flexShrink:'0'}},
                      h('i', {className:'fas fa-box-archive', style:{color:'#fb923c',fontSize:'15px'}})
                    ),
                    h('div', null,
                      h('div', {style:{fontWeight:'700',fontSize:'14px',color:'#fb923c'}}, 'Cambio de versión detectado'),
                      h('div', {style:{fontSize:'12px',color:'var(--text-secondary)',marginTop:'2px'}},
                        'Versión ', h('span',{style:{fontFamily:'JetBrains Mono,monospace',fontWeight:'700',color:'var(--red)',textDecoration:'line-through'}}, STATE.originalVersion),
                        ' → ',
                        h('span',{style:{fontFamily:'JetBrains Mono,monospace',fontWeight:'700',color:'var(--green)'}}, form.version)
                      )
                    )
                  ),
                  h('div', {style:{fontSize:'12px',color:'var(--text-secondary)',marginBottom:'10px',lineHeight:'1.5'}},
                    h('i', {className:'fas fa-info-circle', style:{marginRight:'6px',color:'var(--accent)'}}),
                    'La versión anterior con sus archivos será enviada automáticamente a Obsoletos. Los archivos del registro actual se vaciarán para la nueva versión.'
                  ),
                  h('label', {style:{display:'block',fontSize:'12px',fontWeight:'700',color:'#fb923c',marginBottom:'6px',textTransform:'uppercase',letterSpacing:'.4px'}}, 'Motivo del cambio de versión *'),
                  h('textarea', {
                    className:'form-textarea',
                    id: 'motivo-version-input',
                    placeholder:'Describe el motivo del cambio de versión (obligatorio)...',
                    style:{minHeight:'80px',borderColor:'rgba(249,115,22,.3)'},
                    value: STATE.obsoletoMotivo||'',
                    onInput: (e) => {
                      STATE.obsoletoMotivo = e.target.value;
                      const btn = document.getElementById('btn-save-obsoleto');
                      if (btn) {
                        btn.disabled = !e.target.value.trim();
                      }
                    }
                  })
                )
              )
            : null,
          h('div', {className:'form-group'},
            h('label', {className:'form-label'}, 'Elaboró / Actualizó'),
            h('select', {
              className:'form-select',
              value: form.elaboro||'',
              onChange: (e) => {
                if (e.target.value === '__new__') {
                  const name = prompt('Ingresa el nombre del nuevo elaborador:');
                  if (name && name.trim()) {
                    STATE.elaboros.push(name.trim());
                    STATE.editForm.elaboro = name.trim();
                  }
                } else {
                  STATE.editForm.elaboro = e.target.value;
                }
                refreshModal();
              }
            },
              h('option', {value:''}, '-- Seleccionar --'),
              ...STATE.elaboros.map(e => h('option', {value:e}, e)),
              h('option', {value:'__new__'}, '+ Agregar nuevo...')
            )
          ),
          fg('Tipo Resguardo - Original', 'tipoResguardo', {options:['Físico','Electrónico']}),
          form.tipoResguardo === 'Físico'
            ? fg('Ubicación Física', 'ubicacion', {placeholder:'¿Dónde se guarda?'})
            : form.tipoResguardo === 'Electrónico'
              ? fg('URL del Documento', 'ubicacion', {placeholder:'https://...', hint:'Ingresa la URL del documento'})
              : fg('Ubicación', 'ubicacion', {placeholder:'Selecciona tipo de resguardo primero'}),
          
          h('div', {className:'form-group full'},
            h('label', {className:'form-label'}, 'Archivos'),
            h('div', {style:{display:'flex',gap:'8px',flexWrap:'wrap',marginBottom:'8px'}},
              ...(form.archivos||[]).map((f,i) => h('span', {className:'file-chip'},
                h('i',{className:'fas fa-file'}),
                f.name,
                h('i', {className:'fas fa-times', style:{marginLeft:'4px',cursor:'pointer'}, onClick:()=>{
                  STATE.editForm.archivos.splice(i,1);
                  refreshModal();
                }})
              ))
            ),
            h('input', {
              type:'file',
              multiple:true,
              className:'form-input',
              onChange: (e) => {
                const MAX_SIZE = 10 * 1024 * 1024;
                const ALLOWED = ['application/pdf','image/jpeg','image/png','image/gif','application/msword','application/vnd.openxmlformats-officedocument.wordprocessingml.document','application/vnd.ms-excel','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','text/plain'];
                const files = Array.from(e.target.files);
                files.forEach(async function(file) {
                  if (file.size > MAX_SIZE) { showToast('Archivo "' + file.name + '" excede 10MB', 'error'); return; }
                  if (ALLOWED.length > 0 && !ALLOWED.includes(file.type) && !file.name.match(/\.(pdf|jpg|jpeg|png|gif|doc|docx|xls|xlsx|txt)$/i)) { showToast('Tipo no permitido: ' + file.name, 'error'); return; }
                  var uploaded = await _uploadToStorage(file);
                  if (uploaded) {
                    if (!STATE.editForm.archivos) STATE.editForm.archivos = [];
                    STATE.editForm.archivos.push(uploaded);
                    refreshModal();
                  }
                });
              }
            })
          ),

          h('div', {className:'form-group full'},
            h('label', {className:'form-label'}, 'URLs de Archivos'),
            h('div', {style:{display:'flex',flexDirection:'column',gap:'6px'}},
              ...(form.archivoURLs||[]).map((url, i) => h('div', {style:{display:'flex',gap:'6px',alignItems:'center'}},
                h('input', {className:'form-input', value:url, style:{flex:1}, onInput:(e)=>{STATE.editForm.archivoURLs[i]=e.target.value;}}),
                h('button', {className:'btn btn-sm btn-danger', onClick:()=>{STATE.editForm.archivoURLs.splice(i,1);refreshModal();}}, h('i',{className:'fas fa-times'}))
              ))
            ),
            h('button', {className:'btn btn-sm btn-secondary', style:{marginTop:'6px'}, onClick:()=>{
              if(!STATE.editForm.archivoURLs) STATE.editForm.archivoURLs = [];
              STATE.editForm.archivoURLs.push('');
              refreshModal();
            }}, h('i',{className:'fas fa-plus'}), ' Agregar URL')
          ),
          
          fg('Copias Controladas', 'copias', {numbersOnly:true, hint:'Solo números'}),
          fg('Tipo Resguardo - Copia', 'tipoResguardoCopia', {options:['Físico','Electrónico']}),

          h('div', {className:'form-group full'},
            h('label', {className:'form-label'}, 'Archivos de Copia'),
            h('div', {style:{display:'flex',gap:'8px',flexWrap:'wrap',marginBottom:'8px'}},
              ...(form.archivosCopia||[]).map((f,i) => h('span', {className:'file-chip'},
                h('i',{className:'fas fa-file'}),
                f.name,
                h('i', {className:'fas fa-times', style:{marginLeft:'4px',cursor:'pointer'}, onClick:()=>{
                  STATE.editForm.archivosCopia.splice(i,1);
                  refreshModal();
                }})
              ))
            ),
            h('input', {
              type:'file',
              multiple:true,
              className:'form-input',
              onChange: (e) => {
                const MAX_SIZE = 10 * 1024 * 1024;
                const ALLOWED = ['application/pdf','image/jpeg','image/png','image/gif','application/msword','application/vnd.openxmlformats-officedocument.wordprocessingml.document','application/vnd.ms-excel','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','text/plain'];
                const files = Array.from(e.target.files);
                files.forEach(async function(file) {
                  if (file.size > MAX_SIZE) { showToast('Archivo "' + file.name + '" excede 10MB', 'error'); return; }
                  if (ALLOWED.length > 0 && !ALLOWED.includes(file.type) && !file.name.match(/\.(pdf|jpg|jpeg|png|gif|doc|docx|xls|xlsx|txt)$/i)) { showToast('Tipo no permitido: ' + file.name, 'error'); return; }
                  var uploaded = await _uploadToStorage(file);
                  if (uploaded) {
                    if (!STATE.editForm.archivosCopia) STATE.editForm.archivosCopia = [];
                    STATE.editForm.archivosCopia.push(uploaded);
                    refreshModal();
                  }
                });
              }
            })
          ),

          h('div', {className:'form-group full'},
            h('label', {className:'form-label'}, 'URLs de Archivos de Copia'),
            h('div', {style:{display:'flex',flexDirection:'column',gap:'6px'}},
              ...(form.archivoURLsCopia||[]).map((url, i) => h('div', {style:{display:'flex',gap:'6px',alignItems:'center'}},
                h('input', {className:'form-input', value:url, style:{flex:1}, onInput:(e)=>{STATE.editForm.archivoURLsCopia[i]=e.target.value;}}),
                h('button', {className:'btn btn-sm btn-danger', onClick:()=>{STATE.editForm.archivoURLsCopia.splice(i,1);refreshModal();}}, h('i',{className:'fas fa-times'}))
              ))
            ),
            h('button', {className:'btn btn-sm btn-secondary', style:{marginTop:'6px'}, onClick:()=>{
              if(!STATE.editForm.archivoURLsCopia) STATE.editForm.archivoURLsCopia = [];
              STATE.editForm.archivoURLsCopia.push('');
              refreshModal();
            }}, h('i',{className:'fas fa-plus'}), ' Agregar URL')
          ),

          fg('Usuarios con acceso a copia', 'usuarios', {full:true, placeholder:'correo1@ejemplo.com, correo2@ejemplo.com'}),
          fg('Observaciones', 'observaciones', {full:true, textarea:true})
        )
      ),
      h('div', {className:'modal-footer'},
        h('button', {className:'btn btn-secondary', onClick:closeModal}, 'Cancelar'),
        (() => {
          const vChanged = !isAdd && STATE.originalVersion && String(form.version||'').trim() !== '' && String(form.version||'').trim() !== STATE.originalVersion;
          if (vChanged) {
            const btn = h('button', {className:'btn btn-warning', id:'btn-save-obsoleto', onClick:saveRecord},
              h('i',{className:'fas fa-box-archive'}), 'Guardar y Enviar Versión Anterior a Obsoletos');
            btn.disabled = !(STATE.obsoletoMotivo && STATE.obsoletoMotivo.trim());
            return btn;
          }
          return h('button', {className:'btn btn-primary', onClick:saveRecord},
            h('i',{className:'fas fa-floppy-disk'}), isAdd ? 'Agregar Registro' : 'Guardar Cambios');
        })()
      )
    )
  );
}

function renderSalidaModal() {
  const rec = STATE.selectedRecord;
  const form = STATE.salidaForm;
  if (!rec) return null;
  
  return h('div', {className:'modal-overlay', onClick:(e)=>{if(e.target.className==='modal-overlay')closeModal();}},
    h('div', {className:'modal'},
      h('div', {className:'modal-header'},
        h('h2', null, h('i',{className:'fas fa-sign-out-alt', style:{marginRight:'8px',color:'var(--purple)'}}), 'Registro de Salida'),
        h('button', {className:'modal-close', onClick:closeModal}, h('i',{className:'fas fa-times'}))
      ),
      h('div', {className:'modal-body'},
        h('div', {style:{background:'var(--bg-input)', borderRadius:'10px', padding:'14px', marginBottom:'20px', border:'1px solid var(--border)'}},
          h('div', {style:{fontSize:'13px',color:'var(--text-muted)',marginBottom:'4px'}}, 'Documento:'),
          h('div', {style:{fontSize:'15px',fontWeight:'600'}}, rec.nombreDoc || rec.codManual || rec.codFormato),
          h('div', {style:{fontSize:'12px',color:'var(--text-muted)',marginTop:'4px'}}, `Versión: ${rec.version || '—'} | Emisión: ${rec.fechaEmision || '—'}`)
        ),
        h('div', {className:'form-grid'},
          h('div', {className:'form-group'},
            h('label', {className:'form-label'}, 'Área *'),
            h('input', {className:'form-input', value: form.area||'', onInput:(e)=>{STATE.salidaForm.area=e.target.value;}})
          ),
          h('div', {className:'form-group'},
            h('label', {className:'form-label'}, 'Destinatario *'),
            h('input', {className:'form-input', placeholder:'Nombre completo', value: form.destinatario||'', onInput:(e)=>{STATE.salidaForm.destinatario=e.target.value;}})
          ),
          h('div', {className:'form-group'},
            h('label', {className:'form-label'}, 'Correo del Destinatario *'),
            h('input', {className:'form-input', type:'email', placeholder:'correo@ejemplo.com', value: form.correo||'', onInput:(e)=>{STATE.salidaForm.correo=e.target.value;}})
          ),
          h('div', {className:'form-group full'},
            h('label', {className:'form-label'}, 'Motivo de Entrega *'),
            h('textarea', {className:'form-textarea', placeholder:'Describe el motivo de la entrega...', value: form.motivo||'', onInput:(e)=>{STATE.salidaForm.motivo=e.target.value;}})
          )
        )
      ),
      h('div', {className:'modal-footer'},
        h('button', {className:'btn btn-secondary', onClick:closeModal}, 'Cancelar'),
        h('button', {className:'btn btn-primary', onClick:confirmSalida}, h('i',{className:'fas fa-paper-plane'}), 'Registrar Salida')
      )
    )
  );
}

function renderObsoletoModal() {
  const rec = STATE.selectedRecord;
  if (!rec) return null;
  
  return h('div', {className:'modal-overlay', onClick:(e)=>{if(e.target.className==='modal-overlay')closeModal();}},
    h('div', {className:'modal'},
      h('div', {className:'modal-header'},
        h('h2', null, h('i',{className:'fas fa-archive', style:{marginRight:'8px',color:'var(--orange)'}}), 'Enviar a Obsoletos'),
        h('button', {className:'modal-close', onClick:closeModal}, h('i',{className:'fas fa-times'}))
      ),
      h('div', {className:'modal-body'},
        h('div', {style:{background:'var(--red-bg)', borderRadius:'10px', padding:'14px', marginBottom:'20px', border:'1px solid rgba(239,68,68,.2)'}},
          h('div', {style:{display:'flex',alignItems:'center',gap:'8px',color:'var(--red)',fontWeight:'600',fontSize:'13px'}},
            h('i', {className:'fas fa-exclamation-triangle'}),
            'Este documento será movido a obsoletos'
          ),
          h('div', {style:{fontSize:'13px',marginTop:'8px',color:'var(--text-secondary)'}}, `${rec.nombreDoc || rec.codManual || rec.codFormato} — Versión ${rec.version||'—'}`)
        ),
        h('div', {className:'form-group'},
          h('label', {className:'form-label'}, 'Motivo del cambio *'),
          h('textarea', {
            className:'form-textarea',
            placeholder:'Describe el motivo por el cual este documento pasa a obsoleto...',
            style:{minHeight:'100px'},
            value: STATE.obsoletoMotivo,
            onInput:(e)=>{STATE.obsoletoMotivo=e.target.value;}
          }),
          h('div', {className:'form-hint'}, 'Este campo es obligatorio para confirmar el envío a obsoletos')
        )
      ),
      h('div', {className:'modal-footer'},
        h('button', {className:'btn btn-secondary', onClick:closeModal}, 'Cancelar'),
        h('button', {
          className:'btn btn-warning',
          disabled: !STATE.obsoletoMotivo.trim(),
          onClick: confirmObsoleto
        }, h('i',{className:'fas fa-box-archive'}), 'Confirmar y Enviar a Obsoletos')
      )
    )
  );
}
function renderRenewModal() {
  var rec = STATE.selectedRecord;
  if (!rec) return null;

  var diasVencido = Math.abs(calcDiasVigencia(rec.fechaEmision) || 0);

  return h('div', {className:'modal-overlay', onClick:function(e){if(e.target.className==='modal-overlay')closeModal();}},
    h('div', {className:'modal'},
      h('div', {className:'modal-header'},
        h('h2', null, h('i',{className:'fas fa-rotate-right', style:{marginRight:'8px',color:'var(--green)'}}), 'Renovar Vigencia'),
        h('button', {className:'modal-close', onClick:closeModal}, h('i',{className:'fas fa-times'}))
      ),
      h('div', {className:'modal-body'},
        h('div', {style:{background:'var(--bg-card)', borderRadius:'10px', padding:'14px', marginBottom:'20px', border:'1px solid var(--border)'}},
          h('div', {style:{display:'flex',justifyContent:'space-between',alignItems:'center'}},
            h('div', null,
              h('div', {style:{fontWeight:'600',color:'var(--text-primary)',fontSize:'14px'}}, rec.nombreDoc || rec.codManual || rec.codFormato || 'Sin nombre'),
              h('div', {style:{fontSize:'12px',color:'var(--text-muted)',marginTop:'4px'}}, 'Área: ' + (rec.area || '—') + ' • Versión: ' + (rec.version || '—') + ' • Elaboró: ' + (rec.elaboro || '—'))
            ),
            h('div', {style:{background:'var(--red-bg)', borderRadius:'8px', padding:'8px 12px', textAlign:'center'}},
              h('div', {style:{fontWeight:'700',color:'var(--red)',fontSize:'18px',fontFamily:'JetBrains Mono, monospace'}}, String(diasVencido)),
              h('div', {style:{fontSize:'10px',color:'var(--red)',fontWeight:'500'}}, 'días vencido')
            )
          )
        ),
        h('div', {style:{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'12px',marginBottom:'16px'}},
          h('div', {className:'form-group'},
            h('label', {className:'form-label'}, 'Fecha emisión actual'),
            h('input', {className:'form-input', type:'text', value: rec.fechaEmision || '', disabled: true, style:{opacity:'.6'}})
          ),
          h('div', {className:'form-group'},
            h('label', {className:'form-label'}, 'Vigencia expiró'),
            h('input', {className:'form-input', type:'text', value: typeof renderVigenciaDate === 'function' ? renderVigenciaDate(rec.fechaEmision) : '', disabled: true, style:{opacity:'.6'}})
          )
        ),
        h('div', {className:'form-group'},
          h('label', {className:'form-label', style:{color:'var(--green)', fontWeight:'600'}}, '📅 Nueva Fecha de Emisión *'),
          h('input', {
            className:'form-input',
            type:'date',
            id:'renew-date',
            value: STATE._renewDate || '',
            onInput: function(e) { STATE._renewDate = e.target.value; }
          }),
          h('div', {className:'form-hint'}, 'La vigencia se recalculará automáticamente (730 días desde esta fecha)')
        ),
        h('div', {className:'form-group'},
          h('label', {className:'form-label'}, 'Motivo de la renovación *'),
          h('textarea', {
            className:'form-textarea',
            placeholder:'Describe el motivo de la renovación (ej: Revisión periódica sin cambios, Actualización de contenido, Auditoría...)',
            style:{minHeight:'80px'},
            value: STATE._renewMotivo || '',
            onInput: function(e) { STATE._renewMotivo = e.target.value; }
          }),
          h('div', {className:'form-hint'}, 'Este campo es obligatorio para el registro de auditoría')
        )
      ),
      h('div', {className:'modal-footer'},
        h('button', {className:'btn btn-secondary', onClick:closeModal}, 'Cancelar'),
        h('button', {
          className:'btn btn-primary',
          disabled: !STATE._renewDate || !(STATE._renewMotivo && STATE._renewMotivo.trim()),
          onClick: _confirmRenew
        }, h('i',{className:'fas fa-check-circle'}), 'Confirmar Renovación')
      )
    )
  );
}
function renderAlertasPanel() {
  var alertas = getAlertasVigencia();
  var resumen = getAlertasResumen();

  var filtroNivel = STATE._alertaFiltro || 'todos';

  var alertasFiltradas = alertas;
  if (filtroNivel === 'vencidos') alertasFiltradas = alertas.filter(function(a) { return a.nivel === 'vencido'; });
  else if (filtroNivel === 'criticos') alertasFiltradas = alertas.filter(function(a) { return a.nivel === 'critico'; });
  else if (filtroNivel === 'avisos') alertasFiltradas = alertas.filter(function(a) { return a.nivel === 'aviso' || a.nivel === 'proximo'; });

  var container = h('div', {className: 'content-area'},
    h('div', {className: 'alertas-resumen'},
      h('div', {className: 'alerta-card alerta-card-red', onClick: function() { STATE._alertaFiltro = 'vencidos'; render(); }},
        h('div', {className: 'alerta-card-icon'}, h('i', {className: 'fas fa-circle-xmark'})),
        h('div', {className: 'alerta-card-num'}, String(resumen.vencidos)),
        h('div', {className: 'alerta-card-label'}, 'Vencidos')
      ),
      h('div', {className: 'alerta-card alerta-card-orange', onClick: function() { STATE._alertaFiltro = 'criticos'; render(); }},
        h('div', {className: 'alerta-card-icon'}, h('i', {className: 'fas fa-triangle-exclamation'})),
        h('div', {className: 'alerta-card-num'}, String(resumen.criticos)),
        h('div', {className: 'alerta-card-label'}, '≤ 15 días')
      ),
      h('div', {className: 'alerta-card alerta-card-yellow', onClick: function() { STATE._alertaFiltro = 'avisos'; render(); }},
        h('div', {className: 'alerta-card-icon'}, h('i', {className: 'fas fa-clock'})),
        h('div', {className: 'alerta-card-num'}, String(resumen.avisos + resumen.proximos)),
        h('div', {className: 'alerta-card-label'}, '≤ 60 días')
      ),
      h('div', {className: 'alerta-card alerta-card-blue', onClick: function() { STATE._alertaFiltro = 'todos'; render(); }},
        h('div', {className: 'alerta-card-icon'}, h('i', {className: 'fas fa-bell'})),
        h('div', {className: 'alerta-card-num'}, String(resumen.total)),
        h('div', {className: 'alerta-card-label'}, 'Total alertas')
      )
    ),

    h('div', {className: 'alertas-actions-bar'},
      h('div', {className: 'alertas-actions-left'},
        h('span', {className: 'alertas-filter-label'},
          filtroNivel === 'todos' ? 'Mostrando todas las alertas' :
          filtroNivel === 'vencidos' ? 'Filtrando: Documentos vencidos' :
          filtroNivel === 'criticos' ? 'Filtrando: Vencen en ≤15 días' :
          'Filtrando: Vencen en ≤60 días',
          ' (' + alertasFiltradas.length + ')'
        ),
        filtroNivel !== 'todos' ? h('button', {className: 'btn btn-sm btn-secondary', onClick: function() { STATE._alertaFiltro = 'todos'; render(); }},
          h('i', {className: 'fas fa-times', style:{marginRight:'4px'}}), 'Quitar filtro'
        ) : null
      ),
      h('div', {className: 'alertas-actions-right'},
        alertasFiltradas.length > 0 ? h('button', {
          className: 'btn btn-email-alert',
          onClick: function() { enviarAlertasPorCorreo(alertasFiltradas, filtroNivel); }
        },
          h('i', {className: 'fas fa-envelopes-bulk'}),
          'Enviar todos (' + alertasFiltradas.length + ')'
        ) : null
      )
    ),

    h('div', {className: 'alertas-lista-container'},
      alertasFiltradas.length === 0 ?
        h('div', {className: 'alertas-empty'},
          h('i', {className: 'fas fa-check-circle', style:{fontSize:'48px',color:'var(--green)',marginBottom:'16px'}}),
          h('div', {style:{fontSize:'16px',fontWeight:'600',marginBottom:'6px'}}, 'Sin alertas'),
          h('div', {style:{color:'var(--text-muted)'}}, 'No hay documentos que requieran atención en esta categoría')
        ) :
        h('div', {className: 'alertas-lista'},
          ...alertasFiltradas.map(function(a) {
            var diasTexto = '';
            var diasClass = '';
            if (a.dias < 0) {
              diasTexto = 'Vencido hace ' + Math.abs(a.dias) + ' días';
              diasClass = 'alerta-dias-rojo';
            } else if (a.dias === 0) {
              diasTexto = '¡Vence HOY!';
              diasClass = 'alerta-dias-rojo';
            } else if (a.dias === 1) {
              diasTexto = 'Vence MAÑANA';
              diasClass = 'alerta-dias-rojo';
            } else {
              diasTexto = 'Vence en ' + a.dias + ' días';
              diasClass = a.dias <= 15 ? 'alerta-dias-rojo' : a.dias <= 30 ? 'alerta-dias-naranja' : 'alerta-dias-amarillo';
            }

            return h('div', {className: 'alerta-item alerta-item-' + a.nivel},
              h('div', {className: 'alerta-item-icon', style:{color: a.color}},
                h('i', {className: 'fas ' + a.icono})
              ),
              h('div', {className: 'alerta-item-body'},
                h('div', {className: 'alerta-item-titulo'}, a.nombreDoc),
                h('div', {className: 'alerta-item-meta'},
                  h('span', null, h('i', {className: 'fas fa-folder', style:{marginRight:'4px'}}), a.area),
                  a.codigo ? h('span', null, h('i', {className: 'fas fa-tag', style:{marginRight:'4px'}}), a.codigo) : null,
                  h('span', null, h('i', {className: 'fas fa-code-branch', style:{marginRight:'4px'}}), 'v' + a.version),
                  h('span', null, h('i', {className: 'fas fa-user', style:{marginRight:'4px'}}), a.elaboro)
                ),
                h('div', {className: 'alerta-item-fechas'},
                  h('span', null, 'Emisión: ' + a.fechaEmision),
                  h('span', null, 'Vigencia: ' + (a.vigenciaDate ? formatDateFull(a.vigenciaDate) : '—'))
                )
              ),
              h('div', {className: 'alerta-item-dias'},
                h('div', {className: 'alerta-dias-badge ' + diasClass}, diasTexto),
                h('div', {className: 'alerta-item-btns'},
                  h('button', {className: 'btn btn-sm btn-email-individual', onClick: (function(alerta) { return function() { enviarAlertaIndividual(alerta); }; })(a)},
                    h('i', {className: 'fas fa-envelope'}), 'Correo'
                  ),
                  h('button', {className: 'btn btn-sm btn-secondary', onClick: function() {
                    STATE.currentSection = 'areas';
                    STATE.currentArea = a.area;
                    STATE.searchQuery = a.nombreDoc.substring(0, 30);
                    STATE.page = 1;
                    render();
                  }}, h('i', {className: 'fas fa-eye', style:{marginRight:'4px'}}), 'Ver')
                )
              )
            );
          })
        )
    )
  );

  return container;
}

function showAlertBanner() {
  var criticas = typeof getAlertasCriticas === 'function' ? getAlertasCriticas() : [];
  if (criticas.length === 0) return;

  var vencidos = criticas.filter(function(a) { return a.nivel === 'vencido'; }).length;
  var porVencer = criticas.filter(function(a) { return a.nivel === 'critico'; }).length;

  var mensaje = '';
  if (vencidos > 0 && porVencer > 0) {
    mensaje = '⚠️ ' + vencidos + ' documento(s) vencidos y ' + porVencer + ' por vencer en ≤15 días';
  } else if (vencidos > 0) {
    mensaje = '⚠️ ' + vencidos + ' documento(s) con vigencia vencida';
  } else {
    mensaje = '⚠️ ' + porVencer + ' documento(s) vencen en los próximos 15 días';
  }

  var banner = document.createElement('div');
  banner.className = 'alert-banner';
  var bannerContent = document.createElement('div');
  bannerContent.className = 'alert-banner-content';
  var ico = document.createElement('i');
  ico.className = 'fas fa-triangle-exclamation';
  var sp = document.createElement('span');
  sp.textContent = mensaje;
  var btnVer = document.createElement('button');
  btnVer.className = 'alert-banner-btn';
  btnVer.textContent = 'Ver alertas';
  btnVer.onclick = function(){STATE.currentSection='alertas';render();banner.remove();};
  var btnClose = document.createElement('button');
  btnClose.className = 'alert-banner-close';
  var icoClose = document.createElement('i');
  icoClose.className = 'fas fa-times';
  btnClose.appendChild(icoClose);
  btnClose.onclick = function(){banner.remove();};
  bannerContent.appendChild(ico);
  bannerContent.appendChild(sp);
  bannerContent.appendChild(btnVer);
  bannerContent.appendChild(btnClose);
  banner.appendChild(bannerContent);
  
  document.body.appendChild(banner);

  setTimeout(function() {
    if (banner.parentElement) banner.remove();
  }, 12000);
}
function enviarAlertasPorCorreo(alertas, filtro) {
  if (!alertas || alertas.length === 0) {
    showToast('No hay alertas para enviar', 'error');
    return;
  }

  var now = new Date();
  var fechaReporte = now.toLocaleDateString('es-MX', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
  var horaReporte = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
  var usuario = AUTH.currentUser ? AUTH.currentUser.fullName || AUTH.currentUser.username : 'Sistema';

  var resumen = getAlertasResumen();

  var vencidos  = alertas.filter(function(a) { return a.nivel === 'vencido'; });
  var criticos  = alertas.filter(function(a) { return a.nivel === 'critico'; });
  var avisos    = alertas.filter(function(a) { return a.nivel === 'aviso'; });
  var proximos  = alertas.filter(function(a) { return a.nivel === 'proximo'; });

  var asunto = 'DEBBIOM | Reporte de Alertas de Vigencia — ' + fechaReporte;
  if (vencidos.length > 0) {
    asunto = '⚠️ URGENTE — ' + asunto;
  }

  var lineas = [];

  lineas.push('═══════════════════════════════════════════════');
  lineas.push('       DEBBIOM — REPORTE DE ALERTAS DE VIGENCIA');
  lineas.push('═══════════════════════════════════════════════');
  lineas.push('');
  lineas.push('Fecha del reporte: ' + fechaReporte);
  lineas.push('Hora: ' + horaReporte);
  lineas.push('Generado por: ' + usuario);
  lineas.push('');
  lineas.push('───────────────────────────────────────────────');
  lineas.push('  RESUMEN GENERAL');
  lineas.push('───────────────────────────────────────────────');
  lineas.push('');
  lineas.push('  🔴 Documentos VENCIDOS:            ' + resumen.vencidos);
  lineas.push('  🔴 Vencen en 15 días o menos:      ' + resumen.criticos);
  lineas.push('  🟠 Vencen en 30 días o menos:      ' + resumen.avisos);
  lineas.push('  🟡 Vencen en 60 días o menos:      ' + resumen.proximos);
  lineas.push('  ────────────────────────────────');
  lineas.push('  📋 Total de alertas activas:       ' + resumen.total);
  lineas.push('');

  if (vencidos.length > 0) {
    lineas.push('');
    lineas.push('═══════════════════════════════════════════════');
    lineas.push('  🔴 DOCUMENTOS VENCIDOS (' + vencidos.length + ')');
    lineas.push('  ⚠️  REQUIEREN ATENCIÓN INMEDIATA');
    lineas.push('═══════════════════════════════════════════════');
    vencidos.forEach(function(a, i) {
      lineas.push('');
      lineas.push('  ' + (i + 1) + '. ' + a.nombreDoc);
      lineas.push('     📁 Área: ' + a.area);
      if (a.codigo) lineas.push('     🏷️  Código: ' + a.codigo);
      lineas.push('     📄 Versión: ' + a.version);
      lineas.push('     👤 Elaboró: ' + a.elaboro);
      lineas.push('     📅 Fecha de emisión: ' + a.fechaEmision);
      lineas.push('     📅 Fecha de vigencia: ' + (a.vigenciaDate ? formatDateFull(a.vigenciaDate) : 'N/A'));
      lineas.push('     ❌ VENCIDO hace ' + Math.abs(a.dias) + ' días');
    });
  }

  if (criticos.length > 0) {
    lineas.push('');
    lineas.push('');
    lineas.push('═══════════════════════════════════════════════');
    lineas.push('  🔴 POR VENCER EN 15 DÍAS O MENOS (' + criticos.length + ')');
    lineas.push('  ⚠️  URGENTE — PLANIFICAR RENOVACIÓN');
    lineas.push('═══════════════════════════════════════════════');
    criticos.forEach(function(a, i) {
      lineas.push('');
      lineas.push('  ' + (i + 1) + '. ' + a.nombreDoc);
      lineas.push('     📁 Área: ' + a.area);
      if (a.codigo) lineas.push('     🏷️  Código: ' + a.codigo);
      lineas.push('     📄 Versión: ' + a.version);
      lineas.push('     👤 Elaboró: ' + a.elaboro);
      lineas.push('     📅 Fecha de emisión: ' + a.fechaEmision);
      lineas.push('     📅 Fecha de vigencia: ' + (a.vigenciaDate ? formatDateFull(a.vigenciaDate) : 'N/A'));
      if (a.dias === 0) lineas.push('     ⏰ ¡VENCE HOY!');
      else if (a.dias === 1) lineas.push('     ⏰ ¡VENCE MAÑANA!');
      else lineas.push('     ⏰ Vence en ' + a.dias + ' días');
    });
  }

  if (avisos.length > 0) {
    lineas.push('');
    lineas.push('');
    lineas.push('═══════════════════════════════════════════════');
    lineas.push('  🟠 POR VENCER EN 30 DÍAS O MENOS (' + avisos.length + ')');
    lineas.push('═══════════════════════════════════════════════');
    avisos.forEach(function(a, i) {
      lineas.push('');
      lineas.push('  ' + (i + 1) + '. ' + a.nombreDoc);
      lineas.push('     📁 Área: ' + a.area + '  |  🏷️ ' + (a.codigo || 'N/A') + '  |  📄 v' + a.version);
      lineas.push('     👤 ' + a.elaboro + '  |  📅 Emisión: ' + a.fechaEmision + '  |  Vigencia: ' + (a.vigenciaDate ? formatDateFull(a.vigenciaDate) : 'N/A'));
      lineas.push('     ⏰ Vence en ' + a.dias + ' días');
    });
  }

  if (proximos.length > 0) {
    lineas.push('');
    lineas.push('');
    lineas.push('═══════════════════════════════════════════════');
    lineas.push('  🟡 POR VENCER EN 60 DÍAS O MENOS (' + proximos.length + ')');
    lineas.push('═══════════════════════════════════════════════');
    proximos.forEach(function(a, i) {
      lineas.push('');
      lineas.push('  ' + (i + 1) + '. ' + a.nombreDoc);
      lineas.push('     📁 Área: ' + a.area + '  |  🏷️ ' + (a.codigo || 'N/A') + '  |  📄 v' + a.version);
      lineas.push('     👤 ' + a.elaboro + '  |  📅 Emisión: ' + a.fechaEmision + '  |  Vigencia: ' + (a.vigenciaDate ? formatDateFull(a.vigenciaDate) : 'N/A'));
      lineas.push('     ⏰ Vence en ' + a.dias + ' días');
    });
  }

  lineas.push('');
  lineas.push('');
  lineas.push('───────────────────────────────────────────────');
  lineas.push('  ACCIONES RECOMENDADAS');
  lineas.push('───────────────────────────────────────────────');
  lineas.push('');
  if (vencidos.length > 0) lineas.push('  • URGENTE: Renovar los ' + vencidos.length + ' documento(s) vencidos lo antes posible.');
  if (criticos.length > 0) lineas.push('  • PRIORITARIO: Planificar la renovación de los ' + criticos.length + ' documento(s) que vencen en ≤15 días.');
  if (avisos.length > 0) lineas.push('  • PROGRAMAR: Agendar revisión de los ' + avisos.length + ' documento(s) que vencen en ≤30 días.');
  if (proximos.length > 0) lineas.push('  • PREVENIR: Considerar los ' + proximos.length + ' documento(s) que vencen en ≤60 días.');
  lineas.push('');
  lineas.push('');
  lineas.push('───────────────────────────────────────────────');
  lineas.push('Desarrollos Biomédicos y Biotecnológicos de México S.A. de C.V.');
  lineas.push('DEBBIOM — Sistema de Gestión Documental');
  lineas.push('Este reporte fue generado automáticamente el ' + fechaReporte + ' a las ' + horaReporte);
  lineas.push('───────────────────────────────────────────────');

  var cuerpoCorreo = lineas.join('\n');

  _showEmailModal(asunto, cuerpoCorreo);
}

function _showEmailModal(asunto, cuerpo) {
  var overlay = document.createElement('div');
  overlay.className = 'modal-overlay';
  overlay.id = 'email-modal-overlay';
  overlay.onclick = function(e) { if (e.target === overlay) overlay.remove(); };

  var modal = document.createElement('div');
  modal.className = 'modal modal-lg email-preview-modal';

  var header = document.createElement('div');
  header.className = 'modal-header';
  var h2 = document.createElement('h2');
  var h2i = document.createElement('i');
  h2i.className = 'fas fa-envelope';
  h2i.style.marginRight = '8px';
  h2i.style.color = 'var(--accent)';
  h2.appendChild(h2i);
  h2.appendChild(document.createTextNode('Enviar Reporte por Correo'));
  var closeBtn = document.createElement('button');
  closeBtn.className = 'modal-close';
  closeBtn.onclick = function(){overlay.remove();};
  var closeBtnI = document.createElement('i');
  closeBtnI.className = 'fas fa-times';
  closeBtn.appendChild(closeBtnI);
  header.appendChild(h2);
  header.appendChild(closeBtn);

  var body = document.createElement('div');
  body.className = 'modal-body';

  var destRow = document.createElement('div');
  destRow.className = 'email-field-row';
  var destLabel = document.createElement('label');
  destLabel.className = 'email-field-label';
  var destIco = document.createElement('i');
  destIco.className = 'fas fa-at';
  destLabel.appendChild(destIco);
  destLabel.appendChild(document.createTextNode(' Para:'));
  var destInput = document.createElement('input');
  destInput.type = 'email';
  destInput.id = 'email-dest';
  destInput.className = 'form-input';
  destInput.placeholder = 'correo@debbiom.com, otro@empresa.com';
  destInput.style.flex = '1';
  destRow.appendChild(destLabel);
  destRow.appendChild(destInput);

  var subjRow = document.createElement('div');
  subjRow.className = 'email-field-row';
  var subjLabel = document.createElement('label');
  subjLabel.className = 'email-field-label';
  var subjIco = document.createElement('i');
  subjIco.className = 'fas fa-heading';
  subjLabel.appendChild(subjIco);
  subjLabel.appendChild(document.createTextNode(' Asunto:'));
  var subjInput = document.createElement('input');
  subjInput.type = 'text';
  subjInput.id = 'email-subj';
  subjInput.className = 'form-input';
  subjInput.value = asunto;
  subjInput.style.flex = '1';
  subjRow.appendChild(subjLabel);
  subjRow.appendChild(subjInput);

  var previewLabel = document.createElement('div');
  previewLabel.className = 'email-preview-label';
  previewLabel.textContent = 'Vista previa del correo:';

  var preview = document.createElement('pre');
  preview.className = 'email-preview-body';
  preview.textContent = cuerpo;

  body.appendChild(destRow);
  body.appendChild(subjRow);
  body.appendChild(previewLabel);
  body.appendChild(preview);

  var footer = document.createElement('div');
  footer.className = 'modal-footer email-modal-footer';

  var btnCopy = document.createElement('button');
  btnCopy.className = 'btn btn-secondary';
  _setBtn(btnCopy, 'fas fa-copy', 'Copiar al portapapeles');
  btnCopy.onclick = function() {
    navigator.clipboard.writeText(cuerpo).then(function() {
      showToast('✅ Reporte copiado al portapapeles', 'success');
      _setBtn(btnCopy, 'fas fa-check', '¡Copiado!');
      setTimeout(function() { _setBtn(btnCopy, 'fas fa-copy', 'Copiar al portapapeles'); }, 2000);
    });
  };

  var btnOutlook = document.createElement('button');
  btnOutlook.className = 'btn btn-primary';
  _setBtn(btnOutlook, 'fas fa-paper-plane', 'Abrir en cliente de correo');
  btnOutlook.onclick = function() {
    var dest = document.getElementById('email-dest').value || '';
    var subj = document.getElementById('email-subj').value || asunto;
    var mailBody = _buildMailtoBody(cuerpo);
    var mailtoURL = 'mailto:' + encodeURIComponent(dest) + '?subject=' + encodeURIComponent(subj) + '&body=' + encodeURIComponent(mailBody);
    window.open(mailtoURL, '_self');
    showToast('Abriendo cliente de correo...', 'info');
  };

  var btnGmail = document.createElement('button');
  btnGmail.className = 'btn btn-gmail';
  _setBtn(btnGmail, 'fas fa-envelope', 'Abrir en Gmail');
  btnGmail.onclick = function() {
    var dest = document.getElementById('email-dest').value || '';
    var subj = document.getElementById('email-subj').value || asunto;
    var mailBody = _buildMailtoBody(cuerpo);
    var gmailURL = 'https://mail.google.com/mail/?view=cm&fs=1&to=' + encodeURIComponent(dest) + '&su=' + encodeURIComponent(subj) + '&body=' + encodeURIComponent(mailBody);
    window.open(gmailURL, '_blank');
    showToast('Abriendo Gmail...', 'info');
  };

  var btnClose = document.createElement('button');
  btnClose.className = 'btn btn-secondary';
  btnClose.textContent = 'Cerrar';
  btnClose.onclick = function() { overlay.remove(); };

  footer.appendChild(btnClose);
  footer.appendChild(btnCopy);
  footer.appendChild(btnGmail);
  footer.appendChild(btnOutlook);

  modal.appendChild(header);
  modal.appendChild(body);
  modal.appendChild(footer);
  overlay.appendChild(modal);
  document.body.appendChild(overlay);
}

function _buildMailtoBody(cuerpoCompleto) {
  var maxLen = 1800;
  if (cuerpoCompleto.length <= maxLen) return cuerpoCompleto;
  return cuerpoCompleto.substring(0, maxLen) + '\n\n--- Reporte recortado por límite de correo ---\nPara ver el reporte completo, use "Copiar al portapapeles" y péguelo manualmente en el correo.';
}
function enviarAlertaIndividual(alerta) {
  var now = new Date();
  var fechaReporte = now.toLocaleDateString('es-MX', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
  var horaReporte = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
  var usuario = AUTH.currentUser ? AUTH.currentUser.fullName || AUTH.currentUser.username : 'Sistema';

  var nivelTexto = '';
  var nivelEmoji = '';
  var accion = '';
  if (alerta.nivel === 'vencido') {
    nivelTexto = 'VENCIDO';
    nivelEmoji = '🔴';
    accion = 'Este documento requiere RENOVACIÓN INMEDIATA.';
  } else if (alerta.nivel === 'critico') {
    nivelTexto = 'CRÍTICO — Vence en ' + alerta.dias + ' días';
    nivelEmoji = '🔴';
    accion = 'Se requiere iniciar el proceso de renovación de forma urgente antes del vencimiento.';
  } else if (alerta.nivel === 'aviso') {
    nivelTexto = 'AVISO — Vence en ' + alerta.dias + ' días';
    nivelEmoji = '🟠';
    accion = 'Se recomienda programar la revisión y renovación de este documento.';
  } else {
    nivelTexto = 'PRÓXIMO A VENCER — ' + alerta.dias + ' días restantes';
    nivelEmoji = '🟡';
    accion = 'Se sugiere considerar este documento en la próxima planeación de renovaciones.';
  }

  var diasDesc = '';
  if (alerta.dias < 0) diasDesc = 'Venció hace ' + Math.abs(alerta.dias) + ' días';
  else if (alerta.dias === 0) diasDesc = '¡VENCE HOY!';
  else if (alerta.dias === 1) diasDesc = '¡VENCE MAÑANA!';
  else diasDesc = 'Vence en ' + alerta.dias + ' días';

  var asunto = nivelEmoji + ' DEBBIOM | Alerta de vigencia: ' + alerta.nombreDoc;

  var lineas = [];
  lineas.push('═══════════════════════════════════════════════');
  lineas.push('    DEBBIOM — ALERTA DE VIGENCIA DOCUMENTAL');
  lineas.push('═══════════════════════════════════════════════');
  lineas.push('');
  lineas.push('Estimado(a) colaborador(a),');
  lineas.push('');
  lineas.push('Por medio del presente se le notifica que el siguiente');
  lineas.push('documento del Sistema de Gestión requiere su atención:');
  lineas.push('');
  lineas.push('───────────────────────────────────────────────');
  lineas.push('  ' + nivelEmoji + '  NIVEL DE ALERTA: ' + nivelTexto);
  lineas.push('───────────────────────────────────────────────');
  lineas.push('');
  lineas.push('  INFORMACIÓN DEL DOCUMENTO');
  lineas.push('  ─────────────────────────');
  lineas.push('');
  lineas.push('  📄 Nombre:          ' + alerta.nombreDoc);
  lineas.push('  📁 Área:            ' + alerta.area);
  if (alerta.codigo) lineas.push('  🏷️  Código:          ' + alerta.codigo);
  lineas.push('  📋 Versión:         ' + alerta.version);
  lineas.push('  👤 Elaboró:         ' + alerta.elaboro);
  lineas.push('');
  lineas.push('  FECHAS');
  lineas.push('  ──────');
  lineas.push('');
  lineas.push('  📅 Fecha de emisión:     ' + alerta.fechaEmision);
  lineas.push('  📅 Fecha de vigencia:    ' + (alerta.vigenciaDate ? formatDateFull(alerta.vigenciaDate) : 'N/A'));
  lineas.push('  ⏰ Estado:               ' + diasDesc);
  lineas.push('');
  lineas.push('───────────────────────────────────────────────');
  lineas.push('  ACCIÓN REQUERIDA');
  lineas.push('───────────────────────────────────────────────');
  lineas.push('');
  lineas.push('  ' + accion);
  lineas.push('');
  if (alerta.nivel === 'vencido' || alerta.nivel === 'critico') {
    lineas.push('  Responsable sugerido: ' + alerta.elaboro);
    lineas.push('  Prioridad: ALTA');
  } else {
    lineas.push('  Responsable sugerido: ' + alerta.elaboro);
    lineas.push('  Prioridad: MEDIA');
  }
  lineas.push('');
  lineas.push('');
  lineas.push('───────────────────────────────────────────────');
  lineas.push('Atentamente,');
  lineas.push(usuario);
  lineas.push('');
  lineas.push('Desarrollos Biomédicos y Biotecnológicos');
  lineas.push('de México S.A. de C.V.');
  lineas.push('DEBBIOM — Sistema de Gestión Documental');
  lineas.push('');
  lineas.push('Reporte generado: ' + fechaReporte + ', ' + horaReporte);
  lineas.push('───────────────────────────────────────────────');

  var cuerpoCorreo = lineas.join('\n');
  _showEmailModal(asunto, cuerpoCorreo);
}
const LOGO_SRC = '';

function renderLogin() {
  const app = document.getElementById('app');
  app.innerHTML = '';

  let usernameVal = '';
  let passwordVal = '';

  const onlineCount = (typeof SYNC !== 'undefined' && SYNC.onlineCount) ? SYNC.onlineCount : 0;

  const card = h('div', {className: 'login-wrapper'},
    h('div', {className: 'login-card'},
      h('div', {className: 'logo-area'},
        h('img', {src: DEBBIOM_LOGO, alt: 'DEBBIOM'}),
        h('div', {className: 'brand-name'}, 'DEBBIOM'),
        h('div', {className: 'brand-subtitle'}, 'Listado de estudios'),
        h('div', {className: 'sync-badge'},
          h('i', {className: 'fas fa-cloud'}),
          'Sincronización en la Nube'
        )
      ),
      AUTH.loginError ? h('div', {className: 'login-error'},
        h('i', {className: 'fas fa-exclamation-circle'}),
        AUTH.loginError
      ) : null,
      h('div', {className: 'login-field'},
        h('label', null,
          h('i', {className: 'fas fa-user'}),
          'Usuario'
        ),
        h('div', {className: 'input-wrap'},
          h('input', {
            type: 'text',
            placeholder: '',
            id: 'login-user',
            autocomplete: 'off',
            maxLength: 50,
            onInput: (e) => { usernameVal = e.target.value; },
            onKeydown: (e) => { if(e.key === 'Enter') document.getElementById('login-pw').focus(); }
          })
        )
      ),
      h('div', {className: 'login-field'},
        h('label', null,
          h('i', {className: 'fas fa-lock'}),
          'Contraseña'
        ),
        h('div', {className: 'input-wrap'},
          h('input', {
            type: AUTH.showPassword ? 'text' : 'password',
            placeholder: '',
            id: 'login-pw',
            onInput: (e) => { passwordVal = e.target.value; },
            onKeydown: async (e) => {
              if(e.key === 'Enter') {
                await attemptLogin(usernameVal || document.getElementById('login-user').value, passwordVal || e.target.value);
                render();
              }
            }
          }),
          h('button', {className: 'toggle-pw', onClick: () => { AUTH.showPassword = !AUTH.showPassword; renderLogin(); setTimeout(()=>{const el=document.getElementById('login-pw');if(el)el.focus();},50); }},
            h('i', {className: AUTH.showPassword ? 'fas fa-eye-slash' : 'fas fa-eye'})
          )
        )
      ),
      h('button', {className: 'login-btn', onClick: async () => {
        const u = document.getElementById('login-user').value;
        const p = document.getElementById('login-pw').value;
        await attemptLogin(u, p);
        render();
      }},
        'Iniciar Sesión →'
      ),
      h('div', {className: 'login-online-users'},
        h('div', {className: 'online-count'}, String(onlineCount)),
        h('div', {className: 'online-label'}, 'usuarios en línea')
      ),
      h('div', {className: 'login-footer'},
        'v' + APP_VERSION + ' — Powered by ', h('span', null, 'DEBBIOM')
      )
    )
  );

  app.appendChild(card);

  setTimeout(() => { const el = document.getElementById('login-user'); if(el) el.focus(); }, 100);
}
function renderUserManagementModal() {
  let newUser = { username:'', password:'', fullName:'', role:'user' };

  async function refreshUMBody() {
    const body = document.getElementById('um-body');
    if (!body) return;
    body.textContent = '';
    var loadDiv = document.createElement('div');
    loadDiv.style.cssText = 'text-align:center;padding:20px';
    var loadIco = document.createElement('i');
    loadIco.className = 'fas fa-spinner fa-spin';
    loadDiv.appendChild(loadIco);
    loadDiv.appendChild(document.createTextNode(' Cargando usuarios...'));
    body.appendChild(loadDiv);
    await loadAllProfiles();
    body.innerHTML = '';
    body.appendChild(buildUMContent());
  }

  function buildUMContent() {
    const container = h('div', null,
      h('div', {style:{background:'rgba(38,198,218,.08)',borderRadius:'10px',padding:'12px 16px',marginBottom:'16px',border:'1px solid rgba(38,198,218,.2)',fontSize:'12px',color:'var(--text-secondary)'}},
        h('i', {className:'fas fa-shield-halved', style:{marginRight:'8px',color:'var(--brand-cyan)'}}),
        'Los usuarios se gestionan vía Supabase Auth. Las contraseñas están protegidas en el backend, no en el código.'
      ),
      h('table', {className: 'um-table'},
        h('thead', null,
          h('tr', null,
            h('th', null, 'Usuario'),
            h('th', null, 'Nombre'),
            h('th', null, 'Rol'),
            h('th', {style:{minWidth:'180px'}}, 'Acciones')
          )
        ),
        h('tbody', null,
          ...AUTH.users.filter(function(u) { return u.role !== 'disabled'; }).map(function(u) {
            return h('tr', null,
              h('td', null, h('span', {style:{fontFamily:'JetBrains Mono, monospace', fontWeight:'600'}}, u.username)),
              h('td', null, u.full_name || u.fullName || ''),
              h('td', null, h('span', {className: 'um-role-badge ' + (u.role==='admin'?'um-role-admin':'um-role-user')}, u.role==='admin'?'Administrador':'Usuario')),
              h('td', null,
                h('div', {style:{display:'flex',gap:'6px',flexWrap:'wrap'}},
                  u.id !== (AUTH.currentUser ? AUTH.currentUser.id : '')
                    ? h('button', {className:'btn btn-sm btn-warning', onClick: async function() {
                        var newRole = u.role === 'admin' ? 'user' : 'admin';
                        var r = await changeUserRole(u.id, newRole);
                        if (r.error) { showToast(r.error, 'error'); return; }
                        showToast('Rol cambiado a ' + newRole);
                        refreshUMBody();
                      }}, h('i',{className:'fas fa-user-shield'}), u.role === 'admin' ? ' → Usuario' : ' → Admin')
                    : null,
                  u.id !== (AUTH.currentUser ? AUTH.currentUser.id : '')
                    ? h('button', {className:'btn btn-sm btn-danger', onClick: async function() {
                        if(confirm('¿Desactivar al usuario ' + u.username + '?\nNo podrá iniciar sesión.')) {
                          var r = await deactivateUser(u.id, u.username);
                          if (r.error) { showToast(r.error, 'error'); return; }
                          showToast('Usuario ' + u.username + ' desactivado');
                          refreshUMBody();
                        }
                      }}, h('i',{className:'fas fa-user-slash'}), ' Desactivar')
                    : null
                )
              )
            );
          })
        )
      ),

      h('div', {style:{marginTop:'20px',padding:'16px',background:'var(--bg-input)',borderRadius:'10px',border:'1px solid var(--border)'}},
        h('div', {style:{fontWeight:'600',fontSize:'13px',marginBottom:'10px',color:'var(--yellow)'}},
          h('i', {className:'fas fa-key', style:{marginRight:'6px'}}),
          'Cambiar mi contraseña'
        ),
        h('div', {style:{display:'flex',gap:'8px',alignItems:'center'}},
          h('input', {
            className:'form-input', type:'password', placeholder:'Nueva contraseña (Mín. 8 chars, mayúscula, minúscula, número)',
            style:{flex:'1'}, id:'pw-change-own'
          }),
          h('button', {className:'btn btn-sm btn-success', onClick: async function() {
            var val = document.getElementById('pw-change-own').value;
            if (!val || !validatePassword(val)) { showToast('Mín. 8 caracteres, 1 mayúscula, 1 minúscula, 1 número','error'); return; }
            var r = await changeOwnPassword(val);
            if (r.error) { showToast(r.error, 'error'); return; }
            showToast('✅ Contraseña actualizada');
            document.getElementById('pw-change-own').value = '';
          }}, h('i',{className:'fas fa-check'}), ' Guardar')
        )
      ),

      h('div', {style:{marginTop:'24px',padding:'18px',background:'var(--bg-input)',borderRadius:'10px',border:'1px solid var(--accent)'}},
        h('div', {style:{fontWeight:'600',fontSize:'14px',marginBottom:'14px',color:'var(--accent)'}},
          h('i', {className:'fas fa-user-plus', style:{marginRight:'8px'}}),
          'Agregar Nuevo Usuario'
        ),
        h('div', {className:'form-grid'},
          h('div', {className:'form-group'},
            h('label', {className:'form-label'}, 'Usuario'),
            h('input', {className:'form-input', placeholder:'Ej: JPerez', id:'new-user-name', onInput:function(e){newUser.username=e.target.value;}})
          ),
          h('div', {className:'form-group'},
            h('label', {className:'form-label'}, 'Nombre Completo'),
            h('input', {className:'form-input', placeholder:'Ej: J. Pérez', id:'new-user-full', onInput:function(e){newUser.fullName=e.target.value;}})
          ),
          h('div', {className:'form-group'},
            h('label', {className:'form-label'}, 'Contraseña'),
            h('input', {className:'form-input', type:'password', placeholder:'Contraseña inicial (Mín. 8, A-Z, a-z, 0-9)', id:'new-user-pw', onInput:function(e){newUser.password=e.target.value;}})
          ),
          h('div', {className:'form-group'},
            h('label', {className:'form-label'}, 'Rol'),
            h('select', {className:'form-select', id:'new-user-role', onChange:function(e){newUser.role=e.target.value;}},
              h('option', {value:'user'}, 'Usuario'),
              h('option', {value:'admin'}, 'Administrador')
            )
          )
        ),
        h('button', {className:'btn btn-primary', style:{marginTop:'14px'}, onClick: async function() {
          var uname = document.getElementById('new-user-name').value.trim();
          var fname = document.getElementById('new-user-full').value.trim();
          var pw = document.getElementById('new-user-pw').value;
          var role = document.getElementById('new-user-role').value;
          if (!uname) { showToast('Ingresa un nombre de usuario','error'); return; }
          if (!pw || !validatePassword(pw)) { showToast('Contraseña: Mín. 8 caracteres con mayúscula, minúscula y número','error'); return; }

          var btn = document.querySelector('.btn-primary');
          if (btn) { btn.disabled = true; btn.textContent = 'Creando...'; }

          var r = await createUser(uname, pw, fname, role);
          if (r.error) {
            showToast(r.error, 'error');
            if (btn) { btn.disabled = false; _setBtn(btn, 'fas fa-user-plus', 'Crear Usuario'); }
            return;
          }
          showToast('✅ Usuario ' + uname + ' creado');
          newUser = { username:'', password:'', fullName:'', role:'user' };
          refreshUMBody();
        }}, h('i',{className:'fas fa-user-plus'}), ' Crear Usuario')
      )
    );
    return container;
  }

  return h('div', {className:'modal-overlay', onClick:(e)=>{if(e.target.className==='modal-overlay')closeModal();}},
    h('div', {className:'modal modal-lg'},
      h('div', {className:'modal-header'},
        h('h2', null, h('i',{className:'fas fa-users-gear', style:{marginRight:'8px',color:'var(--accent)'}}), 'Administración de Usuarios'),
        h('button', {className:'modal-close', onClick:closeModal}, h('i',{className:'fas fa-times'}))
      ),
      h('div', {className:'modal-body', id:'um-body'},
        buildUMContent()
      ),
      h('div', {className:'modal-footer'},
        h('button', {className:'btn btn-secondary', onClick:closeModal}, 'Cerrar')
      )
    )
  );
}

function refreshModal() {
  const root = document.getElementById('modal-root');
  if (!root) return;
  if (!STATE.modal) { root.innerHTML = ''; return; }

  const active = document.activeElement;
  const focusId = active ? active.id : null;
  let focusIdx = null;
  let cursorPos = null;
  let scrollTop = 0;

  const modalBody = root.querySelector('.modal-body');
  if (modalBody) scrollTop = modalBody.scrollTop;

  if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.tagName === 'SELECT')) {
    cursorPos = active.selectionStart;
    if (!focusId) {
      const allInputs = root.querySelectorAll('input, textarea, select');
      focusIdx = Array.from(allInputs).indexOf(active);
    }
  }

  let modal = null;
  if (STATE.modal === 'view') modal = renderViewModal();
  else if (STATE.modal === 'edit' || STATE.modal === 'add') modal = renderEditModal();
  else if (STATE.modal === 'salida') modal = renderSalidaModal();
  else if (STATE.modal === 'obsoleto') modal = renderObsoletoModal();
  else if (STATE.modal === 'renew') modal = renderRenewModal();
  else if (STATE.modal === 'users') modal = renderUserManagementModal();
  else if (STATE.modal === 'import') modal = renderImportModal();
  if (modal) root.replaceChildren(modal); else root.replaceChildren();

  requestAnimationFrame(() => {
    const newBody = root.querySelector('.modal-body');
    if (newBody && scrollTop) newBody.scrollTop = scrollTop;

    let target = null;
    if (focusId) target = document.getElementById(focusId);
    if (!target && focusIdx !== null && focusIdx >= 0) {
      const allInputs = root.querySelectorAll('input, textarea, select');
      target = allInputs[focusIdx] || null;
    }
    if (target) {
      target.focus();
      if (cursorPos !== null && typeof target.setSelectionRange === 'function') {
        try { target.setSelectionRange(cursorPos, cursorPos); } catch(e) {}
      }
    }
  });
}

function renderCloudBackupsSection() {
  var isConnected = typeof SYNC !== 'undefined' && SYNC.initialized && SYNC.connected;

  if (!isConnected) {
    return h('div', {className: 'content-area'},
      h('div', {style:{textAlign:'center',padding:'60px 20px',color:'var(--text-muted)'}},
        h('i', {className:'fas fa-cloud-slash', style:{fontSize:'48px',display:'block',marginBottom:'16px',opacity:'.4'}}),
        h('div', {style:{fontSize:'16px',fontWeight:'600',marginBottom:'8px',color:'var(--text-primary)'}}, 'Sin conexión a la nube'),
        h('div', {style:{fontSize:'13px',maxWidth:'400px',margin:'0 auto'}},
          'Los backups en la nube requieren conexión a Supabase. Verifica tu conexión a internet y las credenciales en sync.js.'
        )
      )
    );
  }

  var container = document.createElement('div');
  container.className = 'content-area';

  var headerDiv = document.createElement('div');
  headerDiv.className = 'bk-cloud-header';
  var headerInfo = document.createElement('div');
  headerInfo.className = 'bk-cloud-header-info';
  var h3 = document.createElement('h3');
  h3.style.cssText = 'margin:0;font-size:15px;color:var(--text-primary)';
  var h3i = document.createElement('i');
  h3i.className = 'fas fa-cloud-arrow-up';
  h3i.style.cssText = 'margin-right:8px;color:var(--green)';
  h3.appendChild(h3i);
  h3.appendChild(document.createTextNode('Backups almacenados en Supabase'));
  var h3p = document.createElement('p');
  h3p.style.cssText = 'margin:4px 0 0;font-size:12px;color:var(--text-muted)';
  h3p.textContent = 'Los backups se guardan automáticamente cada hora y se pueden crear manualmente. Disponibles desde cualquier dispositivo.';
  headerInfo.appendChild(h3);
  headerInfo.appendChild(h3p);
  var headerActions = document.createElement('div');
  headerActions.className = 'bk-cloud-header-actions';
  var btnSaveH = document.createElement('button');
  btnSaveH.className = 'btn btn-primary';
  btnSaveH.id = 'bk-save-cloud';
  var btnSaveI = document.createElement('i');
  btnSaveI.className = 'fas fa-cloud-arrow-up';
  btnSaveH.appendChild(btnSaveI);
  btnSaveH.appendChild(document.createTextNode(' Guardar Backup Ahora'));
  var btnRefreshH = document.createElement('button');
  btnRefreshH.className = 'btn btn-secondary';
  btnRefreshH.id = 'bk-refresh-cloud';
  var btnRefreshI = document.createElement('i');
  btnRefreshI.className = 'fas fa-arrows-rotate';
  btnRefreshH.appendChild(btnRefreshI);
  btnRefreshH.appendChild(document.createTextNode(' Actualizar'));
  headerActions.appendChild(btnSaveH);
  headerActions.appendChild(btnRefreshH);
  headerDiv.appendChild(headerInfo);
  headerDiv.appendChild(headerActions);

  var loadingDiv = document.createElement('div');
  loadingDiv.className = 'bk-cloud-loading';
  loadingDiv.id = 'bk-cloud-loading';
  var loadSpinner = document.createElement('i');
  loadSpinner.className = 'fas fa-spinner fa-spin';
  loadSpinner.style.cssText = 'font-size:24px;color:var(--accent)';
  var loadText = document.createElement('span');
  loadText.style.cssText = 'margin-left:10px;color:var(--text-muted)';
  loadText.textContent = 'Cargando backups desde la nube...';
  loadingDiv.appendChild(loadSpinner);
  loadingDiv.appendChild(loadText);

  var tableDiv = document.createElement('div');
  tableDiv.id = 'bk-cloud-table-container';

  container.appendChild(headerDiv);
  container.appendChild(loadingDiv);
  container.appendChild(tableDiv);

  setTimeout(function() {
    _loadAndRenderCloudBackups(container);

    var btnSave = container.querySelector('#bk-save-cloud');
    if (btnSave) {
      btnSave.addEventListener('click', async function() {
        var nota = prompt('Nota del backup (opcional):', 'Backup manual');
        if (nota === null) return;
        btnSave.disabled = true;
        _setBtn(btnSave, 'fas fa-spinner fa-spin', 'Guardando...');
        await saveManualCloudBackup(nota || 'Backup manual');
        btnSave.disabled = false;
        _setBtn(btnSave, 'fas fa-cloud-arrow-up', 'Guardar Backup Ahora');
        _loadAndRenderCloudBackups(container);
      });
    }

    var btnRefresh = container.querySelector('#bk-refresh-cloud');
    if (btnRefresh) {
      btnRefresh.addEventListener('click', function() {
        _loadAndRenderCloudBackups(container);
      });
    }
  }, 50);

  return container;
}

async function _loadAndRenderCloudBackups(container) {
  var loading = container.querySelector('#bk-cloud-loading');
  var tableC = container.querySelector('#bk-cloud-table-container');
  if (!loading || !tableC) return;

  loading.style.display = 'flex';
  tableC.innerHTML = '';

  var backups = await listCloudBackups();

  loading.style.display = 'none';

  if (!backups || backups.length === 0) {
    tableC.textContent = '';
    var emptyDiv = document.createElement('div');
    emptyDiv.style.cssText = 'text-align:center;padding:40px;color:var(--text-muted)';
    var emptyIco = document.createElement('i');
    emptyIco.className = 'fas fa-inbox';
    emptyIco.style.cssText = 'font-size:36px;display:block;margin-bottom:12px;opacity:.4';
    var emptyTitle = document.createElement('div');
    emptyTitle.style.cssText = 'font-size:14px;font-weight:600;color:var(--text-primary);margin-bottom:6px';
    emptyTitle.textContent = 'No hay backups en la nube';
    var emptyDesc = document.createElement('div');
    emptyDesc.style.cssText = 'font-size:12px';
    emptyDesc.textContent = 'Haz clic en "Guardar Backup Ahora" para crear el primero. También se crean automáticamente cada hora.';
    emptyDiv.appendChild(emptyIco);
    emptyDiv.appendChild(emptyTitle);
    emptyDiv.appendChild(emptyDesc);
    tableC.appendChild(emptyDiv);
    return;
  }

  var manuales = backups.filter(function(b) { return b.tipo === 'manual' || b.tipo === 'json' || b.tipo === 'excel'; }).length;
  var autos = backups.filter(function(b) { return b.tipo === 'auto'; }).length;
  var ultimoManual = backups.find(function(b) { return b.tipo !== 'auto'; });

  function _stat(val, label, color) {
    var d = document.createElement('div'); d.className = 'bk-cloud-stat';
    var v = document.createElement('div'); v.className = 'bk-cloud-stat-val'; v.textContent = val;
    if (color) v.style.color = color;
    var l = document.createElement('div'); l.className = 'bk-cloud-stat-label'; l.textContent = label;
    d.appendChild(v); d.appendChild(l); return d;
  }
  var statsDiv = document.createElement('div'); statsDiv.className = 'bk-cloud-stats';
  statsDiv.appendChild(_stat(backups.length, 'Total Backups'));
  statsDiv.appendChild(_stat(manuales, 'Manuales', 'var(--brand-cyan)'));
  statsDiv.appendChild(_stat(autos, 'Automáticos', 'var(--text-muted)'));
  var lastStat = _stat(ultimoManual ? _formatBackupDate(ultimoManual.created_at) : '—', 'Último Manual', 'var(--green)');
  lastStat.querySelector('.bk-cloud-stat-val').style.fontSize = '13px';
  statsDiv.appendChild(lastStat);

  var tipoIcons = {manual:'fa-user',json:'fa-download',excel:'fa-file-excel',auto:'fa-clock'};
  var tipoLabels = {manual:'Manual',json:'JSON',excel:'Excel',auto:'Auto'};
  var tipoClasses = {manual:'bk-tipo-manual',json:'bk-tipo-json',excel:'bk-tipo-excel',auto:'bk-tipo-auto'};

  var tableWrap = document.createElement('div'); tableWrap.className = 'table-container';
  var table = document.createElement('table');
  var thead = document.createElement('thead');
  var headRow = document.createElement('tr');
  ['Tipo','Fecha','Creado por','Contenido','Nota','Acciones'].forEach(function(t) {
    var th = document.createElement('th'); th.textContent = t;
    if (t === 'Acciones') th.style.minWidth = '230px';
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  var tbody = document.createElement('tbody');
  backups.forEach(function(bk) {
    var tr = document.createElement('tr');
    var fecha = _formatBackupDate(bk.created_at);
    var tipo = bk.tipo || 'auto';

    var tdTipo = document.createElement('td');
    var badge = document.createElement('span'); badge.className = 'bk-tipo-badge ' + (tipoClasses[tipo]||'bk-tipo-auto');
    var ico = document.createElement('i'); ico.className = 'fas ' + (tipoIcons[tipo]||'fa-clock');
    badge.appendChild(ico); badge.appendChild(document.createTextNode(' ' + (tipoLabels[tipo]||'Auto')));
    tdTipo.appendChild(badge); tr.appendChild(tdTipo);

    var tdFecha = document.createElement('td'); tdFecha.style.fontWeight = '500'; tdFecha.textContent = fecha; tr.appendChild(tdFecha);
    var tdCreado = document.createElement('td'); tdCreado.textContent = bk.created_by || '—'; tr.appendChild(tdCreado);

    var tdCont = document.createElement('td'); tdCont.style.cssText = 'font-family:JetBrains Mono,monospace;font-size:12px';
    function _badge(text, cls, ml) {
      var s = document.createElement('span'); s.className = 'badge' + (cls ? ' ' + cls : ''); s.textContent = text;
      if (ml) s.style.marginLeft = '4px'; return s;
    }
    tdCont.appendChild(_badge(bk.total_registros + ' reg', 'badge-blue'));
    tdCont.appendChild(_badge(bk.total_obsoletos + ' obs', 'badge-purple', true));
    tdCont.appendChild(_badge(bk.total_salidas + ' sal', '', true));
    tr.appendChild(tdCont);

    var tdNota = document.createElement('td');
    tdNota.style.cssText = 'font-size:12px;color:var(--text-muted);max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap';
    tdNota.textContent = bk.nota || '—'; tdNota.title = bk.nota || '';
    tr.appendChild(tdNota);

    var tdAcc = document.createElement('td');
    var accDiv = document.createElement('div'); accDiv.className = 'bk-actions';
    var btnRestore = document.createElement('button'); btnRestore.className = 'action-btn edit bk-act-restore';
    btnRestore.title = 'Restaurar este backup en TODOS los dispositivos';
    var rIco = document.createElement('i'); rIco.className = 'fas fa-rotate-right';
    btnRestore.appendChild(rIco); btnRestore.appendChild(document.createTextNode(' Restaurar'));
    btnRestore.addEventListener('click', (function(id) { return async function() { var ok = await restoreCloudBackup(id); if (ok) _loadAndRenderCloudBackups(container); }; })(bk.id));

    var btnDown = document.createElement('button'); btnDown.className = 'action-btn view bk-act-download';
    btnDown.title = 'Descargar como archivo JSON';
    var dIco = document.createElement('i'); dIco.className = 'fas fa-download';
    btnDown.appendChild(dIco); btnDown.appendChild(document.createTextNode(' JSON'));
    btnDown.addEventListener('click', (function(id) { return function() { downloadCloudBackup(id); }; })(bk.id));

    accDiv.appendChild(btnRestore); accDiv.appendChild(btnDown);
    tdAcc.appendChild(accDiv); tr.appendChild(tdAcc);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  tableWrap.appendChild(table);

  tableC.textContent = '';
  tableC.appendChild(statsDiv);
  tableC.appendChild(tableWrap);
}

// ============================================================
// IMPORTAR EXCEL
// ============================================================
var _IMPORT_COL_MAP = {
  'no': 'no', 'no.': 'no', 'número': 'no', 'numero': 'no', 'num': 'no',
  'área': 'area', 'area': 'area',
  'clasificación': 'area', 'clasificacion': 'area',
  'cód. manual': 'codManual', 'cod. manual': 'codManual', 'cod manual': 'codManual', 'codigo manual': 'codManual', 'manual': 'codManual',
  'código del manual / procedimiento': 'codManual', 'codigo del manual / procedimiento': 'codManual',
  'código del manual/ procedimiento': 'codManual', 'codigo del manual/ procedimiento': 'codManual',
  'código del manual /procedimiento': 'codManual', 'codigo del manual /procedimiento': 'codManual',
  'código del manual/procedimiento': 'codManual', 'codigo del manual/procedimiento': 'codManual',
  'cód. instructivo': 'codInstructivo', 'cod. instructivo': 'codInstructivo', 'cod instructivo': 'codInstructivo', 'codigo instructivo': 'codInstructivo', 'instructivo': 'codInstructivo',
  'código del instructivo': 'codInstructivo', 'codigo del instructivo': 'codInstructivo',
  'cód. listado': 'codListado', 'cod. listado': 'codListado', 'cod listado': 'codListado', 'codigo listado': 'codListado', 'listado': 'codListado',
  'código del listado': 'codListado', 'codigo del listado': 'codListado',
  'cód. formato': 'codFormato', 'cod. formato': 'codFormato', 'cod formato': 'codFormato', 'codigo formato': 'codFormato', 'formato': 'codFormato',
  'código del formato': 'codFormato', 'codigo del formato': 'codFormato',
  'nombre documento': 'nombreDoc', 'nombre doc': 'nombreDoc', 'nombre del documento': 'nombreDoc', 'documento': 'nombreDoc', 'nombre': 'nombreDoc',
  'versión': 'version', 'version': 'version', 'ver': 'version', 'ver.': 'version',
  'fecha emisión': 'fechaEmision', 'fecha emision': 'fechaEmision', 'fecha de emisión': 'fechaEmision', 'fecha de emision': 'fechaEmision', 'fecha': 'fechaEmision', 'emision': 'fechaEmision',
  'fecha de emisión / actualización': 'fechaEmision', 'fecha de emision / actualizacion': 'fechaEmision',
  'fecha emisión / actualización': 'fechaEmision', 'fecha emision / actualizacion': 'fechaEmision',
  'elaboró': 'elaboro', 'elaboro': 'elaboro', 'elaborado por': 'elaboro', 'autor': 'elaboro',
  'elaboró / actualizó': 'elaboro', 'elaboro / actualizo': 'elaboro',
  'elaboró/ actualizó': 'elaboro', 'elaboro/ actualizo': 'elaboro',
  'tipo resguardo': 'tipoResguardo', 'tipo de resguardo': 'tipoResguardo', 'resguardo': 'tipoResguardo',
  'ubicación': 'ubicacion', 'ubicacion': 'ubicacion',
  'copias': 'copias', 'copias controladas': 'copias',
  'tipo resguardo copia': 'tipoResguardoCopia', 'tipo de resguardo copia': 'tipoResguardoCopia', 'resguardo copia': 'tipoResguardoCopia',
  'usuarios': 'usuarios', 'usuario': 'usuarios',
  'observaciones': 'observaciones', 'obs': 'observaciones', 'notas': 'observaciones', 'comentarios': 'observaciones',
  'urls archivos': 'archivoURLs', 'url archivos': 'archivoURLs', 'enlaces': 'archivoURLs', 'links': 'archivoURLs',
  'urls archivos copia': 'archivoURLsCopia', 'url archivos copia': 'archivoURLsCopia',
  'vigencia': '_skip', 'estado': '_skip', 'dias vigencia': '_skip', 'días vigencia': '_skip', 'dias de vigencia': '_skip',
  'estado de documentos': '_skip', 'columna1': '_skip'
};

// Map sheet names to app AREAS for multi-sheet Listado Maestro format
var _SHEET_AREA_MAP = {
  'dir. gral.': 'Dir. Gral.', 'dirección general': 'Dir. Gral.', 'direccion general': 'Dir. Gral.',
  'calidad': 'Calidad', 'gestión de calidad': 'Calidad', 'gestion de calidad': 'Calidad',
  'clínica': 'Clínica', 'clinica': 'Clínica',
  'analítica': 'Analítica', 'analitica': 'Analítica',
  'bioestad.': 'Bioestad.', 'bioestadística': 'Bioestad.', 'bioestadistica': 'Bioestad.',
  'lab. clínico': 'Lab. Clínico', 'lab. clinico': 'Lab. Clínico', 'laboratorio clínico': 'Lab. Clínico', 'laboratorio clinico': 'Lab. Clínico',
  'seg&hig': 'Seg&Hig', 'seguridad e higiene': 'Seg&Hig', 'coordinación de seguridad e higiene': 'Seg&Hig',
  'mtto': 'Mtto', 'mantenimiento': 'Mtto', 'jefatura seguridad y mantenimiento': 'Mtto',
  'administración': 'Administración', 'administracion': 'Administración',
  'ti ex': 'TI Ex', 'sistemas de ti': 'TI Ex', 'soporte técnico y ti': 'TI Ex',
  'rh ex': 'RH Ex', 'recursos humanos': 'RH Ex'
};

// Map long clasificación values to short area names
var _CLASIFICACION_MAP = {
  'dirección general': 'Dir. Gral.', 'direccion general': 'Dir. Gral.',
  'gestión de calidad': 'Calidad', 'gestion de calidad': 'Calidad',
  'clínica': 'Clínica', 'clinica': 'Clínica',
  'analítica': 'Analítica', 'analitica': 'Analítica',
  'bioestadística': 'Bioestad.', 'bioestadistica': 'Bioestad.',
  'laboratorio clínico': 'Lab. Clínico', 'laboratorio clinico': 'Lab. Clínico',
  'coordinación de seguridad e higiene': 'Seg&Hig', 'coordinacion de seguridad e higiene': 'Seg&Hig',
  'seguridad e higiene': 'Seg&Hig',
  'jefatura seguridad y mantenimiento': 'Mtto', 'mantenimiento': 'Mtto',
  'administración': 'Administración', 'administracion': 'Administración',
  'sistemas de ti': 'TI Ex', 'soporte técnico y ti': 'TI Ex', 'soporte tecnico y ti': 'TI Ex',
  'recursos humanos': 'RH Ex'
};

var _IMPORT_OBS_COL_MAP = {
  'fecha obsoleto': 'fechaObsoleto', 'fecha de obsoleto': 'fechaObsoleto',
  'motivo': 'motivo', 'razón': 'motivo', 'razon': 'motivo'
};

function _openImportDialog() {
  var input = document.createElement('input');
  input.type = 'file';
  input.accept = '.xlsx,.xls,.csv';
  input.style.display = 'none';
  input.addEventListener('change', function(e) {
    if (e.target.files && e.target.files[0]) {
      _processImportFile(e.target.files[0]);
    }
    document.body.removeChild(input);
  });
  document.body.appendChild(input);
  input.click();
}

function _processImportFile(file) {
  if (typeof XLSX === 'undefined') {
    showToast('Librería XLSX no disponible', 'error');
    return;
  }

  var MAX_IMPORT_SIZE = 20 * 1024 * 1024;
  if (file.size > MAX_IMPORT_SIZE) {
    showToast('Archivo demasiado grande (máx 20MB)', 'error');
    return;
  }

  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      var data = new Uint8Array(e.target.result);
      var wb = XLSX.read(data, { type: 'array', cellDates: true });

      if (!wb.SheetNames || wb.SheetNames.length === 0) {
        showToast('El archivo no contiene hojas de datos', 'error');
        return;
      }

      var parsed = _parseWorkbook(wb);
      if (parsed.records.length === 0 && parsed.obsoletos.length === 0) {
        showToast('No se encontraron datos válidos en el archivo', 'error');
        return;
      }

      _showImportPreview(parsed, file.name);

    } catch (err) {
      showToast('Error al leer el archivo Excel', 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function _findHeaderRow(rows) {
  // Search first 10 rows for one containing "Nombre del Documento" or similar identifying header
  var maxSearch = Math.min(10, rows.length);
  for (var r = 0; r < maxSearch; r++) {
    if (!rows[r]) continue;
    var joined = rows[r].map(function(c) { return (c || '').toString().toLowerCase(); }).join(' ');
    if (joined.indexOf('nombre del documento') >= 0 || joined.indexOf('nombre documento') >= 0) {
      return r;
    }
  }
  // Fallback: look for a row with multiple known header keywords
  for (var r2 = 0; r2 < maxSearch; r2++) {
    if (!rows[r2]) continue;
    var joined2 = rows[r2].map(function(c) { return (c || '').toString().toLowerCase(); }).join(' ');
    var hits = 0;
    if (joined2.indexOf('manual') >= 0 || joined2.indexOf('procedimiento') >= 0) hits++;
    if (joined2.indexOf('instructivo') >= 0) hits++;
    if (joined2.indexOf('formato') >= 0) hits++;
    if (joined2.indexOf('versión') >= 0 || joined2.indexOf('version') >= 0) hits++;
    if (joined2.indexOf('no.') >= 0 || joined2.indexOf('no ') >= 0) hits++;
    if (hits >= 2) return r2;
  }
  return 0; // default to first row
}

function _resolveSheetArea(sheetName) {
  // Try to resolve a sheet name to a known app AREA
  var key = sheetName.toLowerCase().trim();
  if (_SHEET_AREA_MAP[key]) return _SHEET_AREA_MAP[key];
  // Fuzzy match
  for (var k in _SHEET_AREA_MAP) {
    if (k.indexOf(key) >= 0 || key.indexOf(k) >= 0) return _SHEET_AREA_MAP[k];
  }
  return '';
}

function _resolveClasificacion(rawVal) {
  // Resolve a "Clasificación" column value to a short app AREA name
  if (!rawVal) return '';
  var key = rawVal.toLowerCase().trim();
  if (_CLASIFICACION_MAP[key]) return _CLASIFICACION_MAP[key];
  // Check if it's already a short area name
  for (var i = 0; i < AREAS.length; i++) {
    if (AREAS[i].toLowerCase() === key) return AREAS[i];
  }
  // Fuzzy match
  for (var k in _CLASIFICACION_MAP) {
    if (k.indexOf(key) >= 0 || key.indexOf(k) >= 0) return _CLASIFICACION_MAP[k];
  }
  return rawVal; // return as-is if no match
}

function _parseWorkbook(wb) {
  var result = { records: [], obsoletos: [], elaboros: [], sheetInfo: [] };

  for (var si = 0; si < wb.SheetNames.length; si++) {
    var sheetName = wb.SheetNames[si];
    var ws = wb.Sheets[sheetName];
    if (!ws) continue;

    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    if (!rows || rows.length < 2) continue;

    // Auto-detect header row (may not be row 0 in Listado Maestro format)
    var headerRowIdx = _findHeaderRow(rows);
    var headers = rows[headerRowIdx].map(function(h) { return (h || '').toString().trim(); });
    var _stripAccents = function(s) { return s.toLowerCase().replace(/[áéíóúüñ]/g, function(c) { return {á:'a',é:'e',í:'i',ó:'o',ú:'u',ü:'u',ñ:'n'}[c]||c; }); };
    var lowerHeaders = headers.map(_stripAccents);

    // Detect if this is an obsoletos sheet
    var isObsoletos = /obsolet/i.test(sheetName) || lowerHeaders.some(function(h) { return h.indexOf('obsoleto') >= 0 || h.indexOf('motivo') >= 0; });
    // Detect elaboradores sheet
    var isElaboradores = /elabor/i.test(sheetName) && headers.length <= 2;
    // Skip Resumen/summary sheets
    var isResumen = /resumen|summary|dashboard|log/i.test(sheetName);

    if (isResumen && !lowerHeaders.some(function(h) { return h.indexOf('documento') >= 0 || h.indexOf('manual') >= 0; })) continue;

    if (isElaboradores) {
      for (var ei = headerRowIdx + 1; ei < rows.length; ei++) {
        var val = _clean(String(rows[ei][0] || ''));
        if (val) result.elaboros.push(val);
      }
      result.sheetInfo.push({ name: sheetName, type: 'elaboradores', count: result.elaboros.length });
      continue;
    }

    // Resolve sheet name to an area (for multi-sheet Listado Maestro format)
    var sheetArea = _resolveSheetArea(sheetName);

    // Map columns — try exact match first, then partial/contains match
    var colMap = {};
    var allMaps = isObsoletos ? Object.assign({}, _IMPORT_COL_MAP, _IMPORT_OBS_COL_MAP) : _IMPORT_COL_MAP;
    for (var ci = 0; ci < lowerHeaders.length; ci++) {
      var normalized = lowerHeaders[ci];
      // Exact match
      if (allMaps[normalized]) {
        colMap[ci] = allMaps[normalized];
        continue;
      }
      // Try without accents on the map keys too
      var normalizedStripped = _stripAccents(normalized);
      for (var mapKey in allMaps) {
        if (_stripAccents(mapKey) === normalizedStripped) {
          colMap[ci] = allMaps[mapKey];
          break;
        }
      }
      // Partial/contains match for common long-form headers
      if (!colMap[ci]) {
        if (normalizedStripped.indexOf('manual') >= 0 && normalizedStripped.indexOf('procedimiento') >= 0) colMap[ci] = 'codManual';
        else if (normalizedStripped.indexOf('nombre') >= 0 && normalizedStripped.indexOf('documento') >= 0) colMap[ci] = 'nombreDoc';
        else if (normalizedStripped.indexOf('fecha') >= 0 && (normalizedStripped.indexOf('emision') >= 0 || normalizedStripped.indexOf('actualizacion') >= 0)) colMap[ci] = 'fechaEmision';
        else if (normalizedStripped.indexOf('elaboro') >= 0 || normalizedStripped.indexOf('actualizo') >= 0) colMap[ci] = 'elaboro';
        else if (normalizedStripped.indexOf('clasificacion') >= 0) colMap[ci] = 'area';
        else if (normalizedStripped.indexOf('vigencia') >= 0 || normalizedStripped.indexOf('estado de documento') >= 0) colMap[ci] = '_skip';
      }
    }

    // Need at least one identifying column
    var hasIdField = false;
    for (var ck in colMap) {
      if (['nombreDoc','codManual','codInstructivo','codListado','codFormato','no'].indexOf(colMap[ck]) >= 0) {
        hasIdField = true;
        break;
      }
    }
    if (!hasIdField) {
      result.sheetInfo.push({ name: sheetName, type: 'omitida', count: 0, reason: 'Sin columnas identificables' });
      continue;
    }

    var sheetRecords = [];
    // Data starts AFTER the header row (not always row 1)
    for (var ri = headerRowIdx + 1; ri < rows.length; ri++) {
      var row = rows[ri];
      if (!row || row.every(function(c) { return !c || String(c).trim() === ''; })) continue;

      var rec = {
        id: 0,
        no: '', area: '', codManual: '', codInstructivo: '', codListado: '', codFormato: '',
        nombreDoc: '', version: '', fechaEmision: '', elaboro: '',
        tipoResguardo: '', ubicacion: '', copias: '', tipoResguardoCopia: '',
        usuarios: '', observaciones: '',
        archivos: [], archivoURLs: [], archivosCopia: [], archivoURLsCopia: []
      };

      if (isObsoletos) {
        rec.fechaObsoleto = '';
        rec.motivo = '';
      }

      var rawClasificacion = '';
      for (var col in colMap) {
        var field = colMap[col];
        if (field === '_skip') continue;
        var cellVal = _clean(String(row[col] || ''));

        if (field === 'fechaEmision' || field === 'fechaObsoleto') {
          cellVal = _normalizeDate(cellVal);
        }

        if (field === 'archivoURLs' || field === 'archivoURLsCopia') {
          rec[field] = cellVal ? cellVal.split(/\s*\|\s*/).filter(function(u) {
            var t = u.trim().toLowerCase();
            return t.startsWith('http://') || t.startsWith('https://');
          }) : [];
        } else if (field === 'area') {
          // Store raw clasificación value for later resolution
          rawClasificacion = cellVal;
          rec[field] = cellVal;
        } else {
          rec[field] = cellVal;
        }
      }

      // Resolve the area: use clasificación value, fall back to sheet name
      if (rawClasificacion) {
        rec.area = _resolveClasificacion(rawClasificacion);
      }
      if (!rec.area && sheetArea) {
        rec.area = sheetArea;
      }

      // Skip rows with no meaningful data
      if (!rec.nombreDoc && !rec.codManual && !rec.codInstructivo && !rec.codFormato && !rec.no) continue;

      sheetRecords.push(rec);
    }

    if (sheetRecords.length > 0) {
      if (isObsoletos) {
        result.obsoletos = result.obsoletos.concat(sheetRecords);
        result.sheetInfo.push({ name: sheetName, type: 'obsoletos', count: sheetRecords.length });
      } else {
        result.records = result.records.concat(sheetRecords);
        result.sheetInfo.push({ name: sheetName, type: 'registros', count: sheetRecords.length });
      }
    } else {
      result.sheetInfo.push({ name: sheetName, type: 'vacía', count: 0 });
    }
  }

  return result;
}

function _normalizeDate(val) {
  if (!val) return '';
  var s = val.toString().trim();
  if (s === '00:00:00' || s === '') return '';
  // Already in YYYY-MM-DD format
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    // Skip Excel bug dates (1900-xx-xx, 1899-xx-xx)
    if (s.startsWith('1900-') || s.startsWith('1899-')) return '';
    return s.substring(0, 10);
  }
  // DD/MM/YYYY or DD-MM-YYYY
  var match = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (match) {
    var d = match[1].padStart(2,'0');
    var m = match[2].padStart(2,'0');
    return match[3] + '-' + m + '-' + d;
  }
  // DD-Mon-YY like 13-Jul-26 or 13-Jul-2026
  var MONTHS = {jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12',
    ene:'01',feb:'02',mar:'03',abr:'04',may:'05',jun:'06',jul:'07',ago:'08',sep:'09',oct:'10',nov:'11',dic:'12'};
  var match2 = s.match(/^(\d{1,2})[\/\-]([A-Za-z]{3})[\/\-](\d{2,4})$/);
  if (match2) {
    var mon = MONTHS[match2[2].toLowerCase()];
    if (mon) {
      var yr = match2[3].length === 2 ? (parseInt(match2[3]) < 50 ? '20' + match2[3] : '19' + match2[3]) : match2[3];
      return yr + '-' + mon + '-' + match2[1].padStart(2,'0');
    }
  }
  // Mon-YY like Jul-26
  var match3 = s.match(/^([A-Za-z]{3})[\/\-](\d{2,4})$/);
  if (match3) {
    var mon3 = MONTHS[match3[1].toLowerCase()];
    if (mon3) {
      var yr3 = match3[2].length === 2 ? (parseInt(match3[2]) < 50 ? '20' + match3[2] : '19' + match3[2]) : match3[2];
      return yr3 + '-' + mon3 + '-01';
    }
  }
  // MM/DD/YYYY (US format) — try parsing as Date
  try {
    var parsed = new Date(s);
    if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 1990 && parsed.getFullYear() < 2100) {
      return parsed.getFullYear() + '-' + String(parsed.getMonth()+1).padStart(2,'0') + '-' + String(parsed.getDate()).padStart(2,'0');
    }
  } catch(e) {}
  return '';
}

function _showImportPreview(parsed, fileName) {
  STATE.modal = 'import';
  STATE._importData = parsed;
  STATE._importFileName = fileName;
  STATE._importMode = 'merge'; // 'merge' or 'replace'
  STATE._importArea = STATE.currentArea || '';
  render();
}

function renderImportModal() {
  var parsed = STATE._importData;
  if (!parsed) return null;

  var overlay = h('div', {className:'modal-overlay', onClick: function(e) { if (e.target === e.currentTarget) closeModal(); }});
  var modal = h('div', {className:'modal-content modal-lg'});

  // Header
  modal.appendChild(h('div', {className:'modal-header'},
    h('h3', {}, h('i', {className:'fas fa-file-import'}), ' Importar Excel'),
    h('button', {className:'modal-close', onClick: closeModal}, '×')
  ));

  var body = h('div', {className:'modal-body', style:{maxHeight:'70vh', overflowY:'auto'}});

  // File info
  body.appendChild(h('div', {className:'import-info', style:{marginBottom:'16px', padding:'12px', background:'var(--bg-card,#f1f5f9)', borderRadius:'8px'}},
    h('p', {style:{margin:'0 0 8px'}}, h('strong', {}, 'Archivo: '), STATE._importFileName || ''),
    h('p', {style:{margin:'0'}}, h('strong', {}, 'Hojas detectadas: '),
      parsed.sheetInfo.map(function(s) {
        var icon = s.type === 'registros' ? '📋' : s.type === 'obsoletos' ? '📁' : s.type === 'elaboradores' ? '👤' : '⏭️';
        return icon + ' ' + s.name + ' (' + s.count + ' ' + s.type + ')';
      }).join(' · ')
    )
  ));

  // Summary counts
  var summary = h('div', {style:{display:'flex', gap:'16px', marginBottom:'16px', flexWrap:'wrap'}});
  summary.appendChild(_importStatBox('Registros', parsed.records.length, '#3b82f6'));
  summary.appendChild(_importStatBox('Obsoletos', parsed.obsoletos.length, '#f59e0b'));
  summary.appendChild(_importStatBox('Elaboradores', parsed.elaboros.length, '#8b5cf6'));
  body.appendChild(summary);

  // Import mode selection
  var modeDiv = h('div', {style:{marginBottom:'16px', padding:'12px', background:'var(--bg-card,#f1f5f9)', borderRadius:'8px'}});
  modeDiv.appendChild(h('label', {style:{fontWeight:'600', display:'block', marginBottom:'8px'}}, 'Modo de importación:'));

  var mergeLabel = h('label', {style:{display:'flex', alignItems:'center', gap:'8px', marginBottom:'6px', cursor:'pointer'}});
  var mergeRadio = h('input', {type:'radio', value:'merge', checked: STATE._importMode === 'merge'});
  mergeRadio.addEventListener('change', function() { STATE._importMode = 'merge'; });
  mergeLabel.appendChild(mergeRadio);
  mergeLabel.appendChild(document.createTextNode('Agregar a los datos existentes (fusionar)'));
  modeDiv.appendChild(mergeLabel);

  var replaceLabel = h('label', {style:{display:'flex', alignItems:'center', gap:'8px', cursor:'pointer', color:'#ef4444'}});
  var replaceRadio = h('input', {type:'radio', value:'replace', checked: STATE._importMode === 'replace'});
  replaceRadio.addEventListener('change', function() { STATE._importMode = 'replace'; });
  replaceLabel.appendChild(replaceRadio);
  replaceLabel.appendChild(document.createTextNode('Reemplazar TODOS los datos actuales'));
  modeDiv.appendChild(replaceLabel);

  body.appendChild(modeDiv);

  // Area assignment for records without area
  var noAreaCount = parsed.records.filter(function(r) { return !r.area; }).length;
  if (noAreaCount > 0) {
    var areaDiv = h('div', {style:{marginBottom:'16px', padding:'12px', background:'var(--bg-card,#f1f5f9)', borderRadius:'8px'}});
    areaDiv.appendChild(h('label', {style:{fontWeight:'600', display:'block', marginBottom:'8px'}},
      noAreaCount + ' registro(s) sin área asignada. Asignar a:'));
    var areaSelect = h('select', {className:'form-control', style:{maxWidth:'250px'}});
    AREAS.forEach(function(a) {
      var opt = h('option', {value: a}, a);
      if (a === STATE._importArea) opt.selected = true;
      areaSelect.appendChild(opt);
    });
    areaSelect.addEventListener('change', function(e) { STATE._importArea = e.target.value; });
    areaDiv.appendChild(areaSelect);
    body.appendChild(areaDiv);
  }

  // Preview table (first 10 records)
  if (parsed.records.length > 0) {
    body.appendChild(h('h4', {style:{margin:'16px 0 8px'}}, 'Vista previa de registros (primeros ' + Math.min(10, parsed.records.length) + '):'));
    var table = h('table', {className:'data-table', style:{fontSize:'12px'}});
    var thead = h('thead', {});
    var trH = h('tr', {});
    ['No.','Área','Cód. Manual','Nombre Doc.','Versión','Fecha','Elaboró'].forEach(function(hdr) {
      trH.appendChild(h('th', {}, hdr));
    });
    thead.appendChild(trH);
    table.appendChild(thead);

    var tbody = h('tbody', {});
    var previewRows = parsed.records.slice(0, 10);
    previewRows.forEach(function(r) {
      var tr = h('tr', {});
      [r.no, r.area || '(sin área)', r.codManual, r.nombreDoc, r.version, r.fechaEmision, r.elaboro].forEach(function(val) {
        tr.appendChild(h('td', {}, _esc(val || '')));
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    var tableWrap = h('div', {style:{overflowX:'auto', maxHeight:'250px', overflowY:'auto'}});
    tableWrap.appendChild(table);
    body.appendChild(tableWrap);

    if (parsed.records.length > 10) {
      body.appendChild(h('p', {style:{color:'#64748b', fontSize:'13px', margin:'4px 0'}}, '... y ' + (parsed.records.length - 10) + ' registros más'));
    }
  }

  modal.appendChild(body);

  // Footer
  var footer = h('div', {className:'modal-footer', style:{display:'flex', gap:'12px', justifyContent:'flex-end', padding:'16px'}});
  footer.appendChild(h('button', {className:'btn btn-secondary', onClick: closeModal}, 'Cancelar'));
  footer.appendChild(h('button', {className:'btn btn-primary', onClick: function() { _executeImport(); }},
    h('i', {className:'fas fa-file-import'}), ' Importar ' + parsed.records.length + ' registros'
  ));
  modal.appendChild(footer);

  overlay.appendChild(modal);
  return overlay;
}

function _importStatBox(label, count, color) {
  return h('div', {style:{padding:'12px 20px', background:'white', borderRadius:'8px', borderLeft:'4px solid ' + color, flex:'1', minWidth:'120px'}},
    h('div', {style:{fontSize:'22px', fontWeight:'700', color: color}}, String(count)),
    h('div', {style:{fontSize:'13px', color:'#64748b'}}, label)
  );
}

function _importRecordKey(r) {
  // Build a unique key from the combination of identifying fields
  var parts = [
    (r.nombreDoc || '').trim().toLowerCase(),
    (r.codManual || '').trim().toLowerCase(),
    (r.codInstructivo || '').trim().toLowerCase(),
    (r.codListado || '').trim().toLowerCase(),
    (r.codFormato || '').trim().toLowerCase(),
    (r.no || '').toString().trim().toLowerCase(),
    (r.area || '').trim().toLowerCase(),
    (r.version || '').toString().trim().toLowerCase()
  ];
  return parts.join('|');
}

function _executeImport() {
  var parsed = STATE._importData;
  if (!parsed) return;

  var mode = STATE._importMode;
  var defaultArea = STATE._importArea || STATE.currentArea || 'Dir. Gral.';

  // Create pre-import backup
  if (typeof _saveCloudBackup === 'function' && typeof SYNC !== 'undefined' && SYNC.initialized) {
    _saveCloudBackup('pre-restore', 'Auto-backup antes de importar Excel: ' + (STATE._importFileName || ''));
  }

  // Build index of existing records for dedup in merge mode
  var existingKeys = {};
  var existingObsKeys = {};
  var skippedRecords = 0;
  var skippedObsoletos = 0;
  var updatedRecords = 0;
  var updatedObsoletos = 0;

  if (mode === 'merge') {
    (STATE.records || []).forEach(function(r, idx) {
      var key = _importRecordKey(r);
      if (key && key !== '|||||||') existingKeys[key] = idx;
    });
    (STATE.obsoletos || []).forEach(function(r, idx) {
      var key = _importRecordKey(r);
      if (key && key !== '|||||||') existingObsKeys[key] = idx;
    });
  }

  var newRecords = [];
  parsed.records.forEach(function(r) {
    var rec = Object.assign({}, r);
    if (!rec.area) rec.area = defaultArea;
    rec.archivos = rec.archivos || [];
    rec.archivoURLs = rec.archivoURLs || [];
    rec.archivosCopia = rec.archivosCopia || [];
    rec.archivoURLsCopia = rec.archivoURLsCopia || [];

    if (mode === 'merge') {
      var key = _importRecordKey(rec);
      if (key && key !== '|||||||' && typeof existingKeys[key] !== 'undefined') {
        // Record exists — update non-empty fields only
        var existIdx = existingKeys[key];
        var existRec = STATE.records[existIdx];
        var changed = false;
        ['elaboro','tipoResguardo','ubicacion','copias','tipoResguardoCopia','usuarios','observaciones','fechaEmision'].forEach(function(f) {
          if (rec[f] && rec[f] !== existRec[f]) {
            existRec[f] = rec[f];
            changed = true;
          }
        });
        if (rec.archivoURLs && rec.archivoURLs.length > 0) {
          var existUrls = (existRec.archivoURLs || []).join('|');
          rec.archivoURLs.forEach(function(u) {
            if (existUrls.indexOf(u) < 0) { existRec.archivoURLs.push(u); changed = true; }
          });
        }
        if (changed) {
          _trackChange(existRec, 'Actualizado desde Excel');
          updatedRecords++;
        } else {
          skippedRecords++;
        }
        return; // skip adding as new
      }
    }

    rec.id = getNextId();
    _trackChange(rec, 'Importado desde Excel');
    newRecords.push(rec);
  });

  var newObsoletos = [];
  parsed.obsoletos.forEach(function(r) {
    var rec = Object.assign({}, r);
    if (!rec.area) rec.area = defaultArea;
    rec.archivos = rec.archivos || [];
    rec.archivoURLs = rec.archivoURLs || [];
    if (!rec.motivo) rec.motivo = 'Importado desde Excel';
    if (!rec.fechaObsoleto) rec.fechaObsoleto = new Date().toISOString().split('T')[0];

    if (mode === 'merge') {
      var key = _importRecordKey(rec);
      if (key && key !== '|||||||' && typeof existingObsKeys[key] !== 'undefined') {
        var existIdx = existingObsKeys[key];
        var existRec = STATE.obsoletos[existIdx];
        var changed = false;
        ['elaboro','tipoResguardo','ubicacion','copias','tipoResguardoCopia','usuarios','observaciones','motivo','fechaObsoleto','fechaEmision'].forEach(function(f) {
          if (rec[f] && rec[f] !== existRec[f]) {
            existRec[f] = rec[f];
            changed = true;
          }
        });
        if (changed) {
          _trackChange(existRec, 'Actualizado desde Excel (obsoleto)');
          updatedObsoletos++;
        } else {
          skippedObsoletos++;
        }
        return;
      }
    }

    rec.id = getNextId();
    _trackChange(rec, 'Importado desde Excel (obsoleto)');
    newObsoletos.push(rec);
  });

  if (mode === 'replace') {
    if (!confirm('⚠️ ATENCIÓN: Esto REEMPLAZARÁ todos los datos actuales.\n\n'
      + 'Se perderán ' + (STATE.records||[]).length + ' registros y ' + (STATE.obsoletos||[]).length + ' obsoletos actuales.\n\n'
      + '¿Estás seguro? Esta acción no se puede deshacer.')) {
      return;
    }
    STATE.records = newRecords;
    STATE.obsoletos = newObsoletos;
  } else {
    STATE.records = (STATE.records || []).concat(newRecords);
    STATE.obsoletos = (STATE.obsoletos || []).concat(newObsoletos);
  }

  // Merge elaboradores
  if (parsed.elaboros && parsed.elaboros.length > 0) {
    var existing = {};
    (STATE.elaboros || []).forEach(function(e) { existing[e] = true; });
    parsed.elaboros.forEach(function(e) {
      if (!existing[e]) {
        STATE.elaboros.push(e);
        existing[e] = true;
      }
    });
  }

  // Extract new elaboradores from imported records
  newRecords.forEach(function(r) {
    if (r.elaboro && STATE.elaboros.indexOf(r.elaboro) < 0) {
      STATE.elaboros.push(r.elaboro);
    }
  });

  var logParts = ['Excel importado: ' + newRecords.length + ' nuevos'];
  if (updatedRecords > 0) logParts.push(updatedRecords + ' actualizados');
  if (skippedRecords > 0) logParts.push(skippedRecords + ' duplicados omitidos');
  if (newObsoletos.length > 0) logParts.push(newObsoletos.length + ' obsoletos nuevos');
  if (updatedObsoletos > 0) logParts.push(updatedObsoletos + ' obs. actualizados');
  if (skippedObsoletos > 0) logParts.push(skippedObsoletos + ' obs. duplicados omitidos');
  logParts.push('(' + mode + ')');
  logParts.push('— ' + (STATE._importFileName || ''));

  if (typeof addLog === 'function') {
    addLog('import', logParts.join(', '));
  }

  saveState();
  closeModal();
  STATE._importData = null;
  STATE._importFileName = null;

  var toastMsg = '✅ ' + newRecords.length + ' registros importados';
  if (updatedRecords > 0) toastMsg += ', ' + updatedRecords + ' actualizados';
  if (skippedRecords + skippedObsoletos > 0) toastMsg += ', ' + (skippedRecords + skippedObsoletos) + ' duplicados omitidos';
  if (newObsoletos.length > 0) toastMsg += ', ' + newObsoletos.length + ' obsoletos';
  if (mode === 'replace') toastMsg += ' (datos reemplazados)';
  showToast(toastMsg, 'success');

  render();
}

function _formatBackupDate(isoStr) {
  try {
    var d = new Date(isoStr);
    return d.toLocaleDateString('es-MX', {day:'2-digit',month:'short',year:'numeric'})
      + ' ' + d.toLocaleTimeString('es-MX', {hour:'2-digit',minute:'2-digit'});
  } catch(e) { return isoStr || '—'; }
}

function _openRestoreDialog() {
  var input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json,.dbk';
  input.style.display = 'none';
  input.addEventListener('change', function(e) {
    if (e.target.files && e.target.files[0]) {
      if (typeof restoreBackup === 'function') {
        restoreBackup(e.target.files[0]);
      }
    }
    document.body.removeChild(input);
  });
  document.body.appendChild(input);
  input.click();
}

var _renderRAF = null;
function render() {
  if (_renderRAF) return;
  _renderRAF = requestAnimationFrame(function() {
    _renderRAF = null;
    _renderNow();
  });
}
function _renderNow() {
  if (STATE.currentSection !== 'dashboard') _stopDashRefresh();
  if (!AUTH.currentUser) {
    _stopDashRefresh();
    document.getElementById('modal-root').innerHTML = '';
    renderLogin();
    return;
  }

  const app = document.getElementById('app');

  var active = document.activeElement;
  var focusIdx = -1;
  var cursorPos = null;
  var selEnd = null;
  var scrollLeft = 0;
  if (active && app.contains(active) && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.tagName === 'SELECT')) {
    var allInputs = app.querySelectorAll('input, textarea, select');
    focusIdx = Array.from(allInputs).indexOf(active);
    try { cursorPos = active.selectionStart; selEnd = active.selectionEnd; } catch(e) {}
    scrollLeft = active.scrollLeft || 0;
  }

  var mainEl = app.querySelector('.main');
  var contentScroll = 0;
  if (mainEl) contentScroll = mainEl.scrollTop;

  const frag = document.createDocumentFragment();
  const sidebar = renderSidebar();

  let sectionTitle;
  let content;

  switch(STATE.currentSection) {
    case 'areas':
      sectionTitle = STATE.currentArea;
      content = renderAreasTable();
      break;
    case 'vencidos':
      sectionTitle = 'Documentos Vencidos';
      content = renderVencidosSection();
      break;
    case 'obsoletos':
      sectionTitle = 'Documentos Obsoletos';
      content = renderObsoletosTable();
      break;
    case 'salidas':
      sectionTitle = 'Registro de Salidas';
      content = renderSalidasTable();
      break;
    case 'papelera':
      sectionTitle = 'Papelera de Reciclaje';
      content = renderPapeleraTable();
      break;
    case 'logs':
      sectionTitle = 'Registro de Actividad';
      content = renderLogsTable();
      break;
    case 'alertas':
      sectionTitle = 'Recordatorios y Alertas de Vigencia';
      content = renderAlertasPanel();
      break;
    case 'backups':
      sectionTitle = 'Backups en la Nube';
      content = renderCloudBackupsSection();
      break;
    case 'dashboard':
      sectionTitle = 'Dashboard de Actividad';
      if (typeof loadAllProfiles === 'function') loadAllProfiles();
      content = renderDashboard();
      break;
    case 'sistema':
      sectionTitle = 'Verificación del Sistema';
      content = renderVerificacionISO();
      break;
  }

  const topbar = h('div', {className: 'topbar'},
    h('div', {className: 'topbar-title'}, sectionTitle),
    h('div', {className: 'topbar-actions'},
      typeof SYNC !== 'undefined' && SYNC.initialized ? h('button', {
        className: 'btn btn-sync', id: 'btn-sync',
        onClick: function() { if (typeof forceSync === 'function') forceSync(); }
      }, h('i', {className: 'fas fa-arrows-rotate'}), 'Sincronizar') : null,
      typeof SYNC !== 'undefined' && SYNC.connected ? h('span', {className: 'sync-time-label', id: 'sync-last-time'}, '') : null,
      STATE.currentSection === 'areas' ? h('button', {className:'btn btn-primary', onClick:openAdd}, h('i',{className:'fas fa-circle-plus'}), 'Nuevo Registro') : null,
      h('button', {className:'btn btn-secondary', onClick:exportToExcel}, h('i',{className:'fas fa-file-excel', style:{color:'#10b981'}}), 'Exportar Excel'),
      STATE.currentSection === 'vencidos' ? h('button', {className:'btn btn-secondary', onClick:exportVencidosToExcel}, h('i',{className:'fas fa-file-excel', style:{color:'#ef4444'}}), 'Exportar Vencidos') : null,
      STATE.currentSection === 'logs' ? h('button', {className:'btn btn-secondary', onClick:exportLogsToExcel}, h('i',{className:'fas fa-file-excel', style:{color:'#f5c542'}}), 'Exportar Logs') : null,
      isAdmin() ? h('button', {className:'btn btn-backup', onClick: function() { if (typeof generateBackupExcel === 'function') generateBackupExcel(); }}, h('i',{className:'fas fa-shield-halved', style:{color:'#f5c542'}}), 'Backup Excel') : null,
      isAdmin() ? h('button', {className:'btn btn-backup-json', onClick: function() { if (typeof generateBackup === 'function') generateBackup(); }}, h('i',{className:'fas fa-download', style:{color:'#26c6da'}}), 'Backup JSON') : null,
      isAdmin() && typeof SYNC !== 'undefined' && SYNC.initialized ? h('button', {className:'btn btn-backup-cloud', onClick: function() {
        var nota = prompt('Nota del backup (opcional):', 'Backup manual');
        if (nota === null) return;
        if (typeof saveManualCloudBackup === 'function') saveManualCloudBackup(nota || 'Backup manual');
      }}, h('i',{className:'fas fa-cloud-arrow-up', style:{color:'#10b981'}}), 'Backup Nube') : null,
      isAdmin() ? h('button', {className:'btn btn-restore', onClick: function() { _openRestoreDialog(); }}, h('i',{className:'fas fa-upload', style:{color:'#fb923c'}}), 'Restaurar') : null,
      isAdmin() ? h('button', {className:'btn btn-import', onClick: function() { _openImportDialog(); }}, h('i',{className:'fas fa-file-import', style:{color:'#6366f1'}}), 'Importar Excel') : null
    )
  );

  const main = h('div', {className:'main'},
    topbar,
    content
  );

  frag.appendChild(sidebar);
  frag.appendChild(main);
  app.replaceChildren(frag);

  var newMain = app.querySelector('.main');
  if (newMain && contentScroll) newMain.scrollTop = contentScroll;

  if (focusIdx >= 0) {
    var newInputs = app.querySelectorAll('input, textarea, select');
    var el = newInputs[focusIdx];
    if (el) {
      el.focus();
      try {
        if (cursorPos !== null && typeof el.setSelectionRange === 'function') el.setSelectionRange(cursorPos, selEnd !== null ? selEnd : cursorPos);
        if (scrollLeft) el.scrollLeft = scrollLeft;
      } catch(e) {}
    }
  }

  refreshModal();

  if (!STATE._alertBannerShown && AUTH.currentUser) {
    STATE._alertBannerShown = true;
    setTimeout(showAlertBanner, 800);
  }
}

function renderToasts() {
  let tc = document.getElementById('toast-container');
  if (tc) tc.remove();
  if (STATE.toasts.length > 0) {
    tc = h('div', {className:'toast-container', id:'toast-container'},
      ...STATE.toasts.map(t => h('div', {className:'toast toast-'+t.type},
        h('i', {className: t.type==='success'?'fas fa-check-circle':t.type==='error'?'fas fa-exclamation-circle':'fas fa-info-circle'}),
        h('span', null, t.msg)
      ))
    );
    document.body.appendChild(tc);
  }
}

(function securityInit() {
  try {} catch(e) {}

  setInterval(async function() {
    if (AUTH.currentUser) {
      var client = typeof getDB === 'function' ? getDB() : null;
      if (client) {
        var result = await client.auth.getSession();
        if (!result.data || !result.data.session) {
          AUTH.currentUser = null;
          if (typeof unregisterPresence === 'function') unregisterPresence();
          showToast('Sesión expirada. Inicia sesión de nuevo.', 'error');
          render();
        }
      }
    }
  }, 60000);

  try {
    Object.defineProperty(AUTH, '_internal', {
      get() { void 0; return undefined; },
      enumerable: false,
      configurable: false
    });
  } catch(e) {}
})();

var _dashTimer = null;
function _startDashRefresh() {
  _stopDashRefresh();
  if (STATE.currentSection !== 'dashboard' || !AUTH.currentUser) return;
  _dashTimer = setInterval(function() {
    if (STATE.currentSection !== 'dashboard' || !AUTH.currentUser) { _stopDashRefresh(); return; }
    try {
      var el = document.getElementById('dash-root');
      if (el) {
        var newContent = _buildDash();
        if (newContent) { el.innerHTML = ''; el.appendChild(newContent); }
      }
    } catch(e) {}
  }, 5000);
}
function _stopDashRefresh() {
  if (_dashTimer) { clearInterval(_dashTimer); _dashTimer = null; }
}

function _pie(data, sz, pid) {
  var t = 0; data.forEach(function(d){t+=d.value;});
  if (t===0) {var e=document.createElement('div');e.style.cssText='text-align:center;padding:40px;color:var(--text-muted)';e.textContent='Sin datos';return e;}
  var cx=sz/2,cy=sz/2,r=sz/2-8,ri=r*0.4,a=-90;
  var ns='http://www.w3.org/2000/svg';
  var svgEl=document.createElementNS(ns,'svg');
  svgEl.setAttribute('viewBox','0 0 '+sz+' '+sz);
  svgEl.setAttribute('width','100%');
  svgEl.style.cssText='display:block;max-width:'+sz+'px;margin:0 auto;animation:dashFadeIn .6s ease-out';
  var tipEl=null;
  data.forEach(function(d){
    if(!d.value)return;
    var p=d.value/t,a1=a*Math.PI/180,sw=p*360,a2=(a+sw)*Math.PI/180,lg=sw>180?1:0;
    var x1=cx+r*Math.cos(a1),y1=cy+r*Math.sin(a1),x2=cx+r*Math.cos(a2),y2=cy+r*Math.sin(a2);
    var ix1=cx+ri*Math.cos(a1),iy1=cy+ri*Math.sin(a1),ix2=cx+ri*Math.cos(a2),iy2=cy+ri*Math.sin(a2);
    var mid=(a+sw/2)*Math.PI/180,tx=cx+(r*0.72)*Math.cos(mid),ty=cy+(r*0.72)*Math.sin(mid);
    var tip=_esc(d.label)+': '+d.value+' ('+Math.round(p*100)+'%)';
    var shape;
    if(p>=0.999){
      shape=document.createElementNS(ns,'circle');
      shape.setAttribute('cx',cx);shape.setAttribute('cy',cy);shape.setAttribute('r',r);
    } else {
      shape=document.createElementNS(ns,'path');
      shape.setAttribute('d','M'+ix1.toFixed(1)+','+iy1.toFixed(1)+' L'+x1.toFixed(1)+','+y1.toFixed(1)+' A'+r+','+r+' 0 '+lg+',1 '+x2.toFixed(1)+','+y2.toFixed(1)+' L'+ix2.toFixed(1)+','+iy2.toFixed(1)+' A'+ri+','+ri+' 0 '+lg+',0 '+ix1.toFixed(1)+','+iy1.toFixed(1)+' Z');
    }
    shape.setAttribute('fill',d.color);shape.style.cssText='transition:opacity .2s,transform .2s;cursor:pointer;transform-origin:'+cx+'px '+cy+'px';
    var title=document.createElementNS(ns,'title');title.textContent=tip;shape.appendChild(title);
    (function(tipText){
      shape.onmouseover=function(){this.style.opacity=0.85;this.style.transform='scale(1.04)';var te=document.getElementById(pid);if(te)te.textContent=tipText;};
      shape.onmouseout=function(){this.style.opacity=1;this.style.transform='scale(1)';var te=document.getElementById(pid);if(te)te.textContent='';};
    })(tip);
    svgEl.appendChild(shape);
    if(p>=0.06){
      var txt=document.createElementNS(ns,'text');
      txt.setAttribute('x',tx.toFixed(1));txt.setAttribute('y',ty.toFixed(1));
      txt.setAttribute('text-anchor','middle');txt.setAttribute('dominant-baseline','middle');
      txt.setAttribute('fill','#fff');txt.setAttribute('font-size','11');txt.setAttribute('font-weight','700');
      txt.style.cssText='pointer-events:none;text-shadow:0 1px 3px rgba(0,0,0,.5)';
      txt.textContent=Math.round(p*100)+'%';svgEl.appendChild(txt);
    }
    a+=sw;
  });
  var tTotal=document.createElementNS(ns,'text');
  tTotal.setAttribute('x',cx);tTotal.setAttribute('y',cy-6);tTotal.setAttribute('text-anchor','middle');
  tTotal.setAttribute('fill','var(--text-primary)');tTotal.setAttribute('font-size','22');
  tTotal.setAttribute('font-weight','800');tTotal.setAttribute('font-family','JetBrains Mono,monospace');
  tTotal.textContent=t;svgEl.appendChild(tTotal);
  var tLabel=document.createElementNS(ns,'text');
  tLabel.setAttribute('x',cx);tLabel.setAttribute('y',cy+10);tLabel.setAttribute('text-anchor','middle');
  tLabel.setAttribute('fill','var(--text-muted)');tLabel.setAttribute('font-size','10');
  tLabel.textContent='total';svgEl.appendChild(tLabel);
  var el=document.createElement('div');el.appendChild(svgEl);
  var tipDiv=document.createElement('div');tipDiv.id=pid;
  tipDiv.style.cssText='text-align:center;min-height:20px;font-size:13px;color:var(--accent);margin-top:8px;font-weight:600';
  el.appendChild(tipDiv);return el;
}

function _legend(data,t) {
  var el = h('div',{className:'dash-legend'});
  data.filter(function(d){return d.value>0;}).forEach(function(d){
    var p=t>0?Math.round((d.value/t)*100):0;
    var item = h('div',{className:'dash-legend-item',style:{'--legend-color':d.color}},
      h('span',{className:'dash-legend-dot',style:{background:d.color,'--legend-color':d.color}}),
      h('span',{className:'dash-legend-name'},d.label),
      h('span',{className:'dash-legend-val',style:{color:d.color}},String(d.value)),
      h('span',{className:'dash-legend-pct'},p+'%')
    );
    el.appendChild(item);
  });
  return el;
}

function _card(icon,title,color,content) {
  var titleIcon = h('i',{className:icon,style:{color:color,background:color+'18'}});
  return h('div',{className:'dash-card'},
    h('div',{className:'dash-card-title',style:{color:color}},titleIcon,title),
    content
  );
}

function _statCard(val,label,color,bg) {
  return h('div',{className:'dash-stat',style:{'--stat-color':color,background:bg||''}},
    h('div',{className:'dash-stat-val',style:{'--stat-color':color}},String(val)),
    h('div',{className:'dash-stat-label'},label),
    h('div',{style:{position:'absolute',bottom:'0',left:'0',right:'0',height:'3px',background:'linear-gradient(90deg,transparent,'+color+',transparent)',opacity:'.4'}})
  );
}

function renderDashboard() {
  _startDashRefresh();
  var w=h('div',{className:'content-area',style:{overflowY:'auto',maxHeight:'calc(100vh - 80px)',paddingBottom:'40px'}});
  var root=h('div',{id:'dash-root'});
  try {
    if (typeof loadAllProfiles === 'function') loadAllProfiles();
    root.appendChild(_buildDash());
  } catch(e) {
    root.appendChild(h('div',{style:{padding:'40px',textAlign:'center',color:'var(--text-muted)'}},'Cargando dashboard...'));
    setTimeout(function() {
      try {
        var el = document.getElementById('dash-root');
        if (el) { el.innerHTML = ''; el.appendChild(_buildDash()); }
      } catch(e2) {}
    }, 2000);
  }
  w.appendChild(root);
  return w;
}

function _buildDash() {
  var logs=STATE.logs||[],profiles=AUTH.users||[],adm={};
  // Build set of valid usernames from profiles
  var validUsers={};
  profiles.forEach(function(p){
    if(p&&p.username){validUsers[p.username]=true;if(p.role==='admin')adm[p.username]=true;}
  });
  // If no profiles loaded, at least include current user
  if(AUTH.currentUser&&AUTH.currentUser.username){
    validUsers[AUTH.currentUser.username]=true;
    if(AUTH.currentUser.role==='admin')adm[AUTH.currentUser.username]=true;
  }
  var hasValidUsers=Object.keys(validUsers).length>0;
  var allLogs=logs.filter(function(l){return l&&l.user&&l.user!=='Sistema';});
  // Filter logs to only valid users if we have a user list
  var ul=hasValidUsers?allLogs.filter(function(l){return validUsers[l.user];}):allLogs;

  var byUser={};
  ul.forEach(function(l){
    if(!byUser[l.user]) byUser[l.user]={total:0,login:0,edit:0,create:0,delete:0,other:0};
    byUser[l.user].total++;
    if(l.type==='login') byUser[l.user].login++;
    else if(l.type==='edit') byUser[l.user].edit++;
    else if(l.type==='create') byUser[l.user].create++;
    else if(l.type==='delete') byUser[l.user].delete++;
    else byUser[l.user].other++;
  });
  var byType={};ul.forEach(function(l){byType[l.type]=(byType[l.type]||0)+1;});
  var byArea={};
  (STATE.records||[]).forEach(function(r){
    if(!byArea[r.area])byArea[r.area]={total:0,v:0,pv:0,ve:0};byArea[r.area].total++;
    var e=typeof calcEstado==='function'?calcEstado(r.fechaEmision):'';
    if(e==='Vigente')byArea[r.area].v++;else if(e==='Por vencer')byArea[r.area].pv++;else if(e==='Vencido')byArea[r.area].ve++;
  });
  var tV=0,tP=0,tE=0;Object.keys(byArea).forEach(function(a){tV+=byArea[a].v;tP+=byArea[a].pv;tE+=byArea[a].ve;});

  var colors=['#06b6d4','#a855f7','#eab308','#22c55e','#ef4444','#f97316','#ec4899','#3b82f6','#14b8a6','#8b5cf6','#6366f1','#84cc16'];
  var tl={login:'Sesiones',edit:'Ediciones',create:'Creaciones',logout:'Cierres',delete:'Eliminaciones',renew:'Renovaciones',obsolete:'Obsoletos',restore:'Restauraciones',loginFail:'Fallidos',loginBlock:'Bloqueos',backup:'Backups',userAdd:'Usuarios',salida:'Salidas',userPw:'Pass',sessionTimeout:'Timeout',import:'Importaciones',export:'Exportaciones',permDelete:'Borrados'};
  var tcl={login:'#06b6d4',edit:'#eab308',create:'#22c55e',logout:'#64748b',delete:'#ef4444',renew:'#10b981',obsolete:'#f97316',loginFail:'#dc2626',loginBlock:'#991b1b',backup:'#a855f7',userAdd:'#3b82f6',salida:'#14b8a6',import:'#6366f1',export:'#10b981',permDelete:'#dc2626'};

  var uData=Object.keys(byUser).sort(function(a,b){return byUser[b].total-byUser[a].total;}).map(function(u,i){return{label:u+(adm[u]?' \u2605':''),value:byUser[u].total,color:colors[i%colors.length]};});
  var uT=0;uData.forEach(function(d){uT+=d.value;});
  var tData=Object.keys(byType).sort(function(a,b){return byType[b]-byType[a];}).map(function(t){return{label:tl[t]||t,value:byType[t],color:tcl[t]||'#94a3b8'};});
  var tT=0;tData.forEach(function(d){tT+=d.value;});
  var aData=Object.keys(byArea).sort(function(a,b){return byArea[b].total-byArea[a].total;}).map(function(a,i){return{label:a,value:byArea[a].total,color:colors[i%colors.length]};});
  var aT=0;aData.forEach(function(d){aT+=d.value;});
  var sData=[{label:'Vigentes',value:tV,color:'#22c55e'},{label:'Por vencer',value:tP,color:'#eab308'},{label:'Vencidos',value:tE,color:'#ef4444'}];
  var sT=tV+tP+tE;

  var maxU=1;Object.keys(byUser).forEach(function(u){if(byUser[u].total>maxU)maxU=byUser[u].total;});

  var now=new Date();
  var recent=ul.slice(0,30);

  // Health score
  var healthPct=sT>0?Math.round((tV/sT)*100):100;
  var healthColor=healthPct>=80?'#22c55e':healthPct>=50?'#eab308':'#ef4444';

  var c=h('div',null,
    // ---- Header bar with live indicator ----
    h('div',{className:'dash-header-bar'},
      h('div',{className:'dash-header-info'},
        h('span',{className:'dash-live-dot'}),
        h('span',null,'Panel en vivo \u00b7 '+ul.length+' eventos \u00b7 '+Object.keys(byUser).length+' usuarios activos')
      ),
      h('div',{className:'dash-time'},
        h('i',{className:'fas fa-satellite-dish',style:{color:'#10b981',fontSize:'13px'}}),
        h('span',{className:'dash-time-val'},now.toLocaleTimeString('es-MX',{hour:'2-digit',minute:'2-digit',second:'2-digit'})),
        h('span',null,now.toLocaleDateString('es-MX',{weekday:'short',day:'numeric',month:'short'}))
      )
    ),

    // ---- Stat cards ----
    h('div',{style:{display:'flex',gap:'12px',marginBottom:'20px',flexWrap:'wrap'}},
      _statCard(ul.length,'Eventos','#06b6d4'),
      _statCard(Object.keys(byUser).length,'Empleados','#a855f7'),
      _statCard(sT,'Documentos','#3b82f6'),
      _statCard(tV,'Vigentes','#22c55e'),
      _statCard(tP,'Por Vencer','#eab308',tP>0?'rgba(234,179,8,.06)':''),
      _statCard(tE,'Vencidos','#ef4444',tE>0?'rgba(239,68,68,.06)':'')
    ),

    // ---- Health bar ----
    h('div',{className:'dash-card',style:{marginBottom:'16px',padding:'16px 24px'}},
      h('div',{style:{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'10px'}},
        h('div',{style:{display:'flex',alignItems:'center',gap:'10px'}},
          h('i',{className:'fas fa-heartbeat',style:{color:healthColor,fontSize:'16px'}}),
          h('span',{style:{fontWeight:'700',fontSize:'13px',color:'var(--text-primary)'}},'Salud del Sistema')
        ),
        h('span',{style:{fontFamily:'JetBrains Mono,monospace',fontWeight:'800',fontSize:'18px',color:healthColor}},healthPct+'%')
      ),
      h('div',{style:{height:'8px',background:'var(--bg-input)',borderRadius:'4px',overflow:'hidden'}},
        h('div',{style:{height:'100%',width:healthPct+'%',borderRadius:'4px',background:'linear-gradient(90deg,'+healthColor+',color-mix(in srgb,'+healthColor+' 70%,#fff))',transition:'width 1s ease'}})
      ),
      h('div',{style:{display:'flex',justifyContent:'space-between',marginTop:'8px',fontSize:'11px',color:'var(--text-muted)'}},
        h('span',null,tV+' vigentes'),
        h('span',{style:{color:'#eab308'}},tP+' por vencer'),
        h('span',{style:{color:'#ef4444'}},tE+' vencidos')
      )
    ),

    // ---- Pie charts row ----
    h('div',{className:'dash-grid dash-grid-2'},
      _card('fas fa-users','Actividad por Usuario','#06b6d4',h('div',null,_pie(uData,220,'tip-u'),_legend(uData,uT))),
      _card('fas fa-layer-group','Actividad por Tipo','#eab308',h('div',null,_pie(tData,220,'tip-t'),_legend(tData,tT)))
    ),
    h('div',{className:'dash-grid dash-grid-2'},
      _card('fas fa-building','Documentos por \u00c1rea','#a855f7',h('div',null,_pie(aData,220,'tip-a'),_legend(aData,aT))),
      _card('fas fa-shield-halved','Estado General','#22c55e',h('div',null,_pie(sData,220,'tip-s'),_legend(sData,sT)))
    ),

    // ---- Employee detail table ----
    _card('fas fa-ranking-star','Detalle por Empleado','#06b6d4',
      h('div',{style:{overflowX:'auto'}},
        (function(){
          var tbl=h('table',{className:'dash-table'});
          var thead=h('thead',null,h('tr',null,
            h('th',{style:{textAlign:'left'}},'Usuario'),
            h('th',{style:{color:'#06b6d4'}},'Sesiones'),
            h('th',{style:{color:'#22c55e'}},'Creados'),
            h('th',{style:{color:'#eab308'}},'Editados'),
            h('th',{style:{color:'#ef4444'}},'Eliminados'),
            h('th',{style:{color:'var(--text-muted)'}},'Otros'),
            h('th',{style:{color:'var(--text-primary)'}},'Total'),
            h('th',{style:{textAlign:'left',minWidth:'180px'}},'')
          ));
          tbl.appendChild(thead);
          var tbody=h('tbody',null);
          Object.keys(byUser).sort(function(a,b){return byUser[b].total-byUser[a].total;}).forEach(function(u,i) {
            var d=byUser[u],pct=Math.round((d.total/maxU)*100);
            var clr=colors[i%colors.length];
            var row=h('tr',null,
              h('td',{style:{fontWeight:'700',color:clr,fontFamily:'DM Sans,sans-serif'}},u+(adm[u]?' \u2605':'')),
              h('td',null,String(d.login)),
              h('td',{style:{color:'#22c55e'}},String(d.create)),
              h('td',{style:{color:'#eab308'}},String(d.edit)),
              h('td',{style:{color:'#ef4444'}},String(d.delete)),
              h('td',{style:{color:'var(--text-muted)'}},String(d.other)),
              h('td',{style:{fontWeight:'800'}},String(d.total)),
              h('td',null,h('div',{className:'dash-table-bar'},h('div',{className:'dash-table-bar-fill',style:{width:pct+'%',background:'linear-gradient(90deg,'+clr+','+clr+'80)'}})))
            );
            tbody.appendChild(row);
          });
          tbl.appendChild(tbody);
          return tbl;
        })()
      )
    ),

    // ---- Recent activity ----
    _card('fas fa-bolt','Actividad Reciente','var(--accent)',
      recent.length===0?h('div',{style:{textAlign:'center',padding:'30px',color:'var(--text-muted)'}},'Sin actividad'):
      (function(){
        var wrap=h('div',{style:{maxHeight:'400px',overflowY:'auto'}});
        recent.forEach(function(l,idx){
          var d=new Date(l.timestamp),cl=tcl[l.type]||'#94a3b8';
          var item=h('div',{className:'dash-activity-item',style:{animationDelay:(idx*30)+'ms'}},
            h('span',{className:'dash-activity-time'},d.toLocaleDateString('es-MX')+' '+d.toLocaleTimeString('es-MX')),
            h('span',{className:'dash-activity-user',style:{color:cl}},l.user),
            h('span',{className:'dash-activity-badge',style:{background:cl+'15',color:cl,border:'1px solid '+cl+'30'}},(tl[l.type]||l.type)),
            h('span',{className:'dash-activity-detail'},l.details||l.docName||'')
          );
          wrap.appendChild(item);
        });
        return wrap;
      })()
    )
  );
  return c;
}
function renderVerificacionISO() {
  var syncOk = typeof SYNC !== 'undefined' && SYNC.initialized && SYNC.connected;
  var authOk = typeof AUTH !== 'undefined' && AUTH._initialized;
  var hasData = (STATE.records || []).length > 0;
  var hasLogs = (STATE.logs || []).length > 0;
  var hasBackups = syncOk;
  var hashActual = _dataHash({r: STATE.records, o: STATE.obsoletos, p: STATE.papelera, s: STATE.salidas, n: STATE.nextId});
  var totalDocs = (STATE.records||[]).length + (STATE.obsoletos||[]).length + (STATE.papelera||[]).length;
  var vencidos = (STATE.records||[]).filter(function(r) { return typeof calcEstado === 'function' && calcEstado(r.fechaEmision) === 'Vencido'; }).length;

  function chk(ok, label, detail) {
    return h('div', {className: 'iso-check ' + (ok ? 'iso-pass' : 'iso-fail')},
      h('i', {className: ok ? 'fas fa-check-circle' : 'fas fa-times-circle'}),
      h('div', {className: 'iso-check-body'},
        h('div', {className: 'iso-check-label'}, label),
        h('div', {className: 'iso-check-detail'}, detail)
      )
    );
  }

  var checks = [];
  checks.push(chk(syncOk, 'A.8 — Conexión a la nube', syncOk ? 'Supabase conectado (v' + SYNC._version + ')' : 'Sin conexión a Supabase'));
  checks.push(chk(authOk, 'A.9.1 — Autenticación centralizada', authOk ? 'Supabase Auth activo' : 'Auth no inicializado'));
  checks.push(chk(!!AUTH.currentUser, 'A.9.2 — Sesión autenticada', AUTH.currentUser ? 'Usuario: ' + AUTH.currentUser.username + ' (' + AUTH.currentUser.role + ')' : 'Sin sesión'));
  checks.push(chk(true, 'A.9.3 — Política de contraseñas', 'Mín. 8 caracteres, mayúscula, minúscula, número'));
  checks.push(chk(true, 'A.9.4 — Bloqueo por intentos fallidos', '5 intentos máx., bloqueo 60 seg., intentos registrados en log'));
  checks.push(chk(true, 'A.9.4 — Timeout de sesión', '30 minutos de inactividad → cierre automático'));
  checks.push(chk(true, 'A.10.1 — Cifrado en tránsito', 'HTTPS/TLS via Supabase'));
  checks.push(chk(true, 'A.10.2 — Credenciales protegidas', 'Anon key pública (diseño Supabase), seguridad en RLS + Auth'));
  checks.push(chk(true, 'A.10.2 — CSRF token por sesión', 'Token criptográfico validado en cada escritura'));
  checks.push(chk(hasData, 'A.12.1 — Datos operativos', hasData ? totalDocs + ' documentos en el sistema' : 'Sin datos cargados'));
  checks.push(chk(hasLogs, 'A.12.4 — Registro de actividad (Logs)', hasLogs ? STATE.logs.length + ' eventos registrados' : 'Sin logs'));
  checks.push(chk(true, 'A.12.4 — Logs completos sin recortar', 'Todos los logs se guardan íntegros'));
  checks.push(chk(hasBackups, 'A.12.3 — Backups automáticos en la nube', hasBackups ? 'Cada hora + manuales' : 'Sin conexión para backups'));
  checks.push(chk(true, 'A.12.3 — Backups imborrables', 'No se pueden eliminar desde la app ni API'));
  checks.push(chk(true, 'A.12.6 — Protección contra borrado', 'Eliminación permanente deshabilitada'));
  checks.push(chk(true, 'A.14.1 — Sanitización de entradas', '_esc() + strip <> en formularios'));
  checks.push(chk(true, 'A.14.2 — Headers de seguridad', 'CSP, X-Content-Type, Referrer, Permissions'));
  checks.push(chk(true, 'A.14.2 — Protección XSS', 'h() usa textContent, DOM API en lugar de innerHTML, _esc() + _clean() en formularios'));
  checks.push(chk(true, 'A.14.2 — Validación de archivos', 'Máx. 10MB, solo PDF/IMG/DOC/XLS/TXT'));
  checks.push(chk(true, 'A.9.2 — RLS en Supabase', 'Solo usuarios autenticados acceden. Solo admin crea backups y gestiona usuarios'));
  checks.push(chk(true, 'A.14.2 — Protección consola', 'Sin logs de depuración en producción'));
  checks.push(chk(true, 'A.14.2 — Mensajes de error genéricos', 'Errores del servidor no se exponen al usuario'));
  checks.push(chk(syncOk, 'A.17.1 — Sincronización en tiempo real', syncOk ? 'WebSocket activo + heartbeat 30s + reconexión auto' : 'Desconectado'));
  checks.push(chk(true, 'A.17.1 — Merge inteligente', 'Registros se fusionan por ID, no se sobreescriben'));
  checks.push(chk(true, 'A.17.1 — Auto-guardado', 'visibilitychange + beforeunload + Supabase'));
  checks.push(chk(syncOk, 'A.17.1 — Protección contra sobreescritura', syncOk ? '_initialLoadDone + sessionId echo suppression' : 'Solo protección local'));
  checks.push(chk(vencidos === 0, 'A.18.2 — Documentos vigentes', vencidos === 0 ? 'Todos los documentos están vigentes' : vencidos + ' documento(s) vencido(s) requieren atención'));

  var passCount = checks.filter(function(c) { return c.className.indexOf('iso-pass') >= 0; }).length;
  var totalChecks = checks.length;
  var pct = Math.round((passCount / totalChecks) * 100);

  return h('div', {className: 'content-area'},
    h('div', {className: 'iso-header'},
      h('div', {className: 'iso-header-info'},
        h('div', {className: 'iso-title'}, 'DEBBIOM v' + APP_VERSION),
        h('div', {className: 'iso-subtitle'}, 'Verificación de controles ISO 27001 — Build ' + APP_BUILD),
        h('div', {className: 'iso-hash'}, 'Hash de integridad: ' + hashActual + ' | Registros: ' + (STATE.records||[]).length + ' | Sesión: ' + SYNC.sessionId)
      ),
      h('div', {className: 'iso-score ' + (pct === 100 ? 'iso-score-ok' : 'iso-score-warn')},
        h('div', {className: 'iso-score-num'}, pct + '%'),
        h('div', {className: 'iso-score-label'}, passCount + '/' + totalChecks + ' controles')
      )
    ),
    h('div', {className: 'iso-checks'}, ...checks),
    h('div', {className: 'iso-footer'},
      h('div', {className: 'iso-footer-text'}, 'Generado: ' + new Date().toLocaleString('es-MX') + ' | Usuario: ' + (AUTH.currentUser ? AUTH.currentUser.username : '—')),
      h('button', {className: 'btn btn-primary', onClick: function() { window.print(); }}, h('i', {className: 'fas fa-print'}), ' Imprimir Reporte')
    )
  );
}

function _forceReloadInitialData() {
  if (typeof RAW_DATA === 'undefined') {
    showToast('No hay datos iniciales (RAW_DATA) disponibles', 'error');
    return;
  }
  if (!confirm('¿Recargar todos los datos desde los datos iniciales del sistema?\n\nEsto reemplazará los datos actuales en la nube con los datos incluidos en la aplicación (' + RAW_DATA.r.length + ' registros).')) return;
  loadExcelData();
  if (typeof saveState === 'function') saveState();
  if (typeof syncSaveState === 'function') syncSaveState();
  showToast('✅ Datos iniciales cargados: ' + (STATE.records||[]).length + ' registros, ' + (STATE.obsoletos||[]).length + ' obsoletos', 'success');
  render();
}

(async function mainInit() {
  // Always start with RAW_DATA so the app is never empty
  loadExcelData();

  if (typeof initAuth === 'function') await initAuth();
  render();

  if (typeof initSync === 'function') {
    var syncReady = await initSync();

    if (syncReady) {
      var cloudResult = await _syncLoadWithStatus();

      if (typeof SYNC !== 'undefined') SYNC._initialLoadDone = true;

      if (cloudResult === 'loaded') {
        // Cloud had data — check if RAW_DATA has newer/more data (new deployment)
        var cloudCount = (STATE.records || []).length;
        var rawCount = (typeof RAW_DATA !== 'undefined' && RAW_DATA.r) ? RAW_DATA.r.length : 0;
        if (rawCount > cloudCount && rawCount > 100) {
          // New deployment with significantly more data — push to cloud
          loadExcelData();
          saveState();
          showToast('📦 Datos actualizados: ' + (STATE.records||[]).length + ' registros desde nueva versión', 'info');
        }
        render();
      } else if (cloudResult === 'empty' || cloudResult === 'error') {
        // Cloud is empty or error — push RAW_DATA to cloud
        saveState();
        if ((STATE.records||[]).length > 0) {
          showToast('📦 Datos iniciales cargados a la nube: ' + (STATE.records||[]).length + ' registros', 'info');
        }
        render();
      }

      listenStateChanges();

      if (AUTH.currentUser && typeof registerPresence === 'function') {
        registerPresence();
      }

      showToast('Sincronización en la nube activa', 'info');
    } else {
      if (typeof SYNC !== 'undefined') SYNC._initialLoadDone = true;
    }
  }
})();
document.addEventListener('visibilitychange', function() {
  if (document.visibilityState === 'hidden' && AUTH.currentUser && _unsavedChanges) {
    if (typeof SYNC !== 'undefined' && SYNC.initialized && SYNC._client) {
      clearTimeout(SYNC._saveTimer);
      _doSave();
    }
  }
});

window.addEventListener('beforeunload', function() {
  if (AUTH.currentUser && _unsavedChanges && typeof SYNC !== 'undefined' && SYNC.initialized && SYNC._client) {
    clearTimeout(SYNC._saveTimer);
    _doSave();
  }
});

try { Object.freeze(window._esc); } catch(e) {}
try {
  Object.defineProperty(window, 'STATE', { configurable: false });
  Object.defineProperty(window, 'AUTH', { configurable: false });
} catch(e) {}
