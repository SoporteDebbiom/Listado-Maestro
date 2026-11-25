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
    var rawTs = l.timestamp || l.date || l.ts;
    if (rawTs) { try { var d = new Date(rawTs); if (!isNaN(d.getTime())) { fecha = d.toLocaleDateString('es-MX'); hora = d.toLocaleTimeString('es-MX'); } } catch(e) { fecha = rawTs; } }
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
  currentSection: 'dashboard',
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
  if (!iso) return '—';
  try {
    const d = new Date(iso);
    if (isNaN(d.getTime())) return String(iso);
    const dd = String(d.getDate()).padStart(2,'0');
    const mm = String(d.getMonth()+1).padStart(2,'0');
    const yy = String(d.getFullYear());
    const hh = String(d.getHours()).padStart(2,'0');
    const mi = String(d.getMinutes()).padStart(2,'0');
    const ss = String(d.getSeconds()).padStart(2,'0');
    return dd+'/'+mm+'/'+yy+' '+hh+':'+mi+':'+ss;
  } catch(e) { return String(iso); }
}
function _formatBackupDate(iso) {
  if (!iso) return '—';
  try {
    var d = new Date(iso);
    if (isNaN(d.getTime())) return String(iso);
    var dd = String(d.getDate()).padStart(2, '0');
    var mm = String(d.getMonth() + 1).padStart(2, '0');
    var yy = d.getFullYear();
    var hh = String(d.getHours()).padStart(2, '0');
    var mi = String(d.getMinutes()).padStart(2, '0');
    return dd + '/' + mm + '/' + yy + ' ' + hh + ':' + mi;
  } catch (e) { return String(iso); }
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
      h('img', {className:'sidebar-logo', src: DEBBIOM_LOGO || 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect fill="#003366" width="100" height="100" rx="12"/><text x="50" y="58" font-family="Arial" font-size="22" font-weight="bold" fill="white" text-anchor="middle">DEBBIOM</text></svg>'), alt:'DEBBIOM'}),
      h('div', {className:'sidebar-header-text'},
        h('h1', null, 'DEBBIOM'),
        h('p', null, 'LISTADO DE ESTUDIOS')
      )
    ),
    h('nav', {className: 'sidebar-nav'},
      h('div', {className: 'nav-section'},
        h('div', {className: 'nav-section-title'}, 'DOCUMENTOS'),
        ...areaItems
      ),
      h('div', {className: 'nav-section'},
        h('div', {className: 'nav-section-title'}, 'GESTIÓN'),
        h('div', {
          className: 'nav-item' + (STATE.currentSection === 'dashboard' ? ' active' : ''),
          onClick: () => { STATE.currentSection='dashboard'; render(); }
        },
          h('i', {className: 'fas fa-chart-pie'}),
          h('span', null, 'Dashboard')
        ),
        h('div', {
          className: 'nav-item nav-item-alertas' + (STATE.currentSection === 'alertas' ? ' active' : ''),
          onClick: () => { STATE.currentSection='alertas'; render(); }
        },
          h('i', {className: 'fas fa-bell'}),
          h('span', null, 'Alertas de Vigencia'),
          (() => {
            var alCount = typeof getAlertasCriticas === 'function' ? getAlertasCriticas().length : 0;
            return alCount > 0 ? h('span', {className: 'nav-badge nav-badge-danger'}, String(alCount)) : null;
          })()
        ),
        h('div', {
          className: 'nav-item' + (STATE.currentSection === 'obsoletos' ? ' active' : ''),
          onClick: () => { STATE.currentSection='obsoletos'; STATE.page=1; render(); }
        },
          h('i', {className: 'fas fa-box-archive'}),
          h('span', null, 'Obsoletos'),
          h('span', {className: 'nav-badge'}, String(obsCount))
        ),
        h('div', {
          className: 'nav-item' + (STATE.currentSection === 'salidas' ? ' active' : ''),
          onClick: () => { STATE.currentSection='salidas'; STATE.page=1; render(); }
        },
          h('i', {className: 'fas fa-sign-out-alt'}),
          h('span', null, 'Salidas'),
          h('span', {className: 'nav-badge'}, String(salCount))
        ),
        h('div', {
          className: 'nav-item' + (STATE.currentSection === 'papelera' ? ' active' : ''),
          onClick: () => { STATE.currentSection='papelera'; STATE.page=1; render(); }
        },
          h('i', {className: 'fas fa-trash-can'}),
          h('span', null, 'Papelera'),
          papCount > 0 ? h('span', {className: 'nav-badge nav-badge-warning'}, String(papCount)) : null
        ),
        h('div', {
          className: 'nav-item' + (STATE.currentSection === 'logs' ? ' active' : ''),
          onClick: () => { STATE.currentSection='logs'; STATE.page=1; render(); }
        },
          h('i', {className: 'fas fa-list-check'}),
          h('span', null, 'Bitácora')
        ),
        h('div', {
          className: 'nav-item nav-item-backups' + (STATE.currentSection === 'backups' ? ' active' : ''),
          onClick: () => { STATE.currentSection='backups'; STATE.page=1; render(); }
        },
          h('i', {className: 'fas fa-cloud'}),
          h('span', null, 'Respaldos')
        )
      )
    ),
    h('div', {className: 'sidebar-user'},
      h('div', {style:{display:'flex',alignItems:'center',gap:'10px',width:'100%',marginBottom:'4px'}},
        h('i', {className: 'fas fa-user-circle', style:{fontSize:'24px',color:'var(--accent)',flexShrink:0}}),
        h('div', {style:{flex:1,minWidth:0}},
          h('div', {style:{fontWeight:'700', fontSize:'13px',whiteSpace:'nowrap',overflow:'hidden',textOverflow:'ellipsis'}}, AUTH.currentUser ? (AUTH.currentUser.fullName || AUTH.currentUser.username) : ''),
          h('div', {style:{fontSize:'11px', color:'var(--text-muted)'}}, isAdmin() ? 'Administrador' : 'Usuario')
        )
      ),
      h('div', {style:{display:'flex',gap:'8px',width:'100%'}},
        isAdmin() ? h('button', {
          className: 'btn btn-sm',
          style:{fontSize:'11px',padding:'6px 12px',flex:1},
          onClick: () => { STATE.modal = 'users'; refreshModal(); }
        }, h('i',{className:'fas fa-users-gear'}), ' Usuarios') : null,
        h('button', {className: 'logout-btn', style:{fontSize:'11px',padding:'6px 12px',flex:1,display:'flex',alignItems:'center',justifyContent:'center',gap:'6px',borderRadius:'8px',border:'1px solid var(--border)',background:'none',color:'var(--text-muted)',cursor:'pointer'}, onClick: () => { if(typeof logout === 'function') logout(); }},
          h('i', {className: 'fas fa-right-from-bracket'}),
          'Cerrar sesión'
        )
      )
    ),
    h('div', {className: 'sidebar-sync'},
      typeof SYNC !== 'undefined' && SYNC.connected
        ? h('div', {className: 'sidebar-online-info'}, h('span', null, '●'), ' En línea')
        : h('div', {className: 'sidebar-online-info', style:{color:'var(--text-muted)'}}, '○ Sin conexión')
    )
  );
}

function _updateObsoletoBtn() {
  var btn = document.getElementById('btn-save-obsoleto');
  if (btn) btn.disabled = !(STATE.obsoletoMotivo && STATE.obsoletoMotivo.trim());
}

function _updateRenewBtn() {
  var btn = document.getElementById('btn-confirm-renew');
  if (btn) btn.disabled = !STATE._renewDate || !(STATE._renewMotivo && STATE._renewMotivo.trim());
}

function closeModal() {
  STATE.modal = null;
  STATE.selectedRecord = null;
  STATE.editForm = {};
  STATE.salidaForm = {};
  STATE.obsoletoMotivo = '';
  STATE._renewDate = '';
  STATE._renewMotivo = '';
  STATE.originalVersion = null;
  refreshModal();
}

async function _uploadToStorage(file) {
  try {
    var reader = new FileReader();
    return new Promise(function(resolve) {
      reader.onload = function(e) {
        resolve({ name: file.name, size: file.size, type: file.type, data: e.target.result });
      };
      reader.onerror = function() { showToast('Error al leer archivo', 'error'); resolve(null); };
      reader.readAsDataURL(file);
    });
  } catch(err) {
    showToast('Error al procesar archivo', 'error');
    return null;
  }
}

function fg(label, field, opts) {
  opts = opts || {};
  var form = STATE.editForm;
  var cls = 'form-group' + (opts.full ? ' full' : '');

  if (opts.options) {
    return h('div', {className: cls},
      h('label', {className: 'form-label'}, label),
      h('select', {
        className: 'form-input',
        value: form[field] || '',
        onChange: function(e) { STATE.editForm[field] = e.target.value; }
      },
        h('option', {value: ''}, '— Seleccionar —'),
        ...opts.options.map(function(o) { return h('option', {value: o}, o); })
      )
    );
  }

  if (opts.textarea) {
    return h('div', {className: cls},
      h('label', {className: 'form-label'}, label),
      h('textarea', {
        className: 'form-textarea',
        value: form[field] || '',
        placeholder: opts.placeholder || '',
        onInput: function(e) { STATE.editForm[field] = e.target.value; }
      })
    );
  }

  if (opts.elaboroList) {
    return h('div', {className: cls},
      h('label', {className: 'form-label'}, label),
      h('input', {
        className: 'form-input',
        list: 'elaboro-list',
        value: form[field] || '',
        placeholder: opts.placeholder || '',
        onInput: function(e) { STATE.editForm[field] = e.target.value; }
      }),
      h('datalist', {id: 'elaboro-list'},
        ...STATE.elaboros.map(function(el) { return h('option', {value: el}); })
      )
    );
  }

  return h('div', {className: cls},
    h('label', {className: 'form-label'}, label),
    h('input', {
      className: 'form-input',
      type: opts.type || 'text',
      value: form[field] || '',
      placeholder: opts.placeholder || '',
      onInput: function(e) {
        var val = e.target.value;
        if (opts.numbersOnly) val = val.replace(/[^0-9]/g, '');
        STATE.editForm[field] = val;
      }
    }),
    opts.hint ? h('div', {className: 'form-hint'}, opts.hint) : null
  );
}

function renderViewModal() {
  var rec = STATE.selectedRecord;
  if (!rec) return null;
  var estado = typeof calcEstado === 'function' ? calcEstado(rec.fechaEmision) : '';
  var diasV = typeof calcDiasVigencia === 'function' ? calcDiasVigencia(rec.fechaEmision) : null;
  var vigF = typeof renderVigenciaDate === 'function' ? renderVigenciaDate(rec.fechaEmision) : '';

  var estadoClass = estado === 'Vigente' ? 'badge-green' : estado === 'Por vencer' ? 'badge-orange' : estado === 'Vencido' ? 'badge-red' : '';
  var code = rec.codManual || rec.codInstructivo || rec.codListado || rec.codFormato || '';

  return h('div', {className:'modal-overlay', onClick:function(e){if(e.target.className==='modal-overlay')closeModal();}},
    h('div', {className:'modal'},
      h('div', {className:'modal-header'},
        h('h2', null, h('i',{className:'fas fa-file-lines', style:{marginRight:'8px',color:'var(--accent)'}}), 'Detalle del Documento'),
        h('button', {className:'modal-close', onClick:closeModal}, h('i',{className:'fas fa-times'}))
      ),
      h('div', {className:'modal-body'},
        h('div', {style:{background:'var(--bg-card)',borderRadius:'10px',padding:'16px',marginBottom:'16px',border:'1px solid var(--border)'}},
          h('div', {style:{display:'flex',justifyContent:'space-between',alignItems:'flex-start'}},
            h('div', null,
              h('div', {style:{fontWeight:'700',fontSize:'16px',color:'var(--text-primary)'}}, rec.nombreDoc || 'Sin nombre'),
              code ? h('div', {style:{fontSize:'13px',color:'var(--accent)',marginTop:'4px',fontFamily:'JetBrains Mono, monospace'}}, code) : null
            ),
            estado ? h('span', {className:'badge ' + estadoClass}, estado) : null
          )
        ),
        h('div', {style:{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'12px',fontSize:'13px'}},
          _viewField('Área', rec.area),
          _viewField('No.', rec.no),
          _viewField('Cód. Manual', rec.codManual),
          _viewField('Cód. Instructivo', rec.codInstructivo),
          _viewField('Cód. Listado', rec.codListado),
          _viewField('Cód. Formato', rec.codFormato),
          _viewField('Versión', rec.version),
          _viewField('Fecha Emisión', rec.fechaEmision),
          _viewField('Vigencia', vigF),
          _viewField('Días restantes', diasV !== null ? String(diasV) : ''),
          _viewField('Elaboró', rec.elaboro),
          _viewField('Tipo Resguardo', rec.tipoResguardo),
          _viewField('Ubicación', rec.ubicacion),
          _viewField('Copias', rec.copias),
          _viewField('Tipo Resguardo Copia', rec.tipoResguardoCopia),
          _viewField('Usuarios', rec.usuarios)
        ),
        rec.observaciones ? h('div', {style:{marginTop:'12px'}}, _viewField('Observaciones', rec.observaciones)) : null,
        (rec.archivos && rec.archivos.length > 0) ? h('div', {style:{marginTop:'12px'}},
          h('div', {style:{fontWeight:'600',fontSize:'12px',color:'var(--text-muted)',marginBottom:'6px'}}, 'Archivos'),
          h('div', {style:{display:'flex',gap:'6px',flexWrap:'wrap'}},
            ...rec.archivos.map(function(f) { return h('span', {className:'file-chip'}, h('i',{className:'fas fa-file'}), f.name); })
          )
        ) : null,
        (rec.archivoURLs && rec.archivoURLs.length > 0) ? h('div', {style:{marginTop:'8px'}},
          h('div', {style:{fontWeight:'600',fontSize:'12px',color:'var(--text-muted)',marginBottom:'6px'}}, 'URLs de archivos'),
          ...rec.archivoURLs.filter(function(u){return u;}).map(function(url) {
            return h('a', {href: _safeHref(url), target:'_blank', rel:'noopener', style:{display:'block',fontSize:'12px',color:'var(--accent)',marginBottom:'4px'}}, url);
          })
        ) : null
      ),
      h('div', {className:'modal-footer'},
        h('button', {className:'btn btn-secondary', onClick:closeModal}, 'Cerrar'),
        isAdmin() ? h('button', {className:'btn btn-primary', onClick:function(){
          STATE.modal = 'edit';
          STATE.editForm = Object.assign({}, rec);
          STATE.originalVersion = rec.version || '';
          refreshModal();
        }}, h('i',{className:'fas fa-pen'}), 'Editar') : null
      )
    )
  );
}

function _viewField(label, value) {
  return h('div', {style:{marginBottom:'8px'}},
    h('div', {style:{fontWeight:'600',fontSize:'11px',color:'var(--text-muted)',textTransform:'uppercase',letterSpacing:'.5px'}}, label),
    h('div', {style:{color:'var(--text-primary)',marginTop:'2px'}}, value || '—')
  );
}

function renderEditModal() {
  var rec = STATE.selectedRecord;
  var form = STATE.editForm;
  var isAdd = STATE.modal === 'add';

  return h('div', {className:'modal-overlay', onClick:function(e){if(e.target.className==='modal-overlay')closeModal();}},
    h('div', {className:'modal', style:{maxWidth:'900px'}},
      h('div', {className:'modal-header'},
        h('h2', null, h('i',{className: isAdd ? 'fas fa-plus-circle' : 'fas fa-pen', style:{marginRight:'8px',color:'var(--accent)'}}), isAdd ? 'Nuevo Registro' : 'Editar Registro'),
        h('button', {className:'modal-close', onClick:closeModal}, h('i',{className:'fas fa-times'}))
      ),
      h('div', {className:'modal-body'},
        !isAdd && STATE.originalVersion ? h('div', {style:{background:'var(--bg-input)',borderRadius:'8px',padding:'10px 14px',marginBottom:'16px',fontSize:'12px',color:'var(--text-muted)',border:'1px solid var(--border)'}},
          h('i', {className:'fas fa-code-branch', style:{marginRight:'6px'}}),
          'Versión original: ' + STATE.originalVersion + ' — Si cambias la versión, la anterior se enviará a obsoletos.'
        ) : null,
        h('div', {className:'form-grid'},
          fg('Área *', 'area', {options: AREAS}),
          fg('No.', 'no'),
          fg('Cód. Manual/Procedimiento', 'codManual'),
          fg('Cód. Instructivo', 'codInstructivo'),
          fg('Cód. Listado', 'codListado'),
          fg('Cód. Formato', 'codFormato'),
          fg('Nombre del Documento *', 'nombreDoc', {full: true}),
          h('div', {className:'form-group'},
            h('label', {className:'form-label'}, 'Versión'),
            h('input', {
              className:'form-input',
              type:'text',
              value: form.version || '',
              onInput: function(e) {
                STATE.editForm.version = e.target.value;
                if (!isAdd && STATE.originalVersion) refreshModal();
              }
            })
          ),
          fg('Fecha de Emisión', 'fechaEmision', {type: 'date'}),
          fg('Elaboró', 'elaboro', {elaboroList: true}),
          fg('Tipo de Resguardo', 'tipoResguardo', {options: ['Físico','Electrónico']}),
          fg('Ubicación del Original', 'ubicacion'),

          !isAdd && STATE.originalVersion && String(form.version||'').trim() !== '' && String(form.version||'').trim() !== STATE.originalVersion
            ? h('div', {className:'form-group full'},
                h('label', {className:'form-label', style:{color:'var(--orange)',fontWeight:'600'}}, '⚠ Motivo del cambio de versión *'),
                h('textarea', {
                  className:'form-textarea',
                  placeholder:'Describe el motivo del cambio de versión...',
                  style:{minHeight:'60px'},
                  value: STATE.obsoletoMotivo || '',
                  onInput: function(e) { STATE.obsoletoMotivo = e.target.value; _updateObsoletoBtn(); }
                }),
                h('div', {className:'form-hint'}, 'Obligatorio cuando se cambia la versión')
              )
            : null,

          h('div', {className:'form-group full'},
            h('label', {className:'form-label'}, 'Archivos'),
            h('div', {style:{display:'flex',gap:'8px',flexWrap:'wrap',marginBottom:'8px'}},
              ...(form.archivos||[]).map(function(f,i) { return h('span', {className:'file-chip'},
                h('i',{className:'fas fa-file'}),
                f.name,
                h('i', {className:'fas fa-times', style:{marginLeft:'4px',cursor:'pointer'}, onClick:function(){
                  STATE.editForm.archivos.splice(i,1);
                  refreshModal();
                }})
              ); })
            ),
            h('input', {
              type:'file',
              multiple:true,
              className:'form-input',
              onChange: function(e) {
                var MAX_SIZE = 10 * 1024 * 1024;
                var ALLOWED = ['application/pdf','image/jpeg','image/png','image/gif','application/msword','application/vnd.openxmlformats-officedocument.wordprocessingml.document','application/vnd.ms-excel','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','text/plain'];
                var files = Array.from(e.target.files);
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
            onInput: function(e) { STATE._renewDate = e.target.value; _updateRenewBtn(); }
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
            onInput: function(e) { STATE._renewMotivo = e.target.value; _updateRenewBtn(); }
          }),
          h('div', {className:'form-hint'}, 'Este campo es obligatorio para el registro de auditoría')
        )
      ),
      h('div', {className:'modal-footer'},
        h('button', {className:'btn btn-secondary', onClick:closeModal}, 'Cancelar'),
        h('button', {
          className:'btn btn-primary',
          id:'btn-confirm-renew',
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

var _alertBannerShown = false;
function showAlertBanner() {
  // Solo mostrar una vez por sesión
  if (_alertBannerShown) return;
  if (sessionStorage.getItem('debbiom_alert_dismissed')) return;

  var criticas = typeof getAlertasCriticas === 'function' ? getAlertasCriticas() : [];
  if (criticas.length === 0) return;

  // Remover banner anterior si existe
  var existing = document.querySelector('.alert-banner');
  if (existing) existing.remove();

  _alertBannerShown = true;

  var vencidos = criticas.filter(function(a) { return a.nivel === 'vencido'; }).length;
  var porVencer = criticas.filter(function(a) { return a.nivel === 'critico'; }).length;

  var mensaje = '';
  if (vencidos > 0 && porVencer > 0) {
    mensaje = vencidos + ' documento(s) vencidos y ' + porVencer + ' por vencer en ≤15 días';
  } else if (vencidos > 0) {
    mensaje = vencidos + ' documento(s) con vigencia vencida';
  } else {
    mensaje = porVencer + ' documento(s) vencen en los próximos 15 días';
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
  btnVer.onclick = function(){STATE.currentSection='alertas';render();banner.remove();sessionStorage.setItem('debbiom_alert_dismissed','1');};
  var btnClose = document.createElement('button');
  btnClose.className = 'alert-banner-close';
  var icoClose = document.createElement('i');
  icoClose.className = 'fas fa-times';
  btnClose.appendChild(icoClose);
  btnClose.onclick = function(){banner.remove();sessionStorage.setItem('debbiom_alert_dismissed','1');};
  bannerContent.appendChild(ico);
  bannerContent.appendChild(sp);
  bannerContent.appendChild(btnVer);
  bannerContent.appendChild(btnClose);
  banner.appendChild(bannerContent);

  document.body.appendChild(banner);

  // Auto-cerrar después de 8 segundos
  setTimeout(function() {
    if (banner.parentElement) {
      banner.style.transition = 'opacity .5s';
      banner.style.opacity = '0';
      setTimeout(function() { if (banner.parentElement) banner.remove(); }, 500);
    }
  }, 8000);
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
  var btnExcelH = document.createElement('button');
  btnExcelH.className = 'btn btn-success';
  btnExcelH.style.cssText = 'background:linear-gradient(135deg,#2ed573,#1fab56);border:none;color:#fff;';
  var btnExcelI = document.createElement('i');
  btnExcelI.className = 'fas fa-file-excel';
  btnExcelH.appendChild(btnExcelI);
  btnExcelH.appendChild(document.createTextNode(' Exportar Backup Excel'));
  btnExcelH.addEventListener('click', function() { exportBackupToExcel(); });
  headerActions.appendChild(btnExcelH);
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

  var maxSearch = Math.min(10, rows.length);
  for (var r = 0; r < maxSearch; r++) {
    if (!rows[r]) continue;
    var joined = rows[r].map(function(c) { return (c || '').toString().toLowerCase(); }).join(' ');
    if (joined.indexOf('nombre del documento') >= 0 || joined.indexOf('nombre documento') >= 0) {
      return r;
    }
  }

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
  return 0;
}

function _parseWorkbook(wb) {
  var result = { records: [], obsoletos: [], elaboros: [] };
  var sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return result;
  var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (rows.length < 2) return result;

  var headerIdx = _findHeaderRow(rows);
  var headers = rows[headerIdx];
  if (!headers) return result;

  var colMap = {};
  var fieldMap = {
    'nombre del documento': 'nombreDoc', 'nombre documento': 'nombreDoc', 'documento': 'nombreDoc',
    'área': 'area', 'area': 'area',
    'no': 'no', 'no.': 'no', 'número': 'no',
    'cód. manual': 'codManual', 'cod. manual': 'codManual', 'manual': 'codManual', 'procedimiento': 'codManual',
    'cód. instructivo': 'codInstructivo', 'cod. instructivo': 'codInstructivo', 'instructivo': 'codInstructivo',
    'cód. listado': 'codListado', 'cod. listado': 'codListado', 'listado': 'codListado',
    'cód. formato': 'codFormato', 'cod. formato': 'codFormato', 'formato': 'codFormato',
    'versión': 'version', 'version': 'version',
    'fecha de emisión': 'fechaEmision', 'fecha emisión': 'fechaEmision', 'fecha emision': 'fechaEmision', 'emisión': 'fechaEmision',
    'elaboró': 'elaboro', 'elaboro': 'elaboro',
    'tipo de resguardo': 'tipoResguardo', 'tipo resguardo': 'tipoResguardo', 'resguardo': 'tipoResguardo',
    'ubicación': 'ubicacion', 'ubicacion': 'ubicacion',
    'copias': 'copias', 'copias controladas': 'copias',
    'observaciones': 'observaciones',
    'usuarios': 'usuarios'
  };

  for (var ci = 0; ci < headers.length; ci++) {
    var hd = (headers[ci] || '').toString().toLowerCase().trim();
    if (fieldMap[hd]) colMap[ci] = fieldMap[hd];
  }

  var elaboroSet = {};
  for (var ri = headerIdx + 1; ri < rows.length; ri++) {
    var row = rows[ri];
    if (!row || row.length === 0) continue;
    var rec = { id: 0, archivos: [], archivoURLs: [] };
    var hasData = false;
    for (var c in colMap) {
      var val = row[c];
      if (val instanceof Date) {
        val = val.toISOString().split('T')[0];
      }
      val = (val || '').toString().trim();
      if (val) hasData = true;
      rec[colMap[c]] = val;
    }
    if (hasData) {
      result.records.push(rec);
      if (rec.elaboro) elaboroSet[rec.elaboro] = true;
    }
  }
  result.elaboros = Object.keys(elaboroSet);
  return result;
}

function _showImportPreview(parsed, fileName) {
  STATE._importData = parsed;
  STATE._importFileName = fileName;
  STATE.modal = 'import';
  refreshModal();
}

function renderImportModal() {
  var data = STATE._importData;
  if (!data) return null;

  return h('div', {className:'modal-overlay', onClick:function(e){if(e.target.className==='modal-overlay')closeModal();}},
    h('div', {className:'modal', style:{maxWidth:'700px'}},
      h('div', {className:'modal-header'},
        h('h2', null, h('i',{className:'fas fa-file-import', style:{marginRight:'8px',color:'var(--accent)'}}), 'Importar Datos'),
        h('button', {className:'modal-close', onClick:closeModal}, h('i',{className:'fas fa-times'}))
      ),
      h('div', {className:'modal-body'},
        h('div', {style:{background:'var(--bg-input)',borderRadius:'10px',padding:'14px',marginBottom:'16px',border:'1px solid var(--border)'}},
          h('div', {style:{fontSize:'14px',fontWeight:'600',color:'var(--text-primary)'}}, 'Archivo: ' + (STATE._importFileName || '')),
          h('div', {style:{fontSize:'13px',color:'var(--text-muted)',marginTop:'4px'}},
            data.records.length + ' registros encontrados' + (data.obsoletos.length > 0 ? ', ' + data.obsoletos.length + ' obsoletos' : '')
          )
        ),
        h('div', {style:{background:'var(--red-bg)',borderRadius:'8px',padding:'12px',marginBottom:'16px',border:'1px solid rgba(239,68,68,.2)',fontSize:'13px',color:'var(--red)'}},
          h('i', {className:'fas fa-exclamation-triangle', style:{marginRight:'6px'}}),
          'Esta acción reemplazará todos los registros actuales. Se recomienda hacer un respaldo antes de continuar.'
        ),
        data.records.length > 0 ? h('div', {style:{maxHeight:'200px',overflowY:'auto',border:'1px solid var(--border)',borderRadius:'8px',fontSize:'12px'}},
          h('table', {style:{width:'100%',borderCollapse:'collapse'}},
            h('thead', null,
              h('tr', null,
                h('th', {style:{padding:'6px 10px',textAlign:'left',borderBottom:'1px solid var(--border)',background:'var(--bg-input)'}}, 'Nombre'),
                h('th', {style:{padding:'6px 10px',textAlign:'left',borderBottom:'1px solid var(--border)',background:'var(--bg-input)'}}, 'Área'),
                h('th', {style:{padding:'6px 10px',textAlign:'left',borderBottom:'1px solid var(--border)',background:'var(--bg-input)'}}, 'Código')
              )
            ),
            h('tbody', null,
              ...data.records.slice(0, 20).map(function(r) {
                return h('tr', null,
                  h('td', {style:{padding:'4px 10px',borderBottom:'1px solid var(--border)'}}, r.nombreDoc || '—'),
                  h('td', {style:{padding:'4px 10px',borderBottom:'1px solid var(--border)'}}, r.area || '—'),
                  h('td', {style:{padding:'4px 10px',borderBottom:'1px solid var(--border)'}}, r.codManual || r.codFormato || r.codInstructivo || '—')
                );
              })
            )
          ),
          data.records.length > 20 ? h('div', {style:{padding:'6px 10px',fontSize:'11px',color:'var(--text-muted)',textAlign:'center'}}, '... y ' + (data.records.length - 20) + ' más') : null
        ) : null
      ),
      h('div', {className:'modal-footer'},
        h('button', {className:'btn btn-secondary', onClick:closeModal}, 'Cancelar'),
        h('button', {className:'btn btn-warning', onClick:function(){
          _applyImport(data);
          closeModal();
        }}, h('i',{className:'fas fa-file-import'}), 'Importar ' + data.records.length + ' registros')
      )
    )
  );
}

function _applyImport(data) {
  if (!data || !data.records) return;
  var nextId = STATE.nextId || 1000;
  data.records.forEach(function(r, i) { r.id = nextId + i; });
  STATE.records = data.records;
  if (data.obsoletos && data.obsoletos.length > 0) {
    data.obsoletos.forEach(function(o, i) { o.id = nextId + data.records.length + i; });
    STATE.obsoletos = data.obsoletos;
  }
  if (data.elaboros && data.elaboros.length > 0) {
    var set = {};
    STATE.elaboros.forEach(function(e) { set[e] = true; });
    data.elaboros.forEach(function(e) { set[e] = true; });
    STATE.elaboros = Object.keys(set);
  }
  STATE.nextId = nextId + data.records.length + (data.obsoletos ? data.obsoletos.length : 0);
  _unsavedChanges = true;
  saveState();
  addLog('import', 'Importados ' + data.records.length + ' registros');
  showToast('Importados ' + data.records.length + ' registros correctamente');
  render();
}

function saveRecord() {
  var form = STATE.editForm;
  var isAdd = STATE.modal === 'add';

  if (!form.area) { showToast('El área es obligatoria', 'error'); return; }
  if (!form.nombreDoc) { showToast('El nombre del documento es obligatorio', 'error'); return; }

  var vChanged = !isAdd && STATE.originalVersion && String(form.version||'').trim() !== '' && String(form.version||'').trim() !== STATE.originalVersion;
  if (vChanged && (!STATE.obsoletoMotivo || !STATE.obsoletoMotivo.trim())) {
    showToast('El motivo del cambio de versión es obligatorio', 'error');
    return;
  }

  if (isAdd) {
    var newRec = Object.assign({}, form, { id: getNextId() });
    STATE.records.push(newRec);
    _trackChange(newRec, 'add');
    addLog('create', 'Registro agregado', newRec.nombreDoc || newRec.codManual);
    if (newRec.elaboro && STATE.elaboros.indexOf(newRec.elaboro) < 0) STATE.elaboros.push(newRec.elaboro);
  } else {
    var idx = STATE.records.findIndex(function(r) { return r.id === form.id; });
    if (idx >= 0) {
      if (vChanged) {
        var oldRec = STATE.records[idx];
        var obsEntry = createObsoletoEntry(oldRec, STATE.obsoletoMotivo);
        STATE.obsoletos.push(obsEntry);
        addLog('obsolete', 'Versión anterior enviada a obsoletos: ' + STATE.obsoletoMotivo, oldRec.nombreDoc || oldRec.codManual);
      }
      STATE.records[idx] = Object.assign({}, form);
      _trackChange(STATE.records[idx], 'edit');
      addLog('edit', 'Registro editado', form.nombreDoc || form.codManual);
      if (form.elaboro && STATE.elaboros.indexOf(form.elaboro) < 0) STATE.elaboros.push(form.elaboro);
    }
  }

  _unsavedChanges = true;
  saveState();
  closeModal();
  render();
}

function confirmSalida() {
  var rec = STATE.selectedRecord;
  var form = STATE.salidaForm;
  if (!rec) return;
  if (!form.area || !form.destinatario || !form.correo || !form.motivo) {
    showToast('Todos los campos marcados con * son obligatorios', 'error');
    return;
  }
  var entry = {
    id: getNextId(),
    fecha: new Date().toISOString(),
    docId: rec.id,
    nombreDoc: rec.nombreDoc || rec.codManual || rec.codFormato || '',
    version: rec.version || '',
    area: form.area,
    destinatario: form.destinatario,
    correo: form.correo,
    motivo: form.motivo,
    usuario: AUTH.currentUser ? (AUTH.currentUser.username || '') : ''
  };
  STATE.salidas.push(entry);
  _trackChange(entry, 'salida');
  addLog('salida', 'Salida registrada a ' + form.destinatario, entry.nombreDoc);
  _unsavedChanges = true;
  saveState();
  closeModal();
  showToast('Salida registrada correctamente');
  render();
}

function confirmObsoleto() {
  var rec = STATE.selectedRecord;
  if (!rec || !STATE.obsoletoMotivo.trim()) return;
  var obsEntry = createObsoletoEntry(rec, STATE.obsoletoMotivo);
  STATE.obsoletos.push(obsEntry);
  var idx = STATE.records.findIndex(function(r) { return r.id === rec.id; });
  if (idx >= 0) STATE.records.splice(idx, 1);
  _trackChange(obsEntry, 'obsolete');
  addLog('obsolete', 'Documento enviado a obsoletos: ' + STATE.obsoletoMotivo, rec.nombreDoc || rec.codManual);
  _unsavedChanges = true;
  saveState();
  closeModal();
  showToast('Documento enviado a obsoletos');
  render();
}

function _confirmRenew() {
  var rec = STATE.selectedRecord;
  if (!rec || !STATE._renewDate || !STATE._renewMotivo || !STATE._renewMotivo.trim()) return;
  var idx = STATE.records.findIndex(function(r) { return r.id === rec.id; });
  if (idx >= 0) {
    STATE.records[idx].fechaEmision = STATE._renewDate;
    _trackChange(STATE.records[idx], 'renew');
    addLog('renew', 'Vigencia renovada: ' + STATE._renewMotivo, rec.nombreDoc || rec.codManual);
  }
  _unsavedChanges = true;
  saveState();
  closeModal();
  showToast('Vigencia renovada correctamente');
  render();
}

function deleteRecord(id) {
  var idx = STATE.records.findIndex(function(r) { return r.id === id; });
  if (idx >= 0) {
    var rec = STATE.records.splice(idx, 1)[0];
    rec._deletedAt = new Date().toISOString();
    STATE.papelera.push(rec);
    _trackChange(rec, 'delete');
    addLog('delete', 'Registro eliminado (papelera)', rec.nombreDoc || rec.codManual);
    _unsavedChanges = true;
    saveState();
    showToast('Registro movido a papelera');
    render();
  }
}

function restoreRecord(id) {
  var idx = STATE.papelera.findIndex(function(r) { return r.id === id; });
  if (idx >= 0) {
    var rec = STATE.papelera.splice(idx, 1)[0];
    delete rec._deletedAt;
    STATE.records.push(rec);
    _trackChange(rec, 'restore');
    addLog('restore', 'Registro restaurado de papelera', rec.nombreDoc || rec.codManual);
    _unsavedChanges = true;
    saveState();
    showToast('Registro restaurado');
    render();
  }
}

function permanentDelete(id) {
  var idx = STATE.papelera.findIndex(function(r) { return r.id === id; });
  if (idx >= 0) {
    var rec = STATE.papelera.splice(idx, 1)[0];
    addLog('permDelete', 'Registro eliminado permanentemente', rec.nombreDoc || rec.codManual);
    _unsavedChanges = true;
    saveState();
    showToast('Registro eliminado permanentemente');
    render();
  }
}

function _sortBy(col) {
  if (STATE.sortCol === col) {
    STATE.sortDir = STATE.sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    STATE.sortCol = col;
    STATE.sortDir = 'asc';
  }
  render();
}

function _renderSortIcon(col) {
  if (STATE.sortCol !== col) return h('i', {className: 'fas fa-sort', style:{opacity:'.3',marginLeft:'4px',fontSize:'10px'}});
  return h('i', {className: STATE.sortDir === 'asc' ? 'fas fa-sort-up' : 'fas fa-sort-down', style:{marginLeft:'4px',fontSize:'10px',color:'var(--accent)'}});
}

function _renderDashboard() {
  var recs = STATE.records || [];
  var total = recs.length;
  var vigentes = 0, porVencer = 0, vencidos = 0, sinFecha = 0;
  var porArea = {};
  var porElaboro = {};

  AREAS.forEach(function(a) { porArea[a] = {total:0,vigentes:0,porVencer:0,vencidos:0}; });

  recs.forEach(function(r) {
    var estado = typeof calcEstado === 'function' ? calcEstado(r.fechaEmision) : '';
    if (estado === 'Vigente') vigentes++;
    else if (estado === 'Por vencer') porVencer++;
    else if (estado === 'Vencido') vencidos++;
    else sinFecha++;

    if (porArea[r.area]) {
      porArea[r.area].total++;
      if (estado === 'Vigente') porArea[r.area].vigentes++;
      else if (estado === 'Por vencer') porArea[r.area].porVencer++;
      else if (estado === 'Vencido') porArea[r.area].vencidos++;
    }

    if (r.elaboro) {
      if (!porElaboro[r.elaboro]) porElaboro[r.elaboro] = 0;
      porElaboro[r.elaboro]++;
    }
  });

  var obsTotal = (STATE.obsoletos || []).length;
  var salTotal = (STATE.salidas || []).length;
  var papTotal = (STATE.papelera || []).length;
  var logTotal = (STATE.logs || []).length;
  var pctVigente = total > 0 ? Math.round((vigentes / total) * 100) : 0;
  var pctPorVencer = total > 0 ? Math.round((porVencer / total) * 100) : 0;
  var pctVencido = total > 0 ? Math.round((vencidos / total) * 100) : 0;

  var topElaboradores = Object.entries(porElaboro).sort(function(a,b){return b[1]-a[1];}).slice(0, 8);
  var maxElab = topElaboradores.length > 0 ? topElaboradores[0][1] : 1;

  var areaEntries = AREAS.map(function(a) { return {name: a, data: porArea[a] || {total:0,vigentes:0,porVencer:0,vencidos:0}}; }).filter(function(e){return e.data.total > 0;});
  var maxArea = Math.max.apply(null, areaEntries.map(function(e){return e.data.total;})) || 1;

  // Últimas actividades (logs recientes)
  var recentLogs = (STATE.logs || []).slice(0, 5);

  // Fecha/hora actual
  var now = new Date();
  var fechaHoy = now.toLocaleDateString('es-MX', {weekday:'long', year:'numeric', month:'long', day:'numeric'});
  var horaHoy = now.toLocaleTimeString('es-MX', {hour:'2-digit', minute:'2-digit'});

  function statCard(icon, label, value, color, sub) {
    return h('div', {style:{background:'var(--bg-card)',borderRadius:'16px',padding:'20px',border:'1px solid var(--border)',position:'relative',overflow:'hidden'}},
      h('div', {style:{position:'absolute',top:0,left:0,right:0,height:'3px',background:color}}),
      h('div', {style:{display:'flex',alignItems:'center',gap:'12px',marginBottom:'8px'}},
        h('div', {style:{width:'40px',height:'40px',borderRadius:'12px',background:color+'18',display:'flex',alignItems:'center',justifyContent:'center'}},
          h('i', {className:'fas ' + icon, style:{color:color,fontSize:'16px'}})
        ),
        h('div', {style:{fontSize:'11px',color:'var(--text-muted)',textTransform:'uppercase',letterSpacing:'.5px',fontWeight:'600'}}, label)
      ),
      h('div', {style:{fontSize:'32px',fontWeight:'800',color:'var(--text-primary)',fontFamily:'JetBrains Mono, monospace'}}, String(value)),
      sub ? h('div', {style:{fontSize:'11px',color:'var(--text-muted)',marginTop:'4px'}}, sub) : null
    );
  }

  function barRow(label, value, max, color) {
    var pct = max > 0 ? Math.round((value / max) * 100) : 0;
    return h('div', {style:{display:'flex',alignItems:'center',gap:'10px',marginBottom:'8px'}},
      h('div', {style:{width:'100px',fontSize:'12px',color:'var(--text-secondary)',textOverflow:'ellipsis',overflow:'hidden',whiteSpace:'nowrap',flexShrink:0}}, label),
      h('div', {style:{flex:1,height:'22px',background:'var(--bg-input)',borderRadius:'6px',overflow:'hidden',position:'relative'}},
        h('div', {style:{width:pct+'%',height:'100%',background:'linear-gradient(90deg,'+color+','+color+'cc)',borderRadius:'6px',transition:'width .6s ease',minWidth: value > 0 ? '2px' : '0'}}),
        h('span', {style:{position:'absolute',right:'8px',top:'3px',fontSize:'11px',fontWeight:'600',color:'var(--text-primary)',fontFamily:'JetBrains Mono, monospace'}}, String(value))
      )
    );
  }

  function stackedBarRow(label, data, maxTotal) {
    var pctV = maxTotal > 0 ? (data.vigentes / maxTotal) * 100 : 0;
    var pctP = maxTotal > 0 ? (data.porVencer / maxTotal) * 100 : 0;
    var pctX = maxTotal > 0 ? (data.vencidos / maxTotal) * 100 : 0;
    var rest = maxTotal > 0 ? ((data.total - data.vigentes - data.porVencer - data.vencidos) / maxTotal) * 100 : 0;
    return h('div', {style:{display:'flex',alignItems:'center',gap:'10px',marginBottom:'6px'}},
      h('div', {style:{width:'110px',fontSize:'12px',color:'var(--text-secondary)',textOverflow:'ellipsis',overflow:'hidden',whiteSpace:'nowrap',flexShrink:0,fontWeight:'500'}}, label),
      h('div', {style:{flex:1,height:'26px',background:'var(--bg-input)',borderRadius:'8px',overflow:'hidden',display:'flex',position:'relative'}},
        pctV > 0 ? h('div', {style:{width:pctV+'%',height:'100%',background:'linear-gradient(90deg, #2ed573, #26c46b)',transition:'width .6s ease'}}) : null,
        pctP > 0 ? h('div', {style:{width:pctP+'%',height:'100%',background:'linear-gradient(90deg, #ffa502, #e89600)',transition:'width .6s ease'}}) : null,
        pctX > 0 ? h('div', {style:{width:pctX+'%',height:'100%',background:'linear-gradient(90deg, #ff4757, #e03e4d)',transition:'width .6s ease'}}) : null,
        rest > 0 ? h('div', {style:{width:rest+'%',height:'100%',background:'rgba(100,100,120,0.3)',transition:'width .6s ease'}}) : null,
        h('span', {style:{position:'absolute',right:'10px',top:'4px',fontSize:'12px',fontWeight:'700',color:'var(--text-primary)',fontFamily:'JetBrains Mono, monospace',textShadow:'0 1px 2px rgba(0,0,0,.5)'}}, String(data.total))
      )
    );
  }

  function circleGauge(pct, color, label, count) {
    return h('div', {style:{background:'var(--bg-card)',borderRadius:'16px',padding:'24px 20px',border:'1px solid var(--border)',textAlign:'center'}},
      h('div', {style:{width:'100px',height:'100px',borderRadius:'50%',margin:'0 auto 12px',background:'conic-gradient('+color+' 0deg, '+color+' '+(pct*3.6)+'deg, var(--bg-input) '+(pct*3.6)+'deg)',display:'flex',alignItems:'center',justifyContent:'center'}},
        h('div', {style:{width:'76px',height:'76px',borderRadius:'50%',background:'var(--bg-card)',display:'flex',alignItems:'center',justifyContent:'center',flexDirection:'column'}},
          h('div', {style:{fontSize:'20px',fontWeight:'800',fontFamily:'JetBrains Mono, monospace',color:color,lineHeight:'1'}}, pct + '%'),
          h('div', {style:{fontSize:'10px',color:'var(--text-muted)',marginTop:'2px'}}, count + ' docs')
        )
      ),
      h('div', {style:{fontSize:'13px',fontWeight:'600',color:'var(--text-secondary)'}}, label)
    );
  }

  return h('div', {className:'content-area', id:'dashboard-content'},
    h('div', {className:'topbar'},
      h('span', {className:'topbar-title'}, h('i', {className:'fas fa-chart-pie', style:{marginRight:'8px'}}), 'Dashboard'),
      h('div', {className:'topbar-actions'},
        h('button', {className:'btn btn-sm btn-secondary', onClick:function(){ _exportDashboardPDF(); }},
          h('i', {className:'fas fa-file-pdf'}), ' Descargar PDF'
        )
      )
    ),
    h('div', {style:{padding:'20px',overflowY:'auto',flex:1}},
      // Fecha y resumen rápido
      h('div', {style:{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'20px',padding:'14px 20px',background:'var(--bg-card)',borderRadius:'12px',border:'1px solid var(--border)'}},
        h('div', null,
          h('div', {style:{fontSize:'15px',fontWeight:'700',color:'var(--text-primary)',textTransform:'capitalize'}}, fechaHoy),
          h('div', {style:{fontSize:'12px',color:'var(--text-muted)',marginTop:'2px'}}, 'Última actualización: ' + horaHoy)
        ),
        h('div', {style:{display:'flex',gap:'16px',fontSize:'12px'}},
          h('span', {style:{color:'var(--green)',fontWeight:'600'}}, '● ' + vigentes + ' vigentes'),
          h('span', {style:{color:'var(--orange)',fontWeight:'600'}}, '● ' + porVencer + ' por vencer'),
          h('span', {style:{color:'var(--red)',fontWeight:'600'}}, '● ' + vencidos + ' vencidos')
        )
      ),

      // Stat cards
      h('div', {style:{display:'grid',gridTemplateColumns:'repeat(auto-fit, minmax(180px, 1fr))',gap:'14px',marginBottom:'24px'}},
        statCard('fa-file-lines', 'Total Documentos', total, 'var(--accent)', (total + obsTotal) + ' total con obsoletos'),
        statCard('fa-circle-check', 'Vigentes', vigentes, 'var(--green)', pctVigente + '% del total'),
        statCard('fa-clock', 'Por Vencer', porVencer, 'var(--orange)', 'Requieren atención'),
        statCard('fa-triangle-exclamation', 'Vencidos', vencidos, 'var(--red)', pctVencido + '% del total'),
        statCard('fa-right-from-bracket', 'Salidas', salTotal, 'var(--purple)', 'Registros de salida'),
        statCard('fa-box-archive', 'Obsoletos', obsTotal, '#8b8b8b', 'Documentos obsoletos'),
        statCard('fa-trash-can', 'Papelera', papTotal, '#ef4444', papTotal > 0 ? 'Pendientes de eliminar' : 'Vacía'),
        statCard('fa-list-check', 'Actividades', logTotal, '#3b82f6', 'Registros en bitácora')
      ),

      // Charts row: Área + Elaboradores
      h('div', {style:{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'20px',marginBottom:'24px'}},
        // Documentos por Área
        h('div', {style:{background:'var(--bg-card)',borderRadius:'16px',padding:'20px',border:'1px solid var(--border)'}},
          h('div', {style:{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'16px'}},
            h('div', {style:{fontSize:'14px',fontWeight:'700',color:'var(--text-primary)'}},
              h('i', {className:'fas fa-building', style:{marginRight:'8px',color:'var(--accent)'}}), 'Documentos por Área'
            ),
            h('div', {style:{display:'flex',gap:'12px',fontSize:'10px',fontWeight:'600'}},
              h('span', {style:{display:'flex',alignItems:'center',gap:'3px'}}, h('span', {style:{width:'8px',height:'8px',borderRadius:'2px',background:'#2ed573',display:'inline-block'}}), 'Vig'),
              h('span', {style:{display:'flex',alignItems:'center',gap:'3px'}}, h('span', {style:{width:'8px',height:'8px',borderRadius:'2px',background:'#ffa502',display:'inline-block'}}), 'P/V'),
              h('span', {style:{display:'flex',alignItems:'center',gap:'3px'}}, h('span', {style:{width:'8px',height:'8px',borderRadius:'2px',background:'#ff4757',display:'inline-block'}}), 'Venc')
            )
          ),
          ...areaEntries.map(function(e) {
            return stackedBarRow(e.name, e.data, maxArea);
          })
        ),
        // Top Elaboradores
        h('div', {style:{background:'var(--bg-card)',borderRadius:'16px',padding:'20px',border:'1px solid var(--border)'}},
          h('div', {style:{fontSize:'14px',fontWeight:'700',marginBottom:'16px',color:'var(--text-primary)'}},
            h('i', {className:'fas fa-user-pen', style:{marginRight:'8px',color:'var(--purple)'}}), 'Top Elaboradores'
          ),
          ...topElaboradores.map(function(e) {
            return barRow(e[0], e[1], maxElab, 'var(--purple)');
          }),
          topElaboradores.length === 0 ? h('div', {style:{textAlign:'center',padding:'20px',color:'var(--text-muted)',fontSize:'13px'}}, 'Sin datos') : null
        )
      ),

      // Gauges + Tabla resumen + Actividad reciente
      h('div', {style:{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'20px',marginBottom:'24px'}},
        // Gauges
        h('div', {style:{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:'14px'}},
          circleGauge(pctVigente, 'var(--green)', 'Vigentes', vigentes),
          circleGauge(pctPorVencer, 'var(--orange)', 'Por Vencer', porVencer),
          circleGauge(pctVencido, 'var(--red)', 'Vencidos', vencidos)
        ),
        // Actividad reciente
        h('div', {style:{background:'var(--bg-card)',borderRadius:'16px',padding:'20px',border:'1px solid var(--border)'}},
          h('div', {style:{fontSize:'14px',fontWeight:'700',marginBottom:'14px',color:'var(--text-primary)'}},
            h('i', {className:'fas fa-clock-rotate-left', style:{marginRight:'8px',color:'#3b82f6'}}), 'Actividad Reciente'
          ),
          recentLogs.length === 0
            ? h('div', {style:{textAlign:'center',padding:'20px',color:'var(--text-muted)',fontSize:'13px'}}, 'Sin actividad registrada')
            : h('div', null, ...recentLogs.map(function(l) {
                var icon = (typeof LOG_TYPES !== 'undefined' && LOG_TYPES[l.type]) || {icon:'fa-circle-info',color:'#888',label:l.type};
                var ts = l.timestamp || l.date || l.ts;
                var timeStr = '';
                if (ts) { try { var d = new Date(ts); if (!isNaN(d.getTime())) timeStr = d.toLocaleTimeString('es-MX', {hour:'2-digit',minute:'2-digit'}); } catch(e){} }
                return h('div', {style:{display:'flex',alignItems:'center',gap:'10px',padding:'8px 0',borderBottom:'1px solid var(--border)'}},
                  h('div', {style:{width:'28px',height:'28px',borderRadius:'8px',background:icon.color+'18',display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0}},
                    h('i', {className:'fas ' + icon.icon, style:{fontSize:'11px',color:icon.color}})
                  ),
                  h('div', {style:{flex:1,minWidth:0}},
                    h('div', {style:{fontSize:'12px',fontWeight:'600',color:'var(--text-primary)',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}, l.details || icon.label),
                    h('div', {style:{fontSize:'10px',color:'var(--text-muted)'}}, (l.user || '') + (timeStr ? ' · ' + timeStr : ''))
                  )
                );
              }))
        )
      ),

      // Tabla resumen por área
      h('div', {style:{background:'var(--bg-card)',borderRadius:'16px',padding:'20px',border:'1px solid var(--border)',marginBottom:'20px'}},
        h('div', {style:{fontSize:'14px',fontWeight:'700',marginBottom:'14px',color:'var(--text-primary)'}},
          h('i', {className:'fas fa-table', style:{marginRight:'8px',color:'var(--accent)'}}), 'Resumen por Área'
        ),
        h('div', {style:{overflowX:'auto'}},
          h('table', {style:{width:'100%',borderCollapse:'collapse',fontSize:'12px'}},
            h('thead', null,
              h('tr', {style:{borderBottom:'2px solid var(--border)'}},
                h('th', {style:{textAlign:'left',padding:'8px 12px',fontWeight:'700',color:'var(--accent)'}}, 'Área'),
                h('th', {style:{textAlign:'center',padding:'8px 12px',fontWeight:'700'}}, 'Total'),
                h('th', {style:{textAlign:'center',padding:'8px 12px',fontWeight:'700',color:'#2ed573'}}, 'Vigentes'),
                h('th', {style:{textAlign:'center',padding:'8px 12px',fontWeight:'700',color:'#ffa502'}}, 'Por Vencer'),
                h('th', {style:{textAlign:'center',padding:'8px 12px',fontWeight:'700',color:'#ff4757'}}, 'Vencidos'),
                h('th', {style:{textAlign:'center',padding:'8px 12px',fontWeight:'700',color:'#2ed573'}}, '% Vigente')
              )
            ),
            h('tbody', null,
              ...areaEntries.map(function(e) {
                var pv = e.data.total > 0 ? Math.round((e.data.vigentes / e.data.total) * 100) : 0;
                return h('tr', {style:{borderBottom:'1px solid var(--border)'}},
                  h('td', {style:{padding:'8px 12px',fontWeight:'600'}}, e.name),
                  h('td', {style:{textAlign:'center',padding:'8px 12px',fontWeight:'700',fontFamily:'JetBrains Mono, monospace'}}, String(e.data.total)),
                  h('td', {style:{textAlign:'center',padding:'8px 12px',color:'#2ed573',fontFamily:'JetBrains Mono, monospace'}}, String(e.data.vigentes)),
                  h('td', {style:{textAlign:'center',padding:'8px 12px',color:'#ffa502',fontFamily:'JetBrains Mono, monospace'}}, String(e.data.porVencer)),
                  h('td', {style:{textAlign:'center',padding:'8px 12px',color:'#ff4757',fontFamily:'JetBrains Mono, monospace'}}, String(e.data.vencidos)),
                  h('td', {style:{textAlign:'center',padding:'8px 12px'}},
                    h('span', {style:{background: pv >= 80 ? '#2ed57322' : pv >= 50 ? '#ffa50222' : '#ff475722', color: pv >= 80 ? '#2ed573' : pv >= 50 ? '#ffa502' : '#ff4757', padding:'2px 10px',borderRadius:'10px',fontWeight:'700',fontSize:'11px',fontFamily:'JetBrains Mono, monospace'}}, pv + '%')
                  )
                );
              }),
              // Fila total
              h('tr', {style:{borderTop:'2px solid var(--border)',fontWeight:'800'}},
                h('td', {style:{padding:'10px 12px',fontWeight:'800',color:'var(--accent)'}}, 'TOTAL'),
                h('td', {style:{textAlign:'center',padding:'10px 12px',fontFamily:'JetBrains Mono, monospace',fontSize:'13px'}}, String(total)),
                h('td', {style:{textAlign:'center',padding:'10px 12px',color:'#2ed573',fontFamily:'JetBrains Mono, monospace',fontSize:'13px'}}, String(vigentes)),
                h('td', {style:{textAlign:'center',padding:'10px 12px',color:'#ffa502',fontFamily:'JetBrains Mono, monospace',fontSize:'13px'}}, String(porVencer)),
                h('td', {style:{textAlign:'center',padding:'10px 12px',color:'#ff4757',fontFamily:'JetBrains Mono, monospace',fontSize:'13px'}}, String(vencidos)),
                h('td', {style:{textAlign:'center',padding:'10px 12px'}},
                  h('span', {style:{background:'#2ed57322',color:'#2ed573',padding:'2px 10px',borderRadius:'10px',fontWeight:'800',fontSize:'12px',fontFamily:'JetBrains Mono, monospace'}}, pctVigente + '%')
                )
              )
            )
          )
        )
      )
    )
  );
}

// Exportar Dashboard a PDF
function _exportDashboardPDF() {
  var _jsPDFClass = null;
  try {
    if (typeof jspdf !== 'undefined' && jspdf.jsPDF) _jsPDFClass = jspdf.jsPDF;
    else if (typeof window.jspdf !== 'undefined' && window.jspdf.jsPDF) _jsPDFClass = window.jspdf.jsPDF;
    else if (typeof window.jsPDF !== 'undefined') _jsPDFClass = window.jsPDF;
  } catch(e) {}
  if (!_jsPDFClass) {
    showToast('Librería jsPDF no disponible. Recarga la página e intenta de nuevo.', 'error');
    console.error('jsPDF not found. window.jspdf:', typeof window.jspdf, 'window.jsPDF:', typeof window.jsPDF);
    return;
  }

  showToast('Generando PDF del Dashboard...', 'info');

  try {
    var recs = STATE.records || [];
    var total = recs.length;
    var vigentes = 0, porVencer = 0, vencidos = 0;
    var porArea = {};
    AREAS.forEach(function(a) { porArea[a] = {total:0,vigentes:0,porVencer:0,vencidos:0}; });
    recs.forEach(function(r) {
      var estado = typeof calcEstado === 'function' ? calcEstado(r.fechaEmision) : '';
      if (estado === 'Vigente') vigentes++;
      else if (estado === 'Por vencer') porVencer++;
      else if (estado === 'Vencido') vencidos++;
      if (porArea[r.area]) {
        porArea[r.area].total++;
        if (estado === 'Vigente') porArea[r.area].vigentes++;
        else if (estado === 'Por vencer') porArea[r.area].porVencer++;
        else if (estado === 'Vencido') porArea[r.area].vencidos++;
      }
    });
    var obsTotal = (STATE.obsoletos || []).length;
    var salTotal = (STATE.salidas || []).length;
    var papTotal = (STATE.papelera || []).length;
    var pctV = total > 0 ? Math.round((vigentes/total)*100) : 0;
    var pctP = total > 0 ? Math.round((porVencer/total)*100) : 0;
    var pctX = total > 0 ? Math.round((vencidos/total)*100) : 0;

    var pdf = new _jsPDFClass({orientation:'landscape',unit:'mm',format:'letter'});
    var W = pdf.internal.pageSize.getWidth();
    var H = pdf.internal.pageSize.getHeight();
    var now = new Date();
    var fechaStr = now.toLocaleDateString('es-MX',{year:'numeric',month:'long',day:'numeric'});
    var horaStr = now.toLocaleTimeString('es-MX',{hour:'2-digit',minute:'2-digit'});

    // Colors
    var darkBg = [10, 22, 40];
    var cardBg = [17, 29, 46];
    var accent = [0, 212, 255];
    var green = [46, 213, 115];
    var orange = [255, 165, 2];
    var red = [255, 71, 87];
    var white = [255, 255, 255];
    var muted = [130, 150, 170];

    // Background
    pdf.setFillColor.apply(pdf, darkBg);
    pdf.rect(0, 0, W, H, 'F');

    // Header
    pdf.setFillColor.apply(pdf, cardBg);
    pdf.roundedRect(10, 8, W-20, 20, 3, 3, 'F');
    pdf.setFontSize(18); pdf.setFont('helvetica','bold'); pdf.setTextColor.apply(pdf, accent);
    pdf.text('DEBBIOM - Dashboard de Control Documental', 18, 20);
    pdf.setFontSize(10); pdf.setTextColor.apply(pdf, muted); pdf.setFont('helvetica','normal');
    pdf.text(fechaStr + ' | ' + horaStr + ' | Generado por: ' + (AUTH.currentUser ? (AUTH.currentUser.fullName || AUTH.currentUser.username) : 'Sistema'), 18, 25);

    // Stat boxes row
    var boxW = (W - 20 - 35) / 6;  // 6 boxes
    var boxY = 34;
    var stats = [
      {label:'Total Docs', val:String(total), color:accent},
      {label:'Vigentes', val:String(vigentes) + ' (' + pctV + '%)', color:green},
      {label:'Por Vencer', val:String(porVencer) + ' (' + pctP + '%)', color:orange},
      {label:'Vencidos', val:String(vencidos) + ' (' + pctX + '%)', color:red},
      {label:'Salidas', val:String(salTotal), color:[168,85,247]},
      {label:'Obsoletos', val:String(obsTotal), color:muted}
    ];
    stats.forEach(function(s, i) {
      var bx = 10 + i * (boxW + 7);
      pdf.setFillColor.apply(pdf, cardBg);
      pdf.roundedRect(bx, boxY, boxW, 22, 2, 2, 'F');
      // Top accent line
      pdf.setFillColor.apply(pdf, s.color);
      pdf.rect(bx, boxY, boxW, 2, 'F');
      pdf.setFontSize(8); pdf.setTextColor.apply(pdf, muted); pdf.setFont('helvetica','normal');
      pdf.text(s.label, bx + 5, boxY + 9);
      pdf.setFontSize(16); pdf.setTextColor.apply(pdf, white); pdf.setFont('helvetica','bold');
      pdf.text(s.val, bx + 5, boxY + 19);
    });

    // Table: Resumen por Área
    var tblY = 62;
    pdf.setFillColor.apply(pdf, cardBg);
    pdf.roundedRect(10, tblY, W - 20, 8 + AREAS.length * 7 + 9, 2, 2, 'F');
    pdf.setFontSize(11); pdf.setTextColor.apply(pdf, accent); pdf.setFont('helvetica','bold');
    pdf.text('Resumen por Area', 16, tblY + 6);

    // Table header
    var colX = [16, 70, 105, 140, 175, 210];
    var hdrY = tblY + 13;
    pdf.setFontSize(8); pdf.setFont('helvetica','bold');
    pdf.setTextColor.apply(pdf, accent);
    pdf.text('Area', colX[0], hdrY);
    pdf.text('Total', colX[1], hdrY);
    pdf.setTextColor.apply(pdf, green); pdf.text('Vigentes', colX[2], hdrY);
    pdf.setTextColor.apply(pdf, orange); pdf.text('Por Vencer', colX[3], hdrY);
    pdf.setTextColor.apply(pdf, red); pdf.text('Vencidos', colX[4], hdrY);
    pdf.setTextColor.apply(pdf, green); pdf.text('% Vigente', colX[5], hdrY);

    // Separator
    pdf.setDrawColor(40, 60, 90); pdf.setLineWidth(0.3);
    pdf.line(colX[0], hdrY + 2, 240, hdrY + 2);

    var rowY = hdrY + 7;
    AREAS.forEach(function(area) {
      var d = porArea[area] || {total:0,vigentes:0,porVencer:0,vencidos:0};
      if (d.total === 0) return;
      var pv = d.total > 0 ? Math.round((d.vigentes/d.total)*100) : 0;
      pdf.setFont('helvetica','normal'); pdf.setFontSize(8);
      pdf.setTextColor.apply(pdf, white); pdf.text(area, colX[0], rowY);
      pdf.setTextColor.apply(pdf, white); pdf.text(String(d.total), colX[1], rowY);
      pdf.setTextColor.apply(pdf, green); pdf.text(String(d.vigentes), colX[2], rowY);
      pdf.setTextColor.apply(pdf, orange); pdf.text(String(d.porVencer), colX[3], rowY);
      pdf.setTextColor.apply(pdf, red); pdf.text(String(d.vencidos), colX[4], rowY);
      // % vigente with color
      var pvColor = pv >= 80 ? green : pv >= 50 ? orange : red;
      pdf.setTextColor.apply(pdf, pvColor); pdf.text(pv + '%', colX[5], rowY);
      rowY += 7;
    });

    // Total row
    pdf.setDrawColor(40, 60, 90); pdf.line(colX[0], rowY - 3, 240, rowY - 3);
    pdf.setFont('helvetica','bold'); pdf.setFontSize(9);
    pdf.setTextColor.apply(pdf, accent); pdf.text('TOTAL', colX[0], rowY + 1);
    pdf.setTextColor.apply(pdf, white); pdf.text(String(total), colX[1], rowY + 1);
    pdf.setTextColor.apply(pdf, green); pdf.text(String(vigentes), colX[2], rowY + 1);
    pdf.setTextColor.apply(pdf, orange); pdf.text(String(porVencer), colX[3], rowY + 1);
    pdf.setTextColor.apply(pdf, red); pdf.text(String(vencidos), colX[4], rowY + 1);
    pdf.setTextColor.apply(pdf, green); pdf.text(pctV + '%', colX[5], rowY + 1);

    // Right side: Stacked bars visualization
    var barX = W/2 + 15;
    var barBoxW = W/2 - 25;
    pdf.setFillColor.apply(pdf, cardBg);
    pdf.roundedRect(barX - 5, tblY, barBoxW, 8 + AREAS.length * 7 + 9, 2, 2, 'F');
    pdf.setFontSize(11); pdf.setTextColor.apply(pdf, accent); pdf.setFont('helvetica','bold');
    pdf.text('Documentos por Area', barX + 2, tblY + 6);

    // Legend
    var lgY = tblY + 5;
    var lgX = barX + barBoxW - 60;
    pdf.setFontSize(6); pdf.setFont('helvetica','normal');
    pdf.setFillColor.apply(pdf, green); pdf.rect(lgX, lgY - 2.5, 4, 3, 'F');
    pdf.setTextColor.apply(pdf, muted); pdf.text('Vig', lgX + 5, lgY);
    pdf.setFillColor.apply(pdf, orange); pdf.rect(lgX + 15, lgY - 2.5, 4, 3, 'F');
    pdf.text('P/V', lgX + 20, lgY);
    pdf.setFillColor.apply(pdf, red); pdf.rect(lgX + 30, lgY - 2.5, 4, 3, 'F');
    pdf.text('Venc', lgX + 35, lgY);

    var bRowY = tblY + 14;
    var maxAreaTotal = Math.max.apply(null, AREAS.map(function(a){ return (porArea[a]||{total:0}).total; })) || 1;
    var barMaxW = barBoxW - 50;

    AREAS.forEach(function(area) {
      var d = porArea[area] || {total:0,vigentes:0,porVencer:0,vencidos:0};
      if (d.total === 0) return;
      pdf.setFontSize(7); pdf.setFont('helvetica','normal'); pdf.setTextColor.apply(pdf, muted);
      pdf.text(area, barX + 2, bRowY + 1);

      var bStartX = barX + 38;
      var fullW = (d.total / maxAreaTotal) * barMaxW;
      var vW = d.total > 0 ? (d.vigentes / d.total) * fullW : 0;
      var pW = d.total > 0 ? (d.porVencer / d.total) * fullW : 0;
      var xW = d.total > 0 ? (d.vencidos / d.total) * fullW : 0;

      if (vW > 0) { pdf.setFillColor.apply(pdf, green); pdf.rect(bStartX, bRowY - 3, vW, 5, 'F'); }
      if (pW > 0) { pdf.setFillColor.apply(pdf, orange); pdf.rect(bStartX + vW, bRowY - 3, pW, 5, 'F'); }
      if (xW > 0) { pdf.setFillColor.apply(pdf, red); pdf.rect(bStartX + vW + pW, bRowY - 3, xW, 5, 'F'); }

      pdf.setFontSize(7); pdf.setTextColor.apply(pdf, white); pdf.setFont('helvetica','bold');
      pdf.text(String(d.total), bStartX + fullW + 3, bRowY + 1);

      bRowY += 7;
    });

    // Footer
    pdf.setFontSize(7); pdf.setTextColor.apply(pdf, muted); pdf.setFont('helvetica','normal');
    pdf.text('DEBBIOM - Listado Maestro de Informacion Documentada | Reporte generado automaticamente', 10, H - 5);
    pdf.text('Pagina 1 de 1', W - 30, H - 5);

    var fname = 'DEBBIOM_Dashboard_' + now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0') + '-' + String(now.getDate()).padStart(2,'0') + '.pdf';

    // Try save with fallback to blob URL
    try {
      pdf.save(fname);
    } catch(saveErr) {
      console.warn('pdf.save() failed, trying blob fallback:', saveErr);
      try {
        var blob = pdf.output('blob');
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = fname;
        document.body.appendChild(a);
        a.click();
        setTimeout(function() { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
      } catch(blobErr) {
        console.error('Blob fallback also failed:', blobErr);
        // Last resort: open in new window
        try {
          var dataUri = pdf.output('datauristring');
          window.open(dataUri, '_blank');
        } catch(uriErr) {
          throw new Error('No se pudo descargar el PDF: ' + saveErr.message);
        }
      }
    }

    showToast('PDF descargado: ' + fname);
    addLog('export', 'Dashboard PDF exportado: ' + fname);
  } catch(e) {
    showToast('Error generando PDF: ' + e.message, 'error');
    console.error('PDF error:', e);
  }
}

function exportBackupToExcel() {
  if (typeof XLSX === 'undefined') { showToast('Librería XLSX no disponible', 'error'); return; }

  var wb = XLSX.utils.book_new();
  var recs = STATE.records || [];
  var areas = AREAS;

  var headerStyle = {font:{bold:true,sz:10,name:'Arial',color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:'003366'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}}};
  var titleStyle = {font:{bold:true,sz:12,name:'Arial'},alignment:{horizontal:'center',vertical:'center'}};

  var copiaHeaders = ['Dirección General','Coordinación Clínica','Coordinación Analítica','Coord. Bioestadística','Coordinacion Administrativa','Soporte Técnico y TI','Recursos Humanos','Aseguramiento de Calidad','Lab. Clínico','Seguridad e Higiene','Mantenimiento'];

  function buildAreaSheet(areaName) {
    var areaRecs = recs.filter(function(r) { return r.area === areaName; });
    var rows = [];

    rows.push(['', '', 'LISTADO MAESTRO DE INFORMACION DOCUMENTADA', '', '', '', '', '', 'Código: L1/AC-P01', 'Versión: 009/Nov-25']);
    rows.push([]);
    var headerRow3 = ['', '', '', '', '', '', '', '', '', '', ''];
    headerRow3[11] = 'COPIA CONTROLADAS / RESPONSABLE';
    rows.push(headerRow3);

    var headers = ['', 'No.', 'Clasificación', 'Código del Manual / Procedimiento', 'Código del Instructivo', 'Código del Listado', 'Código del Formato', 'Nombre del Documento', 'Versión', 'Fecha Emisión', 'Elaboró / Actualizó'];
    copiaHeaders.forEach(function(ch) { headers.push(ch); });
    headers.push('Vigencia', 'Estado de Documentos');
    rows.push(headers);

    areaRecs.forEach(function(r, i) {
      var estado = typeof calcEstado === 'function' ? calcEstado(r.fechaEmision) : '';
      var vigencia = typeof renderVigenciaDate === 'function' ? renderVigenciaDate(r.fechaEmision) : '';
      var row = ['', r.no || String(i+1), r.area || areaName, r.codManual || '', r.codInstructivo || '', r.codListado || '', r.codFormato || '', r.nombreDoc || '', r.version || '', r.fechaEmision || '', r.elaboro || ''];
      copiaHeaders.forEach(function() { row.push(''); });
      row.push(vigencia, estado);
      rows.push(row);
    });

    return rows;
  }

  areas.forEach(function(area) {
    var data = buildAreaSheet(area);
    var ws = XLSX.utils.aoa_to_sheet(data);

    ws['!cols'] = [
      {wch:3}, {wch:6}, {wch:22}, {wch:28}, {wch:20}, {wch:18}, {wch:18},
      {wch:50}, {wch:9}, {wch:14}, {wch:16},
      {wch:14},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14},
      {wch:14}, {wch:18}
    ];

    ws['!merges'] = [
      {s:{r:0,c:2}, e:{r:0,c:7}},
      {s:{r:2,c:11}, e:{r:2,c:21}}
    ];

    XLSX.utils.book_append_sheet(wb, ws, area);
  });

  var obsRecs = STATE.obsoletos || [];
  if (obsRecs.length > 0 || true) {
    var obsRows = [];
    obsRows.push(['', '', 'LISTADO MAESTRO DE INFORMACION DOCUMENTADA', '', '', '', '', '', 'Código: L1/AC-P01', 'Versión: 009/Nov-25']);
    obsRows.push([]);
    var obsH3 = ['','','','','','','','','','',''];
    obsH3[11] = 'COPIA CONTROLADAS / RESPONSABLE';
    obsRows.push(obsH3);
    var obsHeaders = ['', 'No.', 'Clasificación', 'Código del Manual / Procedimiento', 'Código del Instructivo', 'Código del Listado', 'Código del Formato', 'Nombre del Documento', 'Versión', 'Fecha Emisión', 'Elaboró / Actualizó', 'Motivo'];
    obsRows.push(obsHeaders);
    obsRecs.forEach(function(o, i) {
      obsRows.push(['', String(i+1), o.area || '', o.codManual || '', o.codInstructivo || '', o.codListado || '', o.codFormato || '', o.nombreDoc || '', o.version || '', o.fechaEmision || '', o.elaboro || '', o.motivo || '']);
    });
    var obsWs = XLSX.utils.aoa_to_sheet(obsRows);
    obsWs['!cols'] = [{wch:3},{wch:6},{wch:22},{wch:28},{wch:20},{wch:18},{wch:18},{wch:50},{wch:9},{wch:14},{wch:16},{wch:40}];
    obsWs['!merges'] = [{s:{r:0,c:2},e:{r:0,c:7}}];
    XLSX.utils.book_append_sheet(wb, obsWs, 'Obsoletos');
  }

  var now = new Date();
  var fname = 'DEBBIOM_Backup_' + now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0') + '-' + String(now.getDate()).padStart(2,'0') + '.xlsx';
  XLSX.writeFile(wb, fname);
  showToast('Backup Excel exportado: ' + fname);
  addLog('export', 'Backup Excel exportado: ' + fname);
}

function renderMainContent() {
  if (STATE.currentSection === 'dashboard') {
    return _renderDashboard();
  }
  if (STATE.currentSection === 'alertas') {
    return renderAlertasPanel();
  }
  if (STATE.currentSection === 'backups') {
    return renderCloudBackupsSection();
  }
  if (STATE.currentSection === 'logs') {
    return _renderLogsSection();
  }
  if (STATE.currentSection === 'obsoletos') {
    return _renderObsoletosSection();
  }
  if (STATE.currentSection === 'salidas') {
    return _renderSalidasSection();
  }
  if (STATE.currentSection === 'papelera') {
    return _renderPapeleraSection();
  }
  return _renderRecordsSection();
}

function _renderLogsSection() {
  var logs = STATE.logs || [];
  return h('div', {className: 'content-area'},
    h('div', {className: 'topbar'},
      h('span', {className:'topbar-title'}, h('i', {className: 'fas fa-list-check', style:{marginRight:'8px'}}), 'Bitácora de Cambios'),
      h('div', {className: 'topbar-actions'},
        h('button', {className: 'btn btn-sm btn-secondary', onClick: exportLogsToExcel},
          h('i', {className: 'fas fa-file-excel'}), ' Exportar'
        )
      )
    ),
    logs.length === 0
      ? h('div', {style:{textAlign:'center',padding:'60px 20px',color:'var(--text-muted)'}},
          h('i', {className:'fas fa-clipboard-list', style:{fontSize:'48px',marginBottom:'16px',opacity:'.3'}}),
          h('div', null, 'No hay registros en la bitácora')
        )
      : h('div', {className:'table-container'},
          h('table', null,
            h('thead', null,
              h('tr', null,
                h('th', null, 'Fecha'),
                h('th', null, 'Tipo'),
                h('th', null, 'Detalle'),
                h('th', null, 'Documento'),
                h('th', null, 'Usuario')
              )
            ),
            h('tbody', null,
              ...logs.slice(0, 200).map(function(l) {
                var icon = (typeof LOG_TYPES !== 'undefined' && LOG_TYPES[l.type]) || {icon:'fa-circle-info',color:'#888',label:l.type};
                return h('tr', null,
                  h('td', null, formatLogDate(l.timestamp || l.date || l.ts)),
                  h('td', null, h('span', {style:{color:icon.color}}, h('i',{className:'fas ' + icon.icon, style:{marginRight:'4px'}}), icon.label)),
                  h('td', null, l.details || ''),
                  h('td', null, l.docName || ''),
                  h('td', null, l.user || '')
                );
              })
            )
          )
        )
  );
}

function _renderObsoletosSection() {
  var items = STATE.obsoletos || [];
  return h('div', {className: 'content-area'},
    h('div', {className: 'topbar'},
      h('span', {className:'topbar-title'}, h('i', {className: 'fas fa-box-archive', style:{marginRight:'8px',color:'var(--orange)'}}), 'Documentos Obsoletos'),
      h('span', {className: 'badge', style:{marginLeft:'8px'}}, String(items.length))
    ),
    items.length === 0
      ? h('div', {style:{textAlign:'center',padding:'60px 20px',color:'var(--text-muted)'}},
          h('i', {className:'fas fa-box-open', style:{fontSize:'48px',marginBottom:'16px',opacity:'.3'}}),
          h('div', null, 'No hay documentos obsoletos')
        )
      : h('div', {className:'table-container'},
          h('table', null,
            h('thead', null, h('tr', null, h('th',null,'Nombre'), h('th',null,'Código'), h('th',null,'Versión'), h('th',null,'Emisión'), h('th',null,'Motivo'))),
            h('tbody', null, ...items.map(function(o) {
              return h('tr', null,
                h('td', null, o.nombreDoc || '—'),
                h('td', null, o.codManual || o.codInstructivo || o.codFormato || '—'),
                h('td', null, o.version || '—'),
                h('td', null, o.fechaEmision || '—'),
                h('td', null, o.motivo || '—')
              );
            }))
          )
        )
  );
}

function _renderSalidasSection() {
  var items = STATE.salidas || [];
  return h('div', {className: 'content-area'},
    h('div', {className: 'topbar'},
      h('span', {className:'topbar-title'}, h('i', {className: 'fas fa-sign-out-alt', style:{marginRight:'8px',color:'var(--purple)'}}), 'Registros de Salida'),
      h('span', {className: 'badge', style:{marginLeft:'8px'}}, String(items.length))
    ),
    items.length === 0
      ? h('div', {style:{textAlign:'center',padding:'60px 20px',color:'var(--text-muted)'}},
          h('i', {className:'fas fa-inbox', style:{fontSize:'48px',marginBottom:'16px',opacity:'.3'}}),
          h('div', null, 'No hay registros de salida')
        )
      : h('div', {className:'table-container'},
          h('table', null,
            h('thead', null, h('tr', null, h('th',null,'Fecha'), h('th',null,'Documento'), h('th',null,'Destinatario'), h('th',null,'Área'), h('th',null,'Motivo'))),
            h('tbody', null, ...items.map(function(s) {
              return h('tr', null,
                h('td', null, formatLogDate(s.fecha)),
                h('td', null, s.nombreDoc || '—'),
                h('td', null, s.destinatario || '—'),
                h('td', null, s.area || '—'),
                h('td', null, s.motivo || '—')
              );
            }))
          )
        )
  );
}

function _renderPapeleraSection() {
  var items = STATE.papelera || [];
  return h('div', {className: 'content-area'},
    h('div', {className: 'topbar'},
      h('span', {className:'topbar-title'}, h('i', {className: 'fas fa-trash-can', style:{marginRight:'8px',color:'var(--red)'}}), 'Papelera'),
      h('span', {className: 'badge', style:{marginLeft:'8px'}}, String(items.length))
    ),
    items.length === 0
      ? h('div', {style:{textAlign:'center',padding:'60px 20px',color:'var(--text-muted)'}},
          h('i', {className:'fas fa-trash', style:{fontSize:'48px',marginBottom:'16px',opacity:'.3'}}),
          h('div', null, 'La papelera está vacía')
        )
      : h('div', {className:'table-container'},
          h('table', null,
            h('thead', null, h('tr', null, h('th',null,'Nombre'), h('th',null,'Código'), h('th',null,'Área'), h('th',null,'Acciones'))),
            h('tbody', null, ...items.map(function(r) {
              return h('tr', null,
                h('td', null, r.nombreDoc || '—'),
                h('td', null, r.codManual || r.codInstructivo || r.codFormato || '—'),
                h('td', null, r.area || '—'),
                h('td', null,
                  h('button', {className:'btn btn-sm btn-primary', style:{marginRight:'4px'}, onClick:function(){restoreRecord(r.id);}}, h('i',{className:'fas fa-undo'}), ' Restaurar'),
                  isAdmin() ? h('button', {className:'btn btn-sm btn-danger', onClick:function(){if(confirm('¿Eliminar permanentemente?'))permanentDelete(r.id);}}, h('i',{className:'fas fa-trash'})) : null
                )
              );
            }))
          )
        )
  );
}

function _renderRecordsSection() {
  var recs = getFilteredRecords();
  var totalRecs = recs.length;
  var start = (STATE.page - 1) * STATE.perPage;
  var pageRecs = recs.slice(start, start + STATE.perPage);
  var totalPages = Math.ceil(totalRecs / STATE.perPage) || 1;
  var vencidos = typeof getAlertasResumen === 'function' ? getAlertasResumen() : {vencidos:0,criticos:0};

  return h('div', {className: 'content-area'},
    h('div', {className: 'topbar'},
      h('div', {style:{display:'flex',alignItems:'center',gap:'12px',flex:1}},
        h('span', {className:'topbar-title'}, STATE.currentSection === 'areas' ? STATE.currentArea : 'Documentos'),
        h('span', {className: 'badge'}, String(totalRecs)),
        h('input', {
          className: 'form-input',
          type: 'search',
          id: 'search-docs-input',
          placeholder: 'Buscar documentos...',
          style:{maxWidth:'260px',fontSize:'13px'},
          value: STATE.searchQuery || '',
          onInput: function(e) {
            STATE.searchQuery = e.target.value;
            STATE.page = 1;
            clearTimeout(window._searchDebounce);
            window._searchDebounce = setTimeout(function() {
              // Re-render solo la tabla, no el input
              var tableContainer = document.querySelector('.table-container');
              var emptyMsg = document.querySelector('.content-area > div[style*="text-align:center"]');
              if (tableContainer || emptyMsg) {
                var recs = getFilteredRecords();
                var badge = document.querySelector('.topbar .badge');
                if (badge) badge.textContent = String(recs.length);
                render();
                // Restaurar foco al input
                setTimeout(function() {
                  var inp = document.getElementById('search-docs-input');
                  if (inp) { inp.focus(); inp.selectionStart = inp.selectionEnd = inp.value.length; }
                }, 10);
              }
            }, 300);
          }
        })
      ),
      h('div', {className: 'topbar-actions'},
        h('select', {
          className: 'form-input',
          style:{width:'auto',fontSize:'12px'},
          value: STATE.filterEstado || '',
          onChange: function(e) { STATE.filterEstado = e.target.value; STATE.page = 1; render(); }
        },
          h('option', {value: ''}, 'Todos los estados'),
          h('option', {value: 'Vigente'}, 'Vigente'),
          h('option', {value: 'Por vencer'}, 'Por vencer'),
          h('option', {value: 'Vencido'}, 'Vencido')
        ),
        h('select', {
          className: 'form-input',
          style:{width:'auto',fontSize:'12px'},
          value: STATE.filterElaboro || '',
          onChange: function(e) { STATE.filterElaboro = e.target.value; STATE.page = 1; render(); }
        },
          h('option', {value: ''}, 'Todos (Elaboró)'),
          ...STATE.elaboros.map(function(el) { return h('option', {value: el}, el); })
        ),
        vencidos.vencidos > 0 ? h('button', {className:'btn btn-sm btn-danger', onClick:function(){exportVencidosToExcel();}},
          h('i',{className:'fas fa-exclamation-triangle'}), ' ' + vencidos.vencidos + ' Vencidos'
        ) : null,
        isAdmin() ? h('button', {className:'btn btn-sm btn-secondary', onClick:_openImportDialog},
          h('i',{className:'fas fa-file-import'}), ' Importar'
        ) : null,
        isAdmin() ? h('button', {className:'btn btn-sm btn-primary', onClick:function(){
          STATE.modal = 'add';
          STATE.selectedRecord = null;
          STATE.editForm = { area: STATE.currentArea || AREAS[0] };
          STATE.originalVersion = null;
          refreshModal();
        }}, h('i',{className:'fas fa-plus'}), ' Nuevo') : null
      )
    ),
    pageRecs.length === 0
      ? h('div', {style:{textAlign:'center',padding:'60px 20px',color:'var(--text-muted)'}},
          h('i', {className:'fas fa-folder-open', style:{fontSize:'48px',marginBottom:'16px',opacity:'.3'}}),
          h('div', null, STATE.searchQuery ? 'No se encontraron resultados' : 'No hay documentos en esta área')
        )
      : h('div', {className:'table-container'},
          h('table', null,
            h('thead', null,
              h('tr', null,
                h('th', {onClick:function(){_sortBy('no');}, style:{cursor:'pointer'}}, 'No.', _renderSortIcon('no')),
                h('th', {onClick:function(){_sortBy('codManual');}, style:{cursor:'pointer'}}, 'Código', _renderSortIcon('codManual')),
                h('th', {onClick:function(){_sortBy('nombreDoc');}, style:{cursor:'pointer'}}, 'Nombre del Documento', _renderSortIcon('nombreDoc')),
                h('th', null, 'Versión'),
                h('th', {onClick:function(){_sortBy('fechaEmision');}, style:{cursor:'pointer'}}, 'Emisión', _renderSortIcon('fechaEmision')),
                h('th', {onClick:function(){_sortBy('dias');}, style:{cursor:'pointer'}}, 'Estado', _renderSortIcon('dias')),
                h('th', {onClick:function(){_sortBy('elaboro');}, style:{cursor:'pointer'}}, 'Elaboró', _renderSortIcon('elaboro')),
                h('th', null, 'Acciones')
              )
            ),
            h('tbody', null,
              ...pageRecs.map(function(r) {
                var estado = typeof calcEstado === 'function' ? calcEstado(r.fechaEmision) : '';
                var diasV = typeof calcDiasVigencia === 'function' ? calcDiasVigencia(r.fechaEmision) : null;
                var estadoClass = estado === 'Vigente' ? 'badge-green' : estado === 'Por vencer' ? 'badge-orange' : estado === 'Vencido' ? 'badge-red' : '';
                var code = r.codManual || r.codInstructivo || r.codListado || r.codFormato || '';
                return h('tr', null,
                  h('td', null, r.no || ''),
                  h('td', null, h('span', {style:{fontFamily:'JetBrains Mono, monospace',fontSize:'12px',color:'var(--accent)'}}, code)),
                  h('td', null, r.nombreDoc || '—'),
                  h('td', null, r.version || ''),
                  h('td', null, r.fechaEmision || ''),
                  h('td', null,
                    estado ? h('span', {className:'badge ' + estadoClass, style:{fontSize:'11px'}}, estado) : '',
                    diasV !== null && estado === 'Vencido' ? h('span', {style:{fontSize:'10px',color:'var(--red)',marginLeft:'4px'}}, Math.abs(diasV) + 'd') : null
                  ),
                  h('td', null, r.elaboro || ''),
                  h('td', null,
                    h('div', {style:{display:'flex',gap:'4px'}},
                      h('button', {className:'btn btn-sm', title:'Ver', onClick:function(){
                        STATE.selectedRecord = r;
                        STATE.modal = 'view';
                        refreshModal();
                      }}, h('i',{className:'fas fa-eye'})),
                      isAdmin() ? h('button', {className:'btn btn-sm', title:'Editar', onClick:function(){
                        STATE.selectedRecord = r;
                        STATE.modal = 'edit';
                        STATE.editForm = Object.assign({}, r);
                        STATE.originalVersion = r.version || '';
                        refreshModal();
                      }}, h('i',{className:'fas fa-pen'})) : null,
                      isAdmin() ? h('button', {className:'btn btn-sm', title:'Salida', onClick:function(){
                        STATE.selectedRecord = r;
                        STATE.modal = 'salida';
                        STATE.salidaForm = {};
                        refreshModal();
                      }}, h('i',{className:'fas fa-sign-out-alt'})) : null,
                      isAdmin() && estado === 'Vencido' ? h('button', {className:'btn btn-sm', title:'Renovar', onClick:function(){
                        STATE.selectedRecord = r;
                        STATE.modal = 'renew';
                        STATE._renewDate = '';
                        STATE._renewMotivo = '';
                        refreshModal();
                      }}, h('i',{className:'fas fa-rotate-right', style:{color:'var(--green)'}})) : null,
                      isAdmin() ? h('button', {className:'btn btn-sm btn-danger', title:'Eliminar', onClick:function(){
                        if (confirm('¿Mover este documento a papelera?')) deleteRecord(r.id);
                      }}, h('i',{className:'fas fa-trash-can'})) : null
                    )
                  )
                );
              })
            )
          )
        ),
    totalPages > 1 ? h('div', {style:{display:'flex',justifyContent:'center',alignItems:'center',gap:'8px',marginTop:'16px',fontSize:'13px'}},
      h('button', {className:'btn btn-sm', disabled: STATE.page <= 1, onClick:function(){STATE.page--;render();}}, '← Anterior'),
      h('span', null, 'Página ' + STATE.page + ' de ' + totalPages),
      h('button', {className:'btn btn-sm', disabled: STATE.page >= totalPages, onClick:function(){STATE.page++;render();}}, 'Siguiente →')
    ) : null
  );
}

function render() {
  var app = document.getElementById('app');
  if (!app) return;

  if (typeof AUTH === 'undefined' || !AUTH.currentUser) {
    renderLogin();
    return;
  }

  app.innerHTML = '';

  var sidebar = renderSidebar();
  var mainDiv = document.createElement('div');
  mainDiv.className = 'main';
  var content = renderMainContent();
  mainDiv.appendChild(content);

  app.appendChild(sidebar);
  app.appendChild(mainDiv);

  var modalRoot = document.getElementById('modal-root');
  if (!modalRoot) {
    modalRoot = document.createElement('div');
    modalRoot.id = 'modal-root';
    document.body.appendChild(modalRoot);
  }
  refreshModal();

  if (typeof showAlertBanner === 'function') showAlertBanner();
  if (typeof updateOnlineCountUI === 'function') updateOnlineCountUI();
  if (typeof updateSyncIndicator === 'function') updateSyncIndicator();
  if (typeof _updateLastSyncUI === 'function') _updateLastSyncUI();
}

document.addEventListener('DOMContentLoaded', async function() {
  loadExcelData();
  loadSavedState();

  if (typeof initAuth === 'function') {
    await initAuth();
  }

  render();

  if (typeof initSync === 'function') {
    try { await initSync(); } catch(e) {}
  }

  if (typeof _startAutoBackup === 'function') _startAutoBackup();
  if (typeof registerPresence === 'function') registerPresence();

  // Auto-refresh dashboard every 30 seconds
  setInterval(function() {
    if (STATE.currentSection === 'dashboard' && AUTH.currentUser && !STATE.modal) {
      render();
    }
  }, 30000);
});
