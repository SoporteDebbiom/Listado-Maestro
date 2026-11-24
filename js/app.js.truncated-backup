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
      h('img', {className:'sidebar-logo', src:'data:image/jpeg;base64,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
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