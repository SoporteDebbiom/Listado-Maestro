const AREAS = ['Dir. Gral.','Calidad','Clínica','Analítica','Bioestad.','Lab. Clínico','Seg&Hig','Mtto','Administración','TI Ex','RH Ex'];

const DEBBIOM_LOGO = '';

function formatDateFull(d) {
  if (!d) return '';
  var date = (d instanceof Date) ? d : new Date(d);
  if (isNaN(date.getTime())) return '';
  var dd = String(date.getDate()).padStart(2, '0');
  var mm = String(date.getMonth() + 1).padStart(2, '0');
  var yyyy = date.getFullYear();
  return dd + '/' + mm + '/' + yyyy;
}

function calcVigencia(fechaEmision) {
  if (!fechaEmision) return null;
  var parts = fechaEmision.split('-');
  if (parts.length !== 3) return null;
  var fecha = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  if (isNaN(fecha.getTime())) return null;
  fecha.setFullYear(fecha.getFullYear() + 2);
  return fecha;
}

function calcEstado(fechaEmision) {
  var vigencia = calcVigencia(fechaEmision);
  if (!vigencia) return '';
  var hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  var diff = vigencia.getTime() - hoy.getTime();
  var dias = Math.ceil(diff / (1000 * 60 * 60 * 24));
  if (dias < 0) return 'Vencido';
  if (dias <= 90) return 'Por vencer';
  return 'Vigente';
}

function calcDiasVigencia(fechaEmision) {
  var vigencia = calcVigencia(fechaEmision);
  if (!vigencia) return null;
  var hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  var diff = vigencia.getTime() - hoy.getTime();
  return Math.ceil(diff / (1000 * 60 * 60 * 24));
}

function renderVigenciaDate(fechaEmision) {
  var vigencia = calcVigencia(fechaEmision);
  return vigencia ? formatDateFull(vigencia) : '';
}

function getAlertasVigencia() {
  var records = (typeof STATE !== 'undefined' && STATE.records) ? STATE.records : [];
  var alertas = [];
  records.forEach(function(r) {
    var dias = calcDiasVigencia(r.fechaEmision);
    if (dias === null) return;
    var nivel = '';
    if (dias < 0) nivel = 'vencido';
    else if (dias <= 30) nivel = 'critico';
    else if (dias <= 90) nivel = 'proximo';
    else return;
    alertas.push({
      nombreDoc: r.nombreDoc || r.codManual || r.codFormato || '',
      codigo: r.codManual || r.codInstructivo || r.codFormato || '',
      area: r.area || '',
      elaboro: r.elaboro || '',
      fechaEmision: r.fechaEmision || '',
      vigenciaDate: calcVigencia(r.fechaEmision),
      dias: dias,
      nivel: nivel
    });
  });
  alertas.sort(function(a, b) { return a.dias - b.dias; });
  return alertas;
}

function getAlertasResumen() {
  var alertas = getAlertasVigencia();
  return {
    total: alertas.length,
    vencidos: alertas.filter(function(a) { return a.nivel === 'vencido'; }).length,
    criticos: alertas.filter(function(a) { return a.nivel === 'critico'; }).length,
    proximos: alertas.filter(function(a) { return a.nivel === 'proximo'; }).length,
    alertas: alertas
  };
}

function getAlertasCriticas() {
  var alertas = getAlertasVigencia();
  return alertas.filter(function(a) { return a.nivel === 'vencido' || a.nivel === 'critico'; });
}
