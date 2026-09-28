/**
 * SISTEMA PROFESIONAL DE GESTIÓN DE REEMPLAZOS, LICENCIAS Y PERMISOS
 * Escuela J. Fco. Pino M. - Coelemu
 * Backend Principal - Google Apps Script (Modelo 3NF Relacional con Blindaje Inteligente)
 * Base de Datos ID: 1IIPorweYxVVhfR2m3-rTmsUBn_h6d3pfV1iVahLRLV0
 * Versión: 3.0 (Consolidación y Auditoría Integral)
 */

const CONFIG = {
  DB_NAME: 'BD_Sistema_Reemplazos_Docentes',
  DB_ID: '1IIPorweYxVVhfR2m3-rTmsUBn_h6d3pfV1iVahLRLV0',
  CACHE_TTL: 21600,
  CREADOR_MASTER: 'asistenciatecnicafranciscopino@gmail.com'
};

const SHEETS = {
  ROLES: 'Roles',
  USUARIOS: 'Usuarios',
  CARGOS: 'Cargos',
  PERSONAL: 'Personal',
  CURSOS: 'Cursos',
  ASIGNATURAS: 'Asignaturas',
  CURSOS_ASIGNATURAS: 'CursosAsignaturas',
  BLOQUES: 'Bloques',
  HORAS_PEDAGOGICAS: 'HorasPedagogicas',
  HORARIOS: 'Horarios',
  TIPOS_LICENCIA: 'TiposLicencia',
  ENTIDADES_TRAMITADORAS: 'EntidadesTramitadoras',
  TIPOS_JORNADA: 'TiposJornada',
  LICENCIAS_MEDICAS: 'LicenciasMedicas',
  PERMISOS_ADMINISTRATIVOS: 'PermisosAdministrativos',
  ESTADOS_REEMPLAZO: 'EstadosReemplazo',
  REEMPLAZOS: 'Reemplazos',
  TIPOS_NOTIFICACION: 'TiposNotificacion',
  NOTIFICACIONES: 'Notificaciones',
  CONFIGURACION: 'Configuracion',
  AUDITORIA: 'Auditoria'
};

const PREFIXES = {
  ROLES: 'ROL-',
  USUARIOS: 'USR-',
  CARGOS: 'CAR-',
  PERSONAL: 'PER-',
  CURSOS: 'CUR-',
  ASIGNATURAS: 'ASI-',
  CURSOS_ASIGNATURAS: 'CAS-',
  BLOQUES: 'BLO-',
  HORAS_PEDAGOGICAS: 'HPE-',
  HORARIOS: 'HOR-',
  TIPOS_LICENCIA: 'TLI-',
  ENTIDADES_TRAMITADORAS: 'ENT-',
  TIPOS_JORNADA: 'TJO-',
  LICENCIAS_MEDICAS: 'LIC-',
  PERMISOS_ADMINISTRATIVOS: 'PAD-',
  ESTADOS_REEMPLAZO: 'ERE-',
  REEMPLAZOS: 'REP-',
  TIPOS_NOTIFICACION: 'TNO-',
  NOTIFICACIONES: 'NOT-',
  CONFIGURACION: 'CFG-',
  AUDITORIA: 'AUD-'
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistema de Gestión - Escuela J. Fco. Pino M.')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function formatearFechaLimpia(fecha) {
  if (!fecha) return "";
  let d = new Date(fecha);
  if (isNaN(d.getTime())) return String(fecha).split(' ')[0]; 
  let anio = d.getFullYear();
  let mes = String(d.getMonth() + 1).padStart(2, '0');
  let dia = String(d.getDate()).padStart(2, '0');
  return `${anio}-${mes}-${dia}`;
}

function getDbSpreadsheet() {
  try {
    return SpreadsheetApp.openById(CONFIG.DB_ID);
  } catch (err) {
    const files = DriveApp.getFilesByName(CONFIG.DB_NAME);
    if (files.hasNext()) {
      return SpreadsheetApp.open(files.next());
    }
    throw new Error('No se pudo acceder a la base de datos oficial. ID: ' + CONFIG.DB_ID);
  }
}

function apiGetSesionUsuario() {
  const emailAcceso = Session.getActiveUser().getEmail() || '';
  const emailLimpio = emailAcceso.trim().toLowerCase();
  const esMaster = (emailLimpio === CONFIG.CREADOR_MASTER.toLowerCase());

  const personalList = getTableData(SHEETS.PERSONAL);
  const rolesList = getTableData(SHEETS.ROLES);
  
  const funcionario = personalList.find(p => String(p.Email || '').trim().toLowerCase() === emailLimpio && (p.Estado === 'ACTIVO' || esMaster));

  if (!funcionario && !esMaster) {
    throw new Error('ACCESO DENEGADO: Su cuenta (' + emailLimpio + ') no está registrada o autorizada en el sistema.');
  }

  let nombreCompleto = 'Funcionario Autorizado';
  let idRolAsignado = esMaster ? 'ROL-ADMIN' : 'ROL-DOCENTE';
  let nombreRol = esMaster ? 'Administrador' : 'Docente / Funcionario';

  if (funcionario) {
    nombreCompleto = funcionario.Nombre_Completo || `${funcionario.Nombres || ''} ${funcionario.Apellido_Paterno || ''}`.trim();
    if (funcionario.ID_Rol) idRolAsignado = funcionario.ID_Rol;
  } else if (esMaster) {
    nombreCompleto = 'Equipo Directivo (Master)';
  }

  const rolObj = rolesList.find(r => r.ID_Rol === idRolAsignado);
  if (rolObj) nombreRol = rolObj.Nombre_Rol;

  return {
    success: true,
    email: emailLimpio,
    nombreCompleto: nombreCompleto,
    rol: nombreRol,
    idRol: idRolAsignado,
    accesoTotal: (esMaster || idRolAsignado === 'ROL-ADMIN' || idRolAsignado === 'ROL-INSPECTORIA')
  };
}

function generateNextId(sheetName, prefix) {
  const ss = getDbSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return prefix + '0001';
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return prefix + '0001';
  
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const idStr = String(data[i][0] || '');
    if (idStr.startsWith(prefix)) {
      const num = parseInt(idStr.replace(prefix, ''), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }
  return prefix + String(maxNum + 1).padStart(4, '0');
}

function logAuditoria(accion, modulo, idRegistro, valorAnterior, valorNuevo) {
  try {
    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.AUDITORIA);
    if (!sheet) return;
    
    const idAudit = generateNextId(SHEETS.AUDITORIA, PREFIXES.AUDITORIA);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const usuario = Session.getActiveUser().getEmail() || 'SISTEMA_LOCAL';
    
    sheet.appendRow([
      idAudit, timestamp, usuario, accion, modulo, idRegistro,
      typeof valorAnterior === 'object' ? JSON.stringify(valorAnterior) : String(valorAnterior || ''),
      typeof valorNuevo === 'object' ? JSON.stringify(valorNuevo) : String(valorNuevo || '')
    ]);
  } catch (err) {
    Logger.log('Error de auditoría: ' + err.toString());
  }
}

function getTableData(sheetName) {
  const ss = getDbSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      }
      obj[h] = val;
    });
    return obj;
  });
}

function verificarRequiereReemplazo(idPersonal) {
  const personalList = getTableData(SHEETS.PERSONAL);
  const cargosList = getTableData(SHEETS.CARGOS);
  
  const funcionario = personalList.find(p => p.ID_Personal === idPersonal);
  if (!funcionario) return false;

  const cargo = cargosList.find(c => c.ID_Cargo === funcionario.ID_Cargo);
  if (!cargo) return false;

  const permite = cargo.Permite_Reemplazo_Pedagogico;
  return permite === true || permite === "TRUE" || String(permite).toUpperCase() === "SÍ" || String(permite).toUpperCase() === "SI";
}

function apiGetDashboardStats() {
  const personal = getTableData(SHEETS.PERSONAL).filter(p => p.Estado === 'ACTIVO');
  const licencias = getTableData(SHEETS.LICENCIAS_MEDICAS).filter(l => l.Estado === 'ACTIVO');
  const permisos = getTableData(SHEETS.PERMISOS_ADMINISTRATIVOS).filter(p => p.Estado === 'APROBADO' || p.Estado === 'ACTIVO');
  const reemplazos = getTableData(SHEETS.REEMPLAZOS);
  
  const pendientes = reemplazos.filter(r => r.ID_EstadoReemplazo === 'ERE-0001').length;
  const asignados = reemplazos.filter(r => ['ERE-0002', 'ERE-0003'].includes(r.ID_EstadoReemplazo)).length;

  return {
    totalPersonal: personal.length,
    totalAusencias: licencias.length + permisos.length,
    reemplazosPendientes: pendientes,
    reemplazosAsignados: asignados
  };
}

function apiGetCargos() { return getTableData(SHEETS.CARGOS).filter(c => c.Estado === 'ACTIVO'); }
function apiGetTiposLicencia() { return getTableData(SHEETS.TIPOS_LICENCIA).filter(t => t.Estado === 'ACTIVO'); }
function apiGetEntidadesTramitadoras() { return getTableData(SHEETS.ENTIDADES_TRAMITADORAS).filter(e => e.Estado === 'ACTIVO'); }
function apiGetTiposJornada() { return getTableData(SHEETS.TIPOS_JORNADA).filter(j => j.Estado === 'ACTIVO'); }

function apiGetCursos() { 
  const ss = getDbSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CURSOS);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let cursos = [];
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let idCurso = String(row[0] || '').trim();
    let nombreCurso = String(row[1] || '').trim();
    let estado = String(row[3] || 'ACTIVO').trim().toUpperCase();

    if (idCurso && (estado === 'ACTIVO' || estado === '')) {
      cursos.push({
        id: idCurso,
        nombre: nombreCurso || idCurso,
        ID_Curso: idCurso,
        Nombre_Curso: nombreCurso || idCurso,
        Estado: estado || 'ACTIVO'
      });
    }
  }
  return cursos;
}

function apiGetAsignaturas() { 
  const asignaturas = getTableData(SHEETS.ASIGNATURAS);
  if (!asignaturas || asignaturas.length === 0) return [];

  return asignaturas
    .filter(a => !a.Estado || String(a.Estado).toUpperCase() === 'ACTIVO')
    .map(a => ({
      id: a.ID_Asignatura,
      nombre: a.Nombre_Asignatura,
      ID_Asignatura: a.ID_Asignatura,
      Nombre_Asignatura: a.Nombre_Asignatura,
      Area_Conocimiento: a.Area_Conocimiento,
      Estado: a.Estado || 'ACTIVO'
    }));
}

function apiGetAsignaturasPorCurso(idCurso) {
  if (!idCurso) return apiGetAsignaturas();
  
  const ss = getDbSpreadsheet();
  const relSheet = ss.getSheetByName(SHEETS.CURSOS_ASIGNATURAS);
  if (!relSheet) return apiGetAsignaturas();

  const relData = relSheet.getDataRange().getValues();
  if (relData.length <= 1) return apiGetAsignaturas();

  const relHeaders = relData[0];
  const idxCurso = relHeaders.indexOf('ID_Curso') !== -1 ? relHeaders.indexOf('ID_Curso') : 0;
  const idxAsig = relHeaders.indexOf('ID_Asignatura') !== -1 ? relHeaders.indexOf('ID_Asignatura') : 1;
  const idxEstado = relHeaders.indexOf('Estado') !== -1 ? relHeaders.indexOf('Estado') : 2;

  let idsAsignaturasPermitidas = [];
  for (let i = 1; i < relData.length; i++) {
    let row = relData[i];
    let cursoVal = String(row[idxCurso] || '').trim();
    let asigVal = String(row[idxAsig] || '').trim();
    let estadoVal = String(row[idxEstado] || 'ACTIVO').trim().toUpperCase();

    if ((cursoVal === String(idCurso).trim()) && (estadoVal === 'ACTIVO' || estadoVal === '')) {
      if (asigVal) idsAsignaturasPermitidas.push(asigVal);
    }
  }

  const todasAsignaturas = apiGetAsignaturas();
  if (idsAsignaturasPermitidas.length === 0) return todasAsignaturas;

  return todasAsignaturas.filter(a => idsAsignaturasPermitidas.includes(String(a.ID_Asignatura || a.id).trim()));
}

function apiGetPersonal() {
  const personalList = getTableData(SHEETS.PERSONAL);
  const cargosList = getTableData(SHEETS.CARGOS);
  
  const mapCargo = {};
  cargosList.forEach(c => mapCargo[c.ID_Cargo] = c.Nombre_Cargo);

  return personalList.map(p => ({
    ...p,
    Nombre_Completo: p.Nombre_Completo || `${p.Nombres || ''} ${p.Apellido_Paterno || ''}`.trim(),
    Nombre_Cargo: mapCargo[p.ID_Cargo] || p.ID_Cargo
  }));
}

function apiGetPersonalDocente() {
  const ss = getDbSpreadsheet();
  const cargosSheet = ss.getSheetByName(SHEETS.CARGOS);
  if (!cargosSheet) return [];
  const cargosData = cargosSheet.getDataRange().getValues();
  const cargosHeaders = cargosData[0];
  
  const idxCargoId = cargosHeaders.indexOf('ID_Cargo');
  const idxPermiteReemplazo = cargosHeaders.indexOf('Permite_Reemplazo_Pedagogico');

  let cargosDocentesIds = [];
  for (let i = 1; i < cargosData.length; i++) {
    let row = cargosData[i];
    let idCargo = row[idxCargoId];
    let permiteReemplazo = row[idxPermiteReemplazo];
    
    if (permiteReemplazo === true || permiteReemplazo === "TRUE" || String(permiteReemplazo).toUpperCase() === "SÍ" || String(permiteReemplazo).toUpperCase() === "SI") {
      cargosDocentesIds.push(idCargo);
    }
  }

  const personalSheet = ss.getSheetByName(SHEETS.PERSONAL);
  if (!personalSheet) return [];
  const personalData = personalSheet.getDataRange().getValues();
  if (personalData.length <= 1) return [];

  const perHeaders = personalData[0];
  const idxIdProf = perHeaders.indexOf('ID_Personal') !== -1 ? perHeaders.indexOf('ID_Personal') : perHeaders.indexOf('ID_Profesor');
  const idxNombres = perHeaders.indexOf('Nombres');
  const idxApPaterno = perHeaders.indexOf('Apellido_Paterno');
  const idxNomComp = perHeaders.indexOf('Nombre_Completo');
  const idxCargo = perHeaders.indexOf('ID_Cargo');
  const idxEstado = perHeaders.indexOf('Estado');

  let docentes = [];
  for (let i = 1; i < personalData.length; i++) {
    let row = personalData[i];
    let idPersonal = row[idxIdProf];
    let idCargo = row[idxCargo];
    let estado = String(row[idxEstado] || '').trim().toUpperCase();
    
    if (estado === "ACTIVO" && cargosDocentesIds.includes(idCargo)) {
      let nombres = row[idxNombres] || '';
      let apellidoPaterno = row[idxApPaterno] || '';
      let nomComp = (idxNomComp !== -1 ? row[idxNomComp] : '') || `${nombres} ${apellidoPaterno}`.trim();
      
      docentes.push({
        ID_Personal: idPersonal,
        Nombre_Completo: nomComp,
        id: idPersonal,
        nombreCompleto: nomComp
      });
    }
  }
  return docentes;
}

function apiGetLicenciasMedicas() {
  const licencias = getTableData(SHEETS.LICENCIAS_MEDICAS);
  const personal = getTableData(SHEETS.PERSONAL);
  const tipos = getTableData(SHEETS.TIPOS_LICENCIA);
  const entidades = getTableData(SHEETS.ENTIDADES_TRAMITADORAS);

  const mapPersonal = {}; personal.forEach(p => mapPersonal[p.ID_Personal] = p.Nombre_Completo || `${p.Nombres || ''} ${p.Apellido_Paterno || ''}`.trim());
  const mapTipos = {}; tipos.forEach(t => mapTipos[t.ID_TipoLicencia] = t.Nombre_Tipo);
  const mapEntidades = {}; entidades.forEach(e => mapEntidades[e.ID_Entidad] = e.Nombre_Entidad);

  return licencias.map(l => ({
    ...l,
    Fecha_Inicio: formatearFechaLimpia(l.Fecha_Inicio),
    Fecha_Fin: formatearFechaLimpia(l.Fecha_Fin),
    Nombre_Titular: mapPersonal[l.ID_Personal_Titular] || l.ID_Personal_Titular,
    Nombre_TipoLicencia: mapTipos[l.ID_TipoLicencia] || l.ID_TipoLicencia,
    Nombre_Entidad: mapEntidades[l.ID_Entidad] || l.ID_Entidad
  }));
}

function apiSaveLicenciaMedica(licenciaData) {
  const ss = getDbSpreadsheet();
  const sheetLic = ss.getSheetByName(SHEETS.LICENCIAS_MEDICAS);
  const sheetRep = ss.getSheetByName(SHEETS.REEMPLAZOS);
  
  const idLicencia = licenciaData.ID_Licencia || generateNextId(SHEETS.LICENCIAS_MEDICAS, PREFIXES.LICENCIAS_MEDICAS);
  const isNew = !licenciaData.ID_Licencia;
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  let fechaLimpiaInicio = formatearFechaLimpia(licenciaData.Fecha_Inicio);
  let fechaLimpiaFin = formatearFechaLimpia(licenciaData.Fecha_Fin);

  const start = new Date(fechaLimpiaInicio + 'T00:00:00');
  const end = new Date(fechaLimpiaFin + 'T00:00:00');
  if (end < start) throw new Error("La fecha de término no puede ser menor a la fecha de inicio.");

  const diffTime = Math.abs(end - start);
  const diasTotales = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;

  if (isNew) {
    sheetLic.appendRow([
      idLicencia, licenciaData.ID_Personal_Titular, licenciaData.Folio_Licencia || 'SIN-FOLIO',
      licenciaData.ID_TipoLicencia, licenciaData.ID_Entidad, fechaLimpiaInicio,
      fechaLimpiaFin, diasTotales, licenciaData.Observaciones || '', 'ACTIVO', nowStr
    ]);

    let reemplazosGenerados = 0;
    const requiereCobertura = verificarRequiereReemplazo(licenciaData.ID_Personal_Titular);

    if (requiereCobertura) {
      const horarios = getTableData(SHEETS.HORARIOS).filter(h => (h.ID_Profesor === licenciaData.ID_Personal_Titular) && (h.Estado === 'ACTIVO'));
      let current = new Date(start);
      const diasSemana = ['DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO'];

      while (current <= end) {
        const diaNombre = diasSemana[current.getDay()];
        const fechaStr = formatearFechaLimpia(current);

        if (diaNombre !== 'SABADO' && diaNombre !== 'DOMINGO') {
          const horariosDia = horarios.filter(h => h.Dia_Semana === diaNombre);
          horariosDia.forEach(h => {
            const idReemplazo = generateNextId(SHEETS.REEMPLAZOS, PREFIXES.REEMPLAZOS);
            sheetRep.appendRow([
              idReemplazo, fechaStr, fechaStr, licenciaData.ID_Personal_Titular, '',
              h.ID_Curso, h.ID_Asignatura, h.ID_Bloque, '', 'ERE-0001',
              'Generado por Licencia Médica ' + idLicencia, nowStr
            ]);
            reemplazosGenerados++;
          });
        }
        current.setDate(current.getDate() + 1);
      }
    }

    logAuditoria('CREAR', 'LICENCIAS_MEDICAS', idLicencia, null, { licenciaData, reemplazosGenerados, requiereCobertura });
    return { success: true, idLicencia, reemplazosGenerados, requiereCobertura };
  }
}

function apiAnularLicenciaMedica(idLicencia) {
  const ss = getDbSpreadsheet();
  const sheetLic = ss.getSheetByName(SHEETS.LICENCIAS_MEDICAS);
  const data = sheetLic.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idLicencia) {
      sheetLic.getRange(i + 1, 10).setValue('ANULADA');
      
      const sheetRep = ss.getSheetByName(SHEETS.REEMPLAZOS);
      const dataRep = sheetRep.getDataRange().getValues();
      for (let j = 1; j < dataRep.length; j++) {
        if (dataRep[j][10] === ('Generado por Licencia Médica ' + idLicencia) && dataRep[j][9] === 'ERE-0001') {
          sheetRep.getRange(j + 1, 10).setValue('ERE-0006');
        }
      }

      logAuditoria('ANULAR', 'LICENCIAS_MEDICAS', idLicencia, 'ACTIVO', 'ANULADA');
      return { success: true };
    }
  }
  throw new Error('Licencia Médica no encontrada.');
}

function apiGetPermisosAdministrativos() {
  const permisos = getTableData(SHEETS.PERMISOS_ADMINISTRATIVOS);
  const personal = getTableData(SHEETS.PERSONAL);
  const jornadas = getTableData(SHEETS.TIPOS_JORNADA);

  const mapPersonal = {}; personal.forEach(p => mapPersonal[p.ID_Personal] = p.Nombre_Completo || `${p.Nombres || ''} ${p.Apellido_Paterno || ''}`.trim());
  const mapJornadas = {}; jornadas.forEach(j => mapJornadas[j.ID_TipoJornada] = j.Nombre_Jornada);

  return permisos.map(p => ({
    ...p,
    Fecha_Inicio: formatearFechaLimpia(p.Fecha_Inicio),
    Fecha_Fin: formatearFechaLimpia(p.Fecha_Fin),
    Nombre_Titular: mapPersonal[p.ID_Personal_Titular] || p.ID_Personal_Titular,
    Nombre_Jornada: mapJornadas[p.ID_TipoJornada] || p.ID_TipoJornada
  }));
}

function apiSavePermisoAdministrativo(permisoData) {
  const ss = getDbSpreadsheet();
  const sheetPerm = ss.getSheetByName(SHEETS.PERMISOS_ADMINISTRATIVOS);
  const sheetRep = ss.getSheetByName(SHEETS.REEMPLAZOS);

  const idPermiso = permisoData.ID_Permiso || generateNextId(SHEETS.PERMISOS_ADMINISTRATIVOS, PREFIXES.PERMISOS_ADMINISTRATIVOS);
  const isNew = !permisoData.ID_Permiso;
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const usuarioSesion = Session.getActiveUser().getEmail() || 'EQUIPO_DIRECTIVO';

  let fechaLimpiaInicio = formatearFechaLimpia(permisoData.Fecha_Inicio);
  let fechaLimpiaFin = formatearFechaLimpia(permisoData.Fecha_Fin);

  const start = new Date(fechaLimpiaInicio + 'T00:00:00');
  const end = new Date(fechaLimpiaFin + 'T00:00:00');
  if (end < start) throw new Error("La fecha de término no puede ser menor a la fecha de inicio.");

  const jornadas = getTableData(SHEETS.TIPOS_JORNADA);
  const jornadaObj = jornadas.find(j => j.ID_TipoJornada === permisoData.ID_TipoJornada);
  const factor = jornadaObj ? parseFloat(jornadaObj.Factor_Dia) : 1.0;

  const diffTime = Math.abs(end - start);
  const diasCorridos = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
  const diasCalculados = diasCorridos * factor;

  if (isNew) {
    sheetPerm.appendRow([
      idPermiso, permisoData.ID_Personal_Titular, permisoData.ID_TipoJornada || 'TJO-0001',
      fechaLimpiaInicio, fechaLimpiaFin, diasCalculados, permisoData.Con_Goce_Sueldo !== false,
      permisoData.Motivo || '', usuarioSesion, 'APROBADO', nowStr
    ]);

    let reemplazosGenerados = 0;
    const requiereCobertura = verificarRequiereReemplazo(permisoData.ID_Personal_Titular);

    if (requiereCobertura) {
      const horarios = getTableData(SHEETS.HORARIOS).filter(h => h.ID_Profesor === permisoData.ID_Personal_Titular && h.Estado === 'ACTIVO');
      let current = new Date(start);
      const diasSemana = ['DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO'];

      while (current <= end) {
        const diaNombre = diasSemana[current.getDay()];
        const fechaStr = formatearFechaLimpia(current);

        if (diaNombre !== 'SABADO' && diaNombre !== 'DOMINGO') {
          const horariosDia = horarios.filter(h => h.Dia_Semana === diaNombre);
          horariosDia.forEach(h => {
            const idReemplazo = generateNextId(SHEETS.REEMPLAZOS, PREFIXES.REEMPLAZOS);
            sheetRep.appendRow([
              idReemplazo, fechaStr, fechaStr, permisoData.ID_Personal_Titular, '',
              h.ID_Curso, h.ID_Asignatura, h.ID_Bloque, '', 'ERE-0001',
              'Generado por Permiso Admin. ' + idPermiso, nowStr
            ]);
            reemplazosGenerados++;
          });
        }
        current.setDate(current.getDate() + 1);
      }
    }

    logAuditoria('CREAR', 'PERMISOS_ADMINISTRATIVOS', idPermiso, null, { permisoData, diasCalculados, reemplazosGenerados, requiereCobertura });
    return { success: true, idPermiso, reemplazosGenerados, requiereCobertura };
  }
}

function apiAnularPermisoAdministrativo(idPermiso) {
  const ss = getDbSpreadsheet();
  const sheetPerm = ss.getSheetByName(SHEETS.PERMISOS_ADMINISTRATIVOS);
  const data = sheetPerm.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idPermiso) {
      sheetPerm.getRange(i + 1, 10).setValue('ANULADO');
      
      const sheetRep = ss.getSheetByName(SHEETS.REEMPLAZOS);
      const dataRep = sheetRep.getDataRange().getValues();
      for (let j = 1; j < dataRep.length; j++) {
        if (dataRep[j][10] === ('Generado por Permiso Admin. ' + idPermiso) && dataRep[j][9] === 'ERE-0001') {
          sheetRep.getRange(j + 1, 10).setValue('ERE-0006');
        }
      }

      logAuditoria('ANULAR', 'PERMISOS_ADMINISTRATIVOS', idPermiso, 'APROBADO', 'ANULADO');
      return { success: true };
    }
  }
  throw new Error('Permiso Administrativo no encontrado.');
}

function apiGetReemplazos() {
  const ss = getDbSpreadsheet();
  const reemplazos = getTableData(SHEETS.REEMPLAZOS);
  const cursos = getTableData(SHEETS.CURSOS);
  const asignaturas = getTableData(SHEETS.ASIGNATURAS);
  const estados = getTableData(SHEETS.ESTADOS_REEMPLAZO);

  const personalSheet = ss.getSheetByName(SHEETS.PERSONAL);
  const mapPersonal = {};
  if (personalSheet) {
    const personalData = personalSheet.getDataRange().getValues();
    const perHeaders = personalData[0];
    const idxId = perHeaders.indexOf('ID_Personal') !== -1 ? perHeaders.indexOf('ID_Personal') : perHeaders.indexOf('ID_Profesor');
    const idxNom = perHeaders.indexOf('Nombres');
    const idxPat = perHeaders.indexOf('Apellido_Paterno');
    const idxComp = perHeaders.indexOf('Nombre_Completo');

    for (let i = 1; i < personalData.length; i++) {
      let row = personalData[i];
      let idPersonal = String(row[idxId] || '').trim();
      let nombreComp = idxComp !== -1 ? String(row[idxComp] || '').trim() : '';
      let nombreFinal = nombreComp || `${row[idxNom] || ''} ${row[idxPat] || ''}`.trim();
      if (idPersonal) mapPersonal[idPersonal] = nombreFinal || idPersonal;
    }
  }

  const mapCursos = {}; cursos.forEach(c => mapCursos[c.ID_Curso] = c.Nombre_Curso);
  const mapAsig = {}; asignaturas.forEach(a => mapAsig[a.ID_Asignatura] = a.Nombre_Asignatura);

  const mapEstados = {
    'ERE-0001': 'PENDIENTE', 'ERE-0002': 'ASIGNADO', 'ERE-0003': 'CONFIRMADO',
    'ERE-0004': 'EJECUTADO', 'ERE-0005': 'NO REALIZADO', 'ERE-0006': 'CANCELADO', 'ERE-0007': 'MODIFICADO'
  };
  estados.forEach(e => {
    if (e.ID_EstadoReemplazo && e.Nombre_Estado) mapEstados[e.ID_EstadoReemplazo] = e.Nombre_Estado;
  });

  return reemplazos.map(r => ({
    ...r,
    Fecha_Inicio: formatearFechaLimpia(r.Fecha_Inicio || r.Fecha_Reemplazo),
    Nombre_Titular: mapPersonal[String(r.ID_Personal_Titular).trim()] || r.ID_Personal_Titular || 'Sin Titular',
    Nombre_Reemplazante: mapPersonal[String(r.ID_Personal_Reemplazante).trim()] || r.ID_Personal_Reemplazante || 'Sin Asignar',
    Nombre_Curso: mapCursos[r.ID_Curso] || r.ID_Curso,
    Nombre_Asignatura: mapAsig[r.ID_Asignatura] || r.ID_Asignatura,
    Texto_Estado: mapEstados[r.ID_EstadoReemplazo] || r.ID_EstadoReemplazo || 'PENDIENTE'
  }));
}

function apiAsignarReemplazante(idReemplazo, idPersonalReemplazante) {
  const ss = getDbSpreadsheet();
  const sheetRep = ss.getSheetByName(SHEETS.REEMPLAZOS);
  const dataRep = sheetRep.getDataRange().getValues();
  
  let rowIndex = -1;
  let targetRep = null;
  for (let i = 1; i < dataRep.length; i++) {
    if (String(dataRep[i][0]).trim() === String(idReemplazo).trim()) { rowIndex = i + 1; targetRep = dataRep[i]; break; }
  }
  if (rowIndex === -1) throw new Error('Reemplazo no encontrado.');

  const fechaInicioRep = targetRep[1];
  const idBloqueRep = targetRep[7];

  const todosReemplazos = getTableData(SHEETS.REEMPLAZOS);
  const conflicto = todosReemplazos.some(r => 
    r.ID_Personal_Reemplazante === idPersonalReemplazante &&
    r.Fecha_Inicio === fechaInicioRep &&
    r.ID_Bloque === idBloqueRep &&
    r.ID_Reemplazo !== idReemplazo &&
    ['ERE-0002', 'ERE-0003'].includes(r.ID_EstadoReemplazo)
  );

  if (conflicto) throw new Error('CONFLICTO HORARIO: El docente ya posee un reemplazo asignado en este bloque.');

  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheetRep.getRange(rowIndex, 5).setValue(idPersonalReemplazante);
  sheetRep.getRange(rowIndex, 10).setValue('ERE-0002');
  sheetRep.getRange(rowIndex, 12).setValue(nowStr);

  logAuditoria('ASIGNAR', 'REEMPLAZOS', idReemplazo, targetRep[4], idPersonalReemplazante);
  return { success: true };
}

function apiGuardarReemplazoInterno(datos) {
  const ss = getDbSpreadsheet();
  const sheetRep = ss.getSheetByName(SHEETS.REEMPLAZOS);
  
  let fechaLimpiaInicio = formatearFechaLimpia(datos.Fecha_Inicio || datos.fechaInicio);
  let fechaLimpiaTermino = formatearFechaLimpia(datos.Fecha_Termino || datos.fechaTermino || datos.Fecha_Inicio);
  
  let fInicio = new Date(fechaLimpiaInicio + 'T00:00:00');
  let fTermino = new Date(fechaLimpiaTermino + 'T00:00:00');
  
  if (fTermino < fInicio) throw new Error("La fecha de término no puede ser menor a la fecha de inicio.");

  const nuevoId = generateNextId(SHEETS.REEMPLAZOS, PREFIXES.REEMPLAZOS);
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  let nuevaFila = [
    nuevoId, fechaLimpiaInicio, fechaLimpiaTermino,
    datos.ID_Personal_Titular || datos.idPersonalTitular || '',
    datos.ID_Personal_Reemplazante || datos.idPersonalReemplazante,
    datos.ID_Curso || datos.idCurso, datos.ID_Asignatura || datos.idAsignatura,
    datos.ID_Bloque || datos.idBloque, datos.ID_Hora || datos.idHora || '',
    datos.ID_EstadoReemplazo || 'ERE-0003',
    datos.Observaciones || 'Reemplazo directo asignado desde la plataforma', nowStr
  ];

  sheetRep.appendRow(nuevaFila);
  logAuditoria('CREAR_INTERNO', 'REEMPLAZOS', nuevoId, null, datos);
  return { success: true, id: nuevoId, message: "Reemplazo registrado correctamente." };
}

function apiEliminarReemplazoDefinitivo(idReemplazo) {
  const ss = getDbSpreadsheet();
  const sheetRep = ss.getSheetByName(SHEETS.REEMPLAZOS);
  if (!sheetRep) throw new Error('Hoja de Reemplazos no encontrada.');

  const data = sheetRep.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(idReemplazo).trim()) {
      const registroBorrado = data[i];
      sheetRep.deleteRow(i + 1);
      logAuditoria('ELIMINAR_FISICO', 'REEMPLAZOS', idReemplazo, registroBorrado, null);
      return { success: true, message: 'Reemplazo eliminado físicamente de la base de datos.' };
    }
  }
  throw new Error('No se encontró el reemplazo a eliminar.');
}
/**
 * Actualiza o inicializa la estructura de la hoja 'Reemplazos' en Google Sheets
 * con las columnas requeridas para el sistema de reemplazos por bloques y fechas.
 */
function actualizarEstructuraTablaReemplazos() {
  // Intentamos obtener la hoja activa; si es null, abrimos la planilla directamente por su ID
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss) {
    var idSpreadsheet = "1IIPorweYxVVhfR2m3-rTmsUBn_h6d3pfV1iVahLRLV0";
    ss = SpreadsheetApp.openById(idSpreadsheet);
  }
  
  var nombreHoja = "Reemplazos";
  var hoja = ss.getSheetByName(nombreHoja);
  
  // Si la hoja no existe, la creamos
  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
  }
  
  // Definir los nuevos encabezados limpios y normalizados
  var nuevosEncabezados = [
    "ID_Reemplazo",
    "Fecha_Clase",
    "ID_Personal_Titular",
    "ID_Personal_Reemplazante",
    "ID_Horario",
    "ID_Curso",
    "ID_Asignatura",
    "ID_Bloque",
    "ID_EstadoReemplazo",
    "Observaciones",
    "Fecha_Registro"
  ];
  
  // Escribir los nuevos encabezados en la primera fila
  hoja.getRange(1, 1, 1, nuevosEncabezados.length).setValues([nuevosEncabezados]);
  
  // Dar formato estético básico a la cabecera (negrita y fondo gris claro)
  var rangoCabecera = hoja.getRange(1, 1, 1, nuevosEncabezados.length);
  rangoCabecera.setFontWeight("bold");
  rangoCabecera.setBackground("#f3f3f3");
  
  // Inmovilizar la primera fila para mejor visualización
  hoja.setFrozenRows(1);
  
  Logger.log("¡Éxito! La tabla 'Reemplazos' ha sido actualizada con la nueva estructura.");
}

function apiGetCalendarioEventos() {
  const reemplazos = getTableData(SHEETS.REEMPLAZOS);
  const personal = getTableData(SHEETS.PERSONAL);
  const cursos = getTableData(SHEETS.CURSOS);
  const asignaturas = getTableData(SHEETS.ASIGNATURAS);

  const mapPersonal = {}; personal.forEach(p => mapPersonal[p.ID_Personal] = p.Nombre_Completo || `${p.Nombres || ''} ${p.Apellido_Paterno || ''}`.trim());
  const mapCursos = {}; cursos.forEach(c => mapCursos[c.ID_Curso] = c.Nombre_Curso);
  const mapAsig = {}; asignaturas.forEach(a => mapAsig[a.ID_Asignatura] = a.Nombre_Asignatura);

  return reemplazos.map(r => ({
    id: r.ID_Reemplazo,
    title: `${mapCursos[r.ID_Curso] || r.ID_Curso} - ${mapAsig[r.ID_Asignatura] || r.ID_Asignatura}`,
    start: formatearFechaLimpia(r.Fecha_Inicio),
    end: formatearFechaLimpia(r.Fecha_Termino),
    titular: mapPersonal[r.ID_Personal_Titular] || r.ID_Personal_Titular || 'Sin Titular',
    reemplazante: mapPersonal[r.ID_Personal_Reemplazante] || r.ID_Personal_Reemplazante || 'Sin Asignar',
    bloque: 'Bloque ' + r.ID_Bloque,
    estado: r.ID_EstadoReemplazo
  }));
}

function apiGetAuditoria() { return getTableData(SHEETS.AUDITORIA); }
