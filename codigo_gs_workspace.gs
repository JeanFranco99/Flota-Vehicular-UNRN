/**
 * =================================================================
 * VERSIÓN 2: WORKSPACE (DINÁMICO, AUTOGENERADO Y CON CAPTURA DE ERRORES)
 * =================================================================
 */

function obtenerSpreadsheetId() {
  return PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID_GLOBAL');
}

const ENCABEZADOS_HOJAS = {
  'Solicitudes': ['ID', 'Email Solicitante', 'Fecha Salida', 'Fecha Regreso', 'Itinerario', 'Tareas', 'Participantes', 'Conductores', 'Estado', 'Vehículo Asignado', 'Observaciones', 'Gestionado por'],
  'Vehículos': ['ID', 'Patente', 'Marca', 'Modelo', 'Año', 'Capacidad', 'Estado', 'Observaciones', 'Combustible', 'Velocidad Máxima','Fecha VTO RTO', 'Fecha VTO Seguro'],
  'Mantenimientos': ['ID', 'Patente', 'Tipo', 'Detalle', 'Costo', 'Responsable', 'Fecha Inicio', 'Fecha Fin', 'Estado'],
  'Configuración': ['Rol', 'Email'],
  'Controles': ['ID', 'ID_Solicitud', 'Tipo', 'Sección', 'Item', 'Valor', 'Fecha', 'Usuario']
};

function obtenerOInicializarBaseDeDatosGlobal() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var spreadsheetId = scriptProperties.getProperty('DB_SPREADSHEET_ID_GLOBAL');
  
  if (spreadsheetId) {
    try {
      DriveApp.getFileById(spreadsheetId);
      return spreadsheetId; 
    } catch (e) { spreadsheetId = null; }
  }
  
  if (!spreadsheetId) {
    var ss = SpreadsheetApp.create('Flota Universitaria - Base de Datos Global');
    spreadsheetId = ss.getId();
    
    var estructuraBaseDeDatos = {
      'Solicitudes': ["ID", "Email Solicitante", "Fecha Salida", "Fecha Regreso", "Itinerario", "Tareas", "Participantes", "Conductores", "Estado", "Vehículo Asignado", "Observaciones", "Gestionado por"],
      'Vehículos': ["ID", "Patente", "Marca", "Modelo", "Año", "Capacidad", "Estado", "Observaciones", "Combustible", "Velocidad Máxima", "Fecha VTO RTO", "Fecha VTO Seguro"],
      'Mantenimientos': ["ID", "Patente", "Tipo", "Detalle", "Costo", "Responsable", "Fecha Inicio", "Fecha Fin", "Estado"],
      'Controles': ["ID", "ID_Solicitud", "Tipo", "Sección", "Item", "Valor", "Fecha", "Usuario"],
      'Configuración': ["Rol", "Email"],
      '_Logs': ["Timestamp", "Usuario", "Función", "Acción", "Detalles", "Resultado"],
      'Verificaciones': ["ID", "Patente", "Fecha", "Tipo", "ID Mantenimiento", "Resultado", "Verificado por", "Observaciones", "Fotos", "Embriague", "FrenoMotor", "FrenoPie", "FrenoMano", "Direccion", "Velocimetro", "Temperatura", "Tablero", "LimpiaParabrisas", "Bocina", "Espejos", "LucesBajas", "LucesAltas", "LucesIntermitentes", "LucesFreno", "LucesReversa", "GataLlaveCruz", "Triangulos", "Extintor", "Chaleco", "MotorAceite", "Refrigerante", "LiquidoFrenos", "DireccionHidraulica", "Bateria", "LimpiaParabrisasMec", "Neumaticos", "Suspencion", "Escape", "CarroceriaRayones", "CarroceriaAbolladuras", "LimpiezaInterior", "LimpiezaExterior", "DocLicencia", "DocSeguro", "DocRTO", "DocTitulo", "EquipoEmergencia"]
    };
    
    var hojasIniciales = ss.getSheets();
    var primeraHoja = hojasIniciales[0];
    var esPrimeraIteracion = true;
    
    for (var nombreHoja in estructuraBaseDeDatos) {
      var hojaActiva = esPrimeraIteracion ? primeraHoja : ss.insertSheet(nombreHoja);
      if (esPrimeraIteracion) { hojaActiva.setName(nombreHoja); esPrimeraIteracion = false; }
      
      var columnas = estructuraBaseDeDatos[nombreHoja];
      var rangoEncabezado = hojaActiva.getRange(1, 1, 1, columnas.length);
      rangoEncabezado.setValues([columnas]);
      rangoEncabezado.setFontWeight("bold");
      hojaActiva.setFrozenRows(1);
    }
    scriptProperties.setProperty('DB_SPREADSHEET_ID_GLOBAL', spreadsheetId);
  }
  return spreadsheetId;
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }

function doGet(e) {
  obtenerOInicializarBaseDeDatosGlobal();

  const params = e && e.parameter ? e.parameter : {};
  const userEmail = Session.getActiveUser().getEmail();

  if (params.page === 'imprimir' && params.id) {
    if (esEncargado(userEmail) || esSolicitante(userEmail)) { 
      const idSolicitud = params.id;
      const datosSolicitud = obtenerDatosSolicitud(idSolicitud);
      if (datosSolicitud) {
        if (!esEncargado(userEmail) && datosSolicitud.emailUsuario !== userEmail) {
           return HtmlService.createHtmlOutput('No tiene permiso para ver este comprobante.');
        }
        const template = HtmlService.createTemplateFromFile('comprobante');
        Object.keys(datosSolicitud).forEach(key => template[key] = datosSolicitud[key]);
        return template.evaluate().setTitle(`Comprobante Solicitud N° ${idSolicitud}`);
      }
    }
    return HtmlService.createHtmlOutput('No tiene permiso para ver este comprobante o la solicitud no existe.');
  }

  if (params.page === 'imprimirControl' && params.id) {
    if (esEncargado(userEmail)) { 
      const idSolicitud = params.id;
      const controles = obtenerControlesPorSolicitud(idSolicitud);
      const datosSolicitud = obtenerDatosSolicitud(idSolicitud);
      if (datosSolicitud) {
        const html = generarHtmlControl(idSolicitud, datosSolicitud, controles);
        return HtmlService.createHtmlOutput(html).setTitle(`Control Vehicular N° ${idSolicitud}`);
      }
    }
    return HtmlService.createHtmlOutput('No tiene permiso o la solicitud no existe.');
  }

  if (esEncargado(userEmail)) {
    return HtmlService.createTemplateFromFile('Encargado').evaluate().setTitle('Panel Encargado - Flota UNRN');
  } else if (esSolicitante(userEmail)) {
    return HtmlService.createTemplateFromFile('Solicitante').evaluate().setTitle('Panel Solicitante - Flota UNRN');
  } else {
    return HtmlService.createHtmlOutput(`<h1>Acceso Denegado</h1><p>Debes usar una cuenta autorizada para acceder a este sistema.</p>`).setTitle('No Autorizado');
  }
}

function obtenerDatosSolicitud(id) {
    const ss = obtenerSpreadsheet();
    if (!ss) return null;
    const hojaSolicitudes = ss.getSheetByName('Solicitudes');
    const datosSolicitudes = hojaSolicitudes.getDataRange().getValues();

    for (let i = 1; i < datosSolicitudes.length; i++) {
        if (datosSolicitudes[i][0] == id) {
            const solicitudData = mapRowToSolicitudObject(datosSolicitudes[i]);
            if (solicitudData.vehiculoAsignado) {
                const patenteAsignada = solicitudData.vehiculoAsignado.split(' - ')[0].trim();
                const hojaVehiculos = ss.getSheetByName('Vehículos');
                const datosVehiculos = hojaVehiculos.getDataRange().getValues();
                for (let j = 1; j < datosVehiculos.length; j++) {
                    const patenteVehiculo = datosVehiculos[j][1]; 
                    if (patenteVehiculo && patenteVehiculo.toLowerCase() === patenteAsignada.toLowerCase()) {
                        solicitudData.combustible = datosVehiculos[j][8] || 'No especificado';
                        solicitudData.velocidadMaxima = datosVehiculos[j][9] || 'No especificada';
                        break; 
                    }
                }
            }
            if (!solicitudData.combustible) solicitudData.combustible = 'No especificado';
            if (!solicitudData.velocidadMaxima) solicitudData.velocidadMaxima = 'No especificada';
            return solicitudData;
        }
    }
    return null;
}

function obtenerSpreadsheet() {
  try {
    const idDinamico = obtenerSpreadsheetId();
    if (!idDinamico) throw new Error("No hay ID global registrado.");
    
    const ss = SpreadsheetApp.openById(idDinamico);
    Object.keys(ENCABEZADOS_HOJAS).forEach(nombre => {
      let hoja = ss.getSheetByName(nombre);
      if (!hoja) {
        hoja = ss.insertSheet(nombre);
        if (hoja.getLastRow() === 0) hoja.appendRow(ENCABEZADOS_HOJAS[nombre]);
      }
    });
    return ss;
  } catch (e) {
    Logger.log("Error al abrir Spreadsheet por ID. " + e.message);
    return null;
  }
}

function obtenerEncargados() {
  try {
    const hojaConfig = obtenerSpreadsheet().getSheetByName('Configuración');
    if (hojaConfig.getLastRow() < 2) return [];
    const datos = hojaConfig.getRange('B2:B' + hojaConfig.getLastRow()).getValues();
    return datos.map(row => row[0]).filter(email => email && email.trim() !== '');
  } catch (e) { return []; }
}

function esEncargado(email) {
  if (!email) return false;
  return obtenerEncargados().includes(email);
}

function esSolicitante(email) {
  if (!email) return false;
  return !esEncargado(email);
}

function getCurrentUserEmail() { return Session.getActiveUser().getEmail(); }

function obtenerEstadisticasDashboard() {
  if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
  const hojaVehiculos = obtenerSpreadsheet().getSheetByName('Vehículos');
  const datosVehiculos = hojaVehiculos.getLastRow() > 1 ? hojaVehiculos.getRange('G2:G' + hojaVehiculos.getLastRow()).getValues() : [];
  const hojaSolicitudes = obtenerSpreadsheet().getSheetByName('Solicitudes');
  const datosSolicitudes = hojaSolicitudes.getLastRow() > 1 ? hojaSolicitudes.getRange('I2:I' + hojaSolicitudes.getLastRow()).getValues() : [];
  const estadisticas = { vehiculosDisponibles: 0, vehiculosAsignados: 0, vehiculosMantenimiento: 0, solicitudesPendientes: 0 };
  datosVehiculos.forEach(row => {
    const estado = (row[0] || '').toLowerCase();
    if (estado === 'disponible') estadisticas.vehiculosDisponibles++;
    else if (estado === 'asignado') estadisticas.vehiculosAsignados++;
    else if (estado === 'mantenimiento') estadisticas.vehiculosMantenimiento++;
  });
  datosSolicitudes.forEach(row => { if ((row[0] || '').toLowerCase() === 'pendiente') estadisticas.solicitudesPendientes++; });
  return estadisticas;
}

function obtenerMantenimientos() {
  const ss = obtenerSpreadsheet(); // Corrección: Evitar getActiveSpreadsheet()
  if (!ss) return [];
  const hojaMantenimientos = ss.getSheetByName('Mantenimientos');
  const datos = hojaMantenimientos.getDataRange().getValues();
  const cabecera = datos.shift(); 
  return datos.map(fila => {
    const obj = {};
    cabecera.forEach((nombreColumna, i) => obj[nombreColumna.trim().replace(/\s/g, '')] = fila[i]);
    return obj;
  });
}

function aprobarSolicitud(idSolicitud, patenteAsignada, fechaSalida, fechaRegreso, controlData = null) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (controlData && controlData.items && controlData.items.length > 0) {
      try { registrarControl({ idSolicitud: idSolicitud, tipo: 'inicial', items: controlData.items, kilometros: controlData.kilometros || '', observaciones: controlData.observaciones || '' }); } catch (e) {}
    }
    
    if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
    const hojaSolicitudes = obtenerSpreadsheet().getSheetByName('Solicitudes');
    const hojaVehiculos = obtenerSpreadsheet().getSheetByName('Vehículos');

    const datosSolicitudes = hojaSolicitudes.getDataRange().getValues();
    let filaSolicitud = -1, solicitudActual;
    for (let i = 1; i < datosSolicitudes.length; i++) {
      if (datosSolicitudes[i][0] == idSolicitud) {
        solicitudActual = mapRowToSolicitudObject(datosSolicitudes[i]);
        filaSolicitud = i + 1;
        break;
      }
    }

    if (!solicitudActual) throw new Error('Solicitud no encontrada.');
    if (solicitudActual.estado.toLowerCase() !== 'pendiente') throw new Error(`Estado actual es "${solicitudActual.estado}".`);

    const datosVehiculos = hojaVehiculos.getDataRange().getValues();
    let filaVehiculo = -1, modeloVehiculo = '';
    for (let i = 1; i < datosVehiculos.length; i++) {
      if (String(datosVehiculos[i][1]).toLowerCase() === String(patenteAsignada).toLowerCase()) {
        const estadoVehiculo = (datosVehiculos[i][6] || '').toLowerCase();
        if (estadoVehiculo !== 'disponible') throw new Error(`Vehículo en estado "${estadoVehiculo}".`);
        modeloVehiculo = datosVehiculos[i][3];
        filaVehiculo = i + 1;
        break;
      }
    }
    if (filaVehiculo === -1) throw new Error(`Vehículo no encontrado.`);

    let fechaSalidaValidada, fechaRegresoValidada;
    const todasLasSolicitudes = hojaSolicitudes.getDataRange().getValues();
    let solicitudFechaSalida = null, solicitudFechaRegreso = null;
    for (let i = 1; i < todasLasSolicitudes.length; i++) {
      if (todasLasSolicitudes[i][0] == idSolicitud) {
        solicitudFechaSalida = todasLasSolicitudes[i][2];
        solicitudFechaRegreso = todasLasSolicitudes[i][3];
        break;
      }
    }
    
    if (solicitudFechaSalida && solicitudFechaRegreso) {
      fechaSalidaValidada = new Date(solicitudFechaSalida);
      fechaRegresoValidada = new Date(solicitudFechaRegreso);
    } else {
      if (!fechaSalida || !fechaRegreso || fechaSalida.trim() === '' || fechaRegreso.trim() === '') throw new Error('Fechas inválidas.');
      fechaSalidaValidada = parsearFecha(fechaSalida);
      fechaRegresoValidada = parsearFecha(fechaRegreso);
    }
    
    if (isNaN(fechaSalidaValidada.getTime()) || isNaN(fechaRegresoValidada.getTime())) throw new Error('Fechas inválidas.');

    verificarDisponibilidadVehiculo(patenteAsignada, fechaSalidaValidada, fechaRegresoValidada, idSolicitud);

    const descripcionVehiculo = modeloVehiculo ? `${patenteAsignada.toUpperCase()} - ${modeloVehiculo}` : patenteAsignada.toUpperCase();

    hojaVehiculos.getRange(filaVehiculo, 7).setValue('Asignado');
    hojaSolicitudes.getRange(filaSolicitud, 9).setValue('Aprobada');
    hojaSolicitudes.getRange(filaSolicitud, 10).setValue(descripcionVehiculo);
    hojaSolicitudes.getRange(filaSolicitud, 12).setValue(getCurrentUserEmail());

    const asunto = `Solicitud de Vehículo #${idSolicitud} - APROBADA`;
    const cuerpo = `Hola,<br><br>Tu solicitud fue <b>APROBADA.<b><br><br><b>Detalles:</b><br>- ID: ${idSolicitud}<br>- Vehículo: ${descripcionVehiculo}<br>- Fechas: ${formatDateTime(fechaSalida)} a ${formatDateTime(fechaRegreso)}`;
    try { MailApp.sendEmail(solicitudActual.emailUsuario, asunto, "", {htmlBody: cuerpo}); } catch (e) {}

    return { success: true, message: 'Solicitud aprobada.' };
  } catch (error) { return { success: false, message: error.message }; } finally { lock.releaseLock(); }
}

function verificarDisponibilidadVehiculo(patente, fechaSalida, fechaRegreso, idSolicitudExcluida = null) {
  const ss = obtenerSpreadsheet();
  const hojaSolicitudes = ss.getSheetByName('Solicitudes');

  if (hojaSolicitudes.getLastRow() >= 2) {
    const datosSolicitudes = hojaSolicitudes.getDataRange().getValues().slice(1);
    const solicitudesAprobadas = datosSolicitudes.filter(row => (row[8] || '').toLowerCase() === 'aprobada' && (row[9] || '').toUpperCase().startsWith(patente.toUpperCase()) && row[0] != idSolicitudExcluida);
    for (const sol of solicitudesAprobadas) {
      if (fechaSalida.getTime() < new Date(sol[3]).getTime() && fechaRegreso.getTime() > new Date(sol[2]).getTime()) {
        throw new Error(`Conflicto de fechas. Vehículo asignado a la solicitud #${sol[0]}.`);
      }
    }
  }

  const hojaMantenimientos = ss.getSheetByName('Mantenimientos');
  if (hojaMantenimientos.getLastRow() >= 2) {
      const datosMantenimientos = hojaMantenimientos.getDataRange().getValues().slice(1);
      const mantenimientosProgramados = datosMantenimientos.filter(row => (row[1] || '').toUpperCase() === patente.toUpperCase() && (row[8] || '').toLowerCase() === 'en curso');
      for (const mant of mantenimientosProgramados) {
          const inicioMantenimiento = parsearFecha(mant[6]);
          let finMantenimiento = parsearFecha(mant[7]);
          if (isNaN(inicioMantenimiento.getTime())) continue; 
          if (isNaN(finMantenimiento.getTime())) finMantenimiento = new Date(8640000000000000); 
          if (fechaSalida.getTime() < finMantenimiento.getTime() && fechaRegreso.getTime() > inicioMantenimiento.getTime()) {
              throw new Error(`Conflicto. Vehículo con mantenimiento programado.`);
          }
      }
  }
  return true;
}

function obtenerMantenimientosActivosConFechas() {
  if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
  const hojaMantenimientos = obtenerSpreadsheet().getSheetByName('Mantenimientos');
  if (hojaMantenimientos.getLastRow() < 2) return [];
  const datos = hojaMantenimientos.getDataRange().getValues().slice(1);
  return datos.filter(row => (row[8] || '').toLowerCase() === 'en curso').map(row => ({ id: row[0], patente: row[1], tipo: row[2], detalle: row[3], fechaInicio: row[6], fechaInicioFormateada: formatDateTime(row[6]) }));
}

function verificarDisponibilidadEspecifica(patente, fechaSalida, fechaRegreso) {
  if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
  try {
    verificarDisponibilidadVehiculo(patente, fechaSalida, fechaRegreso);
    return { disponible: true, mensaje: 'Vehículo disponible', conflictos: [] };
  } catch (error) { return { disponible: false, mensaje: error.message, conflictos: [{ tipo: error.message.includes('mantenimiento') ? 'mantenimiento' : 'solicitud', detalle: error.message }] }; }
}

function listarSolicitudesPorEstados(estados) {
    if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
    const hoja = obtenerSpreadsheet().getSheetByName('Solicitudes');
    if (hoja.getLastRow() < 2) return [];
    const estadosEnMinuscula = estados.map(e => e.toLowerCase());
    return hoja.getDataRange().getValues().slice(1).filter(row => estadosEnMinuscula.includes((row[8] || '').toLowerCase())).map(mapRowToSolicitudObject);
}

function listarSolicitudesUsuarioPorEstados(estados) {
    const emailUsuario = getCurrentUserEmail();
    if (!esSolicitante(emailUsuario)) throw new Error('No autorizado');
    const hoja = obtenerSpreadsheet().getSheetByName('Solicitudes');
    if (hoja.getLastRow() < 2) return [];
    const estadosEnMinuscula = estados.map(e => e.toLowerCase());
    return hoja.getDataRange().getValues().slice(1).filter(row => row[1] === emailUsuario && estadosEnMinuscula.includes((row[8] || '').toLowerCase())).map(mapRowToSolicitudObject);
}

function listarVehiculos() {
  if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
  const hoja = obtenerSpreadsheet().getSheetByName('Vehículos');
  if (hoja.getLastRow() < 2) return [];
  const vehiculos = [];
  hoja.getDataRange().getValues().slice(1).forEach((row) => {
    if (row.every(cell => cell === "")) return;
    try { vehiculos.push({ id: row[0], patente: row[1], marca: row[2], modelo: row[3], anio: row[4], capacidad: row[5], estado: row[6], observaciones: row[7], combustible: row[8], velocidadMaxima: row[9], fechaVtoRto: formatDateTime(row[10]), fechaVtoSeguro: formatDateTime(row[11]) }); } catch (e) { }
  });
  return vehiculos;
}

function listarMantenimientosActivos() {
    if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
    const hoja = obtenerSpreadsheet().getSheetByName('Mantenimientos');
    if (hoja.getLastRow() < 2) return [];
    return hoja.getDataRange().getValues().slice(1).filter(row => (row[8] || '').toLowerCase() === 'en curso').map(row => ({ id: row[0], patente: row[1], tipo: row[2], detalle: row[3], costo: row[4], responsable: row[5], fechaInicio: formatDateTime(row[6]) }));
}

function listarMantenimientosFinalizados() {
    if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
    const hoja = obtenerSpreadsheet().getSheetByName('Mantenimientos');
    if (hoja.getLastRow() < 2) return [];
    return hoja.getDataRange().getValues().slice(1).filter(row => (row[8] || '').toLowerCase() === 'finalizado').map(row => ({ id: row[0], patente: row[1], tipo: row[2], detalle: row[3], costo: row[4], responsable: row[5], fechaInicio: formatDateTime(row[6]), fechaFin: formatDateTime(row[7]) }));
}

function enviarSolicitud(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const emailUsuario = getCurrentUserEmail();
    
    // Bloqueo de uso y atrapado de errores en interfaz
    if (!esSolicitante(emailUsuario)) throw new Error('Usuario no autorizado para enviar solicitudes. (¿Estás como Encargado?)');
    const itinerario = (formData.itinerario || '').trim();
    const tareas = (formData.tareas || '').trim();
    if (!itinerario || !tareas) throw new Error("Itinerario y tareas son obligatorios.");
    
    const salida = new Date(formData.salida);
    const regreso = new Date(formData.regreso);
    if (regreso <= salida) throw new Error('El regreso debe ser posterior a la salida.');
    if (salida < new Date()) throw new Error('La salida no puede ser en el pasado.');

    const hoja = obtenerSpreadsheet().getSheetByName('Solicitudes');
    const ultimaFila = hoja.getLastRow();
    const nuevaId = ultimaFila > 1 ? (parseInt(hoja.getRange(ultimaFila, 1).getValue(), 10) || 0) + 1 : 1;

    // Se agrega los strings vacios suficientes para 12 columnas exactas
    hoja.appendRow([ nuevaId, emailUsuario, salida, regreso, itinerario, tareas, formData.participantes, formData.conductores, 'Pendiente', '', '', '' ]);
    
    const encargados = obtenerEncargados();
    if (encargados.length > 0) {
      const asunto = `Nueva Solicitud #${nuevaId} - ${emailUsuario}`;
      const cuerpo = `Solicitante: ${emailUsuario}<br>Desde: ${salida}<br>Hasta: ${regreso}`;
      MailApp.sendEmail(encargados.join(','), asunto, "", {htmlBody: cuerpo});
    }
    return { success: true, message: 'Solicitud enviada.' };
  } catch (error) { 
    Logger.log('Error enviarSolicitud: ' + error.message);
    return { success: false, message: error.message }; 
  } finally { lock.releaseLock(); }
}

function rechazarSolicitud(idSolicitud, motivo) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
      const hoja = obtenerSpreadsheet().getSheetByName('Solicitudes');
      const datos = hoja.getDataRange().getValues();
      for (let i = 1; i < datos.length; i++) {
          if (datos[i][0] == idSolicitud) {
              if ((datos[i][8] || '').toLowerCase() !== 'pendiente') throw new Error(`Estado actual inválido.`);
              hoja.getRange(i + 1, 9).setValue('Rechazada');
              hoja.getRange(i + 1, 11).setValue(motivo);
              hoja.getRange(i + 1, 10).setValue('');
              hoja.getRange(i + 1, 12).setValue(getCurrentUserEmail());
              MailApp.sendEmail(datos[i][1], "Solicitud Rechazada", "", {htmlBody: `Motivo: ${motivo}`});
              return { success: true, message: 'Solicitud rechazada.' };
          }
      }
      throw new Error('Solicitud no encontrada.');
    } finally { lock.releaseLock(); }
}

function finalizarSolicitud(idSolicitud, controlData = null) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (controlData && controlData.items && controlData.items.length > 0) {
        try { registrarControl({ idSolicitud: idSolicitud, tipo: 'final', items: controlData.items, kilometros: controlData.kilometros || '', observaciones: controlData.observaciones || '' }); } catch (e) { }
      }
      if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
      
      const hojaSolicitudes = obtenerSpreadsheet().getSheetByName('Solicitudes');
      const datosSolicitudes = hojaSolicitudes.getDataRange().getValues();
      let filaSolicitud = -1, vehiculoAsignado = '';
      for (let i = 1; i < datosSolicitudes.length; i++) {
        if (datosSolicitudes[i][0] == idSolicitud) {
          if ((datosSolicitudes[i][8] || '').toLowerCase() !== 'aprobada') throw new Error('Estado no es Aprobada');
          filaSolicitud = i + 1;
          vehiculoAsignado = datosSolicitudes[i][9];
          break;
        }
      }
      if (filaSolicitud === -1) throw new Error('Solicitud no encontrada.');

      hojaSolicitudes.getRange(filaSolicitud, 9).setValue('Finalizada');
      hojaSolicitudes.getRange(filaSolicitud, 12).setValue(getCurrentUserEmail());
      if (vehiculoAsignado) {
        const patente = vehiculoAsignado.split(' - ')[0].trim();
        const hojaVehiculos = obtenerSpreadsheet().getSheetByName('Vehículos');
        const datosVehiculos = hojaVehiculos.getDataRange().getValues();
        for (let j = 1; j < datosVehiculos.length; j++) {
          if (String(datosVehiculos[j][1]).toLowerCase() === patente.toLowerCase()) {
            hojaVehiculos.getRange(j + 1, 7).setValue('Disponible');
            break;
          }
        }
      }
      return { success: true, message: 'Solicitud finalizada.' };
    } finally { lock.releaseLock(); }
}

function agregarVehiculo(vehiculoData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
      if (!vehiculoData.fechaVtoRto || !vehiculoData.fechaVtoSeguro) throw new Error('Fechas obligatorias.');
      
      const patenteFormateada = validarYFormatearPatente(vehiculoData.patente);
      const hoja = obtenerSpreadsheet().getSheetByName('Vehículos');
      const datosVehiculos = hoja.getLastRow() > 1 ? hoja.getRange(2, 2, hoja.getLastRow() - 1, 1).getValues() : [];
      if (datosVehiculos.flat().map(p => (p || '').toUpperCase()).includes(patenteFormateada)) throw new Error('Patente ya registrada.');

      const ultimaFila = hoja.getLastRow();
      const nuevaId = ultimaFila > 1 ? (parseInt(hoja.getRange(ultimaFila, 1).getValue(), 10) || 0) + 1 : 1;
      hoja.appendRow([ nuevaId, patenteFormateada, vehiculoData.marca, vehiculoData.modelo, Number(vehiculoData.anio), Number(vehiculoData.capacidad), 'Disponible', vehiculoData.observaciones || '', vehiculoData.combustible || '', Number(vehiculoData.velocidadMaxima) || 0, new Date(vehiculoData.fechaVtoRto), new Date(vehiculoData.fechaVtoSeguro) ]);
      return { success: true, message: 'Vehículo agregado.' };
    } finally { lock.releaseLock(); }
}

function validarYFormatearPatente(patente) {
  if (!patente || patente.trim() === '') throw new Error('Patente obligatoria.');
  const patenteUpper = patente.trim().toUpperCase();
  if (!/^(?:[A-Z]{3}\s\d{3}|[A-Z]{2}\s\d{3}\s[A-Z]{2})$/.test(patenteUpper)) throw new Error('Formato no válido.');
  return patenteUpper;
}

function editarVehiculo(vehiculoData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
        const patenteFormateada = validarYFormatearPatente(vehiculoData.patente);
        const hoja = obtenerSpreadsheet().getSheetByName('Vehículos');
        const datos = hoja.getDataRange().getValues();
        let filaVehiculo = -1;

        for (let i = 1; i < datos.length; i++) {
            if (datos[i][0] != vehiculoData.id && (datos[i][1] || '').toUpperCase() === patenteFormateada) throw new Error('Patente registrada en otro vehículo.');
            if (datos[i][0] == vehiculoData.id) {
                if ((datos[i][6] || '').toLowerCase() === 'asignado') throw new Error('No se puede editar vehículo asignado.');
                filaVehiculo = i + 1;
            }
        }
        if (filaVehiculo === -1) throw new Error('Vehículo no encontrado.');
        if (!vehiculoData.fechaVtoRto || !vehiculoData.fechaVtoSeguro) throw new Error('Fechas obligatorias.');

        const FilaActualizada = [ vehiculoData.id, patenteFormateada, vehiculoData.marca, vehiculoData.modelo, Number(vehiculoData.anio), Number(vehiculoData.capacidad), datos[filaVehiculo-1][6], vehiculoData.observaciones || '', vehiculoData.combustible || '', Number(vehiculoData.velocidadMaxima) || 0, new Date(vehiculoData.fechaVtoRto), new Date(vehiculoData.fechaVtoSeguro) ];
        hoja.getRange(filaVehiculo, 1, 1, FilaActualizada.length).setValues([FilaActualizada]);
        return { success: true, message: 'Vehículo actualizado.' };
    } finally { lock.releaseLock(); }
}

function obtenerNotificacionesVencimientos() {
  if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
  const hoja = obtenerSpreadsheet().getSheetByName('Vehículos');
  if (hoja.getLastRow() < 2) return [];
  const notificaciones = [];
  const hoy = new Date();
  const MS_PER_DAY = 1000 * 60 * 60 * 24;

  hoja.getDataRange().getValues().slice(1).forEach(row => {
    const patente = row[1];
    const fechaRto = new Date(row[10]);
    const fechaSeguro = new Date(row[11]);
    if (!isNaN(fechaRto.getTime()) && fechaRto >= hoy && fechaRto <= new Date(hoy.getTime() + 60 * MS_PER_DAY)) notificaciones.push({ patente: patente, tipo: 'RTO', fechaVencimiento: formatDateTime(fechaRto), diasRestantes: Math.round((fechaRto - hoy) / MS_PER_DAY) });
    if (!isNaN(fechaSeguro.getTime()) && fechaSeguro >= hoy && fechaSeguro <= new Date(hoy.getTime() + 30 * MS_PER_DAY)) notificaciones.push({ patente: patente, tipo: 'Seguro', fechaVencimiento: formatDateTime(fechaSeguro), diasRestantes: Math.round((fechaSeguro - hoy) / MS_PER_DAY) });
  });
  return notificaciones;
}

function actualizarEstadoVehiculo(patente, nuevoEstado) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
      if (!['disponible', 'no disponible', 'fuera de servicio'].includes(nuevoEstado.toLowerCase())) throw new Error('Estado no válido.');
      
      const hoja = obtenerSpreadsheet().getSheetByName('Vehículos');
      const datos = hoja.getDataRange().getValues();
      for (let i = 1; i < datos.length; i++) {
          if (datos[i][1].toLowerCase() === patente.toLowerCase()) {
              if (datos[i][6].toLowerCase() === 'asignado') throw new Error('Vehículo actualmente asignado.');
              hoja.getRange(i + 1, 7).setValue(nuevoEstado);
              return { success: true, message: `Estado actualizado a ${nuevoEstado}.` };
          }
      }
      throw new Error('Vehículo no encontrado.');
    } finally { lock.releaseLock(); }
}

function registrarMantenimiento(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
      const hojaMantenimientos = obtenerSpreadsheet().getSheetByName('Mantenimientos');
      const ultimaFila = hojaMantenimientos.getLastRow();
      const nuevaId = ultimaFila > 1 ? (parseInt(hojaMantenimientos.getRange(ultimaFila, 1).getValue(), 10) || 0) + 1 : 1;
      hojaMantenimientos.appendRow([ nuevaId, data.patente, data.tipo, data.detalle, Number(data.costo), data.responsable, new Date(data.fechaInicio), '', 'En curso' ]);
      return { success: true, message: 'Mantenimiento registrado.' };
    } finally { lock.releaseLock(); }
}

function finalizarMantenimiento(idMantenimiento) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
      const hojaMantenimientos = obtenerSpreadsheet().getSheetByName('Mantenimientos');
      const hojaVehiculos = obtenerSpreadsheet().getSheetByName('Vehículos');
      const datosMantenimientos = hojaMantenimientos.getDataRange().getValues();
      for (let i = 1; i < datosMantenimientos.length; i++) {
          if (datosMantenimientos[i][0] == idMantenimiento) {
              const patente = datosMantenimientos[i][1];
              hojaMantenimientos.getRange(i + 1, 9).setValue('Finalizado');
              hojaMantenimientos.getRange(i + 1, 8).setValue(new Date());
              const datosVehiculos = hojaVehiculos.getDataRange().getValues();
              for (let j = 1; j < datosVehiculos.length; j++) {
                  if (String(datosVehiculos[j][1]).toLowerCase() === patente.toLowerCase()) {
                      hojaVehiculos.getRange(j + 1, 7).setValue('Disponible');
                      break;
                  }
              }
              return { success: true, message: 'Mantenimiento finalizado.' };
          }
      }
      throw new Error('Mantenimiento no encontrado.');
    } finally { lock.releaseLock(); }
}

function mapRowToSolicitudObject(row) {
    return { id: row[0], emailUsuario: row[1], salida: formatDateTime(row[2]), regreso: formatDateTime(row[3]), itinerario: row[4], tareas: row[5], participantes: row[6], conductores: row[7], estado: row[8], vehiculoAsignado: row[9], observaciones: row[10], gestionadoPor: row[11] };
}

function formatDateTime(dateValue) {
    if (!dateValue) return 'N/A';
    let date;
    if (dateValue instanceof Date) date = dateValue;
    else if (typeof dateValue === 'string') {
        const parts = dateValue.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s(\d{2}):(\d{2}))?/);
        if (parts) date = new Date(+parts[3], parts[2] - 1, +parts[1], +(parts[4] || 0), +(parts[5] || 0));
        else date = new Date(dateValue);
    } else if (typeof dateValue === 'number') date = new Date(dateValue);
    else return 'Formato desconocido';

    if (isNaN(date.getTime())) return 'Fecha inválida';
    try {
        return date.toLocaleString('es-AR', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false }).replace(',', '');
    } catch (e) { return 'Error de formato'; }
}

function parsearFecha(fechaString) {
    if (!fechaString || typeof fechaString !== 'string') return new Date(fechaString);
    const partes = fechaString.match(/^(\d{2})\/(\d{2})\/(\d{4})(?: (\d{2}):(\d{2}))?$/);
    if (partes) return new Date(parseInt(partes[3], 10), parseInt(partes[2], 10) - 1, parseInt(partes[1], 10), partes[4] ? parseInt(partes[4], 10) : 0, partes[5] ? parseInt(partes[5], 10) : 0);
    return new Date(fechaString);
}

function actualizarEstadosVehiculosPorMantenimiento() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = obtenerSpreadsheet();
    const hojaMantenimientos = ss.getSheetByName('Mantenimientos');
    const hojaVehiculos = ss.getSheetByName('Vehículos');
    if(hojaMantenimientos.getLastRow() < 2) return; 
    
    const datosMantenimientos = hojaMantenimientos.getDataRange().getValues();
    const datosVehiculos = hojaVehiculos.getDataRange().getValues();
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0); 

    for (let i = 1; i < datosMantenimientos.length; i++) {
      if ((datosMantenimientos[i][8] || '').toLowerCase() === 'en curso') {
        const fechaInicio = new Date(datosMantenimientos[i][6]);
        fechaInicio.setHours(0, 0, 0, 0);
        if (fechaInicio <= hoy) {
          for (let j = 1; j < datosVehiculos.length; j++) {
            if (datosVehiculos[j][1] === datosMantenimientos[i][1] && (datosVehiculos[j][6] || '').toLowerCase() === 'disponible') {
              hojaVehiculos.getRange(j + 1, 7).setValue('Mantenimiento');
              break; 
            }
          }
        }
      }
    }
  } catch (e) { } finally { lock.releaseLock(); }
}

function crearTriggerDiario() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) if (trigger.getHandlerFunction() === 'actualizarEstadosVehiculosPorMantenimiento') ScriptApp.deleteTrigger(trigger);
  ScriptApp.newTrigger('actualizarEstadosVehiculosPorMantenimiento').timeBased().everyDays(1).atHour(1).create();
}

function obtenerChecklistEstandar() {
  return {
    motor_y_neumaticos: { label: 'MOTOR Y NEUMÁTICOS', items: [{ nombre: 'Nivel de aceite', key: 'nivel_aceite' }, { nombre: 'Agua radiador', key: 'agua_radiador' }, { nombre: 'Agua limpiaparabrisas', key: 'agua_limpiaparabrisas' }, { nombre: 'Batería adicional', key: 'bateria_adicional' }, { nombre: 'Condición neumáticos', key: 'neumaticos' }] },
    documentacion: { label: 'DOCUMENTACIÓN', items: [{ nombre: 'Cédula verde', key: 'cedula_verde' }, { nombre: 'Seguro', key: 'seguro' }, { nombre: 'Patente', key: 'patente' }, { nombre: 'Inspección Técnica sin vencer', key: 'inspeccion_tecnica' }, { nombre: 'Licencia para conducir vigente', key: 'licencia' }, { nombre: 'Permiso interno', key: 'permiso_interno' }] },
    equipo_de_seguridad: { label: 'EQUIPO DE SEGURIDAD', items: [{ nombre: 'Balizas Portátiles', key: 'balizas' }, { nombre: 'Matafuego', key: 'matafuego' }, { nombre: 'Botiquín de primeros auxilios', key: 'botiquin' }, { nombre: 'Radio BLU / VHF – Celular', key: 'radio' }, { nombre: 'Cinturones de seguridad', key: 'cinturones' }, { nombre: 'Apoya cabeza', key: 'apoya_cabeza' }, { nombre: 'Barra antivuelco externa', key: 'barra_antivuelco' }, { nombre: 'GPS', key: 'gps' }, { nombre: 'Chalecos Fluorescentes', key: 'chalecos' }] },
    cabina: { label: 'CABINA', items: [{ nombre: 'Espejos', key: 'espejos' }, { nombre: 'Luces altas y bajas', key: 'luces_altas_bajas' }, { nombre: 'Luces de posición / luz de stop', key: 'luces_posicion_stop' }, { nombre: 'Guiños y balizas', key: 'guinos_balizas' }, { nombre: 'Bocina', key: 'bocina' }, { nombre: 'Limpia parabrisas', key: 'limpia_parabrisas' }] },
    equipo_auxiliar: { label: 'EQUIPO AUXILIAR', items: [{ nombre: 'Crique', key: 'crique' }, { nombre: 'Llave de ruedas', key: 'llave_ruedas' }, { nombre: 'Rueda de Auxilio', key: 'rueda_auxilio' }, { nombre: 'Cadenas para hielo', key: 'cadenas' }, { nombre: 'Cables auxiliar de batería', key: 'cables_bateria' }, { nombre: 'Barra / eslinga de remolque', key: 'barra_remolque' }, { nombre: 'Tapa de tanque', key: 'tapa_tanque' }, { nombre: 'Caja de herramientas', key: 'caja_herramientas' }, { nombre: 'Linterna', key: 'linterna' }, { nombre: 'Orden y limpieza interior y exterior', key: 'orden_limpieza' }] },
    tanque_combustible: { label: 'TANQUE DE COMBUSTIBLE', items: [{ nombre: 'Nivel de combustible', key: 'nivel_combustible', opciones: ['RESERVA', '1/4', '1/2', '3/4', '4/4'] }] }
  };
}

function registrarControl(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!data.idSolicitud || !data.tipo || !data.items || !data.items.length) throw new Error('Datos inválidos.');
    const hojaControles = obtenerSpreadsheet().getSheetByName('Controles');
    for (let i = 0; i < data.items.length; i++) {
      let valorStr = String(data.items[i].valor || '');
      if (valorStr.match(/^\d+\/\d+$/)) valorStr = "'" + valorStr; 
      hojaControles.appendRow([ null, data.idSolicitud, data.tipo, data.items[i].seccion, data.items[i].item, valorStr, new Date(), getCurrentUserEmail() ]);
    }
    return { success: true, message: `Control registrado.` };
  } finally { lock.releaseLock(); }
}

function obtenerControlesPorSolicitud(idSolicitud) {
  const datos = obtenerSpreadsheet().getSheetByName('Controles').getDataRange().getValues();
  const inicial = [], final = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] == idSolicitud) {
      const registro = { seccion: datos[i][3], item: datos[i][4], valor: datos[i][5], fecha: datos[i][6] };
      if (datos[i][2] === 'inicial') inicial.push(registro);
      else if (datos[i][2] === 'final') final.push(registro);
    }
  }
  return { inicial: inicial, final: final };
}

function generarHtmlControl(idSolicitud, datosSolicitud, controles) {
  const ini = controles.inicial || [], fin = controles.final || [];
  const getValor = (lista, item) => { for (const r of lista) if (r.item === item) return String(r.valor || ''); return '-'; };
  const renderFila = (item, label) => `<tr><td>${label}</td><td>${getValor(ini, item)}</td><td>${label}</td><td>${getValor(fin, item)}</td></tr>`;
  const renderSeccion = (items, titulo) => `<tr><td colspan="4" style="background:#ddd;font-weight:bold;">${titulo}</td></tr>` + items.map(i => renderFila(i.key, i.label)).join('');
  
  const secciones = [
    {key: 'motor', label: 'MOTOR Y NEUMÁTICOS', items: [{key: 'nivel_aceite', label: 'Nivel de aceite'}, {key: 'agua_radiador', label: 'Agua radiador'}, {key: 'agua_limpiaparabrisas', label: 'Agua limpiaparabrisas'}, {key: 'bateria_adicional', label: 'Batería adicional'}, {key: 'neumaticos', label: 'Condición neumáticos'}]},
    {key: 'doc', label: 'DOCUMENTACIÓN', items: [{key: 'cedula_verde', label: 'Cédula verde'}, {key: 'seguro', label: 'Seguro'}, {key: 'patente', label: 'Patente'}, {key: 'inspeccion_tecnica', label: 'Inspección Técnica'}, {key: 'licencia', label: 'Licencia conducir'}, {key: 'permiso_interno', label: 'Permiso interno'}]},
    {key: 'seg', label: 'EQUIPO DE SEGURIDAD', items: [{key: 'balizas', label: 'Balizas Portátiles'}, {key: 'matafuego', label: 'Matafuego'}, {key: 'botiquin', label: 'Botiquín'}, {key: 'radio', label: 'Radio/VHF'}, {key: 'cinturones', label: 'Cinturones'}, {key: 'apoya_cabeza', label: 'Apoya cabeza'}, {key: 'barra_antivuelco', label: 'Barra antivuelco'}, {key: 'gps', label: 'GPS'}, {key: 'chalecos', label: 'Chalecos'}]},
    {key: 'cab', label: 'CABINA', items: [{key: 'espejos', label: 'Espejos'}, {key: 'luces_altas_bajas', label: 'Luces altas/bajas'}, {key: 'luces_posicion_stop', label: 'Luces posición/stop'}, {key: 'guinos_balizas', label: 'Guiños/balizas'}, {key: 'bocina', label: 'Bocina'}, {key: 'limpia_parabrisas', label: 'Limpia parabrisas'}]},
    {key: 'aux', label: 'EQUIPO AUXILIAR', items: [{key: 'crique', label: 'Crique'}, {key: 'llave_ruedas', label: 'Llave de ruedas'}, {key: 'rueda_auxilio', label: 'Rueda Auxilio'}, {key: 'cadenas', label: 'Cadenas hielo'}, {key: 'cables_bateria', label: 'Cables batería'}, {key: 'barra_remolque', label: 'Barra/eslinga'}, {key: 'tapa_tanque', label: 'Tapa tanque'}, {key: 'caja_herramientas', label: 'Caja herramientas'}, {key: 'linterna', label: 'Linterna'}, {key: 'orden_limpieza', label: 'Orden y limpieza'}]},
    {key: 'comb', label: 'TANQUE DE COMBUSTIBLE', items: [{key: 'nivel_combustible', label: 'Nivel de combustible'}]}
  ];
  let rows = secciones.map(sec => renderSeccion(sec.items, sec.label)).join('');
  
  return `<!DOCTYPE html><html><head><style>@page { size: A4; margin: 0.5cm; } @media print { body { margin: 0; padding: 5px; font-size: 10px; } .no-print { display: none !important; } } body { font-family: Arial, sans-serif; margin: 10px; font-size: 10px; } table { width: 100%; border-collapse: collapse; margin-bottom: 6px; font-size: 9px; } th, td { border: 1px solid #000; padding: 2px 4px; text-align: left; } th { background: #eee; } .footer { margin-top: 15px; display: flex; justify-content: space-between; } .firma-box { width: 45%; } .firma-line { border-bottom: 1px solid #000; height: 25px; } .btn { background: #B71C1C; color: white; padding: 8px 16px; border: none; cursor: pointer; margin: 10px; font-size: 12px; }</style></head><body><div class="no-print"><button class="btn" onclick="window.print()">Imprimir</button><button class="btn" onclick="window.close()">Cerrar</button></div><h1>PLANILLA CONTROL VEHÍCULAR</h1><div class="header"><strong>Solicitud N°:</strong> ${idSolicitud} | <strong>Vehículo:</strong> ${datosSolicitud.vehiculoAsignado || ''} | <strong>Solicitante:</strong> ${datosSolicitud.emailUsuario}</div><table><tr><th style="width:25%">ÍTEM</th><th style="width:20%">ENTREGA</th><th style="width:25%">ÍTEM</th><th style="width:20%">DEVOLUCIÓN</th></tr>${rows}</table><div class="footer"><div class="firma-box"><div class="firma-label">Firma Entrega:</div><div class="firma-line"></div></div><div class="firma-box"><div class="firma-label">Firma Recibe:</div><div class="firma-line"></div></div></div></body></html>`;
}

function obtenerDatosControlPorSolicitud(idSolicitud) {
  const controles = obtenerControlesPorSolicitud(idSolicitud);
  const agrupar = (items) => {
    const porSeccion = {};
    for (const item of items) {
      if (!porSeccion[item.seccion]) porSeccion[item.seccion] = [];
      porSeccion[item.seccion].push(item);
    }
    return porSeccion;
  };
  return { inicial: agrupar(controles.inicial), final: agrupar(controles.final) };
}
