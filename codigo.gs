/**
 * =================================================================
 * CONSTANTES Y CONFIGURACIÓN
 * =================================================================
 */
// --- ACCIÓN REQUERIDA ---
// Pegar el ID de la spreadsheet que se usa como DB.
const SPREADSHEET_ID = '1YFeFzAuTC-1vukBVunyZxTMLu35QNyf4j6_13wS2A3g';

// Para restringir el dominio, con workspace no va a hacer falta porque se hace desde la configuracion de implementacionm.
const DOMINIO_REQUERIDO = '@gmail.com';

const ENCABEZADOS_HOJAS = {
  'Solicitudes': ['ID', 'Email Solicitante', 'Fecha Salida', 'Fecha Regreso', 'Itinerario', 'Tareas', 'Participantes', 'Conductores', 'Estado', 'Vehículo Asignado', 'Observaciones', 'Gestionado por'],
  'Vehículos': ['ID', 'Patente', 'Marca', 'Modelo', 'Año', 'Capacidad', 'Estado', 'Observaciones', 'Combustible', 'Velocidad Máxima','Fecha VTO RTO', 'Fecha VTO Seguro'],
  'Mantenimientos': ['ID', 'Patente', 'Tipo', 'Detalle', 'Costo', 'Responsable', 'Fecha Inicio', 'Fecha Fin', 'Estado'],
  'Configuración': ['Rol', 'Email']
};


/**
 * =================================================================
 * FUNCIONES DE INFRAESTRUCTURA Y AUTENTICACIÓN
 * =================================================================
 */

// Esta función devuelve la URL base de la aplicación.
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// Función principal que se ejecuta al cargar la aplicación web.
function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail();

  // Lógica para la página de impresión de comprobantes
  if (e.parameter.page === 'imprimir' && e.parameter.id) {
    if (esEncargado(userEmail) || esSolicitante(userEmail)) { 
      const idSolicitud = e.parameter.id;
      const datosSolicitud = obtenerDatosSolicitud(idSolicitud);
      if (datosSolicitud) {
        // Un solicitante solo puede ver sus propios comprobantes
        if (!esEncargado(userEmail) && datosSolicitud.emailUsuario !== userEmail) {
           return HtmlService.createHtmlOutput('No tiene permiso para ver este comprobante.');
        }
        const template = HtmlService.createTemplateFromFile('comprobante');
        Object.keys(datosSolicitud).forEach(key => {
          template[key] = datosSolicitud[key];
        });
        return template.evaluate().setTitle(`Comprobante Solicitud N° ${idSolicitud}`);
      }
    }
    return HtmlService.createHtmlOutput('No tiene permiso para ver este comprobante o la solicitud no existe.');
  }

  // Lógica principal: decide qué panel mostrar (Encargado o Solicitante)
  if (esEncargado(userEmail)) {
    return HtmlService.createTemplateFromFile('Encargado').evaluate().setTitle('Panel Encargado - Flota UNRN');
  } else if (esSolicitante(userEmail)) {
    return HtmlService.createTemplateFromFile('Solicitante').evaluate().setTitle('Panel Solicitante - Flota UNRN');
  } else {
    // Si el email del usuario no cumple ninguna condición
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
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    Object.keys(ENCABEZADOS_HOJAS).forEach(nombre => {
      let hoja = ss.getSheetByName(nombre);
      if (!hoja) {
        hoja = ss.insertSheet(nombre);
        if (hoja.getLastRow() === 0) {
          hoja.appendRow(ENCABEZADOS_HOJAS[nombre]);
        }
      }
    });
    return ss;
  } catch (e) {
    Logger.log("Error al abrir Spreadsheet por ID. Verifica que el ID sea correcto y tengas permisos. Error: " + e.message);
    return null;
  }
}

function obtenerEncargados() {
  try {
    const hojaConfig = obtenerSpreadsheet().getSheetByName('Configuración');
    if (hojaConfig.getLastRow() < 2) return [];
    const datos = hojaConfig.getRange('B2:B' + hojaConfig.getLastRow()).getValues();
    return datos.map(row => row[0]).filter(email => email && email.trim() !== '');
  } catch (e) {
    Logger.log("Error al leer la hoja de Configuración: " + e.message);
    return [];
  }
}

function esEncargado(email) {
  if (!email) return false;
  const listaEncargados = obtenerEncargados();
  return listaEncargados.includes(email);
}

function esSolicitante(email) {
  if (!email) return false;
  return !esEncargado(email);
}

function getCurrentUserEmail() {
    return Session.getActiveUser().getEmail();
}


/**
 * =================================================================
 * FUNCIONES DE LECTURA (Dashboard, Listas, etc.)
 * =================================================================
 */

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
  datosSolicitudes.forEach(row => {
    if ((row[0] || '').toLowerCase() === 'pendiente') estadisticas.solicitudesPendientes++;
  });
  return estadisticas;
}

/* function verificarDisponibilidadVehiculo(patente, fechaSalida, fechaRegreso, idSolicitudExcluida = null) {
  const hoja = obtenerSpreadsheet().getSheetByName('Solicitudes');
  if (hoja.getLastRow() < 2) return true;
  const datos = hoja.getDataRange().getValues().slice(1);
  const solicitudesAprobadas = datos.filter(row => (row[8] || '').toLowerCase() === 'aprobada' && (row[9] || '').toUpperCase().startsWith(patente.toUpperCase()) && row[0] != idSolicitudExcluida);
  const nuevaSalidaTime = new Date(fechaSalida).getTime();
  const nuevoRegresoTime = new Date(fechaRegreso).getTime();
  for (const sol of solicitudesAprobadas) {
    const salidaExistente = new Date(sol[2]).getTime();
    const regresoExistente = new Date(sol[3]).getTime();
    if (nuevaSalidaTime < regresoExistente && nuevoRegresoTime > salidaExistente) return false;
  }
  return true;
} */


/**
 * Obtiene los mantenimientos de la hoja 'Mantenimientos'.
 * @returns {Array} Un array de objetos con los mantenimientos.
 */
function obtenerMantenimientos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMantenimientos = ss.getSheetByName('Mantenimientos');
  const datos = hojaMantenimientos.getDataRange().getValues();
  const cabecera = datos.shift(); // Quita la cabecera

  const mantenimientos = datos.map(fila => {
    const obj = {};
    cabecera.forEach((nombreColumna, i) => {
      obj[nombreColumna.trim().replace(/\s/g, '')] = fila[i]; // Crea un objeto con nombres de columna sin espacios
    });
    return obj;
  });

  return mantenimientos;
}

// // Ahora también verifica conflictos con mantenimientos programados, incluyendo un búfer de 2 días.
// function verificarDisponibilidadVehiculo(patente, fechaSalida, fechaRegreso, idSolicitudExcluida = null) {
//   const ss = obtenerSpreadsheet();
//   const hojaSolicitudes = ss.getSheetByName('Solicitudes');
//   const nuevoRegreso = new Date(fechaRegreso);

//   // 1. Verificar conflictos con otras solicitudes aprobadas
//   if (hojaSolicitudes.getLastRow() >= 2) {
//     const datosSolicitudes = hojaSolicitudes.getDataRange().getValues().slice(1);
//     const solicitudesAprobadas = datosSolicitudes.filter(row => 
//         (row[8] || '').toLowerCase() === 'aprobada' && 
//         (row[9] || '').toUpperCase().startsWith(patente.toUpperCase()) && 
//         row[0] != idSolicitudExcluida
//     );
//     const nuevoSalidaTime = new Date(fechaSalida).getTime();
//     for (const sol of solicitudesAprobadas) {
//       const salidaExistente = new Date(sol[2]).getTime();
//       const regresoExistente = new Date(sol[3]).getTime();
//       if (nuevoSalidaTime < regresoExistente && nuevoRegreso.getTime() > salidaExistente) {
//         throw new Error(`Conflicto de fechas. El vehículo ya está asignado a la solicitud #${sol[0]} en ese período.`);
//       }
//     }
//   }
//  // backup 05/09/2025
//   // 2. Verificar conflictos con mantenimientos programados
//   const hojaMantenimientos = ss.getSheetByName('Mantenimientos');
//   if (hojaMantenimientos.getLastRow() >= 2) {
//       const datosMantenimientos = hojaMantenimientos.getDataRange().getValues().slice(1);
//       const mantenimientosProgramados = datosMantenimientos.filter(row => 
//           (row[1] || '').toUpperCase() === patente.toUpperCase() &&
//           (row[8] || '').toLowerCase() === 'en curso'
//       );
//       for (const mant of mantenimientosProgramados) {
//           const inicioMantenimiento = new Date(mant[6]);
//           if (isNaN(inicioMantenimiento.getTime())) continue; // Omite fechas inválidas

//           // --- LÓGICA CORREGIDA ---
//           // Normalizamos las fechas a medianoche para una comparación por días.
//           const nuevoRegresoNorm = new Date(nuevoRegreso);
//           nuevoRegresoNorm.setHours(0, 0, 0, 0);

//           const inicioMantenimientoNorm = new Date(inicioMantenimiento);
//           inicioMantenimientoNorm.setHours(0, 0, 0, 0);

//           // Se establece un búfer de 2 días. El vehículo debe regresar ANTES del día que es 2 días antes del mantenimiento.
//           // Ejemplo: Mantenimiento el día 10. El búfer empieza el día 8 (10 - 2). El vehículo debe ser devuelto el 7 como máximo.
//           const dosDiasEnMs = 2 * 24 * 60 * 60 * 1000;
//           const primerDiaNoDisponible = new Date(inicioMantenimientoNorm.getTime() - dosDiasEnMs);

//           // Si la fecha de regreso (normalizada) es igual o posterior al primer día no disponible, hay conflicto.
//           if (nuevoRegresoNorm >= primerDiaNoDisponible) {
//               throw new Error(`Conflicto de mantenimiento. El vehículo debe estar disponible 2 días antes del turno programado para el ${formatDateTime(inicioMantenimiento)}.`);
//           }
//       }
//   }
  
//   return true; // Si no hay conflictos, retorna true.
// }


// function verificarDisponibilidadVehiculo(patente, fechaSalida, fechaRegreso, idSolicitudExcluida = null) {
//   const ss = obtenerSpreadsheet();
//   const hojaSolicitudes = ss.getSheetByName('Solicitudes');
//   const nuevaSalida = new Date(fechaSalida);
//   const nuevoRegreso = new Date(fechaRegreso);

//   // 1. Verificar conflictos con otras solicitudes aprobadas
//   if (hojaSolicitudes.getLastRow() >= 2) {
//     const datosSolicitudes = hojaSolicitudes.getDataRange().getValues().slice(1);
//     const solicitudesAprobadas = datosSolicitudes.filter(row => 
//         (row[8] || '').toLowerCase() === 'aprobada' && 
//         (row[9] || '').toUpperCase().startsWith(patente.toUpperCase()) && 
//         row[0] != idSolicitudExcluida
//     );
//     const nuevoSalidaTime = new Date(fechaSalida).getTime();
//     for (const sol of solicitudesAprobadas) {
//       const salidaExistente = new Date(sol[2]).getTime();
//       const regresoExistente = new Date(sol[3]).getTime();
//       if (nuevoSalidaTime < regresoExistente && nuevoRegreso.getTime() > salidaExistente) {
//         throw new Error(`Conflicto de fechas. El vehículo ya está asignado a la solicitud #${sol[0]} en ese período.`);
//       }
//     }
//   }

//   // 2. Verificar conflictos con mantenimientos programados
//   const hojaMantenimientos = ss.getSheetByName('Mantenimientos');
//   if (hojaMantenimientos.getLastRow() >= 2) {
//       const datosMantenimientos = hojaMantenimientos.getDataRange().getValues().slice(1);
//       const mantenimientosProgramados = datosMantenimientos.filter(row => 
//           (row[1] || '').toUpperCase() === patente.toUpperCase() &&
//           (row[8] || '').toLowerCase() === 'en curso'
//       );
      
//       for (const mant of mantenimientosProgramados) {
//           const inicioMantenimiento = new Date(mant[6]);
//           if (isNaN(inicioMantenimiento.getTime())) continue; // Omite fechas inválidas

//           // Normalizar fechas a medianoche para comparación por días
//           const nuevaSalidaNorm = new Date(nuevaSalida);
//           nuevaSalidaNorm.setHours(0, 0, 0, 0);
          
//           const nuevoRegresoNorm = new Date(nuevoRegreso);
//           nuevoRegresoNorm.setHours(0, 0, 0, 0);

//           const inicioMantenimientoNorm = new Date(inicioMantenimiento);
//           inicioMantenimientoNorm.setHours(0, 0, 0, 0);

//           if (nuevaSalida < inicioMantenimiento && inicioMantenimiento < nuevoRegreso) {
//              throw new Error(`Conflicto de mantenimiento. El vehículo ${patente} tiene un mantenimiento programado entre las fechas de solicitud.`);
//           }
          
//           // LÓGICA CORREGIDA:
//           // El búfer de 2 días significa que el vehículo debe estar libre 2 días ANTES del mantenimiento
//           const dosDiasEnMs = 2 * 24 * 60 * 60 * 1000;
//           const fechaLimiteDevolucion = new Date(inicioMantenimientoNorm.getTime() - dosDiasEnMs);


          
//           // Si la fecha de regreso es DESPUÉS de la fecha límite de devolución, hay conflicto
//           if (nuevoRegresoNorm >= fechaLimiteDevolucion) {
//               // Calcular la fecha límite real para mostrar al usuario
//               const fechaLimiteParaMostrar = new Date(fechaLimiteDevolucion.getTime() - (24 * 60 * 60 * 1000));
              
//               throw new Error(`Conflicto de mantenimiento. El vehículo ${patente} tiene un mantenimiento programado para el ${formatDateTime(inicioMantenimiento)}. Para cumplir con el búfer de seguridad de 2 días, el vehículo debe ser devuelto como máximo el ${formatDateTime(fechaLimiteParaMostrar)}.`);
//           }
//       }
//   }
  
//   return true; // Si no hay conflictos, retorna true.
// }
function aprobarSolicitud(idSolicitud, patenteAsignada, fechaSalida, fechaRegreso) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');

    const hojaSolicitudes = obtenerSpreadsheet().getSheetByName('Solicitudes');
    const hojaVehiculos = obtenerSpreadsheet().getSheetByName('Vehículos');

    // Buscar la solicitud
    const datosSolicitudes = hojaSolicitudes.getDataRange().getValues();
    let filaSolicitud = -1;
    let solicitudActual;
    for (let i = 1; i < datosSolicitudes.length; i++) {
      if (datosSolicitudes[i][0] == idSolicitud) {
        solicitudActual = mapRowToSolicitudObject(datosSolicitudes[i]);
        filaSolicitud = i + 1;
        break;
      }
    }

    if (!solicitudActual) throw new Error('La solicitud no fue encontrada.');
    if (solicitudActual.estado.toLowerCase() !== 'pendiente') {
      throw new Error(`No se puede aprobar esta solicitud porque su estado actual es "${solicitudActual.estado}".`);
    }

    // Verificar que el vehículo existe y está disponible
    const datosVehiculos = hojaVehiculos.getDataRange().getValues();
    let filaVehiculo = -1;
    let modeloVehiculo = '';
    for (let i = 1; i < datosVehiculos.length; i++) {
      if (String(datosVehiculos[i][1]).toLowerCase() === String(patenteAsignada).toLowerCase()) {
        const estadoVehiculo = (datosVehiculos[i][6] || '').toLowerCase();
        if (estadoVehiculo !== 'disponible') {
          throw new Error(`El vehículo ${patenteAsignada} no se puede asignar porque su estado es "${estadoVehiculo}".`);
        }
        modeloVehiculo = datosVehiculos[i][3];
        filaVehiculo = i + 1;
        break;
      }
    }
    if (filaVehiculo === -1) throw new Error(`El vehículo con patente ${patenteAsignada} no fue encontrado.`);

    // Crear objetos de fecha válidos antes de la validación
    const fechaSalidaValidada = parsearFecha(fechaSalida);
    const fechaRegresoValidada = parsearFecha(fechaRegreso);
    if (isNaN(fechaSalidaValidada.getTime()) || isNaN(fechaRegresoValidada.getTime())) {
        throw new Error('Las fechas de salida o regreso no son válidas. Por favor, revisa el formato.');
    }

    // VERIFICACIÓN CRÍTICA: Comprobar disponibilidad incluyendo mantenimientos
    try {
      verificarDisponibilidadVehiculo(patenteAsignada, fechaSalidaValidada, fechaRegresoValidada, idSolicitud);
    } catch (error) {
      // Si hay conflicto con mantenimientos o solicitudes, lanzar error específico
      throw new Error(`No se puede aprobar la solicitud: ${error.message}`);
    }

    // Si llegamos aquí, todo está OK - proceder con la aprobación
    const descripcionVehiculo = modeloVehiculo ? `${patenteAsignada.toUpperCase()} - ${modeloVehiculo}` : patenteAsignada.toUpperCase();

    // Actualizar estado del vehículo y la solicitud
    hojaVehiculos.getRange(filaVehiculo, 7).setValue('Asignado');
    hojaSolicitudes.getRange(filaSolicitud, 9).setValue('Aprobada');
    hojaSolicitudes.getRange(filaSolicitud, 10).setValue(descripcionVehiculo);
    hojaSolicitudes.getRange(filaSolicitud, 12).setValue(getCurrentUserEmail());

    // Enviar email de confirmación
    const emailSolicitante = solicitudActual.emailUsuario;
    const asunto = `Solicitud de Vehículo #${idSolicitud} - APROBADA`;
    const cuerpo = `Hola,<br><br>Tu solicitud de vehículo fue <b>APROBADA.<b><br><br><b>Detalles:</b><br>- ID de Solicitud: ${idSolicitud}<br>- Vehículo Asignado: ${descripcionVehiculo}<br>- Fechas: Desde ${formatDateTime(fechaSalida)} hasta ${formatDateTime(fechaRegreso)}<br><br><br>Saludos,<br>Sistema de Flota Universitaria UNRN.<br>Este correo se genero automaticamente, por favor no responder al mismo.`;

    try {
      MailApp.sendEmail(emailSolicitante, asunto, "", {htmlBody: cuerpo});
    } catch (emailError) {
      Logger.log(`Error enviando email a ${emailSolicitante}: ${emailError.message}`);
      // No lanzamos error aquí para no interrumpir la aprobación
    }

    return { success: true, message: 'Solicitud aprobada correctamente.' };
  } finally {
    lock.releaseLock();
  }
}

function verificarDisponibilidadVehiculo(patente, fechaSalida, fechaRegreso, idSolicitudExcluida = null) {
  const ss = obtenerSpreadsheet();
  const hojaSolicitudes = ss.getSheetByName('Solicitudes');

  // 1. Verificar conflictos con otras solicitudes aprobadas
  if (hojaSolicitudes.getLastRow() >= 2) {
    const datosSolicitudes = hojaSolicitudes.getDataRange().getValues().slice(1);
    const solicitudesAprobadas = datosSolicitudes.filter(row =>
        (row[8] || '').toLowerCase() === 'aprobada' &&
        (row[9] || '').toUpperCase().startsWith(patente.toUpperCase()) &&
        row[0] != idSolicitudExcluida
    );
    for (const sol of solicitudesAprobadas) {
      const salidaExistente = new Date(sol[2]).getTime();
      const regresoExistente = new Date(sol[3]).getTime();
      if (fechaSalida.getTime() < regresoExistente && fechaRegreso.getTime() > salidaExistente) {
        throw new Error(`Conflicto de fechas. El vehículo ya está asignado a la solicitud #${sol[0]} en ese período.`);
      }
    }
  }

  // 2. Verificar conflictos con mantenimientos programados
  const hojaMantenimientos = ss.getSheetByName('Mantenimientos');
  if (hojaMantenimientos.getLastRow() >= 2) {
      const datosMantenimientos = hojaMantenimientos.getDataRange().getValues().slice(1);
      const mantenimientosProgramados = datosMantenimientos.filter(row =>
          (row[1] || '').toUpperCase() === patente.toUpperCase() &&
          (row[8] || '').toLowerCase() === 'en curso'
      );

      for (const mant of mantenimientosProgramados) {
          const inicioMantenimiento = parsearFecha(mant[6]);
          let finMantenimiento = parsearFecha(mant[7]);

          if (isNaN(inicioMantenimiento.getTime())) continue; // Saltar si la fecha de inicio es inválida

          // Lógica corregida: Si no hay fecha de fin, el mantenimiento es indefinido
          if (isNaN(finMantenimiento.getTime())) {
              finMantenimiento = new Date(8640000000000000); // Fecha muy lejana
          }

          // La lógica de solapamiento estándar
          if (fechaSalida.getTime() < finMantenimiento.getTime() && fechaRegreso.getTime() > inicioMantenimiento.getTime()) {
              throw new Error(`Conflicto de mantenimiento. El vehículo ${patente} tiene un mantenimiento programado o en curso que se solapa con el período de solicitud.`);
          }
      }
  }

  return true;
}
function obtenerMantenimientosActivosConFechas() {
  if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
  
  const ss = obtenerSpreadsheet();
  const hojaMantenimientos = ss.getSheetByName('Mantenimientos');
  
  if (hojaMantenimientos.getLastRow() < 2) return [];
  
  const datos = hojaMantenimientos.getDataRange().getValues().slice(1);
  const mantenimientosActivos = datos.filter(row => (row[8] || '').toLowerCase() === 'en curso');
  
  return mantenimientosActivos.map(row => ({
    id: row[0],
    patente: row[1],
    tipo: row[2],
    detalle: row[3],
    fechaInicio: row[6],
    fechaInicioFormateada: formatDateTime(row[6])
  }));
}

/**
 * Verifica si un vehículo específico puede ser asignado en un rango de fechas
 * Retorna información detallada sobre disponibilidad
 */
function verificarDisponibilidadEspecifica(patente, fechaSalida, fechaRegreso) {
  if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
  
  try {
    verificarDisponibilidadVehiculo(patente, fechaSalida, fechaRegreso);
    return { 
      disponible: true, 
      mensaje: 'Vehículo disponible para el período solicitado',
      conflictos: []
    };
  } catch (error) {
    return {
      disponible: false,
      mensaje: error.message,
      conflictos: [{
        tipo: error.message.includes('mantenimiento') ? 'mantenimiento' : 'solicitud',
        detalle: error.message
      }]
    };
  }
}


function listarSolicitudesPorEstados(estados) {
    if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
    const hoja = obtenerSpreadsheet().getSheetByName('Solicitudes');
    if (hoja.getLastRow() < 2) return [];
    const datos = hoja.getDataRange().getValues().slice(1);
    const estadosEnMinuscula = estados.map(e => e.toLowerCase());
    return datos.filter(row => estadosEnMinuscula.includes((row[8] || '').toLowerCase())).map(mapRowToSolicitudObject);
}

function listarSolicitudesUsuarioPorEstados(estados) {
    const emailUsuario = getCurrentUserEmail();
    if (!esSolicitante(emailUsuario)) throw new Error('No autorizado');
    const hoja = obtenerSpreadsheet().getSheetByName('Solicitudes');
    if (hoja.getLastRow() < 2) return [];
    const datos = hoja.getDataRange().getValues().slice(1);
    const estadosEnMinuscula = estados.map(e => e.toLowerCase());
    return datos.filter(row => row[1] === emailUsuario && estadosEnMinuscula.includes((row[8] || '').toLowerCase())).map(mapRowToSolicitudObject);
}

/* function listarVehiculos() {
  if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
  const hoja = obtenerSpreadsheet().getSheetByName('Vehículos');
  if (hoja.getLastRow() < 2) return [];
  const datos = hoja.getDataRange().getValues().slice(1);
  return datos.map(row => ({ id: row[0], patente: row[1], marca: row[2], modelo: row[3], anio: row[4], capacidad: row[5], estado: row[6], observaciones: row[7], combustible: row[8], velocidadMaxima: row[9], fechaVtoRto: row[10], fechaVtoSeguro: row[11]  }));
}
 */

function listarVehiculos() {
  if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
  
  const hoja = obtenerSpreadsheet().getSheetByName('Vehículos');
  if (hoja.getLastRow() < 2) return [];
  
  const datos = hoja.getDataRange().getValues().slice(1);
  const vehiculos = [];

  datos.forEach((row, index) => {
    // Si la fila está completamente vacía, la saltamos.
    if (row.every(cell => cell === "")) return;

    try {
      // Intentamos procesar la fila, si hay un error, se registrará y continuará con la siguiente.
      vehiculos.push({ 
        id: row[0], 
        patente: row[1], 
        marca: row[2], 
        modelo: row[3], 
        anio: row[4], 
        capacidad: row[5], 
        estado: row[6], 
        observaciones: row[7], 
        combustible: row[8], 
        velocidadMaxima: row[9],
        fechaVtoRto: formatDateTime(row[10]),
        fechaVtoSeguro: formatDateTime(row[11])
      });
    } catch (e) {
      // Esto registrará un error detallado en los logs de Apps Script sin detener la app.
      Logger.log(`Error al procesar la fila de vehículo N° ${index + 2}: ${e.message}. Datos de la fila: ${row.join(', ')}`);
    }
  });

  return vehiculos;
}

function listarMantenimientosActivos() {
    if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
    const hoja = obtenerSpreadsheet().getSheetByName('Mantenimientos');
    if (hoja.getLastRow() < 2) return [];
    const datos = hoja.getDataRange().getValues().slice(1);
    return datos.filter(row => (row[8] || '').toLowerCase() === 'en curso')
        .map(row => ({
            id: row[0], patente: row[1], tipo: row[2], detalle: row[3], costo: row[4],
            responsable: row[5], fechaInicio: formatDateTime(row[6])
        }));
}

function listarMantenimientosFinalizados() {
    if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
    const hoja = obtenerSpreadsheet().getSheetByName('Mantenimientos');
    if (hoja.getLastRow() < 2) return [];
    const datos = hoja.getDataRange().getValues().slice(1);
    return datos.filter(row => (row[8] || '').toLowerCase() === 'finalizado')
        .map(row => ({
            id: row[0], patente: row[1], tipo: row[2], detalle: row[3], costo: row[4],
            responsable: row[5], fechaInicio: formatDateTime(row[6]), fechaFin: formatDateTime(row[7])
        }));
}


/**
 * =================================================================
 * FUNCIONES DE ESCRITURA (CON VALIDACIONES MEJORADAS)
 * =================================================================
 */

function enviarSolicitud(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const emailUsuario = getCurrentUserEmail();
    if (!esSolicitante(emailUsuario)) throw new Error('Usuario no autorizado para enviar solicitudes');

    const itinerario = (formData.itinerario || '').trim();
    const tareas = (formData.tareas || '').trim();
    if (!itinerario || !tareas) throw new Error("El itinerario y las tareas a realizar son campos obligatorios.");
    
    const salida = new Date(formData.salida);
    const regreso = new Date(formData.regreso);
    if (regreso <= salida) throw new Error('La fecha de regreso debe ser posterior a la de salida.');
    if (salida < new Date()) throw new Error('La fecha de salida no puede ser en el pasado.');

    const hoja = obtenerSpreadsheet().getSheetByName('Solicitudes');
    const ultimaFila = hoja.getLastRow();
    const nuevaId = ultimaFila > 1 ? (parseInt(hoja.getRange(ultimaFila, 1).getValue(), 10) || 0) + 1 : 1;

    hoja.appendRow([ nuevaId, emailUsuario, salida, regreso, itinerario, tareas, formData.participantes, formData.conductores, 'Pendiente', '', '' ]);
    
    const encargados = obtenerEncargados();
    if (encargados.length > 0) {
      const asunto = `Nueva Solicitud de Vehículo #${nuevaId} - ${emailUsuario}`;
      const cuerpo = `Se ha registrado una nueva solicitud de vehículo:<br><br>- Solicitante: ${emailUsuario}<br>- ID de Solicitud: ${nuevaId}<br>- Fechas: Desde ${salida} hasta ${regreso}<br>- Itinerario: ${formData.itinerario}<br><br>Para gestionar esta solicitud, por favor ingresa al Panel del Encargado.`;
      MailApp.sendEmail(encargados.join(','), asunto, "", {htmlBody: cuerpo});
    }
    
    return { success: true, message: 'Solicitud enviada con éxito.' };
  } finally {
    lock.releaseLock();
  }
}

// function aprobarSolicitud(idSolicitud, patenteAsignada, fechaSalida, fechaRegreso) {
//   const lock = LockService.getScriptLock();
//   lock.waitLock(30000);
//   try {
//     if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
    
//     const hojaSolicitudes = obtenerSpreadsheet().getSheetByName('Solicitudes');
//     const hojaVehiculos = obtenerSpreadsheet().getSheetByName('Vehículos');
    
//     const datosSolicitudes = hojaSolicitudes.getDataRange().getValues();
//     let filaSolicitud = -1;
//     let solicitudActual;
//     for (let i = 1; i < datosSolicitudes.length; i++) {
//       if (datosSolicitudes[i][0] == idSolicitud) {
//         solicitudActual = mapRowToSolicitudObject(datosSolicitudes[i]);
//         filaSolicitud = i + 1;
//         break;
//       }
//     }

//     if (!solicitudActual) throw new Error('La solicitud no fue encontrada.');
//     if (solicitudActual.estado.toLowerCase() !== 'pendiente') {
//       throw new Error(`No se puede aprobar esta solicitud porque su estado actual es "${solicitudActual.estado}".`);
//     }

//     const datosVehiculos = hojaVehiculos.getDataRange().getValues();
//     let filaVehiculo = -1;
//     let modeloVehiculo = '';
//     for (let i = 1; i < datosVehiculos.length; i++) {
//       if (String(datosVehiculos[i][1]).toLowerCase() === String(patenteAsignada).toLowerCase()) {
//         const estadoVehiculo = (datosVehiculos[i][6] || '').toLowerCase();
//         if (estadoVehiculo !== 'disponible') {
//           throw new Error(`El vehículo ${patenteAsignada} no se puede asignar porque su estado es "${estadoVehiculo}".`);
//         }
//         modeloVehiculo = datosVehiculos[i][3];
//         filaVehiculo = i + 1;
//         break;
//       }
//     }
//     if (filaVehiculo === -1) throw new Error(`El vehículo con patente ${patenteAsignada} no fue encontrado.`);

//     if (!verificarDisponibilidadVehiculo(patenteAsignada, fechaSalida, fechaRegreso, idSolicitud)) {
//       throw new Error(`Conflicto de fechas. El vehículo ${patenteAsignada} ya está asignado en ese período.`);
//     }

//     const descripcionVehiculo = modeloVehiculo ? `${patenteAsignada.toUpperCase()} - ${modeloVehiculo}` : patenteAsignada.toUpperCase();
    
//     hojaVehiculos.getRange(filaVehiculo, 7).setValue('Asignado');
//     hojaSolicitudes.getRange(filaSolicitud, 9).setValue('Aprobada');
//     hojaSolicitudes.getRange(filaSolicitud, 10).setValue(descripcionVehiculo);
//     hojaSolicitudes.getRange(filaSolicitud,12).setValue(getCurrentUserEmail());

//     const emailSolicitante = solicitudActual.emailUsuario;
//     const asunto = `Respuesta a tu Solicitud de Vehículo #${idSolicitud}`;
//     const cuerpo = `Hola,<br><br>Tu solicitud de vehículo fue <b>APROBADA.<b><br><br><b>Detalles:</b><br>- ID de Solicitud: ${idSolicitud}<br>- Vehículo Asignado: ${descripcionVehiculo}<br>- Fechas: Desde ${formatDateTime(fechaSalida)} hasta ${formatDateTime(fechaRegreso)}<br><br><br>Saludos,<br>Sistema de Flota Universitaria UNRN.<br>Este correo se genero automaticamente, por favor no responder al mismo.`;
//     MailApp.sendEmail(emailSolicitante, asunto, "", {htmlBody: cuerpo});
    
//     return { success: true, message: 'Solicitud aprobada correctamente.' };
//   } finally {
//     lock.releaseLock();
//   }
// }



function rechazarSolicitud(idSolicitud, motivo) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
      
      const hoja = obtenerSpreadsheet().getSheetByName('Solicitudes');
      const datos = hoja.getDataRange().getValues();
      for (let i = 1; i < datos.length; i++) {
          if (datos[i][0] == idSolicitud) {
              const estadoActual = (datos[i][8] || '').toLowerCase();
              if (estadoActual !== 'pendiente') {
                throw new Error(`No se puede rechazar esta solicitud porque su estado actual es "${estadoActual}".`);
              }
              const emailSolicitante = datos[i][1];
              hoja.getRange(i + 1, 9).setValue('Rechazada');
              hoja.getRange(i + 1, 11).setValue(motivo);
              hoja.getRange(i + 1, 10).setValue('');
              hoja.getRange(i +1,12).setValue(getCurrentUserEmail());
              const asunto = `Respuesta a tu Solicitud de Vehículo #${idSolicitud}`;
              const cuerpo = `Hola,<br><br>Lamentamos informarte que tu solicitud de vehículo ha sido rechazada.<br><br>- ID de Solicitud: ${idSolicitud},<br>- Motivo del rechazo: ${motivo}<br><br><br>Si tienes alguna consulta, por favor, contacta a la oficina de flota.<br><br>Saludos,<br>Oficina de Flota.`;
              MailApp.sendEmail(emailSolicitante, asunto, "", {htmlBody: cuerpo});
              return { success: true, message: 'Solicitud rechazada.' };
          }
      }
      throw new Error('Solicitud no encontrada.');
    } finally {
      lock.releaseLock();
    }
}

function finalizarSolicitud(idSolicitud) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
      
      const hojaSolicitudes = obtenerSpreadsheet().getSheetByName('Solicitudes');
      const datosSolicitudes = hojaSolicitudes.getDataRange().getValues();
      
      let filaSolicitud = -1;
      let vehiculoAsignado = '';
      for (let i = 1; i < datosSolicitudes.length; i++) {
        if (datosSolicitudes[i][0] == idSolicitud) {
          const estadoActual = (datosSolicitudes[i][8] || '').toLowerCase();
          if (estadoActual !== 'aprobada') {
            throw new Error(`No se puede finalizar una solicitud que no está en estado "Aprobada". Estado actual: ${estadoActual}`);
          }
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
      
      return { success: true, message: 'Solicitud finalizada y vehículo liberado.' };
    } finally {
      lock.releaseLock();
    }
}

/* function agregarVehiculo(vehiculoData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');

    const hoja = obtenerSpreadsheet().getSheetByName('Vehículos');
    const ultimaFila = hoja.getLastRow();
    const nuevaId = ultimaFila > 1 ? (parseInt(hoja.getRange(ultimaFila, 1).getValue(), 10) || 0) + 1 : 1;

    hoja.appendRow([
      nuevaId, vehiculoData.patente.toUpperCase(), vehiculoData.marca, vehiculoData.modelo,
      Number(vehiculoData.anio), Number(vehiculoData.capacidad), 'Disponible', vehiculoData.observaciones || '',
      vehiculoData.combustible || '', Number(vehiculoData.velocidadMaxima) || 0
    ]);
    return { success: true, message: 'Vehículo agregado.' };
  } finally {
    lock.releaseLock();
  }
}
 */


function agregarVehiculo(vehiculoData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');

      if (!vehiculoData.fechaVtoRto || !vehiculoData.fechaVtoSeguro) {
        throw new Error('Las fechas de vencimiento de RTO y Seguro son obligatorias.');
      }
      
      // --- INICIO DE VALIDACIÓN DE PATENTE ---
      const patenteFormateada = validarYFormatearPatente(vehiculoData.patente);
      
      const hoja = obtenerSpreadsheet().getSheetByName('Vehículos');
      const datosVehiculos = hoja.getLastRow() > 1 ? hoja.getRange(2, 2, hoja.getLastRow() - 1, 1).getValues() : [];
      const patentesExistentes = datosVehiculos.flat().map(p => (p || '').toUpperCase());

      if (patentesExistentes.includes(patenteFormateada)) {
          throw new Error(`La patente "${patenteFormateada}" ya está registrada.`);
      }
      // --- FIN VALIDACIÓN ---

      const ultimaFila = hoja.getLastRow();
      const nuevaId = ultimaFila > 1 ? (parseInt(hoja.getRange(ultimaFila, 1).getValue(), 10) || 0) + 1 : 1;
      
      hoja.appendRow([
          nuevaId, patenteFormateada, vehiculoData.marca, vehiculoData.modelo, 
          Number(vehiculoData.anio), Number(vehiculoData.capacidad), 'Disponible', vehiculoData.observaciones || '',
          vehiculoData.combustible || '', Number(vehiculoData.velocidadMaxima) || 0,
          new Date(vehiculoData.fechaVtoRto), new Date(vehiculoData.fechaVtoSeguro)
      ]);
      return { success: true, message: 'Vehículo agregado.' };
    } finally {
      lock.releaseLock();
    }
}

/**
 * Valida y formatea la patente.
 * @param {string} patente La patente a validar.
 * @returns {string} La patente formateada en mayúsculas.
 * @throws {Error} Si la patente está vacía o tiene un formato no válido.
 */
function validarYFormatearPatente(patente) {
  if (!patente || typeof patente !== 'string' || patente.trim() === '') {
    throw new Error('La patente es obligatoria.');
  }

  const patenteUpper = patente.trim().toUpperCase();

  // Expresión regular para los dos formatos: AAA 123 y AA 123 AA
  const regexFormato = /^(?:[A-Z]{3}\s\d{3}|[A-Z]{2}\s\d{3}\s[A-Z]{2})$/;

  if (!regexFormato.test(patenteUpper)) {
    throw new Error('Formato de patente no válido. Use "AAA 123" o "AA 123 AA" (mayúsculas y con espacios).');
  }

  return patenteUpper;
}


// --- NUEVA FUNCIÓN ---
function editarVehiculo(vehiculoData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');

        // --- INICIO DE VALIDACIÓN DE PATENTE ---
        const patenteFormateada = validarYFormatearPatente(vehiculoData.patente);

        const hoja = obtenerSpreadsheet().getSheetByName('Vehículos');
        const datos = hoja.getDataRange().getValues();
        let filaVehiculo = -1;

        for (let i = 1; i < datos.length; i++) {
            // Verifica que la patente no esté duplicada en OTRO vehículo
            if (datos[i][0] != vehiculoData.id && (datos[i][1] || '').toUpperCase() === patenteFormateada) {
                throw new Error(`La patente "${patenteFormateada}" ya está registrada en otro vehículo.`);
            }
            // Encuentra la fila del vehículo que se está editando
            if (datos[i][0] == vehiculoData.id) {
                if ((datos[i][6] || '').toLowerCase() === 'asignado') {
                    throw new Error('No se puede editar un vehículo que está actualmente asignado.');
                }
                filaVehiculo = i + 1;
            }
        }
        // --- FIN VALIDACIÓN ---

        if (filaVehiculo === -1) throw new Error('Vehículo no encontrado.');
        
        if (!vehiculoData.fechaVtoRto || !vehiculoData.fechaVtoSeguro) {
            throw new Error('Las fechas de vencimiento de RTO y Seguro son obligatorias.');
        }

        const FilaActualizada = [
            vehiculoData.id, 
            patenteFormateada,
            vehiculoData.marca,
            vehiculoData.modelo,
            Number(vehiculoData.anio),
            Number(vehiculoData.capacidad),
            datos[filaVehiculo-1][6],
            vehiculoData.observaciones || '',
            vehiculoData.combustible || '',
            Number(vehiculoData.velocidadMaxima) || 0,
            new Date(vehiculoData.fechaVtoRto),
            new Date(vehiculoData.fechaVtoSeguro)
        ];
        
        hoja.getRange(filaVehiculo, 1, 1, FilaActualizada.length).setValues([FilaActualizada]);

        return { success: true, message: 'Vehículo actualizado correctamente.' };

    } finally {
        lock.releaseLock();
    }
}


function obtenerNotificacionesVencimientos() {
  if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
  
  const hoja = obtenerSpreadsheet().getSheetByName('Vehículos');
  if (hoja.getLastRow() < 2) return [];
  
  const datos = hoja.getDataRange().getValues().slice(1);
  const notificaciones = [];
  const hoy = new Date();
  const MS_PER_DAY = 1000 * 60 * 60 * 24;

  const limiteRto = new Date(hoy.getTime() + 60 * MS_PER_DAY);
  const limiteSeguro = new Date(hoy.getTime() + 30 * MS_PER_DAY);

  datos.forEach(row => {
    const patente = row[1];
    const fechaRto = new Date(row[10]);
    const fechaSeguro = new Date(row[11]);

    if (!isNaN(fechaRto.getTime()) && fechaRto >= hoy && fechaRto <= limiteRto) {
      const diasRestantes = Math.round((fechaRto - hoy) / MS_PER_DAY);
      notificaciones.push({
        patente: patente,
        tipo: 'RTO',
        fechaVencimiento: formatDateTime(fechaRto),
        diasRestantes: diasRestantes
      });
    }

    if (!isNaN(fechaSeguro.getTime()) && fechaSeguro >= hoy && fechaSeguro <= limiteSeguro) {
      const diasRestantes = Math.round((fechaSeguro - hoy) / MS_PER_DAY);
      notificaciones.push({
        patente: patente,
        tipo: 'Seguro',
        fechaVencimiento: formatDateTime(fechaSeguro),
        diasRestantes: diasRestantes
      });
    }
  });

  return notificaciones;
}

function actualizarEstadoVehiculo(patente, nuevoEstado) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');
      
      const estadosValidos = ['disponible', 'no disponible', 'fuera de servicio'];
      if (!estadosValidos.includes(nuevoEstado.toLowerCase())) {
          throw new Error('Estado no válido.');
      }
      const hoja = obtenerSpreadsheet().getSheetByName('Vehículos');
      const datos = hoja.getDataRange().getValues();
      for (let i = 1; i < datos.length; i++) {
          if (datos[i][1].toLowerCase() === patente.toLowerCase()) {
              if (datos[i][6].toLowerCase() === 'asignado') {
                  throw new Error('No se puede cambiar el estado de un vehículo que está actualmente asignado a una solicitud activa.');
              }
              hoja.getRange(i + 1, 7).setValue(nuevoEstado);
              return { success: true, message: `Estado del vehículo ${patente} actualizado a ${nuevoEstado}.` };
          }
      }
      throw new Error('Vehículo no encontrado.');
    } finally {
      lock.releaseLock();
    }
}

/* function registrarMantenimiento(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');

      const hojaMantenimientos = obtenerSpreadsheet().getSheetByName('Mantenimientos');
      const hojaVehiculos = obtenerSpreadsheet().getSheetByName('Vehículos');
      
      const ultimaFila = hojaMantenimientos.getLastRow();
      const nuevaId = ultimaFila > 1 ? (parseInt(hojaMantenimientos.getRange(ultimaFila, 1).getValue(), 10) || 0) + 1 : 1;

      hojaMantenimientos.appendRow([
          nuevaId, data.patente, data.tipo, data.detalle, Number(data.costo),
          data.responsable, new Date(data.fechaInicio), '', 'En curso'
      ]);
      
      const datosVehiculos = hojaVehiculos.getDataRange().getValues();
      for (let i = 1; i < datosVehiculos.length; i++) {
          if (String(datosVehiculos[i][1]).toLowerCase() === String(data.patente).toLowerCase()) {
              hojaVehiculos.getRange(i + 1, 7).setValue('Mantenimiento');
              break;
          }
      }
      return { success: true, message: 'Mantenimiento registrado con éxito.' };
    } finally {
      lock.releaseLock();
    }
} */

// Ya no cambia el estado del vehículo. Solo registra el mantenimiento.
function registrarMantenimiento(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (!esEncargado(getCurrentUserEmail())) throw new Error('No autorizado');

      const hojaMantenimientos = obtenerSpreadsheet().getSheetByName('Mantenimientos');
      const ultimaFila = hojaMantenimientos.getLastRow();
      const nuevaId = ultimaFila > 1 ? (parseInt(hojaMantenimientos.getRange(ultimaFila, 1).getValue(), 10) || 0) + 1 : 1;

      hojaMantenimientos.appendRow([
          nuevaId, data.patente, data.tipo, data.detalle, Number(data.costo),
          data.responsable, new Date(data.fechaInicio), '', 'En curso'
      ]);
      
      return { success: true, message: 'Mantenimiento registrado con éxito. El estado del vehículo cambiará automáticamente en la fecha de inicio.' };
    } finally {
      lock.releaseLock();
    }
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
    } finally {
      lock.releaseLock();
    }
}


/**
 * =================================================================
 * FUNCIONES DE UTILIDAD Y MAPEO
 * =================================================================
 */

function mapRowToSolicitudObject(row) {
    return {
        id: row[0], emailUsuario: row[1], salida: formatDateTime(row[2]),
        regreso: formatDateTime(row[3]), itinerario: row[4], tareas: row[5],
        participantes: row[6], conductores: row[7], estado: row[8],
        vehiculoAsignado: row[9], observaciones: row[10],
        gestionadoPor: row[11] 
    };
}

function formatDateTime(dateValue) {
    // Si el valor está vacío o nulo, retorna 'N/A'
    if (!dateValue) return 'N/A';

    let date;

    // Si ya es un objeto Date, úsalo directamente.
    if (dateValue instanceof Date) {
        date = dateValue;
    } 
    // Si es un string, intenta analizarlo.
    else if (typeof dateValue === 'string') {
        // Intenta analizar el formato 'dd/MM/yyyy HH:mm:ss' o 'dd/MM/yyyy HH:mm'
        const parts = dateValue.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s(\d{2}):(\d{2}))?/);
        if (parts) {
            // new Date(año, mes-1, día, hora, minuto)
            const hour = parts[4] || 0;
            const minute = parts[5] || 0;
            date = new Date(+parts[3], parts[2] - 1, +parts[1], +hour, +minute);
        } else {
            // Si el formato no coincide, intenta un análisis genérico.
            date = new Date(dateValue);
        }
    } 
    // Si es un número (timestamp), crea la fecha.
    else if (typeof dateValue === 'number') {
        date = new Date(dateValue);
    } 
    // Si no es ninguno de los anteriores, no se puede procesar.
    else {
        return 'Formato desconocido';
    }

    // Verifica si el objeto Date resultante es válido.
    if (isNaN(date.getTime())) {
        return 'Fecha inválida';
    }

    // Si es válido, formatea y devuelve.
    try {
        return date.toLocaleString('es-AR', {
            year: 'numeric', month: '2-digit', day: '2-digit',
            hour: '2-digit', minute: '2-digit', hour12: false
        }).replace(',', '');
    } catch (e) {
        return 'Error de formato';
    }
}

function parsearFecha(fechaString) {
    if (!fechaString || typeof fechaString !== 'string') {
        return new Date(fechaString);
    }
    const partes = fechaString.match(/^(\d{2})\/(\d{2})\/(\d{4})(?: (\d{2}):(\d{2}))?$/);
    if (partes) {
        // new Date(año, mes-1, día, hora, minuto)
        const dia = parseInt(partes[1], 10);
        const mes = parseInt(partes[2], 10) - 1;
        const año = parseInt(partes[3], 10);
        const hora = partes[4] ? parseInt(partes[4], 10) : 0;
        const minuto = partes[5] ? parseInt(partes[5], 10) : 0;
        return new Date(año, mes, dia, hora, minuto);
    }
    // Si no coincide con DD/MM/YYYY, intenta el análisis genérico (ISO 8601)
    return new Date(fechaString);
}

/**
 * Esta función está diseñada para ser ejecutada por un activador diario.
 * Revisa los mantenimientos programados y actualiza el estado del vehículo a "Mantenimiento"
 * si la fecha de inicio del mantenimiento es hoy o ya pasó.
 */
function actualizarEstadosVehiculosPorMantenimiento() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = obtenerSpreadsheet();
    const hojaMantenimientos = ss.getSheetByName('Mantenimientos');
    const hojaVehiculos = ss.getSheetByName('Vehículos');

    if(hojaMantenimientos.getLastRow() < 2) return; // No hay mantenimientos
    
    const datosMantenimientos = hojaMantenimientos.getDataRange().getValues();
    const datosVehiculos = hojaVehiculos.getDataRange().getValues();

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0); // Normalizar a medianoche para comparar solo fechas

    // Itera desde la segunda fila para saltar encabezados
    for (let i = 1; i < datosMantenimientos.length; i++) {
      const estadoMantenimiento = (datosMantenimientos[i][8] || '').toLowerCase();
      
      if (estadoMantenimiento === 'En curso') {
        const fechaInicio = new Date(datosMantenimientos[i][6]);
        fechaInicio.setHours(0, 0, 0, 0);

        if (fechaInicio <= hoy) {
          const patenteMantenimiento = datosMantenimientos[i][1];
          // Buscar el vehículo y actualizar su estado si es "Disponible"
          for (let j = 1; j < datosVehiculos.length; j++) {
            if (datosVehiculos[j][1] === patenteMantenimiento && (datosVehiculos[j][6] || '').toLowerCase() === 'disponible') {
              hojaVehiculos.getRange(j + 1, 7).setValue('Mantenimiento');
              Logger.log(`Vehículo ${patenteMantenimiento} actualizado a "Mantenimiento".`);
              break; // Pasa al siguiente mantenimiento
            }
          }
        }
      }
    }
  } catch (e) {
      Logger.log(`Error en la actualización automática de estados: ${e.message}`);
  } finally {
      lock.releaseLock();
  }
}

/**
 * Función de utilidad para crear el activador diario. 
 * Ejecútala una sola vez desde el editor de Apps Script para configurar la automatización.
 */
function crearTriggerDiario() {
  // Elimina activadores antiguos para evitar duplicados
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'actualizarEstadosVehiculosPorMantenimiento') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  // Crea un nuevo activador que se ejecuta todos los días entre la 1 y 2 a.m.
  ScriptApp.newTrigger('actualizarEstadosVehiculosPorMantenimiento')
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .create();
  Logger.log('Activador diario creado para actualizarEstadosVehiculosPorMantenimiento.');
}

