// MOD-001: ENCABEZADO [INICIO]
/**
*****************************************
PROYECTO: Sistema de Asistencia
ARCHIVO: code.gs
VERSIÓN: 01.00
FECHA: 10/02/2026 15:18 (UTC-5)
*****************************************
*/
// MOD-001: FIN

// MOD-002: CONFIGURACIÓN [INICIO]
const SHEET_ID = '1GI5C5djzMEFCcQEewi-MotEBggQa5VMh0ylchHjV5kM';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('home')
    .setTitle('Sistema de Asistencia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// MOD-002: FIN

// MOD-003: REGISTRAR LLEGADA [INICIO]
function registrarLlegada() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetRegistro = ss.getSheetByName('Registro');
    
    const ahora = new Date();
    const fecha = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const hora = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'HH:mm');
    
    // Buscar si ya existe registro para hoy
    const datos = sheetRegistro.getDataRange().getValues();
    let filaExistente = -1;
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0]) {
        const fechaRegistro = Utilities.formatDate(new Date(datos[i][0]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
        if (fechaRegistro === fecha) {
          filaExistente = i + 1;
          break;
        }
      }
    }
    
    // Generar texto de llegada
    const textoLlegada = generarTextoLlegada(ahora);
    
    if (filaExistente > 0) {
      // Actualizar registro existente
      sheetRegistro.getRange(filaExistente, 2).setValue(hora); // Columna B
      sheetRegistro.getRange(filaExistente, 4).setValue(textoLlegada); // Columna D
    } else {
      // Crear nuevo registro (A a H)
      sheetRegistro.appendRow([fecha, hora, '', textoLlegada, '', '', '', '']);
    }
    
    return {
      success: true,
      texto: textoLlegada,
      mensaje: 'Llegada registrada correctamente'
    };
  } catch (error) {
    Logger.log('Error en registrarLlegada: ' + error.message);
    return {
      success: false,
      mensaje: 'No se pudo registrar la llegada. Por favor, verifica tu conexión e intenta nuevamente.'
    };
  }
}
// MOD-003: FIN

// MOD-004: REGISTRAR SALIDA [INICIO]
function registrarSalida() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetRegistro = ss.getSheetByName('Registro');
    const sheetTarifa = ss.getSheetByName('Tarifa');
    
    const ahora = new Date();
    const fecha = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const hora = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'HH:mm');
    
    // Buscar registro de hoy
    const datos = sheetRegistro.getDataRange().getValues();
    let filaExistente = -1;
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0]) {
        const fechaRegistro = Utilities.formatDate(new Date(datos[i][0]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
        if (fechaRegistro === fecha) {
          filaExistente = i + 1;
          break;
        }
      }
    }
    
    if (filaExistente <= 0) {
      return {
        success: false,
        mensaje: 'No hay registro de llegada para hoy'
      };
    }
    
    // Generar texto de salida
    const textoSalida = generarTextoSalida(ahora);
    
    // Registrar salida y texto
    sheetRegistro.getRange(filaExistente, 3).setValue(hora); // Columna C
    sheetRegistro.getRange(filaExistente, 5).setValue(textoSalida); // Columna E
    
    // Calcular Horas y Monto solo si hay hora de llegada válida
    const horaLlegada = datos[filaExistente - 1][1]; // Columna B
    
    if (horaLlegada) {
      try {
        // Calcular horas trabajadas (sin redondear aún)
        const horasTrabajadas = calcularHorasTrabajadas(horaLlegada, hora);
        
        // Obtener tarifas
        const tarifaHora = sheetTarifa.getRange('A2').getValue();
        const tarifaPasaje = sheetTarifa.getRange('B2').getValue();
        
        // Calcular monto con horas completas (sin redondear) y redondear solo el monto final
        const monto = (horasTrabajadas * tarifaHora) + tarifaPasaje;
        
        // Guardar en columnas F y G como NÚMEROS
        const cellHoras = sheetRegistro.getRange(filaExistente, 6);
        cellHoras.setValue(parseFloat(horasTrabajadas.toFixed(2)));
        cellHoras.setNumberFormat('0.00'); // Formato numérico con 2 decimales
        
        const cellMonto = sheetRegistro.getRange(filaExistente, 7);
        cellMonto.setValue(parseFloat(monto.toFixed(2))); // Redondear solo el monto final
        cellMonto.setNumberFormat('0.00'); // Formato numérico con 2 decimales
        
      } catch (error) {
        Logger.log('Error al calcular horas/monto: ' + error.message);
        // Continuar sin calcular si hay error
      }
    }
    
    return {
      success: true,
      texto: textoSalida,
      mensaje: 'Salida registrada correctamente'
    };
  } catch (error) {
    Logger.log('Error en registrarSalida: ' + error.message);
    return {
      success: false,
      mensaje: 'No se pudo registrar la salida. Por favor, verifica tu conexión e intenta nuevamente.'
    };
  }
}
// MOD-004: FIN

// MOD-005: OBTENER REGISTRO HOY [INICIO]
function obtenerRegistroHoy() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetRegistro = ss.getSheetByName('Registro');
    
    const ahora = new Date();
    const fecha = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    
    const datos = sheetRegistro.getDataRange().getValues();
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0]) {
        const fechaRegistro = Utilities.formatDate(new Date(datos[i][0]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
        if (fechaRegistro === fecha) {
          return {
            success: true,
            textoLlegada: datos[i][3] || '',
            textoSalida: datos[i][4] || ''
          };
        }
      }
    }
    
    return {
      success: true,
      textoLlegada: '',
      textoSalida: ''
    };
  } catch (error) {
    Logger.log('Error en obtenerRegistroHoy: ' + error.message);
    return {
      success: false,
      mensaje: 'No se pudo cargar el registro de hoy. Por favor, recarga la página.'
    };
  }
}
// MOD-005: FIN

// MOD-006: GENERAR REPORTE [INICIO]
function generarReporte() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetRegistro = ss.getSheetByName('Registro');
    
    // Obtener todos los registros
    const datos = sheetRegistro.getDataRange().getValues();
    
    // Procesar registros por semana
    const semanas = {};
    let totalGeneral = 0;
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][1] && datos[i][2]) {
        const fecha = new Date(datos[i][0]);
        const fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        const estado = datos[i][7]; // Columna H
        
        // Verificar si ya está pagado
        if (estado === 'Pagado') {
          continue;
        }
        
        const horaLlegada = datos[i][1];
        const horaSalida = datos[i][2];
        
        // Intentar leer monto de la columna G, si no existe calcularlo
        let monto;
        if (datos[i][6]) {
          monto = parseFloat(datos[i][6]);
        } else {
          // Calcular si no existe
          const sheetTarifa = ss.getSheetByName('Tarifa');
          const tarifaHora = sheetTarifa.getRange('A2').getValue();
          const tarifaPasaje = sheetTarifa.getRange('B2').getValue();
          const horasTrabajadas = calcularHorasTrabajadas(horaLlegada, horaSalida);
          monto = (horasTrabajadas * tarifaHora) + tarifaPasaje;
        }
        
        // Obtener semana
        const inicioSemana = obtenerInicioSemana(fecha);
        const finSemana = obtenerFinSemana(fecha);
        const claveEsemana = Utilities.formatDate(inicioSemana, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        
        if (!semanas[claveEsemana]) {
          semanas[claveEsemana] = {
            inicio: inicioSemana,
            fin: finSemana,
            dias: []
          };
        }
        
        semanas[claveEsemana].dias.push({
          fecha: fecha,
          horaLlegada: horaLlegada,
          horaSalida: horaSalida,
          monto: monto
        });
        
        totalGeneral += monto;
      }
    }
    
    // Generar texto del reporte
    let reporte = 'Hola, comparto el pago pendiente para programación:\n';
    
    const semanasOrdenadas = Object.keys(semanas).sort((a, b) => {
      return semanas[a].inicio - semanas[b].inicio;
    });
    
    for (const claveSemana of semanasOrdenadas) {
      const semana = semanas[claveSemana];
      const inicioStr = formatearFechaSemana(semana.inicio);
      const finStr = formatearFechaSemana(semana.fin);
      
      reporte += `*Sem. del ${inicioStr} a ${finStr}*\n`;
      
      for (const dia of semana.dias) {
        const diaStr = formatearDiaReporte(dia.fecha);
        const llegadaStr = typeof dia.horaLlegada === 'string' ? dia.horaLlegada : Utilities.formatDate(new Date(dia.horaLlegada), Session.getScriptTimeZone(), 'HH:mm');
        const salidaStr = typeof dia.horaSalida === 'string' ? dia.horaSalida : Utilities.formatDate(new Date(dia.horaSalida), Session.getScriptTimeZone(), 'HH:mm');
        
        reporte += `- ${diaStr}: ${llegadaStr} a ${salidaStr} = _*S/${dia.monto.toFixed(2)}*_\n`;
      }
      reporte += '\n';
    }
    
    reporte += `*Total: S/${totalGeneral.toFixed(2)}*`;
    
    return {
      success: true,
      reporte: reporte
    };
  } catch (error) {
    Logger.log('Error en generarReporte: ' + error.message);
    return {
      success: false,
      mensaje: 'No se pudo generar el reporte. Verifica que las hojas de cálculo estén correctamente configuradas.'
    };
  }
}
// MOD-006: FIN

// MOD-007: GENERAR EXCEL [INICIO]
function generarExcel() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetRegistro = ss.getSheetByName('Registro');
    const sheetTarifa = ss.getSheetByName('Tarifa');
    
    // Crear nuevo spreadsheet temporal
    const nuevoSS = SpreadsheetApp.create('Reporte_Asistencia_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    const sheet = nuevoSS.getActiveSheet();
    sheet.setName('Reporte');
    
    // Obtener tarifas para referencia
    const tarifaHora = sheetTarifa.getRange('A2').getValue();
    const tarifaPasaje = sheetTarifa.getRange('B2').getValue();
    
    // Obtener registros
    const datos = sheetRegistro.getDataRange().getValues();
    
    // Encabezados
    sheet.getRange('A1:F1').setValues([['Fecha', 'Día', 'Llegada', 'Salida', 'Horas', 'Monto']]);
    sheet.getRange('A1:F1').setFontWeight('bold');
    sheet.getRange('A1:F1').setBackground('#4285f4');
    sheet.getRange('A1:F1').setFontColor('#ffffff');
    sheet.getRange('A1:F1').setHorizontalAlignment('center');
    
    let fila = 2;
    let semanaAnterior = null;
    
    // Procesar registros
    const registros = [];
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][1] && datos[i][2]) {
        const fecha = new Date(datos[i][0]);
        const estado = datos[i][7]; // Columna H
        
        // Solo registros NO pagados
        if (estado === 'Pagado') {
          continue;
        }
        
        const horaLlegada = datos[i][1];
        const horaSalida = datos[i][2];
        
        // Obtener semana para agrupar
        const inicioSemana = obtenerInicioSemana(fecha);
        const semanaKey = Utilities.formatDate(inicioSemana, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        
        registros.push({
          fecha: fecha,
          horaLlegada: horaLlegada,
          horaSalida: horaSalida,
          semanaKey: semanaKey
        });
      }
    }
    
    // Ordenar por fecha
    registros.sort((a, b) => a.fecha - b.fecha);
    
    // Escribir datos con fórmulas
    for (const registro of registros) {
      // Detectar cambio de semana para agregar divisor
      if (semanaAnterior !== null && registro.semanaKey !== semanaAnterior) {
        // Agregar borde superior grueso para separar semanas
        sheet.getRange(fila, 1, 1, 6).setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
      semanaAnterior = registro.semanaKey;
      
      // Columna A: Fecha (valor Date, formato DD/MM)
      const cellFecha = sheet.getRange(fila, 1);
      cellFecha.setValue(registro.fecha);
      cellFecha.setNumberFormat('dd/mm');
      cellFecha.setHorizontalAlignment('center');
      
      // Columna B: Día (texto)
      const diaStr = obtenerNombreDia(registro.fecha);
      sheet.getRange(fila, 2).setValue(diaStr);
      
      // Columna C: Llegada (valor hora, formato HH:MM)
      const cellLlegada = sheet.getRange(fila, 3);
      if (typeof registro.horaLlegada === 'string') {
        // Convertir string "HH:mm" a valor de hora para Excel
        const [h, m] = registro.horaLlegada.split(':');
        const valorHora = (parseInt(h) * 60 + parseInt(m)) / (24 * 60);
        cellLlegada.setValue(valorHora);
      } else {
        cellLlegada.setValue(registro.horaLlegada);
      }
      cellLlegada.setNumberFormat('hh:mm');
      cellLlegada.setHorizontalAlignment('center');
      
      // Columna D: Salida (valor hora, formato HH:MM)
      const cellSalida = sheet.getRange(fila, 4);
      if (typeof registro.horaSalida === 'string') {
        const [h, m] = registro.horaSalida.split(':');
        const valorHora = (parseInt(h) * 60 + parseInt(m)) / (24 * 60);
        cellSalida.setValue(valorHora);
      } else {
        cellSalida.setValue(registro.horaSalida);
      }
      cellSalida.setNumberFormat('hh:mm');
      cellSalida.setHorizontalAlignment('center');
      
      // Columna E: Horas (FÓRMULA)
      const cellHoras = sheet.getRange(fila, 5);
      cellHoras.setFormula(`=(D${fila}-C${fila})*24`);
      cellHoras.setNumberFormat('0.00');
      
      // Columna F: Monto (FÓRMULA con REDONDEAR)
      const cellMonto = sheet.getRange(fila, 6);
      cellMonto.setFormula(`=ROUND((E${fila}*${tarifaHora})+${tarifaPasaje},2)`);
      cellMonto.setNumberFormat('0.00');
      
      fila++;
    }
    
    // Fila de Total
    if (fila > 2) {
      fila++;
      sheet.getRange(fila, 5).setValue('TOTAL:');
      sheet.getRange(fila, 5).setHorizontalAlignment('right');
      sheet.getRange(fila, 5).setFontWeight('bold');
      
      const cellTotal = sheet.getRange(fila, 6);
      cellTotal.setFormula(`=SUM(F2:F${fila-2})`);
      cellTotal.setNumberFormat('0.00');
      cellTotal.setFontWeight('bold');
      cellTotal.setBackground('#f4b400');
      sheet.getRange(fila, 5).setBackground('#f4b400');
    }
    
    // Ajustar ancho de columnas
    sheet.setColumnWidth(1, 80);  // Fecha
    sheet.setColumnWidth(2, 100); // Día
    sheet.setColumnWidth(3, 80);  // Llegada
    sheet.setColumnWidth(4, 80);  // Salida
    sheet.setColumnWidth(5, 80);  // Horas
    sheet.setColumnWidth(6, 80);  // Monto
    
    // Obtener URL del archivo
    const url = nuevoSS.getUrl();
    
    return {
      success: true,
      url: url,
      mensaje: 'Excel generado correctamente'
    };
  } catch (error) {
    Logger.log('Error en generarExcel: ' + error.message);
    return {
      success: false,
      mensaje: 'No se pudo generar el archivo Excel. Intenta nuevamente en unos momentos.'
    };
  }
}
// MOD-007: FIN

// MOD-008: OBTENER REGISTROS PENDIENTES [INICIO]
function obtenerRegistrosPendientes() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetRegistro = ss.getSheetByName('Registro');
    const sheetTarifa = ss.getSheetByName('Tarifa');
    
    const datos = sheetRegistro.getDataRange().getValues();
    
    const pendientes = [];
    let totalPendiente = 0;
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][1] && datos[i][2]) {
        const fecha = new Date(datos[i][0]);
        const fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        const estado = datos[i][7]; // Columna H
        
        if (!estado || estado !== 'Pagado') {
          const horaLlegada = datos[i][1];
          const horaSalida = datos[i][2];
          
          // Intentar leer horas y monto de las columnas, si no existen calcular
          let horasTrabajadas, monto;
          
          if (datos[i][5]) {
            horasTrabajadas = parseFloat(datos[i][5]);
          } else {
            horasTrabajadas = calcularHorasTrabajadas(horaLlegada, horaSalida);
          }
          
          if (datos[i][6]) {
            monto = parseFloat(datos[i][6]);
          } else {
            const tarifaHora = sheetTarifa.getRange('A2').getValue();
            const tarifaPasaje = sheetTarifa.getRange('B2').getValue();
            monto = (horasTrabajadas * tarifaHora) + tarifaPasaje;
          }
          
          // Calcular semana
          const inicioSemana = obtenerInicioSemana(fecha);
          const finSemana = obtenerFinSemana(fecha);
          const semanaKey = Utilities.formatDate(inicioSemana, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          const semanaTexto = formatearFechaSemana(inicioSemana) + ' - ' + formatearFechaSemana(finSemana);
          
          pendientes.push({
            fecha: fechaStr,
            fechaCompleta: obtenerFechaCompleta(fecha),
            horasTrabajadas: horasTrabajadas.toFixed(2),
            monto: monto.toFixed(2),
            semanaKey: semanaKey,
            semanaTexto: semanaTexto,
            fechaObj: Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          });
          
          totalPendiente += monto;
        }
      }
    }
    
    // Ordenar por fecha
    pendientes.sort((a, b) => {
      return a.fechaObj.localeCompare(b.fechaObj);
    });
    
    return {
      success: true,
      pendientes: pendientes,
      totalPendiente: totalPendiente.toFixed(2)
    };
  } catch (error) {
    Logger.log('Error en obtenerRegistrosPendientes: ' + error.message);
    return {
      success: false,
      mensaje: 'No se pudo cargar la lista de pendientes. Por favor, recarga la página.'
    };
  }
}
// MOD-008: FIN

// MOD-009: MARCAR COMO PAGADO [INICIO]
function marcarComoPagado(fechas) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetRegistro = ss.getSheetByName('Registro');
    
    // Obtener datos actuales
    const datos = sheetRegistro.getDataRange().getValues();
    
    // Crear mapa de filas por fecha
    const mapaFilas = {};
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0]) {
        const fechaRegistro = Utilities.formatDate(new Date(datos[i][0]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
        mapaFilas[fechaRegistro] = i + 1; // guardar número de fila
      }
    }
    
    // Procesar cada fecha
    let actualizados = 0;
    for (const fechaStr of fechas) {
      if (mapaFilas[fechaStr]) {
        // Marcar como Pagado en columna H
        sheetRegistro.getRange(mapaFilas[fechaStr], 8).setValue('Pagado');
        actualizados++;
      }
    }
    
    return {
      success: true,
      mensaje: actualizados + ' registro(s) marcado(s) como Pagado'
    };
  } catch (error) {
    Logger.log('Error en marcarComoPagado: ' + error.message);
    return {
      success: false,
      mensaje: 'No se pudieron actualizar los estados. Verifica tu conexión e intenta nuevamente.'
    };
  }
}
// MOD-009: FIN

// MOD-010: FUNCIONES AUXILIARES TEXTO [INICIO]
function generarTextoLlegada(fecha) {
  const dias = ['do', 'lu', 'ma', 'mi', 'ju', 'vi', 'sa'];
  const dia = dias[fecha.getDay()];
  const ddmm = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM');
  const hhmm = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'HH:mm');
  
  return `Llegada ${dia} ${ddmm}, ${hhmm}`;
}

function generarTextoSalida(fecha) {
  const dias = ['do', 'lu', 'ma', 'mi', 'ju', 'vi', 'sa'];
  const dia = dias[fecha.getDay()];
  const ddmm = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM');
  const hhmm = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'HH:mm');
  
  return `Salida ${dia} ${ddmm}, ${hhmm}`;
}
// MOD-010: FIN

// MOD-011: FUNCIONES AUXILIARES CÁLCULO [INICIO]
function calcularHorasTrabajadas(horaLlegada, horaSalida) {
  const llegada = convertirAMinutos(horaLlegada);
  const salida = convertirAMinutos(horaSalida);
  
  const minutosTrabajados = salida - llegada;
  return minutosTrabajados / 60;
}

function convertirAMinutos(hora) {
  if (typeof hora === 'string') {
    const [h, m] = hora.split(':').map(Number);
    return h * 60 + m;
  } else {
    const horaStr = Utilities.formatDate(new Date(hora), Session.getScriptTimeZone(), 'HH:mm');
    const [h, m] = horaStr.split(':').map(Number);
    return h * 60 + m;
  }
}
// MOD-011: FIN

// MOD-012: FUNCIONES AUXILIARES SEMANA [INICIO]
function obtenerInicioSemana(fecha) {
  const dia = fecha.getDay();
  const diff = dia === 0 ? -6 : 1 - dia; // Lunes como inicio
  const inicio = new Date(fecha);
  inicio.setDate(fecha.getDate() + diff);
  inicio.setHours(0, 0, 0, 0);
  return inicio;
}

function obtenerFinSemana(fecha) {
  const inicio = obtenerInicioSemana(fecha);
  const fin = new Date(inicio);
  fin.setDate(inicio.getDate() + 6); // Domingo
  return fin;
}
// MOD-012: FIN

// MOD-013: FUNCIONES AUXILIARES FORMATO [INICIO]
function formatearFechaSemana(fecha) {
  const dias = ['do', 'lu', 'ma', 'mi', 'ju', 've', 'sa'];
  const dia = dias[fecha.getDay()];
  const ddmm = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM');
  return `${dia} ${ddmm}`;
}

function formatearDiaReporte(fecha) {
  const dias = ['Dom', 'Lun', 'Mar', 'Mie', 'Jue', 'Vie', 'Sab'];
  const dia = dias[fecha.getDay()];
  const ddmm = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM');
  return `${dia} ${ddmm}`;
}

function obtenerNombreDia(fecha) {
  const dias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
  return dias[fecha.getDay()];
}

function obtenerFechaCompleta(fecha) {
  const dias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
  const dia = dias[fecha.getDay()];
  const ddmm = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM');
  return `${dia}, ${ddmm}`;
}
// MOD-013: FIN

// MOD-014: BATERÍA DE PRUEBAS [INICIO]
function testTotal() {
  Logger.clear();
  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('INICIO DE PRUEBAS - SISTEMA DE ASISTENCIA');
  Logger.log('Fecha: 19/12/2025 - 00:17');
  Logger.log('═══════════════════════════════════════════════════════\n');
  
  let totalPruebas = 0;
  let pruebasExitosas = 0;
  let pruebasFallidas = 0;
  
  // ============ PRUEBA 1: Conexión a Google Sheet ============
  Logger.log('📝 PRUEBA 1: Verificación de conexión a Google Sheet');
  totalPruebas++;
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const nombre = ss.getName();
    Logger.log('✅ Conexión exitosa');
    Logger.log('   Nombre del Sheet: ' + nombre);
    pruebasExitosas++;
  } catch (error) {
    Logger.log('❌ ERROR: No se pudo conectar al Sheet');
    Logger.log('   Detalle: ' + error.message);
    pruebasFallidas++;
  }
  Logger.log('');
  
  // ============ PRUEBA 2: Verificación de hojas ============
  Logger.log('📋 PRUEBA 2: Verificación de hojas requeridas');
  totalPruebas++;
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const hojas = ['Registro', 'Tarifa'];
    let hojasEncontradas = 0;
    
    for (const nombreHoja of hojas) {
      const hoja = ss.getSheetByName(nombreHoja);
      if (hoja) {
        Logger.log(`   ✓ Hoja "${nombreHoja}" encontrada`);
        hojasEncontradas++;
      } else {
        Logger.log(`   ✗ Hoja "${nombreHoja}" NO encontrada`);
      }
    }
    
    if (hojasEncontradas === hojas.length) {
      Logger.log('✅ Todas las hojas están presentes');
      pruebasExitosas++;
    } else {
      Logger.log('❌ Faltan ' + (hojas.length - hojasEncontradas) + ' hoja(s)');
      pruebasFallidas++;
    }
  } catch (error) {
    Logger.log('❌ ERROR en verificación de hojas');
    Logger.log('   Detalle: ' + error.message);
    pruebasFallidas++;
  }
  Logger.log('');
  
  // ============ PRUEBA 3: Verificación de tarifas ============
  Logger.log('💰 PRUEBA 3: Verificación de configuración de tarifas');
  totalPruebas++;
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetTarifa = ss.getSheetByName('Tarifa');
    const tarifaHora = sheetTarifa.getRange('A2').getValue();
    const tarifaPasaje = sheetTarifa.getRange('B2').getValue();
    
    if (tarifaHora && tarifaPasaje) {
      Logger.log('✅ Tarifas configuradas correctamente');
      Logger.log('   Tarifa por hora: S/' + tarifaHora);
      Logger.log('   Tarifa pasaje: S/' + tarifaPasaje);
      pruebasExitosas++;
    } else {
      Logger.log('❌ Las tarifas no están configuradas');
      Logger.log('   Tarifa hora: ' + (tarifaHora || 'VACÍO'));
      Logger.log('   Tarifa pasaje: ' + (tarifaPasaje || 'VACÍO'));
      pruebasFallidas++;
    }
  } catch (error) {
    Logger.log('❌ ERROR al verificar tarifas');
    Logger.log('   Detalle: ' + error.message);
    pruebasFallidas++;
  }
  Logger.log('');
  
  // ============ PRUEBA 4: Funciones auxiliares ============
  Logger.log('🔧 PRUEBA 4: Prueba de funciones auxiliares');
  totalPruebas++;
  try {
    const fechaPrueba = new Date();
    
    // Probar generación de textos
    const textoLlegada = generarTextoLlegada(fechaPrueba);
    const textoSalida = generarTextoSalida(fechaPrueba);
    
    Logger.log('   Texto Llegada: ' + textoLlegada);
    Logger.log('   Texto Salida: ' + textoSalida);
    
    // Probar cálculo de horas
    const horas = calcularHorasTrabajadas('09:00', '18:00');
    Logger.log('   Cálculo horas (09:00 a 18:00): ' + horas + ' horas');
    
    // Probar fechas de semana
    const inicioSemana = obtenerInicioSemana(fechaPrueba);
    const finSemana = obtenerFinSemana(fechaPrueba);
    Logger.log('   Inicio de semana: ' + Utilities.formatDate(inicioSemana, Session.getScriptTimeZone(), 'dd/MM/yyyy'));
    Logger.log('   Fin de semana: ' + Utilities.formatDate(finSemana, Session.getScriptTimeZone(), 'dd/MM/yyyy'));
    
    Logger.log('✅ Funciones auxiliares operando correctamente');
    pruebasExitosas++;
  } catch (error) {
    Logger.log('❌ ERROR en funciones auxiliares');
    Logger.log('   Detalle: ' + error.message);
    pruebasFallidas++;
  }
  Logger.log('');
  
  // ============ PRUEBA 5: Registro de llegada (simulación) ============
  Logger.log('🚪 PRUEBA 5: Simulación de registro de llegada');
  totalPruebas++;
  try {
    const resultado = registrarLlegada();
    
    if (resultado.success) {
      Logger.log('✅ Registro de llegada exitoso');
      Logger.log('   Mensaje: ' + resultado.mensaje);
      Logger.log('   Texto generado: ' + resultado.texto);
      pruebasExitosas++;
    } else {
      Logger.log('❌ Fallo en registro de llegada');
      Logger.log('   Mensaje: ' + resultado.mensaje);
      pruebasFallidas++;
    }
  } catch (error) {
    Logger.log('❌ ERROR al probar registro de llegada');
    Logger.log('   Detalle: ' + error.message);
    pruebasFallidas++;
  }
  Logger.log('');
  
  // ============ PRUEBA 6: Obtener registro de hoy ============
  Logger.log('📖 PRUEBA 6: Lectura de registro actual');
  totalPruebas++;
  try {
    const resultado = obtenerRegistroHoy();
    
    if (resultado.success) {
      Logger.log('✅ Lectura de registro exitosa');
      Logger.log('   Texto Llegada: ' + (resultado.textoLlegada || 'Sin registro'));
      Logger.log('   Texto Salida: ' + (resultado.textoSalida || 'Sin registro'));
      pruebasExitosas++;
    } else {
      Logger.log('❌ Fallo en lectura de registro');
      Logger.log('   Mensaje: ' + resultado.mensaje);
      pruebasFallidas++;
    }
  } catch (error) {
    Logger.log('❌ ERROR al leer registro');
    Logger.log('   Detalle: ' + error.message);
    pruebasFallidas++;
  }
  Logger.log('');
  
  // ============ PRUEBA 7: Obtener registros pendientes ============
  Logger.log('💸 PRUEBA 7: Obtención de registros pendientes');
  totalPruebas++;
  try {
    const resultado = obtenerRegistrosPendientes();
    
    if (resultado.success) {
      Logger.log('✅ Consulta de pendientes exitosa');
      Logger.log('   Cantidad de registros pendientes: ' + resultado.pendientes.length);
      Logger.log('   Total pendiente: S/' + resultado.totalPendiente);
      
      if (resultado.pendientes.length > 0) {
        Logger.log('   Primeros 3 registros:');
        for (let i = 0; i < Math.min(3, resultado.pendientes.length); i++) {
          const reg = resultado.pendientes[i];
          Logger.log('   - ' + reg.fechaCompleta + ': S/' + reg.monto);
        }
      }
      pruebasExitosas++;
    } else {
      Logger.log('❌ Fallo en consulta de pendientes');
      Logger.log('   Mensaje: ' + resultado.mensaje);
      pruebasFallidas++;
    }
  } catch (error) {
    Logger.log('❌ ERROR al consultar pendientes');
    Logger.log('   Detalle: ' + error.message);
    pruebasFallidas++;
  }
  Logger.log('');
  
  // ============ PRUEBA 8: Generar reporte ============
  Logger.log('📊 PRUEBA 8: Generación de reporte');
  totalPruebas++;
  try {
    const resultado = generarReporte();
    
    if (resultado.success) {
      Logger.log('✅ Reporte generado exitosamente');
      const lineas = resultado.reporte.split('\n').length;
      Logger.log('   Líneas del reporte: ' + lineas);
      Logger.log('   Primeras 3 líneas:');
      const primerasLineas = resultado.reporte.split('\n').slice(0, 3);
      primerasLineas.forEach(linea => {
        if (linea.trim()) Logger.log('   ' + linea);
      });
      pruebasExitosas++;
    } else {
      Logger.log('❌ Fallo en generación de reporte');
      Logger.log('   Mensaje: ' + resultado.mensaje);
      pruebasFallidas++;
    }
  } catch (error) {
    Logger.log('❌ ERROR al generar reporte');
    Logger.log('   Detalle: ' + error.message);
    pruebasFallidas++;
  }
  Logger.log('');
  
  // ============ RESUMEN FINAL ============
  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('RESUMEN DE PRUEBAS');
  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('Total de pruebas ejecutadas: ' + totalPruebas);
  Logger.log('✅ Pruebas exitosas: ' + pruebasExitosas);
  Logger.log('❌ Pruebas fallidas: ' + pruebasFallidas);
  Logger.log('Porcentaje de éxito: ' + ((pruebasExitosas / totalPruebas) * 100).toFixed(1) + '%');
  Logger.log('═══════════════════════════════════════════════════════');
  
  if (pruebasFallidas === 0) {
    Logger.log('\n🎉 ¡TODAS LAS PRUEBAS PASARON EXITOSAMENTE!');
    Logger.log('El sistema está listo para usarse.');
  } else {
    Logger.log('\n⚠️ ALGUNAS PRUEBAS FALLARON');
    Logger.log('Revisa los errores anteriores y corrige la configuración.');
  }
  
  Logger.log('\n📌 Revisa los logs completos en: Ver > Registros (Ctrl/Cmd + Enter)');
}
// MOD-014: FIN

// MOD-015: MENÚ PERSONALIZADO [INICIO]
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Actualizar')
    .addItem('🔄 Actualizar campos faltantes', 'refreshTextos')
    .addSeparator()
    .addItem('📊 Ejecutar todas las pruebas', 'testTotal')
    .addToUi();
}
// MOD-015: FIN

// MOD-016: REFRESH TEXTOS [INICIO]
function refreshTextos() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Registro');
    const sheetTarifa = ss.getSheetByName('Tarifa');
    const data = sheet.getDataRange().getValues();

    let cambios = 0;
    let errores = 0;

    for (let i = 1; i < data.length; i++) {
      const fechaOriginal = data[i][0];  // A
      const llegada = data[i][1];        // B
      const salida = data[i][2];         // C
      let textoLlegada = data[i][3];     // D
      let textoSalida = data[i][4];      // E
      let horas = data[i][5];            // F
      let monto = data[i][6];            // G

      // Saltar si no hay fecha
      if (!fechaOriginal) continue;

      // Convertir fecha a objeto Date válido
      let fechaObj;
      if (fechaOriginal instanceof Date) {
        fechaObj = new Date(fechaOriginal);
      } else {
        // Si es string, intentar parsear
        fechaObj = new Date(fechaOriginal);
      }
      
      // Verificar que la fecha sea válida
      if (isNaN(fechaObj.getTime())) {
        Logger.log('Fila ' + (i + 1) + ': Fecha inválida - ' + fechaOriginal);
        errores++;
        continue;
      }

      // Texto llegada
      if (!textoLlegada && llegada) {
        try {
          // Crear fecha completa para llegada
          const fechaLlegada = new Date(fechaObj);
          
          // Extraer hora y minutos de llegada
          let horaLlegada, minutoLlegada;
          if (typeof llegada === 'string') {
            const partes = llegada.split(':');
            if (partes.length >= 2) {
              horaLlegada = parseInt(partes[0]);
              minutoLlegada = parseInt(partes[1]);
            } else {
              throw new Error('Formato de hora inválido');
            }
          } else if (llegada instanceof Date) {
            horaLlegada = llegada.getHours();
            minutoLlegada = llegada.getMinutes();
          } else {
            throw new Error('Tipo de dato no válido');
          }
          
          // Validar hora y minutos
          if (horaLlegada >= 0 && horaLlegada <= 23 && minutoLlegada >= 0 && minutoLlegada <= 59) {
            fechaLlegada.setHours(horaLlegada, minutoLlegada, 0, 0);
            
            const textoGenerado = generarTextoLlegada(fechaLlegada);
            sheet.getRange(i + 1, 4).setValue(textoGenerado);
            cambios++;
          } else {
            throw new Error('Hora fuera de rango');
          }
        } catch (error) {
          Logger.log('Fila ' + (i + 1) + ': Error en llegada - ' + error.message);
          errores++;
        }
      }

      // Texto salida
      if (!textoSalida && salida) {
        try {
          // Crear fecha completa para salida
          const fechaSalida = new Date(fechaObj);
          
          // Extraer hora y minutos de salida
          let horaSalida, minutoSalida;
          if (typeof salida === 'string') {
            const partes = salida.split(':');
            if (partes.length >= 2) {
              horaSalida = parseInt(partes[0]);
              minutoSalida = parseInt(partes[1]);
            } else {
              throw new Error('Formato de hora inválido');
            }
          } else if (salida instanceof Date) {
            horaSalida = salida.getHours();
            minutoSalida = salida.getMinutes();
          } else {
            throw new Error('Tipo de dato no válido');
          }
          
          // Validar hora y minutos
          if (horaSalida >= 0 && horaSalida <= 23 && minutoSalida >= 0 && minutoSalida <= 59) {
            fechaSalida.setHours(horaSalida, minutoSalida, 0, 0);
            
            const textoGenerado = generarTextoSalida(fechaSalida);
            sheet.getRange(i + 1, 5).setValue(textoGenerado);
            cambios++;
          } else {
            throw new Error('Hora fuera de rango');
          }
        } catch (error) {
          Logger.log('Fila ' + (i + 1) + ': Error en salida - ' + error.message);
          errores++;
        }
      }
      
      // Calcular Horas y Monto si están vacíos y hay llegada y salida
      if (llegada && salida) {
        try {
          // Horas trabajadas
          if (!horas) {
            const horasTrabajadas = calcularHorasTrabajadas(llegada, salida);
            const cellHoras = sheet.getRange(i + 1, 6);
            cellHoras.setValue(parseFloat(horasTrabajadas.toFixed(2)));
            cellHoras.setNumberFormat('0.00'); // Formato numérico con 2 decimales
            cambios++;
          }
          
          // Monto - calcular con horas completas (sin redondear primero)
          if (!monto) {
            const tarifaHora = sheetTarifa.getRange('A2').getValue();
            const tarifaPasaje = sheetTarifa.getRange('B2').getValue();
            
            // Usar horas SIN redondear para el cálculo
            const horasParaCalculo = horas || calcularHorasTrabajadas(llegada, salida);
            const montoCalculado = (horasParaCalculo * tarifaHora) + tarifaPasaje;
            
            const cellMonto = sheet.getRange(i + 1, 7);
            cellMonto.setValue(parseFloat(montoCalculado.toFixed(2))); // Redondear solo el monto final
            cellMonto.setNumberFormat('0.00'); // Formato numérico con 2 decimales
            cambios++;
          }
        } catch (error) {
          Logger.log('Fila ' + (i + 1) + ': Error en cálculos - ' + error.message);
          errores++;
        }
      }
    }

    // Mensaje final
    let mensaje = '✅ Actualización completada\n\n';
    mensaje += '📝 Campos actualizados: ' + cambios;
    
    if (errores > 0) {
      mensaje += '\n⚠️ Errores encontrados: ' + errores;
      mensaje += '\n\nRevisa los registros (Ver > Registros) para más detalles.';
    }
    
    SpreadsheetApp.getUi().alert('Actualizar Datos', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
    
    Logger.log('refreshTextos completado: ' + cambios + ' campos actualizados, ' + errores + ' errores');
    
  } catch (error) {
    Logger.log('Error en refreshTextos: ' + error.message);
    SpreadsheetApp.getUi().alert(
      'Error al actualizar',
      '❌ Error al actualizar datos.\n\n' + 
      'Verifica que las columnas tengan el formato correcto:\n' +
      '• Columna A: Fecha (formato fecha)\n' +
      '• Columnas B y C: Hora en formato texto "HH:mm" (ejemplo: "15:30")\n\n' +
      'Revisa los registros (Ver > Registros) para más detalles.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
// MOD-016: FIN

// MOD-099: NOTAS [INICIO]
/*
DESCRIPCIÓN:
Sistema de asistencia para registro de entrada/salida con generación de reportes
y seguimiento de pagos. Permite registrar llegada/salida, generar reportes para
WhatsApp, exportar a Excel y gestionar estados de pago.

DEPENDENCIAS:
- MOD-003 y MOD-004: Usan funciones auxiliares de MOD-010, MOD-011
- MOD-006 y MOD-007: Usan funciones de MOD-011, MOD-012, MOD-013
- MOD-008: Usa funciones de MOD-011, MOD-012, MOD-013
- MOD-014: Usa todas las funciones principales para pruebas
- MOD-016: Usa funciones de MOD-010 y MOD-011
*/
// MOD-099: FIN
