/**
 * ============================================
 * SISTEMA DE ASISTENCIA - CODE.GS
 * Versión: V1.01
 * Descripción: Backend del sistema de registro de asistencia
 * Autor: Jorge
 * Fecha: Diciembre 2025
 * ============================================
 */

// ========================================
// CONFIGURACIÓN
// ========================================
const SHEET_ID = '1GI5C5djzMEFCcQEewi-MotEBggQa5VMh0ylchHjV5kM';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('home')
    .setTitle('Sistema de Asistencia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ========================================
// FUNCIONES DE REGISTRO
// ========================================

function registrarLlegada() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetRegistro = ss.getSheetByName('Registro');
    const sheetTarifa = ss.getSheetByName('Tarifa');
    
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
      // Crear nuevo registro
      sheetRegistro.appendRow([fecha, hora, '', textoLlegada, '']);
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

function registrarSalida() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetRegistro = ss.getSheetByName('Registro');
    
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
    
    // Generar texto de salida
    const textoSalida = generarTextoSalida(ahora);
    
    if (filaExistente > 0) {
      sheetRegistro.getRange(filaExistente, 3).setValue(hora); // Columna C
      sheetRegistro.getRange(filaExistente, 5).setValue(textoSalida); // Columna E
    } else {
      return {
        success: false,
        mensaje: 'No hay registro de llegada para hoy'
      };
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

// ========================================
// FUNCIONES DE REPORTE
// ========================================

function generarReporte() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetRegistro = ss.getSheetByName('Registro');
    const sheetTarifa = ss.getSheetByName('Tarifa');
    const sheetEstado = ss.getSheetByName('Estado');
    
    // Obtener tarifas
    const tarifaHora = sheetTarifa.getRange('A2').getValue();
    const tarifaPasaje = sheetTarifa.getRange('B2').getValue();
    
    // Obtener todos los registros
    const datos = sheetRegistro.getDataRange().getValues();
    const datosEstado = sheetEstado.getDataRange().getValues();
    
    // Crear mapa de estados
    const mapaEstados = {};
    for (let i = 1; i < datosEstado.length; i++) {
      if (datosEstado[i][0]) {
        const fechaEstado = Utilities.formatDate(new Date(datosEstado[i][0]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
        mapaEstados[fechaEstado] = datosEstado[i][1];
      }
    }
    
    // Procesar registros por semana
    const semanas = {};
    let totalGeneral = 0;
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][1] && datos[i][2]) {
        const fecha = new Date(datos[i][0]);
        const fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        
        // Verificar si ya está pagado
        if (mapaEstados[fechaStr] === 'Pagado') {
          continue;
        }
        
        const horaLlegada = datos[i][1];
        const horaSalida = datos[i][2];
        
        // Calcular horas trabajadas
        const horasTrabajadas = calcularHorasTrabajadas(horaLlegada, horaSalida);
        const monto = (horasTrabajadas * tarifaHora) + tarifaPasaje;
        
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

function generarExcel() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetRegistro = ss.getSheetByName('Registro');
    const sheetTarifa = ss.getSheetByName('Tarifa');
    const sheetEstado = ss.getSheetByName('Estado');
    
    // Crear nuevo spreadsheet temporal
    const nuevoSS = SpreadsheetApp.create('Reporte_Asistencia_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    const sheet = nuevoSS.getActiveSheet();
    sheet.setName('Reporte');
    
    // Obtener tarifas
    const tarifaHora = sheetTarifa.getRange('A2').getValue();
    const tarifaPasaje = sheetTarifa.getRange('B2').getValue();
    
    // Obtener registros
    const datos = sheetRegistro.getDataRange().getValues();
    const datosEstado = sheetEstado.getDataRange().getValues();
    
    // Crear mapa de estados
    const mapaEstados = {};
    for (let i = 1; i < datosEstado.length; i++) {
      if (datosEstado[i][0]) {
        const fechaEstado = Utilities.formatDate(new Date(datosEstado[i][0]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
        mapaEstados[fechaEstado] = datosEstado[i][1];
      }
    }
    
    // Encabezados
    sheet.getRange('A1:G1').setValues([['Semana', 'Fecha', 'Día', 'Llegada', 'Salida', 'Horas', 'Monto']]);
    sheet.getRange('A1:G1').setFontWeight('bold');
    sheet.getRange('A1:G1').setBackground('#4285f4');
    sheet.getRange('A1:G1').setFontColor('#ffffff');
    
    let fila = 2;
    let totalGeneral = 0;
    
    // Procesar por semanas
    const semanas = {};
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][1] && datos[i][2]) {
        const fecha = new Date(datos[i][0]);
        const fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        
        // Verificar si ya está pagado
        if (mapaEstados[fechaStr] === 'Pagado') {
          continue;
        }
        
        const horaLlegada = datos[i][1];
        const horaSalida = datos[i][2];
        
        // Calcular horas trabajadas
        const horasTrabajadas = calcularHorasTrabajadas(horaLlegada, horaSalida);
        const monto = (horasTrabajadas * tarifaHora) + tarifaPasaje;
        
        // Obtener semana
        const inicioSemana = obtenerInicioSemana(fecha);
        const finSemana = obtenerFinSemana(fecha);
        const claveEsemana = Utilities.formatDate(inicioSemana, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        
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
          horasTrabajadas: horasTrabajadas,
          monto: monto
        });
      }
    }
    
    // Escribir datos
    const semanasOrdenadas = Object.keys(semanas).sort();
    
    for (const claveSemana of semanasOrdenadas) {
      const semana = semanas[claveSemana];
      const semanaStr = `${formatearFechaSemana(semana.inicio)} a ${formatearFechaSemana(semana.fin)}`;
      
      for (const dia of semana.dias) {
        const fechaStr = Utilities.formatDate(dia.fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        const diaStr = obtenerNombreDia(dia.fecha);
        const llegadaStr = typeof dia.horaLlegada === 'string' ? dia.horaLlegada : Utilities.formatDate(new Date(dia.horaLlegada), Session.getScriptTimeZone(), 'HH:mm');
        const salidaStr = typeof dia.horaSalida === 'string' ? dia.horaSalida : Utilities.formatDate(new Date(dia.horaSalida), Session.getScriptTimeZone(), 'HH:mm');
        
        sheet.getRange(fila, 1).setValue(semanaStr);
        sheet.getRange(fila, 2).setValue(fechaStr);
        sheet.getRange(fila, 3).setValue(diaStr);
        sheet.getRange(fila, 4).setValue(llegadaStr);
        sheet.getRange(fila, 5).setValue(salidaStr);
        sheet.getRange(fila, 6).setValue(dia.horasTrabajadas.toFixed(2));
        sheet.getRange(fila, 7).setValue(dia.monto.toFixed(2));
        
        totalGeneral += dia.monto;
        fila++;
      }
    }
    
    // Total
    fila++;
    sheet.getRange(fila, 6).setValue('TOTAL:');
    sheet.getRange(fila, 7).setValue(totalGeneral.toFixed(2));
    sheet.getRange(fila, 6, 1, 2).setFontWeight('bold');
    sheet.getRange(fila, 6, 1, 2).setBackground('#f4b400');
    
    // Ajustar columnas
    sheet.autoResizeColumns(1, 7);
    
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

// ========================================
// FUNCIONES DE ESTADO
// ========================================

function obtenerRegistrosPendientes() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetRegistro = ss.getSheetByName('Registro');
    const sheetTarifa = ss.getSheetByName('Tarifa');
    const sheetEstado = ss.getSheetByName('Estado');
    
    const tarifaHora = sheetTarifa.getRange('A2').getValue();
    const tarifaPasaje = sheetTarifa.getRange('B2').getValue();
    
    const datos = sheetRegistro.getDataRange().getValues();
    const datosEstado = sheetEstado.getDataRange().getValues();
    
    // Crear mapa de estados
    const mapaEstados = {};
    for (let i = 1; i < datosEstado.length; i++) {
      if (datosEstado[i][0]) {
        const fechaEstado = Utilities.formatDate(new Date(datosEstado[i][0]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
        mapaEstados[fechaEstado] = datosEstado[i][1];
      }
    }
    
    const pendientes = [];
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][1] && datos[i][2]) {
        const fecha = new Date(datos[i][0]);
        const fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        
        if (!mapaEstados[fechaStr] || mapaEstados[fechaStr] !== 'Pagado') {
          const horaLlegada = datos[i][1];
          const horaSalida = datos[i][2];
          const horasTrabajadas = calcularHorasTrabajadas(horaLlegada, horaSalida);
          const monto = (horasTrabajadas * tarifaHora) + tarifaPasaje;
          
          pendientes.push({
            fecha: fechaStr,
            fechaCompleta: obtenerFechaCompleta(fecha),
            horasTrabajadas: horasTrabajadas.toFixed(2),
            monto: monto.toFixed(2)
          });
        }
      }
    }
    
    return {
      success: true,
      pendientes: pendientes
    };
  } catch (error) {
    Logger.log('Error en obtenerRegistrosPendientes: ' + error.message);
    return {
      success: false,
      mensaje: 'No se pudo cargar la lista de pendientes. Por favor, recarga la página.'
    };
  }
}

function marcarComoPagado(fechas) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetEstado = ss.getSheetByName('Estado');
    
    // Obtener datos actuales
    const datos = sheetEstado.getDataRange().getValues();
    const mapaEstados = {};
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0]) {
        const fechaEstado = Utilities.formatDate(new Date(datos[i][0]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
        mapaEstados[fechaEstado] = i + 1; // guardar fila
      }
    }
    
    // Procesar cada fecha
    for (const fechaStr of fechas) {
      const [dia, mes, anio] = fechaStr.split('/');
      const fecha = new Date(anio, mes - 1, dia);
      
      if (mapaEstados[fechaStr]) {
        // Actualizar existente
        sheetEstado.getRange(mapaEstados[fechaStr], 2).setValue('Pagado');
      } else {
        // Crear nuevo registro
        sheetEstado.appendRow([fecha, 'Pagado']);
      }
    }
    
    return {
      success: true,
      mensaje: 'Estados actualizados correctamente'
    };
  } catch (error) {
    Logger.log('Error en marcarComoPagado: ' + error.message);
    return {
      success: false,
      mensaje: 'No se pudieron actualizar los estados. Verifica tu conexión e intenta nuevamente.'
    };
  }
}

// ========================================
// FUNCIONES AUXILIARES
// ========================================

function generarTextoLlegada(fecha) {
  const dias = ['do', 'lu', 'ma', 'mi', 'ju', 've', 'sa'];
  const dia = dias[fecha.getDay()];
  const ddmm = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM');
  const hhmm = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'HH:mm');
  
  return `Llegada ${dia} ${ddmm}, ${hhmm}`;
}

function generarTextoSalida(fecha) {
  const dias = ['do', 'lu', 'ma', 'mi', 'ju', 've', 'sa'];
  const dia = dias[fecha.getDay()];
  const ddmm = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM');
  const hhmm = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'HH:mm');
  
  return `Salida ${dia} ${ddmm}, ${hhmm}`;
}

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

// ========================================
// BATERÍA DE PRUEBAS UNIFICADA
// ========================================

/**
 * Función principal de pruebas - Ejecuta todas las pruebas del sistema
 * Para ejecutar: Selecciona "testTotal" en el menú de funciones y haz clic en ▶️
 */
function testTotal() {
  Logger.clear();
  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('INICIO DE PRUEBAS - SISTEMA DE ASISTENCIA V1.01');
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
    const hojas = ['Registro', 'Tarifa', 'Estado'];
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
