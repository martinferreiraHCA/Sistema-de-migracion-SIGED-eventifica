/**
 * Aplicación Web para Conversión de Fichas de Inscripción
 * Colegio Hans Christian Andersen
 *
 * Convierte datos de "Ficha de Inscripción" a:
 * 1. Formato EVENTIFICA (template_estudiantes_padres) - con usuarios CI+hca
 * 2. Formato SIGED (Plantilla_Importar_AlumnosYFamilias)
 *
 * @version 2.1 - Usuarios automáticos y corrección SIGED
 */

// ============================================
// CONFIGURACIÓN Y UTILIDADES
// ============================================

/**
 * Configuración global del sistema
 */
const CONFIG = {
  MAX_RETRIES: 3,
  RETRY_DELAY: 1000, // milliseconds
  TIMEOUT_LIMIT: 300000, // 5 minutos
  LOG_ENABLED: true
};

/**
 * Logger mejorado para debugging
 */
function logInfo(message, data) {
  if (CONFIG.LOG_ENABLED) {
    Logger.log('[INFO] ' + message);
    if (data) Logger.log(JSON.stringify(data));
  }
}

function logError(message, error) {
  Logger.log('[ERROR] ' + message);
  if (error) {
    Logger.log('Error details: ' + error.toString());
    if (error.stack) Logger.log('Stack: ' + error.stack);
  }
}

/**
 * Ejecuta una función con reintentos exponenciales
 */
function retryOperation(operation, operationName, maxRetries) {
  maxRetries = maxRetries || CONFIG.MAX_RETRIES;
  var lastError;

  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      logInfo('Intentando operación: ' + operationName + ' (intento ' + attempt + '/' + maxRetries + ')');
      var result = operation();
      logInfo('Operación exitosa: ' + operationName);
      return result;
    } catch (error) {
      lastError = error;
      logError('Error en operación ' + operationName + ' (intento ' + attempt + ')', error);

      if (attempt < maxRetries) {
        var delay = CONFIG.RETRY_DELAY * Math.pow(2, attempt - 1);
        logInfo('Esperando ' + delay + 'ms antes del siguiente intento...');
        Utilities.sleep(delay);
      }
    }
  }

  throw new Error('Operación "' + operationName + '" falló después de ' + maxRetries + ' intentos. Último error: ' + lastError.toString());
}

/**
 * Verifica los permisos de acceso a Drive y Sheets
 */
function checkPermissions() {
  try {
    logInfo('Verificando permisos de acceso...');

    // Verificar acceso a Drive
    var testFolder = DriveApp.getRootFolder();
    logInfo('Acceso a Drive: OK');

    // Verificar acceso a Sheets
    var testSheet = SpreadsheetApp.create('__PERMISSION_TEST__');
    var testId = testSheet.getId();
    DriveApp.getFileById(testId).setTrashed(true);
    logInfo('Acceso a Sheets: OK');

    return { success: true };
  } catch (error) {
    logError('Error de permisos', error);
    return {
      success: false,
      error: 'Permisos insuficientes. Por favor, autoriza la aplicación para acceder a Google Drive y Sheets.'
    };
  }
}

/**
 * Limpia archivos temporales antiguos (más de 1 hora)
 */
function cleanupOldTempFiles() {
  try {
    logInfo('Limpiando archivos temporales antiguos...');
    var files = DriveApp.searchFiles('title contains "TEMP_" and trashed = false');
    var oneHourAgo = new Date(new Date().getTime() - 3600000);
    var count = 0;

    while (files.hasNext()) {
      var file = files.next();
      if (file.getDateCreated() < oneHourAgo) {
        file.setTrashed(true);
        count++;
      }
    }

    logInfo('Archivos temporales eliminados: ' + count);
  } catch (error) {
    logError('Error limpiando archivos temporales', error);
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Conversor de Fichas de Inscripción - HCA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================
// PROCESAMIENTO DE ARCHIVOS
// ============================================

/**
 * Procesa el archivo subido y extrae los datos (VERSIÓN ROBUSTA)
 */
function processUploadedFile(base64Data, fileName) {
  var tempFile = null;
  var spreadsheet = null;
  var sheetId = null; // ID del Google Sheets convertido

  try {
    logInfo('Iniciando procesamiento de archivo: ' + fileName);

    // Validar entrada
    if (!base64Data || base64Data.length === 0) {
      throw new Error('Datos del archivo vacíos o inválidos');
    }

    // Limpiar archivos temporales viejos antes de empezar
    cleanupOldTempFiles();

    // Decodificar el archivo con reintentos
    var blob = retryOperation(function() {
      var decoded = Utilities.base64Decode(base64Data);
      return Utilities.newBlob(decoded, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'TEMP_' + fileName);
    }, 'Decodificar archivo');

    logInfo('Archivo decodificado exitosamente');

    // Crear archivo temporal en Drive con reintentos
    tempFile = retryOperation(function() {
      return DriveApp.createFile(blob);
    }, 'Crear archivo temporal en Drive');

    var tempFileId = tempFile.getId();
    logInfo('Archivo temporal creado con ID: ' + tempFileId);

    // SOLUCIÓN MEJORADA: Convertir Excel a Google Sheets nativo
    // En lugar de intentar abrir directamente, convertimos explícitamente
    spreadsheet = retryOperation(function() {
      logInfo('Intentando conversión de Excel a Google Sheets...');

      try {
        // MÉTODO 1: Usar Advanced Drive API (más confiable)
        var file = DriveApp.getFileById(tempFileId);
        var blob = file.getBlob();

        var resource = {
          title: 'TEMP_Sheet_' + new Date().getTime(),
          mimeType: 'application/vnd.google-apps.spreadsheet'
        };

        var convertedFile = Drive.Files.insert(resource, blob, {
          convert: true,
          supportsAllDrives: true
        });

        sheetId = convertedFile.id;
        logInfo('Archivo convertido a Google Sheets con ID (Drive API): ' + sheetId);

        // Esperar un momento para que Google procese la conversión
        Utilities.sleep(2000);

        return SpreadsheetApp.openById(sheetId);

      } catch (driveApiError) {
        // MÉTODO 2: Fallback - Intentar método alternativo sin Drive API avanzada
        logError('Error con Drive API, intentando método alternativo', driveApiError);

        // Crear un nuevo spreadsheet vacío
        var newSheet = SpreadsheetApp.create('TEMP_Import_' + new Date().getTime());
        sheetId = newSheet.getId();

        // Obtener el blob del archivo Excel
        var file = DriveApp.getFileById(tempFileId);
        var blob = file.getBlob();

        // Importar datos del Excel al nuevo spreadsheet
        // Convertir el blob a base64
        var bytes = blob.getBytes();
        var base64 = Utilities.base64Encode(bytes);

        // Descargar y parsear Excel usando un método alternativo
        // Crear un spreadsheet temporal importando el archivo
        var importedFile = Drive.Files.insert({
          title: 'TEMP_Import_Alt_' + new Date().getTime(),
          mimeType: MimeType.GOOGLE_SHEETS
        }, blob, { convert: true });

        sheetId = importedFile.id;
        logInfo('Archivo convertido con método alternativo, ID: ' + sheetId);

        Utilities.sleep(2000);

        return SpreadsheetApp.openById(sheetId);
      }
    }, 'Convertir y abrir spreadsheet', 5); // Más intentos para esta operación crítica

    logInfo('Spreadsheet abierto exitosamente');

    // Obtener datos del spreadsheet
    var sheets = spreadsheet.getSheets();
    if (!sheets || sheets.length === 0) {
      throw new Error('El archivo no contiene hojas de cálculo');
    }

    var sheet = sheets[0];
    logInfo('Leyendo datos de la hoja: ' + sheet.getName());

    var data = retryOperation(function() {
      var range = sheet.getDataRange();
      if (!range) throw new Error('No se pudo obtener el rango de datos');
      return range.getValues();
    }, 'Obtener datos del spreadsheet');

    logInfo('Datos leídos: ' + data.length + ' filas');

    // Validar que hay datos
    if (!data || data.length < 2) {
      throw new Error('El archivo no contiene suficientes datos (se necesitan al menos 2 filas: encabezados + 1 registro)');
    }

    // Procesar datos
    var headers = data[0];
    var records = [];
    var errors = [];

    logInfo('Procesando ' + (data.length - 1) + ' registros...');

    for (var i = 1; i < data.length; i++) {
      try {
        var row = data[i];

        // Validar que la fila no esté completamente vacía
        var isEmptyRow = row.every(function(cell) {
          return cell === '' || cell === null || cell === undefined;
        });

        if (!isEmptyRow) {
          var record = extractRecord(headers, row, i);
          if (record && record.estudiante && record.estudiante.nombre) {
            records.push(record);
          } else {
            logInfo('Registro en fila ' + (i + 1) + ' omitido (sin nombre)');
          }
        }
      } catch (recordError) {
        logError('Error procesando fila ' + (i + 1), recordError);
        errors.push({
          row: i + 1,
          error: recordError.toString()
        });
      }
    }

    logInfo('Procesamiento completado. Registros válidos: ' + records.length);

    return {
      success: true,
      records: records,
      totalRows: data.length - 1,
      validRecords: records.length,
      errors: errors.length > 0 ? errors : undefined
    };

  } catch (error) {
    logError('Error en processUploadedFile', error);

    return {
      success: false,
      error: 'Error procesando el archivo: ' + error.toString(),
      details: error.message || error.toString()
    };
  } finally {
    // Limpiar archivos temporales en el bloque finally para asegurar limpieza
    if (tempFile) {
      try {
        logInfo('Eliminando archivo Excel temporal...');
        retryOperation(function() {
          tempFile.setTrashed(true);
        }, 'Eliminar archivo Excel temporal', 2);
        logInfo('Archivo Excel temporal eliminado');
      } catch (cleanupError) {
        logError('Error al limpiar archivo Excel temporal', cleanupError);
      }
    }

    // Limpiar el archivo Google Sheets convertido si existe
    if (sheetId) {
      try {
        logInfo('Eliminando Google Sheets temporal...');
        retryOperation(function() {
          DriveApp.getFileById(sheetId).setTrashed(true);
        }, 'Eliminar Google Sheets temporal', 2);
        logInfo('Google Sheets temporal eliminado');
      } catch (cleanupError) {
        logError('Error al limpiar Google Sheets temporal', cleanupError);
      }
    }
  }
}

/**
 * Extrae un registro individual de una fila
 */
function extractRecord(headers, row, rowIndex) {
  // Función auxiliar para obtener valor por nombre de columna (parcial)
  function getVal(partialName) {
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i] ? headers[i].toString().toLowerCase().trim() : '';
      if (header.includes(partialName.toLowerCase())) {
        return row[i] !== undefined && row[i] !== null ? row[i].toString().trim() : '';
      }
    }
    return '';
  }
  
  // Función para obtener valor exacto por índice conocido
  function getByIndex(idx) {
    return row[idx] !== undefined && row[idx] !== null ? row[idx].toString().trim() : '';
  }
  
  // Formatear fecha
  function formatDate(val) {
    if (!val) return '';
    if (val instanceof Date) {
      var d = val.getDate().toString().padStart(2, '0');
      var m = (val.getMonth() + 1).toString().padStart(2, '0');
      var y = val.getFullYear();
      return d + '/' + m + '/' + y;
    }
    return val.toString();
  }
  
  // Formatear CI (sin puntos ni guiones)
  function formatCI(ci) {
    if (!ci) return '';
    return ci.toString().replace(/[.\-\s]/g, '');
  }
  
  // Extraer nombre y apellido del campo "Nombre completo" de padre/madre
  function splitFullName(fullName) {
    if (!fullName) return { nombre: '', apellido: '' };
    var parts = fullName.trim().split(' ');
    if (parts.length >= 4) {
      // Asumimos: Nombre1 Nombre2 Apellido1 Apellido2
      return {
        primerNombre: parts[0],
        segundoNombre: parts[1],
        primerApellido: parts[2],
        segundoApellido: parts.slice(3).join(' ')
      };
    } else if (parts.length === 3) {
      return {
        primerNombre: parts[0],
        segundoNombre: '',
        primerApellido: parts[1],
        segundoApellido: parts[2]
      };
    } else if (parts.length === 2) {
      return {
        primerNombre: parts[0],
        segundoNombre: '',
        primerApellido: parts[1],
        segundoApellido: ''
      };
    }
    return {
      primerNombre: fullName,
      segundoNombre: '',
      primerApellido: '',
      segundoApellido: ''
    };
  }
  
  // Mapear nivel/grado
  function parseNivelGrado(nivelGrado) {
    var ng = nivelGrado ? nivelGrado.toString().toLowerCase() : '';
    var result = { nivel: '', grado: '', modulo: 'P', curso: '' };
    
    // Patrones comunes
    if (ng.includes('maternal') || ng.includes('2 años')) {
      result.nivel = 'Inicial';
      result.grado = 'Maternal';
      result.modulo = 'P';
      result.curso = 'I2-EBI';
    } else if (ng.includes('3 años') || ng.includes('nivel 3')) {
      result.nivel = 'Inicial';
      result.grado = '3 años';
      result.modulo = 'P';
      result.curso = 'I3-EBI';
    } else if (ng.includes('4 años') || ng.includes('nivel 4')) {
      result.nivel = 'Inicial';
      result.grado = '4 años';
      result.modulo = 'P';
      result.curso = 'I4-EBI';
    } else if (ng.includes('5 años') || ng.includes('nivel 5')) {
      result.nivel = 'Inicial';
      result.grado = '5 años';
      result.modulo = 'P';
      result.curso = 'I5-EBI';
    } else if (ng.includes('1°') || ng.includes('1er') || ng.includes('primero') || ng.includes('1 ')) {
      if (ng.includes('ems') || ng.includes('bach') || ng.includes('human') || ng.includes('cient')) {
        result.nivel = 'Bachillerato';
        result.grado = '1° Bachillerato';
        result.modulo = 'L';
        result.curso = '1-BD';
      } else if (ng.includes('ciclo') || ng.includes('liceo')) {
        result.nivel = 'Ciclo Básico';
        result.grado = '1° CB';
        result.modulo = 'L';
        result.curso = '1-CB';
      } else {
        result.nivel = 'Primaria';
        result.grado = '1°';
        result.modulo = 'P';
        result.curso = '1-EBI';
      }
    } else if (ng.includes('2°') || ng.includes('2do') || ng.includes('segundo') || ng.includes('2 ')) {
      if (ng.includes('ems') || ng.includes('bach') || ng.includes('human') || ng.includes('cient')) {
        result.nivel = 'Bachillerato';
        result.grado = '2° Bachillerato';
        result.modulo = 'L';
        result.curso = ng.includes('human') ? '2-BDH' : '2-BDC';
      } else if (ng.includes('ciclo') || ng.includes('liceo')) {
        result.nivel = 'Ciclo Básico';
        result.grado = '2° CB';
        result.modulo = 'L';
        result.curso = '2-CB';
      } else {
        result.nivel = 'Primaria';
        result.grado = '2°';
        result.modulo = 'P';
        result.curso = '2-EBI';
      }
    } else if (ng.includes('3°') || ng.includes('3er') || ng.includes('tercero') || ng.includes('3 ')) {
      if (ng.includes('ems') || ng.includes('bach')) {
        result.nivel = 'Bachillerato';
        result.grado = '3° Bachillerato';
        result.modulo = 'L';
        result.curso = '3-BD';
      } else if (ng.includes('ciclo') || ng.includes('liceo')) {
        result.nivel = 'Ciclo Básico';
        result.grado = '3° CB';
        result.modulo = 'L';
        result.curso = '3-CB';
      } else {
        result.nivel = 'Primaria';
        result.grado = '3°';
        result.modulo = 'P';
        result.curso = '3-EBI';
      }
    } else if (ng.includes('4°') || ng.includes('4to') || ng.includes('cuarto') || ng.includes('4 ')) {
      result.nivel = 'Primaria';
      result.grado = '4°';
      result.modulo = 'P';
      result.curso = '4-EBI';
    } else if (ng.includes('5°') || ng.includes('5to') || ng.includes('quinto') || ng.includes('5 ')) {
      result.nivel = 'Primaria';
      result.grado = '5°';
      result.modulo = 'P';
      result.curso = '5-EBI';
    } else if (ng.includes('6°') || ng.includes('6to') || ng.includes('sexto') || ng.includes('6 ')) {
      result.nivel = 'Primaria';
      result.grado = '6°';
      result.modulo = 'P';
      result.curso = '6-EBI';
    }
    
    return result;
  }
  
  // Datos del estudiante (columnas conocidas)
  var estudiante = {
    nombre: getByIndex(2), // Nombre
    apellido: getByIndex(3), // Apellido
    fechaNacimiento: formatDate(row[4]), // Fecha de Nacimiento
    edad: getByIndex(5),
    ci: formatCI(getByIndex(6)), // Cédula de Identidad
    nacionalidad: getByIndex(7), // Nacionalidad
    domicilio: getByIndex(12), // Domicilio
    telefono: getByIndex(13), // Teléfono
    emailReferencia: getByIndex(14), // Email de referencia
    emailPropio: getByIndex(15), // Email propio
    asistenciaMedica: getByIndex(16), // Asistencia médica (mutualista)
    emergenciaMovil: getByIndex(17), // Emergencia móvil
    procedencia: getByIndex(18), // Procedencia
    nivelGrado: getByIndex(22), // Nivel / Grado al que se inscribe
    horario: getByIndex(20) // Horario (para Early Years)
  };
  
  // Datos del padre (columnas 23-36)
  var nombreCompletoPadre = getByIndex(23); // Nombre completo padre
  var padreNames = splitFullName(nombreCompletoPadre);
  
  var padre = {
    nombreCompleto: nombreCompletoPadre,
    primerNombre: padreNames.primerNombre,
    segundoNombre: padreNames.segundoNombre,
    primerApellido: padreNames.primerApellido,
    segundoApellido: padreNames.segundoApellido,
    ci: formatCI(getByIndex(24)), // Cédula de Identidad padre
    nacionalidad: getByIndex(25), // Nacionalidad padre
    profesion: getByIndex(26), // Profesión
    lugarTrabajo: getByIndex(27), // Lugar de trabajo
    telefono: getByIndex(28) // Teléfono/Celular padre
  };
  
  // Datos de la madre (columnas 30-36)
  var nombreCompletoMadre = getByIndex(30); // Nombre completo madre
  var madreNames = splitFullName(nombreCompletoMadre);
  
  var madre = {
    nombreCompleto: nombreCompletoMadre,
    primerNombre: madreNames.primerNombre,
    segundoNombre: madreNames.segundoNombre,
    primerApellido: madreNames.primerApellido,
    segundoApellido: madreNames.segundoApellido,
    ci: formatCI(getByIndex(31)), // Cédula de Identidad madre
    nacionalidad: getByIndex(32), // Nacionalidad madre
    profesion: getByIndex(33), // Profesión
    lugarTrabajo: getByIndex(34), // Lugar de trabajo
    telefono: getByIndex(35) // Teléfono/Celular madre
  };
  
  // Fechas de vencimiento
  var carneVacunas = formatDate(row[9]); // Carné de vacunas - Fecha de vencimiento
  var aptitudFisica = formatDate(row[10]); // Aptitud física - Fecha de vencimiento
  
  // Parsear nivel y grado
  var nivelGradoInfo = parseNivelGrado(estudiante.nivelGrado);
  
  // Responsable de pago
  var responsablePago = getByIndex(40);
  
  return {
    rowIndex: rowIndex,
    estudiante: estudiante,
    padre: padre,
    madre: madre,
    nivelGradoInfo: nivelGradoInfo,
    carneVacunas: carneVacunas,
    aptitudFisica: aptitudFisica,
    responsablePago: responsablePago,
    fechaInscripcion: formatDate(row[1]),
    // Validaciones
    validaciones: {
      tieneNombre: !!estudiante.nombre,
      tieneApellido: !!estudiante.apellido,
      tieneCI: !!estudiante.ci && estudiante.ci.length >= 7,
      tieneFechaNac: !!estudiante.fechaNacimiento,
      tieneNivel: !!nivelGradoInfo.nivel,
      tienePadre: !!padre.nombreCompleto || !!padre.ci,
      tieneMadre: !!madre.nombreCompleto || !!madre.ci,
      tieneContacto: !!estudiante.telefono || !!padre.telefono || !!madre.telefono
    }
  };
}

// ============================================
// GENERACIÓN DE ARCHIVOS
// ============================================

/**
 * Genera el archivo EVENTIFICA (VERSIÓN ROBUSTA)
 */
function generateEventifica(records) {
  var spreadsheet = null;
  var fileId = null;

  try {
    logInfo('Iniciando generación de archivo EVENTIFICA');

    // Validar entrada
    if (!records || records.length === 0) {
      throw new Error('No hay registros para generar el archivo');
    }

    logInfo('Generando archivo para ' + records.length + ' registros');

    // Crear spreadsheet con reintentos
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
    var fileName = 'Eventifica_Export_' + timestamp;

    spreadsheet = retryOperation(function() {
      return SpreadsheetApp.create(fileName);
    }, 'Crear spreadsheet EVENTIFICA');

    fileId = spreadsheet.getId();
    logInfo('Spreadsheet creado con ID: ' + fileId);

    var sheet = spreadsheet.getActiveSheet();
    retryOperation(function() {
      sheet.setName('Estudiantes_Padres');
    }, 'Renombrar hoja');

    // Headers según template_estudiantes_padres
    var headers = [
      'Nivel', 'Grado', 'Clase', 'Nombre estudiante', 'Apellido estudiante',
      'Documento estudiante', 'Sexo estudiante', 'Fecha de nacimiento estudiante',
      'Email estudiante', 'Teléfono estudiante', 'Dirección estudiante',
      'Autorización de uso de imagen estudiante', 'Usuario estudiante',
      'Contraseña estudiante', 'Cambiar contraseña estudiante',
      'Nombre padre', 'Apellido padre', 'Documento padre',
      'Fecha de nacimiento padre', 'Email padre', 'Teléfono padre',
      'Dirección padre', 'Autorización de uso de imagen padre',
      'Usuario padre', 'Contraseña padre', 'Cambiar contraseña padre',
      'Nombre madre', 'Apellido madre', 'Documento madre',
      'Fecha de nacimiento madre', 'Email madre', 'Teléfono madre',
      'Dirección madre', 'Autorización de uso de imagen madre',
      'Usuario madre', 'Contraseña madre', 'Cambiar contraseña madre'
    ];

    // Escribir headers con reintentos
    retryOperation(function() {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4');
      sheet.getRange(1, 1, 1, headers.length).setFontColor('white');
    }, 'Escribir headers EVENTIFICA');

    logInfo('Headers escritos exitosamente');

    // Preparar todos los datos en batch para mejorar rendimiento
    var allData = [];

    for (var i = 0; i < records.length; i++) {
      try {
        var r = records[i];

        // Extraer nombre y apellido del padre (con validación)
        var padreNombre = (r.padre.primerNombre || '').trim();
        if (r.padre.segundoNombre) padreNombre += ' ' + r.padre.segundoNombre.trim();
        var padreApellido = (r.padre.primerApellido || '').trim();
        if (r.padre.segundoApellido) padreApellido += ' ' + r.padre.segundoApellido.trim();

        // Extraer nombre y apellido de la madre (con validación)
        var madreNombre = (r.madre.primerNombre || '').trim();
        if (r.madre.segundoNombre) madreNombre += ' ' + r.madre.segundoNombre.trim();
        var madreApellido = (r.madre.primerApellido || '').trim();
        if (r.madre.segundoApellido) madreApellido += ' ' + r.madre.segundoApellido.trim();

        // Generar usuario y contraseña: CI sin puntos ni guiones + "hca"
        var ciLimpia = (r.estudiante.ci || '').replace(/[.\-\s]/g, '');
        var usuarioEstudiante = ciLimpia ? ciLimpia + 'hca' : '';
        var contrasenaEstudiante = ciLimpia ? ciLimpia + 'hca' : '';

        var ciPadreLimpia = (r.padre.ci || '').replace(/[.\-\s]/g, '');
        var usuarioPadre = ciPadreLimpia ? ciPadreLimpia + 'hca' : '';
        var contrasenaPadre = ciPadreLimpia ? ciPadreLimpia + 'hca' : '';

        var ciMadreLimpia = (r.madre.ci || '').replace(/[.\-\s]/g, '');
        var usuarioMadre = ciMadreLimpia ? ciMadreLimpia + 'hca' : '';
        var contrasenaMadre = ciMadreLimpia ? ciMadreLimpia + 'hca' : '';

        var rowData = [
          r.nivelGradoInfo.nivel || '',
          r.nivelGradoInfo.grado || '',
          '', // Clase
          r.estudiante.nombre || '',
          r.estudiante.apellido || '',
          ciLimpia, // CI sin puntos ni guiones
          '', // Sexo
          r.estudiante.fechaNacimiento || '',
          r.estudiante.emailPropio || r.estudiante.emailReferencia || '',
          r.estudiante.telefono || '',
          r.estudiante.domicilio || '',
          '', // Autorización imagen estudiante
          usuarioEstudiante, // Usuario estudiante: CI+hca
          contrasenaEstudiante, // Contraseña estudiante: CI+hca
          'SI', // Cambiar contraseña estudiante (forzar cambio en primer login)
          padreNombre,
          padreApellido,
          ciPadreLimpia, // CI padre sin puntos ni guiones
          '', // Fecha nacimiento padre
          r.estudiante.emailReferencia || '', // Email padre
          r.padre.telefono || '',
          r.estudiante.domicilio || '', // Dirección padre
          '', // Autorización imagen padre
          usuarioPadre, // Usuario padre: CI+hca
          contrasenaPadre, // Contraseña padre: CI+hca
          'SI', // Cambiar contraseña padre
          madreNombre,
          madreApellido,
          ciMadreLimpia, // CI madre sin puntos ni guiones
          '', // Fecha nacimiento madre
          '', // Email madre
          r.madre.telefono || '',
          r.estudiante.domicilio || '', // Dirección madre
          '', // Autorización imagen madre
          usuarioMadre, // Usuario madre: CI+hca
          contrasenaMadre, // Contraseña madre: CI+hca
          'SI'  // Cambiar contraseña madre
        ];

        allData.push(rowData);
      } catch (rowError) {
        logError('Error procesando registro ' + (i + 1) + ' para EVENTIFICA', rowError);
      }
    }

    // Escribir todos los datos de una vez (batch write)
    if (allData.length > 0) {
      retryOperation(function() {
        sheet.getRange(2, 1, allData.length, headers.length).setValues(allData);
      }, 'Escribir datos EVENTIFICA');

      logInfo('Datos escritos: ' + allData.length + ' registros');
    }

    // Ajustar columnas (con manejo de errores, no crítico)
    try {
      sheet.autoResizeColumns(1, Math.min(headers.length, 20)); // Limitar a 20 columnas para evitar timeout
    } catch (resizeError) {
      logError('Error ajustando columnas (no crítico)', resizeError);
    }

    // Verificar que el spreadsheet esté accesible
    retryOperation(function() {
      var testFile = DriveApp.getFileById(fileId);
      if (!testFile) throw new Error('No se pudo verificar el archivo creado');
    }, 'Verificar archivo creado');

    // Obtener URL de descarga
    var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
    var editUrl = spreadsheet.getUrl();

    logInfo('Archivo EVENTIFICA generado exitosamente: ' + fileId);

    return {
      success: true,
      fileId: fileId,
      fileName: fileName,
      downloadUrl: url,
      editUrl: editUrl,
      recordsWritten: allData.length
    };

  } catch (error) {
    logError('Error en generateEventifica', error);

    // Intentar limpiar el archivo si se creó pero falló
    if (fileId) {
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
        logInfo('Archivo parcial eliminado tras error');
      } catch (cleanupError) {
        logError('Error limpiando archivo parcial', cleanupError);
      }
    }

    return {
      success: false,
      error: 'Error generando archivo EVENTIFICA: ' + error.toString(),
      details: error.message || error.toString()
    };
  }
}

/**
 * Genera el archivo AlumnosYFamilias (VERSIÓN ROBUSTA)
 */
function generateAlumnosFamilias(records) {
  var spreadsheet = null;
  var fileId = null;

  try {
    logInfo('Iniciando generación de archivo AlumnosYFamilias');

    // Validar entrada
    if (!records || records.length === 0) {
      throw new Error('No hay registros para generar el archivo');
    }

    logInfo('Generando archivo para ' + records.length + ' registros');

    // Crear spreadsheet con reintentos
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
    var fileName = 'SIGED_Plantilla_Importar_AlumnosYFamilias_' + timestamp;

    spreadsheet = retryOperation(function() {
      return SpreadsheetApp.create(fileName);
    }, 'Crear spreadsheet SIGED');

    fileId = spreadsheet.getId();
    logInfo('Spreadsheet creado con ID: ' + fileId);

    var sheet = spreadsheet.getActiveSheet();
    retryOperation(function() {
      sheet.setName('Alumnos');
    }, 'Renombrar hoja');
    
    // Headers EXACTOS según Plantilla_Importar_AlumnosYFamilias (28).xlsx
    // IMPORTANTE: Copiados directamente de la plantilla oficial de SIGED
    var headers = [
      'FamNro','FamApe','FamEstCiv','FamAAMat','DomLocCod','DomBarCod','DomCalle','DomNroPuerta','DomTel','FamNroCtaBan','FamFecVin',
      'PadPnaDoc','','PadPnaTDoc','PadPnaDocPaiCod','PadPnaPriNom','PadPnaSegNom','PadPnaPriApe','PadPnaSegApe','PadPnaFecNac','PadPnaNacPaiCod','PadPnaNacDepCod','PadPnaNacLug','PadPnaNacionalidad','PadPnaEstCiv','PadPnaNupcias','PadPnaEMail','PadDomLocCod','PadDomBarCod','PadDomCalle','PadDomNroPuerta','PadDomTel','PadPnaTelCel','PadPnaExAlumno','PnaGenEgre','PadProCod','PadPnaOcu','PadPnaEmp','PadPnaTelLab','PadPnaHor','PadPnaForIdPri','PadPnaInstPrimaria','PadPnaForIdSec','PadPnaInstSecundaria','PadPnaForIdNiv','PadPnaNivForEsp','PadPnaRel','PadPnaSerCre','PadPnaNroCre','PadPnaFallecido','PadPnaFecFall','Bautizado','Confirmado','Casado Iglesia','Casado Civil','PadPnaIdExterno',
      'MadPnaDoc','','MadPnaTDoc','MadPnaDocPaiCod','MadPnaPriNom','MadPnaSegNom','MadPnaPriApe','MadPnaSegApe','MadPnaFecNac','MadPnaNacPaiCod','MadPnaNacDepCod','MadPnaNacLug','MadPnaNacionalidad','MadPnaEstCiv','MadPnaNupcias','MadPnaEMail','MadDomLocCod','MadDomBarCod','MadDomCalle','MadDomNroPuerta','MadDomTel','MadPnaTelCel','MadPnaExAlumno','PnaGenEgre','MadProCod','MadPnaOcu','MadPnaEmp','MadPnaTelLab','MadPnaHor','MadPnaForIdPri','MadPnaInstPrimaria','MadPnaForIdSec','MadPnaInstSecundaria','MadPnaForIdNiv','MadPnaNivForEsp','MadPnaRel','MadPnaSerCre','MadPnaNroCre','MadPnaFallecido','MadPnaFecFall','Bautizado','Confirmado','Casado Iglesia','Casado Civil','MadPnaIdExterno',
      'FaluIDLiceo','FAluMat','FAluMatEsc','FAluDoc','FAluTDoc','FAluDocPaiCod','FAluPriApe','FAluSegApe','FAluPriNom','FAluSegNom','FAluFecNac','FAluSexo','FAluNacPaiCod','FAluNacDepCod','FAluNacionalidad','FAluTelCasa','FAluTelCelular','FAluTel1','FAluPer1','FAluTel2','FAluPer2','FAluCP','FAluSecJud','FAluFecIng','FIngCod','FAluComFIng','FAluJBLicCod','FAluFecJB','FAluSerCre','FAluNroCre','FAluEmail','FAluDirLocCod','FAluDirBarCod','FAluCalle','FAluNroPuerta','FAluApto','FAluComDir','Modulo','Alec','InsCurC3PId','InsCurFec','TurCod','GruCod','FAluRelBau','FAluRel1Com','FAluRelConf','FAluRelLugFor','FAluNotEsc','FAluMMVac','FAluAAVac','FAluFecVenCS','FAluFecVenCI','MutCod','FAluNroAfi','EmerCod','FAluObs','FAluMed','FAluAle','FAluDis','FAluOid','FAluVis','FAluAsm','FAluDiabetes','FAluCeliaco','FAluGruSan','FAluVivCon','FAluACargo','FAluTieEscPri','FAluTieEscPub','FAluPadrastro','FAluMadrastra','FAluTutor','FAluTipHijo','FAluDueSolo','FAluDueCon','FAluReligion','FAluPubMatGra','FAluPubMatGraObs'
    ];
    
    // Escribir headers con reintentos
    retryOperation(function() {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.getRange(1, 1, 1, headers.length).setBackground('#34a853');
      sheet.getRange(1, 1, 1, headers.length).setFontColor('white');
    }, 'Escribir headers AlumnosYFamilias');

    logInfo('Headers escritos exitosamente');
    
    // Mapear nacionalidad a código de país
    function getNacionalidadCod(nac) {
      if (!nac) return '858'; // Uruguay por defecto
      var n = nac.toLowerCase();
      if (n.includes('uruguay') || n.includes('oriental')) return '858';
      if (n.includes('argentin')) return '32';
      if (n.includes('brasil')) return '76';
      if (n.includes('chile')) return '152';
      if (n.includes('paragua')) return '600';
      return '858';
    }
    
    // Mapear mutualista
    function getMutCod(mut) {
      if (!mut) return '';
      var m = mut.toLowerCase();
      if (m.includes('casmu')) return 'CASMU';
      if (m.includes('medica')) return 'MU';
      if (m.includes('española')) return 'AES';
      if (m.includes('smu') || m.includes('servicio')) return 'SMU';
      if (m.includes('circulo')) return 'CCOU';
      if (m.includes('evangel')) return 'HE';
      if (m.includes('mp') || m.includes('policial')) return 'MP';
      if (m.includes('militar')) return 'HM';
      if (m.includes('asse')) return 'ASSE';
      return '';
    }
    
    // Mapear emergencia
    function getEmerCod(emer) {
      if (!emer) return '';
      var e = emer.toLowerCase();
      if (e.includes('suat')) return 'SUAT';
      if (e.includes('semm')) return 'SEMM';
      if (e.includes('1727') || e.includes('uca')) return '1727';
      if (e.includes('ucm')) return 'UCM';
      if (e.includes('mp')) return 'MP';
      return '';
    }
    
    // Mapear turno
    function getTurnoCod(horario) {
      if (!horario) return '';
      var h = horario.toLowerCase();
      if (h.includes('matut') || h.includes('mañana')) return '1';
      if (h.includes('vespert') || h.includes('tarde')) return '2';
      if (h.includes('intermedio')) return '3';
      return '';
    }
    
    // Extraer mes y año de fecha de vencimiento de vacunas
    function getVacunasMesAnio(fecha) {
      if (!fecha) return { mes: '', anio: '' };
      var parts = fecha.split('/');
      if (parts.length >= 3) {
        return { mes: parseInt(parts[1]), anio: parseInt(parts[2]) };
      }
      return { mes: '', anio: '' };
    }
    
    // Separar apellidos del estudiante
    function separarApellidos(apellido) {
      if (!apellido) return { primero: '', segundo: '' };
      var parts = apellido.trim().split(' ');
      if (parts.length >= 2) {
        return { primero: parts[0], segundo: parts.slice(1).join(' ') };
      }
      return { primero: apellido, segundo: '' };
    }
    
    // Separar nombres del estudiante
    function separarNombres(nombre) {
      if (!nombre) return { primero: '', segundo: '' };
      var parts = nombre.trim().split(' ');
      if (parts.length >= 2) {
        return { primero: parts[0], segundo: parts.slice(1).join(' ') };
      }
      return { primero: nombre, segundo: '' };
    }
    
    // Preparar todos los datos en batch para mejorar rendimiento
    var allData = [];

    for (var i = 0; i < records.length; i++) {
      try {
        var r = records[i];

        var apellidos = separarApellidos(r.estudiante.apellido);
        var nombres = separarNombres(r.estudiante.nombre);
        var vacunas = getVacunasMesAnio(r.carneVacunas);
        var nacCod = getNacionalidadCod(r.estudiante.nacionalidad);
        var padreNacCod = getNacionalidadCod(r.padre.nacionalidad);
        var madreNacCod = getNacionalidadCod(r.madre.nacionalidad);
        var famApellido = apellidos.primero;
        if (r.madre.primerApellido) {
          famApellido += ' ' + r.madre.primerApellido;
        }
      
      var rowData = [
        '', // FamNro - se asigna automáticamente
        famApellido.trim(), // FamApe
        '', // FamEstCiv
        '', // FamAAMat
        'A 1', // DomLocCod - Montevideo
        '', // DomBarCod
        r.estudiante.domicilio, // DomCalle
        '', // DomNroPuerta
        r.estudiante.telefono, // DomTel
        '', // FamNroCtaBan
        r.fechaInscripcion, // FamFecVin
        // PADRE
        r.padre.ci, // PadPnaDoc
        '', // Sobrescribir
        'CI', // PadPnaTDoc
        '858', // PadPnaDocPaiCod
        r.padre.primerNombre, // PadPnaPriNom
        r.padre.segundoNombre, // PadPnaSegNom
        r.padre.primerApellido, // PadPnaPriApe
        r.padre.segundoApellido, // PadPnaSegApe
        '', // PadPnaFecNac
        padreNacCod, // PadPnaNacPaiCod
        '', // PadPnaNacDepCod
        '', // PadPnaNacLug
        r.padre.nacionalidad, // PadPnaNacionalidad
        '', // PadPnaEstCiv
        '', // PadPnaNupcias
        r.estudiante.emailReferencia, // PadPnaEMail
        'A 1', // PadDomLocCod
        '', // PadDomBarCod
        r.estudiante.domicilio, // PadDomCalle
        '', // PadDomNroPuerta
        '', // PadDomTel
        r.padre.telefono, // PadPnaTelCel
        'N', // PadPnaExAlumno
        '', // PnaGenEgre
        '', // PadProCod
        r.padre.profesion, // PadPnaOcu
        r.padre.lugarTrabajo, // PadPnaEmp
        '', // PadPnaTelLab
        '', // PadPnaHor
        '', '', '', '', '', '', '', '', '', // Estudios padre
        'N', // PadPnaFallecido
        '', // PadPnaFecFall
        '', '', '', '', // Bautizado, Confirmado, etc
        '', // PadPnaIdExterno
        // MADRE
        r.madre.ci, // MadPnaDoc
        '', // Sobrescribir
        'CI', // MadPnaTDoc
        '858', // MadPnaDocPaiCod
        r.madre.primerNombre, // MadPnaPriNom
        r.madre.segundoNombre, // MadPnaSegNom
        r.madre.primerApellido, // MadPnaPriApe
        r.madre.segundoApellido, // MadPnaSegApe
        '', // MadPnaFecNac
        madreNacCod, // MadPnaNacPaiCod
        '', // MadPnaNacDepCod
        '', // MadPnaNacLug
        r.madre.nacionalidad, // MadPnaNacionalidad
        '', // MadPnaEstCiv
        '', // MadPnaNupcias
        '', // MadPnaEMail
        'A 1', // MadDomLocCod
        '', // MadDomBarCod
        r.estudiante.domicilio, // MadDomCalle
        '', // MadDomNroPuerta
        '', // MadDomTel
        r.madre.telefono, // MadPnaTelCel
        'N', // MadPnaExAlumno
        '', // PnaGenEgre
        '', // MadProCod
        r.madre.profesion, // MadPnaOcu
        r.madre.lugarTrabajo, // MadPnaEmp
        '', // MadPnaTelLab
        '', // MadPnaHor
        '', '', '', '', '', '', '', '', '', // Estudios madre
        'N', // MadPnaFallecido
        '', // MadPnaFecFall
        '', '', '', '', // Bautizado, Confirmado, etc
        '', // MadPnaIdExterno
        // ALUMNO
        '', // FaluIDLiceo
        '', // FAluMat
        '', // FAluMatEsc
        r.estudiante.ci, // FAluDoc
        'CI', // FAluTDoc
        '858', // FAluDocPaiCod
        apellidos.primero, // FAluPriApe
        apellidos.segundo, // FAluSegApe
        nombres.primero, // FAluPriNom
        nombres.segundo, // FAluSegNom
        r.estudiante.fechaNacimiento, // FAluFecNac
        '', // FAluSexo
        nacCod, // FAluNacPaiCod
        '', // FAluNacDepCod
        r.estudiante.nacionalidad, // FAluNacionalidad
        r.estudiante.telefono, // FAluTelCasa
        '', // FAluTelCelular
        '', // FAluTel1
        '', // FAluPer1
        '', // FAluTel2
        '', // FAluPer2
        '', // FAluCP
        '', // FAluSecJud
        r.fechaInscripcion, // FAluFecIng
        '', // FIngCod
        r.estudiante.procedencia, // FAluComFIng
        '', // FAluJBLicCod
        '', // FAluFecJB
        '', // FAluSerCre
        '', // FAluNroCre
        r.estudiante.emailPropio || r.estudiante.emailReferencia, // FAluEmail
        'A 1', // FAluDirLocCod
        '', // FAluDirBarCod
        r.estudiante.domicilio, // FAluCalle
        '', // FAluNroPuerta
        '', // FAluApto
        '', // FAluComDir
        r.nivelGradoInfo.modulo, // Modulo
        new Date().getFullYear(), // Alec (Año lectivo)
        r.nivelGradoInfo.curso, // InsCurC3PId
        r.fechaInscripcion, // InsCurFec
        getTurnoCod(r.estudiante.horario), // TurCod
        '', // GruCod
        '', '', '', '', '', // Religión
        vacunas.mes, // FAluMMVac
        vacunas.anio, // FAluAAVac
        r.aptitudFisica, // FAluFecVenCS
        '', // FAluFecVenCI
        getMutCod(r.estudiante.asistenciaMedica), // MutCod
        '', // FAluNroAfi
        getEmerCod(r.estudiante.emergenciaMovil), // EmerCod
        '', // FAluObs
        '', // FAluMed
        '', // FAluAle
        '', '', '', '', '', '', '', '', '', '', // FAluDis, FAluOid, FAluVis, FAluAsm, FAluDiabetes, FAluCeliaco, FAluGruSan, FAluVivCon, FAluACargo, FAluTieEscPri
        '', '', '', '', '', // FAluTieEscPub, FAluPadrastro, FAluMadrastra, FAluTutor, FAluTipHijo
        '', '', '', // FAluDueSolo, FAluDueCon, FAluReligion
        'S', // FAluPubMatGra
        '' // FAluPubMatGraObs
      ];

        allData.push(rowData);
      } catch (rowError) {
        logError('Error procesando registro ' + (i + 1) + ' para AlumnosYFamilias', rowError);
      }
    }

    // Escribir todos los datos de una vez (batch write)
    if (allData.length > 0) {
      retryOperation(function() {
        sheet.getRange(2, 1, allData.length, headers.length).setValues(allData);
      }, 'Escribir datos AlumnosYFamilias');

      logInfo('Datos escritos: ' + allData.length + ' registros');
    }

    // Ajustar columnas (con manejo de errores, no crítico)
    try {
      sheet.autoResizeColumns(1, 20);
    } catch (resizeError) {
      logError('Error ajustando columnas (no crítico)', resizeError);
    }

    // Verificar que el spreadsheet esté accesible
    retryOperation(function() {
      var testFile = DriveApp.getFileById(fileId);
      if (!testFile) throw new Error('No se pudo verificar el archivo creado');
    }, 'Verificar archivo creado');

    // Obtener URL de descarga
    var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
    var editUrl = spreadsheet.getUrl();

    logInfo('Archivo SIGED generado exitosamente: ' + fileId);
    logInfo('URL de descarga SIGED: ' + url);
    logInfo('URL de edición SIGED: ' + editUrl);

    return {
      success: true,
      fileId: fileId,
      fileName: fileName,
      downloadUrl: url,
      editUrl: editUrl,
      recordsWritten: allData.length,
      systemName: 'SIGED'
    };

  } catch (error) {
    logError('Error en generateAlumnosFamilias', error);

    // Intentar limpiar el archivo si se creó pero falló
    if (fileId) {
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
        logInfo('Archivo parcial eliminado tras error');
      } catch (cleanupError) {
        logError('Error limpiando archivo parcial', cleanupError);
      }
    }

    return {
      success: false,
      error: 'Error generando archivo AlumnosYFamilias: ' + error.toString(),
      details: error.message || error.toString()
    };
  }
}

/**
 * Función de diagnóstico para verificar el sistema
 */
function diagnosticoCompleto() {
  Logger.log('=== DIAGNÓSTICO DEL SISTEMA ===');
  Logger.log('Versión: 2.1');
  Logger.log('');

  // 1. Verificar permisos
  Logger.log('1. Verificando permisos...');
  var permisos = checkPermissions();
  Logger.log('   Resultado: ' + (permisos.success ? '✓ OK' : '✗ ERROR'));
  if (!permisos.success) {
    Logger.log('   Error: ' + permisos.error);
  }

  // 2. Verificar Drive API
  Logger.log('');
  Logger.log('2. Verificando Drive API...');
  try {
    var testFile = DriveApp.createFile('TEST_DRIVE_API', 'test');
    var testId = testFile.getId();
    testFile.setTrashed(true);
    Logger.log('   Drive API: ✓ OK');
  } catch (e) {
    Logger.log('   Drive API: ✗ ERROR - ' + e.toString());
  }

  // 3. Limpiar archivos temporales
  Logger.log('');
  Logger.log('3. Limpiando archivos temporales...');
  cleanupOldTempFiles();

  // 4. Verificar configuración
  Logger.log('');
  Logger.log('4. Configuración actual:');
  Logger.log('   MAX_RETRIES: ' + CONFIG.MAX_RETRIES);
  Logger.log('   RETRY_DELAY: ' + CONFIG.RETRY_DELAY + 'ms');
  Logger.log('   LOG_ENABLED: ' + CONFIG.LOG_ENABLED);

  Logger.log('');
  Logger.log('=== FIN DIAGNÓSTICO ===');

  return {
    success: permisos.success,
    timestamp: new Date().toISOString()
  };
}

/**
 * Genera ambos archivos (VERSIÓN ROBUSTA)
 */
function generateBothFiles(records) {
  logInfo('=== INICIANDO GENERACIÓN DE ARCHIVOS ===');
  logInfo('Generando 2 archivos: EVENTIFICA y SIGED');

  try {
    // Validar entrada
    if (!records || records.length === 0) {
      throw new Error('No hay registros para generar los archivos');
    }

    logInfo('Generando 2 archivos para ' + records.length + ' registros');

    // Generar archivo EVENTIFICA
    var eventificaResult = null;
    try {
      eventificaResult = generateEventifica(records);
      if (eventificaResult.success) {
        logInfo('Archivo EVENTIFICA generado exitosamente');
      } else {
        logError('Error generando archivo EVENTIFICA', eventificaResult.error);
      }
    } catch (eventificaError) {
      logError('Excepción generando archivo EVENTIFICA', eventificaError);
      eventificaResult = {
        success: false,
        error: 'Excepción al generar EVENTIFICA: ' + eventificaError.toString()
      };
    }

    // Generar archivo AlumnosYFamilias
    var alumnosResult = null;
    try {
      alumnosResult = generateAlumnosFamilias(records);
      if (alumnosResult.success) {
        logInfo('Archivo AlumnosYFamilias generado exitosamente');
      } else {
        logError('Error generando archivo AlumnosYFamilias', alumnosResult.error);
      }
    } catch (alumnosError) {
      logError('Excepción generando archivo AlumnosYFamilias', alumnosError);
      alumnosResult = {
        success: false,
        error: 'Excepción al generar AlumnosYFamilias: ' + alumnosError.toString()
      };
    }

    var result = {
      eventifica: eventificaResult,
      alumnos: alumnosResult,
      timestamp: new Date().toISOString()
    };

    // Verificar si al menos uno fue exitoso
    if (eventificaResult && eventificaResult.success) {
      result.overallSuccess = true;
      result.message = 'Al menos un archivo fue generado exitosamente';
    } else if (alumnosResult && alumnosResult.success) {
      result.overallSuccess = true;
      result.message = 'Al menos un archivo fue generado exitosamente';
    } else {
      result.overallSuccess = false;
      result.message = 'Error: No se pudo generar ningún archivo';
    }

    logInfo('Generación completada. Estado: ' + (result.overallSuccess ? 'Éxito parcial o total' : 'Fallo'));

    return result;

  } catch (error) {
    logError('Error general en generateBothFiles', error);
    return {
      eventifica: { success: false, error: 'Error general: ' + error.toString() },
      alumnos: { success: false, error: 'Error general: ' + error.toString() },
      overallSuccess: false,
      message: 'Error crítico al generar archivos'
    };
  }
}
