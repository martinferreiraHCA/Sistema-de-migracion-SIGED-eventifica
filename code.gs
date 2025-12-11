/**
 * Aplicación Web para Conversión de Fichas de Inscripción
 * Colegio Hans Christian Andersen
 * 
 * Convierte datos de "Ficha de Inscripción" a:
 * 1. Formato EVENTIFICA (template_estudiantes_padres)
 * 2. Formato AlumnosYFamilias (Plantilla_Importar)
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Conversor de Fichas de Inscripción - HCA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Procesa el archivo subido y extrae los datos
 */
function processUploadedFile(base64Data, fileName) {
  try {
    // Decodificar el archivo
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName);
    
    // Crear archivo temporal en Drive
    var tempFile = DriveApp.createFile(blob);
    var spreadsheet = SpreadsheetApp.open(tempFile);
    var sheet = spreadsheet.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    
    // Eliminar archivo temporal
    tempFile.setTrashed(true);
    
    // Procesar datos
    var headers = data[0];
    var records = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var record = extractRecord(headers, row, i);
      if (record && record.estudiante.nombre) {
        records.push(record);
      }
    }
    
    return {
      success: true,
      records: records,
      totalRows: data.length - 1
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
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

/**
 * Genera el archivo EVENTIFICA
 */
function generateEventifica(records) {
  try {
    var spreadsheet = SpreadsheetApp.create('Eventifica_Export_' + new Date().toISOString().slice(0,10));
    var sheet = spreadsheet.getActiveSheet();
    sheet.setName('Estudiantes_Padres');
    
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
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('white');
    
    // Datos
    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      
      // Extraer nombre y apellido del padre
      var padreNombre = r.padre.primerNombre;
      if (r.padre.segundoNombre) padreNombre += ' ' + r.padre.segundoNombre;
      var padreApellido = r.padre.primerApellido;
      if (r.padre.segundoApellido) padreApellido += ' ' + r.padre.segundoApellido;
      
      // Extraer nombre y apellido de la madre
      var madreNombre = r.madre.primerNombre;
      if (r.madre.segundoNombre) madreNombre += ' ' + r.madre.segundoNombre;
      var madreApellido = r.madre.primerApellido;
      if (r.madre.segundoApellido) madreApellido += ' ' + r.madre.segundoApellido;
      
      var rowData = [
        r.nivelGradoInfo.nivel,
        r.nivelGradoInfo.grado,
        '', // Clase
        r.estudiante.nombre,
        r.estudiante.apellido,
        r.estudiante.ci,
        '', // Sexo
        r.estudiante.fechaNacimiento,
        r.estudiante.emailPropio || r.estudiante.emailReferencia,
        r.estudiante.telefono,
        r.estudiante.domicilio,
        '', // Autorización imagen estudiante
        '', // Usuario estudiante
        '', // Contraseña estudiante
        '', // Cambiar contraseña estudiante
        padreNombre.trim(),
        padreApellido.trim(),
        r.padre.ci,
        '', // Fecha nacimiento padre
        r.estudiante.emailReferencia, // Email padre
        r.padre.telefono,
        r.estudiante.domicilio, // Dirección padre
        '', // Autorización imagen padre
        '', // Usuario padre
        '', // Contraseña padre
        '', // Cambiar contraseña padre
        madreNombre.trim(),
        madreApellido.trim(),
        r.madre.ci,
        '', // Fecha nacimiento madre
        '', // Email madre
        r.madre.telefono,
        r.estudiante.domicilio, // Dirección madre
        '', // Autorización imagen madre
        '', // Usuario madre
        '', // Contraseña madre
        ''  // Cambiar contraseña madre
      ];
      
      sheet.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
    }
    
    // Ajustar columnas
    sheet.autoResizeColumns(1, headers.length);
    
    // Obtener URL de descarga
    var fileId = spreadsheet.getId();
    var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
    
    return {
      success: true,
      fileId: fileId,
      fileName: spreadsheet.getName(),
      downloadUrl: url,
      editUrl: spreadsheet.getUrl()
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Genera el archivo AlumnosYFamilias
 */
function generateAlumnosFamilias(records) {
  try {
    var spreadsheet = SpreadsheetApp.create('AlumnosYFamilias_Export_' + new Date().toISOString().slice(0,10));
    var sheet = spreadsheet.getActiveSheet();
    sheet.setName('Alumnos');
    
    // Headers principales según Plantilla_Importar_AlumnosYFamilias
    var headers = [
      'FamNro', 'FamApe', 'FamEstCiv', 'FamAAMat', 'DomLocCod', 'DomBarCod',
      'DomCalle', 'DomNroPuerta', 'DomTel', 'FamNroCtaBan', 'FamFecVin',
      // Padre (columnas 11-55)
      'PadPnaDoc', '', 'PadPnaTDoc', 'PadPnaDocPaiCod', 'PadPnaPriNom', 'PadPnaSegNom',
      'PadPnaPriApe', 'PadPnaSegApe', 'PadPnaFecNac', 'PadPnaNacPaiCod', 'PadPnaNacDepCod',
      'PadPnaNacLug', 'PadPnaNacionalidad', 'PadPnaEstCiv', 'PadPnaNupcias', 'PadPnaEMail',
      'PadDomLocCod', 'PadDomBarCod', 'PadDomCalle', 'PadDomNroPuerta', 'PadDomTel',
      'PadPnaTelCel', 'PadPnaExAlumno', 'PnaGenEgre', 'PadProCod', 'PadPnaOcu', 'PadPnaEmp',
      'PadPnaTelLab', 'PadPnaHor', 'PadPnaForIdPri', 'PadPnaInstPrimaria', 'PadPnaForIdSec',
      'PadPnaInstSecundaria', 'PadPnaForIdNiv', 'PadPnaNivForEsp', 'PadPnaRel', 'PadPnaSerCre',
      'PadPnaNroCre', 'PadPnaFallecido', 'PadPnaFecFall', 'Bautizado', 'Confirmado',
      'Casado Iglesia', 'Casado Civil', 'PadPnaIdExterno',
      // Madre (columnas 56-100)
      'MadPnaDoc', '', 'MadPnaTDoc', 'MadPnaDocPaiCod', 'MadPnaPriNom', 'MadPnaSegNom',
      'MadPnaPriApe', 'MadPnaSegApe', 'MadPnaFecNac', 'MadPnaNacPaiCod', 'MadPnaNacDepCod',
      'MadPnaNacLug', 'MadPnaNacionalidad', 'MadPnaEstCiv', 'MadPnaNupcias', 'MadPnaEMail',
      'MadDomLocCod', 'MadDomBarCod', 'MadDomCalle', 'MadDomNroPuerta', 'MadDomTel',
      'MadPnaTelCel', 'MadPnaExAlumno', 'PnaGenEgre', 'MadProCod', 'MadPnaOcu', 'MadPnaEmp',
      'MadPnaTelLab', 'MadPnaHor', 'MadPnaForIdPri', 'MadPnaInstPrimaria', 'MadPnaForIdSec',
      'MadPnaInstSecundaria', 'MadPnaForIdNiv', 'MadPnaNivForEsp', 'MadPnaRel', 'MadPnaSerCre',
      'MadPnaNroCre', 'MadPnaFallecido', 'MadPnaFecFall', 'Bautizado', 'Confirmado',
      'Casado Iglesia', 'Casado Civil', 'MadPnaIdExterno',
      // Alumno (columnas 101-179)
      'FaluIDLiceo', 'FAluMat', 'FAluMatEsc', 'FAluDoc', 'FAluTDoc', 'FAluDocPaiCod',
      'FAluPriApe', 'FAluSegApe', 'FAluPriNom', 'FAluSegNom', 'FAluFecNac', 'FAluSexo',
      'FAluNacPaiCod', 'FAluNacDepCod', 'FAluNacionalidad', 'FAluTelCasa', 'FAluTelCelular',
      'FAluTel1', 'FAluPer1', 'FAluTel2', 'FAluPer2', 'FAluCP', 'FAluSecJud', 'FAluFecIng',
      'FIngCod', 'FAluComFIng', 'FAluJBLicCod', 'FAluFecJB', 'FAluSerCre', 'FAluNroCre',
      'FAluEmail', 'FAluDirLocCod', 'FAluDirBarCod', 'FAluCalle', 'FAluNroPuerta', 'FAluApto',
      'FAluComDir', 'Modulo', 'Alec', 'InsCurC3PId', 'InsCurFec', 'TurCod', 'GruCod',
      'FAluRelBau', 'FAluRel1Com', 'FAluRelConf', 'FAluRelLugFor', 'FAluNotEsc',
      'FAluMMVac', 'FAluAAVac', 'FAluFecVenCS', 'FAluFecVenCI', 'MutCod', 'FAluNroAfi',
      'EmerCod', 'FAluObs', 'FAluMed', 'FAluAle', 'FAluDis', 'FAluOid', 'FAluVis', 'FAluAsm',
      'FAluDiabetes', 'FAluCeliaco', 'FAluGruSan', 'FAluVivCon', 'FAluACargo', 'FAluTieEscPri',
      'FAluTieEscPub', 'FAluPadrastro', 'FAluMadrastra', 'FAluTutor', 'FAluTipHijo',
      'FAluDueSolo', 'FAluDueCon', 'FAluReligion', 'FAluPubMatGra', 'FAluPubMatGraObs'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#34a853');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('white');
    
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
    
    // Datos
    for (var i = 0; i < records.length; i++) {
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
        '', '', '', '', '', '', '', '', '', // Condiciones médicas
        '', '', '', '', '', '', '', '', '', '', // Convivencia
        'S', // FAluPubMatGra
        '' // FAluPubMatGraObs
      ];
      
      sheet.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
    }
    
    // Ajustar columnas (solo algunas para no hacer muy lento)
    try {
      sheet.autoResizeColumns(1, 20);
    } catch (e) {}
    
    // Obtener URL de descarga
    var fileId = spreadsheet.getId();
    var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
    
    return {
      success: true,
      fileId: fileId,
      fileName: spreadsheet.getName(),
      downloadUrl: url,
      editUrl: spreadsheet.getUrl()
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Genera ambos archivos
 */
function generateBothFiles(records) {
  var eventificaResult = generateEventifica(records);
  var alumnosResult = generateAlumnosFamilias(records);
  
  return {
    eventifica: eventificaResult,
    alumnos: alumnosResult
  };
}
