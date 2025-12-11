# Sistema de Migración SIGED-Eventifica

## Versión 2.0 - Sistema Robusto

Sistema mejorado para convertir fichas de inscripción del Colegio Hans Christian Andersen a formatos EVENTIFICA y AlumnosYFamilias.

---

## ✨ Mejoras Implementadas

### 🛡️ Manejo de Errores Robusto
- **Reintentos automáticos**: Sistema de reintentos exponenciales para operaciones críticas
- **Validaciones exhaustivas**: Verificación de datos en cada paso del proceso
- **Limpieza automática**: Eliminación de archivos temporales y parciales en caso de error
- **Logging detallado**: Registro completo de operaciones para debugging

### 🚀 Mejoras de Rendimiento
- **Batch Writing**: Escritura de datos en lote para mejorar velocidad
- **Operaciones optimizadas**: Reducción de llamadas individuales a la API
- **Gestión de memoria**: Limpieza automática de archivos temporales antiguos

### 🔒 Seguridad y Permisos
- **Permisos explícitos**: Configuración clara de OAuth scopes
- **Validación de acceso**: Verificación de permisos antes de operaciones críticas
- **Manejo seguro de archivos**: Protección contra pérdida de datos

---

## 📋 Solución al Error de Google Sheets API

### Error Principal Solucionado
```
Error: El servicio Hojas de cálculo falló al acceder al documento con el ID XXXXX
```

**Este error ocurría cuando:**
- El sistema intentaba abrir un archivo Excel directamente como Google Sheets
- No había conversión explícita del formato Excel a formato nativo de Google
- Problemas de timing entre creación y apertura del archivo

**Solución Implementada (v2.0):**
- ✅ Conversión explícita de Excel a Google Sheets usando Drive API
- ✅ Espera de 2 segundos para que Google procese la conversión
- ✅ Método alternativo (fallback) si Drive API falla
- ✅ Reintentos automáticos con delays exponenciales

### Otras Causas y Soluciones

#### 1. **Permisos Insuficientes**
**Solución:**
1. Ve a tu proyecto de Google Apps Script
2. Ejecuta: **Extensiones > Apps Script**
3. Clic en "Ejecutar" en cualquier función (ej: `checkPermissions`)
4. Autoriza todos los permisos solicitados:
   - Google Drive
   - Google Sheets
   - Crear y modificar archivos

#### 2. **Timeout de la API**
**Solución automática implementada:**
- El sistema ahora reintenta automáticamente 3 veces con delays exponenciales
- Si una operación falla, espera 1s, 2s, 4s antes de reintentar
- Logs detallados muestran el progreso de cada intento

#### 3. **Archivo Temporal Corrupto**
**Solución automática implementada:**
- Limpieza automática de archivos temporales al iniciar
- Validación de archivos antes de procesamiento
- Eliminación segura en bloque `finally`

#### 4. **Límites de Cuota de Google**
**Solución:**
- Implementado batch writing para reducir llamadas a API
- Optimización de operaciones para usar menos recursos
- Si persiste, espera unos minutos y reintenta

---

## 🚀 Instalación

### Opción 1: Nuevo Proyecto
1. Ve a [Google Apps Script](https://script.google.com)
2. Crea un nuevo proyecto
3. Copia el contenido de `code.gs` al editor
4. Crea un archivo HTML llamado `index` y copia el contenido de `index.html`
5. Crea un archivo `appsscript.json` y copia su contenido
6. **IMPORTANTE**: Habilita Drive API avanzada:
   - Ve a **Servicios** (+ junto a Servicios en la barra lateral)
   - Busca "Drive API"
   - Selecciona versión v2
   - Haz clic en "Agregar"
7. Guarda y despliega como Web App

### Opción 2: Proyecto Existente
1. Abre tu proyecto en Google Apps Script
2. Reemplaza el código existente con los nuevos archivos
3. Asegúrate de que `appsscript.json` tenga los permisos correctos
4. **IMPORTANTE**: Habilita Drive API avanzada:
   - Ve a **Servicios** (+ junto a Servicios en la barra lateral)
   - Busca "Drive API"
   - Selecciona versión v2
   - Haz clic en "Agregar"
5. Vuelve a desplegar la aplicación

---

## 🔧 Configuración

### Archivo `appsscript.json`

```json
{
  "timeZone": "America/Montevideo",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/script.external_request"
  ],
  "webapp": {
    "access": "ANYONE",
    "executeAs": "USER_DEPLOYING"
  }
}
```

### Parámetros Configurables en `code.gs`

```javascript
const CONFIG = {
  MAX_RETRIES: 3,           // Número máximo de reintentos
  RETRY_DELAY: 1000,        // Delay inicial en ms
  TIMEOUT_LIMIT: 300000,    // Límite de timeout (5 min)
  LOG_ENABLED: true         // Activar/desactivar logging
};
```

---

## 📖 Uso del Sistema

### 1. Preparar el Archivo
- El archivo debe ser Excel (.xlsx o .xls)
- Debe contener las columnas esperadas de la ficha de inscripción
- Mínimo 2 filas (encabezados + 1 registro)

### 2. Subir y Procesar
1. Abre la aplicación web
2. Arrastra o selecciona el archivo Excel
3. El sistema procesará automáticamente:
   - Validará el archivo
   - Extraerá los datos
   - Mostrará un resumen de validaciones

### 3. Verificar Datos
- Revisa la tabla de registros extraídos
- Verifica las validaciones (CI, Fecha Nac, Nivel, Teléfono)
- Expande filas para ver detalles completos

### 4. Generar Archivos
- Haz clic en "Confirmar y Generar Archivos"
- El sistema creará:
  - **Eventifica_Export_YYYY-MM-DD_HHMMSS.xlsx**
  - **AlumnosYFamilias_Export_YYYY-MM-DD_HHMMSS.xlsx**
- Descarga los archivos generados

---

## 🐛 Troubleshooting

### Ver Logs de Ejecución

1. En Google Apps Script, ve a **Ejecuciones**
2. Selecciona la ejecución más reciente
3. Revisa los logs detallados:
   - `[INFO]` - Operaciones exitosas
   - `[ERROR]` - Errores con detalles completos

### Errores Comunes

#### "Datos del archivo vacíos o inválidos"
- **Causa**: Archivo corrupto o formato incorrecto
- **Solución**: Verifica que sea un Excel válido

#### "El archivo no contiene suficientes datos"
- **Causa**: Solo hay encabezados, sin registros
- **Solución**: Agrega al menos 1 fila de datos

#### "Operación falló después de 3 intentos"
- **Causa**: Problema persistente con Google APIs
- **Solución**:
  1. Espera 5 minutos
  2. Verifica cuotas de Google en la consola
  3. Intenta con menos registros

#### "Permisos insuficientes"
- **Causa**: No has autorizado los permisos necesarios
- **Solución**: Ejecuta `checkPermissions()` y autoriza

#### "Drive is not defined" o error con Drive API
- **Causa**: Drive API avanzada no está habilitada
- **Solución**:
  1. En el editor de Apps Script, ve a la barra lateral izquierda
  2. Haz clic en el **+** junto a "Servicios"
  3. Busca "Drive API"
  4. Selecciona versión **v2**
  5. Haz clic en "Agregar"
  6. Guarda y vuelve a ejecutar

### Función de Diagnóstico

Puedes ejecutar esta función desde el editor para verificar el estado del sistema:

```javascript
function diagnosticoSistema() {
  Logger.log('=== DIAGNÓSTICO DEL SISTEMA ===');

  // Verificar permisos
  var permisos = checkPermissions();
  Logger.log('Permisos: ' + (permisos.success ? 'OK' : 'ERROR'));

  // Limpiar archivos temporales
  cleanupOldTempFiles();

  // Verificar configuración
  Logger.log('Configuración: ' + JSON.stringify(CONFIG));

  Logger.log('=== FIN DIAGNÓSTICO ===');
}
```

---

## 📊 Formatos de Salida

### Archivo EVENTIFICA
- **Formato**: template_estudiantes_padres
- **Campos**: 37 columnas
- **Incluye**: Datos de estudiante, padre y madre
- **Uso**: Sistema Eventifica de gestión escolar

### Archivo AlumnosYFamilias
- **Formato**: Plantilla_Importar_AlumnosYFamilias
- **Campos**: ~180 columnas
- **Incluye**: Datos completos de familia y alumno
- **Uso**: Sistema SIGED de gestión educativa

---

## 🔄 Historial de Versiones

### Versión 2.0 (Actual)
- ✅ Sistema de reintentos automáticos
- ✅ Manejo robusto de errores
- ✅ Logging detallado
- ✅ Batch writing para mejor rendimiento
- ✅ Limpieza automática de archivos temporales
- ✅ Validaciones exhaustivas
- ✅ Mejor gestión de permisos

### Versión 1.0
- Funcionalidad básica de conversión
- Generación de ambos formatos
- Interfaz web simple

---

## 📞 Soporte

### Revisar Logs
```javascript
// En Google Apps Script
Ver > Registros
```

### Archivos Importantes
- `code.gs` - Lógica principal del sistema
- `index.html` - Interfaz de usuario
- `appsscript.json` - Configuración y permisos

### Reportar Problemas
Si encuentras un error:
1. Copia los logs de la ejecución
2. Describe los pasos para reproducirlo
3. Incluye el mensaje de error completo

---

## 🎯 Mejores Prácticas

### Al Usar el Sistema
1. **Siempre verifica los datos antes de generar**
2. **Descarga los archivos inmediatamente** (se guardan en tu Google Drive pero pueden acumularse)
3. **Revisa los logs si algo falla**
4. **Mantén copias de seguridad de tus archivos originales**

### Mantenimiento
- Los archivos temporales se limpian automáticamente después de 1 hora
- Revisa periódicamente tu Google Drive por archivos de exportación antiguos
- Actualiza los permisos si cambias de cuenta de Google

---

## 📄 Licencia

Sistema desarrollado para el Colegio y Liceo Hans Christian Andersen.

**Desarrollado por**: Física Simple - Herramientas Educativas © 2024

---

## 🚀 Próximas Mejoras Sugeridas

- [ ] Exportación directa a PDF
- [ ] Validación de datos más avanzada (formato de email, CI válida)
- [ ] Historial de exportaciones
- [ ] Filtrado y búsqueda de registros
- [ ] Edición de datos antes de exportar
- [ ] Importación desde Google Forms directamente
