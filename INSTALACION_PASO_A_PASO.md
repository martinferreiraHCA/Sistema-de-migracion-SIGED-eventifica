# 📝 Guía de Instalación Paso a Paso

## ⚠️ IMPORTANTE: Habilitar Drive API es OBLIGATORIO

Este sistema requiere Drive API avanzada para funcionar correctamente. **No omitas este paso**.

---

## 🎯 Pasos para Aplicar los Cambios

### Paso 1: Abrir tu Proyecto en Google Apps Script

1. Ve a tu aplicación web actual
2. En la URL, busca algo como: `https://script.google.com/home/projects/XXXX/edit`
3. O ve directamente a [Google Apps Script](https://script.google.com) y abre el proyecto

---

### Paso 2: Habilitar Drive API Avanzada (CRÍTICO)

**⚠️ ESTE ES EL PASO MÁS IMPORTANTE ⚠️**

1. En el editor de Google Apps Script, mira la barra lateral **izquierda**
2. Verás una sección llamada **"Servicios"**
3. Haz clic en el botón **"+"** (Agregar un servicio)

   ```
   📁 Archivos
   📋 Bibliotecas
   ⚙️ Servicios  [+]  ← HAZ CLIC AQUÍ
   ```

4. En el diálogo que se abre:
   - Busca: **"Drive API"** o **"Google Drive API"**
   - Selecciona: **"Drive API"**
   - Versión: **v2** (importante, NO v3)
   - Identificador: Déjalo como "Drive"

5. Haz clic en **"Agregar"**

6. Deberías ver ahora en Servicios:
   ```
   ⚙️ Servicios
      └─ Drive (v2)
   ```

**Si no ves la opción de Servicios:**
- Ve a **Configuración del proyecto** (ícono de engranaje ⚙️)
- Activa "Mostrar archivos de manifiesto de proyecto"
- Edita `appsscript.json` manualmente (ver Paso 4)

---

### Paso 3: Actualizar el Código (code.gs)

1. En el editor, abre el archivo `code.gs`
2. **Selecciona TODO el contenido** (Ctrl+A o Cmd+A)
3. **Elimina todo** (Delete)
4. Abre el archivo `code.gs` de este repositorio
5. **Copia TODO el contenido**
6. **Pega** en el editor de Google Apps Script
7. Haz clic en **💾 Guardar** (Ctrl+S)

---

### Paso 4: Actualizar appsscript.json

1. En el editor, ve a **Configuración del proyecto** (ícono de engranaje ⚙️)
2. Marca la casilla: **"Mostrar archivos de manifiesto de proyecto"**
3. En la lista de archivos, haz clic en `appsscript.json`
4. **Selecciona TODO el contenido**
5. **Elimina todo**
6. Copia el siguiente contenido:

```json
{
  "timeZone": "America/Montevideo",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Drive",
        "version": "v2",
        "serviceId": "drive"
      }
    ]
  },
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

7. Haz clic en **💾 Guardar**

---

### Paso 5: Verificar Servicios

Después de guardar `appsscript.json`, verifica que Drive API esté habilitada:

1. Mira la barra lateral izquierda
2. En **Servicios** deberías ver:
   ```
   ⚙️ Servicios
      └─ Drive (v2) ✓
   ```

Si NO aparece:
- Vuelve al **Paso 2** y agrega el servicio manualmente
- O verifica que `appsscript.json` esté guardado correctamente

---

### Paso 6: Autorizar Nuevos Permisos

1. En el editor, selecciona la función: **`checkPermissions`** del menú desplegable
2. Haz clic en **▶️ Ejecutar**
3. Te pedirá autorización:
   - Haz clic en **"Revisar permisos"**
   - Selecciona tu cuenta de Google
   - Verás una advertencia: **"Esta aplicación no ha sido verificada"**
   - Haz clic en **"Configuración avanzada"**
   - Haz clic en **"Ir a [nombre del proyecto] (no seguro)"**
   - Haz clic en **"Permitir"**

4. Verifica en **"Ejecuciones"** que la función se ejecutó exitosamente:
   ```
   [INFO] Verificando permisos de acceso...
   [INFO] Acceso a Drive: OK
   [INFO] Acceso a Sheets: OK
   ```

---

### Paso 7: Volver a Desplegar la Aplicación

1. Haz clic en **"Implementar"** → **"Administrar implementaciones"**
2. Haz clic en el **ícono de lápiz** ✏️ junto a la implementación activa
3. En **"Versión"**, selecciona **"Nueva versión"**
4. **Descripción** (opcional): "v2.0 - Fix Google Sheets API error"
5. Haz clic en **"Implementar"**
6. Copia la **URL de la aplicación web** (la necesitarás)

---

### Paso 8: Probar la Aplicación

1. Abre la **URL de la aplicación web** en tu navegador
2. Sube un archivo Excel de prueba
3. Verifica que:
   - El archivo se procesa correctamente
   - No aparecen errores
   - Los registros se muestran en la tabla
   - Puedes generar los archivos de salida

---

## 🔍 Verificación de Éxito

### ✅ Todo funciona correctamente si:

- [x] Drive API (v2) aparece en Servicios
- [x] `checkPermissions()` se ejecuta sin errores
- [x] Puedes subir archivos Excel sin error
- [x] Los archivos se procesan y generan correctamente
- [x] En "Ejecuciones" ves logs como:
  ```
  [INFO] Archivo convertido a Google Sheets con ID (Drive API): XXXX
  [INFO] Spreadsheet abierto exitosamente
  [INFO] Datos escritos: N registros
  ```

### ❌ Si algo falla:

**Error: "Drive is not defined"**
- Vuelve al **Paso 2** y habilita Drive API
- Verifica que `appsscript.json` tenga la sección `enabledAdvancedServices`

**Error: "El servicio Hojas de cálculo falló..."**
- Verifica que Drive API esté habilitada (Paso 2)
- Verifica que hayas autorizado los nuevos permisos (Paso 6)
- Revisa los logs en "Ejecuciones" para más detalles

**Error: "Permisos insuficientes"**
- Vuelve al **Paso 6** y autoriza todos los permisos

---

## 📊 Logs de Ejecución

Para ver qué está pasando:

1. En el editor, ve a **"Ejecuciones"** (icono de reloj 🕐)
2. Haz clic en la ejecución más reciente
3. Verás logs detallados:
   ```
   [INFO] Iniciando procesamiento de archivo: ejemplo.xlsx
   [INFO] Archivo temporal creado con ID: XXXX
   [INFO] Intentando conversión de Excel a Google Sheets...
   [INFO] Archivo convertido a Google Sheets con ID (Drive API): YYYY
   [INFO] Spreadsheet abierto exitosamente
   [INFO] Datos leídos: 50 filas
   [INFO] Procesamiento completado. Registros válidos: 48
   ```

---

## 🆘 Ayuda

Si después de seguir todos los pasos aún tienes problemas:

1. **Copia los logs** de "Ejecuciones"
2. **Toma una captura** de la sección "Servicios"
3. **Verifica** que `appsscript.json` tenga exactamente el contenido del Paso 4
4. **Comparte** esta información para debugging

---

## 📞 Contacto

Si necesitas ayuda adicional, proporciona:
- Logs completos de la ejecución fallida
- Captura de pantalla de "Servicios"
- Contenido de `appsscript.json`
- Descripción del error exacto

---

## ✨ Resumen Rápido

```bash
1. Habilitar Drive API v2 en Servicios         [CRÍTICO]
2. Actualizar code.gs                          [COPIAR/PEGAR]
3. Actualizar appsscript.json                  [COPIAR/PEGAR]
4. Ejecutar checkPermissions()                 [AUTORIZAR]
5. Volver a desplegar                          [NUEVA VERSIÓN]
6. Probar con archivo de prueba                [VERIFICAR]
```

**Tiempo estimado:** 10-15 minutos

¡Buena suerte! 🚀
