# XML de facturas que llegan por correo

Cada 10 minutos, este script revisa el Gmail donde está instalado. Busca los correos con adjuntos `.xml` o `.zip` y deja los XML de facturas y notas de crédito o débito en Firebase. La página los lee sola y los guarda en el **Lector de XML**. El menú muestra el aviso **"XML nuevos llegados por correo"** con la cantidad, y también aparece en el Inicio.

```
Correo del proveedor ─► Script de Gmail (cada 10 min) ─► Firebase: xmlBuzon ─► Página: lee, guarda en el Lector de XML y avisa
                              │
                              └─► copia del XML en Drive, carpeta "Facturas XML"
```

- Instálalo **solo en el Gmail principal**. Si las facturas llegan copiadas a otras bandejas, no hace falta instalarlo en ellas.
- Si el mismo XML llega dos veces, en dos correos o reenviado, entra una sola vez.
- Se descarta la constancia de SUNAT (CDR) que viene en algunos .zip. Solo entran facturas, NC y ND.
- Los correos ya revisados quedan con la etiqueta **XML-procesado** en Gmail.

## Instalación (una sola vez, unos 10 minutos)

### 1. Crear el script
1. Entra a <https://script.google.com> con el **Gmail principal** y pulsa **Nuevo proyecto**. Ponle de nombre "XML de facturas".
2. Ve a **Configuración del proyecto** (el engranaje) y marca **"Mostrar el archivo de manifiesto appsscript.json en el editor"**.
3. Vuelve al **Editor**:
   - En `Código.gs`, borra todo y pega el contenido de `Codigo.gs`.
   - En `appsscript.json`, borra todo y pega el contenido de `appsscript.json`.
   - Guarda con Ctrl + S.

### 2. Dar permiso a ese Gmail en Firebase
Si ese Gmail es el mismo con el que se creó el proyecto de Firebase, salta este paso.

Si no, en la [consola de Firebase](https://console.firebase.google.com) ve a **Configuración del proyecto → Usuarios y permisos → Agregar miembro**, escribe ese Gmail y elige el rol **Editor**.

> Si prefieres el permiso mínimo, en Google Cloud → IAM dale solo dos roles: **Usuario de Cloud Datastore** y **Consumidor de Service Usage**.

### 3. Agregar la regla de Firestore para la bandeja
En Firebase → **Firestore Database → Reglas**, agrega este bloque dentro de `match /databases/{database}/documents { ... }` y pulsa **Publicar**:

```
    // Bandeja de XML que llegan por correo (los escribe el script de Gmail)
    match /xmlBuzon/{id} {
      allow read: if xbPuede('verXml');
      allow update: if xbPuede('xml');
    }
    function xbUsr() { return get(/databases/$(database)/documents/usuarios/$(request.auth.uid)).data; }
    function xbPuede(p) {
      return request.auth != null && xbUsr().activo == true
        && (xbUsr().area == 'ADMIN' || xbUsr().get('permisos', {}).get(p, false) == true);
    }
```

El script no necesita una regla propia porque escribe con el permiso de tu cuenta de Google, el del paso 2. Solo escribe en `xmlBuzon`.

### 4. Encenderlo
1. En el editor, elige la función **`instalar`** en la lista de arriba y pulsa **Ejecutar**.
2. Google pedirá permisos. Como el script es tuyo y no está publicado, saldrá **"Google no verificó esta app"**: pulsa **Configuración avanzada → Ir a XML de facturas** y acepta.
3. Listo. La primera vuelta revisa los últimos 30 días. Después revisa cada 10 minutos.

Para comprobarlo, en el editor ve a **Ejecuciones** y debe aparecer algo como `5 correo(s) revisado(s) · 7 XML nuevo(s)`. Luego abre la página: en **Lector de XML** aparece la sección **"Llegados por correo"**.

## Qué hace la página
- Cuando entra a la página un usuario con el permiso **"Subir XML y eliminar documentos guardados"**, se leen los XML nuevos y se guardan en el Lector de XML, igual que si los hubiera subido a mano. Al guardarlos queda registrado que llegaron por correo, quién los envió y el asunto.
- Los usuarios con permiso para **ver** el Lector de XML ven el contador en el menú y la lista "Llegados por correo". Al hacer clic en una fila se abre el detalle de productos.
- Si un XML no se puede leer, queda marcado como **"No se pudo leer"**, con el motivo y un botón **Reintentar**.
- Si el XML es de otra empresa (otro RUC de cliente), se marca en rojo.

## Si algo falla
| Mensaje | Qué hacer |
|---|---|
| En Ejecuciones: `Firestore respondió 403 … PERMISSION_DENIED` | Falta el paso 2: el Gmail no tiene permiso en el proyecto de Firebase. |
| En Ejecuciones: `Cloud Firestore API has not been used in project …` | En el script: **Configuración del proyecto → Proyecto de Google Cloud → Cambiar proyecto** y escribe el número de proyecto de Firebase (**20439347200**). Si te lo pide, configura la "pantalla de consentimiento" de ese proyecto como **Interno** o **Externo / en prueba**, agregando tu Gmail. |
| En la página: "No se pudo leer la bandeja de XML del correo" | Falta el paso 3: la regla de `xmlBuzon`. |
| No aparece ningún XML | Revisa que el correo tenga el `.xml` o `.zip` **adjunto**. Si solo trae un enlace para descargarlo, el script no lo ve. |

Para apagarlo, ejecuta la función **`desinstalar`**.

## Configuración (arriba de `Codigo.gs`)
- `DIAS_ATRAS`: cuántos días hacia atrás mira (30).
- `CARPETA_RESPALDO`: carpeta de Drive para la copia de cada XML. Si la dejas vacía (`''`), no guarda copia.
- `MAX_CORREOS`: correos por vuelta (40).
