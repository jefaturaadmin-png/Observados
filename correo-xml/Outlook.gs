/**
 * Farmacia Socorro · XML de facturas que llegan al OUTLOOK
 *
 * Opcional. Si OUTLOOK_CLIENT_ID está lleno y ya conectaste la cuenta (outlookConectar + outlookTerminar),
 * en cada vuelta de revisarCorreo también se revisa el Outlook (cuenta personal o de empresa, Microsoft 365)
 * con Microsoft Graph. Los XML van al mismo buzón de Firebase que los del Gmail: el mismo XML nunca entra dos veces.
 * Los correos revisados quedan con la categoría "XML-procesado" (o "XML-sin-leer") en Outlook. Pasos en LEEME.md.
 */

// ===================== CONFIGURACIÓN =====================
const OUTLOOK_CLIENT_ID = '';   // "Id. de aplicación (cliente)" de la app registrada en Microsoft Entra ('' = no revisar Outlook)
// =========================================================

const MS_LOGIN = 'https://login.microsoftonline.com/common/oauth2/v2.0';
const MS_GRAPH = 'https://graph.microsoft.com/v1.0';
const MS_PERMISOS = 'offline_access Mail.ReadWrite User.Read';

/** PASO 1 · Ejecútalo y sigue lo que dice el registro: abre la página de Microsoft y escribe el código. */
function outlookConectar() {
  if (!OUTLOOK_CLIENT_ID) throw new Error('Primero pega el Id. de aplicación (cliente) en OUTLOOK_CLIENT_ID, arriba de Outlook.gs.');
  const r = UrlFetchApp.fetch(`${MS_LOGIN}/devicecode`, { method: 'post', payload: { client_id: OUTLOOK_CLIENT_ID, scope: MS_PERMISOS }, muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText());
  if (!j.device_code) throw new Error(`Microsoft respondió: ${r.getContentText().slice(0, 400)}`);
  PropertiesService.getScriptProperties().setProperty('outlookDevice', j.device_code);
  console.log(`1) Abre ${j.verification_uri}\n2) Escribe el código: ${j.user_code}\n3) Entra con la cuenta de Outlook y acepta los permisos.\n4) Vuelve aquí y ejecuta outlookTerminar (tienes ${Math.round(j.expires_in / 60)} minutos).`);
}

/** PASO 2 · Después de escribir el código en la página de Microsoft, ejecuta esta función. */
function outlookTerminar() {
  const props = PropertiesService.getScriptProperties(), dev = props.getProperty('outlookDevice');
  if (!dev) throw new Error('Primero ejecuta outlookConectar.');
  const r = UrlFetchApp.fetch(`${MS_LOGIN}/token`, { method: 'post', payload: { grant_type: 'urn:ietf:params:oauth:grant-type:device_code', client_id: OUTLOOK_CLIENT_ID, device_code: dev }, muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText());
  if (j.error === 'authorization_pending') { console.log('Todavía no terminas de iniciar sesión en la página de Microsoft. Termina y vuelve a ejecutar outlookTerminar.'); return; }
  if (!j.refresh_token) throw new Error(`No se pudo conectar: ${j.error_description || r.getContentText().slice(0, 400)}. Vuelve a ejecutar outlookConectar.`);
  props.setProperty('outlookRefresh', j.refresh_token); props.deleteProperty('outlookDevice');
  CacheService.getScriptCache().put('outlookToken', j.access_token, 3000);
  const yo = msGraph('/me?$select=mail,userPrincipalName');
  console.log(`Outlook conectado: ${yo.mail || yo.userPrincipalName}. Desde ahora se revisa junto con el Gmail cada 10 minutos.`);
}

/** Desconecta el Outlook (deja de revisarlo). */
function outlookDesconectar() {
  PropertiesService.getScriptProperties().deleteProperty('outlookRefresh');
  CacheService.getScriptCache().remove('outlookToken');
  console.log('Outlook desconectado.');
}

function msToken() {
  const cache = CacheService.getScriptCache(), t = cache.get('outlookToken');
  if (t) return t;
  const props = PropertiesService.getScriptProperties(), ref = props.getProperty('outlookRefresh');
  if (!ref) return '';
  const r = UrlFetchApp.fetch(`${MS_LOGIN}/token`, { method: 'post', payload: { grant_type: 'refresh_token', client_id: OUTLOOK_CLIENT_ID, refresh_token: ref, scope: MS_PERMISOS }, muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText());
  if (!j.access_token) throw new Error(`Se perdió la conexión con Outlook (${j.error_description || j.error || r.getResponseCode()}). Ejecuta outlookConectar y outlookTerminar otra vez.`);
  if (j.refresh_token) props.setProperty('outlookRefresh', j.refresh_token);
  cache.put('outlookToken', j.access_token, Math.max(60, (j.expires_in || 3600) - 300));
  return j.access_token;
}

function msGraph(ruta, opciones) {
  const o = opciones || {};
  const r = UrlFetchApp.fetch(ruta.startsWith('http') ? ruta : MS_GRAPH + ruta, {
    method: o.method || 'get', contentType: o.payload ? 'application/json' : undefined, payload: o.payload ? JSON.stringify(o.payload) : undefined,
    headers: { Authorization: 'Bearer ' + msToken() }, muteHttpExceptions: true
  });
  const cod = r.getResponseCode();
  if (cod >= 400) throw new Error(`Microsoft Graph respondió ${cod}: ${r.getContentText().slice(0, 300)}`);
  return o.crudo ? r : (cod === 204 ? null : JSON.parse(r.getContentText() || 'null'));
}

/** Revisa los correos del Outlook con adjuntos de los últimos DIAS_ATRAS días que aún no tienen la categoría. */
function revisarOutlook(inicio) {
  if (!OUTLOOK_CLIENT_ID || !PropertiesService.getScriptProperties().getProperty('outlookRefresh')) return;
  const desde = new Date(Date.now() - DIAS_ATRAS * 86400000).toISOString();
  let url = `/me/messages?$filter=${encodeURIComponent(`receivedDateTime ge ${desde} and hasAttachments eq true`)}&$orderby=receivedDateTime%20desc&$select=id,subject,from,receivedDateTime,categories&$top=50`;
  let revisados = 0, enviados = 0, repetidos = 0;
  while (url && revisados < MAX_CORREOS) {
    const pag = msGraph(url);
    for (const c of pag.value || []) {
      if (revisados >= MAX_CORREOS || Date.now() - inicio > 4.5 * 60 * 1000) { url = ''; break; }
      const cats = c.categories || [];
      if (cats.includes(ETIQUETA) || cats.includes(ETIQUETA_REVISAR)) continue;
      revisados += 1;
      const adj = [];
      (msGraph(`/me/messages/${c.id}/attachments`).value || []).forEach((a) => {
        const tipo = a['@odata.type'] || '';
        if (tipo.endsWith('fileAttachment') && a.contentBytes) adj.push({ blob: Utilities.newBlob(Utilities.base64Decode(a.contentBytes), a.contentType || 'application/octet-stream', a.name || ''), nombre: a.name || '' });
        else if (tipo.endsWith('itemAttachment')) { // correo adjunto: se baja en formato MIME y se abre como un .eml
          const mime = msGraph(`/me/messages/${c.id}/attachments/${a.id}/$value`, { crudo: true }).getContent();
          const nombre = (a.name || 'correo') + '.eml';
          adj.push({ blob: Utilities.newBlob(mime, 'message/rfc822', nombre), nombre });
        }
      });
      const m = { // lo que enviarAlBuzon necesita, como un mensaje de Gmail
        getFrom: () => { const e = (c.from || {}).emailAddress || {}; return `${e.name || ''} <${e.address || ''}>`; },
        getSubject: () => c.subject || '',
        getDate: () => new Date(c.receivedDateTime)
      };
      const r = xmlsDeAdjuntos(adj);
      let xmls = 0;
      r.xmls.forEach((x) => { xmls += 1; if (enviarAlBuzon(x, m)) enviados += 1; else repetidos += 1; });
      r.rars.forEach((x) => { if (x.blob.getBytes().length > 650000) { r.raros.push(`${x.nombre} (RAR de más de 650 KB)`); return; } xmls += 1; if (enviarAlBuzon(x, m)) enviados += 1; else repetidos += 1; });
      const cat = !xmls && r.raros.length ? ETIQUETA_REVISAR : ETIQUETA;
      if (cat === ETIQUETA_REVISAR) console.warn(`Outlook · "${c.subject}": no se pudo leer ${r.raros.join(' · ')}`);
      msGraph(`/me/messages/${c.id}`, { method: 'patch', payload: { categories: cats.concat([cat]) } });
    }
    if (url) url = pag['@odata.nextLink'] || '';
  }
  if (revisados) console.log(`Outlook: ${revisados} correo(s) revisado(s) · ${enviados} XML nuevo(s) · ${repetidos} ya estaban`);
}
