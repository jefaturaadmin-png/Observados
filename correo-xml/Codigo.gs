/**
 * Farmacia Socorro · XML de facturas que llegan por correo
 *
 * Revisa el Gmail donde está instalado, busca los correos con .xml o .zip adjuntos,
 * saca los XML de facturas y notas de crédito/débito (UBL 2.1 de SUNAT) y los deja en
 * Firestore (colección xmlBuzon). La página los lee, los guarda en el Lector de XML
 * y avisa en el menú. Instrucciones en LEEME.md.
 */

// ===================== CONFIGURACIÓN =====================
const PROYECTO = 'socorro-19056';          // projectId de Firebase (el mismo de index.html)
const ETIQUETA = 'XML-procesado';          // etiqueta que se pone a los correos ya revisados
const ETIQUETA_REVISAR = 'XML-sin-leer';   // correos con adjuntos que no se pudieron abrir (7z, ZIP dañado, RAR muy grande o sin XML de facturas)
const DIAS_ATRAS = 30;                     // en cada vuelta mira los correos de los últimos N días
const CARPETA_RESPALDO = 'Facturas XML';   // carpeta de Drive para guardar copia de cada XML ('' = no guardar)
const MAX_CORREOS = 40;                    // correos (hilos) por vuelta, para no pasar el límite de 6 minutos
// =========================================================

/** Ejecútalo UNA vez: crea el disparador que revisa el correo cada 10 minutos. */
function instalar() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'revisarCorreo')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('revisarCorreo').timeBased().everyMinutes(10).create();
  revisarCorreo();
}

/** Quita el disparador (deja de revisar el correo). */
function desinstalar() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'revisarCorreo')
    .forEach((t) => ScriptApp.deleteTrigger(t));
}

function revisarCorreo() {
  const candado = LockService.getScriptLock();
  if (!candado.tryLock(1000)) return; // ya hay otra vuelta corriendo
  try {
    const etiqueta = GmailApp.getUserLabelByName(ETIQUETA) || GmailApp.createLabel(ETIQUETA);
    const resto = `newer_than:${DIAS_ATRAS}d -label:${ETIQUETA} -label:${ETIQUETA_REVISAR}`;
    const conAdjunto = GmailApp.search(`has:attachment {filename:xml filename:zip filename:rar filename:7z filename:eml} ${resto}`, 0, MAX_CORREOS);
    // Correos que no traen el XML adjunto sino un botón o enlace para descargarlo ("Ver XML", "Descargar XML")
    const conEnlace = GmailApp.search(`{"Ver XML" "Descargar XML" "Descarga XML" "Descargar el XML" "Ver el XML"} ${resto}`, 0, MAX_CORREOS);
    const vistos = new Set(), hilos = conAdjunto.concat(conEnlace).filter((h) => !vistos.has(h.getId()) && vistos.add(h.getId())).slice(0, MAX_CORREOS);
    let enviados = 0, repetidos = 0;
    const inicio = Date.now();
    for (const hilo of hilos) {
      if (Date.now() - inicio > 4.5 * 60 * 1000) { console.log('Se acaba el tiempo de esta vuelta: lo que falta sigue en la próxima.'); break; }
      let xmls = 0;
      const raros = [];
      hilo.getMessages().forEach((m) => {
        const r = xmlsDe(m);
        raros.push(...r.raros);
        r.xmls.forEach((x) => { xmls += 1; if (enviarAlBuzon(x, m)) enviados += 1; else repetidos += 1; });
        // Los RAR no se pueden abrir desde Google: se mandan tal cual y la página los abre
        r.rars.forEach((x) => { if (x.blob.getBytes().length > 650000) { raros.push(`${x.nombre} (RAR de más de 650 KB)`); return; } xmls += 1; if (enviarAlBuzon(x, m)) enviados += 1; else repetidos += 1; });
      });
      if (!xmls && raros.length) { // nada que leer: se marca para revisarlo a mano (no queda como procesado)
        hilo.addLabel(GmailApp.getUserLabelByName(ETIQUETA_REVISAR) || GmailApp.createLabel(ETIQUETA_REVISAR));
        console.warn(`"${hilo.getFirstMessageSubject()}": no se pudo leer ${raros.join(' · ')}`);
        continue;
      }
      hilo.addLabel(etiqueta); // solo si todo salió bien; si algo falla, se reintenta en la próxima vuelta
    }
    if (hilos.length) console.log(`${hilos.length} correo(s) revisado(s) · ${enviados} XML nuevo(s) · ${repetidos} ya estaban`);
    // El Outlook conectado (si lo hay) se revisa en la misma vuelta: ver Outlook.gs
    if (typeof revisarOutlook === 'function') { try { revisarOutlook(inicio); } catch (e) { console.error(`Outlook: ${e.message || e}`); } }
  } finally {
    candado.releaseLock();
  }
}

/** Los XML de comprobantes que trae un correo: sueltos, dentro de un .zip o de un .zip dentro de otro (hasta 3 niveles).
 *  También abre los correos reenviados como adjunto (.eml, por ejemplo al reenviar desde Outlook) y saca sus XML.
 *  raros = adjuntos que no se pudieron abrir (RAR y 7z no se pueden abrir desde Google: hay que descomprimirlos en la PC). */
function xmlsDe(m) {
  const r = xmlsDeAdjuntos(m.getAttachments({ includeInlineImages: false }).map((a) => ({ blob: a.copyBlob(), nombre: a.getName() || '' })));
  if (r.xmls.length || r.rars.length) return r;
  // Sin XML adjunto: se buscan en el texto del correo los enlaces para descargarlo y se descargan
  const bajados = xmlsDeEnlaces(m.getBody() || '');
  if (!bajados.adjuntos.length) { r.raros.push(...bajados.raros); return r; }
  const r2 = xmlsDeAdjuntos(bajados.adjuntos);
  r2.raros.push(...bajados.raros);
  return r2;
}

/** Enlaces del correo que dicen XML ("Ver XML", "Descargar XML", o que terminan en .xml): se descargan (máximo 5). */
function xmlsDeEnlaces(html) {
  const adjuntos = [], raros = [], vistos = new Set();
  const limpiar = (t) => t.replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
  const re = /<a\b[^>]*?href\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let m;
  while ((m = re.exec(html)) && vistos.size < 5) {
    const url = m[1].replace(/&amp;/g, '&').trim(), texto = limpiar(m[2]);
    if (!/^https?:\/\//i.test(url) || vistos.has(url)) continue;
    if (!/\bxml\b/i.test(texto) && !/\.xml(\?|#|$)/i.test(url)) continue;
    vistos.add(url);
    try {
      const r = UrlFetchApp.fetch(url, { followRedirects: true, muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
      if (r.getResponseCode() !== 200) { raros.push(`enlace "${texto}" (respondió ${r.getResponseCode()})`); continue; }
      const blob = r.getBlob(), bytes = blob.getBytes(), cab = Utilities.newBlob(bytes.slice(0, 400)).getDataAsString('ISO-8859-1');
      const disp = String((r.getHeaders() || {})['Content-Disposition'] || (r.getHeaders() || {})['content-disposition'] || '');
      const nom = ((disp.match(/filename\*?=(?:UTF-8'')?"?([^";]+)"?/i) || [])[1] || url.split('?')[0].split('/').pop() || 'enlace').trim();
      if ((bytes[0] & 255) === 0x50 && (bytes[1] & 255) === 0x4B) adjuntos.push({ blob: Utilities.newBlob(bytes, 'application/zip', /\.zip$/i.test(nom) ? nom : nom + '.zip'), nombre: /\.zip$/i.test(nom) ? nom : nom + '.zip' });
      else if (/<\?xml|<(\w+:)?(Invoice|CreditNote|DebitNote)[\s>]/.test(cab) || (bytes[0] & 255) === 0xFF || (bytes[0] & 255) === 0xFE) adjuntos.push({ blob: Utilities.newBlob(bytes, 'application/xml', /\.xml$/i.test(nom) ? nom : nom + '.xml'), nombre: `enlace → ${/\.xml$/i.test(nom) ? nom : nom + '.xml'}` });
      else raros.push(`enlace "${texto}" (abre una página, no descarga el XML)`);
    } catch (e) { raros.push(`enlace "${texto}" (no se pudo descargar)`); }
  }
  return { adjuntos, raros };
}

/** Lo mismo, a partir de una lista de adjuntos { blob, nombre } (sirve para Gmail y para Outlook). */
function xmlsDeAdjuntos(adjuntos) {
  const out = [], raros = [], rars = [];
  const abrir = (blob, nombre, nivel) => {
    if (/\.xml$/i.test(nombre)) { out.push({ nombre, blob }); return; }
    if ((/\.eml$/i.test(nombre) || /message\/rfc822/i.test(blob.getContentType() || '')) && nivel <= 4) { // correo reenviado como adjunto
      try { adjuntosEml(blob.getBytes()).forEach((b) => abrir(b, `${nombre} → ${b.getName()}`, nivel + 1)); } catch (e) { raros.push(`${nombre} (correo adjunto que no se pudo abrir)`); console.warn(`No se pudo abrir ${nombre}: ${e}`); }
      return;
    }
    if (/\.rar$/i.test(nombre)) { rars.push({ nombre, blob, rar: true }); return; }
    if (/\.7z$/i.test(nombre)) { raros.push(`${nombre} (7z)`); return; }
    if (!/\.zip$/i.test(nombre) || nivel > 3) return;
    try {
      Utilities.unzip(blob.setContentType('application/zip')).forEach((b) => abrir(b, `${nombre} → ${b.getName()}`, nivel + 1));
    } catch (e) { raros.push(`${nombre} (ZIP que no se pudo abrir)`); console.warn(`No se pudo abrir ${nombre}: ${e}`); }
  };
  adjuntos.forEach((a) => abrir(a.blob, a.nombre, 1));
  const xmls = out
    .map((x) => ({ nombre: x.nombre, texto: textoDe(x.blob) }))
    // solo facturas, NC y ND; se descarta la constancia de SUNAT (CDR, ApplicationResponse) y otros XML
    .filter((x) => /<(\w+:)?(Invoice|CreditNote|DebitNote)[\s>]/.test(x.texto.slice(0, 3000)));
  if (out.length && !xmls.length && !rars.length) raros.push('XML que no son facturas ni notas (guías, constancias u otros)');
  return { xmls, raros, rars };
}

/** Archivos adjuntos de un correo guardado como .eml (formato MIME), incluidos los de correos adjuntos dentro de él. */
function adjuntosEml(bytes) {
  const latin = (bs) => Utilities.newBlob(bs).getDataAsString('ISO-8859-1'); // 1 carácter = 1 byte, sin perder nada
  const aBytes = (t) => { const b = new Array(t.length); for (let i = 0; i < t.length; i += 1) { const c = t.charCodeAt(i) & 255; b[i] = c > 127 ? c - 256 : c; } return b; };
  const palabras = (v) => String(v || '').replace(/=\?([^?]+)\?([BQ])\?([^?]*)\?=/gi, (x, cs, enc, txt) => { // =?UTF-8?B?...?=
    try {
      const bs = enc.toUpperCase() === 'B' ? Utilities.base64Decode(txt) : aBytes(txt.replace(/_/g, ' ').replace(/=([0-9A-F]{2})/gi, (y, h) => String.fromCharCode(parseInt(h, 16))));
      return Utilities.newBlob(bs).getDataAsString(cs);
    } catch (e) { return txt; }
  }).replace(/\?=\s+=\?/g, '?==?');
  const param = (h, n) => {
    const ext = h.match(new RegExp(`${n}\\*(?:0\\*?)?=(?:([^']*)'[^']*')?"?([^";]+)"?`, 'i')); // filename*=UTF-8''nombre.xml
    if (ext) { try { return decodeURIComponent(ext[2]); } catch (e) { return ext[2]; } }
    const m = h.match(new RegExp(`(?:^|;)\\s*${n}\\s*=\\s*("([^"]*)"|[^;\\s]+)`, 'i'));
    return m ? palabras(m[2] !== undefined ? m[2] : m[1]) : '';
  };
  const salida = [];
  const parte = (t, nivel) => {
    const corte = t.search(/\r?\n\r?\n/);
    const cab = (corte < 0 ? t : t.slice(0, corte)).replace(/\r?\n[ \t]+/g, ' '), cuerpo = corte < 0 ? '' : t.slice(corte).replace(/^\r?\n\r?\n/, '');
    const h = (n) => { const m = cab.match(new RegExp(`^${n}:\\s*(.*)$`, 'im')); return m ? m[1].trim() : ''; };
    const tipo = h('Content-Type') || 'text/plain', cod = h('Content-Transfer-Encoding').toLowerCase(), disp = h('Content-Disposition');
    const decodificar = () => {
      if (cod === 'base64') return Utilities.base64Decode(cuerpo.replace(/[^A-Za-z0-9+/=]/g, ''));
      if (cod === 'quoted-printable') return aBytes(cuerpo.replace(/=\r?\n/g, '').replace(/=([0-9A-F]{2})/gi, (x, hx) => String.fromCharCode(parseInt(hx, 16))));
      return aBytes(cuerpo);
    };
    if (/^multipart\//i.test(tipo)) {
      const lim = param(tipo, 'boundary');
      if (!lim) return;
      cuerpo.split(new RegExp(`(?:^|\\r?\\n)--${lim.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?:--)?[ \\t]*(?=\\r?\\n|$)`)).slice(1).forEach((p) => { const q = p.replace(/^\r?\n/, ''); if (q.trim()) parte(q, nivel + 1); });
      return;
    }
    if (/^message\/rfc822/i.test(tipo) && nivel < 8) { parte(cod === 'base64' || cod === 'quoted-printable' ? latin(decodificar()) : cuerpo, nivel + 1); return; }
    const nombre = param(disp, 'filename') || param(tipo, 'name');
    if (!nombre) return; // texto del correo, no es un adjunto
    salida.push(Utilities.newBlob(decodificar(), tipo.split(';')[0].trim(), nombre));
  };
  parte(latin(bytes), 0);
  return salida;
}

/** Ejecútalo UNA vez si unos correos reenviados (.eml) quedaron como procesados sin leer sus XML: los vuelve a revisar. */
function revisarDeNuevo() {
  const etiqueta = GmailApp.getUserLabelByName(ETIQUETA);
  if (!etiqueta) return;
  const hilos = GmailApp.search(`filename:eml newer_than:${DIAS_ATRAS}d label:${ETIQUETA}`, 0, 200);
  hilos.forEach((h) => h.removeLabel(etiqueta));
  console.log(`${hilos.length} correo(s) con correos adjuntos se volverán a revisar`);
  revisarCorreo();
}

/** Texto del XML respetando la codificación que declara (muchos vienen en ISO-8859-1). */
function textoDe(blob) {
  const bytes = blob.getBytes();
  if ((bytes[0] & 255) === 0xFF && (bytes[1] & 255) === 0xFE) return blob.getDataAsString('UTF-16LE').replace(/^\uFEFF/, ''); // XML guardado en UTF-16
  if ((bytes[0] & 255) === 0xFE && (bytes[1] & 255) === 0xFF) return blob.getDataAsString('UTF-16BE').replace(/^\uFEFF/, '');
  const cab = Utilities.newBlob(bytes.slice(0, 200)).getDataAsString('ISO-8859-1');
  const m = cab.match(/encoding=["']([\w-]+)["']/i);
  let t;
  try { t = Utilities.newBlob(bytes).getDataAsString(m ? m[1] : 'UTF-8'); } catch (e) { t = Utilities.newBlob(bytes).getDataAsString('UTF-8'); }
  return t.replace(/^﻿/, '');
}

/** Deja el XML en xmlBuzon. El id es la huella del contenido: el mismo XML nunca entra dos veces. */
function enviarAlBuzon(x, m) {
  const rar = x.rar ? Utilities.base64Encode(x.blob.getBytes()) : '';
  if (!x.rar && x.texto.length > 900000) { console.warn(`${x.nombre}: demasiado grande, se omite`); return false; }
  const huella = (x.rar ? 'r' : '') + Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, x.rar ? rar : x.texto, Utilities.Charset.UTF_8)
    .map((b) => ((b + 256) % 256).toString(16).padStart(2, '0')).join('').slice(0, 40);
  if (CARPETA_RESPALDO) respaldar(x, huella);
  const campos = {
    ...(x.rar ? { rar } : { xml: x.texto }),
    archivo: x.nombre.slice(0, 200),
    de: m.getFrom().slice(0, 200),
    asunto: m.getSubject().slice(0, 200),
    recibido: m.getDate().toISOString(),
    estado: 'nuevo'
  };
  const url = `https://firestore.googleapis.com/v1/projects/${PROYECTO}/databases/(default)/documents/xmlBuzon?documentId=x${huella}`;
  const r = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Goog-User-Project': PROYECTO },
    payload: JSON.stringify({ fields: Object.fromEntries(Object.entries(campos).map(([k, v]) => [k, { stringValue: v }])) }),
    muteHttpExceptions: true
  });
  const cod = r.getResponseCode();
  if (cod === 200) return true;
  if (cod === 409) return false; // ya estaba en el buzón
  throw new Error(`Firestore respondió ${cod}: ${r.getContentText().slice(0, 500)}`);
}

let carpetaRespaldo = null;
function respaldar(x, huella) {
  try {
    if (!carpetaRespaldo) { const it = DriveApp.getFoldersByName(CARPETA_RESPALDO); carpetaRespaldo = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_RESPALDO); }
    const carpeta = carpetaRespaldo;
    const nombre = `${huella.slice(0, 8)}_${x.nombre.split('→').pop().trim()}`;
    if (!carpeta.getFilesByName(nombre).hasNext()) { if (x.rar) carpeta.createFile(x.blob.copyBlob().setName(nombre)); else carpeta.createFile(nombre, x.texto, 'application/xml'); }
  } catch (e) { console.warn(`No se pudo guardar la copia en Drive: ${e}`); }
}
