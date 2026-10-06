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
    const hilos = GmailApp.search(`has:attachment {filename:xml filename:zip} newer_than:${DIAS_ATRAS}d -label:${ETIQUETA}`, 0, MAX_CORREOS);
    let enviados = 0, repetidos = 0;
    hilos.forEach((hilo) => {
      hilo.getMessages().forEach((m) => {
        xmlsDe(m).forEach((x) => {
          if (enviarAlBuzon(x, m)) enviados += 1; else repetidos += 1;
        });
      });
      hilo.addLabel(etiqueta); // solo si todo salió bien; si algo falla, se reintenta en la próxima vuelta
    });
    if (hilos.length) console.log(`${hilos.length} correo(s) revisado(s) · ${enviados} XML nuevo(s) · ${repetidos} ya estaban`);
  } finally {
    candado.releaseLock();
  }
}

/** Los XML de comprobantes que trae un correo, sueltos o dentro de un .zip. */
function xmlsDe(m) {
  const out = [];
  m.getAttachments({ includeInlineImages: false }).forEach((a) => {
    const nombre = a.getName() || '';
    if (/\.xml$/i.test(nombre)) out.push({ nombre, blob: a.copyBlob() });
    else if (/\.zip$/i.test(nombre)) {
      try {
        const zip = a.copyBlob().setContentType('application/zip');
        Utilities.unzip(zip).forEach((b) => { if (/\.xml$/i.test(b.getName())) out.push({ nombre: `${nombre} → ${b.getName()}`, blob: b }); });
      } catch (e) { console.warn(`No se pudo abrir ${nombre}: ${e}`); }
    }
  });
  return out
    .map((x) => ({ nombre: x.nombre, texto: textoDe(x.blob) }))
    // solo facturas, NC y ND; se descarta la constancia de SUNAT (CDR, ApplicationResponse) y otros XML
    .filter((x) => /<(\w+:)?(Invoice|CreditNote|DebitNote)[\s>]/.test(x.texto.slice(0, 3000)));
}

/** Texto del XML respetando la codificación que declara (muchos vienen en ISO-8859-1). */
function textoDe(blob) {
  const bytes = blob.getBytes();
  const cab = Utilities.newBlob(bytes.slice(0, 200)).getDataAsString('ISO-8859-1');
  const m = cab.match(/encoding=["']([\w-]+)["']/i);
  let t;
  try { t = Utilities.newBlob(bytes).getDataAsString(m ? m[1] : 'UTF-8'); } catch (e) { t = Utilities.newBlob(bytes).getDataAsString('UTF-8'); }
  return t.replace(/^﻿/, '');
}

/** Deja el XML en xmlBuzon. El id es la huella del contenido: el mismo XML nunca entra dos veces. */
function enviarAlBuzon(x, m) {
  if (x.texto.length > 900000) { console.warn(`${x.nombre}: demasiado grande, se omite`); return false; }
  const huella = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, x.texto, Utilities.Charset.UTF_8)
    .map((b) => ((b + 256) % 256).toString(16).padStart(2, '0')).join('').slice(0, 40);
  if (CARPETA_RESPALDO) respaldar(x, huella);
  const campos = {
    xml: x.texto,
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

function respaldar(x, huella) {
  try {
    const it = DriveApp.getFoldersByName(CARPETA_RESPALDO);
    const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_RESPALDO);
    const nombre = `${huella.slice(0, 8)}_${x.nombre.split('→').pop().trim()}`;
    if (!carpeta.getFilesByName(nombre).hasNext()) carpeta.createFile(nombre, x.texto, 'application/xml');
  } catch (e) { console.warn(`No se pudo guardar la copia en Drive: ${e}`); }
}
