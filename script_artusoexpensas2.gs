/**
 * Script para la casilla CENTRAL: artusoexpensas2@gmail.com
 * - Aquí puedes decidir si solo etiquetas expensas o también las reenvías a otro sistema.
 * - Versión por defecto: SOLO etiqueta los mails de expensas (no los reenvía para evitar duplicados).
 */

// ==================== CONFIGURACIÓN ====================
const CONFIG = {
  EMAILS_ORIGEN: ['artusoexpensas2@gmail.com'],
  EMAIL_DESTINO: 'artusoexpensas2@gmail.com', // se mantiene igual, pero NO se usa para reenviar
  ETIQUETA_PROCESADO: 'ExpensaProcesada',
  ETIQUETA_DESCARTADO: 'ExpensaDescartada',
  ETIQUETA_EN_PROCESO: 'ExpensaEnProceso',
  ETIQUETA_REQUIERE_REVISION: 'REQUIERE REVISION',
  ETIQUETA_MULTIPLES_COMPROBANTES: 'MULTIPLES COMPROBANTES',
  MAX_THREADS_PER_RUN: 10,
  MAX_MESSAGES_PER_THREAD: 20,
  MAX_TOTAL_MESSAGES_PER_RUN: 40,
  DEBUG: true
};

// ==================== CONFIGURACIÓN IA ====================
// Solo cambiar PROVIDER y MODEL. La API_KEY se configura en Propiedades del script.
const AI_CONFIG = {
  PROVIDER: 'openai',       // 'openai', 'gemini', o 'anthropic'
  MODEL: 'gpt-4o-mini'      // Modelo del proveedor elegido
  // API_KEY: Se obtiene de PropertiesService (ver función getAIApiKey)
};

/**
 * Obtiene la API key del proveedor configurado desde las Propiedades del script.
 * Configurar en: Configuración del proyecto > Propiedades del script
 * 
 * Nombres de propiedades esperados:
 * - OPENAI_API_KEY (para provider 'openai')
 * - GEMINI_API_KEY (para provider 'gemini')
 * - ANTHROPIC_API_KEY (para provider 'anthropic')
 * 
 * @returns {string} La API key del proveedor
 */
function getAIApiKey() {
  const propertyNames = {
    'openai': 'OPENAI_API_KEY',
    'gemini': 'GEMINI_API_KEY',
    'anthropic': 'ANTHROPIC_API_KEY'
  };
  
  const propertyName = propertyNames[AI_CONFIG.PROVIDER];
  if (!propertyName) {
    throw new Error(`Proveedor no soportado: ${AI_CONFIG.PROVIDER}`);
  }
  
  const apiKey = PropertiesService.getScriptProperties().getProperty(propertyName);
  
  if (!apiKey) {
    throw new Error(
      `API Key no configurada para ${AI_CONFIG.PROVIDER}. ` +
      `Ir a Configuración del proyecto > Propiedades del script y agregar: ${propertyName}`
    );
  }
  
  return apiKey;
}

// ==================== CONFIGURACIÓN GOOGLE SHEET ====================
const SHEET_CONFIG = {
  SPREADSHEET_ID: '...',    // ID del Google Sheet (se obtiene de la URL)
  SHEET_NAME: 'Expensas'    // Nombre de la hoja
};

const LOCALES_SEMICOLON = ['es', 'fr', 'de', 'it', 'pt', 'nl', 'da', 'fi', 'no', 'sv', 'pl', 'ru'];

// ==================== OCR (Drive) ====================
// OCR via Google Drive -> Google Doc. Costo $0, pero puede ser lento.
const OCR_CONFIG = {
  ENABLED: true,
  OCR_LANGUAGE: 'es',
  MAX_BLOBS: 3,                 // límite por email
  MAX_BYTES: 4 * 1024 * 1024,   // 4MB por archivo (ajustable)
  MIN_INLINE_IMAGE_BYTES: 30 * 1024, // ignora logos/firma
  MAX_OCR_CHARS: 6000,          // limita tokens al pasar al LLM
  MAX_OCR_CALLS_PER_RUN: 15,
  STALE_EN_PROCESO_HOURS: 6
};

var OCR_CALLS_RUN = 0;

const DISCARD_STATE_MONTH_KEY = 'DESCARTADO_STATE_MONTH';
const DISCARD_LASTMSG_PREFIX = 'DESCARTADO_LASTMSG_';

/**
 * Determina el separador de argumentos de fórmula según el locale.
 */
function getFormulaSeparator(locale) {
  if (!locale) return ',';
  const normalized = locale.toLowerCase().split('_')[0];
  return LOCALES_SEMICOLON.includes(normalized) ? ';' : ',';
}

function getMonthStartQuery_() {
  const tz = 'America/Argentina/Buenos_Aires';
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const formatted = Utilities.formatDate(start, tz, 'yyyy/MM/dd');
  return `after:${formatted}`;
}

function getMonthKey_() {
  const tz = 'America/Argentina/Buenos_Aires';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM');
}

function canUseDriveOcr_() {
  // Requiere habilitar el servicio avanzado "Drive API" en Apps Script.
  return typeof Drive !== 'undefined' && Drive && Drive.Files;
}

function isRateLimitError_(e) {
  const msg = (e && (e.message || e.toString())) ? (e.message || e.toString()) : '';
  const m = msg.toLowerCase();
  return m.indexOf('rate limit') !== -1 ||
    m.indexOf('user rate limit') !== -1 ||
    m.indexOf('service invoked too many times') !== -1 ||
    m.indexOf('quota') !== -1 ||
    m.indexOf('429') !== -1 ||
    m.indexOf('403') !== -1;
}

function withBackoff_(fn, attempts) {
  const maxAttempts = typeof attempts === 'number' ? attempts : 6;
  let delay = 800;
  for (var i = 0; i < maxAttempts; i++) {
    try {
      return fn();
    } catch (e) {
      if (!isRateLimitError_(e) || i === maxAttempts - 1) throw e;
      const jitter = Math.floor(Math.random() * 250);
      const sleepMs = Math.min(12000, delay) + jitter;
      log(`(Central) Rate limit, retry ${i + 1}/${maxAttempts} in ${sleepMs}ms`);
      Utilities.sleep(sleepMs);
      delay = delay * 2;
    }
  }
  throw new Error('withBackoff_: unreachable');
}

function getThreadLeaseKey_(threadId) {
  return `ENPROCESO_${threadId}`;
}

function escapeLabelForQuery_(labelName) {
  const s = (labelName || '').toString();
  return `"${s.replace(/"/g, '\\"')}"`;
}

function cleanupDiscardStateIfNeeded_() {
  const props = PropertiesService.getScriptProperties();
  const current = getMonthKey_();
  const last = props.getProperty(DISCARD_STATE_MONTH_KEY);
  if (last === current) return;

  const all = props.getProperties();
  const prefix = DISCARD_LASTMSG_PREFIX;
  Object.keys(all).forEach(key => {
    if (key.indexOf(prefix) === 0) {
      props.deleteProperty(key);
    }
  });
  props.setProperty(DISCARD_STATE_MONTH_KEY, current);
}

function getDiscardLastMsgKey_(threadId) {
  return DISCARD_LASTMSG_PREFIX + threadId;
}

function getDiscardLastMsgMs_(threadId) {
  const raw = PropertiesService.getScriptProperties().getProperty(getDiscardLastMsgKey_(threadId));
  if (!raw) return null;
  const parsed = parseInt(raw, 10);
  return isNaN(parsed) ? null : parsed;
}

function setDiscardLastMsgMs_(threadId, dateObj) {
  if (!dateObj) return;
  PropertiesService.getScriptProperties().setProperty(
    getDiscardLastMsgKey_(threadId),
    String(dateObj.getTime())
  );
}

function clearDiscardLastMsgMs_(threadId) {
  PropertiesService.getScriptProperties().deleteProperty(getDiscardLastMsgKey_(threadId));
}

function setThreadLease_(threadId, when) {
  const iso = (when || new Date()).toISOString();
  PropertiesService.getScriptProperties().setProperty(getThreadLeaseKey_(threadId), iso);
}

function clearThreadLease_(threadId) {
  PropertiesService.getScriptProperties().deleteProperty(getThreadLeaseKey_(threadId));
}

function isThreadLeaseStale_(threadId, staleHours) {
  const key = getThreadLeaseKey_(threadId);
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return true;
  const d = new Date(raw);
  if (isNaN(d.getTime())) return true;
  const hours = (Date.now() - d.getTime()) / (1000 * 60 * 60);
  return hours >= (staleHours || 6);
}

function scheduleRerun_(delayMs) {
  const handler = 'procesarEmailsExpensas';
  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    const t = triggers[i];
    if (t.getHandlerFunction && t.getHandlerFunction() === handler) {
      // Evitar acumulación: dejamos solo uno
      ScriptApp.deleteTrigger(t);
    }
  }
  const ms = Math.max(60000, Number(delayMs || 0)); // mínimo 1 min
  ScriptApp.newTrigger(handler).timeBased().after(ms).create();
  log(`(Central) RERUN programado en ${Math.round(ms / 60000)} min`);
}

function isOcrSupported_(filename, contentType) {
  const lower = (filename || '').toLowerCase();
  if (contentType && contentType.indexOf('image/') === 0) return true;
  if (contentType === 'application/pdf') return true;
  return /\.(pdf|png|jpe?g)$/i.test(lower);
}

function hasRelevantAttachment_(message) {
  const attachments = message.getAttachments({ includeInlineImages: true });
  if (!attachments || attachments.length === 0) return false;

  const minBytes = OCR_CONFIG.MIN_INLINE_IMAGE_BYTES || 0;
  for (let i = 0; i < attachments.length; i++) {
    const att = attachments[i];
    const blob = att.copyBlob();
    const name = att.getName ? (att.getName() || '') : '';
    const contentType = blob.getContentType ? (blob.getContentType() || '') : '';
    const lower = name.toLowerCase();
    const isPdf = contentType === 'application/pdf' || lower.endsWith('.pdf');
    const isImage = contentType.indexOf('image/') === 0 || /\.(jpg|jpeg|png|gif)$/i.test(lower);
    if (!isPdf && !isImage) continue;
    if (isImage) {
      const size = blob.getBytes().length;
      if (size < minBytes) continue;
    }
    return true;
  }

  return false;
}

function reabrirDescartadosConAdjunto_(etiquetaDescartado, labelDescartadoName, limit) {
  const query = `in:inbox label:${escapeLabelForQuery_(labelDescartadoName)} ${getMonthStartQuery_()}`;
  const threads = GmailApp.search(query, 0, limit || CONFIG.MAX_THREADS_PER_RUN);
  if (!threads.length) return;

  threads.forEach(thread => {
    const threadId = thread.getId();
    const lastDate = thread.getLastMessageDate();
    const lastMs = getDiscardLastMsgMs_(threadId);
    if (lastMs && lastDate && lastDate.getTime() <= lastMs) {
      return;
    }

    const messages = thread.getMessages();
    const lastMessage = messages[messages.length - 1];
    if (lastMessage && hasRelevantAttachment_(lastMessage)) {
      thread.removeLabel(etiquetaDescartado);
    }

    setDiscardLastMsgMs_(threadId, lastDate);
  });
}

function truncateText_(text, maxChars) {
  if (!text) return '';
  const limit = maxChars || 0;
  if (!limit || text.length <= limit) return text;
  return text.substring(0, limit);
}

function composeBodyWithOcr_(plainBody, ocrText) {
  const safeBody = plainBody || '';
  const safeOcr = truncateText_(ocrText || '', OCR_CONFIG.MAX_OCR_CHARS);
  if (!safeOcr) return safeBody;
  // Poner OCR primero porque validateExpensePaymentWithLLM trunca el cuerpo.
  return `=== OCR (adjuntos/imagenes) ===\n${safeOcr}\n\n=== CUERPO EMAIL ===\n${safeBody}`;
}

function hasMoneyAmount_(text) {
  const t = (text || '').toString();
  // $ 65.199,02 | 65199.02 | 65.000,00 | 1000000
  return /(?:\$\s*)?\d{1,3}(?:[\.\s]\d{3})+(?:,\d{2})?|\$\s*\d{4,}|\b\d{1,3}(?:\.\d{3})+(?:,\d{2})\b|\b\d{4,}(?:[.,]\d{2})\b/.test(t);
}

function hasPaymentEvidence_(text) {
  const t = (text || '').toString().toLowerCase();
  const hasSignal = /(transferenc|comprobante|voucher|constancia|aviso de transferencia|recibiste un pago|pago fue exitoso|se acredit|realizaste|abon|deposit|mercado\s*pago|cbu|cvu|alias|nro|número de operación)/.test(t);
  return hasSignal && (hasMoneyAmount_(t) || /\b(cbu|cvu|alias)\b/.test(t));
}

// Regla especial: descartar transferencias SALIENTES realizadas por Carlos Artuso.
// Criterio: si aparece "Carlos" + "Artuso" y hay señales fuertes de "transferencia saliente",
// entonces etiquetar como descartado, salvo que el mismo email indique explícitamente que fue ENTRANTE.
function matchesCarlosArtusoName_(text) {
  const t = normalizeForMatch_(text);
  if (!t) return false;
  return /\bcarlos\b/.test(t) && /\bartuso\b/.test(t);
}

function hasIncomingTransferSignal_(text) {
  const t = normalizeForMatch_(text);
  if (!t) return false;
  // Señales típicas de transferencia/pago ENTRANTE
  return (
    /\b(recibiste|has recibido|se acredit|cobraste|importe cobrado|importe acreditado|pago recibido)\b/.test(t) &&
    /\b(pago|transferenc)\b/.test(t)
  );
}

function hasOutgoingTransferSignal_(text) {
  const t = normalizeForMatch_(text);
  if (!t) return false;
  // Señales típicas de transferencia SALIENTE
  const hasStrongPhrase = /\b(te informamos que realizaste|realizaste exitosamente|tu transferencia fue exitosa)\b/.test(t);
  const hasVerbAndTransfer = /\b(realizaste|transferiste|enviaste|ordenaste)\b/.test(t) && /\btransferenc\b/.test(t);
  return hasStrongPhrase || hasVerbAndTransfer;
}

function shouldDiscardOutgoingCarlosArtuso_(subject, body) {
  const combined = `${subject || ''}\n${body || ''}`;
  if (!matchesCarlosArtusoName_(combined)) return false;
  // Precedencia: si es entrante, NO descartar (aunque aparezca el nombre).
  if (hasIncomingTransferSignal_(combined)) return false;
  return hasOutgoingTransferSignal_(combined);
}

function extractSlashAddress_(text) {
  const t = (text || '').toString();
  // Calle/Av + número compuesto tipo 2647/51 (toma el match más largo)
  const re = /\b((?:av\.?\s+|avenida\s+)?[a-záéíóúñ][a-záéíóúñ.'\-\s]{1,40}?)\s+(\d{1,5})\s*\/\s*(\d{1,5})\b/ig;
  let match;
  let best = null;
  while ((match = re.exec(t)) !== null) {
    const street = (match[1] || '').replace(/\s+/g, ' ').trim();
    const n1 = match[2];
    const n2 = match[3];
    const candidate = `${street} ${n1}/${n2}`.replace(/\s+/g, ' ').trim();
    if (!best || candidate.length > best.length) best = candidate;
  }
  return best;
}

function normalizeEdForSheet_(edText) {
  if (edText == null) return edText;
  const s = (edText || '').toString().trim();
  if (!s) return '';
  return s.replace(/\s+/g, ' ').toLowerCase();
}

function extractEdFromEmailText_(text) {
  const t = (text || '').toString();
  if (!t) return null;
  // Dirección simple: "Peron 2250" / "Pte. Perón 2248" / "Av Santa Fe 2647/51"
  const re = /\b((?:av\.?\s+|avda\.?\s+|avenida\s+)?[a-záéíóúñ][a-záéíóúñ.'\-\s]{1,40}?)\s+(\d{1,5})(?:\s*\/\s*(\d{1,5}))?\b/i;
  const m = re.exec(t);
  if (!m) return null;
  const street = (m[1] || '').replace(/\s+/g, ' ').trim();
  const n1 = m[2];
  const n2 = m[3];
  const candidate = n2 ? `${street} ${n1}/${n2}` : `${street} ${n1}`;
  return candidate.replace(/\s+/g, ' ').trim();
}

function extractDptoFromEmailText_(text) {
  const t = (text || '').toString();
  if (!t) return null;
  // Preferir patrón con "/" o "dpto"
  let m = /\/\s*([0-9]{1,2}\s*[A-Z])\b/i.exec(t);
  if (m && m[1]) return m[1].replace(/\s+/g, '').toUpperCase();
  m = /\b(?:dpto|depto|dto)\s*[:\-]?\s*([0-9]{1,2}\s*[A-Z])\b/i.exec(t);
  if (m && m[1]) return m[1].replace(/\s+/g, '').toUpperCase();

  // Caso común en asuntos: "pago expensas 8 E [apellido]"
  // Solo aplicar si hay contexto de expensas para evitar capturas accidentales.
  const tn = normalizeForMatch_(t);
  if (tn && (tn.indexOf('expensa') !== -1 || tn.indexOf('expensas') !== -1)) {
    m = /\bexpensas?\b[^\n]{0,30}\b0*([0-9]{1,2})\s*([a-z])\b/i.exec(t);
    if (m && m[1] && m[2]) return (String(parseInt(m[1], 10)) + String(m[2]).toUpperCase()).replace(/\s+/g, '');
  }

  return null;
}

function normalizeDpto_(raw) {
  const s = (raw || '').toString().trim();
  if (!s) return null;
  // Normalizar guiones unicode comunes del OCR a "-"
  const sClean = s.replace(/[‐‑–—−]/g, '-');
  const n = normalizeForMatch_(s);

  // Formatos especiales: no tocar
  if (/\bpb\b/.test(n) || /\bloc(al)?\b/.test(n) || /\bcochera(s)?\b/.test(n)) return sClean;

  // Caso: ya viene con "piso 0 dpto 4-c" y hay que corregir a piso 4 dpto c
  let m = /^piso\s+0\s+dpto\s+0*([0-9]{1,2})\s*-\s*([a-z])$/i.exec(sClean);
  if (m) {
    const piso = String(parseInt(m[1], 10));
    const letra = (m[2] || '').toUpperCase();
    return `Piso ${piso} Dpto ${letra}`;
  }

  // Caso compacto: 04-C | 4-C | 4C | 2/C
  const compact = sClean.replace(/\s+/g, '');
  m = /^0*([0-9]{1,2})(?:[-\/])?([a-z])$/i.exec(compact);
  if (m) {
    const piso2 = String(parseInt(m[1], 10));
    const letra2 = (m[2] || '').toUpperCase();
    return `Piso ${piso2} Dpto ${letra2}`;
  }

  return sClean;
}

function extractCocheraDptoFromText_(text) {
  const t = normalizeForMatch_(text);
  if (!t) return null;
  if (t.indexOf('cochera') === -1) return null;

  const nums = [];
  const seen = {};

  function pushNum(raw) {
    const s = (raw || '').toString().trim();
    if (!s) return;
    const n = parseInt(s.replace(/^0+/, ''), 10);
    if (!n || isNaN(n)) return;
    if (seen[n]) return;
    seen[n] = true;
    nums.push(n);
  }

  // "cocheras 20 y 21", "cochera20-21", "cochera 20,21", etc.
  const re = /\bcocheras?\b[^0-9]{0,12}((?:0*\d{1,4}\s*(?:,|y|e|\/|-)\s*)*0*\d{1,4})/g;
  let m;
  while ((m = re.exec(t)) !== null) {
    const block = m[1] || '';
    const hits = block.match(/\b0*\d{1,4}\b/g);
    if (!hits) continue;
    for (var i = 0; i < hits.length; i++) pushNum(hits[i]);
  }

  if (!nums.length) return null;
  return nums.map(function(n) { return `cochera${n}`; }).join(', ');
}

function findEvidenceLine_(text, needle) {
  if (!text || !needle) return null;
  const needleNorm = normalizeForMatch_(needle);
  if (!needleNorm) return null;
  const lines = (text || '').split(/\r?\n/);
  for (var i = 0; i < lines.length; i++) {
    const line = (lines[i] || '').trim();
    if (!line) continue;
    const lineNorm = normalizeForMatch_(line);
    if (lineNorm.indexOf(needleNorm) !== -1) return line;
  }
  return null;
}

function collectUnitCandidates_(opts) {
  const dpto = [];
  const uf = [];
  const seenDpto = {};
  const seenUf = {};

  function addDpto_(value, source, evidence) {
    const raw = (value || '').toString().trim();
    if (!raw) return;
    if (/cochera/i.test(raw)) return;
    const norm = normalizeDpto_(raw);
    if (!norm) return;
    const key = normalizeForMatch_(norm);
    if (!key || seenDpto[key]) return;
    seenDpto[key] = true;
    dpto.push({ value: norm, source: source || 'unknown', evidence: evidence || null });
  }

  function addUf_(value, source, evidence) {
    const raw = (value || '').toString().trim();
    if (!raw) return;
    const key = normalizeForMatch_(raw);
    if (!key || seenUf[key]) return;
    seenUf[key] = true;
    uf.push({ value: raw, source: source || 'unknown', evidence: evidence || null });
  }

  addDpto_(opts.dptoFromLlm, 'llm', null);
  addUf_(opts.ufFromLlm, 'llm', null);
  addDpto_(opts.dptoFromEmail, 'email', opts.dptoEmailEvidence);
  addDpto_(opts.dptoFromOcr, 'ocr', opts.dptoOcrEvidence);
  addUf_(opts.ufFromOcr, 'ocr', opts.ufOcrEvidence);

  return { dpto: dpto, uf: uf };
}

function hasUnitConflict_(candidates) {
  if (!candidates) return false;
  return (candidates.dpto && candidates.dpto.length > 1) ||
    (candidates.uf && candidates.uf.length > 1);
}

function reconcileUnitsWithLLM_(subject, emailBody, ocrText, candidates) {
  const dptoList = [];
  const ufList = [];
  const dptoMap = {};
  const ufMap = {};

  for (var i = 0; i < (candidates.dpto || []).length; i++) {
    const c = candidates.dpto[i];
    const id = 'D' + (i + 1);
    dptoList.push({
      id: id,
      value: c.value,
      source: c.source,
      evidence: c.evidence || null
    });
    dptoMap[id] = c.value;
  }

  for (var j = 0; j < (candidates.uf || []).length; j++) {
    const u = candidates.uf[j];
    const uid = 'U' + (j + 1);
    ufList.push({
      id: uid,
      value: u.value,
      source: u.source,
      evidence: u.evidence || null
    });
    ufMap[uid] = u.value;
  }

  const emailShort = truncateText_((emailBody || '').toString(), 1500);
  const ocrShort = truncateText_((ocrText || '').toString(), 2000);

  let dptoLines = dptoList.length ? '' : '(sin candidatos)';
  if (dptoList.length) {
    const parts = [];
    for (var k = 0; k < dptoList.length; k++) {
      const item = dptoList[k];
      const ev = item.evidence ? ` | evidence="${item.evidence}"` : '';
      parts.push(`${item.id}: "${item.value}" | source=${item.source}${ev}`);
    }
    dptoLines = parts.join('\n');
  }

  let ufLines = ufList.length ? '' : '(sin candidatos)';
  if (ufList.length) {
    const partsUf = [];
    for (var m = 0; m < ufList.length; m++) {
      const itemUf = ufList[m];
      const evUf = itemUf.evidence ? ` | evidence="${itemUf.evidence}"` : '';
      partsUf.push(`${itemUf.id}: "${itemUf.value}" | source=${itemUf.source}${evUf}`);
    }
    ufLines = partsUf.join('\n');
  }

  const prompt = `Selecciona DPTO y UF correctos SOLO usando la lista de candidatos.
Si no hay evidencia suficiente o hay ambigüedad, devuelve null y confidence="low".

Responde SOLO este JSON:
{"dpto_id":string|null,"uf_id":string|null,"confidence":"high|medium|low","reason":"muy breve"}

Reglas:
- Usa solo ids D* o U* de la lista (no inventes valores).
- Si confidence es "low", dpto_id y uf_id deben ser null.
- Si no aplica UF o DPTO, usa null.

CANDIDATOS DPTO:
${dptoLines}

CANDIDATOS UF:
${ufLines}

EMAIL_SUBJECT: ${subject || ''}
EMAIL_BODY:
${emailShort}

OCR:
${ocrShort}
`;

  try {
    const response = callAI(prompt);
    const cleanResponse = response.replace(/```json\n?|\n?```/g, '').trim();
    const json = JSON.parse(cleanResponse);
    const confidenceRaw = (json.confidence || '').toString().toLowerCase().trim();
    const confidence = (confidenceRaw === 'high' || confidenceRaw === 'medium' || confidenceRaw === 'low') ? confidenceRaw : 'low';

    if (confidence === 'low') {
      return { dpto: null, uf: null, confidence: 'low', reason: (json.reason || '').toString() };
    }

    const dptoId = (json.dpto_id || '').toString().trim();
    const ufId = (json.uf_id || '').toString().trim();
    const dptoValue = dptoId && dptoMap[dptoId] ? dptoMap[dptoId] : null;
    const ufValue = ufId && ufMap[ufId] ? ufMap[ufId] : null;

    return {
      dpto: dptoValue,
      uf: ufValue,
      confidence: confidence,
      reason: (json.reason || '').toString()
    };
  } catch (e) {
    log(`(Central) Error reconciliando DPTO/UF con LLM: ${e.toString()}`);
    return { dpto: null, uf: null, confidence: 'low', reason: 'LLM_ERROR' };
  }
}

function resolveUnitsWithReconciliation_(subject, emailBody, ocrText, candidates) {
  const result = {
    dpto: null,
    uf: null,
    dptoSource: null,
    ufSource: null,
    usedRecon: false,
    lowConfidence: false
  };

  if (!hasUnitConflict_(candidates)) {
    if (candidates.dpto && candidates.dpto.length === 1) {
      result.dpto = candidates.dpto[0].value;
      result.dptoSource = candidates.dpto[0].source;
    }
    if (candidates.uf && candidates.uf.length === 1) {
      result.uf = candidates.uf[0].value;
      result.ufSource = candidates.uf[0].source;
    }
    return result;
  }

  result.usedRecon = true;
  const recon = reconcileUnitsWithLLM_(subject, emailBody, ocrText, candidates);
  if (recon.confidence === 'low') {
    result.lowConfidence = true;
    return result;
  }

  if (recon.dpto) {
    result.dpto = recon.dpto;
    result.dptoSource = 'recon';
  }
  if (recon.uf) {
    result.uf = recon.uf;
    result.ufSource = 'recon';
  }
  return result;
}

function normalizePayor_(value) {
  let s = (value || '').toString().trim();
  if (!s) return null;
  s = s.replace(/\s+/g, ' ');
  s = s.replace(/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/ig, '');
  s = s.replace(/\b\d{2}-\d{8}-\d\b/g, ''); // CUIT con guiones
  s = s.replace(/\b\d{11}\b/g, ''); // CUIT sin guiones
  s = s.replace(/\s{2,}/g, ' ').trim();
  if (!s) return null;
  if (s.length > 60) s = s.substring(0, 60).trim();
  return s || null;
}

const DIRECCIONES_ED_IGNORAR = [
  // Dirección de la administración (no es el edificio del consorcio)
  'sarmiento 1934'
];

function stripAccents_(text) {
  return (text || '').toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function normalizeForMatch_(text) {
  return stripAccents_(text).toLowerCase().replace(/\s+/g, ' ').trim();
}

function isLikelyHumanNameForEdFallback_(value) {
  const raw = (value || '').toString().trim();
  if (!raw) return false;
  if (raw.length > 80) return false;
  if (/@/.test(raw)) return false;
  if (/\d/.test(raw)) return false;

  const norm = normalizeForMatch_(raw);
  const tokens = norm.split(' ').filter(Boolean);
  if (tokens.length < 2) return false;
  if (tokens.length > 6) return false;

  const banned = [
    'varios',
    'expensa',
    'expensas',
    'pago',
    'pagos',
    'transferencia',
    'transferencias',
    'banco',
    'online',
    'banking',
    'cuenta',
    'cbu',
    'cvu',
    'alias',
    'importe',
    'monto',
    'fecha',
    'hora',
    'operacion',
    'operación',
    'nro',
    'numero',
    'número',
    'concepto',
    'motivo',
    'detalle'
  ];
  for (var i = 0; i < banned.length; i++) {
    if (norm.indexOf(banned[i]) !== -1) return false;
  }

  return isLikelyPersonName_(raw);
}

function extractForwardedOriginalBody_(plainBody) {
  const body = (plainBody || '').toString();
  if (!body) return '';
  const marker = '--- CONTENIDO ORIGINAL ---';
  const idx = body.indexOf(marker);
  if (idx === -1) return body;
  return body.substring(idx + marker.length);
}

function extractForwardedSenderName_(plainBody) {
  const body = (plainBody || '').toString();
  if (!body) return null;

  // Buscar en el wrapper del reenvío, antes del "contenido original".
  const beforeOriginal = body.split('--- CONTENIDO ORIGINAL ---')[0] || '';
  const lines = beforeOriginal.split(/\r?\n/).map(function(l) { return (l || '').trim(); }).filter(Boolean);

  // Ej: "Este email fue reenviado automáticamente desde: Diego ... <mail@...>"
  for (var i = 0; i < lines.length; i++) {
    const line = lines[i];
    const ln = normalizeForMatch_(line);
    if (!ln) continue;
    if (ln.indexOf('reenviado') === -1 || ln.indexOf('desde') === -1) continue;

    const idx = line.indexOf(':');
    const after = idx !== -1 ? line.substring(idx + 1) : line;
    const withoutEmail = after.replace(/<[^>]+>/g, ' ').replace(/\([^)]+\)/g, ' ');
    const cleaned = normalizePayor_(withoutEmail) || withoutEmail;
    const candidate = (cleaned || '').replace(/\s+/g, ' ').trim();
    if (candidate && isLikelyHumanNameForEdFallback_(candidate)) return candidate;
  }

  return null;
}

function extractEmailFromHeader_(header) {
  const raw = (header || '').toString();
  if (!raw) return null;
  const angle = /<([^>]+@[^>]+)>/.exec(raw);
  if (angle && angle[1]) return angle[1].trim().toLowerCase();
  const fallback = /([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/i.exec(raw);
  return fallback && fallback[1] ? fallback[1].trim().toLowerCase() : null;
}

function extractEmailDomain_(email) {
  const e = (email || '').toString().toLowerCase().trim();
  if (!e) return null;
  const at = e.lastIndexOf('@');
  if (at === -1) return null;
  const domain = e.substring(at + 1).trim();
  return domain || null;
}

function extractForwardedSenderEmail_(plainBody) {
  const body = (plainBody || '').toString();
  if (!body) return null;

  const beforeOriginal = body.split('--- CONTENIDO ORIGINAL ---')[0] || '';
  const match = /([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/i.exec(beforeOriginal);
  return match && match[1] ? match[1].trim().toLowerCase() : null;
}

function extractSignatureNameCandidate_(plainBody) {
  const original = extractForwardedOriginalBody_(plainBody);
  const lines = (original || '').split(/\r?\n/).map(function(l) { return (l || '').trim(); }).filter(Boolean);
  if (!lines.length) return null;

  const maxLookback = 14;
  for (var i = lines.length - 1; i >= 0 && (lines.length - i) <= maxLookback; i--) {
    const line = lines[i];
    const n = normalizeForMatch_(line);
    if (!n) continue;
    if (n.indexOf('enviado desde mi iphone') !== -1 || n.indexOf('sent from my iphone') !== -1) continue;
    if (/^saludos\b|^saludos cordiales\b|^cordialmente\b|^atte\b|^atentamente\b|^gracias\b/.test(n)) continue;
    if (/http|www\./.test(n)) continue;
    if (/^--+$/.test(n)) continue;
    if (n.indexOf('administraci') !== -1 || n.indexOf('consorcio') !== -1) continue;

    const cleaned = normalizePayor_(line) || line;
    if (cleaned && isLikelyHumanNameForEdFallback_(cleaned)) return cleaned.trim();
  }

  return null;
}

function extractLabeledFieldValues_(text, labelNames, maxLen) {
  const t = (text || '').toString();
  if (!t) return null;
  const labels = (labelNames || []).map(function(l) { return normalizeForMatch_(l); }).filter(Boolean);
  if (!labels.length) return null;
  const limit = typeof maxLen === 'number' ? maxLen : 120;

  const lines = t.split(/\r?\n/).map(function(l) { return (l || '').trim(); }).filter(Boolean);
  const values = [];

  function looksLikeOtherLabelLine_(lineNorm) {
    if (!lineNorm) return false;
    // Heurística: si contiene ":" y empieza con una palabra corta (label), es probablemente otro campo.
    if (/^[a-záéíóúñ\s]{2,25}:\s*/.test(lineNorm)) return true;
    // Campos comunes en comprobantes
    return /^(fecha|hora|importe|monto|cbu|cvu|alias|banco|cuenta|moneda|nombre|beneficiario|destinatario|operacion|operación|nro|numero|número)\b/.test(lineNorm);
  }

  for (var i = 0; i < lines.length; i++) {
    const line = lines[i];
    const ln = normalizeForMatch_(line);
    for (var j = 0; j < labels.length; j++) {
      const label = labels[j];
      if (!label) continue;
      if (ln.indexOf(label) === -1) continue;

      // Match "label: valor" o "label valor"
      const esc = label.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&');
      const re = new RegExp('^\\s*' + esc + '\\s*[:\\-–—]?\\s*(.{1,' + limit + '})\\s*$', 'i');
      const m = re.exec(line);
      let value = null;
      if (m && m[1] && m[1].trim() && normalizeForMatch_(m[1]) !== label) {
        value = m[1].trim();
      } else if (i + 1 < lines.length) {
        const next = lines[i + 1];
        const nextNorm = normalizeForMatch_(next);
        if (!looksLikeOtherLabelLine_(nextNorm)) value = next.trim().substring(0, limit);
      }

      if (value) values.push({ label: label, value: value });
    }
  }

  return values.length ? values : null;
}

function isBankDomain_(domain) {
  const d = (domain || '').toString().toLowerCase().trim();
  if (!d) return false;
  const list = PALABRAS_CLAVE_EXPENSA.dominios_bancos || [];
  for (var i = 0; i < list.length; i++) {
    const base = (list[i] || '').toString().toLowerCase().trim();
    if (!base) continue;
    if (d === base) return true;
    if (d.endsWith('.' + base)) return true;
  }
  return false;
}

function isBankNotification_(message, subject, plainBody) {
  const remitenteRaw = (message.getFrom() || '').toString().toLowerCase();
  const senderEmail = extractEmailFromHeader_(remitenteRaw);
  const forwardedEmail = extractForwardedSenderEmail_(plainBody);
  const domains = [];

  function pushDomain_(email) {
    const d = extractEmailDomain_(email);
    if (d && domains.indexOf(d) === -1) domains.push(d);
  }

  pushDomain_(senderEmail);
  pushDomain_(forwardedEmail);

  const hasBankDomain = domains.some(isBankDomain_);
  const hasKnownSender = PALABRAS_CLAVE_EXPENSA.remitentes_bancos.some(function(r) {
    return remitenteRaw.indexOf(r) !== -1;
  });

  const bodyCore = extractForwardedOriginalBody_(plainBody || '');
  const text = normalizeForMatch_(`${subject || ''} ${bodyCore || ''}`);
  const bankPhrases = [
    'aviso de transferencia',
    'aviso de pago',
    'recibiste una transferencia',
    'detalle de la operacion',
    'te dejamos el detalle de la operacion',
    'numero de operacion',
    'nro de operacion',
    'tipo de transferencia',
    'datos del destinatario'
  ];
  const hasBankPhrase = bankPhrases.some(function(p) { return text.indexOf(p) !== -1; });
  const hasBankFields = /(cbu|cvu|alias|banco|cuenta|cuit|cuil|numero de operacion|nro de operacion|tipo de transferencia)/.test(text);
  const hasBankName = PALABRAS_CLAVE_EXPENSA.bancos.some(function(p) {
    return text.indexOf(normalizeForMatch_(p)) !== -1;
  });

  return hasBankDomain || hasKnownSender || (hasBankPhrase && (hasBankFields || hasBankName));
}

function extractTransferTypeFromText_(text) {
  const found = extractLabeledFieldValues_(text, [
    'tipo de transferencia',
    'tipo transferencia',
    'tipo de operacion',
    'tipo de operación'
  ], 80);
  if (!found) return null;
  for (var i = 0; i < found.length; i++) {
    const val = (found[i] && found[i].value) ? String(found[i].value).trim() : '';
    if (val) return val;
  }
  return null;
}

function matchesProveedorTransferType_(text) {
  const value = extractTransferTypeFromText_(text);
  if (!value) return null;
  return normalizeForMatch_(value).indexOf('proveedor') !== -1 ? value : null;
}

function extractMotivoDetalleNameFromOcr_(ocrText) {
  const found = extractLabeledFieldValues_(ocrText, ['motivo', 'detalle', 'referencia', 'concepto'], 120);
  if (!found) return null;

  const order = { motivo: 1, detalle: 2, referencia: 3, concepto: 4 };
  found.sort(function(a, b) {
    const la = (a && a.label) ? a.label : '';
    const lb = (b && b.label) ? b.label : '';
    return (order[la] || 99) - (order[lb] || 99);
  });

  for (var i = 0; i < found.length; i++) {
    const v = (found[i] && found[i].value) ? String(found[i].value).trim() : '';
    if (!v) continue;
    const cleaned = normalizePayor_(v) || v;
    if (cleaned && isLikelyHumanNameForEdFallback_(cleaned)) return cleaned.trim();
  }

  return null;
}

function extractConceptFieldValue_(text) {
  // Captura "Referencia/Motivo/Concepto" en la misma línea o en la siguiente.
  const found = extractLabeledFieldValues_(text, ['referencia', 'motivo', 'concepto'], 120);
  if (!found) return null;
  const out = [];
  for (var i = 0; i < found.length; i++) {
    const v = (found[i] && found[i].value) ? String(found[i].value).trim() : '';
    if (!v) continue;
    out.push(v);
  }
  return out.length ? out : null;
}

function matchesNonExpenseConcept_(text) {
  const values = extractConceptFieldValue_(text);
  if (!values) return null;
  const normalizedTerms = CONCEPTOS_NO_EXPENSA.map(normalizeForMatch_);
  for (var i = 0; i < values.length; i++) {
    const vNorm = normalizeForMatch_(values[i]);
    for (var j = 0; j < normalizedTerms.length; j++) {
      const term = normalizedTerms[j];
      if (!term) continue;
      // Match por palabra completa cuando es corto (SAC)
      if (term.length <= 3) {
        const reWord = new RegExp('(^|\\b)' + term.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&') + '(\\b|$)', 'i');
        if (reWord.test(vNorm)) return values[i];
      } else if (vNorm.indexOf(term) !== -1) {
        return values[i];
      }
    }
  }
  return null;
}

function normalizeAccountName_(name) {
  if (!name) return '';
  return normalizeForMatch_(name).replace(/[^a-z0-9\s]/g, '').replace(/\s+/g, ' ').trim();
}

function looksLikeLabelValue_(normLine) {
  // Evita tomar como valor una línea que en realidad es otra etiqueta de comprobante.
  const s = (normLine || '').toString().trim();
  if (!s) return false;
  const labelStarts = [
    'cbu', 'cvu', 'alias', 'banco', 'cuenta', 'cta',
    'importe', 'monto', 'total', 'fecha', 'hora',
    'nro', 'numero', 'operacion', 'transaccion',
    'referencia', 'concepto', 'motivo', 'tipo de cuenta',
    'titular', 'beneficiario', 'destinatario', 'ordenante', 'originante'
  ];
  for (var i = 0; i < labelStarts.length; i++) {
    if (s.indexOf(labelStarts[i]) === 0) return true;
  }
  // Si es solo números (CBU, CUIT, referencias), probablemente no es nombre.
  if (/^[0-9\s.\-]{8,}$/.test(s)) return true;
  return false;
}

function isConsorcioLike_(value) {
  const n = normalizeAccountName_(value);
  if (!n) return false;
  // Formas completas
  if (n.indexOf('consorcio') !== -1) return true;
  if (n.indexOf('propiet') !== -1) return true; // propietarios/propiedad

  // Abreviaciones típicas (OCR):
  // - "cons prop ..."
  // - "cons de prop ..."
  // - "cons d pr ..." (Consorcio de Propietarios)
  if (/\bcons\b/.test(n) && (/\bprop\b/.test(n) || /\bpropiet/.test(n) || /\bpr\b/.test(n))) return true;
  if (/\bcons\b\s*(?:d|de)?\s*(?:prop|propiet|pr)\b/.test(n)) return true;

  return false;
}

function isArtusoLike_(value) {
  const n = normalizeAccountName_(value);
  if (!n) return false;
  return n.indexOf('artuso') !== -1;
}

function extractTitularesFromOcr_(ocrText) {
  if (!ocrText) return { origen: null, destino: null };
  const lines = (ocrText || '').split(/\r?\n/).map(function(l) { return (l || '').trim(); }).filter(Boolean);
  let origen = null;
  let destino = null;

  function getValueSameOrNext_(idx) {
    const line = lines[idx] || '';
    let value = null;
    // Caso "Label: valor"
    const m = /:\s*(.+)$/.exec(line);
    if (m && m[1]) value = m[1].trim();
    // Caso "Label valor" en la misma línea
    if (!value) {
      const parts = line.split(/\s+/);
      if (parts.length >= 2) {
        const maybeVal = parts.slice(1).join(' ').trim();
        const maybeNorm = normalizeAccountName_(maybeVal);
        if (maybeVal && maybeNorm && !looksLikeLabelValue_(maybeNorm)) {
          value = maybeVal;
        }
      }
    }
    // Caso línea siguiente como valor
    if (!value && (idx + 1 < lines.length)) {
      const next = (lines[idx + 1] || '').trim();
      const nextNorm = normalizeAccountName_(next);
      if (next && nextNorm && !looksLikeLabelValue_(nextNorm)) {
        value = next;
      }
    }
    return value;
  }

  let expectTitularAs = null; // 'origen' | 'destino'

  for (var i = 0; i < lines.length; i++) {
    const line = lines[i];
    const ln = normalizeAccountName_(line);
    if (!ln) continue;

    // Setear contexto para "Titular" (algunos comprobantes tienen 2 "Titular")
    if (/(cuenta\s*origen|cuenta\s*de\s*debito|cuenta\s*de\s*debito|cuenta\s*a\s*debitar|cuenta\s*d[eé]bito)/.test(ln)) {
      expectTitularAs = 'origen';
    } else if (/(cbu\s*cvu\s*destino|cbu\s*destino|cvu\s*destino|cuenta\s*a\s*acreditar|cuenta\s*destino)/.test(ln)) {
      expectTitularAs = 'destino';
    }

    // Labels explícitos de destino (preferidos)
    if (/^destinatario\b/.test(ln) || /^beneficiario\b/.test(ln) || /^para\b/.test(ln) || /^titular\s+destino\b/.test(ln)) {
      const vDest = getValueSameOrNext_(i);
      if (vDest && !destino) destino = vDest;
      continue;
    }

    // Labels explícitos de origen
    if (/^ordenante\b/.test(ln) || /^originante\b/.test(ln) || /^titular\s+origen\b/.test(ln)) {
      const vOrg = getValueSameOrNext_(i);
      if (vOrg && !origen) origen = vOrg;
      continue;
    }

    // Caso "Titular" con contexto
    if (/^titular\b/.test(ln)) {
      const vTit = getValueSameOrNext_(i);
      if (expectTitularAs === 'origen' && vTit && !origen) origen = vTit;
      if (expectTitularAs === 'destino' && vTit && !destino) destino = vTit;
      expectTitularAs = null;
    }
  }

  // Fallback: línea suelta tipo "a Cons. ...", si no se encontró destino
  if (!destino) {
    for (var k = 0; k < lines.length; k++) {
      const raw = lines[k] || '';
      const norm = normalizeAccountName_(raw);
      if (!norm) continue;
      if (/^a\s+/.test(norm) || /^para\s+/.test(norm)) {
        const after = raw.replace(/^\s*(a|para)\s+/i, '').trim();
        if (after && (isConsorcioLike_(after) || isArtusoLike_(after))) {
          destino = after;
          break;
        }
      }
    }
  }

  return { origen: origen, destino: destino };
}

function getOcrSection_(bodyForAI) {
  const t = (bodyForAI || '').toString();
  const ocrIdx = t.indexOf('=== OCR');
  if (ocrIdx === -1) return '';
  const bodyIdx = t.indexOf('=== CUERPO EMAIL ===');
  const slice = bodyIdx === -1 ? t.substring(ocrIdx) : t.substring(ocrIdx, bodyIdx);
  return slice.replace(/^=== OCR.*?===\s*/i, '').trim();
}

function uniqueNormalizedAnchors_(anchors) {
  const out = [];
  const seen = {};
  (anchors || []).forEach(function(a) {
    const n = normalizeForMatch_(a);
    if (!n) return;
    if (seen[n]) return;
    seen[n] = true;
    out.push(n);
  });
  return out;
}

function getOcrAnchors_() {
  const base = [
    'depto', 'depto:', 'departamento', 'unidad', 'uf', 'edificio', 'caba',
    'total', 'total a pagar', 'a pagar', 'importe', 'monto',
    'fecha', 'venc', 'vencimiento',
    'titular', 'titular:', 'ordenante', 'ordenante:', 'destinatario', 'destinatario:',
    'cbu', 'cvu', 'alias',
    'operación', 'operacion', 'nro', 'número', 'numero', 'referencia',
    'comprobante', 'transferencia',
    'concepto', 'motivo', 'expensas', 'expensa'
  ];

  const fromProject = []
    .concat(PALABRAS_CLAVE_EXPENSA.referencia || [])
    .concat(PALABRAS_CLAVE_EXPENSA.pago || [])
    .concat(PALABRAS_CLAVE_EXPENSA.comprobante || [])
    .concat(PALABRAS_CLAVE_EXPENSA.expensa || []);

  return uniqueNormalizedAnchors_(base.concat(fromProject));
}

const EXTRACTOR_CONTEXT_LIMITS = {
  EMAIL_BODY_MAX_CHARS: 8000,
  OCR_FULL_MAX_CHARS: 12000
};

function detectReceiptLikeOcr_(ocrText) {
  const t = normalizeForMatch_(ocrText);
  if (!t) return false;
  // Nota: OCR a veces elimina espacios/símbolos, así que evitamos word-boundaries estrictos.
  const hasTransferWord = /(transferenc|comprobant|comprobante|voucher|operaci|transacci|nro|nº|n°|numero)/.test(t);
  const hasBankFields = /(cbu|cvu|alias|banco|cuenta|cuit|cuil)/.test(t);
  const hasAmount = hasMoneyAmount_(t) || /(importe|monto|total|a pagar)/.test(t);
  const hasDate =
    /(fecha|hora|venc)/.test(t) ||
    /\b\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}\b/.test(t) ||
    /\b\d{1,2}:\d{2}\b/.test(t);
  // Considerar "comprobante" si hay monto + (transferencia o campos bancarios) + (fecha/hora o referencia/motivo/concepto)
  return hasAmount && (hasTransferWord || hasBankFields) && (hasDate || /(referencia|motivo|concepto)/.test(t));
}

function countReceiptMarkers_(ocrText) {
  if (!ocrText) return 0;
  const markers = ocrText.match(/\[(?:adjunto|inline): /g);
  return markers ? markers.length : 0;
}

function splitOcrIntoBlocks_(ocrText) {
  const text = (ocrText || '').toString();
  const lines = text.split(/\r?\n/);
  const blocks = [];
  let current = null;

  function pushCurrent_() {
    if (current) {
      current.text = current.lines.join('\n').trim();
      blocks.push(current);
      current = null;
    }
  }

  for (var i = 0; i < lines.length; i++) {
    const line = lines[i] || '';
    const m = /^\[(adjunto|inline):\s*(.+?)\]/i.exec(line);
    if (m) {
      pushCurrent_();
      current = { source: m[1].toLowerCase(), name: m[2], lines: [] };
      continue;
    }
    if (!current) {
      current = { source: 'unknown', name: '', lines: [] };
    }
    current.lines.push(line);
  }
  pushCurrent_();
  return blocks;
}

function getReceiptLikeBlocks_(ocrText) {
  const blocks = splitOcrIntoBlocks_(ocrText);
  const receipts = [];
  for (var i = 0; i < blocks.length; i++) {
    const blk = blocks[i];
    if (blk && blk.text && detectReceiptLikeOcr_(blk.text)) {
      receipts.push(blk);
    }
  }
  return receipts;
}

function extractPayerCuitFromOcr_(ocrText) {
  const t = (ocrText || '').toString();
  if (!t) return null;

  const cuitRe = /\b(cuit\/cuil|cuit|cuil)\b\s*[: ]\s*([0-9]{2}-[0-9]{8}-[0-9])\b/ig;
  let match;
  while ((match = cuitRe.exec(t)) !== null) {
    const cuit = match[2];
    const before = t.substring(0, match.index);
    const near = before.substring(Math.max(0, before.length - 200));
    const nearNorm = normalizeForMatch_(near);

    // Si cerca dice "para/destinatario/beneficiario", probablemente no es pagador
    if (/\b(para|destinatario|beneficiario)\b/.test(nearNorm)) continue;

    // Si cerca dice "de/originante/ordenante/titular", es buena señal de pagador
    const hasPayerLabel = /\b(de|originante|ordenante|titular|origen)\b/.test(nearNorm);

    // Buscar nombre en líneas previas
    const lines = before.split(/\r?\n/).map(function(l) { return (l || '').trim(); }).filter(Boolean);
    for (var i = lines.length - 1; i >= 0; i--) {
      const line = lines[i];
      const lineNorm = normalizeForMatch_(line);
      if (!line) continue;
      if (lineNorm.indexOf('cuit') !== -1 || lineNorm.indexOf('cuil') !== -1) continue;
      if (lineNorm === 'de' || lineNorm === 'para' || lineNorm === 'origen' || lineNorm === 'destino') continue;
      if (/mercado\s*pago/i.test(line)) continue;
      if (!isLikelyPersonName_(line)) continue;
      if (!hasPayerLabel && i >= 1) {
        const prevNorm = normalizeForMatch_(lines[i - 1]);
        if (prevNorm.indexOf('para') !== -1 || prevNorm.indexOf('destinatario') !== -1) continue;
      }
      return { name: line.trim(), cuit: cuit, evidence: `${match[1]}: ${cuit}` };
    }
  }
  return null;
}

function buildOcrEvidence_(ocrText, anchors, windowSize, maxChars) {
  const t = (ocrText || '').toString();
  if (!t) return '';
  const lines = t.split(/\r?\n/).map(function(l) { return (l || '').trim(); }).filter(Boolean);
  if (lines.length === 0) return '';

  const a = anchors && anchors.length ? anchors : getOcrAnchors_();
  const w = typeof windowSize === 'number' ? windowSize : 2;
  const includeIdx = {};

  for (var i = 0; i < lines.length; i++) {
    const ln = normalizeForMatch_(lines[i]);
    let hit = false;
    for (var j = 0; j < a.length; j++) {
      if (ln.indexOf(a[j]) !== -1) { hit = true; break; }
    }
    if (!hit) continue;
    const start = Math.max(0, i - w);
    const end = Math.min(lines.length - 1, i + w);
    for (var k = start; k <= end; k++) includeIdx[k] = true;
  }

  const selected = [];
  for (var idx = 0; idx < lines.length; idx++) {
    if (includeIdx[idx]) selected.push(lines[idx]);
  }
  if (selected.length === 0) return '';

  const joined = selected.join('\n');
  return truncateText_(joined, maxChars || 3500);
}

function extractPayorFromOcr_(ocrText) {
  const t = (ocrText || '').toString();
  if (!t) return null;
  let m = /\bTitular:\s*([^\n\r]{2,80})/i.exec(t);
  if (m && m[1]) return m[1].trim();
  m = /\bOrdenante:\s*([^\n\r]{2,80})/i.exec(t);
  if (m && m[1]) return m[1].trim();
  return null;
}

function extractFechaFromOcr_(ocrText) {
  const t = (ocrText || '').toString();
  if (!t) return null;
  // DD/MM/YYYY o DD/MM/YY
  const m = /\b(\d{2})\/(\d{2})\/(\d{2}|\d{4})\b/.exec(t);
  if (!m) return null;
  const yy = m[3].length === 4 ? m[3].substring(2) : m[3];
  return `${m[1]}-${m[2]}-${yy}`;
}

function parseDateFromText_(value) {
  const s = (value || '').toString().trim();
  if (!s) return null;
  // Acepta DD-MM-YY(YY) o DD/MM/YY(YY) y tolera texto extra (p.ej. hora).
  const m = /(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2}|\d{4})/.exec(s);
  if (!m) return null;
  const day = parseInt(m[1], 10);
  const month = parseInt(m[2], 10);
  let year = parseInt(m[3], 10);
  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
  if (m[3].length === 2) year = 2000 + year;
  if (day < 1 || day > 31 || month < 1 || month > 12) return null;
  return { year: year, month: month, day: day };
}

function ymdInt_(parts) {
  if (!parts) return null;
  return (parts.year * 10000) + (parts.month * 100) + parts.day;
}

function extractMontosFromOcr_(ocrText) {
  const t = (ocrText || '').toString();
  if (!t) return { montos: [], monto_total: null };
  const lines = t.split(/\r?\n/).map(function(l) { return (l || '').trim(); }).filter(Boolean);
  const montos = [];

  const moneyRe = /(?:\$?\s*)(\d{1,3}(?:[.\s]\d{3})+(?:,\d{2})|\d{4,}(?:[.,]\d{2})?)/g;
  for (var i = 0; i < lines.length; i++) {
    const lnNorm = normalizeForMatch_(lines[i]);
    if (!(/\b(total|a pagar|importe|monto)\b/.test(lnNorm))) continue;
    moneyRe.lastIndex = 0;
    let m;
    while ((m = moneyRe.exec(lines[i])) !== null) {
      const parsed = parseAmount_(m[1]);
      if (parsed != null) montos.push(parsed);
    }
  }

  // Monto total: preferir línea con "total" y "a pagar"
  let montoTotal = null;
  for (var j = 0; j < lines.length; j++) {
    const lnN = normalizeForMatch_(lines[j]);
    if (!(lnN.indexOf('total') !== -1 && (lnN.indexOf('a pagar') !== -1 || lnN.indexOf('apagar') !== -1))) continue;
    moneyRe.lastIndex = 0;
    const mm = moneyRe.exec(lines[j]);
    if (mm && mm[1]) {
      montoTotal = parseAmount_(mm[1]);
      break;
    }
  }

  // Dedupe montos
  const uniq = [];
  const seen = {};
  montos.forEach(function(v) {
    const key = String(v);
    if (seen[key]) return;
    seen[key] = true;
    uniq.push(v);
  });

  return { montos: uniq, monto_total: montoTotal };
}

function buildOcrHints_(ocrText) {
  const hints = {};
  const ed = extractEdFromOcr_(ocrText);
  const dpto = extractDeptoFromOcr_(ocrText) || extractDptoFromUnidadLike_(ocrText);
  const uf = extractUfFromOcr_(ocrText);
  const payor = extractPayorFromOcr_(ocrText);
  const fecha = extractFechaFromOcr_(ocrText);
  const amounts = extractMontosFromOcr_(ocrText);

  if (ed) hints.ed = ed;
  if (dpto) hints.dpto = dpto;
  if (uf) hints.uf = uf;
  if (payor) hints.pagador = payor;
  if (fecha) hints.fecha = fecha;
  if (amounts.montos && amounts.montos.length) hints.montos = amounts.montos;
  if (amounts.monto_total != null) hints.monto_total = amounts.monto_total;

  return hints;
}

function buildExtractorContext_(subject, plainBody, ocrFullText) {
  const subjectSafe = (subject || '').toString();
  const emailBodyFull = truncateText_((plainBody || '').toString(), EXTRACTOR_CONTEXT_LIMITS.EMAIL_BODY_MAX_CHARS);
  const ocrFull = truncateText_((ocrFullText || '').toString(), EXTRACTOR_CONTEXT_LIMITS.OCR_FULL_MAX_CHARS);

  const hints = buildOcrHints_(ocrFull);
  const ocrEvidence = buildOcrEvidence_(ocrFull, getOcrAnchors_(), 2, 3500);

  return [
    'EMAIL_SUBJECT:',
    subjectSafe,
    '',
    'EMAIL_BODY_FULL:',
    emailBodyFull,
    '',
    'OCR_FULL:',
    ocrFull,
    '',
    'HINTS_JSON (sugerencias; pueden estar incompletas o equivocadas):',
    JSON.stringify(hints || {}),
    '',
    'OCR_EVIDENCE (recortado por ventanas):',
    ocrEvidence || ''
  ].join('\n');
}

function isIgnoredEd_(edText) {
  const n = normalizeForMatch_(edText);
  return DIRECCIONES_ED_IGNORAR.some(d => n.indexOf(d) !== -1);
}

function isValidEd_(edText) {
  const ed = (edText || '').toString().trim();
  if (!ed) return false;
  if (isIgnoredEd_(ed)) return false;
  // Debe parecer dirección: letras + número (o número compuesto)
  if (!/[a-záéíóúñ].*\d/i.test(ed)) return false;
  const n = normalizeForMatch_(ed);
  // Evitar texto narrativo típico de recargos/vencimientos
  if (/\bpor pago\b|\bfuera de termino\b|\bfuera de término\b|\brecargo\b|\bhasta el\b|\bvencim/i.test(n)) return false;
  return true;
}

function extractEdFromOcr_(ocrText) {
  const t = (ocrText || '').toString();
  if (!t) return null;

  // Prioridad 1: línea con "- CABA" (típicamente dirección del edificio)
  const reCaba = /\b([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s.'\-]{2,60}?\s+\d{1,5}(?:\s*\/\s*\d{1,5})?)\s*-\s*CABA\b/ig;
  let match;
  let best = null;
  while ((match = reCaba.exec(t)) !== null) {
    const candidate = (match[1] || '').replace(/\s+/g, ' ').trim();
    if (isValidEd_(candidate) && (!best || candidate.length > best.length)) best = candidate;
  }
  if (best) return best;

  // Prioridad 2: cualquier "calle + número(/número)"
  const candidates = [];
  const reGeneric = /\b([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s.'\-]{2,60}?\s+\d{1,5}(?:\s*\/\s*\d{1,5})?)\b/g;
  while ((match = reGeneric.exec(t)) !== null) {
    const candidate2 = (match[1] || '').replace(/\s+/g, ' ').trim();
    if (isValidEd_(candidate2)) candidates.push(candidate2);
  }
  if (candidates.length === 0) return null;
  candidates.sort((a, b) => b.length - a.length);
  return candidates[0];
}

function extractDeptoFromOcr_(ocrText) {
  const t = (ocrText || '').toString();
  if (!t) return null;
  // "Depto: 03-C"
  const m = /\bDepto:\s*([0-9]{1,3})\s*-\s*([A-Z])\b/i.exec(t);
  if (!m) return null;
  const num = m[1].padStart(2, '0');
  const letter = (m[2] || '').toUpperCase();
  return `${num}-${letter}`;
}

function extractDeptoFromObservacionesOcr_(ocrText) {
  const t = normalizeForMatch_(ocrText);
  if (!t) return null;
  // Patrones tipo "expensas dic 8 e", "expensas 8e", "expensas 8 e <texto>"
  const re = /\bexpensas?[^\n]{0,40}\b0*([0-9]{1,2})\s*([a-z])\b/i;
  const m = re.exec(ocrText);
  if (!m || !m[1] || !m[2]) return null;
  const num = String(parseInt(m[1], 10));
  const letter = String(m[2]).toUpperCase();
  return `Piso ${num} Dpto ${letter}`;
}

function extractDptoFromUnidadLike_(ocrText) {
  const t = (ocrText || '').toString();
  if (!t) return null;
  // "Unidad: 02-C" | "Unidad: 2C" | "Unidad: 2/C" | "UF: 02-C"
  let m = /\b(?:unidad(?:\s*funcional)?|uf)\b\s*:\s*0*([0-9]{1,3})\s*[-\/]?\s*([A-Z])\b/i.exec(t);
  if (!m) m = /\b(?:unidad(?:\s*funcional)?|uf)\b\s*:\s*0*([0-9]{1,3})\s*\/\s*([A-Z])\b/i.exec(t);
  if (!m) return null;
  const num = String(m[1] || '').padStart(2, '0');
  const letter = (m[2] || '').toUpperCase();
  if (!num || !letter) return null;
  return `${num}-${letter}`;
}

function extractUfFromOcr_(ocrText) {
  const t = (ocrText || '').toString();
  if (!t) return null;
  // "Unidad: 0035" | "Unidad Funcional: 35" | "UF: 0035"
  const m = /\b(?:unidad(?:\s*funcional)?|uf)\b\s*:\s*([0-9]{2,6})\b/i.exec(t);
  if (!m) return null;
  const v = (m[1] || '').trim();
  // Si el valor tiene letra, no es UF numérica
  if (/[A-Z]/i.test(v)) return null;
  return v;
}

function parseAmount_(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number') return isNaN(value) ? null : value;
  let s = String(value).trim();
  if (!s) return null;
  s = s.replace(/\$/g, '').replace(/\s+/g, '');
  // Formatos comunes AR: 65.199,02 o 80000,00
  if (s.indexOf(',') !== -1 && s.indexOf('.') !== -1) {
    // asumir '.' miles y ',' decimal
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.indexOf(',') !== -1 && s.indexOf('.') === -1) {
    s = s.replace(',', '.');
  } else {
    // eliminar separadores de miles por espacios ya removidos; dejar punto decimal si existe
    s = s.replace(/(\d)\.(\d{3})(\D|$)/g, '$1$2$3');
  }
  const n = Number(s);
  return isNaN(n) ? null : n;
}

function extractMontosFromTextLoose_(text) {
  const t = (text || '').toString();
  if (!t) return [];
  const moneyRe = /(?:\$?\s*)(\d{1,3}(?:[.\s]\d{3})+(?:,\d{2})|\d{4,}(?:[.,]\d{2})?)/g;
  const out = [];
  const seen = {};
  let m;
  while ((m = moneyRe.exec(t)) !== null) {
    const parsed = parseAmount_(m[1]);
    if (parsed == null) continue;
    const key = String(parsed);
    if (seen[key]) continue;
    seen[key] = true;
    out.push(parsed);
  }
  return out;
}

function ocrBlobViaDrive_(blob, filename, ocrLanguage) {
  if (!canUseDriveOcr_()) {
    throw new Error('Drive API no habilitada (Advanced Service: Drive)');
  }
  if (OCR_CALLS_RUN >= OCR_CONFIG.MAX_OCR_CALLS_PER_RUN) {
    throw new Error('OCR_BUDGET_EXCEEDED');
  }
  // Contar el intento (aunque falle) para respetar presupuesto y evitar loops
  OCR_CALLS_RUN++;

  const tempFile = DriveApp.createFile(blob.setName(filename));
  let docId = null;

  try {
    const resource = {
      title: filename,
      mimeType: 'application/vnd.google-apps.document'
    };
    const docFile = withBackoff_(function() {
      return Drive.Files.copy(resource, tempFile.getId(), {
        ocr: true,
        ocrLanguage: ocrLanguage || 'es'
      });
    }, 6);
    docId = docFile && docFile.id ? docFile.id : null;
    if (!docId) {
      throw new Error('No se pudo crear el documento OCR');
    }
    const doc = DocumentApp.openById(docId);
    return doc.getBody().getText();
  } finally {
    try {
      if (docId) DriveApp.getFileById(docId).setTrashed(true);
    } catch (e1) {}
    try {
      DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    } catch (e2) {}
  }
}

function ocrAttachments_(attachments, opts) {
  if (!attachments || attachments.length === 0) return '';
  if (!OCR_CONFIG.ENABLED) return '';

  if (!canUseDriveOcr_()) {
    log('(Central) OCR deshabilitado: falta habilitar Advanced Service "Drive API"');
    return '';
  }

  const parts = [];
  let used = 0;

  for (var i = 0; i < attachments.length && used < OCR_CONFIG.MAX_BLOBS; i++) {
    const att = attachments[i];
    const blob = att.copyBlob();
    const name = att.getName ? (att.getName() || ('sin_nombre_' + i)) : ('sin_nombre_' + i);
    const contentType = blob.getContentType ? (blob.getContentType() || '') : '';
    const bytesLen = blob.getBytes().length;

    if (!isOcrSupported_(name, contentType)) continue;
    if (bytesLen > OCR_CONFIG.MAX_BYTES) continue;

    try {
      const text = ocrBlobViaDrive_(blob, name, OCR_CONFIG.OCR_LANGUAGE);
      const cleaned = (text || '').trim();
      if (!cleaned) continue;
      parts.push('[' + ((opts && opts.source) ? opts.source : 'archivo') + ': ' + name + ']\n' + cleaned);
      used++;
    } catch (e) {
      log('(Central) OCR error (' + name + '): ' + e.toString());
    }
  }

  return parts.length ? parts.join('\n\n') : '';
}

function extraerTextoDeAdjuntosExpensa(message) {
  if (!OCR_CONFIG.ENABLED) return '';
  // Adjuntos "reales" (no inline)
  const attachments = message.getAttachments({ includeInlineImages: false });
  return ocrAttachments_(attachments, { source: 'adjunto' });
}

function extraerTextoDeImagenesEnCuerpo(message) {
  if (!OCR_CONFIG.ENABLED) return '';

  const withInline = message.getAttachments({ includeInlineImages: true });
  const withoutInline = message.getAttachments({ includeInlineImages: false });

  const fp = function(att, idx) {
    const blob = att.copyBlob();
    const bytesLen = blob.getBytes().length;
    const name = att.getName ? (att.getName() || ('sin_nombre_' + idx)) : ('sin_nombre_' + idx);
    const ct = blob.getContentType ? (blob.getContentType() || '') : '';
    return [name, ct, bytesLen].join('|');
  };

  const base = {};
  for (var i = 0; i < withoutInline.length; i++) base[fp(withoutInline[i], i)] = true;

  const inlineOnly = [];
  for (var j = 0; j < withInline.length; j++) {
    const key = fp(withInline[j], j);
    if (!base[key]) inlineOnly.push(withInline[j]);
  }

  // Filtrar imágenes chicas típicas de firma/logo
  const filtered = [];
  for (var k = 0; k < inlineOnly.length; k++) {
    const blob = inlineOnly[k].copyBlob();
    if (blob.getBytes().length >= OCR_CONFIG.MIN_INLINE_IMAGE_BYTES) filtered.push(inlineOnly[k]);
  }

  return ocrAttachments_(filtered, { source: 'inline' });
}

// ==================== ADAPTADORES IA (no modificar) ====================
const AI_ADAPTERS = {
  openai: {
    url: 'https://api.openai.com/v1/chat/completions',
    formatRequest: (prompt, model) => ({
      model: model,
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.1
    }),
    getHeaders: (apiKey) => ({
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    }),
    parseResponse: (json) => json.choices[0].message.content
  },
  
  gemini: {
    url: (model, apiKey) => 
      `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`,
    formatRequest: (prompt, model) => ({
      contents: [{ parts: [{ text: prompt }] }]
    }),
    getHeaders: (apiKey) => ({
      'Content-Type': 'application/json'
    }),
    parseResponse: (json) => json.candidates[0].content.parts[0].text
  },
  
  anthropic: {
    url: 'https://api.anthropic.com/v1/messages',
    formatRequest: (prompt, model) => ({
      model: model,
      max_tokens: 1024,
      messages: [{ role: 'user', content: prompt }]
    }),
    getHeaders: (apiKey) => ({
      'x-api-key': apiKey,
      'Content-Type': 'application/json',
      'anthropic-version': '2023-06-01'
    }),
    parseResponse: (json) => json.content[0].text
  }
};

// Reutilizamos el mismo motor de clasificación que en los otros scripts

const PALABRAS_CLAVE_EXPENSA = {
  expensa: [
    'expensa', 'expensas',
    // Typos comunes
    'expesnas',   // intercambio n/s
    'espensas',   // x→s (muy común en español)
    'expnsas',    // falta e
    'expesas',    // falta n
    'expenss',    // falta a
    'expenas',    // falta s
    'exppensas',  // p duplicada
    'expensass',  // s duplicada
    'expenzas',   // s→z
    // Formato especial
    'expensas uf' // formato "Expensas UF 64"
  ],
  pago: [
    'pago', 'pagado', 'pagó', 'pagué', 'pague',
    'aboné', 'abone', 'abonado',
    'transferencia', 'transferí',
    'depósito', 'deposité', 'deposito',
    'aviso de pago', 'aviso de transferencia'
  ],
  comprobante: [
    'comprobante',
    'comprobante de pago',
    'comprobante de transferencia',
    'comprobante adjunto',
    'adjunto comprobante',
    'envío comprobante',
    'envio comprobante',
    'les adjunto el comprobante',
    'recibo',
    'voucher',
    'constancia',
    'ticket'
  ],
  consorcio: [
    'unidad',
    'departamento',
    'depto',
    'ph',
    'casa',
    'cochera',
    'torre',
    'edificio',
    'oficina',
    'local',
    'locales',
    'piso'
  ],
  referencia: [
    'cbu', 'cvu', 'alias', 'referencia', 'nro', 'número',
    'cuit', 'cuil',
    'número de operación', 'n° de operación', 'nro de operación',
    'número de comprobante', 'nro de comprobante',
    'cuenta origen', 'cuenta destino',
    'datos del pago', 'datos del destinatario'
  ],
  frases_pago: [
    // Notificaciones de cobro recibido
    '¡recibiste un pago', 'recibiste un pago', 'has recibido un pago',
    'se acreditó un pago', 'se acredito un pago',
    'recibiste una transferencia',
    'importe cobrado', 'importe acreditado',
    'pago recibido',
    'hemos recibido tu pago',
    'se registró un pago', 'se registro un pago',
    'cobraste $',
    'el dinero se acreditó',
    // Notificaciones de pago enviado
    'realizaste una transferencia',
    'realizaste la siguiente transferencia',
    'te informamos que realizaste',
    'información sobre tu transferencia',
    'se realizó la siguiente transferencia',
    // Notificaciones de sistemas de consorcios
    'se ha generado la siguiente notificación de pago',
    'notificación de pago',
    // Frases manuales de propietarios
    'adjunto comprobante de pago',
    'les adjunto el comprobante de pago',
    'envío comprobante de pago',
    'adjunto pago por expensas',
    'pagué por expensas',
    'van los pagos de expensas',
    'pago expensas' // común en asuntos
  ],
  bancos: [
    // Bancos tradicionales
    'banco ciudad',
    'banco nación', 'banco nacion',
    'banco galicia', 'bancogalicia', 'officebanking',
    'banco santander', 'santander.com.ar',
    'banco macro',
    'banco patagonia',
    'bbva', 'bbva.com.ar',
    // Billeteras digitales
    'mercado pago', 'mercadopago',
    'naranja x', 'naranjax',
    'uala', 'ualá',
    // Sistemas de cobro
    'mi-qr.com.ar', 'mi qr', 'cobros@mi-qr',
    'consorcios en red',
    // Tarjetas
    'visa',
    'mastercard'
  ],
  // NUEVO: Remitentes conocidos de notificaciones bancarias
  remitentes_bancos: [
    'cobros@mi-qr.com.ar',
    'avisos@bbva.com.ar',
    '@bancogalicia.com.ar',
    '@mails.santander.com.ar',
    'go@bancogalicia.com.ar',
    'mensajesyavisos@mails.santander.com.ar',
    '@mercadopago.com',
    '@naranjax.com'
  ],
  // NUEVO: Dominios bancarios confiables (para detectar avisos aunque cambie el remitente)
  dominios_bancos: [
    'bancogalicia.com.ar',
    'mails.santander.com.ar',
    'santander.com.ar',
    'bbva.com.ar',
    'mi-qr.com.ar',
    'mercadopago.com',
    'naranjax.com'
  ],
  // NUEVO: Calles/direcciones comunes de Buenos Aires (para detectar "pago expensas [DIRECCIÓN]")
  direcciones: [
    'lavalle', 'peron', 'perón', 'paraguay', 'ocampo', 'córdoba', 'cordoba',
    'santa fe', 'guemes', 'güemes', 'billinghurst', 'sarmiento', 'alvear',
    'corrientes', 'rivadavia', 'callao', 'florida', 'maipu', 'maipú',
    'libertad', 'talcahuano', 'montevideo', 'rodriguez peña', 'ayacucho',
    'junin', 'junín', 'uriburu', 'pueyrredon', 'pueyrredón', 'sanchez de bustamante'
  ]
};

// NUEVO: Palabras que indican que NO es un pago nuevo (para evitar falsos positivos)
const PALABRAS_EXCLUSION_FUERTE = [
  'gracias por informar el pago',  // es una respuesta, no un pago
  'muchas gracias por informar',
  'recibido, gracias',
  'confirmamos recepción',
  // Excluir resúmenes automáticos internos (ruido)
  'resumen procesamiento expensas',
  // Entidades/temas que NO son pago de expensas (proveedores, cámaras, cobranzas)
  'cámara argentina de la propiedad horizontal',
  'camara argentina de la propiedad horizontal',
  'actividades inmobiliarias',
  'cooperativa de trabajo',
  'mantenimiento integral',
  'dto. comercial y cobranzas',
  'comercial y cobranzas',
  'presupuesto',
  'impermeabilización',
  'pintura en el sector de terrazas',
  'reparacion',
  'reparación'
];

const REGEX_RESUMEN_PROCESAMIENTO = /^resumen\s+procesamiento\s+expensas\s+-\s+\d{2}\/\d{2}\/\d{4}$/i;

// Conceptos típicos de transferencias que NO corresponden a expensas (falsos positivos frecuentes)
// Criterio: pagos de sueldos/aguinaldo/haberes/nómina, aunque venga "consorcio" o dirección.
const CONCEPTOS_NO_EXPENSA = [
  'sac',
  'aguinaldo',
  'sueldo',
  'sueldos',
  'haberes',
  'nomina',
  'nómina',
  'liquidacion',
  'liquidación',
  'payroll'
];

// ==================== PROMPT DE VALIDACIÓN LLM ====================
const VALIDATION_PROMPT = `Tu tarea: entender la INTENCIÓN del email.
Clasifica el email en 3 estados:
- "accept": la intención es INFORMAR/CONFIRMAR un pago de expensas (aunque falten algunos datos).
- "reject": la intención claramente NO es informar pago de expensas (p.ej. presupuesto, reclamo, proveedor, cobranzas/deuda, cuota de cámara, comunicación administrativa sin pago).
- "uncertain": no está claro; ante la duda, dejar pasar como revisión humana.

El email puede incluir una sección "=== OCR (adjuntos/imagenes) ===" con texto extraído.

DEVUELVE SOLO este JSON válido (sin texto extra):
{"decision":"accept|reject|uncertain","reason":"texto muy breve"}

Guía (no exhaustiva):

ACEPTAR si ves intención de pago:
- "pagué/pagado/transferí/aboné", "aviso de transferencia", "tu pago fue exitoso", "recibiste un pago"
- comprobante/voucher/constancia adjunta o datos de operación + monto
- menciona expensas/consorcio/unidad/cochera/dirección del edificio

RECHAZAR solo si la intención NO es pago de expensas:
- presupuesto/pintura/impermeabilización/reparación/proveedor/factura de servicio
- cobranzas/deuda/saldo pendiente/recordatorio de pago
- transferencias cuyo Referencia/Motivo/Concepto sea SAC/aguinaldo/sueldo/haberes/nómina
- recibo/cuota/inscripción/pago a una entidad (ej: Cámara) que no sea expensas del consorcio
- reclamos/avisos (p.ej. "se encuentra esto en la cochera") sin pago

INCERTO (REVIEW) si:
- hay comprobante/transferencia pero no queda claro que sea por expensas
- hay señales mezcladas o falta contexto

Regla: reject SOLO con evidencia negativa clara. Si no, uncertain.

El campo reason debe ser muy conciso (máx 12 palabras).`;

const BANK_ROLES_PROMPT = `Tu tarea: en avisos bancarios de transferencia, detectar pagos a proveedores que NO son expensas.

Devuelve SOLO este JSON válido (sin texto extra):
{"should_discard":boolean,"confidence":"high|medium|low","payer":string|null,"payee":string|null,"transfer_type":"expensa|proveedor|sueldo|unknown","reason":"muy breve"}

Reglas:
- "payer" es quien ENVIA el dinero, "payee" quien RECIBE.
- "should_discard" = true SOLO si hay evidencia clara de que el pagador es consorcio/administración
  y el receptor es persona/empresa externa, o si dice explícitamente "Tipo de transferencia: Proveedor".
- Si el pago va al consorcio/administración (beneficiario/destinatario consorcio), no descartar.
- Si no hay evidencia clara, usa should_discard=false y confidence="low".
`;

function getRequiredConfigString_(value, keyName) {
  const v = (value || '').toString().trim();
  if (!v) throw new Error(`CONFIG inválida: falta ${keyName}`);
  return v;
}

// ==================== FUNCIONES DE IA ====================

/**
 * Llama a la API de IA configurada (vendor-agnostic)
 * @param {string} prompt - El prompt a enviar
 * @returns {string} - La respuesta de la IA
 */
function callAI(prompt) {
  const adapter = AI_ADAPTERS[AI_CONFIG.PROVIDER];
  
  if (!adapter) {
    throw new Error(`Proveedor no soportado: ${AI_CONFIG.PROVIDER}`);
  }
  
  // Obtener API key de forma segura desde PropertiesService
  const apiKey = getAIApiKey();
  
  // Construir URL (algunos proveedores la construyen dinámicamente)
  const url = typeof adapter.url === 'function' 
    ? adapter.url(AI_CONFIG.MODEL, apiKey)
    : adapter.url;
  
  const options = {
    method: 'post',
    headers: adapter.getHeaders(apiKey),
    payload: JSON.stringify(adapter.formatRequest(prompt, AI_CONFIG.MODEL)),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  // Validar respuesta HTTP
  if (statusCode !== 200) {
    log(`Error API IA (${statusCode}): ${responseText}`);
    throw new Error(`Error ${AI_CONFIG.PROVIDER} (${statusCode}): ${responseText.substring(0, 200)}`);
  }
  
  const json = JSON.parse(responseText);
  return adapter.parseResponse(json);
}

/**
 * Valida si el email es un pago de expensas usando LLM (fail-open con 3 estados).
 * @param {string} subject - Asunto del email
 * @param {string} body - Cuerpo del email
 * @returns {Object} - { isValid: boolean, review: boolean, reviewReason?: string }
 */
function validateExpensePaymentWithLLM(subject, body) {
  // Truncar el cuerpo para optimizar tokens (2000 chars es suficiente para contexto)
  const truncatedBody = (body || '').substring(0, 2000);
  const prompt = VALIDATION_PROMPT + `\n\nEmail:\nAsunto: ${subject}\nCuerpo: ${truncatedBody}`;
  
  try {
    const response = callAI(prompt);
    // Limpiar respuesta (a veces la IA agrega backticks de markdown)
    const cleanResponse = response.replace(/```json\n?|\n?```/g, '').trim();
    const json = JSON.parse(cleanResponse);
    const decision = (json.decision || '').toString().toLowerCase().trim();
    const reason = (json.reason || '').toString().trim();
    const shortReason = reason.length > 80 ? reason.substring(0, 80) : reason;

    if (decision === 'reject') {
      return { isValid: false, review: false, reviewReason: shortReason || 'LLM_REJECT', decision };
    }
    if (decision === 'uncertain') {
      return { isValid: true, review: true, reviewReason: shortReason || 'LLM_UNCERTAIN', decision };
    }
    if (decision === 'accept') {
      return { isValid: true, review: false, decision };
    }
    // Respuesta inválida: fail-open como uncertain
    return { isValid: true, review: true, reviewReason: 'LLM_INVALID_DECISION', decision: decision || 'invalid' };
  } catch (e) {
    log(`(Central) Error en validación LLM: ${e.toString()}`);
    // Fail-open, pero marcar como revisión
    return { isValid: true, review: true, reviewReason: 'LLM_ERROR', decision: 'error' };
  }
}

function validateBankTransferRolesWithLLM_(subject, body) {
  const coreBody = extractForwardedOriginalBody_(body || '');
  const truncatedBody = truncateText_(coreBody || body || '', 2200);
  const prompt = BANK_ROLES_PROMPT + `\n\nEmail:\nAsunto: ${subject}\nCuerpo: ${truncatedBody}`;

  try {
    const response = callAI(prompt);
    const cleanResponse = response.replace(/```json\n?|\n?```/g, '').trim();
    const json = JSON.parse(cleanResponse);
    const shouldDiscardRaw = json.should_discard;
    const shouldDiscard = shouldDiscardRaw === true || shouldDiscardRaw === 'true';
    const confidenceRaw = (json.confidence || '').toString().toLowerCase().trim();
    const confidence = (confidenceRaw === 'high' || confidenceRaw === 'medium' || confidenceRaw === 'low') ? confidenceRaw : 'low';
    const reason = (json.reason || '').toString().trim();

    return {
      shouldDiscard: shouldDiscard,
      confidence: confidence,
      reason: reason.length > 80 ? reason.substring(0, 80) : reason,
      payer: (json.payer || '').toString().trim() || null,
      payee: (json.payee || '').toString().trim() || null,
      transferType: (json.transfer_type || '').toString().trim().toLowerCase() || 'unknown'
    };
  } catch (e) {
    log(`(Central) Error en roles LLM (banco): ${e.toString()}`);
    return { shouldDiscard: false, confidence: 'low', reason: 'LLM_ERROR', payer: null, payee: null, transferType: 'unknown' };
  }
}

/**
 * Extrae datos estructurados del email usando IA
 * @param {string} subject - Asunto del email
 * @param {string} body - Cuerpo del email
 * @returns {Object} - {monto, fecha_pago, dpto, estado}
 */
function extractDataWithAI(subject, body) {
  // Limitar para evitar prompts gigantes (p.ej. OCR largo)
  const truncatedBody = truncateText_(body || '', 12000);
  const prompt = `Extrae la siguiente informacion de este email de pago de expensas.
Responde UNICAMENTE con un JSON valido, sin explicaciones ni texto adicional.

Nota: el input incluye EMAIL (asunto/cuerpo) y OCR (texto del comprobante/imagen).
HINTS_JSON y OCR_EVIDENCE son sugerencias/recortes y pueden estar incompletos o equivocados.
Usa el mejor dato disponible por consistencia con el campo pedido.

Campos a extraer:
- monto: numero sin simbolo $ (ejemplo: 164185.30). Si hay múltiples comprobantes, usa monto_total como principal. Si no encuentras monto, usa null.
- montos: array de números (uno por comprobante/transferencia encontrada), sin símbolo $ y usando punto decimal (ej: 65199.02). Si no hay, usa [].
- monto_total: suma de montos (si montos tiene más de 1 valor), mismo formato numérico. Si no aplica, usa null.
- fecha_pago: formato DD-MM-YY. Si no hay fecha de pago explicita en el comprobante, usa null.
- dpto: departamento/unidad/piso. Ejemplos de formatos válidos:
  - "5A", "PB B", "8 PISO", "LOC" (formatos simples)
  - "Piso 5 Dpto 11" (si el texto dice "piso 5 dpto 11", "p 5 dpto 11", "piso 5 departamento 11", etc.)
  - "Cochera 10" o "Cocheras 10 y 15" (si el texto menciona cocheras)
  - Normalizar siempre a "Piso X Dpto Y" cuando haya piso y número de departamento separados.
  - Normalizar siempre a "Cochera X" o "Cocheras X y Y" cuando se mencionen cocheras.
  Si no hay información de departamento/unidad, usa null.
- uf: unidad funcional / unidad (identificador). Regla:
  - Si el texto dice "Unidad:"/"Unidad Funcional:"/"UF:" seguido solo de un número (ej "0035"), eso corresponde a uf.
  - Si dice "Unidad:" con número y letra (ej "02-C", "2C", "2/C"), eso corresponde a dpto (normalizar a "02-C") y uf debe ser null (salvo que haya un uf separado).
  Si no hay UF/Unidad numérica, usa null.
- estado: string en MAYUSCULAS (ej: "PENDIENTE", "APROBADO") si el email contiene un campo tipo "Estado: ...". Si no aparece, usa null. Si aparece "Estado: Pendiente", devuelve "PENDIENTE".
- ed: direccion del edificio/consorcio (solo calle y numero, sin piso ni depto). Debe preservar numeración compuesta tipo "2647/51". Ejemplos: "Paraguay 2949", "Av Santa Fe 2647/51", "Araoz 380". Si no hay dirección, usa null.
- pagador: nombre y apellido o razón social que figure como titular/emisor/ordenante del comprobante (si aparece). Si no hay, usa null.
- observacion_admin: string breve o null. Solo completa si el email contiene un pedido/reclamo/comentario para la administración (ej: no puedo adjuntar comprobante, por favor verificar, no puedo pagar hasta..., reclamo). No repitas montos, DPTO, UF, CBU ni datos sensibles. Si no hay mensaje, usa null.

Input:
${truncatedBody}

JSON:`;

  const response = callAI(prompt);
  // Limpiar respuesta (a veces la IA agrega backticks de markdown)
  const cleanResponse = response.replace(/```json\n?|\n?```/g, '').trim();
  return JSON.parse(cleanResponse);
}

function buildPayerSpecialistPrompt_(subject, emailBody, ocrText) {
  const emailBodyFull = truncateText_((emailBody || '').toString(), 6000);
  const ocrFull = truncateText_((ocrText || '').toString(), 8000);
  const ocrEvidence = buildOcrEvidence_(ocrFull, getOcrAnchors_(), 2, 2500);

  return `Extrae el NOMBRE COMPLETO del PAGADOR (la persona que realiza/ordena la transferencia).
NO es el beneficiario/destinatario/consorcio/administración. Buscamos al ORIGINANTE/ORDENANTE/TITULAR CUENTA DÉBITO.

Si no hay evidencia clara del nombre del pagador, devuelve null.

Devuelve SOLO este JSON válido (sin texto extra):
{"pagador_full_name":string|null,"confidence":"high|medium|low","evidence":string|null}

Reglas:
- "evidence" debe ser una cita EXACTA (substring) del input donde aparece el nombre.
- No inventes nombres.
- No uses nombres de consorcios ("CONS PROP...", "CONSORCIO...", "ADM...") como pagador.

Ejemplos:

INPUT:
EMAIL_SUBJECT: Aviso de transferencia
EMAIL_BODY_FULL:
(vacío)
OCR_FULL:
Comprobante de transferencia
Importe: $ 65.199,02
De: Cecilia Deyheralde
Para: Cons Prop Av Santa Fe 2647 51
Motivo: Varios
OUTPUT:
{"pagador_full_name":"Cecilia Deyheralde","confidence":"high","evidence":"De: Cecilia Deyheralde"}

INPUT:
EMAIL_SUBJECT: Comprobante de Transferencia
EMAIL_BODY_FULL:
(vacío)
OCR_FULL:
BancoCiudad
Importe $ 215.201,03
Destino CONS D PR LM DRAGO 436
Motivo Expensas
Originante CARMEN GRACIELA BURDET
OUTPUT:
{"pagador_full_name":"CARMEN GRACIELA BURDET","confidence":"high","evidence":"Originante CARMEN GRACIELA BURDET"}

INPUT:
EMAIL_SUBJECT: Transferencia
EMAIL_BODY_FULL:
Comparto el comprobante de transferencia de Peron 2250 / 4D
Saludos Cordiales
Silvana
OCR_FULL:
Transferencia
Destinatario: CONS. DE PROP. PTE.
Monto: $70.425,43
Banco: BANCO PATAGONIA
Motivo: Expensas
OUTPUT:
{"pagador_full_name":"Silvana","confidence":"medium","evidence":"Silvana"}

INPUT:
EMAIL_SUBJECT: Comprobante
EMAIL_BODY_FULL:
(vacío)
OCR_FULL:
Cuenta a debitar: CA - PESOS - 4463...
Nombre Beneficiario: CONS PROPIET ARAOZ 378
Importe: 109726.11
OUTPUT:
{"pagador_full_name":null,"confidence":"low","evidence":null}

INPUT:
EMAIL_SUBJECT: ${subject}
EMAIL_BODY_FULL:
${emailBodyFull}

OCR_EVIDENCE:
${ocrEvidence || '(vacío)'}

OCR_FULL:
${ocrFull}
`;
}

function isLikelyPersonName_(name) {
  const s = (name || '').toString().trim();
  if (!s) return false;
  if (/@/.test(s)) return false;
  if (/\d/.test(s)) return false;
  const n = normalizeForMatch_(s);
  if (n.indexOf('consorcio') !== -1) return false;
  if (n.indexOf('cons prop') !== -1) return false;
  if (n.indexOf('cons. de prop') !== -1) return false;
  if (n.indexOf('adm') === 0) return false;
  return /[a-záéíóúñ]/i.test(s);
}

function extractPayerCandidates_(subject, emailBody, ocrText) {
  const candidates = [];
  const seen = {};

  function addCandidate(name, evidence, source) {
    const clean = normalizePayor_(name);
    if (!clean || !isLikelyPersonName_(clean)) return;
    const key = normalizeForMatch_(clean);
    if (!key || seen[key]) return;
    seen[key] = true;
    candidates.push({ name: clean, evidence: (evidence || '').toString().trim(), source: source || 'unknown' });
  }

  const combined = `${subject || ''}\n${emailBody || ''}\n${ocrText || ''}`;
  const lines = combined.split(/\r?\n/).map(function(l) { return (l || '').trim(); }).filter(Boolean);

  // OCR/email labels típicos: "De:", "Originante", "Ordenante", "Titular", "Titular cuenta débito"
  const labelPatterns = [
    /\bDe\s*:\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s.'-]{2,80})/i,
    /\bOriginante\b[:\s]+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s.'-]{2,80})/i,
    /\bOrdenante\b[:\s]+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s.'-]{2,80})/i,
    /\bTitular(?:\s+cuenta\s+d[eé]bito)?\b[:\s]+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s.'-]{2,80})/i,
    /\bCuenta\s+d[eé]bito\b[:\s]+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s.'-]{2,80})/i
  ];

  for (var i = 0; i < lines.length; i++) {
    const line = lines[i];
    for (var j = 0; j < labelPatterns.length; j++) {
      const m = labelPatterns[j].exec(line);
      if (m && m[1]) addCandidate(m[1], line, 'labeled');
    }
  }

  // Firma simple en email: si termina con "Saludos" o "Saludos Cordiales", tomar la siguiente línea como candidato
  for (var k = 0; k < lines.length - 1; k++) {
    const ln = normalizeForMatch_(lines[k]);
    if (ln === 'saludos' || ln === 'saludos cordiales' || ln === 'cordialmente') {
      const next = lines[k + 1];
      // Evitar frases largas y palabras típicas de firmas
      if (next && next.length <= 40 && !/administraci|consorcio|cons prop|banco|cbu|cuit|cuil/i.test(next)) {
        addCandidate(next, next, 'email_signature');
      }
    }
  }

  return candidates;
}

function buildPayerSelectorPrompt_(subject, emailBody, ocrText, candidates) {
  const emailBodyFull = truncateText_((emailBody || '').toString(), 4000);
  const ocrFull = truncateText_((ocrText || '').toString(), 5000);
  const ocrEvidence = buildOcrEvidence_(ocrFull, getOcrAnchors_(), 2, 2000);
  const candJson = JSON.stringify((candidates || []).map(function(c, idx) {
    return { index: idx, name: c.name, source: c.source, evidence: c.evidence };
  }));

  return `Elegí el PAGADOR correcto (persona que realiza/ordena la transferencia) entre los CANDIDATOS.
NO es el beneficiario/destinatario/consorcio/administración. Si ninguno es claramente el pagador, devolvé null.

Devuelve SOLO este JSON válido:
{"selected_index":number|null,"confidence":"high|medium|low","reason":"texto breve"}

Reglas:
- selected_index debe ser un número entero válido del listado o null.
- Preferí candidatos con evidencia rotulada (De/Originante/Ordenante/Titular cuenta débito).
- No elijas nombres de consorcios/administración.

Ejemplo 1:
CANDIDATOS:
[{"index":0,"name":"Cecilia Deyheralde","source":"labeled","evidence":"De: Cecilia Deyheralde"},{"index":1,"name":"Cons Prop Av Santa Fe 2647 51","source":"labeled","evidence":"Para: Cons Prop Av Santa Fe 2647 51"}]
OUTPUT:
{"selected_index":0,"confidence":"high","reason":"Está rotulado como De/originante"}

Ejemplo 2:
CANDIDATOS:
[{"index":0,"name":"CONS PROPIET ARAOZ 378","source":"labeled","evidence":"Nombre Beneficiario: CONS PROPIET ARAOZ 378"}]
OUTPUT:
{"selected_index":null,"confidence":"low","reason":"Solo aparece beneficiario/consorcio"}

INPUT:
EMAIL_SUBJECT: ${subject}
EMAIL_BODY_FULL:
${emailBodyFull}

OCR_EVIDENCE:
${ocrEvidence || '(vacío)'}

OCR_FULL:
${ocrFull}

CANDIDATOS:
${candJson}
`;
}

function selectPayerFromCandidatesWithLLM_(subject, emailBody, ocrText, candidates) {
  if (!candidates || candidates.length === 0) return null;
  if (candidates.length === 1) return candidates[0].name;
  const prompt = buildPayerSelectorPrompt_(subject, emailBody, ocrText, candidates);
  const response = callAI(prompt);
  const cleanResponse = response.replace(/```json\n?|\n?```/g, '').trim();
  const json = JSON.parse(cleanResponse);
  if (json.selected_index == null || json.selected_index === '') return null;
  const idx = Number(json.selected_index);
  if (isNaN(idx) || idx < 0 || idx >= candidates.length) return null;
  return candidates[idx].name;
}

function extractPayerWithLLMSpecialist_(subject, emailBody, ocrText) {
  const prompt = buildPayerSpecialistPrompt_(subject, emailBody, ocrText);
  const response = callAI(prompt);
  const cleanResponse = response.replace(/```json\n?|\n?```/g, '').trim();
  const json = JSON.parse(cleanResponse);
  const payer = normalizePayor_(json.pagador_full_name);
  const confidence = (json.confidence || '').toString().toLowerCase().trim();
  const evidence = (json.evidence || '').toString();

  const inputForEvidence = `${subject}\n${emailBody || ''}\n${ocrText || ''}`;
  if (!payer || !isLikelyPersonName_(payer)) return { payer: null, confidence: confidence || 'low', evidence: null };
  if (evidence && inputForEvidence.indexOf(evidence) === -1) {
    return { payer: null, confidence: 'low', evidence: null };
  }
  return { payer: payer, confidence: confidence || 'medium', evidence: evidence || null };
}

function extractPayerWithCandidateSelector_(subject, emailBody, ocrText) {
  const candidates = extractPayerCandidates_(subject, emailBody, ocrText);
  const selected = selectPayerFromCandidatesWithLLM_(subject, emailBody, ocrText, candidates);
  if (selected) return selected;
  // Fallback: especialista "libre" con evidencia (por si no hubo candidatos claros)
  const res = extractPayerWithLLMSpecialist_(subject, emailBody, ocrText);
  return res && res.payer ? res.payer : null;
}

/**
 * Guarda los datos extraídos en el Google Sheet
 * @param {string} noticeDate - Fecha de recepción del email (DD-MM-YY)
 * @param {string|null} paymentDate - Fecha de pago del comprobante
 * @param {number|null} amount - Monto del pago
 * @param {string|null} apartment - Departamento/unidad
 * @param {string} threadUrl - URL del thread en Gmail
 * @param {string} label - Texto visible del hyperlink
 */
function buildObservacionesRichText_(threadUrl, label, highlightText, highlightColor) {
  const linkLabel = (label || 'Abrir email').toString();
  const prefix = highlightText ? (highlightText + ' | ') : '';
  const full = prefix + linkLabel;

  const builder = SpreadsheetApp.newRichTextValue().setText(full);
  builder.setLinkUrl(prefix.length, full.length, threadUrl);

  if (highlightText) {
    const style = SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setForegroundColor(highlightColor || '#3D85C6')
      .build();
    builder.setTextStyle(0, highlightText.length, style);
  }

  return builder.build();
}

function saveToSheet(noticeDate, paymentDate, amount, building, apartment, uf, highlightStatus, threadUrl, label) {
  const ss = SpreadsheetApp.openById(SHEET_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`No se encontró la hoja "${SHEET_CONFIG.SHEET_NAME}" en el spreadsheet`);
  }
  log(`(Central) CREANDO OBSERVACIONES con RichText`);
  
  // Columnas: TIPO AVISO | FECHA AVISO | FECHA DE PAGO | MONTO | ED | DPTO | UF | COMENTARIO
  const row = [
    '@',                    // TIPO AVISO (siempre email)
    noticeDate,             // FECHA AVISO
    paymentDate || '',      // FECHA DE PAGO
    amount || '',           // MONTO
    building || '',         // ED
    apartment || '',        // DPTO
    uf || '',               // UF
    ''                      // COMENTARIO / OBSERVACIONES (se setea como RichText)
  ];
  
  sheet.appendRow(row);
  log(`(Central) Fila agregada al Sheet: ${JSON.stringify(row)}`);

  const lastRow = sheet.getLastRow();
  const rowRange = sheet.getRange(lastRow, 1, 1, row.length);
  const commentCell = sheet.getRange(lastRow, 8);

  let highlightText = '';
  let highlightColor = '';
  if (highlightStatus === 'missing_no_attachment') {
    highlightText = 'NO SE DETECTO COMPROBANTE DE PAGO';
    highlightColor = '#B45F06'; // amarillo oscuro legible
    rowRange.setBackground('#FFF2CC'); // amarillo suave
  } else if (highlightStatus === 'missing_with_attachment') {
    highlightText = 'NO SE DETECTO COMPROBANTE DE PAGO | SE DETECTO UN ARCHIVO';
    highlightColor = '#3D85C6'; // celeste
    rowRange.setBackground('#D9EAF7'); // celeste suave
  } else if (highlightStatus === 'ok') {
    // sin resaltado
  }

  commentCell.setRichTextValue(buildObservacionesRichText_(threadUrl, label, highlightText, highlightColor));
}

// ==================== FUNCIONES PRINCIPALES ====================

function procesarEmailsExpensas() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) {
    log('=== SKIP: otro proceso está corriendo ===');
    return;
  }
  OCR_CALLS_RUN = 0;
  log('=== Iniciando procesamiento de emails (casilla central) ===');
  let totalProcesados = 0;
  let totalMarcados = 0;
  let totalErrores = 0;
  let needRerun = false;
  let rerunDelayMs = 0;

  try {
    for (var i = 0; i < CONFIG.EMAILS_ORIGEN.length; i++) {
      const emailOrigen = CONFIG.EMAILS_ORIGEN[i];
      try {
        const resultado = procesarEmailsDeCuenta(emailOrigen);
        totalProcesados += resultado.procesados;
        totalMarcados += resultado.reenviados; // aquí reenviados = marcados
        totalErrores += resultado.errores;
        if (resultado.needRerun) {
          needRerun = true;
          rerunDelayMs = Math.max(rerunDelayMs, resultado.rerunDelayMs || 0);
        }
        if (resultado.stopRun) {
          log('=== STOP: se retoma en próxima corrida ===');
          break;
        }
      } catch (e) {
        log(`Error procesando ${emailOrigen}: ${e.toString()}`);
        totalErrores++;
      }
    }
    log(`=== Resumen central: ${totalProcesados} procesados, ${totalMarcados} marcados, ${totalErrores} errores ===`);
    if (needRerun) {
      scheduleRerun_(rerunDelayMs || 2 * 60 * 1000);
    }
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function procesarEmailsDeCuenta(emailOrigen) {
  const stats = { procesados: 0, reenviados: 0, errores: 0, stopRun: false, needRerun: false, rerunDelayMs: 0 };
  try {
    const labelProcesadoName = getRequiredConfigString_(CONFIG.ETIQUETA_PROCESADO, 'ETIQUETA_PROCESADO');
    const labelDescartadoName = (CONFIG.ETIQUETA_DESCARTADO || 'ExpensaDescartada').toString().trim();
    const labelEnProcesoName = (CONFIG.ETIQUETA_EN_PROCESO || 'ExpensaEnProceso').toString().trim();
    const labelRequiereRevisionName = (CONFIG.ETIQUETA_REQUIERE_REVISION || 'REQUIERE REVISION').toString().trim();
    const monthQuery = getMonthStartQuery_();

    cleanupDiscardStateIfNeeded_();
    const etiquetaDescartado = crearObtenerEtiqueta(labelDescartadoName);
    reabrirDescartadosConAdjunto_(etiquetaDescartado, labelDescartadoName, CONFIG.MAX_THREADS_PER_RUN);

    const queryEnProceso = `in:anywhere ${monthQuery} label:${escapeLabelForQuery_(labelEnProcesoName)} -label:${escapeLabelForQuery_(labelProcesadoName)} -label:${escapeLabelForQuery_(labelDescartadoName)} -label:${escapeLabelForQuery_(labelRequiereRevisionName)}`;
    const threadsEnProceso = GmailApp.search(queryEnProceso, 0, CONFIG.MAX_THREADS_PER_RUN);

    const queryNuevos = `in:inbox ${monthQuery} -label:${escapeLabelForQuery_(labelProcesadoName)} -label:${escapeLabelForQuery_(labelDescartadoName)} -label:${escapeLabelForQuery_(labelEnProcesoName)} -label:${escapeLabelForQuery_(labelRequiereRevisionName)}`;
    const threadsNuevos = GmailApp.search(queryNuevos, 0, CONFIG.MAX_THREADS_PER_RUN);
    const threadsNuevosPlusOne = GmailApp.search(queryNuevos, 0, CONFIG.MAX_THREADS_PER_RUN + 1);
    const hayMasNuevos = threadsNuevosPlusOne.length > threadsNuevos.length;

    const byId = {};
    const threads = [];
    for (var i = 0; i < threadsEnProceso.length; i++) {
      const t1 = threadsEnProceso[i];
      const id1 = t1.getId();
      if (byId[id1]) continue;
      byId[id1] = true;
      threads.push(t1);
    }
    for (var j = 0; j < threadsNuevos.length; j++) {
      const t2 = threadsNuevos[j];
      const id2 = t2.getId();
      if (byId[id2]) continue;
      byId[id2] = true;
      threads.push(t2);
    }

    const batch = threads.slice(0, CONFIG.MAX_THREADS_PER_RUN);
    log(`(Central) Threads en proceso: ${threadsEnProceso.length}, nuevos: ${threadsNuevos.length}, procesando: ${batch.length}`);
    if (threadsEnProceso.length > 0 || hayMasNuevos) {
      stats.needRerun = true;
      stats.rerunDelayMs = Math.max(stats.rerunDelayMs, 2 * 60 * 1000);
    }

    const etiqueta = crearObtenerEtiqueta(labelProcesadoName);
    const etiquetaEnProceso = crearObtenerEtiqueta(labelEnProcesoName);
    const etiquetaRequiereRevision = crearObtenerEtiqueta(labelRequiereRevisionName);

    const messagesByThread = GmailApp.getMessagesForThreads(batch);
    let totalMessagesRun = 0;
    for (var ti = 0; ti < batch.length; ti++) {
      const thread = batch[ti];
      const threadId = thread.getId();
      const threadUrl = `https://mail.google.com/mail/u/0/#all/${threadId}`;
      try {
        if (totalMessagesRun >= CONFIG.MAX_TOTAL_MESSAGES_PER_RUN) {
          log('(Central) STOP: MAX_TOTAL_MESSAGES_PER_RUN alcanzado');
          return stats;
        }

        // Lease + EnProceso (checkpoint)
        const isStale = isThreadLeaseStale_(threadId, OCR_CONFIG.STALE_EN_PROCESO_HOURS);
        if (isStale) {
          log(`(Central) Lease stale/empty, retomando: ${threadId}`);
        }
        thread.addLabel(etiquetaEnProceso);
        setThreadLease_(threadId, new Date());

        const messages = messagesByThread[ti] || [];
        const maxMessages = Math.min(messages.length, CONFIG.MAX_MESSAGES_PER_THREAD);
        const processedAllMessages = messages.length <= maxMessages;
        let threadFinalized = false;

        for (var mi = 0; mi < maxMessages; mi++) {
          const message = messages[mi];
          stats.procesados++;
          totalMessagesRun++;
        if (totalMessagesRun > CONFIG.MAX_TOTAL_MESSAGES_PER_RUN) {
          log('(Central) STOP: MAX_TOTAL_MESSAGES_PER_RUN alcanzado');
          stats.needRerun = true;
          stats.rerunDelayMs = Math.max(stats.rerunDelayMs, 2 * 60 * 1000);
          return stats;
        }

          const subjectPre = message.getSubject() || '';
          if (REGEX_RESUMEN_PROCESAMIENTO.test(subjectPre.trim())) {
            log(`(Central) SKIP resumen: ${subjectPre}`);
            thread.addLabel(etiquetaDescartado);
            setDiscardLastMsgMs_(threadId, thread.getLastMessageDate());
            thread.removeLabel(etiquetaEnProceso);
            clearThreadLease_(threadId);
            threadFinalized = true;
            break;
          }
          if (esComprobanteExpensa(message)) {
            log(`(Central) ✓ Expensa detectada por keywords: ${message.getSubject()}`);
            
            const subject = message.getSubject();
            const body = message.getPlainBody();
            const hasInlineImagesOrAttachments = message.getAttachments({ includeInlineImages: true }).length > 0;

            // Señal simple de "pago claro" (para debugging)
            const texto = `${subject} ${body}`.toLowerCase();
            const pagoClaro = /transferenc|recibiste un pago|pago fue exitoso|se acredit|realizaste/.test(texto) &&
              (/(?:\$\s*)?\d{1,3}(?:[\.\s]\d{3})+(?:,\d{2})?|\$\s*\d{4,}/.test(texto) || /\b(cbu|cvu|alias|nro|número)\b/.test(texto));
            
            // Regla: ignorar transferencias salientes realizadas por CARLOS ARTUSO.
            // Importante: NO descartar transferencias entrantes aunque aparezca su nombre.
            if (shouldDiscardOutgoingCarlosArtuso_(subject, body)) {
              log(`(Central) DESCARTADO: transferencia saliente de Carlos Artuso | ${subject}`);
              thread.addLabel(etiquetaDescartado);
              setDiscardLastMsgMs_(threadId, thread.getLastMessageDate());
              thread.removeLabel(etiquetaEnProceso);
              clearThreadLease_(threadId);
              threadFinalized = true;
              break;
            }

            // PASO 0: OCR (solo si parece necesario)
            // - Adjuntos, cuerpo vacío/escaso, o el texto sugiere que "va adjunto".
            let ocrText = '';
            const bodyLooksEmpty = !body || body.trim().length < 120;
            const textSuggestsAttachment = /\badjunt|ver\s+archivo|comprobante\s+adjunto|imagen\s+adjunta/i.test(body || '') ||
              /\badjunt|ver\s+archivo|comprobante\s+adjunto|imagen\s+adjunta/i.test(subject || '');

            // Si hay adjuntos/inline, necesitamos OCR para distinguir "comprobante" vs "archivo adjunto".
            if (OCR_CONFIG.ENABLED && (hasInlineImagesOrAttachments || bodyLooksEmpty || textSuggestsAttachment)) {
              const ocrAdj = extraerTextoDeAdjuntosExpensa(message);
              const ocrInline = extraerTextoDeImagenesEnCuerpo(message);
              ocrText = [ocrAdj, ocrInline].filter(Boolean).join('\n\n');
              if (ocrText) {
                log(`(Central) OCR agregado (chars=${ocrText.length})`);
              }
            }
            const receiptBlocks = getReceiptLikeBlocks_(ocrText);
            const hasComprobante = receiptBlocks.length > 0 || (hasInlineImagesOrAttachments && detectReceiptLikeOcr_(ocrText));
            const bodyForAI = composeBodyWithOcr_(body, ocrText);
            const fullTextForRules = `${subject}\n${bodyForAI}`;
            // Estado único (opción 2): OK vs falta comprobante (con o sin archivo)
            const highlightStatus = hasComprobante
              ? 'ok'
              : (hasInlineImagesOrAttachments ? 'missing_with_attachment' : 'missing_no_attachment');

            // Workflow:
            // 1) Si hay >=2 comprobantes/imagenes OCR -> revisión humana (política cauta)
            const receiptMarkersEarly = receiptBlocks.length || countReceiptMarkers_(ocrText);
            if ((receiptBlocks.length >= 2) || receiptMarkersEarly >= 2) {
              const nBlocks = receiptBlocks.length > 0 ? receiptBlocks.length : receiptMarkersEarly;
              log(`(Central) REQUIERE REVISION: múltiples comprobantes (n=${nBlocks}) | ${subject}`);
              const etiquetaMultiples = crearObtenerEtiqueta(CONFIG.ETIQUETA_MULTIPLES_COMPROBANTES);
              thread.addLabel(etiquetaMultiples);
              thread.addLabel(etiquetaRequiereRevision);
              clearDiscardLastMsgMs_(threadId);
              thread.removeLabel(etiquetaEnProceso);
              clearThreadLease_(threadId);
              threadFinalized = true;
              break;
            }

            // 2) Si el OCR parece comprobante, validar destinatario (fuzzy consorcio/Artuso)
            let acceptByDestino = false;
            var ocrTextForParse = ocrText;
            if (receiptBlocks.length === 1) ocrTextForParse = receiptBlocks[0].text;

            if (hasComprobante) {
              const titularesOcr = extractTitularesFromOcr_(ocrTextForParse || '');
              const titularDestinoOcr = (titularesOcr.destino || '').toString().trim();
              if (!titularDestinoOcr) {
                log(`(Central) REQUIERE REVISION: destino no detectado en OCR | ${subject}`);
                thread.addLabel(etiquetaRequiereRevision);
                clearDiscardLastMsgMs_(threadId);
                thread.removeLabel(etiquetaEnProceso);
                clearThreadLease_(threadId);
                threadFinalized = true;
                break;
              }
              const destinoOk = isConsorcioLike_(titularDestinoOcr) || isArtusoLike_(titularDestinoOcr);
              if (!destinoOk) {
                log(`(Central) DESCARTADO: destino no-consorcio/no-Artuso (${titularDestinoOcr}) | ${subject}`);
                thread.addLabel(etiquetaDescartado);
                setDiscardLastMsgMs_(threadId, thread.getLastMessageDate());
                thread.removeLabel(etiquetaEnProceso);
                clearThreadLease_(threadId);
                threadFinalized = true;
                break;
              }
              acceptByDestino = true;
              log(`(Central) ✓ Expensa confirmada por destinatario OCR: ${titularDestinoOcr}`);
            }

            // 3) Exclusions + LLM solo si NO se confirmó por destinatario OCR
            let validation = { isValid: true, review: false, decision: 'accept', reviewReason: '' };
            if (!acceptByDestino) {
              // Excluir por texto completo (incluye OCR) para frenar falsos positivos de adjuntos
              const textoCompletoLower = fullTextForRules.toLowerCase();
              const tieneExclusionFuerte = PALABRAS_EXCLUSION_FUERTE.some(p => textoCompletoLower.indexOf(p) !== -1);
              if (tieneExclusionFuerte) {
                log(`(Central) DESCARTADO por exclusión fuerte: ${subject}`);
                thread.addLabel(etiquetaDescartado);
                setDiscardLastMsgMs_(threadId, thread.getLastMessageDate());
                thread.removeLabel(etiquetaEnProceso);
                clearThreadLease_(threadId);
                threadFinalized = true;
                break;
              }

              // Regla determinística: transferencias con Referencia/Motivo/Concepto no-expensa (ej: SAC)
              const nonExpenseConcept = matchesNonExpenseConcept_(fullTextForRules);
              if (nonExpenseConcept) {
                log(`(Central) DESCARTADO por concepto no-expensa: ${nonExpenseConcept} | ${subject}`);
                thread.addLabel(etiquetaDescartado);
                setDiscardLastMsgMs_(threadId, thread.getLastMessageDate());
                thread.removeLabel(etiquetaEnProceso);
                clearThreadLease_(threadId);
                threadFinalized = true;
                break;
              }

              const isBankNotice = isBankNotification_(message, subject, body);
              if (isBankNotice) {
                const transferType = matchesProveedorTransferType_(fullTextForRules);
                if (transferType) {
                  log(`(Central) DESCARTADO: tipo de transferencia proveedor (${transferType}) | ${subject}`);
                  thread.addLabel(etiquetaDescartado);
                  setDiscardLastMsgMs_(threadId, thread.getLastMessageDate());
                  thread.removeLabel(etiquetaEnProceso);
                  clearThreadLease_(threadId);
                  threadFinalized = true;
                  break;
                }

                const bankRoles = validateBankTransferRolesWithLLM_(subject, body);
                if (bankRoles.shouldDiscard && (bankRoles.confidence === 'high' || bankRoles.confidence === 'medium')) {
                  const reason = bankRoles.reason ? ` | ${bankRoles.reason}` : '';
                  log(`(Central) DESCARTADO: transferencia proveedor (LLM ${bankRoles.confidence})${reason} | ${subject}`);
                  thread.addLabel(etiquetaDescartado);
                  setDiscardLastMsgMs_(threadId, thread.getLastMessageDate());
                  thread.removeLabel(etiquetaEnProceso);
                  clearThreadLease_(threadId);
                  threadFinalized = true;
                  break;
                }
              }

              // PASO 1: Validar con LLM (fail-open con 3 estados)
              validation = validateExpensePaymentWithLLM(subject, bodyForAI);
              log(`(Central) Validator: decision=${validation.decision || 'unknown'} review=${validation.review ? 'yes' : 'no'} reason=${validation.reviewReason || ''} pago_claro=${pagoClaro ? 'yes' : 'no'}`);
              if (validation.review) {
                log(`(Central) REVIEW (intención no clara): ${validation.reviewReason || 'LLM_UNCERTAIN'}`);
              }
              
              if (!validation.isValid) {
                log(`(Central) ✗ LLM rejected: ${subject} | Reason: ${validation.reviewReason || 'LLM_REJECT'}`);
                // No etiquetar como expensa - evitar re-procesamiento
                thread.addLabel(etiquetaDescartado);
                setDiscardLastMsgMs_(threadId, thread.getLastMessageDate());
                thread.removeLabel(etiquetaEnProceso);
                clearThreadLease_(threadId);
                threadFinalized = true;
                break;
              }
            }
            // Nota: si el LLM está en "review", dejamos pasar y lo marcamos en observaciones.
            
            // PASO 2: Extraer datos con IA y guardar en Sheet
            try {
              const noticeDate = Utilities.formatDate(
                message.getDate(), 
                'America/Argentina/Buenos_Aires', 
                'dd-MM-yy'
              );
              
              const ocrOnly = getOcrSection_(bodyForAI);
              const extractorContext = buildExtractorContext_(subject, body, ocrText);
              const extractedData = extractDataWithAI(subject, extractorContext);
              const estado = (extractedData.estado || '').toString().trim().toUpperCase();

              const qaTags = [];
              const currentEd2 = (extractedData.ed || '').toString();
              const llmDpto = (extractedData.dpto || '').toString().trim();
              const llmUf = (extractedData.uf || '').toString().trim();
              const llmDptoNorm = normalizeDpto_(llmDpto);

              // Determinar si hay múltiples recibos (marcadores de OCR de adjuntos/inline)
              const receiptMarkers = receiptBlocks.length || countReceiptMarkers_(ocrText);
              const multipleReceipts = receiptBlocks.length >= 2 || receiptMarkers >= 2;
              if (multipleReceipts) qaTags.push('QA_MULTIPLE_RECEIPTS');
              if (multipleReceipts) {
                const etiquetaMultiples = crearObtenerEtiqueta(CONFIG.ETIQUETA_MULTIPLES_COMPROBANTES);
                thread.addLabel(etiquetaMultiples);
              }

              // Deduplicación de montos por fuente solo cuando parece un único comprobante.
              if (!multipleReceipts) {
                const montosFromOcr = ocrTextForParse ? (extractMontosFromOcr_(ocrTextForParse).montos || []) : [];
                const montosFromBody = extractMontosFromTextLoose_(body);
                const montosFromSubject = extractMontosFromTextLoose_(subject);
                const montosFromLlm = (extractedData.montos || []).map(parseAmount_).filter(function(v) { return v != null; });

                const merged = [];
                const seenMonto = {};
                let dedupAcrossSources = false;

                function addMonto_(value, source) {
                  const n = parseAmount_(value);
                  if (n == null) return;
                  const key = String(n);
                  if (seenMonto[key]) {
                    dedupAcrossSources = true;
                    return;
                  }
                  seenMonto[key] = source || 'unknown';
                  merged.push(n);
                }

                montosFromOcr.forEach(function(v) { addMonto_(v, 'ocr'); });
                montosFromBody.forEach(function(v) { addMonto_(v, 'body'); });
                montosFromSubject.forEach(function(v) { addMonto_(v, 'subject'); });
                montosFromLlm.forEach(function(v) { addMonto_(v, 'llm'); });

                if (merged.length > 0) {
                  extractedData.montos = merged;
                  extractedData.monto = merged[0];
                  if (merged.length === 1) extractedData.monto_total = merged[0];
                }
                if (dedupAcrossSources) qaTags.push('QA_DEDUP_MONTO_BODY');
              }

              // Auditoría (no corrige): sugerencias desde email + OCR
              const emailTextForExtract = `${subject}\n${body}`;
              const emailEd = extractEdFromEmailText_(emailTextForExtract);
              const emailDptoRaw = extractDptoFromEmailText_(emailTextForExtract);
              const emailDptoNorm = normalizeDpto_(emailDptoRaw);
              if (emailEd && normalizeForMatch_(emailEd) !== normalizeForMatch_(currentEd2)) {
                qaTags.push(`SUG_ED_EMAIL=${truncateText_(emailEd, 45)}`);
              }
              if (emailDptoNorm && normalizeForMatch_(emailDptoNorm) !== normalizeForMatch_(llmDptoNorm || '')) {
                qaTags.push(`SUG_DPTO_EMAIL=${truncateText_(emailDptoNorm, 20)}`);
              }

              if (isIgnoredEd_(currentEd2)) qaTags.push('WARN_ED_ADMIN_ADDR');

              // Direcciones con "/" en el texto: sugerir si el LLM no preservó
              const slashAddress = extractSlashAddress_(fullTextForRules);
              if (slashAddress) {
                const hasSlash = /\d+\s*\/\s*\d+/.test(currentEd2);
                if (!hasSlash) qaTags.push(`SUG_ED_SLASH=${truncateText_(slashAddress, 45)}`);
              }

              let dptoFromOcrRaw = null;
              let dptoFromOcrNorm = null;
              let ufFromOcr = null;
              if (ocrOnly) {
                const edFromOcr = extractEdFromOcr_(ocrOnly);
                if (edFromOcr && normalizeForMatch_(edFromOcr) !== normalizeForMatch_(currentEd2)) {
                  qaTags.push(`SUG_ED_OCR=${truncateText_(edFromOcr, 45)}`);
                }
                dptoFromOcrRaw = extractDeptoFromOcr_(ocrOnly) || extractDptoFromUnidadLike_(ocrOnly) || extractDeptoFromObservacionesOcr_(ocrOnly);
                dptoFromOcrNorm = normalizeDpto_(dptoFromOcrRaw);
                if (dptoFromOcrNorm && normalizeForMatch_(dptoFromOcrNorm) !== normalizeForMatch_(llmDptoNorm || '')) {
                  qaTags.push(`SUG_DPTO_OCR=${truncateText_(dptoFromOcrNorm, 20)}`);
                  if (llmDptoNorm) qaTags.push(`WARN_DPTO_LLM=${truncateText_(llmDptoNorm, 20)}`);
                }

                ufFromOcr = extractUfFromOcr_(ocrOnly);
                if (ufFromOcr) {
                  if (!llmUf) {
                    qaTags.push(`SUG_UF_OCR=${ufFromOcr}`);
                  } else if (normalizeForMatch_(llmUf) !== normalizeForMatch_(ufFromOcr)) {
                    qaTags.push(`SUG_UF_OCR=${ufFromOcr}`);
                    qaTags.push(`WARN_UF_LLM=${truncateText_(llmUf, 20)}`);
                  }
                }
              }

              const dptoEmailEvidence = emailDptoRaw ? findEvidenceLine_(emailTextForExtract, emailDptoRaw) : null;
              const dptoOcrEvidence = dptoFromOcrRaw ? findEvidenceLine_(ocrOnly, dptoFromOcrRaw) : null;
              const ufOcrEvidence = ufFromOcr ? findEvidenceLine_(ocrOnly, ufFromOcr) : null;

              const unitCandidates = collectUnitCandidates_({
                dptoFromLlm: llmDpto,
                ufFromLlm: llmUf,
                dptoFromEmail: emailDptoNorm,
                dptoFromOcr: dptoFromOcrNorm,
                ufFromOcr: ufFromOcr,
                dptoEmailEvidence: dptoEmailEvidence,
                dptoOcrEvidence: dptoOcrEvidence,
                ufOcrEvidence: ufOcrEvidence
              });
              const hasDptoConflict = unitCandidates.dpto.length > 1;
              const hasUfConflict = unitCandidates.uf.length > 1;
              const unitResolution = resolveUnitsWithReconciliation_(subject, body, ocrOnly || '', unitCandidates);
              const unitReconLow = unitResolution.lowConfidence;

              if (unitCandidates.dpto.length > 1) qaTags.push('QA_DPTO_MULTI');

              if (unitReconLow) {
                extractedData.dpto = null;
                extractedData.uf = null;
                qaTags.push('QA_UNIT_RECON_LOW');
              } else {
                if (hasDptoConflict) {
                  extractedData.dpto = unitResolution.dpto || null;
                  if (unitResolution.dpto) {
                    if (!llmDptoNorm || normalizeForMatch_(unitResolution.dpto) !== normalizeForMatch_(llmDptoNorm)) {
                      qaTags.push('FIX_DPTO_LLM_RECON');
                    }
                  }
                } else if (unitResolution.dptoSource) {
                  extractedData.dpto = unitResolution.dpto;
                  if (unitResolution.dptoSource === 'email') qaTags.push('FIX_DPTO_EMAIL');
                  if (unitResolution.dptoSource === 'ocr') qaTags.push('FIX_DPTO_OCR');
                }

                if (hasUfConflict) {
                  extractedData.uf = unitResolution.uf || null;
                  if (unitResolution.uf) {
                    if (!llmUf || normalizeForMatch_(unitResolution.uf) !== normalizeForMatch_(llmUf)) {
                      qaTags.push('FIX_UF_LLM_RECON');
                    }
                  }
                } else if (unitResolution.ufSource) {
                  extractedData.uf = unitResolution.uf;
                  if (unitResolution.ufSource === 'ocr') qaTags.push('FIX_UF_OCR');
                }
              }

              if (!unitReconLow) {
                // Si el LLM puso UF numérica en DPTO, mover a UF (alta confianza)
                const dptoNumericOnly = /^[0-9]{2,6}$/.test((extractedData.dpto || '').toString().trim());
                if (dptoNumericOnly && !extractedData.uf) {
                  extractedData.uf = (extractedData.dpto || '').toString().trim();
                  extractedData.dpto = null;
                  qaTags.push('FIX_DPTO_TO_UF');
                }

                // Cocheras detectadas en el texto completo (email+OCR)
                const cocheraCandidates = [];
                const cocheraFromText = extractCocheraDptoFromText_(fullTextForRules);
                if (cocheraFromText) cocheraCandidates.push({ value: cocheraFromText });

                const MAX_COCHERAS = 4;
                function dedupCandidates_(arr) {
                  const out = [];
                  const seen = {};
                  for (var i = 0; i < arr.length; i++) {
                    const v = (arr[i].value || '').toString().trim();
                    const key = normalizeForMatch_(v);
                    if (!v || !key || seen[key]) continue;
                    seen[key] = true;
                    out.push(arr[i]);
                    if (out.length >= MAX_COCHERAS) break;
                  }
                  return out;
                }

                const cocherasFinal = dedupCandidates_(cocheraCandidates);
                if (cocherasFinal.length > 0 && !extractedData.dpto) qaTags.push('QA_DPTO_COCHERA_ONLY');
                if (cocherasFinal.length > 0 && extractedData.dpto) qaTags.push('QA_DPTO_COCHERA_MERGE');

                const partsDpto = [];
                if (extractedData.dpto) partsDpto.push(extractedData.dpto);
                cocherasFinal.forEach(function(c) { partsDpto.push(c.value); });
                if (partsDpto.length) extractedData.dpto = partsDpto.join(', ');
              }

              const currentDptoNorm = normalizeDpto_(extractedData.dpto);

              // Fallback determinístico de fecha de pago desde OCR si el LLM no la devolvió.
              if (!extractedData.fecha_pago && ocrText) {
                const fechaFromOcr = extractFechaFromOcr_(ocrText);
                if (fechaFromOcr) {
                  extractedData.fecha_pago = fechaFromOcr;
                  qaTags.push('FIX_FECHA_OCR');
                }
              }

              // Control: fecha_pago debe ser <= fecha_aviso (fecha de recepción del mail).
              const noticeParts = parseDateFromText_(noticeDate);
              const pagoParts = parseDateFromText_(extractedData.fecha_pago);
              const noticeYmd = ymdInt_(noticeParts);
              const pagoYmd = ymdInt_(pagoParts);
              if (noticeYmd != null && pagoYmd != null && pagoYmd > noticeYmd) {
                extractedData.fecha_pago = noticeDate;
                qaTags.push('FIX_FECHA_PAGO_GT_AVISO');
              }

              // Múltiples comprobantes: sumar montos si hay array
              let amountToSave = null;
              // Prioridad: OCR > LLM. Si hay montos OCR, usarlos primero.
              if (ocrText) {
                try {
                  const ocrAmounts = extractMontosFromOcr_(ocrText);
                  if (ocrAmounts && ocrAmounts.monto_total != null) {
                    amountToSave = ocrAmounts.monto_total;
                    qaTags.push('FIX_MONTO_OCR');
                  } else if (ocrAmounts && ocrAmounts.montos && ocrAmounts.montos.length) {
                    const maxAmt = Math.max.apply(null, ocrAmounts.montos);
                    if (isFinite(maxAmt)) {
                      amountToSave = maxAmt;
                      qaTags.push('FIX_MONTO_OCR');
                    }
                  }
                } catch (eAmt) {
                  log(`(Central) Error monto OCR: ${eAmt.toString()}`);
                }
              }

              // Si no hay OCR o no trajo valor, recurrir al LLM
              if (amountToSave == null) {
                amountToSave = parseAmount_(extractedData.monto);
              }
              const montos = extractedData.montos;
              const parsedMontoTotal = parseAmount_(extractedData.monto_total);
              if (parsedMontoTotal != null) {
                amountToSave = parsedMontoTotal;
              } else if (montos && montos.length && montos.length > 1) {
                const sum = montos.reduce(function(acc, v) {
                  const n = parseAmount_(v);
                  return acc + (n == null ? 0 : n);
                }, 0);
                if (sum > 0) amountToSave = sum;
              }

              // ED: si no hay dirección válida, fallback determinístico a "Pagador - CUIT/CUIL:.."
              let buildingFinal = ((extractedData.ed || '') + '').toString().trim() || null;
              let edFallbackApplied = false;
              let edFallbackKind = '';

              const hasValidEd = buildingFinal && isValidEd_(buildingFinal);
              const emailEdCandidate = emailEd;
              const ocrEdCandidate = ocrOnly ? extractEdFromOcr_(ocrOnly) : null;
              const hasAnyValidAddressCandidate = (emailEdCandidate && isValidEd_(emailEdCandidate)) || (ocrEdCandidate && isValidEd_(ocrEdCandidate));
              // Fallback: si tenemos DPTO pero no ED, usar Beneficiario/Observaciones como identificador de edificio.
              if (!hasValidEd && !hasAnyValidAddressCandidate && currentDptoNorm) {
                // Prioridad a Beneficiario en OCR
                let ocrBenef = null;
                const benefRe = /\bbeneficiario\b\s*:\s*([A-ZÁÉÍÓÚÑ0-9 .,'/-]{3,80})/i;
                const mb = benefRe.exec(ocrOnly || '');
                if (mb && mb[1]) ocrBenef = mb[1].trim();

                const obsRe = /\bobservac(?:iones)?\b\s*:\s*([A-ZÁÉÍÓÚÑ0-9 .,'/-]{3,120})/i;
                const mo = obsRe.exec(ocrOnly || '');
                let ocrObs = null;
                if (mo && mo[1]) ocrObs = mo[1].trim();

                const fallbackEd = ocrBenef || ocrObs;
                if (fallbackEd) {
                  buildingFinal = fallbackEd;
                  edFallbackApplied = true;
                  edFallbackKind = 'OBS';
                  qaTags.push('FIX_ED_OBS');
                }
              }

              if (!hasValidEd && !hasAnyValidAddressCandidate) {
                const payerCuit = extractPayerCuitFromOcr_(ocrText);
                if (payerCuit && payerCuit.name && payerCuit.cuit) {
                  const payerName = payerCuit.name.trim();
                  buildingFinal = `${payerName} - CUIT/CUIL:${payerCuit.cuit}`;
                  edFallbackApplied = true;
                  edFallbackKind = 'PAGADOR_CUIT';
                  qaTags.push('FIX_ED_PAGADOR_CUIT');
                }
              }

              // Si no hay dirección y no hubo CUIT/CUIL, fallback a pagador desde email/OCR (selector)
              if (!hasValidEd && !hasAnyValidAddressCandidate && !edFallbackApplied) {
                const payerFromExtracted = normalizePayor_(extractedData.pagador);
                let payerName2 = payerFromExtracted;
                if (!payerName2) {
                  try {
                    payerName2 = extractPayerWithCandidateSelector_(subject, body, ocrText);
                  } catch (ePay) {
                    payerName2 = null;
                  }
                }
                if (payerName2) {
                  buildingFinal = payerName2;
                  edFallbackApplied = true;
                  edFallbackKind = 'PAGADOR';
                  qaTags.push('FIX_ED_PAGADOR');
                }
              }

              // Fallback 4C: nombre del remitente capturado por el wrapper de reenvío (casillas receptoras).
              // Ej: "reenviado automáticamente desde: Diego German Scandolo ..."
              if (!hasValidEd && !hasAnyValidAddressCandidate && !edFallbackApplied) {
                const forwardedName = extractForwardedSenderName_(body);
                if (forwardedName) {
                  const parts = forwardedName.split(' ').filter(Boolean);
                  const last = parts.length ? parts[parts.length - 1] : '';
                  buildingFinal = (last && last.length >= 3) ? last : forwardedName;
                  edFallbackApplied = true;
                  edFallbackKind = 'REENVIO';
                  qaTags.push('FIX_ED_REENVIO');
                }
              }

              // Fallback 4B: firma del email
              if (!hasValidEd && !hasAnyValidAddressCandidate && !edFallbackApplied) {
                const sigName = extractSignatureNameCandidate_(body);
                if (sigName) {
                  buildingFinal = sigName;
                  edFallbackApplied = true;
                  edFallbackKind = 'FIRMA';
                  qaTags.push('FIX_ED_FIRMA');
                }
              }

              // Fallback 4A: Motivo/Detalle (OCR) si parece nombre de persona
              if (!hasValidEd && !hasAnyValidAddressCandidate && !edFallbackApplied) {
                const motivoName = extractMotivoDetalleNameFromOcr_(ocrText);
                if (motivoName) {
                  buildingFinal = motivoName;
                  edFallbackApplied = true;
                  edFallbackKind = 'MOTIVO';
                  qaTags.push('FIX_ED_MOTIVO');
                }
              }

              if (!edFallbackApplied && (!buildingFinal || !isValidEd_(buildingFinal))) {
                qaTags.push('WARN_ED_INVALID');
              }

              // Si faltan datos críticos, no completar la fila: etiquetar para revisión humana y continuar
              const missingMonto = (amountToSave == null || amountToSave === '');
              const missingEd = (!buildingFinal || buildingFinal.toString().trim() === '');
              // Política cauta: si hay múltiples comprobantes, forzar revisión.
              if (missingMonto || missingEd || multipleReceipts) {
                const reasons = [];
                if (missingMonto) reasons.push('MONTO');
                if (missingEd) reasons.push('ED');
                if (multipleReceipts) reasons.push('MULTIPLES_COMPROBANTES');
                log(`(Central) REQUIERE REVISION: faltante ${reasons.join(', ')} | ${subject}`);
                thread.addLabel(etiquetaRequiereRevision);
                clearDiscardLastMsgMs_(threadId);
                thread.removeLabel(etiquetaEnProceso);
                clearThreadLease_(threadId);
                threadFinalized = true;
                break;
              }
              if (edFallbackApplied && (edFallbackKind === 'PAGADOR' || edFallbackKind === 'PAGADOR_CUIT' || edFallbackKind === 'FIRMA' || edFallbackKind === 'MOTIVO' || edFallbackKind === 'REENVIO')) {
                qaTags.push('QA_ED_SIN_DIRECCION');
              }
              const commentParts = [];
              if (validation.review) {
                commentParts.push(`REVIEW: ${validation.reviewReason || 'LLM_UNCERTAIN'}`);
              }

              // Si no hay DPTO ni UF, intentar extraer pagador (ayuda a revisión humana)
              if (!extractedData.dpto && !extractedData.uf) {
                let payerCandidate = normalizePayor_(extractedData.pagador);
                if (!payerCandidate) {
                  try {
                    payerCandidate = extractPayerWithCandidateSelector_(subject, body, ocrText);
                    if (payerCandidate) {
                      commentParts.push(`PAGADOR: ${payerCandidate}`);
                    }
                  } catch (eP) {
                    log(`(Central) Error LLM pagador: ${eP.toString()}`);
                  }
                } else {
                  commentParts.push(`PAGADOR: ${payerCandidate}`);
                }
              }

              if (estado === 'PENDIENTE') {
                commentParts.push('ESTADO: PENDIENTE');
              }
              if (montos && montos.length && montos.length > 1) {
                commentParts.push(`MULTI: ${montos.length} comprobantes`);
              }
              if (qaTags && qaTags.length) {
                commentParts.push(`QA: ${qaTags.join(', ')}`);
              }
              // Marcar observación de mensaje a administración si el LLM lo detectó
              const obsAdmin = (extractedData.observacion_admin || '').toString().trim();
              if (obsAdmin) {
                commentParts.push('ADMIN: revisar mensaje del pagador');
                qaTags.push('QA_MSG_ADMIN_LLM');
              }

              const labelSuffix = commentParts.length > 0 ? ` (${commentParts.join(' | ')})` : '';
              const rawLabel = `Abrir email${labelSuffix}`;

              buildingFinal = normalizeEdForSheet_(buildingFinal);
              
              saveToSheet(
                noticeDate, 
                extractedData.fecha_pago, 
                amountToSave, 
                buildingFinal,
                extractedData.dpto,
                extractedData.uf,
                highlightStatus,
                threadUrl,
                rawLabel
              );
              
              log(`(Central) Datos extraídos: ${JSON.stringify(extractedData)}`);
            } catch (extractError) {
              log(`(Central) Error extrayendo datos con IA: ${extractError.toString()}`);
              thread.addLabel(etiquetaRequiereRevision);
              clearDiscardLastMsgMs_(threadId);
              thread.removeLabel(etiquetaEnProceso);
              clearThreadLease_(threadId);
              threadFinalized = true;
              break;
            }
            
            thread.addLabel(etiqueta);
            clearDiscardLastMsgMs_(threadId);
            thread.removeLabel(etiquetaEnProceso);
            clearThreadLease_(threadId);
            stats.reenviados++; // usamos este campo como "marcados"
            threadFinalized = true;
            break;
          } else {
            log(`(Central) ✗ NO es expensa (keywords): ${message.getSubject()}`);
          }
        }

        // Si no se finalizó, liberar lease para que no quede colgado
        if (!threadFinalized) {
          thread.removeLabel(etiquetaEnProceso);
          clearThreadLease_(threadId);
          if (processedAllMessages) {
            thread.addLabel(etiquetaDescartado);
            setDiscardLastMsgMs_(threadId, thread.getLastMessageDate());
          }
        }
      } catch (e) {
        const msg = (e && (e.message || e.toString())) ? (e.message || e.toString()) : '';
        if (msg.indexOf('OCR_BUDGET_EXCEEDED') !== -1) {
          log('(Central) STOP: OCR_BUDGET_EXCEEDED, se retoma en próxima corrida');
          stats.stopRun = true;
          stats.needRerun = true;
          stats.rerunDelayMs = Math.max(stats.rerunDelayMs, 12 * 60 * 1000);
          return stats;
        }
        if (isRateLimitError_(e)) {
          log('(Central) STOP: rate limit Drive, se retoma en próxima corrida');
          stats.stopRun = true;
          stats.needRerun = true;
          stats.rerunDelayMs = Math.max(stats.rerunDelayMs, 12 * 60 * 1000);
          return stats;
        }
        log(`(Central) Error en thread: ${e.toString()}`);
        stats.errores++;
      }
    }
  } catch (e) {
    log(`(Central) Error buscando emails: ${e.toString()}`);
    stats.errores++;
  }
  return stats;
}

function esComprobanteExpensa(message) {
  const asunto = (message.getSubject() || '').toLowerCase();
  const cuerpo = (message.getPlainBody() || '').toLowerCase();
  const remitente = (message.getFrom() || '').toLowerCase();
  const textoCompleto = `${asunto} ${cuerpo}`;
  const tieneAdjuntos = message.getAttachments().length > 0;

  let puntuacion = 0;
  let criteriosCumplidos = [];

  // ===== EXCLUSIONES FUERTES (descarta inmediatamente) =====
  const tieneExclusionFuerte = PALABRAS_EXCLUSION_FUERTE.some(p => textoCompleto.includes(p));
  if (tieneExclusionFuerte) {
    if (CONFIG.DEBUG) {
      log(`(Central) Clasificación: ${message.getSubject()}`);
      log(`  DESCARTADO: Coincide con exclusión fuerte`);
    }
    return false;
  }

  // ===== DETECCIÓN DE CONTEXTO DE EXPENSAS (para validar remitentes bancarios) =====
  const tieneExpensa = PALABRAS_CLAVE_EXPENSA.expensa.some(p => asunto.includes(p) || cuerpo.includes(p));
  const mencionaConsorcio = PALABRAS_CLAVE_EXPENSA.consorcio.some(p => asunto.includes(p) || cuerpo.includes(p));
  const tieneDireccionConExpensa = PALABRAS_CLAVE_EXPENSA.direcciones.some(d => textoCompleto.includes(d)) && tieneExpensa;
  const destinatarioEsArtuso = /(?:destinatario|cuenta\s*destino|para)[:\s]+[^,\n]*artuso/i.test(cuerpo);
  
  const tieneContextoExpensas = tieneExpensa || mencionaConsorcio || tieneDireccionConExpensa || destinatarioEsArtuso;

  // ===== DETECCIÓN DE REMITENTE BANCARIO CONOCIDO =====
  // Solo suma puntos si hay contexto de expensas (evita falsos positivos en cuentas mixtas)
  const esRemitenteBancario = PALABRAS_CLAVE_EXPENSA.remitentes_bancos.some(r => remitente.includes(r));
  if (esRemitenteBancario && tieneContextoExpensas) {
    puntuacion += 4;
    criteriosCumplidos.push('Remitente bancario conocido + contexto expensas');
  } else if (esRemitenteBancario) {
    // Log para debugging: remitente bancario sin contexto
    if (CONFIG.DEBUG) {
      log(`  Remitente bancario detectado pero SIN contexto de expensas - no suma puntos`);
    }
  }

  // ===== PATRÓN DE ASUNTO "pago expensas [DIRECCIÓN]" =====
  const patronAsuntoPagoExpensas = /(?:pago|expensas?)\s+(?:expensas?)?\s*[a-záéíóúñ]+\s+\d+/i;
  if (patronAsuntoPagoExpensas.test(asunto)) {
    puntuacion += 5;
    criteriosCumplidos.push('Asunto con patrón "pago expensas [dirección]"');
  }

  // ===== PATRÓN "Expensas UF XX" en asunto =====
  if (/expensas?\s+uf\s*\d+/i.test(asunto)) {
    puntuacion += 5;
    criteriosCumplidos.push('Asunto con patrón "Expensas UF [número]"');
  }

  // ===== PALABRAS CLAVE BÁSICAS =====
  const palabrasBasicas = [
    ...PALABRAS_CLAVE_EXPENSA.expensa,
    ...PALABRAS_CLAVE_EXPENSA.pago,
    ...PALABRAS_CLAVE_EXPENSA.comprobante
  ];
  const coincidenciasAsunto = palabrasBasicas.filter(p => asunto.includes(p.toLowerCase()));
  if (coincidenciasAsunto.length > 0) {
    puntuacion += 2;
    criteriosCumplidos.push('Palabras clave en asunto');
  }

  const coincidenciasCuerpo = palabrasBasicas.filter(p => cuerpo.includes(p.toLowerCase()));
  if (coincidenciasCuerpo.length >= 2) {
    puntuacion += 3;
    criteriosCumplidos.push('Múltiples palabras clave en cuerpo');
  }

  // ===== COMBINACIONES FUERTES =====
  const tienePago = PALABRAS_CLAVE_EXPENSA.pago.some(p => asunto.includes(p) || cuerpo.includes(p));
  const tieneComprobante = PALABRAS_CLAVE_EXPENSA.comprobante.some(p => asunto.includes(p) || cuerpo.includes(p));
  
  if (tieneExpensa && (tienePago || tieneComprobante)) {
    puntuacion += 5;
    criteriosCumplidos.push('Combinación expensa + pago/comprobante');
  }

  // ===== FRASES DE PAGO/TRANSFERENCIA =====
  const tieneFrasesPago = PALABRAS_CLAVE_EXPENSA.frases_pago.some(p => textoCompleto.includes(p));
  if (tieneFrasesPago) {
    puntuacion += 4;
    criteriosCumplidos.push('Frases típicas de aviso de pago/transferencia');
  }

  // ===== ENTIDADES BANCARIAS =====
  const tieneBanco = PALABRAS_CLAVE_EXPENSA.bancos.some(p => textoCompleto.includes(p) || remitente.includes(p));
  if (tieneBanco) {
    puntuacion += 2;
    criteriosCumplidos.push('Menciona entidad bancaria/financiera');
    
    // Bonus si además tiene contexto de pago
    if (tienePago || tieneComprobante || tieneExpensa) {
      puntuacion += 2;
      criteriosCumplidos.push('Banco + contexto de pago');
    }
  }

  // ===== ADJUNTOS =====
  if (tieneAdjuntos) {
    puntuacion += 2;
    criteriosCumplidos.push('Tiene adjuntos');

    const tiposPermitidos = ['pdf', 'jpg', 'jpeg', 'png', 'gif'];
    const adjuntosValidos = message.getAttachments().some(adj => {
      const nombre = adj.getName().toLowerCase();
      return tiposPermitidos.some(tipo => nombre.endsWith('.' + tipo));
    });
    if (adjuntosValidos) {
      puntuacion += 1;
      criteriosCumplidos.push('Adjuntos de tipo imagen/PDF');
    }
  }

  // ===== REFERENCIAS BANCARIAS (CBU, CVU, CUIT, etc.) =====
  const tieneReferenciaBancaria = PALABRAS_CLAVE_EXPENSA.referencia.some(p => cuerpo.includes(p));
  const tieneNumeroCBU = /\d{20,}/.test(cuerpo); // CBU/CVU tienen 22 dígitos
  const tieneCUIT = /\b\d{2}-?\d{8}-?\d{1}\b/.test(cuerpo); // CUIT formato XX-XXXXXXXX-X
  
  if (tieneReferenciaBancaria && (tieneNumeroCBU || tieneCUIT)) {
    puntuacion += 3;
    criteriosCumplidos.push('Referencia bancaria (CBU/CVU/CUIT) detectada');
  }

  // ===== CONSORCIO/UNIDAD =====
  // (mencionaConsorcio ya fue calculado arriba para contexto de expensas)
  if (mencionaConsorcio) {
    puntuacion += 2;
    criteriosCumplidos.push('Menciona consorcio/unidad/piso');
  }

  // ===== DIRECCIONES EN ASUNTO (indica pago de expensas de una propiedad) =====
  const tieneDireccionEnAsunto = PALABRAS_CLAVE_EXPENSA.direcciones.some(d => asunto.includes(d));
  const tieneNumeroEnAsunto = /\d{2,5}/.test(asunto); // número de calle
  if (tieneDireccionEnAsunto && tieneNumeroEnAsunto) {
    puntuacion += 3;
    criteriosCumplidos.push('Dirección detectada en asunto');
  }

  // ===== MONTOS DE DINERO =====
  // Patrón mejorado para montos argentinos: $164.185,30 o 436.867.55 o $ 360.000
  const patronMonto = /(?:\$\s*)?\d{1,3}(?:[\.\s]\d{3})+(?:,\d{2})?|\$\s*\d{4,}/i;
  const tieneMontoGrande = patronMonto.test(textoCompleto);
  if (tieneMontoGrande && (tieneExpensa || tienePago || tieneComprobante || tieneBanco)) {
    puntuacion += 3;
    criteriosCumplidos.push('Monto de dinero + contexto de pago/expensas');
  }

  // ===== MOTIVO EXPENSAS EN TRANSFERENCIA =====
  if (/motivo:?\s*expensas?/i.test(cuerpo)) {
    puntuacion += 4;
    criteriosCumplidos.push('Motivo de transferencia = EXPENSAS');
  }

  // ===== EXCLUSIONES (reduce puntuación) =====
  const palabrasExcluidas = ['factura de servicio', 'publicidad', 'newsletter', 'promoción', 'suscripción', 'oferta'];
  const tieneExcluidas = palabrasExcluidas.some(p => asunto.includes(p) || cuerpo.includes(p));
  if (tieneExcluidas && puntuacion < 8) {
    puntuacion -= 3;
    criteriosCumplidos.push('Palabras excluidas detectadas (penalización)');
  }

  // ===== LOGGING =====
  if (CONFIG.DEBUG) {
    log(`(Central) Clasificación: ${message.getSubject()}`);
    log(`  Remitente: ${remitente}`);
    log(`  Puntuación: ${puntuacion}`);
    log(`  Criterios: ${criteriosCumplidos.join(', ')}`);
  }

  const UMBRAL_MINIMO = 5;
  return puntuacion >= UMBRAL_MINIMO;
}

function crearObtenerEtiqueta(nombreEtiqueta) {
  let etiqueta = GmailApp.getUserLabelByName(nombreEtiqueta);
  if (!etiqueta) {
    etiqueta = GmailApp.createLabel(nombreEtiqueta);
    log(`Etiqueta creada: ${nombreEtiqueta}`);
  }
  return etiqueta;
}

function log(mensaje) {
  if (CONFIG.DEBUG) {
    console.log(`[${new Date().toISOString()}] ${mensaje}`);
  }
}

function configurarTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'procesarEmailsExpensas') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('procesarEmailsExpensas').timeBased().everyHours(1).create();
  log('(Central) Trigger configurado cada 1 hora');
}

function ejecutarPrueba() {
  log('=== MODO PRUEBA (central) ===');
  CONFIG.DEBUG = true;

  const emailOrigen = CONFIG.EMAILS_ORIGEN[0];
  try {
    const query = `in:inbox newer_than:7d`;
    const threads = GmailApp.search(query, 0, 5);

    log(`\n--- Prueba en cuenta central (origen lógico: ${emailOrigen}) ---`);
    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        const esExpensa = esComprobanteExpensa(message);
        log(`Asunto: ${message.getSubject()}`);
        log(`Resultado: ${esExpensa ? '✓ ES EXPENSA' : '✗ NO ES EXPENSA'}\n`);
      });
    });
  } catch (error) {
    log(`Error en prueba (central): ${error.toString()}`);
  }
}

/**
 * Prueba la extracción con IA sobre los últimos emails detectados como expensa.
 * Ejecutar manualmente para validar que la IA extrae correctamente los datos.
 * NO guarda en el Sheet, solo muestra los resultados en el log.
 */
function testAIExtraction() {
  log('=== PRUEBA DE EXTRACCIÓN CON IA ===');
  CONFIG.DEBUG = true;
  
  try {
    // Buscar emails ya etiquetados como expensas
    const query = `label:${CONFIG.ETIQUETA_PROCESADO} newer_than:7d`;
    const threads = GmailApp.search(query, 0, 3);
    
    if (threads.length === 0) {
      log('No se encontraron emails etiquetados como expensas en los últimos 7 días.');
      log('Ejecuta primero ejecutarPrueba() para detectar expensas.');
      return;
    }
    
    log(`Probando extracción en ${threads.length} threads...\n`);
    
    threads.forEach((thread, index) => {
      const message = thread.getMessages()[0];
      const subject = message.getSubject();
      const body = message.getPlainBody();
      const ocrAdj = extraerTextoDeAdjuntosExpensa(message);
      const ocrInline = extraerTextoDeImagenesEnCuerpo(message);
      const ocrText = [ocrAdj, ocrInline].filter(Boolean).join('\n\n');
      const bodyForAI = composeBodyWithOcr_(body, ocrText);
      const extractorContext = buildExtractorContext_(subject, body, ocrText);
      
      log(`--- Email ${index + 1}: ${subject} ---`);
      
      try {
        const validation = validateExpensePaymentWithLLM(subject, bodyForAI);
        log(`Validador: decision=${validation.isValid ? (validation.review ? 'ACCEPT_WITH_REVIEW' : 'ACCEPT') : 'REJECT'} | review=${validation.review ? (validation.reviewReason || 'LLM_UNCERTAIN') : 'none'}`);
        const extractedData = extractDataWithAI(subject, extractorContext);
        log(`Monto: ${extractedData.monto}`);
        log(`Fecha pago: ${extractedData.fecha_pago}`);
        log(`Departamento: ${extractedData.dpto}`);
        log(`Estado: ${extractedData.estado}`);
        log(`Edificio: ${extractedData.ed}`);
        log(`Extracción exitosa ✓\n`);
      } catch (e) {
        log(`Error en extracción: ${e.toString()}\n`);
      }
    });
    
    log('=== FIN DE PRUEBA ===');
  } catch (error) {
    log(`Error en prueba de IA: ${error.toString()}`);
  }
}

/**
 * Prueba el OCR (Drive) sobre los últimos emails con adjuntos/imagenes inline.
 * No llama a la IA: solo muestra el texto OCR en Logger.
 *
 * Requiere:
 * - Habilitar Advanced Google Service "Drive API" en Apps Script
 */
function testDriveOcr() {
  log('=== PRUEBA OCR (Drive) ===');
  CONFIG.DEBUG = true;

  if (!OCR_CONFIG.ENABLED) {
    log('OCR_CONFIG.ENABLED=false; habilítalo para probar.');
    return;
  }

  if (!canUseDriveOcr_()) {
    log('Falta habilitar el servicio avanzado "Drive API".');
    return;
  }

  try {
    const query = 'in:inbox newer_than:14d has:attachment';
    const threads = GmailApp.search(query, 0, 3);
    log(`Threads encontrados: ${threads.length}`);

    threads.forEach((thread, index) => {
      const message = thread.getMessages()[0];
      const subject = message.getSubject();
      log(`--- Email ${index + 1}: ${subject} ---`);

      const ocrAdj = extraerTextoDeAdjuntosExpensa(message);
      const ocrInline = extraerTextoDeImagenesEnCuerpo(message);
      const ocrText = [ocrAdj, ocrInline].filter(Boolean).join('\n\n');

      if (!ocrText) {
        log('OCR: (vacío)');
      } else {
        log(`OCR chars=${ocrText.length}`);
        log(truncateText_(ocrText, 2500));
      }
    });

    log('=== FIN PRUEBA OCR ===');
  } catch (error) {
    log(`Error en prueba OCR: ${error.toString()}`);
  }
}

/**
 * Prueba rápida de conexión con la API de IA.
 * Ejecutar para verificar que la API key está configurada correctamente.
 */
function testAIConnection() {
  log('=== PRUEBA DE CONEXIÓN CON IA ===');
  log(`Proveedor: ${AI_CONFIG.PROVIDER}`);
  log(`Modelo: ${AI_CONFIG.MODEL}`);
  
  try {
    const testPrompt = 'Responde solo con la palabra "OK" si recibes este mensaje.';
    const response = callAI(testPrompt);
    log(`Respuesta: ${response}`);
    log('Conexión exitosa ✓');
  } catch (e) {
    log(`Error de conexión: ${e.toString()}`);
    log('');
    log('Para configurar la API Key:');
    log('1. Ir a Configuración del proyecto (ícono de engranaje)');
    log('2. Scroll hasta "Propiedades del script"');
    log('3. Agregar propiedad: OPENAI_API_KEY (o GEMINI_API_KEY, ANTHROPIC_API_KEY)');
    log('4. Valor: tu API key');
    log('');
    log('O ejecuta setupAPIKey("tu-api-key") para configurarla desde el código.');
  }
}

/**
 * Configura la API key en las Propiedades del script.
 * Ejecutar UNA VEZ para guardar la API key de forma segura.
 * Después de ejecutar, la API key queda guardada y no necesitas volver a ejecutar.
 * 
 * @param {string} apiKey - La API key a guardar
 * @param {string} provider - (Opcional) El proveedor: 'openai', 'gemini', 'anthropic'. Default: AI_CONFIG.PROVIDER
 * 
 * Ejemplo de uso:
 *   setupAPIKey('sk-1234567890abcdef...');
 *   setupAPIKey('AIzaSy...', 'gemini');
 */
function setupAPIKey(apiKey, provider) {
  if (!apiKey || apiKey.trim() === '') {
    log('Error: Debes proporcionar una API key válida');
    log('Uso: setupAPIKey("tu-api-key")');
    return;
  }
  
  const targetProvider = provider || AI_CONFIG.PROVIDER;
  const propertyNames = {
    'openai': 'OPENAI_API_KEY',
    'gemini': 'GEMINI_API_KEY',
    'anthropic': 'ANTHROPIC_API_KEY'
  };
  
  const propertyName = propertyNames[targetProvider];
  if (!propertyName) {
    log(`Error: Proveedor no soportado: ${targetProvider}`);
    return;
  }
  
  PropertiesService.getScriptProperties().setProperty(propertyName, apiKey);
  log(`✓ API Key guardada exitosamente para ${targetProvider}`);
  log(`  Propiedad: ${propertyName}`);
  log('');
  log('Ahora puedes ejecutar testAIConnection() para verificar la conexión.');
}

/**
 * Prueba la validación LLM sobre los últimos emails de la bandeja de entrada.
 * Ejecutar manualmente para verificar cómo clasifica el LLM antes de activar en producción.
 * NO guarda en el Sheet ni etiqueta emails, solo muestra resultados en el log.
 */
function testLLMValidation() {
  log('=== PRUEBA DE VALIDACIÓN LLM ===');
  CONFIG.DEBUG = true;
  
  try {
    // Buscar los últimos emails (sin filtrar por etiqueta)
    const query = `in:inbox newer_than:7d`;
    const threads = GmailApp.search(query, 0, 5);
    
    if (threads.length === 0) {
      log('No se encontraron emails en los últimos 7 días.');
      return;
    }
    
    log(`Probando validación en ${threads.length} threads...\n`);
    
    let passedKeywords = 0;
    let passedLLM = 0;
    let reviewLLM = 0;
    let rejectedLLM = 0;
    
    threads.forEach((thread, index) => {
      const message = thread.getMessages()[0];
      const subject = message.getSubject();
      const body = message.getPlainBody();
      
      log(`--- Email ${index + 1}: ${subject} ---`);
      
      // Paso 1: Verificar si pasa el filtro de keywords
      const passKeywords = esComprobanteExpensa(message);
      log(`  Keywords: ${passKeywords ? 'PASA' : 'NO PASA'}`);
      
      if (passKeywords) {
        passedKeywords++;
        
        // Paso 2: Verificar con LLM
        try {
          const validation = validateExpensePaymentWithLLM(subject, body);
          const llmStatus = validation.isValid
            ? (validation.review ? 'REVIEW' : 'CONFIRMADO')
            : 'RECHAZADO';
          log(`  LLM: ${llmStatus}`);
          
          if (validation.isValid) {
            if (validation.review) {
              reviewLLM++;
            } else {
              passedLLM++;
            }
          } else {
            rejectedLLM++;
          }
        } catch (e) {
          log(`  LLM Error: ${e.toString()}`);
        }
      }
      
      log('');
    });
    
    log('=== RESUMEN ===');
    log(`Total emails analizados: ${threads.length}`);
    log(`Pasaron filtro keywords: ${passedKeywords}`);
    log(`Confirmados por LLM: ${passedLLM}`);
    log(`En revisión por LLM: ${reviewLLM}`);
    log(`Rechazados por LLM: ${rejectedLLM}`);
    log(`Tasa de rechazo LLM: ${passedKeywords > 0 ? Math.round(rejectedLLM / passedKeywords * 100) : 0}%`);
    log('=== FIN DE PRUEBA ===');
  } catch (error) {
    log(`Error en prueba de validación LLM: ${error.toString()}`);
  }
}
