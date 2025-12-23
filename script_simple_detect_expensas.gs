/**
 * Script simple: detecta expensas por keywords y reenvia con etiqueta.
 */

// ==================== CONFIGURACION ====================
const CONFIG_DETECT = {
  // Email logico de esta cuenta (solo para logging)
  EMAILS_ORIGEN: ['tu-casilla@gmail.com'],
  EMAIL_DESTINO: 'sonntagnahuel@gmail.com',
  ALERT_EMAIL: 'csuarezgurruchaga@gmail.com',
  ETIQUETA_DETECTADA: 'EXPENSA DETECTADA',
  ETIQUETA_NO_PAGO: 'NO ES PAGO DE EXPENSA',
  MAX_REENVIOS_DIARIOS: 80,
  MAX_THREADS_PER_RUN: 200,
  ALERT_STREAK_DAYS: 2,
  ALERT_ENABLED: true,
  TRIGGER_EVERY_HOURS: 2,
  DEBUG: true
};

const PALABRAS_EXPENSA_DETECT = [
  'expensa', 'expensas',
  // Typos comunes
  'expesnas',   // intercambio n/s
  'espensas',   // x->s
  'expnsas',    // falta e
  'expesas',    // falta n
  'expenss',    // falta a
  'expenas',    // falta s
  'exppensas',  // p duplicada
  'expensass',  // s duplicada
  'expenzas',   // s->z
  // Formato especial
  'expensas uf'
];

const EMAILS_EXCLUIR_DETECT = [
  'artusoexpensas2@gmail.com'
];

const EMAILS_REQUIEREN_ADJUNTO_DETECT = [
  // Remitentes automáticos: aceptar solo si traen adjunto real.
  'webmaster@consorciosenredweb.com'
];

const REGEX_RESUMEN_PROCESAMIENTO_DETECT = /^resumen\s+procesamiento\s+expensas\s+-\s+\d{2}\/\d{2}\/\d{4}$/i;

function procesarEmailsExpensasDetectadas() {
  logDetect_('=== Iniciando deteccion simple de expensas ===');

  const etiquetaDetectada = crearObtenerEtiquetaDetect_(CONFIG_DETECT.ETIQUETA_DETECTADA);
  const etiquetaNoPago = crearObtenerEtiquetaDetect_(CONFIG_DETECT.ETIQUETA_NO_PAGO);
  const query = `in:inbox ${getMonthStartQueryDetect_()} -label:${CONFIG_DETECT.ETIQUETA_DETECTADA} -label:${CONFIG_DETECT.ETIQUETA_NO_PAGO}`;
  const threads = GmailApp.search(query, 0, CONFIG_DETECT.MAX_THREADS_PER_RUN);
  const state = getDailySendState_();
  let remaining = Math.max(0, CONFIG_DETECT.MAX_REENVIOS_DIARIOS - state.count);
  let reachedLimit = remaining <= 0;
  let processedThreads = 0;

  logDetect_(`Encontrados ${threads.length} threads para detectar`);
  logDetect_(`Reenvios restantes hoy: ${remaining}`);

  threads.sort(function(a, b) {
    return a.getLastMessageDate().getTime() - b.getLastMessageDate().getTime();
  });

  for (var i = 0; i < threads.length; i++) {
    if (remaining <= 0) {
      reachedLimit = true;
      break;
    }
    const thread = threads[i];
    processedThreads++;
    try {
      const messages = thread.getMessages();
      const resumen = messages.some(function(message) {
        const subject = (message.getSubject() || '').trim();
        return REGEX_RESUMEN_PROCESAMIENTO_DETECT.test(subject);
      });
      if (resumen) {
        thread.addLabel(etiquetaNoPago);
        thread.removeLabel(etiquetaDetectada);
        logDetect_(`SKIP resumen: ${thread.getFirstMessageSubject()}`);
        continue;
      }

      const matchMessage = findEligibleMessageDetect_(messages);

      if (!matchMessage) {
        thread.addLabel(etiquetaNoPago);
        thread.removeLabel(etiquetaDetectada);
        logDetect_(`✗ NO PAGO: ${thread.getFirstMessageSubject()}`);
        continue;
      }

      const attachments = collectThreadAttachmentsDetect_(messages);
      reenviarEmailDetect_(matchMessage, attachments);
      incrementDailySendCount_(state);
      remaining = Math.max(0, remaining - 1);
      thread.addLabel(etiquetaDetectada);
      thread.removeLabel(etiquetaNoPago);
      logDetect_(
        `✓ EXPENSA DETECTADA: ${thread.getFirstMessageSubject()}` +
        (attachments.length ? ` (adjuntos: ${attachments.length})` : '')
      );
    } catch (e) {
      logDetect_(`Error en thread: ${e.toString()}`);
    }
  }

  const limitStreak = updateLimitStreakDetect_(reachedLimit);
  if (reachedLimit) {
    const backlogInfo = estimateBacklogDetect_(query, threads, processedThreads);
    maybeSendLimitAlertDetect_(limitStreak, backlogInfo);
    logDetect_(
      'Limite diario alcanzado, se continua en la proxima ejecucion. ' +
      `Backlog estimado: ${backlogInfo.backlogEstimate}`
    );
    return;
  }

  if (threads.length >= CONFIG_DETECT.MAX_THREADS_PER_RUN) {
    logDetect_('Quedan threads pendientes, se procesaran en la proxima ejecucion.');
  }
}

function containsExpensaKeywordDetect_(text) {
  const t = normalizeForMatchDetect_(text);
  if (!t) return false;
  for (var i = 0; i < PALABRAS_EXPENSA_DETECT.length; i++) {
    const kw = PALABRAS_EXPENSA_DETECT[i];
    if (t.indexOf(kw) !== -1) return true;
  }
  return false;
}

function normalizeForMatchDetect_(text) {
  let t = stripAccentsDetect_(text).toLowerCase();
  for (var i = 0; i < EMAILS_EXCLUIR_DETECT.length; i++) {
    t = t.split(EMAILS_EXCLUIR_DETECT[i]).join(' ');
  }
  return t.replace(/\s+/g, ' ').trim();
}

function stripAccentsDetect_(text) {
  return (text || '').toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function findEligibleMessageDetect_(messages) {
  for (var i = 0; i < messages.length; i++) {
    const message = messages[i];
    const subject = message.getSubject() || '';
    const body = message.getPlainBody() || '';
    if (!containsExpensaKeywordDetect_(subject) && !containsExpensaKeywordDetect_(body)) {
      continue;
    }
    if (senderRequiresAttachmentDetect_(message) && !messageHasRealAttachmentsDetect_(message)) {
      logDetect_(`SKIP remitente sin adjunto: ${message.getSubject()}`);
      continue;
    }
    return message;
  }
  return null;
}

function senderRequiresAttachmentDetect_(message) {
  const from = getSenderEmailDetect_(message.getFrom());
  return EMAILS_REQUIEREN_ADJUNTO_DETECT.indexOf(from) !== -1;
}

function getSenderEmailDetect_(fromValue) {
  const from = (fromValue || '').toString();
  const angleMatch = from.match(/<([^>]+)>/);
  let email = angleMatch && angleMatch[1] ? angleMatch[1] : '';
  if (!email) {
    const match = from.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
    email = match ? match[0] : from;
  }
  return (email || '').toLowerCase().trim();
}

function messageHasRealAttachmentsDetect_(message) {
  const attachments = message.getAttachments({ includeInlineImages: false });
  return attachments && attachments.length > 0;
}

function getMonthStartQueryDetect_() {
  const tz = 'America/Argentina/Buenos_Aires';
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const formatted = Utilities.formatDate(start, tz, 'yyyy/MM/dd');
  return `after:${formatted}`;
}

function reenviarEmailDetect_(message, attachments) {
  const subject = message.getSubject() || '';
  const body = message.getPlainBody() || '';
  const fecha = Utilities.formatDate(
    message.getDate(),
    'America/Argentina/Buenos_Aires',
    'dd/MM/yyyy HH:mm'
  );
  const options = {};
  const attachmentsToSend = Array.isArray(attachments) && attachments.length > 0
    ? attachments
    : message.getAttachments({ includeInlineImages: true });
  if (attachmentsToSend && attachmentsToSend.length > 0) {
    options.attachments = attachmentsToSend;
  }

  GmailApp.sendEmail(
    CONFIG_DETECT.EMAIL_DESTINO,
    `[REENVIADO] ${subject}`,
    `Este email fue reenviado automaticamente desde: ${message.getFrom()}\n\n` +
      `Fecha original: ${fecha}\n` +
      `Asunto original: ${subject}\n\n` +
      `--- CONTENIDO ORIGINAL ---\n${body}`,
    options
  );
}

function collectThreadAttachmentsDetect_(messages) {
  const allowedExt = ['pdf', 'jpg', 'jpeg', 'png', 'gif'];
  const allowedTypes = ['application/pdf', 'image/jpeg', 'image/png', 'image/gif'];
  const all = [];

  for (var i = 0; i < messages.length; i++) {
    const attachments = messages[i].getAttachments({ includeInlineImages: true });
    for (var j = 0; j < attachments.length; j++) {
      const attachment = attachments[j];
      const name = (attachment.getName() || '').toLowerCase();
      const contentType = (attachment.getContentType() || '').toLowerCase();
      const hasAllowedExt = allowedExt.some(function(ext) {
        return name.endsWith('.' + ext);
      });
      const hasAllowedType = allowedTypes.indexOf(contentType) !== -1;
      if (hasAllowedExt || hasAllowedType) {
        all.push(attachment);
      }
    }
  }

  return all;
}

function getDailySendState_() {
  const props = PropertiesService.getScriptProperties();
  const tz = 'America/Argentina/Buenos_Aires';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const savedDay = props.getProperty('SIMPLE_DETECT_DAY') || '';
  let count = Number(props.getProperty('SIMPLE_DETECT_COUNT') || '0');

  if (savedDay !== today) {
    count = 0;
    props.setProperty('SIMPLE_DETECT_DAY', today);
    props.setProperty('SIMPLE_DETECT_COUNT', '0');
  }

  return { count: count, day: today };
}

function incrementDailySendCount_(state) {
  const props = PropertiesService.getScriptProperties();
  const next = (state.count || 0) + 1;
  state.count = next;
  props.setProperty('SIMPLE_DETECT_COUNT', String(next));
}

function estimateBacklogDetect_(query, threads, processedThreads) {
  const pendingInBatch = Math.max(0, threads.length - processedThreads);
  let hasMore = false;
  if (threads.length >= CONFIG_DETECT.MAX_THREADS_PER_RUN && pendingInBatch > 0) {
    const more = GmailApp.search(query, CONFIG_DETECT.MAX_THREADS_PER_RUN, 1);
    hasMore = more.length > 0;
  }
  const backlogEstimate = pendingInBatch + (hasMore ? CONFIG_DETECT.MAX_THREADS_PER_RUN : 0);
  return {
    pendingInBatch: pendingInBatch,
    hasMore: hasMore,
    backlogEstimate: backlogEstimate
  };
}

function updateLimitStreakDetect_(reachedLimit) {
  const props = PropertiesService.getScriptProperties();
  const today = getTodayKeyDetect_();
  const lastLimitDay = props.getProperty('SIMPLE_DETECT_LIMIT_LAST_DAY') || '';
  let streak = Number(props.getProperty('SIMPLE_DETECT_LIMIT_STREAK') || '0');

  if (!reachedLimit) {
    return streak;
  }

  if (lastLimitDay === today) {
    return streak;
  }

  if (lastLimitDay && daysDiffDetect_(today, lastLimitDay) === 1) {
    streak += 1;
  } else {
    streak = 1;
  }

  props.setProperty('SIMPLE_DETECT_LIMIT_LAST_DAY', today);
  props.setProperty('SIMPLE_DETECT_LIMIT_STREAK', String(streak));
  return streak;
}

function maybeSendLimitAlertDetect_(streak, backlogInfo) {
  if (!CONFIG_DETECT.ALERT_ENABLED) return;
  const recipient = CONFIG_DETECT.ALERT_EMAIL || CONFIG_DETECT.EMAIL_DESTINO;
  if (!recipient) return;

  const props = PropertiesService.getScriptProperties();
  const today = getTodayKeyDetect_();
  const lastAlertDay = props.getProperty('SIMPLE_DETECT_ALERT_LAST_DAY') || '';
  if (lastAlertDay === today) return;
  if (streak < CONFIG_DETECT.ALERT_STREAK_DAYS) return;

  const daysRemaining = getDaysRemainingInMonthDetect_();
  const capacityRemaining = daysRemaining * CONFIG_DETECT.MAX_REENVIOS_DIARIOS;
  if (backlogInfo.backlogEstimate <= capacityRemaining) return;

  sendLimitAlertEmailDetect_(recipient, backlogInfo, daysRemaining, capacityRemaining, streak);
  props.setProperty('SIMPLE_DETECT_ALERT_LAST_DAY', today);
}

function sendLimitAlertEmailDetect_(recipient, backlogInfo, daysRemaining, capacityRemaining, streak) {
  const subject = 'Alerta: limite diario alcanzado en deteccion de expensas';
  const body =
    'Se alcanzo el limite diario de reenvios en la deteccion simple.\n\n' +
    `Dias restantes del mes: ${daysRemaining}\n` +
    `Capacidad restante estimada: ${capacityRemaining}\n` +
    `Backlog estimado: ${backlogInfo.backlogEstimate}\n` +
    `Pendientes en batch actual: ${backlogInfo.pendingInBatch}\n` +
    `Hay mas que el batch actual: ${backlogInfo.hasMore ? 'si' : 'no'}\n` +
    `Dias consecutivos con limite: ${streak}\n`;
  MailApp.sendEmail(recipient, subject, body);
}

function getDaysRemainingInMonthDetect_() {
  const tz = 'America/Argentina/Buenos_Aires';
  const now = new Date();
  const year = Number(Utilities.formatDate(now, tz, 'yyyy'));
  const monthIndex = Number(Utilities.formatDate(now, tz, 'MM')) - 1;
  const day = Number(Utilities.formatDate(now, tz, 'dd'));
  const lastDay = new Date(year, monthIndex + 1, 0).getDate();
  return Math.max(0, lastDay - day);
}

function getTodayKeyDetect_() {
  const tz = 'America/Argentina/Buenos_Aires';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

function daysDiffDetect_(dayA, dayB) {
  if (!dayA || !dayB) return 0;
  const a = new Date(dayA + 'T00:00:00');
  const b = new Date(dayB + 'T00:00:00');
  return Math.round((a.getTime() - b.getTime()) / 86400000);
}

function crearObtenerEtiquetaDetect_(nombreEtiqueta) {
  let etiqueta = GmailApp.getUserLabelByName(nombreEtiqueta);
  if (!etiqueta) {
    etiqueta = GmailApp.createLabel(nombreEtiqueta);
    logDetect_(`Etiqueta creada: ${nombreEtiqueta}`);
  }
  return etiqueta;
}

function configurarTriggerDetectCada2Horas() {
  const handler = 'procesarEmailsExpensasDetectadas';
  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    const t = triggers[i];
    if (t.getHandlerFunction && t.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyHours(CONFIG_DETECT.TRIGGER_EVERY_HOURS)
    .create();
  logDetect_(`Trigger configurado: ejecucion cada ${CONFIG_DETECT.TRIGGER_EVERY_HOURS} horas`);
}

function logDetect_(mensaje) {
  if (CONFIG_DETECT.DEBUG) {
    console.log(`[${new Date().toISOString()}] ${mensaje}`);
  }
}
