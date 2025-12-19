/**
 * Script para la casilla: artusocarlos14@gmail.com
 * - Se debe crear un proyecto de Apps Script LOGUEADO como artusocarlos14@gmail.com
 * - Pegar este código en Code.gs
 * - Ejecutar primero ejecutarPrueba() para autorizar
 * - Luego configurarTrigger() para dejarlo automático
 */

// ==================== CONFIGURACIÓN ====================
const CONFIG = {
  // Email lógico de esta cuenta (solo para logging)
  EMAILS_ORIGEN: ['artusocarlos14@gmail.com'],

  // Email destino donde se centralizan las expensas
  EMAIL_DESTINO: 'artusoexpensas2@gmail.com',

  // Etiqueta para marcar emails ya procesados
  ETIQUETA_PROCESADO: 'ExpensaProcesada',
  // Etiqueta para marcar emails descartados
  ETIQUETA_DESCARTADO: 'ExpensaDescartada',

  // Activar logging detallado
  DEBUG: true
};

const MIN_INLINE_IMAGE_BYTES = 30 * 1024;

// ==================== PALABRAS CLAVE PARA DETECCIÓN ====================
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
  // Respuestas/agradecimientos
  'gracias por informar el pago',
  'muchas gracias por informar',
  'recibido, gracias',
  'confirmamos recepción',
  // Actas de asamblea
  'acta de asamblea',
  // Facturas de proveedores (quieren COBRAR, no informan pago)
  'facturas del tad',
  'aguardo comprobante de pago',
  'pendientes de pago',
  // Publicidad
  'gestionás consorcios',
  'tenemos una solución pensada',
  // Notas internas
  'se adjunta una nota para el consorcio',
  'nota para el consorcio',
  // Trámites sindicales/administrativos (FATERYH, SUTERH, etc.)
  'fateryh',
  'suterh',
  'osperyhra',
  'verificación de deuda',
  'procedimiento administrativo',
  'para descargo',
  'aportes y contribuciones'
];


// ==================== FUNCIÓN PRINCIPAL ====================
function procesarEmailsExpensas() {
  log('=== Iniciando procesamiento de emails ===');

  let totalProcesados = 0;
  let totalReenviados = 0;
  let totalErrores = 0;

  CONFIG.EMAILS_ORIGEN.forEach(emailOrigen => {
    try {
      log(`Procesando emails para la cuenta actual (origen lógico: ${emailOrigen})`);
      const resultado = procesarEmailsDeCuenta(emailOrigen);

      totalProcesados += resultado.procesados;
      totalReenviados += resultado.reenviados;
      totalErrores += resultado.errores;

      log(`Resultado: ${resultado.reenviados} reenviados, ${resultado.errores} errores`);
    } catch (error) {
      log(`Error procesando ${emailOrigen}: ${error.toString()}`);
      totalErrores++;
    }
  });

  log(`=== Resumen: ${totalProcesados} procesados, ${totalReenviados} reenviados, ${totalErrores} errores ===`);

  if (CONFIG.DEBUG && totalReenviados > 0) {
    enviarResumenEjecucion(totalProcesados, totalReenviados, totalErrores);
  }
}

// ==================== PROCESAMIENTO POR CUENTA ====================
function procesarEmailsDeCuenta(emailOrigen) {
  const stats = { procesados: 0, reenviados: 0, errores: 0 };

  try {
    cleanupDiscardStateIfNeeded_();
    const etiquetaDescartado = crearObtenerEtiqueta(CONFIG.ETIQUETA_DESCARTADO);
    reabrirDescartadosConAdjunto_(etiquetaDescartado);

    const query = `in:inbox ${getMonthStartQuery_()} -label:${CONFIG.ETIQUETA_PROCESADO} -label:${CONFIG.ETIQUETA_DESCARTADO}`;
    const threads = GmailApp.search(query, 0, 50);

    log(`Encontrados ${threads.length} threads nuevos en inbox para procesar`);

    threads.forEach(thread => {
      try {
        const messages = thread.getMessages();
        
        // Seleccionar el mensaje más relevante del thread (no procesar todos)
        const mensaje = encontrarMensajeComprobante(messages);
        
        // Recopilar TODOS los adjuntos del thread (por si están distribuidos)
        const todosLosAdjuntos = messages.flatMap(m => m.getAttachments());
        
        stats.procesados++;

        const threadId = thread.getId();
        if (esComprobanteExpensa(mensaje)) {
          log(`✓ Email identificado como expensa: ${mensaje.getSubject()}`);
          reenviarEmail(mensaje, todosLosAdjuntos);
          stats.reenviados++;
          thread.addLabel(crearObtenerEtiqueta(CONFIG.ETIQUETA_PROCESADO));
          clearDiscardLastMsgMs_(threadId);
        } else {
          log(`✗ Email NO es expensa: ${mensaje.getSubject()}`);
          thread.addLabel(etiquetaDescartado);
          setDiscardLastMsgMs_(threadId, thread.getLastMessageDate());
        }
      } catch (error) {
        log(`Error procesando thread: ${error.toString()}`);
        stats.errores++;
      }
    });
  } catch (error) {
    log(`Error buscando emails: ${error.toString()}`);
    stats.errores++;
  }

  return stats;
}

/**
 * Encuentra el mensaje más relevante del thread para extraer info del comprobante
 * Prioridad: 1) Con adjuntos PDF/imagen, 2) Con monto, 3) Primer mensaje (más antiguo)
 */
function encontrarMensajeComprobante(messages) {
  // Prioridad 1: Mensaje con adjuntos de comprobante (PDF, imagen)
  const conAdjuntos = messages.find(m => {
    const adjuntos = m.getAttachments();
    return adjuntos.some(a => /\.(pdf|jpg|jpeg|png|gif)$/i.test(a.getName()));
  });
  if (conAdjuntos) {
    log(`  → Seleccionado: mensaje con adjuntos (de ${messages.length} en el thread)`);
    return conAdjuntos;
  }
  
  // Prioridad 2: Mensaje con monto explícito ($XXX.XXX)
  const conMonto = messages.find(m => {
    const cuerpo = m.getPlainBody();
    return /\$\s*[\d.,]+|\d{1,3}(?:\.\d{3})+(?:,\d{2})?/.test(cuerpo);
  });
  if (conMonto) {
    log(`  → Seleccionado: mensaje con monto (de ${messages.length} en el thread)`);
    return conMonto;
  }
  
  // Prioridad 3: Primer mensaje (más antiguo) - quien inicia suele enviar el comprobante
  log(`  → Seleccionado: primer mensaje del thread (de ${messages.length} en el thread)`);
  return messages[0];
}

// ==================== CLASIFICACIÓN ====================
function esComprobanteExpensa(message) {
  const asunto = (message.getSubject() || '').toLowerCase();
  const cuerpo = (message.getPlainBody() || '').toLowerCase();
  const remitente = (message.getFrom() || '').toLowerCase();
  const textoCompleto = `${asunto} ${cuerpo}`;
  const tieneAdjuntos = message.getAttachments().length > 0;

  let puntuacion = 0;
  let criteriosCumplidos = [];

  // ===== EXCLUSIÓN: EMAILS SALIENTES DE ESTA CUENTA =====
  const esEmailSaliente = CONFIG.EMAILS_ORIGEN.some(email => remitente.includes(email.split('@')[0]));
  if (esEmailSaliente) {
    if (CONFIG.DEBUG) {
      log(`Clasificación: ${message.getSubject()}`);
      log(`  DESCARTADO: Email saliente de esta cuenta`);
    }
    return false;
  }

  // ===== EXCLUSIONES FUERTES (descarta inmediatamente) =====
  const contieneExclusionFuerte = PALABRAS_EXCLUSION_FUERTE.some(p => textoCompleto.includes(p));
  if (contieneExclusionFuerte) {
    if (CONFIG.DEBUG) {
      log(`Clasificación: ${message.getSubject()}`);
      log(`  DESCARTADO: Contiene frase de exclusión fuerte`);
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
    log(`Clasificación: ${message.getSubject()}`);
    log(`  Remitente: ${remitente}`);
    log(`  Puntuación: ${puntuacion}`);
    log(`  Criterios: ${criteriosCumplidos.join(', ')}`);
  }

  // ===== UMBRAL DINÁMICO =====
  const UMBRAL_MINIMO = 5;
  const UMBRAL_SIN_EXPENSA = 8; // Umbral más alto si no menciona "expensas"

  // Si no menciona "expensas" y tiene puntuación baja-media, descartar
  if (!tieneExpensa && puntuacion < UMBRAL_SIN_EXPENSA) {
    if (CONFIG.DEBUG) {
      log(`  DESCARTADO: No menciona expensas y puntuación insuficiente (${puntuacion} < ${UMBRAL_SIN_EXPENSA})`);
    }
    return false;
  }

  return puntuacion >= UMBRAL_MINIMO;
}

// ==================== REENVÍO ====================
/**
 * Reenvía un email a la casilla destino
 * @param {GmailMessage} message - El mensaje a reenviar
 * @param {GmailAttachment[]} adjuntosExtra - (Opcional) Adjuntos adicionales del thread
 */
function reenviarEmail(message, adjuntosExtra) {
  try {
    const asunto = message.getSubject();
    const cuerpo = message.getPlainBody();
    
    // Usar adjuntos del thread completo si se proporcionan, sino solo del mensaje
    const adjuntos = adjuntosExtra && adjuntosExtra.length > 0 
      ? adjuntosExtra 
      : message.getAttachments();

    const fecha = Utilities.formatDate(
      message.getDate(),
      'America/Argentina/Buenos_Aires',
      'dd/MM/yyyy HH:mm'
    );

    // Preparar opciones con adjuntos (si existen)
    const options = {};
    if (adjuntos.length > 0) {
      options.attachments = adjuntos;
    }

    // Crear draft con adjuntos como parámetro de opciones
    let emailReenviado = GmailApp.createDraft(
      CONFIG.EMAIL_DESTINO,
      `[REENVIADO] ${asunto}`,
      `Este email fue reenviado automáticamente desde: ${message.getFrom()}\n\n` +
        `Fecha original: ${fecha}\n` +
        `Asunto original: ${asunto}\n\n` +
        `--- CONTENIDO ORIGINAL ---\n${cuerpo}`,
      options
    );

    emailReenviado.send();
    log(`✓ Email reenviado exitosamente: ${asunto}` + (adjuntos.length > 0 ? ` (con ${adjuntos.length} adjuntos)` : ''));
  } catch (error) {
    log(`✗ Error reenviando email: ${error.toString()}`);
    throw error;
  }
}

// ==================== UTILIDADES ====================
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

function getDiscardStateMonthKey_() {
  return 'DESCARTADO_STATE_MONTH';
}

function getDiscardLastMsgPrefix_() {
  return 'DESCARTADO_LASTMSG_';
}

function cleanupDiscardStateIfNeeded_() {
  const props = PropertiesService.getScriptProperties();
  const current = getMonthKey_();
  const last = props.getProperty(getDiscardStateMonthKey_());
  if (last === current) return;

  const all = props.getProperties();
  const prefix = getDiscardLastMsgPrefix_();
  Object.keys(all).forEach(key => {
    if (key.indexOf(prefix) === 0) {
      props.deleteProperty(key);
    }
  });
  props.setProperty(getDiscardStateMonthKey_(), current);
}

function getDiscardLastMsgKey_(threadId) {
  return getDiscardLastMsgPrefix_() + threadId;
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

function hasRelevantAttachment_(message) {
  const attachments = message.getAttachments({ includeInlineImages: true });
  if (!attachments || attachments.length === 0) return false;

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
      if (size < MIN_INLINE_IMAGE_BYTES) continue;
    }
    return true;
  }

  return false;
}

function reabrirDescartadosConAdjunto_(etiquetaDescartado) {
  const query = `in:inbox label:${CONFIG.ETIQUETA_DESCARTADO} ${getMonthStartQuery_()}`;
  const threads = GmailApp.search(query, 0, 50);
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

function enviarResumenEjecucion(procesados, reenviados, errores) {
  const asunto = `Resumen Procesamiento Expensas - ${new Date().toLocaleDateString()}`;
  const cuerpo =
    `Resumen de procesamiento automático de comprobantes de expensas:\n\n` +
    `- Emails procesados: ${procesados}\n` +
    `- Comprobantes reenviados: ${reenviados}\n` +
    `- Errores: ${errores}\n\n` +
    `Fecha: ${new Date().toLocaleString()}\n`;

  MailApp.sendEmail(CONFIG.EMAIL_DESTINO, asunto, cuerpo);
}

// ==================== CONFIGURACIÓN DE TRIGGERS ====================
function configurarTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'procesarEmailsExpensas') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('procesarEmailsExpensas').timeBased().everyMinutes(15).create();

  log('Trigger configurado: ejecución cada 15 minutos');
}

function ejecutarPrueba() {
  log('=== MODO PRUEBA ===');
  CONFIG.DEBUG = true;

  const emailOrigen = CONFIG.EMAILS_ORIGEN[0];
  try {
    const query = `in:inbox newer_than:7d`;
    const threads = GmailApp.search(query, 0, 5);

    log(`\n--- Prueba en cuenta actual (origen lógico: ${emailOrigen}) ---`);
    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        const esExpensa = esComprobanteExpensa(message);
        log(`Asunto: ${message.getSubject()}`);
        log(`Resultado: ${esExpensa ? '✓ ES EXPENSA' : '✗ NO ES EXPENSA'}\n`);
      });
    });
  } catch (error) {
    log(`Error en prueba: ${error.toString()}`);
  }
}
