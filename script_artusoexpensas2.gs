/**
 * Script para la casilla CENTRAL: artusoexpensas2@gmail.com
 * - Aquí puedes decidir si solo etiquetas expensas o también las reenvías a otro sistema.
 * - Versión por defecto: SOLO etiqueta los mails de expensas (no los reenvía para evitar duplicados).
 */

// ==================== CONFIGURACIÓN ====================
const CONFIG = {
  EMAILS_ORIGEN: ['artusoexpensas2@gmail.com'],
  EMAIL_DESTINO: 'artusoexpensas2@gmail.com', // se mantiene igual, pero NO se usa para reenviar
  ETIQUETA_PROCESADO: 'ExpensaDetectada',
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

/**
 * Determina el separador de argumentos de fórmula según el locale.
 */
function getFormulaSeparator(locale) {
  if (!locale) return ',';
  const normalized = locale.toLowerCase().split('_')[0];
  return LOCALES_SEMICOLON.includes(normalized) ? ';' : ',';
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
    'expensa', 'expensas', 'expesnas', // incluye typo común
    'gasto común', 'gastos comunes',
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
  'confirmamos recepción'
];

// ==================== PROMPT DE VALIDACIÓN LLM ====================
const VALIDATION_PROMPT = `Clasifica este email sobre expensas en 3 estados:
- "accept" (pago de expensas)
- "reject" (NO es pago de expensas)
- "uncertain" (no hay evidencia suficiente; si hay duda, usar uncertain)

DEVUELVE SOLO este JSON válido (sin texto extra):
{"decision":"accept|reject|uncertain","reason":"texto breve"}

──────────────────────────────────────────────
EVIDENCIA DE PAGO (aceptar):
- Pago ya realizado / transferencia realizada / se acreditó / pago recibido
- Comprobante adjunto o indicado en texto ("adjunto comprobante", "voucher", "constancia")
- Motivo/Concepto: EXPENSAS
- Datos de transferencia: número de operación, CBU/CVU/alias, importe
- Notificaciones bancarias o de cobro ("Recibiste un pago", "Se realizó la transferencia") con contexto de expensas/consorcio
- Estado APROBADO en sistemas de expensas
- Pagos mixtos (expensas + cochera/otros)

EVIDENCIA NEGATIVA (rechazar SOLO si ves alguna de estas):
- Solicitud/consulta de boleta/liquidación/deuda: "solicitar expensas", "no me han llegado", "reenviar boleta", "cuánto debo", "monto a pagar"
- Factura/Proveedor/Presupuesto/TAD/Publicidad/Newsletter/Acta/Reclamo/Consulta no relacionada a pago
- Comunicaciones administrativas sin confirmación de pago (convocatorias, actas, avisos)
- Pagos ajenos al consorcio

INCERTIDUMBRE (usar "uncertain" si):
- Falta evidencia suficiente de pago y no hay evidencia negativa clara
- Solo adjuntos/imagenes sin texto que confirme pago
- Cuerpo vacío o muy escaso, sin señales claras de pago ni de negativa

REGLA:
- reject: SOLO con evidencia negativa clara.
- accept: si hay evidencia de pago/expensas.
- uncertain: si no hay evidencia suficiente y tampoco negativa clara (ante la duda, usa uncertain).

REGLA ESPECIAL (IMPORTANTE):
- Si el email tiene evidencia clara de PAGO/TRANSFERENCIA (por ejemplo: menciona transferencia realizada, importe, CBU/CVU/alias, número de operación, \"recibiste un pago\", \"se realizó la transferencia\")
  PERO no menciona explícitamente expensas/consorcio/motivo expensas, NO uses \"reject\".
  En ese caso debes responder \"uncertain\" (y el sistema lo registrará como REVIEW para verificación humana).

Formato obligatorio:
{"decision":"accept|reject|uncertain","reason":"texto breve"}`;

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
  const truncatedBody = body.substring(0, 2000);
  const prompt = VALIDATION_PROMPT + `\n\nEmail:\nAsunto: ${subject}\nCuerpo: ${truncatedBody}`;
  
  try {
    const response = callAI(prompt);
    // Limpiar respuesta (a veces la IA agrega backticks de markdown)
    const cleanResponse = response.replace(/```json\n?|\n?```/g, '').trim();
    const json = JSON.parse(cleanResponse);
    const decision = (json.decision || '').toString().toLowerCase().trim();
    const reason = (json.reason || '').toString().trim();
    const shortReason = reason.length > 120 ? reason.substring(0, 120) : reason;

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

/**
 * Extrae datos estructurados del email usando IA
 * @param {string} subject - Asunto del email
 * @param {string} body - Cuerpo del email
 * @returns {Object} - {monto, fecha_pago, dpto, estado}
 */
function extractDataWithAI(subject, body) {
  const prompt = `Extrae la siguiente informacion de este email de pago de expensas.
Responde UNICAMENTE con un JSON valido, sin explicaciones ni texto adicional.

Campos a extraer:
- monto: numero sin simbolo $ (ejemplo: 164185.30). Si no encuentras monto, usa null.
- fecha_pago: formato DD-MM-YY. Si no hay fecha de pago explicita en el comprobante, usa null.
- dpto: departamento/unidad/piso. Ejemplos de formatos válidos:
  - "5A", "PB B", "8 PISO", "LOC" (formatos simples)
  - "Piso 5 Dpto 11" (si el texto dice "piso 5 dpto 11", "p 5 dpto 11", "piso 5 departamento 11", etc.)
  - "Cochera 10" o "Cocheras 10 y 15" (si el texto menciona cocheras)
  - Normalizar siempre a "Piso X Dpto Y" cuando haya piso y número de departamento separados.
  - Normalizar siempre a "Cochera X" o "Cocheras X y Y" cuando se mencionen cocheras.
  Si no hay información de departamento/unidad, usa null.
- estado: string en MAYUSCULAS (ej: "PENDIENTE", "APROBADO") si el email contiene un campo tipo "Estado: ...". Si no aparece, usa null. Si aparece "Estado: Pendiente", devuelve "PENDIENTE".
- ed: direccion del edificio/consorcio (solo calle y numero, sin piso ni depto). Ejemplos: "Paraguay 2949", "Av Santa Fe 2647/51", "Araoz 380". Si no hay dirección, usa null.

Email:
Asunto: ${subject}
Cuerpo: ${body}

JSON:`;

  const response = callAI(prompt);
  // Limpiar respuesta (a veces la IA agrega backticks de markdown)
  const cleanResponse = response.replace(/```json\n?|\n?```/g, '').trim();
  return JSON.parse(cleanResponse);
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
function saveToSheet(noticeDate, paymentDate, amount, building, apartment, threadUrl, label) {
  const ss = SpreadsheetApp.openById(SHEET_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`No se encontró la hoja "${SHEET_CONFIG.SHEET_NAME}" en el spreadsheet`);
  }
  const separator = getFormulaSeparator(ss.getSpreadsheetLocale());
  const safeLabel = (label || '').replace(/"/g, '""');
  const comment = `=HYPERLINK("${threadUrl}"${separator}"${safeLabel}")`;
  log(`(Central) CREANDO OBSERVACIONES con separador "${separator}"`);
  
  // Columnas: TIPO AVISO | FECHA AVISO | FECHA DE PAGO | MONTO | ED | DPTO | UF | COMENTARIO
  const row = [
    '@',                    // TIPO AVISO (siempre email)
    noticeDate,             // FECHA AVISO
    paymentDate || '',      // FECHA DE PAGO
    amount || '',           // MONTO
    building || '',         // ED
    apartment || '',        // DPTO
    '',                     // UF (vacío por ahora)
    comment || ''           // COMENTARIO / OBSERVACIONES
  ];
  
  sheet.appendRow(row);
  log(`(Central) Fila agregada al Sheet: ${JSON.stringify(row)}`);
}

// ==================== FUNCIONES PRINCIPALES ====================

function procesarEmailsExpensas() {
  log('=== Iniciando procesamiento de emails (casilla central) ===');
  let totalProcesados = 0;
  let totalMarcados = 0;
  let totalErrores = 0;

  CONFIG.EMAILS_ORIGEN.forEach(emailOrigen => {
    try {
      const resultado = procesarEmailsDeCuenta(emailOrigen);
      totalProcesados += resultado.procesados;
      totalMarcados += resultado.reenviados; // aquí reenviados = marcados
      totalErrores += resultado.errores;
    } catch (e) {
      log(`Error procesando ${emailOrigen}: ${e.toString()}`);
      totalErrores++;
    }
  });

  log(`=== Resumen central: ${totalProcesados} procesados, ${totalMarcados} marcados, ${totalErrores} errores ===`);
}

function procesarEmailsDeCuenta(emailOrigen) {
  const stats = { procesados: 0, reenviados: 0, errores: 0 };
  try {
    const query = `in:inbox -label:${CONFIG.ETIQUETA_PROCESADO} newer_than:1d`;
    const threads = GmailApp.search(query, 0, 50);
    log(`(Central) Threads nuevos: ${threads.length}`);

    const etiqueta = crearObtenerEtiqueta(CONFIG.ETIQUETA_PROCESADO);

    threads.forEach(thread => {
      const threadUrl = `https://mail.google.com/mail/u/0/#all/${thread.getId()}`;
      try {
        const messages = thread.getMessages();
        messages.forEach(message => {
          stats.procesados++;
          if (esComprobanteExpensa(message)) {
            log(`(Central) ✓ Expensa detectada por keywords: ${message.getSubject()}`);
            
            const subject = message.getSubject();
            const body = message.getPlainBody();

            // Señal simple de "pago claro" (para debugging)
            const texto = `${subject} ${body}`.toLowerCase();
            const pagoClaro = /transferenc|recibiste un pago|pago fue exitoso|se acredit|realizaste/.test(texto) &&
              (/(?:\$\s*)?\d{1,3}(?:[\.\s]\d{3})+(?:,\d{2})?|\$\s*\d{4,}/.test(texto) || /\b(cbu|cvu|alias|nro|número)\b/.test(texto));
            
            // PASO 1: Validar con LLM (fail-open con 3 estados)
            const validation = validateExpensePaymentWithLLM(subject, body);
            log(`(Central) Validator: decision=${validation.decision || 'unknown'} review=${validation.review ? 'yes' : 'no'} reason=${validation.reviewReason || ''} pago_claro=${pagoClaro ? 'yes' : 'no'}`);
            
            if (!validation.isValid) {
              log(`(Central) ✗ LLM rejected: ${subject} | Reason: ${validation.reviewReason || 'LLM_REJECT'}`);
              // No etiquetar ni procesar - el LLM determinó que no es un pago válido
              return;
            }
            
            // PASO 2: Extraer datos con IA y guardar en Sheet
            try {
              const noticeDate = Utilities.formatDate(
                message.getDate(), 
                'America/Argentina/Buenos_Aires', 
                'dd-MM-yy'
              );
              
              const extractedData = extractDataWithAI(subject, body);
              const estado = (extractedData.estado || '').toString().trim().toUpperCase();
              
              const commentParts = [];
              if (validation.review) {
                commentParts.push(`REVIEW: ${validation.reviewReason || 'LLM_UNCERTAIN'}`);
              }
              if (estado === 'PENDIENTE') {
                commentParts.push('ESTADO: PENDIENTE');
              }
              const labelSuffix = commentParts.length > 0 ? ` (${commentParts.join(' | ')})` : '';
              const rawLabel = `Abrir email${labelSuffix}`;
              
              saveToSheet(
                noticeDate, 
                extractedData.fecha_pago, 
                extractedData.monto, 
                extractedData.ed,
                extractedData.dpto,
                threadUrl,
                rawLabel
              );
              
              log(`(Central) Datos extraídos: ${JSON.stringify(extractedData)}`);
            } catch (extractError) {
              log(`(Central) Error extrayendo datos con IA: ${extractError.toString()}`);
              // Continúa aunque falle la extracción - el email ya fue validado como expensa
            }
            
            thread.addLabel(etiqueta);
            stats.reenviados++; // usamos este campo como "marcados"
          } else {
            log(`(Central) ✗ NO es expensa (keywords): ${message.getSubject()}`);
          }
        });
      } catch (e) {
        log(`(Central) Error en thread: ${e.toString()}`);
        stats.errores++;
      }
    });
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
  const esRespuestaAgradecimiento = PALABRAS_EXCLUSION_FUERTE.some(p => textoCompleto.includes(p));
  if (esRespuestaAgradecimiento) {
    if (CONFIG.DEBUG) {
      log(`(Central) Clasificación: ${message.getSubject()}`);
      log(`  DESCARTADO: Es respuesta/agradecimiento, no pago nuevo`);
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
      
      log(`--- Email ${index + 1}: ${subject} ---`);
      
      try {
        const validation = validateExpensePaymentWithLLM(subject, body);
        log(`Validador: decision=${validation.isValid ? (validation.review ? 'ACCEPT_WITH_REVIEW' : 'ACCEPT') : 'REJECT'} | review=${validation.review ? (validation.reviewReason || 'LLM_UNCERTAIN') : 'none'}`);
        const extractedData = extractDataWithAI(subject, body);
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
          const isValidByLLM = validateExpensePaymentWithLLM(subject, body);
          log(`  LLM: ${isValidByLLM ? 'CONFIRMADO' : 'RECHAZADO'}`);
          
          if (isValidByLLM) {
            passedLLM++;
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
    log(`Rechazados por LLM: ${rejectedLLM}`);
    log(`Tasa de rechazo LLM: ${passedKeywords > 0 ? Math.round(rejectedLLM / passedKeywords * 100) : 0}%`);
    log('=== FIN DE PRUEBA ===');
  } catch (error) {
    log(`Error en prueba de validación LLM: ${error.toString()}`);
  }
}
