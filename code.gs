/**
 * បង្ហាញ Dashboard ដោយផ្ទាល់ (មាន Dashboard + All Records បញ្ចូលគ្នាក្នុងឯកសារតែមួយ)
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Ticket Cancellation Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * គណនាចំនួនទឹកប្រាក់ពិន័យ (Penalty) អាស្រ័យលើរយៈពេលលុបចោល
 * Under 10 Minute = 10% នៃ amount
 * Under 30 Minute = 30% នៃ amount
 * Up 30 Minute (Over 30 Minute) = 60% នៃ amount
 */
function calculatePenalty(amount, duration) {
  if (!duration) return 0;
  const d = String(duration);
  if (d.indexOf('Under 10') !== -1) return amount * 0.10;
  if (d.indexOf('Under 30') !== -1) return amount * 0.30;
  if (d.indexOf('Up 30') !== -1 || d.indexOf('Over 30') !== -1) return amount * 0.60;
  return 0;
}

/**
 * ទាញទិន្នន័យពី Sheet ម្តងទុកក្នុង Browser សម្រាប់ឱ្យចុច Chart រត់ឆ្លើយឆ្លងគ្នាភ្លាមៗ (Zero Delay)
 */
function getRawDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Phone call') || ss.getSheets()[0];
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) return { records: [] };

  const headers = values[0].map(h => String(h).trim().toLowerCase());

  const colIndex = {
    ticket: headers.indexOf('ticket code'),
    date: headers.indexOf('travel date') !== -1 ? headers.indexOf('travel date') : headers.indexOf('cancelled date'),
    direction: headers.indexOf('direction'),
    amount: headers.indexOf('amount'),
    agent: headers.indexOf('cancelled by'),
    durationGroup: headers.indexOf('count'),
    reason: headers.indexOf('reason'),
    status: headers.indexOf('status')
  };

  const records = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[colIndex.ticket] && !row[colIndex.agent]) continue;

    // សម្អាតកាលបរិច្ឆេទ YYYY-MM-DD
    let rawDate = row[colIndex.date];
    let isoDate = '';
    if (rawDate instanceof Date && !isNaN(rawDate)) {
      isoDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
    } else if (rawDate) {
      const parsed = new Date(rawDate);
      if (!isNaN(parsed)) {
        isoDate = Utilities.formatDate(parsed, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
      }
    }

    // សម្អាតចំនួនទឹកប្រាក់
    let rawAmount = row[colIndex.amount];
    let numAmount = 0;
    if (typeof rawAmount === 'number') {
      numAmount = rawAmount;
    } else if (rawAmount) {
      numAmount = parseFloat(String(rawAmount).replace(/[^0-9.-]+/g, '')) || 0;
    }

    // សម្អាតឈ្មោះខ្សែផ្លូវ
    let dir = String(row[colIndex.direction] || 'Unknown').trim();
    let simplifiedRoute = dir.split('(')[0].trim() || dir;
    simplifiedRoute = simplifiedRoute.replace(/\s+/g, ' ');

    const durationStr = String(row[colIndex.durationGroup] || 'Other').trim();
    const statusVal = colIndex.status !== -1 ? String(row[colIndex.status] || '').trim() : '';

    records.push({
      rowIndex: i + 1, // លេខជួរដេកពិតប្រាកដក្នុង Sheet (Header = ជួរទី 1)
      ticket: String(row[colIndex.ticket] || ''),
      date: isoDate,
      route: simplifiedRoute,
      amount: numAmount,
      agent: String(row[colIndex.agent] || 'Unknown').trim(),
      duration: durationStr,
      reason: String(row[colIndex.reason] || '-'),
      penalty: calculatePenalty(numAmount, durationStr),
      status: statusVal || 'Pending'
    });
  }

  return { records: records };
}

/**
 * ធ្វើបច្ចុប្បន្នភាព Status (Approved / Rejected) ត្រឡប់ទៅ Sheet វិញ
 * បង្កើតជួរឈរ "Status" ដោយស្វ័យប្រវត្តិ ប្រសិនបើមិនទាន់មាន
 * ប្រសិនបើបើក "Notify on Approve/Reject" នឹងផ្ញើសារទៅ Telegram ផងដែរ
 */
function updateRecordStatus(rowIndex, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Phone call') || ss.getSheets()[0];

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  let statusCol = headers.indexOf('status');

  if (statusCol === -1) {
    statusCol = lastCol + 1; // បន្ថែមជួរឈរថ្មីខាងចុង
    sheet.getRange(1, statusCol).setValue('Status');
  } else {
    statusCol = statusCol + 1; // បម្លែងទៅ 1-based
  }

  sheet.getRange(rowIndex, statusCol).setValue(status);

  // ផ្ញើ Telegram (បើបើកដំណើរការ) - កំហុស Telegram មិនប៉ះពាល់ដល់ការរក្សាទុក Status ទេ
  try {
    notifyDecision_(sheet, headers, rowIndex, status);
  } catch (e) {
    console.error('Telegram notify failed: ' + e.message);
  }

  return { success: true, rowIndex: rowIndex, status: status };
}

/* =====================================================================
 *                         TELEGRAM BOT SERVICE
 * Token និង Chat ID រក្សាទុកក្នុង Script Properties (នៅ Server)
 * មិនដែលបញ្ជូនត្រឡប់ទៅ Browser ជា Token ពេញលេញទេ (បង្ហាញតែ Masked)
 * ===================================================================== */

const TG_KEY_TOKEN  = 'TG_BOT_TOKEN';
const TG_KEY_CHAT   = 'TG_CHAT_ID';
const TG_KEY_NOTIFY = 'TG_NOTIFY_DECISION';

function escHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function maskToken_(token) {
  if (!token) return '';
  const parts = String(token).split(':');
  const tail = String(token).slice(-4);
  return parts[0] + ':••••••••' + tail;
}

/** ហៅ Telegram Bot API */
function tgApi_(token, method, payload) {
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true
  });
  try {
    return JSON.parse(res.getContentText());
  } catch (e) {
    return { ok: false, description: 'Invalid response from Telegram (HTTP ' + res.getResponseCode() + ')' };
  }
}

/** ផ្ញើសារ HTML ទៅ Chat ដែលបានរក្សាទុក */
function sendTelegramMessage_(text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(TG_KEY_TOKEN);
  const chatId = props.getProperty(TG_KEY_CHAT);
  if (!token || !chatId) throw new Error('Telegram is not connected yet.');

  const data = tgApi_(token, 'sendMessage', {
    chat_id: chatId,
    text: text,
    parse_mode: 'HTML',
    disable_web_page_preview: true
  });
  if (!data.ok) throw new Error(data.description || 'Telegram sendMessage failed');
  return true;
}

/** ទាញស្ថានភាពការតភ្ជាប់ (Token ត្រូវបាន Mask) */
function getTelegramConfig() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(TG_KEY_TOKEN) || '';
  const chatId = props.getProperty(TG_KEY_CHAT) || '';
  return {
    connected: !!(token && chatId),
    tokenMasked: maskToken_(token),
    hasToken: !!token,
    chatId: chatId,
    notifyDecision: props.getProperty(TG_KEY_NOTIFY) === 'true'
  };
}

/**
 * ផ្ទៀងផ្ទាត់ Token (getMe) + ផ្ញើសារសាកល្បង រួចទើបរក្សាទុក
 * ប្រសិនបើ token ទទេ ហើយធ្លាប់មាន Token ចាស់ → ប្រើ Token ចាស់
 */
function saveTelegramConfig(token, chatId, notifyDecision) {
  const props = PropertiesService.getScriptProperties();
  token = String(token || '').trim() || props.getProperty(TG_KEY_TOKEN) || '';
  chatId = String(chatId || '').trim();

  if (!token) throw new Error('Please enter the Bot Token.');
  if (!/^\d+:[A-Za-z0-9_-]{30,}$/.test(token)) {
    throw new Error('Bot Token format looks wrong. Example: 123456789:AAH...');
  }
  if (!chatId) throw new Error('Please enter the Chat ID.');
  if (!/^-?\d+$/.test(chatId) && !/^@[A-Za-z0-9_]{5,}$/.test(chatId)) {
    throw new Error('Chat ID must be a number (e.g. 123456789 or -1001234567890) or @channelname.');
  }

  const me = tgApi_(token, 'getMe', {});
  if (!me.ok) throw new Error('Invalid token: ' + (me.description || 'Telegram rejected it'));

  const test = tgApi_(token, 'sendMessage', {
    chat_id: chatId,
    text: '✅ <b>Ticket Cancellation Dashboard</b> connected successfully.',
    parse_mode: 'HTML'
  });
  if (!test.ok) {
    throw new Error('Cannot send to this Chat ID: ' + (test.description || 'unknown error') +
      '. Make sure you pressed Start on the bot (or added it to the group/channel).');
  }

  props.setProperties({
    [TG_KEY_TOKEN]: token,
    [TG_KEY_CHAT]: chatId,
    [TG_KEY_NOTIFY]: notifyDecision ? 'true' : 'false'
  });

  return { success: true, botName: me.result.first_name, botUsername: me.result.username };
}

/** ធ្វើបច្ចុប្បន្នភាពតែ Notify option (មិនប៉ះ Token/Chat) */
function setTelegramNotify(flag) {
  PropertiesService.getScriptProperties().setProperty(TG_KEY_NOTIFY, flag ? 'true' : 'false');
  return { success: true };
}

/** លុបការតភ្ជាប់ */
function disconnectTelegram() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(TG_KEY_TOKEN);
  props.deleteProperty(TG_KEY_CHAT);
  props.deleteProperty(TG_KEY_NOTIFY);
  return { success: true };
}

/**
 * រក Chat ID ដោយស្វ័យប្រវត្តិ ពី getUpdates
 * (អ្នកត្រូវផ្ញើសារ /start ទៅ Bot ជាមុនសិន)
 */
function findTelegramChatId(token) {
  token = String(token || '').trim() || PropertiesService.getScriptProperties().getProperty(TG_KEY_TOKEN) || '';
  if (!token) throw new Error('Enter the Bot Token first.');

  const data = tgApi_(token, 'getUpdates', { limit: 100, timeout: 0 });
  if (!data.ok) throw new Error(data.description || 'Cannot read updates from Telegram.');

  const seen = {};
  const chats = [];
  (data.result || []).forEach(u => {
    const m = u.message || u.edited_message || u.channel_post || u.my_chat_member;
    const c = m && m.chat;
    if (c && !seen[c.id]) {
      seen[c.id] = true;
      const name = c.title || [c.first_name, c.last_name].filter(Boolean).join(' ') || c.username || String(c.id);
      chats.push({ id: String(c.id), name: name, type: c.type });
    }
  });
  return { chats: chats };
}

/** ផ្ញើរបាយការណ៍សង្ខេបពី Dashboard (stats មកពី Browser តាម Filter បច្ចុប្បន្ន) */
function sendTelegramSummary(stats) {
  stats = stats || {};
  const money = n => '$' + Number(n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  let text = '📊 <b>Ticket Cancellation Summary</b>\n';
  if (stats.period) text += '📅 ' + escHtml_(stats.period) + '\n';
  if (stats.filters) text += '🔎 ' + escHtml_(stats.filters) + '\n';
  text += '\n';
  text += '🎫 Total Cancellations: <b>' + Number(stats.total || 0) + '</b>\n';
  text += '💸 Total Lost Value: <b>' + money(stats.amount) + '</b>\n';
  text += '⚠️ Total Penalties: <b>' + money(stats.penalty) + '</b>\n';
  text += '👤 Active Agents: <b>' + Number(stats.agents || 0) + '</b>\n\n';
  text += '🟣 Pending: ' + Number(stats.pending || 0) +
          '   🟢 Approved: ' + Number(stats.approved || 0) +
          '   🔴 Rejected: ' + Number(stats.rejected || 0) + '\n';

  const routes = (stats.topRoutes || []).slice(0, 5);
  if (routes.length) {
    text += '\n🚌 <b>Top 5 Routes</b>\n';
    routes.forEach((r, i) => {
      text += (i + 1) + '. ' + escHtml_(r.name) + ' — <b>' + Number(r.count || 0) + '</b>\n';
    });
  }

  sendTelegramMessage_(text);
  return { success: true };
}

/** សារស្វ័យប្រវត្តិ ពេល Approve / Reject */
function notifyDecision_(sheet, headers, rowIndex, status) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(TG_KEY_NOTIFY) !== 'true') return false;
  if (!props.getProperty(TG_KEY_TOKEN) || !props.getProperty(TG_KEY_CHAT)) return false;

  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const get = name => {
    const i = headers.indexOf(name);
    return i === -1 ? '' : row[i];
  };

  const icon = status === 'Approved' ? '🟢' : '🔴';
  let text = icon + ' <b>' + escHtml_(status) + '</b>\n\n';
  text += '🎫 Ticket: <code>' + escHtml_(get('ticket code')) + '</code>\n';
  text += '🚌 Route: ' + escHtml_(String(get('direction')).split('(')[0].trim()) + '\n';
  text += '💵 Amount: $' + escHtml_(get('amount')) + '\n';
  text += '👤 Agent: ' + escHtml_(get('cancelled by')) + '\n';
  text += '⏱ Duration: ' + escHtml_(get('count'));

  return sendTelegramMessage_(text);
}
/* =====================================================================
 * TELEGRAM PENDING ALERT MENU (ADDED; preserves existing dashboard code)
 * Usage in Telegram: /pending
 * Choose a duration, then choose a date range or enter YYYY-MM-DD to YYYY-MM-DD.
 * Uses existing Script Properties: TG_BOT_TOKEN and TG_CHAT_ID.
 * ===================================================================== */

function doPost(e) {
  try {
    var update = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (update.callback_query) {
      pendingAlertHandleCallback_(update.callback_query);
    } else if (update.message) {
      pendingAlertHandleMessage_(update.message);
    }
  } catch (err) {
    console.error('Pending alert webhook error: ' + (err && err.stack ? err.stack : err));
  }
  return ContentService.createTextOutput('OK');
}

function pendingAlertHandleMessage_(msg) {
  if (!msg || !msg.chat || !msg.chat.id || !msg.text) return;
  var text = String(msg.text).trim();
  var chatId = String(msg.chat.id);
  if (/^\/pending(?:@\w+)?(?:\s|$)/i.test(text) || /^pending alert$/i.test(text)) {
    pendingAlertSendMenu_(chatId);
    return;
  }

  var cache = CacheService.getScriptCache();
  var stateKey = 'PENDING_ALERT_STATE_' + chatId;
  var stateRaw = cache.get(stateKey);
  if (!stateRaw) return;
  var state;
  try { state = JSON.parse(stateRaw); } catch (ignore) { cache.remove(stateKey); return; }
  if (state.awaitingDateRange) {
    var rangeMatch = text.match(/^(\d{4}-\d{2}-\d{2})\s+(?:to|through|-)\s+(\d{4}-\d{2}-\d{2})$/i);
    if (!rangeMatch) {
      pendingAlertSendText_(chatId, 'Please enter a date range like 2026-10-01 to 2026-10-25.');
      return;
    }
    var fromDate = pendingAlertValidateDate_(rangeMatch[1]);
    var toDate = pendingAlertValidateDate_(rangeMatch[2]);
    if (!fromDate || !toDate || rangeMatch[1] > rangeMatch[2]) {
      pendingAlertSendText_(chatId, 'Invalid date range. Make sure both dates are valid and From is not after To.');
      return;
    }
    cache.remove(stateKey);
    pendingAlertSendToConfiguredGroup_(rangeMatch[1], rangeMatch[2], state.category, chatId);
  }
}

function pendingAlertSendMenu_(chatId) {
  var token = pendingAlertGetConfig_().token;
  if (!token) throw new Error('Telegram bot token missing. Save Telegram settings first.');
  var keyboard = { inline_keyboard: [
    [{text:'📋 All Pending Categories', callback_data:'pending:all'}],
    [{text:'🟢 Under 10 Minute', callback_data:'pending:under10'}],
    [{text:'🟡 Under 30 Minute', callback_data:'pending:under30'}],
    [{text:'🔴 30 Minute Up', callback_data:'pending:over30'}]
  ]};
  pendingAlertApi_('sendMessage', {chat_id:chatId, text:'📢 <b>Pending Alert</b>\nChoose a category, then choose a From and To date.', parse_mode:'HTML', reply_markup:keyboard});
}

function pendingAlertHandleCallback_(cq) {
  var token = pendingAlertGetConfig_().token;
  if (token && cq.id) {
    try { pendingAlertApi_('answerCallbackQuery', {callback_query_id:cq.id, text:'Selected'}); } catch (ignore) {}
  }
  var data = String(cq.data || '');
  var chatId = cq.message && cq.message.chat ? String(cq.message.chat.id) : '';
  if (data.indexOf('pendingrange:') === 0) { pendingAlertHandleDateCallback_(cq); return; }
  if (!chatId || data.indexOf('pending:') !== 0) return;
  var category = data.substring('pending:'.length);
  var valid = ['all', 'under10', 'under30', 'over30'];
  if (valid.indexOf(category) === -1) return;
  CacheService.getScriptCache().put('PENDING_ALERT_STATE_' + chatId, JSON.stringify({category:category, direction:'all', awaitingDateRange:true}), 600);
  var keyboard = {inline_keyboard:[
    [{text:'📅 Alert for 2026-10-25 only', callback_data:'pendingrange:2026-10-25:2026-10-25:' + category}],
    [{text:'✏️ Enter From and To dates', callback_data:'pendingrange:custom:custom:' + category}]
  ]};
  pendingAlertApi_('sendMessage', {chat_id:chatId, text:'Choose a date range. For a custom range, enter both dates in one message: YYYY-MM-DD to YYYY-MM-DD.', reply_markup:keyboard});
}

function pendingAlertHandleDateCallback_(cq) {
  var data = String(cq.data || '');
  var chatId = cq.message && cq.message.chat ? String(cq.message.chat.id) : '';
  if (!chatId || data.indexOf('pendingrange:') !== 0) return false;
  var parts = data.split(':');
  var fromDate = parts[1];
  var toDate = parts[2];
  var category = parts[3] || 'all';
  if (fromDate === 'custom' || toDate === 'custom') {
    CacheService.getScriptCache().put('PENDING_ALERT_STATE_' + chatId, JSON.stringify({category:category, direction:'all', awaitingDateRange:true}), 600);
    pendingAlertSendText_(chatId, 'Enter the date range like this: 2026-10-01 to 2026-10-25');
  } else {
    CacheService.getScriptCache().remove('PENDING_ALERT_STATE_' + chatId);
    pendingAlertSendToConfiguredGroup_(fromDate, toDate, category, 'all', chatId);
  }
  return true;
}

function pendingAlertSendToConfiguredGroup_(fromDate, toDate, category, direction, replyChatId) {
  direction = String(direction || 'all').trim();
  var cfg = pendingAlertGetConfig_();
  if (!cfg.token || !cfg.chatId) throw new Error('Telegram token or target group chat ID is missing.');
  var report = pendingAlertBuildReport_(fromDate, toDate, category, direction);
  var messages = report.messages;
  if (!messages.length) messages = ['📢 <b>Pending Alert — ' + pendingAlertEscape_(fromDate) + ' to ' + pendingAlertEscape_(toDate) + '</b>\nNo matching pending records found.'];
  messages.forEach(function(message) {
    pendingAlertApi_('sendMessage', {chat_id:cfg.chatId, text:message, parse_mode:'HTML', disable_web_page_preview:true});
  });
  if (replyChatId && replyChatId !== cfg.chatId) {
    pendingAlertSendText_(replyChatId, '✅ Pending alert sent to the configured Telegram group for ' + fromDate + ' to ' + toDate + '. Matching records: ' + report.count + '.');
  }
}

function pendingAlertBuildReport_(fromDate, toDate, category, direction) {
  direction = String(direction || 'all').trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Phone call') || ss.getSheets()[0];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return {count:0, messages:[]};
  var headers = values[0].map(function(h){return String(h || '').trim().toLowerCase();});
  function col(names) {
    for (var i=0;i<names.length;i++) { var n=headers.indexOf(names[i]); if(n!==-1) return n; }
    return -1;
  }
  var c = {
    ticket:col(['ticket code']), date:col(['travel date','cancelled date','date']), route:col(['direction','route']),
    amount:col(['amount']), agent:col(['cancelled by','agent']), duration:col(['count','duration','cancelled duration']),
    reason:col(['reason']), status:col(['status'])
  };
  if (c.date < 0 || c.duration < 0) throw new Error('Cannot find the date or duration column in the Phone call sheet.');
  var matched = [];
  for (var r=1;r<values.length;r++) {
    var row=values[r];
    if (!row[c.ticket] && (c.agent<0 || !row[c.agent])) continue;
    var rowDate=pendingAlertToIsoDate_(row[c.date]);
    if (!rowDate || rowDate < fromDate || rowDate > toDate) continue;
    var status=c.status<0 ? 'Pending' : String(row[c.status] || 'Pending').trim();
    if (!/^pending$/i.test(status)) continue;
    var dur=String(row[c.duration] || '').toLowerCase();
    var kind=dur.indexOf('under 10')>=0 ? 'under10' : (dur.indexOf('under 30')>=0 ? 'under30' : ((dur.indexOf('up 30')>=0 || dur.indexOf('over 30')>=0 || dur.indexOf('30 minute up')>=0) ? 'over30' : 'other'));
    if (category !== 'all' && kind !== category) continue;
    var routeText = c.route < 0 ? '' : String(row[c.route] || '').trim();
    if (direction !== 'all' && routeText.toLowerCase() !== direction.toLowerCase()) continue;
    var amount=c.amount<0 ? '' : row[c.amount];
    var route=c.route<0 ? '' : row[c.route];
    var agent=c.agent<0 ? '' : row[c.agent];
    var reason=c.reason<0 ? '' : row[c.reason];
    matched.push({ticket:c.ticket<0?'':row[c.ticket], route:route, amount:amount, agent:agent, duration:row[c.duration], reason:reason, row:r+1});
  }
  var catLabel={all:'All Pending Categories',under10:'Under 10 Minute',under30:'Under 30 Minute',over30:'30 Minute Up'}[category] || 'All Pending Categories';
  var header='➡️ <b>Direction: '+pendingAlertEscape_(direction === 'all' ? 'All Directions' : direction)+'</b>\n📢 <b>Pending Alert</b>\n📅 Date range: <b>'+pendingAlertEscape_(fromDate)+' to '+pendingAlertEscape_(toDate)+'</b>\n⏱ Category: <b>'+pendingAlertEscape_(catLabel)+'</b>\n📋 Pending records: <b>'+matched.length+'</b>\n\n';
  var messages=[]; var chunk=header;
  matched.forEach(function(x,i){
    var item=(i+1)+'. 🎫 <b>'+pendingAlertEscape_(x.ticket || ('Row '+x.row))+'</b>\n'+
      '🚌 Route: '+pendingAlertEscape_(x.route || '-')+'\n'+
      '💵 Amount: '+pendingAlertEscape_(x.amount === '' ? '-' : x.amount)+'\n'+
      '👤 Agent: '+pendingAlertEscape_(x.agent || '-')+'\n'+
      '⏱ Duration: '+pendingAlertEscape_(x.duration || '-')+'\n'+
      '📝 Reason: '+pendingAlertEscape_(x.reason || '-')+'\n\n';
    if ((chunk+item).length > 3800) { messages.push(chunk); chunk='📢 <b>Pending Alert — continued</b>\n📅 '+pendingAlertEscape_(fromDate)+' to '+pendingAlertEscape_(toDate)+'\n\n'; }
    chunk+=item;
  });
  if (chunk.trim()) messages.push(chunk);
  return {count:matched.length,messages:messages};
}

function pendingAlertToIsoDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
  if (!value) return '';
  var s=String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var d=new Date(value);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
}

function pendingAlertGetConfig_() {
  var p=PropertiesService.getScriptProperties();
  return {token:p.getProperty(TG_KEY_TOKEN) || '', chatId:p.getProperty(TG_KEY_CHAT) || ''};
}

function pendingAlertApi_(method, payload) {
  var token=pendingAlertGetConfig_().token;
  if (!token) throw new Error('Telegram bot token is not configured.');
  var data=tgApi_(token, method, payload);
  if (!data || !data.ok) throw new Error('Telegram '+method+' failed: '+((data && data.description) || 'unknown error'));
  return data.result;
}

function pendingAlertSendText_(chatId, text) {
  return pendingAlertApi_('sendMessage', {chat_id:chatId, text:text, disable_web_page_preview:true});
}

function pendingAlertEscape_(value) {
  return String(value == null ? '' : value).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

/** Manual test: sends the requested date/category report to configured Telegram group. */
function testPendingAlertFor20261025() {
  pendingAlertSendToConfiguredGroup_('2026-10-25', '2026-10-25', 'all', 'all', '');
}

function getPendingAlertDirections() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Phone call') || ss.getSheets()[0];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return ['All Directions'];
  var headers = values[0].map(function(h){return String(h || '').trim().toLowerCase();});
  var idx = headers.indexOf('direction'); if (idx < 0) idx = headers.indexOf('route');
  if (idx < 0) return ['All Directions'];
  var seen = {}; values.slice(1).forEach(function(row){var v=String(row[idx]||'').trim(); if(v) seen[v]=true;});
  return ['All Directions'].concat(Object.keys(seen).sort());
}


/** Backward-compatible single-date dashboard action. */
function sendPendingAlertByDate(dateIso, category) {
  return sendPendingAlertByDateRange(dateIso, dateIso, category, 'all');
}

/** Dashboard action: send pending alert for an inclusive From/To date range. */
function sendPendingAlertByDateRange(fromDate, toDate, category, direction) {
  fromDate = String(fromDate || '').trim();
  toDate = String(toDate || '').trim();
  category = String(category || 'all').trim();
  direction = String(direction || 'all').trim();
  if (!pendingAlertValidateDate_(fromDate) || !pendingAlertValidateDate_(toDate)) throw new Error('Please choose valid From and To dates.');
  if (fromDate > toDate) throw new Error('From date cannot be after To date.');
  if (['all','under10','under30','over30'].indexOf(category) === -1) throw new Error('Invalid pending category.');
  var report = pendingAlertBuildReport_(fromDate, toDate, category, direction);
  var cfg = pendingAlertGetConfig_();
  if (!cfg.token || !cfg.chatId) throw new Error('Telegram is not connected. Open Telegram settings and save the bot token and target group chat ID.');
  var messages = report.messages;
  if (!messages.length) messages = ['📢 <b>Pending Alert — ' + pendingAlertEscape_(fromDate) + ' to ' + pendingAlertEscape_(toDate) + '</b>\nNo matching pending records found.'];
  messages.forEach(function(message) { pendingAlertApi_('sendMessage', {chat_id:cfg.chatId, text:message, parse_mode:'HTML', disable_web_page_preview:true}); });
  return {success:true, count:report.count, fromDate:fromDate, toDate:toDate, category:category, direction:direction};
}

function pendingAlertValidateDate_(dateIso) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateIso || ''))) return false;
  var d = new Date(dateIso + 'T12:00:00');
  return !isNaN(d.getTime()) && Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd') === dateIso;
}


/** Register this deployed web app as the Telegram webhook for /pending and inline buttons. */
function connectPendingAlertWebhook() {
  var cfg = pendingAlertGetConfig_();
  if (!cfg.token) throw new Error('Save the Telegram bot token first.');
  var url = ScriptApp.getService().getUrl();
  if (!url || !/\/exec(?:$|\?)/.test(url)) throw new Error('Deploy this project as a Web App first, then run this function again.');
  var result = pendingAlertApi_('setWebhook', {url:url});
  return {success:true, url:url, description:'Telegram webhook connected.'};
}


/** Schedule a one-time pending alert. Uses Apps Script time-driven trigger. */
function schedulePendingAlert(fromDate, toDate, category, sendAt, direction) {
  fromDate=String(fromDate||'').trim(); toDate=String(toDate||'').trim(); category=String(category||'all').trim(); direction=String(direction||'all').trim(); sendAt=String(sendAt||'').trim();
  if (!pendingAlertValidateDate_(fromDate)||!pendingAlertValidateDate_(toDate)||fromDate>toDate) throw new Error('Please choose a valid date range.');
  if (['all','under10','under30','over30'].indexOf(category)<0) throw new Error('Invalid category.');
  var when=new Date(sendAt); if (!sendAt || isNaN(when.getTime()) || when.getTime()<=Date.now()) throw new Error('Choose a future automatic send date and time.');
  var props=PropertiesService.getScriptProperties();
  ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='runScheduledPendingAlert_') ScriptApp.deleteTrigger(t);});
  props.setProperty('PENDING_ALERT_SCHEDULE', JSON.stringify({fromDate:fromDate,toDate:toDate,category:category,direction:direction,sendAt:sendAt}));
  ScriptApp.newTrigger('runScheduledPendingAlert_').timeBased().at(when).create();
  return {success:true,fromDate:fromDate,toDate:toDate,category:category,direction:direction,sendAt:sendAt,timezone:Session.getScriptTimeZone()};
}
function runScheduledPendingAlert_() {
  var props=PropertiesService.getScriptProperties(); var raw=props.getProperty('PENDING_ALERT_SCHEDULE'); if(!raw) return;
  props.deleteProperty('PENDING_ALERT_SCHEDULE');
  var cfg=pendingAlertGetConfig_(); if(!cfg.token||!cfg.chatId) throw new Error('Telegram token or chat ID is missing.');
  var x=JSON.parse(raw); var report=pendingAlertBuildReport_(x.fromDate,x.toDate,x.category,x.direction||'all');
  var messages=report.messages; if(!messages.length) messages=['➡️ <b>Direction: '+pendingAlertEscape_(x.direction==='all'?'All Directions':x.direction||'All Directions')+'</b>\n📢 <b>Pending Alert</b>\nNo matching pending records found.'];
  messages.forEach(function(m){pendingAlertApi_('sendMessage',{chat_id:cfg.chatId,text:m,parse_mode:'HTML',disable_web_page_preview:true});});
}

