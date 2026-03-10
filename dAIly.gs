/**
 * dAIly(デイリー) - メインスクリプト
 * * LINEから送られた曖昧な予定をGemini APIで解析し、
 * Googleカレンダーへの登録とスプレッドシート管理による
 * 精度の高いリマインド通知を自動で行います。
 * * @author Honami Kanno
 * @license MIT
 */

// --- 設定情報を取得 ---
const props = PropertiesService.getScriptProperties();
const LINE_ACCESS_TOKEN = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
const CALENDAR_ID = props.getProperty('CALENDAR_ID') || 'primary';
const SHEET_ID = props.getProperty('SHEET_ID');

// LINEからユーザー名を取得する
function getDisplayName(userId) {
  try {
    const url = `https://api.line.me/v2/bot/profile/${userId}`;
    const response = UrlFetchApp.fetch(url, {
      'headers': { 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN },
      'method': 'get'
    });
    const profile = JSON.parse(response.getContentText());
    return profile.displayName + "さま";
  } catch (e) {
    logToSheet("ERROR", "名前取得エラー: " + e.message);
    return "ご主人さま"; 
  }
}

// 制御ワードリスト（正規表現・完全一致）
const IGNORE_REGEX = /^(カスタム|開始時刻|(\d|[０-９])+(分|時間|日)(前|後)|Yes|No|yes|no|はい|いいえ|リセット|キャンセル|やめる)$/;

function doPost(e) {

  // 同時実行防止
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return ContentService.createTextOutput("OK");

  try {
    const json = JSON.parse(e.postData.contents);
    const event = json.events[0];
    if (!event) return ContentService.createTextOutput("OK");

    // LINE公式の「再送フラグ」をチェック
    if (event.deliveryContext && event.deliveryContext.isRedelivery) {
      return ContentService.createTextOutput("OK");
    }

    const replyToken = event.replyToken;
    const userId = event.source.userId;
    const userMessage = event.message && event.message.text ? event.message.text.trim() : "";

    logToSheet("DEBUG-1", "LINEからの受信テキスト: " + userMessage);

    const username = getDisplayName(userId) || "ご主人さま";

    // キャッシュによるリトライをブロック
    const cache = CacheService.getScriptCache();
    const cacheKey = (event.message && event.message.id) || event.webhookEventId || replyToken;
    
    if (cache.get(cacheKey)) {
      logToSheet("INFO", "キャッシュでブロックしました: " + userMessage);
      return ContentService.createTextOutput("OK");
    }
    cache.put(cacheKey, "PROCESSED", 1800);

    // ステータスリセット
    if (userMessage === 'リセット' || userMessage === 'キャンセル' || userMessage === 'やめる') {
      clearUserState(userId);
      sendLineReply(replyToken, "状態をリセットしました。\nもう一度最初から教えていただけますか？");
      return ContentService.createTextOutput("OK");
    }

    // 現在のステータス取得
    let userState = props.getProperty(`STATE_${userId}`);

    // 自動リセット機能
    if (userMessage.match(/(予定|入れて|登録|タスク)/)) {
          logToSheet("INFO", "新規コマンド検知：ステータスを強制リセットします");
          clearUserState(userId);
          userState = null; 
    }

    // ステータスがある時
    if (userState) {
      if (userState === 'awaiting_remind_confirm') {
        const msgLower = userMessage.toLowerCase();
        if (msgLower === 'yes' || userMessage === 'はい') {
          props.setProperty(`STATE_${userId}`, 'awaiting_remind_time');
          sendTimeQuickReply(replyToken, "いつ通知いたしましょう？");
        } else if (msgLower === 'no' || userMessage === 'いいえ') {
          sendLineReply(replyToken, `かしこまりました。\nまたご用の際はお声がけくださいませ✨`);
          clearUserState(userId);
        } else {
          sendLineReply(replyToken, "恐れ入りますが、「Yes」か「No」で教えていただけますか？");
        }
        return ContentService.createTextOutput("OK");
      }

      if (userState === 'awaiting_remind_time') {
        handleTimeSelection(replyToken, userId, userMessage, username);
        return ContentService.createTextOutput("OK");
      }

      if (userState === 'awaiting_custom_time') {
        handleCustomTimeInput(replyToken, userId, userMessage, username);
        return ContentService.createTextOutput("OK"); 
      }
      
      clearUserState(userId);
    }

    // ステータスがない時
    if (IGNORE_REGEX.test(userMessage)) {
      logToSheet("INFO", "制御ワード（完全一致）のためスキップしました: " + userMessage);
      return ContentService.createTextOutput("OK");
    }

    // Gemini解析
    const aiResult = callGeminiAPI(userMessage, username);

    if (aiResult) {
      // もしリストで返ってきたら、一番最初の予定を代表にする
      const targetData = Array.isArray(aiResult) ? aiResult[0] : aiResult;
      
      // カレンダー登録
      const result = createEventFromAi(targetData, userId);

      if(result.success) {
        props.setProperty(`STATE_${userId}`, 'awaiting_remind_confirm');
        props.setProperty(`TEMP_EVENT_${userId}`, JSON.stringify({
          summary: targetData.title,
          start: targetData.start,
          end: targetData.end,
          isAllDay: targetData.isAllDay === "TRUE"
        }));
        sendYesNoQuickReply(replyToken, `${result.message}\n\nこちらの通知は必要でしょうか？`);
      } else {
        sendLineReply(replyToken, result.message);
      }
    }
    return ContentService.createTextOutput("OK");
  } catch (err) {
    logToSheet("ERROR", "doPostに失敗: " + err.message);
    return ContentService.createTextOutput("OK");
  } finally {
    lock.releaseLock();
  }
}

function parseJst(dateStr) {
  if (!dateStr) return null;
  return new Date(
    dateStr.includes("+09:00") ? dateStr : dateStr + "+09:00"
  );
}

// カレンダー登録
function createEventFromAi(data, userId) {
  logToSheet("INFO", "Geminiから受け取ったデータ:", JSON.stringify(data)); 
  try {
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    const startTime = parseJst(data.start);
    const endTime   = parseJst(data.end);


    if (isNaN(startTime.getTime())) throw new Error(`開始時間の形式が不正です！`);
    if (isNaN(endTime.getTime())) throw new Error(`終了時間の形式が不正です！`);
    
    const isAllDay = (String(data.isAllDay).toLowerCase() === "true");

    const searchStart = new Date(startTime.getTime() - 60 * 60 * 1000);
    const searchEnd = new Date(startTime.getTime() + 60 * 60 * 1000);
    const existingEvents = calendar.getEvents(searchStart, searchEnd);
    
    for (const event of existingEvents) {
      const timeDiff = Math.abs(event.getStartTime().getTime() - startTime.getTime());
      const isSameTitle = event.getTitle() === data.title;
      if (isSameTitle && timeDiff < 10 * 60 * 1000) {
        logToSheet("INFO", "重複のためスキップ: " + data.title);
        const username = getDisplayName(userId) || "ご主人さま";
        return { success: false, message: `${username}、そのご予定はもう登録してございます✨` };
      }
    }
  
    const options = {
      location: data.location === "なし" ? "" : data.location,
      description: data.description === "なし" ? "" : data.description
    };

    if (isAllDay) {
      calendar.createAllDayEvent(data.title, startTime, options);
    } else {
      calendar.createEvent(data.title, startTime, endTime, options);
    }
    
    return { success: true, message: data.reply_message, eventData: data };

  } catch (e) {
    logToSheet("ERROR", "エラー発生: " + e.message); 
    return { success: false, message: "申し訳ございません、エラーが出て登録できませんでした💦" };
  }
}

function handleTimeSelection(replyToken, userId, message, username) {
  if (message === 'カスタム') {
      props.setProperty(`STATE_${userId}`, 'awaiting_custom_time');
      sendLineReply(replyToken, "通知したいタイミングを教えていただけますか？\n(例:「3分後」、「来週の金曜夜8時」など)");
      return; 
  }

  try {
    const tempEventStr = props.getProperty(`TEMP_EVENT_${userId}`);
    if (!tempEventStr) {
      sendLineReply(replyToken, "申し訳ございません、予定の情報が見つかりませんでした💦");
      clearUserState(userId);
      return;
    }
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('ReminderQueue');
    if (!sheet) { sendLineReply(replyToken, "エラー: 'ReminderQueue' シートが見つかりません💦"); return; }

    const eventData = JSON.parse(tempEventStr);
    const startDt = parseJst(eventData.start); 
    const startTimeStr = eventData.isAllDay ? "終日" : Utilities.formatDate(startDt, "JST", "M/d HH:mm");

    let remindDt = new Date(startDt.getTime());
    let isDouble = false; 
    let d1, d2;

    // 複雑なセットから順に判定
    // includesにすることで「＆」「&」などの表記ゆれに対応
    if (message.includes('1日前') && message.includes('3時間前')) {
      isDouble = true;
      d1 = new Date(startDt.getTime()); d1.setDate(d1.getDate() - 1);
      d2 = new Date(startDt.getTime()); d2.setHours(d2.getHours() - 3);
    } else if (message === '開始時刻') {
    } else if (message === '15分前') {
      remindDt.setMinutes(remindDt.getMinutes() - 15);
    } else if (message === '1時間前') {
      remindDt.setHours(remindDt.getHours() - 1);
    } else if (message === '3時間前') {
      remindDt.setHours(remindDt.getHours() - 3);
    } else if (message === '1日前') {
      remindDt.setDate(remindDt.getDate() - 1);
    } else {
      logToSheet("ERROR", "想定外の選択肢が届きました" + message);
      clearUserState(userId); return;
    }

    if (isDouble) {
      const d1Str = Utilities.formatDate(d1, "JST", "yyyy/MM/dd HH:mm:ss");
      const d2Str = Utilities.formatDate(d2, "JST", "yyyy/MM/dd HH:mm:ss");
      
      // 重複チェック
      if (isDuplicateEntry(sheet, userId, eventData.summary, startTimeStr, d1Str)) {
        logToSheet("INFO", "重複ブロック（2段通知）");
        clearUserState(userId); return;
      }
      
      sheet.appendRow([d1Str, eventData.summary, userId, startTimeStr]);
      sheet.appendRow([d2Str, eventData.summary, userId, startTimeStr]);
      sendLineReply(replyToken, `${username}、ご予定の1日前と3時間前にしっかりお知らせいたします！✨`);
    } else {
      const remindTimeStr = Utilities.formatDate(remindDt, "JST", "yyyy/MM/dd HH:mm:ss");
      
      // 重複チェック
      if (isDuplicateEntry(sheet, userId, eventData.summary, startTimeStr, remindTimeStr)) {
        logToSheet("INFO", "重複ブロック（単発通知）");
        clearUserState(userId); return;
      }

      sheet.appendRow([remindTimeStr, eventData.summary, userId, startTimeStr]);
      const confirmTime = Utilities.formatDate(remindDt, "JST", "M/d HH:mm");
      sendLineReply(replyToken, `${username}、${confirmTime}にお知らせいたします✨`);
    }
    clearUserState(userId);

  } catch (e) { 
    logToSheet("ERROR", e.message);
    sendLineReply(replyToken, "申し訳ございません、エラーが発生しました💦");
  }
}

function handleCustomTimeInput(replyToken, userId, message, username) {
  const scriptProperties = PropertiesService.getScriptProperties(); 
  const tempEventStr = scriptProperties.getProperty(`TEMP_EVENT_${userId}`);
  
  if (!tempEventStr) {
    sendLineReply(replyToken, "申し訳ございません、予定の情報を見失ってしまいました…💦");
    clearUserState(userId);
    return;
  }

  const API_KEY = scriptProperties.getProperty('GEMINI_API_KEY');
  const MODEL_NAME = 'gemini-flash-latest';
  const URL = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${API_KEY}`;
  const nowStr = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd HH:mm:ss');
  const prompt = `Role: Precise Time Extractor\nCurrent Time: ${nowStr}\nTask: output {"remind_at": "YYYY-MM-DDTHH:mm:ss"}\nUser Input: ${message}`;
  const payload = { "contents": [{ "parts": [{"text": prompt}] }], "generationConfig": { "responseMimeType": "application/json" } };

  try {
    const response = UrlFetchApp.fetch(URL, { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload) });
    const resJson = JSON.parse(response.getContentText());
    let resText = resJson.candidates[0].content.parts[0].text.replace(/```json/g, "").replace(/```/g, "").trim();
    const aiData = JSON.parse(resText);
    
    const remindDt = new Date(aiData.remind_at);
    const remindTimeStr = Utilities.formatDate(remindDt, "JST", "yyyy/MM/dd HH:mm:ss");

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('ReminderQueue');
    const eventData = JSON.parse(tempEventStr);

    const startDtForSheet = parseJst(eventData.start);
    const startTimeStr = eventData.isAllDay ? "終日" : Utilities.formatDate(startDtForSheet, "JST", "M/d HH:mm");

    // 重複チェック
    if (isDuplicateEntry(sheet, userId, eventData.summary, startTimeStr, remindTimeStr)) {
      logToSheet("INFO", "重複ブロック（カスタム）");
      clearUserState(userId); return;
    }

    sheet.appendRow([remindTimeStr, eventData.summary, userId, startTimeStr]);

    const confirmTime = Utilities.formatDate(remindDt, "JST", "M/dのHH:mm");
    sendLineReply(replyToken, `${username}、${confirmTime}にお知らせいたします✨`);
    clearUserState(userId);

  } catch (e) {
    logToSheet("ERROR", "カスタムエラー: " + e.message);
    sendLineReply(replyToken, `申し訳ございません、時間の解析に失敗してしまいました💦`);
  }
}

// 重複チェック用関数
function isDuplicateEntry(sheet, userId, summary, startTimeStr, remindTimeStr) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    const startRow = Math.max(2, lastRow - 10);
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();
    
    for (let i = data.length - 1; i >= 0; i--) {
      const rowRemind = Utilities.formatDate(new Date(data[i][0]), "JST", "yyyy/MM/dd HH:mm:ss");
      const rowSummary = data[i][1];
      const rowId = data[i][2];
      const rowStart = data[i][3];

      if (rowId === userId && rowSummary === summary && rowStart === startTimeStr && rowRemind === remindTimeStr) {
        return true; 
      }
    }
    return false;
  } catch(e) {
    return false; 
  }
} 

function sendLineNotify(token, message) {
  const options = { "method": "post", "headers": {"Authorization": "Bearer " + token}, "payload": {"message": message} };
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}

function postToLine(replyToken, messages) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN },
    'method': 'post',
    'payload': JSON.stringify({ 'replyToken': replyToken, 'messages': messages })
  });
}

function sendLineReply(replyToken, message) { postToLine(replyToken, [{ 'type': 'text', 'text': message }]); }
function sendYesNoQuickReply(replyToken, message) {
  const items = [{ 'type': 'action', 'action': { 'type': 'message', 'label': 'Yes', 'text': 'Yes' } }, { 'type': 'action', 'action': { 'type': 'message', 'label': 'No', 'text': 'No' } }];
  postToLine(replyToken, [{ 'type': 'text', 'text': message, 'quickReply': { 'items': items } }]);
}

function sendTimeQuickReply(replyToken, message) {
  const options = ['開始時刻', '15分前', '1時間前', '3時間前', '1日前', '1日前＆3時間前', 'カスタム'];
  const items = options.map(opt => ({ 'type': 'action', 'action': { 'type': 'message', 'label': opt, 'text': opt } }));
  postToLine(replyToken, [{ 'type': 'text', 'text': message, 'quickReply': { 'items': items } }]);
}

function clearUserState(userId) { props.deleteProperty(`STATE_${userId}`); props.deleteProperty(`TEMP_EVENT_${userId}`); }

function checkAndSendReminders() {
  const props = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID');
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('ReminderQueue');
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = data.length - 1; i >= 1; i--) {
    const remindTimeRaw = data[i][0]; 
    const summary = data[i][1];       
    const userId = data[i][2];        
    const startTimeStr = data[i][3]; // D列
    
    const remindTime = new Date(remindTimeRaw);
    if (now < remindTime) continue;

    const username = getDisplayName(userId) || "ご主人さま";
    let timingInfo = "予定";
    let displayTime = "";

    try {
      if (startTimeStr instanceof Date) {
        // スプシが勝手に日付データに変換していた場合！
        displayTime = Utilities.formatDate(startTimeStr, "JST", "M/d HH:mm");
        const diffMin = Math.round((startTimeStr.getTime() - remindTime.getTime()) / (1000 * 60));
        
        if (diffMin <= 1) timingInfo = "予定時刻";
        else if (diffMin < 60) timingInfo = `${diffMin}分前`;
        else if (diffMin < 1440) timingInfo = `${Math.round(diffMin/60)}時間前`;
        else timingInfo = `${Math.round(diffMin/1440)}日前`;
        
      } else if (startTimeStr === "終日") {
        displayTime = "終日";
        timingInfo = "事前";
        
      } else if (typeof startTimeStr === "string") {
        // スプシに文字列のまま渡された場合の計算ロジック
        displayTime = startTimeStr;
        const parts = startTimeStr.split(/[\/\s:]/);
        if (parts.length >= 4) {
          const month = parseInt(parts[0], 10) - 1;
          const day = parseInt(parts[1], 10);
          const hour = parseInt(parts[2], 10);
          const minute = parseInt(parts[3], 10);
          const startDt = new Date(remindTime.getFullYear(), month, day, hour, minute);
          const diffMin = Math.round((startDt.getTime() - remindTime.getTime()) / (1000 * 60));
          
          if (diffMin <= 1) timingInfo = "予定時刻";
          else if (diffMin < 60) timingInfo = `${diffMin}分前`;
          else if (diffMin < 1440) timingInfo = `${Math.round(diffMin/60)}時間前`;
          else timingInfo = `${Math.round(diffMin/1440)}日前`;
        }
      }
    } catch (e) {
      logToSheet("ERROR", "通知時刻の計算に失敗: " + e.message);
      timingInfo = "予定";
      // 万が一エラーでも時間が空欄にならないようにする
      displayTime = (startTimeStr instanceof Date) ? Utilities.formatDate(startTimeStr, "JST", "M/d HH:mm") : String(startTimeStr);
    }

    const aiMsg = generateAiRemindMessage(summary, displayTime, username, timingInfo);
    
    try {
      pushLineMessage(userId, aiMsg); 
      sheet.deleteRow(i + 1);
      logToSheet("INFO", `通知送信完了: ${summary}`);
    } catch (e) {
      logToSheet("ERROR", `送信失敗につき行を保持します: ${e.message}`);
    }
  }
}

function pushLineMessage(userId, message) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN },
    'method': 'post',
    'payload': JSON.stringify({ 'to': userId, 'messages': [{ 'type': 'text', 'text': message }] })
  });
}

function callGeminiAPI(userMessage, username) {
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const MODEL_NAME = 'gemini-flash-latest'; 
  const URL = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${API_KEY}`;
  let now = new Date();
  let nowStr = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd HH:mm:ss');
  let tomorrow = new Date(now);
  tomorrow.setDate(now.getDate() + 1);
  let tomorrowStr = Utilities.formatDate(tomorrow, 'JST', 'yyyy-MM-dd');

  const prompt = `
  Role: Professional Butler AI ✨
  Current Time (JST): ${nowStr}
  User Name: ${username}
  Task: Extract event details from the user's message and output ONLY in JSON format.
  Rules:
  1. title: Event name. Infer from input.
  2. start/end: ISO 8601 (YYYY-MM-DDTHH:mm:ss). If NO time is specified, set start to 00:00:00. Default duration: 1 hour.
  3. reply_message: Polite butler tone for ${username}. (Example: "かしこまりました。明日の10時より会議のご予定を承ります✨")
  4. isAllDay: true (no time specified) or false (time specified).
  5. location/description: Extract if mentioned. If none, "なし".
  Examples:
  User: 明日の19時に焼肉
  AI: { "title": "焼肉", "start": "${tomorrowStr}T19:00:00", "end": "${tomorrowStr}T21:00:00", "reply_message": "${username}、明日の19時から焼肉のご予定ですね。承知いたしました✨", "isAllDay": false, "location": "なし", "description": "なし" }
  User Input: ${userMessage}
  `;
  const payload = { "contents": [{ "parts": [{"text": prompt}] }], "generationConfig": { "temperature": 0.0, "responseMimeType": "application/json" } };
  const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload), 'muteHttpExceptions': true };

  logToSheet("DEBUG-2", "Geminiへ送信するプロンプト: " + prompt);

  try {
    const response = UrlFetchApp.fetch(URL, options);

    const responseText = response.getContentText(); // debug用
    
    logToSheet("DEBUG-3", "Geminiからの生返答: " + responseText);
    
    // APIレスポンスからコンテンツ部分を抽出
    const json = JSON.parse(response.getContentText());
    if (json.error || !json.candidates || !json.candidates[0].content) return null;
    let aiText = json.candidates[0].content.parts[0].text.trim();

    logToSheet("DEBUG-4", "パース直前のテキスト: " + aiText);

    // 抽出したテキストをJSONオブジェクトとしてパース
    const data = JSON.parse(aiText);
    
    // リスト（配列）形式レスポンスに対する正規化処理
    if (Array.isArray(data)) {
      return data.map(item => ({
        "title": item.title,
        "start": item.start,
        "end": item.end,
        "reply_message": item.reply_message,
        "isAllDay": String(item.isAllDay).toUpperCase(),
        "location": item.location || "なし",
        "description": item.description || "なし"
      }));
    }
    
    // 単一のレスポンスに対する正規化処理
    return { 
      "title": data.title, 
      "start": data.start, 
      "end": data.end, 
      "reply_message": data.reply_message, 
      "isAllDay": String(data.isAllDay).toUpperCase(), 
      "location": data.location || "なし", 
      "description": data.description || "なし" 
    };
  } catch (e) {
    logToSheet("ERROR", 'JSONオブジェクトへのパース失敗:', e);
    return null;
  }
}

function generateAiRemindMessage(summary, dateTimeStr, username, timingInfo) {
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const MODEL_NAME = 'gemini-flash-latest'; 
  const URL = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${API_KEY}`;
  
  // 「事前通知」であることを認識させる
  const prompt = `あなたは有能な「執事」です。
  ${username}に、予定「${summary}」の【${timingInfo}】になったことをお知らせしてください。
  予定の開始時間は【${dateTimeStr}】です。
  
  指示：
  ・「${timingInfo}のお時間です」という言葉を使い、事前通知であることを優雅に伝えてください。
  ・✨を使い、2-3行で出力してください。`;

  const payload = { "contents": [{ "parts": [{"text": prompt}] }], "generationConfig": { "temperature": 0.8 } };
  try {
    const response = UrlFetchApp.fetch(URL, { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload) });
    return JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.trim();
  } catch (e) { 
    return `【${username}、${timingInfo}のお時間です✨】\n予定：${summary}\n開始：${dateTimeStr}`; 
  }
}

// スプレッドシートにログを書き出す
function logToSheet(type, message) {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) return; 

  let ss = SpreadsheetApp.openById(sheetId);
  
  let sheet = ss.getSheetByName('log');
  if (!sheet) {
    sheet = ss.insertSheet('log');
  }

  let now = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
  let logMessage = typeof message === "object" ? JSON.stringify(message) : message;

  sheet.appendRow([now, type, logMessage]);
}