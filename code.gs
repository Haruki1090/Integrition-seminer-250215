/************************************************
 * 1) 定数の設定
 ************************************************/
// Notion
const NOTION_API_TOKEN = PropertiesService.getScriptProperties().getProperty('NOTION_API_TOKEN'); // NotionのInternal Integration Token
const NOTION_DB_ID = PropertiesService.getScriptProperties().getProperty('NOTION_DB_ID'); // 作成したDBのID
const NOTION_VERSION = '2022-06-28'; // Notion APIバージョン

// LINE
const LINE_CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');

// LINEユーザーID
const MY_LINE_USER_ID = PropertiesService.getScriptProperties().getProperty('MY_LINE_USER_ID'); // 単一ユーザーの場合のみ使用

/************************************************
 * 2) doPost - Webhookのエントリポイント
 ************************************************/
function doPost(e) {
  try {
    const contentType = e.postData.type;
    const payload = JSON.parse(e.postData.contents);
    logMessage(`Received payload: ${JSON.stringify(payload, null, 2)}`, 'INFO'); // インデントを追加して可読性を向上
    
    // NotionからのWebhookかLINEからのメッセージかを判別
    if (isNotionWebhook(payload)) {
      handleNotionWebhook(payload);
    } else if (isLineMessage(payload)) {
      handleLineMessage(payload);
    } else {
      logMessage('Unknown webhook source', 'WARNING');
    }

    // 応答（Webhookにはstatus:200を返す）
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    logMessage(`doPost Error: ${error}`, 'ERROR');
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * NotionからのWebhookか判別
 */
function isNotionWebhook(payload) {
  // NotionからのWebhook特有のフィールドをチェック
  // 例: Notionのページ追加イベントには特定のフィールドが含まれる
  return payload.object === 'page' && payload.properties && payload.id;
}

/**
 * LINEからのメッセージか判別
 */
function isLineMessage(payload) {
  // LINEのメッセージには特定のフィールドが含まれる
  return payload.events && payload.events.length > 0 && payload.events[0].source && payload.events[0].source.userId;
}

/************************************************
 * 3) NotionからのWebhook処理
 ************************************************/
function handleNotionWebhook(payload) {
  try {
    // ペイロードの内容に応じて必要なデータを抽出
    const name = payload.properties.Name.title[0].plain_text;
    const amount = payload.properties.Amount.number;
    const type = payload.properties.Type.select.name;
    const startDate = payload.properties.StartDate.date.start;
    const nextBillingDate = payload.properties.NextBillingDate.date.start;
    const pageId = payload.id; // NotionのページID

    logMessage(`Handling Notion webhook for PageID: ${pageId}`, 'INFO');

    // カレンダー関連のコードを削除

    // 必要に応じて他の処理を追加
    // 例えば、ステータスの更新や通知など

  } catch (error) {
    logMessage(`handleNotionWebhook Error: ${error}`, 'ERROR');
  }
}

/************************************************
 * 4) LINEからのメッセージ処理
 ************************************************/
function handleLineMessage(payload) {
  try {
    const events = payload.events;
    events.forEach(event => {
      if (event.type === 'message' && event.message.type === 'text') {
        const userId = event.source.userId; // UserIDを取得
        logMessage(`UserID: ${userId}`, 'INFO');

        const userMessage = event.message.text.trim();

        if (userMessage === 'userID') {
          // UserIDをLINEに返信
          sendTextMessage(`あなたのLINE UserIDは次の通りです: ${userId}`);
        } else if (userMessage.startsWith('登録')) {
          handleRegisterCommand(userMessage, userId);
        } else if (userMessage === '一覧') {
          handleListCommand(userId);
        } else if (userMessage.startsWith('解約')) {
          handleCancelCommand(userMessage, userId);
        } else if (userMessage === '完了') {
          handlePaymentConfirmation(userId);
        } else {
          sendTextMessage(
            '利用できるコマンド:\n' +
            '・ 登録 名前:Netflix, 金額:990, 種類:月額, 開始日:2025-01-01\n' +
            '・ 一覧\n' +
            '・ 解約 Netflix\n' +
            '・ 完了'
          );
        }
      }
    });
  } catch (error) {
    logMessage(`handleLineMessage Error: ${error}`, 'ERROR');
    sendTextMessage(`メッセージ処理中にエラーが発生しました: ${error.message}`);
  }
}

/************************************************
 * 5) サブスク登録
 ************************************************/
function handleRegisterCommand(message, userId) {
  try {
    // 例: 登録 名前:Netflix, 金額:990, 種類:月額, 開始日:2025-01-01
    const regex = /登録\s+名前\s*:\s*(.+?),\s*金額\s*:\s*(\d+),\s*種類\s*:\s*(月額|年額),\s*開始日\s*:\s*(\d{4}-\d{2}-\d{2})/;
    const match = message.match(regex);

    if (!match) {
      sendTextMessage(
        '登録コマンドの形式が正しくありません。\n' +
        '例: 登録 名前:Netflix, 金額:990, 種類:月額, 開始日:2025-01-01'
      );
      return;
    }

    const name = match[1].trim();
    const amount = Number(match[2].trim());
    const type = match[3].trim();
    const startDateStr = match[4].trim();

    // 次回請求日を計算
    let nextDate = new Date(startDateStr);
    if (type === '月額') {
      nextDate.setMonth(nextDate.getMonth() + 1);
    } else {
      nextDate.setFullYear(nextDate.getFullYear() + 1);
    }
    const nextBillingDateStr = Utilities.formatDate(nextDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Notionに新規レコード作成し、ページIDを取得
    const pageId = notionCreateSubscription({
      name,
      amount,
      type,
      startDate: startDateStr,
      nextBillingDate: nextBillingDateStr,
      status: 'Active',
      userId: userId // ユーザーIDを追加
    });

    logMessage(`Subscription created with PageID: ${pageId}`, 'INFO');

    // カレンダー登録機能を削除したため、以下のコードを削除
    /*
    addCalendarEvent({
      name,
      amount,
      type,
      nextBillingDate: nextBillingDateStr,
      status: 'Active',
      pageId: pageId
    });
    */

    // 通知
    sendTextMessage(
      `「${name}」を登録しました。\n次回請求日は ${nextBillingDateStr} です。`
    );
  } catch (error) {
    logMessage(`handleRegisterCommand Error: ${error}`, 'ERROR');
    sendTextMessage(`登録中にエラーが発生しました: ${error.message}`);
  }
}

/************************************************
 * 6) 一覧表示
 ************************************************/
function handleListCommand(userId) {
  try {
    // ステータスが「解約」以外のレコードを取得
    const subs = notionQueryByStatuses(['Active', '確認待ち'], userId);
    if (subs.length === 0) {
      sendTextMessage('現在、有効なサブスクリプションはありません。');
      return;
    }

    let msg = '【サブスク一覧】\n';
    subs.forEach((page, idx) => {
      const p = page.properties;
      const name = getTextValue(p.Name);
      const amount = p.Amount.number ?? 0;
      const type = p.Type.select?.name ?? '';
      const startDate = p.StartDate.date?.start ?? '';
      const nextDate = p.NextBillingDate.date?.start ?? '';
      const status = p.Status.select?.name ?? '';
      msg += `\n${idx + 1}. 名称: ${name}\n`;
      msg += `   金額: ${amount}円\n`;
      msg += `   種類: ${type}\n`;
      msg += `   開始日: ${startDate}\n`;
      msg += `   次回支払日: ${nextDate}\n`;
      msg += `   ステータス: ${status}\n`;
    });

    sendTextMessage(msg);
  } catch (error) {
    logMessage(`handleListCommand Error: ${error}`, 'ERROR');
    sendTextMessage(`一覧取得中にエラーが発生しました: ${error.message}`);
  }
}

/************************************************
 * 7) 解約
 ************************************************/
function handleCancelCommand(message, userId) {
  try {
    // 例: 解約 Netflix
    const regex = /解約\s+(.+)/;
    const match = message.match(regex);
    if (!match) {
      sendTextMessage('解約コマンドの形式が正しくありません。\n例: 解約 Netflix');
      return;
    }
    const targetName = match[1].trim();

    // Active or 確認待ち から、該当のサブスクを探す
    const subs = notionQueryByStatuses(['Active', '確認待ち'], userId);
    // 同名が複数あるとややこしいので、最初にヒットした1件だけ処理
    const page = subs.find(p => getTextValue(p.properties.Name) === targetName);

    if (!page) {
      sendTextMessage(`「${targetName}」を解約対象として見つかりませんでした。`);
      return;
    }

    // 解約ステータスに更新
    notionUpdateSubscription(page.id, { Status: '解約' });

    // カレンダー関連のコードを削除

    sendTextMessage(`「${targetName}」を解約しました。`);
  } catch (error) {
    logMessage(`handleCancelCommand Error: ${error}`, 'ERROR');
    sendTextMessage(`解約中にエラーが発生しました: ${error.message}`);
  }
}

/************************************************
 * 8) 支払い完了 - 「完了」
 ************************************************/
function handlePaymentConfirmation(userId) {
  try {
    // ステータスが「確認待ち」のレコードをすべて取得
    const subs = notionQueryByStatuses(['確認待ち'], userId);
    if (subs.length === 0) {
      sendTextMessage('現在「確認待ち」のサブスクリプションはありません。');
      return;
    }

    subs.forEach(page => {
      const p = page.properties;
      const name = getTextValue(p.Name);
      const type = p.Type.select?.name;
      const nextBillingDateStr = p.NextBillingDate.date?.start;
      if (!type || !nextBillingDateStr) return;

      // 次回請求日を再計算
      let dt = new Date(nextBillingDateStr);
      if (type === '月額') {
        dt.setMonth(dt.getMonth() + 1);
      } else {
        dt.setFullYear(dt.getFullYear() + 1);
      }
      const newDateStr = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      // Notion更新 (ステータスをActiveに戻し, NextBillingDateを更新)
      notionUpdateSubscription(page.id, {
        NextBillingDate: newDateStr,
        Status: 'Active'
      });

      // カレンダー関連のコードを削除

      sendTextMessage(
        `「${name}」の支払いを確認しました。\n次回請求日は ${newDateStr} です。`
      );
    });
  } catch (error) {
    logMessage(`handlePaymentConfirmation Error: ${error}`, 'ERROR');
    sendTextMessage(`支払い確認中にエラーが発生しました: ${error.message}`);
  }
}

/************************************************
 * 9) Notion操作のユーティリティ
 ************************************************/
/**
 * 新しいサブスクをNotionにレコード作成し、ページIDを返す
 */
function notionCreateSubscription({ name, amount, type, startDate, nextBillingDate, status, userId }) {
  try {
    const url = 'https://api.notion.com/v1/pages';
    const payload = {
      parent: { database_id: NOTION_DB_ID },
      properties: {
        Name: {
          title: [{ text: { content: name } }]
        },
        Amount: {
          number: amount
        },
        Type: {
          select: { name: type }
        },
        StartDate: {
          date: { start: startDate }
        },
        NextBillingDate: {
          date: { start: nextBillingDate }
        },
        Status: {
          select: { name: status }
        },
        UserID: { // ユーザーIDを保存するプロパティ
          rich_text: [{ text: { content: userId } }]
        }
      }
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${NOTION_API_TOKEN}`,
        'Notion-Version': NOTION_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true // 詳細なエラーメッセージを取得するために追加
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      const responseData = JSON.parse(responseBody);
      if (!responseData.id) {
        throw new Error('Notionにサブスクリプションを作成できませんでした。');
      }

      logMessage(`Notion Subscription Created: ${responseData.id}`, 'INFO');
      return responseData.id; // ページIDを返す
    } else {
      // エラーメッセージを含む詳細なレスポンスをログに記録
      logMessage(`Notion API Error: ${responseBody}`, 'ERROR');
      throw new Error(`Notion API Error: ${responseBody}`);
    }
  } catch (error) {
    logMessage(`notionCreateSubscription Error: ${error}`, 'ERROR');
    throw error; // エラーを再スローして呼び出し元でキャッチ
  }
}

/**
 * 既存レコードを更新 (部分的)
 */
function notionUpdateSubscription(pageId, { NextBillingDate, Status }) {
  try {
    const url = `https://api.notion.com/v1/pages/${pageId}`;
    const properties = {};

    if (NextBillingDate !== undefined) {
      properties['NextBillingDate'] = {
        date: { start: NextBillingDate }
      };
    }
    if (Status !== undefined) {
      properties['Status'] = {
        select: { name: Status }
      };
    }

    const payload = { properties };
    const options = {
      method: 'patch',
      headers: {
        'Authorization': `Bearer ${NOTION_API_TOKEN}`,
        'Notion-Version': NOTION_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true // 詳細なエラーメッセージを取得するために追加
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      logMessage(`Notion Subscription Updated: PageID=${pageId}, Status=${Status}, NextBillingDate=${NextBillingDate}`, 'INFO');
    } else {
      logMessage(`Notion API Error: ${responseBody}`, 'ERROR');
      throw new Error(`Notion API Error: ${responseBody}`);
    }
  } catch (error) {
    logMessage(`notionUpdateSubscription Error: ${error}`, 'ERROR');
  }
}

/**
 * 指定ステータス(配列)を持つレコードをすべて取得
 */
function notionQueryByStatuses(statusArray, userId) {
  try {
    const url = `https://api.notion.com/v1/databases/${NOTION_DB_ID}/query`;
    const filter = {
      and: [
        {
          property: 'UserID', // ユーザーIDをフィルタに追加
          rich_text: { contains: userId }
        },
        {
          or: statusArray.map(st => ({
            property: 'Status',
            select: { equals: st }
          }))
        }
      ]
    };
    const payload = { filter };
    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${NOTION_API_TOKEN}`,
        'Notion-Version': NOTION_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true // 詳細なエラーメッセージを取得するために追加
    };
    
    const res = UrlFetchApp.fetch(url, options);
    const responseCode = res.getResponseCode();
    const responseBody = res.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      const responseData = JSON.parse(responseBody);
      if (!responseData.results) {
        throw new Error('Notionからデータを取得できませんでした。');
      }
      logMessage(`Notion Query by Statuses: Retrieved ${responseData.results.length} records.`, 'INFO');
      return responseData.results;
    } else {
      logMessage(`Notion API Error: ${responseBody}`, 'ERROR');
      return [];
    }
  } catch (error) {
    logMessage(`notionQueryByStatuses Error: ${error}`, 'ERROR');
    return [];
  }
}

/**
 * 日付一致 & ステータス一致のレコード取得
 */
function notionQueryByDateAndStatus(dateStr, status) {
  try {
    const url = `https://api.notion.com/v1/databases/${NOTION_DB_ID}/query`;
    const filter = {
      and: [
        {
          property: 'NextBillingDate',
          date: { equals: dateStr }
        },
        {
          property: 'Status',
          select: { equals: status }
        }
      ]
    };
    const payload = { filter };
    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${NOTION_API_TOKEN}`,
        'Notion-Version': NOTION_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true // 詳細なエラーメッセージを取得するために追加
    };

    const res = UrlFetchApp.fetch(url, options);
    const responseCode = res.getResponseCode();
    const responseBody = res.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      const responseData = JSON.parse(responseBody);
      if (!responseData.results) {
        throw new Error('Notionからデータを取得できませんでした。');
      }
      logMessage(`Notion Query by Date and Status: Retrieved ${responseData.results.length} records.`, 'INFO');
      return responseData.results;
    } else {
      logMessage(`Notion API Error: ${responseBody}`, 'ERROR');
      return [];
    }
  } catch (error) {
    logMessage(`notionQueryByDateAndStatus Error: ${error}`, 'ERROR');
    return [];
  }
}

/**
 * プロパティが Title / Rich Text いずれの場合でもテキストを取り出す汎用
 */
function getTextValue(prop) {
  if (prop.title && prop.title.length > 0) {
    return prop.title[0].plain_text;
  } else if (prop.rich_text && prop.rich_text.length > 0) {
    return prop.rich_text[0].plain_text;
  }
  return '';
}

/************************************************
 * 11) LINEへテキストをPush送信
 ************************************************/
function sendTextMessage(text) {
  try {
    const url = 'https://api.line.me/v2/bot/message/push';

    // 単一ユーザーの場合
    const userID = MY_LINE_USER_ID;
    if (!userID) {
      logMessage('Error: MY_LINE_USER_ID is not set in script properties.', 'ERROR');
      return;
    }

    const payload = {
      to: userID,
      messages: [{
        type: 'text',
        text
      }]
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    };

    const response = UrlFetchApp.fetch(url, options);
    logMessage(`Response: ${response.getContentText()}`, 'INFO');
  } catch (error) {
    logMessage(`sendTextMessage Error: ${error}`, 'ERROR');
  }
}

/************************************************
 * 12) カスタムログ関数
 ************************************************/
/**
 * カスタムログ関数
 * @param {string} message - ログに記録するメッセージ
 * @param {string} [level='INFO'] - ログのレベル（INFO, WARNING, ERRORなど）
 */
function logMessage(message, level = 'INFO') {
  const logEmail = PropertiesService.getScriptProperties().getProperty('LOG_EMAIL');
  
  if (logEmail) {
    // バッファログを使用する場合は、以下のコメントアウトを外してください
    // bufferLog(message, level);
    
    // 直ちにメール送信する場合は以下を使用
    const subject = `[GAS Log] [${level}] ${new Date().toISOString()}`;
    const body = message;
    
    try {
      MailApp.sendEmail(logEmail, subject, body);
    } catch (e) {
      // メール送信に失敗した場合は、Loggerにエラーを記録
      Logger.log(`Failed to send log email: ${e}`);
    }
  } else {
    Logger.log('LOG_EMAIL is not set in script properties.');
  }
  
  // GASのLoggerにも記録
  switch(level.toUpperCase()) {
    case 'INFO':
      Logger.log(message);
      break;
    case 'WARNING':
      Logger.log(`WARNING: ${message}`);
      break;
    case 'ERROR':
      Logger.log(`ERROR: ${message}`);
      break;
    default:
      Logger.log(message);
  }
}
