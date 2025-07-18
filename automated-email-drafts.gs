const MY_EMAIL = Session.getActiveUser().getEmail();
const ADDITIONAL_CC = 'CCに含めたいアドレスがあればこに記載';

// --- あなたの署名を設定してください ---
const MY_SIGNATURE = `
_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
○○　○○ ＜xxxxx_xxxxx@xxxxxx.co.jp＞ 
Mobile NNN-NNNN-NNNN
 
◇【株式会社○○　○○】◇
◆ http://www.xxxxxxxx.co.jp/◆
  〒000-0000
  ○○○○○○○○○○0-0-　
 Tel 00-0000-0000   Fax 00-0000-0000
_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
`;
// ------------------------------------

// --- 除外リストを保持するスプレッドシートの設定 ---
const EXCLUDE_LIST_SPREADSHEET_ID = 'ここに除外リストのスプレッドシートＩＤ'; // 例: '1ABc2DEfGHIjKLmNOpQR_stuvWXYZ012345'
const EXCLUDE_ADDRESS_SHEET_NAME = '除外アドレス'; // メールアドレス除外リストのシート名
const EXCLUDE_DOMAIN_SHEET_NAME = '除外ドメイン'; // 新しいドメイン除外リストのシート名

let EXCLUDE_SENDER_EMAILS = []; // スプレッドシートから読み込んだ除外アドレスリストを保持する変数
let EXCLUDE_SENDER_DOMAINS = []; // スプレッドシートから読み込んだ除外ドメインリストを保持する変数

// --- 自社ドメインの設定 ---
const MY_COMPANY_DOMAIN = 'xxxxxx.co.jp'; // ここにあなたの会社のドメインを設定してください


// スクリプト実行時に一度だけ除外リストを読み込む
function onOpen() {
  EXCLUDE_SENDER_EMAILS = loadExcludeListFromSpreadsheet(EXCLUDE_ADDRESS_SHEET_NAME);
  EXCLUDE_SENDER_DOMAINS = loadExcludeListFromSpreadsheet(EXCLUDE_DOMAIN_SHEET_NAME); // ドメインリストも読み込む
  console.log('除外アドレスリストが読み込まれました:', EXCLUDE_SENDER_EMAILS);
  console.log('除外ドメインリストが読み込まれました:', EXCLUDE_SENDER_DOMAINS);
}

// タイムドリブントリガーで実行されるメイン関数
function checkAndDraftReply() {
  // トリガー実行時に毎回最新の除外リストを読み込む
  EXCLUDE_SENDER_EMAILS = loadExcludeListFromSpreadsheet(EXCLUDE_ADDRESS_SHEET_NAME);
  EXCLUDE_SENDER_DOMAINS = loadExcludeListFromSpreadsheet(EXCLUDE_DOMAIN_SHEET_NAME); // ドメインリストも読み込む

  const threads = GmailApp.search('is:unread to:me', 0, 10);

  threads.forEach(thread => {
    const message = thread.getMessages()[thread.getMessages().length - 1];
    const to = message.getTo();
    const sender = message.getFrom();
    const senderEmail = extractEmailAddress(sender);
    const senderDomain = extractDomainFromEmail(senderEmail); // 送信元ドメインを抽出

    // ★除外リストに含まれる送信元アドレス または 送信元ドメインからのメールはスキップ
    if (EXCLUDE_SENDER_EMAILS.includes(senderEmail) || EXCLUDE_SENDER_DOMAINS.includes(senderDomain)) {
      console.log(`除外対象の送信元 (${senderEmail} または ${senderDomain}) からのメールのためスキップ: (件名: ${message.getSubject()})`);
      message.markRead(); // 処理しないメールも既読にする
      return; 
    }

    if (to.includes(MY_EMAIL)) {
      const subject = message.getSubject();
      const plainBody = message.getPlainBody(); 
      const htmlBody = message.getBody();     
      const replyTo = extractEmailAddress(sender);
      const originalCc = message.getCc();

      // 社内メールかどうかのフラグ
      const isInternalMail = (senderDomain === MY_COMPANY_DOMAIN);

      // プロンプトを生成する際に、社内/社外のフラグを渡す
      const prompt = buildPrompt(plainBody, isInternalMail);

      const aiReplyBody = generateReplyWithOpenAI(prompt);
      const quotedOriginalMail = formatAsGmailQuote(htmlBody, sender, message.getDate(), subject, MY_EMAIL);
      const finalHtmlBody = `<div>${aiReplyBody.replace(/\n/g, '<br>')}</div><br>${MY_SIGNATURE.replace(/\n\r?/g, '<br>')}<br>${quotedOriginalMail}`; // Windowsの改行コードも考慮

      const replySubject = (subject.startsWith("Re:")) ? subject : "Re: " + subject;
      const mergedCc = mergeCc(originalCc, ADDITIONAL_CC);

      GmailApp.createDraft(
        replyTo,
        replySubject,
        '', 
        {
          htmlBody: finalHtmlBody,
          cc: mergedCc,
          inReplyTo: message.getId()
        }
      );
      message.markRead();
    }
  });
}

/**
 * 指定されたスプレッドシートのシートから除外リストを読み込みます。
 * @param {string} sheetName 読み込むシートの名前
 * @returns {Array<string>} 除外する項目（メールアドレスまたはドメイン）の配列
 */
function loadExcludeListFromSpreadsheet(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(EXCLUDE_LIST_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.warn(`警告: シート名 "${sheetName}" が見つかりません。この除外リストは空として処理されます。`);
      return []; // シートが見つからない場合は空のリストを返す
    }

    const range = sheet.getDataRange();
    const values = range.getValues();
    
    // 1列目の値をフィルタリングし、空でないものだけを配列にする
    return values.map(row => String(row[0]).trim().toLowerCase()).filter(item => item !== ''); // ドメイン比較のため小文字に変換
  } catch (e) {
    console.error(`スプレッドシートからの除外リスト (${sheetName}) 読み込みエラー:`, e.message);
    return []; // エラーが発生した場合は空のリストを返す
  }
}


function buildPrompt(receivedBody, isInternalMail) {
  const commonStyle = `
【私のメール文面の特徴】
---
`;

  const externalStyle = `
◾社外向けメール：
・冒頭で相手の会社名、名前を記載（例：株式会社○○ XX様）
・相手の社名、氏名を書いた後に「お世話になっております。」「{会社名}の{名前}です。」と名乗る
・簡潔で敬意を払ったやりとりを重視するが、遜りすぎない
・丁寧だがフレンドリーな文体で硬すぎない
・相手の書いていることすべてに細々と返答しない、要点のみ簡潔に返答
・要点整理が明確で無駄がない
・文末は「よろしくお願い致します。」で締める
`;

  const internalStyle = `
◾社内向けメール：
・冒頭で相手の名前を記載（例：○○へ、○○さん など）
・冒頭で「お疲れ様です。{自分の名前}です。」
・明確な指示・箇条書き・リンク活用で情報整理
・言い切り・具体的な行動指示・期限の明示
・文末は「以上、よろしくお願いします。」

# 情報
{会社名}="○○○○"
{名前}="○○"
`; // ここに情報セクションを残す

  // スタイルサンプル全体を構築
  const styleSample = commonStyle + 
                      (isInternalMail ? internalStyle : externalStyle) + 
                      (isInternalMail ? "" : `\n# 情報\n{会社名}="○○○○"\n{名前}="○○"\n`); // 社内用には重複しないように情報セクションを分離

  return `
あなたは日本語でビジネスメールの返信を行うアシスタントです。
以下は私が受け取ったメールの内容です。
私の文体の特徴は以下の通りです：

${styleSample}

この内容に対して、私が返信する形で、ビジネスメールとして自然な文章を作ってください。
※宛名と署名は含めなくて大丈夫です。
※返信本文のみを生成してください。元のメールの引用はスクリプトで自動的に追加されます。

--- 受信メールの本文 ---
${receivedBody}
--- ここまで ---
`;
}

function generateReplyWithOpenAI(prompt) {
  const apiKey = getOpenAiApiKey();
  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: "gpt-4",
    messages: [
      { role: "system", content: "あなたはビジネスメールの返信を作成する日本語アシスタントです。" },
      { role: "user", content: prompt }
    ],
    temperature: 0.7
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
      console.error("OpenAI API Error (Status: " + response.getResponseCode() + "): " + (result.error ? result.error.message : response.getContentText()));
      return "返信の下書き作成中にエラーが発生しました。詳細はスクリプトの実行ログを確認してください。";
    }

    if (result.choices && result.choices.length > 0 && result.choices[0].message && result.choices[0].message.content) {
      return result.choices[0].message.content.trim();
    } else {
      console.error("OpenAI API Response Error: Unexpected response format.");
      return "返信の下書き作成中に予期せぬAPIレスポンスがありました。";
    }

  } catch (e) {
    console.error("OpenAI API呼び出し中に例外が発生しました: " + e.toString());
    return "返信の下書き作成中に予期せぬエラーが発生しました。詳細はスクリプトの実行ログを確認してください。";
  }
}

function getOpenAiApiKey() {
  return PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
}

function extractEmailAddress(text) {
  const match = text.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
  return match ? match[1] : text;
}

/**
 * メールアドレスからドメイン部分を抽出します。
 * 例: "user@example.com" -> "example.com"
 * @param {string} emailAddress メールアドレス
 * @returns {string} ドメイン部分
 */
function extractDomainFromEmail(emailAddress) {
  const parts = emailAddress.split('@');
  if (parts.length > 1) {
    return parts[1].toLowerCase(); // 小文字にして比較
  }
  return ''; // ドメインが見つからない場合は空文字列を返す
}

function mergeCc(originalCc, additionalCc) {
  const allCc = [];

  if (originalCc && typeof originalCc === 'string' && originalCc.trim() !== '') {
    originalCc.split(',').forEach(addr => {
      const trimmed = addr.trim();
      if (trimmed) {
        allCc.push(trimmed);
      }
    });
  }

  if (additionalCc && typeof additionalCc === 'string' && additionalCc.trim() !== '') {
    allCc.push(additionalCc.trim());
  }
  
  return Array.from(new Set(allCc)).join(', ');
}

/**
 * 元のメールのHTML本文をGmailの引用形式に整形します。
 * @param {string} originalHtmlBody 元のメールのHTML本文
 * @param {string} sender 元のメールの送信者情報（"Name <email@example.com>"形式）
 * @param {Date} sentDate 元のメールの送信日時
 * @param {string} subject 元のメールの件名
 * @param {string} myEmail 自分のメールアドレス
 * @returns {string} 引用形式に整形されたHTML文字列
 */
function formatAsGmailQuote(originalHtmlBody, sender, sentDate, subject, myEmail) {
  const formattedDate = Utilities.formatDate(sentDate, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");

  const quoteInfoHtml = `
<div class="gmail_extra">
  <div class="quote-header">---------- Forwarded message ---------</div>
  <div class="quote-header">From: ${sender}</div>
  <div class="quote-header">Date: ${formattedDate}</div>
  <div class="quote-header">Subject: ${subject}</div>
  <div class="quote-header">To: &lt;${myEmail}&gt;</div>
  <br>
</div>
`;

  const quotedHtml = `<blockquote class="gmail_quote" style="margin:0px 0px 0px 0.8ex;padding:0px 0px 0px 1ex;border-left:1px solid rgb(204,204,204);">
  ${quoteInfoHtml}
  ${originalHtmlBody}
</blockquote>`;

  return quotedHtml;
}