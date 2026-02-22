/**
 * 名刺書き出し君 v2.0
 * iPhoneショートカット不要 / PWA + GAS 構成
 * ================================================================
 * ⚠️  初回: CONFIG を設定 → initialSetup() を手動実行してください
 * ================================================================
 */

// ================================================================
// ⚙️  設定（ここだけ変更すれば動く）
// ================================================================
const CONFIG = {
  SPREADSHEET_ID: '1CMN7gFPRH-PGTQWH6lSU5K52UJvF8WhHWWLH6hK8oI8',  // SheetsのURL中の長い文字列
  SHEET_NAME:     '名刺データ',
  LOG_SHEET_NAME: '処理ログ',
  // ⚠️ 必ず変更してください: PWAの接続設定画面で入力する「合言葉」と同じ値にする
  // 例: 'meishi2024secret' のような文字列（8文字以上推奨）
  API_SECRET:     'welzowelzowelzo',

  // OCRエンジン: 'DRIVE_OCR'（無料）| 'CLAUDE_API'（高精度・要APIキー）
  OCR_ENGINE:     'DRIVE_OCR',
  CLAUDE_API_KEY: '',
  CLAUDE_MODEL:   'claude-haiku-4-5-20251001',

  TIMEZONE:   'Asia/Tokyo',
  DEDUP_ROWS: 200,
};

// ================================================================
// 📋  列定義（順序・名前の変更はここだけ）
// ================================================================
const COLUMNS = [
  { key: 'requestId',     label: '要求ID（重複防止）'    },
  { key: 'companyName',   label: '会社名'                },
  { key: 'department',    label: '部署名'                },
  { key: 'position',      label: '役職'                  },
  { key: 'fullName',      label: '氏名'                  },
  { key: 'furigana',      label: 'フリガナ'              },
  { key: 'officePhone',   label: '電話番号（オフィス）'  },
  { key: 'mobilePhone',   label: '携帯電話番号'          },
  { key: 'fax',           label: 'FAX'                   },
  { key: 'email',         label: 'メールアドレス'        },
  { key: 'postalCode',    label: '郵便番号'              },
  { key: 'address',       label: '住所'                  },
  { key: 'website',       label: 'WebサイトURL'          },
  { key: 'memo',          label: '備考（手入力）'        },
  { key: 'employeeId',    label: '撮影者ID'              },
  { key: 'timestamp',     label: '撮影日時（ISO）'       },
  { key: 'processStatus', label: '処理ステータス'        },
  { key: 'errorReason',   label: '失敗理由'              },
  { key: 'createdAt',     label: 'GAS処理日時'           },
];

// ================================================================
// 🌐  GET: 管理画面 + 設定API
// ================================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';

  if (action === 'getConfig') {
    const ck = cfg('CLAUDE_API_KEY');
    return ok({
      spreadsheetId: cfg('SPREADSHEET_ID'),
      sheetName:     cfg('SHEET_NAME'),
      logSheetName:  cfg('LOG_SHEET_NAME'),
      ocrEngine:     cfg('OCR_ENGINE'),
      hasApiSecret:  !!cfg('API_SECRET'),
      hasClaudeKey:  ck.length > 0,
      claudeKeyHint: ck.length > 8 ? ck.slice(0,14) + '…' + ck.slice(-4) : '',
    });
  }

  if (action === 'testConnection') {
    if ((e.parameter.secret || '') !== cfg('API_SECRET'))
      return ok({ ok: false, message: '認証エラー（E040）' });
    try {
      const ss    = getSpreadsheet();
      const sheet = ss.getSheetByName(cfg('SHEET_NAME'));
      return ok({
        ok:         true,
        ssName:     ss.getName(),
        sheetFound: !!sheet,
        lastRow:    sheet ? Math.max(0, sheet.getLastRow() - 1) : 0,
        message:    sheet
          ? '接続成功！シート「' + sheet.getName() + '」を確認しました'
          : '⚠️ スプレッドシートは開けましたが、シート「' + cfg('SHEET_NAME') + '」が見つかりません',
      });
    } catch (err) {
      return ok({ ok: false, message: err.message });
    }
  }

  // デフォルト: 管理者設定画面
  return HtmlService.createHtmlOutput(adminHtml())
    .setTitle('名刺書き出し君 — 管理設定')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ================================================================
// 🌐  POST: APIエントリポイント
// ================================================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents)
      return ok({ ok: false, code: 'E001', message: 'リクエストデータが空です（E001）' });

    let payload;
    try { payload = JSON.parse(e.postData.contents); }
    catch (_) { return ok({ ok: false, code: 'E002', message: 'JSON解析エラー（E002）' }); }

    // 管理者設定保存（secretは内部で検証）
    if (payload.adminAction === 'saveConfig') return handleAdminSave(payload);

    // 通常API: 認証チェック
    if (!payload.secret || payload.secret !== cfg('API_SECRET'))
      return ok({ ok: false, code: 'E040', message: '認証に失敗しました（E040）' });

    switch (payload.action) {
      case 'extract': return handleExtract(payload);
      case 'save':    return handleSave(payload);
      case 'test':    return ok({ ok: true,  code: 'OK',   message: '接続テスト成功' });
      default:        return ok({ ok: false, code: 'E003', message: '不明なアクション（E003）' });
    }
  } catch (err) {
    try { log('ERROR', '', '', '予期せぬエラー', err.message); } catch (_) {}
    return ok({ ok: false, code: 'E099', message: '予期せぬエラーが発生しました（E099）' });
  }
}

// ================================================================
// 🔍  OCR実行のみ（Sheets保存なし）
// ================================================================
function handleExtract(payload) {
  const { imageBase64, employeeId } = payload;

  if (!imageBase64 || imageBase64.length < 100)
    return ok({ ok: false, code: 'E020', message: '画像データが不正です（E020）' });
  if (!employeeId || !/^\d{4}$/.test(employeeId))
    return ok({ ok: false, code: 'E011', message: '社員番号は4桁の数字で入力してください（E011）' });

  // OCR
  let rawText;
  try {
    rawText = performOCR(imageBase64);
  } catch (err) {
    const detail = String(err.message || '不明なエラー');
    // Drive Advanced Serviceが未有効の場合に分かりやすいメッセージを表示
    const isDriveNotEnabled =
      detail.includes('is not defined') ||
      detail.includes('Drive') ||
      detail.includes('ReferenceError');
    const userMsg = isDriveNotEnabled
      ? '【Drive OCR未設定】GASエディタ →「サービス」→「Drive API」を追加してください。詳細: ' + detail
      : 'OCR読み取りエラー（E030）: ' + detail;
    log('ERROR', '', employeeId, 'OCRエラー', detail);
    return ok({ ok: false, code: 'E030', message: userMsg, extracted: {} });
  }

  // フィールド抽出
  let extracted = {};
  try {
    extracted = (cfg('OCR_ENGINE') === 'CLAUDE_API' && cfg('CLAUDE_API_KEY'))
      ? extractWithClaude(rawText)
      : extractWithRegex(rawText);
  } catch (_) { extracted = {}; }

  return ok({ ok: true, extracted });
}

// ================================================================
// 💾  Sheets保存
// ================================================================
function handleSave(payload) {
  const { data = {}, employeeId, timestamp, requestId, memo = '' } = payload;

  if (!requestId)
    return ok({ ok: false, code: 'E003', message: '要求IDがありません（E003）' });
  if (!employeeId || !/^\d{4}$/.test(employeeId))
    return ok({ ok: false, code: 'E011', message: '社員番号が正しくありません（E011）' });

  // 重複チェック
  if (isDuplicate(requestId))
    return ok({ ok: true, status: 'DUPLICATE', message: 'すでに登録済みです（重複を防止しました）' });

  const processStatus = determineStatus(data);

  try {
    writeRow(buildRow(data, { employeeId, timestamp, requestId, memo, processStatus, errorReason: '' }));
    log('INFO', requestId, employeeId, '登録成功', processStatus);
  } catch (err) {
    log('ERROR', requestId, employeeId, '書き込みエラー', err.message);
    return ok({ ok: false, code: 'E051', message: 'スプレッドシートへの書き込みに失敗しました（E051）' });
  }

  return ok({
    ok: true,
    status: processStatus,
    message: processStatus === 'SUCCESS'
      ? '登録完了！'
      : '登録しました（スプレッドシートで内容をご確認ください）',
  });
}

// ================================================================
// 🔍  OCR: Google Drive Advanced Service
// ================================================================
function performOCR(imageBase64) {
  let bytes;
  try { bytes = Utilities.base64Decode(imageBase64); }
  catch (_) { throw new Error('画像デコード失敗'); }

  const blob = Utilities.newBlob(bytes, 'image/jpeg', 'meishi_ocr_tmp.jpg');
  let file;
  try {
    file = Drive.Files.insert(
      { title: 'meishi_tmp_' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
      blob,
      { convert: true, ocr: true, ocrLanguage: 'ja' }
    );
  } catch (err) {
    // 'Drive is not defined' = Drive Advanced Serviceが有効化されていない
    throw new Error('OCRアップロード失敗: ' + err.message);
  }

  let text = '';
  try {
    text = DocumentApp.openById(file.id).getBody().getText();
  } finally {
    try { Drive.Files.remove(file.id); } catch (_) {}  // 一時ファイル必ず削除
  }

  if (!text || text.trim().length < 3) throw new Error('読み取りテキストが短すぎます');
  return text;
}

// ================================================================
// 🤖  Claude API 構造化抽出（高精度オプション）
// ================================================================
function extractWithClaude(ocrText) {
  const keys = {
    companyName: '会社名', department: '部署名', position: '役職',
    fullName: '氏名（漢字）', furigana: 'フリガナ（不明なら空文字）',
    officePhone: '電話（オフィス、070/080/090以外）',
    mobilePhone: '携帯（070/080/090始まり）', fax: 'FAX番号',
    email: 'メールアドレス', postalCode: '郵便番号（例:100-0001）',
    address: '住所（〒と郵便番号を除く）', website: 'WebサイトURL',
  };
  const prompt = `名刺OCRテキストから情報を抽出してJSONのみ返してください。不明項目は""。\n\n${ocrText}\n\n返すJSONキー:\n${JSON.stringify(keys, null, 2)}`;

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST', muteHttpExceptions: true,
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': cfg('CLAUDE_API_KEY'),
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify({
      model: cfg('CLAUDE_MODEL') || CONFIG.CLAUDE_MODEL,
      max_tokens: 1024,
      messages: [{ role: 'user', content: prompt }],
    }),
  });

  if (res.getResponseCode() !== 200)
    throw new Error('Claude API HTTP ' + res.getResponseCode());
  const match = JSON.parse(res.getContentText()).content[0].text.match(/\{[\s\S]*\}/);
  if (!match) throw new Error('JSONが返されませんでした');
  return JSON.parse(match[0]);
}

// ================================================================
// 🔧  正規表現抽出（Drive OCR後の無料処理）
// ================================================================
function extractWithRegex(text) {
  const f = {
    companyName:'', department:'', position:'', fullName:'', furigana:'',
    officePhone:'', mobilePhone:'', fax:'', email:'', postalCode:'', address:'', website:'',
  };
  const lines = text.split('\n').map(s => s.trim()).filter(Boolean);

  // メール
  const em = text.match(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/);
  if (em) f.email = em[0];

  // URL
  const ur = text.match(/https?:\/\/[^\s　\n]+|www\.[^\s　\n]+/);
  if (ur) f.website = ur[0].replace(/[。、）)]+$/, '');

  // 郵便番号
  const po = text.match(/〒?\s*(\d{3}[－\-]\d{4}|\d{7})/);
  if (po) f.postalCode = po[1].replace(/[－]/g, '-');

  // 携帯
  const mb = text.match(/0[789]0[－\-\s]?\d{4}[－\-\s]?\d{4}/);
  if (mb) f.mobilePhone = mb[0].replace(/[－\s]/g, '-');

  // FAX
  const fx = text.match(/(?:FAX|fax|ファックス)[：: ]*([0-9０-９\-－]+)/i);
  if (fx) f.fax = z2h(fx[1]);

  // 電話（携帯以外）
  const telRe = /([0-9０-９]{2,5}[－\-][0-9０-９]{2,4}[－\-][0-9０-９]{4})/g;
  let tm;
  while ((tm = telRe.exec(text)) !== null) {
    const n = z2h(tm[1]);
    if (/^0[789]0/.test(n)) { if (!f.mobilePhone) f.mobilePhone = n; }
    else { if (!f.officePhone) f.officePhone = n; }
  }

  // 住所
  const ad = text.match(/(北海道|青森|岩手|宮城|秋田|山形|福島|茨城|栃木|群馬|埼玉|千葉|東京|神奈川|新潟|富山|石川|福井|山梨|長野|岐阜|静岡|愛知|三重|滋賀|京都|大阪|兵庫|奈良|和歌山|鳥取|島根|岡山|広島|山口|徳島|香川|愛媛|高知|福岡|佐賀|長崎|熊本|大分|宮崎|鹿児島|沖縄)[^\n]{4,60}/);
  if (ad) f.address = ad[0];

  // 会社名
  for (const l of lines) {
    if (!f.companyName && /株式会社|有限会社|合同会社|社団法人|LLC|Inc\.|Corp\./i.test(l))
      { f.companyName = l.trim(); break; }
  }

  // 役職・部署・フリガナ・氏名
  for (const l of lines) {
    if (!f.furigana  && /^[\u30A0-\u30FF\u30FC ]+$/.test(l) && l.length >= 3)
      { f.furigana  = l.trim(); continue; }
    if (!f.position  && /部長|課長|係長|取締役|社長|マネージャー|ディレクター|Manager|Director/.test(l))
      { f.position  = l.trim(); continue; }
    if (!f.department && /部|課|室|チーム|Division|Dept/.test(l) && !/株式会社|有限会社/.test(l) && l.length < 30)
      { f.department = l.trim(); continue; }
  }
  if (!f.fullName) {
    for (const l of lines) {
      if (/^[\u4e00-\u9fff]{1,4}[\u3000 ][\u4e00-\u9fff]{1,4}$/.test(l.trim()) &&
          !/株式会社|都|道|府|県|市|区/.test(l))
        { f.fullName = l.trim(); break; }
    }
  }
  return f;
}

// ================================================================
// 📊  ステータス判定
// ================================================================
function determineStatus(ex) {
  const filled = Object.values(ex).filter(v => v && String(v).trim()).length;
  const hasKey  = (ex.fullName && ex.fullName.trim()) || (ex.companyName && ex.companyName.trim());
  if (!hasKey || filled < 2) return 'REVIEW_NEEDED';
  return 'SUCCESS';
}

// ================================================================
// 📝  行データ構築
// ================================================================
function buildRow(ex, meta) {
  const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX");
  const m = {
    requestId: meta.requestId || '', companyName: ex.companyName || '',
    department: ex.department || '', position: ex.position || '',
    fullName: ex.fullName || '', furigana: ex.furigana || '',
    officePhone: ex.officePhone || '', mobilePhone: ex.mobilePhone || '',
    fax: ex.fax || '', email: ex.email || '', postalCode: ex.postalCode || '',
    address: ex.address || '', website: ex.website || '', memo: meta.memo || '',
    employeeId: meta.employeeId || '', timestamp: meta.timestamp || '',
    processStatus: meta.processStatus || '', errorReason: meta.errorReason || '',
    createdAt: now,
  };
  return COLUMNS.map(c => m[c.key] !== undefined ? m[c.key] : '');
}

// ================================================================
// ✍️  Sheets書き込み
// ================================================================
function writeRow(row) {
  const sheet = getDataSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS.map(c => c.label));
    sheet.getRange(1, 1, 1, COLUMNS.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow(row);
}

// ================================================================
// 📜  ログ書き込み
// ================================================================
function log(level, reqId, empId, event, detail) {
  try {
    const ss = getSpreadsheet();
    const ls = ensureSheet(ss, cfg('LOG_SHEET_NAME'));
    if (ls.getLastRow() === 0) {
      ls.appendRow(['日時','レベル','要求ID','社員ID','イベント','詳細']);
      ls.getRange(1,1,1,6).setFontWeight('bold').setBackground('#333').setFontColor('#fff');
      ls.setFrozenRows(1);
    }
    const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    ls.appendRow([now, level, reqId, empId, event, String(detail).slice(0, 500)]);
  } catch (_) {}
}

// ================================================================
// 🔁  重複チェック
// ================================================================
function isDuplicate(requestId) {
  try {
    const sheet = getDataSheet();
    const last  = sheet.getLastRow();
    if (last < 2) return false;
    const count = Math.min(CONFIG.DEDUP_ROWS, last - 1);
    const start = Math.max(2, last - count + 1);
    return sheet.getRange(start, 1, count, 1).getValues().some(r => r[0] === requestId);
  } catch (_) { return false; }
}

// ================================================================
// ⚙️  管理者設定保存
// ================================================================
function handleAdminSave(payload) {
  if (!payload.secret || payload.secret !== cfg('API_SECRET'))
    return ok({ ok: false, message: '認証エラー（E040）' });

  const props   = PropertiesService.getScriptProperties();
  const updates = {};
  const set = (k, v) => { if (typeof v === 'string' && v.trim()) updates[k] = v.trim(); };

  set('SPREADSHEET_ID', payload.spreadsheetId);
  set('SHEET_NAME',     payload.sheetName);
  set('LOG_SHEET_NAME', payload.logSheetName);
  set('OCR_ENGINE',     payload.ocrEngine);
  if (payload.newApiSecret && payload.newApiSecret.length >= 6)
    updates['API_SECRET'] = payload.newApiSecret.trim();
  if (payload.claudeApiKey && payload.claudeApiKey.startsWith('sk-'))
    updates['CLAUDE_API_KEY'] = payload.claudeApiKey.trim();

  if (!Object.keys(updates).length) return ok({ ok: false, message: '変更する項目がありません' });
  props.setProperties(updates);
  log('INFO', 'admin', 'admin', '設定変更', Object.keys(updates).join(', '));
  return ok({ ok: true, message: '設定を保存しました（' + Object.keys(updates).join('、') + '）' });
}

// ================================================================
// 🛠️  ユーティリティ
// ================================================================
function cfg(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || CONFIG[key] || '';
}
function getSpreadsheet() {
  const id = cfg('SPREADSHEET_ID');
  if (!id || id === 'YOUR_SPREADSHEET_ID_HERE')
    throw new Error('E050: スプレッドシートIDが未設定です');
  try { return SpreadsheetApp.openById(id); }
  catch (_) { throw new Error('E051: スプレッドシートを開けません（ID: ' + id + '）'); }
}
function getDataSheet() { return ensureSheet(getSpreadsheet(), cfg('SHEET_NAME') || '名刺データ'); }
function ensureSheet(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function ok(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function z2h(s) {
  return String(s).replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 65248))
                  .replace(/[－]/g, '-');
}

// ================================================================
// 🔧  初回セットアップ（GASエディタから手動実行）
// ================================================================
function initialSetup() {
  try {
    const ss = getSpreadsheet();
    ensureSheet(ss, cfg('SHEET_NAME')     || '名刺データ');
    ensureSheet(ss, cfg('LOG_SHEET_NAME') || '処理ログ');
    Logger.log('✅ セットアップ完了: ' + ss.getName());
    try { SpreadsheetApp.getUi().alert('✅ セットアップ完了！\nスプレッドシート: ' + ss.getName()); }
    catch (_) {}
  } catch (err) {
    Logger.log('❌ エラー: ' + err.message);
    try { SpreadsheetApp.getUi().alert('❌ エラー: ' + err.message); } catch (_) {}
  }
}

// ================================================================
// 🖥️  管理者設定画面 HTML（前回のadminHtml()をそのまま使用）
// ================================================================
function adminHtml() {
  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>名刺書き出し君 — 管理者設定</title>
<style>
/* ===== リセット ===== */
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}

/* ===== 変数 ===== */
:root{
  --blue:#1a73e8; --blue-dark:#1557b0; --blue-light:#e8f0fe;
  --green:#188038; --green-bg:#e6f4ea;
  --red:#c5221f;   --red-bg:#fce8e6;
  --yellow:#b06000;--yellow-bg:#fef7e0;
  --gray:#5f6368;  --border:#dadce0;
  --bg:#f1f3f4;    --card:#fff;
  --text:#202124;
  --radius:12px; --shadow:0 1px 3px rgba(0,0,0,.15);
}

/* ===== レイアウト ===== */
body{font-family:-apple-system,'Hiragino Sans',sans-serif;
  background:var(--bg);color:var(--text);font-size:15px;
  min-height:100vh;padding-bottom:60px}

header{background:var(--blue);color:#fff;padding:18px 24px;
  display:flex;align-items:center;gap:12px;
  position:sticky;top:0;z-index:100;box-shadow:0 2px 6px rgba(0,0,0,.2)}
header h1{font-size:19px;font-weight:700;letter-spacing:.02em}
header .ver{font-size:12px;opacity:.75;margin-left:auto;white-space:nowrap}

main{max-width:680px;margin:28px auto;padding:0 16px;display:flex;flex-direction:column;gap:20px}

/* ===== カード ===== */
.card{background:var(--card);border-radius:var(--radius);
  box-shadow:var(--shadow);overflow:hidden}
.card-header{padding:14px 20px;border-bottom:1px solid var(--border);
  display:flex;align-items:center;gap:8px}
.card-header h2{font-size:15px;font-weight:700}
.card-header .icon{font-size:18px;line-height:1}
.card-body{padding:20px}

/* ===== フォーム部品 ===== */
.field{margin-bottom:18px}
.field:last-child{margin-bottom:0}
.field label{display:block;font-size:13px;font-weight:600;
  color:var(--gray);margin-bottom:5px}
.field .hint{font-size:12px;color:var(--gray);margin-top:4px;line-height:1.5}

.input-wrap{position:relative;display:flex;align-items:center}
.input-wrap input,.input-wrap select{
  width:100%;padding:10px 12px;font-size:14px;
  border:2px solid var(--border);border-radius:8px;
  background:#fff;color:var(--text);transition:border-color .15s}
.input-wrap input:focus,.input-wrap select:focus{
  border-color:var(--blue);outline:none}
.input-wrap input.masked{font-family:monospace;letter-spacing:.05em}

.icon-btn{position:absolute;right:10px;background:none;border:none;
  cursor:pointer;font-size:16px;color:var(--gray);padding:4px;
  border-radius:4px;transition:color .15s}
.icon-btn:hover{color:var(--blue)}

.link-btn{display:inline-flex;align-items:center;gap:5px;
  padding:6px 12px;background:var(--blue-light);color:var(--blue);
  border:none;border-radius:6px;font-size:13px;font-weight:600;
  cursor:pointer;text-decoration:none;transition:background .15s;margin-top:6px}
.link-btn:hover{background:#d2e3fc}

/* ===== OCR選択タブ ===== */
.ocr-tabs{display:flex;gap:1px;background:var(--border);border-radius:8px;overflow:hidden}
.ocr-tab{flex:1;padding:10px;border:none;background:#f8f8f8;font-size:14px;
  cursor:pointer;transition:background .15s;font-weight:600;color:var(--gray)}
.ocr-tab.active{background:var(--blue);color:#fff}
.ocr-tab:first-child{border-radius:7px 0 0 7px}
.ocr-tab:last-child{border-radius:0 7px 7px 0}

.claude-box{margin-top:14px;padding:14px;background:#fafafe;
  border:1px solid #c5d8f7;border-radius:8px}
.cost-table{width:100%;border-collapse:collapse;margin-top:10px;font-size:13px}
.cost-table th{background:var(--blue-light);padding:6px 10px;text-align:left;
  font-weight:600;color:var(--blue)}
.cost-table td{padding:6px 10px;border-top:1px solid #eee}

/* ===== 接続テスト ===== */
.test-btn{display:flex;align-items:center;gap:8px;
  padding:10px 20px;background:#fff;border:2px solid var(--blue);
  color:var(--blue);border-radius:8px;font-size:14px;font-weight:700;
  cursor:pointer;transition:all .15s}
.test-btn:hover{background:var(--blue-light)}
.test-btn:disabled{opacity:.5;cursor:not-allowed}
.test-result{margin-top:12px;padding:12px 14px;border-radius:8px;
  font-size:14px;line-height:1.6;display:none}
.test-result.ok{background:var(--green-bg);color:#155724}
.test-result.ng{background:var(--red-bg);color:#7f1d1d}
.test-result.warn{background:var(--yellow-bg);color:var(--yellow)}
.test-detail{font-size:12px;margin-top:6px;opacity:.8}

/* ===== 保存ボタン ===== */
.save-btn{width:100%;padding:16px;background:var(--blue);color:#fff;
  border:none;border-radius:var(--radius);font-size:17px;font-weight:700;
  cursor:pointer;transition:background .15s;display:flex;align-items:center;
  justify-content:center;gap:10px}
.save-btn:hover{background:var(--blue-dark)}
.save-btn:disabled{opacity:.5;cursor:not-allowed}

/* ===== 保存結果 ===== */
.save-result{padding:14px;border-radius:var(--radius);font-size:15px;
  font-weight:600;text-align:center;display:none;margin-top:4px}
.save-result.ok{background:var(--green-bg);color:#155724}
.save-result.ng{background:var(--red-bg);color:#7f1d1d}

/* ===== スピナー ===== */
.spinner{width:18px;height:18px;border:3px solid rgba(255,255,255,.4);
  border-top-color:#fff;border-radius:50%;animation:spin .7s linear infinite;display:none}
@keyframes spin{to{transform:rotate(360deg)}}
.spinner.dark{border-color:rgba(26,115,232,.3);border-top-color:var(--blue)}

/* ===== ローディングオーバーレイ ===== */
#loadingOverlay{position:fixed;inset:0;background:rgba(255,255,255,.7);
  display:flex;align-items:center;justify-content:center;z-index:200;display:none}
.load-card{background:#fff;padding:28px 40px;border-radius:16px;
  box-shadow:0 4px 20px rgba(0,0,0,.15);text-align:center}
.load-card .spinner{width:36px;height:36px;border-width:4px;margin:0 auto 14px}
.load-card .spinner.dark{display:block}

/* ===== セクション区切り ===== */
.divider{height:1px;background:var(--border);margin:18px 0}

/* ===== セキュリティ状態バッジ ===== */
.badge{display:inline-flex;align-items:center;gap:4px;padding:3px 9px;
  border-radius:20px;font-size:12px;font-weight:600}
.badge.set{background:var(--green-bg);color:var(--green)}
.badge.unset{background:var(--red-bg);color:var(--red)}

/* ===== レスポンシブ ===== */
@media(max-width:500px){
  main{margin:16px auto;gap:14px}
  header h1{font-size:17px}
}
</style>
</head>
<body>

<!-- ===== ヘッダー ===== -->
<header>
  <span style="font-size:24px">🪪</span>
  <h1>名刺書き出し君 — 管理者設定</h1>
  <span class="ver">v1.0.0</span>
</header>

<div id="loadingOverlay">
  <div class="load-card">
    <div class="spinner dark"></div>
    <div id="loadingMsg" style="font-size:15px;color:#444">設定を読み込み中…</div>
  </div>
</div>

<main>

  <!-- ===== 1. スプレッドシート設定 ===== -->
  <div class="card">
    <div class="card-header">
      <span class="icon">📊</span>
      <h2>スプレッドシート設定</h2>
    </div>
    <div class="card-body">

      <div class="field">
        <label>スプレッドシートID <span style="color:var(--red)">*</span></label>
        <div class="input-wrap">
          <input id="spreadsheetId" type="text"
            placeholder="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms"
            autocomplete="off" spellcheck="false">
          <button class="icon-btn" title="コピー" onclick="copyVal('spreadsheetId')">📋</button>
        </div>
        <div class="hint">
          Google Sheets の URL の
          <code>/d/</code>と<code>/edit</code>の間の文字列<br>
          例: docs.google.com/spreadsheets/d/<strong>★ここ★</strong>/edit
        </div>
        <a id="ssLink" class="link-btn" href="#" target="_blank" style="display:none">
          🔗 スプレッドシートを開く
        </a>
      </div>

      <div class="divider"></div>

      <div class="field">
        <label>データシート名（タブ名）</label>
        <input id="sheetName" type="text" placeholder="名刺データ">
        <div class="hint">名刺データを書き込むシートのタブ名（デフォルト: 名刺データ）</div>
      </div>

      <div class="field" style="margin-bottom:0">
        <label>ログシート名（タブ名）</label>
        <input id="logSheetName" type="text" placeholder="処理ログ">
        <div class="hint">処理履歴を記録するシートのタブ名（デフォルト: 処理ログ）</div>
      </div>

      <div class="divider"></div>

      <!-- 接続テスト -->
      <button class="test-btn" id="testBtn" onclick="testConnection()">
        <span id="testSpinner" class="spinner dark" style="border-top-color:var(--blue)"></span>
        🔍 接続テスト
      </button>
      <div class="test-result" id="testResult"></div>

    </div>
  </div>

  <!-- ===== 2. OCRエンジン設定 ===== -->
  <div class="card">
    <div class="card-header">
      <span class="icon">🤖</span>
      <h2>OCRエンジン設定</h2>
    </div>
    <div class="card-body">

      <div class="field">
        <label>使用するOCRエンジン</label>
        <div class="ocr-tabs" id="ocrTabs">
          <button class="ocr-tab" id="tab-drive"  onclick="selectOcr('DRIVE_OCR')">
            ⚡ Drive OCR（無料）
          </button>
          <button class="ocr-tab" id="tab-claude" onclick="selectOcr('CLAUDE_API')">
            ✨ Claude API（高精度）
          </button>
        </div>
        <input type="hidden" id="ocrEngine" value="DRIVE_OCR">
        <div class="hint" id="ocrHint"></div>
      </div>

      <!-- Claude API 設定（条件表示） -->
      <div class="claude-box" id="claudeBox" style="display:none">
        <div class="field" style="margin-bottom:12px">
          <label>Claude API キー</label>
          <div class="input-wrap">
            <input id="claudeApiKey" type="password" class="masked"
              placeholder="sk-ant-api03-…（変更する場合のみ入力）"
              autocomplete="new-password" spellcheck="false">
            <button class="icon-btn" onclick="toggleVisible('claudeApiKey',this)">👁</button>
          </div>
          <div class="hint">
            現在の状態: <span id="claudeKeyStatus"></span><br>
            変更しない場合は空欄のままにしてください<br>
            取得先: <a href="https://console.anthropic.com" target="_blank" style="color:var(--blue)">console.anthropic.com</a>
          </div>
        </div>

        <label style="font-size:13px;font-weight:600;color:var(--gray)">料金目安（claude-haiku使用時）</label>
        <table class="cost-table">
          <tr><th>月間枚数</th><th>概算コスト</th></tr>
          <tr><td>100枚</td><td>約 ¥15</td></tr>
          <tr><td>500枚</td><td>約 ¥75</td></tr>
          <tr><td>1,000枚</td><td>約 ¥150</td></tr>
        </table>
      </div>

    </div>
  </div>

  <!-- ===== 3. セキュリティ設定 ===== -->
  <div class="card">
    <div class="card-header">
      <span class="icon">🔒</span>
      <h2>セキュリティ設定</h2>
    </div>
    <div class="card-body">

      <div class="field" style="margin-bottom:0">
        <label>APIシークレット（合言葉）</label>
        <div style="display:flex;align-items:center;gap:10px;margin-top:4px;flex-wrap:wrap">
          <span id="secretStatus" class="badge"></span>
          <span style="font-size:13px;color:var(--gray)">ショートカットとGASで同じ値を使用</span>
        </div>
        <div style="margin-top:10px">
          <button class="link-btn" onclick="toggleSecretEdit()" id="secretEditBtn">
            🔑 変更する
          </button>
        </div>
        <div id="secretEditBox" style="display:none;margin-top:10px">
          <div class="input-wrap">
            <input id="newApiSecret" type="password" class="masked"
              placeholder="新しいシークレット（6文字以上）"
              autocomplete="new-password" spellcheck="false">
            <button class="icon-btn" onclick="toggleVisible('newApiSecret',this)">👁</button>
          </div>
          <div class="hint" style="color:var(--red)">
            ⚠️ 変更後はショートカット内の secret 値も同じ値に変更してください
          </div>
        </div>
      </div>

    </div>
  </div>

  <!-- ===== 4. 保存 ===== -->
  <button class="save-btn" id="saveBtn" onclick="saveConfig()">
    <span id="saveSpinner" class="spinner"></span>
    <span id="saveBtnText">💾 設定を保存する</span>
  </button>
  <div class="save-result" id="saveResult"></div>

</main>

<script>
// =========================================================
// 設定（GAS URLはデプロイURLを自動取得）
// =========================================================
const GAS_URL = location.href.split('?')[0];
let   currentConfig = {};

// =========================================================
// 初期化
// =========================================================
window.addEventListener('load', loadConfig);

async function loadConfig() {
  showOverlay('設定を読み込み中…');
  try {
    const data = await getJson('?action=getConfig');
    currentConfig = data;
    applyConfig(data);
  } catch(e) {
    alert('設定の読み込みに失敗しました。\\nページを再読み込みしてください。\\n\\n' + e.message);
  } finally {
    hideOverlay();
  }
}

function applyConfig(cfg) {
  v('spreadsheetId', cfg.spreadsheetId || '');
  v('sheetName',     cfg.sheetName     || '');
  v('logSheetName',  cfg.logSheetName  || '');
  selectOcr(cfg.ocrEngine || 'DRIVE_OCR');

  // スプレッドシートリンク
  updateSsLink(cfg.spreadsheetId);

  // Claude API キーのヒント
  const hint = document.getElementById('claudeKeyStatus');
  if (hint) {
    hint.innerHTML = cfg.hasClaudeKey
      ? '<span class="badge set">✅ 設定済み&nbsp;' + (cfg.claudeKeyHint || '') + '</span>'
      : '<span class="badge unset">❌ 未設定</span>';
  }

  // APIシークレット状態
  const sb = document.getElementById('secretStatus');
  if (sb) {
    sb.className = 'badge ' + (cfg.hasApiSecret ? 'set' : 'unset');
    sb.textContent = cfg.hasApiSecret ? '✅ 設定済み' : '❌ 未設定';
  }
}

// =========================================================
// OCR タブ切り替え
// =========================================================
function selectOcr(engine) {
  document.getElementById('ocrEngine').value = engine;
  document.getElementById('tab-drive') .classList.toggle('active', engine === 'DRIVE_OCR');
  document.getElementById('tab-claude').classList.toggle('active', engine === 'CLAUDE_API');

  const box  = document.getElementById('claudeBox');
  const hint = document.getElementById('ocrHint');
  if (engine === 'CLAUDE_API') {
    box.style.display  = 'block';
    hint.textContent   = 'Claude API を使用します。API キーが必要です（要課金）。';
  } else {
    box.style.display  = 'none';
    hint.textContent   = 'Google Drive の無料 OCR を使用します。日本語の精度はやや低めです。';
  }
}

// =========================================================
// 接続テスト
// =========================================================
async function testConnection() {
  const id     = gv('spreadsheetId');
  const secret = currentConfig.hasApiSecret ? promptSecret() : '';

  if (!id) {
    showTestResult('ng', '❌ スプレッドシートIDを入力してください');
    return;
  }

  const btn = document.getElementById('testBtn');
  const sp  = document.getElementById('testSpinner');
  btn.disabled = true; sp.style.display = 'block';
  hideTestResult();

  try {
    const url = '?action=testConnection&secret=' + encodeURIComponent(secret);
    const res = await getJson(url);

    if (res.ok) {
      const cls  = res.sheetFound ? 'ok' : 'warn';
      const icon = res.sheetFound ? '✅' : '⚠️';
      showTestResult(cls,
        icon + ' ' + res.message,
        res.sheetFound
          ? 'データ行数: ' + Math.max(0, res.lastRow - 1) + ' 件 / ログシート: ' + (res.logFound ? 'あり' : '未作成（初回保存時に自動作成）')
          : '「設定を保存」後に initialSetup() を実行してください'
      );
    } else {
      showTestResult('ng', '❌ ' + res.message);
    }
  } catch(e) {
    showTestResult('ng', '❌ 通信エラー: ' + e.message);
  } finally {
    btn.disabled = false; sp.style.display = 'none';
  }
}

function promptSecret() {
  // APIシークレットをその場で入力させる（セキュリティ上ページには保持しない）
  return prompt('APIシークレット（合言葉）を入力してください') || '';
}

function showTestResult(cls, msg, detail) {
  const el = document.getElementById('testResult');
  el.style.display = 'block';
  el.className = 'test-result ' + cls;
  el.innerHTML = msg + (detail ? '<div class="test-detail">' + detail + '</div>' : '');
}
function hideTestResult() {
  const el = document.getElementById('testResult');
  el.style.display = 'none'; el.textContent = '';
}

// =========================================================
// 設定保存
// =========================================================
async function saveConfig() {
  // 簡易バリデーション
  const ssId = gv('spreadsheetId');
  if (!ssId) { alert('スプレッドシートIDを入力してください'); return; }
  if (ssId.includes('/') || ssId.includes('http')) {
    alert('スプレッドシートIDが正しくありません。\\nURLではなく /d/ と /edit の間の文字列を貼り付けてください');
    return;
  }

  const secret = promptSecret();
  if (!secret) { alert('設定を保存するには APIシークレットの入力が必要です'); return; }

  const payload = {
    adminAction:   'saveConfig',
    secret:        secret,
    spreadsheetId: ssId,
    sheetName:     gv('sheetName')     || '名刺データ',
    logSheetName:  gv('logSheetName')  || '処理ログ',
    ocrEngine:     gv('ocrEngine'),
    newApiSecret:  gv('newApiSecret'),
    claudeApiKey:  gv('claudeApiKey'),
  };

  const btn = document.getElementById('saveBtn');
  const sp  = document.getElementById('saveSpinner');
  const tx  = document.getElementById('saveBtnText');
  btn.disabled = true; sp.style.display = 'block'; tx.textContent = '保存中…';
  hideSaveResult();

  try {
    const res = await postJson(GAS_URL, payload);
    if (res.ok) {
      showSaveResult('ok', '✅ ' + res.message);
      updateSsLink(ssId);
      await loadConfig(); // 最新値を再取得
    } else {
      showSaveResult('ng', '❌ ' + res.message);
    }
  } catch(e) {
    showSaveResult('ng', '❌ 通信エラー: ' + e.message);
  } finally {
    btn.disabled = false; sp.style.display = 'none'; tx.textContent = '💾 設定を保存する';
  }
}

function showSaveResult(cls, msg) {
  const el = document.getElementById('saveResult');
  el.style.display = 'block'; el.className = 'save-result ' + cls; el.textContent = msg;
  setTimeout(() => { el.style.display = 'none'; }, 6000);
}
function hideSaveResult() {
  const el = document.getElementById('saveResult');
  el.style.display = 'none'; el.textContent = '';
}

// =========================================================
// UI ヘルパー
// =========================================================
function toggleSecretEdit() {
  const box = document.getElementById('secretEditBox');
  const btn = document.getElementById('secretEditBtn');
  const open = box.style.display === 'none';
  box.style.display = open ? 'block' : 'none';
  btn.textContent   = open ? '✕ キャンセル' : '🔑 変更する';
}

function toggleVisible(inputId, btn) {
  const el = document.getElementById(inputId);
  const show = el.type === 'password';
  el.type = show ? 'text' : 'password';
  btn.textContent = show ? '🙈' : '👁';
}

function copyVal(inputId) {
  const val = document.getElementById(inputId).value;
  if (!val) return;
  navigator.clipboard.writeText(val)
    .then(() => showToast('コピーしました'))
    .catch(() => showToast('コピー失敗'));
}

function updateSsLink(id) {
  const a = document.getElementById('ssLink');
  if (id && !id.includes('/')) {
    a.href = 'https://docs.google.com/spreadsheets/d/' + id + '/edit';
    a.style.display = 'inline-flex';
  } else {
    a.style.display = 'none';
  }
}

function showToast(msg) {
  let t = document.getElementById('toast');
  if (!t) {
    t = document.createElement('div');
    t.id = 'toast';
    Object.assign(t.style, {
      position:'fixed',bottom:'24px',left:'50%',transform:'translateX(-50%)',
      background:'#333',color:'#fff',padding:'8px 18px',borderRadius:'20px',
      fontSize:'13px',zIndex:'999',transition:'opacity .3s',whiteSpace:'nowrap'
    });
    document.body.appendChild(t);
  }
  t.textContent = msg; t.style.opacity = '1';
  clearTimeout(t._timer);
  t._timer = setTimeout(() => { t.style.opacity = '0'; }, 2000);
}

function showOverlay(msg) {
  document.getElementById('loadingMsg').textContent = msg;
  document.getElementById('loadingOverlay').style.display = 'flex';
  document.querySelector('#loadingOverlay .spinner').style.display = 'block';
}
function hideOverlay() {
  document.getElementById('loadingOverlay').style.display = 'none';
}

// =========================================================
// 値の取得・設定ショートハンド
// =========================================================
function gv(id) { const el = document.getElementById(id); return el ? el.value.trim() : ''; }
function v(id, val) { const el = document.getElementById(id); if (el) el.value = val; }

// =========================================================
// Fetch ユーティリティ
// =========================================================
async function getJson(url) {
  const res = await fetch(GAS_URL + (url.startsWith('?') ? url : '?' + url));
  if (!res.ok) throw new Error('HTTP ' + res.status);
  return res.json();
}

async function postJson(url, data) {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  if (!res.ok) throw new Error('HTTP ' + res.status);
  return res.json();
}
</script>
</body>
</html>`;
}
