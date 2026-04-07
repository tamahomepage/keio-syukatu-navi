/**
 * 慶應就活ナビ — GAS メインスクリプト
 *
 * 【セットアップ手順】
 * 1. このファイルをGASプロジェクトに貼り付ける
 * 2. スクリプトプロパティに以下を設定する（プロジェクトの設定 → スクリプトプロパティ）
 *      ANTHROPIC_API_KEY      : sk-ant-api03-... (AnthropicのAPIキー)
 *      AUTH_SPREADSHEET_ID    : 18cQuF3tp4o35rDKCEXwA8c2dkWEjsggA1QHAkoNhnB4（会員データ）
 *      AUTH_LINE_QR_FOLDER_ID : （LINE QR保存先DriveフォルダID）← 省略するとマイドライブ直下
 *    ※ 各シートのスプレッドシートIDは下の SHEET_KEYS に直接記載済み
 * 3. 「デプロイ」→「新しいデプロイ」→ 種類:ウェブアプリ
 *      実行ユーザー: 自分
 *      アクセス: 全員（匿名を含む）
 * 4. デプロイURLをmembers.html / auth.js の GAS_PROXY_URL に貼る
 */

// ─────────────────────────────────────────────
//  シートキー → スプレッドシート設定のマップ
//  各キーに spreadsheetId を直接記載
// ─────────────────────────────────────────────
var SHEET_KEYS = {
  // 選考情報まとめ（A:企業名 B:締切日 C:種別 D:リンク E:業界 F:特集）
  selection:    { spreadsheetId: '1Rm6H6d2BcduDosoagcJjF3I78xOeuimNb3p80R6F3qE', sheetName: 'selection'   },
  // 企業分析ノート（A:企業名 B:業界 C:事業内容 D:強み E:志望動機 F:選考フロー G:URL H:一般呼称）
  company:      { spreadsheetId: '1axjcyNPqBiD4SiNhxDx3NaVALelMEkMt2x2ng0Sk7Ww', sheetName: 'company'    },
  // 会員向けイベント選考案内（A:タイトル B:日時 C:場所 D:詳細 E:リンク F:締切）
  priority:     { spreadsheetId: '1Q8Ixh34Ffzy4imnugj4Nyp9mN65A-PUw_PXOj2ATA58', sheetName: 'priority'   },
  // 限定資料&ファイル（A:タイトル B:種別 C:説明 D:リンク E:日付）
  files:        { spreadsheetId: '1N-9COcRmmf1TeNU1EHqyGqFzHzn02ngNlPXAsBbtnsg', sheetName: 'files'      },
  // 輪読会貸出状況（A:タイトル B:著者 C:状態(貸出可/貸出中) D:保有者 E:説明）
  books:        { spreadsheetId: '1hLFipZMnS03sc5aWb_pmCKXBN7utQaAQRXImg4uQhrs', sheetName: 'books'      },
  // 就活ナビのイベント情報（A:タイトル B:日時 C:場所 D:詳細 E:リンク F:締切）
  event:        { spreadsheetId: '1ASykgCNJirBmB76yqxQ4H0jrpqK8vXbx2izLSdHo1To', sheetName: 'event'      },
  // 企業イベント情報
  corp_event:   { spreadsheetId: '1lr-FNeqRi5l41BafVJRc9Upkwn0VpLNmBvlM7moIgw8', sheetName: 'event'      },
  // 新歓イベント情報
  shinkan:      { spreadsheetId: '1ujQ0RcbAqmIormMO-8I1THUSNvtuu0CjHll6sdc_Wrc', sheetName: 'shinkan'    },
  // 選考体験記
  selection_exp:{ spreadsheetId: '1YRfephgVy8moJLN76MFaIRYUy1xfi4pm7E5JZLvsg34', sheetName: 'selection'  },
  // FAQ
  faq:          { spreadsheetId: '1hJqOds7fOxqLmau-IoHhfQLQC7_ttun-SmGZZSzF76M', sheetName: 'faq'        },
};

var READ_PUBLIC_SHEET_KEYS_ = {
  event: true,
  faq: true,
  shinkan: true,
  orientation: true
};

// ─────────────────────────────────────────────
//  エントリーポイント
// ─────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '慶應就活ナビ API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var payload = {};
  try {
    payload = JSON.parse(e.postData.contents || '{}');
  } catch (err) {
    return jsonResponse_({ status: 'error', message: 'リクエストのパースに失敗しました。' });
  }

  // 認証・AI系アクション
  var authResponse = handleAuthAction_(payload);
  if (authResponse) return authResponse;

  // サイトデータ系アクション
  try {
    var action = String(payload.action || '');
    if (action === 'read')     return jsonResponse_(handleRead_(payload));
    if (action === 'generate') return jsonResponse_(handleGenerate_(payload));
    if (action === 'save')     return jsonResponse_(handleSave_(payload));
    if (action === 'fetchRecruitInfo') return jsonResponse_(handleFetchRecruitInfo_(payload));
    if (action === 'saveRecruitInfo')  return jsonResponse_(handleSaveRecruitInfo_(payload));
    if (action === 'logActivity')       return jsonResponse_(handleLogActivity_(payload));
    if (action === 'getActivitySummary') return jsonResponse_(handleGetActivitySummary_(payload));

    return jsonResponse_({ status: 'error', message: '不明なアクション: ' + action });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: err.message || '処理中にエラーが発生しました。' });
  }
}

// ─────────────────────────────────────────────
//  action: read — シートのデータを返す
// ─────────────────────────────────────────────
function handleRead_(payload) {
  var key    = String(payload.sheetKey || '');
  var config = SHEET_KEYS[key];
  if (!config) return { status: 'error', message: 'sheetKey が不正です: ' + key };
  if (!READ_PUBLIC_SHEET_KEYS_[key]) {
    getActiveSessionOrThrow_(payload.sessionToken);
  }

  var sheet = getSiteSheet_(config.sheetName, false, config.spreadsheetId);
  if (!sheet) return { status: 'ok', rows: [] };

  var values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return { status: 'ok', rows: [] };

  return { status: 'ok', rows: values.slice(1) }; // 1行目はヘッダー
}

// ─────────────────────────────────────────────
//  action: generate — 公式HPから事実を取得し、AIが独自分析
// ─────────────────────────────────────────────
function handleGenerate_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  assertAiUsageAllowed_(session.userId, 'generate');
  var companyName = trimText_(payload.companyName);
  if (!companyName) return { status: 'error', message: '企業名を入力してください。' };

  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { status: 'error', message: 'ANTHROPIC_API_KEY が設定されていません。' };

  // Step 1: Claude に公式HPのURLを推測させる
  var urlResp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: 'claude-sonnet-4-20250514', max_tokens: 300,
      system: '企業の公式サイトURLに詳しいアシスタントです。企業名を受け取り、その企業の公式コーポレートサイトと公式採用ページのURLを返してください。URLのみをJSON形式で返してください。',
      messages: [{ role: 'user', content: companyName + ' の公式サイトURLと採用ページURL\n出力: {"corporate":"https://...","recruit":"https://..."}' }]
    }),
    muteHttpExceptions: true
  });

  var corpUrl = '', recruitUrl = '', sourceUrls = [];
  if (urlResp.getResponseCode() === 200) {
    var urlText = '';
    (JSON.parse(urlResp.getContentText()).content || []).forEach(function(b) { if (b.type === 'text') urlText += b.text; });
    var urlJson = urlText.match(/\{[\s\S]*\}/);
    if (urlJson) {
      try {
        var urls = JSON.parse(urlJson[0]);
        corpUrl = trimText_(urls.corporate || '');
        recruitUrl = trimText_(urls.recruit || '');
      } catch(e) {}
    }
  }

  // Step 1.5: ユーザーが手動URLを指定した場合はそちらを優先
  var manualCorpUrl = trimText_(payload.manualCorpUrl || '');
  var manualRecruitUrl = trimText_(payload.manualRecruitUrl || '');
  if (manualCorpUrl) corpUrl = manualCorpUrl;
  if (manualRecruitUrl) recruitUrl = manualRecruitUrl;

  // Step 2: ドメインチェック＋公式HPからテキストを取得
  corpUrl = normalizeSafeFetchUrl_(corpUrl);
  recruitUrl = normalizeSafeFetchUrl_(recruitUrl);
  var corpText = '', recruitText = '';
  if (corpUrl) {
    if (isDomainBlocked_(corpUrl)) { corpUrl = ''; }
    else {
      try {
        var cResp = UrlFetchApp.fetch(corpUrl, { muteHttpExceptions: true, followRedirects: true });
        if (cResp.getResponseCode() === 200) {
          corpText = cResp.getContentText().replace(/<script[\s\S]*?<\/script>/gi,'').replace(/<style[\s\S]*?<\/style>/gi,'').replace(/<[^>]+>/g,' ').replace(/\s+/g,' ').trim();
          if (corpText.length > 8000) corpText = corpText.substring(0, 8000);
          sourceUrls.push(corpUrl);
        }
      } catch(e) {}
    }
  }
  if (recruitUrl && recruitUrl !== corpUrl) {
    if (isDomainBlocked_(recruitUrl)) { recruitUrl = ''; }
    else {
      try {
        var rResp = UrlFetchApp.fetch(recruitUrl, { muteHttpExceptions: true, followRedirects: true });
        if (rResp.getResponseCode() === 200) {
          recruitText = rResp.getContentText().replace(/<script[\s\S]*?<\/script>/gi,'').replace(/<style[\s\S]*?<\/style>/gi,'').replace(/<[^>]+>/g,' ').replace(/\s+/g,' ').trim();
          if (recruitText.length > 8000) recruitText = recruitText.substring(0, 8000);
          sourceUrls.push(recruitUrl);
        }
      } catch(e) {}
    }
  }

  // Step 3: 取得したテキストを元にAIが独自分析
  var contextBlock = '';
  if (corpText || recruitText) {
    contextBlock = '\n\n【取得した公式サイトの情報（事実データ）】\n';
    if (corpText) contextBlock += '＜コーポレートサイト: ' + corpUrl + '＞\n' + corpText.substring(0, 5000) + '\n\n';
    if (recruitText) contextBlock += '＜採用ページ: ' + recruitUrl + '＞\n' + recruitText.substring(0, 5000) + '\n\n';
    contextBlock += '上記の公式情報を元に、以下のルールで独自分析してください:\n';
    contextBlock += '- データ層: 公式HPから読み取れる事実のみを根拠にする\n';
    contextBlock += '- AI層: 事実を元にした独自の分析・考察を行う（他サイトの表現をコピーしない）\n';
    contextBlock += '- 原文の文章をそのまま引用しない。自分の言葉で要約・分析する\n';
  } else {
    contextBlock = '\n\n※公式HPの取得に失敗しました。一般的に知られている事実のみを元に、独自の分析として出力してください。他サイトの文章を再構成してはいけません。\n';
  }

  var systemPrompt = 'あなたは就活のプロフェッショナルアナリストです。企業の公式情報を元に、独自の視点で企業分析ノートを作成します。\n\n重要なルール:\n- 公式HPから読み取れる事実を根拠に、自分の言葉で分析する\n- 他の就活サイト（ワンキャリア、マイナビ等）の文章を再現・模倣しない\n- 「強み」「志望動機のポイント」は独自の分析観点で記述する\n- 企業が不明な場合は status: "NOT_FOUND" を返す';

  var userMessage = '企業名: ' + companyName + contextBlock +
    '\n\n以下のJSON形式で出力してください:\n{\n  "status": "ok",\n  "officialName": "正式な企業名",\n  "industry": "業界",\n  "business": "公式HPの情報を元にした事業内容の要約（200字以内）",\n  "strength": "公式情報から読み取れる強み・特徴の独自分析（150字以内）",\n  "motive": "この企業を志望する際に考えるべき観点（150字以内。独自の分析視点で）",\n  "process": "採用ページから読み取れる選考フロー（不明なら「要確認」）",\n  "url": "' + (sourceUrls.join(' | ') || '取得失敗') + '",\n  "source": "出典元の説明"\n}\n企業が不明な場合: { "status": "NOT_FOUND" }';

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: 'claude-sonnet-4-20250514', max_tokens: 1200,
      system: systemPrompt,
      messages: [{ role: 'user', content: userMessage }]
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    var errBody = {};
    try { errBody = JSON.parse(response.getContentText()); } catch (e) {}
    throw new Error('AI生成エラー: ' + ((errBody.error && errBody.error.message) || response.getResponseCode()));
  }

  var apiData = JSON.parse(response.getContentText());
  var rawText = '';
  (apiData.content || []).forEach(function (block) {
    if (block.type === 'text') rawText += block.text;
  });

  var match = rawText.match(/\{[\s\S]*\}/);
  if (!match) throw new Error('AIの応答からJSONを取得できませんでした。');

  var note = JSON.parse(match[0]);
  if (note.status === 'NOT_FOUND') return { status: 'error', message: 'NOT_FOUND' };

  // 出典情報＋ウォーターマークを追加
  note.sourceUrls = sourceUrls;
  note.source = sourceUrls.length > 0
    ? 'AI独自分析（出典: ' + sourceUrls.join(', ') + '）'
    : 'AI独自分析（公式HP取得失敗のため一般情報に基づく）';
  note.disclaimer = '※この分析は企業公式HPの情報に基づくAIの独自分析です。正確性は保証されません。必ず公式情報をご確認ください。';

  return { status: 'ok', note: note };
}

// ─────────────────────────────────────────────
//  action: save — 生成した企業ノートをスプシに保存
// ─────────────────────────────────────────────
function handleSave_(payload) {
  getActiveSessionOrThrow_(payload.sessionToken);
  var note        = payload.note || {};
  var searchWord  = trimText_(payload.searchWord);
  var officialName = trimText_(note.officialName || searchWord);
  if (!officialName) return { status: 'error', message: '企業名がありません。' };

  var companyConfig = SHEET_KEYS['company'];
  var sheet = getSiteSheet_(companyConfig.sheetName, true, companyConfig.spreadsheetId);

  // すでに同名の行があれば更新
  var values = sheet.getDataRange().getValues();
  var targetRow = -1;
  for (var i = 1; i < values.length; i++) {
    if (trimText_(values[i][0]) === officialName) { targetRow = i + 1; break; }
  }

  var alias = (searchWord && searchWord !== officialName) ? searchWord : '';
  var row = [
    officialName,
    trimText_(note.industry  || ''),
    trimText_(note.business  || ''),
    trimText_(note.strength  || ''),
    trimText_(note.motive    || ''),
    trimText_(note.process   || ''),
    trimText_(note.url       || ''),
    alias
  ];

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return { status: 'ok' };
}

// ─────────────────────────────────────────────
//  スプレッドシートユーティリティ
// ─────────────────────────────────────────────
function getSiteSheet_(sheetName, createIfMissing, spreadsheetId) {
  var ss = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(sheetName);
    if (sheetName === 'company') {
      sheet.appendRow(['企業名','業界','事業内容','強み','志望動機メモ','選考フロー','URL','一般呼称']);
    }
  }
  return sheet || null;
}

function trimText_(value) {
  return String(value || '').trim();
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}


// ══════════════════════════════════════════════════════════════════════
//  以下：認証・AI呼び出し（gas-auth-extension.gs の内容をそのまま含む）
// ══════════════════════════════════════════════════════════════════════

var AUTH_SHEET_NAMES_ = {
  users:    'auth_users',
  sessions: 'auth_sessions',
  liked:    'auth_liked'
};

var AUTH_USER_HEADERS_ = [
  'id','username','usernameKey','desiredIndustry','passwordHash','salt',
  'createdAt','updatedAt','preferredCompany1','preferredCompany2','preferredCompany3',
  'lineName','lineQrDriveFileId','lineQrDriveUrl'
];

var AUTH_SESSION_TTL_DAYS_ = 7;
var AUTH_LOGIN_RATE_LIMIT_ = {
  maxAttempts: 8,
  windowSeconds: 15 * 60,
  message: 'ログイン試行回数が多すぎます。15分ほど待ってから再試行してください。'
};
var AUTH_AI_RATE_LIMIT_ = {
  maxAttempts: 20,
  windowSeconds: 60 * 60,
  message: 'AI機能の利用回数が上限に達しました。1時間ほど待ってから再試行してください。'
};
var BOARD_SHEET_NAME_          = 'es_board';
var BOARD_COMMENTS_SHEET_NAME_ = 'es_board_comments';
var BOARD_MAX_POSTS_           = 50;

function getAuthRateLimitKey_(action, value) {
  return ['auth-rate', action, normalizeAuthKey_(value || 'unknown')].join(':');
}

function getAuthRateLimitCount_(action, value) {
  var cache = CacheService.getScriptCache();
  return parseInt(cache.get(getAuthRateLimitKey_(action, value)) || '0', 10) || 0;
}

function assertAuthRateLimit_(action, value, options) {
  if (getAuthRateLimitCount_(action, value) >= options.maxAttempts) {
    throw new Error(options.message);
  }
}

function recordAuthRateLimitEvent_(action, value, options) {
  var cache = CacheService.getScriptCache();
  var key = getAuthRateLimitKey_(action, value);
  var count = getAuthRateLimitCount_(action, value) + 1;
  cache.put(key, String(count), options.windowSeconds);
}

function clearAuthRateLimit_(action, value) {
  CacheService.getScriptCache().remove(getAuthRateLimitKey_(action, value));
}

function assertAiUsageAllowed_(userId, actionName) {
  var rateKey = trimAuthText_(userId) + ':' + trimAuthText_(actionName || 'ai');
  assertAuthRateLimit_('ai', rateKey, AUTH_AI_RATE_LIMIT_);
  recordAuthRateLimitEvent_('ai', rateKey, AUTH_AI_RATE_LIMIT_);
}

function handleAuthAction_(payload) {
  if (!payload || typeof payload !== 'object') return null;
  var action = String(payload.action || '');
  if (!action.match(/^(auth|writeBoardPost|readBoardPosts|addBoardComment|callClaude|createPeerFeedback|submitPeerFeedback|getPeerFeedback)/)) return null;

  try {
    switch (action) {
      case 'authRegister':        return jsonResponse_(authRegister_(payload));
      case 'authLogin':           return jsonResponse_(authLogin_(payload));
      case 'authLogout':          return jsonResponse_(authLogout_(payload));
      case 'authUpdateProfile':   return jsonResponse_(authUpdateProfile_(payload));
      case 'authChangePassword':  return jsonResponse_(authChangePassword_(payload));
      case 'authSetLikedCompanies': return jsonResponse_(authSetLikedCompanies_(payload));
      case 'writeBoardPost':      return jsonResponse_(writeBoardPost_(payload));
      case 'readBoardPosts':      return jsonResponse_(readBoardPosts_(payload));
      case 'addBoardComment':     return jsonResponse_(addBoardComment_(payload));
      case 'callClaude':          return jsonResponse_(callClaude_(payload));
      case 'createPeerFeedbackRequest': return jsonResponse_(createPeerFeedbackRequest_(payload));
      case 'submitPeerFeedback':        return jsonResponse_(submitPeerFeedback_(payload));
      case 'getPeerFeedbackResults':    return jsonResponse_(getPeerFeedbackResults_(payload));
      case 'getPeerFeedbackInfo':       return jsonResponse_(getPeerFeedbackInfo_(payload));
      default: return null;
    }
  } catch (error) {
    return jsonResponse_({ status: 'error', message: error && error.message ? error.message : '処理でエラーが発生しました。' });
  }
}

// ── アカウント登録 ──────────────────────────────────
function authRegister_(payload) {
  ensureAuthSheets_();
  var username    = trimAuthText_(payload.username);
  var desiredIndustry = trimAuthText_(payload.desiredIndustry);
  var password    = trimAuthText_(payload.password);
  var preferredCompanies = getPreferredCompanies_(payload);
  var lineName    = trimAuthText_(payload.lineName);
  var lineQrDataUrl  = trimAuthText_(payload.lineQrDataUrl);
  var lineQrFileName = trimAuthText_(payload.lineQrFileName);
  var usernameKey = normalizeAuthKey_(username);

  assertAuth_(username,                       'ユーザー名を入力してください。');
  assertAuth_(desiredIndustry,                '志望業界を選択してください。');
  assertAuth_(preferredCompanies.length > 0,  '第1志望の企業名を入力してください。');
  assertAuth_(lineName,                       'LINE名を入力してください。');
  assertAuth_(lineQrDataUrl,                  'LINE QRをアップロードしてください。');
  assertAuth_(password.length >= 8,           'パスワードは8文字以上で設定してください。');

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  assertAuth_(!users.find(function (r) { return normalizeAuthKey_(r.usernameKey || r.username) === usernameKey; }), 'このユーザー名はすでに使われています。');

  var now    = new Date().toISOString();
  var userId = 'user_' + new Date().getTime().toString(36) + Utilities.getUuid().replace(/-/g,'').slice(0,6);
  var hashRecord = buildPasswordHashRecord_(password);
  var salt   = hashRecord.salt;
  var passwordHash = hashRecord.passwordHash;
  var lineQrAsset  = saveLineQrAsset_(userId, lineQrDataUrl, lineQrFileName, '');

  usersSheet.appendRow([userId, username, usernameKey, desiredIndustry, passwordHash, salt, now, now,
    preferredCompanies[0]||'', preferredCompanies[1]||'', preferredCompanies[2]||'',
    lineName, lineQrAsset.fileId, lineQrAsset.url]);
  saveLikedCompaniesForUser_(userId, []);
  var sessionToken = createSessionForUser_(userId);
  return { status:'ok', sessionToken: sessionToken, user: sanitizeUserRecord_({ id:userId, username, usernameKey, desiredIndustry, createdAt:now, updatedAt:now, preferredCompany1:preferredCompanies[0]||'', preferredCompany2:preferredCompanies[1]||'', preferredCompany3:preferredCompanies[2]||'', lineName, lineQrDriveFileId:lineQrAsset.fileId, lineQrDriveUrl:lineQrAsset.url }), likedCompanies:[] };
}

// ── ログイン ───────────────────────────────────────
function authLogin_(payload) {
  ensureAuthSheets_();
  var usernameKey = normalizeAuthKey_(payload.username);
  var password    = trimAuthText_(payload.password);
  assertAuth_(usernameKey, 'ユーザー名を入力してください。');
  assertAuth_(password,    'パスワードを入力してください。');
  assertAuthRateLimit_('login', usernameKey, AUTH_LOGIN_RATE_LIMIT_);

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var user  = users.find(function (r) { return normalizeAuthKey_(r.usernameKey || r.username) === usernameKey; });
  if (!user || !verifyPasswordRecord_(password, user)) {
    recordAuthRateLimitEvent_('login', usernameKey, AUTH_LOGIN_RATE_LIMIT_);
    throw new Error('ログイン情報が正しくありません。');
  }
  maybeUpgradePasswordHash_(usersSheet, users.indexOf(user), user, password);
  clearAuthRateLimit_('login', usernameKey);

  var sessionToken = createSessionForUser_(user.id);
  return { status:'ok', sessionToken, user: sanitizeUserRecord_(user), likedCompanies: getLikedCompaniesForUser_(user.id) };
}

// ── ログアウト ─────────────────────────────────────
function authLogout_(payload) {
  ensureAuthSheets_();
  var token = trimAuthText_(payload.sessionToken);
  if (!token) return { status:'ok' };
  var sessionsSheet = getAuthSheet_(AUTH_SHEET_NAMES_.sessions);
  var sessions = getSheetRecords_(sessionsSheet);
  var idx = findRecordIndex_(sessions, function (r) { return r.sessionToken === token; });
  if (idx >= 0) sessionsSheet.getRange(idx+2, 6).setValue('0');
  return { status:'ok' };
}

// ── プロフィール更新 ───────────────────────────────
function authUpdateProfile_(payload) {
  ensureAuthSheets_();
  var session  = getActiveSessionOrThrow_(payload.sessionToken);
  var username = trimAuthText_(payload.username);
  var desiredIndustry = trimAuthText_(payload.desiredIndustry);
  var preferredCompanies = getPreferredCompanies_(payload);
  var lineName = trimAuthText_(payload.lineName);
  var lineQrDataUrl  = trimAuthText_(payload.lineQrDataUrl);
  var lineQrFileName = trimAuthText_(payload.lineQrFileName);
  var usernameKey = normalizeAuthKey_(username);

  assertAuth_(username,                      'ユーザー名を入力してください。');
  assertAuth_(desiredIndustry,               '志望業界を選択してください。');
  assertAuth_(preferredCompanies.length > 0, '第1志望の企業名を入力してください。');
  assertAuth_(lineName,                      'LINE名を入力してください。');

  var usersSheet  = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users       = getSheetRecords_(usersSheet);
  var currentIndex = findRecordIndex_(users, function (r) { return r.id === session.userId; });
  assertAuth_(currentIndex >= 0, 'アカウントが見つかりません。');
  assertAuth_(!users.find(function (r) { return r.id !== session.userId && normalizeAuthKey_(r.usernameKey||r.username) === usernameKey; }), 'このユーザー名はすでに使われています。');

  var user = users[currentIndex];
  var lineQrAsset = { fileId: trimAuthText_(user.lineQrDriveFileId), url: trimAuthText_(user.lineQrDriveUrl) };
  if (!lineQrDataUrl) assertAuth_(lineQrAsset.url, 'LINE QRをアップロードしてください。');
  if (lineQrDataUrl)  lineQrAsset = saveLineQrAsset_(session.userId, lineQrDataUrl, lineQrFileName, lineQrAsset.fileId);

  var updatedAt = new Date().toISOString();
  usersSheet.getRange(currentIndex+2, 2, 1, 3).setValues([[username, usernameKey, desiredIndustry]]);
  usersSheet.getRange(currentIndex+2, 8).setValue(updatedAt);
  usersSheet.getRange(currentIndex+2, 9, 1, 6).setValues([[preferredCompanies[0]||'', preferredCompanies[1]||'', preferredCompanies[2]||'', lineName, lineQrAsset.fileId, lineQrAsset.url]]);
  return { status:'ok', user: sanitizeUserRecord_({ id:session.userId, username, usernameKey, desiredIndustry, createdAt:user.createdAt||'', updatedAt, preferredCompany1:preferredCompanies[0]||'', preferredCompany2:preferredCompanies[1]||'', preferredCompany3:preferredCompanies[2]||'', lineName, lineQrDriveFileId:lineQrAsset.fileId, lineQrDriveUrl:lineQrAsset.url }) };
}

// ── パスワード変更 ─────────────────────────────────
function authChangePassword_(payload) {
  ensureAuthSheets_();
  var session         = getActiveSessionOrThrow_(payload.sessionToken);
  var currentPassword = trimAuthText_(payload.currentPassword);
  var nextPassword    = trimAuthText_(payload.nextPassword);
  assertAuth_(currentPassword && nextPassword, '現在のパスワードと新しいパスワードを入力してください。');
  assertAuth_(nextPassword.length >= 8,        '新しいパスワードは8文字以上で設定してください。');

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var idx = findRecordIndex_(users, function (r) { return r.id === session.userId; });
  assertAuth_(idx >= 0, 'アカウントが見つかりません。');
  var user = users[idx];
  assertAuth_(verifyPasswordRecord_(currentPassword, user), '現在のパスワードが正しくありません。');
  var nextHashRecord = buildPasswordHashRecord_(nextPassword);
  var now = new Date().toISOString();
  usersSheet.getRange(idx+2, 5, 1, 2).setValues([[nextHashRecord.passwordHash, nextHashRecord.salt]]);
  usersSheet.getRange(idx+2, 8).setValue(now);
  invalidateSessionsForUser_(session.userId, '');
  return { status:'ok', sessionToken: createSessionForUser_(session.userId) };
}

// ── いいね保存 ─────────────────────────────────────
function authSetLikedCompanies_(payload) {
  ensureAuthSheets_();
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var liked   = Array.isArray(payload.likedCompanies) ? payload.likedCompanies : [];
  saveLikedCompaniesForUser_(session.userId, liked);
  return { status:'ok', likedCompanies: getLikedCompaniesForUser_(session.userId) };
}

// ── ESボード ───────────────────────────────────────
function writeBoardPost_(payload) {
  var session  = getActiveSessionOrThrow_(payload.sessionToken);
  var company  = trimAuthText_(payload.company);
  var question = trimAuthText_(payload.question);
  var esText   = trimAuthText_(payload.esText);
  assertAuth_(company,  '企業名を入力してください。');
  assertAuth_(question, '設問テキストを入力してください。');
  assertAuth_(esText,   'ESの内容を入力してください。');

  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(BOARD_SHEET_NAME_);
  if (!sheet) { sheet = ss.insertSheet(BOARD_SHEET_NAME_); sheet.appendRow(['id','userId','username','company','question','esText','createdAt']); }

  var now      = new Date().toISOString();
  var id       = now.replace(/\D/g,'').slice(0,14)+'_'+Utilities.getUuid().replace(/-/g,'').slice(0,8);
  var users    = getSheetRecords_(getAuthSheet_(AUTH_SHEET_NAMES_.users));
  var userRec  = users.find(function (r) { return r.id === session.userId; });
  var username = userRec ? trimAuthText_(userRec.username) : '';
  sheet.appendRow([id, session.userId, username, company, question, esText, now]);
  return { status:'ok', id };
}

function readBoardPosts_(payload) {
  getActiveSessionOrThrow_(payload.sessionToken);
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(BOARD_SHEET_NAME_);
  if (!sheet) return { status:'ok', posts:[] };
  var posts = getSheetRecords_(sheet);
  if (!posts.length) return { status:'ok', posts:[] };
  posts.sort(function (a,b) { return new Date(b.createdAt||0)-new Date(a.createdAt||0); });
  posts = posts.slice(0, BOARD_MAX_POSTS_);
  var commentsSheet = ss.getSheetByName(BOARD_COMMENTS_SHEET_NAME_);
  var allComments   = commentsSheet ? getSheetRecords_(commentsSheet) : [];
  return { status:'ok', posts: posts.map(function (post) {
    var comments = allComments.filter(function (c) { return c.postId === post.id; });
    comments.sort(function (a,b) { return new Date(a.createdAt||0)-new Date(b.createdAt||0); });
    return { id:trimAuthText_(post.id), userId:trimAuthText_(post.userId), username:trimAuthText_(post.username), company:trimAuthText_(post.company), question:trimAuthText_(post.question), esText:trimAuthText_(post.esText), createdAt:trimAuthText_(post.createdAt),
      comments: comments.map(function (c) { return { id:trimAuthText_(c.id), postId:trimAuthText_(c.postId), userId:trimAuthText_(c.userId), username:trimAuthText_(c.username), commentText:trimAuthText_(c.commentText), createdAt:trimAuthText_(c.createdAt) }; }) };
  }) };
}

function addBoardComment_(payload) {
  var session     = getActiveSessionOrThrow_(payload.sessionToken);
  var postId      = trimAuthText_(payload.postId);
  var commentText = trimAuthText_(payload.commentText);
  assertAuth_(postId,      'postIdが指定されていません。');
  assertAuth_(commentText, 'コメントを入力してください。');
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(BOARD_COMMENTS_SHEET_NAME_);
  if (!sheet) { sheet = ss.insertSheet(BOARD_COMMENTS_SHEET_NAME_); sheet.appendRow(['id','postId','userId','username','commentText','createdAt']); }
  var now     = new Date().toISOString();
  var id      = now.replace(/\D/g,'').slice(0,14)+'_'+Utilities.getUuid().replace(/-/g,'').slice(0,8);
  var users   = getSheetRecords_(getAuthSheet_(AUTH_SHEET_NAMES_.users));
  var userRec = users.find(function (r) { return r.id === session.userId; });
  sheet.appendRow([id, postId, session.userId, userRec ? trimAuthText_(userRec.username) : '', commentText, now]);
  return { status:'ok' };
}

// ── Claude API 呼び出し（AIツール用） ──────────────
function callClaude_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  assertAiUsageAllowed_(session.userId, 'callClaude');
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  assertAuth_(apiKey, 'AI機能が設定されていません。管理者にお問い合わせください。');

  var toolType = trimAuthText_(payload.toolType);
  var input    = payload.input || {};
  var systemPrompt, userMessage;

  if (toolType === 'esReview') {
    var esText = trimAuthText_(input.esText);
    assertAuth_(esText, 'ESの内容を入力してください。');
    systemPrompt = 'あなたは新卒就職活動のプロフェッショナルなESコーチです。慶應義塾大学の学生のESを丁寧に添削してください。フィードバックは日本語で、具体的かつ建設的に行ってください。\n\n【重要】回答の最初に、以下の形式でルーブリック評価スコア（各1〜5の整数）をJSON形式で出力してください。必ずこの形式で始めてください：\n[SCORES]{"structure":X,"specificity":X,"logic":X,"expression":X,"persuasion":X}[/SCORES]\n\n各スコアの基準：\n- structure（構成）: 1=構成が不明確 2=やや整理不足 3=基本的な構成はある 4=論理的で分かりやすい 5=完璧な構成\n- specificity（具体性）: 1=抽象的すぎる 2=具体性が不足 3=一部具体的 4=エピソードが具体的 5=非常に具体的で説得力大\n- logic（論理性）: 1=論理破綻 2=やや飛躍がある 3=概ね論理的 4=論理的で一貫性あり 5=完璧な論理展開\n- expression（表現力）: 1=表現が稚拙 2=やや単調 3=標準的 4=表現が豊か 5=非常に洗練された表現\n- persuasion（説得力）: 1=説得力なし 2=やや弱い 3=一定の説得力 4=説得力がある 5=非常に説得力が高い';
    var direction = trimAuthText_(input.direction || '');
    userMessage  = '【企業名】'+(trimAuthText_(input.company)||'（未記入）')+'\n【設問】'+(trimAuthText_(input.question)||'（未記入）')+'\n\n【ES内容】\n'+esText
      +(direction ? '\n\n【添削の方向性・重点指示】\n'+direction+'\n\n上記の指示を最優先で反映した添削を行ってください。' : '')
      +'\n\n最初に [SCORES]{"structure":X,"specificity":X,"logic":X,"expression":X,"persuasion":X}[/SCORES] の形式でルーブリックスコア（各1〜5）を出力し、その後に以下の観点で詳しく添削してください：\n1. 📌 構成・論理性（PREP法など）\n2. 💡 具体性・エピソードの説得力\n3. 🎯 企業・設問への適合性\n4. ✍️ 表現・文章力\n5. 🔧 改善提案（修正例を示す）\n\n最後に総合評価（S/A/B/C）と一言コメントをつけてください。';

  } else if (toolType === 'interviewCoach') {
    var answer = trimAuthText_(input.answer);
    assertAuth_(answer, '回答を入力してください。');
    systemPrompt = 'あなたは面接官の経験豊富な就活コーチです。慶應義塾大学の学生の面接回答を評価し、具体的な改善アドバイスをしてください。日本語で回答してください。';
    userMessage  = '【面接質問】'+(trimAuthText_(input.question)||'自己PR')+'\n【志望職種】'+(trimAuthText_(input.jobType)||'（未記入）')+'\n\n【学生の回答】\n'+answer+'\n\n以下の観点で評価してください：\n1. ⭐ 総合評価（10点満点）\n2. 📐 構成（STAR法など）\n3. 💪 強みの伝わり方\n4. 🏢 企業・職種への関連性\n5. 😊 熱意・志望度の伝わり方\n6. 🔧 改善アドバイス（具体的なフレーズ例を含む）\n\n最後に「このまま言えばOK」な改善版回答例を提示してください。';

  } else if (toolType === 'esRewrite') {
    var originalEs    = trimAuthText_(input.originalEs);
    var targetCompany = trimAuthText_(input.targetCompany);
    assertAuth_(originalEs,    '元のES内容を入力してください。');
    assertAuth_(targetCompany, '志望企業名を入力してください。');
    systemPrompt = 'あなたは就活のプロフェッショナルです。慶應義塾大学の学生のESを別企業向けに最適化してリライトしてください。日本語で回答してください。';
    userMessage  = '【元のES】\n'+originalEs+'\n\n【新しい志望企業】'+targetCompany+'\n【新しい設問】'+(trimAuthText_(input.targetQuestion)||'（元の設問に準じる）')+'\n\n元のESの強みを活かしながら、新しい企業・設問に最適化したESを作成してください。\n\n出力形式：\n1. 📝 リライト版ES（すぐ使える形で）\n2. 🔄 変更点とその理由\n3. ✅ 追加で強化できるポイント';

  } else if (toolType === 'analyzeSelfProfile') {
    var values       = trimAuthText_(input.values || '');
    var strengths    = trimAuthText_(input.strengths || '');
    var weaknesses   = trimAuthText_(input.weaknesses || '');
    var will         = trimAuthText_(input.will || '');
    var can          = trimAuthText_(input.can || '');
    var must         = trimAuthText_(input.must || '');
    var valuesReason = trimAuthText_(input.valuesReason || '');
    assertAuth_(values || strengths || will, '自己分析データを入力してください。');
    systemPrompt = 'あなたは就活のキャリアカウンセラーです。学生の自己分析データをもとに、強みのパターン、面接での活かし方、おすすめの業界を分析してください。日本語で、親しみやすく、具体的にアドバイスしてください。';
    userMessage  = '以下の自己分析データを総合的に分析してください。\n\n'
      + '【価値観】' + (values || '（未記入）') + '\n'
      + '【価値観の理由】' + (valuesReason || '（未記入）') + '\n'
      + '【強み】' + (strengths || '（未記入）') + '\n'
      + '【弱み・改善】' + (weaknesses || '（未記入）') + '\n'
      + '【Will（やりたいこと）】' + (will || '（未記入）') + '\n'
      + '【Can（できること）】' + (can || '（未記入）') + '\n'
      + '【Must（求められること）】' + (must || '（未記入）') + '\n\n'
      + '以下の4つの観点で分析してください：\n'
      + '1. 🔍 強みのパターン分析（価値観・強み・Canから見えるあなたの一貫した特徴）\n'
      + '2. 🎤 面接で使えるトークポイント（具体的なフレーズ例つき）\n'
      + '3. ⚠️ 気をつけたいポイント（弱みや盲点の対策）\n'
      + '4. 🏢 おすすめの業界・職種（理由つきで3〜5つ）\n\n'
      + 'それぞれ具体的に、すぐ使えるアドバイスをお願いします。';

  } else if (toolType === 'selfAnalysisChat') {
    var context  = trimAuthText_(input.context || '');
    var question = trimAuthText_(input.question || '');
    assertAuth_(question, '質問を入力してください。');
    systemPrompt = 'あなたは就活のキャリアカウンセラーです。学生の自己分析データをもとに、強みのパターン、面接での活かし方、おすすめの業界を分析してください。日本語で、親しみやすく、具体的にアドバイスしてください。以下は直前の分析結果です。この文脈を踏まえて学生の質問に答えてください。\n\n--- 直前の分析 ---\n' + context;
    userMessage  = question;

  } else if (toolType === 'buildMotivation') {
    var industry       = trimAuthText_(input.industry || '');
    var experience     = trimAuthText_(input.experience || '');
    var futureGoal     = trimAuthText_(input.futureGoal || '');
    var companyFeature = trimAuthText_(input.companyFeature || '');
    var charLimit      = parseInt(input.charLimit) || 400;
    assertAuth_(industry || companyFeature, '志望業界または企業の特徴を入力してください。');
    assertAuth_(experience, '自分の経験を入力してください。');
    systemPrompt = 'あなたは就活ESの専門家です。学生の経験・将来の目標・業界の特徴をもとに、説得力のある志望動機を生成してください。PREP法（結論→理由→具体例→結論）で構成し、指定文字数以内で書いてください。';
    userMessage  = '以下の情報をもとに、志望動機を生成してください。\n\n【志望業界】'+(industry||'（未記入）')+'\n【企業/業界の特徴・惹かれた点】'+(companyFeature||'（未記入）')+'\n【自分の経験（ガクチカ等）】\n'+experience+'\n【将来やりたいこと】'+(futureGoal||'（未記入）')+'\n【文字数制限】'+charLimit+'字以内\n\n以下の形式で出力してください：\n1. 📝 志望動機（'+charLimit+'字以内、PREP法で構成。そのまま提出できる形で）\n2. 💡 構成のポイント解説（なぜこの流れにしたか）\n3. ✅ さらに良くするためのヒント';

  } else if (toolType === 'deepDiveGakuchika') {
    var gkTitle = trimAuthText_(input.title || '');
    var gkStar  = input.star || {};
    assertAuth_(gkTitle, 'ガクチカのタイトルを入力してください。');
    var existingQs = trimAuthText_(input.existingQuestions || '');
    systemPrompt = 'あなたは大手企業の採用面接官（10年以上の経験）です。学生のガクチカに対して、実際の面接で聞くような鋭い深掘り質問を10問生成してください。\n\n質問の種類を以下のようにバランスよく含めてください：\n- 動機・背景を問う質問（2問）：「なぜそれを始めたのか」「きっかけは」\n- 行動の詳細を問う質問（3問）：「具体的にどう動いたのか」「他の選択肢はなかったのか」\n- 困難・失敗を問う質問（2問）：「一番苦労したことは」「失敗した経験は」\n- 学び・成長を問う質問（2問）：「何を学んだか」「今の自分にどう活きているか」\n- 価値観・人柄を問う質問（1問）：「チームでのあなたの役割は」「周囲からどう見られていたか」\n\n各質問は具体的で、「はい/いいえ」では答えられない開放型質問にしてください。';
    userMessage = '以下のガクチカに対する深掘り質問を10個生成してください。\n\n【タイトル】'+gkTitle+'\n'
      +(gkStar.situation ? '【状況(S)】'+gkStar.situation+'\n' : '')
      +(gkStar.task ? '【課題(T)】'+gkStar.task+'\n' : '')
      +(gkStar.action ? '【行動(A)】'+gkStar.action+'\n' : '')
      +(gkStar.result ? '【結果(R)】'+gkStar.result+'\n' : '')
      +(gkStar.appeal ? '【学び】'+gkStar.appeal+'\n' : '')
      +(existingQs ? '\n【既出の質問（重複を避けること）】\n'+existingQs+'\n' : '')
      +'\n各質問を1行1つ、番号付きで出力してください（例：1. なぜその活動を始めたのですか？）。質問以外の説明は不要です。';

  } else if (toolType === 'deepDiveChatGk') {
    var ddContext = trimAuthText_(input.context || '');
    var ddQuestion = trimAuthText_(input.question || '');
    assertAuth_(ddQuestion, '質問を入力してください。');
    systemPrompt = 'あなたは就活の面接官であり、キャリアカウンセラーです。学生のガクチカ深掘り回答を添削してください。\n\n以下の形式で添削してください：\n\n【スコア】X/5（1=不十分 2=やや不足 3=普通 4=良い 5=面接官を唸らせるレベル）\n\n【良い点】\n・具体的に何が良いか（1-2点）\n\n【改善点】\n・何が足りないか、どう改善すべきか（1-3点）\n\n【模範回答例】\nこの質問に対する理想的な回答例（200字程度）\n\n【面接官の視点】\nこの回答を聞いた面接官はどう感じるか（1-2文）\n\n上記の形式を必ず守って出力してください。\n\n【コンテキスト】\n'+ddContext;
    userMessage = ddQuestion;

  } else if (toolType === 'refineMotivation') {
    var motivationText = trimAuthText_(input.motivationText || '');
    assertAuth_(motivationText, 'ブラッシュアップする志望動機を入力してください。');
    var refineCharLimit = parseInt(input.charLimit) || 400;
    systemPrompt = 'あなたは就活ESの専門家です。学生の志望動機をさらにブラッシュアップしてください。日本語で回答してください。';
    userMessage  = '以下の志望動機をブラッシュアップしてください。\n\n【現在の志望動機】\n'+motivationText+'\n【文字数制限】'+refineCharLimit+'字以内\n\n以下の形式で出力してください：\n1. 📝 改善版志望動機（'+refineCharLimit+'字以内、そのまま提出できる形で）\n2. 🔄 変更点とその理由\n3. ✅ さらに改善できるポイント';

  } else if (toolType === 'deepDiveExperience') {
    var expTitle  = trimAuthText_(input.title || '');
    var expPeriod = trimAuthText_(input.period || '');
    var expDesc   = trimAuthText_(input.desc || '');
    assertAuth_(expTitle, '経験のタイトルを入力してください。');
    systemPrompt = 'あなたはキャリアカウンセラーです。学生の過去の経験について、その経験から見える価値観・強み・性格特性を引き出すための質問を8問生成してください。\n\n面接対策ではなく、自己理解が目的です。以下の観点でバランスよく：\n- 感情を問う（2問）：「その時どう感じた？」「一番嬉しかった/悔しかった瞬間は？」\n- 選択の理由を問う（2問）：「なぜそれを選んだ？」「他の選択肢はあった？」\n- 価値観を問う（2問）：「何が大事だと思った？」「譲れなかったことは？」\n- パターンを問う（2問）：「他にも似た経験がある？」「この経験と今の自分のつながりは？」\n\n質問は温かみのあるトーンで、内省を促すものにしてください。';
    userMessage = '以下の経験について、自己理解を深めるための振り返り質問を8個生成してください。\n\n【時期】' + (expPeriod || '（未記入）') + '\n【タイトル】' + expTitle + '\n【経験の概要】' + (expDesc || '（未記入）') + '\n\n各質問を1行1つ、番号付きで出力してください（例：1. その経験を始めたきっかけは何でしたか？）。質問以外の説明は不要です。';

  } else if (toolType === 'analyzeExperiences') {
    var expData = trimAuthText_(input.experiencesText || '');
    assertAuth_(expData, '経験データを入力してください。');
    systemPrompt = 'あなたはキャリアカウンセラーです。学生の複数の経験とその振り返りから、以下を分析してください：\n\n1. 【価値観パターン】複数の経験に共通する価値観は何か（2-3個）\n2. 【強みパターン】繰り返し発揮されている強みは何か（2-3個）\n3. 【行動パターン】困難に直面した時の典型的な対処法\n4. 【向いている環境】どんな環境で力を発揮しやすいか\n5. 【自己分析WBへの提案】価値観・強み・Will-Can-Mustに何を書くべきか\n\n具体的な経験の引用を交えて、説得力のある分析をしてください。\n\n【重要】回答の最後に、以下の形式で提案データをJSON形式で出力してください：\n[SUGGESTIONS]{"values":["価値観1","価値観2"],"strengths":"強み1\\n強み2\\n強み3","will":"Willの提案文"}[/SUGGESTIONS]';
    userMessage = '以下の複数の経験とその振り返りから、パターンを分析してください。\n\n' + expData;

  } else {
    throw new Error('不明なツールタイプです。');
  }

  var res  = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model:'claude-opus-4-6', max_tokens:4096, system:systemPrompt, messages:[{ role:'user', content:userMessage }] }),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    var eb = {}; try { eb = JSON.parse(res.getContentText()); } catch(e) {}
    throw new Error('AI処理に失敗しました: '+((eb.error&&eb.error.message)||'HTTP '+res.getResponseCode()));
  }

  var data = JSON.parse(res.getContentText());
  var resultText = '';
  (data.content||[]).forEach(function (b) { if (b.type==='text') resultText += b.text; });
  return { status:'ok', result: resultText };
}

// ── 認証ユーティリティ ─────────────────────────────
function getActiveSessionOrThrow_(sessionToken) {
  var token = trimAuthText_(sessionToken);
  assertAuth_(token, 'ログイン情報が見つかりません。');
  var sessionsSheet = getAuthSheet_(AUTH_SHEET_NAMES_.sessions);
  var sessions = getSheetRecords_(sessionsSheet);
  var now      = new Date();
  var index    = findRecordIndex_(sessions, function (r) { return r.sessionToken === token && String(r.active) !== '0'; });
  assertAuth_(index >= 0, 'ログイン情報が見つかりません。');
  var session  = sessions[index];
  var expiresAt = new Date(session.expiresAt);
  assertAuth_(!isNaN(expiresAt.getTime()) && expiresAt.getTime() > now.getTime(), 'ログインの有効期限が切れました。');
  sessionsSheet.getRange(index+2, 5).setValue(now.toISOString());
  return session;
}

function getConfiguredAdminValues_(key) {
  var raw = trimAuthText_(PropertiesService.getScriptProperties().getProperty(key));
  if (!raw) return [];
  return raw.split(/[\s,;]+/).map(function (value) {
    return trimAuthText_(value);
  }).filter(Boolean);
}

function getAuthUserById_(userId) {
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  return users.find(function (row) {
    return row.id === userId;
  }) || null;
}

function isAdminUserRecord_(user) {
  if (!user) return false;

  var adminIds = getConfiguredAdminValues_('AUTH_ADMIN_USER_IDS');
  if (adminIds.indexOf(trimAuthText_(user.id)) >= 0) return true;

  var adminEmails = getConfiguredAdminValues_('AUTH_ADMIN_EMAILS').map(function (value) {
    return normalizeAuthKey_(value);
  });
  if (!adminEmails.length) return false;

  var email = trimAuthText_(user.email || user.username || user.emailKey || user.usernameKey);
  return adminEmails.indexOf(normalizeAuthKey_(email)) >= 0;
}

function requireAdminSession_(payload) {
  var session = getActiveSessionOrThrow_(payload && payload.sessionToken);
  var user = getAuthUserById_(session.userId);
  assertAuth_(user, 'アカウントが見つかりません。');

  var adminIds = getConfiguredAdminValues_('AUTH_ADMIN_USER_IDS');
  var adminEmails = getConfiguredAdminValues_('AUTH_ADMIN_EMAILS');
  assertAuth_(adminIds.length || adminEmails.length, '管理者が未設定です。AUTH_ADMIN_EMAILS または AUTH_ADMIN_USER_IDS を設定してください。');
  assertAuth_(isAdminUserRecord_(user), '管理者権限が必要です。');
  return { session: session, user: user };
}

function createSessionForUser_(userId) {
  var sessionsSheet = getAuthSheet_(AUTH_SHEET_NAMES_.sessions);
  var token     = Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
  var now       = new Date();
  var expiresAt = new Date(now.getTime()+AUTH_SESSION_TTL_DAYS_*24*60*60*1000);
  sessionsSheet.appendRow([token, userId, now.toISOString(), expiresAt.toISOString(), now.toISOString(), '1']);
  return token;
}

function invalidateSessionsForUser_(userId, keepSessionToken) {
  var sessionsSheet = getAuthSheet_(AUTH_SHEET_NAMES_.sessions);
  var sessions = getSheetRecords_(sessionsSheet);
  sessions.forEach(function (row, index) {
    if (row.userId !== userId) return;
    if (keepSessionToken && row.sessionToken === keepSessionToken) return;
    if (String(row.active) === '0') return;
    sessionsSheet.getRange(index + 2, 6).setValue('0');
  });
}

function getLikedCompaniesForUser_(userId) {
  var sheet = getAuthSheet_(AUTH_SHEET_NAMES_.liked);
  var rows  = getSheetRecords_(sheet);
  var row   = rows.find(function (r) { return r.userId === userId; });
  if (!row || !row.likedCompaniesJson) return [];
  try { return JSON.parse(row.likedCompaniesJson); } catch (e) { return []; }
}

function saveLikedCompaniesForUser_(userId, likedCompanies) {
  var sheet    = getAuthSheet_(AUTH_SHEET_NAMES_.liked);
  var rows     = getSheetRecords_(sheet);
  var rowIndex = findRecordIndex_(rows, function (r) { return r.userId === userId; });
  var payload  = JSON.stringify(likedCompanies||[]);
  var updatedAt = new Date().toISOString();
  if (rowIndex >= 0) { sheet.getRange(rowIndex+2, 2, 1, 2).setValues([[payload, updatedAt]]); return; }
  sheet.appendRow([userId, payload, updatedAt]);
}

function sanitizeUserRecord_(record) {
  var preferredCompanies = getPreferredCompanies_(record);
  var lineQrUrl = trimAuthText_(record.lineQrDriveUrl||record.lineQrUrl);
  return { id:trimAuthText_(record.id), username:trimAuthText_(record.username), usernameKey:trimAuthText_(record.usernameKey), desiredIndustry:trimAuthText_(record.desiredIndustry), preferredCompanies, preferredCompany1:preferredCompanies[0]||'', preferredCompany2:preferredCompanies[1]||'', preferredCompany3:preferredCompanies[2]||'', lineName:trimAuthText_(record.lineName), lineQrUrl, hasLineQr:!!lineQrUrl, createdAt:trimAuthText_(record.createdAt), updatedAt:trimAuthText_(record.updatedAt) };
}

function ensureAuthSheets_() {
  ensureSheet_(AUTH_SHEET_NAMES_.users, AUTH_USER_HEADERS_);
  ensureSheet_(AUTH_SHEET_NAMES_.sessions, ['sessionToken','userId','createdAt','expiresAt','lastSeenAt','active']);
  ensureSheet_(AUTH_SHEET_NAMES_.liked,    ['userId','likedCompaniesJson','updatedAt']);
}

function ensureSheet_(name, headers) {
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) { sheet.appendRow(headers); return; }
  var existing = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  headers.forEach(function (h) { if (existing.indexOf(h) === -1) { sheet.getRange(1,sheet.getLastColumn()+1).setValue(h); existing.push(h); } });
}

function getAuthSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('AUTH_SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getAuthSheet_(name) { return getAuthSpreadsheet_().getSheetByName(name); }

function getSheetRecords_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return [];
  var headers = values[0];
  return values.slice(1).map(function (row) {
    var rec = {};
    headers.forEach(function (h, i) { rec[h] = row[i]; });
    return rec;
  });
}

function findRecordIndex_(rows, predicate) {
  for (var i = 0; i < rows.length; i++) { if (predicate(rows[i])) return i; }
  return -1;
}

function getPreferredCompanies_(payload) {
  var raw = Array.isArray(payload.preferredCompanies) ? payload.preferredCompanies : [payload.preferredCompany1, payload.preferredCompany2, payload.preferredCompany3];
  var unique = [];
  raw.forEach(function (item) {
    var v = trimAuthText_(item); if (!v) return;
    if (!unique.some(function (u) { return normalizeAuthKey_(u) === normalizeAuthKey_(v); })) unique.push(v);
  });
  return unique.slice(0,3);
}

function saveLineQrAsset_(userId, lineQrDataUrl, lineQrFileName, existingFileId) {
  var dataUrl = trimAuthText_(lineQrDataUrl);
  var match   = dataUrl.match(/^data:(image\/[^;]+);base64,(.+)$/);
  assertAuth_(match, 'LINE QRは画像ファイルをアップロードしてください。');
  var mimeType  = match[1];
  var bytes     = Utilities.base64Decode(match[2]);
  var extension = ({ 'image/jpeg':'.jpg','image/png':'.png','image/gif':'.gif','image/webp':'.webp' })[mimeType] || '.img';
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'Asia/Tokyo', 'yyyyMMdd-HHmmss');
  var baseName  = trimAuthText_(lineQrFileName).replace(/[^\w.\-]+/g,'_');
  var fileName  = 'line-qr-'+userId+'-'+timestamp+(baseName ? '-'+baseName.replace(/\.[^.]+$/,'') : '')+extension;
  var blob      = Utilities.newBlob(bytes, mimeType, fileName);
  var folderId  = PropertiesService.getScriptProperties().getProperty('AUTH_LINE_QR_FOLDER_ID');
  var folder    = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var file      = folder.createFile(blob);
  if (existingFileId) { try { DriveApp.getFileById(existingFileId).setTrashed(true); } catch(e){} }
  return { fileId: file.getId(), url: file.getUrl() };
}

function normalizeAuthKey_(value) { return trimAuthText_(value).normalize('NFKC').toLowerCase(); }
function trimAuthText_(value) { return String(value||'').trim(); }
function assertAuth_(condition, message) { if (!condition) throw new Error(message); }

function sha256Hex_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { var n = b < 0 ? b+256 : b; return ('0'+n.toString(16)).slice(-2); }).join('');
}

var PASSWORD_HASH_SCHEME_ = 'iter-sha256-v1';
var PASSWORD_HASH_ITERATIONS_ = 120000;

function hashPasswordWithIterations_(password, salt, iterations) {
  var value = String(salt || '') + ':' + String(password || '');
  var total = Math.max(parseInt(iterations, 10) || 0, 1);
  for (var i = 0; i < total; i++) {
    value = sha256Hex_(value);
  }
  return value;
}

function buildPasswordHashRecord_(password) {
  var salt = Utilities.getUuid().replace(/-/g, '');
  return {
    salt: salt,
    passwordHash: PASSWORD_HASH_SCHEME_ + '$' + PASSWORD_HASH_ITERATIONS_ + '$' + hashPasswordWithIterations_(password, salt, PASSWORD_HASH_ITERATIONS_)
  };
}

function verifyPasswordRecord_(password, user) {
  var storedHash = trimAuthText_(user && user.passwordHash);
  var salt = trimAuthText_(user && user.salt);
  if (storedHash.indexOf(PASSWORD_HASH_SCHEME_ + '$') === 0) {
    var parts = storedHash.split('$');
    if (parts.length === 3) {
      return hashPasswordWithIterations_(password, salt, parseInt(parts[1], 10)) === parts[2];
    }
  }
  return sha256Hex_(salt + ':' + password) === storedHash;
}

function maybeUpgradePasswordHash_(usersSheet, rowIndex, user, password) {
  var storedHash = trimAuthText_(user && user.passwordHash);
  if (storedHash.indexOf(PASSWORD_HASH_SCHEME_ + '$') === 0) return;
  if (rowIndex < 0) return;

  var nextHash = buildPasswordHashRecord_(password);
  usersSheet.getRange(rowIndex + 2, 5, 1, 2).setValues([[nextHash.passwordHash, nextHash.salt]]);
  user.passwordHash = nextHash.passwordHash;
  user.salt = nextHash.salt;
}

// ─────────────────────────────────────────────
//  選考情報自動収集（企業公式HPからAI抽出）
// ─────────────────────────────────────────────

var BLOCKED_DOMAINS_ = [
  // 就活情報サイト
  'onecareer.jp', 'unistyleinc.com', 'mynavi.jp', 'rikunabi.com',
  'shukatsu-kaigi.jp', 'goodfind.jp', 'gaishishukatsu.com',
  'career-tasu.jp', 'en-courage.com', 'offerbox.jp', 'dodacampus.jp',
  'type.jp', 'doda.jp', 'jobway.jp', 'careerpark.jp',
  // 口コミ・評判サイト
  'openwork.jp', 'vorkers.com', 'en-hyouban.com', 'jobtalk.jp',
  'glassdoor.com', 'tenshoku-kaigi.jp', 'kaisha-hyouban.com',
  // まとめ・ニュースサイト
  'wikipedia.org', 'nikkei.com', 'toyokeizai.net', 'diamond.jp',
  'president.jp', 'newspicks.com',
  // SNS・掲示板
  'twitter.com', 'x.com', '5ch.net', '2ch.sc', 'note.com', 'qiita.com'
];

function extractUrlHostname_(url) {
  var raw = trimText_(url);
  var match = raw.match(/^https?:\/\/([^\/?#]+)/i);
  if (!match) return '';
  var authority = String(match[1] || '');
  if (authority.indexOf('@') !== -1) return '';
  var host = authority.replace(/:\d+$/, '').toLowerCase();
  if (host.charAt(0) === '[' && host.charAt(host.length - 1) === ']') {
    host = host.slice(1, -1);
  }
  return host;
}

function isPrivateHost_(hostname) {
  var host = String(hostname || '').toLowerCase();
  if (!host) return true;
  if (host === 'localhost' || host === '0.0.0.0' || host === '::1' || host === '::') return true;
  if (host.endsWith('.localhost')) return true;
  if (/^\d+\.\d+\.\d+\.\d+$/.test(host)) {
    var parts = host.split('.').map(function (part) { return parseInt(part, 10) || 0; });
    if (parts[0] === 10 || parts[0] === 127 || parts[0] === 0) return true;
    if (parts[0] === 169 && parts[1] === 254) return true;
    if (parts[0] === 172 && parts[1] >= 16 && parts[1] <= 31) return true;
    if (parts[0] === 192 && parts[1] === 168) return true;
    if (parts[0] === 100 && parts[1] >= 64 && parts[1] <= 127) return true;
    if (parts[0] === 198 && (parts[1] === 18 || parts[1] === 19)) return true;
  }
  if (host.indexOf(':') !== -1) {
    if (host.indexOf('fe80:') === 0 || host.indexOf('fc') === 0 || host.indexOf('fd') === 0) return true;
  }
  return false;
}

function normalizeSafeFetchUrl_(url) {
  var raw = trimText_(url);
  if (!raw || !/^https?:\/\//i.test(raw)) return '';
  var host = extractUrlHostname_(raw);
  if (!host || isPrivateHost_(host)) return '';
  return raw;
}

function isDomainBlocked_(url) {
  try {
    var hostname = extractUrlHostname_(url);
    if (!hostname) return true;
    return BLOCKED_DOMAINS_.some(function(d) {
      return hostname === d || hostname.endsWith('.' + d);
    });
  } catch (e) {
    return false;
  }
}

function handleFetchRecruitInfo_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  assertAiUsageAllowed_(session.userId, 'fetchRecruitInfo');
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  assertAuth_(apiKey, 'AI機能が設定されていません。');

  var companyName = trimText_(payload.companyName);
  assertAuth_(companyName, '企業名を入力してください。');
  var manualUrl = trimText_(payload.url || '');

  // Step 1: URLが指定されていなければClaudeに推測させる
  var recruitUrl = manualUrl;
  if (!recruitUrl) {
    var urlGuessResp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 300,
        system: 'あなたは日本企業の採用ページURLに詳しいアシスタントです。企業名を受け取り、その企業の公式採用/リクルートページのURLを1つだけ返してください。URLのみを返し、説明は不要です。わからない場合は「不明」と返してください。',
        messages: [{ role: 'user', content: companyName + ' の公式採用ページURL' }]
      }),
      muteHttpExceptions: true
    });
    if (urlGuessResp.getResponseCode() === 200) {
      var urlData = JSON.parse(urlGuessResp.getContentText());
      var guessedUrl = '';
      (urlData.content || []).forEach(function(b) { if (b.type === 'text') guessedUrl += b.text; });
      guessedUrl = guessedUrl.trim();
      if (guessedUrl && guessedUrl !== '不明' && guessedUrl.match(/^https?:\/\//)) {
        recruitUrl = guessedUrl;
      }
    }
  }

  if (!recruitUrl) {
    return { status: 'error', message: companyName + ' の公式採用ページが見つかりませんでした。URLを手動で入力してください。' };
  }

  recruitUrl = normalizeSafeFetchUrl_(recruitUrl);
  if (!recruitUrl) {
    return { status: 'error', message: '安全ではないURLは取得できません。企業の公式採用ページURLを入力してください。' };
  }

  // Step 2: ドメインチェック
  if (isDomainBlocked_(recruitUrl)) {
    return { status: 'error', message: 'このURLは就活情報サイトのため取得できません。企業の公式採用ページのURLを入力してください。' };
  }

  // Step 3: ページ取得
  var pageHtml = '';
  try {
    var fetchResp = UrlFetchApp.fetch(recruitUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; KeioNaviBot/1.0)' }
    });
    if (fetchResp.getResponseCode() !== 200) {
      return { status: 'error', message: 'ページの取得に失敗しました（HTTP ' + fetchResp.getResponseCode() + '）。URLを確認してください。', suggestedUrl: recruitUrl };
    }
    pageHtml = fetchResp.getContentText();
  } catch (e) {
    return { status: 'error', message: 'ページの取得に失敗しました: ' + e.message + '。URLを確認してください。', suggestedUrl: recruitUrl };
  }

  // HTMLを短縮（GASのメモリ制限対策）
  pageHtml = pageHtml.replace(/<script[\s\S]*?<\/script>/gi, '')
                     .replace(/<style[\s\S]*?<\/style>/gi, '')
                     .replace(/<[^>]+>/g, ' ')
                     .replace(/\s+/g, ' ')
                     .trim();
  if (pageHtml.length > 15000) pageHtml = pageHtml.substring(0, 15000);

  // Step 4: Claude APIで選考情報を抽出
  var extractResp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1500,
      system: '企業の採用ページから選考情報を抽出するアシスタントです。必ず以下のJSON形式で返してください。情報が見つからない項目は空文字にしてください。',
      messages: [{ role: 'user', content: '以下は「' + companyName + '」の公式採用ページのテキスト内容です。選考情報を抽出してJSON形式で返してください。\n\n' + pageHtml + '\n\n出力形式（必ずこのJSON形式で）:\n```json\n{\n  "company": "企業名",\n  "deadline": "エントリー締切日（YYYY/MM/DD形式、不明なら空文字）",\n  "type": "本選考 or インターン",\n  "flow": "選考フロー（ES→適性検査→1次面接→... の形式）",\n  "requirements": "応募条件・資格",\n  "url": "' + recruitUrl + '",\n  "industry": "業界",\n  "notes": "その他特記事項"\n}\n```' }]
    }),
    muteHttpExceptions: true
  });

  if (extractResp.getResponseCode() !== 200) {
    return { status: 'error', message: 'AI解析に失敗しました。' };
  }

  var extractData = JSON.parse(extractResp.getContentText());
  var resultText = '';
  (extractData.content || []).forEach(function(b) { if (b.type === 'text') resultText += b.text; });

  // JSONを抽出
  var jsonMatch = resultText.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    return { status: 'error', message: 'AI解析の結果を解析できませんでした。' };
  }

  var info = {};
  try { info = JSON.parse(jsonMatch[0]); } catch (e) {
    return { status: 'error', message: 'AI解析結果のJSON解析に失敗しました。' };
  }

  return {
    status: 'ok',
    recruitInfo: {
      company: trimText_(info.company) || companyName,
      deadline: trimText_(info.deadline),
      type: trimText_(info.type),
      flow: trimText_(info.flow),
      requirements: trimText_(info.requirements),
      url: trimText_(info.url) || recruitUrl,
      industry: trimText_(info.industry),
      notes: trimText_(info.notes)
    },
    sourceUrl: recruitUrl
  };
}

function handleSaveRecruitInfo_(payload) {
  getActiveSessionOrThrow_(payload.sessionToken);
  var info = payload.recruitInfo || {};
  var company = trimText_(info.company);
  assertAuth_(company, '企業名が必要です。');

  var config = SHEET_KEYS['selection'];
  var sheet = getSiteSheet_(config.sheetName, true, config.spreadsheetId);

  // 選考情報スプシに追記（A:企業名 B:締切日 C:種別 D:リンク E:業界 F:特集 G:一般呼称）
  sheet.appendRow([
    company,
    trimText_(info.deadline),
    trimText_(info.type),
    trimText_(info.url),
    trimText_(info.industry),
    '',
    ''
  ]);

  return { status: 'ok', message: company + ' の選考情報をスプシに保存しました。' };
}

// ─────────────────────────────────────────────
//  アクティビティログ / ファネル分析
// ─────────────────────────────────────────────

var ACTIVITY_LOG_SHEET_NAME_ = 'activity_log';
var ACTIVITY_LOG_HEADERS_    = ['timestamp', 'userId', 'event', 'page', 'feature', 'userAgent'];

/**
 * activity_log シートを取得（なければ自動作成）
 */
function getActivityLogSheet_() {
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(ACTIVITY_LOG_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(ACTIVITY_LOG_SHEET_NAME_);
    sheet.appendRow(ACTIVITY_LOG_HEADERS_);
  }
  return sheet;
}

/**
 * action: logActivity — アクティビティをシートに記録
 * セッショントークンがあればユーザーIDを解決、なければ anonymous
 */
function handleLogActivity_(payload) {
  var event   = trimText_(payload.event   || '');
  var page    = trimText_(payload.page    || '');
  var feature = trimText_(payload.feature || '');
  var ua      = trimText_(payload.userAgent || '');

  if (!event) return { status: 'error', message: 'event は必須です。' };

  // セッショントークンからユーザーIDを取得（任意・失敗しても記録する）
  var userId = 'anonymous';
  var sessionToken = trimText_(payload.sessionToken || '');
  if (sessionToken) {
    try {
      var session = getActiveSessionOrThrow_(sessionToken);
      userId = session.userId || 'anonymous';
    } catch (e) {
      // セッション切れ等でもログは記録する
    }
  }

  var sheet = getActivityLogSheet_();
  var now   = new Date().toISOString();
  sheet.appendRow([now, userId, event, page, feature, ua]);

  return { status: 'ok' };
}

/**
 * action: getActivitySummary — 管理者向け集計データを返す
 * 直近7日間のページビュー、機能利用ランキング、ファネル集計
 */
function handleGetActivitySummary_(payload) {
  requireAdminSession_(payload);

  var sheet = getActivityLogSheet_();
  var rows  = getSheetRecords_(sheet);

  var now       = new Date();
  var sevenDaysAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  // 直近7日のログのみ抽出
  var recentRows = rows.filter(function (r) {
    var ts = new Date(r.timestamp);
    return !isNaN(ts.getTime()) && ts.getTime() >= sevenDaysAgo.getTime();
  });

  // (1) 直近7日のページビュー合計
  var pageViews7d = recentRows.filter(function (r) { return r.event === 'page_view'; }).length;

  // (2) 機能利用ランキング（feature別の集計）
  var featureCounts = {};
  recentRows.forEach(function (r) {
    if (r.event === 'feature_use' && r.feature) {
      featureCounts[r.feature] = (featureCounts[r.feature] || 0) + 1;
    }
  });
  var topFeatures = Object.keys(featureCounts).map(function (k) {
    return { feature: k, count: featureCounts[k] };
  });
  topFeatures.sort(function (a, b) { return b.count - a.count; });
  topFeatures = topFeatures.slice(0, 10);

  // (3) ファネル分析: ユニークユーザー数を各段階で計算
  //     全ログ（期間制限なし）を使ってファネルを計算
  var funnelUsers = {
    registered: {},       // auth_users のユーザー数（別途計算）
    firstGakuchika: {},   // feature に 'gakuchika' を含むイベントがあるユーザー
    firstEs: {},          // feature に 'es_' を含むイベントがあるユーザー
    firstInterview: {}    // feature に 'interview' を含むイベントがあるユーザー
  };

  rows.forEach(function (r) {
    var uid = trimText_(r.userId);
    if (!uid || uid === 'anonymous') return;
    var feat = trimText_(r.feature).toLowerCase();
    if (feat.indexOf('gakuchika') >= 0) funnelUsers.firstGakuchika[uid] = true;
    if (feat.indexOf('es_') >= 0 || feat === 'es_save' || feat === 'es_review' || feat === 'es_rewrite') funnelUsers.firstEs[uid] = true;
    if (feat.indexOf('interview') >= 0) funnelUsers.firstInterview[uid] = true;
  });

  // 登録ユーザー数は auth_users シートから取得
  var registeredCount = 0;
  try {
    var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
    if (usersSheet) {
      var lastRow = usersSheet.getLastRow();
      registeredCount = lastRow > 1 ? lastRow - 1 : 0;
    }
  } catch (e) {}

  return {
    status: 'ok',
    summary: {
      pageViews7d: pageViews7d,
      topFeatures: topFeatures,
      funnel: {
        registered:     registeredCount,
        firstGakuchika: Object.keys(funnelUsers.firstGakuchika).length,
        firstEs:        Object.keys(funnelUsers.firstEs).length,
        firstInterview: Object.keys(funnelUsers.firstInterview).length
      }
    }
  };
}

// ══════════════════════════════════════════════════════════════════════
//  他己分析（ピアフィードバック）
// ══════════════════════════════════════════════════════════════════════

var PEER_FEEDBACK_SHEET_NAME_ = 'peer_feedback';

function ensurePeerFeedbackSheet_() {
  var ss = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(PEER_FEEDBACK_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(PEER_FEEDBACK_SHEET_NAME_);
    sheet.appendRow(['token', 'userId', 'displayName', 'createdAt', 'responses']);
  }
  return sheet;
}

/**
 * createPeerFeedbackRequest — ログインユーザーがトークンを生成
 */
function createPeerFeedbackRequest_(payload) {
  ensureAuthSheets_();
  var session = getActiveSessionOrThrow_(payload.sessionToken);

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var user  = users.find(function (r) { return r.id === session.userId; });
  assertAuth_(user, 'アカウントが見つかりません。');

  var sheet = ensurePeerFeedbackSheet_();
  var token = Utilities.getUuid().replace(/-/g, '');
  var displayName = trimAuthText_(user.username) || 'ユーザー';
  var now = new Date().toISOString();

  sheet.appendRow([token, session.userId, displayName, now, '[]']);

  return { status: 'ok', token: token };
}

/**
 * submitPeerFeedback — 友人がトークン付きで回答を送信（認証不要）
 */
function submitPeerFeedback_(payload) {
  var token = trimText_(payload.token);
  if (!token) return { status: 'error', message: 'トークンが見つかりません。' };

  var responses = payload.responses;
  if (!responses || typeof responses !== 'object') {
    return { status: 'error', message: '回答データが不正です。' };
  }

  var sheet = ensurePeerFeedbackSheet_();
  var records = getSheetRecords_(sheet);
  var rowIndex = -1;
  for (var i = 0; i < records.length; i++) {
    if (String(records[i].token) === token) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex < 0) {
    return { status: 'error', message: 'このリンクは無効です。' };
  }

  // Parse existing responses and append
  var existing = [];
  try {
    existing = JSON.parse(records[rowIndex].responses || '[]');
  } catch(e) { existing = []; }
  if (!Array.isArray(existing)) existing = [];

  existing.push(responses);

  // Update the cell (responses is column 5, row = rowIndex + 2 for header offset)
  sheet.getRange(rowIndex + 2, 5).setValue(JSON.stringify(existing));

  return { status: 'ok' };
}

/**
 * getPeerFeedbackResults — ログインユーザーの全フィードバックを返す
 */
function getPeerFeedbackResults_(payload) {
  ensureAuthSheets_();
  var session = getActiveSessionOrThrow_(payload.sessionToken);

  var sheet = ensurePeerFeedbackSheet_();
  var records = getSheetRecords_(sheet);

  var allFeedbacks = [];
  records.forEach(function(r) {
    if (String(r.userId) === session.userId) {
      try {
        var responses = JSON.parse(r.responses || '[]');
        if (Array.isArray(responses)) {
          allFeedbacks = allFeedbacks.concat(responses);
        }
      } catch(e) {}
    }
  });

  return { status: 'ok', feedbacks: allFeedbacks };
}

/**
 * getPeerFeedbackInfo — トークンからdisplayNameを返す（認証不要）
 */
function getPeerFeedbackInfo_(payload) {
  var token = trimText_(payload.token);
  if (!token) return { status: 'error', message: 'トークンが見つかりません。' };

  var sheet = ensurePeerFeedbackSheet_();
  var records = getSheetRecords_(sheet);

  for (var i = 0; i < records.length; i++) {
    if (String(records[i].token) === token) {
      return { status: 'ok', displayName: trimText_(records[i].displayName) || 'ユーザー' };
    }
  }

  return { status: 'error', message: 'このリンクは無効です。' };
}
