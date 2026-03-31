/*
 * Add this file to the existing Google Apps Script project behind GAS_PROXY_URL.
 *
 * In your current doPost(e), parse the JSON body once and insert:
 *
 *   var payload = JSON.parse(e.postData.contents || '{}');
 *   var authResponse = handleAuthAction_(payload);
 *   if (authResponse) return authResponse;
 *
 * Then continue with the existing read / generate / save logic.
 */

var AUTH_SHEET_NAMES_ = {
  users: 'auth_users',
  sessions: 'auth_sessions',
  liked: 'auth_liked'
};

var AUTH_USER_HEADERS_ = [
  'id',
  'email',
  'emailKey',
  'desiredIndustry',
  'passwordHash',
  'salt',
  'createdAt',
  'updatedAt',
  'preferredCompany1',
  'preferredCompany2',
  'preferredCompany3',
  'lineName',
  'lineQrDriveFileId',
  'lineQrDriveUrl',
  'displayName',
  'referralCode'
];

var AUTH_SESSION_TTL_DAYS_ = 30;

function handleAuthAction_(payload) {
  if (!payload || typeof payload !== 'object' || !String(payload.action || '').match(/^(auth|writeBoardPost|readBoardPosts|addBoardComment|callClaude|writeES|readMyES|deleteES|writeGakuchika|readMyGakuchika|deleteGakuchika|readMyIPData|writeIPCompany|deleteIPCompany|writeIPQuestion|deleteIPQuestion|replaceIPGakuchikaQuestions|readMyProgress|writeProgress|deleteProgress|searchMembers|getGroups|joinGroup|leaveGroup|createGroup|getGroupMembers|updateMatchingPrefs|sendWeeklyDigest|readQuestions|postQuestion|postReply|deleteQuestion|postExperience|readExperiences|readGdFeedback|writeGdFeedback|deleteGdFeedback|postPracticeRequest|readPracticeRequests|joinPracticeRequest|leavePracticeRequest|closePracticeRequest|deletePracticeRequest)/)) {
    return null;
  }

  try {
    switch (payload.action) {
      case 'authRegister':
        return jsonResponse_(authRegister_(payload));
      case 'authLogin':
        return jsonResponse_(authLogin_(payload));
      case 'authLogout':
        return jsonResponse_(authLogout_(payload));
      case 'authUpdateProfile':
        return jsonResponse_(authUpdateProfile_(payload));
      case 'authChangePassword':
        return jsonResponse_(authChangePassword_(payload));
      case 'authSetLikedCompanies':
        return jsonResponse_(authSetLikedCompanies_(payload));
      case 'authDeleteAccount':
        return jsonResponse_(authDeleteAccount_(payload));
      case 'authRequestPasswordReset':
        return jsonResponse_(authRequestPasswordReset_(payload));
      case 'authResetPassword':
        return jsonResponse_(authResetPassword_(payload));
      case 'authGetReferralInfo':
        return jsonResponse_(authGetReferralInfo_(payload));
      case 'readMyProgress':  return jsonResponse_(readMyProgress_(payload));
      case 'writeProgress':   return jsonResponse_(writeProgress_(payload));
      case 'deleteProgress':  return jsonResponse_(deleteProgress_(payload));
      case 'searchMembers':       return jsonResponse_(searchMembers_(payload));
      case 'getGroups':           return jsonResponse_(getGroups_(payload));
      case 'joinGroup':           return jsonResponse_(joinGroup_(payload));
      case 'leaveGroup':          return jsonResponse_(leaveGroup_(payload));
      case 'createGroup':         return jsonResponse_(createGroup_(payload));
      case 'getGroupMembers':     return jsonResponse_(getGroupMembers_(payload));
      case 'updateMatchingPrefs': return jsonResponse_(updateMatchingPrefs_(payload));
      case 'writeBoardPost': return jsonResponse_(writeBoardPost_(payload));
      case 'readBoardPosts': return jsonResponse_(readBoardPosts_(payload));
      case 'addBoardComment': return jsonResponse_(addBoardComment_(payload));
      case 'callClaude': return jsonResponse_(callClaude_(payload));
      case 'writeES':    return jsonResponse_(writeES_(payload));
      case 'readMyES':   return jsonResponse_(readMyES_(payload));
      case 'deleteES':   return jsonResponse_(deleteES_(payload));
      case 'writeGakuchika':    return jsonResponse_(writeGakuchika_(payload));
      case 'readMyGakuchika':   return jsonResponse_(readMyGakuchika_(payload));
      case 'deleteGakuchika':   return jsonResponse_(deleteGakuchika_(payload));
      case 'readMyIPData':               return jsonResponse_(readMyIPData_(payload));
      case 'writeIPCompany':             return jsonResponse_(writeIPCompany_(payload));
      case 'deleteIPCompany':            return jsonResponse_(deleteIPCompany_(payload));
      case 'writeIPQuestion':            return jsonResponse_(writeIPQuestion_(payload));
      case 'deleteIPQuestion':           return jsonResponse_(deleteIPQuestion_(payload));
      case 'replaceIPGakuchikaQuestions': return jsonResponse_(replaceIPGakuchikaQuestions_(payload));
      case 'sendWeeklyDigest':   return jsonResponse_(sendWeeklyDigest_(payload));
      case 'readQuestions':      return jsonResponse_(readQuestions_(payload));
      case 'postQuestion':       return jsonResponse_(postQuestion_(payload));
      case 'postReply':          return jsonResponse_(postReply_(payload));
      case 'deleteQuestion':     return jsonResponse_(deleteQuestion_(payload));
      case 'postExperience':     return jsonResponse_(postExperience_(payload));
      case 'readExperiences':    return jsonResponse_(readExperiences_(payload));
      case 'readGdFeedback':     return jsonResponse_(readGdFeedback_(payload));
      case 'writeGdFeedback':    return jsonResponse_(writeGdFeedback_(payload));
      case 'deleteGdFeedback':   return jsonResponse_(deleteGdFeedback_(payload));
      case 'postPracticeRequest':   return jsonResponse_(postPracticeRequest_(payload));
      case 'readPracticeRequests':  return jsonResponse_(readPracticeRequests_(payload));
      case 'joinPracticeRequest':   return jsonResponse_(joinPracticeRequest_(payload));
      case 'leavePracticeRequest':  return jsonResponse_(leavePracticeRequest_(payload));
      case 'closePracticeRequest':  return jsonResponse_(closePracticeRequest_(payload));
      case 'deletePracticeRequest': return jsonResponse_(deletePracticeRequest_(payload));
      default:
        return null;
    }
  } catch (error) {
    return jsonResponse_({
      status: 'error',
      message: error && error.message ? error.message : '認証処理でエラーが発生しました。'
    });
  }
}

function validateEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function authRegister_(payload) {
  ensureAuthSheets_();

  var email = trimAuthText_(payload.email);
  var displayName = trimAuthText_(payload.displayName);
  var desiredIndustry = trimAuthText_(payload.desiredIndustry);
  var password = trimAuthText_(payload.password);
  var preferredCompanies = getPreferredCompanies_(payload);
  var lineName = trimAuthText_(payload.lineName);
  var lineQrDataUrl = trimAuthText_(payload.lineQrDataUrl);
  var lineQrFileName = trimAuthText_(payload.lineQrFileName);
  var emailKey = normalizeAuthKey_(email);
  var referralCodeInput = trimAuthText_(payload.referralCode);

  assertAuth_(email, 'メールアドレスを入力してください。');
  assertAuth_(validateEmail_(email), '正しいメールアドレスの形式で入力してください。');
  assertAuth_(displayName, '表示名を入力してください。');
  assertAuth_(desiredIndustry, '志望業界を選択してください。');
  assertAuth_(preferredCompanies.length > 0, '第1志望の企業名を入力してください。');
  assertAuth_(lineName, 'LINE名を入力してください。');
  assertAuth_(lineQrDataUrl, 'LINE QRをアップロードしてください。');
  assertAuth_(password.length >= 8, 'パスワードは8文字以上で設定してください。');

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var existing = users.find(function (row) {
    return normalizeAuthKey_(row.emailKey || row.email || row.usernameKey || row.username) === emailKey;
  });
  assertAuth_(!existing, 'このメールアドレスはすでに登録されています。');

  var now = new Date().toISOString();
  var userId = 'user_' + new Date().getTime().toString(36) + Utilities.getUuid().replace(/-/g, '').slice(0, 6);
  var salt = Utilities.getUuid().replace(/-/g, '');
  var passwordHash = sha256Hex_(salt + ':' + password);
  var lineQrAsset = saveLineQrAsset_(userId, lineQrDataUrl, lineQrFileName, '');
  var referralCode = 'ref_' + userId.slice(-6);

  usersSheet.appendRow([
    userId,
    email,
    emailKey,
    desiredIndustry,
    passwordHash,
    salt,
    now,
    now,
    preferredCompanies[0] || '',
    preferredCompanies[1] || '',
    preferredCompanies[2] || '',
    lineName,
    lineQrAsset.fileId,
    lineQrAsset.url,
    displayName,
    referralCode
  ]);

  // 紹介コード処理
  if (referralCodeInput) {
    processReferral_(referralCodeInput, userId, users);
  }

  saveLikedCompaniesForUser_(userId, []);
  var sessionToken = createSessionForUser_(userId);

  return {
    status: 'ok',
    sessionToken: sessionToken,
    user: sanitizeUserRecord_({
      id: userId,
      email: email,
      emailKey: emailKey,
      displayName: displayName,
      desiredIndustry: desiredIndustry,
      createdAt: now,
      updatedAt: now,
      preferredCompany1: preferredCompanies[0] || '',
      preferredCompany2: preferredCompanies[1] || '',
      preferredCompany3: preferredCompanies[2] || '',
      lineName: lineName,
      lineQrDriveFileId: lineQrAsset.fileId,
      lineQrDriveUrl: lineQrAsset.url,
      referralCode: referralCode
    }),
    likedCompanies: []
  };
}

function authLogin_(payload) {
  ensureAuthSheets_();

  var emailKey = normalizeAuthKey_(payload.email || payload.username);
  var password = trimAuthText_(payload.password);
  assertAuth_(emailKey, 'メールアドレスを入力してください。');
  assertAuth_(password, 'パスワードを入力してください。');
  assertAuthRateLimit_('login', emailKey, AUTH_LOGIN_RATE_LIMIT_);

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var user = users.find(function (row) {
    return normalizeAuthKey_(row.emailKey || row.email || row.usernameKey || row.username) === emailKey;
  });
  if (!user) {
    recordAuthRateLimitFailure_('login', emailKey, AUTH_LOGIN_RATE_LIMIT_);
    throw new Error('アカウントが見つかりません。');
  }

  var expectedHash = sha256Hex_(String(user.salt || '') + ':' + password);
  if (expectedHash !== String(user.passwordHash || '')) {
    recordAuthRateLimitFailure_('login', emailKey, AUTH_LOGIN_RATE_LIMIT_);
    throw new Error('パスワードが正しくありません。');
  }
  clearAuthRateLimit_('login', emailKey);

  var sessionToken = createSessionForUser_(user.id);
  return {
    status: 'ok',
    sessionToken: sessionToken,
    user: sanitizeUserRecord_(user),
    likedCompanies: getLikedCompaniesForUser_(user.id)
  };
}

var AUTH_LOGIN_RATE_LIMIT_ = {
  maxAttempts: 8,
  windowSeconds: 15 * 60,
  message: 'ログイン試行回数が多すぎます。15分ほど待ってから再試行してください。'
};
var AUTH_RESET_RATE_LIMIT_ = {
  maxAttempts: 5,
  windowSeconds: 15 * 60,
  message: 'パスワードリセット要求が多すぎます。15分ほど待ってから再試行してください。'
};

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

function recordAuthRateLimitFailure_(action, value, options) {
  var cache = CacheService.getScriptCache();
  var key = getAuthRateLimitKey_(action, value);
  var count = getAuthRateLimitCount_(action, value) + 1;
  cache.put(key, String(count), options.windowSeconds);
}

function clearAuthRateLimit_(action, value) {
  CacheService.getScriptCache().remove(getAuthRateLimitKey_(action, value));
}

function authLogout_(payload) {
  ensureAuthSheets_();
  var sessionToken = trimAuthText_(payload.sessionToken);
  if (!sessionToken) return { status: 'ok' };

  var sessionsSheet = getAuthSheet_(AUTH_SHEET_NAMES_.sessions);
  var sessions = getSheetRecords_(sessionsSheet);
  var rowIndex = findRecordIndex_(sessions, function (row) {
    return row.sessionToken === sessionToken;
  });

  if (rowIndex >= 0) {
    sessionsSheet.getRange(rowIndex + 2, 6).setValue('0');
  }

  return { status: 'ok' };
}

function authUpdateProfile_(payload) {
  ensureAuthSheets_();

  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var displayName = trimAuthText_(payload.displayName || payload.username);
  var desiredIndustry = trimAuthText_(payload.desiredIndustry);
  var preferredCompanies = getPreferredCompanies_(payload);
  var lineName = trimAuthText_(payload.lineName);
  var lineQrDataUrl = trimAuthText_(payload.lineQrDataUrl);
  var lineQrFileName = trimAuthText_(payload.lineQrFileName);

  assertAuth_(displayName, '表示名を入力してください。');
  assertAuth_(desiredIndustry, '志望業界を選択してください。');
  assertAuth_(preferredCompanies.length > 0, '第1志望の企業名を入力してください。');
  assertAuth_(lineName, 'LINE名を入力してください。');

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var currentIndex = findRecordIndex_(users, function (row) {
    return row.id === session.userId;
  });
  assertAuth_(currentIndex >= 0, 'アカウントが見つかりません。');

  var user = users[currentIndex];
  var lineQrAsset = {
    fileId: trimAuthText_(user.lineQrDriveFileId),
    url: trimAuthText_(user.lineQrDriveUrl)
  };

  if (!lineQrDataUrl) {
    assertAuth_(lineQrAsset.url, 'LINE QRをアップロードしてください。');
  }

  if (lineQrDataUrl) {
    lineQrAsset = saveLineQrAsset_(session.userId, lineQrDataUrl, lineQrFileName, lineQrAsset.fileId);
  }

  var updatedAt = new Date().toISOString();
  // メールアドレスは変更不可（カラム2,3はemail,emailKeyのまま）
  usersSheet.getRange(currentIndex + 2, 4).setValue(desiredIndustry);
  usersSheet.getRange(currentIndex + 2, 8).setValue(updatedAt);
  usersSheet.getRange(currentIndex + 2, 9, 1, 6).setValues([[
    preferredCompanies[0] || '',
    preferredCompanies[1] || '',
    preferredCompanies[2] || '',
    lineName,
    lineQrAsset.fileId,
    lineQrAsset.url
  ]]);
  usersSheet.getRange(currentIndex + 2, 15).setValue(displayName);

  return {
    status: 'ok',
    user: sanitizeUserRecord_({
      id: session.userId,
      email: trimAuthText_(user.email || user.username),
      emailKey: trimAuthText_(user.emailKey || user.usernameKey),
      displayName: displayName,
      desiredIndustry: desiredIndustry,
      createdAt: user.createdAt || '',
      updatedAt: updatedAt,
      preferredCompany1: preferredCompanies[0] || '',
      preferredCompany2: preferredCompanies[1] || '',
      preferredCompany3: preferredCompanies[2] || '',
      lineName: lineName,
      lineQrDriveFileId: lineQrAsset.fileId,
      lineQrDriveUrl: lineQrAsset.url,
      referralCode: trimAuthText_(user.referralCode)
    })
  };
}

function authChangePassword_(payload) {
  ensureAuthSheets_();

  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var currentPassword = trimAuthText_(payload.currentPassword);
  var nextPassword = trimAuthText_(payload.nextPassword);
  assertAuth_(currentPassword && nextPassword, '現在のパスワードと新しいパスワードを入力してください。');
  assertAuth_(nextPassword.length >= 8, '新しいパスワードは8文字以上で設定してください。');

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var currentIndex = findRecordIndex_(users, function (row) {
    return row.id === session.userId;
  });
  assertAuth_(currentIndex >= 0, 'アカウントが見つかりません。');

  var user = users[currentIndex];
  var currentHash = sha256Hex_(String(user.salt || '') + ':' + currentPassword);
  assertAuth_(currentHash === String(user.passwordHash || ''), '現在のパスワードが正しくありません。');

  usersSheet.getRange(currentIndex + 2, 5).setValue(sha256Hex_(String(user.salt || '') + ':' + nextPassword));
  usersSheet.getRange(currentIndex + 2, 8).setValue(new Date().toISOString());

  return { status: 'ok' };
}

function authDeleteAccount_(payload) {
  ensureAuthSheets_();

  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var password = trimAuthText_(payload.password);
  assertAuth_(password, 'パスワードを入力してください。');

  // パスワード確認
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var userIndex = findRecordIndex_(users, function (row) {
    return row.id === session.userId;
  });
  assertAuth_(userIndex >= 0, 'アカウントが見つかりません。');

  var user = users[userIndex];
  var expectedHash = sha256Hex_(String(user.salt || '') + ':' + password);
  assertAuth_(expectedHash === String(user.passwordHash || ''), 'パスワードが正しくありません。');

  // LINE QRファイル削除
  var lineQrFileId = trimAuthText_(user.lineQrDriveFileId);
  if (lineQrFileId) {
    try { DriveApp.getFileById(lineQrFileId).setTrashed(true); } catch (e) {}
  }

  // ユーザー行を削除
  usersSheet.deleteRow(userIndex + 2);

  // セッション全削除
  var sessionsSheet = getAuthSheet_(AUTH_SHEET_NAMES_.sessions);
  var sessions = getSheetRecords_(sessionsSheet);
  for (var si = sessions.length - 1; si >= 0; si--) {
    if (sessions[si].userId === session.userId) {
      sessionsSheet.deleteRow(si + 2);
    }
  }

  // Liked削除
  var likedSheet = getAuthSheet_(AUTH_SHEET_NAMES_.liked);
  var likedRows = getSheetRecords_(likedSheet);
  for (var li = likedRows.length - 1; li >= 0; li--) {
    if (likedRows[li].userId === session.userId) {
      likedSheet.deleteRow(li + 2);
    }
  }

  // 進捗トラッカー削除
  var spreadsheet = getAuthSpreadsheet_();
  var progressSheet = spreadsheet.getSheetByName(PROGRESS_SHEET_NAME_);
  if (progressSheet) {
    var progressRows = getSheetRecords_(progressSheet);
    for (var pi = progressRows.length - 1; pi >= 0; pi--) {
      if (progressRows[pi].userId === session.userId) {
        progressSheet.deleteRow(pi + 2);
      }
    }
  }

  // ES削除
  var esSheetNames = ['user_es'];
  esSheetNames.forEach(function (name) {
    var sheet = spreadsheet.getSheetByName(name);
    if (!sheet) return;
    var rows = getSheetRecords_(sheet);
    for (var i = rows.length - 1; i >= 0; i--) {
      if (rows[i].userId === session.userId) sheet.deleteRow(i + 2);
    }
  });

  // ガクチカ削除
  var gkSheet = spreadsheet.getSheetByName('user_gakuchika');
  if (gkSheet) {
    var gkRows = getSheetRecords_(gkSheet);
    for (var gi = gkRows.length - 1; gi >= 0; gi--) {
      if (gkRows[gi].userId === session.userId) gkSheet.deleteRow(gi + 2);
    }
  }

  // 面接対策データ削除
  ['ip_companies', 'ip_questions'].forEach(function (name) {
    var sheet = spreadsheet.getSheetByName(name);
    if (!sheet) return;
    var rows = getSheetRecords_(sheet);
    for (var i = rows.length - 1; i >= 0; i--) {
      if (rows[i].userId === session.userId) sheet.deleteRow(i + 2);
    }
  });

  // マッチング設定削除
  var prefsSheet = spreadsheet.getSheetByName(MATCHING_PREFS_SHEET_NAME_);
  if (prefsSheet) {
    var prefsRows = getSheetRecords_(prefsSheet);
    for (var mi = prefsRows.length - 1; mi >= 0; mi--) {
      if (prefsRows[mi].userId === session.userId) prefsSheet.deleteRow(mi + 2);
    }
  }

  // グループメンバーシップ削除
  var gmSheet = spreadsheet.getSheetByName(GROUP_MEMBERS_SHEET_NAME_);
  if (gmSheet) {
    var gmRows = getSheetRecords_(gmSheet);
    for (var gmi = gmRows.length - 1; gmi >= 0; gmi--) {
      if (gmRows[gmi].userId === session.userId) gmSheet.deleteRow(gmi + 2);
    }
  }

  // 掲示板投稿は匿名化（削除ではなくユーザー名を「退会済みユーザー」に）
  var boardSheet = spreadsheet.getSheetByName(BOARD_SHEET_NAME_);
  if (boardSheet) {
    var boardRows = getSheetRecords_(boardSheet);
    boardRows.forEach(function (row, i) {
      if (row.userId === session.userId) {
        boardSheet.getRange(i + 2, 2, 1, 2).setValues([['deleted_user', '退会済みユーザー']]);
      }
    });
  }

  return { status: 'ok', message: 'アカウントを削除しました。ご利用ありがとうございました。' };
}

function authSetLikedCompanies_(payload) {
  ensureAuthSheets_();

  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var likedCompanies = Array.isArray(payload.likedCompanies) ? payload.likedCompanies : [];
  saveLikedCompaniesForUser_(session.userId, likedCompanies);

  return {
    status: 'ok',
    likedCompanies: getLikedCompaniesForUser_(session.userId)
  };
}

function getActiveSessionOrThrow_(sessionToken) {
  var token = trimAuthText_(sessionToken);
  assertAuth_(token, 'ログイン情報が見つかりません。');

  var sessionsSheet = getAuthSheet_(AUTH_SHEET_NAMES_.sessions);
  var sessions = getSheetRecords_(sessionsSheet);
  var now = new Date();
  var index = findRecordIndex_(sessions, function (row) {
    return row.sessionToken === token && String(row.active) !== '0';
  });

  assertAuth_(index >= 0, 'ログイン情報が見つかりません。');

  var session = sessions[index];
  var expiresAt = new Date(session.expiresAt);
  assertAuth_(!isNaN(expiresAt.getTime()) && expiresAt.getTime() > now.getTime(), 'ログインの有効期限が切れました。');

  sessionsSheet.getRange(index + 2, 5).setValue(now.toISOString());
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
  var sessionToken = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var now = new Date();
  var expiresAt = new Date(now.getTime() + AUTH_SESSION_TTL_DAYS_ * 24 * 60 * 60 * 1000);

  sessionsSheet.appendRow([
    sessionToken,
    userId,
    now.toISOString(),
    expiresAt.toISOString(),
    now.toISOString(),
    '1'
  ]);

  return sessionToken;
}

function getLikedCompaniesForUser_(userId) {
  var likedSheet = getAuthSheet_(AUTH_SHEET_NAMES_.liked);
  var rows = getSheetRecords_(likedSheet);
  var row = rows.find(function (record) { return record.userId === userId; });
  if (!row || !row.likedCompaniesJson) return [];

  try {
    return JSON.parse(row.likedCompaniesJson);
  } catch (error) {
    return [];
  }
}

function saveLikedCompaniesForUser_(userId, likedCompanies) {
  var likedSheet = getAuthSheet_(AUTH_SHEET_NAMES_.liked);
  var rows = getSheetRecords_(likedSheet);
  var rowIndex = findRecordIndex_(rows, function (record) { return record.userId === userId; });
  var payload = JSON.stringify(likedCompanies || []);
  var updatedAt = new Date().toISOString();

  if (rowIndex >= 0) {
    likedSheet.getRange(rowIndex + 2, 2, 1, 2).setValues([[payload, updatedAt]]);
    return;
  }

  likedSheet.appendRow([userId, payload, updatedAt]);
}

function sanitizeUserRecord_(record) {
  var preferredCompanies = getPreferredCompanies_(record);
  var lineQrUrl = trimAuthText_(record.lineQrDriveUrl || record.lineQrUrl);
  var email = trimAuthText_(record.email || record.username);
  var displayName = trimAuthText_(record.displayName) || email.split('@')[0];

  return {
    id: trimAuthText_(record.id),
    email: email,
    emailKey: trimAuthText_(record.emailKey || record.usernameKey),
    displayName: displayName,
    // 後方互換性
    username: displayName,
    usernameKey: trimAuthText_(record.emailKey || record.usernameKey),
    desiredIndustry: trimAuthText_(record.desiredIndustry),
    preferredCompanies: preferredCompanies,
    preferredCompany1: preferredCompanies[0] || '',
    preferredCompany2: preferredCompanies[1] || '',
    preferredCompany3: preferredCompanies[2] || '',
    lineName: trimAuthText_(record.lineName),
    lineQrUrl: lineQrUrl,
    hasLineQr: !!lineQrUrl,
    referralCode: trimAuthText_(record.referralCode),
    createdAt: trimAuthText_(record.createdAt),
    updatedAt: trimAuthText_(record.updatedAt)
  };
}

function ensureAuthSheets_() {
  ensureSheet_(AUTH_SHEET_NAMES_.users, AUTH_USER_HEADERS_);

  ensureSheet_(AUTH_SHEET_NAMES_.sessions, [
    'sessionToken',
    'userId',
    'createdAt',
    'expiresAt',
    'lastSeenAt',
    'active'
  ]);

  ensureSheet_(AUTH_SHEET_NAMES_.liked, [
    'userId',
    'likedCompaniesJson',
    'updatedAt'
  ]);
}

function ensureSheet_(name, headers) {
  var spreadsheet = getAuthSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }

  var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach(function (header) {
    if (existingHeaders.indexOf(header) !== -1) return;
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
    existingHeaders.push(header);
  });
}

function getAuthSpreadsheet_() {
  var spreadsheetId = PropertiesService.getScriptProperties().getProperty('AUTH_SPREADSHEET_ID');
  if (spreadsheetId) return SpreadsheetApp.openById(spreadsheetId);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getAuthSheet_(name) {
  return getAuthSpreadsheet_().getSheetByName(name);
}

function getSheetRecords_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return [];
  var headers = values[0];

  return values.slice(1).map(function (row) {
    var record = {};
    headers.forEach(function (header, index) {
      record[header] = row[index];
    });
    return record;
  });
}

function findRecordIndex_(rows, predicate) {
  for (var i = 0; i < rows.length; i += 1) {
    if (predicate(rows[i])) return i;
  }
  return -1;
}

function getPreferredCompanies_(payload) {
  var raw = Array.isArray(payload.preferredCompanies)
    ? payload.preferredCompanies
    : [payload.preferredCompany1, payload.preferredCompany2, payload.preferredCompany3];
  var unique = [];

  raw.forEach(function (item) {
    var value = trimAuthText_(item);
    if (!value) return;
    var exists = unique.some(function (current) {
      return normalizeAuthKey_(current) === normalizeAuthKey_(value);
    });
    if (!exists) unique.push(value);
  });

  return unique.slice(0, 3);
}

function saveLineQrAsset_(userId, lineQrDataUrl, lineQrFileName, existingFileId) {
  var dataUrl = trimAuthText_(lineQrDataUrl);
  var match = dataUrl.match(/^data:(image\/[^;]+);base64,(.+)$/);
  assertAuth_(match, 'LINE QRは画像ファイルをアップロードしてください。');

  var mimeType = match[1];
  var bytes = Utilities.base64Decode(match[2]);
  var extension = getFileExtensionFromMimeType_(mimeType);
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
  var baseName = trimAuthText_(lineQrFileName).replace(/[^\w.\-]+/g, '_');
  var fileName = 'line-qr-' + userId + '-' + timestamp + extension;
  if (baseName) {
    fileName = 'line-qr-' + userId + '-' + timestamp + '-' + baseName.replace(/\.[^.]+$/, '') + extension;
  }

  var blob = Utilities.newBlob(bytes, mimeType, fileName);
  var folder = getLineQrFolder_();
  var file = folder.createFile(blob);

  if (existingFileId) {
    try {
      DriveApp.getFileById(existingFileId).setTrashed(true);
    } catch (error) {}
  }

  return {
    fileId: file.getId(),
    url: file.getUrl()
  };
}

function getLineQrFolder_() {
  var folderId = PropertiesService.getScriptProperties().getProperty('AUTH_LINE_QR_FOLDER_ID');
  if (folderId) return DriveApp.getFolderById(folderId);
  return DriveApp.getRootFolder();
}

function getFileExtensionFromMimeType_(mimeType) {
  switch (mimeType) {
    case 'image/jpeg':
      return '.jpg';
    case 'image/png':
      return '.png';
    case 'image/gif':
      return '.gif';
    case 'image/webp':
      return '.webp';
    default:
      return '.img';
  }
}

function normalizeAuthKey_(value) {
  return trimAuthText_(value).normalize('NFKC').toLowerCase();
}

function trimAuthText_(value) {
  return String(value || '').trim();
}

function assertAuth_(condition, message) {
  if (!condition) throw new Error(message);
}

function sha256Hex_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(function (byte) {
    var normalized = byte < 0 ? byte + 256 : byte;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ES添削掲示板 ────────────────────────────────────────────────────────────

var BOARD_SHEET_NAME_         = 'es_board';
var BOARD_COMMENTS_SHEET_NAME_ = 'es_board_comments';
var BOARD_MAX_POSTS_           = 50;

function writeBoardPost_(payload) {
  var session  = getActiveSessionOrThrow_(payload.sessionToken);
  var company  = trimAuthText_(payload.company);
  var question = trimAuthText_(payload.question);
  var esText   = trimAuthText_(payload.esText);

  assertAuth_(company,          '企業名を入力してください。');
  assertAuth_(question,         '設問テキストを入力してください。');
  assertAuth_(esText.length > 0, 'ESの内容を入力してください。');

  var spreadsheet = getAuthSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(BOARD_SHEET_NAME_);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(BOARD_SHEET_NAME_);
    sheet.appendRow(['id', 'userId', 'username', 'company', 'question', 'esText', 'createdAt']);
  }

  var now       = new Date().toISOString();
  var id        = now.replace(/\D/g, '').slice(0, 14) + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users      = getSheetRecords_(usersSheet);
  var userRecord = users.find(function (row) { return row.id === session.userId; });
  var username   = userRecord ? trimAuthText_(userRecord.displayName || userRecord.username) : '';

  sheet.appendRow([id, session.userId, username, company, question, esText, now]);

  return { status: 'ok', id: id };
}

function readBoardPosts_(payload) {
  getActiveSessionOrThrow_(payload.sessionToken);

  var spreadsheet = getAuthSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(BOARD_SHEET_NAME_);
  if (!sheet) return { status: 'ok', posts: [] };

  var posts = getSheetRecords_(sheet);
  if (!posts.length) return { status: 'ok', posts: [] };

  posts.sort(function (a, b) {
    return new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime();
  });
  posts = posts.slice(0, BOARD_MAX_POSTS_);

  var commentsSheet = spreadsheet.getSheetByName(BOARD_COMMENTS_SHEET_NAME_);
  var allComments   = commentsSheet ? getSheetRecords_(commentsSheet) : [];

  var result = posts.map(function (post) {
    var comments = allComments.filter(function (c) { return c.postId === post.id; });
    comments.sort(function (a, b) {
      return new Date(a.createdAt || 0).getTime() - new Date(b.createdAt || 0).getTime();
    });
    return {
      id:        trimAuthText_(post.id),
      userId:    trimAuthText_(post.userId),
      username:  trimAuthText_(post.username),
      company:   trimAuthText_(post.company),
      question:  trimAuthText_(post.question),
      esText:    trimAuthText_(post.esText),
      createdAt: trimAuthText_(post.createdAt),
      comments:  comments.map(function (c) {
        return {
          id:          trimAuthText_(c.id),
          postId:      trimAuthText_(c.postId),
          userId:      trimAuthText_(c.userId),
          username:    trimAuthText_(c.username),
          commentText: trimAuthText_(c.commentText),
          createdAt:   trimAuthText_(c.createdAt)
        };
      })
    };
  });

  return { status: 'ok', posts: result };
}

function addBoardComment_(payload) {
  var session     = getActiveSessionOrThrow_(payload.sessionToken);
  var postId      = trimAuthText_(payload.postId);
  var commentText = trimAuthText_(payload.commentText);

  assertAuth_(postId,      'postIdが指定されていません。');
  assertAuth_(commentText, 'コメントを入力してください。');

  var spreadsheet = getAuthSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(BOARD_COMMENTS_SHEET_NAME_);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(BOARD_COMMENTS_SHEET_NAME_);
    sheet.appendRow(['id', 'postId', 'userId', 'username', 'commentText', 'createdAt']);
  }

  var now        = new Date().toISOString();
  var id         = now.replace(/\D/g, '').slice(0, 14) + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users      = getSheetRecords_(usersSheet);
  var userRecord = users.find(function (row) { return row.id === session.userId; });
  var username   = userRecord ? trimAuthText_(userRecord.username) : '';

  sheet.appendRow([id, postId, session.userId, username, commentText, now]);

  return { status: 'ok' };
}

// ── Claude API 呼び出し ────────────────────────────────────────────────────

function callClaude_(payload) {
  getActiveSessionOrThrow_(payload.sessionToken);

  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  assertAuth_(apiKey, 'AI機能が設定されていません。管理者にお問い合わせください。');

  var toolType = trimAuthText_(payload.toolType);
  var input    = payload.input || {};
  var systemPrompt, userMessage;

  if (toolType === 'esReview') {
    var company   = trimAuthText_(input.company);
    var question  = trimAuthText_(input.question);
    var esText    = trimAuthText_(input.esText);
    var charLimit = input.charLimit ? parseInt(input.charLimit) : 0;
    assertAuth_(esText, 'ESの内容を入力してください。');

    var limitNote = charLimit > 0 ? '\n【字数制限】' + charLimit + '字（現在' + esText.length + '字）' : '';
    var limitGuide = charLimit > 0 ? '\n6. 📏 字数制限への適合（' + charLimit + '字制限に対して過不足がないか、削る/増やすべき箇所の提案）' : '';
    systemPrompt = 'あなたは新卒就職活動のプロフェッショナルなESコーチです。慶應義塾大学の学生のESを丁寧に添削してください。フィードバックは日本語で、具体的かつ建設的に行ってください。';
    userMessage  = '【企業名】' + (company || '（未記入）') + '\n【設問】' + (question || '（未記入）') + limitNote + '\n\n【ES内容】\n' + esText + '\n\n以下の観点で詳しく添削してください：\n1. 📌 構成・論理性（PREP法など）\n2. 💡 具体性・エピソードの説得力\n3. 🎯 企業・設問への適合性\n4. ✍️ 表現・文章力\n5. 🔧 改善提案（修正例を示す）' + limitGuide + '\n\n最後に総合評価（S/A/B/C）と一言コメントをつけてください。';

  } else if (toolType === 'interviewCoach') {
    var intQuestion = trimAuthText_(input.question);
    var answer      = trimAuthText_(input.answer);
    var jobType     = trimAuthText_(input.jobType);
    assertAuth_(answer, '回答を入力してください。');

    systemPrompt = 'あなたは面接官の経験豊富な就活コーチです。慶應義塾大学の学生の面接回答を評価し、具体的な改善アドバイスをしてください。日本語で回答してください。';
    userMessage  = '【面接質問】' + (intQuestion || '自己PR') + '\n【志望職種】' + (jobType || '（未記入）') + '\n\n【学生の回答】\n' + answer + '\n\n以下の観点で評価してください：\n1. ⭐ 総合評価（10点満点）\n2. 📐 構成（STAR法など）\n3. 💪 強みの伝わり方\n4. 🏢 企業・職種への関連性\n5. 😊 熱意・志望度の伝わり方\n6. 🔧 改善アドバイス（具体的なフレーズ例を含む）\n\n最後に「このまま言えばOK」な改善版回答例を提示してください。';

  } else if (toolType === 'esRewrite') {
    var originalEs     = trimAuthText_(input.originalEs);
    var targetCompany  = trimAuthText_(input.targetCompany);
    var targetQuestion = trimAuthText_(input.targetQuestion);
    assertAuth_(originalEs,    '元のES内容を入力してください。');
    assertAuth_(targetCompany, '志望企業名を入力してください。');

    systemPrompt = 'あなたは就活のプロフェッショナルです。慶應義塾大学の学生のESを別企業向けに最適化してリライトしてください。日本語で回答してください。';
    userMessage  = '【元のES】\n' + originalEs + '\n\n【新しい志望企業】' + targetCompany + '\n【新しい設問】' + (targetQuestion || '（元の設問に準じる）') + '\n\n元のESの強みを活かしながら、新しい企業・設問に最適化したESを作成してください。\n\n出力形式：\n1. 📝 リライト版ES（すぐ使える形で）\n2. 🔄 変更点とその理由\n3. ✅ 追加で強化できるポイント';

  } else if (toolType === 'esMultiReview') {
    var entries = Array.isArray(input.entries) ? input.entries : [];
    assertAuth_(entries.length > 0, '添削するESを入力してください。');

    var esListText = entries.map(function(e, i) {
      var limit = e.charLimit ? parseInt(e.charLimit) : 0;
      var limitNote = limit > 0 ? '（字数制限: ' + limit + '字 / 現在: ' + trimAuthText_(e.esText).length + '字）' : '';
      return '【設問' + (i + 1) + '】' + (trimAuthText_(e.question) || '（設問未記入）') + limitNote + '\n\n' + trimAuthText_(e.esText);
    }).join('\n\n---\n\n');

    systemPrompt = 'あなたは新卒就職活動のプロフェッショナルなESコーチです。慶應義塾大学の学生の複数のES設問を総合的に添削してください。各設問の個別評価に加え、複数設問を通じた一貫性・多面性も評価してください。日本語で回答してください。\n\n【重要】回答の最初に、以下の形式でルーブリック評価スコア（各1〜5の整数）をJSON形式で出力してください。必ずこの形式で始めてください：\n[SCORES]{"structure":X,"specificity":X,"logic":X,"expression":X,"persuasion":X}[/SCORES]\n\n各スコアの基準：\n- structure（構成）: 1=構成が不明確 2=やや整理不足 3=基本的な構成はある 4=論理的で分かりやすい 5=完璧な構成\n- specificity（具体性）: 1=抽象的すぎる 2=具体性が不足 3=一部具体的 4=エピソードが具体的 5=非常に具体的で説得力大\n- logic（論理性）: 1=論理破綻 2=やや飛躍がある 3=概ね論理的 4=論理的で一貫性あり 5=完璧な論理展開\n- expression（表現力）: 1=表現が稚拙 2=やや単調 3=標準的 4=表現が豊か 5=非常に洗練された表現\n- persuasion（説得力）: 1=説得力なし 2=やや弱い 3=一定の説得力 4=説得力がある 5=非常に説得力が高い';
    userMessage  = '【企業名】' + (trimAuthText_(input.company) || '（未記入）') + '\n\n' + esListText + '\n\n' +
      '最初に [SCORES]{"structure":X,"specificity":X,"logic":X,"expression":X,"persuasion":X}[/SCORES] の形式でルーブリックスコア（各1〜5）を出力し、その後に以下を評価してください。\n\n# 評価してください\n\n## ① 各設問の個別評価\n各設問について以下を評価：\n- 構成・論理性（PREP法など）\n- 具体性・エピソードの説得力\n- 字数制限が指定されている場合は字数への適合性（削る/増やすべき箇所の提案）\n- 改善提案（修正例を含む）\n\n## ② 複数設問の総合評価\n1. 🔗 **一貫性** — 自己PR・強みのテーマにブレがないか\n2. 🌐 **多面性** — 異なる強み・経験を引き出せているか\n3. 🎯 **企業適合性** — 全体を通じて志望企業への熱意・適性が伝わるか\n4. ✅ **総合評価（S/A/B/C）と優先改善ポイント**';

  } else if (toolType === 'generateEsFromGakuchika') {
    var gItems = Array.isArray(input.gakuchika) ? input.gakuchika : (input.gakuchika ? [input.gakuchika] : []);
    assertAuth_(gItems.length > 0, 'ガクチカを選択してください。');
    var genCompany  = trimAuthText_(input.company);
    var genQuestion = trimAuthText_(input.question);
    var genLimit    = parseInt(input.charLimit) > 0 ? parseInt(input.charLimit) : 0;

    var gakuchikaText = gItems.map(function(g, i) {
      var parts = [];
      if (g.title)     parts.push('【タイトル】' + g.title);
      if (g.category)  parts.push('【カテゴリ】' + g.category);
      if (g.situation) parts.push('【状況 (S)】' + g.situation);
      if (g.task)      parts.push('【課題・目標 (T)】' + g.task);
      if (g.action)    parts.push('【行動 (A)】' + g.action);
      if (g.result)    parts.push('【結果・成果 (R)】' + g.result);
      if (g.appeal)    parts.push('【アピールポイント・学び】' + g.appeal);
      return (gItems.length > 1 ? '＜エピソード' + (i + 1) + '＞\n' : '') + parts.join('\n');
    }).join('\n\n');

    var genLimitNote = genLimit > 0 ? '字数制限は' + genLimit + '字以内です。必ずこの字数内に収めてください。' : '字数制限は特に指定されていません。';

    systemPrompt = 'あなたは新卒就職活動のプロフェッショナルなESコーチです。慶應義塾大学の学生がメモしたガクチカ（学生時代に力を入れたこと）の情報をもとに、指定された企業・設問に最適化したESを作成してください。日本語で回答してください。';
    userMessage  = '# ガクチカ情報\n\n' + gakuchikaText +
      '\n\n# ES作成条件\n' +
      '【志望企業】' + (genCompany || '（未指定）') + '\n' +
      '【設問】' + (genQuestion || 'ガクチカを教えてください') + '\n' +
      '【字数】' + genLimitNote + '\n\n' +
      '# 出力形式\n\n' +
      '## ① 生成ES（そのまま使えるES本文）\n' +
      '※STAR法（状況→課題→行動→結果）を意識し、数値・固有名詞で具体化してください。\n\n' +
      '## ② 企業・設問への適合ポイント\n' +
      '（なぜこのエピソード・表現を選んだか）\n\n' +
      '## ③ さらに強化するためのアドバイス\n' +
      '（追加で調べると効果的なこと、言い換えられるフレーズなど）';

  } else if (toolType === 'generateInterviewQuestions') {
    var gakuchikaText = trimAuthText_(input.gakuchikaText || '');
    var ipCompany     = trimAuthText_(input.company  || '');
    var ipPosition    = trimAuthText_(input.position || '');
    assertAuth_(gakuchikaText, 'ガクチカを入力してください。');

    systemPrompt = 'あなたは日系大手企業・外資系企業の面接を熟知した採用コンサルタントです。慶應義塾大学の学生が面接に向けて準備できるよう、鋭い深掘り質問を生成してください。日本語で回答してください。';
    userMessage  = '# ガクチカ\n' + gakuchikaText +
      '\n\n# 志望企業・職種\n企業: ' + (ipCompany || '（未記入）') + '\n職種: ' + (ipPosition || '（未記入）') +
      '\n\n# 指示\n上記ガクチカに対して、面接官が実際に聞いてくる「深掘り質問」を10〜12問生成してください。\n\n生成ルール:\n- 「なぜ」「具体的に」「他の選択肢は」「うまくいかなかった点は」「学んだことは」「それが今どう活きているか」など多角的に\n- 圧迫質問・弱点を突く質問も含める\n- 企業・職種が記入されている場合は、その企業文化・職種特性に合わせた質問も加える\n- 各質問はシンプルに1文で\n\n出力形式（必ずこのJSON形式で出力してください）:\n```json\n[\n  {"category": "動機・背景", "question": "なぜその活動を始めたのですか？"},\n  {"category": "具体的行動", "question": "..."},\n  ...\n]\n```\n\ncategory は「動機・背景」「具体的行動」「困難・課題」「思考プロセス」「結果・成果」「自己成長」「企業接続」のいずれかを使用してください。';

  } else if (toolType === 'evaluateInterviewAnswer') {
    var evQuestion = trimAuthText_(input.question || '');
    var evAnswer   = trimAuthText_(input.answer   || '');
    var evCompany  = trimAuthText_(input.company  || '');
    var evPosition = trimAuthText_(input.position || '');
    assertAuth_(evQuestion, '質問を入力してください。');
    assertAuth_(evAnswer,   '回答を入力してください。');

    systemPrompt = 'あなたは日系大手・外資系企業の採用面接を熟知したキャリアコーチです。学生の面接回答を厳しくかつ建設的に評価してください。日本語で回答してください。';
    userMessage  = '# 面接質問\n' + evQuestion +
      '\n\n# 学生の回答\n' + evAnswer +
      '\n\n# 企業・職種（参考）\n企業: ' + (evCompany || '未記入') + ' / 職種: ' + (evPosition || '未記入') +
      '\n\n# 評価してください\n\n## ① 総合評価（S/A/B/C）と一言コメント\n\n## ② 良かった点（具体的に）\n\n## ③ 改善点（具体的に・優先順位順）\n\n## ④ 改善版の回答例\n（「このまま言えばOK」な形で、簡潔に提示してください）\n\n## ⑤ 追加で深掘りされそうな質問\n（1〜2問）';

  } else if (toolType === 'generateJibunshi') {
    var periods = input.periods || {};
    var elem    = trimAuthText_(periods.elementary || '');
    var junior  = trimAuthText_(periods.junior     || '');
    var high    = trimAuthText_(periods.high       || '');
    var college = trimAuthText_(periods.college    || '');
    var traits  = trimAuthText_(input.traits       || '');
    assertAuth_(elem || junior || high || college, '少なくとも1つの時代の情報を入力してください。');

    // ガクチカデータ（任意）
    var gkItems = Array.isArray(input.gakuchika) ? input.gakuchika : [];
    var gkText  = '';
    if (gkItems.length > 0) {
      gkText = '# 登録済みガクチカ（参考情報）\n以下のガクチカを自分史の大学以降・高校時代パートに適切に組み込んでください。\n\n';
      gkItems.forEach(function(g, i) {
        var parts = [];
        if (g.title)     parts.push('【タイトル】' + trimAuthText_(g.title));
        if (g.category)  parts.push('【カテゴリ】' + trimAuthText_(g.category));
        if (g.situation) parts.push('【状況】'     + trimAuthText_(g.situation));
        if (g.task)      parts.push('【課題】'     + trimAuthText_(g.task));
        if (g.action)    parts.push('【行動】'     + trimAuthText_(g.action));
        if (g.result)    parts.push('【結果】'     + trimAuthText_(g.result));
        if (g.appeal)    parts.push('【アピール】' + trimAuthText_(g.appeal));
        gkText += '＜ガクチカ' + (i + 1) + '＞\n' + parts.join('\n') + '\n\n';
      });
    }

    var periodText = '';
    if (elem)    periodText += '## 小学校以前\n' + elem + '\n\n';
    if (junior)  periodText += '## 中学時代\n' + junior + '\n\n';
    if (high)    periodText += '## 高校時代\n' + high + '\n\n';
    if (college) periodText += '## 大学以降\n' + college + '\n\n';

    systemPrompt = 'あなたは三井物産などの総合商社の採用を熟知したキャリアコーチです。慶應義塾大学の学生の「自分史」作成を支援します。\n\n' +
      '【自分史とは】志望動機を述べる書類ではなく、どのような経験・環境の中でその人物が形成されてきたかを時系列で見せる書類です。人物の一貫性・価値観の根幹を採用担当者に伝えることが目的です。\n\n' +
      '【執筆ルール】\n' +
      '① 合計2000字程度（最大2500字）\n' +
      '② 4つの時代に分けて記述し、字数比率は「小学校以前：中学：高校：大学以降 ＝ 10：20：30：40」を厳守する（小学校以前約200字、中学約400字、高校約600字、大学約800字）\n' +
      '③ 時系列で端的に書き、経緯や感想は簡潔に補足する（冗長にならない）\n' +
      '④ 志望動機の記述は不要。あくまで人物形成・価値観の一貫性にフォーカスする\n' +
      '⑤ 各時代の見出しは【小学校以前】【中学時代】【高校時代】【大学以降】の形式を使う\n' +
      '⑥ 全体を通じて「この人物の軸・一貫したテーマ」が自然に浮かび上がるよう構成する\n\n' +
      '日本語で出力してください。';

    userMessage = '# 各時代のエピソード・情報\n\n' + periodText +
      (gkText ? gkText : '') +
      (traits ? '# アピールしたい特性・一貫したテーマ（任意）\n' + traits + '\n\n' : '') +
      '# 指示\n上記の情報をもとに、自分史を執筆してください。\n' +
      '出力は自分史の本文のみとし、各時代の見出し【】を含めてください。\n' +
      '合計2000字程度（最大2500字）、字数比率10:20:30:40を守ってください。\n' +
      '各セクションの末尾に（○字）と文字数を括弧書きで記してください。';

  } else if (toolType === 'esChatMessage') {
    var chatMessages = Array.isArray(input.messages) ? input.messages : [];
    assertAuth_(chatMessages.length > 0, 'メッセージがありません。');
    var chatSystem = trimAuthText_(input.systemPrompt) ||
      'あなたは新卒就職活動のプロフェッショナルなESコーチです。慶應義塾大学の学生のESについて対話形式でサポートしています。日本語で具体的かつ建設的にアドバイスしてください。';

    var chatBody = JSON.stringify({
      model:      'claude-opus-4-6',
      max_tokens: 2000,
      system:     chatSystem,
      messages:   chatMessages
    });
    var chatResp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: chatBody, muteHttpExceptions: true
    });
    if (chatResp.getResponseCode() !== 200) {
      var ce = {}; try { ce = JSON.parse(chatResp.getContentText()); } catch(e2) {}
      throw new Error('AI処理に失敗しました: ' + ((ce.error && ce.error.message) || chatResp.getResponseCode()));
    }
    var chatData = JSON.parse(chatResp.getContentText());
    var chatResult = '';
    (chatData.content || []).forEach(function(b) { if (b.type === 'text') chatResult += b.text; });
    return { status: 'ok', result: chatResult };

  } else {
    throw new Error('不明なツールタイプです。');
  }

  var requestBody = {
    model:      'claude-opus-4-6',
    max_tokens: 2000,
    system:     systemPrompt,
    messages:   [{ role: 'user', content: userMessage }]
  };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:           'post',
    contentType:      'application/json',
    headers: {
      'x-api-key':          apiKey,
      'anthropic-version':  '2023-06-01'
    },
    payload:          JSON.stringify(requestBody),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var body = response.getContentText();

  if (code !== 200) {
    var errData = {};
    try { errData = JSON.parse(body); } catch (e) {}
    var errMsg = (errData.error && errData.error.message) ? errData.error.message : ('HTTP ' + code);
    throw new Error('AI処理に失敗しました: ' + errMsg);
  }

  var data = JSON.parse(body);
  var resultText = '';
  if (data.content && data.content.length) {
    data.content.forEach(function (block) {
      if (block.type === 'text') resultText += block.text;
    });
  }

  return { status: 'ok', result: resultText };
}

// ── ES保管庫 ────────────────────────────────────────────────────────────────

var ES_BANK_SHEET_NAME_ = 'es_bank';
var ES_BANK_HEADERS_    = ['id','userId','company','industry','question','esText','result','memo','createdAt','updatedAt'];

function writeES_(payload) {
  var session  = getActiveSessionOrThrow_(payload.sessionToken);
  var es       = payload.es || {};
  var id       = trimAuthText_(es.id);
  var company  = trimAuthText_(es.company);
  var industry = trimAuthText_(es.industry);
  var question = trimAuthText_(es.question);
  var esText   = trimAuthText_(es.esText);
  var result   = trimAuthText_(es.result);
  var memo     = trimAuthText_(es.memo);

  assertAuth_(company,  '会社名を入力してください。');
  assertAuth_(esText,   'ES内容を入力してください。');

  var sheet = getOrCreateEsSheet_();
  var now   = new Date().toISOString();

  if (id) {
    var rows = getSheetRecords_(sheet);
    var idx  = findRecordIndex_(rows, function(r) { return String(r.id) === id && String(r.userId) === session.userId; });
    assertAuth_(idx >= 0, 'ESが見つかりません。');
    sheet.getRange(idx + 2, 3, 1, 8).setValues([[company, industry, question, esText, result, memo, rows[idx].createdAt, now]]);
    return { status: 'ok', id: id };
  }

  var newId = 'es_' + new Date().getTime().toString(36) + Utilities.getUuid().replace(/-/g, '').slice(0, 6);
  sheet.appendRow([newId, session.userId, company, industry, question, esText, result, memo, now, now]);
  return { status: 'ok', id: newId };
}

function readMyES_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var sheet   = getOrCreateEsSheet_();
  var rows    = getSheetRecords_(sheet);
  var mine    = rows.filter(function(r) { return String(r.userId) === session.userId; });

  return {
    status: 'ok',
    entries: mine.map(function(r) {
      return {
        id:        trimAuthText_(r.id),
        company:   trimAuthText_(r.company),
        industry:  trimAuthText_(r.industry),
        question:  trimAuthText_(r.question),
        esText:    trimAuthText_(r.esText),
        result:    trimAuthText_(r.result),
        memo:      trimAuthText_(r.memo),
        createdAt: trimAuthText_(r.createdAt),
        updatedAt: trimAuthText_(r.updatedAt)
      };
    })
  };
}

function deleteES_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var id      = trimAuthText_(payload.id);
  assertAuth_(id, 'IDが指定されていません。');

  var sheet = getOrCreateEsSheet_();
  var rows  = getSheetRecords_(sheet);
  var idx   = findRecordIndex_(rows, function(r) { return String(r.id) === id && String(r.userId) === session.userId; });
  assertAuth_(idx >= 0, 'ESが見つかりません。');
  sheet.deleteRow(idx + 2);
  return { status: 'ok' };
}

function getOrCreateEsSheet_() {
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(ES_BANK_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(ES_BANK_SHEET_NAME_);
    sheet.appendRow(ES_BANK_HEADERS_);
  }
  return sheet;
}

// ── ガクチカ保管庫 ────────────────────────────────────────────────────────────

var GAKUCHIKA_SHEET_NAME_ = 'gakuchika_bank';
var GAKUCHIKA_HEADERS_    = ['id','userId','title','category','period','situation','task','action','result','appeal','createdAt','updatedAt'];

function writeGakuchika_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var g       = payload.gakuchika || {};
  var sheet   = getOrCreateGakuchikaSheet_();
  var now     = new Date().toISOString();
  var id      = trimAuthText_(g.id);

  if (id) {
    // 更新
    var data  = sheet.getDataRange().getValues();
    var hdr   = data[0];
    var idCol = hdr.indexOf('id');
    var uidCol = hdr.indexOf('userId');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === id && String(data[i][uidCol]) === session.userId) {
        GAKUCHIKA_HEADERS_.forEach(function(h, c) {
          if (h === 'id' || h === 'userId' || h === 'createdAt') return;
          if (h === 'updatedAt') { sheet.getRange(i + 1, c + 1).setValue(now); return; }
          if (g[h] !== undefined) sheet.getRange(i + 1, c + 1).setValue(trimAuthText_(g[h]));
        });
        return { status: 'ok', id: id };
      }
    }
    throw new Error('該当エントリが見つかりません。');
  } else {
    // 新規
    var newId = 'g_' + now.replace(/\D/g, '').slice(0, 14) + '_' + Math.floor(Math.random() * 10000);
    var row = GAKUCHIKA_HEADERS_.map(function(h) {
      if (h === 'id')        return newId;
      if (h === 'userId')    return session.userId;
      if (h === 'createdAt') return now;
      if (h === 'updatedAt') return now;
      return trimAuthText_(g[h] || '');
    });
    sheet.appendRow(row);
    return { status: 'ok', id: newId };
  }
}

function readMyGakuchika_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var sheet   = getOrCreateGakuchikaSheet_();
  var data    = sheet.getDataRange().getValues();
  var hdr     = data[0];
  var uidCol  = hdr.indexOf('userId');

  var entries = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uidCol]) !== session.userId) continue;
    var entry = {};
    hdr.forEach(function(h, c) { entry[h] = String(data[i][c] || ''); });
    entries.push(entry);
  }
  return { status: 'ok', entries: entries };
}

function deleteGakuchika_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var id      = trimAuthText_(payload.id);
  if (!id) throw new Error('IDが指定されていません。');

  var sheet  = getOrCreateGakuchikaSheet_();
  var data   = sheet.getDataRange().getValues();
  var hdr    = data[0];
  var idCol  = hdr.indexOf('id');
  var uidCol = hdr.indexOf('userId');

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idCol]) === id && String(data[i][uidCol]) === session.userId) {
      sheet.deleteRow(i + 1);
      return { status: 'ok' };
    }
  }
  throw new Error('該当エントリが見つかりません。');
}

function getOrCreateGakuchikaSheet_() {
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(GAKUCHIKA_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(GAKUCHIKA_SHEET_NAME_);
    sheet.appendRow(GAKUCHIKA_HEADERS_);
  }
  return sheet;
}

// ── 面接対策ノート ────────────────────────────────────────────────────────────

var IP_COMPANIES_SHEET_ = 'ip_companies';
var IP_QUESTIONS_SHEET_ = 'ip_questions';
var IP_CO_HEADERS_ = ['id','userId','name','industry','position','status','createdAt','updatedAt'];
var IP_Q_HEADERS_  = ['id','companyId','userId','type','category','question','answer','aiFeedback','ready','createdAt','updatedAt'];

var IP_STD_QUESTIONS_ = [
  {category:'定番', question:'自己PRをしてください（1分程度）'},
  {category:'定番', question:'ガクチカを教えてください（1分程度）'},
  {category:'定番', question:'志望動機を教えてください'},
  {category:'定番', question:'当社でやりたいことは何ですか？'},
  {category:'定番', question:'強みと弱みを教えてください'},
  {category:'定番', question:'挫折経験・失敗経験を教えてください'},
  {category:'定番', question:'10年後のキャリアはどのように考えていますか？'},
  {category:'定番', question:'最近気になったニュースはありますか？'},
  {category:'定番', question:'当社への逆質問はありますか？（2〜3個）'}
];

function getOrCreateIPSheet_(name, headers) {
  var ss = getAuthSpreadsheet_();
  var s  = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(headers); }
  return s;
}

function ipSheetToObjects_(sheet, headers) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = String(row[i] !== undefined ? row[i] : ''); });
    return obj;
  });
}

function readMyIPData_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var uid = session.userId;
  var coSheet = getOrCreateIPSheet_(IP_COMPANIES_SHEET_, IP_CO_HEADERS_);
  var qSheet  = getOrCreateIPSheet_(IP_QUESTIONS_SHEET_, IP_Q_HEADERS_);
  var companies = ipSheetToObjects_(coSheet, IP_CO_HEADERS_).filter(function(c) { return c.userId === uid; });
  var questions  = ipSheetToObjects_(qSheet,  IP_Q_HEADERS_).filter(function(q)  { return q.userId === uid; });
  questions.forEach(function(q) { q.ready = q.ready === 'true' || q.ready === true; });
  return { status: 'ok', companies: companies, questions: questions };
}

function writeIPCompany_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var c   = payload.company || {};
  var now = new Date().toISOString();
  var id  = trimAuthText_(c.id);
  var coSheet = getOrCreateIPSheet_(IP_COMPANIES_SHEET_, IP_CO_HEADERS_);

  if (id) {
    // 更新
    var data   = coSheet.getDataRange().getValues();
    var idCol  = 0; var uidCol = 1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === id && String(data[i][uidCol]) === session.userId) {
        coSheet.getRange(i+1,3).setValue(trimAuthText_(c.name     || ''));
        coSheet.getRange(i+1,4).setValue(trimAuthText_(c.industry || ''));
        coSheet.getRange(i+1,5).setValue(trimAuthText_(c.position || ''));
        coSheet.getRange(i+1,6).setValue(trimAuthText_(c.status   || ''));
        coSheet.getRange(i+1,8).setValue(now);
        return { status: 'ok', id: id };
      }
    }
    throw new Error('企業データが見つかりません。');
  } else {
    // 新規作成
    var newId = 'ipc_' + now.replace(/\D/g,'').slice(0,14) + '_' + Math.floor(Math.random()*10000);
    coSheet.appendRow([newId, session.userId,
      trimAuthText_(c.name||''), trimAuthText_(c.industry||''),
      trimAuthText_(c.position||''), trimAuthText_(c.status||''), now, now]);
    // 定番質問を自動挿入
    var qSheet = getOrCreateIPSheet_(IP_QUESTIONS_SHEET_, IP_Q_HEADERS_);
    IP_STD_QUESTIONS_.forEach(function(sq) {
      var qId = 'ipq_' + now.replace(/\D/g,'').slice(0,14) + '_' + Math.floor(Math.random()*10000);
      qSheet.appendRow([qId, newId, session.userId, 'standard', sq.category, sq.question, '', '', 'false', now, now]);
    });
    return { status: 'ok', id: newId };
  }
}

function deleteIPCompany_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var id = trimAuthText_(payload.id);
  if (!id) throw new Error('IDが指定されていません。');

  var coSheet = getOrCreateIPSheet_(IP_COMPANIES_SHEET_, IP_CO_HEADERS_);
  var coData  = coSheet.getDataRange().getValues();
  for (var i = coData.length - 1; i >= 1; i--) {
    if (String(coData[i][0]) === id && String(coData[i][1]) === session.userId) {
      coSheet.deleteRow(i+1); break;
    }
  }
  // 関連質問を全削除
  var qSheet = getOrCreateIPSheet_(IP_QUESTIONS_SHEET_, IP_Q_HEADERS_);
  var qData  = qSheet.getDataRange().getValues();
  for (var j = qData.length - 1; j >= 1; j--) {
    if (String(qData[j][1]) === id && String(qData[j][2]) === session.userId) qSheet.deleteRow(j+1);
  }
  return { status: 'ok' };
}

function writeIPQuestion_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var q   = payload.question || {};
  var now = new Date().toISOString();
  var id  = trimAuthText_(q.id);
  var qSheet = getOrCreateIPSheet_(IP_QUESTIONS_SHEET_, IP_Q_HEADERS_);

  if (id) {
    // 更新
    var data = qSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id && String(data[i][2]) === session.userId) {
        if (q.answer    !== undefined) qSheet.getRange(i+1,7).setValue(trimAuthText_(q.answer));
        if (q.aiFeedback!== undefined) qSheet.getRange(i+1,8).setValue(trimAuthText_(q.aiFeedback));
        if (q.ready     !== undefined) qSheet.getRange(i+1,9).setValue(String(q.ready));
        if (q.question  !== undefined) qSheet.getRange(i+1,6).setValue(trimAuthText_(q.question));
        qSheet.getRange(i+1,11).setValue(now);
        return { status: 'ok', id: id };
      }
    }
    throw new Error('質問データが見つかりません。');
  } else {
    // 新規
    var newId = 'ipq_' + now.replace(/\D/g,'').slice(0,14) + '_' + Math.floor(Math.random()*10000);
    qSheet.appendRow([newId, trimAuthText_(q.companyId||''), session.userId,
      trimAuthText_(q.type||'custom'), trimAuthText_(q.category||'カスタム'),
      trimAuthText_(q.question||''), trimAuthText_(q.answer||''),
      trimAuthText_(q.aiFeedback||''), 'false', now, now]);
    return { status: 'ok', id: newId };
  }
}

function deleteIPQuestion_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var id = trimAuthText_(payload.id);
  if (!id) throw new Error('IDが指定されていません。');
  var qSheet = getOrCreateIPSheet_(IP_QUESTIONS_SHEET_, IP_Q_HEADERS_);
  var data   = qSheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === id && String(data[i][2]) === session.userId) {
      qSheet.deleteRow(i+1);
      return { status: 'ok' };
    }
  }
  throw new Error('質問データが見つかりません。');
}

function replaceIPGakuchikaQuestions_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var companyId = trimAuthText_(payload.companyId);
  var newQuestions = Array.isArray(payload.questions) ? payload.questions : [];
  if (!companyId) throw new Error('companyIdが指定されていません。');

  var qSheet = getOrCreateIPSheet_(IP_QUESTIONS_SHEET_, IP_Q_HEADERS_);
  var data   = qSheet.getDataRange().getValues();
  // 既存のガクチカ質問を削除
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === companyId && String(data[i][2]) === session.userId && String(data[i][3]) === 'gakuchika') {
      qSheet.deleteRow(i+1);
    }
  }
  // 新規質問を一括挿入
  var now = new Date().toISOString();
  var ids = [];
  newQuestions.forEach(function(q) {
    var newId = 'ipq_' + now.replace(/\D/g,'').slice(0,14) + '_' + Math.floor(Math.random()*10000);
    qSheet.appendRow([newId, companyId, session.userId, 'gakuchika',
      trimAuthText_(q.category||'ガクチカ深掘り'), trimAuthText_(q.question||''),
      '', '', 'false', now, now]);
    ids.push(newId);
  });
  return { status: 'ok', ids: ids };
}

// ── パスワードリセット ─────────────────────────────────────────────────────

var RESET_SHEET_NAME_ = 'auth_password_resets';
var RESET_TOKEN_TTL_MINUTES_ = 30;

function authRequestPasswordReset_(payload) {
  ensureAuthSheets_();
  var email = trimAuthText_(payload.email);
  assertAuth_(email, 'メールアドレスを入力してください。');
  var emailKey = normalizeAuthKey_(email);
  assertAuthRateLimit_('reset', emailKey, AUTH_RESET_RATE_LIMIT_);
  recordAuthRateLimitFailure_('reset', emailKey, AUTH_RESET_RATE_LIMIT_);

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var user = users.find(function (row) {
    return normalizeAuthKey_(row.emailKey || row.email || row.usernameKey || row.username) === emailKey;
  });

  // ユーザーが存在しなくてもセキュリティ上同じレスポンスを返す
  if (!user) {
    return { status: 'ok', message: '登録されているメールアドレスであれば、リセット用のメールを送信しました。' };
  }

  var spreadsheet = getAuthSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(RESET_SHEET_NAME_);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(RESET_SHEET_NAME_);
    sheet.appendRow(['resetToken', 'userId', 'createdAt', 'expiresAt', 'used']);
  }

  var resetToken = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var now = new Date();
  var expiresAt = new Date(now.getTime() + RESET_TOKEN_TTL_MINUTES_ * 60 * 1000);

  sheet.appendRow([resetToken, user.id, now.toISOString(), expiresAt.toISOString(), 'false']);

  // メール送信
  var resetUrl = 'https://naotama1123.github.io/keio-syukatu-navi/account.html?mode=reset&token=' + resetToken;
  var userEmail = trimAuthText_(user.email || user.username);
  try {
    MailApp.sendEmail({
      to: userEmail,
      subject: '【慶應就活ナビ】パスワードリセット',
      htmlBody: '<div style="font-family:sans-serif;max-width:480px;margin:0 auto;padding:24px;">'
        + '<h2 style="color:#0a1a3e;">パスワードリセット</h2>'
        + '<p>以下のリンクをクリックして、新しいパスワードを設定してください。</p>'
        + '<p style="margin:24px 0;"><a href="' + resetUrl + '" style="display:inline-block;padding:12px 28px;background:#0a1a3e;color:#fff;text-decoration:none;font-weight:bold;border-radius:4px;">パスワードをリセットする</a></p>'
        + '<p style="font-size:0.85em;color:#666;">このリンクは' + RESET_TOKEN_TTL_MINUTES_ + '分間有効です。心当たりがない場合はこのメールを無視してください。</p>'
        + '</div>'
    });
  } catch (e) {
    // メール送信失敗してもトークンは作成済み
  }

  return { status: 'ok', message: '登録されているメールアドレスであれば、リセット用のメールを送信しました。' };
}

function authResetPassword_(payload) {
  var resetToken = trimAuthText_(payload.resetToken);
  var newPassword = trimAuthText_(payload.newPassword);
  assertAuth_(resetToken, 'リセットトークンが必要です。');
  assertAuth_(newPassword.length >= 8, '新しいパスワードは8文字以上で設定してください。');

  var spreadsheet = getAuthSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(RESET_SHEET_NAME_);
  assertAuth_(sheet, 'パスワードリセットの情報が見つかりません。');

  var rows = getSheetRecords_(sheet);
  var now = new Date();
  var tokenIndex = findRecordIndex_(rows, function (row) {
    return row.resetToken === resetToken && String(row.used) !== 'true';
  });
  assertAuth_(tokenIndex >= 0, 'リセットリンクが無効または期限切れです。');

  var tokenRecord = rows[tokenIndex];
  var expiresAt = new Date(tokenRecord.expiresAt);
  assertAuth_(!isNaN(expiresAt.getTime()) && expiresAt.getTime() > now.getTime(), 'リセットリンクの有効期限が切れています。');

  // トークンを使用済みにする
  sheet.getRange(tokenIndex + 2, 5).setValue('true');

  // パスワード更新
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var userIndex = findRecordIndex_(users, function (row) {
    return row.id === tokenRecord.userId;
  });
  assertAuth_(userIndex >= 0, 'アカウントが見つかりません。');

  var user = users[userIndex];
  var newHash = sha256Hex_(String(user.salt || '') + ':' + newPassword);
  usersSheet.getRange(userIndex + 2, 5).setValue(newHash);
  usersSheet.getRange(userIndex + 2, 8).setValue(now.toISOString());

  // 既存セッションを全て無効化
  var sessionsSheet = getAuthSheet_(AUTH_SHEET_NAMES_.sessions);
  var sessions = getSheetRecords_(sessionsSheet);
  sessions.forEach(function (row, i) {
    if (row.userId === tokenRecord.userId && String(row.active) !== '0') {
      sessionsSheet.getRange(i + 2, 6).setValue('0');
    }
  });

  return { status: 'ok', message: 'パスワードをリセットしました。新しいパスワードでログインしてください。' };
}

// ── 紹介プログラム ─────────────────────────────────────────────────────────

var REFERRAL_SHEET_NAME_ = 'auth_referrals';

function processReferral_(referralCode, newUserId, users) {
  var referrer = users.find(function (row) {
    return trimAuthText_(row.referralCode) === referralCode;
  });
  if (!referrer) return;

  var spreadsheet = getAuthSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(REFERRAL_SHEET_NAME_);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(REFERRAL_SHEET_NAME_);
    sheet.appendRow(['referralCode', 'referrerUserId', 'referredUserId', 'createdAt']);
  }

  sheet.appendRow([referralCode, referrer.id, newUserId, new Date().toISOString()]);
}

function authGetReferralInfo_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var user = users.find(function (row) { return row.id === session.userId; });
  assertAuth_(user, 'アカウントが見つかりません。');

  var referralCode = trimAuthText_(user.referralCode);

  var spreadsheet = getAuthSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(REFERRAL_SHEET_NAME_);
  var count = 0;
  if (sheet) {
    var rows = getSheetRecords_(sheet);
    count = rows.filter(function (row) {
      return row.referrerUserId === session.userId;
    }).length;
  }

  return {
    status: 'ok',
    referralCode: referralCode,
    referralCount: count
  };
}

// ── 就活進捗トラッカー ─────────────────────────────────────────────────────

var PROGRESS_SHEET_NAME_ = 'progress_tracker';
var PROGRESS_HEADERS_ = ['id', 'userId', 'company', 'industry', 'deadline', 'stage', 'status', 'memo', 'createdAt', 'updatedAt'];

function ensureProgressSheet_() {
  ensureSheet_(PROGRESS_SHEET_NAME_, PROGRESS_HEADERS_);
}

function readMyProgress_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  ensureProgressSheet_();

  var sheet = getAuthSheet_(PROGRESS_SHEET_NAME_);
  var rows = getSheetRecords_(sheet);
  var myRows = rows.filter(function (row) {
    return row.userId === session.userId;
  });

  return {
    status: 'ok',
    entries: myRows.map(function (row) {
      return {
        id: trimAuthText_(row.id),
        company: trimAuthText_(row.company),
        industry: trimAuthText_(row.industry),
        deadline: trimAuthText_(row.deadline),
        stage: trimAuthText_(row.stage),
        status: trimAuthText_(row.status),
        memo: trimAuthText_(row.memo),
        createdAt: trimAuthText_(row.createdAt),
        updatedAt: trimAuthText_(row.updatedAt)
      };
    })
  };
}

function writeProgress_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  ensureProgressSheet_();

  var sheet = getAuthSheet_(PROGRESS_SHEET_NAME_);
  var id = trimAuthText_(payload.id);
  var company = trimAuthText_(payload.company);
  assertAuth_(company, '企業名を入力してください。');

  var now = new Date().toISOString();
  var rowData = {
    company: company,
    industry: trimAuthText_(payload.industry),
    deadline: trimAuthText_(payload.deadline),
    stage: trimAuthText_(payload.stage) || '未応募',
    status: trimAuthText_(payload.status) || 'active',
    memo: trimAuthText_(payload.memo)
  };

  if (id) {
    var rows = getSheetRecords_(sheet);
    var idx = findRecordIndex_(rows, function (row) {
      return row.id === id && row.userId === session.userId;
    });
    if (idx >= 0) {
      sheet.getRange(idx + 2, 3, 1, 6).setValues([[
        rowData.company, rowData.industry, rowData.deadline,
        rowData.stage, rowData.status, rowData.memo
      ]]);
      sheet.getRange(idx + 2, 10).setValue(now);
      return { status: 'ok', id: id };
    }
  }

  var newId = 'prog_' + new Date().getTime().toString(36) + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 6);
  sheet.appendRow([
    newId, session.userId, rowData.company, rowData.industry, rowData.deadline,
    rowData.stage, rowData.status, rowData.memo, now, now
  ]);

  return { status: 'ok', id: newId };
}

function deleteProgress_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  ensureProgressSheet_();

  var id = trimAuthText_(payload.id);
  assertAuth_(id, 'IDが必要です。');

  var sheet = getAuthSheet_(PROGRESS_SHEET_NAME_);
  var rows = getSheetRecords_(sheet);
  var idx = findRecordIndex_(rows, function (row) {
    return row.id === id && row.userId === session.userId;
  });
  assertAuth_(idx >= 0, 'エントリーが見つかりません。');

  sheet.deleteRow(idx + 2);
  return { status: 'ok' };
}

// ── メンバーマッチング・グループ ───────────────────────────────────────────

var GROUPS_SHEET_NAME_ = 'member_groups';
var GROUP_MEMBERS_SHEET_NAME_ = 'member_group_members';
var MATCHING_PREFS_SHEET_NAME_ = 'member_matching_prefs';

function ensureMatchingSheets_() {
  ensureSheet_(GROUPS_SHEET_NAME_, ['id', 'name', 'type', 'industry', 'company', 'description', 'createdBy', 'createdAt']);
  ensureSheet_(GROUP_MEMBERS_SHEET_NAME_, ['groupId', 'userId', 'joinedAt']);
  ensureSheet_(MATCHING_PREFS_SHEET_NAME_, ['userId', 'optIn', 'updatedAt']);
}

function searchMembers_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  ensureMatchingSheets_();

  var filterIndustry = trimAuthText_(payload.industry);
  var filterCompany = trimAuthText_(payload.company);

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);

  var prefsSheet = getAuthSheet_(MATCHING_PREFS_SHEET_NAME_);
  var prefs = getSheetRecords_(prefsSheet);
  var optInMap = {};
  prefs.forEach(function (row) {
    if (String(row.optIn) === 'true') optInMap[row.userId] = true;
  });

  var results = [];
  users.forEach(function (user) {
    if (user.id === session.userId) return;
    if (!optInMap[user.id]) return;

    var match = false;
    if (filterIndustry && normalizeAuthKey_(trimAuthText_(user.desiredIndustry)) === normalizeAuthKey_(filterIndustry)) {
      match = true;
    }
    if (filterCompany) {
      var companies = [user.preferredCompany1, user.preferredCompany2, user.preferredCompany3].map(function (c) { return normalizeAuthKey_(trimAuthText_(c)); });
      if (companies.indexOf(normalizeAuthKey_(filterCompany)) >= 0) match = true;
    }
    if (!filterIndustry && !filterCompany) match = true;

    if (match) {
      results.push({
        id: user.id,
        displayName: trimAuthText_(user.displayName || user.username),
        desiredIndustry: trimAuthText_(user.desiredIndustry),
        preferredCompanies: getPreferredCompanies_(user),
        lineName: trimAuthText_(user.lineName),
        lineQrUrl: trimAuthText_(user.lineQrDriveUrl)
      });
    }
  });

  return { status: 'ok', members: results.slice(0, 50) };
}

function getGroups_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  ensureMatchingSheets_();

  var groupsSheet = getAuthSheet_(GROUPS_SHEET_NAME_);
  var groups = getSheetRecords_(groupsSheet);
  var membersSheet = getAuthSheet_(GROUP_MEMBERS_SHEET_NAME_);
  var memberships = getSheetRecords_(membersSheet);

  var myGroupIds = {};
  memberships.forEach(function (m) {
    if (m.userId === session.userId) myGroupIds[m.groupId] = true;
  });

  var memberCounts = {};
  memberships.forEach(function (m) {
    memberCounts[m.groupId] = (memberCounts[m.groupId] || 0) + 1;
  });

  return {
    status: 'ok',
    groups: groups.map(function (g) {
      return {
        id: g.id,
        name: trimAuthText_(g.name),
        type: trimAuthText_(g.type),
        industry: trimAuthText_(g.industry),
        company: trimAuthText_(g.company),
        description: trimAuthText_(g.description),
        memberCount: memberCounts[g.id] || 0,
        isMember: !!myGroupIds[g.id]
      };
    })
  };
}

function createGroup_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  ensureMatchingSheets_();

  var name = trimAuthText_(payload.name);
  var type = trimAuthText_(payload.type) || 'industry';
  assertAuth_(name, 'グループ名を入力してください。');

  var groupsSheet = getAuthSheet_(GROUPS_SHEET_NAME_);
  var now = new Date().toISOString();
  var groupId = 'grp_' + new Date().getTime().toString(36) + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 6);

  groupsSheet.appendRow([
    groupId, name, type,
    trimAuthText_(payload.industry),
    trimAuthText_(payload.company),
    trimAuthText_(payload.description),
    session.userId, now
  ]);

  // 作成者を自動参加
  var membersSheet = getAuthSheet_(GROUP_MEMBERS_SHEET_NAME_);
  membersSheet.appendRow([groupId, session.userId, now]);

  return { status: 'ok', groupId: groupId };
}

function joinGroup_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  ensureMatchingSheets_();

  var groupId = trimAuthText_(payload.groupId);
  assertAuth_(groupId, 'グループIDが必要です。');

  var membersSheet = getAuthSheet_(GROUP_MEMBERS_SHEET_NAME_);
  var memberships = getSheetRecords_(membersSheet);
  var already = memberships.find(function (m) {
    return m.groupId === groupId && m.userId === session.userId;
  });
  if (already) return { status: 'ok' };

  membersSheet.appendRow([groupId, session.userId, new Date().toISOString()]);
  return { status: 'ok' };
}

function leaveGroup_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  ensureMatchingSheets_();

  var groupId = trimAuthText_(payload.groupId);
  assertAuth_(groupId, 'グループIDが必要です。');

  var membersSheet = getAuthSheet_(GROUP_MEMBERS_SHEET_NAME_);
  var memberships = getSheetRecords_(membersSheet);
  var idx = findRecordIndex_(memberships, function (m) {
    return m.groupId === groupId && m.userId === session.userId;
  });
  if (idx >= 0) membersSheet.deleteRow(idx + 2);

  return { status: 'ok' };
}

function getGroupMembers_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  ensureMatchingSheets_();

  var groupId = trimAuthText_(payload.groupId);
  assertAuth_(groupId, 'グループIDが必要です。');

  var membersSheet = getAuthSheet_(GROUP_MEMBERS_SHEET_NAME_);
  var memberships = getSheetRecords_(membersSheet);
  var memberUserIds = memberships.filter(function (m) {
    return m.groupId === groupId;
  }).map(function (m) { return m.userId; });

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);

  var prefsSheet = getAuthSheet_(MATCHING_PREFS_SHEET_NAME_);
  var prefs = getSheetRecords_(prefsSheet);
  var optInMap = {};
  prefs.forEach(function (row) {
    if (String(row.optIn) === 'true') optInMap[row.userId] = true;
  });

  var members = [];
  memberUserIds.forEach(function (uid) {
    var user = users.find(function (u) { return u.id === uid; });
    if (!user) return;
    members.push({
      id: user.id,
      displayName: trimAuthText_(user.displayName || user.username),
      desiredIndustry: trimAuthText_(user.desiredIndustry),
      preferredCompanies: getPreferredCompanies_(user),
      lineName: optInMap[uid] ? trimAuthText_(user.lineName) : '',
      lineQrUrl: optInMap[uid] ? trimAuthText_(user.lineQrDriveUrl) : '',
      isMe: uid === session.userId
    });
  });

  return { status: 'ok', members: members };
}

function updateMatchingPrefs_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  ensureMatchingSheets_();

  var optIn = payload.optIn === true || payload.optIn === 'true';
  var prefsSheet = getAuthSheet_(MATCHING_PREFS_SHEET_NAME_);
  var prefs = getSheetRecords_(prefsSheet);
  var idx = findRecordIndex_(prefs, function (row) {
    return row.userId === session.userId;
  });

  var now = new Date().toISOString();
  if (idx >= 0) {
    prefsSheet.getRange(idx + 2, 2, 1, 2).setValues([[String(optIn), now]]);
  } else {
    prefsSheet.appendRow([session.userId, String(optIn), now]);
  }

  return { status: 'ok', optIn: optIn };
}

/* ===== 週次ダイジェストメール ===== */

function sendWeeklyDigest_(payload) {
  requireAdminSession_(payload);
  var result = sendWeeklyDigest();
  return { status: 'ok', sent: result };
}

/**
 * 週次ダイジェストメールを全ユーザーに送信する。
 * GAS時間ベーストリガー（毎週）で sendWeeklyDigest() を呼び出してください。
 */
function sendWeeklyDigest() {
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var sevenDaysAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString();
  var sentCount = 0;

  var newBoardPosts = countRecentRows_('es_board', sevenDaysAgo);
  var newConsultations = countRecentRows_('consultations', sevenDaysAgo);
  var activeTimeline = countRecentRows_('timeline', sevenDaysAgo);

  users.forEach(function(user) {
    var email = trimAuthText_(user.email);
    if (!email || !validateEmail_(email)) return;

    var displayName = trimAuthText_(user.displayName) || 'メンバー';
    var siteUrl = 'https://naotama1123.github.io/keio-syukatu-navi/public/members.html';

    var htmlBody = '\x3c!DOCTYPE html\x3e\x3chtml\x3e\x3chead\x3e\x3cmeta charset="UTF-8"\x3e\x3c/head\x3e\x3cbody style="font-family:sans-serif;background:#faf7f2;padding:20px;"\x3e'
      + '\x3cdiv style="max-width:500px;margin:0 auto;background:white;border-radius:8px;overflow:hidden;"\x3e'
      + '\x3cdiv style="background:#0a1a3e;padding:20px;text-align:center;"\x3e'
      + '\x3ch1 style="color:#c9a84c;font-size:18px;margin:0;"\x3e今週の就活ナビ\x3c/h1\x3e'
      + '\x3c/div\x3e'
      + '\x3cdiv style="padding:20px;"\x3e'
      + '\x3cp style="font-size:14px;color:#1a1a2e;"\x3e' + displayName + ' さん、こんにちは！\x3c/p\x3e'
      + '\x3cp style="font-size:13px;color:#6b6b7e;"\x3e今週のコミュニティの動きをお届けします。\x3c/p\x3e'
      + '\x3cdiv style="background:#faf7f2;padding:15px;border-radius:6px;margin:15px 0;"\x3e'
      + '\x3cp style="font-size:13px;margin:6px 0;"\x3eES掲示板の新規投稿: \x3cstrong\x3e' + newBoardPosts + '件\x3c/strong\x3e\x3c/p\x3e'
      + '\x3cp style="font-size:13px;margin:6px 0;"\x3e就活相談の新規投稿: \x3cstrong\x3e' + newConsultations + '件\x3c/strong\x3e\x3c/p\x3e'
      + '\x3cp style="font-size:13px;margin:6px 0;"\x3eタイムライン投稿: \x3cstrong\x3e' + activeTimeline + '件\x3c/strong\x3e\x3c/p\x3e'
      + '\x3c/div\x3e'
      + '\x3cdiv style="text-align:center;margin:20px 0;"\x3e'
      + '\x3ca href="' + siteUrl + '" style="display:inline-block;background:#0a1a3e;color:#c9a84c;padding:12px 30px;text-decoration:none;font-weight:bold;font-size:14px;border-radius:4px;"\x3e今すぐサイトを開く\x3c/a\x3e'
      + '\x3c/div\x3e'
      + '\x3c/div\x3e'
      + '\x3c/div\x3e'
      + '\x3c/body\x3e\x3c/html\x3e';

    try {
      MailApp.sendEmail({
        to: email,
        subject: '今週の就活ナビ',
        htmlBody: htmlBody
      });
      sentCount++;
    } catch (e) {
      // メール送信エラーはスキップ
    }
  });

  return sentCount;
}

function countRecentRows_(sheetName, sinceIso) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  var rows = getSheetRecords_(sheet);
  var count = 0;
  rows.forEach(function(row) {
    var createdAt = row.createdAt || row.timestamp || '';
    if (createdAt >= sinceIso) count++;
  });
  return count;
}

// ── 質問・相談掲示板 (Q&A Board) ──────────────────────────────────────────────

var QA_BOARD_SHEET_NAME_    = 'qa_board';
var QA_BOARD_MAX_POSTS_     = 100;
var QA_BOARD_CATEGORIES_    = ['ES・ガクチカ','面接','企業研究','その他'];
var QA_BOARD_HEADERS_       = ['id','userId','displayName','title','body','category','anonymous','createdAt','replies'];

function getOrCreateQABoardSheet_() {
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(QA_BOARD_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(QA_BOARD_SHEET_NAME_);
    sheet.appendRow(QA_BOARD_HEADERS_);
  }
  return sheet;
}

function postQuestion_(payload) {
  var session  = getActiveSessionOrThrow_(payload.sessionToken);
  var title    = trimAuthText_(payload.title);
  var body     = trimAuthText_(payload.body);
  var category = trimAuthText_(payload.category);
  var anonymous = payload.anonymous === true || payload.anonymous === 'true';

  assertAuth_(title, 'タイトルを入力してください。');
  assertAuth_(body,  '質問内容を入力してください。');
  if (QA_BOARD_CATEGORIES_.indexOf(category) < 0) category = 'その他';

  var sheet = getOrCreateQABoardSheet_();
  var now   = new Date().toISOString();
  var id    = now.replace(/\D/g, '').slice(0, 14) + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users      = getSheetRecords_(usersSheet);
  var userRecord = users.find(function (row) { return row.id === session.userId; });
  var displayName = userRecord ? trimAuthText_(userRecord.displayName || userRecord.username) : '';

  sheet.appendRow([id, session.userId, displayName, title, body, category, anonymous ? 'true' : 'false', now, '[]']);

  return { status: 'ok', id: id };
}

function readQuestions_(payload) {
  getActiveSessionOrThrow_(payload.sessionToken);

  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(QA_BOARD_SHEET_NAME_);
  if (!sheet) return { status: 'ok', questions: [] };

  var posts = getSheetRecords_(sheet);
  if (!posts.length) return { status: 'ok', questions: [] };

  posts.sort(function (a, b) {
    return new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime();
  });
  posts = posts.slice(0, QA_BOARD_MAX_POSTS_);

  var result = posts.map(function (post) {
    var replies = [];
    try { replies = JSON.parse(post.replies || '[]'); } catch (e) { replies = []; }
    var isAnon = String(post.anonymous) === 'true';
    return {
      id:          trimAuthText_(post.id),
      userId:      trimAuthText_(post.userId),
      displayName: isAnon ? '匿名' : trimAuthText_(post.displayName),
      title:       trimAuthText_(post.title),
      body:        trimAuthText_(post.body),
      category:    trimAuthText_(post.category),
      anonymous:   isAnon,
      createdAt:   trimAuthText_(post.createdAt),
      replyCount:  replies.length,
      replies:     replies
    };
  });

  return { status: 'ok', questions: result };
}

function postReply_(payload) {
  var session    = getActiveSessionOrThrow_(payload.sessionToken);
  var questionId = trimAuthText_(payload.questionId);
  var body       = trimAuthText_(payload.body);

  assertAuth_(questionId, '質問IDが指定されていません。');
  assertAuth_(body,       '返信内容を入力してください。');

  var sheet = getOrCreateQABoardSheet_();
  var rows  = getSheetRecords_(sheet);
  var idx   = findRecordIndex_(rows, function (r) { return r.id === questionId; });
  assertAuth_(idx >= 0, '質問が見つかりませんでした。');

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users      = getSheetRecords_(usersSheet);
  var userRecord = users.find(function (row) { return row.id === session.userId; });
  var displayName = userRecord ? trimAuthText_(userRecord.displayName || userRecord.username) : '';

  var now     = new Date().toISOString();
  var replyId = now.replace(/\D/g, '').slice(0, 14) + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);

  var replies = [];
  try { replies = JSON.parse(rows[idx].replies || '[]'); } catch (e) { replies = []; }
  replies.push({ id: replyId, userId: session.userId, displayName: displayName, body: body, createdAt: now });

  // replies column index (9th column = index 9)
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var repliesCol = headers.indexOf('replies') + 1;
  if (repliesCol < 1) repliesCol = 9;
  sheet.getRange(idx + 2, repliesCol).setValue(JSON.stringify(replies));

  return { status: 'ok' };
}

function deleteQuestion_(payload) {
  var session    = getActiveSessionOrThrow_(payload.sessionToken);
  var questionId = trimAuthText_(payload.questionId);

  assertAuth_(questionId, '質問IDが指定されていません。');

  var sheet = getOrCreateQABoardSheet_();
  var rows  = getSheetRecords_(sheet);
  var idx   = findRecordIndex_(rows, function (r) { return r.id === questionId; });
  assertAuth_(idx >= 0, '質問が見つかりませんでした。');
  assertAuth_(rows[idx].userId === session.userId, '自分の質問のみ削除できます。');

  sheet.deleteRow(idx + 2);

  return { status: 'ok' };
}

/* ==============================
   内定体験記 (Experience Stories)
   ============================== */

var EXP_SHEET_NAME_    = 'experience_stories';
var EXP_MAX_STORIES_   = 200;
var EXP_HEADERS_       = ['id','userId','displayName','industry','companyCount','startMonth','result','duration','keyLearning','advice','tools','anonymous','createdAt'];

function getOrCreateExpSheet_() {
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(EXP_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(EXP_SHEET_NAME_);
    sheet.appendRow(EXP_HEADERS_);
  }
  return sheet;
}

function postExperience_(payload) {
  var session      = getActiveSessionOrThrow_(payload.sessionToken);
  var industry     = trimAuthText_(payload.industry);
  var companyCount = parseInt(payload.companyCount, 10) || 0;
  var startMonth   = trimAuthText_(payload.startMonth);
  var result       = parseInt(payload.result, 10) || 0;
  var duration     = trimAuthText_(payload.duration);
  var keyLearning  = trimAuthText_(payload.keyLearning);
  var advice       = trimAuthText_(payload.advice);
  var tools        = payload.tools;
  var anonymous    = payload.anonymous === true || payload.anonymous === 'true';

  assertAuth_(industry,    '業界を選択してください。');
  assertAuth_(startMonth,  '就活開始時期を選択してください。');
  assertAuth_(duration,    '就活期間を選択してください。');
  assertAuth_(keyLearning, '最も大事だったことを入力してください。');
  assertAuth_(advice,      '後輩へのアドバイスを入力してください。');

  if (!Array.isArray(tools)) tools = [];
  var toolsJson = JSON.stringify(tools);

  var sheet = getOrCreateExpSheet_();
  var now   = new Date().toISOString();
  var id    = now.replace(/\D/g, '').slice(0, 14) + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);

  var usersSheet  = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users       = getSheetRecords_(usersSheet);
  var userRecord  = users.find(function (row) { return row.id === session.userId; });
  var displayName = userRecord ? trimAuthText_(userRecord.displayName || userRecord.username) : '';

  sheet.appendRow([id, session.userId, displayName, industry, companyCount, startMonth, result, duration, keyLearning, advice, toolsJson, anonymous ? 'true' : 'false', now]);

  return { status: 'ok', id: id };
}

function readExperiences_(payload) {
  getActiveSessionOrThrow_(payload.sessionToken);

  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(EXP_SHEET_NAME_);
  if (!sheet) return { status: 'ok', experiences: [] };

  var rows = getSheetRecords_(sheet);
  if (!rows.length) return { status: 'ok', experiences: [] };

  rows.sort(function (a, b) {
    return new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime();
  });
  rows = rows.slice(0, EXP_MAX_STORIES_);

  var result = rows.map(function (row) {
    var toolsParsed = [];
    try { toolsParsed = JSON.parse(row.tools || '[]'); } catch (e) { toolsParsed = []; }
    var isAnon = String(row.anonymous) === 'true';
    return {
      id:           trimAuthText_(row.id),
      userId:       trimAuthText_(row.userId),
      displayName:  isAnon ? '匿名' : trimAuthText_(row.displayName),
      industry:     trimAuthText_(row.industry),
      companyCount: parseInt(row.companyCount, 10) || 0,
      startMonth:   trimAuthText_(row.startMonth),
      result:       parseInt(row.result, 10) || 0,
      duration:     trimAuthText_(row.duration),
      keyLearning:  trimAuthText_(row.keyLearning),
      advice:       trimAuthText_(row.advice),
      tools:        toolsParsed,
      anonymous:    isAnon,
      createdAt:    trimAuthText_(row.createdAt)
    };
  });

  return { status: 'ok', experiences: result };
}

// ── GD練習会フィードバック ─────────────────────────────────────

var GD_FB_SHEET_ = 'gd_feedback';
var GD_FB_HEADERS_ = ['id','userId','date','theme','role','otherFeedback','selfReflection','rating','createdAt','updatedAt'];

function getOrCreateGdFbSheet_() {
  var ssId = PropertiesService.getScriptProperties().getProperty('AUTH_SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName(GD_FB_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(GD_FB_SHEET_);
    sheet.getRange(1, 1, 1, GD_FB_HEADERS_.length).setValues([GD_FB_HEADERS_]);
  }
  return sheet;
}

function writeGdFeedback_(payload) {
  var user = resolveSession_(payload.sessionToken);
  if (!user) throw new Error('ログインが必要です。');
  var entry = payload.entry || {};
  if (!entry.theme) throw new Error('テーマを入力してください。');

  var sheet = getOrCreateGdFbSheet_();
  var now = new Date().toISOString();

  if (entry.id) {
    // 更新
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === entry.id && data[i][1] === user.userId) {
        sheet.getRange(i + 1, 3, 1, 8).setValues([[
          entry.date || '', entry.theme, entry.role || '',
          entry.otherFeedback || '', entry.selfReflection || '',
          entry.rating || 0, data[i][8], now
        ]]);
        return { status: 'ok', id: entry.id };
      }
    }
    throw new Error('エントリーが見つかりません。');
  } else {
    // 新規
    var id = Utilities.getUuid();
    sheet.appendRow([
      id, user.userId, entry.date || '', entry.theme, entry.role || '',
      entry.otherFeedback || '', entry.selfReflection || '',
      entry.rating || 0, now, now
    ]);
    return { status: 'ok', id: id };
  }
}

function readGdFeedback_(payload) {
  var user = resolveSession_(payload.sessionToken);
  if (!user) throw new Error('ログインが必要です。');

  var sheet;
  try { sheet = getOrCreateGdFbSheet_(); } catch(e) { return { status: 'ok', entries: [] }; }
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: 'ok', entries: [] };

  var entries = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] !== user.userId) continue;
    entries.push({
      id: data[i][0], date: data[i][2], theme: data[i][3], role: data[i][4],
      otherFeedback: data[i][5], selfReflection: data[i][6], rating: data[i][7],
      createdAt: data[i][8], updatedAt: data[i][9]
    });
  }
  entries.sort(function(a, b) { return (b.date || '').localeCompare(a.date || ''); });
  return { status: 'ok', entries: entries };
}

function deleteGdFeedback_(payload) {
  var user = resolveSession_(payload.sessionToken);
  if (!user) throw new Error('ログインが必要です。');
  var id = payload.id;
  if (!id) throw new Error('IDが必要です。');

  var sheet = getOrCreateGdFbSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id && data[i][1] === user.userId) {
      sheet.deleteRow(i + 1);
      return { status: 'ok' };
    }
  }
  throw new Error('エントリーが見つかりません。');
}

/* ==============================
   面接練習マッチング (Practice Matching)
   ============================== */

var PM_SHEET_NAME_  = 'practice_matching';
var PM_MAX_ITEMS_   = 200;
var PM_HEADERS_     = ['id','userId','displayName','type','industry','preferredDate','message','maxMembers','participants','status','createdAt'];

function getOrCreatePMSheet_() {
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(PM_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(PM_SHEET_NAME_);
    sheet.appendRow(PM_HEADERS_);
  }
  return sheet;
}

function postPracticeRequest_(payload) {
  var session       = getActiveSessionOrThrow_(payload.sessionToken);
  var type          = trimAuthText_(payload.type);
  var industry      = trimAuthText_(payload.industry);
  var preferredDate = trimAuthText_(payload.preferredDate);
  var message       = trimAuthText_(payload.message);
  var maxMembers    = parseInt(payload.maxMembers, 10) || 2;

  var validTypes = ['面接練習', 'GD練習', 'ケース面接'];
  assertAuth_(validTypes.indexOf(type) >= 0, '練習タイプを選択してください。');
  assertAuth_(preferredDate, '希望日時を入力してください。');
  if (maxMembers < 2) maxMembers = 2;
  if (maxMembers > 6) maxMembers = 6;

  var sheet = getOrCreatePMSheet_();
  var now   = new Date().toISOString();
  var id    = now.replace(/\D/g, '').slice(0, 14) + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);

  var usersSheet  = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users       = getSheetRecords_(usersSheet);
  var userRecord  = users.find(function (row) { return row.id === session.userId; });
  var displayName = userRecord ? trimAuthText_(userRecord.displayName || '') : '';

  sheet.appendRow([id, session.userId, displayName, type, industry, preferredDate, message, maxMembers, '[]', '募集中', now]);

  return { status: 'ok', id: id };
}

function readPracticeRequests_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);

  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(PM_SHEET_NAME_);
  if (!sheet) return { status: 'ok', requests: [] };

  var rows = getSheetRecords_(sheet);
  if (!rows.length) return { status: 'ok', requests: [] };

  // Look up user lineName map for participants
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var allUsers   = getSheetRecords_(usersSheet);
  var userMap    = {};
  allUsers.forEach(function (u) {
    userMap[u.id] = { displayName: trimAuthText_(u.displayName || ''), lineName: trimAuthText_(u.lineName || '') };
  });

  // Sort: 募集中 first, then newest first
  rows.sort(function (a, b) {
    var sa = String(a.status) === '募集中' ? 0 : 1;
    var sb = String(b.status) === '募集中' ? 0 : 1;
    if (sa !== sb) return sa - sb;
    return new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime();
  });
  rows = rows.slice(0, PM_MAX_ITEMS_);

  var result = rows.map(function (row) {
    var participants = [];
    try { participants = JSON.parse(row.participants || '[]'); } catch (e) { participants = []; }

    var isParticipant = participants.some(function (p) { return p.userId === session.userId; });
    var creatorInfo   = userMap[row.userId] || {};

    // Enrich participants with display names
    var enrichedParticipants = participants.map(function (p) {
      var info = userMap[p.userId] || {};
      return { userId: p.userId, displayName: info.displayName || p.displayName || '' };
    });

    return {
      id:            trimAuthText_(row.id),
      userId:        trimAuthText_(row.userId),
      displayName:   trimAuthText_(row.displayName),
      type:          trimAuthText_(row.type),
      industry:      trimAuthText_(row.industry),
      preferredDate: trimAuthText_(row.preferredDate),
      message:       trimAuthText_(row.message),
      maxMembers:    parseInt(row.maxMembers, 10) || 2,
      participants:  enrichedParticipants,
      status:        trimAuthText_(row.status),
      createdAt:     trimAuthText_(row.createdAt),
      isOwner:       row.userId === session.userId,
      isParticipant: isParticipant,
      creatorLineName: isParticipant ? (creatorInfo.lineName || '') : ''
    };
  });

  return { status: 'ok', requests: result };
}

function joinPracticeRequest_(payload) {
  var session   = getActiveSessionOrThrow_(payload.sessionToken);
  var requestId = trimAuthText_(payload.requestId);
  assertAuth_(requestId, '募集IDが指定されていません。');

  var sheet = getOrCreatePMSheet_();
  var rows  = getSheetRecords_(sheet);
  var idx   = findRecordIndex_(rows, function (r) { return r.id === requestId; });
  assertAuth_(idx >= 0, '募集が見つかりませんでした。');

  var row = rows[idx];
  assertAuth_(row.userId !== session.userId, '自分の募集には参加できません。');
  assertAuth_(String(row.status) === '募集中', 'この募集は締め切られています。');

  var participants = [];
  try { participants = JSON.parse(row.participants || '[]'); } catch (e) { participants = []; }

  var alreadyJoined = participants.some(function (p) { return p.userId === session.userId; });
  assertAuth_(!alreadyJoined, 'すでに参加しています。');

  var maxMembers = parseInt(row.maxMembers, 10) || 2;
  assertAuth_(participants.length < maxMembers, '定員に達しています。');

  var usersSheet  = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var allUsers    = getSheetRecords_(usersSheet);
  var userRecord  = allUsers.find(function (u) { return u.id === session.userId; });
  var displayName = userRecord ? trimAuthText_(userRecord.displayName || '') : '';

  participants.push({ userId: session.userId, displayName: displayName, joinedAt: new Date().toISOString() });

  var headers       = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var partCol       = headers.indexOf('participants') + 1;
  var statusCol     = headers.indexOf('status') + 1;
  if (partCol < 1) partCol = 9;
  if (statusCol < 1) statusCol = 10;

  sheet.getRange(idx + 2, partCol).setValue(JSON.stringify(participants));

  // Auto-close if full
  if (participants.length >= maxMembers) {
    sheet.getRange(idx + 2, statusCol).setValue('締切');
  }

  return { status: 'ok' };
}

function leavePracticeRequest_(payload) {
  var session   = getActiveSessionOrThrow_(payload.sessionToken);
  var requestId = trimAuthText_(payload.requestId);
  assertAuth_(requestId, '募集IDが指定されていません。');

  var sheet = getOrCreatePMSheet_();
  var rows  = getSheetRecords_(sheet);
  var idx   = findRecordIndex_(rows, function (r) { return r.id === requestId; });
  assertAuth_(idx >= 0, '募集が見つかりませんでした。');

  var row = rows[idx];
  var participants = [];
  try { participants = JSON.parse(row.participants || '[]'); } catch (e) { participants = []; }

  var newParticipants = participants.filter(function (p) { return p.userId !== session.userId; });
  assertAuth_(newParticipants.length < participants.length, '参加していません。');

  var headers   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var partCol   = headers.indexOf('participants') + 1;
  var statusCol = headers.indexOf('status') + 1;
  if (partCol < 1) partCol = 9;
  if (statusCol < 1) statusCol = 10;

  sheet.getRange(idx + 2, partCol).setValue(JSON.stringify(newParticipants));

  // Re-open if was closed due to capacity and now has room
  var maxMembers = parseInt(row.maxMembers, 10) || 2;
  if (String(row.status) === '締切' && newParticipants.length < maxMembers) {
    sheet.getRange(idx + 2, statusCol).setValue('募集中');
  }

  return { status: 'ok' };
}

function closePracticeRequest_(payload) {
  var session   = getActiveSessionOrThrow_(payload.sessionToken);
  var requestId = trimAuthText_(payload.requestId);
  assertAuth_(requestId, '募集IDが指定されていません。');

  var sheet = getOrCreatePMSheet_();
  var rows  = getSheetRecords_(sheet);
  var idx   = findRecordIndex_(rows, function (r) { return r.id === requestId; });
  assertAuth_(idx >= 0, '募集が見つかりませんでした。');
  assertAuth_(rows[idx].userId === session.userId, '自分の募集のみ締切にできます。');

  var headers   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statusCol = headers.indexOf('status') + 1;
  if (statusCol < 1) statusCol = 10;

  sheet.getRange(idx + 2, statusCol).setValue('締切');

  return { status: 'ok' };
}

function deletePracticeRequest_(payload) {
  var session   = getActiveSessionOrThrow_(payload.sessionToken);
  var requestId = trimAuthText_(payload.requestId);
  assertAuth_(requestId, '募集IDが指定されていません。');

  var sheet = getOrCreatePMSheet_();
  var rows  = getSheetRecords_(sheet);
  var idx   = findRecordIndex_(rows, function (r) { return r.id === requestId; });
  assertAuth_(idx >= 0, '募集が見つかりませんでした。');
  assertAuth_(rows[idx].userId === session.userId, '自分の募集のみ削除できます。');

  sheet.deleteRow(idx + 2);

  return { status: 'ok' };
}
