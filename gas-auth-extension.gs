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
  liked: 'auth_liked',
  emailVerifications: 'auth_email_verifications',
  audit: 'auth_audit_log'
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
  'referralCode',
  'lineNotifyToken',
  'emailVerified',
  'emailVerifiedAt'
];

var AUTH_SESSION_TTL_DAYS_ = 7;
var EMAIL_VERIFICATION_TOKEN_TTL_HOURS_ = 72;
var AUTH_AUDIT_HEADERS_ = ['timestamp', 'action', 'status', 'userId', 'emailKey', 'sessionRef', 'detail'];
var AUTH_EMAIL_VERIFICATION_HEADERS_ = ['verificationToken', 'userId', 'email', 'emailKey', 'createdAt', 'expiresAt', 'used'];

function handleAuthAction_(payload) {
  if (!payload || typeof payload !== 'object' || !String(payload.action || '').match(/^(auth|writeBoardPost|readBoardPosts|addBoardComment|callClaude|writeES|readMyES|deleteES|writeGakuchika|readMyGakuchika|deleteGakuchika|readMyIPData|writeIPCompany|deleteIPCompany|writeIPQuestion|deleteIPQuestion|replaceIPGakuchikaQuestions|readMyProgress|writeProgress|deleteProgress|writeSelfPR|readMySelfPR|deleteSelfPR|searchMembers|getGroups|joinGroup|leaveGroup|createGroup|getGroupMembers|updateMatchingPrefs|writeTimeline|readTimeline|getMyStats|writeInterviewReview|readMyInterviewReviews|deleteInterviewReview|writeConsultation|readConsultations|addConsultationComment|createGDSession|readGDSessions|joinGDSession|leaveGDSession|generateGDTheme|writeOBVisit|readMyOBVisits|deleteOBVisit|shareOBInfo|readSharedOBInfo|readPassedES|submitNPS|readNPSSummary|sendWeeklyDigest|writeSelectionExperience|readSelectionExperiences|getAdminStats|getModerationQueue|resolveModerationReport|getAuthAuditLog|reportContent|readMutedMembers|muteMember|unmuteMember|acceptQuestionReply|authSaveLineNotifyToken|sendWeeklyDigestLine|readQuestions|postQuestion|postReply|deleteQuestion|postExperience|readExperiences|readGdFeedback|writeGdFeedback|deleteGdFeedback|postPracticeRequest|readPracticeRequests|joinPracticeRequest|leavePracticeRequest|closePracticeRequest|deletePracticeRequest)/)) {
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
      case 'authVerifyEmail':
        return jsonResponse_(authVerifyEmail_(payload));
      case 'authResendVerificationEmail':
        return jsonResponse_(authResendVerificationEmail_(payload));
      case 'authRequestPasswordReset':
        return jsonResponse_(authRequestPasswordReset_(payload));
      case 'authResetPassword':
        return jsonResponse_(authResetPassword_(payload));
      case 'authGetReferralInfo':
        return jsonResponse_(authGetReferralInfo_(payload));
      case 'authListSessions':
        return jsonResponse_(authListSessions_(payload));
      case 'authRevokeSession':
        return jsonResponse_(authRevokeSession_(payload));
      case 'authRevokeOtherSessions':
        return jsonResponse_(authRevokeOtherSessions_(payload));
      case 'readMyProgress':  return jsonResponse_(readMyProgress_(payload));
      case 'writeProgress':   return jsonResponse_(writeProgress_(payload));
      case 'deleteProgress':  return jsonResponse_(deleteProgress_(payload));
      case 'writeSelfPR':     return jsonResponse_(writeSelfPR_(payload));
      case 'readMySelfPR':    return jsonResponse_(readMySelfPR_(payload));
      case 'deleteSelfPR':    return jsonResponse_(deleteSelfPR_(payload));
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
      case 'writeTimeline':  return jsonResponse_(writeTimeline_(payload));
      case 'readTimeline':   return jsonResponse_(readTimeline_(payload));
      case 'getMyStats':     return jsonResponse_(getMyStats_(payload));
      case 'writeInterviewReview':    return jsonResponse_(writeInterviewReview_(payload));
      case 'readMyInterviewReviews':  return jsonResponse_(readMyInterviewReviews_(payload));
      case 'deleteInterviewReview':   return jsonResponse_(deleteInterviewReview_(payload));
      case 'writeConsultation':       return jsonResponse_(writeConsultation_(payload));
      case 'readConsultations':       return jsonResponse_(readConsultations_(payload));
      case 'addConsultationComment':  return jsonResponse_(addConsultationComment_(payload));
      case 'createGDSession':   return jsonResponse_(createGDSession_(payload));
      case 'readGDSessions':    return jsonResponse_(readGDSessions_(payload));
      case 'joinGDSession':     return jsonResponse_(joinGDSession_(payload));
      case 'leaveGDSession':    return jsonResponse_(leaveGDSession_(payload));
      case 'generateGDTheme':   return jsonResponse_(generateGDTheme_(payload));
      case 'writeOBVisit':       return jsonResponse_(writeOBVisit_(payload));
      case 'readMyOBVisits':     return jsonResponse_(readMyOBVisits_(payload));
      case 'deleteOBVisit':      return jsonResponse_(deleteOBVisit_(payload));
      case 'shareOBInfo':        return jsonResponse_(shareOBInfo_(payload));
      case 'readSharedOBInfo':   return jsonResponse_(readSharedOBInfo_(payload));
      case 'readPassedES':       return jsonResponse_(readPassedES_(payload));
      case 'submitNPS':          return jsonResponse_(submitNPS_(payload));
      case 'readNPSSummary':     return jsonResponse_(readNPSSummary_(payload));
      case 'sendWeeklyDigest':   return jsonResponse_(sendWeeklyDigest_(payload));
      case 'writeSelectionExperience': return jsonResponse_(writeSelectionExperience_(payload));
      case 'readSelectionExperiences': return jsonResponse_(readSelectionExperiences_(payload));
      case 'getAdminStats':            return jsonResponse_(getAdminStats_(payload));
      case 'getModerationQueue':       return jsonResponse_(getModerationQueue_(payload));
      case 'resolveModerationReport':  return jsonResponse_(resolveModerationReport_(payload));
      case 'getAuthAuditLog':          return jsonResponse_(getAuthAuditLog_(payload));
      case 'reportContent':            return jsonResponse_(reportContent_(payload));
      case 'readMutedMembers':         return jsonResponse_(readMutedMembers_(payload));
      case 'muteMember':               return jsonResponse_(muteMember_(payload));
      case 'unmuteMember':             return jsonResponse_(unmuteMember_(payload));
      case 'acceptQuestionReply':      return jsonResponse_(acceptQuestionReply_(payload));
      case 'authSaveLineNotifyToken':   return jsonResponse_(authSaveLineNotifyToken_(payload));
      case 'sendWeeklyDigestLine':      return jsonResponse_(sendWeeklyDigestLineAction_(payload));
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
    if (payload && /^auth/.test(trimAuthText_(payload.action))) {
      recordAuthAudit_(payload.action, 'error', {
        userId: trimAuthText_(payload.userId),
        emailKey: normalizeAuthKey_(payload.emailKey || payload.email || payload.username),
        sessionToken: trimAuthText_(payload.sessionToken),
        detail: error && error.message ? error.message : 'auth_error'
      });
    }
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
  var hashRecord = buildPasswordHashRecord_(password);
  var salt = hashRecord.salt;
  var passwordHash = hashRecord.passwordHash;
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
    referralCode,
    '',
    '0',
    ''
  ]);

  // 紹介コード処理
  if (referralCodeInput) {
    processReferral_(referralCodeInput, userId, users);
  }

  saveLikedCompaniesForUser_(userId, []);
  var sessionToken = createSessionForUser_(userId, payload.userAgent);

  try {
    sendEmailVerificationEmail_({
      id: userId,
      email: email,
      emailKey: emailKey,
      displayName: displayName
    });
  } catch (error) {}

  recordAuthAudit_('authRegister', 'ok', {
    userId: userId,
    emailKey: emailKey,
    sessionToken: sessionToken
  });

  // ウェルカムメール送信
  try {
    MailApp.sendEmail({
      to: email,
      subject: '【慶應就活ナビ】ようこそ！最初にやること3つ',
      htmlBody: '<div style="font-family:sans-serif;max-width:520px;margin:0 auto;padding:24px;color:#1a1a2e;">'
        + '<h1 style="color:#0a1a3e;font-size:1.3rem;">ようこそ、' + displayName + ' さん！</h1>'
        + '<p style="line-height:1.8;">慶應就活ナビへの登録ありがとうございます。まずはこの3つから始めましょう。</p>'
        + '<div style="margin:24px 0;padding:16px;background:#f7f3ed;border-left:4px solid #c9a84c;">'
        + '<p style="margin:0 0 12px;font-weight:700;color:#0a1a3e;">Step 1: ガクチカを1つ書いてみよう</p>'
        + '<p style="margin:0 0 4px;font-size:0.9em;color:#555;">STAR法でエピソードを整理すると、ESや面接でそのまま使えます。</p>'
        + '</div>'
        + '<div style="margin:24px 0;padding:16px;background:#f7f3ed;border-left:4px solid #c9a84c;">'
        + '<p style="margin:0 0 12px;font-weight:700;color:#0a1a3e;">Step 2: ESをAIに添削してもらおう</p>'
        + '<p style="margin:0 0 4px;font-size:0.9em;color:#555;">AIが構成・具体性・表現力の観点でフィードバックしてくれます。何度でも無料。</p>'
        + '</div>'
        + '<div style="margin:24px 0;padding:16px;background:#f7f3ed;border-left:4px solid #c9a84c;">'
        + '<p style="margin:0 0 12px;font-weight:700;color:#0a1a3e;">Step 3: 仲間とつながろう</p>'
        + '<p style="margin:0 0 4px;font-size:0.9em;color:#555;">コミュニティで就活仲間を見つけて、情報交換しましょう。</p>'
        + '</div>'
        + '<p style="margin-top:24px;font-size:0.85em;color:#888;">就活を、ひとりで戦わない。— 慶應就活ナビ</p>'
        + '</div>'
    });
  } catch (e) { /* メール送信失敗しても登録は成功させる */ }

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
      referralCode: referralCode,
      lineNotifyToken: '',
      emailVerified: '0',
      emailVerifiedAt: ''
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
  if (!user || !verifyPasswordRecord_(password, user)) {
    recordAuthRateLimitFailure_('login', emailKey, AUTH_LOGIN_RATE_LIMIT_);
    throw new Error('ログイン情報が正しくありません。');
  }
  maybeUpgradePasswordHash_(usersSheet, users.indexOf(user), user, password);
  clearAuthRateLimit_('login', emailKey);

  var sessionToken = createSessionForUser_(user.id, payload.userAgent);
  recordAuthAudit_('authLogin', 'ok', {
    userId: trimAuthText_(user.id),
    emailKey: emailKey,
    sessionToken: sessionToken
  });
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
var AUTH_AI_RATE_LIMIT_ = {
  maxAttempts: 20,
  windowSeconds: 60 * 60,
  message: 'AI機能の利用回数が上限に達しました。1時間ほど待ってから再試行してください。'
};
var AUTH_RESET_RATE_LIMIT_ = {
  maxAttempts: 5,
  windowSeconds: 15 * 60,
  message: 'パスワードリセット要求が多すぎます。15分ほど待ってから再試行してください。'
};
var AUTH_EMAIL_VERIFICATION_RATE_LIMIT_ = {
  maxAttempts: 5,
  windowSeconds: 60 * 60,
  message: '確認メールの再送回数が多すぎます。しばらく待ってから再試行してください。'
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

function recordAuthRateLimitEvent_(action, value, options) {
  recordAuthRateLimitFailure_(action, value, options);
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

  recordAuthAudit_('authLogout', 'ok', {
    sessionToken: sessionToken
  });

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
  assertAuth_(verifyPasswordRecord_(currentPassword, user), '現在のパスワードが正しくありません。');

  var nextHashRecord = buildPasswordHashRecord_(nextPassword);
  var updatedAt = new Date().toISOString();
  usersSheet.getRange(currentIndex + 2, 5, 1, 2).setValues([[nextHashRecord.passwordHash, nextHashRecord.salt]]);
  usersSheet.getRange(currentIndex + 2, 8).setValue(updatedAt);

  invalidateSessionsForUser_(session.userId, '');

  var nextSessionToken = createSessionForUser_(session.userId, payload.userAgent);
  recordAuthAudit_('authChangePassword', 'ok', {
    userId: session.userId,
    sessionToken: nextSessionToken
  });
  return { status: 'ok', sessionToken: nextSessionToken };
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
  assertAuth_(verifyPasswordRecord_(password, user), 'パスワードが正しくありません。');

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

  cleanupAuthRecordsForUser_(spreadsheet, session.userId, normalizeAuthKey_(user.emailKey || user.email));
  recordAuthAudit_('authDeleteAccount', 'ok', {
    userId: session.userId,
    sessionToken: session.sessionToken,
    detail: 'account_deleted'
  });

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

function getVerifiedUserForSessionToken_(sessionToken) {
  var session = getActiveSessionOrThrow_(sessionToken);
  var user = getAuthUserById_(session.userId);
  assertAuth_(user, 'アカウントが見つかりません。');
  assertAuth_(isEmailVerifiedRecord_(user), 'メールアドレスの確認後に利用できます。');
  return { session: session, user: user };
}

function getVerifiedActionContext_(payload, actionLabel) {
  var session = getActiveSessionOrThrow_(payload && payload.sessionToken);
  var user = getAuthUserById_(session.userId);
  assertAuth_(user, 'アカウントが見つかりません。');
  assertAuth_(isEmailVerifiedRecord_(user), trimAuthText_(actionLabel || 'この操作') + 'はメールアドレス確認後に利用できます。');
  return { session: session, user: user };
}

function assertAiUsageAllowedForSession_(sessionToken, actionName) {
  var context = getVerifiedUserForSessionToken_(sessionToken);
  var rateKey = trimAuthText_(context.user.id) + ':' + trimAuthText_(actionName || 'ai');
  assertAuthRateLimit_('ai', rateKey, AUTH_AI_RATE_LIMIT_);
  recordAuthRateLimitEvent_('ai', rateKey, AUTH_AI_RATE_LIMIT_);
  return context;
}

function createSessionForUser_(userId, userAgent) {
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
    '1',
    trimUserAgent_(userAgent)
  ]);

  return sessionToken;
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
    emailVerified: isEmailVerifiedRecord_(record),
    emailVerifiedAt: trimAuthText_(record.emailVerifiedAt),
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
    'active',
    'userAgent'
  ]);

  ensureSheet_(AUTH_SHEET_NAMES_.liked, [
    'userId',
    'likedCompaniesJson',
    'updatedAt'
  ]);

  ensureSheet_(AUTH_SHEET_NAMES_.emailVerifications, AUTH_EMAIL_VERIFICATION_HEADERS_);
  ensureSheet_(AUTH_SHEET_NAMES_.audit, AUTH_AUDIT_HEADERS_);
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

function trimUserAgent_(value) {
  return trimAuthText_(value).slice(0, 400);
}

function isEmailVerifiedRecord_(record) {
  var verifiedFlag = trimAuthText_(record && record.emailVerified).toLowerCase();
  var verifiedAt = trimAuthText_(record && record.emailVerifiedAt);
  if (!verifiedFlag && !verifiedAt) return true;
  return verifiedFlag === '1' || verifiedFlag === 'true';
}

function getSessionReference_(sessionToken) {
  var token = trimAuthText_(sessionToken);
  if (!token) return '';
  return sha256Hex_('session-ref:' + token).slice(0, 24);
}

function recordAuthAudit_(action, status, meta) {
  try {
    ensureSheet_(AUTH_SHEET_NAMES_.audit, AUTH_AUDIT_HEADERS_);
    getAuthSheet_(AUTH_SHEET_NAMES_.audit).appendRow([
      new Date().toISOString(),
      trimAuthText_(action).slice(0, 80),
      trimAuthText_(status).slice(0, 24),
      trimAuthText_(meta && meta.userId),
      normalizeAuthKey_((meta && meta.emailKey) || ''),
      getSessionReference_(meta && meta.sessionToken),
      trimAuthText_(meta && meta.detail).slice(0, 500)
    ]);
  } catch (error) {}
}

function createEmailVerificationToken_(user) {
  ensureSheet_(AUTH_SHEET_NAMES_.emailVerifications, AUTH_EMAIL_VERIFICATION_HEADERS_);
  var sheet = getAuthSheet_(AUTH_SHEET_NAMES_.emailVerifications);
  var rows = getSheetRecords_(sheet);
  rows.forEach(function (row, index) {
    if (row.userId !== trimAuthText_(user.id)) return;
    if (String(row.used) === 'true') return;
    sheet.getRange(index + 2, 7).setValue('true');
  });

  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var now = new Date();
  var expiresAt = new Date(now.getTime() + EMAIL_VERIFICATION_TOKEN_TTL_HOURS_ * 60 * 60 * 1000);
  sheet.appendRow([
    token,
    trimAuthText_(user.id),
    trimAuthText_(user.email),
    normalizeAuthKey_(user.emailKey || user.email),
    now.toISOString(),
    expiresAt.toISOString(),
    'false'
  ]);
  return token;
}

function sendEmailVerificationEmail_(user) {
  var email = trimAuthText_(user && user.email);
  assertAuth_(email, '確認メールの送信先が見つかりません。');

  var token = createEmailVerificationToken_(user);
  var verifyUrl = getConfiguredSiteBaseUrl_() + '/account.html?mode=verify&token=' + token;
  var displayName = trimAuthText_(user && user.displayName) || email.split('@')[0] || '会員';

  MailApp.sendEmail({
    to: email,
    subject: '【慶應就活ナビ】メールアドレスの確認',
    htmlBody: '<div style="font-family:sans-serif;max-width:480px;margin:0 auto;padding:24px;">'
      + '<h2 style="color:#0a1a3e;">メールアドレスの確認</h2>'
      + '<p>' + displayName + ' さん、登録ありがとうございます。以下のボタンからメールアドレスを確認してください。</p>'
      + '<p style="margin:24px 0;"><a href="' + verifyUrl + '" style="display:inline-block;padding:12px 28px;background:#0a1a3e;color:#fff;text-decoration:none;font-weight:bold;border-radius:4px;">メールアドレスを確認する</a></p>'
      + '<p style="font-size:0.85em;color:#666;">このリンクは' + EMAIL_VERIFICATION_TOKEN_TTL_HOURS_ + '時間有効です。心当たりがない場合はこのメールを無視してください。</p>'
      + '</div>'
  });
}

function getUserSessions_(userId, currentSessionToken) {
  var sessionsSheet = getAuthSheet_(AUTH_SHEET_NAMES_.sessions);
  var rows = getSheetRecords_(sessionsSheet);
  var now = new Date();
  var items = [];

  rows.forEach(function (row, index) {
    if (row.userId !== userId) return;
    var expiresAt = new Date(row.expiresAt);
    var expired = isNaN(expiresAt.getTime()) || expiresAt.getTime() <= now.getTime();
    if (expired && String(row.active) !== '0') {
      sessionsSheet.getRange(index + 2, 6).setValue('0');
    }
    if (String(row.active) === '0' || expired) return;
    items.push({
      sessionRef: getSessionReference_(row.sessionToken),
      createdAt: trimAuthText_(row.createdAt),
      expiresAt: trimAuthText_(row.expiresAt),
      lastSeenAt: trimAuthText_(row.lastSeenAt),
      userAgent: trimUserAgent_(row.userAgent),
      current: trimAuthText_(row.sessionToken) === trimAuthText_(currentSessionToken)
    });
  });

  items.sort(function (left, right) {
    if (left.current && !right.current) return -1;
    if (!left.current && right.current) return 1;
    return new Date(right.lastSeenAt || right.createdAt || 0).getTime() - new Date(left.lastSeenAt || left.createdAt || 0).getTime();
  });

  return items;
}

function cleanupAuthRecordsForUser_(spreadsheet, userId, emailKey) {
  [
    { name: RESET_SHEET_NAME_, userIdHeader: 'userId' },
    { name: AUTH_SHEET_NAMES_.emailVerifications, userIdHeader: 'userId' }
  ].forEach(function (entry) {
    var sheet = spreadsheet.getSheetByName(entry.name);
    if (!sheet) return;
    var rows = getSheetRecords_(sheet);
    for (var i = rows.length - 1; i >= 0; i--) {
      if (trimAuthText_(rows[i][entry.userIdHeader]) === trimAuthText_(userId)) {
        sheet.deleteRow(i + 2);
      }
    }
  });

  var auditSheet = spreadsheet.getSheetByName(AUTH_SHEET_NAMES_.audit);
  if (auditSheet) {
    var auditRows = getSheetRecords_(auditSheet);
    for (var ai = auditRows.length - 1; ai >= 0; ai--) {
      if (trimAuthText_(auditRows[ai].userId) === trimAuthText_(userId) ||
          normalizeAuthKey_(auditRows[ai].emailKey) === normalizeAuthKey_(emailKey)) {
        auditSheet.deleteRow(ai + 2);
      }
    }
  }
}

function getConfiguredSiteBaseUrl_() {
  var siteBaseUrl = trimAuthText_(PropertiesService.getScriptProperties().getProperty('SITE_BASE_URL'));
  assertAuth_(siteBaseUrl, 'SITE_BASE_URL が設定されていません。');
  return siteBaseUrl.replace(/\/+$/, '');
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── モデレーション / ミュート ────────────────────────────────────────────────

var MODERATION_REPORTS_SHEET_NAME_ = 'moderation_reports';
var MODERATION_CONTENT_STATE_SHEET_NAME_ = 'moderation_content_state';
var MODERATION_MUTES_SHEET_NAME_ = 'member_mutes';
var MODERATION_REPORT_HEADERS_ = ['id', 'reporterUserId', 'contentType', 'contentId', 'targetUserId', 'reason', 'details', 'status', 'resolution', 'resolvedBy', 'resolvedAt', 'createdAt'];
var MODERATION_CONTENT_STATE_HEADERS_ = ['contentType', 'contentId', 'state', 'note', 'updatedBy', 'updatedAt'];
var MODERATION_MUTES_HEADERS_ = ['userId', 'mutedUserId', 'createdAt', 'active'];
var MODERATION_REPORTABLE_TYPES_ = ['board_post', 'consultation', 'qa_question', 'experience', 'practice_request'];

function ensureModerationSheets_() {
  ensureSheet_(MODERATION_REPORTS_SHEET_NAME_, MODERATION_REPORT_HEADERS_);
  ensureSheet_(MODERATION_CONTENT_STATE_SHEET_NAME_, MODERATION_CONTENT_STATE_HEADERS_);
  ensureSheet_(MODERATION_MUTES_SHEET_NAME_, MODERATION_MUTES_HEADERS_);
}

function readMutedUserIdsForUser_(userId) {
  ensureModerationSheets_();
  var rows = getSheetRecords_(getAuthSheet_(MODERATION_MUTES_SHEET_NAME_));
  var map = {};
  rows.forEach(function (row) {
    if (trimAuthText_(row.userId) !== trimAuthText_(userId)) return;
    var mutedUserId = trimAuthText_(row.mutedUserId);
    if (!mutedUserId) return;
    if (String(row.active) === '0' || String(row.active).toLowerCase() === 'false') {
      delete map[mutedUserId];
      return;
    }
    map[mutedUserId] = true;
  });
  return map;
}

function getContentStateMap_(contentType) {
  ensureModerationSheets_();
  var rows = getSheetRecords_(getAuthSheet_(MODERATION_CONTENT_STATE_SHEET_NAME_));
  var map = {};
  rows.forEach(function (row) {
    if (trimAuthText_(row.contentType) !== trimAuthText_(contentType)) return;
    var contentId = trimAuthText_(row.contentId);
    if (!contentId) return;
    map[contentId] = trimAuthText_(row.state || 'visible') || 'visible';
  });
  return map;
}

function isContentHidden_(stateMap, contentId) {
  return trimAuthText_(stateMap && stateMap[trimAuthText_(contentId)]) === 'hidden';
}

function setContentState_(contentType, contentId, state, note, adminUserId) {
  ensureModerationSheets_();
  var sheet = getAuthSheet_(MODERATION_CONTENT_STATE_SHEET_NAME_);
  var rows = getSheetRecords_(sheet);
  var type = trimAuthText_(contentType);
  var id = trimAuthText_(contentId);
  if (!type || !id) return;
  var now = new Date().toISOString();
  var idx = findRecordIndex_(rows, function (row) {
    return trimAuthText_(row.contentType) === type && trimAuthText_(row.contentId) === id;
  });
  if (idx >= 0) {
    sheet.getRange(idx + 2, 3, 1, 4).setValues([[trimAuthText_(state || 'visible'), trimAuthText_(note), trimAuthText_(adminUserId), now]]);
    return;
  }
  sheet.appendRow([type, id, trimAuthText_(state || 'visible'), trimAuthText_(note), trimAuthText_(adminUserId), now]);
}

function reportContent_(payload) {
  var context = getVerifiedActionContext_(payload, '通報');
  ensureModerationSheets_();

  var contentType = trimAuthText_(payload.contentType);
  var contentId = trimAuthText_(payload.contentId);
  var reason = trimAuthText_(payload.reason);
  var details = trimAuthText_(payload.details);
  var targetUserId = trimAuthText_(payload.targetUserId);

  assertAuth_(MODERATION_REPORTABLE_TYPES_.indexOf(contentType) >= 0, '通報対象の種類が不正です。');
  assertAuth_(contentId, '通報対象を指定してください。');
  assertAuth_(reason, '通報理由を入力してください。');

  var sheet = getAuthSheet_(MODERATION_REPORTS_SHEET_NAME_);
  var rows = getSheetRecords_(sheet);
  var duplicate = rows.find(function (row) {
    return trimAuthText_(row.reporterUserId) === trimAuthText_(context.session.userId) &&
      trimAuthText_(row.contentType) === contentType &&
      trimAuthText_(row.contentId) === contentId &&
      trimAuthText_(row.status || 'open') === 'open';
  });
  if (duplicate) {
    return { status: 'ok', id: trimAuthText_(duplicate.id), alreadyReported: true };
  }

  var id = 'rep_' + new Date().getTime().toString(36) + Utilities.getUuid().replace(/-/g, '').slice(0, 6);
  var now = new Date().toISOString();
  sheet.appendRow([id, context.session.userId, contentType, contentId, targetUserId, reason, details, 'open', '', '', '', now]);

  recordAuthAudit_('reportContent', 'ok', {
    userId: context.session.userId,
    sessionToken: context.session.sessionToken,
    detail: contentType + ':' + contentId
  });

  return { status: 'ok', id: id };
}

function readMutedMembers_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  return {
    status: 'ok',
    mutedUserIds: Object.keys(readMutedUserIdsForUser_(session.userId))
  };
}

function muteMember_(payload) {
  var context = getVerifiedActionContext_(payload, 'ミュート');
  ensureModerationSheets_();

  var mutedUserId = trimAuthText_(payload.mutedUserId);
  assertAuth_(mutedUserId, 'ミュートする相手を指定してください。');
  assertAuth_(mutedUserId !== trimAuthText_(context.session.userId), '自分自身はミュートできません。');

  var sheet = getAuthSheet_(MODERATION_MUTES_SHEET_NAME_);
  var rows = getSheetRecords_(sheet);
  var idx = findRecordIndex_(rows, function (row) {
    return trimAuthText_(row.userId) === trimAuthText_(context.session.userId) &&
      trimAuthText_(row.mutedUserId) === mutedUserId;
  });
  var now = new Date().toISOString();
  if (idx >= 0) {
    sheet.getRange(idx + 2, 3, 1, 2).setValues([[now, '1']]);
  } else {
    sheet.appendRow([context.session.userId, mutedUserId, now, '1']);
  }
  return { status: 'ok', mutedUserId: mutedUserId };
}

function unmuteMember_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  ensureModerationSheets_();

  var mutedUserId = trimAuthText_(payload.mutedUserId);
  assertAuth_(mutedUserId, 'ミュート解除する相手を指定してください。');

  var sheet = getAuthSheet_(MODERATION_MUTES_SHEET_NAME_);
  var rows = getSheetRecords_(sheet);
  rows.forEach(function (row, index) {
    if (trimAuthText_(row.userId) !== trimAuthText_(session.userId)) return;
    if (trimAuthText_(row.mutedUserId) !== mutedUserId) return;
    sheet.getRange(index + 2, 4).setValue('0');
  });
  return { status: 'ok', mutedUserId: mutedUserId };
}

function getModerationQueue_(payload) {
  var admin = requireAdminSession_(payload);
  ensureModerationSheets_();

  var reports = getSheetRecords_(getAuthSheet_(MODERATION_REPORTS_SHEET_NAME_));
  var users = getSheetRecords_(getAuthSheet_(AUTH_SHEET_NAMES_.users));
  var stateRows = getSheetRecords_(getAuthSheet_(MODERATION_CONTENT_STATE_SHEET_NAME_));
  var userMap = {};
  var contentStateMap = {};
  users.forEach(function (user) {
    userMap[trimAuthText_(user.id)] = trimAuthText_(user.displayName || user.username || user.email);
  });
  stateRows.forEach(function (row) {
    var contentType = trimAuthText_(row.contentType);
    var contentId = trimAuthText_(row.contentId);
    if (!contentType || !contentId) return;
    contentStateMap[contentType + '::' + contentId] = trimAuthText_(row.state || 'visible') || 'visible';
  });

  reports.sort(function (a, b) {
    return new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime();
  });

  var queue = reports.slice(0, 200).map(function (row) {
    return {
      id: trimAuthText_(row.id),
      reporterUserId: trimAuthText_(row.reporterUserId),
      reporterName: userMap[trimAuthText_(row.reporterUserId)] || '不明',
      contentType: trimAuthText_(row.contentType),
      contentId: trimAuthText_(row.contentId),
      targetUserId: trimAuthText_(row.targetUserId),
      targetUserName: userMap[trimAuthText_(row.targetUserId)] || '',
      reason: trimAuthText_(row.reason),
      details: trimAuthText_(row.details),
      status: trimAuthText_(row.status || 'open'),
      resolution: trimAuthText_(row.resolution),
      resolvedBy: trimAuthText_(row.resolvedBy),
      resolvedAt: trimAuthText_(row.resolvedAt),
      createdAt: trimAuthText_(row.createdAt),
      contentState: contentStateMap[trimAuthText_(row.contentType) + '::' + trimAuthText_(row.contentId)] || 'visible'
    };
  });

  var hiddenCount = stateRows.filter(function (row) { return trimAuthText_(row.state) === 'hidden'; }).length;

  recordAuthAudit_('getModerationQueue', 'ok', {
    userId: admin.session.userId,
    sessionToken: admin.session.sessionToken,
    detail: 'reports=' + queue.length
  });

  return {
    status: 'ok',
    reports: queue,
    summary: {
      openCount: queue.filter(function (row) { return row.status === 'open'; }).length,
      hiddenCount: hiddenCount
    }
  };
}

function resolveModerationReport_(payload) {
  var admin = requireAdminSession_(payload);
  ensureModerationSheets_();

  var reportId = trimAuthText_(payload.reportId);
  var resolution = trimAuthText_(payload.resolution);
  var note = trimAuthText_(payload.note);

  assertAuth_(reportId, '通報IDを指定してください。');
  assertAuth_(resolution, '対応内容を指定してください。');

  var reportsSheet = getAuthSheet_(MODERATION_REPORTS_SHEET_NAME_);
  var reports = getSheetRecords_(reportsSheet);
  var idx = findRecordIndex_(reports, function (row) {
    return trimAuthText_(row.id) === reportId;
  });
  assertAuth_(idx >= 0, '通報が見つかりません。');

  var report = reports[idx];
  var now = new Date().toISOString();
  reportsSheet.getRange(idx + 2, 8, 1, 4).setValues([['resolved', resolution, admin.session.userId, now]]);

  if (resolution === 'hidden') {
    setContentState_(report.contentType, report.contentId, 'hidden', note, admin.session.userId);
  } else if (resolution === 'restored') {
    setContentState_(report.contentType, report.contentId, 'visible', note, admin.session.userId);
  }

  recordAuthAudit_('resolveModerationReport', 'ok', {
    userId: admin.session.userId,
    sessionToken: admin.session.sessionToken,
    detail: reportId + ':' + resolution
  });

  return { status: 'ok', resolution: resolution };
}

function getAuthAuditLog_(payload) {
  var admin = requireAdminSession_(payload);
  ensureSheet_(AUTH_SHEET_NAMES_.audit, AUTH_AUDIT_HEADERS_);
  var rows = getSheetRecords_(getAuthSheet_(AUTH_SHEET_NAMES_.audit));
  rows.sort(function (a, b) {
    return new Date(b.timestamp || 0).getTime() - new Date(a.timestamp || 0).getTime();
  });
  return {
    status: 'ok',
    entries: rows.slice(0, 80).map(function (row) {
      return {
        timestamp: trimAuthText_(row.timestamp),
        action: trimAuthText_(row.action),
        status: trimAuthText_(row.status),
        userId: trimAuthText_(row.userId),
        emailKey: trimAuthText_(row.emailKey),
        sessionRef: trimAuthText_(row.sessionRef),
        detail: trimAuthText_(row.detail)
      };
    })
  };
}

// ── ES添削掲示板 ────────────────────────────────────────────────────────────

var BOARD_SHEET_NAME_         = 'es_board';
var BOARD_COMMENTS_SHEET_NAME_ = 'es_board_comments';
var BOARD_MAX_POSTS_           = 50;

function writeBoardPost_(payload) {
  var context  = getVerifiedActionContext_(payload, 'ES投稿');
  var session  = context.session;
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
  var session = getActiveSessionOrThrow_(payload.sessionToken);

  var spreadsheet = getAuthSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(BOARD_SHEET_NAME_);
  if (!sheet) return { status: 'ok', posts: [] };

  var posts = getSheetRecords_(sheet);
  if (!posts.length) return { status: 'ok', posts: [] };
  var hiddenMap = getContentStateMap_('board_post');
  var mutedUserIds = readMutedUserIdsForUser_(session.userId);

  posts.sort(function (a, b) {
    return new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime();
  });
  posts = posts.filter(function (post) {
    return !isContentHidden_(hiddenMap, post.id) && !mutedUserIds[trimAuthText_(post.userId)];
  });
  posts = posts.slice(0, BOARD_MAX_POSTS_);

  var commentsSheet = spreadsheet.getSheetByName(BOARD_COMMENTS_SHEET_NAME_);
  var allComments   = commentsSheet ? getSheetRecords_(commentsSheet) : [];

  var result = posts.map(function (post) {
    var comments = allComments.filter(function (c) { return c.postId === post.id; });
    comments.sort(function (a, b) {
      return new Date(a.createdAt || 0).getTime() - new Date(b.createdAt || 0).getTime();
    });
    comments = comments.filter(function (comment) {
      return !mutedUserIds[trimAuthText_(comment.userId)];
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
  var context     = getVerifiedActionContext_(payload, 'ESへのコメント');
  var session     = context.session;
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

// ── Claude API リトライ付き呼び出し ────────────────────────────────────────

function fetchClaudeWithRetry_(apiKey, requestBody, maxRetries) {
  var retries = maxRetries || 3;
  for (var attempt = 0; attempt < retries; attempt++) {
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code === 200) return response;
    if (code === 529 || code === 503 || code === 429) {
      // Overloaded / Service Unavailable / Rate Limited → リトライ
      if (attempt < retries - 1) {
        Utilities.sleep((attempt + 1) * 3000); // 3秒、6秒、9秒と増やす
        continue;
      }
    }
    // リトライ不可のエラー or 最後のリトライ
    var errBody = {};
    try { errBody = JSON.parse(response.getContentText()); } catch(e) {}
    var errMsg = (errBody.error && errBody.error.message) ? errBody.error.message : ('HTTP ' + code);
    throw new Error('AI処理に失敗しました: ' + errMsg);
  }
  throw new Error('AI処理に失敗しました: リトライ上限に達しました。');
}

// ── Claude API 呼び出し ────────────────────────────────────────────────────

function callClaude_(payload) {
  assertAiUsageAllowedForSession_(payload.sessionToken, 'callClaude');

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
    var limitStrict = charLimit > 0 ? '\n\n⚠️ 最重要ルール: 改善版ESは必ず' + charLimit + '字以内に収めてください。1字でも超えてはいけません。字数制限は' + charLimit + '字です。改善版を書いた後、必ず文字数を数え直して' + charLimit + '字以内であることを確認してください。超えている場合は削って調整してください。' : '';
    systemPrompt = 'あなたは新卒就職活動のプロフェッショナルなESコーチです。学生のESを丁寧に添削してください。フィードバックは日本語で、具体的かつ建設的に行ってください。' + (charLimit > 0 ? '改善版ESを提示する場合、必ず' + charLimit + '字以内に収めてください。字数超過は絶対に避けてください。' : '');
    userMessage  = '【企業名】' + (company || '（未記入）') + '\n【設問】' + (question || '（未記入）') + limitNote + '\n\n【ES内容】\n' + esText + '\n\n以下の観点で詳しく添削してください：\n1. 📌 構成・論理性（PREP法など）\n2. 💡 具体性・エピソードの説得力\n3. 🎯 企業・設問への適合性\n4. ✍️ 表現・文章力\n5. 🔧 改善提案（修正例を示す）' + limitGuide + '\n\n最後に総合評価（S/A/B/C）と一言コメントをつけてください。' + limitStrict;

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

    systemPrompt = 'あなたは新卒就職活動のプロフェッショナルなESコーチです。学生の複数のES設問を総合的に添削してください。各設問の個別評価に加え、複数設問を通じた一貫性・多面性も評価してください。日本語で回答してください。重要: 改善版ESを提示する場合、各設問の字数制限を必ず守ってください。1字でも超えてはいけません。\n\n【重要】回答の最初に、以下の形式でルーブリック評価スコア（各1〜5の整数）をJSON形式で出力してください。必ずこの形式で始めてください：\n[SCORES]{"structure":X,"specificity":X,"logic":X,"expression":X,"persuasion":X}[/SCORES]\n\n各スコアの基準：\n- structure（構成）: 1=構成が不明確 2=やや整理不足 3=基本的な構成はある 4=論理的で分かりやすい 5=完璧な構成\n- specificity（具体性）: 1=抽象的すぎる 2=具体性が不足 3=一部具体的 4=エピソードが具体的 5=非常に具体的で説得力大\n- logic（論理性）: 1=論理破綻 2=やや飛躍がある 3=概ね論理的 4=論理的で一貫性あり 5=完璧な論理展開\n- expression（表現力）: 1=表現が稚拙 2=やや単調 3=標準的 4=表現が豊か 5=非常に洗練された表現\n- persuasion（説得力）: 1=説得力なし 2=やや弱い 3=一定の説得力 4=説得力がある 5=非常に説得力が高い';
    userMessage  = '【企業名】' + (trimAuthText_(input.company) || '（未記入）') + '\n\n' + esListText + '\n\n' +
      '最初に [SCORES]{"structure":X,"specificity":X,"logic":X,"expression":X,"persuasion":X}[/SCORES] の形式でルーブリックスコア（各1〜5）を出力し、その後に以下を評価してください。\n\n# 評価してください\n\n## ① 各設問の個別評価\n各設問について以下を評価：\n- 構成・論理性（PREP法など）\n- 具体性・エピソードの説得力\n- 字数制限が指定されている場合は字数への適合性（削る/増やすべき箇所の提案）\n- 改善提案（修正例を含む）\n- ⚠️ 改善版ESは必ず指定字数以内に収めること。超過厳禁。\n\n## ② 複数設問の総合評価\n1. 🔗 **一貫性** — 自己PR・強みのテーマにブレがないか\n2. 🌐 **多面性** — 異なる強み・経験を引き出せているか\n3. 🎯 **企業適合性** — 全体を通じて志望企業への熱意・適性が伝わるか\n4. ✅ **総合評価（S/A/B/C）と優先改善ポイント**';

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

    var genLimitNote = genLimit > 0 ? '字数制限は' + genLimit + '字以内です。1字でも超えてはいけません。生成後に必ず文字数を数え直して' + genLimit + '字以内であることを確認してください。超えていたら削って調整してください。' : '字数制限は特に指定されていません。';

    systemPrompt = 'あなたは新卒就職活動のプロフェッショナルなESコーチです。学生がメモしたガクチカ（学生時代に力を入れたこと）の情報をもとに、指定された設問に最適化したESを作成してください。日本語で回答してください。' + (genLimit > 0 ? '生成するESは必ず' + genLimit + '字以内に収めてください。字数超過は絶対に避けてください。' : '');
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

  var response = fetchClaudeWithRetry_(apiKey, requestBody, 3);
  var data = JSON.parse(response.getContentText());
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

function authVerifyEmail_(payload) {
  ensureAuthSheets_();

  var verificationToken = trimAuthText_(payload.verificationToken || payload.token);
  assertAuth_(verificationToken, '確認トークンが必要です。');

  var sheet = getAuthSheet_(AUTH_SHEET_NAMES_.emailVerifications);
  var rows = getSheetRecords_(sheet);
  var now = new Date();
  var tokenIndex = findRecordIndex_(rows, function (row) {
    return row.verificationToken === verificationToken && String(row.used) !== 'true';
  });
  assertAuth_(tokenIndex >= 0, '確認リンクが無効または期限切れです。');

  var tokenRecord = rows[tokenIndex];
  var expiresAt = new Date(tokenRecord.expiresAt);
  assertAuth_(!isNaN(expiresAt.getTime()) && expiresAt.getTime() > now.getTime(), '確認リンクの有効期限が切れています。');

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var userIndex = findRecordIndex_(users, function (row) {
    return row.id === tokenRecord.userId;
  });
  assertAuth_(userIndex >= 0, 'アカウントが見つかりません。');

  var verifiedAt = now.toISOString();
  var verifiedCol = AUTH_USER_HEADERS_.indexOf('emailVerified') + 1;
  var verifiedAtCol = AUTH_USER_HEADERS_.indexOf('emailVerifiedAt') + 1;
  sheet.getRange(tokenIndex + 2, 7).setValue('true');
  usersSheet.getRange(userIndex + 2, verifiedCol).setValue('1');
  usersSheet.getRange(userIndex + 2, verifiedAtCol).setValue(verifiedAt);

  var user = users[userIndex];
  user.emailVerified = '1';
  user.emailVerifiedAt = verifiedAt;

  recordAuthAudit_('authVerifyEmail', 'ok', {
    userId: trimAuthText_(user.id),
    emailKey: trimAuthText_(user.emailKey || user.email),
    sessionToken: trimAuthText_(payload.sessionToken)
  });

  return {
    status: 'ok',
    message: 'メールアドレスを確認しました。',
    user: sanitizeUserRecord_(user)
  };
}

function authResendVerificationEmail_(payload) {
  ensureAuthSheets_();

  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var user = users.find(function (row) {
    return row.id === session.userId;
  });
  assertAuth_(user, 'アカウントが見つかりません。');

  if (isEmailVerifiedRecord_(user)) {
    return {
      status: 'ok',
      message: 'メールアドレスはすでに確認済みです。',
      user: sanitizeUserRecord_(user)
    };
  }

  var emailKey = normalizeAuthKey_(user.emailKey || user.email);
  assertAuthRateLimit_('verify-email', emailKey, AUTH_EMAIL_VERIFICATION_RATE_LIMIT_);
  recordAuthRateLimitFailure_('verify-email', emailKey, AUTH_EMAIL_VERIFICATION_RATE_LIMIT_);
  sendEmailVerificationEmail_(user);
  recordAuthAudit_('authResendVerificationEmail', 'ok', {
    userId: trimAuthText_(user.id),
    emailKey: emailKey,
    sessionToken: session.sessionToken
  });

  return {
    status: 'ok',
    message: '確認メールを再送しました。メールをご確認ください。',
    user: sanitizeUserRecord_(user)
  };
}

function authListSessions_(payload) {
  ensureAuthSheets_();
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var sessions = getUserSessions_(session.userId, session.sessionToken);

  recordAuthAudit_('authListSessions', 'ok', {
    userId: session.userId,
    sessionToken: session.sessionToken,
    detail: 'count=' + sessions.length
  });

  return {
    status: 'ok',
    currentSessionRef: getSessionReference_(session.sessionToken),
    sessions: sessions
  };
}

function authRevokeSession_(payload) {
  ensureAuthSheets_();

  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var sessionRef = trimAuthText_(payload.sessionRef);
  assertAuth_(sessionRef, 'セッションを指定してください。');

  var sessionsSheet = getAuthSheet_(AUTH_SHEET_NAMES_.sessions);
  var rows = getSheetRecords_(sessionsSheet);
  var targetIndex = findRecordIndex_(rows, function (row) {
    return row.userId === session.userId &&
      getSessionReference_(row.sessionToken) === sessionRef &&
      String(row.active) !== '0';
  });
  assertAuth_(targetIndex >= 0, '対象のセッションが見つかりません。');

  var target = rows[targetIndex];
  sessionsSheet.getRange(targetIndex + 2, 6).setValue('0');
  var currentRevoked = trimAuthText_(target.sessionToken) === trimAuthText_(session.sessionToken);

  recordAuthAudit_('authRevokeSession', 'ok', {
    userId: session.userId,
    sessionToken: session.sessionToken,
    detail: 'target=' + sessionRef
  });

  return {
    status: 'ok',
    currentRevoked: currentRevoked
  };
}

function authRevokeOtherSessions_(payload) {
  ensureAuthSheets_();

  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var sessionsSheet = getAuthSheet_(AUTH_SHEET_NAMES_.sessions);
  var rows = getSheetRecords_(sessionsSheet);
  var revokedCount = 0;

  rows.forEach(function (row, index) {
    if (row.userId !== session.userId) return;
    if (trimAuthText_(row.sessionToken) === trimAuthText_(session.sessionToken)) return;
    if (String(row.active) === '0') return;
    sessionsSheet.getRange(index + 2, 6).setValue('0');
    revokedCount += 1;
  });

  recordAuthAudit_('authRevokeOtherSessions', 'ok', {
    userId: session.userId,
    sessionToken: session.sessionToken,
    detail: 'count=' + revokedCount
  });

  return {
    status: 'ok',
    revokedCount: revokedCount
  };
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
    recordAuthAudit_('authRequestPasswordReset', 'ok', {
      emailKey: emailKey,
      detail: 'user_not_found'
    });
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
  var resetUrl = getConfiguredSiteBaseUrl_() + '/account.html?mode=reset&token=' + resetToken;
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

  recordAuthAudit_('authRequestPasswordReset', 'ok', {
    userId: trimAuthText_(user.id),
    emailKey: emailKey
  });

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
  var nextHashRecord = buildPasswordHashRecord_(newPassword);
  usersSheet.getRange(userIndex + 2, 5, 1, 2).setValues([[nextHashRecord.passwordHash, nextHashRecord.salt]]);
  usersSheet.getRange(userIndex + 2, 8).setValue(now.toISOString());

  // 既存セッションを全て無効化
  var sessionsSheet = getAuthSheet_(AUTH_SHEET_NAMES_.sessions);
  var sessions = getSheetRecords_(sessionsSheet);
  sessions.forEach(function (row, i) {
    if (row.userId === tokenRecord.userId && String(row.active) !== '0') {
      sessionsSheet.getRange(i + 2, 6).setValue('0');
    }
  });

  recordAuthAudit_('authResetPassword', 'ok', {
    userId: trimAuthText_(user.id),
    emailKey: trimAuthText_(user.emailKey || user.email)
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
    if (idx < 0) {
      idx = findRecordIndex_(rows, function (row) {
        return row.company === company && row.userId === session.userId;
      });
    }
    if (idx >= 0) {
      sheet.getRange(idx + 2, 3, 1, 6).setValues([[
        rowData.company, rowData.industry, rowData.deadline,
        rowData.stage, rowData.status, rowData.memo
      ]]);
      sheet.getRange(idx + 2, 10).setValue(now);
      return { status: 'ok', id: rows[idx].id || id };
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

// ── 自己PR保管庫 ───────────────────────────────────────────────────────────

var SELFPR_SHEET_NAME_ = 'selfpr_bank';
var SELFPR_HEADERS_ = ['id','userId','title','point1','reason','example','point2','fullText','targetIndustry','memo','createdAt','updatedAt'];

function writeSelfPR_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var pr = payload.selfPR || {};
  var now = new Date().toISOString();
  var id = trimAuthText_(pr.id);

  var ss = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(SELFPR_SHEET_NAME_);
  if (!sheet) { sheet = ss.insertSheet(SELFPR_SHEET_NAME_); sheet.appendRow(SELFPR_HEADERS_); }

  if (id) {
    var data = sheet.getDataRange().getValues();
    var hdr = data[0];
    var idCol = hdr.indexOf('id');
    var uidCol = hdr.indexOf('userId');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === id && String(data[i][uidCol]) === session.userId) {
        SELFPR_HEADERS_.forEach(function(h, c) {
          if (h === 'id' || h === 'userId' || h === 'createdAt') return;
          if (h === 'updatedAt') { sheet.getRange(i + 1, c + 1).setValue(now); return; }
          if (pr[h] !== undefined) sheet.getRange(i + 1, c + 1).setValue(trimAuthText_(pr[h]));
        });
        return { status: 'ok', id: id };
      }
    }
    throw new Error('自己PRが見つかりません。');
  }

  var newId = 'pr_' + now.replace(/\D/g, '').slice(0, 14) + '_' + Math.floor(Math.random() * 10000);
  var row = SELFPR_HEADERS_.map(function(h) {
    if (h === 'id') return newId;
    if (h === 'userId') return session.userId;
    if (h === 'createdAt') return now;
    if (h === 'updatedAt') return now;
    return trimAuthText_(pr[h] || '');
  });
  sheet.appendRow(row);
  return { status: 'ok', id: newId };
}

function readMySelfPR_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var ss = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(SELFPR_SHEET_NAME_);
  if (!sheet) return { status: 'ok', entries: [] };

  var data = sheet.getDataRange().getValues();
  var hdr = data[0];
  var uidCol = hdr.indexOf('userId');
  var entries = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uidCol]) !== session.userId) continue;
    var entry = {};
    hdr.forEach(function(h, c) { entry[h] = String(data[i][c] || ''); });
    entries.push(entry);
  }
  return { status: 'ok', entries: entries };
}

function deleteSelfPR_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var id = trimAuthText_(payload.id);
  if (!id) throw new Error('IDが指定されていません。');

  var ss = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(SELFPR_SHEET_NAME_);
  if (!sheet) throw new Error('自己PRが見つかりません。');

  var data = sheet.getDataRange().getValues();
  var hdr = data[0];
  var idCol = hdr.indexOf('id');
  var uidCol = hdr.indexOf('userId');
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idCol]) === id && String(data[i][uidCol]) === session.userId) {
      sheet.deleteRow(i + 1);
      return { status: 'ok' };
    }
  }
  throw new Error('自己PRが見つかりません。');
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
  var context = getVerifiedActionContext_(payload, 'グループ作成');
  var session = context.session;
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
  var context = getVerifiedActionContext_(payload, 'グループ参加');
  var session = context.session;
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
  var context = getVerifiedActionContext_(payload, '公開設定の変更');
  var session = context.session;
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

// ── #今日の就活 Timeline ──

function writeTimeline_(payload) {
  var context = getVerifiedActionContext_(payload, '#今日の就活投稿');
  var session = context.session;
  var text = trimAuthText_(payload.text);
  assertAuth_(text, '投稿内容を入力してください。');
  assertAuth_(text.length <= 140, '140字以内で入力してください。');

  // Get user's display name
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var user = users.find(function(r) { return r.id === session.userId; });
  var displayName = user ? trimAuthText_(user.displayName || user.username) : '';

  var ss = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName('activity_timeline');
  if (!sheet) { sheet = ss.insertSheet('activity_timeline'); sheet.appendRow(['id','userId','displayName','text','createdAt']); }

  var now = new Date().toISOString();
  var id = 'tl_' + now.replace(/\D/g,'').slice(0,14) + '_' + Math.floor(Math.random()*10000);
  sheet.appendRow([id, session.userId, displayName, text, now]);

  return { status: 'ok', id: id };
}

function readTimeline_(payload) {
  getActiveSessionOrThrow_(payload.sessionToken);
  var ss = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName('activity_timeline');
  if (!sheet) return { status: 'ok', entries: [] };

  var rows = getSheetRecords_(sheet);
  rows.sort(function(a,b) { return new Date(b.createdAt||0) - new Date(a.createdAt||0); });
  var recent = rows.slice(0, 50);

  return {
    status: 'ok',
    entries: recent.map(function(r) {
      return {
        id: trimAuthText_(r.id),
        displayName: trimAuthText_(r.displayName),
        text: trimAuthText_(r.text),
        createdAt: trimAuthText_(r.createdAt)
      };
    })
  };
}

// ── Gamification / Stats ──

function getMyStats_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var uid = session.userId;
  var ss = getAuthSpreadsheet_();

  var esCount = 0, gkCount = 0, boardCount = 0, ipCount = 0, prCount = 0, timelineCount = 0;

  // Count ES entries
  var esSheet = ss.getSheetByName('es_bank');
  if (esSheet) { esCount = getSheetRecords_(esSheet).filter(function(r){return r.userId===uid;}).length; }

  // Count Gakuchika
  var gkSheet = ss.getSheetByName('gakuchika_bank');
  if (gkSheet) { gkCount = getSheetRecords_(gkSheet).filter(function(r){return r.userId===uid;}).length; }

  // Count board posts
  var boardSheet = ss.getSheetByName('es_board');
  if (boardSheet) { boardCount = getSheetRecords_(boardSheet).filter(function(r){return r.userId===uid;}).length; }

  // Count interview prep questions answered
  var ipSheet = ss.getSheetByName('ip_questions');
  if (ipSheet) { ipCount = getSheetRecords_(ipSheet).filter(function(r){return r.userId===uid && trimAuthText_(r.answer);}).length; }

  // Count self PRs
  var prSheet = ss.getSheetByName('selfpr_bank');
  if (prSheet) { prCount = getSheetRecords_(prSheet).filter(function(r){return r.userId===uid;}).length; }

  // Count timeline posts
  var tlSheet = ss.getSheetByName('activity_timeline');
  if (tlSheet) { timelineCount = getSheetRecords_(tlSheet).filter(function(r){return r.userId===uid;}).length; }

  var points = esCount * 10 + gkCount * 10 + boardCount * 15 + ipCount * 5 + prCount * 10 + timelineCount * 2;

  var badges = [];
  if (esCount >= 1) badges.push('ES初投稿');
  if (esCount >= 5) badges.push('ES5件達成');
  if (esCount >= 10) badges.push('ESマスター');
  if (gkCount >= 1) badges.push('ガクチカ初登録');
  if (gkCount >= 3) badges.push('ガクチカ3件');
  if (boardCount >= 1) badges.push('掲示板デビュー');
  if (boardCount >= 5) badges.push('掲示板常連');
  if (ipCount >= 10) badges.push('面接準備10問');
  if (ipCount >= 30) badges.push('面接マスター');
  if (prCount >= 1) badges.push('自己PR完成');
  if (timelineCount >= 5) badges.push('毎日就活');
  if (points >= 100) badges.push('100pt達成');
  if (points >= 300) badges.push('300pt達成');

  return {
    status: 'ok',
    stats: { esCount:esCount, gkCount:gkCount, boardCount:boardCount, ipCount:ipCount, prCount:prCount, timelineCount:timelineCount },
    points: points,
    badges: badges
  };
}

// ── 面接振り返りノート ─────────────────────────────────────────────────────────

var INTERVIEW_REVIEW_SHEET_NAME_ = 'interview_reviews';
var INTERVIEW_REVIEW_HEADERS_    = ['id','userId','date','company','interviewType','questionsAsked','myAnswers','reflection','improvement','rating','createdAt','updatedAt'];

function getOrCreateInterviewReviewSheet_() {
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(INTERVIEW_REVIEW_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(INTERVIEW_REVIEW_SHEET_NAME_);
    sheet.appendRow(INTERVIEW_REVIEW_HEADERS_);
  }
  return sheet;
}

function writeInterviewReview_(payload) {
  var session       = getActiveSessionOrThrow_(payload.sessionToken);
  var review        = payload.review || {};
  var id            = trimAuthText_(review.id);
  var date          = trimAuthText_(review.date);
  var company       = trimAuthText_(review.company);
  var interviewType = trimAuthText_(review.interviewType);
  var questionsAsked = trimAuthText_(review.questionsAsked);
  var myAnswers     = trimAuthText_(review.myAnswers);
  var reflection    = trimAuthText_(review.reflection);
  var improvement   = trimAuthText_(review.improvement);
  var rating        = trimAuthText_(review.rating);

  assertAuth_(company,  '企業名を入力してください。');
  assertAuth_(reflection, '振り返り内容を入力してください。');

  var sheet = getOrCreateInterviewReviewSheet_();
  var now   = new Date().toISOString();

  if (id) {
    var rows = getSheetRecords_(sheet);
    var idx  = findRecordIndex_(rows, function(r) { return String(r.id) === id && String(r.userId) === session.userId; });
    assertAuth_(idx >= 0, '振り返りが見つかりません。');
    sheet.getRange(idx + 2, 3, 1, 10).setValues([[date, company, interviewType, questionsAsked, myAnswers, reflection, improvement, rating, rows[idx].createdAt, now]]);
    return { status: 'ok', id: id };
  }

  var newId = 'ir_' + new Date().getTime().toString(36) + Utilities.getUuid().replace(/-/g, '').slice(0, 6);
  sheet.appendRow([newId, session.userId, date, company, interviewType, questionsAsked, myAnswers, reflection, improvement, rating, now, now]);
  return { status: 'ok', id: newId };
}

function readMyInterviewReviews_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var sheet   = getOrCreateInterviewReviewSheet_();
  var rows    = getSheetRecords_(sheet);
  var mine    = rows.filter(function(r) { return String(r.userId) === session.userId; });

  return {
    status: 'ok',
    entries: mine.map(function(r) {
      return {
        id:             trimAuthText_(r.id),
        date:           trimAuthText_(r.date),
        company:        trimAuthText_(r.company),
        interviewType:  trimAuthText_(r.interviewType),
        questionsAsked: trimAuthText_(r.questionsAsked),
        myAnswers:      trimAuthText_(r.myAnswers),
        reflection:     trimAuthText_(r.reflection),
        improvement:    trimAuthText_(r.improvement),
        rating:         trimAuthText_(r.rating),
        createdAt:      trimAuthText_(r.createdAt),
        updatedAt:      trimAuthText_(r.updatedAt)
      };
    })
  };
}

function deleteInterviewReview_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var id      = trimAuthText_(payload.id);
  assertAuth_(id, 'IDが指定されていません。');

  var sheet = getOrCreateInterviewReviewSheet_();
  var rows  = getSheetRecords_(sheet);
  var idx   = findRecordIndex_(rows, function(r) { return String(r.id) === id && String(r.userId) === session.userId; });
  assertAuth_(idx >= 0, '振り返りが見つかりません。');
  sheet.deleteRow(idx + 2);
  return { status: 'ok' };
}

// ── 就活相談掲示板 ─────────────────────────────────────────────────────────────

var CONSULTATION_SHEET_NAME_          = 'consultation_board';
var CONSULTATION_COMMENTS_SHEET_NAME_ = 'consultation_comments';
var CONSULTATION_MAX_POSTS_           = 50;
var CONSULTATION_CATEGORIES_          = ['業界選び','ES・書類','面接','内定・承諾','メンタル','その他'];

function writeConsultation_(payload) {
  var context     = getVerifiedActionContext_(payload, '相談投稿');
  var session     = context.session;
  var category    = trimAuthText_(payload.category);
  var title       = trimAuthText_(payload.title);
  var text        = trimAuthText_(payload.text || payload.body);

  assertAuth_(title, 'タイトルを入力してください。');
  assertAuth_(text,  '相談内容を入力してください。');
  if (CONSULTATION_CATEGORIES_.indexOf(category) < 0) category = 'その他';

  var spreadsheet = getAuthSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(CONSULTATION_SHEET_NAME_);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONSULTATION_SHEET_NAME_);
    sheet.appendRow(['id','userId','displayName','category','title','text','createdAt']);
  }

  var now        = new Date().toISOString();
  var id         = now.replace(/\D/g, '').slice(0, 14) + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users      = getSheetRecords_(usersSheet);
  var userRecord = users.find(function (row) { return row.id === session.userId; });
  var displayName = userRecord ? trimAuthText_(userRecord.displayName || userRecord.username) : '';

  sheet.appendRow([id, session.userId, displayName, category, title, text, now]);

  return { status: 'ok', id: id };
}

function readConsultations_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);

  var spreadsheet = getAuthSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(CONSULTATION_SHEET_NAME_);
  if (!sheet) return { status: 'ok', posts: [] };

  var posts = getSheetRecords_(sheet);
  if (!posts.length) return { status: 'ok', posts: [] };
  var hiddenMap = getContentStateMap_('consultation');
  var mutedUserIds = readMutedUserIdsForUser_(session.userId);

  posts.sort(function (a, b) {
    return new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime();
  });
  posts = posts.filter(function (post) {
    return !isContentHidden_(hiddenMap, post.id) && !mutedUserIds[trimAuthText_(post.userId)];
  });
  posts = posts.slice(0, CONSULTATION_MAX_POSTS_);

  var commentsSheet = spreadsheet.getSheetByName(CONSULTATION_COMMENTS_SHEET_NAME_);
  var allComments   = commentsSheet ? getSheetRecords_(commentsSheet) : [];

  var result = posts.map(function (post) {
    var comments = allComments.filter(function (c) { return c.postId === post.id; });
    comments.sort(function (a, b) {
      return new Date(a.createdAt || 0).getTime() - new Date(b.createdAt || 0).getTime();
    });
    comments = comments.filter(function (comment) {
      return !mutedUserIds[trimAuthText_(comment.userId)];
    });
    return {
      id:          trimAuthText_(post.id),
      userId:      trimAuthText_(post.userId),
      displayName: trimAuthText_(post.displayName),
      category:    trimAuthText_(post.category),
      title:       trimAuthText_(post.title),
      text:        trimAuthText_(post.text),
      createdAt:   trimAuthText_(post.createdAt),
      comments:    comments.map(function (c) {
        return {
          id:          trimAuthText_(c.id),
          postId:      trimAuthText_(c.postId),
          userId:      trimAuthText_(c.userId),
          displayName: trimAuthText_(c.displayName),
          text:        trimAuthText_(c.text),
          createdAt:   trimAuthText_(c.createdAt)
        };
      })
    };
  });

  return { status: 'ok', posts: result };
}

function addConsultationComment_(payload) {
  var context = getVerifiedActionContext_(payload, '相談へのコメント');
  var session = context.session;
  var postId  = trimAuthText_(payload.postId);
  var text    = trimAuthText_(payload.text);

  assertAuth_(postId, 'postIdが指定されていません。');
  assertAuth_(text,   'コメントを入力してください。');

  var spreadsheet = getAuthSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(CONSULTATION_COMMENTS_SHEET_NAME_);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONSULTATION_COMMENTS_SHEET_NAME_);
    sheet.appendRow(['id','postId','userId','displayName','text','createdAt']);
  }

  var now        = new Date().toISOString();
  var id         = now.replace(/\D/g, '').slice(0, 14) + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users      = getSheetRecords_(usersSheet);
  var userRecord = users.find(function (row) { return row.id === session.userId; });
  var displayName = userRecord ? trimAuthText_(userRecord.displayName || userRecord.username) : '';

  sheet.appendRow([id, postId, session.userId, displayName, text, now]);

  return { status: 'ok' };
}

// ── GD練習マッチング ───────────────────────────────────────────────────────────

var GD_SESSIONS_SHEET_NAME_      = 'gd_sessions';
var GD_PARTICIPANTS_SHEET_NAME_  = 'gd_participants';
var GD_SESSIONS_HEADERS_         = ['id','creatorId','creatorName','title','scheduledDate','scheduledTime','maxParticipants','meetUrl','aiTheme','status','createdAt'];
var GD_PARTICIPANTS_HEADERS_     = ['sessionId','userId','displayName','joinedAt'];

function getOrCreateGDSessionsSheet_() {
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(GD_SESSIONS_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(GD_SESSIONS_SHEET_NAME_);
    sheet.appendRow(GD_SESSIONS_HEADERS_);
  }
  return sheet;
}

function getOrCreateGDParticipantsSheet_() {
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(GD_PARTICIPANTS_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(GD_PARTICIPANTS_SHEET_NAME_);
    sheet.appendRow(GD_PARTICIPANTS_HEADERS_);
  }
  return sheet;
}

function createGDSession_(payload) {
  var context         = getVerifiedActionContext_(payload, 'GDセッション作成');
  var session         = context.session;
  var title           = trimAuthText_(payload.title);
  var scheduledDate   = trimAuthText_(payload.scheduledDate || payload.date);
  var scheduledTime   = trimAuthText_(payload.scheduledTime || payload.time);
  var maxParticipants = trimAuthText_(payload.maxParticipants) || '6';
  var meetUrl         = trimAuthText_(payload.meetUrl);

  assertAuth_(title,         'タイトルを入力してください。');
  assertAuth_(scheduledDate, '日付を入力してください。');
  assertAuth_(scheduledTime, '時間を入力してください。');

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users      = getSheetRecords_(usersSheet);
  var userRecord = users.find(function (row) { return row.id === session.userId; });
  var creatorName = userRecord ? trimAuthText_(userRecord.displayName || userRecord.username) : '';

  var sheet = getOrCreateGDSessionsSheet_();
  var now   = new Date().toISOString();
  var id    = 'gd_' + new Date().getTime().toString(36) + Utilities.getUuid().replace(/-/g, '').slice(0, 6);

  sheet.appendRow([id, session.userId, creatorName, title, scheduledDate, scheduledTime, maxParticipants, meetUrl, '', 'open', now]);

  // Creator auto-joins the session
  var pSheet = getOrCreateGDParticipantsSheet_();
  pSheet.appendRow([id, session.userId, creatorName, now]);

  return { status: 'ok', id: id };
}

function readGDSessions_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var mutedUserIds = readMutedUserIdsForUser_(session.userId);

  var sessionsSheet = getOrCreateGDSessionsSheet_();
  var sessions      = getSheetRecords_(sessionsSheet);

  // Filter to open sessions and sort by scheduled date
  sessions = sessions.filter(function(s) {
    return s.status === 'open' && !mutedUserIds[trimAuthText_(s.creatorId)];
  });
  sessions.sort(function(a, b) {
    var da = (a.scheduledDate || '') + (a.scheduledTime || '');
    var db = (b.scheduledDate || '') + (b.scheduledTime || '');
    return da > db ? 1 : da < db ? -1 : 0;
  });

  var pSheet          = getOrCreateGDParticipantsSheet_();
  var allParticipants = getSheetRecords_(pSheet);

  var result = sessions.map(function(s) {
    var participants = allParticipants.filter(function(p) { return p.sessionId === s.id; });
    participants = participants.filter(function (participant) {
      return !mutedUserIds[trimAuthText_(participant.userId)];
    });
    var isParticipant = participants.some(function (p) { return trimAuthText_(p.userId) === session.userId; });
    var canViewMeetUrl = isParticipant || trimAuthText_(s.creatorId) === session.userId;
    return {
      id:              trimAuthText_(s.id),
      creatorId:       trimAuthText_(s.creatorId),
      creatorName:     trimAuthText_(s.creatorName),
      title:           trimAuthText_(s.title),
      scheduledDate:   trimAuthText_(s.scheduledDate),
      scheduledTime:   trimAuthText_(s.scheduledTime),
      maxParticipants: trimAuthText_(s.maxParticipants),
      meetUrl:         canViewMeetUrl ? trimAuthText_(s.meetUrl) : '',
      aiTheme:         trimAuthText_(s.aiTheme),
      status:          trimAuthText_(s.status),
      createdAt:       trimAuthText_(s.createdAt),
      canViewMeetUrl:  canViewMeetUrl,
      isParticipant:   isParticipant,
      participantCount: participants.length,
      participants:    participants.map(function(p) {
        return { userId: trimAuthText_(p.userId), displayName: trimAuthText_(p.displayName), joinedAt: trimAuthText_(p.joinedAt) };
      })
    };
  });

  return { status: 'ok', sessions: result };
}

function joinGDSession_(payload) {
  var context   = getVerifiedActionContext_(payload, 'GDセッション参加');
  var session   = context.session;
  var sessionId = trimAuthText_(payload.sessionId);
  assertAuth_(sessionId, 'セッションIDが指定されていません。');

  var sessionsSheet = getOrCreateGDSessionsSheet_();
  var sessions      = getSheetRecords_(sessionsSheet);
  var gdSession     = sessions.find(function(s) { return String(s.id) === sessionId; });
  assertAuth_(gdSession, 'セッションが見つかりません。');
  assertAuth_(gdSession.status === 'open', 'このセッションは締め切られています。');

  var pSheet       = getOrCreateGDParticipantsSheet_();
  var participants = getSheetRecords_(pSheet);
  var already      = participants.find(function(p) { return p.sessionId === sessionId && p.userId === session.userId; });
  if (already) return { status: 'ok', message: '既に参加しています。' };

  var currentCount = participants.filter(function(p) { return p.sessionId === sessionId; }).length;
  var max          = parseInt(gdSession.maxParticipants) || 6;
  assertAuth_(currentCount < max, '定員に達しています。');

  var usersSheet  = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users       = getSheetRecords_(usersSheet);
  var userRecord  = users.find(function(row) { return row.id === session.userId; });
  var displayName = userRecord ? trimAuthText_(userRecord.displayName || userRecord.username) : '';

  var now = new Date().toISOString();
  pSheet.appendRow([sessionId, session.userId, displayName, now]);

  return { status: 'ok' };
}

function leaveGDSession_(payload) {
  var session   = getActiveSessionOrThrow_(payload.sessionToken);
  var sessionId = trimAuthText_(payload.sessionId);
  assertAuth_(sessionId, 'セッションIDが指定されていません。');

  var pSheet       = getOrCreateGDParticipantsSheet_();
  var participants = getSheetRecords_(pSheet);
  var idx          = findRecordIndex_(participants, function(p) { return p.sessionId === sessionId && p.userId === session.userId; });
  assertAuth_(idx >= 0, 'このセッションに参加していません。');
  pSheet.deleteRow(idx + 2);

  return { status: 'ok' };
}

function generateGDTheme_(payload) {
  assertAiUsageAllowedForSession_(payload.sessionToken, 'generateGDTheme');
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  assertAuth_(apiKey, 'AI機能が設定されていません。');

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:'post', contentType:'application/json',
    headers:{'x-api-key':apiKey,'anthropic-version':'2023-06-01'},
    payload:JSON.stringify({
      model:'claude-sonnet-4-20250514', max_tokens:300,
      system:'就活のグループディスカッション練習用のテーマを1つ提案してください。テーマと簡単な背景説明を日本語で返してください。',
      messages:[{role:'user',content:'GDのテーマを1つ提案してください。'}]
    }),
    muteHttpExceptions:true
  });

  var text = '';
  if (resp.getResponseCode() === 200) {
    (JSON.parse(resp.getContentText()).content||[]).forEach(function(b){if(b.type==='text')text+=b.text;});
  }
  return { status:'ok', theme: text || 'テーマの生成に失敗しました。' };
}

// ── OB訪問対策 ──────────────────────────────────────────────────────────────

var OB_VISIT_SHEET_NAME_ = 'ob_visits';
var OB_VISIT_HEADERS_    = ['id','userId','company','obName','obDepartment','visitDate','status','prepQuestions','prepNotes','reviewNotes','impressions','nextActions','rating','createdAt','updatedAt'];

function getOrCreateOBVisitSheet_() {
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(OB_VISIT_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(OB_VISIT_SHEET_NAME_);
    sheet.appendRow(OB_VISIT_HEADERS_);
  }
  return sheet;
}

function writeOBVisit_(payload) {
  var session       = getActiveSessionOrThrow_(payload.sessionToken);
  var visit         = payload.visit || {};
  var id            = trimAuthText_(visit.id);
  var company       = trimAuthText_(visit.company);
  var obName        = trimAuthText_(visit.obName);
  var obDepartment  = trimAuthText_(visit.obDepartment);
  var visitDate     = trimAuthText_(visit.visitDate);
  var status        = trimAuthText_(visit.status);
  var prepQuestions = trimAuthText_(visit.prepQuestions);
  var prepNotes     = trimAuthText_(visit.prepNotes);
  var reviewNotes   = trimAuthText_(visit.reviewNotes);
  var impressions   = trimAuthText_(visit.impressions);
  var nextActions   = trimAuthText_(visit.nextActions);
  var rating        = trimAuthText_(visit.rating);

  assertAuth_(company, '企業名（テーマ）を入力してください。');
  if (['planned','completed'].indexOf(status) < 0) status = 'planned';

  var sheet = getOrCreateOBVisitSheet_();
  var now   = new Date().toISOString();

  if (id) {
    var rows = getSheetRecords_(sheet);
    var idx  = findRecordIndex_(rows, function(r) { return String(r.id) === id && String(r.userId) === session.userId; });
    assertAuth_(idx >= 0, 'OB訪問記録が見つかりません。');
    sheet.getRange(idx + 2, 3, 1, 13).setValues([[company, obName, obDepartment, visitDate, status, prepQuestions, prepNotes, reviewNotes, impressions, nextActions, rating, rows[idx].createdAt, now]]);
    return { status: 'ok', id: id };
  }

  var newId = 'ob_' + new Date().getTime().toString(36) + Utilities.getUuid().replace(/-/g, '').slice(0, 6);
  sheet.appendRow([newId, session.userId, company, obName, obDepartment, visitDate, status, prepQuestions, prepNotes, reviewNotes, impressions, nextActions, rating, now, now]);
  return { status: 'ok', id: newId };
}

function readMyOBVisits_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var sheet   = getOrCreateOBVisitSheet_();
  var rows    = getSheetRecords_(sheet);
  var mine    = rows.filter(function(r) { return String(r.userId) === session.userId; });

  return {
    status: 'ok',
    entries: mine.map(function(r) {
      return {
        id:            trimAuthText_(r.id),
        company:       trimAuthText_(r.company),
        obName:        trimAuthText_(r.obName),
        obDepartment:  trimAuthText_(r.obDepartment),
        visitDate:     trimAuthText_(r.visitDate),
        status:        trimAuthText_(r.status),
        prepQuestions: trimAuthText_(r.prepQuestions),
        prepNotes:     trimAuthText_(r.prepNotes),
        reviewNotes:   trimAuthText_(r.reviewNotes),
        impressions:   trimAuthText_(r.impressions),
        nextActions:   trimAuthText_(r.nextActions),
        rating:        trimAuthText_(r.rating),
        createdAt:     trimAuthText_(r.createdAt),
        updatedAt:     trimAuthText_(r.updatedAt)
      };
    })
  };
}

function deleteOBVisit_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var id      = trimAuthText_(payload.id);
  assertAuth_(id, 'IDが指定されていません。');

  var sheet = getOrCreateOBVisitSheet_();
  var rows  = getSheetRecords_(sheet);
  var idx   = findRecordIndex_(rows, function(r) { return String(r.id) === id && String(r.userId) === session.userId; });
  assertAuth_(idx >= 0, 'OB訪問記録が見つかりません。');
  sheet.deleteRow(idx + 2);
  return { status: 'ok' };
}

// ── OB訪問 匿名共有 ─────────────────────────────────────────────────────────

var SHARED_OB_SHEET_NAME_ = 'shared_ob_info';
var SHARED_OB_HEADERS_    = ['id','industry','department','yearsOfExp','keyInsights','tips','sharedAt'];

function getOrCreateSharedOBSheet_() {
  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(SHARED_OB_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(SHARED_OB_SHEET_NAME_);
    sheet.appendRow(SHARED_OB_HEADERS_);
  }
  return sheet;
}

function shareOBInfo_(payload) {
  getActiveSessionOrThrow_(payload.sessionToken);
  var info        = payload.info || {};
  var industry    = trimAuthText_(info.industry);
  var department  = trimAuthText_(info.department);
  var yearsOfExp  = trimAuthText_(info.yearsOfExp);
  var keyInsights = trimAuthText_(info.keyInsights);
  var tips        = trimAuthText_(info.tips);

  assertAuth_(industry || department || keyInsights, '共有する情報を入力してください。');

  var sheet = getOrCreateSharedOBSheet_();
  var now   = new Date().toISOString();
  var newId = 'sob_' + new Date().getTime().toString(36) + Utilities.getUuid().replace(/-/g, '').slice(0, 6);
  sheet.appendRow([newId, industry, department, yearsOfExp, keyInsights, tips, now]);
  return { status: 'ok', id: newId };
}

function readSharedOBInfo_(payload) {
  getActiveSessionOrThrow_(payload.sessionToken);
  var sheet = getOrCreateSharedOBSheet_();
  var rows  = getSheetRecords_(sheet);

  return {
    status: 'ok',
    entries: rows.map(function(r) {
      return {
        id:          trimAuthText_(r.id),
        industry:    trimAuthText_(r.industry),
        department:  trimAuthText_(r.department),
        yearsOfExp:  trimAuthText_(r.yearsOfExp),
        keyInsights: trimAuthText_(r.keyInsights),
        tips:        trimAuthText_(r.tips),
        sharedAt:    trimAuthText_(r.sharedAt)
      };
    })
  };
}

// ── 通過ES閲覧 ──────────────────────────────────────────────────────────────

function readPassedES_(payload) {
  getActiveSessionOrThrow_(payload.sessionToken);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('es_bank');
  if (!sheet) return { status: 'ok', entries: [] };

  var rows = getSheetRecords_(sheet);
  var passed = rows.filter(function(r) {
    var result = String(r.result || '');
    return result.indexOf('通過') >= 0 || result.indexOf('合格') >= 0;
  });

  return {
    status: 'ok',
    entries: passed.map(function(r) {
      return {
        id:       trimAuthText_(r.id),
        company:  trimAuthText_(r.company),
        industry: trimAuthText_(r.industry),
        question: trimAuthText_(r.question),
        esText:   trimAuthText_(r.esText),
        result:   trimAuthText_(r.result)
      };
    })
  };
}

// ── NPS計測 ──────────────────────────────────────────────────────────────────

var NPS_SHEET_NAME_ = 'nps_responses';
var NPS_HEADERS_    = ['id', 'userId', 'score', 'comment', 'createdAt'];

function getOrCreateNPSSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NPS_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(NPS_SHEET_NAME_);
    sheet.appendRow(NPS_HEADERS_);
  }
  return sheet;
}

function submitNPS_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var score   = parseInt(payload.score, 10);
  if (isNaN(score) || score < 0 || score > 10) {
    throw new Error('スコアは0〜10の整数で入力してください。');
  }
  var comment = trimAuthText_(payload.comment || '');
  var sheet   = getOrCreateNPSSheet_();
  var newId   = 'nps_' + new Date().getTime().toString(36) + Utilities.getUuid().replace(/-/g, '').slice(0, 6);
  sheet.appendRow([newId, session.userId, score, comment, new Date().toISOString()]);
  return { status: 'ok', id: newId };
}

function readNPSSummary_(payload) {
  requireAdminSession_(payload);
  var sheet = getOrCreateNPSSheet_();
  var rows  = getSheetRecords_(sheet);
  if (!rows.length) return { status: 'ok', average: 0, count: 0 };
  var total = 0;
  rows.forEach(function(r) { total += parseInt(r.score, 10) || 0; });
  return {
    status: 'ok',
    average: Math.round((total / rows.length) * 10) / 10,
    count: rows.length
  };
}

// ── LINE Notify ──────────────────────────────────────────────────────────

function sendLineNotify_(token, message) {
  if (!token) return false;
  try {
    var resp = UrlFetchApp.fetch('https://notify-api.line.me/api/notify', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: { message: message },
      muteHttpExceptions: true
    });
    return resp.getResponseCode() === 200;
  } catch (e) {
    return false;
  }
}

function sendLineNotifyToUser_(userId, message) {
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var user = users.find(function(r) { return r.id === userId; });
  if (!user) return false;
  var token = trimAuthText_(user.lineNotifyToken);
  return sendLineNotify_(token, message);
}

function sendLineNotifyToAll_(message) {
  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var sent = 0;
  users.forEach(function(user) {
    var token = trimAuthText_(user.lineNotifyToken);
    if (token && sendLineNotify_(token, message)) sent++;
  });
  return sent;
}

// LINE Notifyトークン保存（プロフィール更新時に使う）
function authSaveLineNotifyToken_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var token = trimAuthText_(payload.lineNotifyToken);

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var idx = findRecordIndex_(users, function(r) { return r.id === session.userId; });
  assertAuth_(idx >= 0, 'アカウントが見つかりません。');

  // lineNotifyTokenカラムの位置を探す
  var headers = usersSheet.getRange(1, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  var tokenCol = headers.indexOf('lineNotifyToken');
  if (tokenCol >= 0) {
    usersSheet.getRange(idx + 2, tokenCol + 1).setValue(token);
  }

  // トークンが有効か検証
  if (token) {
    var valid = sendLineNotify_(token, '\n【慶應就活ナビ】LINE通知の設定が完了しました！');
    if (!valid) {
      return { status: 'error', message: 'LINE Notifyトークンが無効です。もう一度発行してください。' };
    }
  }

  return { status: 'ok', message: token ? 'LINE通知を有効にしました。' : 'LINE通知を無効にしました。' };
}

// 掲示板にコメントがついた時の通知
function notifyBoardComment_(postUserId, commenterName, postTitle) {
  sendLineNotifyToUser_(postUserId, '\n【就活ナビ】あなたの投稿にコメントがつきました\n投稿者: ' + commenterName + '\n内容: ' + postTitle);
}

// 週次ダイジェストのLINE版
function sendWeeklyDigestLine_() {
  var ss = getAuthSpreadsheet_();
  var now = new Date();
  var weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  var boardCount = 0, consultCount = 0, tlCount = 0;

  var boardSheet = ss.getSheetByName('es_board');
  if (boardSheet) {
    boardCount = getSheetRecords_(boardSheet).filter(function(r) {
      return new Date(r.createdAt) > weekAgo;
    }).length;
  }

  var consultSheet = ss.getSheetByName('consultation_board');
  if (consultSheet) {
    consultCount = getSheetRecords_(consultSheet).filter(function(r) {
      return new Date(r.createdAt) > weekAgo;
    }).length;
  }

  var tlSheet = ss.getSheetByName('activity_timeline');
  if (tlSheet) {
    tlCount = getSheetRecords_(tlSheet).filter(function(r) {
      return new Date(r.createdAt) > weekAgo;
    }).length;
  }

  var message = '\n📊 今週の就活ナビ\n'
    + '・ES添削掲示板: ' + boardCount + '件の新規投稿\n'
    + '・相談掲示板: ' + consultCount + '件\n'
    + '・タイムライン: ' + tlCount + '件の投稿\n'
    + '\n今すぐチェック →';

  return sendLineNotifyToAll_(message);
}

/* ===== 週次ダイジェストメール ===== */

function sendWeeklyDigest_(payload) {
  requireAdminSession_(payload);
  var result = sendWeeklyDigest();
  return { status: 'ok', sent: result };
}

function sendWeeklyDigestLineAction_(payload) {
  requireAdminSession_(payload);
  return { status: 'ok', sent: sendWeeklyDigestLine_() };
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

  // 各シートから過去7日間のデータをカウント
  var newBoardPosts = countRecentRows_('es_board', sevenDaysAgo);
  var newConsultations = countRecentRows_('consultations', sevenDaysAgo);
  var activeTimeline = countRecentRows_('timeline', sevenDaysAgo);

  users.forEach(function(user) {
    var email = trimAuthText_(user.email);
    if (!email || !validateEmail_(email)) return;

    var displayName = trimAuthText_(user.displayName) || 'メンバー';
    var siteUrl = getConfiguredSiteBaseUrl_() + '/members.html';

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

// ── 選考体験記 ──
var SEL_EXP_SHEET_ = 'selection_experiences';
var SEL_EXP_HEADERS_ = ['id','userId','displayName','stage','questions','atmosphere','tips','result','createdAt'];

function writeSelectionExperience_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);
  var text = trimAuthText_(payload.stage);
  assertAuth_(text, '選考段階を選択してください。');

  var usersSheet = getAuthSheet_(AUTH_SHEET_NAMES_.users);
  var users = getSheetRecords_(usersSheet);
  var user = users.find(function(r){return r.id===session.userId;});
  var displayName = user ? trimAuthText_(user.displayName||user.username) : '匿名';

  var ss = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(SEL_EXP_SHEET_);
  if(!sheet){sheet=ss.insertSheet(SEL_EXP_SHEET_);sheet.appendRow(SEL_EXP_HEADERS_);}

  var now = new Date().toISOString();
  var id = 'sexp_'+now.replace(/\D/g,'').slice(0,14)+'_'+Math.floor(Math.random()*10000);
  sheet.appendRow([id,session.userId,displayName,
    trimAuthText_(payload.stage),trimAuthText_(payload.questions),
    trimAuthText_(payload.atmosphere),trimAuthText_(payload.tips),
    trimAuthText_(payload.result),now]);
  return {status:'ok',id:id};
}

function readSelectionExperiences_(payload) {
  getActiveSessionOrThrow_(payload.sessionToken);
  var ss = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(SEL_EXP_SHEET_);
  if(!sheet) return {status:'ok',entries:[]};
  var rows = getSheetRecords_(sheet);
  rows.sort(function(a,b){return new Date(b.createdAt||0)-new Date(a.createdAt||0);});
  return {status:'ok',entries:rows.slice(0,30).map(function(r){
    return {id:trimAuthText_(r.id),displayName:trimAuthText_(r.displayName),
      stage:trimAuthText_(r.stage),questions:trimAuthText_(r.questions),
      atmosphere:trimAuthText_(r.atmosphere),tips:trimAuthText_(r.tips),
      result:trimAuthText_(r.result),createdAt:trimAuthText_(r.createdAt)};
  })};
}

// ── 管理者ダッシュボード ──
function getAdminStats_(payload) {
  requireAdminSession_(payload);
  ensureModerationSheets_();
  var ss = getAuthSpreadsheet_();
  var userCount = 0, esCount = 0, gkCount = 0, boardCount = 0, tlCount = 0, npsAvg = 0, npsCount = 0;
  var unverifiedUserCount = 0, activeSessionCount = 0, pendingReportCount = 0, hiddenContentCount = 0;

  var usersSheet = ss.getSheetByName('auth_users');
  if(usersSheet) {
    var users = getSheetRecords_(usersSheet);
    userCount = users.length;
    unverifiedUserCount = users.filter(function (user) {
      return !isEmailVerifiedRecord_(user);
    }).length;
  }
  var esSheet = ss.getSheetByName('es_bank');
  if(esSheet) esCount = Math.max(0, esSheet.getLastRow()-1);
  var gkSheet = ss.getSheetByName('gakuchika_bank');
  if(gkSheet) gkCount = Math.max(0, gkSheet.getLastRow()-1);
  var boardSheet = ss.getSheetByName('es_board');
  if(boardSheet) boardCount = Math.max(0, boardSheet.getLastRow()-1);
  var tlSheet = ss.getSheetByName('activity_timeline');
  if(tlSheet) tlCount = Math.max(0, tlSheet.getLastRow()-1);

  var npsSheet = ss.getSheetByName('nps_responses');
  if(npsSheet) {
    var npsRows = getSheetRecords_(npsSheet);
    npsCount = npsRows.length;
    if(npsCount>0) {
      var total = npsRows.reduce(function(s,r){return s+parseInt(r.score||0);},0);
      npsAvg = Math.round(total/npsCount*10)/10;
    }
  }

  var sessionsSheet = ss.getSheetByName(AUTH_SHEET_NAMES_.sessions);
  if (sessionsSheet) {
    activeSessionCount = getSheetRecords_(sessionsSheet).filter(function (row) {
      return String(row.active) !== '0' && new Date(row.expiresAt || 0).getTime() > Date.now();
    }).length;
  }

  pendingReportCount = getSheetRecords_(getAuthSheet_(MODERATION_REPORTS_SHEET_NAME_)).filter(function (row) {
    return trimAuthText_(row.status || 'open') === 'open';
  }).length;
  hiddenContentCount = getSheetRecords_(getAuthSheet_(MODERATION_CONTENT_STATE_SHEET_NAME_)).filter(function (row) {
    return trimAuthText_(row.state) === 'hidden';
  }).length;

  return {status:'ok',stats:{
    userCount:userCount,
    esCount:esCount,
    gkCount:gkCount,
    boardCount:boardCount,
    tlCount:tlCount,
    npsAvg:npsAvg,
    npsCount:npsCount,
    unverifiedUserCount:unverifiedUserCount,
    activeSessionCount:activeSessionCount,
    pendingReportCount:pendingReportCount,
    hiddenContentCount:hiddenContentCount
  }};
}

// ── 質問・相談掲示板 (Q&A Board) ──────────────────────────────────────────────

var QA_BOARD_SHEET_NAME_    = 'qa_board';
var QA_BOARD_MAX_POSTS_     = 100;
var QA_BOARD_CATEGORIES_    = ['ES・ガクチカ','面接','企業研究','その他'];
var QA_BOARD_HEADERS_       = ['id','userId','displayName','title','body','category','anonymous','createdAt','replies','acceptedReplyId'];

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
  var context  = getVerifiedActionContext_(payload, '質問投稿');
  var session  = context.session;
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

  sheet.appendRow([id, session.userId, displayName, title, body, category, anonymous ? 'true' : 'false', now, '[]', '']);

  return { status: 'ok', id: id };
}

function readQuestions_(payload) {
  var session = getActiveSessionOrThrow_(payload.sessionToken);

  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(QA_BOARD_SHEET_NAME_);
  if (!sheet) return { status: 'ok', questions: [] };

  var posts = getSheetRecords_(sheet);
  if (!posts.length) return { status: 'ok', questions: [] };
  var hiddenMap = getContentStateMap_('qa_question');
  var mutedUserIds = readMutedUserIdsForUser_(session.userId);

  posts.sort(function (a, b) {
    return new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime();
  });
  posts = posts.filter(function (post) {
    return !isContentHidden_(hiddenMap, post.id) && !mutedUserIds[trimAuthText_(post.userId)];
  });
  posts = posts.slice(0, QA_BOARD_MAX_POSTS_);

  var result = posts.map(function (post) {
    var replies = [];
    try { replies = JSON.parse(post.replies || '[]'); } catch (e) { replies = []; }
    replies = replies.filter(function (reply) {
      return !mutedUserIds[trimAuthText_(reply.userId)];
    });
    var acceptedReplyId = trimAuthText_(post.acceptedReplyId);
    replies.sort(function (a, b) {
      var aAccepted = trimAuthText_(a.id) === acceptedReplyId ? 0 : 1;
      var bAccepted = trimAuthText_(b.id) === acceptedReplyId ? 0 : 1;
      if (aAccepted !== bAccepted) return aAccepted - bAccepted;
      return new Date(a.createdAt || 0).getTime() - new Date(b.createdAt || 0).getTime();
    });
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
      acceptedReplyId: acceptedReplyId,
      canAcceptAnswer: trimAuthText_(post.userId) === trimAuthText_(session.userId),
      replyCount:  replies.length,
      replies:     replies.map(function (reply) {
        return {
          id: trimAuthText_(reply.id),
          userId: trimAuthText_(reply.userId),
          displayName: trimAuthText_(reply.displayName),
          body: trimAuthText_(reply.body),
          createdAt: trimAuthText_(reply.createdAt),
          accepted: trimAuthText_(reply.id) === acceptedReplyId
        };
      })
    };
  });

  return { status: 'ok', questions: result };
}

function postReply_(payload) {
  var context    = getVerifiedActionContext_(payload, '返信投稿');
  var session    = context.session;
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

  // replies column index
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var repliesCol = headers.indexOf('replies') + 1;
  if (repliesCol < 1) repliesCol = 9;
  sheet.getRange(idx + 2, repliesCol).setValue(JSON.stringify(replies));

  return { status: 'ok' };
}

function acceptQuestionReply_(payload) {
  var context = getVerifiedActionContext_(payload, 'ベストアンサー選択');
  var questionId = trimAuthText_(payload.questionId);
  var replyId = trimAuthText_(payload.replyId);
  assertAuth_(questionId, '質問IDが指定されていません。');
  assertAuth_(replyId, '返信IDが指定されていません。');

  var sheet = getOrCreateQABoardSheet_();
  var rows = getSheetRecords_(sheet);
  var idx = findRecordIndex_(rows, function (row) { return trimAuthText_(row.id) === questionId; });
  assertAuth_(idx >= 0, '質問が見つかりませんでした。');
  assertAuth_(trimAuthText_(rows[idx].userId) === trimAuthText_(context.session.userId), '自分の質問のみ操作できます。');

  var replies = [];
  try { replies = JSON.parse(rows[idx].replies || '[]'); } catch (error) { replies = []; }
  var exists = replies.some(function (reply) { return trimAuthText_(reply.id) === replyId; });
  assertAuth_(exists, '返信が見つかりませんでした。');

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var acceptedCol = headers.indexOf('acceptedReplyId') + 1;
  if (acceptedCol < 1) acceptedCol = 10;
  var nextReplyId = trimAuthText_(rows[idx].acceptedReplyId) === replyId ? '' : replyId;
  sheet.getRange(idx + 2, acceptedCol).setValue(nextReplyId);
  return { status: 'ok', acceptedReplyId: nextReplyId };
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
  var context      = getVerifiedActionContext_(payload, '体験記投稿');
  var session      = context.session;
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
  var session = getActiveSessionOrThrow_(payload.sessionToken);

  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(EXP_SHEET_NAME_);
  if (!sheet) return { status: 'ok', experiences: [] };

  var rows = getSheetRecords_(sheet);
  if (!rows.length) return { status: 'ok', experiences: [] };
  var hiddenMap = getContentStateMap_('experience');
  var mutedUserIds = readMutedUserIdsForUser_(session.userId);

  rows.sort(function (a, b) {
    return new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime();
  });
  rows = rows.filter(function (row) {
    return !isContentHidden_(hiddenMap, row.id) && !mutedUserIds[trimAuthText_(row.userId)];
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
  var context       = getVerifiedActionContext_(payload, '練習募集投稿');
  var session       = context.session;
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
  var hiddenMap = getContentStateMap_('practice_request');
  var mutedUserIds = readMutedUserIdsForUser_(session.userId);

  var ss    = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(PM_SHEET_NAME_);
  if (!sheet) return { status: 'ok', requests: [] };

  var rows = getSheetRecords_(sheet);
  if (!rows.length) return { status: 'ok', requests: [] };
  rows = rows.filter(function (row) {
    return !isContentHidden_(hiddenMap, row.id) && !mutedUserIds[trimAuthText_(row.userId)];
  });

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
    participants = participants.filter(function (participant) {
      return !mutedUserIds[trimAuthText_(participant.userId)];
    });

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
  var context   = getVerifiedActionContext_(payload, '練習募集への参加');
  var session   = context.session;
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
