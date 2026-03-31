(function () {
  const GAS_PROXY_URL = 'https://script.google.com/macros/s/AKfycbymaWYmys7-r9SnUk8o6tIx5SnNAdcKD2en4pi_tTnNtZF7QI_zf6_H6gvjAxQ__plYAA/exec';
  const SESSION_TOKEN_KEY = 'keio_navi_session_token_v1';
  const USER_CACHE_KEY = 'keio_navi_current_user_cache_v1';
  const LIKED_CACHE_KEY = 'keio_navi_liked_cache_v1';
  const REDIRECT_KEY = 'keio_navi_redirect_after_login_v1';
  const LEGACY_LIKED_KEY = 'keio_navi_liked_v2';

  let likedSyncQueue = Promise.resolve();

  function readJSON(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch (error) {
      return fallback;
    }
  }

  function writeJSON(key, value) {
    localStorage.setItem(key, JSON.stringify(value));
  }

  function trimText(value) {
    return String(value || '').trim();
  }

  function normalizeTextKey(value) {
    return trimText(value).normalize('NFKC').toLowerCase();
  }

  function normalizePreferredCompanies(source) {
    const raw = Array.isArray(source)
      ? source
      : source && typeof source === 'object'
        ? (Array.isArray(source.preferredCompanies)
          ? source.preferredCompanies
          : [source.preferredCompany1, source.preferredCompany2, source.preferredCompany3])
        : [];

    const unique = [];
    raw.forEach(item => {
      const name = trimText(item);
      if (!name) return;
      if (unique.some(existing => normalizeTextKey(existing) === normalizeTextKey(name))) return;
      unique.push(name);
    });
    return unique.slice(0, 3);
  }

  function currentRelativeUrl() {
    const fileName = window.location.pathname.split('/').pop() || 'index.html';
    return fileName + window.location.search + window.location.hash;
  }

  function setReturnTo(url) {
    if (url) sessionStorage.setItem(REDIRECT_KEY, url);
  }

  function getStoredReturnTo() {
    return sessionStorage.getItem(REDIRECT_KEY) || '';
  }

  function consumeReturnTo() {
    const url = getStoredReturnTo();
    sessionStorage.removeItem(REDIRECT_KEY);
    return url;
  }

  function getSessionToken() {
    return localStorage.getItem(SESSION_TOKEN_KEY) || '';
  }

  function setSessionToken(token) {
    localStorage.setItem(SESSION_TOKEN_KEY, token);
  }

  function clearSessionCache() {
    localStorage.removeItem(SESSION_TOKEN_KEY);
    localStorage.removeItem(USER_CACHE_KEY);
    localStorage.removeItem(LIKED_CACHE_KEY);
  }

  function normalizeUser(user) {
    if (!user || typeof user !== 'object') return null;
    const email = trimText(user.email || user.username || user.name);
    const displayName = trimText(user.displayName) || (email ? email.split('@')[0] : '');
    const desiredIndustry = trimText(user.desiredIndustry || user.desiredIndustries);
    if (!email) return null;

    const preferredCompanies = normalizePreferredCompanies(user);
    const lineName = trimText(user.lineName);
    const lineQrUrl = trimText(user.lineQrUrl);
    const hasLineQr = user.hasLineQr !== undefined ? Boolean(user.hasLineQr) : Boolean(lineQrUrl);

    return {
      id: trimText(user.id),
      email,
      emailKey: normalizeTextKey(user.emailKey || user.usernameKey || email),
      displayName,
      // 後方互換性
      username: displayName,
      usernameKey: normalizeTextKey(user.emailKey || user.usernameKey || email),
      desiredIndustry,
      preferredCompanies,
      preferredCompany1: preferredCompanies[0] || '',
      preferredCompany2: preferredCompanies[1] || '',
      preferredCompany3: preferredCompanies[2] || '',
      lineName,
      lineQrUrl,
      hasLineQr,
      referralCode: trimText(user.referralCode),
      createdAt: user.createdAt || '',
      updatedAt: user.updatedAt || '',
      sessionToken: trimText(user.sessionToken) || getSessionToken(),
      name: displayName,
      desiredIndustries: desiredIndustry,
    };
  }

  function buildProfilePayload(data) {
    const preferredCompanies = normalizePreferredCompanies(
      Array.isArray(data && data.preferredCompanies)
        ? data.preferredCompanies
        : [data && data.preferredCompany1, data && data.preferredCompany2, data && data.preferredCompany3]
    );

    return {
      email: trimText(data && data.email),
      displayName: trimText(data && (data.displayName || data.username || data.name)),
      // 後方互換性
      username: trimText(data && (data.displayName || data.username || data.name)),
      desiredIndustry: trimText(data && (data.desiredIndustry || data.desiredIndustries)),
      preferredCompanies,
      preferredCompany1: preferredCompanies[0] || '',
      preferredCompany2: preferredCompanies[1] || '',
      preferredCompany3: preferredCompanies[2] || '',
      lineName: trimText(data && data.lineName),
      lineQrDataUrl: trimText(data && data.lineQrDataUrl),
      lineQrFileName: trimText(data && data.lineQrFileName),
    };
  }

  function sanitizeLikedCompany(company) {
    if (!company || typeof company !== 'object') return null;
    return {
      name: trimText(company.name),
      industry: trimText(company.industry),
      business: trimText(company.business),
      strength: trimText(company.strength),
      motive: trimText(company.motive),
      process: trimText(company.process),
      url: trimText(company.url),
      alias: trimText(company.alias),
    };
  }

  function normalizeLikedCompanies(items) {
    return (Array.isArray(items) ? items : [])
      .map(sanitizeLikedCompany)
      .filter(item => item && item.name);
  }

  function getCurrentUser() {
    return normalizeUser(readJSON(USER_CACHE_KEY, null));
  }

  function getLikedCompanies() {
    const cached = normalizeLikedCompanies(readJSON(LIKED_CACHE_KEY, []));
    if (cached.length) return cached;

    const legacyLiked = normalizeLikedCompanies(readJSON(LEGACY_LIKED_KEY, []));
    if (!legacyLiked.length) return [];

    writeJSON(LIKED_CACHE_KEY, legacyLiked);
    localStorage.removeItem(LEGACY_LIKED_KEY);
    scheduleLikedSync(legacyLiked);
    return legacyLiked;
  }

  function setCachedSession(sessionToken, user, likedCompanies) {
    setSessionToken(sessionToken);
    writeJSON(USER_CACHE_KEY, normalizeUser(user));
    writeJSON(LIKED_CACHE_KEY, normalizeLikedCompanies(likedCompanies));
    localStorage.removeItem(LEGACY_LIKED_KEY);
  }

  async function postToGas(payload) {
    if (!GAS_PROXY_URL) {
      throw new Error('GAS_PROXY_URL が設定されていません。');
    }

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 30000);
    let response;
    try {
      response = await fetch(GAS_PROXY_URL, {
        method: 'POST',
        body: JSON.stringify(payload),
        redirect: 'follow',
        signal: controller.signal,
      });
    } catch (fetchError) {
      clearTimeout(timer);
      if (fetchError.name === 'AbortError') throw new Error('サーバーへの接続がタイムアウトしました。時間をおいて再試行してください。');
      throw fetchError;
    }
    clearTimeout(timer);

    const data = await response.json().catch(() => null);
    if (!data || data.status !== 'ok') {
      throw new Error((data && data.message) || 'サーバー保存に失敗しました。');
    }

    return data;
  }

  function validateEmail(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  }

  function validateProfile(profile, options) {
    const password = trimText(options && options.password);

    if (options && options.requireEmail) {
      if (!profile.email) return 'メールアドレスを入力してください。';
      if (!validateEmail(profile.email)) return '正しいメールアドレスの形式で入力してください。';
    }
    if (!profile.displayName && !profile.username) return '表示名を入力してください。';
    if (!profile.desiredIndustry) return '志望業界を選択してください。';
    if (!profile.preferredCompany1) return '第1志望を入力してください。';
    if (!profile.lineName) return 'LINE名を入力してください。';
    if (options && options.requireLineQr && !profile.lineQrDataUrl) {
      return 'LINE QRをアップロードしてください。';
    }
    if (options && options.requirePassword && password.length < 8) {
      return 'パスワードは8文字以上で設定してください。';
    }

    return '';
  }

  function scheduleLikedSync(nextLiked) {
    const sessionToken = getSessionToken();
    if (!sessionToken) return;

    const payload = normalizeLikedCompanies(nextLiked);
    likedSyncQueue = likedSyncQueue
      .then(() => postToGas({ action: 'authSetLikedCompanies', sessionToken, likedCompanies: payload }))
      .then(data => {
        if (data.likedCompanies) writeJSON(LIKED_CACHE_KEY, normalizeLikedCompanies(data.likedCompanies));
      })
      .catch(error => {
        console.warn('liked sync failed', error);
      });
  }

  async function createAccount(data) {
    const profile = buildProfilePayload(data);
    const error = validateProfile(profile, {
      requireEmail: true,
      requirePassword: true,
      password: data && data.password,
      requireLineQr: true,
    });
    if (error) return { ok: false, error };

    try {
      const result = await postToGas({
        action: 'authRegister',
        email: profile.email,
        displayName: profile.displayName,
        desiredIndustry: profile.desiredIndustry,
        preferredCompanies: profile.preferredCompanies,
        lineName: profile.lineName,
        lineQrDataUrl: profile.lineQrDataUrl,
        lineQrFileName: profile.lineQrFileName,
        password: trimText(data && data.password),
        referralCode: trimText(data && data.referralCode),
      });

      setCachedSession(result.sessionToken, result.user, result.likedCompanies || []);
      return { ok: true, user: getCurrentUser() };
    } catch (serverError) {
      return { ok: false, error: serverError.message };
    }
  }

  async function login(email, password) {
    const normalizedEmail = trimText(email);
    const normalizedPassword = trimText(password);

    if (!normalizedEmail || !normalizedPassword) {
      return { ok: false, error: 'メールアドレスとパスワードを入力してください。' };
    }

    try {
      const result = await postToGas({
        action: 'authLogin',
        email: normalizedEmail,
        password: normalizedPassword,
      });

      setCachedSession(result.sessionToken, result.user, result.likedCompanies || []);
      return { ok: true, user: getCurrentUser() };
    } catch (serverError) {
      return { ok: false, error: serverError.message };
    }
  }

  function logout() {
    const sessionToken = getSessionToken();
    clearSessionCache();

    if (!sessionToken) return;
    fetch(GAS_PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'authLogout', sessionToken }),
      redirect: 'follow',
    }).catch(() => {});
  }

  async function updateCurrentUser(data) {
    const sessionToken = getSessionToken();
    const currentUser = getCurrentUser();
    if (!sessionToken || !currentUser) return { ok: false, error: 'ログインが必要です。' };

    const profile = buildProfilePayload(data);
    const validationError = validateProfile(profile, {
      requirePassword: false,
      requireLineQr: !currentUser.hasLineQr,
    });
    if (validationError) return { ok: false, error: validationError };

    try {
      const result = await postToGas({
        action: 'authUpdateProfile',
        sessionToken,
        displayName: profile.displayName,
        desiredIndustry: profile.desiredIndustry,
        preferredCompanies: profile.preferredCompanies,
        lineName: profile.lineName,
        lineQrDataUrl: profile.lineQrDataUrl,
        lineQrFileName: profile.lineQrFileName,
      });

      writeJSON(USER_CACHE_KEY, normalizeUser(result.user));
      return { ok: true, user: getCurrentUser() };
    } catch (serverError) {
      return { ok: false, error: serverError.message };
    }
  }

  async function changePassword(currentPassword, nextPassword) {
    const sessionToken = getSessionToken();
    if (!sessionToken) return { ok: false, error: 'ログインが必要です。' };

    const oldPassword = trimText(currentPassword);
    const newPassword = trimText(nextPassword);
    if (!oldPassword || !newPassword) {
      return { ok: false, error: '現在のパスワードと新しいパスワードを入力してください。' };
    }
    if (newPassword.length < 8) {
      return { ok: false, error: '新しいパスワードは8文字以上で設定してください。' };
    }

    try {
      await postToGas({
        action: 'authChangePassword',
        sessionToken,
        currentPassword: oldPassword,
        nextPassword: newPassword,
      });
      return { ok: true };
    } catch (serverError) {
      return { ok: false, error: serverError.message };
    }
  }

  async function deleteAccount(password) {
    const sessionToken = getSessionToken();
    if (!sessionToken) return { ok: false, error: 'ログインが必要です。' };
    const pw = trimText(password);
    if (!pw) return { ok: false, error: 'パスワードを入力してください。' };

    try {
      const result = await postToGas({ action: 'authDeleteAccount', sessionToken, password: pw });
      clearSessionCache();
      return { ok: true, message: result.message };
    } catch (serverError) {
      return { ok: false, error: serverError.message };
    }
  }

  function setLikedCompanies(items) {
    const normalized = normalizeLikedCompanies(items);
    writeJSON(LIKED_CACHE_KEY, normalized);
    scheduleLikedSync(normalized);
  }

  function clearLikedCompanies() {
    setLikedCompanies([]);
  }

  function addLikedCompany(company) {
    const liked = getLikedCompanies();
    const normalizedCompany = sanitizeLikedCompany(company);
    if (!normalizedCompany || !normalizedCompany.name) return liked;
    if (liked.some(item => item.name === normalizedCompany.name)) return liked;

    const next = liked.concat(normalizedCompany);
    setLikedCompanies(next);
    return next;
  }

  async function requestPasswordReset(email) {
    const normalizedEmail = trimText(email);
    if (!normalizedEmail) return { ok: false, error: 'メールアドレスを入力してください。' };
    if (!validateEmail(normalizedEmail)) return { ok: false, error: '正しいメールアドレスの形式で入力してください。' };

    try {
      const result = await postToGas({ action: 'authRequestPasswordReset', email: normalizedEmail });
      return { ok: true, message: result.message };
    } catch (serverError) {
      return { ok: false, error: serverError.message };
    }
  }

  async function resetPassword(resetToken, newPassword) {
    const token = trimText(resetToken);
    const password = trimText(newPassword);
    if (!token) return { ok: false, error: 'リセットトークンが必要です。' };
    if (password.length < 8) return { ok: false, error: '新しいパスワードは8文字以上で設定してください。' };

    try {
      const result = await postToGas({ action: 'authResetPassword', resetToken: token, newPassword: password });
      return { ok: true, message: result.message };
    } catch (serverError) {
      return { ok: false, error: serverError.message };
    }
  }

  async function getReferralInfo() {
    const sessionToken = getSessionToken();
    if (!sessionToken) return { ok: false, error: 'ログインが必要です。' };

    try {
      const result = await postToGas({ action: 'authGetReferralInfo', sessionToken });
      return { ok: true, referralCode: result.referralCode, referralCount: result.referralCount };
    } catch (serverError) {
      return { ok: false, error: serverError.message };
    }
  }

  async function readMyProgress() {
    const sessionToken = getSessionToken();
    if (!sessionToken) return { ok: false, error: 'ログインが必要です。' };
    try {
      const result = await postToGas({ action: 'readMyProgress', sessionToken });
      return { ok: true, entries: result.entries };
    } catch (e) { return { ok: false, error: e.message }; }
  }

  async function writeProgress(entry) {
    const sessionToken = getSessionToken();
    if (!sessionToken) return { ok: false, error: 'ログインが必要です。' };
    try {
      const result = await postToGas({ action: 'writeProgress', sessionToken, ...entry });
      return { ok: true, id: result.id };
    } catch (e) { return { ok: false, error: e.message }; }
  }

  async function deleteProgress(id) {
    const sessionToken = getSessionToken();
    if (!sessionToken) return { ok: false, error: 'ログインが必要です。' };
    try {
      await postToGas({ action: 'deleteProgress', sessionToken, id });
      return { ok: true };
    } catch (e) { return { ok: false, error: e.message }; }
  }

  function getRedirectTarget(explicitReturnTo) {
    return explicitReturnTo || getStoredReturnTo() || 'account.html';
  }

  function requireAuth(options) {
    const user = getCurrentUser();
    if (user) return user;

    const settings = options || {};
    const returnTo = settings.returnTo || currentRelativeUrl();
    setReturnTo(returnTo);

    if (settings.redirect !== false) {
      const target = 'account.html?mode=login&returnTo=' + encodeURIComponent(returnTo);
      window.location.href = target;
    }

    return null;
  }

  function formatDate(dateString) {
    if (!dateString) return '';
    const date = new Date(dateString);
    if (isNaN(date)) return '';
    return date.toLocaleDateString('ja-JP', {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    });
  }

  // ── アクティビティトラッキング ──────────────────────
  const activityDebounceMap = {};
  const ACTIVITY_DEBOUNCE_MS = 30000; // 30秒

  function logActivity(event, page, feature) {
    if (!event) return;

    // デバウンス: 同一 event+page+feature は30秒に1回まで
    const debounceKey = [event, page || '', feature || ''].join('|');
    const now = Date.now();
    if (activityDebounceMap[debounceKey] && now - activityDebounceMap[debounceKey] < ACTIVITY_DEBOUNCE_MS) {
      return;
    }
    activityDebounceMap[debounceKey] = now;

    // Fire-and-forget: 結果を待たず、エラーも無視
    const payload = {
      action: 'logActivity',
      event: event,
      page: page || '',
      feature: feature || '',
      sessionToken: getSessionToken(),
      userAgent: navigator.userAgent || '',
    };

    fetch(GAS_PROXY_URL, {
      method: 'POST',
      body: JSON.stringify(payload),
      redirect: 'follow',
    }).catch(function () {});
  }

  // ── 通知機能 ──────────────────────────────────────
  const NOTIFICATION_PREFS_KEY = 'keio_navi_notification_prefs_v1';
  const NOTIFICATION_SENT_KEY = 'keio_navi_notification_sent_v1';

  function getNotificationStatus() {
    if (!('Notification' in window)) return 'unsupported';
    return Notification.permission; // 'granted' | 'denied' | 'default'
  }

  async function requestNotificationPermission() {
    if (!('Notification' in window)) {
      return { ok: false, error: 'このブラウザは通知に対応していません。' };
    }
    try {
      const permission = await Notification.requestPermission();
      return { ok: permission === 'granted', permission: permission };
    } catch (e) {
      return { ok: false, error: e.message };
    }
  }

  function getNotificationPrefs() {
    return readJSON(NOTIFICATION_PREFS_KEY, {
      enabled: false,
      deadlineReminder: true,
      newExperience: true,
      boardReply: true,
    });
  }

  function saveNotificationPrefs(prefs) {
    writeJSON(NOTIFICATION_PREFS_KEY, prefs);
  }

  function scheduleLocalNotification(title, body, delayMs, tag) {
    if (!('Notification' in window)) return null;
    if (Notification.permission !== 'granted') return null;

    const prefs = getNotificationPrefs();
    if (!prefs.enabled) return null;

    const timerId = setTimeout(function () {
      try {
        new Notification(title, {
          body: body,
          icon: 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect width="100" height="100" rx="20" fill="%230a1a3e"/><text x="50" y="65" font-size="50" text-anchor="middle" fill="%23c9a84c" font-family="serif">就</text></svg>',
          tag: tag || 'keio-navi-local',
          requireInteraction: false,
        });
      } catch (e) {
        // Service Worker環境でなければ直接Notificationを使う
        // SW経由の場合はregistration.showNotificationを使う
        if ('serviceWorker' in navigator && navigator.serviceWorker.ready) {
          navigator.serviceWorker.ready.then(function (reg) {
            reg.showNotification(title, {
              body: body,
              icon: 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect width="100" height="100" rx="20" fill="%230a1a3e"/><text x="50" y="65" font-size="50" text-anchor="middle" fill="%23c9a84c" font-family="serif">就</text></svg>',
              tag: tag || 'keio-navi-local',
            });
          }).catch(function () {});
        }
      }
    }, Math.max(0, delayMs || 0));

    return timerId;
  }

  function checkDeadlineReminders() {
    if (!('Notification' in window)) return;
    if (Notification.permission !== 'granted') return;

    var prefs = getNotificationPrefs();
    if (!prefs.enabled || !prefs.deadlineReminder) return;

    // 既に送信済みの通知タグを取得（日付ベースで重複防止）
    var today = new Date().toISOString().slice(0, 10);
    var sentMap = readJSON(NOTIFICATION_SENT_KEY, {});
    // 古い日付のエントリをクリーンアップ
    Object.keys(sentMap).forEach(function (key) {
      if (sentMap[key] < today) delete sentMap[key];
    });

    // 選考管理データ（localStorageに保存されている場合）を確認
    // members.htmlのinitProgressTrackerが使うキーを参照
    var progressEntries = [];
    try {
      // GAS経由で取得されたデータはmembers.htmlで管理されるが、
      // ローカルにキャッシュされた選考データも確認する
      var selectionData = readJSON('keio_navi_selection_cache_v1', []);
      if (Array.isArray(selectionData)) {
        selectionData.forEach(function (entry) {
          if (entry && entry.deadline) {
            progressEntries.push({
              company: entry.company || entry.name || '',
              deadline: entry.deadline,
              step: entry.step || entry.stage || '',
            });
          }
        });
      }
    } catch (e) {}

    // progress tracker のエントリもチェック
    try {
      var progressData = readJSON('keio_navi_progress_cache_v1', []);
      if (Array.isArray(progressData)) {
        progressData.forEach(function (entry) {
          if (entry && entry.deadline) {
            progressEntries.push({
              company: entry.company || entry.name || '',
              deadline: entry.deadline,
              step: entry.step || entry.stage || entry.status || '',
            });
          }
        });
      }
    } catch (e) {}

    if (!progressEntries.length) return;

    var now = new Date();
    var threeDaysLater = new Date(now.getTime() + 3 * 24 * 60 * 60 * 1000);

    progressEntries.forEach(function (entry) {
      if (!entry.deadline || !entry.company) return;
      var deadlineDate = new Date(entry.deadline);
      if (isNaN(deadlineDate.getTime())) return;

      var deadlineDateStr = deadlineDate.toISOString().slice(0, 10);
      var todayStr = now.toISOString().slice(0, 10);
      var diffMs = deadlineDate.getTime() - now.getTime();
      var diffDays = Math.ceil(diffMs / (24 * 60 * 60 * 1000));

      // 当日または3日以内
      if (diffDays < 0 || diffDays > 3) return;

      var notifTag = 'deadline-' + entry.company + '-' + deadlineDateStr;
      if (sentMap[notifTag]) return; // 既に通知済み

      var title, body;
      if (diffDays === 0) {
        title = '本日締切です！';
        body = entry.company + (entry.step ? '（' + entry.step + '）' : '') + 'の締切が本日です。忘れずに提出しましょう。';
      } else {
        title = '締切まであと' + diffDays + '日';
        body = entry.company + (entry.step ? '（' + entry.step + '）' : '') + 'の締切が' + deadlineDateStr + 'です。準備を進めましょう。';
      }

      // 少しずらして表示（重複を避けるため）
      scheduleLocalNotification(title, body, 2000, notifTag);
      sentMap[notifTag] = todayStr;
    });

    writeJSON(NOTIFICATION_SENT_KEY, sentMap);
  }

  window.KeioNaviAuth = {
    createAccount,
    login,
    logout,
    requireAuth,
    getCurrentUser,
    getUser: getCurrentUser,
    updateCurrentUser,
    changePassword,
    deleteAccount,
    requestPasswordReset,
    resetPassword,
    getReferralInfo,
    readMyProgress,
    writeProgress,
    deleteProgress,
    postToGas,
    setReturnTo,
    consumeReturnTo,
    getRedirectTarget,
    currentRelativeUrl,
    getLikedCompanies,
    setLikedCompanies,
    clearLikedCompanies,
    addLikedCompany,
    formatDate,
    getSessionToken,
    logActivity,
    getNotificationStatus,
    requestNotificationPermission,
    getNotificationPrefs,
    saveNotificationPrefs,
    scheduleLocalNotification,
    checkDeadlineReminders,
  };

  // ── ページ読み込み時に page_view を自動記録 ─────────
  try {
    const pageName = (window.location.pathname.split('/').pop() || 'index.html').replace(/\.html$/, '');
    logActivity('page_view', pageName, '');
  } catch (e) {}
})();
