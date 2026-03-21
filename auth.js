(function () {
  const GAS_PROXY_URL = 'https://script.google.com/macros/s/AKfycbwMc8d4p_bLOJcMa8Vs_Egg5lsjl5UDwIOWT4cGFmNPWiQSdie-rT6oSYJ-pBiLCudzCw/exec';
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
    const username = trimText(user.username || user.name);
    const desiredIndustry = trimText(user.desiredIndustry || user.desiredIndustries);
    if (!username) return null;

    const preferredCompanies = normalizePreferredCompanies(user);
    const lineName = trimText(user.lineName);
    const lineQrUrl = trimText(user.lineQrUrl);
    const hasLineQr = user.hasLineQr !== undefined ? Boolean(user.hasLineQr) : Boolean(lineQrUrl);

    return {
      id: trimText(user.id),
      username,
      usernameKey: normalizeTextKey(user.usernameKey || username),
      desiredIndustry,
      preferredCompanies,
      preferredCompany1: preferredCompanies[0] || '',
      preferredCompany2: preferredCompanies[1] || '',
      preferredCompany3: preferredCompanies[2] || '',
      lineName,
      lineQrUrl,
      hasLineQr,
      createdAt: user.createdAt || '',
      updatedAt: user.updatedAt || '',
      name: username,
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
      username: trimText(data && (data.username || data.name)),
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

    const response = await fetch(GAS_PROXY_URL, {
      method: 'POST',
      body: JSON.stringify(payload),
      redirect: 'follow',
    });

    const data = await response.json().catch(() => null);
    if (!data || data.status !== 'ok') {
      throw new Error((data && data.message) || 'サーバー保存に失敗しました。');
    }

    return data;
  }

  function validateProfile(profile, options) {
    const password = trimText(options && options.password);

    if (!profile.username) return 'ユーザー名を入力してください。';
    if (!profile.desiredIndustry) return '志望業界を選択してください。';
    if (!profile.preferredCompany1) return '第1志望の企業名を入力してください。';
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
      requirePassword: true,
      password: data && data.password,
      requireLineQr: true,
    });
    if (error) return { ok: false, error };

    try {
      const result = await postToGas({
        action: 'authRegister',
        username: profile.username,
        desiredIndustry: profile.desiredIndustry,
        preferredCompanies: profile.preferredCompanies,
        lineName: profile.lineName,
        lineQrDataUrl: profile.lineQrDataUrl,
        lineQrFileName: profile.lineQrFileName,
        password: trimText(data && data.password),
      });

      setCachedSession(result.sessionToken, result.user, result.likedCompanies || []);
      return { ok: true, user: getCurrentUser() };
    } catch (serverError) {
      return { ok: false, error: serverError.message };
    }
  }

  async function login(username, password) {
    const normalizedUsername = trimText(username);
    const normalizedPassword = trimText(password);

    if (!normalizedUsername || !normalizedPassword) {
      return { ok: false, error: 'ユーザー名とパスワードを入力してください。' };
    }

    try {
      const result = await postToGas({
        action: 'authLogin',
        username: normalizedUsername,
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
        username: profile.username,
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

  window.KeioNaviAuth = {
    createAccount,
    login,
    logout,
    requireAuth,
    getCurrentUser,
    updateCurrentUser,
    changePassword,
    setReturnTo,
    consumeReturnTo,
    getRedirectTarget,
    currentRelativeUrl,
    getLikedCompanies,
    setLikedCompanies,
    clearLikedCompanies,
    addLikedCompany,
    formatDate,
  };
})();
