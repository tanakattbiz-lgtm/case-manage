function getAuthState(sessionToken) {
  const users = listUserRecords_();
  if (users.length === 0) {
    return {
      requiresSetup: true,
      authenticated: false,
      settings: getPublicAuthSettings_(),
    };
  }

  const sessionContext = requireSession_(sessionToken, { silent: true, touch: true });
  if (!sessionContext) {
    return {
      requiresSetup: false,
      authenticated: false,
      settings: getPublicAuthSettings_(),
    };
  }

  return buildAuthPayload_(sessionToken, sessionContext.user, sessionContext.session);
}

function bootstrapAdmin(payload) {
  return runWithScriptLock_(function () {
    if (listUserRecords_().length > 0) {
      throwAppError_('BOOTSTRAP_LOCKED', '初回セットアップは完了しています。');
    }

    const input = payload || {};
    const name = normalizeString_(input.name);
    const email = normalizeEmail_(input.email);
    const password = String(input.password || '');
    const userAgent = normalizeUserAgent_(input.userAgent);

    validateIdentityInput_(name, email);
    assertPasswordPolicy_(password);

    const now = nowDateTimeStr_();
    const salt = createSecretChunk_();
    const userRecord = {
      ID: genId('USR'),
      '氏名': name,
      'メール': email,
      'ロール': USER_ROLES.admin,
      'パスワードハッシュ': hashPassword_(password, salt),
      'ソルト': salt,
      'ステータス': USER_STATUSES.active,
      '失敗回数': 0,
      'ロック期限': '',
      '最終ログイン': now,
      '最終PW変更': now,
      '登録日': now,
      '更新日': now,
    };

    const storedUser = appendUserRecord_(userRecord);
    const createdSession = createSessionForUser_(storedUser, userAgent);
    addAuditLogEntry_('bootstrap', 'SUCCESS', storedUser, '初回管理者を作成しました。', userAgent);
    return buildAuthPayload_(createdSession.token, storedUser, createdSession.session);
  });
}

function login(payload) {
  return runWithScriptLock_(function () {
    const input = payload || {};
    const email = normalizeEmail_(input.email);
    const password = String(input.password || '');
    const userAgent = normalizeUserAgent_(input.userAgent);

    if (!email || !password) {
      throwAppError_('LOGIN_REQUIRED', 'メールアドレスとパスワードを入力してください。');
    }
    if (listUserRecords_().length === 0) {
      throwAppError_('BOOTSTRAP_REQUIRED', '初回セットアップが必要です。');
    }

    const userRecord = findUserRecordByEmail_(email);
    const genericMessage = 'メールアドレスまたはパスワードが正しくありません。';
    if (!userRecord) {
      addAuditLogEntry_('login', 'FAIL', null, '存在しないメールアドレスでログインが試行されました。', userAgent, email);
      throwAppError_('LOGIN_FAILED', genericMessage);
    }
    if (normalizeUserStatus_(userRecord['ステータス']) !== USER_STATUSES.active) {
      addAuditLogEntry_('login', 'DENY', userRecord, '無効アカウントでログインが試行されました。', userAgent);
      throwAppError_('ACCOUNT_DISABLED', 'このアカウントは無効です。管理者へ連絡してください。');
    }

    const lockedUntil = parseDateValue_(userRecord['ロック期限']);
    if (lockedUntil && lockedUntil.getTime() > Date.now()) {
      addAuditLogEntry_('login', 'LOCKED', userRecord, 'ロック中アカウントでログインが試行されました。', userAgent);
      throwAppError_('ACCOUNT_LOCKED', 'ログインに失敗したため一時的にロックされています。');
    }

    const passwordVerification = verifyPassword_(password, userRecord);
    if (!passwordVerification.matched) {
      registerFailedLogin_(userRecord, userAgent);
      throwAppError_('LOGIN_FAILED', genericMessage);
    }

    const mutableUser = copyObject_(userRecord);
    mutableUser['失敗回数'] = 0;
    mutableUser['ロック期限'] = '';
    mutableUser['最終ログイン'] = nowDateTimeStr_();
    mutableUser['更新日'] = nowDateTimeStr_();
    if (passwordVerification.legacy) {
      const upgradeSalt = mutableUser['ソルト'] || createSecretChunk_();
      mutableUser['ソルト'] = upgradeSalt;
      mutableUser['パスワードハッシュ'] = hashPassword_(password, upgradeSalt);
    }
    const updatedUser = updateUserRecordById_(mutableUser.ID, mutableUser);
    const createdSession = createSessionForUser_(updatedUser, userAgent);
    addAuditLogEntry_('login', 'SUCCESS', updatedUser, 'ログインしました。', userAgent);
    return buildAuthPayload_(createdSession.token, updatedUser, createdSession.session);
  });
}

function logout(sessionToken) {
  return runWithScriptLock_(function () {
    if (!sessionToken) return { success: true };

    const sessionRecord = findSessionRecordByHash_(hashToken_(sessionToken));
    if (sessionRecord && sessionRecord['状態'] === 'active') {
      const mutableSession = copyObject_(sessionRecord);
      mutableSession['状態'] = 'logged_out';
      mutableSession['更新日'] = nowDateTimeStr_();
      updateSessionRecordById_(mutableSession.ID, mutableSession);

      const userRecord = findUserRecordById_(sessionRecord['ユーザーID']);
      if (userRecord) {
        addAuditLogEntry_('logout', 'SUCCESS', userRecord, 'ログアウトしました。', sessionRecord['ユーザーエージェント']);
      }
    }

    return { success: true };
  });
}

function changeMyPassword(sessionToken, currentPassword, newPassword) {
  return runWithScriptLock_(function () {
    const sessionContext = requireReadAccess_(sessionToken);
    const userRecord = findUserRecordById_(sessionContext.user.ID);
    if (!userRecord) {
      throwAppError_('USER_NOT_FOUND', 'ユーザーが見つかりません。');
    }

    const current = String(currentPassword || '');
    const next = String(newPassword || '');
    if (!verifyPassword_(current, userRecord).matched) {
      throwAppError_('PASSWORD_MISMATCH', '現在のパスワードが正しくありません。');
    }
    if (current === next) {
      throwAppError_('PASSWORD_REUSED', '新しいパスワードは現在のものと別にしてください。');
    }

    assertPasswordPolicy_(next);

    const mutableUser = copyObject_(userRecord);
    const salt = createSecretChunk_();
    mutableUser['ソルト'] = salt;
    mutableUser['パスワードハッシュ'] = hashPassword_(next, salt);
    mutableUser['最終PW変更'] = nowDateTimeStr_();
    mutableUser['失敗回数'] = 0;
    mutableUser['ロック期限'] = '';
    mutableUser['更新日'] = nowDateTimeStr_();

    const updatedUser = updateUserRecordById_(mutableUser.ID, mutableUser);
    revokeUserSessions_(updatedUser.ID, hashToken_(sessionToken));
    addAuditLogEntry_('password.change', 'SUCCESS', updatedUser, '本人がパスワードを変更しました。', sessionContext.session['ユーザーエージェント']);
    return { success: true };
  });
}

function requireReadAccess_(sessionToken) {
  const sessionContext = requireSession_(sessionToken, { touch: true });
  if (!getRolePermissions_(sessionContext.user['ロール']).canView) {
    throwAppError_('FORBIDDEN', '閲覧権限がありません。');
  }
  return sessionContext;
}

function requireEditAccess_(sessionToken) {
  const sessionContext = requireSession_(sessionToken, { touch: true });
  if (!getRolePermissions_(sessionContext.user['ロール']).canEdit) {
    addAuditLogEntry_('authorize', 'DENY', sessionContext.user, '編集権限のない操作がブロックされました。', sessionContext.session['ユーザーエージェント']);
    throwAppError_('FORBIDDEN', '編集権限がありません。');
  }
  return sessionContext;
}

function requireAdminAccess_(sessionToken) {
  const sessionContext = requireSession_(sessionToken, { touch: true });
  if (!getRolePermissions_(sessionContext.user['ロール']).canManageUsers) {
    addAuditLogEntry_('authorize', 'DENY', sessionContext.user, '管理者権限のない操作がブロックされました。', sessionContext.session['ユーザーエージェント']);
    throwAppError_('FORBIDDEN', '管理者権限がありません。');
  }
  return sessionContext;
}

function requireSession_(sessionToken, options) {
  const opts = options || {};
  if (!sessionToken) {
    if (opts.silent) return null;
    throwAppError_('AUTH_REQUIRED', 'ログインしてください。');
  }

  const tokenHash = hashToken_(sessionToken);
  const sessionRecord = findSessionRecordByHash_(tokenHash);
  if (!sessionRecord || sessionRecord['状態'] !== 'active') {
    if (opts.silent) return null;
    throwAppError_('AUTH_REQUIRED', 'セッションが無効です。再度ログインしてください。');
  }

  const expiresAt = parseDateValue_(sessionRecord['有効期限']);
  if (!expiresAt || expiresAt.getTime() <= Date.now()) {
    const expiredSession = copyObject_(sessionRecord);
    expiredSession['状態'] = 'expired';
    expiredSession['更新日'] = nowDateTimeStr_();
    updateSessionRecordById_(expiredSession.ID, expiredSession);
    if (opts.silent) return null;
    throwAppError_('AUTH_EXPIRED', 'セッションの有効期限が切れました。');
  }

  const userRecord = findUserRecordById_(sessionRecord['ユーザーID']);
  if (!userRecord || normalizeUserStatus_(userRecord['ステータス']) !== USER_STATUSES.active) {
    const revokedSession = copyObject_(sessionRecord);
    revokedSession['状態'] = 'revoked';
    revokedSession['更新日'] = nowDateTimeStr_();
    updateSessionRecordById_(revokedSession.ID, revokedSession);
    if (opts.silent) return null;
    throwAppError_('AUTH_REQUIRED', 'アカウントが無効です。');
  }

  const activeSession = opts.touch ? touchSessionIfNeeded_(sessionRecord) : sessionRecord;
  return {
    user: userRecord,
    session: activeSession,
  };
}

function buildAuthPayload_(token, userRecord, sessionRecord) {
  return {
    requiresSetup: false,
    authenticated: true,
    token: token,
    user: sanitizeUser_(userRecord),
    permissions: getRolePermissions_(userRecord['ロール']),
    session: {
      expiresAt: sessionRecord['有効期限'] || '',
      lastSeenAt: sessionRecord['最終アクセス'] || '',
    },
    settings: getPublicAuthSettings_(),
  };
}

function getPublicAuthSettings_() {
  return {
    passwordMinLength: FIXED_VALUES.auth.passwordMinLength,
    maxFailedAttempts: FIXED_VALUES.auth.maxFailedAttempts,
    lockoutMinutes: FIXED_VALUES.auth.lockoutMinutes,
    sessionTtlMinutes: FIXED_VALUES.auth.sessionTtlMinutes,
    roles: {
      admin: USER_ROLE_LABELS[USER_ROLES.admin],
      editor: USER_ROLE_LABELS[USER_ROLES.editor],
      viewer: USER_ROLE_LABELS[USER_ROLES.viewer],
    },
  };
}

function getRolePermissions_(role) {
  const normalizedRole = normalizeRole_(role);
  return {
    role: normalizedRole,
    roleLabel: USER_ROLE_LABELS[normalizedRole] || normalizedRole,
    canView: true,
    canEdit: normalizedRole === USER_ROLES.admin || normalizedRole === USER_ROLES.editor,
    canManageUsers: normalizedRole === USER_ROLES.admin,
  };
}

function sanitizeUser_(userRecord) {
  if (!userRecord) return null;
  const role = normalizeRole_(userRecord['ロール']);
  return {
    id: userRecord.ID,
    name: userRecord['氏名'] || '',
    email: userRecord['メール'] || '',
    role: role,
    roleLabel: USER_ROLE_LABELS[role] || role,
    status: normalizeUserStatus_(userRecord['ステータス']),
    failedCount: Number(userRecord['失敗回数']) || 0,
    lockedUntil: userRecord['ロック期限'] || '',
    lastLoginAt: userRecord['最終ログイン'] || '',
    passwordChangedAt: userRecord['最終PW変更'] || '',
    createdAt: userRecord['登録日'] || '',
    updatedAt: userRecord['更新日'] || '',
  };
}

function registerFailedLogin_(userRecord, userAgent) {
  const mutableUser = copyObject_(userRecord);
  const attempts = (Number(mutableUser['失敗回数']) || 0) + 1;
  const shouldLock = attempts >= FIXED_VALUES.auth.maxFailedAttempts;
  mutableUser['失敗回数'] = attempts;
  mutableUser['ロック期限'] = shouldLock
    ? formatDateTime_(addMinutes_(new Date(), FIXED_VALUES.auth.lockoutMinutes))
    : '';
  mutableUser['更新日'] = nowDateTimeStr_();
  updateUserRecordById_(mutableUser.ID, mutableUser);

  addAuditLogEntry_(
    'login',
    shouldLock ? 'LOCKED' : 'FAIL',
    mutableUser,
    shouldLock ? 'ログイン失敗が上限に達したためロックしました。' : 'ログイン失敗回数を更新しました。',
    userAgent
  );
}

function createSessionForUser_(userRecord, userAgent) {
  const token = createSessionToken_();
  const now = nowDateTimeStr_();
  const sessionRecord = {
    ID: genId('SES'),
    'ユーザーID': userRecord.ID,
    'トークンハッシュ': hashToken_(token),
    '状態': 'active',
    '有効期限': formatDateTime_(addMinutes_(new Date(), FIXED_VALUES.auth.sessionTtlMinutes)),
    '最終アクセス': now,
    'ユーザーエージェント': normalizeUserAgent_(userAgent),
    '登録日': now,
    '更新日': now,
  };

  const storedSession = appendSessionRecord_(sessionRecord);
  return {
    token: token,
    session: storedSession || sessionRecord,
  };
}

function revokeUserSessions_(userId, exceptTokenHash) {
  const skipHash = String(exceptTokenHash || '');
  listSessionRecordsByUserId_(userId).forEach(function (sessionRecord) {
    if (sessionRecord['状態'] !== 'active') return;
    if (skipHash && sessionRecord['トークンハッシュ'] === skipHash) return;

    const mutableSession = copyObject_(sessionRecord);
    mutableSession['状態'] = 'revoked';
    mutableSession['更新日'] = nowDateTimeStr_();
    updateSessionRecordById_(mutableSession.ID, mutableSession);
  });
}

function touchSessionIfNeeded_(sessionRecord) {
  const expiresAt = parseDateValue_(sessionRecord['有効期限']);
  if (!expiresAt) return sessionRecord;

  const remainingMinutes = Math.floor((expiresAt.getTime() - Date.now()) / 60000);
  if (remainingMinutes > FIXED_VALUES.auth.sessionRefreshThresholdMinutes) {
    return sessionRecord;
  }

  const mutableSession = copyObject_(sessionRecord);
  mutableSession['最終アクセス'] = nowDateTimeStr_();
  mutableSession['有効期限'] = formatDateTime_(addMinutes_(new Date(), FIXED_VALUES.auth.sessionTtlMinutes));
  mutableSession['更新日'] = nowDateTimeStr_();
  return updateSessionRecordById_(mutableSession.ID, mutableSession);
}

function addAuditLogEntry_(eventName, result, userRecord, detail, userAgent, emailFallback) {
  appendAuditRecord_({
    ID: genId('AUD'),
    '日時': nowDateTimeStr_(),
    'イベント': String(eventName || ''),
    '結果': String(result || ''),
    'ユーザーID': userRecord ? userRecord.ID || '' : '',
    '氏名': userRecord ? userRecord['氏名'] || '' : '',
    'メール': userRecord ? userRecord['メール'] || '' : String(emailFallback || ''),
    '詳細': String(detail || ''),
    'ユーザーエージェント': normalizeUserAgent_(userAgent),
  });
}

function validateIdentityInput_(name, email) {
  if (!name) {
    throwAppError_('NAME_REQUIRED', '氏名を入力してください。');
  }
  if (!email) {
    throwAppError_('EMAIL_REQUIRED', 'メールアドレスを入力してください。');
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throwAppError_('EMAIL_INVALID', 'メールアドレスの形式が正しくありません。');
  }
}

function assertPasswordPolicy_(password) {
  const value = String(password || '');
  if (value.length < FIXED_VALUES.auth.passwordMinLength) {
    throwAppError_('PASSWORD_WEAK', 'パスワードは ' + FIXED_VALUES.auth.passwordMinLength + ' 文字以上にしてください。');
  }
  if (!/[A-Za-z]/.test(value) || !/\d/.test(value)) {
    throwAppError_('PASSWORD_WEAK', 'パスワードには英字と数字を含めてください。');
  }
}

function normalizeRole_(role) {
  const value = normalizeString_(role).toLowerCase();
  if (value === USER_ROLES.admin || value === USER_ROLES.editor || value === USER_ROLES.viewer) {
    return value;
  }
  return USER_ROLES.viewer;
}

function normalizeUserStatus_(status) {
  return normalizeString_(status).toLowerCase() === USER_STATUSES.disabled
    ? USER_STATUSES.disabled
    : USER_STATUSES.active;
}

function createSessionToken_() {
  return sha256Hex_([Utilities.getUuid(), Date.now(), Math.random(), createSecretChunk_()].join(':'))
    + sha256Hex_([Utilities.getUuid(), Math.random(), Date.now()].join(':'));
}

function createSecretChunk_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function hashPassword_(password, salt) {
  return sha256Hex_([getAuthSecret_(), getAuthPepper_(), String(salt || ''), String(password || '')].join(':'));
}

function hashLegacyPassword_(password, salt) {
  return sha256Hex_([getAuthSecret_(), String(salt || ''), String(password || '')].join(':'));
}

// Keep pre-pepper accounts valid and upgrade them on the next successful login.
function verifyPassword_(password, userRecord) {
  const salt = userRecord ? userRecord['ソルト'] : '';
  const storedHash = userRecord ? String(userRecord['パスワードハッシュ'] || '') : '';
  if (!storedHash) {
    return { matched: false, legacy: false };
  }

  if (hashPassword_(password, salt) === storedHash) {
    return { matched: true, legacy: false };
  }

  if (hashLegacyPassword_(password, salt) === storedHash) {
    return { matched: true, legacy: true };
  }

  return { matched: false, legacy: false };
}

function hashToken_(token) {
  return sha256Hex_([getAuthSecret_(), 'token', String(token || '')].join(':'));
}

function getAuthSecret_() {
  const properties = PropertiesService.getScriptProperties();
  let secret = properties.getProperty(SCRIPT_PROPERTY_KEYS.authSecret);
  if (!secret) {
    secret = createSecretChunk_();
    properties.setProperty(SCRIPT_PROPERTY_KEYS.authSecret, secret);
  }
  return secret;
}

function getAuthPepper_() {
  const properties = PropertiesService.getScriptProperties();
  let pepper = properties.getProperty(SCRIPT_PROPERTY_KEYS.authPepper);
  if (!pepper) {
    pepper = createSecretChunk_();
    properties.setProperty(SCRIPT_PROPERTY_KEYS.authPepper, pepper);
  }
  return pepper;
}

function sha256Hex_(value) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value || ''),
    Utilities.Charset.UTF_8
  );
  return digest.map(function (byte) {
    const normalized = byte < 0 ? byte + 256 : byte;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');
}
