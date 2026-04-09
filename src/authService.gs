function getAuthState(sessionToken) {
  const users = getAllUsers_();
  if (users.length === 0) {
    return {
      requiresSetup: true,
      authenticated: false,
      settings: getPublicAuthSettings_(),
    };
  }

  const session = requireSession_(sessionToken, { silent: true, touch: true });
  if (!session) {
    return {
      requiresSetup: false,
      authenticated: false,
      settings: getPublicAuthSettings_(),
    };
  }

  return buildAuthPayload_(sessionToken, session.user, session.session);
}

function bootstrapAdmin(payload) {
  return withAuthLock_(function () {
    if (getAllUsers_().length > 0) {
      throwAuthError_('BOOTSTRAP_LOCKED', '初回セットアップは完了しています。');
    }

    const input = payload || {};
    const name = normalizeName_(input.name || input['氏名']);
    const email = normalizeEmail_(input.email || input['メール']);
    const password = String(input.password || '');
    const userAgent = normalizeUserAgent_(input.userAgent);

    validateIdentityInput_(name, email);
    assertPasswordPolicy_(password);

    const now = nowDateTimeStr_();
    const salt = createSecretChunk_();
    const user = {
      ID: genId('USR'),
      氏名: name,
      メール: email,
      ロール: AUTH_ROLES.admin,
      パスワードハッシュ: hashPassword_(password, salt),
      ソルト: salt,
      ステータス: AUTH_STATUSES.active,
      失敗回数: 0,
      ロック期限: '',
      最終ログイン: now,
      最終PW変更: now,
      登録日: now,
      更新日: now,
    };

    userSheet().appendRow(userRecordToRow_(user));
    const storedUser = findUserById_(user.ID);
    const created = createSessionForUser_(storedUser, userAgent);
    addAuditLog_('bootstrap', 'SUCCESS', storedUser, '初回管理者を作成しました。', userAgent);
    return buildAuthPayload_(created.token, storedUser, created.session);
  });
}

function login(payload) {
  return withAuthLock_(function () {
    const input = payload || {};
    const email = normalizeEmail_(input.email || input['メール']);
    const password = String(input.password || '');
    const userAgent = normalizeUserAgent_(input.userAgent);

    if (!email || !password) {
      throwAuthError_('LOGIN_REQUIRED', 'メールアドレスとパスワードを入力してください。');
    }
    if (getAllUsers_().length === 0) {
      throwAuthError_('BOOTSTRAP_REQUIRED', '初回セットアップが必要です。');
    }

    const user = findUserByEmail_(email);
    const genericMessage = 'メールアドレスまたはパスワードが正しくありません。';
    if (!user) {
      addAuditLog_('login', 'FAIL', null, '存在しないメールアドレスでログインが試行されました。', userAgent, email);
      throwAuthError_('LOGIN_FAILED', genericMessage);
    }
    if (user['ステータス'] !== AUTH_STATUSES.active) {
      addAuditLog_('login', 'DENY', user, '無効化されたアカウントでログインが試行されました。', userAgent);
      throwAuthError_('ACCOUNT_DISABLED', 'このアカウントは無効です。管理者に連絡してください。');
    }

    const lockedUntil = parseDateValue_(user['ロック期限']);
    if (lockedUntil && lockedUntil.getTime() > Date.now()) {
      addAuditLog_('login', 'LOCKED', user, 'ロック中のアカウントでログインが試行されました。', userAgent);
      throwAuthError_('ACCOUNT_LOCKED', 'ログイン失敗が続いたため、一時的にロックされています。');
    }

    const passwordHash = hashPassword_(password, user['ソルト']);
    if (passwordHash !== user['パスワードハッシュ']) {
      registerFailedLogin_(user, userAgent);
      throwAuthError_('LOGIN_FAILED', genericMessage);
    }

    const now = nowDateTimeStr_();
    const mutableUser = copySheetObject_(user);
    mutableUser['失敗回数'] = 0;
    mutableUser['ロック期限'] = '';
    mutableUser['最終ログイン'] = now;
    mutableUser['更新日'] = now;
    writeUserRecord_(mutableUser._row, mutableUser);

    const storedUser = findUserById_(user['ID']);
    const created = createSessionForUser_(storedUser, userAgent);
    addAuditLog_('login', 'SUCCESS', storedUser, 'ログインしました。', userAgent);
    return buildAuthPayload_(created.token, storedUser, created.session);
  });
}

function logout(sessionToken) {
  return withAuthLock_(function () {
    if (!sessionToken) return { success: true };

    const tokenHash = hashToken_(sessionToken);
    const session = findSessionByHash_(tokenHash);
    if (session && session['状態'] === 'active') {
      const mutableSession = copySheetObject_(session);
      mutableSession['状態'] = 'logged_out';
      mutableSession['更新日'] = nowDateTimeStr_();
      writeSessionRecord_(mutableSession._row, mutableSession);

      const user = findUserById_(session['ユーザーID']);
      if (user) addAuditLog_('logout', 'SUCCESS', user, 'ログアウトしました。', session['ユーザーエージェント']);
    }

    return { success: true };
  });
}

function getUsers(sessionToken) {
  requireAdminAccess_(sessionToken);
  const roleOrder = {};
  roleOrder[AUTH_ROLES.admin] = 0;
  roleOrder[AUTH_ROLES.editor] = 1;
  roleOrder[AUTH_ROLES.viewer] = 2;

  return getAllUsers_()
    .sort(function (a, b) {
      const ra = roleOrder[String(a['ロール'] || '')] != null ? roleOrder[String(a['ロール'] || '')] : 99;
      const rb = roleOrder[String(b['ロール'] || '')] != null ? roleOrder[String(b['ロール'] || '')] : 99;
      if (ra !== rb) return ra - rb;
      return String(a['氏名'] || '').localeCompare(String(b['氏名'] || ''), 'ja');
    })
    .map(sanitizeUser_);
}

function saveUser(sessionToken, payload) {
  return withAuthLock_(function () {
    const adminSession = requireAdminAccess_(sessionToken);
    const input = payload || {};
    const existingId = String(input.id || input['ID'] || '');
    const existingUser = existingId ? findUserById_(existingId) : null;
    if (existingId && !existingUser) {
      throwAuthError_('USER_NOT_FOUND', '対象ユーザーが見つかりません。');
    }

    const name = normalizeName_(input.name || input['氏名'] || (existingUser ? existingUser['氏名'] : ''));
    const email = normalizeEmail_(input.email || input['メール'] || (existingUser ? existingUser['メール'] : ''));
    const role = normalizeRole_(input.role || input['ロール'] || (existingUser ? existingUser['ロール'] : AUTH_ROLES.viewer));
    const status = normalizeStatus_(input.status || input['ステータス'] || (existingUser ? existingUser['ステータス'] : AUTH_STATUSES.active));
    const password = String(input.password || '');
    const unlock = Boolean(input.unlock);

    validateIdentityInput_(name, email);
    assertEmailUnique_(email, existingId || null);
    if (!existingUser && !password) {
      throwAuthError_('PASSWORD_REQUIRED', '新規ユーザーには初期パスワードが必要です。');
    }
    if (password) {
      assertPasswordPolicy_(password);
    }

    const now = nowDateTimeStr_();
    if (existingUser) {
      ensureAdminRetention_(existingUser, role, status);

      const mutableUser = copySheetObject_(existingUser);
      mutableUser['氏名'] = name;
      mutableUser['メール'] = email;
      mutableUser['ロール'] = role;
      mutableUser['ステータス'] = status;
      mutableUser['更新日'] = now;

      if (unlock) {
        mutableUser['失敗回数'] = 0;
        mutableUser['ロック期限'] = '';
      }

      if (password) {
        const salt = createSecretChunk_();
        mutableUser['ソルト'] = salt;
        mutableUser['パスワードハッシュ'] = hashPassword_(password, salt);
        mutableUser['最終PW変更'] = now;
        mutableUser['失敗回数'] = 0;
        mutableUser['ロック期限'] = '';
      }

      writeUserRecord_(mutableUser._row, mutableUser);
      if (status !== AUTH_STATUSES.active || password) {
        revokeUserSessions_(mutableUser['ID'], '');
      }

      addAuditLog_(
        password ? 'user.password_reset' : 'user.update',
        'SUCCESS',
        adminSession.user,
        'target=' + email + ', role=' + role + ', status=' + status,
        ''
      );
      return { success: true, user: sanitizeUser_(findUserById_(mutableUser['ID'])) };
    }

    const salt = createSecretChunk_();
    const user = {
      ID: genId('USR'),
      氏名: name,
      メール: email,
      ロール: role,
      パスワードハッシュ: hashPassword_(password, salt),
      ソルト: salt,
      ステータス: status,
      失敗回数: 0,
      ロック期限: '',
      最終ログイン: '',
      最終PW変更: now,
      登録日: now,
      更新日: now,
    };

    userSheet().appendRow(userRecordToRow_(user));
    addAuditLog_('user.create', 'SUCCESS', adminSession.user, 'target=' + email + ', role=' + role, '');
    return { success: true, user: sanitizeUser_(findUserById_(user.ID)) };
  });
}

function unlockUser(sessionToken, userId) {
  return withAuthLock_(function () {
    const adminSession = requireAdminAccess_(sessionToken);
    const user = findUserById_(userId);
    if (!user) {
      throwAuthError_('USER_NOT_FOUND', '対象ユーザーが見つかりません。');
    }

    const mutableUser = copySheetObject_(user);
    mutableUser['失敗回数'] = 0;
    mutableUser['ロック期限'] = '';
    mutableUser['更新日'] = nowDateTimeStr_();
    writeUserRecord_(mutableUser._row, mutableUser);

    addAuditLog_('user.unlock', 'SUCCESS', adminSession.user, 'target=' + user['メール'], '');
    return { success: true, user: sanitizeUser_(findUserById_(user['ID'])) };
  });
}

function getAuditLogs(sessionToken) {
  requireAdminAccess_(sessionToken);
  return sheetToObjects(auditSheet())
    .sort(function (a, b) {
      const da = parseDateValue_(a['日時']);
      const db = parseDateValue_(b['日時']);
      return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
    })
    .slice(0, AUTH_SETTINGS.auditLogLimit)
    .map(function (row) {
      return {
        id: row['ID'],
        timestamp: row['日時'] || '',
        event: row['イベント'] || '',
        result: row['結果'] || '',
        userId: row['ユーザーID'] || '',
        name: row['氏名'] || '',
        email: row['メール'] || '',
        detail: row['詳細'] || '',
        userAgent: row['ユーザーエージェント'] || '',
      };
    });
}

function changeMyPassword(sessionToken, currentPassword, newPassword) {
  return withAuthLock_(function () {
    const session = requireReadAccess_(sessionToken);
    const user = findUserById_(session.user['ID']);
    if (!user) {
      throwAuthError_('USER_NOT_FOUND', 'ユーザーが見つかりません。');
    }

    const current = String(currentPassword || '');
    const next = String(newPassword || '');
    if (hashPassword_(current, user['ソルト']) !== user['パスワードハッシュ']) {
      throwAuthError_('PASSWORD_MISMATCH', '現在のパスワードが正しくありません。');
    }
    if (current === next) {
      throwAuthError_('PASSWORD_REUSED', '新しいパスワードは現在のものと変えてください。');
    }

    assertPasswordPolicy_(next);

    const mutableUser = copySheetObject_(user);
    const now = nowDateTimeStr_();
    const salt = createSecretChunk_();
    mutableUser['ソルト'] = salt;
    mutableUser['パスワードハッシュ'] = hashPassword_(next, salt);
    mutableUser['最終PW変更'] = now;
    mutableUser['失敗回数'] = 0;
    mutableUser['ロック期限'] = '';
    mutableUser['更新日'] = now;
    writeUserRecord_(mutableUser._row, mutableUser);

    revokeUserSessions_(mutableUser['ID'], hashToken_(sessionToken));
    addAuditLog_('password.change', 'SUCCESS', mutableUser, '本人がパスワードを変更しました。', session.session['ユーザーエージェント']);
    return { success: true };
  });
}

function requireReadAccess_(sessionToken) {
  const session = requireSession_(sessionToken, { touch: true });
  if (!getRolePermissions_(session.user['ロール']).canView) {
    throwAuthError_('FORBIDDEN', '閲覧権限がありません。');
  }
  return session;
}

function requireEditAccess_(sessionToken) {
  const session = requireSession_(sessionToken, { touch: true });
  if (!getRolePermissions_(session.user['ロール']).canEdit) {
    addAuditLog_('authorize', 'DENY', session.user, '編集権限がない操作がブロックされました。', session.session['ユーザーエージェント']);
    throwAuthError_('FORBIDDEN', '編集権限がありません。');
  }
  return session;
}

function requireAdminAccess_(sessionToken) {
  const session = requireSession_(sessionToken, { touch: true });
  if (!getRolePermissions_(session.user['ロール']).canManageUsers) {
    addAuditLog_('authorize', 'DENY', session.user, '管理者専用操作がブロックされました。', session.session['ユーザーエージェント']);
    throwAuthError_('FORBIDDEN', '管理者権限がありません。');
  }
  return session;
}

function requireSession_(sessionToken, options) {
  const opts = options || {};
  if (!sessionToken) {
    if (opts.silent) return null;
    throwAuthError_('AUTH_REQUIRED', 'ログインが必要です。');
  }

  const tokenHash = hashToken_(sessionToken);
  const session = findSessionByHash_(tokenHash);
  if (!session || session['状態'] !== 'active') {
    if (opts.silent) return null;
    throwAuthError_('AUTH_REQUIRED', 'ログイン状態を確認できませんでした。');
  }

  const expiresAt = parseDateValue_(session['有効期限']);
  if (!expiresAt || expiresAt.getTime() <= Date.now()) {
    const mutableSession = copySheetObject_(session);
    mutableSession['状態'] = 'expired';
    mutableSession['更新日'] = nowDateTimeStr_();
    writeSessionRecord_(mutableSession._row, mutableSession);
    if (opts.silent) return null;
    throwAuthError_('AUTH_EXPIRED', 'セッションの有効期限が切れました。');
  }

  const user = findUserById_(session['ユーザーID']);
  if (!user || user['ステータス'] !== AUTH_STATUSES.active) {
    const mutableSession = copySheetObject_(session);
    mutableSession['状態'] = 'revoked';
    mutableSession['更新日'] = nowDateTimeStr_();
    writeSessionRecord_(mutableSession._row, mutableSession);
    if (opts.silent) return null;
    throwAuthError_('AUTH_REQUIRED', 'アカウントが無効です。');
  }

  if (opts.touch) {
    touchSessionIfNeeded_(session);
  }

  return {
    user: user,
    session: findSessionByHash_(tokenHash) || session,
  };
}

function buildAuthPayload_(token, user, session) {
  return {
    requiresSetup: false,
    authenticated: true,
    token: token,
    user: sanitizeUser_(user),
    permissions: getRolePermissions_(user['ロール']),
    session: {
      expiresAt: session['有効期限'] || '',
      lastSeenAt: session['最終アクセス'] || '',
    },
    settings: getPublicAuthSettings_(),
  };
}

function getPublicAuthSettings_() {
  return {
    passwordMinLength: AUTH_SETTINGS.passwordMinLength,
    maxFailedAttempts: AUTH_SETTINGS.maxFailedAttempts,
    lockoutMinutes: AUTH_SETTINGS.lockoutMinutes,
    sessionTtlMinutes: AUTH_SETTINGS.sessionTtlMinutes,
    roles: {
      admin: AUTH_ROLE_LABELS[AUTH_ROLES.admin],
      editor: AUTH_ROLE_LABELS[AUTH_ROLES.editor],
      viewer: AUTH_ROLE_LABELS[AUTH_ROLES.viewer],
    },
  };
}

function getRolePermissions_(role) {
  const normalizedRole = normalizeRole_(role || AUTH_ROLES.viewer);
  return {
    role: normalizedRole,
    roleLabel: AUTH_ROLE_LABELS[normalizedRole] || normalizedRole,
    canView: true,
    canEdit: normalizedRole === AUTH_ROLES.admin || normalizedRole === AUTH_ROLES.editor,
    canManageUsers: normalizedRole === AUTH_ROLES.admin,
  };
}

function sanitizeUser_(user) {
  if (!user) return null;
  return {
    row: user._row,
    id: user['ID'],
    name: user['氏名'] || '',
    email: user['メール'] || '',
    role: normalizeRole_(user['ロール'] || AUTH_ROLES.viewer),
    roleLabel: AUTH_ROLE_LABELS[normalizeRole_(user['ロール'] || AUTH_ROLES.viewer)] || '',
    status: normalizeStatus_(user['ステータス'] || AUTH_STATUSES.active),
    failedCount: Number(user['失敗回数']) || 0,
    lockedUntil: user['ロック期限'] || '',
    lastLoginAt: user['最終ログイン'] || '',
    passwordChangedAt: user['最終PW変更'] || '',
    createdAt: user['登録日'] || '',
    updatedAt: user['更新日'] || '',
  };
}

function registerFailedLogin_(user, userAgent) {
  const mutableUser = copySheetObject_(user);
  const attempts = (Number(mutableUser['失敗回数']) || 0) + 1;
  const shouldLock = attempts >= AUTH_SETTINGS.maxFailedAttempts;
  mutableUser['失敗回数'] = attempts;
  mutableUser['ロック期限'] = shouldLock ? dateTimeStr_(addMinutes_(new Date(), AUTH_SETTINGS.lockoutMinutes)) : '';
  mutableUser['更新日'] = nowDateTimeStr_();
  writeUserRecord_(mutableUser._row, mutableUser);

  const detail = shouldLock
    ? 'ログイン失敗が ' + attempts + ' 回に達したためロックしました。'
    : 'ログイン失敗 ' + attempts + ' 回目。';
  addAuditLog_('login', shouldLock ? 'LOCKED' : 'FAIL', mutableUser, detail, userAgent);
}

function createSessionForUser_(user, userAgent) {
  const token = createSessionToken_();
  const now = nowDateTimeStr_();
  const session = {
    ID: genId('SES'),
    ユーザーID: user['ID'],
    トークンハッシュ: hashToken_(token),
    状態: 'active',
    有効期限: dateTimeStr_(addMinutes_(new Date(), AUTH_SETTINGS.sessionTtlMinutes)),
    最終アクセス: now,
    ユーザーエージェント: userAgent,
    登録日: now,
    更新日: now,
  };

  sessionSheet().appendRow(sessionRecordToRow_(session));
  return {
    token: token,
    session: findSessionByHash_(session['トークンハッシュ']) || session,
  };
}

function revokeUserSessions_(userId, exceptHash) {
  const targetHash = String(exceptHash || '');
  getAllSessions_().forEach(function (session) {
    if (session['ユーザーID'] !== userId) return;
    if (session['状態'] !== 'active') return;
    if (targetHash && session['トークンハッシュ'] === targetHash) return;

    const mutableSession = copySheetObject_(session);
    mutableSession['状態'] = 'revoked';
    mutableSession['更新日'] = nowDateTimeStr_();
    writeSessionRecord_(mutableSession._row, mutableSession);
  });
}

function touchSessionIfNeeded_(session) {
  const expiresAt = parseDateValue_(session['有効期限']);
  if (!expiresAt) return;

  const remainingMinutes = Math.floor((expiresAt.getTime() - Date.now()) / 60000);
  if (remainingMinutes > AUTH_SETTINGS.sessionRefreshThresholdMinutes) return;

  const mutableSession = copySheetObject_(session);
  mutableSession['最終アクセス'] = nowDateTimeStr_();
  mutableSession['有効期限'] = dateTimeStr_(addMinutes_(new Date(), AUTH_SETTINGS.sessionTtlMinutes));
  mutableSession['更新日'] = nowDateTimeStr_();
  writeSessionRecord_(mutableSession._row, mutableSession);
}

function addAuditLog_(eventName, result, user, detail, userAgent, emailFallback) {
  auditSheet().appendRow(auditRecordToRow_({
    ID: genId('AUD'),
    日時: nowDateTimeStr_(),
    イベント: eventName || '',
    結果: result || '',
    ユーザーID: user ? user['ID'] || '' : '',
    氏名: user ? user['氏名'] || '' : '',
    メール: user ? user['メール'] || '' : (emailFallback || ''),
    詳細: detail || '',
    ユーザーエージェント: normalizeUserAgent_(userAgent),
  }));
}

function ensureAdminRetention_(currentUser, nextRole, nextStatus) {
  const willRemainAdmin = nextRole === AUTH_ROLES.admin && nextStatus === AUTH_STATUSES.active;
  if (willRemainAdmin) return;
  if (currentUser['ロール'] !== AUTH_ROLES.admin || currentUser['ステータス'] !== AUTH_STATUSES.active) return;

  const activeAdminCount = getAllUsers_().filter(function (user) {
    return user['ID'] !== currentUser['ID']
      && user['ロール'] === AUTH_ROLES.admin
      && user['ステータス'] === AUTH_STATUSES.active;
  }).length;

  if (activeAdminCount === 0) {
    throwAuthError_('LAST_ADMIN', '最後の有効な管理者は降格または無効化できません。');
  }
}

function assertEmailUnique_(email, ignoreUserId) {
  const normalizedEmail = normalizeEmail_(email);
  const duplicated = getAllUsers_().some(function (user) {
    return normalizeEmail_(user['メール']) === normalizedEmail && user['ID'] !== ignoreUserId;
  });
  if (duplicated) {
    throwAuthError_('EMAIL_DUPLICATED', '同じメールアドレスのユーザーが既に存在します。');
  }
}

function validateIdentityInput_(name, email) {
  if (!name) {
    throwAuthError_('NAME_REQUIRED', '氏名を入力してください。');
  }
  if (!email) {
    throwAuthError_('EMAIL_REQUIRED', 'メールアドレスを入力してください。');
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throwAuthError_('EMAIL_INVALID', 'メールアドレスの形式が正しくありません。');
  }
}

function assertPasswordPolicy_(password) {
  const value = String(password || '');
  if (value.length < AUTH_SETTINGS.passwordMinLength) {
    throwAuthError_(
      'PASSWORD_WEAK',
      'パスワードは ' + AUTH_SETTINGS.passwordMinLength + ' 文字以上にしてください。'
    );
  }
  if (!/[A-Za-z]/.test(value) || !/\d/.test(value)) {
    throwAuthError_('PASSWORD_WEAK', 'パスワードには英字と数字を両方含めてください。');
  }
}

function findUserByEmail_(email) {
  const normalizedEmail = normalizeEmail_(email);
  return getAllUsers_().find(function (user) {
    return normalizeEmail_(user['メール']) === normalizedEmail;
  }) || null;
}

function findUserById_(userId) {
  return getAllUsers_().find(function (user) {
    return user['ID'] === userId;
  }) || null;
}

function findSessionByHash_(tokenHash) {
  return getAllSessions_().find(function (session) {
    return session['トークンハッシュ'] === tokenHash;
  }) || null;
}

function getAllUsers_() {
  return sheetToObjects(userSheet());
}

function getAllSessions_() {
  return sheetToObjects(sessionSheet());
}

function writeUserRecord_(rowNumber, user) {
  userSheet().getRange(Number(rowNumber), 1, 1, USER_COLS.length).setValues([userRecordToRow_(user)]);
}

function writeSessionRecord_(rowNumber, session) {
  sessionSheet().getRange(Number(rowNumber), 1, 1, SESSION_COLS.length).setValues([sessionRecordToRow_(session)]);
}

function userRecordToRow_(user) {
  return USER_COLS.map(function (col) {
    return user[col] == null ? '' : user[col];
  });
}

function sessionRecordToRow_(session) {
  return SESSION_COLS.map(function (col) {
    return session[col] == null ? '' : session[col];
  });
}

function auditRecordToRow_(audit) {
  return AUDIT_COLS.map(function (col) {
    return audit[col] == null ? '' : audit[col];
  });
}

function normalizeName_(name) {
  return String(name || '').trim();
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function normalizeRole_(role) {
  const value = String(role || '').trim().toLowerCase();
  if (value === AUTH_ROLES.admin || value === AUTH_ROLES.editor || value === AUTH_ROLES.viewer) {
    return value;
  }
  return AUTH_ROLES.viewer;
}

function normalizeStatus_(status) {
  const value = String(status || '').trim().toLowerCase();
  if (value === AUTH_STATUSES.disabled) return AUTH_STATUSES.disabled;
  return AUTH_STATUSES.active;
}

function normalizeUserAgent_(userAgent) {
  return String(userAgent || '').trim().slice(0, 300);
}

function throwAuthError_(code, message) {
  throw new Error(String(code || 'APP_ERROR') + ': ' + String(message || 'エラーが発生しました。'));
}

function createSessionToken_() {
  return sha256Hex_([Utilities.getUuid(), Date.now(), Math.random(), createSecretChunk_()].join(':'))
    + sha256Hex_([Utilities.getUuid(), Math.random(), Date.now()].join(':'));
}

function createSecretChunk_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function hashPassword_(password, salt) {
  return sha256Hex_([getAuthSecret_(), String(salt || ''), String(password || '')].join(':'));
}

function hashToken_(token) {
  return sha256Hex_([getAuthSecret_(), 'token', String(token || '')].join(':'));
}

function getAuthSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('AUTH_SECRET');
  if (!secret) {
    secret = createSecretChunk_();
    props.setProperty('AUTH_SECRET', secret);
  }
  return secret;
}

function sha256Hex_(value) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value || ''),
    Utilities.Charset.UTF_8
  );

  return digest.map(function (b) {
    const byte = b < 0 ? b + 256 : b;
    return ('0' + byte.toString(16)).slice(-2);
  }).join('');
}

function parseDateValue_(value) {
  if (!value) return null;
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  if (typeof value === 'number') {
    const numericDate = new Date(value);
    return isNaN(numericDate.getTime()) ? null : numericDate;
  }

  const normalized = String(value).trim();
  if (!normalized) return null;

  const matched = normalized.match(
    /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
  );
  if (matched) {
    const parsedManual = new Date(
      Number(matched[1]),
      Number(matched[2]) - 1,
      Number(matched[3]),
      Number(matched[4] || 0),
      Number(matched[5] || 0),
      Number(matched[6] || 0)
    );
    return isNaN(parsedManual.getTime()) ? null : parsedManual;
  }

  const parsed = new Date(normalized.replace(/\//g, '-'));
  return isNaN(parsed.getTime()) ? null : parsed;
}

function nowDateTimeStr_() {
  return dateTimeStr_(new Date());
}

function dateTimeStr_(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

function addMinutes_(date, minutes) {
  return new Date(date.getTime() + Number(minutes || 0) * 60000);
}

function withAuthLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function copySheetObject_(obj) {
  return Object.assign({}, obj || {});
}
