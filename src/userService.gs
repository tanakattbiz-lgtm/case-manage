function getUsers(sessionToken, query) {
  requireAdminAccess_(sessionToken);
  const normalizedQuery = normalizeUserListQuery_(query);
  const allItems = listUserDtos_();
  const filteredItems = allItems.filter(function (user) {
    if (normalizedQuery.role && user.role !== normalizedQuery.role) return false;
    if (normalizedQuery.status && user.status !== normalizedQuery.status) return false;
    if (!normalizedQuery.query) return true;

    const joined = [
      user.name,
      user.email,
      user.roleLabel,
    ].join(' ').toLowerCase();
    return joined.indexOf(normalizedQuery.query.toLowerCase()) >= 0;
  });
  const paged = paginateItems_(filteredItems, normalizedQuery.page, normalizedQuery.pageSize);

  return {
    items: paged.items,
    pagination: paged.pagination,
    filters: normalizedQuery,
    summary: buildUserListSummary_(filteredItems, allItems.length),
    fetchedAt: nowDateTimeStr_(),
  };
}

function getUserDetail(sessionToken, userId) {
  requireAdminAccess_(sessionToken);
  const targetId = normalizeString_(userId);
  const user = listUserDtos_().find(function (item) {
    return item.id === targetId;
  }) || null;

  if (!user) {
    throwAppError_('USER_NOT_FOUND', 'ユーザーが見つかりません。');
  }

  const audits = listAuditDtos_().filter(function (audit) {
    return audit.userId === user.id;
  }).slice(0, FIXED_VALUES.lists.detailRelatedLimit);

  return {
    user: user,
    recentAudits: audits,
  };
}

function saveUser(sessionToken, payload) {
  return runWithScriptLock_(function () {
    const adminContext = requireAdminAccess_(sessionToken);
    const input = payload || {};
    const existingId = normalizeString_(input.id);
    const existing = existingId ? findUserRecordById_(existingId) : null;
    if (existingId && !existing) {
      throwAppError_('USER_NOT_FOUND', 'ユーザーが見つかりません。');
    }

    const name = normalizeString_(input.name || (existing ? existing['氏名'] : ''));
    const email = normalizeEmail_(input.email || (existing ? existing['メール'] : ''));
    const role = normalizeRole_(input.role || (existing ? existing['ロール'] : USER_ROLES.viewer));
    const status = normalizeUserStatus_(input.status || (existing ? existing['ステータス'] : USER_STATUSES.active));
    const password = String(input.password || '');
    const unlock = Boolean(input.unlock);

    validateIdentityInput_(name, email);
    assertUserEmailAvailable_(email, existingId);
    if (!existing && !password) {
      throwAppError_('PASSWORD_REQUIRED', '新規ユーザーには初期パスワードが必要です。');
    }
    if (password) {
      assertPasswordPolicy_(password);
    }

    if (existing) {
      ensureAdminRetention_(existing, role, status);
      const mutableUser = copyObject_(existing);
      mutableUser['氏名'] = name;
      mutableUser['メール'] = email;
      mutableUser['ロール'] = role;
      mutableUser['ステータス'] = status;
      mutableUser['更新日'] = nowDateTimeStr_();

      if (unlock) {
        mutableUser['失敗回数'] = 0;
        mutableUser['ロック期限'] = '';
      }
      if (password) {
        const salt = createSecretChunk_();
        mutableUser['ソルト'] = salt;
        mutableUser['パスワードハッシュ'] = hashPassword_(password, salt);
        mutableUser['最終PW変更'] = nowDateTimeStr_();
        mutableUser['失敗回数'] = 0;
        mutableUser['ロック期限'] = '';
      }

      const updatedUser = updateUserRecordById_(mutableUser.ID, mutableUser);
      if (status !== USER_STATUSES.active || password) {
        revokeUserSessions_(updatedUser.ID, '');
      }
      addAuditLogEntry_('user.update', 'SUCCESS', adminContext.user, 'target=' + email + ', role=' + role + ', status=' + status, '');
      return { success: true, user: sanitizeUser_(updatedUser) };
    }

    const salt = createSecretChunk_();
    const now = nowDateTimeStr_();
    const userRecord = {
      ID: genId('USR'),
      '氏名': name,
      'メール': email,
      'ロール': role,
      'パスワードハッシュ': hashPassword_(password, salt),
      'ソルト': salt,
      'ステータス': status,
      '失敗回数': 0,
      'ロック期限': '',
      '最終ログイン': '',
      '最終PW変更': now,
      '登録日': now,
      '更新日': now,
    };
    const storedUser = appendUserRecord_(userRecord);
    addAuditLogEntry_('user.create', 'SUCCESS', adminContext.user, 'target=' + email + ', role=' + role, '');
    return { success: true, user: sanitizeUser_(storedUser) };
  });
}

function unlockUser(sessionToken, userId) {
  return runWithScriptLock_(function () {
    const adminContext = requireAdminAccess_(sessionToken);
    const userRecord = findUserRecordById_(String(userId || ''));
    if (!userRecord) {
      throwAppError_('USER_NOT_FOUND', 'ユーザーが見つかりません。');
    }

    const mutableUser = copyObject_(userRecord);
    mutableUser['失敗回数'] = 0;
    mutableUser['ロック期限'] = '';
    mutableUser['更新日'] = nowDateTimeStr_();
    const updatedUser = updateUserRecordById_(mutableUser.ID, mutableUser);
    addAuditLogEntry_('user.unlock', 'SUCCESS', adminContext.user, 'target=' + updatedUser['メール'], '');
    return { success: true, user: sanitizeUser_(updatedUser) };
  });
}

function getAuditLogs(sessionToken, query) {
  requireAdminAccess_(sessionToken);
  const normalizedQuery = normalizeAuditListQuery_(query);
  const allItems = listAuditDtos_();
  const filteredItems = allItems.filter(function (audit) {
    if (normalizedQuery.result && audit.result !== normalizedQuery.result) return false;
    if (!normalizedQuery.query) return true;

    const joined = [
      audit.event,
      audit.result,
      audit.name,
      audit.email,
      audit.detail,
    ].join(' ').toLowerCase();
    return joined.indexOf(normalizedQuery.query.toLowerCase()) >= 0;
  });
  const paged = paginateItems_(filteredItems, normalizedQuery.page, normalizedQuery.pageSize);

  return {
    items: paged.items,
    pagination: paged.pagination,
    filters: normalizedQuery,
    summary: {
      totalItems: allItems.length,
      filteredItems: filteredItems.length,
    },
    fetchedAt: nowDateTimeStr_(),
  };
}

function listUserDtos_() {
  const roleOrder = {};
  roleOrder[USER_ROLES.admin] = 0;
  roleOrder[USER_ROLES.editor] = 1;
  roleOrder[USER_ROLES.viewer] = 2;

  return listUserRecords_()
    .sort(function (a, b) {
      const leftRole = roleOrder[normalizeRole_(a['ロール'])] != null ? roleOrder[normalizeRole_(a['ロール'])] : 99;
      const rightRole = roleOrder[normalizeRole_(b['ロール'])] != null ? roleOrder[normalizeRole_(b['ロール'])] : 99;
      if (leftRole !== rightRole) return leftRole - rightRole;
      return String(a['氏名'] || '').localeCompare(String(b['氏名'] || ''), 'ja');
    })
    .map(sanitizeUser_);
}

function listAuditDtos_() {
  return listAuditRecords_()
    .sort(function (a, b) {
      const left = parseDateValue_(a['日時']);
      const right = parseDateValue_(b['日時']);
      return (right ? right.getTime() : 0) - (left ? left.getTime() : 0);
    })
    .slice(0, FIXED_VALUES.auth.auditLogLimit)
    .map(function (record) {
      return {
        id: record.ID,
        timestamp: record['日時'] || '',
        event: record['イベント'] || '',
        result: record['結果'] || '',
        userId: record['ユーザーID'] || '',
        name: record['氏名'] || '',
        email: record['メール'] || '',
        detail: record['詳細'] || '',
        userAgent: record['ユーザーエージェント'] || '',
      };
    });
}

function normalizeUserListQuery_(query) {
  const input = query || {};
  return {
    page: normalizePositiveInteger_(input.page, 1, 1, 999),
    pageSize: normalizePositiveInteger_(
      input.pageSize,
      FIXED_VALUES.lists.defaultPageSize,
      1,
      FIXED_VALUES.lists.maxPageSize
    ),
    query: normalizeString_(input.query),
    role: normalizeString_(input.role),
    status: normalizeString_(input.status),
    sort: {
      field: 'role,name',
      direction: 'asc',
      label: '権限順',
    },
  };
}

function normalizeAuditListQuery_(query) {
  const input = query || {};
  return {
    page: normalizePositiveInteger_(input.page, 1, 1, 999),
    pageSize: normalizePositiveInteger_(
      input.pageSize,
      FIXED_VALUES.lists.defaultPageSize,
      1,
      FIXED_VALUES.lists.maxPageSize
    ),
    query: normalizeString_(input.query),
    result: normalizeString_(input.result).toUpperCase(),
    sort: {
      field: 'timestamp',
      direction: 'desc',
      label: '新着順',
    },
  };
}

function buildUserListSummary_(filteredItems, totalItems) {
  return {
    totalItems: Number(totalItems) || 0,
    filteredItems: filteredItems.length,
    activeCount: filteredItems.filter(function (user) {
      return user.status === USER_STATUSES.active;
    }).length,
    lockedCount: filteredItems.filter(function (user) {
      return isLockedUserDto_(user);
    }).length,
    adminCount: filteredItems.filter(function (user) {
      return user.role === USER_ROLES.admin;
    }).length,
  };
}

function isLockedUserDto_(user) {
  const lockedUntil = parseDateValue_(user && user.lockedUntil);
  return Boolean(lockedUntil && lockedUntil.getTime() > Date.now());
}

function ensureAdminRetention_(currentUser, nextRole, nextStatus) {
  const remainsActiveAdmin = nextRole === USER_ROLES.admin && nextStatus === USER_STATUSES.active;
  if (remainsActiveAdmin) return;
  if (normalizeRole_(currentUser['ロール']) !== USER_ROLES.admin || normalizeUserStatus_(currentUser['ステータス']) !== USER_STATUSES.active) {
    return;
  }

  const otherActiveAdmins = listUserRecords_().filter(function (record) {
    return record.ID !== currentUser.ID
      && normalizeRole_(record['ロール']) === USER_ROLES.admin
      && normalizeUserStatus_(record['ステータス']) === USER_STATUSES.active;
  }).length;

  if (otherActiveAdmins === 0) {
    throwAppError_('LAST_ADMIN', '最後の有効な管理者を無効化できません。');
  }
}

function assertUserEmailAvailable_(email, ignoreUserId) {
  const duplicated = listUserRecords_().some(function (record) {
    return normalizeEmail_(record['メール']) === normalizeEmail_(email)
      && record.ID !== ignoreUserId;
  });
  if (duplicated) {
    throwAppError_('EMAIL_DUPLICATED', '同じメールアドレスのユーザーが既に存在します。');
  }
}
