const STORAGE_KEYS = {
  currentUser: 'lsh15_current_user_clean',
  tasks: 'lsh15_tasks_clean',
  modReports: 'lsh15_mod_reports_clean',
  users: 'lsh13_users'
};

const FIREBASE_COLLECTIONS = {
  users: 'users',
  tasks: 'tasks_clean',
  modReports: 'modReports_clean'
};
const FIREBASE_LOGIN_DOMAIN = 'laya.local';
const FIREBASE_STORAGE_FOLDERS = {
  taskMedia: 'task-media',
  modMedia: 'mod-media'
};

const OVERDUE_MINUTES = {
  Low: 180,
  Medium: 120,
  High: 60,
  Urgent: 30
};

const DEPARTMENTS = ['FO', 'FB', 'HK', 'Engineering'];
const STATUSES = ['New', 'Accepted', 'In Progress', 'Waiting Supervisor Review', 'Returned for Rework', 'Pending FO Closure', 'Closed'];
const PRIORITIES = ['Urgent', 'High', 'Medium', 'Low'];
const MOD_STATUSES = ['Open', 'In Progress', 'Resolved', 'Closed'];
const MOD_MEDIA_INLINE_LIMIT = 2 * 1024 * 1024;
const DEPARTMENT_COLORS = {
  FO: '#315f9f',
  FB: '#b08b57',
  HK: '#2f8f5b',
  Engineering: '#c64747'
};
const STATUS_COLORS = {
  New: '#3178c6',
  Accepted: '#7c4dcb',
  'In Progress': '#d88a1d',
  'Waiting Supervisor Review': '#b26b00',
  'Returned for Rework': '#c64747',
  'Pending FO Closure': '#2f8f5b',
  Closed: '#6b736d'
};
const PRIORITY_COLORS = {
  Urgent: '#c64747',
  High: '#d88a1d',
  Medium: '#3178c6',
  Low: '#6b736d'
};

const defaultUsers = [
  { employeeId: '11025', password: '1234', name: 'Noi', role: 'CGM', department: 'FO' },
  { employeeId: '11026', password: '1234', name: 'Pim', role: 'DOR', department: 'FO' },
  { employeeId: '12001', password: '1234', name: 'Anna', role: 'FO Staff', department: 'FO' },
  { employeeId: '12002', password: '1234', name: 'Beam', role: 'FO Staff', department: 'FO' },
  { employeeId: '22001', password: '1234', name: 'Joy', role: 'Housekeeping Manager', department: 'HK' },
  { employeeId: '22018', password: '1234', name: 'May', role: 'HK Staff', department: 'HK' },
  { employeeId: '22019', password: '1234', name: 'Fon', role: 'HK Staff', department: 'HK' },
  { employeeId: '33001', password: '1234', name: 'Lek', role: 'ENG Manager', department: 'Engineering' },
  { employeeId: '33007', password: '1234', name: 'Art', role: 'ENG Staff', department: 'Engineering' },
  { employeeId: '33008', password: '1234', name: 'Tom', role: 'ENG Staff', department: 'Engineering' },
  { employeeId: '44001', password: '1234', name: 'Mint', role: 'FB Manager', department: 'FB' },
  { employeeId: '44005', password: '1234', name: 'Ben', role: 'FB Staff', department: 'FB' },
  { employeeId: '44006', password: '1234', name: 'Nan', role: 'FB Staff', department: 'FB' },
  { employeeId: '99001', password: '1234', name: 'Mook', role: 'MOD', department: 'Hotel Ops' }
];
let users = [];

function createSeedTask(config) {
  const createdAt = new Date(Date.now() - (config.createdAtOffsetMin || 0) * 60000).toISOString();
  const logs = (config.logs || []).map((log) => ({
    action: log.action,
    note: log.note,
    byName: log.byName,
    byDepartment: log.byDepartment,
    createdAt: new Date(Date.now() - (log.offsetMin || 0) * 60000).toISOString()
  })).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  let assignedAt = '';
  let startedAt = '';
  let doneAt = '';
  let closedAt = '';

  logs.forEach((log) => {
    if (log.action === 'Accepted') assignedAt = log.createdAt;
    if (log.action === 'In Progress') startedAt = log.createdAt;
    if (log.action === 'Done' || log.action === 'Pending FO Closure' || log.action === 'Submitted to FO') doneAt = log.createdAt;
    if (log.action === 'Closed') closedAt = log.createdAt;
  });

  return normalizeTask({
    id: config.ticketNo,
    ticketNo: config.ticketNo,
    location: config.location,
    department: config.department,
    category: config.category,
    priority: config.priority,
    subject: config.subject,
    detail: config.detail,
    status: config.status,
    openedByName: config.openedByName,
    openedByDepartment: config.openedByDepartment,
    assignedToName: config.assignedToName || '',
    sourceType: config.sourceType || '',
    sourceReference: config.sourceReference || '',
    mediaAttachments: config.mediaAttachments || [],
    createdAt,
    updatedAt: logs[0]?.createdAt || createdAt,
    assignedAt,
    startedAt,
    doneAt,
    closedAt,
    logs
  });
}

const seedTasks = [];

const seedModReports = [];
const extraSeedModReports = [];

const state = {
  currentUser: null,
  currentView: 'home',
  previousView: 'home',
  currentTaskId: null,
  taskStatusFilter: 'All',
  taskSearch: '',
  highOnly: false,
  taskContext: 'all',
  teamQuickFilter: 'all',
  teamAssigneeFilter: '__ALL__',
  historyPreset: 'today',
  reportPreset: 'today',
  reportViewMode: 'daily',
  currentModReportId: null,
  modFilter: 'all',
  modSearch: '',
  modDraftMedia: [],
  detailDraftMedia: [],
  pendingRegistrationProfile: null
};

document.addEventListener('DOMContentLoaded', init);

async function init() {
  bindEvents();
  setupChipGroup(document.getElementById('department-chips'), els.taskDepartment, 'FO', false, renderCreateAssignmentState);
  setupChipGroup(document.getElementById('priority-chips'), els.taskPriority, 'Medium');
  setupChipGroup(document.getElementById('mod-area-chips'), els.modArea, 'Lobby');
  setupChipGroup(document.getElementById('mod-department-chips'), els.modDepartment, 'Engineering');
  setupChipGroup(document.getElementById('mod-priority-chips'), els.modPriority, 'High');
  setDatePreset('history', 'today', true);
  setDatePreset('report', 'today', true);
  await initializeBackend();
}

function bindEvents() {
  els.loginBtn.addEventListener('click', onLogin);
  els.loginPassword.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') onLogin();
  });
  if (els.openRegisterBtn) els.openRegisterBtn.addEventListener('click', () => toggleRegisterPanel(true));
  if (els.registerCancelBtn) els.registerCancelBtn.addEventListener('click', () => toggleRegisterPanel(false));
  if (els.registerSubmitBtn) els.registerSubmitBtn.addEventListener('click', onRegister);
  if (els.registerPasswordConfirm) els.registerPasswordConfirm.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') onRegister();
  });
  els.logoutBtn.addEventListener('click', logout);
  els.backBtn.addEventListener('click', onBack);
  els.navHome.addEventListener('click', () => showPage(getHomeNavPage()));
  els.navTasks.addEventListener('click', () => { clearTaskContext(); showPage('tasks'); });
  els.navCreate.addEventListener('click', () => showPage('create'));
  els.navHistory.addEventListener('click', () => showPage('history'));
  els.navMod.addEventListener('click', () => showPage('mod'));
  els.navReports.addEventListener('click', () => showPage('reports'));
  els.homeCreateBtn.addEventListener('click', () => showPage(canCreateTasks() ? 'create' : getDefaultLandingPage()));
  els.homeViewTasksBtn.addEventListener('click', openDepartmentTasksFromHome);
  els.cancelCreateBtn.addEventListener('click', () => showPage(getDefaultLandingPage()));
  els.createTaskForm.addEventListener('submit', onCreateTask);
  els.taskStatusTabs.addEventListener('click', onTaskTabClick);
  els.taskFilterHigh.addEventListener('click', toggleHighFilter);
  els.taskContextClear.addEventListener('click', () => { clearTaskContext(); renderTaskList(); updateTopbar('tasks'); });
  els.teamQuickFilters.addEventListener('click', onTeamQuickFilterClick);
  els.teamAssigneeFilter.addEventListener('change', onTeamAssigneeFilterChange);
  els.teamFilterClear.addEventListener('click', resetSupervisorTaskBoardFilters);
  els.supervisorWorkloadCards.addEventListener('click', onWorkloadCardClick);
  els.supervisorBalanceActions.addEventListener('click', onSupervisorBalanceActionClick);
  els.supervisorBoardLanes.addEventListener('click', onSupervisorBoardClick);
  els.homeSection1Btn.addEventListener('click', () => openTaskPreset(getHomeConfig().sections[0].preset));
  els.homeSection2Btn.addEventListener('click', () => openTaskPreset(getHomeConfig().sections[1].preset));
  els.homeSection3Btn.addEventListener('click', () => openTaskPreset(getHomeConfig().sections[2].preset));
  els.tasksSearch.addEventListener('input', onTasksSearch);
  els.tasksList.addEventListener('click', onTaskCardClick);
  els.homeSection1List.addEventListener('click', onTaskCardClick);
  els.homeSection2List.addEventListener('click', onTaskCardClick);
  els.homeSection3List.addEventListener('click', onTaskCardClick);
  els.dashboardOverdue.addEventListener('click', onTaskCardClick);
  els.historyResults.addEventListener('click', onTaskCardClick);
  els.reportResults.addEventListener('click', onTaskCardClick);
  els.detailActions.addEventListener('click', onDetailActionClick);
  els.detailAssignBtn.addEventListener('click', assignCurrentTask);
  els.detailAssignFocusBtn.addEventListener('click', () => els.detailNoteInput.focus());
  els.detailSaveNoteBtn.addEventListener('click', saveDetailNote);
  if (els.teamAdminForm) els.teamAdminForm.addEventListener('submit', onAddTeamMember);
  if (els.teamMemberResetBtn) els.teamMemberResetBtn.addEventListener('click', resetTeamAdminForm);
  if (els.teamAdminList) els.teamAdminList.addEventListener('click', onTeamAdminListClick);
  if (els.detailWorkMediaInput) els.detailWorkMediaInput.addEventListener('change', onDetailWorkMediaSelected);
  if (els.detailWorkMediaPreview) els.detailWorkMediaPreview.addEventListener('click', onDetailWorkMediaPreviewClick);
  if (els.detailWorkMediaSaveBtn) els.detailWorkMediaSaveBtn.addEventListener('click', saveDetailWorkMedia);
  if (els.detailWorkMediaClearBtn) els.detailWorkMediaClearBtn.addEventListener('click', clearDetailWorkMediaDraft);
  els.historyPresets.addEventListener('click', onHistoryPresetClick);
  els.reportPresets.addEventListener('click', onReportPresetClick);
  els.reportViewModes.addEventListener('click', onReportViewModeClick);
  els.historyStartDate.addEventListener('change', renderHistoryPage);
  els.historyEndDate.addEventListener('change', renderHistoryPage);
  els.historySearch.addEventListener('input', renderHistoryPage);
  els.historyDepartment.addEventListener('change', renderHistoryPage);
  els.historyStatus.addEventListener('change', renderHistoryPage);
  els.historyPriority.addEventListener('change', renderHistoryPage);
  els.reportStartDate.addEventListener('change', renderReportsPage);
  els.reportEndDate.addEventListener('change', renderReportsPage);
  els.reportDepartment.addEventListener('change', renderReportsPage);
  els.reportStatus.addEventListener('change', renderReportsPage);
  els.reportExportCsv.addEventListener('click', exportReportCsv);
  els.reportExportPdf.addEventListener('click', exportReportPdf);
  els.modMediaInput.addEventListener('change', onModMediaSelected);
  els.modMediaPreview.addEventListener('click', onModMediaPreviewClick);
  els.modSubmitBtn.addEventListener('click', submitModReport);
  els.modResetBtn.addEventListener('click', resetModForm);
  els.modFilterChips.addEventListener('click', onModFilterClick);
  els.modSearch.addEventListener('input', onModSearchInput);
  els.modReportsList.addEventListener('click', onModReportListClick);
  els.modDetailPanel.addEventListener('click', onModDetailClick);
  els.dashGoTasks.addEventListener('click', () => { clearTaskContext(); showPage('tasks'); });
  els.dashGoHistory.addEventListener('click', () => showPage('history'));
  els.dashGoReports.addEventListener('click', () => showPage('reports'));
}


async function initializeBackend() {
  if (hasFirebaseConfig()) {
    try {
      initializeFirebaseServices();
      updateFirebaseStatus('กำลังเชื่อมต่อ Firebase…', 'info');
      try {
        await state.firebaseAuth.setPersistence(firebase.auth.Auth.Persistence.LOCAL);
      } catch (error) {
        console.warn('Auth persistence setup failed', error);
      }
      attachFirebaseAuthObserver();
      return;
    } catch (error) {
      console.error('Firebase initialization failed', error);
      updateFirebaseStatus('เชื่อม Firebase ไม่สำเร็จ กำลังใช้โหมด Local Demo', 'warn');
    }
  }

  state.backendMode = 'local';
  initializeUsers();
  initializeTasks();
  initializeModReports();
  restoreSession();
  updateFirebaseStatus('Local Demo / ยังไม่ได้ใส่ Firebase Config', 'warn');
}

function hasFirebaseConfig() {
  const cfg = window.LSH_FIREBASE_CONFIG || {};
  const enabled = (window.LSH_FIREBASE_OPTIONS?.useFirebase ?? true) !== false;
  return enabled && typeof window.firebase !== 'undefined' && cfg.apiKey && !String(cfg.apiKey).includes('PASTE_') && cfg.projectId;
}

function initializeFirebaseServices() {
  const cfg = window.LSH_FIREBASE_CONFIG;
  if (!firebase.apps.length) {
    firebase.initializeApp(cfg);
  }
  state.firebaseApp = firebase.app();
  state.firebaseAuth = firebase.auth();
  state.firebaseDb = firebase.firestore();
  state.firebaseStorage = firebase.storage();
  state.backendMode = 'firebase';
}

function attachFirebaseAuthObserver() {
  if (state.authObserverReady || !state.firebaseAuth) return;
  state.authObserverReady = true;
  state.firebaseAuth.onAuthStateChanged(async (authUser) => {
    await handleFirebaseAuthState(authUser);
  });
}

async function handleFirebaseAuthState(authUser) {
  teardownFirebaseSubscriptions();

  if (!authUser) {
    state.currentUser = null;
    localStorage.removeItem(STORAGE_KEYS.currentUser);
    screens.app.classList.remove('screen--active');
    screens.login.classList.add('screen--active');
    updateFirebaseStatus(hasFirebaseConfig() ? 'Firebase พร้อมใช้งาน / กรุณาเข้าสู่ระบบ' : 'Local Demo', hasFirebaseConfig() ? 'info' : 'warn');
    return;
  }

  updateFirebaseStatus('กำลังโหลดข้อมูลจาก Firebase…', 'info');
  const employeeId = emailToEmployeeId(authUser.email || '');
  let profileSnap = await state.firebaseDb.collection(FIREBASE_COLLECTIONS.users).doc(employeeId).get();

  if (!profileSnap.exists) {
    const pendingProfile = state.pendingRegistrationProfile;
    if (pendingProfile && pendingProfile.employeeId === employeeId) {
      await state.firebaseDb.collection(FIREBASE_COLLECTIONS.users).doc(employeeId).set(stripUserForFirestore(pendingProfile), { merge: true });
      profileSnap = await state.firebaseDb.collection(FIREBASE_COLLECTIONS.users).doc(employeeId).get();
    }
  }

  if (!profileSnap.exists) {
    const fallbackUser = defaultUsers.find((user) => user.employeeId === employeeId);
    if (fallbackUser) {
      await state.firebaseDb.collection(FIREBASE_COLLECTIONS.users).doc(employeeId).set(stripUserForFirestore(fallbackUser));
      profileSnap = await state.firebaseDb.collection(FIREBASE_COLLECTIONS.users).doc(employeeId).get();
    }
  }

  if (!profileSnap.exists) {
    els.loginError.textContent = 'ไม่พบข้อมูลผู้ใช้ใน Firestore (users collection)';
    els.loginError.classList.remove('hidden');
    updateFirebaseStatus('ไม่พบ users/{employeeId} ใน Firestore', 'error');
    await state.firebaseAuth.signOut();
    return;
  }

  state.currentUser = normalizeFirebaseUser(profileSnap.data());
  state.pendingRegistrationProfile = null;
  localStorage.setItem(STORAGE_KEYS.currentUser, JSON.stringify(state.currentUser));
  await bootstrapFirebaseDemoDataIfNeeded();
  await loadFirebaseCollections();
  subscribeFirebaseCollections();
  els.loginError.classList.add('hidden');
  updateFirebaseStatus(`เชื่อม Firebase แล้ว / ${state.firebaseApp.options.projectId}`, 'success');
  showApp();
}

async function bootstrapFirebaseDemoDataIfNeeded() {
  if (!state.firebaseDb) return;
  const shouldSeed = window.LSH_FIREBASE_OPTIONS?.seedDemoDataIfEmpty !== false;
  if (!shouldSeed) return;

  const [usersSnap, tasksSnap, modSnap] = await Promise.all([
    state.firebaseDb.collection(FIREBASE_COLLECTIONS.users).limit(1).get(),
    state.firebaseDb.collection(FIREBASE_COLLECTIONS.tasks).limit(1).get(),
    state.firebaseDb.collection(FIREBASE_COLLECTIONS.modReports).limit(1).get()
  ]);

  if (usersSnap.empty) {
    await syncCollectionDiff(FIREBASE_COLLECTIONS.users, [], defaultUsers, stripUserForFirestore, getUserDocId);
  }
  if (tasksSnap.empty) {
    await syncCollectionDiff(FIREBASE_COLLECTIONS.tasks, [], seedTasks, stripTaskForFirestore, getTaskDocId);
  }
  if (modSnap.empty) {
    await syncCollectionDiff(FIREBASE_COLLECTIONS.modReports, [], [...seedModReports, ...extraSeedModReports], stripModReportForFirestore, getModReportDocId);
  }
}

async function loadFirebaseCollections() {
  const [usersSnap, tasksSnap, modSnap] = await Promise.all([
    state.firebaseDb.collection(FIREBASE_COLLECTIONS.users).get(),
    state.firebaseDb.collection(FIREBASE_COLLECTIONS.tasks).get(),
    state.firebaseDb.collection(FIREBASE_COLLECTIONS.modReports).get()
  ]);

  users = usersSnap.docs.map((doc) => normalizeFirebaseUser({ ...doc.data(), employeeId: doc.id }));
  state.usersCache = users.map((user) => ({ ...user }));
  state.tasksCache = tasksSnap.docs.map((doc) => normalizeTask({ ...doc.data(), id: doc.id, ticketNo: doc.data().ticketNo || doc.id }));
  state.modReportsCache = modSnap.docs.map((doc) => normalizeModReport({ ...doc.data(), id: doc.id }));
}

function subscribeFirebaseCollections() {
  if (!state.firebaseDb) return;
  teardownFirebaseSubscriptions();

  state.unsubscribeFns = [
    state.firebaseDb.collection(FIREBASE_COLLECTIONS.users).onSnapshot((snapshot) => {
      users = snapshot.docs.map((doc) => normalizeFirebaseUser({ ...doc.data(), employeeId: doc.id }));
      state.usersCache = users.map((user) => ({ ...user }));
      if (state.currentUser) {
        const refreshed = users.find((user) => user.employeeId === state.currentUser.employeeId);
        if (refreshed) state.currentUser = { ...refreshed };
      }
      safeRenderAfterSync();
    }),
    state.firebaseDb.collection(FIREBASE_COLLECTIONS.tasks).onSnapshot((snapshot) => {
      state.tasksCache = snapshot.docs.map((doc) => normalizeTask({ ...doc.data(), id: doc.id, ticketNo: doc.data().ticketNo || doc.id }));
      safeRenderAfterSync();
    }),
    state.firebaseDb.collection(FIREBASE_COLLECTIONS.modReports).onSnapshot((snapshot) => {
      state.modReportsCache = snapshot.docs.map((doc) => normalizeModReport({ ...doc.data(), id: doc.id }));
      safeRenderAfterSync();
    })
  ];
}

function teardownFirebaseSubscriptions() {
  (state.unsubscribeFns || []).forEach((unsub) => {
    try { if (typeof unsub === 'function') unsub(); } catch (error) { console.warn(error); }
  });
  state.unsubscribeFns = [];
}

function safeRenderAfterSync() {
  if (!state.currentUser || !screens.app.classList.contains('screen--active')) return;
  renderApp();
}

function updateFirebaseStatus(message, tone = 'info') {
  if (!els.firebaseStatus) return;
  els.firebaseStatus.textContent = message;
  els.firebaseStatus.dataset.tone = tone;
}

function employeeIdToEmail(employeeId) {
  return `${String(employeeId).trim()}@${FIREBASE_LOGIN_DOMAIN}`;
}

function emailToEmployeeId(email) {
  return String(email || '').split('@')[0] || '';
}

function normalizeFirebaseUser(user) {
  return {
    employeeId: String(user.employeeId || ''),
    name: user.name || '-',
    role: user.role || 'FO Staff',
    department: user.department || 'FO',
    active: user.active !== false,
    password: '',
    createdAt: user.createdAt || '',
    updatedAt: user.updatedAt || ''
  };
}

function getUserDocId(user) {
  return String(user.employeeId || '').trim();
}

function getTaskDocId(task) {
  return String(task.id || task.ticketNo || '').trim();
}

function getModReportDocId(report) {
  return String(report.id || report.reportNo || '').trim();
}

function stripUserForFirestore(user) {
  return {
    employeeId: String(user.employeeId || '').trim(),
    name: user.name || '-',
    role: user.role || 'FO Staff',
    department: user.department || 'FO',
    active: user.active !== false,
    updatedAt: new Date().toISOString(),
    createdAt: user.createdAt || new Date().toISOString()
  };
}

function stripTaskForFirestore(task) {
  const draft = JSON.parse(JSON.stringify(normalizeTask(task)));
  delete draft.password;
  return draft;
}

function stripModReportForFirestore(report) {
  return JSON.parse(JSON.stringify(normalizeModReport(report)));
}

async function syncCollectionDiff(collectionName, previousItems, nextItems, serializer, getId) {
  if (state.backendMode !== 'firebase' || !state.firebaseDb) return;
  const colRef = state.firebaseDb.collection(collectionName);
  const batch = state.firebaseDb.batch();
  const previousIds = new Set((previousItems || []).map((item) => getId(item)).filter(Boolean));
  const nextIds = new Set();

  (nextItems || []).forEach((item) => {
    const docId = getId(item);
    if (!docId) return;
    nextIds.add(docId);
    batch.set(colRef.doc(docId), serializer(item), { merge: false });
  });

  previousIds.forEach((docId) => {
    if (!nextIds.has(docId)) batch.delete(colRef.doc(docId));
  });

  await batch.commit();
}

function handleBackgroundSyncError(error, context) {
  console.error(`Firebase sync failed: ${context}`, error);
  updateFirebaseStatus(`Firebase sync error: ${context}`, 'error');
}

async function prepareAttachmentsForSave(items, folderPrefix) {
  const list = Array.isArray(items) ? items : [];
  if (state.backendMode !== 'firebase' || !state.firebaseStorage) {
    return list.map((item) => ({
      id: item.id,
      name: item.name,
      type: item.type,
      size: item.size,
      kind: item.kind,
      dataUrl: item.dataUrl || '',
      inlineSaved: !!item.dataUrl
    }));
  }

  const uploaded = [];
  for (const item of list) {
    if (!item.file) {
      uploaded.push({
        id: item.id,
        name: item.name,
        type: item.type,
        size: item.size,
        kind: item.kind,
        url: item.url || item.dataUrl || '',
        storagePath: item.storagePath || '',
        inlineSaved: false
      });
      continue;
    }
    const safeName = sanitizeFileName(item.name || `${item.kind || 'file'}-${Date.now()}`);
    const filePath = `${folderPrefix}/${Date.now()}_${Math.random().toString(36).slice(2, 8)}_${safeName}`;
    const ref = state.firebaseStorage.ref(filePath);
    const snapshot = await ref.put(item.file);
    const url = await snapshot.ref.getDownloadURL();
    uploaded.push({
      id: item.id,
      name: item.name,
      type: item.type,
      size: item.size,
      kind: item.kind,
      url,
      storagePath: filePath,
      inlineSaved: false
    });
  }
  return uploaded;
}

function sanitizeFileName(name) {
  return String(name || 'file').replace(/[^a-zA-Z0-9._-]+/g, '_');
}

function getMediaPreviewSrc(item) {
  return item.url || item.downloadURL || item.dataUrl || '';
}

function getFriendlyAuthError(error) {
  const code = error?.code || '';
  const map = {
    'auth/invalid-credential': 'รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง',
    'auth/invalid-email': 'รูปแบบ Employee ID สำหรับ Firebase ไม่ถูกต้อง',
    'auth/user-not-found': 'ไม่พบบัญชีผู้ใช้นี้ใน Firebase Authentication',
    'auth/wrong-password': 'รหัสผ่านไม่ถูกต้อง',
    'auth/email-already-in-use': 'รหัสพนักงานนี้ถูกสมัครไว้แล้ว',
    'auth/weak-password': 'รหัสผ่านควรมีอย่างน้อย 6 ตัวอักษร',
    'auth/too-many-requests': 'มีการลองเข้าสู่ระบบหลายครั้งเกินไป กรุณาลองใหม่ภายหลัง',
    'auth/network-request-failed': 'เครือข่ายมีปัญหา กรุณาตรวจสอบอินเทอร์เน็ต'
  };
  return map[code] || error?.message || 'เข้าสู่ระบบไม่สำเร็จ';
}

function getStaffRoleForDepartment(department) {
  return {
    FO: 'FO Staff',
    HK: 'HK Staff',
    Engineering: 'ENG Staff',
    FB: 'FB Staff'
  }[department] || 'FO Staff';
}

function canSelfRegisterRole(role) {
  return ['FO Staff', 'HK Staff', 'ENG Staff', 'FB Staff'].includes(role);
}

function resetRegisterForm() {
  if (els.registerName) els.registerName.value = '';
  if (els.registerEmployeeId) els.registerEmployeeId.value = '';
  if (els.registerDepartment) els.registerDepartment.value = 'FO';
  if (els.registerPassword) els.registerPassword.value = '';
  if (els.registerPasswordConfirm) els.registerPasswordConfirm.value = '';
  if (els.registerError) {
    els.registerError.textContent = 'ไม่สามารถสมัครสมาชิกได้';
    els.registerError.classList.add('hidden');
  }
  if (els.registerSuccess) {
    els.registerSuccess.textContent = 'สมัครสมาชิกสำเร็จ กรุณาเข้าสู่ระบบ';
    els.registerSuccess.classList.add('hidden');
  }
}

function toggleRegisterPanel(forceOpen) {
  if (!els.registerPanel) return;
  const shouldOpen = typeof forceOpen === 'boolean' ? forceOpen : els.registerPanel.classList.contains('hidden');
  els.registerPanel.classList.toggle('hidden', !shouldOpen);
  if (!shouldOpen) {
    resetRegisterForm();
    return;
  }
  if (els.registerName) els.registerName.focus();
}

function validateRegisterInput(name, employeeId, department, password, confirmPassword) {
  if (!name || !employeeId || !department || !password || !confirmPassword) return 'กรุณากรอกข้อมูลสมัครสมาชิกให้ครบ';
  if (password.length < 6) return 'รหัสผ่านควรมีอย่างน้อย 6 ตัวอักษร';
  if (password !== confirmPassword) return 'ยืนยันรหัสผ่านไม่ตรงกัน';
  if (!/^\d{4,}$/.test(employeeId)) return 'รหัสพนักงานควรเป็นตัวเลขอย่างน้อย 4 หลัก';
  return '';
}

async function onRegister() {
  const name = (els.registerName?.value || '').trim();
  const employeeId = (els.registerEmployeeId?.value || '').trim();
  const department = els.registerDepartment?.value || 'FO';
  const password = (els.registerPassword?.value || '').trim();
  const confirmPassword = (els.registerPasswordConfirm?.value || '').trim();

  if (els.registerError) els.registerError.classList.add('hidden');
  if (els.registerSuccess) els.registerSuccess.classList.add('hidden');

  const validationError = validateRegisterInput(name, employeeId, department, password, confirmPassword);
  if (validationError) {
    if (els.registerError) {
      els.registerError.textContent = validationError;
      els.registerError.classList.remove('hidden');
    }
    return;
  }

  if (state.backendMode === 'firebase' && state.firebaseAuth && state.firebaseDb) {
    try {
      if (els.registerSubmitBtn) els.registerSubmitBtn.disabled = true;
      const email = employeeIdToEmail(employeeId);
      state.pendingRegistrationProfile = {
        employeeId,
        name,
        role: getStaffRoleForDepartment(department),
        department,
        active: true,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };
      const credential = await state.firebaseAuth.createUserWithEmailAndPassword(email, password);
      const docRef = state.firebaseDb.collection(FIREBASE_COLLECTIONS.users).doc(employeeId);
      const existingDoc = await docRef.get();
      const existing = existingDoc.exists ? normalizeFirebaseUser({ ...existingDoc.data(), employeeId }) : null;
      if (existing && existing.role && !canSelfRegisterRole(existing.role)) {
        state.pendingRegistrationProfile = null;
        await state.firebaseAuth.signOut();
        if (els.registerError) {
          els.registerError.textContent = 'บัญชีระดับหัวหน้างานหรือผู้จัดการ กรุณาให้แอดมินสร้างในระบบ';
          els.registerError.classList.remove('hidden');
        }
        return;
      }
      const profile = {
        employeeId,
        name: existing?.name || name,
        role: existing?.role || getStaffRoleForDepartment(existing?.department || department),
        department: existing?.department || department,
        active: true,
        createdAt: existing?.createdAt || new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };
      await docRef.set(stripUserForFirestore(profile), { merge: true });
      if (els.registerSuccess) {
        els.registerSuccess.textContent = 'สมัครสมาชิกสำเร็จ ระบบกำลังเข้าสู่ระบบให้';
        els.registerSuccess.classList.remove('hidden');
      }
      if (els.loginEmployeeId) els.loginEmployeeId.value = employeeId;
      if (els.loginPassword) els.loginPassword.value = password;
      return;
    } catch (error) {
      console.error(error);
      state.pendingRegistrationProfile = null;
      if (els.registerError) {
        els.registerError.textContent = getFriendlyAuthError(error) || 'สมัครสมาชิกไม่สำเร็จ';
        els.registerError.classList.remove('hidden');
      }
    } finally {
      if (els.registerSubmitBtn) els.registerSubmitBtn.disabled = false;
    }
    return;
  }

  const existingUser = users.find((user) => user.employeeId === employeeId);
  if (existingUser) {
    if (els.registerError) {
      els.registerError.textContent = 'รหัสพนักงานนี้มีอยู่แล้ว';
      els.registerError.classList.remove('hidden');
    }
    return;
  }

  users.push({
    employeeId,
    password,
    name,
    role: getStaffRoleForDepartment(department),
    department,
    active: true
  });
  saveUsers(users);
  if (els.registerSuccess) {
    els.registerSuccess.textContent = 'สมัครสมาชิกสำเร็จ กรุณาเข้าสู่ระบบ';
    els.registerSuccess.classList.remove('hidden');
  }
  if (els.loginEmployeeId) els.loginEmployeeId.value = employeeId;
  if (els.loginPassword) els.loginPassword.value = password;
  resetRegisterForm();
  if (els.registerSuccess) {
    els.registerSuccess.textContent = 'สมัครสมาชิกสำเร็จ กรุณาเข้าสู่ระบบ';
    els.registerSuccess.classList.remove('hidden');
  }
}

function initializeTasks() {
  const existing = localStorage.getItem(STORAGE_KEYS.tasks);
  if (!existing) {
    localStorage.setItem(STORAGE_KEYS.tasks, JSON.stringify(seedTasks));
  }
}

function initializeModReports() {
  const existing = localStorage.getItem(STORAGE_KEYS.modReports);
  if (!existing) {
    localStorage.setItem(STORAGE_KEYS.modReports, JSON.stringify([...seedModReports, ...extraSeedModReports]));
  }
}

function initializeUsers() {
  const existing = localStorage.getItem(STORAGE_KEYS.users);
  if (!existing) {
    localStorage.setItem(STORAGE_KEYS.users, JSON.stringify(defaultUsers));
  }
  users = getUsers();
}

function getUsers() {
  if (state.backendMode === 'firebase') return state.usersCache.map((user) => ({ ...user }));
  return JSON.parse(localStorage.getItem(STORAGE_KEYS.users) || '[]');
}

function saveUsers(nextUsers) {
  const normalized = nextUsers.map((user) => ({ ...user }));
  const previous = users.map((user) => ({ ...user }));
  users = normalized.map((user) => ({ ...user }));

  if (state.backendMode === 'firebase') {
    state.usersCache = users.map((user) => ({ ...user }));
    syncCollectionDiff(FIREBASE_COLLECTIONS.users, previous, state.usersCache, stripUserForFirestore, getUserDocId)
      .catch((error) => handleBackgroundSyncError(error, 'users'));
  } else {
    localStorage.setItem(STORAGE_KEYS.users, JSON.stringify(users));
  }

  if (state.currentUser) {
    const refreshed = users.find((user) => user.employeeId === state.currentUser.employeeId);
    if (refreshed) {
      state.currentUser = refreshed;
      localStorage.setItem(STORAGE_KEYS.currentUser, JSON.stringify(refreshed));
    }
  }
}

function restoreSession() {
  if (state.backendMode === 'firebase') return;
  const savedUser = localStorage.getItem(STORAGE_KEYS.currentUser);
  if (!savedUser) return;
  const parsed = JSON.parse(savedUser);
  const refreshed = users.find((user) => user.employeeId === parsed.employeeId) || parsed;
  state.currentUser = refreshed;
  localStorage.setItem(STORAGE_KEYS.currentUser, JSON.stringify(refreshed));
  showApp();
}

async function onLogin() {
  const employeeId = els.loginEmployeeId.value.trim();
  const password = els.loginPassword.value.trim();

  if (!employeeId || !password) {
    els.loginError.textContent = 'กรุณากรอกรหัสพนักงานและรหัสผ่าน';
    els.loginError.classList.remove('hidden');
    return;
  }

  if (state.backendMode === 'firebase' && state.firebaseAuth) {
    try {
      els.loginBtn.disabled = true;
      els.loginError.classList.add('hidden');
      await state.firebaseAuth.signInWithEmailAndPassword(employeeIdToEmail(employeeId), password);
    } catch (error) {
      console.error(error);
      els.loginError.textContent = getFriendlyAuthError(error);
      els.loginError.classList.remove('hidden');
    } finally {
      els.loginBtn.disabled = false;
    }
    return;
  }

  const matchedUser = users.find((u) => u.employeeId === employeeId && u.password === password);
  if (!matchedUser) {
    els.loginError.textContent = 'รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง';
    els.loginError.classList.remove('hidden');
    return;
  }

  els.loginError.classList.add('hidden');
  state.currentUser = matchedUser;
  localStorage.setItem(STORAGE_KEYS.currentUser, JSON.stringify(matchedUser));
  showApp();
}

async function logout() {
  state.currentUser = null;
  state.pendingRegistrationProfile = null;
  localStorage.removeItem(STORAGE_KEYS.currentUser);
  teardownFirebaseSubscriptions();
  if (state.backendMode === 'firebase' && state.firebaseAuth) {
    try {
      await state.firebaseAuth.signOut();
    } catch (error) {
      console.error('Firebase sign-out failed', error);
    }
  }
  screens.app.classList.remove('screen--active');
  screens.login.classList.add('screen--active');
  els.loginPassword.value = '';
  toggleRegisterPanel(false);
}

function showApp() {
  screens.login.classList.remove('screen--active');
  screens.app.classList.add('screen--active');
  toggleRegisterPanel(false);
  configureNavigation();
  showPage(getDefaultLandingPage());
  renderApp();
}

function configureNavigation() {
  const modPageAccess = canAccessMod();
  const reportsAccess = canViewReports();
  const tasksAccess = canViewTaskList();
  const createAccess = canCreateTasks();
  const historyAccess = canViewHistory();
  els.navHomeLabel.textContent = isMOD() ? 'MOD' : isCGM() ? 'Dash' : 'Home';
  els.navHistoryLabel.textContent = 'History';
  els.navTasks.classList.toggle('hidden', !tasksAccess);
  els.navCreate.classList.toggle('hidden', !createAccess);
  els.navHistory.classList.toggle('hidden', !historyAccess);
  els.navMod.classList.toggle('hidden', !modPageAccess || isMOD());
  els.navReports.classList.toggle('hidden', !reportsAccess);
  els.bottomNav.classList.remove('bottom-nav--count-1', 'bottom-nav--count-2', 'bottom-nav--count-3', 'bottom-nav--count-4', 'bottom-nav--count-5', 'bottom-nav--count-6');
  const visibleCount = [els.navHome, els.navTasks, els.navCreate, els.navHistory, els.navMod, els.navReports].filter((btn) => !btn.classList.contains('hidden')).length;
  els.bottomNav.classList.add(`bottom-nav--count-${Math.max(1, visibleCount)}`);
}

function showPage(pageName) {
  if (state.currentView !== pageName) {
    state.previousView = state.currentView;
  }
  state.currentView = pageName;

  Object.entries(pages).forEach(([name, element]) => {
    element.classList.toggle('page-view--active', name === pageName);
  });

  setActiveNav(pageName);
  updateTopbar(pageName);
  els.backBtn.classList.toggle('hidden', pageName === 'home' || pageName === 'dashboard' || pageName === 'mod');

  if (pageName === 'create') {
    els.openedByText.textContent = `Opened by: ${state.currentUser.name} / ${state.currentUser.department} / ${formatDateTime(new Date().toISOString())}`;
    renderCreateAssignmentState();
  }

  if (pageName === 'tasks') renderTaskList();
  if (pageName === 'detail') renderTaskDetail();
  if (pageName === 'history') renderHistoryPage();
  if (pageName === 'dashboard') renderDashboardPage();
  if (pageName === 'mod') renderModPage();
  if (pageName === 'reports') renderReportsPage();
}

function onBack() {
  if (state.currentView === 'detail') {
    showPage(state.previousView === 'history' ? 'history' : state.previousView === 'reports' ? 'reports' : state.previousView === 'mod' ? 'mod' : 'tasks');
    return;
  }
  showPage(getDefaultLandingPage());
}

function setActiveNav(pageName) {
  const activeMap = {
    home: els.navHome,
    dashboard: els.navHome,
    tasks: els.navTasks,
    detail: els.navTasks,
    create: els.navCreate,
    history: els.navHistory,
    mod: isMOD() ? els.navHome : els.navMod,
    reports: els.navReports
  };
  [els.navHome, els.navTasks, els.navCreate, els.navHistory, els.navMod, els.navReports].forEach((btn) => btn.classList.remove('is-active'));
  activeMap[pageName]?.classList.add('is-active');
}

function updateTopbar(pageName) {
  const titleMap = {
    home: [`Good Morning, ${state.currentUser.name}`, `${state.currentUser.department} / ${state.currentUser.role}`],
    dashboard: ['CGM Dashboard', 'Hotel operations overview'],
    create: ['Create Task', 'Open a new request'],
    tasks: ['Tasks', getTaskPageSubtitle()],
    detail: ['Task Detail', 'View status, notes, and action'],
    history: ['History', 'Search previous tasks'],
    mod: ['MOD Checklist Report', 'Open findings from daily inspection'],
    reports: ['Reports', 'Summary and export']
  };
  const [title, subtitle] = titleMap[pageName] || ['Laya Service Hub', ''];
  els.topbarTitle.textContent = title;
  els.topbarSubtitle.textContent = subtitle;
}

function renderApp() {
  const tasks = getTasks();
  renderSummary(tasks);
  renderHomeContent(tasks);
  renderCreateDepartmentChoices();
  renderCreateAssignmentState();
  renderTeamAdminSection();
  renderSupervisorTaskBoard();
  renderTaskList();
  renderHistoryPage();
  if (canAccessMod()) renderModPage();
  if (isCGM()) renderDashboardPage();
  if (canViewReports()) renderReportsPage();
  if (state.currentTaskId) renderTaskDetail();
}

function renderSummary(tasks) {
  const visibleTasks = getVisibleTasks(tasks);
  const todayKey = new Date().toDateString();
  els.statNew.textContent = visibleTasks.filter((t) => t.status === 'New').length;
  els.statProgress.textContent = visibleTasks.filter((t) => t.status === 'In Progress').length;
  els.statDone.textContent = visibleTasks.filter((t) => t.doneAt && new Date(t.doneAt).toDateString() === todayKey).length;
  els.statUrgent.textContent = visibleTasks.filter((t) => t.priority === 'Urgent' && t.status !== 'Closed').length;
}

function renderHomeContent(tasks) {
  if (isCGM()) return;

  const visibleTasks = getHomeScopedTasks(tasks);
  const config = getHomeConfig();
  const sectionBindings = [
    {
      title: els.homeSection1Title,
      hint: els.homeSection1Hint,
      button: els.homeSection1Btn,
      list: els.homeSection1List,
      config: config.sections[0]
    },
    {
      title: els.homeSection2Title,
      hint: els.homeSection2Hint,
      button: els.homeSection2Btn,
      list: els.homeSection2List,
      config: config.sections[1]
    },
    {
      title: els.homeSection3Title,
      hint: els.homeSection3Hint,
      button: els.homeSection3Btn,
      list: els.homeSection3List,
      config: config.sections[2]
    }
  ];

  sectionBindings.forEach((section) => {
    section.title.textContent = section.config.title;
    section.hint.textContent = section.config.hint;
    section.button.textContent = section.config.buttonLabel || 'Open';
    const sectionTasks = getHomeTasksByPreset(visibleTasks, section.config.preset)
      .filter((task) => task.status !== 'Closed')
      .sort(sortTasks)
      .slice(0, 4);

    section.list.innerHTML = sectionTasks.length
      ? sectionTasks.map((task) => taskCardHTML(task)).join('')
      : emptyStateHTML(section.config.emptyTitle || 'No tasks found', section.config.emptyDescription || 'Nothing to follow up right now.');
  });

  els.homeViewTasksBtn.textContent = config.tasksButtonLabel;
  els.recentActivityTitle.textContent = config.activityTitle;
  els.recentActivityHint.textContent = config.activityHint;
  renderRecentActivity(visibleTasks, config.activityPreset);
}

function getHomeScopedTasks(tasks) {
  if (isDepartmentStaff()) return tasks.filter((task) => task.department === state.currentUser.department);
  if (isDepartmentManager()) return tasks.filter((task) => task.department === state.currentUser.department);
  if (isFODispatcher() || isCGM()) return tasks;
  return getVisibleTasks(tasks);
}

function renderRecentActivity(tasks, preset = 'departmentRecent') {
  const allLogs = getHomeActivityEntries(tasks, preset)
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))
    .slice(0, 6);

  if (!allLogs.length) {
    els.recentActivity.innerHTML = '<div class="activity-row"><div class="activity-row__text">No recent activity</div></div>';
    return;
  }

  els.recentActivity.innerHTML = allLogs.map((entry) => `
    <div class="activity-row">
      <div class="activity-row__text">
        <strong>${escapeHtml(entry.task.ticketNo)}</strong> / ${escapeHtml(entry.action)} / ${escapeHtml(entry.note || entry.task.subject)}
      </div>
      <div class="activity-row__time">${timeAgo(entry.createdAt)}</div>
    </div>
  `).join('');
}

function getHomeConfig() {
  const department = state.currentUser?.department;
  const base = {
    tasksButtonLabel: 'View Tasks',
    tasksButtonPreset: 'deptOnly',
    activityTitle: 'Recent Activity',
    activityHint: 'Latest updates',
    activityPreset: 'departmentRecent',
    sections: []
  };

  if (isFODispatcher()) {
    const dor = isDOR();
    return {
      ...base,
      tasksButtonLabel: dor ? 'Open Hotel Task Board' : 'Open Hotel Tasks',
      tasksButtonPreset: 'hotelAllOpen',
      activityTitle: dor ? 'FO / Room Division Follow-up' : 'My Dispatch Activity',
      activityHint: dor ? 'Monitor tasks opened by FO and room division' : 'Tasks opened and followed by this account',
      activityPreset: dor ? 'foDispatchRecent' : 'createdByMe',
      sections: [
        {
          title: dor ? 'FO Opened Tasks' : 'Tasks Opened by Me',
          hint: dor ? 'All requests created by FO / DOR' : 'Track every request you dispatched to other departments',
          buttonLabel: 'Open Board',
          preset: dor ? 'foOpenedActive' : 'openedByMeActive',
          emptyTitle: 'No open dispatch task',
          emptyDescription: 'There is no active task opened from FO right now.'
        },
        {
          title: 'Pending FO Close',
          hint: 'Department manager approved the work and sent it back to FO for closure',
          buttonLabel: 'Open Pending',
          preset: dor ? 'doneWaitingFO' : 'doneWaitingMeClose',
          emptyTitle: 'Nothing waiting for FO close',
          emptyDescription: 'No manager-approved task is waiting for FO confirmation.'
        },
        {
          title: 'Urgent from MOD',
          hint: 'Urgent MOD findings across the hotel',
          buttonLabel: 'Open MOD',
          preset: 'modUrgentHotel',
          emptyTitle: 'No urgent MOD item',
          emptyDescription: 'There is no urgent MOD-driven task right now.'
        }
      ]
    };
  }

  if (isDepartmentManager()) {
    return {
      ...base,
      tasksButtonLabel: `Open ${department} Board`,
      tasksButtonPreset: 'deptBoard',
      activityTitle: `${department} Team Updates`,
      activityHint: 'Live activity from your department team',
      activityPreset: 'departmentRecent',
      sections: [
        {
          title: `${department} Open Board`,
          hint: 'All open work orders for your department',
          buttonLabel: 'Open Board',
          preset: 'deptBoard',
          emptyTitle: `No open ${department} task`,
          emptyDescription: 'Your department queue is currently clear.'
        },
        {
          title: 'Assigned / In Progress',
          hint: 'Work already delegated to your team',
          buttonLabel: 'Open Team',
          preset: 'deptAssignedActive',
          emptyTitle: 'No active delegated work',
          emptyDescription: 'No assigned or in-progress task in your department now.'
        },
        {
          title: 'Urgent from MOD',
          hint: 'MOD findings assigned to your department',
          buttonLabel: 'Open MOD',
          preset: 'modUrgentDept',
          emptyTitle: 'No urgent MOD task',
          emptyDescription: 'No urgent MOD item in your department queue.'
        }
      ]
    };
  }

  if (isDepartmentStaff()) {
    return {
      ...base,
      tasksButtonLabel: 'Open My Assigned Tasks',
      tasksButtonPreset: 'mineActive',
      activityTitle: `${department} Work Updates`,
      activityHint: 'Only tasks assigned to your name',
      activityPreset: 'mineActive',
      sections: [
        {
          title: 'Assigned to Me',
          hint: 'Tasks with your name as owner',
          buttonLabel: 'Open Mine',
          preset: 'mineActive',
          emptyTitle: 'No task assigned to you',
          emptyDescription: 'Your queue is clear at the moment.'
        },
        {
          title: `${department} Board`,
          hint: 'Department-wide board for visibility',
          buttonLabel: 'Open Dept',
          preset: 'deptOpen',
          emptyTitle: `No active ${department} task`,
          emptyDescription: 'No active department task is visible now.'
        },
        {
          title: 'Urgent from MOD',
          hint: 'Important MOD findings for your department',
          buttonLabel: 'Open MOD',
          preset: 'modUrgentDept',
          emptyTitle: 'No urgent MOD item',
          emptyDescription: 'No urgent MOD task for your department now.'
        }
      ]
    };
  }

  return {
    ...base,
    sections: [
      { title: 'Need Attention', hint: 'Open tasks for this department', buttonLabel: 'Open Queue', preset: 'deptOpen' },
      { title: 'My In Progress', hint: 'Active tasks assigned to me', buttonLabel: 'Open Mine', preset: 'mineActive' },
      { title: 'Urgent / Overdue', hint: 'Items requiring attention first', buttonLabel: 'Open Urgent', preset: 'deptUrgentOverdue' }
    ]
  };
}

function getHomeTasksByPreset(tasks, preset) {
  const department = state.currentUser?.department;
  const mine = (task) => task.assignedToName === state.currentUser?.name;
  const openedByMe = (task) => task.openedByName === state.currentUser?.name;
  const deptTask = (task) => task.department === department;
  const activeTask = (task) => !['Pending FO Closure', 'Closed'].includes(task.status);
  const urgentTask = (task) => ['Urgent', 'High'].includes(task.priority) || isTaskOverdue(task);
  const modTask = (task) => task.sourceType === 'MOD';
  const openedByFO = (task) => task.openedByDepartment === 'FO';

  const presets = {
    all: tasks,
    deptOnly: tasks.filter((task) => deptTask(task)),
    deptOpen: tasks.filter((task) => deptTask(task) && activeTask(task)),
    deptBoard: tasks.filter((task) => deptTask(task)).sort(sortTasks),
    mineActive: tasks.filter((task) => mine(task) && activeTask(task)),
    createdByMe: tasks.filter((task) => openedByMe(task)),
    openedByMeActive: tasks.filter((task) => openedByMe(task) && activeTask(task)),
    hotelAllOpen: tasks.filter((task) => activeTask(task)),
    foOpenedActive: tasks.filter((task) => openedByFO(task) && activeTask(task)),
    doneWaitingFO: tasks.filter((task) => task.status === 'Pending FO Closure' && openedByFO(task)),
    doneWaitingMeClose: tasks.filter((task) => task.status === 'Pending FO Closure' && openedByMe(task)),
    deptAssignedActive: tasks.filter((task) => deptTask(task) && ['Accepted', 'In Progress', 'Waiting Supervisor Review', 'Returned for Rework', 'Pending FO Closure'].includes(task.status)),
    modUrgentDept: tasks.filter((task) => deptTask(task) && modTask(task) && activeTask(task) && urgentTask(task)),
    modUrgentHotel: tasks.filter((task) => modTask(task) && activeTask(task) && urgentTask(task)),
    deptUrgentOverdue: tasks.filter((task) => deptTask(task) && activeTask(task) && urgentTask(task)),
    departmentRecent: tasks.filter((task) => deptTask(task) || task.openedByDepartment === department),
    foDispatchRecent: tasks.filter((task) => openedByFO(task) || task.openedByName === state.currentUser?.name)
  };

  return presets[preset] || presets.all;
}

function getHomeActivityEntries(tasks, preset) {
  return getHomeTasksByPreset(tasks, preset)
    .flatMap((task) => (task.logs || []).map((log) => ({ ...log, task })));
}

function openDepartmentTasksFromHome() {
  const preset = getHomeConfig().tasksButtonPreset || 'deptOnly';
  if (isMOD()) {
    showPage('mod');
    return;
  }
  openTaskPreset(preset);
}

function openTaskPreset(preset) {
  resetTaskListControls();
  const safePreset = isDepartmentStaff() && preset !== 'mineActive' ? 'mineActive' : preset;
  setTaskContext(safePreset);
  showPage('tasks');
}

function setTaskContext(preset) {
  state.taskContext = preset || 'all';
}

function clearTaskContext() {
  state.taskContext = 'all';
}

function resetTaskListControls() {
  state.taskStatusFilter = 'All';
  state.highOnly = false;
  state.taskSearch = '';
  if (els.tasksSearch) els.tasksSearch.value = '';
  if (els.taskFilterHigh) els.taskFilterHigh.classList.remove('is-active');
  resetSupervisorTaskBoardFilters(true);
  Array.from(els.taskStatusTabs.querySelectorAll('.tab')).forEach((tab) => {
    tab.classList.toggle('is-active', tab.dataset.status === 'All');
  });
}

function getTaskContextFilteredTasks(tasks) {
  return getHomeTasksByPreset(tasks, state.taskContext || 'all');
}

function getTaskContextLabel() {
  const labels = {
    all: '',
    deptOnly: `${state.currentUser?.department || 'Department'} tasks`,
    deptOpen: `${state.currentUser?.department || 'Department'} open tasks`,
    deptBoard: `${state.currentUser?.department || 'Department'} board`,
    mineActive: 'Tasks assigned to me',
    createdByMe: 'Recently created by me',
    openedByMeActive: 'Tasks opened by me',
    hotelAllOpen: 'All hotel open tasks',
    foOpenedActive: 'FO opened tasks',
    doneWaitingFO: 'Pending FO closure',
    doneWaitingMeClose: 'Pending FO closure / opened by me',
    deptAssignedActive: `${state.currentUser?.department || 'Department'} assigned / active`,
    modUrgentDept: `${state.currentUser?.department || 'Department'} urgent from MOD`,
    modUrgentHotel: 'Urgent from MOD / hotel-wide',
    deptUrgentOverdue: `${state.currentUser?.department || 'Department'} urgent / overdue`
  };
  return labels[state.taskContext] || '';
}

function getTaskPageSubtitle() {
  const label = getTaskContextLabel();
  return label ? `Filter and track work orders / ${label}` : 'Filter and track work orders';
}

function renderDashboardPage() {
  if (!isCGM()) return;
  const tasks = getVisibleTasks(getTasks());
  const openTasks = tasks.filter((t) => t.status !== 'Closed');
  const todayKey = new Date().toDateString();

  els.dashOpen.textContent = openTasks.length;
  els.dashProgress.textContent = tasks.filter((t) => t.status === 'In Progress').length;
  els.dashOverdue.textContent = tasks.filter((t) => isTaskOverdue(t)).length;
  els.dashClosedToday.textContent = tasks.filter((t) => t.closedAt && new Date(t.closedAt).toDateString() === todayKey).length;

  els.dashboardDepartments.innerHTML = DEPARTMENTS.map((department) => {
    const deptTasks = tasks.filter((t) => t.department === department);
    const open = deptTasks.filter((t) => t.status !== 'Closed').length;
    const doneToday = deptTasks.filter((t) => t.doneAt && new Date(t.doneAt).toDateString() === todayKey).length;
    return `
      <div class="overview-row">
        <span class="overview-row__label">${escapeHtml(department)}</span>
        <span class="overview-row__value">Open ${open} / Done ${doneToday}</span>
      </div>
    `;
  }).join('');

  const overdueTasks = tasks.filter((t) => isTaskOverdue(t)).sort(sortTasks).slice(0, 4);
  els.dashboardOverdue.innerHTML = overdueTasks.length
    ? overdueTasks.map((task) => taskCardHTML(task, { singleAction: 'View Detail' })).join('')
    : emptyStateHTML('No overdue tasks', 'Everything is currently within expected time.');

  const latestLogs = tasks
    .flatMap((task) => (task.logs || []).map((log) => ({ ...log, task })))
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))
    .slice(0, 6);

  els.dashboardActivity.innerHTML = latestLogs.length
    ? latestLogs.map((entry) => `
        <div class="activity-row">
          <div class="activity-row__text"><strong>${escapeHtml(entry.task.ticketNo)}</strong> / ${escapeHtml(entry.action)} / ${escapeHtml(entry.note || '')}</div>
          <div class="activity-row__time">${timeAgo(entry.createdAt)}</div>
        </div>
      `).join('')
    : '<div class="activity-row"><div class="activity-row__text">No recent activity</div></div>';
}

function renderTaskList() {
  renderSupervisorTaskBoard();
  const tasks = getTaskListResults();
  const contextLabel = getTaskContextLabel();
  els.taskContextBar.classList.toggle('hidden', !contextLabel);
  els.taskContextLabel.textContent = contextLabel || 'Filtered view';
  els.tasksList.innerHTML = tasks.length
    ? tasks.map((task) => taskCardHTML(task)).join('')
    : emptyStateHTML('No matching tasks', 'Try another status, team filter, or search keyword.');
}


function renderSupervisorTaskBoard() {
  if (!els.supervisorBoard) return;
  if (!isDepartmentManager()) {
    els.supervisorBoard.classList.add('hidden');
    return;
  }

  els.supervisorBoard.classList.remove('hidden');
  const summaryTasks = getSupervisorBoardSummaryTasks();
  const boardTasks = getSupervisorBoardLaneTasks();
  const balancingTasks = getSupervisorBalancingTasks();
  const snapshot = buildTeamWorkloadSnapshot(balancingTasks, state.currentUser?.department);
  const department = state.currentUser?.department || 'Department';

  els.teamFilterHint.textContent = buildSupervisorFilterHint();
  renderSupervisorBoardSummary(summaryTasks, department);
  renderSupervisorAssigneeOptions(summaryTasks);
  renderSupervisorWorkloadCards(snapshot);
  renderSupervisorBalancePanel(snapshot);
  renderSupervisorBoardLanes(boardTasks);
}

function renderSupervisorBoardSummary(tasks, department) {
  const items = [
    { label: 'Open Team', value: tasks.filter((task) => !['Pending FO Closure', 'Closed'].includes(task.status)).length, tone: '' },
    { label: 'Unassigned', value: tasks.filter((task) => task.status === 'New' && !task.assignedToName).length, tone: '' },
    { label: 'In Progress', value: tasks.filter((task) => task.status === 'In Progress').length, tone: '' },
    { label: 'Overdue', value: tasks.filter((task) => isTaskOverdue(task)).length, tone: ' stat-card--danger' }
  ];
  els.supervisorBoardSummary.innerHTML = items.map((item) => `
    <article class="card stat-card${item.tone}">
      <span class="stat-card__label">${escapeHtml(item.label)}</span>
      <strong class="stat-card__value">${escapeHtml(String(item.value))}</strong>
      <span class="helper-text">${escapeHtml(department)} team scope</span>
    </article>
  `).join('');
}

function renderSupervisorAssigneeOptions(tasks) {
  const options = [
    { value: '__ALL__', label: 'All team members' },
    { value: '__UNASSIGNED__', label: 'Unassigned only' },
    { value: '__ME__', label: `Assigned to me (${state.currentUser.name})` }
  ];
  getTeamUsers(state.currentUser.department).forEach((user) => {
    options.push({ value: `name:${user.name}`, label: `${user.name} / ${user.role}` });
  });

  const currentValue = options.some((option) => option.value === state.teamAssigneeFilter)
    ? state.teamAssigneeFilter
    : '__ALL__';
  if (currentValue !== state.teamAssigneeFilter) state.teamAssigneeFilter = currentValue;

  els.teamAssigneeFilter.innerHTML = options.map((option) => `
    <option value="${escapeHtml(option.value)}" ${option.value === currentValue ? 'selected' : ''}>${escapeHtml(option.label)}</option>
  `).join('');
}

function renderSupervisorWorkloadCards(snapshot) {
  const cards = snapshot.metrics.map((item) => {
    const user = item.user;
    const fill = Math.max(8, Math.round((item.score / snapshot.maxScore) * 100));
    const activeClass = matchesSelectedAssigneeName(user.name) ? ' is-active' : '';
    return `
      <button class="workload-card workload-card--${escapeHtml(item.bandKey)}${activeClass}" type="button" data-workload-owner="name:${escapeHtml(user.name)}">
        <div class="workload-card__top">
          <div>
            <div class="workload-card__name">${escapeHtml(user.name)}</div>
            <div class="workload-card__role">${escapeHtml(user.role)} / ${escapeHtml(user.department)}</div>
          </div>
          <div class="workload-card__count">${escapeHtml(String(item.activeCount))}</div>
        </div>
        <div class="workload-card__status-row">
          <span class="balance-badge balance-badge--${escapeHtml(item.bandKey)}">${escapeHtml(item.bandLabel)}</span>
          <span class="workload-card__score">Load ${escapeHtml(formatScore(item.score))}</span>
        </div>
        <div class="workload-card__bar"><span class="workload-card__fill" style="--fill:${fill}%"></span></div>
        <div class="workload-card__meta">
          <span>Accepted ${item.accepted}</span>
          <span>In Progress ${item.inProgress}</span>
          <span>Overdue ${item.overdue}</span>
        </div>
      </button>
    `;
  });

  els.supervisorWorkloadCards.innerHTML = cards.length
    ? cards.join('')
    : emptyStateHTML('No team members', 'Add team users to see workload cards.');
}

function renderSupervisorBalancePanel(snapshot) {
  if (!els.supervisorBalancePanel) return;
  const statusItems = [
    { label: 'Overloaded', value: snapshot.overloaded.length, meta: 'Needs relief', band: 'overloaded' },
    { label: 'Light / Available', value: snapshot.light.length + snapshot.available.length, meta: 'Ready for more', band: 'light' },
    { label: 'Unassigned Queue', value: snapshot.unassignedTasks.length, meta: 'Waiting for owner', band: 'balanced' },
    { label: 'Team Avg Load', value: formatScore(snapshot.avgScore), meta: `${formatScore(snapshot.avgActiveCount)} active avg`, band: 'heavy' }
  ];

  els.supervisorBalanceStatus.innerHTML = statusItems.map((item) => `
    <article class="balance-status-card balance-status-card--${escapeHtml(item.band)}">
      <span class="balance-status-card__label">${escapeHtml(item.label)}</span>
      <strong class="balance-status-card__value">${escapeHtml(String(item.value))}</strong>
      <span class="balance-status-card__meta">${escapeHtml(item.meta)}</span>
    </article>
  `).join('');

  const recommendations = buildBalanceRecommendations(snapshot);
  els.supervisorBalanceRecommendations.innerHTML = recommendations.length
    ? recommendations.map((item) => `
      <article class="balance-reco-card">
        <div class="balance-reco-card__head">
          <div>
            <h4 class="balance-reco-card__title">${escapeHtml(item.title)}</h4>
            <p class="balance-reco-card__meta">${escapeHtml(item.meta)}</p>
          </div>
          <button class="btn btn-secondary" type="button" data-balance-action="${escapeHtml(item.action)}">${escapeHtml(item.buttonLabel || 'Run')}</button>
        </div>
      </article>
    `).join('')
    : emptyStateHTML('Team is balanced', 'No urgent rebalancing action is needed right now.');

  const suggestedTarget = snapshot.lightestMetric?.user?.name || 'best target';
  const heaviestOwner = snapshot.heaviestMetric?.user?.name || 'busy owner';
  const moveDisabled = getRebalanceMovePlan(snapshot) ? '' : 'disabled';
  const assignDisabled = snapshot.unassignedTasks.length ? '' : 'disabled';
  const autoDisabled = getAutoBalancePlan(snapshot, 2).steps.length ? '' : 'disabled';

  els.supervisorBalanceActions.innerHTML = `
    <button class="btn btn-primary" type="button" data-balance-action="assign-unassigned" ${assignDisabled}>Assign oldest unassigned → ${escapeHtml(suggestedTarget)}</button>
    <button class="btn btn-secondary" type="button" data-balance-action="rebalance-one" ${moveDisabled}>Move 1 task from ${escapeHtml(heaviestOwner)} → ${escapeHtml(suggestedTarget)}</button>
    <button class="btn btn-secondary" type="button" data-balance-action="auto-balance" ${autoDisabled}>Auto Balance 2 steps</button>
  `;
}

function renderSupervisorBoardLanes(tasks) {
  const lanes = [
    { title: 'New / Queue', status: 'New' },
    { title: 'Accepted', status: 'Accepted' },
    { title: 'In Progress', status: 'In Progress' },
    { title: 'Waiting Review', status: 'Waiting Supervisor Review' }
  ];

  els.supervisorBoardLanes.innerHTML = lanes.map((lane) => {
    const laneTasks = tasks.filter((task) => task.status === lane.status).slice(0, 6);
    return `
      <section class="board-lane">
        <div class="board-lane__head">
          <div class="board-lane__title">${escapeHtml(lane.title)}</div>
          <div class="board-lane__count">${escapeHtml(String(tasks.filter((task) => task.status === lane.status).length))}</div>
        </div>
        <div class="board-lane__body">
          ${laneTasks.length ? laneTasks.map((task) => boardMiniCardHTML(task)).join('') : '<div class="helper-text">No task in this lane.</div>'}
        </div>
      </section>
    `;
  }).join('');
}

function boardMiniCardHTML(task) {
  const overdueClass = isTaskOverdue(task) ? ' is-overdue' : '';
  return `
    <article class="board-mini-card${overdueClass}">
      <div class="board-mini-card__top">
        <span class="board-mini-card__ticket">${escapeHtml(task.ticketNo)}</span>
        <span class="board-mini-card__ticket">${timeAgo(task.updatedAt || task.createdAt)}</span>
      </div>
      <h4 class="board-mini-card__title">${escapeHtml(task.subject)}</h4>
      <div class="board-mini-card__meta">
        <span>${escapeHtml(task.location)}</span>
        <span>${escapeHtml(task.assignedToName || 'Unassigned')}</span>
      </div>
      <div class="board-mini-card__badges">
        <span class="badge ${priorityBadgeClass(task.priority)}">${escapeHtml(task.priority)}</span>
        <span class="badge ${statusBadgeClass(task.status)}">${escapeHtml(task.status)}</span>
      </div>
      <div class="board-mini-card__foot">
        <span>Opened by ${escapeHtml(task.openedByDepartment)}</span>
        <button class="btn btn-secondary board-mini-card__btn" type="button" data-task-view="${escapeHtml(task.id)}">Open</button>
      </div>
    </article>
  `;
}

function buildSupervisorFilterHint() {
  const quickMap = {
    all: 'all team tasks',
    active: 'active team tasks',
    unassigned: 'unassigned queue',
    assignedByMe: 'tasks assigned by me',
    overdue: 'overdue tasks'
  };
  return `${state.currentUser.department} supervisor board / ${quickMap[state.teamQuickFilter] || 'all team tasks'} / ${getTeamAssigneeLabel()}`;
}

function onTeamQuickFilterClick(event) {
  const chip = event.target.closest('[data-team-quick]');
  if (!chip) return;
  state.teamQuickFilter = chip.dataset.teamQuick;
  Array.from(els.teamQuickFilters.querySelectorAll('[data-team-quick]')).forEach((node) => node.classList.toggle('is-active', node === chip));
  renderTaskList();
}

function onTeamAssigneeFilterChange(event) {
  state.teamAssigneeFilter = event.target.value;
  renderTaskList();
}

function onWorkloadCardClick(event) {
  const card = event.target.closest('[data-workload-owner]');
  if (!card) return;
  state.teamAssigneeFilter = card.dataset.workloadOwner;
  els.teamAssigneeFilter.value = state.teamAssigneeFilter;
  renderTaskList();
}

function onSupervisorBalanceActionClick(event) {
  const button = event.target.closest('[data-balance-action]');
  if (!button || button.disabled) return;

  const snapshot = buildTeamWorkloadSnapshot(getSupervisorBalancingTasks(), state.currentUser?.department);
  let message = '';

  switch (button.dataset.balanceAction) {
    case 'assign-unassigned':
      message = runAssignOldestUnassigned(snapshot);
      break;
    case 'rebalance-one':
      message = runRebalanceOne(snapshot);
      break;
    case 'auto-balance':
      message = runAutoBalance(snapshot, 2);
      break;
    default:
      return;
  }

  if (message) alert(message);
  renderTaskList();
}

function onSupervisorBoardClick(event) {
  const viewButton = event.target.closest('[data-task-view]');
  if (!viewButton) return;
  openTaskDetail(viewButton.dataset.taskView);
}

function resetSupervisorTaskBoardFilters(silent = false) {
  if (!isDepartmentManager()) return;
  state.teamQuickFilter = 'all';
  state.teamAssigneeFilter = '__ALL__';
  if (els.teamQuickFilters) {
    Array.from(els.teamQuickFilters.querySelectorAll('[data-team-quick]')).forEach((chip) => chip.classList.toggle('is-active', chip.dataset.teamQuick === 'all'));
  }
  if (els.teamAssigneeFilter) {
    els.teamAssigneeFilter.value = '__ALL__';
  }
  if (!silent) renderTaskList();
}

function getSupervisorBoardSummaryTasks() {
  return getSupervisorTaskScope({ ignoreAssigneeFilter: true, ignoreStatusFilter: true, ignoreHighOnly: true, ignoreSearch: true }).filter((task) => task.department === state.currentUser.department);
}

function getSupervisorBalancingTasks() {
  return getSupervisorTaskScope({ ignoreAssigneeFilter: true, ignoreStatusFilter: true, ignoreHighOnly: true, ignoreSearch: true }).filter((task) => task.department === state.currentUser.department);
}

function getSupervisorBoardLaneTasks() {
  return getSupervisorTaskScope({ ignoreStatusFilter: true }).filter((task) => task.department === state.currentUser.department);
}

function getSupervisorTaskScope(options = {}) {
  const {
    ignoreAssigneeFilter = false,
    ignoreStatusFilter = false,
    ignoreHighOnly = false,
    ignoreSearch = false
  } = options;

  let tasks = getTaskContextFilteredTasks(getVisibleTasks(getTasks())).filter((task) => task.department === state.currentUser.department);
  tasks = tasks.filter((task) => matchesSupervisorQuickFilter(task));
  if (!ignoreAssigneeFilter) tasks = tasks.filter((task) => matchesSupervisorAssigneeFilter(task));
  if (!ignoreStatusFilter) tasks = tasks.filter((task) => state.taskStatusFilter === 'All' || task.status === state.taskStatusFilter);
  if (!ignoreHighOnly) tasks = tasks.filter((task) => !state.highOnly || ['High', 'Urgent'].includes(task.priority));
  if (!ignoreSearch) tasks = tasks.filter((task) => matchesSearch(task, state.taskSearch));
  return tasks.sort(sortTasks);
}

function buildTeamWorkloadSnapshot(tasks, department) {
  const teamUsers = getTeamUsers(department);
  const activeTasks = tasks.filter((task) => !['Pending FO Closure', 'Closed'].includes(task.status));
  const unassignedTasks = activeTasks
    .filter((task) => !task.assignedToName)
    .sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt) || priorityWeight(b.priority) - priorityWeight(a.priority));

  const rawMetrics = teamUsers.map((user) => {
    const assignedTasks = activeTasks.filter((task) => task.assignedToName === user.name);
    const score = assignedTasks.reduce((total, task) => total + taskLoadScore(task), 0);
    return {
      user,
      tasks: assignedTasks.sort(sortTasks),
      activeCount: assignedTasks.length,
      accepted: assignedTasks.filter((task) => task.status === 'Accepted').length,
      inProgress: assignedTasks.filter((task) => task.status === 'In Progress').length,
      overdue: assignedTasks.filter((task) => isTaskOverdue(task)).length,
      score
    };
  });

  const avgScore = rawMetrics.length ? rawMetrics.reduce((sum, item) => sum + item.score, 0) / rawMetrics.length : 0;
  const avgActiveCount = rawMetrics.length ? rawMetrics.reduce((sum, item) => sum + item.activeCount, 0) / rawMetrics.length : 0;
  const metrics = rawMetrics
    .map((item) => {
      const bandKey = classifyWorkloadBand(item, avgScore, avgActiveCount);
      return {
        ...item,
        bandKey,
        bandLabel: getWorkloadBandLabel(bandKey)
      };
    })
    .sort((a, b) => b.score - a.score || b.activeCount - a.activeCount || a.user.name.localeCompare(b.user.name));

  const available = metrics.filter((item) => item.bandKey === 'available');
  const light = metrics.filter((item) => item.bandKey === 'light');
  const balanced = metrics.filter((item) => item.bandKey === 'balanced');
  const heavy = metrics.filter((item) => item.bandKey === 'heavy');
  const overloaded = metrics.filter((item) => item.bandKey === 'overloaded');
  const metricsAsc = [...metrics].sort((a, b) => a.score - b.score || a.activeCount - b.activeCount || a.user.name.localeCompare(b.user.name));

  return {
    department,
    teamUsers,
    activeTasks,
    unassignedTasks,
    metrics,
    avgScore,
    avgActiveCount,
    maxScore: Math.max(1, ...metrics.map((item) => item.score)),
    available,
    light,
    balanced,
    heavy,
    overloaded,
    heaviestMetric: metrics[0] || null,
    lightestMetric: metricsAsc[0] || null,
    orderedTargets: metricsAsc,
    orderedSources: [...metrics]
  };
}

function taskLoadScore(task) {
  const statusScore = { New: 1, Accepted: 1.15, 'In Progress': 1.45 }[task.status] || 1;
  const priorityScore = { Low: 0.05, Medium: 0.25, High: 0.55, Urgent: 0.9 }[task.priority] || 0.2;
  const overdueScore = isTaskOverdue(task) ? 0.95 : 0;
  return Number((statusScore + priorityScore + overdueScore).toFixed(2));
}

function classifyWorkloadBand(item, avgScore, avgActiveCount) {
  if (item.activeCount === 0) return 'available';
  const overloadThreshold = Math.max(avgScore * 1.65, avgScore + 1.5, 3.2);
  const heavyThreshold = Math.max(avgScore * 1.2, avgScore + 0.6, 2.1);
  const lightThreshold = Math.max(0.85, avgScore * 0.65);

  if (item.score >= overloadThreshold || item.activeCount >= avgActiveCount + 2) return 'overloaded';
  if (item.score >= heavyThreshold || item.activeCount >= avgActiveCount + 1) return 'heavy';
  if (item.score <= lightThreshold || item.activeCount <= Math.max(0, avgActiveCount - 1.1)) return 'light';
  return 'balanced';
}

function getWorkloadBandLabel(bandKey) {
  return {
    overloaded: 'Overloaded',
    heavy: 'Heavy',
    balanced: 'Balanced',
    light: 'Light',
    available: 'Available'
  }[bandKey] || 'Balanced';
}

function buildBalanceRecommendations(snapshot) {
  const items = [];
  const assignTarget = pickBestTargetMetric(snapshot);
  if (snapshot.unassignedTasks.length && assignTarget) {
    const task = snapshot.unassignedTasks[0];
    items.push({
      action: 'assign-unassigned',
      title: `Assign oldest unassigned to ${assignTarget.user.name}`,
      meta: `${task.ticketNo} · ${task.subject}`,
      buttonLabel: 'Assign now'
    });
  }

  const movePlan = getRebalanceMovePlan(snapshot);
  if (movePlan) {
    items.push({
      action: 'rebalance-one',
      title: `Move 1 task from ${movePlan.from.user.name} to ${movePlan.to.user.name}`,
      meta: `${movePlan.task.ticketNo} · ${movePlan.task.subject}`,
      buttonLabel: 'Move now'
    });
  }

  if (getAutoBalancePlan(snapshot, 2).steps.length) {
    items.push({
      action: 'auto-balance',
      title: 'Auto balance recommended queue',
      meta: 'Runs up to 2 safe balancing steps and keeps a full timeline log.',
      buttonLabel: 'Run 2 steps'
    });
  }

  return items;
}

function pickBestTargetMetric(snapshot, { excludeName = '' } = {}) {
  return snapshot.orderedTargets.find((item) => item.user.name !== excludeName) || null;
}

function pickRebalanceCandidateTask(tasks) {
  const safeTasks = tasks.filter((task) => ['New', 'Accepted'].includes(task.status));
  const pool = safeTasks.length ? safeTasks : tasks.filter((task) => task.status !== 'In Progress');
  return (pool.length ? pool : [])
    .slice()
    .sort((a, b) => Number(isTaskOverdue(b)) - Number(isTaskOverdue(a)) || priorityWeight(b.priority) - priorityWeight(a.priority) || new Date(a.createdAt) - new Date(b.createdAt))[0] || null;
}

function getRebalanceMovePlan(snapshot) {
  const donor = snapshot.orderedSources.find((item) => ['overloaded', 'heavy'].includes(item.bandKey) && item.activeCount > 0);
  if (!donor) return null;
  const target = pickBestTargetMetric(snapshot, { excludeName: donor.user.name });
  if (!target) return null;
  if (donor.score - target.score < 0.9 && donor.activeCount - target.activeCount < 2) return null;
  const task = pickRebalanceCandidateTask(donor.tasks);
  if (!task) return null;
  return { from: donor, to: target, task };
}

function getAutoBalancePlan(snapshot, maxSteps = 2) {
  const steps = [];
  let draft = JSON.parse(JSON.stringify(snapshot));
  for (let index = 0; index < maxSteps; index += 1) {
    const assignTarget = pickBestTargetMetric(draft);
    if (draft.unassignedTasks.length && assignTarget) {
      const task = draft.unassignedTasks.shift();
      steps.push({ type: 'assign-unassigned', taskId: task.id, assigneeName: assignTarget.user.name, taskLabel: `${task.ticketNo}` });
      draft = buildTeamWorkloadSnapshot(simulateAssignment(draft.activeTasks, task.id, assignTarget.user.name), draft.department);
      continue;
    }
    const movePlan = getRebalanceMovePlan(draft);
    if (!movePlan) break;
    steps.push({ type: 'rebalance-one', taskId: movePlan.task.id, assigneeName: movePlan.to.user.name, fromName: movePlan.from.user.name, taskLabel: `${movePlan.task.ticketNo}` });
    draft = buildTeamWorkloadSnapshot(simulateAssignment(draft.activeTasks, movePlan.task.id, movePlan.to.user.name), draft.department);
  }
  return { steps };
}

function simulateAssignment(tasks, taskId, newOwnerName) {
  return tasks.map((task) => task.id === taskId || task.ticketNo === taskId ? { ...task, assignedToName: newOwnerName, status: task.status === 'New' ? 'Accepted' : task.status } : task);
}

function runAssignOldestUnassigned(snapshot) {
  const task = snapshot.unassignedTasks[0];
  const target = pickBestTargetMetric(snapshot);
  if (!task || !target) return 'No unassigned task can be balanced right now.';
  performAssignment(task.id, target.user, 'Supervisor workload balancing / oldest unassigned task auto-assigned.');
  return `${task.ticketNo} assigned to ${target.user.name} from the unassigned queue.`;
}

function runRebalanceOne(snapshot) {
  const plan = getRebalanceMovePlan(snapshot);
  if (!plan) return 'No safe rebalance move is available right now.';
  performAssignment(plan.task.id, plan.to.user, `Supervisor workload balancing / moved from ${plan.from.user.name} to ${plan.to.user.name}.`);
  return `${plan.task.ticketNo} moved from ${plan.from.user.name} to ${plan.to.user.name}.`;
}

function runAutoBalance(snapshot, maxSteps = 2) {
  const plan = getAutoBalancePlan(snapshot, maxSteps);
  if (!plan.steps.length) return 'Team queue is already balanced enough for now.';

  const messages = [];
  plan.steps.forEach((step, index) => {
    const assignee = users.find((user) => user.name === step.assigneeName && user.department === state.currentUser.department);
    if (!assignee) return;
    const note = step.type === 'assign-unassigned'
      ? `Supervisor workload balancing / step ${index + 1}: picked from unassigned queue.`
      : `Supervisor workload balancing / step ${index + 1}: auto-rebalanced from ${step.fromName} to ${step.assigneeName}.`;
    performAssignment(step.taskId, assignee, note);
    messages.push(`${index + 1}. ${step.taskLabel} → ${step.assigneeName}`);
  });

  return messages.length
    ? `Auto balance completed:
${messages.join('\n')}`
    : 'No balancing step was applied.';
}

function matchesSupervisorQuickFilter(task) {
  switch (state.teamQuickFilter) {
    case 'active':
      return ['Accepted', 'In Progress', 'Waiting Supervisor Review', 'Returned for Rework'].includes(task.status);
    case 'unassigned':
      return task.status === 'New' && !task.assignedToName;
    case 'assignedByMe':
      return !['Pending FO Closure', 'Closed'].includes(task.status) && task.assignedByName === state.currentUser.name;
    case 'overdue':
      return isTaskOverdue(task);
    default:
      return true;
  }
}

function matchesSupervisorAssigneeFilter(task) {
  if (state.teamAssigneeFilter === '__ALL__') return true;
  if (state.teamAssigneeFilter === '__UNASSIGNED__') return !task.assignedToName;
  if (state.teamAssigneeFilter === '__ME__') return task.assignedToName === state.currentUser.name;
  if (state.teamAssigneeFilter.startsWith('name:')) return task.assignedToName === state.teamAssigneeFilter.replace('name:', '');
  return true;
}

function getTeamUsers(department) {
  return users.filter((user) => user.department === department && isAssignableRole(user.role));
}

function getTeamAssigneeLabel() {
  if (state.teamAssigneeFilter === '__ALL__') return 'all owners';
  if (state.teamAssigneeFilter === '__UNASSIGNED__') return 'unassigned only';
  if (state.teamAssigneeFilter === '__ME__') return `assigned to me (${state.currentUser.name})`;
  if (state.teamAssigneeFilter.startsWith('name:')) return `owner ${state.teamAssigneeFilter.replace('name:', '')}`;
  return 'all owners';
}

function matchesSelectedAssigneeName(name) {
  return state.teamAssigneeFilter === `name:${name}` || (state.teamAssigneeFilter === '__ME__' && name === state.currentUser.name);
}

function formatScore(value) {
  return Number(value || 0).toFixed(1);
}

function renderTaskDetail() {
  const task = getTaskById(state.currentTaskId);
  if (!task) {
    els.detailLocation.textContent = '-';
    els.detailSubject.textContent = 'Task not found';
    els.detailTicket.textContent = '-';
    els.detailDescription.textContent = 'This task may have been removed.';
    els.detailDepartment.textContent = '-';
    els.detailCategory.textContent = '-';
    els.detailOpenedBy.textContent = '-';
    els.detailAssignedTo.textContent = '-';
    els.detailAssignedBy.textContent = '-';
    els.detailSource.textContent = '-';
    els.detailCreatedAt.textContent = '-';
    els.detailUpdatedAt.textContent = '-';
    els.detailPriorityBadge.className = 'badge badge-priority-low';
    els.detailPriorityBadge.textContent = '-';
    els.detailStatusBadge.className = 'badge badge-status-new';
    els.detailStatusBadge.textContent = '-';
    els.detailActions.innerHTML = '<div class="helper-text">No action available.</div>';
    els.detailAssignCard.classList.add('hidden');
    els.detailMediaCard.classList.add('hidden');
    els.detailMediaList.innerHTML = '';
    els.detailNoteInput.value = '';
    els.detailTimeline.innerHTML = emptyStateHTML('Task not found', 'This task may have been removed.');
    return;
  }

  els.detailLocation.textContent = `${task.location} / ${task.department}`;
  els.detailSubject.textContent = task.subject;
  els.detailTicket.textContent = task.ticketNo;
  els.detailDescription.textContent = task.detail || 'No detail provided.';
  els.detailDepartment.textContent = task.department;
  els.detailCategory.textContent = task.category;
  els.detailOpenedBy.textContent = `${task.openedByName} / ${task.openedByDepartment}`;
  els.detailAssignedTo.textContent = task.assignedToName || '-';
  els.detailAssignedBy.textContent = task.assignedByName ? `${task.assignedByName} / ${task.assignedByDepartment || '-'}` : '-';
  els.detailSource.textContent = task.sourceType ? `${task.sourceType}${task.sourceReference ? ` / ${task.sourceReference}` : ''}` : '-';
  els.detailCreatedAt.textContent = formatDateTime(task.createdAt);
  els.detailUpdatedAt.textContent = formatDateTime(task.updatedAt || task.createdAt);

  els.detailPriorityBadge.className = `badge ${priorityBadgeClass(task.priority)}`;
  els.detailPriorityBadge.textContent = task.priority;
  els.detailStatusBadge.className = `badge ${statusBadgeClass(task.status)}`;
  els.detailStatusBadge.textContent = task.status;

  const actions = getDetailActions(task);
  els.detailActions.innerHTML = actions.length
    ? actions.map((action) => `<button class="btn ${action.primary ? 'btn-primary' : 'btn-secondary'} ${action.full ? 'btn-full' : ''}" type="button" data-detail-action="${action.value}">${action.label}</button>`).join('')
    : '<div class="helper-text">No direct action available for this user in current status.</div>';

  renderDetailAssignmentState(task);
  renderDetailWorkUploadCard(task);
  renderTaskMedia(task);

  const logs = [...(task.logs || [])].sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  els.detailTimeline.innerHTML = logs.length
    ? logs.map((log) => timelineItemHTML(log)).join('')
    : emptyStateHTML('No timeline yet', 'Updates will appear here.');
}

function renderHistoryPage() {
  const tasks = getHistoryResults();
  els.historyCount.textContent = `${tasks.length} item${tasks.length === 1 ? '' : 's'}`;
  els.historyResults.innerHTML = tasks.length
    ? tasks.map((task) => taskCardHTML(task, { singleAction: 'View Detail' })).join('')
    : emptyStateHTML('No history found', 'Try another date range or filter.');
}

function renderReportsPage() {
  if (!canViewReports()) {
    els.reportResults.innerHTML = emptyStateHTML('CGM only', 'This page is available for CGM role only.');
    if (els.reportDepartmentChart) els.reportDepartmentChart.innerHTML = chartEmptyStateHTML('CGM only');
    if (els.reportTrendChart) els.reportTrendChart.innerHTML = chartEmptyStateHTML('CGM only');
    if (els.reportPriorityChart) els.reportPriorityChart.innerHTML = chartEmptyStateHTML('CGM only');
    if (els.reportPrioritySummary) els.reportPrioritySummary.innerHTML = '';
    if (els.reportStatusDonut) els.reportStatusDonut.innerHTML = chartEmptyStateHTML('CGM only');
    if (els.reportStatusLegend) els.reportStatusLegend.innerHTML = '';
    return;
  }

  const tasks = getReportResults();
  const closed = tasks.filter((t) => t.status === 'Closed').length;
  const open = tasks.filter((t) => t.status !== 'Closed').length;
  const overdue = tasks.filter((t) => isTaskOverdue(t)).length;
  const closureRate = tasks.length ? Math.round((closed / tasks.length) * 100) : 0;

  els.reportTotal.textContent = tasks.length;
  els.reportClosed.textContent = closed;
  els.reportOpen.textContent = open;
  els.reportOverdue.textContent = overdue;
  els.reportRangeLabel.textContent = buildReportRangeLabel();
  if (els.reportViewModeLabel) {
    els.reportViewModeLabel.textContent = `${getReportViewModeLabel()} view - open vs closed`;
  }

  const deptStats = DEPARTMENTS.map((department) => {
    const deptTasks = tasks.filter((t) => t.department === department);
    const deptClosed = deptTasks.filter((t) => t.status === 'Closed').length;
    const deptOverdue = deptTasks.filter((t) => isTaskOverdue(t)).length;
    return {
      department,
      count: deptTasks.length,
      closed: deptClosed,
      overdue: deptOverdue,
      share: tasks.length ? Math.round((deptTasks.length / tasks.length) * 100) : 0
    };
  });

  const statusStats = STATUSES.map((status) => {
    const count = tasks.filter((t) => t.status === status).length;
    return {
      status,
      count,
      share: tasks.length ? Math.round((count / tasks.length) * 100) : 0
    };
  });

  const priorityStats = PRIORITIES.map((priority) => {
    const priorityTasks = tasks.filter((t) => t.priority === priority);
    const count = priorityTasks.length;
    return {
      priority,
      count,
      active: priorityTasks.filter((t) => !['Closed', 'Pending FO Closure'].includes(t.status)).length,
      share: tasks.length ? Math.round((count / tasks.length) * 100) : 0
    };
  });

  const trendStats = getTrendStats(tasks, state.reportViewMode);

  els.reportDepartmentSummary.innerHTML = deptStats.map((item) => metricRowHTML({
    label: item.department,
    value: `${item.count} task${item.count === 1 ? '' : 's'}`,
    meta: `${item.closed} closed / ${item.overdue} overdue`,
    share: item.share,
    tone: 'department'
  })).join('');

  els.reportStatusSummary.innerHTML = statusStats.map((item) => metricRowHTML({
    label: item.status,
    value: `${item.count}`,
    meta: `${item.share}% of selected workload`,
    share: item.share,
    tone: 'status'
  })).join('');

  if (els.reportPrioritySummary) {
    els.reportPrioritySummary.innerHTML = priorityStats.map((item) => metricRowHTML({
      label: item.priority,
      value: `${item.count} task${item.count === 1 ? '' : 's'}`,
      meta: `${item.active} active / ${item.share}% of selected workload`,
      share: item.share,
      tone: 'priority'
    })).join('');
  }

  els.reportExecSummary.innerHTML = buildExecutiveSummary(tasks, deptStats, { closed, open, overdue, closureRate })
    .map((line) => `<li>${escapeHtml(line)}</li>`).join('');

  if (els.reportDepartmentChart) {
    els.reportDepartmentChart.innerHTML = tasks.length
      ? renderDepartmentChart(deptStats)
      : chartEmptyStateHTML('No department volume in current filter');
  }
  if (els.reportTrendChart) {
    els.reportTrendChart.innerHTML = tasks.length
      ? renderTrendChart(trendStats)
      : chartEmptyStateHTML('No workload trend in current filter');
  }
  if (els.reportPriorityChart) {
    els.reportPriorityChart.innerHTML = tasks.length
      ? renderPriorityChart(priorityStats)
      : chartEmptyStateHTML('No priority mix in current filter');
  }
  if (els.reportStatusDonut) {
    els.reportStatusDonut.innerHTML = tasks.length
      ? renderStatusDonutChart(statusStats, tasks.length, closureRate)
      : chartEmptyStateHTML('No status mix in current filter');
  }
  if (els.reportStatusLegend) {
    els.reportStatusLegend.innerHTML = tasks.length
      ? renderStatusLegend(statusStats)
      : '';
  }

  els.reportResults.innerHTML = tasks.length
    ? tasks.map((task) => taskCardHTML(task, { singleAction: 'View Detail' })).join('')
    : emptyStateHTML('No tasks in report', 'Adjust the report filters and try again.');
}

function getReportViewModeLabel() {
  return { daily: 'Daily', weekly: 'Weekly', monthly: 'Monthly' }[state.reportViewMode] || 'Daily';
}

function getWeekStart(date) {
  const copy = new Date(date);
  const day = copy.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  copy.setDate(copy.getDate() + diff);
  copy.setHours(0, 0, 0, 0);
  return copy;
}

function getTrendBucketMeta(date, viewMode) {
  const source = new Date(date);
  let bucketDate;

  if (viewMode === 'weekly') {
    bucketDate = getWeekStart(source);
  } else if (viewMode === 'monthly') {
    bucketDate = new Date(source.getFullYear(), source.getMonth(), 1);
    bucketDate.setHours(0, 0, 0, 0);
  } else {
    bucketDate = new Date(source.getFullYear(), source.getMonth(), source.getDate());
    bucketDate.setHours(0, 0, 0, 0);
  }

  const key = viewMode === 'monthly'
    ? `${bucketDate.getFullYear()}-${String(bucketDate.getMonth() + 1).padStart(2, '0')}`
    : formatDateInput(bucketDate);

  const label = viewMode === 'weekly'
    ? `${String(bucketDate.getDate()).padStart(2, '0')} ${bucketDate.toLocaleDateString('en-GB', { month: 'short' })}`
    : viewMode === 'monthly'
      ? bucketDate.toLocaleDateString('en-GB', { month: 'short', year: '2-digit' })
      : bucketDate.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });

  return {
    key,
    label,
    sortValue: bucketDate.getTime(),
    bucketDate
  };
}

function getNextTrendBucketDate(date, viewMode) {
  const next = new Date(date);
  if (viewMode === 'weekly') next.setDate(next.getDate() + 7);
  else if (viewMode === 'monthly') next.setMonth(next.getMonth() + 1);
  else next.setDate(next.getDate() + 1);
  next.setHours(0, 0, 0, 0);
  return next;
}

function getTrendRangeBounds(tasks, viewMode) {
  const parseBound = (value, end = false) => {
    if (!value) return null;
    const date = new Date(`${value}T00:00:00`);
    if (end) date.setHours(23, 59, 59, 999);
    return date;
  };

  let start = parseBound(els.reportStartDate.value, false);
  let end = parseBound(els.reportEndDate.value, true);

  const allDates = [];
  tasks.forEach((task) => {
    if (task.createdAt) allDates.push(new Date(task.createdAt));
    if (task.closedAt) allDates.push(new Date(task.closedAt));
  });

  if (!start && allDates.length) start = new Date(Math.min(...allDates.map((item) => item.getTime())));
  if (!end && allDates.length) end = new Date(Math.max(...allDates.map((item) => item.getTime())));
  if (!start || !end) return null;

  return {
    start: getTrendBucketMeta(start, viewMode).bucketDate,
    end: getTrendBucketMeta(end, viewMode).bucketDate
  };
}

function getTrendStats(tasks, viewMode) {
  const grouped = new Map();
  const ensureBucket = (dateValue) => {
    const meta = getTrendBucketMeta(dateValue, viewMode);
    if (!grouped.has(meta.key)) {
      grouped.set(meta.key, {
        key: meta.key,
        label: meta.label,
        sortValue: meta.sortValue,
        open: 0,
        closed: 0
      });
    }
    return grouped.get(meta.key);
  };

  const bounds = getTrendRangeBounds(tasks, viewMode);
  if (bounds) {
    let cursor = new Date(bounds.start);
    let safety = 0;
    while (cursor.getTime() <= bounds.end.getTime() && safety < 180) {
      ensureBucket(cursor);
      cursor = getNextTrendBucketDate(cursor, viewMode);
      safety += 1;
    }
  }

  tasks.forEach((task) => {
    const openBucket = ensureBucket(task.createdAt || task.updatedAt);
    openBucket.open += 1;

    if (task.status === 'Closed' && task.closedAt) {
      const closedBucket = ensureBucket(task.closedAt);
      closedBucket.closed += 1;
    }
  });

  const maxBuckets = { daily: 8, weekly: 10, monthly: 12 }[viewMode] || 8;
  return Array.from(grouped.values())
    .sort((a, b) => a.sortValue - b.sortValue)
    .slice(-maxBuckets);
}

function renderTrendChart(stats) {
  const width = 560;
  const height = 300;
  const chartX = 44;
  const chartY = 72;
  const chartW = width - 88;
  const chartH = 160;
  const maxValue = Math.max(...stats.map((item) => Math.max(item.open, item.closed)), 1);
  const groupGap = 18;
  const groupWidth = Math.max(34, Math.min(62, (chartW - (groupGap * Math.max(stats.length - 1, 0))) / Math.max(stats.length, 1)));
  const barGap = 8;
  const barWidth = Math.max(12, (groupWidth - barGap) / 2);
  const totalGroupsWidth = (stats.length * groupWidth) + (Math.max(stats.length - 1, 0) * groupGap);
  const startX = chartX + ((chartW - totalGroupsWidth) / 2);
  const openTotal = stats.reduce((sum, item) => sum + item.open, 0);
  const closedTotal = stats.reduce((sum, item) => sum + item.closed, 0);
  const gridLines = [0.25, 0.5, 0.75, 1].map((ratio) => {
    const y = chartY + chartH - (chartH * ratio);
    const value = Math.round(maxValue * ratio);
    return `
      <g>
        <line x1="${chartX}" y1="${y}" x2="${chartX + chartW}" y2="${y}" stroke="#e7ece7" stroke-width="1"></line>
        <text x="${chartX - 8}" y="${y + 4}" text-anchor="end" class="axis-label">${value}</text>
      </g>`;
  }).join('');

  const groups = stats.map((item, index) => {
    const groupX = startX + index * (groupWidth + groupGap);
    const openHeight = item.open ? Math.max(6, (item.open / maxValue) * chartH) : 0;
    const closedHeight = item.closed ? Math.max(6, (item.closed / maxValue) * chartH) : 0;
    const openY = chartY + chartH - openHeight;
    const closedY = chartY + chartH - closedHeight;

    return `
      <g>
        <rect x="${groupX}" y="${openY}" width="${barWidth}" height="${openHeight}" rx="10" fill="#3178c6"></rect>
        <rect x="${groupX + barWidth + barGap}" y="${closedY}" width="${barWidth}" height="${closedHeight}" rx="10" fill="#2f8f5b"></rect>
        ${item.open ? `<text x="${groupX + (barWidth / 2)}" y="${openY - 8}" text-anchor="middle" class="bar-value">${item.open}</text>` : ''}
        ${item.closed ? `<text x="${groupX + barWidth + barGap + (barWidth / 2)}" y="${closedY - 8}" text-anchor="middle" class="bar-value">${item.closed}</text>` : ''}
        <text x="${groupX + (groupWidth / 2)}" y="${chartY + chartH + 22}" text-anchor="middle" class="axis-label">${escapeHtml(item.label)}</text>
      </g>`;
  }).join('');

  return `
    <svg class="report-chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeHtml(getReportViewModeLabel())} open versus closed trend chart">
      <text x="44" y="20" font-size="12" font-weight="700" fill="#7c867f">${escapeHtml(getReportViewModeLabel())} open vs closed trend</text>
      <g>
        <circle cx="44" cy="42" r="6" fill="#3178c6"></circle>
        <text x="56" y="46" class="trend-legend-label">Open</text>
        <circle cx="108" cy="42" r="6" fill="#2f8f5b"></circle>
        <text x="120" y="46" class="trend-legend-label">Closed</text>
      </g>
      <text x="516" y="22" text-anchor="end" font-size="12" font-weight="800" fill="#1f2a23">${openTotal} open / ${closedTotal} closed</text>
      ${gridLines}
      <line x1="${chartX}" y1="${chartY + chartH}" x2="${chartX + chartW}" y2="${chartY + chartH}" stroke="#d9dfd6" stroke-width="1.5"></line>
      ${groups}
    </svg>`;
}

function renderPriorityChart(stats) {
  const width = 520;
  const rowHeight = 52;
  const height = Math.max(230, 44 + (stats.length * rowHeight));
  const labelX = 12;
  const barX = 136;
  const barMaxWidth = 300;
  const countX = 462;
  const maxValue = Math.max(...stats.map((item) => item.count), 1);

  const rows = stats.map((item, index) => {
    const y = 34 + (index * rowHeight);
    const barWidth = item.count ? Math.max(14, (item.count / maxValue) * barMaxWidth) : 0;
    const fill = PRIORITY_COLORS[item.priority] || '#1f4d3a';
    return `
      <g>
        <text x="${labelX}" y="${y + 18}" font-size="13" font-weight="700" fill="#1f2a23">${escapeHtml(item.priority)}</text>
        <text x="${labelX}" y="${y + 35}" font-size="11" font-weight="600" fill="#7c867f">${item.active} active / ${item.share}%</text>
        <rect x="${barX}" y="${y + 6}" width="${barMaxWidth}" height="14" rx="7" fill="#edf1ee"></rect>
        <rect x="${barX}" y="${y + 6}" width="${barWidth}" height="14" rx="7" fill="${fill}"></rect>
        <text x="${countX}" y="${y + 18}" text-anchor="end" font-size="14" font-weight="800" fill="#1f2a23">${item.count}</text>
        <text x="${countX}" y="${y + 35}" text-anchor="end" font-size="11" font-weight="600" fill="#7c867f">selected tasks</text>
      </g>`;
  }).join('');

  return `
    <svg class="report-chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Priority breakdown chart">
      <text x="12" y="20" font-size="12" font-weight="700" fill="#7c867f">Workload urgency by priority</text>
      ${rows}
    </svg>`;
}

function renderDepartmentChart(stats) {
  const sorted = [...stats].sort((a, b) => b.count - a.count);
  const maxValue = Math.max(...sorted.map((item) => item.count), 1);
  const rowHeight = 52;
  const width = 520;
  const height = Math.max(230, 44 + (sorted.length * rowHeight));
  const labelX = 12;
  const barX = 142;
  const barMaxWidth = 300;
  const countX = 462;

  const rows = sorted.map((item, index) => {
    const y = 34 + (index * rowHeight);
    const barWidth = item.count ? Math.max(14, (item.count / maxValue) * barMaxWidth) : 0;
    const fill = DEPARTMENT_COLORS[item.department] || '#1f4d3a';
    return `
      <g>
        <text x="${labelX}" y="${y + 18}" font-size="13" font-weight="700" fill="#1f2a23">${escapeHtml(item.department)}</text>
        <text x="${labelX}" y="${y + 35}" font-size="11" font-weight="600" fill="#7c867f">${item.share}% of workload</text>
        <rect x="${barX}" y="${y + 6}" width="${barMaxWidth}" height="14" rx="7" fill="#edf1ee"></rect>
        <rect x="${barX}" y="${y + 6}" width="${barWidth}" height="14" rx="7" fill="${fill}"></rect>
        <text x="${countX}" y="${y + 18}" text-anchor="end" font-size="14" font-weight="800" fill="#1f2a23">${item.count}</text>
        <text x="${countX}" y="${y + 35}" text-anchor="end" font-size="11" font-weight="600" fill="#7c867f">${item.closed} closed / ${item.overdue} overdue</text>
      </g>`;
  }).join('');

  return `
    <svg class="report-chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Department workload chart">
      <text x="12" y="20" font-size="12" font-weight="700" fill="#7c867f">Task volume by department</text>
      ${rows}
    </svg>`;
}

function renderStatusDonutChart(stats, total, closureRate) {
  const radius = 72;
  const circumference = 2 * Math.PI * radius;
  const centerX = 110;
  const centerY = 110;
  let offset = 0;

  const arcs = stats.filter((item) => item.count > 0).map((item) => {
    const portion = item.count / Math.max(total, 1);
    const dash = portion * circumference;
    const arc = `<circle cx="${centerX}" cy="${centerY}" r="${radius}" fill="none" stroke="${STATUS_COLORS[item.status] || '#1f4d3a'}" stroke-width="22" stroke-linecap="butt" stroke-dasharray="${dash} ${circumference - dash}" stroke-dashoffset="${-offset}" transform="rotate(-90 ${centerX} ${centerY})"></circle>`;
    offset += dash;
    return arc;
  }).join('');

  return `
    <svg class="report-chart-svg" viewBox="0 0 220 220" role="img" aria-label="Status distribution chart">
      <circle cx="${centerX}" cy="${centerY}" r="${radius}" fill="none" stroke="#edf1ee" stroke-width="22"></circle>
      ${arcs}
      <circle cx="${centerX}" cy="${centerY}" r="48" fill="#ffffff"></circle>
      <text x="${centerX}" y="${centerY - 2}" text-anchor="middle" font-size="28" class="report-donut-center">${total}</text>
      <text x="${centerX}" y="${centerY + 18}" text-anchor="middle" font-size="12" class="report-donut-subtext">tasks in scope</text>
      <text x="${centerX}" y="${centerY + 36}" text-anchor="middle" font-size="11" class="report-donut-subtext">${closureRate}% closure</text>
    </svg>`;
}

function renderStatusLegend(stats) {
  return stats.map((item) => `
    <div class="report-legend__item">
      <div class="report-legend__main">
        <span class="report-legend__swatch" style="background:${STATUS_COLORS[item.status] || '#1f4d3a'}"></span>
        <div>
          <div class="report-legend__label">${escapeHtml(item.status)}</div>
          <div class="report-legend__meta">${item.share}% of selected workload</div>
        </div>
      </div>
      <div class="report-legend__value">${item.count}</div>
    </div>
  `).join('');
}

function chartEmptyStateHTML(message) {
  return `<div class="report-chart-empty">${escapeHtml(message)}</div>`;
}

function renderTaskMedia(task) {
  const media = Array.isArray(task.mediaAttachments) ? task.mediaAttachments : [];
  const hasMedia = media.length > 0;
  els.detailMediaCard.classList.toggle('hidden', !hasMedia);
  els.detailMediaList.innerHTML = hasMedia ? renderMediaGallery(media) : '';
}

function getModReports() {
  if (state.backendMode === 'firebase') return state.modReportsCache.map((report) => normalizeModReport(report));
  return (JSON.parse(localStorage.getItem(STORAGE_KEYS.modReports) || '[]')).map(normalizeModReport);
}

function saveModReports(reports) {
  const normalized = reports.map((report) => normalizeModReport(report));
  if (state.backendMode === 'firebase') {
    const previous = state.modReportsCache.map((report) => ({ ...report }));
    state.modReportsCache = normalized.map((report) => ({ ...report }));
    syncCollectionDiff(FIREBASE_COLLECTIONS.modReports, previous, state.modReportsCache, stripModReportForFirestore, getModReportDocId)
      .catch((error) => handleBackgroundSyncError(error, 'modReports'));
    return;
  }
  localStorage.setItem(STORAGE_KEYS.modReports, JSON.stringify(normalized));
}

function normalizeModReport(report) {
  return {
    id: report.id || `mod_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    reportNo: report.reportNo || generateModReportNo(getModReportsRawSafe()),
    area: report.area || 'Lobby',
    location: report.location || '-',
    department: report.department || 'Engineering',
    category: report.category || 'Other',
    priority: report.priority || 'High',
    subject: report.subject || '-',
    detail: report.detail || '',
    actionNote: report.actionNote || '',
    status: report.status || 'Open',
    openedByName: report.openedByName || '-',
    openedByDepartment: report.openedByDepartment || '-',
    createdAt: report.createdAt || new Date().toISOString(),
    updatedAt: report.updatedAt || report.createdAt || new Date().toISOString(),
    linkedTaskId: report.linkedTaskId || '',
    linkedTaskTicketNo: report.linkedTaskTicketNo || '',
    attachments: Array.isArray(report.attachments) ? report.attachments : [],
    logs: Array.isArray(report.logs) ? report.logs : []
  };
}

function getModReportsRawSafe() {
  try {
    if (state.backendMode === 'firebase') return state.modReportsCache.map((report) => ({ ...report }));
    return JSON.parse(localStorage.getItem(STORAGE_KEYS.modReports) || '[]');
  } catch (error) {
    return [];
  }
}

function generateModReportNo(reports) {
  const today = new Date();
  const datePart = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
  const count = reports.filter((item) => String(item.reportNo || '').includes(datePart)).length + 1;
  return `MOD-${datePart}-${String(count).padStart(4, '0')}`;
}

async function onModMediaSelected(event) {
  const files = Array.from(event.target.files || []);
  state.modDraftMedia = await Promise.all(files.map(serializeMediaFile));
  renderModMediaPreview();
}

async function serializeMediaFile(file) {
  const base = {
    id: `media_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    name: file.name,
    type: file.type,
    size: file.size,
    kind: file.type.startsWith('video/') ? 'video' : 'image',
    dataUrl: '',
    inlineSaved: false,
    file
  };
  if (file.size > MOD_MEDIA_INLINE_LIMIT) return base;
  const dataUrl = await readFileAsDataUrl(file);
  return { ...base, dataUrl, inlineSaved: true };
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(reader.error || new Error('File read failed'));
    reader.readAsDataURL(file);
  });
}

function renderModMediaPreview() {
  const items = state.modDraftMedia || [];
  els.modMediaPreview.innerHTML = items.length
    ? renderMediaGallery(items, { removable: true })
    : emptyStateHTML('No media selected', 'Attach photo or video evidence for this MOD finding.');
}

function renderMediaGallery(items, options = {}) {
  return items.map((item, index) => `
    <div class="mod-media-card">
      <div class="mod-media-card__preview">
        ${item.kind === 'video'
          ? (getMediaPreviewSrc(item)
              ? `<video src="${escapeHtml(getMediaPreviewSrc(item))}" controls muted playsinline></video>`
              : `<div class="mod-media-placeholder">Video file<br>${escapeHtml(item.name)}</div>`)
          : (getMediaPreviewSrc(item)
              ? `<img src="${escapeHtml(getMediaPreviewSrc(item))}" alt="${escapeHtml(item.name)}" />`
              : `<div class="mod-media-placeholder">Image file<br>${escapeHtml(item.name)}</div>`)}
      </div>
      <div class="mod-media-card__meta">
        <div class="mod-media-card__name">${escapeHtml(item.name)}</div>
        <div class="mod-media-card__sub">${item.kind === 'video' ? 'Video' : 'Photo'} / ${formatFileSize(item.size)}${item.inlineSaved ? '' : ' / filename only'}</div>
      </div>
      ${options.removable ? `<button class="link-btn" type="button" data-mod-remove-media="${index}">Remove</button>` : ''}
    </div>
  `).join('');
}

function formatFileSize(size) {
  if (!size) return '0 KB';
  if (size >= 1024 * 1024) return `${(size / (1024 * 1024)).toFixed(1)} MB`;
  return `${Math.max(1, Math.round(size / 1024))} KB`;
}

async function submitModReport() {
  if (!canAccessMod()) {
    alert('This page is available for MOD / CGM role.');
    return;
  }

  const area = els.modArea.value.trim();
  const location = document.getElementById('mod-location').value.trim();
  const department = els.modDepartment.value.trim();
  const category = document.getElementById('mod-category').value;
  const priority = els.modPriority.value.trim();
  const subject = document.getElementById('mod-subject').value.trim();
  const detail = document.getElementById('mod-detail').value.trim();
  const actionNote = document.getElementById('mod-action-note').value.trim();

  if (!area || !location || !department || !priority || !subject) {
    alert('Please fill required fields: area, location, owner department, severity, and title.');
    return;
  }

  const reports = getModReports();
  const now = new Date().toISOString();
  const reportNo = generateModReportNo(reports);
  const attachments = await prepareAttachmentsForSave(state.modDraftMedia || [], `${FIREBASE_STORAGE_FOLDERS.modMedia}/${reportNo}`);

  let linkedTaskId = '';
  let linkedTaskTicketNo = '';
  if (els.modCreateTaskToggle.checked) {
    const task = createTaskFromModFinding({ reportNo, location, department, category, priority, subject, detail, actionNote, attachments });
    linkedTaskId = task.id;
    linkedTaskTicketNo = task.ticketNo;
  }

  const report = normalizeModReport({
    id: `mod_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    reportNo,
    area,
    location,
    department,
    category,
    priority,
    subject,
    detail,
    actionNote,
    status: linkedTaskId ? 'In Progress' : 'Open',
    openedByName: state.currentUser.name,
    openedByDepartment: state.currentUser.department,
    createdAt: now,
    updatedAt: now,
    linkedTaskId,
    linkedTaskTicketNo,
    attachments,
    logs: [{ action: 'Reported', note: actionNote || subject, byName: state.currentUser.name, byDepartment: state.currentUser.department, createdAt: now }]
  });

  if (linkedTaskId) {
    report.logs.unshift({ action: 'Task Opened', note: `Linked follow-up task ${linkedTaskTicketNo} created for ${department}.`, byName: state.currentUser.name, byDepartment: state.currentUser.department, createdAt: now });
  }

  reports.unshift(report);
  saveModReports(reports);
  state.currentModReportId = report.id;
  resetModForm(false);
  renderApp();
  showPage('mod');
  alert(linkedTaskTicketNo
    ? `MOD report submitted: ${report.reportNo} / Linked task: ${linkedTaskTicketNo}`
    : `MOD report submitted: ${report.reportNo}`);
}

function createTaskFromModFinding({ reportNo, location, department, category, priority, subject, detail, actionNote, attachments }) {
  const tasks = getTasks();
  const now = new Date().toISOString();
  const newTask = normalizeTask({
    id: generateId(),
    ticketNo: generateTicketNo(tasks),
    location,
    department,
    category: category === 'Engineering' ? 'Repair' : category,
    priority,
    subject: `[MOD] ${subject}`,
    detail: [detail, actionNote ? `MOD instruction: ${actionNote}` : '', `Source: ${reportNo}`].filter(Boolean).join('\n'),
    status: 'New',
    openedByName: state.currentUser.name,
    openedByDepartment: state.currentUser.department,
    sourceType: 'MOD',
    sourceReference: reportNo,
    mediaAttachments: attachments,
    createdAt: now,
    updatedAt: now,
    logs: [{ action: 'Created', note: `Task opened from MOD report ${reportNo}.`, byName: state.currentUser.name, byDepartment: state.currentUser.department, createdAt: now }]
  });
  tasks.unshift(newTask);
  saveTasks(tasks);
  return newTask;
}

function resetModForm(resetSelection = true) {
  if (els.modReportForm) els.modReportForm.reset();
  state.modDraftMedia = [];
  if (els.modMediaInput) els.modMediaInput.value = '';
  renderModMediaPreview();
  if (resetSelection) state.currentModReportId = '';
  setupChipGroup(document.getElementById('mod-area-chips'), els.modArea, 'Lobby', true);
  setupChipGroup(document.getElementById('mod-department-chips'), els.modDepartment, 'Engineering', true);
  setupChipGroup(document.getElementById('mod-priority-chips'), els.modPriority, 'High', true);
  if (els.modCreateTaskToggle) els.modCreateTaskToggle.checked = true;
}

function renderModPage() {
  if (!canAccessMod()) {
    els.modReportsList.innerHTML = emptyStateHTML('MOD access only', 'This page is available for MOD / CGM role.');
    els.modDetailPanel.innerHTML = emptyStateHTML('No access', 'Please login with MOD or CGM account.');
    els.modSummaryGrid.innerHTML = '';
    return;
  }
  renderModMediaPreview();
  renderModSummary();
  renderModReportList();
  renderModReportDetail();
}

function renderModSummary() {
  const reports = getModReports();
  const todayKey = new Date().toDateString();
  const openCount = reports.filter((item) => !['Resolved', 'Closed'].includes(item.status)).length;
  const highCount = reports.filter((item) => ['High', 'Urgent'].includes(item.priority) && !['Resolved', 'Closed'].includes(item.status)).length;
  const mediaCount = reports.filter((item) => (item.attachments || []).length > 0).length;
  const todayCount = reports.filter((item) => new Date(item.createdAt).toDateString() === todayKey).length;
  els.modSummaryGrid.innerHTML = [
    { label: 'Open Findings', value: openCount },
    { label: 'High Risk', value: highCount },
    { label: 'With Media', value: mediaCount },
    { label: 'Today Submitted', value: todayCount }
  ].map((card) => `
    <article class="card stat-card">
      <span class="stat-card__label">${escapeHtml(card.label)}</span>
      <strong class="stat-card__value">${card.value}</strong>
    </article>
  `).join('');
}

function getFilteredModReports() {
  const search = state.modSearch || '';
  return getModReports()
    .filter((item) => {
      if (state.modFilter === 'open') return !['Resolved', 'Closed'].includes(item.status);
      if (state.modFilter === 'high') return ['High', 'Urgent'].includes(item.priority);
      if (state.modFilter === 'media') return (item.attachments || []).length > 0;
      if (state.modFilter === 'mine') return item.openedByName === state.currentUser?.name;
      return true;
    })
    .filter((item) => (`${item.reportNo} ${item.location} ${item.subject} ${item.detail} ${item.area}`).toLowerCase().includes(search))
    .sort((a, b) => new Date(b.updatedAt || b.createdAt) - new Date(a.updatedAt || a.createdAt));
}

function renderModReportList() {
  const reports = getFilteredModReports();
  els.modCount.textContent = `${reports.length} report${reports.length === 1 ? '' : 's'}`;
  els.modReportsList.innerHTML = reports.length
    ? reports.map((report) => modReportCardHTML(report)).join('')
    : emptyStateHTML('No MOD report found', 'Try another filter or create a new finding above.');
}

function modReportCardHTML(report) {
  const active = state.currentModReportId === report.id ? ' is-selected' : '';
  return `
    <article class="card task-card mod-report-card${active}">
      <div class="task-card__top">
        <span class="task-card__ticket">${escapeHtml(report.reportNo)}</span>
        <span class="task-card__time">${timeAgo(report.updatedAt || report.createdAt)}</span>
      </div>
      <div class="task-card__meta">${escapeHtml(report.area)} / ${escapeHtml(report.location)}</div>
      <h3 class="task-card__title">${escapeHtml(report.subject)}</h3>
      <div class="task-card__badges">
        <span class="badge ${priorityBadgeClass(report.priority)}">${escapeHtml(report.priority)}</span>
        <span class="badge ${modStatusBadgeClass(report.status)}">${escapeHtml(report.status)}</span>
      </div>
      <div class="task-card__info">
        <span>${escapeHtml(report.department)} / ${escapeHtml(report.category)}</span>
        <span>Media: ${(report.attachments || []).length}</span>
      </div>
      <div class="task-card__actions">
        <button class="btn btn-secondary" type="button" data-mod-open="${escapeHtml(report.id)}">Open</button>
        ${report.linkedTaskId ? `<button class="btn btn-primary" type="button" data-task-view="${escapeHtml(report.linkedTaskId)}">View Task</button>` : `<button class="btn btn-primary" type="button" data-mod-open="${escapeHtml(report.id)}">Review</button>`}
      </div>
    </article>
  `;
}

function renderModReportDetail() {
  const reports = getModReports();
  const report = reports.find((item) => item.id === state.currentModReportId) || reports[0] || null;
  if (report && !state.currentModReportId) state.currentModReportId = report.id;
  if (!report) {
    els.modDetailHint.textContent = 'Open one report to review media and follow-up';
    els.modDetailPanel.innerHTML = emptyStateHTML('No MOD report yet', 'Create the first inspection finding from the form above.');
    return;
  }
  els.modDetailHint.textContent = `${report.reportNo} / ${report.department}`;
  const logs = [...(report.logs || [])].sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  els.modDetailPanel.innerHTML = `
    <div class="mod-detail-head">
      <div>
        <p class="task-card__ticket">${escapeHtml(report.reportNo)}</p>
        <h3 class="task-hero__title">${escapeHtml(report.subject)}</h3>
        <p class="task-hero__meta">${escapeHtml(report.area)} / ${escapeHtml(report.location)}</p>
      </div>
      <div class="task-card__badges">
        <span class="badge ${priorityBadgeClass(report.priority)}">${escapeHtml(report.priority)}</span>
        <span class="badge ${modStatusBadgeClass(report.status)}">${escapeHtml(report.status)}</span>
      </div>
    </div>
    <dl class="detail-list">
      <div><dt>Owner Dept</dt><dd>${escapeHtml(report.department)}</dd></div>
      <div><dt>Category</dt><dd>${escapeHtml(report.category)}</dd></div>
      <div><dt>Opened by</dt><dd>${escapeHtml(report.openedByName)} / ${escapeHtml(report.openedByDepartment)}</dd></div>
      <div><dt>Linked Task</dt><dd>${report.linkedTaskId ? `<button class="link-btn" type="button" data-task-view="${escapeHtml(report.linkedTaskId)}">${escapeHtml(report.linkedTaskTicketNo || 'Open task')}</button>` : '-'}</dd></div>
      <div><dt>Created</dt><dd>${formatDateTime(report.createdAt)}</dd></div>
      <div><dt>Updated</dt><dd>${formatDateTime(report.updatedAt)}</dd></div>
    </dl>
    <div class="card mod-detail-block">
      <h4 class="block-title">Issue Found</h4>
      <p class="detail-description">${escapeHtml(report.detail || 'No detail provided.')}</p>
      <h4 class="block-title">Immediate Action / Instruction</h4>
      <p class="detail-description">${escapeHtml(report.actionNote || '-')}</p>
    </div>
    <div class="card mod-detail-block">
      <div class="section__header section__header--home">
        <div>
          <h4 class="block-title">Media Evidence</h4>
          <span class="section__hint">${(report.attachments || []).length} file(s)</span>
        </div>
      </div>
      <div class="mod-media-grid">${(report.attachments || []).length ? renderMediaGallery(report.attachments) : emptyStateHTML('No media attached', 'No photo or video was attached to this MOD report.')}</div>
    </div>
    <div class="action-grid mod-detail-actions">
      ${getModDetailActions(report).map((action) => `<button class="btn ${action.primary ? 'btn-primary' : 'btn-secondary'}" type="button" data-mod-action="${action.value}" data-mod-id="${report.id}">${action.label}</button>`).join('')}
    </div>
    <div class="card mod-detail-block">
      <h4 class="block-title">Timeline</h4>
      <div class="timeline">${logs.length ? logs.map((log) => timelineItemHTML(log)).join('') : emptyStateHTML('No timeline yet', 'Updates will appear here.')}</div>
    </div>
  `;
}

function getModDetailActions(report) {
  const actions = [];
  if (report.status === 'Open') actions.push({ value: 'start', label: 'Start Follow-up', primary: true });
  if (report.status === 'In Progress') actions.push({ value: 'resolve', label: 'Mark Resolved', primary: true });
  if (report.status === 'Resolved') actions.push({ value: 'close', label: 'Close Report', primary: true });
  if (report.status === 'Closed') actions.push({ value: 'reopen', label: 'Reopen', primary: false });
  if (!report.linkedTaskId) actions.push({ value: 'createTask', label: 'Create Task', primary: false });
  return actions;
}

function onModFilterClick(event) {
  const chip = event.target.closest('[data-mod-filter]');
  if (!chip) return;
  state.modFilter = chip.dataset.modFilter;
  Array.from(els.modFilterChips.querySelectorAll('.chip')).forEach((node) => node.classList.toggle('is-active', node === chip));
  renderModReportList();
  renderModReportDetail();
}

function onModMediaPreviewClick(event) {
  const removeBtn = event.target.closest('[data-mod-remove-media]');
  if (!removeBtn) return;
  state.modDraftMedia.splice(Number(removeBtn.dataset.modRemoveMedia), 1);
  renderModMediaPreview();
}

function onModSearchInput(event) {
  state.modSearch = event.target.value.trim().toLowerCase();
  renderModReportList();
  renderModReportDetail();
}

function onModReportListClick(event) {
  const openBtn = event.target.closest('[data-mod-open]');
  const taskBtn = event.target.closest('[data-task-view]');
  if (taskBtn) {
    openTaskDetail(taskBtn.dataset.taskView);
    return;
  }
  if (!openBtn) return;
  state.currentModReportId = openBtn.dataset.modOpen;
  renderModReportList();
  renderModReportDetail();
}

function onModDetailClick(event) {
  const taskBtn = event.target.closest('[data-task-view]');
  if (taskBtn) {
    openTaskDetail(taskBtn.dataset.taskView);
    return;
  }
  const actionBtn = event.target.closest('[data-mod-action]');
  if (!actionBtn) return;
  runModReportAction(actionBtn.dataset.modAction, actionBtn.dataset.modId);
}

function runModReportAction(action, reportId) {
  const reports = getModReports();
  const report = reports.find((item) => item.id === reportId);
  if (!report) return;
  const now = new Date().toISOString();
  if (action === 'createTask' && !report.linkedTaskId) {
    const task = createTaskFromModFinding({
      reportNo: report.reportNo,
      location: report.location,
      department: report.department,
      category: report.category,
      priority: report.priority,
      subject: report.subject,
      detail: report.detail,
      actionNote: report.actionNote,
      attachments: report.attachments || []
    });
    updateModReport(reportId, (draft) => {
      draft.linkedTaskId = task.id;
      draft.linkedTaskTicketNo = task.ticketNo;
      draft.status = draft.status === 'Open' ? 'In Progress' : draft.status;
      draft.updatedAt = now;
      draft.logs.unshift({ action: 'Task Opened', note: `Linked follow-up task ${task.ticketNo} created.`, byName: state.currentUser.name, byDepartment: state.currentUser.department, createdAt: now });
      return draft;
    });
    renderApp();
    return;
  }

  const transitionMap = {
    start: { status: 'In Progress', label: 'Follow-up Started' },
    resolve: { status: 'Resolved', label: 'Resolved' },
    close: { status: 'Closed', label: 'Closed' },
    reopen: { status: 'In Progress', label: 'Reopened' }
  };
  const step = transitionMap[action];
  if (!step) return;
  updateModReport(reportId, (draft) => {
    draft.status = step.status;
    draft.updatedAt = now;
    draft.logs.unshift({ action: step.label, note: `${step.label} by ${state.currentUser.name}.`, byName: state.currentUser.name, byDepartment: state.currentUser.department, createdAt: now });
    return draft;
  });
  renderApp();
}

function updateModReport(reportId, updater) {
  const updated = getModReports().map((report) => {
    if (report.id !== reportId) return report;
    return normalizeModReport(updater({ ...report, logs: [...(report.logs || [])], attachments: [...(report.attachments || [])] }));
  });
  saveModReports(updated);
}

function modStatusBadgeClass(status) {
  return {
    'Open': 'badge-status-new',
    'In Progress': 'badge-status-progress',
    'Resolved': 'badge-status-done',
    'Closed': 'badge-status-closed'
  }[status] || 'badge-status-new';
}

function onTaskCardClick(event) {
  const viewButton = event.target.closest('[data-task-view]');
  const actionButton = event.target.closest('[data-task-action]');
  if (viewButton) {
    openTaskDetail(viewButton.dataset.taskView);
    return;
  }
  if (actionButton) {
    runTaskTransition(actionButton.dataset.taskAction, actionButton.dataset.taskId);
  }
}

function onDetailActionClick(event) {
  const button = event.target.closest('[data-detail-action]');
  if (!button) return;
  if (button.dataset.detailAction === 'focus-note') {
    els.detailNoteInput.focus();
    return;
  }
  runTaskTransition(button.dataset.detailAction, state.currentTaskId);
}

function onTaskTabClick(event) {
  const tab = event.target.closest('[data-status]');
  if (!tab) return;
  state.taskStatusFilter = tab.dataset.status;
  Array.from(els.taskStatusTabs.querySelectorAll('.tab')).forEach((node) => node.classList.toggle('is-active', node === tab));
  renderTaskList();
}

function onTasksSearch(event) {
  state.taskSearch = event.target.value.trim().toLowerCase();
  renderTaskList();
}

function toggleHighFilter() {
  state.highOnly = !state.highOnly;
  els.taskFilterHigh.classList.toggle('is-active', state.highOnly);
  renderTaskList();
}

function onCreateTask(event) {
  event.preventDefault();

  const location = document.getElementById('task-location').value.trim();
  const department = els.taskDepartment.value.trim();
  const category = document.getElementById('task-category').value;
  const priority = els.taskPriority.value.trim();
  const subject = document.getElementById('task-subject').value.trim();
  const detail = document.getElementById('task-detail').value.trim();
  const selectedAssigneeId = els.createAssignSelect?.value || '';
  const allowedDepartments = getAllowedCreateDepartments();
  const canAssignNow = selectedAssigneeId && canAssignToDepartment(department);
  const selectedAssignee = canAssignNow ? users.find((user) => user.employeeId === selectedAssigneeId) : null;

  if (!canCreateTasks()) {
    alert('This role cannot create tasks.');
    return;
  }

  if (!location || !department || !priority || !subject) {
    alert('Please fill required fields: location, department, priority, and subject.');
    return;
  }

  if (!allowedDepartments.includes(department)) {
    alert(`This role cannot dispatch task to ${department}.`);
    return;
  }

  const tasks = getTasks();
  const now = new Date().toISOString();
  const logs = [
    {
      action: 'Created',
      note: subject,
      byName: state.currentUser.name,
      byDepartment: state.currentUser.department,
      createdAt: now
    }
  ];

  const newTask = normalizeTask({
    id: generateId(),
    ticketNo: generateTicketNo(tasks),
    location,
    department,
    category,
    priority,
    subject,
    detail,
    status: canAssignNow && selectedAssignee ? 'Accepted' : 'New',
    openedByName: state.currentUser.name,
    openedByDepartment: state.currentUser.department,
    assignedToName: selectedAssignee?.name || '',
    assignedByName: selectedAssignee ? state.currentUser.name : '',
    assignedByDepartment: selectedAssignee ? state.currentUser.department : '',
    createdAt: now,
    updatedAt: now,
    assignedAt: selectedAssignee ? now : '',
    logs
  });

  if (canAssignNow && selectedAssignee) {
    newTask.logs.unshift({
      action: 'Assigned',
      note: `Assigned to ${selectedAssignee.name} during task creation.`,
      byName: state.currentUser.name,
      byDepartment: state.currentUser.department,
      createdAt: now
    });
  }

  tasks.unshift(newTask);
  saveTasks(tasks);
  els.createTaskForm.reset();
  if (els.createAssignSelect) els.createAssignSelect.value = '';
  setupChipGroup(document.getElementById('department-chips'), els.taskDepartment, getAllowedCreateDepartments()[0] || 'FO', true, renderCreateAssignmentState);
  setupChipGroup(document.getElementById('priority-chips'), els.taskPriority, 'Medium', true);
  renderApp();
  alert(selectedAssignee
    ? `Task created and assigned to ${selectedAssignee.name}: ${newTask.ticketNo}`
    : `Task created successfully: ${newTask.ticketNo}`);
  showPage('tasks');
}

function renderCreateDepartmentChoices() {
  const allowed = getAllowedCreateDepartments();
  const chips = Array.from(document.querySelectorAll('#department-chips [data-value]'));
  chips.forEach((chip) => {
    const isAllowed = allowed.includes(chip.dataset.value);
    chip.classList.toggle('hidden', !isAllowed);
  });
  if (allowed.length && !allowed.includes(els.taskDepartment.value)) {
    els.taskDepartment.value = allowed[0];
  }
}

function renderCreateAssignmentState() {
  if (!els.createAssignWrap || !state.currentUser) return;
  const selectedDepartment = els.taskDepartment?.value || 'FO';
  const canAssign = selectedDepartment && canAssignToDepartment(selectedDepartment);
  els.createAssignWrap.classList.toggle('hidden', !canAssign);
  if (!canAssign) {
    populateAssigneeSelect(els.createAssignSelect, [], '', 'Immediate assign available for own team manager / CGM only');
    return;
  }

  const assignees = getAssignableUsers(selectedDepartment);
  populateAssigneeSelect(els.createAssignSelect, assignees, '', `Optional: assign immediately to ${selectedDepartment} team`);
  if (els.createAssignHelper) {
    els.createAssignHelper.textContent = isCGM()
      ? 'CGM can assign immediately to any department team.'
      : 'DOR can assign immediately only to FO team.';
  }
}

function renderDetailAssignmentState(task) {
  if (!els.detailAssignCard) return;
  const canAssign = !!task && canManageAssignments(task) && !['Pending FO Closure', 'Closed'].includes(task.status);
  els.detailAssignCard.classList.toggle('hidden', !canAssign);
  if (!canAssign) return;

  const assignees = getAssignableUsers(task.department);
  const selectedUser = assignees.find((user) => user.name === task.assignedToName);
  populateAssigneeSelect(
    els.detailAssignSelect,
    assignees,
    selectedUser?.employeeId || '',
    task.assignedToName ? `Currently assigned to ${task.assignedToName}. Choose another team member to reassign.` : 'No owner yet. Assign this task to a team member.'
  );
  if (els.detailAssignHint) {
    els.detailAssignHint.textContent = task.assignedToName
      ? `Current owner: ${task.assignedToName} / ${task.department}`
      : `No owner yet / ${task.department} team`;
  }
  if (els.detailAssignBtn) {
    els.detailAssignBtn.textContent = task.assignedToName ? 'Reassign Task' : 'Assign Task';
  }
}

function assignCurrentTask() {
  const task = getTaskById(state.currentTaskId);
  if (!task) return;
  if (!canManageAssignments(task)) {
    alert('You cannot assign this task.');
    return;
  }

  const selectedAssigneeId = els.detailAssignSelect?.value || '';
  const assignee = users.find((user) => user.employeeId === selectedAssigneeId);
  if (!assignee) {
    alert('Please select a team member first.');
    return;
  }

  const note = els.detailAssignNote?.value.trim() || '';
  performAssignment(task.id, assignee, note);
  if (els.detailAssignNote) els.detailAssignNote.value = '';
  renderApp();
  showPage('detail');
}

function performAssignment(taskId, assignee, note = '') {
  updateTask(taskId, (draft) => {
    const now = new Date().toISOString();
    const previousAssignee = draft.assignedToName || '';
    const action = previousAssignee && previousAssignee !== assignee.name ? 'Reassigned' : 'Assigned';
    const baseNote = previousAssignee && previousAssignee !== assignee.name
      ? `Reassigned from ${previousAssignee} to ${assignee.name}`
      : `Assigned to ${assignee.name}`;
    draft.assignedToName = assignee.name;
    draft.assignedByName = state.currentUser.name;
    draft.assignedByDepartment = state.currentUser.department;
    draft.assignedAt = now;
    if (isAssignableRole(assignee.role) && draft.status !== 'Closed') {
      draft.status = 'New';
      draft.startedAt = '';
      draft.doneAt = '';
      draft.closedAt = '';
    }
    draft.updatedAt = now;
    draft.logs.unshift({
      action,
      note: note ? `${baseNote} / ${note}` : `${baseNote}.`,
      byName: state.currentUser.name,
      byDepartment: state.currentUser.department,
      createdAt: now
    });
    return draft;
  });
}

function populateAssigneeSelect(selectEl, people, selectedValue = '', emptyLabel = 'Select team member') {
  if (!selectEl) return;
  const options = [`<option value="">${escapeHtml(emptyLabel)}</option>`].concat(
    people.map((user) => `<option value="${escapeHtml(user.employeeId)}" ${selectedValue === user.employeeId ? 'selected' : ''}>${escapeHtml(user.name)} / ${escapeHtml(user.role)}</option>`)
  );
  selectEl.innerHTML = options.join('');
}

function getAssignableUsers(department) {
  return users
    .filter((user) => user.department === department && isAssignableRole(user.role))
    .sort((a, b) => a.name.localeCompare(b.name));
}

function canAssignToDepartment(department) {
  if (isCGM()) return true;
  if (isDOR()) return department === 'FO';
  return isDepartmentManager() && state.currentUser?.department === department;
}

function canManageAssignments(task) {
  if (isCGM()) return true;
  if (isDOR()) return task.department === 'FO';
  return isDepartmentManager() && state.currentUser?.department === task.department;
}

function canWorkOnTask(task) {
  if (isCGM()) return true;
  return task.assignedToName === state.currentUser?.name;
}

function saveDetailNote() {
  const note = els.detailNoteInput.value.trim();
  if (!note || !state.currentTaskId) {
    alert('Please type a note first.');
    return;
  }

  updateTask(state.currentTaskId, (task) => {
    task.logs.unshift(createLog('Note Added', note));
    task.updatedAt = new Date().toISOString();
    return task;
  });

  els.detailNoteInput.value = '';
  renderApp();
  showPage('detail');
}

function runTaskTransition(action, taskId) {
  const task = getTaskById(taskId);
  if (!task) return;
  const transitions = {
    accept: { status: 'Accepted', logAction: 'Accepted', note: 'Task accepted by staff.' },
    start: { status: 'In Progress', logAction: 'In Progress', note: 'Work started.' },
    submitReview: { status: 'Waiting Supervisor Review', logAction: 'Submitted to Manager', note: 'Work submitted to department manager for review.' },
    resubmitReview: { status: 'Waiting Supervisor Review', logAction: 'Resubmitted to Manager', note: 'Rework completed and submitted back to department manager.' },
    managerReturn: { status: 'Returned for Rework', logAction: 'Returned for Rework', note: 'Department manager returned this task for rework.' },
    managerApprove: { status: 'Pending FO Closure', logAction: 'Submitted to FO', note: 'Department manager approved the work and sent it to FO / DOR for closure.' },
    close: { status: 'Closed', logAction: 'Closed', note: 'Task closed by FO / DOR.' },
    reopen: { status: 'In Progress', logAction: 'Reopened', note: 'Task reopened.' }
  };
  if (!transitions[action]) return;
  if (!canTransitionTask(task, action)) {
    alert('This action is not allowed for the current user or task status.');
    return;
  }

  const draftNote = (els.detailNoteInput?.value || '').trim();
  if ((action === 'submitReview' || action === 'resubmitReview') && !(task.mediaAttachments || []).length) {
    alert('Please upload at least 1 photo or video before sending work back to manager.');
    return;
  }

  updateTask(taskId, (draft) => {
    const now = new Date().toISOString();
    const transition = transitions[action];
    draft.status = transition.status;
    draft.updatedAt = now;
    if (action === 'accept') {
      draft.assignedToName = state.currentUser.name;
      draft.assignedAt = now;
    }
    if (action === 'start' || action === 'reopen') draft.startedAt = now;
    if (action === 'submitReview' || action === 'resubmitReview' || action === 'managerApprove') draft.doneAt = now;
    if (action === 'managerReturn') {
      draft.doneAt = '';
      draft.closedAt = '';
    }
    if (action === 'close') draft.closedAt = now;
    if (action === 'reopen') {
      draft.closedAt = '';
      draft.doneAt = '';
    }
    const combinedNote = draftNote ? `${transition.note} / ${draftNote}` : transition.note;
    draft.logs.unshift(createLog(transition.logAction, combinedNote));
    return draft;
  });
  if (els.detailNoteInput) els.detailNoteInput.value = '';
  renderApp();
  if (state.currentView === 'detail') showPage('detail');
}

function openTaskDetail(taskId) {
  state.currentTaskId = taskId;
  showPage('detail');
}

function getDetailActions(task) {
  const actions = [];
  const assignedWorker = isDepartmentStaff() && task.assignedToName === state.currentUser?.name;
  const deptManager = isDepartmentManager() && state.currentUser?.department === task.department;
  const canClose = isCGM() || isFODispatcher();

  if (task.status === 'New' && assignedWorker) {
    actions.push({ value: 'accept', label: 'รับงาน', primary: true, full: true });
  }
  if (task.status === 'Accepted' && assignedWorker) {
    actions.push({ value: 'start', label: 'เริ่มทำงาน', primary: true, full: true });
  }
  if (task.status === 'In Progress' && assignedWorker) {
    actions.push({ value: 'focus-note', label: 'เพิ่มหมายเหตุ', primary: false });
    actions.push({ value: 'submitReview', label: 'ส่งงานกลับหัวหน้า', primary: true });
  }
  if (task.status === 'Returned for Rework' && assignedWorker) {
    actions.push({ value: 'focus-note', label: 'เพิ่มหมายเหตุ', primary: false });
    actions.push({ value: 'resubmitReview', label: 'แก้ไขแล้วส่งกลับหัวหน้า', primary: true });
  }
  if (task.status === 'Waiting Supervisor Review' && deptManager) {
    actions.push({ value: 'managerReturn', label: 'ส่งกลับให้แก้ไข', primary: false });
    actions.push({ value: 'managerApprove', label: 'รับรองงาน / ส่งต่อ FO', primary: true });
  }
  if (task.status === 'Pending FO Closure' && canClose) {
    actions.push({ value: 'close', label: 'ปิดงาน', primary: true, full: true });
  }
  if (task.status === 'Closed' && canClose) {
    actions.push({ value: 'reopen', label: 'เปิดงานใหม่', primary: false, full: true });
  }
  return actions;
}

function canTransitionTask(task, action) {
  return getDetailActions(task).some((item) => item.value === action);
}

function getTaskListResults() {
  if (isDepartmentManager()) return getSupervisorTaskScope();
  return getTaskContextFilteredTasks(getVisibleTasks(getTasks()))
    .filter((task) => state.taskStatusFilter === 'All' || task.status === state.taskStatusFilter)
    .filter((task) => !state.highOnly || ['High', 'Urgent'].includes(task.priority))
    .filter((task) => matchesSearch(task, state.taskSearch))
    .sort(sortTasks);
}

function getHistoryResults() {
  const search = els.historySearch.value.trim().toLowerCase();
  return getVisibleTasks(getTasks())
    .filter((task) => taskWithinRange(task, els.historyStartDate.value, els.historyEndDate.value))
    .filter((task) => els.historyDepartment.value === 'All' || task.department === els.historyDepartment.value)
    .filter((task) => els.historyStatus.value === 'All' || task.status === els.historyStatus.value)
    .filter((task) => els.historyPriority.value === 'All' || task.priority === els.historyPriority.value)
    .filter((task) => matchesSearch(task, search))
    .sort((a, b) => new Date(b.updatedAt || b.createdAt) - new Date(a.updatedAt || a.createdAt));
}

function getReportResults() {
  return getVisibleTasks(getTasks())
    .filter((task) => taskWithinRange(task, els.reportStartDate.value, els.reportEndDate.value))
    .filter((task) => els.reportDepartment.value === 'All' || task.department === els.reportDepartment.value)
    .filter((task) => els.reportStatus.value === 'All' || task.status === els.reportStatus.value)
    .sort((a, b) => new Date(b.updatedAt || b.createdAt) - new Date(a.updatedAt || a.createdAt));
}

function onHistoryPresetClick(event) {
  const chip = event.target.closest('[data-preset]');
  if (!chip) return;
  setDatePreset('history', chip.dataset.preset);
  renderHistoryPage();
}

function onReportPresetClick(event) {
  const chip = event.target.closest('[data-preset]');
  if (!chip) return;
  setDatePreset('report', chip.dataset.preset);
  renderReportsPage();
}

function onReportViewModeClick(event) {
  const chip = event.target.closest('[data-view]');
  if (!chip) return;
  state.reportViewMode = chip.dataset.view;
  Array.from(els.reportViewModes.querySelectorAll('.chip')).forEach((node) => node.classList.toggle('is-active', node === chip));
  renderReportsPage();
}

function setDatePreset(type, preset, silent = false) {
  const startEl = type === 'history' ? els.historyStartDate : els.reportStartDate;
  const endEl = type === 'history' ? els.historyEndDate : els.reportEndDate;
  const chipContainer = type === 'history' ? els.historyPresets : els.reportPresets;
  const today = new Date();
  const endDate = formatDateInput(today);
  let startDate = endDate;

  if (preset === '7days') {
    const start = new Date(today);
    start.setDate(today.getDate() - 6);
    startDate = formatDateInput(start);
  } else if (preset === 'month') {
    const start = new Date(today.getFullYear(), today.getMonth(), 1);
    startDate = formatDateInput(start);
  } else if (preset === 'all') {
    startDate = '';
  }

  startEl.value = startDate;
  endEl.value = preset === 'all' ? '' : endDate;
  Array.from(chipContainer.querySelectorAll('.chip')).forEach((node) => node.classList.toggle('is-active', node.dataset.preset === preset));
  if (type === 'history') state.historyPreset = preset;
  if (type === 'report') state.reportPreset = preset;
  if (!silent && type === 'report') renderReportsPage();
}

function exportReportCsv() {
  const rows = getReportResults();
  const header = ['Ticket No', 'Location', 'Department', 'Category', 'Priority', 'Status', 'Opened By', 'Assigned To', 'Created At', 'Updated At'];
  const csvRows = [header.join(',')].concat(rows.map((task) => [
    task.ticketNo,
    task.location,
    task.department,
    task.category,
    task.priority,
    task.status,
    `${task.openedByName} / ${task.openedByDepartment}`,
    task.assignedToName || '-',
    formatDateTime(task.createdAt),
    formatDateTime(task.updatedAt || task.createdAt)
  ].map(csvEscape).join(',')));

  downloadBlob(csvRows.join('\n'), `laya-service-hub-report-${formatFileDate(new Date())}.csv`, 'text/csv;charset=utf-8;');
}

function exportReportPdf() {
  const rows = getReportResults();
  const jspdfLib = window.jspdf?.jsPDF;
  if (!jspdfLib || typeof window.jspdf?.jsPDF !== 'function' || typeof window.jspdf?.jsPDF?.API === 'undefined') {
    alert('PDF library not loaded. Please connect to the internet once or deploy this app online, then try again.');
    return;
  }

  const doc = new jspdfLib({ orientation: 'p', unit: 'pt', format: 'a4' });
  if (typeof doc.autoTable !== 'function') {
    alert('PDF table plugin not loaded. Please refresh the page and try again.');
    return;
  }

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 40;
  const summaryTop = 118;
  const summaryWidth = (pageWidth - (margin * 2) - 18) / 2;
  const summaryHeight = 60;

  const reportMeta = {
    total: rows.length,
    closed: rows.filter((t) => t.status === 'Closed').length,
    open: rows.filter((t) => t.status !== 'Closed').length,
    overdue: rows.filter((t) => isTaskOverdue(t)).length
  };
  const closureRate = rows.length ? Math.round((reportMeta.closed / rows.length) * 100) : 0;

  const deptStatsDetailed = DEPARTMENTS.map((department) => {
    const deptTasks = rows.filter((t) => t.department === department);
    return {
      department,
      count: deptTasks.length,
      closed: deptTasks.filter((t) => t.status === 'Closed').length,
      overdue: deptTasks.filter((t) => isTaskOverdue(t)).length,
      share: rows.length ? Math.round((deptTasks.length / rows.length) * 100) : 0
    };
  });

  const statusStats = STATUSES.map((status) => {
    const count = rows.filter((t) => t.status === status).length;
    return {
      status,
      count,
      share: rows.length ? Math.round((count / rows.length) * 100) : 0
    };
  });

  const priorityStatsDetailed = PRIORITIES.map((priority) => {
    const priorityTasks = rows.filter((t) => t.priority === priority);
    const count = priorityTasks.length;
    return {
      priority,
      count,
      active: priorityTasks.filter((t) => !['Closed', 'Pending FO Closure'].includes(t.status)).length,
      share: rows.length ? Math.round((count / rows.length) * 100) : 0
    };
  });

  const trendStatsDetailed = getTrendStats(rows, state.reportViewMode);

  doc.setFillColor(24, 55, 43);
  doc.rect(0, 0, pageWidth, 92, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(22);
  doc.text('Laya Service Hub Performance Report', margin, 44);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  doc.text(`Excom-ready export - ${buildReportRangeLabel()}`, margin, 66);
  doc.text(`Generated ${formatDateTime(new Date().toISOString())}`, margin, 82);

  drawPdfSummaryBox(doc, margin, summaryTop, summaryWidth, summaryHeight, 'Total Tasks', String(reportMeta.total), 'Selected workload');
  drawPdfSummaryBox(doc, margin + summaryWidth + 18, summaryTop, summaryWidth, summaryHeight, 'Closed', `${reportMeta.closed} (${closureRate}%)`, 'Completion rate');
  drawPdfSummaryBox(doc, margin, summaryTop + summaryHeight + 14, summaryWidth, summaryHeight, 'Open', String(reportMeta.open), 'Still active');
  drawPdfSummaryBox(doc, margin + summaryWidth + 18, summaryTop + summaryHeight + 14, summaryWidth, summaryHeight, 'Overdue', String(reportMeta.overdue), 'Escalation candidates');

  let cursorY = summaryTop + (summaryHeight * 2) + 42;
  doc.setTextColor(31, 42, 35);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(14);
  doc.text('Executive Summary', margin, cursorY);
  cursorY += 16;

  const execLines = buildExecutiveSummary(rows, deptStatsDetailed, { closed: reportMeta.closed, open: reportMeta.open, overdue: reportMeta.overdue, closureRate });
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  execLines.forEach((line) => {
    const wrapped = doc.splitTextToSize(`- ${line}`, pageWidth - (margin * 2));
    doc.text(wrapped, margin, cursorY);
    cursorY += (wrapped.length * 14);
  });
  cursorY += 10;

  const chartTop = cursorY;
  const chartHeight = 184;
  const chartGap = 18;
  const chartWidth = (pageWidth - (margin * 2) - chartGap) / 2;
  drawPdfBarChart(doc, margin, chartTop, chartWidth, chartHeight, deptStatsDetailed);
  drawPdfDonutChart(doc, margin + chartWidth + chartGap, chartTop, chartWidth, chartHeight, statusStats, reportMeta.total, closureRate);
  const trendTop = chartTop + chartHeight + 20;
  drawPdfTrendChart(doc, margin, trendTop, pageWidth - (margin * 2), 196, trendStatsDetailed);
  cursorY = trendTop + 196 + 24;

  const deptStatsTable = deptStatsDetailed.map((item) => [item.department, String(item.count), `${item.closed}`, `${item.overdue}`]);
  doc.autoTable({
    startY: cursorY,
    head: [['Department', 'Tasks', 'Closed', 'Overdue']],
    body: deptStatsTable,
    margin: { left: margin, right: margin },
    styles: { fontSize: 10, cellPadding: 6 },
    headStyles: { fillColor: [31, 77, 58] },
    alternateRowStyles: { fillColor: [248, 250, 248] },
    theme: 'grid'
  });

  const afterDeptY = doc.lastAutoTable.finalY + 18;
  const priorityTable = priorityStatsDetailed.map((item) => [item.priority, String(item.count), `${item.active}`, `${item.share}%`]);
  doc.autoTable({
    startY: afterDeptY,
    head: [['Priority', 'Tasks', 'Active', 'Share']],
    body: priorityTable,
    margin: { left: margin, right: margin },
    styles: { fontSize: 10, cellPadding: 6 },
    headStyles: { fillColor: [176, 139, 87] },
    alternateRowStyles: { fillColor: [250, 248, 244] },
    theme: 'grid'
  });

  const afterPriorityY = doc.lastAutoTable.finalY + 18;
  const taskRows = rows.length ? rows.map((task) => ([
    task.ticketNo,
    `${task.location} / ${task.department}`,
    task.subject,
    task.priority,
    task.status,
    formatDateTime(task.updatedAt || task.createdAt)
  ])) : [['-', '-', 'No tasks in selected report range', '-', '-', '-']];

  doc.autoTable({
    startY: afterPriorityY,
    head: [['Ticket', 'Location / Dept', 'Subject', 'Priority', 'Status', 'Updated']],
    body: taskRows,
    margin: { left: margin, right: margin },
    styles: { fontSize: 9, cellPadding: 5, overflow: 'linebreak' },
    headStyles: { fillColor: [176, 139, 87] },
    alternateRowStyles: { fillColor: [250, 248, 244] },
    columnStyles: {
      0: { cellWidth: 86 },
      1: { cellWidth: 98 },
      2: { cellWidth: 160 },
      3: { cellWidth: 52 },
      4: { cellWidth: 62 },
      5: { cellWidth: 76 }
    },
    theme: 'grid',
    didDrawPage: ({ pageNumber }) => {
      doc.setFontSize(9);
      doc.setTextColor(124, 134, 127);
      doc.text(`Laya Service Hub / Page ${pageNumber}`, margin, pageHeight - 18);
    }
  });

  doc.save(`laya-service-hub-report-${formatFileDate(new Date())}.pdf`);
}

function buildReportRangeLabel() {
  let range = 'All dates';
  if (els.reportStartDate.value && els.reportEndDate.value) range = `${els.reportStartDate.value} to ${els.reportEndDate.value}`;
  else if (els.reportStartDate.value || els.reportEndDate.value) range = els.reportStartDate.value || els.reportEndDate.value;
  return `${range} - ${getReportViewModeLabel()} view`;
}

function drawPdfBarChart(doc, x, y, width, height, stats) {
  doc.setDrawColor(222, 229, 223);
  doc.setFillColor(255, 255, 255);
  doc.roundedRect(x, y, width, height, 14, 14, 'FD');
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  doc.setTextColor(31, 42, 35);
  doc.text('Department Volume', x + 14, y + 18);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.setTextColor(124, 134, 127);
  doc.text('Bar chart - workload by team', x + 14, y + 32);

  const chartX = x + 92;
  const chartY = y + 52;
  const chartW = width - 124;
  const rowGap = 26;
  const barH = 10;
  const sorted = [...stats].sort((a, b) => b.count - a.count);
  const maxCount = Math.max(...sorted.map((item) => item.count), 1);

  sorted.forEach((item, index) => {
    const rowY = chartY + (index * rowGap);
    const fillColor = hexToRgb(DEPARTMENT_COLORS[item.department] || '#1f4d3a');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    doc.setTextColor(31, 42, 35);
    doc.text(item.department, x + 14, rowY + 8);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(124, 134, 127);
    doc.text(`${item.share}%`, x + 14, rowY + 20);

    doc.setFillColor(237, 241, 238);
    doc.roundedRect(chartX, rowY, chartW, barH, 5, 5, 'F');
    const fillW = item.count ? Math.max(12, (item.count / maxCount) * chartW) : 0;
    doc.setFillColor(fillColor.r, fillColor.g, fillColor.b);
    if (fillW > 0) doc.roundedRect(chartX, rowY, fillW, barH, 5, 5, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(31, 42, 35);
    doc.text(String(item.count), chartX + chartW + 8, rowY + 8);
  });
}

function drawPdfDonutChart(doc, x, y, width, height, stats, total, closureRate) {
  doc.setDrawColor(222, 229, 223);
  doc.setFillColor(255, 255, 255);
  doc.roundedRect(x, y, width, height, 14, 14, 'FD');
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  doc.setTextColor(31, 42, 35);
  doc.text('Status Distribution', x + 14, y + 18);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.setTextColor(124, 134, 127);
  doc.text('Donut chart - current task mix', x + 14, y + 32);

  const centerX = x + 74;
  const centerY = y + 102;
  const radius = 42;
  const innerRadius = 24;
  let startAngle = -90;

  doc.setDrawColor(237, 241, 238);
  doc.setLineWidth(14);
  doc.circle(centerX, centerY, radius, 'S');

  stats.filter((item) => item.count > 0).forEach((item) => {
    const sweep = (item.count / Math.max(total, 1)) * 360;
    const rgb = hexToRgb(STATUS_COLORS[item.status] || '#1f4d3a');
    drawPdfArc(doc, centerX, centerY, radius, startAngle, startAngle + sweep, 14, rgb);
    startAngle += sweep;
  });

  doc.setFillColor(255, 255, 255);
  doc.circle(centerX, centerY, innerRadius, 'F');
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(20);
  doc.setTextColor(31, 42, 35);
  doc.text(String(total), centerX, centerY - 2, { align: 'center' });
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.setTextColor(124, 134, 127);
  doc.text('tasks', centerX, centerY + 12, { align: 'center' });
  doc.text(`${closureRate}% closure`, centerX, centerY + 24, { align: 'center' });

  let legendY = y + 58;
  stats.forEach((item) => {
    const rgb = hexToRgb(STATUS_COLORS[item.status] || '#1f4d3a');
    doc.setFillColor(rgb.r, rgb.g, rgb.b);
    doc.circle(x + 156, legendY - 3, 4, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    doc.setTextColor(31, 42, 35);
    doc.text(item.status, x + 166, legendY);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(124, 134, 127);
    doc.text(`${item.count} / ${item.share}%`, x + width - 16, legendY, { align: 'right' });
    legendY += 21;
  });
}

function drawPdfTrendChart(doc, x, y, width, height, stats) {
  doc.setDrawColor(222, 229, 223);
  doc.setFillColor(255, 255, 255);
  doc.roundedRect(x, y, width, height, 14, 14, 'FD');
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  doc.setTextColor(31, 42, 35);
  doc.text('Open vs Closed Trend', x + 14, y + 18);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.setTextColor(124, 134, 127);
  doc.text(`${getReportViewModeLabel()} view - opened vs closed tickets`, x + 14, y + 32);

  if (!stats.length) {
    doc.text('No trend data in current filter.', x + 14, y + 58);
    return;
  }

  const legendY = y + 48;
  doc.setFillColor(49, 120, 198);
  doc.circle(x + 18, legendY, 4, 'F');
  doc.setTextColor(31, 42, 35);
  doc.setFont('helvetica', 'bold');
  doc.text('Open', x + 28, legendY + 3);
  doc.setFillColor(47, 143, 91);
  doc.circle(x + 82, legendY, 4, 'F');
  doc.text('Closed', x + 92, legendY + 3);

  const chartX = x + 34;
  const chartY = y + 74;
  const chartW = width - 68;
  const chartH = 86;
  const maxValue = Math.max(...stats.map((item) => Math.max(item.open, item.closed)), 1);
  const groupGap = 10;
  const groupWidth = Math.max(22, Math.min(42, (chartW - (groupGap * Math.max(stats.length - 1, 0))) / Math.max(stats.length, 1)));
  const barGap = 4;
  const barWidth = Math.max(8, (groupWidth - barGap) / 2);
  const totalGroupsWidth = (stats.length * groupWidth) + (Math.max(stats.length - 1, 0) * groupGap);
  const startX = chartX + ((chartW - totalGroupsWidth) / 2);

  doc.setDrawColor(231, 236, 231);
  [0.25, 0.5, 0.75, 1].forEach((ratio) => {
    const lineY = chartY + chartH - (chartH * ratio);
    doc.line(chartX, lineY, chartX + chartW, lineY);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.setTextColor(124, 134, 127);
    doc.text(String(Math.round(maxValue * ratio)), chartX - 6, lineY + 3, { align: 'right' });
  });

  stats.forEach((item, index) => {
    const groupX = startX + index * (groupWidth + groupGap);
    const openH = item.open ? Math.max(4, (item.open / maxValue) * chartH) : 0;
    const closedH = item.closed ? Math.max(4, (item.closed / maxValue) * chartH) : 0;

    doc.setFillColor(49, 120, 198);
    if (openH > 0) doc.roundedRect(groupX, chartY + chartH - openH, barWidth, openH, 3, 3, 'F');
    doc.setFillColor(47, 143, 91);
    if (closedH > 0) doc.roundedRect(groupX + barWidth + barGap, chartY + chartH - closedH, barWidth, closedH, 3, 3, 'F');

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.setTextColor(124, 134, 127);
    doc.text(item.label, groupX + (groupWidth / 2), chartY + chartH + 12, { align: 'center' });
  });
}

function drawPdfArc(doc, cx, cy, radius, startDeg, endDeg, lineWidth, rgb) {
  const step = 4;
  const points = [];
  for (let angle = startDeg; angle <= endDeg; angle += step) {
    const rad = (Math.PI / 180) * angle;
    points.push([cx + Math.cos(rad) * radius, cy + Math.sin(rad) * radius]);
  }
  const endRad = (Math.PI / 180) * endDeg;
  points.push([cx + Math.cos(endRad) * radius, cy + Math.sin(endRad) * radius]);
  doc.setDrawColor(rgb.r, rgb.g, rgb.b);
  doc.setLineWidth(lineWidth);
  for (let i = 1; i < points.length; i += 1) {
    doc.line(points[i - 1][0], points[i - 1][1], points[i][0], points[i][1]);
  }
}

function hexToRgb(hex) {
  const normalized = hex.replace('#', '');
  const bigint = parseInt(normalized, 16);
  return {
    r: (bigint >> 16) & 255,
    g: (bigint >> 8) & 255,
    b: bigint & 255
  };
}


function buildExecutiveSummary(tasks, deptStats, meta) {
  const sortedDepartments = [...deptStats].sort((a, b) => b.count - a.count);
  const busiest = sortedDepartments.find((item) => item.count > 0);
  const mostOverdue = [...deptStats].sort((a, b) => b.overdue - a.overdue)[0];
  const urgentOpen = tasks.filter((task) => ['High', 'Urgent'].includes(task.priority) && !['Pending FO Closure', 'Closed'].includes(task.status)).length;
  const oldestOpen = [...tasks]
    .filter((task) => !['Pending FO Closure', 'Closed'].includes(task.status))
    .sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt))[0];

  const lines = [
    `Closure rate is ${meta.closureRate}% with ${meta.closed} closed task(s) out of ${tasks.length || 0} in the selected period.`,
    busiest ? `${busiest.department} is currently the busiest department with ${busiest.count} task(s) in scope.` : 'No department activity is recorded in the selected report range.',
    `${urgentOpen} open high-priority / urgent task(s) remain active and may need close monitoring.`,
    mostOverdue?.overdue ? `${mostOverdue.department} currently carries the most overdue pressure with ${mostOverdue.overdue} overdue task(s).` : 'No overdue pressure is visible in the selected report range.',
    oldestOpen ? `Oldest active ticket is ${oldestOpen.ticketNo} for ${oldestOpen.location} (${oldestOpen.subject}).` : 'There are no aging open tickets right now.'
  ];

  return lines;
}

function metricRowHTML({ label, value, meta, share, tone = 'department' }) {
  return `
    <div class="metric-row metric-row--${escapeHtml(tone)}">
      <div class="metric-row__head">
        <span class="metric-row__label">${escapeHtml(label)}</span>
        <span class="metric-row__value">${escapeHtml(value)}</span>
      </div>
      <div class="metric-row__bar"><span class="metric-row__fill" style="--fill:${Math.max(0, Math.min(100, share || 0))}%"></span></div>
      <div class="metric-row__meta">${escapeHtml(meta)}</div>
    </div>
  `;
}

function drawPdfSummaryBox(doc, x, y, width, height, label, value, foot) {
  doc.setDrawColor(219, 226, 220);
  doc.setFillColor(255, 255, 255);
  doc.roundedRect(x, y, width, height, 12, 12, 'FD');
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);
  doc.setTextColor(95, 107, 99);
  doc.text(label, x + 14, y + 18);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(22);
  doc.setTextColor(31, 42, 35);
  doc.text(value, x + 14, y + 40);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.setTextColor(124, 134, 127);
  doc.text(foot, x + 14, y + 54);
}

function getVisibleTasks(tasks) {
  if (isMOD()) return [];
  if (canSeeAllHotelTasks()) return tasks;
  if (isDepartmentManager()) return tasks.filter((task) => task.department === state.currentUser.department);
  if (isDepartmentStaff()) return tasks.filter((task) => task.assignedToName === state.currentUser.name);
  return tasks.filter((task) => task.openedByName === state.currentUser.name || task.assignedToName === state.currentUser.name);
}

function getTasks() {
  if (state.backendMode === 'firebase') return state.tasksCache.map((task) => normalizeTask(task));
  return (JSON.parse(localStorage.getItem(STORAGE_KEYS.tasks) || '[]')).map(normalizeTask);
}

function saveTasks(tasks) {
  const normalized = tasks.map((task) => normalizeTask(task));
  if (state.backendMode === 'firebase') {
    const previous = state.tasksCache.map((task) => ({ ...task }));
    state.tasksCache = normalized.map((task) => ({ ...task }));
    syncCollectionDiff(FIREBASE_COLLECTIONS.tasks, previous, state.tasksCache, stripTaskForFirestore, getTaskDocId)
      .catch((error) => handleBackgroundSyncError(error, 'tasks'));
    return;
  }
  localStorage.setItem(STORAGE_KEYS.tasks, JSON.stringify(normalized));
}

function getTaskById(taskId) {
  return getTasks().find((task) => task.id === taskId || task.ticketNo === taskId) || null;
}

function updateTask(taskId, updater) {
  const tasks = getTasks();
  const updated = tasks.map((task) => {
    if (task.id !== taskId && task.ticketNo !== taskId) return task;
    return normalizeTask(updater({ ...task, logs: [...(task.logs || [])] }));
  });
  saveTasks(updated);
}

function renderTeamAdminSection() {
  if (!els.teamAdminSection) return;
  const canManage = canManageOwnTeamDirectory();
  els.teamAdminSection.classList.toggle('hidden', !canManage);
  if (!canManage) return;
  const department = getManagedDepartment();
  const roleLabel = getManagedStaffRoleLabel();
  els.teamAdminTitle.textContent = `${department} Team Setup`;
  els.teamAdminHint.textContent = `Add or remove ${roleLabel} accounts in your team.`;
  const teamMembers = users.filter((user) => user.department === department && user.role === roleLabel).sort((a, b) => a.name.localeCompare(b.name));
  els.teamAdminList.innerHTML = teamMembers.length
    ? teamMembers.map((member) => `
        <div class="team-admin-row">
          <div class="team-admin-row__meta">
            <strong>${escapeHtml(member.name)}</strong>
            <div class="team-admin-row__sub">${escapeHtml(member.employeeId)} / ${escapeHtml(member.role)}</div>
          </div>
          <button class="btn btn-secondary" type="button" data-team-remove="${escapeHtml(member.employeeId)}">Remove</button>
        </div>
      `).join('')
    : emptyStateHTML('No team member yet', 'Add the first team member for this department.');
}

function resetTeamAdminForm() {
  if (els.teamMemberName) els.teamMemberName.value = '';
  if (els.teamMemberId) els.teamMemberId.value = '';
  if (els.teamMemberPassword) els.teamMemberPassword.value = '1234';
}

function onAddTeamMember(event) {
  event.preventDefault();
  if (!canManageOwnTeamDirectory()) return;
  const name = els.teamMemberName.value.trim();
  const employeeId = els.teamMemberId.value.trim();
  const password = (els.teamMemberPassword.value || '1234').trim() || '1234';
  if (!name || !employeeId) {
    alert('Please enter staff name and employee ID.');
    return;
  }
  if (users.some((user) => user.employeeId === employeeId)) {
    alert('Employee ID already exists.');
    return;
  }
  const department = getManagedDepartment();
  users.push({ employeeId, password, name, role: getManagedStaffRoleLabel(), department });
  saveUsers(users);
  resetTeamAdminForm();
  renderTeamAdminSection();
}

function onTeamAdminListClick(event) {
  const button = event.target.closest('[data-team-remove]');
  if (!button || !canManageOwnTeamDirectory()) return;
  const employeeId = button.dataset.teamRemove;
  users = users.filter((user) => user.employeeId !== employeeId);
  saveUsers(users);
  renderTeamAdminSection();
}

async function onDetailWorkMediaSelected(event) {
  const files = Array.from(event.target.files || []);
  state.detailDraftMedia = await Promise.all(files.map(serializeMediaFile));
  renderDetailWorkMediaPreview();
}

function onDetailWorkMediaPreviewClick(event) {
  const button = event.target.closest('[data-mod-remove-media]');
  if (!button) return;
  const index = Number(button.dataset.modRemoveMedia);
  state.detailDraftMedia.splice(index, 1);
  renderDetailWorkMediaPreview();
}

function renderDetailWorkMediaPreview() {
  if (!els.detailWorkMediaPreview) return;
  const items = state.detailDraftMedia || [];
  els.detailWorkMediaPreview.innerHTML = items.length
    ? renderMediaGallery(items, { removable: true })
    : emptyStateHTML('No work media selected', 'Attach photo or video from the department before sending work back to manager.');
}

function clearDetailWorkMediaDraft() {
  state.detailDraftMedia = [];
  if (els.detailWorkMediaInput) els.detailWorkMediaInput.value = '';
  renderDetailWorkMediaPreview();
}

async function saveDetailWorkMedia() {
  const task = getTaskById(state.currentTaskId);
  if (!task) return;
  if (!state.detailDraftMedia.length) {
    alert('Please choose photo or video first.');
    return;
  }
  const uploadedAttachments = await prepareAttachmentsForSave(state.detailDraftMedia, `${FIREBASE_STORAGE_FOLDERS.taskMedia}/${task.id}`);
  updateTask(task.id, (draft) => {
    draft.mediaAttachments = [...(draft.mediaAttachments || []), ...uploadedAttachments];
    draft.logs.unshift(createLog('Media Added', `Attached ${uploadedAttachments.length} work media file(s).`));
    draft.updatedAt = new Date().toISOString();
    return draft;
  });
  clearDetailWorkMediaDraft();
  renderApp();
  showPage('detail');
}

function renderDetailWorkUploadCard(task) {
  if (!els.detailWorkUploadCard) return;
  const visible = isDepartmentStaff() && task.assignedToName === state.currentUser?.name && ['Accepted', 'In Progress', 'Returned for Rework'].includes(task.status);
  els.detailWorkUploadCard.classList.toggle('hidden', !visible);
  if (visible && els.detailWorkUploadHint) {
    els.detailWorkUploadHint.textContent = 'Attach photo or video evidence before sending this task back to manager.';
  }
  renderDetailWorkMediaPreview();
}

function normalizeTask(task) {
  const normalizedStatus = task.status === 'Done' ? 'Pending FO Closure' : (task.status || 'New');
  return {
    id: task.id || task.ticketNo || generateId(),
    ticketNo: task.ticketNo,
    location: task.location || '-',
    department: task.department || 'FO',
    category: task.category || 'Other',
    priority: task.priority || 'Medium',
    subject: task.subject || '-',
    detail: task.detail || '',
    status: normalizedStatus,
    openedByName: task.openedByName || '-',
    openedByDepartment: task.openedByDepartment || '-',
    assignedToName: task.assignedToName || '',
    assignedByName: task.assignedByName || '',
    assignedByDepartment: task.assignedByDepartment || '',
    sourceType: task.sourceType || '',
    sourceReference: task.sourceReference || '',
    mediaAttachments: Array.isArray(task.mediaAttachments) ? task.mediaAttachments : [],
    createdAt: task.createdAt || new Date().toISOString(),
    updatedAt: task.updatedAt || task.createdAt || new Date().toISOString(),
    assignedAt: task.assignedAt || '',
    startedAt: task.startedAt || '',
    doneAt: task.doneAt || '',
    closedAt: task.closedAt || '',
    logs: Array.isArray(task.logs) ? task.logs : []
  };
}

function taskCardHTML(task, options = {}) {
  const overdueClass = isTaskOverdue(task) ? ' is-overdue' : '';
  const quickAction = getCardAction(task, options.singleAction);
  const modBadge = task.sourceType === 'MOD' ? '<span class="badge badge-status-accepted">MOD</span>' : '';
  return `
    <article class="card task-card${overdueClass}">
      <div class="task-card__top">
        <span class="task-card__ticket">${escapeHtml(task.ticketNo)}</span>
        <span class="task-card__time">${formatClock(task.updatedAt || task.createdAt)}</span>
      </div>
      <div class="task-card__meta">${escapeHtml(task.location)} / ${escapeHtml(task.department)}</div>
      <h3 class="task-card__title">${escapeHtml(task.subject)}</h3>
      <div class="task-card__badges">
        ${modBadge}
        <span class="badge ${priorityBadgeClass(task.priority)}">${escapeHtml(task.priority)}</span>
        <span class="badge ${statusBadgeClass(task.status)}">${escapeHtml(task.status)}</span>
      </div>
      <div class="task-card__info">
        <span>Opened by ${escapeHtml(task.openedByDepartment)}</span>
        <span>Assigned: ${escapeHtml(task.assignedToName || '-')}</span>
      </div>
      <div class="task-card__actions">
        <button class="btn btn-secondary" type="button" data-task-view="${escapeHtml(task.id)}">View</button>
        ${quickAction}
      </div>
    </article>
  `;
}

function getCardAction(task, forcedLabel) {
  if (forcedLabel) {
    return `<button class="btn btn-primary" type="button" data-task-view="${escapeHtml(task.id)}">${escapeHtml(forcedLabel)}</button>`;
  }
  const detailActions = getDetailActions(task);
  const primary = detailActions.find((action) => action.primary) || detailActions[0];
  if (!primary) {
    return `<button class="btn btn-primary" type="button" data-task-view="${escapeHtml(task.id)}">Open</button>`;
  }
  if (primary.value === 'focus-note') {
    return `<button class="btn btn-primary" type="button" data-task-view="${escapeHtml(task.id)}">Open</button>`;
  }
  return `<button class="btn btn-primary" type="button" data-task-action="${escapeHtml(primary.value)}" data-task-id="${escapeHtml(task.id)}">${escapeHtml(primary.label)}</button>`;
}

function timelineItemHTML(log) {
  return `
    <div class="timeline-item">
      <div class="timeline-item__dot"></div>
      <div class="timeline-item__content">
        <div class="timeline-item__title">${escapeHtml(log.action)} / ${escapeHtml(log.byName || '-')}</div>
        <div class="timeline-item__note">${escapeHtml(log.note || '-')}</div>
        <div class="timeline-item__time">${formatDateTime(log.createdAt)}</div>
      </div>
    </div>
  `;
}

function createLog(action, note) {
  return {
    action,
    note,
    byName: state.currentUser.name,
    byDepartment: state.currentUser.department,
    createdAt: new Date().toISOString()
  };
}

function setupChipGroup(container, hiddenInput, defaultValue, forceReset = false, onChange = null) {
  if (!hiddenInput) return;
  if (!container) {
    if (forceReset || !hiddenInput.value) {
      hiddenInput.value = defaultValue;
    }
    if (typeof onChange === 'function') onChange(hiddenInput.value);
    return;
  }

  const chips = Array.from(container.querySelectorAll('.chip'));
  if (forceReset || !hiddenInput.value) {
    hiddenInput.value = defaultValue;
    chips.forEach((chip) => chip.classList.toggle('is-active', chip.dataset.value === defaultValue));
  }

  chips.forEach((chip) => {
    if (chip.dataset.bound === 'true') return;
    chip.dataset.bound = 'true';
    chip.addEventListener('click', () => {
      chips.forEach((c) => c.classList.remove('is-active'));
      chip.classList.add('is-active');
      hiddenInput.value = chip.dataset.value;
      if (typeof onChange === 'function') onChange(hiddenInput.value);
    });
  });
}

function generateTicketNo(tasks) {
  const today = new Date();
  const y = today.getFullYear();
  const m = String(today.getMonth() + 1).padStart(2, '0');
  const d = String(today.getDate()).padStart(2, '0');
  const datePart = `${y}${m}${d}`;
  const sameDayCount = tasks.filter((task) => task.ticketNo && task.ticketNo.includes(datePart)).length + 1;
  return `LSH-${datePart}-${String(sameDayCount).padStart(4, '0')}`;
}

function generateId() {
  return `task_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function formatClock(dateString) {
  return new Date(dateString).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
}

function formatDateTime(dateString) {
  const date = new Date(dateString);
  return date.toLocaleString([], {
    year: 'numeric',
    month: 'short',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });
}

function formatDateInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatFileDate(date) {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`;
}

function priorityWeight(priority) {
  return { Urgent: 4, High: 3, Medium: 2, Low: 1 }[priority] || 0;
}

function sortTasks(a, b) {
  return priorityWeight(b.priority) - priorityWeight(a.priority) || new Date(b.updatedAt || b.createdAt) - new Date(a.updatedAt || a.createdAt);
}

function priorityBadgeClass(priority) {
  return {
    Low: 'badge-priority-low',
    Medium: 'badge-priority-medium',
    High: 'badge-priority-high',
    Urgent: 'badge-priority-urgent'
  }[priority] || 'badge-priority-low';
}

function statusBadgeClass(status) {
  return {
    New: 'badge-status-new',
    Accepted: 'badge-status-accepted',
    'In Progress': 'badge-status-progress',
    'Waiting Supervisor Review': 'badge-status-review',
    'Returned for Rework': 'badge-status-rework',
    'Pending FO Closure': 'badge-status-done',
    Closed: 'badge-status-closed'
  }[status] || 'badge-status-new';
}

function timeAgo(dateString) {
  const diffMs = Date.now() - new Date(dateString).getTime();
  const mins = Math.max(1, Math.floor(diffMs / 60000));
  if (mins < 60) return `${mins}m ago`;
  const hours = Math.floor(mins / 60);
  if (hours < 24) return `${hours}h ago`;
  const days = Math.floor(hours / 24);
  return `${days}d ago`;
}

function isTaskOverdue(task) {
  if (task.status === 'Closed' || task.status === 'Pending FO Closure') return false;
  const limit = OVERDUE_MINUTES[task.priority] || 120;
  const ageMinutes = (Date.now() - new Date(task.createdAt).getTime()) / 60000;
  return ageMinutes > limit;
}

function matchesSearch(task, search) {
  if (!search) return true;
  const haystack = `${task.ticketNo} ${task.location} ${task.subject} ${task.detail}`.toLowerCase();
  return haystack.includes(search);
}

function taskWithinRange(task, startDate, endDate) {
  const taskDate = getTaskRangeDate(task);
  if (startDate) {
    const start = new Date(`${startDate}T00:00:00`);
    if (taskDate < start) return false;
  }
  if (endDate) {
    const end = new Date(`${endDate}T23:59:59`);
    if (taskDate > end) return false;
  }
  return true;
}

function getTaskRangeDate(task) {
  if (task.closedAt) return new Date(task.closedAt);
  return new Date(task.updatedAt || task.createdAt);
}

function isCGM() {
  return state.currentUser?.role === 'CGM';
}

function isMOD() {
  return state.currentUser?.role === 'MOD';
}

function isFOStaff() {
  return state.currentUser?.role === 'FO Staff';
}

function isDOR() {
  return state.currentUser?.role === 'DOR';
}

function isFODispatcher() {
  return isFOStaff() || isDOR();
}

function isDepartmentStaff() {
  return ['HK Staff', 'ENG Staff', 'FB Staff'].includes(state.currentUser?.role || '');
}

function isDepartmentManager() {
  return ['Housekeeping Manager', 'ENG Manager', 'FB Manager'].includes(state.currentUser?.role || '');
}

function canAccessMod() {
  return isCGM() || isMOD();
}

function canViewReports() {
  return isCGM();
}

function canViewTaskList() {
  return !isMOD();
}

function canViewHistory() {
  return !isMOD();
}

function canCreateTasks() {
  return isFODispatcher() || isCGM();
}

function canSeeAllHotelTasks() {
  return isFODispatcher() || isCGM();
}

function canManageOwnTeamDirectory() {
  return isDOR() || isDepartmentManager();
}

function getManagedDepartment() {
  if (isDOR()) return 'FO';
  if (isDepartmentManager()) return state.currentUser?.department || '';
  return '';
}

function getManagedStaffRoleLabel() {
  return {
    FO: 'FO Staff',
    HK: 'HK Staff',
    Engineering: 'ENG Staff',
    FB: 'FB Staff'
  }[getManagedDepartment()] || 'Staff';
}

function isAssignableRole(role) {
  return ['FO Staff', 'HK Staff', 'ENG Staff', 'FB Staff'].includes(role);
}

function getAllowedCreateDepartments() {
  if (isFOStaff()) return ['FB', 'HK', 'Engineering'];
  if (isDOR() || isCGM()) return ['FO', 'FB', 'HK', 'Engineering'];
  return [];
}

function getDefaultLandingPage() {
  if (isCGM()) return 'dashboard';
  if (isMOD()) return 'mod';
  return 'home';
}

function getHomeNavPage() {
  if (isCGM()) return 'dashboard';
  if (isMOD()) return 'mod';
  return 'home';
}

function emptyStateHTML(title, description) {
  return `
    <div class="card empty-state">
      <h3>${escapeHtml(title)}</h3>
      <p>${escapeHtml(description)}</p>
    </div>
  `;
}

function csvEscape(value) {
  const safe = String(value ?? '').replace(/"/g, '""');
  return `"${safe}"`;
}

function downloadBlob(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function capitalize(value) {
  return value.charAt(0).toUpperCase() + value.slice(1);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}






