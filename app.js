const STORAGE_KEYS = {
  currentUser: 'lsh_current_user',
  tasks: 'lsh_tasks',
  modReports: 'lsh_mod_reports'
};

const OVERDUE_MINUTES = {
  Low: 180,
  Medium: 120,
  High: 60,
  Urgent: 30
};

const DEPARTMENTS = ['FO', 'FB', 'HK', 'Engineering'];
const STATUSES = ['New', 'Accepted', 'In Progress', 'Done', 'Closed'];
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
  Done: '#2f8f5b',
  Closed: '#6b736d'
};
const PRIORITY_COLORS = {
  Urgent: '#c64747',
  High: '#d88a1d',
  Medium: '#3178c6',
  Low: '#6b736d'
};

const users = [
  { employeeId: '11025', password: '1234', name: 'Noi', role: 'Manager', department: 'FO' },
  { employeeId: '11026', password: '1234', name: 'Pim', role: 'Supervisor', department: 'FO' },
  { employeeId: '12001', password: '1234', name: 'Anna', role: 'Staff', department: 'FO' },
  { employeeId: '12002', password: '1234', name: 'Beam', role: 'Staff', department: 'FO' },
  { employeeId: '22001', password: '1234', name: 'Joy', role: 'Supervisor', department: 'HK' },
  { employeeId: '22018', password: '1234', name: 'May', role: 'Staff', department: 'HK' },
  { employeeId: '22019', password: '1234', name: 'Fon', role: 'Staff', department: 'HK' },
  { employeeId: '33001', password: '1234', name: 'Lek', role: 'Supervisor', department: 'Engineering' },
  { employeeId: '33007', password: '1234', name: 'Art', role: 'Staff', department: 'Engineering' },
  { employeeId: '33008', password: '1234', name: 'Tom', role: 'Staff', department: 'Engineering' },
  { employeeId: '44001', password: '1234', name: 'Mint', role: 'Supervisor', department: 'FB' },
  { employeeId: '44005', password: '1234', name: 'Ben', role: 'Staff', department: 'FB' },
  { employeeId: '44006', password: '1234', name: 'Nan', role: 'Staff', department: 'FB' },
  { employeeId: '99001', password: '1234', name: 'Mook', role: 'MOD', department: 'Hotel Ops' }
];

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
    if (log.action === 'Done') doneAt = log.createdAt;
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
    createdAt,
    updatedAt: logs[0]?.createdAt || createdAt,
    assignedAt,
    startedAt,
    doneAt,
    closedAt,
    logs
  });
}

const seedTasks = [
  createSeedTask({
    ticketNo: 'LSH-20260325-0001',
    location: 'A203',
    department: 'Engineering',
    category: 'Repair',
    priority: 'High',
    subject: 'Air conditioner not cold',
    detail: 'Guest reported AC is not cold since morning.',
    status: 'New',
    openedByName: 'Noi',
    openedByDepartment: 'FO',
    assignedToName: '',
    createdAtOffsetMin: 90,
    logs: [
      { action: 'Created', note: 'Task opened by FO after guest call.', byName: 'Noi', byDepartment: 'FO', offsetMin: 90 }
    ]
  }),
  createSeedTask({
    ticketNo: 'LSH-20260325-0002',
    location: 'B112',
    department: 'HK',
    category: 'Guest Request',
    priority: 'Medium',
    subject: 'Extra towels request',
    detail: 'Guest requested two more bath towels.',
    status: 'Accepted',
    openedByName: 'Anna',
    openedByDepartment: 'FO',
    assignedToName: 'May',
    createdAtOffsetMin: 40,
    logs: [
      { action: 'Created', note: 'Opened after guest request.', byName: 'Anna', byDepartment: 'FO', offsetMin: 40 },
      { action: 'Accepted', note: 'Task accepted by HK.', byName: 'May', byDepartment: 'HK', offsetMin: 32 }
    ]
  }),
  createSeedTask({
    ticketNo: 'LSH-20260325-0003',
    location: 'VIP Arrival',
    department: 'FB',
    category: 'Setup',
    priority: 'Urgent',
    subject: 'Welcome amenity setup',
    detail: 'Please prepare VIP welcome fruit and sparkling water.',
    status: 'In Progress',
    openedByName: 'Noi',
    openedByDepartment: 'FO',
    assignedToName: 'Ben',
    createdAtOffsetMin: 22,
    logs: [
      { action: 'Created', note: 'VIP arrival task opened.', byName: 'Noi', byDepartment: 'FO', offsetMin: 22 },
      { action: 'Accepted', note: 'Accepted by FB.', byName: 'Ben', byDepartment: 'FB', offsetMin: 20 },
      { action: 'In Progress', note: 'Amenity setup is being prepared.', byName: 'Ben', byDepartment: 'FB', offsetMin: 16 }
    ]
  }),
  createSeedTask({
    ticketNo: 'LSH-20260325-0004',
    location: 'Lobby',
    department: 'FO',
    category: 'Other',
    priority: 'Low',
    subject: 'Bell desk follow-up',
    detail: 'Guest requested taxi status update.',
    status: 'Done',
    openedByName: 'Anna',
    openedByDepartment: 'FO',
    assignedToName: 'Anna',
    createdAtOffsetMin: 70,
    logs: [
      { action: 'Created', note: 'Taxi follow-up requested.', byName: 'Anna', byDepartment: 'FO', offsetMin: 70 },
      { action: 'Accepted', note: 'Accepted by FO.', byName: 'Anna', byDepartment: 'FO', offsetMin: 68 },
      { action: 'Done', note: 'Taxi confirmed with guest.', byName: 'Anna', byDepartment: 'FO', offsetMin: 64 }
    ]
  }),
  createSeedTask({
    ticketNo: 'LSH-20260324-0010',
    location: 'C305',
    department: 'Engineering',
    category: 'Repair',
    priority: 'Medium',
    subject: 'Bathroom light issue',
    detail: 'Bathroom light flickering on and off.',
    status: 'Closed',
    openedByName: 'Anna',
    openedByDepartment: 'FO',
    assignedToName: 'Art',
    createdAtOffsetMin: 1440,
    logs: [
      { action: 'Created', note: 'Guest reported bathroom light issue.', byName: 'Anna', byDepartment: 'FO', offsetMin: 1440 },
      { action: 'Accepted', note: 'Accepted by Engineering.', byName: 'Art', byDepartment: 'Engineering', offsetMin: 1430 },
      { action: 'In Progress', note: 'Checked ceiling light fixture.', byName: 'Art', byDepartment: 'Engineering', offsetMin: 1410 },
      { action: 'Done', note: 'Bulb replaced and tested.', byName: 'Art', byDepartment: 'Engineering', offsetMin: 1390 },
      { action: 'Closed', note: 'FO confirmed guest satisfied.', byName: 'Noi', byDepartment: 'FO', offsetMin: 1380 }
    ]
  })
];

const seedModReports = [
  normalizeModReport({
    id: 'mod_seed_1',
    reportNo: 'MOD-20260325-0001',
    area: 'Lobby',
    location: 'Lobby entrance',
    department: 'HK',
    category: 'Cleanliness',
    priority: 'High',
    subject: 'Dust on main entrance glass panel',
    detail: 'Dust and fingerprint marks clearly visible from guest arrival side.',
    actionNote: 'Please wipe glass and check brass handle polish before evening peak.',
    status: 'Open',
    openedByName: 'Mook',
    openedByDepartment: 'Hotel Ops',
    createdAt: new Date(Date.now() - 75 * 60000).toISOString(),
    updatedAt: new Date(Date.now() - 75 * 60000).toISOString(),
    attachments: [],
    logs: [{ action: 'Reported', note: 'Opened during AM MOD round.', byName: 'Mook', byDepartment: 'Hotel Ops', createdAt: new Date(Date.now() - 75 * 60000).toISOString() }]
  }),
  normalizeModReport({
    id: 'mod_seed_2',
    reportNo: 'MOD-20260325-0002',
    area: 'Pool / Outdoor',
    location: 'Pool shower',
    department: 'Engineering',
    category: 'Safety',
    priority: 'Urgent',
    subject: 'Loose shower handle at pool area',
    detail: 'Metal shower handle is loose and could come off when guest uses it.',
    actionNote: 'Block use immediately and repair before afternoon opening.',
    status: 'In Progress',
    openedByName: 'Noi',
    openedByDepartment: 'FO',
    createdAt: new Date(Date.now() - 35 * 60000).toISOString(),
    updatedAt: new Date(Date.now() - 20 * 60000).toISOString(),
    attachments: [],
    linkedTaskId: 'LSH-20260325-0001',
    linkedTaskTicketNo: 'LSH-20260325-0001',
    logs: [
      { action: 'Reported', note: 'Opened during MOD pool inspection.', byName: 'Noi', byDepartment: 'FO', createdAt: new Date(Date.now() - 35 * 60000).toISOString() },
      { action: 'Follow-up Started', note: 'Engineering informed and area blocked.', byName: 'Noi', byDepartment: 'FO', createdAt: new Date(Date.now() - 20 * 60000).toISOString() }
    ]
  })
];

const screens = {
  login: document.getElementById('screen-login'),
  app: document.getElementById('screen-app')
};

const pages = {
  home: document.getElementById('page-home'),
  dashboard: document.getElementById('page-dashboard'),
  tasks: document.getElementById('page-tasks'),
  detail: document.getElementById('page-detail'),
  create: document.getElementById('page-create'),
  history: document.getElementById('page-history'),
  mod: document.getElementById('page-mod'),
  reports: document.getElementById('page-reports')
};

const els = {
  loginBtn: document.getElementById('login-btn'),
  loginError: document.getElementById('login-error'),
  loginEmployeeId: document.getElementById('login-employee-id'),
  loginPassword: document.getElementById('login-password'),
  logoutBtn: document.getElementById('logout-btn'),
  backBtn: document.getElementById('back-btn'),
  navHome: document.getElementById('nav-home'),
  navHomeLabel: document.getElementById('nav-home-label'),
  navTasks: document.getElementById('nav-tasks'),
  navCreate: document.getElementById('nav-create'),
  navHistory: document.getElementById('nav-history'),
  navHistoryLabel: document.getElementById('nav-history-label'),
  navMod: document.getElementById('nav-mod'),
  navReports: document.getElementById('nav-reports'),
  bottomNav: document.getElementById('bottom-nav'),
  homeCreateBtn: document.getElementById('home-create-btn'),
  topbarTitle: document.getElementById('topbar-title'),
  topbarSubtitle: document.getElementById('topbar-subtitle'),
  openedByText: document.getElementById('opened-by-text'),
  homeViewTasksBtn: document.getElementById('home-view-tasks-btn'),
  homeSection1Title: document.getElementById('home-section-1-title'),
  homeSection1Hint: document.getElementById('home-section-1-hint'),
  homeSection1Btn: document.getElementById('home-section-1-btn'),
  homeSection1List: document.getElementById('home-section-1-list'),
  homeSection2Title: document.getElementById('home-section-2-title'),
  homeSection2Hint: document.getElementById('home-section-2-hint'),
  homeSection2Btn: document.getElementById('home-section-2-btn'),
  homeSection2List: document.getElementById('home-section-2-list'),
  homeSection3Title: document.getElementById('home-section-3-title'),
  homeSection3Hint: document.getElementById('home-section-3-hint'),
  homeSection3Btn: document.getElementById('home-section-3-btn'),
  homeSection3List: document.getElementById('home-section-3-list'),
  recentActivityTitle: document.getElementById('recent-activity-title'),
  recentActivityHint: document.getElementById('recent-activity-hint'),
  recentActivity: document.getElementById('recent-activity'),
  homeModBoard: document.getElementById('home-mod-board'),
  homeModBoardTitle: document.getElementById('home-mod-board-title'),
  homeModBoardHint: document.getElementById('home-mod-board-hint'),
  homeModBoardBtn: document.getElementById('home-mod-board-btn'),
  homeModBoardList: document.getElementById('home-mod-board-list'),
  statNew: document.getElementById('stat-new'),
  statProgress: document.getElementById('stat-progress'),
  statDone: document.getElementById('stat-done'),
  statUrgent: document.getElementById('stat-urgent'),
  dashOpen: document.getElementById('dash-open'),
  dashProgress: document.getElementById('dash-progress'),
  dashOverdue: document.getElementById('dash-overdue'),
  dashClosedToday: document.getElementById('dash-closed-today'),
  dashboardDepartments: document.getElementById('dashboard-departments'),
  dashboardOverdue: document.getElementById('dashboard-overdue'),
  dashboardActivity: document.getElementById('dashboard-activity'),
  dashGoTasks: document.getElementById('dash-go-tasks'),
  dashGoHistory: document.getElementById('dash-go-history'),
  dashGoReports: document.getElementById('dash-go-reports'),
  createTaskForm: document.getElementById('create-task-form'),
  cancelCreateBtn: document.getElementById('cancel-create-btn'),
  createAssignWrap: document.getElementById('create-assign-wrap'),
  createAssignSelect: document.getElementById('create-assign-select'),
  createAssignHelper: document.getElementById('create-assign-helper'),
  taskDepartment: document.getElementById('task-department'),
  taskPriority: document.getElementById('task-priority'),
  taskStatusTabs: document.getElementById('task-status-tabs'),
  taskFilterHigh: document.getElementById('task-filter-high'),
  tasksSearch: document.getElementById('tasks-search'),
  taskContextBar: document.getElementById('task-context-bar'),
  taskContextLabel: document.getElementById('task-context-label'),
  taskContextClear: document.getElementById('task-context-clear'),
  modActionBoard: document.getElementById('mod-action-board'),
  modActionBoardHint: document.getElementById('mod-action-board-hint'),
  modActionSummary: document.getElementById('mod-action-summary'),
  modActionCount: document.getElementById('mod-action-count'),
  modActionFilters: document.getElementById('mod-action-filters'),
  modActionOpenAll: document.getElementById('mod-action-open-all'),
  modActionList: document.getElementById('mod-action-list'),
  supervisorBoard: document.getElementById('supervisor-board'),
  supervisorBoardSummary: document.getElementById('supervisor-board-summary'),
  teamQuickFilters: document.getElementById('team-quick-filters'),
  teamAssigneeFilter: document.getElementById('team-assignee-filter'),
  teamFilterHint: document.getElementById('team-filter-hint'),
  teamFilterClear: document.getElementById('team-filter-clear'),
  supervisorWorkloadCards: document.getElementById('supervisor-workload-cards'),
  supervisorBalancePanel: document.getElementById('supervisor-balance-panel'),
  supervisorBalanceStatus: document.getElementById('supervisor-balance-status'),
  supervisorBalanceRecommendations: document.getElementById('supervisor-balance-recommendations'),
  supervisorBalanceActions: document.getElementById('supervisor-balance-actions'),
  supervisorBoardLanes: document.getElementById('supervisor-board-lanes'),
  tasksList: document.getElementById('tasks-list'),
  detailTicket: document.getElementById('detail-ticket'),
  detailSubject: document.getElementById('detail-subject'),
  detailLocation: document.getElementById('detail-location'),
  detailPriorityBadge: document.getElementById('detail-priority-badge'),
  detailStatusBadge: document.getElementById('detail-status-badge'),
  detailDescription: document.getElementById('detail-description'),
  detailDepartment: document.getElementById('detail-department'),
  detailCategory: document.getElementById('detail-category'),
  detailOpenedBy: document.getElementById('detail-opened-by'),
  detailAssignedTo: document.getElementById('detail-assigned-to'),
  detailAssignedBy: document.getElementById('detail-assigned-by'),
  detailSource: document.getElementById('detail-source'),
  detailCreatedAt: document.getElementById('detail-created-at'),
  detailUpdatedAt: document.getElementById('detail-updated-at'),
  detailActions: document.getElementById('detail-actions'),
  detailAssignCard: document.getElementById('detail-assign-card'),
  detailAssignHint: document.getElementById('detail-assign-hint'),
  detailAssignSelect: document.getElementById('detail-assign-select'),
  detailAssignNote: document.getElementById('detail-assign-note'),
  detailAssignBtn: document.getElementById('detail-assign-btn'),
  detailAssignFocusBtn: document.getElementById('detail-assign-focus-btn'),
  detailMediaCard: document.getElementById('detail-media-card'),
  detailMediaList: document.getElementById('detail-media-list'),
  detailTimeline: document.getElementById('detail-timeline'),
  detailNoteInput: document.getElementById('detail-note-input'),
  detailSaveNoteBtn: document.getElementById('detail-save-note-btn'),
  historyStartDate: document.getElementById('history-start-date'),
  historyEndDate: document.getElementById('history-end-date'),
  historyPresets: document.getElementById('history-presets'),
  historySearch: document.getElementById('history-search'),
  historyDepartment: document.getElementById('history-department'),
  historyStatus: document.getElementById('history-status'),
  historyPriority: document.getElementById('history-priority'),
  historyResults: document.getElementById('history-results'),
  historyCount: document.getElementById('history-count'),
  reportStartDate: document.getElementById('report-start-date'),
  reportEndDate: document.getElementById('report-end-date'),
  reportPresets: document.getElementById('report-presets'),
  reportViewModes: document.getElementById('report-view-modes'),
  reportViewModeLabel: document.getElementById('report-view-mode-label'),
  reportDepartment: document.getElementById('report-department'),
  reportStatus: document.getElementById('report-status'),
  reportExportCsv: document.getElementById('report-export-csv'),
  reportExportPdf: document.getElementById('report-export-pdf'),
  reportTotal: document.getElementById('report-total'),
  reportClosed: document.getElementById('report-closed'),
  reportOpen: document.getElementById('report-open'),
  reportOverdue: document.getElementById('report-overdue'),
  reportDepartmentSummary: document.getElementById('report-department-summary'),
  reportStatusSummary: document.getElementById('report-status-summary'),
  reportDepartmentChart: document.getElementById('report-department-chart'),
  reportTrendChart: document.getElementById('report-trend-chart'),
  reportPriorityChart: document.getElementById('report-priority-chart'),
  reportPrioritySummary: document.getElementById('report-priority-summary'),
  reportStatusDonut: document.getElementById('report-status-donut'),
  reportStatusLegend: document.getElementById('report-status-legend'),
  reportRangeLabel: document.getElementById('report-range-label'),
  reportResults: document.getElementById('report-results'),
  reportExecSummary: document.getElementById('report-exec-summary'),
  modReportForm: document.getElementById('mod-report-form'),
  modArea: document.getElementById('mod-area'),
  modDepartment: document.getElementById('mod-department'),
  modPriority: document.getElementById('mod-priority'),
  modMediaInput: document.getElementById('mod-media-input'),
  modMediaPreview: document.getElementById('mod-media-preview'),
  modCreateTaskToggle: document.getElementById('mod-create-task-toggle'),
  modSubmitBtn: document.getElementById('mod-submit-btn'),
  modResetBtn: document.getElementById('mod-reset-btn'),
  modFilterChips: document.getElementById('mod-filter-chips'),
  modSearch: document.getElementById('mod-search'),
  modReportsList: document.getElementById('mod-reports-list'),
  modCount: document.getElementById('mod-count'),
  modDetailPanel: document.getElementById('mod-detail-panel'),
  modDetailHint: document.getElementById('mod-detail-hint'),
  modSummaryGrid: document.getElementById('mod-summary-grid')
};

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
  modActionFilter: 'all'
};

document.addEventListener('DOMContentLoaded', init);

function init() {
  initializeTasks();
  initializeModReports();
  bindEvents();
  setupChipGroup(document.getElementById('department-chips'), els.taskDepartment, 'FO', false, renderCreateAssignmentState);
  setupChipGroup(document.getElementById('priority-chips'), els.taskPriority, 'Medium');
  setupChipGroup(document.getElementById('mod-area-chips'), els.modArea, 'Lobby');
  setupChipGroup(document.getElementById('mod-department-chips'), els.modDepartment, 'Engineering');
  setupChipGroup(document.getElementById('mod-priority-chips'), els.modPriority, 'High');
  setDatePreset('history', 'today', true);
  setDatePreset('report', 'today', true);
  restoreSession();
}

function bindEvents() {
  els.loginBtn.addEventListener('click', onLogin);
  els.loginPassword.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') onLogin();
  });
  els.logoutBtn.addEventListener('click', logout);
  els.backBtn.addEventListener('click', onBack);
  els.navHome.addEventListener('click', () => showPage(getHomeNavPage()));
  els.navTasks.addEventListener('click', () => { clearTaskContext(); showPage('tasks'); });
  els.navCreate.addEventListener('click', () => showPage('create'));
  els.navHistory.addEventListener('click', () => showPage('history'));
  els.navMod.addEventListener('click', () => showPage('mod'));
  els.navReports.addEventListener('click', () => showPage('reports'));
  els.homeCreateBtn.addEventListener('click', () => showPage('create'));
  els.homeViewTasksBtn.addEventListener('click', openDepartmentTasksFromHome);
  if (els.homeModBoardBtn) els.homeModBoardBtn.addEventListener('click', openModBoardFromHome);
  if (els.homeModBoardList) els.homeModBoardList.addEventListener('click', onModActionBoardClick);
  els.cancelCreateBtn.addEventListener('click', () => showPage(getDefaultLandingPage()));
  els.createTaskForm.addEventListener('submit', onCreateTask);
  els.taskStatusTabs.addEventListener('click', onTaskTabClick);
  els.taskFilterHigh.addEventListener('click', toggleHighFilter);
  els.taskContextClear.addEventListener('click', () => { clearTaskContext(); renderTaskList(); updateTopbar('tasks'); });
  if (els.modActionFilters) els.modActionFilters.addEventListener('click', onModActionBoardFilterClick);
  if (els.modActionList) els.modActionList.addEventListener('click', onModActionBoardClick);
  if (els.modActionOpenAll) els.modActionOpenAll.addEventListener('click', () => openTaskPreset('modDept'));
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

function initializeTasks() {
  const existing = localStorage.getItem(STORAGE_KEYS.tasks);
  if (!existing) {
    localStorage.setItem(STORAGE_KEYS.tasks, JSON.stringify(seedTasks));
  }
}

function initializeModReports() {
  const existing = localStorage.getItem(STORAGE_KEYS.modReports);
  if (!existing) {
    localStorage.setItem(STORAGE_KEYS.modReports, JSON.stringify(seedModReports));
  }
}

function restoreSession() {
  const savedUser = localStorage.getItem(STORAGE_KEYS.currentUser);
  if (!savedUser) return;
  state.currentUser = JSON.parse(savedUser);
  showApp();
}

function onLogin() {
  const employeeId = els.loginEmployeeId.value.trim();
  const password = els.loginPassword.value.trim();
  const matchedUser = users.find((u) => u.employeeId === employeeId && u.password === password);

  if (!matchedUser) {
    els.loginError.classList.remove('hidden');
    return;
  }

  els.loginError.classList.add('hidden');
  state.currentUser = matchedUser;
  localStorage.setItem(STORAGE_KEYS.currentUser, JSON.stringify(matchedUser));
  showApp();
}

function logout() {
  state.currentUser = null;
  localStorage.removeItem(STORAGE_KEYS.currentUser);
  screens.app.classList.remove('screen--active');
  screens.login.classList.add('screen--active');
  els.loginPassword.value = '';
}

function showApp() {
  screens.login.classList.remove('screen--active');
  screens.app.classList.add('screen--active');
  configureNavigation();
  showPage(getDefaultLandingPage());
  renderApp();
}

function configureNavigation() {
  const manager = isManager();
  const modAccess = canAccessMod();
  const reportsAccess = canViewReports();
  els.navHomeLabel.textContent = manager ? 'Dash' : 'Home';
  els.navHistoryLabel.textContent = 'History';
  els.navMod.classList.toggle('hidden', !modAccess);
  els.navReports.classList.toggle('hidden', !reportsAccess);
  els.bottomNav.classList.remove('bottom-nav--count-4', 'bottom-nav--count-5', 'bottom-nav--count-6');
  const visibleCount = 4 + (modAccess ? 1 : 0) + (reportsAccess ? 1 : 0);
  els.bottomNav.classList.add(`bottom-nav--count-${visibleCount}`);
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
    mod: els.navMod,
    reports: els.navReports
  };
  [els.navHome, els.navTasks, els.navCreate, els.navHistory, els.navMod, els.navReports].forEach((btn) => btn.classList.remove('is-active'));
  activeMap[pageName]?.classList.add('is-active');
}

function updateTopbar(pageName) {
  const titleMap = {
    home: [`Good Morning, ${state.currentUser.name}`, `${state.currentUser.department} / ${state.currentUser.role}`],
    dashboard: ['Manager Dashboard', 'Hotel operations overview'],
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
  renderHomeModActionBoard();
  renderCreateAssignmentState();
  renderSupervisorTaskBoard();
  renderTaskList();
  renderHistoryPage();
  if (canAccessMod()) renderModPage();
  if (isManager()) renderDashboardPage();
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
  if (isManager()) return;

  const visibleTasks = getVisibleTasks(tasks);
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


function renderHomeModActionBoard() {
  if (!els.homeModBoard) return;
  if (!shouldShowModActionBoard()) {
    els.homeModBoard.classList.add('hidden');
    return;
  }

  const reports = getFilteredModActionReports();
  const department = state.currentUser?.department || 'Department';
  els.homeModBoard.classList.remove('hidden');
  if (els.homeModBoardTitle) els.homeModBoardTitle.textContent = `MOD Action Board / ${department}`;
  if (els.homeModBoardHint) {
    els.homeModBoardHint.textContent = isSupervisor()
      ? 'Department findings from MOD round for supervisor follow-up'
      : 'Issues from MOD round for your department responsibility';
  }
  const preview = reports.slice(0, 3);
  els.homeModBoardList.innerHTML = preview.length
    ? preview.map((report) => modActionReportCardHTML(report, { compact: true })).join('')
    : emptyStateHTML('No MOD follow-up now', 'No open MOD finding for your department right now.');
}

function openModBoardFromHome() {
  state.modActionFilter = 'open';
  syncModActionFilterChips();
  openTaskPreset('modDept');
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
    tasksButtonLabel: `View ${department} Tasks`,
    activityTitle: `Latest ${department} Updates`,
    activityHint: 'Recent team activity',
    activityPreset: 'departmentRecent',
    sections: []
  };

  if (isSupervisor()) {
    return {
      ...base,
      tasksButtonLabel: `View ${department} Team Queue`,
      activityTitle: `Latest ${department} Team Updates`,
      activityHint: 'Team workload and supervisor follow-up',
      sections: [
        {
          title: 'New / Unassigned',
          hint: 'Open tasks waiting for supervisor assignment',
          buttonLabel: 'Open Queue',
          preset: 'supervisorNewUnassigned',
          emptyTitle: 'No new unassigned task',
          emptyDescription: 'All fresh requests already have an owner.'
        },
        {
          title: 'Team In Progress',
          hint: 'Accepted and in-progress work for your team',
          buttonLabel: 'Open Team',
          preset: 'supervisorTeamActive',
          emptyTitle: 'No active team task',
          emptyDescription: 'No accepted or in-progress work in the team right now.'
        },
        {
          title: 'Assigned by Me / Overdue',
          hint: 'Track delegation and escalations',
          buttonLabel: 'Open Follow-up',
          preset: 'supervisorAssignedOrOverdue',
          emptyTitle: 'No supervisor follow-up',
          emptyDescription: 'Nothing overdue and no active delegated work from this account.'
        }
      ]
    };
  }

  if (department === 'FO') {
    return {
      ...base,
      tasksButtonLabel: 'View FO Tasks',
      activityTitle: 'Latest FO Updates',
      sections: [
        {
          title: 'Need Follow-up',
          hint: 'Opened by FO or waiting on FO follow-up',
          buttonLabel: 'Open Follow-up',
          preset: 'foFollowUp',
          emptyTitle: 'No follow-up now',
          emptyDescription: 'FO follow-up queue is clear.'
        },
        {
          title: 'Recently Created by Me',
          hint: 'Newest tickets opened by this account',
          buttonLabel: 'Open Mine',
          preset: 'createdByMe',
          emptyTitle: 'No recent tickets',
          emptyDescription: 'This account has not created a task yet.'
        },
        {
          title: 'Urgent / Overdue',
          hint: 'Guest-facing cases needing escalation',
          buttonLabel: 'Open Urgent',
          preset: 'foUrgentOverdue',
          emptyTitle: 'No urgent FO cases',
          emptyDescription: 'Nothing overdue or urgent in FO right now.'
        }
      ]
    };
  }

  if (department === 'HK') {
    return {
      ...base,
      tasksButtonLabel: 'View HK Tasks',
      activityTitle: 'Latest HK Updates',
      sections: [
        {
          title: 'New Tasks for HK',
          hint: 'Fresh requests waiting for housekeeping',
          buttonLabel: 'Open New',
          preset: 'deptNew',
          emptyTitle: 'No new HK tasks',
          emptyDescription: 'No new housekeeping request in queue.'
        },
        {
          title: 'My In Progress',
          hint: 'Tasks currently worked by me',
          buttonLabel: 'Open Mine',
          preset: 'myInProgress',
          emptyTitle: 'Nothing in progress',
          emptyDescription: 'You do not have an active HK task yet.'
        },
        {
          title: 'Urgent / Overdue',
          hint: 'High-pressure tasks needing faster closure',
          buttonLabel: 'Open Urgent',
          preset: 'deptUrgentOverdue',
          emptyTitle: 'No urgent HK tasks',
          emptyDescription: 'HK queue is under control.'
        }
      ]
    };
  }

  if (department === 'Engineering') {
    return {
      ...base,
      tasksButtonLabel: 'View Engineering Tasks',
      activityTitle: 'Latest Engineering Updates',
      sections: [
        {
          title: 'Urgent Repairs',
          hint: 'Critical engineering tickets first',
          buttonLabel: 'Open Repairs',
          preset: 'engUrgentRepairs',
          emptyTitle: 'No urgent repairs',
          emptyDescription: 'No urgent repair ticket at the moment.'
        },
        {
          title: 'My In Progress',
          hint: 'Repairs currently assigned to me',
          buttonLabel: 'Open Mine',
          preset: 'myInProgress',
          emptyTitle: 'No engineering task in progress',
          emptyDescription: 'Your active repair queue is empty.'
        },
        {
          title: 'New Eng / Overdue',
          hint: 'New requests and overdue items',
          buttonLabel: 'Open Queue',
          preset: 'engNewOrOverdue',
          emptyTitle: 'No new or overdue Eng task',
          emptyDescription: 'Engineering queue is balanced right now.'
        }
      ]
    };
  }

  if (department === 'FB') {
    return {
      ...base,
      tasksButtonLabel: 'View FB Tasks',
      activityTitle: 'Latest FB Updates',
      sections: [
        {
          title: 'New FB Requests',
          hint: 'Fresh requests for food and beverage team',
          buttonLabel: 'Open New',
          preset: 'deptNew',
          emptyTitle: 'No new FB requests',
          emptyDescription: 'There is no new FB request waiting now.'
        },
        {
          title: 'VIP Requests',
          hint: 'VIP arrivals and special setup tasks',
          buttonLabel: 'Open VIP',
          preset: 'fbVip',
          emptyTitle: 'No VIP request now',
          emptyDescription: 'No VIP setup or special request at the moment.'
        },
        {
          title: 'My In Progress / Overdue',
          hint: 'Active work and anything needing escalation',
          buttonLabel: 'Open Active',
          preset: 'fbMyInProgressOrOverdue',
          emptyTitle: 'No active FB follow-up',
          emptyDescription: 'Your FB queue is clear for now.'
        }
      ]
    };
  }

  return {
    ...base,
    sections: [
      {
        title: 'Need Attention',
        hint: 'Open tasks for this department',
        buttonLabel: 'Open Queue',
        preset: 'deptOpen'
      },
      {
        title: 'My In Progress',
        hint: 'Active tasks assigned to me',
        buttonLabel: 'Open Mine',
        preset: 'myInProgress'
      },
      {
        title: 'Urgent / Overdue',
        hint: 'Items requiring attention first',
        buttonLabel: 'Open Urgent',
        preset: 'deptUrgentOverdue'
      }
    ]
  };
}

function getHomeTasksByPreset(tasks, preset) {
  const department = state.currentUser?.department;
  const mine = (task) => task.assignedToName === state.currentUser?.name;
  const openedByMe = (task) => task.openedByName === state.currentUser?.name;
  const assignedByMe = (task) => task.assignedByName === state.currentUser?.name;
  const deptTask = (task) => task.department === department;
  const activeTask = (task) => !['Done', 'Closed'].includes(task.status);
  const teamActiveTask = (task) => deptTask(task) && ['Accepted', 'In Progress'].includes(task.status);
  const urgentTask = (task) => ['Urgent', 'High'].includes(task.priority) || isTaskOverdue(task);
  const vipTask = (task) => /vip|amenity|welcome/i.test(`${task.subject} ${task.detail} ${task.location}`);

  const presets = {
    all: tasks,
    deptOnly: tasks.filter((task) => deptTask(task)),
    deptOpen: tasks.filter((task) => deptTask(task) && activeTask(task)),
    modDept: tasks.filter((task) => deptTask(task) && activeTask(task) && isTaskFromMod(task)),
    modMine: tasks.filter((task) => deptTask(task) && activeTask(task) && isTaskFromMod(task) && (mine(task) || !task.assignedToName)),
    modHighRisk: tasks.filter((task) => deptTask(task) && activeTask(task) && isTaskFromMod(task) && urgentTask(task)),
    deptNew: tasks.filter((task) => deptTask(task) && task.status === 'New'),
    myInProgress: tasks.filter((task) => mine(task) && task.status === 'In Progress'),
    createdByMe: tasks.filter((task) => openedByMe(task)),
    supervisorNewUnassigned: tasks.filter((task) => deptTask(task) && task.status === 'New' && !task.assignedToName),
    supervisorTeamActive: tasks.filter((task) => teamActiveTask(task)),
    supervisorAssignedOrOverdue: tasks.filter((task) => deptTask(task) && activeTask(task) && (assignedByMe(task) || urgentTask(task))),
    deptUrgentOverdue: tasks.filter((task) => deptTask(task) && activeTask(task) && urgentTask(task)),
    foFollowUp: tasks.filter((task) => activeTask(task) && (task.openedByDepartment === 'FO' || mine(task) || task.department === 'FO')),
    foUrgentOverdue: tasks.filter((task) => activeTask(task) && urgentTask(task) && (task.openedByDepartment === 'FO' || task.department === 'FO')),
    engUrgentRepairs: tasks.filter((task) => task.department === 'Engineering' && activeTask(task) && (urgentTask(task) || task.category === 'Repair')),
    engNewOrOverdue: tasks.filter((task) => task.department === 'Engineering' && activeTask(task) && (task.status === 'New' || isTaskOverdue(task))),
    fbVip: tasks.filter((task) => task.department === 'FB' && activeTask(task) && vipTask(task)),
    fbMyInProgressOrOverdue: tasks.filter((task) => task.department === 'FB' && activeTask(task) && (mine(task) || isTaskOverdue(task) || task.status === 'In Progress')),
    departmentRecent: tasks.filter((task) => deptTask(task) || task.openedByDepartment === department)
  };

  return presets[preset] || presets.all;
}

function getHomeActivityEntries(tasks, preset) {
  return getHomeTasksByPreset(tasks, preset)
    .flatMap((task) => (task.logs || []).map((log) => ({ ...log, task })));
}

function openDepartmentTasksFromHome() {
  openTaskPreset('deptOnly');
}

function openTaskPreset(preset) {
  resetTaskListControls();
  setTaskContext(preset);
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
    modDept: `${state.currentUser?.department || 'Department'} MOD action board`,
    modMine: 'My MOD follow-up',
    modHighRisk: `${state.currentUser?.department || 'Department'} MOD high risk`,
    deptNew: `${state.currentUser?.department || 'Department'} new tasks`,
    myInProgress: 'My in progress tasks',
    createdByMe: 'Recently created by me',
    supervisorNewUnassigned: `${state.currentUser?.department || 'Department'} new / unassigned`,
    supervisorTeamActive: `${state.currentUser?.department || 'Department'} team active`,
    supervisorAssignedOrOverdue: 'Assigned by me / overdue',
    deptUrgentOverdue: `${state.currentUser?.department || 'Department'} urgent / overdue`,
    foFollowUp: 'FO need follow-up',
    foUrgentOverdue: 'FO urgent / overdue',
    engUrgentRepairs: 'Engineering urgent repairs',
    engNewOrOverdue: 'Engineering new / overdue',
    fbVip: 'FB VIP requests',
    fbMyInProgressOrOverdue: 'FB active / overdue'
  };
  return labels[state.taskContext] || '';
}

function getTaskPageSubtitle() {
  const label = getTaskContextLabel();
  return label ? `Filter and track work orders / ${label}` : 'Filter and track work orders';
}

function renderDashboardPage() {
  if (!isManager()) return;
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
  renderModActionBoard();
  renderSupervisorTaskBoard();
  const tasks = getTaskListResults();
  const contextLabel = getTaskContextLabel();
  els.taskContextBar.classList.toggle('hidden', !contextLabel);
  els.taskContextLabel.textContent = contextLabel || 'Filtered view';
  els.tasksList.innerHTML = tasks.length
    ? tasks.map((task) => taskCardHTML(task)).join('')
    : emptyStateHTML('No matching tasks', 'Try another status, team filter, or search keyword.');
}



function renderModActionBoard() {
  if (!els.modActionBoard) return;
  if (!shouldShowModActionBoard()) {
    els.modActionBoard.classList.add('hidden');
    return;
  }

  const reports = getFilteredModActionReports();
  const allReports = getVisibleModActionReports();
  const department = state.currentUser?.department || 'Department';
  els.modActionBoard.classList.remove('hidden');
  if (els.modActionBoardHint) {
    els.modActionBoardHint.textContent = isSupervisor()
      ? `${department} findings from MOD round for team action`
      : `${department} findings from MOD round in your responsibility`;
  }

  const assignedToMe = allReports.filter((report) => {
    const linkedTask = getLinkedTaskForModReport(report);
    return linkedTask?.assignedToName === state.currentUser?.name && !['Done', 'Closed'].includes(linkedTask.status);
  }).length;

  els.modActionSummary.innerHTML = [
    { label: 'Open Findings', value: allReports.filter((report) => !['Resolved', 'Closed'].includes(report.status)).length },
    { label: 'High Risk', value: allReports.filter((report) => ['High', 'Urgent'].includes(report.priority) && !['Resolved', 'Closed'].includes(report.status)).length },
    { label: 'Linked Tasks', value: allReports.filter((report) => !!getLinkedTaskForModReport(report)).length },
    { label: 'Assigned to Me', value: assignedToMe }
  ].map((card) => `
    <article class="card stat-card mod-action-stat">
      <span class="stat-card__label">${escapeHtml(card.label)}</span>
      <strong class="stat-card__value">${card.value}</strong>
    </article>
  `).join('');

  els.modActionCount.textContent = `${reports.length} item${reports.length === 1 ? '' : 's'}`;
  els.modActionList.innerHTML = reports.length
    ? reports.map((report) => modActionReportCardHTML(report)).join('')
    : emptyStateHTML('No matching MOD follow-up', 'Try another MOD filter or check all department tasks.');
  syncModActionFilterChips();
}

function syncModActionFilterChips() {
  if (!els.modActionFilters) return;
  Array.from(els.modActionFilters.querySelectorAll('[data-mod-board]')).forEach((chip) => {
    chip.classList.toggle('is-active', chip.dataset.modBoard === state.modActionFilter);
  });
}

function getVisibleModActionReports() {
  const department = state.currentUser?.department;
  if (!department || !shouldShowModActionBoard()) return [];
  return getModReports()
    .filter((report) => report.department === department)
    .sort((a, b) => new Date(b.updatedAt || b.createdAt) - new Date(a.updatedAt || a.createdAt));
}

function getFilteredModActionReports() {
  return getVisibleModActionReports().filter(matchesModActionFilter);
}

function matchesModActionFilter(report) {
  const linkedTask = getLinkedTaskForModReport(report);
  switch (state.modActionFilter) {
    case 'open':
      return !['Resolved', 'Closed'].includes(report.status);
    case 'high':
      return ['High', 'Urgent'].includes(report.priority) && !['Resolved', 'Closed'].includes(report.status);
    case 'media':
      return (report.attachments || []).length > 0;
    case 'mine':
      return linkedTask?.assignedToName === state.currentUser?.name;
    default:
      return true;
  }
}

function getLinkedTaskForModReport(report) {
  if (!report?.linkedTaskId && !report?.linkedTaskTicketNo) return null;
  return getTasks().find((task) => (
    (report.linkedTaskId && (task.id === report.linkedTaskId || task.ticketNo === report.linkedTaskId))
    || (report.linkedTaskTicketNo && task.ticketNo === report.linkedTaskTicketNo)
  )) || null;
}

function isTaskFromMod(task) {
  if (!task) return false;
  if (task.sourceType === 'MOD') return true;
  if (String(task.sourceReference || '').startsWith('MOD-')) return true;
  return getModReports().some((report) => (
    report.linkedTaskId && (report.linkedTaskId === task.id || report.linkedTaskId === task.ticketNo || report.linkedTaskTicketNo === task.ticketNo)
  ));
}

function modActionReportCardHTML(report, options = {}) {
  const linkedTask = getLinkedTaskForModReport(report);
  const compact = Boolean(options.compact);
  const assignedLabel = linkedTask?.assignedToName || 'Unassigned';
  const actionNote = report.actionNote || 'No immediate action note.';
  return `
    <article class="card task-card mod-action-card${linkedTask && isTaskOverdue(linkedTask) ? ' is-overdue' : ''}">
      <div class="task-card__top">
        <span class="task-card__ticket">${escapeHtml(report.reportNo)}</span>
        <span class="task-card__time">${timeAgo(report.updatedAt || report.createdAt)}</span>
      </div>
      <div class="task-card__meta">${escapeHtml(report.area)} / ${escapeHtml(report.location)} / ${escapeHtml(report.department)}</div>
      <h3 class="task-card__title">${escapeHtml(report.subject)}</h3>
      <div class="task-card__badges">
        <span class="badge badge-source-mod">MOD</span>
        <span class="badge ${priorityBadgeClass(report.priority)}">${escapeHtml(report.priority)}</span>
        <span class="badge ${modStatusBadgeClass(report.status)}">${escapeHtml(report.status)}</span>
        ${linkedTask ? `<span class="badge ${statusBadgeClass(linkedTask.status)}">${escapeHtml(linkedTask.status)}</span>` : ''}
      </div>
      <div class="task-card__info">
        <span>Owner: ${escapeHtml(assignedLabel)}</span>
        <span>${linkedTask ? `Task ${escapeHtml(linkedTask.ticketNo)}` : 'Task not opened yet'}</span>
      </div>
      <div class="mod-action-card__note">${escapeHtml(actionNote)}</div>
      <div class="task-card__actions">
        ${linkedTask
          ? `<button class="btn btn-primary" type="button" data-task-view="${escapeHtml(linkedTask.id)}">Open Task</button>`
          : `<button class="btn btn-secondary" type="button" data-mod-board-open="modDept">Show MOD Tasks</button>`}
        ${compact
          ? `<button class="btn btn-secondary" type="button" data-mod-board-open="modDept">Open Board</button>`
          : `<button class="btn btn-secondary" type="button" data-mod-board-open="${escapeHtml(linkedTask ? 'modMine' : 'modDept')}">Focus Queue</button>`}
      </div>
    </article>
  `;
}

function onModActionBoardFilterClick(event) {
  const chip = event.target.closest('[data-mod-board]');
  if (!chip) return;
  state.modActionFilter = chip.dataset.modBoard;
  renderModActionBoard();
}

function onModActionBoardClick(event) {
  const taskBtn = event.target.closest('[data-task-view]');
  if (taskBtn) {
    openTaskDetail(taskBtn.dataset.taskView);
    return;
  }
  const boardBtn = event.target.closest('[data-mod-board-open]');
  if (!boardBtn) return;
  openTaskPreset(boardBtn.dataset.modBoardOpen || 'modDept');
}

function shouldShowModActionBoard() {
  return isSupervisor() || isStaff();
}

function renderSupervisorTaskBoard() {
  if (!els.supervisorBoard) return;
  if (!isSupervisor()) {
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
    { label: 'Open Team', value: tasks.filter((task) => !['Done', 'Closed'].includes(task.status)).length, tone: '' },
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
    { title: 'Done', status: 'Done' }
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
  if (!isSupervisor()) return;
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
  const activeTasks = tasks.filter((task) => !['Done', 'Closed'].includes(task.status));
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
      return ['Accepted', 'In Progress'].includes(task.status);
    case 'unassigned':
      return task.status === 'New' && !task.assignedToName;
    case 'assignedByMe':
      return !['Done', 'Closed'].includes(task.status) && task.assignedByName === state.currentUser.name;
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
  return users.filter((user) => user.department === department && user.role !== 'Manager');
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
    els.reportResults.innerHTML = emptyStateHTML('Manager / MOD only', 'This page is available for manager-level roles.');
    if (els.reportDepartmentChart) els.reportDepartmentChart.innerHTML = chartEmptyStateHTML('Manager / MOD only');
    if (els.reportTrendChart) els.reportTrendChart.innerHTML = chartEmptyStateHTML('Manager / MOD only');
    if (els.reportPriorityChart) els.reportPriorityChart.innerHTML = chartEmptyStateHTML('Manager / MOD only');
    if (els.reportPrioritySummary) els.reportPrioritySummary.innerHTML = '';
    if (els.reportStatusDonut) els.reportStatusDonut.innerHTML = chartEmptyStateHTML('Manager / MOD only');
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
      active: priorityTasks.filter((t) => !['Closed', 'Done'].includes(t.status)).length,
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
  return (JSON.parse(localStorage.getItem(STORAGE_KEYS.modReports) || '[]')).map(normalizeModReport);
}

function saveModReports(reports) {
  localStorage.setItem(STORAGE_KEYS.modReports, JSON.stringify(reports));
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
    inlineSaved: false
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
          ? (item.dataUrl
              ? `<video src="${escapeHtml(item.dataUrl)}" controls muted playsinline></video>`
              : `<div class="mod-media-placeholder">Video file<br>${escapeHtml(item.name)}</div>`)
          : (item.dataUrl
              ? `<img src="${escapeHtml(item.dataUrl)}" alt="${escapeHtml(item.name)}" />`
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
    alert('This page is available for MOD / Manager role.');
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
  const attachments = (state.modDraftMedia || []).map((item) => ({ ...item }));

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
    els.modReportsList.innerHTML = emptyStateHTML('MOD access only', 'This page is available for MOD / Manager role.');
    els.modDetailPanel.innerHTML = emptyStateHTML('No access', 'Please login with MOD or manager-level account.');
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
  const canAssignNow = selectedAssigneeId && canAssignToDepartment(department);
  const selectedAssignee = canAssignNow ? users.find((user) => user.employeeId === selectedAssigneeId) : null;

  if (!location || !department || !priority || !subject) {
    alert('Please fill required fields: location, department, priority, and subject.');
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
  setupChipGroup(document.getElementById('department-chips'), els.taskDepartment, 'FO', true, renderCreateAssignmentState);
  setupChipGroup(document.getElementById('priority-chips'), els.taskPriority, 'Medium', true);
  renderApp();
  alert(selectedAssignee
    ? `Task created and assigned to ${selectedAssignee.name}: ${newTask.ticketNo}`
    : `Task created successfully: ${newTask.ticketNo}`);
  showPage('tasks');
}

function renderCreateAssignmentState() {
  if (!els.createAssignWrap || !state.currentUser) return;
  const selectedDepartment = els.taskDepartment?.value || 'FO';
  const canAssign = canAssignToDepartment(selectedDepartment);
  els.createAssignWrap.classList.toggle('hidden', !canAssign);
  if (!canAssign) {
    populateAssigneeSelect(els.createAssignSelect, [], '', 'Assignment available for supervisor own department or manager only');
    return;
  }

  const assignees = getAssignableUsers(selectedDepartment);
  populateAssigneeSelect(els.createAssignSelect, assignees, '', `Optional: assign immediately to ${selectedDepartment} team`);
  if (els.createAssignHelper) {
    els.createAssignHelper.textContent = canSeeAllHotelTasks()
      ? `${state.currentUser.role} can assign tasks to any department during creation.`
      : `Supervisor can assign only to ${state.currentUser.department} team during creation.`;
  }
}

function renderDetailAssignmentState(task) {
  if (!els.detailAssignCard) return;
  const canAssign = !!task && canManageAssignments(task) && !['Done', 'Closed'].includes(task.status);
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
    if (draft.status === 'New') draft.status = 'Accepted';
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
    .filter((user) => user.department === department && user.role !== 'Manager')
    .sort((a, b) => {
      const roleWeight = (role) => ({ Supervisor: 0, Staff: 1 }[role] ?? 2);
      return roleWeight(a.role) - roleWeight(b.role) || a.name.localeCompare(b.name);
    });
}

function canAssignToDepartment(department) {
  return canSeeAllHotelTasks() || (isSupervisor() && state.currentUser?.department === department);
}

function canManageAssignments(task) {
  return canSeeAllHotelTasks() || (isSupervisor() && state.currentUser?.department === task.department);
}

function canWorkOnTask(task) {
  return canSeeAllHotelTasks() || canManageAssignments(task) || task.assignedToName === state.currentUser?.name;
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
    accept: { status: 'Accepted', note: 'Task accepted.' },
    start: { status: 'In Progress', note: 'Work started.' },
    done: { status: 'Done', note: 'Work marked done.' },
    close: { status: 'Closed', note: 'Task closed.' },
    reopen: { status: 'In Progress', note: 'Task reopened.' }
  };
  if (!transitions[action]) return;
  if (!canTransitionTask(task, action)) {
    alert('This action is not allowed for the current user or task status.');
    return;
  }
  updateTask(taskId, (draft) => {
    const now = new Date().toISOString();
    draft.status = transitions[action].status;
    draft.updatedAt = now;
    if (action === 'accept') {
      draft.assignedToName = state.currentUser.name;
      draft.assignedAt = now;
    }
    if (action === 'start') draft.startedAt = now;
    if (action === 'done') draft.doneAt = now;
    if (action === 'close') draft.closedAt = now;
    if (action === 'reopen') {
      draft.closedAt = '';
      draft.doneAt = '';
    }
    draft.logs.unshift(createLog(capitalize(action), transitions[action].note));
    return draft;
  });
  renderApp();
  if (state.currentView === 'detail') showPage('detail');
}

function openTaskDetail(taskId) {
  state.currentTaskId = taskId;
  showPage('detail');
}

function getDetailActions(task) {
  const actions = [];
  const manager = canSeeAllHotelTasks();
  const sameDepartment = state.currentUser.department === task.department;
  const canManage = canManageAssignments(task);
  const canWork = canWorkOnTask(task);

  if (task.status === 'New' && sameDepartment && !canManage) {
    actions.push({ value: 'accept', label: 'Accept Task', primary: true, full: true });
  }
  if (task.status === 'Accepted' && canWork) {
    actions.push({ value: 'start', label: 'Start Work', primary: true, full: true });
  }
  if (task.status === 'In Progress' && canWork) {
    actions.push({ value: 'focus-note', label: 'Add Note', primary: false });
    actions.push({ value: 'done', label: 'Mark Done', primary: true });
  }
  if (task.status === 'Done' && (manager || state.currentUser.department === task.openedByDepartment)) {
    actions.push({ value: 'close', label: 'Close Task', primary: true, full: true });
  }
  if (task.status === 'Closed' && (manager || canManage)) {
    actions.push({ value: 'reopen', label: 'Reopen Task', primary: false, full: true });
  }
  return actions;
}

function canTransitionTask(task, action) {
  return getDetailActions(task).some((item) => item.value === action);
}

function getTaskListResults() {
  if (isSupervisor()) return getSupervisorTaskScope();
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
      active: priorityTasks.filter((t) => !['Closed', 'Done'].includes(t.status)).length,
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
  const urgentOpen = tasks.filter((task) => ['High', 'Urgent'].includes(task.priority) && !['Done', 'Closed'].includes(task.status)).length;
  const oldestOpen = [...tasks]
    .filter((task) => !['Done', 'Closed'].includes(task.status))
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
  if (canSeeAllHotelTasks()) return tasks;
  return tasks.filter((task) => task.department === state.currentUser.department || task.openedByDepartment === state.currentUser.department);
}

function getTasks() {
  return (JSON.parse(localStorage.getItem(STORAGE_KEYS.tasks) || '[]')).map(normalizeTask);
}

function saveTasks(tasks) {
  localStorage.setItem(STORAGE_KEYS.tasks, JSON.stringify(tasks));
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

function normalizeTask(task) {
  return {
    id: task.id || task.ticketNo || generateId(),
    ticketNo: task.ticketNo,
    location: task.location || '-',
    department: task.department || 'FO',
    category: task.category || 'Other',
    priority: task.priority || 'Medium',
    subject: task.subject || '-',
    detail: task.detail || '',
    status: task.status || 'New',
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
  return `
    <article class="card task-card${overdueClass}">
      <div class="task-card__top">
        <span class="task-card__ticket">${escapeHtml(task.ticketNo)}</span>
        <span class="task-card__time">${formatClock(task.updatedAt || task.createdAt)}</span>
      </div>
      <div class="task-card__meta">${escapeHtml(task.location)} / ${escapeHtml(task.department)}</div>
      <h3 class="task-card__title">${escapeHtml(task.subject)}</h3>
      <div class="task-card__badges">
        ${isTaskFromMod(task) ? '<span class="badge badge-source-mod">MOD</span>' : ''}
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
    Done: 'badge-status-done',
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
  if (task.status === 'Closed' || task.status === 'Done') return false;
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

function isManager() {
  return state.currentUser?.role === 'Manager';
}

function isMOD() {
  return state.currentUser?.role === 'MOD';
}

function isSupervisor() {
  return state.currentUser?.role === 'Supervisor';
}

function isStaff() {
  return state.currentUser?.role === 'Staff';
}

function canAccessMod() {
  return isManager() || isMOD();
}

function canViewReports() {
  return isManager() || isMOD();
}

function canSeeAllHotelTasks() {
  return isManager() || isMOD();
}

function getDefaultLandingPage() {
  if (isManager()) return 'dashboard';
  if (isMOD()) return 'mod';
  return 'home';
}

function getHomeNavPage() {
  if (isManager()) return 'dashboard';
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






