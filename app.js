const STORAGE_KEYS = {
  currentUser: 'lsh_current_user',
  tasks: 'lsh_tasks'
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
  { employeeId: '12001', password: '1234', name: 'Anna', role: 'Staff', department: 'FO' },
  { employeeId: '22018', password: '1234', name: 'May', role: 'Staff', department: 'HK' },
  { employeeId: '33007', password: '1234', name: 'Art', role: 'Staff', department: 'Engineering' },
  { employeeId: '44005', password: '1234', name: 'Ben', role: 'Staff', department: 'FB' }
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
  navReports: document.getElementById('nav-reports'),
  bottomNav: document.getElementById('bottom-nav'),
  homeCreateBtn: document.getElementById('home-create-btn'),
  homeViewTasksBtn: document.getElementById('home-view-tasks-btn'),
  topbarTitle: document.getElementById('topbar-title'),
  topbarSubtitle: document.getElementById('topbar-subtitle'),
  openedByText: document.getElementById('opened-by-text'),
  homeTaskList: document.getElementById('home-task-list'),
  recentActivity: document.getElementById('recent-activity'),
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
  taskDepartment: document.getElementById('task-department'),
  taskPriority: document.getElementById('task-priority'),
  taskStatusTabs: document.getElementById('task-status-tabs'),
  taskFilterHigh: document.getElementById('task-filter-high'),
  tasksSearch: document.getElementById('tasks-search'),
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
  detailCreatedAt: document.getElementById('detail-created-at'),
  detailUpdatedAt: document.getElementById('detail-updated-at'),
  detailActions: document.getElementById('detail-actions'),
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
  reportExecSummary: document.getElementById('report-exec-summary')
};

const state = {
  currentUser: null,
  currentView: 'home',
  previousView: 'home',
  currentTaskId: null,
  taskStatusFilter: 'All',
  taskSearch: '',
  highOnly: false,
  historyPreset: 'today',
  reportPreset: 'today',
  reportViewMode: 'daily'
};

document.addEventListener('DOMContentLoaded', init);

function init() {
  initializeTasks();
  bindEvents();
  setupChipGroup(document.getElementById('department-chips'), els.taskDepartment, 'FO');
  setupChipGroup(document.getElementById('priority-chips'), els.taskPriority, 'Medium');
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
  els.navHome.addEventListener('click', () => showPage(isManager() ? 'dashboard' : 'home'));
  els.navTasks.addEventListener('click', () => showPage('tasks'));
  els.navCreate.addEventListener('click', () => showPage('create'));
  els.navHistory.addEventListener('click', () => showPage('history'));
  els.navReports.addEventListener('click', () => showPage('reports'));
  els.homeCreateBtn.addEventListener('click', () => showPage('create'));
  els.homeViewTasksBtn.addEventListener('click', () => showPage('tasks'));
  els.cancelCreateBtn.addEventListener('click', () => showPage(isManager() ? 'dashboard' : 'home'));
  els.createTaskForm.addEventListener('submit', onCreateTask);
  els.taskStatusTabs.addEventListener('click', onTaskTabClick);
  els.taskFilterHigh.addEventListener('click', toggleHighFilter);
  els.tasksSearch.addEventListener('input', onTasksSearch);
  els.tasksList.addEventListener('click', onTaskCardClick);
  els.homeTaskList.addEventListener('click', onTaskCardClick);
  els.dashboardOverdue.addEventListener('click', onTaskCardClick);
  els.historyResults.addEventListener('click', onTaskCardClick);
  els.reportResults.addEventListener('click', onTaskCardClick);
  els.detailActions.addEventListener('click', onDetailActionClick);
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
  els.dashGoTasks.addEventListener('click', () => showPage('tasks'));
  els.dashGoHistory.addEventListener('click', () => showPage('history'));
  els.dashGoReports.addEventListener('click', () => showPage('reports'));
}

function initializeTasks() {
  const existing = localStorage.getItem(STORAGE_KEYS.tasks);
  if (!existing) {
    localStorage.setItem(STORAGE_KEYS.tasks, JSON.stringify(seedTasks));
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
  showPage(isManager() ? 'dashboard' : 'home');
  renderApp();
}

function configureNavigation() {
  const manager = isManager();
  els.navHomeLabel.textContent = manager ? 'Dash' : 'Home';
  els.navHistoryLabel.textContent = 'History';
  els.navReports.classList.toggle('hidden', !manager);
  els.bottomNav.classList.toggle('bottom-nav--count-5', manager);
  els.bottomNav.classList.toggle('bottom-nav--count-4', !manager);
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
  els.backBtn.classList.toggle('hidden', pageName === 'home' || pageName === 'dashboard');

  if (pageName === 'create') {
    els.openedByText.textContent = `Opened by: ${state.currentUser.name} / ${state.currentUser.department} / ${formatDateTime(new Date().toISOString())}`;
  }

  if (pageName === 'tasks') renderTaskList();
  if (pageName === 'detail') renderTaskDetail();
  if (pageName === 'history') renderHistoryPage();
  if (pageName === 'dashboard') renderDashboardPage();
  if (pageName === 'reports') renderReportsPage();
}

function onBack() {
  if (state.currentView === 'detail') {
    showPage(state.previousView === 'history' ? 'history' : state.previousView === 'reports' ? 'reports' : 'tasks');
    return;
  }
  showPage(isManager() ? 'dashboard' : 'home');
}

function setActiveNav(pageName) {
  const activeMap = {
    home: els.navHome,
    dashboard: els.navHome,
    tasks: els.navTasks,
    detail: els.navTasks,
    create: els.navCreate,
    history: els.navHistory,
    reports: els.navReports
  };
  [els.navHome, els.navTasks, els.navCreate, els.navHistory, els.navReports].forEach((btn) => btn.classList.remove('is-active'));
  activeMap[pageName]?.classList.add('is-active');
}

function updateTopbar(pageName) {
  const titleMap = {
    home: [`Good Morning, ${state.currentUser.name}`, `${state.currentUser.department} • ${state.currentUser.role}`],
    dashboard: ['Manager Dashboard', 'Hotel operations overview'],
    create: ['Create Task', 'Open a new request'],
    tasks: ['Tasks', 'Filter and track work orders'],
    detail: ['Task Detail', 'View status, notes, and action'],
    history: ['History', 'Search previous tasks'],
    reports: ['Reports', 'Summary and export']
  };
  const [title, subtitle] = titleMap[pageName] || ['Laya Service Hub', ''];
  els.topbarTitle.textContent = title;
  els.topbarSubtitle.textContent = subtitle;
}

function renderApp() {
  const tasks = getTasks();
  renderSummary(tasks);
  renderHomeTasks(tasks);
  renderRecentActivity(tasks);
  renderTaskList();
  renderHistoryPage();
  if (isManager()) {
    renderDashboardPage();
    renderReportsPage();
  }
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

function renderHomeTasks(tasks) {
  const visibleTasks = getVisibleTasks(tasks)
    .filter((task) => task.status !== 'Closed')
    .sort(sortTasks)
    .slice(0, 5);

  if (!visibleTasks.length) {
    els.homeTaskList.innerHTML = emptyStateHTML('No tasks found', 'No visible tasks for this user yet.');
    return;
  }

  els.homeTaskList.innerHTML = visibleTasks.map((task) => taskCardHTML(task)).join('');
}

function renderRecentActivity(tasks) {
  const allLogs = getVisibleTasks(tasks)
    .flatMap((task) => (task.logs || []).map((log) => ({ ...log, task })))
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))
    .slice(0, 5);

  if (!allLogs.length) {
    els.recentActivity.innerHTML = '<div class="activity-row"><div class="activity-row__text">No recent activity</div></div>';
    return;
  }

  els.recentActivity.innerHTML = allLogs.map((entry) => `
    <div class="activity-row">
      <div class="activity-row__text">
        <strong>${escapeHtml(entry.task.ticketNo)}</strong> · ${escapeHtml(entry.action)} · ${escapeHtml(entry.note || entry.task.subject)}
      </div>
      <div class="activity-row__time">${timeAgo(entry.createdAt)}</div>
    </div>
  `).join('');
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
        <span class="overview-row__value">Open ${open} · Done ${doneToday}</span>
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
          <div class="activity-row__text"><strong>${escapeHtml(entry.task.ticketNo)}</strong> · ${escapeHtml(entry.action)} · ${escapeHtml(entry.note || '')}</div>
          <div class="activity-row__time">${timeAgo(entry.createdAt)}</div>
        </div>
      `).join('')
    : '<div class="activity-row"><div class="activity-row__text">No recent activity</div></div>';
}

function renderTaskList() {
  const tasks = getTaskListResults();
  els.tasksList.innerHTML = tasks.length
    ? tasks.map((task) => taskCardHTML(task)).join('')
    : emptyStateHTML('No matching tasks', 'Try another status or search keyword.');
}

function renderTaskDetail() {
  const task = getTaskById(state.currentTaskId);
  if (!task) {
    els.detailTimeline.innerHTML = emptyStateHTML('Task not found', 'This task may have been removed.');
    return;
  }

  els.detailLocation.textContent = `${task.location} • ${task.department}`;
  els.detailSubject.textContent = task.subject;
  els.detailTicket.textContent = task.ticketNo;
  els.detailDescription.textContent = task.detail || 'No detail provided.';
  els.detailDepartment.textContent = task.department;
  els.detailCategory.textContent = task.category;
  els.detailOpenedBy.textContent = `${task.openedByName} / ${task.openedByDepartment}`;
  els.detailAssignedTo.textContent = task.assignedToName || '-';
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
  if (!isManager()) {
    els.reportResults.innerHTML = emptyStateHTML('Manager only', 'This page is available for manager role.');
    if (els.reportDepartmentChart) els.reportDepartmentChart.innerHTML = chartEmptyStateHTML('Manager only');
    if (els.reportTrendChart) els.reportTrendChart.innerHTML = chartEmptyStateHTML('Manager only');
    if (els.reportPriorityChart) els.reportPriorityChart.innerHTML = chartEmptyStateHTML('Manager only');
    if (els.reportPrioritySummary) els.reportPrioritySummary.innerHTML = '';
    if (els.reportStatusDonut) els.reportStatusDonut.innerHTML = chartEmptyStateHTML('Manager only');
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
    els.reportViewModeLabel.textContent = `${getReportViewModeLabel()} view • created tasks`;
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
    meta: `${item.closed} closed · ${item.overdue} overdue`,
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
      meta: `${item.active} active · ${item.share}% of selected workload`,
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

function getTrendStats(tasks, viewMode) {
  const grouped = new Map();
  tasks.forEach((task) => {
    const date = new Date(task.createdAt || task.updatedAt);
    let key;
    let label;
    let sortValue;

    if (viewMode === 'weekly') {
      const start = getWeekStart(date);
      key = formatDateInput(start);
      label = `${String(start.getDate()).padStart(2, '0')} ${start.toLocaleDateString('en-GB', { month: 'short' })}`;
      sortValue = start.getTime();
    } else if (viewMode === 'monthly') {
      const start = new Date(date.getFullYear(), date.getMonth(), 1);
      key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
      label = start.toLocaleDateString('en-GB', { month: 'short', year: '2-digit' });
      sortValue = start.getTime();
    } else {
      key = formatDateInput(date);
      label = date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
      sortValue = new Date(`${key}T00:00:00`).getTime();
    }

    if (!grouped.has(key)) grouped.set(key, { key, label, sortValue, count: 0 });
    grouped.get(key).count += 1;
  });

  const maxBuckets = { daily: 8, weekly: 10, monthly: 12 }[viewMode] || 8;
  return Array.from(grouped.values()).sort((a, b) => a.sortValue - b.sortValue).slice(-maxBuckets);
}

function renderTrendChart(stats) {
  const width = 520;
  const height = 250;
  const chartX = 28;
  const chartY = 26;
  const chartW = width - 56;
  const chartH = 166;
  const maxValue = Math.max(...stats.map((item) => item.count), 1);
  const gap = 12;
  const barWidth = Math.max(24, Math.min(56, (chartW - (gap * Math.max(stats.length - 1, 0))) / Math.max(stats.length, 1)));
  const totalBarsWidth = (stats.length * barWidth) + (Math.max(stats.length - 1, 0) * gap);
  const startX = chartX + ((chartW - totalBarsWidth) / 2);

  const bars = stats.map((item, index) => {
    const x = startX + index * (barWidth + gap);
    const barHeight = item.count ? Math.max(14, (item.count / maxValue) * (chartH - 22)) : 0;
    const y = chartY + chartH - barHeight;
    return `
      <g>
        <rect x="${x}" y="${chartY + 10}" width="${barWidth}" height="${chartH - 10}" rx="14" fill="#f3f5f3"></rect>
        <rect x="${x}" y="${y}" width="${barWidth}" height="${barHeight}" rx="14" fill="#1f4d3a"></rect>
        <text x="${x + barWidth / 2}" y="${y - 8}" text-anchor="middle" class="bar-value">${item.count}</text>
        <text x="${x + barWidth / 2}" y="${chartY + chartH + 20}" text-anchor="middle" class="axis-label">${escapeHtml(item.label)}</text>
      </g>`;
  }).join('');

  return `
    <svg class="report-chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeHtml(getReportViewModeLabel())} workload trend chart">
      <text x="28" y="16" font-size="12" font-weight="700" fill="#7c867f">${escapeHtml(getReportViewModeLabel())} created volume</text>
      ${bars}
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
        <text x="${labelX}" y="${y + 35}" font-size="11" font-weight="600" fill="#7c867f">${item.active} active • ${item.share}%</text>
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
        <text x="${countX}" y="${y + 35}" text-anchor="end" font-size="11" font-weight="600" fill="#7c867f">${item.closed} closed · ${item.overdue} overdue</text>
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

  if (!location || !department || !priority || !subject) {
    alert('Please fill required fields: location, department, priority, and subject.');
    return;
  }

  const tasks = getTasks();
  const now = new Date().toISOString();
  const newTask = normalizeTask({
    id: generateId(),
    ticketNo: generateTicketNo(tasks),
    location,
    department,
    category,
    priority,
    subject,
    detail,
    status: 'New',
    openedByName: state.currentUser.name,
    openedByDepartment: state.currentUser.department,
    assignedToName: '',
    createdAt: now,
    updatedAt: now,
    logs: [
      {
        action: 'Created',
        note: subject,
        byName: state.currentUser.name,
        byDepartment: state.currentUser.department,
        createdAt: now
      }
    ]
  });

  tasks.unshift(newTask);
  saveTasks(tasks);
  els.createTaskForm.reset();
  setupChipGroup(document.getElementById('department-chips'), els.taskDepartment, 'FO', true);
  setupChipGroup(document.getElementById('priority-chips'), els.taskPriority, 'Medium', true);
  renderApp();
  alert(`Task created successfully: ${newTask.ticketNo}`);
  showPage('tasks');
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
    if (action === 'reopen') draft.closedAt = '';

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
  const manager = isManager();
  const sameDepartment = state.currentUser.department === task.department;

  if (task.status === 'New' && (sameDepartment || manager)) {
    actions.push({ value: 'accept', label: 'Accept Task', primary: true, full: true });
  }
  if (task.status === 'Accepted' && (sameDepartment || manager)) {
    actions.push({ value: 'start', label: 'Start Work', primary: true, full: true });
  }
  if (task.status === 'In Progress' && (sameDepartment || manager)) {
    actions.push({ value: 'focus-note', label: 'Add Note', primary: false });
    actions.push({ value: 'done', label: 'Mark Done', primary: true });
  }
  if (task.status === 'Done' && (manager || state.currentUser.department === task.openedByDepartment)) {
    actions.push({ value: 'close', label: 'Close Task', primary: true, full: true });
  }
  if (task.status === 'Closed' && manager) {
    actions.push({ value: 'reopen', label: 'Reopen Task', primary: false, full: true });
  }
  return actions;
}

function getTaskListResults() {
  return getVisibleTasks(getTasks())
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

  doc.setFillColor(24, 55, 43);
  doc.rect(0, 0, pageWidth, 92, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(22);
  doc.text('Laya Service Hub Performance Report', margin, 44);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  doc.text(`Excom-ready export • ${buildReportRangeLabel()}`, margin, 66);
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
    const wrapped = doc.splitTextToSize(`• ${line}`, pageWidth - (margin * 2));
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
  cursorY = chartTop + chartHeight + 24;

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
      doc.text(`Laya Service Hub • Page ${pageNumber}`, margin, pageHeight - 18);
    }
  });

  doc.save(`laya-service-hub-report-${formatFileDate(new Date())}.pdf`);
}

function buildReportRangeLabel() {
  let range = 'All dates';
  if (els.reportStartDate.value && els.reportEndDate.value) range = `${els.reportStartDate.value} to ${els.reportEndDate.value}`;
  else if (els.reportStartDate.value || els.reportEndDate.value) range = els.reportStartDate.value || els.reportEndDate.value;
  return `${range} • ${getReportViewModeLabel()} view`;
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
  doc.text('Bar chart • workload by team', x + 14, y + 32);

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
  doc.text('Donut chart • current task mix', x + 14, y + 32);

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
    doc.text(`${item.count} • ${item.share}%`, x + width - 16, legendY, { align: 'right' });
    legendY += 21;
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
  if (isManager()) return tasks;
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
      <div class="task-card__meta">${escapeHtml(task.location)} • ${escapeHtml(task.department)}</div>
      <h3 class="task-card__title">${escapeHtml(task.subject)}</h3>
      <div class="task-card__badges">
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
        <div class="timeline-item__title">${escapeHtml(log.action)} · ${escapeHtml(log.byName || '-')}</div>
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

function setupChipGroup(container, hiddenInput, defaultValue, forceReset = false) {
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
  return date.toISOString().slice(0, 10);
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
  const taskDate = new Date(task.updatedAt || task.createdAt);
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

function isManager() {
  return state.currentUser?.role === 'Manager';
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
