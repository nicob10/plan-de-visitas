const SERVICE_TYPES = [
  { key: "fixedFire", label: "Instalaciones Fijas", short: "IF", css: "sf" },
  { key: "extinguishers", label: "Extintores", short: "EX", css: "se" },
  { key: "works", label: "Obras C.I.", short: "OB", css: "so" }
];

const MEETING_COLOR_OPTIONS = [
  { value: "yellow", label: "Amarillo" },
  { value: "orange", label: "Naranja" },
  { value: "blue", label: "Azul" },
  { value: "purple", label: "Violeta" },
  { value: "green", label: "Verde" },
  { value: "red", label: "Rojo" }
];

const DEFAULT_OPPORTUNITY_STATUSES = ["Abierta", "Calificada", "Propuesta", "Negociación", "Ganada", "Perdida"];
const DEFAULT_OPPORTUNITY_TYPES = ["Proyecto", "Negociación"];
const DEFAULT_SERVICE_LINES = ["Instalaciones Fijas", "Extintores", "Obras C.I.", "Multiservicio", "Otro"];
const DEFAULT_FOLLOW_UP_TYPES = ["Llamada", "Email", "Visita", "Cotización", "Recordatorio", "Reunión"];
const DEFAULT_FOLLOW_UP_STATUSES = ["Pendiente", "Hecho"];
const VISIT_RULE_OBJECTIVE_OPTIONS = ["Desarrollo de cuentas", "Nuevo negocio"];
const BITRIX_COMPLAINT_RESPONSIBLES = [
  { id: "44", name: "Gonzalo Garcia" },
  { id: "10542", name: "Guillermo Saralegui" },
  { id: "1720", name: "Guido Avalo Ceci" }
];
const CRM_ENABLED = false;

const loginScreen = document.getElementById("loginScreen");
const appContent = document.getElementById("appContent");
const loginForm = document.getElementById("loginForm");
const loginEmail = document.getElementById("loginEmail");
const loginPassword = document.getElementById("loginPassword");
const loginRemember = document.getElementById("loginRemember");
const loginSubmit = document.getElementById("loginSubmit");
const loginStatus = document.getElementById("loginStatus");
const currentUserBadge = document.getElementById("currentUserBadge");
const logoutBtn = document.getElementById("logoutBtn");
const showClientsBtn = document.getElementById("showClientsBtn");
const showVisitsBtn = document.getElementById("showVisitsBtn");
const showPipelineBtn = document.getElementById("showPipelineBtn");
const crmMenu = document.getElementById("crmMenu");
const showVisitsGridBtn = document.getElementById("showVisitsGridBtn");
const showCalendarBtn = document.getElementById("showCalendarBtn");
const showVisitsStatsBtn = document.getElementById("showVisitsStatsBtn");
const showUsersBtn = document.getElementById("showUsersBtn");
const showSettingsBtn = document.getElementById("showSettingsBtn");
const settingsTabButtons = Array.from(document.querySelectorAll(".settings-tab-btn"));
const settingsPaneCatalogs = document.getElementById("settingsPaneCatalogs");
const settingsPaneImports = document.getElementById("settingsPaneImports");
const settingsPaneControl = document.getElementById("settingsPaneControl");
const openPipelineFromCrmBtn = document.getElementById("openPipelineFromCrmBtn");
const openUsersFromSettingsBtn = document.getElementById("openUsersFromSettingsBtn");
const backToSettingsBtn = document.getElementById("backToSettingsBtn");

const tableBody = document.getElementById("companyTableBody");
const searchInput = document.getElementById("searchInput");
const riskFilter = document.getElementById("riskFilter");
const segmentFilter = document.getElementById("segmentFilter");
const fixedFireFilter = document.getElementById("fixedFireFilter");
const extinguishersFilter = document.getElementById("extinguishersFilter");
const worksFilter = document.getElementById("worksFilter");
const visibleCount = document.getElementById("visibleCount");
const globalKpis = document.getElementById("globalKpis");
const addCompanyBtn = document.getElementById("addCompanyBtn");
const usersTableBody = document.getElementById("usersTableBody");
const userCreateForm = document.getElementById("userCreateForm");
const newUserName = document.getElementById("newUserName");
const newUserEmail = document.getElementById("newUserEmail");
const newUserRole = document.getElementById("newUserRole");
const newUserBitrixSearch = document.getElementById("newUserBitrixSearch");
const newUserBitrixOptions = document.getElementById("newUserBitrixOptions");
const newUserBitrixId = document.getElementById("newUserBitrixId");
const newUserPassword = document.getElementById("newUserPassword");
const permCreateClients = document.getElementById("permCreateClients");
const permEditClients = document.getElementById("permEditClients");
const permHideClients = document.getElementById("permHideClients");
const permCreateBranches = document.getElementById("permCreateBranches");
const permEditBranches = document.getElementById("permEditBranches");
const permHideBranches = document.getElementById("permHideBranches");
const createUserBtn = document.getElementById("createUserBtn");
const createUserStatus = document.getElementById("createUserStatus");
const userFormTitle = document.getElementById("userFormTitle");
const cancelUserEditBtn = document.getElementById("cancelUserEditBtn");
const sectorForm = document.getElementById("sectorForm");
const newSectorName = document.getElementById("newSectorName");
const createSectorBtn = document.getElementById("createSectorBtn");
const sectorStatus = document.getElementById("sectorStatus");
const sectorsTableBody = document.getElementById("sectorsTableBody");
const downloadClientsTemplateBtn = document.getElementById("downloadClientsTemplateBtn");
const clientsImportFile = document.getElementById("clientsImportFile");
const importClientsBtn = document.getElementById("importClientsBtn");
const clientsImportStatus = document.getElementById("clientsImportStatus");
const clientsImportDetails = document.getElementById("clientsImportDetails");
const downloadBranchesTemplateBtn = document.getElementById("downloadBranchesTemplateBtn");
const branchesImportFile = document.getElementById("branchesImportFile");
const importBranchesBtn = document.getElementById("importBranchesBtn");
const branchesImportStatus = document.getElementById("branchesImportStatus");
const branchesImportDetails = document.getElementById("branchesImportDetails");
const downloadUsersTemplateBtn = document.getElementById("downloadUsersTemplateBtn");
const usersImportFile = document.getElementById("usersImportFile");
const importUsersBtn = document.getElementById("importUsersBtn");
const usersImportStatus = document.getElementById("usersImportStatus");
const usersImportDetails = document.getElementById("usersImportDetails");
const meetingTypeForm = document.getElementById("meetingTypeForm");
const meetingTypeValue = document.getElementById("meetingTypeValue");
const meetingTypeLabel = document.getElementById("meetingTypeLabel");
const meetingTypeColor = document.getElementById("meetingTypeColor");
const meetingTypeFormTitle = document.getElementById("meetingTypeFormTitle");
const saveMeetingTypeBtn = document.getElementById("saveMeetingTypeBtn");
const cancelMeetingTypeEditBtn = document.getElementById("cancelMeetingTypeEditBtn");
const meetingTypeStatus = document.getElementById("meetingTypeStatus");
const meetingTypesTableBody = document.getElementById("meetingTypesTableBody");
const meetingReasonForm = document.getElementById("meetingReasonForm");
const meetingReasonName = document.getElementById("meetingReasonName");
const meetingReasonFormTitle = document.getElementById("meetingReasonFormTitle");
const saveMeetingReasonBtn = document.getElementById("saveMeetingReasonBtn");
const cancelMeetingReasonEditBtn = document.getElementById("cancelMeetingReasonEditBtn");
const meetingReasonStatus = document.getElementById("meetingReasonStatus");
const meetingReasonsTableBody = document.getElementById("meetingReasonsTableBody");
const contactRoleForm = document.getElementById("contactRoleForm");
const contactRoleName = document.getElementById("contactRoleName");
const contactRoleFormTitle = document.getElementById("contactRoleFormTitle");
const saveContactRoleBtn = document.getElementById("saveContactRoleBtn");
const cancelContactRoleEditBtn = document.getElementById("cancelContactRoleEditBtn");
const contactRoleStatus = document.getElementById("contactRoleStatus");
const contactRolesTableBody = document.getElementById("contactRolesTableBody");
const trashMeetingsTableBody = document.getElementById("trashMeetingsTableBody");
const trashMeetingsStatus = document.getElementById("trashMeetingsStatus");
const auditLogsTableBody = document.getElementById("auditLogsTableBody");
const auditLogsStatus = document.getElementById("auditLogsStatus");
const visitRulesFeatureForm = document.getElementById("visitRulesFeatureForm");
const visitRulesFeatureToggle = document.getElementById("visitRulesFeatureToggle");
const saveVisitRulesFeatureBtn = document.getElementById("saveVisitRulesFeatureBtn");
const visitRulesFeatureStatus = document.getElementById("visitRulesFeatureStatus");
const bitrixUsersPreviewForm = document.getElementById("bitrixUsersPreviewForm");
const bitrixWebhookUrl = document.getElementById("bitrixWebhookUrl");
const loadBitrixUsersBtn = document.getElementById("loadBitrixUsersBtn");
const saveBitrixWebhookBtn = document.getElementById("saveBitrixWebhookBtn");
const bitrixUsersStatus = document.getElementById("bitrixUsersStatus");
const bitrixLocalUsersCount = document.getElementById("bitrixLocalUsersCount");
const bitrixRemoteUsersCount = document.getElementById("bitrixRemoteUsersCount");
const bitrixLocalClientsCount = document.getElementById("bitrixLocalClientsCount");
const bitrixRemoteCompaniesCount = document.getElementById("bitrixRemoteCompaniesCount");
const bitrixRemoteLeadsCount = document.getElementById("bitrixRemoteLeadsCount");
const bitrixUserMappingsTableBody = document.getElementById("bitrixUserMappingsTableBody");
const bitrixClientMappingsTableBody = document.getElementById("bitrixClientMappingsTableBody");
const bitrixTaskTestForm = document.getElementById("bitrixTaskTestForm");
const bitrixTaskTitle = document.getElementById("bitrixTaskTitle");
const bitrixTaskResponsible = document.getElementById("bitrixTaskResponsible");
const bitrixTaskDescription = document.getElementById("bitrixTaskDescription");
const createBitrixTaskBtn = document.getElementById("createBitrixTaskBtn");
const bitrixTaskStatus = document.getElementById("bitrixTaskStatus");
const bitrixUsersTableBody = document.getElementById("bitrixUsersTableBody");
const visitsTableBody = document.getElementById("visitsTableBody");
const visitsVisibleCount = document.getElementById("visitsVisibleCount");
const visitsSearchInput = document.getElementById("visitsSearchInput");
const visitsStatusFilter = document.getElementById("visitsStatusFilter");
const visitsTypeFilter = document.getElementById("visitsTypeFilter");
const visitsModalityFilter = document.getElementById("visitsModalityFilter");
const visitsExecutiveFilter = document.getElementById("visitsExecutiveFilter");
const visitsSupervisorFilter = document.getElementById("visitsSupervisorFilter");
const visitsParticipantFilter = document.getElementById("visitsParticipantFilter");
const visitsDateFromFilter = document.getElementById("visitsDateFromFilter");
const visitsDateToFilter = document.getElementById("visitsDateToFilter");
const exportVisitsBtn = document.getElementById("exportVisitsBtn");
const statsExecutiveFilter = document.getElementById("statsExecutiveFilter");
const statsDateFrom = document.getElementById("statsDateFrom");
const statsDateTo = document.getElementById("statsDateTo");
const applyStatsFiltersBtn = document.getElementById("applyStatsFiltersBtn");
const resetStatsFiltersBtn = document.getElementById("resetStatsFiltersBtn");
const statsUsersStatus = document.getElementById("statsUsersStatus");
const statsTypesStatus = document.getElementById("statsTypesStatus");
const statsRulesStatus = document.getElementById("statsRulesStatus");
const statsSemaphoreStatus = document.getElementById("statsSemaphoreStatus");
const statsRulesCard = document.getElementById("statsRulesCard");
const statsSemaphoreCard = document.getElementById("statsSemaphoreCard");
const statsRulesTotal = document.getElementById("statsRulesTotal");
const statsRulesCompleted = document.getElementById("statsRulesCompleted");
const statsRulesCompletionRate = document.getElementById("statsRulesCompletionRate");
const statsSemaphoreWhiteRate = document.getElementById("statsSemaphoreWhiteRate");
const statsSemaphoreGreenRate = document.getElementById("statsSemaphoreGreenRate");
const statsSemaphoreYellowRate = document.getElementById("statsSemaphoreYellowRate");
const statsSemaphoreRedRate = document.getElementById("statsSemaphoreRedRate");
const statsSemaphoreWhiteCount = document.getElementById("statsSemaphoreWhiteCount");
const statsSemaphoreGreenCount = document.getElementById("statsSemaphoreGreenCount");
const statsSemaphoreYellowCount = document.getElementById("statsSemaphoreYellowCount");
const statsSemaphoreRedCount = document.getElementById("statsSemaphoreRedCount");
const statsUsersTableBody = document.getElementById("statsUsersTableBody");
const statsTypesTableBody = document.getElementById("statsTypesTableBody");

const listScreen = document.getElementById("listScreen");
const visitsScreen = document.getElementById("visitsScreen");
const pipelineScreen = document.getElementById("pipelineScreen");
const detailScreen = document.getElementById("detailScreen");
const editScreen = document.getElementById("editScreen");
const meetingScreen = document.getElementById("meetingScreen");
const calendarScreen = document.getElementById("calendarScreen");
const visitsGridScreen = document.getElementById("visitsGridScreen");
const visitsStatsScreen = document.getElementById("visitsStatsScreen");
const usersScreen = document.getElementById("usersScreen");
const settingsScreen = document.getElementById("settingsScreen");
const visitsGridWrap = document.getElementById("visitsGridWrap");
const visitsGridLegend = document.getElementById("visitsGridLegend");
const calendarLegend = document.getElementById("calendarLegend");

const backToList = document.getElementById("backToList");
const backFromEdit = document.getElementById("backFromEdit");
const backFromMeeting = document.getElementById("backFromMeeting");
const goToEditBtn = document.getElementById("goToEditBtn");
const goToMeetingBtn = document.getElementById("goToMeetingBtn");

const detailCompanyName = document.getElementById("detailCompanyName");
const detailMeta = document.getElementById("detailMeta");
const detailSegment = document.getElementById("detailSegment");
const detailKpis = document.getElementById("detailKpis");
const detailServices = document.getElementById("detailServices");
const crmSummaryGrid = document.getElementById("crmSummaryGrid");
const showDetailOpportunitiesBtn = document.getElementById("showDetailOpportunitiesBtn");
const showDetailVisitsBtn = document.getElementById("showDetailVisitsBtn");
const detailOpportunitiesSection = document.getElementById("detailOpportunitiesSection");
const detailVisitsSection = document.getElementById("detailVisitsSection");
const newOpportunityBtn = document.getElementById("newOpportunityBtn");
const opportunityForm = document.getElementById("opportunityForm");
const opportunityFormTitle = document.getElementById("opportunityFormTitle");
const opportunityTitle = document.getElementById("opportunityTitle");
const opportunityType = document.getElementById("opportunityType");
const opportunityServiceLine = document.getElementById("opportunityServiceLine");
const opportunityStatusSelect = document.getElementById("opportunityStatusSelect");
const opportunityAmount = document.getElementById("opportunityAmount");
const opportunityProbability = document.getElementById("opportunityProbability");
const opportunityExpectedCloseDate = document.getElementById("opportunityExpectedCloseDate");
const opportunityOwnerUserId = document.getElementById("opportunityOwnerUserId");
const opportunityBranchId = document.getElementById("opportunityBranchId");
const opportunitySource = document.getElementById("opportunitySource");
const opportunityDescription = document.getElementById("opportunityDescription");
const opportunityLossReasonGroup = document.getElementById("opportunityLossReasonGroup");
const opportunityLossReason = document.getElementById("opportunityLossReason");
const saveOpportunityBtn = document.getElementById("saveOpportunityBtn");
const cancelOpportunityEditBtn = document.getElementById("cancelOpportunityEditBtn");
const opportunityStatusMessage = document.getElementById("opportunityStatusMessage");
const followUpForm = document.getElementById("followUpForm");
const followUpFormTitle = document.getElementById("followUpFormTitle");
const followUpType = document.getElementById("followUpType");
const followUpTitle = document.getElementById("followUpTitle");
const followUpDueDate = document.getElementById("followUpDueDate");
const followUpAssignedUserId = document.getElementById("followUpAssignedUserId");
const followUpStatusSelect = document.getElementById("followUpStatusSelect");
const followUpNotes = document.getElementById("followUpNotes");
const saveFollowUpBtn = document.getElementById("saveFollowUpBtn");
const cancelFollowUpEditBtn = document.getElementById("cancelFollowUpEditBtn");
const followUpStatusMessage = document.getElementById("followUpStatusMessage");
const opportunityList = document.getElementById("opportunityList");
const meetingList = document.getElementById("meetingList");
const branchList = document.getElementById("branchList");
const meetingCardTemplate = document.getElementById("meetingCardTemplate");

const pipelineSearchInput = document.getElementById("pipelineSearchInput");
const pipelineOwnerFilter = document.getElementById("pipelineOwnerFilter");
const pipelineStatusFilter = document.getElementById("pipelineStatusFilter");
const pipelineVisibleCount = document.getElementById("pipelineVisibleCount");
const pipelineSummaryProjects = document.getElementById("pipelineSummaryProjects");
const pipelineProjectsBoard = document.getElementById("pipelineProjectsBoard");
const pipelineSummaryServices = document.getElementById("pipelineSummaryServices");
const pipelineServicesBoard = document.getElementById("pipelineServicesBoard");

const calendarGrid = document.getElementById("calendarGrid");
const calendarDayView = document.getElementById("calendarDayView");
const calendarMonthLabel = document.getElementById("calendarMonthLabel");
const prevMonthBtn = document.getElementById("prevMonthBtn");
const nextMonthBtn = document.getElementById("nextMonthBtn");
const backToMonthBtn = document.getElementById("backToMonthBtn");
const calendarParticipantFilter = document.getElementById("calendarParticipantFilter");

const editScreenTitle = document.getElementById("editScreenTitle");
const companyEditForm = document.getElementById("companyEditForm");
const editName = document.getElementById("editName");
const editSector = document.getElementById("editSector");
const editCompanyType = document.getElementById("editCompanyType");
const editCountry = document.getElementById("editCountry");
const editAccountStage = document.getElementById("editAccountStage");
const editBitrixLeadId = document.getElementById("editBitrixLeadId");
const editBitrixCompanySearch = document.getElementById("editBitrixCompanySearch");
const editBitrixCompanyId = document.getElementById("editBitrixCompanyId");
const editBitrixCompanyOptions = document.getElementById("editBitrixCompanyOptions");
const editManualBranchIdGroup = document.getElementById("editManualBranchIdGroup");
const editManualBranchId = document.getElementById("editManualBranchId");
const editManager = document.getElementById("editManager");
const editRisk = document.getElementById("editRisk");
const editSegment = document.getElementById("editSegment");
const accountRolesGrid = document.getElementById("accountRolesGrid");
const editFixedFire = document.getElementById("editFixedFire");
const editExtinguishers = document.getElementById("editExtinguishers");
const editWorks = document.getElementById("editWorks");
const fixedFireSupervisorGroup = document.getElementById("fixedFireSupervisorGroup");
const extinguishersSupervisorGroup = document.getElementById("extinguishersSupervisorGroup");
const worksSupervisorGroup = document.getElementById("worksSupervisorGroup");
const editSupervisorFixedFire = document.getElementById("editSupervisorFixedFire");
const editSupervisorExtinguishers = document.getElementById("editSupervisorExtinguishers");
const editSupervisorWorks = document.getElementById("editSupervisorWorks");
const editNotes = document.getElementById("editNotes");
const addVisitRuleBtn = document.getElementById("addVisitRuleBtn");
const visitRulesList = document.getElementById("visitRulesList");
const visitRulesEditorSection = document.getElementById("visitRulesEditorSection");
const saveCompanyBtn = document.getElementById("saveCompanyBtn");
const hideCompanyBtn = document.getElementById("hideCompanyBtn");
const editCompanyStatus = document.getElementById("editCompanyStatus");
const addBranchBtn = document.getElementById("addBranchBtn");

const meetingScreenTitle = document.getElementById("meetingScreenTitle");
const meetingScreenMeta = document.getElementById("meetingScreenMeta");
const meetingForm = document.getElementById("meetingForm");
const meetingKind = document.getElementById("meetingKind");
const meetingSubject = document.getElementById("meetingSubject");
const meetingObjective = document.getElementById("meetingObjective");
const meetingDate = document.getElementById("meetingDate");
const meetingModality = document.getElementById("meetingModality");
const meetingScope = document.getElementById("meetingScope");
const meetingParticipantsPicker = document.getElementById("meetingParticipantsPicker");
const meetingParticipantsSummary = document.getElementById("meetingParticipantsSummary");
const meetingParticipantsSearch = document.getElementById("meetingParticipantsSearch");
const meetingParticipantsList = document.getElementById("meetingParticipantsList");
const meetingContactName = document.getElementById("meetingContactName");
const meetingContactRole = document.getElementById("meetingContactRole");
const meetingStatusSelect = document.getElementById("meetingStatusSelect");
const meetingOpportunity = document.getElementById("meetingOpportunity");
const meetingNextDateGroup = document.getElementById("meetingNextDateGroup");
const meetingNextDate = document.getElementById("meetingNextDate");
const meetingMinutes = document.getElementById("meetingMinutes");
const meetingNegotiationStatus = document.getElementById("meetingNegotiationStatus");
const meetingOpportunities = document.getElementById("meetingOpportunities");
const meetingSubstituteRecovery = document.getElementById("meetingSubstituteRecovery");
const meetingGlobalContactsGroup = document.getElementById("meetingGlobalContactsGroup");
const meetingGlobalContacts = document.getElementById("meetingGlobalContacts");
const meetingServiceHealthStatusGroup = document.getElementById("meetingServiceHealthStatusGroup");
const meetingServiceHealthStatus = document.getElementById("meetingServiceHealthStatus");
const meetingServiceStatusGroup = document.getElementById("meetingServiceStatusGroup");
const meetingServiceStatus = document.getElementById("meetingServiceStatus");
const meetingComplaintBitrixResponsibleGroup = document.getElementById("meetingComplaintBitrixResponsibleGroup");
const meetingComplaintBitrixResponsible = document.getElementById("meetingComplaintBitrixResponsible");
const saveMeetingBtn = document.getElementById("saveMeetingBtn");
const meetingStatus = document.getElementById("meetingStatus");
const meetingDetailView = document.getElementById("meetingDetailView");
const editMeetingFromDetailBtn = document.getElementById("editMeetingFromDetailBtn");

let clients = [];
let selectedId = null;
let selectedClient = null;
let selectedBranchView = null;
let currentUser = null;
let editMode = "edit";
let editEntityType = "client";
let meetingMode = "create";
let meetingScreenMode = "form";
let editingMeetingId = null;
let editingBranchId = null;
let searchTimer = null;
let meetingTypes = [];
let meetingReasons = [];
let contactRoles = [];
let meetingStatuses = [];
let meetingModalities = [];
let editingMeetingTypeId = null;
let editingMeetingReasonId = null;
let editingContactRoleId = null;
let deletedMeetings = [];
let auditLogs = [];
let bitrixPreviewUsers = [];
let bitrixUserMappings = [];
let bitrixClientMappings = [];
let bitrixDirectory = {
  users: [],
  companies: []
};
let bitrixDirectoryPermissions = {
  users: true,
  companies: true
};
let bitrixDirectoryErrors = {
  users: "",
  companies: ""
};
let activeBitrixSearchDropdown = null;
let bitrixCompanySearchResults = [];
let bitrixCompanySearchRequestId = 0;
let bitrixCounts = {
  localUsers: 0,
  localClients: 0,
  bitrixUsers: 0,
  bitrixCompanies: 0,
  bitrixLeads: 0
};
let bitrixPermissions = {
  users: false,
  companies: false,
  leads: false
};
let bitrixErrors = {
  companies: "",
  leads: ""
};
let visitRulesFeatureEnabled = true;
let currentSettingsTab = "catalogs";
let calendarMeetings = [];
let calendarDate = new Date();
let calendarView = "month";
let selectedCalendarDay = null;
let suppressHashRouting = false;
let users = [];
let userRoles = [];
let userFormMode = "create";
let editingUserId = null;
let sectorOptions = [];
let currentVisitRules = [];
let visits = [];
let visitsGridData = [];
let visitStats = {
  byUser: [],
  byType: [],
  byRule: {
    totalCount: 0,
    completedCount: 0,
    completionRate: 0
  }
};
let visitsSearchTimer = null;
let pipelineSearchTimer = null;
let visitsSort = {
  key: "scheduledFor",
  direction: "desc"
};
let selectedParticipantUserIds = [];
let detailSection = CRM_ENABLED ? "opportunities" : "visits";
let crmCatalogs = {
  opportunityStatuses: [],
  openOpportunityStatuses: [],
  opportunityTypes: [],
  serviceLines: [],
  followUpStatuses: [],
  followUpTypes: []
};
let editingOpportunityId = null;
let editingFollowUpId = null;
let followUpOpportunityId = null;
let pipelineOpportunities = [];
let pipelineSummaryData = {
  openCount: 0,
  totalAmount: 0,
  weightedAmount: 0,
  overdueFollowUps: 0,
  dueThisWeek: 0
};
let assignmentOptions = {
  allUsers: [],
  executives: [],
  supervisors: {
    fixedFire: [],
    extinguishers: [],
    works: []
  }
};

function canAccessSettings() {
  const email = String(currentUser?.email || "").trim().toLowerCase();
  const name = String(currentUser?.name || "").trim();
  return !!currentUser?.canManageSettings || (email === "nicolas@maxiseguridad.com" && name === "Nicolas Beguelman");
}

function getCurrentEntityPermissions() {
  return {
    createClients: currentUser?.permissions?.createClients !== false,
    editClients: currentUser?.permissions?.editClients !== false,
    hideClients: currentUser?.permissions?.hideClients !== false,
    createBranches: currentUser?.permissions?.createBranches !== false,
    editBranches: currentUser?.permissions?.editBranches !== false,
    hideBranches: currentUser?.permissions?.hideBranches !== false
  };
}

function getUserPermissionsPayload() {
  return {
    createClients: !!permCreateClients?.checked,
    editClients: !!permEditClients?.checked,
    hideClients: !!permHideClients?.checked,
    createBranches: !!permCreateBranches?.checked,
    editBranches: !!permEditBranches?.checked,
    hideBranches: !!permHideBranches?.checked
  };
}

function applyUserPermissionsForm(permissions = {}) {
  const normalized = {
    createClients: permissions.createClients !== false,
    editClients: permissions.editClients !== false,
    hideClients: permissions.hideClients !== false,
    createBranches: permissions.createBranches !== false,
    editBranches: permissions.editBranches !== false,
    hideBranches: permissions.hideBranches !== false
  };
  permCreateClients.checked = normalized.createClients;
  permEditClients.checked = normalized.editClients;
  permHideClients.checked = normalized.hideClients;
  permCreateBranches.checked = normalized.createBranches;
  permEditBranches.checked = normalized.editBranches;
  permHideBranches.checked = normalized.hideBranches;
}

function summarizePermissions(permissions = {}) {
  const companyPermissions = [];
  const branchPermissions = [];
  if (permissions.createClients !== false) companyPermissions.push("crear");
  if (permissions.editClients !== false) companyPermissions.push("editar");
  if (permissions.hideClients !== false) companyPermissions.push("ocultar");
  if (permissions.createBranches !== false) branchPermissions.push("crear");
  if (permissions.editBranches !== false) branchPermissions.push("editar");
  if (permissions.hideBranches !== false) branchPermissions.push("ocultar");
  return [
    `Compañías: ${companyPermissions.length ? companyPermissions.join(", ") : "sin acceso"}`,
    `Sucursales: ${branchPermissions.length ? branchPermissions.join(", ") : "sin acceso"}`
  ].join(" · ");
}

function notifyError(message) {
  window.alert(message);
}

function updateAuthUi() {
  const loggedIn = !!currentUser;
  loginScreen.classList.toggle("hidden", loggedIn);
  appContent.classList.toggle("hidden", !loggedIn);
  crmMenu.classList.toggle("hidden", !CRM_ENABLED);
  showSettingsBtn.classList.toggle("hidden", !loggedIn || !canAccessSettings());
  currentUserBadge.innerHTML = loggedIn
    ? `<strong>${currentUser.name}</strong><small>${currentUser.role}</small>`
    : "";
  applyEntityPermissionsUi();
}

function applyFeatureFlagsUi() {
  visitRulesEditorSection.classList.toggle("hidden", !visitRulesFeatureEnabled);
  statsRulesCard.classList.toggle("hidden", !visitRulesFeatureEnabled);
  statsSemaphoreCard.classList.toggle("hidden", !visitRulesFeatureEnabled);
  if (visitRulesFeatureToggle) {
    visitRulesFeatureToggle.checked = !!visitRulesFeatureEnabled;
  }
}

function applyEntityPermissionsUi() {
  const permissions = getCurrentEntityPermissions();
  addCompanyBtn.classList.toggle("hidden", !currentUser || !permissions.createClients);
}

function syncModalState() {
  const editOpen = !editScreen.classList.contains("hidden");
  const meetingOpen = !meetingScreen.classList.contains("hidden");
  const opportunityOpen = !opportunityForm.classList.contains("hidden");
  const followUpOpen = !followUpForm.classList.contains("hidden");
  document.body.classList.toggle("modal-active", editOpen || meetingOpen || opportunityOpen || followUpOpen);
}

function parseNumericInput(input, { integer = false, min = null, max = null } = {}) {
  const raw = String(input?.value ?? "").trim();
  if (!raw) return { ok: false };

  const normalized = raw.replace(",", ".");
  const value = Number(normalized);
  if (!Number.isFinite(value)) return { ok: false };
  if (integer && !Number.isInteger(value)) return { ok: false };
  if (min !== null && value < min) return { ok: false };
  if (max !== null && value > max) return { ok: false };

  return { ok: true, value };
}

function formatDate(dateValue) {
  if (!dateValue) return "Sin fecha";
  const date = new Date(`${dateValue}T00:00:00`);
  return new Intl.DateTimeFormat("es-AR", {
    day: "2-digit",
    month: "short",
    year: "numeric"
  }).format(date);
}

function monthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function monthLabel(date) {
  return new Intl.DateTimeFormat("es-AR", {
    month: "long",
    year: "numeric"
  }).format(date);
}

function dateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function fullDateLabel(dateValue) {
  if (!dateValue) return "";
  const date = new Date(`${dateValue}T00:00:00`);
  return new Intl.DateTimeFormat("es-AR", {
    weekday: "long",
    day: "2-digit",
    month: "long",
    year: "numeric"
  }).format(date);
}

function getMeetingTypeMeta(kind) {
  return meetingTypes.find((type) => type.value === kind) || {
    value: kind,
    label: kind,
    color: "yellow"
  };
}

function getOpenOpportunityStatuses() {
  return crmCatalogs.openOpportunityStatuses || [];
}

function isOpenOpportunity(status) {
  return getOpenOpportunityStatuses().includes(status);
}

function formatMoney(amount) {
  const numericAmount = Number(amount || 0);
  return new Intl.NumberFormat("es-AR", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0
  }).format(numericAmount);
}

function getOpportunityStatusClass(status) {
  const normalized = String(status || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "-");
  return `opportunity-status-${normalized || "default"}`;
}

function getFollowUpStatusClass(status) {
  const normalized = String(status || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "-");
  return `follow-up-status-${normalized || "default"}`;
}

function getMeetingColorLabel(colorValue) {
  return MEETING_COLOR_OPTIONS.find((option) => option.value === colorValue)?.label || colorValue;
}

function normalizeParticipantIds(ids) {
  return [...new Set((ids || []).map((id) => Number(id)).filter((id) => Number.isInteger(id) && id > 0))];
}

function getParticipantUsers() {
  return assignmentOptions.allUsers || [];
}

function getParticipantNamesFromIds(ids) {
  const selectedIds = new Set(normalizeParticipantIds(ids));
  return getParticipantUsers()
    .filter((user) => selectedIds.has(Number(user.id)))
    .map((user) => user.name);
}

function setSelectedParticipantUserIds(ids) {
  selectedParticipantUserIds = normalizeParticipantIds(ids);
  const selectedNames = getParticipantNamesFromIds(selectedParticipantUserIds);
  meetingParticipantsSummary.textContent = selectedNames.length ? selectedNames.join(", ") : "Seleccionar participantes";
}

function renderMeetingParticipantsPicker() {
  const searchTerm = (meetingParticipantsSearch.value || "").trim().toLowerCase();
  const selectedIds = new Set(selectedParticipantUserIds);
  const matchingUsers = getParticipantUsers().filter((user) => {
    if (!searchTerm) return true;
    return (
      user.name.toLowerCase().includes(searchTerm) ||
      String(user.email || "").toLowerCase().includes(searchTerm) ||
      String(user.role || "").toLowerCase().includes(searchTerm)
    );
  });

  if (!matchingUsers.length) {
    meetingParticipantsList.innerHTML = '<div class="multi-select-empty">No hay usuarios que coincidan con la búsqueda.</div>';
    return;
  }

  meetingParticipantsList.innerHTML = matchingUsers
    .map(
      (user) => `
        <label class="multi-select-option">
          <input type="checkbox" value="${user.id}" ${selectedIds.has(Number(user.id)) ? "checked" : ""} />
          <span>
            <strong>${user.name}</strong>
            <small>${user.role}</small>
          </span>
        </label>
      `
    )
    .join("");

  meetingParticipantsList.querySelectorAll('input[type="checkbox"]').forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      const nextIds = new Set(selectedParticipantUserIds);
      const userId = Number(checkbox.value);
      if (checkbox.checked) nextIds.add(userId);
      else nextIds.delete(userId);
      setSelectedParticipantUserIds([...nextIds]);
      renderMeetingParticipantsPicker();
    });
  });
}

function getCalendarStatusClass(status) {
  if (status === "Realizada") return "calendar-status-completed";
  if (status === "Confirmada") return "calendar-status-confirmed";
  return "calendar-status-scheduled";
}

function renderMeetingTypeColorOptions() {
  const currentColor = meetingTypeColor.value || "yellow";
  meetingTypeColor.innerHTML = MEETING_COLOR_OPTIONS
    .map((option) => `<option value="${option.value}">${option.label}</option>`)
    .join("");
  meetingTypeColor.value = MEETING_COLOR_OPTIONS.some((option) => option.value === currentColor) ? currentColor : "yellow";
}

function renderTypeSelectOptions() {
  const currentVisitType = visitsTypeFilter.value;
  const currentVisitStatus = visitsStatusFilter.value;
  const currentContactRole = meetingContactRole.value;
  meetingKind.innerHTML = meetingTypes
    .map((type) => `<option value="${type.value}">${type.label}</option>`)
    .join("");

  const currentSubject = meetingSubject.value;
  meetingSubject.innerHTML = meetingReasons
    .map((reason) => {
      const name = typeof reason === "string" ? reason : reason.name;
      return `<option value="${name}">${name}</option>`;
    })
    .join("");
  if (currentSubject && meetingReasons.some((reason) => (typeof reason === "string" ? reason : reason.name) === currentSubject)) {
    meetingSubject.value = currentSubject;
  }

  meetingStatusSelect.innerHTML = meetingStatuses
    .map((status) => `<option value="${status}">${status}</option>`)
    .join("");

  meetingModality.innerHTML = meetingModalities
    .map((modality) => `<option value="${modality}">${modality}</option>`)
    .join("");

  meetingContactRole.innerHTML = [
    '<option value="">Seleccionar función</option>',
    ...contactRoles.map((role) => `<option value="${role.name}">${role.name}</option>`)
  ].join("");
  if (currentContactRole && contactRoles.some((role) => role.name === currentContactRole)) {
    meetingContactRole.value = currentContactRole;
  }

  visitsTypeFilter.innerHTML = ['<option value="todos">Todos</option>']
    .concat(meetingTypes.map((type) => `<option value="${type.value}">${type.label}</option>`))
    .join("");

  visitsStatusFilter.innerHTML = ['<option value="todos">Todos</option>']
    .concat(meetingStatuses.map((status) => `<option value="${status}">${status}</option>`))
    .join("");

  visitsTypeFilter.value = currentVisitType || "todos";
  visitsStatusFilter.value = currentVisitStatus || "todos";
}

function renderOpportunitySelects() {
  const opportunityStatuses = crmCatalogs.opportunityStatuses?.length ? crmCatalogs.opportunityStatuses : DEFAULT_OPPORTUNITY_STATUSES;
  const opportunityTypes = crmCatalogs.opportunityTypes?.length ? crmCatalogs.opportunityTypes : DEFAULT_OPPORTUNITY_TYPES;
  const serviceLines = crmCatalogs.serviceLines?.length ? crmCatalogs.serviceLines : DEFAULT_SERVICE_LINES;
  const followUpTypes = crmCatalogs.followUpTypes?.length ? crmCatalogs.followUpTypes : DEFAULT_FOLLOW_UP_TYPES;
  const followUpStatuses = crmCatalogs.followUpStatuses?.length ? crmCatalogs.followUpStatuses : DEFAULT_FOLLOW_UP_STATUSES;
  const currentPipelineStatus = pipelineStatusFilter.value;
  const currentOpportunityStatus = opportunityStatusSelect.value;
  const currentOpportunityType = opportunityType.value;
  const currentServiceLine = opportunityServiceLine.value;
  const currentFollowUpType = followUpType.value;
  const currentFollowUpStatus = followUpStatusSelect.value;

  opportunityStatusSelect.innerHTML = opportunityStatuses
    .map((status) => `<option value="${status}">${status}</option>`)
    .join("");
  opportunityType.innerHTML = opportunityTypes
    .map((type) => `<option value="${type}">${type}</option>`)
    .join("");
  pipelineStatusFilter.innerHTML = ['<option value="todos">Todas</option>']
    .concat(opportunityStatuses.map((status) => `<option value="${status}">${status}</option>`))
    .join("");
  opportunityServiceLine.innerHTML = serviceLines
    .map((serviceLine) => `<option value="${serviceLine}">${serviceLine}</option>`)
    .join("");
  followUpType.innerHTML = followUpTypes
    .map((type) => `<option value="${type}">${type}</option>`)
    .join("");
  followUpStatusSelect.innerHTML = followUpStatuses
    .map((status) => `<option value="${status}">${status}</option>`)
    .join("");

  if (currentOpportunityStatus && opportunityStatuses.includes(currentOpportunityStatus)) {
    opportunityStatusSelect.value = currentOpportunityStatus;
  }
  if (currentOpportunityType && opportunityTypes.includes(currentOpportunityType)) {
    opportunityType.value = currentOpportunityType;
  }
  if (currentPipelineStatus && opportunityStatuses.includes(currentPipelineStatus)) {
    pipelineStatusFilter.value = currentPipelineStatus;
  }
  if (currentServiceLine && serviceLines.includes(currentServiceLine)) {
    opportunityServiceLine.value = currentServiceLine;
  }
  if (currentFollowUpType && followUpTypes.includes(currentFollowUpType)) {
    followUpType.value = currentFollowUpType;
  }
  if (currentFollowUpStatus && followUpStatuses.includes(currentFollowUpStatus)) {
    followUpStatusSelect.value = currentFollowUpStatus;
  }

  syncOpportunityLossReasonVisibility();
}

function renderOpportunityOwnerSelects() {
  const users = assignmentOptions.allUsers || [];
  const currentOwner = opportunityOwnerUserId.value;
  const currentAssigned = followUpAssignedUserId.value;
  const currentPipelineOwner = pipelineOwnerFilter.value;

  const options = ['<option value="">Sin asignar</option>']
    .concat(users.map((user) => `<option value="${user.id}">${user.name}</option>`))
    .join("");

  opportunityOwnerUserId.innerHTML = options;
  followUpAssignedUserId.innerHTML = options;
  pipelineOwnerFilter.innerHTML = ['<option value="todos">Todos</option>']
    .concat(users.map((user) => `<option value="${user.id}">${user.name}</option>`))
    .join("");

  opportunityOwnerUserId.value = currentOwner || "";
  followUpAssignedUserId.value = currentAssigned || "";
  pipelineOwnerFilter.value = currentPipelineOwner || "todos";
}

function renderMeetingOpportunityOptions(selectedOpportunityId = "") {
  const opportunities = Array.isArray(selectedClient?.opportunities) ? selectedClient.opportunities : [];
  meetingOpportunity.innerHTML = ['<option value="">Sin vincular</option>']
    .concat(
      opportunities.map(
        (opportunity) =>
          `<option value="${opportunity.id}">${opportunity.title} · ${opportunity.status}${opportunity.branchName ? ` · ${opportunity.branchName}` : ""}</option>`
      )
    )
    .join("");
  meetingOpportunity.value = selectedOpportunityId ? String(selectedOpportunityId) : "";
}

function renderOpportunityBranchOptions(selectedBranchId = "") {
  if (!selectedClient) {
    opportunityBranchId.innerHTML = '<option value="">Casa matriz</option>';
    return;
  }

  opportunityBranchId.innerHTML = [
    '<option value="">Casa matriz</option>',
    ...(selectedClient.branches || []).map((branch) => `<option value="${branch.id}">${branch.name}</option>`)
  ].join("");
  opportunityBranchId.value = selectedBranchId ? String(selectedBranchId) : "";
}

function getMeetingScopeEntity() {
  if (!selectedClient) return null;
  if (meetingScope.value.startsWith("branch:")) {
    const branchId = Number(meetingScope.value.split(":")[1]);
    return selectedClient.branches?.find((branch) => Number(branch.id) === branchId) || null;
  }
  return selectedBranchView || selectedClient;
}

function syncMeetingContextBlocks() {
  const scopeEntity = getMeetingScopeEntity();
  const isGlobalCompany = selectedClient?.companyType === "Global";
  const hasActiveServices =
    selectedClient?.accountStage !== "Prospecto" &&
    !!scopeEntity &&
    Object.values(scopeEntity.services || {}).some(Boolean);

  meetingGlobalContactsGroup.classList.toggle("hidden", !isGlobalCompany);
  meetingServiceHealthStatusGroup.classList.toggle("hidden", !hasActiveServices);
  meetingServiceStatusGroup.classList.toggle("hidden", !hasActiveServices);
  meetingComplaintBitrixResponsibleGroup.classList.toggle("hidden", !hasActiveServices);

  if (!isGlobalCompany) meetingGlobalContacts.value = "";
  if (!hasActiveServices) meetingServiceHealthStatus.value = "";
  if (!hasActiveServices) meetingServiceStatus.value = "";
  if (!hasActiveServices) meetingComplaintBitrixResponsible.value = "";
}

function renderComplaintBitrixResponsibleOptions(selectedValue = "") {
  meetingComplaintBitrixResponsible.innerHTML = ['<option value="">Seleccionar responsable</option>']
    .concat(
      BITRIX_COMPLAINT_RESPONSIBLES.map((item) => `<option value="${item.id}">${item.name}</option>`)
    )
    .join("");
  meetingComplaintBitrixResponsible.value = selectedValue || "";
}

function syncOpportunityLossReasonVisibility() {
  const showLossReason = opportunityStatusSelect.value === "Perdida";
  opportunityLossReasonGroup.classList.toggle("hidden", !showLossReason);
  if (!showLossReason) {
    opportunityLossReason.value = "";
  }
}

function resetMeetingTypeForm() {
  editingMeetingTypeId = null;
  meetingTypeFormTitle.textContent = "Nuevo tipo de reunión";
  saveMeetingTypeBtn.textContent = "Guardar tipo";
  cancelMeetingTypeEditBtn.classList.add("hidden");
  meetingTypeValue.value = "";
  meetingTypeLabel.value = "";
  meetingTypeColor.value = "yellow";
  meetingTypeStatus.textContent = "";
}

function resetMeetingReasonForm() {
  editingMeetingReasonId = null;
  meetingReasonFormTitle.textContent = "Nuevo motivo de reunión";
  saveMeetingReasonBtn.textContent = "Guardar motivo";
  cancelMeetingReasonEditBtn.classList.add("hidden");
  meetingReasonName.value = "";
  meetingReasonStatus.textContent = "";
}

function resetContactRoleForm() {
  editingContactRoleId = null;
  contactRoleFormTitle.textContent = "Nueva función del contacto";
  saveContactRoleBtn.textContent = "Guardar función";
  cancelContactRoleEditBtn.classList.add("hidden");
  contactRoleName.value = "";
  contactRoleStatus.textContent = "";
}

function renderMeetingTypesConfig() {
  meetingTypesTableBody.innerHTML = "";

  meetingTypes.forEach((type) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td><b>${type.value}</b></td>
      <td>${type.label}</td>
      <td><span class="meeting-type-badge ${type.color}">${getMeetingColorLabel(type.color)}</span></td>
      <td>
        <div class="table-actions">
          <button class="secondary-btn edit-meeting-type-btn" type="button">Editar</button>
          <button class="ghost-btn delete-meeting-type-btn" type="button">Eliminar</button>
        </div>
      </td>
    `;

    row.querySelector(".edit-meeting-type-btn").addEventListener("click", () => {
      editingMeetingTypeId = type.id;
      meetingTypeFormTitle.textContent = `Editar tipo · ${type.label}`;
      saveMeetingTypeBtn.textContent = "Guardar cambios";
      cancelMeetingTypeEditBtn.classList.remove("hidden");
      meetingTypeValue.value = type.value;
      meetingTypeLabel.value = type.label;
      meetingTypeColor.value = type.color;
      meetingTypeStatus.textContent = "";
      window.scrollTo({ top: 0, behavior: "smooth" });
    });

    row.querySelector(".delete-meeting-type-btn").addEventListener("click", async () => {
      const confirmed = window.confirm(`Se va a eliminar el tipo de reunión ${type.label}.`);
      if (!confirmed) return;

      try {
        await deleteMeetingType(type.id);
        await loadMeetingTypesConfig();
        resetMeetingTypeForm();
      } catch (error) {
        notifyError(error.message);
      }
    });

    meetingTypesTableBody.appendChild(row);
  });
}

function renderMeetingReasonsConfig() {
  meetingReasonsTableBody.innerHTML = "";

  meetingReasons.forEach((reason) => {
    const reasonName = typeof reason === "string" ? reason : reason.name;
    const reasonId = typeof reason === "string" ? null : reason.id;
    const row = document.createElement("tr");
    row.innerHTML = `
      <td><b>${reasonName}</b></td>
      <td>
        <div class="table-actions">
          <button class="secondary-btn edit-meeting-reason-btn" type="button">Editar</button>
          <button class="ghost-btn delete-meeting-reason-btn" type="button">Eliminar</button>
        </div>
      </td>
    `;

    row.querySelector(".edit-meeting-reason-btn").addEventListener("click", () => {
      editingMeetingReasonId = reasonId;
      meetingReasonFormTitle.textContent = `Editar motivo · ${reasonName}`;
      saveMeetingReasonBtn.textContent = "Guardar cambios";
      cancelMeetingReasonEditBtn.classList.remove("hidden");
      meetingReasonName.value = reasonName;
      meetingReasonStatus.textContent = "";
      window.scrollTo({ top: 0, behavior: "smooth" });
    });

    row.querySelector(".delete-meeting-reason-btn").addEventListener("click", async () => {
      const confirmed = window.confirm(`Se va a eliminar el motivo ${reasonName}.`);
      if (!confirmed) return;

      try {
        await deleteMeetingReason(reasonId);
        await loadMeetingReasonsConfig();
        resetMeetingReasonForm();
      } catch (error) {
        notifyError(error.message);
      }
    });

    meetingReasonsTableBody.appendChild(row);
  });
}

function renderContactRolesConfig() {
  contactRolesTableBody.innerHTML = "";

  contactRoles.forEach((role) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td><b>${role.name}</b></td>
      <td>
        <div class="table-actions">
          <button class="secondary-btn edit-contact-role-btn" type="button">Editar</button>
          <button class="ghost-btn delete-contact-role-btn" type="button">Eliminar</button>
        </div>
      </td>
    `;

    row.querySelector(".edit-contact-role-btn").addEventListener("click", () => {
      editingContactRoleId = role.id;
      contactRoleFormTitle.textContent = `Editar función · ${role.name}`;
      saveContactRoleBtn.textContent = "Guardar cambios";
      cancelContactRoleEditBtn.classList.remove("hidden");
      contactRoleName.value = role.name;
      contactRoleStatus.textContent = "";
      window.scrollTo({ top: 0, behavior: "smooth" });
    });

    row.querySelector(".delete-contact-role-btn").addEventListener("click", async () => {
      const confirmed = window.confirm(`Se va a eliminar la función ${role.name}.`);
      if (!confirmed) return;

      try {
        await deleteContactRole(role.id);
        await loadContactRolesConfig();
        resetContactRoleForm();
      } catch (error) {
        notifyError(error.message);
      }
    });

    contactRolesTableBody.appendChild(row);
  });
}

function formatDateTime(dateValue) {
  if (!dateValue) return "-";
  const date = new Date(dateValue);
  if (Number.isNaN(date.getTime())) return String(dateValue);
  return new Intl.DateTimeFormat("es-AR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  }).format(date);
}

function renderDeletedMeetings() {
  trashMeetingsTableBody.innerHTML = "";

  if (!deletedMeetings.length) {
    const row = document.createElement("tr");
    row.innerHTML = '<td colspan="5" class="muted-inline">No hay reuniones en papelera.</td>';
    trashMeetingsTableBody.appendChild(row);
    return;
  }

  deletedMeetings.forEach((meeting) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${formatDate(meeting.scheduledFor)}</td>
      <td>${meeting.branchName ? `${meeting.clientName} · ${meeting.branchName}` : meeting.clientName}</td>
      <td><b>${meeting.subject}</b><br /><small>${meeting.kindLabel || meeting.kind}</small></td>
      <td>${meeting.deletedBy || "Sin dato"}<br /><small>${formatDateTime(meeting.deletedAt)}</small></td>
      <td>
        <div class="table-actions">
          <button class="secondary-btn restore-meeting-btn" type="button">Recuperar</button>
        </div>
      </td>
    `;

    row.querySelector(".restore-meeting-btn").addEventListener("click", async () => {
      const confirmed = window.confirm(`Se va a recuperar la reunión "${meeting.subject}".`);
      if (!confirmed) return;

      try {
        trashMeetingsStatus.textContent = "Recuperando reunión...";
        await restoreDeletedMeeting(meeting.id);
        await loadSettingsCatalogs();
        await loadClients({ clearSelectionWhenMissing: false });
        await loadCalendar();
        renderTable();
        trashMeetingsStatus.textContent = "Reunión recuperada";
      } catch (error) {
        trashMeetingsStatus.textContent = error.message;
      }
    });

    trashMeetingsTableBody.appendChild(row);
  });
}

function formatAuditAction(log) {
  const actionLabels = {
    "auth.login": "Inicio de sesión",
    "auth.logout": "Cierre de sesión",
    "client.create": "Alta de compañía",
    "client.update": "Edición de compañía",
    "client.hide": "Ocultar compañía",
    "branch.create": "Alta de sucursal",
    "branch.update": "Edición de sucursal",
    "meeting.create": "Alta de reunión",
    "meeting.update": "Edición de reunión",
    "meeting.delete": "Enviar a papelera",
    "meeting.restore": "Recuperar reunión",
    "user.create": "Alta de usuario",
    "user.update": "Edición de usuario",
    "user.delete": "Baja de usuario",
    "settings.sector.create": "Alta de sector",
    "settings.sector.delete": "Baja de sector",
    "settings.meeting_type.create": "Alta de tipo de reunión",
    "settings.meeting_type.update": "Edición de tipo de reunión",
    "settings.meeting_type.delete": "Baja de tipo de reunión",
    "settings.meeting_reason.create": "Alta de motivo",
    "settings.meeting_reason.update": "Edición de motivo",
    "settings.meeting_reason.delete": "Baja de motivo",
    "settings.contact_role.create": "Alta de función de contacto",
    "settings.contact_role.update": "Edición de función de contacto",
    "settings.contact_role.delete": "Baja de función de contacto",
    "settings.bitrix.config.update": "Actualización de webhook Bitrix",
    "settings.bitrix.users_preview": "Consulta de usuarios Bitrix",
    "settings.bitrix.mappings_preview": "Validación de mapeos Bitrix",
    "settings.bitrix.task_test_create": "Alta de tarea de prueba en Bitrix",
    "settings.features.update": "Actualización de módulos"
  };
  return actionLabels[log.action] || log.action;
}

function renderAuditLogs() {
  auditLogsTableBody.innerHTML = "";

  if (!auditLogs.length) {
    const row = document.createElement("tr");
    row.innerHTML = '<td colspan="5" class="muted-inline">Todavía no hay movimientos auditados.</td>';
    auditLogsTableBody.appendChild(row);
    return;
  }

  auditLogs.forEach((log) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${formatDateTime(log.createdAt)}</td>
      <td><b>${log.userName || "Sistema"}</b><br /><small>${log.userEmail || "-"}</small></td>
      <td>${formatAuditAction(log)}</td>
      <td>${log.entityName || log.entityType || "-"}</td>
      <td>${log.details && Object.keys(log.details).length ? Object.entries(log.details).map(([key, value]) => `${key}: ${value}`).join(" · ") : "-"}</td>
    `;
    auditLogsTableBody.appendChild(row);
  });
}

function renderBitrixUsersPreview() {
  bitrixUsersTableBody.innerHTML = "";

  if (!bitrixPreviewUsers.length) {
    const row = document.createElement("tr");
    row.innerHTML = '<td colspan="5" class="muted-inline">Todavía no consultaste usuarios de Bitrix.</td>';
    bitrixUsersTableBody.appendChild(row);
    return;
  }

  bitrixPreviewUsers.forEach((user) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td><b>${user.id || "-"}</b></td>
      <td>${user.fullName || "-"}</td>
      <td>${user.email || "-"}</td>
      <td>${user.workPosition || "-"}</td>
      <td>${user.active ? "Activo" : "Inactivo"}${user.admin ? " · Admin" : ""}</td>
    `;
    bitrixUsersTableBody.appendChild(row);
  });
}

function renderBitrixTaskResponsibleOptions() {
  const currentValue = bitrixTaskResponsible.value;
  bitrixTaskResponsible.innerHTML = ['<option value="">Seleccionar responsable</option>']
    .concat(bitrixPreviewUsers.map((user) => `<option value="${user.id}">${user.fullName} (${user.id})</option>`))
    .join("");

  if (currentValue && bitrixPreviewUsers.some((user) => String(user.id) === String(currentValue))) {
    bitrixTaskResponsible.value = currentValue;
  } else if (bitrixPreviewUsers.length) {
    bitrixTaskResponsible.value = String(bitrixPreviewUsers[0].id);
  }
}

function renderBitrixMappings() {
  bitrixLocalUsersCount.textContent = String(bitrixCounts.localUsers || 0);
  bitrixRemoteUsersCount.textContent = String(bitrixCounts.bitrixUsers || 0);
  bitrixLocalClientsCount.textContent = String(bitrixCounts.localClients || 0);
  bitrixRemoteCompaniesCount.textContent = String(bitrixCounts.bitrixCompanies || 0);
  bitrixRemoteLeadsCount.textContent = String(bitrixCounts.bitrixLeads || 0);

  bitrixUserMappingsTableBody.innerHTML = "";
  if (!bitrixUserMappings.length) {
    const row = document.createElement("tr");
    row.innerHTML = '<td colspan="5" class="muted-inline">Todavía no se generó el mapeo de usuarios.</td>';
    bitrixUserMappingsTableBody.appendChild(row);
  } else {
    bitrixUserMappings.forEach((mapping) => {
      const row = document.createElement("tr");
      const statusLabel =
        mapping.status === "mapped"
          ? "Vinculado"
          : mapping.status === "missing"
            ? "No encontrado"
            : "Sin ID";
      row.innerHTML = `
        <td><b>${mapping.appUser?.name || "-"}</b><br /><small>${mapping.appUser?.email || "-"}</small></td>
        <td>${mapping.localBitrixUserId || "-"}</td>
        <td>${mapping.bitrixUser?.fullName || "-"}</td>
        <td>${mapping.bitrixUser?.email || "-"}</td>
        <td>${statusLabel}</td>
      `;
      bitrixUserMappingsTableBody.appendChild(row);
    });
  }

  bitrixClientMappingsTableBody.innerHTML = "";
  if (!bitrixClientMappings.length) {
    const row = document.createElement("tr");
    row.innerHTML = '<td colspan="5" class="muted-inline">Todavía no se generó el mapeo de clientes.</td>';
    bitrixClientMappingsTableBody.appendChild(row);
  } else {
    bitrixClientMappings.forEach((mapping) => {
      const row = document.createElement("tr");
      const matchedLabel = mapping.matchedEntity
        ? `${mapping.matchedEntityType === "company" ? "Compañía" : "Lead"} · ${mapping.matchedEntity.title}`
        : "-";
      const statusLabel =
        mapping.status === "mapped"
          ? "Vinculado"
          : mapping.status === "missing"
            ? "No encontrado"
            : mapping.status === "scope_missing"
              ? "Sin permiso CRM"
              : "Sin ID";
      row.innerHTML = `
        <td><b>${mapping.clientName}</b></td>
        <td>${mapping.bitrixCompanyId || "-"}</td>
        <td>${mapping.bitrixLeadId || "-"}</td>
        <td>${matchedLabel}</td>
        <td>${statusLabel}</td>
      `;
      bitrixClientMappingsTableBody.appendChild(row);
    });
  }

  renderBitrixTaskResponsibleOptions();
}

function renderVisitsResponsibleFilters() {
  const currentExecutive = visitsExecutiveFilter.value;
  const currentSupervisor = visitsSupervisorFilter.value;
  const currentParticipant = visitsParticipantFilter.value;
  const currentCalendarParticipant = calendarParticipantFilter.value;
  const currentStatsExecutive = statsExecutiveFilter.value;
  visitsExecutiveFilter.innerHTML = ['<option value="todos">Todos</option>']
    .concat(
      assignmentOptions.executives.map((user) => `<option value="${user.id}">${user.name}</option>`)
    )
    .join("");

  const supervisors = [
    ...assignmentOptions.supervisors.fixedFire,
    ...assignmentOptions.supervisors.extinguishers,
    ...assignmentOptions.supervisors.works
  ].filter((user, index, array) => array.findIndex((candidate) => Number(candidate.id) === Number(user.id)) === index);

  visitsSupervisorFilter.innerHTML = ['<option value="todos">Todos</option>']
    .concat(supervisors.map((user) => `<option value="${user.id}">${user.name}</option>`))
    .join("");

  visitsParticipantFilter.innerHTML = ['<option value="todos">Todos</option>']
    .concat(getParticipantUsers().map((user) => `<option value="${user.id}">${user.name}</option>`))
    .join("");

  calendarParticipantFilter.innerHTML = ['<option value="todos">Todos</option>']
    .concat(getParticipantUsers().map((user) => `<option value="${user.id}">${user.name}</option>`))
    .join("");

  statsExecutiveFilter.innerHTML = ['<option value="todos">Todos</option>']
    .concat(assignmentOptions.executives.map((user) => `<option value="${user.id}">${user.name}</option>`))
    .join("");

  visitsExecutiveFilter.value = currentExecutive || "todos";
  visitsSupervisorFilter.value = currentSupervisor || "todos";
  visitsParticipantFilter.value = currentParticipant || "todos";
  calendarParticipantFilter.value = currentCalendarParticipant || "todos";
  statsExecutiveFilter.value = currentStatsExecutive || "todos";
  renderOpportunityOwnerSelects();
}

function syncMeetingCompletionFields() {
  const isCompleted = meetingStatusSelect.value === "Realizada";
  meetingNextDateGroup.classList.toggle("hidden", !isCompleted);
  meetingNextDate.required = false;
  if (!isCompleted) {
    meetingNextDate.value = "";
  }
}

function renderServices(company, compact = false) {
  return SERVICE_TYPES.filter((service) => company.services[service.key])
    .map(
      (service) =>
        `<span class="service-badge ${service.css}">${compact ? service.short : service.label}</span>`
    )
    .join("");
}

function renderMeetingDots(dots) {
  if (!dots?.length) return '<span class="muted-inline">-</span>';
  return dots
    .map(
      (dot, index) =>
        `<span class="meeting-dot ${dot.color}" title="${dot.kind} ${index + 1}"></span>`
    )
    .join("");
}

function renderUserRoleOptions(selectedRole) {
  return userRoles
    .map((role) => `<option value="${role}" ${role === selectedRole ? "selected" : ""}>${role}</option>`)
    .join("");
}

function renderUserSelect(select, options, selectedId, placeholder) {
  const normalizedSelectedId = selectedId ? Number(selectedId) : null;
  select.innerHTML = [`<option value="">${placeholder}</option>`]
    .concat(
      options.map(
        (user) =>
          `<option value="${user.id}" ${normalizedSelectedId === Number(user.id) ? "selected" : ""}>${user.name}</option>`
      )
    )
    .join("");
}

function renderSectorSelect(selectedSector = "") {
  editSector.innerHTML = [
    '<option value="">Seleccionar sector</option>',
    ...sectorOptions.map(
      (sector) =>
        `<option value="${sector.name}" ${sector.name === selectedSector ? "selected" : ""}>${sector.name}</option>`
    )
  ].join("");
}

function createEmptyVisitRule() {
  return {
    periodicityDays: 30,
    contactRole: contactRoles[0]?.name || "",
    objective: VISIT_RULE_OBJECTIVE_OPTIONS[0]
  };
}

function renderVisitRulesEditor() {
  if (!visitRulesFeatureEnabled) {
    visitRulesList.innerHTML = "";
    return;
  }

  visitRulesList.innerHTML = "";

  if (!currentVisitRules.length) {
    const empty = document.createElement("div");
    empty.className = "muted-inline";
    empty.textContent = "Todavía no hay reglas definidas.";
    visitRulesList.appendChild(empty);
    return;
  }

  currentVisitRules.forEach((rule, index) => {
    const item = document.createElement("div");
    item.className = "visit-rule-item";
    item.innerHTML = `
      <div class="edit-grid visit-rule-grid">
        <label>
          Periodicidad (días)
          <input class="visit-rule-periodicity" type="number" min="1" step="1" value="${rule.periodicityDays || 30}" />
        </label>
        <label>
          A quién ir a ver
          <select class="visit-rule-contact-role">
            <option value="">Seleccionar función</option>
            ${contactRoles
              .map(
                (contactRole) =>
                  `<option value="${contactRole.name}" ${contactRole.name === rule.contactRole ? "selected" : ""}>${contactRole.name}</option>`
              )
              .join("")}
          </select>
        </label>
        <label class="span-2">
          Objetivo
          <select class="visit-rule-objective">
            ${VISIT_RULE_OBJECTIVE_OPTIONS.map(
              (objectiveOption) =>
                `<option value="${objectiveOption}" ${objectiveOption === rule.objective ? "selected" : ""}>${objectiveOption}</option>`
            ).join("")}
            ${
              rule.objective && !VISIT_RULE_OBJECTIVE_OPTIONS.includes(rule.objective)
                ? `<option value="${rule.objective}" selected>${rule.objective}</option>`
                : ""
            }
          </select>
        </label>
      </div>
      <div class="table-actions">
        <button class="ghost-btn remove-visit-rule-btn" type="button">Eliminar regla</button>
      </div>
    `;

    item.querySelector(".visit-rule-periodicity").addEventListener("input", (event) => {
      currentVisitRules[index].periodicityDays = Number(event.target.value || 0);
    });
    item.querySelector(".visit-rule-contact-role").addEventListener("change", (event) => {
      currentVisitRules[index].contactRole = event.target.value;
    });
    item.querySelector(".visit-rule-objective").addEventListener("change", (event) => {
      currentVisitRules[index].objective = event.target.value;
    });
    item.querySelector(".remove-visit-rule-btn").addEventListener("click", () => {
      currentVisitRules.splice(index, 1);
      renderVisitRulesEditor();
    });

    visitRulesList.appendChild(item);
  });
}

function renderSectors() {
  sectorsTableBody.innerHTML = "";

  sectorOptions.forEach((sector) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td><b>${sector.name}</b></td>
      <td>
        <div class="table-actions">
          <button class="ghost-btn delete-sector-btn" type="button">Eliminar</button>
        </div>
      </td>
    `;

    row.querySelector(".delete-sector-btn").addEventListener("click", async () => {
      const confirmed = window.confirm(`Se va a eliminar el sector ${sector.name}.`);
      if (!confirmed) return;

      try {
        await deleteSector(sector.id);
        sectorStatus.textContent = "Sector eliminado";
        await loadSectorOptions();
      } catch (error) {
        notifyError(error.message);
      }
    });

    sectorsTableBody.appendChild(row);
  });
}

function syncSupervisorVisibility() {
  fixedFireSupervisorGroup.classList.toggle("hidden", !editFixedFire.checked);
  extinguishersSupervisorGroup.classList.toggle("hidden", !editExtinguishers.checked);
  worksSupervisorGroup.classList.toggle("hidden", !editWorks.checked);
  editSupervisorFixedFire.disabled = !editFixedFire.checked;
  editSupervisorExtinguishers.disabled = !editExtinguishers.checked;
  editSupervisorWorks.disabled = !editWorks.checked;
  editSupervisorFixedFire.required = editFixedFire.checked;
  editSupervisorExtinguishers.required = editExtinguishers.checked;
  editSupervisorWorks.required = editWorks.checked;
}

function showScreen(screen) {
  const requestedScreen = !CRM_ENABLED && screen === "pipeline" ? "list" : screen;
  const nextScreen = !canAccessSettings() && (requestedScreen === "settings" || requestedScreen === "users") ? "list" : requestedScreen;
  listScreen.classList.toggle("hidden", nextScreen !== "list");
  visitsScreen.classList.toggle("hidden", nextScreen !== "visits");
  pipelineScreen.classList.toggle("hidden", nextScreen !== "pipeline");
  detailScreen.classList.toggle("hidden", !["detail", "edit", "meeting"].includes(nextScreen));
  editScreen.classList.toggle("hidden", nextScreen !== "edit");
  meetingScreen.classList.toggle("hidden", nextScreen !== "meeting");
  calendarScreen.classList.toggle("hidden", nextScreen !== "calendar");
  visitsGridScreen.classList.toggle("hidden", nextScreen !== "visits-grid");
  visitsStatsScreen.classList.toggle("hidden", nextScreen !== "visits-stats");
  usersScreen.classList.toggle("hidden", nextScreen !== "users");
  settingsScreen.classList.toggle("hidden", nextScreen !== "settings");
  showClientsBtn.classList.toggle(
    "active-view",
    nextScreen === "list" || nextScreen === "detail" || nextScreen === "edit" || nextScreen === "meeting"
  );
  showVisitsBtn.classList.toggle(
    "active-view",
    nextScreen === "visits" || nextScreen === "visits-grid" || nextScreen === "calendar" || nextScreen === "visits-stats"
  );
  showPipelineBtn.classList.toggle("active-view", CRM_ENABLED && nextScreen === "pipeline");
  showVisitsGridBtn.classList.toggle("active-view", nextScreen === "visits-grid");
  showCalendarBtn.classList.toggle("active-view", nextScreen === "calendar");
  showVisitsStatsBtn.classList.toggle("active-view", nextScreen === "visits-stats");
  showSettingsBtn.classList.toggle("active-view", nextScreen === "settings" || nextScreen === "users");
  syncModalState();
}

function showSettingsTab(tab) {
  currentSettingsTab = ["catalogs", "imports", "control"].includes(tab) ? tab : "catalogs";
  settingsPaneCatalogs.classList.toggle("hidden", currentSettingsTab !== "catalogs");
  settingsPaneImports.classList.toggle("hidden", currentSettingsTab !== "imports");
  settingsPaneControl.classList.toggle("hidden", currentSettingsTab !== "control");
  settingsTabButtons.forEach((button) => {
    button.classList.toggle("active-view", button.dataset.settingsTab === currentSettingsTab);
  });
}

function buildCompanyHash(clientId, branchId = null) {
  if (!clientId) return "#/clientes";
  return branchId ? `#/companias/${clientId}/sucursales/${branchId}` : `#/companias/${clientId}`;
}

function replaceHash(hash) {
  if (window.location.hash === hash) return;
  suppressHashRouting = true;
  window.location.hash = hash;
  window.setTimeout(() => {
    suppressHashRouting = false;
  }, 0);
}

function syncUrlWithState(screen) {
  if (!currentUser) return;
  if (screen === "detail" && selectedClient) {
    replaceHash(buildCompanyHash(selectedClient.id, selectedBranchView?.id || null));
    return;
  }
  if (screen === "calendar") {
    replaceHash("#/calendario");
    return;
  }
  if (screen === "visits-stats") {
    replaceHash("#/visitas/estadisticas");
    return;
  }
  if (screen === "visitas" || screen === "visits") {
    replaceHash("#/visitas");
    return;
  }
  if (screen === "pipeline") {
    if (!CRM_ENABLED) {
      replaceHash("#/clientes");
      return;
    }
    replaceHash("#/pipeline");
    return;
  }
  if (screen === "users") {
    if (!canAccessSettings()) {
      replaceHash("#/clientes");
      return;
    }
    replaceHash("#/configuracion/usuarios");
    return;
  }
  if (screen === "settings") {
    if (!canAccessSettings()) {
      replaceHash("#/clientes");
      return;
    }
    replaceHash("#/configuracion");
    return;
  }
  replaceHash("#/clientes");
}

async function openClientDetail(clientId, { branchId = null, smooth = false } = {}) {
  selectedId = Number(clientId);
  const loaded = await loadClientDetail(selectedId);
  if (!loaded) {
    notifyError("No pudimos abrir la ficha del cliente. Revisá la consola del navegador si vuelve a pasar.");
    return false;
  }

  selectedBranchView = branchId
    ? selectedClient?.branches?.find((branch) => Number(branch.id) === Number(branchId)) || null
    : null;
  renderTable();
  showScreen("detail");
  syncUrlWithState("detail");

  if (smooth) {
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
  return true;
}

function parseAppHash() {
  const hash = window.location.hash || "#/clientes";
  const normalized = hash.startsWith("#") ? hash.slice(1) : hash;
  const companyBranchMatch = normalized.match(/^\/companias\/(\d+)\/sucursales\/(\d+)$/);
  if (companyBranchMatch) {
    return {
      screen: "detail",
      clientId: Number(companyBranchMatch[1]),
      branchId: Number(companyBranchMatch[2])
    };
  }

  const companyMatch = normalized.match(/^\/companias\/(\d+)$/);
  if (companyMatch) {
    return {
      screen: "detail",
      clientId: Number(companyMatch[1]),
      branchId: null
    };
  }

  if (normalized === "/calendario") return { screen: "calendar" };
  if (normalized === "/visitas/estadisticas") return { screen: "visits-stats" };
  if (normalized === "/visitas") return { screen: "visits" };
  if (normalized === "/pipeline") return { screen: "pipeline" };
  if (normalized === "/configuracion/usuarios") return { screen: "users" };
  if (normalized === "/configuracion") return { screen: "settings" };
  return { screen: "list" };
}

async function applyRouteFromHash() {
  if (!currentUser) return;

  const route = parseAppHash();
  if (route.screen === "detail") {
    await openClientDetail(route.clientId, { branchId: route.branchId });
    return;
  }

  if (route.screen === "calendar") {
    calendarView = "month";
    selectedCalendarDay = null;
    await loadCalendar();
    showScreen("calendar");
    return;
  }

  if (route.screen === "visits") {
    await loadVisits();
    showScreen("visits");
    return;
  }

  if (route.screen === "visits-stats") {
    await loadVisitsStats();
    showScreen("visits-stats");
    return;
  }

  if (route.screen === "pipeline") {
    if (!CRM_ENABLED) {
      showScreen("list");
      replaceHash("#/clientes");
      return;
    }
    await loadPipeline();
    showScreen("pipeline");
    return;
  }

  if (route.screen === "users") {
    if (!canAccessSettings()) {
      showScreen("list");
      replaceHash("#/clientes");
      return;
    }
    await loadUsers();
    showScreen("users");
    return;
  }

  if (route.screen === "settings") {
    if (!canAccessSettings()) {
      showScreen("list");
      replaceHash("#/clientes");
      return;
    }
    await loadSettingsCatalogs();
    showSettingsTab("catalogs");
    showScreen("settings");
    return;
  }

  selectedBranchView = null;
  showScreen("list");
}

function computeGlobalKpis(globalData) {
  globalKpis.innerHTML = "";
  const items = [
    { label: "Reuniones agendadas", value: globalData.meetingsScheduled || 0 },
    { label: "Reuniones confirmadas", value: globalData.meetingsConfirmed || 0 },
    { label: "Reuniones realizadas", value: globalData.meetingsCompleted || 0 },
    { label: "Clientes visitados", value: globalData.clientsVisited || 0 }
  ];

  if (CRM_ENABLED) {
    items.push(
      { label: "Oportunidades abiertas", value: globalData.openOpportunities || 0 },
      { label: "Follow-ups vencidos", value: globalData.overdueFollowUps || 0 }
    );
  }

  items.forEach((kpi) => {
    const chip = document.createElement("div");
    chip.className = "kpi-chip";
    chip.innerHTML = `<b>${kpi.value}</b><span>${kpi.label}</span>`;
    globalKpis.appendChild(chip);
  });
}

function renderTable() {
  visibleCount.textContent = `${clients.length} clientes`;
  tableBody.innerHTML = "";

  clients.forEach((company) => {
    const tr = document.createElement("tr");
    if (company.id === selectedId) tr.classList.add("active");

    tr.innerHTML = `
      <td><b>${company.name}</b><br /><span class="muted-inline">${company.companyType} · ${company.country} · ${company.accountStage}</span></td>
      <td>${company.sector}</td>
      <td>${company.executiveName || company.manager}</td>
      <td><span class="risk ${company.risk.toLowerCase()}">${company.risk}</span></td>
      <td><div class="service-row compact">${renderServices(company, true)}</div></td>
      <td><div class="meeting-dot-row">${renderMeetingDots(company.meetingDots)}</div></td>
      <td>${company.nextMeeting}</td>
    `;

    tr.addEventListener("click", async () => {
      await openClientDetail(company.id, { smooth: true });
    });

    tableBody.appendChild(tr);
  });
}

function renderVisitsTable() {
  const sortedVisits = [...visits].sort((left, right) => {
    const getComparable = (visit) => {
      if (visitsSort.key === "scheduledFor") return visit.scheduledFor || "";
      if (visitsSort.key === "scopeLabel") return visit.scopeLabel || "";
      if (visitsSort.key === "kindLabel") return visit.kindLabel || visit.kind || "";
      if (visitsSort.key === "contactName") return [visit.contactName, visit.contactRole].filter(Boolean).join(" · ");
      return String(visit[visitsSort.key] || "");
    };

    const leftValue = getComparable(left);
    const rightValue = getComparable(right);
    const comparison =
      typeof leftValue === "string" && typeof rightValue === "string"
        ? leftValue.localeCompare(rightValue, "es", { numeric: true, sensitivity: "base" })
        : String(leftValue).localeCompare(String(rightValue), "es", { numeric: true, sensitivity: "base" });

    return visitsSort.direction === "asc" ? comparison : comparison * -1;
  });

  visitsVisibleCount.textContent = `${sortedVisits.length} visitas`;
  visitsTableBody.innerHTML = "";

  document.querySelectorAll(".table-sort-btn").forEach((button) => {
    const isActive = button.dataset.sortKey === visitsSort.key;
    button.classList.toggle("active", isActive);
    button.dataset.direction = isActive ? visitsSort.direction : "";
  });

  if (!sortedVisits.length) {
    visitsTableBody.innerHTML = `
      <tr>
        <td colspan="7" class="empty-table-cell">No encontramos visitas con esos filtros.</td>
      </tr>
    `;
    return;
  }

  sortedVisits.forEach((visit) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td><b>${formatDate(visit.scheduledFor)}</b></td>
      <td><b>${visit.clientName}</b></td>
      <td>${visit.branchName || "Casa matriz"}</td>
      <td>${visit.kindLabel || visit.kind}</td>
      <td>${visit.modality || "-"}</td>
      <td><span class="visit-status-badge ${String(visit.status || "").toLowerCase()}">${visit.status}</span></td>
      <td>${[visit.contactName, visit.contactRole].filter(Boolean).join(" · ") || "-"}</td>
    `;

    row.addEventListener("click", async () => {
      await openClientDetail(visit.clientId, { branchId: visit.branchId || null, smooth: true });
    });

    visitsTableBody.appendChild(row);
  });
}

function formatGridDate(dateValue) {
  if (!dateValue) return "";
  const [year, month, day] = dateValue.split("-");
  if (!year || !month || !day) return dateValue;
  return `${Number(day)}/${Number(month)}/${String(year).slice(2)}`;
}

function renderVisitsGridLegend() {
  visitsGridLegend.innerHTML = meetingTypes
    .map((type) => `<span class="meeting-type-badge ${type.color}">${type.label}</span>`)
    .join("");
}

function renderCalendarLegend() {
  calendarLegend.innerHTML = meetingTypes
    .map((type) => `<span class="meeting-type-badge ${type.color}">${type.label}</span>`)
    .join("");
}

function renderVisitsGrid() {
  visitsGridWrap.innerHTML = "";
  renderVisitsGridLegend();

  if (!visitsGridData.length) {
    visitsGridWrap.innerHTML = `
      <article class="empty-meeting-state">
        <h4>No hay visitas realizadas</h4>
        <p>Cuando existan reuniones marcadas como realizadas, van a aparecer en esta grilla.</p>
      </article>
    `;
    return;
  }

  const groupedByClient = visitsGridData.reduce((accumulator, visit) => {
    const key = `${visit.clientId}`;
    if (!accumulator.has(key)) {
      accumulator.set(key, {
        clientId: visit.clientId,
        clientName: visit.clientName,
        visits: []
      });
    }
    accumulator.get(key).visits.push(visit);
    return accumulator;
  }, new Map());

  const rows = [...groupedByClient.values()]
    .map((row) => ({
      ...row,
      visits: row.visits.sort((left, right) => left.scheduledFor.localeCompare(right.scheduledFor))
    }))
    .sort((left, right) => left.clientName.localeCompare(right.clientName, "es"));

  const maxColumns = Math.max(...rows.map((row) => row.visits.length), 1);
  const table = document.createElement("table");
  table.className = "visits-grid-table";

  const headerCells = ['<th class="visits-grid-corner" colspan="1">Cliente</th>']
    .concat(Array.from({ length: maxColumns }, () => "<th></th>"))
    .join("");

  table.innerHTML = `
    <thead>
      <tr>${headerCells}</tr>
    </thead>
    <tbody></tbody>
  `;

  const tbody = table.querySelector("tbody");

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    const cells = [`<th class="visits-grid-client" title="${row.clientName}"><span>${row.clientName}</span></th>`];

    for (let index = 0; index < maxColumns; index += 1) {
      const visit = row.visits[index];
      if (!visit) {
        cells.push('<td class="visits-grid-empty"></td>');
        continue;
      }

      cells.push(`
        <td>
          <button type="button" class="visits-grid-cell ${visit.color}" data-client-id="${visit.clientId}" data-meeting-id="${visit.id}">
            ${formatGridDate(visit.scheduledFor)}
          </button>
        </td>
      `);
    }

    tr.innerHTML = cells.join("");
    tr.querySelectorAll(".visits-grid-cell").forEach((button) => {
      button.addEventListener("click", async () => {
        const visit = row.visits.find((item) => Number(item.id) === Number(button.dataset.meetingId));
        if (!visit) return;
        await openClientDetail(visit.clientId, { branchId: visit.branchId || null, smooth: true });
        openMeetingDetailScreen(visit);
      });
    });
    tbody.appendChild(tr);
  });

  visitsGridWrap.appendChild(table);
}

function getVisitsFilters() {
  return {
    search: visitsSearchInput.value.trim(),
    status: visitsStatusFilter.value,
    kind: visitsTypeFilter.value,
    modality: visitsModalityFilter.value,
    executiveUserId: visitsExecutiveFilter.value,
    supervisorUserId: visitsSupervisorFilter.value,
    participantUserId: visitsParticipantFilter.value,
    dateFrom: visitsDateFromFilter.value,
    dateTo: visitsDateToFilter.value
  };
}

function getCalendarFilters() {
  return {
    participantUserId: calendarParticipantFilter.value
  };
}

function getVisitsStatsFilters() {
  return {
    executiveUserId: statsExecutiveFilter.value,
    dateFrom: statsDateFrom.value,
    dateTo: statsDateTo.value
  };
}

function renderVisitStats() {
  statsUsersTableBody.innerHTML = "";
  statsTypesTableBody.innerHTML = "";
  applyFeatureFlagsUi();
  statsRulesTotal.textContent = String(visitStats.byRule.totalCount || 0);
  statsRulesCompleted.textContent = String(visitStats.byRule.completedCount || 0);
  statsRulesCompletionRate.textContent = `${Number(visitStats.byRule.completionRate || 0).toFixed(0)}%`;
  statsRulesStatus.textContent =
    visitStats.byRule.totalCount > 0 ? `${visitStats.byRule.completedCount} de ${visitStats.byRule.totalCount}` : "Sin visitas por regla";
  const semaphores = visitStats.byRule.semaphores || {
    totalRules: 0,
    whiteCount: 0,
    greenCount: 0,
    yellowCount: 0,
    redCount: 0,
    whiteRate: 0,
    greenRate: 0,
    yellowRate: 0,
    redRate: 0
  };
  statsSemaphoreWhiteRate.textContent = `${Number(semaphores.whiteRate || 0).toFixed(0)}%`;
  statsSemaphoreGreenRate.textContent = `${Number(semaphores.greenRate || 0).toFixed(0)}%`;
  statsSemaphoreYellowRate.textContent = `${Number(semaphores.yellowRate || 0).toFixed(0)}%`;
  statsSemaphoreRedRate.textContent = `${Number(semaphores.redRate || 0).toFixed(0)}%`;
  statsSemaphoreWhiteCount.textContent = `${semaphores.whiteCount || 0} reglas`;
  statsSemaphoreGreenCount.textContent = `${semaphores.greenCount || 0} reglas`;
  statsSemaphoreYellowCount.textContent = `${semaphores.yellowCount || 0} reglas`;
  statsSemaphoreRedCount.textContent = `${semaphores.redCount || 0} reglas`;
  statsSemaphoreStatus.textContent = semaphores.totalRules > 0 ? `${semaphores.totalRules} reglas activas` : "Sin reglas activas";

  if (!visitStats.byUser.length) {
    const row = document.createElement("tr");
    row.innerHTML = '<td colspan="6" class="muted-inline">No hay visitas en el período seleccionado.</td>';
    statsUsersTableBody.appendChild(row);
  } else {
    visitStats.byUser.forEach((item) => {
      const row = document.createElement("tr");
      row.innerHTML = `
        <td><b>${item.name}</b></td>
        <td>${item.role}</td>
        <td>${item.scheduledCount}</td>
        <td>${item.confirmedCount}</td>
        <td>${item.completedCount}</td>
        <td><b>${item.totalCount}</b></td>
      `;
      statsUsersTableBody.appendChild(row);
    });
  }

  if (!visitStats.byType.length) {
    const row = document.createElement("tr");
    row.innerHTML = '<td colspan="5" class="muted-inline">No hay tipos con visitas en el período seleccionado.</td>';
    statsTypesTableBody.appendChild(row);
  } else {
    visitStats.byType.forEach((item) => {
      const row = document.createElement("tr");
      row.innerHTML = `
        <td><span class="meeting-kind-badge ${item.color}">${item.label}</span></td>
        <td>${item.scheduledCount}</td>
        <td>${item.confirmedCount}</td>
        <td>${item.completedCount}</td>
        <td><b>${item.totalCount}</b></td>
      `;
      statsTypesTableBody.appendChild(row);
    });
  }

  statsUsersStatus.textContent = `${visitStats.byUser.length} usuarios`;
  statsTypesStatus.textContent = `${visitStats.byType.length} tipos`;
}

function kpiItem(label, value) {
  const div = document.createElement("div");
  div.className = "kpi-box";
  div.innerHTML = `<span>${label}</span><b>${value}</b>`;
  return div;
}

function renderRoleAssignmentsCard(label, value) {
  const div = document.createElement("div");
  div.className = "kpi-box";
  div.innerHTML = `<span>${label}</span><b>${value || "Sin asignar"}</b>`;
  return div;
}

function buildCrmSummaryClientSide(opportunities) {
  const openOpportunities = opportunities.filter((opportunity) => isOpenOpportunity(opportunity.status));
  const today = new Date(`${dateKey(new Date())}T00:00:00`);
  const dueThisWeek = openOpportunities.reduce((total, opportunity) => {
    const dueDate = opportunity.nextFollowUp?.dueDate;
    if (!dueDate || opportunity.nextFollowUp?.visualStatus === "Vencido") return total;
    const diffDays = Math.floor((new Date(`${dueDate}T00:00:00`) - today) / 86400000);
    return diffDays >= 0 && diffDays <= 7 ? total + 1 : total;
  }, 0);

  return {
    openCount: openOpportunities.length,
    totalAmount: openOpportunities.reduce((total, opportunity) => total + Number(opportunity.amount || 0), 0),
    weightedAmount: openOpportunities.reduce((total, opportunity) => total + Number(opportunity.weightedAmount || 0), 0),
    overdueFollowUps: openOpportunities.reduce((total, opportunity) => total + Number(opportunity.overdueFollowUps || 0), 0),
    dueThisWeek
  };
}

function getVisibleOpportunities() {
  const opportunities = Array.isArray(selectedClient?.opportunities) ? selectedClient.opportunities : [];
  if (!selectedBranchView) return opportunities;
  return opportunities.filter((opportunity) => Number(opportunity.branchId) === Number(selectedBranchView.id));
}

function renderCrmSummary(opportunities) {
  const summary = buildCrmSummaryClientSide(opportunities);
  crmSummaryGrid.innerHTML = "";
  crmSummaryGrid.append(
    kpiItem("Oportunidades abiertas", summary.openCount),
    kpiItem("Pipeline abierto", formatMoney(summary.totalAmount)),
    kpiItem("Pipeline ponderado", formatMoney(summary.weightedAmount)),
    kpiItem("Seguimientos vencidos", summary.overdueFollowUps),
    kpiItem("Seguimientos esta semana", summary.dueThisWeek)
  );
}

function renderDetailSectionNav() {
  if (!CRM_ENABLED) {
    detailOpportunitiesSection.classList.add("hidden");
    showDetailOpportunitiesBtn.classList.add("hidden");
    detailVisitsSection.classList.remove("hidden");
    showDetailVisitsBtn.classList.add("active-view");
    return;
  }
  const showingOpportunities = detailSection === "opportunities";
  detailOpportunitiesSection.classList.toggle("hidden", !showingOpportunities);
  detailVisitsSection.classList.toggle("hidden", showingOpportunities);
  showDetailOpportunitiesBtn.classList.toggle("active-view", showingOpportunities);
  showDetailVisitsBtn.classList.toggle("active-view", !showingOpportunities);
}

function resetOpportunityForm() {
  renderOpportunitySelects();
  const opportunityTypes = crmCatalogs.opportunityTypes?.length ? crmCatalogs.opportunityTypes : DEFAULT_OPPORTUNITY_TYPES;
  const serviceLines = crmCatalogs.serviceLines?.length ? crmCatalogs.serviceLines : DEFAULT_SERVICE_LINES;
  editingOpportunityId = null;
  opportunityFormTitle.textContent = "Nueva oportunidad";
  saveOpportunityBtn.textContent = "Guardar oportunidad";
  opportunityTitle.value = "";
  opportunityAmount.value = "";
  opportunityProbability.value = "50";
  opportunityExpectedCloseDate.value = "";
  opportunitySource.value = "";
  opportunityDescription.value = "";
  opportunityLossReason.value = "";
  opportunityStatusMessage.textContent = "";
  opportunityStatusSelect.value = opportunityStatusSelect.options[0]?.value || "";
  opportunityType.value = opportunityTypes[0] || "";
  opportunityServiceLine.value = serviceLines[0] || "";
  opportunityOwnerUserId.value = currentUser?.id ? String(currentUser.id) : "";
  renderOpportunityBranchOptions();
  syncOpportunityLossReasonVisibility();
  opportunityForm.classList.add("hidden");
  syncModalState();
}

function openOpportunityForm(opportunity = null) {
  if (!selectedClient) return;
  renderOpportunitySelects();
  const opportunityTypes = crmCatalogs.opportunityTypes?.length ? crmCatalogs.opportunityTypes : DEFAULT_OPPORTUNITY_TYPES;
  const serviceLines = crmCatalogs.serviceLines?.length ? crmCatalogs.serviceLines : DEFAULT_SERVICE_LINES;
  editingOpportunityId = opportunity?.id || null;
  opportunityFormTitle.textContent = opportunity ? `Editar oportunidad · ${opportunity.title}` : "Nueva oportunidad";
  saveOpportunityBtn.textContent = opportunity ? "Guardar cambios" : "Guardar oportunidad";
  opportunityTitle.value = opportunity?.title || "";
  opportunityType.value = opportunity?.opportunityType || opportunityTypes[0] || "";
  opportunityServiceLine.value = opportunity?.serviceLine || serviceLines[0] || "";
  opportunityStatusSelect.value = opportunity?.status || opportunityStatusSelect.options[0]?.value || "";
  opportunityAmount.value = opportunity?.amount ?? "";
  opportunityProbability.value = opportunity?.probability ?? 50;
  opportunityExpectedCloseDate.value = opportunity?.expectedCloseDate || "";
  opportunityOwnerUserId.value = opportunity?.ownerUserId ? String(opportunity.ownerUserId) : currentUser?.id ? String(currentUser.id) : "";
  opportunitySource.value = opportunity?.source || "";
  opportunityDescription.value = opportunity?.description || "";
  opportunityLossReason.value = opportunity?.lossReason || "";
  renderOpportunityBranchOptions(opportunity?.branchId || (selectedBranchView ? selectedBranchView.id : ""));
  opportunityStatusMessage.textContent = "";
  syncOpportunityLossReasonVisibility();
  opportunityForm.classList.remove("hidden");
  followUpForm.classList.add("hidden");
  syncModalState();
}

function resetFollowUpForm() {
  renderOpportunitySelects();
  editingFollowUpId = null;
  followUpOpportunityId = null;
  followUpFormTitle.textContent = "Nuevo seguimiento";
  saveFollowUpBtn.textContent = "Guardar seguimiento";
  followUpType.value = followUpType.options[0]?.value || "";
  followUpTitle.value = "";
  followUpDueDate.value = "";
  followUpAssignedUserId.value = currentUser?.id ? String(currentUser.id) : "";
  followUpStatusSelect.value = followUpStatusSelect.options[0]?.value || "";
  followUpNotes.value = "";
  followUpStatusMessage.textContent = "";
  followUpForm.classList.add("hidden");
  syncModalState();
}

function openFollowUpForm(opportunity, followUp = null) {
  if (!opportunity) return;
  renderOpportunitySelects();
  editingFollowUpId = followUp?.id || null;
  followUpOpportunityId = opportunity.id;
  followUpFormTitle.textContent = followUp ? `Editar seguimiento · ${opportunity.title}` : `Nuevo seguimiento · ${opportunity.title}`;
  saveFollowUpBtn.textContent = followUp ? "Guardar cambios" : "Guardar seguimiento";
  followUpType.value = followUp?.type || followUpType.options[0]?.value || "";
  followUpTitle.value = followUp?.title || "";
  followUpDueDate.value = followUp?.dueDate || "";
  followUpAssignedUserId.value = followUp?.assignedUserId ? String(followUp.assignedUserId) : currentUser?.id ? String(currentUser.id) : "";
  followUpStatusSelect.value = followUp?.status || followUpStatusSelect.options[0]?.value || "";
  followUpNotes.value = followUp?.notes || "";
  followUpStatusMessage.textContent = "";
  followUpForm.classList.remove("hidden");
  opportunityForm.classList.add("hidden");
  syncModalState();
}

function renderFollowUpsList(opportunity) {
  if (!opportunity.followUps?.length) {
    return `
      <div class="crm-empty-inline">
        <strong>Sin seguimientos cargados</strong>
        <span>Definí el próximo paso para que la oportunidad no quede sin dueño ni fecha.</span>
      </div>
    `;
  }

  return opportunity.followUps
    .map(
      (followUp) => `
        <article class="follow-up-item">
          <div class="follow-up-item-main">
            <div class="follow-up-item-head">
              <strong>${followUp.title}</strong>
              <span class="follow-up-badge ${getFollowUpStatusClass(followUp.visualStatus || followUp.status)}">${followUp.visualStatus || followUp.status}</span>
            </div>
            <p class="follow-up-item-meta">${[
              followUp.type,
              followUp.dueDate ? `Para ${formatDate(followUp.dueDate)}` : "",
              followUp.assignedUserName ? `Responsable: ${followUp.assignedUserName}` : ""
            ]
              .filter(Boolean)
              .join(" · ")}</p>
            <p class="follow-up-item-notes">${followUp.notes || "Sin comentarios cargados."}</p>
          </div>
          <div class="follow-up-item-actions">
            <button class="ghost-btn edit-follow-up-btn" type="button" data-follow-up-id="${followUp.id}">Editar</button>
            <button class="ghost-btn toggle-follow-up-btn" type="button" data-follow-up-id="${followUp.id}" data-next-status="${followUp.status === "Hecho" ? "Pendiente" : "Hecho"}">${followUp.status === "Hecho" ? "Reabrir" : "Marcar hecho"}</button>
          </div>
        </article>
      `
    )
    .join("");
}

function renderOpportunityCard(opportunity) {
  const article = document.createElement("article");
  article.className = "opportunity-card";
  article.innerHTML = `
    <div class="opportunity-card-head">
      <div>
        <div class="opportunity-card-badges">
          <span class="opportunity-badge ${getOpportunityStatusClass(opportunity.status)}">${opportunity.status}</span>
          <span class="pill">${opportunity.opportunityType}</span>
          <span class="service-badge">${opportunity.serviceLine}</span>
          ${opportunity.branchName ? `<span class="pill">${opportunity.branchName}</span>` : '<span class="pill">Casa matriz</span>'}
        </div>
        <h4 class="opportunity-card-title">${opportunity.title}</h4>
        <p class="opportunity-card-meta">${[
          opportunity.clientName || selectedClient?.name || "",
          opportunity.ownerName ? `Responsable: ${opportunity.ownerName}` : "",
          opportunity.source ? `Origen: ${opportunity.source}` : ""
        ]
          .filter(Boolean)
          .join(" · ")}</p>
      </div>
      <div class="opportunity-card-actions">
        <button class="secondary-btn edit-opportunity-btn" type="button">Editar</button>
        <button class="ghost-btn add-follow-up-btn" type="button">+ Seguimiento</button>
      </div>
    </div>
    <div class="opportunity-kpis">
      <div class="opportunity-kpi">
        <span>Monto</span>
        <b>${formatMoney(opportunity.amount)}</b>
      </div>
      <div class="opportunity-kpi">
        <span>Probabilidad</span>
        <b>${opportunity.probability}%</b>
      </div>
      <div class="opportunity-kpi">
        <span>Ponderado</span>
        <b>${formatMoney(opportunity.weightedAmount)}</b>
      </div>
      <div class="opportunity-kpi">
        <span>Cierre estimado</span>
        <b>${opportunity.expectedCloseDate ? formatDate(opportunity.expectedCloseDate) : "Sin fecha"}</b>
      </div>
      <div class="opportunity-kpi">
        <span>Próximo paso</span>
        <b>${opportunity.nextFollowUp ? `${opportunity.nextFollowUp.title} · ${formatDate(opportunity.nextFollowUp.dueDate)}` : "Sin seguimiento"}</b>
      </div>
      <div class="opportunity-kpi">
        <span>Alertas</span>
        <b>${opportunity.overdueFollowUps ? `${opportunity.overdueFollowUps} vencido(s)` : "Al día"}</b>
      </div>
    </div>
    <div class="opportunity-description">
      <p>${opportunity.description || "Sin comentarios comerciales cargados."}</p>
      ${opportunity.lossReason ? `<p class="opportunity-loss-reason"><strong>Motivo de pérdida:</strong> ${opportunity.lossReason}</p>` : ""}
    </div>
    <div class="opportunity-follow-ups">
      <div class="panel-title-row compact-row">
        <h5>Seguimientos</h5>
      </div>
      <div class="follow-up-list">${renderFollowUpsList(opportunity)}</div>
    </div>
  `;

  article.querySelector(".edit-opportunity-btn").addEventListener("click", () => {
    openOpportunityForm(opportunity);
  });
  article.querySelector(".add-follow-up-btn").addEventListener("click", () => {
    openFollowUpForm(opportunity);
  });
  article.querySelectorAll(".edit-follow-up-btn").forEach((button) => {
    button.addEventListener("click", () => {
      const followUp = opportunity.followUps.find((item) => Number(item.id) === Number(button.dataset.followUpId));
      if (!followUp) return;
      openFollowUpForm(opportunity, followUp);
    });
  });
  article.querySelectorAll(".toggle-follow-up-btn").forEach((button) => {
    button.addEventListener("click", async () => {
      const followUp = opportunity.followUps.find((item) => Number(item.id) === Number(button.dataset.followUpId));
      if (!followUp) return;
      try {
        await saveFollowUp(opportunity.id, followUp.id, {
          type: followUp.type,
          title: followUp.title,
          dueDate: followUp.dueDate,
          assignedUserId: followUp.assignedUserId,
          status: button.dataset.nextStatus,
          notes: followUp.notes
        });
        await loadClientDetail(selectedClient.id);
        renderTable();
      } catch (error) {
        notifyError(error.message);
      }
    });
  });

  return article;
}

function renderOpportunityList() {
  const opportunities = getVisibleOpportunities();
  renderCrmSummary(opportunities);
  opportunityList.innerHTML = "";

  if (!opportunities.length) {
    opportunityList.innerHTML = `
      <article class="empty-meeting-state">
        <h4>No hay oportunidades cargadas</h4>
        <p>${selectedBranchView ? "Las oportunidades de esta sucursal van a aparecer acá." : "Creá la primera oportunidad para empezar a gestionar pipeline, comentarios y próximos pasos."}</p>
      </article>
    `;
    return;
  }

  opportunities.forEach((opportunity) => {
    opportunityList.appendChild(renderOpportunityCard(opportunity));
  });
}

function renderPipelineSummary(targetNode, summaryData) {
  targetNode.innerHTML = "";
  [
    { label: "Oportunidades abiertas", value: summaryData.openCount || 0 },
    { label: "Monto abierto", value: formatMoney(summaryData.totalAmount || 0) },
    { label: "Monto ponderado", value: formatMoney(summaryData.weightedAmount || 0) },
    { label: "Seguimientos vencidos", value: summaryData.overdueFollowUps || 0 },
    { label: "Seguimientos esta semana", value: summaryData.dueThisWeek || 0 }
  ].forEach((item) => {
    const node = document.createElement("div");
    node.className = "pipeline-summary-card";
    node.innerHTML = `<span>${item.label}</span><b>${item.value}</b>`;
    targetNode.appendChild(node);
  });
}

function renderPipelineBoardGroup(targetNode, opportunities) {
  targetNode.innerHTML = "";

  if (!opportunities.length) {
    targetNode.innerHTML = `
      <article class="empty-meeting-state">
        <h4>No encontramos oportunidades</h4>
        <p>Probá cambiando los filtros o cargá una nueva oportunidad en esta categoría.</p>
      </article>
    `;
    return;
  }

  crmCatalogs.opportunityStatuses.forEach((status) => {
    const column = document.createElement("section");
    column.className = "pipeline-column";
    const items = opportunities.filter((opportunity) => opportunity.status === status);
    column.innerHTML = `
      <div class="pipeline-column-head">
        <div>
          <span class="opportunity-badge ${getOpportunityStatusClass(status)}">${status}</span>
          <h3>${items.length}</h3>
        </div>
        <strong>${formatMoney(items.reduce((total, opportunity) => total + Number(opportunity.amount || 0), 0))}</strong>
      </div>
      <div class="pipeline-column-body"></div>
    `;

    const body = column.querySelector(".pipeline-column-body");
    if (!items.length) {
      body.innerHTML = '<div class="crm-empty-inline"><strong>Sin oportunidades</strong><span>Esta etapa todavía no tiene movimiento.</span></div>';
    } else {
      items.forEach((opportunity) => {
        const cardClientName = opportunity.clientName || selectedClient?.name || "Cliente";
        const cardType = opportunity.opportunityType || "Negociación";
        const cardServiceLine = opportunity.serviceLine || "Sin línea";
        const cardOwner = opportunity.ownerName || "Sin responsable";
        const card = document.createElement("article");
        card.className = "pipeline-card";
        card.innerHTML = `
          <p class="pipeline-card-company">${cardClientName}</p>
          <h4>${opportunity.title}</h4>
          <p class="pipeline-card-meta">${[cardType, cardServiceLine, cardOwner].join(" · ")}</p>
          <div class="pipeline-card-kpis">
            <strong>${formatMoney(opportunity.amount)}</strong>
            <span>${opportunity.probability}%</span>
          </div>
          <p class="pipeline-card-next">${opportunity.nextFollowUp ? `Próximo: ${opportunity.nextFollowUp.title} · ${formatDate(opportunity.nextFollowUp.dueDate)}` : "Sin próximo paso"}</p>
        `;

        card.addEventListener("click", async () => {
          await openClientDetail(opportunity.clientId, { branchId: opportunity.branchId || null, smooth: true });
        });

        body.appendChild(card);
      });
    }

    targetNode.appendChild(column);
  });
}

function renderPipelineBoard() {
  pipelineVisibleCount.textContent = `${pipelineOpportunities.length} oportunidades`;
  const projectOpportunities = pipelineOpportunities.filter((opportunity) => opportunity.serviceLine === "Obras C.I.");
  const serviceOpportunities = pipelineOpportunities.filter((opportunity) =>
    ["Instalaciones Fijas", "Extintores", "Multiservicio"].includes(opportunity.serviceLine)
  );

  renderPipelineSummary(pipelineSummaryProjects, buildCrmSummaryClientSide(projectOpportunities));
  renderPipelineSummary(pipelineSummaryServices, buildCrmSummaryClientSide(serviceOpportunities));
  renderPipelineBoardGroup(pipelineProjectsBoard, projectOpportunities);
  renderPipelineBoardGroup(pipelineServicesBoard, serviceOpportunities);
}

function renderMeetingCard(meeting) {
  const fragment = meetingCardTemplate.content.cloneNode(true);
  const card = fragment.querySelector(".meeting-card");
  const dateNode = fragment.querySelector(".meeting-date");
  const titleNode = fragment.querySelector(".meeting-title");
  const kindNode = fragment.querySelector(".meeting-kind-badge");
  const statusNode = fragment.querySelector(".meeting-status-badge");
  const metaNode = fragment.querySelector(".meeting-meta");
  const objectiveNode = fragment.querySelector(".meeting-objective");
  const minutesNode = fragment.querySelector(".meeting-minutes");
  const findingsNode = fragment.querySelector(".meeting-findings");
  const opportunitiesNode = fragment.querySelector(".meeting-opportunities");
  const substituteRecoveryNode = fragment.querySelector(".meeting-substitute-recovery");
  const globalContactsBlock = fragment.querySelector(".meeting-global-contacts-block");
  const globalContactsNode = fragment.querySelector(".meeting-global-contacts");
  const serviceHealthStatusBlock = fragment.querySelector(".meeting-service-health-status-block");
  const serviceHealthStatusNode = fragment.querySelector(".meeting-service-health-status");
  const serviceStatusBlock = fragment.querySelector(".meeting-service-status-block");
  const serviceStatusNode = fragment.querySelector(".meeting-service-status");
  const editButton = fragment.querySelector(".edit-meeting-btn");
  const deleteButton = fragment.querySelector(".delete-meeting-btn");

  const typeMeta = getMeetingTypeMeta(meeting.kind);

  dateNode.textContent = formatDate(meeting.scheduledFor);
  titleNode.textContent = meeting.subject;
  kindNode.textContent = typeMeta.label;
  kindNode.classList.add(typeMeta.color);
  statusNode.textContent = meeting.status;
  statusNode.classList.add(
    meeting.status === "Realizada" ? "done" : meeting.status === "Confirmada" ? "confirmed" : "scheduled"
  );

  const meta = [];
  if (meeting.visitId) meta.push(`ID visita: ${meeting.visitId}`);
  if (meeting.branchId && meeting.branchName) meta.push(`Sucursal: ${meeting.branchName}`);
  const linkedOpportunityTitle =
    meeting.opportunityTitle ||
    selectedClient?.opportunities?.find((opportunity) => Number(opportunity.id) === Number(meeting.opportunityId))?.title;
  if (linkedOpportunityTitle) meta.push(`Oportunidad: ${linkedOpportunityTitle}`);
  if (meeting.participants) meta.push(`Participantes: ${meeting.participants}`);
  if (meeting.contactName || meeting.contactRole) {
    meta.push(`Contacto: ${[meeting.contactName, meeting.contactRole].filter(Boolean).join(" · ")}`);
  }
  if (meeting.modality) meta.push(`Modalidad: ${meeting.modality}`);
  if (meeting.createdBy) meta.push(`Cargado por: ${meeting.createdBy}`);
  metaNode.textContent = meta.join(" · ");

  objectiveNode.textContent = meeting.objective || "Sin objetivo registrado";
  minutesNode.textContent = meeting.minutes || "Todavía no se cargó la minuta.";
  findingsNode.textContent = meeting.activeNegotiationsStatus || "Sin status cargado.";
  opportunitiesNode.textContent = meeting.opportunities || "Sin oportunidades registradas.";
  substituteRecoveryNode.textContent = meeting.substituteRecovery || "Sin recupero de sustitutos cargado.";
  globalContactsNode.textContent = meeting.globalContacts || "Sin contactos globales registrados.";
  serviceHealthStatusNode.textContent = meeting.serviceHealthStatus || "Sin status del servicio registrado.";
  serviceStatusNode.textContent = meeting.serviceStatus || "Sin reclamos registrados.";
  globalContactsBlock.classList.toggle("hidden", !meeting.globalContacts);
  serviceHealthStatusBlock.classList.toggle("hidden", !meeting.serviceHealthStatus);
  serviceStatusBlock.classList.toggle("hidden", !meeting.serviceStatus);

  if (meeting.status === "Realizada") {
    card.classList.add("meeting-card-done");
  }

  editButton.addEventListener("click", () => {
    openMeetingScreen(meeting);
  });

  deleteButton.addEventListener("click", async () => {
    const confirmed = window.confirm(
      `La reunión "${meeting.subject}" del ${formatDate(meeting.scheduledFor)} se va a enviar a la papelera.`
    );
    if (!confirmed) return;

    try {
      await deleteMeeting(meeting.id);
      await loadClients({ clearSelectionWhenMissing: false });
      await loadClientDetail(selectedClient.id);
      await loadCalendar();
      renderTable();
      showScreen("detail");
      syncUrlWithState("detail");
    } catch (error) {
      notifyError(error.message);
    }
  });

  return fragment;
}

function renderCompactMeetingListItem(meeting) {
  const article = document.createElement("article");
  const typeMeta = getMeetingTypeMeta(meeting.kind);
  article.className = `meeting-list-item ${meeting.status === "Realizada" ? "done" : ""}`;
  article.innerHTML = `
    <div class="meeting-list-item-main">
      <p class="meeting-date">${formatDate(meeting.scheduledFor)}</p>
      <h4 class="meeting-list-item-title">${meeting.subject}</h4>
      <p class="meeting-list-item-meta">${[
        meeting.visitId ? `ID visita: ${meeting.visitId}` : "",
        meeting.kindLabel || typeMeta.label,
        meeting.opportunityId
          ? `Oportunidad: ${
              meeting.opportunityTitle ||
              selectedClient?.opportunities?.find((opportunity) => Number(opportunity.id) === Number(meeting.opportunityId))?.title ||
              "Vinculada"
            }`
          : "",
        meeting.contactName ? `${meeting.contactName}${meeting.contactRole ? ` · ${meeting.contactRole}` : ""}` : "",
        meeting.modality || ""
      ]
        .filter(Boolean)
        .join(" · ")}</p>
    </div>
    <div class="meeting-list-item-actions">
      <span class="meeting-kind-badge ${typeMeta.color}">${typeMeta.label}</span>
      <span class="meeting-status-badge ${
        meeting.status === "Realizada" ? "done" : meeting.status === "Confirmada" ? "confirmed" : "scheduled"
      }">${meeting.status}</span>
      <button class="ghost-btn delete-meeting-btn" type="button" aria-label="Enviar reunión a papelera" title="Enviar reunión a papelera">🗑</button>
    </div>
  `;

  article.querySelector(".delete-meeting-btn").addEventListener("click", async (event) => {
    event.stopPropagation();
    const confirmed = window.confirm(
      `La reunión "${meeting.subject}" del ${formatDate(meeting.scheduledFor)} se va a enviar a la papelera.`
    );
    if (!confirmed) return;

    try {
      await deleteMeeting(meeting.id);
      await loadClients({ clearSelectionWhenMissing: false });
      await loadClientDetail(selectedClient.id);
      await loadCalendar();
      renderTable();
      showScreen("detail");
      syncUrlWithState("detail");
    } catch (error) {
      notifyError(error.message);
    }
  });

  article.addEventListener("click", () => {
    openMeetingDetailScreen(meeting);
  });

  return article;
}

function renderBranchCard(branch) {
  const article = document.createElement("article");
  const canEditBranch = getCurrentEntityPermissions().editBranches;
  article.className = "meeting-card";
  article.innerHTML = `
    <div class="meeting-card-head">
      <div>
        <h4 class="meeting-title">${branch.name}</h4>
        <p class="meeting-meta">${[
          branch.manualBranchId ? `ID sucursal: ${branch.manualBranchId}` : "",
          branch.sector,
          `Segmento ${branch.segment}`,
          `Riesgo ${branch.risk}`
        ]
          .filter(Boolean)
          .join(" · ")}</p>
      </div>
      ${canEditBranch ? '<button class="secondary-btn edit-branch-btn" type="button">Editar sucursal</button>' : ""}
    </div>
    <div class="service-row">${renderServices(branch)}</div>
    <p class="meeting-objective">Ejecutivo: ${branch.executiveName || "Sin asignar"}</p>
    <p class="meeting-findings">Sup. IFCI: ${branch.supervisors?.fixedFire?.name || "Sin asignar"} · Sup. EXT: ${branch.supervisors?.extinguishers?.name || "Sin asignar"} · Sup. Obra: ${branch.supervisors?.works?.name || "Sin asignar"}</p>
  `;

  article.addEventListener("click", () => {
    selectedBranchView = branch;
    renderDetail();
    syncUrlWithState("detail");
    window.scrollTo({ top: 0, behavior: "smooth" });
  });

  const editBranchButton = article.querySelector(".edit-branch-btn");
  if (editBranchButton) {
    editBranchButton.addEventListener("click", (event) => {
      event.stopPropagation();
      selectedBranchView = branch;
      openBranchEditScreen(branch);
    });
  }

  return article;
}

function renderCalendar() {
  renderCalendarLegend();
  calendarMonthLabel.textContent =
    calendarView === "day" && selectedCalendarDay ? fullDateLabel(selectedCalendarDay) : monthLabel(calendarDate);
  calendarGrid.innerHTML = "";
  calendarDayView.innerHTML = "";
  calendarGrid.classList.toggle("hidden", calendarView === "day");
  calendarDayView.classList.toggle("hidden", calendarView !== "day");
  backToMonthBtn.classList.toggle("hidden", calendarView !== "day");
  prevMonthBtn.classList.toggle("hidden", calendarView === "day");
  nextMonthBtn.classList.toggle("hidden", calendarView === "day");

  if (calendarView === "day" && selectedCalendarDay) {
    renderCalendarDayView();
    return;
  }

  const firstDay = new Date(calendarDate.getFullYear(), calendarDate.getMonth(), 1);
  const startWeekday = (firstDay.getDay() + 6) % 7;
  const gridStart = new Date(firstDay);
  gridStart.setDate(firstDay.getDate() - startWeekday);

  const meetingsByDate = new Map();
  calendarMeetings.forEach((meeting) => {
    if (!meetingsByDate.has(meeting.scheduledFor)) {
      meetingsByDate.set(meeting.scheduledFor, []);
    }
    meetingsByDate.get(meeting.scheduledFor).push(meeting);
  });

  ["Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Dom"].forEach((label) => {
    const cell = document.createElement("div");
    cell.className = "calendar-weekday";
    cell.textContent = label;
    calendarGrid.appendChild(cell);
  });

  for (let index = 0; index < 42; index += 1) {
    const day = new Date(gridStart);
    day.setDate(gridStart.getDate() + index);
    const dayKey = dateKey(day);
    const cell = document.createElement("article");
    const isCurrentMonth = day.getMonth() === calendarDate.getMonth();
    const dayMeetings = meetingsByDate.get(dayKey) || [];

    cell.className = `calendar-day${isCurrentMonth ? "" : " muted-day"}`;
    cell.innerHTML = `
      <div class="calendar-day-head">
        <span>${day.getDate()}</span>
        <span>${dayMeetings.length ? `${dayMeetings.length} ag.` : ""}</span>
      </div>
      <div class="calendar-events"></div>
    `;

    cell.addEventListener("click", () => {
      selectedCalendarDay = dayKey;
      calendarView = "day";
      renderCalendar();
    });

    const events = cell.querySelector(".calendar-events");
    dayMeetings.slice(0, 3).forEach((meeting) => {
      const event = document.createElement("button");
      event.type = "button";
      event.className = `calendar-event ${meeting.color} ${getCalendarStatusClass(meeting.status)}`;
      event.textContent = `${meeting.scopeLabel || meeting.clientName} · ${meeting.subject}`;
      event.addEventListener("click", async () => {
        event.stopPropagation();
        await openMeetingContext(meeting);
      });
      events.appendChild(event);
    });

    if (dayMeetings.length > 3) {
      const moreButton = document.createElement("button");
      moreButton.type = "button";
      moreButton.className = "calendar-more";
      moreButton.textContent = `+${dayMeetings.length - 3} más`;
      moreButton.addEventListener("click", (event) => {
        event.stopPropagation();
        selectedCalendarDay = dayKey;
        calendarView = "day";
        renderCalendar();
      });
      events.appendChild(moreButton);
    }

    calendarGrid.appendChild(cell);
  }
}

function renderCalendarDayView() {
  const dayMeetings = calendarMeetings.filter((meeting) => meeting.scheduledFor === selectedCalendarDay);

  if (!dayMeetings.length) {
    calendarDayView.innerHTML = `
      <div class="empty-meeting-state compact-empty-state">
        <h3>Sin reuniones para esta fecha</h3>
        <p>Podés volver a la vista mensual para elegir otro día.</p>
      </div>
    `;
    return;
  }

  const wrapper = document.createElement("div");
  wrapper.className = "calendar-day-stack";

  dayMeetings.forEach((meeting) => {
    const card = document.createElement("article");
    card.className = `meeting-card day-meeting-card ${meeting.color} ${getCalendarStatusClass(meeting.status)}`;
    card.innerHTML = `
      <div class="meeting-head">
        <div>
          <p class="meeting-date">${formatDate(meeting.scheduledFor)}</p>
          <h3>${meeting.subject}</h3>
          <p class="meeting-meta">${meeting.scopeLabel || meeting.clientName} · ${meeting.kindLabel || meeting.kind} · ${meeting.modality || "Presencial"}</p>
        </div>
        <div class="meeting-head-actions">
          <span class="meeting-type-badge ${meeting.color}">${meeting.kindLabel || meeting.kind}</span>
        </div>
      </div>
      <p class="meeting-objective">${meeting.contactName ? `Con ${meeting.contactName}` : "Sin contacto definido"}${meeting.contactRole ? ` · ${meeting.contactRole}` : ""}</p>
    `;

    card.addEventListener("click", async () => {
      await openMeetingContext(meeting);
    });

    wrapper.appendChild(card);
  });

  calendarDayView.appendChild(wrapper);
}

async function openMeetingContext(meeting) {
  await openClientDetail(meeting.clientId, { branchId: meeting.branchId || null });
}

function openMeetingDetailScreen(meeting) {
  if (!selectedClient) return;
  meetingScreenMode = "detail";
  meetingMode = "edit";
  editingMeetingId = meeting?.id || null;
  meetingScreenTitle.textContent = `Reunión · ${selectedClient.name}`;
  meetingScreenMeta.textContent = meeting?.visitId ? `ID visita: ${meeting.visitId} · Detalle completo de la reunión.` : "Detalle completo de la reunión.";
  meetingForm.classList.add("hidden");
  meetingDetailView.classList.remove("hidden");
  editMeetingFromDetailBtn.classList.remove("hidden");
  meetingDetailView.innerHTML = "";
  meetingDetailView.appendChild(renderMeetingCard(meeting));
  showScreen("meeting");
}

function renderDetail() {
  if (!selectedClient) return;
  opportunityForm.classList.add("hidden");
  followUpForm.classList.add("hidden");
  syncModalState();
  const meetings = Array.isArray(selectedClient.meetings) ? selectedClient.meetings : [];
  const parentMeetings = meetings.filter((meeting) => !meeting.branchId);
  const opportunities = CRM_ENABLED ? getVisibleOpportunities() : [];
  const crmSummary = CRM_ENABLED ? buildCrmSummaryClientSide(opportunities) : { openCount: 0 };
  const currentEntity = selectedBranchView || selectedClient;
  const entityMeetings = selectedBranchView
    ? meetings.filter((meeting) => Number(meeting.branchId) === Number(selectedBranchView.id))
    : parentMeetings;
  const meetingsCompleted = Number(selectedClient.meetingsCompleted || 0);
  const meetingsCount = Number(selectedClient.meetingsCount || meetings.length);
  const supervisorSummary = [
    currentEntity.supervisors?.fixedFire ? `Sup. IFCI: ${currentEntity.supervisors.fixedFire.name}` : "",
    currentEntity.supervisors?.extinguishers ? `Sup. EXT: ${currentEntity.supervisors.extinguishers.name}` : "",
    currentEntity.supervisors?.works ? `Sup. Obra: ${currentEntity.supervisors.works.name}` : ""
  ]
    .filter(Boolean)
    .join(" · ");

  backToList.textContent = selectedBranchView ? "← Volver a la casa matriz" : "← Volver al listado";
  goToEditBtn.textContent = selectedBranchView ? "Editar sucursal" : "Editar compañía";
  const permissions = getCurrentEntityPermissions();
  const canEditCurrentEntity = selectedBranchView ? permissions.editBranches : permissions.editClients;
  goToEditBtn.classList.toggle("hidden", !canEditCurrentEntity);
  detailCompanyName.textContent = selectedBranchView ? `${selectedBranchView.name} · ${selectedClient.name}` : selectedClient.name;
  detailMeta.textContent = selectedBranchView
    ? `${[
        currentEntity.manualBranchId ? `ID sucursal: ${currentEntity.manualBranchId}` : "",
        currentEntity.sector,
        `Ejecutivo: ${currentEntity.executiveName || currentEntity.manager}`,
        `Riesgo ${currentEntity.risk}`,
        supervisorSummary
      ]
        .filter(Boolean)
        .join(" · ")}`
    : `${currentEntity.sector} · ${currentEntity.companyType} · ${currentEntity.country} · ${currentEntity.accountStage} · Ejecutivo: ${currentEntity.executiveName || currentEntity.manager} · Riesgo ${currentEntity.risk}${supervisorSummary ? ` · ${supervisorSummary}` : ""}`;
  detailSegment.textContent = `Segmento ${currentEntity.segment}`;
  detailServices.innerHTML = renderServices(currentEntity);
  renderMeetingOpportunityOptions();
  renderOpportunityBranchOptions(selectedBranchView ? selectedBranchView.id : "");
  renderDetailSectionNav();

  detailKpis.innerHTML = "";
  if (selectedBranchView) {
    detailKpis.append(
      renderRoleAssignmentsCard("Ejecutivo", currentEntity.executiveName),
      renderRoleAssignmentsCard("Supervisor IFCI", currentEntity.supervisors?.fixedFire?.name),
      renderRoleAssignmentsCard("Supervisor EXT", currentEntity.supervisors?.extinguishers?.name),
      renderRoleAssignmentsCard("Supervisor Obra", currentEntity.supervisors?.works?.name),
      kpiItem("Reuniones", `${entityMeetings.filter((meeting) => meeting.status === "Realizada").length}/${entityMeetings.length}`)
    );
  } else {
    detailKpis.append(
      renderRoleAssignmentsCard("Ejecutivo", selectedClient.executiveName),
      renderRoleAssignmentsCard("Supervisor IFCI", selectedClient.supervisors?.fixedFire?.name),
      renderRoleAssignmentsCard("Supervisor EXT", selectedClient.supervisors?.extinguishers?.name),
      renderRoleAssignmentsCard("Supervisor Obra", selectedClient.supervisors?.works?.name),
      kpiItem("Reuniones", `${meetingsCompleted}/${meetingsCount}`)
    );
  }

  if (CRM_ENABLED) {
    renderOpportunityList();
  }

  meetingList.innerHTML = "";

  if (!entityMeetings.length) {
    meetingList.innerHTML = `
      <article class="empty-meeting-state">
        <h4>No hay reuniones cargadas</h4>
        <p>${selectedBranchView ? "Las reuniones de esta sucursal van a aparecer acá." : "Las reuniones de casa matriz van a aparecer acá."}</p>
      </article>
    `;
  } else {
    entityMeetings.forEach((meeting) => {
      meetingList.appendChild(renderCompactMeetingListItem(meeting));
    });
  }

  branchList.innerHTML = "";
  if (selectedBranchView) {
    branchList.innerHTML = "";
    addBranchBtn.classList.add("hidden");
    return;
  }

  addBranchBtn.classList.toggle("hidden", !permissions.createBranches);
  if (!selectedClient.branches?.length) {
    branchList.innerHTML = `
      <article class="empty-meeting-state">
        <h4>No hay sucursales cargadas</h4>
        <p>Agregá la primera sucursal para gestionar responsables y servicios por ubicación.</p>
      </article>
    `;
  } else {
    selectedClient.branches.forEach((branch) => {
      branchList.appendChild(renderBranchCard(branch));
    });
  }
}

function renderUsers() {
  const currentRoleValue = newUserRole.value;
  newUserRole.innerHTML = userRoles.map((role) => `<option value="${role}">${role}</option>`).join("");
  if (currentRoleValue && userRoles.includes(currentRoleValue)) {
    newUserRole.value = currentRoleValue;
  }
  usersTableBody.innerHTML = "";

  users.forEach((user) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td><b>${user.name}</b></td>
      <td>${user.email}</td>
      <td>${user.role}</td>
      <td>${user.bitrixUserId || "-"}</td>
      <td>${summarizePermissions(user.permissions || {})}</td>
      <td>
        <div class="table-actions">
          <button class="secondary-btn edit-user-btn" type="button">Editar</button>
          <button class="ghost-btn delete-user-btn" type="button">Eliminar</button>
        </div>
      </td>
    `;

    row.querySelector(".edit-user-btn").addEventListener("click", () => {
      startUserEdit(user);
    });

    row.querySelector(".delete-user-btn").addEventListener("click", async () => {
      const confirmed = window.confirm(`Se va a eliminar el usuario ${user.name}. Esta acción no se puede deshacer.`);
      if (!confirmed) return;

      try {
        await deleteUser(user.id);
        if (editingUserId === user.id) resetUserForm();
        createUserStatus.textContent = "Usuario eliminado";
        await loadUsers();
        await loadAssignmentOptions();
      } catch (error) {
        notifyError(error.message);
      }
    });

    usersTableBody.appendChild(row);
  });
}

async function createUser(payload) {
  const response = await fetch("/api/users", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo crear el usuario");
  }
  return data.user;
}

async function updateUser(userId, payload) {
  const response = await fetch(`/api/users/${userId}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo actualizar el usuario");
  }
  return data.user;
}

async function deleteUser(userId) {
  const response = await fetch(`/api/users/${userId}`, {
    method: "DELETE"
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo eliminar el usuario");
  }
  return data;
}

async function deleteMeeting(meetingId) {
  if (!selectedClient?.id) {
    throw new Error("No hay un cliente seleccionado");
  }

  const response = await fetch(`/api/clients/${selectedClient.id}/meetings/${meetingId}`, {
    method: "DELETE"
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo enviar la reunión a la papelera");
  }
  return data;
}

async function loadSectorOptions() {
  const response = await fetch("/api/sector-options");
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudieron cargar los sectores");
  }
  sectorOptions = data.sectors || [];
  renderSectorSelect(editSector.value);
  if (canAccessSettings()) {
    renderSectors();
  }
}

async function loadMeetingTypesConfig() {
  const response = await fetch("/api/settings/meeting-types");
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudieron cargar los tipos de reunión");
  }
  meetingTypes = data.meetingTypes || [];
  renderTypeSelectOptions();
  renderMeetingTypesConfig();
}

async function loadMeetingReasonsConfig() {
  const response = await fetch("/api/settings/meeting-reasons");
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudieron cargar los motivos de reunión");
  }
  meetingReasons = data.meetingReasons || [];
  renderTypeSelectOptions();
  renderMeetingReasonsConfig();
}

async function loadContactRolesConfig() {
  const response = await fetch("/api/settings/contact-roles");
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudieron cargar las funciones del contacto");
  }
  contactRoles = data.contactRoles || [];
  renderTypeSelectOptions();
  renderContactRolesConfig();
}

async function loadDeletedMeetings() {
  const response = await fetch("/api/settings/trash/meetings");
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo cargar la papelera");
  }
  deletedMeetings = data.meetings || [];
  trashMeetingsStatus.textContent = deletedMeetings.length
    ? `${deletedMeetings.length} reuniones en papelera`
    : "Papelera vacía";
  renderDeletedMeetings();
}

async function loadAuditLogs() {
  const response = await fetch("/api/settings/audit-logs?limit=150");
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo cargar la auditoría");
  }
  auditLogs = data.logs || [];
  auditLogsStatus.textContent = auditLogs.length ? `${auditLogs.length} movimientos recientes` : "Sin movimientos recientes";
  renderAuditLogs();
}

async function loadFeatureSettings() {
  const response = await fetch("/api/settings/features");
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo cargar la configuración de módulos");
  }
  visitRulesFeatureEnabled = Boolean(data?.featureFlags?.visitRulesEnabled ?? true);
  applyFeatureFlagsUi();
}

async function loadSettingsCatalogs() {
  await loadSectorOptions();
  await loadMeetingTypesConfig();
  await loadMeetingReasonsConfig();
  await loadContactRolesConfig();
  await loadDeletedMeetings();
  await loadAuditLogs();
  await loadFeatureSettings();
  try {
    const bitrixConfig = await loadBitrixConfig();
    bitrixWebhookUrl.value = bitrixConfig.webhookUrl || "";
  } catch (_error) {
    // Si falla la lectura de la configuración, no bloqueamos el resto de la pantalla.
  }
  renderBitrixUsersPreview();
  renderBitrixMappings();
}

async function createSector(name) {
  const response = await fetch("/api/settings/sectors", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name })
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo crear el sector");
  }
  return data.sector;
}

async function deleteSector(sectorId) {
  const response = await fetch(`/api/settings/sectors/${sectorId}`, {
    method: "DELETE"
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo eliminar el sector");
  }
  return data;
}

async function saveMeetingType(payload) {
  const endpoint = editingMeetingTypeId ? `/api/settings/meeting-types/${editingMeetingTypeId}` : "/api/settings/meeting-types";
  const method = editingMeetingTypeId ? "PATCH" : "POST";
  const response = await fetch(endpoint, {
    method,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo guardar el tipo de reunión");
  }
  return data.meetingTypes || [];
}

async function deleteMeetingType(id) {
  const response = await fetch(`/api/settings/meeting-types/${id}`, {
    method: "DELETE"
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo eliminar el tipo de reunión");
  }
  return data.meetingTypes || [];
}

async function saveMeetingReason(payload) {
  const endpoint = editingMeetingReasonId
    ? `/api/settings/meeting-reasons/${editingMeetingReasonId}`
    : "/api/settings/meeting-reasons";
  const method = editingMeetingReasonId ? "PATCH" : "POST";
  const response = await fetch(endpoint, {
    method,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo guardar el motivo de reunión");
  }
  return data.meetingReasons || [];
}

async function deleteMeetingReason(id) {
  const response = await fetch(`/api/settings/meeting-reasons/${id}`, {
    method: "DELETE"
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo eliminar el motivo de reunión");
  }
  return data.meetingReasons || [];
}

async function saveContactRole(payload) {
  const endpoint = editingContactRoleId ? `/api/settings/contact-roles/${editingContactRoleId}` : "/api/settings/contact-roles";
  const method = editingContactRoleId ? "PATCH" : "POST";
  const response = await fetch(endpoint, {
    method,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo guardar la función del contacto");
  }
  return data.contactRoles || [];
}

async function deleteContactRole(id) {
  const response = await fetch(`/api/settings/contact-roles/${id}`, {
    method: "DELETE"
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo eliminar la función del contacto");
  }
  return data.contactRoles || [];
}

async function downloadClientsTemplate() {
  const response = await fetch("/api/settings/clients-import-template");
  if (!response.ok) {
    throw new Error("No se pudo descargar la plantilla");
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "plantilla-clientes.xlsx";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function downloadBranchesTemplate() {
  const response = await fetch("/api/settings/branches-import-template");
  if (!response.ok) {
    throw new Error("No se pudo descargar la plantilla de sucursales");
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "plantilla-sucursales.xlsx";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function importClientsFromExcel(file) {
  const fileData = await file.arrayBuffer();
  const bytes = new Uint8Array(fileData);
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  const base64 = window.btoa(binary);

  const response = await fetch("/api/settings/clients-import", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileName: file.name,
      fileData: base64
    })
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo importar el archivo");
  }
  return data;
}

async function importBranchesFromExcel(file) {
  const fileData = await file.arrayBuffer();
  const bytes = new Uint8Array(fileData);
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  const base64 = window.btoa(binary);

  const response = await fetch("/api/settings/branches-import", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileName: file.name,
      fileData: base64
    })
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo importar el archivo de sucursales");
  }
  return data;
}

async function downloadUsersTemplate() {
  const response = await fetch("/api/settings/users-import-template");
  if (!response.ok) {
    throw new Error("No se pudo descargar la plantilla de usuarios");
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "plantilla-usuarios.xlsx";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function importUsersFromExcel(file) {
  const fileData = await file.arrayBuffer();
  const bytes = new Uint8Array(fileData);
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  const base64 = window.btoa(binary);

  const response = await fetch("/api/settings/users-import", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileName: file.name,
      fileData: base64
    })
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo importar el archivo de usuarios");
  }
  return data;
}

async function restoreDeletedMeeting(meetingId) {
  const response = await fetch(`/api/settings/trash/meetings/${meetingId}/restore`, {
    method: "POST"
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo recuperar la reunión");
  }
  return data;
}

async function previewBitrixUsers(webhookUrlValue) {
  const response = await fetch("/api/settings/bitrix/users-preview", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      webhookUrl: webhookUrlValue
    })
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudieron consultar los usuarios de Bitrix");
  }
  return data;
}

async function loadBitrixConfig() {
  const response = await fetch("/api/settings/bitrix/config");
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo cargar la configuración de Bitrix");
  }
  return data;
}

async function saveBitrixConfig(webhookUrlValue) {
  const response = await fetch("/api/settings/bitrix/config", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      webhookUrl: webhookUrlValue
    })
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo guardar la configuración de Bitrix");
  }
  return data;
}

async function saveFeatureSettings(payload) {
  const response = await fetch("/api/settings/features", {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo guardar la configuración de módulos");
  }
  return data;
}

async function previewBitrixMappings(webhookUrlValue) {
  const response = await fetch("/api/settings/bitrix/mappings-preview", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      webhookUrl: webhookUrlValue
    })
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudieron generar los mapeos de Bitrix");
  }
  return data;
}

async function createBitrixTaskTest(payload) {
  const response = await fetch("/api/settings/bitrix/tasks-test", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo crear la tarea de prueba en Bitrix");
  }
  return data;
}

function clearImportFeedback(statusNode, detailsNode) {
  statusNode.textContent = "";
  detailsNode.innerHTML = "";
  detailsNode.classList.add("hidden");
}

function renderImportFeedback(statusNode, detailsNode, result, entityLabel) {
  statusNode.textContent = `Importación finalizada: ${result.createdCount} ${entityLabel} creados${result.errorCount ? `, ${result.errorCount} con error` : ""}.`;
  const errors = Array.isArray(result.errors) ? result.errors : [];
  if (!errors.length) {
    detailsNode.innerHTML = "";
    detailsNode.classList.add("hidden");
    return;
  }

  detailsNode.classList.remove("hidden");
  const visibleErrors = errors.slice(0, 8);
  detailsNode.innerHTML = `
    <strong>Filas con observaciones</strong>
    <ul>
      ${visibleErrors.map((error) => `<li>${error}</li>`).join("")}
    </ul>
    ${errors.length > visibleErrors.length ? `<span>Y ${errors.length - visibleErrors.length} más.</span>` : ""}
  `;
}

function resetUserForm() {
  userFormMode = "create";
  editingUserId = null;
  userFormTitle.textContent = "Nuevo usuario";
  createUserBtn.textContent = "Crear usuario";
  cancelUserEditBtn.classList.add("hidden");
  newUserPassword.required = true;
  newUserPassword.placeholder = "";
  newUserName.value = "";
  newUserEmail.value = "";
  newUserBitrixSearch.value = "";
  newUserBitrixId.value = "";
  closeBitrixSearchDropdown(newUserBitrixOptions);
  newUserPassword.value = "";
  applyUserPermissionsForm();
  if (userRoles.length) {
    newUserRole.value = userRoles[0];
  }
}

function startUserEdit(user) {
  userFormMode = "edit";
  editingUserId = user.id;
  userFormTitle.textContent = `Editar usuario · ${user.name}`;
  createUserBtn.textContent = "Guardar usuario";
  cancelUserEditBtn.classList.remove("hidden");
  newUserPassword.required = false;
  newUserPassword.placeholder = "Dejar vacío para mantener la actual";
  newUserName.value = user.name;
  newUserEmail.value = user.email;
  newUserRole.value = user.role;
  newUserBitrixSearch.value = getBitrixUserLabel(user.bitrixUserId || "");
  newUserBitrixId.value = user.bitrixUserId || "";
  applyUserPermissionsForm(user.permissions || {});
  closeBitrixSearchDropdown(newUserBitrixOptions);
  newUserPassword.value = "";
  createUserStatus.textContent = "";
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function fillEditForm(client) {
  const entity = client || {};
  renderSectorSelect(entity?.sector || "");
  renderUserSelect(editManager, assignmentOptions.executives, entity?.executiveUserId, "Seleccionar ejecutivo");
  renderUserSelect(
    editSupervisorFixedFire,
    assignmentOptions.supervisors.fixedFire,
    entity?.supervisors?.fixedFire?.userId,
    "Seleccionar supervisor IFCI"
  );
  renderUserSelect(
    editSupervisorExtinguishers,
    assignmentOptions.supervisors.extinguishers,
    entity?.supervisors?.extinguishers?.userId,
    "Seleccionar supervisor EXT"
  );
  renderUserSelect(
    editSupervisorWorks,
    assignmentOptions.supervisors.works,
    entity?.supervisors?.works?.userId,
    "Seleccionar supervisor Obra"
  );

  editName.value = entity?.name || "";
  editCompanyType.value = entity?.companyType || "Local";
  editCountry.value = entity?.country || "Argentina";
  editAccountStage.value = entity?.accountStage || "Activa";
  editBitrixCompanySearch.value = getBitrixCompanyLabel(entity?.bitrixCompanyId || "");
  editBitrixLeadId.value = entity?.bitrixLeadId || "";
  editBitrixCompanyId.value = entity?.bitrixCompanyId || "";
  editManualBranchId.value = entity?.manualBranchId || "";
  closeBitrixSearchDropdown(editBitrixCompanyOptions);
  editRisk.value = entity?.risk || "Bajo";
  editSegment.value = entity?.segment || "C";
  editFixedFire.checked = !!entity?.services?.fixedFire;
  editExtinguishers.checked = !!entity?.services?.extinguishers;
  editWorks.checked = !!entity?.services?.works;
  editNotes.value = entity?.notes || "";
  currentVisitRules = Array.isArray(entity?.visitRules)
    ? entity.visitRules.map((rule) => ({
        id: rule.id || null,
        periodicityDays: Number(rule.periodicityDays || 30),
        contactRole: rule.contactRole || "",
        objective: rule.objective || ""
      }))
    : [];
  editCompanyStatus.textContent = "";
  editManager.required = editEntityType !== "branch";
  editCompanyType.required = editEntityType !== "branch";
  editCountry.required = editEntityType !== "branch";
  editAccountStage.required = editEntityType !== "branch";
  editManager.closest("label").classList.toggle("hidden", editEntityType === "branch");
  document.getElementById("editCompanyTypeGroup").classList.toggle("hidden", editEntityType === "branch");
  document.getElementById("editCountryGroup").classList.toggle("hidden", editEntityType === "branch");
  document.getElementById("editAccountStageGroup").classList.toggle("hidden", editEntityType === "branch");
  editManualBranchIdGroup.classList.toggle("hidden", editEntityType !== "branch");
  editBitrixCompanySearch.closest("label").classList.toggle("hidden", false);
  syncSupervisorVisibility();
  renderAccountRoleSummary();
  renderVisitRulesEditor();
}

function fillMeetingForm(meeting) {
  const defaultKind = meetingTypes[0]?.value || "Comercial";
  const defaultStatus = meetingStatuses[0] || "Agendada";
  const defaultModality = meetingModalities[0] || "Presencial";
  meetingScope.innerHTML = [
    `<option value="">Seleccionar alcance</option>`,
    `<option value="client" ${!meeting?.branchId ? "selected" : ""}>${selectedClient.name} (Casa matriz)</option>`,
    ...(selectedClient?.branches || []).map(
      (branch) =>
        `<option value="branch:${branch.id}" ${Number(meeting?.branchId) === Number(branch.id) ? "selected" : ""}>${branch.name}</option>`
    )
  ].join("");
  meetingKind.value = meeting?.kind || defaultKind;
  meetingSubject.value = meeting?.subject || "";
  meetingObjective.value = meeting?.objective || "";
  meetingDate.value = meeting?.scheduledFor || "";
  meetingModality.value = meeting?.modality || defaultModality;
  const participantNames = String(meeting?.participants || "")
    .split(",")
    .map((name) => name.trim().toLowerCase())
    .filter(Boolean);
  const fallbackParticipantIds = getParticipantUsers()
    .filter((user) => participantNames.includes(user.name.trim().toLowerCase()))
    .map((user) => user.id);
  setSelectedParticipantUserIds(meeting?.participantUserIds?.length ? meeting.participantUserIds : fallbackParticipantIds);
  meetingParticipantsSearch.value = "";
  renderMeetingParticipantsPicker();
  meetingContactName.value = meeting?.contactName || "";
  meetingContactRole.value = "";
  const nextContactRole = meeting?.contactRole || "";
  if (nextContactRole) {
    if (!contactRoles.some((role) => role.name === nextContactRole)) {
      meetingContactRole.innerHTML += `<option value="${nextContactRole}">${nextContactRole}</option>`;
    }
    meetingContactRole.value = nextContactRole;
  }
  meetingStatusSelect.value = meeting?.status || defaultStatus;
  renderMeetingOpportunityOptions(meeting?.opportunityId || "");
  meetingNextDate.value = meeting?.nextMeetingDate || "";
  meetingMinutes.value = meeting?.minutes || "";
  meetingNegotiationStatus.value = meeting?.activeNegotiationsStatus || meeting?.findings || "";
  meetingOpportunities.value = meeting?.opportunities || "";
  meetingSubstituteRecovery.value = meeting?.substituteRecovery || "";
  meetingGlobalContacts.value = meeting?.globalContacts || "";
  meetingServiceHealthStatus.value = meeting?.serviceHealthStatus || "";
  meetingServiceStatus.value = meeting?.serviceStatus || "";
  renderComplaintBitrixResponsibleOptions(meeting?.complaintBitrixResponsibleId || "");
  meetingStatus.textContent = "";
  syncMeetingCompletionFields();
  syncMeetingContextBlocks();
}

function openCreateScreen() {
  if (!getCurrentEntityPermissions().createClients) {
    notifyError("No tenés permisos para crear compañías.");
    return;
  }
  editMode = "create";
  editEntityType = "client";
  editingBranchId = null;
  editScreenTitle.textContent = "Nueva compañía";
  saveCompanyBtn.textContent = "Crear compañía";
  hideCompanyBtn.classList.add("hidden");
  fillEditForm(null);
  showScreen("edit");
}

function openEditScreen() {
  if (!selectedClient) return;
  if (!getCurrentEntityPermissions().editClients) {
    notifyError("No tenés permisos para editar compañías.");
    return;
  }
  editMode = "edit";
  editEntityType = "client";
  editingBranchId = null;
  editScreenTitle.textContent = `Editar compañía · ${selectedClient.name}`;
  saveCompanyBtn.textContent = "Guardar cambios";
  hideCompanyBtn.textContent = "Ocultar compañía";
  hideCompanyBtn.classList.toggle("hidden", !getCurrentEntityPermissions().hideClients);
  fillEditForm(selectedClient);
  showScreen("edit");
}

function openBranchCreateScreen() {
  if (!selectedClient) return;
  if (!getCurrentEntityPermissions().createBranches) {
    notifyError("No tenés permisos para crear sucursales.");
    return;
  }
  editMode = "create";
  editEntityType = "branch";
  editingBranchId = null;
  editScreenTitle.textContent = `Nueva sucursal · ${selectedClient.name}`;
  saveCompanyBtn.textContent = "Crear sucursal";
  hideCompanyBtn.textContent = "Ocultar sucursal";
  hideCompanyBtn.classList.add("hidden");
  fillEditForm({
    sector: selectedClient.sector,
    companyType: selectedClient.companyType,
    country: selectedClient.country,
    accountStage: selectedClient.accountStage,
    risk: selectedClient.risk,
    segment: selectedClient.segment,
    services: { fixedFire: false, extinguishers: false, works: false },
    supervisors: { fixedFire: null, extinguishers: null, works: null },
    executiveUserId: selectedClient.executiveUserId,
    notes: "",
    visitRules: []
  });
  showScreen("edit");
}

function openBranchEditScreen(branch) {
  if (!getCurrentEntityPermissions().editBranches) {
    notifyError("No tenés permisos para editar sucursales.");
    return;
  }
  editMode = "edit";
  editEntityType = "branch";
  editingBranchId = branch.id;
  editScreenTitle.textContent = `Editar sucursal · ${branch.name}`;
  saveCompanyBtn.textContent = "Guardar sucursal";
  hideCompanyBtn.textContent = "Ocultar sucursal";
  hideCompanyBtn.classList.toggle("hidden", !getCurrentEntityPermissions().hideBranches);
  fillEditForm(branch);
  showScreen("edit");
}

function renderAccountRoleSummary() {
  accountRolesGrid.innerHTML = "";
  const executiveName =
    editEntityType === "branch"
      ? selectedClient?.executiveName || "Sin asignar"
      : editManager.selectedOptions[0]?.textContent || "Sin asignar";
  const fixedFireName = editFixedFire.checked ? editSupervisorFixedFire.selectedOptions[0]?.textContent || "Sin asignar" : "No aplica";
  const extinguishersName = editExtinguishers.checked
    ? editSupervisorExtinguishers.selectedOptions[0]?.textContent || "Sin asignar"
    : "No aplica";
  const worksName = editWorks.checked ? editSupervisorWorks.selectedOptions[0]?.textContent || "Sin asignar" : "No aplica";

  accountRolesGrid.append(
    renderRoleAssignmentsCard("Ejecutivo", executiveName),
    renderRoleAssignmentsCard("Supervisor IFCI", fixedFireName),
    renderRoleAssignmentsCard("Supervisor EXT", extinguishersName),
    renderRoleAssignmentsCard("Supervisor Obra", worksName)
  );
}

function openMeetingScreen(meeting = null) {
  if (!selectedClient) return;
  meetingScreenMode = "form";
  meetingMode = meeting ? "edit" : "create";
  editingMeetingId = meeting?.id || null;
  meetingScreenTitle.textContent = meeting
    ? `Editar reunión · ${selectedClient.name}`
    : `Nueva reunión · ${selectedClient.name}`;
  meetingScreenMeta.textContent = meeting
    ? `${meeting?.visitId ? `ID visita: ${meeting.visitId} · ` : ""}Actualizá tipo, estado, participantes, minuta y hallazgos.`
    : "Podés dejarla agendada en una fecha futura o marcarla como realizada si ya ocurrió.";
  saveMeetingBtn.textContent = meeting ? "Guardar cambios" : "Guardar reunión";
  meetingDetailView.classList.add("hidden");
  editMeetingFromDetailBtn.classList.add("hidden");
  meetingForm.classList.remove("hidden");
  fillMeetingForm(meeting);
  if (!meeting && selectedBranchView) {
    meetingScope.value = `branch:${selectedBranchView.id}`;
  }
  showScreen("meeting");
}

function getFilters() {
  return {
    search: searchInput.value.trim(),
    risk: riskFilter.value,
    segment: segmentFilter.value,
    fixedFire: fixedFireFilter.value,
    extinguishers: extinguishersFilter.value,
    works: worksFilter.value
  };
}

function buildQuery(params) {
  const qs = new URLSearchParams();
  Object.entries(params).forEach(([key, value]) => {
    if (value && value !== "todos") qs.set(key, value);
  });
  return qs.toString();
}

function getBitrixUserLabel(bitrixUserId) {
  const normalizedId = String(bitrixUserId || "").trim();
  const match = bitrixDirectory.users.find((item) => item.id === normalizedId);
  return match?.label || "";
}

function getBitrixCompanyLabel(bitrixCompanyId) {
  const normalizedId = String(bitrixCompanyId || "").trim();
  const match = bitrixDirectory.companies.find((item) => item.id === normalizedId);
  return match?.label || "";
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getBitrixSearchMatches(options, query) {
  const normalizedQuery = String(query || "").trim().toLowerCase();
  if (!normalizedQuery) return options.slice(0, 25);

  return options
    .filter((item) =>
      [item.label, item.name, item.email, item.title, item.id]
        .filter(Boolean)
        .some((value) => String(value).toLowerCase().includes(normalizedQuery))
    )
    .slice(0, 25);
}

function getBitrixUsersEmptyText(query = "") {
  if (!bitrixDirectoryPermissions.users) {
    return bitrixDirectoryErrors.users || "No se pudieron cargar usuarios de Bitrix.";
  }
  if (!bitrixDirectory.users.length) {
    return "No hay usuarios activos de Bitrix disponibles.";
  }
  if (String(query || "").trim()) {
    return "No encontramos usuarios activos de Bitrix para esa búsqueda.";
  }
  return "No hay usuarios activos de Bitrix disponibles.";
}

function getBitrixCompaniesEmptyText(query = "") {
  if (!bitrixDirectoryPermissions.companies) {
    return bitrixDirectoryErrors.companies || "El webhook actual no tiene permisos para leer compañías de Bitrix.";
  }
  const availableCompanies = bitrixCompanySearchResults.length ? bitrixCompanySearchResults : bitrixDirectory.companies;
  if (!availableCompanies.length) {
    return "No hay compañías de Bitrix disponibles.";
  }
  if (String(query || "").trim()) {
    return "No encontramos compañías de Bitrix para esa búsqueda.";
  }
  return "No hay compañías de Bitrix disponibles.";
}

function openBitrixSearchDropdown(container) {
  if (activeBitrixSearchDropdown && activeBitrixSearchDropdown !== container) {
    activeBitrixSearchDropdown.classList.add("hidden");
  }
  activeBitrixSearchDropdown = container;
  container.classList.remove("hidden");
}

function closeBitrixSearchDropdown(container) {
  container.classList.add("hidden");
  if (activeBitrixSearchDropdown === container) {
    activeBitrixSearchDropdown = null;
  }
}

function renderBitrixSearchDropdown(container, options, query, emptyText) {
  const matches = getBitrixSearchMatches(options, query);

  if (!matches.length) {
    container.innerHTML = `<div class="bitrix-search-empty">${escapeHtml(emptyText)}</div>`;
    openBitrixSearchDropdown(container);
    return;
  }

  container.innerHTML = matches
    .map(
      (item) => `
        <button
          class="bitrix-search-option"
          type="button"
          data-id="${escapeHtml(item.id)}"
          data-label="${escapeHtml(item.label)}"
        >
          <strong>${escapeHtml(item.name || item.title || item.label)}</strong>
          <small>${escapeHtml(item.email || item.label)}</small>
        </button>
      `
    )
    .join("");

  openBitrixSearchDropdown(container);
}

function renderBitrixDirectoryDatalists() {
  if (activeBitrixSearchDropdown === newUserBitrixOptions) {
    renderBitrixSearchDropdown(
      newUserBitrixOptions,
      bitrixDirectory.users,
      newUserBitrixSearch.value,
      getBitrixUsersEmptyText(newUserBitrixSearch.value)
    );
  }

  if (activeBitrixSearchDropdown === editBitrixCompanyOptions) {
    renderBitrixSearchDropdown(
      editBitrixCompanyOptions,
      bitrixCompanySearchResults.length ? bitrixCompanySearchResults : bitrixDirectory.companies,
      editBitrixCompanySearch.value,
      getBitrixCompaniesEmptyText(editBitrixCompanySearch.value)
    );
  }
}

function syncBitrixUserSelectionFromSearch() {
  const normalizedValue = newUserBitrixSearch.value.trim();
  const match = bitrixDirectory.users.find((item) => item.label === normalizedValue);
  newUserBitrixId.value = match?.id || "";
}

function syncBitrixCompanySelectionFromSearch() {
  const normalizedValue = editBitrixCompanySearch.value.trim();
  const companyOptions = [...bitrixCompanySearchResults, ...bitrixDirectory.companies];
  const match = companyOptions.find((item) => item.label === normalizedValue);
  editBitrixCompanyId.value = match?.id || "";
}

function mergeBitrixCompanies(companies) {
  const existing = new Map(bitrixDirectory.companies.map((item) => [item.id, item]));
  companies.forEach((company) => {
    if (!company?.id) return;
    existing.set(company.id, company);
  });
  bitrixDirectory.companies = Array.from(existing.values());
}

async function searchBitrixCompaniesRemote(query) {
  const normalizedQuery = String(query || "").trim();
  if (!normalizedQuery) {
    bitrixCompanySearchResults = [];
    renderBitrixDirectoryDatalists();
    return;
  }

  const requestId = ++bitrixCompanySearchRequestId;
  try {
    const response = await fetch(`/api/bitrix/companies-search?q=${encodeURIComponent(normalizedQuery)}`);
    const data = await response.json();
    if (requestId !== bitrixCompanySearchRequestId) return;

    bitrixDirectoryPermissions.companies = data.permissions?.companies ?? bitrixDirectoryPermissions.companies;
    bitrixDirectoryErrors.companies = data.errors?.companies || "";
    bitrixCompanySearchResults = Array.isArray(data.companies) ? data.companies : [];
    mergeBitrixCompanies(bitrixCompanySearchResults);
    renderBitrixDirectoryDatalists();
  } catch (_error) {
    if (requestId !== bitrixCompanySearchRequestId) return;
    bitrixCompanySearchResults = [];
    renderBitrixDirectoryDatalists();
  }
}

async function loadCatalogs() {
  const response = await fetch("/api/meeting-types");
  const data = await response.json();
  meetingTypes = data.meetingTypes || [];
  meetingReasons = data.meetingReasons || [];
  contactRoles = data.contactRoles || [];
  meetingStatuses = data.statuses || [];
  meetingModalities = data.modalities || ["Presencial", "Virtual"];
  renderTypeSelectOptions();
  await loadSectorOptions();
  try {
    const bitrixResponse = await fetch("/api/bitrix/options");
    const bitrixData = await bitrixResponse.json();
    if (bitrixResponse.ok) {
      bitrixDirectory = {
        users: bitrixData.users || [],
        companies: bitrixData.companies || []
      };
      bitrixCompanySearchResults = [];
      bitrixDirectoryPermissions = bitrixData.permissions || {
        users: true,
        companies: true
      };
      bitrixDirectoryErrors = bitrixData.errors || {
        users: "",
        companies: ""
      };
      renderBitrixDirectoryDatalists();
    }
  } catch (_error) {
    // No bloqueamos el uso local si Bitrix no está disponible.
  }
}

async function loadCrmCatalogs() {
  const response = await fetch("/api/crm/catalogs");
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudieron cargar los catálogos comerciales");
  }
  crmCatalogs = {
    opportunityStatuses: data.opportunityStatuses || [],
    openOpportunityStatuses: data.openOpportunityStatuses || [],
    opportunityTypes: data.opportunityTypes || [],
    serviceLines: data.serviceLines || [],
    followUpStatuses: data.followUpStatuses || [],
    followUpTypes: data.followUpTypes || []
  };
  renderOpportunitySelects();
}

async function loadUsers() {
  const response = await fetch("/api/users");
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudieron cargar los usuarios");
  }
  users = data.users || [];
  userRoles = data.roles || [];
  renderUsers();
  if (userFormMode === "edit") {
    const currentEditingUser = users.find((user) => user.id === editingUserId);
    if (currentEditingUser) {
      startUserEdit(currentEditingUser);
    } else {
      resetUserForm();
    }
  } else {
    resetUserForm();
  }
}

async function loadAssignmentOptions() {
  const response = await fetch("/api/company-assignment-options");
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudieron cargar las asignaciones");
  }
  assignmentOptions = data;
  renderVisitsResponsibleFilters();
  renderMeetingParticipantsPicker();
  setSelectedParticipantUserIds(selectedParticipantUserIds);
}

async function updateVisitStatus(meetingId, status) {
  const response = await fetch(`/api/meetings/${meetingId}/status`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ status })
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo actualizar el estado de la visita");
  }
  return data.meeting;
}

async function hideClient(clientId) {
  const response = await fetch(`/api/clients/${clientId}/hide`, {
    method: "PATCH"
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo ocultar la compañía");
  }
  return data;
}

async function hideBranch(clientId, branchId) {
  const response = await fetch(`/api/clients/${clientId}/branches/${branchId}/hide`, {
    method: "PATCH"
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo ocultar la sucursal");
  }
  return data;
}

async function saveOpportunity(payload) {
  const endpoint = editingOpportunityId
    ? `/api/clients/${selectedClient.id}/opportunities/${editingOpportunityId}`
    : `/api/clients/${selectedClient.id}/opportunities`;
  const method = editingOpportunityId ? "PATCH" : "POST";
  const response = await fetch(endpoint, {
    method,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo guardar la oportunidad");
  }
  return data.opportunity;
}

async function saveFollowUp(opportunityId, followUpId, payload) {
  const endpoint = followUpId
    ? `/api/opportunities/${opportunityId}/follow-ups/${followUpId}`
    : `/api/opportunities/${opportunityId}/follow-ups`;
  const method = followUpId ? "PATCH" : "POST";
  const response = await fetch(endpoint, {
    method,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || "No se pudo guardar el seguimiento");
  }
  return data.followUp;
}

async function exportVisits() {
  const query = buildQuery(getVisitsFilters());
  const response = await fetch(`/api/visits/export${query ? `?${query}` : ""}`);
  if (!response.ok) {
    let message = "No se pudo exportar la vista de visitas";
    try {
      const data = await response.json();
      message = data?.error || message;
    } catch (_error) {
      // ignore non-json export error payloads
    }
    throw new Error(message);
  }

  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "visitas.xlsx";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function loadVisitsStats() {
  statsUsersStatus.textContent = "Cargando...";
  statsTypesStatus.textContent = "Cargando...";
  statsRulesStatus.textContent = "Cargando...";
  statsSemaphoreStatus.textContent = "Cargando...";

  try {
    const query = buildQuery(getVisitsStatsFilters());
    const response = await fetch(`/api/visits/stats${query ? `?${query}` : ""}`);
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data?.error || "No se pudieron cargar las estadísticas");
    }

    visitRulesFeatureEnabled = Boolean(data?.featureFlags?.visitRulesEnabled ?? visitRulesFeatureEnabled);
    applyFeatureFlagsUi();

    visitStats = {
      byUser: data.byUser || [],
      byType: data.byType || [],
      byRule: data.byRule || {
        totalCount: 0,
        completedCount: 0,
        completionRate: 0,
        semaphores: {
          totalRules: 0,
          whiteCount: 0,
          greenCount: 0,
          yellowCount: 0,
          redCount: 0,
          whiteRate: 0,
          greenRate: 0,
          yellowRate: 0,
          redRate: 0
        }
      }
    };
    renderVisitStats();
  } catch (error) {
    statsUsersStatus.textContent = "Error";
    statsTypesStatus.textContent = "Error";
    statsRulesStatus.textContent = "Error";
    statsSemaphoreStatus.textContent = "Error";
    notifyError(error.message);
  }
}

async function loadClients(options = {}) {
  const { clearSelectionWhenMissing = true } = options;
  visibleCount.textContent = "Cargando...";
  try {
    const query = buildQuery(getFilters());
    const response = await fetch(`/api/clients${query ? `?${query}` : ""}`);
    const data = await response.json();

    if (!response.ok) {
      if (response.status === 401) {
        currentUser = null;
        updateAuthUi();
        throw new Error("Tu sesión venció. Volvé a ingresar.");
      }
      throw new Error(data?.error || "No se pudo cargar la lista de clientes");
    }

    clients = data.clients;
    computeGlobalKpis(data.globalKpis);
    renderTable();

    if (clearSelectionWhenMissing && selectedId && !clients.some((client) => client.id === selectedId)) {
      selectedId = null;
      selectedClient = null;
      showScreen("list");
    }
  } catch (error) {
    console.error(error);
    visibleCount.textContent = "Error al cargar";
    notifyError("No pudimos cargar los clientes. Confirmá que el servidor local esté corriendo en localhost:3000.");
  }
}

async function loadCalendar() {
  try {
    const query = buildQuery({ month: monthKey(calendarDate), ...getCalendarFilters() });
    const response = await fetch(`/api/calendar?${query}`);
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data?.error || "No se pudo cargar el calendario");
    }
    calendarMeetings = data.meetings || [];
    renderCalendar();
  } catch (error) {
    console.error(error);
    notifyError("No pudimos cargar el calendario de reuniones.");
  }
}

async function loadVisits() {
  try {
    const response = await fetch(`/api/visits?${buildQuery(getVisitsFilters())}`);
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data?.error || "No se pudo cargar la vista de visitas");
    }
    visits = data.visits || [];
    renderVisitsTable();
  } catch (error) {
    console.error(error);
    visitsVisibleCount.textContent = "Error al cargar";
    notifyError("No pudimos cargar la vista de visitas.");
  }
}

async function loadVisitsGrid() {
  try {
    const response = await fetch("/api/visits?status=Realizada");
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data?.error || "No se pudo cargar la grilla de visitas");
    }
    visitsGridData = data.visits || [];
    renderVisitsGrid();
  } catch (error) {
    console.error(error);
    visitsGridWrap.innerHTML = `
      <article class="empty-meeting-state">
        <h4>Error al cargar la grilla</h4>
        <p>No pudimos cargar las visitas realizadas.</p>
      </article>
    `;
  }
}

function getPipelineFilters() {
  return {
    search: pipelineSearchInput.value.trim(),
    ownerUserId: pipelineOwnerFilter.value,
    status: pipelineStatusFilter.value
  };
}

async function loadPipeline() {
  try {
    const response = await fetch(`/api/pipeline?${buildQuery(getPipelineFilters())}`);
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data?.error || "No se pudo cargar el pipeline");
    }
    pipelineOpportunities = data.opportunities || [];
    pipelineSummaryData = data.summary || pipelineSummaryData;
    renderPipelineBoard();
  } catch (error) {
    console.error(error);
    pipelineBoard.innerHTML = `
      <article class="empty-meeting-state">
        <h4>Error al cargar el pipeline</h4>
        <p>No pudimos cargar las oportunidades comerciales.</p>
      </article>
    `;
  }
}

async function loadCurrentUser() {
  try {
    const response = await fetch("/api/auth/me");
    if (!response.ok) {
      currentUser = null;
      updateAuthUi();
      return false;
    }

    const data = await response.json();
    currentUser = data.user;
    visitRulesFeatureEnabled = Boolean(data?.featureFlags?.visitRulesEnabled ?? true);
    updateAuthUi();
    applyFeatureFlagsUi();
    return true;
  } catch (error) {
    console.error("Error cargando sesión", error);
    currentUser = null;
    updateAuthUi();
    return false;
  }
}

async function loadClientDetail(clientId) {
  try {
    const response = await fetch(`/api/clients/${clientId}`);
    if (!response.ok) {
      if (response.status === 401) {
        currentUser = null;
        updateAuthUi();
      }
      selectedClient = null;
      console.error("Detalle no disponible", clientId, response.status);
      return false;
    }
    const data = await response.json();
    selectedClient = data.client;
    renderDetail();
    return true;
  } catch (error) {
    console.error("Error cargando detalle del cliente", clientId, error);
    selectedClient = null;
    return false;
  }
}

const filters = [riskFilter, segmentFilter, fixedFireFilter, extinguishersFilter, worksFilter];

filters.forEach((input) => {
  input.addEventListener("change", async () => {
    await loadClients();
  });
});

searchInput.addEventListener("input", () => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => {
    loadClients();
  }, 250);
});

backToList.addEventListener("click", () => {
  if (selectedBranchView) {
    selectedBranchView = null;
    renderDetail();
    syncUrlWithState("detail");
    return;
  }
  showScreen("list");
  syncUrlWithState("list");
});

backFromEdit.addEventListener("click", () => {
  if (selectedClient) {
    showScreen("detail");
    syncUrlWithState("detail");
    return;
  }
  showScreen("list");
  syncUrlWithState("list");
});

backFromMeeting.addEventListener("click", () => {
  showScreen("detail");
  syncUrlWithState("detail");
});

editMeetingFromDetailBtn.addEventListener("click", () => {
  const meetings = Array.isArray(selectedClient?.meetings) ? selectedClient.meetings : [];
  const meeting = meetings.find((item) => Number(item.id) === Number(editingMeetingId));
  if (!meeting) return;
  openMeetingScreen(meeting);
});

goToEditBtn.addEventListener("click", () => {
  if (selectedBranchView) {
    openBranchEditScreen(selectedBranchView);
    return;
  }
  openEditScreen();
});

goToMeetingBtn.addEventListener("click", () => {
  detailSection = "visits";
  renderDetailSectionNav();
  openMeetingScreen();
});

showClientsBtn.addEventListener("click", () => {
  selectedBranchView = null;
  showScreen("list");
  syncUrlWithState("list");
});

showVisitsBtn.addEventListener("click", async () => {
  await loadVisits();
  showScreen("visits");
  syncUrlWithState("visits");
});

showPipelineBtn.addEventListener("click", async () => {
  if (!CRM_ENABLED) return;
  await loadPipeline();
  showScreen("pipeline");
  syncUrlWithState("pipeline");
});

openPipelineFromCrmBtn.addEventListener("click", async () => {
  if (!CRM_ENABLED) return;
  await loadPipeline();
  showScreen("pipeline");
  syncUrlWithState("pipeline");
});

showVisitsGridBtn.addEventListener("click", async () => {
  await loadVisitsGrid();
  showScreen("visits-grid");
});

showCalendarBtn.addEventListener("click", async () => {
  calendarView = "month";
  selectedCalendarDay = null;
  await loadCalendar();
  showScreen("calendar");
  syncUrlWithState("calendar");
});

showVisitsStatsBtn.addEventListener("click", async () => {
  await loadVisitsStats();
  showScreen("visits-stats");
  syncUrlWithState("visits-stats");
});

showSettingsBtn.addEventListener("click", async () => {
  if (!canAccessSettings()) return;
  await loadSettingsCatalogs();
  showSettingsTab("catalogs");
  showScreen("settings");
  syncUrlWithState("settings");
});

openUsersFromSettingsBtn.addEventListener("click", async () => {
  if (!canAccessSettings()) return;
  await loadUsers();
  showScreen("users");
  syncUrlWithState("users");
});

settingsTabButtons.forEach((button) => {
  button.addEventListener("click", () => {
    showSettingsTab(button.dataset.settingsTab);
  });
});

backToSettingsBtn.addEventListener("click", async () => {
  if (!canAccessSettings()) {
    showScreen("list");
    syncUrlWithState("list");
    return;
  }
  await loadSettingsCatalogs();
  showSettingsTab("catalogs");
  showScreen("settings");
  syncUrlWithState("settings");
});

downloadClientsTemplateBtn.addEventListener("click", async () => {
  try {
    clearImportFeedback(clientsImportStatus, clientsImportDetails);
    downloadClientsTemplateBtn.disabled = true;
    await downloadClientsTemplate();
  } catch (error) {
    clientsImportStatus.textContent = error.message;
  } finally {
    downloadClientsTemplateBtn.disabled = false;
  }
});

downloadBranchesTemplateBtn.addEventListener("click", async () => {
  try {
    clearImportFeedback(branchesImportStatus, branchesImportDetails);
    downloadBranchesTemplateBtn.disabled = true;
    await downloadBranchesTemplate();
  } catch (error) {
    branchesImportStatus.textContent = error.message;
  } finally {
    downloadBranchesTemplateBtn.disabled = false;
  }
});

importClientsBtn.addEventListener("click", async () => {
  const file = clientsImportFile.files?.[0];
  if (!file) {
    clientsImportStatus.textContent = "Seleccioná un archivo Excel antes de importar.";
    clientsImportDetails.classList.add("hidden");
    return;
  }

  try {
    clearImportFeedback(clientsImportStatus, clientsImportDetails);
    importClientsBtn.disabled = true;
    clientsImportStatus.textContent = "Importando clientes...";
    const result = await importClientsFromExcel(file);
    renderImportFeedback(clientsImportStatus, clientsImportDetails, result, "clientes");
    await loadClients({ clearSelectionWhenMissing: false });
  } catch (error) {
    clientsImportStatus.textContent = error.message;
    clientsImportDetails.classList.add("hidden");
  } finally {
    importClientsBtn.disabled = false;
  }
});

importBranchesBtn.addEventListener("click", async () => {
  const file = branchesImportFile.files?.[0];
  if (!file) {
    branchesImportStatus.textContent = "Seleccioná un archivo Excel antes de importar.";
    branchesImportDetails.classList.add("hidden");
    return;
  }

  try {
    clearImportFeedback(branchesImportStatus, branchesImportDetails);
    importBranchesBtn.disabled = true;
    branchesImportStatus.textContent = "Importando sucursales...";
    const result = await importBranchesFromExcel(file);
    renderImportFeedback(branchesImportStatus, branchesImportDetails, result, "sucursales");
    await loadClients({ clearSelectionWhenMissing: false });
  } catch (error) {
    branchesImportStatus.textContent = error.message;
    branchesImportDetails.classList.add("hidden");
  } finally {
    importBranchesBtn.disabled = false;
  }
});

downloadUsersTemplateBtn.addEventListener("click", async () => {
  try {
    clearImportFeedback(usersImportStatus, usersImportDetails);
    downloadUsersTemplateBtn.disabled = true;
    await downloadUsersTemplate();
  } catch (error) {
    usersImportStatus.textContent = error.message;
  } finally {
    downloadUsersTemplateBtn.disabled = false;
  }
});

importUsersBtn.addEventListener("click", async () => {
  const file = usersImportFile.files?.[0];
  if (!file) {
    usersImportStatus.textContent = "Seleccioná un archivo Excel antes de importar.";
    usersImportDetails.classList.add("hidden");
    return;
  }

  try {
    clearImportFeedback(usersImportStatus, usersImportDetails);
    importUsersBtn.disabled = true;
    usersImportStatus.textContent = "Importando usuarios...";
    const result = await importUsersFromExcel(file);
    renderImportFeedback(usersImportStatus, usersImportDetails, result, "usuarios");
    await loadUsers();
    await loadAssignmentOptions();
  } catch (error) {
    usersImportStatus.textContent = error.message;
    usersImportDetails.classList.add("hidden");
  } finally {
    importUsersBtn.disabled = false;
  }
});

applyStatsFiltersBtn.addEventListener("click", async () => {
  await loadVisitsStats();
});

resetStatsFiltersBtn.addEventListener("click", async () => {
  statsExecutiveFilter.value = "todos";
  statsDateFrom.value = "";
  statsDateTo.value = "";
  await loadVisitsStats();
});

statsExecutiveFilter.addEventListener("change", async () => {
  await loadVisitsStats();
});

document.querySelectorAll(".table-sort-btn").forEach((button) => {
  button.addEventListener("click", () => {
    const nextKey = button.dataset.sortKey;
    if (!nextKey) return;
    if (visitsSort.key === nextKey) {
      visitsSort.direction = visitsSort.direction === "asc" ? "desc" : "asc";
    } else {
      visitsSort.key = nextKey;
      visitsSort.direction = nextKey === "scheduledFor" ? "desc" : "asc";
    }
    renderVisitsTable();
  });
});

visitsSearchInput.addEventListener("input", () => {
  clearTimeout(visitsSearchTimer);
  visitsSearchTimer = setTimeout(() => {
    loadVisits();
  }, 250);
});

[visitsStatusFilter, visitsTypeFilter, visitsModalityFilter, visitsDateFromFilter, visitsDateToFilter].forEach((input) => {
  input.addEventListener("change", () => {
    loadVisits();
  });
});

[visitsExecutiveFilter, visitsSupervisorFilter].forEach((input) => {
  input.addEventListener("change", () => {
    loadVisits();
  });
});

visitsParticipantFilter.addEventListener("change", () => {
  loadVisits();
});

exportVisitsBtn.addEventListener("click", async () => {
  try {
    exportVisitsBtn.disabled = true;
    await exportVisits();
  } catch (error) {
    notifyError(error.message);
  } finally {
    exportVisitsBtn.disabled = false;
  }
});

hideCompanyBtn.addEventListener("click", async () => {
  if (!selectedClient || editMode !== "edit") return;
  const isBranch = editEntityType === "branch" && selectedBranchView;
  const targetName = isBranch ? selectedBranchView?.name : selectedClient.name;
  const targetLabel = isBranch ? "sucursal" : "compañía";
  const confirmed = window.confirm(`La ${targetLabel} ${targetName} se va a ocultar del listado, pero no se va a borrar.`);
  if (!confirmed) return;

  try {
    hideCompanyBtn.disabled = true;
    if (isBranch) {
      await hideBranch(selectedClient.id, selectedBranchView.id);
      selectedBranchView = null;
      await loadClients({ clearSelectionWhenMissing: false });
      await loadClientDetail(selectedClient.id);
      renderTable();
      showScreen("detail");
      syncUrlWithState("detail");
    } else {
      await hideClient(selectedClient.id);
      selectedId = null;
      selectedClient = null;
      selectedBranchView = null;
      await loadClients({ clearSelectionWhenMissing: false });
      showScreen("list");
      syncUrlWithState("list");
    }
    editCompanyStatus.textContent = "";
  } catch (error) {
    editCompanyStatus.textContent = error.message;
  } finally {
    hideCompanyBtn.disabled = false;
  }
});

prevMonthBtn.addEventListener("click", async () => {
  calendarDate = new Date(calendarDate.getFullYear(), calendarDate.getMonth() - 1, 1);
  await loadCalendar();
});

nextMonthBtn.addEventListener("click", async () => {
  calendarDate = new Date(calendarDate.getFullYear(), calendarDate.getMonth() + 1, 1);
  await loadCalendar();
});

calendarParticipantFilter.addEventListener("change", async () => {
  calendarView = "month";
  selectedCalendarDay = null;
  await loadCalendar();
});

meetingParticipantsSearch.addEventListener("input", () => {
  renderMeetingParticipantsPicker();
});

backToMonthBtn.addEventListener("click", () => {
  calendarView = "month";
  selectedCalendarDay = null;
  renderCalendar();
});

addCompanyBtn.addEventListener("click", () => {
  openCreateScreen();
});

addVisitRuleBtn.addEventListener("click", () => {
  currentVisitRules.push(createEmptyVisitRule());
  renderVisitRulesEditor();
});

addBranchBtn.addEventListener("click", () => {
  openBranchCreateScreen();
});

newOpportunityBtn.addEventListener("click", () => {
  if (!CRM_ENABLED) return;
  detailSection = "opportunities";
  renderDetailSectionNav();
  openOpportunityForm();
});

showDetailOpportunitiesBtn.addEventListener("click", () => {
  if (!CRM_ENABLED) return;
  detailSection = "opportunities";
  renderDetailSectionNav();
});

showDetailVisitsBtn.addEventListener("click", () => {
  detailSection = "visits";
  renderDetailSectionNav();
});

newUserBitrixSearch.addEventListener("change", syncBitrixUserSelectionFromSearch);
newUserBitrixSearch.addEventListener("input", () => {
  renderBitrixSearchDropdown(
    newUserBitrixOptions,
    bitrixDirectory.users,
    newUserBitrixSearch.value,
    getBitrixUsersEmptyText(newUserBitrixSearch.value)
  );
  if (!newUserBitrixSearch.value.trim()) {
    newUserBitrixId.value = "";
  }
});
newUserBitrixSearch.addEventListener("focus", () => {
  renderBitrixSearchDropdown(
    newUserBitrixOptions,
    bitrixDirectory.users,
    newUserBitrixSearch.value,
    getBitrixUsersEmptyText(newUserBitrixSearch.value)
  );
});

editBitrixCompanySearch.addEventListener("change", syncBitrixCompanySelectionFromSearch);
editBitrixCompanySearch.addEventListener("input", async () => {
  renderBitrixSearchDropdown(
    editBitrixCompanyOptions,
    bitrixCompanySearchResults.length ? bitrixCompanySearchResults : bitrixDirectory.companies,
    editBitrixCompanySearch.value,
    getBitrixCompaniesEmptyText(editBitrixCompanySearch.value)
  );
  if (!editBitrixCompanySearch.value.trim()) {
    editBitrixCompanyId.value = "";
    bitrixCompanySearchResults = [];
    renderBitrixDirectoryDatalists();
    return;
  }
  await searchBitrixCompaniesRemote(editBitrixCompanySearch.value);
});
editBitrixCompanySearch.addEventListener("focus", () => {
  bitrixCompanySearchResults = [];
  renderBitrixSearchDropdown(
    editBitrixCompanyOptions,
    bitrixDirectory.companies,
    editBitrixCompanySearch.value,
    getBitrixCompaniesEmptyText(editBitrixCompanySearch.value)
  );
});

newUserBitrixOptions.addEventListener("click", (event) => {
  const option = event.target.closest(".bitrix-search-option");
  if (!option) return;

  newUserBitrixSearch.value = option.dataset.label || "";
  newUserBitrixId.value = option.dataset.id || "";
  closeBitrixSearchDropdown(newUserBitrixOptions);
});

editBitrixCompanyOptions.addEventListener("click", (event) => {
  const option = event.target.closest(".bitrix-search-option");
  if (!option) return;

  editBitrixCompanySearch.value = option.dataset.label || "";
  editBitrixCompanyId.value = option.dataset.id || "";
  mergeBitrixCompanies([
    {
      id: option.dataset.id || "",
      label: option.dataset.label || "",
      title: option.dataset.label ? option.dataset.label.split(" · ")[0] : option.dataset.label || ""
    }
  ]);
  bitrixCompanySearchResults = [];
  closeBitrixSearchDropdown(editBitrixCompanyOptions);
});

document.addEventListener("click", (event) => {
  if (!newUserBitrixSearch.contains(event.target) && !newUserBitrixOptions.contains(event.target)) {
    closeBitrixSearchDropdown(newUserBitrixOptions);
  }

  if (!editBitrixCompanySearch.contains(event.target) && !editBitrixCompanyOptions.contains(event.target)) {
    closeBitrixSearchDropdown(editBitrixCompanyOptions);
  }
});

cancelOpportunityEditBtn.addEventListener("click", () => {
  resetOpportunityForm();
});

cancelFollowUpEditBtn.addEventListener("click", () => {
  resetFollowUpForm();
});

opportunityStatusSelect.addEventListener("change", () => {
  syncOpportunityLossReasonVisibility();
});

pipelineSearchInput.addEventListener("input", () => {
  clearTimeout(pipelineSearchTimer);
  pipelineSearchTimer = setTimeout(() => {
    loadPipeline();
  }, 250);
});

[pipelineOwnerFilter, pipelineStatusFilter].forEach((input) => {
  input.addEventListener("change", () => {
    loadPipeline();
  });
});

userCreateForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  syncBitrixUserSelectionFromSearch();
  createUserBtn.disabled = true;
  createUserStatus.textContent = userFormMode === "edit" ? "Guardando..." : "Creando...";

  try {
    const payload = {
      name: newUserName.value.trim(),
      email: newUserEmail.value.trim(),
      role: newUserRole.value,
      bitrixUserId: newUserBitrixId.value.trim(),
      permissions: getUserPermissionsPayload(),
      password: newUserPassword.value
    };

    if (userFormMode === "edit") {
      if (!payload.password) delete payload.password;
      await updateUser(editingUserId, payload);
      createUserStatus.textContent = "Usuario actualizado";
    } else {
      await createUser(payload);
      createUserStatus.textContent = "Usuario creado";
    }

    await loadUsers();
    await loadAssignmentOptions();
    resetUserForm();
  } catch (error) {
    createUserStatus.textContent = error.message;
  } finally {
    createUserBtn.disabled = false;
  }
});

cancelUserEditBtn.addEventListener("click", () => {
  createUserStatus.textContent = "";
  resetUserForm();
});

saveBitrixWebhookBtn.addEventListener("click", async () => {
  saveBitrixWebhookBtn.disabled = true;
  bitrixUsersStatus.textContent = "Guardando webhook...";

  try {
    const result = await saveBitrixConfig(bitrixWebhookUrl.value.trim());
    bitrixWebhookUrl.value = result.webhookUrl || bitrixWebhookUrl.value.trim();
    try {
      const bitrixResponse = await fetch("/api/bitrix/options");
      const bitrixData = await bitrixResponse.json();
      if (bitrixResponse.ok) {
        bitrixDirectory = {
          users: bitrixData.users || [],
          companies: bitrixData.companies || []
        };
        bitrixCompanySearchResults = [];
        bitrixDirectoryPermissions = bitrixData.permissions || {
          users: true,
          companies: true
        };
        bitrixDirectoryErrors = bitrixData.errors || {
          users: "",
          companies: ""
        };
        renderBitrixDirectoryDatalists();
      }
    } catch (_error) {
      // No bloqueamos la pantalla por un refresh fallido del directorio.
    }
    bitrixUsersStatus.textContent = "Webhook guardado correctamente";
  } catch (error) {
    bitrixUsersStatus.textContent = error.message;
  } finally {
    saveBitrixWebhookBtn.disabled = false;
  }
});

bitrixUsersPreviewForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  loadBitrixUsersBtn.disabled = true;
  bitrixUsersStatus.textContent = "Consultando Bitrix...";

  try {
    const webhook = bitrixWebhookUrl.value.trim();
    const savedConfig = await saveBitrixConfig(webhook);
    const normalizedWebhook = savedConfig.webhookUrl || webhook;
    bitrixWebhookUrl.value = normalizedWebhook;
    const [usersResult, mappingsResult] = await Promise.all([
      previewBitrixUsers(normalizedWebhook),
      previewBitrixMappings(normalizedWebhook)
    ]);

    bitrixPreviewUsers = Array.isArray(usersResult.users) ? usersResult.users : [];
    bitrixUserMappings = Array.isArray(mappingsResult.userMappings) ? mappingsResult.userMappings : [];
    bitrixClientMappings = Array.isArray(mappingsResult.clientMappings) ? mappingsResult.clientMappings : [];
    bitrixCounts = mappingsResult.counts || bitrixCounts;
    bitrixPermissions = mappingsResult.permissions || bitrixPermissions;
    bitrixErrors = mappingsResult.errors || bitrixErrors;
    bitrixDirectory = {
      users: (bitrixPreviewUsers || [])
        .filter((user) => user.active)
        .map((user) => ({
          id: String(user.id || ""),
          label: `${user.fullName}${user.email ? ` · ${user.email}` : ""}${user.id ? ` · ${user.id}` : ""}`.trim(),
          name: user.fullName,
          email: user.email || ""
        })),
      companies: Array.isArray(mappingsResult.bitrixCompanies)
        ? mappingsResult.bitrixCompanies.map((company) => ({
            id: String(company.id || ""),
            label: `${company.title}${company.id ? ` · ${company.id}` : ""}`.trim(),
            title: company.title
          }))
        : bitrixDirectory.companies
    };
    bitrixCompanySearchResults = [];
    bitrixDirectoryPermissions = {
      users: true,
      companies: Boolean(bitrixPermissions.companies)
    };
    bitrixDirectoryErrors = {
      users: "",
      companies: bitrixErrors.companies || ""
    };
    renderBitrixDirectoryDatalists();

    renderBitrixUsersPreview();
    renderBitrixMappings();
    const warnings = [];
    if (!bitrixPermissions.companies && bitrixErrors.companies) warnings.push(`Compañías: ${bitrixErrors.companies}`);
    if (!bitrixPermissions.leads && bitrixErrors.leads) warnings.push(`Leads: ${bitrixErrors.leads}`);
    bitrixUsersStatus.textContent = warnings.length
      ? `${usersResult.count || bitrixPreviewUsers.length} usuarios Bitrix cargados. CRM sin permiso: ${warnings.join(" · ")}`
      : `${usersResult.count || bitrixPreviewUsers.length} usuarios Bitrix y mapeos actualizados`;
  } catch (error) {
    bitrixPreviewUsers = [];
    bitrixUserMappings = [];
    bitrixClientMappings = [];
    bitrixCounts = {
      localUsers: 0,
      localClients: 0,
      bitrixUsers: 0,
      bitrixCompanies: 0,
      bitrixLeads: 0
    };
    bitrixPermissions = {
      users: false,
      companies: false,
      leads: false
    };
    bitrixErrors = {
      companies: "",
      leads: ""
    };
    renderBitrixUsersPreview();
    renderBitrixMappings();
    bitrixUsersStatus.textContent = error.message;
  } finally {
    loadBitrixUsersBtn.disabled = false;
  }
});

bitrixTaskTestForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  createBitrixTaskBtn.disabled = true;
  bitrixTaskStatus.textContent = "Creando tarea en Bitrix...";

  try {
    const result = await createBitrixTaskTest({
      webhookUrl: bitrixWebhookUrl.value.trim(),
      title: bitrixTaskTitle.value.trim(),
      description: bitrixTaskDescription.value.trim(),
      responsibleId: bitrixTaskResponsible.value
    });
    bitrixTaskStatus.textContent = result.task?.id
      ? `Tarea creada en Bitrix con ID ${result.task.id}`
      : "Tarea creada en Bitrix";
  } catch (error) {
    bitrixTaskStatus.textContent = error.message;
  } finally {
    createBitrixTaskBtn.disabled = false;
  }
});

visitRulesFeatureForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  saveVisitRulesFeatureBtn.disabled = true;
  visitRulesFeatureStatus.textContent = "Guardando...";

  try {
    const result = await saveFeatureSettings({
      visitRulesEnabled: visitRulesFeatureToggle.checked
    });
    visitRulesFeatureEnabled = Boolean(result?.featureFlags?.visitRulesEnabled ?? visitRulesFeatureToggle.checked);
    applyFeatureFlagsUi();
    visitRulesFeatureStatus.textContent = visitRulesFeatureEnabled
      ? "Reglas automáticas activadas"
      : "Reglas automáticas desactivadas";
  } catch (error) {
    visitRulesFeatureStatus.textContent = error.message;
  } finally {
    saveVisitRulesFeatureBtn.disabled = false;
  }
});

sectorForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  createSectorBtn.disabled = true;
  sectorStatus.textContent = "Guardando...";

  try {
    await createSector(newSectorName.value.trim());
    newSectorName.value = "";
    sectorStatus.textContent = "Sector agregado";
    await loadSectorOptions();
  } catch (error) {
    sectorStatus.textContent = error.message;
  } finally {
    createSectorBtn.disabled = false;
  }
});

meetingTypeForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  saveMeetingTypeBtn.disabled = true;
  meetingTypeStatus.textContent = "Guardando...";

  try {
    await saveMeetingType({
      value: meetingTypeValue.value.trim(),
      label: meetingTypeLabel.value.trim(),
      color: meetingTypeColor.value
    });
    await loadMeetingTypesConfig();
    resetMeetingTypeForm();
    meetingTypeStatus.textContent = "Tipo guardado";
  } catch (error) {
    meetingTypeStatus.textContent = error.message;
  } finally {
    saveMeetingTypeBtn.disabled = false;
  }
});

cancelMeetingTypeEditBtn.addEventListener("click", () => {
  resetMeetingTypeForm();
});

meetingReasonForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  saveMeetingReasonBtn.disabled = true;
  meetingReasonStatus.textContent = "Guardando...";

  try {
    await saveMeetingReason({
      name: meetingReasonName.value.trim()
    });
    await loadMeetingReasonsConfig();
    resetMeetingReasonForm();
    meetingReasonStatus.textContent = "Motivo guardado";
  } catch (error) {
    meetingReasonStatus.textContent = error.message;
  } finally {
    saveMeetingReasonBtn.disabled = false;
  }
});

cancelMeetingReasonEditBtn.addEventListener("click", () => {
  resetMeetingReasonForm();
});

contactRoleForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  saveContactRoleBtn.disabled = true;
  contactRoleStatus.textContent = "Guardando...";

  try {
    await saveContactRole({
      name: contactRoleName.value.trim()
    });
    await loadContactRolesConfig();
    resetContactRoleForm();
    contactRoleStatus.textContent = "Función guardada";
  } catch (error) {
    contactRoleStatus.textContent = error.message;
  } finally {
    saveContactRoleBtn.disabled = false;
  }
});

cancelContactRoleEditBtn.addEventListener("click", () => {
  resetContactRoleForm();
});

[editFixedFire, editExtinguishers, editWorks].forEach((checkbox) => {
  checkbox.addEventListener("change", () => {
    syncSupervisorVisibility();
    renderAccountRoleSummary();
  });
});

[editManager, editSupervisorFixedFire, editSupervisorExtinguishers, editSupervisorWorks].forEach((select) => {
  select.addEventListener("change", () => {
    renderAccountRoleSummary();
  });
});

meetingStatusSelect.addEventListener("change", () => {
  syncMeetingCompletionFields();
});

meetingScope.addEventListener("change", () => {
  syncMeetingContextBlocks();
});

opportunityForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!selectedClient) return;

  const amountResult = parseNumericInput(opportunityAmount, { min: 0 });
  const probabilityResult = parseNumericInput(opportunityProbability, { integer: true, min: 0, max: 100 });

  if (!amountResult.ok || !probabilityResult.ok) {
    opportunityStatusMessage.textContent = "Completá monto y probabilidad con valores válidos.";
    return;
  }

  saveOpportunityBtn.disabled = true;
  opportunityStatusMessage.textContent = "Guardando oportunidad...";

  try {
    await saveOpportunity({
      title: opportunityTitle.value.trim(),
      opportunityType: opportunityType.value,
      serviceLine: opportunityServiceLine.value,
      status: opportunityStatusSelect.value,
      amount: amountResult.value,
      probability: probabilityResult.value,
      expectedCloseDate: opportunityExpectedCloseDate.value,
      ownerUserId: opportunityOwnerUserId.value ? Number(opportunityOwnerUserId.value) : null,
      branchId: opportunityBranchId.value ? Number(opportunityBranchId.value) : null,
      source: opportunitySource.value.trim(),
      description: opportunityDescription.value.trim(),
      lossReason: opportunityLossReason.value.trim()
    });
    await loadClientDetail(selectedClient.id);
    await loadClients({ clearSelectionWhenMissing: false });
    await loadPipeline();
    renderTable();
    resetOpportunityForm();
  } catch (error) {
    opportunityStatusMessage.textContent = error.message;
  } finally {
    saveOpportunityBtn.disabled = false;
  }
});

followUpForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!selectedClient || !followUpOpportunityId) return;

  saveFollowUpBtn.disabled = true;
  followUpStatusMessage.textContent = "Guardando seguimiento...";

  try {
    await saveFollowUp(followUpOpportunityId, editingFollowUpId, {
      type: followUpType.value,
      title: followUpTitle.value.trim(),
      dueDate: followUpDueDate.value,
      assignedUserId: followUpAssignedUserId.value ? Number(followUpAssignedUserId.value) : null,
      status: followUpStatusSelect.value,
      notes: followUpNotes.value.trim()
    });
    await loadClientDetail(selectedClient.id);
    await loadClients({ clearSelectionWhenMissing: false });
    await loadPipeline();
    renderTable();
    resetFollowUpForm();
  } catch (error) {
    followUpStatusMessage.textContent = error.message;
  } finally {
    saveFollowUpBtn.disabled = false;
  }
});

companyEditForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  syncBitrixCompanySelectionFromSearch();

  if (!editSector.value) {
    editCompanyStatus.textContent = "Seleccioná un sector válido.";
    return;
  }
  if (editEntityType !== "branch" && !editCompanyType.value) {
    editCompanyStatus.textContent = "Seleccioná si la compañía es local o global.";
    return;
  }
  if (editEntityType !== "branch" && !editCountry.value) {
    editCompanyStatus.textContent = "Seleccioná el país de la compañía.";
    return;
  }
  if (editEntityType !== "branch" && !editAccountStage.value) {
    editCompanyStatus.textContent = "Seleccioná si la compañía es activa o prospecto.";
    return;
  }
  if (editFixedFire.checked && !editSupervisorFixedFire.value) {
    editCompanyStatus.textContent = "Debes asignar un supervisor IFCI.";
    return;
  }
  if (editExtinguishers.checked && !editSupervisorExtinguishers.value) {
    editCompanyStatus.textContent = "Debes asignar un supervisor EXT.";
    return;
  }
  if (editWorks.checked && !editSupervisorWorks.value) {
    editCompanyStatus.textContent = "Debes asignar un supervisor Obra.";
    return;
  }

  const payload = {
    name: editName.value.trim(),
    sector: editSector.value.trim(),
    risk: editRisk.value,
    segment: editSegment.value,
    services: {
      fixedFire: editFixedFire.checked,
      extinguishers: editExtinguishers.checked,
      works: editWorks.checked
    },
    supervisors: {
      fixedFire: { userId: editSupervisorFixedFire.value ? Number(editSupervisorFixedFire.value) : null },
      extinguishers: { userId: editSupervisorExtinguishers.value ? Number(editSupervisorExtinguishers.value) : null },
      works: { userId: editSupervisorWorks.value ? Number(editSupervisorWorks.value) : null }
    },
    notes: editNotes.value.trim()
  };

  if (visitRulesFeatureEnabled) {
    payload.visitRules = currentVisitRules.map((rule) => ({
      periodicityDays: Number(rule.periodicityDays || 0),
      contactRole: String(rule.contactRole || "").trim(),
      objective: String(rule.objective || "").trim()
    }));

    if (
      payload.visitRules.some(
        (rule) => !Number.isInteger(rule.periodicityDays) || rule.periodicityDays <= 0 || !rule.contactRole || !rule.objective
      )
    ) {
      editCompanyStatus.textContent = "Completá correctamente todas las reglas de visita.";
      return;
    }
  }

  if (editEntityType !== "branch") {
    payload.companyType = editCompanyType.value;
    payload.country = editCountry.value;
    payload.accountStage = editAccountStage.value;
    payload.bitrixLeadId = editBitrixLeadId.value.trim();
    payload.bitrixCompanyId = editBitrixCompanyId.value.trim();
    payload.executiveUserId = editManager.value ? Number(editManager.value) : null;
  } else {
    payload.bitrixCompanyId = editBitrixCompanyId.value.trim();
    payload.manualBranchId = editManualBranchId.value.trim();
  }

  saveCompanyBtn.disabled = true;
  editCompanyStatus.textContent = "Guardando...";

  try {
    const isCreate = editMode === "create";
    const isBranch = editEntityType === "branch";
    const endpoint = isBranch
      ? isCreate
        ? `/api/clients/${selectedClient.id}/branches`
        : `/api/clients/${selectedClient.id}/branches/${editingBranchId}`
      : isCreate
        ? "/api/clients"
        : `/api/clients/${selectedClient.id}`;
    const method = isCreate ? "POST" : "PATCH";

    const response = await fetch(endpoint, {
      method,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const responseData = await response.json();
    if (!response.ok) {
      throw new Error(responseData?.error || `No se pudo guardar la ${isBranch ? "sucursal" : "compañía"}`);
    }

    selectedId = selectedClient?.id || responseData.client?.id;
    await loadClients({ clearSelectionWhenMissing: false });
    await loadClientDetail(selectedId);
    renderTable();
    editCompanyStatus.textContent = "Cambios guardados correctamente";
    showScreen("detail");
    syncUrlWithState("detail");
  } catch (error) {
    editCompanyStatus.textContent = error.message;
  } finally {
    saveCompanyBtn.disabled = false;
  }
});

meetingForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!selectedClient) return;

  const participantUserIds = normalizeParticipantIds(selectedParticipantUserIds);
  const participantNames = getParticipantNamesFromIds(participantUserIds);

  const payload = {
    kind: meetingKind.value,
    subject: meetingSubject.value.trim(),
    objective: meetingObjective.value.trim(),
    scheduledFor: meetingDate.value,
    modality: meetingModality.value,
    branchId: meetingScope.value.startsWith("branch:") ? Number(meetingScope.value.split(":")[1]) : null,
    opportunityId: meetingOpportunity.value ? Number(meetingOpportunity.value) : null,
    participants: participantNames.join(", "),
    participantUserIds,
    contactName: meetingContactName.value.trim(),
    contactRole: meetingContactRole.value.trim(),
    nextMeetingDate: meetingNextDate.value,
    status: meetingStatusSelect.value,
    minutes: meetingMinutes.value.trim(),
    findings: meetingNegotiationStatus.value.trim(),
    activeNegotiationsStatus: meetingNegotiationStatus.value.trim(),
    opportunities: meetingOpportunities.value.trim(),
    substituteRecovery: meetingSubstituteRecovery.value.trim(),
    globalContacts: meetingGlobalContacts.value.trim(),
    serviceHealthStatus: meetingServiceHealthStatus.value.trim(),
    serviceStatus: meetingServiceStatus.value.trim(),
    complaintBitrixResponsibleId: meetingComplaintBitrixResponsible.value
  };

  if (
    !payload.kind ||
    !payload.subject ||
    !payload.objective ||
    !payload.scheduledFor ||
    !payload.modality ||
    !meetingScope.value ||
    !payload.participants ||
    !payload.contactName ||
    !payload.contactRole
  ) {
    meetingStatus.textContent = "Completá tipo, motivo, objetivo, fecha, modalidad, alcance, participantes, contacto y función.";
    return;
  }

  if (payload.serviceStatus && !payload.complaintBitrixResponsibleId) {
    meetingStatus.textContent = "Seleccioná un responsable Bitrix para el reclamo.";
    return;
  }

  saveMeetingBtn.disabled = true;
  meetingStatus.textContent = "Guardando reunión...";

  try {
    const endpoint =
      meetingMode === "create"
        ? `/api/clients/${selectedClient.id}/meetings`
        : `/api/clients/${selectedClient.id}/meetings/${editingMeetingId}`;
    const method = meetingMode === "create" ? "POST" : "PATCH";

    const response = await fetch(endpoint, {
      method,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const responseData = await response.json();
    if (!response.ok) {
      throw new Error(responseData?.error || "No se pudo guardar la reunión");
    }

    await loadClients({ clearSelectionWhenMissing: false });
    await loadClientDetail(selectedClient.id);
    await loadCalendar();
    renderTable();
    meetingStatus.textContent = "Reunión guardada correctamente";
    showScreen("detail");
    syncUrlWithState("detail");
  } catch (error) {
    meetingStatus.textContent = error.message;
  } finally {
    saveMeetingBtn.disabled = false;
  }
});

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  loginSubmit.disabled = true;
  loginStatus.textContent = "Ingresando...";

  try {
    const response = await fetch("/api/auth/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        email: loginEmail.value.trim(),
        password: loginPassword.value,
        rememberMe: loginRemember.checked
      })
    });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data?.error || "No se pudo iniciar sesión");
    }

    currentUser = data.user;
    visitRulesFeatureEnabled = Boolean(data?.featureFlags?.visitRulesEnabled ?? true);
    updateAuthUi();
    applyFeatureFlagsUi();
    loginStatus.textContent = "";
    loginPassword.value = "";
    showScreen("list");
    await loadCatalogs();
    await loadCrmCatalogs();
    await loadAssignmentOptions();
    if (canAccessSettings()) {
      await loadSettingsCatalogs();
      await loadUsers();
    }
    await loadClients();
    await loadCalendar();
    await applyRouteFromHash();
  } catch (error) {
    loginStatus.textContent = error.message;
  } finally {
    loginSubmit.disabled = false;
  }
});

logoutBtn.addEventListener("click", async () => {
  await fetch("/api/auth/logout", { method: "POST" });
  currentUser = null;
  visitRulesFeatureEnabled = true;
  selectedId = null;
  selectedClient = null;
  clients = [];
  users = [];
  calendarMeetings = [];
  pipelineOpportunities = [];
  resetOpportunityForm();
  applyFeatureFlagsUi();
  resetFollowUpForm();
  updateAuthUi();
  showScreen("list");
  replaceHash("#/clientes");
});

window.addEventListener("hashchange", async () => {
  if (suppressHashRouting) return;
  await applyRouteFromHash();
});

showScreen("list");
updateAuthUi();
renderMeetingTypeColorOptions();
loadCurrentUser().then(async (loggedIn) => {
  if (loggedIn) {
    await loadCatalogs();
    await loadCrmCatalogs();
    await loadAssignmentOptions();
    if (canAccessSettings()) {
      await loadSettingsCatalogs();
      await loadUsers();
    }
    await loadClients();
    await loadCalendar();
    await applyRouteFromHash();
  }
});
