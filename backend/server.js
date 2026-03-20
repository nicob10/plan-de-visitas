const express = require("express");
const cors = require("cors");
const path = require("path");
const session = require("express-session");
const XLSX = require("xlsx");
const { query, hashPassword, initDb, withTransaction } = require("./db");
const {
  MEETING_STATUS,
  MEETING_TYPES,
  MEETING_REASONS,
  USER_ROLES,
  USER_ROLE_OPTIONS,
  OPPORTUNITY_STATUS,
  OPPORTUNITY_STATUS_OPTIONS,
  OPPORTUNITY_OPEN_STATUSES,
  OPPORTUNITY_SERVICE_LINES,
  OPPORTUNITY_TYPES,
  FOLLOW_UP_STATUS,
  FOLLOW_UP_STATUS_OPTIONS,
  FOLLOW_UP_TYPES
} = require("./constants");

const app = express();
const PORT = process.env.PORT || 3000;
const VALID_RISKS = new Set(["Bajo", "Medio", "Alto"]);
const VALID_SEGMENTS = new Set(["A", "B", "C"]);
const VALID_COMPANY_TYPES = new Set(["Local", "Global"]);
const VALID_COUNTRIES = new Set(["Argentina", "Bolivia", "Mexico"]);
const VALID_ACCOUNT_STAGES = new Set(["Activa", "Prospecto"]);
const VALID_MEETING_STATUS = new Set(Object.values(MEETING_STATUS));
const MEETING_MODALITIES = ["Virtual", "Presencial"];
const VALID_MEETING_MODALITIES = new Set(MEETING_MODALITIES);
const VALID_OPPORTUNITY_STATUS = new Set(OPPORTUNITY_STATUS_OPTIONS);
const VALID_SERVICE_LINES = new Set(OPPORTUNITY_SERVICE_LINES);
const VALID_OPPORTUNITY_TYPES = new Set(OPPORTUNITY_TYPES);
const VALID_FOLLOW_UP_STATUS = new Set(FOLLOW_UP_STATUS_OPTIONS);
const VALID_FOLLOW_UP_TYPES = new Set(FOLLOW_UP_TYPES);
const OPEN_OPPORTUNITY_STATUS_SET = new Set(OPPORTUNITY_OPEN_STATUSES);
const OPPORTUNITY_STATUS_SORT_ORDER = OPPORTUNITY_STATUS_OPTIONS.reduce((accumulator, status, index) => {
  accumulator[status] = index;
  return accumulator;
}, {});
const EXECUTIVE_ROLE = USER_ROLES.EXECUTIVE;
const SUPERVISOR_ROLE_BY_SERVICE = {
  fixedFire: USER_ROLES.SUPERVISOR_IFCI,
  extinguishers: USER_ROLES.SUPERVISOR_EXT,
  works: USER_ROLES.SUPERVISOR_WORKS
};
let meetingTypeOptionsCache = [...MEETING_TYPES];
let meetingReasonOptionsCache = [...MEETING_REASONS];

app.use(cors());
app.use(express.json());
app.use(
  session({
    secret: process.env.SESSION_SECRET || "top-clientes-dev-secret",
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax"
    }
  })
);
app.use((_req, res, next) => {
  res.setHeader("Cache-Control", "no-store");
  next();
});
app.use(express.static(path.join(__dirname, "..", "public")));

function asyncHandler(handler) {
  return (req, res, next) => {
    Promise.resolve(handler(req, res, next)).catch(next);
  };
}

function sanitizeUser(row) {
  if (!row) return null;
  return {
    id: row.id,
    name: row.name,
    email: row.email,
    role: row.role
  };
}

async function refreshMeetingTypesCache() {
  const { rows } = await query(
    `
    SELECT id, value, label, color, sort_order
    FROM meeting_type_options
    ORDER BY sort_order ASC, id ASC
    `
  );

  meetingTypeOptionsCache = rows.length
    ? rows.map((row) => ({
        id: row.id,
        value: row.value,
        label: row.label,
        color: row.color
      }))
    : [...MEETING_TYPES];
}

function getMeetingTypes() {
  return meetingTypeOptionsCache;
}

async function refreshMeetingReasonsCache() {
  const { rows } = await query(
    `
    SELECT id, name, sort_order
    FROM meeting_reason_options
    ORDER BY sort_order ASC, id ASC
    `
  );

  meetingReasonOptionsCache = rows.length
    ? rows.map((row) => ({ id: row.id, name: row.name }))
    : [...MEETING_REASONS];
}

function getMeetingReasons() {
  return meetingReasonOptionsCache;
}

function isValidMeetingReason(subject) {
  return getMeetingReasons().some((reason) => (typeof reason === "string" ? reason : reason.name) === subject);
}

function isValidMeetingType(kind) {
  return getMeetingTypes().some((type) => type.value === kind);
}

function validateUserPayload(payload, { requirePassword = true } = {}) {
  const errors = [];
  const name = String(payload?.name || "").trim();
  const email = String(payload?.email || "").trim().toLowerCase();
  const password = String(payload?.password || "");
  const role = String(payload?.role || "").trim();

  if (!name) errors.push("El nombre es obligatorio");
  if (!email || !email.includes("@")) errors.push("El email es inválido");
  if (requirePassword && password.length < 6) errors.push("La contraseña debe tener al menos 6 caracteres");
  if (!requirePassword && password && password.length < 6) {
    errors.push("La contraseña debe tener al menos 6 caracteres");
  }
  if (!USER_ROLE_OPTIONS.includes(role)) errors.push("Rol inválido");

  return {
    errors,
    values: { name, email, password, role }
  };
}

function requireAuth(req, res, next) {
  if (!req.session.userId) {
    res.status(401).json({ error: "Sesión requerida" });
    return;
  }
  next();
}

function isSettingsAdminUser(user) {
  return (
    String(user?.email || "").trim().toLowerCase() === "nicolas@maxiseguridad.com"
  );
}

async function requireSettingsAdmin(req, res, next) {
  const { rows } = await query("SELECT id, name, email, role FROM users WHERE id = $1", [req.session.userId]);
  const user = rows[0];

  if (!user) {
    req.session.destroy(() => {});
    res.status(401).json({ error: "Sesión inválida" });
    return;
  }

  if (!isSettingsAdminUser(user)) {
    res.status(403).json({ error: "No tenés permisos para acceder a esta sección" });
    return;
  }

  req.authUser = user;
  next();
}

function toClientObject(row) {
  return {
    id: row.id,
    rank: row.position,
    name: row.name,
    sector: row.sector,
    companyType: row.company_type || "Local",
    country: row.country || "Argentina",
    accountStage: row.account_stage || "Activa",
    manager: row.manager,
    executiveUserId: row.executive_user_id,
    executiveName: row.executive_name || row.manager,
    billing: Number(row.billing_2025),
    potential: Math.round(Number(row.billing_2025) * 1.12),
    segment: row.segment,
    risk: row.risk,
    services: {
      fixedFire: !!row.service_fixed_fire,
      extinguishers: !!row.service_extinguishers,
      works: !!row.service_works
    },
    supervisors: {
      fixedFire: row.supervisor_ifci_user_id
        ? {
            userId: row.supervisor_ifci_user_id,
            name: row.supervisor_ifci_name || "",
            role: USER_ROLES.SUPERVISOR_IFCI
          }
        : null,
      extinguishers: row.supervisor_ext_user_id
        ? {
            userId: row.supervisor_ext_user_id,
            name: row.supervisor_ext_name || "",
            role: USER_ROLES.SUPERVISOR_EXT
          }
        : null,
      works: row.supervisor_works_user_id
        ? {
            userId: row.supervisor_works_user_id,
            name: row.supervisor_works_name || "",
            role: USER_ROLES.SUPERVISOR_WORKS
          }
        : null
    },
    walletShare: row.wallet_share,
    nps: row.nps,
    openOpportunities: row.open_opportunities,
    notes: row.notes
  };
}

function toBranchObject(row) {
  return {
    id: row.id,
    clientId: row.client_id,
    name: row.name,
    sector: row.sector,
    manager: row.manager,
    executiveUserId: row.parent_executive_user_id,
    executiveName: row.parent_executive_name || row.parent_manager || row.manager,
    segment: row.segment,
    risk: row.risk,
    services: {
      fixedFire: !!row.service_fixed_fire,
      extinguishers: !!row.service_extinguishers,
      works: !!row.service_works
    },
    supervisors: {
      fixedFire: row.supervisor_ifci_user_id
        ? {
            userId: row.supervisor_ifci_user_id,
            name: row.supervisor_ifci_name || "",
            role: USER_ROLES.SUPERVISOR_IFCI
          }
        : null,
      extinguishers: row.supervisor_ext_user_id
        ? {
            userId: row.supervisor_ext_user_id,
            name: row.supervisor_ext_name || "",
            role: USER_ROLES.SUPERVISOR_EXT
          }
        : null,
      works: row.supervisor_works_user_id
        ? {
            userId: row.supervisor_works_user_id,
            name: row.supervisor_works_name || "",
            role: USER_ROLES.SUPERVISOR_WORKS
          }
        : null
    },
    notes: row.notes,
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

function buildClientSelectQuery(whereClause = "", orderClause = "ORDER BY clients.billing_2025 DESC, clients.id ASC") {
  return `
    SELECT
      clients.*,
      executive_user.name AS executive_name,
      supervisor_ifci.name AS supervisor_ifci_name,
      supervisor_ext.name AS supervisor_ext_name,
      supervisor_works.name AS supervisor_works_name
    FROM clients
    LEFT JOIN users AS executive_user ON executive_user.id = clients.executive_user_id
    LEFT JOIN users AS supervisor_ifci ON supervisor_ifci.id = clients.supervisor_ifci_user_id
    LEFT JOIN users AS supervisor_ext ON supervisor_ext.id = clients.supervisor_ext_user_id
    LEFT JOIN users AS supervisor_works ON supervisor_works.id = clients.supervisor_works_user_id
    ${whereClause}
    ${orderClause}
  `;
}

function buildBranchSelectQuery(whereClause = "", orderClause = "ORDER BY client_branches.id ASC") {
  return `
    SELECT
      client_branches.*,
      parent_client.manager AS parent_manager,
      parent_client.executive_user_id AS parent_executive_user_id,
      parent_executive_user.name AS parent_executive_name,
      supervisor_ifci.name AS supervisor_ifci_name,
      supervisor_ext.name AS supervisor_ext_name,
      supervisor_works.name AS supervisor_works_name
    FROM client_branches
    INNER JOIN clients AS parent_client ON parent_client.id = client_branches.client_id
    LEFT JOIN users AS parent_executive_user ON parent_executive_user.id = parent_client.executive_user_id
    LEFT JOIN users AS supervisor_ifci ON supervisor_ifci.id = client_branches.supervisor_ifci_user_id
    LEFT JOIN users AS supervisor_ext ON supervisor_ext.id = client_branches.supervisor_ext_user_id
    LEFT JOIN users AS supervisor_works ON supervisor_works.id = client_branches.supervisor_works_user_id
    ${whereClause}
    ${orderClause}
  `;
}

function parseBooleanCell(value) {
  const normalized = String(value ?? "").trim().toLowerCase();
  return ["si", "sí", "true", "1", "x"].includes(normalized);
}

async function findUserByEmailOrName(value, expectedRole = null) {
  const normalized = String(value || "").trim();
  if (!normalized) return null;
  const { rows } = await query(
    `
    SELECT id, name, email, role
    FROM users
    WHERE lower(email) = lower($1) OR lower(name) = lower($1)
    ORDER BY id ASC
    LIMIT 1
    `,
    [normalized]
  );
  const user = rows[0] || null;
  if (!user) return null;
  if (expectedRole && user.role !== expectedRole) return null;
  return user;
}

async function createClientRecord(payload) {
  const nextPositionResult = await query("SELECT COALESCE(MAX(position), 0) + 1 AS next FROM clients");
  const nextPosition = nextPositionResult.rows[0].next;
  const {
    name,
    billing,
    sector,
    companyType,
    country,
    accountStage,
    executiveUserId,
    risk,
    segment,
    services,
    supervisors,
    notes
  } = payload;

  const normalizedBilling = typeof billing === "number" && !Number.isNaN(billing) ? billing : 0;
  const executiveResult = await query("SELECT name FROM users WHERE id = $1", [Number(executiveUserId)]);
  const executiveName = executiveResult.rows[0]?.name || "";
  const fixedFireSupervisorId = services.fixedFire ? normalizeNullableUserId(supervisors.fixedFire.userId) : null;
  const extinguishersSupervisorId = services.extinguishers ? normalizeNullableUserId(supervisors.extinguishers.userId) : null;
  const worksSupervisorId = services.works ? normalizeNullableUserId(supervisors.works.userId) : null;

  const createdResult = await query(
    `
    INSERT INTO clients (
      position, name, billing_2025, sector, company_type, country, account_stage, manager, executive_user_id, risk, segment,
      service_fixed_fire, service_extinguishers, service_works,
      wallet_share, nps, open_opportunities, notes,
      supervisor_ifci_user_id, supervisor_ext_user_id, supervisor_works_user_id
    ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21)
    RETURNING *
    `,
    [
      nextPosition,
      name.trim(),
      normalizedBilling,
      sector.trim(),
      companyType,
      country,
      accountStage,
      executiveName,
      Number(executiveUserId),
      risk,
      segment,
      services.fixedFire,
      services.extinguishers,
      services.works,
      "",
      0,
      0,
      notes.trim(),
      fixedFireSupervisorId,
      extinguishersSupervisorId,
      worksSupervisorId
    ]
  );

  const selectedCreated = await query(buildClientSelectQuery("WHERE clients.id = $1", ""), [createdResult.rows[0].id]);
  return buildClientResponse(selectedCreated.rows[0]);
}

async function createUserRecord(payload) {
  const { values } = validateUserPayload(payload, { requirePassword: true });
  const result = await query(
    `
    INSERT INTO users (name, email, password_hash, role)
    VALUES ($1, $2, $3, $4)
    RETURNING id, name, email, role, created_at
    `,
    [values.name, values.email, hashPassword(values.password), values.role]
  );

  return {
    ...sanitizeUser(result.rows[0]),
    createdAt: result.rows[0].created_at
  };
}

async function getUsersByRoles(roles) {
  const { rows } = await query(
    `
    SELECT id, name, email, role
    FROM users
    WHERE role = ANY($1::text[])
    ORDER BY name ASC, id ASC
    `,
    [roles]
  );
  return rows.map(sanitizeUser);
}

async function getSectorOptions() {
  const { rows } = await query(
    `
    SELECT id, name, created_at
    FROM sector_options
    ORDER BY lower(name) ASC, id ASC
    `
  );

  return rows.map((row) => ({
    id: row.id,
    name: row.name,
    createdAt: row.created_at
  }));
}

async function getUserById(userId) {
  const { rows } = await query(
    `
    SELECT id, name, email, role, created_at
    FROM users
    WHERE id = $1
    `,
    [userId]
  );
  return rows[0] || null;
}

async function getUserAssignments(userId) {
  const { rows } = await query(
    `
    SELECT
      id,
      name,
      'company' AS entity_type,
      executive_user_id,
      supervisor_ifci_user_id,
      supervisor_ext_user_id,
      supervisor_works_user_id
    FROM clients
    WHERE executive_user_id = $1
       OR supervisor_ifci_user_id = $1
       OR supervisor_ext_user_id = $1
       OR supervisor_works_user_id = $1
    UNION ALL
    SELECT
      id,
      name,
      'branch' AS entity_type,
      executive_user_id,
      supervisor_ifci_user_id,
      supervisor_ext_user_id,
      supervisor_works_user_id
    FROM client_branches
    WHERE executive_user_id = $1
       OR supervisor_ifci_user_id = $1
       OR supervisor_ext_user_id = $1
       OR supervisor_works_user_id = $1
    ORDER BY name ASC, id ASC
    `,
    [userId]
  );

  return rows.map((row) => ({
    clientId: row.id,
    clientName: `${row.name}${row.entity_type === "branch" ? " (Sucursal)" : ""}`,
    asExecutive: Number(row.executive_user_id) === Number(userId),
    asSupervisorIfci: Number(row.supervisor_ifci_user_id) === Number(userId),
    asSupervisorExt: Number(row.supervisor_ext_user_id) === Number(userId),
    asSupervisorWorks: Number(row.supervisor_works_user_id) === Number(userId)
  }));
}

function validateRoleChangeAgainstAssignments(assignments, nextRole) {
  if (!assignments.length) return null;

  if (assignments.some((assignment) => assignment.asExecutive) && nextRole !== USER_ROLES.EXECUTIVE) {
    const company = assignments.find((assignment) => assignment.asExecutive);
    return `No podés cambiar el rol: el usuario está asignado como Ejecutivo en ${company.clientName}`;
  }
  if (assignments.some((assignment) => assignment.asSupervisorIfci) && nextRole !== USER_ROLES.SUPERVISOR_IFCI) {
    const company = assignments.find((assignment) => assignment.asSupervisorIfci);
    return `No podés cambiar el rol: el usuario está asignado como Supervisor IFCI en ${company.clientName}`;
  }
  if (assignments.some((assignment) => assignment.asSupervisorExt) && nextRole !== USER_ROLES.SUPERVISOR_EXT) {
    const company = assignments.find((assignment) => assignment.asSupervisorExt);
    return `No podés cambiar el rol: el usuario está asignado como Supervisor EXT en ${company.clientName}`;
  }
  if (assignments.some((assignment) => assignment.asSupervisorWorks) && nextRole !== USER_ROLES.SUPERVISOR_WORKS) {
    const company = assignments.find((assignment) => assignment.asSupervisorWorks);
    return `No podés cambiar el rol: el usuario está asignado como Supervisor Obra en ${company.clientName}`;
  }

  return null;
}

async function validateUserAssignment(userId, expectedRole, label) {
  if (userId === null || userId === undefined) return null;
  const numericUserId = Number(userId);
  if (!Number.isInteger(numericUserId) || numericUserId <= 0) {
    return `${label} inválido`;
  }

  const { rows } = await query("SELECT id, role FROM users WHERE id = $1", [numericUserId]);
  const user = rows[0];
  if (!user) return `${label} no encontrado`;
  if (user.role !== expectedRole) return `${label} debe tener el rol ${expectedRole}`;
  return null;
}

function normalizeNullableUserId(value) {
  if (value === null || value === undefined || value === "") return null;
  const numericValue = Number(value);
  if (!Number.isInteger(numericValue) || numericValue <= 0) return NaN;
  return numericValue;
}

async function validateSectorExists(sectorName) {
  const normalizedSector = String(sectorName || "").trim();
  if (!normalizedSector) return "El campo sector es obligatorio";

  const { rows } = await query("SELECT id FROM sector_options WHERE lower(name) = lower($1)", [normalizedSector]);
  if (!rows[0]) {
    return "El sector seleccionado no existe en la configuración";
  }
  return null;
}

async function fetchClientBranches(clientId) {
  const { rows } = await query(buildBranchSelectQuery("WHERE client_branches.client_id = $1"), [clientId]);
  return rows.map(toBranchObject);
}

async function withBillingRank(client) {
  const { rows } = await query(
    `
    SELECT COUNT(*)::int + 1 AS rank
    FROM clients
    WHERE billing_2025 > $1
       OR (billing_2025 = $2 AND id < $3)
    `,
    [client.billing, client.billing, client.id]
  );

  return {
    ...client,
    rank: rows[0]?.rank || 1
  };
}

function toMeetingObject(row) {
  return {
    id: row.id,
    clientId: row.client_id,
    clientName: row.client_name || "",
    branchId: row.branch_id,
    opportunityId: row.opportunity_id || null,
    kind: row.kind,
    subject: row.subject,
    objective: row.objective,
    scheduledFor: row.scheduled_for,
    participants: row.participants,
    participantUserIds: Array.isArray(row.participant_user_ids) ? row.participant_user_ids.map((id) => Number(id)) : [],
    contactName: row.contact_name,
    contactRole: row.contact_role,
    modality: row.modality || "Presencial",
    nextMeetingDate: row.next_meeting_date || "",
    followUpFromMeetingId: row.follow_up_from_meeting_id || null,
    minutes: row.minutes,
    findings: row.findings,
    activeNegotiationsStatus: row.active_negotiations_status || row.findings || "",
    opportunities: row.opportunities || "",
    substituteRecovery: row.substitute_recovery || "",
    globalContacts: row.global_contacts || "",
    serviceStatus: row.service_status || "",
    createdBy: row.created_by,
    status: row.status,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
    opportunityTitle: row.opportunity_title || ""
  };
}

function toOpportunityObject(row) {
  return {
    id: row.id,
    clientId: row.client_id,
    clientName: row.client_name || "",
    branchId: row.branch_id,
    branchName: row.branch_name || "",
    title: row.title,
    opportunityType: row.opportunity_type,
    serviceLine: row.service_line,
    status: row.status,
    amount: Number(row.amount || 0),
    probability: Number(row.probability || 0),
    expectedCloseDate: row.expected_close_date || "",
    ownerUserId: row.owner_user_id || null,
    ownerName: row.owner_name || "",
    source: row.source || "",
    description: row.description || "",
    lossReason: row.loss_reason || "",
    createdBy: row.created_by || "",
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

function toFollowUpObject(row) {
  return {
    id: row.id,
    opportunityId: row.opportunity_id,
    type: row.type,
    title: row.title,
    dueDate: row.due_date || "",
    status: row.status,
    assignedUserId: row.assigned_user_id || null,
    assignedUserName: row.assigned_user_name || "",
    notes: row.notes || "",
    completedAt: row.completed_at,
    createdBy: row.created_by || "",
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

function todayKey() {
  return new Date().toISOString().slice(0, 10);
}

function isOpenOpportunityStatus(status) {
  return OPEN_OPPORTUNITY_STATUS_SET.has(status);
}

function classifyFollowUpStatus(followUp) {
  if (!followUp) return FOLLOW_UP_STATUS.PENDING;
  if (followUp.status === FOLLOW_UP_STATUS.DONE) return FOLLOW_UP_STATUS.DONE;
  if (followUp.dueDate && followUp.dueDate < todayKey()) return "Vencido";
  return FOLLOW_UP_STATUS.PENDING;
}

function enrichOpportunity(opportunity, followUps = []) {
  const sortedFollowUps = [...followUps].sort((left, right) => {
    const leftPriority = classifyFollowUpStatus(left) === "Vencido" ? 0 : left.status === FOLLOW_UP_STATUS.PENDING ? 1 : 2;
    const rightPriority = classifyFollowUpStatus(right) === "Vencido" ? 0 : right.status === FOLLOW_UP_STATUS.PENDING ? 1 : 2;
    if (leftPriority !== rightPriority) return leftPriority - rightPriority;
    return String(left.dueDate || "").localeCompare(String(right.dueDate || ""));
  });

  const pendingFollowUps = sortedFollowUps.filter((followUp) => followUp.status === FOLLOW_UP_STATUS.PENDING);
  const overdueFollowUps = pendingFollowUps.filter((followUp) => followUp.dueDate && followUp.dueDate < todayKey());
  const nextFollowUp = sortedFollowUps.find((followUp) => followUp.status === FOLLOW_UP_STATUS.PENDING) || null;

  return {
    ...opportunity,
    followUps: sortedFollowUps.map((followUp) => ({
      ...followUp,
      visualStatus: classifyFollowUpStatus(followUp)
    })),
    pendingFollowUps: pendingFollowUps.length,
    overdueFollowUps: overdueFollowUps.length,
    nextFollowUp: nextFollowUp
      ? {
          ...nextFollowUp,
          visualStatus: classifyFollowUpStatus(nextFollowUp)
        }
      : null,
    weightedAmount: Math.round(opportunity.amount * (opportunity.probability / 100))
  };
}

function buildCrmSummary(opportunities) {
  const openOpportunities = opportunities.filter((opportunity) => isOpenOpportunityStatus(opportunity.status));
  const overdueFollowUps = openOpportunities.reduce((total, opportunity) => total + Number(opportunity.overdueFollowUps || 0), 0);
  const dueThisWeek = openOpportunities.reduce((total, opportunity) => {
    const nextFollowUp = opportunity.nextFollowUp;
    if (!nextFollowUp || nextFollowUp.status !== FOLLOW_UP_STATUS.PENDING || !nextFollowUp.dueDate) return total;
    const delta = Math.floor((new Date(`${nextFollowUp.dueDate}T00:00:00`) - new Date(`${todayKey()}T00:00:00`)) / 86400000);
    return delta >= 0 && delta <= 7 ? total + 1 : total;
  }, 0);

  return {
    openCount: openOpportunities.length,
    totalAmount: openOpportunities.reduce((total, opportunity) => total + opportunity.amount, 0),
    weightedAmount: openOpportunities.reduce((total, opportunity) => total + opportunity.weightedAmount, 0),
    overdueFollowUps,
    dueThisWeek
  };
}

async function attachFollowUpsToOpportunities(opportunities) {
  if (!opportunities.length) return [];

  const opportunityIds = opportunities.map((opportunity) => opportunity.id);
  const { rows } = await query(
    `
    SELECT
      opportunity_follow_ups.*,
      assigned_user.name AS assigned_user_name
    FROM opportunity_follow_ups
    LEFT JOIN users AS assigned_user ON assigned_user.id = opportunity_follow_ups.assigned_user_id
    WHERE opportunity_follow_ups.opportunity_id = ANY($1::INTEGER[])
    ORDER BY opportunity_follow_ups.due_date ASC, opportunity_follow_ups.id ASC
    `,
    [opportunityIds]
  );

  const followUpsByOpportunity = rows.reduce((accumulator, row) => {
    const followUp = toFollowUpObject(row);
    if (!accumulator.has(followUp.opportunityId)) {
      accumulator.set(followUp.opportunityId, []);
    }
    accumulator.get(followUp.opportunityId).push(followUp);
    return accumulator;
  }, new Map());

  return opportunities.map((opportunity) => enrichOpportunity(opportunity, followUpsByOpportunity.get(opportunity.id) || []));
}

async function fetchOpportunities({
  clientId = null,
  search = "",
  ownerUserId = null,
  status = null
} = {}) {
  const where = ["clients.is_hidden = FALSE"];
  const params = [];

  if (clientId !== null) {
    params.push(Number(clientId));
    where.push(`opportunities.client_id = $${params.length}`);
  }

  if (search.trim()) {
    params.push(`%${search.trim()}%`);
    where.push(
      `(opportunities.title ILIKE $${params.length} OR COALESCE(opportunities.source, '') ILIKE $${params.length} OR clients.name ILIKE $${params.length})`
    );
  }

  if (ownerUserId !== null) {
    params.push(Number(ownerUserId));
    where.push(`opportunities.owner_user_id = $${params.length}`);
  }

  if (status !== null) {
    params.push(status);
    where.push(`opportunities.status = $${params.length}`);
  }

  const { rows } = await query(
    `
    SELECT
      opportunities.*,
      clients.name AS client_name,
      client_branches.name AS branch_name,
      owner_user.name AS owner_name
    FROM opportunities
    INNER JOIN clients ON clients.id = opportunities.client_id
    LEFT JOIN client_branches ON client_branches.id = opportunities.branch_id
    LEFT JOIN users AS owner_user ON owner_user.id = opportunities.owner_user_id
    WHERE ${where.join(" AND ")}
    ORDER BY
      CASE opportunities.status
        ${OPPORTUNITY_STATUS_OPTIONS.map((currentStatus, index) => `WHEN '${currentStatus}' THEN ${index + 1}`).join(" ")}
        ELSE 99
      END ASC,
      opportunities.expected_close_date ASC,
      opportunities.updated_at DESC,
      opportunities.id DESC
    `,
    params
  );

  return attachFollowUpsToOpportunities(rows.map(toOpportunityObject));
}

function normalizeParticipantUserIds(values) {
  return [...new Set((values || []).map((value) => Number(value)).filter((value) => Number.isInteger(value) && value > 0))];
}

async function resolveParticipantNames(participantUserIds) {
  const normalizedIds = normalizeParticipantUserIds(participantUserIds);
  if (!normalizedIds.length) return [];
  const { rows } = await query(
    `
    SELECT id, name
    FROM users
    WHERE id = ANY($1::INTEGER[])
    ORDER BY name ASC, id ASC
    `,
    [normalizedIds]
  );
  return rows.map((row) => row.name);
}

function getMeetingColor(kind) {
  return getMeetingTypes().find((type) => type.value === kind)?.color || "yellow";
}

function getMeetingLabel(kind) {
  return getMeetingTypes().find((type) => type.value === kind)?.label || kind;
}

function buildVisitsWhereClause(filters = {}) {
  const {
    search = "",
    status = "todos",
    kind = "todos",
    modality = "todos",
    executiveUserId = "todos",
    supervisorUserId = "todos",
    participantUserId = "todos",
    dateFrom = "",
    dateTo = ""
  } = filters;

  const where = [];
  const params = [];

  if (search) {
    params.push(`%${String(search).trim()}%`);
    where.push(
      `(clients.name ILIKE $${params.length} OR COALESCE(client_branches.name, '') ILIKE $${params.length} OR COALESCE(meetings.contact_name, '') ILIKE $${params.length} OR COALESCE(meetings.participants, '') ILIKE $${params.length})`
    );
  }

  if (status !== "todos") {
    params.push(status);
    where.push(`meetings.status = $${params.length}`);
  }

  if (kind !== "todos") {
    params.push(kind);
    where.push(`meetings.kind = $${params.length}`);
  }

  if (modality !== "todos") {
    params.push(modality);
    where.push(`meetings.modality = $${params.length}`);
  }

  if (executiveUserId !== "todos") {
    params.push(Number(executiveUserId));
    where.push(`clients.executive_user_id = $${params.length}`);
  }

  if (supervisorUserId !== "todos") {
    params.push(Number(supervisorUserId));
    where.push(`(
      COALESCE(client_branches.supervisor_ifci_user_id, clients.supervisor_ifci_user_id) = $${params.length}
      OR COALESCE(client_branches.supervisor_ext_user_id, clients.supervisor_ext_user_id) = $${params.length}
      OR COALESCE(client_branches.supervisor_works_user_id, clients.supervisor_works_user_id) = $${params.length}
    )`);
  }

  if (participantUserId !== "todos") {
    params.push(Number(participantUserId));
    where.push(`meetings.participant_user_ids @> ARRAY[$${params.length}]::INTEGER[]`);
  }

  if (dateFrom) {
    params.push(dateFrom);
    where.push(`meetings.scheduled_for >= $${params.length}`);
  }

  if (dateTo) {
    params.push(dateTo);
    where.push(`meetings.scheduled_for <= $${params.length}`);
  }

  return { where, params };
}

async function fetchVisits(filters = {}) {
  const { where, params } = buildVisitsWhereClause(filters);
  const { rows } = await query(
    `
    SELECT
      meetings.id,
      meetings.client_id,
      meetings.branch_id,
      meetings.opportunity_id,
      meetings.kind,
      meetings.subject,
      meetings.objective,
      meetings.scheduled_for,
      meetings.participants,
      meetings.participant_user_ids,
      meetings.contact_name,
      meetings.contact_role,
      meetings.modality,
      meetings.next_meeting_date,
      meetings.follow_up_from_meeting_id,
      meetings.minutes,
      meetings.findings,
      meetings.created_by,
      meetings.status,
      meetings.created_at,
      meetings.updated_at,
      clients.name AS client_name,
      clients.executive_user_id AS executive_user_id,
      executive_user.name AS executive_name,
      client_branches.name AS branch_name,
      COALESCE(branch_supervisor_ifci.name, client_supervisor_ifci.name) AS supervisor_ifci_name,
      COALESCE(branch_supervisor_ext.name, client_supervisor_ext.name) AS supervisor_ext_name,
      COALESCE(branch_supervisor_works.name, client_supervisor_works.name) AS supervisor_works_name
    FROM meetings
    INNER JOIN clients ON clients.id = meetings.client_id
    LEFT JOIN users AS executive_user ON executive_user.id = clients.executive_user_id
    LEFT JOIN client_branches ON client_branches.id = meetings.branch_id
    LEFT JOIN users AS client_supervisor_ifci ON client_supervisor_ifci.id = clients.supervisor_ifci_user_id
    LEFT JOIN users AS client_supervisor_ext ON client_supervisor_ext.id = clients.supervisor_ext_user_id
    LEFT JOIN users AS client_supervisor_works ON client_supervisor_works.id = clients.supervisor_works_user_id
    LEFT JOIN users AS branch_supervisor_ifci ON branch_supervisor_ifci.id = client_branches.supervisor_ifci_user_id
    LEFT JOIN users AS branch_supervisor_ext ON branch_supervisor_ext.id = client_branches.supervisor_ext_user_id
    LEFT JOIN users AS branch_supervisor_works ON branch_supervisor_works.id = client_branches.supervisor_works_user_id
    ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
    ORDER BY meetings.scheduled_for DESC, meetings.id DESC
    `,
    params
  );

  return rows.map((row) => ({
    ...toMeetingObject(row),
    clientName: row.client_name,
    branchName: row.branch_name,
    scopeLabel: row.branch_name ? `${row.client_name} · ${row.branch_name}` : `${row.client_name} · Casa matriz`,
    kindLabel: getMeetingLabel(row.kind),
    color: getMeetingColor(row.kind),
    executiveName: row.executive_name || "",
    supervisors: [row.supervisor_ifci_name, row.supervisor_ext_name, row.supervisor_works_name].filter(Boolean)
  }));
}

function isUpcomingMeetingStatus(status) {
  return status === MEETING_STATUS.SCHEDULED || status === MEETING_STATUS.CONFIRMED;
}

function buildMeetingDots(meetings) {
  return meetings
    .filter((meeting) => meeting.status === MEETING_STATUS.COMPLETED)
    .slice(0, 8)
    .map((meeting) => ({
      kind: meeting.kind,
      color: getMeetingColor(meeting.kind)
    }));
}

function buildMeetingSummary(meetings) {
  const upcoming = meetings
    .filter((meeting) => meeting.scheduledFor && isUpcomingMeetingStatus(meeting.status))
    .sort((a, b) => a.scheduledFor.localeCompare(b.scheduledFor));

  const latestCompleted = meetings
    .filter((meeting) => meeting.scheduledFor && meeting.status === MEETING_STATUS.COMPLETED)
    .sort((a, b) => b.scheduledFor.localeCompare(a.scheduledFor));

  if (upcoming.length) {
    return `${upcoming[0].subject} · ${upcoming[0].scheduledFor}`;
  }

  if (latestCompleted.length) {
    return `Última: ${latestCompleted[0].scheduledFor}`;
  }

  return "Sin reunión agendada";
}

async function attachMeetingSummary(client) {
  const { rows } = await query(
    `
    SELECT meetings.id, meetings.client_id, meetings.branch_id, meetings.opportunity_id, meetings.kind, meetings.subject, meetings.objective, meetings.scheduled_for, meetings.participants, meetings.participant_user_ids, meetings.contact_name, meetings.contact_role,
           meetings.modality, meetings.next_meeting_date, meetings.follow_up_from_meeting_id, meetings.minutes, meetings.findings,
           meetings.active_negotiations_status, meetings.opportunities, meetings.substitute_recovery, meetings.global_contacts, meetings.service_status,
           meetings.created_by, meetings.status, meetings.created_at, meetings.updated_at,
           opportunities_table.title AS opportunity_title
    FROM meetings
    LEFT JOIN opportunities AS opportunities_table ON opportunities_table.id = meetings.opportunity_id
    WHERE meetings.client_id = $1
    ORDER BY meetings.scheduled_for DESC, meetings.id DESC
    `,
    [client.id]
  );

  const meetings = rows.map(toMeetingObject);

  return {
    ...client,
    meetings,
    nextMeeting: buildMeetingSummary(meetings),
    meetingDots: buildMeetingDots(meetings),
    meetingsCount: meetings.length,
    meetingsCompleted: meetings.filter((meeting) => meeting.status === MEETING_STATUS.COMPLETED).length,
    upcomingMeetings: meetings.filter((meeting) => isUpcomingMeetingStatus(meeting.status)).length
  };
}

async function buildClientResponse(row, { includeCrm = true } = {}) {
  const ranked = await withBillingRank(toClientObject(row));
  const clientWithMeetings = await attachMeetingSummary(ranked);
  const branches = await fetchClientBranches(clientWithMeetings.id);

  if (!includeCrm) {
    return {
      ...clientWithMeetings,
      branches
    };
  }

  const opportunities = await fetchOpportunities({ clientId: clientWithMeetings.id });
  const crmSummary = buildCrmSummary(opportunities);
  return {
    ...clientWithMeetings,
    branches,
    opportunities,
    crmSummary,
    openOpportunities: crmSummary.openCount
  };
}

function mergeWhereClause(baseCondition, whereClause = "") {
  if (!whereClause) return `WHERE ${baseCondition}`;
  if (/^where\s+/i.test(whereClause.trim())) {
    return `${whereClause} AND ${baseCondition}`;
  }
  return `WHERE ${baseCondition} AND (${whereClause})`;
}

async function validateAccountPayload(payload, { requireCommercialMetrics = false, requireBilling = true, requireExecutive = true } = {}) {
  const errors = [];
  const {
    name,
    billing,
    sector,
    companyType,
    country,
    accountStage,
    executiveUserId,
    risk,
    segment,
    services,
    supervisors,
    notes
  } = payload || {};

  if (typeof name !== "string" || !name.trim()) errors.push("El campo name es obligatorio");
  if (requireBilling && (typeof billing !== "number" || Number.isNaN(billing) || billing < 0)) {
    errors.push("El campo billing debe ser un número mayor o igual a 0");
  }
  const sectorError = await validateSectorExists(sector);
  if (sectorError) errors.push(sectorError);
  if (requireExecutive && !VALID_COMPANY_TYPES.has(companyType)) errors.push("El tipo de compañía es inválido");
  if (requireExecutive && !VALID_COUNTRIES.has(country)) errors.push("El país de la compañía es inválido");
  if (requireExecutive && !VALID_ACCOUNT_STAGES.has(accountStage)) errors.push("El estado de la compañía es inválido");
  if (!VALID_RISKS.has(risk)) errors.push("El campo risk es inválido");
  if (!VALID_SEGMENTS.has(segment)) errors.push("El campo segment es inválido");

  if (!services || typeof services !== "object") {
    errors.push("El campo services es obligatorio");
  } else {
    const { fixedFire, extinguishers, works } = services;
    if (typeof fixedFire !== "boolean" || typeof extinguishers !== "boolean" || typeof works !== "boolean") {
      errors.push("Los servicios deben ser booleanos");
    }
  }

  if (requireCommercialMetrics) {
    const walletShare = String(payload?.walletShare || "").trim();
    const nps = Number(payload?.nps);
    const openOpportunities = Number(payload?.openOpportunities);
    if (!walletShare) errors.push("El campo walletShare es obligatorio");
    if (!Number.isInteger(nps) || nps < 0 || nps > 100) errors.push("El campo nps debe ser un entero entre 0 y 100");
    if (!Number.isInteger(openOpportunities) || openOpportunities < 0) {
      errors.push("El campo openOpportunities debe ser un entero mayor o igual a 0");
    }
  }
  if (typeof notes !== "string") errors.push("El campo notes debe ser texto");

  if (requireExecutive) {
    const executiveError = await validateUserAssignment(executiveUserId, EXECUTIVE_ROLE, "El ejecutivo");
    if (executiveError) errors.push(executiveError);
  }

  if (!supervisors || typeof supervisors !== "object") {
    errors.push("El bloque de supervisores es obligatorio");
  } else {
    for (const [serviceKey, expectedRole] of Object.entries(SUPERVISOR_ROLE_BY_SERVICE)) {
      if (services?.[serviceKey]) {
        const normalizedSupervisorId = normalizeNullableUserId(supervisors[serviceKey]?.userId);
        if (normalizedSupervisorId === null) {
          errors.push(
            `Debes asignar un supervisor de ${serviceKey === "fixedFire" ? "IFCI" : serviceKey === "extinguishers" ? "EXT" : "Obra"}`
          );
          continue;
        }
        if (Number.isNaN(normalizedSupervisorId)) {
          errors.push(
            `El supervisor de ${serviceKey === "fixedFire" ? "IFCI" : serviceKey === "extinguishers" ? "EXT" : "Obra"} es inválido`
          );
          continue;
        }
        const error = await validateUserAssignment(
          normalizedSupervisorId,
          expectedRole,
          `El supervisor de ${serviceKey === "fixedFire" ? "IFCI" : serviceKey === "extinguishers" ? "EXT" : "Obra"}`
        );
        if (error) errors.push(error);
      }
    }
  }

  return errors;
}

async function validateClientPayload(payload) {
  return validateAccountPayload(payload, { requireCommercialMetrics: false, requireBilling: false, requireExecutive: true });
}

async function validateBranchPayload(payload) {
  return validateAccountPayload(payload, { requireCommercialMetrics: false, requireBilling: false, requireExecutive: false });
}

async function validateParticipantUsers(participantUserIds) {
  const normalizedIds = normalizeParticipantUserIds(participantUserIds);
  if (!normalizedIds.length) {
    return "Debes seleccionar al menos un participante";
  }

  const { rows } = await query(
    `
    SELECT id
    FROM users
    WHERE id = ANY($1::INTEGER[])
    `,
    [normalizedIds]
  );

  if (rows.length !== normalizedIds.length) {
    return "Uno o más participantes no existen en el sistema";
  }

  return null;
}

async function validateAnyUser(userId, label) {
  if (userId === null || userId === undefined || userId === "") return null;
  const numericUserId = Number(userId);
  if (!Number.isInteger(numericUserId) || numericUserId <= 0) {
    return `${label} inválido`;
  }

  const { rows } = await query("SELECT id FROM users WHERE id = $1", [numericUserId]);
  if (!rows[0]) return `${label} no encontrado`;
  return null;
}

async function validateOpportunityPayload(payload, clientId) {
  const errors = [];

  if (typeof payload?.title !== "string" || !payload.title.trim()) {
    errors.push("El nombre de la oportunidad es obligatorio");
  }

  if (!VALID_OPPORTUNITY_TYPES.has(payload?.opportunityType)) {
    errors.push("El tipo de oportunidad es inválido");
  }

  if (!VALID_SERVICE_LINES.has(payload?.serviceLine)) {
    errors.push("La línea de servicio es inválida");
  }

  if (!VALID_OPPORTUNITY_STATUS.has(payload?.status)) {
    errors.push("El estado de la oportunidad es inválido");
  }

  const amount = Number(payload?.amount);
  if (!Number.isFinite(amount) || amount < 0) {
    errors.push("El monto debe ser un número mayor o igual a 0");
  }

  const probability = Number(payload?.probability);
  if (!Number.isInteger(probability) || probability < 0 || probability > 100) {
    errors.push("La probabilidad debe ser un entero entre 0 y 100");
  }

  if (payload?.expectedCloseDate !== undefined && typeof payload.expectedCloseDate !== "string") {
    errors.push("La fecha estimada de cierre es inválida");
  }

  if (payload?.source !== undefined && typeof payload.source !== "string") {
    errors.push("El origen debe ser texto");
  }

  if (payload?.description !== undefined && typeof payload.description !== "string") {
    errors.push("La descripción debe ser texto");
  }

  if (payload?.lossReason !== undefined && typeof payload.lossReason !== "string") {
    errors.push("El motivo de pérdida debe ser texto");
  }

  if (payload?.status === OPPORTUNITY_STATUS.LOST && !String(payload?.lossReason || "").trim()) {
    errors.push("Debes indicar el motivo de pérdida");
  }

  const ownerError = await validateAnyUser(payload?.ownerUserId, "El responsable");
  if (ownerError) errors.push(ownerError);

  const branchId = payload?.branchId ? Number(payload.branchId) : null;
  if (branchId) {
    const branchResult = await query("SELECT id FROM client_branches WHERE id = $1 AND client_id = $2", [branchId, clientId]);
    if (!branchResult.rows[0]) {
      errors.push("La sucursal seleccionada no pertenece a esta compañía");
    }
  }

  return errors;
}

async function validateFollowUpPayload(payload) {
  const errors = [];

  if (typeof payload?.title !== "string" || !payload.title.trim()) {
    errors.push("El título del seguimiento es obligatorio");
  }

  if (typeof payload?.dueDate !== "string" || !payload.dueDate.trim()) {
    errors.push("La fecha del seguimiento es obligatoria");
  }

  if (!VALID_FOLLOW_UP_TYPES.has(payload?.type)) {
    errors.push("El tipo de seguimiento es inválido");
  }

  if (!VALID_FOLLOW_UP_STATUS.has(payload?.status)) {
    errors.push("El estado del seguimiento es inválido");
  }

  if (payload?.notes !== undefined && typeof payload.notes !== "string") {
    errors.push("Los comentarios del seguimiento deben ser texto");
  }

  const assignedError = await validateAnyUser(payload?.assignedUserId, "El usuario asignado");
  if (assignedError) errors.push(assignedError);

  return errors;
}

function normalizeMeetingStatus(payload) {
  const hasContent = [payload.minutes, payload.findings].some((value) => String(value || "").trim());
  if (payload.status && VALID_MEETING_STATUS.has(payload.status)) return payload.status;
  const today = new Date().toISOString().slice(0, 10);
  if (payload.scheduledFor && payload.scheduledFor > today) return MEETING_STATUS.SCHEDULED;
  return hasContent ? MEETING_STATUS.COMPLETED : MEETING_STATUS.SCHEDULED;
}

async function validateMeetingPayload(payload) {
  const errors = [];

  if (typeof payload?.subject !== "string" || !payload.subject.trim()) {
    errors.push("El motivo de la reunión es obligatorio");
  }
  if (typeof payload?.subject === "string" && payload.subject.trim() && !isValidMeetingReason(payload.subject.trim())) {
    errors.push("El motivo de la reunión es inválido");
  }
  if (typeof payload?.objective !== "string" || !payload.objective.trim()) {
    errors.push("El objetivo de la reunión es obligatorio");
  }
  if (typeof payload?.scheduledFor !== "string" || !payload.scheduledFor.trim()) {
    errors.push("La fecha de la reunión es obligatoria");
  }
  const participantError = await validateParticipantUsers(payload?.participantUserIds);
  if (participantError) errors.push(participantError);
  if (typeof payload?.contactName !== "string" || !payload.contactName.trim()) {
    errors.push("Debes indicar con quién nos vamos a reunir");
  }
  if (typeof payload?.contactRole !== "string" || !payload.contactRole.trim()) {
    errors.push("Debes indicar la función del contacto");
  }
  if (!VALID_MEETING_MODALITIES.has(payload?.modality)) {
    errors.push("La modalidad de la reunión es inválida");
  }
  if (!isValidMeetingType(payload?.kind)) {
    errors.push("El tipo de reunión es inválido");
  }
  if (payload?.createdBy !== undefined && typeof payload.createdBy !== "string") {
    errors.push("El campo cargado por debe ser texto");
  }
  if (payload?.minutes !== undefined && typeof payload.minutes !== "string") {
    errors.push("La minuta debe ser texto");
  }
  if (payload?.findings !== undefined && typeof payload.findings !== "string") {
    errors.push("Los hallazgos deben ser texto");
  }
  for (const field of ["activeNegotiationsStatus", "opportunities", "substituteRecovery", "globalContacts", "serviceStatus"]) {
    if (payload?.[field] !== undefined && typeof payload[field] !== "string") {
      errors.push(`El campo ${field} debe ser texto`);
    }
  }
  if (payload?.status && !VALID_MEETING_STATUS.has(payload.status)) {
    errors.push("El estado de la reunión es inválido");
  }
  if (payload?.nextMeetingDate !== undefined && typeof payload.nextMeetingDate !== "string") {
    errors.push("La fecha de próxima reunión es inválida");
  }
  if (typeof payload?.scheduledFor === "string" && typeof payload?.nextMeetingDate === "string") {
    const nextMeetingDate = payload.nextMeetingDate.trim();
    if (nextMeetingDate && nextMeetingDate < payload.scheduledFor.trim()) {
      errors.push("La próxima reunión no puede ser anterior a la reunión actual");
    }
  }

  return errors;
}

async function upsertMeetingFollowUp(client, meeting, createdBy) {
  const nextMeetingDate = String(meeting.nextMeetingDate || "").trim();
  const followUpResult = await client.query(
    "SELECT id FROM meetings WHERE follow_up_from_meeting_id = $1 ORDER BY id ASC LIMIT 1",
    [meeting.id]
  );
  const existingFollowUp = followUpResult.rows[0] || null;

  if (meeting.status !== MEETING_STATUS.COMPLETED || !nextMeetingDate) {
    if (existingFollowUp) {
      await client.query("DELETE FROM meetings WHERE id = $1", [existingFollowUp.id]);
    }
    return;
  }

  const followUpValues = [
    meeting.clientId,
    meeting.branchId,
    meeting.opportunityId,
    meeting.kind,
    meeting.subject,
    meeting.objective,
    nextMeetingDate,
    meeting.participants,
    meeting.participantUserIds,
    meeting.contactName,
    meeting.contactRole,
    meeting.modality,
    "",
    "",
    meeting.activeNegotiationsStatus || "",
    meeting.opportunities || "",
    meeting.substituteRecovery || "",
    meeting.globalContacts || "",
    meeting.serviceStatus || "",
    String(createdBy || "").trim(),
    MEETING_STATUS.SCHEDULED,
    meeting.id
  ];

  if (existingFollowUp) {
    await client.query(
      `
      UPDATE meetings
      SET
        client_id = $1,
        branch_id = $2,
        opportunity_id = $3,
        kind = $4,
        subject = $5,
        objective = $6,
        scheduled_for = $7,
        participants = $8,
        participant_user_ids = $9,
        contact_name = $10,
        contact_role = $11,
        modality = $12,
        minutes = $13,
        findings = $14,
        active_negotiations_status = $15,
        opportunities = $16,
        substitute_recovery = $17,
        global_contacts = $18,
        service_status = $19,
        created_by = $20,
        status = $21,
        updated_at = NOW()
      WHERE id = $22
      `,
      [...followUpValues.slice(0, 21), existingFollowUp.id]
    );
    return;
  }

  await client.query(
    `
    INSERT INTO meetings (
      client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids,
      contact_name, contact_role, modality, minutes, findings,
      active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_status,
      created_by, status, follow_up_from_meeting_id, created_at, updated_at
    ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21, $22, NOW(), NOW())
    `,
    followUpValues
  );
}

app.get("/api/health", (_req, res) => {
  res.json({ ok: true });
});

app.get(
  "/api/sector-options",
  asyncHandler(async (_req, res) => {
    res.json({ sectors: await getSectorOptions() });
  })
);

app.get(
  "/api/auth/me",
  asyncHandler(async (req, res) => {
    if (!req.session.userId) {
      res.status(401).json({ error: "Sesión requerida" });
      return;
    }

    const { rows } = await query("SELECT id, name, email, role FROM users WHERE id = $1", [req.session.userId]);
    const user = rows[0];

    if (!user) {
      req.session.destroy(() => {});
      res.status(401).json({ error: "Sesión inválida" });
      return;
    }

    res.json({
      user: {
        ...sanitizeUser(user),
        canManageSettings: isSettingsAdminUser(user)
      }
    });
  })
);

app.post(
  "/api/auth/login",
  asyncHandler(async (req, res) => {
    const email = String(req.body?.email || "").trim().toLowerCase();
    const password = String(req.body?.password || "");

    if (!email || !password) {
      res.status(400).json({ error: "Email y contraseña son obligatorios" });
      return;
    }

    const { rows } = await query("SELECT * FROM users WHERE lower(email) = $1", [email]);
    const user = rows[0];

    if (!user || user.password_hash !== hashPassword(password)) {
      res.status(401).json({ error: "Credenciales inválidas" });
      return;
    }

    req.session.userId = user.id;
    res.json({
      user: {
        ...sanitizeUser(user),
        canManageSettings: isSettingsAdminUser(user)
      }
    });
  })
);

app.post("/api/auth/logout", (req, res) => {
  req.session.destroy(() => {
    res.json({ ok: true });
  });
});

app.get("/api/meeting-types", (_req, res) => {
  res.json({
    meetingTypes: getMeetingTypes(),
    meetingReasons: getMeetingReasons(),
    statuses: Object.values(MEETING_STATUS),
    modalities: MEETING_MODALITIES
  });
});

app.use("/api", requireAuth);
app.use("/api/settings", requireSettingsAdmin);

app.get(
  "/api/crm/catalogs",
  asyncHandler(async (_req, res) => {
    res.json({
      opportunityStatuses: OPPORTUNITY_STATUS_OPTIONS,
      openOpportunityStatuses: OPPORTUNITY_OPEN_STATUSES,
      opportunityTypes: OPPORTUNITY_TYPES,
      serviceLines: OPPORTUNITY_SERVICE_LINES,
      followUpStatuses: FOLLOW_UP_STATUS_OPTIONS,
      followUpTypes: FOLLOW_UP_TYPES
    });
  })
);

app.get(
  "/api/settings/sectors",
  asyncHandler(async (_req, res) => {
    res.json({ sectors: await getSectorOptions() });
  })
);

app.get(
  "/api/settings/meeting-types",
  asyncHandler(async (_req, res) => {
    res.json({ meetingTypes: getMeetingTypes() });
  })
);

app.get(
  "/api/settings/meeting-reasons",
  asyncHandler(async (_req, res) => {
    res.json({ meetingReasons: getMeetingReasons() });
  })
);

app.get(
  "/api/settings/clients-import-template",
  asyncHandler(async (_req, res) => {
    const rows = [
      {
        name: "Cliente Ejemplo",
        sector: "Energía",
        companyType: "Local",
        country: "Argentina",
        accountStage: "Activa",
        executive: "comercial@maxi.local",
        risk: "Medio",
        segment: "B",
        fixedFire: "No",
        extinguishers: "Si",
        works: "No",
        supervisorIfci: "",
        supervisorExt: "ext@maxi.local",
        supervisorWorks: "",
        notes: "Importado desde Excel"
      }
    ];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Clientes");
    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="plantilla-clientes.xlsx"');
    res.send(buffer);
  })
);

app.get(
  "/api/settings/users-import-template",
  asyncHandler(async (_req, res) => {
    const rows = [
      {
        name: "Nuevo Comercial",
        email: "nuevo.comercial@empresa.com",
        role: USER_ROLES.EXECUTIVE,
        password: "Temporal1234"
      }
    ];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Usuarios");
    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="plantilla-usuarios.xlsx"');
    res.send(buffer);
  })
);

app.post(
  "/api/settings/clients-import",
  asyncHandler(async (req, res) => {
    const fileData = String(req.body?.fileData || "").trim();
    if (!fileData) {
      res.status(400).json({ error: "Debes adjuntar un archivo Excel" });
      return;
    }

    const workbook = XLSX.read(Buffer.from(fileData, "base64"), { type: "buffer" });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    if (!rows.length) {
      res.status(400).json({ error: "El archivo no tiene filas para importar" });
      return;
    }

    let createdCount = 0;
    const errors = [];

    for (const [index, row] of rows.entries()) {
      const line = index + 2;
      const executive = await findUserByEmailOrName(row.executive, EXECUTIVE_ROLE);
      if (!executive) {
        errors.push(`Fila ${line}: no se encontró un Ejecutivo válido en "executive"`);
        continue;
      }

      const services = {
        fixedFire: parseBooleanCell(row.fixedFire),
        extinguishers: parseBooleanCell(row.extinguishers),
        works: parseBooleanCell(row.works)
      };

      const supervisors = {
        fixedFire: { userId: null },
        extinguishers: { userId: null },
        works: { userId: null }
      };

      if (services.fixedFire) {
        const supervisor = await findUserByEmailOrName(row.supervisorIfci, USER_ROLES.SUPERVISOR_IFCI);
        supervisors.fixedFire.userId = supervisor?.id || null;
      }
      if (services.extinguishers) {
        const supervisor = await findUserByEmailOrName(row.supervisorExt, USER_ROLES.SUPERVISOR_EXT);
        supervisors.extinguishers.userId = supervisor?.id || null;
      }
      if (services.works) {
        const supervisor = await findUserByEmailOrName(row.supervisorWorks, USER_ROLES.SUPERVISOR_WORKS);
        supervisors.works.userId = supervisor?.id || null;
      }

      const payload = {
        name: String(row.name || "").trim(),
        billing: 0,
        sector: String(row.sector || "").trim(),
        companyType: String(row.companyType || "Local").trim() || "Local",
        country: String(row.country || "Argentina").trim() || "Argentina",
        accountStage: String(row.accountStage || "Activa").trim() || "Activa",
        executiveUserId: executive.id,
        risk: String(row.risk || "Bajo").trim() || "Bajo",
        segment: String(row.segment || "C").trim() || "C",
        services,
        supervisors,
        notes: String(row.notes || "").trim()
      };

      const duplicateResult = await query("SELECT id FROM clients WHERE lower(name) = lower($1) LIMIT 1", [payload.name]);
      if (duplicateResult.rows[0]) {
        errors.push(`Fila ${line}: ya existe un cliente con el nombre "${payload.name}"`);
        continue;
      }

      const validationErrors = await validateClientPayload(payload);
      if (validationErrors.length) {
        errors.push(`Fila ${line}: ${validationErrors[0]}`);
        continue;
      }

      await createClientRecord(payload);
      createdCount += 1;
    }

    res.json({
      createdCount,
      errorCount: errors.length,
      errors
    });
  })
);

app.post(
  "/api/settings/users-import",
  asyncHandler(async (req, res) => {
    const fileData = String(req.body?.fileData || "").trim();
    if (!fileData) {
      res.status(400).json({ error: "Debes adjuntar un archivo Excel" });
      return;
    }

    const workbook = XLSX.read(Buffer.from(fileData, "base64"), { type: "buffer" });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    if (!rows.length) {
      res.status(400).json({ error: "El archivo no tiene filas para importar" });
      return;
    }

    let createdCount = 0;
    const errors = [];

    for (const [index, row] of rows.entries()) {
      const line = index + 2;
      const payload = {
        name: String(row.name || "").trim(),
        email: String(row.email || "").trim(),
        role: String(row.role || "").trim(),
        password: String(row.password || "").trim()
      };

      const validation = validateUserPayload(payload, { requirePassword: true });
      if (validation.errors.length) {
        errors.push(`Fila ${line}: ${validation.errors[0]}`);
        continue;
      }

      const existing = await query("SELECT id FROM users WHERE lower(email) = lower($1) LIMIT 1", [payload.email]);
      if (existing.rows[0]) {
        errors.push(`Fila ${line}: ya existe un usuario con el email "${payload.email}"`);
        continue;
      }

      try {
        await createUserRecord(payload);
        createdCount += 1;
      } catch (error) {
        if (error.code === "23505") {
          errors.push(`Fila ${line}: ya existe un usuario con el email "${payload.email}"`);
          continue;
        }
        throw error;
      }
    }

    res.json({
      createdCount,
      errorCount: errors.length,
      errors
    });
  })
);

app.post(
  "/api/settings/meeting-types",
  asyncHandler(async (req, res) => {
    const value = String(req.body?.value || "").trim();
    const label = String(req.body?.label || "").trim();
    const color = String(req.body?.color || "yellow").trim();

    if (!value || !label) {
      res.status(400).json({ error: "El nombre interno y la etiqueta son obligatorios" });
      return;
    }

    await query(
      `
      INSERT INTO meeting_type_options (value, label, color, sort_order)
      VALUES ($1, $2, $3, COALESCE((SELECT MAX(sort_order) FROM meeting_type_options), 0) + 1)
      `,
      [value, label, color]
    );

    await refreshMeetingTypesCache();
    res.status(201).json({ meetingTypes: getMeetingTypes() });
  })
);

app.patch(
  "/api/settings/meeting-types/:id",
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    const value = String(req.body?.value || "").trim();
    const label = String(req.body?.label || "").trim();
    const color = String(req.body?.color || "yellow").trim();

    if (!Number.isInteger(id) || id <= 0) {
      res.status(400).json({ error: "Tipo de reunión inválido" });
      return;
    }

    if (!value || !label) {
      res.status(400).json({ error: "El nombre interno y la etiqueta son obligatorios" });
      return;
    }

    const result = await query(
      `
      UPDATE meeting_type_options
      SET value = $1, label = $2, color = $3
      WHERE id = $4
      RETURNING id
      `,
      [value, label, color, id]
    );

    if (!result.rows[0]) {
      res.status(404).json({ error: "Tipo de reunión no encontrado" });
      return;
    }

    await refreshMeetingTypesCache();
    res.json({ meetingTypes: getMeetingTypes() });
  })
);

app.delete(
  "/api/settings/meeting-types/:id",
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) {
      res.status(400).json({ error: "Tipo de reunión inválido" });
      return;
    }

    const existing = getMeetingTypes().find((type) => Number(type.id) === id);
    if (!existing) {
      res.status(404).json({ error: "Tipo de reunión no encontrado" });
      return;
    }

    const usage = await query("SELECT id FROM meetings WHERE kind = $1 LIMIT 1", [existing.value]);
    if (usage.rows[0]) {
      res.status(409).json({ error: "No se puede eliminar un tipo de reunión que ya está en uso" });
      return;
    }

    await query("DELETE FROM meeting_type_options WHERE id = $1", [id]);
    await refreshMeetingTypesCache();
    res.json({ meetingTypes: getMeetingTypes() });
  })
);

app.post(
  "/api/settings/meeting-reasons",
  asyncHandler(async (req, res) => {
    const name = String(req.body?.name || "").trim();
    if (!name) {
      res.status(400).json({ error: "El motivo de reunión es obligatorio" });
      return;
    }

    await query(
      `
      INSERT INTO meeting_reason_options (name, sort_order)
      VALUES ($1, COALESCE((SELECT MAX(sort_order) FROM meeting_reason_options), 0) + 1)
      `,
      [name]
    );

    await refreshMeetingReasonsCache();
    res.status(201).json({ meetingReasons: getMeetingReasons() });
  })
);

app.patch(
  "/api/settings/meeting-reasons/:id",
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    const name = String(req.body?.name || "").trim();
    if (!Number.isInteger(id) || id <= 0) {
      res.status(400).json({ error: "Motivo de reunión inválido" });
      return;
    }
    if (!name) {
      res.status(400).json({ error: "El motivo de reunión es obligatorio" });
      return;
    }

    const result = await query(
      `
      UPDATE meeting_reason_options
      SET name = $1
      WHERE id = $2
      RETURNING id
      `,
      [name, id]
    );

    if (!result.rows[0]) {
      res.status(404).json({ error: "Motivo de reunión no encontrado" });
      return;
    }

    await refreshMeetingReasonsCache();
    res.json({ meetingReasons: getMeetingReasons() });
  })
);

app.delete(
  "/api/settings/meeting-reasons/:id",
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) {
      res.status(400).json({ error: "Motivo de reunión inválido" });
      return;
    }

    const existing = getMeetingReasons().find((reason) => Number(reason.id) === id);
    if (!existing) {
      res.status(404).json({ error: "Motivo de reunión no encontrado" });
      return;
    }

    const usage = await query("SELECT id FROM meetings WHERE subject = $1 LIMIT 1", [existing.name]);
    if (usage.rows[0]) {
      res.status(409).json({ error: "No se puede eliminar un motivo de reunión que ya está en uso" });
      return;
    }

    await query("DELETE FROM meeting_reason_options WHERE id = $1", [id]);
    await refreshMeetingReasonsCache();
    res.json({ meetingReasons: getMeetingReasons() });
  })
);

app.post(
  "/api/settings/sectors",
  asyncHandler(async (req, res) => {
    const name = String(req.body?.name || "").trim();
    if (!name) {
      res.status(400).json({ error: "El nombre del sector es obligatorio" });
      return;
    }

    try {
      const result = await query(
        `
        INSERT INTO sector_options (name)
        VALUES ($1)
        RETURNING id, name, created_at
        `,
        [name]
      );

      res.status(201).json({
        sector: {
          id: result.rows[0].id,
          name: result.rows[0].name,
          createdAt: result.rows[0].created_at
        }
      });
    } catch (error) {
      if (error.code === "23505") {
        res.status(409).json({ error: "Ese sector ya existe" });
        return;
      }
      throw error;
    }
  })
);

app.delete(
  "/api/settings/sectors/:id",
  asyncHandler(async (req, res) => {
    const sectorId = Number(req.params.id);
    if (!Number.isInteger(sectorId) || sectorId <= 0) {
      res.status(400).json({ error: "Sector inválido" });
      return;
    }

    const sectorResult = await query("SELECT id, name FROM sector_options WHERE id = $1", [sectorId]);
    const sector = sectorResult.rows[0];
    if (!sector) {
      res.status(404).json({ error: "Sector no encontrado" });
      return;
    }

    const usageResult = await query("SELECT id FROM clients WHERE lower(sector) = lower($1) LIMIT 1", [sector.name]);
    if (usageResult.rows[0]) {
      res.status(409).json({ error: "No se puede eliminar un sector que ya está asignado a compañías" });
      return;
    }

    const branchUsageResult = await query("SELECT id FROM client_branches WHERE lower(sector) = lower($1) LIMIT 1", [sector.name]);
    if (branchUsageResult.rows[0]) {
      res.status(409).json({ error: "No se puede eliminar un sector que ya está asignado a sucursales" });
      return;
    }

    await query("DELETE FROM sector_options WHERE id = $1", [sectorId]);
    res.json({ ok: true });
  })
);

app.get(
  "/api/users",
  requireSettingsAdmin,
  asyncHandler(async (_req, res) => {
    const { rows } = await query(
      `
      SELECT id, name, email, role, created_at
      FROM users
      ORDER BY name ASC, id ASC
      `
    );

    res.json({
      users: rows.map((row) => ({
        ...sanitizeUser(row),
        createdAt: row.created_at
      })),
      roles: USER_ROLE_OPTIONS
    });
  })
);

app.post(
  "/api/users",
  requireSettingsAdmin,
  asyncHandler(async (req, res) => {
    const { errors, values } = validateUserPayload(req.body, { requirePassword: true });
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    try {
      res.status(201).json({
        user: await createUserRecord(values)
      });
    } catch (error) {
      if (error.code === "23505") {
        res.status(409).json({ error: "Ya existe un usuario con ese email" });
        return;
      }
      throw error;
    }
  })
);

app.patch(
  "/api/users/:id",
  requireSettingsAdmin,
  asyncHandler(async (req, res) => {
    const userId = Number(req.params.id);
    if (!Number.isInteger(userId) || userId <= 0) {
      res.status(400).json({ error: "Usuario inválido" });
      return;
    }

    const existingUser = await getUserById(userId);
    if (!existingUser) {
      res.status(404).json({ error: "Usuario no encontrado" });
      return;
    }

    const { errors, values } = validateUserPayload(req.body, { requirePassword: false });
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const assignments = await getUserAssignments(userId);
    const assignmentError = validateRoleChangeAgainstAssignments(assignments, values.role);
    if (assignmentError) {
      res.status(409).json({ error: assignmentError });
      return;
    }

    const updateFields = ["name = $1", "email = $2", "role = $3"];
    const params = [values.name, values.email, values.role];

    if (values.password) {
      updateFields.push(`password_hash = $${params.length + 1}`);
      params.push(hashPassword(values.password));
    }

    params.push(userId);

    let result;
    try {
      result = await query(
        `
        UPDATE users
        SET ${updateFields.join(", ")}
        WHERE id = $${params.length}
        RETURNING id, name, email, role, created_at
        `,
        params
      );
    } catch (error) {
      if (error.code === "23505") {
        res.status(409).json({ error: "Ya existe un usuario con ese email" });
        return;
      }
      throw error;
    }

    res.json({
      user: {
        ...sanitizeUser(result.rows[0]),
        createdAt: result.rows[0].created_at
      }
    });
  })
);

app.delete(
  "/api/users/:id",
  requireSettingsAdmin,
  asyncHandler(async (req, res) => {
    const userId = Number(req.params.id);
    if (!Number.isInteger(userId) || userId <= 0) {
      res.status(400).json({ error: "Usuario inválido" });
      return;
    }

    if (Number(req.session.userId) === userId) {
      res.status(400).json({ error: "No podés eliminar tu propio usuario mientras tenés la sesión abierta" });
      return;
    }

    const existingUser = await getUserById(userId);
    if (!existingUser) {
      res.status(404).json({ error: "Usuario no encontrado" });
      return;
    }

    const assignments = await getUserAssignments(userId);
    if (assignments.length) {
      res.status(409).json({
        error: `No podés eliminar este usuario porque está asignado a ${assignments[0].clientName}`
      });
      return;
    }

    await query("DELETE FROM users WHERE id = $1", [userId]);
    res.json({ ok: true });
  })
);

app.get(
  "/api/company-assignment-options",
  asyncHandler(async (_req, res) => {
    res.json({
      allUsers: await getUsersByRoles(USER_ROLE_OPTIONS),
      executives: await getUsersByRoles([EXECUTIVE_ROLE]),
      supervisors: {
        fixedFire: await getUsersByRoles([USER_ROLES.SUPERVISOR_IFCI]),
        extinguishers: await getUsersByRoles([USER_ROLES.SUPERVISOR_EXT]),
        works: await getUsersByRoles([USER_ROLES.SUPERVISOR_WORKS])
      }
    });
  })
);

app.get(
  "/api/pipeline",
  asyncHandler(async (req, res) => {
    const search = String(req.query.search || "").trim();
    const ownerUserId = String(req.query.ownerUserId || "todos").trim();
    const status = String(req.query.status || "todos").trim();

    const opportunities = await fetchOpportunities({
      search,
      ownerUserId: ownerUserId !== "todos" ? Number(ownerUserId) : null,
      status: status !== "todos" ? status : null
    });

    res.json({
      opportunities,
      summary: buildCrmSummary(opportunities)
    });
  })
);

app.get(
  "/api/clients",
  asyncHandler(async (req, res) => {
    const {
      search = "",
      risk = "todos",
      segment = "todos",
      fixedFire = "todos",
      extinguishers = "todos",
      works = "todos"
    } = req.query;

    const where = [];
    const params = [];

    if (search.trim()) {
      const like = `%${search.trim()}%`;
      params.push(like, like, like);
      where.push(
        `(clients.name ILIKE $${params.length - 2} OR clients.sector ILIKE $${params.length - 1} OR COALESCE(executive_user.name, clients.manager) ILIKE $${params.length})`
      );
    }
    if (risk !== "todos") {
      params.push(risk);
      where.push(`clients.risk = $${params.length}`);
    }
    if (segment !== "todos") {
      params.push(segment);
      where.push(`clients.segment = $${params.length}`);
    }
    if (fixedFire !== "todos") {
      params.push(fixedFire === "si");
      where.push(`clients.service_fixed_fire = $${params.length}`);
    }
    if (extinguishers !== "todos") {
      params.push(extinguishers === "si");
      where.push(`clients.service_extinguishers = $${params.length}`);
    }
    if (works !== "todos") {
      params.push(works === "si");
      where.push(`clients.service_works = $${params.length}`);
    }

    const { rows } = await query(
      buildClientSelectQuery(mergeWhereClause("clients.is_hidden = FALSE", where.length ? `WHERE ${where.join(" AND ")}` : "")),
      params
    );

    const clients = await Promise.all(rows.map((row) => buildClientResponse(row, { includeCrm: false })));

    const totalsResult = await query(
      "SELECT COUNT(*)::int AS total_clients, COALESCE(SUM(billing_2025), 0) AS total_billing FROM clients WHERE is_hidden = FALSE"
    );
    const totals = totalsResult.rows[0];

    const meetingSummaryResult = await query(`
      SELECT
        COUNT(*)::int AS total_meetings,
        COALESCE(SUM(CASE WHEN status = 'Agendada' THEN 1 ELSE 0 END), 0)::int AS scheduled_meetings,
        COALESCE(SUM(CASE WHEN status = 'Confirmada' THEN 1 ELSE 0 END), 0)::int AS confirmed_meetings,
        COALESCE(SUM(CASE WHEN status = 'Realizada' THEN 1 ELSE 0 END), 0)::int AS completed_meetings
        ,COUNT(DISTINCT CASE WHEN status = 'Realizada' THEN client_id END)::int AS visited_clients
      FROM meetings
    `);
    const meetingSummary = meetingSummaryResult.rows[0];
    const crmSummary = buildCrmSummary(await fetchOpportunities());

    res.json({
      clients,
      globalKpis: {
        totalClients: totals.total_clients || 0,
        totalBilling: Number(totals.total_billing) || 0,
        totalPotential: Math.round((Number(totals.total_billing) || 0) * 1.12),
        meetingsScheduled: meetingSummary.scheduled_meetings || 0,
        meetingsConfirmed: meetingSummary.confirmed_meetings || 0,
        meetingsCompleted: meetingSummary.completed_meetings || 0,
        meetingsTotal: meetingSummary.total_meetings || 0,
        clientsVisited: meetingSummary.visited_clients || 0,
        openOpportunities: crmSummary.openCount || 0,
        pipelineAmount: crmSummary.totalAmount || 0,
        weightedPipelineAmount: crmSummary.weightedAmount || 0,
        overdueFollowUps: crmSummary.overdueFollowUps || 0,
        visitsPerVisitedClient: meetingSummary.visited_clients
          ? Number(meetingSummary.completed_meetings || 0) / Number(meetingSummary.visited_clients)
          : 0
      }
    });
  })
);

app.get(
  "/api/calendar",
  asyncHandler(async (req, res) => {
    const month = String(req.query.month || "").trim();
    const participantUserId = String(req.query.participantUserId || "todos").trim();
    const where = ["meetings.status = ANY($1::text[])"];
    const params = [[MEETING_STATUS.SCHEDULED, MEETING_STATUS.CONFIRMED, MEETING_STATUS.COMPLETED]];

    if (/^\d{4}-\d{2}$/.test(month)) {
      params.push(month);
      where.push(`LEFT(meetings.scheduled_for, 7) = $${params.length}`);
    }

    if (participantUserId !== "todos") {
      params.push(Number(participantUserId));
      where.push(`meetings.participant_user_ids @> ARRAY[$${params.length}]::INTEGER[]`);
    }

    const { rows } = await query(
      `
      SELECT
        meetings.id,
        meetings.client_id,
        meetings.branch_id,
        meetings.opportunity_id,
        meetings.kind,
        meetings.subject,
        meetings.objective,
        meetings.scheduled_for,
        meetings.participants,
        meetings.participant_user_ids,
        meetings.contact_name,
        meetings.contact_role,
        meetings.modality,
        meetings.next_meeting_date,
        meetings.follow_up_from_meeting_id,
        meetings.minutes,
        meetings.findings,
        meetings.created_by,
        meetings.status,
        meetings.created_at,
        meetings.updated_at,
        clients.name AS client_name,
        client_branches.name AS branch_name
      FROM meetings
      INNER JOIN clients ON clients.id = meetings.client_id
      LEFT JOIN client_branches ON client_branches.id = meetings.branch_id
      WHERE ${where.join(" AND ")}
      ORDER BY meetings.scheduled_for ASC, meetings.id ASC
      `,
      params
    );

    const meetings = rows.map((row) => ({
      ...toMeetingObject(row),
      clientId: row.client_id,
      branchId: row.branch_id,
      clientName: row.client_name,
      branchName: row.branch_name,
      scopeLabel: row.branch_name ? `${row.client_name} · ${row.branch_name}` : row.client_name,
      color: getMeetingColor(row.kind)
    }));

    res.json({ meetings });
  })
);

app.get(
  "/api/visits",
  asyncHandler(async (req, res) => {
    res.json({ visits: await fetchVisits(req.query) });
  })
);

app.get(
  "/api/visits/export",
  asyncHandler(async (req, res) => {
    const visits = await fetchVisits(req.query);
    const rows = visits.map((visit) => ({
      Fecha: visit.scheduledFor,
      Cliente: visit.clientName,
      Alcance: visit.branchName || "Casa matriz",
      Tipo: visit.kindLabel || visit.kind,
      Modalidad: visit.modality || "",
      Estado: visit.status,
      Contacto: visit.contactName || "",
      Funcion: visit.contactRole || "",
      Participantes: visit.participants || "",
      Ejecutivo: visit.executiveName || "",
      Supervisores: visit.supervisors.join(" / ")
    }));

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Visitas");
    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="visitas.xlsx"');
    res.send(buffer);
  })
);

app.patch(
  "/api/meetings/:meetingId/status",
  asyncHandler(async (req, res) => {
    const meetingId = Number(req.params.meetingId);
    const status = String(req.body?.status || "").trim();
    const nextMeetingDate = String(req.body?.nextMeetingDate || "").trim();

    if (!Number.isInteger(meetingId) || meetingId <= 0) {
      res.status(400).json({ error: "Reunión inválida" });
      return;
    }

    if (!VALID_MEETING_STATUS.has(status)) {
      res.status(400).json({ error: "Estado de visita inválido" });
      return;
    }

    const existingResult = await query(
      `
      SELECT id, client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids, contact_name, contact_role,
             modality, next_meeting_date, follow_up_from_meeting_id, minutes, findings,
             active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_status,
             created_by, status, created_at, updated_at
      FROM meetings
      WHERE id = $1
      `,
      [meetingId]
    );
    const existingMeeting = existingResult.rows[0];

    if (!existingMeeting) {
      res.status(404).json({ error: "Reunión no encontrada" });
      return;
    }

    const updatedMeeting = await withTransaction(async (client) => {
      const result = await client.query(
        `
        UPDATE meetings
        SET
          status = $1,
          next_meeting_date = $2,
          updated_at = NOW()
        WHERE id = $3
        RETURNING id, client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids, contact_name, contact_role,
                  modality, next_meeting_date, follow_up_from_meeting_id, minutes, findings,
                  active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_status,
                  created_by, status, created_at, updated_at
        `,
        [status, nextMeetingDate, meetingId]
      );

      const meeting = toMeetingObject(result.rows[0]);
      await upsertMeetingFollowUp(client, meeting, meeting.createdBy || "");
      return meeting;
    });

    res.json({ meeting: updatedMeeting });
  })
);

app.get(
  "/api/clients/:id",
  asyncHandler(async (req, res) => {
    const { rows } = await query(buildClientSelectQuery("WHERE clients.id = $1 AND clients.is_hidden = FALSE", ""), [Number(req.params.id)]);
    const row = rows[0];

    if (!row) {
      res.status(404).json({ error: "Cliente no encontrado" });
      return;
    }

    res.json({ client: await buildClientResponse(row) });
  })
);

app.post(
  "/api/clients",
  asyncHandler(async (req, res) => {
    const payload = req.body || {};
    const errors = await validateClientPayload(payload);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    res.status(201).json({ client: await createClientRecord(payload) });
  })
);

app.patch(
  "/api/clients/:id",
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const existingClientResult = await query("SELECT id, billing_2025 FROM clients WHERE id = $1", [clientId]);

    if (!existingClientResult.rows[0]) {
      res.status(404).json({ error: "Cliente no encontrado" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateClientPayload(payload);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const {
      name,
      billing,
      sector,
      companyType,
      country,
      accountStage,
      executiveUserId,
      risk,
      segment,
      services,
      supervisors,
      notes
    } = payload;

    const normalizedBilling =
      typeof billing === "number" && !Number.isNaN(billing) ? billing : Number(existingClientResult.rows[0].billing_2025 || 0);
    const executiveResult = await query("SELECT name FROM users WHERE id = $1", [Number(executiveUserId)]);
    const executiveName = executiveResult.rows[0]?.name || "";
    const fixedFireSupervisorId = services.fixedFire ? normalizeNullableUserId(supervisors.fixedFire.userId) : null;
    const extinguishersSupervisorId = services.extinguishers ? normalizeNullableUserId(supervisors.extinguishers.userId) : null;
    const worksSupervisorId = services.works ? normalizeNullableUserId(supervisors.works.userId) : null;

    const updatedResult = await query(
      `
      UPDATE clients
      SET
        name = $1,
        billing_2025 = $2,
        sector = $3,
        company_type = $4,
        country = $5,
        account_stage = $6,
        manager = $7,
        executive_user_id = $8,
        risk = $9,
        segment = $10,
        service_fixed_fire = $11,
        service_extinguishers = $12,
        service_works = $13,
        wallet_share = $14,
        nps = $15,
        open_opportunities = $16,
        notes = $17,
        supervisor_ifci_user_id = $18,
        supervisor_ext_user_id = $19,
        supervisor_works_user_id = $20
      WHERE id = $21
      RETURNING *
      `,
      [
        name.trim(),
        normalizedBilling,
        sector.trim(),
        companyType,
        country,
        accountStage,
        executiveName,
        Number(executiveUserId),
        risk,
        segment,
        services.fixedFire,
        services.extinguishers,
        services.works,
        "",
        0,
        0,
        notes.trim(),
        fixedFireSupervisorId,
        extinguishersSupervisorId,
        worksSupervisorId,
        clientId
      ]
    );

    const selectedUpdated = await query(buildClientSelectQuery("WHERE clients.id = $1", ""), [updatedResult.rows[0].id]);
    res.json({ client: await buildClientResponse(selectedUpdated.rows[0]) });
  })
);

app.patch(
  "/api/clients/:id/hide",
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const existingClientResult = await query("SELECT id, name FROM clients WHERE id = $1 AND is_hidden = FALSE", [clientId]);

    if (!existingClientResult.rows[0]) {
      res.status(404).json({ error: "Cliente no encontrado" });
      return;
    }

    await query(
      `
      UPDATE clients
      SET is_hidden = TRUE
      WHERE id = $1
      `,
      [clientId]
    );

    res.json({ ok: true });
  })
);

app.post(
  "/api/clients/:id/branches",
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const existingClientResult = await query("SELECT id FROM clients WHERE id = $1", [clientId]);

    if (!existingClientResult.rows[0]) {
      res.status(404).json({ error: "Cliente no encontrado" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateBranchPayload(payload);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const { name, sector, risk, segment, services, supervisors, notes } = payload;
    const parentClientResult = await query("SELECT manager, billing_2025, executive_user_id FROM clients WHERE id = $1", [clientId]);
    const executiveName = parentClientResult.rows[0]?.manager || "";
    const parentBilling = Number(parentClientResult.rows[0]?.billing_2025 || 0);
    const parentExecutiveUserId = parentClientResult.rows[0]?.executive_user_id || null;
    const fixedFireSupervisorId = services.fixedFire ? normalizeNullableUserId(supervisors.fixedFire.userId) : null;
    const extinguishersSupervisorId = services.extinguishers ? normalizeNullableUserId(supervisors.extinguishers.userId) : null;
    const worksSupervisorId = services.works ? normalizeNullableUserId(supervisors.works.userId) : null;

    const result = await query(
      `
      INSERT INTO client_branches (
        client_id, name, billing_2025, sector, manager, executive_user_id, risk, segment,
        service_fixed_fire, service_extinguishers, service_works, notes,
        supervisor_ifci_user_id, supervisor_ext_user_id, supervisor_works_user_id,
        created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, NOW(), NOW())
      RETURNING id
      `,
      [
        clientId,
        name.trim(),
        parentBilling,
        sector.trim(),
        executiveName,
        parentExecutiveUserId,
        risk,
        segment,
        services.fixedFire,
        services.extinguishers,
        services.works,
        notes.trim(),
        fixedFireSupervisorId,
        extinguishersSupervisorId,
        worksSupervisorId
      ]
    );

    const selectedBranch = await query(buildBranchSelectQuery("WHERE client_branches.id = $1", ""), [result.rows[0].id]);
    res.status(201).json({ branch: toBranchObject(selectedBranch.rows[0]) });
  })
);

app.patch(
  "/api/clients/:id/branches/:branchId",
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const branchId = Number(req.params.branchId);
    const existingBranchResult = await query("SELECT id FROM client_branches WHERE id = $1 AND client_id = $2", [branchId, clientId]);

    if (!existingBranchResult.rows[0]) {
      res.status(404).json({ error: "Sucursal no encontrada" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateBranchPayload(payload);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const { name, sector, risk, segment, services, supervisors, notes } = payload;
    const parentClientResult = await query("SELECT manager, billing_2025, executive_user_id FROM clients WHERE id = $1", [clientId]);
    const executiveName = parentClientResult.rows[0]?.manager || "";
    const parentBilling = Number(parentClientResult.rows[0]?.billing_2025 || 0);
    const parentExecutiveUserId = parentClientResult.rows[0]?.executive_user_id || null;
    const fixedFireSupervisorId = services.fixedFire ? normalizeNullableUserId(supervisors.fixedFire.userId) : null;
    const extinguishersSupervisorId = services.extinguishers ? normalizeNullableUserId(supervisors.extinguishers.userId) : null;
    const worksSupervisorId = services.works ? normalizeNullableUserId(supervisors.works.userId) : null;

    const result = await query(
      `
      UPDATE client_branches
      SET
        name = $1,
        billing_2025 = $2,
        sector = $3,
        manager = $4,
        executive_user_id = $5,
        risk = $6,
        segment = $7,
        service_fixed_fire = $8,
        service_extinguishers = $9,
        service_works = $10,
        notes = $11,
        supervisor_ifci_user_id = $12,
        supervisor_ext_user_id = $13,
        supervisor_works_user_id = $14,
        updated_at = NOW()
      WHERE id = $15 AND client_id = $16
      RETURNING id
      `,
      [
        name.trim(),
        parentBilling,
        sector.trim(),
        executiveName,
        parentExecutiveUserId,
        risk,
        segment,
        services.fixedFire,
        services.extinguishers,
        services.works,
        notes.trim(),
        fixedFireSupervisorId,
        extinguishersSupervisorId,
        worksSupervisorId,
        branchId,
        clientId
      ]
    );

    const selectedBranch = await query(buildBranchSelectQuery("WHERE client_branches.id = $1", ""), [result.rows[0].id]);
    res.json({ branch: toBranchObject(selectedBranch.rows[0]) });
  })
);

app.post(
  "/api/clients/:id/opportunities",
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const existingClientResult = await query("SELECT id FROM clients WHERE id = $1 AND is_hidden = FALSE", [clientId]);

    if (!existingClientResult.rows[0]) {
      res.status(404).json({ error: "Cliente no encontrado" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateOpportunityPayload(payload, clientId);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const currentUserResult = await query("SELECT name FROM users WHERE id = $1", [req.session.userId]);
    const currentUser = currentUserResult.rows[0];
    const result = await query(
      `
      INSERT INTO opportunities (
        client_id, branch_id, title, opportunity_type, service_line, status, amount, probability, expected_close_date,
        owner_user_id, source, description, loss_reason, created_by, created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, NOW(), NOW())
      RETURNING id
      `,
      [
        clientId,
        payload.branchId ? Number(payload.branchId) : null,
        payload.title.trim(),
        payload.opportunityType,
        payload.serviceLine,
        payload.status,
        Number(payload.amount),
        Number(payload.probability),
        String(payload.expectedCloseDate || "").trim(),
        payload.ownerUserId ? Number(payload.ownerUserId) : null,
        String(payload.source || "").trim(),
        String(payload.description || "").trim(),
        String(payload.lossReason || "").trim(),
        String(currentUser?.name || "").trim()
      ]
    );

    const opportunities = await fetchOpportunities({ clientId });
    const opportunity = opportunities.find((item) => Number(item.id) === Number(result.rows[0].id));
    res.status(201).json({ opportunity });
  })
);

app.patch(
  "/api/clients/:id/opportunities/:opportunityId",
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const opportunityId = Number(req.params.opportunityId);
    const existingOpportunityResult = await query(
      "SELECT id FROM opportunities WHERE id = $1 AND client_id = $2",
      [opportunityId, clientId]
    );

    if (!existingOpportunityResult.rows[0]) {
      res.status(404).json({ error: "Oportunidad no encontrada" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateOpportunityPayload(payload, clientId);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    await query(
      `
      UPDATE opportunities
      SET
        branch_id = $1,
        title = $2,
        opportunity_type = $3,
        service_line = $4,
        status = $5,
        amount = $6,
        probability = $7,
        expected_close_date = $8,
        owner_user_id = $9,
        source = $10,
        description = $11,
        loss_reason = $12,
        updated_at = NOW()
      WHERE id = $13 AND client_id = $14
      `,
      [
        payload.branchId ? Number(payload.branchId) : null,
        payload.title.trim(),
        payload.opportunityType,
        payload.serviceLine,
        payload.status,
        Number(payload.amount),
        Number(payload.probability),
        String(payload.expectedCloseDate || "").trim(),
        payload.ownerUserId ? Number(payload.ownerUserId) : null,
        String(payload.source || "").trim(),
        String(payload.description || "").trim(),
        String(payload.lossReason || "").trim(),
        opportunityId,
        clientId
      ]
    );

    const opportunities = await fetchOpportunities({ clientId });
    const opportunity = opportunities.find((item) => Number(item.id) === opportunityId);
    res.json({ opportunity });
  })
);

app.post(
  "/api/opportunities/:opportunityId/follow-ups",
  asyncHandler(async (req, res) => {
    const opportunityId = Number(req.params.opportunityId);
    const existingOpportunityResult = await query("SELECT id, client_id FROM opportunities WHERE id = $1", [opportunityId]);
    const existingOpportunity = existingOpportunityResult.rows[0];

    if (!existingOpportunity) {
      res.status(404).json({ error: "Oportunidad no encontrada" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateFollowUpPayload(payload);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const currentUserResult = await query("SELECT name FROM users WHERE id = $1", [req.session.userId]);
    const currentUser = currentUserResult.rows[0];
    const result = await query(
      `
      INSERT INTO opportunity_follow_ups (
        opportunity_id, type, title, due_date, status, assigned_user_id, notes, completed_at, created_by, created_at, updated_at
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, NOW(), NOW())
      RETURNING id
      `,
      [
        opportunityId,
        payload.type,
        payload.title.trim(),
        payload.dueDate.trim(),
        payload.status,
        payload.assignedUserId ? Number(payload.assignedUserId) : null,
        String(payload.notes || "").trim(),
        payload.status === FOLLOW_UP_STATUS.DONE ? new Date().toISOString() : null,
        String(currentUser?.name || "").trim()
      ]
    );

    const opportunities = await fetchOpportunities({ clientId: existingOpportunity.client_id });
    const opportunity = opportunities.find((item) => Number(item.id) === opportunityId);
    const followUp = opportunity?.followUps?.find((item) => Number(item.id) === Number(result.rows[0].id)) || null;
    res.status(201).json({ followUp, opportunity });
  })
);

app.patch(
  "/api/opportunities/:opportunityId/follow-ups/:followUpId",
  asyncHandler(async (req, res) => {
    const opportunityId = Number(req.params.opportunityId);
    const followUpId = Number(req.params.followUpId);
    const existingResult = await query(
      `
      SELECT opportunity_follow_ups.id, opportunities.client_id
      FROM opportunity_follow_ups
      INNER JOIN opportunities ON opportunities.id = opportunity_follow_ups.opportunity_id
      WHERE opportunity_follow_ups.id = $1 AND opportunity_follow_ups.opportunity_id = $2
      `,
      [followUpId, opportunityId]
    );
    const existingFollowUp = existingResult.rows[0];

    if (!existingFollowUp) {
      res.status(404).json({ error: "Seguimiento no encontrado" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateFollowUpPayload(payload);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    await query(
      `
      UPDATE opportunity_follow_ups
      SET
        type = $1,
        title = $2,
        due_date = $3,
        status = $4,
        assigned_user_id = $5,
        notes = $6,
        completed_at = CASE
          WHEN $4 = '${FOLLOW_UP_STATUS.DONE}' THEN COALESCE(completed_at, NOW())
          ELSE NULL
        END,
        updated_at = NOW()
      WHERE id = $7 AND opportunity_id = $8
      `,
      [
        payload.type,
        payload.title.trim(),
        payload.dueDate.trim(),
        payload.status,
        payload.assignedUserId ? Number(payload.assignedUserId) : null,
        String(payload.notes || "").trim(),
        followUpId,
        opportunityId
      ]
    );

    const opportunities = await fetchOpportunities({ clientId: existingFollowUp.client_id });
    const opportunity = opportunities.find((item) => Number(item.id) === opportunityId);
    const followUp = opportunity?.followUps?.find((item) => Number(item.id) === followUpId) || null;
    res.json({ followUp, opportunity });
  })
);

app.post(
  "/api/clients/:id/meetings",
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const existingClientResult = await query("SELECT id FROM clients WHERE id = $1", [clientId]);
    if (!existingClientResult.rows[0]) {
      res.status(404).json({ error: "Cliente no encontrado" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateMeetingPayload(payload);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const branchId = payload.branchId ? Number(payload.branchId) : null;
    if (branchId) {
      const branchResult = await query("SELECT id FROM client_branches WHERE id = $1 AND client_id = $2", [branchId, clientId]);
      if (!branchResult.rows[0]) {
        res.status(400).json({ error: "La sucursal seleccionada no pertenece a esta compañía" });
        return;
      }
    }
    const opportunityId = payload.opportunityId ? Number(payload.opportunityId) : null;
    if (opportunityId) {
      const opportunityResult = await query("SELECT id FROM opportunities WHERE id = $1 AND client_id = $2", [opportunityId, clientId]);
      if (!opportunityResult.rows[0]) {
        res.status(400).json({ error: "La oportunidad seleccionada no pertenece a esta compañía" });
        return;
      }
    }

    const currentUserResult = await query("SELECT name FROM users WHERE id = $1", [req.session.userId]);
    const currentUser = currentUserResult.rows[0];
    const status = normalizeMeetingStatus(payload);
    const nextMeetingDate = String(payload.nextMeetingDate || "").trim();
    const participantUserIds = normalizeParticipantUserIds(payload.participantUserIds);
    const participantNames = await resolveParticipantNames(participantUserIds);

    const createdMeeting = await withTransaction(async (client) => {
      const result = await client.query(
        `
        INSERT INTO meetings (
          client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids,
          contact_name, contact_role, modality, next_meeting_date, minutes, findings,
          active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_status,
          created_by, status, created_at, updated_at
        ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21, $22, NOW(), NOW())
        RETURNING id, client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids, contact_name, contact_role,
                  modality, next_meeting_date, follow_up_from_meeting_id, minutes, findings,
                  active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_status,
                  created_by, status, created_at, updated_at
        `,
        [
          clientId,
          branchId,
          opportunityId,
          payload.kind,
          payload.subject.trim(),
          payload.objective.trim(),
          payload.scheduledFor.trim(),
          participantNames.join(", "),
          participantUserIds,
          payload.contactName.trim(),
          payload.contactRole.trim(),
          payload.modality,
          nextMeetingDate,
          String(payload.minutes || "").trim(),
          String(payload.findings || "").trim(),
          String(payload.activeNegotiationsStatus || "").trim(),
          String(payload.opportunities || "").trim(),
          String(payload.substituteRecovery || "").trim(),
          String(payload.globalContacts || "").trim(),
          String(payload.serviceStatus || "").trim(),
          String(currentUser?.name || "").trim(),
          status
        ]
      );

      const meeting = toMeetingObject(result.rows[0]);
      await upsertMeetingFollowUp(client, meeting, currentUser?.name || "");
      return meeting;
    });

    res.status(201).json({ meeting: createdMeeting });
  })
);

app.patch(
  "/api/clients/:id/meetings/:meetingId",
  asyncHandler(async (req, res) => {
    const clientId = Number(req.params.id);
    const meetingId = Number(req.params.meetingId);
    const existingResult = await query("SELECT id FROM meetings WHERE id = $1 AND client_id = $2", [meetingId, clientId]);

    if (!existingResult.rows[0]) {
      res.status(404).json({ error: "Reunión no encontrada" });
      return;
    }

    const payload = req.body || {};
    const errors = await validateMeetingPayload(payload);
    if (errors.length) {
      res.status(400).json({ error: errors[0] });
      return;
    }

    const branchId = payload.branchId ? Number(payload.branchId) : null;
    if (branchId) {
      const branchResult = await query("SELECT id FROM client_branches WHERE id = $1 AND client_id = $2", [branchId, clientId]);
      if (!branchResult.rows[0]) {
        res.status(400).json({ error: "La sucursal seleccionada no pertenece a esta compañía" });
        return;
      }
    }
    const opportunityId = payload.opportunityId ? Number(payload.opportunityId) : null;
    if (opportunityId) {
      const opportunityResult = await query("SELECT id FROM opportunities WHERE id = $1 AND client_id = $2", [opportunityId, clientId]);
      if (!opportunityResult.rows[0]) {
        res.status(400).json({ error: "La oportunidad seleccionada no pertenece a esta compañía" });
        return;
      }
    }

    const currentUserResult = await query("SELECT name FROM users WHERE id = $1", [req.session.userId]);
    const currentUser = currentUserResult.rows[0];
    const status = normalizeMeetingStatus(payload);
    const nextMeetingDate = String(payload.nextMeetingDate || "").trim();
    const participantUserIds = normalizeParticipantUserIds(payload.participantUserIds);
    const participantNames = await resolveParticipantNames(participantUserIds);

    const updatedMeeting = await withTransaction(async (client) => {
      const result = await client.query(
        `
        UPDATE meetings
        SET
          branch_id = $1,
          opportunity_id = $2,
          kind = $3,
          subject = $4,
          objective = $5,
          scheduled_for = $6,
          participants = $7,
          participant_user_ids = $8,
          contact_name = $9,
          contact_role = $10,
          modality = $11,
          next_meeting_date = $12,
          minutes = $13,
          findings = $14,
          active_negotiations_status = $15,
          opportunities = $16,
          substitute_recovery = $17,
          global_contacts = $18,
          service_status = $19,
          created_by = $20,
          status = $21,
          updated_at = NOW()
        WHERE id = $22 AND client_id = $23
        RETURNING id, client_id, branch_id, opportunity_id, kind, subject, objective, scheduled_for, participants, participant_user_ids, contact_name, contact_role,
                  modality, next_meeting_date, follow_up_from_meeting_id, minutes, findings,
                  active_negotiations_status, opportunities, substitute_recovery, global_contacts, service_status,
                  created_by, status, created_at, updated_at
        `,
        [
          branchId,
          opportunityId,
          payload.kind,
          payload.subject.trim(),
          payload.objective.trim(),
          payload.scheduledFor.trim(),
          participantNames.join(", "),
          participantUserIds,
          payload.contactName.trim(),
          payload.contactRole.trim(),
          payload.modality,
          nextMeetingDate,
          String(payload.minutes || "").trim(),
          String(payload.findings || "").trim(),
          String(payload.activeNegotiationsStatus || "").trim(),
          String(payload.opportunities || "").trim(),
          String(payload.substituteRecovery || "").trim(),
          String(payload.globalContacts || "").trim(),
          String(payload.serviceStatus || "").trim(),
          String(currentUser?.name || "").trim(),
          status,
          meetingId,
          clientId
        ]
      );

      const meeting = toMeetingObject(result.rows[0]);
      await upsertMeetingFollowUp(client, meeting, currentUser?.name || "");
      return meeting;
    });

    res.json({ meeting: updatedMeeting });
  })
);

app.use("/api", (_req, res) => {
  res.status(404).json({ error: "Ruta API no encontrada" });
});

app.use((_req, res) => {
  res.sendFile(path.join(__dirname, "..", "public", "index.html"));
});

app.use((error, _req, res, _next) => {
  console.error(error);
  res.status(500).json({ error: "Error interno del servidor" });
});

async function start() {
  await initDb();
  await refreshMeetingTypesCache();
  await refreshMeetingReasonsCache();
  app.listen(PORT, () => {
    console.log(`Servidor iniciado en http://localhost:${PORT}`);
  });
}

start().catch((error) => {
  console.error("No se pudo iniciar el servidor:", error);
  process.exit(1);
});
