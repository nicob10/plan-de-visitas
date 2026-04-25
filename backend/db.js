const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const Database = require("better-sqlite3");
const { Pool } = require("pg");
const {
  MEETING_TYPES,
  MEETING_REASONS,
  USER_ROLES,
  USER_ROLE_OPTIONS,
  OPPORTUNITY_STATUS_OPTIONS,
  FOLLOW_UP_STATUS_OPTIONS,
  OPPORTUNITY_SERVICE_LINES,
  FOLLOW_UP_TYPES,
  OPPORTUNITY_TYPES
} = require("./constants");

const DEFAULT_DB_USER = encodeURIComponent(process.env.PGUSER || process.env.USER || "postgres");
const DATABASE_URL = process.env.DATABASE_URL || `postgres://${DEFAULT_DB_USER}@localhost:5432/top_clients`;
const LEGACY_DB_PATH = path.join(__dirname, "..", "data", "top_clients.sqlite");

const pool = new Pool({
  connectionString: DATABASE_URL,
  ssl: process.env.PGSSLMODE === "require" ? { rejectUnauthorized: false } : false
});

const DEFAULT_SECTOR_OPTIONS = [
  "Alimentos y Bebidas",
  "Automotriz",
  "Centros Comerciales",
  "Construcción",
  "Data Centers",
  "Depósitos y Logística",
  "Educación",
  "Energía",
  "Farmacéutica",
  "Hotelería y Gastronomía",
  "Manufactura",
  "Minería",
  "Oil & Gas",
  "Papel y Celulosa",
  "Petroquímica",
  "Puertos",
  "Química",
  "Retail",
  "Salud",
  "Telecomunicaciones",
  "Transporte",
  "Utilities"
];

const DEFAULT_CONTACT_ROLE_OPTIONS = ["Usuario", "Comprador", "Nuevo contacto"];

function hashPassword(password) {
  return crypto.createHash("sha256").update(password).digest("hex");
}

async function query(text, params = []) {
  const result = await pool.query(text, params);
  return result;
}

async function withTransaction(callback) {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    const result = await callback(client);
    await client.query("COMMIT");
    return result;
  } catch (error) {
    await client.query("ROLLBACK");
    throw error;
  } finally {
    client.release();
  }
}

function inferMeetingKind(text) {
  const source = String(text || "").toLowerCase();
  if (source.includes("directorio")) return "Directorio";
  if (source.includes("gerencia")) return "Gerencia Comercial";
  if (source.includes("operativa") || source.includes("operaciones")) return "Operaciones";
  return "Comercial";
}

async function seedDefaultUsers() {
  const existingUsers = await query("SELECT COUNT(*)::int AS count FROM users");
  if (Number(existingUsers.rows[0]?.count || 0) > 0) return;

  const users = [
    ["Admin MAXI", "admin@maxi.local", "admin1234", USER_ROLES.COMMERCIAL_MANAGER],
    ["Comercial Maxi", "comercial@maxi.local", "comercial1234", USER_ROLES.EXECUTIVE],
    ["Supervisor IFCI", "ifci@maxi.local", "ifci1234", USER_ROLES.SUPERVISOR_IFCI],
    ["Supervisor EXT", "ext@maxi.local", "ext1234", USER_ROLES.SUPERVISOR_EXT],
    ["Supervisor Obra", "obra@maxi.local", "obra1234", USER_ROLES.SUPERVISOR_WORKS],
    ["Gerencia Comercial", "gerencia.comercial@maxi.local", "gerencia1234", USER_ROLES.COMMERCIAL_MANAGER]
  ];

  await withTransaction(async (client) => {
    for (const [name, email, password, role] of users) {
      await client.query(
        `
        INSERT INTO users (name, email, password_hash, role)
        VALUES ($1, $2, $3, $4)
        ON CONFLICT (email) DO NOTHING
        `,
        [name, email, hashPassword(password), role]
      );
    }
  });
}

async function seedSectorOptions() {
  await withTransaction(async (client) => {
    for (const sectorName of DEFAULT_SECTOR_OPTIONS) {
      await client.query(
        `
        INSERT INTO sector_options (name)
        VALUES ($1)
        ON CONFLICT (name) DO NOTHING
        `,
        [sectorName]
      );
    }

    await client.query(
      `
      INSERT INTO sector_options (name)
      SELECT DISTINCT TRIM(sector)
      FROM clients
      WHERE COALESCE(TRIM(sector), '') <> ''
      ON CONFLICT (name) DO NOTHING
      `
    );
  });
}

async function seedMeetingTypeOptions() {
  const existingTypes = await query("SELECT COUNT(*)::int AS count FROM meeting_type_options");
  if (Number(existingTypes.rows[0]?.count || 0) > 0) return;

  await withTransaction(async (client) => {
    for (const [index, type] of MEETING_TYPES.entries()) {
      await client.query(
        `
        INSERT INTO meeting_type_options (value, label, color, sort_order)
        VALUES ($1, $2, $3, $4)
        ON CONFLICT (value) DO NOTHING
        `,
        [type.value, type.label, type.color, index + 1]
      );
    }
  });
}

async function seedMeetingReasonOptions() {
  const existingReasons = await query("SELECT COUNT(*)::int AS count FROM meeting_reason_options");
  if (Number(existingReasons.rows[0]?.count || 0) > 0) return;

  await withTransaction(async (client) => {
    for (const [index, reason] of MEETING_REASONS.entries()) {
      await client.query(
        `
        INSERT INTO meeting_reason_options (name, sort_order)
        VALUES ($1, $2)
        ON CONFLICT (name) DO NOTHING
        `,
        [reason, index + 1]
      );
    }
  });
}

async function seedContactRoleOptions() {
  const existingRoles = await query("SELECT COUNT(*)::int AS count FROM contact_role_options");
  if (Number(existingRoles.rows[0]?.count || 0) > 0) return;

  await withTransaction(async (client) => {
    for (const [index, role] of DEFAULT_CONTACT_ROLE_OPTIONS.entries()) {
      await client.query(
        `
        INSERT INTO contact_role_options (name, sort_order)
        VALUES ($1, $2)
        ON CONFLICT (name) DO NOTHING
        `,
        [role, index + 1]
      );
    }
  });
}

function readLegacyMeetings(legacyDb) {
  const meetingsTable = legacyDb
    .prepare("SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'meetings'")
    .get();

  if (meetingsTable) {
    const columns = legacyDb.prepare("PRAGMA table_info(meetings)").all().map((column) => column.name);
    const selectColumn = (name, fallbackSql) => (columns.includes(name) ? name : `${fallbackSql} AS ${name}`);

    return legacyDb
      .prepare(
        `
        SELECT
          ${selectColumn("id", "NULL")},
          ${selectColumn("client_id", "NULL")},
          ${selectColumn("kind", "'Comercial'")},
          ${selectColumn("subject", "''")},
          ${selectColumn("objective", "''")},
          ${selectColumn("scheduled_for", "''")},
          ${selectColumn("participants", "''")},
          ${selectColumn("contact_name", "''")},
          ${selectColumn("contact_role", "''")},
          ${selectColumn("modality", "'Presencial'")},
          ${selectColumn("next_meeting_date", "''")},
          ${selectColumn("follow_up_from_meeting_id", "NULL")},
          ${selectColumn("minutes", "''")},
          ${selectColumn("findings", "''")},
          ${selectColumn("active_negotiations_status", "''")},
          ${selectColumn("opportunities", "''")},
          ${selectColumn("substitute_recovery", "''")},
          ${selectColumn("global_contacts", "''")},
          ${selectColumn("service_health_status", "''")},
          ${selectColumn("service_status", "''")},
          ${selectColumn("complaint_bitrix_responsible_id", "''")},
          ${selectColumn("complaint_bitrix_task_id", "''")},
          ${selectColumn("created_by", "''")},
          ${selectColumn("status", "'Agendada'")},
          ${selectColumn("created_at", "CURRENT_TIMESTAMP")},
          ${selectColumn("updated_at", "CURRENT_TIMESTAMP")}
        FROM meetings
        ORDER BY id ASC
        `
      )
      .all();
  }

  const visitsTable = legacyDb
    .prepare("SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'visits'")
    .get();

  if (!visitsTable) return [];

  return legacyDb
    .prepare(
      `
      SELECT id, client_id, type, done, date, minutes, findings
      FROM visits
      WHERE COALESCE(date, '') <> ''
         OR COALESCE(minutes, '') <> ''
         OR COALESCE(findings, '') <> ''
         OR done = 1
      ORDER BY id ASC
      `
    )
    .all()
    .map((visit) => ({
      id: visit.id,
      client_id: visit.client_id,
      kind: inferMeetingKind(visit.type),
      subject: visit.type,
      objective: visit.type,
      scheduled_for: visit.date || "",
      participants: "",
      contact_name: "",
      contact_role: "",
      modality: "Presencial",
      next_meeting_date: "",
      follow_up_from_meeting_id: null,
      minutes: visit.minutes || "",
      findings: visit.findings || "",
      active_negotiations_status: visit.findings || "",
      opportunities: "",
      substitute_recovery: "",
      global_contacts: "",
      service_health_status: "",
      service_status: "",
      complaint_bitrix_responsible_id: "",
      complaint_bitrix_task_id: "",
      created_by: "",
      status: visit.done ? "Realizada" : "Agendada",
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    }));
}

async function syncSequence(client, tableName) {
  await client.query(
    `
    SELECT setval(
      pg_get_serial_sequence($1, 'id'),
      GREATEST(COALESCE((SELECT MAX(id) FROM ${tableName}), 0), 1),
      (SELECT COUNT(*) > 0 FROM ${tableName})
    )
    `,
    [tableName]
  );
}

async function migrateLegacySqliteIfNeeded() {
  if (!fs.existsSync(LEGACY_DB_PATH)) return;

  const { rows } = await query("SELECT COUNT(*)::int AS total FROM clients");
  if ((rows[0]?.total || 0) > 0) return;

  const legacyDb = new Database(LEGACY_DB_PATH, { readonly: true });

  try {
    const clients = legacyDb
      .prepare(
        `
        SELECT
          id, position, name, billing_2025, sector, manager, risk, segment,
          service_fixed_fire, service_extinguishers, service_works,
          wallet_share, nps, open_opportunities, notes
        FROM clients
        ORDER BY id ASC
        `
      )
      .all();

    const usersTable = legacyDb
      .prepare("SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'users'")
      .get();
    const users = usersTable
      ? legacyDb
          .prepare(
            `
            SELECT id, name, email, password_hash, role, created_at
            FROM users
            ORDER BY id ASC
            `
          )
          .all()
      : [];

    const meetings = readLegacyMeetings(legacyDb);

    await withTransaction(async (client) => {
      for (const row of clients) {
        await client.query(
          `
          INSERT INTO clients (
            id, position, name, billing_2025, sector, manager, risk, segment,
            company_type, country, account_stage,
            service_fixed_fire, service_extinguishers, service_works,
            wallet_share, nps, open_opportunities, notes,
            executive_user_id, supervisor_ifci_user_id, supervisor_ext_user_id, supervisor_works_user_id
          ) VALUES (
            $1, $2, $3, $4, $5, $6, $7, $8,
            $9, $10, $11, $12, $13, $14,
            $15, $16, $17, $18, $19, $20,
            $21, $22
          )
          ON CONFLICT (id) DO NOTHING
          `,
          [
            row.id,
            row.position,
            row.name,
            row.billing_2025,
            row.sector,
            row.manager,
            row.risk,
            row.segment,
            row.company_type || "Local",
            row.country || "Argentina",
            row.account_stage || "Activa",
            !!row.service_fixed_fire,
            !!row.service_extinguishers,
            !!row.service_works,
            row.wallet_share,
            row.nps,
            row.open_opportunities,
            row.notes || "",
            null,
            null,
            null,
            null
          ]
        );
      }

      for (const row of users) {
        const legacyRoleMap = {
          Administrador: USER_ROLES.COMMERCIAL_MANAGER,
          Comercial: USER_ROLES.EXECUTIVE,
          Operativo: USER_ROLES.SUPERVISOR_IFCI
        };
        await client.query(
          `
          INSERT INTO users (id, name, email, password_hash, role, created_at)
          VALUES ($1, $2, $3, $4, $5, COALESCE($6::timestamptz, NOW()))
          ON CONFLICT (id) DO NOTHING
          `,
          [row.id, row.name, row.email, row.password_hash, legacyRoleMap[row.role] || row.role, row.created_at || null]
        );
      }

      for (const row of meetings) {
        await client.query(
          `
          INSERT INTO meetings (
            id, client_id, kind, subject, objective, scheduled_for, participants,
            contact_name, contact_role, modality, next_meeting_date, follow_up_from_meeting_id,
            minutes, findings, active_negotiations_status, opportunities, substitute_recovery,
            global_contacts, service_health_status, service_status, complaint_bitrix_responsible_id, complaint_bitrix_task_id,
            created_by, status, created_at, updated_at
          ) VALUES (
            $1, $2, $3, $4, $5, $6, $7,
            $8, $9, $10, $11, $12,
            $13, $14, $15, $16, $17, $18, $19, $20, $21, $22,
            $23, $24, COALESCE($25::timestamptz, NOW()), COALESCE($26::timestamptz, NOW())
          )
          ON CONFLICT (id) DO NOTHING
          `,
          [
            row.id,
            row.client_id,
            MEETING_TYPES.some((type) => type.value === row.kind) ? row.kind : "Comercial",
            row.subject,
            row.objective || "",
            row.scheduled_for || "",
            row.participants || "",
            row.contact_name || "",
            row.contact_role || "",
            row.modality || "Presencial",
            row.next_meeting_date || "",
            row.follow_up_from_meeting_id || null,
            row.minutes || "",
            row.findings || "",
            row.active_negotiations_status || row.findings || "",
            row.opportunities || "",
            row.substitute_recovery || "",
            row.global_contacts || "",
            row.service_health_status || "",
            row.service_status || "",
            row.complaint_bitrix_responsible_id || "",
            row.complaint_bitrix_task_id || "",
            row.created_by || "",
            row.status === "Realizada" ? "Realizada" : "Agendada",
            row.created_at || null,
            row.updated_at || null
          ]
        );
      }

      await syncSequence(client, "clients");
      await syncSequence(client, "users");
      await syncSequence(client, "meetings");
    });
  } finally {
    legacyDb.close();
  }
}

async function initDb() {
  await query(`
    CREATE TABLE IF NOT EXISTS users (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL,
      email TEXT NOT NULL UNIQUE,
      password_hash TEXT NOT NULL,
      role TEXT NOT NULL DEFAULT '${USER_ROLES.EXECUTIVE}',
      can_create_clients BOOLEAN NOT NULL DEFAULT TRUE,
      can_edit_clients BOOLEAN NOT NULL DEFAULT TRUE,
      can_hide_clients BOOLEAN NOT NULL DEFAULT TRUE,
      can_create_branches BOOLEAN NOT NULL DEFAULT TRUE,
      can_edit_branches BOOLEAN NOT NULL DEFAULT TRUE,
      can_hide_branches BOOLEAN NOT NULL DEFAULT TRUE,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS clients (
      id SERIAL PRIMARY KEY,
      position INTEGER NOT NULL UNIQUE,
      name TEXT NOT NULL,
      billing_2025 NUMERIC(14, 2) NOT NULL,
      sector TEXT NOT NULL,
      manager TEXT NOT NULL,
      risk TEXT NOT NULL CHECK (risk IN ('Bajo', 'Medio', 'Alto')),
      segment TEXT NOT NULL CHECK (segment IN ('A', 'B', 'C')),
      company_type TEXT NOT NULL DEFAULT 'Local',
      country TEXT NOT NULL DEFAULT 'Argentina',
      account_stage TEXT NOT NULL DEFAULT 'Activa',
      is_hidden BOOLEAN NOT NULL DEFAULT FALSE,
      service_fixed_fire BOOLEAN NOT NULL DEFAULT FALSE,
      service_extinguishers BOOLEAN NOT NULL DEFAULT FALSE,
      service_works BOOLEAN NOT NULL DEFAULT FALSE,
      wallet_share TEXT NOT NULL,
      nps INTEGER NOT NULL,
      open_opportunities INTEGER NOT NULL,
      notes TEXT NOT NULL DEFAULT '',
      executive_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
      supervisor_ifci_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
      supervisor_ext_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
      supervisor_works_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS client_branches (
      id SERIAL PRIMARY KEY,
      client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
      name TEXT NOT NULL,
      billing_2025 NUMERIC(14, 2) NOT NULL DEFAULT 0,
      bitrix_company_id TEXT NOT NULL DEFAULT '',
      manual_branch_id TEXT NOT NULL DEFAULT '',
      is_hidden BOOLEAN NOT NULL DEFAULT FALSE,
      sector TEXT NOT NULL,
      manager TEXT NOT NULL,
      risk TEXT NOT NULL CHECK (risk IN ('Bajo', 'Medio', 'Alto')),
      segment TEXT NOT NULL CHECK (segment IN ('A', 'B', 'C')),
      service_fixed_fire BOOLEAN NOT NULL DEFAULT FALSE,
      service_extinguishers BOOLEAN NOT NULL DEFAULT FALSE,
      service_works BOOLEAN NOT NULL DEFAULT FALSE,
      notes TEXT NOT NULL DEFAULT '',
      supervisor_ifci_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
      supervisor_ext_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
      supervisor_works_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS visit_rules (
      id SERIAL PRIMARY KEY,
      entity_type TEXT NOT NULL,
      entity_id INTEGER NOT NULL,
      periodicity_days INTEGER NOT NULL DEFAULT 30,
      contact_role TEXT NOT NULL DEFAULT '',
      objective TEXT NOT NULL DEFAULT '',
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS meetings (
      id SERIAL PRIMARY KEY,
      visit_id TEXT NOT NULL DEFAULT '',
      client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
      branch_id INTEGER REFERENCES client_branches(id) ON DELETE CASCADE,
      opportunity_id INTEGER,
      kind TEXT NOT NULL DEFAULT 'Comercial',
      subject TEXT NOT NULL,
      objective TEXT NOT NULL DEFAULT '',
      scheduled_for TEXT NOT NULL DEFAULT '',
      participants TEXT NOT NULL DEFAULT '',
      participant_user_ids INTEGER[] NOT NULL DEFAULT '{}'::INTEGER[],
      contact_name TEXT NOT NULL DEFAULT '',
      contact_role TEXT NOT NULL DEFAULT '',
      modality TEXT NOT NULL DEFAULT 'Presencial',
      next_meeting_date TEXT NOT NULL DEFAULT '',
      follow_up_from_meeting_id INTEGER REFERENCES meetings(id) ON DELETE SET NULL,
      minutes TEXT NOT NULL DEFAULT '',
      findings TEXT NOT NULL DEFAULT '',
      active_negotiations_status TEXT NOT NULL DEFAULT '',
      opportunities TEXT NOT NULL DEFAULT '',
      substitute_recovery TEXT NOT NULL DEFAULT '',
      global_contacts TEXT NOT NULL DEFAULT '',
      service_health_status TEXT NOT NULL DEFAULT '',
      service_status TEXT NOT NULL DEFAULT '',
      complaint_bitrix_responsible_id TEXT NOT NULL DEFAULT '',
      complaint_bitrix_task_id TEXT NOT NULL DEFAULT '',
      created_by TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT 'Agendada' CHECK (status IN ('Agendada', 'Realizada')),
      is_deleted BOOLEAN NOT NULL DEFAULT FALSE,
      deleted_at TIMESTAMPTZ,
      deleted_by TEXT NOT NULL DEFAULT '',
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS audit_logs (
      id SERIAL PRIMARY KEY,
      user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
      user_name TEXT NOT NULL DEFAULT '',
      user_email TEXT NOT NULL DEFAULT '',
      action TEXT NOT NULL,
      entity_type TEXT NOT NULL DEFAULT '',
      entity_id INTEGER,
      entity_name TEXT NOT NULL DEFAULT '',
      details JSONB NOT NULL DEFAULT '{}'::jsonb,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS opportunities (
      id SERIAL PRIMARY KEY,
      client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
      branch_id INTEGER REFERENCES client_branches(id) ON DELETE SET NULL,
      title TEXT NOT NULL,
      opportunity_type TEXT NOT NULL DEFAULT 'Negociación',
      service_line TEXT NOT NULL DEFAULT 'Otro',
      status TEXT NOT NULL DEFAULT 'Abierta',
      amount NUMERIC(14, 2) NOT NULL DEFAULT 0,
      probability INTEGER NOT NULL DEFAULT 0,
      expected_close_date TEXT NOT NULL DEFAULT '',
      owner_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
      source TEXT NOT NULL DEFAULT '',
      description TEXT NOT NULL DEFAULT '',
      loss_reason TEXT NOT NULL DEFAULT '',
      created_by TEXT NOT NULL DEFAULT '',
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS opportunity_follow_ups (
      id SERIAL PRIMARY KEY,
      opportunity_id INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
      type TEXT NOT NULL DEFAULT 'Recordatorio',
      title TEXT NOT NULL,
      due_date TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT 'Pendiente',
      assigned_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
      notes TEXT NOT NULL DEFAULT '',
      completed_at TIMESTAMPTZ,
      created_by TEXT NOT NULL DEFAULT '',
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS sector_options (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL UNIQUE,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS meeting_type_options (
      id SERIAL PRIMARY KEY,
      value TEXT NOT NULL UNIQUE,
      label TEXT NOT NULL,
      color TEXT NOT NULL DEFAULT 'yellow',
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS meeting_reason_options (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL UNIQUE,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS contact_role_options (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL UNIQUE,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS app_settings (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL DEFAULT '',
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE INDEX IF NOT EXISTS idx_clients_filters
      ON clients(risk, segment, service_fixed_fire, service_extinguishers, service_works);
    CREATE INDEX IF NOT EXISTS idx_branches_client ON client_branches(client_id);
    CREATE INDEX IF NOT EXISTS idx_visit_rules_entity ON visit_rules(entity_type, entity_id, id);
    CREATE INDEX IF NOT EXISTS idx_meetings_client ON meetings(client_id);
    CREATE INDEX IF NOT EXISTS idx_meetings_schedule ON meetings(scheduled_for, status);
    CREATE INDEX IF NOT EXISTS idx_opportunities_client ON opportunities(client_id);
    CREATE INDEX IF NOT EXISTS idx_opportunities_status ON opportunities(status, expected_close_date);
    CREATE INDEX IF NOT EXISTS idx_opportunities_owner ON opportunities(owner_user_id);
    CREATE INDEX IF NOT EXISTS idx_follow_ups_opportunity ON opportunity_follow_ups(opportunity_id, due_date);
    CREATE INDEX IF NOT EXISTS idx_sector_options_name ON sector_options(name);
    CREATE INDEX IF NOT EXISTS idx_meeting_type_options_sort ON meeting_type_options(sort_order, id);
    CREATE INDEX IF NOT EXISTS idx_meeting_reason_options_sort ON meeting_reason_options(sort_order, id);
    CREATE INDEX IF NOT EXISTS idx_contact_role_options_sort ON contact_role_options(sort_order, id);
    CREATE INDEX IF NOT EXISTS idx_audit_logs_created_at ON audit_logs(created_at DESC);
    CREATE INDEX IF NOT EXISTS idx_audit_logs_user_id ON audit_logs(user_id, created_at DESC);
  `);

  await query(`
    ALTER TABLE users
      ADD COLUMN IF NOT EXISTS bitrix_user_id TEXT NOT NULL DEFAULT '';
    ALTER TABLE users
      ADD COLUMN IF NOT EXISTS can_create_clients BOOLEAN NOT NULL DEFAULT TRUE;
    ALTER TABLE users
      ADD COLUMN IF NOT EXISTS can_edit_clients BOOLEAN NOT NULL DEFAULT TRUE;
    ALTER TABLE users
      ADD COLUMN IF NOT EXISTS can_hide_clients BOOLEAN NOT NULL DEFAULT TRUE;
    ALTER TABLE users
      ADD COLUMN IF NOT EXISTS can_create_branches BOOLEAN NOT NULL DEFAULT TRUE;
    ALTER TABLE users
      ADD COLUMN IF NOT EXISTS can_edit_branches BOOLEAN NOT NULL DEFAULT TRUE;
    ALTER TABLE users
      ADD COLUMN IF NOT EXISTS can_hide_branches BOOLEAN NOT NULL DEFAULT TRUE;
    ALTER TABLE clients
      ADD COLUMN IF NOT EXISTS executive_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL;
    ALTER TABLE clients
      ADD COLUMN IF NOT EXISTS supervisor_ifci_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL;
    ALTER TABLE clients
      ADD COLUMN IF NOT EXISTS supervisor_ext_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL;
    ALTER TABLE clients
      ADD COLUMN IF NOT EXISTS supervisor_works_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL;
    ALTER TABLE clients
      ADD COLUMN IF NOT EXISTS company_type TEXT NOT NULL DEFAULT 'Local';
    ALTER TABLE clients
      ADD COLUMN IF NOT EXISTS country TEXT NOT NULL DEFAULT 'Argentina';
    ALTER TABLE clients
      ADD COLUMN IF NOT EXISTS account_stage TEXT NOT NULL DEFAULT 'Activa';
    ALTER TABLE clients
      ADD COLUMN IF NOT EXISTS is_hidden BOOLEAN NOT NULL DEFAULT FALSE;
    ALTER TABLE clients
      ADD COLUMN IF NOT EXISTS bitrix_lead_id TEXT NOT NULL DEFAULT '';
    ALTER TABLE clients
      ADD COLUMN IF NOT EXISTS bitrix_company_id TEXT NOT NULL DEFAULT '';
    CREATE INDEX IF NOT EXISTS idx_clients_executive ON clients(executive_user_id);
  `);

  await query(`
    ALTER TABLE client_branches
      ADD COLUMN IF NOT EXISTS billing_2025 NUMERIC(14, 2) NOT NULL DEFAULT 0;
    ALTER TABLE client_branches
      ADD COLUMN IF NOT EXISTS executive_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL;
    ALTER TABLE client_branches
      ADD COLUMN IF NOT EXISTS bitrix_company_id TEXT NOT NULL DEFAULT '';
    ALTER TABLE client_branches
      ADD COLUMN IF NOT EXISTS manual_branch_id TEXT NOT NULL DEFAULT '';
    ALTER TABLE client_branches
      ADD COLUMN IF NOT EXISTS is_hidden BOOLEAN NOT NULL DEFAULT FALSE;
    ALTER TABLE client_branches
      ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ NOT NULL DEFAULT NOW();
    ALTER TABLE client_branches
      ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW();
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS visit_id TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS branch_id INTEGER REFERENCES client_branches(id) ON DELETE CASCADE;
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS opportunity_id INTEGER;
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS participant_user_ids INTEGER[] NOT NULL DEFAULT '{}'::INTEGER[];
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS contact_name TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS contact_role TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS modality TEXT NOT NULL DEFAULT 'Presencial';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS next_meeting_date TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS follow_up_from_meeting_id INTEGER REFERENCES meetings(id) ON DELETE SET NULL;
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS active_negotiations_status TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS opportunities TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS substitute_recovery TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS global_contacts TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS service_health_status TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS service_status TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS complaint_bitrix_responsible_id TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS complaint_bitrix_task_id TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS is_deleted BOOLEAN NOT NULL DEFAULT FALSE;
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS deleted_at TIMESTAMPTZ;
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS deleted_by TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS is_automatic BOOLEAN NOT NULL DEFAULT FALSE;
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS automatic_rule_id INTEGER;
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS automatic_rule_entity_type TEXT NOT NULL DEFAULT '';
    ALTER TABLE meetings
      ADD COLUMN IF NOT EXISTS automatic_rule_entity_id INTEGER;
    CREATE INDEX IF NOT EXISTS idx_client_branches_executive ON client_branches(executive_user_id);
    CREATE INDEX IF NOT EXISTS idx_client_branches_hidden ON client_branches(client_id, is_hidden);
    CREATE INDEX IF NOT EXISTS idx_meetings_opportunity ON meetings(opportunity_id);
    CREATE INDEX IF NOT EXISTS idx_meetings_deleted ON meetings(is_deleted, deleted_at);
    CREATE INDEX IF NOT EXISTS idx_meetings_automatic_rule ON meetings(automatic_rule_id, is_deleted, status);
  `);

  await query(`
    ALTER TABLE meetings
      DROP CONSTRAINT IF EXISTS meetings_opportunity_id_fkey;
    ALTER TABLE meetings
      ADD CONSTRAINT meetings_opportunity_id_fkey
      FOREIGN KEY (opportunity_id) REFERENCES opportunities(id) ON DELETE SET NULL;

    ALTER TABLE opportunities
      ADD COLUMN IF NOT EXISTS opportunity_type TEXT NOT NULL DEFAULT 'Negociación';
    ALTER TABLE opportunities
      DROP CONSTRAINT IF EXISTS opportunities_status_check;
    ALTER TABLE opportunities
      ADD CONSTRAINT opportunities_status_check CHECK (status = ANY(ARRAY['${OPPORTUNITY_STATUS_OPTIONS.join("','")}']::text[]));
    ALTER TABLE opportunities
      DROP CONSTRAINT IF EXISTS opportunities_opportunity_type_check;
    ALTER TABLE opportunities
      ADD CONSTRAINT opportunities_opportunity_type_check CHECK (opportunity_type = ANY(ARRAY['${OPPORTUNITY_TYPES.join("','")}']::text[]));
    ALTER TABLE opportunities
      DROP CONSTRAINT IF EXISTS opportunities_service_line_check;
    ALTER TABLE opportunities
      ADD CONSTRAINT opportunities_service_line_check CHECK (service_line = ANY(ARRAY['${OPPORTUNITY_SERVICE_LINES.join("','")}']::text[]));
    ALTER TABLE opportunities
      DROP CONSTRAINT IF EXISTS opportunities_probability_check;
    ALTER TABLE opportunities
      ADD CONSTRAINT opportunities_probability_check CHECK (probability >= 0 AND probability <= 100);
    ALTER TABLE opportunities
      DROP CONSTRAINT IF EXISTS opportunities_amount_check;
    ALTER TABLE opportunities
      ADD CONSTRAINT opportunities_amount_check CHECK (amount >= 0);

    ALTER TABLE opportunity_follow_ups
      DROP CONSTRAINT IF EXISTS opportunity_follow_ups_status_check;
    ALTER TABLE opportunity_follow_ups
      ADD CONSTRAINT opportunity_follow_ups_status_check CHECK (status = ANY(ARRAY['${FOLLOW_UP_STATUS_OPTIONS.join("','")}']::text[]));
    ALTER TABLE opportunity_follow_ups
      DROP CONSTRAINT IF EXISTS opportunity_follow_ups_type_check;
    ALTER TABLE opportunity_follow_ups
      ADD CONSTRAINT opportunity_follow_ups_type_check CHECK (type = ANY(ARRAY['${FOLLOW_UP_TYPES.join("','")}']::text[]));
  `);

  await query(`
    ALTER TABLE meetings
      DROP CONSTRAINT IF EXISTS meetings_status_check;
    ALTER TABLE meetings
      ADD CONSTRAINT meetings_status_check CHECK (status IN ('Agendada', 'Confirmada', 'Realizada'));
  `);

  await query(`
    UPDATE clients
    SET country = 'Argentina'
    WHERE COALESCE(country, '') = '';
  `);

  await query(`
    UPDATE clients
    SET company_type = 'Local'
    WHERE COALESCE(company_type, '') NOT IN ('Local', 'Global');
  `);

  await query(`
    UPDATE clients
    SET account_stage = 'Activa'
    WHERE COALESCE(account_stage, '') NOT IN ('Activa', 'Prospecto');
  `);

  await query(`
    UPDATE meetings
    SET contact_name = participants
    WHERE COALESCE(contact_name, '') = '' AND COALESCE(participants, '') <> '';
  `);

  await query(`
    UPDATE meetings
    SET visit_id = 'VIS-' || LPAD(id::text, 6, '0')
    WHERE COALESCE(visit_id, '') = '';
  `);

  await query(`
    CREATE UNIQUE INDEX IF NOT EXISTS idx_meetings_visit_id ON meetings(visit_id);
  `);

  await query(`
    CREATE OR REPLACE FUNCTION set_meetings_visit_id()
    RETURNS trigger AS $$
    BEGIN
      IF COALESCE(NEW.visit_id, '') = '' THEN
        NEW.visit_id := 'VIS-' || LPAD(NEW.id::text, 6, '0');
      END IF;
      RETURN NEW;
    END;
    $$ LANGUAGE plpgsql;
  `);

  await query(`
    DROP TRIGGER IF EXISTS trg_meetings_visit_id ON meetings;
    CREATE TRIGGER trg_meetings_visit_id
    BEFORE INSERT ON meetings
    FOR EACH ROW
    EXECUTE FUNCTION set_meetings_visit_id();
  `);

  await query(`
    UPDATE meetings
    SET active_negotiations_status = findings
    WHERE COALESCE(active_negotiations_status, '') = '' AND COALESCE(findings, '') <> '';
  `);

  await query(`
    UPDATE meetings
    SET modality = 'Presencial'
    WHERE COALESCE(modality, '') NOT IN ('Virtual', 'Presencial');
  `);

  await query(`
    UPDATE opportunities
    SET opportunity_type = 'Negociación'
    WHERE COALESCE(opportunity_type, '') NOT IN ('${OPPORTUNITY_TYPES.join("','")}');
  `);

  await query(`
    UPDATE opportunities
    SET service_line = 'Otro'
    WHERE COALESCE(service_line, '') NOT IN ('${OPPORTUNITY_SERVICE_LINES.join("','")}');
  `);

  await query(`
    UPDATE opportunities
    SET status = 'Abierta'
    WHERE COALESCE(status, '') NOT IN ('${OPPORTUNITY_STATUS_OPTIONS.join("','")}');
  `);

  await query(`
    UPDATE opportunities
    SET probability = LEAST(GREATEST(COALESCE(probability, 0), 0), 100);
  `);

  await query(`
    UPDATE opportunity_follow_ups
    SET status = 'Pendiente'
    WHERE COALESCE(status, '') NOT IN ('${FOLLOW_UP_STATUS_OPTIONS.join("','")}');
  `);

  await query(`
    UPDATE opportunity_follow_ups
    SET type = 'Recordatorio'
    WHERE COALESCE(type, '') NOT IN ('${FOLLOW_UP_TYPES.join("','")}');
  `);

  await migrateLegacySqliteIfNeeded();
  await seedDefaultUsers();
  await normalizeRolesAndAssignments();
  await backfillMeetingParticipants();
  await seedSectorOptions();
  await seedMeetingTypeOptions();
  await seedMeetingReasonOptions();
  await seedContactRoleOptions();
}

async function normalizeRolesAndAssignments() {
  const legacyRoleMap = {
    Administrador: USER_ROLES.COMMERCIAL_MANAGER,
    Comercial: USER_ROLES.EXECUTIVE,
    Operativo: USER_ROLES.SUPERVISOR_IFCI
  };

  for (const [legacyRole, nextRole] of Object.entries(legacyRoleMap)) {
    await query("UPDATE users SET role = $1 WHERE role = $2", [nextRole, legacyRole]);
  }

  await query(
    `
    UPDATE users
    SET role = $1
    WHERE role <> ALL($2::text[])
    `,
    [USER_ROLES.EXECUTIVE, USER_ROLE_OPTIONS]
  );

  const executivesResult = await query(
    `
    SELECT id, name
    FROM users
    WHERE role = $1
    ORDER BY id ASC
    `,
    [USER_ROLES.EXECUTIVE]
  );
  const executives = executivesResult.rows;
  if (!executives.length) return;

  const companiesResult = await query(
    `
    SELECT id, manager, executive_user_id
    FROM clients
    WHERE executive_user_id IS NULL
    ORDER BY id ASC
    `
  );

  for (const company of companiesResult.rows) {
    const matchedExecutive =
      executives.find((user) => user.name.trim().toLowerCase() === String(company.manager || "").trim().toLowerCase()) ||
      executives[company.id % executives.length];

    await query(
      `
      UPDATE clients
      SET executive_user_id = $1
      WHERE id = $2
      `,
      [matchedExecutive.id, company.id]
    );
  }
}

async function backfillMeetingParticipants() {
  const usersResult = await query("SELECT id, name FROM users ORDER BY id ASC");
  const usersByName = new Map(
    usersResult.rows.map((user) => [String(user.name || "").trim().toLowerCase(), Number(user.id)])
  );

  const meetingsResult = await query(
    `
    SELECT id, participants, participant_user_ids
    FROM meetings
    WHERE COALESCE(participants, '') <> ''
    `
  );

  for (const meeting of meetingsResult.rows) {
    const currentIds = Array.isArray(meeting.participant_user_ids) ? meeting.participant_user_ids : [];
    if (currentIds.length) continue;

    const nextIds = String(meeting.participants || "")
      .split(",")
      .map((name) => name.trim().toLowerCase())
      .filter(Boolean)
      .map((name) => usersByName.get(name))
      .filter((id) => Number.isInteger(id));

    if (!nextIds.length) continue;

    await query(
      `
      UPDATE meetings
      SET participant_user_ids = $1::INTEGER[]
      WHERE id = $2
      `,
      [[...new Set(nextIds)], meeting.id]
    );
  }
}

module.exports = {
  hashPassword,
  initDb,
  pool,
  query,
  withTransaction
};
