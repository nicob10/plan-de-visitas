const MEETING_STATUS = {
  SCHEDULED: "Agendada",
  CONFIRMED: "Confirmada",
  COMPLETED: "Realizada"
};

const MEETING_TYPES = [
  { value: "Comercial", label: "Visita Comercial", color: "yellow" },
  { value: "Gerencia Comercial", label: "Visita Gerencia Comercial", color: "orange" },
  { value: "Directorio", label: "Visita Directorio", color: "blue" },
  { value: "Operaciones", label: "Visita Operaciones", color: "purple" }
];

const MEETING_REASONS = [
  "Acercamiento al cliente",
  "Negociacion",
  "Servicio"
];

const USER_ROLES = {
  EXECUTIVE: "Ejecutivo",
  DIRECTOR: "Director",
  SUPERVISOR_EXT: "Supervisor EXT",
  SUPERVISOR_IFCI: "Supervisor IFCI",
  SUPERVISOR_WORKS: "Supervisor Obra",
  MANAGER_IFCI: "Gerente IFCI",
  MANAGER_EXT: "Gerente EXT",
  MANAGER_WORKS: "Gerente Obra",
  COMMERCIAL_MANAGER: "Gerente Comercial"
};

const USER_ROLE_OPTIONS = Object.values(USER_ROLES);

const OPPORTUNITY_STATUS = {
  OPEN: "Abierta",
  QUALIFIED: "Calificada",
  PROPOSAL: "Propuesta",
  NEGOTIATION: "Negociación",
  WON: "Ganada",
  LOST: "Perdida"
};

const OPPORTUNITY_STATUS_OPTIONS = Object.values(OPPORTUNITY_STATUS);

const OPPORTUNITY_OPEN_STATUSES = [
  OPPORTUNITY_STATUS.OPEN,
  OPPORTUNITY_STATUS.QUALIFIED,
  OPPORTUNITY_STATUS.PROPOSAL,
  OPPORTUNITY_STATUS.NEGOTIATION
];

const OPPORTUNITY_SERVICE_LINES = [
  "Instalaciones Fijas",
  "Extintores",
  "Obras C.I.",
  "Multiservicio",
  "Otro"
];

const OPPORTUNITY_TYPES = [
  "Proyecto",
  "Negociación"
];

const FOLLOW_UP_STATUS = {
  PENDING: "Pendiente",
  DONE: "Hecho"
};

const FOLLOW_UP_STATUS_OPTIONS = Object.values(FOLLOW_UP_STATUS);

const FOLLOW_UP_TYPES = [
  "Llamada",
  "Email",
  "Visita",
  "Cotización",
  "Recordatorio",
  "Reunión"
];

module.exports = {
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
};
