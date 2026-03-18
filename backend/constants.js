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

module.exports = {
  MEETING_STATUS,
  MEETING_TYPES,
  MEETING_REASONS,
  USER_ROLES,
  USER_ROLE_OPTIONS
};
