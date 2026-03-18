const path = require("path");
const XLSX = require("xlsx");
const { initDb, withTransaction } = require("./db");

const sectors = ["Consumo Masivo", "Retail", "Industria", "Servicios", "Energía", "Tecnología", "Salud"];
const managerNames = [
  "María Gómez",
  "Juan Pérez",
  "Lucía Fernández",
  "Nicolás Herrera",
  "Sofía Álvarez",
  "Mateo López",
  "Valentina Ruiz"
];

function parseMoney(value) {
  const clean = String(value || "")
    .replace(/\$/g, "")
    .replace(/,/g, "")
    .trim();
  return Number.parseFloat(clean) || 0;
}

function generateServices(rank) {
  const services = {
    fixedFire: rank % 2 === 0 || rank % 7 === 0,
    extinguishers: rank % 3 !== 0,
    works: rank % 4 === 0 || rank % 9 === 0
  };

  if (!services.fixedFire && !services.extinguishers && !services.works) {
    services.fixedFire = true;
  }

  return services;
}

function riskByPosition(position) {
  if (position % 7 === 0) return "Alto";
  if (position % 3 === 0) return "Medio";
  return "Bajo";
}

function segmentByPosition(position) {
  if (position <= 20) return "A";
  if (position <= 35) return "B";
  return "C";
}

async function seedFromExcel(excelPath) {
  await initDb();

  const workbook = XLSX.readFile(excelPath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  if (!rows.length) {
    throw new Error("El Excel no tiene filas de clientes");
  }

  await withTransaction(async (client) => {
    await client.query("TRUNCATE TABLE meetings, clients RESTART IDENTITY CASCADE");

    for (const [idx, row] of rows.entries()) {
      const position = Number(row.Posicion || idx + 1);
      const name = String(row.Cliente || `Compañía ${position}`).trim();
      const billing = parseMoney(row["Facturacion 2025"]);
      const services = generateServices(position);

      await client.query(
        `
        INSERT INTO clients (
          position, name, billing_2025, sector, manager, risk, segment,
          service_fixed_fire, service_extinguishers, service_works,
          wallet_share, nps, open_opportunities, notes
        ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, '')
        `,
        [
          position,
          name,
          billing,
          sectors[position % sectors.length],
          managerNames[position % managerNames.length],
          riskByPosition(position),
          segmentByPosition(position),
          services.fixedFire,
          services.extinguishers,
          services.works,
          `${Math.min(95, 35 + position)}%`,
          42 + (position % 30),
          (position % 8) + 1
        ]
      );
    }
  });

  return rows.length;
}

if (require.main === module) {
  const defaultPath = path.join(process.env.HOME || "", "Downloads", "top50.xlsx");
  const excelPath = process.argv[2] || defaultPath;
  seedFromExcel(excelPath)
    .then((total) => {
      console.log(`Seed completado con ${total} clientes desde ${excelPath}`);
    })
    .catch((error) => {
      console.error(error);
      process.exit(1);
    });
}

module.exports = {
  seedFromExcel
};
