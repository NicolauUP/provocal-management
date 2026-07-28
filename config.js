
// -----------------------------
// Folhas
// -----------------------------

const SHEETS = {
  GENERAL_ATTENDANCE: "Presencas_Gerais",
  SECTION_ATTENDANCE: [
    "Ensaios_Sopranos",
    "Ensaios_Contraltos",
    "Ensaios_Tenores",
    "Ensaios_Baixos"
  ],
  MEMBERS: "Membros",
  MOVEMENTS: "Movimentos",
  DISTRIBUTION: "Distribuição",
  SUMMARY: "Resumo"
};

// -----------------------------
// Intervalo de datas
// -----------------------------

const START_DATE = new Date(2026, 4, 1);
const END_DATE = new Date(2026, 5, 30);

const ACTIVITIES_START_DATE = new Date(2026, 4, 1); // 01/05/2026

// -----------------------------
// Publicação manual
// -----------------------------

const PUBLICATION = {
  SPREADSHEET_ID: "1QNh6oYrC72tSmD9Vy8aOg6CwFpj-AWVqv2qF_Lf3eAM",
  AUTHORIZED_PUBLISHERS: [
    "nicolau.23.sobrosa@gmail.com",
    "catarinagferreiralopes@gmail.com"
  ],
  INFO_SHEET: "Informação",
  DATA_SHEET: "_Dados_Consulta",
  PRIVATE_PORTAL_SPREADSHEET_ID:
    "1UwlcSqP5xDMIaSGmxwvRlu8R3jsp3U2PYHAl1-Nl4ME",
  PRIVATE_PORTAL_DATA_SHEET: "_Dados_Portal_Privado",
  PRIVATE_PORTAL_VALUES_SHEET: "_Valores_Portal_Privado",
  PRIVATE_PORTAL_GLOBAL_SHEET: "_Resumo_Coro_Privado"
};
