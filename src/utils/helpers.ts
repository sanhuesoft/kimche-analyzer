import type { ObservationKind } from "@/types";

export const REQUIRED_HEADERS = [
  "Curso",
  "No. Lista",
  "Primer Apellido Estudiante",
  "Segundo Apellido Estudiante",
  "Nombre Estudiante",
  "Fecha",
  "Tipo de observación",
  "Descripción",
] as const;

export const PIE_COLORS: Record<string, string> = {
  "Anotaciones positivas": "#10b981",
  "Anotaciones negativas": "#ef4444",
  "Observaciones": "#f59e0b",
  "Citaciones a apoderado": "#8b5cf6",
  "Derivaciones": "#06b6d4",
  "Entrevistas con apoderado": "#6366f1",
  "Entrevistas con estudiante": "#ec4899",
};
export const DIACRITICS_REGEX = /[\u0300-\u036f]/g;

export function parseDateToSortable(value: string) {
  const clean = value.trim();
  if (!clean) return "";

  const isoMatch = clean.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    const [, y, m, d] = isoMatch;
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }

  const latamMatch = clean.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/);
  if (latamMatch) {
    const [, d, m, y] = latamMatch;
    const year = y.length === 2 ? `20${y}` : y;
    return `${year}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }

  return "";
}

export function formatDateToVerbal(dateStr: string): string {
  if (!dateStr) return "";
  const clean = dateStr.trim();
  const monthsNames = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
  ];

  const latamMatch = clean.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/);
  if (latamMatch) {
    const [, d, m] = latamMatch;
    const day = parseInt(d, 10);
    const monthNum = parseInt(m, 10);
    const monthName = monthsNames[monthNum - 1] || "";
    if (monthName) return `${day} de ${monthName}`;
  }

  const isoMatch = clean.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    const [, , m, d] = isoMatch;
    const day = parseInt(d, 10);
    const monthNum = parseInt(m, 10);
    const monthName = monthsNames[monthNum - 1] || "";
    if (monthName) return `${day} de ${monthName}`;
  }

  return dateStr;
}

export function normalizeType(typeValue: string): ObservationKind {
  const value = typeValue.toLowerCase().normalize("NFD").replace(DIACRITICS_REGEX, "");
  if (value.includes("anotacion positiva")) return "positiva";
  if (value.includes("anotacion negativa")) return "negativa";
  return "otros";
}

export function capitalizeProperName(name: string): string {
  if (!name) return "";
  return name
    .toLowerCase()
    .split(/\s+/)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}

export function buildStudentName(row: Record<string, string>) {
  const rawName = [
    row["Nombre Estudiante"],
    row["Primer Apellido Estudiante"],
    row["Segundo Apellido Estudiante"],
  ]
    .map((item) => item?.trim())
    .filter(Boolean)
    .join(" ");

  return capitalizeProperName(rawName);
}

export function typeBadgeClasses(type: ObservationKind) {
  if (type === "positiva") return "bg-emerald-100 text-emerald-700";
  if (type === "negativa") return "bg-rose-100 text-rose-700";
  return "bg-indigo-100 text-indigo-700";
}

export function typeLabel(type: ObservationKind) {
  if (type === "positiva") return "Positiva";
  if (type === "negativa") return "Negativa";
  return "Otros / Entrevistas";
}

export function calculateTrendLine(data: { value: number }[]) {
  const n = data.length;
  if (n === 0) return [];
  let sumX = 0;
  let sumY = 0;
  let sumXY = 0;
  let sumXX = 0;

  for (let i = 0; i < n; i++) {
    sumX += i;
    sumY += data[i].value;
    sumXY += i * data[i].value;
    sumXX += i * i;
  }

  const slope = n > 1 ? (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX) : 0;
  const intercept = (sumY - slope * sumX) / n;

  return data.map((d, i) => Math.max(0, parseFloat((slope * i + intercept).toFixed(2))));
}


