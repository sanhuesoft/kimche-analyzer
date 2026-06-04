"use client";

import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  AlertCircle,
  CheckCircle2,
  Filter,
  Search,
  Star,
  Upload,
  Users,
} from "lucide-react";
import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import { ChangeEvent, DragEvent, useEffect, useMemo, useRef, useState } from "react";

type ObservationKind = "positiva" | "negativa" | "otros";
type QuickFilter = "all" | ObservationKind;

type Observation = {
  id: string;
  curso: string;
  numeroLista: string;
  nombreCompleto: string;
  fechaTexto: string;
  fechaOrdenable: string;
  tipoOriginal: string;
  tipo: ObservationKind;
  descripcion: string;
  asignaturaOrCategorizacion: string;
  funcionario: string;
  falta: string;
};

type PendienteItem = {
  id: string;
  tipo: "firma" | "leccionario";
  docente: string;
  curso: string;
  asignatura: string;
  fecha: string;
  hora: string;
  estadoFirma?: string;
  estadoRegistro?: string;
  checked: boolean;
};

const REQUIRED_HEADERS = [
  "Curso",
  "No. Lista",
  "Primer Apellido Estudiante",
  "Segundo Apellido Estudiante",
  "Nombre Estudiante",
  "Fecha",
  "Tipo de observación",
  "Descripción",
] as const;

const PIE_COLORS = ["#10b981", "#ef4444", "#6366f1", "#f59e0b"];
const DIACRITICS_REGEX = /[\u0300-\u036f]/g;

function parseDateToSortable(value: string) {
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

function formatDateToVerbal(dateStr: string): string {
  if (!dateStr) return "";
  const clean = dateStr.trim();
  const monthsNames = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
  ];

  const latamMatch = clean.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/);
  if (latamMatch) {
    const [, d, m] = latamMatch;
    const day = parseInt(d, 10);
    const monthNum = parseInt(m, 10);
    const monthName = monthsNames[monthNum - 1] || "";
    if (monthName) {
      return `${day} de ${monthName}`;
    }
  }

  const isoMatch = clean.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    const [, , m, d] = isoMatch;
    const day = parseInt(d, 10);
    const monthNum = parseInt(m, 10);
    const monthName = monthsNames[monthNum - 1] || "";
    if (monthName) {
      return `${day} de ${monthName}`;
    }
  }

  return dateStr;
}

function normalizeType(typeValue: string): ObservationKind {
  const value = typeValue.toLowerCase().normalize("NFD").replace(DIACRITICS_REGEX, "");
  if (value.includes("anotacion positiva")) return "positiva";
  if (value.includes("anotacion negativa")) return "negativa";
  return "otros";
}

function capitalizeProperName(name: string): string {
  if (!name) return "";
  return name
    .toLowerCase()
    .split(/\s+/)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}

function buildStudentName(row: Record<string, string>) {
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

function typeBadgeClasses(type: ObservationKind) {
  if (type === "positiva") return "bg-emerald-100 text-emerald-700";
  if (type === "negativa") return "bg-rose-100 text-rose-700";
  return "bg-indigo-100 text-indigo-700";
}

function typeLabel(type: ObservationKind) {
  if (type === "positiva") return "Positiva";
  if (type === "negativa") return "Negativa";
  return "Otros / Entrevistas";
}

function calculateTrendLine(data: { value: number }[]) {
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

export default function Home() {
  const [observations, setObservations] = useState<Observation[]>([]);
  const [errorMessage, setErrorMessage] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [quickFilter, setQuickFilter] = useState<QuickFilter>("all");
  const [showPositiveTrend, setShowPositiveTrend] = useState(true);
  const [showNegativeTrend, setShowNegativeTrend] = useState(true);
  const [timeResolution, setTimeResolution] = useState<"daily" | "monthly">("daily");


  const [appMode, setAppMode] = useState<"observations" | "pendientes">("observations");
  const [pendientes, setPendientes] = useState<PendienteItem[]>([]);
  const [pendientesSearch, setPendientesSearch] = useState("");
  const [pendientesFilter, setPendientesFilter] = useState<"all" | "firma" | "leccionario">("all");
  const [selectedAsignatura, setSelectedAsignatura] = useState<string>("all");
  const [selectedCurso, setSelectedCurso] = useState<string>("all");
  const [selectedFecha, setSelectedFecha] = useState<string>("all");
  const [selectedAsignaturaObs, setSelectedAsignaturaObs] = useState<string>("all");
  const [selectedCursoObs, setSelectedCursoObs] = useState<string>("all");
  const [selectedFechaObs, setSelectedFechaObs] = useState<string>("all");
  const [isLoading, setIsLoading] = useState(false);



  const total = observations.length;

  const summary = useMemo(() => {
    const result = { positivas: 0, negativas: 0, otros: 0 };
    for (const row of observations) {
      if (row.tipo === "positiva") result.positivas += 1;
      if (row.tipo === "negativa") result.negativas += 1;
      if (row.tipo === "otros") result.otros += 1;
    }
    return result;
  }, [observations]);

  const positiveObservationsPercentage =
    (summary.positivas + summary.negativas) > 0
      ? Math.round((summary.positivas / (summary.positivas + summary.negativas)) * 100)
      : 0;

  const pieData = useMemo(() => {
    let positivas = 0;
    let negativas = 0;
    let entrevistas = 0;
    let otros = 0;

    for (const row of observations) {
      if (row.tipo === "positiva") {
        positivas += 1;
      } else if (row.tipo === "negativa") {
        negativas += 1;
      } else {
        const raw = row.tipoOriginal.toLowerCase().normalize("NFD").replace(DIACRITICS_REGEX, "");
        if (raw.includes("entrevista")) {
          entrevistas += 1;
        } else {
          otros += 1;
        }
      }
    }

    return [
      { name: "Positivas", value: positivas },
      { name: "Negativas", value: negativas },
      { name: "Entrevistas", value: entrevistas },
      { name: "Otros", value: otros },
    ].filter((item) => item.value > 0);
  }, [observations]);

  const trendData = useMemo(() => {
    const map = new Map<string, { fecha: string; positivas: number; negativas: number }>();

    for (const item of observations) {
      if (!item.fechaOrdenable) continue;
      if (!map.has(item.fechaOrdenable)) {
        map.set(item.fechaOrdenable, {
          fecha: item.fechaOrdenable,
          positivas: 0,
          negativas: 0,
        });
      }

      const current = map.get(item.fechaOrdenable);
      if (!current) continue;
      if (item.tipo === "positiva") current.positivas += 1;
      if (item.tipo === "negativa") current.negativas += 1;
    }

    const sorted = [...map.values()].sort((a, b) => a.fecha.localeCompare(b.fecha));

    const positiveValues = sorted.map((d) => ({ value: d.positivas }));
    const negativeValues = sorted.map((d) => ({ value: d.negativas }));

    const posTrend = calculateTrendLine(positiveValues);
    const negTrend = calculateTrendLine(negativeValues);

    return sorted.map((item, index) => ({
      ...item,
      tendenciaPositiva: posTrend[index] ?? 0,
      tendenciaNegativa: negTrend[index] ?? 0,
    }));
  }, [observations]);

  const monthlyData = useMemo(() => {
    const map = new Map<string, { mes: string; positivas: number; negativas: number }>();

    for (const item of observations) {
      if (!item.fechaOrdenable) continue;
      const mesKey = item.fechaOrdenable.substring(0, 7);
      if (!map.has(mesKey)) {
        map.set(mesKey, {
          mes: mesKey,
          positivas: 0,
          negativas: 0,
        });
      }

      const current = map.get(mesKey);
      if (current) {
        if (item.tipo === "positiva") current.positivas += 1;
        if (item.tipo === "negativa") current.negativas += 1;
      }
    }

    const sorted = [...map.values()].sort((a, b) => a.mes.localeCompare(b.mes));

    const positiveValues = sorted.map((d) => ({ value: d.positivas }));
    const negativeValues = sorted.map((d) => ({ value: d.negativas }));

    const posTrend = calculateTrendLine(positiveValues);
    const negTrend = calculateTrendLine(negativeValues);

    return sorted.map((item, index) => ({
      ...item,
      tendenciaPositiva: posTrend[index] ?? 0,
      tendenciaNegativa: negTrend[index] ?? 0,
    }));
  }, [observations]);

  const { topPositive, topNegative } = useMemo(() => {
    const positives = new Map<string, number>();
    const negatives = new Map<string, number>();

    for (const item of observations) {
      if (item.tipo === "positiva") {
        positives.set(item.nombreCompleto, (positives.get(item.nombreCompleto) ?? 0) + 1);
      }
      if (item.tipo === "negativa") {
        negatives.set(item.nombreCompleto, (negatives.get(item.nombreCompleto) ?? 0) + 1);
      }
    }

    const toTop = (source: Map<string, number>) =>
      [...source.entries()]
        .map(([estudiante, totalObservaciones]) => ({ estudiante, totalObservaciones }))
        .sort((a, b) => b.totalObservaciones - a.totalObservaciones)
        .slice(0, 5);

    return { topPositive: toTop(positives), topNegative: toTop(negatives) };
  }, [observations]);

  const topFuncionarios = useMemo(() => {
    const counts = new Map<string, { funcionario: string; positivas: number; negativas: number; total: number }>();

    for (const item of observations) {
      const func = item.funcionario || "No especificado";
      if (!counts.has(func)) {
        counts.set(func, { funcionario: func, positivas: 0, negativas: 0, total: 0 });
      }
      const data = counts.get(func)!;
      if (item.tipo === "positiva") {
        data.positivas += 1;
        data.total += 1;
      } else if (item.tipo === "negativa") {
        data.negativas += 1;
        data.total += 1;
      }
    }

    return [...counts.values()]
      .sort((a, b) => b.total - a.total)
      .slice(0, 5);
  }, [observations]);

  const faltasStats = useMemo(() => {
    const counts = new Map<string, number>();
    let totalValidos = 0;

    for (const item of observations) {
      if (!item.falta) continue;
      const f = item.falta.trim();
      if (!f) continue;
      counts.set(f, (counts.get(f) ?? 0) + 1);
      totalValidos += 1;
    }

    const data = [...counts.entries()]
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    return { data, totalValidos };
  }, [observations]);

  const { latestPositive, latestNegative } = useMemo(() => {
    let latestPos: Observation | null = null;
    let latestNeg: Observation | null = null;

    const sorted = [...observations]
      .filter((obs) => obs.fechaOrdenable)
      .sort((a, b) => b.fechaOrdenable.localeCompare(a.fechaOrdenable));

    for (const obs of sorted) {
      if (!latestPos && obs.tipo === "positiva") {
        latestPos = obs;
      }
      if (!latestNeg && obs.tipo === "negativa") {
        latestNeg = obs;
      }
      if (latestPos && latestNeg) break;
    }

    if (!latestPos) {
      latestPos = observations.find((obs) => obs.tipo === "positiva") ?? null;
    }
    if (!latestNeg) {
      latestNeg = observations.find((obs) => obs.tipo === "negativa") ?? null;
    }

    return { latestPositive: latestPos, latestNegative: latestNeg };
  }, [observations]);

  const filteredRows = useMemo(() => {
    const term = searchQuery.trim().toLowerCase();

    return observations.filter((item) => {
      const matchesSearch = !term || item.nombreCompleto.toLowerCase().includes(term);
      const matchesType = quickFilter === "all" || item.tipo === quickFilter;
      const matchesAsignatura = selectedAsignaturaObs === "all" || item.asignaturaOrCategorizacion === selectedAsignaturaObs;
      const matchesCurso = selectedCursoObs === "all" || item.curso === selectedCursoObs;
      const matchesFecha = selectedFechaObs === "all" || item.fechaTexto === selectedFechaObs;
      return matchesSearch && matchesType && matchesAsignatura && matchesCurso && matchesFecha;
    });
  }, [observations, quickFilter, searchQuery, selectedAsignaturaObs, selectedCursoObs, selectedFechaObs]);



  const uniqueAsignaturas = useMemo(() => {
    return Array.from(new Set(pendientes.map((p) => p.asignatura).filter(Boolean))).sort();
  }, [pendientes]);

  const uniqueCursos = useMemo(() => {
    return Array.from(new Set(pendientes.map((p) => p.curso).filter(Boolean))).sort();
  }, [pendientes]);

  const uniqueFechas = useMemo(() => {
    const rawFechas = Array.from(new Set(pendientes.map((p) => p.fecha).filter(Boolean)));
    return rawFechas.sort((a, b) => {
      const dateA = parseDateToSortable(a);
      const dateB = parseDateToSortable(b);
      return dateA.localeCompare(dateB);
    });
  }, [pendientes]);

  const uniqueAsignaturasObs = useMemo(() => {
    return Array.from(new Set(observations.map((o) => o.asignaturaOrCategorizacion).filter(Boolean))).sort();
  }, [observations]);

  const uniqueCursosObs = useMemo(() => {
    return Array.from(new Set(observations.map((o) => o.curso).filter(Boolean))).sort();
  }, [observations]);

  const uniqueFechasObs = useMemo(() => {
    const rawFechas = Array.from(new Set(observations.map((o) => o.fechaTexto).filter(Boolean)));
    return rawFechas.sort((a, b) => {
      const dateA = parseDateToSortable(a);
      const dateB = parseDateToSortable(b);
      return dateA.localeCompare(dateB);
    });
  }, [observations]);

  const sortedPendientes = useMemo(() => {
    const term = pendientesSearch.toLowerCase().trim();
    const filtered = pendientes.filter((item) => {
      const matchesType = pendientesFilter === "all" || item.tipo === pendientesFilter;
      const matchesAsignatura = selectedAsignatura === "all" || item.asignatura === selectedAsignatura;
      const matchesCurso = selectedCurso === "all" || item.curso === selectedCurso;
      const matchesFecha = selectedFecha === "all" || item.fecha === selectedFecha;
      const matchesSearch =
        !term ||
        item.docente.toLowerCase().includes(term) ||
        item.curso.toLowerCase().includes(term) ||
        item.asignatura.toLowerCase().includes(term) ||
        item.fecha.toLowerCase().includes(term) ||
        item.hora.toLowerCase().includes(term);
      return matchesType && matchesAsignatura && matchesCurso && matchesFecha && matchesSearch;
    });

    return [...filtered].sort((a, b) => {
      if (a.checked === b.checked) {
        return a.fecha.localeCompare(b.fecha) || a.hora.localeCompare(b.hora);
      }
      return a.checked ? 1 : -1;
    });
  }, [pendientes, pendientesSearch, pendientesFilter, selectedAsignatura, selectedCurso, selectedFecha]);

  const pendientesSummary = useMemo(() => {
    const totalItems = pendientes.length;
    const completados = pendientes.filter((p) => p.checked).length;
    const firmasPendientes = pendientes.filter((p) => p.tipo === "firma" && !p.checked).length;
    const leccionariosPendientes = pendientes.filter((p) => p.tipo === "leccionario" && !p.checked).length;
    return { totalItems, completados, firmasPendientes, leccionariosPendientes };
  }, [pendientes]);

  const togglePendiente = (id: string) => {
    setPendientes((prev) =>
      prev.map((item) => (item.id === id ? { ...item, checked: !item.checked } : item))
    );
  };

  const parseAndLoadCsv = (file: File) => {
    Papa.parse<Record<string, string>>(file, {
      header: true,
      skipEmptyLines: true,
      complete: ({ data, meta }) => {
        const missingHeaders = REQUIRED_HEADERS.filter((header) => !meta.fields?.includes(header));
        if (missingHeaders.length > 0) {
          setObservations([]);
          setErrorMessage("Kimche Analyzer no puede procesar ese tipo de planilla aún");
          return;
        }

        const parsed = data
          .map((row, index) => {
            const nombreCompleto = buildStudentName(row);
            const fechaTexto = row["Fecha"]?.trim() ?? "";
            const fechaOrdenable = parseDateToSortable(fechaTexto);

            const rawFunc =
              row["Nombre docente autor"]?.trim() ||
              row["Nombre Funcionario"]?.trim() ||
              row["Funcionario"]?.trim() ||
              row["Creado por"]?.trim() ||
              row["Profesor"]?.trim() ||
              row["Docente"]?.trim() ||
              "No especificado";
            const funcionario = rawFunc !== "No especificado" ? capitalizeProperName(rawFunc) : rawFunc;

            const falta =
              row["Falta"]?.trim() ||
              row["Detalle de la falta"]?.trim() ||
              row["Gravedad"]?.trim() ||
              row["Gravedad de la falta"]?.trim() ||
              row["Tipo de falta"]?.trim() ||
              "";

            return {
              id: `${fechaTexto}-${nombreCompleto}-${row["Tipo de observación"]?.trim() ?? ""}-${index}`,
              curso: row["Curso"]?.trim() ?? "",
              numeroLista: row["No. Lista"]?.trim() ?? "",
              nombreCompleto,
              fechaTexto,
              fechaOrdenable,
              tipoOriginal: row["Tipo de observación"]?.trim() ?? "",
              tipo: normalizeType(row["Tipo de observación"]?.trim() ?? ""),
              descripcion: row["Descripción"]?.trim() ?? "",
              asignaturaOrCategorizacion: row["Asignatura"]?.trim() || row["Categorización"]?.trim() || "",
              funcionario,
              falta,
            } satisfies Observation;
          })
          .filter((item) => item.nombreCompleto && item.descripcion && item.tipoOriginal);

        setObservations(parsed);
      },
      error: (error) => {
        setObservations([]);
        setErrorMessage(`No fue posible leer el archivo: ${error.message}`);
      },
    });
  };

  const processObservationsWorkbook = (workbook: XLSX.WorkBook) => {
    try {
      const firstSheetName = workbook.SheetNames[0];
      if (!firstSheetName) {
        setErrorMessage("El archivo Excel no contiene hojas de trabajo.");
        return;
      }
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json<Record<string, any>>(worksheet, {
        defval: "",
        raw: false,
      });

      if (jsonData.length === 0) {
        setErrorMessage("El archivo Excel está vacío.");
        return;
      }

      const headers = Object.keys(jsonData[0]);
      const missingHeaders = REQUIRED_HEADERS.filter((header) => !headers.includes(header));
      if (missingHeaders.length > 0) {
        setObservations([]);
        setErrorMessage("Kimche Analyzer no puede procesar ese tipo de planilla aún");
        return;
      }

      const parsed = jsonData
        .map((row, index) => {
          const normalizedRow: Record<string, string> = {};
          for (const key of Object.keys(row)) {
            normalizedRow[key] = String(row[key]);
          }

          const nombreCompleto = buildStudentName(normalizedRow);
          const fechaTexto = normalizedRow["Fecha"]?.trim() ?? "";
          const fechaOrdenable = parseDateToSortable(fechaTexto);

          const rawFunc =
            normalizedRow["Nombre docente autor"]?.trim() ||
            normalizedRow["Nombre Funcionario"]?.trim() ||
            normalizedRow["Funcionario"]?.trim() ||
            normalizedRow["Creado por"]?.trim() ||
            normalizedRow["Profesor"]?.trim() ||
            normalizedRow["Docente"]?.trim() ||
            "No especificado";
          const funcionario = rawFunc !== "No especificado" ? capitalizeProperName(rawFunc) : rawFunc;

          const falta =
            normalizedRow["Falta"]?.trim() ||
            normalizedRow["Detalle de la falta"]?.trim() ||
            normalizedRow["Gravedad"]?.trim() ||
            normalizedRow["Gravedad de la falta"]?.trim() ||
            normalizedRow["Tipo de falta"]?.trim() ||
            "";

          return {
            id: `${fechaTexto}-${nombreCompleto}-${normalizedRow["Tipo de observación"]?.trim() ?? ""}-${index}`,
            curso: normalizedRow["Curso"]?.trim() ?? "",
            numeroLista: normalizedRow["No. Lista"]?.trim() ?? "",
            nombreCompleto,
            fechaTexto,
            fechaOrdenable,
            tipoOriginal: normalizedRow["Tipo de observación"]?.trim() ?? "",
            tipo: normalizeType(normalizedRow["Tipo de observación"]?.trim() ?? ""),
            descripcion: normalizedRow["Descripción"]?.trim() ?? "",
            asignaturaOrCategorizacion: normalizedRow["Asignatura"]?.trim() || normalizedRow["Categorización"]?.trim() || "",
            funcionario,
            falta,
          } satisfies Observation;
        })
        .filter((item) => item.nombreCompleto && item.descripcion && item.tipoOriginal);

      setObservations(parsed);
    } catch (error: any) {
      setObservations([]);
      setErrorMessage(`No fue posible leer el archivo Excel: ${error.message}`);
    }
  };

  const processPendientesWorkbook = (workbook: XLSX.WorkBook) => {
    try {
      const items: PendienteItem[] = [];

      const sheet1Name = workbook.SheetNames.find((name) =>
        name.toLowerCase().includes("firma")
      ) || workbook.SheetNames[0];

      if (sheet1Name) {
        const sheet1 = workbook.Sheets[sheet1Name];
        const data1 = XLSX.utils.sheet_to_json<Record<string, any>>(sheet1, {
          defval: "",
          raw: false,
        });
        data1.forEach((row, index) => {
          const docente = row["Docente titular"]?.trim() ?? row["Docente"]?.trim() ?? "";
          const curso = row["Curso"]?.trim() ?? "";
          const asignatura = row["Asignatura"]?.trim() ?? "";
          const fecha = row["Fecha"]?.trim() ?? "";
          const hora = row["Hora de clase"]?.trim() ?? row["Hora"]?.trim() ?? "";

          if (docente || curso || asignatura) {
            items.push({
              id: `firma-${fecha}-${hora}-${asignatura}-${index}`,
              tipo: "firma",
              docente,
              curso,
              asignatura,
              fecha,
              hora,
              estadoFirma: row["Estado de la firma"]?.trim() ?? "",
              estadoRegistro: row["Estado del registro"]?.trim() ?? "",
              checked: false,
            });
          }
        });
      }

      const sheet2Name = workbook.SheetNames.find((name) =>
        name.toLowerCase().includes("leccionario") || name.toLowerCase().includes("registro")
      ) || workbook.SheetNames[1];

      if (sheet2Name) {
        const sheet2 = workbook.Sheets[sheet2Name];
        const data2 = XLSX.utils.sheet_to_json<Record<string, any>>(sheet2, {
          defval: "",
          raw: false,
        });
        data2.forEach((row, index) => {
          const docente = row["Docente titular"]?.trim() ?? row["Docente"]?.trim() ?? "";
          const curso = row["Curso"]?.trim() ?? "";
          const asignatura = row["Asignatura"]?.trim() ?? "";
          const fecha = row["Fecha"]?.trim() ?? "";
          const hora = row["Hora de clase"]?.trim() ?? row["Hora"]?.trim() ?? "";

          if (docente || curso || asignatura) {
            items.push({
              id: `leccionario-${fecha}-${hora}-${asignatura}-${index}`,
              tipo: "leccionario",
              docente,
              curso,
              asignatura,
              fecha,
              hora,
              estadoRegistro: row["Estado del registro"]?.trim() ?? "",
              checked: false,
            });
          }
        });
      }

      if (items.length === 0) {
        setErrorMessage("No se encontraron registros pendientes en el archivo.");
        return;
      }

      setPendientes(items);
    } catch (error: any) {
      setErrorMessage(`Error al leer el archivo de pendientes: ${error.message}`);
    }
  };

  const handleFile = (file: File) => {
    setErrorMessage("");
    const fileName = file.name.toLowerCase();

    if (fileName.endsWith(".csv")) {
      setAppMode("observations");
      parseAndLoadCsv(file);
    } else if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: "array" });

          const firstSheetName = workbook.SheetNames[0];
          let isPendientes = false;

          const hasPendientesSheetNames = workbook.SheetNames.some(
            (name) =>
              name.toLowerCase().includes("firma") ||
              name.toLowerCase().includes("leccionario") ||
              name.toLowerCase().includes("registro")
          );

          if (hasPendientesSheetNames) {
            isPendientes = true;
          } else if (firstSheetName) {
            const worksheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json<Record<string, any>>(worksheet, { header: 1 });
            const headers = (json[0] || []).map((h: any) => String(h).trim().toLowerCase());
            isPendientes = headers.includes("docente titular") || headers.includes("docente") || headers.includes("estado de la firma") || headers.includes("hora de clase");
          }

          if (isPendientes) {
            setAppMode("pendientes");
            processPendientesWorkbook(workbook);
          } else {
            setAppMode("observations");
            processObservationsWorkbook(workbook);
          }
        } catch (error: any) {
          setErrorMessage(`No fue posible leer el archivo Excel: ${error.message}`);
        }
      };
      reader.onerror = () => {
        setErrorMessage("Error al leer el archivo.");
      };
      reader.readAsArrayBuffer(file);
    } else {
      setErrorMessage("Por favor, sube un archivo CSV o Excel (.xlsx, .xls).");
    }
  };

  useEffect(() => {
    const handleExtensionMessage = (event: MessageEvent) => {
      const message = event.data;
      if (message && message.type === "KIMCHE_EXT_FILE") {
        setIsLoading(true);
        // Delay slightly so the loading animation renders beautifully before processing
        setTimeout(() => {
          try {
            const fileName = message.name || "archivo.xlsx";

            if (message.base64) {
              const binaryString = window.atob(message.base64);
              const len = binaryString.length;
              const bytes = new Uint8Array(len);
              for (let i = 0; i < len; i++) {
                bytes[i] = binaryString.charCodeAt(i);
              }

              const file = new File([bytes], fileName, {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              });
              handleFile(file);
            } else if (message.data instanceof Uint8Array || message.data instanceof ArrayBuffer) {
              const bytes = message.data instanceof Uint8Array ? message.data : new Uint8Array(message.data);
              const file = new File([bytes], fileName, {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              });
              handleFile(file);
            }
          } catch (err: any) {
            setErrorMessage(`Error al procesar archivo de la extensión: ${err.message}`);
          } finally {
            setIsLoading(false);
          }
        }, 1200);
      }
    };

    window.addEventListener("message", handleExtensionMessage);
    return () => {
      window.removeEventListener("message", handleExtensionMessage);
    };
  }, [handleFile]);

  const handleInputChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    handleFile(file);
    event.target.value = "";
  };

  const handleDrop = (event: DragEvent<HTMLLabelElement>) => {
    event.preventDefault();
    setIsDragging(false);
    const file = event.dataTransfer.files?.[0];
    if (!file) return;
    handleFile(file);
  };

  if (observations.length === 0 && pendientes.length === 0) {
    return (
      <main className="flex-grow w-full bg-slate-100 px-4 py-10 text-slate-900 flex items-center justify-center">
        {isLoading && (
          <div className="fixed inset-0 z-[100] flex flex-col items-center justify-center bg-slate-950/70 backdrop-blur-md select-none transition-all duration-300">
            <div className="relative flex items-center justify-center h-20 w-20">
              {/* Spinning outer gradient ring */}
              <div className="absolute inset-0 rounded-full border-4 border-transparent border-t-indigo-500 border-r-purple-500 animate-spin" />
              {/* Secondary counter-spinning ring */}
              <div className="absolute inset-1.5 rounded-full border-4 border-transparent border-b-blue-500 border-l-pink-500 animate-spin [animation-duration:1.5s]" />
              {/* Inner glowing core */}
              <div className="h-4 w-4 rounded-full bg-indigo-500 animate-ping" />
            </div>
            <p className="mt-6 text-[11px] font-extrabold uppercase tracking-widest text-slate-200 animate-pulse">
              Cargando planilla automáticamente...
            </p>
          </div>
        )}
        <section className="w-full max-w-3xl rounded-2xl border border-slate-200 bg-white p-8 shadow-sm">
          <div className="mb-8 flex items-center gap-3">
            <div className="rounded-xl bg-indigo-100 p-2 text-indigo-600">
              <Users className="h-5 w-5" />
            </div>
            <div>
              <h1 className="text-2xl font-semibold">Kimche Analyzer</h1>
              <p className="text-sm text-slate-500">
                Carga tu Registro de firmas pendientes o el Registro de observaciones de tu curso (detección automática) y obtén estadísticas. El análisis se realiza de manera local, por lo que no se comparte ningún dato con nadie.
              </p>
            </div>
          </div>

          <label
            onDragOver={(event) => {
              event.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
            className={`flex cursor-pointer flex-col items-center justify-center gap-3 rounded-2xl border-2 border-dashed p-12 text-center transition ${isDragging
              ? "border-indigo-500 bg-indigo-50"
              : "border-slate-300 bg-slate-50 hover:border-indigo-400"
              }`}
          >
            <div className="rounded-full bg-white p-3 shadow-sm">
              <Upload className="h-6 w-6 text-indigo-500" />
            </div>
            <div>
              <p className="font-medium">Arrastra y suelta tu archivo aquí</p>
              <p className="text-sm text-slate-500">Acepta CSV de observaciones y Excel (.xlsx, .xls) de pendientes</p>
            </div>
            <input
              type="file"
              accept=".csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              className="hidden"
              onChange={handleInputChange}
            />
          </label>

          {errorMessage ? (
            <p className="mt-5 flex items-center gap-2 text-sm text-rose-600">
              <AlertCircle className="h-4 w-4" />
              {errorMessage}
            </p>
          ) : null}
        </section>
      </main>
    );
  }

  if (appMode === "pendientes" && pendientes.length > 0) {
    return (
      <main className="flex-grow w-full bg-slate-100 px-4 py-8 text-slate-900">
        {isLoading && (
          <div className="fixed inset-0 z-[100] flex flex-col items-center justify-center bg-slate-950/70 backdrop-blur-md select-none transition-all duration-300">
            <div className="relative flex items-center justify-center h-20 w-20">
              {/* Spinning outer gradient ring */}
              <div className="absolute inset-0 rounded-full border-4 border-transparent border-t-indigo-500 border-r-purple-500 animate-spin" />
              {/* Secondary counter-spinning ring */}
              <div className="absolute inset-1.5 rounded-full border-4 border-transparent border-b-blue-500 border-l-pink-500 animate-spin [animation-duration:1.5s]" />
              {/* Inner glowing core */}
              <div className="h-4 w-4 rounded-full bg-indigo-500 animate-ping" />
            </div>
            <p className="mt-6 text-[11px] font-extrabold uppercase tracking-widest text-slate-200 animate-pulse">
              Cargando planilla automáticamente...
            </p>
          </div>
        )}
        <div className="mx-auto flex max-w-7xl flex-col gap-6">
          <header className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
            <div>
              <h1 className="text-2xl font-semibold">Control de Firmas y Leccionarios Pendientes</h1>
              <p className="mt-1 text-sm text-slate-500">
                Visualiza y gestiona las actividades pendientes en tu libro de clases digital. Todo de forma local y temporal.
              </p>
            </div>
            <button
              onClick={() => {
                setObservations([]);
                setPendientes([]);
                setErrorMessage("");
                setSelectedAsignatura("all");
                setSelectedCurso("all");
                setSelectedFecha("all");
                setPendientesSearch("");
                setPendientesFilter("all");
                setSelectedAsignaturaObs("all");
                setSelectedCursoObs("all");
                setSelectedFechaObs("all");
                setSearchQuery("");
                setQuickFilter("all");
              }}
              className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2.5 text-sm font-semibold text-slate-700 hover:bg-slate-100 transition active:scale-95 shadow-sm whitespace-nowrap"
            >
              Subir otro archivo
            </button>
          </header>

          <section className="grid gap-4 md:grid-cols-2 xl:grid-cols-3">
            {/* Firmas Pendientes */}
            <article className="relative overflow-hidden rounded-xl bg-blue-600 px-4 py-3.5 text-white shadow-sm transition hover:shadow-md">
              <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
              <div className="relative z-10 flex items-center justify-between">
                <div className="flex items-center gap-2.5">
                  <div className="rounded-lg bg-white/10 p-1.5 text-white">
                    <Users className="h-4 w-4" />
                  </div>
                  <p className="text-xs font-semibold uppercase tracking-wider text-blue-100">Firmas pendientes</p>
                </div>
                <p className="text-2xl font-extrabold">{pendientesSummary.firmasPendientes}</p>
              </div>
            </article>

            {/* Leccionarios Pendientes */}
            <article className="relative overflow-hidden rounded-xl bg-purple-600 px-4 py-3.5 text-white shadow-sm transition hover:shadow-md">
              <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
              <div className="relative z-10 flex items-center justify-between">
                <div className="flex items-center gap-2.5">
                  <div className="rounded-lg bg-white/10 p-1.5 text-white">
                    <AlertCircle className="h-4 w-4" />
                  </div>
                  <p className="text-xs font-semibold uppercase tracking-wider text-purple-100">Leccionarios pendientes</p>
                </div>
                <p className="text-2xl font-extrabold">{pendientesSummary.leccionariosPendientes}</p>
              </div>
            </article>

            {/* Tareas Completadas */}
            <article className="relative overflow-hidden rounded-xl bg-emerald-600 px-4 py-3.5 text-white shadow-sm transition hover:shadow-md">
              <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
              <div className="relative z-10 flex items-center justify-between">
                <div className="flex items-center gap-2.5">
                  <div className="rounded-lg bg-white/10 p-1.5 text-white">
                    <CheckCircle2 className="h-4 w-4" />
                  </div>
                  <p className="text-xs font-semibold uppercase tracking-wider text-emerald-100">Tareas completadas</p>
                </div>
                <p className="text-2xl font-extrabold">{pendientesSummary.completados}</p>
              </div>
            </article>
          </section>

          <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <div className="mb-6 flex flex-col gap-4 border-b border-slate-100 pb-5">
              <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                <h3 className="font-semibold text-slate-800 flex items-center gap-2">
                  <Filter className="h-4 w-4 text-indigo-500" />
                  Panel de Filtros Rápidos
                </h3>
                {(selectedAsignatura !== "all" || selectedCurso !== "all" || selectedFecha !== "all" || pendientesSearch !== "" || pendientesFilter !== "all") && (
                  <button
                    onClick={() => {
                      setSelectedAsignatura("all");
                      setSelectedCurso("all");
                      setSelectedFecha("all");
                      setPendientesSearch("");
                      setPendientesFilter("all");
                    }}
                    className="text-xs text-indigo-600 hover:text-indigo-800 font-semibold underline cursor-pointer"
                  >
                    Restablecer todos los filtros
                  </button>
                )}
              </div>

              <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-5">
                {/* 1. Búsqueda de texto */}
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                    Buscar
                  </label>
                  <div className="relative">
                    <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                    <input
                      value={pendientesSearch}
                      onChange={(event) => setPendientesSearch(event.target.value)}
                      placeholder="Docente o palabra..."
                      className="w-full rounded-xl border border-slate-300 bg-white py-2 pl-9 pr-3 text-xs outline-none ring-indigo-500 transition focus:ring"
                    />
                  </div>
                </div>

                {/* 2. Filtro de Tipo */}
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                    Tipo
                  </label>
                  <select
                    value={pendientesFilter}
                    onChange={(e) => setPendientesFilter(e.target.value as any)}
                    className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-xs outline-none ring-indigo-500 focus:ring cursor-pointer"
                  >
                    <option value="all">Todos</option>
                    <option value="firma">Solo Firmas</option>
                    <option value="leccionario">Solo Leccionarios</option>
                  </select>
                </div>

                {/* 3. Filtro de Asignatura */}
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                    Asignatura
                  </label>
                  <select
                    value={selectedAsignatura}
                    onChange={(e) => setSelectedAsignatura(e.target.value)}
                    className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-xs outline-none ring-indigo-500 focus:ring cursor-pointer"
                  >
                    <option value="all">Todas</option>
                    {uniqueAsignaturas.map((asig) => (
                      <option key={asig} value={asig}>
                        {asig}
                      </option>
                    ))}
                  </select>
                </div>

                {/* 4. Filtro de Curso */}
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                    Curso
                  </label>
                  <select
                    value={selectedCurso}
                    onChange={(e) => setSelectedCurso(e.target.value)}
                    className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-xs outline-none ring-indigo-500 focus:ring cursor-pointer"
                  >
                    <option value="all">Todos</option>
                    {uniqueCursos.map((c) => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>

                {/* 5. Filtro de Fecha */}
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                    Fecha
                  </label>
                  <select
                    value={selectedFecha}
                    onChange={(e) => setSelectedFecha(e.target.value)}
                    className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-xs outline-none ring-indigo-500 focus:ring cursor-pointer"
                  >
                    <option value="all">Todas</option>
                    {uniqueFechas.map((f) => (
                      <option key={f} value={f}>
                        {formatDateToVerbal(f)}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {sortedPendientes.length > 0 ? (
                sortedPendientes.map((item) => (
                  <div
                    key={item.id}
                    onClick={() => togglePendiente(item.id)}
                    className={`flex items-start gap-3 rounded-xl border p-4 cursor-pointer transition select-none ${item.checked
                      ? "bg-slate-50 border-slate-200 text-slate-400 opacity-60 line-through"
                      : "bg-white border-slate-200 hover:border-indigo-300 hover:bg-indigo-50/10 shadow-sm hover:shadow"
                      }`}
                  >
                    <div className="flex items-start pt-1">
                      <input
                        type="checkbox"
                        checked={item.checked}
                        onChange={() => { }}
                        className="h-5 w-5 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                      />
                    </div>

                    <div className="flex-1 flex flex-col justify-between min-h-[90px] gap-2">
                      <div>
                        <div className="mb-1.5">
                          <span
                            className={`inline-block rounded-md px-2 py-0.5 text-[9px] font-bold uppercase tracking-wider text-white ${item.tipo === "firma"
                              ? "bg-blue-600"
                              : "bg-purple-600"
                              }`}
                          >
                            {item.tipo === "firma" ? "Firma" : "Leccionario"}
                          </span>
                        </div>
                        <h4 className="font-bold text-base text-slate-800 leading-tight">
                          {item.asignatura}
                        </h4>
                      </div>

                      {/* Unified 3-column metadata block */}
                      <div className={`grid grid-cols-3 text-center text-[11px] bg-slate-50 border border-slate-200/60 rounded-lg divide-x divide-slate-200 overflow-hidden mt-1.5 shadow-sm ${item.checked ? "opacity-75" : ""}`}>
                        <div className="py-1.5 px-1 font-semibold text-slate-700 truncate" title={item.curso}>
                          {item.curso}
                        </div>
                        <div className="py-1.5 px-1 font-medium text-slate-600 truncate" title={item.fecha}>
                          {item.fecha}
                        </div>
                        <div className={`py-1.5 px-1 font-semibold truncate ${item.checked ? "text-slate-500 bg-slate-105" : "text-indigo-700 bg-indigo-50/20"}`} title={item.hora ? `Bloque ${item.hora}` : "Sin bloque"}>
                          {item.hora ? `Bloque ${item.hora}` : "-"}
                        </div>
                      </div>

                      <p className="text-[10px] text-slate-400 font-light truncate">
                        Docente: {item.docente}
                      </p>
                    </div>
                  </div>
                ))
              ) : (
                <div className="col-span-full text-center py-12 text-slate-400 italic">
                  No se encontraron pendientes.
                </div>
              )}
            </div>
          </section>
        </div>
      </main>
    );
  }

  return (
    <main className="flex-grow w-full bg-slate-100 px-4 py-8 text-slate-900">
      {isLoading && (
        <div className="fixed inset-0 z-[100] flex flex-col items-center justify-center bg-slate-950/70 backdrop-blur-md select-none transition-all duration-300">
          <div className="relative flex items-center justify-center h-20 w-20">
            {/* Spinning outer gradient ring */}
            <div className="absolute inset-0 rounded-full border-4 border-transparent border-t-indigo-500 border-r-purple-500 animate-spin" />
            {/* Secondary counter-spinning ring */}
            <div className="absolute inset-1.5 rounded-full border-4 border-transparent border-b-blue-500 border-l-pink-500 animate-spin [animation-duration:1.5s]" />
            {/* Inner glowing core */}
            <div className="h-4 w-4 rounded-full bg-indigo-500 animate-ping" />
          </div>
          <p className="mt-6 text-[11px] font-extrabold uppercase tracking-widest text-slate-200 animate-pulse">
            Cargando planilla automáticamente...
          </p>
        </div>
      )}
      <div className="mx-auto flex max-w-7xl flex-col gap-6">
        <header className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
          <div>
            <h1 className="text-2xl font-semibold">Dashboard de Convivencia Escolar</h1>
            <p className="mt-1 text-sm text-slate-500">
              Visualización y métricas de observaciones de convivencia escolar cargadas.
            </p>
          </div>
          <button
            onClick={() => {
              setObservations([]);
              setPendientes([]);
              setErrorMessage("");
              setSelectedAsignatura("all");
              setSelectedCurso("all");
              setSelectedFecha("all");
              setPendientesSearch("");
              setPendientesFilter("all");
              setSelectedAsignaturaObs("all");
              setSelectedCursoObs("all");
              setSelectedFechaObs("all");
              setSearchQuery("");
              setQuickFilter("all");
            }}
            className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2.5 text-sm font-semibold text-slate-700 hover:bg-slate-100 transition active:scale-95 shadow-sm whitespace-nowrap"
          >
            Subir otro archivo
          </button>
        </header>

        <section className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
          {/* Observaciones procesadas */}
          <article className="relative overflow-hidden rounded-xl bg-blue-600 px-4 py-3.5 text-white shadow-sm transition hover:shadow-md">
            <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
            <div className="relative z-10 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="rounded-lg bg-white/10 p-1.5 text-white">
                  <Users className="h-4 w-4" />
                </div>
                <p className="text-xs font-semibold uppercase tracking-wider text-blue-100">Observaciones procesadas</p>
              </div>
              <p className="text-2xl font-extrabold">{total}</p>
            </div>
          </article>

          {/* Anotaciones positivas */}
          <article className="relative overflow-hidden rounded-xl bg-emerald-600 px-4 py-3.5 text-white shadow-sm transition hover:shadow-md">
            <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
            <div className="relative z-10 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="rounded-lg bg-white/10 p-1.5 text-white">
                  <CheckCircle2 className="h-4 w-4" />
                </div>
                <p className="text-xs font-semibold uppercase tracking-wider text-emerald-100">Anotaciones positivas</p>
              </div>
              <p className="text-2xl font-extrabold">{summary.positivas}</p>
            </div>
          </article>

          {/* Anotaciones negativas */}
          <article className="relative overflow-hidden rounded-xl bg-rose-600 px-4 py-3.5 text-white shadow-sm transition hover:shadow-md">
            <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
            <div className="relative z-10 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="rounded-lg bg-white/10 p-1.5 text-white">
                  <AlertCircle className="h-4 w-4" />
                </div>
                <p className="text-xs font-semibold uppercase tracking-wider text-rose-100">Anotaciones negativas</p>
              </div>
              <p className="text-2xl font-extrabold">{summary.negativas}</p>
            </div>
          </article>

          {/* Ratio de convivencia */}
          <article className="group cursor-help relative overflow-visible rounded-xl bg-indigo-600 px-4 py-3.5 text-white shadow-sm transition hover:shadow-md">
            <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none rounded-xl" />
            <div className="relative z-10 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="rounded-lg bg-white/10 p-1.5 text-white">
                  <Star className="h-4 w-4" />
                </div>
                <p className="text-xs font-semibold uppercase tracking-wider text-indigo-100">Ratio de convivencia</p>
              </div>
              <p className="text-2xl font-extrabold">{positiveObservationsPercentage}%</p>
            </div>

            {/* Hover Explanatory Popup Tooltip */}
            <div className="pointer-events-none absolute bottom-full left-1/2 z-50 mb-3 w-64 -translate-x-1/2 rounded-xl bg-slate-950 p-3 text-[11px] text-slate-100 shadow-xl opacity-0 transition-all duration-200 scale-95 origin-bottom group-hover:opacity-100 group-hover:scale-100 select-none">
              <p className="font-bold text-white mb-1">¿Cómo se calcula este ratio?</p>
              <p className="text-slate-350 leading-relaxed">
                Representa el porcentaje de anotaciones positivas sobre el total combinado de anotaciones positivas y negativas (excluyendo entrevistas u otros registros).
              </p>
              <div className="absolute top-full left-1/2 h-2 w-2 -translate-x-1/2 -translate-y-1 rotate-45 bg-slate-950" />
            </div>
          </article>
        </section>

        <section className="grid gap-4 md:grid-cols-2">
          <article className="rounded-2xl border-l-4 border-l-emerald-500 border-y border-r border-slate-200 bg-white p-5 shadow-sm flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-semibold uppercase tracking-wider text-emerald-600 bg-emerald-50 px-2.5 py-1 rounded-full">
                  Última Anotación Positiva
                </span>
                {latestPositive && (
                  <span className="text-xs text-slate-500">
                    {latestPositive.fechaTexto}
                  </span>
                )}
              </div>
              {latestPositive ? (
                <div>
                  <h3 className="text-lg font-semibold text-slate-800">
                    {latestPositive.nombreCompleto}
                  </h3>

                  <p className="text-sm text-slate-600 italic line-clamp-3">
                    "{latestPositive.descripcion}"
                  </p>
                </div>
              ) : (
                <p className="text-sm text-slate-400 py-4 italic text-center">
                  No se encontraron anotaciones positivas.
                </p>
              )}
            </div>
          </article>

          <article className="rounded-2xl border-l-4 border-l-rose-500 border-y border-r border-slate-200 bg-white p-5 shadow-sm flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-semibold uppercase tracking-wider text-rose-600 bg-rose-50 px-2.5 py-1 rounded-full">
                  Última Anotación Negativa
                </span>
                {latestNegative && (
                  <span className="text-xs text-slate-500">
                    {latestNegative.fechaTexto}
                  </span>
                )}
              </div>
              {latestNegative ? (
                <div>
                  <h3 className="text-lg font-semibold text-slate-800">
                    {latestNegative.nombreCompleto}
                  </h3>

                  <p className="text-sm text-slate-600 italic line-clamp-3">
                    "{latestNegative.descripcion}"
                  </p>
                </div>
              ) : (
                <p className="text-sm text-slate-400 py-4 italic text-center">
                  No se encontraron anotaciones negativas.
                </p>
              )}
            </div>
          </article>
        </section>

        <section className="grid gap-4 xl:grid-cols-3">
          <article className="h-80 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
            <h2 className="mb-3 font-semibold">Distribución de observaciones</h2>
            <ResponsiveContainer width="100%" height="90%">
              <PieChart>
                <Pie data={pieData} dataKey="value" nameKey="name" outerRadius={90} innerRadius={45}>
                  {pieData.map((entry, index) => (
                    <Cell key={entry.name} fill={PIE_COLORS[index % PIE_COLORS.length]} />
                  ))}
                </Pie>
                <Legend />
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          </article>

          <article className="h-80 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm xl:col-span-2">
            <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
              <h2 className="font-semibold">Evolución temporal (positivas vs negativas)</h2>
              <div className="flex flex-wrap items-center gap-3">
                {/* Selector de resolución temporal */}
                <div className="flex rounded-lg bg-slate-100 p-0.5 border border-slate-200">
                  <button
                    type="button"
                    onClick={() => setTimeResolution("daily")}
                    className={`rounded-md px-2.5 py-1 text-xs font-semibold transition ${timeResolution === "daily"
                      ? "bg-white text-slate-800 shadow-sm"
                      : "text-slate-500 hover:text-slate-850"
                      }`}
                  >
                    Diario (Líneas)
                  </button>
                  <button
                    type="button"
                    onClick={() => setTimeResolution("monthly")}
                    className={`rounded-md px-2.5 py-1 text-xs font-semibold transition ${timeResolution === "monthly"
                      ? "bg-white text-slate-800 shadow-sm"
                      : "text-slate-500 hover:text-slate-850"
                      }`}
                  >
                    Mensual (Líneas)
                  </button>
                </div>

                {/* Filtros de la leyenda */}
                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={() => setShowPositiveTrend(!showPositiveTrend)}
                    className={`flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold border transition ${showPositiveTrend
                      ? "bg-emerald-100 text-emerald-800 border-emerald-300"
                      : "bg-slate-50 text-slate-400 border-slate-200 line-through"
                      }`}
                  >
                    <span className={`h-2 w-2 rounded-full ${showPositiveTrend ? "bg-emerald-500" : "bg-slate-400"}`} />
                    Positivas
                  </button>
                  <button
                    type="button"
                    onClick={() => setShowNegativeTrend(!showNegativeTrend)}
                    className={`flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold border transition ${showNegativeTrend
                      ? "bg-rose-100 text-rose-800 border-rose-300"
                      : "bg-slate-50 text-slate-400 border-slate-200 line-through"
                      }`}
                  >
                    <span className={`h-2 w-2 rounded-full ${showNegativeTrend ? "bg-rose-500" : "bg-slate-400"}`} />
                    Negativas
                  </button>
                </div>
              </div>
            </div>
            <ResponsiveContainer width="100%" height="80%">
              {timeResolution === "daily" ? (
                <LineChart data={trendData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                  <XAxis dataKey="fecha" />
                  <YAxis allowDecimals={false} />
                  <Tooltip />
                  {showPositiveTrend && (
                    <Line type="monotone" dataKey="positivas" stroke="#10b981" strokeWidth={2} dot={{ r: 4, fill: "#10b981", stroke: "#10b981" }} name="Positivas" />
                  )}
                  {showPositiveTrend && (
                    <Line type="monotone" dataKey="tendenciaPositiva" stroke="#059669" strokeDasharray="4 4" dot={false} activeDot={false} name="Tendencia Positivas" />
                  )}
                  {showNegativeTrend && (
                    <Line type="monotone" dataKey="negativas" stroke="#ef4444" strokeWidth={2} dot={{ r: 4, fill: "#ef4444", stroke: "#ef4444" }} name="Negativas" />
                  )}
                  {showNegativeTrend && (
                    <Line type="monotone" dataKey="tendenciaNegativa" stroke="#dc2626" strokeDasharray="4 4" dot={false} activeDot={false} name="Tendencia Negativas" />
                  )}
                </LineChart>
              ) : (
                <LineChart data={monthlyData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                  <XAxis dataKey="mes" />
                  <YAxis allowDecimals={false} />
                  <Tooltip />
                  {showPositiveTrend && (
                    <Line type="monotone" dataKey="positivas" stroke="#10b981" strokeWidth={2} dot={{ r: 4, fill: "#10b981", stroke: "#10b981" }} name="Positivas" />
                  )}
                  {showPositiveTrend && (
                    <Line type="monotone" dataKey="tendenciaPositiva" stroke="#059669" strokeDasharray="4 4" dot={false} activeDot={false} name="Tendencia Positivas" />
                  )}
                  {showNegativeTrend && (
                    <Line type="monotone" dataKey="negativas" stroke="#ef4444" strokeWidth={2} dot={{ r: 4, fill: "#ef4444", stroke: "#ef4444" }} name="Negativas" />
                  )}
                  {showNegativeTrend && (
                    <Line type="monotone" dataKey="tendenciaNegativa" stroke="#dc2626" strokeDasharray="4 4" dot={false} activeDot={false} name="Tendencia Negativas" />
                  )}
                </LineChart>
              )}
            </ResponsiveContainer>
          </article>
        </section>

        <section className="grid gap-4 lg:grid-cols-2">
          {/* Tarjeta 1: Estudiantes con más positivas */}
          <article className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="mb-4 text-lg font-semibold text-slate-800 flex items-center gap-2">
              <CheckCircle2 className="h-5 w-5 text-emerald-500" />
              Top 5 estudiantes con más positivas
            </h2>
            <div className="overflow-x-auto">
              <table className="min-w-full text-left text-sm">
                <thead>
                  <tr className="border-b border-slate-200 text-slate-500 font-medium">
                    <th className="px-4 py-3 text-center w-16">Puesto</th>
                    <th className="px-4 py-3">Estudiante</th>
                    <th className="px-4 py-3 text-center w-28">Anotaciones</th>
                  </tr>
                </thead>
                <tbody>
                  {Array.from({ length: 5 }).map((_, index) => {
                    const pos = topPositive[index];
                    const rank = index + 1;

                    let rankClass = "text-slate-500 bg-slate-50";
                    if (rank === 1) rankClass = "bg-amber-100 text-amber-800 font-bold ring-1 ring-amber-200";
                    else if (rank === 2) rankClass = "bg-slate-100 text-slate-700 font-bold ring-1 ring-slate-200";
                    else if (rank === 3) rankClass = "bg-orange-50 text-orange-700 font-bold ring-1 ring-orange-100";

                    return (
                      <tr key={index} className="border-b border-slate-100 hover:bg-slate-50/50 transition">
                        <td className="px-4 py-3.5 text-center">
                          <span className={`inline-flex items-center justify-center h-6 w-6 rounded-full text-xs ${rankClass}`}>
                            {rank}
                          </span>
                        </td>
                        <td className="px-4 py-3.5 font-medium text-slate-800">
                          {pos ? pos.estudiante : <span className="text-slate-400 font-normal">-</span>}
                        </td>
                        <td className="px-4 py-3.5 text-center">
                          {pos ? (
                            <span className="inline-block rounded-full bg-emerald-100 px-3 py-1 text-xs font-bold text-emerald-700">
                              {pos.totalObservaciones}
                            </span>
                          ) : (
                            <span className="text-slate-400">-</span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </article>

          {/* Tarjeta 2: Estudiantes con más negativas */}
          <article className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="mb-4 text-lg font-semibold text-slate-800 flex items-center gap-2">
              <AlertCircle className="h-5 w-5 text-rose-500" />
              Top 5 estudiantes con más negativas
            </h2>
            <div className="overflow-x-auto">
              <table className="min-w-full text-left text-sm">
                <thead>
                  <tr className="border-b border-slate-200 text-slate-500 font-medium">
                    <th className="px-4 py-3 text-center w-16">Puesto</th>
                    <th className="px-4 py-3">Estudiante</th>
                    <th className="px-4 py-3 text-center w-28">Anotaciones</th>
                  </tr>
                </thead>
                <tbody>
                  {Array.from({ length: 5 }).map((_, index) => {
                    const neg = topNegative[index];
                    const rank = index + 1;

                    let rankClass = "text-slate-500 bg-slate-50";
                    if (rank === 1) rankClass = "bg-amber-100 text-amber-800 font-bold ring-1 ring-amber-200";
                    else if (rank === 2) rankClass = "bg-slate-100 text-slate-700 font-bold ring-1 ring-slate-200";
                    else if (rank === 3) rankClass = "bg-orange-50 text-orange-700 font-bold ring-1 ring-orange-100";

                    return (
                      <tr key={index} className="border-b border-slate-100 hover:bg-slate-50/50 transition">
                        <td className="px-4 py-3.5 text-center">
                          <span className={`inline-flex items-center justify-center h-6 w-6 rounded-full text-xs ${rankClass}`}>
                            {rank}
                          </span>
                        </td>
                        <td className="px-4 py-3.5 font-medium text-slate-800">
                          {neg ? neg.estudiante : <span className="text-slate-400 font-normal">-</span>}
                        </td>
                        <td className="px-4 py-3.5 text-center">
                          {neg ? (
                            <span className="inline-block rounded-full bg-rose-100 px-3 py-1 text-xs font-bold text-rose-700">
                              {neg.totalObservaciones}
                            </span>
                          ) : (
                            <span className="text-slate-400">-</span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </article>
        </section>

        <section className="grid gap-4">
          <article className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="mb-4 text-lg font-semibold text-slate-800 flex items-center gap-2">
              <Users className="h-5 w-5 text-indigo-500" />
              Top funcionarios con más anotaciones registradas
            </h2>
            <div className="overflow-x-auto">
              <table className="min-w-full text-left text-sm">
                <thead>
                  <tr className="border-b border-slate-200 text-slate-500 font-medium">
                    <th className="px-4 py-3 text-center w-16">Puesto</th>
                    <th className="px-4 py-3">Funcionario / Profesor</th>
                    <th className="px-4 py-3 text-center w-28">Positivas</th>
                    <th className="px-4 py-3 text-center w-28">Negativas</th>
                    <th className="px-4 py-3 text-center w-28">Total</th>
                  </tr>
                </thead>
                <tbody>
                  {topFuncionarios.length > 0 ? (
                    topFuncionarios.map((func, index) => {
                      const rank = index + 1;
                      let rankClass = "text-slate-500 bg-slate-50";
                      if (rank === 1) rankClass = "bg-amber-100 text-amber-800 font-bold ring-1 ring-amber-200";
                      else if (rank === 2) rankClass = "bg-slate-100 text-slate-700 font-bold ring-1 ring-slate-200";
                      else if (rank === 3) rankClass = "bg-orange-50 text-orange-700 font-bold ring-1 ring-orange-100";

                      return (
                        <tr key={index} className="border-b border-slate-100 hover:bg-slate-50/50 transition">
                          <td className="px-4 py-3.5 text-center">
                            <span className={`inline-flex items-center justify-center h-6 w-6 rounded-full text-xs ${rankClass}`}>
                              {rank}
                            </span>
                          </td>
                          <td className="px-4 py-3.5 font-medium text-slate-800">
                            {func.funcionario}
                          </td>
                          <td className="px-4 py-3.5 text-center">
                            <span className="inline-block rounded-full bg-emerald-100 px-3 py-1 text-xs font-bold text-emerald-700">
                              {func.positivas}
                            </span>
                          </td>
                          <td className="px-4 py-3.5 text-center">
                            <span className="inline-block rounded-full bg-rose-100 px-3 py-1 text-xs font-bold text-rose-700">
                              {func.negativas}
                            </span>
                          </td>
                          <td className="px-4 py-3.5 text-center">
                            <span className="inline-block rounded-full bg-indigo-100 px-3 py-1 text-xs font-bold text-indigo-700">
                              {func.total}
                            </span>
                          </td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={5} className="text-center py-6 text-slate-400 italic">
                        No se detectó información de funcionarios en la planilla cargada.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </article>
        </section>

        <section className="grid gap-4">
          <article className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="mb-4 text-lg font-semibold text-slate-800 flex items-center gap-2">
              <AlertCircle className="h-5 w-5 text-rose-500" />
              Detalle de faltas registradas
            </h2>
            <div className="overflow-x-auto max-h-72 overflow-y-auto">
              <table className="min-w-full text-left text-sm">
                <thead>
                  <tr className="border-b border-slate-200 text-slate-500 font-medium">
                    <th className="px-4 py-3">Tipo / Gravedad de Falta</th>
                    <th className="px-4 py-3 text-center w-28">Cantidad</th>
                    <th className="px-4 py-3 text-center w-28">Porcentaje</th>
                  </tr>
                </thead>
                <tbody>
                  {faltasStats.totalValidos > 0 ? (
                    faltasStats.data.map((item, index) => {
                      const percentage = Math.round((item.value / faltasStats.totalValidos) * 100);
                      return (
                        <tr key={index} className="border-b border-slate-100 hover:bg-slate-50/50 transition">
                          <td className="px-4 py-3 font-medium text-slate-800">{item.name}</td>
                          <td className="px-4 py-3 text-center font-bold text-slate-700">{item.value}</td>
                          <td className="px-4 py-3 text-center">
                            <span className="inline-block rounded-full bg-indigo-50 px-2.5 py-0.5 text-xs font-semibold text-indigo-600">
                              {percentage}%
                            </span>
                          </td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={3} className="text-center py-12 text-slate-400 italic">
                        No se detectó información de faltas en la planilla cargada (se requieren columnas como "Falta" o "Gravedad").
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </article>
        </section>

        <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
          <div className="mb-6 flex flex-col gap-4 border-b border-slate-100 pb-5">
            <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
              <h3 className="font-semibold text-slate-800 flex items-center gap-2">
                <Filter className="h-4 w-4 text-indigo-500" />
                Panel de Filtros Rápidos
              </h3>
              {(selectedAsignaturaObs !== "all" || selectedCursoObs !== "all" || selectedFechaObs !== "all" || searchQuery !== "" || quickFilter !== "all") && (
                <button
                  onClick={() => {
                    setSelectedAsignaturaObs("all");
                    setSelectedCursoObs("all");
                    setSelectedFechaObs("all");
                    setSearchQuery("");
                    setQuickFilter("all");
                  }}
                  className="text-xs text-indigo-600 hover:text-indigo-800 font-semibold underline cursor-pointer"
                >
                  Restablecer todos los filtros
                </button>
              )}
            </div>

            <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
              {/* 1. Búsqueda de texto */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                  Buscar
                </label>
                <div className="relative">
                  <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                  <input
                    value={searchQuery}
                    onChange={(event) => setSearchQuery(event.target.value)}
                    placeholder="Estudiante..."
                    className="w-full rounded-xl border border-slate-300 bg-white py-2 pl-9 pr-3 text-xs outline-none ring-indigo-500 transition focus:ring"
                  />
                </div>
              </div>

              {/* 2. Filtro de Tipo */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                  Tipo
                </label>
                <select
                  value={quickFilter}
                  onChange={(e) => setQuickFilter(e.target.value as any)}
                  className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-xs outline-none ring-indigo-500 focus:ring cursor-pointer"
                >
                  <option value="all">Ver todo</option>
                  <option value="positiva">Solo positivas</option>
                  <option value="negativa">Solo negativas</option>
                  <option value="otros">Otros / Entrevistas</option>
                </select>
              </div>

              {/* 3. Filtro de Asignatura */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                  Gravedad
                </label>
                <select
                  value={selectedAsignaturaObs}
                  onChange={(e) => setSelectedAsignaturaObs(e.target.value)}
                  className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-xs outline-none ring-indigo-500 focus:ring cursor-pointer"
                >
                  <option value="all">Todas</option>
                  {uniqueAsignaturasObs.map((asig) => (
                    <option key={asig} value={asig}>
                      {asig}
                    </option>
                  ))}
                </select>
              </div>

              {/* 5. Filtro de Fecha */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                  Fecha
                </label>
                <select
                  value={selectedFechaObs}
                  onChange={(e) => setSelectedFechaObs(e.target.value)}
                  className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-xs outline-none ring-indigo-500 focus:ring cursor-pointer"
                >
                  <option value="all">Todas</option>
                  {uniqueFechasObs.map((f) => (
                    <option key={f} value={f}>
                      {formatDateToVerbal(f)}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="min-w-full text-left text-sm">
              <thead>
                <tr className="border-b border-slate-200 text-slate-500 font-medium">
                  <th className="px-4 py-3 min-w-[100px]">Fecha</th>
                  <th className="px-4 py-3 min-w-[200px]">Estudiante</th>
                  <th className="px-4 py-3 min-w-[140px] whitespace-nowrap">Tipo</th>
                  <th className="px-4 py-3">Descripción</th>
                </tr>
              </thead>
              <tbody>
                {filteredRows.map((row) => (
                  <tr key={row.id} className="border-b border-slate-100 hover:bg-slate-50/50 transition">
                    <td className="px-4 py-3 text-slate-600 whitespace-nowrap">{row.fechaTexto || "-"}</td>
                    <td className="px-4 py-3 font-medium whitespace-nowrap">{row.nombreCompleto}</td>
                    <td className="px-4 py-3 whitespace-nowrap">
                      <span className={`inline-block rounded-full px-2.5 py-1 text-xs font-semibold ${typeBadgeClasses(row.tipo)}`}>
                        {typeLabel(row.tipo)}
                      </span>
                    </td>
                    <td className="px-4 py-3 text-slate-600">{row.descripcion}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </main>
  );
}
