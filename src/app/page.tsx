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

const PIE_COLORS = ["#10b981", "#ef4444", "#6366f1"];
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
  const [visibleCount, setVisibleCount] = useState(25);
  const observerRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    setVisibleCount(25);
  }, [searchQuery, quickFilter, observations]);

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
    total > 0 ? Math.round((summary.positivas / total) * 100) : 0;

  const pieData = useMemo(
    () => [
      { name: "Positivas", value: summary.positivas },
      { name: "Negativas", value: summary.negativas },
      { name: "Otros/Entrevistas", value: summary.otros },
    ],
    [summary],
  );

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
      return matchesSearch && matchesType;
    });
  }, [observations, quickFilter, searchQuery]);

  useEffect(() => {
    if (!observerRef.current) return;
    const observer = new IntersectionObserver(
      (entries) => {
        if (entries[0].isIntersecting) {
          setVisibleCount((prev) => prev + 25);
        }
      },
      { threshold: 0.1 }
    );
    observer.observe(observerRef.current);
    return () => observer.disconnect();
  }, [filteredRows, visibleCount]);

  const parseAndLoadCsv = (file: File) => {
    Papa.parse<Record<string, string>>(file, {
      header: true,
      skipEmptyLines: true,
      complete: ({ data, meta }) => {
        const missingHeaders = REQUIRED_HEADERS.filter((header) => !meta.fields?.includes(header));
        if (missingHeaders.length > 0) {
          setObservations([]);
          setErrorMessage(`Faltan columnas obligatorias: ${missingHeaders.join(", ")}`);
          return;
        }

        const parsed = data
          .map((row, index) => {
            const nombreCompleto = buildStudentName(row);
            const fechaTexto = row["Fecha"]?.trim() ?? "";
            const fechaOrdenable = parseDateToSortable(fechaTexto);

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

  const parseAndLoadExcel = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
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
          setErrorMessage(`Faltan columnas obligatorias: ${missingHeaders.join(", ")}`);
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
            } satisfies Observation;
          })
          .filter((item) => item.nombreCompleto && item.descripcion && item.tipoOriginal);

        setObservations(parsed);
      } catch (error: any) {
        setObservations([]);
        setErrorMessage(`No fue posible leer el archivo Excel: ${error.message}`);
      }
    };
    reader.onerror = () => {
      setObservations([]);
      setErrorMessage("Error al leer el archivo.");
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFile = (file: File) => {
    setErrorMessage("");
    const fileName = file.name.toLowerCase();
    if (fileName.endsWith(".csv")) {
      parseAndLoadCsv(file);
    } else if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
      parseAndLoadExcel(file);
    } else {
      setErrorMessage("Por favor, sube un archivo CSV o Excel (.xlsx, .xls).");
    }
  };

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

  if (observations.length === 0) {
    return (
      <main className="min-h-screen bg-slate-100 px-4 py-10 text-slate-900">
        <section className="mx-auto max-w-3xl rounded-2xl border border-slate-200 bg-white p-8 shadow-sm">
          <div className="mb-8 flex items-center gap-3">
            <div className="rounded-xl bg-indigo-100 p-2 text-indigo-600">
              <Users className="h-5 w-5" />
            </div>
            <div>
              <h1 className="text-2xl font-semibold">Kimche Analyzer</h1>
              <p className="text-sm text-slate-500">
                Carga un CSV o Excel de observaciones escolares para visualizar métricas y tendencias.
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
              <p className="font-medium">Arrastra y suelta tu CSV o Excel aquí</p>
              <p className="text-sm text-slate-500">o haz clic para seleccionarlo desde tu equipo</p>
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

  return (
    <main className="min-h-screen bg-slate-100 px-4 py-8 text-slate-900">
      <div className="mx-auto flex max-w-7xl flex-col gap-6">
        <header className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
          <h1 className="text-2xl font-semibold">Dashboard de Convivencia Escolar</h1>
          <p className="mt-1 text-sm text-slate-500">
            Total de registros cargados: <span className="font-semibold text-slate-700">{total}</span>
          </p>
        </header>

        <section className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
          <article className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <p className="text-sm text-slate-500">Observaciones procesadas</p>
            <p className="mt-3 text-3xl font-semibold">{total}</p>
          </article>
          <article className="rounded-2xl border border-emerald-200 bg-emerald-50 p-5 shadow-sm">
            <div className="flex items-center justify-between">
              <p className="text-sm text-emerald-700">Anotaciones positivas</p>
              <CheckCircle2 className="h-5 w-5 text-emerald-600" />
            </div>
            <p className="mt-3 text-3xl font-semibold text-emerald-700">{summary.positivas}</p>
          </article>
          <article className="rounded-2xl border border-rose-200 bg-rose-50 p-5 shadow-sm">
            <div className="flex items-center justify-between">
              <p className="text-sm text-rose-700">Anotaciones negativas</p>
              <AlertCircle className="h-5 w-5 text-rose-600" />
            </div>
            <p className="mt-3 text-3xl font-semibold text-rose-700">{summary.negativas}</p>
          </article>
          <article className="rounded-2xl border border-indigo-200 bg-indigo-50 p-5 shadow-sm">
            <div className="flex items-center justify-between">
              <p className="text-sm text-indigo-700">Ratio de convivencia</p>
              <Star className="h-5 w-5 text-indigo-600" />
            </div>
            <p className="mt-3 text-3xl font-semibold text-indigo-700">
              {positiveObservationsPercentage}%
            </p>
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
                    className={`rounded-md px-2.5 py-1 text-xs font-semibold transition ${
                      timeResolution === "daily"
                        ? "bg-white text-slate-800 shadow-sm"
                        : "text-slate-500 hover:text-slate-850"
                    }`}
                  >
                    Diario (Líneas)
                  </button>
                  <button
                    type="button"
                    onClick={() => setTimeResolution("monthly")}
                    className={`rounded-md px-2.5 py-1 text-xs font-semibold transition ${
                      timeResolution === "monthly"
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
          <article className="h-80 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
            <h2 className="mb-3 font-semibold">Top 5 estudiantes con más positivas</h2>
            <ResponsiveContainer width="100%" height="90%">
              <BarChart data={topPositive} layout="vertical" margin={{ left: 30 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis type="number" allowDecimals={false} />
                <YAxis type="category" dataKey="estudiante" width={160} />
                <Tooltip />
                <Bar dataKey="totalObservaciones" fill="#10b981" radius={[0, 8, 8, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </article>

          <article className="h-80 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
            <h2 className="mb-3 font-semibold">Top 5 estudiantes con más negativas</h2>
            <ResponsiveContainer width="100%" height="90%">
              <BarChart data={topNegative} layout="vertical" margin={{ left: 30 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis type="number" allowDecimals={false} />
                <YAxis type="category" dataKey="estudiante" width={160} />
                <Tooltip />
                <Bar dataKey="totalObservaciones" fill="#ef4444" radius={[0, 8, 8, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </article>
        </section>

        <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
          <div className="mb-4 flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
            <div className="relative w-full lg:max-w-sm">
              <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
              <input
                value={searchQuery}
                onChange={(event) => setSearchQuery(event.target.value)}
                placeholder="Buscar estudiante..."
                className="w-full rounded-xl border border-slate-300 bg-white py-2 pl-9 pr-3 text-sm outline-none ring-indigo-500 transition focus:ring"
              />
            </div>

            <div className="flex items-center gap-2 text-sm">
              <Filter className="h-4 w-4 text-slate-500" />
              {([
                { value: "all", label: "Ver todo" },
                { value: "positiva", label: "Solo positivas" },
                { value: "negativa", label: "Solo negativas" },
              ] as const).map((option) => (
                <button
                  key={option.value}
                  type="button"
                  onClick={() => setQuickFilter(option.value)}
                  className={`rounded-xl px-3 py-2 transition ${quickFilter === option.value
                    ? "bg-indigo-600 text-white"
                    : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                    }`}
                >
                  {option.label}
                </button>
              ))}
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
                {filteredRows.slice(0, visibleCount).map((row) => (
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

          {filteredRows.length > visibleCount && (
            <div ref={observerRef} className="mt-4 py-6 text-center flex flex-col items-center justify-center gap-2">
              <span className="text-xs text-slate-400">
                Mostrando {visibleCount} de {filteredRows.length} registros
              </span>
              <button
                type="button"
                onClick={() => setVisibleCount((prev) => prev + 25)}
                className="rounded-xl bg-slate-100 px-5 py-2.5 text-xs font-semibold text-slate-700 hover:bg-slate-200 transition active:scale-95 shadow-sm border border-slate-200"
              >
                Cargar más registros...
              </button>
            </div>
          )}
        </section>
      </div>
    </main>
  );
}
