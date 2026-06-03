"use client";

import Papa from "papaparse";
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
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import { ChangeEvent, DragEvent, useMemo, useState } from "react";

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

function buildStudentName(row: Record<string, string>) {
  return [
    row["Nombre Estudiante"],
    row["Primer Apellido Estudiante"],
    row["Segundo Apellido Estudiante"],
  ]
    .map((item) => item?.trim())
    .filter(Boolean)
    .join(" ");
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

export default function Home() {
  const [observations, setObservations] = useState<Observation[]>([]);
  const [errorMessage, setErrorMessage] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [quickFilter, setQuickFilter] = useState<QuickFilter>("all");

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

    return [...map.values()].sort((a, b) => a.fecha.localeCompare(b.fecha));
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

  const filteredRows = useMemo(() => {
    const term = searchQuery.trim().toLowerCase();

    return observations.filter((item) => {
      const matchesSearch = !term || item.nombreCompleto.toLowerCase().includes(term);
      const matchesType = quickFilter === "all" || item.tipo === quickFilter;
      return matchesSearch && matchesType;
    });
  }, [observations, quickFilter, searchQuery]);

  const parseAndLoadCsv = (file: File) => {
    setErrorMessage("");
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

  const handleInputChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    parseAndLoadCsv(file);
    event.target.value = "";
  };

  const handleDrop = (event: DragEvent<HTMLLabelElement>) => {
    event.preventDefault();
    setIsDragging(false);
    const file = event.dataTransfer.files?.[0];
    if (!file) return;
    parseAndLoadCsv(file);
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
                Carga un CSV de observaciones escolares para visualizar métricas y tendencias.
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
            className={`flex cursor-pointer flex-col items-center justify-center gap-3 rounded-2xl border-2 border-dashed p-12 text-center transition ${
              isDragging
                ? "border-indigo-500 bg-indigo-50"
                : "border-slate-300 bg-slate-50 hover:border-indigo-400"
            }`}
          >
            <div className="rounded-full bg-white p-3 shadow-sm">
              <Upload className="h-6 w-6 text-indigo-500" />
            </div>
            <div>
              <p className="font-medium">Arrastra y suelta tu CSV aquí</p>
              <p className="text-sm text-slate-500">o haz clic para seleccionarlo desde tu equipo</p>
            </div>
            <input
              type="file"
              accept=".csv,text/csv"
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
            <h2 className="mb-3 font-semibold">Evolución temporal (positivas vs negativas)</h2>
            <ResponsiveContainer width="100%" height="90%">
              <AreaChart data={trendData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis dataKey="fecha" />
                <YAxis allowDecimals={false} />
                <Tooltip />
                <Legend />
                <Area type="monotone" dataKey="positivas" stroke="#10b981" fill="#10b98133" />
                <Area type="monotone" dataKey="negativas" stroke="#ef4444" fill="#ef444433" />
              </AreaChart>
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
                  className={`rounded-xl px-3 py-2 transition ${
                    quickFilter === option.value
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
                <tr className="border-b border-slate-200 text-slate-500">
                  <th className="px-2 py-3">Fecha</th>
                  <th className="px-2 py-3">Estudiante</th>
                  <th className="px-2 py-3">Tipo</th>
                  <th className="px-2 py-3">Descripción</th>
                </tr>
              </thead>
              <tbody>
                {filteredRows.map((row) => (
                  <tr key={row.id} className="border-b border-slate-100">
                    <td className="px-2 py-3 text-slate-600">{row.fechaTexto || "-"}</td>
                    <td className="px-2 py-3 font-medium">{row.nombreCompleto}</td>
                    <td className="px-2 py-3">
                      <span className={`rounded-full px-2 py-1 text-xs font-medium ${typeBadgeClasses(row.tipo)}`}>
                        {typeLabel(row.tipo)}
                      </span>
                    </td>
                    <td className="px-2 py-3 text-slate-600">{row.descripcion}</td>
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
