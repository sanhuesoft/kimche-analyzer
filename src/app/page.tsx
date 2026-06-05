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
  Info,
  BookOpen,
} from "lucide-react";
import {
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
import { ChangeEvent, DragEvent, useEffect, useMemo, useState, useCallback } from "react";

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

type GradeItem = {
  label: string;
  weight: string;
  value: number | null;
};

type CalificacionSubjectData = {
  subjectName: string;
  grades: GradeItem[];
  p1: number | null;
  p2: number | null;
  pf: number | null;
};

type CalificacionStudent = {
  id: string;
  lista: string;
  estudiante: string;
  run: string;
  dv: string;
  periodo1: number | null;
  periodo2: number | null;
  promedioGeneral: number | null;
  subjects: CalificacionSubjectData[];
};

interface ExcelSectionValue {
  subjectName?: string;
  type: "grade" | "p1" | "p2" | "pf" | "p1_gen" | "p2_gen" | "pf_gen";
  label?: string;
  value: number | string | null;
}

interface ExcelSectionRow {
  label: string;
  values: ExcelSectionValue[];
}

interface ExcelExtraData {
  indicadores: ExcelSectionRow[];
  categorias: ExcelSectionRow[];
  leyendas: string[];
}

interface RiesgoStudent {
  id: string;
  lista: string;
  estudiante: string;
  run: string;
  dv: string;
  promedioGeneral: number | null;
  asistencia: string | null;
  dynamicFields: Record<string, string | number | null>;
}

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


  const [appMode, setAppMode] = useState<"observations" | "pendientes" | "calificaciones" | "panorama_riesgo">("observations");
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
  const [uploadedFilesCount, setUploadedFilesCount] = useState<number>(0);
  const [activeCourseTab, setActiveCourseTab] = useState<string>("all");

  const [calificaciones, setCalificaciones] = useState<CalificacionStudent[]>([]);
  const [selectedCalificacionesAsignatura, setSelectedCalificacionesAsignatura] = useState<string>("all");
  const [onlyShowAtRiskCalificaciones, setOnlyShowAtRiskCalificaciones] = useState<boolean>(false);
  const [calificacionesSearch, setCalificacionesSearch] = useState<string>("");
  const [excelExtraData, setExcelExtraData] = useState<ExcelExtraData | null>(null);

  const [riesgoStudents, setRiesgoStudents] = useState<RiesgoStudent[]>([]);
  const [riesgoHeaders, setRiesgoHeaders] = useState<string[]>([]);
  const [panoramaActiveTab, setPanoramaActiveTab] = useState<"panorama" | "riesgo">("panorama");
  const [riesgoSearch, setRiesgoSearch] = useState<string>("");
  const [riesgoFilter, setRiesgoFilter] = useState<"all" | "average" | "attendance">("all");

  const resetAll = () => {
    setObservations([]);
    setPendientes([]);
    setCalificaciones([]);
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
    setUploadedFilesCount(0);
    setActiveCourseTab("all");
    setSelectedCalificacionesAsignatura("all");
    setOnlyShowAtRiskCalificaciones(false);
    setCalificacionesSearch("");
    setExcelExtraData(null);
    setRiesgoStudents([]);
    setRiesgoHeaders([]);
    setPanoramaActiveTab("panorama");
    setRiesgoSearch("");
    setRiesgoFilter("all");
  };

  const promedioGeneralCurso = useMemo(() => {
    const validGrades = calificaciones.map(c => c.promedioGeneral).filter((g): g is number => g !== null && g > 0);
    if (validGrades.length === 0) return 0;
    const sum = validGrades.reduce((a, b) => a + b, 0);
    return parseFloat((sum / validGrades.length).toFixed(2));
  }, [calificaciones]);

  const tasaAprobacion = useMemo(() => {
    const validGrades = calificaciones.map(c => c.promedioGeneral).filter((g): g is number => g !== null && g > 0);
    if (validGrades.length === 0) return 0;
    const passing = validGrades.filter(g => g >= 4.0).length;
    return Math.round((passing / validGrades.length) * 100);
  }, [calificaciones]);

  const uniqueAsignaturasCalificaciones = useMemo(() => {
    const subjectsSet = new Set<string>();
    for (const student of calificaciones) {
      for (const sub of student.subjects) {
        if (sub.subjectName) {
          subjectsSet.add(sub.subjectName);
        }
      }
    }
    return Array.from(subjectsSet).sort();
  }, [calificaciones]);

  const distributionData = useMemo(() => {
    let reprobados = 0;
    let suficiente = 0;
    let bueno = 0;
    let excelente = 0;

    for (const c of calificaciones) {
      const avg = c.promedioGeneral;
      if (avg === null || avg === 0) continue;
      if (avg < 4.0) reprobados++;
      else if (avg < 5.0) suficiente++;
      else if (avg < 6.0) bueno++;
      else excelente++;
    }

    return [
      { name: "Reprobado (1.0-3.9)", value: reprobados, color: "#ef4444" },
      { name: "Suficiente (4.0-4.9)", value: suficiente, color: "#f59e0b" },
      { name: "Bueno (5.0-5.9)", value: bueno, color: "#3b82f6" },
      { name: "Excelente (6.0-7.0)", value: excelente, color: "#10b981" },
    ].filter(item => item.value > 0);
  }, [calificaciones]);

  const subjectPerformanceData = useMemo(() => {
    const map = new Map<string, { sum: number; count: number }>();

    for (const student of calificaciones) {
      for (const sub of student.subjects) {
        if (sub.pf !== null && sub.pf > 0) {
          const existing = map.get(sub.subjectName) || { sum: 0, count: 0 };
          map.set(sub.subjectName, {
            sum: existing.sum + sub.pf,
            count: existing.count + 1,
          });
        }
      }
    }

    return Array.from(map.entries()).map(([name, stats]) => ({
      name,
      promedio: parseFloat((stats.sum / stats.count).toFixed(2)),
    })).sort((a, b) => a.name.localeCompare(b.name));
  }, [calificaciones]);

  const filteredCalificaciones = useMemo(() => {
    const term = calificacionesSearch.trim().toLowerCase();

    return calificaciones.filter(student => {
      const matchesSearch = !term || student.estudiante.toLowerCase().includes(term);

      let matchesSubject = true;
      if (selectedCalificacionesAsignatura !== "all") {
        const sub = student.subjects.find(s => s.subjectName === selectedCalificacionesAsignatura);
        matchesSubject = !!sub;
      }

      let matchesAtRisk = true;
      if (onlyShowAtRiskCalificaciones) {
        if (selectedCalificacionesAsignatura === "all") {
          matchesAtRisk = student.promedioGeneral !== null && student.promedioGeneral < 4.0;
        } else {
          const sub = student.subjects.find(s => s.subjectName === selectedCalificacionesAsignatura);
          matchesAtRisk = sub ? (sub.pf !== null && sub.pf < 4.0) : false;
        }
      }

      return matchesSearch && matchesSubject && matchesAtRisk;
    });
  }, [calificaciones, calificacionesSearch, selectedCalificacionesAsignatura, onlyShowAtRiskCalificaciones]);

  const riesgoMetrics = useMemo(() => {
    const totalRiesgo = riesgoStudents.length;

    const validGrades = riesgoStudents.map(s => s.promedioGeneral).filter((g): g is number => g !== null && g > 0);
    const avgGrade = validGrades.length > 0 ? validGrades.reduce((a, b) => a + b, 0) / validGrades.length : 0;

    let totalAsist = 0;
    let asistCount = 0;
    let countLowAsist = 0;
    let countLowGrade = 0;

    riesgoStudents.forEach(s => {
      if (s.promedioGeneral !== null && s.promedioGeneral < 4.0) {
        countLowGrade++;
      }

      if (s.asistencia) {
        const clean = s.asistencia.replace("%", "").trim();
        const num = parseFloat(clean);
        if (!isNaN(num)) {
          totalAsist += num;
          asistCount++;
          if (num < 85) {
            countLowAsist++;
          }
        }
      }
    });

    const avgAsist = asistCount > 0 ? Math.round(totalAsist / asistCount) : 0;

    return {
      totalRiesgo,
      avgGrade: parseFloat(avgGrade.toFixed(2)),
      avgAsist,
      countLowAsist,
      countLowGrade,
    };
  }, [riesgoStudents]);

  const filteredRiesgoStudents = useMemo(() => {
    const term = riesgoSearch.trim().toLowerCase();
    return riesgoStudents.filter(s => {
      const matchesSearch = !term || s.estudiante.toLowerCase().includes(term);

      let matchesFilter = true;
      if (riesgoFilter === "average") {
        matchesFilter = s.promedioGeneral !== null && s.promedioGeneral < 4.0;
      } else if (riesgoFilter === "attendance") {
        if (s.asistencia) {
          const num = parseFloat(s.asistencia.replace("%", "").trim());
          matchesFilter = !isNaN(num) && num < 85;
        } else {
          matchesFilter = false;
        }
      }

      return matchesSearch && matchesFilter;
    });
  }, [riesgoStudents, riesgoSearch, riesgoFilter]);

  const visibleRiesgoHeaders = useMemo(() => {
    return riesgoHeaders
      .filter(h => {
        const lower = h.toLowerCase();
        return !(
          lower === "año" ||
          lower.includes("tipo de enseñanza") ||
          lower === "nivel educativo" ||
          lower === "run" ||
          lower.includes("dígito") ||
          lower.includes("verificador") ||
          lower.includes("apellido paterno") ||
          lower.includes("apellido materno") ||
          lower === "paterno" ||
          lower === "materno"
        );
      })
      .map(h => {
        let display = h;
        if (h.toLowerCase() === "nombres" || h.toLowerCase() === "nombre") {
          display = "Estudiante";
        }
        return { original: h, display };
      });
  }, [riesgoHeaders]);

  const subjectColumnStats = useMemo(() => {
    let columns: { label: string, type: "grade" | "p1" | "p2" | "pf" | "p1_gen" | "p2_gen" | "pf_gen" }[] = [];

    if (selectedCalificacionesAsignatura !== "all") {
      const selectedSubObj = calificaciones.find(s => s.subjects.some(sub => sub.subjectName === selectedCalificacionesAsignatura))
        ?.subjects.find(sub => sub.subjectName === selectedCalificacionesAsignatura);
      if (!selectedSubObj) return null;

      columns = [
        ...selectedSubObj.grades.map(g => ({ label: g.label, type: "grade" as const })),
        { label: "P1", type: "p1" as const },
        { label: "P2", type: "p2" as const },
        { label: "PF", type: "pf" as const },
      ];
    } else {
      columns = [
        { label: "P1 General", type: "p1_gen" as const },
        { label: "P2 General", type: "p2_gen" as const },
        { label: "Promedio General", type: "pf_gen" as const },
      ];
    }

    const result: Record<string, { promedio: number | null, minimo: number | null, maximo: number | null, desviacion: number | null }> = {};

    for (const col of columns) {
      const grades: number[] = [];
      for (const student of filteredCalificaciones) {
        let val: number | null = null;

        if (selectedCalificacionesAsignatura !== "all") {
          const studentSub = student.subjects.find(s => s.subjectName === selectedCalificacionesAsignatura);
          if (studentSub) {
            if (col.type === "grade") {
              val = studentSub.grades.find(g => g.label === col.label)?.value ?? null;
            } else if (col.type === "p1") {
              val = studentSub.p1;
            } else if (col.type === "p2") {
              val = studentSub.p2;
            } else if (col.type === "pf") {
              val = studentSub.pf;
            }
          }
        } else {
          if (col.type === "p1_gen") {
            val = student.periodo1;
          } else if (col.type === "p2_gen") {
            val = student.periodo2;
          } else if (col.type === "pf_gen") {
            val = student.promedioGeneral;
          }
        }

        if (val !== null && val > 0) {
          grades.push(val);
        }
      }

      if (grades.length > 0) {
        const sum = grades.reduce((a, b) => a + b, 0);
        const avg = sum / grades.length;
        const min = Math.min(...grades);
        const max = Math.max(...grades);

        const variance = grades.reduce((a, b) => a + Math.pow(b - avg, 2), 0) / grades.length;
        const stdDev = Math.sqrt(variance);

        result[col.label] = {
          promedio: parseFloat(avg.toFixed(2)),
          minimo: min,
          maximo: max,
          desviacion: parseFloat(stdDev.toFixed(2)),
        };
      } else {
        result[col.label] = { promedio: null, minimo: null, maximo: null, desviacion: null };
      }
    }

    return result;
  }, [filteredCalificaciones, selectedCalificacionesAsignatura, calificaciones]);



  const getExcelCategoryVal = useCallback((
    rowLabel: string,
    type: ExcelSectionValue["type"],
    subjectName?: string,
    gradeLabel?: string
  ): number | string | null => {
    if (!excelExtraData) return null;
    const cleanLabel = rowLabel.trim().toLowerCase();
    const row = excelExtraData.categorias.find(r => {
      const l = r.label.trim().toLowerCase();
      return l.includes(cleanLabel) || cleanLabel.includes(l);
    });
    if (!row) return null;

    const match = row.values.find(v => {
      if (type === "grade") {
        return v.type === "grade" && v.subjectName === subjectName && v.label === gradeLabel;
      }
      if (subjectName) {
        return v.type === type && v.subjectName === subjectName;
      }
      return v.type === type;
    });
    return match ? match.value : null;
  }, [excelExtraData]);

  const displayObservations = useMemo(() => {
    if (activeCourseTab === "all") return observations;
    return observations.filter((o) => o.curso === activeCourseTab);
  }, [observations, activeCourseTab]);

  const total = displayObservations.length;

  const summary = useMemo(() => {
    const result = { positivas: 0, negativas: 0, otros: 0 };
    for (const row of displayObservations) {
      if (row.tipo === "positiva") result.positivas += 1;
      if (row.tipo === "negativa") result.negativas += 1;
      if (row.tipo === "otros") result.otros += 1;
    }
    return result;
  }, [displayObservations]);

  const positiveObservationsPercentage =
    (summary.positivas + summary.negativas) > 0
      ? Math.round((summary.positivas / (summary.positivas + summary.negativas)) * 100)
      : 0;

  const pieData = useMemo(() => {
    let positivas = 0;
    let negativas = 0;
    let entrevistas = 0;
    let otros = 0;

    for (const row of displayObservations) {
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
  }, [displayObservations]);

  const trendData = useMemo(() => {
    const map = new Map<string, { fecha: string; positivas: number; negativas: number }>();

    for (const item of displayObservations) {
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
  }, [displayObservations]);

  const monthlyData = useMemo(() => {
    const map = new Map<string, { mes: string; positivas: number; negativas: number }>();

    for (const item of displayObservations) {
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
  }, [displayObservations]);

  const { topPositive, topNegative } = useMemo(() => {
    const positives = new Map<string, number>();
    const negatives = new Map<string, number>();

    for (const item of displayObservations) {
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
  }, [displayObservations]);

  const { topPositiveFuncionarios, topNegativeFuncionarios } = useMemo(() => {
    const positives = new Map<string, number>();
    const negatives = new Map<string, number>();

    for (const item of displayObservations) {
      const func = item.funcionario || "No especificado";
      if (item.tipo === "positiva") {
        positives.set(func, (positives.get(func) ?? 0) + 1);
      }
      if (item.tipo === "negativa") {
        negatives.set(func, (negatives.get(func) ?? 0) + 1);
      }
    }

    const toTop = (source: Map<string, number>) =>
      [...source.entries()]
        .map(([funcionario, totalObservaciones]) => ({ funcionario, totalObservaciones }))
        .sort((a, b) => b.totalObservaciones - a.totalObservaciones)
        .slice(0, 5);

    return {
      topPositiveFuncionarios: toTop(positives),
      topNegativeFuncionarios: toTop(negatives),
    };
  }, [displayObservations]);

  const faltasStats = useMemo(() => {
    const counts = new Map<string, number>();
    let totalValidos = 0;

    for (const item of displayObservations) {
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
  }, [displayObservations]);

  const { latestPositive, latestNegative } = useMemo(() => {
    let latestPos: Observation | null = null;
    let latestNeg: Observation | null = null;

    const sorted = [...displayObservations]
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
      latestPos = displayObservations.find((obs) => obs.tipo === "positiva") ?? null;
    }
    if (!latestNeg) {
      latestNeg = displayObservations.find((obs) => obs.tipo === "negativa") ?? null;
    }

    return { latestPositive: latestPos, latestNegative: latestNeg };
  }, [displayObservations]);

  const filteredRows = useMemo(() => {
    const term = searchQuery.trim().toLowerCase();

    return displayObservations.filter((item) => {
      const matchesSearch = !term || item.nombreCompleto.toLowerCase().includes(term);
      const matchesType = quickFilter === "all" || item.tipo === quickFilter;
      const matchesAsignatura = selectedAsignaturaObs === "all" || item.asignaturaOrCategorizacion === selectedAsignaturaObs;
      const matchesCurso = selectedCursoObs === "all" || item.curso === selectedCursoObs;
      const matchesFecha = selectedFechaObs === "all" || item.fechaTexto === selectedFechaObs;
      return matchesSearch && matchesType && matchesAsignatura && matchesCurso && matchesFecha;
    });
  }, [displayObservations, quickFilter, searchQuery, selectedAsignaturaObs, selectedCursoObs, selectedFechaObs]);



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
    const counts = new Map<string, number>();
    for (const o of displayObservations) {
      if (o.asignaturaOrCategorizacion) {
        counts.set(o.asignaturaOrCategorizacion, (counts.get(o.asignaturaOrCategorizacion) ?? 0) + 1);
      }
    }
    return [...counts.entries()]
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [displayObservations]);

  const uniqueCursosObs = useMemo(() => {
    return Array.from(new Set(observations.map((o) => o.curso).filter(Boolean))).sort();
  }, [observations]);

  const uniqueFechasObs = useMemo(() => {
    const rawFechas = Array.from(new Set(displayObservations.map((o) => o.fechaTexto).filter(Boolean)));
    return rawFechas.sort((a, b) => {
      const dateA = parseDateToSortable(a);
      const dateB = parseDateToSortable(b);
      return dateA.localeCompare(dateB);
    });
  }, [displayObservations]);

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

  const processSingleFile = useCallback((file: File): Promise<{
    fileName: string;
    type: "observations" | "pendientes" | "calificaciones" | "panorama_riesgo";
    observationsData?: Observation[];
    pendientesData?: PendienteItem[];
    calificacionesData?: CalificacionStudent[];
    excelExtraData?: ExcelExtraData;
    riesgoData?: RiesgoStudent[];
    riesgoHeaders?: string[];
    error?: string;
  }> => {
    return new Promise((resolve) => {
      const fileName = file.name.toLowerCase();
      if (fileName.endsWith(".csv")) {
        Papa.parse<Record<string, string>>(file, {
          header: true,
          skipEmptyLines: true,
          complete: ({ data, meta }) => {
            const missingHeaders = REQUIRED_HEADERS.filter((header) => !meta.fields?.includes(header));
            if (missingHeaders.length > 0) {
              resolve({ fileName: file.name, type: "observations", error: "Kimche Analyzer no puede procesar ese tipo de planilla aún" });
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
                  id: `${file.name}-${fechaTexto}-${nombreCompleto}-${row["Tipo de observación"]?.trim() ?? ""}-${index}`,
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

            resolve({ fileName: file.name, type: "observations", observationsData: parsed });
          },
          error: (error) => {
            resolve({ fileName: file.name, type: "observations", error: `No fue posible leer el archivo: ${error.message}` });
          },
        });
      } else if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: "array" });

            const firstSheetName = workbook.SheetNames[0];
            const sheetNames = workbook.SheetNames;
            const hasRiesgoSheet = sheetNames.some(
              (name) =>
                name.toLowerCase().includes("riesgo") &&
                name.toLowerCase().includes("promoci")
            );

            let isPanoramaRiesgo = hasRiesgoSheet;
            let isPendientes = false;
            let isCalificaciones = false;
            let jsonGrid: unknown[][] = [];

            if (!isPanoramaRiesgo && firstSheetName) {
              const worksheet = workbook.Sheets[firstSheetName];
              jsonGrid = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: "" });

              if (jsonGrid.length >= 3) {
                const row1 = jsonGrid[0].map((h: unknown) => String(h || "").trim().toLowerCase());
                const row2 = jsonGrid[1].map((h: unknown) => String(h || "").trim().toLowerCase());

                const hasEstudiante = row1.some(h => h.includes("estudiante") || h.includes("alumno"));
                const hasPromedio = row1.some(h => h.includes("promedio general") || h.includes("promediogeneral"));
                const hasN1 = row2.some(h => h === "n1" || h === "n2" || h === "n°1");

                if (hasEstudiante && (hasPromedio || hasN1)) {
                  isCalificaciones = true;
                }
              }
            }

            const hasPendientesSheetNames = workbook.SheetNames.some(
              (name) =>
                name.toLowerCase().includes("firma") ||
                name.toLowerCase().includes("leccionario") ||
                name.toLowerCase().includes("registro")
            );

            if (!isPanoramaRiesgo && !isCalificaciones) {
              if (hasPendientesSheetNames) {
                isPendientes = true;
              } else if (jsonGrid.length > 0) {
                const headers = jsonGrid[0].map((h: unknown) => String(h || "").trim().toLowerCase());
                isPendientes = headers.includes("docente titular") || headers.includes("docente") || headers.includes("estado de la firma") || headers.includes("hora de clase");
              }
            }

            if (isPanoramaRiesgo) {
              const firstSheet = workbook.Sheets[firstSheetName];
              const grid1 = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, { header: 1, defval: "" });

              // Robust header row detection
              let mainHeaderRowIdx1 = 4;
              for (let r = 0; r < Math.min(5, grid1.length); r++) {
                const row = grid1[r] || [];
                const nonEmptyCount = row.filter(cell => String(cell || "").trim() !== "").length;
                if (nonEmptyCount >= 4) {
                  const hasEst = row.some(val => {
                    const s = String(val || "").toLowerCase();
                    return s.includes("estudiante") || s.includes("alumno") || s.includes("nombre") || s.includes("apellido");
                  });
                  if (hasEst) {
                    mainHeaderRowIdx1 = r;
                    break;
                  }
                }
              }

              const headersRow1 = grid1[mainHeaderRowIdx1] || [];
              const subheadersRow1 = grid1[mainHeaderRowIdx1 + 1] || [];

              const colInfo1 = headersRow1.map((h, cIdx) => ({
                label: String(h || "").trim(),
                colIndex: cIdx,
              })).filter(c => c.label !== "");

              let colLista1 = -1;
              let colEstudiante1 = -1;
              let colPaterno1 = -1;
              let colMaterno1 = -1;
              let colNombres1 = -1;
              let colRun1 = -1;
              let colDv1 = -1;
              let colProm1 = -1;
              let colPeriodo1_1 = -1;
              let colPeriodo2_1 = -1;

              colInfo1.forEach(c => {
                const l = c.label.toLowerCase();
                if (l.includes("lista") || l.includes("nº") || l.includes("n°")) {
                  colLista1 = c.colIndex;
                } else if (l.includes("paterno")) {
                  colPaterno1 = c.colIndex;
                } else if (l.includes("materno")) {
                  colMaterno1 = c.colIndex;
                } else if (l.includes("nombres") || l.includes("nombre")) {
                  colNombres1 = c.colIndex;
                } else if (l.includes("estudiante") || l.includes("alumno")) {
                  colEstudiante1 = c.colIndex;
                } else if (l.includes("run")) {
                  colRun1 = c.colIndex;
                } else if (l.includes("dv") || l.includes("dígito") || l.includes("verific")) {
                  colDv1 = c.colIndex;
                } else if (l.includes("promedio general") || l.includes("promediogeneral") || l.includes("promedio final") || l.includes("prom. gral")) {
                  colProm1 = c.colIndex;
                } else if (l.includes("periodo 1") || l.includes("periodo1") || l.includes("período 1")) {
                  colPeriodo1_1 = c.colIndex;
                } else if (l.includes("periodo 2") || l.includes("periodo2") || l.includes("período 2")) {
                  colPeriodo2_1 = c.colIndex;
                }
              });

              // Subjects matching with subheader P1, P2, PF lookup
              const subjectCols1: {
                subjectName: string;
                p1ColIndex: number;
                p2ColIndex: number;
                pfColIndex: number;
              }[] = [];
              headersRow1.forEach((h, cIdx) => {
                const label = String(h || "").trim();
                if (!label) return;

                const l = label.toLowerCase();
                const isStudentDetail = cIdx === colLista1 || cIdx === colEstudiante1 || cIdx === colPaterno1 || cIdx === colMaterno1 || cIdx === colNombres1 || cIdx === colRun1 || cIdx === colDv1;
                const isGeneralStat = l.includes("promedio") || l.includes("prom.") || l.includes("asistencia") || l.includes("situac") || l.includes("result") || l.includes("observ") || l.includes("periodo") || l.includes("año") || l.includes("tipo de enseñanza") || l.includes("nivel educativo") || l.includes("curso") || l.includes("días") || l.includes("asistido") || l.includes("ausente") || l.includes("justificado");

                if (!isStudentDetail && !isGeneralStat) {
                  let p1ColIndex = -1;
                  let p2ColIndex = -1;
                  let pfColIndex = cIdx;

                  // Look ahead to find subheaders P1, P2, PF in the next rows of same block
                  for (let k = cIdx; k < grid1[mainHeaderRowIdx1].length; k++) {
                    if (k > cIdx && String(grid1[mainHeaderRowIdx1][k] || "").trim() !== "") {
                      break; // next header block started
                    }
                    const sub = String(subheadersRow1[k] || "").trim().toLowerCase();
                    if (sub === "p1") {
                      p1ColIndex = k;
                    } else if (sub === "p2") {
                      p2ColIndex = k;
                    } else if (sub === "pf" || sub === "promedio final" || sub === "final" || sub === "promedio") {
                      pfColIndex = k;
                    }
                  }
                  subjectCols1.push({
                    subjectName: label,
                    p1ColIndex,
                    p2ColIndex,
                    pfColIndex,
                  });
                }
              });

              const parsedPanoramaStudents: CalificacionStudent[] = [];

              // Decide data starting row based on presence of subheaders
              let startRow1 = mainHeaderRowIdx1 + 1;
              const nextRow1 = grid1[mainHeaderRowIdx1 + 1] || [];
              const hasSubheaders = nextRow1.some(val => {
                const s = String(val || "").toLowerCase();
                return s === "p1" || s === "p2" || s === "pf" || s === "promedio final";
              });
              if (hasSubheaders) {
                startRow1 = mainHeaderRowIdx1 + 2;
              }

              for (let r = startRow1; r < grid1.length; r++) {
                const row = grid1[r];
                if (!row || row.length === 0) continue;

                let nameVal = "";
                if (colPaterno1 !== -1 || colMaterno1 !== -1 || colNombres1 !== -1) {
                  const pat = colPaterno1 !== -1 ? String(row[colPaterno1] || "").trim() : "";
                  const mat = colMaterno1 !== -1 ? String(row[colMaterno1] || "").trim() : "";
                  const nom = colNombres1 !== -1 ? String(row[colNombres1] || "").trim() : "";
                  nameVal = [pat, mat, nom].filter(Boolean).join(" ");
                } else if (colEstudiante1 !== -1) {
                  nameVal = String(row[colEstudiante1] || "").trim();
                }

                if (!nameVal) {
                  const isRowEmpty = row.every(cell => String(cell || "").trim() === "");
                  if (isRowEmpty) continue;
                  const firstVal = String(row[0] || "").trim().toLowerCase();
                  if (firstVal.startsWith("total") || firstVal.startsWith("promedio") || firstVal.startsWith("resumen")) {
                    break;
                  }
                  continue;
                }

                const nameLower = nameVal.toLowerCase();
                if (nameLower.startsWith("total") || nameLower.startsWith("promedio") || nameLower.startsWith("resumen")) {
                  break;
                }

                const run = colRun1 !== -1 ? String(row[colRun1] || "").trim() : "";
                const dv = colDv1 !== -1 ? String(row[colDv1] || "").trim() : "";
                const listNum = colLista1 !== -1 ? String(row[colLista1] || "").trim() : String(r - startRow1 + 1);

                const parseGrade = (val: unknown): number | null => {
                  if (val === undefined || val === null || String(val).trim() === "" || String(val).trim() === "-") {
                    return null;
                  }
                  const str = String(val).trim().replace(",", ".");
                  const num = parseFloat(str);
                  return isNaN(num) ? null : num;
                };

                const promedioGeneral = colProm1 !== -1 ? parseGrade(row[colProm1]) : null;
                const p1General = colPeriodo1_1 !== -1 ? parseGrade(row[colPeriodo1_1]) : null;
                const p2General = colPeriodo2_1 !== -1 ? parseGrade(row[colPeriodo2_1]) : null;

                const subjects: CalificacionSubjectData[] = [];
                subjectCols1.forEach(s => {
                  const p1Val = s.p1ColIndex !== -1 ? parseGrade(row[s.p1ColIndex]) : null;
                  const p2Val = s.p2ColIndex !== -1 ? parseGrade(row[s.p2ColIndex]) : null;
                  const pfVal = parseGrade(row[s.pfColIndex]);
                  subjects.push({
                    subjectName: s.subjectName,
                    grades: [],
                    p1: p1Val,
                    p2: p2Val,
                    pf: pfVal,
                  });
                });

                parsedPanoramaStudents.push({
                  id: `panorama-${file.name}-${nameVal}-${r}`,
                  lista: listNum,
                  estudiante: capitalizeProperName(nameVal),
                  run,
                  dv,
                  periodo1: p1General,
                  periodo2: p2General,
                  promedioGeneral,
                  subjects,
                });
              }

              const riesgoSheetName = workbook.SheetNames.find(
                name => name.toLowerCase().includes("riesgo") && name.toLowerCase().includes("promoci")
              );
              let parsedRiesgoRows: RiesgoStudent[] = [];
              let parsedRiesgoHeaders: string[] = [];

              if (riesgoSheetName) {
                const sheet = workbook.Sheets[riesgoSheetName];
                const grid2 = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });

                let mainHeaderRowIdx2 = 7;
                for (let r = 0; r < Math.min(8, grid2.length); r++) {
                  const row = grid2[r] || [];
                  const nonEmptyCount = row.filter(cell => String(cell || "").trim() !== "").length;
                  if (nonEmptyCount >= 4) {
                    const hasEst = row.some(val => {
                      const s = String(val || "").toLowerCase();
                      return s.includes("estudiante") || s.includes("alumno") || s.includes("nombre") || s.includes("apellido");
                    });
                    if (hasEst) {
                      mainHeaderRowIdx2 = r;
                      break;
                    }
                  }
                }

                const headersRow2 = grid2[mainHeaderRowIdx2] || [];
                const colInfo2 = headersRow2.map((h, cIdx) => ({
                  label: String(h || "").trim(),
                  colIndex: cIdx,
                })).filter(c => c.label !== "");

                parsedRiesgoHeaders = colInfo2.map(c => c.label);

                let colLista2 = -1;
                let colEstudiante2 = -1;
                let colPaterno2 = -1;
                let colMaterno2 = -1;
                let colNombres2 = -1;
                let colRun2 = -1;
                let colDv2 = -1;
                let colProm2 = -1;
                let colAsist2 = -1;

                colInfo2.forEach(c => {
                  const l = c.label.toLowerCase();
                  if (l.includes("lista") || l.includes("nº") || l.includes("n°")) {
                    colLista2 = c.colIndex;
                  } else if (l.includes("paterno")) {
                    colPaterno2 = c.colIndex;
                  } else if (l.includes("materno")) {
                    colMaterno2 = c.colIndex;
                  } else if (l.includes("nombres") || l.includes("nombre")) {
                    colNombres2 = c.colIndex;
                  } else if (l.includes("estudiante") || l.includes("alumno")) {
                    colEstudiante2 = c.colIndex;
                  } else if (l.includes("run")) {
                    colRun2 = c.colIndex;
                  } else if (l.includes("dv") || l.includes("dígito") || l.includes("verific")) {
                    colDv2 = c.colIndex;
                  } else if (l.includes("promedio") || l.includes("prom")) {
                    colProm2 = c.colIndex;
                  } else if (l.includes("asistencia") || l.includes("asist")) {
                    colAsist2 = c.colIndex;
                  }
                });

                const startRow2 = mainHeaderRowIdx2 + 1;
                for (let r = startRow2; r < grid2.length; r++) {
                  const row = grid2[r];
                  if (!row || row.length === 0) continue;

                  let nameVal = "";
                  if (colPaterno2 !== -1 || colMaterno2 !== -1 || colNombres2 !== -1) {
                    const pat = colPaterno2 !== -1 ? String(row[colPaterno2] || "").trim() : "";
                    const mat = colMaterno2 !== -1 ? String(row[colMaterno2] || "").trim() : "";
                    const nom = colNombres2 !== -1 ? String(row[colNombres2] || "").trim() : "";
                    nameVal = [pat, mat, nom].filter(Boolean).join(" ");
                  } else if (colEstudiante2 !== -1) {
                    nameVal = String(row[colEstudiante2] || "").trim();
                  }

                  if (!nameVal) {
                    const isRowEmpty = row.every(cell => String(cell || "").trim() === "");
                    if (isRowEmpty) continue;
                    const firstVal = String(row[0] || "").trim().toLowerCase();
                    if (firstVal.startsWith("total") || firstVal.startsWith("promedio") || firstVal.startsWith("resumen")) {
                      break;
                    }
                    continue;
                  }

                  const nameLower = nameVal.toLowerCase();
                  if (nameLower.startsWith("total") || nameLower.startsWith("promedio") || nameLower.startsWith("resumen")) {
                    break;
                  }

                  const run = colRun2 !== -1 ? String(row[colRun2] || "").trim() : "";
                  const dv = colDv2 !== -1 ? String(row[colDv2] || "").trim() : "";
                  const listNum = colLista2 !== -1 ? String(row[colLista2] || "").trim() : String(r - startRow2 + 1);

                  const parseGrade = (val: unknown): number | null => {
                    if (val === undefined || val === null || String(val).trim() === "" || String(val).trim() === "-") {
                      return null;
                    }
                    const str = String(val).trim().replace(",", ".");
                    const num = parseFloat(str);
                    return isNaN(num) ? null : num;
                  };

                  const promedioGeneral = colProm2 !== -1 ? parseGrade(row[colProm2]) : null;
                  const asistencia = colAsist2 !== -1 ? String(row[colAsist2] || "").trim() : null;

                  const dynamicFields: Record<string, string | number | null> = {};
                  colInfo2.forEach(c => {
                    dynamicFields[c.label] = String(row[c.colIndex] || "").trim();
                  });

                  parsedRiesgoRows.push({
                    id: `riesgo-${file.name}-${nameVal}-${r}`,
                    lista: listNum,
                    estudiante: capitalizeProperName(nameVal),
                    run,
                    dv,
                    promedioGeneral,
                    asistencia,
                    dynamicFields,
                  });
                }
              }

              if (parsedPanoramaStudents.length === 0) {
                resolve({ fileName: file.name, type: "panorama_riesgo", error: "No se encontraron estudiantes en el panorama global." });
              } else {
                resolve({
                  fileName: file.name,
                  type: "panorama_riesgo",
                  calificacionesData: parsedPanoramaStudents,
                  riesgoData: parsedRiesgoRows,
                  riesgoHeaders: parsedRiesgoHeaders,
                });
              }
            } else if (isCalificaciones) {
              const headersRow1 = jsonGrid[0] || [];
              const headersRow2 = jsonGrid[1] || [];
              const headersRow3 = jsonGrid[2] || [];

              let colLista = -1;
              let colEstudiante = -1;
              let colRun = -1;
              let colDv = -1;
              let colP1 = -1;
              let colP2 = -1;
              let colPF = -1;

              for (let c = 0; c < Math.min(15, headersRow1.length); c++) {
                const h1 = String(headersRow1[c] || "").trim().toLowerCase();
                if (h1.includes("lista") || h1.includes("nº") || h1.includes("n°")) {
                  colLista = c;
                } else if (h1.includes("estudiante") && !h1.includes("run")) {
                  colEstudiante = c;
                } else if (h1.includes("run")) {
                  colRun = c;
                } else if (h1.includes("digito") || h1.includes("dígito") || h1.includes("dv") || h1.includes("verific")) {
                  colDv = c;
                } else if (h1.includes("periodo 1") || h1.includes("periodo1")) {
                  colP1 = c;
                } else if (h1.includes("periodo 2") || h1.includes("periodo2")) {
                  colP2 = c;
                } else if (h1.includes("promedio general") || h1.includes("promediogeneral")) {
                  colPF = c;
                }
              }

              for (let c = 0; c < Math.min(15, headersRow2.length); c++) {
                const h2 = String(headersRow2[c] || "").trim().toLowerCase();
                if (h2 === "p1" && colP1 === -1 && c !== colLista && c !== colEstudiante && c !== colRun && c !== colDv) {
                  colP1 = c;
                } else if (h2 === "p2" && colP2 === -1 && c !== colLista && c !== colEstudiante && c !== colRun && c !== colDv) {
                  colP2 = c;
                } else if (h2 === "pf" && colPF === -1 && c !== colLista && c !== colEstudiante && c !== colRun && c !== colDv) {
                  colPF = c;
                }
              }

              interface SubjectColumnInfo {
                subjectName: string;
                colIndex: number;
                type: "grade" | "p1" | "p2" | "pf";
                gradeLabel?: string;
                weight?: string;
              }

              const subjectCols: SubjectColumnInfo[] = [];
              let currentSubject = "";

              for (let c = 0; c < headersRow1.length; c++) {
                if (c === colLista || c === colEstudiante || c === colRun || c === colDv || c === colP1 || c === colP2 || c === colPF) {
                  continue;
                }

                const h1Val = String(headersRow1[c] || "").trim();
                if (h1Val) {
                  currentSubject = h1Val;
                }

                if (!currentSubject) continue;

                const h2Val = String(headersRow2[c] || "").trim();
                const h3Val = String(headersRow3[c] || "").trim();

                const h2Lower = h2Val.toLowerCase();
                if (h2Lower === "p1") {
                  subjectCols.push({ subjectName: currentSubject, colIndex: c, type: "p1" });
                } else if (h2Lower === "p2") {
                  subjectCols.push({ subjectName: currentSubject, colIndex: c, type: "p2" });
                } else if (h2Lower === "pf") {
                  subjectCols.push({ subjectName: currentSubject, colIndex: c, type: "pf" });
                } else if (h2Val) {
                  subjectCols.push({
                    subjectName: currentSubject,
                    colIndex: c,
                    type: "grade",
                    gradeLabel: h2Val,
                    weight: h3Val,
                  });
                }
              }

              const studentRows: CalificacionStudent[] = [];
              let breakIndex = jsonGrid.length;

              for (let r = 3; r < jsonGrid.length; r++) {
                const row = jsonGrid[r];
                if (!row || row.length === 0) continue;

                // Stop parsing when encountering indicators, categories, or legend sections
                const valColA = String(row[0] || "").trim().toLowerCase();
                const valColEstudiante = colEstudiante !== -1 ? String(row[colEstudiante] || "").trim().toLowerCase() : "";
                if (
                  valColA.startsWith("indicador") || valColEstudiante.startsWith("indicador") ||
                  valColA.startsWith("categor") || valColEstudiante.startsWith("categor") ||
                  valColA.startsWith("leyenda") || valColEstudiante.startsWith("leyenda")
                ) {
                  breakIndex = r;
                  break;
                }

                const studentName = colEstudiante !== -1 ? String(row[colEstudiante] || "").trim() : "";
                if (!studentName) continue;

                const listNum = colLista !== -1 ? String(row[colLista] || "").trim() : String(r - 2);
                const run = colRun !== -1 ? String(row[colRun] || "").trim() : "";
                const dv = colDv !== -1 ? String(row[colDv] || "").trim() : "";

                const parseGrade = (val: unknown): number | null => {
                  if (val === undefined || val === null || String(val).trim() === "" || String(val).trim() === "-") {
                    return null;
                  }
                  const str = String(val).trim().replace(",", ".");
                  const num = parseFloat(str);
                  return isNaN(num) ? null : num;
                };

                const p1General = colP1 !== -1 ? parseGrade(row[colP1]) : null;
                const p2General = colP2 !== -1 ? parseGrade(row[colP2]) : null;
                const pfGeneral = colPF !== -1 ? parseGrade(row[colPF]) : null;

                const subjectsMap: Record<string, CalificacionSubjectData> = {};

                for (const sCol of subjectCols) {
                  const subName = sCol.subjectName;
                  if (!subjectsMap[subName]) {
                    subjectsMap[subName] = {
                      subjectName: subName,
                      grades: [],
                      p1: null,
                      p2: null,
                      pf: null,
                    };
                  }

                  const val = row[sCol.colIndex];
                  const gradeVal = parseGrade(val);

                  if (sCol.type === "grade") {
                    subjectsMap[subName].grades.push({
                      label: sCol.gradeLabel || "",
                      weight: sCol.weight || "",
                      value: gradeVal,
                    });
                  } else if (sCol.type === "p1") {
                    subjectsMap[subName].p1 = gradeVal;
                  } else if (sCol.type === "p2") {
                    subjectsMap[subName].p2 = gradeVal;
                  } else if (sCol.type === "pf") {
                    subjectsMap[subName].pf = gradeVal;
                  }
                }

                studentRows.push({
                  id: `calificacion-${file.name}-${studentName}-${r}`,
                  lista: listNum,
                  estudiante: capitalizeProperName(studentName),
                  run,
                  dv,
                  periodo1: p1General,
                  periodo2: p2General,
                  promedioGeneral: pfGeneral,
                  subjects: Object.values(subjectsMap),
                });
              }

              // Parse extra indicators, categories, and legends
              const excelExtra: ExcelExtraData = {
                indicadores: [],
                categorias: [],
                leyendas: [],
              };

              let currentSection: "indicadores" | "categorias" | "leyendas" | "" = "";

              for (let idx = breakIndex; idx < jsonGrid.length; idx++) {
                const row = jsonGrid[idx];
                if (!row || row.length === 0) continue;

                const valColA = String(row[0] || "").trim().toLowerCase();
                const valColEstudiante = colEstudiante !== -1 ? String(row[colEstudiante] || "").trim().toLowerCase() : "";

                if (valColA.startsWith("indicador") || valColEstudiante.startsWith("indicador")) {
                  currentSection = "indicadores";
                } else if (valColA.startsWith("categor") || valColEstudiante.startsWith("categor")) {
                  currentSection = "categorias";
                } else if (valColA.startsWith("leyenda") || valColEstudiante.startsWith("leyenda")) {
                  currentSection = "leyendas";
                }

                const colLabelIndex = colEstudiante !== -1 ? colEstudiante : 1;
                const label = String(row[colLabelIndex] || row[1] || row[0] || "").trim();

                if (currentSection === "indicadores" || currentSection === "categorias") {
                  if (!label || label.toLowerCase().startsWith("indicador") || label.toLowerCase().startsWith("categor")) {
                    continue;
                  }

                  const values: ExcelSectionValue[] = [];

                  if (colP1 !== -1 && row[colP1] !== undefined) {
                    const v = row[colP1];
                    const numVal = parseFloat(String(v || "").trim().replace(",", "."));
                    values.push({
                      type: "p1_gen",
                      value: isNaN(numVal) ? String(v || "").trim() : numVal,
                    });
                  }
                  if (colP2 !== -1 && row[colP2] !== undefined) {
                    const v = row[colP2];
                    const numVal = parseFloat(String(v || "").trim().replace(",", "."));
                    values.push({
                      type: "p2_gen",
                      value: isNaN(numVal) ? String(v || "").trim() : numVal,
                    });
                  }
                  if (colPF !== -1 && row[colPF] !== undefined) {
                    const v = row[colPF];
                    const numVal = parseFloat(String(v || "").trim().replace(",", "."));
                    values.push({
                      type: "pf_gen",
                      value: isNaN(numVal) ? String(v || "").trim() : numVal,
                    });
                  }

                  for (const sCol of subjectCols) {
                    const v = row[sCol.colIndex];
                    if (v !== undefined) {
                      const strVal = String(v || "").trim();
                      const numVal = parseFloat(strVal.replace(",", "."));
                      const finalVal = isNaN(numVal) ? strVal : numVal;

                      if (sCol.type === "grade") {
                        values.push({
                          subjectName: sCol.subjectName,
                          type: "grade",
                          label: sCol.gradeLabel,
                          value: finalVal,
                        });
                      } else {
                        values.push({
                          subjectName: sCol.subjectName,
                          type: sCol.type,
                          value: finalVal,
                        });
                      }
                    }
                  }

                  if (currentSection === "indicadores") {
                    excelExtra.indicadores.push({ label, values });
                  } else {
                    excelExtra.categorias.push({ label, values });
                  }
                } else if (currentSection === "leyendas") {
                  const text = String(row[1] || row[0] || "").trim();
                  if (text && text.toLowerCase() !== "leyenda" && !text.toLowerCase().startsWith("leyenda")) {
                    excelExtra.leyendas.push(text);
                  }
                }
              }

              if (studentRows.length === 0) {
                resolve({ fileName: file.name, type: "calificaciones", error: "No se encontraron estudiantes o calificaciones en el archivo." });
              } else {
                resolve({ fileName: file.name, type: "calificaciones", calificacionesData: studentRows, excelExtraData: excelExtra });
              }
            } else if (isPendientes) {
              const items: PendienteItem[] = [];

              const sheet1Name = workbook.SheetNames.find((name) =>
                name.toLowerCase().includes("firma")
              ) || workbook.SheetNames[0];

              if (sheet1Name) {
                const sheet1 = workbook.Sheets[sheet1Name];
                const data1 = XLSX.utils.sheet_to_json<Record<string, string>>(sheet1, {
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
                      id: `firma-${file.name}-${fecha}-${hora}-${asignatura}-${index}`,
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
                const data2 = XLSX.utils.sheet_to_json<Record<string, string>>(sheet2, {
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
                      id: `leccionario-${file.name}-${fecha}-${hora}-${asignatura}-${index}`,
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
                resolve({ fileName: file.name, type: "pendientes", error: "No se encontraron registros pendientes en el archivo." });
              } else {
                resolve({ fileName: file.name, type: "pendientes", pendientesData: items });
              }
            } else {
              const worksheet = workbook.Sheets[firstSheetName];
              const jsonData = XLSX.utils.sheet_to_json<Record<string, string>>(worksheet, {
                defval: "",
                raw: false,
              });

              if (jsonData.length === 0) {
                resolve({ fileName: file.name, type: "observations", error: "El archivo Excel está vacío." });
                return;
              }

              const headers = Object.keys(jsonData[0]);
              const missingHeaders = REQUIRED_HEADERS.filter((header) => !headers.includes(header));
              if (missingHeaders.length > 0) {
                resolve({ fileName: file.name, type: "observations", error: "Kimche Analyzer no puede procesar ese tipo de planilla aún" });
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
                    id: `${file.name}-${fechaTexto}-${nombreCompleto}-${normalizedRow["Tipo de observación"]?.trim() ?? ""}-${index}`,
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

              resolve({ fileName: file.name, type: "observations", observationsData: parsed });
            }
          } catch (error) {
            const msg = error instanceof Error ? error.message : String(error);
            resolve({ fileName: file.name, type: "observations", error: `No fue posible leer el archivo Excel: ${msg}` });
          }
        };
        reader.onerror = () => {
          resolve({ fileName: file.name, type: "observations", error: "Error al leer el archivo." });
        };
        reader.readAsArrayBuffer(file);
      } else {
        resolve({ fileName: file.name, type: "observations", error: "Por favor, sube un archivo CSV o Excel (.xlsx, .xls)." });
      }
    });
  }, []);

  const handleFiles = useCallback(async (files: File[]) => {
    if (files.length === 0) return;
    setErrorMessage("");
    setIsLoading(true);

    try {
      const results = await Promise.all(files.map(processSingleFile));

      const failed = results.find((r) => r.error);
      if (failed) {
        setErrorMessage(`Error en archivo ${failed.fileName}: ${failed.error}`);
        return;
      }

      const firstType = results[0].type;
      const allSameType = results.every((r) => r.type === firstType);
      if (!allSameType) {
        setErrorMessage("Todos los archivos cargados al mismo tiempo deben ser del mismo tipo (Firmas/Leccionarios, Observaciones, Calificaciones o Panorama y Riesgo)");
        return;
      }

      if (firstType === "observations") {
        const mergedObs: Observation[] = [];
        results.forEach((r) => {
          if (r.observationsData) {
            mergedObs.push(...r.observationsData);
          }
        });
        setUploadedFilesCount(files.length);
        setActiveCourseTab("all");
        setAppMode("observations");
        setObservations(mergedObs);
      } else if (firstType === "calificaciones") {
        const mergedCalificaciones: CalificacionStudent[] = [];
        const mergedExtraData: ExcelExtraData = { indicadores: [], categorias: [], leyendas: [] };
        results.forEach((r) => {
          if (r.calificacionesData) {
            mergedCalificaciones.push(...r.calificacionesData);
          }
          if (r.excelExtraData) {
            for (const item of r.excelExtraData.indicadores) {
              const existing = mergedExtraData.indicadores.find((ind) => ind.label === item.label);
              if (existing) {
                existing.values.push(...item.values);
              } else {
                mergedExtraData.indicadores.push({ label: item.label, values: [...item.values] });
              }
            }
            for (const item of r.excelExtraData.categorias) {
              const existing = mergedExtraData.categorias.find((cat) => cat.label === item.label);
              if (existing) {
                existing.values.push(...item.values);
              } else {
                mergedExtraData.categorias.push({ label: item.label, values: [...item.values] });
              }
            }
            mergedExtraData.leyendas.push(...r.excelExtraData.leyendas);
          }
        });

        mergedExtraData.leyendas = Array.from(new Set(mergedExtraData.leyendas));

        setAppMode("calificaciones");
        setCalificaciones(mergedCalificaciones);
        setExcelExtraData(mergedExtraData);
      } else if (firstType === "panorama_riesgo") {
        const mergedCalificaciones: CalificacionStudent[] = [];
        const mergedRiesgo: RiesgoStudent[] = [];
        let finalRiesgoHeaders: string[] = [];

        results.forEach((r) => {
          if (r.calificacionesData) {
            mergedCalificaciones.push(...r.calificacionesData);
          }
          if (r.riesgoData) {
            mergedRiesgo.push(...r.riesgoData);
          }
          if (r.riesgoHeaders && r.riesgoHeaders.length > 0) {
            finalRiesgoHeaders = r.riesgoHeaders;
          }
        });

        setAppMode("panorama_riesgo");
        setPanoramaActiveTab("panorama");
        setCalificaciones(mergedCalificaciones);
        setRiesgoStudents(mergedRiesgo);
        setRiesgoHeaders(finalRiesgoHeaders);
      } else {
        const mergedPendientes: PendienteItem[] = [];
        results.forEach((r) => {
          if (r.pendientesData) {
            mergedPendientes.push(...r.pendientesData);
          }
        });
        setAppMode("pendientes");
        setPendientes(mergedPendientes);
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      setErrorMessage(`Error procesando los archivos: ${msg}`);
    } finally {
      setIsLoading(false);
    }
  }, [processSingleFile]);

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
              handleFiles([file]);
            } else if (message.data instanceof Uint8Array || message.data instanceof ArrayBuffer) {
              const bytes = message.data instanceof Uint8Array ? message.data : new Uint8Array(message.data);
              const file = new File([bytes], fileName, {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              });
              handleFiles([file]);
            }
          } catch (err) {
            const msg = err instanceof Error ? err.message : String(err);
            setErrorMessage(`Error al procesar archivo de la extensión: ${msg}`);
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
  }, [handleFiles]);

  const handleInputChange = (event: ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files ? Array.from(event.target.files) : [];
    if (files.length === 0) return;
    handleFiles(files);
    event.target.value = "";
  };

  const handleDrop = (event: DragEvent<HTMLLabelElement>) => {
    event.preventDefault();
    setIsDragging(false);
    const files = event.dataTransfer.files ? Array.from(event.dataTransfer.files) : [];
    if (files.length === 0) return;
    handleFiles(files);
  };

  if (observations.length === 0 && pendientes.length === 0 && calificaciones.length === 0) {
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
                Carga las planillas que exportas desde Kimche y obtén estadísticas en un formato visual más amigable. El análisis se realiza de manera local, por lo que no se comparte ningún dato con nadie. La app puede procesar los siguientes tipos de planillas:
              </p>
            </div>
          </div>

          <div className="mb-6 flex flex-wrap gap-2.5 justify-center">
            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-indigo-50 text-indigo-700 border border-indigo-200 shadow-sm">
              <span className="h-1.5 w-1.5 rounded-full bg-indigo-500" />
              Observaciones
            </span>
            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200 shadow-sm">
              <span className="h-1.5 w-1.5 rounded-full bg-emerald-500" />
              Calificaciones
            </span>
            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-amber-50 text-amber-700 border border-amber-200 shadow-sm">
              <span className="h-1.5 w-1.5 rounded-full bg-amber-500" />
              Registros pendientes
            </span>
            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-violet-50 text-violet-700 border border-violet-200 shadow-sm">
              <span className="h-1.5 w-1.5 rounded-full bg-violet-500" />
              Promedios y situación final
            </span>
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
              <p className="text-sm text-slate-500">o haz clic para abrir el explorador de archivos</p>
            </div>
            <input
              type="file"
              accept=".csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              className="hidden"
              onChange={handleInputChange}
              multiple
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
              onClick={resetAll}
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
                    onChange={(e) => setPendientesFilter(e.target.value as "all" | "firma" | "leccionario")}
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

  if ((appMode === "calificaciones" || appMode === "panorama_riesgo") && calificaciones.length > 0) {
    const selectedSubObj = selectedCalificacionesAsignatura !== "all"
      ? calificaciones.find(s => s.subjects.some(sub => sub.subjectName === selectedCalificacionesAsignatura))
        ?.subjects.find(sub => sub.subjectName === selectedCalificacionesAsignatura)
      : null;
    const gradeHeaders = selectedSubObj ? selectedSubObj.grades : [];

    return (
      <main className="flex-grow w-full bg-slate-100 px-4 py-8 text-slate-900">
        {isLoading && (
          <div className="fixed inset-0 z-[100] flex flex-col items-center justify-center bg-slate-950/70 backdrop-blur-md select-none transition-all duration-300">
            <div className="relative flex items-center justify-center h-20 w-20">
              <div className="absolute inset-0 rounded-full border-4 border-transparent border-t-indigo-500 border-r-purple-500 animate-spin" />
              <div className="absolute inset-1.5 rounded-full border-4 border-transparent border-b-blue-500 border-l-pink-500 animate-spin [animation-duration:1.5s]" />
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
              <h1 className="text-2xl font-semibold">
                {appMode === "panorama_riesgo" ? "Visión del Curso y Riesgo de Promoción" : "Registro de Calificaciones y Rendimiento"}
              </h1>
              <p className="mt-1 text-sm text-slate-500">
                {appMode === "panorama_riesgo"
                  ? "Visualiza la visión general del curso y detecta a los estudiantes en peligro de repitencia."
                  : "Visualiza y analiza el rendimiento general, promedios, aprobaciones y calificaciones por asignatura. Todo de forma local."}
              </p>
            </div>
            <button
              onClick={resetAll}
              className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2.5 text-sm font-semibold text-slate-700 hover:bg-slate-100 transition active:scale-95 shadow-sm whitespace-nowrap"
            >
              Subir otro archivo
            </button>
          </header>

          {appMode === "panorama_riesgo" && (
            <div className="flex border-b border-slate-200 bg-white p-2 rounded-2xl shadow-sm gap-2 print:hidden">
              <button
                onClick={() => setPanoramaActiveTab("panorama")}
                className={`rounded-xl px-6 py-2.5 text-sm font-bold transition-all ${panoramaActiveTab === "panorama"
                  ? "bg-indigo-600 text-white shadow-sm"
                  : "text-slate-600 hover:text-slate-800 hover:bg-slate-50"
                  }`}
              >
                Panorama Global
              </button>
              <button
                onClick={() => setPanoramaActiveTab("riesgo")}
                className={`rounded-xl px-6 py-2.5 text-sm font-bold transition-all ${panoramaActiveTab === "riesgo"
                  ? "bg-indigo-600 text-white shadow-sm"
                  : "text-slate-600 hover:text-slate-800 hover:bg-slate-50"
                  }`}
              >
                Riesgo de Promoción ({riesgoStudents.length})
              </button>
            </div>
          )}

          {appMode !== "panorama_riesgo" || panoramaActiveTab === "panorama" ? (
            <>
              {/* Métricas del Curso */}
              <section className="grid gap-4 md:grid-cols-2 lg:grid-cols-4">
                <article className="relative overflow-hidden rounded-xl bg-indigo-600 px-4 py-4.5 text-white shadow-sm transition hover:shadow-md">
                  <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
                  <div className="relative z-10">
                    <p className="text-xs font-semibold uppercase tracking-wider text-indigo-100">Promedio General Curso</p>
                    <p className="text-3xl font-extrabold mt-1">{promedioGeneralCurso.toFixed(2)}</p>
                  </div>
                </article>

                <article className="relative overflow-hidden rounded-xl bg-emerald-600 px-4 py-4.5 text-white shadow-sm transition hover:shadow-md">
                  <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
                  <div className="relative z-10">
                    <p className="text-xs font-semibold uppercase tracking-wider text-emerald-105">Tasa de Aprobación</p>
                    <p className="text-3xl font-extrabold mt-1">{tasaAprobacion}%</p>
                  </div>
                </article>

                <article className="relative overflow-hidden rounded-xl bg-blue-600 px-4 py-4.5 text-white shadow-sm transition hover:shadow-md">
                  <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
                  <div className="relative z-10">
                    <p className="text-xs font-semibold uppercase tracking-wider text-blue-101">Total Alumnos</p>
                    <p className="text-3xl font-extrabold mt-1">{calificaciones.length}</p>
                  </div>
                </article>

                <article className="relative overflow-hidden rounded-xl bg-purple-600 px-4 py-4.5 text-white shadow-sm transition hover:shadow-md">
                  <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
                  <div className="relative z-10">
                    <p className="text-xs font-semibold uppercase tracking-wider text-purple-101">Asignaturas Evaluadas</p>
                    <p className="text-3xl font-extrabold mt-1">{uniqueAsignaturasCalificaciones.length}</p>
                  </div>
                </article>
              </section>

              {/* Gráficos */}
              <section className="grid gap-6 lg:grid-cols-2">
                <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                  <h3 className="font-semibold text-slate-800 mb-4">Distribución de Promedios Generales</h3>
                  <div className="w-full flex items-center justify-center">
                    <ResponsiveContainer width="100%" height={240}>
                      <BarChart data={distributionData}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                        <XAxis dataKey="name" stroke="#64748b" fontSize={11} tickLine={false} />
                        <YAxis stroke="#64748b" fontSize={11} tickLine={false} />
                        <Tooltip cursor={{ fill: '#f8fafc' }} />
                        <Bar dataKey="value" radius={[6, 6, 0, 0]}>
                          {distributionData.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={entry.color} />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                  <h3 className="font-semibold text-slate-800 mb-4">Rendimiento Promedio por Asignatura</h3>
                  <div className="w-full flex items-center justify-center">
                    {subjectPerformanceData.length > 0 ? (
                      <ResponsiveContainer width="100%" height={240}>
                        <BarChart data={subjectPerformanceData}>
                          <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                          <XAxis dataKey="name" stroke="#64748b" fontSize={11} tickLine={false} />
                          <YAxis domain={[1.0, 7.0]} stroke="#64748b" fontSize={11} tickLine={false} />
                          <Tooltip />
                          <Bar dataKey="promedio" fill="#6366f1" radius={[6, 6, 0, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    ) : (
                      <div className="h-[240px] flex items-center justify-center text-slate-400 italic text-sm">
                        No hay asignaturas con promedios válidos
                      </div>
                    )}
                  </div>
                </div>
              </section>

              {/* Filtros e Historial */}
              <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                <div className="mb-6 flex flex-col gap-4 border-b border-slate-100 pb-5">
                  <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                    <h3 className="font-semibold text-slate-800 flex items-center gap-2">
                      <Filter className="h-4 w-4 text-indigo-500" />
                      Panel de Filtros de Calificaciones
                    </h3>
                    {(selectedCalificacionesAsignatura !== "all" || onlyShowAtRiskCalificaciones || calificacionesSearch !== "") && (
                      <button
                        onClick={() => {
                          setSelectedCalificacionesAsignatura("all");
                          setOnlyShowAtRiskCalificaciones(false);
                          setCalificacionesSearch("");
                        }}
                        className="text-xs text-indigo-600 hover:text-indigo-800 font-semibold underline cursor-pointer"
                      >
                        Restablecer todos los filtros
                      </button>
                    )}
                  </div>

                  <div className="grid gap-4 sm:grid-cols-3">
                    {/* 1. Búsqueda por Estudiante */}
                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                        Buscar Alumno
                      </label>
                      <div className="relative">
                        <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                        <input
                          value={calificacionesSearch}
                          onChange={(e) => setCalificacionesSearch(e.target.value)}
                          placeholder="Nombre estudiante..."
                          className="w-full rounded-xl border border-slate-300 bg-white py-2 pl-9 pr-3 text-xs outline-none ring-indigo-500 transition focus:ring"
                        />
                      </div>
                    </div>

                    {/* 2. Filtro por Asignatura */}
                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                        Ver Asignatura
                      </label>
                      <select
                        value={selectedCalificacionesAsignatura}
                        onChange={(e) => setSelectedCalificacionesAsignatura(e.target.value)}
                        className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-xs outline-none ring-indigo-500 focus:ring cursor-pointer"
                      >
                        <option value="all">Promedios Generales (Todos)</option>
                        {uniqueAsignaturasCalificaciones.map((subj) => (
                          <option key={subj} value={subj}>
                            {subj}
                          </option>
                        ))}
                      </select>
                    </div>

                    {/* 3. Filtro Alumnos en Riesgo */}
                    <div className="flex flex-col justify-end pb-1.5">
                      <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-slate-700 select-none">
                        <input
                          type="checkbox"
                          checked={onlyShowAtRiskCalificaciones}
                          onChange={(e) => setOnlyShowAtRiskCalificaciones(e.target.checked)}
                          className="h-4 w-4 rounded border-slate-300 text-indigo-650 focus:ring-indigo-500"
                        />
                        <span>Mostrar solo estudiantes en riesgo (&lt; 4.0)</span>
                      </label>
                    </div>
                  </div>
                </div>

                {/* Tabla de Resultados */}
                <div className="overflow-x-auto">
                  <table className="min-w-full text-left text-sm border-collapse table-fixed">
                    <thead>
                      <tr className="border-b border-slate-200 text-slate-500 font-medium">
                        <th className="px-3 py-3 w-16 min-w-[64px] text-center font-bold">Nº</th>
                        <th className="px-4 py-3 w-[300px] min-w-[200px] font-bold">Estudiante</th>

                        {/* Dynamic Header Columns for subject-specific view */}
                        {selectedCalificacionesAsignatura !== "all" ? (
                          appMode === "panorama_riesgo" ? (
                            <>
                              <th className="px-3 py-3 text-center font-bold w-24 min-w-[96px]">P1</th>
                              <th className="px-3 py-3 text-center font-bold w-24 min-w-[96px]">P2</th>
                              <th className="px-3 py-3 text-center font-bold w-24 min-w-[96px]">PF</th>
                            </>
                          ) : (
                            <>
                              {gradeHeaders.map((grade, gIdx) => (
                                <th key={gIdx} className="px-2 py-3 text-center w-20 min-w-[80px]">
                                  <div className="font-bold">{grade.label}</div>
                                  {grade.weight && (
                                    <div className="text-[10px] text-slate-400 font-normal">{grade.weight}</div>
                                  )}
                                </th>
                              ))}
                              <th className="px-3 py-3 text-center font-bold w-24 min-w-[96px]">P1</th>
                              <th className="px-3 py-3 text-center font-bold w-24 min-w-[96px]">P2</th>
                              <th className="px-3 py-3 text-center font-bold w-24 min-w-[96px]">PF</th>
                            </>
                          )
                        ) : (
                          <>
                            <th className="px-4 py-3 text-center font-bold w-36 min-w-[144px]">P1 General</th>
                            <th className="px-4 py-3 text-center font-bold w-36 min-w-[144px]">P2 General</th>
                            <th className="px-4 py-3 text-center font-bold w-40 min-w-[160px]">Promedio General</th>
                          </>
                        )}

                        <th className="px-4 py-3 text-center font-bold w-36 min-w-[144px]">Estado</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredCalificaciones.length > 0 ? (
                        filteredCalificaciones.map((student) => {
                          let finalGrade: number | null = null;
                          let p1Value: number | null = null;
                          let p2Value: number | null = null;
                          let studentGrades: GradeItem[] = [];

                          if (selectedCalificacionesAsignatura !== "all") {
                            const studentSub = student.subjects.find(s => s.subjectName === selectedCalificacionesAsignatura);
                            if (studentSub) {
                              finalGrade = studentSub.pf;
                              p1Value = studentSub.p1;
                              p2Value = studentSub.p2;
                              studentGrades = studentSub.grades;
                            }
                          } else {
                            finalGrade = student.promedioGeneral;
                            p1Value = student.periodo1;
                            p2Value = student.periodo2;
                          }

                          const isApproved = finalGrade !== null && finalGrade >= 4.0;
                          const hasGrade = finalGrade !== null && finalGrade > 0;

                          const getGradeClass = (gradeVal: number | null) => {
                            if (gradeVal === null || gradeVal === 0) return "text-slate-400";
                            if (gradeVal < 4.0) return "text-rose-600 font-extrabold";
                            if (gradeVal >= 6.0) return "text-emerald-600 font-bold";
                            return "text-slate-800 font-semibold";
                          };

                          return (
                            <tr key={student.id} className="border-b border-slate-100 hover:bg-slate-50/50 transition">
                              <td className="px-3 py-3.5 text-center text-slate-500 whitespace-nowrap font-medium">
                                {student.lista}
                              </td>
                              <td className="px-4 py-3.5 font-bold text-slate-800 whitespace-nowrap">
                                {student.estudiante}
                              </td>

                              {selectedCalificacionesAsignatura !== "all" ? (
                                appMode === "panorama_riesgo" ? (
                                  <>
                                    <td className={`px-3 py-3.5 text-center whitespace-nowrap ${getGradeClass(p1Value)}`}>
                                      {p1Value !== null ? p1Value.toFixed(1).replace(".", ",") : "-"}
                                    </td>
                                    <td className={`px-3 py-3.5 text-center whitespace-nowrap ${getGradeClass(p2Value)}`}>
                                      {p2Value !== null ? p2Value.toFixed(1).replace(".", ",") : "-"}
                                    </td>
                                    <td className={`px-3 py-3.5 text-center whitespace-nowrap ${getGradeClass(finalGrade)} bg-slate-50/40`}>
                                      {finalGrade !== null ? finalGrade.toFixed(1).replace(".", ",") : "-"}
                                    </td>
                                  </>
                                ) : (
                                  <>
                                    {/* Render individual grades for subject */}
                                    {gradeHeaders.map((header, gIdx) => {
                                      const matchedGrade = studentGrades.find(g => g.label === header.label);
                                      const gValue = matchedGrade ? matchedGrade.value : null;
                                      return (
                                        <td key={gIdx} className={`px-2 py-3.5 text-center whitespace-nowrap ${getGradeClass(gValue)}`}>
                                          {gValue !== null ? gValue.toFixed(1).replace(".", ",") : "-"}
                                        </td>
                                      );
                                    })}

                                    <td className={`px-3 py-3.5 text-center whitespace-nowrap ${getGradeClass(p1Value)}`}>
                                      {p1Value !== null ? p1Value.toFixed(1).replace(".", ",") : "-"}
                                    </td>
                                    <td className={`px-3 py-3.5 text-center whitespace-nowrap ${getGradeClass(p2Value)}`}>
                                      {p2Value !== null ? p2Value.toFixed(1).replace(".", ",") : "-"}
                                    </td>
                                    <td className={`px-3 py-3.5 text-center whitespace-nowrap ${getGradeClass(finalGrade)} bg-slate-50/40`}>
                                      {finalGrade !== null ? finalGrade.toFixed(1).replace(".", ",") : "-"}
                                    </td>
                                  </>
                                )
                              ) : (
                                <>
                                  <td className={`px-4 py-3.5 text-center whitespace-nowrap ${getGradeClass(p1Value)}`}>
                                    {p1Value !== null ? p1Value.toFixed(1).replace(".", ",") : "-"}
                                  </td>
                                  <td className={`px-4 py-3.5 text-center whitespace-nowrap ${getGradeClass(p2Value)}`}>
                                    {p2Value !== null ? p2Value.toFixed(1).replace(".", ",") : "-"}
                                  </td>
                                  <td className={`px-4 py-3.5 text-center whitespace-nowrap ${getGradeClass(finalGrade)} bg-indigo-50/10`}>
                                    {finalGrade !== null ? finalGrade.toFixed(1).replace(".", ",") : "-"}
                                  </td>
                                </>
                              )}

                              <td className="px-4 py-3.5 text-center whitespace-nowrap">
                                {hasGrade ? (
                                  <span className={`inline-block rounded-full px-2.5 py-0.5 text-xs font-bold ${isApproved
                                    ? "bg-emerald-100 text-emerald-800"
                                    : "bg-rose-100 text-rose-800"
                                    }`}>
                                    {isApproved ? "Aprobado" : "Reprobado"}
                                  </span>
                                ) : (
                                  <span className="inline-block rounded-full px-2.5 py-0.5 text-xs font-medium bg-slate-100 text-slate-500">
                                    Sin Nota
                                  </span>
                                )}
                              </td>
                            </tr>
                          );
                        })
                      ) : (
                        <tr>
                          <td colSpan={selectedCalificacionesAsignatura !== "all" ? (appMode === "panorama_riesgo" ? 6 : 5 + gradeHeaders.length) : 6} className="text-center py-12 text-slate-400 italic">
                            No se encontraron estudiantes que coincidan con los filtros.
                          </td>
                        </tr>
                      )}
                    </tbody>
                    {subjectColumnStats && (
                      <tfoot className="border-t-2 border-slate-300 bg-slate-50 font-semibold text-slate-800">
                        {/* Row 1: Promedio */}
                        <tr>
                          <td className="px-3 py-2.5 text-center whitespace-nowrap text-slate-400 font-bold">-</td>
                          <td className="px-4 py-2.5 whitespace-nowrap font-bold text-slate-800">Promedio Curso</td>
                          {selectedCalificacionesAsignatura !== "all" ? (
                            appMode === "panorama_riesgo" ? (
                              <>
                                <td className="px-3 py-2.5 text-center whitespace-nowrap font-bold text-slate-800">
                                  {subjectColumnStats["P1"]?.promedio !== null ? subjectColumnStats["P1"]?.promedio?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2.5 text-center whitespace-nowrap font-bold text-slate-800">
                                  {subjectColumnStats["P2"]?.promedio !== null ? subjectColumnStats["P2"]?.promedio?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2.5 text-center whitespace-nowrap font-bold text-indigo-800 bg-indigo-50">
                                  {subjectColumnStats["PF"]?.promedio !== null ? subjectColumnStats["PF"]?.promedio?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                              </>
                            ) : (
                              <>
                                {gradeHeaders.map((header, gIdx) => {
                                  const stat = subjectColumnStats[header.label];
                                  const val = stat ? stat.promedio : null;
                                  return (
                                    <td key={`promedio-${gIdx}`} className="px-2 py-2.5 text-center whitespace-nowrap font-bold text-indigo-700">
                                      {val !== null ? val.toFixed(1).replace(".", ",") : "-"}
                                    </td>
                                  );
                                })}
                                <td className="px-3 py-2.5 text-center whitespace-nowrap font-bold text-slate-800">
                                  {subjectColumnStats["P1"]?.promedio !== null ? subjectColumnStats["P1"]?.promedio?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2.5 text-center whitespace-nowrap font-bold text-slate-800">
                                  {subjectColumnStats["P2"]?.promedio !== null ? subjectColumnStats["P2"]?.promedio?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2.5 text-center whitespace-nowrap font-bold text-indigo-800 bg-indigo-50">
                                  {subjectColumnStats["PF"]?.promedio !== null ? subjectColumnStats["PF"]?.promedio?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                              </>
                            )
                          ) : (
                            <>
                              <td className="px-4 py-2.5 text-center whitespace-nowrap font-bold text-slate-800">
                                {subjectColumnStats["P1 General"]?.promedio !== null ? subjectColumnStats["P1 General"]?.promedio?.toFixed(1).replace(".", ",") : "-"}
                              </td>
                              <td className="px-4 py-2.5 text-center whitespace-nowrap font-bold text-slate-800">
                                {subjectColumnStats["P2 General"]?.promedio !== null ? subjectColumnStats["P2 General"]?.promedio?.toFixed(1).replace(".", ",") : "-"}
                              </td>
                              <td className="px-4 py-2.5 text-center whitespace-nowrap font-bold text-indigo-800 bg-indigo-50">
                                {subjectColumnStats["Promedio General"]?.promedio !== null ? subjectColumnStats["Promedio General"]?.promedio?.toFixed(1).replace(".", ",") : "-"}
                              </td>
                            </>
                          )}
                          <td className="px-4 py-2.5 text-center whitespace-nowrap font-bold text-slate-400">-</td>
                        </tr>

                        {/* Row 2: Desviación Estándar */}
                        <tr className="text-slate-600 bg-slate-50/70 border-t border-slate-200">
                          <td className="px-3 py-2 text-center whitespace-nowrap text-slate-400 font-normal">-</td>
                          <td className="px-4 py-2 whitespace-nowrap font-medium text-slate-600 text-xs">Desviación Estándar</td>
                          {selectedCalificacionesAsignatura !== "all" ? (
                            appMode === "panorama_riesgo" ? (
                              <>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-medium text-slate-600 text-xs">
                                  {subjectColumnStats["P1"]?.desviacion !== null ? subjectColumnStats["P1"]?.desviacion?.toFixed(2).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-medium text-slate-600 text-xs">
                                  {subjectColumnStats["P2"]?.desviacion !== null ? subjectColumnStats["P2"]?.desviacion?.toFixed(2).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-semibold text-slate-700 bg-indigo-50/40 text-xs">
                                  {subjectColumnStats["PF"]?.desviacion !== null ? subjectColumnStats["PF"]?.desviacion?.toFixed(2).replace(".", ",") : "-"}
                                </td>
                              </>
                            ) : (
                              <>
                                {gradeHeaders.map((header, gIdx) => {
                                  const stat = subjectColumnStats[header.label];
                                  const val = stat ? stat.desviacion : null;
                                  return (
                                    <td key={`desv-${gIdx}`} className="px-2 py-2 text-center whitespace-nowrap font-medium text-slate-600 text-xs">
                                      {val !== null ? val.toFixed(2).replace(".", ",") : "-"}
                                    </td>
                                  );
                                })}
                                <td className="px-3 py-2 text-center whitespace-nowrap font-medium text-slate-600 text-xs">
                                  {subjectColumnStats["P1"]?.desviacion !== null ? subjectColumnStats["P1"]?.desviacion?.toFixed(2).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-medium text-slate-600 text-xs">
                                  {subjectColumnStats["P2"]?.desviacion !== null ? subjectColumnStats["P2"]?.desviacion?.toFixed(2).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-semibold text-slate-700 bg-indigo-50/40 text-xs">
                                  {subjectColumnStats["PF"]?.desviacion !== null ? subjectColumnStats["PF"]?.desviacion?.toFixed(2).replace(".", ",") : "-"}
                                </td>
                              </>
                            )
                          ) : (
                            <>
                              <td className="px-4 py-2 text-center whitespace-nowrap font-medium text-slate-600 text-xs">
                                {subjectColumnStats["P1 General"]?.desviacion !== null ? subjectColumnStats["P1 General"]?.desviacion?.toFixed(2).replace(".", ",") : "-"}
                              </td>
                              <td className="px-4 py-2 text-center whitespace-nowrap font-medium text-slate-600 text-xs">
                                {subjectColumnStats["P2 General"]?.desviacion !== null ? subjectColumnStats["P2 General"]?.desviacion?.toFixed(2).replace(".", ",") : "-"}
                              </td>
                              <td className="px-4 py-2 text-center whitespace-nowrap font-semibold text-slate-700 bg-indigo-50/40 text-xs">
                                {subjectColumnStats["Promedio General"]?.desviacion !== null ? subjectColumnStats["Promedio General"]?.desviacion?.toFixed(2).replace(".", ",") : "-"}
                              </td>
                            </>
                          )}
                          <td className="px-4 py-2 text-center whitespace-nowrap font-normal text-slate-400">-</td>
                        </tr>

                        {/* Row 3: Mínimo */}
                        <tr className="text-slate-600 bg-slate-50/70 border-t border-slate-200">
                          <td className="px-3 py-2 text-center whitespace-nowrap text-slate-400 font-normal">-</td>
                          <td className="px-4 py-2 whitespace-nowrap font-medium text-slate-600 text-xs">Nota Mínima</td>
                          {selectedCalificacionesAsignatura !== "all" ? (
                            appMode === "panorama_riesgo" ? (
                              <>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-medium text-rose-600 text-xs">
                                  {subjectColumnStats["P1"]?.minimo !== null ? subjectColumnStats["P1"]?.minimo?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-medium text-rose-600 text-xs">
                                  {subjectColumnStats["P2"]?.minimo !== null ? subjectColumnStats["P2"]?.minimo?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-semibold text-rose-700 bg-indigo-50/40 text-xs">
                                  {subjectColumnStats["PF"]?.minimo !== null ? subjectColumnStats["PF"]?.minimo?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                              </>
                            ) : (
                              <>
                                {gradeHeaders.map((header, gIdx) => {
                                  const stat = subjectColumnStats[header.label];
                                  const val = stat ? stat.minimo : null;
                                  return (
                                    <td key={`min-${gIdx}`} className="px-2 py-2 text-center whitespace-nowrap font-medium text-rose-600 text-xs">
                                      {val !== null ? val.toFixed(1).replace(".", ",") : "-"}
                                    </td>
                                  );
                                })}
                                <td className="px-3 py-2 text-center whitespace-nowrap font-medium text-rose-600 text-xs">
                                  {subjectColumnStats["P1"]?.minimo !== null ? subjectColumnStats["P1"]?.minimo?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-medium text-rose-600 text-xs">
                                  {subjectColumnStats["P2"]?.minimo !== null ? subjectColumnStats["P2"]?.minimo?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-semibold text-rose-700 bg-indigo-50/40 text-xs">
                                  {subjectColumnStats["PF"]?.minimo !== null ? subjectColumnStats["PF"]?.minimo?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                              </>
                            )
                          ) : (
                            <>
                              <td className="px-4 py-2 text-center whitespace-nowrap font-medium text-rose-600 text-xs">
                                {subjectColumnStats["P1 General"]?.minimo !== null ? subjectColumnStats["P1 General"]?.minimo?.toFixed(1).replace(".", ",") : "-"}
                              </td>
                              <td className="px-4 py-2 text-center whitespace-nowrap font-medium text-rose-600 text-xs">
                                {subjectColumnStats["P2 General"]?.minimo !== null ? subjectColumnStats["P2 General"]?.minimo?.toFixed(1).replace(".", ",") : "-"}
                              </td>
                              <td className="px-4 py-2 text-center whitespace-nowrap font-semibold text-rose-700 bg-indigo-50/40 text-xs">
                                {subjectColumnStats["Promedio General"]?.minimo !== null ? subjectColumnStats["Promedio General"]?.minimo?.toFixed(1).replace(".", ",") : "-"}
                              </td>
                            </>
                          )}
                          <td className="px-4 py-2 text-center whitespace-nowrap font-normal text-slate-400">-</td>
                        </tr>

                        {/* Row 4: Máximo */}
                        <tr className="text-slate-600 bg-slate-50/70 border-t border-slate-200">
                          <td className="px-3 py-2 text-center whitespace-nowrap text-slate-400 font-normal">-</td>
                          <td className="px-4 py-2 whitespace-nowrap font-medium text-slate-600 text-xs">Nota Máxima</td>
                          {selectedCalificacionesAsignatura !== "all" ? (
                            appMode === "panorama_riesgo" ? (
                              <>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-medium text-emerald-600 text-xs">
                                  {subjectColumnStats["P1"]?.maximo !== null ? subjectColumnStats["P1"]?.maximo?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-medium text-emerald-600 text-xs">
                                  {subjectColumnStats["P2"]?.maximo !== null ? subjectColumnStats["P2"]?.maximo?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-semibold text-emerald-700 bg-indigo-50/40 text-xs">
                                  {subjectColumnStats["PF"]?.maximo !== null ? subjectColumnStats["PF"]?.maximo?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                              </>
                            ) : (
                              <>
                                {gradeHeaders.map((header, gIdx) => {
                                  const stat = subjectColumnStats[header.label];
                                  const val = stat ? stat.maximo : null;
                                  return (
                                    <td key={`max-${gIdx}`} className="px-2 py-2 text-center whitespace-nowrap font-medium text-emerald-600 text-xs">
                                      {val !== null ? val.toFixed(1).replace(".", ",") : "-"}
                                    </td>
                                  );
                                })}
                                <td className="px-3 py-2 text-center whitespace-nowrap font-medium text-emerald-600 text-xs">
                                  {subjectColumnStats["P1"]?.maximo !== null ? subjectColumnStats["P1"]?.maximo?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-medium text-emerald-600 text-xs">
                                  {subjectColumnStats["P2"]?.maximo !== null ? subjectColumnStats["P2"]?.maximo?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                                <td className="px-3 py-2 text-center whitespace-nowrap font-semibold text-emerald-700 bg-indigo-50/40 text-xs">
                                  {subjectColumnStats["PF"]?.maximo !== null ? subjectColumnStats["PF"]?.maximo?.toFixed(1).replace(".", ",") : "-"}
                                </td>
                              </>
                            )
                          ) : (
                            <>
                              <td className="px-4 py-2 text-center whitespace-nowrap font-medium text-emerald-600 text-xs">
                                {subjectColumnStats["P1 General"]?.maximo !== null ? subjectColumnStats["P1 General"]?.maximo?.toFixed(1).replace(".", ",") : "-"}
                              </td>
                              <td className="px-4 py-2 text-center whitespace-nowrap font-medium text-emerald-600 text-xs">
                                {subjectColumnStats["P2 General"]?.maximo !== null ? subjectColumnStats["P2 General"]?.maximo?.toFixed(1).replace(".", ",") : "-"}
                              </td>
                              <td className="px-4 py-2 text-center whitespace-nowrap font-semibold text-emerald-700 bg-indigo-50/40 text-xs">
                                {subjectColumnStats["Promedio General"]?.maximo !== null ? subjectColumnStats["Promedio General"]?.maximo?.toFixed(1).replace(".", ",") : "-"}
                              </td>
                            </>
                          )}
                          <td className="px-4 py-2 text-center whitespace-nowrap font-normal text-slate-400">-</td>
                        </tr>
                      </tfoot>
                    )}
                  </table>
                </div>
              </section>

              {/* Card: Datos de Referencia y Leyendas */}
              {excelExtraData && (
                <section className="grid gap-6 md:grid-cols-2 print:hidden">
                  {/* Categorías Académicas */}
                  {excelExtraData.categorias.length > 0 && (
                    <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                      <h3 className="font-semibold text-slate-800 flex items-center gap-2 mb-4">
                        <Info className="h-5 w-5 text-indigo-500" />
                        Categorías Académicas (Reporte Excel)
                      </h3>
                      {(() => {
                        const categoriesToShow = excelExtraData.categorias.map(cat => {
                          const val = getExcelCategoryVal(
                            cat.label,
                            selectedCalificacionesAsignatura !== "all" ? "pf" : "pf_gen",
                            selectedCalificacionesAsignatura !== "all" ? selectedCalificacionesAsignatura : undefined
                          );
                          return {
                            label: cat.label,
                            value: val !== null ? Number(val) : 0,
                          };
                        }).filter(c => !isNaN(c.value) && c.value > 0);

                        if (categoriesToShow.length === 0) {
                          return (
                            <p className="text-xs text-slate-400 italic">
                              No hay datos de categorías registradas en Excel para esta asignatura/vista.
                            </p>
                          );
                        }

                        const maxVal = Math.max(...categoriesToShow.map(c => c.value), 1);

                        return (
                          <div className="space-y-4">
                            {categoriesToShow.map((cat, cIdx) => {
                              let barColor = "bg-indigo-500";
                              let bgLight = "bg-indigo-50";
                              let textColor = "text-indigo-700";
                              const lbl = cat.label.toLowerCase();
                              if (lbl.includes("insuficiente") || lbl.includes("bajo") || lbl.includes("reprobado")) {
                                barColor = "bg-rose-500";
                                bgLight = "bg-rose-50";
                                textColor = "text-rose-700";
                              } else if (lbl.includes("elemental") || lbl.includes("medio") || lbl.includes("regular")) {
                                barColor = "bg-amber-500";
                                bgLight = "bg-amber-50";
                                textColor = "text-amber-700";
                              } else if (lbl.includes("adecuado") || lbl.includes("bueno") || lbl.includes("aprobado")) {
                                barColor = "bg-emerald-500";
                                bgLight = "bg-emerald-50";
                                textColor = "text-emerald-700";
                              } else if (lbl.includes("destacado") || lbl.includes("excelente") || lbl.includes("alto")) {
                                barColor = "bg-purple-500";
                                bgLight = "bg-purple-50";
                                textColor = "text-purple-700";
                              }

                              const pct = (cat.value / maxVal) * 100;

                              return (
                                <div key={cIdx} className="space-y-1">
                                  <div className="flex items-center justify-between text-xs">
                                    <span className="font-medium text-slate-700">{cat.label}</span>
                                    <span className={`font-bold px-2 py-0.5 rounded-full text-[10px] ${bgLight} ${textColor}`}>
                                      {cat.value} alumnos
                                    </span>
                                  </div>
                                  <div className="h-2 w-full rounded-full bg-slate-100 overflow-hidden">
                                    <div
                                      className={`h-full rounded-full transition-all duration-500 ${barColor}`}
                                      style={{ width: `${pct}%` }}
                                    />
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        );
                      })()}
                    </div>
                  )}

                  {/* Leyenda de Siglas */}
                  {excelExtraData.leyendas.length > 0 && (
                    <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                      <h3 className="font-semibold text-slate-800 flex items-center gap-2 mb-3">
                        <BookOpen className="h-5 w-5 text-indigo-500" />
                        Leyenda de Siglas
                      </h3>
                      <div className="space-y-2">
                        {excelExtraData.leyendas.map((leyenda, lIdx) => {
                          const parts = leyenda.split(":");
                          const abbreviation = parts[0] ? parts[0].trim() : "";
                          const definition = parts[1] ? parts[1].trim() : "";
                          return (
                            <div key={lIdx} className="flex items-start gap-2 text-xs">
                              <span className="inline-flex items-center justify-center font-bold px-1.5 py-0.5 rounded bg-slate-100 text-slate-700 text-[10px] min-w-[32px] text-center">
                                {abbreviation}
                              </span>
                              <span className="text-slate-600 font-medium">{definition}</span>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                </section>
              )}
            </>
          ) : (
            <>
              {/* Dashboard y tabla de Riesgo de Promoción */}
              <section className="grid gap-4 md:grid-cols-2 lg:grid-cols-4">
                <article className="relative overflow-hidden rounded-xl bg-rose-600 px-4 py-4.5 text-white shadow-sm transition hover:shadow-md">
                  <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
                  <div className="relative z-10 flex flex-col justify-between h-full">
                    <p className="text-xs font-semibold uppercase tracking-wider text-rose-100">Alumnos en Riesgo</p>
                    <p className="text-3xl font-extrabold mt-1">{riesgoMetrics.totalRiesgo}</p>
                  </div>
                </article>

                <article className="relative overflow-hidden rounded-xl bg-amber-600 px-4 py-4.5 text-white shadow-sm transition hover:shadow-md">
                  <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
                  <div className="relative z-10 flex flex-col justify-between h-full">
                    <p className="text-xs font-semibold uppercase tracking-wider text-amber-100">Promedio de Notas Grupo</p>
                    <p className="text-3xl font-extrabold mt-1">{riesgoMetrics.avgGrade !== 0 ? riesgoMetrics.avgGrade.toFixed(2) : "-"}</p>
                  </div>
                </article>

                <article className="relative overflow-hidden rounded-xl bg-blue-600 px-4 py-4.5 text-white shadow-sm transition hover:shadow-md">
                  <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
                  <div className="relative z-10 flex flex-col justify-between h-full">
                    <p className="text-xs font-semibold uppercase tracking-wider text-blue-100">Asistencia Promedio</p>
                    <p className="text-3xl font-extrabold mt-1">{riesgoMetrics.avgAsist !== 0 ? riesgoMetrics.avgAsist + "%" : "-"}</p>
                  </div>
                </article>

                <article className="relative overflow-hidden rounded-xl bg-indigo-650 px-4 py-4.5 text-white bg-indigo-600 shadow-sm transition hover:shadow-md">
                  <div className="absolute inset-0 bg-[radial-gradient(#ffffff_1.5px,transparent_1.5px)] [background-size:12px_12px] opacity-15 pointer-events-none" />
                  <div className="relative z-10 flex flex-col justify-between h-full">
                    <p className="text-xs font-semibold uppercase tracking-wider text-indigo-100">Casos Muy Críticos</p>
                    <p className="text-3xl font-extrabold mt-1">{riesgoMetrics.countLowGrade + riesgoMetrics.countLowAsist}</p>
                  </div>
                </article>
              </section>

              {/* Leyenda Criterios de Riesgo de Repitencia */}
              <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm print:hidden">
                <div className="flex items-start gap-3 mb-4">
                  <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-rose-50 text-rose-600">
                    <Info className="h-5 w-5" />
                  </div>
                  <div>
                    <h3 className="font-bold text-slate-800 text-sm">
                      Criterios para el Riesgo de Repitencia
                    </h3>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Un estudiante se considera en riesgo de repitencia si cumple con al menos uno de los siguientes criterios:
                    </p>
                  </div>
                </div>

                <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                  <div className="flex flex-col rounded-xl bg-slate-50 p-4 border border-slate-100 hover:border-rose-100 hover:bg-rose-50/20 transition-all duration-300">
                    <span className="text-[10px] font-bold text-rose-600 bg-rose-50 px-2.5 py-0.5 rounded-full uppercase tracking-wider self-start">
                      Criterio 1
                    </span>
                    <p className="text-xs font-semibold text-slate-700 mt-2.5 leading-relaxed">
                      Reprobar tres o más asignaturas
                    </p>
                  </div>

                  <div className="flex flex-col rounded-xl bg-slate-50 p-4 border border-slate-100 hover:border-rose-100 hover:bg-rose-50/20 transition-all duration-300">
                    <span className="text-[10px] font-bold text-rose-600 bg-rose-50 px-2.5 py-0.5 rounded-full uppercase tracking-wider self-start">
                      Criterio 2
                    </span>
                    <p className="text-xs font-semibold text-slate-700 mt-2.5 leading-relaxed">
                      Un promedio insuficiente y promedio general menor a 4,5
                    </p>
                  </div>

                  <div className="flex flex-col rounded-xl bg-slate-50 p-4 border border-slate-100 hover:border-rose-100 hover:bg-rose-50/20 transition-all duration-300">
                    <span className="text-[10px] font-bold text-rose-600 bg-rose-50 px-2.5 py-0.5 rounded-full uppercase tracking-wider self-start">
                      Criterio 3
                    </span>
                    <p className="text-xs font-semibold text-slate-700 mt-2.5 leading-relaxed">
                      Dos promedios insuficientes y promedio general menor a 5,0
                    </p>
                  </div>

                  <div className="flex flex-col rounded-xl bg-slate-50 p-4 border border-slate-100 hover:border-rose-100 hover:bg-rose-50/20 transition-all duration-300">
                    <span className="text-[10px] font-bold text-rose-600 bg-rose-50 px-2.5 py-0.5 rounded-full uppercase tracking-wider self-start">
                      Criterio 4
                    </span>
                    <p className="text-xs font-semibold text-slate-700 mt-2.5 leading-relaxed">
                      Asistencia menor a 85%
                    </p>
                  </div>
                </div>
              </div>

              {/* Filtros de Riesgo */}
              <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                <div className="mb-6 flex flex-col gap-4 border-b border-slate-100 pb-5">
                  <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                    <h3 className="font-semibold text-slate-800 flex items-center gap-2">
                      <Filter className="h-4 w-4 text-rose-500" />
                      Filtros de Riesgo de Promoción
                    </h3>
                    {(riesgoSearch !== "" || riesgoFilter !== "all") && (
                      <button
                        onClick={() => {
                          setRiesgoSearch("");
                          setRiesgoFilter("all");
                        }}
                        className="text-xs text-indigo-600 hover:text-indigo-850 font-semibold underline cursor-pointer"
                      >
                        Restablecer todos los filtros
                      </button>
                    )}
                  </div>

                  <div className="grid gap-4 sm:grid-cols-2">
                    {/* Búsqueda */}
                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                        Buscar Estudiante
                      </label>
                      <div className="relative">
                        <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                        <input
                          value={riesgoSearch}
                          onChange={(e) => setRiesgoSearch(e.target.value)}
                          placeholder="Nombre estudiante..."
                          className="w-full rounded-xl border border-slate-300 bg-white py-2 pl-9 pr-3 text-xs outline-none ring-rose-500 transition focus:ring"
                        />
                      </div>
                    </div>

                    {/* Filtro de Criterio */}
                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">
                        Criterio de Riesgo
                      </label>
                      <select
                        value={riesgoFilter}
                        onChange={(e) => setRiesgoFilter(e.target.value as any)}
                        className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-xs outline-none ring-rose-500 focus:ring cursor-pointer"
                      >
                        <option value="all">Ver todos los casos</option>
                        <option value="average">Solo Promedio General &lt; 4.0</option>
                        <option value="attendance">Solo Asistencia &lt; 85%</option>
                      </select>
                    </div>
                  </div>
                </div>

                {/* Tabla de Resultados de Riesgo */}
                <div className="overflow-x-auto">
                  <table className="min-w-full text-left text-sm border-collapse">
                    <thead>
                      <tr className="border-b border-slate-200 text-slate-500 font-medium">
                        {visibleRiesgoHeaders.map((vh, idx) => (
                          <th key={idx} className="px-4 py-3 font-bold">{vh.display}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRiesgoStudents.length > 0 ? (
                        filteredRiesgoStudents.map((student) => (
                          <tr key={student.id} className="border-b border-slate-100 hover:bg-slate-50/50 transition">
                            {visibleRiesgoHeaders.map((vh, hIdx) => {
                              const header = vh.original;
                              const val = student.dynamicFields[header] ?? "-";

                              const isAvg = header.toLowerCase().includes("promedio") || header.toLowerCase().includes("prom");
                              const isAsist = header.toLowerCase().includes("asistencia") || header.toLowerCase().includes("asist");

                              let cellClass = "text-slate-800 font-medium";
                              let displayVal = String(val);

                              if (isAvg) {
                                const num = parseFloat(String(val).replace(",", "."));
                                if (!isNaN(num)) {
                                  cellClass = num < 4.0 ? "text-rose-600 font-extrabold" : num >= 6.0 ? "text-emerald-600 font-bold" : "text-slate-800 font-semibold";
                                  displayVal = num.toFixed(1).replace(".", ",");
                                }
                              } else if (isAsist) {
                                const num = parseFloat(String(val).replace("%", "").trim());
                                if (!isNaN(num)) {
                                  cellClass = num < 85 ? "text-rose-650 font-extrabold animate-pulse" : "text-slate-800 font-semibold";
                                  displayVal = `${Math.round(num)}%`;
                                }
                              }

                              const isName = vh.display.toLowerCase() === "estudiante" || header.toLowerCase().includes("estudiante") || header.toLowerCase().includes("alumno") || header.toLowerCase().includes("nombre");
                              if (isName) {
                                cellClass = "text-slate-900 font-bold";
                              }

                              return (
                                <td key={hIdx} className={`px-4 py-3.5 whitespace-nowrap ${cellClass}`}>
                                  {isName ? student.estudiante : displayVal}
                                </td>
                              );
                            })}
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan={visibleRiesgoHeaders.length || 1} className="text-center py-12 text-slate-400 italic">
                            No se encontraron estudiantes que coincidan con los filtros de riesgo.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </section>
            </>
          )}
        </div>
      </main>
    );
  }

  return (
    <>
      <main className="flex-grow w-full bg-slate-100 px-4 py-8 text-slate-900 print:hidden">
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
            <div className="flex items-center gap-3 print:hidden">
              <button
                onClick={resetAll}
                className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2.5 text-sm font-semibold text-slate-700 hover:bg-slate-100 transition active:scale-95 shadow-sm whitespace-nowrap"
              >
                Subir otro archivo
              </button>
              <button
                onClick={() => window.print()}
                className="rounded-xl bg-indigo-600 px-4 py-2.5 text-sm font-semibold text-white hover:bg-indigo-700 transition active:scale-95 shadow-sm whitespace-nowrap"
              >
                Generar PDF
              </button>
            </div>
          </header>

          {uploadedFilesCount > 1 && uniqueCursosObs.length > 0 && (
            <div className="flex flex-wrap gap-2 p-2 bg-white border border-slate-200 rounded-2xl shadow-sm print:hidden">
              <button
                onClick={() => setActiveCourseTab("all")}
                className={`rounded-xl px-4 py-2 text-xs font-semibold transition ${activeCourseTab === "all"
                  ? "bg-indigo-600 text-white shadow-sm"
                  : "text-slate-600 hover:text-slate-800 hover:bg-slate-50"
                  }`}
              >
                General
              </button>
              {uniqueCursosObs.map((curso) => (
                <button
                  key={curso}
                  onClick={() => setActiveCourseTab(curso)}
                  className={`rounded-xl px-4 py-2 text-xs font-semibold transition ${activeCourseTab === curso
                    ? "bg-indigo-600 text-white shadow-sm"
                    : "text-slate-600 hover:text-slate-800 hover:bg-slate-50"
                    }`}
                >
                  {curso}
                </button>
              ))}
            </div>
          )}

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
                      &quot;{latestPositive.descripcion}&quot;
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
                      &quot;{latestNegative.descripcion}&quot;
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

          <section className="grid gap-4 lg:grid-cols-2">
            {/* Tarjeta 1: Funcionarios con más positivas */}
            <article className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <h2 className="mb-4 text-lg font-semibold text-slate-800 flex items-center gap-2">
                <CheckCircle2 className="h-5 w-5 text-emerald-500" />
                Top 5 funcionarios con más positivas
              </h2>
              <div className="overflow-x-auto">
                <table className="min-w-full text-left text-sm">
                  <thead>
                    <tr className="border-b border-slate-200 text-slate-500 font-medium">
                      <th className="px-4 py-3 text-center w-16">Puesto</th>
                      <th className="px-4 py-3">Funcionario / Docente</th>
                      <th className="px-4 py-3 text-center w-28">Anotaciones</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Array.from({ length: 5 }).map((_, index) => {
                      const pos = topPositiveFuncionarios[index];
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
                            {pos ? pos.funcionario : <span className="text-slate-400 font-normal">-</span>}
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

            {/* Tarjeta 2: Funcionarios con más negativas */}
            <article className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <h2 className="mb-4 text-lg font-semibold text-slate-800 flex items-center gap-2">
                <AlertCircle className="h-5 w-5 text-rose-500" />
                Top 5 funcionarios con más negativas
              </h2>
              <div className="overflow-x-auto">
                <table className="min-w-full text-left text-sm">
                  <thead>
                    <tr className="border-b border-slate-200 text-slate-500 font-medium">
                      <th className="px-4 py-3 text-center w-16">Puesto</th>
                      <th className="px-4 py-3">Funcionario / Docente</th>
                      <th className="px-4 py-3 text-center w-28">Anotaciones</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Array.from({ length: 5 }).map((_, index) => {
                      const neg = topNegativeFuncionarios[index];
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
                            {neg ? neg.funcionario : <span className="text-slate-400 font-normal">-</span>}
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

          <section className="grid gap-4 print:hidden">
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
                          No se detectó información de faltas en la planilla cargada (se requieren columnas como &quot;Falta&quot; o &quot;Gravedad&quot;).
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </article>
          </section>

          <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm print:hidden">
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
                    onChange={(e) => setQuickFilter(e.target.value as QuickFilter)}
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
                    {uniqueAsignaturasObs.map((item) => (
                      <option key={item.name} value={item.name}>
                        {item.name} ({item.count})
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

      {/* Reporte de impresión limpio y compacto */}
      <div className="hidden print:block bg-white text-black font-sans p-2 text-xs leading-normal">
        {appMode === "calificaciones" || (appMode === "panorama_riesgo" && panoramaActiveTab === "panorama") ? (
          <>
            <header className="border-b border-slate-400 pb-2 mb-3 flex justify-between items-end">
              <div>
                <h1 className="text-lg font-bold text-slate-900">
                  {appMode === "panorama_riesgo" ? "Reporte de Panorama Global del Curso" : "Reporte de Calificaciones y Rendimiento"}
                </h1>
                <p className="text-[10px] text-slate-500">
                  {appMode === "panorama_riesgo" ? "Kimche Analyzer - Visión General de Asignaturas" : "Kimche Analyzer - Análisis de Rendimiento Escolar"}
                </p>
              </div>
              <div className="text-right text-[9px] text-slate-600">
                <p><strong>Fecha Emisión:</strong> {new Date().toLocaleDateString("es-CL", { day: "numeric", month: "long", year: "numeric" })}</p>
                <p><strong>Asignatura Activa:</strong> {selectedCalificacionesAsignatura === "all" ? "Todas las asignaturas" : selectedCalificacionesAsignatura}</p>
              </div>
            </header>

            {/* Métricas clave */}
            <div className="grid grid-cols-4 gap-3 mb-3">
              <div className="border border-slate-300 p-1.5 text-center rounded">
                <p className="text-[8px] font-semibold text-slate-500 uppercase">Promedio General</p>
                <p className="text-base font-bold text-slate-900 mt-0.5">{promedioGeneralCurso.toFixed(2)}</p>
              </div>
              <div className="border border-slate-300 p-1.5 text-center rounded">
                <p className="text-[8px] font-semibold text-slate-500 uppercase">Tasa de Aprobación</p>
                <p className="text-base font-bold text-slate-800 mt-0.5">{tasaAprobacion}%</p>
              </div>
              <div className="border border-slate-300 p-1.5 text-center rounded">
                <p className="text-[8px] font-semibold text-slate-500 uppercase">Total Alumnos</p>
                <p className="text-base font-bold text-slate-800 mt-0.5">{calificaciones.length}</p>
              </div>
              <div className="border border-slate-300 p-1.5 text-center rounded">
                <p className="text-[8px] font-semibold text-slate-500 uppercase">Asignaturas</p>
                <p className="text-base font-bold text-slate-800 mt-0.5">{uniqueAsignaturasCalificaciones.length}</p>
              </div>
            </div>

            {/* Tabla de Alumnos */}
            <div>
              <h3 className="text-[10px] font-bold text-slate-800 border-b border-slate-400 pb-0.5 mb-1">Listado de Calificaciones</h3>
              <table className="w-full text-[9px] border border-slate-300 border-collapse">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-300">
                    <th className="py-0.5 px-1 text-center font-semibold w-8">Nº</th>
                    <th className="py-0.5 px-1 text-left font-semibold">Estudiante</th>
                    {selectedCalificacionesAsignatura !== "all" ? (
                      appMode === "panorama_riesgo" ? (
                        <th className="py-0.5 px-1 text-center font-semibold" style={{ minWidth: '40px' }}>{selectedCalificacionesAsignatura}</th>
                      ) : (
                        <>
                          {calificaciones.find(s => s.subjects.some(sub => sub.subjectName === selectedCalificacionesAsignatura))
                            ?.subjects.find(sub => sub.subjectName === selectedCalificacionesAsignatura)
                            ?.grades.map((grade, gIdx) => (
                              <th key={gIdx} className="py-0.5 px-1 text-center font-semibold" style={{ minWidth: '24px' }}>{grade.label}</th>
                            ))
                          }
                          <th className="py-0.5 px-1 text-center font-semibold w-10">P1</th>
                          <th className="py-0.5 px-1 text-center font-semibold w-10">P2</th>
                          <th className="py-0.5 px-1 text-center font-semibold w-10">PF</th>
                        </>
                      )
                    ) : (
                      <>
                        <th className="py-0.5 px-1 text-center font-semibold w-16">P1 Gral</th>
                        <th className="py-0.5 px-1 text-center font-semibold w-16">P2 Gral</th>
                        <th className="py-0.5 px-1 text-center font-semibold w-20">Promedio Gral</th>
                      </>
                    )}
                    <th className="py-0.5 px-1 text-center font-semibold w-16">Estado</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredCalificaciones.map((student) => {
                    let finalGrade: number | null = null;
                    let p1Value: number | null = null;
                    let p2Value: number | null = null;
                    let studentGrades: GradeItem[] = [];

                    if (selectedCalificacionesAsignatura !== "all") {
                      const studentSub = student.subjects.find(s => s.subjectName === selectedCalificacionesAsignatura);
                      if (studentSub) {
                        finalGrade = studentSub.pf;
                        p1Value = studentSub.p1;
                        p2Value = studentSub.p2;
                        studentGrades = studentSub.grades;
                      }
                    } else {
                      finalGrade = student.promedioGeneral;
                      p1Value = student.periodo1;
                      p2Value = student.periodo2;
                    }

                    const isApproved = finalGrade !== null && finalGrade >= 4.0;
                    const hasGrade = finalGrade !== null && finalGrade > 0;

                    return (
                      <tr key={student.id} className="border-b border-slate-200">
                        <td className="py-0.5 px-1 text-center">{student.lista}</td>
                        <td className="py-0.5 px-1 font-medium">{student.estudiante}</td>

                        {selectedCalificacionesAsignatura !== "all" ? (
                          appMode === "panorama_riesgo" ? (
                            <td className={`py-0.5 px-1 text-center ${finalGrade !== null && finalGrade < 4.0 ? "text-rose-600 font-bold" : ""}`}>
                              {finalGrade !== null ? finalGrade.toFixed(1).replace(".", ",") : "-"}
                            </td>
                          ) : (
                            <>
                              {studentGrades.map((g, gIdx) => (
                                <td key={gIdx} className={`py-0.5 px-1 text-center ${g.value !== null && g.value < 4.0 ? "text-rose-600 font-bold" : ""}`}>
                                  {g.value !== null ? g.value.toFixed(1).replace(".", ",") : "-"}
                                </td>
                              ))}
                              <td className={`py-0.5 px-1 text-center ${p1Value !== null && p1Value < 4.0 ? "text-rose-600 font-bold" : ""}`}>
                                {p1Value !== null ? p1Value.toFixed(1).replace(".", ",") : "-"}
                              </td>
                              <td className={`py-0.5 px-1 text-center ${p2Value !== null && p2Value < 4.0 ? "text-rose-600 font-bold" : ""}`}>
                                {p2Value !== null ? p2Value.toFixed(1).replace(".", ",") : "-"}
                              </td>
                              <td className={`py-0.5 px-1 text-center font-bold ${finalGrade !== null && finalGrade < 4.0 ? "text-rose-600 font-bold" : ""}`}>
                                {finalGrade !== null ? finalGrade.toFixed(1).replace(".", ",") : "-"}
                              </td>
                            </>
                          )
                        ) : (
                          <>
                            <td className={`py-0.5 px-1 text-center ${p1Value !== null && p1Value < 4.0 ? "text-rose-600 font-bold" : ""}`}>
                              {p1Value !== null ? p1Value.toFixed(1).replace(".", ",") : "-"}
                            </td>
                            <td className={`py-0.5 px-1 text-center ${p2Value !== null && p2Value < 4.0 ? "text-rose-600 font-bold" : ""}`}>
                              {p2Value !== null ? p2Value.toFixed(1).replace(".", ",") : "-"}
                            </td>
                            <td className={`py-0.5 px-1 text-center font-bold ${finalGrade !== null && finalGrade < 4.0 ? "text-rose-600 font-bold" : ""}`}>
                              {finalGrade !== null ? finalGrade.toFixed(1).replace(".", ",") : "-"}
                            </td>
                          </>
                        )}

                        <td className="py-0.5 px-1 text-center font-semibold">
                          {hasGrade ? (isApproved ? "Aprobado" : "Reprobado") : "Sin Nota"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </>
        ) : appMode === "panorama_riesgo" && panoramaActiveTab === "riesgo" ? (
          <>
            <header className="border-b border-slate-400 pb-2 mb-3 flex justify-between items-end">
              <div>
                <h1 className="text-lg font-bold text-slate-900">Reporte de Alumnos en Riesgo de Promoción</h1>
                <p className="text-[10px] text-slate-500">Kimche Analyzer - Análisis de Repitencia y Asistencia</p>
              </div>
              <div className="text-right text-[9px] text-slate-600">
                <p><strong>Fecha Emisión:</strong> {new Date().toLocaleDateString("es-CL", { day: "numeric", month: "long", year: "numeric" })}</p>
              </div>
            </header>

            {/* Filtros aplicados */}
            {(riesgoSearch !== "" || riesgoFilter !== "all") && (
              <div className="mb-3 bg-slate-50 border border-slate-200 p-1.5 rounded text-[9px] text-slate-700">
                <p className="font-semibold mb-0.5">Filtros aplicados en la consulta:</p>
                <div className="grid grid-cols-2 gap-x-2 gap-y-0.5">
                  {riesgoSearch !== "" && <p>• <strong>Búsqueda:</strong> &quot;{riesgoSearch}&quot;</p>}
                  {riesgoFilter !== "all" && <p>• <strong>Criterio:</strong> {riesgoFilter === "average" ? "Promedio < 4.0" : "Asistencia < 85%"}</p>}
                </div>
              </div>
            )}

            {/* Métricas clave */}
            <div className="grid grid-cols-4 gap-3 mb-3">
              <div className="border border-slate-300 p-1.5 text-center rounded">
                <p className="text-[8px] font-semibold text-slate-500 uppercase">Alumnos en Riesgo</p>
                <p className="text-base font-bold text-slate-900 mt-0.5">{riesgoMetrics.totalRiesgo}</p>
              </div>
              <div className="border border-slate-300 p-1.5 text-center rounded">
                <p className="text-[8px] font-semibold text-slate-500 uppercase">Promedio Notas Grupo</p>
                <p className="text-base font-bold text-slate-800 mt-0.5">{riesgoMetrics.avgGrade !== 0 ? riesgoMetrics.avgGrade.toFixed(2) : "-"}</p>
              </div>
              <div className="border border-slate-300 p-1.5 text-center rounded">
                <p className="text-[8px] font-semibold text-slate-500 uppercase">Asistencia Promedio</p>
                <p className="text-base font-bold text-slate-800 mt-0.5">{riesgoMetrics.avgAsist !== 0 ? riesgoMetrics.avgAsist + "%" : "-"}</p>
              </div>
              <div className="border border-slate-300 p-1.5 text-center rounded">
                <p className="text-[8px] font-semibold text-slate-500 uppercase">Casos Críticos</p>
                <p className="text-base font-bold text-slate-800 mt-0.5">{riesgoMetrics.countLowGrade + riesgoMetrics.countLowAsist}</p>
              </div>
            </div>

            {/* Tabla de Riesgo */}
            <div>
              <h3 className="text-[10px] font-bold text-slate-800 border-b border-slate-400 pb-0.5 mb-1">Listado de Estudiantes en Peligro de Repitencia</h3>
              <table className="w-full text-[9px] border border-slate-300 border-collapse">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-300">
                    {visibleRiesgoHeaders.map((vh, idx) => (
                      <th key={idx} className="py-0.5 px-1 text-left font-semibold">{vh.display}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredRiesgoStudents.map((student) => (
                    <tr key={student.id} className="border-b border-slate-200">
                      {visibleRiesgoHeaders.map((vh, hIdx) => {
                        const header = vh.original;
                        const val = student.dynamicFields[header] ?? "-";
                        const isAvg = header.toLowerCase().includes("promedio") || header.toLowerCase().includes("prom");
                        const isAsist = header.toLowerCase().includes("asistencia") || header.toLowerCase().includes("asist");

                        let cellClass = "";
                        let displayVal = String(val);

                        if (isAvg) {
                          const num = parseFloat(String(val).replace(",", "."));
                          if (!isNaN(num)) {
                            if (num < 4.0) {
                              cellClass = "text-rose-600 font-bold";
                            }
                            displayVal = num.toFixed(1).replace(".", ",");
                          }
                        } else if (isAsist) {
                          const num = parseFloat(String(val).replace("%", "").trim());
                          if (!isNaN(num)) {
                            if (num < 85) {
                              cellClass = "text-rose-600 font-bold";
                            }
                            displayVal = `${Math.round(num)}%`;
                          }
                        }

                        const isName = vh.display.toLowerCase() === "estudiante" || header.toLowerCase().includes("estudiante") || header.toLowerCase().includes("alumno") || header.toLowerCase().includes("nombre");

                        return (
                          <td key={hIdx} className={`py-0.5 px-1 ${cellClass} ${isName ? "font-bold" : ""}`}>
                            {isName ? student.estudiante : displayVal}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        ) : (
          <>
            <header className="border-b border-slate-400 pb-2 mb-3 flex justify-between items-end">
              <div>
                <h1 className="text-lg font-bold text-slate-900">Reporte de Convivencia Escolar</h1>
                <p className="text-[10px] text-slate-500">Kimche Analyzer - Análisis de Observaciones</p>
              </div>
              <div className="text-right text-[9px] text-slate-650">
                <p><strong>Fecha Emisión:</strong> {new Date().toLocaleDateString("es-CL", { day: "numeric", month: "long", year: "numeric" })}</p>
                <p><strong>Curso Activo:</strong> {activeCourseTab === "all" ? "Todos los cursos" : activeCourseTab}</p>
              </div>
            </header>

            {/* Filtros aplicados */}
            {(selectedCursoObs !== "all" || selectedAsignaturaObs !== "all" || selectedFechaObs !== "all" || searchQuery !== "" || quickFilter !== "all") && (
              <div className="mb-3 bg-slate-50 border border-slate-200 p-1.5 rounded text-[9px] text-slate-700">
                <p className="font-semibold mb-0.5">Filtros aplicados en la consulta:</p>
                <div className="grid grid-cols-3 gap-x-2 gap-y-0.5">
                  {selectedCursoObs !== "all" && <p>• <strong>Curso:</strong> {selectedCursoObs}</p>}
                  {selectedAsignaturaObs !== "all" && <p>• <strong>Gravedad/Asig:</strong> {selectedAsignaturaObs}</p>}
                  {selectedFechaObs !== "all" && <p>• <strong>Fecha:</strong> {formatDateToVerbal(selectedFechaObs)}</p>}
                  {quickFilter !== "all" && <p>• <strong>Tipo:</strong> {quickFilter === "positiva" ? "Solo positivas" : quickFilter === "negativa" ? "Solo negativas" : "Otros"}</p>}
                  {searchQuery !== "" && <p>• <strong>Búsqueda:</strong> &quot;{searchQuery}&quot;</p>}
                </div>
              </div>
            )}

            {/* Métricas clave */}
            <div className="grid grid-cols-4 gap-3 mb-3">
              <div className="border border-slate-300 p-1.5 text-center rounded">
                <p className="text-[8px] font-semibold text-slate-500 uppercase">Obs. Procesadas</p>
                <p className="text-base font-bold text-slate-900 mt-0.5">{total}</p>
              </div>
              <div className="border border-slate-300 p-1.5 text-center rounded">
                <p className="text-[8px] font-semibold text-slate-500 uppercase">Anot. Positivas</p>
                <p className="text-base font-bold text-slate-800 mt-0.5">{summary.positivas}</p>
              </div>
              <div className="border border-slate-300 p-1.5 text-center rounded">
                <p className="text-[8px] font-semibold text-slate-500 uppercase">Anot. Negativas</p>
                <p className="text-base font-bold text-slate-800 mt-0.5">{summary.negativas}</p>
              </div>
              <div className="border border-slate-300 p-1.5 text-center rounded">
                <p className="text-[8px] font-semibold text-slate-500 uppercase">Ratio Convivencia</p>
                <p className="text-base font-bold text-slate-800 mt-0.5">{positiveObservationsPercentage}%</p>
              </div>
            </div>

            {/* Últimas anotaciones */}
            <div className="grid grid-cols-2 gap-3 mb-3 print-avoid-break">
              <div className="border border-slate-300 p-2 rounded">
                <p className="font-bold text-[9px] text-slate-700 uppercase mb-0.5">Última Anotación Positiva</p>
                {latestPositive ? (
                  <div>
                    <p className="font-bold text-[10px] text-slate-900">{latestPositive.nombreCompleto}</p>
                    <p className="text-[8px] text-slate-500 mb-0.5">{latestPositive.fechaTexto}</p>
                    <p className="text-[9px] text-slate-750 italic">&quot;{latestPositive.descripcion}&quot;</p>
                  </div>
                ) : (
                  <p className="text-[9px] text-slate-400 italic">No se encontraron anotaciones positivas.</p>
                )}
              </div>
              <div className="border border-slate-300 p-2 rounded">
                <p className="font-bold text-[9px] text-slate-700 uppercase mb-0.5">Última Anotación Negativa</p>
                {latestNegative ? (
                  <div>
                    <p className="font-bold text-[10px] text-slate-900">{latestNegative.nombreCompleto}</p>
                    <p className="text-[8px] text-slate-500 mb-0.5">{latestNegative.fechaTexto}</p>
                    <p className="text-[9px] text-slate-750 italic">&quot;{latestNegative.descripcion}&quot;</p>
                  </div>
                ) : (
                  <p className="text-[9px] text-slate-400 italic">No se encontraron anotaciones negativas.</p>
                )}
              </div>
            </div>

            {/* Tops de Estudiantes */}
            <div className="grid grid-cols-2 gap-3 mb-3 print-avoid-break">
              <div>
                <h3 className="text-[10px] font-bold text-slate-800 border-b border-slate-400 pb-0.5 mb-1">Top 5 Estudiantes (Positivas)</h3>
                <table className="w-full text-[9px] border border-slate-300 border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-300">
                      <th className="py-0.5 px-1 text-center font-semibold w-6">#</th>
                      <th className="py-0.5 px-1 text-left font-semibold">Estudiante</th>
                      <th className="py-0.5 px-1 text-center font-semibold w-10">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Array.from({ length: 5 }).map((_, index) => {
                      const item = topPositive[index];
                      return (
                        <tr key={index} className={index === 4 ? "" : "border-b border-slate-200"}>
                          <td className="py-0.5 px-1 text-center">{index + 1}</td>
                          <td className="py-0.5 px-1 font-medium">{item ? item.estudiante : "-"}</td>
                          <td className="py-0.5 px-1 text-center font-bold">{item ? item.totalObservaciones : "-"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <div>
                <h3 className="text-[10px] font-bold text-slate-800 border-b border-slate-400 pb-0.5 mb-1">Top 5 Estudiantes (Negativas)</h3>
                <table className="w-full text-[9px] border border-slate-300 border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-300">
                      <th className="py-0.5 px-1 text-center font-semibold w-6">#</th>
                      <th className="py-0.5 px-1 text-left font-semibold">Estudiante</th>
                      <th className="py-0.5 px-1 text-center font-semibold w-10">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Array.from({ length: 5 }).map((_, index) => {
                      const item = topNegative[index];
                      return (
                        <tr key={index} className={index === 4 ? "" : "border-b border-slate-200"}>
                          <td className="py-0.5 px-1 text-center">{index + 1}</td>
                          <td className="py-0.5 px-1 font-medium">{item ? item.estudiante : "-"}</td>
                          <td className="py-0.5 px-1 text-center font-bold">{item ? item.totalObservaciones : "-"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Tops de Funcionarios */}
            <div className="grid grid-cols-2 gap-3 mb-3 print-avoid-break">
              <div>
                <h3 className="text-[10px] font-bold text-slate-800 border-b border-slate-400 pb-0.5 mb-1">Top 5 Funcionarios (Positivas)</h3>
                <table className="w-full text-[9px] border border-slate-300 border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-300">
                      <th className="py-0.5 px-1 text-center font-semibold w-6">#</th>
                      <th className="py-0.5 px-1 text-left font-semibold">Funcionario</th>
                      <th className="py-0.5 px-1 text-center font-semibold w-10">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Array.from({ length: 5 }).map((_, index) => {
                      const item = topPositiveFuncionarios[index];
                      return (
                        <tr key={index} className={index === 4 ? "" : "border-b border-slate-200"}>
                          <td className="py-0.5 px-1 text-center">{index + 1}</td>
                          <td className="py-0.5 px-1 font-medium">{item ? item.funcionario : "-"}</td>
                          <td className="py-0.5 px-1 text-center font-bold">{item ? item.totalObservaciones : "-"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <div>
                <h3 className="text-[10px] font-bold text-slate-800 border-b border-slate-400 pb-0.5 mb-1">Top 5 Funcionarios (Negativas)</h3>
                <table className="w-full text-[9px] border border-slate-300 border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-300">
                      <th className="py-0.5 px-1 text-center font-semibold w-6">#</th>
                      <th className="py-0.5 px-1 text-left font-semibold">Funcionario</th>
                      <th className="py-0.5 px-1 text-center font-semibold w-10">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Array.from({ length: 5 }).map((_, index) => {
                      const item = topNegativeFuncionarios[index];
                      return (
                        <tr key={index} className={index === 4 ? "" : "border-b border-slate-200"}>
                          <td className="py-0.5 px-1 text-center">{index + 1}</td>
                          <td className="py-0.5 px-1 font-medium">{item ? item.funcionario : "-"}</td>
                          <td className="py-0.5 px-1 text-center font-bold">{item ? item.totalObservaciones : "-"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}
      </div>
    </>
  );
}
