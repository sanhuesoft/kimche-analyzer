import Papa from "papaparse";
import * as XLSX from "xlsx";
import type {
  Observation,
  PendienteItem,
  CalificacionStudent,
  CalificacionSubjectData,
  ExcelExtraData,
  ExcelSectionValue,
  RiesgoStudent,
} from "@/types";
import {
  REQUIRED_HEADERS,
  parseDateToSortable,
  normalizeType,
  capitalizeProperName,
  buildStudentName,
} from "@/utils/helpers";

export type ParseResult = {
  fileName: string;
  type: "observations" | "pendientes" | "calificaciones" | "panorama_riesgo";
  observationsData?: Observation[];
  pendientesData?: PendienteItem[];
  calificacionesData?: CalificacionStudent[];
  excelExtraData?: ExcelExtraData;
  riesgoData?: RiesgoStudent[];
  riesgoHeaders?: string[];
  error?: string;
};

export function processSingleFile(file: File): Promise<ParseResult> {
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
            resolve(parsePanoramaRiesgo(file, workbook, firstSheetName));
          } else if (isCalificaciones) {
            resolve(parseCalificaciones(file, jsonGrid));
          } else if (isPendientes) {
            resolve(parsePendientes(file, workbook));
          } else {
            resolve(parseObservationsXlsx(file, workbook, firstSheetName));
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
}

function parseObservationsXlsx(file: File, workbook: XLSX.WorkBook, firstSheetName: string): ParseResult {
  const worksheet = workbook.Sheets[firstSheetName];
  const jsonData = XLSX.utils.sheet_to_json<Record<string, string>>(worksheet, { defval: "", raw: false });

  if (jsonData.length === 0) {
    return { fileName: file.name, type: "observations", error: "El archivo Excel está vacío." };
  }

  const headers = Object.keys(jsonData[0]);
  const missingHeaders = REQUIRED_HEADERS.filter((header) => !headers.includes(header));
  if (missingHeaders.length > 0) {
    return { fileName: file.name, type: "observations", error: "Kimche Analyzer no puede procesar ese tipo de planilla aún" };
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

  return { fileName: file.name, type: "observations", observationsData: parsed };
}

function parsePendientes(file: File, workbook: XLSX.WorkBook): ParseResult {
  const items: PendienteItem[] = [];

  const sheet1Name = workbook.SheetNames.find((name) =>
    name.toLowerCase().includes("firma")
  ) || workbook.SheetNames[0];

  if (sheet1Name) {
    const sheet1 = workbook.Sheets[sheet1Name];
    const data1 = XLSX.utils.sheet_to_json<Record<string, string>>(sheet1, { defval: "", raw: false });
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
    const data2 = XLSX.utils.sheet_to_json<Record<string, string>>(sheet2, { defval: "", raw: false });
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
    return { fileName: file.name, type: "pendientes", error: "No se encontraron registros pendientes en el archivo." };
  }
  return { fileName: file.name, type: "pendientes", pendientesData: items };
}

function parseCalificaciones(file: File, jsonGrid: unknown[][]): ParseResult {
  const parseGrade = (val: unknown): number | null => {
    if (val === undefined || val === null || String(val).trim() === "" || String(val).trim() === "-") return null;
    const str = String(val).trim().replace(",", ".");
    const num = parseFloat(str);
    return isNaN(num) ? null : num;
  };

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
    if (h1.includes("lista") || h1.includes("nº") || h1.includes("n°")) colLista = c;
    else if (h1.includes("estudiante") && !h1.includes("run")) colEstudiante = c;
    else if (h1.includes("run")) colRun = c;
    else if (h1.includes("digito") || h1.includes("dígito") || h1.includes("dv") || h1.includes("verific")) colDv = c;
    else if (h1.includes("periodo 1") || h1.includes("periodo1")) colP1 = c;
    else if (h1.includes("periodo 2") || h1.includes("periodo2")) colP2 = c;
    else if (h1.includes("promedio general") || h1.includes("promediogeneral")) colPF = c;
  }

  for (let c = 0; c < Math.min(15, headersRow2.length); c++) {
    const h2 = String(headersRow2[c] || "").trim().toLowerCase();
    if (h2 === "p1" && colP1 === -1 && c !== colLista && c !== colEstudiante && c !== colRun && c !== colDv) colP1 = c;
    else if (h2 === "p2" && colP2 === -1 && c !== colLista && c !== colEstudiante && c !== colRun && c !== colDv) colP2 = c;
    else if (h2 === "pf" && colPF === -1 && c !== colLista && c !== colEstudiante && c !== colRun && c !== colDv) colPF = c;
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
    if (c === colLista || c === colEstudiante || c === colRun || c === colDv || c === colP1 || c === colP2 || c === colPF) continue;

    const h1Val = String(headersRow1[c] || "").trim();
    if (h1Val) currentSubject = h1Val;
    if (!currentSubject) continue;

    const h2Val = String(headersRow2[c] || "").trim();
    const h3Val = String(headersRow3[c] || "").trim();
    const h2Lower = h2Val.toLowerCase();

    if (h2Lower === "p1") subjectCols.push({ subjectName: currentSubject, colIndex: c, type: "p1" });
    else if (h2Lower === "p2") subjectCols.push({ subjectName: currentSubject, colIndex: c, type: "p2" });
    else if (h2Lower === "pf") subjectCols.push({ subjectName: currentSubject, colIndex: c, type: "pf" });
    else if (h2Val) subjectCols.push({ subjectName: currentSubject, colIndex: c, type: "grade", gradeLabel: h2Val, weight: h3Val });
  }

  const studentRows: CalificacionStudent[] = [];
  let breakIndex = jsonGrid.length;

  for (let r = 3; r < jsonGrid.length; r++) {
    const row = jsonGrid[r];
    if (!row || row.length === 0) continue;

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

    const p1General = colP1 !== -1 ? parseGrade(row[colP1]) : null;
    const p2General = colP2 !== -1 ? parseGrade(row[colP2]) : null;
    const pfGeneral = colPF !== -1 ? parseGrade(row[colPF]) : null;

    const subjectsMap: Record<string, CalificacionSubjectData> = {};

    for (const sCol of subjectCols) {
      const subName = sCol.subjectName;
      if (!subjectsMap[subName]) {
        subjectsMap[subName] = { subjectName: subName, grades: [], p1: null, p2: null, pf: null };
      }

      const val = row[sCol.colIndex];
      const gradeVal = parseGrade(val);

      if (sCol.type === "grade") {
        subjectsMap[subName].grades.push({ label: sCol.gradeLabel || "", weight: sCol.weight || "", value: gradeVal });
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
  const excelExtra: ExcelExtraData = { indicadores: [], categorias: [], leyendas: [] };
  let currentSection: "indicadores" | "categorias" | "leyendas" | "" = "";

  for (let idx = breakIndex; idx < jsonGrid.length; idx++) {
    const row = jsonGrid[idx];
    if (!row || row.length === 0) continue;

    const valColA = String(row[0] || "").trim().toLowerCase();
    const valColEstudiante = colEstudiante !== -1 ? String(row[colEstudiante] || "").trim().toLowerCase() : "";

    if (valColA.startsWith("indicador") || valColEstudiante.startsWith("indicador")) currentSection = "indicadores";
    else if (valColA.startsWith("categor") || valColEstudiante.startsWith("categor")) currentSection = "categorias";
    else if (valColA.startsWith("leyenda") || valColEstudiante.startsWith("leyenda")) currentSection = "leyendas";

    const colLabelIndex = colEstudiante !== -1 ? colEstudiante : 1;
    const label = String(row[colLabelIndex] || row[1] || row[0] || "").trim();

    if (currentSection === "indicadores" || currentSection === "categorias") {
      if (!label || label.toLowerCase().startsWith("indicador") || label.toLowerCase().startsWith("categor")) continue;

      const values: ExcelSectionValue[] = [];

      if (colP1 !== -1 && row[colP1] !== undefined) {
        const v = row[colP1];
        const numVal = parseFloat(String(v || "").trim().replace(",", "."));
        values.push({ type: "p1_gen", value: isNaN(numVal) ? String(v || "").trim() : numVal });
      }
      if (colP2 !== -1 && row[colP2] !== undefined) {
        const v = row[colP2];
        const numVal = parseFloat(String(v || "").trim().replace(",", "."));
        values.push({ type: "p2_gen", value: isNaN(numVal) ? String(v || "").trim() : numVal });
      }
      if (colPF !== -1 && row[colPF] !== undefined) {
        const v = row[colPF];
        const numVal = parseFloat(String(v || "").trim().replace(",", "."));
        values.push({ type: "pf_gen", value: isNaN(numVal) ? String(v || "").trim() : numVal });
      }

      for (const sCol of subjectCols) {
        const v = row[sCol.colIndex];
        if (v !== undefined) {
          const strVal = String(v || "").trim();
          const numVal = parseFloat(strVal.replace(",", "."));
          const finalVal = isNaN(numVal) ? strVal : numVal;

          if (sCol.type === "grade") {
            values.push({ subjectName: sCol.subjectName, type: "grade", label: sCol.gradeLabel, value: finalVal });
          } else {
            values.push({ subjectName: sCol.subjectName, type: sCol.type, value: finalVal });
          }
        }
      }

      if (currentSection === "indicadores") excelExtra.indicadores.push({ label, values });
      else excelExtra.categorias.push({ label, values });
    } else if (currentSection === "leyendas") {
      const text = String(row[1] || row[0] || "").trim();
      if (text && text.toLowerCase() !== "leyenda" && !text.toLowerCase().startsWith("leyenda")) {
        excelExtra.leyendas.push(text);
      }
    }
  }

  if (studentRows.length === 0) {
    return { fileName: file.name, type: "calificaciones", error: "No se encontraron estudiantes o calificaciones en el archivo." };
  }
  return { fileName: file.name, type: "calificaciones", calificacionesData: studentRows, excelExtraData: excelExtra };
}

function parsePanoramaRiesgo(file: File, workbook: XLSX.WorkBook, firstSheetName: string): ParseResult {
  const parseGrade = (val: unknown): number | null => {
    if (val === undefined || val === null || String(val).trim() === "" || String(val).trim() === "-") return null;
    const str = String(val).trim().replace(",", ".");
    const num = parseFloat(str);
    return isNaN(num) ? null : num;
  };

  const firstSheet = workbook.Sheets[firstSheetName];
  const grid1 = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, { header: 1, defval: "" });

  let mainHeaderRowIdx1 = 4;
  for (let r = 0; r < Math.min(5, grid1.length); r++) {
    const row = grid1[r] || [];
    const nonEmptyCount = row.filter(cell => String(cell || "").trim() !== "").length;
    if (nonEmptyCount >= 4) {
      const hasEst = row.some(val => {
        const s = String(val || "").toLowerCase();
        return s.includes("estudiante") || s.includes("alumno") || s.includes("nombre") || s.includes("apellido");
      });
      if (hasEst) { mainHeaderRowIdx1 = r; break; }
    }
  }

  const headersRow1 = grid1[mainHeaderRowIdx1] || [];
  const subheadersRow1 = grid1[mainHeaderRowIdx1 + 1] || [];

  const colInfo1 = headersRow1.map((h, cIdx) => ({
    label: String(h || "").trim(),
    colIndex: cIdx,
  })).filter(c => c.label !== "");

  let colLista1 = -1, colEstudiante1 = -1, colPaterno1 = -1, colMaterno1 = -1;
  let colNombres1 = -1, colRun1 = -1, colDv1 = -1, colProm1 = -1;
  let colPeriodo1_1 = -1, colPeriodo2_1 = -1;

  colInfo1.forEach(c => {
    const l = c.label.toLowerCase();
    if (l.includes("lista") || l.includes("nº") || l.includes("n°")) colLista1 = c.colIndex;
    else if (l.includes("paterno")) colPaterno1 = c.colIndex;
    else if (l.includes("materno")) colMaterno1 = c.colIndex;
    else if (l.includes("nombres") || l.includes("nombre")) colNombres1 = c.colIndex;
    else if (l.includes("estudiante") || l.includes("alumno")) colEstudiante1 = c.colIndex;
    else if (l.includes("run")) colRun1 = c.colIndex;
    else if (l.includes("dv") || l.includes("dígito") || l.includes("verific")) colDv1 = c.colIndex;
    else if (l.includes("promedio general") || l.includes("promediogeneral") || l.includes("promedio final") || l.includes("prom. gral")) colProm1 = c.colIndex;
    else if (l.includes("periodo 1") || l.includes("periodo1") || l.includes("período 1")) colPeriodo1_1 = c.colIndex;
    else if (l.includes("periodo 2") || l.includes("periodo2") || l.includes("período 2")) colPeriodo2_1 = c.colIndex;
  });

  const subjectCols1: { subjectName: string; p1ColIndex: number; p2ColIndex: number; pfColIndex: number; }[] = [];
  headersRow1.forEach((h, cIdx) => {
    const label = String(h || "").trim();
    if (!label) return;
    const l = label.toLowerCase();
    const isStudentDetail = cIdx === colLista1 || cIdx === colEstudiante1 || cIdx === colPaterno1 || cIdx === colMaterno1 || cIdx === colNombres1 || cIdx === colRun1 || cIdx === colDv1;
    const isGeneralStat = l.includes("promedio") || l.includes("prom.") || l.includes("asistencia") || l.includes("situac") || l.includes("result") || l.includes("observ") || l.includes("periodo") || l.includes("año") || l.includes("tipo de enseñanza") || l.includes("nivel educativo") || l.includes("curso") || l.includes("días") || l.includes("asistido") || l.includes("ausente") || l.includes("justificado");

    if (!isStudentDetail && !isGeneralStat) {
      let p1ColIndex = -1, p2ColIndex = -1, pfColIndex = cIdx;
      for (let k = cIdx; k < grid1[mainHeaderRowIdx1].length; k++) {
        if (k > cIdx && String(grid1[mainHeaderRowIdx1][k] || "").trim() !== "") break;
        const sub = String(subheadersRow1[k] || "").trim().toLowerCase();
        if (sub === "p1") p1ColIndex = k;
        else if (sub === "p2") p2ColIndex = k;
        else if (sub === "pf" || sub === "promedio final" || sub === "final" || sub === "promedio") pfColIndex = k;
      }
      subjectCols1.push({ subjectName: label, p1ColIndex, p2ColIndex, pfColIndex });
    }
  });

  const parsedPanoramaStudents: CalificacionStudent[] = [];
  let startRow1 = mainHeaderRowIdx1 + 1;
  const nextRow1 = grid1[mainHeaderRowIdx1 + 1] || [];
  const hasSubheaders = nextRow1.some(val => {
    const s = String(val || "").toLowerCase();
    return s === "p1" || s === "p2" || s === "pf" || s === "promedio final";
  });
  if (hasSubheaders) startRow1 = mainHeaderRowIdx1 + 2;

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
      if (firstVal.startsWith("total") || firstVal.startsWith("promedio") || firstVal.startsWith("resumen")) break;
      continue;
    }

    const nameLower = nameVal.toLowerCase();
    if (nameLower.startsWith("total") || nameLower.startsWith("promedio") || nameLower.startsWith("resumen")) break;

    const run = colRun1 !== -1 ? String(row[colRun1] || "").trim() : "";
    const dv = colDv1 !== -1 ? String(row[colDv1] || "").trim() : "";
    const listNum = colLista1 !== -1 ? String(row[colLista1] || "").trim() : String(r - startRow1 + 1);

    const promedioGeneral = colProm1 !== -1 ? parseGrade(row[colProm1]) : null;
    const p1General = colPeriodo1_1 !== -1 ? parseGrade(row[colPeriodo1_1]) : null;
    const p2General = colPeriodo2_1 !== -1 ? parseGrade(row[colPeriodo2_1]) : null;

    const subjects: CalificacionSubjectData[] = subjectCols1.map(s => ({
      subjectName: s.subjectName,
      grades: [],
      p1: s.p1ColIndex !== -1 ? parseGrade(row[s.p1ColIndex]) : null,
      p2: s.p2ColIndex !== -1 ? parseGrade(row[s.p2ColIndex]) : null,
      pf: parseGrade(row[s.pfColIndex]),
    }));

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

  // Parse riesgo sheet
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
        if (hasEst) { mainHeaderRowIdx2 = r; break; }
      }
    }

    const headersRow2 = grid2[mainHeaderRowIdx2] || [];
    const colInfo2 = headersRow2.map((h, cIdx) => ({
      label: String(h || "").trim(),
      colIndex: cIdx,
    })).filter(c => c.label !== "");

    parsedRiesgoHeaders = colInfo2.map(c => c.label);

    let colLista2 = -1, colEstudiante2 = -1, colPaterno2 = -1, colMaterno2 = -1;
    let colNombres2 = -1, colRun2 = -1, colDv2 = -1, colProm2 = -1, colAsist2 = -1;

    colInfo2.forEach(c => {
      const l = c.label.toLowerCase();
      if (l.includes("lista") || l.includes("nº") || l.includes("n°")) colLista2 = c.colIndex;
      else if (l.includes("paterno")) colPaterno2 = c.colIndex;
      else if (l.includes("materno")) colMaterno2 = c.colIndex;
      else if (l.includes("nombres") || l.includes("nombre")) colNombres2 = c.colIndex;
      else if (l.includes("estudiante") || l.includes("alumno")) colEstudiante2 = c.colIndex;
      else if (l.includes("run")) colRun2 = c.colIndex;
      else if (l.includes("dv") || l.includes("dígito") || l.includes("verific")) colDv2 = c.colIndex;
      else if (l.includes("promedio") || l.includes("prom")) colProm2 = c.colIndex;
      else if (l.includes("asistencia") || l.includes("asist")) colAsist2 = c.colIndex;
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
        if (firstVal.startsWith("total") || firstVal.startsWith("promedio") || firstVal.startsWith("resumen")) break;
        continue;
      }

      const nameLower = nameVal.toLowerCase();
      if (nameLower.startsWith("total") || nameLower.startsWith("promedio") || nameLower.startsWith("resumen")) break;

      const run = colRun2 !== -1 ? String(row[colRun2] || "").trim() : "";
      const dv = colDv2 !== -1 ? String(row[colDv2] || "").trim() : "";
      const listNum = colLista2 !== -1 ? String(row[colLista2] || "").trim() : String(r - startRow2 + 1);
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
    return { fileName: file.name, type: "panorama_riesgo", error: "No se encontraron estudiantes en el panorama global." };
  }

  return {
    fileName: file.name,
    type: "panorama_riesgo",
    calificacionesData: parsedPanoramaStudents,
    riesgoData: parsedRiesgoRows,
    riesgoHeaders: parsedRiesgoHeaders,
  };
}
