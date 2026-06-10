export type ObservationKind = "positiva" | "negativa" | "otros";
export type QuickFilter = "all" | ObservationKind;

export type Observation = {
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

export type PendienteItem = {
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

export type GradeItem = {
  label: string;
  weight: string;
  value: number | null;
};

export type CalificacionSubjectData = {
  subjectName: string;
  grades: GradeItem[];
  p1: number | null;
  p2: number | null;
  pf: number | null;
};

export type CalificacionStudent = {
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

export interface ExcelSectionValue {
  subjectName?: string;
  type: "grade" | "p1" | "p2" | "pf" | "p1_gen" | "p2_gen" | "pf_gen";
  label?: string;
  value: number | string | null;
}

export interface ExcelSectionRow {
  label: string;
  values: ExcelSectionValue[];
}

export interface ExcelExtraData {
  indicadores: ExcelSectionRow[];
  categorias: ExcelSectionRow[];
  leyendas: string[];
}

export interface RiesgoStudent {
  id: string;
  lista: string;
  estudiante: string;
  run: string;
  dv: string;
  promedioGeneral: number | null;
  asistencia: string | null;
  dynamicFields: Record<string, string | number | null>;
}

export type AppMode = "observations" | "pendientes" | "calificaciones" | "panorama_riesgo";
