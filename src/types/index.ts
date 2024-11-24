export interface Alumno {
  Nombre: string;
  Calificaciones: (string | null)[];
}

export interface Curso {
  Text: string;
  Materias: string[];
  Alumnos: Alumno[];
}

export interface ProcessingError {
  message: string;
  details?: string;
}