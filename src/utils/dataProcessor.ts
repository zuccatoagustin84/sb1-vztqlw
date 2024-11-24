import * as XLSX from 'xlsx';
import { Curso, Alumno } from '../types';

export function processDataFile(dataWorkbook: XLSX.WorkBook): Curso[] {
  const cursos: Curso[] = [];

  dataWorkbook.SheetNames.forEach((sheetName) => {
    const worksheet = dataWorkbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

    if (!jsonData || jsonData.length < 3) {
      throw new Error(`Formato inválido en la hoja ${sheetName}`);
    }

    const materias = jsonData[2].slice(2).filter(Boolean) as string[];
    const alumnos: Alumno[] = jsonData.slice(3)
      .map((row) => ({
        Nombre: row[1] as string,
        Calificaciones: row.slice(2, 2 + materias.length) as (string | null)[],
      }))
      .filter((alumno) => alumno.Nombre);

    cursos.push({
      Text: sheetName,
      Materias: materias,
      Alumnos: alumnos,
    });
  });

  return cursos;
}