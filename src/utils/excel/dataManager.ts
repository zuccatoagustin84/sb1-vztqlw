import * as ExcelJS from 'exceljs';
import { Curso } from '../../types';
import { formatGrade } from '../helpers';

export function updateStudentData(
  sheet: ExcelJS.Worksheet,
  alumno: Curso['Alumnos'][0],
  curso: Curso
): void {
  try {
    // Update student name and course while preserving styles
    const nombreCell = sheet.getCell('B16'); // Changed from E16 to B16
    const cursoCell = sheet.getCell('H16');  // Changed from Q16 to H16
    
    const nombreStyle = { ...nombreCell.style };
    const cursoStyle = { ...cursoCell.style };
    
    nombreCell.value = alumno.Nombre;
    cursoCell.value = curso.Text;
    
    nombreCell.style = nombreStyle;
    cursoCell.style = cursoStyle;

    // Update subjects and grades starting from the correct position
    curso.Materias.forEach((materia, index) => {
      const row = 20 + index; // Start from row 20
      const materiaCell = sheet.getCell(`B${row}`);
      const calificacionCell = sheet.getCell(`H${row}`);
      
      const materiaStyle = { ...materiaCell.style };
      const calificacionStyle = { ...calificacionCell.style };
      
      materiaCell.value = materia;
      calificacionCell.value = formatGrade(alumno.Calificaciones[index]);
      
      materiaCell.style = materiaStyle;
      calificacionCell.style = calificacionStyle;
    });
  } catch (error) {
    console.error('Error updating student data:', error);
    throw new Error('Error al actualizar datos del estudiante');
  }
}