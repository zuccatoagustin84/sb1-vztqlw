import * as ExcelJS from 'exceljs';
import { Curso } from '../types';

export async function generateExcelFiles(templateData: ArrayBuffer, cursos: Curso[]): Promise<Blob[]> {
  const processedFiles: Blob[] = [];

  if (!templateData || templateData.byteLength === 0) {
    throw new Error('La plantilla Excel está vacía o es inválida');
  }

  if (!cursos || !Array.isArray(cursos) || cursos.length === 0) {
    throw new Error('No hay datos de cursos para procesar');
  }

  try {
    // Load template workbook once
    const templateWorkbook = new ExcelJS.Workbook();
    await templateWorkbook.xlsx.load(templateData);

    for (const curso of cursos) {
      for (const alumno of curso.Alumnos) {
        try {
          // Create new workbook from template buffer for each student
          const newWorkbook = new ExcelJS.Workbook();
          await newWorkbook.xlsx.load(templateData);
          
          // Get the first worksheet
          const sheet = newWorkbook.worksheets[0];
          if (!sheet) {
            throw new Error('No se encontró la hoja de trabajo principal');
          }

          // Update student data
          const nombreCell = sheet.getCell('B16');
          const cursoCell = sheet.getCell('H16');
          
          nombreCell.value = alumno.Nombre;
          cursoCell.value = curso.Text;

          // Update subjects and grades
          curso.Materias.forEach((materia, index) => {
            const row = 20 + index; // Start from row 20
            const materiaCell = sheet.getCell(`B${row}`);
            const calificacionCell = sheet.getCell(`H${row}`);
            
            materiaCell.value = materia;
            calificacionCell.value = alumno.Calificaciones[index] || 'A';
          });

          // Generate Excel file
          const buffer = await newWorkbook.xlsx.writeBuffer();
          const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          });
          
          processedFiles.push(blob);
          console.log(`Successfully processed: ${alumno.Nombre}`);
        } catch (error) {
          console.error(`Error processing student ${alumno.Nombre}:`, error);
          // Continue with next student instead of throwing
          continue;
        }
      }
    }

    if (processedFiles.length === 0) {
      throw new Error('No se pudo generar ningún archivo Excel');
    }

    return processedFiles;
  } catch (error) {
    console.error('Error generating Excel files:', error);
    throw new Error(`Error al generar archivos Excel: ${(error as Error).message}`);
  }
}