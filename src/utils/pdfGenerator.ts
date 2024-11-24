import { jsPDF } from 'jspdf';
import 'jspdf-autotable';
import JSZip from 'jszip';
import { Curso } from '../types';
import { formatGrade, sanitizeFileName } from './helpers';

export async function generatePDFFiles(cursos: Curso[]): Promise<Blob> {
  try {
    const zip = new JSZip();

    cursos.forEach((curso) => {
      curso.Alumnos.forEach((alumno) => {
        const pdf = new jsPDF({
          orientation: 'portrait',
          unit: 'mm',
          format: 'a4'
        });

        pdf.setFontSize(20);
        pdf.setTextColor(44, 62, 80);
        pdf.text('Reporte de Calificaciones', pdf.internal.pageSize.width / 2, 20, { align: 'center' });

        pdf.setFontSize(12);
        pdf.setTextColor(52, 73, 94);
        pdf.text(`Alumno: ${alumno.Nombre}`, 20, 40);
        pdf.text(`Curso: ${curso.Text}`, 20, 50);

        const tableData = curso.Materias.map((materia, index) => [
          materia,
          formatGrade(alumno.Calificaciones[index])
        ]);

        pdf.autoTable({
          startY: 60,
          head: [['Materia', 'Calificación']],
          body: tableData,
          theme: 'grid',
          headStyles: {
            fillColor: [63, 81, 181],
            textColor: 255,
            fontSize: 12,
            fontStyle: 'bold',
            halign: 'center'
          },
          bodyStyles: {
            fontSize: 11,
            textColor: 44,
            cellPadding: 5
          },
          columnStyles: {
            0: { cellWidth: 100 },
            1: { cellWidth: 40, halign: 'center' }
          },
          margin: { top: 60, left: 20, right: 20 },
          didDrawPage: function(data) {
            const pageCount = pdf.internal.getNumberOfPages();
            pdf.setFontSize(10);
            pdf.setTextColor(128, 128, 128);
            pdf.text(
              `Página ${data.pageNumber} de ${pageCount}`,
              pdf.internal.pageSize.width / 2,
              pdf.internal.pageSize.height - 10,
              { align: 'center' }
            );
          }
        });

        const pdfFileName = `${sanitizeFileName(curso.Text)}_${sanitizeFileName(alumno.Nombre)}.pdf`;
        zip.file(pdfFileName, pdf.output('arraybuffer'));
      });
    });

    return await zip.generateAsync({ type: 'blob' });
  } catch (error) {
    console.error('Error en generación de PDF:', error);
    throw new Error(`Error al generar PDFs: ${(error as Error).message}`);
  }
}