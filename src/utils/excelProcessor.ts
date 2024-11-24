import * as XLSX from 'xlsx';
import { Curso } from '../types';
import { processDataFile } from './dataProcessor';
import { generateExcelFiles } from './excelGenerator';
import { generatePDFFiles } from './pdfGenerator';

export async function processExcelData(templateFile: File, dataFile: File): Promise<Curso[]> {
  try {
    const dataFileData = await dataFile.arrayBuffer();
    const dataWorkbook = XLSX.read(dataFileData, {
      cellDates: true,
      cellNF: true,
      cellStyles: true
    });
    return processDataFile(dataWorkbook);
  } catch (error) {
    console.error('Error processing Excel data:', error);
    throw new Error(`Error al procesar datos Excel: ${(error as Error).message}`);
  }
}

export async function generateProcessedExcel(templateFile: File, cursos: Curso[]): Promise<Blob[]> {
  try {
    const templateData = await templateFile.arrayBuffer();
    return await generateExcelFiles(templateData, cursos);
  } catch (error) {
    console.error('Error generating Excel files:', error);
    throw new Error(`Error al generar archivos Excel: ${(error as Error).message}`);
  }
}

export async function generatePDFsAndZip(cursos: Curso[]): Promise<Blob> {
  try {
    return await generatePDFFiles(cursos);
  } catch (error) {
    console.error('Error generating PDFs:', error);
    throw new Error(`Error al generar PDFs: ${(error as Error).message}`);
  }
}