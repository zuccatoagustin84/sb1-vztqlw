import React, { useState } from 'react';
import { FileUpload } from './components/FileUpload';
import { ProcessButton } from './components/ProcessButton';
import { DownloadButton } from './components/DownloadButton';
import { processExcelData, generatePDFsAndZip, generateProcessedExcel } from './utils/excelProcessor';
import { Curso, ProcessingError } from './types';
import { FileSpreadsheet } from 'lucide-react';
import { sanitizeFileName } from './utils/helpers';

function App() {
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [dataFile, setDataFile] = useState<File | null>(null);
  const [processedData, setProcessedData] = useState<Curso[] | null>(null);
  const [zipUrl, setZipUrl] = useState<string | null>(null);
  const [excelUrls, setExcelUrls] = useState<string[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<ProcessingError | null>(null);

  const handleTemplateUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFile = event.target.files?.[0];
    if (uploadedFile) {
      setTemplateFile(uploadedFile);
      setError(null);
      // Reset processed data when new file is uploaded
      setProcessedData(null);
      setZipUrl(null);
      setExcelUrls([]);
    }
  };

  const handleDataFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFile = event.target.files?.[0];
    if (uploadedFile) {
      setDataFile(uploadedFile);
      setError(null);
      // Reset processed data when new file is uploaded
      setProcessedData(null);
      setZipUrl(null);
      setExcelUrls([]);
    }
  };

  const processExcel = async () => {
    if (!templateFile || !dataFile) return;
    
    setIsProcessing(true);
    setError(null);
    
    try {
      const cursos = await processExcelData(templateFile, dataFile);
      setProcessedData(cursos);
      
      // Generate Excel files
      const excelBlobs = await generateProcessedExcel(templateFile, cursos);
      const newExcelUrls = excelBlobs.map((blob, index) => {
        const url = URL.createObjectURL(blob);
        return url;
      });
      
      // Generate ZIP with PDFs
      const zipContent = await generatePDFsAndZip(cursos);
      const newZipUrl = URL.createObjectURL(zipContent);
      
      // Clean up old URLs
      if (zipUrl) URL.revokeObjectURL(zipUrl);
      excelUrls.forEach(url => URL.revokeObjectURL(url));
      
      setExcelUrls(newExcelUrls);
      setZipUrl(newZipUrl);
    } catch (err) {
      setError({
        message: 'Error al procesar los archivos',
        details: (err as Error).message
      });
      console.error('Error processing files:', err);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-indigo-50 to-blue-50 flex items-center justify-center p-4">
      <div className="bg-white p-8 rounded-xl shadow-lg w-full max-w-md">
        <h1 className="text-2xl font-bold mb-6 text-center text-indigo-900">
          Procesador de Calificaciones
        </h1>
        
        <div className="space-y-6">
          <FileUpload
            id="template-upload"
            label="Subir Plantilla Excel"
            file={templateFile}
            onChange={handleTemplateUpload}
          />

          <FileUpload
            id="data-file-upload"
            label="Subir Archivo de Datos Excel"
            file={dataFile}
            onChange={handleDataFileUpload}
          />

          {error && (
            <div className="p-3 bg-red-50 border border-red-200 rounded-lg">
              <p className="text-sm text-red-800 font-medium">{error.message}</p>
              {error.details && (
                <p className="text-xs text-red-600 mt-1">{error.details}</p>
              )}
            </div>
          )}

          <ProcessButton
            onClick={processExcel}
            disabled={!templateFile || !dataFile}
            isProcessing={isProcessing}
          />

          {processedData && excelUrls.length > 0 && (
            <div className="space-y-2">
              {processedData.map((curso, cursoIndex) => (
                <div key={curso.Text} className="space-y-2">
                  <h3 className="font-medium text-gray-700">{curso.Text}</h3>
                  {curso.Alumnos.map((alumno, alumnoIndex) => {
                    const urlIndex = cursoIndex * curso.Alumnos.length + alumnoIndex;
                    return (
                      <a
                        key={alumno.Nombre}
                        href={excelUrls[urlIndex]}
                        download={`${sanitizeFileName(curso.Text)}_${sanitizeFileName(alumno.Nombre)}.xlsx`}
                        className="flex items-center px-3 py-2 text-sm text-gray-700 hover:bg-gray-50 rounded-lg
                          border border-gray-200 transition-colors duration-200"
                      >
                        <FileSpreadsheet className="h-4 w-4 mr-2 text-gray-500" />
                        {alumno.Nombre}
                      </a>
                    );
                  })}
                </div>
              ))}
            </div>
          )}

          {zipUrl && <DownloadButton zipUrl={zipUrl} />}
        </div>
      </div>
    </div>
  );
}

export default App;