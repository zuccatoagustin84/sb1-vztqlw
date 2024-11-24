import React from 'react';
import { FileDown } from 'lucide-react';

interface DownloadButtonProps {
  zipUrl: string;
}

export function DownloadButton({ zipUrl }: DownloadButtonProps) {
  return (
    <a
      href={zipUrl}
      download="informes_calificaciones.zip"
      className="flex items-center justify-center w-full px-4 py-3 rounded-lg text-sm font-medium text-white
        bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500
        transition-colors duration-200"
    >
      <FileDown className="mr-2 h-5 w-5" />
      Descargar todos los informes (ZIP)
    </a>
  );
}