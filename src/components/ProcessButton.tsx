import React from 'react';
import { FileSpreadsheet } from 'lucide-react';

interface ProcessButtonProps {
  onClick: () => void;
  disabled: boolean;
  isProcessing: boolean;
}

export function ProcessButton({ onClick, disabled, isProcessing }: ProcessButtonProps) {
  return (
    <button
      onClick={onClick}
      disabled={disabled || isProcessing}
      className={`w-full flex items-center justify-center px-4 py-3 rounded-lg text-sm font-medium text-white
        ${isProcessing || disabled
          ? 'bg-gray-400 cursor-not-allowed' 
          : 'bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500'}
        transition-colors duration-200`}
    >
      <FileSpreadsheet className="mr-2 h-5 w-5" />
      {isProcessing ? 'Procesando...' : 'Procesar Excel y Generar PDFs'}
    </button>
  );
}