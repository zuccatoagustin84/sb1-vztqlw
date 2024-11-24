import React from 'react';
import { Upload } from 'lucide-react';

interface FileUploadProps {
  id: string;
  label: string;
  file: File | null;
  onChange: (event: React.ChangeEvent<HTMLInputElement>) => void;
}

export function FileUpload({ id, label, file, onChange }: FileUploadProps) {
  return (
    <div className="relative">
      <label 
        htmlFor={id} 
        className={`flex items-center justify-center w-full px-4 py-3 border-2 border-dashed rounded-lg
          ${file ? 'border-green-500 bg-green-50' : 'border-gray-300 hover:border-indigo-500'}
          transition-colors duration-200 cursor-pointer group`}
      >
        <Upload className={`mr-2 h-5 w-5 ${file ? 'text-green-500' : 'text-gray-400 group-hover:text-indigo-500'}`} />
        <span className={`text-sm font-medium ${file ? 'text-green-700' : 'text-gray-600 group-hover:text-indigo-600'}`}>
          {file ? file.name : label}
        </span>
      </label>
      <input 
        id={id} 
        type="file" 
        accept=".xlsx,.xls" 
        onChange={onChange} 
        className="sr-only" 
      />
    </div>
  );
}