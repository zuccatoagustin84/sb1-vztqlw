export function sanitizeFileName(name: string): string {
  return name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
}

export function formatGrade(grade: string | null): string {
  if (!grade || grade.trim() === '' || grade.trim() === '-') {
    return 'A';
  }
  return grade.trim();
}