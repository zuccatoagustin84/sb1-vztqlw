import * as ExcelJS from 'exceljs';

export async function createWorkbook(templateData: ArrayBuffer): Promise<ExcelJS.Workbook> {
  if (!templateData || templateData.byteLength === 0) {
    throw new Error('Template data is empty or invalid');
  }

  try {
    console.log('Creating workbook from template...');
    const workbook = new ExcelJS.Workbook();
    
    // Load with specific options to ensure proper image handling
    await workbook.xlsx.load(templateData, {
      base64: false,
      password: '',
      ignoreNodes: ['extLst', 'ext'],
    });

    if (!workbook.worksheets || workbook.worksheets.length === 0) {
      throw new Error('Template contains no worksheets');
    }

    // Log media information
    const mediaCount = Object.keys(workbook.model.media || {}).length;
    console.log(`Template loaded with ${workbook.worksheets.length} worksheets and ${mediaCount} media items`);

    return workbook;
  } catch (error) {
    console.error('Error creating workbook:', error);
    throw new Error(`Failed to create workbook: ${(error as Error).message}`);
  }
}

export function initializeNewWorkbook(templateWorkbook: ExcelJS.Workbook): ExcelJS.Workbook {
  if (!templateWorkbook) {
    throw new Error('Invalid template workbook provided');
  }

  try {
    console.log('Initializing new workbook...');
    const newWorkbook = new ExcelJS.Workbook();
    
    // Copy workbook properties
    newWorkbook.creator = templateWorkbook.creator || 'Grade System';
    newWorkbook.lastModifiedBy = 'Grade System';
    newWorkbook.created = templateWorkbook.created || new Date();
    newWorkbook.modified = new Date();
    newWorkbook.properties = { ...templateWorkbook.properties };
    
    // Initialize media container
    if (!newWorkbook.model) {
      newWorkbook.model = {} as any;
    }
    if (!newWorkbook.model.media) {
      newWorkbook.model.media = {};
    }

    console.log('New workbook initialized successfully');
    return newWorkbook;
  } catch (error) {
    console.error('Error initializing new workbook:', error);
    throw new Error(`Failed to initialize workbook: ${(error as Error).message}`);
  }
}