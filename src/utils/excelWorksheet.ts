import * as ExcelJS from 'exceljs';

interface CellStyle {
  numFmt?: string;
  font?: Partial<ExcelJS.Font>;
  alignment?: Partial<ExcelJS.Alignment>;
  border?: Partial<ExcelJS.Borders>;
  fill?: ExcelJS.Fill;
  protection?: Partial<ExcelJS.Protection>;
}

function copyCellStyle(source: ExcelJS.Cell, target: ExcelJS.Cell) {
  try {
    const style: CellStyle = {};
    
    // Copy number format
    if (source.numFmt) {
      style.numFmt = source.numFmt;
    }

    // Copy font
    if (source.style.font) {
      style.font = {
        name: source.style.font.name,
        size: source.style.font.size,
        family: source.style.font.family,
        scheme: source.style.font.scheme,
        charset: source.style.font.charset,
        color: source.style.font.color,
        bold: source.style.font.bold,
        italic: source.style.font.italic,
        underline: source.style.font.underline,
        strike: source.style.font.strike,
        outline: source.style.font.outline,
        vertAlign: source.style.font.vertAlign
      };
    }

    // Copy alignment
    if (source.style.alignment) {
      style.alignment = {
        horizontal: source.style.alignment.horizontal,
        vertical: source.style.alignment.vertical,
        wrapText: source.style.alignment.wrapText,
        indent: source.style.alignment.indent,
        readingOrder: source.style.alignment.readingOrder,
        textRotation: source.style.alignment.textRotation
      };
    }

    // Copy borders
    if (source.style.border) {
      style.border = {
        top: source.style.border.top,
        left: source.style.border.left,
        bottom: source.style.border.bottom,
        right: source.style.border.right,
        diagonal: source.style.border.diagonal,
        diagonalDown: source.style.border.diagonalDown,
        diagonalUp: source.style.border.diagonalUp
      };
    }

    // Copy fill
    if (source.style.fill) {
      style.fill = source.style.fill;
    }

    // Copy protection
    if (source.style.protection) {
      style.protection = {
        locked: source.style.protection.locked,
        hidden: source.style.protection.hidden
      };
    }

    // Apply the copied style
    target.style = style as ExcelJS.Style;
    
    return true;
  } catch (error) {
    console.warn('Error copying cell style:', error);
    return false;
  }
}

export async function copyWorksheetStyles(source: ExcelJS.Worksheet, target: ExcelJS.Worksheet) {
  try {
    // Copy worksheet properties
    target.properties = { ...source.properties };
    target.state = source.state;
    target.views = [...(source.views || [])];

    // Copy column properties
    source.columns.forEach((col, index) => {
      try {
        const targetCol = target.getColumn(index + 1);
        if (col.width) targetCol.width = col.width;
        if (col.hidden) targetCol.hidden = col.hidden;
        if (col.style) targetCol.style = col.style;
      } catch (error) {
        console.warn(`Error copying column ${index + 1}:`, error);
      }
    });

    // Copy row properties and cell styles
    source.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      try {
        const targetRow = target.getRow(rowNumber);
        
        // Copy row properties
        targetRow.height = row.height;
        if (row.hidden) targetRow.hidden = row.hidden;
        if (row.style) targetRow.style = row.style;
        
        // Copy cells
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          try {
            const targetCell = targetRow.getCell(colNumber);
            
            // Copy value with proper type handling
            if (cell.type === ExcelJS.ValueType.Formula) {
              targetCell.value = { formula: cell.formula, result: cell.result };
            } else {
              targetCell.value = cell.value;
            }
            
            // Copy style
            copyCellStyle(cell, targetCell);
            
            // Copy data validation
            if (cell.dataValidation) {
              targetCell.dataValidation = { ...cell.dataValidation };
            }
          } catch (error) {
            console.warn(`Error copying cell at row ${rowNumber}, col ${colNumber}:`, error);
          }
        });
      } catch (error) {
        console.warn(`Error copying row ${rowNumber}:`, error);
      }
    });

    // Copy merged cells
    try {
      // Access the internal model's merges
      const merges = source.model?.merges;
      if (merges) {
        // Iterate through merge ranges
        for (const mergeAddress of Object.keys(merges)) {
          try {
            // Extract the range from the merge address
            const range = mergeAddress.split(':');
            if (range.length === 2) {
              target.mergeCells(range[0], range[1]);
            } else {
              target.mergeCells(mergeAddress);
            }
          } catch (error) {
            console.warn(`Error copying merge range ${mergeAddress}:`, error);
          }
        }
      }
    } catch (error) {
      console.warn('Error copying merged cells:', error);
    }

    // Copy conditional formatting
    try {
      if (source.conditionalFormattings) {
        target.conditionalFormattings = source.conditionalFormattings;
      }
    } catch (error) {
      console.warn('Error copying conditional formatting:', error);
    }

    // Copy page setup
    try {
      if (source.pageSetup) {
        target.pageSetup = { ...source.pageSetup };
      }
    } catch (error) {
      console.warn('Error copying page setup:', error);
    }

    return true;
  } catch (error) {
    console.error('Error in copyWorksheetStyles:', error);
    return false;
  }
}