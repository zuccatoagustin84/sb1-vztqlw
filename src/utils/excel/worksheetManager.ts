import * as ExcelJS from 'exceljs';

export async function copyWorksheet(source: ExcelJS.Worksheet, target: ExcelJS.Worksheet): Promise<void> {
  if (!source || !target) {
    throw new Error('Invalid source or target worksheet');
  }

  try {
    // Copy worksheet properties
    target.properties = { ...source.properties };
    target.state = source.state;
    target.views = [...(source.views || [])];
    target.pageSetup = { ...source.pageSetup };
    target.headerFooter = { ...source.headerFooter };

    // Copy column properties and widths
    source.columns.forEach((col, index) => {
      try {
        const targetCol = target.getColumn(index + 1);
        if (col.width !== undefined) targetCol.width = col.width;
        if (col.hidden !== undefined) targetCol.hidden = col.hidden;
        if (col.outlineLevel !== undefined) targetCol.outlineLevel = col.outlineLevel;
        if (col.style) targetCol.style = { ...col.style };
      } catch (err) {
        console.warn(`Warning: Failed to copy column ${index + 1}`, err);
      }
    });

    // Copy rows and cells with improved error handling
    source.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      try {
        const targetRow = target.getRow(rowNumber);
        
        // Copy row properties
        if (row.height !== undefined) targetRow.height = row.height;
        if (row.hidden !== undefined) targetRow.hidden = row.hidden;
        if (row.outlineLevel !== undefined) targetRow.outlineLevel = row.outlineLevel;
        if (row.style) targetRow.style = { ...row.style };

        // Copy cells
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          try {
            const targetCell = targetRow.getCell(colNumber);

            // Copy value with type checking
            if (cell.type === ExcelJS.ValueType.Formula) {
              targetCell.value = {
                formula: cell.formula,
                result: cell.result
              };
            } else if (cell.value !== undefined) {
              targetCell.value = cell.value;
            }

            // Copy style directly without JSON parsing
            if (cell.style) {
              targetCell.style = { ...cell.style };
            }

            // Copy number format
            if (cell.numFmt) {
              targetCell.numFmt = cell.numFmt;
            }

            // Copy data validation
            if (cell.dataValidation) {
              targetCell.dataValidation = { ...cell.dataValidation };
            }

            // Copy hyperlinks
            if (cell.hyperlink) {
              targetCell.hyperlink = cell.hyperlink;
            }

            // Copy text rotation
            if (cell.style?.alignment?.textRotation !== undefined) {
              targetCell.style = {
                ...targetCell.style,
                alignment: {
                  ...targetCell.style?.alignment,
                  textRotation: cell.style.alignment.textRotation
                }
              };
            }
          } catch (err) {
            console.warn(`Warning: Failed to copy cell at row ${rowNumber}, column ${colNumber}`, err);
          }
        });
      } catch (err) {
        console.warn(`Warning: Failed to copy row ${rowNumber}`, err);
      }
    });

    // Copy merged cells
    try {
      const merges = source.model?.merges;
      if (merges) {
        Object.keys(merges).forEach(mergeRange => {
          try {
            target.mergeCells(mergeRange);
          } catch (err) {
            console.warn(`Warning: Failed to merge cells ${mergeRange}`, err);
          }
        });
      }
    } catch (err) {
      console.warn('Warning: Failed to copy merged cells', err);
    }

    // Copy images with correct positioning
    try {
      if (source.model?.drawing?.elements) {
        source.model.drawing.elements.forEach((drawing: any) => {
          if (drawing.type === 'picture') {
            const imageId = drawing.imageId;
            const image = source.workbook.model.media[imageId];
            if (image) {
              const newImageId = target.workbook.addImage({
                buffer: image.buffer,
                extension: image.extension
              });
              target.addImage(newImageId, {
                tl: { col: drawing.from.col, row: drawing.from.row },
                br: { col: drawing.to.col, row: drawing.to.row },
                editAs: drawing.editAs
              });
            }
          }
        });
      }
    } catch (err) {
      console.warn('Warning: Failed to copy images', err);
    }

    // Copy conditional formatting
    if (source.conditionalFormattings) {
      target.conditionalFormattings = { ...source.conditionalFormattings };
    }

    // Copy autofilter
    if (source.autoFilter) {
      target.autoFilter = { ...source.autoFilter };
    }

  } catch (error) {
    console.error('Critical error in copyWorksheet:', error);
    throw new Error(`Failed to copy worksheet: ${(error as Error).message}`);
  }
}