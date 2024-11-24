import * as ExcelJS from 'exceljs';

export async function copyWorksheetImages(
  sourceWorkbook: ExcelJS.Workbook,
  targetWorkbook: ExcelJS.Workbook,
  sourceSheet: ExcelJS.Worksheet,
  targetSheet: ExcelJS.Worksheet
): Promise<boolean> {
  try {
    const media = sourceWorkbook.model.media;
    if (!media || Object.keys(media).length === 0) {
      return true;
    }

    // Create image ID mapping
    const imageMap = new Map<string, number>();
    
    for (const [id, image] of Object.entries(media)) {
      try {
        const newImageId = targetWorkbook.addImage({
          buffer: image.buffer,
          extension: image.extension,
        });
        imageMap.set(id, newImageId);
      } catch (err) {
        console.warn(`Failed to copy image ${id}:`, err);
        continue;
      }
    }

    // Copy image placements
    const drawings = sourceSheet.model.drawing?.elements || [];
    for (const drawing of drawings) {
      if (drawing.type === 'picture') {
        const newImageId = imageMap.get(String(drawing.imageId));
        if (typeof newImageId === 'number') {
          try {
            targetSheet.addImage(newImageId, {
              tl: { col: drawing.from.col, row: drawing.from.row },
              br: { col: drawing.to.col, row: drawing.to.row },
              editAs: drawing.editAs,
            });
          } catch (err) {
            console.warn(`Failed to place image ${drawing.imageId}:`, err);
            continue;
          }
        }
      }
    }

    return true;
  } catch (error) {
    console.error('Error copying worksheet images:', error);
    return false;
  }
}