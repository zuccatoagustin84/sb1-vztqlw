import * as ExcelJS from 'exceljs';
import { Buffer } from 'buffer';

interface ImageInfo {
  buffer: Buffer;
  extension: string;
  width?: number;
  height?: number;
}

export async function copyImages(sourceWorkbook: ExcelJS.Workbook, targetWorkbook: ExcelJS.Workbook): Promise<void> {
  if (!sourceWorkbook?.model?.media || !targetWorkbook) {
    console.warn('No media found in source workbook or invalid target workbook');
    return;
  }

  const imageMap = new Map<string, number>();

  try {
    // First pass: Copy all images from the media collection
    for (const [mediaId, mediaItem] of Object.entries(sourceWorkbook.model.media)) {
      try {
        if (!mediaItem.buffer) {
          console.warn(`No buffer found for media item ${mediaId}`);
          continue;
        }

        const imageInfo: ImageInfo = {
          buffer: mediaItem.buffer,
          extension: mediaItem.extension,
          width: mediaItem.width,
          height: mediaItem.height
        };

        const newImageId = targetWorkbook.addImage(imageInfo);
        imageMap.set(mediaId, newImageId);
        console.log(`Successfully copied image ${mediaId} to new workbook`);
      } catch (err) {
        console.error(`Failed to copy media item ${mediaId}:`, err);
      }
    }

    // Second pass: Copy image placements in each worksheet
    sourceWorkbook.worksheets.forEach((sourceSheet, sheetIndex) => {
      const targetSheet = targetWorkbook.worksheets[sheetIndex];
      if (!targetSheet) {
        console.warn(`Target sheet ${sheetIndex} not found`);
        return;
      }

      const drawings = sourceSheet.model?.drawing?.elements || [];
      drawings.forEach((drawing: any) => {
        if (drawing.type === 'picture') {
          const newImageId = imageMap.get(String(drawing.imageId));
          if (typeof newImageId === 'number') {
            try {
              targetSheet.addImage(newImageId, {
                tl: {
                  col: drawing.from.col,
                  row: drawing.from.row,
                  colOff: drawing.from.colOff || 0,
                  rowOff: drawing.from.rowOff || 0
                },
                br: {
                  col: drawing.to.col,
                  row: drawing.to.row,
                  colOff: drawing.to.colOff || 0,
                  rowOff: drawing.to.rowOff || 0
                },
                editAs: drawing.editAs || 'oneCell'
              });
              console.log(`Successfully placed image ${drawing.imageId} in sheet ${sourceSheet.name}`);
            } catch (err) {
              console.error(`Failed to place image ${drawing.imageId}:`, err);
            }
          }
        }
      });
    });
  } catch (error) {
    console.error('Error in copyImages:', error);
    throw new Error(`Failed to copy images: ${(error as Error).message}`);
  }
}