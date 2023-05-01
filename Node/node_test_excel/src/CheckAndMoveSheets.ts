import path from 'path';
import {
  changeSalesDocNameAndBuildXlsx,
  createDirectoryIfNotExists,
  deleteEmptyDirectories,
  deleteFile,
  getMapOfExcelPath,
  isSheetEmpty,
  moveFile,
} from './services/SheetManagement';

export function CheckAndMoveSheets(
  basePathFrom: string,
  basePathTo: string,
  basePathToOut: string
) {
  console.time('Total time');

  // Map all folders
  console.time('Mapeia as pastas');
  const mapOfSheets = getMapOfExcelPath(basePathFrom);
  console.timeEnd('Mapeia as pastas');

  // Delete folders without files
  console.time('Deleta pasta');
  for (const [folderName, listOfShets] of mapOfSheets) {
    if (listOfShets.length === 0) {
      deleteEmptyDirectories(path.join(basePathFrom, folderName));
      mapOfSheets.delete(folderName);
    }
  }
  console.timeEnd('Deleta pasta');

  // Delete empty sheets or that has only the first row filled
  console.time('Deleta planilhas vazias');
  for (const [folderName, listOfShets] of mapOfSheets) {
    const filesToRemoveFromList = [];
    for (const [index, eachFile] of listOfShets.entries()) {
      const filePath = path.join(basePathFrom, folderName, eachFile);
      if (isSheetEmpty(filePath)) {
        deleteFile(filePath);
        filesToRemoveFromList.push(index);
      }
    }
    for (const index of filesToRemoveFromList) {
      listOfShets.splice(index, 1);
    }
  }
  console.timeEnd('Deleta planilhas vazias');

  console.time('Novo XLSX');
  for (const [folderName, listOfShets] of mapOfSheets) {
    for (const eachFile of listOfShets) {
      if (eachFile.toLocaleLowerCase().includes('list')) {
        const originalFilePath = path.resolve(
          basePathFrom,
          folderName,
          eachFile
        );
        changeSalesDocNameAndBuildXlsx(
          originalFilePath,
          basePathTo,
          folderName,
          eachFile
        );
      }
    }
  }
  console.timeEnd('Novo XLSX');

  console.time('Move outbound');
  for (const [folderName, listOfShets] of mapOfSheets) {
    for (const eachFile of listOfShets) {
      if (eachFile.toLocaleLowerCase().includes('outbound')) {
        const originalFilePath = path.resolve(
          basePathFrom,
          folderName,
          eachFile
        );
        const newPath = path.resolve(basePathToOut, folderName);
        try {
          createDirectoryIfNotExists(newPath);
          moveFile(originalFilePath, newPath);
        } catch (error) {
          console.log(`Erro ao mover arquivo: ${error}`);
          continue;
        }
      }
    }
  }
  console.timeEnd('Move outbound');
  console.timeEnd('Total time');
}
