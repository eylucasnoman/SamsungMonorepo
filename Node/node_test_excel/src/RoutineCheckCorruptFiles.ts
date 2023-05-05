import fs from 'fs';
import path from 'path';
import nodeXLSX from 'node-xlsx';
import { writeLog } from './utils/util';

export function CheckForCorruptSheet(basePath: string) {
  const listOfFolders = fs.readdirSync(basePath);

  for (const folder of listOfFolders) {
    const soListFolderPath = path.join(basePath, folder);
    const soListSheets = fs.readdirSync(soListFolderPath);

    for (const sheetName of soListSheets) {
      const pathToFile = path.join(soListFolderPath, sheetName);
      const fileName = path.basename(pathToFile);

      try {
        const currentFileName = fileName.toLocaleLowerCase();
        if (
          currentFileName.includes('out') ||
          currentFileName.includes('list')
        ) {
          nodeXLSX.parse(pathToFile);
          writeLog('NotCorruptedFiles.log', `${folder} - ${fileName}`);
        }
      } catch (error) {
        writeLog('CorruptedFiles.log', `${folder} - ${fileName}`);
      }
    }
  }
}
