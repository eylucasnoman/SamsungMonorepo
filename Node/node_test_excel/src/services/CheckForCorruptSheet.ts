import fs from 'fs';
import path from 'path';
import nodeXLSX from 'node-xlsx';

export function CheckForCorruptSheet(basePath: string) {
  const listOfFolders = fs.readdirSync(basePath);

  for (const folder of listOfFolders) {
    const soListFolderPath = path.join(basePath, folder);
    const soListSheets = fs.readdirSync(soListFolderPath);

    for (const sheetName of soListSheets) {
      const pathToFile = path.join(soListFolderPath, sheetName);
      const fileName = path.basename(pathToFile);

      try {
        if (pathToFile.toLowerCase().includes('list')) {
          nodeXLSX.parse(pathToFile);
          console.log('Funcionando:', fileName);
        }
      } catch (error) {
        console.log('Corrompida ou não é arquivo:', fileName);
      }
    }
  }
}
