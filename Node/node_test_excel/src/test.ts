import fs from 'fs';
import path from 'path';

const pathFrom = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas\BCK\Teste\CE`;
const pathTo = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas E-NERP\BASE SOLIST\CE`;

function getDirectParentDirectoryName(filePath: string): string {
  const directoryTillFile = path.dirname(filePath);
  const fileParentFolderName = path.basename(directoryTillFile);
  return fileParentFolderName;
}

function getMapOfExcelPath(basePath: string): Map<string, string> {
  const listOfFoldersInsidePath = fs.readdirSync(basePath);
  const listOfExcelPath: string[] = [];
  const mapOfExcelByDirectory = new Map();

  for (const folderName of listOfFoldersInsidePath) {
    if (path.extname(folderName.toLowerCase()).includes('xls')) {
      listOfExcelPath.push(folderName);
      const currentPath = path.join(basePath, folderName);
      const folderNameOfFile = getDirectParentDirectoryName(currentPath);
      mapOfExcelByDirectory.set(folderNameOfFile, listOfExcelPath);
    }
  }

  for (const folderName of listOfFoldersInsidePath) {
    const currentFolder = path.join(pathFrom, folderName);

    try {
      const listOfFilesInsideFolder = fs.readdirSync(currentFolder);
      mapOfExcelByDirectory.set(folderName, listOfFilesInsideFolder);
    } catch (error) {
      continue;
    }
  }

  return mapOfExcelByDirectory;
}

getMapOfExcelPath(pathFrom);
