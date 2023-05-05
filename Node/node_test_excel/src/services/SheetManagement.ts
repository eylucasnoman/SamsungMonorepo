import fs from 'node:fs';
import path from 'node:path';
import nodeXLSX from 'node-xlsx';
import { writeLog } from '../utils/util';
import yaml from 'js-yaml';

// TODO - alterar para função recursiva que adiciona ao Map todos os .xlsx dentro de todas as pastas
export function getMapOfExcelPath(basePath: string): Map<string, string[]> {
  const listOfFoldersInsidePath = fs.readdirSync(basePath);
  const listOfExcelPath: string[] = [];
  const mapOfExcelByDirectory = new Map<string, string[]>();

  for (const folderName of listOfFoldersInsidePath) {
    if (path.extname(folderName.toLowerCase()).includes('xls')) {
      listOfExcelPath.push(folderName);
      const currentPath = path.join(basePath, folderName);
      const folderNameOfFile = getDirectParentDirectoryName(currentPath);
      mapOfExcelByDirectory.set(folderNameOfFile, listOfExcelPath);
    }
  }

  for (const folderName of listOfFoldersInsidePath) {
    const currentFolder = path.join(basePath, folderName);

    try {
      const listOfFilesInsideFolder = fs.readdirSync(currentFolder);
      mapOfExcelByDirectory.set(folderName, listOfFilesInsideFolder);
    } catch (error) {
      continue;
    }
  }

  return mapOfExcelByDirectory;
}
// const mapResult = getMapOfExcelPath(pathFrom);

export function getDirectParentDirectoryName(filePath: string): string {
  const directoryTillFile = path.dirname(filePath);
  const fileParentFolderName = path.basename(directoryTillFile);
  return fileParentFolderName;
}

export function deleteEmptyDirectories(folderPath: string) {
  if (fs.readdirSync(folderPath).length !== 0) {
    console.log('O diretório não está vazio!');
    return;
  }

  try {
    fs.rmdir(path.resolve(folderPath), (err) => {
      if (err !== null) console.log(err);
    });
  } catch (error) {
    console.log('Erro ao tentar deletar a pasta:', error);
  }
}

export function isSheetEmpty(filePath: string): Boolean {
  const [{ name, data }] = nodeXLSX.parse(filePath, { sheetRows: 3 });

  let isEmpty = false;
  if (data.length <= 1) isEmpty = true;

  return isEmpty;
}

export function deleteFile(filePath: string) {
  fs.rm(filePath, (err) => {
    if (err !== null) console.log(err);
  });
}

export function createDirectoryIfNotExists(newFolderPath: string) {
  try {
    if (!fs.existsSync(newFolderPath)) fs.mkdirSync(newFolderPath);
  } catch (error) {
    console.log(`Diretório não foi criado. Erro: ${error}`);
  }
}

export function changeSalesDocNameAndBuildXlsx(
  originalFilePath: string,
  newBasePath: string,
  newFolderName: string,
  newFileName: string,
  newFileExtension?: string
) {
  newFileName = path.parse(newFileName).name;

  const [{ name, data }] = nodeXLSX.parse(originalFilePath);
  data[0][0] = 'Sales document';

  const newSheet = nodeXLSX.build([{ name, data }]);

  const pathToBeWritten = path.join(newBasePath, newFolderName);
  createDirectoryIfNotExists(pathToBeWritten);

  const originalFileExtension = path.extname(originalFilePath);
  newFileExtension === undefined
    ? (newFileName += originalFileExtension)
    : (newFileName += newFileExtension);

  const newFilePath = path.join(pathToBeWritten, newFileName);
  try {
    fs.writeFileSync(newFilePath, Buffer.from(newSheet));
    writeLog(
      'NewlyCreatedFiles.log',
      `Arquivo "${newFileName}" criado em: ${newFolderName}.`
    );
  } catch (error) {
    writeLog(
      'ErrorFileCreation.log',
      `Arquivo "${newFileName}" - Erro: ${error}.`
    );
  }
}

export function moveFile(filePath: string, newDir: string) {
  const fileName = path.basename(filePath);
  const dest = path.resolve(newDir, fileName);

  fs.rename(filePath, dest, (err) => {
    if (err) console.warn(err);
  });
}

export function editCellValue(
  filePath: string,
  dictionaryFile: fs.PathOrFileDescriptor
) {
  const fileContents = fs.readFileSync(dictionaryFile, 'utf-8');

  const toFromFile = yaml.load(fileContents) as object;
  const excelTitles = new Map<string, [string]>(Object.entries(toFromFile));

  const [{ name, data }] = nodeXLSX.parse(filePath);

  for (const cellName of data[0]) {
    for (const [newTitle, originalTitle] of excelTitles) {
      if (originalTitle.includes(cellName) && newTitle !== cellName) {
        const cellIndex = data[0].findIndex((cell) => cell === cellName);
        data[0][cellIndex] = newTitle;
      }
    }
  }

  return [{ name, data }];
}

export function buildSheet(
  pathToSaveFile: string,
  data: { name: string; data: any[][] }[]
) {
  const newSheet = nodeXLSX.build(data);
  const fileName = path.basename(pathToSaveFile);
  const dirName = path.dirname(pathToSaveFile);

  try {
    fs.writeFileSync(pathToSaveFile, Buffer.from(newSheet));
    writeLog(
      'NewlyCreatedFiles.log',
      `Arquivo "${fileName}" criado em: ${dirName}.`
    );
  } catch (error) {
    writeLog(
      'ErrorFileCreation.log',
      `Arquivo "${fileName}" - Erro: ${error}.`
    );
  }
}
