import fs from 'fs';
import path from 'path';
import nodeXLSX from 'node-xlsx';

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
    console.log(`Arquivo "${newFileName}" criado em: ${newFolderName}.`);
  } catch (error) {
    console.log(`Falha ao escrever o arquivo "${newFileName}."`);
    console.log(error);
  }
}

export function moveFile(filePath: string, newDir: string) {
  const fileName = path.basename(filePath);
  const dest = path.resolve(newDir, fileName);

  fs.rename(filePath, dest, (err) => {
    if (err) console.warn(err);
    else console.log(`${fileName} movido.`);
  });
}

/*
- Pegar todos os arquivos XLSX e nomes de pastas diretas ✅
- Checar pastas sem nada ✅
  - Excluir ✅
- Checar arquivos XLSX sem conteúdo ✅
  - Excluir ✅
- Alterar: Sales Document -> Sales document ✅
  - OBS.: para alterar, eu tenho que gerar um novo arquivo ✅
    - Garantir que a extensão desse arquivo seja a mesma do original ✅
- Mover XLSX com nome SOLIST para outro diretório com o mesmo nome de pasta ✅
*/

// NOTE - fazer um loop para testar com mais de um arquivo antes de rodar na pasta CE
