import fs from 'node:fs';
import path from 'node:path';
import nodeXLSX from 'node-xlsx';
import { writeLog } from '../utils/util';
import yaml from 'js-yaml';

// TODO - alterar para função recursiva que adiciona ao Map todos os .xlsx dentro de todas as pastas
/*
  Retorna um map com nome da pasta e lista de planilhas. Ex.: {11.01.2022: [outbound.xlsx, solist ce.xlsx]}
  basePath: é a string do caminho inteiro até a pasta
 */
export function getMapOfExcelPath(basePath: string): Map<string, string[]> {
  // "fs.readdirSync" lista todos os arquivos e pastas dentro do diretório
  const listOfFoldersInsidePath = fs.readdirSync(basePath);
  const listOfExcelPath: string[] = [];
  const mapOfExcelByDirectory = new Map<string, string[]>();

  // Itera sobre as pastas encontradas
  for (const folderName of listOfFoldersInsidePath) {
    if (path.extname(folderName.toLowerCase()).includes('xls')) {
      // Se o conteúdo for um arquivo com extensão "xls", eu pego o nome da pasta
      // e pego os nomes de todos os arquivos dentro da pasta e adiciono no Map da linha 26
      listOfExcelPath.push(folderName);
      const currentPath = path.join(basePath, folderName);
      const folderNameOfFile = getDirectParentDirectoryName(currentPath);
      mapOfExcelByDirectory.set(folderNameOfFile, listOfExcelPath);
    }
  }

  // Caso o conteúdo seja uma pasta, eu pego os nomes de todos os arquivos dentro dela
  // e também adiciono junto com o nome da pastas, no Map na linha 36
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

// filePath: é o caminho completo até o arquivo. Esse retorna o nome da pasta
// direta que contém o arquivo
export function getDirectParentDirectoryName(filePath: string): string {
  const directoryTillFile = path.dirname(filePath);
  const fileParentFolderName = path.basename(directoryTillFile);
  return fileParentFolderName;
}

// folderPath: caminho completo até a pasta
// Essa função verifica se a pasta está vazia e a deleta
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

// Retorna "true" se o arquivo tiver vazio
export function isSheetEmpty(filePath: string): Boolean {
  const [{ name, data }] = nodeXLSX.parse(filePath, { sheetRows: 3 });

  let isEmpty = false;
  if (data.length <= 1) isEmpty = true;

  return isEmpty;
}

// Deleta qualquer arquivo
export function deleteFile(filePath: string) {
  fs.rm(filePath, (err) => {
    if (err !== null) console.log(err);
  });
}

// Verifica se já existe uma pasta com o nome definido, se não, cria a pasta.
// Aqui deve-se passar o caminho completo, incluindo o nome da nova pasta
// ex.: C:/Users/Desktop/Teste
// "Teste" é a pasta que não existe e será criada agora.
export function createDirectoryIfNotExists(newFolderPath: string) {
  try {
    if (!fs.existsSync(newFolderPath)) fs.mkdirSync(newFolderPath);
  } catch (error) {
    console.log(`Diretório não foi criado. Erro: ${error}`);
  }
}

// Função antiga que muda o "Sales Document" para "Sales document"
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

// Move o arquivo para a pasta desejada
// 1 parâmetro é o caminho até o arquivo e o 2 é o novo caminho
export function moveFile(filePath: string, newDir: string) {
  const fileName = path.basename(filePath);
  const dest = path.resolve(newDir, fileName);

  fs.rename(filePath, dest, (err) => {
    if (err) console.warn(err);
  });
}

// Edita os títulos de dada planilha de acordo com o dicionário YAML passado
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

// Recria a planilha no diretório passado. Os dados da planilha tem que vir
// exatamente como o "nodeXLSX.parse()" retorna
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
