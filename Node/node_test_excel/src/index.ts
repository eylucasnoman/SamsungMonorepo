import fs from 'fs';
import path from 'path';
import nodeXLSX from 'node-xlsx';

console.time('Tempo gasto');

const ceDir = String.raw`C:\Users\GB675AG\Downloads\excel check\excel check\excel`;
const listOfFolders = fs.readdirSync(ceDir);

for (const folder of listOfFolders) {
  const soListFolderPath = path.join(ceDir, folder);
  const soListSheets = fs.readdirSync(soListFolderPath);

  for (const sheetName of soListSheets) {
    const pathToFile = path.join(soListFolderPath, sheetName);
    const fileName = path.basename(pathToFile);

    try {
      nodeXLSX.parse(pathToFile);
      console.log('Funcionando:', fileName);
    } catch (error) {
      console.log('Corrompida:', fileName);
    }
  }
}

console.timeEnd('Tempo gasto');

const pathFrom = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas\BCK\Teste\CE`;
const pathTo = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas E-NERP\BASE SOLIST\CE`;

function createListOfPathToExcel(basePath: string) {
  const listOfFoldersInsidePath = fs.readdirSync(basePath);
}
