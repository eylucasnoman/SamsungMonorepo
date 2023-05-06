import path from 'path';
import { getMapOfExcelPath } from './services/SheetManagement';
import { pathToBckOutboundMX } from './utils/defaultPaths';
import xlsx from 'node-xlsx';
import ExcelJS from 'exceljs';

const workbook = new ExcelJS.Workbook();

const mapOfSheets = getMapOfExcelPath(pathToBckOutboundMX);
console.time();
outerloop: for (const [folderName, listOfFiles] of mapOfSheets) {
  for (const fileName of listOfFiles) {
    if (fileName.toLowerCase().includes('out')) {
      const outboundFile = path.join(pathToBckOutboundMX, folderName, fileName);

      const [{ name, data }] = xlsx.parse(outboundFile);

      if (name !== 'Sheet1')
        console.log(`${folderName} - ${fileName} | Nome: ${name}`);
      else console.log(`OK: ${folderName} - ${fileName} | Nome: ${name}`);

      // -------------------------------------
      // Esse é o novo código ainda em teste.
      // Para usá-lo, é só comentar acima das linhas 16 a 20 e descomentar aqui da 26 a 34

      // workbook.xlsx
      //   .readFile(outboundFile)
      //   .then(() => {
      //     const names = workbook.worksheets.map((worksheet) => worksheet.name);

      //     if (!names.includes('Sheet1'))
      //       console.log(`${folderName} - ${fileName} | Nomes: ${names}`);
      //   })
      //   .catch((error) => console.log(error));
    }
  }
}
console.timeEnd();
