import path from 'path';
import {
  buildSheet,
  editCellValue,
  getMapOfExcelPath,
} from './services/SheetManagement';

// NOTE - arquivos SO List e dicionários
const sheetPathSoListCE = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas E-NERP\BASE SOLIST\CE`;
const sheetPathSoListMX = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas E-NERP\BASE SOLIST\MX`;

const dictionaryPathSoListCE = String.raw`C:\Users\GB675AG\OneDrive - EY\Desktop\Tests\excel check\Node\node_test_excel\CE_to_from.yaml`;
const dictionaryPathSoListMX = String.raw`C:\Users\GB675AG\OneDrive - EY\Desktop\Tests\excel check\Node\node_test_excel\MX_to_from.yaml`;

// NOTE - arquivos Outbound e dicionários
const sheetPathOutboundCE = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas\BCK\CE`;
const sheetPathOutboundMX = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas\BCK\IM`;

const dictionaryPathOutboundCE = String.raw`C:\Users\GB675AG\OneDrive - EY\Desktop\Tests\excel check\Node\node_test_excel\OUT_CE_to_from.yaml`;
const dictionaryPathOutboundMX = String.raw`C:\Users\GB675AG\OneDrive - EY\Desktop\Tests\excel check\Node\node_test_excel\OUT_MX_to_from.yaml`;

// NOTE - Ajuste da planilha original
const mapOfSheets = getMapOfExcelPath(sheetPathOutboundMX);

for (const [folderName, listOfShets] of mapOfSheets) {
  for (const eachFile of listOfShets) {
    if (eachFile.toLowerCase().includes('out')) {
      const filePath = path.resolve(sheetPathOutboundMX, folderName, eachFile);
      const newValues = editCellValue(filePath, dictionaryPathOutboundMX);
      buildSheet(filePath, newValues);
    }
  }
}
