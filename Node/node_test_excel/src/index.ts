import { CheckAndMoveSheets } from './CheckAndMoveSheets';

const basePathFrom = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas\BCK\Teste\CE`;
const basePathTo = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas E-NERP\BASE SOLIST\CE`;
const basePathToOut = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas\BCK\OUTBOUND`;
// const basePathFrom = String.raw`C:\Users\GB675AG\OneDrive - EY\Desktop\bkp_tst\CE_from`;
// const basePathTo = String.raw`C:\Users\GB675AG\OneDrive - EY\Desktop\bkp_tst\CE_to`;
// const basePathToOut = String.raw`C:\Users\GB675AG\OneDrive - EY\Desktop\bkp_tst\bck`;

CheckAndMoveSheets(basePathFrom, basePathTo, basePathToOut);
