import { CheckForCorruptSheet } from './RoutineCheckCorruptFiles';

const basePathFromCE = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas\BCK\Teste\CE`;
const basePathToCE = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas E-NERP\BASE SOLIST\CE`;
const basePathToOutCE = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas\BCK\OUTBOUND\CE`;

const basePathFromMX = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas\BCK\Teste\MX`;
const basePathToMX = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas E-NERP\BASE SOLIST\MX`;
const basePathToOutMX = String.raw`C:\Users\GB675AG\EY\Projeto Samsung Order Mgmt - General\03. Gestão da Rotina\Automações\Ferramentas\BCK\OUTBOUND\MX`;

CheckForCorruptSheet(basePathFromMX);
