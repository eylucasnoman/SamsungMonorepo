import os from 'os';
import path from 'path';

// NOTE - base paths
const rootDir = process.cwd();
const osUser = os.userInfo().username;

const pathToAutomation = path.join('C:', 'Users', osUser, 'EY', 'Projeto Samsung Order Mgmt - General', '03. Gestão da Rotina', 'Automações');

const pathToBCK = path.join(pathToAutomation, 'ferramentas', 'BCK');

const pathToBaseSoList = path.join(pathToAutomation, 'Ferramentas E-NERP', 'BASE SOLIST');

// NOTE - path to sheet files
export const pathToBckSoListCE = path.join(pathToBCK, 'Teste', 'CE');
export const pathToBckSoListMX = path.join(pathToBCK, 'Teste', 'MX');

export const pathToBckOutboundCE = path.join(pathToBCK, 'CE');
export const pathToBckOutboundMX = path.join(pathToBCK, 'IM');

export const pathToBaseSoListCE = path.join(pathToBaseSoList, 'CE');
export const pathToBaseSoListMX = path.join(pathToBaseSoList, 'MX');

// NOTE - path to "to from\from to" dictionaries
export const pathToDictSoListCE = path.resolve(rootDir, 'public', 'CE_to_from.yaml');
export const pathToDictSoListMX = path.resolve(rootDir, 'public', 'MX_to_from.yaml');

export const pathToDictOutboundCE = path.join(rootDir, 'public', 'OUT_CE_to_from.yaml');
export const pathToDictOutboundMX = path.join(rootDir, 'public', 'OUT_MX_to_from.yaml');
