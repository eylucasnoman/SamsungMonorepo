import path from 'path';
import fs from 'fs';

export function writeLog(logFileName: string, content: string) {
  const date = new Date().toLocaleString('pt-br');
  const logPath = path.resolve(__dirname, '../../logs');

  const fileExists = fs.readdirSync(logPath).includes(logFileName);
  const logFilePath = path.join(logPath, logFileName);

  const contentWithDate = `[${date}]: ${content}\r\n`;

  if (!fileExists) fs.writeFileSync(logFilePath, contentWithDate);
  else fs.appendFileSync(logFilePath, contentWithDate);
}
