
import fs from 'fs';
import path from 'path';
import * as XLSX_Module from 'xlsx';
import { fileURLToPath } from 'url';

const XLSX = XLSX_Module.default || XLSX_Module;
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const RAW_DATA_DIR = path.join(__dirname, '../raw-data');

const file = 'Book1.xlsx';
const workbook = XLSX.readFile(path.join(RAW_DATA_DIR, file));
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

console.log(JSON.stringify(rows.slice(0, 30), null, 2));
