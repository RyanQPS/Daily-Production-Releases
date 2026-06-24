#!/usr/bin/env node
// QPS Production Release Routine — spreadsheet helper.
//
// Usage:
//   node releases.mjs init
//   node releases.mjs read
//   node releases.mjs replace <json-file>

import { readFileSync, existsSync } from 'node:fs';
import { resolve } from 'node:path';
import xlsx from 'xlsx';

const FILE = resolve(process.cwd(), 'releases.xlsx');
const SHEET = 'Releases';
const COLUMNS = ['Date', 'Type', 'Project', 'Sender', 'Subject', 'MessageId'];

function readSheet() {
  if (!existsSync(FILE)) return [];
  const wb = xlsx.readFile(FILE);
  const ws = wb.Sheets[SHEET] ?? wb.Sheets[wb.SheetNames[0]];
  if (!ws) return [];
  return xlsx.utils.sheet_to_json(ws, { defval: '' });
}

function writeSheet(rows) {
  const normalized = rows.map((r) => {
    const out = {};
    for (const c of COLUMNS) out[c] = r[c] ?? '';
    return out;
  });
  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.json_to_sheet(normalized, { header: COLUMNS });
  ws['!cols'] = [
    { wch: 12 }, { wch: 8 }, { wch: 40 },
    { wch: 24 }, { wch: 60 }, { wch: 40 },
  ];
  xlsx.utils.book_append_sheet(wb, ws, SHEET);
  xlsx.writeFile(wb, FILE);
}

const [, , cmd, arg] = process.argv;

switch (cmd) {
  case 'init': {
    if (existsSync(FILE)) {
      console.log(`exists: ${FILE}`);
    } else {
      writeSheet([]);
      console.log(`created: ${FILE}`);
    }
    break;
  }
  case 'read': {
    process.stdout.write(JSON.stringify(readSheet(), null, 2) + '\n');
    break;
  }
  case 'replace': {
    if (!arg) {
      console.error('usage: node releases.mjs replace <json-file>');
      process.exit(2);
    }
    const rows = JSON.parse(readFileSync(arg, 'utf8'));
    if (!Array.isArray(rows)) {
      console.error('input must be a JSON array of row objects');
      process.exit(2);
    }
    writeSheet(rows);
    console.log(`wrote ${rows.length} rows to ${FILE}`);
    break;
  }
  default: {
    console.error('usage: node releases.mjs [init|read|replace <json-file>]');
    process.exit(2);
  }
}
