// convert.js
import XLSX from "xlsx";
import fs from "fs";

const wb = XLSX.readFile("NTWordsonly(unicode).xls");
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws);

const cleaned = data.map(row => ({
  id: crypto.randomUUID(),
  greek: row.Word,
  fullWords: row["Full words"] || "",
  meaning: row["Meaning (English Gloss)"],
  type: row.Type || "",
  frequency: Number(row.Frequency || 0),
  difficulty: 50,
  known: 0,
  missed: 0
}));

fs.writeFileSync("nt_words.json", JSON.stringify(cleaned, null, 2));