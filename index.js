const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const port = 3000;

const upload = multer({ dest: "uploads/" });
app.use(express.static("public"));

function processWorkbook(filePath, multiplier = 1, itemsMap = {}, fileMap = {}) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    console.log(`\n📄 Processing file: ${path.basename(filePath)} | multiplier = ${multiplier}`);

    // -------------------------------
    // 🔹 Сборочные единицы
    let foundSubassemblies = false;
    let subColIndex = null;
    let skipNextSub = false;

    for (let row of rows) {
        if (!row) continue;

        if (!foundSubassemblies) {
            for (let i = 0; i < row.length; i++) {
                const cell = row[i];
                if (cell && cell.toString().toLowerCase().includes("сборочные единицы")) {
                    foundSubassemblies = true;
                    subColIndex = i;
                    skipNextSub = true;
                    console.log(`🔹 Found "Subassemblies" section in column ${subColIndex}`);
                    break;
                }
            }
            continue;
        }

        if (foundSubassemblies) {
            if (skipNextSub) { skipNextSub = false; continue; }
            if (row.every(c => !c || c.toString().trim() === "")) break;

            const name = row[subColIndex] ? row[subColIndex].toString().trim() : null;
            const qty = row[subColIndex + 1] ? parseFloat(row[subColIndex + 1].toString().replace(',', '.')) : 1;

            if (name && !isNaN(qty)) {
                console.log(`🔹 Subassembly: "${name}", qty = ${qty} | multiplier = ${multiplier}`);
                const clean = str => str.toString().trim().toLowerCase();
                const subFileKey = Object.keys(fileMap).find(f => clean(f) === clean(name));

                if (subFileKey) {
                    processWorkbook(fileMap[subFileKey].path, multiplier * qty, itemsMap, fileMap);
                } else {
                    console.log("❌ File not found among uploaded files. Searching keys:");
                    Object.keys(fileMap).forEach(f => console.log("   -", f));
                }
            }
        }
    }

    // -------------------------------
    // 🔹 Стандартные изделия
    processSection(rows, "стандартные изделия", multiplier, itemsMap);

    // -------------------------------
    // 🔹 Прочие изделия
    processSection(rows, "прочие изделия", multiplier, itemsMap);

    return itemsMap;
}

// 🔧 Универсальная функция для сбора изделий из секции
// 🔧 Универсальная функция для сбора изделий из секции
function processSection(rows, keyword, multiplier, itemsMap) {
    let found = false;
    let colIndex = null;
    let skipNext = false;

    for (let row of rows) {
        if (!row) continue;

        if (!found) {
            for (let i = 0; i < row.length; i++) {
                const cell = row[i];
                if (cell && cell.toString().toLowerCase().includes(keyword)) {
                    found = true;
                    colIndex = i;
                    skipNext = true;
                    console.log(`🔹 Found "${keyword}" section in column ${colIndex}`);
                    break;
                }
            }
            continue;
        }

        if (found) {
            if (skipNext) { skipNext = false; continue; }
            if (row.every(c => !c || c.toString().trim() === "")) break;

            // Основное имя (в колонке с названием изделия)
            let name = row[colIndex] ? row[colIndex].toString().trim() : null;

            // Если в ЛЕВОЙ ячейке есть данные → добавить их через "_"
            if (name && colIndex > 0 && row[colIndex - 1]) {
                name = name + "_" + row[colIndex - 1].toString().trim();
            }

            const qty = row[colIndex + 1] 
                ? parseFloat(row[colIndex + 1].toString().replace(',', '.')) 
                : 1;

            if (name && !isNaN(qty)) {
                const totalQty = qty * multiplier;
                console.log(`✅ ${keyword}: "${name}", qty = ${qty}, multiplier = ${multiplier}, total = ${totalQty}`);
                if (itemsMap[name]) itemsMap[name] += totalQty;
                else itemsMap[name] = totalQty;
            }
        }
    }
}




app.post("/upload", upload.array("files", 50), (req, res) => {
  const files = req.files;
  let multipliers = req.body.multipliers || [];
  let names = req.body.names || [];

  if (!Array.isArray(multipliers)) multipliers = [multipliers];
  if (!Array.isArray(names)) names = [names];
  if (!files || files.length === 0) return res.status(400).send("No files uploaded");

  const itemsMap = {};

  try {
    files.forEach((f, i) => {
      const multiplier = parseFloat(multipliers[i]) || 1;
      console.log(`📂 Root file: ${names[i] || f.originalname}, multiplier = ${multiplier}`);
      processWorkbook(f.path, multiplier, itemsMap, {}); 
    });

    const newWorkbook = XLSX.utils.book_new();
    const data = [["Item Name", "Quantity"]];
    Object.keys(itemsMap).forEach(name => data.push([name, itemsMap[name]]));
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(newWorkbook, ws, "Merged Items");

    const outFile = `merged_${Date.now()}.xlsx`;
    const outPath = path.join(__dirname, "uploads", outFile);
    XLSX.writeFile(newWorkbook, outPath);

    res.download(outPath, outFile, err => {
      if (err) console.error(err);
      files.forEach(f => fs.unlinkSync(f.path));
      fs.unlinkSync(outPath);
    });
  } catch (err) {
    files.forEach(f => fs.unlinkSync(f.path));
    res.status(500).send("Error: " + err.message);
  }
});


app.listen(port, () => console.log(`Server running at http://localhost:${port}`));
