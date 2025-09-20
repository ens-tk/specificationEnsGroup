const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const port = 3000;

const upload = multer({ dest: "uploads/" });

app.use(express.static("public"));

app.post("/upload", upload.array("files", 20), (req, res) => {
    const files = req.files;
    if (!files || files.length === 0) return res.status(400).send("Файлы не загружены");

    const itemsMap = {};

    try {
        files.forEach(file => {
            const workbook = XLSX.readFile(file.path);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            let foundSection = false;
            let skipNext = false;
            let colIndex = null;

            for (let row of rows) {
                if (!row) continue;

                if (!foundSection) {
                    for (let i = 0; i < row.length; i++) {
                        const cell = row[i];
                        if (cell && cell.toString().toLowerCase().includes("стандартные изделия")) {
                            foundSection = true;
                            colIndex = i;
                            skipNext = true;
                            break;
                        }
                    }
                    continue;
                }

                if (foundSection) {
                    if (skipNext) {
                        skipNext = false;
                        continue;
                    }

                    if (row.every(cell => !cell || cell.toString().trim() === "")) break;

                    const name = row[colIndex] ? row[colIndex].toString().trim() : null;
                    let qty = row[colIndex + 1] ? parseFloat(row[colIndex + 1].toString().trim().replace(',', '.')) : null;

                    if (name && !isNaN(qty)) {
                        if (itemsMap[name]) itemsMap[name] += qty;
                        else itemsMap[name] = qty;
                    }
                }
            }

            fs.unlinkSync(file.path);
        });

        const newWorkbook = XLSX.utils.book_new();
        const newSheetData = [["Название изделия", "Количество"]];
        Object.keys(itemsMap).forEach(name => {
            newSheetData.push([name, itemsMap[name]]);
        });
        const newWorksheet = XLSX.utils.aoa_to_sheet(newSheetData);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Объединённые изделия");

        const outputFileName = `merged_${Date.now()}.xlsx`;
        const outputPath = path.join(__dirname, "uploads", outputFileName);
        XLSX.writeFile(newWorkbook, outputPath);

        res.download(outputPath, outputFileName, (err) => {
            if (err) console.error(err);
            fs.unlinkSync(outputPath);
        });

    } catch (err) {
        files.forEach(f => fs.unlinkSync(f.path));
        res.status(500).send("Ошибка при обработке файлов: " + err.message);
    }
});

app.listen(port, () => console.log(`Сервер запущен: http://localhost:${port}`));
