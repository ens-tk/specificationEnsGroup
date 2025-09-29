const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const port = 3000;

const upload = multer({ dest: "uploads/" });
app.use(express.static("public"));

function processWorkbook(filePath, multiplier = 1, itemsMap = {}, fileMap = {}, parentChain = [], relations = [], fileName = null) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const currentNode = { name: fileName, qty: multiplier };
    const newParentChain = [...parentChain, currentNode];

    console.log(`Processing: ${currentNode.name}, multiplier=${multiplier}`);

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
                const clean = str => str.toString().trim().toLowerCase();
                const subFileKey = Object.keys(fileMap).find(f => clean(f) === clean(name));

                if (subFileKey) {
                    processWorkbook(
                        fileMap[subFileKey].path,
                        multiplier * qty,
                        itemsMap,
                        fileMap,
                        newParentChain,
                        relations,
                        subFileKey
                    );
                }
            }
        }
    }

    processSection(rows, "стандартные изделия", multiplier, itemsMap, newParentChain, relations);
    processSection(rows, "прочие изделия", multiplier, itemsMap, newParentChain, relations);

    return { itemsMap, relations };
}

function processSection(rows, keyword, multiplier, itemsMap, parentChain, relations) {
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
                    break;
                }
            }
            continue;
        }

        if (found) {
            if (skipNext) { skipNext = false; continue; }
            if (row.every(c => !c || c.toString().trim() === "")) break;

            let itemName = row[colIndex] ? row[colIndex].toString().trim() : null;
            if (itemName && colIndex > 0 && row[colIndex - 1]) {
                itemName = itemName + "_" + row[colIndex - 1].toString().trim();
            }

            const itemQty = row[colIndex + 1] ? parseFloat(row[colIndex + 1].toString().replace(',', '.')) : 1;

            if (itemName && !isNaN(itemQty)) {
                const totalItemQty = itemQty * multiplier;

                if (itemsMap[itemName]) itemsMap[itemName] += totalItemQty;
                else itemsMap[itemName] = totalItemQty;

                const childNode = parentChain[parentChain.length - 1];
                const parentNode = parentChain.length >= 2 
                    ? parentChain[parentChain.length - 2] 
                    : null;

                relations.push([
                    parentNode ? parentNode.name : "",
                    parentNode ? parentNode.qty : 1,
                    childNode.name,
                    childNode.qty,
                    itemName,
                    totalItemQty
                ]);
            }
        }
    }
}

app.post("/upload", upload.array("files", 50), (req, res) => {
    const files = req.files;
    let multipliers = req.body.multipliers || [];
    let names = req.body.names || [];
    let parents = req.body.parents || [];
    let parentQtys = req.body.parentQtys || [];

    if (!Array.isArray(multipliers)) multipliers = [multipliers];
    if (!Array.isArray(names)) names = [names];
    if (!Array.isArray(parents)) parents = [parents];
    if (!Array.isArray(parentQtys)) parentQtys = [parentQtys];

    if (!files || files.length === 0) return res.status(400).send("No files uploaded");

    const itemsMap = {};
    const relations = [];
    const fileMap = {};

    files.forEach((f, i) => {
        const originalName = names[i] || f.originalname;
        fileMap[originalName] = f;
    });

    try {
        files.forEach((f, i) => {
            const multiplier = parseFloat(multipliers[i]) || 1;
            const fileName = names[i] || f.originalname;
            const parentName = parents[i] || "";
            const parentQty = parseFloat(parentQtys[i]) || 1;

            const initialParentChain = parentName 
                ? [{ name: parentName, qty: parentQty }] 
                : [];

            const { itemsMap: im, relations: rels } = processWorkbook(
                f.path,
                multiplier,
                itemsMap,
                fileMap,
                initialParentChain,
                [],
                fileName
            );

            relations.push([
                parentName,
                parentQty,
                fileName,
                multiplier,
                "",
                multiplier
            ]);

            rels.forEach(r => relations.push(r));
        });

        const newWorkbook = XLSX.utils.book_new();

        const allSpecs = new Set();
        const specQuantities = {};

        relations.forEach(r => {
            const itemName = r[4];
            const specName = r[2];
            const qty = r[5];

            if (!itemName) return;

            allSpecs.add(specName);

            if (!specQuantities[itemName]) specQuantities[itemName] = {};
            if (!specQuantities[itemName][specName]) specQuantities[itemName][specName] = 0;

            specQuantities[itemName][specName] += qty;
        });

        const specList = Array.from(allSpecs);
        const header = ["Item Name", "Total Quantity", ...specList];
        const data = [header];

        Object.keys(specQuantities).forEach(itemName => {
            const row = [];
            row.push(itemName);

            const totalQty = specList.reduce((sum, spec) => sum + (specQuantities[itemName][spec] || 0), 0);
            row.push(totalQty);

            specList.forEach(spec => row.push(specQuantities[itemName][spec] || 0));
            data.push(row);
        });

        const ws1 = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(newWorkbook, ws1, "Merged Items");

        const ws2 = XLSX.utils.aoa_to_sheet([["Parent","ParentQty","Child","ChildQty","Item","ItemQty"]]);
        XLSX.utils.sheet_add_aoa(ws2, relations, { origin: -1 });
        XLSX.utils.book_append_sheet(newWorkbook, ws2, "Relations");

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
