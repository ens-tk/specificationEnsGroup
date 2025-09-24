const rootInput = document.getElementById("rootFileInput");
const rootCounter = document.getElementById("rootFileCounter");
const fileTree = document.getElementById("fileTree");
const mergeBtn = document.getElementById("mergeBtn");

let rootFileNode = null;

function FileNode(file, name = null) {
  this.file = file;
  this.name = name || (file ? file.name : "");
  this.children = [];
  this.subassemblies = [];
  this.qty = 1;
  this.expanded = true; // 🔹 по умолчанию раскрыт
}


// Выбор корневого файла
rootInput.addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;

  rootFileNode = new FileNode(file);
  rootCounter.textContent = `Выбран файл: ${file.name}`;
  readSubAssemblies(file, rootFileNode).then(() => renderTree());
});

// Считываем сборочные единицы и их количество
async function readSubAssemblies(file, node) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const subassemblies = [];
  let foundSection = false;
  let skipNext = false;
  let colIndex = null;

  for (let row of rows) {
    if (!row) continue;

    if (!foundSection) {
      for (let i = 0; i < row.length; i++) {
        const cell = row[i];
        if (cell && cell.toString().toLowerCase().includes("сборочные единицы")) {
          foundSection = true;
          skipNext = true;
          colIndex = i;
          break;
        }
      }
      continue;
    }

    if (foundSection) {
      if (skipNext) { skipNext = false; continue; }
      if (row.every(c => !c || c.toString().trim() === "")) break;

      const name = row[colIndex] ? row[colIndex].toString().trim() : null;
const qty = row[colIndex + 1] 
  ? parseFloat(row[colIndex + 1].toString().replace(',', '.')) 
  : 1;


      if (name) {
        const nodeSub = new FileNode(null, name);
        nodeSub.qty = qty;
        subassemblies.push(nodeSub);
      }
    }
  }

  node.subassemblies = subassemblies;
}

// Отрисовка дерева
function renderTree() {
  fileTree.innerHTML = "";
  if (rootFileNode) fileTree.appendChild(createNodeElement(rootFileNode));
}

function createNodeElement(node) {
  const div = document.createElement("div");
  div.className = "file-node";

  const header = document.createElement("div");
  header.className = "node-header";

const toggle = document.createElement("span");
toggle.className = "toggle-btn";
toggle.textContent = node.subassemblies.length > 0
  ? (node.expanded ? "▼" : "▶")
  : "";

toggle.onclick = () => {
  node.expanded = !node.expanded; // 🔹 сохраняем состояние
  renderTree(); // перерисовываем
};

  const span = document.createElement("span");
  span.textContent = node.name;

  const uploadLabel = document.createElement("label");
  uploadLabel.className = "custom-btn";
  uploadLabel.style.cursor = "pointer";
  uploadLabel.textContent = node.file ? `✅ Файл загружен (${node.file.name})` : `Загрузить файл`;

  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".xls,.xlsx";
  input.style.display = "none";

  uploadLabel.onclick = () => input.click();
  input.addEventListener("change", e => {
    const file = e.target.files[0];
    if (file) {
      node.file = file;
      node.name = file.name;
      readSubAssemblies(file, node).then(() => renderTree());
    }
  });

  header.appendChild(toggle);
  header.appendChild(span);
  header.appendChild(uploadLabel);
  header.appendChild(input);

  div.appendChild(header);

const childrenDiv = document.createElement("div");
childrenDiv.className = "children";
childrenDiv.style.marginLeft = "20px";
childrenDiv.style.display = node.expanded ? "block" : "none"; // 🔹 учёт expanded
node.subassemblies.forEach(sub => childrenDiv.appendChild(createNodeElement(sub)));
div.appendChild(childrenDiv);

  return div;
}

// Сбор всех файлов
// Сбор всех файлов с multiplier
function collectFileNodes(node, currentMultiplier = 1) {
  let result = [];
  if (node && node.file) {
    result.push({
      file: node.file,
      multiplier: currentMultiplier * node.qty,
      name: node.name // чисто для логов
    });
  }
  node.subassemblies.forEach(sub => {
    result = result.concat(collectFileNodes(sub, currentMultiplier * node.qty));
  });
  return result;
}

mergeBtn.addEventListener("click", async () => {
  if (!rootFileNode) return alert("Выберите корневой файл");

  const allNodes = collectFileNodes(rootFileNode);
  const formData = new FormData();
  allNodes.forEach(n => {
    formData.append("files", n.file);
    formData.append("multipliers", n.multiplier);
    formData.append("names", n.name); // чисто для отладки
  });

  try {
    const res = await fetch("/upload", { method: "POST", body: formData });
    if (!res.ok) throw new Error("Ошибка при загрузке");

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "merged.xlsx";
    a.click();
    window.URL.revokeObjectURL(url);
  } catch (err) {
    alert(err.message);
  }
});
