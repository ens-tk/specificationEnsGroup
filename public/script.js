const fileInput = document.getElementById("fileInput");
const fileList = document.getElementById("fileList");
const fileCounter = document.getElementById("fileCounter");

let selectedFiles = [];

fileInput.addEventListener("change", (e) => {
  for (let file of e.target.files) {
    selectedFiles.push(file);
  }
  updateFileList();
  fileInput.value = "";
});

function updateFileList() {
  fileList.innerHTML = "";

  if (selectedFiles.length === 0) {
    fileCounter.textContent = "Файл не выбран";
  } else {
    fileCounter.textContent = `Выбрано файлов: ${selectedFiles.length}`;
  }

  selectedFiles.forEach((file, index) => {
    const div = document.createElement("div");
    div.classList.add("file-item");

    const span = document.createElement("span");
    span.textContent = file.name;

    const btn = document.createElement("button");
    btn.textContent = "Удалить";
    btn.classList.add("remove-btn");
    btn.onclick = (e) => {
      e.preventDefault();
      selectedFiles.splice(index, 1);
      updateFileList();
    };

    div.appendChild(span);
    div.appendChild(btn);
    fileList.appendChild(div);
  });
}

document.getElementById("uploadForm").addEventListener("submit", (e) => {
  e.preventDefault();

  if (selectedFiles.length === 0) {
    alert("Выберите хотя бы один файл!");
    return;
  }

  const formData = new FormData();
  selectedFiles.forEach(file => {
    formData.append("files", file);
  });

  fetch("/upload", {
    method: "POST",
    body: formData
  })
  .then(res => {
    if (!res.ok) throw new Error("Ошибка при загрузке");
    return res.blob();
  })
  .then(blob => {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "merged.xlsx";
    a.click();
    window.URL.revokeObjectURL(url);
  })
  .catch(err => alert(err.message));
});
