const CURRENCY_PRICE_COL = "Текущая цена";
const NEW_PRICE_COL = "Новая цена, RUB";
const QUANTITY_COL = "Количество";

let workbook = null;
let processedWorkbook = null;
let workbook2 = null;
let processedWorkbook2 = null;

// Загрузка и парсинг при выборе файла
document.getElementById("fileInput").addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      workbook = XLSX.read(ev.target.result, { type: "array" });
      processedWorkbook = null;
      alert("Файл загружен. Введите число и нажмите Calculate.");
    } catch (err) {
      alert("Ошибка чтения файла: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
});

// Расчёт и обновление цен
document.getElementById("calculateButton").addEventListener("click", () => {
  if (!workbook) {
    alert("Сначала загрузите файл");
    return;
  }
  const numberInput = document.getElementById("number");
  const value = parseFloat(numberInput.value);
  if (isNaN(value)) {
    alert("Введите число в поле");
    return;
  }
  const operator = document.getElementById("select").value;
  const delta = operator === "+" ? value : -value;

  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];

  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  if (aoa.length === 0) {
    alert("Таблица пуста");
    return;
  }

  const headers = aoa[0];
  const currentPriceIdx = headers.findIndex(
    h => h && String(h).trim() === CURRENCY_PRICE_COL,
  );
  const newPriceIdx = headers.findIndex(
    h => h && String(h).trim() === NEW_PRICE_COL,
  );

  if (currentPriceIdx === -1) {
    alert(`Колонка "${CURRENCY_PRICE_COL}" не найдена`);
    return;
  }

  const encodeCell = XLSX.utils.encode_cell;
  let newPriceColIdx = newPriceIdx;

  if (newPriceIdx === -1) {
    headers.push(NEW_PRICE_COL);
    newPriceColIdx = headers.length - 1;
    ws["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(aoa.length - 1, 0), c: newPriceColIdx },
    });
    const headerCellRef = encodeCell({ r: 0, c: newPriceColIdx });
    ws[headerCellRef] = { t: "s", v: NEW_PRICE_COL };
  }

  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r];
    const currentVal = row[currentPriceIdx];
    let newPrice;

    if (currentVal == null || currentVal === "") {
      newPrice = "";
    } else {
      const num = parseFloat(String(currentVal).replace(/\s/g, ""));
      if (isNaN(num)) {
        newPrice = currentVal;
      } else if (num === 0) {
        newPrice = 0;
      } else {
        newPrice = Math.round((num + delta) * 100) / 100;
      }
    }

    const cellRef = encodeCell({ r, c: newPriceColIdx });
    if (!ws[cellRef]) ws[cellRef] = {};
    ws[cellRef].v = newPrice;
    ws[cellRef].t = typeof newPrice === "number" ? "n" : "s";
  }

  processedWorkbook = workbook;
  alert("Расчёт выполнен. Нажмите Download для скачивания.");
});

// Скачивание обработанного файла
document.getElementById("downloadButton").addEventListener("click", () => {
  const wb = processedWorkbook || workbook;
  if (!wb) {
    alert("Сначала загрузите файл и выполните расчёт");
    return;
  }
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "prices_updated.xlsx";
  a.click();
  URL.revokeObjectURL(url);
});

// === Форма изменения количества ===
document.getElementById("fileInput2").addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      workbook2 = XLSX.read(ev.target.result, { type: "array" });
      processedWorkbook2 = null;
      alert("Файл загружен. Введите число и нажмите Рассчитать.");
    } catch (err) {
      alert("Ошибка чтения файла: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
});

document.getElementById("calculateButton2").addEventListener("click", () => {
  if (!workbook2) {
    alert("Сначала загрузите файл");
    return;
  }
  const numberInput = document.getElementById("number2");
  const value = parseInt(numberInput.value, 10);
  if (isNaN(value)) {
    alert("Введите число в поле");
    return;
  }
  const operator = document.getElementById("select2").value;
  const delta = operator === "+" ? value : -value;

  const sheetName = workbook2.SheetNames[0];
  const ws = workbook2.Sheets[sheetName];

  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  if (aoa.length === 0) {
    alert("Таблица пуста");
    return;
  }

  const headers = aoa[0];
  const quantityIdx = headers.findIndex(
    h => h && String(h).trim() === QUANTITY_COL,
  );

  if (quantityIdx === -1) {
    alert(`Колонка "${QUANTITY_COL}" не найдена`);
    return;
  }

  const encodeCell = XLSX.utils.encode_cell;

  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r];
    const currentVal = row[quantityIdx];
    let newQty;

    if (currentVal == null || currentVal === "") {
      newQty = "";
    } else {
      const num = parseInt(String(currentVal).replace(/\s/g, ""), 10);
      if (isNaN(num)) {
        newQty = currentVal;
      } else if (num === 0) {
        newQty = 0;
      } else {
        newQty = Math.max(0, num + delta);
      }
    }

    const cellRef = encodeCell({ r, c: quantityIdx });
    if (!ws[cellRef]) ws[cellRef] = {};
    ws[cellRef].v = newQty;
    ws[cellRef].t = typeof newQty === "number" ? "n" : "s";
  }

  processedWorkbook2 = workbook2;
  alert("Расчёт выполнен. Нажмите кнопку для скачивания.");
});

document.getElementById("downloadButton2").addEventListener("click", () => {
  const wb = processedWorkbook2 || workbook2;
  if (!wb) {
    alert("Сначала загрузите файл и выполните расчёт");
    return;
  }
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "quantity_updated.xlsx";
  a.click();
  URL.revokeObjectURL(url);
});
