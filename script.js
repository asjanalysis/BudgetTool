const budgetInput = document.getElementById("budgetFile");
const versionSelect = document.getElementById("version");
const loadButton = document.getElementById("loadBudget");
const generateButton = document.getElementById("generateReport");
const savePointButton = document.getElementById("downloadSavePoint");
const expensesTableBody = document.querySelector("#expensesTable tbody");

let expenses = [];
let attachments = [];

const MONEY_FORMAT = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
});

function showMessage(text, tone = "info") {
  generateButton.textContent = tone === "loading" ? text : "Generate report PDF";
  generateButton.disabled = tone === "loading";
}

function parseAmount(raw) {
  if (raw === null || raw === undefined || raw === "") return 0;
  const valueStr = typeof raw === "string" ? raw : String(raw);

  let clean = valueStr.replace(/,/g, "").replace(/\$/g, "").trim();
  if (clean.startsWith("(") && clean.endsWith(")")) {
    clean = "-" + clean.slice(1, -1);
  }
  const value = parseFloat(clean);
  return Number.isFinite(value) ? value : 0;
}

function addRows(sheet, nameCols, amtCol, rows) {
  const list = [];
  for (let i = 6; i < rows.length; i++) {
    const row = rows[i] || [];
    const nameParts = [];
    for (let j = 0; j < nameCols; j++) {
      if (row[j]) nameParts.push(row[j]);
    }
    const amount = parseAmount(row[amtCol]);
    if (amount === 0) continue;
    list.push({
      name: nameParts.join(" - "),
      amount,
      sheet,
    });
  }
  return list;
}

function parseExpenseName(name) {
  const fallback = {
    category: "(category)",
    subCategory: "(sub-category)",
    phase: "(phase)",
    details: "(details)",
    original: name || "(unnamed)",
  };

  if (!name) return fallback;

  const parts = String(name)
    .split(" - ")
    .map((part) => part.trim())
    .filter(Boolean);

  const [category, subCategory, phase, ...rest] = parts;
  return {
    category: category || fallback.category,
    subCategory: subCategory || fallback.subCategory,
    phase: phase || fallback.phase,
    details: rest.length ? rest.join(" - ") : fallback.details,
    original: name,
  };
}

function renderExpenseName(name) {
  const parsed = parseExpenseName(name);
  return `
    <div class="expense-name">
      <span class="pill category">${parsed.category}</span>
      <span class="pill sub-category">${parsed.subCategory}</span>
      <span class="pill phase">Phase ${parsed.phase}</span>
      <span class="pill details">${parsed.details}</span>
    </div>
    <div class="expense-name-raw">${parsed.original}</div>
  `;
}

function addExpenseIds(list) {
  const counts = new Map();
  return list.map((exp) => {
    const base = `${exp.sheet}||${exp.name}||${exp.amount}`;
    const n = (counts.get(base) || 0) + 1;
    counts.set(base, n);
    return { ...exp, id: `${base}||${n}` };
  });
}

function isSavePointFile(file) {
  const name = (file?.name || "").toLowerCase();
  return name.endsWith(".btsp") || name.endsWith(".zip");
}

async function loadBudgetFromFile(file, version) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  const getRows = (sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    return sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1 }) : [];
  };

  const parsed = [];
  if (version === 2) {
    parsed.push(
      ...addRows("Personnel_Expenses", 6, 6, getRows("Personnel_Expenses")),
      ...addRows("NonPersonnel_Expenses", 3, 9, getRows("NonPersonnel_Expenses"))
    );
  } else {
    const expensesSheetName =
      workbook.SheetNames.find((name) => name.trim().toLowerCase() === "expenses") || "Expenses";
    const expenseSheet = workbook.Sheets[expensesSheetName];
    const expenseRows = expenseSheet
      ? XLSX.utils.sheet_to_json(expenseSheet, { header: 1, defval: "" })
      : [];
    const headerTargets = [
      "budget category",
      "sub-category",
      "project phase",
      "vendor",
      "item",
      "amount",
    ];
    let headerRowIndex = -1;
    let headerMap = {};

    for (let i = 0; i < expenseRows.length; i++) {
      const row = expenseRows[i] || [];
      const normalized = row.map((cell) => String(cell || "").trim().toLowerCase());
      const hits = headerTargets.filter((target) =>
        normalized.some((cell) => cell.includes(target))
      );
      if (hits.length >= 3) {
        headerRowIndex = i;
        headerMap = normalized.reduce((acc, cell, index) => {
          if (!cell) return acc;
          if (cell.includes("budget category")) acc.category = index;
          if (cell.includes("sub-category")) acc.subCategory = index;
          if (cell.includes("project phase")) acc.phase = index;
          if (cell.includes("vendor")) acc.vendor = index;
          if (cell.includes("item")) acc.item = index;
          if (cell.includes("invoice") && cell.includes("credit")) acc.invoice = index;
          if (cell.includes("invoice") && cell.includes("date")) acc.invoiceDate = index;
          if (cell.includes("transaction type")) acc.transaction = index;
          if (cell.includes("check") || cell.includes("voucher")) acc.check = index;
          if (cell.includes("amount")) acc.amount = index;
          return acc;
        }, {});
        break;
      }
    }

    const amountCol = Number.isInteger(headerMap.amount) ? headerMap.amount : 10;
    const startRow = headerRowIndex >= 0 ? headerRowIndex + 1 : 6;
    const list = [];

    for (let i = startRow; i < expenseRows.length; i++) {
      const row = expenseRows[i] || [];
      const amount = parseAmount(row[amountCol]);
      if (amount === 0) continue;

      const nameParts = [
        row[headerMap.category],
        row[headerMap.subCategory],
        row[headerMap.phase],
        row[headerMap.vendor],
        row[headerMap.item],
        row[headerMap.invoice],
        row[headerMap.invoiceDate],
        row[headerMap.transaction],
        row[headerMap.check],
      ]
        .map((value) => (value === null || value === undefined ? "" : String(value).trim()))
        .filter(Boolean);

      list.push({
        name: nameParts.join(" - "),
        amount,
        sheet: expensesSheetName,
      });
    }

    parsed.push(...list);
  }

  return parsed;
}

function clearTable() {
  expensesTableBody.innerHTML = `
    <tr class="empty">
      <td colspan="6">Load a budget spreadsheet to see expenses.</td>
    </tr>
  `;
  expenses = [];
  attachments = [];
  generateButton.disabled = true;
  if (savePointButton) savePointButton.disabled = true;
}

function renderTable(data, existingAttachments = null) {
  if (!data.length) {
    clearTable();
    return;
  }

  expensesTableBody.innerHTML = "";
  attachments = existingAttachments || data.map(() => ({ invoice: null, proof: null }));

  data.forEach((exp, index) => {
    const row = document.createElement("tr");
    const id = index + 1;

    const invoiceInputId = `invoice-${id}`;
    const proofInputId = `proof-${id}`;

    row.innerHTML = `
      <td>${id}</td>
      <td>${renderExpenseName(exp.name)}</td>
      <td>${MONEY_FORMAT.format(exp.amount)}</td>
      <td><span class="badge invoice">${exp.sheet}</span></td>
      <td>
        <input id="${invoiceInputId}" type="file" accept="application/pdf,image/png,image/jpeg" />
        <label for="${invoiceInputId}" class="button">Select invoice</label>
        <div class="hint" data-role="invoice-name"></div>
      </td>
      <td>
        <input id="${proofInputId}" type="file" accept="application/pdf,image/png,image/jpeg" />
        <label for="${proofInputId}" class="button">Select proof</label>
        <div class="hint" data-role="proof-name"></div>
      </td>
    `;

    row.querySelector(`#${invoiceInputId}`).addEventListener("change", (e) => {
      const file = e.target.files?.[0] || null;
      attachments[index].invoice = file;
      const label = row.querySelector('[data-role="invoice-name"]');
      label.textContent = file ? file.name : "No file";
    });

    row.querySelector(`#${proofInputId}`).addEventListener("change", (e) => {
      const file = e.target.files?.[0] || null;
      attachments[index].proof = file;
      const label = row.querySelector('[data-role="proof-name"]');
      label.textContent = file ? file.name : "No file";
    });

    const invLabel = row.querySelector('[data-role="invoice-name"]');
    if (invLabel) {
      invLabel.textContent = attachments[index]?.invoice ? attachments[index].invoice.name : "No file";
    }

    const proofLabel = row.querySelector('[data-role="proof-name"]');
    if (proofLabel) {
      proofLabel.textContent = attachments[index]?.proof ? attachments[index].proof.name : "No file";
    }

    expensesTableBody.appendChild(row);
  });

  generateButton.disabled = false;
  if (savePointButton) savePointButton.disabled = false;
}

function sanitizeFilename(name) {
  return String(name || "file").replace(/[^\w.\-() ]+/g, "_");
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

async function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

async function buildProgressPdfBytes() {
  const doc = await PDFLib.PDFDocument.create();

  for (let i = 0; i < expenses.length; i++) {
    const exp = expenses[i];
    const page = await addDetailPage(doc, i + 1, exp);

    const invoiceName = attachments[i]?.invoice?.name || "NOT UPLOADED";
    const proofName = attachments[i]?.proof?.name || "NOT UPLOADED";

    page.drawText(`Invoice: ${invoiceName}`, {
      x: 50,
      y: 90,
      size: 10,
      color: PDFLib.rgb(0.2, 0.2, 0.2),
    });

    page.drawText(`Proof: ${proofName}`, {
      x: 50,
      y: 75,
      size: 10,
      color: PDFLib.rgb(0.2, 0.2, 0.2),
    });
  }

  return await doc.save();
}

async function appendPdfAttachment(doc, file) {
  const bytes = await readFileAsArrayBuffer(file);
  const attachmentPdf = await PDFLib.PDFDocument.load(bytes);
  const pages = await doc.copyPages(attachmentPdf, attachmentPdf.getPageIndices());
  pages.forEach((p) => doc.addPage(p));
}

async function appendImageAttachment(doc, file, label) {
  const bytes = await readFileAsArrayBuffer(file);
  const page = doc.addPage([595.28, 841.89]); // A4
  const mime = file.type;

  let image;
  if (mime === "image/png") {
    image = await doc.embedPng(bytes);
  } else {
    image = await doc.embedJpg(bytes);
  }

  const { width, height } = image.scale(1);
  const maxWidth = 480;
  const maxHeight = 640;
  const scale = Math.min(maxWidth / width, maxHeight / height, 1);
  const scaled = image.scale(scale);

  page.drawText(label, {
    x: 40,
    y: 780,
    size: 14,
    color: PDFLib.rgb(1, 1, 1),
  });

  page.drawImage(image, {
    x: (595.28 - scaled.width) / 2,
    y: (841.89 - scaled.height) / 2,
    width: scaled.width,
    height: scaled.height,
  });
}

function addBlankPage(doc, message) {
  const page = doc.addPage([595.28, 841.89]);
  page.drawText(message, {
    x: 50,
    y: 760,
    size: 14,
    color: PDFLib.rgb(1, 1, 1),
  });
}

async function addDetailPage(doc, index, exp) {
  const page = doc.addPage([595.28, 841.89]);
  page.drawText(`Expense ${index}`, { x: 40, y: 780, size: 18, color: PDFLib.rgb(0.14, 0.52, 0.92) });
  const parsedName = parseExpenseName(exp.name);
  const yStart = 740;
  const lineGap = 26;

  const drawLine = (label, value, color, indexOffset = 0) => {
    const y = yStart - lineGap * indexOffset;
    page.drawText(label, { x: 40, y, size: 12, color: PDFLib.rgb(0.8, 0.86, 0.95) });
    page.drawText(value, { x: 140, y, size: 12, color });
  };

  drawLine("Category", parsedName.category, PDFLib.rgb(0.23, 0.51, 0.96), 0);
  drawLine("Sub-category", parsedName.subCategory, PDFLib.rgb(0.92, 0.28, 0.6), 1);
  drawLine("Phase", `Phase ${parsedName.phase}`, PDFLib.rgb(0.98, 0.45, 0.09), 2);
  drawLine("Details", parsedName.details, PDFLib.rgb(0.13, 0.77, 0.36), 3);
  drawLine("Amount", MONEY_FORMAT.format(exp.amount), PDFLib.rgb(0, 0.32, 0.71), 4);
  drawLine("Sheet", exp.sheet, PDFLib.rgb(0.2, 0.2, 0.2), 5);

  return page;
}

async function downloadSavePoint() {
  if (!expenses.length) return;
  if (typeof JSZip === "undefined") {
    alert("JSZip failed to load. Check index.html script tag.");
    return;
  }

  const oldText = savePointButton?.textContent;
  if (savePointButton) {
    savePointButton.disabled = true;
    savePointButton.textContent = "Packaging...";
  }

  try {
    const zip = new JSZip();
    const templateVersion = Number(versionSelect.value);

    const manifest = attachments.map((att, idx) => {
      const out = {};
      if (att?.invoice) {
        const safe = sanitizeFilename(att.invoice.name);
        out.invoice = {
          name: att.invoice.name,
          type: att.invoice.type || "application/octet-stream",
          path: `attachments/${idx + 1}/invoice_${safe}`,
        };
      }
      if (att?.proof) {
        const safe = sanitizeFilename(att.proof.name);
        out.proof = {
          name: att.proof.name,
          type: att.proof.type || "application/octet-stream",
          path: `attachments/${idx + 1}/proof_${safe}`,
        };
      }
      return out;
    });

    const state = {
      schemaVersion: 1,
      createdAt: new Date().toISOString(),
      templateVersion,
      expenses: expenses.map((exp) => ({
        id: exp.id,
        name: exp.name,
        amount: exp.amount,
        sheet: exp.sheet,
      })),
      attachments: manifest,
    };

    zip.file("state.json", JSON.stringify(state, null, 2));

    const progressPdfBytes = await buildProgressPdfBytes();
    zip.file("progress-report.pdf", progressPdfBytes);

    for (let i = 0; i < attachments.length; i++) {
      const inv = attachments[i]?.invoice;
      if (inv) {
        zip.file(manifest[i].invoice.path, await readFileAsArrayBuffer(inv));
      }

      const pr = attachments[i]?.proof;
      if (pr) {
        zip.file(manifest[i].proof.path, await readFileAsArrayBuffer(pr));
      }
    }

    const blob = await zip.generateAsync({ type: "blob" });
    const date = new Date().toISOString().slice(0, 10);
    downloadBlob(blob, `BudgetTool_SavePoint_${date}.btsp`);
  } catch (err) {
    console.error(err);
    alert(`Failed to create Save Point: ${err.message}`);
  } finally {
    if (savePointButton) {
      savePointButton.textContent = oldText;
      savePointButton.disabled = !expenses.length;
    }
  }
}

async function loadSavePoint(file) {
  if (typeof JSZip === "undefined") {
    throw new Error("JSZip is not available.");
  }

  const zipBytes = await readFileAsArrayBuffer(file);
  const zip = await JSZip.loadAsync(zipBytes);

  const stateText = await zip.file("state.json")?.async("string");
  if (!stateText) throw new Error("Save Point missing state.json");

  const state = JSON.parse(stateText);
  if (state.schemaVersion !== 1) {
    throw new Error(`Unsupported schemaVersion: ${state.schemaVersion}`);
  }

  if (state.templateVersion) {
    versionSelect.value = String(state.templateVersion);
  }

  expenses = (state.expenses || []).map((exp) => ({
    id: exp.id,
    name: exp.name,
    amount: exp.amount,
    sheet: exp.sheet,
  }));

  const restoredAttachments = expenses.map(() => ({ invoice: null, proof: null }));

  for (let i = 0; i < (state.attachments || []).length; i++) {
    const att = state.attachments[i] || {};

    if (att.invoice?.path) {
      const bytes = await zip.file(att.invoice.path).async("uint8array");
      restoredAttachments[i].invoice = new File([bytes], att.invoice.name, { type: att.invoice.type });
    }
    if (att.proof?.path) {
      const bytes = await zip.file(att.proof.path).async("uint8array");
      restoredAttachments[i].proof = new File([bytes], att.proof.name, { type: att.proof.type });
    }
  }

  renderTable(expenses, restoredAttachments);
}

async function generateReport() {
  if (!expenses.length) return;

  try {
    showMessage("Generating...", "loading");
    const doc = await PDFLib.PDFDocument.create();

    for (let i = 0; i < expenses.length; i++) {
      const exp = expenses[i];
      await addDetailPage(doc, i + 1, exp);

      const invoice = attachments[i]?.invoice;
      if (invoice) {
        if (invoice.type === "application/pdf") {
          await appendPdfAttachment(doc, invoice);
        } else {
          await appendImageAttachment(doc, invoice, `Invoice for expense ${i + 1}`);
        }
      } else {
        addBlankPage(doc, `Expense ${i + 1}: no invoice uploaded.`);
      }

      const proof = attachments[i]?.proof;
      if (proof) {
        if (proof.type === "application/pdf") {
          await appendPdfAttachment(doc, proof);
        } else {
          await appendImageAttachment(doc, proof, `Proof of payment for expense ${i + 1}`);
        }
      } else {
        addBlankPage(doc, `Expense ${i + 1}: no proof uploaded.`);
      }
    }

    const pdfBytes = await doc.save();
    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");
    link.href = url;
    link.download = "expense-report.pdf";
    link.click();
    URL.revokeObjectURL(url);
  } catch (err) {
    alert(`Something went wrong while creating the report: ${err.message}`);
    console.error(err);
  } finally {
    showMessage("Generate report PDF", "idle");
    generateButton.disabled = !expenses.length;
  }
}

loadButton.addEventListener("click", async () => {
  const file = budgetInput.files?.[0];
  if (!file) {
    alert("Please choose a budget spreadsheet first.");
    return;
  }

  try {
    loadButton.disabled = true;
    loadButton.textContent = "Loading...";

    if (isSavePointFile(file)) {
      await loadSavePoint(file);
    } else {
      const version = Number(versionSelect.value);
      const data = await loadBudgetFromFile(file, version);
      expenses = addExpenseIds(data);
      renderTable(expenses);
    }
  } catch (err) {
    alert(`Unable to load the file: ${err.message}`);
    console.error(err);
    clearTable();
  } finally {
    loadButton.disabled = false;
    loadButton.textContent = "Load expenses";
  }
});

budgetInput.addEventListener("change", () => {
  const file = budgetInput.files?.[0];
  if (file) {
    clearTable();
    const label = document.querySelector("label[for='budgetFile']");
    label.textContent = `Selected: ${file.name}`;
  }
});

if (savePointButton) {
  savePointButton.addEventListener("click", downloadSavePoint);
}

generateButton.addEventListener("click", generateReport);
