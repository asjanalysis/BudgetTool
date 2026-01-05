const budgetInput = document.getElementById("budgetFile");
const versionSelect = document.getElementById("version");
const loadButton = document.getElementById("loadBudget");
const saveButton = document.getElementById("saveProgress");
const generateButton = document.getElementById("generateReport");
const expensesTableBody = document.querySelector("#expensesTable tbody");

const SAVE_ATTACHMENT_NAME = "budget-tool-save.json";

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

function updateActionAvailability() {
  const hasExpenses = Array.isArray(expenses) && expenses.length > 0;
  generateButton.disabled = !hasExpenses;
  saveButton.disabled = !hasExpenses;
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
    parsed.push(...addRows("Expenses", 10, 10, getRows("Expenses")));
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
  updateActionAvailability();
}

function renderTable(data, existingAttachments = []) {
  if (!data.length) {
    clearTable();
    return;
  }

  expensesTableBody.innerHTML = "";
  attachments = data.map((_, index) => existingAttachments[index] || { invoice: null, proof: null });

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

    const invoiceLabel = row.querySelector('[data-role="invoice-name"]');
    const proofLabel = row.querySelector('[data-role="proof-name"]');
    const currentInvoice = attachments[index].invoice;
    const currentProof = attachments[index].proof;
    invoiceLabel.textContent = currentInvoice ? currentInvoice.name : "No file";
    proofLabel.textContent = currentProof ? currentProof.name : "No file";

    expensesTableBody.appendChild(row);
  });

  updateActionAvailability();
}

async function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
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
}

async function createReportDocument() {
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

  return doc;
}

async function downloadPdf(doc, filename) {
  const pdfBytes = await doc.save();
  const blob = new Blob([pdfBytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

async function serializeAttachment(file) {
  if (!file) return null;
  const buffer = await readFileAsArrayBuffer(file);
  return {
    name: file.name,
    type: file.type || "application/octet-stream",
    data: arrayBufferToBase64(buffer),
  };
}

function base64ToFile(b64, name, type) {
  if (!b64) return null;
  const binary = atob(b64);
  const length = binary.length;
  const bytes = new Uint8Array(length);
  for (let i = 0; i < length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  const blob = new Blob([bytes], { type: type || "application/octet-stream" });
  try {
    return new File([blob], name, { type: blob.type });
  } catch (err) {
    // Older browsers may not support File constructor
    blob.name = name;
    return blob;
  }
}

async function buildSaveMetadata() {
  const serializedAttachments = [];
  for (const pair of attachments) {
    serializedAttachments.push({
      invoice: await serializeAttachment(pair.invoice),
      proof: await serializeAttachment(pair.proof),
    });
  }

  return {
    kind: "BudgetToolSave",
    schemaVersion: 1,
    templateVersion: Number(versionSelect.value),
    budgetFileName: budgetInput.files?.[0]?.name || "",
    expenses,
    attachments: serializedAttachments,
  };
}

async function attachSaveData(doc, metadata) {
  const encoder = new TextEncoder();
  const payload = encoder.encode(JSON.stringify(metadata));

  // Store a lightweight copy of the save data in the PDF metadata as a fallback
  // for environments that cannot read attachments. Attachments still carry the
  // full save payload (including encoded files) to restore uploads when
  // reloading progress.
  const lightMetadata = {
    kind: metadata.kind,
    schemaVersion: metadata.schemaVersion,
    templateVersion: metadata.templateVersion,
    budgetFileName: metadata.budgetFileName,
    expenses: metadata.expenses,
  };

  doc.setSubject(JSON.stringify(lightMetadata));
  doc.setTitle("Budget Tool progress save");
  doc.attach(payload, SAVE_ATTACHMENT_NAME, {
    mimeType: "application/json",
    description: "Budget Tool save data",
  });
}

async function generateReport() {
  if (!expenses.length) return;

  try {
    showMessage("Generating...", "loading");
    const doc = await createReportDocument();
    await downloadPdf(doc, "expense-report.pdf");
  } catch (err) {
    alert(`Something went wrong while creating the report: ${err.message}`);
    console.error(err);
  } finally {
    showMessage("Generate report PDF", "idle");
    updateActionAvailability();
  }
}

async function saveProgress() {
  if (!expenses.length) return;
  saveButton.textContent = "Saving...";
  saveButton.disabled = true;

  try {
    const doc = await createReportDocument();
    const metadata = await buildSaveMetadata();
    await attachSaveData(doc, metadata);
    await downloadPdf(doc, "budget-progress.pdf");
  } catch (err) {
    alert(`Unable to save your progress: ${err.message}`);
    console.error(err);
  } finally {
    saveButton.textContent = "Save progress PDF";
    updateActionAvailability();
  }
}

async function loadSavedProgress(file) {
  const buffer = await readFileAsArrayBuffer(file);
  const doc = await PDFLib.PDFDocument.load(buffer);

  const attachments = typeof doc.getAttachments === "function" ? doc.getAttachments() : [];
  const saveAttachment = attachments?.find((att) => att.name === SAVE_ATTACHMENT_NAME);
  const subject = doc.getSubject();

  if (!saveAttachment && !subject) {
    throw new Error("This PDF does not contain Budget Tool save data.");
  }

  let payload;
  try {
    if (saveAttachment) {
      const decoder = new TextDecoder();
      const json = decoder.decode(saveAttachment.content || saveAttachment.data || saveAttachment.bytes);
      payload = JSON.parse(json);
    } else {
      payload = JSON.parse(subject);
    }
  } catch (err) {
    throw new Error("Unable to read save metadata from PDF.");
  }

  if (payload.kind !== "BudgetToolSave") {
    throw new Error("The selected PDF is not a Budget Tool progress file.");
  }

  const restoredAttachments = (payload.attachments || []).map((pair) => ({
    invoice: pair?.invoice ? base64ToFile(pair.invoice.data, pair.invoice.name, pair.invoice.type) : null,
    proof: pair?.proof ? base64ToFile(pair.proof.data, pair.proof.name, pair.proof.type) : null,
  }));

  versionSelect.value = String(payload.templateVersion || 1);
  expenses = payload.expenses || [];
  renderTable(expenses, restoredAttachments);
  updateActionAvailability();
}

loadButton.addEventListener("click", async () => {
  const file = budgetInput.files?.[0];
  if (!file) {
    alert("Please choose a budget spreadsheet or saved progress PDF first.");
    return;
  }

  try {
    loadButton.disabled = true;
    loadButton.textContent = "Loading...";
    if (file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) {
      await loadSavedProgress(file);
    } else {
      const version = Number(versionSelect.value);
      const data = await loadBudgetFromFile(file, version);
      expenses = data;
      renderTable(data);
      updateActionAvailability();
    }
  } catch (err) {
    alert(`Unable to read the file: ${err.message}`);
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

generateButton.addEventListener("click", generateReport);
saveButton.addEventListener("click", saveProgress);
