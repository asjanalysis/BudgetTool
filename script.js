const budgetInput = document.getElementById("budgetFile");
const versionSelect = document.getElementById("version");
const loadButton = document.getElementById("loadBudget");
const generateButton = document.getElementById("generateReport");
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
  generateButton.disabled = true;
}

function renderTable(data) {
  if (!data.length) {
    clearTable();
    return;
  }

  expensesTableBody.innerHTML = "";
  attachments = data.map(() => ({ invoice: null, proof: null }));

  data.forEach((exp, index) => {
    const row = document.createElement("tr");
    const id = index + 1;

    const invoiceInputId = `invoice-${id}`;
    const proofInputId = `proof-${id}`;

    row.innerHTML = `
      <td>${id}</td>
      <td>${exp.name || "(unnamed)"}</td>
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

    expensesTableBody.appendChild(row);
  });

  generateButton.disabled = false;
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
  page.drawText(`Expense ${index}`, { x: 40, y: 780, size: 18, color: PDFLib.rgb(0.33, 0.84, 1) });
  page.drawText("Name", { x: 40, y: 740, size: 12, color: PDFLib.rgb(0.8, 0.86, 0.95) });
  page.drawText(exp.name || "(unnamed)", { x: 40, y: 720, size: 12, color: PDFLib.rgb(1, 1, 1) });
  page.drawText("Amount", { x: 40, y: 690, size: 12, color: PDFLib.rgb(0.8, 0.86, 0.95) });
  page.drawText(MONEY_FORMAT.format(exp.amount), { x: 40, y: 670, size: 12, color: PDFLib.rgb(1, 1, 1) });
  page.drawText("Sheet", { x: 40, y: 640, size: 12, color: PDFLib.rgb(0.8, 0.86, 0.95) });
  page.drawText(exp.sheet, { x: 40, y: 620, size: 12, color: PDFLib.rgb(1, 1, 1) });
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
    const version = Number(versionSelect.value);
    const data = await loadBudgetFromFile(file, version);
    expenses = data;
    renderTable(data);
  } catch (err) {
    alert(`Unable to read the spreadsheet: ${err.message}`);
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
