const comparisonInput = document.getElementById("comparisonFile");
const truthInput = document.getElementById("truthFile");
const compareButton = document.getElementById("compareButton");
const exportButton = document.getElementById("exportButton");
const installButton = document.getElementById("installButton");
const message = document.getElementById("message");
const results = document.getElementById("results");
const resultsBody = document.getElementById("resultsBody");
const actionHint = document.getElementById("actionHint");
const checkedCount = document.getElementById("checkedCount");
const differenceCount = document.getElementById("differenceCount");
const resultStatus = document.getElementById("resultStatus");
const resultNote = document.getElementById("resultNote");
let differences = [];

function showMessage(text, kind) {
  message.replaceChildren();
  if (!text) {
    message.className = "message";
    message.removeAttribute("aria-busy");
    return;
  }

  const icon = document.createElement("span");
  icon.className = "message-icon";
  icon.setAttribute("aria-hidden", "true");
  icon.textContent = kind === "error" ? "!" : kind === "loading" ? "…" : "✓";
  const content = document.createElement("span");
  content.textContent = text;
  message.append(icon, content);
  message.className = `message visible ${kind}`;
  message.setAttribute("aria-busy", String(kind === "loading"));
}

function updateFileState(input) {
  const file = input.files[0];
  const box = input.closest("[data-file-box]");
  const dropArea = input.closest(".file-drop");
  const name = dropArea.querySelector(".file-name");
  const detail = dropArea.querySelector(".file-detail");
  const state = dropArea.querySelector(".file-state");

  if (file && !/\.(xlsx|xls|csv)$/i.test(file.name)) {
    input.value = "";
    updateFileState(input);
    showMessage(`« ${file.name} » n’est pas un format accepté. Choisissez un fichier XLSX, XLS ou CSV.`, "error");
    return;
  }

  box.classList.remove("is-dragging");
  dropArea.classList.toggle("has-file", Boolean(file));
  if (file) {
    name.textContent = file.name;
    detail.textContent = `${(file.size / 1024).toLocaleString("fr-FR", { maximumFractionDigits: 0 })} Ko · cliquer pour remplacer`;
    state.textContent = "PRÊT";
  } else {
    name.textContent = "Choisir un fichier";
    detail.textContent = "ou déposer ici";
    state.textContent = "REQUIS";
  }

  differences = [];
  results.hidden = true;
  exportButton.disabled = true;
  resultsBody.replaceChildren();
  showMessage("", "");
  const ready = Boolean(comparisonInput.files[0] && truthInput.files[0]);
  compareButton.disabled = !ready;
  actionHint.textContent = ready
    ? "Vos fichiers sont prêts. Lancez la comparaison."
    : "Ajoutez les deux fichiers pour commencer.";

  if (file) {
    showMessage(ready
      ? "Les deux fichiers sont prêts. Vous pourrez vérifier les écarts avant l’export."
      : `Fichier chargé : ${file.name}. Ajoutez l’autre fichier pour continuer.`,
    "success");
  }
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error(`Impossible de lire « ${file.name} ».`));
    reader.onload = event => {
      try {
        const workbook = XLSX.read(event.target.result, { type: "array", cellDates: false });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        if (!firstSheet) throw new Error(`Le fichier « ${file.name} » ne contient aucune feuille.`);
        resolve(XLSX.utils.sheet_to_json(firstSheet, {
          header: 1,
          raw: false,
          defval: "",
          blankrows: false
        }));
      } catch (error) {
        reject(new Error(`Impossible de traiter « ${file.name} » : ${error.message}`));
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function parseStock(rows, fileName) {
  const stock = new Map();
  const errors = [];

  rows.forEach((row, index) => {
    const referenceValue = row[0] == null ? "" : String(row[0]).trim();
    const quantityValue = row[1] == null ? "" : String(row[1]).trim();
    if (!referenceValue && !quantityValue) return;

    if (index === 0 && !/^\d{7}$/.test(referenceValue) && !Number.isFinite(Number(quantityValue))) {
      return;
    }

    const reference = referenceValue.replace(/\.0$/, "");
    const quantity = Number(quantityValue.replace(",", "."));
    if (!/^\d{7}$/.test(reference)) {
      errors.push(`ligne ${index + 1} : référence « ${referenceValue} » invalide (7 chiffres requis)`);
      return;
    }
    if (!quantityValue || !Number.isFinite(quantity) || quantity < 0) {
      errors.push(`ligne ${index + 1} : quantité « ${quantityValue} » invalide (nombre positif ou nul requis)`);
      return;
    }
    stock.set(reference, (stock.get(reference) || 0) + quantity);
  });

  if (errors.length) {
    throw new Error(`${fileName} — ${errors.slice(0, 5).join("; ")}${errors.length > 5 ? `; et ${errors.length - 5} autre(s) erreur(s)` : ""}.`);
  }
  if (stock.size === 0) throw new Error(`${fileName} ne contient aucune ligne de stock valide.`);
  return stock;
}

function formatQuantity(value) {
  return new Intl.NumberFormat("fr-FR", { maximumFractionDigits: 3 }).format(value);
}

function addCell(row, value, className) {
  const cell = document.createElement("td");
  if (className === "status") {
    const pill = document.createElement("span");
    pill.className = "status-pill";
    pill.textContent = value;
    cell.appendChild(pill);
  } else {
    cell.textContent = value;
    if (className) cell.className = className;
  }
  row.appendChild(cell);
}

for (const input of [comparisonInput, truthInput]) {
  input.addEventListener("change", () => updateFileState(input));
}

document.querySelectorAll("[data-file-box]").forEach(box => {
  box.addEventListener("dragover", event => {
    event.preventDefault();
    box.classList.add("is-dragging");
  });
  box.addEventListener("dragleave", event => {
    if (!box.contains(event.relatedTarget)) box.classList.remove("is-dragging");
  });
  box.addEventListener("drop", event => {
    event.preventDefault();
    const input = box.querySelector('input[type="file"]');
    const file = event.dataTransfer.files[0];
    if (!file) return;
    const transfer = new DataTransfer();
    transfer.items.add(file);
    input.files = transfer.files;
    updateFileState(input);
  });
});

compareButton.addEventListener("click", async () => {
  const comparisonFile = comparisonInput.files[0];
  const truthFile = truthInput.files[0];
  if (!comparisonFile || !truthFile) {
    showMessage("Sélectionnez les deux fichiers avant de lancer la comparaison.", "error");
    return;
  }
  if (typeof XLSX === "undefined") {
    showMessage("Le composant de lecture Excel n’a pas pu être chargé. Vérifiez votre connexion Internet puis rechargez cette page.", "error");
    return;
  }

  compareButton.disabled = true;
  compareButton.querySelector("span").textContent = "Comparaison en cours…";
  exportButton.disabled = true;
  results.hidden = true;
  showMessage("Lecture des fichiers et calcul des écarts…", "loading");

  try {
    const [comparisonRows, truthRows] = await Promise.all([
      readWorkbook(comparisonFile),
      readWorkbook(truthFile)
    ]);
    const comparison = parseStock(comparisonRows, comparisonFile.name);
    const truth = parseStock(truthRows, truthFile.name);
    const references = new Set([...comparison.keys(), ...truth.keys()]);
    differences = [...references]
      .sort()
      .map(reference => {
        const comparisonQuantity = comparison.get(reference) || 0;
        const truthQuantity = truth.get(reference) || 0;
        return {
          reference,
          comparisonQuantity,
          truthQuantity,
          difference: comparisonQuantity - truthQuantity,
          adjustment: truthQuantity - comparisonQuantity,
          status: !comparison.has(reference)
            ? "Absente du stock à comparer"
            : !truth.has(reference)
              ? "Absente du fichier de vérité"
              : "Quantité différente"
        };
      })
      .filter(row => row.difference !== 0);

    resultsBody.replaceChildren();
    differences.forEach(item => {
      const row = document.createElement("tr");
      addCell(row, item.reference);
      addCell(row, formatQuantity(item.comparisonQuantity), "number");
      addCell(row, formatQuantity(item.truthQuantity), "number");
      addCell(row, formatQuantity(item.difference), "number");
      addCell(row, item.adjustment > 0 ? `+${formatQuantity(item.adjustment)}` : formatQuantity(item.adjustment), "number");
      addCell(row, item.status, "status");
      resultsBody.appendChild(row);
    });

    if (differences.length === 0) {
      const row = document.createElement("tr");
      const cell = document.createElement("td");
      cell.colSpan = 6;
      row.className = "empty-row";
      cell.textContent = "Aucun écart : les deux stocks correspondent.";
      row.appendChild(cell);
      resultsBody.appendChild(row);
    }

    checkedCount.textContent = references.size.toLocaleString("fr-FR");
    differenceCount.textContent = differences.length.toLocaleString("fr-FR");
    resultStatus.textContent = differences.length ? "À vérifier" : "Tout est juste";
    resultNote.textContent = differences.length ? "aperçu prêt à vérifier" : "aucun écart détecté";
    results.hidden = false;
    exportButton.disabled = differences.length === 0;
    showMessage(differences.length
      ? "Comparaison terminée. Vérifiez l’aperçu des écarts avant d’exporter. Les références répétées ont été additionnées."
      : "Comparaison terminée. Aucun écart : vos stocks correspondent. Les références répétées ont été additionnées.",
    "success");
    results.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (error) {
    differences = [];
    showMessage(error.message, "error");
  } finally {
    compareButton.disabled = false;
    compareButton.querySelector("span").textContent = "Comparer les stocks";
  }
});

exportButton.addEventListener("click", () => {
  if (!differences.length || typeof XLSX === "undefined") return;
  const rows = differences.map(item => ({
    "Référence": item.reference,
    "Stock à comparer": item.comparisonQuantity,
    "Stock vérité": item.truthQuantity,
    "Écart (comparé - vérité)": item.difference,
    "Ajustement recommandé (vérité - comparé)": item.adjustment,
    "État": item.status
  }));
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(rows);
  worksheet["!cols"] = [
    { wch: 14 }, { wch: 18 }, { wch: 15 },
    { wch: 24 }, { wch: 38 }, { wch: 30 }
  ];
  XLSX.utils.book_append_sheet(workbook, worksheet, "Différences");
  XLSX.writeFile(workbook, "differences_stock.xlsx");
});

let installPrompt;

window.addEventListener("beforeinstallprompt", event => {
  event.preventDefault();
  installPrompt = event;
  installButton.hidden = false;
});

installButton.addEventListener("click", async () => {
  if (!installPrompt) return;
  await installPrompt.prompt();
  installPrompt = null;
  installButton.hidden = true;
});

window.addEventListener("appinstalled", () => {
  installButton.hidden = true;
});

if ("serviceWorker" in navigator && location.protocol !== "file:") {
  navigator.serviceWorker.register("./sw.js").catch(error => {
    console.error("Échec de l’activation du mode hors ligne :", error);
    showMessage("Le mode hors ligne n’a pas pu être activé. Rechargez la page via une connexion sécurisée (HTTPS) ou contactez la personne qui héberge l’application.", "error");
  });
}
