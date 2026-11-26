const comparisonInput = document.getElementById("comparisonFile");
const truthInput = document.getElementById("truthFile");
const compareButton = document.getElementById("compareButton");
const exportButton = document.getElementById("exportButton");
const installButton = document.getElementById("installButton");
const message = document.getElementById("message");
const results = document.getElementById("results");
const resultsTitle = document.getElementById("resultsTitle");
const resultsBody = document.getElementById("resultsBody");
let differences = [];

function showMessage(text, kind) {
  message.textContent = text;
  message.className = kind;
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
  cell.textContent = value;
  if (className) cell.className = className;
  row.appendChild(cell);
}

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
  exportButton.disabled = true;
  results.hidden = true;
  showMessage("Lecture et comparaison des fichiers…", "success");

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
      addCell(row, item.status);
      resultsBody.appendChild(row);
    });

    if (differences.length === 0) {
      const row = document.createElement("tr");
      const cell = document.createElement("td");
      cell.colSpan = 6;
      cell.className = "empty";
      cell.textContent = "Aucun écart : les deux stocks correspondent.";
      row.appendChild(cell);
      resultsBody.appendChild(row);
    }

    resultsTitle.textContent = `${differences.length} référence(s) avec un écart sur ${references.size} référence(s) comparée(s).`;
    results.hidden = false;
    exportButton.disabled = differences.length === 0;
    showMessage("Comparaison terminée. Les quantités des références répétées ont été additionnées.", "success");
  } catch (error) {
    differences = [];
    showMessage(error.message, "error");
  } finally {
    compareButton.disabled = false;
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
