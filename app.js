/* ============================================================
   Générateur Snake Excel par FCI (Version absolue)
   ============================================================ */

const state = {
  csvFile: null,
  zip: null,
  zipBlob: null,
  modeleBuffer: null,
};

const $ = id => document.getElementById(id);

/* ---------- Pré-chargement du modèle ---------- */
window.addEventListener("DOMContentLoaded", () => {
  try {
    if (typeof getModeleArrayBuffer !== "function") {
      throw new Error("modele-b64.js manquant ou invalide");
    }
    state.modeleBuffer = getModeleArrayBuffer();
    const sizeKB = getModeleSizeKB();
    const info = $("modele-info");
    if (info) {
      info.innerHTML = `✓ Modèle intégré chargé (${sizeKB} Ko) — feuille <code>SNK_1</code> prête`;
      info.style.color = "#2E7D32";
    }
  } catch (err) {
    console.error(err);
    const info = $("modele-info");
    if (info) {
      info.textContent = "❌ Impossible de décoder le modèle intégré : " + err.message;
      info.style.color = "#f44336";
    }
  }
  checkReady();
});

/* ---------- Gestion CSV : input + drag & drop ---------- */
const dropZone = $("drop-zone");
const fileInput = $("file-csv");
const csvInfo = $("csv-info");

dropZone.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", e => handleCsvFile(e.target.files[0]));

["dragenter", "dragover"].forEach(evt => {
  dropZone.addEventListener(evt, e => {
    e.preventDefault(); e.stopPropagation();
    dropZone.style.background = "#C8E6C9";
  });
});

["dragleave", "drop"].forEach(evt => {
  dropZone.addEventListener(evt, e => {
    e.preventDefault(); e.stopPropagation();
    dropZone.style.background = "#F1F8E9";
  });
});

dropZone.addEventListener("drop", e => {
  const files = e.dataTransfer.files;
  if (files && files.length) handleCsvFile(files[0]);
});

function handleCsvFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".csv")) {
    csvInfo.textContent = "❌ Le fichier doit être un .csv";
    csvInfo.style.color = "#f44336";
    return;
  }
  state.csvFile = file;
  csvInfo.textContent = `✓ ${file.name} (${(file.size / 1024).toFixed(1)} Ko)`;
  csvInfo.style.color = "#2E7D32";
  checkReady();
}

function checkReady() {
  $("btn-generer").disabled = !(state.csvFile && state.modeleBuffer);
}

function log(msg, cls = "") {
  const box = $("log");
  if (!box) { console.log(msg); return; }
  const line = document.createElement("div");
  if (cls) line.className = cls;
  line.textContent = msg;
  box.appendChild(line);
  box.scrollTop = box.scrollHeight;
}

function setProgress(pct) {
  const fill = $("progress-fill");
  if (fill) fill.style.width = pct + "%";
}

/* ============================================================
   UTILITAIRES
   ============================================================ */
const s = v => (v == null ? "" : String(v).trim());

function safeInt(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = parseInt(parseFloat(String(v).trim()));
  return isNaN(n) ? null : n;
}

function sanitizeFilename(name) {
  return String(name).replace(/[<>:"/\\|?*]/g, "_");
}

function normaliser(txt) {
  return s(txt).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function trouverColonne(entetes, ...candidats) {
  const table = {};
  entetes.forEach(h => { if (h) table[normaliser(h)] = h; });
  for (const cand of candidats) {
    const k = normaliser(cand);
    if (table[k] !== undefined) return table[k];
  }
  return null;
}

function trouverColonneFuzzy(entetes, ...mots) {
  for (const h of entetes) {
    if (!h) continue;
    const hn = normaliser(h);
    for (const m of mots) if (hn.includes(normaliser(m))) return h;
  }
  return null;
}

function splitMasque(val) {
  const m = s(val).match(/^([A-Za-z]+)\s*(\d+)?$/);
  if (!m) return [s(val), null];
  return [m[1], m[2] ? parseInt(m[2]) : null];
}

/* ============================================================
   PARSING CSV
   ============================================================ */
function parseCSVFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      delimiter: ";",
      skipEmptyLines: true,
      transformHeader: h => h.trim().replace(/^\uFEFF/, ""),
      complete: r => resolve(r.data),
      error: reject,
    });
  });
}

/* ============================================================
   GROUPEMENT PAR FCI
   ============================================================ */
function preparerParFCI(rows, reverse = false) {
  if (!rows || !rows.length) throw new Error("CSV vide ou invalide");
  const cols = Object.keys(rows[0]);

  const cOp      = trouverColonne(cols, "OP");
  const cFci     = trouverColonne(cols, "FCI", "N°_FCI", "N_FCI") || trouverColonneFuzzy(cols, "fci");
  const cCh1     = trouverColonne(cols, "CODE_CH1", "CODE CH1");
  const cCh2     = trouverColonne(cols, "CODE_CH2", "INSEE");
  const cRef     = trouverColonne(cols, "REF_CHAMBR");
  const cBpe     = trouverColonne(cols, "BPE");
  const cMasH    = trouverColonne(cols, "MASQUE", "Masque H");
  const cMasB    = trouverColonne(cols, "MASQUE B", "Masque B");
  const cDiaH    = trouverColonne(cols, "DIAMETRE", "Diamètre H");
  const cDiaB    = trouverColonne(cols, "DIAMETRE B", "Diamètre B");
  const cEtiq2   = trouverColonneFuzzy(cols, "etiquette");
  const cOrdre   = trouverColonne(cols, "ORDRE", "Ordre");
  const cCommune = trouverColonne(cols, "COMMUNE");

  if (!cFci) throw new Error("Colonne FCI introuvable dans le CSV");

  const groupes = {};
  const ordreGroupes = [];

  rows.forEach((r, i) => {
    const opVal = s(r[cOp]);
    if (opVal && opVal.toUpperCase() !== "FT") return;

    const fciVal = s(r[cFci]);
    if (!fciVal) return;

    const obj = {
      FCI: fciVal,
      CODE_CH1: s(r[cCh1]),
      CODE_CH2: s(r[cCh2]),
      REF_CHAMBR: s(r[cRef]),
      BPE: s(r[cBpe]),
      MASQUE: s(r[cMasH]),
      DIAMETRE: s(r[cDiaH]),
      "MASQUE B": s(r[cMasB]),
      "DIAMETRE B": s(r[cDiaB]),
      ETIQUETTE_2: s(r[cEtiq2]),
      ORDRE: s(r[cOrdre]),
      COMMUNE: s(r[cCommune]),
      _idx: i,
    };

    if (!groupes[fciVal]) {
      groupes[fciVal] = [];
      ordreGroupes.push(fciVal);
    }
    groupes[fciVal].push(obj);
  });

  const sorter = (a, b) => {
    const na = safeInt(a.ORDRE);
    const nb = safeInt(b.ORDRE);
    const va = na === null ? 999999 : na;
    const vb = nb === null ? 999999 : nb;
    return reverse ? vb - va : va - vb;
  };
  for (const fci of ordreGroupes) groupes[fci].sort(sorter);

  return { groupes, ordreGroupes };
}

/* ============================================================
   GÉNÉRATION EXCEL SNAKE
   ============================================================ */
const DIAM_OFFSET = { 32: 9, 45: 10, 60: 11, 80: 12, 100: 13, 150: 16 };

async function generateExcelSnake(modeleBuffer, lignes, fciVal, codeCh2, numProjet) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(modeleBuffer);

  const ws = wb.getWorksheet("SNK_1");
  if (!ws) throw new Error("Feuille SNK_1 introuvable dans le modèle");

  // Cette fonction écrase la cellule de force, même si elle est fusionnée
  function forceWrite(addr, value) {
    if (value === null || value === undefined || value === "") return;
    
    // On tente de supprimer la fusion si elle existe pour cette cellule
    const mergeKey = Object.keys(ws._merges || {}).find(key => {
      const m = ws._merges[key];
      const model = m.model || m;
      // On vérifie si la cellule cible fait partie d'une fusion
      const cell = ws.getCell(addr);
      return cell.row >= model.top && cell.row <= model.bottom &&
             cell.col >= model.left && cell.col <= model.right;
    });

    if (mergeKey) {
      // Si on a trouvé une fusion, on écrit dans la cellule en haut à gauche
      const model = ws._merges[mergeKey].model || ws._merges[mergeKey];
      ws.getCell(model.top, model.left).value = value;
    } else {
      // Sinon on écrit normalement
      ws.getCell(addr).value = value;
    }
  }

  lignes.forEach((r, idx) => {
    const base = 4 + (idx * 19);

    const ref    = s(r.REF_CHAMBR);     // D4
    const code1  = s(r.CODE_CH1);       // 586
    const bpe    = s(r.BPE);
    const mas    = s(r.MASQUE);
    const diam   = safeInt(r.DIAMETRE);
    const masB   = s(r["MASQUE B"]);
    const diamB  = safeInt(r["DIAMETRE B"]);
    const etiq   = s(r.ETIQUETTE_2);    // 106
    const commune = s(r.COMMUNE);

    if (ref)   forceWrite(`M${base}`, ref);
    if (code1) forceWrite(`R${base}`, code1);
    
    // C'est ICI qu'on écrase le D4 par la vraie longueur
    if (etiq)  forceWrite(`T${base + 3}`, etiq);

    if (bpe) {
      forceWrite(`L${base}`, "M");
      const parts = [bpe, code1, commune, fciVal].filter(x => x);
      forceWrite(`Z${base + 2}`, parts.join(" - "));
    }

    forceWrite(`D${base + 6}`, "X");
    forceWrite(`X${base + 12}`, "X");

    if (mas) {
      const [l, n] = splitMasque(mas);
      forceWrite(`B${base + 3}`, l);
      if (n !== null) forceWrite(`D${base + 3}`, n);
    }

    if (masB) {
      const [l, n] = splitMasque(masB);
      forceWrite(`B${base + 16}`, l);
      if (n !== null) forceWrite(`D${base + 16}`, n);
    }

    if (diam !== null && DIAM_OFFSET[diam] !== undefined) {
      forceWrite(`H${base + DIAM_OFFSET[diam]}`, "X");
    }
    if (diamB !== null && DIAM_OFFSET[diamB] !== undefined) {
      forceWrite(`H${base + DIAM_OFFSET[diamB]}`, "X");
    }
  });

  const fciPropre  = sanitizeFilename(fciVal).replace(/[^a-zA-Z0-9_\-]/g, "");
  const codePropre = sanitizeFilename(codeCh2).replace(/[^a-zA-Z0-9_\-]/g, "");
  const filename = `Snake_${fciPropre}_${codePropre}.xlsx`;

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  return { blob, filename };
}

/* ============================================================
   ORCHESTRATION
   ============================================================ */
$("btn-generer").addEventListener("click", async () => {
  $("step-1").classList.add("hidden");
  $("step-2").classList.remove("hidden");
  $("btn-generer").disabled = true;
  $("log").innerHTML = "";

  try {
    log("📖 Lecture du CSV...");
    const rows = await parseCSVFile(state.csvFile);
    log(`   → ${rows.length} lignes lues`, "ok");

    const reverse = $("opt-sens").checked;
    const nomProjet = s($("inp-projet").value);

    log(`📊 Regroupement par FCI${reverse ? " (sens inversé)" : ""}...`);
    const { groupes, ordreGroupes } = preparerParFCI(rows, reverse);

    if (!ordreGroupes.length) {
      throw new Error("Aucun FCI trouvé dans le CSV");
    }
    log(`   → ${ordreGroupes.length} FCI détecté(s)`, "ok");
    ordreGroupes.forEach(fci => {
      log(`      • ${fci} : ${groupes[fci].length} chambre(s)`);
    });

    log("\n📥 Utilisation du modèle intégré...");
    const modeleBuffer = state.modeleBuffer;

    state.zip = new JSZip();
    let derniereChambrePrecedente = null;

    for (let i = 0; i < ordreGroupes.length; i++) {
      const fci = ordreGroupes[i];
      const lignes = groupes[fci];
      const codeCh2 = lignes[0]?.CODE_CH2 || "";

      log(`\n🔨 Traitement FCI ${fci} (commune ${codeCh2})...`);

      const lignesAvecRappel = [];
      if (derniereChambrePrecedente) {
        log(`   ↪ Rappel de CH${derniereChambrePrecedente.CODE_CH1}`);
        lignesAvecRappel.push({ ...derniereChambrePrecedente });
      }
      lignesAvecRappel.push(...lignes);

      const { blob, filename } = await generateExcelSnake(
        modeleBuffer, lignesAvecRappel, fci, codeCh2, nomProjet
      );
      state.zip.file(filename, blob);
      log(`   ✔ ${filename}`, "ok");

      if (lignes.length) {
        derniereChambrePrecedente = lignes[lignes.length - 1];
      }
      setProgress(((i + 1) / ordreGroupes.length) * 100);
    }

    log("\n📦 Préparation du ZIP...");
    state.zipBlob = await state.zip.generateAsync({ type: "blob" });

    $("step-2").classList.add("hidden");
    $("step-3").classList.remove("hidden");
    $("result-msg").textContent = `${ordreGroupes.length} fichier(s) Excel généré(s). Cliquez pour télécharger.`;

  } catch (err) {
    console.error(err);
    log("\n❌ ERREUR : " + err.message, "err");
    alert("Erreur : " + err.message);
    $("btn-generer").disabled = false;
    $("step-2").classList.add("hidden");
    $("step-1").classList.remove("hidden");
  }
});

$("btn-download").addEventListener("click", () => {
  if (!state.zipBlob) return;
  const url = URL.createObjectURL(state.zipBlob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Snakes_${new Date().toISOString().slice(0, 10)}.zip`;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 200);
});