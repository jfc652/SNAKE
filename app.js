/* ============================================================
   Générateur Snake Excel par FCI - Version Finale (Mobile)
   ============================================================ */

const state = {
  csvFile: null,
  zip: null,
  zipBlob: null,
  modeleBuffer: null,
  photoFiles: [],
};

const $ = id => document.getElementById(id);

/* ---------- Détection Mobile ---------- */
const isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);

/* ---------- Pré-chargement du modèle ---------- */
window.addEventListener("DOMContentLoaded", () => {
  try {
    if (typeof getModeleArrayBuffer !== "function") throw new Error("modele-b64.js manquant");
    state.modeleBuffer = getModeleArrayBuffer();
    const info = $("modele-info");
    if (info) {
      info.innerHTML = `✓ Modèle intégré chargé (${getModeleSizeKB()} Ko)`;
      info.style.color = "#2E7D32";
    }
    
    if (isMobile) {
      const hint = $("photo-hint");
      if (hint) hint.innerHTML = "⚠️ Sur mobile, la sélection de dossier n'est pas supportée. Sélectionnez les photos manuellement ou utilisez un PC.";
    }
  } catch (err) { 
    console.error(err); 
  }
  checkReady();
});

/* ---------- Gestion des fichiers ---------- */
const dropZone = $("drop-zone");
const fileInput = $("file-csv");
const csvInfo = $("csv-info");

dropZone.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", e => {
  if (e.target.files[0]) {
    state.csvFile = e.target.files[0];
    csvInfo.textContent = `✓ ${state.csvFile.name}`;
    csvInfo.style.color = "#2E7D32";
    checkReady();
  }
});

dropZone.addEventListener("dragover", (e) => { 
  e.preventDefault(); 
  dropZone.style.background = "#E8F5E9"; 
  dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => { 
  dropZone.style.background = "#F1F8E9"; 
  dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (e) => {
  e.preventDefault(); 
  dropZone.style.background = "#F1F8E9"; 
  dropZone.classList.remove("dragover");
  const file = e.dataTransfer.files[0];
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".csv")) {
    alert("Veuillez déposer un fichier CSV");
    return;
  }
  state.csvFile = file;
  csvInfo.textContent = `✓ ${file.name}`;
  csvInfo.style.color = "#2E7D32";
  checkReady();
});

const photoInput = $("file-photos");
if (photoInput) {
  photoInput.addEventListener("change", e => {
    state.photoFiles = Array.from(e.target.files);
    console.log("Nombre de fichiers photos :", state.photoFiles.length);
    log(`📷 ${state.photoFiles.length} fichiers photos chargés`);
  });
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
const safeInt = v => { const n = parseInt(String(v).trim()); return isNaN(n) ? null : n; };
const sanitizeFilename = name => String(name).replace(/[<>:"/\\|?*]/g, "_");
const normaliser = txt => s(txt).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

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
      header: true, delimiter: ";", skipEmptyLines: true,
      transformHeader: h => h.trim().replace(/^\uFEFF/, ""),
      complete: r => resolve(r.data), error: reject,
    });
  });
}

/* ============================================================
   GROUPEMENT PAR FCI
   ============================================================ */
function preparerParFCI(rows, reverse = false) {
  const cols = Object.keys(rows[0]);
  const cOp = trouverColonne(cols, "OP");
  const cFci = trouverColonne(cols, "FCI", "N°_FCI") || trouverColonneFuzzy(cols, "fci");
  const cCh1 = trouverColonne(cols, "CODE_CH1", "CODE CH1");
  const cCh2 = trouverColonne(cols, "CODE_CH2", "INSEE");
  const cRef = trouverColonne(cols, "REF_CHAMBR");
  const cBpe = trouverColonne(cols, "BPE");
  const cMasH = trouverColonne(cols, "MASQUE", "Masque H");
  const cMasB = trouverColonne(cols, "MASQUE B");
  const cDiaH = trouverColonne(cols, "DIAMETRE", "Diamètre H");
  const cDiaB = trouverColonne(cols, "DIAMETRE B");
  const cEtiq2 = trouverColonne(cols, "ETIQUETTE_2", "étiquette_2") || trouverColonneFuzzy(cols, "etiquette");
  const cOrdre = trouverColonne(cols, "ORDRE");
  const cCommune = trouverColonne(cols, "COMMUNE");

  const groupes = {};
  const ordreGroupes = [];

  rows.forEach((r, i) => {
    if (s(r[cOp]) && s(r[cOp]).toUpperCase() !== "FT") return;
    const fciVal = s(r[cFci]);
    if (!fciVal) return;

    const obj = {
      FCI: fciVal, CODE_CH1: s(r[cCh1]), CODE_CH2: s(r[cCh2]),
      REF_CHAMBR: s(r[cRef]), BPE: s(r[cBpe]), MASQUE: s(r[cMasH]),
      DIAMETRE: s(r[cDiaH]), "MASQUE B": s(r[cMasB]), "DIAMETRE B": s(r[cDiaB]),
      ETIQUETTE_2: s(r[cEtiq2]), ORDRE: s(r[cOrdre]), COMMUNE: s(r[cCommune]),
    };

    if (!groupes[fciVal]) { groupes[fciVal] = []; ordreGroupes.push(fciVal); }
    groupes[fciVal].push(obj);
  });

  for (const fci of ordreGroupes) {
    groupes[fci].sort((a, b) => {
      const va = safeInt(a.ORDRE) ?? 999999;
      const vb = safeInt(b.ORDRE) ?? 999999;
      return reverse ? vb - va : va - vb;
    });
  }
  return { groupes, ordreGroupes };
}

/* ============================================================
   PHOTOS
   ============================================================ */
async function compressImage(file, maxDim = 150, quality = 0.3) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      let w = img.width, h = img.height;
      if (w > maxDim || h > maxDim) {
        if (w > h) { h = Math.round(h * maxDim / w); w = maxDim; }
        else { w = Math.round(w * maxDim / h); h = maxDim; }
      }
      canvas.width = w; canvas.height = h;
      canvas.getContext("2d").drawImage(img, 0, 0, w, h);
      resolve(canvas.toDataURL("image/jpeg", quality));
    };
    img.onerror = () => resolve(null);
    img.src = URL.createObjectURL(file);
  });
}

function findPhotosForChambre(ordre, ch1, codeInsee, photoFiles) {
  const photos = { MA: [], MB: [], MC: [], MD: [], MIRE_LG: [], MIRE_L: [], MIRE_P: [], BOITE: [], CASSETTE: [], VUE_ENSEMBLE: [], ETIQ_CABLE: [], ETIQ_BOITE: [] };
  
  const chStr = String(ch1).trim(); 
  const inseeStr = String(codeInsee).trim();

  const filesForCh = photoFiles.filter(f => {
    let path = f.webkitRelativePath.replace(/\\/g, "/").toUpperCase();
    const hasInsee = path.includes(inseeStr);
    if (!hasInsee) return false;

    const pathAvecCH = path.split("/").map(part => " CH " + part).join("/");
    const regexChambre = new RegExp(`CH[\\s_-]*${chStr}(?![0-9A-Za-z])`, "i");

    return regexChambre.test(pathAvecCH);
  });

  if (filesForCh.length > 0) {
    console.log(`[TROUVÉ] Dossier pour CH ${chStr} (INSEE ${inseeStr}) :`, filesForCh[0].webkitRelativePath);
  }

  for (const file of filesForCh) {
    const path = file.webkitRelativePath.toUpperCase().replace(/\\/g, "/");
    const inFolder = (folderName) => path.includes("/" + folderName + "/");
    
    if (inFolder("MA")) photos.MA.push(file);
    if (inFolder("MB")) photos.MB.push(file);
    if (inFolder("MC")) photos.MC.push(file);
    if (inFolder("MD")) photos.MD.push(file);
    if (inFolder("MIRE") || inFolder("MIRE_LG")) photos.MIRE_LG.push(file);
    if (inFolder("MIRE_L")) photos.MIRE_L.push(file);
    if (inFolder("MIRE_P")) photos.MIRE_P.push(file);
    if (inFolder("BOITE")) photos.BOITE.push(file);
    if (inFolder("CASSETTE")) photos.CASSETTE.push(file);
    if (inFolder("VUE_ENSEMBLE")) photos.VUE_ENSEMBLE.push(file);
    if (inFolder("ETIQ_CABLE")) photos.ETIQ_CABLE.push(file);
    if (inFolder("ETIQ_BOITE")) photos.ETIQ_BOITE.push(file);
  }
  
  return photos;
}

/* ============================================================
   GÉNÉRATION EXCEL SNAKE
   ============================================================ */
const DIAM_OFFSET = { 32: 9, 45: 10, 60: 11, 80: 12, 100: 13, 150: 16 };

async function generateExcelSnake(modeleBuffer, lignes, fciVal, codeCh2, numProjet, photoFiles) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(modeleBuffer);
  const ws = wb.getWorksheet("SNK_1");

  // --- 1. Colonnes et Titres ---
  const colsBase = { MA: 34, MB: 35, MC: 36, MD: 37 }; 
  const colsBoite = { MIRE_LG: 38, MIRE_L: 39, MIRE_P: 40, BOITE: 41, CASSETTE: 42, VUE_ENSEMBLE: 43, ETIQ_CABLE: 44, ETIQ_BOITE: 45};
  const allCols = { ...colsBase, ...colsBoite };
  
  for (const [key, colIdx] of Object.entries(allCols)) {
    ws.getColumn(colIdx).width = 20;
    const cell = ws.getCell(2, colIdx);
    cell.value = key.replace("_", " "); 
    cell.font = { bold: true, color: { argb: "FF2E7D32" } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  }

  let totalLongueur = 0;
  let formuleM2 = [];

  // --- 2. Boucle Chambres ---
  for (let idx = 0; idx < lignes.length; idx++) {
    log(`📍 Chambre ${lignes[idx].CODE_CH1} en cours...`);
    const r = lignes[idx];
    const base = 4 + (idx * 19);
    const hasBpe = !!s(r.BPE);

    function writeSafe(addr, value) {
      if (value === null || value === undefined || value === "") return;
      const cell = ws.getCell(addr);
      if (cell.isMerged) ws.getCell(cell.master.address).value = value;
      else cell.value = value;
    }

    // Textes
    if (s(r.REF_CHAMBR)) writeSafe(`M${base}`, s(r.REF_CHAMBR));
    if (s(r.CODE_CH1)) writeSafe(`R${base}`, s(r.CODE_CH1));
    
    // Longueur (T) et Formule M2
    const etiq = s(r.ETIQUETTE_2);
    if (etiq) {
      const tCell = ws.getCell(`T${base + 3}`);
      const etiqNum = parseFloat(etiq.replace(",", "."));
      const valToWrite = isNaN(etiqNum) ? etiq : etiqNum;
      
      if (tCell.isMerged) {
        ws.getCell(tCell.master.address).value = valToWrite;
        ws.getCell(tCell.master.address).numFmt = "0";
      } else {
        tCell.value = valToWrite;
        tCell.numFmt = "0";
      }
      
      if (!isNaN(etiqNum)) {
        formuleM2.push(`T${base + 3}`);
        totalLongueur += etiqNum;
      }
    }

    if (hasBpe) {
      writeSafe(`L${base}`, "M");
    }

    const cellZ = ws.getCell(`Z${base + 2}`);
    let commentaire = "";
    if (s(r.BPE)) {
      commentaire = `BOITE : ${s(r.BPE)}\n` + `${s(r.COMMUNE)}\n` + `${fciVal}`;
    } else {
      commentaire = `${s(r.COMMUNE)}\n` + `${fciVal}`;
    }
    
    if (cellZ.isMerged) {
      const master = ws.getCell(cellZ.master.address);
      master.value = commentaire;
      master.font = { bold: true, size: 10 };
      master.alignment = { vertical: "middle", horizontal: "center", wrapText: true, textRotation: 90 };
    } else {
      cellZ.value = commentaire;
      cellZ.font = { bold: true, size: 10 };
      cellZ.alignment = { vertical: "middle", horizontal: "center", wrapText: true, textRotation: 90 };
    }

    writeSafe(`D${base + 6}`, "X");
    writeSafe(`X${base + 12}`, "X");

    if (s(r.MASQUE)) {
      const masqueTxt = s(r.MASQUE);
      const masqueNorm = normaliser(masqueTxt);

      if (masqueNorm.includes("non empruntee") || masqueNorm.includes("non empruntée")) {
        writeSafe(`H${base + 14}`, "X");
        const cellObs = ws.getCell(`Z${base}`);
        if (cellObs.isMerged) ws.getCell(cellObs.master.address).value = masqueTxt.toUpperCase().replace(/\s+/g, "\n");
        else cellObs.value = masqueTxt.toUpperCase().replace(/\s+/g, "\n");
      } else {
        const [l, n] = splitMasque(masqueTxt);
        if (n !== null) {
          writeSafe(`B${base + 3}`, l);
          writeSafe(`D${base + 3}`, n);
        } else {
          writeSafe(`O${base + 14}`, "X");
          const cellObs = ws.getCell(`Z${base}`);
          if (cellObs.isMerged) ws.getCell(cellObs.master.address).value = masqueTxt.toUpperCase().replace(/\s+/g, "\n");
          else cellObs.value = masqueTxt.toUpperCase().replace(/\s+/g, "\n");
        }
      }
    }

    if (s(r["MASQUE B"])) {
      const [l, n] = splitMasque(r["MASQUE B"]);
      writeSafe(`B${base + 16}`, l);
      if (n !== null) writeSafe(`D${base + 16}`, n);
    }
    
    const diam = safeInt(r.DIAMETRE);
    if (diam && DIAM_OFFSET[diam]) writeSafe(`H${base + DIAM_OFFSET[diam]}`, "X");
    
    const diamB = safeInt(r["DIAMETRE B"]);
    if (diamB && DIAM_OFFSET[diamB]) writeSafe(`H${base + DIAM_OFFSET[diamB]}`, "X");

    // --- 3. Photos (avec la méthode "Bottom-Right") ---
    if (photoFiles && photoFiles.length > 0) {
      const photos = findPhotosForChambre(r.ORDRE, r.CODE_CH1, r.CODE_CH2, photoFiles);
      const activeCols = hasBpe ? { ...colsBase, ...colsBoite } : colsBase;

      for (const [key, colIdx] of Object.entries(activeCols)) {
        if (photos[key] && photos[key].length > 0) {
          try {
            const base64 = await compressImage(photos[key][0]);
            if (!base64) continue;

            const imgId = wb.addImage({ base64: base64, extension: 'jpeg' });
            const startRow = base + 2; 
            
            // L'ASTUCE FINALE : Bottom-Right (br). On dit à Excel 
            // "L'image commence ici, et se termine à cette cellule exacte".
            // C'est ainsi qu'Excel stocke lui-même ses images, donc c'est le mieux lu par le mobile.
            ws.addImage(imgId, {
              tl: { col: colIdx - 1, row: startRow - 1 },
              br: { col: colIdx, row: startRow + 6 },
              editAs: 'oneCell' 
            });

            // Liseré vert
            const endRow = startRow + 6; 
            for (let rLigne = startRow; rLigne <= endRow; rLigne++) {
              const cell = ws.getCell(rLigne, colIdx);
              if (!cell.border) cell.border = {};
              cell.border.left = { style: "medium", color: { argb: "FF4CAF50" } };
              cell.border.right = { style: "medium", color: { argb: "FF4CAF50" } };
              if (rLigne === startRow) cell.border.top = { style: "medium", color: { argb: "FF4CAF50" } };
              if (rLigne === endRow) cell.border.bottom = { style: "medium", color: { argb: "FF4CAF50" } };
            }

            // Pause anti-crash
            await new Promise(resolve => setTimeout(resolve, 10));

          } catch (err) {
            console.warn(`Erreur photo pour CH ${r.CODE_CH1} (${key})`, err);
          }
        }
      }
    }
  }

  // --- 4. Totaux M2 ---
  const m2Cell = ws.getCell("M2");
  const m2Master = m2Cell.isMerged ? ws.getCell(m2Cell.master.address) : m2Cell;
  
  if (formuleM2.length > 0) {
    m2Master.value = { formula: `SUM(${formuleM2.join(",")})` };
  } else {
    m2Master.value = totalLongueur;
  }
  m2Master.numFmt = "0";

  const fciPropre = sanitizeFilename(fciVal).replace(/[^a-zA-Z0-9_\-]/g, "");
  const codePropre = sanitizeFilename(codeCh2).replace(/[^a-zA-Z0-9_\-]/g, "");
  const filename = `Snake_${fciPropre}_${codePropre}.xlsx`;

  const buf = await wb.xlsx.writeBuffer();
  return { blob: new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filename };
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
    const reverse = $("opt-sens").checked;
    const nomProjet = s($("inp-projet").value);

    log(`📊 Regroupement par FCI...`);
    const { groupes, ordreGroupes } = preparerParFCI(rows, reverse);

    state.zip = new JSZip();
    let derniereChambrePrecedente = null;

    for (let i = 0; i < ordreGroupes.length; i++) {
      const fci = ordreGroupes[i];
      const lignes = groupes[fci];
      const codeCh2 = lignes[0]?.CODE_CH2 || "";

      log(`\n🔨 Traitement FCI ${fci}...`);

      const lignesAvecRappel = [];
      if (derniereChambrePrecedente) lignesAvecRappel.push({ ...derniereChambrePrecedente });
      lignesAvecRappel.push(...lignes);

      const { blob, filename } = await generateExcelSnake(
        state.modeleBuffer, lignesAvecRappel, fci, codeCh2, nomProjet, state.photoFiles
      );
      
      state.zip.file(filename, blob);
      log(`   ✔ ${filename}`, "ok");

      if (lignes.length) derniereChambrePrecedente = lignes[lignes.length - 1];
      setProgress(((i + 1) / ordreGroupes.length) * 100);
    }

    log("\n📦 Préparation du ZIP...");
    state.zipBlob = await state.zip.generateAsync({ type: "blob" });

    $("step-2").classList.add("hidden");
    $("step-3").classList.remove("hidden");
    $("result-msg").textContent = `${ordreGroupes.length} fichier(s) Excel généré(s).`;

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
  a.href = url; a.download = `Snakes_${new Date().toISOString().slice(0, 10)}.zip`;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 200);
});
