/* ============================================================
   Générateur Snake Excel par FCI
   Version : tri chronologique strict + photos + caméra
   + regroupement FCI sur 12 caractères
   + exclusion des POP (NRO) et de certaines chambres
   ============================================================ */

const state = {
  csvFile: null,
  zip: null,
  zipBlob: null,
  modeleBuffer: null,
  photoFiles: [],
  photoIndex: [],
  chambresDisponibles: [],
};

const $ = id => document.getElementById(id);

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
  } catch (err) { console.error(err); }
  checkReady();
});

/* ---------- Gestion des fichiers ---------- */
const dropZone = $("drop-zone");
const fileInput = $("file-csv");
const csvInfo = $("csv-info");

function setCsv(file) {
  state.csvFile = file;
  state.chambresDisponibles = []; // force le rechargement pour la caméra
  csvInfo.textContent = `✓ ${file.name}`;
  csvInfo.style.color = "#2E7D32";
  checkReady();
}

dropZone.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", e => {
  if (e.target.files[0]) setCsv(e.target.files[0]);
});

dropZone.addEventListener("dragover", e => { e.preventDefault(); dropZone.style.background = "#E8F5E9"; });
dropZone.addEventListener("dragleave", () => { dropZone.style.background = "#F1F8E9"; });
dropZone.addEventListener("drop", e => {
  e.preventDefault();
  dropZone.style.background = "#F1F8E9";
  const file = e.dataTransfer.files[0];
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".csv")) {
    alert("Veuillez déposer un fichier CSV");
    return;
  }
  setCsv(file);
});

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
const safeInt = v => { const n = parseInt(String(v).trim(), 10); return isNaN(n) ? null : n; };
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
  return [m[1], m[2] ? parseInt(m[2], 10) : null];
}

// Nettoie le FCI : garde les 12 premiers caractères (F compris) si l'option est cochée
function nettoyerFCI(fciBrut, doitRegrouper) {
  if (!doitRegrouper) return String(fciBrut).trim();
  const str = String(fciBrut).trim();
  if (str.length > 12) {
    return str.substring(0, 12).replace(/[_-]+$/, "");
  }
  return str;
}

/* ============================================================
   FILTRES : POP et chambres exclues
   ============================================================ */

// Les POP / NRO (CODE_CH1 du type POP_xxx_0001) ne sont pas des chambres Orange
function estPop(codeCh1, refChambr) {
  const code = s(codeCh1).toUpperCase();
  const ref = s(refChambr).toUpperCase();
  return code.startsWith("POP_") || ref === "NRO";
}

// Clé de comparaison tolérante : sans accents, tirets, espaces ni majuscules
// ("VERRIERES-LE-BUISSON", "Verrières le Buisson" → "verriereslebuisson")
const cle = txt => normaliser(txt).replace(/[^a-z0-9]/g, "");

// Chambres intermédiaires : elles n'apparaissent PAS dans le fichier principal (Massy),
// seulement dans le fichier de leur propre FCI (Verrières). [COMMUNE, CODE_CH1]
const CHAMBRES_HORS_PRINCIPAL = [
  ["verrieres-le-buisson", "825"],
  ["verrieres-le-buisson", "832"],
].map(([com, ch]) => [cle(com), ch]);

// Communes "enclavées" : leur entrée et leur sortie sont insérées dans le fichier de la
// commune précédente (par ORDRE), et elles ont en plus leur propre fichier (FCI propre)
// qui commence par un rappel de la dernière chambre du fichier principal.
const COMMUNES_RATTACHEES = ["verrieres-le-buisson"].map(cle);

function estRattachee(commune) {
  return COMMUNES_RATTACHEES.includes(cle(commune));
}

function estHorsPrincipal(commune, codeCh1) {
  const c = cle(commune);
  const code = s(codeCh1);
  return CHAMBRES_HORS_PRINCIPAL.some(([com, ch]) => com === c && ch === code);
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
   CHARGEMENT POUR LA CAMÉRA
   ============================================================ */
async function chargerChambresDansState() {
  if (!state.csvFile) return;
  if (state.chambresDisponibles.length > 0) return;
  try {
    const rows = await parseCSVFile(state.csvFile);
    const cols = Object.keys(rows[0]);
    const cFci = trouverColonne(cols, "FCI", "N°_FCI") || trouverColonneFuzzy(cols, "fci");
    const cCh1 = trouverColonne(cols, "CODE_CH1", "CODE CH1");
    const cCh2 = trouverColonne(cols, "CODE_CH2", "INSEE");
    const cOrdre = trouverColonne(cols, "ORDRE");
    const cCommune = trouverColonne(cols, "COMMUNE");
    const cRef = trouverColonne(cols, "REF_CHAMBR");

    state.chambresDisponibles = [];
    rows.forEach(r => {
      const ch1 = s(r[cCh1]);
      if (!ch1) return;
      if (estPop(r[cCh1], r[cRef])) return;          // pas de POP
      state.chambresDisponibles.push({
        label: `CH ${ch1} (${s(r[cCommune])}) - Ordre ${s(r[cOrdre])}`,
        CODE_CH1: ch1,
        CODE_CH2: s(r[cCh2]),
        ORDRE: s(r[cOrdre]),
        FCI: s(r[cFci]),
        COMMUNE: s(r[cCommune])
      });
    });
  } catch (e) { console.error(e); }
}

/* ============================================================
   GROUPEMENT PAR FCI
   ============================================================ */
function preparerParFCI(rows, reverse = false, doitRegrouper = true) {
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

  // On parcourt les lignes dans l'ordre du réseau pour pouvoir rattacher
  // certaines communes au FCI de la chambre précédente
  const rowsTriees = [...rows].sort((a, b) =>
    (safeInt(a[cOrdre]) ?? 999999) - (safeInt(b[cOrdre]) ?? 999999)
  );
  const ajouter = (obj, cleGroupe) => {
    if (!groupes[cleGroupe]) { groupes[cleGroupe] = []; ordreGroupes.push(cleGroupe); }
    groupes[cleGroupe].push(obj);
  };

  const fabriquer = (r, fciVal, fciBrut) => ({
    FCI: fciVal,           // FCI affiché pour cette chambre
    _fciOrigine: fciBrut,
    CODE_CH1: s(r[cCh1]), CODE_CH2: s(r[cCh2]),
    REF_CHAMBR: s(r[cRef]), BPE: s(r[cBpe]), MASQUE: s(r[cMasH]),
    DIAMETRE: s(r[cDiaH]), "MASQUE B": s(r[cMasB]), "DIAMETRE B": s(r[cDiaB]),
    ETIQUETTE_2: s(r[cEtiq2]), ORDRE: s(r[cOrdre]), COMMUNE: s(r[cCommune]),
  });

  let dernierePrincipale = null;     // dernière chambre du flux principal (ex : 1159 à Massy)
  const rappelFait = new Set();      // groupes enclavés qui ont déjà reçu leur rappel

  rowsTriees.forEach((r) => {
    if (s(r[cOp]) && s(r[cOp]).toUpperCase() !== "FT") return;

    // On ne garde que les chambres Orange : pas de POP / NRO
    if (estPop(r[cCh1], r[cRef])) return;

    const fciBrut = s(r[cFci]);
    if (!fciBrut) return;

    // FCI propre de la chambre (regroupé sur 12 caractères si l'option est cochée)
    const fciVal = nettoyerFCI(fciBrut, doitRegrouper);

    // --- Commune enclavée (ex : Verrières au milieu du fichier de Massy) ---
    if (doitRegrouper && estRattachee(r[cCommune]) && dernierePrincipale) {
      // 1. Rappel de la dernière chambre du fichier principal, en tête du fichier enclavé
      //    (elle garde sa propre commune et son propre FCI, pas ceux de l'enclave)
      if (!rappelFait.has(fciVal)) {
        rappelFait.add(fciVal);
        ajouter({ ...dernierePrincipale, _rappel: true }, fciVal);
      }

      // 2. Fichier principal : uniquement l'entrée et la sortie (pas les intermédiaires)
      if (!estHorsPrincipal(r[cCommune], r[cCh1])) {
        const objP = fabriquer(r, dernierePrincipale.FCI, fciBrut);
        ajouter(objP, dernierePrincipale.FCI);
        dernierePrincipale = objP;
      }

      // 3. Fichier de la commune enclavée : toutes ses chambres
      ajouter(fabriquer(r, fciVal, fciBrut), fciVal);
      return;
    }

    // --- Cas normal ---
    const obj = fabriquer(r, fciVal, fciBrut);

    // Nouveau fichier (changement de FCI / commune) : on commence par un rappel de la
    // dernière chambre du fichier précédent (ex : dernière de Chilly 91161 en tête de
    // Longjumeau 91345). Elle garde sa propre commune et son propre FCI.
    if (!groupes[fciVal] && dernierePrincipale && dernierePrincipale.FCI !== fciVal) {
      ajouter({ ...dernierePrincipale, _rappel: true }, fciVal);
    }

    ajouter(obj, fciVal);
    dernierePrincipale = obj;
  });

  // --- TRI STRICT PAR ORDRE ---
  // On ne trie plus par FCI d'origine, on respecte l'ordre numérique absolu
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
const TYPES = ["MA", "MB", "MC", "MD", "MIRE_LG", "MIRE_L", "MIRE_P", "BOITE", "CASSETTE", "VUE_ENSEMBLE", "ETIQ_CABLE", "ETIQ_BOITE"];
const ALIAS = { MIRE: "MIRE_LG" };
const photoCache = new Map();

function indexPhotos(files) {
  const idx = [];
  for (const f of files) {
    if (!/\.(jpe?g|png|webp|bmp)$/i.test(f.name)) continue;
    const segs = f.webkitRelativePath
      .replace(/\\/g, "/")
      .split("/")
      .map(x => x.trim().toUpperCase())
      .filter(Boolean);

    if (segs.length > 2 && segs[0].startsWith("PHOTO")) segs.shift();

    segs.pop();
    const raw = segs[segs.length - 1];
    const type = ALIAS[raw] || raw;
    if (!TYPES.includes(type)) continue;
    idx.push({ file: f, type, parents: segs.slice(0, -1) });
  }
  return idx;
}

function findPhotosForChambre(ordre, ch1, index) {
  const photos = Object.fromEntries(TYPES.map(t => [t, []]));
  const chRaw = s(ch1).toUpperCase();
  if (!chRaw) return photos;
  const ch = chRaw.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const reCh = new RegExp(`(^|[^A-Z0-9])CH[ _-]*0*${ch}(?![0-9A-Z])`);
  const num = safeInt(ordre);

  for (const it of index) {
    if (!it.parents.some(p => reCh.test(p))) continue;
    const ordres = it.parents.filter(p => /^0*\d+$/.test(p)).map(p => parseInt(p, 10));
    if (ordres.length && !ordres.includes(num)) continue;
    photos[it.type].push(it.file);
  }
  for (const t of TYPES) {
    photos[t].sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
  }
  return photos;
}

async function compressedBase64(file, maxDim = 400, quality = 0.7) {
  if (photoCache.has(file)) return photoCache.get(file);
  let bmp;
  try {
    bmp = await createImageBitmap(file, { imageOrientation: "from-image" });
  } catch (e) {
    console.warn("Photo illisible :", file.webkitRelativePath || file.name);
    return null;
  }
  const r = Math.min(1, maxDim / Math.max(bmp.width, bmp.height));
  const c = document.createElement("canvas");
  c.width = Math.round(bmp.width * r);
  c.height = Math.round(bmp.height * r);
  c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
  bmp.close();
  const b64 = c.toDataURL("image/jpeg", quality);
  photoCache.set(file, b64);
  return b64;
}

/* ---------- Dossier de photos (PC) ---------- */
const photoInput = $("file-photos");
if (photoInput) {
  photoInput.addEventListener("change", e => {
    state.photoFiles = Array.from(e.target.files);
    state.photoIndex = indexPhotos(state.photoFiles);
    log(`📷 ${state.photoIndex.length} photos indexées (dossier)`);
  });
}

/* ---------- ZIP de photos (mobile) ---------- */
const zipInput = $("file-zip");
if (zipInput) {
  zipInput.addEventListener("change", async e => {
    const zipFile = e.target.files[0];
    if (!zipFile) return;
    try {
      log("📦 Lecture du ZIP photos...");
      const zip = await JSZip.loadAsync(zipFile);
      const entries = Object.values(zip.files).filter(z =>
        !z.dir &&
        /\.(jpe?g|png|webp)$/i.test(z.name) &&
        !z.name.includes("__MACOSX") &&
        !z.name.split("/").pop().startsWith("._")
      );

      const stubs = [];
      let n = 0;
      for (const entry of entries) {
        const blob = await entry.async("blob");
        const realFile = new File([blob], entry.name.split("/").pop(), { type: blob.type || "image/jpeg" });
        const b64 = await compressedBase64(realFile);
        if (!b64) continue;
        const stub = { name: realFile.name, webkitRelativePath: entry.name };
        photoCache.set(stub, b64);
        stubs.push(stub);
        if (++n % 20 === 0) log(`   ${n}/${entries.length} photos traitées...`);
      }
      state.photoFiles = state.photoFiles.concat(stubs);
      state.photoIndex = indexPhotos(state.photoFiles);
      log(`📷 ${state.photoIndex.length} photos indexées au total (ZIP ajouté)`, "ok");
    } catch (err) {
      console.error(err);
      log("❌ ZIP illisible : " + err.message, "err");
      alert("Impossible de lire le ZIP : " + err.message);
    }
  });
}

/* ============================================================
   CAMÉRA
   ============================================================ */
const cameraDialog = $("camera-dialog");
const btnCamera = $("btn-camera");
const btnCameraGo = $("btn-camera-go");
const btnCameraClose = $("btn-camera-close");
const cameraInput = $("camera-input");
const selChambre = $("sel-chambre");
const selType = $("sel-type");

if (btnCamera) {
  btnCamera.addEventListener("click", async () => {
    await chargerChambresDansState();
    if (state.chambresDisponibles.length === 0) {
      alert("Veuillez d'abord sélectionner un fichier CSV valide.");
      return;
    }
    selChambre.innerHTML = state.chambresDisponibles
      .map((c, i) => `<option value="${i}">${c.label}</option>`)
      .join("");
    cameraDialog.showModal();
  });
}

if (btnCameraClose) {
  btnCameraClose.addEventListener("click", () => cameraDialog.close());
}

if (btnCameraGo) {
  btnCameraGo.addEventListener("click", () => cameraInput.click());
}

if (cameraInput) {
  cameraInput.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const chambre = state.chambresDisponibles[selChambre.value];
    const type = selType.value;

    const dossierChambre = `CH ${chambre.CODE_CH1}`;
    const dossierOrdre = String(safeInt(chambre.ORDRE) ?? "ZZZ").padStart(3, "0");
    const webkitRelativePath = `${dossierOrdre}/${dossierChambre}/${type}/${file.name}`;

    const stub = { name: file.name, webkitRelativePath: webkitRelativePath };

    const b64 = await compressedBase64(file);
    if (b64) {
      photoCache.set(stub, b64);
      state.photoFiles.push(stub);
      state.photoIndex = indexPhotos(state.photoFiles);
      log(`📸 Photo ajoutée : CH ${chambre.CODE_CH1} → ${type}`, "ok");
    } else {
      alert("Format d'image non supporté.");
    }

    cameraDialog.close();
    cameraInput.value = "";
  });
}

/* ============================================================
   GÉNÉRATION EXCEL SNAKE
   ============================================================ */
const DIAM_OFFSET = { 32: 9, 45: 10, 60: 11, 80: 12, 100: 13, 150: 16 };

async function generateExcelSnake(modeleBuffer, lignes, fciVal, codeCh2, numProjet, photoIndex) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(modeleBuffer);
  const ws = wb.getWorksheet("SNK_1");

  const colsBase = { MA: 34, MB: 35, MC: 36, MD: 37 };
  const colsBoite = { MIRE_LG: 38, MIRE_L: 39, MIRE_P: 40, BOITE: 41, CASSETTE: 42, VUE_ENSEMBLE: 43, ETIQ_CABLE: 44, ETIQ_BOITE: 45 };
  const allCols = { ...colsBase, ...colsBoite };

  for (const [key, colIdx] of Object.entries(allCols)) {
    ws.getColumn(colIdx).width = 20;
    const cell = ws.getCell(2, colIdx);
    cell.value = key.replace(/_/g, " ");
    cell.font = { bold: true, color: { argb: "FF2E7D32" } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  }

  let totalLongueur = 0;
  const formuleM2 = [];

  function writeSafe(addr, value) {
    if (value === null || value === undefined || value === "") return;
    const cell = ws.getCell(addr);
    if (cell.isMerged) ws.getCell(cell.master.address).value = value;
    else cell.value = value;
  }

  function writeMasterStyled(addr, value, font, alignment) {
    const cell = ws.getCell(addr);
    const target = cell.isMerged ? ws.getCell(cell.master.address) : cell;
    target.value = value;
    if (font) target.font = font;
    if (alignment) target.alignment = alignment;
  }

  const green = { style: "medium", color: { argb: "FF4CAF50" } };

  for (let idx = 0; idx < lignes.length; idx++) {
    const r = lignes[idx];
    const base = 4 + (idx * 19);
    const hasBpe = !!s(r.BPE);

    if (s(r.REF_CHAMBR)) writeSafe(`M${base}`, s(r.REF_CHAMBR));
    if (s(r.CODE_CH1)) writeSafe(`R${base}`, s(r.CODE_CH1));

    const etiq = s(r.ETIQUETTE_2);
    if (etiq) {
      const tCell = ws.getCell(`T${base + 3}`);
      const etiqNum = parseFloat(etiq.replace(",", "."));
      const valToWrite = isNaN(etiqNum) ? etiq : etiqNum;
      const target = tCell.isMerged ? ws.getCell(tCell.master.address) : tCell;
      target.value = valToWrite;
      target.numFmt = "0";
      // Le rappel n'est pas compté dans le total (déjà compté dans son propre fichier)
      if (!isNaN(etiqNum) && !r._rappel) {
        formuleM2.push(`T${base + 3}`);
        totalLongueur += etiqNum;
      }
    }

    if (hasBpe) writeSafe(`L${base}`, "M");

    // FCI propre à la chambre (le rappel garde le FCI et la commune de son fichier d'origine)
    const fciLigne = s(r.FCI) || fciVal;
    const commentaire = hasBpe
      ? `BOITE : ${s(r.BPE)}\n${s(r.COMMUNE)}\n${fciLigne}`
      : `${s(r.COMMUNE)}\n${fciLigne}`;
    writeMasterStyled(
      `Z${base + 2}`, commentaire,
      { bold: true, size: 10 },
      { vertical: "middle", horizontal: "center", wrapText: true, textRotation: 90 }
    );

    writeSafe(`D${base + 6}`, "X");
    writeSafe(`X${base + 12}`, "X");

    if (s(r.MASQUE)) {
      const masqueTxt = s(r.MASQUE);
      const masqueNorm = normaliser(masqueTxt);
      const obsTxt = masqueTxt.toUpperCase().replace(/\s+/g, "\n");

      if (masqueNorm.includes("non emprunte")) {
        writeSafe(`H${base}`, "X");
        writeMasterStyled(`Z${base}`, obsTxt);
      } else {
        const [l, n] = splitMasque(masqueTxt);
        if (n !== null) {
          writeSafe(`B${base + 3}`, l);
          writeSafe(`D${base + 3}`, n);
        } else {
          writeSafe(`O${base + 14}`, "X");
          writeMasterStyled(`Z${base}`, obsTxt);
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

    // --- 3. Photos ---
    if (photoIndex && photoIndex.length > 0) {
      const photos = findPhotosForChambre(r.ORDRE, r.CODE_CH1, photoIndex);
      const activeCols = hasBpe
        ? allCols
        : {
            ...colsBase,
            VUE_ENSEMBLE: colsBoite.VUE_ENSEMBLE
          };
      const startRow = base + 2;
      const endRow = startRow + 6;

      for (const [key, colIdx] of Object.entries(activeCols)) {
        const list = photos[key];

        if (!list.length) {
          const c = ws.getCell(startRow, colIdx);
          c.value = "Pas de photo";
          c.alignment = { horizontal: "center", vertical: "middle" };
          c.font = { italic: true, color: { argb: "FF808080" } };
          continue;
        }

        const b64 = await compressedBase64(list[0]);
        if (!b64) continue;

        const imgId = wb.addImage({ base64: b64, extension: "jpeg" });
        ws.addImage(imgId, {
          tl: { col: colIdx - 1, row: startRow - 1 },
          ext: { width: 130, height: 100 },
        });

        for (let rl = startRow; rl <= endRow; rl++) {
          const cell = ws.getCell(rl, colIdx);
          cell.border = {
            left: green, right: green,
            ...(rl === startRow && { top: green }),
            ...(rl === endRow && { bottom: green }),
          };
        }
      }
    }
  }

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
  return {
    blob: new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    filename,
  };
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
    if (!rows.length) throw new Error("CSV vide");
    const reverse = $("opt-sens").checked;
    const nomProjet = s($("inp-projet").value);

    const doitRegrouper = $("opt-regrouper") ? $("opt-regrouper").checked : true;
    if (doitRegrouper) log("🔗 Regroupement des FCI identiques (12 premiers caractères) activé.");

    log("📊 Regroupement par FCI...");
    const { groupes, ordreGroupes } = preparerParFCI(rows, reverse, doitRegrouper);
    if (!ordreGroupes.length) throw new Error("Aucune ligne FT avec un numéro FCI");

    state.zip = new JSZip();

    for (let i = 0; i < ordreGroupes.length; i++) {
      const fci = ordreGroupes[i];
      const lignes = groupes[fci];
      // Le rappel vient d'une autre commune : on prend le code INSEE de la 1re vraie chambre
      const codeCh2 = (lignes.find(l => !l._rappel) || lignes[0])?.CODE_CH2 || "";

      log(`\n🔨 Traitement FCI ${fci} (${lignes.length} chambres)...`);

      const { blob, filename } = await generateExcelSnake(
        state.modeleBuffer, [...lignes], fci, codeCh2, nomProjet, state.photoIndex
      );

      state.zip.file(filename, blob);
      log(`   ✔ ${filename}`, "ok");

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
  a.href = url;
  a.download = `Snakes_${new Date().toISOString().slice(0, 10)}.zip`;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 200);
});
