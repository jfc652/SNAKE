/* ============================================================
   Fusion de deux dossiers de photos (port JS de fusion_dossiers_chambres.py)
   Structure attendue dans chaque dossier : Dossier/Chambre/.../photo.jpg
   Règles :
    - tout le dossier 1 est conservé
    - les photos du dossier 2 sont ajoutées si leur contenu (SHA-256) est absent
    - même nom + contenu différent => les 2 sont gardées (suffixe _dossier2)
    - les chambres présentes dans un seul dossier sont reprises
   Sortie : un ZIP (les dossiers d'origine ne sont jamais touchés)
   ============================================================ */

const FUSION_EXT = /\.(jpe?g|png|gif|bmp|tiff?|webp|heic|heif|raw|cr2|nef|arw|dng)$/i;
const fusionState = { files1: [], files2: [], blob: null, merged: [] };

const fe = id => document.getElementById(id);

function fusionStatus(msg, cls = "") {
  const el = fe("fusion-status");
  if (!el) return;
  el.textContent = msg;
  el.style.color = cls === "err" ? "#f44336" : cls === "ok" ? "#2E7D32" : "#666";
}

async function fusionHash(file) {
  const buf = await file.arrayBuffer();
  const h = await crypto.subtle.digest("SHA-256", buf);
  return Array.from(new Uint8Array(h)).map(b => b.toString(16).padStart(2, "0")).join("");
}

// Regroupe les fichiers par chambre (1er sous-dossier sous la racine choisie)
function fusionGrouper(files) {
  const chambres = new Map(); // cle minuscule -> { nom, items: [{file, rel}] }
  for (const f of files) {
    if (!FUSION_EXT.test(f.name)) continue;
    const segs = (f.webkitRelativePath || "").replace(/\\/g, "/").split("/").filter(Boolean);
    if (segs.length < 3) continue;            // racine/chambre/photo minimum
    const nom = segs[1];
    const rel = segs.slice(2).join("/");
    const cle = nom.toLowerCase();
    if (!chambres.has(cle)) chambres.set(cle, { nom, items: [] });
    chambres.get(cle).items.push({ file: f, rel });
  }
  for (const c of chambres.values()) c.items.sort((a, b) => a.rel.localeCompare(b.rel, undefined, { numeric: true }));
  return chambres;
}

function fusionNomLibre(path, suffixe, utilises) {
  const i = path.lastIndexOf(".");
  const base = i > path.lastIndexOf("/") ? path.slice(0, i) : path;
  const ext = i > path.lastIndexOf("/") ? path.slice(i) : "";
  let cand = `${base}${suffixe}${ext}`;
  let n = 2;
  while (utilises.has(cand.toLowerCase())) cand = `${base}${suffixe}${n++}${ext}`;
  return cand;
}

async function fusionner() {
  const ch1 = fusionGrouper(fusionState.files1);
  const ch2 = fusionGrouper(fusionState.files2);
  const toutes = [...new Set([...ch1.keys(), ...ch2.keys()])].sort();
  const zip = new JSZip();
  const journal = [];
  const merged = [];
  const stats = { d1: 0, d2: 0, doublons: 0, renommees: 0, chambres: toutes.length };

  let done = 0;
  for (const cle of toutes) {
    const c1 = ch1.get(cle), c2 = ch2.get(cle);
    const nom = (c1 || c2).nom;
    journal.push(`\n=== ${nom} ===`);
    const hashes = new Set();
    const utilises = new Set();

    const ajouter = (file, rel) => {
      const path = `${nom}/${rel}`;
      zip.file(path, file);
      utilises.add(path.toLowerCase());
      merged.push({ file, path });
      return path;
    };

    if (c1) for (const { file, rel } of c1.items) {
      const h = await fusionHash(file);
      if (hashes.has(h)) { stats.doublons++; journal.push(`  = (déjà là)  ${rel}`); continue; }
      let r = rel;
      if (utilises.has(`${nom}/${r}`.toLowerCase())) { r = fusionNomLibre(`${nom}/${r}`, "_dossier1", utilises).slice(nom.length + 1); stats.renommees++; }
      ajouter(file, r);
      hashes.add(h); stats.d1++; journal.push(`  + D1         ${rel}`);
    }

    if (c2) for (const { file, rel } of c2.items) {
      const h = await fusionHash(file);
      if (hashes.has(h)) { stats.doublons++; journal.push(`  = (existe)   ${rel}`); continue; }
      let r = rel;
      if (utilises.has(`${nom}/${r}`.toLowerCase())) {
        r = fusionNomLibre(`${nom}/${r}`, "_dossier2", utilises).slice(nom.length + 1);
        stats.renommees++;
        journal.push(`  ! même nom, contenu différent -> ${r}`);
      }
      ajouter(file, r);
      hashes.add(h); stats.d2++; journal.push(`  + D2         ${rel}`);
    }

    fusionStatus(`Fusion... ${++done}/${toutes.length} chambres`);
  }

  zip.file("journal_fusion.txt", journal.join("\n"));
  fusionState.blob = await zip.generateAsync({ type: "blob" });
  fusionState.merged = merged;
  return stats;
}

/* ---------- Interface ---------- */
function fusionCheckReady() {
  fe("btn-fusion").disabled = !(fusionState.files1.length && fusionState.files2.length);
}

fe("fusion-d1").addEventListener("change", e => {
  fusionState.files1 = Array.from(e.target.files);
  fusionCheckReady();
});
fe("fusion-d2").addEventListener("change", e => {
  fusionState.files2 = Array.from(e.target.files);
  fusionCheckReady();
});

fe("btn-fusion").addEventListener("click", async () => {
  const btn = fe("btn-fusion");
  btn.disabled = true;
  fe("fusion-result").classList.add("hidden");
  try {
    if (!window.crypto || !crypto.subtle) throw new Error("SHA-256 indisponible (ouvrez l'appli en HTTPS)");
    fusionStatus("Fusion en cours...");
    const st = await fusionner();
    fusionStatus(
      `✓ ${st.chambres} chambres — ${st.d1} photos du dossier 1, ${st.d2} ajoutées du dossier 2, ` +
      `${st.doublons} doublons ignorés, ${st.renommees} renommées.`, "ok");
    fe("fusion-result").classList.remove("hidden");
  } catch (err) {
    console.error(err);
    fusionStatus("❌ " + err.message, "err");
  }
  btn.disabled = false;
});

fe("btn-fusion-dl").addEventListener("click", () => {
  if (!fusionState.blob) return;
  const url = URL.createObjectURL(fusionState.blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Photos_fusionnees_${new Date().toISOString().slice(0, 10)}.zip`;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 200);
});

// Utilise directement le résultat comme photos du Snake (sans repasser par le ZIP)
fe("btn-fusion-use").addEventListener("click", async () => {
  if (!fusionState.merged.length) return;
  const btn = fe("btn-fusion-use");
  btn.disabled = true;
  try {
    const stubs = [];
    let n = 0;
    for (const { file, path } of fusionState.merged) {
      if (!/\.(jpe?g|png|webp|bmp)$/i.test(file.name)) continue;
      const b64 = await compressedBase64(file);
      if (!b64) continue;
      const stub = { name: file.name, webkitRelativePath: `PHOTOS/${path}` };
      photoCache.set(stub, b64);
      stubs.push(stub);
      if (++n % 20 === 0) fusionStatus(`Préparation des photos... ${n}/${fusionState.merged.length}`);
    }
    state.photoFiles = state.photoFiles.concat(stubs);
    state.photoIndex = indexPhotos(state.photoFiles);
    fusionStatus(`✓ ${state.photoIndex.length} photos indexées au total pour la génération.`, "ok");
  } catch (err) {
    console.error(err);
    fusionStatus("❌ " + err.message, "err");
  }
  btn.disabled = false;
});
