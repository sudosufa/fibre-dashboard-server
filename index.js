/* ==========================================================================
   Serveur backend minimal pour le Dashboard Fibre Sonatel.
   - GET  /api/data    : renvoie les dernières données importées (public, lecture seule)
   - POST /api/upload  : reçoit le(s) fichier(s) .xlsx et met à jour les données
                         (protégé par un mot de passe admin)
   Stockage : Upstash Redis (gratuit), pour que les données survivent aux
   redémarrages du serveur gratuit (le disque de Render free n'est pas persistant).
   ========================================================================== */

const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fetch = require('node-fetch');
const { buildDataFromWorkbook, buildPBOZeroData, buildPenetrationBelow50Data } = require('./logic.js');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 30 * 1024 * 1024 } });

const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || '';
const UPSTASH_URL = process.env.UPSTASH_REDIS_REST_URL || '';
const UPSTASH_TOKEN = process.env.UPSTASH_REDIS_REST_TOKEN || '';
const DATA_KEY = 'fibre_dashboard_latest_data';

if (!ADMIN_PASSWORD) console.warn('[ATTENTION] ADMIN_PASSWORD non défini — /api/upload sera inaccessible tant que ce n\'est pas configuré.');
if (!UPSTASH_URL || !UPSTASH_TOKEN) console.warn('[ATTENTION] UPSTASH_REDIS_REST_URL / UPSTASH_REDIS_REST_TOKEN non définis — le stockage ne fonctionnera pas.');

// --- CORS (le dashboard HTML peut être ouvert depuis n'importe où) ---
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

app.use(express.json());

async function upstash(command) {
  const res = await fetch(UPSTASH_URL, {
    method: 'POST',
    headers: { Authorization: `Bearer ${UPSTASH_TOKEN}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(command),
  });
  const json = await res.json();
  if (json.error) throw new Error('Upstash: ' + json.error);
  return json.result;
}

const path = require('path');

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'Dashboard_Fibre_import_auto.html'));
});

app.get('/api/data', async (req, res) => {
  try {
    const raw = await upstash(['GET', DATA_KEY]);
    if (!raw) return res.status(404).json({ error: 'Aucune donnée importée pour le moment.' });
    res.json(JSON.parse(raw));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Erreur serveur : ' + err.message });
  }
});

app.post(
  '/api/upload',
  upload.fields([
    { name: 'mainFile', maxCount: 1 }, // Suivi_Taux_d_occupation_de_la_fibre_...xlsx
    { name: 'pboFile', maxCount: 1 },  // Fichier_Du_Taux_Pénétration_...xlsx (optionnel)
  ]),
  async (req, res) => {
    try {
      if (!ADMIN_PASSWORD) return res.status(503).json({ error: "ADMIN_PASSWORD n'est pas configuré côté serveur." });
      if (req.body.password !== ADMIN_PASSWORD) return res.status(401).json({ error: 'Mot de passe incorrect.' });

      const mainFile = req.files && req.files.mainFile && req.files.mainFile[0];
      const pboFile = req.files && req.files.pboFile && req.files.pboFile[0];
      if (!mainFile && !pboFile) return res.status(400).json({ error: 'Aucun fichier reçu.' });

      let data;
      if (mainFile) {
        const wbMain = XLSX.read(mainFile.buffer, { type: 'buffer', cellDates: false });
        const sheetNames = new Set(wbMain.SheetNames);
        if (!sheetNames.has('TO_Plaques') || !sheetNames.has('TO_Communes')) {
          return res.status(400).json({ error: "Le fichier principal ne contient pas les feuilles attendues (TO_Plaques / TO_Communes)." });
        }
        data = buildDataFromWorkbook(wbMain);
        data.isSeed = false;
      } else {
        // on ne met à jour que le fichier PBO/pénétration : on repart des dernières données connues
        const prevRaw = await upstash(['GET', DATA_KEY]);
        if (!prevRaw) return res.status(400).json({ error: "Aucune donnée principale existante — importez d'abord le fichier Suivi_Taux_d_occupation...xlsx." });
        data = JSON.parse(prevRaw);
      }

      if (pboFile) {
        const wbPbo = XLSX.read(pboFile.buffer, { type: 'buffer', cellDates: false });
        const pboSheets = new Set(wbPbo.SheetNames);
        if (pboSheets.has('PBO_0_Client')) data.pbo_zero = buildPBOZeroData(wbPbo);
        if (pboSheets.has('Pénétration < à 50%')) data.pen50 = buildPenetrationBelow50Data(wbPbo);
      } else if (mainFile) {
        // fichier principal réimporté seul : on garde le PBO/pen50 déjà connus
        try {
          const prevRaw = await upstash(['GET', DATA_KEY]);
          if (prevRaw) {
            const prev = JSON.parse(prevRaw);
            if (prev.pbo_zero) data.pbo_zero = prev.pbo_zero;
            if (prev.pen50) data.pen50 = prev.pen50;
          }
        } catch (e) { /* pas grave si indisponible */ }
      }

      await upstash(['SET', DATA_KEY, JSON.stringify(data)]);
      res.json({ ok: true, week: data.week, plaques: data.plaques.length });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'Erreur serveur : ' + err.message });
    }
  }
);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Fibre Dashboard API sur le port ${PORT}`));
