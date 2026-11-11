#!/usr/bin/env node
'use strict';
// Modified in this fork (roelschouten/local-modelling-tool): adds /api/concept-models.

const express = require('express');
const path = require('path');
const fs = require('fs');
const zlib = require('zlib');

const app = express();
const PORT = process.env.PORT || 8000;

const publicDir = path.join(__dirname, 'public');
const indexHtml = path.join(publicDir, 'index.html');

app.get('/api/health', (req, res) => {
  res.json({ ok: true });
});

// The bundled model ships gzipped (public/model.json.gz, ~3 MB) to keep the
// deployed artifact small. A raw public/model.json is also honoured if present.
const rawModelPath = path.join(publicDir, 'model.json');
const gzModelPath = path.join(publicDir, 'model.json.gz');
let gzBytesCache = null, rawBytesCache = null;
function hasBundledModel() { return fs.existsSync(gzModelPath) || fs.existsSync(rawModelPath); }
function gzBytes() { if (!gzBytesCache && fs.existsSync(gzModelPath)) gzBytesCache = fs.readFileSync(gzModelPath); return gzBytesCache; }

app.get('/model.json', (req, res) => {
  const acceptsGzip = (req.headers['accept-encoding'] || '').includes('gzip');
  res.set('Content-Type', 'application/json; charset=utf-8');
  res.set('Cache-Control', 'public, max-age=300');
  if (fs.existsSync(gzModelPath)) {
    if (acceptsGzip) { res.set('Content-Encoding', 'gzip'); return res.end(gzBytes()); }
    if (!rawBytesCache) rawBytesCache = zlib.gunzipSync(gzBytes());
    return res.end(rawBytesCache);
  }
  if (fs.existsSync(rawModelPath)) {
    if (acceptsGzip) {
      if (!gzBytesCache) gzBytesCache = zlib.gzipSync(fs.readFileSync(rawModelPath));
      res.set('Content-Encoding', 'gzip'); return res.end(gzBytesCache);
    }
    return res.sendFile(rawModelPath);
  }
  res.status(404).json({ error: 'not found' });
});

app.get('/api/bundled-models', (req, res) => {
  const models = [];
  if (hasBundledModel()) models.push({ path: '/model.json', label: 'ddx_gtm_small2_v7' });
  res.json({ models });
});

// Fork addition: YAML domain models for the Concepts / Processes views.
// Every *.yaml / *.yml under CONCEPT_MODELS_DIR (default ./concept-models) is
// returned as text; the browser parses them. Read-only, size-capped.
const conceptDir = path.resolve(process.env.CONCEPT_MODELS_DIR || path.join(__dirname, 'concept-models'));
const CONCEPT_MAX_BYTES = 5 * 1024 * 1024;
function listYaml(dir, base, out) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.name.startsWith('.')) continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) listYaml(full, base, out);
    else if (/\.ya?ml$/i.test(entry.name)) out.push(full);
  }
  return out;
}
app.get('/api/concept-models', (req, res) => {
  if (!fs.existsSync(conceptDir)) return res.json({ files: [] });
  try {
    let total = 0;
    const files = [];
    for (const full of listYaml(conceptDir, conceptDir, []).sort()) {
      const text = fs.readFileSync(full, 'utf8');
      total += Buffer.byteLength(text);
      if (total > CONCEPT_MAX_BYTES) break;
      files.push({ name: path.relative(conceptDir, full).split(path.sep).join('/'), text });
    }
    res.set('Cache-Control', 'no-cache');
    res.json({ files });
  } catch (e) {
    res.status(500).json({ error: 'could not read concept models' });
  }
});

app.use('/static', express.static(path.join(publicDir, 'static')));
app.get(/.*/, (req, res) => res.sendFile(indexHtml));

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Model Viewer running on http://0.0.0.0:${PORT}`);
});
