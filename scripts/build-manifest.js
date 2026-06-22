#!/usr/bin/env node
/**
 * Generates an environment-specific Outlook add-in manifest from the
 * localhost-based source template (src/addin/manifest.xml) by substituting the
 * add-in host into every URL, AppDomain, and the WebApplicationInfo App ID URI.
 *
 * Environments (match infra/envs/<env>.tfvars and src/addin/.env.<env>):
 *   local — add-in served from https://localhost:3000 (Vite dev server),
 *           functions run locally via `func start`.   → manifest.local.xml
 *   dev   — add-in served from the Azure Static Web App,
 *           functions deployed on Azure.              → manifest.dev.xml
 *
 *   (Service Bus and the Storage Account are always in Azure in both envs;
 *    only the add-in host and the Function App location differ.)
 *
 * The add-in host for each env is resolved, in priority order, from:
 *   1. a --host <hostname> flag
 *   2. VITE_API_SCOPE in src/addin/.env.<env>  (api://<host>/<client-id>/...)
 *   3. localhost:3000  (local only)
 * so the .env files stay the single source of truth.
 *
 * Usage:
 *   node scripts/build-manifest.js local
 *   node scripts/build-manifest.js dev
 *   node scripts/build-manifest.js dev --host proud-pond-0123.azurestaticapps.net
 */

const fs = require('fs');
const path = require('path');

// The host baked into the source manifest (manifest.xml). All env builds are
// produced by replacing this token.
const SOURCE_HOST = 'localhost:3000';
const ENVIRONMENTS = ['local', 'dev'];

const env = (process.argv[2] || '').toLowerCase();
if (!ENVIRONMENTS.includes(env)) {
  console.error(`Usage: node scripts/build-manifest.js <${ENVIRONMENTS.join('|')}> [--host <hostname>]`);
  process.exit(1);
}

const root = path.resolve(__dirname, '..');
const addinDir = path.join(root, 'src', 'addin');
const srcXml = path.join(addinDir, 'manifest.xml');
const outXml = path.join(addinDir, `manifest.${env}.xml`);

// ── Resolve the add-in host for this environment ──────────────────────────────
const host =
  flagValue('--host') ??
  hostFromEnvFile(path.join(addinDir, `.env.${env}`)) ??
  (env === 'local' ? SOURCE_HOST : undefined);

if (!host) {
  console.error(`Error: could not determine the add-in host for "${env}".`);
  console.error(`  Fill VITE_API_SCOPE in src/addin/.env.${env} (api://<host>/<client-id>/access_as_user),`);
  console.error(`  or pass it explicitly: node scripts/build-manifest.js ${env} --host <hostname>`);
  process.exit(1);
}
validateHost(host);

// ── Substitute and write ──────────────────────────────────────────────────────
if (!fs.existsSync(srcXml)) {
  console.error(`Error: source manifest not found at ${srcXml}`);
  process.exit(1);
}
const xml = fs.readFileSync(srcXml, 'utf8').replaceAll(SOURCE_HOST, host);
fs.writeFileSync(outXml, xml, 'utf8');

console.log(`Wrote ${path.relative(root, outXml)}  (env=${env}, host=${host})`);
if (host === SOURCE_HOST) {
  console.log('  host unchanged — identical to manifest.xml');
}
console.log('');
console.log('Next: sideload this manifest in Outlook (remove any previously sideloaded copy first).');
if (env === 'dev') {
  console.log(`Ensure the App Registration has, as Single-page application redirect URIs:`);
  console.log(`  https://${host}/commands.html`);
  console.log(`  brk-multihub://${host}   (NAA broker — origin only, no path)`);
}

// ── helpers ───────────────────────────────────────────────────────────────────
function flagValue(name) {
  const i = process.argv.indexOf(name);
  return i !== -1 ? process.argv[i + 1] : undefined;
}

function hostFromEnvFile(file) {
  if (!fs.existsSync(file)) return undefined;
  const text = fs.readFileSync(file, 'utf8');
  // api://<host>/<client-id>/...  — last definition wins (.env files may repeat keys)
  const matches = [...text.matchAll(/^\s*VITE_API_SCOPE\s*=\s*api:\/\/([^/\s]+)\//gim)];
  return matches.length ? matches[matches.length - 1][1] : undefined;
}

function validateHost(host) {
  if (host.startsWith('http://') || host.startsWith('https://')) {
    console.error('Error: host must be a bare hostname without scheme (e.g. localhost:3000 or proud-pond-0123.azurestaticapps.net).');
    process.exit(1);
  }
  if (host.includes('/')) {
    console.error('Error: host must be origin-only (no path).');
    process.exit(1);
  }
}
