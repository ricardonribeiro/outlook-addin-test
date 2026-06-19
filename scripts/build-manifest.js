#!/usr/bin/env node
/**
 * Generates manifest.prod.xml from manifest.xml by substituting the production
 * add-in host for all localhost:3000 references (URLs, App ID URI, AppDomain).
 *
 * Usage:
 *   ADDIN_HOST=proud-pond-0123.azurestaticapps.net node scripts/build-manifest.js
 *   node scripts/build-manifest.js proud-pond-0123.azurestaticapps.net
 *
 * The host should be the bare hostname (no https:// prefix, no trailing slash).
 * After running, upload src/addin/manifest.prod.xml to the M365 Admin Center.
 */

const fs   = require('fs');
const path = require('path');

const host = process.env.ADDIN_HOST ?? process.argv[2];
if (!host) {
  console.error('Error: production host not provided.');
  console.error('  ADDIN_HOST=<hostname> node scripts/build-manifest.js');
  console.error('  node scripts/build-manifest.js <hostname>');
  process.exit(1);
}
if (host.startsWith('https://') || host.startsWith('http://')) {
  console.error('Error: provide the bare hostname without scheme (e.g. proud-pond-0123.azurestaticapps.net).');
  process.exit(1);
}

const root   = path.resolve(__dirname, '..');
const srcXml = path.join(root, 'src', 'addin', 'manifest.xml');
const outXml = path.join(root, 'src', 'addin', 'manifest.prod.xml');

const xml = fs.readFileSync(srcXml, 'utf8').replaceAll('localhost:3000', host);
fs.writeFileSync(outXml, xml, 'utf8');

console.log(`manifest.prod.xml written for host: ${host}`);
console.log(`  source : ${srcXml}`);
console.log(`  output : ${outXml}`);
console.log('');
console.log('Next steps:');
console.log('  1. Update your App Registration — add the production App ID URI:');
console.log(`       api://${host}/<client-id>`);
console.log('  2. Add the SWA hostname as a SPA redirect URI on the App Registration.');
console.log('  3. Upload manifest.prod.xml to M365 Admin Center → Settings → Integrated apps.');
