// Serves the add-in files to Word over https://localhost:4810 using the Office dev certificate.
import https from 'node:https';
import { readFile, writeFile, stat } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join, extname, dirname, basename } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = dirname(fileURLToPath(import.meta.url));
const certs = join(homedir(), '.office-addin-dev-certs');
const types = { '.html': 'text/html', '.js': 'text/javascript', '.png': 'image/png', '.xml': 'text/xml' };
const allowed = new Set(['taskpane.html', 'functions.html', 'harper.bundle.js', 'icon-16.png', 'icon-32.png', 'icon-80.png', 'manifest.xml']);

let tls;
try {
  tls = { key: await readFile(join(certs, 'localhost.key')), cert: await readFile(join(certs, 'localhost.crt')) };
} catch {
  console.log('Certificate not found. Run Setup.cmd first.');
  process.exit(1);
}

// Your own words, one per line; you can edit this file by hand.
const dictionary = join(root, 'dictionary.txt');
const validWord = /^[\p{L}\p{N}][\p{L}\p{N}'’.\-]{0,63}$/u;

// Hand-edited files may have a BOM (Notepad), stray spaces or blank lines.
const readWords = async () => (await readFile(dictionary, 'utf8').catch(() => '')).replace(/^﻿/, '')
  .split(/\r?\n/).map(w => w.trim()).filter(Boolean);
let writes = Promise.resolve(); // one dictionary write at a time, so an add during a remove isn't lost

async function handleDictionary(req, res) {
  if (req.method === 'GET') {
    res.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-cache' }).end((await readWords()).join('\r\n'));
    return;
  }
  // The custom header makes browsers block other websites from posting here (it forces a CORS preflight we never allow).
  if (!['POST', 'DELETE'].includes(req.method) || req.headers['x-harper'] !== '1') { res.writeHead(403).end(); return; }
  let body = '';
  for await (const chunk of req) { body += chunk; if (body.length > 200) { res.writeHead(413).end(); return; } }
  const word = body.trim();
  // removing only needs an exact match, so words typed into the file by hand can be removed too
  if (req.method === 'POST' ? !validWord.test(word) : !word) { res.writeHead(400).end(); return; }
  const op = writes.then(async () => {
    const words = await readWords();
    if (req.method === 'DELETE') {
      if (words.includes(word)) await writeFile(dictionary, words.filter(w => w !== word).map(w => w + '\r\n').join(''), 'utf8');
    } else if (!words.includes(word)) await writeFile(dictionary, [...words, word].map(w => w + '\r\n').join(''), 'utf8');
  });
  writes = op.catch(() => {});
  await op;
  res.writeHead(204).end();
}

// Errors (a locked dictionary.txt, an aborted request) answer 500 instead of crashing the server.
const server = https.createServer(tls, (req, res) => handle(req, res).catch(() => {
  if (!res.headersSent) res.writeHead(500);
  res.end();
}));

async function handle(req, res) {
  const name = basename(new URL(req.url, 'https://localhost').pathname);
  if (name === 'dictionary') return handleDictionary(req, res);
  if (!allowed.has(name)) { res.writeHead(404).end(); return; }
  // ETag lets Word revalidate with a 304 instead of re-reading and re-sending the 27 MB bundle on every panel load
  const { size, mtimeMs } = await stat(join(root, name));
  const etag = `"${size}-${mtimeMs}"`;
  if (req.headers['if-none-match'] === etag) { res.writeHead(304, { ETag: etag }).end(); return; }
  res.writeHead(200, { 'Content-Type': types[extname(name)], 'Cache-Control': 'no-cache', ETag: etag });
  res.end(await readFile(join(root, name)));
}

server.on('error', e => {
  if (e.code === 'EADDRINUSE') { console.log('Harper is already running.'); process.exit(0); }
  throw e;
});
server.listen(4810, '127.0.0.1', () => console.log('Harper for Word is running. Keep this window open while you use Word.'));
