// Dev only: serves the add-in next to a fake Word document so the panel can be debugged in a normal browser.
// Usage: node dev/serve.mjs [add-in folder] [port]   then open http://127.0.0.1:<port>/
import http from 'node:http';
import { readFile } from 'node:fs/promises';
import { join, dirname, resolve, basename, extname } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const addin = resolve(process.argv[2] || join(here, '..', 'word-addin'));
const port = Number(process.argv[3]) || 4820;
const types = { '.html': 'text/html', '.js': 'text/javascript', '.png': 'image/png', '.css': 'text/css' };
const words = []; // in-memory stand-in for dictionary.txt

http.createServer(async (req, res) => {
  const name = basename(new URL(req.url, 'http://x').pathname) || 'index.html';
  try {
    if (name === 'dictionary') {
      if (req.method === 'GET') return res.writeHead(200, { 'Content-Type': 'text/plain' }).end(words.join('\n'));
      if (req.headers['x-harper'] !== '1') return res.writeHead(403).end();
      let body = ''; for await (const c of req) body += c;
      const i = words.indexOf(body.trim());
      if (req.method === 'DELETE') { if (i !== -1) words.splice(i, 1); }
      else if (i === -1) words.push(body.trim());
      return res.writeHead(204).end();
    }
    if (name === 'index.html' || name === 'fake-word.js') {
      return res.writeHead(200, { 'Content-Type': types[extname(name)], 'Cache-Control': 'no-cache' }).end(await readFile(join(here, name)));
    }
    let body = await readFile(join(addin, name));
    if (name === 'taskpane.html') {
      // swap the real office.js (it blanks the frame outside Office) for the fake Office + Word API
      body = body.toString().replace(/<script src="https:\/\/appsforoffice[^"]*office\.js"><\/script>/, '<script src="fake-word.js"></script>');
    }
    res.writeHead(200, { 'Content-Type': types[extname(name)] || 'application/octet-stream', 'Cache-Control': 'no-cache' }).end(body);
  } catch { res.writeHead(404).end(); }
}).listen(port, '127.0.0.1', () => console.log(`Fake Word: http://127.0.0.1:${port}/  (add-in: ${addin})`));
