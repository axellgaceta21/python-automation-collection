/* Dependency-free local preview of the static files allowed by vercel.json.
 * Run from any directory: node scripts/serve-portfolio.mjs
 * A generic repository-root server would also expose private/legacy folders.
 */
import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

export const portfolioRoot = fileURLToPath(new URL('../', import.meta.url));
const config = JSON.parse(await fs.readFile(path.join(portfolioRoot, 'vercel.json'), 'utf8'));
if (!config.builds?.length || config.builds.some(rule => rule.use !== '@vercel/static')) {
  throw new Error('Local preview requires an explicit static-file deployment allowlist.');
}

export function isPublicPath(relative) {
  const normalized = relative.replaceAll('\\', '/');
  return !normalized.split('/').some(segment => segment.startsWith('.') || !segment)
    && config.builds.some(rule => path.matchesGlob(normalized, rule.src));
}

const types = {
  '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8', '.png': 'image/png',
  '.pdf': 'application/pdf',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

export async function handlePortfolioRequest(req, res) {
  try {
    if (!['GET', 'HEAD'].includes(req.method)) {
      res.writeHead(405, { Allow: 'GET, HEAD' }); res.end(); return;
    }
    let relative = decodeURIComponent(new URL(req.url, 'http://localhost').pathname).slice(1);
    if (!relative || relative.endsWith('/')) relative += 'index.html';
    const filename = path.resolve(portfolioRoot, relative);
    const canonical = path.relative(portfolioRoot, filename).replaceAll('\\', '/');
    if (!isPublicPath(canonical) || !filename.startsWith(portfolioRoot)) throw new Error('Not public');
    const real = await fs.realpath(filename);
    if (real !== filename) throw new Error('Symlink not served');
    const stat = await fs.stat(filename);
    if (!stat.isFile()) throw new Error('Not a file');
    // Read once before headers so file access errors produce a plain 404.
    const body = req.method === 'HEAD' ? null : await fs.readFile(filename);
    res.writeHead(200, {
      'Content-Type': types[path.extname(filename)] || 'application/octet-stream',
      'Content-Length': stat.size, 'X-Content-Type-Options': 'nosniff'
    });
    res.end(body);
  } catch {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end(req.method === 'HEAD' ? undefined : 'Not found');
  }
}

export const createPortfolioServer = () => http.createServer(handlePortfolioRequest);

if (process.argv[1] && pathToFileURL(path.resolve(process.argv[1])).href === import.meta.url) {
  const server = createPortfolioServer();
  server.on('error', error => { console.error(error.message); process.exitCode = 1; });
  server.listen(5500, '127.0.0.1', () => console.log('Portfolio: http://127.0.0.1:5500/ (non-public folders excluded)'));
}
