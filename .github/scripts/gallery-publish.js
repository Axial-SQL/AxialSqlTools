'use strict';

const crypto = require('node:crypto');
const { productVersion } = require('./release');

const galleries = new Set([
  'https://ssmsgallery.azurewebsites.net',
  'https://www.vsixgallery.com',
]);
const sha256Pattern = /^[a-f0-9]{64}$/i;

async function checkGallery({ galleryUrl, extensionId, version, sha256, fetchImpl = globalThis.fetch }) {
  const origin = typeof galleryUrl === 'string' ? galleryUrl.replace(/\/$/, '') : '';
  if (!galleries.has(origin)) throw new Error(`Unsupported gallery URL: ${galleryUrl}`);
  if (typeof extensionId !== 'string' || !/^[a-z0-9._-]+$/i.test(extensionId)) {
    throw new Error('The gallery extension ID is invalid.');
  }
  if (typeof sha256 !== 'string' || !sha256Pattern.test(sha256)) {
    throw new Error('The signed VSIX SHA-256 must contain 64 hexadecimal characters.');
  }
  const currentVersion = productVersion(version);
  const expectedHash = sha256.toLowerCase();

  async function get(url, accept) {
    try {
      return await fetchImpl(url, {
        method: 'GET',
        headers: { Accept: accept, 'Cache-Control': 'no-cache' },
        cache: 'no-store',
        redirect: 'error',
        signal: AbortSignal.timeout(60_000),
      });
    } catch (error) {
      throw new Error(`Could not read ${url}: ${error.message}`, { cause: error });
    }
  }

  const apiUrl = `${origin}/api/${encodeURIComponent(extensionId)}`;
  const response = await get(apiUrl, 'application/json');
  if (response.status === 404) {
    return { publish: true, reason: `${extensionId} is not listed at ${origin}.` };
  }
  if (!response.ok) throw new Error(`Gallery lookup failed at ${origin}: HTTP ${response.status}.`);

  const metadata = await response.json();
  if (!metadata || Array.isArray(metadata) || (metadata.id ?? metadata.ID) !== extensionId) {
    throw new Error(`Gallery metadata at ${origin} does not identify ${extensionId}.`);
  }
  const remoteVersionText = metadata.version ?? metadata.Version;
  const remoteVersion = productVersion(remoteVersionText);
  const difference = currentVersion.map((part, index) => part - remoteVersion[index]).find(part => part !== 0) || 0;
  if (difference < 0) {
    throw new Error(`${origin} already has version ${remoteVersionText}; refusing to replace it with older version ${version}.`);
  }
  if (difference > 0) {
    return { publish: true, reason: `${origin} has version ${remoteVersionText}; publishing ${version}.` };
  }

  let remoteHash = metadata.sha256 ?? metadata.Sha256;
  if (typeof remoteHash !== 'string' || !sha256Pattern.test(remoteHash)) {
    // The older SSMS gallery does not include a digest in its API response.
    // Compare the stored VSIX itself before treating a retry as complete.
    const downloadLink = metadata.downloadLink ?? metadata.DownloadLink;
    if (typeof downloadLink !== 'string' || !downloadLink) {
      throw new Error(`${origin} does not provide a SHA-256 or download link for version ${version}.`);
    }
    const downloadUrl = new URL(downloadLink, `${origin}/`);
    if (downloadUrl.origin !== origin || downloadUrl.username || downloadUrl.password || downloadUrl.hash) {
      throw new Error(`The gallery download URL must stay on ${origin}.`);
    }
    const download = await get(downloadUrl.href, 'application/octet-stream');
    if (!download.ok) throw new Error(`Gallery VSIX download failed at ${origin}: HTTP ${download.status}.`);
    remoteHash = crypto.createHash('sha256').update(Buffer.from(await download.arrayBuffer())).digest('hex');
  }

  if (remoteHash.toLowerCase() !== expectedHash) {
    throw new Error(`${origin} already has different VSIX bytes for version ${version}. Bump the product version before publishing changed files.`);
  }
  return { publish: false, reason: `${origin} already has version ${version} with the same signed VSIX SHA-256.` };
}

module.exports = { checkGallery };
