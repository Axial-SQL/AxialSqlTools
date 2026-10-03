'use strict';

const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const test = require('node:test');
const { checkGallery } = require('./gallery-publish');

const openGallery = 'https://www.vsixgallery.com';
const ssmsGallery = 'https://ssmsgallery.azurewebsites.net';
const extensionId = 'AxialSqlTools';
const vsix = Buffer.from('the exact signed VSIX bytes');
const sha256 = crypto.createHash('sha256').update(vsix).digest('hex');
const differentHash = crypto.createHash('sha256').update('another VSIX').digest('hex');
const downloadLink = '/extensions/AxialSqlTools/Axial%20SQL%20Tools%20v4.14.vsix';

function jsonResponse(value, status = 200) {
  return new Response(JSON.stringify(value), { status, headers: { 'Content-Type': 'application/json' } });
}

function mockFetch(responses) {
  const calls = [];
  const fetchImpl = async (url, options) => {
    calls.push({ url, options });
    assert.ok(options.signal instanceof AbortSignal, 'Requests must have a timeout signal.');
    assert.equal(options.redirect, 'error', 'Redirects must not leave the selected gallery.');
    assert.equal(options.method, 'GET', 'The preflight must never upload or modify a gallery.');
    assert.equal(options.cache, 'no-store');
    const response = responses[calls.length - 1];
    assert.ok(response, `Unexpected request to ${url}`);
    if (response instanceof Error) throw response;
    return response;
  };
  return { fetchImpl, calls };
}

async function check(responses, overrides = {}) {
  const mock = mockFetch(responses);
  const result = await checkGallery({ galleryUrl: openGallery, extensionId, version: '4.14', sha256, ...overrides, fetchImpl: mock.fetchImpl });
  return { result, calls: mock.calls };
}

for (const galleryUrl of [openGallery, ssmsGallery]) {
  test(`a missing extension can be published to ${galleryUrl}`, async () => {
    const { result, calls } = await check([jsonResponse({}, 404)], { galleryUrl });
    assert.equal(result.publish, true);
    assert.equal(calls.length, 1);
    assert.equal(calls[0].url, `${galleryUrl}/api/AxialSqlTools`);
  });
}

test('a lower numeric gallery version can be replaced without downloading it', async () => {
  const { result, calls } = await check([jsonResponse({ id: extensionId, version: '4.9' })]);
  assert.equal(result.publish, true);
  assert.equal(calls.length, 1);
});

for (const remoteVersion of ['4.15', '4.100', '5.0', '4.14.0.1']) {
  test(`a retry cannot downgrade gallery version ${remoteVersion}`, async () => {
    await assert.rejects(check([jsonResponse({ id: extensionId, version: remoteVersion })]), /refusing to replace.*older version 4\.14/);
  });
}

test('an identical signed package is skipped even with normalized version and uppercase digest', async () => {
  const { result, calls } = await check([jsonResponse({ id: extensionId, version: '4.14.0.0', sha256: sha256.toUpperCase() })]);
  assert.equal(result.publish, false);
  assert.match(result.reason, /same signed VSIX SHA-256/);
  assert.equal(calls.length, 1);
});

test('different signed bytes under the same version require a version bump', async () => {
  await assert.rejects(check([jsonResponse({ id: extensionId, version: '4.14', sha256: differentHash })]), /Bump the product version/);
});

test('the older SSMS API can confirm identical bytes by downloading its VSIX', async () => {
  const { result, calls } = await check([
    jsonResponse({ id: extensionId, version: '4.14', downloadLink }),
    new Response(vsix),
  ], { galleryUrl: ssmsGallery });
  assert.equal(result.publish, false);
  assert.equal(calls.length, 2);
  assert.equal(calls[1].url, `${ssmsGallery}${downloadLink}`);
});

test('older PascalCase API fields and an unusable digest fall back to a download', async () => {
  const { result } = await check([
    jsonResponse({ ID: extensionId, Version: '4.14', Sha256: 'unavailable', DownloadLink: `${openGallery}${downloadLink}` }),
    new Response(vsix),
  ]);
  assert.equal(result.publish, false);
});

test('a fallback download with different bytes cannot overwrite the existing version', async () => {
  await assert.rejects(check([
    jsonResponse({ id: extensionId, version: '4.14', downloadLink }),
    new Response('different package'),
  ]), /Bump the product version/);
});

for (const status of [401, 429, 500]) {
  test(`a gallery HTTP ${status} response fails instead of being treated as absent`, async () => {
    await assert.rejects(check([jsonResponse({}, status)]), new RegExp(`Gallery lookup failed.*HTTP ${status}`));
  });
}

test('a missing gallery download fails instead of permitting replacement', async () => {
  await assert.rejects(check([
    jsonResponse({ id: extensionId, version: '4.14', downloadLink }),
    new Response('', { status: 404 }),
  ]), /VSIX download failed.*HTTP 404/);
});

for (const badLink of [
  'https://example.com/extension.vsix',
  '//ssmsgallery.azurewebsites.net/extension.vsix',
  'http://www.vsixgallery.com/extension.vsix',
  'https://someone:password@www.vsixgallery.com/extension.vsix',
]) {
  test(`a download link cannot redirect the preflight to ${new URL(badLink, openGallery).origin}`, async () => {
    await assert.rejects(check([jsonResponse({ id: extensionId, version: '4.14', downloadLink: badLink })]), /download URL must stay on/);
  });
}

for (const galleryUrl of ['https://example.com', 'http://www.vsixgallery.com', `${openGallery}/api`, `${openGallery}?query=value`]) {
  test(`an unsupported gallery URL is rejected before a request: ${galleryUrl}`, async () => {
    await assert.rejects(check([], { galleryUrl }), /Unsupported gallery URL/);
  });
}

test('a trailing slash on an allowed gallery URL resolves to the exact endpoint', async () => {
  const { result, calls } = await check([jsonResponse({}, 404)], { galleryUrl: `${openGallery}/` });
  assert.equal(result.publish, true);
  assert.equal(calls[0].url, `${openGallery}/api/AxialSqlTools`);
});

test('the local package must have a valid SHA-256 before querying the gallery', async () => {
  await assert.rejects(check([], { sha256: 'bad-hash' }), /64 hexadecimal/);
});

test('the package identifier cannot alter the API path', async () => {
  await assert.rejects(check([], { extensionId: '../upload' }), /extension ID is invalid/);
});

test('metadata for another package is rejected', async () => {
  await assert.rejects(check([jsonResponse({ id: 'OtherPackage', version: '4.13' })]), /does not identify AxialSqlTools/);
});

test('malformed gallery metadata cannot be treated as a successful publish', async () => {
  await assert.rejects(check([jsonResponse(null)]), /does not identify AxialSqlTools/);
  await assert.rejects(check([new Response('<html>not JSON</html>')]), SyntaxError);
});

test('an unknown remote version must be resolved before publishing', async () => {
  await assert.rejects(check([jsonResponse({ id: extensionId, version: 'preview' })]), /Unsupported release version/);
});

test('an equal version without a digest or downloadable package cannot be confirmed', async () => {
  await assert.rejects(check([jsonResponse({ id: extensionId, version: '4.14' })]), /does not provide a SHA-256 or download link/);
});

test('network errors and timeouts fail with the requested gallery URL', async () => {
  await assert.rejects(check([new Error('request timed out')]), /Could not read https:\/\/www\.vsixgallery\.com\/api\/AxialSqlTools: request timed out/);
});
