// 로컬 확인용 서버. 정적 파일과 가짜 Apps Script를 묶어 제공한다. 데이터는 메모리에만 있다.
//   POST /__fail?on=1   /api가 500을 돌려주게 한다 (?on=0으로 복구)
//   POST /__bad-row     형식이 틀린 행을 시트에 넣는다 (skipped 경고 확인용)
//   POST /__clock?advance=ms  가짜 서버의 시계를 앞으로 돌린다 (비밀번호 잠금 해제 확인용)

import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { extname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createFakeAppsScript } from './fake-apps-script.mjs';

// `node --test`가 폴더를 훑어 이 파일을 실행해도 서버를 띄워 멈추지 않게 한다.
if (process.env.NODE_TEST_CONTEXT) {
  process.exit(0);
}

const PORT = 4173;
const HOST = '127.0.0.1';
const ROOT = fileURLToPath(new URL('..', import.meta.url));
const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.md': 'text/markdown; charset=utf-8',
  '.gs': 'text/plain; charset=utf-8',
};

const script = createFakeAppsScript({ password: 'test-password' });
script.setup();
let failing = false;

function send(response, status, contentType, body) {
  response.writeHead(status, { 'Content-Type': contentType, 'Cache-Control': 'no-store' });
  response.end(body);
}

const sendText = (response, status, text) => send(response, status, 'text/plain; charset=utf-8', text);

async function readBody(request) {
  const chunks = [];
  for await (const chunk of request) {
    chunks.push(chunk);
  }
  return Buffer.concat(chunks).toString('utf8');
}

async function handleApi(request, response, url) {
  if (failing) {
    sendText(response, 500, '일부러 낸 고장입니다');
    return;
  }
  let output;
  if (request.method === 'GET') {
    output = script.doGet(Object.fromEntries(url.searchParams));
  } else if (request.method === 'POST') {
    output = script.doPost(await readBody(request));
  } else {
    sendText(response, 405, '허용되지 않는 메서드');
    return;
  }
  send(response, 200, `${output.getMimeType()}; charset=utf-8`, output.getContent());
}

async function handleStatic(response, pathname) {
  const relative = pathname === '/' ? 'index.html' : `.${decodeURIComponent(pathname)}`;
  const filePath = resolve(ROOT, relative);
  if (!filePath.startsWith(ROOT)) {
    sendText(response, 403, '허용되지 않는 경로');
    return;
  }
  const contentType = MIME_TYPES[extname(filePath)] ?? 'application/octet-stream';
  send(response, 200, contentType, await readFile(filePath));
}

const server = createServer(async (request, response) => {
  const url = new URL(request.url, `http://${HOST}:${PORT}`);
  try {
    if (url.pathname === '/config.js') {
      send(response, 200, MIME_TYPES['.js'], `export const apiUrl = '/api';\n`);
    } else if (url.pathname === '/api') {
      await handleApi(request, response, url);
    } else if (url.pathname === '/__fail' && request.method === 'POST') {
      failing = url.searchParams.get('on') === '1';
      sendText(response, 200, `failing=${failing}`);
    } else if (url.pathname === '/__bad-row' && request.method === 'POST') {
      const sheet = script.sheet();
      sheet.humanRow(sheet.getLastRow() + 1, ['bad-row', '내일', '가나다', '많이', '']);
      sendText(response, 200, 'ok');
    } else if (url.pathname === '/__clock' && request.method === 'POST') {
      script.advanceClock(Number(url.searchParams.get('advance')));
      sendText(response, 200, 'ok');
    } else {
      await handleStatic(response, url.pathname);
    }
  } catch (error) {
    // 진짜 Apps Script도 처리 못 한 예외는 JSON이 아닌 응답이 된다. 화면에는 실패로 보여야 한다.
    const status = error.code === 'ENOENT' || error.code === 'EISDIR' ? 404 : 500;
    sendText(response, status, status === 404 ? '찾을 수 없음' : `서버 오류: ${error.message}`);
  }
});

server.listen(PORT, HOST, () => {
  console.log(`http://${HOST}:${PORT} 에서 실행 중 (비밀번호: test-password, 데이터는 메모리에만)`);
});
