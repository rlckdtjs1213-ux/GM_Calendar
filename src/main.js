'use strict';

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const chokidar = require('chokidar');
const XLSX = require('xlsx');

// ── 설정 파일 (userData 폴더에 저장: 팀원마다 공유폴더 경로가 다를 수 있음) ──
const CONFIG_PATH = path.join(app.getPath('userData'), 'config.json');
const TEAMS = ['GMK', 'GMAP']; // 토글 대상 팀

function defaultConfig() {
  return { active: 'GMK', teams: { GMK: { watchDir: '' }, GMAP: { watchDir: '' } } };
}

function loadConfig() {
  let raw;
  try {
    raw = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf-8'));
  } catch {
    return defaultConfig();
  }
  // 구버전 마이그레이션: { watchDir } → { active, teams }
  if (raw && typeof raw.watchDir === 'string' && !raw.teams) {
    return { active: 'GMK', teams: { GMK: { watchDir: raw.watchDir }, GMAP: { watchDir: '' } } };
  }
  // 누락 필드 보정
  const cfg = defaultConfig();
  if (raw.active && TEAMS.includes(raw.active)) cfg.active = raw.active;
  for (const t of TEAMS) {
    if (raw.teams && raw.teams[t] && typeof raw.teams[t].watchDir === 'string') {
      cfg.teams[t].watchDir = raw.teams[t].watchDir;
    }
  }
  return cfg;
}

function activeDir(config) {
  return config.teams[config.active].watchDir;
}

function saveConfig(config) {
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2), 'utf-8');
}

let mainWindow = null;
let watcher = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 820,
    minWidth: 900,
    minHeight: 600,
    title: '배포일정 캘린더',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.removeMenu();
  mainWindow.loadFile(path.join(__dirname, 'index.html'));
}

// ── 엑셀 파싱 ──────────────────────────────────────────────
// 기대 컬럼: 날짜 / 제목 / 담당자 / 비고
const COLUMN_ALIASES = {
  date: ['날짜', '배포날짜', '배포일', 'date', '일자'],
  title: ['제목', '시스템', '시스템명', '내용', 'title'],
  owner: ['담당자', '담당', 'owner', '작성자'],
  note: ['비고', '메모', 'note', '설명'],
};

function pickColumn(row, aliases) {
  for (const key of Object.keys(row)) {
    const norm = String(key).trim().toLowerCase();
    if (aliases.some((a) => a.toLowerCase() === norm)) return row[key];
  }
  return '';
}

// 엑셀 날짜를 YYYY-MM-DD 문자열로 정규화
function normalizeDate(value) {
  if (value == null || value === '') return null;

  // 엑셀 직렬 숫자 (날짜 셀)
  if (typeof value === 'number') {
    const d = XLSX.SSF.parse_date_code(value);
    if (!d) return null;
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }

  const s = String(value).trim();
  // 2026-06-16 / 2026.06.16 / 2026/6/16 등 모두 허용
  const m = s.match(/(\d{4})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})/);
  if (m) {
    return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  }
  const parsed = new Date(s);
  if (!isNaN(parsed)) {
    return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, '0')}-${String(parsed.getDate()).padStart(2, '0')}`;
  }
  return null;
}

// CSV는 인코딩이 제각각(한국 Excel은 CP949로 저장) → BOM/UTF-8/euc-kr 순으로 판별
function decodeCsv(buf) {
  if (buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) {
    return new TextDecoder('utf-8').decode(buf); // UTF-8 BOM
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(buf); // 정상 UTF-8
  } catch {
    return new TextDecoder('euc-kr').decode(buf); // CP949 폴백
  }
}

function readWorkbook(filePath) {
  if (/\.csv$/i.test(filePath)) {
    return XLSX.read(decodeCsv(fs.readFileSync(filePath)), { type: 'string' });
  }
  return XLSX.readFile(filePath);
}

// ── RFC 시트 파서 (GMK 기획팀 RFC 관리 시트 형식) ───────────
// 섹션헤더 예) "<6/30(화) 09시 LIVE> 6월 16일(화) RFC&UAT (6/30(화) 09시~7/1(수) 19시)"
//   - 괄호 안에 M/D~M/D 범위가 있으면 그 기간으로 연결
//   - 없으면 <M/D LIVE>의 단일 날짜
const pad2 = (n) => String(n).padStart(2, '0');

// 시트 연도(sy)·월(sm) 기준으로 M/D의 연도를 보정 (연말연초 넘어가는 경우)
function resolveYear(m, sy, sm) {
  if (m - sm > 6) return sy - 1; // 1월 시트의 12월 날짜
  if (sm - m > 6) return sy + 1; // 12월 시트의 1월 날짜
  return sy;
}

function parseSectionDates(text, sy, sm) {
  // 괄호 안 날짜 범위: (M/D ... ~ M/D ...)  — 두 번째 M/D가 있어야 '연결'
  const range = text.match(/\((\d{1,2})\/(\d{1,2})[^~]*~\s*(\d{1,2})\/(\d{1,2})/);
  if (range) {
    const sM = +range[1], sD = +range[2], eM = +range[3], eD = +range[4];
    const startY = resolveYear(sM, sy, sm);
    let endY = resolveYear(eM, sy, sm);
    if (eM < sM) endY = startY + 1; // 12/31 ~ 1/1
    return { start: `${startY}-${pad2(sM)}-${pad2(sD)}`, end: `${endY}-${pad2(eM)}-${pad2(eD)}` };
  }
  // 괄호 안 단일 날짜 (예: (6/4(목) 09시~19시)) → 그 날짜 사용 (LIVE보다 우선)
  const single = text.match(/\((\d{1,2})\/(\d{1,2})/);
  if (single) {
    const m = +single[1], d = +single[2], y = resolveYear(m, sy, sm);
    const s = `${y}-${pad2(m)}-${pad2(d)}`;
    return { start: s, end: s };
  }
  // 괄호 날짜가 전혀 없을 때만 LIVE 날짜 (<M/D ...>)
  const live = text.match(/<\s*(\d{1,2})\/(\d{1,2})/);
  if (live) {
    const m = +live[1], d = +live[2], y = resolveYear(m, sy, sm);
    const s = `${y}-${pad2(m)}-${pad2(d)}`;
    return { start: s, end: s };
  }
  return null;
}

const isSectionHeader = (v) =>
  typeof v === 'string' && /<\s*\d{1,2}\/\d{1,2}.*LIVE/i.test(v);

// 시트명이 YYYY.MM(또는 YYYY.M) 이면 {y, mo}, 아니면 null
function parseSheetYM(name) {
  const m = String(name).match(/^(\d{4})\.(\d{1,2})$/);
  return m ? { y: +m[1], mo: +m[2] } : null;
}

// 'YYYY-MM-DD' 에 하루 더하기 (FullCalendar allDay 종료일은 exclusive)
function addOneDay(ymd) {
  const d = new Date(ymd + 'T00:00:00');
  d.setDate(d.getDate() + 1);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function isRfcSheet(rows) {
  return rows.some((r) => r.some((c) => String(c).trim() === 'Task Code'));
}

function parseRfcSheet(rows, sheetName, ym) {
  const hRow = rows.findIndex((r) => r.some((c) => String(c).trim() === 'Task Code'));
  if (hRow < 0) return [];
  const H = rows[hRow].map((c) => String(c).trim());
  const findCol = (needle) =>
    H.findIndex((h) => h.replace(/\s+/g, '').includes(needle.replace(/\s+/g, '')));
  const cTask = findCol('TaskCode');
  // 상태 컬럼: 시트마다 'STATUS' / 'Staus'(오타) 등 표기가 달라, '요청중' 포함 여부로도 감지
  const cStatus = H.findIndex((h) => h.includes('요청중') || /sta.?us/i.test(h));
  const cOwner = H.findIndex((h) => h.includes('기획팀'));
  const cMemo = findCol('메모');

  const sections = [];
  let cur = null;
  for (let i = hRow + 1; i < rows.length; i++) {
    const r = rows[i];
    const c1 = String(r[1] || '').trim();
    if (isSectionHeader(c1)) {
      const dates = parseSectionDates(c1, ym.y, ym.mo);
      cur = dates ? { header: c1, ...dates, tasks: [], sheet: sheetName } : null;
      if (cur) sections.push(cur);
      continue;
    }
    if (!cur) continue;
    const name = c1;
    const code = String(r[cTask] || '').trim();
    if (!name && !code) continue; // 빈/구분 행
    // 한글 정규화(NFC) — OneDrive/맥 경유 셀이 NFD일 수 있어 비교·표시 일관성 확보
    const nfc = (v) => String(v || '').trim().normalize('NFC');
    cur.tasks.push({
      gubun: nfc(r[0]),
      name: nfc(name),
      code: nfc(code),
      status: cStatus >= 0 ? nfc(r[cStatus]) : '',
      owner: cOwner >= 0 ? nfc(r[cOwner]) : '',
      memo: cMemo >= 0 ? nfc(r[cMemo]) : '',
    });
  }

  // 업무(Task Code 또는 STATUS가 있는 실제 배포 건)마다 캘린더 바를 하나씩 생성
  return sections.flatMap((s) => {
    const multiDayEnd = s.start !== s.end ? addOneDay(s.end) : undefined;
    return s.tasks
      .filter((t) => t.code || t.status) // 코드/상태 없는 소제목·연속행 제외
      .map((t) => ({
        title: `🚀 ${t.name.replace(/^\s*\d+\.\s*/, '') || '(제목 없음)'}`,
        start: s.start,
        end: multiDayEnd,
        allDay: true,
        color: '#4c7dff',
        extendedProps: {
          type: 'rfc',
          header: s.header,
          startDisp: s.start,
          endDisp: s.end,
          task: t,
          sheet: s.sheet,
        },
      }));
  });
}

// ── 단순 형식 파서 (날짜/제목/담당자/비고 — 샘플·CSV용) ──────
function parseSimpleSheet(sheetObj, sheetName) {
  const events = [];
  const rows = XLSX.utils.sheet_to_json(sheetObj, { defval: '' });
  for (const row of rows) {
    const date = normalizeDate(pickColumn(row, COLUMN_ALIASES.date));
    const title = String(pickColumn(row, COLUMN_ALIASES.title) || '').trim();
    if (!date || !title) continue;

    const owner = String(pickColumn(row, COLUMN_ALIASES.owner) || '').trim();
    const note = String(pickColumn(row, COLUMN_ALIASES.note) || '').trim();
    events.push({
      title: owner ? `${title} (${owner})` : title,
      start: date,
      allDay: true,
      extendedProps: { type: 'simple', rawTitle: title, owner, note, sheet: sheetName },
    });
  }
  return events;
}

// 최근 N개월 + 미래 시트만 (월별 시트일 때만 적용). 옵션.
const RECENT_MONTHS = 6;
function isSheetInRange(ym) {
  if (!ym) return true; // 날짜형 시트명이 아니면 항상 포함
  const now = new Date();
  const cutoff = now.getFullYear() * 12 + now.getMonth() - RECENT_MONTHS;
  return ym.y * 12 + (ym.mo - 1) >= cutoff;
}

function readExcelFile(filePath) {
  const wb = readWorkbook(filePath);
  let events = [];

  for (const sheetName of wb.SheetNames) {
    const ym = parseSheetYM(sheetName);
    if (!isSheetInRange(ym)) continue; // 최근 6개월+미래만

    const sheetObj = wb.Sheets[sheetName];
    const rows2d = XLSX.utils.sheet_to_json(sheetObj, { defval: '', header: 1 });

    if (isRfcSheet(rows2d) && ym) {
      events = events.concat(parseRfcSheet(rows2d, sheetName, ym));
    } else {
      events = events.concat(parseSimpleSheet(sheetObj, sheetName));
    }
  }
  return events;
}

// 월별 시트가 월말/월초 경계에서 겹쳐 같은 배포가 두 시트에 들어 있음
// → Task Code(고유) 기준으로 중복 제거. 코드가 없으면 날짜+제목으로.
function dedupEvents(events) {
  const seen = new Set();
  const out = [];
  for (const e of events) {
    const code = e.extendedProps && e.extendedProps.task && e.extendedProps.task.code;
    const key = code ? `c:${code}` : `t:${e.start}|${e.end || ''}|${e.title}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(e);
  }
  return out;
}

// 폴더 안 모든 엑셀(.xlsx/.xls)을 읽어 합침
function readAllExcels(dir) {
  if (!dir || !fs.existsSync(dir)) return { events: [], error: '공유 폴더가 설정되지 않았거나 존재하지 않습니다.' };

  let events = [];
  const errors = [];
  const files = fs
    .readdirSync(dir)
    .filter((f) => /\.(xlsx|xls|csv)$/i.test(f) && !f.startsWith('~$')); // 임시잠금파일 제외

  for (const file of files) {
    try {
      events = events.concat(readExcelFile(path.join(dir, file)));
    } catch (e) {
      errors.push(`${file}: ${e.message}`);
    }
  }
  events = dedupEvents(events);
  return { events, fileCount: files.length, error: errors.length ? errors.join('\n') : null };
}

function pushEvents() {
  if (!mainWindow) return;
  const config = loadConfig();
  const dir = activeDir(config);
  const result = readAllExcels(dir);
  mainWindow.webContents.send('events:update', { ...result, watchDir: dir, active: config.active });
}

// ── 공유 폴더 감시 ─────────────────────────────────────────
function startWatching(dir) {
  if (watcher) {
    watcher.close();
    watcher = null;
  }
  if (!dir || !fs.existsSync(dir)) return;

  watcher = chokidar.watch(dir, {
    depth: 0,
    ignoreInitial: true,
    awaitWriteFinish: { stabilityThreshold: 800, pollInterval: 200 },
  });

  const onChange = (p) => {
    if (/\.(xlsx|xls|csv)$/i.test(p) && !path.basename(p).startsWith('~$')) {
      pushEvents();
    }
  };
  watcher.on('add', onChange).on('change', onChange).on('unlink', onChange);
}

// ── IPC ────────────────────────────────────────────────────
ipcMain.handle('config:get', () => loadConfig());

// 현재 활성 팀의 공유 폴더 지정
ipcMain.handle('config:pickFolder', async () => {
  const config = loadConfig();
  const res = await dialog.showOpenDialog(mainWindow, {
    title: `[${config.active}] 배포일정 엑셀이 있는 공유 폴더 선택`,
    properties: ['openDirectory'],
  });
  if (res.canceled || !res.filePaths[0]) return config;

  config.teams[config.active].watchDir = res.filePaths[0];
  saveConfig(config);
  startWatching(activeDir(config));
  pushEvents();
  return config;
});

// 팀 전환 (GMK / GMAP)
ipcMain.handle('team:set', (_e, team) => {
  const config = loadConfig();
  if (!TEAMS.includes(team)) return config;
  config.active = team;
  saveConfig(config);
  startWatching(activeDir(config));
  pushEvents();
  return config;
});

ipcMain.handle('events:reload', () => {
  pushEvents();
  return true;
});

app.whenReady().then(() => {
  createWindow();
  const config = loadConfig();
  startWatching(activeDir(config));

  mainWindow.webContents.once('did-finish-load', () => pushEvents());

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (watcher) watcher.close();
  if (process.platform !== 'darwin') app.quit();
});
