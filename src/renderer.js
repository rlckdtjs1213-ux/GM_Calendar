'use strict';

let calendar;

function initCalendar() {
  const el = document.getElementById('calendar');
  calendar = new FullCalendar.Calendar(el, {
    initialView: 'dayGridMonth',
    height: '100%',
    locale: 'ko',
    firstDay: 0,
    headerToolbar: {
      left: 'prev,next today',
      center: 'title',
      right: 'dayGridMonth,timeGridWeek,listMonth',
    },
    buttonText: { today: '오늘', month: '월', week: '주', list: '목록' },
    titleFormat: { year: 'numeric', month: 'long' },
    dayMaxEvents: 4,
    events: [],
    eventClick(info) {
      showDetail(info.event);
    },
  });
  calendar.render();
}

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}

function showDetail(event) {
  const p = event.extendedProps;
  const titleEl = document.getElementById('m-title');
  const metaEl = document.getElementById('m-meta');
  const body = document.getElementById('m-body');
  body.replaceChildren();
  metaEl.replaceChildren();

  if (p.type === 'rfc') {
    const t = p.task;
    const period = p.startDisp === p.endDisp ? p.startDisp : `${p.startDisp} ~ ${p.endDisp}`;
    titleEl.textContent = t.name.replace(/^\s*\d+\.\s*/, '');
    metaEl.append(el('div', 'm-period', period));
    metaEl.append(el('div', 'm-header', p.header));

    const dl = el('dl', 'm-dl');
    const add = (k, v) => {
      if (!v) return;
      dl.append(el('dt', null, k));
      dl.append(el('dd', null, v));
    };
    add('구분', t.gubun);
    add('Task Code', t.code);
    add('담당자', t.owner);
    add('상태', t.status);
    add('메모', t.memo);
    body.append(dl);
    body.append(el('div', 'm-sheet', `출처: ${p.sheet}`));
  } else {
    titleEl.textContent = p.rawTitle || event.title;
    metaEl.append(el('div', 'm-period', event.startStr));
    const dl = el('dl', 'm-dl');
    const add = (k, v) => {
      dl.append(el('dt', null, k));
      dl.append(el('dd', null, v || '-'));
    };
    add('담당자', p.owner);
    add('비고', p.note);
    add('출처', p.sheet);
    body.append(dl);
  }

  document.getElementById('overlay').classList.remove('hidden');
}

function setStatus(text, isError) {
  const el = document.getElementById('status');
  el.textContent = text;
  el.classList.toggle('status-error', !!isError);
}

function applyTheme(theme) {
  const dark = theme === 'dark';
  document.body.classList.toggle('dark', dark);
  const btn = document.getElementById('themeBtn');
  if (btn) {
    btn.textContent = dark ? '☀️' : '🌙';
    btn.title = dark ? '라이트모드 전환' : '다크모드 전환';
  }
}

function toggleTheme() {
  const next = document.body.classList.contains('dark') ? 'light' : 'dark';
  localStorage.setItem('theme', next);
  applyTheme(next);
}

function setActiveTeam(team) {
  document.querySelectorAll('.team-btn').forEach((b) => {
    b.classList.toggle('active', b.dataset.team === team);
  });
}

function applyEvents(payload) {
  if (!calendar) return;
  if (payload.active) setActiveTeam(payload.active);
  calendar.removeAllEvents();
  if (payload.events && payload.events.length) {
    calendar.addEventSource(payload.events);
  }

  const prefix = payload.active ? `[${payload.active}] ` : '';
  if (payload.error && (!payload.events || !payload.events.length)) {
    setStatus(prefix + payload.error, true);
  } else if (!payload.watchDir) {
    setStatus(prefix + '공유 폴더를 설정하세요', true);
  } else {
    const now = new Date().toLocaleTimeString('ko-KR');
    setStatus(
      `${prefix}엑셀 ${payload.fileCount || 0}개 · 일정 ${payload.events.length}건 · ${now} 갱신`,
      false
    );
  }
}

window.addEventListener('DOMContentLoaded', async () => {
  applyTheme(localStorage.getItem('theme') || 'light');
  document.getElementById('themeBtn').addEventListener('click', toggleTheme);

  initCalendar();

  window.api.onEventsUpdate(applyEvents);

  document.getElementById('reloadBtn').addEventListener('click', () => window.api.reload());
  document.getElementById('folderBtn').addEventListener('click', () => window.api.pickFolder());
  document.querySelectorAll('.team-btn').forEach((b) =>
    b.addEventListener('click', () => window.api.setTeam(b.dataset.team))
  );

  const updateBtn = document.getElementById('updateBtn');
  updateBtn.addEventListener('click', () => {
    setStatus('업데이트 확인 중…', false);
    window.api.checkUpdate();
  });
  window.api.onUpdateStatus(({ state, info }) => {
    const v = info && info.version ? ` (v${info.version})` : '';
    switch (state) {
      case 'checking': setStatus('업데이트 확인 중…', false); break;
      case 'available': setStatus(`새 버전${v} 다운로드 중…`, false); break;
      case 'downloading': setStatus(`업데이트 다운로드 ${info.percent}%`, false); break;
      case 'downloaded': setStatus(`새 버전${v} 다운로드 완료 — 재시작하면 설치됩니다`, false); break;
      case 'latest': setStatus(`최신 버전입니다${v}`, false); break;
      case 'dev': setStatus('개발 모드에서는 업데이트 확인 불가 (설치본에서만 동작)', true); break;
      case 'error': setStatus(`업데이트 확인 실패: ${info && info.message ? info.message : ''}`, true); break;
    }
  });
  document.getElementById('m-close').addEventListener('click', () =>
    document.getElementById('overlay').classList.add('hidden')
  );
  document.getElementById('overlay').addEventListener('click', (e) => {
    if (e.target.id === 'overlay') e.currentTarget.classList.add('hidden');
  });

  const config = await window.api.getConfig();
  setActiveTeam(config.active);
  if (!config.teams[config.active].watchDir) setStatus(`[${config.active}] 공유 폴더를 설정하세요`, true);
});
