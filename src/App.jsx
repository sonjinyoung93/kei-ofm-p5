import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Plus, Trash2, Download, ClipboardList, LayoutDashboard, Loader2, AlertCircle, CheckCircle2, X, Settings, Save, LogOut, Building2 } from 'lucide-react';

const ROWS = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i)); // A~Z
const COLS = Array.from({ length: 70 }, (_, i) => String(i + 1)); // 1~70
const PROCESSES = ['자탐', '유도등', '무통'];
const UNITS = ['EA', 'M', 'SET', 'BOX', 'ROLL', 'KG'];
const STATUS_FLOW = ['요청됨', '확인됨', '발주완료', '입고완료'];
const STATUS_COLOR = {
  '요청됨':   { bg: '#EEF0EC', fg: '#5C6B73', bd: '#C9CFC6' },
  '확인됨':   { bg: '#E3ECF5', fg: '#2B5A8C', bd: '#9FBEDC' },
  '발주완료': { bg: '#FBE7DA', fg: '#B84B10', bd: '#F0A97A' },
  '입고완료': { bg: '#E1EBE3', fg: '#2E6B47', bd: '#9FC7AC' },
};

// 품목 카탈로그
const CATALOG = [
  { name: '무나사전선관', spec: 'E19', color: 'N/A' },
  { name: '무나사전선관', spec: 'E25', color: 'N/A' },
  { name: '무나사전선관', spec: 'E31', color: 'N/A' },
  { name: '무나사전선관', spec: 'E39', color: 'N/A' },
  { name: '커플링', spec: 'E19', color: 'N/A' },
  { name: '커플링', spec: 'E25', color: 'N/A' },
  { name: '커플링', spec: 'E31', color: 'N/A' },
  { name: '커플링', spec: 'E39', color: 'N/A' },
  { name: '박스커넥터', spec: 'E19', color: 'N/A' },
  { name: '박스커넥터', spec: 'E25', color: 'N/A' },
  { name: '박스커넥터', spec: 'E31', color: 'N/A' },
  { name: '박스커넥터', spec: 'E39', color: 'N/A' },
  { name: '통신케이블', spec: '14TP', color: '적' },
  { name: '통신케이블', spec: '14TP', color: '흑' },
  { name: '통신케이블', spec: '16TSP', color: '황' },
  { name: '내화케이블', spec: '4mm 3C', color: 'N/A' },
  { name: '내화케이블', spec: '4mm 2C', color: 'N/A' },
  { name: '내화케이블', spec: '2.5mm 2C', color: 'N/A' },
];
const CATALOG_NAMES = [...new Set(CATALOG.map(c => c.name))];
function getSpecs(name) { return [...new Set(CATALOG.filter(c => c.name === name).map(c => c.spec))]; }
function getColors(name, spec) { return [...new Set(CATALOG.filter(c => c.name === name && c.spec === spec).map(c => c.color))]; }

const DEFAULT_PROJECTS = [
  { id: 'proj-1', name: 'P4 Ph4 (삼성물산)' },
  { id: 'proj-2', name: '신규 프로젝트' },
];

// ⚠️ 구글 Apps Script를 "웹 앱으로 배포"한 뒤 나오는 URL로 반드시 교체하세요.
// 작업매뉴얼.md 2~3단계 참고.
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwVPkUPP93WpzWhZLZ4F4fOkQkC5HQ-h1UM7cK72MzdCJ4oxfT5ALwA6GsDj-LSJbpW/exec?action=projects';

async function apiGet(action) {
  const res = await fetch(`${APPS_SCRIPT_URL}?action=${action}`);
  if (!res.ok) throw new Error(`GET ${action} failed: ${res.status}`);
  return res.json();
}

async function apiPost(action, data) {
  // Content-Type을 text/plain으로 보내면 브라우저의 CORS preflight(OPTIONS)를
  // 건너뛸 수 있습니다. Apps Script 쪽에서는 그대로 JSON.parse 해서 읽습니다.
  const res = await fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, ...data }),
  });
  if (!res.ok) throw new Error(`POST ${action} failed: ${res.status}`);
  return res.json();
}

function pad(n) { return String(n).padStart(2, '0'); }
function genReqNo() {
  const d = new Date();
  const stamp = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
  const rand = Math.random().toString(36).slice(2, 6).toUpperCase();
  return `MRS-${stamp}-${rand}`;
}
function genId() { return Math.random().toString(36).slice(2, 10); }
function fmtDate(iso) {
  const d = new Date(iso);
  return `${d.getFullYear()}.${pad(d.getMonth() + 1)}.${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function fmtZone(rowFrom, rowTo, colFrom, colTo) {
  const rowPart = rowFrom === rowTo ? `${rowFrom}행` : `${rowFrom}~${rowTo}행`;
  const colPart = colFrom === colTo ? `${colFrom}열` : `${colFrom}~${colTo}열`;
  return `${rowPart} ${colPart}`;
}

function newItemRow() {
  const name = CATALOG_NAMES[0];
  const spec = getSpecs(name)[0];
  const color = getColors(name, spec)[0];
  return { id: genId(), name, spec, color, qty: '', unit: 'EA' };
}

const GlobalStyle = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');
    .mrs-root {
      --paper: #F0EEE6;
      --paper-line: #D8D4C6;
      --ink: #1C2A33;
      --ink-soft: #5C6B73;
      --steel: #8A94A6;
      --accent: #D9601B;
      --accent-dark: #B84B10;
      --line: #C9CFC6;
      --card: #FBFAF6;
      font-family: 'Apple SD Gothic Neo','Malgun Gothic',-apple-system,sans-serif;
      color: var(--ink);
      background:
        repeating-linear-gradient(0deg, transparent, transparent 27px, rgba(28,42,51,0.035) 28px),
        repeating-linear-gradient(90deg, transparent, transparent 27px, rgba(28,42,51,0.035) 28px),
        var(--paper);
      min-height: 100%;
      padding: 0;
    }
    .mrs-mono { font-family: 'JetBrains Mono', monospace; }
    .mrs-display { font-family: 'Oswald', 'Apple SD Gothic Neo', sans-serif; letter-spacing: 0.02em; }

    .mrs-header {
      background: var(--ink);
      color: #EDEAE0;
      padding: 18px 20px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      flex-wrap: wrap;
      border-bottom: 3px solid var(--accent);
    }
    .mrs-project-chip {
      display: flex; align-items: center; gap: 8px;
      border: 1px solid #3A4954;
      border-radius: 20px;
      padding: 6px 14px;
      font-size: 12px;
      color: #C9CFC6;
    }
    .mrs-project-chip b { color: #EDEAE0; font-weight: 600; font-size: 13px; }
    .mrs-header-right { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
    .mrs-user-chip { font-size: 12px; color: #C9CFC6; display: flex; align-items: center; gap: 6px; }
    .mrs-logout-btn {
      display: flex; align-items: center; gap: 6px;
      background: none; border: 1px solid #3A4954; color: #C9CFC6;
      border-radius: 20px; padding: 6px 12px; font-size: 12px; cursor: pointer;
    }
    .mrs-logout-btn:hover { border-color: var(--accent); color: #EDEAE0; }

    .mrs-tabs {
      display: flex;
      gap: 0;
      background: var(--card);
      border-bottom: 1px solid var(--line);
      padding: 0 16px;
    }
    .mrs-tab {
      display: flex; align-items: center; gap: 7px;
      padding: 12px 18px;
      font-size: 14px;
      font-weight: 600;
      color: var(--ink-soft);
      cursor: pointer;
      border-bottom: 3px solid transparent;
      transition: all .15s ease;
      background: none; border-top:none; border-left:none; border-right:none;
    }
    .mrs-tab.active { color: var(--ink); border-bottom-color: var(--accent); }
    .mrs-tab:hover:not(.active) { color: var(--ink); background: rgba(217,96,27,0.05); }

    .mrs-body { padding: 20px; max-width: 1150px; margin: 0 auto; }

    .mrs-card {
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 3px;
      box-shadow: 0 1px 2px rgba(28,42,51,0.06);
    }

    .mrs-field-label {
      font-size: 12px; font-weight: 600; color: var(--ink-soft);
      margin-bottom: 6px; display: block;
      text-transform: uppercase; letter-spacing: 0.04em;
    }
    .mrs-input, .mrs-select, .mrs-textarea {
      width: 100%; box-sizing: border-box;
      padding: 9px 10px;
      border: 1px solid var(--line);
      border-radius: 2px;
      background: #fff;
      font-size: 14px;
      color: var(--ink);
      font-family: inherit;
    }
    .mrs-input:focus, .mrs-select:focus, .mrs-textarea:focus {
      outline: none; border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(217,96,27,0.12);
    }
    .mrs-input:disabled, .mrs-select:disabled { background: #F1EFE9; color: var(--ink-soft); }

    .mrs-zone-grid {
      display: grid;
      grid-template-columns: 1fr 1fr 1fr 1fr;
      gap: 8px;
    }
    @media (max-width: 640px) { .mrs-zone-grid { grid-template-columns: 1fr 1fr; } }

    .mrs-item-row {
      display: grid;
      grid-template-columns: 1.6fr 1fr 0.8fr 0.7fr 0.7fr auto;
      gap: 8px;
      align-items: end;
      padding: 10px 0;
      border-bottom: 1px dashed var(--line);
    }
    @media (max-width: 820px) {
      .mrs-item-row { grid-template-columns: 1fr 1fr; }
    }

    .mrs-btn {
      display: inline-flex; align-items: center; gap: 6px;
      padding: 10px 16px;
      border-radius: 2px;
      font-size: 14px; font-weight: 600;
      cursor: pointer; border: 1px solid transparent;
      transition: all .15s ease;
    }
    .mrs-btn-primary { background: var(--accent); color: #fff; }
    .mrs-btn-primary:hover { background: var(--accent-dark); }
    .mrs-btn-ghost { background: transparent; color: var(--ink-soft); border-color: var(--line); }
    .mrs-btn-ghost:hover { border-color: var(--ink-soft); color: var(--ink); }
    .mrs-btn-danger { background: transparent; color: #B84B10; }
    .mrs-btn:disabled { opacity: 0.5; cursor: not-allowed; }

    .mrs-chip {
      display: inline-flex; align-items: center;
      padding: 3px 9px; border-radius: 20px;
      font-size: 12px; font-weight: 600;
      border: 1px solid;
      white-space: nowrap;
    }

    .mrs-table-wrap { overflow-x: auto; border: 1px solid var(--line); border-radius: 3px; background: var(--card); }
    table.mrs-table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 950px; }
    table.mrs-table th {
      text-align: left; padding: 10px 12px;
      background: var(--ink); color: #C9CFC6;
      font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em;
      position: sticky; top: 0;
    }
    table.mrs-table td { padding: 10px 12px; border-bottom: 1px solid var(--paper-line); vertical-align: middle; }
    table.mrs-table tr:hover td { background: rgba(217,96,27,0.04); }

    .mrs-stat {
      display: flex; flex-direction: column; gap: 2px;
      padding: 10px 16px; border-right: 1px solid var(--line);
    }
    .mrs-stat:last-child { border-right: none; }
    .mrs-stat .n { font-size: 22px; font-weight: 700; }
    .mrs-stat .l { font-size: 11px; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.04em; }

    select.mrs-status-select {
      font-size: 12px; font-weight: 600; border-radius: 20px;
      padding: 4px 24px 4px 10px; border: 1px solid; cursor: pointer;
      -webkit-appearance: none; appearance: none;
      background-repeat: no-repeat; background-position: right 8px center; background-size: 10px;
    }

    .mrs-empty { text-align: center; padding: 40px 20px; color: var(--ink-soft); }
    .mrs-spin { animation: mrs-spin 1s linear infinite; }
    @keyframes mrs-spin { to { transform: rotate(360deg); } }

    .mrs-login-wrap {
      min-height: 100vh;
      display: flex; align-items: center; justify-content: center;
      padding: 20px;
    }
    .mrs-login-card {
      width: 100%; max-width: 380px;
      background: var(--card);
      border: 1px solid var(--line);
      border-top: 4px solid var(--accent);
      border-radius: 3px;
      padding: 32px 28px;
      box-shadow: 0 4px 16px rgba(28,42,51,0.08);
    }
  `}</style>
);

function StatusSelect({ value, onChange }) {
  const c = STATUS_COLOR[value];
  return (
    <select
      className="mrs-status-select"
      value={value}
      onChange={e => onChange(e.target.value)}
      style={{ background: c.bg, color: c.fg, borderColor: c.bd }}
    >
      {STATUS_FLOW.map(s => <option key={s} value={s}>{s}</option>)}
    </select>
  );
}

function ItemRowEditor({ item, onChange, onRemove, removable }) {
  const specs = getSpecs(item.name);
  const colors = getColors(item.name, item.spec);

  function handleNameChange(name) {
    const spec = getSpecs(name)[0];
    const color = getColors(name, spec)[0];
    onChange({ ...item, name, spec, color });
  }
  function handleSpecChange(spec) {
    const color = getColors(item.name, spec)[0];
    onChange({ ...item, spec, color });
  }

  return (
    <div className="mrs-item-row">
      <div>
        <select className="mrs-select" value={item.name} onChange={e => handleNameChange(e.target.value)}>
          {CATALOG_NAMES.map(n => <option key={n} value={n}>{n}</option>)}
        </select>
      </div>
      <div>
        <select className="mrs-select" value={item.spec} onChange={e => handleSpecChange(e.target.value)}>
          {specs.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
      </div>
      <div>
        <select className="mrs-select" value={item.color} onChange={e => onChange({ ...item, color: e.target.value })} disabled={colors.length <= 1}>
          {colors.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
      </div>
      <div>
        <input className="mrs-input" type="number" min="0" value={item.qty} onChange={e => onChange({ ...item, qty: e.target.value })} placeholder="0" />
      </div>
      <div>
        <select className="mrs-select" value={item.unit} onChange={e => onChange({ ...item, unit: e.target.value })}>
          {UNITS.map(u => <option key={u} value={u}>{u}</option>)}
        </select>
      </div>
      <button className="mrs-btn mrs-btn-danger" onClick={onRemove} disabled={!removable} title="삭제" style={{ padding: 8 }}>
        <Trash2 size={16} />
      </button>
    </div>
  );
}

function LoginScreen({ projects, onLogin, loading }) {
  const [name, setName] = useState('');
  const [projectId, setProjectId] = useState(projects[0]?.id || '');

  useEffect(() => {
    if (!projectId && projects.length) setProjectId(projects[0].id);
  }, [projects]);

  function handleStart() {
    if (!name.trim()) { alert('팀장명을 입력해주세요.'); return; }
    if (!projectId) { alert('프로젝트를 선택해주세요.'); return; }
    onLogin({ requester: name.trim(), projectId });
  }

  return (
    <div className="mrs-root">
      <GlobalStyle />
      <div className="mrs-login-wrap">
        <div className="mrs-login-card">
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
            <Building2 size={20} color="var(--accent)" />
            <span className="mrs-display" style={{ fontSize: 18, fontWeight: 600 }}>자재 요청 관리 시스템</span>
          </div>
          <div style={{ fontSize: 12, color: 'var(--ink-soft)', marginBottom: 24 }}>팀장 로그인 · 건국이엔아이 공무팀</div>

          {loading ? (
            <div className="mrs-empty"><Loader2 size={18} className="mrs-spin" /></div>
          ) : (
            <>
              <label className="mrs-field-label">팀장명</label>
              <input className="mrs-input" value={name} onChange={e => setName(e.target.value)} placeholder="예: 손진영 차장" style={{ marginBottom: 16 }} />

              <label className="mrs-field-label">담당 프로젝트</label>
              <select className="mrs-select" value={projectId} onChange={e => setProjectId(e.target.value)} style={{ marginBottom: 22 }}>
                {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>

              <button className="mrs-btn mrs-btn-primary" style={{ width: '100%', justifyContent: 'center' }} onClick={handleStart}>
                시작하기
              </button>
              <div style={{ fontSize: 11, color: 'var(--ink-soft)', marginTop: 14, lineHeight: 1.5 }}>
                선택한 프로젝트로 요청이 자동 분류되어 취합됩니다. 담당 프로젝트가 바뀌면 우측 상단 로그아웃 후 다시 선택해주세요.
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function RequestForm({ session, projectName, onSubmit, saving }) {
  const [rowFrom, setRowFrom] = useState('A');
  const [rowTo, setRowTo] = useState('A');
  const [colFrom, setColFrom] = useState('1');
  const [colTo, setColTo] = useState('1');
  const [process, setProcess] = useState(PROCESSES[0]);
  const [note, setNote] = useState('');
  const [items, setItems] = useState([newItemRow()]);
  const [justSubmitted, setJustSubmitted] = useState(false);

  function updateItem(id, updated) {
    setItems(items.map(it => it.id === id ? updated : it));
  }
  function addItem() { setItems([...items, newItemRow()]); }
  function removeItem(id) {
    if (items.length === 1) return;
    setItems(items.filter(it => it.id !== id));
  }

  function validate() {
    const valid = items.filter(it => it.name && it.qty !== '' && Number(it.qty) > 0);
    if (valid.length === 0) return '최소 1개 이상의 품목에 수량을 입력해주세요.';
    return null;
  }

  async function handleSubmit() {
    const err = validate();
    if (err) { alert(err); return; }
    const payload = {
      id: genId(),
      reqNo: genReqNo(),
      requester: session.requester,
      projectId: session.projectId,
      zone: fmtZone(rowFrom, rowTo, colFrom, colTo),
      process,
      note: note.trim(),
      createdAt: new Date().toISOString(),
      items: items
        .filter(it => it.name && it.qty !== '' && Number(it.qty) > 0)
        .map(it => ({ id: genId(), name: it.name, spec: it.spec, color: it.color, qty: it.qty, unit: it.unit, status: '요청됨' })),
    };
    await onSubmit(payload);
    setRowFrom('A'); setRowTo('A'); setColFrom('1'); setColTo('1');
    setProcess(PROCESSES[0]); setNote(''); setItems([newItemRow()]);
    setJustSubmitted(true);
    setTimeout(() => setJustSubmitted(false), 3000);
  }

  return (
    <div className="mrs-card" style={{ padding: 22 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 18 }}>
        <h2 className="mrs-display" style={{ fontSize: 18, margin: 0, fontWeight: 600 }}>자재 요청 입력</h2>
        <span className="mrs-mono" style={{ fontSize: 11, color: 'var(--ink-soft)' }}>NEW REQUEST</span>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14, marginBottom: 14 }}>
        <div>
          <label className="mrs-field-label">요청자</label>
          <input className="mrs-input" value={session.requester} disabled />
        </div>
        <div>
          <label className="mrs-field-label">프로젝트</label>
          <input className="mrs-input" value={projectName} disabled />
        </div>
      </div>

      <div style={{ marginBottom: 14 }}>
        <label className="mrs-field-label">공정</label>
        <select className="mrs-select" value={process} onChange={e => setProcess(e.target.value)}>
          {PROCESSES.map(p => <option key={p} value={p}>{p}</option>)}
        </select>
      </div>

      <div style={{ marginBottom: 14 }}>
        <label className="mrs-field-label">구역 (행 / 열)</label>
        <div className="mrs-zone-grid">
          <div>
            <span style={{ fontSize: 11, color: 'var(--ink-soft)' }}>행 시작</span>
            <select className="mrs-select" value={rowFrom} onChange={e => setRowFrom(e.target.value)}>
              {ROWS.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
          <div>
            <span style={{ fontSize: 11, color: 'var(--ink-soft)' }}>행 끝</span>
            <select className="mrs-select" value={rowTo} onChange={e => setRowTo(e.target.value)}>
              {ROWS.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
          <div>
            <span style={{ fontSize: 11, color: 'var(--ink-soft)' }}>열 시작</span>
            <select className="mrs-select" value={colFrom} onChange={e => setColFrom(e.target.value)}>
              {COLS.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <span style={{ fontSize: 11, color: 'var(--ink-soft)' }}>열 끝</span>
            <select className="mrs-select" value={colTo} onChange={e => setColTo(e.target.value)}>
              {COLS.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
        </div>
      </div>

      <div style={{ marginBottom: 8 }}>
        <label className="mrs-field-label">요청 품목</label>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1.6fr 1fr 0.8fr 0.7fr 0.7fr auto', gap: 8, marginBottom: 4 }}>
        <span style={{ fontSize: 11, color: 'var(--ink-soft)' }}>품목</span>
        <span style={{ fontSize: 11, color: 'var(--ink-soft)' }}>규격</span>
        <span style={{ fontSize: 11, color: 'var(--ink-soft)' }}>색상</span>
        <span style={{ fontSize: 11, color: 'var(--ink-soft)' }}>수량</span>
        <span style={{ fontSize: 11, color: 'var(--ink-soft)' }}>단위</span>
        <span></span>
      </div>
      {items.map(it => (
        <ItemRowEditor
          key={it.id}
          item={it}
          onChange={updated => updateItem(it.id, updated)}
          onRemove={() => removeItem(it.id)}
          removable={items.length > 1}
        />
      ))}
      <button className="mrs-btn mrs-btn-ghost" style={{ marginTop: 10 }} onClick={addItem}>
        <Plus size={15} /> 품목 추가
      </button>

      <div style={{ marginTop: 18 }}>
        <label className="mrs-field-label">특이사항 (선택)</label>
        <input className="mrs-input" value={note} onChange={e => setNote(e.target.value)} placeholder="예: 3층 반입 엘리베이터 이용 불가, 인력으로 운반 필요" />
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 22 }}>
        <button className="mrs-btn mrs-btn-primary" onClick={handleSubmit} disabled={saving}>
          {saving ? <Loader2 size={15} className="mrs-spin" /> : <ClipboardList size={15} />}
          요청 제출
        </button>
        {justSubmitted && (
          <span style={{ display: 'flex', alignItems: 'center', gap: 6, color: '#2E6B47', fontSize: 13, fontWeight: 600 }}>
            <CheckCircle2 size={16} /> 요청이 접수되었습니다
          </span>
        )}
      </div>
    </div>
  );
}

function ProjectManager({ projects, onSave, saving }) {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState(projects);
  useEffect(() => setDraft(projects), [projects]);

  if (!open) {
    return (
      <button className="mrs-btn mrs-btn-ghost" onClick={() => setOpen(true)} style={{ marginBottom: 14 }}>
        <Settings size={15} /> 프로젝트명 관리
      </button>
    );
  }
  return (
    <div className="mrs-card" style={{ padding: 14, marginBottom: 14 }}>
      <label className="mrs-field-label">프로젝트 목록</label>
      {draft.map((p, idx) => (
        <div key={p.id} style={{ display: 'flex', gap: 10, marginBottom: 8, alignItems: 'center' }}>
          <input
            className="mrs-input"
            value={p.name}
            onChange={e => {
              const next = [...draft];
              next[idx] = { ...p, name: e.target.value };
              setDraft(next);
            }}
          />
        </div>
      ))}
      <div style={{ display: 'flex', gap: 10, marginTop: 10 }}>
        <button className="mrs-btn mrs-btn-primary" disabled={saving || draft.some(p => !p.name.trim())} onClick={() => { onSave(draft); setOpen(false); }}>
          <Save size={15} /> 저장
        </button>
        <button className="mrs-btn mrs-btn-ghost" onClick={() => { setDraft(projects); setOpen(false); }}>취소</button>
      </div>
    </div>
  );
}

function AdminDashboard({ requests, projects, onUpdateStatus, onDelete, onSaveProjects, savingSettings }) {
  const [projectFilter, setProjectFilter] = useState('전체');
  const [zoneQuery, setZoneQuery] = useState('');
  const [processFilter, setProcessFilter] = useState('전체');
  const [statusFilter, setStatusFilter] = useState('전체');

  const projectNameById = {};
  projects.forEach(p => { projectNameById[p.id] = p.name; });

  const rows = [];
  requests.forEach(r => {
    r.items.forEach(it => {
      rows.push({
        reqId: r.id, itemId: it.id, reqNo: r.reqNo, requester: r.requester,
        projectId: r.projectId, projectName: projectNameById[r.projectId] || '(삭제된 프로젝트)',
        zone: r.zone, process: r.process, name: it.name, spec: it.spec, color: it.color,
        qty: it.qty, unit: it.unit, status: it.status,
        note: r.note, createdAt: r.createdAt,
      });
    });
  });

  const statuses = ['전체', ...STATUS_FLOW];

  const filtered = rows.filter(r =>
    (projectFilter === '전체' || r.projectId === projectFilter) &&
    (zoneQuery.trim() === '' || r.zone.includes(zoneQuery.trim())) &&
    (processFilter === '전체' || r.process === processFilter) &&
    (statusFilter === '전체' || r.status === statusFilter)
  ).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  const counts = STATUS_FLOW.reduce((acc, s) => { acc[s] = rows.filter(r => (projectFilter === '전체' || r.projectId === projectFilter) && r.status === s).length; return acc; }, {});
  const totalCount = rows.filter(r => projectFilter === '전체' || r.projectId === projectFilter).length;

  function exportExcel() {
    const data = filtered.map(r => ({
      '프로젝트': r.projectName,
      '요청번호': r.reqNo,
      '요청일시': fmtDate(r.createdAt),
      '요청자': r.requester,
      '구역': r.zone,
      '공정': r.process,
      '품목명': r.name,
      '규격': r.spec,
      '색상': r.color,
      '수량': r.qty,
      '단위': r.unit,
      '상태': r.status,
      '특이사항': r.note || '',
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{wch:16},{wch:18},{wch:16},{wch:12},{wch:14},{wch:8},{wch:16},{wch:12},{wch:8},{wch:8},{wch:8},{wch:10},{wch:30}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '자재요청');
    const d = new Date();
    const label = projectFilter === '전체' ? '전체프로젝트' : (projectNameById[projectFilter] || '프로젝트');
    XLSX.writeFile(wb, `${label}_자재요청_취합_${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}.xlsx`);
  }

  return (
    <div>
      <ProjectManager projects={projects} onSave={onSaveProjects} saving={savingSettings} />

      <div style={{ display: 'flex', gap: 10, marginBottom: 14, flexWrap: 'wrap' }}>
        <select className="mrs-select" style={{ width: 'auto', fontWeight: 600 }} value={projectFilter} onChange={e => setProjectFilter(e.target.value)}>
          <option value="전체">전체 프로젝트</option>
          {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
      </div>

      <div className="mrs-card" style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 16 }}>
        <div className="mrs-stat"><span className="n mrs-mono">{totalCount}</span><span className="l">전체 요청</span></div>
        {STATUS_FLOW.map(s => (
          <div className="mrs-stat" key={s}>
            <span className="n mrs-mono" style={{ color: STATUS_COLOR[s].fg }}>{counts[s]}</span>
            <span className="l">{s}</span>
          </div>
        ))}
      </div>

      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 14, alignItems: 'center' }}>
        <input className="mrs-input" style={{ width: 160 }} value={zoneQuery} onChange={e => setZoneQuery(e.target.value)} placeholder="구역 검색 (예: A행)" />
        <select className="mrs-select" style={{ width: 'auto' }} value={processFilter} onChange={e => setProcessFilter(e.target.value)}>
          <option value="전체">전체 공정</option>
          {PROCESSES.map(p => <option key={p} value={p}>{p}</option>)}
        </select>
        <select className="mrs-select" style={{ width: 'auto' }} value={statusFilter} onChange={e => setStatusFilter(e.target.value)}>
          {statuses.map(s => <option key={s} value={s}>{s === '전체' ? '전체 상태' : s}</option>)}
        </select>
        <div style={{ flex: 1 }} />
        <button className="mrs-btn mrs-btn-primary" onClick={exportExcel} disabled={filtered.length === 0}>
          <Download size={15} /> 엑셀로 내보내기
        </button>
      </div>

      {filtered.length === 0 ? (
        <div className="mrs-card mrs-empty">표시할 요청이 없습니다.</div>
      ) : (
        <div className="mrs-table-wrap">
          <table className="mrs-table">
            <thead>
              <tr>
                <th>프로젝트</th><th>요청번호</th><th>요청일시</th><th>요청자</th><th>구역</th><th>공정</th>
                <th>품목명</th><th>규격</th><th>색상</th><th>수량</th><th>상태</th><th></th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(r => (
                <tr key={r.itemId}>
                  <td style={{ fontWeight: 600 }}>{r.projectName}</td>
                  <td className="mrs-mono" style={{ fontSize: 12, color: 'var(--ink-soft)' }}>{r.reqNo}</td>
                  <td className="mrs-mono" style={{ fontSize: 12 }}>{fmtDate(r.createdAt)}</td>
                  <td>{r.requester}</td>
                  <td>{r.zone}</td>
                  <td>{r.process}</td>
                  <td style={{ fontWeight: 600 }}>{r.name}</td>
                  <td style={{ color: 'var(--ink-soft)' }}>{r.spec || '-'}</td>
                  <td style={{ color: 'var(--ink-soft)' }}>{r.color || '-'}</td>
                  <td className="mrs-mono">{r.qty} {r.unit}</td>
                  <td><StatusSelect value={r.status} onChange={v => onUpdateStatus(r.reqId, r.itemId, v)} /></td>
                  <td>
                    <button className="mrs-btn-danger" style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4 }}
                      onClick={() => { if (confirm('이 요청 전체를 삭제할까요?')) onDelete(r.reqId); }} title="요청 삭제">
                      <X size={15} />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [session, setSession] = useState(null);
  const [tab, setTab] = useState('form');
  const [requests, setRequests] = useState([]);
  const [projects, setProjects] = useState(DEFAULT_PROJECTS);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [savingSettings, setSavingSettings] = useState(false);
  const [error, setError] = useState(null);

  useEffect(() => { load(); }, []);

  async function load() {
    setLoading(true); setError(null);
    try {
      const [reqData, projData] = await Promise.all([
        apiGet('list'),
        apiGet('projects'),
      ]);
      setRequests(Array.isArray(reqData) ? reqData : []);
      setProjects(Array.isArray(projData) && projData.length ? projData : DEFAULT_PROJECTS);
    } catch (e) {
      setError('구글시트 연결에 실패했습니다. APPS_SCRIPT_URL 설정과 네트워크를 확인해주세요.');
      setRequests([]);
    } finally {
      setLoading(false);
    }
  }

  async function handleSubmit(payload) {
    const next = [payload, ...requests];
    setRequests(next); // 낙관적 업데이트
    setSaving(true); setError(null);
    try {
      await apiPost('addRequest', { payload });
    } catch (e) {
      setError('요청 저장에 실패했습니다. 다시 시도해주세요.');
      await load(); // 서버 상태로 복구
    } finally {
      setSaving(false);
    }
  }

  async function handleUpdateStatus(reqId, itemId, status) {
    const next = requests.map(r => r.id !== reqId ? r : {
      ...r, items: r.items.map(it => it.id !== itemId ? it : { ...it, status }),
    });
    setRequests(next);
    setSaving(true); setError(null);
    try {
      await apiPost('updateStatus', { itemId, status });
    } catch (e) {
      setError('상태 변경에 실패했습니다.');
      await load();
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete(reqId) {
    const next = requests.filter(r => r.id !== reqId);
    setRequests(next);
    setSaving(true); setError(null);
    try {
      await apiPost('deleteRequest', { reqId });
    } catch (e) {
      setError('삭제에 실패했습니다.');
      await load();
    } finally {
      setSaving(false);
    }
  }

  async function handleSaveProjects(next) {
    setSavingSettings(true); setError(null);
    try {
      await apiPost('saveProjects', { projects: next });
      setProjects(next);
    } catch (e) {
      setError('프로젝트 저장에 실패했습니다.');
    } finally {
      setSavingSettings(false);
    }
  }

  if (!session) {
    return <LoginScreen projects={projects} onLogin={setSession} loading={loading} />;
  }

  const currentProjectName = (projects.find(p => p.id === session.projectId) || {}).name || '(알 수 없음)';

  return (
    <div className="mrs-root">
      <GlobalStyle />
      <div className="mrs-header">
        <div>
          <div className="mrs-display" style={{ fontSize: 21, fontWeight: 600, letterSpacing: '0.03em' }}>자재 요청 관리 시스템</div>
          <div style={{ fontSize: 12, color: '#9AA5AC', marginTop: 2 }}>Material Requisition System · 건국이엔아이 공무팀</div>
        </div>
        <div className="mrs-header-right">
          <div className="mrs-project-chip">
            PROJECT&nbsp;<b>{currentProjectName}</b>
          </div>
          <span className="mrs-user-chip">{session.requester}</span>
          <button className="mrs-logout-btn" onClick={() => setSession(null)}>
            <LogOut size={13} /> 로그아웃
          </button>
        </div>
      </div>

      <div className="mrs-tabs">
        <button className={`mrs-tab ${tab === 'form' ? 'active' : ''}`} onClick={() => setTab('form')}>
          <ClipboardList size={16} /> 요청 입력 (팀장용)
        </button>
        <button className={`mrs-tab ${tab === 'admin' ? 'active' : ''}`} onClick={() => setTab('admin')}>
          <LayoutDashboard size={16} /> 관리자 현황판
        </button>
        {saving && <span style={{ marginLeft: 'auto', alignSelf: 'center', fontSize: 12, color: 'var(--ink-soft)', display:'flex', alignItems:'center', gap:6 }}><Loader2 size={13} className="mrs-spin" /> 저장 중...</span>}
      </div>

      <div className="mrs-body">
        {error && (
          <div className="mrs-card" style={{ padding: '10px 14px', marginBottom: 14, borderColor: '#F0A97A', background: '#FBE7DA', display: 'flex', alignItems: 'center', gap: 8, color: '#B84B10', fontSize: 13 }}>
            <AlertCircle size={16} /> {error}
          </div>
        )}
        {loading ? (
          <div className="mrs-empty"><Loader2 size={20} className="mrs-spin" /><div style={{ marginTop: 8 }}>불러오는 중...</div></div>
        ) : tab === 'form' ? (
          <RequestForm session={session} projectName={currentProjectName} onSubmit={handleSubmit} saving={saving} />
        ) : (
          <AdminDashboard
            requests={requests}
            projects={projects}
            onUpdateStatus={handleUpdateStatus}
            onDelete={handleDelete}
            onSaveProjects={handleSaveProjects}
            savingSettings={savingSettings}
          />
        )}
      </div>
    </div>
  );
}
