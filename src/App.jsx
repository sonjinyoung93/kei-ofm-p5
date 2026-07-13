import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Plus, Trash2, Download, ClipboardList, LayoutDashboard, Loader2, AlertCircle, CheckCircle2, X, Settings, Save, LogOut, Building2, Users, ListChecks, CalendarDays, Lock } from 'lucide-react';

const ROWS = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));
const COLS = Array.from({ length: 70 }, (_, i) => String(i + 1));
const PROCESSES = ['자탐', '유도등', '무통'];
const UNITS = ['EA', 'M', 'SET', 'BOX', 'ROLL', 'KG'];
const STATUS_FLOW = ['요청됨', '확인됨', '발주완료', '입고완료'];
const STATUS_COLOR = {
  '요청됨':   { bg: '#EEF0EC', fg: '#5C6B73', bd: '#C9CFC6' },
  '확인됨':   { bg: '#E3ECF5', fg: '#2B5A8C', bd: '#9FBEDC' },
  '발주완료': { bg: '#FBE7DA', fg: '#B84B10', bd: '#F0A97A' },
  '입고완료': { bg: '#E1EBE3', fg: '#2E6B47', bd: '#9FC7AC' },
};

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

const ITEM_UNITS = {
  '무나사전선관': ['본', 'M'],
  '커플링': ['EA', 'BOX'],
  '박스커넥터': ['EA', 'BOX'],
  '통신케이블': ['롤', 'M'],
  '내화케이블': ['M'],
};
function getUnits(name) { return ITEM_UNITS[name] || UNITS; }

const DEFAULT_PROJECTS = [
  { id: 'proj-1', name: 'P4 Ph4 (삼성물산)' },
  { id: 'proj-2', name: '신규 프로젝트' },
];

// ⚠️ 구글 Apps Script를 "웹 앱으로 배포"한 뒤 나오는 URL로 반드시 교체하세요.
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwVPkUPP93WpzWhZLZ4F4fOkQkC5HQ-h1UM7cK72MzdCJ4oxfT5ALwA6GsDj-LSJbpW/exec';

async function apiGet(action) {
  const res = await fetch(`${APPS_SCRIPT_URL}?action=${action}`);
  if (!res.ok) throw new Error(`GET ${action} failed: ${res.status}`);
  return res.json();
}
async function apiPost(action, data) {
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
function isToday(iso) {
  const d = new Date(iso);
  const now = new Date();
  return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth() && d.getDate() === now.getDate();
}
function fmtZone(rowFrom, rowTo, colFrom, colTo) {
  const rowPart = rowFrom === rowTo ? `${rowFrom}행` : `${rowFrom}~${rowTo}행`;
  const colPart = colFrom === colTo ? `${colFrom}열` : `${colFrom}~${colTo}열`;
  return `${rowPart} ${colPart}`;
}
function rowToNum(r) { return r ? r.charCodeAt(0) - 64 : null; }
// req의 구역이 order의 구역 범위 안에 완전히 포함되는지 확인
function zoneContained(req, order) {
  if (!req.rowFrom || !req.rowTo || !req.colFrom || !req.colTo) return false;
  if (!order.rowFrom || !order.rowTo || !order.colFrom || !order.colTo) return false;
  if ((req.floor || '') !== (order.floor || '')) return false;
  const rF = rowToNum(req.rowFrom), rT = rowToNum(req.rowTo);
  const cF = Number(req.colFrom), cT = Number(req.colTo);
  const oRF = rowToNum(order.rowFrom), oRT = rowToNum(order.rowTo);
  const oCF = Number(order.colFrom), oCT = Number(order.colTo);
  return rF >= oRF && rT <= oRT && cF >= oCF && cT <= oCT;
}
function sameItem(a, b) { return a.name === b.itemName && a.spec === b.itemSpec && a.color === b.itemColor; }
function newItemRow() {
  const name = CATALOG_NAMES[0];
  const spec = getSpecs(name)[0];
  const color = getColors(name, spec)[0];
  const unit = getUnits(name)[0];
  return { id: genId(), name, spec, color, qty: '', unit };
}

const GlobalStyle = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');
    .mrs-root {
      --paper: #F0EEE6; --paper-line: #D8D4C6; --ink: #1C2A33; --ink-soft: #5C6B73;
      --steel: #8A94A6; --accent: #D9601B; --accent-dark: #B84B10; --line: #C9CFC6; --card: #FBFAF6;
      font-family: 'Apple SD Gothic Neo','Malgun Gothic',-apple-system,sans-serif;
      color: var(--ink); min-height: 100%;
      background:
        repeating-linear-gradient(0deg, transparent, transparent 27px, rgba(28,42,51,0.035) 28px),
        repeating-linear-gradient(90deg, transparent, transparent 27px, rgba(28,42,51,0.035) 28px),
        var(--paper);
    }
    .mrs-mono { font-family: 'JetBrains Mono', monospace; }
    .mrs-display { font-family: 'Oswald', 'Apple SD Gothic Neo', sans-serif; letter-spacing: 0.02em; }
    .mrs-header { background: var(--ink); color: #EDEAE0; padding: 18px 20px; display: flex; justify-content: space-between; align-items: center; gap: 16px; flex-wrap: wrap; border-bottom: 3px solid var(--accent); }
    .mrs-project-chip { display: flex; align-items: center; gap: 8px; border: 1px solid #3A4954; border-radius: 20px; padding: 6px 14px; font-size: 12px; color: #C9CFC6; }
    .mrs-project-chip b { color: #EDEAE0; font-weight: 600; font-size: 13px; }
    .mrs-header-right { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
    .mrs-user-chip { font-size: 12px; color: #C9CFC6; display: flex; align-items: center; gap: 6px; }
    .mrs-logout-btn { display: flex; align-items: center; gap: 6px; background: none; border: 1px solid #3A4954; color: #C9CFC6; border-radius: 20px; padding: 6px 12px; font-size: 12px; cursor: pointer; }
    .mrs-logout-btn:hover { border-color: var(--accent); color: #EDEAE0; }
    .mrs-tabs { display: flex; gap: 0; background: var(--card); border-bottom: 1px solid var(--line); padding: 0 16px; flex-wrap: wrap; }
    .mrs-tab { display: flex; align-items: center; gap: 7px; padding: 12px 18px; font-size: 14px; font-weight: 600; color: var(--ink-soft); cursor: pointer; border-bottom: 3px solid transparent; transition: all .15s ease; background: none; border-top:none; border-left:none; border-right:none; }
    .mrs-tab.active { color: var(--ink); border-bottom-color: var(--accent); }
    .mrs-tab:hover:not(.active) { color: var(--ink); background: rgba(217,96,27,0.05); }
    .mrs-body { padding: 20px; max-width: 1150px; margin: 0 auto; }
    .mrs-card { background: var(--card); border: 1px solid var(--line); border-radius: 3px; box-shadow: 0 1px 2px rgba(28,42,51,0.06); }
    .mrs-field-label { font-size: 12px; font-weight: 600; color: var(--ink-soft); margin-bottom: 6px; display: block; text-transform: uppercase; letter-spacing: 0.04em; }
    .mrs-input, .mrs-select, .mrs-textarea { width: 100%; box-sizing: border-box; padding: 9px 10px; border: 1px solid var(--line); border-radius: 2px; background: #fff; font-size: 14px; color: var(--ink); font-family: inherit; }
    .mrs-input:focus, .mrs-select:focus, .mrs-textarea:focus { outline: none; border-color: var(--accent); box-shadow: 0 0 0 3px rgba(217,96,27,0.12); }
    .mrs-input:disabled, .mrs-select:disabled { background: #F1EFE9; color: var(--ink-soft); }
    .mrs-top-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px; }
    @media (max-width: 480px) { .mrs-top-grid { gap: 6px; } .mrs-top-grid .mrs-field-label { font-size: 10px; } .mrs-top-grid input, .mrs-top-grid select { font-size: 12px; padding: 7px 6px; } }
    .mrs-zone-grid { display: grid; grid-template-columns: 1fr 1fr 1fr 1fr; gap: 8px; }
    @media (max-width: 480px) { .mrs-zone-grid { gap: 5px; } .mrs-zone-grid select { font-size: 12px; padding: 7px 3px; } }
    .mrs-item-row { display: grid; grid-template-columns: 1.6fr 1fr 0.7fr 0.6fr 0.7fr auto; gap: 6px; align-items: end; padding: 10px 0; border-bottom: 1px dashed var(--line); }
    @media (max-width: 480px) { .mrs-item-row { gap: 4px; } .mrs-item-row select, .mrs-item-row input { font-size: 12px; padding: 7px 3px; } }
    .mrs-btn { display: inline-flex; align-items: center; gap: 6px; padding: 10px 16px; border-radius: 2px; font-size: 14px; font-weight: 600; cursor: pointer; border: 1px solid transparent; transition: all .15s ease; }
    .mrs-btn-primary { background: var(--accent); color: #fff; }
    .mrs-btn-primary:hover { background: var(--accent-dark); }
    .mrs-btn-ghost { background: transparent; color: var(--ink-soft); border-color: var(--line); }
    .mrs-btn-ghost:hover { border-color: var(--ink-soft); color: var(--ink); }
    .mrs-btn-danger { background: transparent; color: #B84B10; }
    .mrs-btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .mrs-chip { display: inline-flex; align-items: center; padding: 3px 9px; border-radius: 20px; font-size: 12px; font-weight: 600; border: 1px solid; white-space: nowrap; }
    .mrs-table-wrap { overflow-x: auto; border: 1px solid var(--line); border-radius: 3px; background: var(--card); }
    table.mrs-table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 950px; }
    table.mrs-table th { text-align: left; padding: 10px 12px; background: var(--ink); color: #C9CFC6; font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; position: sticky; top: 0; }
    table.mrs-table td { padding: 10px 12px; border-bottom: 1px solid var(--paper-line); vertical-align: middle; }
    table.mrs-table tr:hover td { background: rgba(217,96,27,0.04); }
    .mrs-stat { display: flex; flex-direction: column; gap: 2px; padding: 10px 16px; border-right: 1px solid var(--line); }
    .mrs-stat:last-child { border-right: none; }
    .mrs-stat .n { font-size: 22px; font-weight: 700; }
    .mrs-stat .l { font-size: 11px; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.04em; }
    select.mrs-status-select { font-size: 12px; font-weight: 600; border-radius: 20px; padding: 4px 24px 4px 10px; border: 1px solid; cursor: pointer; -webkit-appearance: none; appearance: none; background-repeat: no-repeat; background-position: right 8px center; background-size: 10px; }
    .mrs-empty { text-align: center; padding: 40px 20px; color: var(--ink-soft); }
    .mrs-spin { animation: mrs-spin 1s linear infinite; }
    @keyframes mrs-spin { to { transform: rotate(360deg); } }
    .mrs-login-wrap { min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 20px; }
    .mrs-login-card { width: 100%; max-width: 380px; background: var(--card); border: 1px solid var(--line); border-top: 4px solid var(--accent); border-radius: 3px; padding: 32px 28px; box-shadow: 0 4px 16px rgba(28,42,51,0.08); }
    .mrs-role-badge { display: inline-flex; align-items: center; gap: 5px; font-size: 11px; font-weight: 600; padding: 2px 9px; border-radius: 20px; }
  `}</style>
);

const FALLBACK_STATUS_COLOR = { bg: '#EEF0EC', fg: '#5C6B73', bd: '#C9CFC6' };

function StatusSelect({ value, onChange }) {
  const c = STATUS_COLOR[value] || FALLBACK_STATUS_COLOR;
  const safeValue = STATUS_FLOW.includes(value) ? value : STATUS_FLOW[0];
  return (
    <select className="mrs-status-select" value={safeValue} onChange={e => onChange(e.target.value)} style={{ background: c.bg, color: c.fg, borderColor: c.bd }}>
      {STATUS_FLOW.map(s => <option key={s} value={s}>{s}</option>)}
    </select>
  );
}

function StatusBadge({ value }) {
  const c = STATUS_COLOR[value] || FALLBACK_STATUS_COLOR;
  return <span className="mrs-chip" style={{ background: c.bg, color: c.fg, borderColor: c.bd }}>{value || '알 수 없음'}</span>;
}

function ItemRowEditor({ item, onChange, onRemove, removable }) {
  const specs = getSpecs(item.name);
  const colors = getColors(item.name, item.spec);
  const units = getUnits(item.name);
  function handleNameChange(name) {
    const spec = getSpecs(name)[0];
    const color = getColors(name, spec)[0];
    const unit = getUnits(name)[0];
    onChange({ ...item, name, spec, color, unit });
  }
  function handleSpecChange(spec) {
    const color = getColors(item.name, spec)[0];
    onChange({ ...item, spec, color });
  }
  return (
    <div className="mrs-item-row">
      <div><select className="mrs-select" value={item.name} onChange={e => handleNameChange(e.target.value)}>{CATALOG_NAMES.map(n => <option key={n} value={n}>{n}</option>)}</select></div>
      <div><select className="mrs-select" value={item.spec} onChange={e => handleSpecChange(e.target.value)}>{specs.map(s => <option key={s} value={s}>{s}</option>)}</select></div>
      <div><input className="mrs-input" type="number" min="0" value={item.qty} onChange={e => onChange({ ...item, qty: e.target.value })} placeholder="0" /></div>
      <div><select className="mrs-select" value={item.unit} onChange={e => onChange({ ...item, unit: e.target.value })} disabled={units.length <= 1}>{units.map(u => <option key={u} value={u}>{u}</option>)}</select></div>
      <div><select className="mrs-select" value={item.color} onChange={e => onChange({ ...item, color: e.target.value })} disabled={colors.length <= 1}>{colors.map(c => <option key={c} value={c}>{c}</option>)}</select></div>
      <button className="mrs-btn mrs-btn-danger" onClick={onRemove} disabled={!removable} title="삭제" style={{ padding: 8 }}><Trash2 size={16} /></button>
    </div>
  );
}

// ── 로그인 화면 ─────────────────────────────────────────
function LoginScreen({ onLogin, loading, error }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');

  function handleSubmit() {
    if (!username.trim() || !password) { alert('아이디와 비밀번호를 입력해주세요.'); return; }
    onLogin(username.trim(), password);
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
          <div style={{ fontSize: 12, color: 'var(--ink-soft)', marginBottom: 24 }}>로그인 · 건국이엔아이 공무팀</div>

          <label className="mrs-field-label">아이디</label>
          <input className="mrs-input" value={username} onChange={e => setUsername(e.target.value)} onKeyDown={e => e.key === 'Enter' && handleSubmit()} placeholder="아이디" style={{ marginBottom: 14 }} />

          <label className="mrs-field-label">비밀번호</label>
          <input className="mrs-input" type="password" value={password} onChange={e => setPassword(e.target.value)} onKeyDown={e => e.key === 'Enter' && handleSubmit()} placeholder="비밀번호" style={{ marginBottom: 14 }} />

          {error && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, color: '#B84B10', fontSize: 12, marginBottom: 14 }}>
              <AlertCircle size={14} /> {error}
            </div>
          )}

          <button className="mrs-btn mrs-btn-primary" style={{ width: '100%', justifyContent: 'center' }} onClick={handleSubmit} disabled={loading}>
            {loading ? <Loader2 size={15} className="mrs-spin" /> : <Lock size={15} />}
            로그인
          </button>
        </div>
      </div>
    </div>
  );
}

// ── 팀장: 요청 입력 폼 ──────────────────────────────────
function RequestForm({ session, projectName, onSubmit, saving }) {
  const [floor, setFloor] = useState('1F');
  const [rowFrom, setRowFrom] = useState('A');
  const [rowTo, setRowTo] = useState('A');
  const [colFrom, setColFrom] = useState('1');
  const [colTo, setColTo] = useState('1');
  const [process, setProcess] = useState(PROCESSES[0]);
  const [note, setNote] = useState('');
  const [items, setItems] = useState([newItemRow()]);
  const [justSubmitted, setJustSubmitted] = useState(false);

  function updateItem(id, updated) { setItems(items.map(it => it.id === id ? updated : it)); }
  function addItem() { setItems([...items, newItemRow()]); }
  function removeItem(id) { if (items.length === 1) return; setItems(items.filter(it => it.id !== id)); }

  function validate() {
    const valid = items.filter(it => it.name && it.qty !== '' && Number(it.qty) > 0);
    if (valid.length === 0) return '최소 1개 이상의 품목에 수량을 입력해주세요.';
    return null;
  }

  async function handleSubmit() {
    const err = validate();
    if (err) { alert(err); return; }
    const payload = {
      id: genId(), reqNo: genReqNo(), requester: session.name, username: session.username,
      projectId: session.projectId, zone: fmtZone(rowFrom, rowTo, colFrom, colTo),
      floor, rowFrom, rowTo, colFrom, colTo, process,
      note: note.trim(), createdAt: new Date().toISOString(),
      items: items.filter(it => it.name && it.qty !== '' && Number(it.qty) > 0)
        .map(it => ({ id: genId(), name: it.name, spec: it.spec, color: it.color, qty: it.qty, unit: it.unit, status: '요청됨' })),
    };
    await onSubmit(payload);
    setFloor('1F'); setRowFrom('A'); setRowTo('A'); setColFrom('1'); setColTo('1');
    setProcess(PROCESSES[0]); setNote(''); setItems([newItemRow()]);
    setJustSubmitted(true);
    setTimeout(() => setJustSubmitted(false), 3000);
  }

  return (
    <div className="mrs-card" style={{ padding: 22 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 18 }}>
        <h2 className="mrs-display" style={{ fontSize: 18, margin: 0, fontWeight: 600 }}>자재 요청 시트</h2>
        <span className="mrs-mono" style={{ fontSize: 11, color: 'var(--ink-soft)' }}>NEW REQUEST</span>
      </div>

      <div className="mrs-top-grid" style={{ marginBottom: 14 }}>
        <div><label className="mrs-field-label">요청자</label><input className="mrs-input" value={session.name} disabled /></div>
        <div><label className="mrs-field-label">프로젝트</label><input className="mrs-input" value={projectName} disabled /></div>
        <div><label className="mrs-field-label">공정</label><select className="mrs-select" value={process} onChange={e => setProcess(e.target.value)}>{PROCESSES.map(p => <option key={p} value={p}>{p}</option>)}</select></div>
      </div>

      <div style={{ marginBottom: 14 }}>
        <label className="mrs-field-label">층 / 구역 (행 / 열)</label>
        <input className="mrs-input" value={floor} onChange={e => setFloor(e.target.value)} placeholder="층 (예: 1F)" style={{ marginBottom: 6, maxWidth: 140 }} />
        <div className="mrs-zone-grid">
          <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>행 시작</span><select className="mrs-select" value={rowFrom} onChange={e => setRowFrom(e.target.value)}>{ROWS.map(r => <option key={r} value={r}>{r}</option>)}</select></div>
          <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>행 끝</span><select className="mrs-select" value={rowTo} onChange={e => setRowTo(e.target.value)}>{ROWS.map(r => <option key={r} value={r}>{r}</option>)}</select></div>
          <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>열 시작</span><select className="mrs-select" value={colFrom} onChange={e => setColFrom(e.target.value)}>{COLS.map(c => <option key={c} value={c}>{c}</option>)}</select></div>
          <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>열 끝</span><select className="mrs-select" value={colTo} onChange={e => setColTo(e.target.value)}>{COLS.map(c => <option key={c} value={c}>{c}</option>)}</select></div>
        </div>
      </div>

      <div style={{ marginBottom: 8 }}><label className="mrs-field-label">요청 품목</label></div>
      <div className="mrs-item-row" style={{ borderBottom: 'none', paddingBottom: 0 }}>
        <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>품목</span><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>규격</span>
        <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>수량</span><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>단위</span>
        <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>색상</span><span></span>
      </div>
      {items.map(it => <ItemRowEditor key={it.id} item={it} onChange={u => updateItem(it.id, u)} onRemove={() => removeItem(it.id)} removable={items.length > 1} />)}
      <button className="mrs-btn mrs-btn-ghost" style={{ marginTop: 10 }} onClick={addItem}><Plus size={15} /> 품목 추가</button>

      <div style={{ marginTop: 18 }}>
        <label className="mrs-field-label">특이사항 (선택)</label>
        <input className="mrs-input" value={note} onChange={e => setNote(e.target.value)} placeholder="" />
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 22 }}>
        <button className="mrs-btn mrs-btn-primary" onClick={handleSubmit} disabled={saving}>
          {saving ? <Loader2 size={15} className="mrs-spin" /> : <ClipboardList size={15} />} 요청 제출
        </button>
        {justSubmitted && <span style={{ display: 'flex', alignItems: 'center', gap: 6, color: '#2E6B47', fontSize: 13, fontWeight: 600 }}><CheckCircle2 size={16} /> 요청이 접수되었습니다</span>}
      </div>
    </div>
  );
}

// ── 팀장: 본인 요청 리스트 ──────────────────────────────
function MyRequestList({ requests, session }) {
  const rows = [];
  requests
    .filter(r => (r.username ? r.username === session.username : r.requester === session.name))
    .forEach(r => {
      r.items.forEach(it => {
        rows.push({ reqNo: r.reqNo, createdAt: r.createdAt, zone: r.zone, process: r.process, note: r.note, ...it });
      });
    });
  rows.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  return (
    <div>
      <h2 className="mrs-display" style={{ fontSize: 18, margin: '0 0 14px', fontWeight: 600 }}>요청리스트</h2>
      {rows.length === 0 ? (
        <div className="mrs-card mrs-empty">제출한 요청이 없습니다.</div>
      ) : (
        <div className="mrs-table-wrap">
          <table className="mrs-table">
            <thead><tr><th>요청번호</th><th>요청일시</th><th>구역</th><th>공정</th><th>품목명</th><th>규격</th><th>수량</th><th>상태</th></tr></thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={i}>
                  <td className="mrs-mono" style={{ fontSize: 12, color: 'var(--ink-soft)' }}>{r.reqNo}</td>
                  <td className="mrs-mono" style={{ fontSize: 12 }}>{fmtDate(r.createdAt)}</td>
                  <td>{r.zone}</td><td>{r.process}</td>
                  <td style={{ fontWeight: 600 }}>{r.name}</td><td style={{ color: 'var(--ink-soft)' }}>{r.spec || '-'}</td>
                  <td className="mrs-mono">{r.qty} {r.unit}</td>
                  <td><StatusBadge value={r.status} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ── 팀장 전체 화면 ──────────────────────────────────────
function LeaderApp({ session, requests, projects, onSubmit, saving }) {
  const [tab, setTab] = useState('form');
  const projectName = (projects.find(p => p.id === session.projectId) || {}).name || '(알 수 없음)';
  return (
    <div className="mrs-body">
      <div className="mrs-tabs" style={{ margin: '-20px -20px 20px', padding: '0 20px' }}>
        <button className={`mrs-tab ${tab === 'form' ? 'active' : ''}`} onClick={() => setTab('form')}><ClipboardList size={16} /> 자재요청 시트</button>
        <button className={`mrs-tab ${tab === 'mylist' ? 'active' : ''}`} onClick={() => setTab('mylist')}><ListChecks size={16} /> 요청리스트</button>
      </div>
      {tab === 'form'
        ? <RequestForm session={session} projectName={projectName} onSubmit={onSubmit} saving={saving} />
        : <MyRequestList requests={requests} session={session} />}
    </div>
  );
}

// ── 관리자: 요청 테이블 (금일/누계 공용) ─────────────────
function RequestsTable({ requests, projects, onUpdateStatus, onDelete, scope }) {
  const [projectFilter, setProjectFilter] = useState('전체');
  const [zoneQuery, setZoneQuery] = useState('');
  const [processFilter, setProcessFilter] = useState('전체');
  const [statusFilter, setStatusFilter] = useState('전체');

  const projectNameById = {};
  projects.forEach(p => { projectNameById[p.id] = p.name; });

  const rows = [];
  requests.forEach(r => {
    if (scope === 'today' && !isToday(r.createdAt)) return;
    r.items.forEach(it => {
      rows.push({
        reqId: r.id, itemId: it.id, reqNo: r.reqNo, requester: r.requester,
        projectId: r.projectId, projectName: projectNameById[r.projectId] || '(삭제된 프로젝트)',
        zone: r.zone, process: r.process, name: it.name, spec: it.spec, color: it.color,
        qty: it.qty, unit: it.unit, status: it.status, note: r.note, createdAt: r.createdAt,
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
      '프로젝트': r.projectName, '요청번호': r.reqNo, '요청일시': fmtDate(r.createdAt), '요청자': r.requester,
      '구역': r.zone, '공정': r.process, '품목명': r.name, '규격': r.spec, '색상': r.color,
      '수량': r.qty, '단위': r.unit, '상태': r.status, '특이사항': r.note || '',
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{wch:16},{wch:18},{wch:16},{wch:12},{wch:14},{wch:8},{wch:16},{wch:12},{wch:8},{wch:8},{wch:8},{wch:10},{wch:30}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '자재요청');
    const d = new Date();
    const label = (scope === 'today' ? '금일_' : '누계_') + (projectFilter === '전체' ? '전체프로젝트' : (projectNameById[projectFilter] || '프로젝트'));
    XLSX.writeFile(wb, `${label}_자재요청_${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}.xlsx`);
  }

  return (
    <div>
      <div style={{ display: 'flex', gap: 10, marginBottom: 14, flexWrap: 'wrap' }}>
        <select className="mrs-select" style={{ width: 'auto', fontWeight: 600 }} value={projectFilter} onChange={e => setProjectFilter(e.target.value)}>
          <option value="전체">전체 프로젝트</option>
          {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
      </div>

      <div className="mrs-card" style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 16 }}>
        <div className="mrs-stat"><span className="n mrs-mono">{totalCount}</span><span className="l">{scope === 'today' ? '금일 요청' : '전체 요청'}</span></div>
        {STATUS_FLOW.map(s => (
          <div className="mrs-stat" key={s}><span className="n mrs-mono" style={{ color: STATUS_COLOR[s].fg }}>{counts[s]}</span><span className="l">{s}</span></div>
        ))}
      </div>

      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 14, alignItems: 'center' }}>
        <input className="mrs-input" style={{ width: 160 }} value={zoneQuery} onChange={e => setZoneQuery(e.target.value)} placeholder="구역 검색 (예: A행)" />
        <select className="mrs-select" style={{ width: 'auto' }} value={processFilter} onChange={e => setProcessFilter(e.target.value)}>
          <option value="전체">전체 공정</option>{PROCESSES.map(p => <option key={p} value={p}>{p}</option>)}
        </select>
        <select className="mrs-select" style={{ width: 'auto' }} value={statusFilter} onChange={e => setStatusFilter(e.target.value)}>
          {statuses.map(s => <option key={s} value={s}>{s === '전체' ? '전체 상태' : s}</option>)}
        </select>
        <div style={{ flex: 1 }} />
        <button className="mrs-btn mrs-btn-primary" onClick={exportExcel} disabled={filtered.length === 0}><Download size={15} /> 엑셀로 내보내기</button>
      </div>

      {filtered.length === 0 ? (
        <div className="mrs-card mrs-empty">표시할 요청이 없습니다.</div>
      ) : (
        <div className="mrs-table-wrap">
          <table className="mrs-table">
            <thead><tr><th>프로젝트</th><th>요청번호</th><th>요청일시</th><th>요청자</th><th>구역</th><th>공정</th><th>품목명</th><th>규격</th><th>색상</th><th>수량</th><th>상태</th><th></th></tr></thead>
            <tbody>
              {filtered.map(r => (
                <tr key={r.itemId}>
                  <td style={{ fontWeight: 600 }}>{r.projectName}</td>
                  <td className="mrs-mono" style={{ fontSize: 12, color: 'var(--ink-soft)' }}>{r.reqNo}</td>
                  <td className="mrs-mono" style={{ fontSize: 12 }}>{fmtDate(r.createdAt)}</td>
                  <td>{r.requester}</td><td>{r.zone}</td><td>{r.process}</td>
                  <td style={{ fontWeight: 600 }}>{r.name}</td><td style={{ color: 'var(--ink-soft)' }}>{r.spec || '-'}</td>
                  <td style={{ color: 'var(--ink-soft)' }}>{r.color || '-'}</td>
                  <td className="mrs-mono">{r.qty} {r.unit}</td>
                  <td><StatusSelect value={r.status} onChange={v => onUpdateStatus(r.reqId, r.itemId, v)} /></td>
                  <td><button style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4, color: '#B84B10' }} onClick={() => { if (confirm('이 요청 전체를 삭제할까요?')) onDelete(r.reqId); }} title="요청 삭제"><X size={15} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ── 관리자: 프로젝트명 관리 ─────────────────────────────
function ProjectManager({ projects, onSave, saving }) {
  const [draft, setDraft] = useState(projects);
  useEffect(() => setDraft(projects), [projects]);
  return (
    <div className="mrs-card" style={{ padding: 18, marginBottom: 20 }}>
      <h3 className="mrs-display" style={{ fontSize: 15, margin: '0 0 12px', fontWeight: 600 }}>프로젝트명 관리</h3>
      {draft.map((p, idx) => (
        <div key={p.id} style={{ display: 'flex', gap: 10, marginBottom: 8, alignItems: 'center' }}>
          <input className="mrs-input" value={p.name} onChange={e => { const next = [...draft]; next[idx] = { ...p, name: e.target.value }; setDraft(next); }} />
        </div>
      ))}
      <button className="mrs-btn mrs-btn-primary" disabled={saving || draft.some(p => !p.name.trim())} onClick={() => onSave(draft)}><Save size={15} /> 저장</button>
    </div>
  );
}

// ── 관리자: 팀장 계정 관리 ──────────────────────────────
function StaffManager({ users, projects, onAdd, onUpdate, onDelete, saving }) {
  const staff = users.filter(u => u.role === 'leader' || u.role === 'material');
  const [editingId, setEditingId] = useState(null);
  const [editDraft, setEditDraft] = useState({});
  const [newStaff, setNewStaff] = useState({ name: '', username: '', password: '', role: 'leader', projectId: projects[0]?.id || '' });

  useEffect(() => {
    if (!newStaff.projectId && projects.length) setNewStaff(n => ({ ...n, projectId: projects[0].id }));
  }, [projects]);

  function startEdit(u) { setEditingId(u.id); setEditDraft({ name: u.name, username: u.username, password: '', role: u.role, projectId: u.projectId }); }
  function saveEdit(id) { onUpdate({ id, ...editDraft, projectId: editDraft.role === 'leader' ? editDraft.projectId : '' }); setEditingId(null); }

  function submitNew() {
    if (!newStaff.name.trim() || !newStaff.username.trim() || !newStaff.password) {
      alert('이름, 아이디, 비밀번호를 모두 입력해주세요.'); return;
    }
    onAdd({ ...newStaff, projectId: newStaff.role === 'leader' ? newStaff.projectId : '' });
    setNewStaff({ name: '', username: '', password: '', role: 'leader', projectId: projects[0]?.id || '' });
  }

  const projectName = id => (projects.find(p => p.id === id) || {}).name || '-';
  const roleLabel = r => r === 'material' ? '자재팀' : '팀장';

  return (
    <div className="mrs-card" style={{ padding: 18, marginBottom: 20 }}>
      <h3 className="mrs-display" style={{ fontSize: 15, margin: '0 0 12px', fontWeight: 600 }}>팀장 · 자재팀 계정 관리</h3>

      {staff.length === 0 ? (
        <p style={{ fontSize: 13, color: 'var(--ink-soft)', marginBottom: 16 }}>등록된 계정이 없습니다.</p>
      ) : (
        <div style={{ marginBottom: 18 }}>
          {staff.map(u => (
            <div key={u.id} style={{ borderBottom: '1px solid var(--paper-line)', padding: '10px 0' }}>
              {editingId === u.id ? (
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 1fr 1fr auto auto', gap: 8, alignItems: 'center' }}>
                  <input className="mrs-input" value={editDraft.name} onChange={e => setEditDraft({ ...editDraft, name: e.target.value })} placeholder="이름" />
                  <input className="mrs-input" value={editDraft.username} onChange={e => setEditDraft({ ...editDraft, username: e.target.value })} placeholder="아이디" />
                  <input className="mrs-input" type="password" value={editDraft.password} onChange={e => setEditDraft({ ...editDraft, password: e.target.value })} placeholder="새 비밀번호(변경시)" />
                  <select className="mrs-select" value={editDraft.role} onChange={e => setEditDraft({ ...editDraft, role: e.target.value })}>
                    <option value="leader">팀장</option><option value="material">자재팀</option>
                  </select>
                  {editDraft.role === 'leader' ? (
                    <select className="mrs-select" value={editDraft.projectId} onChange={e => setEditDraft({ ...editDraft, projectId: e.target.value })}>
                      {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                    </select>
                  ) : <span style={{ fontSize: 12, color: 'var(--ink-soft)' }}>전체 프로젝트</span>}
                  <button className="mrs-btn mrs-btn-primary" onClick={() => saveEdit(u.id)} disabled={saving}>저장</button>
                  <button className="mrs-btn mrs-btn-ghost" onClick={() => setEditingId(null)}>취소</button>
                </div>
              ) : (
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
                  <div style={{ fontSize: 14 }}>
                    <span className="mrs-chip" style={{ marginRight: 8, background: u.role === 'material' ? '#FBE7DA' : '#E3ECF5', color: u.role === 'material' ? '#B84B10' : '#2B5A8C', borderColor: u.role === 'material' ? '#F0A97A' : '#9FBEDC' }}>{roleLabel(u.role)}</span>
                    <b>{u.name}</b>
                    <span style={{ color: 'var(--ink-soft)', marginLeft: 8 }}>@{u.username}</span>
                    {u.role === 'leader' && <span className="mrs-chip" style={{ marginLeft: 10, background: '#EEF0EC', color: '#5C6B73', borderColor: '#C9CFC6' }}>{projectName(u.projectId)}</span>}
                  </div>
                  <div style={{ display: 'flex', gap: 8 }}>
                    <button className="mrs-btn mrs-btn-ghost" onClick={() => startEdit(u)}>수정</button>
                    <button className="mrs-btn mrs-btn-danger" onClick={() => { if (confirm(`${u.name} 계정을 삭제할까요?`)) onDelete(u.id); }}>삭제</button>
                  </div>
                </div>
              )}
            </div>
          ))}
        </div>
      )}

      <div style={{ borderTop: '1px solid var(--line)', paddingTop: 14 }}>
        <label className="mrs-field-label">새 계정 추가</label>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 1fr 1fr auto', gap: 8 }}>
          <input className="mrs-input" value={newStaff.name} onChange={e => setNewStaff({ ...newStaff, name: e.target.value })} placeholder="이름" />
          <input className="mrs-input" value={newStaff.username} onChange={e => setNewStaff({ ...newStaff, username: e.target.value })} placeholder="아이디" />
          <input className="mrs-input" type="password" value={newStaff.password} onChange={e => setNewStaff({ ...newStaff, password: e.target.value })} placeholder="비밀번호" />
          <select className="mrs-select" value={newStaff.role} onChange={e => setNewStaff({ ...newStaff, role: e.target.value })}>
            <option value="leader">팀장</option><option value="material">자재팀</option>
          </select>
          {newStaff.role === 'leader' ? (
            <select className="mrs-select" value={newStaff.projectId} onChange={e => setNewStaff({ ...newStaff, projectId: e.target.value })}>
              {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          ) : <span style={{ fontSize: 12, color: 'var(--ink-soft)', alignSelf: 'center' }}>전체 프로젝트</span>}
          <button className="mrs-btn mrs-btn-primary" onClick={submitNew} disabled={saving}><Plus size={15} /> 추가</button>
        </div>
      </div>
    </div>
  );
}

// ── 관리자: 본인 계정(아이디/비번) 변경 ──────────────────
function AdminAccountManager({ session, onUpdate, saving }) {
  const [username, setUsername] = useState(session.username);
  const [password, setPassword] = useState('');
  return (
    <div className="mrs-card" style={{ padding: 18 }}>
      <h3 className="mrs-display" style={{ fontSize: 15, margin: '0 0 12px', fontWeight: 600 }}>관리자 계정</h3>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr auto', gap: 8, alignItems: 'end' }}>
        <div><label className="mrs-field-label">아이디</label><input className="mrs-input" value={username} onChange={e => setUsername(e.target.value)} /></div>
        <div><label className="mrs-field-label">새 비밀번호 (변경시에만 입력)</label><input className="mrs-input" type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="변경하지 않으면 비워두세요" /></div>
        <button className="mrs-btn mrs-btn-primary" disabled={saving} onClick={() => { onUpdate({ id: session.id, username, password: password || undefined }); setPassword(''); }}><Save size={15} /> 저장</button>
      </div>
    </div>
  );
}

// ── 공무팀: 발주요청(지급자재 산출 물량) 입력 ─────────────
function OrderManager({ orders, session, onAdd, onDelete, saving }) {
  const [floor, setFloor] = useState('1F');
  const [rowFrom, setRowFrom] = useState('A');
  const [rowTo, setRowTo] = useState('A');
  const [colFrom, setColFrom] = useState('1');
  const [colTo, setColTo] = useState('1');
  const [name, setName] = useState(CATALOG_NAMES[0]);
  const [spec, setSpec] = useState(getSpecs(CATALOG_NAMES[0])[0]);
  const [color, setColor] = useState(getColors(CATALOG_NAMES[0], getSpecs(CATALOG_NAMES[0])[0])[0]);
  const [unit, setUnit] = useState(getUnits(CATALOG_NAMES[0])[0]);
  const [qty, setQty] = useState('');
  const [note, setNote] = useState('');

  const specs = getSpecs(name);
  const colors = getColors(name, spec);
  const units = getUnits(name);

  function handleNameChange(n) {
    const s = getSpecs(n)[0]; setName(n); setSpec(s);
    setColor(getColors(n, s)[0]); setUnit(getUnits(n)[0]);
  }
  function handleSpecChange(s) { setSpec(s); setColor(getColors(name, s)[0]); }

  function submit() {
    if (!qty || Number(qty) <= 0) { alert('발주수량을 입력해주세요.'); return; }
    onAdd({
      itemName: name, itemSpec: spec, itemColor: color,
      floor, rowFrom, rowTo, colFrom, colTo, qty, unit, note: note.trim(), createdBy: session.name,
    });
    setQty(''); setNote('');
  }

  return (
    <div>
      <div className="mrs-card" style={{ padding: 18, marginBottom: 20 }}>
        <h3 className="mrs-display" style={{ fontSize: 15, margin: '0 0 12px', fontWeight: 600 }}>발주요청 등록 (지급자재 산출 물량)</h3>

        <label className="mrs-field-label">층 / 구간 (행 / 열)</label>
        <input className="mrs-input" value={floor} onChange={e => setFloor(e.target.value)} placeholder="층 (예: 1F)" style={{ marginBottom: 6, maxWidth: 140 }} />
        <div className="mrs-zone-grid" style={{ marginBottom: 12 }}>
          <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>행 시작</span><select className="mrs-select" value={rowFrom} onChange={e => setRowFrom(e.target.value)}>{ROWS.map(r => <option key={r} value={r}>{r}</option>)}</select></div>
          <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>행 끝</span><select className="mrs-select" value={rowTo} onChange={e => setRowTo(e.target.value)}>{ROWS.map(r => <option key={r} value={r}>{r}</option>)}</select></div>
          <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>열 시작</span><select className="mrs-select" value={colFrom} onChange={e => setColFrom(e.target.value)}>{COLS.map(c => <option key={c} value={c}>{c}</option>)}</select></div>
          <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>열 끝</span><select className="mrs-select" value={colTo} onChange={e => setColTo(e.target.value)}>{COLS.map(c => <option key={c} value={c}>{c}</option>)}</select></div>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1.4fr 1fr 0.8fr 0.8fr 0.7fr', gap: 8, marginBottom: 12 }}>
          <div><label className="mrs-field-label">품목</label><select className="mrs-select" value={name} onChange={e => handleNameChange(e.target.value)}>{CATALOG_NAMES.map(n => <option key={n} value={n}>{n}</option>)}</select></div>
          <div><label className="mrs-field-label">규격</label><select className="mrs-select" value={spec} onChange={e => handleSpecChange(e.target.value)}>{specs.map(s => <option key={s} value={s}>{s}</option>)}</select></div>
          <div><label className="mrs-field-label">색상</label><select className="mrs-select" value={color} onChange={e => setColor(e.target.value)} disabled={colors.length <= 1}>{colors.map(c => <option key={c} value={c}>{c}</option>)}</select></div>
          <div><label className="mrs-field-label">발주수량</label><input className="mrs-input" type="number" min="0" value={qty} onChange={e => setQty(e.target.value)} placeholder="0" /></div>
          <div><label className="mrs-field-label">단위</label><select className="mrs-select" value={unit} onChange={e => setUnit(e.target.value)} disabled={units.length <= 1}>{units.map(u => <option key={u} value={u}>{u}</option>)}</select></div>
        </div>

        <label className="mrs-field-label">비고 (선택)</label>
        <input className="mrs-input" value={note} onChange={e => setNote(e.target.value)} placeholder="" style={{ marginBottom: 14 }} />

        <button className="mrs-btn mrs-btn-primary" onClick={submit} disabled={saving}><Plus size={15} /> 발주요청 등록</button>
      </div>

      <div className="mrs-table-wrap">
        <table className="mrs-table">
          <thead><tr><th>층</th><th>구간</th><th>품목</th><th>규격</th><th>색상</th><th>발주수량</th><th>등록자</th><th>등록일</th><th></th></tr></thead>
          <tbody>
            {orders.length === 0 ? (
              <tr><td colSpan={9} style={{ textAlign: 'center', padding: 20, color: 'var(--ink-soft)' }}>등록된 발주요청이 없습니다.</td></tr>
            ) : orders.slice().sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt)).map(o => (
              <tr key={o.id}>
                <td>{o.floor}</td>
                <td>{fmtZone(o.rowFrom, o.rowTo, o.colFrom, o.colTo)}</td>
                <td style={{ fontWeight: 600 }}>{o.itemName}</td>
                <td style={{ color: 'var(--ink-soft)' }}>{o.itemSpec}</td>
                <td style={{ color: 'var(--ink-soft)' }}>{o.itemColor}</td>
                <td className="mrs-mono">{o.qty} {o.unit}</td>
                <td>{o.createdBy}</td>
                <td className="mrs-mono" style={{ fontSize: 12 }}>{fmtDate(o.createdAt)}</td>
                <td><button style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4, color: '#B84B10' }} onClick={() => { if (confirm('이 발주요청을 삭제할까요?')) onDelete(o.id); }}><X size={15} /></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}


// ── 관리자 전체 화면 ────────────────────────────────────
function AdminApp({ session, requests, projects, users, orders, onUpdateStatus, onDelete, onSaveProjects, onAddUser, onUpdateUser, onDeleteUser, onAddOrder, onDeleteOrder, savingSettings }) {
  const [tab, setTab] = useState('today');
  return (
    <div className="mrs-body">
      <div className="mrs-tabs" style={{ margin: '-20px -20px 20px', padding: '0 20px' }}>
        <button className={`mrs-tab ${tab === 'today' ? 'active' : ''}`} onClick={() => setTab('today')}><CalendarDays size={16} /> 금일 자재요청</button>
        <button className={`mrs-tab ${tab === 'all' ? 'active' : ''}`} onClick={() => setTab('all')}><LayoutDashboard size={16} /> 누계 요청리스트</button>
        <button className={`mrs-tab ${tab === 'orders' ? 'active' : ''}`} onClick={() => setTab('orders')}><ClipboardList size={16} /> 발주요청</button>
        <button className={`mrs-tab ${tab === 'manage' ? 'active' : ''}`} onClick={() => setTab('manage')}><Users size={16} /> 관리 설정</button>
      </div>

      {tab === 'today' && <RequestsTable requests={requests} projects={projects} onUpdateStatus={onUpdateStatus} onDelete={onDelete} scope="today" />}
      {tab === 'all' && <RequestsTable requests={requests} projects={projects} onUpdateStatus={onUpdateStatus} onDelete={onDelete} scope="all" />}
      {tab === 'orders' && <OrderManager orders={orders} session={session} onAdd={onAddOrder} onDelete={onDeleteOrder} saving={savingSettings} />}
      {tab === 'manage' && (
        <div>
          <ProjectManager projects={projects} onSave={onSaveProjects} saving={savingSettings} />
          <StaffManager users={users} projects={projects} onAdd={onAddUser} onUpdate={onUpdateUser} onDelete={onDeleteUser} saving={savingSettings} />
          <AdminAccountManager session={session} onUpdate={onUpdateUser} saving={savingSettings} />
        </div>
      )}
    </div>
  );
}

// ── 자재팀: 현장 출고 대기 ──────────────────────────────
function MaterialOutbound({ requests, projects, onUpdateStatus }) {
  const projectNameById = {};
  projects.forEach(p => { projectNameById[p.id] = p.name; });

  const rows = [];
  requests.forEach(r => {
    r.items.forEach(it => {
      if (it.status === '입고완료') return;
      rows.push({
        reqId: r.id, itemId: it.id, requester: r.requester, projectName: projectNameById[r.projectId] || '-',
        zone: r.zone, name: it.name, spec: it.spec, color: it.color, qty: it.qty, unit: it.unit, status: it.status, createdAt: r.createdAt,
      });
    });
  });
  rows.sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));

  return (
    <div>
      {rows.length === 0 ? (
        <div className="mrs-card mrs-empty">출고 대기 중인 요청이 없습니다.</div>
      ) : (
        <div className="mrs-table-wrap">
          <table className="mrs-table">
            <thead><tr><th>요청자</th><th>프로젝트</th><th>구역</th><th>품목</th><th>규격</th><th>색상</th><th>수량</th><th>상태</th><th></th></tr></thead>
            <tbody>
              {rows.map(r => (
                <tr key={r.itemId}>
                  <td>{r.requester}</td><td>{r.projectName}</td><td>{r.zone}</td>
                  <td style={{ fontWeight: 600 }}>{r.name}</td><td style={{ color: 'var(--ink-soft)' }}>{r.spec}</td><td style={{ color: 'var(--ink-soft)' }}>{r.color}</td>
                  <td className="mrs-mono">{r.qty} {r.unit}</td>
                  <td><StatusBadge value={r.status} /></td>
                  <td><button className="mrs-btn mrs-btn-primary" style={{ padding: '6px 10px', fontSize: 12 }} onClick={() => onUpdateStatus(r.reqId, r.itemId, '입고완료')}>출고 처리</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ── 자재팀: 잔여물량 현황 (발주물량 - 현장입고 = 잔여물량) ──
function MaterialBalance({ orders, requests }) {
  const shippedItems = [];
  requests.forEach(r => {
    r.items.forEach(it => {
      if (it.status !== '입고완료') return;
      shippedItems.push({ name: it.name, spec: it.spec, color: it.color, qty: Number(it.qty) || 0, floor: r.floor, rowFrom: r.rowFrom, rowTo: r.rowTo, colFrom: r.colFrom, colTo: r.colTo, requester: r.requester, zone: r.zone });
    });
  });

  const rows = orders.map(o => {
    const matched = shippedItems.filter(it => sameItem(it, o) && zoneContained(it, o));
    const shipped = matched.reduce((sum, it) => sum + it.qty, 0);
    return { ...o, shipped, remain: (Number(o.qty) || 0) - shipped };
  });

  const unmatched = shippedItems.filter(it => !orders.some(o => sameItem(it, o) && zoneContained(it, o)));

  return (
    <div>
      <div className="mrs-table-wrap" style={{ marginBottom: 20 }}>
        <table className="mrs-table">
          <thead><tr><th>층</th><th>구간</th><th>품목</th><th>규격</th><th>색상</th><th>발주물량</th><th>현장입고</th><th>잔여물량</th></tr></thead>
          <tbody>
            {rows.length === 0 ? (
              <tr><td colSpan={8} style={{ textAlign: 'center', padding: 20, color: 'var(--ink-soft)' }}>등록된 발주요청이 없습니다.</td></tr>
            ) : rows.map(r => (
              <tr key={r.id}>
                <td>{r.floor}</td>
                <td>{fmtZone(r.rowFrom, r.rowTo, r.colFrom, r.colTo)}</td>
                <td style={{ fontWeight: 600 }}>{r.itemName}</td>
                <td style={{ color: 'var(--ink-soft)' }}>{r.itemSpec}</td>
                <td style={{ color: 'var(--ink-soft)' }}>{r.itemColor}</td>
                <td className="mrs-mono">{r.qty} {r.unit}</td>
                <td className="mrs-mono">{r.shipped} {r.unit}</td>
                <td className="mrs-mono" style={{ fontWeight: 700, color: r.remain < 0 ? '#B84B10' : '#2E6B47' }}>{r.remain} {r.unit}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <h3 className="mrs-display" style={{ fontSize: 14, margin: '0 0 10px', fontWeight: 600, color: 'var(--ink-soft)' }}>미지정 출고내역 (매칭되는 발주요청 없음)</h3>
      {unmatched.length === 0 ? (
        <div className="mrs-card mrs-empty" style={{ padding: 20 }}>없습니다.</div>
      ) : (
        <div className="mrs-table-wrap">
          <table className="mrs-table">
            <thead><tr><th>요청자</th><th>구역</th><th>품목</th><th>규격</th><th>색상</th><th>수량</th></tr></thead>
            <tbody>
              {unmatched.map((it, i) => (
                <tr key={i}>
                  <td>{it.requester}</td><td>{it.zone}</td>
                  <td style={{ fontWeight: 600 }}>{it.name}</td><td style={{ color: 'var(--ink-soft)' }}>{it.spec}</td><td style={{ color: 'var(--ink-soft)' }}>{it.color}</td>
                  <td className="mrs-mono">{it.qty}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ── 자재팀 전체 화면 ────────────────────────────────────
function MaterialApp({ requests, projects, orders, onUpdateStatus }) {
  const [tab, setTab] = useState('outbound');
  return (
    <div className="mrs-body">
      <div className="mrs-tabs" style={{ margin: '-20px -20px 20px', padding: '0 20px' }}>
        <button className={`mrs-tab ${tab === 'outbound' ? 'active' : ''}`} onClick={() => setTab('outbound')}><ClipboardList size={16} /> 현장 출고 대기</button>
        <button className={`mrs-tab ${tab === 'balance' ? 'active' : ''}`} onClick={() => setTab('balance')}><LayoutDashboard size={16} /> 잔여물량 현황</button>
      </div>
      {tab === 'outbound'
        ? <MaterialOutbound requests={requests} projects={projects} onUpdateStatus={onUpdateStatus} />
        : <MaterialBalance orders={orders} requests={requests} />}
    </div>
  );
}


export default function App() {
  const [session, setSession] = useState(null);
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginError, setLoginError] = useState(null);

  const [requests, setRequests] = useState([]);
  const [projects, setProjects] = useState(DEFAULT_PROJECTS);
  const [users, setUsers] = useState([]);
  const [orders, setOrders] = useState([]);
  const [dataLoading, setDataLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [savingSettings, setSavingSettings] = useState(false);
  const [error, setError] = useState(null);

  async function handleLogin(username, password) {
    setLoginLoading(true); setLoginError(null);
    try {
      const result = await apiPost('login', { username, password });
      if (result.error) { setLoginError(result.error); return; }
      setSession(result);
      await loadData(result.role);
    } catch (e) {
      setLoginError('로그인 중 오류가 발생했습니다. 네트워크를 확인해주세요.');
    } finally {
      setLoginLoading(false);
    }
  }

  async function loadData(role) {
    setDataLoading(true); setError(null);
    try {
      const calls = [apiGet('list'), apiGet('projects')];
      if (role === 'admin') calls.push(apiGet('users'));
      if (role === 'admin' || role === 'material') calls.push(apiGet('orders'));
      const results = await Promise.all(calls);
      setRequests(Array.isArray(results[0]) ? results[0] : []);
      setProjects(Array.isArray(results[1]) && results[1].length ? results[1] : DEFAULT_PROJECTS);
      let idx = 2;
      if (role === 'admin') { setUsers(Array.isArray(results[idx]) ? results[idx] : []); idx++; }
      if (role === 'admin' || role === 'material') { setOrders(Array.isArray(results[idx]) ? results[idx] : []); }
    } catch (e) {
      setError('데이터를 불러오지 못했습니다. 네트워크를 확인해주세요.');
    } finally {
      setDataLoading(false);
    }
  }

  async function handleSubmit(payload) {
    const next = [payload, ...requests];
    setRequests(next);
    setSaving(true); setError(null);
    try { await apiPost('addRequest', { payload }); }
    catch (e) { setError('요청 저장에 실패했습니다.'); await loadData(session.role); }
    finally { setSaving(false); }
  }

  async function handleUpdateStatus(reqId, itemId, status) {
    const next = requests.map(r => r.id !== reqId ? r : { ...r, items: r.items.map(it => it.id !== itemId ? it : { ...it, status }) });
    setRequests(next);
    setSaving(true); setError(null);
    try { await apiPost('updateStatus', { itemId, status }); }
    catch (e) { setError('상태 변경에 실패했습니다.'); await loadData(session.role); }
    finally { setSaving(false); }
  }

  async function handleDelete(reqId) {
    const next = requests.filter(r => r.id !== reqId);
    setRequests(next);
    setSaving(true); setError(null);
    try { await apiPost('deleteRequest', { reqId }); }
    catch (e) { setError('삭제에 실패했습니다.'); await loadData(session.role); }
    finally { setSaving(false); }
  }

  async function handleSaveProjects(next) {
    setSavingSettings(true); setError(null);
    try { await apiPost('saveProjects', { projects: next }); setProjects(next); }
    catch (e) { setError('프로젝트 저장에 실패했습니다.'); }
    finally { setSavingSettings(false); }
  }

  async function handleAddUser(user) {
    setSavingSettings(true); setError(null);
    try {
      const res = await apiPost('addUser', { user });
      if (res.error) { alert(res.error); return; }
      await loadData('admin');
    } catch (e) { setError('팀장 추가에 실패했습니다.'); }
    finally { setSavingSettings(false); }
  }

  async function handleUpdateUser(user) {
    setSavingSettings(true); setError(null);
    try {
      const res = await apiPost('updateUser', { user });
      if (res.error) { alert(res.error); return; }
      if (session.id === user.id) setSession({ ...session, username: user.username || session.username });
      await loadData('admin');
    } catch (e) { setError('계정 수정에 실패했습니다.'); }
    finally { setSavingSettings(false); }
  }

  async function handleDeleteUser(id) {
    setSavingSettings(true); setError(null);
    try {
      const res = await apiPost('deleteUser', { id });
      if (res.error) { alert(res.error); return; }
      await loadData('admin');
    } catch (e) { setError('계정 삭제에 실패했습니다.'); }
    finally { setSavingSettings(false); }
  }

  async function handleAddOrder(order) {
    setSavingSettings(true); setError(null);
    try {
      const res = await apiPost('addOrder', { order });
      if (res.error) { alert(res.error); return; }
      await loadData(session.role);
    } catch (e) { setError('발주요청 등록에 실패했습니다.'); }
    finally { setSavingSettings(false); }
  }

  async function handleDeleteOrder(id) {
    setSavingSettings(true); setError(null);
    try {
      const res = await apiPost('deleteOrder', { id });
      if (res.error) { alert(res.error); return; }
      await loadData(session.role);
    } catch (e) { setError('발주요청 삭제에 실패했습니다.'); }
    finally { setSavingSettings(false); }
  }

  if (!session) {
    return <LoginScreen onLogin={handleLogin} loading={loginLoading} error={loginError} />;
  }

  const currentProjectName = session.role === 'leader'
    ? (projects.find(p => p.id === session.projectId) || {}).name || '(알 수 없음)'
    : null;

  return (
    <div className="mrs-root">
      <GlobalStyle />
      <div className="mrs-header">
        <div>
          <div className="mrs-display" style={{ fontSize: 21, fontWeight: 600, letterSpacing: '0.03em' }}>자재 요청 관리 시스템</div>
          <div style={{ fontSize: 12, color: '#9AA5AC', marginTop: 2 }}>Material Requisition System · 건국이엔아이 공무팀</div>
        </div>
        <div className="mrs-header-right">
          {session.role === 'leader' && <div className="mrs-project-chip">PROJECT&nbsp;<b>{currentProjectName}</b></div>}
          <span className="mrs-role-badge" style={{ background: session.role === 'admin' ? '#FBE7DA' : session.role === 'material' ? '#E1EBE3' : '#E3ECF5', color: session.role === 'admin' ? '#B84B10' : session.role === 'material' ? '#2E6B47' : '#2B5A8C' }}>
            {session.role === 'admin' ? '관리자' : session.role === 'material' ? '자재팀' : '팀장'}
          </span>
          <span className="mrs-user-chip">{session.name}</span>
          <button className="mrs-logout-btn" onClick={() => { setSession(null); setRequests([]); setUsers([]); setOrders([]); }}><LogOut size={13} /> 로그아웃</button>
        </div>
      </div>

      {error && (
        <div className="mrs-card" style={{ margin: '16px 20px 0', padding: '10px 14px', borderColor: '#F0A97A', background: '#FBE7DA', display: 'flex', alignItems: 'center', gap: 8, color: '#B84B10', fontSize: 13 }}>
          <AlertCircle size={16} /> {error}
        </div>
      )}

      {dataLoading ? (
        <div className="mrs-empty"><Loader2 size={20} className="mrs-spin" /><div style={{ marginTop: 8 }}>불러오는 중...</div></div>
      ) : session.role === 'admin' ? (
        <AdminApp
          session={session} requests={requests} projects={projects} users={users} orders={orders}
          onUpdateStatus={handleUpdateStatus} onDelete={handleDelete} onSaveProjects={handleSaveProjects}
          onAddUser={handleAddUser} onUpdateUser={handleUpdateUser} onDeleteUser={handleDeleteUser}
          onAddOrder={handleAddOrder} onDeleteOrder={handleDeleteOrder}
          savingSettings={savingSettings}
        />
      ) : session.role === 'material' ? (
        <MaterialApp requests={requests} projects={projects} orders={orders} onUpdateStatus={handleUpdateStatus} />
      ) : (
        <LeaderApp session={session} requests={requests} projects={projects} onSubmit={handleSubmit} saving={saving} />
      )}
    </div>
  );
}

