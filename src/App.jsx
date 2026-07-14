import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import { NanumGothicBase64 } from './NanumGothicFont';
import { Plus, Trash2, Download, ClipboardList, LayoutDashboard, Loader2, AlertCircle, CheckCircle2, X, Save, LogOut, Building2, Users, ListChecks, CalendarDays, Lock, PenLine, PackageMinus, RefreshCw, MapPin, Package } from 'lucide-react';

const UNITS = ['EA', 'M', 'SET', 'BOX', 'ROLL', 'KG', '본', '롤'];
const PROCESSES = ['자탐', '유도등', '무통'];
const STATUS_FLOW = ['요청됨', '확인됨', '입고완료'];
const RETURN_STATUS_FLOW = ['반출요청', '반출확인완료'];
const STATUS_COLOR = {
  '요청됨':       { bg: '#EEF0EC', fg: '#5C6B73', bd: '#C9CFC6' },
  '확인됨':       { bg: '#E3ECF5', fg: '#2B5A8C', bd: '#9FBEDC' },
  '입고완료':     { bg: '#E1EBE3', fg: '#2E6B47', bd: '#9FC7AC' },
  '반출요청':     { bg: '#FBE7DA', fg: '#B84B10', bd: '#F0A97A' },
  '반출확인완료': { bg: '#E1EBE3', fg: '#2E6B47', bd: '#9FC7AC' },
};
const FALLBACK_STATUS_COLOR = { bg: '#EEF0EC', fg: '#5C6B73', bd: '#C9CFC6' };

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
function genId(prefix = '') { return (prefix ? prefix + '-' : '') + Math.random().toString(36).slice(2, 10); }
function fmtDate(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return String(iso);
  return `${d.getFullYear()}.${pad(d.getMonth() + 1)}.${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function isToday(iso) {
  const d = new Date(iso);
  const now = new Date();
  return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth() && d.getDate() === now.getDate();
}

// ── 카탈로그 유틸 ─────────────────────────────────────────
function catalogNames(catalog) { return [...new Set(catalog.map(c => c.name))]; }
function catalogSpecs(catalog, name) { return [...new Set(catalog.filter(c => c.name === name).map(c => c.spec))]; }
function catalogColors(catalog, name, spec) { return [...new Set(catalog.filter(c => c.name === name && c.spec === spec).map(c => c.color))]; }
function catalogUnit(catalog, name, spec, color) {
  const found = catalog.find(c => c.name === name && c.spec === spec && c.color === color);
  return found ? found.unit : 'EA';
}
function newItemRow(catalog) {
  if (catalog.length === 0) return { id: genId(), name: '', spec: '', color: 'N/A', qty: '', unit: 'EA' };
  const name = catalogNames(catalog)[0];
  const spec = catalogSpecs(catalog, name)[0];
  const color = catalogColors(catalog, name, spec)[0];
  const unit = catalogUnit(catalog, name, spec, color);
  return { id: genId(), name, spec, color, qty: '', unit };
}

// ── 글로벌 CSS ──────────────────────────────────────────
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
    .mrs-logout-btn, .mrs-refresh-btn { display: flex; align-items: center; gap: 6px; background: none; border: 1px solid #3A4954; color: #C9CFC6; border-radius: 20px; padding: 6px 12px; font-size: 12px; cursor: pointer; }
    .mrs-logout-btn:hover, .mrs-refresh-btn:hover { border-color: var(--accent); color: #EDEAE0; }
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
    .mrs-zone-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px; }
    @media (max-width: 480px) { .mrs-zone-grid { gap: 6px; } .mrs-zone-grid select { font-size: 12px; padding: 7px 4px; } }
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

function StatusBadge({ value }) {
  const c = STATUS_COLOR[value] || FALLBACK_STATUS_COLOR;
  return <span className="mrs-chip" style={{ background: c.bg, color: c.fg, borderColor: c.bd }}>{value || '알 수 없음'}</span>;
}

function StatusSelect({ value, options, onChange }) {
  const c = STATUS_COLOR[value] || FALLBACK_STATUS_COLOR;
  const safe = options.includes(value) ? value : options[0];
  return (
    <select className="mrs-status-select" value={safe} onChange={e => onChange(e.target.value)} style={{ background: c.bg, color: c.fg, borderColor: c.bd }}>
      {options.map(s => <option key={s} value={s}>{s}</option>)}
    </select>
  );
}

// ── 서명패드 ─────────────────────────────────────────────
function SignaturePad({ onChange }) {
  const canvasRef = useRef(null);
  const drawing = useRef(false);
  function getPos(e, canvas) {
    const rect = canvas.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  }
  function start(e) {
    drawing.current = true;
    const ctx = canvasRef.current.getContext('2d');
    const { x, y } = getPos(e, canvasRef.current);
    ctx.beginPath(); ctx.moveTo(x, y);
  }
  function move(e) {
    if (!drawing.current) return;
    const ctx = canvasRef.current.getContext('2d');
    const { x, y } = getPos(e, canvasRef.current);
    ctx.lineTo(x, y); ctx.strokeStyle = '#1C2A33'; ctx.lineWidth = 2.2; ctx.lineCap = 'round'; ctx.stroke();
  }
  function end() {
    if (!drawing.current) return;
    drawing.current = false;
    onChange(canvasRef.current.toDataURL('image/png'));
  }
  function clear() {
    const canvas = canvasRef.current;
    canvas.getContext('2d').clearRect(0, 0, canvas.width, canvas.height);
    onChange('');
  }
  return (
    <div>
      <canvas
        ref={canvasRef} width={280} height={120}
        style={{ width: '100%', maxWidth: 280, height: 120, background: '#fff', border: '1px solid var(--line)', borderRadius: 4, touchAction: 'none', cursor: 'crosshair' }}
        onPointerDown={start} onPointerMove={move} onPointerUp={end} onPointerLeave={end}
      />
      <button type="button" className="mrs-btn mrs-btn-ghost" style={{ marginTop: 6, padding: '4px 10px', fontSize: 12 }} onClick={clear}>지우고 다시 서명</button>
    </div>
  );
}

function ConfirmSignaturePanel({ deliverLabel, receiveLabel, deliverNameDefault, receiveNameDefault, onCancel, onSubmit, saving }) {
  const [deliverName, setDeliverName] = useState(deliverNameDefault || '');
  const [receiveName, setReceiveName] = useState(receiveNameDefault || '');
  const [deliverSig, setDeliverSig] = useState('');
  const [receiveSig, setReceiveSig] = useState('');
  function submit() {
    if (!deliverName.trim() || !receiveName.trim()) { alert('인도자, 인수자 이름을 모두 입력해주세요.'); return; }
    if (!deliverSig || !receiveSig) { alert('양쪽 서명을 모두 받아주세요.'); return; }
    onSubmit({ deliverName: deliverName.trim(), deliverSignature: deliverSig, receiveName: receiveName.trim(), receiveSignature: receiveSig });
  }
  return (
    <div className="mrs-card" style={{ padding: 16, marginTop: 10, background: '#F1EFE9' }}>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
        <div>
          <label className="mrs-field-label">{deliverLabel}</label>
          <input className="mrs-input" value={deliverName} onChange={e => setDeliverName(e.target.value)} placeholder="이름" style={{ marginBottom: 8 }} />
          <SignaturePad onChange={setDeliverSig} />
        </div>
        <div>
          <label className="mrs-field-label">{receiveLabel}</label>
          <input className="mrs-input" value={receiveName} onChange={e => setReceiveName(e.target.value)} placeholder="이름" style={{ marginBottom: 8 }} />
          <SignaturePad onChange={setReceiveSig} />
        </div>
      </div>
      <div style={{ display: 'flex', gap: 8, marginTop: 14 }}>
        <button className="mrs-btn mrs-btn-primary" onClick={submit} disabled={saving}><PenLine size={15} /> 서명 완료 및 확인</button>
        <button className="mrs-btn mrs-btn-ghost" onClick={onCancel}>취소</button>
      </div>
    </div>
  );
}

// ── PDF 생성 (한글 폰트 임베드) ────────────────────────────
let pdfFontLoaded = false;
function ensurePdfFont(doc) {
  if (!pdfFontLoaded) {
    doc.addFileToVFS('NanumGothic.ttf', NanumGothicBase64);
    doc.addFont('NanumGothic.ttf', 'NanumGothic', 'normal');
    pdfFontLoaded = true;
  }
  doc.setFont('NanumGothic');
}

function generateDocPdf({ title, docNo, dateStr, projectName, zoneStr, items, deliverLabel, deliverName, deliverSignature, receiveLabel, receiveName, receiveSignature }) {
  const doc = new jsPDF({ unit: 'mm', format: 'a5' });
  // 한글 폰트를 매 인스턴스마다 등록
  doc.addFileToVFS('NanumGothic.ttf', NanumGothicBase64);
  doc.addFont('NanumGothic.ttf', 'NanumGothic', 'normal');
  doc.setFont('NanumGothic');

  const pw = doc.internal.pageSize.getWidth();
  let y = 16;

  doc.setFontSize(16);
  doc.text(title, pw / 2, y, { align: 'center' });
  y += 6;
  doc.setFontSize(9);
  doc.setTextColor(120);
  doc.text(docNo, pw / 2, y, { align: 'center' });
  doc.setTextColor(0);
  y += 10;

  doc.setFontSize(10);
  doc.text(`일자: ${dateStr}`, 14, y);
  doc.text(`프로젝트: ${projectName || '-'}`, pw - 14, y, { align: 'right' });
  y += 6;
  doc.text(`구역: ${zoneStr || '-'}`, 14, y);
  y += 8;

  doc.setDrawColor(200);
  doc.line(14, y, pw - 14, y);
  y += 6;
  doc.setFontSize(9);
  doc.text('품목', 14, y);
  doc.text('규격', 55, y);
  doc.text('색상', 85, y);
  doc.text('수량', pw - 14, y, { align: 'right' });
  y += 4;
  doc.line(14, y, pw - 14, y);
  y += 6;
  items.forEach(it => {
    doc.text(String(it.name || ''), 14, y);
    doc.text(String(it.spec || '-'), 55, y);
    doc.text(String(it.color || '-'), 85, y);
    doc.text(`${it.qty} ${it.unit}`, pw - 14, y, { align: 'right' });
    y += 6;
  });
  y += 6;

  const boxW = (pw - 14 * 2 - 8) / 2;
  const boxH = 40;
  doc.setDrawColor(150);
  doc.rect(14, y, boxW, boxH);
  doc.rect(14 + boxW + 8, y, boxW, boxH);
  doc.setFontSize(9);
  doc.text(deliverLabel, 14 + boxW / 2, y + 6, { align: 'center' });
  doc.text(receiveLabel, 14 + boxW + 8 + boxW / 2, y + 6, { align: 'center' });
  if (deliverSignature) { try { doc.addImage(deliverSignature, 'PNG', 16, y + 9, boxW - 4, boxH - 18); } catch (e) {} }
  if (receiveSignature) { try { doc.addImage(receiveSignature, 'PNG', 14 + boxW + 10, y + 9, boxW - 4, boxH - 18); } catch (e) {} }
  doc.setFontSize(8);
  doc.text(deliverName || '', 14 + boxW / 2, y + boxH - 3, { align: 'center' });
  doc.text(receiveName || '', 14 + boxW + 8 + boxW / 2, y + boxH - 3, { align: 'center' });

  doc.save(`${title}_${docNo}.pdf`);
}

// ── 품목 행 편집기 (카탈로그 기반) ────────────────────────
function ItemRowEditor({ item, catalog, onChange, onRemove, removable }) {
  const names = catalogNames(catalog);
  const specs = catalogSpecs(catalog, item.name);
  const colors = catalogColors(catalog, item.name, item.spec);

  function handleNameChange(name) {
    const spec = catalogSpecs(catalog, name)[0] || '';
    const color = catalogColors(catalog, name, spec)[0] || 'N/A';
    const unit = catalogUnit(catalog, name, spec, color);
    onChange({ ...item, name, spec, color, unit });
  }
  function handleSpecChange(spec) {
    const color = catalogColors(catalog, item.name, spec)[0] || 'N/A';
    const unit = catalogUnit(catalog, item.name, spec, color);
    onChange({ ...item, spec, color, unit });
  }
  function handleColorChange(color) {
    const unit = catalogUnit(catalog, item.name, item.spec, color);
    onChange({ ...item, color, unit });
  }

  return (
    <div className="mrs-item-row">
      <div><select className="mrs-select" value={item.name} onChange={e => handleNameChange(e.target.value)}>{names.length === 0 ? <option value="">품목 없음</option> : names.map(n => <option key={n} value={n}>{n}</option>)}</select></div>
      <div><select className="mrs-select" value={item.spec} onChange={e => handleSpecChange(e.target.value)}>{specs.map(s => <option key={s} value={s}>{s}</option>)}</select></div>
      <div><input className="mrs-input" type="number" min="0" value={item.qty} onChange={e => onChange({ ...item, qty: e.target.value })} placeholder="0" /></div>
      <div><select className="mrs-select" value={item.unit} onChange={e => onChange({ ...item, unit: e.target.value })}>{UNITS.map(u => <option key={u} value={u}>{u}</option>)}</select></div>
      <div><select className="mrs-select" value={item.color} onChange={e => handleColorChange(e.target.value)} disabled={colors.length <= 1}>{colors.map(c => <option key={c} value={c}>{c}</option>)}</select></div>
      <button className="mrs-btn mrs-btn-danger" onClick={onRemove} disabled={!removable} title="삭제" style={{ padding: 8 }}><Trash2 size={16} /></button>
    </div>
  );
}

// ── 로그인 화면 ─────────────────────────────────────────
function LoginScreen({ onLogin, loading, error }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  function submit() {
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
          <input className="mrs-input" value={username} onChange={e => setUsername(e.target.value)} onKeyDown={e => e.key === 'Enter' && submit()} placeholder="아이디" style={{ marginBottom: 14 }} />
          <label className="mrs-field-label">비밀번호</label>
          <input className="mrs-input" type="password" value={password} onChange={e => setPassword(e.target.value)} onKeyDown={e => e.key === 'Enter' && submit()} placeholder="비밀번호" style={{ marginBottom: 14 }} />
          {error && <div style={{ display: 'flex', alignItems: 'center', gap: 6, color: '#B84B10', fontSize: 12, marginBottom: 14 }}><AlertCircle size={14} /> {error}</div>}
          <button className="mrs-btn mrs-btn-primary" style={{ width: '100%', justifyContent: 'center' }} onClick={submit} disabled={loading}>
            {loading ? <Loader2 size={15} className="mrs-spin" /> : <Lock size={15} />} 로그인
          </button>
        </div>
      </div>
    </div>
  );
}

// ── 팀장: 자재요청 시트 ─────────────────────────────────
function RequestForm({ session, projectName, catalog, zones, onSubmit, saving }) {
  const myZones = zones.filter(z => z.projectId === session.projectId);
  const floors = [...new Set(myZones.map(z => z.floor).filter(Boolean))];
  const [floor, setFloor] = useState('');
  const rooms = [...new Set(myZones.filter(z => z.floor === floor).map(z => z.room).filter(Boolean))];
  const [room, setRoom] = useState('');
  const zoneNames = [...new Set(myZones.filter(z => z.floor === floor && z.room === room).map(z => z.zone).filter(Boolean))];
  const [zone, setZone] = useState('');

  const [process, setProcess] = useState(PROCESSES[0]);
  const [note, setNote] = useState('');
  const [items, setItems] = useState([newItemRow(catalog)]);
  const [justSubmitted, setJustSubmitted] = useState(false);

  useEffect(() => { if (!floor && floors.length > 0) setFloor(floors[0]); }, [floors.join('|')]);
  useEffect(() => { if (rooms.length > 0 && !rooms.includes(room)) setRoom(rooms[0]); }, [rooms.join('|')]);
  useEffect(() => { if (zoneNames.length > 0 && !zoneNames.includes(zone)) setZone(zoneNames[0]); }, [zoneNames.join('|')]);
  useEffect(() => { setItems([newItemRow(catalog)]); }, [catalog.length]);

  function updateItem(id, updated) { setItems(items.map(it => it.id === id ? updated : it)); }
  function addItem() { setItems([...items, newItemRow(catalog)]); }
  function removeItem(id) { if (items.length === 1) return; setItems(items.filter(it => it.id !== id)); }

  async function handleSubmit() {
    // 드롭다운에는 첫 옵션이 표시되고 있지만 state가 아직 빈 경우가 있을 수 있어 실제 사용 값을 여기서 확정
    const effFloor = floor || floors[0] || '';
    const effRoom = room || rooms[0] || '';
    const effZone = zone || zoneNames[0] || '';
    if (!effFloor || !effRoom || !effZone) { alert('구역을 모두 선택해주세요. 관리자에게 구역 등록을 요청하세요.'); return; }
    if (catalog.length === 0) { alert('품목이 등록되어 있지 않습니다.'); return; }
    const valid = items.filter(it => it.name && it.qty !== '' && Number(it.qty) > 0);
    if (valid.length === 0) { alert('최소 1개 이상의 품목에 수량을 입력해주세요.'); return; }
    const zoneName = `${effFloor} · ${effRoom} · ${effZone}`;
    const payload = {
      id: genId(), reqNo: genReqNo(), requester: session.name, username: session.username,
      projectId: session.projectId, zoneName, process, note: note.trim(), createdAt: new Date().toISOString(),
      items: valid.map(it => ({ id: genId(), name: it.name, spec: it.spec, color: it.color, qty: it.qty, unit: it.unit, status: '요청됨' })),
    };
    await onSubmit(payload);
    setNote(''); setItems([newItemRow(catalog)]);
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

      <label className="mrs-field-label">구역 (층 / 실 / 구역)</label>
      {myZones.length === 0 ? (
        <div style={{ padding: 10, background: '#FBE7DA', color: '#B84B10', fontSize: 13, borderRadius: 3, marginBottom: 14 }}>
          이 프로젝트에 등록된 구역이 없습니다. 관리자에게 구역 등록을 요청하세요.
        </div>
      ) : (
        <div className="mrs-zone-grid" style={{ marginBottom: 14 }}>
          <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>층</span>
            <select className="mrs-select" value={floor} onChange={e => { setFloor(e.target.value); setRoom(''); setZone(''); }}>
              {floors.map(f => <option key={f} value={f}>{f}</option>)}
            </select>
          </div>
          <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>실</span>
            <select className="mrs-select" value={room} onChange={e => { setRoom(e.target.value); setZone(''); }}>
              {rooms.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
          <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>구역</span>
            <select className="mrs-select" value={zone} onChange={e => setZone(e.target.value)}>
              {zoneNames.map(z => <option key={z} value={z}>{z}</option>)}
            </select>
          </div>
        </div>
      )}

      <div style={{ marginBottom: 8 }}><label className="mrs-field-label">요청 품목</label></div>
      {catalog.length === 0 ? (
        <div style={{ padding: 10, background: '#FBE7DA', color: '#B84B10', fontSize: 13, borderRadius: 3, marginBottom: 10 }}>
          등록된 품목이 없습니다.
        </div>
      ) : (
        <>
          <div className="mrs-item-row" style={{ borderBottom: 'none', paddingBottom: 0 }}>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>품목</span>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>규격</span>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>수량</span>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>단위</span>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>색상</span>
            <span></span>
          </div>
          {items.map(it => <ItemRowEditor key={it.id} item={it} catalog={catalog} onChange={u => updateItem(it.id, u)} onRemove={() => removeItem(it.id)} removable={items.length > 1} />)}
          <button className="mrs-btn mrs-btn-ghost" style={{ marginTop: 10 }} onClick={addItem}><Plus size={15} /> 품목 추가</button>
        </>
      )}

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

// ── 팀장: 요청리스트 (본인 요청 조회 + 삭제 + 입고확인 + PDF) ──
function MyRequestList({ requests, session, projects, onConfirmReceipt, onDelete, saving }) {
  const [confirmingReqId, setConfirmingReqId] = useState(null);
  const projectNameById = {};
  projects.forEach(p => { projectNameById[p.id] = p.name; });

  const mine = requests
    .filter(r => (r.username ? r.username === session.username : r.requester === session.name))
    .slice()
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  function allRequested(r) { return r.items.every(it => it.status === '요청됨'); }
  function allConfirmedOrDelivered(r) { return r.items.every(it => it.status === '확인됨' || it.status === '입고완료'); }
  function allDelivered(r) { return r.items.every(it => it.status === '입고완료'); }

  async function handleConfirm(r, sig) {
    await onConfirmReceipt(r.id, sig);
    setConfirmingReqId(null);
  }

  function downloadPdf(r) {
    generateDocPdf({
      title: '거래명세표', docNo: r.reqNo, dateStr: fmtDate(r.confirmedAt || r.createdAt),
      projectName: projectNameById[r.projectId], zoneStr: r.zoneName,
      items: r.items, deliverLabel: '인도자 (자재팀)', deliverName: r.deliverName, deliverSignature: r.deliverSignature,
      receiveLabel: '인수자 (현장팀장)', receiveName: r.receiveName, receiveSignature: r.receiveSignature,
    });
  }

  return (
    <div>
      <h2 className="mrs-display" style={{ fontSize: 18, margin: '0 0 14px', fontWeight: 600 }}>요청리스트</h2>
      {mine.length === 0 ? (
        <div className="mrs-card mrs-empty">제출한 요청이 없습니다.</div>
      ) : mine.map(r => (
        <div className="mrs-card" key={r.id} style={{ padding: 16, marginBottom: 12 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8, flexWrap: 'wrap', gap: 6 }}>
            <span className="mrs-mono" style={{ fontSize: 12, color: 'var(--ink-soft)' }}>{r.reqNo}</span>
            <span className="mrs-mono" style={{ fontSize: 12, color: 'var(--ink-soft)' }}>{fmtDate(r.createdAt)}</span>
          </div>
          <div style={{ fontSize: 13, color: 'var(--ink-soft)', marginBottom: 10 }}>{r.zoneName} · {r.process}</div>
          {r.items.map(it => (
            <div key={it.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '5px 0', borderBottom: '1px dashed var(--paper-line)' }}>
              <span style={{ fontSize: 13 }}><b>{it.name}</b> {it.spec}{it.color && it.color !== 'N/A' ? ` ${it.color}` : ''} · {it.qty}{it.unit}</span>
              <StatusBadge value={it.status} />
            </div>
          ))}

          <div style={{ display: 'flex', gap: 8, marginTop: 12, flexWrap: 'wrap' }}>
            {allRequested(r) && (
              <button className="mrs-btn mrs-btn-danger" style={{ borderColor: '#F0A97A' }} onClick={() => { if (confirm('이 요청을 삭제할까요?')) onDelete(r.id); }}>
                <Trash2 size={15} /> 요청 삭제
              </button>
            )}
            {allDelivered(r) ? (
              <button className="mrs-btn mrs-btn-ghost" onClick={() => downloadPdf(r)}><Download size={15} /> 거래명세표 다운로드</button>
            ) : allConfirmedOrDelivered(r) && confirmingReqId !== r.id ? (
              <button className="mrs-btn mrs-btn-primary" onClick={() => setConfirmingReqId(r.id)}><PenLine size={15} /> 입고확인</button>
            ) : null}
          </div>

          {confirmingReqId === r.id && (
            <ConfirmSignaturePanel
              deliverLabel="인도자 (자재팀)" receiveLabel="인수자 (현장팀장)"
              receiveNameDefault={session.name}
              onCancel={() => setConfirmingReqId(null)}
              onSubmit={sig => handleConfirm(r, sig)}
              saving={saving}
            />
          )}
        </div>
      ))}
    </div>
  );
}

// ── 팀장: 물량반출 ──────────────────────────────────────
function ReturnPanel({ session, projectName, catalog, zones, returns, onSubmit, saving }) {
  const myZones = zones.filter(z => z.projectId === session.projectId);
  const floors = [...new Set(myZones.map(z => z.floor).filter(Boolean))];
  const [floor, setFloor] = useState('');
  const rooms = [...new Set(myZones.filter(z => z.floor === floor).map(z => z.room).filter(Boolean))];
  const [room, setRoom] = useState('');
  const zoneNames = [...new Set(myZones.filter(z => z.floor === floor && z.room === room).map(z => z.zone).filter(Boolean))];
  const [zone, setZone] = useState('');

  const [reason, setReason] = useState('');
  const [items, setItems] = useState([newItemRow(catalog)]);
  const [justSubmitted, setJustSubmitted] = useState(false);

  useEffect(() => { if (!floor && floors.length > 0) setFloor(floors[0]); }, [floors.join('|')]);
  useEffect(() => { if (rooms.length > 0 && !rooms.includes(room)) setRoom(rooms[0]); }, [rooms.join('|')]);
  useEffect(() => { if (zoneNames.length > 0 && !zoneNames.includes(zone)) setZone(zoneNames[0]); }, [zoneNames.join('|')]);
  useEffect(() => { setItems([newItemRow(catalog)]); }, [catalog.length]);

  function updateItem(id, updated) { setItems(items.map(it => it.id === id ? updated : it)); }
  function addItem() { setItems([...items, newItemRow(catalog)]); }
  function removeItem(id) { if (items.length === 1) return; setItems(items.filter(it => it.id !== id)); }

  async function handleSubmit() {
    const effFloor = floor || floors[0] || '';
    const effRoom = room || rooms[0] || '';
    const effZone = zone || zoneNames[0] || '';
    if (!effFloor || !effRoom || !effZone) { alert('구역을 모두 선택해주세요.'); return; }
    if (catalog.length === 0) { alert('품목이 등록되어 있지 않습니다.'); return; }
    const valid = items.filter(it => it.name && it.qty !== '' && Number(it.qty) > 0);
    if (valid.length === 0) { alert('최소 1개 이상의 품목에 수량을 입력해주세요.'); return; }
    const zoneName = `${effFloor} · ${effRoom} · ${effZone}`;
    const payload = {
      id: genId(), reqNo: genReqNo(), requester: session.name, username: session.username,
      projectId: session.projectId, zoneName, reason: reason.trim(), createdAt: new Date().toISOString(),
      items: valid.map(it => ({ id: genId(), name: it.name, spec: it.spec, color: it.color, qty: it.qty, unit: it.unit })),
    };
    await onSubmit(payload);
    setReason(''); setItems([newItemRow(catalog)]);
    setJustSubmitted(true);
    setTimeout(() => setJustSubmitted(false), 3000);
  }

  const mine = returns
    .filter(r => (r.username ? r.username === session.username : r.requester === session.name))
    .slice().sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  function downloadPdf(r) {
    generateDocPdf({
      title: '반출확인서', docNo: r.reqNo, dateStr: fmtDate(r.confirmedAt || r.createdAt),
      projectName, zoneStr: r.zoneName,
      items: r.items, deliverLabel: '인도자 (현장팀장)', deliverName: r.deliverName, deliverSignature: r.deliverSignature,
      receiveLabel: '인수자 (자재팀)', receiveName: r.receiveName, receiveSignature: r.receiveSignature,
    });
  }

  return (
    <div>
      <div className="mrs-card" style={{ padding: 22, marginBottom: 20 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 18 }}>
          <h2 className="mrs-display" style={{ fontSize: 18, margin: 0, fontWeight: 600 }}>물량반출 요청</h2>
          <span className="mrs-mono" style={{ fontSize: 11, color: 'var(--ink-soft)' }}>RETURN REQUEST</span>
        </div>

        <label className="mrs-field-label">구역 (층 / 실 / 구역)</label>
        {myZones.length === 0 ? (
          <div style={{ padding: 10, background: '#FBE7DA', color: '#B84B10', fontSize: 13, borderRadius: 3, marginBottom: 14 }}>
            이 프로젝트에 등록된 구역이 없습니다.
          </div>
        ) : (
          <div className="mrs-zone-grid" style={{ marginBottom: 14 }}>
            <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>층</span>
              <select className="mrs-select" value={floor} onChange={e => { setFloor(e.target.value); setRoom(''); setZone(''); }}>
                {floors.map(f => <option key={f} value={f}>{f}</option>)}
              </select>
            </div>
            <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>실</span>
              <select className="mrs-select" value={room} onChange={e => { setRoom(e.target.value); setZone(''); }}>
                {rooms.map(r => <option key={r} value={r}>{r}</option>)}
              </select>
            </div>
            <div><span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>구역</span>
              <select className="mrs-select" value={zone} onChange={e => setZone(e.target.value)}>
                {zoneNames.map(z => <option key={z} value={z}>{z}</option>)}
              </select>
            </div>
          </div>
        )}

        <div style={{ marginBottom: 8 }}><label className="mrs-field-label">반출 품목</label></div>
        {catalog.length === 0 ? (
          <div style={{ padding: 10, background: '#FBE7DA', color: '#B84B10', fontSize: 13, borderRadius: 3, marginBottom: 10 }}>
            등록된 품목이 없습니다.
          </div>
        ) : (
          <>
            <div className="mrs-item-row" style={{ borderBottom: 'none', paddingBottom: 0 }}>
              <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>품목</span>
              <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>규격</span>
              <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>수량</span>
              <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>단위</span>
              <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>색상</span>
              <span></span>
            </div>
            {items.map(it => <ItemRowEditor key={it.id} item={it} catalog={catalog} onChange={u => updateItem(it.id, u)} onRemove={() => removeItem(it.id)} removable={items.length > 1} />)}
            <button className="mrs-btn mrs-btn-ghost" style={{ marginTop: 10 }} onClick={addItem}><Plus size={15} /> 품목 추가</button>
          </>
        )}

        <div style={{ marginTop: 18 }}>
          <label className="mrs-field-label">반출 사유 (선택)</label>
          <input className="mrs-input" value={reason} onChange={e => setReason(e.target.value)} placeholder="" />
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 22 }}>
          <button className="mrs-btn mrs-btn-primary" onClick={handleSubmit} disabled={saving}>
            {saving ? <Loader2 size={15} className="mrs-spin" /> : <PackageMinus size={15} />} 반출 요청 제출
          </button>
          {justSubmitted && <span style={{ display: 'flex', alignItems: 'center', gap: 6, color: '#2E6B47', fontSize: 13, fontWeight: 600 }}><CheckCircle2 size={16} /> 요청이 접수되었습니다</span>}
        </div>
      </div>

      <h3 className="mrs-display" style={{ fontSize: 15, margin: '0 0 12px', fontWeight: 600 }}>반출내역</h3>
      {mine.length === 0 ? (
        <div className="mrs-card mrs-empty">제출한 반출요청이 없습니다.</div>
      ) : mine.map(r => (
        <div className="mrs-card" key={r.id} style={{ padding: 16, marginBottom: 10 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
            <span className="mrs-mono" style={{ fontSize: 12, color: 'var(--ink-soft)' }}>{r.reqNo}</span>
            <StatusBadge value={r.status} />
          </div>
          <div style={{ fontSize: 13, color: 'var(--ink-soft)', marginBottom: 8 }}>{r.zoneName}</div>
          {r.items.map(it => (
            <div key={it.id} style={{ fontSize: 13, padding: '3px 0' }}>{it.name} {it.spec} · {it.qty}{it.unit}</div>
          ))}
          {r.status === '반출확인완료' && (
            <button className="mrs-btn mrs-btn-ghost" style={{ marginTop: 10 }} onClick={() => downloadPdf(r)}><Download size={15} /> 반출확인서 다운로드</button>
          )}
        </div>
      ))}
    </div>
  );
}

function LeaderApp({ session, requests, returns, projects, catalog, zones, onSubmit, onSubmitReturn, onConfirmReceipt, onDeleteRequest, saving }) {
  const [tab, setTab] = useState('form');
  const projectName = (projects.find(p => p.id === session.projectId) || {}).name || '(알 수 없음)';
  // 팀장 자신의 반출내역만 필터링
  const myReturns = returns.filter(r => (r.username ? r.username === session.username : r.requester === session.name));
  return (
    <div className="mrs-body">
      <div className="mrs-tabs" style={{ margin: '-20px -20px 20px', padding: '0 20px' }}>
        <button className={`mrs-tab ${tab === 'form' ? 'active' : ''}`} onClick={() => setTab('form')}><ClipboardList size={16} /> 자재요청 시트</button>
        <button className={`mrs-tab ${tab === 'mylist' ? 'active' : ''}`} onClick={() => setTab('mylist')}><ListChecks size={16} /> 요청리스트</button>
        <button className={`mrs-tab ${tab === 'return' ? 'active' : ''}`} onClick={() => setTab('return')}><PackageMinus size={16} /> 물량반출</button>
        <button className={`mrs-tab ${tab === 'returnlist' ? 'active' : ''}`} onClick={() => setTab('returnlist')}><CalendarDays size={16} /> 반출리스트</button>
      </div>
      {tab === 'form' && <RequestForm session={session} projectName={projectName} catalog={catalog} zones={zones} onSubmit={onSubmit} saving={saving} />}
      {tab === 'mylist' && <MyRequestList requests={requests} session={session} projects={projects} onConfirmReceipt={onConfirmReceipt} onDelete={onDeleteRequest} saving={saving} />}
      {tab === 'return' && <ReturnPanel session={session} projectName={projectName} catalog={catalog} zones={zones} returns={returns} onSubmit={onSubmitReturn} saving={saving} />}
      {tab === 'returnlist' && <ReturnsTable returns={myReturns} projects={projects} />}
    </div>
  );
}

// ── 누계요청리스트 (관리자/자재팀 공용) ─────────────────────
function RequestsTable({ requests, projects, onUpdateStatus, onDelete, scope, allowPdf }) {
  const [projectFilter, setProjectFilter] = useState('전체');
  const [zoneQuery, setZoneQuery] = useState('');
  const [processFilter, setProcessFilter] = useState('전체');
  const [statusFilter, setStatusFilter] = useState('전체');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');

  const projectNameById = {};
  projects.forEach(p => { projectNameById[p.id] = p.name; });

  const rows = [];
  requests.forEach(r => {
    if (scope === 'today' && !isToday(r.createdAt)) return;
    r.items.forEach(it => {
      rows.push({
        req: r, reqId: r.id, itemId: it.id, reqNo: r.reqNo, requester: r.requester,
        projectId: r.projectId, projectName: projectNameById[r.projectId] || '(삭제된 프로젝트)',
        zoneName: r.zoneName || '', process: r.process,
        name: it.name, spec: it.spec, color: it.color, qty: it.qty, unit: it.unit,
        status: it.status, note: r.note, createdAt: r.createdAt,
      });
    });
  });

  const statuses = ['전체', ...STATUS_FLOW];
  const filtered = rows.filter(r =>
    (projectFilter === '전체' || r.projectId === projectFilter) &&
    (zoneQuery.trim() === '' || (r.zoneName || '').includes(zoneQuery.trim())) &&
    (processFilter === '전체' || r.process === processFilter) &&
    (statusFilter === '전체' || r.status === statusFilter) &&
    (dateFrom === '' || new Date(r.createdAt) >= new Date(dateFrom + 'T00:00:00')) &&
    (dateTo === '' || new Date(r.createdAt) <= new Date(dateTo + 'T23:59:59'))
  ).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  const counts = STATUS_FLOW.reduce((acc, s) => { acc[s] = rows.filter(r => (projectFilter === '전체' || r.projectId === projectFilter) && r.status === s).length; return acc; }, {});
  const totalCount = rows.filter(r => projectFilter === '전체' || r.projectId === projectFilter).length;

  function exportExcel() {
    const data = filtered.map(r => ({
      '프로젝트': r.projectName, '요청번호': r.reqNo, '요청일시': fmtDate(r.createdAt), '요청자': r.requester,
      '구역': r.zoneName, '공정': r.process, '품목명': r.name, '규격': r.spec, '색상': r.color,
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

  function downloadPdf(req) {
    generateDocPdf({
      title: '거래명세표', docNo: req.reqNo, dateStr: fmtDate(req.confirmedAt || req.createdAt),
      projectName: projectNameById[req.projectId], zoneStr: req.zoneName,
      items: req.items, deliverLabel: '인도자 (자재팀)', deliverName: req.deliverName, deliverSignature: req.deliverSignature,
      receiveLabel: '인수자 (현장팀장)', receiveName: req.receiveName, receiveSignature: req.receiveSignature,
    });
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
        <input className="mrs-input" style={{ width: 160 }} value={zoneQuery} onChange={e => setZoneQuery(e.target.value)} placeholder="구역명 검색" />
        {scope === 'all' && (
          <>
            <input className="mrs-input" type="date" style={{ width: 145 }} value={dateFrom} onChange={e => setDateFrom(e.target.value)} />
            <span style={{ fontSize: 12, color: 'var(--ink-soft)' }}>~</span>
            <input className="mrs-input" type="date" style={{ width: 145 }} value={dateTo} onChange={e => setDateTo(e.target.value)} />
          </>
        )}
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
            <thead><tr><th>프로젝트</th><th>요청번호</th><th>요청일시</th><th>요청자</th><th>구역</th><th>공정</th><th>품목명</th><th>규격</th><th>색상</th><th>수량</th><th>상태</th><th></th><th></th></tr></thead>
            <tbody>
              {filtered.map(r => (
                <tr key={r.itemId}>
                  <td style={{ fontWeight: 600 }}>{r.projectName}</td>
                  <td className="mrs-mono" style={{ fontSize: 12, color: 'var(--ink-soft)' }}>{r.reqNo}</td>
                  <td className="mrs-mono" style={{ fontSize: 12 }}>{fmtDate(r.createdAt)}</td>
                  <td>{r.requester}</td><td>{r.zoneName}</td><td>{r.process}</td>
                  <td style={{ fontWeight: 600 }}>{r.name}</td><td style={{ color: 'var(--ink-soft)' }}>{r.spec || '-'}</td>
                  <td style={{ color: 'var(--ink-soft)' }}>{r.color || '-'}</td>
                  <td className="mrs-mono">{r.qty} {r.unit}</td>
                  <td><StatusSelect value={r.status} options={STATUS_FLOW} onChange={v => onUpdateStatus(r.reqId, r.itemId, v)} /></td>
                  <td>
                    {allowPdf && r.status === '입고완료' && (
                      <button className="mrs-btn mrs-btn-ghost" style={{ padding: '4px 8px', fontSize: 12 }} onClick={() => downloadPdf(r.req)} title="거래명세표"><Download size={13} /></button>
                    )}
                  </td>
                  <td>
                    {onDelete && (
                      <button style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4, color: '#B84B10' }} onClick={() => { if (confirm('이 요청 전체를 삭제할까요?')) onDelete(r.reqId); }} title="요청 삭제"><X size={15} /></button>
                    )}
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

// ── 누계반출리스트 (관리자/자재팀 공용) ─────────────────────
function ReturnsTable({ returns, projects }) {
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const [statusFilter, setStatusFilter] = useState('전체');
  const projectNameById = {};
  projects.forEach(p => { projectNameById[p.id] = p.name; });

  const rows = [];
  returns.forEach(r => {
    r.items.forEach(it => {
      rows.push({
        ret: r, reqId: r.id, reqNo: r.reqNo, requester: r.requester,
        projectName: projectNameById[r.projectId] || '-',
        zoneName: r.zoneName || '', name: it.name, spec: it.spec, color: it.color, qty: it.qty, unit: it.unit,
        status: r.status, reason: r.reason, createdAt: r.createdAt,
      });
    });
  });

  const filtered = rows.filter(r =>
    (statusFilter === '전체' || r.status === statusFilter) &&
    (dateFrom === '' || new Date(r.createdAt) >= new Date(dateFrom + 'T00:00:00')) &&
    (dateTo === '' || new Date(r.createdAt) <= new Date(dateTo + 'T23:59:59'))
  ).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  function exportExcel() {
    const data = filtered.map(r => ({
      '프로젝트': r.projectName, '반출번호': r.reqNo, '반출일시': fmtDate(r.createdAt), '요청자': r.requester,
      '구역': r.zoneName, '품목명': r.name, '규격': r.spec, '색상': r.color,
      '수량': r.qty, '단위': r.unit, '상태': r.status, '사유': r.reason || '',
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{wch:16},{wch:18},{wch:16},{wch:12},{wch:14},{wch:16},{wch:12},{wch:8},{wch:8},{wch:8},{wch:12},{wch:30}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '물량반출');
    const d = new Date();
    XLSX.writeFile(wb, `누계반출리스트_${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}.xlsx`);
  }

  function downloadPdf(r) {
    generateDocPdf({
      title: '반출확인서', docNo: r.reqNo, dateStr: fmtDate(r.confirmedAt || r.createdAt),
      projectName: projectNameById[r.projectId], zoneStr: r.zoneName,
      items: r.items, deliverLabel: '인도자 (현장팀장)', deliverName: r.deliverName, deliverSignature: r.deliverSignature,
      receiveLabel: '인수자 (자재팀)', receiveName: r.receiveName, receiveSignature: r.receiveSignature,
    });
  }

  return (
    <div>
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 14, alignItems: 'center' }}>
        <input className="mrs-input" type="date" style={{ width: 145 }} value={dateFrom} onChange={e => setDateFrom(e.target.value)} />
        <span style={{ fontSize: 12, color: 'var(--ink-soft)' }}>~</span>
        <input className="mrs-input" type="date" style={{ width: 145 }} value={dateTo} onChange={e => setDateTo(e.target.value)} />
        <select className="mrs-select" style={{ width: 'auto' }} value={statusFilter} onChange={e => setStatusFilter(e.target.value)}>
          <option value="전체">전체 상태</option>
          {RETURN_STATUS_FLOW.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
        <div style={{ flex: 1 }} />
        <button className="mrs-btn mrs-btn-primary" onClick={exportExcel} disabled={filtered.length === 0}><Download size={15} /> 엑셀로 내보내기</button>
      </div>

      {filtered.length === 0 ? (
        <div className="mrs-card mrs-empty">표시할 반출요청이 없습니다.</div>
      ) : (
        <div className="mrs-table-wrap">
          <table className="mrs-table">
            <thead><tr><th>프로젝트</th><th>반출번호</th><th>반출일시</th><th>요청자</th><th>구역</th><th>품목명</th><th>규격</th><th>색상</th><th>수량</th><th>상태</th><th></th></tr></thead>
            <tbody>
              {filtered.map((r, i) => (
                <tr key={i}>
                  <td style={{ fontWeight: 600 }}>{r.projectName}</td>
                  <td className="mrs-mono" style={{ fontSize: 12, color: 'var(--ink-soft)' }}>{r.reqNo}</td>
                  <td className="mrs-mono" style={{ fontSize: 12 }}>{fmtDate(r.createdAt)}</td>
                  <td>{r.requester}</td><td>{r.zoneName}</td>
                  <td style={{ fontWeight: 600 }}>{r.name}</td><td style={{ color: 'var(--ink-soft)' }}>{r.spec || '-'}</td>
                  <td style={{ color: 'var(--ink-soft)' }}>{r.color || '-'}</td>
                  <td className="mrs-mono">{r.qty} {r.unit}</td>
                  <td><StatusBadge value={r.status} /></td>
                  <td>
                    {r.status === '반출확인완료' && (
                      <button className="mrs-btn mrs-btn-ghost" style={{ padding: '4px 8px', fontSize: 12 }} onClick={() => downloadPdf(r.ret)} title="반출확인서"><Download size={13} /></button>
                    )}
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

// ── 관리 설정: 프로젝트 ────────────────────────────────
function ProjectManager({ projects, onSave, saving }) {
  const [draft, setDraft] = useState(projects);
  useEffect(() => setDraft(projects), [projects]);
  function addRow() { setDraft([...draft, { id: genId('proj'), name: '' }]); }
  function remove(idx) { setDraft(draft.filter((_, i) => i !== idx)); }
  return (
    <div className="mrs-card" style={{ padding: 18, marginBottom: 20 }}>
      <h3 className="mrs-display" style={{ fontSize: 15, margin: '0 0 12px', fontWeight: 600 }}>프로젝트 관리</h3>
      {draft.map((p, idx) => (
        <div key={p.id} style={{ display: 'flex', gap: 10, marginBottom: 8, alignItems: 'center' }}>
          <input className="mrs-input" value={p.name} onChange={e => { const next = [...draft]; next[idx] = { ...p, name: e.target.value }; setDraft(next); }} />
          <button className="mrs-btn mrs-btn-danger" onClick={() => remove(idx)} style={{ padding: 8 }}><Trash2 size={15} /></button>
        </div>
      ))}
      <div style={{ display: 'flex', gap: 8, marginTop: 10 }}>
        <button className="mrs-btn mrs-btn-ghost" onClick={addRow}><Plus size={15} /> 프로젝트 추가</button>
        <button className="mrs-btn mrs-btn-primary" disabled={saving || draft.some(p => !p.name.trim())} onClick={() => onSave(draft)}><Save size={15} /> 저장</button>
      </div>
    </div>
  );
}

// ── 관리 설정: 구역명 (프로젝트/층/실/구역 + 엑셀 업로드) ──
function ZoneManager({ zones, projects, onSave, saving }) {
  const [draft, setDraft] = useState(zones);
  const fileRef = useRef(null);
  useEffect(() => setDraft(zones), [zones]);

  function addRow() { setDraft([...draft, { id: genId('zone'), projectId: projects[0]?.id || '', floor: '', room: '', zone: '' }]); }
  function remove(idx) { setDraft(draft.filter((_, i) => i !== idx)); }
  function update(idx, field, val) { const next = [...draft]; next[idx] = { ...next[idx], [field]: val }; setDraft(next); }

  function downloadTemplate() {
    const rows = [
      { '프로젝트명': 'P4 Ph4 (삼성물산)', '층': '1F', '실': 'N실', '구역': 'A구역' },
      { '프로젝트명': 'P4 Ph4 (삼성물산)', '층': '1F', '실': 'N실', '구역': 'B구역' },
    ];
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [{ wch: 24 }, { wch: 10 }, { wch: 12 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '구역');
    XLSX.writeFile(wb, '구역_업로드_양식.xlsx');
  }

  function handleUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        const projectByName = {};
        projects.forEach(p => { projectByName[p.name] = p.id; });
        const parsed = rows.map(row => ({
          id: genId('zone'),
          projectId: projectByName[String(row['프로젝트명'] || row['프로젝트'] || '').trim()] || projects[0]?.id || '',
          floor: String(row['층'] || '').trim(),
          room: String(row['실'] || '').trim(),
          zone: String(row['구역'] || '').trim(),
        })).filter(z => z.floor && z.room && z.zone);
        if (parsed.length === 0) { alert('업로드할 데이터가 없습니다. 양식을 확인해주세요.'); return; }
        if (confirm(`${parsed.length}건의 구역을 추가할까요? (기존 구역은 유지됩니다)`)) {
          setDraft([...draft, ...parsed]);
        }
      } catch (err) { alert('엑셀 읽기 실패: ' + err.message); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  }

  const projectName = id => (projects.find(p => p.id === id) || {}).name || '-';

  return (
    <div className="mrs-card" style={{ padding: 18, marginBottom: 20 }}>
      <h3 className="mrs-display" style={{ fontSize: 15, margin: '0 0 12px', fontWeight: 600 }}><MapPin size={15} style={{ verticalAlign: -2 }} /> 구역 관리 (프로젝트/층/실/구역)</h3>
      <div style={{ display: 'flex', gap: 8, marginBottom: 12, flexWrap: 'wrap' }}>
        <button className="mrs-btn mrs-btn-ghost" onClick={addRow}><Plus size={15} /> 구역 추가</button>
        <button className="mrs-btn mrs-btn-primary" onClick={() => fileRef.current?.click()}><Download size={15} style={{ transform: 'rotate(180deg)' }} /> 엑셀 업로드</button>
        <button className="mrs-btn mrs-btn-ghost" onClick={downloadTemplate}><Download size={15} /> 양식 다운로드</button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleUpload} />
      </div>

      {draft.length === 0 ? (
        <div style={{ fontSize: 13, color: 'var(--ink-soft)', padding: '10px 0' }}>등록된 구역이 없습니다.</div>
      ) : (
        <>
          <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 0.8fr 1fr 1fr auto', gap: 6, marginBottom: 4 }}>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>프로젝트</span>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>층</span>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>실</span>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>구역</span>
            <span></span>
          </div>
          {draft.map((z, idx) => (
            <div key={z.id} style={{ display: 'grid', gridTemplateColumns: '1.5fr 0.8fr 1fr 1fr auto', gap: 6, marginBottom: 6, alignItems: 'center' }}>
              <select className="mrs-select" value={z.projectId} onChange={e => update(idx, 'projectId', e.target.value)}>
                {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
              <input className="mrs-input" value={z.floor} onChange={e => update(idx, 'floor', e.target.value)} placeholder="1F" />
              <input className="mrs-input" value={z.room} onChange={e => update(idx, 'room', e.target.value)} placeholder="N실" />
              <input className="mrs-input" value={z.zone} onChange={e => update(idx, 'zone', e.target.value)} placeholder="A구역" />
              <button className="mrs-btn mrs-btn-danger" onClick={() => remove(idx)} style={{ padding: 8 }}><Trash2 size={15} /></button>
            </div>
          ))}
        </>
      )}
      <button className="mrs-btn mrs-btn-primary" style={{ marginTop: 10 }} disabled={saving || draft.some(z => !z.floor.trim() || !z.room.trim() || !z.zone.trim() || !z.projectId)} onClick={() => onSave(draft)}><Save size={15} /> 전체 저장</button>
    </div>
  );
}

// ── 관리 설정: 품목/규격/색상/단위 (+ 엑셀 업로드) ──────────
function CatalogManager({ catalog, onSave, saving }) {
  const [draft, setDraft] = useState(catalog);
  const fileRef = useRef(null);
  useEffect(() => setDraft(catalog), [catalog]);

  function addRow() { setDraft([...draft, { id: genId('cat'), name: '', spec: '', color: 'N/A', unit: 'EA' }]); }
  function remove(idx) { setDraft(draft.filter((_, i) => i !== idx)); }
  function update(idx, field, val) { const next = [...draft]; next[idx] = { ...next[idx], [field]: val }; setDraft(next); }

  function downloadTemplate() {
    const rows = [
      { '품목명': '무나사전선관', '규격': 'E19', '색상': 'N/A', '단위': '본' },
      { '품목명': '통신케이블', '규격': '14TP', '색상': '적', '단위': '롤' },
    ];
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [{ wch: 16 }, { wch: 12 }, { wch: 8 }, { wch: 8 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '품목');
    XLSX.writeFile(wb, '품목_업로드_양식.xlsx');
  }

  function handleUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        const parsed = rows.map(row => ({
          id: genId('cat'),
          name: String(row['품목명'] || row['품목'] || '').trim(),
          spec: String(row['규격'] || '').trim(),
          color: String(row['색상'] || 'N/A').trim() || 'N/A',
          unit: String(row['단위'] || 'EA').trim() || 'EA',
        })).filter(it => it.name && it.spec);
        if (parsed.length === 0) { alert('업로드할 데이터가 없습니다.'); return; }
        if (confirm(`${parsed.length}건의 품목을 추가할까요? (기존 품목은 유지됩니다)`)) {
          setDraft([...draft, ...parsed]);
        }
      } catch (err) { alert('엑셀 읽기 실패: ' + err.message); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  }

  return (
    <div className="mrs-card" style={{ padding: 18, marginBottom: 20 }}>
      <h3 className="mrs-display" style={{ fontSize: 15, margin: '0 0 12px', fontWeight: 600 }}><Package size={15} style={{ verticalAlign: -2 }} /> 품목 관리</h3>
      <div style={{ fontSize: 12, color: 'var(--ink-soft)', marginBottom: 10 }}>색상이 없는 품목은 N/A로 두세요.</div>
      <div style={{ display: 'flex', gap: 8, marginBottom: 12, flexWrap: 'wrap' }}>
        <button className="mrs-btn mrs-btn-ghost" onClick={addRow}><Plus size={15} /> 품목 추가</button>
        <button className="mrs-btn mrs-btn-primary" onClick={() => fileRef.current?.click()}><Download size={15} style={{ transform: 'rotate(180deg)' }} /> 엑셀 업로드</button>
        <button className="mrs-btn mrs-btn-ghost" onClick={downloadTemplate}><Download size={15} /> 양식 다운로드</button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleUpload} />
      </div>

      {draft.length === 0 ? (
        <div style={{ fontSize: 13, color: 'var(--ink-soft)', padding: '10px 0' }}>등록된 품목이 없습니다.</div>
      ) : (
        <>
          <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 1fr 0.8fr 0.8fr auto', gap: 6, marginBottom: 4 }}>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>품목명</span>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>규격</span>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>단위</span>
            <span style={{ fontSize: 10, color: 'var(--ink-soft)' }}>색상</span>
            <span></span>
          </div>
          {draft.map((it, idx) => (
            <div key={it.id} style={{ display: 'grid', gridTemplateColumns: '1.5fr 1fr 0.8fr 0.8fr auto', gap: 6, marginBottom: 6, alignItems: 'center' }}>
              <input className="mrs-input" value={it.name} onChange={e => update(idx, 'name', e.target.value)} placeholder="무나사전선관" />
              <input className="mrs-input" value={it.spec} onChange={e => update(idx, 'spec', e.target.value)} placeholder="E19" />
              <select className="mrs-select" value={it.unit} onChange={e => update(idx, 'unit', e.target.value)}>
                {UNITS.map(u => <option key={u} value={u}>{u}</option>)}
              </select>
              <input className="mrs-input" value={it.color} onChange={e => update(idx, 'color', e.target.value)} placeholder="N/A" />
              <button className="mrs-btn mrs-btn-danger" onClick={() => remove(idx)} style={{ padding: 8 }}><Trash2 size={15} /></button>
            </div>
          ))}
        </>
      )}
      <button className="mrs-btn mrs-btn-primary" style={{ marginTop: 10 }} disabled={saving || draft.some(it => !it.name.trim() || !it.spec.trim())} onClick={() => onSave(draft)}><Save size={15} /> 전체 저장</button>
    </div>
  );
}

// ── 관리 설정: 팀장/자재팀 계정 ─────────────────────────
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
    if (!newStaff.name.trim() || !newStaff.username.trim() || !newStaff.password) { alert('이름, 아이디, 비밀번호를 모두 입력해주세요.'); return; }
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

// ── 관리자 전체 화면 (자재팀과 동일한 구성 + 관리 설정) ───
function AdminApp({ session, requests, returns, projects, users, zones, catalog, onUpdateStatus, onDelete, onSaveProjects, onSaveZones, onSaveCatalog, onAddUser, onUpdateUser, onDeleteUser, onConfirmReturn, savingSettings }) {
  const [tab, setTab] = useState('outbound');
  return (
    <div className="mrs-body">
      <div className="mrs-tabs" style={{ margin: '-20px -20px 20px', padding: '0 20px' }}>
        <button className={`mrs-tab ${tab === 'outbound' ? 'active' : ''}`} onClick={() => setTab('outbound')}><ClipboardList size={16} /> 발주확인 대기</button>
        <button className={`mrs-tab ${tab === 'all' ? 'active' : ''}`} onClick={() => setTab('all')}><LayoutDashboard size={16} /> 누계 요청리스트</button>
        <button className={`mrs-tab ${tab === 'return' ? 'active' : ''}`} onClick={() => setTab('return')}><PackageMinus size={16} /> 반출확인 대기</button>
        <button className={`mrs-tab ${tab === 'returns' ? 'active' : ''}`} onClick={() => setTab('returns')}><CalendarDays size={16} /> 누계 반출리스트</button>
        <button className={`mrs-tab ${tab === 'manage' ? 'active' : ''}`} onClick={() => setTab('manage')}><Users size={16} /> 관리 설정</button>
      </div>

      {tab === 'outbound' && <MaterialOutbound requests={requests} projects={projects} onUpdateStatus={onUpdateStatus} />}
      {tab === 'all' && <RequestsTable requests={requests} projects={projects} onUpdateStatus={onUpdateStatus} onDelete={onDelete} scope="all" allowPdf />}
      {tab === 'return' && <MaterialReturns returns={returns} onConfirmReturn={onConfirmReturn} saving={savingSettings} />}
      {tab === 'returns' && <ReturnsTable returns={returns} projects={projects} />}
      {tab === 'manage' && (
        <div>
          <ProjectManager projects={projects} onSave={onSaveProjects} saving={savingSettings} />
          <ZoneManager zones={zones} projects={projects} onSave={onSaveZones} saving={savingSettings} />
          <CatalogManager catalog={catalog} onSave={onSaveCatalog} saving={savingSettings} />
          <StaffManager users={users} projects={projects} onAdd={onAddUser} onUpdate={onUpdateUser} onDelete={onDeleteUser} saving={savingSettings} />
          <AdminAccountManager session={session} onUpdate={onUpdateUser} saving={savingSettings} />
        </div>
      )}
    </div>
  );
}

// ── 자재팀: 발주확인 대기 ──────────────────────────────
function MaterialOutbound({ requests, projects, onUpdateStatus }) {
  const projectNameById = {};
  projects.forEach(p => { projectNameById[p.id] = p.name; });
  const rows = [];
  requests.forEach(r => {
    r.items.forEach(it => {
      if (it.status !== '요청됨') return;
      rows.push({
        reqId: r.id, itemId: it.id, requester: r.requester, projectName: projectNameById[r.projectId] || '-',
        zoneName: r.zoneName || '', name: it.name, spec: it.spec, color: it.color, qty: it.qty, unit: it.unit, status: it.status, createdAt: r.createdAt,
      });
    });
  });
  rows.sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
  return (
    <div>
      {rows.length === 0 ? (
        <div className="mrs-card mrs-empty">발주확인 대기 중인 요청이 없습니다.</div>
      ) : (
        <div className="mrs-table-wrap">
          <table className="mrs-table">
            <thead><tr><th>요청자</th><th>프로젝트</th><th>구역</th><th>품목</th><th>규격</th><th>색상</th><th>수량</th><th>상태</th><th></th></tr></thead>
            <tbody>
              {rows.map(r => (
                <tr key={r.itemId}>
                  <td>{r.requester}</td><td>{r.projectName}</td><td>{r.zoneName}</td>
                  <td style={{ fontWeight: 600 }}>{r.name}</td><td style={{ color: 'var(--ink-soft)' }}>{r.spec}</td><td style={{ color: 'var(--ink-soft)' }}>{r.color}</td>
                  <td className="mrs-mono">{r.qty} {r.unit}</td>
                  <td><StatusBadge value={r.status} /></td>
                  <td><button className="mrs-btn mrs-btn-primary" style={{ padding: '6px 10px', fontSize: 12 }} onClick={() => onUpdateStatus(r.reqId, r.itemId, '확인됨')}>발주확인</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ── 자재팀: 반출확인 대기 ──────────────────────────────
function MaterialReturns({ returns, onConfirmReturn, saving }) {
  const [confirmingId, setConfirmingId] = useState(null);
  const pending = returns.filter(r => r.status !== '반출확인완료').slice().sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
  async function handleConfirm(r, sig) {
    await onConfirmReturn(r.id, sig);
    setConfirmingId(null);
  }
  return (
    <div>
      {pending.length === 0 ? (
        <div className="mrs-card mrs-empty">반출확인 대기 중인 요청이 없습니다.</div>
      ) : pending.map(r => (
        <div className="mrs-card" key={r.id} style={{ padding: 16, marginBottom: 12 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
            <span style={{ fontSize: 14, fontWeight: 600 }}>{r.requester}</span>
            <span className="mrs-mono" style={{ fontSize: 12, color: 'var(--ink-soft)' }}>{fmtDate(r.createdAt)}</span>
          </div>
          <div style={{ fontSize: 13, color: 'var(--ink-soft)', marginBottom: 8 }}>{r.zoneName}{r.reason ? ` · ${r.reason}` : ''}</div>
          {r.items.map(it => (
            <div key={it.id} style={{ fontSize: 13, padding: '3px 0' }}>{it.name} {it.spec} · {it.qty}{it.unit}</div>
          ))}
          {confirmingId !== r.id ? (
            <button className="mrs-btn mrs-btn-primary" style={{ marginTop: 10 }} onClick={() => setConfirmingId(r.id)}><PenLine size={15} /> 반출확인</button>
          ) : (
            <ConfirmSignaturePanel
              deliverLabel="인도자 (현장팀장)" receiveLabel="인수자 (자재팀)"
              deliverNameDefault={r.requester}
              onCancel={() => setConfirmingId(null)}
              onSubmit={sig => handleConfirm(r, sig)}
              saving={saving}
            />
          )}
        </div>
      ))}
    </div>
  );
}

function MaterialApp({ requests, projects, returns, onUpdateStatus, onDelete, onConfirmReturn, savingSettings }) {
  const [tab, setTab] = useState('outbound');
  return (
    <div className="mrs-body">
      <div className="mrs-tabs" style={{ margin: '-20px -20px 20px', padding: '0 20px' }}>
        <button className={`mrs-tab ${tab === 'outbound' ? 'active' : ''}`} onClick={() => setTab('outbound')}><ClipboardList size={16} /> 발주확인 대기</button>
        <button className={`mrs-tab ${tab === 'all' ? 'active' : ''}`} onClick={() => setTab('all')}><LayoutDashboard size={16} /> 누계 요청리스트</button>
        <button className={`mrs-tab ${tab === 'return' ? 'active' : ''}`} onClick={() => setTab('return')}><PackageMinus size={16} /> 반출확인 대기</button>
        <button className={`mrs-tab ${tab === 'returns' ? 'active' : ''}`} onClick={() => setTab('returns')}><CalendarDays size={16} /> 누계 반출리스트</button>
      </div>
      {tab === 'outbound' && <MaterialOutbound requests={requests} projects={projects} onUpdateStatus={onUpdateStatus} />}
      {tab === 'all' && <RequestsTable requests={requests} projects={projects} onUpdateStatus={onUpdateStatus} onDelete={onDelete} scope="all" allowPdf />}
      {tab === 'return' && <MaterialReturns returns={returns} onConfirmReturn={onConfirmReturn} saving={savingSettings} />}
      {tab === 'returns' && <ReturnsTable returns={returns} projects={projects} />}
    </div>
  );
}

// ── 최상위 App ──────────────────────────────────────────
export default function App() {
  const [session, setSession] = useState(null);
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginError, setLoginError] = useState(null);

  const [requests, setRequests] = useState([]);
  const [returns, setReturns] = useState([]);
  const [projects, setProjects] = useState(DEFAULT_PROJECTS);
  const [users, setUsers] = useState([]);
  const [zones, setZones] = useState([]);
  const [catalog, setCatalog] = useState([]);
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
      const calls = [apiGet('list'), apiGet('projects'), apiGet('returns'), apiGet('zones'), apiGet('catalog')];
      if (role === 'admin') calls.push(apiGet('users'));
      const results = await Promise.all(calls);
      setRequests(Array.isArray(results[0]) ? results[0] : []);
      setProjects(Array.isArray(results[1]) && results[1].length ? results[1] : DEFAULT_PROJECTS);
      setReturns(Array.isArray(results[2]) ? results[2] : []);
      setZones(Array.isArray(results[3]) ? results[3] : []);
      setCatalog(Array.isArray(results[4]) ? results[4] : []);
      if (role === 'admin') setUsers(Array.isArray(results[5]) ? results[5] : []);
    } catch (e) {
      setError('데이터를 불러오지 못했습니다. 네트워크를 확인해주세요.');
    } finally {
      setDataLoading(false);
    }
  }

  async function handleRefresh() {
    if (!session) return;
    await loadData(session.role);
  }

  async function handleSubmit(payload) {
    setRequests([payload, ...requests]);
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

  async function handleConfirmReceipt(reqId, sig) {
    const next = requests.map(r => r.id !== reqId ? r : {
      ...r, items: r.items.map(it => ({ ...it, status: '입고완료' })),
      deliverName: sig.deliverName, deliverSignature: sig.deliverSignature,
      receiveName: sig.receiveName, receiveSignature: sig.receiveSignature, confirmedAt: new Date().toISOString(),
    });
    setRequests(next);
    setSaving(true); setError(null);
    try { await apiPost('confirmReceipt', { data: { reqId, ...sig } }); }
    catch (e) { setError('입고확인에 실패했습니다.'); await loadData(session.role); }
    finally { setSaving(false); }
  }

  async function handleDelete(reqId) {
    setRequests(requests.filter(r => r.id !== reqId));
    setSaving(true); setError(null);
    try { await apiPost('deleteRequest', { reqId }); }
    catch (e) { setError('삭제에 실패했습니다.'); await loadData(session.role); }
    finally { setSaving(false); }
  }

  // 관리자 비밀번호 확인 후 요청 삭제
  async function handleDeleteWithPassword(reqId) {
    const pwd = prompt('정말 삭제하시겠습니까?\n관리자 비밀번호를 입력해주세요.');
    if (pwd === null) return; // 취소
    if (!pwd) { alert('비밀번호를 입력해주세요.'); return; }
    // 서버에 재로그인으로 비밀번호 검증
    try {
      const check = await apiPost('login', { username: session.username, password: pwd });
      if (check.error) { alert('비밀번호가 올바르지 않습니다.'); return; }
    } catch (e) { alert('비밀번호 확인 중 오류가 발생했습니다.'); return; }
    await handleDelete(reqId);
  }

  async function handleSubmitReturn(payload) {
    setReturns([{ ...payload, status: '반출요청' }, ...returns]);
    setSaving(true); setError(null);
    try { await apiPost('addReturn', { payload }); }
    catch (e) { setError('반출요청 저장에 실패했습니다.'); await loadData(session.role); }
    finally { setSaving(false); }
  }

  async function handleConfirmReturn(reqId, sig) {
    const next = returns.map(r => r.id !== reqId ? r : {
      ...r, status: '반출확인완료',
      deliverName: sig.deliverName, deliverSignature: sig.deliverSignature,
      receiveName: sig.receiveName, receiveSignature: sig.receiveSignature, confirmedAt: new Date().toISOString(),
    });
    setReturns(next);
    setSavingSettings(true); setError(null);
    try { await apiPost('confirmReturn', { data: { reqId, ...sig } }); }
    catch (e) { setError('반출확인에 실패했습니다.'); await loadData(session.role); }
    finally { setSavingSettings(false); }
  }

  async function handleSaveProjects(next) {
    setSavingSettings(true); setError(null);
    try { await apiPost('saveProjects', { projects: next }); setProjects(next); }
    catch (e) { setError('프로젝트 저장에 실패했습니다.'); }
    finally { setSavingSettings(false); }
  }

  async function handleSaveZones(next) {
    setSavingSettings(true); setError(null);
    try { await apiPost('saveZones', { zones: next }); setZones(next); }
    catch (e) { setError('구역 저장에 실패했습니다.'); }
    finally { setSavingSettings(false); }
  }

  async function handleSaveCatalog(next) {
    setSavingSettings(true); setError(null);
    try { await apiPost('saveCatalog', { items: next }); setCatalog(next); }
    catch (e) { setError('품목 저장에 실패했습니다.'); }
    finally { setSavingSettings(false); }
  }

  async function handleAddUser(user) {
    setSavingSettings(true); setError(null);
    try {
      const res = await apiPost('addUser', { user });
      if (res.error) { alert(res.error); return; }
      await loadData('admin');
    } catch (e) { setError('계정 추가에 실패했습니다.'); }
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
          <button className="mrs-refresh-btn" onClick={handleRefresh} disabled={dataLoading} title="새로고침">
            <RefreshCw size={13} className={dataLoading ? 'mrs-spin' : ''} /> 새로고침
          </button>
          <button className="mrs-logout-btn" onClick={() => { setSession(null); setRequests([]); setReturns([]); setUsers([]); setZones([]); setCatalog([]); }}><LogOut size={13} /> 로그아웃</button>
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
          session={session} requests={requests} returns={returns} projects={projects} users={users} zones={zones} catalog={catalog}
          onUpdateStatus={handleUpdateStatus} onDelete={handleDeleteWithPassword} onConfirmReturn={handleConfirmReturn}
          onSaveProjects={handleSaveProjects} onSaveZones={handleSaveZones} onSaveCatalog={handleSaveCatalog}
          onAddUser={handleAddUser} onUpdateUser={handleUpdateUser} onDeleteUser={handleDeleteUser}
          savingSettings={savingSettings}
        />
      ) : session.role === 'material' ? (
        <MaterialApp requests={requests} projects={projects} returns={returns} onUpdateStatus={handleUpdateStatus} onDelete={handleDelete} onConfirmReturn={handleConfirmReturn} savingSettings={savingSettings} />
      ) : (
        <LeaderApp session={session} requests={requests} returns={returns} projects={projects} catalog={catalog} zones={zones} onSubmit={handleSubmit} onSubmitReturn={handleSubmitReturn} onConfirmReceipt={handleConfirmReceipt} onDeleteRequest={handleDelete} saving={saving} />
      )}
    </div>
  );
}
