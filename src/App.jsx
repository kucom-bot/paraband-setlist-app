import { useState, useEffect, useRef } from 'react';
import { DragDropContext, Droppable, Draggable } from '@hello-pangea/dnd';
import {
  GripVertical, Music, Plus, ArrowLeft, ZoomIn, ZoomOut, Upload,
  Image as ImageIcon, Trash2, ChevronLeft, ChevronRight, Radio, Search, Play, Square,
  Type, FileText, Save, PlusCircle, MinusCircle, Sparkles, Clock, Tag, Settings,
  X, Edit2, Link, RotateCcw, Eye, EyeOff, Zap, Activity,
  Folder, FolderPlus, Download, AlertCircle,
  Table, CheckSquare, ArrowUpFromLine, LogIn, LogOut, User, Menu,
  Repeat, Scissors, BookOpen, Printer, ChevronDown, Copy
} from 'lucide-react';
import { db } from './firebase';
import {
  collection, onSnapshot, addDoc, updateDoc, doc, query, orderBy,
  setDoc, deleteDoc, writeBatch, where, getDoc, getDocs, serverTimestamp
} from 'firebase/firestore';
import {
  getAuth, signInWithEmailAndPassword, createUserWithEmailAndPassword,
  signOut, onAuthStateChanged, updateProfile
} from 'firebase/auth';
import * as XLSX from 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm';

// ==========================================
// Helpers
// ==========================================
const parseDuration = (durStr) => {
  if (!durStr) return 0;
  const parts = String(durStr).split(':');
  if (parts.length === 2) return (parseInt(parts[0]) || 0) * 60 + (parseInt(parts[1]) || 0);
  return 0;
};
const formatDuration = (totalSeconds) => {
  const m = Math.floor(totalSeconds / 60);
  const s = totalSeconds % 60;
  return `${m} นาที ${s < 10 ? '0' : ''}${s} วินาที`;
};
const NOTES = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'];
const FLATS = { 'Db': 'C#', 'Eb': 'D#', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#' };
const transposeText = (text, steps) => {
  if (!text) return "";
  const chordRegex = /([^a-zA-Z0-9]|^)([CDEFGAB][#b]?(?:m|maj|min|dim|aug|sus)?\d*(?:\/[CDEFGAB][#b]?)?)(?=[^a-zA-Z0-9]|$)/g;
  return text.replace(chordRegex, (match, prefix, chord) => {
    const rootMatch = chord.match(/^([CDEFGAB][#b]?)(.*)$/);
    if (!rootMatch) return match;
    let root = rootMatch[1]; let rest = rootMatch[2] || '';
    let rootNote = FLATS[root] || root;
    let index = NOTES.indexOf(rootNote);
    if (index === -1) return match;
    let newRoot = NOTES[(index + steps + 12) % 12];
    if (rest.includes('/')) {
      const parts = rest.split('/'); const bass = parts[1];
      let bassNote = FLATS[bass] || bass; let bassIdx = NOTES.indexOf(bassNote);
      if (bassIdx !== -1) rest = parts[0] + '/' + NOTES[(bassIdx + steps + 12) % 12];
    }
    return prefix + newRoot + rest;
  });
};
const detectKey = (text) => {
  if (!text) return null;
  const keyMatch = text.match(/(?:Key|คีย์|key)\s*[:=]?\s*([CDEFGAB][#b]?m?)/i);
  if (keyMatch) return keyMatch[1];
  const chordRegex = /(?:^|\s)([CDEFGAB][#b]?m?)(?:\s|$)/g;
  const counts = {};
  let m;
  while ((m = chordRegex.exec(text)) !== null) counts[m[1]] = (counts[m[1]] || 0) + 1;
  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  return sorted.length > 0 ? sorted[0][0] : null;
};

const TAG_COLORS = [
  { bg: 'bg-blue-900/60', border: 'border-blue-700', text: 'text-blue-300' },
  { bg: 'bg-green-900/60', border: 'border-green-700', text: 'text-green-300' },
  { bg: 'bg-yellow-900/60', border: 'border-yellow-700', text: 'text-yellow-300' },
  { bg: 'bg-red-900/60', border: 'border-red-700', text: 'text-red-300' },
  { bg: 'bg-purple-900/60', border: 'border-purple-700', text: 'text-purple-300' },
  { bg: 'bg-pink-900/60', border: 'border-pink-700', text: 'text-pink-300' },
  { bg: 'bg-orange-900/60', border: 'border-orange-700', text: 'text-orange-300' },
  { bg: 'bg-cyan-900/60', border: 'border-cyan-700', text: 'text-cyan-300' },
];
const getTagColor = (tag) => TAG_COLORS[Math.abs(tag.split('').reduce((a, c) => a + c.charCodeAt(0), 0)) % TAG_COLORS.length];

// ==========================================
// Excel Template + Import
// ==========================================
const downloadTemplate = () => {
  const ws = XLSX.utils.aoa_to_sheet([
    ['ชื่อเพลง (title)*', 'ศิลปิน (artist)', 'คีย์ (key)', 'BPM', 'ความยาว (duration mm:ss)', 'แท็ก (tags คั่นด้วย,)'],
    ['สักวันหนึ่ง', 'Bodyslam', 'C', '90', '04:20', 'เพลงช้า,เปิดตัว'],
    ['มาตาม', 'Labanoon', 'Am', '120', '03:45', 'เพลงเร็ว'],
  ]);
  ws['!cols'] = [{ wch: 22 }, { wch: 16 }, { wch: 8 }, { wch: 6 }, { wch: 22 }, { wch: 24 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Setlist Template');
  XLSX.writeFile(wb, 'setlist_template.xlsx');
};

const parseImportFile = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (rows.length < 2) throw new Error('ไฟล์ว่างหรือไม่มีข้อมูล');
      const header = rows[0].map(h => String(h).toLowerCase().trim());
      const colIdx = {
        title: header.findIndex(h => h.includes('title') || h.includes('ชื่อ')),
        artist: header.findIndex(h => h.includes('artist') || h.includes('ศิลปิน')),
        key: header.findIndex(h => h.includes('key') || h.includes('คีย์')),
        bpm: header.findIndex(h => h.includes('bpm')),
        duration: header.findIndex(h => h.includes('duration') || h.includes('ความยาว')),
        tags: header.findIndex(h => h.includes('tag') || h.includes('แท็ก')),
      };
      if (colIdx.title === -1) throw new Error('ไม่พบคอลัมน์ "ชื่อเพลง (title)"');
      const songs = rows.slice(1)
        .filter(row => String(row[colIdx.title] || '').trim())
        .map((row, i) => ({
          title: String(row[colIdx.title] || '').trim(),
          artist: colIdx.artist >= 0 ? String(row[colIdx.artist] || '').trim() || 'ไม่ระบุ' : 'ไม่ระบุ',
          key: colIdx.key >= 0 ? String(row[colIdx.key] || '').trim() || 'C' : 'C',
          bpm: colIdx.bpm >= 0 ? String(row[colIdx.bpm] || '').trim() || '120' : '120',
          duration: colIdx.duration >= 0 ? String(row[colIdx.duration] || '').trim() || '03:30' : '03:30',
          tags: colIdx.tags >= 0 ? String(row[colIdx.tags] || '').split(',').map(t => t.trim()).filter(Boolean) : [],
          order: i, imageUrl: null, chordText: null, links: [], notes: '', sharedCues: '', history: [], sections: [],
        }));
      resolve(songs);
    } catch (err) { reject(err); }
  };
  reader.onerror = () => reject(new Error('อ่านไฟล์ไม่สำเร็จ'));
  reader.readAsArrayBuffer(file);
});

// ==========================================
// PDF Export
// ==========================================
const exportSetlistPDF = (songs, playlist, includeChords = false) => {
  const win = window.open('', '_blank');
  const chordRegex = /([^a-zA-Z0-9]|^)([CDEFGAB][#b]?(?:m|maj|min|dim|aug|sus)?\d*(?:\/[CDEFGAB][#b]?)?)(?=[^a-zA-Z0-9]|$)/g;

  const renderChordLine = (line) => {
    return line.replace(chordRegex, (match, prefix, chord) =>
      `${prefix}<span style="color:#b45309;font-weight:bold;background:#fef3c7;padding:0 3px;border-radius:3px;">${chord}</span>`
    );
  };

  const totalSecs = songs.reduce((a, s) => a + parseDuration(s.duration), 0);

  const songHTML = songs.map((song, i) => `
    <div class="song" style="page-break-inside:avoid; margin-bottom:${includeChords && song.chordText ? '24px' : '8px'}; padding:${includeChords && song.chordText ? '12px' : '8px 10px'}; border:1px solid #e5e7eb; border-radius:8px; background:${includeChords && song.chordText ? '#fff' : '#f9fafb'};">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <div style="display:flex;align-items:center;gap:10px;">
          <span style="font-size:18px;font-weight:bold;color:#6b7280;min-width:28px;">${i + 1}.</span>
          <div>
            <span style="font-size:${includeChords ? '16px' : '14px'};font-weight:bold;color:#111;">${song.title}</span>
            <span style="font-size:12px;color:#6b7280;margin-left:8px;">${song.artist}</span>
          </div>
        </div>
        <div style="display:flex;gap:8px;align-items:center;flex-shrink:0;">
          ${song.tags?.length ? song.tags.map(t => `<span style="font-size:10px;padding:2px 7px;background:#eff6ff;color:#3b82f6;border-radius:99px;border:1px solid #bfdbfe;">${t}</span>`).join('') : ''}
          <span style="font-size:12px;background:#1e3a8a;color:#bfdbfe;padding:2px 8px;border-radius:5px;font-family:monospace;">Key ${song.key}</span>
          ${song.bpm ? `<span style="font-size:11px;color:#6b7280;">♩${song.bpm}</span>` : ''}
          ${song.duration ? `<span style="font-size:11px;color:#d97706;">⏱${song.duration}</span>` : ''}
        </div>
      </div>
      ${song.sharedCues ? `<div style="margin-top:6px;padding:4px 8px;background:#eff6ff;border-left:3px solid #3b82f6;font-size:11px;color:#1d4ed8;border-radius:0 4px 4px 0;">${song.sharedCues}</div>` : ''}
      ${includeChords && song.chordText ? `
        <div style="margin-top:10px;padding:10px;background:#f8f9fa;border-radius:6px;border:1px solid #e5e7eb;font-family:monospace;font-size:12px;line-height:1.9;white-space:pre-wrap;word-break:break-word;">
          ${song.chordText.split('\n').map(line => `<div>${renderChordLine(line) || '&nbsp;'}</div>`).join('')}
        </div>` : ''}
    </div>
  `).join('');

  win.document.write(`<!DOCTYPE html>
<html lang="th"><head>
<meta charset="UTF-8">
<title>Setlist: ${playlist?.name || 'Setlist'}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Sarabun', sans-serif; padding: 24px 32px; color: #111; background: #fff; }
  .header { border-bottom: 2px solid #1e3a8a; padding-bottom: 12px; margin-bottom: 20px; }
  .header h1 { font-size: 24px; font-weight: bold; color: #1e3a8a; }
  .header .meta { font-size: 13px; color: #6b7280; margin-top: 4px; display: flex; gap: 16px; }
  @media print { body { padding: 12px 18px; } .no-print { display: none; } }
</style>
</head><body>
<div class="header">
  <div style="display:flex;justify-content:space-between;align-items:flex-end;">
    <div>
      <h1>${playlist?.icon || '🎵'} ${playlist?.name || 'Setlist'}</h1>
      ${playlist?.description ? `<p style="font-size:13px;color:#6b7280;margin-top:2px;">${playlist.description}</p>` : ''}
      <div class="meta">
        <span>📋 ${songs.length} เพลง</span>
        <span>⏱ รวม ${formatDuration(totalSecs)}</span>
        <span>📅 พิมพ์: ${new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' })}</span>
      </div>
    </div>
    <button class="no-print" onclick="window.print()" style="padding:8px 18px;background:#1e3a8a;color:#fff;border:none;border-radius:8px;cursor:pointer;font-size:14px;">🖨️ พิมพ์</button>
  </div>
</div>
${songHTML}
</body></html>`);
  win.document.close();
};

// ==========================================
// Sub-components
// ==========================================

// TagBadge
const TagBadge = ({ tag, onRemove, small = false }) => {
  const c = getTagColor(tag);
  return (
    <span className={`flex items-center gap-1 rounded-full border ${c.bg} ${c.border} ${c.text} ${small ? 'text-[10px] px-1.5 py-0' : 'text-xs px-2 py-0.5'}`}>
      {tag}
      {onRemove && <button onClick={() => onRemove(tag)}><X size={small ? 8 : 10} /></button>}
    </span>
  );
};

// ---- Auth Modal ----
function AuthModal({ onClose }) {
  const auth = getAuth();
  const [mode, setMode] = useState('login'); // login | register
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [displayName, setDisplayName] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSubmit = async () => {
    setError(''); setLoading(true);
    try {
      if (mode === 'login') {
        await signInWithEmailAndPassword(auth, email, password);
      } else {
        const cred = await createUserWithEmailAndPassword(auth, email, password);
        if (displayName) await updateProfile(cred.user, { displayName });
      }
      onClose();
    } catch (e) {
      const msgs = {
        'auth/invalid-email': 'อีเมลไม่ถูกต้อง',
        'auth/wrong-password': 'รหัสผ่านไม่ถูกต้อง',
        'auth/user-not-found': 'ไม่พบบัญชีนี้',
        'auth/email-already-in-use': 'อีเมลนี้มีบัญชีแล้ว',
        'auth/weak-password': 'รหัสผ่านสั้นเกินไป (อย่างน้อย 6 ตัว)',
      };
      setError(msgs[e.code] || e.message);
    } finally { setLoading(false); }
  };

  return (
    <div className="fixed inset-0 bg-black/85 z-50 flex items-center justify-center p-4" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="bg-gray-800 rounded-2xl w-full max-w-sm border border-gray-700 shadow-2xl">
        <div className="flex justify-between items-center p-5 border-b border-gray-700">
          <h2 className="text-xl font-bold flex items-center gap-2">
            <User size={20} className="text-blue-400" />
            {mode === 'login' ? 'เข้าสู่ระบบ' : 'สร้างบัญชีใหม่'}
          </h2>
          <button onClick={onClose}><X size={20} className="text-gray-400 hover:text-white" /></button>
        </div>
        <div className="p-5 space-y-3">
          {mode === 'register' && (
            <div>
              <label className="text-xs text-gray-400 mb-1 block">ชื่อ (แสดงในวง)</label>
              <input value={displayName} onChange={e => setDisplayName(e.target.value)} placeholder="เช่น นัท กีตาร์"
                className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none" />
            </div>
          )}
          <div>
            <label className="text-xs text-gray-400 mb-1 block">อีเมล</label>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="your@email.com"
              className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none" />
          </div>
          <div>
            <label className="text-xs text-gray-400 mb-1 block">รหัสผ่าน</label>
            <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="••••••"
              onKeyDown={e => e.key === 'Enter' && handleSubmit()}
              className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none" />
          </div>
          {error && (
            <div className="flex items-center gap-2 bg-red-900/40 border border-red-700 rounded-lg p-2.5">
              <AlertCircle size={14} className="text-red-400 flex-shrink-0" />
              <p className="text-red-300 text-sm">{error}</p>
            </div>
          )}
          <button onClick={handleSubmit} disabled={loading || !email || !password}
            className="w-full py-2.5 bg-blue-600 hover:bg-blue-500 disabled:bg-gray-600 rounded-xl font-semibold transition mt-1">
            {loading ? 'กำลังดำเนินการ...' : mode === 'login' ? 'เข้าสู่ระบบ' : 'สร้างบัญชี'}
          </button>
          <button onClick={() => { setMode(m => m === 'login' ? 'register' : 'login'); setError(''); }}
            className="w-full text-center text-sm text-gray-400 hover:text-gray-200 transition py-1">
            {mode === 'login' ? 'ยังไม่มีบัญชี? สร้างบัญชีใหม่' : 'มีบัญชีแล้ว? เข้าสู่ระบบ'}
          </button>
        </div>
      </div>
    </div>
  );
}

// ---- Rehearsal Mode Modal ----
function RehearsalModal({ song, onClose, theme }) {
  const [sections, setSections] = useState(song.sections || []);
  const [loopSection, setLoopSection] = useState(null);
  const [isLooping, setIsLooping] = useState(false);
  const [loopCount, setLoopCount] = useState(0);
  const [maxLoops, setMaxLoops] = useState(4);
  const [newSectionName, setNewSectionName] = useState('');
  const [newSectionStart, setNewSectionStart] = useState('');
  const [newSectionEnd, setNewSectionEnd] = useState('');
  const [editMode, setEditMode] = useState(false);
  const [zoomLevel, setZoomLevel] = useState(100);
  const [highlightedSection, setHighlightedSection] = useState(null);
  const loopRef = useRef(null);

  // Default section names
  const SECTION_PRESETS = ['Intro', 'Verse 1', 'Pre-Chorus', 'Chorus', 'Verse 2', 'Bridge', 'Solo', 'Outro', 'Hook'];

  const chordLines = (song.chordText || '').split('\n');

  const addSection = () => {
    if (!newSectionName.trim()) return;
    const start = parseInt(newSectionStart) || 0;
    const end = parseInt(newSectionEnd) || chordLines.length - 1;
    const newSec = { id: Date.now(), name: newSectionName.trim(), start, end, color: SECTION_COLORS[sections.length % SECTION_COLORS.length] };
    const updated = [...sections, newSec].sort((a, b) => a.start - b.start);
    setSections(updated);
    setNewSectionName(''); setNewSectionStart(''); setNewSectionEnd('');
  };

  const removeSection = (id) => setSections(s => s.filter(x => x.id !== id));

  const saveSections = async () => {
    await updateDoc(doc(db, 'songs', song.id), { sections });
    onClose();
  };

  const startLoop = (sec) => {
    setLoopSection(sec); setIsLooping(true); setLoopCount(0);
    setHighlightedSection(sec.id);
    const el = document.getElementById(`line-${sec.start}`);
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  };

  const stopLoop = () => { setIsLooping(false); setLoopSection(null); setHighlightedSection(null); };

  // Loop counter (visual only — counts based on time if bpm available)
  useEffect(() => {
    if (!isLooping || !loopSection || !song.bpm) return;
    const lineCount = loopSection.end - loopSection.start + 1;
    const beatsPerLine = 4;
    const msPerBeat = (60000 / parseInt(song.bpm || 120));
    const loopDuration = lineCount * beatsPerLine * msPerBeat;
    loopRef.current = setInterval(() => {
      setLoopCount(c => {
        if (c + 1 >= maxLoops) { stopLoop(); return 0; }
        return c + 1;
      });
    }, loopDuration);
    return () => clearInterval(loopRef.current);
  }, [isLooping, loopSection, maxLoops, song.bpm]);

  const SECTION_COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4', '#f97316'];

  const getLineSection = (lineIdx) => sections.find(s => lineIdx >= s.start && lineIdx <= s.end);

  const renderChordLine = (line, lineIdx) => {
    const chordRegex = /([^a-zA-Z0-9]|^)([CDEFGAB][#b]?(?:m|maj|min|dim|aug|sus)?\d*(?:\/[CDEFGAB][#b]?)?)(?=[^a-zA-Z0-9]|$)/g;
    const sec = getLineSection(lineIdx);
    const isHighlighted = highlightedSection && sec?.id === highlightedSection;
    const parts = []; let lastIndex = 0; let match;
    const regex = new RegExp(chordRegex);
    while ((match = regex.exec(line)) !== null) {
      const before = line.substring(lastIndex, match.index + match[1].length);
      if (before) parts.push(before);
      parts.push(<span key={match.index} className="font-bold px-1 rounded mx-0.5 inline-block" style={{ color: theme.chordColor, background: theme.chordColor + '22', border: `1px solid ${theme.chordColor}44` }}>{match[2]}</span>);
      lastIndex = match.index + match[0].length;
    }
    const after = line.substring(lastIndex);
    if (after) parts.push(after);

    return (
      <div key={lineIdx} id={`line-${lineIdx}`}
        className={`whitespace-pre-wrap break-words px-2 py-0.5 rounded transition-all ${isHighlighted ? 'ring-1' : ''}`}
        style={{
          background: isHighlighted ? sec.color + '22' : (sec ? sec.color + '11' : 'transparent'),
          borderLeft: sec ? `3px solid ${sec.color}` : '3px solid transparent',
          ringColor: isHighlighted ? sec?.color : undefined,
        }}>
        {editMode && (
          <span className="text-xs text-gray-600 select-none mr-2 font-mono">{lineIdx + 1}</span>
        )}
        {parts.length > 0 ? parts : (line || '\u00a0')}
      </div>
    );
  };

  return (
    <div className="fixed inset-0 z-50 flex flex-col" style={{ background: theme.bg }}>
      {/* Header */}
      <div className="flex items-center justify-between px-4 py-3 border-b border-gray-800 flex-shrink-0">
        <div className="flex items-center gap-3">
          <button onClick={onClose} className="text-gray-400 hover:text-white"><ArrowLeft size={20} /></button>
          <div>
            <h2 className="text-base font-bold" style={{ color: theme.chordColor }}>{song.title}</h2>
            <p className="text-xs text-gray-500">Rehearsal Mode — Key {song.key} {song.bpm && `| ${song.bpm} BPM`}</p>
          </div>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={() => setZoomLevel(z => Math.max(60, z - 10))} className="p-1.5 bg-gray-800 rounded-lg hover:bg-gray-700"><ZoomOut size={15} /></button>
          <span className="text-xs text-gray-400 font-mono w-10 text-center">{zoomLevel}%</span>
          <button onClick={() => setZoomLevel(z => Math.min(200, z + 10))} className="p-1.5 bg-gray-800 rounded-lg hover:bg-gray-700"><ZoomIn size={15} /></button>
          <button onClick={() => setEditMode(e => !e)} className={`px-3 py-1.5 rounded-lg text-sm font-semibold ${editMode ? 'bg-yellow-600 text-white' : 'bg-gray-800 text-gray-400 hover:bg-gray-700'}`}>
            <Scissors size={14} className="inline mr-1" />{editMode ? 'โหมดตัดท่อน' : 'จัดท่อน'}
          </button>
          <button onClick={saveSections} className="px-3 py-1.5 rounded-lg text-sm font-semibold bg-blue-600 hover:bg-blue-500 text-white">
            <Save size={14} className="inline mr-1" />บันทึก
          </button>
        </div>
      </div>

      <div className="flex flex-1 overflow-hidden">
        {/* Chord content */}
        <div className="flex-1 overflow-y-auto p-4">
          {!song.chordText ? (
            <div className="flex items-center justify-center h-full text-gray-600">
              <p>ยังไม่มีข้อความคอร์ด — กรอกคอร์ดก่อนใช้ Rehearsal Mode</p>
            </div>
          ) : (
            <div style={{ fontSize: `${zoomLevel}%`, color: theme.textColor }} className="leading-relaxed tracking-wide font-medium">
              {chordLines.map((line, i) => renderChordLine(line, i))}
            </div>
          )}
        </div>

        {/* Right Panel: Sections */}
        <div className="w-64 flex-shrink-0 border-l border-gray-800 flex flex-col bg-gray-900/50 overflow-hidden">
          <div className="p-3 border-b border-gray-800 flex-shrink-0">
            <p className="text-xs text-gray-400 font-semibold mb-2 flex items-center gap-1"><Repeat size={12} /> ท่อนเพลง (Sections)</p>

            {/* Loop Status */}
            {isLooping && loopSection && (
              <div className="mb-2 p-2 rounded-lg border text-xs" style={{ borderColor: loopSection.color, background: loopSection.color + '22' }}>
                <div className="flex justify-between items-center">
                  <span style={{ color: loopSection.color }} className="font-bold">🔁 {loopSection.name}</span>
                  <button onClick={stopLoop} className="text-gray-400 hover:text-red-400"><Square size={12} /></button>
                </div>
                <div className="mt-1.5 flex gap-1">
                  {Array.from({ length: maxLoops }).map((_, i) => (
                    <div key={i} className="flex-1 h-1.5 rounded-full" style={{ background: i < loopCount ? loopSection.color : '#374151' }} />
                  ))}
                </div>
                <p className="text-gray-400 mt-1">{loopCount}/{maxLoops} รอบ</p>
              </div>
            )}

            {/* Max loops control */}
            <div className="flex items-center gap-2 mb-2">
              <span className="text-xs text-gray-400">ลูป:</span>
              {[2, 4, 8, 16].map(n => (
                <button key={n} onClick={() => setMaxLoops(n)} className={`text-xs px-1.5 py-0.5 rounded ${maxLoops === n ? 'bg-blue-700 text-white' : 'bg-gray-700 text-gray-400'}`}>{n}x</button>
              ))}
            </div>
          </div>

          {/* Section list */}
          <div className="flex-1 overflow-y-auto p-2 space-y-1.5">
            {sections.length === 0 && !editMode && (
              <p className="text-xs text-gray-600 text-center py-4">กดปุ่ม "จัดท่อน"<br />เพื่อเพิ่มท่อน</p>
            )}
            {sections.map(sec => (
              <div key={sec.id} className="rounded-lg p-2 border text-xs" style={{ borderColor: sec.color + '66', background: sec.color + '11' }}>
                <div className="flex items-center justify-between">
                  <span className="font-bold" style={{ color: sec.color }}>{sec.name}</span>
                  <div className="flex gap-1">
                    <button onClick={() => { setHighlightedSection(sec.id); const el = document.getElementById(`line-${sec.start}`); el?.scrollIntoView({ behavior: 'smooth', block: 'center' }); }}
                      className="text-gray-500 hover:text-white p-0.5"><Eye size={11} /></button>
                    {!isLooping ? (
                      <button onClick={() => startLoop(sec)} className="text-gray-500 hover:text-green-400 p-0.5"><Play size={11} /></button>
                    ) : loopSection?.id === sec.id ? (
                      <button onClick={stopLoop} className="text-red-400 p-0.5"><Square size={11} /></button>
                    ) : null}
                    {editMode && <button onClick={() => removeSection(sec.id)} className="text-gray-600 hover:text-red-400 p-0.5"><X size={11} /></button>}
                  </div>
                </div>
                <p className="text-gray-500 mt-0.5">บรรทัด {sec.start + 1}–{sec.end + 1}</p>
              </div>
            ))}
          </div>

          {/* Add section form */}
          {editMode && (
            <div className="p-3 border-t border-gray-800 flex-shrink-0 space-y-2">
              <p className="text-xs text-gray-400 font-semibold">เพิ่มท่อนใหม่</p>
              <div className="flex flex-wrap gap-1">
                {SECTION_PRESETS.filter(p => !sections.find(s => s.name === p)).slice(0, 6).map(p => (
                  <button key={p} onClick={() => setNewSectionName(p)} className="text-xs px-2 py-0.5 bg-gray-700 hover:bg-gray-600 rounded-full text-gray-300">{p}</button>
                ))}
              </div>
              <input value={newSectionName} onChange={e => setNewSectionName(e.target.value)} placeholder="ชื่อท่อน (เช่น Chorus)"
                className="w-full bg-gray-900 border border-gray-700 rounded-lg px-2 py-1.5 text-white text-xs outline-none focus:border-blue-500" />
              <div className="flex gap-1.5">
                <input value={newSectionStart} onChange={e => setNewSectionStart(e.target.value)} placeholder="บรรทัดเริ่ม" type="number"
                  className="flex-1 bg-gray-900 border border-gray-700 rounded-lg px-2 py-1.5 text-white text-xs outline-none" />
                <input value={newSectionEnd} onChange={e => setNewSectionEnd(e.target.value)} placeholder="บรรทัดจบ" type="number"
                  className="flex-1 bg-gray-900 border border-gray-700 rounded-lg px-2 py-1.5 text-white text-xs outline-none" />
              </div>
              <button onClick={addSection} disabled={!newSectionName.trim()}
                className="w-full py-1.5 bg-blue-600 hover:bg-blue-500 disabled:bg-gray-700 rounded-lg text-xs font-semibold text-white">
                + เพิ่มท่อน
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ---- PDF Export Modal ----
function PDFExportModal({ songs, playlist, onClose }) {
  const [includeChords, setIncludeChords] = useState(false);
  const [selectedIds, setSelectedIds] = useState(new Set(songs.map(s => s.id)));

  const toggleSong = (id) => setSelectedIds(prev => {
    const next = new Set(prev);
    if (next.has(id)) next.delete(id); else next.add(id);
    return next;
  });

  const selectedSongs = songs.filter(s => selectedIds.has(s.id));
  const totalSecs = selectedSongs.reduce((a, s) => a + parseDuration(s.duration), 0);

  return (
    <div className="fixed inset-0 bg-black/85 z-50 flex items-center justify-center p-4" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="bg-gray-800 rounded-2xl w-full max-w-lg border border-gray-700 shadow-2xl max-h-[90vh] flex flex-col">
        <div className="flex justify-between items-center p-5 border-b border-gray-700 flex-shrink-0">
          <h2 className="text-xl font-bold flex items-center gap-2">
            <Printer size={20} className="text-blue-400" /> Export PDF / พิมพ์แจก
          </h2>
          <button onClick={onClose}><X size={20} className="text-gray-400 hover:text-white" /></button>
        </div>

        <div className="flex-1 overflow-y-auto p-5 space-y-4">
          {/* Options */}
          <div className="bg-gray-900 rounded-xl p-4 space-y-3">
            <p className="text-sm font-semibold text-gray-300">ตัวเลือก</p>
            <label className="flex items-center gap-3 cursor-pointer">
              <div onClick={() => setIncludeChords(c => !c)}
                className={`w-10 h-6 rounded-full transition-colors ${includeChords ? 'bg-blue-600' : 'bg-gray-700'} relative`}>
                <div className={`absolute top-1 w-4 h-4 bg-white rounded-full transition-all ${includeChords ? 'left-5' : 'left-1'}`} />
              </div>
              <div>
                <p className="text-sm text-gray-200">รวมเนื้อเพลง/คอร์ด</p>
                <p className="text-xs text-gray-500">แสดงคอร์ดในหน้า PDF (ไฟล์จะใหญ่ขึ้น)</p>
              </div>
            </label>
          </div>

          {/* Song selection */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <p className="text-sm font-semibold text-gray-300">เลือกเพลง ({selectedIds.size}/{songs.length})</p>
              <div className="flex gap-2 text-xs">
                <button onClick={() => setSelectedIds(new Set(songs.map(s => s.id)))} className="text-blue-400 hover:text-blue-300">เลือกทั้งหมด</button>
                <button onClick={() => setSelectedIds(new Set())} className="text-gray-400 hover:text-gray-300">ยกเลิกทั้งหมด</button>
              </div>
            </div>
            <div className="space-y-1.5 max-h-52 overflow-y-auto">
              {songs.map((song, i) => (
                <label key={song.id} className={`flex items-center gap-3 p-2.5 rounded-xl cursor-pointer transition ${selectedIds.has(song.id) ? 'bg-blue-900/30 border border-blue-800' : 'bg-gray-900 border border-gray-700 opacity-50'}`}>
                  <input type="checkbox" checked={selectedIds.has(song.id)} onChange={() => toggleSong(song.id)} className="accent-blue-500 w-4 h-4" />
                  <span className="text-gray-500 text-sm w-5 text-center">{i + 1}</span>
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-semibold text-white truncate">{song.title}</p>
                    <p className="text-xs text-gray-400">{song.artist}</p>
                  </div>
                  <span className="text-xs font-mono text-blue-300">{song.key}</span>
                  {song.duration && <span className="text-xs text-yellow-600">{song.duration}</span>}
                </label>
              ))}
            </div>
          </div>

          {/* Summary */}
          <div className="bg-gray-900 rounded-xl p-3 flex items-center justify-between">
            <span className="text-sm text-gray-400">{selectedIds.size} เพลง • {formatDuration(totalSecs)}</span>
            {includeChords && <span className="text-xs text-yellow-500 flex items-center gap-1"><AlertCircle size={12} /> รวมคอร์ด</span>}
          </div>
        </div>

        <div className="flex gap-3 p-5 border-t border-gray-700 flex-shrink-0">
          <button onClick={onClose} className="flex-1 py-2.5 bg-gray-700 hover:bg-gray-600 rounded-xl font-semibold">ยกเลิก</button>
          <button
            onClick={() => { exportSetlistPDF(selectedSongs, playlist, includeChords); onClose(); }}
            disabled={selectedIds.size === 0}
            className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-500 disabled:bg-gray-600 rounded-xl font-semibold flex items-center justify-center gap-2">
            <Printer size={16} /> เปิด PDF / พิมพ์
          </button>
        </div>
      </div>
    </div>
  );
}

// ---- Import Modal ----
function ImportModal({ onClose, onImport, playlists }) {
  const [step, setStep] = useState('upload');
  const [previewRows, setPreviewRows] = useState([]);
  const [errors, setErrors] = useState([]);
  const [targetPlaylistId, setTargetPlaylistId] = useState(playlists[0]?.id || '');
  const [progress, setProgress] = useState(0);
  const [isDrag, setIsDrag] = useState(false);

  const handleFile = async (file) => {
    try { const rows = await parseImportFile(file); setPreviewRows(rows); setErrors([]); setStep('preview'); }
    catch (e) { setErrors([e.message]); }
  };

  const handleImport = async () => {
    setStep('importing');
    const q = query(collection(db, 'songs'), where('playlistId', '==', targetPlaylistId), orderBy('order', 'desc'));
    const snap = await getDocs(q);
    const startOrder = snap.empty ? 0 : (snap.docs[0].data().order || 0) + 1;
    let count = 0;
    for (const [i, song] of previewRows.entries()) {
      await addDoc(collection(db, 'songs'), { ...song, order: startOrder + i, playlistId: targetPlaylistId });
      count++;
      setProgress(Math.round((count / previewRows.length) * 100));
    }
    setStep('done');
    onImport();
  };

  return (
    <div className="fixed inset-0 bg-black/85 z-50 flex items-center justify-center p-4" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="bg-gray-800 rounded-2xl w-full max-w-2xl border border-gray-700 shadow-2xl max-h-[90vh] flex flex-col">
        <div className="flex justify-between items-center p-5 border-b border-gray-700 flex-shrink-0">
          <h2 className="text-xl font-bold flex items-center gap-2"><ArrowUpFromLine size={20} className="text-green-400" />Import จาก Excel / CSV</h2>
          <button onClick={onClose}><X size={20} className="text-gray-400 hover:text-white" /></button>
        </div>
        <div className="flex-1 overflow-y-auto p-5">
          {step === 'upload' && (
            <div className="space-y-5">
              <div className="bg-blue-900/30 border border-blue-700 rounded-xl p-4 flex items-start gap-3">
                <Table size={20} className="text-blue-400 flex-shrink-0 mt-0.5" />
                <div className="flex-1">
                  <p className="text-blue-200 font-semibold text-sm">ขั้นตอนที่ 1: ดาวน์โหลด Template</p>
                  <p className="text-blue-400 text-xs mt-1">แจกให้สมาชิกไปกรอก แล้ว Import กลับมา</p>
                  <button onClick={downloadTemplate} className="mt-3 flex items-center gap-2 px-4 py-2 bg-blue-700 hover:bg-blue-600 rounded-lg text-sm font-semibold text-white">
                    <Download size={14} /> ดาวน์โหลด Template (.xlsx)
                  </button>
                </div>
              </div>
              <div>
                <p className="text-sm text-gray-300 font-semibold mb-2">ขั้นตอนที่ 2: อัปโหลดไฟล์</p>
                <div className={`border-2 border-dashed rounded-xl p-10 text-center transition-all ${isDrag ? 'border-green-500 bg-green-900/20' : 'border-gray-600 hover:border-gray-400'}`}
                  onDragOver={e => { e.preventDefault(); setIsDrag(true); }} onDragLeave={() => setIsDrag(false)}
                  onDrop={e => { e.preventDefault(); setIsDrag(false); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); }}>
                  <ArrowUpFromLine size={36} className={`mx-auto mb-3 ${isDrag ? 'text-green-400' : 'text-gray-500'}`} />
                  <p className="text-gray-300 font-semibold">ลากไฟล์มาวางที่นี่</p>
                  <label className="mt-3 inline-flex items-center gap-2 px-4 py-2 bg-gray-700 hover:bg-gray-600 rounded-lg text-sm cursor-pointer font-semibold">
                    <Upload size={14} /> เลือกไฟล์ .xlsx / .csv
                    <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={e => e.target.files[0] && handleFile(e.target.files[0])} />
                  </label>
                </div>
              </div>
              {errors.length > 0 && (
                <div className="bg-red-900/40 border border-red-700 rounded-xl p-3 flex items-start gap-2">
                  <AlertCircle size={16} className="text-red-400 flex-shrink-0 mt-0.5" /><p className="text-red-300 text-sm">{errors[0]}</p>
                </div>
              )}
            </div>
          )}
          {step === 'preview' && (
            <div className="space-y-4">
              <div className="flex items-center justify-between">
                <p className="text-green-400 font-semibold flex items-center gap-2"><CheckSquare size={16} /> พบ {previewRows.length} เพลง</p>
                <button onClick={() => setStep('upload')} className="text-xs text-gray-400 hover:text-white flex items-center gap-1"><ArrowLeft size={12} />เลือกไฟล์ใหม่</button>
              </div>
              <div>
                <label className="text-xs text-gray-400 mb-1 block">Import เข้า Playlist</label>
                <select value={targetPlaylistId} onChange={e => setTargetPlaylistId(e.target.value)}
                  className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white text-sm outline-none focus:border-blue-500">
                  {playlists.map(p => <option key={p.id} value={p.id}>{p.icon} {p.name}</option>)}
                </select>
              </div>
              <div className="overflow-x-auto rounded-xl border border-gray-700">
                <table className="w-full text-xs">
                  <thead className="bg-gray-900 text-gray-400">
                    <tr>{['ชื่อเพลง', 'ศิลปิน', 'Key', 'BPM', 'ความยาว', 'แท็ก'].map(h => <th key={h} className="text-left px-3 py-2 whitespace-nowrap">{h}</th>)}</tr>
                  </thead>
                  <tbody>
                    {previewRows.map((r, i) => (
                      <tr key={i} className={`border-t border-gray-700 ${i % 2 === 0 ? 'bg-gray-800/50' : ''}`}>
                        <td className="px-3 py-2 font-semibold text-white">{r.title}</td>
                        <td className="px-3 py-2 text-gray-300">{r.artist}</td>
                        <td className="px-3 py-2 font-mono text-blue-300">{r.key}</td>
                        <td className="px-3 py-2 text-gray-400">{r.bpm}</td>
                        <td className="px-3 py-2 text-yellow-600">{r.duration}</td>
                        <td className="px-3 py-2"><div className="flex flex-wrap gap-1">{r.tags.map((t, ti) => <TagBadge key={ti} tag={t} small />)}</div></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
          {step === 'importing' && (
            <div className="py-10 text-center space-y-4">
              <div className="w-16 h-16 mx-auto rounded-full border-4 border-blue-500 border-t-transparent animate-spin" />
              <p className="text-white font-semibold">กำลัง Import... {progress}%</p>
              <div className="h-2 bg-gray-700 rounded-full overflow-hidden mx-8">
                <div className="h-full bg-blue-500 rounded-full transition-all" style={{ width: `${progress}%` }} />
              </div>
            </div>
          )}
          {step === 'done' && (
            <div className="py-10 text-center space-y-3">
              <div className="w-16 h-16 mx-auto rounded-full bg-green-700 flex items-center justify-center"><CheckSquare size={32} className="text-white" /></div>
              <p className="text-white font-bold text-lg">Import สำเร็จ!</p>
              <p className="text-gray-400 text-sm">เพิ่ม {previewRows.length} เพลงเรียบร้อยแล้ว</p>
            </div>
          )}
        </div>
        <div className="flex gap-3 p-5 border-t border-gray-700 flex-shrink-0">
          {step === 'preview' && (
            <>
              <button onClick={() => setStep('upload')} className="flex-1 py-2.5 bg-gray-700 hover:bg-gray-600 rounded-xl font-semibold">ย้อนกลับ</button>
              <button onClick={handleImport} disabled={!targetPlaylistId} className="flex-1 py-2.5 bg-green-600 hover:bg-green-500 disabled:bg-gray-600 rounded-xl font-semibold flex items-center justify-center gap-2">
                <ArrowUpFromLine size={16} /> Import {previewRows.length} เพลง
              </button>
            </>
          )}
          {(step === 'upload' || step === 'done') && (
            <button onClick={onClose} className="flex-1 py-2.5 bg-gray-700 hover:bg-gray-600 rounded-xl font-semibold">{step === 'done' ? 'ปิด' : 'ยกเลิก'}</button>
          )}
        </div>
      </div>
    </div>
  );
}

// ---- Playlist Form ----
function PlaylistFormModal({ playlist, onClose, onSave }) {
  const [name, setName] = useState(playlist?.name || '');
  const [description, setDescription] = useState(playlist?.description || '');
  const [color, setColor] = useState(playlist?.color || '#3b82f6');
  const [icon, setIcon] = useState(playlist?.icon || '🎵');
  const ICONS = ['🎵', '🎸', '🥁', '🎹', '🎺', '🎻', '🎤', '🎧', '⭐', '🔥', '💫', '🎭', '🎪', '🏟️', '🌙', '☀️'];
  const COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4', '#f97316'];
  return (
    <div className="fixed inset-0 bg-black/85 z-50 flex items-center justify-center p-4" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="bg-gray-800 rounded-2xl w-full max-w-md border border-gray-700 shadow-2xl">
        <div className="flex justify-between items-center p-5 border-b border-gray-700">
          <h2 className="text-xl font-bold flex items-center gap-2"><FolderPlus size={20} className="text-blue-400" />{playlist ? 'แก้ไข Playlist' : 'สร้าง Playlist ใหม่'}</h2>
          <button onClick={onClose}><X size={20} className="text-gray-400 hover:text-white" /></button>
        </div>
        <div className="p-5 space-y-4">
          <div className="flex items-center gap-3 p-3 rounded-xl border border-gray-700" style={{ borderColor: color + '66' }}>
            <div className="w-10 h-10 rounded-xl flex items-center justify-center text-xl flex-shrink-0" style={{ background: color + '33' }}>{icon}</div>
            <div><p className="font-bold text-white">{name || 'ชื่อ Playlist'}</p><p className="text-xs text-gray-400">{description || 'คำอธิบาย...'}</p></div>
          </div>
          <div><label className="text-xs text-gray-400 mb-1 block">ชื่อ *</label>
            <input autoFocus value={name} onChange={e => setName(e.target.value)} placeholder="เช่น งานแต่งงาน, ซ้อมประจำ" className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none" /></div>
          <div><label className="text-xs text-gray-400 mb-1 block">คำอธิบาย</label>
            <input value={description} onChange={e => setDescription(e.target.value)} placeholder="รายละเอียด..." className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none" /></div>
          <div><label className="text-xs text-gray-400 mb-2 block">ไอคอน</label>
            <div className="flex flex-wrap gap-2">{ICONS.map(ic => <button key={ic} onClick={() => setIcon(ic)} className={`w-9 h-9 rounded-lg text-lg flex items-center justify-center ${icon === ic ? 'bg-blue-700 ring-2 ring-blue-400' : 'bg-gray-700 hover:bg-gray-600'}`}>{ic}</button>)}</div></div>
          <div><label className="text-xs text-gray-400 mb-2 block">สี</label>
            <div className="flex gap-2">{COLORS.map(c => <button key={c} onClick={() => setColor(c)} className={`w-8 h-8 rounded-full transition ${color === c ? 'ring-2 ring-white scale-110' : 'hover:scale-105'}`} style={{ background: c }} />)}</div></div>
        </div>
        <div className="flex gap-3 p-5 border-t border-gray-700">
          <button onClick={onClose} className="flex-1 py-2.5 bg-gray-700 hover:bg-gray-600 rounded-xl font-semibold">ยกเลิก</button>
          <button onClick={() => { onSave({ name, description, color, icon }); onClose(); }} disabled={!name.trim()} className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-500 disabled:bg-gray-600 rounded-xl font-semibold">บันทึก</button>
        </div>
      </div>
    </div>
  );
}

// ---- AddSongModal ----
function AddSongModal({ onClose, onSave, allSongs }) {
  const [title, setTitle] = useState('');
  const [artist, setArtist] = useState('');
  const [key, setKey] = useState('C');
  const [bpm, setBpm] = useState('120');
  const [duration, setDuration] = useState('03:30');
  const [tagInput, setTagInput] = useState('');
  const [tags, setTags] = useState([]);
  const allTags = [...new Set(allSongs.flatMap(s => s.tags || []))];
  const addTag = (t) => { const c = t.trim(); if (c && !tags.includes(c)) setTags([...tags, c]); setTagInput(''); };
  return (
    <div className="fixed inset-0 bg-black/80 z-50 flex items-center justify-center p-4" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="bg-gray-800 rounded-2xl w-full max-w-lg border border-gray-700 shadow-2xl">
        <div className="flex justify-between items-center p-5 border-b border-gray-700">
          <h2 className="text-xl font-bold flex items-center gap-2"><Music size={20} className="text-blue-400" /> เพิ่มเพลงใหม่</h2>
          <button onClick={onClose}><X size={20} className="text-gray-400 hover:text-white" /></button>
        </div>
        <div className="p-5 space-y-4">
          <div><label className="text-xs text-gray-400 mb-1 block">ชื่อเพลง *</label>
            <input autoFocus value={title} onChange={e => setTitle(e.target.value)} placeholder="เช่น สักวันหนึ่ง" className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none" /></div>
          <div><label className="text-xs text-gray-400 mb-1 block">ศิลปิน</label>
            <input value={artist} onChange={e => setArtist(e.target.value)} placeholder="เช่น Bodyslam" className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none" /></div>
          <div className="grid grid-cols-3 gap-3">
            <div><label className="text-xs text-gray-400 mb-1 block">คีย์</label>
              <select value={key} onChange={e => setKey(e.target.value)} className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none">
                {[...NOTES, 'Cm','Dm','Em','Fm','Gm','Am','Bm'].map(n => <option key={n}>{n}</option>)}
              </select></div>
            <div><label className="text-xs text-gray-400 mb-1 block">BPM</label>
              <input value={bpm} onChange={e => setBpm(e.target.value)} type="number" className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white outline-none" /></div>
            <div><label className="text-xs text-gray-400 mb-1 block">ความยาว</label>
              <input value={duration} onChange={e => setDuration(e.target.value)} placeholder="03:30" className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white outline-none" /></div>
          </div>
          <div>
            <label className="text-xs text-gray-400 mb-1 block">แท็ก</label>
            <div className="flex gap-2">
              <input value={tagInput} onChange={e => setTagInput(e.target.value)} onKeyDown={e => { if (e.key === 'Enter' || e.key === ',') { e.preventDefault(); addTag(tagInput); } }} placeholder="พิมพ์แล้วกด Enter" className="flex-1 bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none text-sm" />
              <button onClick={() => addTag(tagInput)} className="px-3 py-2 bg-gray-700 hover:bg-gray-600 rounded-lg text-sm">เพิ่ม</button>
            </div>
            {allTags.filter(t => !tags.includes(t)).length > 0 && <div className="flex flex-wrap gap-1 mt-2">{allTags.filter(t => !tags.includes(t)).map(t => <button key={t} onClick={() => addTag(t)} className={`text-xs px-2 py-0.5 rounded-full border opacity-60 hover:opacity-100 ${getTagColor(t).bg} ${getTagColor(t).border} ${getTagColor(t).text}`}>+ {t}</button>)}</div>}
            {tags.length > 0 && <div className="flex flex-wrap gap-1 mt-2">{tags.map(t => <TagBadge key={t} tag={t} onRemove={t2 => setTags(tags.filter(x => x !== t2))} />)}</div>}
          </div>
        </div>
        <div className="flex gap-3 p-5 border-t border-gray-700">
          <button onClick={onClose} className="flex-1 py-2.5 bg-gray-700 hover:bg-gray-600 rounded-xl font-semibold">ยกเลิก</button>
          <button onClick={() => { if (!title.trim()) return; onSave({ title: title.trim(), artist: artist.trim() || 'ไม่ระบุ', key, bpm, duration, tags, imageUrl: null, chordText: null, links: [], notes: '', sharedCues: '', history: [], sections: [] }); onClose(); }} disabled={!title.trim()} className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-500 disabled:bg-gray-600 rounded-xl font-semibold">บันทึก</button>
        </div>
      </div>
    </div>
  );
}

// ---- EditSongModal ----
function EditSongModal({ song, onClose, onSave, allSongs }) {
  const [title, setTitle] = useState(song.title || '');
  const [artist, setArtist] = useState(song.artist || '');
  const [key, setKey] = useState(song.key || 'C');
  const [bpm, setBpm] = useState(song.bpm || '120');
  const [duration, setDuration] = useState(song.duration || '03:30');
  const [tags, setTags] = useState(song.tags || []);
  const [tagInput, setTagInput] = useState('');
  const [links, setLinks] = useState(song.links || []);
  const [linkInput, setLinkInput] = useState('');
  const [linkLabel, setLinkLabel] = useState('');
  const [notes, setNotes] = useState(song.notes || '');
  const [sharedCues, setSharedCues] = useState(song.sharedCues || '');
  const [activeTab, setActiveTab] = useState('info');
  const allTags = [...new Set(allSongs.flatMap(s => s.tags || []))];
  const addTag = (t) => { const c = t.trim(); if (c && !tags.includes(c)) setTags([...tags, c]); setTagInput(''); };
  const addLink = () => { if (!linkInput.trim()) return; setLinks([...links, { url: linkInput.trim(), label: linkLabel.trim() || linkInput.trim() }]); setLinkInput(''); setLinkLabel(''); };
  const TABS = [{ id: 'info', label: 'ข้อมูล' }, { id: 'tags', label: 'แท็ก' }, { id: 'links', label: 'ลิงก์' }, { id: 'notes', label: 'โน้ต' }];
  return (
    <div className="fixed inset-0 bg-black/80 z-50 flex items-center justify-center p-4" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="bg-gray-800 rounded-2xl w-full max-w-lg border border-gray-700 shadow-2xl max-h-[90vh] flex flex-col">
        <div className="flex justify-between items-center p-5 border-b border-gray-700 flex-shrink-0">
          <h2 className="text-lg font-bold flex items-center gap-2"><Edit2 size={17} className="text-yellow-400" /> {song.title}</h2>
          <button onClick={onClose}><X size={20} className="text-gray-400 hover:text-white" /></button>
        </div>
        <div className="flex border-b border-gray-700 flex-shrink-0">
          {TABS.map(t => <button key={t.id} onClick={() => setActiveTab(t.id)} className={`flex-1 py-2.5 text-sm font-semibold transition ${activeTab === t.id ? 'text-blue-400 border-b-2 border-blue-400' : 'text-gray-400 hover:text-gray-200'}`}>{t.label}</button>)}
        </div>
        <div className="flex-1 overflow-y-auto p-5">
          {activeTab === 'info' && (
            <div className="space-y-3">
              <div><label className="text-xs text-gray-400 mb-1 block">ชื่อเพลง</label><input value={title} onChange={e => setTitle(e.target.value)} className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none" /></div>
              <div><label className="text-xs text-gray-400 mb-1 block">ศิลปิน</label><input value={artist} onChange={e => setArtist(e.target.value)} className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none" /></div>
              <div className="grid grid-cols-3 gap-3">
                <div><label className="text-xs text-gray-400 mb-1 block">คีย์</label>
                  <select value={key} onChange={e => setKey(e.target.value)} className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none">
                    {[...NOTES, 'Cm','Dm','Em','Fm','Gm','Am','Bm'].map(n => <option key={n}>{n}</option>)}
                  </select></div>
                <div><label className="text-xs text-gray-400 mb-1 block">BPM</label><input value={bpm} onChange={e => setBpm(e.target.value)} type="number" className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white outline-none" /></div>
                <div><label className="text-xs text-gray-400 mb-1 block">ความยาว</label><input value={duration} onChange={e => setDuration(e.target.value)} placeholder="03:30" className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white outline-none" /></div>
              </div>
            </div>
          )}
          {activeTab === 'tags' && (
            <div className="space-y-3">
              <div className="flex gap-2"><input value={tagInput} onChange={e => setTagInput(e.target.value)} onKeyDown={e => { if (e.key === 'Enter' || e.key === ',') { e.preventDefault(); addTag(tagInput); } }} placeholder="พิมพ์แล้วกด Enter" className="flex-1 bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white focus:border-blue-500 outline-none text-sm" /><button onClick={() => addTag(tagInput)} className="px-3 py-2 bg-gray-700 hover:bg-gray-600 rounded-lg text-sm">เพิ่ม</button></div>
              {allTags.filter(t => !tags.includes(t)).length > 0 && <div className="flex flex-wrap gap-1">{allTags.filter(t => !tags.includes(t)).map(t => <button key={t} onClick={() => addTag(t)} className={`text-xs px-2 py-0.5 rounded-full border opacity-60 hover:opacity-100 ${getTagColor(t).bg} ${getTagColor(t).border} ${getTagColor(t).text}`}>+ {t}</button>)}</div>}
              <div className="flex flex-wrap gap-1 min-h-8">{tags.length === 0 ? <span className="text-gray-600 text-sm">ยังไม่มีแท็ก</span> : tags.map(t => <TagBadge key={t} tag={t} onRemove={t2 => setTags(tags.filter(x => x !== t2))} />)}</div>
            </div>
          )}
          {activeTab === 'links' && (
            <div className="space-y-3">
              <input value={linkLabel} onChange={e => setLinkLabel(e.target.value)} placeholder="ชื่อลิงก์ (เช่น YouTube)" className="w-full bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white outline-none text-sm" />
              <div className="flex gap-2"><input value={linkInput} onChange={e => setLinkInput(e.target.value)} onKeyDown={e => e.key === 'Enter' && addLink()} placeholder="https://..." className="flex-1 bg-gray-900 border border-gray-600 rounded-lg px-3 py-2 text-white outline-none text-sm" /><button onClick={addLink} className="px-3 py-2 bg-blue-700 hover:bg-blue-600 rounded-lg text-sm">เพิ่ม</button></div>
              <div className="space-y-2">{links.length === 0 && <p className="text-gray-600 text-sm">ยังไม่มีลิงก์</p>}{links.map((lk, i) => <div key={i} className="flex items-center gap-2 bg-gray-900 rounded-lg p-2 border border-gray-700"><Link size={12} className="text-blue-400 flex-shrink-0" /><a href={lk.url} target="_blank" rel="noopener noreferrer" className="flex-1 text-blue-300 hover:text-blue-100 text-sm truncate">{lk.label}</a><button onClick={() => setLinks(links.filter((_, idx) => idx !== i))}><X size={14} className="text-gray-500 hover:text-red-400" /></button></div>)}</div>
            </div>
          )}
          {activeTab === 'notes' && (
            <div className="space-y-3">
              <div><label className="text-xs text-gray-400 mb-1 flex items-center gap-1"><Eye size={11} /> Shared Cues (ทั้งวงเห็น)</label><textarea value={sharedCues} onChange={e => setSharedCues(e.target.value)} rows={3} placeholder="เช่น โซโล่กีตาร์ 2 รอบ..." className="w-full bg-gray-900 border border-blue-800 rounded-lg px-3 py-2 text-white outline-none text-sm resize-none" /></div>
              <div><label className="text-xs text-gray-400 mb-1 flex items-center gap-1"><EyeOff size={11} /> Personal Notes (เห็นแค่ตัวเอง)</label><textarea value={notes} onChange={e => setNotes(e.target.value)} rows={3} placeholder="เช่น เหยียบ delay ท่อนฮุค..." className="w-full bg-gray-900 border border-gray-700 rounded-lg px-3 py-2 text-white outline-none text-sm resize-none" /></div>
            </div>
          )}
        </div>
        <div className="flex gap-3 p-5 border-t border-gray-700 flex-shrink-0">
          <button onClick={onClose} className="flex-1 py-2.5 bg-gray-700 hover:bg-gray-600 rounded-xl font-semibold">ยกเลิก</button>
          <button onClick={() => { onSave({ title, artist, key, bpm, duration, tags, links, notes, sharedCues }); onClose(); }} className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-500 rounded-xl font-semibold">บันทึก</button>
        </div>
      </div>
    </div>
  );
}

// ---- Settings Modal ----
function SettingsModal({ theme, onClose, onSave, geminiKeyCount, onManageGeminiKeys }) {
  const [bg, setBg] = useState(theme.bg || '#000000');
  const [textColor, setTextColor] = useState(theme.textColor || '#ffffff');
  const [chordColor, setChordColor] = useState(theme.chordColor || '#facc15');
  const [bgImage, setBgImage] = useState(theme.bgImage || '');
  const [bgOpacity, setBgOpacity] = useState(theme.bgOpacity ?? 0.15);
  const [preset, setPreset] = useState('custom');
  const presets = [
    { id: 'dark', label: 'Dark', bg: '#000000', textColor: '#ffffff', chordColor: '#facc15' },
    { id: 'navy', label: 'Navy', bg: '#0f172a', textColor: '#e2e8f0', chordColor: '#38bdf8' },
    { id: 'forest', label: 'Forest', bg: '#052e16', textColor: '#dcfce7', chordColor: '#86efac' },
    { id: 'warm', label: 'Warm', bg: '#1c0a00', textColor: '#fef3c7', chordColor: '#fb923c' },
    { id: 'purple', label: 'Purple', bg: '#1e1b4b', textColor: '#e0e7ff', chordColor: '#c084fc' },
  ];

  const handleBgImage = (file) => {
    if (!file) return;
    if (file.size > 500 * 1024) { alert('รูปใหญ่เกิน 500KB ครับ'); return; }
    const reader = new FileReader();
    reader.onload = e => setBgImage(e.target.result);
    reader.readAsDataURL(file);
  };

  return (
    <div className="fixed inset-0 bg-black/80 z-50 flex items-center justify-center p-4" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="bg-gray-800 rounded-2xl w-full max-w-md border border-gray-700 shadow-2xl max-h-[90vh] flex flex-col">
        <div className="flex justify-between items-center p-5 border-b border-gray-700 flex-shrink-0">
          <h2 className="text-xl font-bold flex items-center gap-2"><Settings size={18} className="text-gray-300" /> ตั้งค่าธีม Stage</h2>
          <button onClick={onClose}><X size={20} className="text-gray-400 hover:text-white" /></button>
        </div>
        <div className="flex-1 overflow-y-auto p-5 space-y-4">
          {/* Presets */}
          <div>
            <p className="text-xs text-gray-400 mb-2">Presets</p>
            <div className="grid grid-cols-3 gap-2 sm:grid-cols-5">
              {presets.map(p => (
                <button key={p.id} onClick={() => { setPreset(p.id); setBg(p.bg); setTextColor(p.textColor); setChordColor(p.chordColor); }}
                  style={{ background: p.bg, borderColor: preset === p.id ? p.chordColor : '#374151' }} className="border-2 rounded-lg p-2 text-left transition">
                  <span style={{ color: p.textColor }} className="text-xs font-semibold block truncate">{p.label}</span>
                  <div className="flex gap-1 mt-1"><span style={{ background: p.textColor }} className="w-2.5 h-2.5 rounded-full" /><span style={{ background: p.chordColor }} className="w-2.5 h-2.5 rounded-full" /></div>
                </button>
              ))}
            </div>
          </div>

          {/* Colors */}
          <div className="space-y-3">
            <p className="text-xs text-gray-400">สีตัวอักษรและคอร์ด</p>
            {[['สีพื้นหลัง', bg, setBg], ['สีตัวอักษร', textColor, setTextColor], ['สีคอร์ด', chordColor, setChordColor]].map(([label, val, setter]) => (
              <div key={label} className="flex items-center justify-between">
                <label className="text-sm text-gray-300">{label}</label>
                <div className="flex items-center gap-2">
                  <input type="color" value={val} onChange={e => setter(e.target.value)} className="w-10 h-8 rounded cursor-pointer border-0 bg-transparent" />
                  <span className="text-xs text-gray-400 font-mono">{val}</span>
                </div>
              </div>
            ))}
          </div>

          {/* Background Image */}
          <div className="space-y-2">
            <p className="text-xs text-gray-400">รูปพื้นหลัง Stage (ไม่บังคับ)</p>
            <div className="flex gap-2 items-center">
              <label className="flex-1 flex items-center gap-2 px-3 py-2 bg-gray-700 hover:bg-gray-600 rounded-lg cursor-pointer text-sm transition">
                <ImageIcon size={14} className="text-gray-400" />
                {bgImage ? 'เปลี่ยนรูป' : 'เลือกรูปพื้นหลัง'}
                <input type="file" accept="image/*" className="hidden" onChange={e => handleBgImage(e.target.files[0])} />
              </label>
              {bgImage && (
                <button onClick={() => setBgImage('')} className="px-3 py-2 bg-red-900/40 text-red-400 hover:bg-red-800 rounded-lg text-sm transition">ลบ</button>
              )}
            </div>
            {bgImage && (
              <div className="space-y-1">
                <div className="flex items-center justify-between text-xs text-gray-400">
                  <span>ความเข้มรูป</span>
                  <span>{Math.round(bgOpacity * 100)}%</span>
                </div>
                <input type="range" min="0.03" max="0.5" step="0.01" value={bgOpacity} onChange={e => setBgOpacity(parseFloat(e.target.value))} className="w-full accent-blue-500" />
              </div>
            )}
          </div>

          {/* Live Preview */}
          <div>
            <p className="text-xs text-gray-400 mb-2">Preview</p>
            <div className="rounded-xl p-4 border border-gray-600 relative overflow-hidden min-h-16" style={{ background: bg }}>
              {bgImage && (
                <div className="absolute inset-0 bg-cover bg-center" style={{ backgroundImage: `url(${bgImage})`, opacity: bgOpacity }} />
              )}
              <div className="relative z-10">
                <p style={{ color: chordColor }} className="text-lg font-bold">ชื่อเพลง</p>
                <p style={{ color: textColor }} className="text-sm mt-1">
                  เนื้อเพลงบรรทัดนี้{' '}
                  <span style={{ color: chordColor, fontWeight: 'bold', background: chordColor + '22', padding: '0 4px', borderRadius: 4 }}>Am</span>
                  {' '}ต่อด้วย{' '}
                  <span style={{ color: chordColor, fontWeight: 'bold', background: chordColor + '22', padding: '0 4px', borderRadius: 4 }}>G</span>
                </p>
                <p style={{ color: chordColor, opacity: 0.3 }} className="text-sm mt-1 truncate">Next: เพลงถัดไป</p>
              </div>
            </div>
          </div>

          {/* Gemini Keys Management */}
          <div className="bg-gray-900 rounded-xl p-3 flex items-center justify-between">
            <div>
              <p className="text-sm text-gray-300 font-semibold">Gemini API Keys</p>
              <p className="text-xs text-gray-500">{geminiKeyCount} key{geminiKeyCount !== 1 ? 's' : ''} — สุ่มหมุนเวียนอัตโนมัติ</p>
            </div>
            <button onClick={onManageGeminiKeys} className="px-3 py-1.5 bg-yellow-700 hover:bg-yellow-600 text-yellow-100 rounded-lg text-xs font-semibold transition">
              จัดการ Keys
            </button>
          </div>
        </div>

        <div className="flex gap-3 p-5 border-t border-gray-700 flex-shrink-0">
          <button onClick={onClose} className="flex-1 py-2.5 bg-gray-700 hover:bg-gray-600 rounded-xl font-semibold">ยกเลิก</button>
          <button onClick={() => { onSave({ bg, textColor, chordColor, bgImage, bgOpacity }); onClose(); }} className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-500 rounded-xl font-semibold">บันทึก</button>
        </div>
      </div>
    </div>
  );
}

// ---- Metronome ----
function Metronome({ bpm, active }) {
  const [beat, setBeat] = useState(false);
  useEffect(() => {
    if (!active || !bpm) return;
    const iv = setInterval(() => setBeat(b => !b), (60000 / parseInt(bpm)) / 2);
    return () => clearInterval(iv);
  }, [active, bpm]);
  if (!active) return null;
  return <div className={`w-4 h-4 rounded-full transition-all duration-75 ${beat ? 'bg-green-400 shadow-[0_0_12px_rgba(74,222,128,0.8)]' : 'bg-green-900'}`} title={`${bpm} BPM`} />;
}

// ==========================================
// Main App
// ==========================================
function App() {
  const auth = getAuth();

  // --- Auth ---
  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [showAuth, setShowAuth] = useState(false);

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, u => { setUser(u); setAuthLoading(false); });
    return () => unsub();
  }, []);

  // --- Playlists ---
  const [playlists, setPlaylists] = useState([]);
  const [activePlaylistId, setActivePlaylistId] = useState(null);
  const [showPlaylistForm, setShowPlaylistForm] = useState(false);
  const [editingPlaylist, setEditingPlaylist] = useState(null);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);

  // --- Songs ---
  const [songs, setSongs] = useState([]);
  const [allSongsAcrossPlaylists, setAllSongsAcrossPlaylists] = useState([]); // สำหรับ Panic Jump ข้าม playlist
  const [isLoading, setIsLoading] = useState(true);
  const [selectedSong, setSelectedSong] = useState(null);

  // --- Stage ---
  const [zoomLevel, setZoomLevel] = useState(100);
  const [isUploading, setIsUploading] = useState(false);
  const [isDragActiveImg, setIsDragActiveImg] = useState(false);
  const [isDragActiveAI, setIsDragActiveAI] = useState(false);
  const [isConductor, setIsConductor] = useState(false);
  const [showPanicJump, setShowPanicJump] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [isAutoScrolling, setIsAutoScrolling] = useState(false);
  const [scrollSpeed, setScrollSpeed] = useState(1);
  const [isEditingText, setIsEditingText] = useState(false);
  const [tempText, setTempText] = useState('');
  const [showScrollControls, setShowScrollControls] = useState(false);
  const [showMetronome, setShowMetronome] = useState(false);
  const [showSharedCue, setShowSharedCue] = useState(true);
  const [showRehearsalMode, setShowRehearsalMode] = useState(false);

  // --- UI ---
  const [showAddSong, setShowAddSong] = useState(false);
  const [editingSong, setEditingSong] = useState(null);
  const [showSettings, setShowSettings] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [showPDFExport, setShowPDFExport] = useState(false);
  const [songToCopy, setSongToCopy] = useState(null);
  const [filterTag, setFilterTag] = useState('');
  const [theme, setTheme] = useState(() => {
    try { return JSON.parse(localStorage.getItem('stage_theme')) || { bg: '#000000', textColor: '#ffffff', chordColor: '#facc15' }; } catch { return { bg: '#000000', textColor: '#ffffff', chordColor: '#facc15' }; }
  });

  // Firestore: playlists
  useEffect(() => {
    const q = query(collection(db, 'playlists'), orderBy('order'));
    const unsub = onSnapshot(q, snap => {
      const data = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      setPlaylists(data);
      if (data.length > 0 && !activePlaylistId) setActivePlaylistId(data[0].id);
    });
    return () => unsub();
  }, []);

  // Firestore: songs
  useEffect(() => {
    if (!activePlaylistId) { setSongs([]); setIsLoading(false); return; }
    setIsLoading(true);
    const q = query(collection(db, 'songs'), where('playlistId', '==', activePlaylistId), orderBy('order'));
    const unsub = onSnapshot(q, snap => {
      const data = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      setSongs(data);
      if (selectedSong && !isConductor && !isEditingText) {
        const updated = data.find(s => s.id === selectedSong.id);
        if (updated) setSelectedSong(updated);
      }
      setIsLoading(false);
    });
    return () => unsub();
  }, [activePlaylistId, isConductor, isEditingText]);

  // โหลดเพลงทุกเพลงจากทุก Playlist เพื่อใช้ใน Panic Jump
  useEffect(() => {
    const unsub = onSnapshot(collection(db, 'songs'), snap => {
      const data = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      setAllSongsAcrossPlaylists(data);
    });
    return () => unsub();
  }, []);

  // Conductor sync
  useEffect(() => {
    const unsub = onSnapshot(doc(db, 'settings', 'live_status'), snap => {
      if (snap.exists() && !isConductor) {
        const data = snap.data();
        if (data.activeSongId) {
          const song = songs.find(s => s.id === data.activeSongId);
          if (song && (!selectedSong || selectedSong.id !== data.activeSongId)) {
            setSelectedSong(song); setZoomLevel(100); setIsAutoScrolling(false); setIsEditingText(false);
          }
        }
      }
    });
    return () => unsub();
  }, [songs, isConductor, selectedSong?.id]);

  // Wake Lock
  useEffect(() => {
    let wl = null;
    if ('wakeLock' in navigator && selectedSong) navigator.wakeLock.request('screen').then(w => { wl = w; }).catch(() => {});
    return () => { if (wl) wl.release(); };
  }, [selectedSong?.id]);

  // Paste image
  useEffect(() => {
    const handlePaste = (e) => {
      if (!selectedSong || isUploading || isEditingText) return;
      for (let i = 0; i < e.clipboardData.items.length; i++) {
        if (e.clipboardData.items[i].type.startsWith('image/')) { processImageFile(e.clipboardData.items[i].getAsFile()); break; }
      }
    };
    window.addEventListener('paste', handlePaste);
    return () => window.removeEventListener('paste', handlePaste);
  }, [selectedSong, isUploading, isEditingText]);

  // Auto scroll
  useEffect(() => {
    if (!isAutoScrolling) return;
    const iv = setInterval(() => { const c = document.getElementById('chord-container'); if (c) c.scrollTop += scrollSpeed; }, 40);
    return () => clearInterval(iv);
  }, [isAutoScrolling, scrollSpeed]);

  const saveTheme = (t) => { setTheme(t); localStorage.setItem('stage_theme', JSON.stringify(t)); };

  const getGeminiKeyCount = () => { try { return JSON.parse(localStorage.getItem('gemini_api_keys') || '[]').length; } catch { return 0; } };
  const [geminiKeyCount, setGeminiKeyCount] = useState(getGeminiKeyCount);

  const handleManageGeminiKeys = () => {
    const current = (() => { try { return JSON.parse(localStorage.getItem('gemini_api_keys') || '[]'); } catch { return []; } })();
    const input = prompt(
      `จัดการ Gemini API Keys (${current.length} keys ปัจจุบัน)\n\nใส่ keys ทั้งหมด คั่นด้วย Enter หรือ comma\n(ลบออกทั้งหมดแล้วใส่ใหม่)\n\nKeys ปัจจุบัน:\n${current.map((k, i) => `${i + 1}. ${k.slice(0, 8)}...`).join('\n') || '(ว่าง)'}`,
      current.join('\n')
    );
    if (input === null) return;
    const parsed = input.split(/[\n,]+/).map(k => k.trim()).filter(Boolean);
    localStorage.setItem('gemini_api_keys', JSON.stringify(parsed));
    localStorage.setItem('gemini_key_idx', '0');
    setGeminiKeyCount(parsed.length);
    alert(`บันทึก ${parsed.length} key เรียบร้อย`);
  };

  // Playlist handlers
  const handleCreatePlaylist = async (data) => { await addDoc(collection(db, 'playlists'), { ...data, order: playlists.length, createdAt: Date.now() }); };
  const handleUpdatePlaylist = async (id, data) => { await updateDoc(doc(db, 'playlists', id), data); };
  const handleDeletePlaylist = async (id) => {
    if (!window.confirm('ลบ Playlist นี้? เพลงทั้งหมดจะถูกลบด้วย')) return;
    const q = query(collection(db, 'songs'), where('playlistId', '==', id));
    const snap = await getDocs(q);
    const batch = writeBatch(db);
    snap.docs.forEach(d => batch.delete(d.ref));
    batch.delete(doc(db, 'playlists', id));
    await batch.commit();
    if (activePlaylistId === id) setActivePlaylistId(playlists.find(p => p.id !== id)?.id || null);
  };

  const handleDuplicatePlaylist = async (playlistToCopy) => {
    if (!window.confirm(`ต้องการทำสำเนา Playlist "${playlistToCopy.name}" และเพลงทั้งหมดข้างใน ใช่หรือไม่?`)) return;
    try {
      // 1. แยก id เก่าทิ้งไป เอาเฉพาะข้อมูลที่เหลือมาใช้ (ป้องกัน Error undefined)
      const { id, ...playlistData } = playlistToCopy;
      
      // 2. สร้าง Playlist ใหม่
      const newPlaylistRef = await addDoc(collection(db, 'playlists'), {
        ...playlistData,
        name: `${playlistData.name} (สำเนา)`,
        order: playlists.length,
        createdAt: serverTimestamp()
      });

      // 3. ดึงเพลงทั้งหมดจาก Playlist เดิม
      const q = query(collection(db, 'songs'), where('playlistId', '==', playlistToCopy.id));
      const snapshot = await getDocs(q);
      
      // 4. ก๊อปปี้เพลงไปใส่ Playlist ใหม่พร้อมๆ กัน
      const batch = writeBatch(db);
      snapshot.forEach((docSnap) => {
        const newSongRef = doc(collection(db, 'songs'));
        batch.set(newSongRef, { 
          ...docSnap.data(), 
          playlistId: newPlaylistRef.id, 
          createdAt: serverTimestamp() 
        });
      });
      await batch.commit();
      
      alert(`ทำสำเนาเรียบร้อย! (${snapshot.size} เพลง)`);
    } catch (error) { 
      alert("Error: " + error.message); 
    }
  };

  const handleCopySongToPlaylist = async (song, targetPlaylistId) => {
    try {
      const { id, ...songData } = song;
      await addDoc(collection(db, 'songs'), { ...songData, playlistId: targetPlaylistId, order: 9999, createdAt: serverTimestamp() });
      setSongToCopy(null);
      alert("ส่งเพลงไป Playlist ใหม่สำเร็จ!");
    } catch (e) { alert("Error: " + e.message); }
  };

  // Song handlers
  const handleSelectSong = (song) => {
    setSelectedSong(song); setZoomLevel(100); setIsAutoScrolling(false);
    setShowPanicJump(false); setIsEditingText(false); setShowSharedCue(true);
    if (isConductor && song) setDoc(doc(db, 'settings', 'live_status'), { activeSongId: song.id, timestamp: Date.now() }, { merge: true }).catch(console.error);
  };
  const navigateSong = (dir) => {
    const idx = songs.findIndex(s => s.id === selectedSong.id);
    const next = idx + dir;
    if (next >= 0 && next < songs.length) handleSelectSong(songs[next]);
  };
  const toggleConductorMode = () => {
    const next = !isConductor; setIsConductor(next);
    setDoc(doc(db, 'settings', 'live_status'), { activeSongId: next && selectedSong ? selectedSong.id : null, timestamp: Date.now() }, { merge: true });
  };
  const handleAddSong = async (data) => { await addDoc(collection(db, 'songs'), { ...data, order: songs.length, playlistId: activePlaylistId }); };
  const handleUpdateSong = async (id, data) => { await updateDoc(doc(db, 'songs', id), data); };
  const handleDeleteSong = async (id) => { if (!window.confirm('ลบเพลงนี้?')) return; await deleteDoc(doc(db, 'songs', id)); };
  const handleOnDragEnd = async (result) => {
    if (!result.destination) return;
    const items = Array.from(songs);
    const [moved] = items.splice(result.source.index, 1);
    items.splice(result.destination.index, 0, moved);
    setSongs(items);
    items.forEach((item, i) => updateDoc(doc(db, 'songs', item.id), { order: i }).catch(console.error));
  };

  const processImageFile = (file) => {
    if (!file || !selectedSong) return;
    if (file.size > 700 * 1024) { alert("รูปภาพใหญ่เกิน 700KB"); return; }
    setIsUploading(true);
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = async () => {
      try { await updateDoc(doc(db, 'songs', selectedSong.id), { imageUrl: reader.result }); setSelectedSong(s => ({ ...s, imageUrl: reader.result })); }
      catch { alert("อัปโหลดไม่สำเร็จ"); } finally { setIsUploading(false); }
    };
    reader.onerror = () => { alert("อ่านไฟล์ไม่สำเร็จ"); setIsUploading(false); };
  };

  // --- Gemini multi-key rotation ---
  const getGeminiKeys = () => {
    try { return JSON.parse(localStorage.getItem('gemini_api_keys') || '[]'); } catch { return []; }
  };
  const saveGeminiKeys = (keys) => localStorage.setItem('gemini_api_keys', JSON.stringify(keys));
  const pickGeminiKey = () => {
    const keys = getGeminiKeys();
    if (keys.length === 0) return null;
    // rotate: เก็บ index ปัจจุบันไว้
    const idx = (parseInt(localStorage.getItem('gemini_key_idx') || '0')) % keys.length;
    localStorage.setItem('gemini_key_idx', (idx + 1) % keys.length);
    return keys[idx];
  };
  const ensureGeminiKeys = () => {
    let keys = getGeminiKeys();
    if (keys.length > 0) return true;
    const input = prompt("🔑 ใส่ Gemini API Key (รับฟรีที่ aistudio.google.com)\nใส่ได้หลาย key คั่นด้วย Enter หรือ comma เพื่อสุ่มหมุนเวียน:");
    if (!input) return false;
    const parsed = input.split(/[\n,]+/).map(k => k.trim()).filter(Boolean);
    saveGeminiKeys(parsed);
    return parsed.length > 0;
  };

  const processImageWithAI = async (file) => {
    if (!file || !selectedSong) return;
    if (!ensureGeminiKeys()) return;
    setIsUploading(true);
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = async () => {
      const base64Data = reader.result.split(',')[1];
      const apiKey = pickGeminiKey();
      try {
        const res = await fetch(`https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key=${apiKey}`, {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ contents: [{ parts: [{ text: "Extract the lyrics and chords from this image. Output ONLY the raw text. Do NOT use markdown code blocks. Preserve the exact placement of chords relative to the lyrics." }, { inline_data: { mime_type: file.type, data: base64Data } }] }] })
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error.message);
        let text = data.candidates[0].content.parts[0].text.replace(/```/g, '').trim();
        const detectedKey = detectKey(text);
        setTempText(text); setIsEditingText(true);
        if (detectedKey && detectedKey !== selectedSong.key && window.confirm(`AI เจอคีย์ "${detectedKey}" อัปเดตไหม?`)) {
          await updateDoc(doc(db, 'songs', selectedSong.id), { key: detectedKey });
          setSelectedSong(s => ({ ...s, key: detectedKey }));
        }
      } catch (e) {
        // ถ้า key นี้ใช้ไม่ได้ (429 rate limit) บอก user
        const msg = e.message || '';
        if (msg.includes('429') || msg.includes('quota')) {
          alert(`Key หมด quota ชั่วคราว กำลังลอง key ถัดไปในครั้งหน้า\n(${msg})`);
        } else {
          alert("AI Error: " + msg);
          // ถ้า error อื่น (invalid key) ให้ล้าง keys ทั้งหมด
          if (msg.includes('API_KEY') || msg.includes('invalid')) {
            localStorage.removeItem('gemini_api_keys');
            localStorage.removeItem('gemini_key_idx');
          }
        }
      }
      finally { setIsUploading(false); }
    };
  };

  const handleDeleteAsset = async (type) => {
    if (!window.confirm(`ลบ${type === 'image' ? 'รูปภาพ' : 'ข้อความ'}?`)) return;
    const update = type === 'image' ? { imageUrl: null } : { chordText: null };
    await updateDoc(doc(db, 'songs', selectedSong.id), update);
    setSelectedSong(s => ({ ...s, ...update }));
  };

  const handleSaveText = async () => {
    const history = [...(selectedSong.history || []), { text: selectedSong.chordText || '', savedAt: Date.now() }].slice(-10);
    await updateDoc(doc(db, 'songs', selectedSong.id), { chordText: tempText, history });
    setSelectedSong(s => ({ ...s, chordText: tempText, history }));
    setIsEditingText(false);
  };

  const handleRestoreHistory = async (oldText) => {
    if (!window.confirm('กู้คืนเวอร์ชันนี้?')) return;
    await updateDoc(doc(db, 'songs', selectedSong.id), { chordText: oldText });
    setSelectedSong(s => ({ ...s, chordText: oldText }));
  };

  const handleTranspose = async (steps) => {
    if (!selectedSong.chordText) return;
    const newText = transposeText(selectedSong.chordText, steps);
    await updateDoc(doc(db, 'songs', selectedSong.id), { chordText: newText });
    setSelectedSong(s => ({ ...s, chordText: newText }));
  };

  // Personal notes (stored locally per user per song)
  const personalNotesKey = user ? `pnotes_${user.uid}_${selectedSong?.id}` : null;
  const getPersonalNotes = () => { if (!personalNotesKey) return ''; try { return localStorage.getItem(personalNotesKey) || ''; } catch { return ''; } };
  const [showPersonalNotes, setShowPersonalNotes] = useState(false);
  const [personalNotesText, setPersonalNotesText] = useState('');
  useEffect(() => { if (selectedSong && user) setPersonalNotesText(getPersonalNotes()); }, [selectedSong?.id, user?.uid]);

  const renderTextWithChords = (text) => {
    if (!text) return null;
    const chordRegex = /([^a-zA-Z0-9]|^)([CDEFGAB][#b]?(?:m|maj|min|dim|aug|sus)?\d*(?:\/[CDEFGAB][#b]?)?)(?=[^a-zA-Z0-9]|$)/g;
    return (
      <div className="w-full rounded-xl p-4 md:p-6 leading-relaxed tracking-wide font-medium border border-gray-800"
        style={{ fontSize: `${zoomLevel}%`, background: theme.bg, color: theme.textColor }}>
        {text.split('\n').map((line, i) => {
          const parts = []; let lastIndex = 0; let match; const regex = new RegExp(chordRegex);
          while ((match = regex.exec(line)) !== null) {
            const before = line.substring(lastIndex, match.index + match[1].length);
            if (before) parts.push(before);
            parts.push(<span key={match.index} className="font-bold px-1.5 rounded mx-1 shadow-sm inline-block border" style={{ color: theme.chordColor, background: theme.chordColor + '22', borderColor: theme.chordColor + '55' }}>{match[2]}</span>);
            lastIndex = match.index + match[0].length;
          }
          const after = line.substring(lastIndex);
          if (after) parts.push(after);
          return <div key={i} className="whitespace-pre-wrap break-words">{parts.length > 0 ? parts : (line || '\u00a0')}</div>;
        })}
      </div>
    );
  };

  const allTags = [...new Set(songs.flatMap(s => s.tags || []))];
  const filteredSongs = filterTag ? songs.filter(s => (s.tags || []).includes(filterTag)) : songs;
  const totalSeconds = songs.reduce((acc, s) => acc + parseDuration(s.duration), 0);
  const activePlaylist = playlists.find(p => p.id === activePlaylistId);

  // Loading screen
  if (authLoading) return <div className="h-screen bg-gray-900 flex items-center justify-center"><div className="w-10 h-10 border-4 border-blue-500 border-t-transparent rounded-full animate-spin" /></div>;

  // ==========================================
  // Rehearsal Mode (full screen)
  // ==========================================
  if (showRehearsalMode && selectedSong) {
    return <RehearsalModal song={selectedSong} theme={theme} onClose={() => { setShowRehearsalMode(false); }}/>;
  }

  // ==========================================
  // Stage Mode
  // ==========================================
  if (selectedSong) {
    return (
      <div className="h-screen text-white p-2 md:p-4 flex flex-col relative overflow-hidden" style={{ background: theme.bg }}>
        {/* Background image layer */}
        {theme.bgImage && (
          <div className="absolute inset-0 bg-cover bg-center pointer-events-none" style={{ backgroundImage: `url(${theme.bgImage})`, opacity: theme.bgOpacity ?? 0.15 }} />
        )}
        {/* Top Bar */}
        <div className="flex justify-between items-center mb-2 z-20 relative flex-shrink-0">
          <button onClick={() => handleSelectSong(null)} className="flex items-center gap-1 text-gray-400 hover:text-white px-2 py-1">
            <ArrowLeft size={20} /> <span className="hidden md:inline text-sm">ออก</span>
          </button>
          <div className="flex gap-1.5 items-center">
            <Metronome bpm={selectedSong.bpm} active={showMetronome} />
            <button onClick={() => setShowMetronome(m => !m)} className={`p-2 rounded-lg transition ${showMetronome ? 'bg-green-700 text-white' : 'bg-gray-800/60 text-gray-400'}`}><Activity size={16} /></button>
            {selectedSong.chordText && (
              <button onClick={() => setShowRehearsalMode(true)} className="flex items-center gap-1 p-2 px-2.5 bg-gray-800/60 text-purple-400 hover:bg-gray-700 rounded-lg font-bold">
                <Repeat size={16} /> <span className="hidden md:inline text-sm">ซ้อม</span>
              </button>
            )}
            <button onClick={() => setShowPanicJump(true)} className="flex items-center gap-1 p-2 px-2.5 bg-gray-800/60 text-yellow-400 rounded-lg hover:bg-gray-700 font-bold">
              <Zap size={16} /> <span className="hidden md:inline text-sm">ด่วน</span>
            </button>
            <button onClick={() => setShowScrollControls(s => !s)} className={`flex items-center gap-1 p-2 px-2.5 rounded-lg font-bold transition ${isAutoScrolling ? 'bg-red-600 text-white animate-pulse' : 'bg-gray-800/60 text-gray-400'}`}>
              {isAutoScrolling ? <Square size={16} /> : <Play size={16} />}
              <span className="hidden md:inline text-sm">{isAutoScrolling ? 'หยุด' : 'เลื่อน'}</span>
            </button>
            <button onClick={toggleConductorMode} className={`flex items-center gap-1 p-2 px-2.5 rounded-lg font-bold transition ${isConductor ? 'bg-green-600 text-white' : 'bg-gray-800/60 text-gray-400'}`}>
              <Radio size={16} className={isConductor ? 'animate-pulse' : ''} />
              <span className="hidden md:inline text-sm">{isConductor ? 'คุมคิว' : 'รับคิว'}</span>
            </button>
            <button onClick={() => setShowSettings(true)} className="p-2 bg-gray-800/60 text-gray-400 hover:bg-gray-700 rounded-lg"><Settings size={16} /></button>
          </div>
        </div>

        {/* Scroll Controls */}
        {showScrollControls && (
          <div className="flex items-center gap-3 mb-2 px-3 py-2 bg-gray-900/80 rounded-xl border border-gray-700 flex-shrink-0 z-20">
            <span className="text-xs text-gray-400 whitespace-nowrap">ความเร็ว:</span>
            <input type="range" min="0.2" max="5" step="0.2" value={scrollSpeed} onChange={e => setScrollSpeed(parseFloat(e.target.value))} className="flex-1 accent-blue-500" />
            <span className="text-xs text-blue-300 font-mono w-8">{scrollSpeed.toFixed(1)}x</span>
            <button onClick={() => setIsAutoScrolling(s => !s)} className={`px-3 py-1 rounded-lg text-sm font-bold ${isAutoScrolling ? 'bg-red-600' : 'bg-blue-600'} text-white`}>
              {isAutoScrolling ? '⏹ หยุด' : '▶ เริ่ม'}
            </button>
          </div>
        )}

        {/* Song Info */}
        <div className="flex justify-between items-end mb-3 pb-3 border-b border-gray-800 flex-wrap gap-3 z-20 relative flex-shrink-0">
          <div className="flex-1 min-w-0">
            {/* ชื่อเพลงปัจจุบัน + เพลงถัดไป */}
            {(() => {
              const currentIndex = filteredSongs.findIndex(s => s.id === selectedSong.id);
              const nextSong = currentIndex >= 0 && currentIndex < filteredSongs.length - 1 ? filteredSongs[currentIndex + 1] : null;
              return (
                <div className="flex w-full justify-between items-baseline gap-4 mb-1">
                  
                  {/* ชื่อเพลงปัจจุบัน (ชิดซ้าย) */}
                  <h2 className="text-2xl md:text-3xl font-bold tracking-wide truncate" style={{ color: theme.chordColor }}>
                    {selectedSong.title}
                  </h2>
                  
                  {/* ชื่อเพลงถัดไป (ชิดขวา, ขนาดเท่ากัน, สีตามธีมแต่จางลง) */}
                  {nextSong && (
                    <div className="flex items-baseline gap-2 min-w-0 text-right">
                      <span className="text-gray-500 text-[10px] md:text-xs uppercase tracking-wider font-semibold hidden md:inline flex-shrink-0">
                        Next:
                      </span>
                      <span className="text-2xl md:text-3xl font-bold truncate opacity-30" style={{ color: theme.chordColor }}>
                        {nextSong.title}
                      </span>
                      {nextSong.key && (
                        <span className="text-[10px] md:text-xs bg-gray-800/50 px-1.5 py-0.5 rounded font-mono opacity-50 flex-shrink-0" style={{ color: theme.chordColor }}>
                          {nextSong.key}
                        </span>
                      )}
                    </div>
                  )}
                  
                </div>
              );
            })()}
            <div className="flex flex-wrap items-center gap-3 mt-1">
              <p className="text-xs md:text-sm text-gray-400">
                Key: <span className="text-white font-bold">{selectedSong.key}</span>
                {selectedSong.bpm && <> | BPM: <span className="text-white font-bold">{selectedSong.bpm}</span></>}
                {selectedSong.duration && <> | ⏳ {selectedSong.duration}</>}
              </p>
              {selectedSong.chordText && !isEditingText && (
                <div className="flex items-center gap-1 bg-gray-800/60 rounded-lg p-1">
                  <span className="text-xs text-gray-400 px-1">คีย์:</span>
                  <button onClick={() => handleTranspose(-1)} className="p-1 hover:text-yellow-400"><MinusCircle size={16} /></button>
                  <button onClick={() => handleTranspose(1)} className="p-1 hover:text-yellow-400"><PlusCircle size={16} /></button>
                </div>
              )}
              {(selectedSong.links || []).map((lk, i) => (
                <a key={i} href={lk.url} target="_blank" rel="noopener noreferrer" className="flex items-center gap-1 text-xs px-2 py-0.5 bg-blue-900/40 text-blue-300 rounded border border-blue-800 hover:bg-blue-800">
                  <Link size={10} /> {lk.label}
                </a>
              ))}
              {user && (
                <button onClick={() => setShowPersonalNotes(n => !n)} className={`flex items-center gap-1 text-xs px-2 py-0.5 rounded border transition ${showPersonalNotes ? 'bg-yellow-900/40 text-yellow-300 border-yellow-700' : 'bg-gray-800/60 text-gray-500 border-gray-700 hover:border-gray-500'}`}>
                  <EyeOff size={10} /> โน้ตส่วนตัว
                </button>
              )}
            </div>
          </div>
          <div className="flex gap-1.5 items-center flex-shrink-0">
            {(selectedSong.imageUrl || selectedSong.chordText) && (
              <button onClick={() => handleDeleteAsset(selectedSong.imageUrl ? 'image' : 'text')} className="p-1.5 bg-red-900/30 text-red-400 rounded-lg hover:bg-red-800 hover:text-white transition"><Trash2 size={16} /></button>
            )}
            {selectedSong.chordText && (selectedSong.history || []).length > 0 && !isEditingText && (
              <details className="relative">
                <summary className="p-1.5 bg-gray-800/60 rounded-lg text-gray-400 hover:bg-gray-700 cursor-pointer list-none"><RotateCcw size={16} /></summary>
                <div className="absolute right-0 top-full mt-1 bg-gray-800 border border-gray-700 rounded-xl p-2 w-52 z-50 shadow-2xl">
                  <p className="text-xs text-gray-400 mb-2">Version History</p>
                  {(selectedSong.history || []).slice().reverse().map((h, i) => (
                    <button key={i} onClick={() => handleRestoreHistory(h.text)} className="w-full text-left text-xs p-2 hover:bg-gray-700 rounded-lg mb-1 text-gray-300">{new Date(h.savedAt).toLocaleString('th-TH')}</button>
                  ))}
                </div>
              </details>
            )}
            {!isEditingText && (
              <>
                <button onClick={() => setZoomLevel(z => Math.max(50, z - 10))} className="p-1.5 bg-gray-800/60 rounded-lg hover:bg-gray-700"><ZoomOut size={16} /></button>
                <span className="text-gray-400 text-xs font-mono w-10 text-center">{zoomLevel}%</span>
                <button onClick={() => setZoomLevel(z => Math.min(300, z + 10))} className="p-1.5 bg-gray-800/60 rounded-lg hover:bg-gray-700"><ZoomIn size={16} /></button>
              </>
            )}
          </div>
        </div>

        {/* Shared Cue */}
        {selectedSong.sharedCues && showSharedCue && (
          <div className="mb-2 px-4 py-2 bg-blue-900/50 border border-blue-700 rounded-xl text-sm text-blue-200 flex justify-between items-start gap-2 z-20 flex-shrink-0">
            <div className="flex items-start gap-2"><Eye size={14} className="flex-shrink-0 mt-0.5" /><span>{selectedSong.sharedCues}</span></div>
            <button onClick={() => setShowSharedCue(false)}><X size={14} className="text-blue-400" /></button>
          </div>
        )}

        {/* Personal Notes (only visible to logged-in user) */}
        {showPersonalNotes && user && (
          <div className="mb-2 px-3 py-2 bg-yellow-900/30 border border-yellow-700/50 rounded-xl z-20 flex-shrink-0">
            <div className="flex items-center gap-2 mb-1">
              <EyeOff size={12} className="text-yellow-500" />
              <span className="text-xs text-yellow-400 font-semibold">โน้ตส่วนตัว ({user.displayName || user.email})</span>
            </div>
            <textarea value={personalNotesText} onChange={e => setPersonalNotesText(e.target.value)}
              onBlur={() => { if (personalNotesKey) localStorage.setItem(personalNotesKey, personalNotesText); }}
              placeholder="พิมพ์โน้ตส่วนตัวที่นี่ (บันทึกอัตโนมัติ, มองเห็นเฉพาะคุณ)..."
              rows={2} className="w-full bg-transparent text-yellow-100 text-xs outline-none resize-none placeholder-yellow-900" />
          </div>
        )}

        {/* Panic Jump — ค้นหาข้าม Playlist */}
        {showPanicJump && (
          <div className="absolute inset-0 bg-black/95 z-50 flex flex-col p-4 md:p-10">
            <div className="max-w-2xl w-full mx-auto flex flex-col h-full">
              <div className="flex justify-between items-center mb-4">
                <h3 className="text-xl font-bold text-yellow-400 flex items-center gap-2"><Zap size={22} /> ค้นหาเพลงด่วน <span className="text-xs text-gray-500 font-normal">(ทุก Playlist)</span></h3>
                <button onClick={() => setShowPanicJump(false)}><X size={20} className="text-gray-400 hover:text-white" /></button>
              </div>
              <input type="text" placeholder="พิมพ์ชื่อเพลง หรือชื่อศิลปิน..." value={searchQuery} onChange={e => setSearchQuery(e.target.value)} autoFocus
                className="w-full bg-gray-800 text-white text-xl p-4 rounded-xl border-2 border-gray-700 focus:border-yellow-400 outline-none mb-3" />
              <div className="flex-1 overflow-auto space-y-2">
                {(() => {
                  const q = searchQuery.toLowerCase();
                  const results = q.length < 1 ? songs : allSongsAcrossPlaylists.filter(s =>
                    s.title?.toLowerCase().includes(q) || s.artist?.toLowerCase().includes(q)
                  ).sort((a, b) => {
                    // เพลงใน playlist ปัจจุบันขึ้นก่อน
                    const aInCurrent = a.playlistId === activePlaylistId ? 0 : 1;
                    const bInCurrent = b.playlistId === activePlaylistId ? 0 : 1;
                    return aInCurrent - bInCurrent || a.title?.localeCompare(b.title);
                  });
                  if (results.length === 0) return <p className="text-gray-600 text-center py-10">ไม่พบเพลง "{searchQuery}"</p>;
                  return results.map(song => {
                    const pl = playlists.find(p => p.id === song.playlistId);
                    const isOtherPlaylist = song.playlistId !== activePlaylistId;
                    return (
                      <div key={song.id}
                        onClick={() => {
                          if (isOtherPlaylist) setActivePlaylistId(song.playlistId);
                          handleSelectSong(song);
                          setShowPanicJump(false);
                        }}
                        className={`p-4 rounded-xl border cursor-pointer flex justify-between items-center group transition ${isOtherPlaylist ? 'bg-gray-900/80 border-gray-700 hover:border-orange-400' : 'bg-gray-900 border-gray-700 hover:border-yellow-400'}`}>
                        <div className="min-w-0">
                          <span className={`text-base font-bold group-hover:text-yellow-400 ${isOtherPlaylist ? 'text-gray-300' : 'text-white'}`}>{song.title}</span>
                          <div className="flex items-center gap-2 mt-0.5">
                            <p className="text-gray-500 text-xs truncate">{song.artist}</p>
                            {pl && isOtherPlaylist && (
                              <span className="text-[10px] px-1.5 py-0.5 rounded-full border flex-shrink-0" style={{ color: pl.color, borderColor: pl.color + '66', background: pl.color + '22' }}>
                                {pl.icon} {pl.name}
                              </span>
                            )}
                          </div>
                        </div>
                        <span className="text-gray-400 text-sm font-mono flex-shrink-0 ml-3">Key {song.key}</span>
                      </div>
                    );
                  });
                })()}
              </div>
            </div>
          </div>
        )}

        {/* Tap Zones */}
        <div className="absolute left-0 top-36 bottom-0 w-12 md:w-20 z-10 flex items-center justify-start cursor-pointer hover:bg-white/5 group" onClick={() => navigateSong(-1)}>
          <ChevronLeft className="text-white/10 group-hover:text-white/40" size={52} />
        </div>
        <div className="absolute right-0 top-36 bottom-0 w-12 md:w-20 z-10 flex items-center justify-end cursor-pointer hover:bg-white/5 group" onClick={() => navigateSong(1)}>
          <ChevronRight className="text-white/10 group-hover:text-white/40" size={52} />
        </div>

        {/* Content */}
        <div id="chord-container" className="flex-1 overflow-y-auto p-2 md:p-4 flex justify-center items-start z-0 relative">
          {selectedSong.imageUrl && !isEditingText ? (
            <div className="w-full max-w-5xl flex justify-center pb-20">
              <img src={selectedSong.imageUrl} alt="Chord" style={{ width: `${zoomLevel}%`, maxWidth: 'none' }} className="rounded-xl shadow-2xl pointer-events-none" />
            </div>
          ) : selectedSong.chordText && !isEditingText ? (
            <div className="w-full max-w-5xl pb-20">{renderTextWithChords(selectedSong.chordText)}</div>
          ) : isEditingText ? (
            <div className="w-full max-w-4xl flex flex-col h-full z-20 relative">
              <textarea value={tempText} onChange={e => setTempText(e.target.value)} placeholder="วางเนื้อเพลงพร้อมคอร์ดที่นี่..."
                className="flex-1 bg-gray-800 text-white font-mono p-4 rounded-t-xl outline-none border border-gray-700 focus:border-blue-500 w-full resize-none leading-relaxed text-sm" />
              <div className="flex gap-2 mt-2">
                <button onClick={() => setIsEditingText(false)} className="flex-1 py-3 bg-gray-700 hover:bg-gray-600 rounded-b-xl font-bold">ยกเลิก</button>
                <button onClick={handleSaveText} className="flex-1 py-3 bg-blue-600 hover:bg-blue-500 rounded-b-xl font-bold flex items-center justify-center gap-2"><Save size={16} /> บันทึก</button>
              </div>
            </div>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 w-full max-w-5xl z-20 relative mt-6">
              <div className={`border-2 border-dashed rounded-xl p-6 flex flex-col items-center justify-center gap-4 transition-all text-center ${isDragActiveImg ? 'border-blue-500 bg-gray-800 scale-105' : 'border-gray-700 bg-gray-900 hover:border-gray-500'}`}
                onDragOver={e => { e.preventDefault(); setIsDragActiveImg(true); }} onDragLeave={() => setIsDragActiveImg(false)}
                onDrop={e => { e.preventDefault(); setIsDragActiveImg(false); if (e.dataTransfer.files[0]) processImageFile(e.dataTransfer.files[0]); }}>
                <ImageIcon className={isDragActiveImg ? 'text-blue-500' : 'text-gray-500'} size={32} />
                <div><p className="text-gray-300 font-bold">1. บันทึกเป็นรูปภาพ</p><p className="text-gray-500 text-xs mt-1">ลากวาง หรือ Ctrl+V</p></div>
                <label className={`px-4 py-2 rounded-lg flex items-center gap-2 cursor-pointer text-sm ${isUploading ? 'bg-gray-600 text-gray-400' : 'bg-gray-800 hover:bg-gray-700 border border-gray-600'}`}>
                  <Upload size={14} /> เลือกไฟล์{!isUploading && <input type="file" accept="image/*" onChange={e => processImageFile(e.target.files[0])} className="hidden" />}
                </label>
              </div>
              <div className={`border-2 rounded-xl p-6 flex flex-col items-center justify-center gap-4 transition-all text-center ${isDragActiveAI ? 'border-yellow-400 bg-yellow-900/20 scale-105' : 'border-yellow-700/50 bg-gray-900 hover:border-yellow-500'}`}
                onDragOver={e => { e.preventDefault(); setIsDragActiveAI(true); }} onDragLeave={() => setIsDragActiveAI(false)}
                onDrop={e => { e.preventDefault(); setIsDragActiveAI(false); if (e.dataTransfer.files[0]) processImageWithAI(e.dataTransfer.files[0]); }}>
                <Sparkles className={isDragActiveAI ? 'text-yellow-300' : 'text-yellow-500'} size={32} />
                <div><p className="text-yellow-400 font-bold">2. แปลงรูปเป็นคอร์ด (AI)</p><p className="text-gray-500 text-xs mt-1">ลากวาง หรือ Ctrl+V หรือเลือกไฟล์</p></div>
                <label className={`px-4 py-2 rounded-lg flex items-center gap-2 cursor-pointer text-sm font-bold ${isUploading ? 'bg-gray-600 text-gray-400' : 'bg-yellow-600 hover:bg-yellow-500 text-white'}`}>
                  {isUploading ? 'กำลังประมวลผล...' : <><Sparkles size={14} /> อัปโหลดให้ AI</>}
                  {!isUploading && <input type="file" accept="image/*" onChange={e => processImageWithAI(e.target.files[0])} className="hidden" />}
                </label>
              </div>
              <div className="border-2 border-gray-700 rounded-xl p-6 flex flex-col items-center justify-center gap-4 bg-gray-900 hover:border-blue-500 transition-colors text-center">
                <Type className="text-blue-500" size={32} />
                <div><p className="text-gray-300 font-bold">3. เพิ่มคอร์ดแบบข้อความ</p><p className="text-gray-500 text-xs mt-1">ก๊อปปี้เนื้อเพลง/คอร์ดมาวางเอง</p></div>
                <button onClick={() => { setTempText(''); setIsEditingText(true); }} className="px-4 py-2 rounded-lg flex items-center gap-2 bg-blue-600 hover:bg-blue-500 font-bold text-sm">
                  <FileText size={14} /> วางข้อความ
                </button>
              </div>
            </div>
          )}
        </div>

        {showSettings && <SettingsModal theme={theme} onClose={() => setShowSettings(false)} onSave={saveTheme} geminiKeyCount={geminiKeyCount} onManageGeminiKeys={handleManageGeminiKeys} />}
      </div>
    );
  }

  // ==========================================
  // Setlist Screen (MERGED — Responsive + Duplicate + Copy Song)
  // ==========================================
  return (
    <div className="h-screen bg-gray-900 text-white flex overflow-hidden relative">

      {/* ฉากหลังสีดำตอนสไลด์เมนูออกมา (เฉพาะบนมือถือ) */}
      {isSidebarOpen && (
        <div className="fixed inset-0 bg-black/60 z-30 md:hidden" onClick={() => setIsSidebarOpen(false)} />
      )}

      {/* Sidebar (เมนูซ้ายมือ) */}
      <aside className={`fixed inset-y-0 left-0 z-40 w-60 bg-gray-950 border-r border-gray-800 flex flex-col h-full transition-transform duration-300 ease-in-out md:relative md:translate-x-0 ${isSidebarOpen ? 'translate-x-0' : '-translate-x-full'}`}>
        <div className="p-4 border-b border-gray-800">
          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center gap-2">
              <Music size={18} className="text-blue-400" />
              <span className="font-bold text-white text-sm">BandSetlist</span>
            </div>
            {/* Auth button */}
            {user ? (
              <div className="flex items-center gap-1">
                <span className="text-xs text-gray-400 truncate max-w-16">{user.displayName || user.email?.split('@')[0]}</span>
                <button onClick={() => signOut(auth)} className="p-1 text-gray-500 hover:text-red-400" title="ออกจากระบบ"><LogOut size={13} /></button>
              </div>
            ) : (
              <button onClick={() => setShowAuth(true)} className="flex items-center gap-1 text-xs text-blue-400 hover:text-blue-300 transition">
                <LogIn size={13} /> Login
              </button>
            )}
          </div>
          <button onClick={() => { setEditingPlaylist(null); setShowPlaylistForm(true); }}
            className="w-full flex items-center gap-2 px-3 py-2 bg-blue-700 hover:bg-blue-600 rounded-lg text-sm font-semibold transition">
            <FolderPlus size={14} /> สร้าง Playlist
          </button>
        </div>

        <nav className="flex-1 p-3 space-y-1 overflow-y-auto">
          {playlists.length === 0 && <p className="text-gray-600 text-xs text-center py-6">ยังไม่มี Playlist<br />กดสร้างด้านบน</p>}
          {playlists.map(pl => {
            const isActive = activePlaylistId === pl.id;
            return (
              <div key={pl.id}
                // เมื่อกดเลือก Playlist บนมือถือ ให้ลิ้นชักหุบเก็บอัตโนมัติ
                onClick={() => { setActivePlaylistId(pl.id); setIsSidebarOpen(false); }}
                className={`group flex items-center gap-2 px-3 py-2.5 rounded-xl cursor-pointer transition ${isActive ? 'bg-gray-800 text-white' : 'text-gray-400 hover:bg-gray-800/50 hover:text-gray-200'}`}
                style={isActive ? { borderLeft: `3px solid ${pl.color}` } : {}}>
                <span className="text-lg flex-shrink-0">{pl.icon}</span>
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-semibold truncate">{pl.name}</p>
                  {isActive && pl.description && <p className="text-xs text-gray-500 truncate">{pl.description}</p>}
                </div>
                {isActive && (
                  <div className="flex gap-0.5 opacity-100 md:opacity-0 md:group-hover:opacity-100 transition flex-shrink-0">
                    <button onClick={e => { e.stopPropagation(); setEditingPlaylist(pl); setShowPlaylistForm(true); }} className="p-1 hover:text-yellow-400" title="แก้ไข"><Edit2 size={11} /></button>
                    <button onClick={e => { e.stopPropagation(); handleDuplicatePlaylist(pl); }} className="p-1 hover:text-green-400" title="ทำสำเนา"><Copy size={11} /></button>
                    <button onClick={e => { e.stopPropagation(); handleDeletePlaylist(pl.id); }} className="p-1 hover:text-red-400" title="ลบ"><Trash2 size={11} /></button>
                  </div>
                )}
              </div>
            );
          })}
        </nav>

        <div className="p-3 border-t border-gray-800 space-y-1 flex-shrink-0">
          <button onClick={() => setShowImport(true)} className="w-full flex items-center gap-2 px-3 py-2 text-gray-400 hover:bg-gray-800 hover:text-green-400 rounded-lg text-sm transition">
            <ArrowUpFromLine size={14} /> Import Excel/CSV
          </button>
          <button onClick={() => { if (songs.length > 0) setShowPDFExport(true); else alert('ยังไม่มีเพลงใน Playlist นี้'); }} className="w-full flex items-center gap-2 px-3 py-2 text-gray-400 hover:bg-gray-800 hover:text-blue-400 rounded-lg text-sm transition">
            <Printer size={14} /> Export PDF
          </button>
          <button onClick={() => setShowSettings(true)} className="w-full flex items-center gap-2 px-3 py-2 text-gray-400 hover:bg-gray-800 rounded-lg text-sm transition">
            <Settings size={14} /> ตั้งค่าธีม
          </button>
        </div>
      </aside>

      {/* Main (พื้นที่แสดงรายชื่อเพลง) */}
      <main className="flex-1 min-w-0 h-full overflow-y-auto p-4 md:p-8">

        {/* แถบเมนูด้านบนที่มีปุ่ม Hamburger (เห็นเฉพาะบนจอมือถือ) */}
        <div className="md:hidden flex items-center gap-3 mb-4 pb-3 border-b border-gray-800">
          <button onClick={() => setIsSidebarOpen(true)} className="p-1 text-gray-400 hover:text-white">
            <Menu size={26} />
          </button>
          <div className="flex items-center gap-2">
            <Music size={18} className="text-blue-400" />
            <span className="font-bold text-white text-base">BandSetlist</span>
          </div>
        </div>

        {!activePlaylistId ? (
          <div className="flex flex-col items-center justify-center h-full text-gray-600 gap-4 py-20">
            <Folder size={64} className="opacity-30" />
            <p className="text-lg font-semibold">เลือกหรือสร้าง Playlist ก่อน</p>
            <button onClick={() => setShowPlaylistForm(true)} className="flex items-center gap-2 px-5 py-2.5 bg-blue-700 hover:bg-blue-600 rounded-xl text-white font-semibold"><FolderPlus size={18} /> สร้าง Playlist แรก</button>
          </div>
        ) : (
          <>
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-5 gap-3">
              <div>
                <h1 className="text-xl md:text-2xl font-bold flex items-center gap-2">
                  <span className="text-2xl">{activePlaylist?.icon}</span>
                  <span style={{ color: activePlaylist?.color || 'white' }}>{activePlaylist?.name}</span>
                </h1>
                {activePlaylist?.description && <p className="text-gray-500 text-sm mt-0.5">{activePlaylist.description}</p>}
                <div className="flex items-center gap-3 mt-1.5 text-gray-400 text-sm">
                  <span className="flex items-center gap-1"><Clock size={13} className="text-yellow-500" />{formatDuration(totalSeconds)}</span>
                  <span className="text-gray-600">•</span>
                  <span>{songs.length} เพลง</span>
                  {user && <span className="text-green-500 text-xs flex items-center gap-1"><User size={11} />{user.displayName || user.email?.split('@')[0]}</span>}
                </div>
              </div>
              <div className="flex gap-2 flex-wrap">
                <button onClick={() => setShowPDFExport(true)} className="flex items-center gap-1.5 px-3 py-2 bg-gray-800 hover:bg-gray-700 text-blue-400 rounded-lg text-sm font-semibold transition">
                  <Printer size={14} /> PDF
                </button>
                <button onClick={() => setShowImport(true)} className="flex items-center gap-1.5 px-3 py-2 bg-gray-800 hover:bg-gray-700 text-green-400 rounded-lg text-sm font-semibold transition">
                  <ArrowUpFromLine size={14} /> Import
                </button>
                <button onClick={() => setShowAddSong(true)} className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 hover:bg-blue-500 rounded-lg font-semibold shadow-lg text-sm whitespace-nowrap">
                  <Plus size={15} /> เพิ่มเพลง
                </button>
              </div>
            </div>

            {allTags.length > 0 && (
              <div className="flex gap-2 flex-wrap mb-4">
                <button onClick={() => setFilterTag('')} className={`text-xs px-3 py-1 rounded-full border transition ${!filterTag ? 'bg-blue-700 border-blue-600 text-white' : 'bg-gray-800 border-gray-700 text-gray-400 hover:border-gray-500'}`}>ทั้งหมด</button>
                {allTags.map(t => { const c = getTagColor(t); return <button key={t} onClick={() => setFilterTag(filterTag === t ? '' : t)} className={`text-xs px-3 py-1 rounded-full border transition ${filterTag === t ? `${c.bg} ${c.border} ${c.text}` : 'bg-gray-800 border-gray-700 text-gray-400 hover:border-gray-500'}`}>#{t}</button>; })}
              </div>
            )}

            {isLoading ? (
              <div className="text-center text-gray-500 py-16">กำลังโหลด...</div>
            ) : filteredSongs.length === 0 ? (
              <div className="flex flex-col items-center justify-center py-20 text-gray-600 gap-3">
                <Music size={48} className="opacity-30" />
                <p>{filterTag ? `ไม่มีเพลงที่มีแท็ก "${filterTag}"` : 'ยังไม่มีเพลง'}</p>
                {!filterTag && <div className="flex gap-2">
                  <button onClick={() => setShowAddSong(true)} className="flex items-center gap-1 px-4 py-2 bg-blue-700 hover:bg-blue-600 rounded-xl text-sm font-semibold"><Plus size={14} /> เพิ่มเพลง</button>
                  <button onClick={() => setShowImport(true)} className="flex items-center gap-1 px-4 py-2 bg-green-800 hover:bg-green-700 rounded-xl text-sm font-semibold text-green-300"><ArrowUpFromLine size={14} /> Import Excel</button>
                </div>}
              </div>
            ) : (
              <DragDropContext onDragEnd={handleOnDragEnd}>
                <Droppable droppableId="setlist">
                  {(provided) => (
                    <div {...provided.droppableProps} ref={provided.innerRef} className="space-y-2 pb-20">
                      {filteredSongs.map((song, index) => (
                        <Draggable draggableId={song.id} index={index} key={song.id}>
                          {(provided, snapshot) => (
                            <div ref={provided.innerRef} {...provided.draggableProps}
                              className={`flex items-center gap-2 md:gap-3 p-3 rounded-xl border transition-colors ${snapshot.isDragging ? 'bg-gray-700 border-blue-500 shadow-2xl' : 'bg-gray-800 border-gray-700 hover:border-gray-600'}`}>
                              <div {...provided.dragHandleProps} className="cursor-grab text-gray-600 hover:text-gray-300 p-1 flex-shrink-0"><GripVertical size={18} /></div>
                              <div className="flex-1 min-w-0 cursor-pointer" onClick={() => handleSelectSong(song)}>
                                <h3 className="text-sm md:text-base font-semibold truncate flex items-center gap-1.5 hover:text-blue-400 transition">
                                  {song.title}
                                  {song.imageUrl && <ImageIcon className="text-green-400 flex-shrink-0" size={11} />}
                                  {song.chordText && <FileText className="text-blue-400 flex-shrink-0" size={11} />}
                                  {song.sharedCues && <Eye className="text-yellow-500 flex-shrink-0" size={11} />}
                                  {(song.sections || []).length > 0 && <Repeat className="text-purple-400 flex-shrink-0" size={11} />}
                                </h3>
                                <div className="flex items-center gap-2 text-gray-400 text-xs mt-0.5">
                                  <span className="truncate">{song.artist}</span>
                                  {song.duration && <><span className="text-gray-600">•</span><span className="flex items-center gap-0.5 text-yellow-600"><Clock size={9} />{song.duration}</span></>}
                                  {(song.links || []).length > 0 && <><span className="text-gray-600">•</span><Link size={9} className="text-blue-500" /><span className="text-blue-500">{song.links.length}</span></>}
                                </div>
                                {(song.tags || []).length > 0 && (
                                  <div className="flex gap-1 mt-1.5 overflow-x-auto">
                                    {song.tags.map((t, i) => { const c = getTagColor(t); return <span key={i} className={`flex-shrink-0 text-[10px] px-1.5 py-0 rounded-full border ${c.bg} ${c.border} ${c.text}`}>{t}</span>; })}
                                  </div>
                                )}
                              </div>
                              <span className="bg-blue-900/50 text-blue-300 text-xs px-2 py-1 rounded-md font-mono border border-blue-800/50 flex-shrink-0">{song.key}</span>
                              <button onClick={e => { e.stopPropagation(); setSongToCopy(song); }} className="p-1.5 text-gray-500 hover:text-green-400 hover:bg-gray-700 rounded-lg transition flex-shrink-0" title="ส่งไป Playlist อื่น"><Copy size={14} /></button>
                              <button onClick={e => { e.stopPropagation(); setEditingSong(song); }} className="p-1.5 text-gray-500 hover:text-yellow-400 hover:bg-gray-700 rounded-lg transition flex-shrink-0" title="แก้ไข"><Edit2 size={14} /></button>
                              <button onClick={e => { e.stopPropagation(); handleDeleteSong(song.id); }} className="p-1.5 text-gray-600 hover:text-red-400 hover:bg-gray-700 rounded-lg transition flex-shrink-0" title="ลบ"><Trash2 size={14} /></button>
                            </div>
                          )}
                        </Draggable>
                      ))}
                      {provided.placeholder}
                    </div>
                  )}
                </Droppable>
              </DragDropContext>
            )}
          </>
        )}
      </main>

      {/* ===== Modals ทั้งหมด ===== */}
      {showAuth && <AuthModal onClose={() => setShowAuth(false)} />}
      {showPlaylistForm && <PlaylistFormModal playlist={editingPlaylist} onClose={() => { setShowPlaylistForm(false); setEditingPlaylist(null); }} onSave={data => editingPlaylist ? handleUpdatePlaylist(editingPlaylist.id, data) : handleCreatePlaylist(data)} />}
      {showAddSong && <AddSongModal onClose={() => setShowAddSong(false)} onSave={handleAddSong} allSongs={songs} />}
      {editingSong && <EditSongModal song={editingSong} allSongs={songs} onClose={() => setEditingSong(null)} onSave={data => { handleUpdateSong(editingSong.id, data); setEditingSong(null); }} />}
      {showSettings && <SettingsModal theme={theme} onClose={() => setShowSettings(false)} onSave={saveTheme} geminiKeyCount={geminiKeyCount} onManageGeminiKeys={handleManageGeminiKeys} />}
      {showPDFExport && <PDFExportModal songs={filteredSongs.length > 0 ? filteredSongs : songs} playlist={activePlaylist} onClose={() => setShowPDFExport(false)} />}
      {showImport && playlists.length > 0 && <ImportModal playlists={playlists} onClose={() => setShowImport(false)} onImport={() => setTimeout(() => setShowImport(false), 2000)} />}

      {/* Import แต่ยังไม่มี Playlist */}
      {showImport && playlists.length === 0 && (
        <div className="fixed inset-0 bg-black/80 z-50 flex items-center justify-center p-4">
          <div className="bg-gray-800 rounded-2xl p-8 max-w-sm text-center border border-gray-700">
            <AlertCircle size={40} className="text-yellow-400 mx-auto mb-3" />
            <h3 className="text-lg font-bold mb-2">ต้องมี Playlist ก่อน</h3>
            <p className="text-gray-400 text-sm mb-4">กรุณาสร้าง Playlist ก่อน Import</p>
            <div className="flex gap-2">
              <button onClick={() => setShowImport(false)} className="flex-1 py-2 bg-gray-700 rounded-xl text-sm">ปิด</button>
              <button onClick={() => { setShowImport(false); setShowPlaylistForm(true); }} className="flex-1 py-2 bg-blue-600 rounded-xl text-sm font-semibold">สร้าง Playlist</button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: ส่งเพลงไป Playlist อื่น */}
      {songToCopy && (
        <div className="fixed inset-0 bg-black/80 z-50 flex items-center justify-center p-4" onClick={e => e.target === e.currentTarget && setSongToCopy(null)}>
          <div className="bg-gray-800 p-6 rounded-2xl w-full max-w-sm border border-gray-700 shadow-2xl">
            <h3 className="text-lg font-bold mb-1 flex items-center gap-2"><Copy size={16} className="text-green-400" /> ส่งเพลงไป Playlist อื่น</h3>
            <p className="mb-4 text-sm text-blue-400 truncate">"{songToCopy.title}"</p>
            <div className="space-y-2 max-h-60 overflow-y-auto pr-1">
              {playlists.filter(p => p.id !== activePlaylistId).length === 0 ? (
                <p className="text-gray-500 text-sm text-center py-4">ไม่มี Playlist อื่นให้เลือก<br /><span className="text-xs text-gray-600">สร้าง Playlist ใหม่ก่อน</span></p>
              ) : (
                playlists.filter(p => p.id !== activePlaylistId).map(pl => (
                  <button key={pl.id} onClick={() => handleCopySongToPlaylist(songToCopy, pl.id)}
                    className="w-full text-left px-4 py-3 bg-gray-700 hover:bg-gray-600 rounded-xl text-sm font-medium transition flex items-center gap-3">
                    <span className="text-lg">{pl.icon}</span>
                    <div>
                      <p className="font-semibold">{pl.name}</p>
                      {pl.description && <p className="text-xs text-gray-400">{pl.description}</p>}
                    </div>
                  </button>
                ))
              )}
            </div>
            <button onClick={() => setSongToCopy(null)} className="w-full mt-4 py-2.5 bg-gray-700 text-gray-300 hover:bg-gray-600 rounded-xl font-semibold transition">ยกเลิก</button>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
