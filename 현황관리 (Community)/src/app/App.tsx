import { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  Search, X, Lock, Eye, EyeOff, Shield, Settings, Save,
  Plus, Trash2, ChevronUp, ChevronDown, Edit2, Check,
  KeyRound, Image as ImageIcon, Upload, AlertCircle,
  RefreshCw, Home, FileSpreadsheet, CheckSquare, Square,
  Download, GripVertical, Columns,
} from "lucide-react";

// ─── Types ─────────────────────────────────────────────────────────────────────
type ColType  = "status" | "station" | "text" | "number" | "note" | "date";
type ColAlign = "left" | "right" | "center";
type Status   = "판매중" | "판매완료";

interface LineTab  { id: string; label: string; color: string; }
interface MediaTab { id: string; label: string; lines: LineTab[]; }
interface Column   { id: string; key: string; label: string; type: ColType; align: ColAlign; fixedWidth?: number; }
interface Row      { id: string; [key: string]: string; }

interface AppData {
  mediaTabs: MediaTab[];
  defaultColumns: Column[];
  tabColumns: Record<string, Column[]>;   // key: `${mediaId}__${lineId}`
  tableData: Record<string, Record<string, Row[]>>;
  stationImages: Record<string, string[]>;
  adminPassword: string;
}

// ─── Station lists ──────────────────────────────────────────────────────────────
const STATIONS: Record<string, string[]> = {
  분당선:     ["왕십리","서울숲","압구정로데오","강남구청","선정릉","선릉","한티","도곡","구룡","개포동","대모산입구","수서","복정","가천대","야탑","이매","서현","수내","정자","미금"],
  경원선:     ["용산","서빙고","한남","옥수","응봉","왕십리","청량리","회기","외대앞","신이문","석계","광운대","월계","녹천","창동"],
  경부선:     ["서울","남영","용산","노량진","영등포","신도림","구로","가산디지털단지","금천구청","안양","군포","의왕","성균관대","화서","수원"],
  경인선:     ["구로","구일","개봉","오류동","온수","역곡","소사","부천","중동","송내","부개","부평","백운","동암","간석","주안","도화","제물포","인천"],
  "8호선":    ["암사","천호","강동구청","몽촌토성","잠실","석촌","송파","가락시장","문정","장지","복정","산성","단대오거리","신흥","수진","모란"],
  경원선조명: ["용산","서빙고","한남","옥수","응봉","왕십리","청량리","회기","외대앞","신이문","석계","광운대","월계","녹천","창동"],
  "5호선":    ["방화","개화산","김포공항","송정","마곡","발산","화곡","까치산","목동","오목교","양평","영등포구청","여의도","여의나루","마포","공덕","애오개","충정로","광화문","종로3가","동대문역사문화공원","왕십리","천호","강동"],
  "6호선":    ["응암","불광","연신내","구산","새절","증산","디지털미디어시티","마포구청","합정","상수","광흥창","대흥","공덕","삼각지","이태원","한강진","약수","청구","신당","동묘앞","고려대","월곡","석계","태릉입구","화랑대"],
};

const ADVERTISERS = ["삼성전자","LG전자","현대자동차","롯데백화점","신세계","이마트","CJ제일제당","네이버","카카오","SK텔레콤","KT","우리은행","신한은행","국민은행","하나은행","GS칼텍스","아모레퍼시픽","오리온","빙그레","KB증권"];
const MANAGERS    = ["김민준","이서연","박지훈","최수아","정우진","강예린","윤하준","임지유","한민서","오준혁"];
const PERIODS     = ["1개월","2개월","3개월","6개월","12개월","1개월","2개월","3개월","1개월","6개월"];
const SIZES       = ["500×700mm","600×900mm","400×600mm","700×1000mm","300×500mm","1000×1400mm","1200×1600mm"];
const AD_FEES     = ["850,000","1,200,000","450,000","1,500,000","520,000","2,200,000","2,800,000"];
const PROD_FEES   = ["120,000","150,000","80,000","180,000","90,000","280,000","320,000"];
const NOTES       = ["","위치우수","유동인구多","환승역","","출구인접","역세권","광고효과우수","","접근성우수","유동인구多","","환승역","","역세권"];
const STATUS_PAT: Status[] = ["판매중","판매중","판매완료","판매중","판매완료","판매완료","판매중","판매중","판매완료","판매중","판매중","판매완료","판매중","판매완료","판매중"];
const SAMPLE_IMGS = [
  "https://images.unsplash.com/photo-1544620626-fca4d2b56c78?w=800&h=600&fit=crop&auto=format",
  "https://images.unsplash.com/photo-1555993539-1732b0258235?w=800&h=600&fit=crop&auto=format",
];

function fmtDate(y: number, m: number, d: number) {
  return `${y}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
}

const DEFAULT_COLS: Column[] = [
  { id:"cc4",  key:"status",     label:"판매현황",type:"status",  align:"center", fixedWidth:90 },
  { id:"cc1",  key:"station",    label:"역명",    type:"station", align:"left",   fixedWidth:150 },
  { id:"cc2",  key:"media",      label:"매체",    type:"text",    align:"left" },
  { id:"cc3",  key:"number",     label:"번호",    type:"text",    align:"center" },
  { id:"cc5",  key:"advertiser", label:"광고주",  type:"text",    align:"left" },
  { id:"cc6",  key:"manager",    label:"담당자",  type:"text",    align:"left" },
  { id:"cc7",  key:"period",     label:"기간",    type:"text",    align:"center" },
  { id:"cc8",  key:"startDate",  label:"게첨일자",type:"date",    align:"center" },
  { id:"cc9",  key:"endDate",    label:"종료일자",type:"date",    align:"center" },
  { id:"cc10", key:"size",       label:"사이즈",  type:"text",    align:"left" },
  { id:"cc11", key:"adFee",      label:"광고료",  type:"number",  align:"right" },
  { id:"cc12", key:"prodFee",    label:"제작비",  type:"number",  align:"right" },
  { id:"cc13", key:"note",       label:"비고",    type:"note",    align:"left",  fixedWidth:160 },
];

function uid() { return Math.random().toString(36).slice(2,9); }

// ── Effective status (종료일자가 미래이면 자동 판매완료) ─────────────────────────
function getEffectiveStatus(row: Row): Status {
  if (row.endDate) {
    const n = parseDate(row.endDate);
    if (/^\d{4}-\d{2}-\d{2}$/.test(n)) {
      const end = new Date(n + "T23:59:59");
      if (!isNaN(end.getTime()) && end > new Date()) return "판매완료";
    }
  }
  return (row.status || "판매중") as Status;
}

// ── Arrow-key navigation between table cells ────────────────────────────────
function navigateCell(e: React.KeyboardEvent<HTMLInputElement|HTMLTextAreaElement>, ri: number, ci: number) {
  const el = e.currentTarget;
  const key = e.key;
  let dRow = 0, dCol = 0;

  if (key === "ArrowUp")   { dRow = -1; }
  else if (key === "ArrowDown")  { dRow = 1; }
  else if (key === "ArrowLeft" && el.selectionStart === 0 && el.selectionEnd === 0) { dCol = -1; }
  else if (key === "ArrowRight" && el.selectionStart === el.value.length && el.selectionEnd === el.value.length) { dCol = 1; }
  else return;

  e.preventDefault();
  const target = document.querySelector<HTMLElement>(
    `[data-cell-row="${ri + dRow}"][data-cell-col="${ci + dCol}"]`
  );
  if (target) {
    target.focus();
    if (dCol !== 0) {
      try {
        const inp = target as HTMLInputElement;
        const len = inp.value.length;
        inp.setSelectionRange(dCol > 0 ? 0 : len, dCol > 0 ? 0 : len);
      } catch {}
    }
  }
}

// ── Number formatting (천단위 콤마) ────────────────────────────────────────────
function formatNumber(val: string): string {
  const raw = val.replace(/,/g, "").trim();
  if (!raw || isNaN(Number(raw))) return val;
  return Number(raw).toLocaleString("ko-KR");
}
function stripCommas(val: string): string {
  return val.replace(/,/g, "");
}

// ── Date normalizer (YYYY-MM-DD / YYYY.MM.DD / YYYY년MM월DD일 → YYYY-MM-DD) ──
function parseDate(val: string): string {
  const s = val.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const dot = s.match(/^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})$/);
  if (dot) return `${dot[1]}-${dot[2].padStart(2,"0")}-${dot[3].padStart(2,"0")}`;
  const kor = s.match(/^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일?$/);
  if (kor) return `${kor[1]}-${kor[2].padStart(2,"0")}-${kor[3].padStart(2,"0")}`;
  return val;
}

function makeRows(mediaLabel: string, lineKey: string, mediaId: string): [Row[], Record<string,string[]>] {
  const stKey = lineKey === "경원선" && mediaId === "m3" ? "경원선조명" : lineKey;
  const stations = STATIONS[stKey] ?? [];
  const imgs: Record<string,string[]> = {};
  const rows: Row[] = [];

  stations.forEach((st, si) => {
    const count = si % 3 === 0 ? 2 : 1;
    for (let ni = 0; ni < count; ni++) {
      const id   = `${mediaId}-${lineKey}-${si}-${ni}`;
      const idx  = (si * 2 + ni) % STATUS_PAT.length;
      const sold = STATUS_PAT[idx] === "판매완료";
      const iAdv = (si * 3 + ni * 7) % ADVERTISERS.length;
      const iMgr = (si + ni * 4)     % MANAGERS.length;
      const iPer = (si + ni)          % PERIODS.length;
      const month = (si % 12) + 1;
      const perMonths = [1,2,3,6,12,1,2,3,1,6][iPer];
      const endM = ((month - 1 + perMonths) % 12) + 1;
      const endY = month + perMonths > 12 ? 2025 : 2024;

      if (si < 3 && ni === 0) imgs[id] = [SAMPLE_IMGS[si % 2]];

      rows.push({
        id,
        station:    st,
        media:      mediaLabel,
        number:     `${String.fromCharCode(65 + (si % 5))}-${ni + 1}`,
        status:     STATUS_PAT[idx],
        advertiser: sold ? ADVERTISERS[iAdv] : "",
        manager:    sold ? MANAGERS[iMgr]    : "",
        period:     sold ? PERIODS[iPer]     : "",
        startDate:  sold ? fmtDate(2024, month, 1)      : "",
        endDate:    sold ? fmtDate(endY, endM, 28)       : "",
        size:       SIZES[si    % SIZES.length],
        adFee:      AD_FEES[si  % AD_FEES.length],
        prodFee:    PROD_FEES[si % PROD_FEES.length],
        note:       NOTES[si    % NOTES.length],
      } as Row);
    }
  });
  return [rows, imgs];
}

function buildInitialData(): AppData {
  const mediaTabs: MediaTab[] = [
    { id:"m1", label:"종합안내도", lines:[{id:"l11",label:"분당선",color:"#F5A200"},{id:"l12",label:"경원선",color:"#77C4A3"}] },
    { id:"m2", label:"소방함",    lines:[{id:"l21",label:"경부선",color:"#0052A4"},{id:"l22",label:"경인선",color:"#5BA3D9"}] },
    { id:"m3", label:"조명",      lines:[{id:"l31",label:"8호선", color:"#E6186C"},{id:"l32",label:"경원선",color:"#77C4A3"}] },
    { id:"m4", label:"승강장",    lines:[{id:"l41",label:"5호선", color:"#996CAC"},{id:"l42",label:"6호선", color:"#CD7C2F"}] },
  ];
  const tableData: AppData["tableData"] = {};
  let stationImages: AppData["stationImages"] = {};
  mediaTabs.forEach(m => {
    tableData[m.id] = {};
    m.lines.forEach(l => {
      const [rows, imgs] = makeRows(m.label, l.label, m.id);
      tableData[m.id][l.id] = rows;
      stationImages = { ...stationImages, ...imgs };
    });
  });
  return { mediaTabs, defaultColumns: DEFAULT_COLS, tabColumns: {}, tableData, stationImages, adminPassword:"0000" };
}

const SK = "subway-ad-v6";
function loadData(): AppData {
  try {
    const r = localStorage.getItem(SK);
    if (r) {
      const d = JSON.parse(r) as AppData & { columns?: Column[] };
      // migrate old format
      if (d.columns && !d.defaultColumns) {
        d.defaultColumns = d.columns;
        delete d.columns;
      }
      if (!d.tabColumns) d.tabColumns = {};
      if (!d.defaultColumns) d.defaultColumns = DEFAULT_COLS;
      return d as AppData;
    }
  } catch {}
  return buildInitialData();
}
function persist(d: AppData) { localStorage.setItem(SK, JSON.stringify(d)); }
function cloneData<T>(d: T): T { return JSON.parse(JSON.stringify(d)); }

// ─── Consecutive row grouping (preserves user order) ───────────────────────────
interface GroupedRow { row: Row; stationSpan: number; showStation: boolean; }

function groupConsecutive(rows: Row[]): GroupedRow[] {
  const result: GroupedRow[] = [];
  let i = 0;
  while (i < rows.length) {
    const station = rows[i].station;
    let j = i + 1;
    while (j < rows.length && rows[j].station === station) j++;
    const span = j - i;
    for (let k = i; k < j; k++) {
      result.push({ row: rows[k], stationSpan: k === i ? span : 0, showStation: k === i });
    }
    i = j;
  }
  return result;
}

// ─── App ────────────────────────────────────────────────────────────────────────
export default function App() {
  const [published, setPublished] = useState<AppData>(() => loadData());
  const [draft,     setDraft]     = useState<AppData>(() => loadData());
  const [isAdmin,   setIsAdmin]   = useState(false);

  const data       = isAdmin ? draft : published;
  const hasUnsaved = isAdmin && JSON.stringify(draft) !== JSON.stringify(published);

  const [activeMedia, setActiveMedia] = useState(0);
  const [activeLine,  setActiveLine]  = useState(0);
  const [search,      setSearch]      = useState("");
  const [highlightId, setHighlightId] = useState<string|null>(null);

  // Pending navigation when leaving an editing table
  type PendingNav = {type:"media";idx:number} | {type:"line";idx:number};
  const [pendingNav, setPendingNav] = useState<PendingNav|null>(null);

  const [showLogin, setShowLogin] = useState(false);
  const [loginPw,   setLoginPw]   = useState("");
  const [loginErr,  setLoginErr]  = useState("");
  const [showPw,    setShowPw]    = useState(false);

  const [showAdmin, setShowAdmin] = useState(false);
  const [adminTab,  setAdminTab]  = useState<"media"|"line"|"password">("media");

  const [editTabId,   setEditTabId]   = useState<string|null>(null);
  const [tabVal,      setTabVal]      = useState("");
  const [newPw1,      setNewPw1]      = useState("");
  const [newPw2,      setNewPw2]      = useState("");
  const [pwMsg,       setPwMsg]       = useState("");

  const [selected,     setSelected]     = useState<Set<string>>(new Set());
  const [imgRow,       setImgRow]       = useState<Row|null>(null);
  const [imgFullSrc,   setImgFullSrc]   = useState<string|null>(null);
  const [dlConfirmSrc, setDlConfirmSrc] = useState<string|null>(null);

  // Non-admin table edit mode (per-table buffer)
  const [editingTables, setEditingTables] = useState<Set<string>>(new Set());
  const [tableBuffer,   setTableBuffer]   = useState<Record<string, Row[]>>({});

  // Inline column editor panel
  const [showColEditor, setShowColEditor] = useState(false);
  const [editColId,     setEditColId]     = useState<string|null>(null);
  const [colLabelVal,   setColLabelVal]   = useState("");

  // Drag-to-reorder
  const [dragRowId,  setDragRowId]  = useState<string|null>(null);
  const [dragOverId, setDragOverId] = useState<string|null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const xlsxInputRef = useRef<HTMLInputElement>(null);
  const hlRowRef     = useRef<HTMLTableRowElement>(null);

  const mediaTabs    = data.mediaTabs;
  const activeMedTab = mediaTabs[activeMedia] ?? mediaTabs[0];
  const lineTab      = activeMedTab?.lines[activeLine] ?? activeMedTab?.lines[0];
  const lineColor    = lineTab?.color ?? "#1a365d";

  const curKey = `${activeMedTab?.id}__${lineTab?.id}`;

  // Resolve columns for current tab (per-tab or fallback to default)
  const curCols: Column[] = data.tabColumns[curKey] ?? data.defaultColumns;

  const isEditingTable = !isAdmin && editingTables.has(curKey);

  const baseRows: Row[]    = data.tableData[activeMedTab?.id]?.[lineTab?.id] ?? [];
  const displayRows: Row[] = isEditingTable ? (tableBuffer[curKey] ?? baseRows) : baseRows;
  const groupedRows        = useMemo(() => groupConsecutive(displayRows), [displayRows]);

  // Search
  const searchResults = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return [];
    const out: Array<Row & {mLabel:string;lLabel:string;mi:number;li:number;lColor:string}> = [];
    data.mediaTabs.forEach((m,mi) => m.lines.forEach((l,li) => {
      (data.tableData[m.id]?.[l.id] ?? []).forEach(r => {
        if (r.station?.toLowerCase().includes(q))
          out.push({ ...r, mLabel:m.label, lLabel:l.label, mi, li, lColor:l.color });
      });
    }));
    return out;
  }, [search, data]);

  useEffect(() => {
    if (highlightId && hlRowRef.current) {
      setTimeout(() => hlRowRef.current?.scrollIntoView({ behavior:"smooth", block:"center" }), 100);
      const t = setTimeout(() => setHighlightId(null), 3500);
      return () => clearTimeout(t);
    }
  }, [highlightId]);

  // ── Draft helpers ────────────────────────────────────────────────────────────
  function mutateDraft(fn: (d: AppData) => void) {
    setDraft(prev => { const c = cloneData(prev); fn(c); return c; });
  }
  function handleSave() { persist(draft); setPublished(cloneData(draft)); }

  // ── Per-tab column helpers ────────────────────────────────────────────────────
  // Ensure the current tab has its own column copy (splitting from default)
  function ensureTabCols(d: AppData) {
    if (!d.tabColumns[curKey]) {
      d.tabColumns[curKey] = cloneData(d.defaultColumns);
    }
  }
  function tabColMutate(fn: (cols: Column[]) => void) {
    mutateDraft(d => { ensureTabCols(d); fn(d.tabColumns[curKey]); });
  }
  function addTabCol() {
    tabColMutate(cols => cols.push({ id:"c"+uid(), key:"col_"+uid(), label:"새 열", type:"text", align:"left" }));
  }
  function removeTabCol(id: string) {
    tabColMutate(cols => { const i = cols.findIndex(c=>c.id===id); if(i>=0) cols.splice(i,1); });
    setEditColId(null);
  }
  function moveTabCol(idx: number, dir: -1|1) {
    const ni = idx+dir;
    tabColMutate(cols => { if(ni<0||ni>=cols.length) return; [cols[idx],cols[ni]]=[cols[ni],cols[idx]]; });
  }
  function updateTabCol(id: string, patch: Partial<Column>) {
    tabColMutate(cols => { const c = cols.find(x=>x.id===id); if(c) Object.assign(c,patch); });
  }

  // ── Admin cell editing (direct to draft, no edit mode needed) ────────────────
  function adminUpdateCell(rowId: string, colKey: string, value: string) {
    if (!activeMedTab || !lineTab) return;
    mutateDraft(d => {
      const row = d.tableData[activeMedTab.id]?.[lineTab.id]?.find(r=>r.id===rowId);
      if (row) row[colKey] = value;
    });
  }

  // ── Non-admin table edit mode ─────────────────────────────────────────────────
  function startTableEdit() {
    setTableBuffer(prev => ({ ...prev, [curKey]: cloneData(baseRows) }));
    setEditingTables(prev => new Set([...prev, curKey]));
  }
  function cancelTableEdit() {
    setEditingTables(prev => { const n=new Set(prev); n.delete(curKey); return n; });
  }
  function saveTableEdit() {
    const buf = tableBuffer[curKey];
    if (!buf || !activeMedTab || !lineTab) return;
    setPublished(prev => { const c=cloneData(prev); c.tableData[activeMedTab.id][lineTab.id]=buf; persist(c); return c; });
    setDraft(prev => { const c=cloneData(prev); c.tableData[activeMedTab.id][lineTab.id]=buf; return c; });
    setEditingTables(prev => { const n=new Set(prev); n.delete(curKey); return n; });
  }
  function updateBufferCell(rowId: string, colKey: string, value: string) {
    setTableBuffer(prev => {
      const buf = [...(prev[curKey] ?? baseRows)];
      const ri = buf.findIndex(r=>r.id===rowId);
      if (ri>=0) buf[ri] = { ...buf[ri], [colKey]: value };
      return { ...prev, [curKey]: buf };
    });
  }

  // ── Apply rows to current table (admin → draft, non-admin → buffer) ─────────
  function applyRows(newRows: Row[]) {
    if (!activeMedTab || !lineTab) return;
    if (isAdmin) {
      mutateDraft(d => { d.tableData[activeMedTab.id][lineTab.id] = newRows; });
    } else {
      setTableBuffer(prev => ({ ...prev, [curKey]: newRows }));
    }
  }

  // ── Excel-style multi-cell paste ──────────────────────────────────────────────
  function handleCellPaste(
    e: React.ClipboardEvent<HTMLInputElement|HTMLTextAreaElement>,
    ri: number,
    ci: number
  ) {
    const text = e.clipboardData.getData("text/plain");
    const lines = text.split(/\r?\n/).filter(l => l.length > 0);
    // If single value with no tabs, let browser handle normally
    if (lines.length <= 1 && !text.includes("\t")) return;

    e.preventDefault();
    const pastedRows = lines.map(l => l.split("\t"));
    const allCols    = curCols;
    const workRows   = cloneData(displayRows);

    pastedRows.forEach((pastedRow, dr) => {
      const targetRi = ri + dr;
      let targetRow: Row;

      if (targetRi < workRows.length) {
        targetRow = { ...workRows[targetRi] };
      } else {
        targetRow = {
          id:"r"+uid(), station:"새 역명", media:activeMedTab?.label||"",
          number:"A-1", status:"판매중", advertiser:"", manager:"",
          period:"", startDate:"", endDate:"", size:"", adFee:"", prodFee:"", note:""
        };
      }

      pastedRow.forEach((val, dc) => {
        const targetCi = ci + dc;
        if (targetCi < allCols.length) {
          targetRow[allCols[targetCi].key] = val.trim();
        }
      });

      if (targetRi < workRows.length) workRows[targetRi] = targetRow;
      else workRows.push(targetRow);
    });

    applyRows(workRows);
  }

  // ── Status toggle — only in edit mode ────────────────────────────────────────
  function toggleStatus(rowId: string) {
    if (!activeMedTab || !lineTab) return;
    // Non-admin: only allowed when table is in edit mode
    if (!isAdmin && !isEditingTable) return;
    if (isEditingTable) {
      const cur = (tableBuffer[curKey] ?? baseRows).find(r=>r.id===rowId);
      if (cur) updateBufferCell(rowId, "status", cur.status==="판매중" ? "판매완료" : "판매중");
      return;
    }
    // Admin direct edit
    adminUpdateCell(rowId, "status",
      (draft.tableData[activeMedTab.id]?.[lineTab.id]?.find(r=>r.id===rowId)?.status ?? "판매중") === "판매중" ? "판매완료" : "판매중"
    );
  }

  // ── Tab navigation with edit-mode guard ───────────────────────────────────────
  function tryNavigate(nav: PendingNav) {
    if (isEditingTable) { setPendingNav(nav); return; }
    applyNav(nav);
  }
  function applyNav(nav: PendingNav) {
    if (nav.type==="media") { setActiveMedia(nav.idx); setActiveLine(0); setSelected(new Set()); setShowColEditor(false); }
    else                    { setActiveLine(nav.idx); setSelected(new Set()); }
    setPendingNav(null);
  }
  function confirmNavSave() { saveTableEdit(); applyNav(pendingNav!); }
  function confirmNavDiscard() { cancelTableEdit(); applyNav(pendingNav!); }

  // ── Drag-to-reorder (admin only) ──────────────────────────────────────────────
  function handleDragStart(e: React.DragEvent, rowId: string) {
    setDragRowId(rowId);
    e.dataTransfer.effectAllowed = "move";
  }
  function handleDragOver(e: React.DragEvent, rowId: string) {
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    if (rowId !== dragOverId) setDragOverId(rowId);
  }
  function handleDrop(e: React.DragEvent, targetId: string) {
    e.preventDefault();
    if (!dragRowId || dragRowId===targetId || !activeMedTab || !lineTab) { setDragRowId(null); setDragOverId(null); return; }
    mutateDraft(d => {
      const rows = d.tableData[activeMedTab.id]?.[lineTab.id];
      if (!rows) return;
      const from = rows.findIndex(r=>r.id===dragRowId);
      const to   = rows.findIndex(r=>r.id===targetId);
      if (from<0||to<0) return;
      const [item] = rows.splice(from, 1);
      rows.splice(to, 0, item);
    });
    setDragRowId(null); setDragOverId(null);
  }
  function handleDragEnd() { setDragRowId(null); setDragOverId(null); }

  // ── Row management ────────────────────────────────────────────────────────────
  function addRow() {
    if (!activeMedTab || !lineTab) return;
    mutateDraft(d => {
      if (!d.tableData[activeMedTab.id]) d.tableData[activeMedTab.id]={};
      if (!d.tableData[activeMedTab.id][lineTab.id]) d.tableData[activeMedTab.id][lineTab.id]=[];
      d.tableData[activeMedTab.id][lineTab.id].push({
        id:"r"+uid(), station:"새 역명", media:activeMedTab.label,
        number:"A-1", status:"판매중", advertiser:"", manager:"", period:"",
        startDate:"", endDate:"", size:"500×700mm", adFee:"0", prodFee:"0", note:""
      });
    });
  }
  function deleteRows(ids: Set<string>) {
    if (!activeMedTab || !lineTab) return;
    mutateDraft(d => { d.tableData[activeMedTab.id][lineTab.id] = d.tableData[activeMedTab.id][lineTab.id].filter(r=>!ids.has(r.id)); });
    setSelected(new Set());
  }
  function bulkStatus(ids: Set<string>, status: Status) {
    if (!activeMedTab || !lineTab) return;
    mutateDraft(d => { d.tableData[activeMedTab.id]?.[lineTab.id]?.forEach(r=>{if(ids.has(r.id)) r.status=status;}); });
    setSelected(new Set());
  }

  // ── Image management ──────────────────────────────────────────────────────────
  function handleImgUpload(e: React.ChangeEvent<HTMLInputElement>, rowId: string) {
    const files = e.target.files; if (!files) return;
    Array.from(files).forEach(f => {
      const reader = new FileReader();
      reader.onload = ev => {
        const url = ev.target?.result as string;
        mutateDraft(d => { if(!d.stationImages[rowId]) d.stationImages[rowId]=[]; d.stationImages[rowId].push(url); });
      };
      reader.readAsDataURL(f);
    });
    e.target.value="";
  }
  function deleteImg(rowId: string, ii: number) {
    mutateDraft(d => { d.stationImages[rowId]?.splice(ii,1); });
  }

  // ── Excel import / export ─────────────────────────────────────────────────────
  function handleXlsxImport(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]; if (!file || !activeMedTab || !lineTab) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target?.result, {type:"binary"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows: Record<string,string>[] = XLSX.utils.sheet_to_json(ws, {defval:""});
        if (!rows.length) { alert("데이터가 없습니다."); return; }
        const newRows: Row[] = rows.map(r => {
          const row: Row = {id:"r"+uid()};
          curCols.forEach(c => { row[c.key] = String(r[c.label]??r[c.key]??""); });
          if (!row.status) row.status="판매중";
          return row;
        });
        mutateDraft(d => {
          if (!d.tableData[activeMedTab.id]) d.tableData[activeMedTab.id]={};
          if (!d.tableData[activeMedTab.id][lineTab.id]) d.tableData[activeMedTab.id][lineTab.id]=[];
          d.tableData[activeMedTab.id][lineTab.id].push(...newRows);
        });
        alert(`${newRows.length}개 행이 추가되었습니다. 저장 버튼을 눌러 적용하세요.`);
      } catch { alert("엑셀 파일을 읽는 중 오류가 발생했습니다."); }
    };
    reader.readAsBinaryString(file);
    e.target.value="";
  }
  function handleXlsxExport() {
    const wsData = [curCols.map(c=>c.label), ...displayRows.map(r=>curCols.map(c=>r[c.key]??""))];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `${activeMedTab?.label}-${lineTab?.label}`.slice(0,31));
    XLSX.writeFile(wb, `${activeMedTab?.label}_${lineTab?.label}.xlsx`);
  }

  // ── Login / logout ────────────────────────────────────────────────────────────
  function handleLogin() {
    if (loginPw === published.adminPassword) {
      setDraft(cloneData(published)); setIsAdmin(true);
      setShowLogin(false); setLoginPw(""); setLoginErr(""); setSelected(new Set());
    } else setLoginErr("비밀번호가 올바르지 않습니다.");
  }
  function handleLogout() {
    if (hasUnsaved && !confirm("저장되지 않은 변경사항이 있습니다. 로그아웃 하시겠습니까?")) return;
    setIsAdmin(false); setShowAdmin(false); setSelected(new Set()); setShowColEditor(false);
  }

  // ── Media tab management ──────────────────────────────────────────────────────
  function addMediaTab() {
    const id="m"+uid();
    mutateDraft(d => { d.mediaTabs.push({id,label:"새 매체",lines:[]}); d.tableData[id]={}; });
  }
  function removeMediaTab(id: string) {
    if (draft.mediaTabs.length<=1) return;
    mutateDraft(d => { d.mediaTabs=d.mediaTabs.filter(m=>m.id!==id); delete d.tableData[id]; });
    if (activeMedia >= draft.mediaTabs.length-1) setActiveMedia(Math.max(0,activeMedia-1));
  }
  function renameMediaTab(id: string, label: string) {
    mutateDraft(d => { const t=d.mediaTabs.find(m=>m.id===id); if(t) t.label=label; }); setEditTabId(null);
  }
  function moveMedia(idx: number, dir: -1|1) {
    const ni=idx+dir; if(ni<0||ni>=draft.mediaTabs.length) return;
    mutateDraft(d => { [d.mediaTabs[idx],d.mediaTabs[ni]]=[d.mediaTabs[ni],d.mediaTabs[idx]]; });
    if(activeMedia===idx) setActiveMedia(ni); else if(activeMedia===ni) setActiveMedia(idx);
  }

  // ── Line tab management ───────────────────────────────────────────────────────
  function addLineTab(mediaId: string) {
    const id="l"+uid();
    mutateDraft(d => {
      d.mediaTabs.find(m=>m.id===mediaId)?.lines.push({id,label:"새 노선",color:"#888888"});
      if (!d.tableData[mediaId]) d.tableData[mediaId]={};
      d.tableData[mediaId][id]=[];
    });
  }
  function removeLineTab(mediaId: string, lineId: string) {
    const m=draft.mediaTabs.find(x=>x.id===mediaId);
    if (!m||m.lines.length<=1) return;
    mutateDraft(d => {
      const m2=d.mediaTabs.find(x=>x.id===mediaId);
      if(m2) m2.lines=m2.lines.filter(l=>l.id!==lineId);
      delete d.tableData[mediaId]?.[lineId];
    });
    if (activeLine >= m.lines.length-1) setActiveLine(Math.max(0,activeLine-1));
  }
  function renameLineTab(mediaId: string, lineId: string, label: string) {
    mutateDraft(d => { const l=d.mediaTabs.find(x=>x.id===mediaId)?.lines.find(x=>x.id===lineId); if(l) l.label=label; }); setEditTabId(null);
  }
  function changeLineColor(mediaId: string, lineId: string, color: string) {
    mutateDraft(d => { const l=d.mediaTabs.find(x=>x.id===mediaId)?.lines.find(x=>x.id===lineId); if(l) l.color=color; });
  }
  function moveLine(mediaId: string, idx: number, dir: -1|1) {
    const m=draft.mediaTabs.find(x=>x.id===mediaId); if(!m) return;
    const ni=idx+dir; if(ni<0||ni>=m.lines.length) return;
    mutateDraft(d => { const m2=d.mediaTabs.find(x=>x.id===mediaId); if(!m2) return; [m2.lines[idx],m2.lines[ni]]=[m2.lines[ni],m2.lines[idx]]; });
    if(activeLine===idx) setActiveLine(ni); else if(activeLine===ni) setActiveLine(idx);
  }

  // ── Password ──────────────────────────────────────────────────────────────────
  function handlePwChange() {
    if (!newPw1) { setPwMsg("새 비밀번호를 입력하세요."); return; }
    if (newPw1.length<4) { setPwMsg("4자 이상 입력하세요."); return; }
    if (newPw1!==newPw2) { setPwMsg("비밀번호가 일치하지 않습니다."); return; }
    mutateDraft(d => { d.adminPassword=newPw1; }); setNewPw1(""); setNewPw2("");
    setPwMsg("✓ 비밀번호가 변경되었습니다. 저장 버튼을 눌러 적용하세요.");
  }

  // ── Checkbox ──────────────────────────────────────────────────────────────────
  const allChecked  = displayRows.length>0 && displayRows.every(r=>selected.has(r.id));
  const someChecked = !allChecked && displayRows.some(r=>selected.has(r.id));
  function toggleSelect(id: string) { setSelected(prev=>{const n=new Set(prev); n.has(id)?n.delete(id):n.add(id); return n;}); }
  function toggleAll() { allChecked ? setSelected(new Set()) : setSelected(new Set(displayRows.map(r=>r.id))); }

  // ── Download image ────────────────────────────────────────────────────────────
  function downloadImage(src: string) {
    const a=document.createElement("a"); a.href=src; a.download=`station-image-${Date.now()}.jpg`; a.target="_blank";
    document.body.appendChild(a); a.click(); document.body.removeChild(a); setDlConfirmSrc(null);
  }

  // ─────────────────────────────────────────────────────────────────────────────
  return (
    <div className="min-h-screen bg-[#f0f4f8] font-['Noto_Sans_KR',sans-serif] flex flex-col text-[15px]">

      {/* ══ HEADER ══ */}
      <header className="bg-[#1a365d] text-white shadow-xl sticky top-0 z-40">
        <div className="max-w-screen-2xl mx-auto px-4 py-3 flex items-center gap-3 flex-wrap">
          <button
            onClick={() => { setActiveMedia(0); setActiveLine(0); setSearch(""); window.scrollTo({top:0,behavior:"smooth"}); }}
            className="flex items-center gap-2 shrink-0 hover:opacity-80 transition-opacity" title="홈으로">
            <div className="w-8 h-8 rounded bg-white/15 flex items-center justify-center">
              <Home className="w-4 h-4"/>
            </div>
            <div className="hidden sm:block">
              <div className="font-bold text-[14px] leading-tight">지하철 광고매체 현황관리</div>
              <div className="text-[10px] text-blue-200">Subway AD Media Management</div>
            </div>
          </button>

          <div className="flex-1 max-w-sm relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-blue-200 pointer-events-none"/>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="역명 검색... (예: 강남, 서울)"
              className="w-full pl-9 pr-8 py-2 bg-white/10 border border-white/20 rounded text-sm text-white placeholder:text-blue-300 focus:outline-none focus:ring-2 focus:ring-white/30"/>
            {search && <button onClick={()=>setSearch("")} className="absolute right-2 top-1/2 -translate-y-1/2 text-blue-200 hover:text-white"><X className="w-4 h-4"/></button>}
          </div>

          <div className="ml-auto flex items-center gap-2 shrink-0">
            {hasUnsaved && (
              <span className="flex items-center gap-1 text-[12px] text-amber-300 bg-amber-500/20 border border-amber-400/30 px-2 py-1 rounded animate-pulse">
                <AlertCircle className="w-3 h-3"/> 미저장
              </span>
            )}
            {isAdmin ? (
              <>
                <button onClick={handleSave} className="flex items-center gap-1 text-xs bg-green-600 hover:bg-green-700 text-white px-3 py-1.5 rounded font-semibold transition-colors">
                  <Save className="w-3.5 h-3.5"/> 저장
                </button>
                <button onClick={()=>setShowAdmin(true)} className="flex items-center gap-1 text-xs bg-amber-500/20 hover:bg-amber-500/30 text-amber-200 border border-amber-400/30 px-3 py-1.5 rounded">
                  <Settings className="w-3.5 h-3.5"/> 관리설정
                </button>
                <span className="hidden sm:flex items-center gap-1 text-[12px] bg-amber-500/20 text-amber-200 border border-amber-400/30 px-2 py-1 rounded">
                  <Shield className="w-3 h-3"/> 관리자 모드
                </span>
                <button onClick={handleLogout} className="text-xs bg-white/10 hover:bg-white/20 border border-white/20 px-3 py-1.5 rounded">로그아웃</button>
              </>
            ) : (
              <button onClick={()=>setShowLogin(true)} className="flex items-center gap-1 text-xs bg-white/10 hover:bg-white/20 border border-white/20 px-3 py-1.5 rounded">
                <Lock className="w-3.5 h-3.5"/> 관리자
              </button>
            )}
          </div>
        </div>
      </header>

      {/* ══ SEARCH RESULTS ══ */}
      {search.trim() && (
        <div className="bg-white border-b border-gray-200 shadow-md z-30">
          <div className="max-w-screen-2xl mx-auto px-4 py-3">
            <p className="text-xs font-bold text-gray-600 mb-2">&ldquo;{search}&rdquo; 검색 결과 ({searchResults.length}건)</p>
            {searchResults.length===0
              ? <p className="text-xs text-gray-400 py-2">검색 결과가 없습니다.</p>
              : (
                <div className="flex flex-col gap-0.5 max-h-72 overflow-y-auto">
                  {searchResults.slice(0,50).map(r => (
                    <button key={r.id}
                      onClick={() => { setActiveMedia(r.mi); setActiveLine(r.li); setSearch(""); setHighlightId(r.id); }}
                      className="flex items-center gap-3 text-left px-3 py-2 rounded hover:bg-blue-50 border border-transparent hover:border-blue-200 transition-all">
                      <span className="w-3 h-3 rounded-full shrink-0" style={{background:r.lColor}}/>
                      <span className="font-semibold text-[#1a365d] w-24 truncate">{r.station}</span>
                      <span className="text-gray-400 text-[12px]">{r.mLabel} · {r.lLabel}</span>
                      <span className="text-gray-500 text-[12px]">번호 {r.number}</span>
                      <span className={`ml-auto px-2 py-0.5 rounded text-[12px] font-bold shrink-0 ${r.status==="판매완료"?"bg-red-100 text-red-700":"bg-green-100 text-green-700"}`}>
                        {r.status}
                      </span>
                    </button>
                  ))}
                  {searchResults.length>50 && <p className="text-xs text-gray-400 text-center py-2">외 {searchResults.length-50}건</p>}
                </div>
              )
            }
          </div>
        </div>
      )}

      {/* ══ MAIN ══ */}
      <div className="flex-1 max-w-screen-2xl mx-auto w-full px-4 py-4">

        {/* Media Tabs */}
        <div className="flex gap-1 overflow-x-auto mb-0">
          {data.mediaTabs.map((m,i) => (
            <button key={m.id} onClick={()=>tryNavigate({type:"media",idx:i})}
              className={`shrink-0 px-5 py-2 rounded-t font-semibold text-xs border transition-all
                ${activeMedia===i?"bg-[#1a365d] text-white border-[#1a365d] shadow-sm":"bg-white text-gray-500 border-gray-200 hover:text-[#1a365d] hover:bg-blue-50"}`}>
              {m.label}
            </button>
          ))}
        </div>

        {/* Card */}
        <div className="bg-white border border-gray-200 rounded-b rounded-tr shadow-sm overflow-hidden">

          {/* Line Tabs */}
          <div className="flex border-b-2 border-gray-200 bg-gray-50 overflow-x-auto">
            {activeMedTab?.lines.map((l,i) => (
              <button key={l.id} onClick={()=>tryNavigate({type:"line",idx:i})}
                className={`flex items-center gap-1.5 px-5 py-2.5 text-xs font-semibold border-b-2 whitespace-nowrap transition-all ${activeLine===i?"bg-white":"border-transparent text-gray-500 hover:bg-gray-100"}`}
                style={activeLine===i?{borderColor:l.color,color:l.color}:{}}>
                <span className="w-2.5 h-2.5 rounded-full" style={{background:l.color}}/>
                {l.label}
              </button>
            ))}
          </div>

          {/* Toolbar */}
          <div className="flex flex-wrap items-center gap-2 px-4 py-2.5 bg-gray-50 border-b border-gray-200">
            <span className="font-['JetBrains_Mono',monospace] text-[12px] text-gray-400">총 {baseRows.length}개</span>
            <span className="flex items-center gap-1 text-[12px] text-gray-500"><span className="w-2 h-2 rounded-full bg-green-500"/>{baseRows.filter(r=>r.status==="판매중").length} 판매중</span>
            <span className="flex items-center gap-1 text-[12px] text-gray-500"><span className="w-2 h-2 rounded-full bg-red-500"/>{baseRows.filter(r=>r.status==="판매완료").length} 판매완료</span>

            <div className="ml-auto flex items-center gap-2 flex-wrap">
              {/* Admin bulk actions */}
              {isAdmin && selected.size>0 && (
                <>
                  <span className="text-amber-700 bg-amber-100 border border-amber-200 px-2 py-0.5 rounded text-[12px]">{selected.size}개 선택</span>
                  <button onClick={()=>bulkStatus(selected,"판매중")} className="text-[12px] text-green-700 bg-green-100 border border-green-200 px-2 py-1 rounded hover:bg-green-200">→ 판매중</button>
                  <button onClick={()=>bulkStatus(selected,"판매완료")} className="text-[12px] text-red-700 bg-red-100 border border-red-200 px-2 py-1 rounded hover:bg-red-200">→ 판매완료</button>
                  <button onClick={()=>{if(confirm(`선택한 ${selected.size}개 행을 삭제하시겠습니까?`)) deleteRows(selected);}}
                    className="flex items-center gap-1 text-[12px] text-red-600 bg-red-50 border border-red-200 px-2 py-1 rounded hover:bg-red-100">
                    <Trash2 className="w-3 h-3"/> 선택 삭제
                  </button>
                </>
              )}

              {isAdmin && (
                <>
                  <button onClick={addRow}
                    className="flex items-center gap-1 text-[12px] bg-[#1a365d] text-white px-2.5 py-1 rounded hover:bg-[#2d4a7a]">
                    <Plus className="w-3.5 h-3.5"/> 행 추가
                  </button>
                  <label className="flex items-center gap-1 text-[12px] bg-emerald-700 text-white px-2.5 py-1 rounded hover:bg-emerald-800 cursor-pointer">
                    <FileSpreadsheet className="w-3.5 h-3.5"/> 엑셀 가져오기
                    <input ref={xlsxInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleXlsxImport}/>
                  </label>
                  <button onClick={handleXlsxExport}
                    className="flex items-center gap-1 text-[12px] bg-teal-700 text-white px-2.5 py-1 rounded hover:bg-teal-800">
                    <Download className="w-3.5 h-3.5"/> 엑셀 다운로드
                  </button>
                  <button onClick={()=>setShowColEditor(v=>!v)}
                    className={`flex items-center gap-1 text-[12px] px-2.5 py-1 rounded border transition-colors
                      ${showColEditor?"bg-indigo-700 text-white border-indigo-700":"bg-white text-indigo-700 border-indigo-300 hover:bg-indigo-50"}`}>
                    <Columns className="w-3.5 h-3.5"/> 열 관리
                  </button>
                  <span className="text-[12px] text-amber-600 bg-amber-50 border border-amber-200 px-2 py-1 rounded flex items-center gap-1">
                    <Edit2 className="w-3 h-3"/> 셀 직접 편집 가능
                  </span>
                </>
              )}

              {/* Non-admin edit buttons */}
              {!isAdmin && (
                isEditingTable ? (
                  <>
                    <button onClick={saveTableEdit} className="flex items-center gap-1 text-[12px] bg-blue-600 text-white px-3 py-1 rounded hover:bg-blue-700 font-semibold">
                      <Save className="w-3.5 h-3.5"/> 저장
                    </button>
                    <button onClick={cancelTableEdit} className="flex items-center gap-1 text-[12px] bg-gray-200 text-gray-700 px-3 py-1 rounded hover:bg-gray-300">
                      <X className="w-3.5 h-3.5"/> 취소
                    </button>
                    <span className="text-[12px] text-amber-600 bg-amber-50 border border-amber-200 px-2 py-1 rounded">편집 중</span>
                  </>
                ) : (
                  <button onClick={startTableEdit} className="flex items-center gap-1 text-[12px] bg-white border border-gray-300 text-gray-600 px-3 py-1 rounded hover:bg-gray-50">
                    <Edit2 className="w-3.5 h-3.5"/> 편집
                  </button>
                )
              )}
            </div>
          </div>

          {/* ── Inline Column Editor Panel (admin only) ── */}
          {isAdmin && showColEditor && (
            <div className="border-b border-indigo-200 bg-indigo-50 px-4 py-3">
              <div className="flex items-center justify-between mb-2">
                <h4 className="text-xs font-bold text-indigo-800 flex items-center gap-1.5">
                  <Columns className="w-3.5 h-3.5"/> 열 관리 — {activeMedTab?.label} / {lineTab?.label}
                  {data.tabColumns[curKey] && <span className="text-[12px] bg-indigo-200 text-indigo-700 px-1.5 py-0.5 rounded">이 탭 전용 설정</span>}
                </h4>
                <div className="flex items-center gap-2">
                  <button onClick={addTabCol}
                    className="flex items-center gap-1 text-[12px] bg-indigo-700 text-white px-2.5 py-1 rounded hover:bg-indigo-800">
                    <Plus className="w-3 h-3"/> 열 추가
                  </button>
                  {data.tabColumns[curKey] && (
                    <button
                      onClick={()=>{if(confirm("이 탭의 열 설정을 기본값으로 초기화하시겠습니까?")) mutateDraft(d=>{delete d.tabColumns[curKey];});}}
                      className="text-[12px] text-red-500 border border-red-200 px-2 py-1 rounded hover:bg-red-50">
                      기본값으로 초기화
                    </button>
                  )}
                </div>
              </div>
              <div className="flex flex-col gap-1 max-h-64 overflow-y-auto">
                {curCols.map((col,i) => (
                  <div key={col.id} className="bg-white border border-indigo-100 rounded">
                    <div className="flex items-center gap-2 px-2.5 py-1.5">
                      <span className="flex-1 text-[12px] font-semibold text-gray-800">{col.label}</span>
                      <span className="text-[12px] bg-blue-100 text-blue-700 px-1.5 py-0.5 rounded font-['JetBrains_Mono',monospace]">{col.type}</span>
                      <span className="text-[12px] bg-gray-100 text-gray-600 px-1.5 py-0.5 rounded">{col.align}</span>
                      {col.fixedWidth && <span className="text-[12px] bg-amber-100 text-amber-700 px-1.5 py-0.5 rounded">{col.fixedWidth}px</span>}
                      <button onClick={()=>{setEditColId(editColId===col.id?null:col.id);setColLabelVal(col.label);}} className="text-indigo-400 hover:text-indigo-600"><Edit2 className="w-3.5 h-3.5"/></button>
                      <button onClick={()=>moveTabCol(i,-1)} disabled={i===0} className="text-gray-400 disabled:opacity-30"><ChevronUp className="w-3.5 h-3.5"/></button>
                      <button onClick={()=>moveTabCol(i,1)} disabled={i===curCols.length-1} className="text-gray-400 disabled:opacity-30"><ChevronDown className="w-3.5 h-3.5"/></button>
                      <button onClick={()=>{if(confirm("이 열을 삭제하시겠습니까?")) removeTabCol(col.id);}} disabled={curCols.length<=1} className="text-red-400 hover:text-red-600 disabled:opacity-30"><Trash2 className="w-3.5 h-3.5"/></button>
                    </div>
                    {editColId===col.id && (
                      <div className="grid grid-cols-2 gap-2 px-3 pb-2 pt-1 border-t border-indigo-100 bg-indigo-50/50">
                        <div>
                          <label className="text-[12px] text-gray-500 block mb-0.5">열 이름</label>
                          <div className="flex gap-1">
                            <input value={colLabelVal} onChange={e=>setColLabelVal(e.target.value)}
                              className="flex-1 text-xs border border-gray-200 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-400"/>
                            <button onClick={()=>{updateTabCol(col.id,{label:colLabelVal});setEditColId(null);}}
                              className="text-xs bg-green-600 text-white px-2 py-1 rounded hover:bg-green-700">적용</button>
                          </div>
                        </div>
                        <div>
                          <label className="text-[12px] text-gray-500 block mb-0.5">타입</label>
                          <select value={col.type} onChange={e=>updateTabCol(col.id,{type:e.target.value as ColType})}
                            className="w-full text-xs border border-gray-200 rounded px-2 py-1 bg-white focus:outline-none">
                            <option value="text">text — 텍스트</option>
                            <option value="number">number — 숫자</option>
                            <option value="date">date — 날짜 (달력)</option>
                            <option value="status">status — 판매현황</option>
                            <option value="station">station — 역명</option>
                            <option value="note">note — 비고 (줄바꿈)</option>
                          </select>
                        </div>
                        <div>
                          <label className="text-[12px] text-gray-500 block mb-0.5">정렬</label>
                          <select value={col.align} onChange={e=>updateTabCol(col.id,{align:e.target.value as ColAlign})}
                            className="w-full text-xs border border-gray-200 rounded px-2 py-1 bg-white focus:outline-none">
                            <option value="left">좌측</option>
                            <option value="center">가운데</option>
                            <option value="right">우측</option>
                          </select>
                        </div>
                        <div>
                          <label className="text-[12px] text-gray-500 block mb-0.5">고정 너비 (px)</label>
                          <input type="number" value={col.fixedWidth??""} placeholder="없음"
                            onChange={e=>updateTabCol(col.id,{fixedWidth:e.target.value?Number(e.target.value):undefined})}
                            className="w-full text-xs border border-gray-200 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-400"/>
                        </div>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Table */}
          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-xs">
              <thead>
                <tr className="bg-[#1a365d] text-white">
                  {isAdmin && (
                    <th className="px-2 py-2.5 w-10 text-center shrink-0 border-r border-white/10">
                      <button onClick={toggleAll} className="text-white/70 hover:text-white">
                        {allChecked ? <CheckSquare className="w-4 h-4"/> : someChecked ? <CheckSquare className="w-4 h-4 opacity-50"/> : <Square className="w-4 h-4"/>}
                      </button>
                    </th>
                  )}
                  {isAdmin && <th className="px-1 py-2.5 w-6 border-r border-white/10"/>}
                  {curCols.map(col => (
                    <th key={col.id}
                      className={`px-3 py-2.5 font-semibold border-r border-white/10 last:border-r-0 whitespace-nowrap text-${col.align}`}
                      style={col.fixedWidth?{width:col.fixedWidth,minWidth:col.fixedWidth,maxWidth:col.fixedWidth}:undefined}>
                      {col.label}
                    </th>
                  ))}
                  {isAdmin && <th className="px-2 py-2.5 w-10 text-center">삭제</th>}
                </tr>
              </thead>
              <tbody>
                {displayRows.length===0
                  ? <tr><td colSpan={curCols.length+(isAdmin?4:0)} className="py-14 text-center text-gray-400">데이터가 없습니다.{isAdmin&&" 행 추가 버튼 또는 엑셀 가져오기로 추가하세요."}</td></tr>
                  : (isAdmin ? displayRows.map((row,ri) => ({row,stationSpan:1,showStation:true,ri})) : groupedRows.map((g,ri)=>({...g,ri}))).map(({row,stationSpan,showStation,ri}) => {
                      const effectiveStatus = getEffectiveStatus(row);
                      const isSold = effectiveStatus === "판매완료";
                      const isSel  = selected.has(row.id);
                      const isHL   = highlightId===row.id;
                      const isDrag = dragRowId===row.id;
                      const isOver = dragOverId===row.id;

                      return (
                        <tr
                          key={row.id}
                          ref={isHL ? hlRowRef : undefined}
                          draggable={isAdmin}
                          onDragStart={isAdmin ? e=>handleDragStart(e,row.id) : undefined}
                          onDragOver={isAdmin ? e=>handleDragOver(e,row.id) : undefined}
                          onDrop={isAdmin ? e=>handleDrop(e,row.id) : undefined}
                          onDragEnd={isAdmin ? handleDragEnd : undefined}
                          className={[
                            "border-b border-gray-100 transition-colors",
                            isHL ? "outline outline-2 outline-yellow-400 bg-yellow-50/80" : "",
                            !isHL && isSel ? "bg-blue-50" : "",
                            !isHL && !isSel && isSold ? "bg-red-50/40" : "",
                            !isHL && !isSel && !isSold && ri%2===0 ? "bg-white" : "",
                            !isHL && !isSel && !isSold && ri%2===1 ? "bg-gray-50/40" : "",
                            isDrag ? "opacity-40" : "",
                            isOver ? "border-t-2 border-t-blue-500" : "",
                            isAdmin ? "cursor-default" : "",
                          ].join(" ")}>

                          {isAdmin && (
                            <td className="px-2 py-2 text-center w-10 align-top border-r border-gray-100">
                              <button onClick={()=>toggleSelect(row.id)}>
                                {isSel ? <CheckSquare className="w-4 h-4 text-blue-500"/> : <Square className="w-4 h-4 text-gray-300 hover:text-gray-500"/>}
                              </button>
                            </td>
                          )}
                          {isAdmin && (
                            <td className="px-1 py-2 w-6 align-top border-r border-gray-100 cursor-grab active:cursor-grabbing text-gray-300 hover:text-gray-500">
                              <GripVertical className="w-3.5 h-3.5"/>
                            </td>
                          )}

                          {curCols.map((col, ci) => {
                            if (col.type==="station" && !showStation) return null;

                            const tdStyle: React.CSSProperties = col.fixedWidth
                              ? {width:col.fixedWidth,minWidth:col.fixedWidth,maxWidth:col.fixedWidth}
                              : {};
                            const tdClass = [
                              "px-3 py-2 border-r border-gray-100 last:border-r-0 align-top",
                              col.type==="note" ? "whitespace-normal break-words" : "whitespace-nowrap",
                              col.align==="right" ? "text-right font-['JetBrains_Mono',monospace] text-gray-700" : col.align==="center" ? "text-center" : "",
                            ].join(" ");
                            const spanProp = col.type==="station" && stationSpan>1 ? {rowSpan:stationSpan} : {};

                            // Decide if this cell is in edit mode
                            const inEdit = isAdmin || isEditingTable;

                            const cellVal = row[col.key] ?? "";
                            const onCellChange = (val: string) => {
                              if (isAdmin) adminUpdateCell(row.id, col.key, val);
                              else updateBufferCell(row.id, col.key, val);
                            };

                            return (
                              <td key={col.id} className={tdClass} style={tdStyle} {...spanProp}>

                                {col.type==="status" && (() => {
                                  const canToggle = isAdmin || isEditingTable;
                                  return (
                                    <button
                                      onClick={canToggle ? ()=>toggleStatus(row.id) : undefined}
                                      title={canToggle ? undefined : "편집 모드에서만 변경 가능합니다"}
                                      className={[
                                        "px-2 py-0.5 rounded border font-bold text-[12px] whitespace-nowrap w-full transition-all",
                                        isSold ? "bg-red-100 text-red-700 border-red-300" : "bg-green-100 text-green-700 border-green-300",
                                        canToggle ? (isSold ? "hover:bg-red-200 cursor-pointer" : "hover:bg-green-200 cursor-pointer") : "cursor-default opacity-80",
                                      ].join(" ")}>
                                      {effectiveStatus}
                                    </button>
                                  );
                                })()}

                                {col.type==="station" && (
                                  <div className="flex items-center gap-1.5">
                                    <span className="w-2 h-2 rounded-full shrink-0" style={{background:lineColor}}/>
                                    {inEdit ? (
                                      <input value={cellVal} onChange={e=>onCellChange(e.target.value)}
                                        data-cell-row={ri} data-cell-col={ci}
                                        onKeyDown={e=>navigateCell(e,ri,ci)}
                                        onPaste={e=>handleCellPaste(e,ri,ci)}
                                        className="border border-blue-300 rounded px-1.5 py-0.5 bg-yellow-50/80 focus:outline-none focus:ring-1 focus:ring-blue-400 w-28"/>
                                    ) : (
                                      <button onClick={()=>setImgRow(row)} className="text-[#1a365d] hover:text-blue-600 font-semibold hover:underline flex items-center gap-1">
                                        {row[col.key]}
                                        <ImageIcon className="w-3 h-3 text-gray-300 hover:text-blue-400"/>
                                      </button>
                                    )}
                                  </div>
                                )}

                                {col.type==="date" && (
                                  inEdit ? (
                                    <input type="date" value={parseDate(cellVal)}
                                      onChange={e=>onCellChange(e.target.value)}
                                      onBlur={e=>{ const n=parseDate(e.target.value); if(n!==cellVal) onCellChange(n); }}
                                      data-cell-row={ri} data-cell-col={ci}
                                      onKeyDown={e=>navigateCell(e,ri,ci)}
                                      onPaste={e=>handleCellPaste(e,ri,ci)}
                                      className="border border-blue-300 rounded px-1.5 py-0.5 bg-yellow-50/80 focus:outline-none focus:ring-1 focus:ring-blue-400 text-[14px] w-full"/>
                                  ) : (
                                    <span className="text-[14px]">{parseDate(cellVal)||<span className="text-gray-300">—</span>}</span>
                                  )
                                )}

                                {col.type==="note" && (
                                  inEdit ? (
                                    <textarea value={cellVal} onChange={e=>onCellChange(e.target.value)}
                                      data-cell-row={ri} data-cell-col={ci}
                                      onKeyDown={e=>navigateCell(e as unknown as React.KeyboardEvent<HTMLInputElement|HTMLTextAreaElement>,ri,ci)}
                                      onPaste={e=>handleCellPaste(e as unknown as React.ClipboardEvent<HTMLInputElement|HTMLTextAreaElement>,ri,ci)}
                                      rows={2} className="w-full border border-blue-300 rounded px-1.5 py-0.5 bg-yellow-50/80 focus:outline-none focus:ring-1 focus:ring-blue-400 text-[14px] resize-none"/>
                                  ) : (
                                    <span className="text-[14px] text-gray-600 leading-relaxed">{cellVal}</span>
                                  )
                                )}

                                {(col.type==="text"||col.type==="number") && (() => {
                                  const isNum = col.type==="number";
                                  const displayVal = isNum ? formatNumber(cellVal) : cellVal;
                                  const editVal   = isNum ? stripCommas(cellVal)  : cellVal;
                                  return inEdit ? (
                                    <input value={editVal} onChange={e=>onCellChange(isNum ? stripCommas(e.target.value) : e.target.value)}
                                      onBlur={isNum ? e=>{ const f=formatNumber(e.target.value); onCellChange(stripCommas(f)||e.target.value); } : undefined}
                                      data-cell-row={ri} data-cell-col={ci}
                                      onKeyDown={e=>navigateCell(e,ri,ci)}
                                      onPaste={e=>handleCellPaste(e,ri,ci)}
                                      className="border border-blue-300 rounded px-1.5 py-0.5 bg-yellow-50/80 focus:outline-none focus:ring-1 focus:ring-blue-400 w-full min-w-[60px] text-[14px]"/>
                                  ) : (
                                    <span className="text-[14px]">{displayVal||<span className="text-gray-300">—</span>}</span>
                                  );
                                })()}
                              </td>
                            );
                          })}

                          {isAdmin && (
                            <td className="px-2 py-2 text-center w-10 align-top">
                              <button onClick={()=>{if(confirm("이 행을 삭제하시겠습니까?")) deleteRows(new Set([row.id]));}}
                                className="text-red-400 hover:text-red-600"><Trash2 className="w-3.5 h-3.5"/></button>
                            </td>
                          )}
                        </tr>
                      );
                    })
                }
              </tbody>
            </table>
          </div>

          {/* Legend */}
          <div className="flex flex-wrap items-center gap-4 px-4 py-2 border-t border-gray-100 bg-gray-50 text-[12px] text-gray-400">
            <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-green-100 border border-green-300"/>판매중: 광고 가능</span>
            <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-red-100 border border-red-300"/>판매완료: 광고 진행 중</span>
            <span className="flex items-center gap-1"><ImageIcon className="w-3 h-3 text-blue-300"/>역명 클릭 시 이미지 조회</span>
            {isAdmin && <span className="ml-auto text-amber-600 font-semibold flex items-center gap-1"><Shield className="w-3 h-3"/>관리자: 셀 직접 편집 · 행 드래그로 순서 변경 · 상단 저장으로 적용</span>}
          </div>
        </div>
      </div>

      <footer className="border-t border-gray-200 bg-white text-center text-[12px] text-gray-400 py-2">
        지하철 광고매체 현황관리 시스템 · Subway AD Media Management System
      </footer>

      {/* ══ STATION IMAGE POPUP ══ */}
      {imgRow && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50 p-4" onClick={()=>setImgRow(null)}>
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg max-h-[85vh] flex flex-col" onClick={e=>e.stopPropagation()}>
            <div className="flex items-center justify-between px-5 py-3 bg-[#1a365d] text-white rounded-t-xl shrink-0">
              <div className="flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full" style={{background:lineColor}}/>
                <span className="font-bold">{imgRow.station}</span>
                <span className="text-blue-200 text-[12px]">· {activeMedTab?.label} · {lineTab?.label}</span>
              </div>
              <button onClick={()=>setImgRow(null)} className="text-blue-200 hover:text-white"><X className="w-5 h-5"/></button>
            </div>
            <div className="flex-1 overflow-y-auto p-4">
              {!(draft.stationImages[imgRow.id]?.length) ? (
                <div className="flex flex-col items-center py-12 text-gray-300">
                  <ImageIcon className="w-14 h-14 mb-2"/>
                  <p className="text-sm text-gray-400">등록된 이미지가 없습니다.</p>
                  {isAdmin && <p className="text-xs text-gray-400 mt-1">아래 버튼으로 이미지를 추가하세요.</p>}
                </div>
              ) : (
                <div className="space-y-3">
                  {draft.stationImages[imgRow.id].map((src,ii) => (
                    <div key={ii} className="relative group rounded-lg overflow-hidden border border-gray-200 bg-gray-100">
                      <img src={src} alt={`${imgRow.station} ${ii+1}`}
                        className="w-full object-cover max-h-60 cursor-zoom-in"
                        onClick={()=>setImgFullSrc(src)}
                        onContextMenu={e=>{e.preventDefault();setDlConfirmSrc(src);}}/>
                      <div className="absolute bottom-2 left-2 opacity-0 group-hover:opacity-100 flex gap-1 transition-opacity">
                        <button onClick={()=>setImgFullSrc(src)} className="bg-black/60 text-white rounded px-2 py-1 text-[12px] hover:bg-black/80">크게 보기</button>
                        <button onClick={()=>setDlConfirmSrc(src)} className="bg-black/60 text-white rounded px-2 py-1 text-[12px] flex items-center gap-1 hover:bg-black/80"><Download className="w-3 h-3"/> 다운로드</button>
                      </div>
                      {isAdmin && (
                        <button onClick={()=>deleteImg(imgRow.id,ii)}
                          className="absolute top-2 right-2 bg-red-600 text-white rounded-full p-1 opacity-0 group-hover:opacity-100 transition-opacity shadow hover:bg-red-700">
                          <Trash2 className="w-3.5 h-3.5"/>
                        </button>
                      )}
                    </div>
                  ))}
                </div>
              )}
              {isAdmin && (
                <div className="mt-3 pt-3 border-t border-gray-100">
                  <input ref={fileInputRef} type="file" accept="image/*" multiple className="hidden" onChange={e=>handleImgUpload(e,imgRow.id)}/>
                  <button onClick={()=>fileInputRef.current?.click()}
                    className="w-full flex items-center justify-center gap-2 py-3 border-2 border-dashed border-blue-300 rounded-lg text-blue-500 hover:bg-blue-50 transition-colors text-sm font-medium">
                    <Upload className="w-4 h-4"/> 이미지 추가 (다중 선택 가능)
                  </button>
                  <p className="text-[12px] text-gray-400 text-center mt-1">업로드 후 저장 버튼을 눌러야 적용됩니다.</p>
                </div>
              )}
            </div>
            <div className="px-5 py-2 border-t border-gray-100 bg-gray-50 rounded-b-xl flex items-center justify-between text-[12px] text-gray-400">
              <span>{draft.stationImages[imgRow.id]?.length ?? 0}개 이미지</span>
              <span>클릭: 크게 보기 · 우클릭: 다운로드</span>
            </div>
          </div>
        </div>
      )}

      {/* ══ FULL-SIZE IMAGE ══ */}
      {imgFullSrc && (
        <div className="fixed inset-0 bg-black/90 flex items-center justify-center z-[60] p-4 cursor-zoom-out" onClick={()=>setImgFullSrc(null)}>
          <img src={imgFullSrc} alt="full size" className="max-w-full max-h-full object-contain rounded-lg shadow-2xl"
            onContextMenu={e=>{e.preventDefault();setDlConfirmSrc(imgFullSrc);}} onClick={e=>e.stopPropagation()}/>
          <button onClick={()=>setImgFullSrc(null)} className="absolute top-4 right-4 bg-white/20 hover:bg-white/40 text-white rounded-full p-2"><X className="w-5 h-5"/></button>
          <button onClick={()=>setDlConfirmSrc(imgFullSrc)} className="absolute bottom-4 right-4 bg-white/20 hover:bg-white/40 text-white rounded px-3 py-2 flex items-center gap-2 text-sm"><Download className="w-4 h-4"/> 다운로드</button>
        </div>
      )}

      {/* ══ DOWNLOAD CONFIRM ══ */}
      {dlConfirmSrc && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[70]" onClick={()=>setDlConfirmSrc(null)}>
          <div className="bg-white rounded-xl shadow-2xl p-6 w-72" onClick={e=>e.stopPropagation()}>
            <h3 className="font-bold text-gray-900 mb-2 flex items-center gap-2"><Download className="w-4 h-4 text-[#1a365d]"/> 이미지 다운로드</h3>
            <p className="text-sm text-gray-500 mb-5">이 이미지를 다운로드 하시겠습니까?</p>
            <div className="flex gap-2">
              <button onClick={()=>downloadImage(dlConfirmSrc)} className="flex-1 bg-[#1a365d] text-white py-2 rounded font-semibold text-sm hover:bg-[#2d4a7a]">다운로드</button>
              <button onClick={()=>setDlConfirmSrc(null)} className="flex-1 bg-gray-100 text-gray-700 py-2 rounded font-semibold text-sm hover:bg-gray-200">취소</button>
            </div>
          </div>
        </div>
      )}

      {/* ══ UNSAVED EDIT NAVIGATION CONFIRM ══ */}
      {pendingNav && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[80]">
          <div className="bg-white rounded-xl shadow-2xl w-80 p-6">
            <h3 className="font-bold text-gray-900 mb-1 flex items-center gap-2">
              <AlertCircle className="w-4 h-4 text-amber-500"/> 편집 중인 내용이 있습니다
            </h3>
            <p className="text-sm text-gray-500 mb-5">
              저장되지 않은 변경사항이 있습니다.<br/>저장하고 이동하시겠습니까?
            </p>
            <div className="flex flex-col gap-2">
              <button onClick={confirmNavSave}
                className="w-full bg-blue-600 text-white py-2 rounded font-semibold text-sm hover:bg-blue-700">
                저장하고 이동
              </button>
              <button onClick={confirmNavDiscard}
                className="w-full bg-gray-100 text-gray-700 py-2 rounded font-semibold text-sm hover:bg-gray-200">
                저장하지 않고 이동
              </button>
              <button onClick={()=>setPendingNav(null)}
                className="w-full bg-white border border-gray-200 text-gray-500 py-2 rounded text-sm hover:bg-gray-50">
                취소 (계속 편집)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ══ ADMIN LOGIN ══ */}
      {showLogin && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50" onClick={()=>{setShowLogin(false);setLoginPw("");setLoginErr("");}}>
          <div className="bg-white rounded-xl shadow-2xl w-80 p-6" onClick={e=>e.stopPropagation()}>
            <div className="flex items-center justify-between mb-4">
              <h2 className="font-bold text-gray-900 flex items-center gap-2"><Lock className="w-4 h-4 text-[#1a365d]"/> 관리자 로그인</h2>
              <button onClick={()=>{setShowLogin(false);setLoginPw("");setLoginErr("");}}><X className="w-4 h-4 text-gray-400"/></button>
            </div>
            <div className="relative mb-3">
              <input type={showPw?"text":"password"} value={loginPw} onChange={e=>setLoginPw(e.target.value)}
                placeholder="관리자 비밀번호" autoFocus onKeyDown={e=>e.key==="Enter"&&handleLogin()}
                className="w-full border border-gray-200 rounded px-3 py-2.5 pr-10 text-sm focus:outline-none focus:ring-2 focus:ring-[#1a365d]/30"/>
              <button onClick={()=>setShowPw(v=>!v)} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400">
                {showPw?<EyeOff className="w-4 h-4"/>:<Eye className="w-4 h-4"/>}
              </button>
            </div>
            {loginErr && <p className="text-xs text-red-600 mb-3">{loginErr}</p>}
            <button onClick={handleLogin} className="w-full bg-[#1a365d] text-white py-2.5 rounded font-semibold hover:bg-[#2d4a7a] text-sm">로그인</button>
          </div>
        </div>
      )}

      {/* ══ ADMIN MANAGEMENT MODAL ══ */}
      {showAdmin && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-3xl max-h-[90vh] flex flex-col">
            <div className="flex items-center justify-between px-5 py-3 bg-[#1a365d] text-white rounded-t-xl shrink-0">
              <h2 className="font-bold flex items-center gap-2"><Settings className="w-4 h-4"/> 관리자 설정</h2>
              <div className="flex items-center gap-2">
                <button onClick={()=>{handleSave();}} className="flex items-center gap-1 text-xs bg-green-500 hover:bg-green-600 text-white px-3 py-1.5 rounded font-semibold">
                  <Save className="w-3.5 h-3.5"/> 저장
                </button>
                <button onClick={()=>setShowAdmin(false)} className="text-blue-200 hover:text-white"><X className="w-5 h-5"/></button>
              </div>
            </div>

            <div className="flex flex-1 overflow-hidden">
              <div className="w-36 bg-gray-50 border-r border-gray-200 flex flex-col shrink-0">
                {([{key:"media",label:"매체 탭"},{key:"line",label:"노선 탭"},{key:"password",label:"비밀번호"}] as const).map(t => (
                  <button key={t.key} onClick={()=>setAdminTab(t.key)}
                    className={`px-4 py-3 text-xs font-semibold text-left border-b border-gray-200 transition-colors ${adminTab===t.key?"bg-[#1a365d] text-white":"text-gray-600 hover:bg-gray-100"}`}>
                    {t.label}
                  </button>
                ))}
                {hasUnsaved && (
                  <div className="p-2 mt-auto">
                    <div className="flex items-start gap-1 text-[12px] text-amber-700 bg-amber-50 border border-amber-200 rounded p-2">
                      <AlertCircle className="w-3 h-3 shrink-0 mt-0.5"/> 저장하지 않은 변경사항이 있습니다.
                    </div>
                  </div>
                )}
              </div>

              <div className="flex-1 overflow-y-auto p-5">

                {adminTab==="media" && (
                  <div>
                    <div className="flex items-center justify-between mb-3">
                      <h3 className="font-bold text-gray-800">매체 탭 관리</h3>
                      <button onClick={addMediaTab} className="flex items-center gap-1 text-xs bg-[#1a365d] text-white px-2.5 py-1.5 rounded hover:bg-[#2d4a7a]"><Plus className="w-3.5 h-3.5"/> 탭 추가</button>
                    </div>
                    <p className="text-[12px] text-gray-400 mb-3">변경 후 저장 버튼을 눌러야 적용됩니다.</p>
                    <div className="space-y-2">
                      {draft.mediaTabs.map((m,i) => (
                        <div key={m.id} className="flex items-center gap-2 bg-gray-50 border border-gray-200 rounded px-3 py-2">
                          {editTabId===m.id ? (
                            <>
                              <input autoFocus value={tabVal} onChange={e=>setTabVal(e.target.value)}
                                onKeyDown={e=>{if(e.key==="Enter")renameMediaTab(m.id,tabVal);if(e.key==="Escape")setEditTabId(null);}}
                                className="flex-1 text-xs border border-blue-300 rounded px-2 py-1 focus:outline-none"/>
                              <button onClick={()=>renameMediaTab(m.id,tabVal)} className="text-green-600"><Check className="w-4 h-4"/></button>
                              <button onClick={()=>setEditTabId(null)} className="text-gray-400"><X className="w-4 h-4"/></button>
                            </>
                          ) : (
                            <>
                              <span className="flex-1 text-xs font-medium text-gray-800">{m.label}</span>
                              <span className="text-[12px] text-gray-400 bg-gray-200 px-1.5 py-0.5 rounded">{m.lines.length}개 노선</span>
                              <button onClick={()=>{setEditTabId(m.id);setTabVal(m.label);}} className="text-blue-400 hover:text-blue-600"><Edit2 className="w-3.5 h-3.5"/></button>
                              <button onClick={()=>moveMedia(i,-1)} disabled={i===0} className="text-gray-400 disabled:opacity-30"><ChevronUp className="w-3.5 h-3.5"/></button>
                              <button onClick={()=>moveMedia(i,1)} disabled={i===draft.mediaTabs.length-1} className="text-gray-400 disabled:opacity-30"><ChevronDown className="w-3.5 h-3.5"/></button>
                              <button onClick={()=>{if(confirm(`"${m.label}" 매체 탭과 모든 데이터를 삭제하시겠습니까?`))removeMediaTab(m.id);}} disabled={draft.mediaTabs.length<=1} className="text-red-400 hover:text-red-600 disabled:opacity-30"><Trash2 className="w-3.5 h-3.5"/></button>
                            </>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {adminTab==="line" && (
                  <div>
                    <h3 className="font-bold text-gray-800 mb-3">노선 탭 관리</h3>
                    <div className="space-y-5">
                      {draft.mediaTabs.map(m => (
                        <div key={m.id} className="border border-gray-200 rounded-lg overflow-hidden">
                          <div className="flex items-center justify-between px-3 py-2 bg-[#1a365d] text-white">
                            <span className="text-xs font-bold">{m.label}</span>
                            <button onClick={()=>addLineTab(m.id)} className="flex items-center gap-1 text-[12px] bg-white/20 hover:bg-white/30 px-2 py-1 rounded"><Plus className="w-3 h-3"/> 노선 추가</button>
                          </div>
                          <div className="p-2 space-y-1 bg-gray-50">
                            {m.lines.map((l,i) => (
                              <div key={l.id} className="flex items-center gap-2 bg-white border border-gray-200 rounded px-2.5 py-1.5">
                                {editTabId===l.id ? (
                                  <>
                                    <input autoFocus value={tabVal} onChange={e=>setTabVal(e.target.value)}
                                      onKeyDown={e=>{if(e.key==="Enter")renameLineTab(m.id,l.id,tabVal);if(e.key==="Escape")setEditTabId(null);}}
                                      className="flex-1 text-xs border border-blue-300 rounded px-2 py-1 focus:outline-none"/>
                                    <input type="color" value={l.color} onChange={e=>changeLineColor(m.id,l.id,e.target.value)} className="w-7 h-7 rounded cursor-pointer p-0.5 border border-gray-200"/>
                                    <button onClick={()=>renameLineTab(m.id,l.id,tabVal)} className="text-green-600"><Check className="w-4 h-4"/></button>
                                    <button onClick={()=>setEditTabId(null)} className="text-gray-400"><X className="w-4 h-4"/></button>
                                  </>
                                ) : (
                                  <>
                                    <span className="w-3 h-3 rounded-full shrink-0" style={{background:l.color}}/>
                                    <span className="flex-1 text-xs font-medium text-gray-800">{l.label}</span>
                                    <span className="text-[12px] text-gray-400 font-['JetBrains_Mono',monospace]">{draft.tableData[m.id]?.[l.id]?.length ?? 0}역</span>
                                    <button onClick={()=>{setEditTabId(l.id);setTabVal(l.label);}} className="text-blue-400"><Edit2 className="w-3 h-3"/></button>
                                    <button onClick={()=>moveLine(m.id,i,-1)} disabled={i===0} className="text-gray-400 disabled:opacity-30"><ChevronUp className="w-3 h-3"/></button>
                                    <button onClick={()=>moveLine(m.id,i,1)} disabled={i===m.lines.length-1} className="text-gray-400 disabled:opacity-30"><ChevronDown className="w-3 h-3"/></button>
                                    <button onClick={()=>{if(confirm(`"${l.label}" 노선과 데이터를 삭제하시겠습니까?`))removeLineTab(m.id,l.id);}} disabled={m.lines.length<=1} className="text-red-400 disabled:opacity-30"><Trash2 className="w-3 h-3"/></button>
                                  </>
                                )}
                              </div>
                            ))}
                            {m.lines.length===0 && <p className="text-xs text-gray-400 text-center py-2">노선이 없습니다. 추가해주세요.</p>}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {adminTab==="password" && (
                  <div className="max-w-xs">
                    <h3 className="font-bold text-gray-800 mb-1 flex items-center gap-2"><KeyRound className="w-4 h-4"/> 비밀번호 관리</h3>
                    <p className="text-[12px] text-gray-400 mb-4">변경 후 반드시 저장 버튼을 눌러야 적용됩니다.</p>
                    <div className="space-y-3">
                      <div>
                        <label className="text-xs font-semibold text-gray-700 block mb-1">새 비밀번호 (4자 이상)</label>
                        <input type="password" value={newPw1} onChange={e=>setNewPw1(e.target.value)} placeholder="새 비밀번호"
                          className="w-full border border-gray-200 rounded px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300"/>
                      </div>
                      <div>
                        <label className="text-xs font-semibold text-gray-700 block mb-1">비밀번호 확인</label>
                        <input type="password" value={newPw2} onChange={e=>setNewPw2(e.target.value)} placeholder="동일하게 입력"
                          onKeyDown={e=>e.key==="Enter"&&handlePwChange()}
                          className="w-full border border-gray-200 rounded px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300"/>
                      </div>
                      {pwMsg && <p className={`text-xs ${pwMsg.startsWith("✓")?"text-green-600":"text-red-600"}`}>{pwMsg}</p>}
                      <button onClick={handlePwChange} className="w-full bg-[#1a365d] text-white py-2.5 rounded text-sm font-semibold hover:bg-[#2d4a7a]">비밀번호 변경</button>
                    </div>
                    <div className="mt-6 pt-4 border-t border-gray-100">
                      <p className="text-[12px] text-gray-500 font-semibold mb-2">⚠ 위험 구역</p>
                      <button onClick={()=>{if(confirm("모든 데이터가 초기 상태로 돌아갑니다. 계속하시겠습니까?")){const d=buildInitialData();setPublished(d);setDraft(d);persist(d);setShowAdmin(false);}}}
                        className="flex items-center gap-1 text-xs text-red-500 hover:text-red-700 border border-red-200 hover:border-red-400 px-3 py-1.5 rounded">
                        <RefreshCw className="w-3.5 h-3.5"/> 전체 데이터 초기화
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
