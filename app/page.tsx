'use client';

import { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { BarChart3, CalendarDays, Check, ChevronDown, Download, FileSpreadsheet, Filter, Info, RotateCcw, Search, ShieldCheck, Sparkles, Upload, X } from 'lucide-react';

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const GAME_NAMES = ['FB','GB','GL','DS'];
const MONTH_ALIASES: Record<string, number> = Object.fromEntries(MONTHS.flatMap((m, i) => [[m.toLowerCase(), i + 1], [m.slice(0, 3).toLowerCase(), i + 1]]));
type DayRecord = { year: number; month: number; day: number; games: string[] };
type Match = { current: string; currentGame: number; compared: string; comparedGame: number; kind: 'Exact'|'Reverse'; source: 'Previous year'|'Previous month'; date: string };
type AnalysisRow = DayRecord & { matches: Match[] };
type FileState = { name: string; records: DayRecord[] } | null;
type Candidate = { value: string; score: number; evidence: string[] };
type ForecastDay = { month: number; day: number; overall: Candidate[]; byGame: Candidate[][]; actual?: string[] };
type StrategyRow = { forecast: ForecastDay; score: number; candidate: string; hit: boolean; recommendation: string; unitNet: number };
type CommonRank = { value:string; counts:{year:number;month:number;current:number}; agreement:number; total:number; score:number };
type ParityRow = { label:string; weight:number; games:{odd:number;even:number;special:number;total:number}[] };

const numberValue = (value: unknown): string | null => {
  if (value === null || value === undefined || value === '') return null;
  const match = String(value).trim().match(/^0*(\d{1,3})(?:\.0+)?$/);
  if (!match) return null;
  const n = Number(match[1]);
  return n >= 1 && n <= 100 ? (n === 100 ? '100' : String(n).padStart(2, '0')) : null;
};
const reversed = (value: string) => numberValue(value.split('').reverse().join('')) || value;
const isDoubleOrZeroNumber = (value:string) => value === '100' || value.startsWith('0') || value.endsWith('0') || (value.length===2 && value[0]===value[1]);
const keyOf = (y: number, m: number, d: number) => `${y}-${m}-${d}`;
const prettyDate = (y: number, m: number, d: number) => `${d} ${MONTHS[m - 1]} ${y}`;

function dateParts(value: unknown, fallbackYear: number, fallbackMonth?: number): [number, number, number] | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return [value.getFullYear(), value.getMonth() + 1, value.getDate()];
  if (typeof value === 'number' && value > 20000) {
    const d = XLSX.SSF.parse_date_code(value);
    return d ? [d.y, d.m, d.d] : null;
  }
  const text = String(value ?? '').trim();
  if (!text) return null;
  if (fallbackMonth && /^\d{1,2}$/.test(text)) {
    const day = Number(text); return day >= 1 && day <= 31 ? [fallbackYear, fallbackMonth, day] : null;
  }
  const named = text.toLowerCase().match(/(\d{1,2})[\s\-/]+([a-z]{3,9})[\s\-/]+(\d{2,4})/);
  if (named && MONTH_ALIASES[named[2]] ) return [Number(named[3]) < 100 ? 2000 + Number(named[3]) : Number(named[3]), MONTH_ALIASES[named[2]], Number(named[1])];
  const iso = text.match(/^(\d{4})[\-/](\d{1,2})[\-/](\d{1,2})/);
  if (iso) return [Number(iso[1]), Number(iso[2]), Number(iso[3])];
  const dmy = text.match(/^(\d{1,2})[\-/](\d{1,2})[\-/](\d{2,4})/);
  if (dmy) return [Number(dmy[3]) < 100 ? 2000 + Number(dmy[3]) : Number(dmy[3]), Number(dmy[2]), Number(dmy[1])];
  return null;
}

function parseWorkbook(buffer: ArrayBuffer, expectedYear: number): DayRecord[] {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const found = new Map<string, DayRecord>();
  for (const sheetName of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
    const sheetMonth = Object.entries(MONTH_ALIASES).find(([name]) => sheetName.toLowerCase().includes(name))?.[1];

    // Shared-date horizontal format: Date | Jan (4 games) | Feb (4 games) ... Dec.
    // Month names may be merged across their four columns, so only the first cell
    // of each month block is expected to contain the month label.
    let handledWideMonths = false;
    for (let monthRow = 0; monthRow < Math.min(rows.length, 8); monthRow++) {
      const monthBlocks = rows[monthRow].map((value, col) => {
        const label = String(value ?? '').trim().toLowerCase();
        return MONTH_ALIASES[label] ? { month: MONTH_ALIASES[label], col } : null;
      }).filter((block): block is { month: number; col: number } => Boolean(block));
      if (monthBlocks.length < 2) continue;

      const headerRow = rows.slice(monthRow + 1, monthRow + 4).findIndex(row =>
        row.some(value => /^(date|day)$/i.test(String(value ?? '').trim()))
      );
      if (headerRow < 0) continue;
      const actualHeaderRow = monthRow + 1 + headerRow;
      const dateCol = rows[actualHeaderRow].findIndex(value => /^(date|day)$/i.test(String(value ?? '').trim()));

      for (let r = actualHeaderRow + 1; r < rows.length; r++) {
        for (const block of monthBlocks) {
          const parts = dateParts(rows[r][dateCol], expectedYear, block.month);
          const maxDay = new Date(expectedYear, block.month, 0).getDate();
          if (!parts || parts[2] > maxDay) continue;
          const games = rows[r].slice(block.col, block.col + 4).map(numberValue).filter((v): v is string => Boolean(v));
          if (games.length !== 4) continue;
          found.set(keyOf(expectedYear, block.month, parts[2]), { year: expectedYear, month: block.month, day: parts[2], games });
        }
      }
      handledWideMonths = true;
      break;
    }
    if (handledWideMonths) continue;

    let headerIndex = -1, dateCol = 0, gameCols: number[] = [];
    for (let r = 0; r < Math.min(rows.length, 12); r++) {
      const labels = rows[r].map(v => String(v ?? '').trim().toLowerCase());
      const dc = labels.findIndex(v => /^(date|day|draw date|result date)$/.test(v));
      const gc = labels.map((v, i) => (/^(game|g|result|draw|slot)[\s_-]*[1-4]$/.test(v) ? i : -1)).filter(i => i >= 0);
      if (dc >= 0 || gc.length >= 2) { headerIndex = r; dateCol = dc >= 0 ? dc : 0; gameCols = gc; break; }
    }
    for (let r = headerIndex + 1; r < rows.length; r++) {
      const row = rows[r];
      const parts = dateParts(row[dateCol], expectedYear, sheetMonth);
      if (!parts || parts[0] !== expectedYear || parts[1] < 1 || parts[1] > 12 || parts[2] < 1 || parts[2] > 31) continue;
      const candidateCols = gameCols.length ? gameCols : row.map((_, i) => i).filter(i => i !== dateCol);
      const games = candidateCols.map(i => numberValue(row[i])).filter((v): v is string => Boolean(v)).slice(0, 4);
      if (games.length !== 4) continue;
      found.set(keyOf(...parts), { year: parts[0], month: parts[1], day: parts[2], games });
    }
  }
  return [...found.values()].sort((a,b) => a.month - b.month || a.day - b.day);
}

function buildAnalysis(data25: DayRecord[], data26: DayRecord[]): AnalysisRow[] {
  const index = new Map([...data25, ...data26].map(r => [keyOf(r.year,r.month,r.day), r]));
  return data26.map(current => {
    const comparisons: { record?: DayRecord; source: Match['source'] }[] = [
      { record: index.get(keyOf(2025,current.month,current.day)), source: 'Previous year' },
      { record: current.month > 1 ? index.get(keyOf(2026,current.month - 1,current.day)) : undefined, source: 'Previous month' },
    ];
    const matches: Match[] = [];
    current.games.forEach((value, ci) => comparisons.forEach(({record,source}) => record?.games.forEach((other, oi) => {
      const kind = value === other ? 'Exact' : (reversed(value) === other ? 'Reverse' : null);
      if (kind) matches.push({ current:value, currentGame:ci+1, compared:other, comparedGame:oi+1, kind, source, date:prettyDate(record.year,record.month,record.day) });
    })));
    return { ...current, matches };
  });
}

function buildForecasts(data24: DayRecord[], data25: DayRecord[], data26: DayRecord[]): ForecastDay[] {
  if (!data24.length || !data25.length) return [];
  const index = new Map([...data24, ...data25, ...data26].map(r => [keyOf(r.year,r.month,r.day), r]));
  const forecasts: ForecastDay[] = [];
  for (let month=1; month<=12; month++) {
    const maxDay = new Date(2026,month,0).getDate();
    for (let day=1; day<=maxDay; day++) {
      const positionMaps = Array.from({length:4},()=>new Map<string,{score:number;evidence:Set<string>}>());
      const add = (position:number,value:string,points:number,reason:string) => {
        const map=positionMaps[position], current=map.get(value)||{score:0,evidence:new Set<string>()}; current.score+=points; current.evidence.add(reason); map.set(value,current);
      };
      const addRecord = (record:DayRecord|undefined, exactWeight:number, reverseWeight:number, label:string) => record?.games.forEach((value,p)=>{
        add(p,value,exactWeight,`${label}: ${GAME_NAMES[p]} ${value}`);
        const rev=reversed(value); if(rev!==value) add(p,rev,reverseWeight,`${label} reverse: ${value}→${rev}`);
      });
      const same25=index.get(keyOf(2025,month,day)), same24=index.get(keyOf(2024,month,day));
      addRecord(same25,5,3,'Same date 2025');
      addRecord(same24,3.5,2,'Same date 2024');
      addRecord(month>1?index.get(keyOf(2026,month-1,day)):undefined,4,2.5,'Previous month 2026');
      for(let back=2;back<=4;back++) if(month-back>=1) addRecord(index.get(keyOf(2026,month-back,day)),1.25,.65,`${MONTHS[month-back-1]} 2026, same day`);
      if(same24&&same25) same25.games.forEach((value,p)=>{
        if(value===same24.games[p]) add(p,value,3,'Repeated in the same position in 2024 and 2025');
        else if(value===reversed(same24.games[p])) add(p,value,2,'2024→2025 reverse-position pattern');
      });
      const byGame=positionMaps.map(map=>[...map.entries()].map(([value,item])=>({value,score:Number(item.score.toFixed(2)),evidence:[...item.evidence]})).sort((a,b)=>b.score-a.score||a.value.localeCompare(b.value)).slice(0,5));
      const overallMap=new Map<string,{score:number;evidence:Set<string>}>();
      byGame.forEach((list,p)=>list.forEach(c=>{const current=overallMap.get(c.value)||{score:0,evidence:new Set<string>()};current.score+=c.score; c.evidence.forEach(e=>current.evidence.add(`${GAME_NAMES[p]} · ${e}`));overallMap.set(c.value,current)}));
      const overall=[...overallMap.entries()].map(([value,item])=>({value,score:Number(item.score.toFixed(2)),evidence:[...item.evidence]})).sort((a,b)=>b.score-a.score||a.value.localeCompare(b.value)).slice(0,10);
      forecasts.push({month,day,overall,byGame,actual:index.get(keyOf(2026,month,day))?.games});
    }
  }
  return forecasts;
}

function signalScore(forecast: ForecastDay) {
  const top=forecast.overall[0], runner=forecast.overall[1];
  if(!top) return 0;
  const sources=new Set(top.evidence.map(e=>e.includes('2024')?'2024 same date':e.includes('Previous month')?'Previous month':e.includes('2026, same day')?'Earlier 2026 months':e.includes('Repeated')||e.includes('pattern')?'Multi-year pattern':'2025 same date'));
  const separation=Math.max(0,top.score-(runner?.score||0));
  return Math.min(100,Math.round(Math.min(60,top.score*4)+Math.min(20,sources.size*5)+Math.min(12,top.evidence.length*3)+Math.min(8,separation*2)));
}
function signalLabel(score:number) { return score>=78?'High signal':score>=58?'Watch':'Low signal'; }
function strategyRecommendation(score:number) { return score>=58&&score<70?'Backtest-qualified: top 1 only':'Skip under conservative rule'; }

function PeriodControls({month,cutoff,setMonth,setCutoff,game,setGame,dateMode='before'}:{month:number;cutoff:number;setMonth:(value:number)=>void;setCutoff:(value:number)=>void;game?:string;setGame?:(value:string)=>void;dateMode?:'before'|'day'}) {
  return <div className="pattern-controls"><label className="control"><CalendarDays size={13}/><select value={month} onChange={e=>setMonth(Number(e.target.value))}>{MONTHS.map((name,i)=><option key={name} value={i+1}>{name}</option>)}</select><ChevronDown size={12}/></label><label className="control"><select value={cutoff} onChange={e=>setCutoff(Number(e.target.value))}>{Array.from({length:new Date(2026,month,0).getDate()},(_,i)=>i+1).map(day=><option key={day} value={day}>{dateMode==='day'?`Day ${day}`:`Before day ${day}`}</option>)}</select><ChevronDown size={12}/></label>{game&&setGame&&<label className="control"><select value={game} onChange={e=>setGame(e.target.value)}><option>All games</option>{GAME_NAMES.map(name=><option key={name}>{name}</option>)}</select><ChevronDown size={12}/></label>}</div>;
}

function CommonNumberView({month,cutoff,setMonth,setCutoff,game,setGame,rank}:{month:number;cutoff:number;setMonth:(value:number)=>void;setCutoff:(value:number)=>void;game:string;setGame:(value:string)=>void;rank:CommonRank[]}) {
  return <section className="pattern-view"><div className="pattern-view-head"><div><p>Common number weights</p><h3>Exact-value ranking before {cutoff} {MONTHS[month-1]}</h3><span>Same-month previous year ×2, previous month ×3, earlier current-month dates ×4, plus 5 points for every period containing the value.</span></div><PeriodControls month={month} cutoff={cutoff} setMonth={setMonth} setCutoff={setCutoff} game={game} setGame={setGame}/></div><div className="pattern-warning"><Info size={15}/><span>Weighted score measures historical agreement and frequency. It is not a probability or a guaranteed next result.</span></div><div className="common-rank-grid">{rank.slice(0,12).map((item,i)=><article key={item.value}><span>#{i+1}</span><strong>{item.value}</strong><div><b>{item.score} pts</b><small>{item.agreement}/3 periods · {item.total} occurrences</small></div><dl><div><dt>Previous year</dt><dd>{item.counts.year}</dd></div><div><dt>Previous month</dt><dd>{item.counts.month}</dd></div><div><dt>Current month</dt><dd>{item.counts.current}</dd></div></dl></article>)}</div>{rank.length===0&&<div className="pattern-empty">No complete records are available before this cutoff.</div>}<footer>Use the rank to compare patterns consistently across any month. Existing exact, reverse, digit, forecast, and strategy calculations remain unchanged.</footer></section>;
}

function SingleDigitView({month,cutoff,setMonth,setCutoff,comparisons,digitCounts,periodLabels}:{month:number;cutoff:number;setMonth:(value:number)=>void;setCutoff:(value:number)=>void;comparisons:{game:string;current:string;year:string;previousMonth:string;shared:string[];currentRepeats:string[]}[];digitCounts:number[][];periodLabels:string[]}) {
  const digits=(value:string,shared:string[])=><span className="split-digits">{value?value.split('').map((digit,i)=><b key={i} className={shared.includes(digit)?'shared-digit':''}>{digit}</b>):<b>—</b>}</span>;
  return <section className="pattern-view"><div className="pattern-view-head"><div><p>Single-digit patterns</p><h3>Same-date digit comparison</h3><span>Finds digits common to the previous year and previous month for {cutoff} {MONTHS[month-1]}, separately for FB, GB, GL and DS.</span></div><PeriodControls month={month} cutoff={cutoff} setMonth={setMonth} setCutoff={setCutoff} dateMode="day"/></div><div className="pattern-warning"><Info size={15}/><span>Every value is split into individual digits. A digit is common when it appears in both the previous-year value and the previous-month value; the current result is shown only to verify whether that historical digit repeated.</span></div><div className="same-date-digit-grid">{comparisons.map(item=><article key={item.game}><header><strong>{item.game}</strong><span>Historical comparison</span></header><dl><div><dt>Current 2026</dt><dd>{digits(item.current,item.currentRepeats)}</dd></div><div><dt>Previous year</dt><dd>{digits(item.year,item.shared)}</dd></div><div><dt>Previous month</dt><dd>{digits(item.previousMonth,item.shared)}</dd></div></dl><footer><span>Historical common digits</span><strong>{item.shared.length?item.shared.join(', '):'No common digit'}</strong>{item.current&&<small>Repeated in current: {item.currentRepeats.length?item.currentRepeats.join(', '):'None'}</small>}</footer></article>)}</div><div className="digit-compare-table"><table><thead><tr><th>Digit</th><th>{periodLabels[0]}</th><th>{periodLabels[1]}</th><th>{periodLabels[2]}</th><th>Historical match</th></tr></thead><tbody>{Array.from({length:10},(_,digit)=>{const y=digitCounts[0]?.[digit]||0,m=digitCounts[1]?.[digit]||0,c=digitCounts[2]?.[digit]||0,matched=y>0&&m>0,repeated=matched&&c>0;return <tr key={digit}><td><b>{digit}</b></td><td>{y}</td><td>{m}</td><td>{c}</td><td><span className={matched?'digit-match-all':'digit-match-one'}>{matched?(repeated?'Common · also current':'Common'):'No historical match'}</span></td></tr>})}</tbody></table></div><footer>Shared digits are historical pattern evidence, not guaranteed future numbers.</footer></section>;
}

function OddEvenView({month,cutoff,setMonth,setCutoff,rows,finalLeans}:{month:number;cutoff:number;setMonth:(value:number)=>void;setCutoff:(value:number)=>void;rows:{label:string;games:{value:string;lean:string}[]}[];finalLeans:{game:string;result:string}[]}) {
  return <section className="pattern-view"><div className="pattern-view-head"><div><p>Odd / even patterns</p><h3>Direct number classification</h3><span>Classifies each FB, GB, GL and DS value on the selected date.</span></div><PeriodControls month={month} cutoff={cutoff} setMonth={setMonth} setCutoff={setCutoff} dateMode="day"/></div><div className="special-rule"><strong>Classification rule</strong><span>01–09, 10, 11, 20, 22, 30, 33, 40, 44, 50, 55, 60, 66, 70, 77, 80, 88, 90, 99 and 100 are Double Lean. Every other value is classified directly as Odd Lean or Even Lean.</span></div><div className="upcoming-lean final-lean-panel"><header><div><strong>Final upcoming result</strong><span>Uses same-date history, prior 2026 months, the previous month, and earlier current-month results. The algorithm resolves every tie to one result.</span></div><small>Pattern-based estimate · not a guarantee</small></header><div className="upcoming-lean-grid final-lean-grid">{finalLeans.map(item=><article key={item.game}><b>{item.game}</b><strong className={`final-lean-result final-${item.result.toLowerCase()}`}>{item.result}</strong></article>)}</div></div><div className="parity-table lean-table"><table><thead><tr><th>Comparison date</th>{GAME_NAMES.map(name=><th key={name}>{name}</th>)}</tr></thead><tbody>{rows.map(row=><tr key={row.label}><td><b>{row.label}</b></td>{row.games.map((game,p)=><td key={GAME_NAMES[p]}><strong className="lean-value">{game.value||'—'}</strong><span className={`lean-badge ${game.lean==='Odd Lean'?'lean-odd':game.lean==='Even Lean'?'lean-even':game.lean==='Double Lean'?'lean-double':'lean-empty'}`}>{game.lean}</span></td>)}</tr>)}</tbody></table></div><footer>The final result is selected from historical pattern evidence only; it does not guarantee a future draw.</footer></section>;
}

function UploadCard({year, file, onFile}:{year:number; file:FileState; onFile:(f:File)=>void}) {
  const input = useRef<HTMLInputElement>(null);
  return <div onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault(); const f=e.dataTransfer.files[0]; if(f) onFile(f)}} className={`relative flex min-h-40 flex-col items-center justify-center rounded-2xl border p-5 text-center transition ${file ? 'border-[#8db5a2] bg-[#f7fbf8]' : 'border-dashed border-[#b9c8bf] bg-white hover:border-[#1e6d53]'}`}>
    {file ? <><span className="mb-3 grid h-10 w-10 place-items-center rounded-full bg-[#dceee4] text-[#176044]"><Check size={20}/></span><strong className="text-sm">{year} ready</strong><span className="mt-1 max-w-full truncate text-xs text-[#6b7b72]">{file.name} · {file.records.length} dates</span><button onClick={()=>input.current?.click()} className="mt-3 text-xs font-bold text-[#176044] underline underline-offset-4">Replace file</button></> : <><span className="mb-3 grid h-10 w-10 place-items-center rounded-xl bg-[#edf4ef] text-[#1e6d53]"><FileSpreadsheet size={21}/></span><strong className="text-sm">Upload {year} workbook</strong><span className="mt-1 text-xs text-[#718078]">Drop .xlsx, .xls or .csv here</span><button onClick={()=>input.current?.click()} className="mt-3 inline-flex items-center gap-2 rounded-lg bg-[#183e32] px-4 py-2 text-xs font-bold text-white"><Upload size={14}/> Choose file</button></>}
    <input ref={input} onChange={e=>e.target.files?.[0]&&onFile(e.target.files[0])} className="sr-only" type="file" accept=".xlsx,.xls,.csv" />
  </div>;
}

export default function Home() {
  const [file24,setFile24] = useState<FileState>(null), [file25,setFile25] = useState<FileState>(null), [file26,setFile26] = useState<FileState>(null);
  const [error,setError] = useState(''), [month,setMonth] = useState(1), [focus,setFocus] = useState('Monthly matches'), [commonDay,setCommonDay] = useState<number|null>(null), [forecastMonth,setForecastMonth] = useState(9), [forecastDate,setForecastDate] = useState('9-9'), [source,setSource] = useState('All sources'), [kind,setKind] = useState('All matches'), [query,setQuery] = useState(''), [digitMonth,setDigitMonth] = useState(9), [digitCutoff,setDigitCutoff] = useState(19), [digitSource,setDigitSource] = useState('All sources'), [commonGame,setCommonGame] = useState('All games');
  const analysis = useMemo(()=>buildAnalysis(file25?.records||[],file26?.records||[]),[file25,file26]);
  const forecasts = useMemo(()=>buildForecasts(file24?.records||[],file25?.records||[],file26?.records||[]),[file24,file25,file26]);
  const ready = Boolean(file25 && file26);
  const dayPatterns = Array.from({length:31},(_,i)=>i+1).map(day=>{
    const rows=analysis.filter(r=>r.day===day&&r.matches.length>0);
    return {day,months:new Set(rows.map(r=>r.month)).size,exactMonths:new Set(rows.filter(r=>r.matches.some(m=>m.kind==='Exact')).map(r=>r.month)).size,reverseMonths:new Set(rows.filter(r=>r.matches.some(m=>m.kind==='Reverse')).map(r=>r.month)).size,matches:rows.reduce((n,r)=>n+r.matches.length,0)};
  }).filter(p=>p.months>1).sort((a,b)=>b.months-a.months||b.matches-a.matches||a.day-b.day);
  const selectedCommonDay = commonDay ?? dayPatterns[0]?.day ?? 1;
  const focusedRows = (focus === 'Monthly matches' ? analysis.filter(r=>r.month===month) : focus === 'Common day patterns' ? analysis.filter(r=>r.day===selectedCommonDay) : analysis)
    .filter(r => focus !== 'Exact + reverse dates' || (r.matches.some(m=>m.kind==='Exact') && r.matches.some(m=>m.kind==='Reverse')))
    .filter(r => focus !== 'Same-position exact' || r.matches.some(m=>m.kind==='Exact' && m.currentGame===m.comparedGame));
  const filteredRows = focusedRows.filter(r => {
    const ms = r.matches.filter(m => (focus !== 'Same-position exact' || (m.kind==='Exact' && m.currentGame===m.comparedGame)) && (source==='All sources'||m.source===source) && (kind==='All matches'||m.kind===kind));
    const q = query.trim().toLowerCase(); return ms.length > 0 && (!q || String(r.day).includes(q) || r.games.some(g=>g.includes(q)) || ms.some(m=>m.current.includes(q)||m.compared.includes(q)));
  }).map(r=>({...r,matches:r.matches.filter(m => (focus !== 'Same-position exact' || (m.kind==='Exact' && m.currentGame===m.comparedGame)) && (source==='All sources'||m.source===source)&&(kind==='All matches'||m.kind===kind))}))
    .sort((a,b)=>focus==='Exact + reverse dates' ? b.matches.length-a.matches.length : a.month-b.month||a.day-b.day);
  const totals = MONTHS.map((_,i)=>analysis.filter(r=>r.month===i+1).reduce((n,r)=>n+r.matches.length,0));
  const exact = analysis.reduce((n,r)=>n+r.matches.filter(m=>m.kind==='Exact').length,0), reverseCount = analysis.reduce((n,r)=>n+r.matches.filter(m=>m.kind==='Reverse').length,0);
  const mixedDates = analysis.filter(r=>r.matches.some(m=>m.kind==='Exact')&&r.matches.some(m=>m.kind==='Reverse')).length;
  const positionExact = analysis.reduce((n,r)=>n+r.matches.filter(m=>m.kind==='Exact'&&m.currentGame===m.comparedGame).length,0);
  const activePattern = dayPatterns.find(p=>p.day===selectedCommonDay);
  const viewTitle = focus==='Monthly matches' ? `${MONTHS[month-1]} 2026` : focus==='Common day patterns' ? `Day ${String(selectedCommonDay).padStart(2,'0')} pattern` : focus;
  const viewSubtitle = focus==='Monthly matches' ? `Compared with ${MONTHS[month-1]} 2025${month>1?` and ${MONTHS[month-2]} 2026`:''}` : focus==='Exact + reverse dates' ? `${filteredRows.length} strongest dates across January–December 2026, ranked by matches` : focus==='Same-position exact' ? `${filteredRows.length} dates with an exact value in the same game position` : `Matches on day ${selectedCommonDay} appear in ${activePattern?.months||0} of 12 months · ${activePattern?.exactMonths||0} exact-months · ${activePattern?.reverseMonths||0} reverse-months`;
  const monthForecasts=forecasts.filter(f=>f.month===forecastMonth);
  const activeForecast=monthForecasts.find(f=>`${f.month}-${f.day}`===forecastDate)||monthForecasts[0];
  const activeSignal=activeForecast?signalScore(activeForecast):0;
  const activeSignalLabel=signalLabel(activeSignal);
  const bestUpcomingDays=monthForecasts.filter(f=>!f.actual).map(f=>({forecast:f,score:signalScore(f)})).sort((a,b)=>b.score-a.score||a.forecast.day-b.forecast.day).slice(0,2);
  const activeForecastMatches=activeForecast?analysis.find(r=>r.month===activeForecast.month&&r.day===activeForecast.day)?.matches||[]:[];
  const candidateResult=(candidate:string,actual?:string[])=>!actual?'Pending':actual.includes(candidate)?'Exact':actual.includes(reversed(candidate))?'Reverse':'No match';
  const strength=(candidate:Candidate,topScore:number)=>candidate.score>=Math.max(8,topScore*.75)&&candidate.evidence.length>=3?'Strong':candidate.score>=Math.max(5,topScore*.45)?'Moderate':'Exploratory';
  const strategyRows:StrategyRow[]=forecasts.filter(f=>f.actual&&f.overall[0]).map(f=>{const score=signalScore(f),candidate=f.overall[0].value,hit=Boolean(f.actual?.includes(candidate)||f.actual?.includes(reversed(candidate)));return {forecast:f,score,candidate,hit,recommendation:strategyRecommendation(score),unitNet:hit?94:-1};});
  const qualifiedRows=strategyRows.filter(r=>r.score>=58&&r.score<70), qualifiedWins=qualifiedRows.filter(r=>r.hit).length, qualifiedNet=qualifiedRows.reduce((n,r)=>n+r.unitNet,0);
  const longestQualifiedDryRun=qualifiedRows.reduce((state,row)=>{const run=row.hit?0:state.run+1;return {run,max:Math.max(state.max,run)}},{run:0,max:0}).max;
  const highScoreRows=strategyRows.filter(r=>r.score>=70), highScoreWins=highScoreRows.filter(r=>r.hit).length;
  const elapsedDays=strategyRows.length?Math.max(1,Math.round((Date.UTC(2026,strategyRows.at(-1)!.forecast.month-1,strategyRows.at(-1)!.forecast.day)-Date.UTC(2026,0,1))/86400000)+1):0;
  const targetUnit=qualifiedNet>0&&elapsedDays?Math.ceil(100000/(qualifiedNet/(elapsedDays/30.44))):null;
  const digitSourceRows=[...(file25?.records||[]).filter(r=>r.month===digitMonth&&r.day<digitCutoff).map(record=>({record,source:'Previous year'})),...(file26?.records||[]).filter(r=>r.month===digitMonth-1&&r.day<digitCutoff).map(record=>({record,source:'Previous month'})),...(file26?.records||[]).filter(r=>r.month===digitMonth&&r.day<digitCutoff).map(record=>({record,source:'Earlier current month'}))].filter(r=>digitSource==='All sources'||r.source===digitSource);
  const digitCounts=Array.from({length:4},()=>Array(10).fill(0) as number[]), digitTotals=Array(4).fill(0) as number[];
  digitSourceRows.forEach(({record})=>record.games.forEach((value,p)=>value.split('').forEach(char=>{const digit=Number(char);if(Number.isInteger(digit)){digitCounts[p][digit]++;digitTotals[p]++;}})));
  const digitLeaders=digitCounts.map((counts,p)=>counts.map((count,digit)=>({digit,count,share:digitTotals[p]?count/digitTotals[p]:0})).sort((a,b)=>b.count-a.count||a.digit-b.digit).slice(0,3));
  const periodGroups=[{key:'year',label:`${MONTHS[digitMonth-1]} 2025`,weight:2,records:(file25?.records||[]).filter(r=>r.month===digitMonth&&r.day<digitCutoff)},{key:'month',label:digitMonth>1?`${MONTHS[digitMonth-2]} 2026`:'No previous month',weight:3,records:(file26?.records||[]).filter(r=>r.month===digitMonth-1&&r.day<digitCutoff)},{key:'current',label:`Earlier ${MONTHS[digitMonth-1]} 2026`,weight:4,records:(file26?.records||[]).filter(r=>r.month===digitMonth&&r.day<digitCutoff)}];
  const selectedDigitRecords=[(file25?.records||[]).find(r=>r.month===digitMonth&&r.day===digitCutoff),(file26?.records||[]).find(r=>r.month===digitMonth-1&&r.day===digitCutoff),(file26?.records||[]).find(r=>r.month===digitMonth&&r.day===digitCutoff)];
  const comparisonDigitCounts=selectedDigitRecords.map(record=>{const counts=Array(10).fill(0) as number[];record?.games.forEach(value=>value.split('').forEach(char=>counts[Number(char)]++));return counts;});
  const digitComparisons=GAME_NAMES.map((game,p)=>{const year=selectedDigitRecords[0]?.games[p]||'',previousMonth=selectedDigitRecords[1]?.games[p]||'',current=selectedDigitRecords[2]?.games[p]||'',shared=year&&previousMonth?[...new Set(year.split('').filter(digit=>previousMonth.includes(digit)))]:[],currentRepeats=current?shared.filter(digit=>current.includes(digit)):[];return {game,current,year,previousMonth,shared,currentRepeats};});
  const commonNumberMap=new Map<string,{year:number;month:number;current:number}>();
  periodGroups.forEach(group=>group.records.forEach(record=>record.games.forEach((value,p)=>{if(commonGame!=='All games'&&commonGame!==GAME_NAMES[p])return;const item=commonNumberMap.get(value)||{year:0,month:0,current:0};item[group.key as 'year'|'month'|'current']++;commonNumberMap.set(value,item);})));
  const commonNumberRank=[...commonNumberMap.entries()].map(([value,counts])=>{const agreement=[counts.year,counts.month,counts.current].filter(Boolean).length;return {value,counts,agreement,total:counts.year+counts.month+counts.current,score:counts.year*2+counts.month*3+counts.current*4+agreement*5};}).sort((a,b)=>b.score-a.score||b.agreement-a.agreement||b.total-a.total||a.value.localeCompare(b.value));
  const classifyLean=(value:string)=>!value?'No data':isDoubleOrZeroNumber(value)?'Double Lean':Number(value)%2===1?'Odd Lean':'Even Lean';
  const directLeanRows=[{label:`${digitCutoff} ${MONTHS[digitMonth-1]} 2025`,record:selectedDigitRecords[0]},{label:digitMonth>1?`${digitCutoff} ${MONTHS[digitMonth-2]} 2026`:'No previous month',record:selectedDigitRecords[1]},{label:`${digitCutoff} ${MONTHS[digitMonth-1]} 2026`,record:selectedDigitRecords[2]}].map(item=>({label:item.label,games:GAME_NAMES.map((_,p)=>{const value=item.record?.games[p]||'';return {value,lean:classifyLean(value)};})}));
  const finalLeans=GAME_NAMES.map((game,p)=>{const labels=['Odd Lean','Even Lean','Double Lean'],scores=new Map(labels.map(label=>[label,0])),add=(value:string|undefined,weight:number)=>{if(!value)return;const label=classifyLean(value);scores.set(label,(scores.get(label)||0)+weight);};add(selectedDigitRecords[0]?.games[p],8);add(selectedDigitRecords[1]?.games[p],10);for(let m=1;m<digitMonth;m++){const record=(file26?.records||[]).find(r=>r.month===m&&r.day===digitCutoff);add(record?.games[p],2+4*m/Math.max(1,digitMonth-1));}const groups=[{records:(file25?.records||[]).filter(r=>r.month===digitMonth&&r.day<digitCutoff),weight:4},{records:(file26?.records||[]).filter(r=>r.month===digitMonth-1&&r.day<=digitCutoff),weight:6},{records:(file26?.records||[]).filter(r=>r.month===digitMonth&&r.day<digitCutoff),weight:7}];groups.forEach(group=>{const values=group.records.map(r=>r.games[p]).filter(Boolean);values.forEach(value=>add(value,group.weight/Math.max(1,values.length)));});const history=[...(file25?.records||[]),...(file26?.records||[]).filter(r=>r.month<digitMonth||(r.month===digitMonth&&r.day<digitCutoff))],frequency=new Map(labels.map(label=>[label,0]));history.forEach(record=>{const label=classifyLean(record.games[p]);frequency.set(label,(frequency.get(label)||0)+1);});const recent=classifyLean(selectedDigitRecords[1]?.games[p]||selectedDigitRecords[0]?.games[p]||'');const winner=[...labels].sort((a,b)=>(scores.get(b)||0)-(scores.get(a)||0)||(a===recent?-1:0)-(b===recent?-1:0)||(frequency.get(b)||0)-(frequency.get(a)||0)||labels.indexOf(a)-labels.indexOf(b))[0];return {game,result:winner.replace(' Lean','')};});

  async function load(file:File, year:number) {
    setError('');
    try { const records=parseWorkbook(await file.arrayBuffer(),year); if(!records.length) throw new Error(`No valid ${year} rows found. Use a Date/Day column followed by four game columns.`); (year===2024?setFile24:year===2025?setFile25:setFile26)({name:file.name,records}); }
    catch(e){ setError(e instanceof Error?e.message:'Could not read that workbook.'); }
  }
  function demo() {
    const c:DayRecord[] = [], a:DayRecord[] = [], b:DayRecord[] = [];
    for(let m=1;m<=12;m++) for(let d=1;d<=Math.min(m===2?28:30,31);d++) {
      c.push({year:2024,month:m,day:d,games:[d+2,d+13,d+31,d+49].map(n=>numberValue((n%99)+1)!) });
      a.push({year:2025,month:m,day:d,games:[d,d+11,d+29,d+47].map(n=>numberValue((n%99)+1)!) });
      if(m<9||d<=8) b.push({year:2026,month:m,day:d,games:[d%2?String(((d+10)%99)+1).padStart(2,'0'):reversed(a[a.length-1].games[0]),a[a.length-1].games[2],String(((m*7+d)%99)+1).padStart(2,'0'),String(((m*11+d*3)%99)+1).padStart(2,'0')]});
    }
    setFile24({name:'sample-2024.xlsx',records:c}); setFile25({name:'sample-2025.xlsx',records:a}); setFile26({name:'sample-2026.xlsx',records:b}); setMonth(8); setError('');
  }
  function exportCsv() {
    const lines=[['Current date','Current game','Current value','Match type','Comparison period','Comparison date','Comparison game','Comparison value']];
    analysis.forEach(r=>r.matches.forEach(m=>lines.push([prettyDate(r.year,r.month,r.day),GAME_NAMES[m.currentGame-1],m.current,m.kind,m.source,m.date,GAME_NAMES[m.comparedGame-1],m.compared])));
    const csv=lines.map(row=>row.map(v=>`"${String(v).replaceAll('"','""')}"`).join(',')).join('\n'); const url=URL.createObjectURL(new Blob([csv],{type:'text/csv'})); const a=document.createElement('a');a.href=url;a.download='excel-match-analysis-2026.csv';a.click();URL.revokeObjectURL(url);
  }
  function exportForecastCsv() {
    const lines=[['Forecast date','Status','Actual FB','Actual GB','Actual GL','Actual DS','Scope','Candidate','Candidate result','Pattern score','Strength','Strategy / recommendation','Supporting evidence']];
    forecasts.forEach(f=>{
      const top=f.overall[0]?.score||1;
      const base=[prettyDate(2026,f.month,f.day),f.actual?'Historical':'Upcoming',...(f.actual||['','','',''])];
      f.overall.forEach(c=>lines.push([...base,'Overall',c.value,candidateResult(c.value,f.actual),String(c.score),strength(c,top),c===f.overall[0]?strategyRecommendation(signalScore(f)):'Supporting candidate only',c.evidence.join(' | ')]));
      f.byGame.forEach((list,p)=>list.forEach(c=>lines.push([...base,GAME_NAMES[p],c.value,f.actual?(f.actual[p]===c.value?'Exact':f.actual[p]===reversed(c.value)?'Reverse':'No match'):'Pending',String(c.score),strength(c,list[0]?.score||1),'Position evidence only',c.evidence.join(' | ')])));
    });
    const csv=lines.map(row=>row.map(v=>`"${String(v).replaceAll('"','""')}"`).join(',')).join('\n'); const url=URL.createObjectURL(new Blob([csv],{type:'text/csv'})); const a=document.createElement('a');a.href=url;a.download='pattern-candidates-complete-2026.csv';a.click();URL.revokeObjectURL(url);
  }

  return <main className="min-h-screen bg-[#f5f7f4] text-[#17231d]">
    <header className="sticky top-0 z-20 border-b border-[#dce4dd] bg-[#fbfcfa]/95 backdrop-blur">
      <div className="mx-auto flex max-w-[1500px] items-center justify-between px-4 py-3 sm:px-6 lg:px-9"><div className="flex items-center gap-3"><span className="grid h-9 w-9 place-items-center rounded-xl bg-[#163f32] text-white"><BarChart3 size={19}/></span><div><p className="text-[9px] font-bold uppercase tracking-[.18em] text-[#718078]">Number intelligence</p><h1 className="text-base font-bold tracking-tight">Excel Match Analyzer</h1></div></div><div className="flex items-center gap-3"><span className="hidden items-center gap-2 text-xs font-semibold text-[#5b7065] md:flex"><ShieldCheck size={14}/> Local browser analysis</span>{ready&&<button onClick={exportCsv} className="inline-flex items-center gap-2 rounded-lg bg-[#173f32] px-3.5 py-2 text-xs font-bold text-white hover:bg-[#0f3025]"><Download size={14}/> <span className="hidden sm:inline">Export report</span></button>}</div></div>
    </header>
    <div className="mx-auto max-w-[1500px] px-4 py-6 sm:px-6 lg:px-9">
      {!ready ? <section>
        <div className="mb-6 flex flex-col justify-between gap-3 md:flex-row md:items-end"><div><p className="mb-2 flex items-center gap-2 text-[11px] font-bold uppercase tracking-[.16em] text-[#b45a32]"><Sparkles size={13}/> 2026 analysis workspace</p><h2 className="max-w-2xl text-3xl font-bold tracking-[-.04em] sm:text-4xl">Find exact and reversed numbers across years.</h2></div><p className="max-w-md text-sm leading-6 text-[#65736c]">Upload both workbooks. Each 2026 date is checked against the same date in 2025 and the previous month of 2026, across all four games.</p></div>
        <div className="grid gap-4 lg:grid-cols-3"><UploadCard year={2024} file={file24} onFile={f=>load(f,2024)}/><UploadCard year={2025} file={file25} onFile={f=>load(f,2025)}/><UploadCard year={2026} file={file26} onFile={f=>load(f,2026)}/></div>
        {error&&<div className="mt-4 flex items-start gap-3 rounded-xl border border-[#edc8b8] bg-[#fff8f4] p-4 text-sm text-[#8d3d1f]"><Info className="mt-0.5 shrink-0" size={16}/>{error}<button aria-label="Close error" onClick={()=>setError('')} className="ml-auto"><X size={15}/></button></div>}
        <div className="mt-4 flex items-center justify-center gap-3 text-xs text-[#718078]"><span>Want to explore first?</span><button onClick={demo} className="font-bold text-[#176044] underline underline-offset-4">Load sample analysis</button></div>
        <div className="mt-8 grid gap-3 md:grid-cols-3">{[['1','Upload both files','Excel or CSV, with a date and four result columns.'],['2','Automatic comparison','Every game checks all four games on each matching date.'],['3','Review & export','Filter exact or reverse matches and download a complete CSV.']].map(([n,t,d])=><div key={n} className="rounded-xl border border-[#dce4dd] bg-[#fafcfa] p-4"><span className="mb-3 grid h-7 w-7 place-items-center rounded-full bg-[#e4efe8] text-xs font-black text-[#176044]">{n}</span><h3 className="text-sm font-bold">{t}</h3><p className="mt-1 text-xs leading-5 text-[#718078]">{d}</p></div>)}</div>
      </section> : <section>
        <div className="mb-5 flex flex-col justify-between gap-4 xl:flex-row xl:items-center"><div><p className="mb-1 text-[10px] font-bold uppercase tracking-[.18em] text-[#aa5b37]">Analysis complete</p><h2 className="text-2xl font-bold tracking-[-.03em]">2026 match report</h2><p className="mt-1 text-xs text-[#6d7c74]">{file25?.records.length} dates from 2025 · {file26?.records.length} dates from 2026</p></div><div className="grid grid-cols-3 gap-2 sm:min-w-[430px]"><div className="stat"><span>Total matches</span><strong>{exact+reverseCount}</strong></div><div className="stat"><span>Exact</span><strong>{exact}</strong></div><div className="stat"><span>Reversed</span><strong className="text-[#b2552d]">{reverseCount}</strong></div></div></div>
        <div className="grid gap-5 lg:grid-cols-[220px_minmax(0,1fr)]">
          <aside className="h-fit rounded-2xl border border-[#d9e2dc] bg-[#fbfcfa] p-3 lg:sticky lg:top-20">
            <button onClick={()=>setFocus('Forecast lab')} className={`mb-3 flex w-full items-center justify-between rounded-xl border px-3 py-3 text-left ${focus==='Forecast lab'?'border-[#aa5b37] bg-[#aa5b37] text-white':'border-[#edc9b8] bg-[#fff8f4] text-[#7c3f24]'}`}><span><strong className="flex items-center gap-2 text-xs"><Sparkles size={14}/> Forecast Lab</strong><small className={`mt-1 block text-[10px] ${focus==='Forecast lab'?'text-[#ffe0d1]':'text-[#9a6c57]'}`}>Full-year history + future candidates</small></span><b className="rounded-md bg-white/20 px-2 py-1 text-[10px]">NEW</b></button>
            <div className="mb-2 px-2 py-1 text-[10px] font-bold uppercase tracking-[.14em] text-[#738078]">Accuracy filters</div>
            <div className="mb-3 grid gap-1">
              <button onClick={()=>setFocus('Exact + reverse dates')} className={`insight-filter ${focus==='Exact + reverse dates'?'insight-active':''}`}><span><strong>Exact + reverse dates</strong><small>Most accurate across all months</small></span><b>{mixedDates}</b></button>
              <button onClick={()=>setFocus('Same-position exact')} className={`insight-filter ${focus==='Same-position exact'?'insight-active':''}`}><span><strong>Same-position exact</strong><small>FB→FB, GB→GB, GL→GL, DS→DS</small></span><b>{positionExact}</b></button>
              <button onClick={()=>setFocus('Common day patterns')} className={`insight-filter ${focus==='Common day patterns'?'insight-active':''}`}><span><strong>Common day patterns</strong><small>Recurring day numbers across months</small></span><b>{dayPatterns[0]?.months||0}/12</b></button>
              <button onClick={()=>setFocus('Single-digit patterns')} className={`insight-filter ${focus==='Single-digit patterns'?'insight-active':''}`}><span><strong>Single-digit patterns</strong><small>Digit counts by FB, GB, GL and DS</small></span><b>0–9</b></button>
              <button onClick={()=>setFocus('Odd/even patterns')} className={`insight-filter ${focus==='Odd/even patterns'?'insight-active':''}`}><span><strong>Odd / even patterns</strong><small>Parity comparison by game and period</small></span><b>O/E</b></button>
            </div>
            <div className="mb-2 flex items-center justify-between border-t border-[#e1e7e3] px-2 pt-3"><span className="text-[10px] font-bold uppercase tracking-[.14em] text-[#738078]">Monthly summary</span><CalendarDays size={14} className="text-[#718078]"/></div>
            <div className="grid grid-cols-2 gap-1 lg:grid-cols-1">{MONTHS.map((name,i)=><button key={name} onClick={()=>{setMonth(i+1);setFocus('Monthly matches')}} className={`flex items-center justify-between rounded-lg px-2.5 py-2 text-left text-xs font-semibold ${focus==='Monthly matches'&&month===i+1?'bg-[#173f32] text-white':'text-[#53665c] hover:bg-[#edf3ef]'}`}><span>{name}</span><span className={`rounded-md px-1.5 py-0.5 text-[10px] ${focus==='Monthly matches'&&month===i+1?'bg-white/15':'bg-[#e5ece7]'}`}>{totals[i]}</span></button>)}</div>
            <button onClick={()=>{setFile24(null);setFile25(null);setFile26(null)}} className="mt-3 flex w-full items-center justify-center gap-2 border-t border-[#e1e7e3] pt-3 text-xs font-bold text-[#6c7a72]"><RotateCcw size={13}/> Start over</button>
          </aside>
          <div className={`min-w-0 ${focus==='Single-digit patterns'||focus==='Odd/even patterns'?'pattern-special':''}`}>
            {focus==='Single-digit patterns'&&<SingleDigitView month={digitMonth} cutoff={digitCutoff} setMonth={setDigitMonth} setCutoff={setDigitCutoff} comparisons={digitComparisons} digitCounts={comparisonDigitCounts} periodLabels={periodGroups.map(group=>group.label)}/>}
            {focus==='Odd/even patterns'&&<OddEvenView month={digitMonth} cutoff={digitCutoff} setMonth={setDigitMonth} setCutoff={setDigitCutoff} rows={directLeanRows} finalLeans={finalLeans}/>}
            {focus==='Forecast lab'?<div className="forecast-shell">
            {!file24?<div className="rounded-2xl border border-[#d9e2dc] bg-white p-5"><div className="mb-5"><p className="mb-1 text-[10px] font-bold uppercase tracking-[.16em] text-[#aa5b37]">Additional history required</p><h3 className="text-xl font-bold">Add the complete 2024 workbook</h3><p className="mt-2 max-w-2xl text-sm leading-6 text-[#65736c]">The existing 2025–2026 report stays unchanged. Forecast Lab uses 2024 as a second historical reference to score missing September–December 2026 dates.</p></div><UploadCard year={2024} file={file24} onFile={f=>load(f,2024)}/>{error&&<p className="mt-3 text-sm text-[#9b4524]">{error}</p>}</div>:<>
              <div className="mb-3 rounded-2xl border border-[#e4cabb] bg-[#fffaf7] p-4"><div className="flex flex-col justify-between gap-3 md:flex-row md:items-center"><div><p className="mb-1 text-[10px] font-bold uppercase tracking-[.16em] text-[#aa5b37]">Historical pattern forecast</p><h3 className="text-xl font-bold">Complete 2026 candidate history</h3><p className="mt-1 text-xs text-[#765f54]">Review actual results and matches for completed dates, then continue into upcoming dates with the same scoring rules.</p></div><button onClick={exportForecastCsv} className="inline-flex items-center justify-center gap-2 rounded-lg bg-[#8f4728] px-3.5 py-2 text-xs font-bold text-white"><Download size={14}/> Export candidates</button></div><div className="mt-3 flex items-start gap-2 rounded-lg bg-[#fff1e8] p-3 text-xs leading-5 text-[#7c482f]"><Info className="mt-0.5 shrink-0" size={14}/><span>Historical dates support backtesting. Upcoming candidates remain pattern-based and are not guaranteed outcomes.</span></div></div>
              <div className="mb-3 flex flex-col gap-3 rounded-xl border border-[#dce4dd] bg-white p-3 lg:flex-row lg:items-center lg:justify-between"><div className="flex flex-wrap gap-1.5">{MONTHS.map((name,i)=><button key={name} onClick={()=>{setForecastMonth(i+1);setForecastDate('')}} className={`forecast-month ${forecastMonth===i+1?'forecast-month-active':''}`}>{name.slice(0,3)} <span>{forecasts.filter(f=>f.month===i+1&&f.actual).length}/{forecasts.filter(f=>f.month===i+1).length}</span></button>)}</div><label className="control min-w-[180px]"><CalendarDays size={13}/><select value={activeForecast?`${activeForecast.month}-${activeForecast.day}`:''} onChange={e=>setForecastDate(e.target.value)}>{monthForecasts.map(f=><option key={`${f.month}-${f.day}`} value={`${f.month}-${f.day}`}>{prettyDate(2026,f.month,f.day)} · {f.actual?'Actual':'Upcoming'}</option>)}</select><ChevronDown size={12}/></label></div>
              {activeForecast?<><div className="mb-3 flex items-end justify-between"><div><p className="text-[10px] font-bold uppercase tracking-[.14em] text-[#718078]">Selected 2026 date</p><h4 className="text-2xl font-bold tracking-tight">{prettyDate(2026,activeForecast.month,activeForecast.day)}</h4></div><span className={`badge ${activeForecast.actual?'badge-exact':'badge-reverse'}`}>{activeForecast.actual?'Historical result':'Upcoming date'}</span></div>
                <section className={`decision-lens decision-${activeSignalLabel.split(' ')[0].toLowerCase()}`}><div className="decision-head"><div><p>Conservative decision lens</p><h5>{activeSignalLabel}</h5><span>{activeSignalLabel==='High signal'?'Multiple independent patterns agree and the leaders are clearly separated.':activeSignalLabel==='Watch'?'Some useful agreement exists, but this is not one of the clearest opportunities.':'The historical signals do not separate strongly enough. Skipping is the clearest choice.'}</span></div><div className="signal-meter"><strong>{activeSignal}</strong><small>signal score</small></div></div><div className="decision-grid"><div className="shortlist-block"><p>Top two overall</p><div className="shortlist-values">{activeForecast.overall.slice(0,2).map((c,i)=><div key={c.value}><span>#{i+1}</span><strong>{c.value}</strong><small>{c.score} pts · {c.evidence.length} signals</small></div>)}</div></div><div className="shortlist-block"><p>One leader per game</p><div className="game-leaders">{activeForecast.byGame.map((list,p)=><div key={GAME_NAMES[p]}><span>{GAME_NAMES[p]}</span><strong>{list[0]?.value||'—'}</strong><small>{list[0]?.score||0} pts</small></div>)}</div></div><div className="shortlist-block"><p>Best upcoming days in {MONTHS[forecastMonth-1]}</p><div className="opportunity-days">{bestUpcomingDays.length?bestUpcomingDays.map(({forecast,score})=><button key={forecast.day} onClick={()=>setForecastDate(`${forecast.month}-${forecast.day}`)}><strong>{String(forecast.day).padStart(2,'0')}</strong><span>{signalLabel(score)}</span><small>{score}/100</small></button>):<span className="text-xs text-[#718078]">No upcoming dates in this month.</span>}</div></div></div><footer>This score is a relative pattern-strength measure, not a probability or guarantee.</footer></section>
                {activeForecast.actual&&<div className="mb-4 rounded-xl border border-[#cfe0d6] bg-[#f7fbf8] p-4"><div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between"><div><strong className="text-sm">Actual results</strong><p className="mt-1 text-xs text-[#718078]">{activeForecastMatches.length} exact/reverse comparison matches recorded by the existing analyzer</p></div><div className="flex gap-2">{activeForecast.actual.map((value,p)=><span key={GAME_NAMES[p]} className="actual-result" data-game={GAME_NAMES[p]}>{value}</span>)}</div></div>{activeForecastMatches.length>0&&<div className="mt-3 flex flex-wrap gap-1.5 border-t border-[#dce8e0] pt-3">{activeForecastMatches.map((m,i)=><span key={i} className={`badge ${m.kind==='Exact'?'badge-exact':'badge-reverse'}`}>{GAME_NAMES[m.currentGame-1]} {m.current} · {m.kind} · {m.source}</span>)}</div>}</div>}
                <details className="full-analysis"><summary>Show all candidates and supporting evidence</summary><div className="mb-4 grid gap-2 sm:grid-cols-2 xl:grid-cols-5">{activeForecast.overall.slice(0,10).map((c,i)=>{const level=strength(c,activeForecast.overall[0]?.score||1),result=candidateResult(c.value,activeForecast.actual);return <details key={c.value} className={`candidate-card ${result==='Exact'||result==='Reverse'?'candidate-hit':''}`}><summary><span className="candidate-rank">#{i+1}</span><strong>{c.value}</strong><span className={`strength ${result==='Exact'?'strength-strong':result==='Reverse'?'strength-moderate':`strength-${level.toLowerCase()}`}`}>{activeForecast.actual?result:level}</span><small>Score {c.score}</small></summary><div className="evidence-list">{c.evidence.map((e,j)=><p key={j}>{e}</p>)}</div></details>})}</div>
                <div className="grid gap-3 xl:grid-cols-4">{activeForecast.byGame.map((list,p)=><section key={GAME_NAMES[p]} className="position-panel"><div className="position-title"><span>{GAME_NAMES[p]}</span><small>{activeForecast.actual?`Actual ${activeForecast.actual[p]}`:'Position candidates'}</small></div>{list.map((c,i)=>{const result=activeForecast.actual?(activeForecast.actual[p]===c.value?'Exact':activeForecast.actual[p]===reversed(c.value)?'Reverse':''):'Pending';return <details key={c.value} className={`position-candidate ${result==='Exact'||result==='Reverse'?'position-hit':''}`}><summary><b>{c.value}</b><span>#{i+1}</span><em>{result&&result!=='Pending'?`${result} · `:''}{c.score} pts</em></summary><div className="evidence-list">{c.evidence.map((e,j)=><p key={j}>{e}</p>)}</div></details>})}</section>)}</div></details>
                <section className="strategy-audit"><div className="strategy-audit-head"><div><p>Historical strategy audit</p><h5>January 1 to {strategyRows.length?prettyDate(2026,strategyRows.at(-1)!.forecast.month,strategyRows.at(-1)!.forecast.day):'latest uploaded date'}</h5></div><span>₹1 stake · ₹95 gross return</span></div><div className="strategy-stats"><div><strong>{qualifiedRows.length}</strong><span>qualified historical dates</span></div><div><strong>{qualifiedWins}</strong><span>top-1 winning dates</span></div><div><strong>{qualifiedRows.length?`${(qualifiedWins/qualifiedRows.length*100).toFixed(1)}%`:'—'}</strong><span>observed hit rate</span></div><div><strong>{qualifiedNet>=0?'+':''}{qualifiedNet}</strong><span>net units in replay</span></div><div><strong>{longestQualifiedDryRun}</strong><span>longest qualified losing run</span></div></div><div className="strategy-callout"><div><strong>Conservative experimental rule</strong><p>Only scores from 58 to 69 qualify, and only the single top overall candidate is considered. All other dates are marked Skip. The 70+ group produced {highScoreWins} hits from {highScoreRows.length} historical selections, so a higher score must not be treated as a safer bet.</p></div><div><strong>₹1 lakh target check</strong><p>{targetUnit?`Repeating this historical average would require roughly ₹${targetUnit.toLocaleString('en-IN')} on every qualified selection. A ${longestQualifiedDryRun}-selection losing run occurred in the same sample, equal to roughly ₹${(targetUnit*longestQualifiedDryRun).toLocaleString('en-IN')} at that size.`:'There is not enough positive historical evidence to calculate a target scenario.'} This is a risk illustration, not a staking recommendation.</p></div></div><p className="strategy-now"><b>Selected-date recommendation:</b> {activeForecast.actual?'Historical replay':strategyRecommendation(activeSignal)} · Top candidate {activeForecast.overall[0]?.value||'—'} · Signal {activeSignal}/100</p><details className="strategy-table"><summary>Show date-by-date strategy and recommendation column</summary><div className="strategy-table-scroll"><table><thead><tr><th>Date</th><th>Signal</th><th>Top 1</th><th>Actual FB / GB / GL / DS</th><th>Result</th><th>Net units</th><th>Strategy / recommendation</th></tr></thead><tbody>{[...strategyRows].reverse().map(r=><tr key={`${r.forecast.month}-${r.forecast.day}`}><td>{prettyDate(2026,r.forecast.month,r.forecast.day)}</td><td>{r.score}/100</td><td><b>{r.candidate}</b></td><td>{r.forecast.actual?.join(' · ')}</td><td><span className={`strategy-result ${r.hit?'strategy-hit':'strategy-miss'}`}>{r.hit?'Hit':'Miss'}</span></td><td>{r.unitNet>0?'+':''}{r.unitNet}</td><td>{r.recommendation}</td></tr>)}</tbody></table></div></details><footer>Backtest figures reuse the existing forecast logic and uploaded workbooks. The rule was identified from this same sample, so it may be overfit and can fail in future dates.</footer></section>
              </>:<div className="rounded-xl border border-[#dce4dd] bg-white p-10 text-center text-sm text-[#718078]">No missing dates found for {MONTHS[forecastMonth-1]} 2026.</div>}
            </>}
          </div>:focus==='Single-digit patterns'?<section className="digit-pattern-view"><div className="digit-view-head"><div><p>Single-digit frequency filter</p><h3>Digits before {digitCutoff} {MONTHS[digitMonth-1]}</h3><span>Counts each digit separately in the same game position. A result such as 34 contributes one count to digit 3 and one to digit 4.</span></div><div className="digit-controls"><label className="control"><CalendarDays size={13}/><select value={digitMonth} onChange={e=>setDigitMonth(Number(e.target.value))}>{MONTHS.map((name,i)=><option key={name} value={i+1}>{name}</option>)}</select><ChevronDown size={12}/></label><label className="control"><select value={digitCutoff} onChange={e=>setDigitCutoff(Number(e.target.value))}>{Array.from({length:new Date(2026,digitMonth,0).getDate()},(_,i)=>i+1).map(day=><option key={day} value={day}>Before day {day}</option>)}</select><ChevronDown size={12}/></label><label className="control"><Filter size={13}/><select value={digitSource} onChange={e=>setDigitSource(e.target.value)}><option>All sources</option><option>Previous year</option><option>Previous month</option><option>Earlier current month</option></select><ChevronDown size={12}/></label></div></div><div className="digit-source-note"><Info size={15}/><span>Using {digitSourceRows.length} dated records from {digitSource==='All sources'?`${MONTHS[digitMonth-1]} 2025, ${digitMonth>1?MONTHS[digitMonth-2]+' 2026, ':''}and earlier ${MONTHS[digitMonth-1]} 2026 dates`:digitSource}. Percentages are observed digit frequency, not the probability of a future result.</span></div><div className="digit-game-grid">{digitLeaders.map((leaders,p)=><article key={GAME_NAMES[p]}><header><strong>{GAME_NAMES[p]}</strong><span>{digitTotals[p]} digit observations</span></header><div>{leaders.map((item,i)=><div key={item.digit} className={i===0?'digit-leader':''}><b>{item.digit}</b><span>{item.count} matches</span><em>{(item.share*100).toFixed(1)}%</em></div>)}</div></article>)}</div><div className="digit-table-wrap"><table><thead><tr><th>Digit</th>{GAME_NAMES.map(name=><th key={name}>{name} count</th>)}{GAME_NAMES.map(name=><th key={`${name}-share`}>{name} frequency</th>)}</tr></thead><tbody>{Array.from({length:10},(_,digit)=><tr key={digit}><td><b>{digit}</b></td>{digitCounts.map((counts,p)=><td key={`${p}-c`}>{counts[digit]}</td>)}{digitCounts.map((counts,p)=><td key={`${p}-s`}><span className="digit-bar"><i style={{width:`${digitTotals[p]?counts[digit]/digitTotals[p]*100:0}%`}}></i></span><strong>{digitTotals[p]?`${(counts[digit]/digitTotals[p]*100).toFixed(1)}%`:'—'}</strong></td>)}</tr>)}</tbody></table></div><footer>This filter does not change the existing exact, reverse, position, forecast, or strategy calculations.</footer></section>:<><div className="mb-3 flex flex-col justify-between gap-3 rounded-xl border border-[#dce4dd] bg-white p-3 xl:flex-row xl:items-center"><div><h3 className="font-bold">{viewTitle}</h3><p className="text-[11px] text-[#718078]">{viewSubtitle}</p></div><div className="flex flex-wrap gap-2"><label className="control"><Filter size={13}/><select value={source} onChange={e=>setSource(e.target.value)}><option>All sources</option><option>Previous year</option><option>Previous month</option></select><ChevronDown size={12}/></label><label className="control"><select value={kind} onChange={e=>setKind(e.target.value)} disabled={focus==='Same-position exact'}><option>All matches</option><option>Exact</option><option>Reverse</option></select><ChevronDown size={12}/></label><label className="control min-w-[145px]"><Search size={13}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Date or number"/></label></div></div>
            {focus==='Common day patterns'&&<div className="mb-3 rounded-xl border border-[#dce4dd] bg-white p-3"><div className="mb-2 flex items-center justify-between"><div><strong className="text-xs">Ranked common days</strong><p className="text-[10px] text-[#7a887f]">Coverage shows how many 2026 months contain at least one exact or reverse match on that day.</p></div><span className="badge badge-exact">Top: {dayPatterns[0]?.months||0}/12 months</span></div><div className="flex flex-wrap gap-1.5">{dayPatterns.map(p=><button key={p.day} onClick={()=>setCommonDay(p.day)} className={`day-chip ${selectedCommonDay===p.day?'day-chip-active':''}`}><b>{String(p.day).padStart(2,'0')}</b><span>{p.months}/12</span><small>{p.matches} matches</small></button>)}</div></div>}
            <div className="overflow-hidden rounded-2xl border border-[#d9e2dc] bg-white"><div className="overflow-x-auto"><table className="w-full min-w-[900px] border-collapse text-left"><thead><tr className="border-b border-[#dfe6e1] bg-[#f2f6f3] text-[10px] uppercase tracking-[.11em] text-[#697970]"><th>Date</th><th>2026 game results</th><th>Matched value</th><th>Match type</th><th>Found in</th><th>Comparison</th></tr></thead><tbody>{filteredRows.map(row=>row.matches.map((m,i)=><tr key={`${row.month}-${row.day}-${i}`} className="border-b border-[#edf1ee] last:border-0 hover:bg-[#fafcfa]"><td>{i===0&&<div><strong className="text-sm">{String(row.day).padStart(2,'0')} {MONTHS[row.month-1].slice(0,3)}</strong><span className="block text-[10px] text-[#819087]">2026 · {row.matches.length} matches</span></div>}</td><td>{i===0&&<div className="flex gap-1.5">{row.games.map((g,gi)=><span key={gi} className={`game game-labeled ${row.matches.some(x=>x.currentGame===gi+1)?'game-hit':''}`} data-game={GAME_NAMES[gi]} title={GAME_NAMES[gi]}>{g}</span>)}</div>}</td><td><div className="flex items-center gap-2"><span className="match-number">{m.current}</span><span className="text-[10px] font-bold text-[#7a887f]">{GAME_NAMES[m.currentGame-1]}</span></div></td><td><span className={`badge ${m.kind==='Reverse'?'badge-reverse':'badge-exact'}`}>{m.kind==='Reverse'&&<RotateCcw size={11}/>} {m.kind}</span>{m.kind==='Exact'&&m.currentGame===m.comparedGame&&<span className="ml-1 badge badge-position">Same position</span>}</td><td><span className={`badge ${m.source==='Previous year'?'badge-year':'badge-month'}`}>{m.source}</span></td><td><strong className="text-xs">{m.compared}</strong><span className="ml-1 text-[10px] font-bold text-[#718078]">{GAME_NAMES[m.comparedGame-1]}</span><span className="block text-[10px] text-[#718078]">{m.date}</span></td></tr>))}{filteredRows.length===0&&<tr><td colSpan={6}><div className="grid place-items-center py-16 text-center"><span className="mb-3 grid h-10 w-10 place-items-center rounded-full bg-[#edf3ef] text-[#6f8076]"><Search size={18}/></span><strong className="text-sm">No matches for this view</strong><p className="mt-1 text-xs text-[#7a887f]">Try another month or clear the filters.</p></div></td></tr>}</tbody></table></div></div></>}
          </div>
        </div>
      </section>}
    </div>
  </main>;
}
