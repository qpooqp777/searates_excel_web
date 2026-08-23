// Chartroom Ledger（簡潔藍白版）：Excel → 地點對照 → SeaRates URL → 匯出。
import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { Check, ClipboardPaste, Download, FileSpreadsheet, Link2, MapPin, Upload, X, Plus, Pencil, Trash2, Search, Save, Database } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";

const PASTE_HEADERS = ["報單號碼", "出口港（裝貨港）", "出口港位置", "進口港（卸存地）", "港口位置", "預估運輸距離", "查詢連結", "查詢畫面", "案件日期", "運輸項目", "總重量"];

const HEADER_CANDIDATES = {
  origin: ["出口港（裝貨港）", "出口港", "出發港"],
  destination: ["港口位置", "進口港（卸存地）", "目的港", "目的地"],
};

const LOCATIONS: Record<string, { sea?: string; air?: string; label: string }> = {
  "基隆港": { sea: "Keelung, TW", label: "Keelung, TW" },
  "台北港": { sea: "Taipei, Taipei, TW", label: "Taipei, Taipei, TW" },
  "高雄港": { sea: "Kaohsiung, TW", label: "Kaohsiung, TW" },
  "高雄機場": { air: "Kaohsiung, TW", label: "Kaohsiung, TW" },
  "桃園機場": { air: "Taoyuan, Taoyuan City, TW", label: "Taoyuan, Taoyuan City, TW" },
  "VALENCIA": { sea: "Valencia, Valencian Community, ES", label: "Valencia, Valencian Community, ES" },
  "MUNDRA": { sea: "Mundra, Gujarat, IN", label: "Mundra, Gujarat, IN" },
  "SALALAH": { sea: "Salalah, Dhofar, OM", label: "Salalah, Dhofar, OM" },
  "SOHAR": { sea: "Sohar, Al Batinah, OM", label: "Sohar, Al Batinah, OM" },
  "JEDDAH": { sea: "Jeddah, Makkah Region, SA", label: "Jeddah, Makkah Region, SA" },
  "JAKARTA, JAVA": { sea: "Jakarta, Java, Java, ID", air: "Jakarta, Java, Java, ID", label: "Jakarta, Java, Java, ID" },
  "JAKARTA, JAVA/TANI": { sea: "Jakarta, Java, Java, ID", air: "Jakarta, Java, Java, ID", label: "Jakarta, Java, Java, ID" },
  "BARCELONA": { sea: "Barcelona, Catalonia, ES", label: "Barcelona, Catalonia, ES" },
  "NHAVA SHEVA (JAWAHAR": { sea: "Nhava Sheva, Maharashtra, IN", label: "Nhava Sheva, Maharashtra, IN" },
  "NHAVA SHEVA (JAWA": { sea: "Nhava Sheva, Maharashtra, IN", label: "Nhava Sheva, Maharashtra, IN" },
  "HONG KONG": { sea: "Hong Kong, HK", air: "Hong Kong, HK", label: "Hong Kong, HK" },
  "SOEKARNO-HATTA APT": { air: "Soekarno-Hatta International Airport, Java, ID", label: "Soekarno-Hatta International Airport, Java, ID" },
  "BILBAO": { air: "Bilbao, ES", sea: "Bilbao, ES", label: "Bilbao, ES" },
  "DAMMAM": { air: "Ad Dammam, Eastern Province, SA", sea: "Ad Dammam, Eastern Province, SA", label: "Ad Dammam, Eastern Province, SA" },
};

type LocationEntry = { type: "seaport" | "airport"; sea?: string; air?: string; aliases: string[]; unlocode?: string; iata?: string };
type LocationCodes = { schema_version: string; description?: string; locations: Record<string, LocationEntry> };

const normalize = (value: unknown) => String(value ?? "").trim();
const findHeader = (headers: string[], candidates: string[]) => candidates.find((x) => headers.includes(x));
const findLocation = (raw: string) => {
  const exact = LOCATIONS[raw];
  if (exact) return exact;
  const upper = raw.toUpperCase();
  const key = Object.keys(LOCATIONS).find((item) => upper.includes(item));
  return key ? LOCATIONS[key] : undefined;
};

function makeSeaRatesUrl(from: string, to: string, mode: "海運" | "空運") {
  const params = new URLSearchParams({
    from,
    to,
    transportMode: mode === "空運" ? "Air" : "Sea",
    routingMode: "short",
    fromPlaceType: mode === "空運" ? "airport" : "seaport",
    toPlaceType: mode === "空運" ? "airport" : "seaport",
  });
  return `https://www.searates.com/distance-time?${params.toString()}`;
}

function makeSeaRatesUrlFromCodes(from: string, to: string, mode: "海運" | "空運") {
  const normalizeCode = (value: string) => value.trim().replace(/\s+/g, "+");
  const transportMode = mode === "空運" ? "Air" : "Sea";
  const placeType = mode === "空運" ? "airport" : "seaport";
  return `https://www.searates.com/distance-time?from=${normalizeCode(from)}&to=${normalizeCode(to)}&transportMode=${transportMode}&routingMode=short&fromPlaceType=${placeType}&toPlaceType=${placeType}`;
}

function CodebookManager() {
  const [codes, setCodes] = useState<LocationCodes | null>(null);
  const [query, setQuery] = useState("");
  const [editingKey, setEditingKey] = useState<string | null>(null);
  const [error, setError] = useState("");
  const blank: LocationEntry = { type: "seaport", aliases: [] };
  const [form, setForm] = useState<LocationEntry>(blank);
  const [name, setName] = useState("");

  useEffect(() => {
    fetch("/searates_location_codes.json")
      .then((response) => response.json() as Promise<LocationCodes>)
      .then(setCodes)
      .catch(() => setError("代碼表載入失敗，請確認 JSON 檔案存在。"));
  }, []);

  const entries = useMemo(() => Object.entries(codes?.locations ?? {}).filter(([key, value]) => {
    const needle = query.toLowerCase().trim();
    if (!needle) return true;
    return [key, value.type, value.sea, value.air, value.unlocode, value.iata, ...(value.aliases ?? [])].join(" ").toLowerCase().includes(needle);
  }), [codes, query]);

  const startNew = () => { setEditingKey(null); setName(""); setForm({ ...blank }); setError(""); };
  const startEdit = (key: string, value: LocationEntry) => { setEditingKey(key); setName(key); setForm({ ...value, aliases: [...(value.aliases ?? [])] }); setError(""); window.scrollTo({ top: 0, behavior: "smooth" }); };
  const saveEntry = () => {
    const trimmed = name.trim();
    if (!trimmed || !form.aliases?.length) { setError("請填寫中文／顯示名稱與至少一個別名。別名可使用逗號分隔。"); return; }
    if (!form.sea && !form.air) { setError("海運或空運至少需要一個英文 URL 參數。"); return; }
    setCodes((current) => current ? { ...current, locations: { ...current.locations, [trimmed]: { ...form, aliases: form.aliases.map((item) => item.trim()).filter(Boolean) } } } : current);
    setEditingKey(trimmed); setName(trimmed); setError("");
  };
  const deleteEntry = (key: string) => { if (window.confirm(`確定刪除「${key}」？`)) { setCodes((current) => { if (!current) return current; const next = { ...current.locations }; delete next[key]; return { ...current, locations: next }; }); if (editingKey === key) startNew(); } };
  const exportJson = () => { if (!codes) return; const blob = new Blob([JSON.stringify(codes, null, 2)], { type: "application/json;charset=utf-8" }); const url = URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = url; anchor.download = "searates_location_codes_updated.json"; anchor.click(); URL.revokeObjectURL(url); };

  return <div className="codebook-manager">
    <div className="codebook-toolbar"><div><p className="eyebrow">JSON 對照表 · 本機編輯</p><h3>港口與機場代碼</h3><p className="subhead">修改只會留在目前瀏覽器工作階段，按匯出 JSON 才會下載檔案。</p></div><div className="codebook-actions"><Button variant="outline" onClick={startNew}><Plus size={15} />新增地點</Button><Button onClick={exportJson} disabled={!codes} className="export-button"><Download size={15} />匯出 JSON</Button></div></div>
    {error && <div className="error-banner"><X size={16} />{error}</div>}
    <div className="codebook-layout">
      <Card className="editor-card"><CardHeader><div><p className="eyebrow">編輯欄位</p><CardTitle>{editingKey ? "修改地點" : "新增地點"}</CardTitle></div><Database size={19} className="section-icon" /></CardHeader><CardContent>
        <label className="field-label">中文／顯示名稱<input className="code-input" value={name} onChange={(event) => setName(event.target.value)} placeholder="例如：新加坡港" /></label>
        <label className="field-label">地點類型<select className="code-input" value={form.type} onChange={(event) => setForm({ ...form, type: event.target.value as LocationEntry["type"] })}><option value="seaport">seaport · 海運</option><option value="airport">airport · 空運</option></select></label>
        <label className="field-label">海運 URL 參數<input className="code-input mono" value={form.sea ?? ""} onChange={(event) => setForm({ ...form, sea: event.target.value || undefined })} placeholder="Singapore,+SG" /></label>
        <label className="field-label">空運 URL 參數<input className="code-input mono" value={form.air ?? ""} onChange={(event) => setForm({ ...form, air: event.target.value || undefined })} placeholder="Singapore,+SG" /></label>
        <label className="field-label">UN/LOCODE<input className="code-input" value={form.unlocode ?? ""} onChange={(event) => setForm({ ...form, unlocode: event.target.value || undefined })} placeholder="SGSIN" /></label>
        <label className="field-label">IATA<input className="code-input" value={form.iata ?? ""} onChange={(event) => setForm({ ...form, iata: event.target.value || undefined })} placeholder="SIN" /></label>
        <label className="field-label">別名（逗號分隔）<textarea className="code-input code-textarea" value={(form.aliases ?? []).join(", ")} onChange={(event) => setForm({ ...form, aliases: event.target.value.split(",") })} placeholder="SINGAPORE, Singapore" /></label>
        <div className="editor-actions"><Button onClick={saveEntry} className="save-button"><Save size={15} />儲存對應</Button>{editingKey && <Button variant="outline" onClick={startNew}>清除編輯</Button>}</div>
      </CardContent></Card>
      <Card className="directory-card"><CardHeader><div><p className="eyebrow">目前資料 · {codes ? Object.keys(codes.locations).length : "—"} 筆</p><CardTitle>搜尋代碼表</CardTitle></div><div className="search-box"><Search size={15} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="搜尋名稱、別名或代碼" /></div></CardHeader><CardContent><div className="code-table-wrap"><table className="code-table"><thead><tr><th>名稱</th><th>類型</th><th>SeaRates 參數</th><th>代碼</th><th>操作</th></tr></thead><tbody>{entries.map(([key, value]) => <tr key={key}><td><strong>{key}</strong><small>{value.aliases?.slice(0, 3).join(" · ")}</small></td><td><span className={`mode-pill ${value.type === "airport" ? "air" : "sea"}`}>{value.type === "airport" ? "空運" : "海運"}</span></td><td className="mono code-cell">{value.sea || value.air || "—"}</td><td className="mono">{value.iata || value.unlocode || "—"}</td><td><div className="row-actions"><button type="button" onClick={() => startEdit(key, value)} aria-label={`修改 ${key}`}><Pencil size={14} /></button><button type="button" onClick={() => deleteEntry(key)} aria-label={`刪除 ${key}`}><Trash2 size={14} /></button></div></td></tr>)}</tbody></table>{!entries.length && <div className="empty-state"><Search size={24} /><strong>找不到符合的地點</strong><span>試試中文名稱、英文別名或 IATA／UN/LOCODE。</span></div>}</div></CardContent></Card>
    </div>
  </div>;
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [fileName, setFileName] = useState("");
  const [rows, setRows] = useState<Record<string, unknown>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [sheetName, setSheetName] = useState("");
  const [error, setError] = useState("");
  const [isReady, setIsReady] = useState(false);
  const [inputMode, setInputMode] = useState<"excel" | "text">("excel");
  const [pastedText, setPastedText] = useState("");
  const [activeView, setActiveView] = useState<"generator" | "codebook">("generator");

  const outputRows = useMemo<Record<string, unknown>[]>(() => {
    if (!rows.length) return [];
    if (inputMode === "text") {
      return rows.map((row) => {
        const rawMode = normalize(row.運輸方式).toLowerCase();
        const mode: "海運" | "空運" = rawMode.includes("air") || rawMode.includes("空") ? "空運" : "海運";
        const from = normalize(row.出發地英文代碼);
        const to = normalize(row.目的地英文代碼);
        const url = from && to ? makeSeaRatesUrlFromCodes(from, to, mode) : "";
        return { ...row, 運輸方式: mode, SeaRates查詢連結: url, 查詢狀態: url ? "完成" : "待補三欄資料" };
      });
    }
    const originHeader = findHeader(headers, HEADER_CANDIDATES.origin);
    const destinationHeader = findHeader(headers, HEADER_CANDIDATES.destination);
    return rows.map((row) => {
      const originRaw = normalize(originHeader ? row[originHeader] : "");
      const destinationRaw = normalize(destinationHeader ? row[destinationHeader] : "");
      const mode: "海運" | "空運" = destinationRaw.includes("機場") ? "空運" : "海運";
      const origin = findLocation(originRaw);
      const destination = findLocation(destinationRaw);
      const from = origin?.[mode === "空運" ? "air" : "sea"];
      const to = destination?.[mode === "空運" ? "air" : "sea"];
      const url = from && to ? makeSeaRatesUrl(from, to, mode) : "";
      return {
        ...row,
        運輸方式: mode,
        出發地英文代碼: from || "",
        目的地英文代碼: to || "",
        SeaRates查詢連結: url,
        查詢狀態: url ? "完成" : `待補對照：${originRaw || "出發地"} → ${destinationRaw || "目的地"}`,
      };
    });
  }, [headers, inputMode, rows]);

  const resultHeaders = useMemo(() => {
    const extra = ["運輸方式", "出發地英文代碼", "目的地英文代碼", "SeaRates查詢連結", "查詢狀態"];
    return Array.from(new Set([...headers, ...extra]));
  }, [headers]);

  const successCount = outputRows.filter((row) => row.查詢狀態 === "完成").length;
  const pendingCount = outputRows.length - successCount;

  const loadFile = async (file?: File) => {
    if (!file) return;
    setError("");
    if (!/\.xlsx?$/i.test(file.name)) {
      setError("請選擇 .xlsx 或 .xls 檔案。");
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const name = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[name];
      const matrix = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: "" });
      const rawHeaders = (matrix[0] || []).map((cell) => normalize(cell));
      const data = matrix.slice(1).filter((line) => line.some((cell) => normalize(cell) !== ""));
      setHeaders(rawHeaders);
      setRows(data.map((line) => Object.fromEntries(rawHeaders.map((header, index) => [header, line[index] ?? ""]))));
      setSheetName(name);
      setFileName(file.name);
      setIsReady(true);
    } catch {
      setError("檔案讀取失敗，請確認 Excel 格式與第一列欄位標題。");
    }
  };

  const loadPastedText = () => {
    setError("");
    const text = pastedText.trim();
    if (!text) {
      setError("請先貼上一行三欄資料：運輸方式、出發地英文代碼、目的地英文代碼。");
      return;
    }
    const lines = text.split(/\r?\n/).filter((line) => line.trim());
    const delimiter = lines[0].includes("\t") ? "\t" : lines[0].includes(",") ? "," : "\t";
    const parsed = lines.map((line) => line.split(delimiter).map((cell) => cell.trim().replace(/^"|"$/g, "")));
    const data = parsed.filter((line) => normalize(line[0]) !== "報單號碼" && line.some((cell) => normalize(cell) !== ""));
    if (!data.length) {
      setError("請貼上至少一行三欄資料：運輸方式、出發地英文代碼、目的地英文代碼。");
      return;
    }
    const rawHeaders = ["運輸方式", "出發地英文代碼", "目的地英文代碼"];
    setHeaders(rawHeaders);
    setRows(data.map((line) => Object.fromEntries(rawHeaders.map((header, index) => [header, line[index] ?? ""]))));
    setSheetName("貼上資料");
    setFileName("貼上文字資料.txt");
    setIsReady(true);
  };

  const exportFile = () => {
    if (!outputRows.length) return;
    const worksheet = XLSX.utils.json_to_sheet(outputRows, { header: resultHeaders });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName || "報單彙整");
    XLSX.writeFile(workbook, `${fileName.replace(/\.(xlsx?|xls)$/i, "")}_SeaRates網址.xlsx`);
  };

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand-lockup">
          <div className="brand-mark" aria-hidden="true"><span /><span /></div>
          <div><p className="eyebrow">SEA ROUTE WORKBENCH</p><h1>SeaRates Excel 網址產生器</h1></div>
        </div>
        <div className="topbar-note"><span className="status-dot" />本機處理，不上傳檔案</div>
      </header>

      <main className="workspace">
        <aside className="workflow-panel">
          <div className="panel-intro"><p className="eyebrow">工作流程</p><h2>把報單帶上船，連結一次整理。</h2><p>上傳 Excel，系統會依目的地判斷海運／空運，帶出英文代碼與 SeaRates 查詢網址。</p></div>
          <div className="steps">
            <div className={`step ${fileName ? "done" : "active"}`}><span className="step-number">01</span><div><strong>上傳報單</strong><small>{fileName || "選擇 Excel 檔案"}</small></div></div>
            <div className={`step ${isReady ? "active" : ""}`}><span className="step-number">02</span><div><strong>對照網址</strong><small>{isReady ? `${outputRows.length} 列已讀取` : "等待檔案"}</small></div></div>
            <div className="step"><span className="step-number">03</span><div><strong>匯出結果</strong><small>下載更新後 Excel</small></div></div>
          </div>
          <div className="code-note"><MapPin size={16} /><span>港口位置含「機場」時，自動使用空運；其他目的地使用海運。</span></div>
        </aside>

        <section className="content-panel">
          <div className="content-heading"><div><p className="eyebrow">批次作業 · SeaRates</p><h2>Excel 轉 SeaRates 連結</h2><p className="subhead">欄位會在瀏覽器內完成對照，原始檔案不會離開你的裝置。</p></div><Badge variant="outline" className="secure-badge"><span className="status-dot" />安全處理</Badge></div>

          <div className="workspace-tabs" role="tablist" aria-label="工具功能"><button type="button" className={`workspace-tab ${activeView === "generator" ? "selected" : ""}`} onClick={() => setActiveView("generator")}><Link2 size={15} />網址產生器</button><button type="button" className={`workspace-tab ${activeView === "codebook" ? "selected" : ""}`} onClick={() => setActiveView("codebook")}><Database size={15} />代碼表管理</button></div>
          {activeView === "codebook" ? <CodebookManager /> : <>
          <div className="input-tabs" role="tablist" aria-label="資料輸入方式">
            <button type="button" role="tab" aria-selected={inputMode === "excel"} className={`input-tab ${inputMode === "excel" ? "selected" : ""}`} onClick={() => setInputMode("excel")}><FileSpreadsheet size={15} />上傳 Excel</button>
            <button type="button" role="tab" aria-selected={inputMode === "text"} className={`input-tab ${inputMode === "text" ? "selected" : ""}`} onClick={() => setInputMode("text")}><ClipboardPaste size={15} />貼上文字</button>
          </div>
          {inputMode === "excel" ? <div className="upload-card" onDragOver={(event) => event.preventDefault()} onDrop={(event) => { event.preventDefault(); void loadFile(event.dataTransfer.files[0]); }}>
            <div className="upload-icon"><Upload size={22} /></div>
            <div className="upload-copy"><strong>{fileName || "拖曳 Excel 檔案到這裡"}</strong><span>{fileName ? "已載入，可立即檢查並匯出" : "支援 .xlsx 與 .xls 格式"}</span></div>
            <Button onClick={() => inputRef.current?.click()} variant="outline" className="choose-button"><FileSpreadsheet size={16} />{fileName ? "更換檔案" : "選擇檔案"}</Button>
            <Input ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden-input" onChange={(event) => void loadFile(event.target.files?.[0])} />
          </div> : <div className="paste-card">
            <div className="paste-heading"><div className="upload-icon"><ClipboardPaste size={22} /></div><div><strong>貼上三欄資料</strong><span>運輸方式、出發地英文代碼、目的地英文代碼</span></div></div>
            <textarea className="paste-area" value={pastedText} onChange={(event) => setPastedText(event.target.value)} placeholder="請貼上三欄資料，不需要欄位名稱……\n海運    Valencia,+Valencian+Community,+ES    Keelung,+TW\n海運    Valencia,+Valencian+Community,+ES    Keelung,+TW" aria-label="貼上 Excel 文字內容" />
            <div className="paste-actions"><span>每行一筆；Tab 或逗號分隔</span><Button onClick={loadPastedText} className="paste-button"><ClipboardPaste size={16} />讀取貼上資料</Button></div>
          </div>}
          {error && <div className="error-banner"><X size={16} />{error}</div>}

          <div className="stat-grid">
            <Card><CardContent><span className="stat-label">已讀取列數</span><strong>{outputRows.length}</strong><small>{sheetName || "尚未載入"}</small></CardContent></Card>
            <Card><CardContent><span className="stat-label">可匯出網址</span><strong className="blue-number">{successCount}</strong><small>海運與空運皆可</small></CardContent></Card>
            <Card><CardContent><span className="stat-label">待補對照</span><strong className={pendingCount ? "orange-number" : ""}>{pendingCount}</strong><small>{pendingCount ? "請更新代碼表" : "全部完成"}</small></CardContent></Card>
          </div>

          <Card className="preview-card"><CardHeader><div><p className="eyebrow">結果預覽 · 逐列核對</p><CardTitle>網址欄位預覽</CardTitle></div><Button onClick={exportFile} disabled={!isReady || !outputRows.length} className="export-button"><Download size={16} />匯出更新後 Excel</Button></CardHeader><CardContent>
            {!outputRows.length ? <div className="empty-state"><FileSpreadsheet size={28} /><strong>等待 Excel 檔案</strong><span>上傳後會在這裡顯示前幾筆自動產生的結果。</span></div> : <div className="table-wrap"><table><thead><tr><th>列</th><th>運輸</th><th>出發地英文代碼</th><th>目的地英文代碼</th><th>SeaRates 查詢網址</th><th>狀態</th></tr></thead><tbody>{outputRows.slice(0, 8).map((row, index) => <tr key={index}><td className="row-index">{index + 2}</td><td><span className={`mode-pill ${row.運輸方式 === "空運" ? "air" : "sea"}`}>{row.運輸方式 as string}</span></td><td className="mono">{(row.出發地英文代碼 as string) || "—"}</td><td className="mono">{(row.目的地英文代碼 as string) || "—"}</td><td>{row.SeaRates查詢連結 ? <a className="url-link" href={row.SeaRates查詢連結 as string} target="_blank" rel="noreferrer"><Link2 size={14} />開啟查詢網址</a> : <span className="muted">尚待對照</span>}</td><td><span className={row.查詢狀態 === "完成" ? "success-state" : "pending-state"}>{row.查詢狀態 === "完成" && <Check size={13} />}{row.查詢狀態 as string}</span></td></tr>)}</tbody></table></div>}
          </CardContent></Card>
          {outputRows.length > 8 && <p className="table-footnote">目前預覽前 8 列，匯出時會包含全部 {outputRows.length} 列。</p>}
          </>}
        </section>
      </main>
    </div>
  );
}
