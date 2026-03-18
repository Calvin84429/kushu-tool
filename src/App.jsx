import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { dbGetClients, dbUpsertClient, dbDeleteClient, dbGetMeta, dbSetMeta, dbGetShortcuts, dbSetShortcuts, dbGetTrash, dbSetTrash } from "./supabase.js";

// ─── SYSTEM PROMPTS ───────────────────────────────────────────────────────────
const BASE = `你是空書道的AI客服助手，幫助客服人員Lucy生成回覆訊息。

━━━━━━━━━━━━━━━━━━━━━━━━━━━
【空書道核心思想體系】
━━━━━━━━━━━━━━━━━━━━━━━━━━━

空書道不是算命，是一套「個體能量狀態管理學」。

核心比喻：把人比作一部智能手機
- 身體與先天稟賦 = 硬件
- 思維、情緒、能量運行模式 = 操作系統
- 事業、財富、感情、健康 = 各個APP
- 精力值與對外界機會的感知能力 = 電量與信號強度

大多數人的問題不是「APP壞了」，而是：手機電量長期低於20%、後台堆滿未清理緩存、網絡信號時斷時續。
這就是為什麼無論多努力，事情都容易卡頓、出錯、不順。

「沐氣」= 清理緩存、關閉無用後台進程、連接穩定電源
「啟運」= 在系統潔淨的基礎上，精準校準個人頻率與外界良性互動節點，讓能量順流

核心主張：你的不順，不是天注定，是你的「系統」需要一次深度的專業維護與升級。
我們不改命，我們重啟氣運、對齊時勢，讓你成為命運的主導者。

━━━━━━━━━━━━━━━━━━━━━━━━━━━
【六大客戶問題的空書道解讀框架】
━━━━━━━━━━━━━━━━━━━━━━━━━━━

當客戶說出問題時，用以下框架「重新命名」問題，讓客戶感受到被深度看見：

1.【事業走不順、做不起來、一直在耗】
解讀：你不是能力不夠，是你的氣場節律跟當前的時機錯位了。
每一個行動都在逆著氣走，消耗的是你自己。不是沒有能力，是力氣全花在對抗上。

2.【財運不好、留不住錢、負債壓身】
解讀：不是你賺不到，是身上有漏財的破口，財進得來但守不住。
這個破口通常跟「人」有關——某段關係、某個決策模式、某種長期的內耗正在悄悄抽走你的財氣。

3.【感情、家庭關係紊亂】
解讀：關係裡的摩擦，本質是兩個人的頻率對不上，不是誰對誰錯的問題。
是能量場在互相消耗——當一方的氣場持續低迷，另一方再有耐心都會被拖著走。

4.【身體狀態差、疑難雜症長治不癒】
解讀：身體是最誠實的傳感器。它告訴你，你的內在系統已經長期超載。
這不只是身體的問題，身體只是最後一個發出警報的地方，能量層面的堵塞早就存在了。

5.【內耗、沒方向、心態消極】
解讀：你不是懶，也不是廢，是你的氣被堵住了。
所有能量都在原地打轉出不來——這種狀態不是性格問題，是系統需要清理和重啟。

6.【工作壓力大、人際關係複雜】
解讀：周圍的人際消耗，是因為你的磁場現在是散的。
吸引不到對的人，卻一直招來消耗你的——當一個人的氣場渾濁，就像手機信號差，只會連上雜訊。

━━━━━━━━━━━━━━━━━━━━━━━━━━━
【服務方案】
━━━━━━━━━━━━━━━━━━━━━━━━━━━

- 一季 $36,000（3個月）
- 二季 $68,000（6個月）
- 一年 $120,000（12個月）
- 付款：匯款、信用卡、無卡分期（36期月付約3,900元）

服務內容：每月老師遠距啟運、定期通話指導、氣盤節奏指引、老師訂製字畫、擴能水氧機、金屬卡

━━━━━━━━━━━━━━━━━━━━━━━━━━━
【客戶分級】
━━━━━━━━━━━━━━━━━━━━━━━━━━━

A級：職業穩定、需求明確、問題超過3個月 → 主推一年方案＋無卡分期
B級：收入普通有需求但預算有限 → 推一季方案
C級：嚴重債務或需求模糊 → 快速禮貌結束

━━━━━━━━━━━━━━━━━━━━━━━━━━━
【語氣與風格】
━━━━━━━━━━━━━━━━━━━━━━━━━━━

- 精準直接，每句話有目的
- 帶神秘感與權威感，不說廢話
- 先讓客戶感受到「被看見」，再引導需求
- 用「老師」代稱自己、「小編」代稱客服
- 說話像一個真正懂玄學、懂人的人，不像推銷員

格式規定：絕對不使用任何 Markdown 符號（*、**、-列表、#、_）。只輸出純文字，段落之間用換行分隔，可以直接複製貼給客戶。`;

const REPLY_SYS = BASE + `

根據「當前對話階段」提供對應風格的回覆：

【建立信任階段】
用空書道思想框架「重新解讀」客戶的問題，讓他感受到被深度看見。
不急著介紹服務，先幫他把問題說清楚——用能量、氣場、系統失序的語言重新命名他的困境。
例如：客戶說事業不順 → 不說「我們可以幫你」，而是說「你的情況不是能力問題，是氣場節律跟時機錯位了，你的力氣都花在對抗上，而不是前進」
讓客戶的反應是：「對，就是這樣，你怎麼這麼懂我」

【觸發需求階段】
在建立共鳴後，點出問題不處理的代價。
強調這種狀態會繼續累積，拖越久耗越深，時機窗口在流失。
用「再這樣下去」「你這個狀況已經在消耗你的XX」等語氣製造緊迫感。

【報價切入階段】
用無卡分期降低門檻，強調每月不到4千。
把「啟運」說成一次系統性的深度調整，不是花錢買玄學，是為自己的狀態做一次真正的投資。
推進到具體行動：「我幫你跟老師安排一個時間」

直接輸出1-3段可複製貼上的回覆話術。
若需判斷等級，第一行只寫 GRADE:A 或 GRADE:B 或 GRADE:C，換行再寫回覆。
絕不在話術內出現 GRADE 字樣，不要輸出分析說明。`;

const ANALYSE_SYS = BASE + `

用空書道視角深度分析客戶，輸出：
【客戶問題本質】用空書道框架解讀這個人的狀態（對應六大問題類型哪一類，根源是什麼）
【能量卡點】從氣場、節律、系統失序角度判斷他真正被什麼卡住
【成交切入點】他最容易被打動的點是什麼，用哪個角度切入最有共鳴
【建議話術】1-2段用空書道語言寫的回覆，讓他感受到被深度看見`;

const IMG_REPLY_SYS = BASE + `

識別截圖對話（深色/紫色氣泡=小編，淺色/白色氣泡=客戶），直接輸出1-3段回覆話術。
若需判斷等級，第一行只寫 GRADE:A/B/C，換行再寫回覆。不要輸出分析或截圖文字。`;

const IMG_ANALYSE_SYS = BASE + `

識別截圖完整對話，輸出：
【截圖對話內容】逐字讀出所有對話
【客戶分析】等級 / 階段 / 情緒 / 卡點
【建議話術】1-3段可複製的回覆`;

const ASST_SYS = BASE + `
你是AI客服助理，協助Lucy制定成交策略、分析截圖、討論跟進方式。
你完全理解空書道的核心思想：沐氣啟運、六大問題解讀框架、手機系統比喻。
分析客戶時，要從能量狀態、氣場節律、系統失序的角度切入，不只是表面的銷售建議。
繁體中文，語氣專業直接，像一個真正懂空書道思想的顧問。`;

// ─── STORAGE KEYS ─────────────────────────────────────────────────────────────
const SK = "ksd_v7";
const SK_TRASH = "ksd_v7_trash";
const TRASH_TTL = 7 * 24 * 60 * 60 * 1000; // 7 days in ms

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function ls(key, fb) { try { const r = localStorage.getItem(key); return r ? JSON.parse(r) : fb; } catch { return fb; } }
function lsSave(key, v) { try { localStorage.setItem(key, JSON.stringify(v)); } catch {} }
async function wsSave(key, v) { try { await window.storage.set(key, JSON.stringify(v)); } catch {} }
async function wsLoad(key) { try { { const r = await window.storage.get(key); if (r?.value) return JSON.parse(r.value); } } catch {} return null; }

async function callAPI(system, messages, maxTokens = 1200) {
  try {
    const r = await fetch("/api/chat", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: "claude-sonnet-4-5-20250929", max_tokens: maxTokens, system, messages }),
    });
    if (!r.ok) {
      const err = await r.text();
      console.error("API error:", r.status, err);
      return "（連線錯誤：" + r.status + "）";
    }
    const d = await r.json();
    if (d.error) { console.error("API error:", d.error); return "（API錯誤：" + (d.error.message || JSON.stringify(d.error)) + "）"; }
    return d.content?.[0]?.text || "";
  } catch(e) {
    console.error("callAPI exception:", e);
    return "（網路錯誤，請重試）";
  }
}

function imgMsg(b64obj, text) {
  const data = typeof b64obj === "object" ? b64obj.data : b64obj;
  const mime = typeof b64obj === "object" ? b64obj.mime : "image/jpeg";
  return [
    { type: "image", source: { type: "base64", media_type: mime, data } },
    { type: "text", text: text || "請分析截圖。" },
  ];
}

function readImgFile(file, cb) {
  if (!file || !file.type.startsWith("image/")) return;
  const mime = file.type;
  const reader = new FileReader();
  reader.onload = ev => cb({ data: ev.target.result.split(",")[1], mime, preview: ev.target.result, name: file.name, id: "i" + Date.now() + Math.random() });
  reader.readAsDataURL(file);
}

// ─── SMALL COMPONENTS ─────────────────────────────────────────────────────────
function Tag({ label, color, bg, border }) {
  return (
    <span style={{ fontSize: 10, padding: "2px 7px", borderRadius: 10, color, background: bg, border: `1px solid ${border}`, fontWeight: 500, whiteSpace: "nowrap" }}>
      {label}
    </span>
  );
}

function CopyText({ text, label = "複製回覆" }) {
  const [done, setDone] = useState(false);
  function doCopy() {
    function fallback() {
      const el = document.createElement("textarea");
      el.value = text; el.style.cssText = "position:fixed;top:-9999px;opacity:0";
      document.body.appendChild(el); el.focus(); el.select();
      try { document.execCommand("copy"); } catch (e) {}
      document.body.removeChild(el);
      setDone(true); setTimeout(() => setDone(false), 2000);
    }
    if (navigator.clipboard?.writeText) {
      navigator.clipboard.writeText(text).then(() => { setDone(true); setTimeout(() => setDone(false), 2000); }).catch(fallback);
    } else fallback();
  }
  return (
    <button onClick={doCopy} style={{ padding: "4px 12px", borderRadius: 6, fontSize: 11, cursor: "pointer", fontFamily: "inherit", fontWeight: 500, border: done ? "1px solid #10b981" : "1px solid #6B21C8", background: done ? "#f0fdf4" : "white", color: done ? "#10b981" : "#6B21C8" }}>
      {done ? "已複製 ✓" : label}
    </button>
  );
}

function CopyImg({ dataUrl }) {
  const [done, setDone] = useState(false);
  async function doCopy() {
    try {
      const img = new Image(); img.src = dataUrl;
      await new Promise(r => { img.onload = r; });
      const canvas = document.createElement("canvas");
      canvas.width = img.naturalWidth; canvas.height = img.naturalHeight;
      canvas.getContext("2d").drawImage(img, 0, 0);
      canvas.toBlob(async blob => {
        try { await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]); }
        catch { window.open(URL.createObjectURL(blob), "_blank"); }
        setDone(true); setTimeout(() => setDone(false), 2000);
      }, "image/png");
    } catch { setDone(true); setTimeout(() => setDone(false), 2000); }
  }
  return (
    <button onClick={doCopy} style={{ padding: "4px 10px", borderRadius: 6, fontSize: 11, cursor: "pointer", fontFamily: "inherit", fontWeight: 500, border: done ? "1px solid #10b981" : "1px solid #6B21C8", background: done ? "#f0fdf4" : "white", color: done ? "#10b981" : "#6B21C8" }}>
      {done ? "已複製 ✓" : "複製圖片"}
    </button>
  );
}

function Dots() {
  return (
    <div style={{ display: "flex", gap: 4, padding: "10px 14px", borderRadius: 14, background: "#f9fafb", border: "1px solid #e5e7eb" }}>
      {[0, 0.2, 0.4].map((d, i) => (
        <div key={i} style={{ width: 5, height: 5, borderRadius: "50%", background: "#6B21C8", animation: `dot 1.2s ${d}s infinite` }} />
      ))}
    </div>
  );
}

function Popup({ title, onClose, children }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: "white", borderRadius: 16, padding: 24, width: "100%", maxWidth: 420, boxShadow: "0 20px 60px rgba(0,0,0,0.15)", maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <div style={{ fontSize: 15, fontWeight: 700 }}>{title}</div>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#999", cursor: "pointer", fontSize: 18 }}>✕</button>
        </div>
        {children}
      </div>
    </div>
  );
}

// ─── IMAGE GRID ───────────────────────────────────────────────────────────────
function ImgGrid({ images, onPreview }) {
  const [selected, setSelected] = useState(new Set());
  if (!images || images.length === 0) return null;
  const allSel = selected.size === images.length;
  function toggle(id) { setSelected(p => { const n = new Set(p); n.has(id) ? n.delete(id) : n.add(id); return n; }); }
  function toggleAll() { setSelected(allSel ? new Set() : new Set(images.map(i => i.id))); }
  return (
    <div style={{ marginTop: 8 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
        <span style={{ fontSize: 11, color: "#6b7280" }}>圖片（{images.length}張）</span>
        <button onClick={toggleAll} style={{ padding: "2px 8px", borderRadius: 5, fontSize: 11, cursor: "pointer", fontFamily: "inherit", border: "1px solid #e5e7eb", background: allSel ? "#f3f0ff" : "white", color: allSel ? "#6B21C8" : "#6b7280" }}>
          {allSel ? "取消全選" : "全選"}
        </button>
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
        {images.map(img => {
          const isSel = selected.has(img.id);
          return (
            <div key={img.id} style={{ position: "relative", cursor: "pointer" }} onClick={() => toggle(img.id)}>
              <img src={img.dataUrl || img.preview} alt={img.name} style={{ width: 72, height: 72, objectFit: "cover", borderRadius: 8, border: isSel ? "2px solid #6B21C8" : "2px solid transparent", opacity: isSel ? 1 : 0.75, transition: "all 0.15s" }} />
              {isSel && (
                <div style={{ position: "absolute", top: 3, right: 3, width: 16, height: 16, borderRadius: "50%", background: "#6B21C8", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 9, color: "white", fontWeight: 700 }}>✓</div>
              )}
              <button onClick={e => { e.stopPropagation(); onPreview(img.dataUrl || img.preview); }} style={{ position: "absolute", bottom: 3, left: 3, padding: "1px 4px", borderRadius: 4, fontSize: 9, background: "rgba(0,0,0,0.5)", color: "white", border: "none", cursor: "pointer" }}>
                預覽
              </button>
            </div>
          );
        })}
      </div>
      {selected.size > 0 && (
        <div style={{ display: "flex", gap: 6, marginTop: 8, flexWrap: "wrap" }}>
          {images.filter(i => selected.has(i.id)).map(img => (
            <CopyImg key={img.id} dataUrl={img.dataUrl || img.preview} />
          ))}
        </div>
      )}
    </div>
  );
}

// ─── SHORTCUTS PANEL ──────────────────────────────────────────────────────────
const INP = { width: "100%", border: "1px solid #e5e7eb", borderRadius: 8, padding: "8px 12px", fontSize: 13, outline: "none", fontFamily: "inherit", background: "white", boxSizing: "border-box" };

function ShortcutsPanel({ sc, setSc }) {
  const [newCatName, setNewCatName] = useState("");
  const [addingCat, setAddingCat] = useState(false);
  const [addingItem, setAddingItem] = useState(null);
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState({ title: "", text: "", images: [] });
  const [collapsed, setCollapsed] = useState({});
  const [lightbox, setLightbox] = useState(null);
  const [pinnedCats, setPinnedCats] = useState(new Set());
  const [catCtx, setCatCtx] = useState(null);
  const [editCatId, setEditCatId] = useState(null);
  const [editCatName, setEditCatName] = useState("");
  const [dragCat, setDragCat] = useState(null);
  const [overCat, setOverCat] = useState(null);
  const [dragItm, setDragItm] = useState(null);
  const [overItm, setOverItm] = useState(null);
  const addImgRef = useRef(null);
  const editImgRef = useRef(null);

  const cats = sc.categories || [];

  function addCat() {
    const n = newCatName.trim(); if (!n) return;
    const id = "c" + Date.now();
    setSc(p => ({ ...p, categories: [...(p.categories || []), { id, name: n }], items: { ...(p.items || {}), [id]: [] } }));
    setNewCatName(""); setAddingCat(false);
  }
  function delCat(id) {
    const cat = (sc.categories || []).find(c => c.id === id);
    const catItems = sc.items?.[id] || [];
    if (cat) setTrash(p => [...p, { id, name: cat.name, items: catItems, deletedAt: Date.now(), trashType: "shortcut" }]);
    setSc(p => { const cats = (p.categories || []).filter(c => c.id !== id); const items = { ...p.items }; delete items[id]; return { ...p, categories: cats, items }; });
  }
  function restoreCat(item) {
    const { deletedAt, trashType, items: catItems, ...cat } = item;
    setSc(p => ({ ...p, categories: [...(p.categories || []), { id: cat.id, name: cat.name }], items: { ...p.items, [cat.id]: catItems } }));
    setTrash(p => p.filter(t => !(t.id === item.id && t.deletedAt === item.deletedAt)));
  }
  function permDelCat(item) {
    setTrash(p => p.filter(t => !(t.id === item.id && t.deletedAt === item.deletedAt)));
  }
  function pinCat(id) {
    setSc(p => { const cats = [...(p.categories || [])]; const idx = cats.findIndex(c => c.id === id); if (idx < 0) return p; const [m] = cats.splice(idx, 1); cats.unshift(m); return { ...p, categories: cats }; });
    setPinnedCats(p => { const n = new Set(p); n.add(id); return n; });
    setCatCtx(null);
  }
  function unpinCat(id) {
    setPinnedCats(p => { const n = new Set(p); n.delete(id); return n; });
    setCatCtx(null);
  }
  function saveCatName(id, name) {
    if (!name.trim()) return;
    setSc(p => ({ ...p, categories: (p.categories||[]).map(c => c.id === id ? { ...c, name: name.trim() } : c) }));
    setEditCatId(null); setEditCatName("");
  }
  function addItm(catId) {
    if (!form.title.trim() || (!form.text.trim() && form.images.length === 0)) return;
    const id = "i" + Date.now();
    setSc(p => ({ ...p, items: { ...p.items, [catId]: [...(p.items?.[catId] || []), { id, title: form.title.trim(), text: form.text.trim(), images: form.images }] } }));
    setForm({ title: "", text: "", images: [] }); setAddingItem(null);
  }
  function updItm(catId, itemId) {
    if (!form.title.trim()) return;
    setSc(p => ({ ...p, items: { ...p.items, [catId]: (p.items?.[catId] || []).map(it => it.id === itemId ? { ...it, title: form.title.trim(), text: form.text.trim(), images: form.images } : it) } }));
    setEditItem(null); setForm({ title: "", text: "", images: [] });
  }
  function delItm(catId, itemId) {
    setSc(p => ({ ...p, items: { ...p.items, [catId]: (p.items?.[catId] || []).filter(it => it.id !== itemId) } }));
  }
  function addImgs(files) {
    Array.from(files).forEach(f => readImgFile(f, img => setForm(p => ({ ...p, images: [...p.images, img] }))));
  }
  function dropCat(e, toId) {
    e.preventDefault();
    if (!dragCat || dragCat === toId) { setDragCat(null); setOverCat(null); return; }
    setSc(p => { const cats = [...(p.categories || [])]; const fi = cats.findIndex(c => c.id === dragCat); const ti = cats.findIndex(c => c.id === toId); if (fi < 0 || ti < 0) return p; const [m] = cats.splice(fi, 1); cats.splice(ti, 0, m); return { ...p, categories: cats }; });
    setDragCat(null); setOverCat(null);
  }
  function dropItm(e, catId, toId) {
    e.preventDefault();
    if (!dragItm || dragItm.itemId === toId || dragItm.catId !== catId) { setDragItm(null); setOverItm(null); return; }
    setSc(p => { const items = [...(p.items?.[catId] || [])]; const fi = items.findIndex(i => i.id === dragItm.itemId); const ti = items.findIndex(i => i.id === toId); if (fi < 0 || ti < 0) return p; const [m] = items.splice(fi, 1); items.splice(ti, 0, m); return { ...p, items: { ...p.items, [catId]: items } }; });
    setDragItm(null); setOverItm(null);
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
      {lightbox && (
        <div onClick={() => setLightbox(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)", zIndex: 300, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" }}>
          <img src={lightbox} alt="預覽" style={{ maxWidth: "90vw", maxHeight: "90vh", borderRadius: 12 }} />
          <div style={{ position: "absolute", top: 20, right: 20, color: "white", fontSize: 24 }}>✕</div>
        </div>
      )}
      {catCtx && (
        <div onClick={() => setCatCtx(null)} style={{ position: "fixed", inset: 0, zIndex: 400 }}>
          <div onClick={e => e.stopPropagation()} style={{ position: "fixed", top: catCtx.y, left: catCtx.x, background: "white", border: "1px solid #e5e7eb", borderRadius: 10, padding: "4px 0", boxShadow: "0 8px 30px rgba(0,0,0,0.15)", zIndex: 401, minWidth: 140 }}>
            {pinnedCats.has(catCtx.id) ? (
              <button onClick={() => unpinCat(catCtx.id)} style={{ display: "block", width: "100%", padding: "9px 16px", textAlign: "left", background: "none", border: "none", fontSize: 13, cursor: "pointer", fontFamily: "inherit", color: "#374151" }}>📌 取消置頂</button>
            ) : (
              <button onClick={() => pinCat(catCtx.id)} style={{ display: "block", width: "100%", padding: "9px 16px", textAlign: "left", background: "none", border: "none", fontSize: 13, cursor: "pointer", fontFamily: "inherit", color: "#374151" }}>📌 置頂</button>
            )}
          </div>
        </div>
      )}
      {cats.length === 0 && !addingCat && (
        <div style={{ textAlign: "center", color: "#9ca3af", fontSize: 13, padding: "24px 0" }}>還沒有分類，點下方「+ 新增分類」開始</div>
      )}
      {cats.map(cat => {
        const items = sc.items?.[cat.id] || [];
        const isCol = !!collapsed[cat.id];
        return (
          <div key={cat.id} draggable onDragStart={() => setDragCat(cat.id)} onDragOver={e => { e.preventDefault(); setOverCat(cat.id); }} onDrop={e => dropCat(e, cat.id)} onDragEnd={() => { setDragCat(null); setOverCat(null); }} style={{ marginBottom: 10, border: "1px solid #e5e7eb", borderRadius: 10, overflow: "hidden", opacity: dragCat === cat.id ? 0.4 : 1, borderTop: overCat === cat.id ? "2px solid #6B21C8" : "1px solid #e5e7eb" }}>
            <div onContextMenu={e => { e.preventDefault(); e.stopPropagation(); setCatCtx({ x: e.clientX, y: e.clientY, id: cat.id }); }} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "9px 12px", background: "#f9fafb", borderBottom: isCol ? "none" : "1px solid #e5e7eb", cursor: "pointer" }} onClick={() => setCollapsed(p => ({ ...p, [cat.id]: !p[cat.id] }))}>
              <div style={{ fontSize: 13, fontWeight: 700, color: "#374151", display: "flex", alignItems: "center", gap: 6, flex: 1, minWidth: 0 }}>
                <span style={{ fontSize: 10, color: "#9ca3af", display: "inline-block", flexShrink: 0, transform: isCol ? "rotate(-90deg)" : "rotate(0deg)", transition: "transform 0.2s" }}>▼</span>
                <span style={{ width: 4, height: 14, borderRadius: 2, background: "#6B21C8", display: "inline-block", flexShrink: 0 }} />
                {pinnedCats.has(cat.id) && <span style={{ fontSize: 9, flexShrink: 0 }}>📌</span>}
                {editCatId === cat.id ? (
                  <input autoFocus value={editCatName} onChange={e => setEditCatName(e.target.value)}
                    onKeyDown={e => { e.stopPropagation(); if (e.key === "Enter") saveCatName(cat.id, editCatName); if (e.key === "Escape") { setEditCatId(null); setEditCatName(""); } }}
                    onBlur={() => saveCatName(cat.id, editCatName || cat.name)}
                    onClick={e => e.stopPropagation()}
                    style={{ fontSize: 13, fontWeight: 700, border: "1px solid #6B21C8", borderRadius: 5, padding: "1px 6px", outline: "none", fontFamily: "inherit", color: "#6B21C8", minWidth: 0, flex: 1 }} />
                ) : (
                  <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{cat.name}</span>
                )}
                <span style={{ fontSize: 11, color: "#9ca3af", fontWeight: 400, flexShrink: 0 }}>（{items.length}則）</span>
              </div>
              <div style={{ display: "flex", gap: 5, flexShrink: 0 }} onClick={e => e.stopPropagation()}>
                <button onClick={() => { setEditCatId(cat.id); setEditCatName(cat.name); }} style={{ padding: "3px 7px", borderRadius: 6, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", fontSize: 13, cursor: "pointer", fontFamily: "inherit" }} title="編輯分類名稱">✎</button>
                <button onClick={() => { setAddingItem(cat.id); setForm({ title: "", text: "", images: [] }); setCollapsed(p => ({ ...p, [cat.id]: false })); }} style={{ padding: "3px 8px", borderRadius: 6, border: "1px solid #6B21C8", background: "white", color: "#6B21C8", fontSize: 13, cursor: "pointer", fontFamily: "inherit" }} title="新增快捷鍵">+</button>
                <button onClick={() => delCat(cat.id)} style={{ padding: "3px 7px", borderRadius: 6, border: "1px solid #fecaca", background: "white", color: "#ef4444", fontSize: 13, cursor: "pointer", fontFamily: "inherit" }} title="刪除分類">✕</button>
              </div>
            </div>
            {!isCol && (
              <div style={{ padding: "10px 12px", display: "flex", flexDirection: "column", gap: 6 }}>
                {addingItem === cat.id && (
                  <div style={{ background: "#f9fafb", border: "1px solid #e5e7eb", borderRadius: 8, padding: 12, display: "flex", flexDirection: "column", gap: 8 }}>
                    <input style={INP} placeholder="快捷鍵名稱 *" value={form.title} onChange={e => setForm(p => ({ ...p, title: e.target.value }))} autoFocus />
                    <textarea style={{ ...INP, resize: "none", minHeight: 72 }} placeholder="回覆文字（可只上傳圖片）" value={form.text} onChange={e => setForm(p => ({ ...p, text: e.target.value }))} />
                    <div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                        <span style={{ fontSize: 12, color: "#374151", fontWeight: 500 }}>圖片（可多張）</span>
                        <button onClick={() => addImgRef.current?.click()} style={{ padding: "3px 10px", borderRadius: 6, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>+ 上傳</button>
                        <input ref={addImgRef} type="file" accept="image/*" multiple style={{ display: "none" }} onChange={e => { addImgs(e.target.files); e.target.value = ""; }} />
                      </div>
                      {form.images.length > 0 && (
                        <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                          {form.images.map((img, idx) => (
                            <div key={img.id} style={{ position: "relative" }}>
                              <img src={img.preview} alt="" style={{ width: 60, height: 60, objectFit: "cover", borderRadius: 6, border: "1px solid #e5e7eb" }} />
                              <button onClick={() => setForm(p => ({ ...p, images: p.images.filter((_, i) => i !== idx) }))} style={{ position: "absolute", top: -5, right: -5, width: 16, height: 16, borderRadius: "50%", background: "#ef4444", border: "none", color: "white", fontSize: 9, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      <button onClick={() => { setAddingItem(null); setForm({ title: "", text: "", images: [] }); }} style={{ padding: "5px 12px", borderRadius: 6, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", fontSize: 12, cursor: "pointer", fontFamily: "inherit" }}>取消</button>
                      <button onClick={() => addItm(cat.id)} style={{ padding: "5px 12px", borderRadius: 6, border: "none", background: "#6B21C8", color: "white", fontSize: 12, cursor: "pointer", fontFamily: "inherit" }}>儲存</button>
                    </div>
                  </div>
                )}
                {items.length === 0 && addingItem !== cat.id && (
                  <div style={{ fontSize: 12, color: "#9ca3af" }}>此分類還沒有快捷鍵</div>
                )}
                {items.map(item => (
                  <div key={item.id} draggable onDragStart={() => setDragItm({ catId: cat.id, itemId: item.id })} onDragOver={e => { e.preventDefault(); setOverItm(item.id); }} onDrop={e => dropItm(e, cat.id, item.id)} onDragEnd={() => { setDragItm(null); setOverItm(null); }} style={{ background: "white", border: overItm === item.id && dragItm?.catId === cat.id ? "2px solid #6B21C8" : "1px solid #e5e7eb", borderRadius: 8, padding: "10px 12px", opacity: dragItm?.itemId === item.id ? 0.4 : 1, cursor: "grab" }}>
                    {editItem?.catId === cat.id && editItem?.itemId === item.id ? (
                      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                        <input style={INP} value={form.title} onChange={e => setForm(p => ({ ...p, title: e.target.value }))} autoFocus />
                        <textarea style={{ ...INP, resize: "none", minHeight: 72 }} value={form.text} onChange={e => setForm(p => ({ ...p, text: e.target.value }))} />
                        <div>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                            <span style={{ fontSize: 12, color: "#374151", fontWeight: 500 }}>圖片</span>
                            <button onClick={() => editImgRef.current?.click()} style={{ padding: "3px 10px", borderRadius: 6, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>+ 新增</button>
                            <input ref={editImgRef} type="file" accept="image/*" multiple style={{ display: "none" }} onChange={e => { addImgs(e.target.files); e.target.value = ""; }} />
                          </div>
                          {form.images.length > 0 && (
                            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                              {form.images.map((img, idx) => (
                                <div key={img.id} style={{ position: "relative" }}>
                                  <img src={img.dataUrl || img.preview} alt="" style={{ width: 60, height: 60, objectFit: "cover", borderRadius: 6, border: "1px solid #e5e7eb" }} />
                                  <button onClick={() => setForm(p => ({ ...p, images: p.images.filter((_, i) => i !== idx) }))} style={{ position: "absolute", top: -5, right: -5, width: 16, height: 16, borderRadius: "50%", background: "#ef4444", border: "none", color: "white", fontSize: 9, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                        <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                          <button onClick={() => { setEditItem(null); setForm({ title: "", text: "", images: [] }); }} style={{ padding: "4px 10px", borderRadius: 6, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>取消</button>
                          <button onClick={() => updItm(cat.id, item.id)} style={{ padding: "4px 10px", borderRadius: 6, border: "none", background: "#6B21C8", color: "white", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>儲存</button>
                        </div>
                      </div>
                    ) : (
                      <div>
                        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 10 }}>
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontSize: 12, fontWeight: 600, color: "#374151", marginBottom: 4 }}>{item.title}</div>
                            {item.text && <div style={{ fontSize: 12, color: "#6b7280", lineHeight: 1.6, whiteSpace: "pre-wrap" }}>{item.text}</div>}
                          </div>
                          <div style={{ display: "flex", gap: 5, flexShrink: 0 }}>
                            {item.text && <CopyText text={item.text} label="複製文字" />}
                            <button onClick={() => { setEditItem({ catId: cat.id, itemId: item.id }); setForm({ title: item.title, text: item.text || "", images: item.images || [] }); }} style={{ padding: "4px 7px", borderRadius: 6, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", fontSize: 13, cursor: "pointer", fontFamily: "inherit" }} title="編輯">✎</button>
                            <button onClick={() => delItm(cat.id, item.id)} style={{ padding: "4px 7px", borderRadius: 6, border: "1px solid #fecaca", background: "white", color: "#ef4444", fontSize: 13, cursor: "pointer", fontFamily: "inherit" }} title="刪除">✕</button>
                          </div>
                        </div>
                        <ImgGrid images={item.images || []} onPreview={setLightbox} />
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })}
      {addingCat ? (
        <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
          <input style={{ ...INP, flex: 1 }} placeholder="輸入分類名稱" value={newCatName} onChange={e => setNewCatName(e.target.value)} onKeyDown={e => { if (e.key === "Enter") addCat(); if (e.key === "Escape") setAddingCat(false); }} autoFocus />
          <button onClick={addCat} style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: "#6B21C8", color: "white", fontSize: 12, cursor: "pointer", fontFamily: "inherit", whiteSpace: "nowrap" }}>建立</button>
          <button onClick={() => setAddingCat(false)} style={{ padding: "8px 10px", borderRadius: 8, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", fontSize: 12, cursor: "pointer", fontFamily: "inherit" }}>取消</button>
        </div>
      ) : (
        <button onClick={() => setAddingCat(true)} style={{ padding: "8px 0", borderRadius: 8, border: "1px dashed #d1d5db", background: "white", color: "#6b7280", fontSize: 12, cursor: "pointer", fontFamily: "inherit", marginTop: 4 }}>+ 新增分類</button>
      )}
    </div>
  );
}

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const STAGES = ["觸發需求", "建立信任", "推服務", "報價切入", "促成交易", "報價後追蹤"];
const WEIGHTS = [
  { id: "low", label: "低", color: "#ef4444", bg: "#fef2f2", border: "#fecaca" },
  { id: "mid", label: "中", color: "#f59e0b", bg: "#fffbeb", border: "#fde68a" },
  { id: "high", label: "高", color: "#10b981", bg: "#f0fdf4", border: "#a7f3d0" },
];
// Weight display name
const WEIGHT_NAME = "成交率";
const GS = {
  A: { color: "#10b981", bg: "#f0fdf4", border: "#a7f3d0", label: "A級｜一年方案" },
  B: { color: "#f59e0b", bg: "#fffbeb", border: "#fde68a", label: "B級｜一季方案" },
  C: { color: "#ef4444", bg: "#fef2f2", border: "#fecaca", label: "C級｜結束" },
};

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  // ── State ──
  const [clients, setClients] = useState({});
  const [order, setOrder] = useState([]);
  const [asstMsgs, setAsstMsgs] = useState([]);
  const [sc, setSc] = useState({ categories: [], items: {} });
  const [activeId, setActiveId] = useState("__asst__");
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [modal, setModal] = useState(null);
  const [mForm, setMForm] = useState({});
  const [sideOpen, setSideOpen] = useState(true);
  const [scOpen, setScOpen] = useState(false);
  const [scSearch, setScSearch] = useState("");
  const [sideTab, setSideTab] = useState("visitor");
  const [selectedProspects, setSelectedProspects] = useState(new Set());
  const [exportMode, setExportMode] = useState(false);
  const [weightFilter, setWeightFilter] = useState(null);
  const [pinnedIds, setPinnedIds] = useState(new Set());
  const [loaded, setLoaded] = useState(false);
  const [trash, setTrash] = useState([]);
  const [showTrash, setShowTrash] = useState(false);
  const [trashTab, setTrashTab] = useState("visitor");
  const [noteOpen, setNoteOpen] = useState(true);
  const [editCId, setEditCId] = useState(null);
  const [editCName, setEditCName] = useState("");
  const [imgPrev, setImgPrev] = useState(null);
  const [imgB64, setImgB64] = useState(null);
  const [dragging, setDragging] = useState(false);
  const [dragCli, setDragCli] = useState(null);
  const [overCli, setOverCli] = useState(null);
  const [ctxMenu, setCtxMenu] = useState(null);
  const chatRef = useRef(null);
  const asstFileRef = useRef(null);
  const cliFileRef = useRef(null);
  const lpTimer = useRef(null);
  const importRef = useRef(null);

  // ── Persist to Supabase ──
  const saveTimer = useRef({});
  function debounceSave(key, fn, delay = 1000) {
    clearTimeout(saveTimer.current[key]);
    saveTimer.current[key] = setTimeout(fn, delay);
  }
  useEffect(() => { if (!loaded) return; Object.values(clients).forEach(c => debounceSave('c_' + c.id, () => dbUpsertClient(c))); }, [clients, loaded]);
  useEffect(() => { if (!loaded) return; debounceSave('order', () => dbSetMeta('order', order)); }, [order, loaded]);
  useEffect(() => { if (!loaded) return; debounceSave('sc', () => dbSetShortcuts(sc)); }, [sc, loaded]);
  useEffect(() => { if (!loaded) return; debounceSave('pinned', () => dbSetMeta('pinnedIds', [...pinnedIds])); }, [pinnedIds, loaded]);
  useEffect(() => {
    if (!loaded) return;
    const now = Date.now();
    const clean = trash.filter(t => now - t.deletedAt < TRASH_TTL);
    debounceSave('trash', () => dbSetTrash(clean));
  }, [trash, loaded]);
  useEffect(() => {
    async function load() {
      try {
        const [c, orderVal, pinnedVal, s, tr] = await Promise.all([
          dbGetClients(),
          dbGetMeta('order'),
          dbGetMeta('pinnedIds'),
          dbGetShortcuts(),
          dbGetTrash(),
        ]);
        if (c) setClients(c);
        if (orderVal) setOrder(orderVal);
        if (pinnedVal) setPinnedIds(new Set(pinnedVal));
        if (s) setSc(s);
        if (tr) {
          const now = Date.now();
          setTrash(tr.filter(t => now - t.deletedAt < TRASH_TTL));
        }
      } catch(e) { console.error('Load error:', e); }
      setLoaded(true);
    }
    load();
  }, []);
  useEffect(() => { if (chatRef.current) chatRef.current.scrollTop = chatRef.current.scrollHeight; }, [clients, asstMsgs, activeId, loading]);

  // ── Client helpers ──
  function updC(id, patch) { setClients(p => ({ ...p, [id]: { ...p[id], ...patch } })); }
  function addMsg(id, msg) { setClients(p => ({ ...p, [id]: { ...p[id], messages: [...(p[id].messages || []), msg] } })); }
  function replaceLastAI(id, msg) {
    setClients(p => {
      const msgs = [...(p[id].messages || [])];
      for (let i = msgs.length - 1; i >= 0; i--) { if (msgs[i].type === "ai") { msgs[i] = msg; break; } }
      return { ...p, [id]: { ...p[id], messages: msgs } };
    });
  }
  function delClient(id) {
    const c = clients[id]; if (!c) return;
    setTrash(p => [...p, { ...c, deletedAt: Date.now(), trashType: "client" }]);
    setClients(p => { const n = { ...p }; delete n[id]; return n; });
    setOrder(p => p.filter(x => x !== id));
    setPinnedIds(p => { const n = new Set(p); n.delete(id); return n; });
    dbDeleteClient(id);
    if (activeId === id) setActiveId("__asst__");
    setCtxMenu(null);
  }
  function restoreClient(item) {
    const { deletedAt, trashType, ...client } = item;
    setClients(p => ({ ...p, [client.id]: client }));
    setOrder(p => [...p, client.id]);
    setTrash(p => p.filter(t => t.id !== client.id || t.deletedAt !== deletedAt));
  }
  function permDelClient(item) {
    setTrash(p => p.filter(t => !(t.id === item.id && t.deletedAt === item.deletedAt)));
  }
  function createVisitor(name) {
    const id = "v" + Date.now();
    setClients(p => ({ ...p, [id]: { id, type: "visitor", name, grade: "auto", stage: "建立信任", messages: [], createdAt: new Date().toLocaleDateString("zh-TW") } }));
    setOrder(p => [...p, id]);
    setActiveId(id); setModal(null); setMForm({});
  }
  function createProspect(name, account, note, weight) {
    const id = "p" + Date.now();
    setClients(p => ({ ...p, [id]: { id, type: "prospect", name, account, note, weight: weight || "mid", grade: "auto", stage: "建立信任", messages: [], createdAt: new Date().toLocaleDateString("zh-TW") } }));
    setOrder(p => [...p, id]);
    setActiveId(id); setModal(null); setMForm({});
  }
  function upgradeVisitor(id) {
    if (!mForm.name || !mForm.account) return;
    updC(id, { type: "prospect", name: mForm.name, account: mForm.account, note: mForm.note || "", weight: mForm.weight || "mid" });
    setModal(null); setMForm({});
  }

  // ── Drag client ──
  const orderedIds = [...order.filter(id => clients[id]), ...Object.keys(clients).filter(id => !order.includes(id))];
  const clientList = orderedIds.map(id => clients[id]).filter(Boolean);
  function dropCli(e, toId) {
    e.preventDefault();
    if (!dragCli || dragCli === toId) { setDragCli(null); setOverCli(null); return; }
    setOrder(prev => {
      const o = [...orderedIds];
      const fi = o.indexOf(dragCli), ti = o.indexOf(toId);
      if (fi < 0 || ti < 0) return prev;
      const [m] = o.splice(fi, 1); o.splice(ti, 0, m); return o;
    });
    setDragCli(null); setOverCli(null);
  }
  function pinClient(id) {
    setOrder(prev => { const o = orderedIds.filter(x => x !== id); return [id, ...o]; });
    setPinnedIds(p => { const n = new Set(p); n.add(id); return n; });
    setCtxMenu(null);
  }
  function unpinClient(id) {
    setPinnedIds(p => { const n = new Set(p); n.delete(id); return n; });
    setCtxMenu(null);
  }
  function openCtx(e, id) { e.preventDefault(); setCtxMenu({ x: e.clientX, y: e.clientY, id }); }
  function startLP(id) { lpTimer.current = setTimeout(() => setCtxMenu({ x: 60, y: 200, id }), 600); }
  function cancelLP() { clearTimeout(lpTimer.current); }

  // ── File handlers ──
  function handleFile(e, forAsst) {
    const f = e.target.files?.[0]; if (!f) return;
    readImgFile(f, b => { setImgB64(b); setImgPrev(b.preview); });
    e.target.value = "";
  }
  function handleDrop(e) {
    e.preventDefault(); setDragging(false);
    const f = e.dataTransfer.files?.[0];
    if (f) readImgFile(f, b => { setImgB64(b); setImgPrev(b.preview); });
  }

  // ── API calls ──
  async function sendAsst(text, img) {
    setLoading(true);
    const prospects = Object.values(clients).filter(c => c.type === "prospect");
    const visitors = Object.values(clients).filter(c => c.type === "visitor");
    let ctx = "";
    if (prospects.length) ctx += "\n\n【潛在客戶】\n" + prospects.map(p => `- ${p.name}（${p.account}）${p.weight === "high" ? "高" : p.weight === "mid" ? "中" : "低"}質量：${p.note || ""}`).join("\n");
    if (visitors.length) ctx += "\n\n【遊客對話】\n" + visitors.map(v => `- ${v.name}，階段：${v.stage || "觸發需求"}，等級：${v.grade || "未判斷"}`).join("\n");
    const msgContent = img ? imgMsg(img, text || "請分析截圖對話，判斷客戶狀況，建議下一步。") : text;
    const hist = asstMsgs.slice(-8).map(m => ({ role: m.role, content: m.content }));
    try {
      const reply = await callAPI(ASST_SYS + ctx, [...hist, { role: "user", content: msgContent }], 1500);
      const time = new Date().toLocaleTimeString("zh-TW", { hour: "2-digit", minute: "2-digit" });
      setAsstMsgs(p => [...p, { role: "user", content: typeof msgContent === "string" ? msgContent : "[截圖]", display: text || "（截圖）", imgSrc: img ? img.preview : null, time }, { role: "assistant", content: reply, display: reply, time }]);
    } catch {
      setAsstMsgs(p => [...p, { role: "assistant", content: "連線錯誤，請重試。", display: "連線錯誤，請重試。", time: "" }]);
    }
    setLoading(false);
  }

  async function sendClient(clientId, text, img, isRegen = false, isAnalyse = false) {
    setLoading(true);
    const client = clients[clientId];
    const grade = client.grade || "auto";
    const validStages = ["建立信任", "觸發需求", "報價切入"]; const stage = validStages.includes(client.stage) ? client.stage : "建立信任";
    const gradeInstr = grade === "auto" ? "若需判斷等級，第一行只寫 GRADE:A/B/C，換行再寫回覆。" : `客戶等級：${grade}級。`;
    const hist = (client.messages || []).slice(-10).map(m => ({ role: m.type === "user" ? "user" : "assistant", content: m.text || "" }));
    const time = new Date().toLocaleTimeString("zh-TW", { hour: "2-digit", minute: "2-digit" });
    let system, userContent;
    if (img && isAnalyse) { system = IMG_ANALYSE_SYS; userContent = imgMsg(img, text || "請分析截圖。"); }
    else if (img) { system = IMG_REPLY_SYS + `\n當前階段：${stage}\n${gradeInstr}`; userContent = imgMsg(img, text || "請識別截圖並給出回覆。"); }
    else if (isAnalyse) { system = ANALYSE_SYS + `\n當前階段：${stage}`; userContent = text; }
    else { system = REPLY_SYS + `\n當前階段：${stage}\n${gradeInstr}`; userContent = text; }
    try {
      const raw = await callAPI(system, [...(img || isAnalyse ? [] : hist), { role: "user", content: userContent }]);
      let reply = raw, detectedGrade = grade === "auto" ? "" : grade;
      const gm = raw.match(/^GRADE:\s*([ABC])\s*\n?/i);
      if (gm) { detectedGrade = gm[1].toUpperCase(); reply = raw.slice(gm[0].length).trim(); if (grade === "auto") updC(clientId, { grade: detectedGrade }); }
      reply = reply.replace(/^GRADE:\s*[ABC]\s*$/gim, "").trim();
      const aiMsg = { type: "ai", text: reply, grade: detectedGrade, time, isAnalyse };
      if (isRegen) replaceLastAI(clientId, aiMsg); else addMsg(clientId, aiMsg);
    } catch {
      const err = { type: "ai", text: "連線錯誤，請重試。", grade: "", time, isAnalyse: false };
      if (isRegen) replaceLastAI(clientId, err); else addMsg(clientId, err);
    }
    setLoading(false);
  }

  async function send() {
    const text = input.trim();
    if ((!text && !imgB64) || loading) return;
    const img = imgB64;
    setInput(""); setImgPrev(null); setImgB64(null);
    if (activeId === "__asst__") { await sendAsst(text, img); }
    else {
      const display = text || "（截圖）";
      addMsg(activeId, { type: "user", text: display, imgSrc: img ? img.preview : null, time: new Date().toLocaleTimeString("zh-TW", { hour: "2-digit", minute: "2-digit" }) });
      await sendClient(activeId, text, img);
    }
  }
  function handleKey(e) { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } }

  // ── Derived ──
  const activeClient = activeId !== "__asst__" ? clients[activeId] : null;
  const messages = activeClient?.messages || [];
  const allScItems = (sc.categories || []).flatMap(cat => (sc.items?.[cat.id] || []).map(item => ({ ...item, catName: cat.name })));
  const filteredSc = scSearch.trim() ? allScItems.filter(i => i.title.includes(scSearch) || i.text?.includes(scSearch) || i.catName.includes(scSearch)) : allScItems;

  // ── Excel export ──
  function exportProspects() {
    exportSelectedProspects();
    return;
    const XL = xlsxRef.current; // kept for reference
    const prospects = clientList.filter(c => c.type === "prospect");
    if (prospects.length === 0) { alert("目前沒有潛在客戶資料可匯出"); return; }
    const gradeLabel = { A: "A級", B: "B級", C: "C級", auto: "未判斷" };
    const weightLabel = { low: "低", mid: "中", high: "高" };
    const rows = prospects.map(c => ({
      "姓名": c.name || "",
      "IG帳號": c.account || "",
      "說明": c.note || "",
      "成交率": weightLabel[c.weight] || "中",
      "客戶等級": gradeLabel[c.grade] || "未判斷",
      "目前階段": c.stage || "觸發需求",
      "建立日期": c.createdAt || "",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    // Set column widths
    ws["!cols"] = [{ wch: 12 }, { wch: 20 }, { wch: 30 }, { wch: 10 }, { wch: 10 }, { wch: 14 }, { wch: 12 }];
    // Style header row (bold)
    const range = XLSX.utils.decode_range(ws["!ref"]);
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c: C });
      if (!ws[addr]) continue;
      ws[addr].s = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "6B21C8" } } };
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "潛在客戶");
    XLSX.writeFile(wb, `空書道_潛在客戶_${new Date().toLocaleDateString("zh-TW").replace(/\//g,"-")}.xlsx`);
  }

  function exportSelectedProspects() {
    const XL = xlsxRef.current;
    if (!XL) { alert("Excel 功能載入中，請稍後再試"); return; }
    const ids = selectedProspects.size > 0 ? [...selectedProspects] : clientList.filter(c => c.type === "prospect").map(c => c.id);
    const prospects = ids.map(id => clients[id]).filter(Boolean);
    if (prospects.length === 0) { alert("沒有可匯出的資料"); return; }
    const gradeLabel = { A: "A級", B: "B級", C: "C級", auto: "未判斷" };
    const weightLabel = { low: "低", mid: "中", high: "高" };
    const rows = prospects.map(c => ({
      "姓名": c.name || "", "IG帳號": c.account || "", "說明": c.note || "",
      "成交率": weightLabel[c.weight] || "中", "客戶等級": gradeLabel[c.grade] || "未判斷",
      "目前階段": c.stage || "觸發需求", "建立日期": c.createdAt || "",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 12 }, { wch: 20 }, { wch: 30 }, { wch: 10 }, { wch: 10 }, { wch: 14 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "潛在客戶");
    XLSX.writeFile(wb, `空書道_潛在客戶_${new Date().toLocaleDateString("zh-TW").replace(/\//g,"-")}.xlsx`);
    setSelectedProspects(new Set());
  }

  function importProspects(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        const weightMap = { "低": "low", "中": "mid", "高": "high" };
        const gradeMap = { "A級": "A", "B級": "B", "C級": "C", "未判斷": "auto" };
        let imported = 0;
        rows.forEach(row => {
          const name = row["姓名"] || row["名稱"] || "";
          const account = row["IG帳號"] || row["帳號"] || "";
          if (!name) return;
          const id = "p" + Date.now() + Math.random();
          const newClient = {
            id, type: "prospect",
            name: String(name),
            account: String(account || ""),
            note: String(row["說明"] || ""),
            weight: weightMap[row["成交率"]] || "mid",
            grade: gradeMap[row["客戶等級"]] || "auto",
            stage: row["目前階段"] || "觸發需求",
            messages: [],
            createdAt: row["建立日期"] || new Date().toLocaleDateString("zh-TW"),
          };
          setClients(p => ({ ...p, [id]: newClient }));
          setOrder(p => [...p, id]);
          imported++;
        });
        alert(`成功匯入 ${imported} 筆潛在客戶資料`);
      } catch (e) {
        alert("匯入失敗，請確認檔案格式正確（需為匯出的Excel格式）");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ── Style helpers ──
  const SIDE_ITEM = (active) => ({ padding: "10px 14px", cursor: "pointer", transition: "all 0.15s", borderLeft: active ? "3px solid #6B21C8" : "3px solid transparent", background: active ? "#f3f0ff" : "transparent", display: "flex", alignItems: "center", justifyContent: "space-between" });
  const BTN = (primary) => ({ padding: "8px 16px", borderRadius: 8, fontSize: 13, cursor: "pointer", fontFamily: "inherit", fontWeight: 500, border: primary ? "none" : "1px solid #e5e7eb", background: primary ? "#6B21C8" : "white", color: primary ? "white" : "#374151" });

  // ─── RENDER ───────────────────────────────────────────────────────────────
  if (!loaded) return (
    <div style={{ fontFamily: "'Noto Sans TC','PingFang TC',sans-serif", display: "flex", alignItems: "center", justifyContent: "center", minHeight: "100vh", flexDirection: "column", gap: 16, background: "white", color: "#6B21C8" }}>
      <div style={{ width: 40, height: 40, borderRadius: 12, background: "linear-gradient(135deg,#6B21C8,#9B59F5)", display: "flex", alignItems: "center", justifyContent: "center", color: "white", fontSize: 20, fontWeight: 700 }}>空</div>
      <div style={{ fontSize: 14, color: "#6b7280" }}>載入中，正在同步數據...</div>
    </div>
  );

  return (
    <div style={{ fontFamily: "'Noto Sans TC','PingFang TC',sans-serif", fontSize: 13, background: "white", minHeight: "100vh", display: "flex", flexDirection: "column", color: "#111" }}>

      {/* HEADER */}
      <div style={{ background: "white", borderBottom: "1px solid #e5e7eb", padding: "12px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 50 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <button onClick={() => setSideOpen(v => !v)} style={{ width: 28, height: 28, borderRadius: 7, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11 }} title={sideOpen ? "隱藏側欄" : "顯示側欄"}>
            {sideOpen ? "◀" : "▶"}
          </button>
          <div style={{ width: 28, height: 28, borderRadius: 8, background: "linear-gradient(135deg,#6B21C8,#9B59F5)", display: "flex", alignItems: "center", justifyContent: "center", color: "white", fontSize: 13, fontWeight: 700 }}>空</div>
          <span style={{ fontWeight: 700, fontSize: 15, letterSpacing: 1 }}>空書道 AI 客服助手</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button onClick={() => setScOpen(v => !v)} style={{ padding: "5px 12px", borderRadius: 8, border: "1px solid #e5e7eb", background: "white", color: "#374151", fontSize: 12, cursor: "pointer", fontFamily: "inherit", fontWeight: 500, display: "flex", alignItems: "center", gap: 5 }}>⚡ 回覆快捷鍵</button>
          <button onClick={() => setShowTrash(true)} style={{ padding: "5px 12px", borderRadius: 8, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", fontSize: 12, cursor: "pointer", fontFamily: "inherit", display: "flex", alignItems: "center", gap: 5 }} title="垃圾桶">
            🗑 垃圾桶{trash.length > 0 && <span style={{ background: "#ef4444", color: "white", borderRadius: "50%", width: 16, height: 16, fontSize: 10, display: "flex", alignItems: "center", justifyContent: "center" }}>{trash.length}</span>}
          </button>
          <div style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11, color: "#6b7280" }}>
            <div style={{ width: 6, height: 6, borderRadius: "50%", background: "#10b981" }} />運行中
          </div>
        </div>
      </div>

      <div style={{ flex: 1, display: "flex", overflow: "hidden", maxHeight: "calc(100vh - 53px)" }}>

        {/* SIDEBAR */}
        <div style={{ width: sideOpen ? 220 : 0, minWidth: sideOpen ? 220 : 0, overflow: "hidden", background: "#fafafa", borderRight: sideOpen ? "1px solid #e5e7eb" : "none", display: "flex", flexDirection: "column", flexShrink: 0, transition: "width 0.25s, min-width 0.25s" }}>
          <div style={{ width: 220, display: "flex", flexDirection: "column", flex: 1, minHeight: 0 }}>
            {/* AI Assistant entry */}
            <div style={SIDE_ITEM(activeId === "__asst__")} onClick={() => setActiveId("__asst__")}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <div style={{ width: 26, height: 26, borderRadius: 8, fontSize: 12, background: activeId === "__asst__" ? "#6B21C8" : "#e5e7eb", color: activeId === "__asst__" ? "white" : "#6b7280", display: "flex", alignItems: "center", justifyContent: "center" }}>✦</div>
                <div>
                  <div style={{ fontSize: 12, fontWeight: 600, color: activeId === "__asst__" ? "#6B21C8" : "#111" }}>AI 客服助理</div>
                  <div style={{ fontSize: 10, color: "#9ca3af" }}>策略討論 / 截圖分析</div>
                </div>
              </div>
            </div>
            <div style={{ height: 1, background: "#e5e7eb", margin: "4px 0" }} />

            {/* Tab switcher: 遊客 / 潛在客戶 */}
            <div style={{ display: "flex", borderBottom: "1px solid #e5e7eb" }}>
              <button onClick={() => setSideTab("visitor")} style={{ flex: 1, padding: "8px 0", fontSize: 11, fontWeight: sideTab === "visitor" ? 700 : 400, cursor: "pointer", fontFamily: "inherit", border: "none", background: "transparent", color: sideTab === "visitor" ? "#6B21C8" : "#6b7280", borderBottom: sideTab === "visitor" ? "2px solid #6B21C8" : "2px solid transparent", transition: "all 0.15s" }}>
                遊客（{clientList.filter(c => c.type === "visitor").length}）
              </button>
              <button onClick={() => setSideTab("prospect")} style={{ flex: 1, padding: "8px 0", fontSize: 11, fontWeight: sideTab === "prospect" ? 700 : 400, cursor: "pointer", fontFamily: "inherit", border: "none", background: "transparent", color: sideTab === "prospect" ? "#6B21C8" : "#6b7280", borderBottom: sideTab === "prospect" ? "2px solid #6B21C8" : "2px solid transparent", transition: "all 0.15s" }}>
                潛在（{clientList.filter(c => c.type === "prospect").length}）
              </button>
            </div>

            {/* VISITOR TAB */}
            {sideTab === "visitor" && (
              <>
                {/* Add + header */}
                <div style={{ padding: "8px 12px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <span style={{ fontSize: 10, color: "#9ca3af", fontWeight: 600, letterSpacing: 1 }}>遊客對話</span>
                  <button onClick={() => setModal("visitor")} style={{ padding: "3px 10px", borderRadius: 6, border: "1px solid #6B21C8", background: "#6B21C8", color: "white", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>+ 新增</button>
                </div>
                <div style={{ flex: 1, overflowY: "auto" }}>
                  {clientList.filter(c => c.type === "visitor").length === 0 && (
                    <div style={{ padding: "16px 14px", fontSize: 12, color: "#9ca3af", textAlign: "center" }}>還沒有遊客對話</div>
                  )}
                  {clientList.filter(c => c.type === "visitor").map(c => (
                    <div key={c.id} draggable onDragStart={() => setDragCli(c.id)} onDragOver={e => { e.preventDefault(); setOverCli(c.id); }} onDrop={e => dropCli(e, c.id)} onDragEnd={() => { setDragCli(null); setOverCli(null); }} onContextMenu={e => openCtx(e, c.id)} onTouchStart={() => startLP(c.id)} onTouchEnd={cancelLP} onTouchMove={cancelLP}
                      style={{ ...SIDE_ITEM(activeId === c.id), opacity: dragCli === c.id ? 0.4 : 1, borderTop: overCli === c.id ? "2px solid #6B21C8" : undefined, cursor: "grab" }}
                      onClick={() => setActiveId(c.id)}>
                      <div style={{ flexShrink: 0, color: "#d1d5db", fontSize: 11, paddingRight: 4 }}>⠿</div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        {editCId === c.id ? (
                          <input autoFocus value={editCName} onChange={e => setEditCName(e.target.value)} onKeyDown={e => { e.stopPropagation(); if (e.key === "Enter" && editCName.trim()) { updC(c.id, { name: editCName.trim() }); setEditCId(null); } if (e.key === "Escape") setEditCId(null); }} onClick={e => e.stopPropagation()} onBlur={() => { if (editCName.trim()) updC(c.id, { name: editCName.trim() }); setEditCId(null); }} style={{ width: "100%", fontSize: 12, fontWeight: 500, border: "1px solid #6B21C8", borderRadius: 5, padding: "2px 6px", outline: "none", fontFamily: "inherit", color: "#6B21C8" }} />
                        ) : (
                          <div style={{ fontSize: 12, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", color: activeId === c.id ? "#6B21C8" : "#111" }}>
                            {pinnedIds.has(c.id) && <span style={{ fontSize: 9, marginRight: 3 }}>📌</span>}{c.name}
                          </div>
                        )}
                        <div style={{ display: "flex", gap: 4, marginTop: 2, alignItems: "center" }}>
                          <span style={{ fontSize: 10, color: "#9ca3af" }}>{c.stage || "觸發需求"}</span>
                          {c.grade && c.grade !== "auto" && GS[c.grade] && <Tag label={c.grade + "級"} {...GS[c.grade]} />}
                        </div>
                      </div>
                      <div style={{ display: "flex", gap: 2, flexShrink: 0 }}>
                        <button onClick={e => { e.stopPropagation(); setEditCId(c.id); setEditCName(c.name); }} style={{ background: "none", border: "none", color: "#9ca3af", cursor: "pointer", fontSize: 14, padding: "2px 4px", letterSpacing: 1 }}>⋯</button>
                        <button onClick={e => { e.stopPropagation(); delClient(c.id); }} style={{ background: "none", border: "none", color: "#d1d5db", cursor: "pointer", fontSize: 12, padding: "2px 4px" }}>✕</button>
                      </div>
                    </div>
                  ))}
                </div>
              </>
            )}

            {/* PROSPECT TAB */}
            {sideTab === "prospect" && (
              <>
                {/* Toolbar: add + export + import */}
                <div style={{ padding: "8px 12px", display: "flex", alignItems: "center", gap: 5, borderBottom: "1px solid #f3f4f6" }}>
                  <button onClick={() => setModal("prospect")} style={{ flex: 1, padding: "4px 0", borderRadius: 6, border: "none", background: "#6B21C8", color: "white", fontSize: 11, cursor: "pointer", fontFamily: "inherit", fontWeight: 500 }}>+ 新增</button>
                  <button onClick={() => { setExportMode(v => !v); setSelectedProspects(new Set()); }} style={{ padding: "4px 8px", borderRadius: 6, border: exportMode ? "1px solid #6B21C8" : "1px solid #e5e7eb", background: exportMode ? "#f3f0ff" : "white", color: exportMode ? "#6B21C8" : "#374151", fontSize: 11, cursor: "pointer", fontFamily: "inherit", fontWeight: exportMode ? 600 : 400 }}>↓ 匯出</button>
                  <button onClick={() => importRef.current?.click()} style={{ padding: "4px 8px", borderRadius: 6, border: "1px solid #e5e7eb", background: "white", color: "#374151", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>↑ 匯入</button>
                  <input ref={importRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => { importProspects(e.target.files?.[0]); e.target.value = ""; }} />
                </div>
                {/* Weight filter bar - 成交率篩選 */}
                <div style={{ padding: "5px 12px", display: "flex", gap: 4, borderBottom: "1px solid #f3f4f6" }}>
                  <span style={{ fontSize: 10, color: "#9ca3af", alignSelf: "center", marginRight: 2 }}>成交率</span>
                  {WEIGHTS.map(w => (
                    <button key={w.id} onClick={() => setWeightFilter(p => p === w.id ? null : w.id)} style={{ padding: "2px 8px", borderRadius: 10, fontSize: 10, cursor: "pointer", fontFamily: "inherit", border: `1px solid ${weightFilter === w.id ? w.border : "#e5e7eb"}`, background: weightFilter === w.id ? w.bg : "white", color: weightFilter === w.id ? w.color : "#9ca3af", fontWeight: weightFilter === w.id ? 700 : 400, transition: "all 0.15s" }}>{w.label}</button>
                  ))}
                  {weightFilter && <button onClick={() => setWeightFilter(null)} style={{ padding: "2px 6px", borderRadius: 10, fontSize: 10, cursor: "pointer", fontFamily: "inherit", border: "none", background: "transparent", color: "#9ca3af" }}>✕</button>}
                </div>
                {/* Export mode: select-all bar */}
                {exportMode && (
                  <div style={{ padding: "6px 12px", background: "#f3f0ff", borderBottom: "1px solid #e5e7eb", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 7, cursor: "pointer" }} onClick={() => {
                      const all = clientList.filter(c => c.type === "prospect").map(c => c.id);
                      setSelectedProspects(p => p.size === all.length ? new Set() : new Set(all));
                    }}>
                      {(() => {
                        const all = clientList.filter(c => c.type === "prospect");
                        const allSel = all.length > 0 && selectedProspects.size === all.length;
                        return (
                          <div style={{ width: 14, height: 14, borderRadius: 3, border: allSel ? "none" : "1px solid #6B21C8", background: allSel ? "#6B21C8" : "white", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                            {allSel && <span style={{ color: "white", fontSize: 9, fontWeight: 700 }}>✓</span>}
                          </div>
                        );
                      })()}
                      <span style={{ fontSize: 11, color: "#6B21C8", fontWeight: 500 }}>
                        {selectedProspects.size > 0 ? `已選 ${selectedProspects.size} 筆` : "全選"}
                      </span>
                    </div>
                    <div style={{ display: "flex", gap: 5 }}>
                      <button onClick={() => { exportSelectedProspects(); setExportMode(false); }} style={{ padding: "3px 8px", borderRadius: 5, border: "none", background: "#6B21C8", color: "white", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>
                        {selectedProspects.size > 0 ? `匯出 ${selectedProspects.size} 筆` : "匯出全部"}
                      </button>
                      <button onClick={() => { setExportMode(false); setSelectedProspects(new Set()); }} style={{ padding: "3px 8px", borderRadius: 5, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>取消</button>
                    </div>
                  </div>
                )}
                <div style={{ flex: 1, overflowY: "auto" }}>
                  {clientList.filter(c => c.type === "prospect" && (!weightFilter || c.weight === weightFilter)).length === 0 && (
                    <div style={{ padding: "16px 14px", fontSize: 12, color: "#9ca3af", textAlign: "center" }}>{weightFilter ? `沒有${WEIGHTS.find(w=>w.id===weightFilter)?.label}成交率的客戶` : "還沒有潛在客戶"}</div>
                  )}
                  {clientList.filter(c => c.type === "prospect" && (!weightFilter || c.weight === weightFilter)).map(c => {
                    const w = WEIGHTS.find(w => w.id === c.weight) || WEIGHTS[1];
                    const isSel = selectedProspects.has(c.id);
                    return (
                      <div key={c.id} draggable onDragStart={() => setDragCli(c.id)} onDragOver={e => { e.preventDefault(); setOverCli(c.id); }} onDrop={e => dropCli(e, c.id)} onDragEnd={() => { setDragCli(null); setOverCli(null); }} onContextMenu={e => openCtx(e, c.id)} onTouchStart={() => startLP(c.id)} onTouchEnd={cancelLP} onTouchMove={cancelLP}
                        style={{ ...SIDE_ITEM(activeId === c.id), opacity: dragCli === c.id ? 0.4 : 1, borderTop: overCli === c.id ? "2px solid #6B21C8" : undefined, cursor: "grab", background: (exportMode && isSel) ? "#f3f0ff" : activeId === c.id ? "#f3f0ff" : "transparent" }}
                        onClick={() => setActiveId(c.id)}>
                        {/* Checkbox - only visible in exportMode */}
                        {exportMode && (
                          <div onClick={e => { e.stopPropagation(); setSelectedProspects(p => { const n = new Set(p); n.has(c.id) ? n.delete(c.id) : n.add(c.id); return n; }); }} style={{ width: 14, height: 14, borderRadius: 3, border: isSel ? "none" : "1px solid #d1d5db", background: isSel ? "#6B21C8" : "white", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, marginRight: 4, cursor: "pointer" }}>
                            {isSel && <span style={{ color: "white", fontSize: 9, fontWeight: 700 }}>✓</span>}
                          </div>
                        )}
                        <div style={{ flexShrink: 0, color: "#d1d5db", fontSize: 11, paddingRight: 4 }}>⠿</div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                            {editCId === c.id ? (
                              <input autoFocus value={editCName} onChange={e => setEditCName(e.target.value)} onKeyDown={e => { e.stopPropagation(); if (e.key === "Enter" && editCName.trim()) { updC(c.id, { name: editCName.trim() }); setEditCId(null); } if (e.key === "Escape") setEditCId(null); }} onClick={e => e.stopPropagation()} onBlur={() => { if (editCName.trim()) updC(c.id, { name: editCName.trim() }); setEditCId(null); }} style={{ flex: 1, fontSize: 12, fontWeight: 500, border: "1px solid #6B21C8", borderRadius: 5, padding: "2px 6px", outline: "none", fontFamily: "inherit", color: "#6B21C8" }} />
                            ) : (
                              <div style={{ fontSize: 12, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", color: activeId === c.id ? "#6B21C8" : "#111" }}>
                                {pinnedIds.has(c.id) && <span style={{ fontSize: 9, marginRight: 3 }}>📌</span>}{c.name}
                              </div>
                            )}
                            <Tag label={w.label} {...w} />
                          </div>
                          <div style={{ fontSize: 10, color: "#9ca3af", marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{c.account}</div>
                        </div>
                        <div style={{ display: "flex", gap: 2, flexShrink: 0 }}>
                          <button onClick={e => { e.stopPropagation(); setModal("edit_prospect"); setMForm({ id: c.id, name: c.name, account: c.account || "", note: c.note || "", weight: c.weight || "mid" }); }} style={{ background: "none", border: "none", color: "#9ca3af", cursor: "pointer", fontSize: 14, padding: "2px 4px", letterSpacing: 1 }}>⋯</button>
                          <button onClick={e => { e.stopPropagation(); delClient(c.id); }} style={{ background: "none", border: "none", color: "#d1d5db", cursor: "pointer", fontSize: 12, padding: "2px 4px" }}>✕</button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </>
            )}

            {/* Bottom shortcuts */}
            <div style={{ marginTop: "auto", borderTop: "1px solid #e5e7eb", padding: "8px 12px" }}>
              <button onClick={() => setScOpen(v => !v)} style={{ width: "100%", padding: "7px 0", borderRadius: 8, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", fontSize: 11, cursor: "pointer", fontFamily: "inherit", display: "flex", alignItems: "center", justifyContent: "center", gap: 5 }}>
                ⚡ 快捷鍵管理（{allScItems.length}則）
              </button>
            </div>
          </div>
        </div>

        {/* MAIN CONTENT */}
        <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>

          {/* ASSISTANT VIEW */}
          {activeId === "__asst__" && (
            <>
              <div style={{ padding: "12px 20px", borderBottom: "1px solid #f3f4f6", background: "#fafafa", display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ width: 32, height: 32, borderRadius: 10, background: "linear-gradient(135deg,#6B21C8,#9B59F5)", display: "flex", alignItems: "center", justifyContent: "center", color: "white", fontSize: 14 }}>✦</div>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 14 }}>AI 客服助理</div>
                  <div style={{ fontSize: 11, color: "#6b7280" }}>策略討論、截圖分析、成交建議</div>
                </div>
              </div>
              <div ref={chatRef} style={{ flex: 1, overflowY: "auto", padding: 20, display: "flex", flexDirection: "column", gap: 16 }}>
                {asstMsgs.length === 0 && (
                  <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 12, padding: "40px 0", color: "#9ca3af" }}>
                    <div style={{ fontSize: 36 }}>✦</div>
                    <div style={{ fontSize: 13, textAlign: "center", lineHeight: 1.8 }}>你好，我是你的AI客服助理<br />可以幫你分析截圖、制定成交策略<br />或討論特定客戶的跟進方式</div>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap", justifyContent: "center" }}>
                      {["這個客戶怎麼成交", "客戶說太貴了怎麼辦", "報價後消失怎麼追蹤"].map(q => (
                        <button key={q} onClick={() => setInput(q)} style={{ padding: "5px 12px", borderRadius: 14, border: "1px solid #e5e7eb", background: "white", color: "#374151", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>{q}</button>
                      ))}
                    </div>
                  </div>
                )}
                {asstMsgs.map((m, i) => (
                  <div key={i} style={{ display: "flex", flexDirection: "column", alignItems: m.role === "user" ? "flex-end" : "flex-start", gap: 3 }}>
                    <div style={{ fontSize: 10, color: "#9ca3af" }}>{m.role === "user" ? "Lucy" : "AI助理"} {m.time}</div>
                    {m.imgSrc && <img src={m.imgSrc} alt="截圖" style={{ maxWidth: 240, borderRadius: 10, border: "1px solid #e5e7eb", marginBottom: 4 }} />}
                    <div style={{ maxWidth: "82%", padding: "10px 14px", borderRadius: 14, borderBottomRightRadius: m.role === "user" ? 3 : 14, borderBottomLeftRadius: m.role === "assistant" ? 3 : 14, background: m.role === "user" ? "#6B21C8" : "#f9fafb", border: m.role === "assistant" ? "1px solid #e5e7eb" : "none", color: m.role === "user" ? "white" : "#111", fontSize: 13, lineHeight: 1.8, whiteSpace: "pre-wrap" }}>{m.display}</div>
                    {m.role === "assistant" && <CopyText text={m.display} />}
                  </div>
                ))}
                {loading && activeId === "__asst__" && (
                  <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-start", gap: 3 }}>
                    <div style={{ fontSize: 10, color: "#9ca3af" }}>AI助理</div>
                    <Dots />
                  </div>
                )}
              </div>
              <div style={{ padding: "12px 20px 16px", borderTop: "1px solid #f3f4f6" }}>
                {imgPrev && (
                  <div style={{ position: "relative", display: "inline-block", marginBottom: 8 }}>
                    <img src={imgPrev} alt="預覽" style={{ height: 72, borderRadius: 8, border: "1px solid #e5e7eb" }} />
                    <button onClick={() => { setImgPrev(null); setImgB64(null); }} style={{ position: "absolute", top: -6, right: -6, width: 18, height: 18, borderRadius: "50%", background: "#ef4444", border: "none", color: "white", fontSize: 10, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
                  </div>
                )}
                <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
                  <button onClick={() => asstFileRef.current?.click()} style={{ width: 38, height: 38, borderRadius: 8, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", cursor: "pointer", fontSize: 16, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }} title="上傳截圖">📷</button>
                  <input ref={asstFileRef} type="file" accept="image/*" onChange={e => handleFile(e, true)} style={{ display: "none" }} />
                  <textarea value={input} onChange={e => setInput(e.target.value)} onKeyDown={handleKey} placeholder="輸入問題，或上傳截圖讓AI分析..." style={{ flex: 1, border: "1px solid #e5e7eb", borderRadius: 10, padding: "9px 13px", color: "#111", fontFamily: "inherit", fontSize: 13, lineHeight: 1.6, resize: "none", minHeight: 60, outline: "none", background: "white" }} />
                  <button onClick={send} disabled={loading || (!input.trim() && !imgB64)} style={{ width: 38, height: 38, borderRadius: 8, border: "none", background: loading || (!input.trim() && !imgB64) ? "#e5e7eb" : "#6B21C8", color: "white", cursor: "pointer", flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15 }}>➤</button>
                </div>
              </div>
            </>
          )}

          {/* CLIENT VIEW */}
          {activeClient && (
            <>
              <div style={{ borderBottom: "1px solid #f3f4f6", background: "#fafafa" }}>
                {/* Top row: name + controls */}
                <div style={{ padding: "10px 20px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <div style={{ fontWeight: 700, fontSize: 14 }}>{activeClient.name}</div>
                    {activeClient.type === "visitor" && (
                      <button onClick={() => { setModal("upgrade"); setMForm({ name: activeClient.name }); }} style={{ padding: "2px 8px", borderRadius: 6, border: "1px solid #6B21C8", background: "white", color: "#6B21C8", fontSize: 10, cursor: "pointer", fontFamily: "inherit" }}>升級為潛在客戶</button>
                    )}
                    {activeClient.type === "prospect" && (
                      <div style={{ display: "flex", gap: 5, alignItems: "center" }}>
                        <span style={{ fontSize: 10, color: "#9ca3af" }}>{activeClient.account}</span>
                        {WEIGHTS.map(w => (
                          <button key={w.id} onClick={() => updC(activeId, { weight: w.id })} style={{ padding: "2px 7px", borderRadius: 8, fontSize: 10, cursor: "pointer", fontFamily: "inherit", border: `1px solid ${activeClient.weight === w.id ? w.border : "#e5e7eb"}`, background: activeClient.weight === w.id ? w.bg : "white", color: activeClient.weight === w.id ? w.color : "#9ca3af", fontWeight: activeClient.weight === w.id ? 600 : 400 }}>{w.label}</button>
                        ))}
                      </div>
                    )}
                  </div>
                  {/* Expand/collapse toggle — only for prospects with note */}
                  {activeClient.type === "prospect" && activeClient.note && (
                    <button onClick={() => setNoteOpen(v => !v)} style={{ padding: "2px 8px", borderRadius: 6, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", fontSize: 11, cursor: "pointer", fontFamily: "inherit", display: "flex", alignItems: "center", gap: 4 }}>
                      {noteOpen ? "▲ 收起" : "▼ 說明"}
                    </button>
                  )}
                </div>
                {/* Note section — expandable, default open */}
                {activeClient.type === "prospect" && activeClient.note && noteOpen && (
                  <div style={{ margin: "0 20px 10px", padding: "8px 12px", background: "white", borderRadius: 8, border: "1px solid #e5e7eb", fontSize: 12, color: "#374151", lineHeight: 1.7, whiteSpace: "pre-wrap" }}>
                    {activeClient.note}
                  </div>
                )}
              </div>

              <div ref={chatRef} style={{ flex: 1, overflowY: "auto", padding: 20, display: "flex", flexDirection: "column", gap: 14 }}>
                {messages.length === 0 && (
                  <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 10, padding: "40px 0", color: "#9ca3af" }}>
                    <div style={{ fontSize: 28 }}>💬</div>
                    <div style={{ fontSize: 12, textAlign: "center", lineHeight: 1.8 }}>貼上 {activeClient.name} 的訊息，或拖移截圖到輸入欄</div>
                  </div>
                )}
                {messages.map((msg, i) => {
                  const isLast = i === messages.length - 1;
                  if (msg.type === "user") {
                    return (
                      <div key={i} style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 3 }}>
                        <div style={{ fontSize: 10, color: "#9ca3af" }}>客戶訊息 {msg.time}</div>
                        {msg.imgSrc && <img src={msg.imgSrc} alt="截圖" style={{ maxWidth: 240, borderRadius: 10, border: "1px solid #e5e7eb", marginBottom: 2 }} />}
                        {msg.text && msg.text !== "（截圖）" && <div style={{ maxWidth: "80%", padding: "10px 14px", borderRadius: 14, borderBottomRightRadius: 3, background: "#f3f4f6", border: "1px solid #e5e7eb", fontSize: 13, lineHeight: 1.7, whiteSpace: "pre-wrap", color: "#111" }}>{msg.text}</div>}
                      </div>
                    );
                  }
                  return (
                    <div key={i} style={{ background: "#fafafa", border: "1px solid #e5e7eb", borderRadius: 14, padding: 14, display: "flex", flexDirection: "column", gap: 10 }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                        <div style={{ fontSize: 10, color: "#6B21C8", fontWeight: 600, letterSpacing: 1 }}>{msg.isAnalyse ? "🔍 客戶分析" : "✦ 建議回覆"} {msg.time}</div>
                        {msg.grade && GS[msg.grade] && <Tag label={GS[msg.grade].label} {...GS[msg.grade]} />}
                      </div>
                      <div style={{ fontSize: 13, lineHeight: 1.8, whiteSpace: "pre-wrap", color: "#111" }}>{msg.text}</div>
                      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                        <CopyText text={msg.text} />
                        {isLast && !loading && (
                          <button onClick={() => {
                            const lu = [...messages].reverse().find(m => m.type === "user");
                            if (!lu) return;
                            const imgObj = lu.imgSrc ? (() => { const p = lu.imgSrc.split(","); return { data: p[1], mime: p[0].split(":")[1]?.split(";")[0] || "image/jpeg", preview: lu.imgSrc }; })() : null;
                            sendClient(activeId, lu.text === "（截圖）" ? "" : lu.text, imgObj, false, true);
                          }} style={{ padding: "4px 12px", borderRadius: 6, border: "1px solid #6B21C8", background: "white", color: "#6B21C8", fontSize: 11, cursor: "pointer", fontFamily: "inherit", fontWeight: 500 }}>🔍 分析</button>
                        )}
                        {isLast && !loading && !msg.isAnalyse && (
                          <button onClick={() => {
                            const lu = [...messages].reverse().find(m => m.type === "user");
                            if (!lu) return;
                            const imgObj = lu.imgSrc ? (() => { const p = lu.imgSrc.split(","); return { data: p[1], mime: p[0].split(":")[1]?.split(";")[0] || "image/jpeg", preview: lu.imgSrc }; })() : null;
                            sendClient(activeId, lu.text === "（截圖）" ? "" : lu.text, imgObj, true);
                          }} style={{ padding: "4px 12px", borderRadius: 6, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>重新生成</button>
                        )}
                      </div>
                    </div>
                  );
                })}
                {loading && activeId !== "__asst__" && (
                  <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-start", gap: 3 }}>
                    <div style={{ fontSize: 10, color: "#9ca3af" }}>AI 生成中...</div>
                    <Dots />
                  </div>
                )}
              </div>

              <div onDragOver={e => { e.preventDefault(); setDragging(true); }} onDragLeave={e => { if (!e.currentTarget.contains(e.relatedTarget)) setDragging(false); }} onDrop={handleDrop} style={{ padding: "8px 20px 14px", borderTop: "1px solid #f3f4f6" }}>
                {/* Stage selector above input */}
                {(() => {
                  const validStages = ["建立信任", "觸發需求", "報價切入"];
                  const curStage = validStages.includes(activeClient.stage) ? activeClient.stage : "建立信任";
                  // Auto-fix stage if it's an old value
                  if (!validStages.includes(activeClient.stage)) {
                    setTimeout(() => updC(activeId, { stage: "建立信任" }), 0);
                  }
                  return (
                    <div style={{ display: "flex", gap: 5, marginBottom: 8 }}>
                      {validStages.map(s => (
                        <button key={s} onClick={() => updC(activeId, { stage: s })} style={{ padding: "4px 12px", borderRadius: 16, fontSize: 11, cursor: "pointer", fontFamily: "inherit", fontWeight: curStage === s ? 600 : 400, border: curStage === s ? "1px solid #6B21C8" : "1px solid #e5e7eb", background: curStage === s ? "#6B21C8" : "white", color: curStage === s ? "white" : "#6b7280", transition: "all 0.15s" }}>{s}</button>
                      ))}
                    </div>
                  );
                })()}
                {imgPrev && (
                  <div style={{ position: "relative", display: "inline-block", marginBottom: 8 }}>
                    <img src={imgPrev} alt="預覽" style={{ height: 72, borderRadius: 8, border: "1px solid #e5e7eb" }} />
                    <button onClick={() => { setImgPrev(null); setImgB64(null); }} style={{ position: "absolute", top: -6, right: -6, width: 18, height: 18, borderRadius: "50%", background: "#ef4444", border: "none", color: "white", fontSize: 10, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
                  </div>
                )}
                {dragging && <div style={{ border: "2px dashed #6B21C8", borderRadius: 10, padding: 14, textAlign: "center", color: "#6B21C8", fontSize: 12, marginBottom: 8, background: "#f3f0ff" }}>放開滑鼠上傳截圖</div>}
                <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
                  <button onClick={() => cliFileRef.current?.click()} style={{ width: 38, height: 38, borderRadius: 8, border: "1px solid #e5e7eb", background: "white", color: "#6b7280", cursor: "pointer", fontSize: 16, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }} title="上傳截圖">📷</button>
                  <input ref={cliFileRef} type="file" accept="image/*" onChange={e => handleFile(e, false)} style={{ display: "none" }} />
                  <textarea value={input} onChange={e => setInput(e.target.value)} onKeyDown={handleKey} placeholder={`貼上 ${activeClient.name} 說的話，或拖移截圖到此處...`} style={{ flex: 1, border: dragging ? "1px solid #6B21C8" : "1px solid #e5e7eb", borderRadius: 10, padding: "9px 13px", color: "#111", fontFamily: "inherit", fontSize: 13, lineHeight: 1.6, resize: "none", minHeight: 64, outline: "none", background: dragging ? "#f3f0ff" : "white", transition: "all 0.15s" }} />
                  <button onClick={send} disabled={loading || (!input.trim() && !imgB64)} style={{ width: 38, height: 38, borderRadius: 8, border: "none", background: loading || (!input.trim() && !imgB64) ? "#e5e7eb" : "#6B21C8", color: "white", cursor: "pointer", flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15 }}>➤</button>
                </div>
              </div>
            </>
          )}
        </div>

        {/* SHORTCUTS OVERLAY PANEL */}
        <div style={{ position: "fixed", top: 53, right: 0, bottom: 0, width: 360, borderLeft: "1px solid #e5e7eb", background: "white", display: "flex", flexDirection: "column", zIndex: 40, boxShadow: scOpen ? "-4px 0 20px rgba(0,0,0,0.08)" : "none", transform: scOpen ? "translateX(0)" : "translateX(100%)", transition: "transform 0.25s cubic-bezier(0.4,0,0.2,1), box-shadow 0.25s", pointerEvents: scOpen ? "auto" : "none" }}>
            <div style={{ padding: "14px 16px", borderBottom: "1px solid #e5e7eb", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ fontWeight: 700, fontSize: 14 }}>⚡ 回覆快捷鍵</div>
              <button onClick={() => setScOpen(false)} style={{ background: "none", border: "none", color: "#9ca3af", cursor: "pointer", fontSize: 16 }}>✕</button>
            </div>
            <div style={{ padding: "10px 16px", borderBottom: "1px solid #f3f4f6" }}>
              <input value={scSearch} onChange={e => setScSearch(e.target.value)} placeholder="搜尋快捷鍵..." style={{ width: "100%", border: "1px solid #e5e7eb", borderRadius: 8, padding: "7px 12px", fontSize: 12, outline: "none", fontFamily: "inherit", background: "#fafafa" }} />
            </div>
            <div style={{ flex: 1, overflowY: "auto", padding: "12px 16px" }}>
              {scSearch.trim() ? (
                filteredSc.length === 0 ? (
                  <div style={{ textAlign: "center", color: "#9ca3af", fontSize: 12, padding: "20px 0" }}>找不到符合的快捷鍵</div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    {filteredSc.map(item => (
                      <div key={item.id} style={{ background: "#fafafa", border: "1px solid #e5e7eb", borderRadius: 8, padding: "10px 12px" }}>
                        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8 }}>
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 2 }}>{item.catName}</div>
                            <div style={{ fontSize: 12, fontWeight: 600, color: "#374151", marginBottom: 4 }}>{item.title}</div>
                            {item.text && <div style={{ fontSize: 12, color: "#6b7280", lineHeight: 1.6, whiteSpace: "pre-wrap" }}>{item.text}</div>}
                          </div>
                          {item.text && <CopyText text={item.text} label="複製" />}
                        </div>
                        {item.images && item.images.length > 0 && (
                          <ImgGrid images={item.images} onPreview={url => window.open(url, "_blank")} />
                        )}
                      </div>
                    ))}
                  </div>
                )
              ) : (
                <ShortcutsPanel sc={sc} setSc={setSc} />
              )}
            </div>
          </div>
        )}
      </div>

      {/* CONTEXT MENU */}
      {ctxMenu && (
        <div onClick={() => setCtxMenu(null)} style={{ position: "fixed", inset: 0, zIndex: 150 }}>
          <div onClick={e => e.stopPropagation()} style={{ position: "fixed", top: ctxMenu.y, left: ctxMenu.x, background: "white", border: "1px solid #e5e7eb", borderRadius: 10, padding: "4px 0", boxShadow: "0 8px 30px rgba(0,0,0,0.15)", zIndex: 151, minWidth: 140 }}>
            {pinnedIds.has(ctxMenu.id) ? (
              <button onClick={() => unpinClient(ctxMenu.id)} style={{ display: "block", width: "100%", padding: "9px 16px", textAlign: "left", background: "none", border: "none", fontSize: 13, cursor: "pointer", fontFamily: "inherit", color: "#374151" }}>📌 取消置頂</button>
            ) : (
              <button onClick={() => pinClient(ctxMenu.id)} style={{ display: "block", width: "100%", padding: "9px 16px", textAlign: "left", background: "none", border: "none", fontSize: 13, cursor: "pointer", fontFamily: "inherit", color: "#374151" }}>📌 置頂</button>
            )}
            <button onClick={() => delClient(ctxMenu.id)} style={{ display: "block", width: "100%", padding: "9px 16px", textAlign: "left", background: "none", border: "none", fontSize: 13, cursor: "pointer", fontFamily: "inherit", color: "#ef4444" }}>🗑 刪除</button>
          </div>
        </div>
      )}

      {/* MODALS */}
      {modal === "visitor" && (
        <Popup title="新增遊客對話" onClose={() => { setModal(null); setMForm({}); }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <input style={INP} placeholder="客戶名稱或IG帳號" value={mForm.name || ""} onChange={e => setMForm(p => ({ ...p, name: e.target.value }))} onKeyDown={e => e.key === "Enter" && mForm.name?.trim() && createVisitor(mForm.name.trim())} autoFocus />
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button style={BTN(false)} onClick={() => { setModal(null); setMForm({}); }}>取消</button>
              <button style={BTN(true)} onClick={() => mForm.name?.trim() && createVisitor(mForm.name.trim())}>建立</button>
            </div>
          </div>
        </Popup>
      )}
      {modal === "prospect" && (
        <Popup title="新增潛在客戶" onClose={() => { setModal(null); setMForm({}); }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <input style={INP} placeholder="姓名 *" value={mForm.name || ""} onChange={e => setMForm(p => ({ ...p, name: e.target.value }))} autoFocus />
            <input style={INP} placeholder="IG帳號 *" value={mForm.account || ""} onChange={e => setMForm(p => ({ ...p, account: e.target.value }))} />
            <textarea style={{ ...INP, resize: "none", minHeight: 72 }} placeholder="說明（背景、需求、狀況）" value={mForm.note || ""} onChange={e => setMForm(p => ({ ...p, note: e.target.value }))} />
            <div>
              <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 6 }}>成交率</div>
              <div style={{ display: "flex", gap: 8 }}>
                {WEIGHTS.map(w => (
                  <button key={w.id} onClick={() => setMForm(p => ({ ...p, weight: w.id }))} style={{ flex: 1, padding: "6px 0", borderRadius: 8, fontSize: 12, cursor: "pointer", fontFamily: "inherit", border: `1px solid ${mForm.weight === w.id ? w.border : "#e5e7eb"}`, background: mForm.weight === w.id ? w.bg : "white", color: mForm.weight === w.id ? w.color : "#6b7280", fontWeight: mForm.weight === w.id ? 700 : 400 }}>{w.label}</button>
                ))}
              </div>
            </div>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button style={BTN(false)} onClick={() => { setModal(null); setMForm({}); }}>取消</button>
              <button style={BTN(true)} onClick={() => mForm.name?.trim() && mForm.account?.trim() && createProspect(mForm.name.trim(), mForm.account.trim(), mForm.note || "", mForm.weight || "mid")}>建立</button>
            </div>
          </div>
        </Popup>
      )}
      {modal === "edit_prospect" && (
        <Popup title="編輯潛在客戶" onClose={() => { setModal(null); setMForm({}); }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <input style={INP} placeholder="姓名 *" value={mForm.name || ""} onChange={e => setMForm(p => ({ ...p, name: e.target.value }))} autoFocus />
            <input style={INP} placeholder="IG帳號 *" value={mForm.account || ""} onChange={e => setMForm(p => ({ ...p, account: e.target.value }))} />
            <textarea style={{ ...INP, resize: "none", minHeight: 80 }} placeholder="說明（背景、需求、狀況）" value={mForm.note || ""} onChange={e => setMForm(p => ({ ...p, note: e.target.value }))} />
            <div>
              <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 6 }}>成交率</div>
              <div style={{ display: "flex", gap: 8 }}>
                {WEIGHTS.map(w => (
                  <button key={w.id} onClick={() => setMForm(p => ({ ...p, weight: w.id }))} style={{ flex: 1, padding: "6px 0", borderRadius: 8, fontSize: 12, cursor: "pointer", fontFamily: "inherit", border: `1px solid ${mForm.weight === w.id ? w.border : "#e5e7eb"}`, background: mForm.weight === w.id ? w.bg : "white", color: mForm.weight === w.id ? w.color : "#6b7280", fontWeight: mForm.weight === w.id ? 700 : 400 }}>{w.label}</button>
                ))}
              </div>
            </div>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button style={BTN(false)} onClick={() => { setModal(null); setMForm({}); }}>取消</button>
              <button style={BTN(true)} onClick={() => {
                if (!mForm.name?.trim()) return;
                updC(mForm.id, { name: mForm.name.trim(), account: mForm.account.trim(), note: mForm.note, weight: mForm.weight || "mid" });
                setModal(null); setMForm({});
              }}>儲存</button>
            </div>
          </div>
        </Popup>
      )}

      {modal === "upgrade" && (
        <Popup title="升級為潛在客戶" onClose={() => { setModal(null); setMForm({}); }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <input style={INP} placeholder="姓名 *" value={mForm.name || ""} onChange={e => setMForm(p => ({ ...p, name: e.target.value }))} autoFocus />
            <input style={INP} placeholder="IG帳號 *" value={mForm.account || ""} onChange={e => setMForm(p => ({ ...p, account: e.target.value }))} />
            <textarea style={{ ...INP, resize: "none", minHeight: 72 }} placeholder="說明" value={mForm.note || ""} onChange={e => setMForm(p => ({ ...p, note: e.target.value }))} />
            <div>
              <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 6 }}>成交率</div>
              <div style={{ display: "flex", gap: 8 }}>
                {WEIGHTS.map(w => (
                  <button key={w.id} onClick={() => setMForm(p => ({ ...p, weight: w.id }))} style={{ flex: 1, padding: "6px 0", borderRadius: 8, fontSize: 12, cursor: "pointer", fontFamily: "inherit", border: `1px solid ${mForm.weight === w.id ? w.border : "#e5e7eb"}`, background: mForm.weight === w.id ? w.bg : "white", color: mForm.weight === w.id ? w.color : "#6b7280", fontWeight: mForm.weight === w.id ? 700 : 400 }}>{w.label}</button>
                ))}
              </div>
            </div>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button style={BTN(false)} onClick={() => { setModal(null); setMForm({}); }}>取消</button>
              <button style={BTN(true)} onClick={() => upgradeVisitor(activeId)}>確認升級</button>
            </div>
          </div>
        </Popup>
      )}

      {/* TRASH MODAL */}
      {showTrash && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
          <div style={{ background: "white", borderRadius: 16, width: "100%", maxWidth: 520, maxHeight: "80vh", display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,0.15)" }}>
            {/* Header */}
            <div style={{ padding: "16px 20px", borderBottom: "1px solid #e5e7eb", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ fontWeight: 700, fontSize: 15 }}>🗑 垃圾桶</div>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontSize: 11, color: "#9ca3af" }}>項目保留7天後自動刪除</span>
                <button onClick={() => setShowTrash(false)} style={{ background: "none", border: "none", color: "#999", cursor: "pointer", fontSize: 18 }}>✕</button>
              </div>
            </div>
            {/* Tabs */}
            <div style={{ display: "flex", borderBottom: "1px solid #e5e7eb" }}>
              {[["visitor","遊客"], ["prospect","潛在客戶"], ["shortcut","快捷鍵分類"]].map(([tab, label]) => {
                const count = trash.filter(t => t.trashType === (tab === "visitor" || tab === "prospect" ? "client" : "shortcut") && (tab === "shortcut" || t.type === tab)).length;
                return (
                  <button key={tab} onClick={() => setTrashTab(tab)} style={{ flex: 1, padding: "10px 0", fontSize: 12, fontWeight: trashTab === tab ? 700 : 400, cursor: "pointer", fontFamily: "inherit", border: "none", background: "transparent", color: trashTab === tab ? "#6B21C8" : "#6b7280", borderBottom: trashTab === tab ? "2px solid #6B21C8" : "2px solid transparent" }}>
                    {label}{count > 0 && ` (${count})`}
                  </button>
                );
              })}
            </div>
            {/* Content */}
            <div style={{ flex: 1, overflowY: "auto", padding: 16 }}>
              {(() => {
                const now = Date.now();
                const items = trash.filter(t => {
                  if (trashTab === "shortcut") return t.trashType === "shortcut";
                  return t.trashType === "client" && t.type === trashTab;
                });
                if (items.length === 0) return <div style={{ textAlign: "center", color: "#9ca3af", fontSize: 13, padding: "24px 0" }}>沒有已刪除的項目</div>;
                return items.map((item, i) => {
                  const daysLeft = Math.ceil((TRASH_TTL - (now - item.deletedAt)) / (24 * 60 * 60 * 1000));
                  const deletedDate = new Date(item.deletedAt).toLocaleDateString("zh-TW");
                  return (
                    <div key={i} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 12px", border: "1px solid #e5e7eb", borderRadius: 10, marginBottom: 8 }}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: 13, fontWeight: 600, color: "#111", marginBottom: 2 }}>
                          {item.name}
                          {item.type === "prospect" && item.account && <span style={{ fontSize: 11, color: "#9ca3af", fontWeight: 400, marginLeft: 6 }}>{item.account}</span>}
                          {item.trashType === "shortcut" && <span style={{ fontSize: 11, color: "#9ca3af", fontWeight: 400, marginLeft: 6 }}>{(item.items||[]).length}則快捷鍵</span>}
                        </div>
                        <div style={{ fontSize: 11, color: "#9ca3af" }}>
                          刪除於 {deletedDate} · 還有 <span style={{ color: daysLeft <= 2 ? "#ef4444" : "#f59e0b" }}>{daysLeft} 天</span>自動清除
                        </div>
                      </div>
                      <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                        <button onClick={() => item.trashType === "client" ? restoreClient(item) : restoreCat(item)} style={{ padding: "4px 10px", borderRadius: 6, border: "1px solid #6B21C8", background: "white", color: "#6B21C8", fontSize: 11, cursor: "pointer", fontFamily: "inherit", fontWeight: 500 }}>還原</button>
                        <button onClick={() => item.trashType === "client" ? permDelClient(item) : permDelCat(item)} style={{ padding: "4px 10px", borderRadius: 6, border: "1px solid #fecaca", background: "white", color: "#ef4444", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>永久刪除</button>
                      </div>
                    </div>
                  );
                });
              })()}
            </div>
            {/* Footer */}
            {trash.length > 0 && (
              <div style={{ padding: "12px 20px", borderTop: "1px solid #f3f4f6", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <span style={{ fontSize: 11, color: "#9ca3af" }}>
                  {(() => {
                    const count = trash.filter(t => {
                      if (trashTab === "shortcut") return t.trashType === "shortcut";
                      return t.trashType === "client" && t.type === trashTab;
                    }).length;
                    return count > 0 ? `目前分類 ${count} 筆` : "目前分類無資料";
                  })()}
                </span>
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={() => {
                    setTrash(p => p.filter(t => {
                      if (trashTab === "shortcut") return t.trashType !== "shortcut";
                      return !(t.trashType === "client" && t.type === trashTab);
                    }));
                  }} style={{ padding: "6px 14px", borderRadius: 8, border: "1px solid #fecaca", background: "white", color: "#ef4444", fontSize: 12, cursor: "pointer", fontFamily: "inherit", fontWeight: 500 }}>
                    刪除此分類全部
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      <style>{`
        @keyframes dot { 0%,60%,100%{transform:translateY(0);opacity:0.4} 30%{transform:translateY(-5px);opacity:1} }
        textarea::placeholder, input::placeholder { color:#9ca3af; }
        textarea:focus, input:focus { border-color:#6B21C8 !important; }
        ::-webkit-scrollbar { width:3px; }
        ::-webkit-scrollbar-thumb { background:#e5e7eb; border-radius:2px; }
      `}</style>
    </div>
  );
}
