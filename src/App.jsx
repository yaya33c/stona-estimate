import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ── Fonts & CSS ───────────────────────────────────────────────────────────────
const fontLink = document.createElement("link");
fontLink.rel = "stylesheet";
fontLink.href = "https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@300;400;500;600&family=IBM+Plex+Mono:wght@300;400;500&family=Noto+Sans+JP:wght@300;400;500&display=swap";
document.head.appendChild(fontLink);

const styleEl = document.createElement("style");
styleEl.textContent = `
  * { box-sizing:border-box; -webkit-tap-highlight-color:transparent; }
  body { margin:0; }
  .tap-btn:active { opacity:0.6; transform:scale(0.97); }
  .row-card:active { background:#F0EBE3 !important; }
  input:focus,select:focus,textarea:focus { outline:2px solid #C17B2F; outline-offset:-1px; }
  .cell-in { border:none;background:transparent;width:100%;font-family:'Noto Sans JP',sans-serif;font-size:13px;color:#1A1A1A;outline:none;padding:0; }
  .num-in  { border:none;background:transparent;width:100%;font-family:'IBM Plex Mono',monospace;font-size:13px;color:#1A1A1A;outline:none;text-align:right;padding:0; }
  .slide-up { animation:su 0.22s ease; }
  @keyframes su { from{transform:translateY(100%);opacity:0} to{transform:translateY(0);opacity:1} }
  .fade-in  { animation:fi 0.18s ease; }
  @keyframes fi { from{opacity:0} to{opacity:1} }
  /* Print */
  @media print {
    @page { size:A4; margin:12mm 10mm; }
    body { background:#fff !important; }
    .no-print { display:none !important; }
    .print-only { display:block !important; }
    .page-break { page-break-after:always; break-after:page; }
    .print-th { background:#1A1A1A !important; color:#C17B2F !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    .print-hbar { background:#1A1A1A !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    .bg-stripe { background:#FDFBF8 !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    .badge-g { background:#e8f5ec !important; color:#2a8a4a !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    .badge-o { background:#fdf0e0 !important; color:#C17B2F !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    .badge-r { background:#fde8e8 !important; color:#cc4444 !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    .p-cop { color:#C17B2F !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
  }
  .print-only { display:none; }
`;
document.head.appendChild(styleEl);

// ── Constants ─────────────────────────────────────────────────────────────────
const COP="#C17B2F", DARK="#1A1A1A", BG="#F7F4F0", BORDER="#E5E0D8", MUTED="#999";
const UNITS=["式","㎡","m","m²","m³","本","個","枚","ヶ所","箇所","セット","台","組"];
const uid=()=>Math.random().toString(36).slice(2,9);
const fmt=n=>(n===""||n===null||n===undefined||isNaN(Number(n)))?"―":Number(n).toLocaleString("ja-JP");
const pct=n=>(!isFinite(n)||isNaN(n)||n===null)?"―":(n*100).toFixed(1)+"%";
const useIsMobile=()=>{
  const[m,setM]=useState(window.innerWidth<768);
  useEffect(()=>{const h=()=>setM(window.innerWidth<768);window.addEventListener("resize",h);return()=>window.removeEventListener("resize",h);},[]);
  return m;
};

// ── Data ──────────────────────────────────────────────────────────────────────
const INIT_COVER={estimateNo:"0000001",estimateDate:new Date().toISOString().slice(0,10),client:"",projectName:"",location:"",content:"",paymentTerms:"",validityPeriod:"見積り提出後30日間",constructionPeriod:"別途ご相談の上決定",note:"",taxRate:10,discount:0};
const INIT_SECS=["仮設工事","解体工事","大工工事","内装工事","電気工事","設備工事","左官工事","タイル工事","衛生設備機器","建材","諸経費"].map(name=>({id:uid(),name,items:[]}));
const blankItem=()=>({id:uid(),name:"",spec:"",qty:"",unit:"式",unitPrice:"",note:"",costUnitPrice:""});

// ── Calc ──────────────────────────────────────────────────────────────────────
const calcItem=it=>{const q=Number(it.qty)||0,u=Number(it.unitPrice)||0,c=Number(it.costUnitPrice)||0,amount=q*u,cost=q*c,gross=amount-cost;return{amount,cost,gross,margin:amount>0?gross/amount:null};};
const calcSec=s=>{const r=s.items.reduce((a,i)=>{const c=calcItem(i);return{amount:a.amount+c.amount,cost:a.cost+c.cost}},{amount:0,cost:0});return{...r,gross:r.amount-r.cost,margin:r.amount>0?(r.amount-r.cost)/r.amount:null};};

// ── UI Atoms ──────────────────────────────────────────────────────────────────
const Btn=({style,children,...p})=><button className="tap-btn" style={{border:"none",cursor:"pointer",fontFamily:"'Noto Sans JP',sans-serif",...style}} {...p}>{children}</button>;
const Inp=({style,...p})=><input style={{width:"100%",padding:"10px 12px",border:`1px solid ${BORDER}`,background:"#fff",fontSize:15,color:DARK,...style}} {...p}/>;
const Label=({c,children})=><div style={{fontSize:11,color:c||MUTED,letterSpacing:"0.1em",marginBottom:5,textTransform:"uppercase"}}>{children}</div>;
const Badge=({margin,cls})=>{
  if(margin===null)return<span style={{color:MUTED,fontSize:11,fontFamily:"'IBM Plex Mono',monospace"}}>―</span>;
  const col=margin>=0.2?"#2a8a4a":margin>=0.1?COP:"#cc4444";
  const bc=margin>=0.2?"badge-g":margin>=0.1?"badge-o":"badge-r";
  return<span className={cls||""} style={{background:col+"22",color:col,fontSize:11,fontFamily:"'IBM Plex Mono',monospace",padding:"2px 7px"}}>{pct(margin)}</span>;
};
const Card=({title,action,children,style})=>(
  <div style={{background:"#fff",border:`1px solid ${BORDER}`,marginBottom:16,...style}}>
    <div style={{padding:"12px 20px",borderBottom:`1px solid ${BORDER}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <span style={{fontFamily:"'Noto Serif JP',serif",fontSize:12,color:MUTED,letterSpacing:"0.15em",textTransform:"uppercase"}}>{title}</span>
      {action}
    </div>
    <div style={{padding:"18px 20px"}}>{children}</div>
  </div>
);

// ── Item Modal ────────────────────────────────────────────────────────────────
function ItemModal({item,onSave,onDelete,onClose}){
  const[d,setD]=useState({...item});
  const u=(k,v)=>setD(p=>({...p,[k]:v}));
  const c=calcItem(d);
  return(
    <div style={{position:"fixed",inset:0,zIndex:200,display:"flex",flexDirection:"column"}}>
      <div style={{flex:1,background:"rgba(0,0,0,0.55)"}} onClick={onClose}/>
      <div className="slide-up" style={{background:"#fff",borderRadius:"20px 20px 0 0",paddingBottom:"env(safe-area-inset-bottom,16px)"}}>
        <div style={{display:"flex",justifyContent:"center",padding:"12px 0 4px"}}><div style={{width:36,height:4,background:"#DDD",borderRadius:2}}/></div>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 20px 14px",borderBottom:`1px solid ${BORDER}`}}>
          <span style={{fontFamily:"'Noto Serif JP',serif",fontSize:16,fontWeight:500}}>項目を編集</span>
          <div style={{display:"flex",gap:10}}>
            <Btn style={{color:"#cc4444",background:"none",fontSize:14,padding:"6px 10px"}} onClick={()=>{onDelete(item.id);onClose();}}>削除</Btn>
            <Btn style={{background:COP,color:"#fff",padding:"8px 22px",fontSize:14}} onClick={()=>{onSave(d);onClose();}}>保存</Btn>
          </div>
        </div>
        <div style={{padding:"16px 20px",display:"flex",flexDirection:"column",gap:14,maxHeight:"67vh",overflowY:"auto"}}>
          <div><Label>項目名</Label><Inp value={d.name} placeholder="例：床直貼り" onChange={e=>u("name",e.target.value)}/></div>
          <div><Label>仕様・摘要</Label><Inp value={d.spec} placeholder="例：朝日ウッドテック" onChange={e=>u("spec",e.target.value)}/></div>
          <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:12}}>
            <div><Label>数量</Label><Inp style={{fontFamily:"'IBM Plex Mono',monospace"}} type="number" value={d.qty} placeholder="0" onChange={e=>u("qty",e.target.value)}/></div>
            <div><Label>単位</Label><select style={{width:"100%",padding:"10px 8px",border:`1px solid ${BORDER}`,background:"#fff",fontSize:15}} value={d.unit} onChange={e=>u("unit",e.target.value)}>{UNITS.map(v=><option key={v}>{v}</option>)}</select></div>
          </div>
          <div><Label>単価</Label><Inp style={{fontFamily:"'IBM Plex Mono',monospace"}} type="number" value={d.unitPrice} placeholder="0" onChange={e=>u("unitPrice",e.target.value)}/></div>
          <div style={{background:BG,padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{color:MUTED,fontSize:13}}>金額（自動）</span>
            <span style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:18,fontWeight:500}}>¥{fmt(c.amount)}</span>
          </div>
          <div style={{borderTop:`1px solid ${BORDER}`,paddingTop:14}}>
            <Label c={COP}>▸ 原価管理</Label>
            <Inp style={{fontFamily:"'IBM Plex Mono',monospace",marginTop:6}} type="number" value={d.costUnitPrice} placeholder="0" onChange={e=>u("costUnitPrice",e.target.value)}/>
            <div style={{marginTop:10,background:BG,padding:"10px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{color:MUTED,fontSize:13}}>原価 / 粗利率</span>
              <div style={{display:"flex",gap:12,alignItems:"center"}}>
                <span style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:14,color:"#8B6030"}}>¥{fmt(c.cost)}</span>
                <Badge margin={c.margin}/>
              </div>
            </div>
          </div>
          <div><Label>備考</Label><Inp value={d.note} onChange={e=>u("note",e.target.value)}/></div>
        </div>
      </div>
    </div>
  );
}

// ── Cover Page ────────────────────────────────────────────────────────────────
function CoverPage({cover,setCover,isMobile}){
  const u=(k,v)=>setCover(c=>({...c,[k]:v}));
  const f=(key,label,type="text")=>(
    <div><Label>{label}</Label><Inp type={type} value={cover[key]} onChange={e=>u(key,e.target.value)}/></div>
  );
  const g=isMobile?{display:"flex",flexDirection:"column",gap:14}:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"14px 28px"};
  return(
    <div>
      <Card title="見積情報">
        <div style={g}>
          {f("estimateNo","見積番号")} {f("estimateDate","見積作成日","date")}
          <div><Label>消費税率</Label><select style={{width:"100%",padding:"10px 12px",border:`1px solid ${BORDER}`,background:"#fff",fontSize:15}} value={cover.taxRate} onChange={e=>u("taxRate",Number(e.target.value))}>{[8,10].map(r=><option key={r} value={r}>{r}%</option>)}</select></div>
          {f("validityPeriod","見積有効期限")}
        </div>
      </Card>
      <Card title="顧客・工事情報">
        <div style={g}>
          {f("client","顧客名（宛先）")} {f("projectName","件名")}
          {f("location","工事場所")} {f("content","工事内容")}
          {f("constructionPeriod","工期")} {f("paymentTerms","支払条件")}
        </div>
      </Card>
      <Card title="その他">
        <div style={g}>
          {f("note","備考")}
          <div><Label>値引き額（税抜・マイナスで入力）</Label><Inp style={{fontFamily:"'IBM Plex Mono',monospace"}} type="number" value={cover.discount} placeholder="-886789" onChange={e=>u("discount",Number(e.target.value))}/></div>
        </div>
      </Card>
    </div>
  );
}

// ── Summary Page ──────────────────────────────────────────────────────────────
function SummaryPage({sections,setSections,cover,setView,setActiveSection,isMobile}){
  const[newName,setNewName]=useState("");
  const[editId,setEditId]=useState(null);
  const[editName,setEditName]=useState("");
  const totals=sections.map(s=>({...s,...calcSec(s)}));
  const subtotal=totals.reduce((a,s)=>a+s.amount,0)+Number(cover.discount||0);
  const tax=Math.floor(subtotal*(cover.taxRate/100));
  const total=subtotal+tax;
  const tCost=totals.reduce((a,s)=>a+s.cost,0);
  const add=()=>{if(!newName.trim())return;setSections(s=>[...s,{id:uid(),name:newName.trim(),items:[]}]);setNewName("");};
  const del=id=>{if(!confirm("削除しますか？"))return;setSections(s=>s.filter(x=>x.id!==id));};
  const saveEdit=id=>{setSections(p=>p.map(s=>s.id===id?{...s,name:editName}:s));setEditId(null);};
  const move=(id,d)=>setSections(p=>{const i=p.findIndex(s=>s.id===id),nx=i+d;if(nx<0||nx>=p.length)return p;const a=[...p];[a[i],a[nx]]=[a[nx],a[i]];return a;});
  return(
    <div>
      <div style={{background:DARK,color:"#fff",padding:isMobile?"18px 16px":"22px 28px",marginBottom:20}}>
        <div style={{fontSize:11,color:"#666",letterSpacing:"0.1em",marginBottom:4}}>御見積金額（税込）</div>
        <div style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:isMobile?28:32,color:COP}}>¥{fmt(total)}</div>
        <div style={{marginTop:14,display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
          {[["税抜金額",`¥${fmt(subtotal)}`],["消費税",`¥${fmt(tax)}`],["原価合計",`¥${fmt(tCost)}`]].map(([l,v])=>(
            <div key={l}><div style={{fontSize:10,color:"#666",marginBottom:3}}>{l}</div><div style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:14,color:"#ccc"}}>{v}</div></div>
          ))}
        </div>
      </div>
      <Card title="工事区分">
        <div>
          {totals.map((s,i)=>(
            <div key={s.id} style={{borderBottom:i<totals.length-1?`1px solid ${BORDER}`:"none",padding:"12px 0"}}>
              {editId===s.id?(
                <div style={{display:"flex",gap:8}}>
                  <Inp value={editName} onChange={e=>setEditName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&saveEdit(s.id)} autoFocus style={{flex:1}}/>
                  <Btn style={{background:COP,color:"#fff",padding:"8px 14px"}} onClick={()=>saveEdit(s.id)}>保存</Btn>
                  <Btn style={{background:"none",border:`1px solid ${BORDER}`,padding:"8px 10px",color:MUTED}} onClick={()=>setEditId(null)}>✕</Btn>
                </div>
              ):(
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <div style={{display:"flex",flexDirection:"column",gap:2,flexShrink:0}}>
                    <Btn style={{background:"none",border:`1px solid ${BORDER}`,color:MUTED,padding:"1px 6px",fontSize:11,lineHeight:1.4}} onClick={()=>move(s.id,-1)}>▲</Btn>
                    <Btn style={{background:"none",border:`1px solid ${BORDER}`,color:MUTED,padding:"1px 6px",fontSize:11,lineHeight:1.4}} onClick={()=>move(s.id,1)}>▼</Btn>
                  </div>
                  <div style={{flex:1,cursor:"pointer",minWidth:0}} onClick={()=>{setActiveSection(s.id);setView("detail");}}>
                    <div style={{fontSize:15,fontWeight:500}}>{s.name}<span style={{marginLeft:8,fontSize:12,color:MUTED,fontWeight:300}}>{s.items.length}項目</span></div>
                    <div style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:13,color:s.amount>0?COP:"#ccc",marginTop:2,display:"flex",alignItems:"center",gap:8}}>
                      ¥{fmt(s.amount)}{s.margin!==null&&<Badge margin={s.margin}/>}
                    </div>
                  </div>
                  <div style={{display:"flex",gap:6,flexShrink:0}}>
                    <Btn style={{background:"none",border:`1px solid ${BORDER}`,color:MUTED,padding:"6px 10px",fontSize:12}} onClick={()=>{setEditId(s.id);setEditName(s.name);}}>名称</Btn>
                    <Btn style={{background:DARK,color:"#fff",padding:"8px 12px",fontSize:13}} onClick={()=>{setActiveSection(s.id);setView("detail");}}>明細→</Btn>
                    <Btn style={{background:"none",color:"#cc4444",fontSize:18,padding:"4px 6px",lineHeight:1}} onClick={()=>del(s.id)}>×</Btn>
                  </div>
                </div>
              )}
            </div>
          ))}
        </div>
      </Card>
      <Card title="工事区分を追加">
        <div style={{display:"flex",gap:10}}>
          <Inp value={newName} placeholder="例：塗装工事" onChange={e=>setNewName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&add()} style={{flex:1}}/>
          <Btn style={{background:COP,color:"#fff",padding:"10px 20px",fontSize:14,whiteSpace:"nowrap"}} onClick={add}>追加</Btn>
        </div>
      </Card>
    </div>
  );
}

// ── Detail Page ───────────────────────────────────────────────────────────────
function DetailPage({sections,setSections,activeSection,isMobile}){
  const[modal,setModal]=useState(null);
  const sec=sections.find(s=>s.id===activeSection);
  if(!sec)return<div style={{color:MUTED,padding:32}}>工事区分を選んでください</div>;
  const upSec=useCallback(fn=>setSections(p=>p.map(s=>s.id===activeSection?fn(s):s)),[setSections,activeSection]);
  const addItem=()=>{const it=blankItem();upSec(s=>({...s,items:[...s.items,it]}));setModal(it);};
  const saveItem=it=>upSec(s=>({...s,items:s.items.map(i=>i.id===it.id?it:i)}));
  const delItem=id=>upSec(s=>({...s,items:s.items.filter(i=>i.id!==id)}));
  const moveItem=(id,dir)=>upSec(s=>{const items=[...s.items],idx=items.findIndex(i=>i.id===id),nx=idx+dir;if(nx<0||nx>=items.length)return s;[items[idx],items[nx]]=[items[nx],items[idx]];return{...s,items};});
  const tot=calcSec(sec);
  return(
    <div>
      <div style={{background:DARK,color:"#fff",padding:isMobile?"14px 16px":"16px 24px",marginBottom:16,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <div style={{fontSize:11,color:"#666",marginBottom:3}}>{sec.name}　小計（税抜）</div>
          <div style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:isMobile?22:26,color:COP}}>¥{fmt(tot.amount)}</div>
        </div>
        <div style={{textAlign:"right"}}>
          <div style={{fontSize:11,color:"#666",marginBottom:3}}>粗利 / 粗利率</div>
          <div style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:13,color:"#ccc"}}>¥{fmt(tot.gross)} / <span style={{color:COP}}>{pct(tot.margin)}</span></div>
        </div>
      </div>
      {isMobile?(
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {sec.items.length===0&&<div style={{textAlign:"center",color:MUTED,padding:"48px 20px",fontSize:14,background:"#fff",border:`1px solid ${BORDER}`}}>まだ項目がありません<br/><span style={{fontSize:12}}>＋ボタンで追加</span></div>}
          {sec.items.map((it,i)=>{
            const c=calcItem(it);
            return(
              <div key={it.id} className="row-card" style={{background:"#fff",border:`1px solid ${BORDER}`,padding:"14px 16px",cursor:"pointer",display:"flex",alignItems:"center",gap:10}}>
                <div style={{display:"flex",flexDirection:"column",gap:3,flexShrink:0}}>
                  <Btn style={{background:"none",border:`1px solid ${BORDER}`,color:MUTED,padding:"2px 7px",fontSize:11,lineHeight:1.4}} onClick={()=>moveItem(it.id,-1)}>▲</Btn>
                  <Btn style={{background:"none",border:`1px solid ${BORDER}`,color:MUTED,padding:"2px 7px",fontSize:11,lineHeight:1.4}} onClick={()=>moveItem(it.id,1)}>▼</Btn>
                </div>
                <div style={{flex:1,minWidth:0}} onClick={()=>setModal(it)}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:15,fontWeight:it.name?500:300,color:it.name?DARK:MUTED,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{it.name||"（未入力）"}</div>
                      {it.spec&&<div style={{fontSize:12,color:MUTED,marginTop:1}}>{it.spec}</div>}
                      {(it.qty||it.unitPrice)&&<div style={{fontSize:12,color:MUTED,fontFamily:"'IBM Plex Mono',monospace",marginTop:2}}>{it.qty||"―"}{it.unit} × ¥{fmt(it.unitPrice)}</div>}
                    </div>
                    <div style={{textAlign:"right",flexShrink:0}}>
                      <div style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:15,fontWeight:500}}>{c.amount>0?`¥${fmt(c.amount)}`:<span style={{color:MUTED}}>―</span>}</div>
                      <Badge margin={c.margin}/>
                    </div>
                  </div>
                </div>
              </div>
            );
          })}
          <Btn onClick={addItem} style={{position:"fixed",right:20,bottom:76,zIndex:50,background:COP,color:"#fff",width:56,height:56,borderRadius:"50%",fontSize:26,boxShadow:"0 4px 16px rgba(193,123,47,0.45)",display:"flex",alignItems:"center",justifyContent:"center"}}>＋</Btn>
        </div>
      ):(
        <div style={{background:"#fff",border:`1px solid ${BORDER}`,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:DARK}}>
                {[["No.",36,"c"],["項目",160,"l"],["仕様・摘要",120,"l"],["数量",70,"r"],["単位",60,"l"],["単価",90,"r"],["金額",100,"r"],["原単価",90,"r"],["原価",90,"r"],["粗利率",76,"c"],["備考",100,"l"],["",70,"c"]].map(([h,w,a])=>(
                  <th key={h+w} style={{padding:"9px 10px",color:COP,fontFamily:"'IBM Plex Mono',monospace",fontSize:11,fontWeight:400,textAlign:a==="r"?"right":a==="c"?"center":"left",width:w,whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sec.items.map((it,i)=>{
                const c=calcItem(it),bg=i%2===1?"#FDFBF8":"#fff",cb=i%2===1?"#F6F1E8":"#FAF6EC";
                const td=(a,cost)=>({padding:"6px 10px",borderBottom:`1px solid #F0EDE8`,background:cost?cb:bg,textAlign:a==="r"?"right":a==="c"?"center":"left",verticalAlign:"middle"});
                return(
                  <tr key={it.id}>
                    <td style={td("c")}><span style={{fontFamily:"'IBM Plex Mono',monospace",color:MUTED,fontSize:12}}>{i+1}</span></td>
                    <td style={td("l")}><input className="cell-in" value={it.name} placeholder="項目名" onChange={e=>saveItem({...it,name:e.target.value})}/></td>
                    <td style={td("l")}><input className="cell-in" value={it.spec} onChange={e=>saveItem({...it,spec:e.target.value})}/></td>
                    <td style={td("r")}><input className="num-in" type="number" value={it.qty} placeholder="0" onChange={e=>saveItem({...it,qty:e.target.value})}/></td>
                    <td style={td("l")}><select className="cell-in" style={{cursor:"pointer"}} value={it.unit} onChange={e=>saveItem({...it,unit:e.target.value})}>{UNITS.map(v=><option key={v}>{v}</option>)}</select></td>
                    <td style={td("r")}><input className="num-in" type="number" value={it.unitPrice} placeholder="0" onChange={e=>saveItem({...it,unitPrice:e.target.value})}/></td>
                    <td style={{...td("r"),fontFamily:"'IBM Plex Mono',monospace",fontWeight:500}}>¥{fmt(c.amount)}</td>
                    <td style={td("r",true)}><input className="num-in" style={{color:"#8B6030"}} type="number" value={it.costUnitPrice} placeholder="0" onChange={e=>saveItem({...it,costUnitPrice:e.target.value})}/></td>
                    <td style={{...td("r",true),fontFamily:"'IBM Plex Mono',monospace",color:"#8B6030"}}>¥{fmt(c.cost)}</td>
                    <td style={td("c",true)}><Badge margin={c.margin}/></td>
                    <td style={td("l")}><input className="cell-in" value={it.note} onChange={e=>saveItem({...it,note:e.target.value})}/></td>
                    <td style={td("c")}>
                      <div style={{display:"flex",gap:5,justifyContent:"center"}}>
                        <Btn style={{background:"none",border:`1px solid ${BORDER}`,color:MUTED,padding:"3px 7px",fontSize:11}} onClick={()=>moveItem(it.id,-1)}>▲</Btn>
                        <Btn style={{background:"none",border:`1px solid ${BORDER}`,color:MUTED,padding:"3px 7px",fontSize:11}} onClick={()=>moveItem(it.id,1)}>▼</Btn>
                        <Btn style={{background:"none",border:"none",color:"#cc4444",fontSize:16,padding:"0 4px",lineHeight:1}} onClick={()=>delItem(it.id)}>×</Btn>
                      </div>
                    </td>
                  </tr>
                );
              })}
              <tr style={{background:DARK}}>
                <td colSpan={6} style={{padding:"10px 12px",color:"#fff",fontFamily:"'Noto Serif JP',serif",letterSpacing:"0.1em",fontSize:13}}>小計（税抜）</td>
                <td style={{padding:"10px 12px",color:COP,fontFamily:"'IBM Plex Mono',monospace",textAlign:"right"}}>¥{fmt(tot.amount)}</td>
                <td/><td style={{padding:"10px 12px",color:"#aaa",fontFamily:"'IBM Plex Mono',monospace",textAlign:"right"}}>¥{fmt(tot.cost)}</td>
                <td style={{padding:"10px 12px",color:COP,fontFamily:"'IBM Plex Mono',monospace",textAlign:"center"}}>{pct(tot.margin)}</td>
                <td colSpan={2}/>
              </tr>
            </tbody>
          </table>
          <div style={{padding:"10px 16px",borderTop:`1px solid ${BORDER}`,background:BG}}>
            <Btn style={{background:"none",border:`1px solid ${BORDER}`,padding:"6px 16px",fontSize:13,color:DARK}} onClick={addItem}>＋ 行を追加</Btn>
          </div>
        </div>
      )}
      {modal&&<ItemModal item={sec.items.find(i=>i.id===modal.id)||modal} onSave={saveItem} onDelete={delItem} onClose={()=>setModal(null)}/>}
    </div>
  );
}

// ── Print Document ────────────────────────────────────────────────────────────
function PrintDoc({cover,sections,preview}){
  const st=sections.map(s=>({...s,...calcSec(s)}));
  const sub=st.reduce((a,s)=>a+s.amount,0)+Number(cover.discount||0);
  const tax=Math.floor(sub*(cover.taxRate/100));
  const total=sub+tax;
  const TH=({children,style})=><th className="print-th" style={{padding:"5pt 7pt",textAlign:"left",fontWeight:500,letterSpacing:"0.04em",...style}}>{children}</th>;
  const TD=({children,style})=><td style={{padding:"5pt 7pt",border:"0.5pt solid #D8D4CE",...style}}>{children}</td>;
  const mono={fontFamily:"'Noto Serif JP',serif"};
  const serif={fontFamily:"'Noto Serif JP',serif"};
  const wrap=preview?{...serif,color:DARK,fontSize:"10pt",background:"#fff",padding:"10mm",maxWidth:780}:{};

  return(
    <div className={preview?"fade-in":"print-only"} style={wrap}>

      {/* ══ 表紙 ══════════════════════════════════════════════════════════════ */}
      <div className={preview?"":"page-break"} style={{marginBottom:preview?48:0}}>

        {/* タイトルブロック */}
        <div style={{borderBottom:"2pt solid #1A1A1A",paddingBottom:"6mm",marginBottom:"8mm",display:"flex",alignItems:"flex-end",justifyContent:"space-between"}}>
          <div>
            <div style={{...serif,fontSize:"7pt",letterSpacing:"0.3em",color:"#888",marginBottom:"4pt"}}>ESTIMATE</div>
            <div style={{...serif,fontSize:"18pt",fontWeight:400,letterSpacing:"0.22em",color:DARK}}>御　見　積　書</div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{...mono,fontSize:"8pt",color:"#888",marginBottom:"2pt"}}>STONA</div>
            <div style={{...mono,fontSize:"8.5pt",color:"#555"}}>No. {cover.estimateNo}</div>
            <div style={{...mono,fontSize:"8pt",color:"#888"}}>{cover.estimateDate}</div>
          </div>
        </div>

        {/* 本文グリッド */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"10mm",marginBottom:"8mm"}}>

          {/* 左：宛先 + 御見積金額 */}
          <div>
            {/* 宛先 */}
            <div style={{marginBottom:"7mm"}}>
              <div style={{...serif,fontSize:"14pt",fontWeight:500,borderBottom:"1pt solid #1A1A1A",paddingBottom:"4pt",marginBottom:"4pt",display:"inline-block",minWidth:"60%"}}>
                {cover.client||"　"}&nbsp;御中
              </div>
              <div style={{fontSize:"8.5pt",color:"#666",marginTop:"4pt"}}>下記の通り御見積り申し上げます。</div>
            </div>

            {/* 金額ボックス */}
            <div style={{border:"1pt solid #1A1A1A",padding:"10pt 14pt",background:"#FAFAF8"}}>
              <div style={{fontSize:"7.5pt",color:"#888",letterSpacing:"0.15em",marginBottom:"4pt"}}>御　見　積　金　額（消費税込）</div>
              <div style={{...mono,fontSize:"24pt",fontWeight:400,color:DARK,letterSpacing:"0.04em"}}>
                ¥ {fmt(total)} ―
              </div>
              <div style={{borderTop:"0.5pt solid #DDD",marginTop:"8pt",paddingTop:"6pt",display:"flex",gap:"16pt",fontSize:"8.5pt",color:"#666"}}>
                <span>税抜金額　¥ {fmt(sub)}</span>
                <span>消費税（{cover.taxRate}%）　¥ {fmt(tax)}</span>
              </div>
            </div>
          </div>

          {/* 右：会社情報 + 見積条件 */}
          <div>
            {/* 発行元 */}
            <div style={{border:"1pt solid #1A1A1A",padding:"10pt 12pt",marginBottom:"5mm"}}>
              <div style={{...serif,fontSize:"11pt",fontWeight:500,color:DARK,marginBottom:"4pt",letterSpacing:"0.1em"}}>STONA</div>
              <div style={{fontSize:"8pt",color:"#555",lineHeight:1.8}}>
                <div>〒584-0036　大阪府富田林市甲田1-4-38-10</div>
                <div>TEL : 0721-55-3673　FAX : 0721-55-3674</div>
                <div>登録番号：T1120101068155</div>
              </div>
            </div>
            {/* 条件 */}
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"8.5pt"}}>
              <tbody>
                {[["有効期限",cover.validityPeriod],["工期",cover.constructionPeriod],["支払条件",cover.paymentTerms]].filter(([,v])=>v).map(([l,v])=>(
                  <tr key={l}>
                    <td style={{padding:"4pt 8pt",background:"#F2EFE9",color:"#555",whiteSpace:"nowrap",width:"35%",borderBottom:"0.5pt solid #DDD"}}>{l}</td>
                    <td style={{padding:"4pt 8pt",borderBottom:"0.5pt solid #DDD"}}>{v}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* 工事情報テーブル */}
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:"9pt"}}>
          <tbody>
            {[["件　名",cover.projectName],["工事場所",cover.location],["工事内容",cover.content],["備　考",cover.note]].filter(([,v])=>v).map(([l,v])=>(
              <tr key={l}>
                <td className="print-th" style={{width:"15%",padding:"6pt 10pt",border:"0.5pt solid #555",whiteSpace:"nowrap",textAlign:"center",letterSpacing:"0.15em"}}>{l}</td>
                <td style={{padding:"6pt 10pt",border:"0.5pt solid #D8D4CE"}}>{v}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* ══ 見積内訳書 ══════════════════════════════════════════════════════ */}
      <div className={preview?"":"page-break"} style={{marginBottom:preview?48:0}}>
        {/* ページヘッダー */}
        <div style={{borderBottom:"2pt solid #1A1A1A",paddingBottom:"4pt",marginBottom:"6mm",display:"flex",justifyContent:"space-between",alignItems:"flex-end"}}>
          <div style={{...serif,fontSize:"15pt",fontWeight:300,letterSpacing:"0.15em"}}>見　積　内　訳　書</div>
          <div style={{...mono,fontSize:"8pt",color:"#888"}}>{cover.projectName}　／　{cover.estimateNo}</div>
        </div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:"9.5pt"}}>
          <thead>
            <tr>
              <TH style={{width:28,textAlign:"center"}}>No.</TH>
              <TH>工　事　区　分</TH>
              <TH style={{width:44,textAlign:"right"}}>数量</TH>
              <TH style={{width:36,textAlign:"center"}}>単位</TH>
              <TH style={{width:110,textAlign:"right"}}>金　額（税抜）</TH>
              <TH style={{width:80,textAlign:"right"}}>備　考</TH>
            </tr>
          </thead>
          <tbody>
            {st.map((s,i)=>(
              <tr key={s.id} className={i%2===1?"bg-stripe":""}>
                <TD style={{textAlign:"center",...mono,color:"#aaa",fontSize:"8pt"}}>{i+1}</TD>
                <TD style={{fontWeight:500,letterSpacing:"0.02em"}}>{s.name}</TD>
                <TD style={{textAlign:"right",...mono}}>1</TD>
                <TD style={{textAlign:"center"}}>式</TD>
                <TD style={{textAlign:"right",...mono,fontSize:"10pt"}}>¥ {fmt(s.amount)}</TD>
                <TD style={{color:"#888",fontSize:"8.5pt"}}></TD>
              </tr>
            ))}
            {Number(cover.discount)!==0&&(
              <tr>
                <TD/><TD style={{letterSpacing:"0.05em"}}>値　引　き</TD><TD/><TD style={{textAlign:"center"}}>式</TD>
                <TD style={{textAlign:"right",...mono,fontSize:"10pt"}}>▲ {fmt(Math.abs(Number(cover.discount)))}</TD><TD/>
              </tr>
            )}
          </tbody>
          <tfoot>
            <tr className="print-hbar">
              <td colSpan={4} style={{padding:"7pt 10pt",color:"#fff",...serif,fontSize:"10pt",letterSpacing:"0.15em"}}>合　計（税　抜）</td>
              <td className="p-cop" style={{padding:"7pt 10pt",textAlign:"right",...mono,fontSize:"12pt",fontWeight:500,color:COP}}>¥ {fmt(sub)}</td>
              <td style={{padding:"7pt 10pt"}}/>
            </tr>
            <tr style={{background:"#F2EFE9"}}>
              <td colSpan={4} style={{padding:"5pt 10pt",color:"#555",fontSize:"8.5pt",letterSpacing:"0.1em"}}>消費税（{cover.taxRate}%）</td>
              <td style={{padding:"5pt 10pt",textAlign:"right",...mono,fontSize:"9.5pt"}}>¥ {fmt(tax)}</td>
              <td/>
            </tr>
            <tr style={{borderTop:"2pt solid #1A1A1A"}}>
              <td colSpan={4} style={{padding:"8pt 10pt",...serif,fontSize:"11pt",letterSpacing:"0.15em",fontWeight:500}}>御　見　積　金　額（税　込）</td>
              <td style={{padding:"8pt 10pt",textAlign:"right",...mono,fontSize:"13pt",fontWeight:500}}>¥ {fmt(total)}</td>
              <td/>
            </tr>
          </tfoot>
        </table>
      </div>

      {/* ══ 各工事明細 ═══════════════════════════════════════════════════════ */}
      {sections.filter(s=>s.items.length>0).map((s,si,arr)=>{
        const tot=calcSec(s);
        return(
          <div key={s.id} className={si<arr.length-1&&!preview?"page-break":""} style={{marginBottom:preview?48:0}}>
            <div style={{borderBottom:"2pt solid #1A1A1A",paddingBottom:"4pt",marginBottom:"6mm",display:"flex",justifyContent:"space-between",alignItems:"flex-end"}}>
              <div style={{...serif,fontSize:"14pt",fontWeight:300,letterSpacing:"0.12em"}}>見　積　明　細　書　─　{s.name}</div>
              <div style={{...mono,fontSize:"8pt",color:"#888"}}>{cover.projectName}　／　{cover.estimateNo}</div>
            </div>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"9pt"}}>
              <thead>
                <tr>
                  <TH style={{width:24,textAlign:"center"}}>No.</TH>
                  <TH style={{minWidth:120}}>項　目</TH>
                  <TH style={{minWidth:90}}>仕様・摘要</TH>
                  <TH style={{width:44,textAlign:"right"}}>数量</TH>
                  <TH style={{width:36,textAlign:"center"}}>単位</TH>
                  <TH style={{width:80,textAlign:"right"}}>単　価</TH>
                  <TH style={{width:96,textAlign:"right"}}>金　額</TH>
                  <TH style={{minWidth:70}}>備　考</TH>
                </tr>
              </thead>
              <tbody>
                {s.items.map((it,i)=>{
                  const c=calcItem(it);
                  return(
                    <tr key={it.id} className={i%2===1?"bg-stripe":""}>
                      <TD style={{textAlign:"center",...mono,color:"#aaa",fontSize:"8pt"}}>{i+1}</TD>
                      <TD style={{fontWeight:it.name?500:300,color:it.name?DARK:"#bbb"}}>{it.name||"―"}</TD>
                      <TD style={{color:"#666",fontSize:"8.5pt"}}>{it.spec}</TD>
                      <TD style={{textAlign:"right",...mono}}>{it.qty||"―"}</TD>
                      <TD style={{textAlign:"center"}}>{it.unit}</TD>
                      <TD style={{textAlign:"right",...mono}}>{it.unitPrice?`¥ ${fmt(it.unitPrice)}`:"―"}</TD>
                      <TD style={{textAlign:"right",...mono,fontWeight:500}}>{c.amount>0?`¥ ${fmt(c.amount)}`:"―"}</TD>
                      <TD style={{color:"#888",fontSize:"8.5pt"}}>{it.note}</TD>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr className="print-hbar">
                  <td colSpan={6} style={{padding:"6pt 8pt",color:"#fff",...serif,letterSpacing:"0.12em",fontSize:"9.5pt"}}>小　計（税　抜）</td>
                  <td className="p-cop" style={{padding:"6pt 8pt",textAlign:"right",...mono,fontWeight:500,fontSize:"10.5pt",color:COP}}>¥ {fmt(tot.amount)}</td>
                  <td style={{padding:"6pt 8pt"}}/>
                </tr>
              </tfoot>
            </table>
          </div>
        );
      })}
    </div>
  );
}

// ── Excel Export ──────────────────────────────────────────────────────────────
function exportExcel(cover,sections){
  const wb=XLSX.utils.book_new();
  const st=sections.map(s=>({...s,...calcSec(s)}));
  const sub=st.reduce((a,s)=>a+s.amount,0)+Number(cover.discount||0);
  const tax=Math.floor(sub*(cover.taxRate/100));
  const total=sub+tax;

  // ── ヘルパー ──────────────────────────────────────────────────────────────
  const cell=(v,t="s")=>({v,t});
  const num=v=>({v:Number(v)||0,t:"n",z:"#,##0"});
  const bold=v=>({v,t:"s",s:{font:{bold:true}}});
  const setColW=(ws,widths)=>{ ws["!cols"]=widths.map(w=>({wch:w})); };
  const setRow=(ws,r,h)=>{ if(!ws["!rows"])ws["!rows"]=[]; ws["!rows"][r]={hpt:h}; };
  const merge=(ws,r1,c1,r2,c2)=>{ if(!ws["!merges"])ws["!merges"]=[]; ws["!merges"].push({s:{r:r1,c:c1},e:{r:r2,c:c2}}); };

  // ── 表紙シート ────────────────────────────────────────────────────────────
  const coverAoa=[
    ["御　見　積　書","","","","STONA"],           // 0
    ["","","","",""],                                // 1
    [cover.client?""+cover.client+" 御中":"","","","",""],// 2
    ["下記の通り御見積り申し上げます。","","","",""],// 3
    ["","","","",""],                                // 4
    ["御見積金額（消費税込）","","","",""],          // 5
    [total,"","","",""],                             // 6  ★金額
    ["","","","",""],                                // 7
    ["税抜金額","",sub,"消費税（"+cover.taxRate+"%）",tax],// 8
    ["","","","",""],                                // 9
    ["件　名","",cover.projectName||"","",""],       // 10
    ["工事場所","",cover.location||"","",""],        // 11
    ["工事内容","",cover.content||"","",""],         // 12
    ["支払条件","",cover.paymentTerms||"","",""],    // 13
    ["工　期","",cover.constructionPeriod||"","",""],// 14
    ["有効期限","",cover.validityPeriod||"","",""],  // 15
    ["備　考","",cover.note||"","",""],              // 16
    ["","","","",""],                                // 17
    ["見積番号","",cover.estimateNo||"","",""],      // 18
    ["見積作成日","",cover.estimateDate||"","",""],  // 19
    ["","","","",""],                                // 20
    ["STONA","","","",""],                           // 21
    ["大阪府富田林市甲田1-4-38-10","","","",""],     // 22
    ["TEL : 0721-55-3673","","","",""],              // 23
    ["登録番号：T1120101068155","","","",""],        // 24
  ];
  const wsCover=XLSX.utils.aoa_to_sheet(coverAoa);
  setColW(wsCover,[16,2,28,16,14]);
  setRow(wsCover,0,28); setRow(wsCover,6,32);
  XLSX.utils.book_append_sheet(wb,wsCover,"表紙");

  // ── 小計シート ────────────────────────────────────────────────────────────
  const sumAoa=[
    ["見　積　内　訳　書","","","","",""],
    [cover.projectName||"","","","","",""],
    ["","","","","",""],
    ["No.","工　事　区　分","数量","単位","金　額（税抜）","備考"],
  ];
  st.forEach((s,i)=>{
    sumAoa.push([i+1, s.name, 1, "式", s.amount||0, ""]);
  });
  if(Number(cover.discount)!==0){
    sumAoa.push(["","値　引　き","","式",Number(cover.discount),"値引き"]);
  }
  sumAoa.push(["","","","","",""]);
  sumAoa.push(["","合　計（税　抜）","","",sub,""]);
  sumAoa.push(["","消費税（"+cover.taxRate+"%）","","",tax,""]);
  sumAoa.push(["","御見積金額（税込）","","",total,""]);

  const wsSum=XLSX.utils.aoa_to_sheet(sumAoa);
  setColW(wsSum,[6,28,8,8,16,16]);
  setRow(wsSum,0,24);
  XLSX.utils.book_append_sheet(wb,wsSum,"小計");

  // ── 各工事明細シート ──────────────────────────────────────────────────────
  sections.forEach(s=>{
    if(s.items.length===0)return;
    const tot=calcSec(s);
    const sheetAoa=[
      ["見　積　明　細　書　─　"+s.name,"","","","","","",""],
      [cover.projectName||"","","","","","","",""],
      ["","","","","","","",""],
      ["No.","項　目","仕様・摘要","数量","単位","単　価","金　額","備考"],
    ];
    s.items.forEach((it,i)=>{
      const c=calcItem(it);
      sheetAoa.push([
        i+1,
        it.name||"",
        it.spec||"",
        Number(it.qty)||0,
        it.unit||"式",
        Number(it.unitPrice)||0,
        c.amount||0,
        it.note||"",
      ]);
    });
    sheetAoa.push(["","","","","","","",""]);
    sheetAoa.push(["","小　計（税　抜）","","","","",tot.amount,""]);

    const ws=XLSX.utils.aoa_to_sheet(sheetAoa);
    setColW(ws,[5,22,18,8,6,14,14,16]);
    setRow(ws,0,22);
    const name=s.name.slice(0,31);
    XLSX.utils.book_append_sheet(wb,ws,name);
  });

  XLSX.writeFile(wb,`見積書_${cover.projectName||"STONA"}_${cover.estimateDate||"draft"}.xlsx`);
}

// ── Excel Import (SheetJS + Claude API) ──────────────────────────────────────
async function importExcel(file, setCover, setSections, setImportStatus){
  setImportStatus({state:"loading",msg:"ファイルを読み込み中..."});
  try{
    const buf=await file.arrayBuffer();
    const wb=XLSX.read(buf,{type:"array"});
    // 全シートをテキストで抽出
    const sheets={};
    wb.SheetNames.forEach(name=>{
      const ws=wb.Sheets[name];
      sheets[name]=XLSX.utils.sheet_to_csv(ws,{skipHidden:true});
    });
    const preview=Object.entries(sheets).map(([n,c])=>`=== シート: ${n} ===\n${c.slice(0,800)}`).join("\n\n");

    setImportStatus({state:"loading",msg:"Claude AIがデータを解析中..."});

    const prompt=`以下は建設工事の見積書Excelファイルのデータです。
このデータから見積情報を抽出し、必ず以下のJSON形式のみで返してください。他の文章は一切不要です。

JSON形式:
{
  "cover": {
    "estimateNo": "見積番号",
    "estimateDate": "YYYY-MM-DD",
    "client": "顧客名",
    "projectName": "件名・工事名",
    "location": "工事場所",
    "content": "工事内容",
    "paymentTerms": "支払条件",
    "validityPeriod": "有効期限",
    "constructionPeriod": "工期",
    "note": "備考",
    "taxRate": 10,
    "discount": 0
  },
  "sections": [
    {
      "name": "工事区分名（例：仮設工事）",
      "items": [
        {
          "name": "項目名",
          "spec": "仕様・摘要",
          "qty": 1,
          "unit": "式",
          "unitPrice": 50000,
          "note": "備考",
          "costUnitPrice": 0
        }
      ]
    }
  ]
}

Excelデータ:
${preview}`;

    const res=await fetch("https://api.anthropic.com/v1/messages",{
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify({
        model:"claude-sonnet-4-20250514",
        max_tokens:4000,
        messages:[{role:"user",content:prompt}]
      })
    });
    const data=await res.json();
    const text=data.content?.[0]?.text||"";
    const jsonMatch=text.match(/\{[\s\S]*\}/);
    if(!jsonMatch)throw new Error("JSONの抽出に失敗しました");
    const parsed=JSON.parse(jsonMatch[0]);

    if(parsed.cover) setCover(c=>({...c,...parsed.cover}));
    if(parsed.sections){
      setSections(parsed.sections.map(s=>({
        id:uid(), name:s.name||"工事区分",
        items:(s.items||[]).map(it=>({
          id:uid(), name:it.name||"", spec:it.spec||"",
          qty:it.qty||"", unit:it.unit||"式",
          unitPrice:it.unitPrice||"", note:it.note||"",
          costUnitPrice:it.costUnitPrice||""
        }))
      })));
    }
    setImportStatus({state:"success",msg:`✓ ${wb.SheetNames.length}シートを取り込みました`});
  }catch(e){
    setImportStatus({state:"error",msg:`エラー: ${e.message}`});
  }
}

// ── Output Modal ──────────────────────────────────────────────────────────────
function OutputModal({cover,sections,onClose}){
  const[tab,setTab]=useState("preview");
  const[importStatus,setImportStatus]=useState(null);
  const[setCoverExt,setSectionsExt]=[null,null]; // passed from App
  const fileRef=useRef();

  const tabs=[{k:"preview",l:"📄 印刷プレビュー"},{k:"excel",l:"📊 Excel出力"},{k:"import",l:"📥 Excelインポート"}];
  return(
    <div style={{position:"fixed",inset:0,zIndex:300,display:"flex",flexDirection:"column",background:"rgba(0,0,0,0.7)"}}>
      <div style={{background:"#fff",display:"flex",flexDirection:"column",flex:1,maxHeight:"96vh",margin:"2vh auto",width:"min(900px,96vw)",borderRadius:4,overflow:"hidden"}}>
        {/* ヘッダー */}
        <div style={{background:DARK,display:"flex",alignItems:"center",padding:"0 16px",gap:0,flexShrink:0}}>
          {tabs.map(t=>(
            <Btn key={t.k} style={{padding:"14px 18px",background:tab===t.k?COP:"transparent",color:tab===t.k?"#fff":"#888",fontSize:13,borderRadius:0}} onClick={()=>setTab(t.k)}>{t.l}</Btn>
          ))}
          <Btn style={{marginLeft:"auto",background:"none",color:"#888",fontSize:20,padding:"8px 14px"}} onClick={onClose}>×</Btn>
        </div>

        {/* コンテンツ */}
        <div style={{flex:1,overflowY:"auto",padding:"24px"}}>
          {tab==="preview"&&(
            <div>
              <div style={{display:"flex",justifyContent:"flex-end",marginBottom:16}}>
                <Btn style={{background:COP,color:"#fff",padding:"10px 24px",fontSize:14,display:"flex",alignItems:"center",gap:8}} onClick={()=>window.print()}>
                  🖨 このまま印刷 / PDF保存
                </Btn>
              </div>
              <div style={{border:`1px solid ${BORDER}`,background:"#fff"}}>
                <PrintDoc cover={cover} sections={sections} preview/>
              </div>
            </div>
          )}

          {tab==="excel"&&(
            <div style={{display:"flex",flexDirection:"column",gap:16,maxWidth:480}}>
              <div style={{background:BG,padding:"20px 24px",borderLeft:`4px solid ${COP}`}}>
                <div style={{fontFamily:"'Noto Serif JP',serif",fontSize:15,marginBottom:8}}>Excel出力について</div>
                <div style={{fontSize:13,color:"#555",lineHeight:1.7}}>
                  現在の見積データを複数シート構成のExcelファイルとして出力します。<br/>
                  構成：表紙 / 小計 / 各工事区分の明細シート
                </div>
              </div>
              <div style={{background:"#fff",border:`1px solid ${BORDER}`,padding:"20px 24px"}}>
                <div style={{fontSize:13,color:MUTED,marginBottom:12}}>ファイル名プレビュー</div>
                <div style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:14,color:DARK,background:BG,padding:"10px 14px"}}>
                  見積書_{cover.projectName||"STONA"}_{cover.estimateDate||"draft"}.xlsx
                </div>
              </div>
              <Btn style={{background:DARK,color:"#fff",padding:"14px 32px",fontSize:15,display:"flex",alignItems:"center",gap:10,alignSelf:"flex-start"}}
                onClick={()=>exportExcel(cover,sections)}>
                ⬇ Excelをダウンロード
              </Btn>
            </div>
          )}

          {tab==="import"&&(
            <div style={{display:"flex",flexDirection:"column",gap:16,maxWidth:560}}>
              <div style={{background:BG,padding:"20px 24px",borderLeft:`4px solid ${COP}`}}>
                <div style={{fontFamily:"'Noto Serif JP',serif",fontSize:15,marginBottom:8}}>Excelインポート（AIアシスト）</div>
                <div style={{fontSize:13,color:"#555",lineHeight:1.7}}>
                  既存の見積Excelファイルをアップロードすると、Claude AIが内容を解析して自動で項目を取り込みます。<br/>
                  フォーマットが違っても対応します。
                </div>
              </div>
              <div style={{background:"#fff",border:`2px dashed ${BORDER}`,padding:"36px 24px",textAlign:"center",cursor:"pointer"}}
                onClick={()=>fileRef.current?.click()}>
                <div style={{fontSize:32,marginBottom:8}}>📂</div>
                <div style={{fontSize:15,fontWeight:500,marginBottom:4}}>クリックしてExcelを選択</div>
                <div style={{fontSize:12,color:MUTED}}>.xlsx / .xls 対応</div>
                <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}}
                  onChange={e=>{
                    const f=e.target.files?.[0];
                    if(f) importExcel(f,
                      c=>{ /* cover更新はApp側でやる */ },
                      s=>{ /* sections更新はApp側でやる */ },
                      setImportStatus
                    );
                  }}
                />
              </div>
              {importStatus&&(
                <div style={{
                  padding:"14px 18px",background:"#fff",border:`1px solid ${BORDER}`,
                  borderLeft:`4px solid ${importStatus.state==="success"?"#2a8a4a":importStatus.state==="error"?"#cc4444":COP}`,
                  fontFamily:"'IBM Plex Mono',monospace",fontSize:13,
                  color:importStatus.state==="error"?"#cc4444":DARK,
                  display:"flex",alignItems:"center",gap:10
                }}>
                  {importStatus.state==="loading"&&<span style={{animation:"spin 1s linear infinite",display:"inline-block"}}>⟳</span>}
                  {importStatus.msg}
                </div>
              )}
              <div style={{fontSize:12,color:MUTED,lineHeight:1.7}}>
                ※ インポートすると現在のデータは上書きされます。<br/>
                ※ Claude APIを使用するため、インポート完了まで数秒かかります。
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ── App ───────────────────────────────────────────────────────────────────────
export default function App(){
  const isMobile=useIsMobile();
  const[view,setView]=useState("cover");
  const[cover,setCover]=useState(()=>{try{return JSON.parse(localStorage.getItem("stona_cover"))||INIT_COVER;}catch{return INIT_COVER;}});
  const[sections,setSections]=useState(()=>{try{return JSON.parse(localStorage.getItem("stona_sections"))||INIT_SECS;}catch{return INIT_SECS;}});
  const[activeSection,setActiveSection]=useState(null);
  const[showOutput,setShowOutput]=useState(false);
  const[importStatus,setImportStatus]=useState(null);
  const fileRef=useRef();

  useEffect(()=>{localStorage.setItem("stona_cover",JSON.stringify(cover));},[cover]);
  useEffect(()=>{localStorage.setItem("stona_sections",JSON.stringify(sections));},[sections]);

  const sec=sections.find(s=>s.id===activeSection);
  const title=view==="cover"?"表紙":view==="summary"?"見積内訳":sec?.name||"明細";
  const navItems=[{key:"cover",label:"表紙",icon:"📋"},{key:"summary",label:"中項目",icon:"📊"},...(activeSection&&sec?[{key:"detail",label:sec.name.slice(0,4),icon:"✏️"}]:[])];

  // インポートのセット関数をモーダルに渡す
  const handleImport=async(file)=>{
    await importExcel(file,setCover,setSections,setImportStatus);
  };

  return(
    <div style={{minHeight:"100vh",background:BG,fontFamily:"'Noto Sans JP',sans-serif",color:DARK}}>
      {/* Header */}
      <div className="no-print" style={{background:DARK,height:52,display:"flex",alignItems:"center",padding:"0 20px",borderBottom:`2px solid ${COP}`,position:"sticky",top:0,zIndex:40}}>
        <span style={{fontFamily:"'Noto Serif JP',serif",fontSize:16,color:COP,letterSpacing:"0.15em",marginRight:16}}>STONA</span>
        {isMobile
          ?<span style={{fontSize:15,fontWeight:500,color:"#fff"}}>{title}</span>
          :<div style={{display:"flex",gap:4,marginLeft:8}}>
            {["cover","summary"].map(k=>(
              <Btn key={k} style={{padding:"5px 18px",background:view===k?COP:"transparent",color:view===k?"#fff":"#aaa",fontSize:13}} onClick={()=>setView(k)}>
                {k==="cover"?"表紙":"中項目一覧"}
              </Btn>
            ))}
          </div>
        }
        <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center"}}>
          {/* インポートボタン（ヘッダー） */}
          <Btn style={{background:"none",border:"1px solid #444",color:"#aaa",padding:"5px 12px",fontSize:12,display:"flex",alignItems:"center",gap:5}}
            onClick={()=>{fileRef.current?.click();}}>
            📥 取込
          </Btn>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}}
            onChange={e=>{const f=e.target.files?.[0];if(f)handleImport(f);}}
          />
          {/* 出力ボタン */}
          <Btn style={{background:COP,color:"#fff",padding:"5px 16px",fontSize:12,display:"flex",alignItems:"center",gap:6}}
            onClick={()=>setShowOutput(true)}>
            🖨 出力
          </Btn>
        </div>
      </div>

      {/* インポートステータスバー */}
      {importStatus&&(
        <div className="no-print" style={{background:importStatus.state==="success"?"#e8f5ec":importStatus.state==="error"?"#fde8e8":"#fdf0e0",borderBottom:`1px solid ${BORDER}`,padding:"10px 20px",display:"flex",alignItems:"center",gap:12,fontSize:13}}>
          {importStatus.state==="loading"&&<span>⟳</span>}
          <span>{importStatus.msg}</span>
          {importStatus.state!=="loading"&&<Btn style={{marginLeft:"auto",background:"none",color:MUTED,fontSize:16,padding:"0 6px"}} onClick={()=>setImportStatus(null)}>×</Btn>}
        </div>
      )}

      {/* PC layout */}
      {!isMobile?(
        <div className="no-print" style={{display:"flex"}}>
          {view!=="cover"&&(
            <div style={{width:208,background:DARK,minHeight:"calc(100vh - 54px)",padding:"20px 0",flexShrink:0,position:"sticky",top:54,alignSelf:"flex-start",overflowY:"auto"}}>
              <div style={{fontSize:10,color:"#555",letterSpacing:"0.15em",fontFamily:"'IBM Plex Mono',monospace",padding:"0 18px 10px"}}>SECTIONS</div>
              {sections.map(s=>{
                const t=calcSec(s),active=view==="detail"&&activeSection===s.id;
                return(
                  <div key={s.id} onClick={()=>{setActiveSection(s.id);setView("detail");}}
                    style={{padding:"9px 18px",cursor:"pointer",borderLeft:active?`3px solid ${COP}`:"3px solid transparent",background:active?"rgba(193,123,47,0.12)":"transparent"}}>
                    <div style={{fontSize:13,color:active?"#E8A85F":"#999",fontWeight:active?500:300}}>{s.name}</div>
                    {t.amount>0&&<div style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:11,color:COP,marginTop:1}}>¥{fmt(t.amount)}</div>}
                  </div>
                );
              })}
              <div style={{borderTop:"1px solid #2a2a2a",margin:"12px 0 0",padding:"12px 18px 0"}}>
                <Btn style={{background:"none",border:"1px solid #333",color:"#777",width:"100%",padding:"8px 0",fontSize:12}} onClick={()=>setView("summary")}>← 一覧</Btn>
              </div>
            </div>
          )}
          <div style={{flex:1,padding:"36px 40px",maxWidth:1100}}>
            {view==="cover"&&<CoverPage cover={cover} setCover={setCover} isMobile={false}/>}
            {view==="summary"&&<SummaryPage sections={sections} setSections={setSections} cover={cover} setView={setView} setActiveSection={setActiveSection} isMobile={false}/>}
            {view==="detail"&&<DetailPage sections={sections} setSections={setSections} activeSection={activeSection} isMobile={false}/>}
          </div>
        </div>
      ):(
        <div className="no-print">
          <div style={{padding:"14px 14px",paddingBottom:80}}>
            {view==="cover"&&<CoverPage cover={cover} setCover={setCover} isMobile/>}
            {view==="summary"&&<SummaryPage sections={sections} setSections={setSections} cover={cover} setView={setView} setActiveSection={setActiveSection} isMobile/>}
            {view==="detail"&&<DetailPage sections={sections} setSections={setSections} activeSection={activeSection} isMobile/>}
          </div>
          <div style={{position:"fixed",bottom:0,left:0,right:0,height:60,background:DARK,borderTop:`2px solid ${COP}`,display:"flex",zIndex:40,paddingBottom:"env(safe-area-inset-bottom,0px)"}}>
            {navItems.map(n=>(
              <Btn key={n.key} style={{flex:1,background:"none",color:view===n.key?COP:"#666",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:2,fontSize:10,borderTop:view===n.key?`2px solid ${COP}`:"2px solid transparent",marginTop:-2,borderLeft:"none",borderRight:"none",borderBottom:"none"}}
                onClick={()=>setView(n.key)}>
                <span style={{fontSize:18}}>{n.icon}</span>
                <span>{n.label}</span>
              </Btn>
            ))}
            <Btn style={{flex:1,background:COP,color:"#fff",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:2,fontSize:10,borderTop:"2px solid transparent",marginTop:-2,borderLeft:"none",borderRight:"none",borderBottom:"none"}}
              onClick={()=>setShowOutput(true)}>
              <span style={{fontSize:18}}>🖨</span>
              <span>出力</span>
            </Btn>
          </div>
        </div>
      )}

      {/* 印刷用ドキュメント */}
      <div className="print-only">
        <PrintDoc cover={cover} sections={sections}/>
      </div>

      {/* 出力モーダル */}
      {showOutput&&(
        <OutputModal
          cover={cover} sections={sections}
          onClose={()=>setShowOutput(false)}
          onImport={handleImport}
          importStatus={importStatus}
          setImportStatus={setImportStatus}
        />
      )}
    </div>
  );
}
