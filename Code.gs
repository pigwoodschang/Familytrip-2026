const SHEET_NAME = "expenses";

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === "getAll")     return jsonOut(getAllExpenses());
  if (action === "getConfig")  return jsonOut(getConfig());
  if (action === "init")       return jsonOut(initSheet());
  if (action === "save") {
    const exp = JSON.parse(e.parameter.expense);
    return jsonOut(saveExpense(exp));
  }
  if (action === "delete")     return jsonOut(deleteExpense(e.parameter.id));
  if (action === "saveConfig") {
    const cfg = JSON.parse(e.parameter.config);
    return jsonOut(saveConfig(cfg));
  }
  // No action = serve the HTML app
  return HtmlService.createHtmlOutput(getHTML())
    .setTitle("家族旅遊記帳")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    if (action === "getAll")     return jsonOut(getAllExpenses());
    if (action === "getConfig")  return jsonOut(getConfig());
    if (action === "init")       return jsonOut(initSheet());
    if (action === "save")       return jsonOut(saveExpense(data.expense));
    if (action === "delete")     return jsonOut(deleteExpense(data.id));
    if (action === "saveConfig") return jsonOut(saveConfig(data.config));
    return jsonOut({ error: "unknown action" });
  } catch(err) {
    return jsonOut({ error: err.message });
  }
}

function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function initSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["id","title","amount","originalAmount","originalCurrency","category","paidBy","date","note","splitAmong"]);
    sheet.setFrozenRows(1);
  }
  let cfg = ss.getSheetByName("config");
  if (!cfg) {
    cfg = ss.insertSheet("config");
    cfg.appendRow(["key","value"]);
    cfg.appendRow(["tripName","家族旅遊"]);
    cfg.appendRow(["budget","300000"]);
    cfg.appendRow(["eurRate","35.5"]);
    cfg.appendRow(["members","爸爸,媽媽,孩子1,孩子2"]);
  }
  return { message: "initialized" };
}

function getAllExpenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(r => ({
    id: String(r[0]), title: r[1], amount: Number(r[2]),
    originalAmount: Number(r[3]), originalCurrency: r[4],
    category: r[5], paidBy: r[6], date: r[7], note: r[8],
    splitAmong: r[9] ? String(r[9]).split(",") : []
  }));
}

function saveExpense(exp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(exp.id)) {
      sheet.getRange(i+1,1,1,10).setValues([[
        exp.id, exp.title, exp.amount, exp.originalAmount,
        exp.originalCurrency, exp.category, exp.paidBy,
        exp.date, exp.note||"", (exp.splitAmong||[]).join(",")
      ]]);
      return { message: "updated" };
    }
  }
  sheet.appendRow([
    exp.id, exp.title, exp.amount, exp.originalAmount,
    exp.originalCurrency, exp.category, exp.paidBy,
    exp.date, exp.note||"", (exp.splitAmong||[]).join(",")
  ]);
  return { message: "saved" };
}

function deleteExpense(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i+1);
      return { message: "deleted" };
    }
  }
  return { message: "not found" };
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName("config");
  if (!cfg) return {};
  const rows = cfg.getDataRange().getValues();
  const result = {};
  rows.slice(1).forEach(r => { result[r[0]] = r[1]; });
  if (result.members) result.members = String(result.members).split(",");
  if (result.budget)  result.budget  = Number(result.budget);
  if (result.eurRate) result.eurRate = Number(result.eurRate);
  return result;
}

function saveConfig(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName("config");
  const rows = cfg.getDataRange().getValues();
  const updates = {
    ...config,
    members: Array.isArray(config.members) ? config.members.join(",") : config.members
  };
  Object.entries(updates).forEach(([key, value]) => {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === key) {
        cfg.getRange(i+1,2).setValue(value);
        return;
      }
    }
    cfg.appendRow([key, value]);
  });
  return { message: "config saved" };
}

function getHTML() {
  return `<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0">
<title>家族旅遊記帳</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Noto Sans TC','PingFang TC',sans-serif;background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%);min-height:100vh;color:#e8e8f0}
#app{max-width:480px;margin:0 auto;min-height:100vh;display:flex;flex-direction:column}
#toast{position:fixed;top:16px;left:50%;transform:translateX(-50%);background:rgba(20,20,40,0.96);border:1px solid rgba(255,255,255,0.15);border-radius:20px;padding:8px 20px;font-size:13px;z-index:999;white-space:nowrap;backdrop-filter:blur(10px);display:none}
#loading{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:100vh;gap:16px}
.lb{width:180px;height:4px;background:rgba(255,255,255,0.1);border-radius:2px;overflow:hidden}
.lbi{height:100%;width:55%;background:linear-gradient(90deg,#54a0ff,#5f27cd);border-radius:2px;animation:sl 1.2s ease-in-out infinite alternate}
@keyframes sl{from{transform:translateX(-60px)}to{transform:translateX(130px)}}
#header{padding:22px 20px 14px;background:rgba(255,255,255,0.04);border-bottom:1px solid rgba(255,255,255,0.08)}
.hrow{display:flex;justify-content:space-between;align-items:center}
.tsub{font-size:10px;letter-spacing:3px;color:#7f8fa6;text-transform:uppercase;margin-bottom:3px}
.tname{font-size:19px;font-weight:700;color:#fff}
.hbtns{display:flex;gap:8px;align-items:center}
.ibtn{background:rgba(255,255,255,0.08);border:none;color:#aaa;border-radius:10px;padding:7px 10px;cursor:pointer;font-size:15px}
.slbl{font-size:11px;color:#54a0ff}
.brow{display:flex;justify-content:space-between;font-size:12px;color:#7f8fa6;margin:14px 0 5px}
.bbg{height:7px;background:rgba(255,255,255,0.1);border-radius:4px;overflow:hidden}
.bfill{height:100%;border-radius:4px;transition:width 0.5s}
.brem{margin-top:5px;font-size:12px;text-align:right}
.badge{margin-top:10px;display:flex;align-items:center;gap:6px;font-size:11px;color:#7f8fa6}
.dot{width:7px;height:7px;border-radius:50%;background:#00d2d3;flex-shrink:0}
#nav{display:flex;background:rgba(0,0,0,0.3);border-bottom:1px solid rgba(255,255,255,0.06)}
.nbtn{flex:1;padding:10px 0;background:transparent;border:none;color:#7f8fa6;cursor:pointer;font-size:11px;border-bottom:2px solid transparent;transition:all 0.2s;font-family:inherit}
.nbtn.active{background:rgba(84,160,255,0.15);color:#54a0ff;border-bottom-color:#54a0ff}
#content{flex:1;padding:18px;overflow-y:auto}
.card{background:rgba(255,255,255,0.05);border-radius:16px;padding:15px;border:1px solid rgba(255,255,255,0.08);margin-bottom:16px}
.ctitle{font-size:12px;color:#7f8fa6;margin-bottom:12px}
.sgrid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px}
.scard{background:rgba(255,255,255,0.05);border-radius:14px;padding:14px;border:1px solid rgba(255,255,255,0.08)}
.slabel{font-size:11px;color:#7f8fa6;margin-bottom:5px}
.sval{font-size:18px;font-weight:700}
.crow{margin-bottom:11px}
.cinfo{display:flex;justify-content:space-between;margin-bottom:4px;font-size:13px}
.mbbg{height:5px;background:rgba(255,255,255,0.08);border-radius:3px;overflow:hidden}
.mbf{height:100%;border-radius:3px}
.ri{display:flex;align-items:center;gap:10px;padding:9px 0;border-bottom:1px solid rgba(255,255,255,0.05)}
.ci{width:34px;height:34px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:17px;flex-shrink:0}
.flbl{font-size:11px;color:#7f8fa6;margin-bottom:5px}
.fg{margin-bottom:14px}
input,select,textarea{width:100%;padding:11px 13px;background:rgba(255,255,255,0.07);border:1px solid rgba(255,255,255,0.12);border-radius:12px;color:#e8e8f0;font-size:14px;outline:none;font-family:inherit}
textarea{resize:none}
input::placeholder,textarea::placeholder{color:#7f8fa6}
.r2{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.ctog{display:flex;border-radius:12px;overflow:hidden;border:1px solid rgba(255,255,255,0.12);flex-shrink:0}
.cbtn{padding:0 14px;background:rgba(255,255,255,0.05);border:none;color:#aaa;cursor:pointer;font-size:13px;font-family:inherit}
.cbtn.active{background:rgba(84,160,255,0.3);color:#54a0ff;font-weight:700}
.cgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:7px}
.catbtn{padding:9px 4px;border-radius:10px;border:1px solid rgba(255,255,255,0.1);background:rgba(255,255,255,0.04);color:#aaa;cursor:pointer;font-size:12px;font-family:inherit;transition:all 0.2s}
.chips{display:flex;flex-wrap:wrap;gap:7px}
.chip{padding:6px 14px;border-radius:20px;border:1px solid rgba(255,255,255,0.1);background:transparent;color:#aaa;cursor:pointer;font-size:13px;font-family:inherit}
.chip.active{border-color:#54a0ff;background:rgba(84,160,255,0.2);color:#54a0ff}
.bprimary{width:100%;padding:13px;background:linear-gradient(90deg,#54a0ff,#5f27cd);border:none;border-radius:14px;color:#fff;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit}
.bprimary:disabled{background:rgba(84,160,255,0.4);cursor:not-allowed}
.bghost{width:100%;padding:11px;background:transparent;border:1px solid rgba(255,255,255,0.15);border-radius:14px;color:#aaa;font-size:13px;cursor:pointer;margin-top:9px;font-family:inherit}
.fwrap{display:flex;gap:7px;margin-bottom:14px;overflow-x:auto;padding-bottom:4px}
.fbtn{white-space:nowrap;padding:5px 12px;border-radius:20px;border:1px solid rgba(255,255,255,0.1);background:transparent;color:#aaa;cursor:pointer;font-size:12px;font-family:inherit}
.fbtn.active{border-color:#54a0ff;background:rgba(84,160,255,0.2);color:#54a0ff}
.eitem{background:rgba(255,255,255,0.04);border-radius:14px;padding:13px;margin-bottom:9px;border:1px solid rgba(255,255,255,0.07);display:flex;gap:10px;align-items:center}
.eicon{width:40px;height:40px;border-radius:11px;display:flex;align-items:center;justify-content:center;font-size:19px;flex-shrink:0}
.eacts{display:flex;gap:5px;margin-top:5px}
.bedit{background:rgba(84,160,255,0.15);border:none;color:#54a0ff;border-radius:6px;padding:3px 7px;cursor:pointer;font-size:12px}
.bdel{background:rgba(255,107,107,0.15);border:none;color:#ff6b6b;border-radius:6px;padding:3px 7px;cursor:pointer;font-size:12px}
.srow{display:flex;align-items:center;gap:8px;padding:12px 0;border-bottom:1px solid rgba(255,255,255,0.05)}
.sfrom{background:rgba(255,107,107,0.15);border-radius:8px;padding:5px 12px;font-size:13px;color:#ff6b6b;font-weight:600}
.sto{background:rgba(0,210,211,0.15);border-radius:8px;padding:5px 12px;font-size:13px;color:#00d2d3;font-weight:600}
.smid{flex:1;text-align:center}
.fxcard{background:linear-gradient(135deg,rgba(84,160,255,0.15),rgba(95,39,205,0.15));border-radius:16px;padding:16px;border:1px solid rgba(84,160,255,0.25);margin-bottom:18px;text-align:center}
.fxtog{display:flex;border-radius:14px;overflow:hidden;border:1px solid rgba(255,255,255,0.12);margin-bottom:16px}
.fxbtn{flex:1;padding:12px 8px;background:rgba(255,255,255,0.04);border:none;color:#aaa;cursor:pointer;font-size:14px;font-family:inherit}
.fxbtn.active{background:rgba(84,160,255,0.2);color:#54a0ff;font-weight:700}
.fxres{background:rgba(255,255,255,0.06);border-radius:16px;padding:20px;border:1px solid rgba(255,255,255,0.1);text-align:center;margin-bottom:20px}
.fxref{display:flex;justify-content:space-between;padding:9px 0;border-bottom:1px solid rgba(255,255,255,0.05);cursor:pointer}
.mrow{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.mrow input{flex:1;padding:8px 12px;font-size:13px}
.brem2{width:32px;height:32px;border-radius:8px;background:rgba(255,107,107,0.15);border:1px solid rgba(255,107,107,0.25);color:#ff6b6b;cursor:pointer;font-size:14px;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-family:inherit}
.badd{padding:8px 14px;border-radius:10px;background:rgba(84,160,255,0.2);border:1px solid rgba(84,160,255,0.35);color:#54a0ff;cursor:pointer;font-size:18px;flex-shrink:0;font-family:inherit}
.dzone{margin-top:22px;padding:13px;background:rgba(255,107,107,0.08);border-radius:14px;border:1px solid rgba(255,107,107,0.2)}
.empty{text-align:center;color:#7f8fa6;padding:30px 0;font-size:13px}
.hint{font-size:11px;color:#7f8fa6;margin-top:6px}
.ehint{font-size:11px;color:#7f8fa6;margin-top:5px;text-align:right}
</style>
</head>
<body>
<div id="toast"></div>
<div id="app">
  <div id="loading">
    <div style="font-size:44px">✈️</div>
    <div style="font-size:15px;color:#7f8fa6">載入中…</div>
    <div class="lb"><div class="lbi"></div></div>
  </div>
</div>
<script>
(function(){
const SCRIPT_URL = window.location.href.split('?')[0];
const CATS=[
  {id:"food",label:"🍜 餐飲",color:"#FF6B6B"},
  {id:"transport",label:"🚌 交通",color:"#4ECDC4"},
  {id:"hotel",label:"🏨 住宿",color:"#45B7D1"},
  {id:"ticket",label:"🎡 門票",color:"#96CEB4"},
  {id:"shopping",label:"🛍 購物",color:"#FFEAA7"},
  {id:"other",label:"📦 其他",color:"#DDA0DD"},
];
const DEF={tripName:"家族旅遊",budget:300000,eurRate:35.5,members:["爸爸","媽媽","孩子1","孩子2"]};
let config=JSON.parse(JSON.stringify(DEF));
let expenses=[];
let view="dashboard";
let filterCat="all";
let editId=null;
let syncing=false;
let convDir="TWD_TO_EUR";
let fxInput="";
let form=emptyForm();
let settingsMembers=[];

function emptyForm(){return{title:"",amount:"",currency:"TWD",category:"food",paidBy:"全家",date:today(),note:"",splitAmong:[]};}
function today(){return new Date().toISOString().split("T")[0];}
const fmt=n=>new Intl.NumberFormat("zh-TW",{style:"currency",currency:"TWD",minimumFractionDigits:0}).format(n);
const fmtEUR=n=>new Intl.NumberFormat("de-DE",{style:"currency",currency:"EUR",minimumFractionDigits:2}).format(n);
const getCat=id=>CATS.find(c=>c.id===id)||CATS[5];
const getMembers=()=>(config.members&&config.members.length>0)?config.members:DEF.members;
const esc=s=>String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");

// API — same-origin fetch, no CORS issues
async function api(params){
  const qs=Object.entries(params).map(([k,v])=>encodeURIComponent(k)+"="+encodeURIComponent(typeof v==="object"?JSON.stringify(v):v)).join("&");
  const res=await fetch(SCRIPT_URL+"?"+qs);
  return res.json();
}

let toastTimer;
function showToast(msg,ms=2500){
  const el=document.getElementById("toast");
  if(!el)return;
  el.textContent=msg;el.style.display="block";
  clearTimeout(toastTimer);
  toastTimer=setTimeout(()=>{el.style.display="none";},ms);
}
function setSyncing(v){
  syncing=v;
  const el=document.querySelector(".slbl");
  if(el)el.style.display=v?"inline":"none";
}

async function loadAll(){
  try{
    await api({action:"init"});
    const[cfg,exps]=await Promise.all([api({action:"getConfig"}),api({action:"getAll"})]);
    if(cfg&&!cfg.error)config={...DEF,...cfg};
    if(Array.isArray(exps))expenses=exps;
  }catch(e){showToast("⚠️ 載入失敗："+e.message,5000);}
  render();
}

function calcBalances(){
  const members=getMembers();
  const paid={},sp={};
  members.forEach(m=>{paid[m]=0;sp[m]=0;});
  expenses.forEach(exp=>{
    const among=exp.splitAmong&&exp.splitAmong.length>0?exp.splitAmong:members;
    const share=exp.amount/among.length;
    if(exp.paidBy==="全家")members.forEach(m=>{paid[m]+=exp.amount/members.length;});
    else if(members.includes(exp.paidBy))paid[exp.paidBy]+=exp.amount;
    else members.forEach(m=>{paid[m]+=exp.amount/members.length;});
    among.forEach(m=>{if(members.includes(m))sp[m]+=share;});
  });
  const b={};members.forEach(m=>{b[m]=(paid[m]||0)-(sp[m]||0);});
  return b;
}

function calcSettlements(bal){
  const cr=[],db=[];
  Object.entries(bal).forEach(([n,b])=>{if(b>0.5)cr.push({name:n,amount:b});else if(b<-0.5)db.push({name:n,amount:-b});});
  cr.sort((a,b)=>b.amount-a.amount);db.sort((a,b)=>b.amount-a.amount);
  const out=[];let ci=0,di=0;
  while(ci<cr.length&&di<db.length){
    const amt=Math.min(cr[ci].amount,db[di].amount);
    if(amt>0.5)out.push({from:db[di].name,to:cr[ci].name,amount:amt});
    cr[ci].amount-=amt;db[di].amount-=amt;
    if(cr[ci].amount<0.5)ci++;if(db[di].amount<0.5)di++;
  }
  return out;
}

function render(){
  const el=document.getElementById("loading");
  if(el)el.style.display="none";
  const members=getMembers();
  const total=expenses.reduce((s,e)=>s+e.amount,0);
  const rem=config.budget-total;
  const pct=Math.min((total/config.budget)*100,100);
  const bc=pct>90?"linear-gradient(90deg,#ff6b6b,#ee5a24)":pct>70?"linear-gradient(90deg,#ffd32a,#ff9f43)":"linear-gradient(90deg,#00d2d3,#54a0ff)";
  let html=\`
<div id="header">
  <div class="hrow">
    <div><div class="tsub">家族旅遊記帳</div><div class="tname">\${esc(config.tripName)}</div></div>
    <div class="hbtns">
      <span class="slbl" style="display:\${syncing?"inline":"none"}">同步中…</span>
      <button class="ibtn" id="btn-reload">🔄</button>
      <button class="ibtn" id="btn-settings">⚙️</button>
    </div>
  </div>
  <div class="brow"><span>已花費 \${fmt(total)}</span><span>預算 \${fmt(config.budget)}</span></div>
  <div class="bbg"><div class="bfill" style="width:\${pct}%;background:\${bc}"></div></div>
  <div class="brem"><span style="color:\${rem<0?"#ff6b6b":"#00d2d3"}">\${rem<0?"⚠ 超支 "+fmt(-rem):"剩餘 "+fmt(rem)}</span><span style="color:#7f8fa6;margin-left:8px">≈ \${fmtEUR(rem/config.eurRate)}</span></div>
  <div class="badge"><div class="dot"></div>已連線 Google Sheets・多裝置同步</div>
</div>
<div id="nav">
  \${[["dashboard","📊","總覽"],["add","➕","新增"],["list","📋","明細"],["split","⚖️","分帳"],["fx","💱","換算"]].map(([v,icon,label])=>\`<button class="nbtn \${view===v?"active":""}" data-view="\${v}"><div style="font-size:16px">\${icon}</div>\${label}</button>\`).join("")}
</div>
<div id="content">\`;
  if(view==="dashboard")html+=renderDashboard(total);
  else if(view==="add")html+=renderAdd(members);
  else if(view==="list")html+=renderList();
  else if(view==="split")html+=renderSplit(members);
  else if(view==="fx")html+=renderFX();
  else if(view==="settings")html+=renderSettings();
  html+=\`</div>\`;
  document.getElementById("app").innerHTML=html;
  attachEvents();
}

function renderDashboard(total){
  const cats=CATS.map(c=>({...c,t:expenses.filter(e=>e.category===c.id).reduce((s,e)=>s+e.amount,0)})).filter(c=>c.t>0).sort((a,b)=>b.t-a.t);
  const recent=[...expenses].sort((a,b)=>new Date(b.date)-new Date(a.date)).slice(0,5);
  return \`
<div class="sgrid">
  <div class="scard"><div class="slabel">總筆數</div><div class="sval" style="color:#54a0ff">\${expenses.length} 筆</div></div>
  <div class="scard"><div class="slabel">總花費(€)</div><div class="sval" style="color:#ffd32a">\${fmtEUR(total/config.eurRate)}</div></div>
</div>
<div class="card">
  <div class="ctitle">類別分析</div>
  \${cats.length===0?\`<div class="empty">還沒有記錄，開始記帳吧！</div>\`:cats.map(c=>\`
  <div class="crow">
    <div class="cinfo"><span>\${c.label}</span><div><span style="color:\${c.color};font-weight:600">\${fmt(c.t)}</span><span style="color:#7f8fa6;font-size:11px;margin-left:6px">≈\${fmtEUR(c.t/config.eurRate)}</span></div></div>
    <div class="mbbg"><div class="mbf" style="width:\${total>0?(c.t/total)*100:0}%;background:\${c.color}"></div></div>
  </div>\`).join("")}
</div>
<div class="card">
  <div class="ctitle">最近記錄</div>
  \${recent.length===0?\`<div class="empty">尚無記錄</div>\`:recent.map(e=>{const cat=getCat(e.category);return\`
  <div class="ri">
    <div class="ci" style="background:\${cat.color}22">\${cat.label.split(" ")[0]}</div>
    <div style="flex:1"><div style="font-size:13px;font-weight:500">\${esc(e.title)}</div><div style="font-size:11px;color:#7f8fa6">\${e.date} · \${esc(e.paidBy)}</div></div>
    <div style="text-align:right"><div style="color:\${cat.color};font-weight:700;font-size:13px">\${fmt(e.amount)}</div><div style="color:#7f8fa6;font-size:10px">\${fmtEUR(e.amount/config.eurRate)}</div></div>
  </div>\`;}).join("")}
</div>\`;}

function renderAdd(members){
  const eff=form.splitAmong.length>0?form.splitAmong:[...members];
  const n=Number(form.amount);
  const eh=form.amount&&!isNaN(n)?(form.currency==="TWD"?\`≈ \${fmtEUR(n/config.eurRate)}\`:\`≈ \${fmt(Math.round(n*config.eurRate))}\`):"";
  const pp=eff.length>0&&form.amount&&!isNaN(n)?Math.round((form.currency==="EUR"?n*config.eurRate:n)/eff.length):0;
  return \`
<div style="font-size:17px;font-weight:700;margin-bottom:18px">\${editId?"✏️ 編輯記錄":"➕ 新增支出"}</div>
<div class="fg"><div class="flbl">名稱</div><input id="f-title" value="\${esc(form.title)}" placeholder="例：羅浮宮門票"></div>
<div class="fg">
  <div class="flbl">金額</div>
  <div style="display:flex;gap:8px">
    <input id="f-amount" type="number" value="\${esc(form.amount)}" placeholder="0" style="flex:1">
    <div class="ctog">
      <button class="cbtn \${form.currency==="TWD"?"active":""}" data-cur="TWD">NT\$</button>
      <button class="cbtn \${form.currency==="EUR"?"active":""}" data-cur="EUR">€</button>
    </div>
  </div>
  <div class="ehint" id="eur-hint">\${eh}</div>
</div>
<div class="r2 fg">
  <div><div class="flbl">日期</div><input id="f-date" type="date" value="\${form.date}"></div>
  <div><div class="flbl">付款人</div><select id="f-paidby">\${["全家",...members].map(m=>\`<option\${m===form.paidBy?" selected":""}>\${esc(m)}</option>\`).join("")}</select></div>
</div>
<div class="fg">
  <div class="flbl">類別</div>
  <div class="cgrid">\${CATS.map(c=>\`<button class="catbtn" data-cat="\${c.id}" style="border-color:\${form.category===c.id?c.color:"rgba(255,255,255,0.1)"};background:\${form.category===c.id?c.color+"22":"rgba(255,255,255,0.04)"};color:\${form.category===c.id?c.color:"#aaa"}">\${c.label}</button>\`).join("")}</div>
</div>
<div class="fg">
  <div class="flbl">分攤對象</div>
  <div class="chips">\${members.map(m=>\`<button class="chip \${eff.includes(m)?"active":""}" data-member="\${esc(m)}">\${esc(m)}</button>\`).join("")}</div>
  <div class="hint" id="pp-hint">\${pp>0?\`每人負擔 ≈ \${fmt(pp)} \${fmtEUR(pp/config.eurRate)}\`:"每人平均分攤"}</div>
</div>
<div class="fg"><div class="flbl">備註（選填）</div><textarea id="f-note" rows="2" placeholder="備註…">\${esc(form.note)}</textarea></div>
<button class="bprimary" id="btn-submit" \${syncing?"disabled":""}>\${syncing?"同步中…":editId?"儲存修改":"新增記錄"}</button>
\${editId?\`<button class="bghost" id="btn-cancel">取消</button>\`:""}
\`;}

function renderList(){
  const filtered=[...expenses].filter(e=>filterCat==="all"||e.category===filterCat).sort((a,b)=>new Date(b.date)-new Date(a.date));
  return \`
<div class="fwrap">
  <button class="fbtn \${filterCat==="all"?"active":""}" data-filter="all">全部</button>
  \${CATS.map(c=>\`<button class="fbtn \${filterCat===c.id?"active":""}" data-filter="\${c.id}">\${c.label}</button>\`).join("")}
</div>
\${filtered.length===0?\`<div class="empty">沒有符合的記錄</div>\`:filtered.map(e=>{
  const cat=getCat(e.category);
  return \`<div class="eitem">
    <div class="eicon" style="background:\${cat.color}22">\${cat.label.split(" ")[0]}</div>
    <div style="flex:1;min-width:0">
      <div style="font-weight:600;font-size:14px">\${esc(e.title)}</div>
      <div style="font-size:11px;color:#7f8fa6;margin-top:2px">\${e.date} · \${esc(e.paidBy)} · \${cat.label.split(" ")[1]}</div>
      \${e.splitAmong&&e.splitAmong.length>0?\`<div style="font-size:10px;color:#54a0ff;margin-top:2px">分攤：\${e.splitAmong.map(esc).join("、")}</div>\`:""}
      \${e.note?\`<div style="font-size:11px;color:#aaa;margin-top:2px">\${esc(e.note)}</div>\`:""}
    </div>
    <div style="text-align:right;flex-shrink:0">
      <div style="color:\${cat.color};font-weight:700;font-size:14px">\${fmt(e.amount)}</div>
      <div style="color:#7f8fa6;font-size:10px">\${fmtEUR(e.amount/config.eurRate)}</div>
      \${e.originalCurrency==="EUR"?\`<div style="font-size:10px;color:#ffd32a">原 €\${e.originalAmount}</div>\`:""}
      <div class="eacts">
        <button class="bedit" data-id="\${e.id}">✏️</button>
        <button class="bdel" data-id="\${e.id}">🗑</button>
      </div>
    </div>
  </div>\`;}).join("")}\`;}

function renderSplit(members){
  const bal=calcBalances();const sets=calcSettlements(bal);
  return \`
<div style="font-size:17px;font-weight:700;margin-bottom:4px">⚖️ 分帳計算</div>
<div style="font-size:12px;color:#7f8fa6;margin-bottom:18px">依每筆支出的付款人與分攤對象自動計算</div>
<div class="card">
  <div class="ctitle">個人收支餘額</div>
  \${members.map(m=>{const b=bal[m]||0,isP=b>0.5,isM=b<-0.5;return\`
  <div style="display:flex;align-items:center;justify-content:space-between;padding:10px 0;border-bottom:1px solid rgba(255,255,255,0.05)">
    <div style="display:flex;align-items:center;gap:10px">
      <div style="width:36px;height:36px;border-radius:10px;background:\${isP?"rgba(0,210,211,0.15)":isM?"rgba(255,107,107,0.15)":"rgba(255,255,255,0.08)"};display:flex;align-items:center;justify-content:center;font-size:18px">\${m==="爸爸"?"👨":m==="媽媽"?"👩":"🧒"}</div>
      <div><div style="font-weight:600;font-size:14px">\${esc(m)}</div><div style="font-size:11px;color:#7f8fa6">\${isP?"應收回":isM?"應付出":"已結清"}</div></div>
    </div>
    <div style="text-align:right"><div style="font-weight:700;font-size:15px;color:\${isP?"#00d2d3":isM?"#ff6b6b":"#7f8fa6"}">\${isP?"+":""}\${fmt(Math.round(b))}</div><div style="font-size:11px;color:#7f8fa6">\${fmtEUR(b/config.eurRate)}</div></div>
  </div>\`;}).join("")}
</div>
<div class="card">
  <div class="ctitle">最佳結清方案</div>
  \${sets.length===0?\`<div style="text-align:center;padding:20px 0"><div style="font-size:28px;margin-bottom:8px">🎉</div><div style="color:#00d2d3;font-weight:600">已全部結清！</div></div>\`
  :sets.map(s=>\`<div class="srow"><div class="sfrom">\${esc(s.from)}</div><div class="smid"><div style="font-size:10px;color:#7f8fa6">支付給</div><div style="font-size:14px;font-weight:700;color:#ffd32a">\${fmt(Math.round(s.amount))}</div><div style="font-size:10px;color:#7f8fa6">\${fmtEUR(s.amount/config.eurRate)}</div></div><div class="sto">\${esc(s.to)}</div></div>\`).join("")}
</div>\`;}

function renderFX(){
  const inp=parseFloat(fxInput)||0;
  const res=inp?(convDir==="TWD_TO_EUR"?inp/config.eurRate:inp*config.eurRate):null;
  const nums=convDir==="TWD_TO_EUR"?[500,1000,3000,5000,10000,30000,50000]:[5,10,20,50,100,200,500];
  return \`
<div style="font-size:17px;font-weight:700;margin-bottom:4px">💱 匯率換算</div>
<div style="font-size:12px;color:#7f8fa6;margin-bottom:18px">新台幣（TWD）↔ 歐元（EUR）</div>
<div class="fxcard"><div style="font-size:11px;color:#7f8fa6;margin-bottom:6px">目前設定匯率</div><div style="font-size:28px;font-weight:800;color:#54a0ff">1 € = NT\$ \${config.eurRate}</div><div style="font-size:11px;color:#7f8fa6;margin-top:4px">可在設定中修改匯率</div></div>
<div class="fxtog">
  <button class="fxbtn \${convDir==="TWD_TO_EUR"?"active":""}" data-dir="TWD_TO_EUR"><div>NT\$ → €</div><div style="font-size:10px;margin-top:2px">新台幣換歐元</div></button>
  <button class="fxbtn \${convDir==="EUR_TO_TWD"?"active":""}" data-dir="EUR_TO_TWD"><div>€ → NT\$</div><div style="font-size:10px;margin-top:2px">歐元換新台幣</div></button>
</div>
<div class="fg"><div class="flbl">\${convDir==="TWD_TO_EUR"?"輸入新台幣金額 (NT\$)":"輸入歐元金額 (€)"}</div><input id="fx-input" type="number" placeholder="0" value="\${fxInput}" style="font-size:22px;text-align:center;font-weight:700"></div>
\${res!==null?\`<div class="fxres"><div style="font-size:12px;color:#7f8fa6;margin-bottom:8px">換算結果</div><div style="font-size:32px;font-weight:800;color:#ffd32a">\${convDir==="TWD_TO_EUR"?fmtEUR(res):fmt(Math.round(res))}</div><div style="font-size:12px;color:#7f8fa6;margin-top:6px">\${convDir==="TWD_TO_EUR"?\`NT\$ \${inp.toLocaleString()} ÷ \${config.eurRate} = €\${res.toFixed(2)}\`:\`€\${inp} × \${config.eurRate} = NT\$ \${Math.round(res).toLocaleString()}\`}</div></div>\`:""}
<div class="card"><div class="ctitle">常用換算參考 <span style="font-size:10px">(點擊帶入)</span></div>
\${nums.map(v=>\`<div class="fxref" data-fxval="\${v}"><span style="color:#aaa;font-size:14px">\${convDir==="TWD_TO_EUR"?\`NT\$ \${v.toLocaleString()}\`:\`€ \${v}\`}</span><span style="color:#ffd32a;font-weight:600;font-size:14px">\${convDir==="TWD_TO_EUR"?\`€ \${(v/config.eurRate).toFixed(2)}\`:\`NT\$ \${Math.round(v*config.eurRate).toLocaleString()}\`}</span></div>\`).join("")}
</div>\`;}

function renderSettings(){
  const members=settingsMembers.length>0?settingsMembers:[...getMembers()];
  return \`
<div style="font-size:17px;font-weight:700;margin-bottom:18px">⚙️ 旅遊設定</div>
<div class="fg"><div class="flbl">旅遊名稱</div><input id="s-name" value="\${esc(config.tripName)}" placeholder="例：歐洲家族旅遊"></div>
<div class="fg"><div class="flbl">總預算 (TWD)</div><input id="s-budget" type="number" value="\${config.budget}"></div>
<div class="fg"><div class="flbl">歐元匯率（1 EUR = ? TWD）</div><input id="s-rate" type="number" value="\${config.eurRate}"></div>
<div class="fg">
  <div class="flbl">家族成員</div>
  <div class="card" style="margin-bottom:0">
    <div id="members-list">
      \${members.map((m,i)=>\`<div class="mrow"><input class="s-member" data-idx="\${i}" value="\${esc(m)}"><button class="brem2" data-remove="\${i}">✕</button></div>\`).join("")}
    </div>
    <div style="display:flex;gap:8px;margin-top:4px">
      <input id="new-member" placeholder="輸入新成員名稱" style="flex:1;padding:8px 12px;font-size:13px">
      <button class="badd" id="btn-add-member">＋</button>
    </div>
    <div style="font-size:10px;color:#7f8fa6;margin-top:8px">共 \${members.length} 位成員・按 Enter 或 ＋ 新增</div>
  </div>
</div>
<button class="bprimary" id="btn-save-settings" \${syncing?"disabled":""}>\${syncing?"同步中…":"儲存並同步至 Sheets"}</button>
<div class="dzone">
  <div style="font-size:13px;color:#ff6b6b;margin-bottom:9px">⚠️ 危險區域</div>
  <button id="btn-clear-all" style="width:100%;padding:10px;background:rgba(255,107,107,0.15);border:1px solid rgba(255,107,107,0.3);border-radius:10px;color:#ff6b6b;font-size:13px;cursor:pointer;font-family:inherit">清除所有記錄</button>
</div>\`;}

function attachEvents(){
  document.querySelectorAll(".nbtn[data-view]").forEach(btn=>{
    btn.addEventListener("click",()=>{
      const v=btn.dataset.view;
      if(v==="add"){editId=null;form=emptyForm();}
      if(v==="settings")settingsMembers=[...getMembers()];
      view=v;render();
    });
  });
  const br=document.getElementById("btn-reload");
  if(br)br.addEventListener("click",()=>{render();loadAll();});
  const bs=document.getElementById("btn-settings");
  if(bs)bs.addEventListener("click",()=>{settingsMembers=[...getMembers()];view="settings";render();});

  if(view==="add"){
    const ft=document.getElementById("f-title");
    const fa=document.getElementById("f-amount");
    const fd=document.getElementById("f-date");
    const fp=document.getElementById("f-paidby");
    const fn=document.getElementById("f-note");
    if(ft)ft.addEventListener("input",e=>{form.title=e.target.value;});
    if(fa)fa.addEventListener("input",e=>{form.amount=e.target.value;updateHints();});
    if(fd)fd.addEventListener("change",e=>{form.date=e.target.value;});
    if(fp)fp.addEventListener("change",e=>{form.paidBy=e.target.value;});
    if(fn)fn.addEventListener("input",e=>{form.note=e.target.value;});
    document.querySelectorAll(".cbtn[data-cur]").forEach(b=>{
      b.addEventListener("click",()=>{form.currency=b.dataset.cur;document.querySelectorAll(".cbtn").forEach(x=>x.classList.toggle("active",x.dataset.cur===form.currency));updateHints();});
    });
    document.querySelectorAll(".catbtn[data-cat]").forEach(b=>{
      b.addEventListener("click",()=>{
        form.category=b.dataset.cat;
        document.querySelectorAll(".catbtn").forEach(x=>{const c=getCat(x.dataset.cat);const a=x.dataset.cat===form.category;x.style.borderColor=a?c.color:"rgba(255,255,255,0.1)";x.style.background=a?c.color+"22":"rgba(255,255,255,0.04)";x.style.color=a?c.color:"#aaa";});
      });
    });
    document.querySelectorAll(".chip[data-member]").forEach(b=>{
      b.addEventListener("click",()=>{
        const m=b.dataset.member;const members=getMembers();
        const eff=form.splitAmong.length>0?[...form.splitAmong]:[...members];
        form.splitAmong=eff.includes(m)?eff.filter(x=>x!==m):[...eff,m];
        document.querySelectorAll(".chip[data-member]").forEach(c=>{const en=form.splitAmong.length>0?form.splitAmong:members;c.classList.toggle("active",en.includes(c.dataset.member));});
        updateHints();
      });
    });
    const bsub=document.getElementById("btn-submit");
    if(bsub)bsub.addEventListener("click",submitForm);
    const bcan=document.getElementById("btn-cancel");
    if(bcan)bcan.addEventListener("click",()=>{editId=null;form=emptyForm();view="list";render();});
  }

  if(view==="list"){
    document.querySelectorAll(".fbtn[data-filter]").forEach(b=>{b.addEventListener("click",()=>{filterCat=b.dataset.filter;render();});});
    document.querySelectorAll(".bedit[data-id]").forEach(b=>{
      b.addEventListener("click",()=>{
        const e=expenses.find(x=>x.id===b.dataset.id);if(!e)return;
        form={...e,amount:String(e.originalAmount||e.amount),currency:e.originalCurrency||"TWD",splitAmong:e.splitAmong||[...getMembers()]};
        editId=e.id;view="add";render();
      });
    });
    document.querySelectorAll(".bdel[data-id]").forEach(b=>{b.addEventListener("click",()=>delExp(b.dataset.id));});
  }

  if(view==="fx"){
    document.querySelectorAll(".fxbtn[data-dir]").forEach(b=>{b.addEventListener("click",()=>{convDir=b.dataset.dir;fxInput="";render();});});
    const fi=document.getElementById("fx-input");
    if(fi)fi.addEventListener("input",e=>{fxInput=e.target.value;render();});
    document.querySelectorAll(".fxref[data-fxval]").forEach(r=>{r.addEventListener("click",()=>{fxInput=r.dataset.fxval;render();});});
  }

  if(view==="settings"){
    document.querySelectorAll(".s-member[data-idx]").forEach(inp=>{inp.addEventListener("input",e=>{settingsMembers[parseInt(inp.dataset.idx)]=e.target.value;});});
    document.querySelectorAll(".brem2[data-remove]").forEach(b=>{
      b.addEventListener("click",()=>{const i=parseInt(b.dataset.remove);if(settingsMembers.length<=1)return;settingsMembers.splice(i,1);render();});
    });
    const bam=document.getElementById("btn-add-member");
    const nm=document.getElementById("new-member");
    function addMember(){if(!nm||!nm.value.trim())return;settingsMembers.push(nm.value.trim());render();}
    if(bam)bam.addEventListener("click",addMember);
    if(nm)nm.addEventListener("keydown",e=>{if(e.key==="Enter")addMember();});
    const bss=document.getElementById("btn-save-settings");
    if(bss)bss.addEventListener("click",saveSettings);
    const bca=document.getElementById("btn-clear-all");
    if(bca)bca.addEventListener("click",clearAll);
  }
}

function updateHints(){
  const eh=document.getElementById("eur-hint");
  const pp=document.getElementById("pp-hint");
  const n=Number(form.amount);
  if(eh){eh.textContent=form.amount&&!isNaN(n)?(form.currency==="TWD"?\`≈ \${fmtEUR(n/config.eurRate)}\`:\`≈ \${fmt(Math.round(n*config.eurRate))}\`):"" ;}
  if(pp){const members=getMembers();const eff=form.splitAmong.length>0?form.splitAmong:members;const p=eff.length>0&&form.amount&&!isNaN(n)?Math.round((form.currency==="EUR"?n*config.eurRate:n)/eff.length):0;pp.textContent=p>\`每人負擔 ≈ \${fmt(p)} \${fmtEUR(p/config.eurRate)}\`:"每人平均分攤";}
}

async function submitForm(){
  const ft=document.getElementById("f-title");const fa=document.getElementById("f-amount");const fd=document.getElementById("f-date");const fp=document.getElementById("f-paidby");const fn=document.getElementById("f-note");
  const title=(ft?ft.value:form.title).trim();const amount=fa?fa.value:form.amount;const date=fd?fd.value:form.date;const paidBy=fp?fp.value:form.paidBy;const note=fn?fn.value:form.note;
  if(!title||!amount){showToast("⚠️ 請填寫名稱和金額");return;}
  const members=getMembers();let amtTWD=Number(amount);
  if(form.currency==="EUR")amtTWD=Math.round(amtTWD*config.eurRate);
  const entry={title,date,paidBy,note,amount:amtTWD,originalAmount:Number(amount),originalCurrency:form.currency,category:form.category,splitAmong:form.splitAmong.length>0?[...form.splitAmong]:[...members],id:editId||String(Date.now())};
  setSyncing(true);
  expenses=editId?expenses.map(e=>e.id===editId?entry:e):[...expenses,entry];
  form=emptyForm();editId=null;view="list";render();
  try{await api({action:"save",expense:JSON.stringify(entry)});showToast("✅ 已同步至 Sheets");}
  catch(e){showToast("⚠️ 同步失敗："+e.message);const exps=await api({action:"getAll"}).catch(()=>null);if(Array.isArray(exps))expenses=exps;}
  setSyncing(false);render();
}

async function delExp(id){
  if(!confirm("確定刪除這筆記錄？"))return;
  setSyncing(true);expenses=expenses.filter(e=>e.id!==id);render();
  try{await api({action:"delete",id});showToast("🗑 已刪除");}
  catch(e){showToast("⚠️ 刪除失敗："+e.message);}
  setSyncing(false);render();
}

async function saveSettings(){
  document.querySelectorAll(".s-member[data-idx]").forEach(inp=>{settingsMembers[parseInt(inp.dataset.idx)]=inp.value;});
  const name=document.getElementById("s-name")?.value;const budget=document.getElementById("s-budget")?.value;const rate=document.getElementById("s-rate")?.value;
  if(name)config.tripName=name;if(budget)config.budget=Number(budget);if(rate)config.eurRate=Number(rate);if(settingsMembers.length>0)config.members=[...settingsMembers];
  setSyncing(true);
  try{await api({action:"saveConfig",config:JSON.stringify(config)});showToast("✅ 設定已同步");}
  catch(e){showToast("⚠️ 同步失敗："+e.message);}
  setSyncing(false);settingsMembers=[];view="dashboard";render();
}

async function clearAll(){
  if(!confirm("確定要清除所有記錄嗎？"))return;
  setSyncing(true);
  for(const e of[...expenses]){await api({action:"delete",id:e.id}).catch(()=>{});}
  expenses=[];setSyncing(false);showToast("🗑 已清除所有記錄");view="dashboard";render();
}

loadAll();
})();
<\/script>
</body>
</html>`;
}
