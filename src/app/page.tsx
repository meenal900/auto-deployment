
"use client";
import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

const SITES = ["SBCZ","MBOX","MHYD","MDEA","SXVD","MPNA","MMAA","MAMA","DELU","NCRU"];
const SHIFTS = [
  { value: "Early Morning", label: "Early Morning (06:00-10:00)" },
  { value: "Day", label: "Day (10:00-20:00)" },
  { value: "Evening PT", label: "Evening PT (16:00-20:00)" },
  { value: "Night", label: "Night (20:00-06:00)" },
];
const PP_LIST = ["Inbound","Outbound","PS","ZBS","Sort Slide","Bagging","BigBox","Bag Pulling"];
const LC_OPTIONS = ["LC0","LC1","LC2","LC3","LC4","LC5"];
const FASTSTART_BASE = "https://admin.eu-west-1.faststart.ats.amazon.dev/#/DEFAULT";

export default function Home() {
  const [site, setSite] = useState("SBCZ");
  const [showAdmin, setShowAdmin] = useState(false);
  const [adminTab, setAdminTab] = useState<"upload"|"edit">("upload");
  const [adminFile, setAdminFile] = useState<File|null>(null);
  const [adminMsg, setAdminMsg] = useState("");
  const [adminError, setAdminError] = useState("");
  const [adminLoading, setAdminLoading] = useState(false);
  const [fileStatus, setFileStatus] = useState<any>(null);
  const [searchAlc, setSearchAlc] = useState("");
  const [alcResults, setAlcResults] = useState<any[]>([]);
  const [alcSearching, setAlcSearching] = useState(false);
  const [editRow, setEditRow] = useState<any>(null);
  const [editMsg, setEditMsg] = useState("");
  const [editError, setEditError] = useState("");
  const [addMode, setAddMode] = useState(false);
  const [newAA, setNewAA] = useState({alias:"",badgeID:"",employeeID:"",empName:"",primaryPP:PP_LIST[0],secondaryPP:"",empType:"FT",primaryLC:"LC0"});
  const [step, setStep] = useState(1);
  const [date, setDate] = useState("");
  const [shift, setShift] = useState("Day");
  const [alfaLogins, setAlfaLogins] = useState("");
  const [ppReqs, setPPReqs] = useState<Record<string,number>>(Object.fromEntries(PP_LIST.map(p=>[p,0])));
  const [scanResult, setScanResult] = useState<any>(null);
  const [deployResult, setDeployResult] = useState<any>(null);
  const [fclmFile, setFclmFile] = useState<File|null>(null);
  const [actualResult, setActualResult] = useState<any>(null);
  const [actualView, setActualView] = useState<"summary"|"badges"|"absent"|"audit">("summary");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  useEffect(()=>{checkFileStatus();},[site]);

  const checkFileStatus=async()=>{
    try{const fd=new FormData();fd.append("action","status");fd.append("site",site);
    const res=await fetch("/api/deploy",{method:"POST",body:fd});setFileStatus(await res.json());}catch(e){console.error(e);}
  };

  const downloadCSV=(data:any[],filename:string)=>{
    if(!data||data.length===0)return;const headers=Object.keys(data[0]);
    const csv=[headers.join(","),...data.map((r:any)=>headers.map(h=>`"${r[h]??""}"`).join(","))].join("\n");
    const blob=new Blob([csv],{type:"text/csv"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=filename;a.click();
  };

  const downloadFSK=()=>{
    if(!deployResult?.fskData)return;
    const sl=shift.replace(/\s+/g,"_");
    const wb=XLSX.utils.book_new();
    const wsData:any[][]=[];
    wsData.push(["BadgeId","EmployeeName","EmployeeLogin","EmployeeID","AssignedRole"]);
    for(const row of deployResult.fskData){
      wsData.push([row.BadgeId||"",row.EmployeeName||"",row.EmployeeLogin||"",row.EmployeeID||"",row.AssignedRole||""]);
    }
    const ws=XLSX.utils.aoa_to_sheet(wsData);
    ws["!cols"]=[{wch:14},{wch:25},{wch:16},{wch:14},{wch:35}];
    XLSX.utils.book_append_sheet(wb,ws,"FSK");
    const buf=XLSX.write(wb,{type:"array",bookType:"xlsx"});
    const blob=new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`${site}_FSK_${date}_${sl}.xlsx`;a.click();
  };

  const downloadFullReport=()=>{
    if(!actualResult)return;const sl=shift.replace(/\s+/g,"_");const lines:string[]=[];
    const t=actualResult.totals;const lc=t.lc||{};
    lines.push("=== SHIFT HEALTH REPORT ===",`Site,${site}`,`Date,${date}`,`Shift,${shift}`,
      `Total HC,${t.totalHC}`,`Planned,${t.planned}`,`Present,${t.present}`,`Absent,${t.absent}`,
      `Extra,${t.extra}`,`FT (Full Time),${t.totalFT}`,`PT (Part Time),${t.totalPT}`,`Attendance,${t.attendance}`,"");
    lines.push("=== LC WISE BREAKDOWN ===","LC5,LC4,LC3,LC2,LC1,LC0",
      `${lc.LC5||0},${lc.LC4||0},${lc.LC3||0},${lc.LC2||0},${lc.LC1||0},${lc.LC0||0}`,"");
    lines.push("=== PP WISE SUMMARY ===","PP,Required,Deployed,FT,PT,Present,Extra,Absent,Shortage,Overflow,LC5,LC4,LC3,LC2,LC1,LC0");
    let tR=0,tD=0,tF=0,tP=0,tPr=0,tE=0,tA=0,tS=0,tO=0;
    const tL:Record<string,number>={LC5:0,LC4:0,LC3:0,LC2:0,LC1:0,LC0:0};
    for(const[pp,s]of Object.entries(actualResult.summary)as any){
      lines.push(`${pp},${s.required},${s.deployed},${s.ftDeployed},${s.ptDeployed},${s.present},${s.extra},${s.absent},${s.shortage},${s.overflow},${s.lc?.LC5||0},${s.lc?.LC4||0},${s.lc?.LC3||0},${s.lc?.LC2||0},${s.lc?.LC1||0},${s.lc?.LC0||0}`);
      tR+=s.required;tD+=s.deployed;tF+=s.ftDeployed;tP+=s.ptDeployed;tPr+=s.present;tE+=s.extra;tA+=s.absent;tS+=s.shortage;tO+=s.overflow;
      for(const l of Object.keys(tL))tL[l]+=(s.lc?.[l]||0);
    }
    lines.push(`TOTAL,${tR},${tD},${tF},${tP},${tPr},${tE},${tA},${tS},${tO},${tL.LC5},${tL.LC4},${tL.LC3},${tL.LC2},${tL.LC1},${tL.LC0}`,"");
    lines.push("=== PP WISE BADGE IDs ===");
    for(const[pp,badges]of Object.entries(actualResult.badgesByPP)as any){
      if(!badges||badges.length===0)continue;
      lines.push(`--- ${pp} (${badges.length}) ---`,"Badge ID");
      for(const b of badges)lines.push(b);lines.push("");
    }
    lines.push("=== ABSENT LIST ===");
    if(actualResult.absentList.length>0){lines.push("Badge ID,Emp ID,Name,Alias,Planned PP,LC");
      for(const a of actualResult.absentList)lines.push(`${a.badgeID},${a.empID},"${a.name}","${a.alias}",${a.plannedPP},${a.lc||""}`);}
    else lines.push("No absent employees");lines.push("");
    if(actualResult.skippedNotInALC&&actualResult.skippedNotInALC.length>0){
      lines.push("=== SKIPPED (NOT IN ALC MAPPING) ===","Badge/ID");
      for(const s of actualResult.skippedNotInALC)lines.push(s);lines.push("");
    }
    lines.push("=== AUDIT REPORT ===","Source,Emp ID,Alias,Badge ID,Status,Reason,Primary PP,Assigned PP,FT/PT,LC");
    for(const e of actualResult.auditLog)lines.push(`${e.inFSK==="YES"?"Planned":"Extra"},${e.empID},"${e.alias}",${e.badgeID},${e.status},"${e.reason}",${e.primaryPP},${e.assignedPP},${e.empType},${e.lc||""}`);
    const blob=new Blob([lines.join("\n")],{type:"text/csv"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);
    a.download=`${site}_${date}_${sl}_Shift_Health_Report.csv`;a.click();
  };

  const handleAdminUpload=async()=>{
    if(!adminFile){setAdminError("Select a file");return;}setAdminError("");setAdminMsg("");setAdminLoading(true);
    const fd=new FormData();fd.append("action","admin-upload");fd.append("site",site);fd.append("file",adminFile);
    try{const res=await fetch("/api/deploy",{method:"POST",body:fd});const data=await res.json();if(!res.ok)throw new Error(data.error);
    setAdminMsg(`Uploaded! ALC: ${data.alcRows} rows | Roster: ${data.rosterRows} rows`);checkFileStatus();}catch(e:any){setAdminError(e.message);}finally{setAdminLoading(false);}
  };

  const handleSearchALC=async()=>{
    if(!searchAlc.trim())return;setAlcSearching(true);setAlcResults([]);setEditMsg("");setEditError("");
    const fd=new FormData();fd.append("action","search-alc");fd.append("site",site);fd.append("query",searchAlc.trim());
    try{const res=await fetch("/api/deploy",{method:"POST",body:fd});const data=await res.json();if(!res.ok)throw new Error(data.error);setAlcResults(data.results||[]);}catch(e:any){setEditError(e.message);}finally{setAlcSearching(false);}
  };

  const handleSaveEdit=async()=>{
    if(!editRow)return;setEditMsg("");setEditError("");
    const fd=new FormData();fd.append("action","edit-alc");fd.append("site",site);fd.append("data",JSON.stringify(editRow));
    try{const res=await fetch("/api/deploy",{method:"POST",body:fd});const data=await res.json();if(!res.ok)throw new Error(data.error);
    setEditMsg("Saved!");setEditRow(null);if(searchAlc)handleSearchALC();}catch(e:any){setEditError(e.message);}
  };

  const handleAddAA=async()=>{
    if(!newAA.alias&&!newAA.badgeID){setEditError("Enter alias or badge ID");return;}setEditMsg("");setEditError("");
    const fd=new FormData();fd.append("action","add-alc");fd.append("site",site);fd.append("data",JSON.stringify(newAA));
    try{const res=await fetch("/api/deploy",{method:"POST",body:fd});const data=await res.json();if(!res.ok)throw new Error(data.error);
    setEditMsg("AA added!");setNewAA({alias:"",badgeID:"",employeeID:"",empName:"",primaryPP:PP_LIST[0],secondaryPP:"",empType:"FT",primaryLC:"LC0"});setAddMode(false);}catch(e:any){setEditError(e.message);}
  };

  const handleScan=async()=>{
    if(!date||!shift){setError("Select date and shift");return;}setError("");setLoading(true);
    const fd=new FormData();fd.append("action","scan");fd.append("date",date);fd.append("shift",shift);fd.append("alfaLogins",alfaLogins);fd.append("site",site);
    try{const res=await fetch("/api/deploy",{method:"POST",body:fd});const data=await res.json();if(!res.ok)throw new Error(data.error);setScanResult(data);setStep(2);}catch(e:any){setError(e.message);}finally{setLoading(false);}
  };

  const handleDeploy=async()=>{
    setLoading(true);setError("");const fd=new FormData();fd.append("action","deploy");fd.append("date",date);fd.append("shift",shift);
    fd.append("alfaLogins",alfaLogins);fd.append("ppRequirements",JSON.stringify(ppReqs));fd.append("site",site);
    try{const res=await fetch("/api/deploy",{method:"POST",body:fd});const data=await res.json();if(!res.ok)throw new Error(data.error);setDeployResult(data);setStep(3);}catch(e:any){setError(e.message);}finally{setLoading(false);}
  };

  const handleActual=async()=>{
    if(!fclmFile){setError("Upload FCLM Punch file");return;}setLoading(true);setError("");
    const fd=new FormData();fd.append("action","actual");fd.append("fclmFile",fclmFile);fd.append("fskData",JSON.stringify(deployResult?.fskData||[]));
    fd.append("ppRequirements",JSON.stringify(ppReqs));fd.append("date",date);fd.append("shift",shift);fd.append("site",site);
    try{const res=await fetch("/api/deploy",{method:"POST",body:fd});const data=await res.json();if(!res.ok)throw new Error(data.error);setActualResult(data);setStep(5);}catch(e:any){setError(e.message);}finally{setLoading(false);}
  };

  const handleNewDeploy=()=>{setStep(1);setScanResult(null);setDeployResult(null);setActualResult(null);setAlfaLogins("");setFclmFile(null);setError("");setPPReqs(Object.fromEntries(PP_LIST.map(p=>[p,0])));};
  const alfaCount=alfaLogins.split(/[\n,;]+/).filter(s=>s.trim()).length;
  const B=(bg:string,c="#fff"):React.CSSProperties=>({padding:"12px 24px",fontSize:15,background:bg,color:c,border:"none",borderRadius:8,cursor:"pointer",fontWeight:"bold"});
  const C:React.CSSProperties={border:"1px solid #ddd",borderRadius:12,padding:24,marginBottom:20,background:"#fff"};
  const I:React.CSSProperties={padding:8,fontSize:14,borderRadius:6,border:"1px solid #ddd",boxSizing:"border-box" as any};
  const SB=(a:boolean,d:boolean):React.CSSProperties=>({width:32,height:32,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:"bold",fontSize:14,color:"#fff",background:d?"#4caf50":a?"#ff9900":"#ccc"});

  return (
    <div style={{maxWidth:1200,margin:"0 auto",padding:"16px 24px",fontFamily:"Arial"}}>
      {/* HEADER */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,borderBottom:"2px solid #ff9900",paddingBottom:12}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <h1 style={{color:"#232f3e",margin:0,fontSize:22}}>🚀 Auto Deployment Tool</h1>
          <select value={site} onChange={e=>setSite(e.target.value)} style={{padding:"6px 12px",fontSize:16,fontWeight:"bold",borderRadius:8,border:"2px solid #ff9900",background:"#fff",color:"#232f3e",cursor:"pointer"}}>
            {SITES.map(s=><option key={s} value={s}>{s}</option>)}
          </select>
        </div>
        <button onClick={()=>{setShowAdmin(!showAdmin);setAdminTab("upload");}} style={{...B(showAdmin?"#f44336":"#232f3e"),padding:"8px 16px",fontSize:13}}>
          {showAdmin?"Close Admin":"Admin Panel"}
        </button>
      </div>

      {fileStatus&&<div style={{background:fileStatus.alcMapping?"#e8f5e9":"#fff3cd",padding:"8px 16px",borderRadius:8,marginBottom:16,fontSize:13,display:"flex",justifyContent:"space-between"}}>
        <span>{fileStatus.alcMapping?`Data loaded for ${site} (ALC: ${fileStatus.alcCount} | Roster: ${fileStatus.rosterCount})`:`No data for ${site} — Upload via Admin Panel`}</span>
        {fileStatus.lastUpdated&&<span style={{color:"#888"}}>Updated: {new Date(fileStatus.lastUpdated).toLocaleString()}</span>}
      </div>}

      {/* ADMIN PANEL */}
      {showAdmin&&<div style={{...C,border:"2px solid #ff9900"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <h2 style={{margin:0,color:"#232f3e"}}>Admin — {site}</h2>
          <div style={{display:"flex",gap:6}}>
            {(["upload","edit"] as const).map(t=><button key={t} onClick={()=>{setAdminTab(t);setEditMsg("");setEditError("");}} style={{padding:"6px 16px",background:adminTab===t?"#ff9900":"#eee",color:adminTab===t?"#fff":"#333",border:"none",borderRadius:6,cursor:"pointer",fontWeight:"bold",fontSize:13}}>{t==="upload"?"Upload Data":"Edit / Add AA"}</button>)}
          </div>
        </div>
        {adminTab==="upload"&&<div>
          <p style={{color:"#666",fontSize:14}}>Upload Excel with: <strong>ALC Mapping</strong> and <strong>Roster_Dump</strong></p>
          {adminError&&<div style={{background:"#fee",border:"1px solid #f66",padding:10,borderRadius:8,marginBottom:12,color:"red"}}>{adminError}</div>}
          {adminMsg&&<div style={{background:"#e8f5e9",border:"1px solid #4caf50",padding:10,borderRadius:8,marginBottom:12}}>{adminMsg}</div>}
          <div style={{display:"flex",gap:12,alignItems:"center"}}>
            <input type="file" accept=".xlsx,.xls" onChange={e=>setAdminFile(e.target.files?.[0]||null)}/>
            <button onClick={handleAdminUpload} disabled={adminLoading} style={B("#ff9900")}>{adminLoading?"Uploading...":"Upload"}</button>
          </div>
        </div>}
        {adminTab==="edit"&&<div>
          {editMsg&&<div style={{background:"#e8f5e9",border:"1px solid #4caf50",padding:10,borderRadius:8,marginBottom:12}}>{editMsg}</div>}
          {editError&&<div style={{background:"#fee",border:"1px solid #f66",padding:10,borderRadius:8,marginBottom:12,color:"red"}}>{editError}</div>}
          <div style={{display:"flex",gap:8,marginBottom:16}}>
            <input type="text" placeholder="Search by Alias, Badge ID, or Name..." value={searchAlc} onChange={e=>setSearchAlc(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleSearchALC()} style={{...I,flex:1}}/>
            <button onClick={handleSearchALC} disabled={alcSearching} style={B("#232f3e")}>{alcSearching?"...":"Search"}</button>
            <button onClick={()=>{setAddMode(true);setEditRow(null);setEditMsg("");setEditError("");}} style={B("#4caf50")}>+ Add AA</button>
          </div>
          {alcResults.length>0&&!editRow&&!addMode&&<table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:"#232f3e",color:"#fff"}}>{["Alias","Badge","Emp ID","Name","Primary PP","Sec PP","LC","FT/PT",""].map(h=><th key={h} style={{padding:6,border:"1px solid #555"}}>{h}</th>)}</tr></thead>
            <tbody>{alcResults.map((r:any,i:number)=><tr key={i}>
              <td style={{padding:5,border:"1px solid #eee"}}>{r.alias}</td><td style={{padding:5,border:"1px solid #eee"}}>{r.badgeID}</td>
              <td style={{padding:5,border:"1px solid #eee"}}>{r.employeeID}</td><td style={{padding:5,border:"1px solid #eee"}}>{r.empName}</td>
              <td style={{padding:5,border:"1px solid #eee"}}>{r.primaryPP}</td><td style={{padding:5,border:"1px solid #eee"}}>{r.secondaryPP}</td>
              <td style={{padding:5,border:"1px solid #eee",fontWeight:"bold",color:"#1565c0"}}>{r.primaryLC}</td><td style={{padding:5,border:"1px solid #eee"}}>{r.empType}</td>
              <td style={{padding:5,border:"1px solid #eee",textAlign:"center"}}><button onClick={()=>{setEditRow({...r,originalAlias:r.alias,originalBadge:r.badgeID,originalEmpID:r.employeeID});setAddMode(false);setEditMsg("");setEditError("");}} style={{...B("#ff9900"),padding:"4px 12px",fontSize:12}}>Edit</button></td>
            </tr>)}</tbody>
          </table>}
          {alcResults.length===0&&searchAlc&&!alcSearching&&!addMode&&!editRow&&<p style={{color:"#888",textAlign:"center"}}>No results</p>}
          {editRow&&!addMode&&<div style={{background:"#f8f9fa",border:"1px solid #ddd",borderRadius:8,padding:16,marginTop:12}}>
            <h3 style={{marginTop:0,color:"#232f3e"}}>Edit: {editRow.originalAlias||editRow.alias||editRow.badgeID}</h3>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
              {[{k:"alias",l:"Alias"},{k:"badgeID",l:"Badge ID"},{k:"employeeID",l:"Employee ID"},{k:"empName",l:"Name"}].map(f=><div key={f.k}><label style={{fontSize:12,fontWeight:"bold"}}>{f.l}</label><input value={editRow[f.k]||""} onChange={e=>setEditRow({...editRow,[f.k]:e.target.value})} style={{...I,width:"100%",display:"block",marginTop:4}}/></div>)}
              <div><label style={{fontSize:12,fontWeight:"bold"}}>Primary PP</label><select value={editRow.primaryPP||""} onChange={e=>setEditRow({...editRow,primaryPP:e.target.value})} style={{...I,width:"100%",display:"block",marginTop:4}}><option value="">--</option>{PP_LIST.map(p=><option key={p} value={p}>{p}</option>)}</select></div>
              <div><label style={{fontSize:12,fontWeight:"bold"}}>Secondary PP</label><select value={editRow.secondaryPP||""} onChange={e=>setEditRow({...editRow,secondaryPP:e.target.value})} style={{...I,width:"100%",display:"block",marginTop:4}}><option value="">None</option>{PP_LIST.map(p=><option key={p} value={p}>{p}</option>)}</select></div>
              <div><label style={{fontSize:12,fontWeight:"bold"}}>LC Level</label><select value={editRow.primaryLC||"LC0"} onChange={e=>setEditRow({...editRow,primaryLC:e.target.value})} style={{...I,width:"100%",display:"block",marginTop:4}}>{LC_OPTIONS.map(l=><option key={l} value={l}>{l}</option>)}</select></div>
              <div><label style={{fontSize:12,fontWeight:"bold"}}>FT / PT</label><select value={editRow.empType||"FT"} onChange={e=>setEditRow({...editRow,empType:e.target.value})} style={{...I,width:"100%",display:"block",marginTop:4}}><option value="FT">FT</option><option value="PT">PT</option></select></div>
            </div>
            <div style={{display:"flex",gap:8,marginTop:16}}><button onClick={handleSaveEdit} style={B("#4caf50")}>Save</button><button onClick={()=>setEditRow(null)} style={B("#999")}>Cancel</button></div>
          </div>}
          {addMode&&!editRow&&<div style={{background:"#f0fff0",border:"1px solid #4caf50",borderRadius:8,padding:16,marginTop:12}}>
            <h3 style={{marginTop:0,color:"#4caf50"}}>Add New AA</h3>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
              {[{k:"alias",l:"Alias *"},{k:"badgeID",l:"Badge ID"},{k:"employeeID",l:"Employee ID"},{k:"empName",l:"Name"}].map(f=><div key={f.k}><label style={{fontSize:12,fontWeight:"bold"}}>{f.l}</label><input value={(newAA as any)[f.k]} onChange={e=>setNewAA({...newAA,[f.k]:e.target.value})} style={{...I,width:"100%",display:"block",marginTop:4}}/></div>)}
              <div><label style={{fontSize:12,fontWeight:"bold"}}>Primary PP *</label><select value={newAA.primaryPP} onChange={e=>setNewAA({...newAA,primaryPP:e.target.value})} style={{...I,width:"100%",display:"block",marginTop:4}}>{PP_LIST.map(p=><option key={p} value={p}>{p}</option>)}</select></div>
              <div><label style={{fontSize:12,fontWeight:"bold"}}>Secondary PP</label><select value={newAA.secondaryPP} onChange={e=>setNewAA({...newAA,secondaryPP:e.target.value})} style={{...I,width:"100%",display:"block",marginTop:4}}><option value="">None</option>{PP_LIST.map(p=><option key={p} value={p}>{p}</option>)}</select></div>
              <div><label style={{fontSize:12,fontWeight:"bold"}}>LC Level</label><select value={newAA.primaryLC} onChange={e=>setNewAA({...newAA,primaryLC:e.target.value})} style={{...I,width:"100%",display:"block",marginTop:4}}>{LC_OPTIONS.map(l=><option key={l} value={l}>{l}</option>)}</select></div>
              <div><label style={{fontSize:12,fontWeight:"bold"}}>FT / PT</label><select value={newAA.empType} onChange={e=>setNewAA({...newAA,empType:e.target.value})} style={{...I,width:"100%",display:"block",marginTop:4}}><option value="FT">FT</option><option value="PT">PT</option></select></div>
            </div>
            <div style={{display:"flex",gap:8,marginTop:16}}><button onClick={handleAddAA} style={B("#4caf50")}>Add AA</button><button onClick={()=>setAddMode(false)} style={B("#999")}>Cancel</button></div>
          </div>}
        </div>}
      </div>}

      {/* DEPLOYMENT FLOW */}
      {!showAdmin&&<>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:24}}>
          {[{n:1,l:"Shift & Alfa OT"},{n:2,l:"PP Requirements"},{n:3,l:"Summary & FSK"},{n:4,l:"FCLM Punch"},{n:5,l:"Shift Health Report"}].map((s,i)=>
            <div key={s.n} style={{display:"flex",alignItems:"center",gap:6}}>
              <div style={SB(step===s.n,step>s.n)}>{step>s.n?"✓":s.n}</div>
              <span style={{fontSize:13,fontWeight:step===s.n?"bold":"normal",color:step===s.n?"#ff9900":step>s.n?"#4caf50":"#999"}}>{s.l}</span>
              {i<4&&<div style={{width:30,height:2,background:step>s.n?"#4caf50":"#ddd"}}/>}
            </div>)}
        </div>
        {error&&<div style={{background:"#fee",border:"1px solid #f66",padding:12,borderRadius:8,marginBottom:16,color:"red"}}>{error}</div>}

        {step===1&&<div style={{...C,maxWidth:600,margin:"0 auto"}}>
          <h2 style={{marginTop:0,color:"#232f3e"}}>Step 1: Shift and Alfa/OT Logins</h2>
          <div style={{display:"flex",gap:16,marginBottom:16}}>
            <div style={{flex:1}}><label style={{fontWeight:"bold",fontSize:14}}>Date:</label><input type="date" value={date} onChange={e=>setDate(e.target.value)} style={{display:"block",width:"100%",padding:10,fontSize:15,borderRadius:8,border:"1px solid #ddd",boxSizing:"border-box",marginTop:4}}/></div>
            <div style={{flex:1}}><label style={{fontWeight:"bold",fontSize:14}}>Shift:</label><select value={shift} onChange={e=>setShift(e.target.value)} style={{display:"block",width:"100%",padding:10,fontSize:15,borderRadius:8,border:"1px solid #ddd",marginTop:4}}>{SHIFTS.map(s=><option key={s.value} value={s.value}>{s.label}</option>)}</select></div>
          </div>
          <label style={{fontWeight:"bold",fontSize:14}}>Alfa / OT Login IDs:</label>
          <textarea value={alfaLogins} onChange={e=>setAlfaLogins(e.target.value)} placeholder="one per line or comma separated" rows={5} style={{display:"block",width:"100%",padding:10,marginTop:4,marginBottom:4,fontSize:14,borderRadius:8,border:"1px solid #ddd",fontFamily:"monospace",boxSizing:"border-box",resize:"vertical"}}/>
          <p style={{color:"#888",fontSize:12,margin:"0 0 16px 0"}}>{alfaCount>0?`${alfaCount} login(s)`:"No logins"}</p>
          <button onClick={handleScan} disabled={loading} style={{...B("#ff9900"),width:"100%"}}>{loading?"Scanning...":"Next: Scan Roster and Alfa"}</button>
        </div>}

        {step===2&&scanResult&&<div style={{...C,maxWidth:600,margin:"0 auto"}}>
          <h2 style={{marginTop:0,color:"#232f3e"}}>Step 2: PP-wise Headcount</h2>
          <div style={{background:"#e8f5e9",padding:12,borderRadius:8,marginBottom:8}}>
            <strong style={{fontSize:18}}>Available: {scanResult.totalAvailable}</strong>
            <span style={{marginLeft:16,color:"#666"}}>Roster: {scanResult.rosterCount} | Alfa: {scanResult.alfaCount}</span>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px 24px",marginBottom:20}}>
            {PP_LIST.map(pp=><div key={pp} style={{display:"flex",alignItems:"center",gap:8}}>
              <label style={{width:100,fontWeight:"bold",fontSize:14}}>{pp}:</label>
              <input type="number" min={0} value={ppReqs[pp]} onChange={e=>setPPReqs({...ppReqs,[pp]:parseInt(e.target.value)||0})} style={{padding:8,fontSize:15,width:80,borderRadius:6,border:"1px solid #ddd",textAlign:"center"}}/>
            </div>)}
          </div>
          <p style={{fontSize:13,color:"#666",marginBottom:16}}>Total: <strong>{Object.values(ppReqs).reduce((a,b)=>a+b,0)}</strong> / {scanResult.totalAvailable}</p>
          <div style={{display:"flex",gap:12}}>
            <button onClick={()=>setStep(1)} style={B("#ddd","#333")}>Back</button>
            <button onClick={handleDeploy} disabled={loading} style={{...B("#ff9900"),flex:1}}>{loading?"Deploying...":"Run Auto Deployment"}</button>
          </div>
        </div>}

        {step===3&&deployResult&&<div>
          <div style={{...C,background:"#f0f7ff"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
            <h2 style={{margin:0,color:"#232f3e"}}>Deployed — {site} | {date} | {shift}</h2>
            <button onClick={downloadFSK} style={B("#4caf50")}>Download FSK (.xlsx)</button>
          </div></div>
          <div style={{display:"flex",gap:16,marginBottom:16,flexWrap:"wrap"}}>
            {[{l:"Total Deployed",v:deployResult.totalDeployed,c:"#232f3e"},{l:"Roster",v:deployResult.rosterCount,c:"#2196f3"},{l:"Alfa OT",v:deployResult.alfaCount,c:"#ff9900"}].map(s=>
              <div key={s.l} style={{background:"#fff",border:"1px solid #ddd",borderRadius:8,padding:"12px 24px",textAlign:"center",flex:1,minWidth:120}}>
                <div style={{fontSize:28,fontWeight:"bold",color:s.c}}>{s.v}</div><div style={{fontSize:12,color:"#888"}}>{s.l}</div></div>)}
          </div>
          {deployResult.skippedNotInALC&&deployResult.skippedNotInALC.length>0&&<div style={{...C,background:"#fff3cd",border:"1px solid #ffc107"}}>
            <strong style={{color:"#856404"}}>⚠️ Skipped ({deployResult.skippedNotInALC.length}) — Not in ALC Mapping:</strong>
            <p style={{color:"#856404",fontSize:13,margin:"8px 0 0 0"}}>{deployResult.skippedNotInALC.join(", ")}</p>
          </div>}
          <div style={C}><h3 style={{marginTop:0}}>Summary</h3><div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr style={{background:"#232f3e",color:"#fff"}}>{["PP","Req","Deployed","Primary","Secondary","Overflow","Shortage","LC5","LC4","LC3","LC2","LC1","LC0"].map(h=><th key={h} style={{padding:8,border:"1px solid #555"}}>{h}</th>)}</tr></thead>
              <tbody>{Object.entries(deployResult.summary).map(([pp,s]:any)=><tr key={pp}>
                <td style={{padding:6,border:"1px solid #eee",fontWeight:"bold"}}>{pp}</td>
                {[s.required,s.deployed,s.primaryCount,s.secondaryCount,s.overflowCount].map((v:number,i:number)=><td key={i} style={{padding:6,border:"1px solid #eee",textAlign:"center",fontWeight:i===1?"bold":"normal"}}>{v}</td>)}
                <td style={{padding:6,border:"1px solid #eee",textAlign:"center",color:s.shortage>0?"red":"green",fontWeight:"bold"}}>{s.shortage}</td>
                {["LC5","LC4","LC3","LC2","LC1","LC0"].map(l=><td key={l} style={{padding:6,border:"1px solid #eee",textAlign:"center"}}>{s.lc?.[l]||0}</td>)}
              </tr>)}</tbody>
            </table></div>
          </div>
          <div style={{...C,background:"#e3f2fd",border:"2px solid #2196f3",textAlign:"center"}}>
            <h3 style={{marginTop:0,color:"#1565c0"}}>📋 Upload FSK to FastStart Portal</h3>
            <p style={{color:"#666",marginBottom:12}}>Download the FSK file above and upload it here:</p>
            <a href={`${FASTSTART_BASE}/${site}/audit`} target="_blank" rel="noopener noreferrer" style={{display:"inline-block",padding:"12px 32px",background:"#1565c0",color:"#fff",borderRadius:8,fontWeight:"bold",fontSize:16,textDecoration:"none"}}>Open FastStart Portal → {site}</a>
            <p style={{color:"#999",fontSize:12,marginTop:8,marginBottom:0}}>admin.eu-west-1.faststart.ats.amazon.dev/#/DEFAULT/{site}/audit</p>
          </div>
          <div style={{...C,textAlign:"center",border:"2px dashed #ff9900",background:"#fffbf0"}}>
            <h3 style={{marginTop:0,color:"#ff9900"}}>Waiting for Punch-In...</h3>
            <button onClick={()=>setStep(4)} style={{...B("#ff9900"),padding:"16px 48px",fontSize:18}}>Punch-In Done → Upload FCLM</button>
          </div>
        </div>}

        {step===4&&<div style={{...C,maxWidth:550,margin:"0 auto",textAlign:"center"}}>
          <h2 style={{marginTop:0,color:"#232f3e"}}>Step 4: Upload FCLM Punch</h2>
          <div style={{border:"2px dashed #ddd",borderRadius:12,padding:32,marginBottom:20,background:fclmFile?"#e8f5e9":"#fafafa"}}>
            <input type="file" accept=".xlsx,.xls" onChange={e=>setFclmFile(e.target.files?.[0]||null)} style={{fontSize:15}}/>
            {fclmFile&&<p style={{color:"#4caf50",marginTop:8,marginBottom:0}}>{fclmFile.name}</p>}
          </div>
          <div style={{display:"flex",gap:12,justifyContent:"center"}}>
            <button onClick={()=>setStep(3)} style={B("#ddd","#333")}>Back</button>
            <button onClick={handleActual} disabled={loading||!fclmFile} style={B("#ff9900")}>{loading?"Processing...":"Generate Shift Health Report"}</button>
          </div>
        </div>}


        {step===5&&actualResult&&(()=>{
          const t=actualResult.totals;const lc=t.lc||{};
          let tR=0,tD=0,tF=0,tP=0,tPr=0,tE=0,tA=0,tS=0,tO=0;
          const tLC:Record<string,number>={LC5:0,LC4:0,LC3:0,LC2:0,LC1:0,LC0:0};
          for(const[pp,s]of Object.entries(actualResult.summary)as any){
            tR+=s.required;tD+=s.deployed;tF+=s.ftDeployed;tP+=s.ptDeployed;tPr+=s.present;tE+=s.extra;tA+=s.absent;tS+=s.shortage;tO+=s.overflow;
            for(const l of Object.keys(tLC))tLC[l]+=(s.lc?.[l]||0);
          }
          return(<div>
            <div style={{...C,background:"linear-gradient(135deg,#232f3e 0%,#37475a 100%)",color:"#fff",border:"none"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
                <div>
                  <h2 style={{margin:0}}>Shift Health Report — {site}</h2>
                  <p style={{margin:"4px 0 0 0",opacity:0.8,fontSize:14}}>{date} | {shift} Shift</p>
                </div>
                <div style={{display:"flex",gap:8}}>
                  <button onClick={downloadFullReport} style={{...B("#ff9900"),padding:"14px 32px",fontSize:16}}>Download Full Report</button>
                  <button onClick={handleNewDeploy} style={B("#4caf50")}>New Deployment</button>
                </div>
              </div>
            </div>
            <div style={{...C,background:"#f8f9fa",border:"2px solid #232f3e"}}>
              <div style={{textAlign:"center",marginBottom:16}}>
                <div style={{fontSize:14,color:"#888",textTransform:"uppercase",letterSpacing:1}}>Total Headcount</div>
                <div style={{fontSize:56,fontWeight:"bold",color:"#232f3e"}}>{t.totalHC}</div>
              </div>
              <div style={{display:"flex",gap:12,justifyContent:"center",flexWrap:"wrap"}}>
                {[{l:"Planned",v:t.planned,c:"#232f3e"},{l:"Present",v:t.present,c:"#4caf50"},{l:"Absent",v:t.absent,c:"#f44336"},{l:"Extra",v:t.extra,c:"#2196f3"},{l:"FT",v:t.totalFT,c:"#ff9900"},{l:"PT",v:t.totalPT,c:"#9c27b0"},{l:"Attendance",v:t.attendance,c:"#009688"}].map(s=>
                  <div key={s.l} style={{background:"#fff",border:"1px solid #ddd",borderRadius:8,padding:"10px 20px",textAlign:"center",flex:1,minWidth:90}}>
                    <div style={{fontSize:24,fontWeight:"bold",color:s.c}}>{s.v}</div>
                    <div style={{fontSize:11,color:"#888"}}>{s.l}</div>
                  </div>
                )}
              </div>
            </div>
            <div style={C}>
              <h3 style={{marginTop:0}}>LC-wise Headcount</h3>
              <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
                {["LC5","LC4","LC3","LC2","LC1","LC0"].map(l=>
                  <div key={l} style={{background:"#f0f7ff",border:"1px solid #bbdefb",borderRadius:8,padding:"12px 24px",textAlign:"center",flex:1,minWidth:80}}>
                    <div style={{fontSize:26,fontWeight:"bold",color:"#1565c0"}}>{lc[l]||0}</div>
                    <div style={{fontSize:13,color:"#666",fontWeight:"bold"}}>{l}</div>
                  </div>
                )}
              </div>
            </div>
            {actualResult.skippedNotInALC&&actualResult.skippedNotInALC.length>0&&<div style={{...C,background:"#fff3cd",border:"1px solid #ffc107"}}>
              <strong style={{color:"#856404"}}>Skipped ({actualResult.skippedNotInALC.length}) — Not in ALC Mapping (not deployed):</strong>
              <p style={{color:"#856404",fontSize:13,margin:"8px 0 0 0"}}>{actualResult.skippedNotInALC.join(", ")}</p>
            </div>}
            <div style={{display:"flex",gap:6,marginBottom:16}}>
              {(["summary","badges","absent","audit"]as const).map(tab=>
                <button key={tab} onClick={()=>setActualView(tab)}
                  style={{padding:"8px 20px",background:actualView===tab?"#232f3e":"#eee",color:actualView===tab?"#fff":"#333",border:"none",borderRadius:8,cursor:"pointer",fontWeight:actualView===tab?"bold":"normal",fontSize:13}}>
                  {tab.charAt(0).toUpperCase()+tab.slice(1)}
                </button>
              )}
            </div>
            {actualView==="summary"&&<div style={C}>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead>
                    <tr style={{background:"#232f3e",color:"#fff"}}>
                      {["PP","Req","Deployed","FT","PT","Present","Extra","Absent","Shortage","Overflow","LC5","LC4","LC3","LC2","LC1","LC0"].map(h=>
                        <th key={h} style={{padding:6,border:"1px solid #555",whiteSpace:"nowrap"}}>{h}</th>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {Object.entries(actualResult.summary).map(([pp,s]:any)=>
                      (s.deployed>0||s.required>0||s.absent>0)?
                      <tr key={pp}>
                        <td style={{padding:5,border:"1px solid #eee",fontWeight:"bold"}}>{pp}</td>
                        <td style={{padding:5,border:"1px solid #eee",textAlign:"center"}}>{s.required}</td>
                        <td style={{padding:5,border:"1px solid #eee",textAlign:"center",fontWeight:"bold"}}>{s.deployed}</td>
                        <td style={{padding:5,border:"1px solid #eee",textAlign:"center"}}>{s.ftDeployed}</td>
                        <td style={{padding:5,border:"1px solid #eee",textAlign:"center"}}>{s.ptDeployed}</td>
                        <td style={{padding:5,border:"1px solid #eee",textAlign:"center",color:"#4caf50"}}>{s.present}</td>
                        <td style={{padding:5,border:"1px solid #eee",textAlign:"center",color:"#2196f3"}}>{s.extra}</td>
                        <td style={{padding:5,border:"1px solid #eee",textAlign:"center",color:"#f44336"}}>{s.absent}</td>
                        <td style={{padding:5,border:"1px solid #eee",textAlign:"center",color:s.shortage>0?"red":"green",fontWeight:"bold"}}>{s.shortage}</td>
                        <td style={{padding:5,border:"1px solid #eee",textAlign:"center"}}>{s.overflow}</td>
                        {["LC5","LC4","LC3","LC2","LC1","LC0"].map(l=>
                          <td key={l} style={{padding:5,border:"1px solid #eee",textAlign:"center"}}>{s.lc?.[l]||0}</td>
                        )}
                      </tr>:null
                    )}
                    <tr style={{background:"#fff3e0",fontWeight:"bold"}}>
                      <td style={{padding:6,border:"1px solid #ddd"}}>TOTAL</td>
                      <td style={{padding:6,border:"1px solid #ddd",textAlign:"center"}}>{tR}</td>
                      <td style={{padding:6,border:"1px solid #ddd",textAlign:"center"}}>{tD}</td>
                      <td style={{padding:6,border:"1px solid #ddd",textAlign:"center"}}>{tF}</td>
                      <td style={{padding:6,border:"1px solid #ddd",textAlign:"center"}}>{tP}</td>
                      <td style={{padding:6,border:"1px solid #ddd",textAlign:"center",color:"#4caf50"}}>{tPr}</td>
                      <td style={{padding:6,border:"1px solid #ddd",textAlign:"center",color:"#2196f3"}}>{tE}</td>
                      <td style={{padding:6,border:"1px solid #ddd",textAlign:"center",color:"#f44336"}}>{tA}</td>
                      <td style={{padding:6,border:"1px solid #ddd",textAlign:"center",color:tS>0?"red":"green"}}>{tS}</td>
                      <td style={{padding:6,border:"1px solid #ddd",textAlign:"center"}}>{tO}</td>
                      {["LC5","LC4","LC3","LC2","LC1","LC0"].map(l=>
                        <td key={l} style={{padding:6,border:"1px solid #ddd",textAlign:"center"}}>{tLC[l]}</td>
                      )}
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>}
            {actualView==="badges"&&<div style={{...C,display:"flex",gap:12,overflowX:"auto"}}>
              {Object.entries(actualResult.badgesByPP).map(([pp,badges]:any)=>
                badges.length>0?
                <div key={pp} style={{minWidth:120,flexShrink:0}}>
                  <div style={{background:"#232f3e",color:"#fff",padding:6,fontWeight:"bold",textAlign:"center",borderRadius:"8px 8px 0 0",fontSize:13}}>
                    {pp} ({badges.length})
                  </div>
                  <div style={{border:"1px solid #ddd",borderTop:"none",borderRadius:"0 0 8px 8px",maxHeight:300,overflowY:"auto"}}>
                    {badges.map((b:string,i:number)=>
                      <div key={i} style={{padding:"4px 8px",borderBottom:"1px solid #f0f0f0",textAlign:"center",fontSize:12}}>{b}</div>
                    )}
                  </div>
                </div>:null
              )}
            </div>}
            {actualView==="absent"&&<div style={C}>
              <h3 style={{color:"#f44336",marginTop:0}}>Absent ({actualResult.absentList.length})</h3>
              {actualResult.absentList.length>0?
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead>
                    <tr style={{background:"#ffebee"}}>
                      {["Badge ID","Emp ID","Name","Alias","Planned PP","LC"].map(h=>
                        <th key={h} style={{padding:8,border:"1px solid #ddd",textAlign:"left"}}>{h}</th>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {actualResult.absentList.map((a:any,i:number)=>
                      <tr key={i}>
                        <td style={{padding:6,border:"1px solid #eee",color:"red",fontWeight:"bold"}}>{a.badgeID}</td>
                        <td style={{padding:6,border:"1px solid #eee"}}>{a.empID}</td>
                        <td style={{padding:6,border:"1px solid #eee"}}>{a.name}</td>
                        <td style={{padding:6,border:"1px solid #eee"}}>{a.alias}</td>
                        <td style={{padding:6,border:"1px solid #eee"}}>{a.plannedPP}</td>
                        <td style={{padding:6,border:"1px solid #eee"}}>{a.lc}</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              :<p style={{color:"#4caf50"}}>No absent employees!</p>}
            </div>}
            {actualView==="audit"&&<div style={{...C,overflowX:"auto"}}>
              <table style={{borderCollapse:"collapse",fontSize:11,whiteSpace:"nowrap"}}>
                <thead>
                  <tr style={{background:"#fff8e1"}}>
                    {["Source","Emp ID","Alias","Badge","Status","Reason","Primary PP","Assigned PP","FT/PT","LC"].map(h=>
                      <th key={h} style={{padding:6,border:"1px solid #ddd"}}>{h}</th>
                    )}
                  </tr>
                </thead>
                <tbody>
                  {actualResult.auditLog.map((e:any,i:number)=>
                    <tr key={i} style={{background:e.status==="SKIPPED"?"#fff3cd":undefined}}>
                      <td style={{padding:4,border:"1px solid #eee"}}>{e.inFSK==="YES"?"Planned":"Extra"}</td>
                      <td style={{padding:4,border:"1px solid #eee"}}>{e.empID}</td>
                      <td style={{padding:4,border:"1px solid #eee"}}>{e.alias}</td>
                      <td style={{padding:4,border:"1px solid #eee"}}>{e.badgeID}</td>
                      <td style={{padding:4,border:"1px solid #eee",color:e.status==="Present"?"green":e.status==="ABSENT"?"red":e.status==="Extra"?"blue":e.status==="SKIPPED"?"#856404":"orange",fontWeight:"bold"}}>{e.status}</td>
                      <td style={{padding:4,border:"1px solid #eee"}}>{e.reason}</td>
                      <td style={{padding:4,border:"1px solid #eee"}}>{e.primaryPP}</td>
                      <td style={{padding:4,border:"1px solid #eee"}}>{e.assignedPP}</td>
                      <td style={{padding:4,border:"1px solid #eee"}}>{e.empType}</td>
                      <td style={{padding:4,border:"1px solid #eee"}}>{e.lc}</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>}
          </div>);
        })()}
      </>}
    </div>
  );
}

