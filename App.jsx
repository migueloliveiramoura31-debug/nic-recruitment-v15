import React, { useState, useEffect, useCallback, useMemo } from "react";
import Papa from "papaparse";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

// ── Export helpers ────────────────────────────────────────────────────────────
function exportExcel({candidates,allScores,interviewData,chosenCandidates,members,aiScores,promoted}){
  const wb=XLSX.utils.book_new();

  // Sheet 1: All candidates with scores
  const appRows=candidates.map((c,i)=>{
    const avg=allScores.filter(s=>s.candidate_id===c.id);
    const score=avg.length?((avg.reduce((a,r)=>a+parseFloat(r.score),0)/avg.length/4)*100).toFixed(1):null;
    const ai=aiScores[c.id];
    return{
      Rank:i+1,
      "Student No":c.student_number,
      Name:c.full_name!==c.student_number?c.full_name:"",
      Email:c.email,
      "Avg Score (%)":score?parseFloat(score):null,
      "AI Detection (%)":ai?.overall_pct??null,
      "Round":promoted[c.id]||"Application only",
      "Chosen":chosenCandidates.some(x=>x.id===c.id)?"Yes":"No",
    };
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(appRows),"Applications");

  // Sheet 2: Interview feedback
  const ivRows=interviewData.map(iv=>{
    const member=members.find(m=>m.id===iv.interviewer_id);
    const candidate=candidates.find(c=>c.id===iv.candidate_id);
    return{
      "Student No":candidate?.student_number||iv.candidate_id,
      "Name":candidate?.full_name!==candidate?.student_number?candidate?.full_name:"",
      "Round":iv.round==="president"?"Final Round":"Member Interview",
      "Interviewer":member?.name||iv.interviewer_id,
      "Personal":iv.personal_score,
      "Technical":iv.technical_score,
      "Brainstormer":iv.brainstormer_score,
      "Verdict":iv.verdict,
      "Interview Date":iv.interview_date||"",
      "Feedback":iv.feedback||"",
    };
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(ivRows),"Interview Feedback");

  // Sheet 3: Chosen candidates summary
  const chosenRows=chosenCandidates.map(c=>{
    const presIv=interviewData.filter(f=>f.candidate_id===c.id&&f.round==="president");
    const memIv=interviewData.filter(f=>f.candidate_id===c.id&&f.round==="interview");
    const avg=allScores.filter(s=>s.candidate_id===c.id);
    const appScore=avg.length?((avg.reduce((a,r)=>a+parseFloat(r.score),0)/avg.length/4)*100).toFixed(1):null;
    return{
      "Student No":c.student_number,
      "Name":c.full_name!==c.student_number?c.full_name:"",
      "Email":c.email,
      "App Score (%)":appScore?parseFloat(appScore):null,
      "Member Verdict 1":memIv[0]?.verdict||"",
      "Member Verdict 2":memIv[1]?.verdict||"",
      "President Verdict":presIv[0]?.verdict||"",
      "Final Personal":presIv[0]?.personal_score||"",
      "Final Technical":presIv[0]?.technical_score||"",
      "Final Brainstormer":presIv[0]?.brainstormer_score||"",
    };
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(chosenRows),"Chosen Candidates");

  XLSX.writeFile(wb,`NIC-UD_Recruitment_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function exportPDF({candidates,allScores,interviewData,chosenCandidates,members,aiScores,promoted,ranked}){
  // Build HTML and open as print dialog — no external PDF lib needed
  const fmtS=v=>v!=null?`${parseFloat(v).toFixed(1)}%`:"—";
  const rows=ranked.map((c,i)=>{
    const ai=aiScores[c.id];
    const status=promoted[c.id]||"—";
    const chosen=chosenCandidates.some(x=>x.id===c.id);
    const presIv=interviewData.filter(f=>f.candidate_id===c.id&&f.round==="president");
    const memIv=interviewData.filter(f=>f.candidate_id===c.id&&f.round==="interview");
    return `<tr style="background:${chosen?"#fef9c3":i%2===0?"#fff":"#f8fafc"}">
      <td style="font-weight:700;color:#0f2952">#${i+1}</td>
      <td><strong>${c.full_name!==c.student_number?c.full_name:""}</strong><br><span style="color:#7a90b0;font-size:11px">${c.student_number} · ${c.email||""}</span></td>
      <td style="text-align:center;font-weight:700;color:${parseFloat(c.avg)>=75?"#16a34a":parseFloat(c.avg)>=50?"#d97706":"#dc2626"}">${fmtS(c.avg)}</td>
      <td style="text-align:center">${ai?.overall_pct!=null?`${ai.overall_pct}%`:"—"}</td>
      <td style="text-align:center">${status}</td>
      <td style="text-align:center">${memIv.map(f=>`<span style="color:${f.verdict==="pass"?"#16a34a":f.verdict==="borderline"?"#d97706":"#dc2626"}">${f.verdict||"—"}</span>`).join(" / ")||"—"}</td>
      <td style="text-align:center">${presIv.map(f=>`<span style="color:${f.verdict==="pass"?"#16a34a":f.verdict==="borderline"?"#d97706":"#dc2626"}">${f.verdict||"—"}</span>`).join(" / ")||"—"}</td>
      <td style="text-align:center">${chosen?"⭐ YES":"—"}</td>
    </tr>`;
  }).join("");

  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>NIC-UD Recruitment Report</title>
  <style>
    body{font-family:Arial,sans-serif;font-size:12px;margin:0;padding:20px;color:#0d1b2a}
    h1{color:#0f2952;font-size:22px;margin:0 0 4px}
    p{color:#7a90b0;font-size:12px;margin:0 0 20px}
    table{width:100%;border-collapse:collapse;margin-bottom:24px}
    th{background:#0f2952;color:#fff;padding:8px 10px;text-align:left;font-size:11px;letter-spacing:1px}
    td{padding:8px 10px;border-bottom:1px solid #dde3ef;vertical-align:top;font-size:11px}
    .section{font-size:13px;font-weight:700;color:#0f2952;margin:20px 0 8px;letter-spacing:1px;text-transform:uppercase;border-bottom:2px solid #0f2952;padding-bottom:4px}
    @media print{body{padding:10px}.no-print{display:none}}
  </style></head><body>
  <div style="display:flex;align-items:center;gap:16px;margin-bottom:20px">
    <div><h1>NIC-UD Recruitment Report</h1><p>Generated ${new Date().toLocaleDateString("en-GB",{day:"numeric",month:"long",year:"numeric"})} · ${candidates.length} applicants · ${chosenCandidates.length} chosen</p></div>
  </div>

  <div class="section">Recruitment Funnel</div>
  <table><tr>
    ${[["Applications",candidates.length],["Interview Round",interviewData.filter(f=>f.round==="interview").map(f=>f.candidate_id).filter((v,i,a)=>a.indexOf(v)===i).length],["Final Round",interviewData.filter(f=>f.round==="president").map(f=>f.candidate_id).filter((v,i,a)=>a.indexOf(v)===i).length],["Chosen",chosenCandidates.length]].map(([l,n])=>`<td style="text-align:center"><div style="font-size:28px;font-weight:900;color:#0f2952">${n}</div><div style="color:#7a90b0">${l}</div></td>`).join("")}
  </tr></table>

  <div class="section">All Candidates</div>
  <table><thead><tr><th>#</th><th>Candidate</th><th>App Score</th><th>AI %</th><th>Stage</th><th>Member Verdict</th><th>Pres. Verdict</th><th>Chosen</th></tr></thead><tbody>${rows}</tbody></table>

  <div class="section">Chosen Candidates (${chosenCandidates.length})</div>
  <table><thead><tr><th>Candidate</th><th>App Score</th><th>Member Verdicts</th><th>Pres. Verdict</th><th>P</th><th>T</th><th>B</th></tr></thead><tbody>
  ${chosenCandidates.map(c=>{
    const presIv=interviewData.filter(f=>f.candidate_id===c.id&&f.round==="president");
    const memIv=interviewData.filter(f=>f.candidate_id===c.id&&f.round==="interview");
    const avg=allScores.filter(s=>s.candidate_id===c.id);
    const appScore=avg.length?((avg.reduce((a,r)=>a+parseFloat(r.score),0)/avg.length/4)*100).toFixed(1):null;
    return`<tr style="background:#fef9c3"><td><strong>${c.full_name!==c.student_number?c.full_name:c.student_number}</strong><br><span style="color:#7a90b0">${c.email||""}</span></td>
    <td style="text-align:center;font-weight:700">${fmtS(appScore)}</td>
    <td style="text-align:center">${memIv.map(f=>f.verdict||"—").join(" / ")||"—"}</td>
    <td style="text-align:center">${presIv.map(f=>f.verdict||"—").join(" / ")||"—"}</td>
    <td style="text-align:center">${presIv[0]?.personal_score??"-"}</td>
    <td style="text-align:center">${presIv[0]?.technical_score??"-"}</td>
    <td style="text-align:center">${presIv[0]?.brainstormer_score??"-"}</td></tr>`;
  }).join("")}
  </tbody></table>
  <script>window.onload=()=>window.print();</script>
  </body></html>`;

  const w=window.open("","_blank");
  if(!w){alert("Please allow popups to export PDF.");return;}
  w.document.write(html);
  w.document.close();
}

const SUPABASE_URL = "https://cfijvlomsugjwdikphpe.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImNmaWp2bG9tc3VnandkaWtwaHBlIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk0NzQ5ODAsImV4cCI6MjA5NTA1MDk4MH0.w3z2O8G-A_SWKIXWvG2Ul9dBQZz7VpT-215U0BAK8QI";
const sb = createClient(SUPABASE_URL, SUPABASE_KEY);

// ── CSV column mapping (supports both old Google Forms and new Microsoft Forms) ──
// Microsoft Forms exports use the question text as column headers.
// We try multiple variants to match flexibly.
const COL_VARIANTS = {
  name:      ["Name","Full name","Full Name","Nome","Nome completo","What is your full name?"],
  student:   ["Student ID","Student number","What is your student number?","Número de aluno"],
  email:     ["Student email address","What is your institutional email?","Email","E-mail"],
  phone:     ["Phone number","What is your phone number?","Telefone","Phone"],
  cv:        ["Upload CV","Please upload your CV.","CV","Curriculum"],
  b1:        ["Tell us about yourself. What motivates you? How are you different from other applicants? Where do you see yourself in five years? (Max 750 characters)"],
  b2:        ["What makes you want to join NIC-UD? Why do you think you could be a valuable member for the club? How are you different from other applicants? (Max 750 characters)",
              "What makes you want to join NIC-UD? Why do you think you could be a valuable member for the club? (Max 750 characters)"],
  b3:        ["What's one thing you pursued during your Bachelor's that wasn't required, and what motivated you to do it? (Max 750 characters)",
              "What\u2019s one thing you pursued during your Bachelor\u2019s that wasn\u2019t required, and what motivated you to do it? (Max 750 characters)"],
  t1:        ["Choose a stock from the NIC-UD fund that you would follow more closely due to its potential long-term risks. Explain your choice and foreseen risks. (Max 1000 characters)",
              "Choose a stock from the NIC-UD fund and explain to us why it might be a bad investment. (Max 500 characters)",
              "If you had \u20AC100,000 to invest, how would you allocate this capital across different asset classes or markets to build a solid, future-proof portfolio? Please explain your choices by considering current macroeconomic trends and some key fundamentals. (Max 1000 characters)"],
  comments:  ["Additional Comments/Questions","Additional Comments","Comments"],
  timestamp: ["Coluna 1","Start time","Completion time","Timestamp","ID"],
};

// Flexibly find a column value from a CSV row
function getCol(row, key) {
  const variants = COL_VARIANTS[key] || [];
  for (const v of variants) {
    if (row[v] !== undefined && row[v] !== null) return row[v];
  }
  // Fuzzy: try partial match on column headers
  const headers = Object.keys(row);
  for (const v of variants) {
    const found = headers.find(h => h.toLowerCase().includes(v.toLowerCase().slice(0, 30)));
    if (found && row[found]) return row[found];
  }
  return "";
}
function extractName(row) { return getCol(row, "name") || null; }
function displayName(c) { return c?.full_name && c.full_name !== c.student_number ? c.full_name : `#${c?.student_number}`; }

// ── Constants ─────────────────────────────────────────────────────────────────
const QUESTIONS = [
  { id:"b1", label:"About Yourself",    sublabel:"Motivation & Differentiation",
    fullText:"Tell us about yourself. What motivates you? How are you different from other applicants? Where do you see yourself in five years? (Max 750 characters)" },
  { id:"b2", label:"Why NIC-UD?",       sublabel:"Fit & Value Add",
    fullText:"What makes you want to join NIC-UD? Why do you think you could be a valuable member for the club? (Max 750 characters)" },
  { id:"b3", label:"Beyond Required",   sublabel:"Initiative & Self-Motivation",
    fullText:"What\u2019s one thing you pursued during your Bachelor\u2019s that wasn\u2019t required, and what motivated you to do it? (Max 750 characters)" },
  { id:"t1", label:"Technical",         sublabel:"Stock Risk Analysis",
    fullText:"Choose a stock from the NIC-UD fund that you would follow more closely due to its potential long-term risks. Explain your choice and foreseen risks. (Max 1000 characters)" },
];
const VERDICTS = [
  { id:"pass",       label:"Pass",       color:"#16a34a", bg:"#dcfce7" },
  { id:"borderline", label:"Borderline", color:"#d97706", bg:"#fef3c7" },
  { id:"fail",       label:"Fail",       color:"#dc2626", bg:"#fee2e2" },
];
const C = {
  navy:"#0f2952", navyMid:"#1a3a6b", navyLt:"#2451a0", accent:"#2d6ae0",
  bg:"#f4f6fb", white:"#ffffff", border:"#dde3ef",
  text:"#0d1b2a", textMid:"#3d5278", textLt:"#7a90b0",
  green:"#16a34a", amber:"#d97706", red:"#dc2626",
};

// ── AI heuristic detection ────────────────────────────────────────────────────
function scoreAI(text) {
  if (!text||text.trim().length<20) return 0;
  const t=text.toLowerCase(); let s=0;
  ["in today's","in the current","it is worth noting","it is important to","furthermore","moreover","additionally","in terms of","leveraging","utilize","multifaceted","holistic","synergy","robust","comprehensive","nuanced","delve","foster","pivotal","navigate","landscape","needless to say","going forward","in light of","it should be noted","one must consider"].forEach(p=>{if(t.includes(p))s+=6;});
  const sents=text.split(/[.!?]+/).filter(x=>x.trim().length>10);
  const avg=sents.reduce((a,x)=>a+x.length,0)/(sents.length||1);
  if(avg>180)s+=20; else if(avg>130)s+=10; else if(avg>90)s+=5;
  if((text.match(/\d+%/g)||[]).length>=4)s+=12;
  if(text.length>700)s+=8;
  ["i've","i'm","i'd","i'll","my father","my brother","my family","growing up","honestly","to be honest","i wasn't","i didn't","i don't","i can't"].forEach(p=>{if(t.includes(p))s-=8;});
  if(text.length<150)s-=15;
  return Math.max(0,Math.min(100,s));
}
function detectAI(candidate) {
  const b1=scoreAI(candidate.b1),b2=scoreAI(candidate.b2),b3=scoreAI(candidate.b3),t1=scoreAI(candidate.t1);
  const overall=Math.round((b1+b2+b3+t1)/4);
  const flags=overall>=70?"Multiple AI-associated patterns detected.":overall>=40?"Some formal phrasing — review manually.":"No strong AI patterns detected.";
  return {b1,b2,b3,t1,overall,flags};
}

// ── Pure helpers ──────────────────────────────────────────────────────────────
const aiColor    = p => p==null?C.textLt:p>=70?C.red:p>=40?C.amber:C.green;
const scoreColor = s => s==null?C.textLt:s>=75?C.green:s>=50?C.amber:C.red;
const initials   = n => (n||"?").split(" ").map(w=>w[0]).join("").slice(0,2).toUpperCase();
const verdictColor = v => VERDICTS.find(x=>x.id===v)?.color||C.textLt;
const verdictBg    = v => VERDICTS.find(x=>x.id===v)?.bg||"#f1f5f9";
const fmtScore     = v => v!=null ? parseFloat(v).toFixed(parseFloat(v)%1===0?0:1) : "—";
const fmtDate      = d => d ? new Date(d).toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"}) : null;

function avgScore(cid, scores) {
  const rows=scores.filter(s=>s.candidate_id===cid);
  if(!rows.length)return null;
  return((rows.reduce((a,r)=>a+parseFloat(r.score),0)/rows.length/4)*100).toFixed(1);
}
function memberProgress(mid,candidates,scores){
  const my=scores.filter(s=>s.member_id===mid);
  const done=(candidates||[]).filter(c=>QUESTIONS.every(q=>my.some(s=>s.candidate_id===c.id&&s.question_id===q.id))).length;
  return{done,total:(candidates||[]).length};
}

// ── UI atoms ──────────────────────────────────────────────────────────────────
function Spinner({size=20}){
  return <div style={{width:size,height:size,border:`2px solid ${C.border}`,borderTop:`2px solid ${C.navy}`,borderRadius:"50%",animation:"spin 0.7s linear infinite",flexShrink:0}}/>;
}
function AiBadge({pct}){
  if(pct==null)return <span style={{color:C.textLt,fontSize:12}}>—</span>;
  const col=aiColor(pct);
  return <span style={{background:col+"15",color:col,border:`1px solid ${col}44`,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>{pct}% AI</span>;
}
function VerdictBadge({verdict}){
  if(!verdict)return <span style={{color:C.textLt,fontSize:12}}>—</span>;
  const col=verdictColor(verdict),bg=verdictBg(verdict);
  return <span style={{background:bg,color:col,border:`1px solid ${col}44`,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>{verdict.charAt(0).toUpperCase()+verdict.slice(1)}</span>;
}
function Toast({toast}){
  if(!toast)return null;
  const ok=toast.type!=="err";
  return <div style={{position:"fixed",top:18,right:18,zIndex:9999,background:ok?"#f0fdf4":"#fef2f2",border:`1px solid ${ok?"#86efac":"#fca5a5"}`,color:ok?C.green:C.red,borderRadius:10,padding:"11px 20px",fontSize:14,fontWeight:600,fontFamily:"system-ui,sans-serif",animation:"fadeUp 0.2s ease",boxShadow:"0 4px 16px rgba(0,0,0,0.10)"}}>{toast.msg}</div>;
}
function ScoreInput({value,onChange}){
  return(
    <div style={{display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>
      {[0,1,2,3,4].map(n=>(
        <button key={n} onClick={()=>onChange(n)}
          style={{width:34,height:34,borderRadius:7,border:`1.5px solid ${parseFloat(value)===n?C.navy:C.border}`,background:parseFloat(value)===n?C.navy:"#fff",color:parseFloat(value)===n?"#fff":C.textMid,fontWeight:800,fontSize:14,cursor:"pointer",flexShrink:0}}>
          {n}
        </button>
      ))}
      <div style={{display:"flex",alignItems:"center",gap:4,background:C.bg,border:`1.5px solid ${C.border}`,borderRadius:7,padding:"2px 8px"}}>
        <span style={{fontSize:10,color:C.textLt,fontWeight:700}}>CUSTOM</span>
        <input type="number" min="0" max="4" step="0.5" value={value??""} placeholder="2.5"
          onChange={e=>{const v=parseFloat(e.target.value);if(!isNaN(v)&&v>=0&&v<=4)onChange(v);}}
          style={{width:46,border:"none",background:"transparent",fontSize:14,fontWeight:700,color:C.navy,textAlign:"center",outline:"none"}}/>
      </div>
      {value!=null&&<span style={{fontSize:11,color:C.navy,fontWeight:600,marginLeft:2}}>{parseFloat(value)===0?"No merit":parseFloat(value)<=1?"Below avg":parseFloat(value)<=2?"Average":parseFloat(value)<=3?"Good":"Excellent"}</span>}
    </div>
  );
}

// ── Answers Modal (standalone, not inside render) ─────────────────────────────
function AnswersModal({candidate,aiScores,onClose}){
  const c=candidate;
  if(!c)return null;
  const ai=aiScores[c.id];
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={onClose}>
      <div style={{background:"#fff",borderRadius:16,width:"100%",maxWidth:720,maxHeight:"90vh",overflow:"hidden",display:"flex",flexDirection:"column",boxShadow:"0 8px 40px rgba(0,0,0,0.25)"}} onClick={e=>e.stopPropagation()}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"16px 22px",background:C.navy,flexShrink:0,gap:10,flexWrap:"wrap"}}>
          <div>
            <div style={{fontWeight:700,fontSize:16,color:"#fff"}}>{displayName(c)}</div>
            <div style={{fontSize:12,color:"rgba(255,255,255,0.5)",marginTop:2}}>#{c.student_number}{c.email?` · ${c.email}`:""}</div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            {c.cv_link&&<a href={c.cv_link} target="_blank" rel="noopener noreferrer" style={{background:"rgba(255,255,255,0.15)",color:"#fff",borderRadius:7,padding:"6px 12px",fontSize:12,fontWeight:700,textDecoration:"none"}}>📄 CV ↗</a>}
            <button onClick={onClose} style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"none",borderRadius:7,padding:"6px 12px",fontSize:14,fontWeight:700,cursor:"pointer"}}>✕</button>
          </div>
        </div>
        {ai?.overall_pct!=null&&(
          <div style={{display:"flex",gap:12,padding:"10px 22px",background:C.bg,borderBottom:`1px solid ${C.border}`,flexShrink:0,flexWrap:"wrap",alignItems:"center"}}>
            <span style={{fontSize:10,color:C.textLt,fontWeight:700,letterSpacing:1}}>AI DETECTION</span>
            {[["B1",ai.b1_pct],["B2",ai.b2_pct],["B3",ai.b3_pct],["T1",ai.t1_pct],["Overall",ai.overall_pct]].map(([lbl,val])=>(
              <span key={lbl} style={{fontSize:12,fontWeight:700,color:aiColor(val)}}>{lbl}: {val}%</span>
            ))}
            {ai.flags&&<span style={{fontSize:11,color:C.textLt}}>· {ai.flags}</span>}
          </div>
        )}
        <div style={{overflowY:"auto",padding:"18px 22px",display:"flex",flexDirection:"column",gap:14}}>
          {QUESTIONS.map(q=>(
            <div key={q.id} style={{background:C.bg,borderRadius:10,padding:16,borderLeft:`3px solid ${C.border}`}}>
              <div style={{fontSize:10,color:C.navy,letterSpacing:2,fontWeight:700,marginBottom:4}}>{q.label.toUpperCase()}</div>
              <div style={{fontSize:11,color:C.textLt,lineHeight:1.5,marginBottom:8,fontStyle:"italic"}}>"{q.fullText}"</div>
              <div style={{fontSize:14,color:C.text,lineHeight:1.85,whiteSpace:"pre-wrap",wordBreak:"break-word"}}>
                {c[q.id]?.trim()||<em style={{color:C.textLt}}>No answer provided</em>}
              </div>
            </div>
          ))}
          {/* Additional comments — not scored */}
          {c.comments?.trim()&&(
            <div style={{background:"#fffbeb",borderRadius:10,padding:16,borderLeft:"3px solid #fcd34d"}}>
              <div style={{fontSize:10,color:C.amber,letterSpacing:2,fontWeight:700,marginBottom:8}}>ADDITIONAL COMMENTS (not scored)</div>
              <div style={{fontSize:14,color:C.textMid,lineHeight:1.85,whiteSpace:"pre-wrap",wordBreak:"break-word"}}>{c.comments}</div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ── Assign Modal ──────────────────────────────────────────────────────────────
function AssignModal({assignModal,candidates,assignments,members,onAssign,onClose}){
  const candidate=candidates?.find(c=>c.id===assignModal?.id);
  const round=assignModal?.round||"interview";
  const current=useMemo(()=>assignments.filter(a=>a.candidate_id===assignModal?.id&&a.round===round).map(a=>a.interviewer_id),[assignments,assignModal,round]);
  const [sel,setSel]=useState(current);
  const eligible=round==="president"?members.filter(m=>m.role==="president"):members.filter(m=>m.role!=="president");
  if(!candidate)return null;
  const maxSel=2;
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={onClose}>
      <div style={{background:"#fff",borderRadius:16,padding:28,width:"100%",maxWidth:440,boxShadow:"0 8px 40px rgba(0,0,0,0.2)"}} onClick={e=>e.stopPropagation()}>
        <h3 style={{margin:"0 0 4px",color:C.navy,fontSize:17}}>Assign Interviewers</h3>
        <p style={{color:C.textMid,fontSize:13,marginBottom:18}}>{displayName(candidate)} · {round==="president"?"Final Round":"Member Round"}</p>
        <p style={{fontSize:11,color:C.textLt,marginBottom:10,fontWeight:700,letterSpacing:1}}>SELECT UP TO 2:</p>
        <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:20,maxHeight:260,overflowY:"auto"}}>
          {eligible.map(m=>{
            const on=sel.includes(m.id);
            const disabled=!on&&sel.length>=maxSel;
            return(
              <button key={m.id} onClick={()=>!disabled&&setSel(prev=>on?prev.filter(x=>x!==m.id):[...prev,m.id])}
                style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",background:on?C.navy:C.bg,border:`1.5px solid ${on?C.navy:C.border}`,borderRadius:9,cursor:disabled?"not-allowed":"pointer",color:on?"#fff":C.text,opacity:disabled?0.4:1}}>
                <span style={{width:28,height:28,borderRadius:"50%",background:on?"rgba(255,255,255,0.2)":C.border,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800,color:on?"#fff":C.navy,flexShrink:0}}>{initials(m.name)}</span>
                <span style={{fontSize:13,fontWeight:600}}>{m.name}</span>
                {on&&<span style={{marginLeft:"auto"}}>✓</span>}
              </button>
            );
          })}
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={onClose} style={{flex:1,padding:"10px",background:C.bg,border:`1px solid ${C.border}`,borderRadius:9,fontSize:14,fontWeight:600,cursor:"pointer",color:C.textMid}}>Cancel</button>
          <button onClick={()=>onAssign(assignModal.id,round,sel)} disabled={sel.length===0}
            style={{flex:1,padding:"10px",background:sel.length>0?C.navy:"#c5cedc",color:"#fff",border:"none",borderRadius:9,fontSize:14,fontWeight:700,cursor:sel.length>0?"pointer":"not-allowed"}}>
            Confirm ({sel.length} selected)
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Interview feedback form ────────────────────────────────────────────────────
function InterviewForm({candidate,interviewData,user,onSave,round,onViewAnswers}){
  const existing=interviewData.find(i=>i.interviewer_id===user.id&&i.candidate_id===candidate.id&&i.round===round);
  const [personal,      setPersonal]      = useState(existing?.personal_score??null);
  const [technical,     setTechnical]     = useState(existing?.technical_score??null);
  const [brainstormer,  setBrainstormer]  = useState(existing?.brainstormer_score??null);
  const [feedback,      setFeedback]      = useState(existing?.feedback??"");
  const [verdict,       setVerdict]       = useState(existing?.verdict??"");
  const [interviewDate, setInterviewDate] = useState(existing?.interview_date??"");
  const [saving,        setSaving]        = useState(false);

  const save=async()=>{
    setSaving(true);
    await onSave({candidate_id:candidate.id,interviewer_id:user.id,round,personal_score:personal,technical_score:technical,brainstormer_score:brainstormer,feedback,verdict,interview_date:interviewDate||null});
    setSaving(false);
  };

  return(
    <div style={{padding:20,background:C.white,border:`1px solid ${C.border}`,borderTop:"none",borderRadius:"0 0 12px 12px"}}>
      {/* Quick access */}
      <div style={{display:"flex",gap:8,marginBottom:18,padding:"10px 14px",background:C.bg,borderRadius:9,alignItems:"center",flexWrap:"wrap"}}>
        <span style={{fontSize:10,color:C.textLt,fontWeight:700,letterSpacing:1}}>QUICK ACCESS</span>
        {candidate.cv_link&&<a href={candidate.cv_link} target="_blank" rel="noopener noreferrer" style={{background:C.navy,color:"#fff",borderRadius:7,padding:"6px 12px",fontSize:12,fontWeight:700,textDecoration:"none"}}>📄 CV →</a>}
        <button onClick={()=>onViewAnswers(candidate)} style={{background:"#eff6ff",color:C.navy,border:`1px solid #bfdbfe`,borderRadius:7,padding:"6px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>📋 Answers</button>
        <div style={{display:"flex",alignItems:"center",gap:6,marginLeft:"auto",flexWrap:"wrap"}}>
          <span style={{fontSize:10,color:C.textLt,fontWeight:700,letterSpacing:1}}>📅 DATE</span>
          <input type="date" value={interviewDate} onChange={e=>setInterviewDate(e.target.value)}
            style={{border:`1px solid ${C.border}`,borderRadius:7,padding:"5px 9px",fontSize:13,color:C.text,background:"#fff",cursor:"pointer",outline:"none"}}/>
        </div>
      </div>
      {/* Scores */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12,marginBottom:18}}>
        {[["Personal",personal,setPersonal],["Technical",technical,setTechnical],["Brainstormer",brainstormer,setBrainstormer]].map(([lbl,val,set])=>(
          <div key={lbl} style={{background:C.bg,borderRadius:9,padding:12}}>
            <div style={{fontSize:10,color:C.navy,letterSpacing:1.5,fontWeight:700,marginBottom:8}}>{lbl.toUpperCase()}</div>
            <ScoreInput value={val} onChange={set}/>
          </div>
        ))}
      </div>
      {/* Verdict */}
      <div style={{marginBottom:14}}>
        <div style={{fontSize:10,color:C.textLt,letterSpacing:1.5,fontWeight:700,marginBottom:8}}>YOUR VERDICT</div>
        <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
          {VERDICTS.map(v=>(
            <button key={v.id} onClick={()=>setVerdict(v.id)}
              style={{flex:"1 1 90px",padding:"9px 8px",borderRadius:9,border:`2px solid ${verdict===v.id?v.color:C.border}`,background:verdict===v.id?v.bg:"#fff",color:verdict===v.id?v.color:C.textMid,fontWeight:700,fontSize:13,cursor:"pointer"}}>
              {v.id==="pass"?"✅":v.id==="borderline"?"🟡":"❌"} {v.label}
            </button>
          ))}
        </div>
      </div>
      {/* Feedback */}
      <div style={{marginBottom:14}}>
        <div style={{fontSize:10,color:C.textLt,letterSpacing:1.5,fontWeight:700,marginBottom:8}}>YOUR FEEDBACK</div>
        <textarea value={feedback} onChange={e=>setFeedback(e.target.value)} rows={4} placeholder="Write your interview notes here…"
          style={{width:"100%",border:`1px solid ${C.border}`,borderRadius:9,padding:"10px 13px",fontSize:14,color:C.text,resize:"vertical",fontFamily:"system-ui,sans-serif",outline:"none",background:C.bg}}/>
      </div>
      <button onClick={save} disabled={saving}
        style={{background:C.navy,color:"#fff",border:"none",borderRadius:9,padding:"10px 22px",fontSize:14,fontWeight:700,cursor:"pointer"}}>
        {saving?"Saving…":"💾 Save"}
      </button>
    </div>
  );
}

// ── Feedback read-only view (presidents) ──────────────────────────────────────
function FeedbackView({candidate,interviewData,members,round,onViewAnswers}){
  const all=interviewData.filter(i=>i.candidate_id===candidate.id&&i.round===round);
  if(!all.length)return(
    <div style={{padding:20,background:C.white,border:`1px solid ${C.border}`,borderTop:"none",borderRadius:"0 0 12px 12px",color:C.textLt,fontSize:13,textAlign:"center"}}>
      No feedback submitted yet.
    </div>
  );
  return(
    <div style={{background:C.white,border:`1px solid ${C.border}`,borderTop:"none",borderRadius:"0 0 12px 12px",overflow:"hidden"}}>
      {all.map((iv,idx)=>{
        const name=members.find(m=>m.id===iv.interviewer_id)?.name||iv.interviewer_id;
        return(
          <div key={iv.interviewer_id} style={{padding:18,borderTop:idx>0?`1px solid ${C.border}`:"none",background:idx%2===0?"#fff":"#fafbff"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
              <div style={{fontSize:11,color:C.navy,letterSpacing:1.5,fontWeight:700}}>{name.toUpperCase()}</div>
              <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                {iv.interview_date&&<span style={{fontSize:12,color:C.textMid,fontWeight:600}}>📅 {fmtDate(iv.interview_date)}</span>}
                <VerdictBadge verdict={iv.verdict}/>
              </div>
            </div>
            <div style={{display:"flex",gap:10,marginBottom:12,flexWrap:"wrap"}}>
              {[["Personal",iv.personal_score],["Technical",iv.technical_score],["Brainstormer",iv.brainstormer_score]].map(([lbl,val])=>(
                <div key={lbl} style={{background:C.bg,borderRadius:8,padding:"10px 14px",minWidth:90,textAlign:"center"}}>
                  <div style={{fontSize:9,color:C.textLt,marginBottom:4,fontWeight:700,letterSpacing:1}}>{lbl.toUpperCase()}</div>
                  <div style={{fontSize:22,fontWeight:900,color:val!=null?scoreColor((parseFloat(val)/4)*100):C.textLt}}>{fmtScore(val)}<span style={{fontSize:11,color:C.textLt,fontWeight:400}}>/4</span></div>
                </div>
              ))}
            </div>
            {iv.feedback&&<div style={{fontSize:13,color:C.textMid,lineHeight:1.75,background:C.bg,borderRadius:8,padding:"12px 14px",whiteSpace:"pre-wrap"}}>{iv.feedback}</div>}
          </div>
        );
      })}
    </div>
  );
}

// ── Import panel ──────────────────────────────────────────────────────────────
function ImportPanel({onImportReplace,onImportAppend,loading,hasCandidates}){
  const [drag,setDrag]=useState(false);
  const [err,setErr]=useState(null);
  const [parsedRows,setParsedRows]=useState(null);
  const [mode,setMode]=useState(null); // null | "replace" | "append"
  const [stats,setStats]=useState(null); // {total, new, duplicate}

  const parse=(f,existingCandidates)=>{
    if(!f)return;
    setErr(null);setParsedRows(null);setMode(null);setStats(null);
    Papa.parse(f,{header:true,skipEmptyLines:true,complete:r=>{
      if(!r.data?.length){setErr("CSV appears empty.");return;}
      const seen={};
      r.data.forEach(row=>{const sn=getCol(row,"student")?.trim();if(!sn||sn==="teste")return;if(!seen[sn]||getCol(row,"timestamp")>getCol(seen[sn],"timestamp"))seen[sn]=row;});
      const deduped=Object.values(seen);
      if(!deduped.length){setErr("No valid candidates found.");return;}
      // Calculate stats if we have existing candidates
      if(hasCandidates&&existingCandidates){
        const existingIds=new Set((existingCandidates||[]).map(c=>c.student_number));
        const newOnly=deduped.filter(r=>!existingIds.has(getCol(r,"student")?.trim()));
        const dupes=deduped.length-newOnly.length;
        setStats({total:deduped.length,new:newOnly.length,duplicate:dupes});
      }
      setParsedRows(deduped);
    },error:()=>setErr("Failed to parse CSV.")});
  };

  const handleConfirm=()=>{
    if(!parsedRows)return;
    if(mode==="replace")onImportReplace(parsedRows);
    else if(mode==="append")onImportAppend(parsedRows);
    setParsedRows(null);setMode(null);setStats(null);
  };

  return(
    <div style={{background:C.white,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:22}}>
      <div style={{fontSize:11,fontWeight:700,color:C.navy,letterSpacing:1,marginBottom:12}}>📂 IMPORT APPLICATIONS</div>

      {/* Dropzone */}
      {!parsedRows&&(
        <div style={{border:`2px dashed ${drag?C.navy:C.border}`,borderRadius:9,padding:"24px 16px",cursor:"pointer",textAlign:"center",background:drag?"#eef3ff":C.bg,transition:"all 0.2s"}}
          onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
          onDrop={e=>{e.preventDefault();setDrag(false);parse(e.dataTransfer.files[0]);}}
          onClick={()=>document.getElementById("csvImport").click()}>
          <div style={{fontSize:32,marginBottom:6}}>📄</div>
          <p style={{margin:"0 0 3px",fontWeight:700,color:C.text,fontSize:14}}>Drop CSV or click to upload</p>
          <p style={{margin:0,fontSize:12,color:C.textLt}}>Duplicates auto-removed by Student ID</p>
          <input id="csvImport" type="file" accept=".csv" style={{display:"none"}} onChange={e=>parse(e.target.files[0])}/>
        </div>
      )}

      {/* After parsing — show options */}
      {parsedRows&&!mode&&(
        <div>
          <div style={{background:C.bg,borderRadius:9,padding:"14px 16px",marginBottom:14}}>
            <div style={{fontSize:14,fontWeight:700,color:C.navy,marginBottom:6}}>📋 {parsedRows.length} candidates found in CSV</div>
            {stats&&(
              <div style={{fontSize:13,color:C.textMid}}>
                <span style={{color:C.green,fontWeight:700}}>{stats.new} new</span>
                {stats.duplicate>0&&<span> · <span style={{color:C.amber,fontWeight:600}}>{stats.duplicate} already exist</span></span>}
              </div>
            )}
          </div>

          <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
            {hasCandidates&&(
              <button onClick={()=>setMode("append")}
                style={{flex:1,padding:"14px 16px",background:"#f0fdf4",border:`2px solid #86efac`,borderRadius:10,cursor:"pointer",textAlign:"left"}}>
                <div style={{fontSize:14,fontWeight:700,color:C.green,marginBottom:4}}>➕ Add New Only</div>
                <div style={{fontSize:12,color:C.textMid}}>
                  {stats?`Add ${stats.new} new candidates. Skip ${stats.duplicate} duplicates. Keep all existing scores and rounds.`:"Add only candidates that don't exist yet. Keep everything else."}
                </div>
              </button>
            )}
            <button onClick={()=>setMode("replace")}
              style={{flex:1,padding:"14px 16px",background:"#fef2f2",border:`2px solid #fca5a5`,borderRadius:10,cursor:"pointer",textAlign:"left"}}>
              <div style={{fontSize:14,fontWeight:700,color:C.red,marginBottom:4}}>🔄 Replace All</div>
              <div style={{fontSize:12,color:C.textMid}}>Delete everything and start fresh with {parsedRows.length} candidates. Scores, interviews, groups — all reset.</div>
            </button>
          </div>
        </div>
      )}

      {/* Confirmation */}
      {parsedRows&&mode&&(
        <div style={{background:mode==="append"?"#f0fdf4":"#fef2f2",border:`1px solid ${mode==="append"?"#86efac":"#fca5a5"}`,borderRadius:9,padding:"14px 16px"}}>
          <div style={{fontSize:14,fontWeight:700,color:mode==="append"?C.green:C.red,marginBottom:8}}>
            {mode==="append"?`➕ Add ${stats?.new||parsedRows.length} new candidates?`:`🔄 Replace all with ${parsedRows.length} candidates?`}
          </div>
          {mode==="replace"&&<div style={{fontSize:12,color:C.red,marginBottom:10}}>⚠️ This will delete ALL existing candidates, scores, interviews, and groups.</div>}
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>{setMode(null);}} style={{flex:1,padding:"9px",background:"#fff",border:`1px solid ${C.border}`,borderRadius:8,fontSize:13,fontWeight:600,cursor:"pointer",color:C.textMid}}>← Back</button>
            <button onClick={handleConfirm} disabled={loading}
              style={{flex:1,padding:"9px",background:mode==="append"?C.green:C.red,color:"#fff",border:"none",borderRadius:8,fontSize:13,fontWeight:700,cursor:"pointer"}}>
              {loading?"Importing…":"✓ Confirm"}
            </button>
          </div>
        </div>
      )}

      {err&&<p style={{color:C.red,fontSize:13,marginTop:10}}>{err}</p>}
      {loading&&<div style={{display:"flex",justifyContent:"center",marginTop:12}}><Spinner/></div>}

      {/* Reset to re-upload */}
      {parsedRows&&!loading&&(
        <button onClick={()=>{setParsedRows(null);setMode(null);setStats(null);setErr(null);}}
          style={{background:"none",border:"none",color:C.textLt,fontSize:12,cursor:"pointer",marginTop:10,padding:0}}>
          ← Upload different file
        </button>
      )}
    </div>
  );
}

// ── Password Gate ─────────────────────────────────────────────────────────────
const PORTAL_PASSWORD = "NIC_UD_26/27";

function PasswordGate({onPass}){
  const [input, setInput] = useState("");
  const [error, setError] = useState(false);
  const [shake, setShake] = useState(false);

  const attempt = () => {
    if(input === PORTAL_PASSWORD){
      sessionStorage.setItem("nic_auth","1");
      onPass();
    } else {
      setError(true);
      setShake(true);
      setInput("");
      setTimeout(()=>setShake(false), 600);
    }
  };

  return(
    <div style={{minHeight:"100vh",background:C.navy,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"system-ui,sans-serif",padding:16,position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",inset:0,backgroundImage:"url(/team-bg.jpg)",backgroundSize:"cover",backgroundPosition:"center",opacity:0.35}}/>
      <div style={{position:"absolute",inset:0,background:"linear-gradient(180deg,rgba(15,41,82,0.55) 0%,rgba(15,41,82,0.75) 100%)"}}/>
      <style>{`@keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-8px)}40%,80%{transform:translateX(8px)}}`}</style>
      <div style={{position:"relative",zIndex:1,background:"rgba(255,255,255,0.07)",backdropFilter:"blur(12px)",border:"1px solid rgba(255,255,255,0.12)",borderRadius:16,padding:"40px 36px",width:"100%",maxWidth:400,boxShadow:"0 8px 32px rgba(0,0,0,0.3)",textAlign:"center",animation:shake?"shake 0.5s ease":"none"}}>
        <div style={{marginBottom:20}}>
          <div style={{display:"inline-block",background:"#fff",borderRadius:12,padding:"10px 18px"}}><img src="/nic-logo.png" alt="NIC" style={{height:60,objectFit:"contain"}}/></div>
        </div>
        <div style={{height:1,background:"rgba(255,255,255,0.1)",marginBottom:24}}/>
        <h2 style={{fontSize:18,fontWeight:700,color:"#fff",margin:"0 0 6px"}}>Recruitment Portal</h2>
        <p style={{fontSize:13,color:"rgba(255,255,255,0.5)",marginBottom:24}}>Enter the access password to continue</p>
        <input
          type="password"
          value={input}
          onChange={e=>{setInput(e.target.value);setError(false);}}
          onKeyDown={e=>e.key==="Enter"&&attempt()}
          placeholder="Password"
          autoFocus
          style={{width:"100%",border:`1.5px solid ${error?"#ef4444":"rgba(255,255,255,0.2)"}`,borderRadius:9,padding:"11px 14px",fontSize:15,color:"#fff",outline:"none",marginBottom:error?8:16,textAlign:"center",letterSpacing:2,background:error?"rgba(239,68,68,0.15)":"rgba(255,255,255,0.08)",transition:"border 0.2s"}}
        />
        {error&&<p style={{color:"#fca5a5",fontSize:13,marginBottom:12,fontWeight:600}}>Incorrect password. Try again.</p>}
        <button onClick={attempt}
          style={{width:"100%",padding:"12px",background:"#fff",color:C.navy,border:"none",borderRadius:9,fontSize:15,fontWeight:700,cursor:"pointer"}}>
          Enter →
        </button>
        <p style={{fontSize:11,color:"rgba(255,255,255,0.3)",marginTop:16}}>NIC — Undergraduate Division</p>
      </div>
    </div>
  );
}

// ── Login ─────────────────────────────────────────────────────────────────────
function LoginScreen({members,onLogin}){
  const [sel,setSel]=useState(null);
  const presidents=members.filter(m=>m.role==="president");
  const regular=members.filter(m=>m.role!=="president");

  const handleLogin=()=>{if(sel)onLogin(sel);};

  return(
    <div style={{minHeight:"100vh",background:C.navy,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"system-ui,sans-serif",padding:16,position:"relative",overflow:"hidden"}}>
      {/* Background team photo */}
      <div style={{position:"absolute",inset:0,backgroundImage:"url(/team-bg.jpg)",backgroundSize:"cover",backgroundPosition:"center",opacity:0.35}}/>
      <div style={{position:"absolute",inset:0,background:"linear-gradient(180deg,rgba(15,41,82,0.55) 0%,rgba(15,41,82,0.75) 100%)"}}/>

      <div style={{position:"relative",zIndex:1,width:"100%",maxWidth:560}}>
        {/* Header */}
        <div style={{textAlign:"center",marginBottom:20}}>
          <div style={{display:"inline-block",background:"#fff",borderRadius:12,padding:"8px 16px"}}><img src="/nic-logo.png" alt="NIC" style={{height:54,objectFit:"contain"}}/></div>
          <h2 style={{fontSize:20,fontWeight:700,color:"#fff",margin:"12px 0 4px"}}>Recruitment Portal</h2>
          <p style={{fontSize:13,color:"rgba(255,255,255,0.5)"}}>Select your profile to continue</p>
        </div>

        {/* Card with scrollable member list */}
        <div style={{background:"rgba(255,255,255,0.07)",backdropFilter:"blur(12px)",border:"1px solid rgba(255,255,255,0.12)",borderRadius:16,padding:"20px 22px",maxHeight:"60vh",overflowY:"auto"}}>
          {/* Presidents */}
          {presidents.length>0&&(
            <div style={{marginBottom:14}}>
              <div style={{fontSize:10,color:"rgba(255,255,255,0.4)",letterSpacing:2,fontWeight:700,marginBottom:8}}>CO-PRESIDENTS</div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                {presidents.map(m=>(
                  <button key={m.id} onClick={()=>setSel(m)}
                    style={{flex:"1 1 180px",display:"flex",alignItems:"center",gap:8,padding:"9px 12px",background:sel?.id===m.id?"rgba(255,255,255,0.2)":"rgba(255,255,255,0.06)",border:`1.5px solid ${sel?.id===m.id?"rgba(255,255,255,0.5)":"rgba(255,255,255,0.1)"}`,borderRadius:9,cursor:"pointer",color:"#fff"}}>
                    <span style={{width:26,height:26,borderRadius:"50%",background:sel?.id===m.id?"rgba(255,255,255,0.3)":"rgba(255,255,255,0.1)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,flexShrink:0}}>{initials(m.name)}</span>
                    <span style={{fontSize:12,fontWeight:600}}>{m.name}</span>
                    <span style={{marginLeft:"auto",fontSize:8,background:"rgba(255,255,255,0.15)",borderRadius:3,padding:"2px 5px",letterSpacing:1,fontWeight:700}}>PRES</span>
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* Members */}
          <div style={{fontSize:10,color:"rgba(255,255,255,0.4)",letterSpacing:2,fontWeight:700,marginBottom:8}}>MEMBERS</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:5}}>
            {regular.map(m=>(
              <button key={m.id} onClick={()=>setSel(m)}
                style={{display:"flex",alignItems:"center",gap:7,padding:"8px 10px",background:sel?.id===m.id?"rgba(255,255,255,0.2)":"rgba(255,255,255,0.04)",border:`1.5px solid ${sel?.id===m.id?"rgba(255,255,255,0.5)":"rgba(255,255,255,0.08)"}`,borderRadius:8,cursor:"pointer",color:sel?.id===m.id?"#fff":"rgba(255,255,255,0.7)"}}>
                <span style={{width:24,height:24,borderRadius:"50%",background:sel?.id===m.id?"rgba(255,255,255,0.3)":"rgba(255,255,255,0.08)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:800,flexShrink:0,color:sel?.id===m.id?"#fff":"rgba(255,255,255,0.5)"}}>{initials(m.name)}</span>
                <span style={{fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.name}</span>
              </button>
            ))}
          </div>
        </div>

        {/* Enter button — always visible, fixed outside scroll area */}
        <button onClick={handleLogin} disabled={!sel}
          style={{width:"100%",padding:"13px",marginTop:16,background:sel?"#fff":"rgba(255,255,255,0.15)",color:sel?C.navy:"rgba(255,255,255,0.3)",border:"none",borderRadius:10,fontSize:15,fontWeight:700,cursor:sel?"pointer":"not-allowed",boxShadow:sel?"0 4px 16px rgba(0,0,0,0.2)":"none",transition:"all 0.2s"}}>
          {sel?`Enter as ${sel.name} →`:"Select your name above"}
        </button>
      </div>
    </div>
  );
}

// ══ MAIN APP ══════════════════════════════════════════════════════════════════
// Error boundary wrapper
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { hasError: false, error: null }; }
  static getDerivedStateFromError(error) { return { hasError: true, error }; }
  render() {
    if (this.state.hasError) {
      return React.createElement('div', {style:{minHeight:"100vh",background:C.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"system-ui,sans-serif",padding:20,textAlign:"center"}},
        React.createElement('div', {style:{fontSize:48,marginBottom:16}}, '⚠️'),
        React.createElement('h2', {style:{color:C.navy,margin:"0 0 8px"}}, 'Something went wrong'),
        React.createElement('p', {style:{color:C.textMid,fontSize:14,marginBottom:16}}, this.state.error?.message || 'Unknown error'),
        React.createElement('button', {onClick:()=>{this.setState({hasError:false});window.location.reload();},style:{background:C.navy,color:"#fff",border:"none",borderRadius:8,padding:"10px 20px",fontSize:14,fontWeight:700,cursor:"pointer"}}, 'Reload')
      );
    }
    return this.props.children;
  }
}

function AppInner(){
  const [authed,        setAuthed]        = useState(()=>sessionStorage.getItem("nic_auth")==="1");
  const [members,       setMembers]       = useState([]);
  const [user,          setUser]          = useState(null);
  const [candidates,    setCandidates]    = useState(null);
  const [allScores,     setAllScores]     = useState([]);
  const [aiScores,      setAiScores]      = useState({});
  const [revealed,      setRevealed]      = useState(false);
  const [topN,          setTopN]          = useState(20);
  const [view,          setView]          = useState("list");
  const [selected,      setSelected]      = useState(null);
  const [detecting,     setDetecting]     = useState(false);
  const [importing,     setImporting]     = useState(false);
  const [appLoading,    setAppLoading]    = useState(true);
  const [toast,         setToast]         = useState(null);
  const [search,        setSearch]        = useState("");
  const [showImport,    setShowImport]    = useState(false);
  const [interviewData, setInterviewData] = useState([]);
  const [assignments,   setAssignments]   = useState([]);
  const [promoted,      setPromoted]      = useState({});
  const [assignModal,   setAssignModal]   = useState(null);
  const [answersModal,  setAnswersModal]  = useState(null);
  const [chosenSet,     setChosenSet]     = useState({}); // {candidateId: true}
  const [sidebarOpen,   setSidebarOpen]   = useState(false);
  const [deadline,      setDeadline]      = useState("");
  const [evalGroups,    setEvalGroups]    = useState([]);     // [{id,name}]
  const [groupMembers,  setGroupMembers]  = useState([]);     // [{group_id,member_id}]
  const [groupCandidates,setGroupCandidates]=useState([]);    // [{group_id,candidate_id}]
  const [showAllMode,   setShowAllMode]   = useState(false);

  const showToast=(msg,type="ok")=>{setToast({msg,type});setTimeout(()=>setToast(null),3000);};

  useEffect(()=>{
    sb.from("members").select("*").order("name").then(({data})=>{if(data)setMembers(data);});
  },[]);

  useEffect(()=>{
    if(!user)return;
    setAppLoading(true);
    // Safe loader — if a table doesn't exist, returns empty array instead of crashing
    const safe=async(query)=>{try{const r=await query;return r.data||[];}catch{return[];}};
    Promise.all([
      safe(sb.from("candidates").select("*").order("student_number")),
      safe(sb.from("scores").select("*")),
      safe(sb.from("ai_scores").select("*")),
      safe(sb.from("settings").select("*")),
      safe(sb.from("interview_feedback").select("*")),
      safe(sb.from("interview_assignments").select("*")),
      safe(sb.from("candidate_promotions").select("*")),
      safe(sb.from("chosen_candidates").select("*")),
      safe(sb.from("eval_groups").select("*")),
      safe(sb.from("eval_group_members").select("*")),
      safe(sb.from("eval_group_candidates").select("*")),
    ]).then(([c,sc,ai,cfg,ivf,iva,promo,chosen,eg,egm,egc])=>{
      setCandidates(c.length?c:null);
      if(sc.length)setAllScores(sc);
      if(ai.length){const m={};ai.forEach(r=>{m[r.candidate_id]=r;});setAiScores(m);}
      if(cfg.length){
        const rev=cfg.find(r=>r.key==="revealed");
        const tn=cfg.find(r=>r.key==="top_n");
        const dl=cfg.find(r=>r.key==="deadline");
        if(rev)setRevealed(rev.value==="true");
        if(tn)setTopN(parseInt(tn.value)||20);
        if(dl)setDeadline(dl.value||"");
        const saMode=cfg.find(r=>r.key==="show_all_candidates");
        if(saMode)setShowAllMode(saMode.value==="true");
      }
      if(ivf.length)setInterviewData(ivf);
      if(iva.length)setAssignments(iva);
      if(promo.length){const m={};promo.forEach(r=>{m[r.candidate_id]=r.round;});setPromoted(m);}
      if(chosen.length){const m={};chosen.forEach(r=>{m[r.candidate_id]=true;});setChosenSet(m);}
      if(eg.length)setEvalGroups(eg);
      if(egm.length)setGroupMembers(egm);
      if(egc.length)setGroupCandidates(egc);
      setAppLoading(false);
    }).catch(()=>setAppLoading(false));
  },[user]);

  // ── Realtime ─────────────────────────────────────────────────────────────
  useEffect(()=>{
    if(!user)return;
    const ch=sb.channel("nic-live-v3")
      .on("postgres_changes",{event:"*",schema:"public",table:"scores"},p=>{
        setAllScores(prev=>{
          if(p.eventType==="DELETE")return prev.filter(r=>r.id!==p.old.id);
          const f=prev.filter(r=>!(r.member_id===p.new.member_id&&r.candidate_id===p.new.candidate_id&&r.question_id===p.new.question_id));
          return[...f,p.new];
        });
      })
      .on("postgres_changes",{event:"*",schema:"public",table:"settings"},p=>{
        if(p.new?.key==="revealed")setRevealed(p.new.value==="true");
        if(p.new?.key==="top_n")setTopN(parseInt(p.new.value)||20);
        if(p.new?.key==="deadline")setDeadline(p.new.value||"");
        if(p.new?.key==="show_all_candidates")setShowAllMode(p.new.value==="true");
      })
      .on("postgres_changes",{event:"INSERT",schema:"public",table:"candidates"},p=>{
        setCandidates(prev=>prev?[...prev,p.new]:[p.new]);
      })
      .on("postgres_changes",{event:"*",schema:"public",table:"ai_scores"},p=>{
        if(p.new)setAiScores(prev=>({...prev,[p.new.candidate_id]:p.new}));
      })
      .on("postgres_changes",{event:"*",schema:"public",table:"interview_feedback"},p=>{
        // FIX: handle all event types correctly
        setInterviewData(prev=>{
          if(p.eventType==="DELETE")return prev.filter(r=>!(r.interviewer_id===p.old.interviewer_id&&r.candidate_id===p.old.candidate_id&&r.round===p.old.round));
          const f=prev.filter(r=>!(r.interviewer_id===p.new.interviewer_id&&r.candidate_id===p.new.candidate_id&&r.round===p.new.round));
          return[...f,p.new];
        });
      })
      .on("postgres_changes",{event:"*",schema:"public",table:"interview_assignments"},p=>{
        // FIX: handle DELETE correctly
        setAssignments(prev=>{
          if(p.eventType==="DELETE")return prev.filter(r=>!(r.candidate_id===p.old.candidate_id&&r.round===p.old.round&&r.interviewer_id===p.old.interviewer_id));
          return[...prev.filter(r=>!(r.candidate_id===p.new.candidate_id&&r.round===p.new.round&&r.interviewer_id===p.new.interviewer_id)),p.new];
        });
      })
      .on("postgres_changes",{event:"*",schema:"public",table:"candidate_promotions"},p=>{
        if(p.eventType==="DELETE")setPromoted(prev=>{const n={...prev};delete n[p.old.candidate_id];return n;});
        else if(p.new)setPromoted(prev=>({...prev,[p.new.candidate_id]:p.new.round}));
      })
      .on("postgres_changes",{event:"*",schema:"public",table:"chosen_candidates"},p=>{
        if(p.eventType==="DELETE")setChosenSet(prev=>{const n={...prev};delete n[p.old.candidate_id];return n;});
        else if(p.new)setChosenSet(prev=>({...prev,[p.new.candidate_id]:true}));
      })
      .subscribe();
    return()=>{sb.removeChannel(ch);};
  },[user]);

  // ── Handlers ─────────────────────────────────────────────────────────────
  // Map CSV rows to candidate objects
  const mapRows=(rows)=>rows.map((r,i)=>({
    id:`c_${Date.now()}_${i}_${Math.random().toString(36).slice(2,6)}`,
    full_name:extractName(r)||getCol(r,"student")||`Candidate ${i+1}`,
    student_number:getCol(r,"student")||"",
    email:getCol(r,"email")||"",
    phone:getCol(r,"phone")||"",
    cv_link:getCol(r,"cv")||"",
    b1:getCol(r,"b1")||"",
    b2:getCol(r,"b2")||"",
    b3:getCol(r,"b3")||"",
    t1:getCol(r,"t1")||"",
    comments:getCol(r,"comments")||"",
    submitted_at:getCol(r,"timestamp")||"",
  }));

  const insertCandidates=async(mapped)=>{
    let{error}=await sb.from("candidates").insert(mapped);
    if(error?.message?.includes("column")){
      const fallback=mapped.map(({email,phone,submitted_at,comments,b3,...rest})=>rest);
      ({error}=await sb.from("candidates").insert(fallback));
    }
    return error;
  };

  // Replace all — wipe everything and start fresh
  const handleImportReplace=useCallback(async rows=>{
    setImporting(true);
    await Promise.all([
      sb.from("eval_group_candidates").delete().neq("group_id","_"),
      sb.from("chosen_candidates").delete().neq("candidate_id","_"),
      sb.from("interview_feedback").delete().neq("candidate_id","_"),
      sb.from("interview_assignments").delete().neq("candidate_id","_"),
      sb.from("candidate_promotions").delete().neq("candidate_id","_"),
      sb.from("scores").delete().neq("id","00000000-0000-0000-0000-000000000000"),
      sb.from("ai_scores").delete().neq("candidate_id","_"),
      sb.from("candidates").delete().neq("id","_"),
    ]);
    await sb.from("settings").upsert({key:"revealed",value:"false"},{onConflict:"key"});
    setAllScores([]); setAiScores({}); setInterviewData([]); setAssignments([]); setPromoted({}); setChosenSet({}); setGroupCandidates([]); setRevealed(false);
    const mapped=mapRows(rows);
    const error=await insertCandidates(mapped);
    if(error)showToast("Import failed: "+error.message,"err");
    else{setCandidates(mapped);setShowImport(false);showToast(`${mapped.length} candidates imported (fresh start)`,"ok");}
    setImporting(false);
  },[]);

  // Append — add only new candidates (skip duplicates by student_number)
  const handleImportAppend=useCallback(async rows=>{
    setImporting(true);
    const existingIds=new Set((candidates||[]).map(c=>c.student_number));
    const newRows=rows.filter(r=>{
      const sn=getCol(r,"student")?.trim();
      return sn&&!existingIds.has(sn);
    });
    if(!newRows.length){
      showToast("No new candidates to add — all already exist","err");
      setImporting(false);
      return;
    }
    const mapped=mapRows(newRows);
    const error=await insertCandidates(mapped);
    if(error){
      showToast("Import failed: "+error.message,"err");
    } else {
      setCandidates(prev=>prev?[...prev,...mapped]:mapped);
      // Auto-distribute new candidates to groups if groups exist
      if(evalGroups.length>0&&groupCandidates.length>0){
        // Count candidates per group
        const groupCounts={};
        evalGroups.forEach(g=>{groupCounts[g.id]=groupCandidates.filter(gc=>gc.group_id===g.id).length;});
        // Distribute new candidates to least-loaded groups
        const newGroupAssignments=[];
        mapped.forEach(c=>{
          // Find group with fewest candidates
          const minGroup=evalGroups.reduce((min,g)=>(groupCounts[g.id]||0)<(groupCounts[min.id]||0)?g:min,evalGroups[0]);
          newGroupAssignments.push({group_id:minGroup.id,candidate_id:c.id});
          groupCounts[minGroup.id]=(groupCounts[minGroup.id]||0)+1;
        });
        if(newGroupAssignments.length){
          await sb.from("eval_group_candidates").insert(newGroupAssignments);
          setGroupCandidates(prev=>[...prev,...newGroupAssignments]);
        }
        showToast(`${mapped.length} new candidates added & distributed to groups`,"ok");
      } else {
        showToast(`${mapped.length} new candidates added`,"ok");
      }
      setShowImport(false);
    }
    setImporting(false);
  },[candidates,evalGroups,groupCandidates]);

  const handleScore=useCallback(async(candidateId,questionId,value)=>{
    setAllScores(prev=>{
      const idx=prev.findIndex(s=>s.member_id===user.id&&s.candidate_id===candidateId&&s.question_id===questionId);
      const row={member_id:user.id,candidate_id:candidateId,question_id:questionId,score:value,id:`opt_${Date.now()}`};
      return idx>=0?prev.map((s,i)=>i===idx?{...s,score:value}:s):[...prev,row];
    });
    await sb.from("scores").upsert({member_id:user.id,candidate_id:candidateId,question_id:questionId,score:value,updated_at:new Date().toISOString()},{onConflict:"member_id,candidate_id,question_id"});
  },[user]);

  const runDetection=useCallback(async candidate=>{
    if(aiScores[candidate.id])return;
    setDetecting(true);
    const result=detectAI(candidate);
    setDetecting(false);
    const row={candidate_id:candidate.id,b1_pct:result.b1,b2_pct:result.b2,b3_pct:result.b3,t1_pct:result.t1,overall_pct:result.overall,flags:result.flags};
    await sb.from("ai_scores").upsert(row,{onConflict:"candidate_id"});
    setAiScores(prev=>({...prev,[candidate.id]:row}));
  },[aiScores]);

  const handleReveal=useCallback(async()=>{
    await sb.from("settings").upsert({key:"revealed",value:"true"},{onConflict:"key"});
    setRevealed(true);showToast("Results revealed to all members","ok");
  },[]);
  const handleUnreveal=useCallback(async()=>{
    await sb.from("settings").upsert({key:"revealed",value:"false"},{onConflict:"key"});
    setRevealed(false);showToast("Results hidden from members","ok");
  },[]);
  const handleTopN=useCallback(async n=>{
    setTopN(n);
    await sb.from("settings").upsert({key:"top_n",value:String(n)},{onConflict:"key"});
  },[]);
  const handlePromote=useCallback(async(candidateId,round)=>{
    await sb.from("candidate_promotions").upsert({candidate_id:candidateId,round},{onConflict:"candidate_id"});
    setPromoted(prev=>({...prev,[candidateId]:round}));
    showToast(`Promoted to ${round==="interview"?"Interview":"Final"} Round`,"ok");
  },[]);
  const handleDemote=useCallback(async(candidateId)=>{
    await sb.from("candidate_promotions").delete().eq("candidate_id",candidateId);
    await sb.from("interview_assignments").delete().eq("candidate_id",candidateId);
    await sb.from("interview_feedback").delete().eq("candidate_id",candidateId);
    setPromoted(prev=>{const n={...prev};delete n[candidateId];return n;});
    setAssignments(prev=>prev.filter(a=>a.candidate_id!==candidateId));
    setInterviewData(prev=>prev.filter(i=>i.candidate_id!==candidateId));
    showToast("Promotion cancelled","ok");
  },[]);
  const handleAssign=useCallback(async(candidateId,round,interviewerIds)=>{
    await sb.from("interview_assignments").delete().eq("candidate_id",candidateId).eq("round",round);
    if(interviewerIds.length)await sb.from("interview_assignments").insert(interviewerIds.map(id=>({candidate_id:candidateId,round,interviewer_id:id})));
    setAssignments(prev=>[...prev.filter(a=>!(a.candidate_id===candidateId&&a.round===round)),...interviewerIds.map(id=>({candidate_id:candidateId,round,interviewer_id:id}))]);
    showToast("Interviewers assigned","ok");
    setAssignModal(null);
  },[]);
  const handleInterviewSave=useCallback(async(data)=>{
    const payload={...data,updated_at:new Date().toISOString()};
    await sb.from("interview_feedback").upsert(payload,{onConflict:"candidate_id,interviewer_id,round"});
    setInterviewData(prev=>{
      const f=prev.filter(r=>!(r.interviewer_id===data.interviewer_id&&r.candidate_id===data.candidate_id&&r.round===data.round));
      return[...f,payload];
    });
    showToast("Feedback saved","ok");
  },[]);
  const handleSetDeadline=useCallback(async(date)=>{
    setDeadline(date);
    await sb.from("settings").upsert({key:"deadline",value:date},{onConflict:"key"});
    showToast(date?"Deadline set for all members":"Deadline removed","ok");
  },[]);

  // ── Group management ────────────────────────────────────────────────────
  const handleCreateGroup=useCallback(async(name)=>{
    const id=`g_${Date.now()}`;
    const {error}=await sb.from("eval_groups").insert({id,name});
    if(error){showToast("Failed to create group","err");return;}
    setEvalGroups(prev=>[...prev,{id,name}]);
    showToast(`Group "${name}" created`,"ok");
  },[]);

  const handleDeleteGroup=useCallback(async(groupId)=>{
    await sb.from("eval_group_candidates").delete().eq("group_id",groupId);
    await sb.from("eval_group_members").delete().eq("group_id",groupId);
    await sb.from("eval_groups").delete().eq("id",groupId);
    setEvalGroups(prev=>prev.filter(g=>g.id!==groupId));
    setGroupMembers(prev=>prev.filter(m=>m.group_id!==groupId));
    setGroupCandidates(prev=>prev.filter(c=>c.group_id!==groupId));
    showToast("Group deleted","ok");
  },[]);

  const handleAssignMemberToGroup=useCallback(async(groupId,memberId,add)=>{
    if(add){
      // Remove from other groups first
      await sb.from("eval_group_members").delete().eq("member_id",memberId);
      await sb.from("eval_group_members").insert({group_id:groupId,member_id:memberId});
      setGroupMembers(prev=>[...prev.filter(m=>m.member_id!==memberId),{group_id:groupId,member_id:memberId}]);
    } else {
      await sb.from("eval_group_members").delete().eq("group_id",groupId).eq("member_id",memberId);
      setGroupMembers(prev=>prev.filter(m=>!(m.group_id===groupId&&m.member_id===memberId)));
    }
  },[]);

  const handleDistributeCandidates=useCallback(async()=>{
    if(!evalGroups.length||!candidates?.length)return;
    // Delete existing distribution
    await sb.from("eval_group_candidates").delete().neq("group_id","_");
    // Shuffle candidates
    const shuffled=[...candidates].sort(()=>Math.random()-0.5);
    const perGroup=Math.ceil(shuffled.length/evalGroups.length);
    const rows=[];
    shuffled.forEach((c,i)=>{
      const groupIdx=Math.floor(i/perGroup);
      const group=evalGroups[groupIdx]||evalGroups[evalGroups.length-1];
      rows.push({group_id:group.id,candidate_id:c.id});
    });
    await sb.from("eval_group_candidates").insert(rows);
    setGroupCandidates(rows);
    showToast(`${candidates.length} candidates distributed across ${evalGroups.length} groups`,"ok");
  },[evalGroups,candidates]);

  const handleToggleShowAll=useCallback(async()=>{
    const newVal=!showAllMode;
    setShowAllMode(newVal);
    await sb.from("settings").upsert({key:"show_all_candidates",value:newVal?"true":"false"},{onConflict:"key"});
    showToast(newVal?"All candidates visible to everyone":"Group-based view active","ok");
  },[showAllMode]);

  const handleToggleChosen=useCallback(async(candidateId)=>{
    if(chosenSet[candidateId]){
      await sb.from("chosen_candidates").delete().eq("candidate_id",candidateId);
      setChosenSet(prev=>{const n={...prev};delete n[candidateId];return n;});
      showToast("Removed from chosen","ok");
    } else {
      await sb.from("chosen_candidates").upsert({candidate_id:candidateId},{onConflict:"candidate_id"});
      setChosenSet(prev=>({...prev,[candidateId]:true}));
      showToast("Added to chosen","ok");
    }
  },[chosenSet]);

  // ── Derived ──────────────────────────────────────────────────────────────
  const isPresident=user?.role==="president";
  const myScores=useMemo(()=>allScores.filter(s=>s.member_id===user?.id),[allScores,user]);

  // Group filtering — must be defined before progress
  const myGroupId=useMemo(()=>{
    const gm=groupMembers.find(m=>m.member_id===user?.id);
    return gm?.group_id||null;
  },[groupMembers,user]);
  const myCandidates=useMemo(()=>{
    if(!candidates)return [];
    if(isPresident)return candidates;
    if(showAllMode)return candidates;
    if(!evalGroups.length||!groupCandidates.length)return candidates;
    if(!myGroupId)return [];
    const myIds=new Set(groupCandidates.filter(gc=>gc.group_id===myGroupId).map(gc=>gc.candidate_id));
    return candidates.filter(c=>myIds.has(c.id));
  },[candidates,isPresident,showAllMode,evalGroups,groupCandidates,myGroupId]);

  const progress=useMemo(()=>user?memberProgress(user.id,myCandidates,allScores):{done:0,total:0},[user,myCandidates,allScores]);
  const allDone=useMemo(()=>members.length>0&&(candidates?.length||0)>0&&members.every(m=>{const p=memberProgress(m.id,candidates,allScores);return p.done===p.total&&p.total>0;}),[members,candidates,allScores]);
  const filtered=useMemo(()=>myCandidates.filter(c=>!search||c.student_number?.includes(search)||c.email?.toLowerCase().includes(search.toLowerCase())||c.full_name?.toLowerCase().includes(search.toLowerCase())),[myCandidates,search]);
  const ranked=useMemo(()=>(candidates||[]).map(c=>({...c,avg:avgScore(c.id,allScores)})).sort((a,b)=>(parseFloat(b.avg)||0)-(parseFloat(a.avg)||0)),[candidates,allScores]);
  const candidateAI=selected?(aiScores[selected.id]||{}):{};
  const interviewCandidates=useMemo(()=>(candidates||[]).filter(c=>promoted[c.id]==="interview"||promoted[c.id]==="president"),[candidates,promoted]);
  const presidentCandidates=useMemo(()=>(candidates||[]).filter(c=>promoted[c.id]==="president"),[candidates,promoted]);
  const chosenCandidates=useMemo(()=>(candidates||[]).filter(c=>chosenSet[c.id]),[candidates,chosenSet]);
  const myAssigned=useMemo(()=>assignments.filter(a=>a.interviewer_id===user?.id),[assignments,user]);

  const navTo=(v)=>{setView(v);setSelected(null);setShowImport(false);setSidebarOpen(false);};

  const TABS=[
    {id:"list",       label:"📋 Applications"},
    {id:"results",    label:"🏆 Results"},
    {id:"interviews", label:"🎤 Interviews"},
    {id:"presidents", label:"👔 Final Round"},
    {id:"chosen",     label:"⭐ Chosen"},
    {id:"stats",      label:"📊 Stats"},
    ...(isPresident?[{id:"groups",label:"👥 Groups"}]:[]),
  ];

  if(!authed)return<PasswordGate onPass={()=>setAuthed(true)}/>;
  if(!user)return<LoginScreen members={members} onLogin={setUser}/>;
  if(appLoading)return(
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:16,fontFamily:"system-ui,sans-serif"}}>
      <div style={{display:"inline-block",background:"#fff",borderRadius:12,padding:"8px 16px"}}><img src="/nic-logo.png" alt="NIC" style={{height:52,objectFit:"contain"}}/></div><Spinner size={28}/>
      <div style={{color:C.textMid,fontSize:14}}>Loading portal…</div>
    </div>
  );

  // ─────────────────────────────────────────────────────────────────────────
  return(
    <div style={{display:"flex",minHeight:"100vh",background:C.bg,fontFamily:"system-ui,sans-serif",color:C.text}}>
      <style>{`
        @keyframes spin    { to { transform:rotate(360deg); } }
        @keyframes fadeUp  { from { opacity:0; transform:translateY(10px); } to { opacity:1; transform:translateY(0); } }
        @keyframes slideIn { from { transform:translateX(-100%); } to { transform:translateX(0); } }
        * { box-sizing:border-box; }
        ::-webkit-scrollbar { width:4px; }
        ::-webkit-scrollbar-track { background:${C.bg}; }
        ::-webkit-scrollbar-thumb { background:${C.border}; border-radius:4px; }
        button { transition:all 0.15s; cursor:pointer; font-family:system-ui,sans-serif; }
        button:hover:not(:disabled) { filter:brightness(0.93); }
        a:hover { opacity:0.75; }
        textarea, input { font-family:system-ui,sans-serif; }
        @media (max-width:768px) {
          .sidebar { transform:translateX(-100%); position:fixed!important; z-index:1000; height:100vh; transition:transform 0.25s; }
          .sidebar.open { transform:translateX(0); animation:slideIn 0.25s ease; }
          .main-content { padding:16px!important; }
          .table-row { flex-wrap:wrap; gap:6px; }
          .hide-mobile { display:none!important; }
        }
      `}</style>

      <Toast toast={toast}/>

      {/* Modals */}
      {answersModal&&<AnswersModal candidate={answersModal} aiScores={aiScores} onClose={()=>setAnswersModal(null)}/>}
      {assignModal&&<AssignModal assignModal={assignModal} candidates={candidates} assignments={assignments} members={members} onAssign={handleAssign} onClose={()=>setAssignModal(null)}/>}

      {/* Mobile overlay */}
      {sidebarOpen&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:999}} onClick={()=>setSidebarOpen(false)}/>}

      {/* ── SIDEBAR ── */}
      <aside className={`sidebar${sidebarOpen?" open":""}`} style={{width:236,background:C.navy,display:"flex",flexDirection:"column",flexShrink:0,position:"sticky",top:0,height:"100vh",overflowY:"auto"}}>
        <div style={{background:C.navyMid,padding:"18px 16px 14px",textAlign:"center"}}>
          <div style={{display:"inline-block",background:"#fff",borderRadius:10,padding:"6px 14px"}}><img src="/nic-logo.png" alt="NIC" style={{height:44,objectFit:"contain"}}/></div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:9,background:"rgba(255,255,255,0.07)",margin:"12px 12px 0",borderRadius:9,padding:"9px 11px"}}>
          <div style={{width:32,height:32,borderRadius:"50%",background:"rgba(255,255,255,0.15)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:800,color:"#fff",flexShrink:0}}>{initials(user.name)}</div>
          <div style={{minWidth:0}}>
            <div style={{fontSize:12,fontWeight:700,color:"#fff",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{user.name}</div>
            {isPresident&&<div style={{fontSize:9,color:"rgba(255,255,255,0.5)",letterSpacing:1}}>CO-PRESIDENT</div>}
          </div>
        </div>
        <nav style={{display:"flex",flexDirection:"column",gap:2,margin:"14px 8px 0"}}>
          {TABS.map(({id,label})=>(
            <button key={id} onClick={()=>navTo(id)}
              style={{display:"flex",alignItems:"center",gap:7,padding:"9px 12px",background:view===id?"rgba(255,255,255,0.14)":"transparent",border:"none",borderRadius:8,color:view===id?"#fff":"rgba(255,255,255,0.5)",fontSize:13,fontWeight:600,textAlign:"left",position:"relative"}}>
              {label}
              {id==="interviews"&&myAssigned.length>0&&<span style={{marginLeft:"auto",background:"#ef4444",color:"#fff",borderRadius:20,padding:"1px 7px",fontSize:10,fontWeight:800}}>{myAssigned.length}</span>}
              {id==="chosen"&&chosenCandidates.length>0&&<span style={{marginLeft:"auto",background:"#f59e0b",color:"#fff",borderRadius:20,padding:"1px 7px",fontSize:10,fontWeight:800}}>{chosenCandidates.length}</span>}
            </button>
          ))}
        </nav>

        <div style={{background:"rgba(255,255,255,0.06)",borderRadius:9,padding:12,margin:"14px 12px 0"}}>
          <div style={{fontSize:9,color:"rgba(255,255,255,0.4)",letterSpacing:2,marginBottom:7}}>MY PROGRESS</div>
          <div style={{fontSize:28,fontWeight:800,color:"#fff",lineHeight:1}}>{progress.done}<span style={{fontSize:13,color:"rgba(255,255,255,0.35)",fontWeight:400}}>/{progress.total}</span></div>
          <div style={{height:3,background:"rgba(255,255,255,0.1)",borderRadius:2,marginTop:8}}><div style={{height:"100%",background:"#fff",borderRadius:2,width:progress.total?`${progress.done/progress.total*100}%`:"0%",transition:"width 0.6s"}}/></div>
          <div style={{fontSize:10,color:"rgba(255,255,255,0.3)",marginTop:4}}>evaluated</div>
        </div>

        {/* Deadline countdown — visible to all */}
        {deadline&&(()=>{
          const dl=new Date(deadline+"T23:59:59");
          const now=new Date();
          const diff=dl-now;
          const days=Math.floor(diff/86400000);
          const hours=Math.floor((diff%86400000)/3600000);
          const expired=diff<0;
          return(
            <div style={{background:expired?"rgba(239,68,68,0.15)":"rgba(245,158,11,0.15)",borderRadius:9,padding:12,margin:"10px 12px 0",border:`1px solid ${expired?"rgba(239,68,68,0.3)":"rgba(245,158,11,0.3)"}`}}>
              <div style={{fontSize:9,color:expired?"#fca5a5":"#fcd34d",letterSpacing:2,marginBottom:6}}>EVALUATION DEADLINE</div>
              {expired
                ?<div style={{fontSize:13,color:"#fca5a5",fontWeight:700}}>⚠️ Deadline passed</div>
                :<div style={{fontSize:13,color:"#fcd34d",fontWeight:700}}>{days}d {hours}h remaining</div>
              }
              <div style={{fontSize:10,color:"rgba(255,255,255,0.3)",marginTop:3}}>{new Date(deadline).toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"})}</div>
            </div>
          );
        })()}
        {isPresident&&(
          <div style={{background:"rgba(255,255,255,0.06)",borderRadius:9,padding:12,margin:"10px 12px 0"}}>
            <div style={{fontSize:9,color:"rgba(255,255,255,0.4)",letterSpacing:2,marginBottom:8}}>TEAM · LIVE</div>
            <div style={{maxHeight:140,overflowY:"auto"}}>
              {members.map(m=>{const p=memberProgress(m.id,candidates,allScores);const done=p.done===p.total&&p.total>0;return(
                <div key={m.id} style={{display:"flex",justifyContent:"space-between",marginBottom:4,fontSize:11,color:done?"#4ade80":"rgba(255,255,255,0.35)"}}>
                  <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:120}}>{m.name.split(" ")[0]}</span>
                  <span style={{fontFamily:"monospace",flexShrink:0}}>{p.done}/{p.total}</span>
                </div>
              );})}
            </div>
            <div style={{marginTop:10,display:"flex",flexDirection:"column",gap:6}}>
              {!revealed
                ?(<><div style={{fontSize:10,color:"rgba(255,255,255,0.35)"}}>{!allDone?"⏳ Not all done yet":""}</div><button onClick={handleReveal} style={{background:"#fff",color:C.navy,border:"none",borderRadius:7,padding:"8px 10px",width:"100%",fontSize:11,fontWeight:700,cursor:"pointer"}}>🔓 Reveal Results</button></>)
                :(<><div style={{color:"#4ade80",fontSize:11,fontWeight:700}}>✓ Results Revealed</div><button onClick={handleUnreveal} style={{background:"rgba(239,68,68,0.2)",color:"#fca5a5",border:"1px solid rgba(239,68,68,0.3)",borderRadius:7,padding:"7px 10px",width:"100%",fontSize:11,fontWeight:600,cursor:"pointer"}}>🔒 Hide</button></>)
              }
              <button onClick={()=>setShowImport(v=>!v)} style={{background:"rgba(255,255,255,0.10)",color:"rgba(255,255,255,0.7)",border:"1px solid rgba(255,255,255,0.15)",borderRadius:7,padding:"7px 10px",width:"100%",fontSize:11,fontWeight:600}}>
                {showImport?"✕ Close":"📂 Import CSV"}
              </button>
              <div style={{display:"flex",alignItems:"center",gap:5,marginTop:2}}>
                <span style={{fontSize:9,color:"rgba(255,255,255,0.4)",letterSpacing:1,whiteSpace:"nowrap"}}>⏰ DEADLINE</span>
                <input type="date" value={deadline} onChange={e=>handleSetDeadline(e.target.value)}
                  style={{flex:1,border:"1px solid rgba(255,255,255,0.2)",borderRadius:6,padding:"4px 7px",fontSize:11,color:"#fff",background:"rgba(255,255,255,0.08)",cursor:"pointer",outline:"none"}}/>
                {deadline&&<button onClick={()=>handleSetDeadline("")} style={{background:"none",border:"none",color:"rgba(255,255,255,0.3)",fontSize:14,cursor:"pointer",padding:"0 2px"}}>✕</button>}
              </div>
            </div>
          </div>
        )}

        <button onClick={()=>{setUser(null);setView("list");setSelected(null);setAuthed(false);sessionStorage.removeItem("nic_auth");}} style={{background:"none",border:"none",color:"rgba(255,255,255,0.2)",cursor:"pointer",fontSize:11,marginTop:"auto",padding:"12px 16px",textAlign:"left"}}>← Logout</button>
      </aside>

      {/* ── MAIN ── */}
      <main className="main-content" style={{flex:1,padding:"32px 36px",overflowY:"auto",minWidth:0}}>
        {/* Mobile top bar */}
        <div style={{display:"none"}} className="hide-desktop">
          <button onClick={()=>setSidebarOpen(true)} style={{background:C.navy,color:"#fff",border:"none",borderRadius:8,padding:"8px 14px",fontSize:13,fontWeight:700,marginBottom:16}}>☰ Menu</button>
        </div>

        {isPresident&&showImport&&<ImportPanel onImportReplace={handleImportReplace} onImportAppend={handleImportAppend} loading={importing} hasCandidates={!!(candidates?.length)}/>}

        {!candidates&&!showImport&&(
          <div style={{textAlign:"center",padding:"70px 20px"}}>
            <div style={{fontSize:44,marginBottom:14}}>📋</div>
            <h3 style={{color:C.navyMid,margin:"0 0 8px",fontSize:18}}>No applications yet</h3>
            <p style={{color:C.textMid,fontSize:14}}>{isPresident?"Use the Import CSV button in the sidebar.":"Waiting for Co-Presidents to import applications."}</p>
          </div>
        )}

        {/* ══ APPLICATIONS ══ */}
        {candidates&&view==="list"&&!selected&&(
          <div style={{animation:"fadeUp 0.2s ease"}}>
            {/* Group info banner */}
            {!isPresident&&myGroupId&&!showAllMode&&(()=>{
              const myGroup=evalGroups.find(g=>g.id===myGroupId);
              return myGroup?(
                <div style={{background:"#eff6ff",border:"1px solid #bfdbfe",borderRadius:9,padding:"10px 16px",marginBottom:14,fontSize:13,color:C.navy,fontWeight:600}}>
                  👥 You are in <strong>{myGroup.name}</strong> — showing {myCandidates.length} of {candidates.length} candidates
                </div>
              ):null;
            })()}

            <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:20,gap:12,flexWrap:"wrap"}}>
              <div>
                <h1 style={{fontSize:24,fontWeight:800,color:C.navy,margin:0}}>Applications</h1>
                <p style={{color:C.textMid,fontSize:13,marginTop:3}}>{myCandidates.length} candidates · {progress.done} evaluated by you{myCandidates.length<(candidates||[]).length?` (filtered by group)`:""}</p>
              </div>
              <input type="text" placeholder="Search name, student no. or email…" value={search} onChange={e=>setSearch(e.target.value)}
                style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 13px",fontSize:13,color:C.text,width:260,outline:"none",flexShrink:0}}/>
            </div>
            <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:11,overflow:"hidden",boxShadow:"0 1px 4px rgba(0,0,0,0.04)"}}>
              <div style={{display:"flex",padding:"9px 18px",background:C.bg,borderBottom:`1px solid ${C.border}`,fontSize:10,color:C.textLt,letterSpacing:1.5,fontWeight:700}}>
                <span style={{flex:0.4}}>#</span><span style={{flex:1.8}}>STUDENT</span><span style={{flex:2}} className="hide-mobile">EMAIL</span><span style={{flex:0.9}}>AI</span><span style={{flex:1}}>SCORES</span><span style={{flex:0.5}} className="hide-mobile">CV</span><span style={{flex:0.8,textAlign:"right"}}></span>
              </div>
              {filtered.length===0&&<div style={{padding:"28px",textAlign:"center",color:C.textLt,fontSize:14}}>No candidates match your search.</div>}
              {filtered.map((c,i)=>{
                const cs=myScores.filter(s=>s.candidate_id===c.id);
                const done=QUESTIONS.every(q=>cs.some(s=>s.question_id===q.id));
                const ai=aiScores[c.id];
                return(
                  <div key={c.id} className="table-row" style={{display:"flex",alignItems:"center",padding:"11px 18px",borderBottom:`1px solid #f1f5f9`,background:done?"#f0fdf4":"#fff"}}>
                    <span style={{flex:0.4,color:C.textLt,fontSize:11,fontFamily:"monospace"}}>{String(i+1).padStart(2,"0")}</span>
                    <span style={{flex:1.8,fontWeight:700,color:C.navy,fontSize:13}}>{displayName(c)}</span>
                    <span className="hide-mobile" style={{flex:2,color:C.textMid,fontSize:12}}>{c.email||"—"}</span>
                    <span style={{flex:0.9}}><AiBadge pct={ai?.overall_pct??null}/></span>
                    <span style={{flex:1,display:"flex",gap:2}}>
                      {QUESTIONS.map(q=>{const sc=cs.find(x=>x.question_id===q.id);return<span key={q.id} title={q.label} style={{width:24,height:24,borderRadius:5,background:sc!=null?C.navy:"#f1f5f9",border:`1px solid ${sc!=null?C.navy:C.border}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,color:sc!=null?"#fff":C.textLt,fontWeight:800}}>{sc!=null?fmtScore(sc.score):"·"}</span>;})}
                    </span>
                    <span className="hide-mobile" style={{flex:0.5}}>{c.cv_link?<a href={c.cv_link} target="_blank" rel="noopener noreferrer" style={{color:C.accent,fontSize:12,textDecoration:"none",fontWeight:600}}>CV ↗</a>:<span style={{color:C.border,fontSize:12}}>—</span>}</span>
                    <span style={{flex:0.8,textAlign:"right"}}><button onClick={()=>{setSelected(c);setView("evaluate");if(!aiScores[c.id])runDetection(c);}} style={{background:done?"#dcfce7":C.bg,color:done?C.green:C.navy,border:`1px solid ${done?"#86efac":C.border}`,borderRadius:6,padding:"5px 11px",fontSize:11,fontWeight:700}}>{done?"✓ Edit":"Evaluate"}</button></span>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ══ EVALUATE ══ */}
        {candidates&&view==="evaluate"&&selected&&(
          <div style={{animation:"fadeUp 0.2s ease"}}>
            <button onClick={()=>{setSelected(null);setView("list");}} style={{background:"none",border:"none",color:C.textMid,cursor:"pointer",fontSize:13,marginBottom:14,padding:0}}>← Back</button>
            {/* Candidate info card */}
            <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:18,boxShadow:"0 2px 8px rgba(15,41,82,0.06)"}}>
              <div style={{display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
                {/* Avatar */}
                <div style={{width:56,height:56,borderRadius:"50%",background:C.navy,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,fontWeight:900,color:"#fff",flexShrink:0}}>
                  {initials(displayName(selected))}
                </div>
                {/* Info */}
                <div style={{flex:1,minWidth:0}}>
                  <h1 style={{fontSize:22,fontWeight:800,color:C.navy,margin:0}}>{displayName(selected)}</h1>
                  <div style={{display:"flex",gap:16,marginTop:8,flexWrap:"wrap"}}>
                    {selected.student_number&&(
                      <div style={{display:"flex",alignItems:"center",gap:5}}>
                        <span style={{fontSize:11,color:C.textLt,fontWeight:700,letterSpacing:1}}>ID</span>
                        <span style={{fontSize:14,color:C.text,fontWeight:600}}>#{selected.student_number}</span>
                      </div>
                    )}
                    {selected.email&&(
                      <div style={{display:"flex",alignItems:"center",gap:5}}>
                        <span style={{fontSize:11,color:C.textLt,fontWeight:700}}>📧</span>
                        <span style={{fontSize:14,color:C.text,fontWeight:600}}>{selected.email}</span>
                      </div>
                    )}
                    {selected.phone&&(
                      <div style={{display:"flex",alignItems:"center",gap:5}}>
                        <span style={{fontSize:11,color:C.textLt,fontWeight:700}}>📱</span>
                        <span style={{fontSize:14,color:C.navy,fontWeight:700}}>{selected.phone}</span>
                      </div>
                    )}
                  </div>
                </div>
                {/* CV button */}
                {selected.cv_link&&(
                  <a href={selected.cv_link} target="_blank" rel="noopener noreferrer"
                    style={{background:C.navy,color:"#fff",borderRadius:9,padding:"10px 18px",fontSize:13,fontWeight:700,textDecoration:"none",whiteSpace:"nowrap",flexShrink:0}}>
                    📄 Open CV →
                  </a>
                )}
              </div>
            </div>
            {/* AI panel */}
            <div style={{background:"#f8fafc",border:`1px solid ${C.border}`,borderRadius:11,padding:18,marginBottom:14}}>
              <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
                <span style={{fontSize:10,letterSpacing:2,color:C.textMid,fontWeight:700}}>AI DETECTION</span>
                {detecting&&<Spinner size={12}/>}
              </div>
              {candidateAI.overall_pct!=null?(
                <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                  {QUESTIONS.map(q=>(
                    <div key={q.id} style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 13px",minWidth:85}}>
                      <div style={{fontSize:9,color:C.textLt,marginBottom:3,fontWeight:700}}>{q.label.toUpperCase()}</div>
                      <div style={{fontSize:20,fontWeight:900,color:aiColor(candidateAI[`${q.id}_pct`])}}>{candidateAI[`${q.id}_pct`]??"-"}%</div>
                    </div>
                  ))}
                  <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 13px",minWidth:85}}>
                    <div style={{fontSize:9,color:C.textLt,marginBottom:3,fontWeight:700}}>OVERALL</div>
                    <div style={{fontSize:20,fontWeight:900,color:aiColor(candidateAI.overall_pct)}}>{candidateAI.overall_pct??"-"}%</div>
                  </div>
                  <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 13px",flex:2,minWidth:160}}>
                    <div style={{fontSize:9,color:C.textLt,marginBottom:3,fontWeight:700}}>FLAGS</div>
                    <div style={{fontSize:12,color:C.textMid,lineHeight:1.5}}>{candidateAI.flags||"None"}</div>
                  </div>
                </div>
              ):<span style={{color:C.textLt,fontSize:12}}>{detecting?"Analysing…":"Run detection by opening a candidate."}</span>}
            </div>
            {QUESTIONS.map(q=>{
              const existing=myScores.find(s=>s.candidate_id===selected.id&&s.question_id===q.id);
              const cur=existing?.score;
              return(
                <div key={q.id} style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:11,padding:20,marginBottom:10,boxShadow:"0 1px 3px rgba(0,0,0,0.03)"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6,flexWrap:"wrap",gap:8}}>
                    <div><div style={{fontSize:10,color:C.navy,letterSpacing:2,fontWeight:700}}>{q.label.toUpperCase()}</div></div>
                    {cur!=null&&<span style={{background:C.navy,color:"#fff",borderRadius:20,padding:"3px 12px",fontSize:12,fontWeight:700}}>{fmtScore(cur)}/4</span>}
                  </div>
                  <div style={{fontSize:12,color:C.textLt,lineHeight:1.5,marginBottom:14,fontStyle:"italic"}}>"{q.fullText}"</div>
                  <div style={{fontSize:14,color:C.text,lineHeight:1.85,marginBottom:16,padding:"13px 15px",background:C.bg,borderRadius:8,borderLeft:`3px solid ${C.border}`,whiteSpace:"pre-wrap",wordBreak:"break-word"}}>
                    {selected[q.id]?.trim()||<em style={{color:C.textLt}}>No answer provided</em>}
                  </div>
                  <ScoreInput value={cur} onChange={v=>handleScore(selected.id,q.id,v)}/>
                </div>
              );
            })}
            <button onClick={()=>{setSelected(null);setView("list");}} style={{background:C.navy,color:"#fff",border:"none",borderRadius:8,padding:"11px 24px",fontSize:14,fontWeight:700,marginTop:6}}>Save & Return →</button>
          </div>
        )}

        {/* ══ RESULTS ══ */}
        {candidates&&view==="results"&&(
          <div style={{animation:"fadeUp 0.2s ease"}}>
            <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:20,gap:12,flexWrap:"wrap"}}>
              <div>
                <h1 style={{fontSize:24,fontWeight:800,color:C.navy,margin:0}}>Results</h1>
                <p style={{color:C.textMid,fontSize:13,marginTop:3}}>{revealed?"Visible to all members":isPresident?"Preview mode":"Locked"}</p>
              </div>
              {isPresident&&(
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{fontSize:13,color:C.textMid}}>Top</span>
                  <input type="number" min={1} max={candidates.length} value={topN} onChange={e=>handleTopN(Math.max(1,parseInt(e.target.value)||1))} style={{width:56,background:"#fff",border:`1px solid ${C.border}`,borderRadius:7,padding:"6px 8px",color:C.text,fontSize:14,textAlign:"center"}}/>
                </div>
              )}
            </div>
            {!revealed&&!isPresident?(
              <div style={{textAlign:"center",padding:"70px 20px"}}>
                <div style={{fontSize:48}}>🔒</div>
                <h3 style={{color:C.textMid,margin:"14px 0 8px",fontSize:18}}>Results are locked</h3>
                <p style={{color:C.textLt,fontSize:14}}>Co-Presidents will reveal results when ready.</p>
              </div>
            ):(
              <>
                {!revealed&&<div style={{background:"#fffbeb",border:"1px solid #fcd34d",borderRadius:8,padding:"9px 14px",fontSize:13,color:"#92400e",marginBottom:16,fontWeight:600}}>⚠️ Preview — members see locked screen</div>}
                <div style={{display:"flex",flexDirection:"column",gap:7}}>
                  {ranked.slice(0,isPresident?topN:ranked.length).map((c,i)=>{
                    const ai=aiScores[c.id];
                    const isTop=i<topN;
                    const pStatus=promoted[c.id];
                    return(
                      <div key={c.id} style={{display:"flex",alignItems:"center",gap:12,background:"#fff",border:`1px solid ${isTop?C.navyMid+"44":C.border}`,borderLeft:`4px solid ${isTop?C.navy:C.border}`,borderRadius:9,padding:"13px 18px",boxShadow:isTop?"0 2px 6px rgba(15,41,82,0.06)":"none",flexWrap:"wrap",gap:10}}>
                        <div style={{fontSize:14,color:isTop?C.navy:C.textLt,fontWeight:900,fontFamily:"monospace",width:30,flexShrink:0}}>#{i+1}</div>
                        <div style={{flex:"1 1 160px",minWidth:0,cursor:"pointer"}} onClick={()=>setAnswersModal(c)}>
                          <div style={{fontWeight:700,fontSize:14,color:C.navy,textDecoration:"underline",textDecorationStyle:"dotted",textUnderlineOffset:2}}>{displayName(c)}</div>
                          <div style={{fontSize:11,color:C.textLt,marginTop:1}}>{c.email} <span style={{color:C.accent}}>· view answers</span></div>
                        </div>
                        <AiBadge pct={ai?.overall_pct??null}/>
                        {c.cv_link&&<a href={c.cv_link} target="_blank" rel="noopener noreferrer" style={{color:C.accent,fontSize:12,textDecoration:"none",fontWeight:600}} className="hide-mobile">CV ↗</a>}
                        {isPresident&&(
                          <div style={{display:"flex",gap:5,flexShrink:0,alignItems:"center",flexWrap:"wrap"}}>
                            {pStatus==="president"?(<><span style={{background:"#ede9fe",color:"#7c3aed",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>👔 Final</span><button onClick={()=>handleDemote(c.id)} style={{background:"#fee2e2",color:C.red,border:"1px solid #fca5a5",borderRadius:6,padding:"3px 8px",fontSize:11,fontWeight:700}}>✕</button></>)
                            :pStatus==="interview"?(<><span style={{background:"#dcfce7",color:C.green,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>🎤 Interview</span><button onClick={()=>handleDemote(c.id)} style={{background:"#fee2e2",color:C.red,border:"1px solid #fca5a5",borderRadius:6,padding:"3px 8px",fontSize:11,fontWeight:700}}>✕</button></>)
                            :(<button onClick={()=>handlePromote(c.id,"interview")} style={{background:"#f0fdf4",color:C.green,border:`1px solid #86efac`,borderRadius:7,padding:"5px 10px",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>+ Interview</button>)}
                          </div>
                        )}
                        <div style={{textAlign:"right",minWidth:76,flexShrink:0}}>
                          <div style={{fontSize:24,fontWeight:900,color:scoreColor(parseFloat(c.avg))}}>{c.avg??"-"}{c.avg?"%":""}</div>
                          <div style={{fontSize:9,color:C.textLt,letterSpacing:1}}>AVG</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </>
            )}
          </div>
        )}

        {/* ══ INTERVIEWS ══ */}
        {candidates&&view==="interviews"&&(
          <div style={{animation:"fadeUp 0.2s ease"}}>
            <h1 style={{fontSize:24,fontWeight:800,color:C.navy,margin:"0 0 4px"}}>Member Interviews</h1>
            <p style={{color:C.textMid,fontSize:13,marginBottom:20}}>{interviewCandidates.length} candidates in this round</p>
            {interviewCandidates.length===0?(
              <div style={{textAlign:"center",padding:"60px 20px",background:"#fff",borderRadius:12,border:`1px solid ${C.border}`}}>
                <div style={{fontSize:36,marginBottom:10}}>🎤</div>
                <h3 style={{color:C.navyMid,margin:"0 0 8px",fontSize:16}}>No candidates promoted yet</h3>
                <p style={{color:C.textMid,fontSize:13}}>{isPresident?"Go to Results → + Interview":"Waiting for Co-Presidents to promote candidates."}</p>
              </div>
            ):(
              interviewCandidates.map(c=>{
                const cAssign=assignments.filter(a=>a.candidate_id===c.id&&a.round==="interview");
                const cFeedback=interviewData.filter(i=>i.candidate_id===c.id&&i.round==="interview");
                const assignedNames=cAssign.map(a=>members.find(m=>m.id===a.interviewer_id)?.name).filter(Boolean);
                const amAssigned=cAssign.some(a=>a.interviewer_id===user.id);
                const allSubmitted=cAssign.length===2&&cAssign.every(a=>cFeedback.some(f=>f.interviewer_id===a.interviewer_id));
                return(
                  <div key={c.id} style={{marginBottom:20}}>
                    <div style={{display:"flex",alignItems:"center",padding:"11px 16px",background:C.navy,borderRadius:"11px 11px 0 0",flexWrap:"wrap",gap:8}}>
                      <div style={{flex:1,minWidth:0}}>
                        <span style={{fontWeight:700,color:"#fff",fontSize:14}}>{displayName(c)}</span>
                        <span style={{color:"rgba(255,255,255,0.4)",fontSize:11,marginLeft:8}}>{c.email}</span>
                      </div>
                      <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
                        {assignedNames.map((n,i)=><span key={i} style={{background:"rgba(255,255,255,0.15)",color:"#fff",borderRadius:20,padding:"2px 9px",fontSize:11,fontWeight:600}}>{n.split(" ")[0]}</span>)}
                        {allSubmitted&&<span style={{background:"#dcfce7",color:C.green,borderRadius:20,padding:"2px 9px",fontSize:11,fontWeight:700}}>✓ Done</span>}
                      </div>
                      {isPresident&&(
                        <div style={{display:"flex",gap:5}}>
                          <button onClick={()=>setAssignModal({id:c.id,round:"interview"})} style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"1px solid rgba(255,255,255,0.2)",borderRadius:6,padding:"5px 10px",fontSize:11,fontWeight:600}}>{cAssign.length===2?"✏️ Reassign":"👥 Assign"}</button>
                          {promoted[c.id]!=="president"?(
                            <button onClick={()=>handlePromote(c.id,"president")} style={{background:"#7c3aed",color:"#fff",border:"none",borderRadius:6,padding:"5px 10px",fontSize:11,fontWeight:700}}>+ Final Round</button>
                          ):(
                            <><span style={{background:"#ede9fe",color:"#7c3aed",borderRadius:20,padding:"2px 9px",fontSize:11,fontWeight:700}}>👔 Final</span><button onClick={()=>handlePromote(c.id,"interview")} style={{background:"rgba(239,68,68,0.2)",color:"#fca5a5",border:"1px solid rgba(239,68,68,0.3)",borderRadius:6,padding:"5px 9px",fontSize:11,fontWeight:700}}>✕</button></>
                          )}
                        </div>
                      )}
                    </div>
                    {isPresident
                      ?<FeedbackView candidate={c} interviewData={interviewData} members={members} round="interview" onViewAnswers={setAnswersModal}/>
                      :amAssigned
                        ?<InterviewForm candidate={c} interviewData={interviewData} user={user} onSave={handleInterviewSave} round="interview" onViewAnswers={setAnswersModal}/>
                        :<div style={{padding:16,background:"#fff",border:`1px solid ${C.border}`,borderTop:"none",borderRadius:"0 0 11px 11px",color:C.textLt,fontSize:13,textAlign:"center"}}>You are not assigned to interview this candidate.</div>
                    }
                  </div>
                );
              })
            )}
          </div>
        )}

        {/* ══ FINAL ROUND ══ */}
        {candidates&&view==="presidents"&&(
          <div style={{animation:"fadeUp 0.2s ease"}}>
            <h1 style={{fontSize:24,fontWeight:800,color:C.navy,margin:"0 0 4px"}}>Final Round</h1>
            <p style={{color:C.textMid,fontSize:13,marginBottom:20}}>{presidentCandidates.length} candidates · Co-President interviews</p>
            {presidentCandidates.length===0?(
              <div style={{textAlign:"center",padding:"60px 20px",background:"#fff",borderRadius:12,border:`1px solid ${C.border}`}}>
                <div style={{fontSize:36,marginBottom:10}}>👔</div>
                <h3 style={{color:C.navyMid,margin:"0 0 8px",fontSize:16}}>No candidates in Final Round</h3>
                <p style={{color:C.textMid,fontSize:13}}>Go to Interviews → + Final Round after member interviews.</p>
              </div>
            ):(
              presidentCandidates.map(c=>{
                const cAssign=assignments.filter(a=>a.candidate_id===c.id&&a.round==="president");
                const cFeedback=interviewData.filter(i=>i.candidate_id===c.id&&i.round==="president");
                const assignedNames=cAssign.map(a=>members.find(m=>m.id===a.interviewer_id)?.name).filter(Boolean);
                const amAssigned=cAssign.some(a=>a.interviewer_id===user.id);
                const allDoneHere=cAssign.length>0&&cAssign.every(a=>cFeedback.some(f=>f.interviewer_id===a.interviewer_id));
                const isChosen=!!chosenSet[c.id];
                return(
                  <div key={c.id} style={{marginBottom:20}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,padding:"11px 16px",background:"#3d1c7a",borderRadius:"11px 11px 0 0",flexWrap:"wrap"}}>
                      <div style={{flex:1,minWidth:0}}>
                        <span style={{fontWeight:700,color:"#fff",fontSize:14}}>{displayName(c)}</span>
                        <span style={{color:"rgba(255,255,255,0.4)",fontSize:11,marginLeft:8}}>{c.email}</span>
                      </div>
                      <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
                        {assignedNames.map((n,i)=><span key={i} style={{background:"rgba(255,255,255,0.15)",color:"#fff",borderRadius:20,padding:"2px 9px",fontSize:11,fontWeight:600}}>{n.split(" ")[0]}</span>)}
                        {allDoneHere&&<span style={{background:"#dcfce7",color:C.green,borderRadius:20,padding:"2px 9px",fontSize:11,fontWeight:700}}>✓ Done</span>}
                        {isChosen&&<span style={{background:"#fef9c3",color:"#854d0e",borderRadius:20,padding:"2px 9px",fontSize:11,fontWeight:700}}>⭐ Chosen</span>}
                      </div>
                      {c.cv_link&&<a href={c.cv_link} target="_blank" rel="noopener noreferrer" style={{background:"rgba(255,255,255,0.15)",color:"#fff",borderRadius:6,padding:"5px 10px",fontSize:11,fontWeight:700,textDecoration:"none"}}>CV ↗</a>}
                      <button onClick={()=>setAnswersModal(c)} style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"none",borderRadius:6,padding:"5px 10px",fontSize:11,fontWeight:600}}>📋 Answers</button>
                      {isPresident&&(
                        <>
                          <button onClick={()=>setAssignModal({id:c.id,round:"president"})} style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"1px solid rgba(255,255,255,0.2)",borderRadius:6,padding:"5px 10px",fontSize:11,fontWeight:600}}>{cAssign.length>0?"✏️ Reassign":"👥 Assign"}</button>
                          <button onClick={()=>handlePromote(c.id,"interview")} style={{background:"rgba(239,68,68,0.2)",color:"#fca5a5",border:"1px solid rgba(239,68,68,0.3)",borderRadius:6,padding:"5px 9px",fontSize:11,fontWeight:700}}>✕ Cancel</button>
                          <button onClick={()=>handleToggleChosen(c.id)} style={{background:isChosen?"#fef9c3":"rgba(255,255,255,0.1)",color:isChosen?"#854d0e":"rgba(255,255,255,0.7)",border:`1px solid ${isChosen?"#fcd34d":"rgba(255,255,255,0.15)"}`,borderRadius:6,padding:"5px 10px",fontSize:11,fontWeight:700}}>
                            {isChosen?"★ Unchose":"☆ Choose"}
                          </button>
                        </>
                      )}
                    </div>
                    {isPresident&&!amAssigned?(
                      <div style={{background:"#fff",border:`1px solid ${C.border}`,borderTop:"none",borderRadius:"0 0 11px 11px",overflow:"hidden"}}>
                        <div style={{padding:"11px 16px",background:"#faf8ff",borderBottom:`1px solid ${C.border}`,fontSize:13,color:"#7c3aed",fontWeight:600}}>Not assigned — assign yourself to evaluate.</div>
                        {cFeedback.length>0&&<FeedbackView candidate={c} interviewData={interviewData} members={members} round="president" onViewAnswers={setAnswersModal}/>}
                      </div>
                    ):isPresident&&amAssigned?(
                      <div>
                        <InterviewForm candidate={c} interviewData={interviewData} user={user} onSave={handleInterviewSave} round="president" onViewAnswers={setAnswersModal}/>
                        {cFeedback.filter(f=>f.interviewer_id!==user.id).length>0&&(
                          <div>
                            <div style={{padding:"9px 16px",background:"#f5f3ff",fontSize:10,color:"#7c3aed",letterSpacing:1,fontWeight:700}}>OTHER PRESIDENT'S FEEDBACK</div>
                            <FeedbackView candidate={c} interviewData={interviewData.filter(f=>f.interviewer_id!==user.id)} members={members} round="president" onViewAnswers={setAnswersModal}/>
                          </div>
                        )}
                      </div>
                    ):(
                      <div style={{padding:16,background:"#fff",border:`1px solid ${C.border}`,borderTop:"none",borderRadius:"0 0 11px 11px",color:C.textLt,fontSize:13,textAlign:"center"}}>Final Round interviews are conducted by Co-Presidents only.</div>
                    )}
                  </div>
                );
              })
            )}
          </div>
        )}

        {/* ══ CHOSEN ══ */}
        {candidates&&view==="chosen"&&(
          <div style={{animation:"fadeUp 0.2s ease"}}>
            <h1 style={{fontSize:24,fontWeight:800,color:C.navy,margin:"0 0 4px"}}>⭐ Chosen Candidates</h1>
            <p style={{color:C.textMid,fontSize:13,marginBottom:20}}>{chosenCandidates.length} candidate{chosenCandidates.length!==1?"s":""} selected for NIC-UD</p>
            {chosenCandidates.length===0?(
              <div style={{textAlign:"center",padding:"60px 20px",background:"#fff",borderRadius:12,border:`1px solid ${C.border}`}}>
                <div style={{fontSize:36,marginBottom:10}}>⭐</div>
                <h3 style={{color:C.navyMid,margin:"0 0 8px",fontSize:16}}>No candidates chosen yet</h3>
                <p style={{color:C.textMid,fontSize:13}}>Go to Final Round and click ☆ Choose on candidates that are accepted.</p>
              </div>
            ):(
              <div style={{display:"flex",flexDirection:"column",gap:10}}>
                {chosenCandidates.map((c,i)=>{
                  const presidentFeedback=interviewData.filter(f=>f.candidate_id===c.id&&f.round==="president");
                  const memberFeedback=interviewData.filter(f=>f.candidate_id===c.id&&f.round==="interview");
                  const avgFinal=presidentFeedback.length?((presidentFeedback.reduce((a,f)=>{const s=[f.personal_score,f.technical_score,f.brainstormer_score].filter(x=>x!=null);return a+s.reduce((a,x)=>a+parseFloat(x),0)/(s.length||1);},0)/presidentFeedback.length/4)*100).toFixed(0):null;
                  return(
                    <div key={c.id} style={{background:"#fff",border:`2px solid #fcd34d`,borderRadius:12,overflow:"hidden",boxShadow:"0 2px 8px rgba(245,158,11,0.10)"}}>
                      {/* Header */}
                      <div style={{display:"flex",alignItems:"center",padding:"14px 20px",background:"linear-gradient(135deg,#0f2952,#2451a0)",flexWrap:"wrap",gap:10}}>
                        <div style={{width:38,height:38,borderRadius:"50%",background:"rgba(255,255,255,0.15)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,fontWeight:900,color:"#fff",flexShrink:0}}>{initials(displayName(c))}</div>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{fontWeight:800,color:"#fff",fontSize:16}}>{displayName(c)}</div>
                          <div style={{fontSize:12,color:"rgba(255,255,255,0.5)",marginTop:1}}>#{c.student_number}{c.email?` · ${c.email}`:""}</div>
                        </div>
                        <span style={{background:"#fef9c3",color:"#854d0e",borderRadius:20,padding:"4px 12px",fontSize:12,fontWeight:700}}>⭐ Chosen</span>
                        {avgFinal&&<div style={{textAlign:"right"}}><div style={{fontSize:22,fontWeight:900,color:"#fcd34d"}}>{avgFinal}%</div><div style={{fontSize:9,color:"rgba(255,255,255,0.4)",letterSpacing:1}}>FINAL AVG</div></div>}
                        <div style={{display:"flex",gap:7}}>
                          {c.cv_link&&<a href={c.cv_link} target="_blank" rel="noopener noreferrer" style={{background:"rgba(255,255,255,0.15)",color:"#fff",borderRadius:7,padding:"6px 11px",fontSize:12,fontWeight:700,textDecoration:"none"}}>📄 CV</a>}
                          <button onClick={()=>setAnswersModal(c)} style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"none",borderRadius:7,padding:"6px 11px",fontSize:12,fontWeight:600}}>📋 Answers</button>
                          {isPresident&&<button onClick={()=>handleToggleChosen(c.id)} style={{background:"rgba(239,68,68,0.2)",color:"#fca5a5",border:"1px solid rgba(239,68,68,0.3)",borderRadius:7,padding:"6px 11px",fontSize:12,fontWeight:700}}>✕ Remove</button>}
                        </div>
                      </div>
                      {/* Summary */}
                      <div style={{padding:"14px 20px",display:"flex",gap:14,flexWrap:"wrap"}}>
                        {/* Written application score */}
                        <div style={{background:C.bg,borderRadius:9,padding:"10px 14px",minWidth:100,textAlign:"center"}}>
                          <div style={{fontSize:9,color:C.textLt,fontWeight:700,letterSpacing:1,marginBottom:4}}>APP SCORE</div>
                          <div style={{fontSize:20,fontWeight:900,color:scoreColor(parseFloat(avgScore(c.id,allScores)))}}>{avgScore(c.id,allScores)??"-"}{avgScore(c.id,allScores)?"%":""}</div>
                        </div>
                        {/* Member interview verdicts */}
                        {memberFeedback.length>0&&(
                          <div style={{background:C.bg,borderRadius:9,padding:"10px 14px"}}>
                            <div style={{fontSize:9,color:C.textLt,fontWeight:700,letterSpacing:1,marginBottom:6}}>MEMBER VERDICTS</div>
                            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                              {memberFeedback.map(f=>(
                                <div key={f.interviewer_id} style={{textAlign:"center"}}>
                                  <div style={{fontSize:10,color:C.textLt,marginBottom:2}}>{members.find(m=>m.id===f.interviewer_id)?.name.split(" ")[0]}</div>
                                  <VerdictBadge verdict={f.verdict}/>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}
                        {/* President interview verdicts */}
                        {presidentFeedback.length>0&&(
                          <div style={{background:C.bg,borderRadius:9,padding:"10px 14px"}}>
                            <div style={{fontSize:9,color:C.textLt,fontWeight:700,letterSpacing:1,marginBottom:6}}>PRESIDENT VERDICTS</div>
                            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                              {presidentFeedback.map(f=>(
                                <div key={f.interviewer_id} style={{textAlign:"center"}}>
                                  <div style={{fontSize:10,color:C.textLt,marginBottom:2}}>{members.find(m=>m.id===f.interviewer_id)?.name.split(" ")[0]}</div>
                                  <VerdictBadge verdict={f.verdict}/>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}
                        {/* President scores */}
                        {presidentFeedback.length>0&&(
                          <div style={{background:C.bg,borderRadius:9,padding:"10px 14px"}}>
                            <div style={{fontSize:9,color:C.textLt,fontWeight:700,letterSpacing:1,marginBottom:6}}>FINAL SCORES</div>
                            <div style={{display:"flex",gap:8}}>
                              {[["P",presidentFeedback.reduce((a,f)=>a+(f.personal_score!=null?1:0),0)?((presidentFeedback.reduce((a,f)=>a+parseFloat(f.personal_score||0),0)/presidentFeedback.filter(f=>f.personal_score!=null).length)).toFixed(1):null],
                                ["T",presidentFeedback.reduce((a,f)=>a+(f.technical_score!=null?1:0),0)?((presidentFeedback.reduce((a,f)=>a+parseFloat(f.technical_score||0),0)/presidentFeedback.filter(f=>f.technical_score!=null).length)).toFixed(1):null],
                                ["B",presidentFeedback.reduce((a,f)=>a+(f.brainstormer_score!=null?1:0),0)?((presidentFeedback.reduce((a,f)=>a+parseFloat(f.brainstormer_score||0),0)/presidentFeedback.filter(f=>f.brainstormer_score!=null).length)).toFixed(1):null]
                              ].map(([lbl,val])=>(
                                <div key={lbl} style={{textAlign:"center"}}>
                                  <div style={{fontSize:9,color:C.textLt,fontWeight:700}}>{lbl}</div>
                                  <div style={{fontSize:16,fontWeight:800,color:val?scoreColor((parseFloat(val)/4)*100):C.textLt}}>{val??"-"}</div>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}
        {/* ══ GROUPS ══ */}
        {candidates&&view==="groups"&&isPresident&&(
          <div style={{animation:"fadeUp 0.2s ease"}}>
            <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:20,gap:12,flexWrap:"wrap"}}>
              <div>
                <h1 style={{fontSize:24,fontWeight:800,color:C.navy,margin:0}}>👥 Evaluation Groups</h1>
                <p style={{color:C.textMid,fontSize:13,marginTop:3}}>
                  {evalGroups.length} groups · {candidates.length} candidates
                  {groupCandidates.length>0?` · ${groupCandidates.length} distributed`:" · not distributed yet"}
                </p>
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <button onClick={handleToggleShowAll}
                  style={{background:showAllMode?"#dcfce7":"#fee2e2",color:showAllMode?C.green:C.red,border:`1px solid ${showAllMode?"#86efac":"#fca5a5"}`,borderRadius:8,padding:"9px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
                  {showAllMode?"👁️ Show All: ON":"👁️ Show All: OFF"}
                </button>
              </div>
            </div>

            {/* Create group */}
            <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:16}}>
              <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
                <div style={{fontSize:11,color:C.textLt,letterSpacing:2,fontWeight:700}}>CREATE GROUP</div>
                <input id="newGroupName" type="text" placeholder="Group name (e.g. Group A)" 
                  style={{flex:"1 1 200px",border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 12px",fontSize:13,color:C.text,outline:"none"}}
                  onKeyDown={e=>{if(e.key==="Enter"&&e.target.value.trim()){handleCreateGroup(e.target.value.trim());e.target.value="";}}}
                />
                <button onClick={()=>{const inp=document.getElementById("newGroupName");if(inp?.value.trim()){handleCreateGroup(inp.value.trim());inp.value="";}}}
                  style={{background:C.navy,color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  + Create
                </button>
              </div>
            </div>

            {/* Groups list */}
            {evalGroups.length===0?(
              <div style={{textAlign:"center",padding:"50px 20px",background:"#fff",borderRadius:12,border:`1px solid ${C.border}`}}>
                <div style={{fontSize:36,marginBottom:10}}>👥</div>
                <h3 style={{color:C.navyMid,margin:"0 0 8px",fontSize:16}}>No groups created</h3>
                <p style={{color:C.textMid,fontSize:13}}>Create groups above, assign members, then distribute candidates.</p>
              </div>
            ):(
              <div style={{display:"flex",flexDirection:"column",gap:14}}>
                {evalGroups.map(group=>{
                  const gMembers=groupMembers.filter(m=>m.group_id===group.id);
                  const gCandidates=groupCandidates.filter(c=>c.group_id===group.id);
                  const memberNames=gMembers.map(gm=>members.find(m=>m.id===gm.member_id)).filter(Boolean);
                  const unassigned=members.filter(m=>!groupMembers.some(gm=>gm.member_id===m.id));

                  return(
                    <div key={group.id} style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,overflow:"hidden"}}>
                      {/* Group header */}
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"14px 18px",background:C.bg,borderBottom:`1px solid ${C.border}`,flexWrap:"wrap",gap:8}}>
                        <div>
                          <div style={{fontWeight:700,fontSize:16,color:C.navy}}>{group.name}</div>
                          <div style={{fontSize:12,color:C.textLt,marginTop:2}}>{memberNames.length} members · {gCandidates.length} candidates</div>
                        </div>
                        <button onClick={()=>handleDeleteGroup(group.id)}
                          style={{background:"#fee2e2",color:C.red,border:"1px solid #fca5a5",borderRadius:7,padding:"5px 12px",fontSize:11,fontWeight:700,cursor:"pointer"}}>
                          🗑️ Delete
                        </button>
                      </div>

                      {/* Members */}
                      <div style={{padding:18}}>
                        <div style={{fontSize:10,color:C.textLt,letterSpacing:2,fontWeight:700,marginBottom:10}}>MEMBERS IN THIS GROUP</div>
                        <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:12}}>
                          {memberNames.map(m=>(
                            <div key={m.id} style={{display:"flex",alignItems:"center",gap:6,background:C.navy,color:"#fff",borderRadius:20,padding:"5px 12px 5px 6px",fontSize:12,fontWeight:600}}>
                              <span style={{width:22,height:22,borderRadius:"50%",background:"rgba(255,255,255,0.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:800}}>{initials(m.name)}</span>
                              {m.name.split(" ")[0]}
                              <button onClick={()=>handleAssignMemberToGroup(group.id,m.id,false)}
                                style={{background:"none",border:"none",color:"rgba(255,255,255,0.5)",fontSize:14,cursor:"pointer",padding:"0 2px",marginLeft:2}}>✕</button>
                            </div>
                          ))}
                          {memberNames.length===0&&<span style={{color:C.textLt,fontSize:12,fontStyle:"italic"}}>No members assigned yet</span>}
                        </div>

                        {/* Add member dropdown */}
                        {unassigned.length>0&&(
                          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                            <span style={{fontSize:10,color:C.textLt,fontWeight:700,letterSpacing:1,alignSelf:"center"}}>ADD:</span>
                            {unassigned.map(m=>(
                              <button key={m.id} onClick={()=>handleAssignMemberToGroup(group.id,m.id,true)}
                                style={{background:C.bg,color:C.textMid,border:`1px solid ${C.border}`,borderRadius:20,padding:"4px 10px 4px 6px",fontSize:11,fontWeight:600,cursor:"pointer",display:"flex",alignItems:"center",gap:4}}>
                                <span style={{fontSize:12}}>+</span> {m.name.split(" ")[0]}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  );
                })}

                {/* Distribute button */}
                <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:20,textAlign:"center"}}>
                  <p style={{color:C.textMid,fontSize:13,marginBottom:12}}>
                    {groupCandidates.length>0
                      ?`Currently ${groupCandidates.length} candidates distributed. Redistribute will reshuffle.`
                      :`Distribute ${candidates.length} candidates equally across ${evalGroups.length} groups.`
                    }
                  </p>
                  <button onClick={handleDistributeCandidates}
                    style={{background:C.navy,color:"#fff",border:"none",borderRadius:9,padding:"11px 28px",fontSize:14,fontWeight:700,cursor:"pointer"}}>
                    🔀 {groupCandidates.length>0?"Redistribute":"Distribute"} Candidates
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ══ STATS ══ */}
        {candidates&&view==="stats"&&(
          <div style={{animation:"fadeUp 0.2s ease"}}>
            <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:20,gap:12,flexWrap:"wrap"}}>
              <div>
                <h1 style={{fontSize:24,fontWeight:800,color:C.navy,margin:0}}>📊 Statistics & Report</h1>
                <p style={{color:C.textMid,fontSize:13,marginTop:3}}>Recruitment cycle overview</p>
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>exportExcel({candidates,allScores,interviewData,chosenCandidates,members,aiScores,promoted})}
                  style={{background:"#16a34a",color:"#fff",border:"none",borderRadius:8,padding:"9px 16px",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  ⬇️ Export Excel
                </button>
                <button onClick={()=>exportPDF({candidates,allScores,interviewData,chosenCandidates,members,aiScores,promoted,ranked})}
                  style={{background:C.navy,color:"#fff",border:"none",borderRadius:8,padding:"9px 16px",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  ⬇️ Export PDF
                </button>
              </div>
            </div>

            {/* Funnel */}
            {(()=>{
              const total=candidates.length;
              const interviewed=interviewCandidates.length;
              const finalRound=presidentCandidates.length;
              const chosen=chosenCandidates.length;
              const stages=[
                {label:"Applications",n:total,color:C.navy},
                {label:"Interview Round",n:interviewed,color:C.navyLt},
                {label:"Final Round",n:finalRound,color:"#7c3aed"},
                {label:"Chosen",n:chosen,color:C.green},
              ];
              return(
                <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:22,marginBottom:16}}>
                  <div style={{fontSize:11,color:C.textLt,letterSpacing:2,fontWeight:700,marginBottom:16}}>RECRUITMENT FUNNEL</div>
                  <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                    {stages.map((s,i)=>(
                      <div key={s.label} style={{flex:"1 1 120px",textAlign:"center"}}>
                        <div style={{fontSize:36,fontWeight:900,color:s.color}}>{s.n}</div>
                        <div style={{fontSize:12,color:C.textMid,fontWeight:600,marginTop:2}}>{s.label}</div>
                        {i>0&&stages[i-1].n>0&&<div style={{fontSize:11,color:C.textLt,marginTop:2}}>{((s.n/stages[i-1].n)*100).toFixed(0)}% of prev</div>}
                        <div style={{height:4,background:s.color,borderRadius:2,marginTop:8,opacity:0.7}}/>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}

            {/* Score distribution */}
            {(()=>{
              const buckets={"0-25%":0,"25-50%":0,"50-75%":0,"75-100%":0};
              ranked.forEach(c=>{
                const a=parseFloat(c.avg);
                if(isNaN(a))return;
                if(a<25)buckets["0-25%"]++;
                else if(a<50)buckets["25-50%"]++;
                else if(a<75)buckets["50-75%"]++;
                else buckets["75-100%"]++;
              });
              const maxB=Math.max(...Object.values(buckets),1);
              return(
                <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:22,marginBottom:16}}>
                  <div style={{fontSize:11,color:C.textLt,letterSpacing:2,fontWeight:700,marginBottom:16}}>SCORE DISTRIBUTION</div>
                  <div style={{display:"flex",gap:12,alignItems:"flex-end",height:120}}>
                    {Object.entries(buckets).map(([lbl,n])=>(
                      <div key={lbl} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
                        <div style={{fontSize:12,fontWeight:700,color:C.navy}}>{n}</div>
                        <div style={{width:"100%",background:C.navyLt,borderRadius:"4px 4px 0 0",height:`${(n/maxB)*80}px`,minHeight:n>0?4:0,transition:"height 0.4s"}}/>
                        <div style={{fontSize:11,color:C.textLt,textAlign:"center"}}>{lbl}</div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}

            {/* AI detection summary */}
            {(()=>{
              const vals=Object.values(aiScores).map(a=>a.overall_pct).filter(x=>x!=null);
              if(!vals.length)return null;
              const high=vals.filter(v=>v>=70).length;
              const med=vals.filter(v=>v>=40&&v<70).length;
              const low=vals.filter(v=>v<40).length;
              return(
                <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:22,marginBottom:16}}>
                  <div style={{fontSize:11,color:C.textLt,letterSpacing:2,fontWeight:700,marginBottom:14}}>AI DETECTION SUMMARY ({vals.length} analysed)</div>
                  <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
                    {[["🔴 Likely AI",high,C.red],["🟡 Suspicious",med,C.amber],["🟢 Likely Human",low,C.green]].map(([lbl,n,col])=>(
                      <div key={lbl} style={{flex:"1 1 120px",background:col+"10",border:`1px solid ${col}33`,borderRadius:9,padding:"14px 16px",textAlign:"center"}}>
                        <div style={{fontSize:28,fontWeight:900,color:col}}>{n}</div>
                        <div style={{fontSize:12,color:col,fontWeight:600,marginTop:2}}>{lbl}</div>
                        <div style={{fontSize:11,color:C.textLt,marginTop:2}}>{vals.length?((n/vals.length)*100).toFixed(0):0}%</div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}

            {/* Member evaluation progress */}
            <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:22,marginBottom:16}}>
              <div style={{fontSize:11,color:C.textLt,letterSpacing:2,fontWeight:700,marginBottom:14}}>MEMBER EVALUATION COMPLETION</div>
              <div style={{display:"flex",flexDirection:"column",gap:8}}>
                {members.map(m=>{
                  const p=memberProgress(m.id,candidates,allScores);
                  const pct=p.total?Math.round(p.done/p.total*100):0;
                  return(
                    <div key={m.id} style={{display:"flex",alignItems:"center",gap:12}}>
                      <div style={{width:130,fontSize:13,color:C.text,fontWeight:600,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.name}</div>
                      <div style={{flex:1,height:8,background:C.bg,borderRadius:4,overflow:"hidden"}}>
                        <div style={{height:"100%",background:pct===100?C.green:C.navy,borderRadius:4,width:`${pct}%`,transition:"width 0.5s"}}/>
                      </div>
                      <div style={{width:70,fontSize:12,color:C.textMid,textAlign:"right",flexShrink:0}}>{p.done}/{p.total} ({pct}%)</div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Top member interview contributors */}
            {(()=>{
              const contrib={};
              interviewData.forEach(f=>{
                if(!contrib[f.interviewer_id])contrib[f.interviewer_id]={name:members.find(m=>m.id===f.interviewer_id)?.name||f.interviewer_id,count:0,passes:0,verdicts:[]};
                contrib[f.interviewer_id].count++;
                if(f.verdict)contrib[f.interviewer_id].verdicts.push(f.verdict);
                if(f.verdict==="pass")contrib[f.interviewer_id].passes++;
              });
              const sorted=Object.values(contrib).sort((a,b)=>b.count-a.count);
              if(!sorted.length)return null;
              return(
                <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:22,marginBottom:16}}>
                  <div style={{fontSize:11,color:C.textLt,letterSpacing:2,fontWeight:700,marginBottom:14}}>INTERVIEW CONTRIBUTIONS</div>
                  <div style={{display:"flex",flexDirection:"column",gap:8}}>
                    {sorted.map((m,i)=>(
                      <div key={m.name} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 14px",background:i===0?C.bg:"#fff",borderRadius:8,border:`1px solid ${C.border}`}}>
                        <div style={{fontSize:16,width:24,textAlign:"center"}}>{i===0?"🥇":i===1?"🥈":i===2?"🥉":"  "}</div>
                        <div style={{flex:1,fontWeight:600,fontSize:13,color:C.text}}>{m.name}</div>
                        <div style={{fontSize:12,color:C.textMid}}>{m.count} interview{m.count!==1?"s":""}</div>
                        <div style={{display:"flex",gap:5}}>
                          {m.verdicts.filter((v,i,a)=>a.indexOf(v)===i).map(v=><VerdictBadge key={v} verdict={v}/>)}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}
          </div>
        )}
      </main>
    </div>
  );
}

export default function App() {
  return React.createElement(ErrorBoundary, null, React.createElement(AppInner));
}
