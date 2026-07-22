import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import { BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, Tooltip, ResponsiveContainer, AreaChart, Area, Legend } from "recharts";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════════════════
const CATS = ["Mão de Obra","Material","Equipamento","Comissão","Outros"];
const SUBCATS = {
  "Mão de Obra":["Pedreiro","Servente","Eletricista","Encanador","Pintor","Carpinteiro","Gesseiro","Outro"],
  "Material":["Cimento","Aço/Ferro","Tijolo","Areia","Brita","Madeira","Elétrico","Hidráulico","Acabamento","Tinta","Outro"],
  "Equipamento":["Betoneira","Andaime","Escavadeira","Guindaste","Compactador","Outro"],
  "Comissão":["Vendedor","Indicação","Corretor","Engenheiro","Arquiteto","Captador","Sócio","Outro"],
  "Outros":["Transporte","Alimentação","Taxa/Imposto","Projeto","Outro"],
};
const PAGAMENTOS = ["PIX","Dinheiro","Transferência","Boleto","Cartão"];
const TIPOS_OBRA = ["Alto Padrão","Padrão","Econômico"];
const STATUS_OBRA = ["Em Andamento","Finalizada","Pausada"];
const TIPOS_EQUIPE = ["Alvenaria","Elétrica","Hidráulica","Acabamento","Pintura","Estrutura","Gesso","Outro"];
const PIE_COLORS = ["#f59e0b","#10b981","#3b82f6","#ef4444","#8b5cf6","#06b6d4","#f97316","#84cc16"];
const MODALIDADES = ["Empreitada","Construção e Venda"];
const SOCIO_COLORS = ["#f59e0b","#3b82f6","#10b981","#8b5cf6","#ef4444","#06b6d4","#f97316","#84cc16"];

// ═══════════════════════════════════════════════════════
// PIN DE ACESSO — Tela de bloqueio
// ═══════════════════════════════════════════════════════
const PIN_CORRETO = "540707"; // Altere aqui para o PIN que quiser
const PIN_STORAGE_KEY = "og_auth";
const PIN_VALIDADE_HORAS = 12;

function LoginScreen({onSuccess}){
  const [pin,setPin]=useState("");
  const [erro,setErro]=useState("");
  const [tentativas,setTentativas]=useState(0);
  const submit=()=>{
    if(pin===PIN_CORRETO){
      localStorage.setItem(PIN_STORAGE_KEY,String(Date.now()));
      onSuccess();
    }else{
      setTentativas(t=>t+1);
      setErro("PIN incorreto. Tente novamente.");
      setPin("");
    }
  };
  return(
    <div style={{background:"#060b14",height:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:20,fontFamily:"'DM Sans',sans-serif"}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');`}</style>
      <div style={{textAlign:"center",marginBottom:32}}>
        <div style={{fontSize:36,fontWeight:800,color:"#f59e0b",letterSpacing:1}}>OBRA GESTÃO</div>
        <div style={{fontSize:12,color:"#334155",letterSpacing:3,marginTop:4}}>FINANCEIRO</div>
      </div>
      <div style={{background:"#0d1220",border:"1px solid #1e2a3a",borderRadius:16,padding:"32px 28px",width:"100%",maxWidth:340}}>
        <div style={{fontSize:14,color:"#94a3b8",marginBottom:16,textAlign:"center",fontWeight:600}}>Digite o PIN de acesso</div>
        <input 
          type="password"
          value={pin}
          onChange={e=>{setPin(e.target.value);setErro("");}}
          onKeyDown={e=>{if(e.key==="Enter")submit();}}
          autoFocus
          maxLength={10}
          style={{width:"100%",background:"#111827",border:`1px solid ${erro?"#7f1d1d":"#1e2a3a"}`,borderRadius:10,padding:"14px 16px",color:"#e2e8f0",fontSize:20,textAlign:"center",letterSpacing:8,fontWeight:700,boxSizing:"border-box"}}
          placeholder="••••"
        />
        {erro&&<div style={{color:"#f87171",fontSize:13,textAlign:"center",marginTop:10}}>⚠ {erro}</div>}
        <button onClick={submit} style={{width:"100%",marginTop:16,padding:"12px 20px",borderRadius:10,border:"none",background:"#f59e0b",color:"#0d0d0d",fontSize:15,fontWeight:700,cursor:"pointer"}}>Entrar</button>
        {tentativas>=3&&<div style={{color:"#64748b",fontSize:11,textAlign:"center",marginTop:12}}>Após 5 tentativas você será bloqueado por 5 minutos.</div>}
      </div>
      <div style={{marginTop:24,fontSize:11,color:"#334155",textAlign:"center"}}>Sessão válida por {PIN_VALIDADE_HORAS}h</div>
    </div>
  );
}
// ═══════════════════════════════════════════════════════
const SUPABASE_URL = "https://cbrfmimwnpvhttrwyaib.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImNicmZtaW13bnB2aHR0cnd5YWliIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQ2Nzk1NjIsImV4cCI6MjEwMDI1NTU2Mn0.KAYrKdKN2XURRcbShNab551i-qi7p1sMG0WB4MVWIks";

const sbHeaders = {
  "apikey": SUPABASE_KEY,
  "Authorization": `Bearer ${SUPABASE_KEY}`,
  "Content-Type": "application/json",
  "Prefer": "resolution=merge-duplicates"
};

const TABLE_MAP = {"obras":"obras","lancamentos":"lancamentos","equipes":"equipes"};

const db = {
  async get(k){
    try{
      const table=TABLE_MAP[k];
      if(!table){const v=localStorage.getItem(k);return v?JSON.parse(v):null;}
      const res=await fetch(`${SUPABASE_URL}/rest/v1/${table}?select=data`,{headers:sbHeaders});
      if(!res.ok)throw new Error("get fail");
      const rows=await res.json();
      const arr=rows.map(r=>r.data);
      localStorage.setItem(k,JSON.stringify(arr));
      return arr;
    }catch(e){
      console.warn("db.get fallback:",e);
      try{const v=localStorage.getItem(k);return v?JSON.parse(v):null;}catch{return null;}
    }
  },
  async set(k,v){
    try{
      localStorage.setItem(k,JSON.stringify(v));
      const table=TABLE_MAP[k];
      if(!table)return;
      if(!Array.isArray(v))return;
      // Upsert todos
      if(v.length>0){
        const rows=v.map(item=>({id:item.id,data:item}));
        await fetch(`${SUPABASE_URL}/rest/v1/${table}`,{method:"POST",headers:sbHeaders,body:JSON.stringify(rows)});
      }
      // Apagar os que não estão no array
      const ids=v.map(x=>`"${x.id}"`).join(",")||`""`;
      await fetch(`${SUPABASE_URL}/rest/v1/${table}?id=not.in.(${ids})`,{method:"DELETE",headers:sbHeaders});
    }catch(e){console.warn("db.set:",e);}
  }
};

// ═══════════════════════════════════════════════════════
// UTILS
// ═══════════════════════════════════════════════════════
const fmt = v => (v||0).toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
const fmtN = v => (v||0).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtDate = d => d?new Date(d+"T00:00:00").toLocaleDateString("pt-BR"):"-";
const uid = () => Math.random().toString(36).slice(2,10)+Date.now().toString(36);
const today = () => new Date().toISOString().split("T")[0];
const pct = (a,b) => b>0?Math.min(100,Math.round((a/b)*100)):0;

// ═══════════════════════════════════════════════════════
// BACKUP EXCEL
// ═══════════════════════════════════════════════════════
function exportarExcel(obras, lancamentos, equipes){
  const wb = XLSX.utils.book_new();
  const obraMap = Object.fromEntries(obras.map(o=>[o.id,o.nome]));
  const wsObras = XLSX.utils.json_to_sheet(obras.map(o=>({
    ID:o.id,Nome:o.nome,Cliente:o.cliente,"Endereço":o.endereco||"",
    "Data Início":o.dataInicio||"","Data Término":o.dataFim||"",
    "Valor Contrato":o.valorContrato||0,Tipo:o.tipo,Status:o.status,
  })));
  XLSX.utils.book_append_sheet(wb,wsObras,"Obras");
  const wsLanc = XLSX.utils.json_to_sheet(lancamentos.map(l=>({
    ID:l.id,"ID Obra":l.obraId,Obra:obraMap[l.obraId]||"",
    Tipo:l.tipo,Categoria:l.categoria||"",Subcategoria:l.subcategoria||"",
    "Descrição":l.descricao||"",Valor:l.valor||0,Data:l.data||"",
    Pagamento:l.pagamento||"","Responsável":l.responsavel||"","Criado Em":l.criadoEm||"",
  })));
  XLSX.utils.book_append_sheet(wb,wsLanc,"Lançamentos");
  const wsEq = XLSX.utils.json_to_sheet(equipes.map(e=>({
    ID:e.id,Nome:e.nome,Tipo:e.tipo,"Responsável":e.responsavel||"",
  })));
  XLSX.utils.book_append_sheet(wb,wsEq,"Equipes");
  const resumo = obras.map(o=>{
    const lans=lancamentos.filter(l=>l.obraId===o.id);
    const recebido=lans.filter(l=>l.tipo==="Receita").reduce((s,l)=>s+l.valor,0);
    const despesas=lans.filter(l=>l.tipo==="Despesa").reduce((s,l)=>s+l.valor,0);
    return{Obra:o.nome,Cliente:o.cliente,Status:o.status,
      "Valor Contrato":o.valorContrato||0,"Total Recebido":recebido,
      "Total Despesas":despesas,"Saldo":recebido-despesas,
      "Lucro Projetado":(o.valorContrato||0)-despesas,
      "Margem %":o.valorContrato>0?(((o.valorContrato-despesas)/o.valorContrato)*100).toFixed(1)+"%":"0%"};
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(resumo),"Resumo Financeiro");
  const d=new Date().toLocaleDateString("pt-BR").replace(/\//g,"-");
  XLSX.writeFile(wb,`Backup_ObraGestao_${d}.xlsx`);
}

function importarExcel(file,setObras,setLancamentos,setEquipes,onDone){
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:"binary"});
      const obras=(XLSX.utils.sheet_to_json(wb.Sheets["Obras"]||{})||[]).map(r=>({
        id:r.ID||uid(),nome:r.Nome||"",cliente:r.Cliente||"",
        endereco:r["Endereço"]||"",dataInicio:r["Data Início"]||"",
        dataFim:r["Data Término"]||"",valorContrato:parseFloat(r["Valor Contrato"])||0,
        tipo:r.Tipo||"Padrão",status:r.Status||"Em Andamento",
      }));
      const lancamentos=(XLSX.utils.sheet_to_json(wb.Sheets["Lançamentos"]||{})||[]).map(r=>({
        id:r.ID||uid(),obraId:r["ID Obra"]||"",tipo:r.Tipo||"Despesa",
        categoria:r.Categoria||"",subcategoria:r.Subcategoria||"",
        descricao:r["Descrição"]||"",valor:parseFloat(r.Valor)||0,
        data:r.Data||"",pagamento:r.Pagamento||"PIX",
        responsavel:r["Responsável"]||"",criadoEm:r["Criado Em"]||new Date().toISOString(),
      }));
      const equipes=(XLSX.utils.sheet_to_json(wb.Sheets["Equipes"]||{})||[]).map(r=>({
        id:r.ID||uid(),nome:r.Nome||"",tipo:r.Tipo||"Alvenaria",responsavel:r["Responsável"]||"",
      }));
      if(obras.length>0)setObras(obras);
      if(lancamentos.length>0)setLancamentos(lancamentos);
      if(equipes.length>0)setEquipes(equipes);
      onDone(true,`${obras.length} obras, ${lancamentos.length} lançamentos e ${equipes.length} equipes restaurados!`);
    }catch{
      onDone(false,"Erro ao ler o arquivo. Use apenas backups gerados pelo sistema.");
    }
  };
  reader.readAsBinaryString(file);
}

// ═══════════════════════════════════════════════════════
// STYLES
// ═══════════════════════════════════════════════════════
const S = {
  app:{display:"flex",minHeight:"100vh",background:"#060b14",fontFamily:"'DM Sans',sans-serif",color:"#e2e8f0"},
  sidebar:{width:220,background:"#0d1220",borderRight:"1px solid #1e2a3a",display:"flex",flexDirection:"column",flexShrink:0},
  main:{flex:1,overflow:"auto",padding:20,background:"#060b14",minWidth:0},
  card:{background:"#0d1220",border:"1px solid #1e2a3a",borderRadius:12,padding:"18px 20px"},
  cardSm:{background:"#0d1220",border:"1px solid #1e2a3a",borderRadius:10,padding:"14px 16px"},
  input:{width:"100%",background:"#111827",border:"1px solid #1e2a3a",borderRadius:8,padding:"10px 14px",color:"#e2e8f0",fontSize:14,boxSizing:"border-box"},
  btn:{padding:"10px 20px",borderRadius:8,fontSize:14,fontWeight:600,cursor:"pointer",border:"none"},
  btnPrimary:{padding:"10px 20px",borderRadius:8,fontSize:14,fontWeight:600,cursor:"pointer",border:"none",background:"#f59e0b",color:"#0d0d0d"},
  btnDanger:{padding:"8px 16px",borderRadius:8,fontSize:13,fontWeight:600,cursor:"pointer",border:"none",background:"#7f1d1d",color:"#fca5a5"},
  btnGhost:{padding:"8px 14px",borderRadius:8,fontSize:13,cursor:"pointer",border:"1px solid #1e2a3a",background:"transparent",color:"#94a3b8"},
  tag:{display:"inline-flex",alignItems:"center",padding:"3px 10px",borderRadius:20,fontSize:12,fontWeight:600},
  modal:{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100,padding:16},
  modalBox:{background:"#0d1220",border:"1px solid #1e2a3a",borderRadius:16,padding:24,width:"100%",maxWidth:560,maxHeight:"90vh",overflow:"auto"},
  label:{display:"block",fontSize:13,color:"#64748b",marginBottom:6,fontWeight:500},
  row:{display:"flex",gap:16},
  h1:{fontFamily:"'DM Sans',sans-serif",fontSize:24,fontWeight:700,color:"#f1f5f9",letterSpacing:"-0.3px"},
  h2:{fontFamily:"'DM Sans',sans-serif",fontSize:18,fontWeight:700,color:"#f1f5f9",letterSpacing:"-0.2px"},
  h3:{fontFamily:"'DM Sans',sans-serif",fontSize:15,fontWeight:700,color:"#f1f5f9",letterSpacing:0},
  muted:{color:"#64748b",fontSize:14},
  grid2:{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))",gap:14},
  grid3:{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:12},
  grid4:{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12},
};

const statusColor = s => s==="Em Andamento"?{bg:"#052e16",text:"#4ade80"}:s==="Finalizada"?{bg:"#1e3a5f",text:"#93c5fd"}:{bg:"#431407",text:"#fdba74"};
const tipoColor = t => t==="Alto Padrão"?{bg:"#2d1b69",text:"#c4b5fd"}:t==="Padrão"?{bg:"#1e3a5f",text:"#93c5fd"}:{bg:"#1a2e1a",text:"#86efac"};

// ═══════════════════════════════════════════════════════
// SHARED COMPONENTS
// ═══════════════════════════════════════════════════════
function KpiCard({label,value,sub,color="#f59e0b",icon}){
  return(
    <div style={{...S.cardSm,position:"relative",overflow:"hidden",minWidth:0}}>
      <div style={{position:"absolute",top:-8,right:-8,fontSize:36,opacity:.05,lineHeight:1}}>{icon}</div>
      <div style={{fontSize:11,color:"#64748b",marginBottom:8,fontWeight:600,textTransform:"uppercase",letterSpacing:.8,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{label}</div>
      <div style={{fontSize:15,fontWeight:700,color,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{value}</div>
      {sub&&<div style={{fontSize:11,color:"#475569",marginTop:4,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{sub}</div>}
    </div>
  );
}

function Tag({s,style}){
  const c=statusColor(s);
  return <span style={{...S.tag,background:c.bg,color:c.text,...style}}>{s}</span>;
}

function Modal({open,onClose,title,children}){
  if(!open)return null;
  return(
    <div style={S.modal} onClick={e=>{if(e.target===e.currentTarget)onClose()}}>
      <div style={S.modalBox}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <span style={S.h2}>{title}</span>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#64748b",fontSize:22,cursor:"pointer",lineHeight:1}}>×</button>
        </div>
        {children}
      </div>
    </div>
  );
}

function FormField({label,children}){
  return <div style={{marginBottom:14}}><label style={S.label}>{label}</label>{children}</div>;
}

function Alert({type,msg}){
  const c=type==="danger"?{bg:"#450a0a",border:"#7f1d1d",text:"#fca5a5",icon:"⚠"}:{bg:"#431407",border:"#7c2d12",text:"#fdba74",icon:"🔔"};
  return(
    <div style={{background:c.bg,border:`1px solid ${c.border}`,borderRadius:8,padding:"10px 14px",display:"flex",gap:8,alignItems:"center",marginBottom:8}}>
      <span style={{fontSize:14}}>{c.icon}</span>
      <span style={{color:c.text,fontSize:13}}>{msg}</span>
    </div>
  );
}

// ═══════════════════════════════════════════════════════
// SIDEBAR
// ═══════════════════════════════════════════════════════
const NAV = [
  {id:"dashboard",label:"Dashboard Geral",icon:"⊞"},
  {id:"obras",label:"Obras",icon:"◧"},
  {id:"financeiro",label:"Financeiro",icon:"◈"},
  {id:"mao-de-obra",label:"Mão de Obra",icon:"◉"},
  {id:"relatorios",label:"Relatórios",icon:"◳"},
];

function Sidebar({view,setView,onExportar,onImportar,onLogout}){
  const inputRef=useRef();
  return(
    <>
      <div style={{padding:"24px 20px 16px"}}>
        <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:18,fontWeight:800,color:"#f59e0b",letterSpacing:1}}>OBRA</div>
        <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:18,fontWeight:800,color:"#f1f5f9",letterSpacing:1,marginTop:-4}}>GESTÃO</div>
        <div style={{fontSize:11,color:"#334155",marginTop:2,letterSpacing:2}}>FINANCEIRO</div>
      </div>
      <div style={{height:1,background:"#1e2a3a",margin:"0 16px 12px"}}/>
      <nav style={{flex:1,padding:"0 8px"}}>
        {NAV.map(n=>(
          <button key={n.id} onClick={()=>setView(n.id)}
            style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"10px 12px",borderRadius:8,border:"none",cursor:"pointer",marginBottom:2,
              background:view===n.id?"#1a2b45":"transparent",
              color:view===n.id?"#f59e0b":"#64748b",fontSize:14,fontWeight:view===n.id?600:400,textAlign:"left",transition:"all .15s"}}>
            <span style={{fontSize:16,opacity:.9}}>{n.icon}</span>{n.label}
          </button>
        ))}
      </nav>
      <div style={{padding:"16px 12px",borderTop:"1px solid #1e2a3a",display:"flex",flexDirection:"column",gap:8}}>
        <div style={{fontSize:11,color:"#334155",textAlign:"center",marginBottom:4,letterSpacing:1}}>BACKUP DE DADOS</div>
        <button onClick={onExportar}
          style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1px solid #166534",background:"#052e16",color:"#4ade80",fontSize:13,fontWeight:600,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
          ⬇ Exportar Excel
        </button>
        <button onClick={()=>inputRef.current.click()}
          style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1px solid #1e3a5f",background:"#0c1f3a",color:"#93c5fd",fontSize:13,fontWeight:600,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
          ⬆ Restaurar Backup
        </button>
        <input ref={inputRef} type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>{if(e.target.files[0])onImportar(e.target.files[0]);e.target.value="";}}/>
        <button onClick={onLogout}
          style={{width:"100%",marginTop:8,padding:"8px 12px",borderRadius:8,border:"1px solid #7f1d1d",background:"#450a0a",color:"#fca5a5",fontSize:12,fontWeight:600,cursor:"pointer"}}>
          🔒 Bloquear Acesso
        </button>
        <div style={{fontSize:11,color:"#1e3a5f",textAlign:"center",marginTop:2}}>v1.0 · Gestão de Obras</div>
      </div>
    </>
  );
}

// ═══════════════════════════════════════════════════════
// OBRA FORM
// ═══════════════════════════════════════════════════════
function ObraForm({obra,onSave,onClose}){
  const init = obra||{nome:"",cliente:"",endereco:"",dataInicio:today(),dataFim:"",valorContrato:"",metrosQuadrados:"",tipo:"Padrão",status:"Em Andamento",modalidade:"Empreitada",socios:[],valorVendaEstimado:"",valorVendaReal:""};
  const [f,setF]=useState({...init,socios:init.socios||[]});
  const [novoSocio,setNovoSocio]=useState({nome:"",percentual:""});
  const set = k=>e=>setF(p=>({...p,[k]:e.target.value}));
  
  const contrato=parseFloat(f.valorContrato)||0;
  const m2=parseFloat(f.metrosQuadrados)||0;
  const valorM2=m2>0?(contrato/m2):0;
  const totalPerc=f.socios.reduce((s,x)=>s+(parseFloat(x.percentual)||0),0);
  
  const addSocio=()=>{
    if(!novoSocio.nome)return alert("Nome do sócio obrigatório");
    const p=parseFloat(novoSocio.percentual)||0;
    if(p<=0||p>100)return alert("Percentual deve estar entre 0 e 100");
    if(totalPerc+p>100)return alert(`Total dos percentuais não pode passar de 100%. Disponível: ${(100-totalPerc).toFixed(2)}%`);
    setF(p=>({...p,socios:[...p.socios,{...novoSocio,percentual:parseFloat(novoSocio.percentual),id:uid()}]}));
    setNovoSocio({nome:"",percentual:""});
  };
  const removeSocio=id=>setF(p=>({...p,socios:p.socios.filter(s=>s.id!==id)}));
  
  const save=()=>{
    if(!f.nome||!f.cliente)return alert("Nome e cliente são obrigatórios");
    if(f.modalidade==="Construção e Venda"&&f.socios.length>0&&Math.abs(totalPerc-100)>0.01)
      return alert(`Os percentuais dos sócios precisam somar 100%. Atual: ${totalPerc.toFixed(2)}%`);
    onSave({...f,valorContrato:contrato,metrosQuadrados:m2,valorVendaEstimado:parseFloat(f.valorVendaEstimado)||0,valorVendaReal:parseFloat(f.valorVendaReal)||0,id:f.id||uid()});
  };
  
  return(
    <>
      <FormField label="Modalidade *">
        <select style={S.input} value={f.modalidade} onChange={set("modalidade")}>{MODALIDADES.map(t=><option key={t}>{t}</option>)}</select>
        <div style={{fontSize:11,color:"#475569",marginTop:4}}>
          {f.modalidade==="Empreitada"?"Obra para terceiros com contrato fechado.":"Obra própria que será vendida — permite gestão de sócios."}
        </div>
      </FormField>
      <div style={S.grid2}>
        <FormField label="Nome da Obra *"><input style={S.input} value={f.nome} onChange={set("nome")} placeholder="Ex: Residência Silva"/></FormField>
        <FormField label={f.modalidade==="Construção e Venda"?"Proprietário / Empreendimento *":"Cliente *"}><input style={S.input} value={f.cliente} onChange={set("cliente")} placeholder="Nome"/></FormField>
      </div>
      <FormField label="Endereço"><input style={S.input} value={f.endereco} onChange={set("endereco")} placeholder="Rua, número, bairro, cidade"/></FormField>
      <div style={S.grid2}>
        <FormField label="Data de Início"><input style={S.input} type="date" value={f.dataInicio} onChange={set("dataInicio")}/></FormField>
        <FormField label="Previsão de Término"><input style={S.input} type="date" value={f.dataFim} onChange={set("dataFim")}/></FormField>
      </div>
      <div style={S.grid2}>
        <FormField label={f.modalidade==="Construção e Venda"?"Orçamento de Construção (R$)":"Valor do Contrato (R$)"}><input style={S.input} type="number" value={f.valorContrato} onChange={set("valorContrato")} placeholder="0,00"/></FormField>
        <FormField label="Área da Construção (m²)"><input style={S.input} type="number" step="0.01" value={f.metrosQuadrados} onChange={set("metrosQuadrados")} placeholder="Ex: 85,50"/></FormField>
      </div>
      {contrato>0&&m2>0&&(
        <div style={{background:"#0c1f3a",border:"1px solid #1e3a5f",borderRadius:8,padding:"10px 16px",marginBottom:14,display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:12,color:"#64748b"}}>Valor por m²:</span>
          <span style={{fontSize:15,fontWeight:700,color:"#f59e0b"}}>{fmt(valorM2)}/m²</span>
        </div>
      )}
      
      {f.modalidade==="Construção e Venda"&&(
        <>
          <div style={S.grid2}>
            <FormField label="Valor de Venda Estimado (R$)"><input style={S.input} type="number" value={f.valorVendaEstimado} onChange={set("valorVendaEstimado")} placeholder="0,00"/></FormField>
            <FormField label="Valor de Venda Real (R$)"><input style={S.input} type="number" value={f.valorVendaReal} onChange={set("valorVendaReal")} placeholder="Quando vender"/></FormField>
          </div>
          
          <div style={{background:"#0c1f3a",border:"1px solid #1e3a5f",borderRadius:8,padding:"14px 16px",marginBottom:14}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
              <div>
                <div style={{fontSize:13,color:"#93c5fd",fontWeight:700}}>SÓCIOS DO EMPREENDIMENTO</div>
                <div style={{fontSize:11,color:"#475569",marginTop:2}}>Define quem participa e em qual percentual. Total: <span style={{color:Math.abs(totalPerc-100)<0.01?"#4ade80":totalPerc>100?"#f87171":"#f59e0b",fontWeight:700}}>{totalPerc.toFixed(2)}%</span></div>
              </div>
            </div>
            {f.socios.length>0&&(
              <div style={{marginBottom:12,display:"flex",flexDirection:"column",gap:6}}>
                {f.socios.map((s,i)=>(
                  <div key={s.id} style={{display:"flex",alignItems:"center",gap:10,background:"#0d1a2e",borderRadius:6,padding:"8px 12px"}}>
                    <div style={{width:10,height:10,borderRadius:5,background:SOCIO_COLORS[i%SOCIO_COLORS.length],flexShrink:0}}/>
                    <span style={{flex:1,color:"#f1f5f9",fontSize:13,fontWeight:600}}>{s.nome}</span>
                    <span style={{color:"#f59e0b",fontWeight:700,fontSize:13}}>{s.percentual}%</span>
                    {contrato>0&&<span style={{color:"#64748b",fontSize:12}}>· {fmt(contrato*s.percentual/100)}</span>}
                    <button onClick={()=>removeSocio(s.id)} style={{background:"none",border:"none",color:"#475569",cursor:"pointer",fontSize:14}}>✕</button>
                  </div>
                ))}
              </div>
            )}
            <div style={{display:"flex",gap:8}}>
              <input style={{...S.input,flex:1,padding:"8px 12px",fontSize:13}} placeholder="Nome do sócio" value={novoSocio.nome} onChange={e=>setNovoSocio(p=>({...p,nome:e.target.value}))}/>
              <input style={{...S.input,width:80,padding:"8px 12px",fontSize:13}} type="number" placeholder="% " value={novoSocio.percentual} onChange={e=>setNovoSocio(p=>({...p,percentual:e.target.value}))}/>
              <button onClick={addSocio} style={{padding:"8px 14px",borderRadius:8,border:"1px solid #1e3a5f",background:"#0d1a2e",color:"#93c5fd",fontSize:13,fontWeight:600,cursor:"pointer"}}>+ Adicionar</button>
            </div>
          </div>
        </>
      )}
      
      <div style={S.grid2}>
        <FormField label="Tipo de Obra">
          <select style={S.input} value={f.tipo} onChange={set("tipo")}>{TIPOS_OBRA.map(t=><option key={t}>{t}</option>)}</select>
        </FormField>
        <FormField label="Status">
          <select style={S.input} value={f.status} onChange={set("status")}>{STATUS_OBRA.map(s=><option key={s}>{s}</option>)}</select>
        </FormField>
      </div>
      <div style={{display:"flex",gap:12,justifyContent:"flex-end",marginTop:20}}>
        <button style={{...S.btnGhost}} onClick={onClose}>Cancelar</button>
        <button style={S.btnPrimary} onClick={save}>{obra?"Salvar Alterações":"Criar Obra"}</button>
      </div>
    </>
  );
}

// ═══════════════════════════════════════════════════════
// OBRAS VIEW
// ═══════════════════════════════════════════════════════
function ObrasView({obras,setObras,obraStats,setView,setSelectedObra,goToObra}){
  const [modal,setModal]=useState(false);
  const [editing,setEditing]=useState(null);
  const [filter,setFilter]=useState("Todos");
  const [del,setDel]=useState(null);

  const filtered = filter==="Todos"?obras:obras.filter(o=>o.status===filter);
  const save = o=>{
    setObras(p=>editing?p.map(x=>x.id===o.id?o:x):[...p,o]);
    setModal(false);setEditing(null);
  };
  const remove=id=>{setObras(p=>p.filter(x=>x.id!==id));setDel(null)};

  return(
    <div>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500;600&display=swap');`}</style>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:24}}>
        <div>
          <h1 style={S.h1}>Obras</h1>
          <p style={S.muted}>{obras.length} obra{obras.length!==1?"s":""} cadastrada{obras.length!==1?"s":""}</p>
        </div>
        <button style={S.btnPrimary} onClick={()=>{setEditing(null);setModal(true)}}>+ Nova Obra</button>
      </div>
      <div style={{display:"flex",gap:8,marginBottom:20}}>
        {["Todos",...STATUS_OBRA].map(s=>(
          <button key={s} onClick={()=>setFilter(s)}
            style={{...S.btnGhost,background:filter===s?"#1a2b45":"transparent",color:filter===s?"#f59e0b":"#64748b",borderColor:filter===s?"#1e3a5f":"#1e2a3a"}}>
            {s}
          </button>
        ))}
      </div>
      {filtered.length===0&&<div style={{...S.card,textAlign:"center",padding:48,color:"#334155"}}>Nenhuma obra encontrada. Crie a primeira!</div>}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:16}}>
        {filtered.map(o=>{
          const st=obraStats[o.id]||{};
          const lucro=o.valorContrato-(st.despesas||0);
          const positivo=lucro>=0;
          const tc=tipoColor(o.tipo);
          return(
            <div key={o.id} style={{...S.card,cursor:"pointer",transition:"border-color .2s",borderColor:"#1e2a3a"}}
              onMouseEnter={e=>e.currentTarget.style.borderColor="#f59e0b"}
              onMouseLeave={e=>e.currentTarget.style.borderColor="#1e2a3a"}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
                <div style={{flex:1}}>
                  <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:16,fontWeight:700,color:"#f1f5f9",marginBottom:4}}>{o.nome}</div>
                  <div style={{fontSize:13,color:"#64748b"}}>{o.cliente}</div>
                </div>
                <Tag s={o.status}/>
              </div>
              {o.endereco&&<div style={{fontSize:12,color:"#475569",marginBottom:10}}>📍 {o.endereco}</div>}
              <span style={{...S.tag,...tc,marginBottom:12,display:"inline-flex"}}>{o.tipo}</span>
              <div style={{height:1,background:"#1e2a3a",margin:"12px 0"}}/>
              <div style={S.grid2}>
                <div><div style={{fontSize:11,color:"#475569",marginBottom:2}}>CONTRATO</div><div style={{fontSize:15,fontWeight:700,color:"#f59e0b"}}>{fmt(o.valorContrato)}</div></div>
                <div><div style={{fontSize:11,color:"#475569",marginBottom:2}}>SALDO</div><div style={{fontSize:15,fontWeight:700,color:positivo?"#4ade80":"#f87171"}}>{fmt(st.saldo||0)}</div></div>
                <div><div style={{fontSize:11,color:"#475569",marginBottom:2}}>RECEBIDO</div><div style={{fontSize:13,color:"#94a3b8"}}>{fmt(st.recebido||0)}</div></div>
                <div><div style={{fontSize:11,color:"#475569",marginBottom:2}}>GASTO</div><div style={{fontSize:13,color:"#94a3b8"}}>{fmt(st.despesas||0)}</div></div>
              </div>
              <div style={{marginTop:12}}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:"#475569",marginBottom:4}}>
                  <span>Execução orçamentária</span><span>{pct(st.despesas||0,o.valorContrato)}%</span>
                </div>
                <div style={{background:"#1e2a3a",borderRadius:4,height:6,overflow:"hidden"}}>
                  <div style={{height:"100%",background:pct(st.despesas||0,o.valorContrato)>100?"#ef4444":"#f59e0b",width:`${Math.min(100,pct(st.despesas||0,o.valorContrato))}%`,borderRadius:4,transition:"width .5s"}}/>
                </div>
              </div>
              <div style={{display:"flex",gap:8,marginTop:16}}>
                <button style={{...S.btnPrimary,flex:1,padding:"8px 0",fontSize:13}} onClick={()=>goToObra(o.id)}>Ver Dashboard</button>
                <button style={{...S.btnGhost,padding:"8px 14px",fontSize:13}} onClick={e=>{e.stopPropagation();setEditing(o);setModal(true)}}>Editar</button>
                <button style={{...S.btnDanger,padding:"8px 14px",fontSize:13}} onClick={e=>{e.stopPropagation();setDel(o.id)}}>✕</button>
              </div>
            </div>
          );
        })}
      </div>
      <Modal open={modal} onClose={()=>{setModal(false);setEditing(null)}} title={editing?"Editar Obra":"Nova Obra"}>
        <ObraForm obra={editing} onSave={save} onClose={()=>{setModal(false);setEditing(null)}}/>
      </Modal>
      <Modal open={!!del} onClose={()=>setDel(null)} title="Confirmar Exclusão">
        <p style={{color:"#94a3b8",marginBottom:20}}>Tem certeza? Todos os lançamentos desta obra também serão removidos.</p>
        <div style={{display:"flex",gap:12,justifyContent:"flex-end"}}>
          <button style={S.btnGhost} onClick={()=>setDel(null)}>Cancelar</button>
          <button style={S.btnDanger} onClick={()=>remove(del)}>Excluir</button>
        </div>
      </Modal>
    </div>
  );
}

// ═══════════════════════════════════════════════════════
// RELATORIO INTERNO DA OBRA — Com filtros e PDF
// ═══════════════════════════════════════════════════════
function RelatorioObraSection({obra,lancamentos,obraStats}){
  const [tipo,setTipo]=useState("");
  const [categoria,setCategoria]=useState("");
  const [subcategoria,setSubcategoria]=useState("");
  const [dataIni,setDataIni]=useState("");
  const [dataFim,setDataFim]=useState("");
  const [search,setSearch]=useState("");
  const [gerandoPDF,setGerandoPDF]=useState(false);

  const filtrados=useMemo(()=>{
    return lancamentos.filter(l=>{
      if(tipo&&l.tipo!==tipo)return false;
      if(categoria&&l.categoria!==categoria)return false;
      if(subcategoria&&l.subcategoria!==subcategoria)return false;
      if(dataIni&&l.data<dataIni)return false;
      if(dataFim&&l.data>dataFim)return false;
      if(search&&!`${l.descricao}${l.categoria}${l.subcategoria}${l.responsavel}`.toLowerCase().includes(search.toLowerCase()))return false;
      return true;
    }).sort((a,b)=>b.data>a.data?1:-1);
  },[lancamentos,tipo,categoria,subcategoria,dataIni,dataFim,search]);

  const totalRec=filtrados.filter(l=>l.tipo==="Receita").reduce((s,l)=>s+l.valor,0);
  const totalDesp=filtrados.filter(l=>l.tipo==="Despesa").reduce((s,l)=>s+l.valor,0);
  const limpar=()=>{setTipo("");setCategoria("");setSubcategoria("");setDataIni("");setDataFim("");setSearch("");};
  const temFiltro=tipo||categoria||subcategoria||dataIni||dataFim||search;

  const handlePDF=async()=>{
    setGerandoPDF(true);
    try{ await gerarPDFObra(obra,filtrados.length<lancamentos.length?filtrados.map(l=>({...l,obraId:obra.id})):lancamentos,obraStats); }
    catch{ alert("Erro ao gerar PDF."); }
    setGerandoPDF(false);
  };

  return(
    <div style={S.card}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:8}}>
        <div>
          <h3 style={S.h3}>Relatório de Lançamentos</h3>
          <p style={{fontSize:12,color:"#64748b",marginTop:2}}>Filtre os lançamentos e gere relatório em PDF</p>
        </div>
        <button onClick={handlePDF} disabled={gerandoPDF||lancamentos.length===0}
          style={{padding:"9px 18px",borderRadius:8,border:"none",cursor:gerandoPDF||lancamentos.length===0?"not-allowed":"pointer",
            background:lancamentos.length===0?"#1e2a3a":"#450a0a",color:lancamentos.length===0?"#475569":"#fca5a5",
            fontSize:13,fontWeight:600,display:"flex",alignItems:"center",gap:6}}>
          {gerandoPDF?"⏳ Gerando...":"📄 Gerar Relatório PDF"}
        </button>
      </div>

      {/* Filtros */}
      <div style={{background:"#0d1220",border:"1px solid #1e2a3a",borderRadius:10,padding:"12px",marginBottom:14}}>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:8}}>
          <input style={{...S.input,flex:1,minWidth:160,padding:"8px 12px",fontSize:13}} placeholder="Buscar descrição, responsável..." value={search} onChange={e=>setSearch(e.target.value)}/>
          <select style={{...S.input,width:130,padding:"8px 12px",fontSize:13}} value={tipo} onChange={e=>setTipo(e.target.value)}>
            <option value="">Todos tipos</option>
            <option>Receita</option><option>Despesa</option>
          </select>
          <select style={{...S.input,width:150,padding:"8px 12px",fontSize:13}} value={categoria} onChange={e=>{setCategoria(e.target.value);setSubcategoria("");}}>
            <option value="">Toda categoria</option>
            {CATS.map(c=><option key={c}>{c}</option>)}
          </select>
          {categoria&&(
            <select style={{...S.input,width:150,padding:"8px 12px",fontSize:13}} value={subcategoria} onChange={e=>setSubcategoria(e.target.value)}>
              <option value="">Toda subcategoria</option>
              {(SUBCATS[categoria]||[]).map(s=><option key={s}>{s}</option>)}
            </select>
          )}
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <span style={{fontSize:12,color:"#64748b"}}>Período:</span>
          <input style={{...S.input,width:150,padding:"8px 12px",fontSize:13}} type="date" value={dataIni} onChange={e=>setDataIni(e.target.value)}/>
          <span style={{fontSize:12,color:"#475569"}}>até</span>
          <input style={{...S.input,width:150,padding:"8px 12px",fontSize:13}} type="date" value={dataFim} onChange={e=>setDataFim(e.target.value)}/>
          {temFiltro&&<button style={{...S.btnGhost,padding:"6px 12px",fontSize:12}} onClick={limpar}>✕ Limpar filtros</button>}
          <div style={{marginLeft:"auto",fontSize:12,color:"#64748b"}}>
            <span style={{color:"#4ade80",fontWeight:600}}>+{fmt(totalRec)}</span>
            <span style={{margin:"0 6px",color:"#334155"}}>·</span>
            <span style={{color:"#f87171",fontWeight:600}}>-{fmt(totalDesp)}</span>
            <span style={{margin:"0 6px",color:"#334155"}}>·</span>
            <span style={{color:(totalRec-totalDesp)>=0?"#4ade80":"#f87171",fontWeight:700}}>{fmt(totalRec-totalDesp)}</span>
          </div>
        </div>
      </div>

      {filtrados.length===0?<div style={{color:"#334155",fontSize:13,textAlign:"center",padding:24}}>Nenhum lançamento encontrado com esses filtros</div>:(
        <div style={{overflowX:"auto",maxHeight:500}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead style={{position:"sticky",top:0,background:"#0d1220",zIndex:1}}>
              <tr style={{borderBottom:"1px solid #1e2a3a"}}>
                {["Data","Tipo","Categoria","Subcategoria","Descrição","Responsável","Pgto","Valor"].map(h=>(
                  <th key={h} style={{textAlign:"left",padding:"8px 12px",color:"#475569",fontWeight:500,whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtrados.map(l=>(
                <tr key={l.id} style={{borderBottom:"1px solid #111827",transition:"background .1s"}}
                  onMouseEnter={e=>e.currentTarget.style.background="#111827"}
                  onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <td style={{padding:"10px 12px",color:"#64748b",whiteSpace:"nowrap"}}>{fmtDate(l.data)}</td>
                  <td style={{padding:"10px 12px"}}><span style={{...S.tag,background:l.tipo==="Receita"?"#052e16":"#450a0a",color:l.tipo==="Receita"?"#4ade80":"#f87171"}}>{l.tipo}</span></td>
                  <td style={{padding:"10px 12px",color:"#94a3b8"}}>{l.categoria||"—"}</td>
                  <td style={{padding:"10px 12px",color:"#64748b"}}>{l.subcategoria||"—"}</td>
                  <td style={{padding:"10px 12px",color:"#e2e8f0",maxWidth:200}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{l.descricao||"—"}</div></td>
                  <td style={{padding:"10px 12px",color:"#64748b"}}>{l.responsavel||"—"}</td>
                  <td style={{padding:"10px 12px",color:"#475569",whiteSpace:"nowrap"}}>{l.pagamento||"—"}</td>
                  <td style={{padding:"10px 12px",fontWeight:700,color:l.tipo==="Receita"?"#4ade80":"#f87171",whiteSpace:"nowrap"}}>{l.tipo==="Receita"?"+":"-"}{fmt(l.valor)}</td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr style={{borderTop:"2px solid #1e2a3a",fontWeight:700}}>
                <td colSpan={7} style={{padding:"10px 12px",color:"#64748b",fontSize:12}}>{filtrados.length} lançamento(s)</td>
                <td style={{padding:"10px 12px",color:(totalRec-totalDesp)>=0?"#4ade80":"#f87171",whiteSpace:"nowrap"}}>{fmt(totalRec-totalDesp)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════
// OBRA DETAIL DASHBOARD
// ═══════════════════════════════════════════════════════
function ObraDetail({obra,stats,lancamentos,setView,onUpdateObra,addLancamento,allObras}){
  if(!obra)return null;
  const st=stats||{};

  const [modalAditivo,setModalAditivo]=useState(false);
  const [novoAditivo,setNovoAditivo]=useState({descricao:"",valor:"",data:today(),dataCompensacao:"",tipo:"Acréscimo"});
  const [modalLanc,setModalLanc]=useState(false);
  const [defaultCatLanc,setDefaultCatLanc]=useState(null);
  const [defaultTipoLanc,setDefaultTipoLanc]=useState(null);

  const abrirLancamento=(cat=null,tipo=null)=>{setDefaultCatLanc(cat);setDefaultTipoLanc(tipo);setModalLanc(true);};

  const aditivos=obra.aditivos||[];
  const totalAditivos=aditivos.reduce((s,a)=>s+(a.tipo==="Acréscimo"?a.valor:-a.valor),0);
  const contratoAtualizado=(obra.valorContrato||0)+totalAditivos;
  const lucro=contratoAtualizado-(st.despesas||0);

  const saveAditivo=()=>{
    if(!novoAditivo.descricao)return alert("Descrição obrigatória");
    if(!novoAditivo.valor||parseFloat(novoAditivo.valor)<=0)return alert("Valor inválido");
    const novo={...novoAditivo,valor:parseFloat(novoAditivo.valor),id:uid(),criadoEm:new Date().toISOString()};
    onUpdateObra({...obra,aditivos:[...aditivos,novo]});
    setNovoAditivo({descricao:"",valor:"",data:today(),dataCompensacao:"",tipo:"Acréscimo"});
    setModalAditivo(false);
  };
  const removeAditivo=id=>onUpdateObra({...obra,aditivos:aditivos.filter(a=>a.id!==id)});

  const receitas=lancamentos.filter(l=>l.tipo==="Receita");
  const despesas=lancamentos.filter(l=>l.tipo==="Despesa");
  const byCategoria=CATS.map(c=>({name:c,value:despesas.filter(l=>l.categoria===c).reduce((s,l)=>s+l.valor,0)})).filter(x=>x.value>0);

  const byMonth=useMemo(()=>{
    const map={};
    lancamentos.forEach(l=>{
      const m=l.data?.slice(0,7)||"";
      if(!m)return;
      if(!map[m])map[m]={mes:m,receita:0,despesa:0};
      if(l.tipo==="Receita")map[m].receita+=l.valor;
      else map[m].despesa+=l.valor;
    });
    return Object.values(map).sort((a,b)=>a.mes>b.mes?1:-1).map(x=>({...x,mes:new Date(x.mes+"-01").toLocaleDateString("pt-BR",{month:"short",year:"2-digit"})}));
  },[lancamentos]);

  return(
    <div>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:16}}>
        <div style={{flex:1,minWidth:0}}>
          {(()=>{const len=(obra.nome||"").length;const fs=len>30?18:len>20?20:len>14?22:26;return(
            <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:fs,fontWeight:800,color:"#f1f5f9",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{obra.nome}</div>
          );})()} 
          <div style={{display:"flex",gap:8,alignItems:"center",marginTop:4,flexWrap:"wrap"}}>
            <Tag s={obra.status}/>
            {obra.modalidade==="Construção e Venda"&&<span style={{...S.tag,background:"#2d1b69",color:"#c4b5fd"}}>🏗 Construção e Venda</span>}
            <span style={S.muted}>· {obra.cliente}</span>
            {obra.endereco&&<span style={S.muted}>· 📍 {obra.endereco}</span>}
            {obra.metrosQuadrados>0&&<span style={S.muted}>· {obra.metrosQuadrados} m²</span>}
          </div>
        </div>
        <div style={{textAlign:"right",flexShrink:0}}>
          <div style={{fontSize:11,color:"#475569"}}>Contrato Original</div>
          {(()=>{const len=fmt(obra.valorContrato).length;const fs=len>16?13:len>13?15:17;return(
            <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:fs,fontWeight:700,color:"#64748b",letterSpacing:0,textDecoration:totalAditivos!==0?"line-through":"none"}}>{fmt(obra.valorContrato)}</div>
          );})()} 
          {totalAditivos!==0&&(
            <>
              <div style={{fontSize:11,color:"#475569",marginTop:2}}>Contrato Atualizado</div>
              {(()=>{const len=fmt(contratoAtualizado).length;const fs=len>16?14:len>13?16:18;return(
                <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:fs,fontWeight:700,color:"#f59e0b",letterSpacing:0}}>{fmt(contratoAtualizado)}</div>
              );})()} 
              <div style={{fontSize:11,color:totalAditivos>=0?"#4ade80":"#f87171",marginTop:1}}>{totalAditivos>=0?"+":""}{fmt(totalAditivos)} em aditivos</div>
            </>
          )}
          {totalAditivos===0&&(()=>{const len=fmt(obra.valorContrato).length;const fs=len>16?14:len>13?16:18;return(
            <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:fs,fontWeight:700,color:"#f59e0b",letterSpacing:0}}>{fmt(obra.valorContrato)}</div>
          );})()} 
          {obra.metrosQuadrados>0&&<div style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,fontWeight:600,color:"#64748b",marginTop:2}}>{fmt(contratoAtualizado/obra.metrosQuadrados)}/m²</div>}
        </div>
      </div>

      {/* Barra de Lançamento Rápido para esta obra */}
      <div style={{background:"linear-gradient(135deg,#0d1a2e 0%,#111827 100%)",border:"1px solid #1e2a3a",borderRadius:12,padding:"14px 20px",marginBottom:16}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,gap:12,flexWrap:"wrap"}}>
          <div>
            <div style={{fontSize:12,color:"#f59e0b",fontWeight:700,letterSpacing:.8}}>⚡ LANÇAMENTO RÁPIDO NESTA OBRA</div>
            <div style={{fontSize:11,color:"#475569",marginTop:2}}>Registre receitas e despesas já vinculadas a {obra.nome}</div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {[
            {cat:null,label:"💰 Receita",bg:"#052e16",border:"#166534",text:"#4ade80",tipo:"Receita"},
            {cat:"Material",label:"🧱 Material",bg:"#422006",border:"#7c2d12",text:"#fdba74"},
            {cat:"Mão de Obra",label:"👷 Mão de Obra",bg:"#2d1b69",border:"#4c1d95",text:"#c4b5fd"},
            {cat:"Equipamento",label:"🔧 Equipamento",bg:"#0c1f3a",border:"#1e3a5f",text:"#93c5fd"},
            {cat:"Comissão",label:"💼 Comissão",bg:"#451a03",border:"#92400e",text:"#fcd34d"},
            {cat:"Outros",label:"📦 Outros",bg:"#1a2e1a",border:"#166534",text:"#86efac"},
          ].map(b=>(
            <button key={b.label} onClick={()=>abrirLancamento(b.cat,b.tipo)}
              style={{padding:"9px 16px",borderRadius:8,border:`1px solid ${b.border}`,background:b.bg,color:b.text,fontSize:13,fontWeight:600,cursor:"pointer",transition:"transform .15s, opacity .15s"}}
              onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-1px)";e.currentTarget.style.opacity=".9";}}
              onMouseLeave={e=>{e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.opacity="1";}}>
              {b.label}
            </button>
          ))}
        </div>
      </div>

      {/* Painel de Aditivos */}
      <div style={{background:"#0c1f3a",border:"1px solid #1e3a5f",borderRadius:12,padding:"16px 20px",marginBottom:16}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <div>
            <div style={{fontSize:12,color:"#93c5fd",fontWeight:700,letterSpacing:.8}}>ADITIVOS DO CONTRATO</div>
            <div style={{fontSize:12,color:"#475569",marginTop:2}}>{aditivos.length} aditivo(s) · Impacto: <span style={{color:totalAditivos>=0?"#4ade80":"#f87171",fontWeight:600}}>{totalAditivos>=0?"+":""}{fmt(totalAditivos)}</span></div>
          </div>
          <button onClick={()=>setModalAditivo(true)} style={{...S.btnGhost,padding:"7px 14px",fontSize:13,borderColor:"#1e3a5f",color:"#93c5fd"}}>+ Novo Aditivo</button>
        </div>
        {aditivos.length===0?(
          <div style={{fontSize:13,color:"#334155",padding:"8px 0"}}>Nenhum aditivo registrado. Clique em "+ Novo Aditivo" para adicionar acréscimos ou supressões ao contrato.</div>
        ):(
          <div style={{display:"flex",flexDirection:"column",gap:6}}>
            {aditivos.map((a,i)=>{
              const acrescimo=a.tipo==="Acréscimo";
              const vencido=a.dataCompensacao&&new Date(a.dataCompensacao+"T00:00:00")<new Date();
              return(
                <div key={a.id} style={{background:"#0d1a2e",borderRadius:8,padding:"10px 14px",display:"flex",alignItems:"center",gap:12,border:`1px solid ${acrescimo?"#166534":"#7f1d1d"}`}}>
                  <div style={{width:6,height:36,borderRadius:3,background:acrescimo?"#4ade80":"#f87171",flexShrink:0}}/>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:13,fontWeight:600,color:"#f1f5f9",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.descricao}</div>
                    <div style={{display:"flex",gap:12,marginTop:3,flexWrap:"wrap"}}>
                      <span style={{fontSize:11,color:"#475569"}}>📅 Assinado: {fmtDate(a.data)}</span>
                      {a.dataCompensacao&&(
                        <span style={{fontSize:11,color:vencido?"#f87171":"#f59e0b",fontWeight:600}}>
                          ⏰ Compensação: {fmtDate(a.dataCompensacao)}{vencido?" · VENCIDO":""}
                        </span>
                      )}
                    </div>
                  </div>
                  <div style={{textAlign:"right",flexShrink:0}}>
                    <span style={{...S.tag,background:acrescimo?"#052e16":"#450a0a",color:acrescimo?"#4ade80":"#f87171",fontSize:11}}>{a.tipo}</span>
                    <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:15,fontWeight:700,color:acrescimo?"#4ade80":"#f87171",marginTop:4}}>{acrescimo?"+":"-"}{fmt(a.valor)}</div>
                  </div>
                  <button onClick={()=>removeAditivo(a.id)} style={{background:"none",border:"none",color:"#475569",cursor:"pointer",fontSize:16,padding:"0 4px",flexShrink:0}}>✕</button>
                </div>
              );
            })}
            {/* Linha de total */}
            <div style={{background:"#111827",borderRadius:8,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:4,border:"1px solid #1e2a3a"}}>
              <span style={{fontSize:12,color:"#64748b",fontWeight:600}}>CONTRATO ORIGINAL: {fmt(obra.valorContrato)}</span>
              <span style={{fontSize:13,fontWeight:700,color:"#f59e0b"}}>CONTRATO ATUALIZADO: {fmt(contratoAtualizado)}</span>
            </div>
          </div>
        )}
      </div>

      {/* m² panel */}
      {obra.metrosQuadrados>0&&(()=>{
        const custoM2=st.despesas>0?(st.despesas/obra.metrosQuadrados):0;
        const contratoM2=contratoAtualizado/obra.metrosQuadrados;
        const saldoM2=contratoM2-custoM2;
        const pctCusto=pct(custoM2,contratoM2);
        return(
          <div style={{background:"#0c1f3a",border:"1px solid #1e3a5f",borderRadius:12,padding:"16px 20px",marginBottom:16}}>
            <div style={{fontSize:12,color:"#64748b",fontWeight:600,letterSpacing:.8,textTransform:"uppercase",marginBottom:12}}>Análise por Metro Quadrado — {obra.metrosQuadrados} m²</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:14}}>
              {[
                {label:"CONTRATO / m²",value:fmt(contratoM2),color:"#f59e0b"},
                {label:"CUSTO ATUAL / m²",value:custoM2>0?fmt(custoM2):"—",color:"#f87171"},
                {label:"MARGEM / m²",value:custoM2>0?fmt(saldoM2):"—",color:saldoM2>=0?"#4ade80":"#f87171"},
                {label:"CUSTO EXECUTADO",value:custoM2>0?`${pctCusto}%`:"—",color:pctCusto>100?"#ef4444":pctCusto>80?"#f59e0b":"#93c5fd"},
              ].map(item=>{
                const len=(item.value||"").toString().length;
                const fs=len>14?14:len>11?16:len>9?18:20;
                return(
                  <div key={item.label} style={{minWidth:0}}>
                    <div style={{fontSize:10,color:"#475569",marginBottom:4,fontWeight:600,letterSpacing:.8,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{item.label}</div>
                    <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:fs,fontWeight:700,color:item.color,wordBreak:"break-word",lineHeight:1.2,letterSpacing:0}}>{item.value}</div>
                  </div>
                );
              })}
            </div>
            <div style={{background:"#0d1a2e",borderRadius:6,height:8,overflow:"hidden"}}>
              <div style={{height:"100%",background:pctCusto>100?"#ef4444":pctCusto>80?"#f59e0b":"#10b981",width:`${Math.min(100,pctCusto)}%`,borderRadius:6,transition:"width .8s"}}/>
            </div>
            <div style={{fontSize:11,color:"#475569",marginTop:6}}>
              Custo atual: {fmt(st.despesas||0)} de {fmt(contratoAtualizado)} do contrato ({pctCusto}% do valor/m² consumido)
            </div>
          </div>
        );
      })()}

      {lucro<0&&<Alert type="danger" msg={`PREJUÍZO: Esta obra está com deficit de ${fmt(Math.abs(lucro))} em relação ao contrato atualizado.`}/>}
      {pct(st.despesas||0,contratoAtualizado)>100&&<Alert type="warn" msg="ORÇAMENTO ESTOURADO: As despesas ultrapassaram o valor do contrato atualizado."/>}

      {/* PAINEL DE SÓCIOS — só para Construção e Venda */}
      {obra.modalidade==="Construção e Venda"&&obra.socios&&obra.socios.length>0&&(()=>{
        const totalDespesas=st.despesas||0;
        const totalReceitas=st.recebido||0;
        const valorVenda=obra.valorVendaReal||obra.valorVendaEstimado||0;
        const lucroVenda=valorVenda-totalDespesas;
        const acertos=obra.socios.map(s=>{
          const pago=despesas.filter(l=>l.pagoPor===s.id).reduce((sum,l)=>sum+l.valor,0);
          const recebido=receitas.filter(l=>l.pagoPor===s.id).reduce((sum,l)=>sum+l.valor,0);
          const deveriaPagar=totalDespesas*s.percentual/100;
          const saldoAcerto=pago-deveriaPagar;
          const cotaLucro=lucroVenda*s.percentual/100;
          return{...s,pago,recebido,deveriaPagar,saldoAcerto,cotaLucro};
        });
        const naoAtribuido=despesas.filter(l=>!l.pagoPor||!obra.socios.find(s=>s.id===l.pagoPor)).reduce((sum,l)=>sum+l.valor,0);
        const valorVendaTotal=obra.valorVendaReal>0?"real":"estimado";
        return(
          <div style={{background:"#1a0f2e",border:"1px solid #5b21b6",borderRadius:12,padding:"16px 20px",marginBottom:16}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:8}}>
              <div>
                <div style={{fontSize:12,color:"#c4b5fd",fontWeight:700,letterSpacing:.8}}>ACERTO DE CONTAS — SÓCIOS DO EMPREENDIMENTO</div>
                <div style={{fontSize:11,color:"#64748b",marginTop:2}}>
                  Valor de venda ({valorVendaTotal}): <span style={{color:"#f59e0b",fontWeight:700}}>{fmt(valorVenda)}</span>
                  {valorVenda>0&&<> · Lucro projetado: <span style={{color:lucroVenda>=0?"#4ade80":"#f87171",fontWeight:700}}>{fmt(lucroVenda)}</span></>}
                </div>
              </div>
            </div>

            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead>
                  <tr style={{borderBottom:"1px solid #4c1d95"}}>
                    {["Sócio","%","Deveria Pagar","Pagou","Acerto","Cota do Lucro"].map(h=>(
                      <th key={h} style={{textAlign:"left",padding:"8px 12px",color:"#a78bfa",fontWeight:600,fontSize:11,letterSpacing:.5}}>{h.toUpperCase()}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {acertos.map((s,i)=>(
                    <tr key={s.id} style={{borderBottom:"1px solid #2d1b69"}}>
                      <td style={{padding:"10px 12px"}}>
                        <div style={{display:"flex",alignItems:"center",gap:8}}>
                          <div style={{width:10,height:10,borderRadius:5,background:SOCIO_COLORS[i%SOCIO_COLORS.length]}}/>
                          <span style={{color:"#f1f5f9",fontWeight:600}}>{s.nome}</span>
                        </div>
                      </td>
                      <td style={{padding:"10px 12px",color:"#a78bfa",fontWeight:600}}>{s.percentual}%</td>
                      <td style={{padding:"10px 12px",color:"#94a3b8"}}>{fmt(s.deveriaPagar)}</td>
                      <td style={{padding:"10px 12px",color:"#e2e8f0",fontWeight:600}}>{fmt(s.pago)}</td>
                      <td style={{padding:"10px 12px",fontWeight:700,color:s.saldoAcerto>0?"#4ade80":s.saldoAcerto<0?"#f87171":"#64748b"}}>
                        {s.saldoAcerto>0?"+":""}{fmt(s.saldoAcerto)}
                        <div style={{fontSize:10,color:"#475569",fontWeight:400,marginTop:1}}>
                          {Math.abs(s.saldoAcerto)<1?"Em dia":s.saldoAcerto>0?"A receber dos sócios":"A pagar aos sócios"}
                        </div>
                      </td>
                      <td style={{padding:"10px 12px",fontWeight:700,color:s.cotaLucro>=0?"#4ade80":"#f87171"}}>{valorVenda>0?fmt(s.cotaLucro):"—"}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr style={{borderTop:"2px solid #4c1d95"}}>
                    <td style={{padding:"10px 12px",color:"#c4b5fd",fontWeight:700,fontSize:12}}>TOTAL</td>
                    <td style={{padding:"10px 12px",color:"#a78bfa",fontWeight:700}}>{obra.socios.reduce((s,x)=>s+x.percentual,0)}%</td>
                    <td style={{padding:"10px 12px",color:"#94a3b8",fontWeight:700}}>{fmt(totalDespesas)}</td>
                    <td style={{padding:"10px 12px",color:"#e2e8f0",fontWeight:700}}>{fmt(acertos.reduce((s,x)=>s+x.pago,0))}</td>
                    <td/>
                    <td style={{padding:"10px 12px",color:lucroVenda>=0?"#4ade80":"#f87171",fontWeight:700}}>{valorVenda>0?fmt(lucroVenda):"—"}</td>
                  </tr>
                </tfoot>
              </table>
            </div>
            {naoAtribuido>0&&(
              <div style={{marginTop:12,background:"#451a03",border:"1px solid #92400e",borderRadius:6,padding:"8px 12px",fontSize:12,color:"#fcd34d"}}>
                ⚠ <strong>{fmt(naoAtribuido)}</strong> em despesas ainda não foram atribuídas a nenhum sócio. Edite os lançamentos para incluir o campo "Pago Por".
              </div>
            )}
          </div>
        );
      })()}

      <div style={{...S.grid4,marginBottom:20}}>
        <KpiCard label="Total Recebido" value={fmt(st.recebido||0)} icon="💰" color="#4ade80" sub={`${receitas.length} entrada(s)`}/>
        <KpiCard label="Total Gasto" value={fmt(st.despesas||0)} icon="💸" color="#f87171" sub={`${despesas.length} saída(s)`}/>
        <KpiCard label="Saldo Atual" value={fmt(st.saldo||0)} color={(st.saldo||0)>=0?"#4ade80":"#f87171"} icon="⚖" sub="Recebido - Despesas"/>
        <KpiCard label="Lucro Estimado" value={fmt(lucro)} color={lucro>=0?"#4ade80":"#f87171"} icon="📈" sub="Contrato atual - Gastos"/>
      </div>

      <div style={{...S.grid2,marginBottom:20}}>
        <div style={S.card}>
          <h3 style={{...S.h3,marginBottom:4}}>Execução Orçamentária</h3>
          <div style={{fontSize:13,color:"#64748b",marginBottom:16}}>{fmt(st.despesas||0)} de {fmt(contratoAtualizado)}</div>
          <div style={{background:"#1e2a3a",borderRadius:6,height:12,overflow:"hidden",marginBottom:8}}>
            <div style={{height:"100%",background:pct(st.despesas||0,contratoAtualizado)>100?"#ef4444":"#f59e0b",width:`${Math.min(100,pct(st.despesas||0,contratoAtualizado))}%`,borderRadius:6,transition:"width .8s"}}/>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:13}}>
            <span style={{color:"#64748b"}}>Executado</span>
            <span style={{color:"#f59e0b",fontWeight:700}}>{pct(st.despesas||0,contratoAtualizado)}%</span>
          </div>
          {obra.dataInicio&&<div style={{marginTop:16,fontSize:12,color:"#475569"}}>Início: {fmtDate(obra.dataInicio)} · Previsão: {fmtDate(obra.dataFim)}</div>}
        </div>
        <div style={S.card}>
          <h3 style={{...S.h3,marginBottom:4}}>Gastos por Categoria</h3>
          <div style={{fontSize:13,color:"#64748b",marginBottom:12}}>Distribuição das despesas</div>
          {byCategoria.length===0?<div style={{color:"#334155",fontSize:13,marginTop:20}}>Sem despesas registradas</div>:(
            <>
              <ResponsiveContainer width="100%" height={140}>
                <PieChart>
                  <Pie data={byCategoria} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={60} innerRadius={30}>
                    {byCategoria.map((_,i)=><Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>)}
                  </Pie>
                  <Tooltip formatter={v=>fmt(v)} contentStyle={{background:"#0d1220",border:"1px solid #1e2a3a",borderRadius:8,color:"#e2e8f0"}}/>
                </PieChart>
              </ResponsiveContainer>
              <div style={{display:"flex",flexWrap:"wrap",gap:8,marginTop:8}}>
                {byCategoria.map((c,i)=>(
                  <div key={c.name} style={{display:"flex",alignItems:"center",gap:4,fontSize:11,color:"#94a3b8"}}>
                    <div style={{width:8,height:8,borderRadius:2,background:PIE_COLORS[i%PIE_COLORS.length]}}/>
                    {c.name}
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
      </div>

      {byMonth.length>0&&(
        <div style={{...S.card,marginBottom:20}}>
          <h3 style={{...S.h3,marginBottom:16}}>Evolução Mensal — Receitas vs Despesas</h3>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={byMonth} barGap={4}>
              <XAxis dataKey="mes" tick={{fill:"#64748b",fontSize:12}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fill:"#64748b",fontSize:11}} axisLine={false} tickLine={false} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`}/>
              <Tooltip formatter={v=>fmt(v)} contentStyle={{background:"#0d1220",border:"1px solid #1e2a3a",borderRadius:8,color:"#e2e8f0"}}/>
              <Legend wrapperStyle={{fontSize:12,color:"#64748b"}}/>
              <Bar dataKey="receita" fill="#10b981" name="Receita" radius={[4,4,0,0]}/>
              <Bar dataKey="despesa" fill="#f59e0b" name="Despesa" radius={[4,4,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </div>
      )}

      <RelatorioObraSection obra={obra} lancamentos={lancamentos} obraStats={{[obra.id]:st}}/>

      {/* Modal Aditivo */}
      <Modal open={modalAditivo} onClose={()=>setModalAditivo(false)} title="Novo Aditivo de Contrato">
        <div style={{background:"#0c1f3a",border:"1px solid #1e3a5f",borderRadius:8,padding:"10px 14px",marginBottom:16,fontSize:13,color:"#93c5fd"}}>
          ℹ Aditivos atualizam o valor total do contrato sem gerar lançamento financeiro. Use para registrar acréscimos ou supressões formais.
        </div>
        <FormField label="Tipo de Aditivo">
          <select style={S.input} value={novoAditivo.tipo} onChange={e=>setNovoAditivo(p=>({...p,tipo:e.target.value}))}>
            <option>Acréscimo</option>
            <option>Supressão</option>
          </select>
        </FormField>
        <FormField label="Descrição *">
          <input style={S.input} value={novoAditivo.descricao} onChange={e=>setNovoAditivo(p=>({...p,descricao:e.target.value}))} placeholder="Ex: Aditivo 01 — Ampliação de área"/>
        </FormField>
        <div style={S.grid2}>
          <FormField label="Valor (R$) *">
            <input style={S.input} type="number" step="0.01" value={novoAditivo.valor} onChange={e=>setNovoAditivo(p=>({...p,valor:e.target.value}))} placeholder="0,00"/>
          </FormField>
          <FormField label="Data de Assinatura">
            <input style={S.input} type="date" value={novoAditivo.data} onChange={e=>setNovoAditivo(p=>({...p,data:e.target.value}))}/>
          </FormField>
        </div>
        <FormField label="Data de Compensação">
          <input style={S.input} type="date" value={novoAditivo.dataCompensacao} onChange={e=>setNovoAditivo(p=>({...p,dataCompensacao:e.target.value}))}/>
          <div style={{fontSize:11,color:"#475569",marginTop:4}}>Data limite para o valor do aditivo ser compensado/pago. Exibirá alerta quando vencida.</div>
        </FormField>
        {novoAditivo.valor>0&&(
          <div style={{background:"#111827",borderRadius:8,padding:"10px 14px",marginBottom:8,fontSize:13}}>
            <span style={{color:"#64748b"}}>Novo contrato: </span>
            <span style={{color:"#f59e0b",fontWeight:700}}>{fmt(contratoAtualizado+(novoAditivo.tipo==="Acréscimo"?parseFloat(novoAditivo.valor||0):-parseFloat(novoAditivo.valor||0)))}</span>
          </div>
        )}
        <div style={{display:"flex",gap:12,justifyContent:"flex-end",marginTop:20}}>
          <button style={S.btnGhost} onClick={()=>setModalAditivo(false)}>Cancelar</button>
          <button style={S.btnPrimary} onClick={saveAditivo}>Registrar Aditivo</button>
        </div>
      </Modal>

      {/* Modal Lançamento Rápido */}
      <Modal open={modalLanc} onClose={()=>{setModalLanc(false);setDefaultCatLanc(null);setDefaultTipoLanc(null);}} title={`Novo ${defaultTipoLanc||"Lançamento"} — ${obra.nome}`}>
        <LancamentoForm 
          obras={allObras||[obra]} 
          defaultObraId={obra.id}
          defaultCategoria={defaultCatLanc}
          defaultTipo={defaultTipoLanc}
          onSave={l=>{addLancamento(l);setModalLanc(false);setDefaultCatLanc(null);setDefaultTipoLanc(null);}} 
          onClose={()=>{setModalLanc(false);setDefaultCatLanc(null);setDefaultTipoLanc(null);}}
        />
      </Modal>
    </div>
  );
}

// ═══════════════════════════════════════════════════════
// LANCAMENTO FORM
// ═══════════════════════════════════════════════════════
function LancamentoForm({obras,onSave,onClose,defaultObraId,defaultTipo,defaultCategoria}){
  const [f,setF]=useState({obraId:defaultObraId||obras[0]?.id||"",tipo:defaultTipo||(defaultCategoria?"Despesa":"Despesa"),categoria:defaultCategoria||"Material",subcategoria:"",descricao:"",valor:"",data:today(),pagamento:"PIX",responsavel:"",pagoPor:""});
  const set=k=>e=>setF(p=>({...p,[k]:e.target.value}));
  const obraAtual=obras.find(o=>o.id===f.obraId);
  const temSocios=obraAtual?.modalidade==="Construção e Venda"&&obraAtual?.socios?.length>0;
  const save=()=>{
    if(!f.obraId)return alert("Selecione uma obra");
    if(!f.valor||parseFloat(f.valor)<=0)return alert("Valor inválido");
    onSave({...f,valor:parseFloat(f.valor),id:uid()});
  };
  return(
    <>
      <div style={S.grid2}>
        <FormField label="Obra *">
          <select style={S.input} value={f.obraId} onChange={set("obraId")}>
            <option value="">Selecione...</option>
            {obras.map(o=><option key={o.id} value={o.id}>{o.nome}</option>)}
          </select>
        </FormField>
        <FormField label="Tipo *">
          <select style={S.input} value={f.tipo} onChange={set("tipo")}>
            <option>Receita</option><option>Despesa</option>
          </select>
        </FormField>
      </div>
      {f.tipo==="Despesa"&&(
        <div style={S.grid2}>
          <FormField label="Categoria">
            <select style={S.input} value={f.categoria} onChange={e=>{setF(p=>({...p,categoria:e.target.value,subcategoria:""}));}}>
              {CATS.map(c=><option key={c}>{c}</option>)}
            </select>
          </FormField>
          <FormField label="Subcategoria">
            <select style={S.input} value={f.subcategoria} onChange={set("subcategoria")}>
              <option value="">Selecione...</option>
              {(SUBCATS[f.categoria]||[]).map(s=><option key={s}>{s}</option>)}
            </select>
          </FormField>
        </div>
      )}
      {f.tipo==="Receita"&&(
        <FormField label="Descrição">
          <select style={S.input} value={f.descricao} onChange={set("descricao")}>
            <option value="">Tipo de receita...</option>
            {["Medição","Adiantamento","Parcela","Sinal","Pagamento final","Outro"].map(s=><option key={s}>{s}</option>)}
          </select>
        </FormField>
      )}
      <FormField label="Descrição / Observação"><input style={S.input} value={f.tipo==="Despesa"?f.descricao:undefined||f.descricao} onChange={set("descricao")} placeholder="Detalhes do lançamento"/></FormField>
      <div style={S.grid2}>
        <FormField label="Valor (R$) *"><input style={S.input} type="number" step="0.01" value={f.valor} onChange={set("valor")} placeholder="0,00"/></FormField>
        <FormField label="Data *"><input style={S.input} type="date" value={f.data} onChange={set("data")}/></FormField>
      </div>
      <div style={S.grid2}>
        <FormField label="Forma de Pagamento">
          <select style={S.input} value={f.pagamento} onChange={set("pagamento")}>{PAGAMENTOS.map(p=><option key={p}>{p}</option>)}</select>
        </FormField>
        <FormField label="Responsável"><input style={S.input} value={f.responsavel} onChange={set("responsavel")} placeholder="Nome"/></FormField>
      </div>
      {temSocios&&(
        <div style={{background:"#0c1f3a",border:"1px solid #1e3a5f",borderRadius:8,padding:"10px 14px",marginBottom:14}}>
          <FormField label="Pago Por (Sócio)">
            <select style={S.input} value={f.pagoPor} onChange={set("pagoPor")}>
              <option value="">— Conta da obra / Não atribuído —</option>
              {obraAtual.socios.map(s=><option key={s.id} value={s.id}>{s.nome} ({s.percentual}%)</option>)}
            </select>
            <div style={{fontSize:11,color:"#93c5fd",marginTop:4}}>
              ℹ Define qual sócio efetuou o pagamento. Essencial para o acerto de contas final.
            </div>
          </FormField>
        </div>
      )}
      <div style={{display:"flex",gap:12,justifyContent:"flex-end",marginTop:20}}>
        <button style={S.btnGhost} onClick={onClose}>Cancelar</button>
        <button style={S.btnPrimary} onClick={save}>Registrar {f.tipo}</button>
      </div>
    </>
  );
}

// ═══════════════════════════════════════════════════════
// FINANCEIRO VIEW — UNIFICADO
// ═══════════════════════════════════════════════════════
function FinanceiroView({obras,lancamentos,addLancamento,deleteLancamento}){
  const [modal,setModal]=useState(false);
  const [defaultCat,setDefaultCat]=useState(null);
  const [filterObra,setFilterObra]=useState("");
  const [filterTipo,setFilterTipo]=useState("");
  const [filterCat,setFilterCat]=useState("");
  const [search,setSearch]=useState("");

  const filtered=lancamentos.filter(l=>{
    if(filterObra&&l.obraId!==filterObra)return false;
    if(filterTipo&&l.tipo!==filterTipo)return false;
    if(filterCat&&l.categoria!==filterCat)return false;
    if(search&&!`${l.descricao}${l.categoria}${l.subcategoria}${l.responsavel}`.toLowerCase().includes(search.toLowerCase()))return false;
    return true;
  }).sort((a,b)=>b.data>a.data?1:-1);

  const totalRec=filtered.filter(l=>l.tipo==="Receita").reduce((s,l)=>s+l.valor,0);
  const totalDesp=filtered.filter(l=>l.tipo==="Despesa").reduce((s,l)=>s+l.valor,0);
  const obraMap=Object.fromEntries(obras.map(o=>[o.id,o.nome]));

  // Totais por categoria (filtro aplicado)
  const catTotals=CATS.map(c=>({cat:c,val:filtered.filter(l=>l.tipo==="Despesa"&&l.categoria===c).reduce((s,l)=>s+l.valor,0)}));

  const abrirLancamento=(cat=null)=>{setDefaultCat(cat);setModal(true);};

  return(
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20}}>
        <div><h1 style={S.h1}>Financeiro</h1><p style={S.muted}>Central de lançamentos — receitas, despesas e materiais</p></div>
        <button style={S.btnPrimary} onClick={()=>abrirLancamento(null)}>+ Novo Lançamento</button>
      </div>

      {/* KPIs principais */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:12,marginBottom:20}}>
        <KpiCard label="Total Receitas" value={fmt(totalRec)} color="#4ade80" icon="↑" sub={`${filtered.filter(l=>l.tipo==="Receita").length} entradas`}/>
        <KpiCard label="Total Despesas" value={fmt(totalDesp)} color="#f87171" icon="↓" sub={`${filtered.filter(l=>l.tipo==="Despesa").length} saídas`}/>
        <KpiCard label="Resultado" value={fmt(totalRec-totalDesp)} color={(totalRec-totalDesp)>=0?"#4ade80":"#f87171"} icon="="/>
        {catTotals.filter(c=>c.cat==="Material").map(c=>(
          <KpiCard key={c.cat} label="Materiais" value={fmt(c.val)} color="#f59e0b" icon="🧱" sub={`${filtered.filter(l=>l.categoria==="Material").length} reg.`}/>
        ))}
        {catTotals.filter(c=>c.cat==="Mão de Obra").map(c=>(
          <KpiCard key={c.cat} label="Mão de Obra" value={fmt(c.val)} color="#8b5cf6" icon="👷" sub={`${filtered.filter(l=>l.categoria==="Mão de Obra").length} reg.`}/>
        ))}
        {catTotals.filter(c=>c.cat==="Comissão").map(c=>(
          <KpiCard key={c.cat} label="Comissões" value={fmt(c.val)} color="#fcd34d" icon="💼" sub={`${filtered.filter(l=>l.categoria==="Comissão").length} reg.`}/>
        ))}
      </div>

      {/* Botões rápidos por categoria */}
      <div style={{...S.card,marginBottom:16,padding:"14px 20px"}}>
        <div style={{fontSize:12,color:"#64748b",marginBottom:10,fontWeight:600,letterSpacing:.8}}>LANÇAMENTO RÁPIDO POR CATEGORIA</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {[
            {cat:"Material",label:"🧱 Material",bg:"#422006",border:"#7c2d12",text:"#fdba74"},
            {cat:"Mão de Obra",label:"👷 Mão de Obra",bg:"#2d1b69",border:"#4c1d95",text:"#c4b5fd"},
            {cat:"Equipamento",label:"🔧 Equipamento",bg:"#0c1f3a",border:"#1e3a5f",text:"#93c5fd"},
            {cat:"Comissão",label:"💼 Comissão",bg:"#451a03",border:"#92400e",text:"#fcd34d"},
            {cat:"Outros",label:"📦 Outros",bg:"#1a2e1a",border:"#166534",text:"#86efac"},
            {cat:null,label:"💰 Receita",bg:"#052e16",border:"#166534",text:"#4ade80",tipo:"Receita"},
          ].map(b=>(
            <button key={b.label} onClick={()=>abrirLancamento(b.cat)}
              style={{padding:"8px 16px",borderRadius:8,border:`1px solid ${b.border}`,background:b.bg,color:b.text,fontSize:13,fontWeight:600,cursor:"pointer",transition:"opacity .15s"}}
              onMouseEnter={e=>e.currentTarget.style.opacity=".8"} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
              {b.label}
            </button>
          ))}
        </div>
      </div>

      {/* Filtros */}
      <div style={{...S.card,marginBottom:16}}>
        <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
          <input style={{...S.input,flex:1,minWidth:180}} placeholder="Buscar descrição, categoria, responsável..." value={search} onChange={e=>setSearch(e.target.value)}/>
          <select style={{...S.input,width:190}} value={filterObra} onChange={e=>setFilterObra(e.target.value)}>
            <option value="">Todas as obras</option>
            {obras.map(o=><option key={o.id} value={o.id}>{o.nome}</option>)}
          </select>
          <select style={{...S.input,width:140}} value={filterTipo} onChange={e=>setFilterTipo(e.target.value)}>
            <option value="">Tipo</option>
            <option>Receita</option><option>Despesa</option>
          </select>
          <select style={{...S.input,width:160}} value={filterCat} onChange={e=>setFilterCat(e.target.value)}>
            <option value="">Categoria</option>
            {CATS.map(c=><option key={c}>{c}</option>)}
          </select>
          {(filterObra||filterTipo||filterCat||search)&&(
            <button style={{...S.btnGhost,padding:"8px 12px",fontSize:12}} onClick={()=>{setFilterObra("");setFilterTipo("");setFilterCat("");setSearch("");}}>✕ Limpar</button>
          )}
        </div>
        {/* Totais por categoria visíveis */}
        <div style={{display:"flex",gap:16,marginTop:12,flexWrap:"wrap"}}>
          {catTotals.filter(c=>c.val>0).map((c,i)=>(
            <div key={c.cat} style={{display:"flex",alignItems:"center",gap:6,fontSize:12}}>
              <div style={{width:8,height:8,borderRadius:2,background:PIE_COLORS[i%PIE_COLORS.length]}}/>
              <span style={{color:"#64748b"}}>{c.cat}:</span>
              <span style={{color:"#f1f5f9",fontWeight:600}}>{fmt(c.val)}</span>
            </div>
          ))}
        </div>
      </div>

      {/* Tabela */}
      <div style={S.card}>
        {filtered.length===0
          ?<div style={{textAlign:"center",padding:40,color:"#334155"}}>Nenhum lançamento encontrado</div>
          :(
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{borderBottom:"1px solid #1e2a3a"}}>
                  {["Data","Obra","Tipo","Categoria","Subcategoria","Descrição","Responsável","Pgto","Valor",""].map(h=>(
                    <th key={h} style={{textAlign:"left",padding:"8px 12px",color:"#475569",fontWeight:500,whiteSpace:"nowrap"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.map(l=>(
                  <tr key={l.id} style={{borderBottom:"1px solid #0d1220",transition:"background .1s"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#111827"}
                    onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                    <td style={{padding:"10px 12px",color:"#64748b",whiteSpace:"nowrap"}}>{fmtDate(l.data)}</td>
                    <td style={{padding:"10px 12px",color:"#94a3b8",maxWidth:140}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{obraMap[l.obraId]||"—"}</div></td>
                    <td style={{padding:"10px 12px"}}><span style={{...S.tag,background:l.tipo==="Receita"?"#052e16":"#450a0a",color:l.tipo==="Receita"?"#4ade80":"#f87171"}}>{l.tipo}</span></td>
                    <td style={{padding:"10px 12px",color:"#94a3b8"}}>{l.categoria||"—"}</td>
                    <td style={{padding:"10px 12px",color:"#64748b"}}>{l.subcategoria||"—"}</td>
                    <td style={{padding:"10px 12px",color:"#e2e8f0",maxWidth:200}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{l.descricao||"—"}</div></td>
                    <td style={{padding:"10px 12px",color:"#64748b"}}>{l.responsavel||"—"}</td>
                    <td style={{padding:"10px 12px",color:"#475569",whiteSpace:"nowrap"}}>{l.pagamento||"—"}</td>
                    <td style={{padding:"10px 12px",fontWeight:700,color:l.tipo==="Receita"?"#4ade80":"#f87171",whiteSpace:"nowrap"}}>{l.tipo==="Receita"?"+":"-"}{fmt(l.valor)}</td>
                    <td style={{padding:"10px 12px"}}><button onClick={()=>deleteLancamento(l.id)} style={{background:"none",border:"none",color:"#475569",cursor:"pointer",fontSize:15,padding:2}}>✕</button></td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr style={{borderTop:"2px solid #1e2a3a"}}>
                  <td colSpan={8} style={{padding:"10px 12px",color:"#64748b",fontSize:12}}>{filtered.length} lançamento(s)</td>
                  <td style={{padding:"10px 12px",fontWeight:700,color:(totalRec-totalDesp)>=0?"#4ade80":"#f87171",whiteSpace:"nowrap"}}>{fmt(totalRec-totalDesp)}</td>
                  <td/>
                </tr>
              </tfoot>
            </table>
          </div>
        )}
      </div>

      <Modal open={modal} onClose={()=>{setModal(false);setDefaultCat(null);}} title="Novo Lançamento">
        <LancamentoForm obras={obras} onSave={l=>{addLancamento(l);setModal(false);setDefaultCat(null);}} onClose={()=>{setModal(false);setDefaultCat(null);}} defaultCategoria={defaultCat}/>
      </Modal>
    </div>
  );
}

// ═══════════════════════════════════════════════════════
// MÃO DE OBRA VIEW
// ═══════════════════════════════════════════════════════
function MaoDeObraView({obras,equipes,setEquipes,addLancamento,lancamentos}){
  const [modalEquipe,setModalEquipe]=useState(false);
  const [modalLanc,setModalLanc]=useState(false);
  const [novaEquipe,setNovaEquipe]=useState({nome:"",tipo:"Alvenaria",responsavel:""});
  const [lanc,setLanc]=useState({obraId:obras[0]?.id||"",equipeId:"",periodo:"",valor:"",obs:"",data:today()});

  const mdoLancs=lancamentos.filter(l=>l.categoria==="Mão de Obra").sort((a,b)=>b.data>a.data?1:-1);
  const obraMap=Object.fromEntries(obras.map(o=>[o.id,o.nome]));
  const eqMap=Object.fromEntries(equipes.map(e=>[e.id,e]));

  const saveEquipe=()=>{
    if(!novaEquipe.nome)return alert("Nome obrigatório");
    setEquipes(p=>[...p,{...novaEquipe,id:uid()}]);
    setNovaEquipe({nome:"",tipo:"Alvenaria",responsavel:""});
    setModalEquipe(false);
  };
  const saveLanc=()=>{
    if(!lanc.obraId)return alert("Selecione uma obra");
    if(!lanc.valor||parseFloat(lanc.valor)<=0)return alert("Valor inválido");
    const eq=eqMap[lanc.equipeId];
    addLancamento({obraId:lanc.obraId,tipo:"Despesa",categoria:"Mão de Obra",subcategoria:eq?.tipo||"",descricao:`${eq?.nome||"Equipe"} · ${lanc.periodo}`,valor:parseFloat(lanc.valor),data:lanc.data,pagamento:"PIX",responsavel:eq?.responsavel||lanc.responsavel||""});
    setLanc({obraId:obras[0]?.id||"",equipeId:"",periodo:"",valor:"",obs:"",data:today()});
    setModalLanc(false);
  };

  return(
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:24}}>
        <div><h1 style={S.h1}>Mão de Obra</h1><p style={S.muted}>Controle de equipes e pagamentos</p></div>
        <div style={{display:"flex",gap:8}}>
          <button style={S.btnGhost} onClick={()=>setModalEquipe(true)}>+ Cadastrar Equipe</button>
          <button style={S.btnPrimary} onClick={()=>setModalLanc(true)}>+ Lançar Mão de Obra</button>
        </div>
      </div>

      <div style={{marginBottom:20}}>
        <h2 style={{...S.h2,marginBottom:12}}>Equipes Cadastradas</h2>
        {equipes.length===0?<div style={{...S.card,color:"#334155",textAlign:"center",padding:32}}>Nenhuma equipe cadastrada</div>:(
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:12}}>
            {equipes.map(e=>(
              <div key={e.id} style={{...S.cardSm,display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div>
                  <div style={{fontWeight:700,color:"#f1f5f9",fontSize:15,marginBottom:4}}>{e.nome}</div>
                  <div style={{fontSize:12,color:"#64748b"}}>{e.tipo}</div>
                  {e.responsavel&&<div style={{fontSize:12,color:"#475569",marginTop:4}}>👤 {e.responsavel}</div>}
                </div>
                <button onClick={()=>setEquipes(p=>p.filter(x=>x.id!==e.id))} style={{background:"none",border:"none",color:"#475569",cursor:"pointer",fontSize:15}}>✕</button>
              </div>
            ))}
          </div>
        )}
      </div>

      <div>
        <h2 style={{...S.h2,marginBottom:12}}>Lançamentos de Mão de Obra</h2>
        <div style={S.card}>
          {mdoLancs.length===0?<div style={{color:"#334155",textAlign:"center",padding:32}}>Nenhum lançamento de mão de obra registrado</div>:(
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr style={{borderBottom:"1px solid #1e2a3a"}}>
                {["Data","Obra","Equipe/Descrição","Período","Responsável","Valor"].map(h=><th key={h} style={{textAlign:"left",padding:"8px 12px",color:"#475569",fontWeight:500}}>{h}</th>)}
              </tr></thead>
              <tbody>{mdoLancs.map(l=>(
                <tr key={l.id} style={{borderBottom:"1px solid #111827"}}>
                  <td style={{padding:"10px 12px",color:"#64748b"}}>{fmtDate(l.data)}</td>
                  <td style={{padding:"10px 12px",color:"#94a3b8"}}>{obraMap[l.obraId]||"—"}</td>
                  <td style={{padding:"10px 12px",color:"#e2e8f0"}}>{l.descricao||"—"}</td>
                  <td style={{padding:"10px 12px",color:"#64748b"}}>{l.subcategoria||"—"}</td>
                  <td style={{padding:"10px 12px",color:"#64748b"}}>{l.responsavel||"—"}</td>
                  <td style={{padding:"10px 12px",fontWeight:700,color:"#f87171"}}>-{fmt(l.valor)}</td>
                </tr>
              ))}</tbody>
            </table>
          )}
        </div>
      </div>

      <Modal open={modalEquipe} onClose={()=>setModalEquipe(false)} title="Nova Equipe">
        <FormField label="Nome da Equipe *"><input style={S.input} value={novaEquipe.nome} onChange={e=>setNovaEquipe(p=>({...p,nome:e.target.value}))} placeholder="Ex: Equipe de Pedreiros"/></FormField>
        <FormField label="Tipo">
          <select style={S.input} value={novaEquipe.tipo} onChange={e=>setNovaEquipe(p=>({...p,tipo:e.target.value}))}>{TIPOS_EQUIPE.map(t=><option key={t}>{t}</option>)}</select>
        </FormField>
        <FormField label="Responsável"><input style={S.input} value={novaEquipe.responsavel} onChange={e=>setNovaEquipe(p=>({...p,responsavel:e.target.value}))} placeholder="Nome do encarregado"/></FormField>
        <div style={{display:"flex",gap:12,justifyContent:"flex-end",marginTop:20}}>
          <button style={S.btnGhost} onClick={()=>setModalEquipe(false)}>Cancelar</button>
          <button style={S.btnPrimary} onClick={saveEquipe}>Cadastrar Equipe</button>
        </div>
      </Modal>

      <Modal open={modalLanc} onClose={()=>setModalLanc(false)} title="Lançar Mão de Obra">
        <div style={{background:"#052e16",border:"1px solid #166534",borderRadius:8,padding:"10px 14px",marginBottom:16,fontSize:13,color:"#4ade80"}}>
          ✓ O valor será automaticamente lançado como DESPESA e irá abater do saldo da obra.
        </div>
        <div style={S.grid2}>
          <FormField label="Obra *">
            <select style={S.input} value={lanc.obraId} onChange={e=>setLanc(p=>({...p,obraId:e.target.value}))}>
              <option value="">Selecione...</option>
              {obras.map(o=><option key={o.id} value={o.id}>{o.nome}</option>)}
            </select>
          </FormField>
          <FormField label="Equipe">
            <select style={S.input} value={lanc.equipeId} onChange={e=>setLanc(p=>({...p,equipeId:e.target.value}))}>
              <option value="">Selecione...</option>
              {equipes.map(e=><option key={e.id} value={e.id}>{e.nome}</option>)}
            </select>
          </FormField>
        </div>
        <div style={S.grid2}>
          <FormField label="Período (ex: Semana 01/01-07/01)"><input style={S.input} value={lanc.periodo} onChange={e=>setLanc(p=>({...p,periodo:e.target.value}))} placeholder="Semana ou período"/></FormField>
          <FormField label="Data do Pagamento"><input style={S.input} type="date" value={lanc.data} onChange={e=>setLanc(p=>({...p,data:e.target.value}))}/></FormField>
        </div>
        <FormField label="Valor (R$) *"><input style={S.input} type="number" step="0.01" value={lanc.valor} onChange={e=>setLanc(p=>({...p,valor:e.target.value}))} placeholder="0,00"/></FormField>
        <div style={{display:"flex",gap:12,justifyContent:"flex-end",marginTop:20}}>
          <button style={S.btnGhost} onClick={()=>setModalLanc(false)}>Cancelar</button>
          <button style={S.btnPrimary} onClick={saveLanc}>Lançar e Abater do Saldo</button>
        </div>
      </Modal>
    </div>
  );
}

// ═══════════════════════════════════════════════════════
// DASHBOARD GERAL
// ═══════════════════════════════════════════════════════
function DashboardView({obras,obraStats,lancamentos,setView,setSelectedObra,goToObra}){
  const [filtroObra,setFiltroObra]=useState("");

  const lans=filtroObra?lancamentos.filter(l=>l.obraId===filtroObra):lancamentos;
  const obrasVisiveis=filtroObra?obras.filter(o=>o.id===filtroObra):obras;

  const totalContrato=obrasVisiveis.reduce((s,o)=>s+o.valorContrato,0);
  const totalRec=obrasVisiveis.reduce((s,o)=>s+(obraStats[o.id]?.recebido||0),0);
  const totalDesp=obrasVisiveis.reduce((s,o)=>s+(obraStats[o.id]?.despesas||0),0);
  const totalLucro=obrasVisiveis.reduce((s,o)=>s+(o.valorContrato-(obraStats[o.id]?.despesas||0)),0);
  const saldoAtual=totalRec-totalDesp;

  const alerts=obrasVisiveis.filter(o=>{
    const st=obraStats[o.id]||{};
    return(o.valorContrato-(st.despesas||0))<0||(st.despesas||0)>o.valorContrato;
  });

  const ranking=[...obrasVisiveis].sort((a,b)=>{
    const la=a.valorContrato-(obraStats[a.id]?.despesas||0);
    const lb=b.valorContrato-(obraStats[b.id]?.despesas||0);
    return lb-la;
  });

  const barData=obrasVisiveis.map(o=>({
    nome:o.nome.length>14?o.nome.slice(0,14)+"…":o.nome,
    recebido:obraStats[o.id]?.recebido||0,
    despesas:obraStats[o.id]?.despesas||0,
  }));

  const catData=CATS.map(c=>({name:c,value:lans.filter(l=>l.tipo==="Despesa"&&l.categoria===c).reduce((s,l)=>s+l.valor,0)})).filter(x=>x.value>0);

  // Gráfico de linha — acumulado mês a mês
  const lineData=useMemo(()=>{
    const map={};
    lans.forEach(l=>{
      const m=l.data?.slice(0,7)||"";
      if(!m)return;
      if(!map[m])map[m]={mes:m,receita:0,despesa:0};
      if(l.tipo==="Receita")map[m].receita+=l.valor;
      else map[m].despesa+=l.valor;
    });
    const sorted=Object.values(map).sort((a,b)=>a.mes>b.mes?1:-1);
    let acRec=0,acDesp=0;
    return sorted.map(x=>{
      acRec+=x.receita; acDesp+=x.despesa;
      return{
        mes:new Date(x.mes+"-01").toLocaleDateString("pt-BR",{month:"short",year:"2-digit"}),
        "Receita Acumulada":Math.round(acRec),
        "Despesa Acumulada":Math.round(acDesp),
        "Resultado":Math.round(acRec-acDesp),
      };
    });
  },[lans]);

  return(
    <div>
      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:24,gap:16,flexWrap:"wrap"}}>
        <div>
          <h1 style={S.h1}>Dashboard Geral</h1>
          <p style={S.muted}>{obras.length} obra{obras.length!==1?"s":""} · {new Date().toLocaleDateString("pt-BR",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</p>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:13,color:"#64748b",whiteSpace:"nowrap"}}>Filtrar por obra:</span>
          <select value={filtroObra} onChange={e=>setFiltroObra(e.target.value)}
            style={{...S.input,width:220,padding:"8px 12px",fontSize:13}}>
            <option value="">Todas as obras</option>
            {obras.map(o=><option key={o.id} value={o.id}>{o.nome}</option>)}
          </select>
          {filtroObra&&<button onClick={()=>setFiltroObra("")} style={{...S.btnGhost,padding:"8px 12px",fontSize:12,whiteSpace:"nowrap"}}>✕ Limpar</button>}
        </div>
      </div>

      {filtroObra&&(()=>{const o=obras.find(x=>x.id===filtroObra);return o&&(
        <div style={{background:"#0c1f3a",border:"1px solid #1e3a5f",borderRadius:10,padding:"12px 18px",marginBottom:20,display:"flex",gap:16,alignItems:"center",flexWrap:"wrap"}}>
          <div style={{fontSize:13,color:"#93c5fd",fontWeight:600}}>📍 {o.nome}</div>
          <div style={{fontSize:12,color:"#475569"}}>Cliente: {o.cliente}</div>
          <div style={{fontSize:12,color:"#475569"}}>Tipo: {o.tipo}</div>
          <Tag s={o.status}/>
          {o.dataInicio&&<div style={{fontSize:12,color:"#475569"}}>Início: {fmtDate(o.dataInicio)}</div>}
          {o.dataFim&&<div style={{fontSize:12,color:"#475569"}}>Previsão: {fmtDate(o.dataFim)}</div>}
          <button onClick={()=>goToObra(filtroObra)} style={{...S.btnGhost,padding:"6px 12px",fontSize:12,marginLeft:"auto"}}>Ver dashboard completo →</button>
        </div>
      );})()}

      {alerts.length>0&&(
        <div style={{marginBottom:20}}>
          {alerts.map(o=>{
            const st=obraStats[o.id]||{};
            const lucro=o.valorContrato-(st.despesas||0);
            return <Alert key={o.id} type="danger" msg={`${o.nome}: ${lucro<0?`PREJUÍZO de ${fmt(Math.abs(lucro))}`:"Orçamento estourado"} — Despesas: ${fmt(st.despesas||0)} / Contrato: ${fmt(o.valorContrato)}`}/>;
          })}
        </div>
      )}

      <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12,marginBottom:24}}>
        <KpiCard label="Contratos" value={fmt(totalContrato)} color="#f59e0b" icon="📋" sub={`${obrasVisiveis.length} obra(s)`}/>
        <KpiCard label="Recebido" value={fmt(totalRec)} color="#4ade80" icon="↑"/>
        <KpiCard label="Gasto" value={fmt(totalDesp)} color="#f87171" icon="↓"/>
        <KpiCard label="Saldo Atual" value={fmt(saldoAtual)} color={saldoAtual>=0?"#4ade80":"#f87171"} icon="⚖"/>
        <KpiCard label="Lucro Projetado" value={fmt(totalLucro)} color={totalLucro>=0?"#4ade80":"#f87171"} icon="★"/>
      </div>

      {lineData.length>0&&(
        <div style={{...S.card,marginBottom:24}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
            <h3 style={S.h3}>Evolução Acumulada — Receitas vs Despesas</h3>
            <span style={{fontSize:12,color:"#475569"}}>{filtroObra?obras.find(o=>o.id===filtroObra)?.nome:"Todas as obras"}</span>
          </div>
          <div style={{fontSize:13,color:"#64748b",marginBottom:16}}>Acúmulo mês a mês ao longo do tempo</div>
          <div style={{display:"flex",gap:20,marginBottom:12}}>
            {[{cor:"#10b981",label:"Receita Acumulada"},{cor:"#f59e0b",label:"Despesa Acumulada"},{cor:"#3b82f6",label:"Resultado"}].map(x=>(
              <div key={x.label} style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:"#94a3b8"}}>
                <div style={{width:24,height:3,borderRadius:2,background:x.cor}}/>
                {x.label}
              </div>
            ))}
          </div>
          <ResponsiveContainer width="100%" height={260}>
            <AreaChart data={lineData} margin={{top:4,right:8,left:0,bottom:0}}>
              <defs>
                <linearGradient id="gRec" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#10b981" stopOpacity={0.15}/>
                  <stop offset="95%" stopColor="#10b981" stopOpacity={0}/>
                </linearGradient>
                <linearGradient id="gDesp" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#f59e0b" stopOpacity={0.15}/>
                  <stop offset="95%" stopColor="#f59e0b" stopOpacity={0}/>
                </linearGradient>
                <linearGradient id="gRes" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#3b82f6" stopOpacity={0.15}/>
                  <stop offset="95%" stopColor="#3b82f6" stopOpacity={0}/>
                </linearGradient>
              </defs>
              <XAxis dataKey="mes" tick={{fill:"#64748b",fontSize:12}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fill:"#64748b",fontSize:11}} axisLine={false} tickLine={false} tickFormatter={v=>`R$${(v/1000).toFixed(0)}k`} width={60}/>
              <Tooltip formatter={(v,n)=>[fmt(v),n]} contentStyle={{background:"#0d1220",border:"1px solid #1e2a3a",borderRadius:8,color:"#e2e8f0",fontSize:12}}/>
              <Area type="monotone" dataKey="Receita Acumulada" stroke="#10b981" strokeWidth={2.5} fill="url(#gRec)" dot={{fill:"#10b981",r:3}} activeDot={{r:5}}/>
              <Area type="monotone" dataKey="Despesa Acumulada" stroke="#f59e0b" strokeWidth={2.5} fill="url(#gDesp)" dot={{fill:"#f59e0b",r:3}} activeDot={{r:5}}/>
              <Area type="monotone" dataKey="Resultado" stroke="#3b82f6" strokeWidth={2} strokeDasharray="5 3" fill="url(#gRes)" dot={{fill:"#3b82f6",r:3}} activeDot={{r:5}}/>
            </AreaChart>
          </ResponsiveContainer>
        </div>
      )}

      {barData.length>0&&(
        <div style={{...S.grid2,marginBottom:24}}>
          <div style={S.card}>
            <h3 style={{...S.h3,marginBottom:16}}>Recebido vs Gasto por Obra</h3>
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={barData} barGap={4}>
                <XAxis dataKey="nome" tick={{fill:"#64748b",fontSize:11}} axisLine={false} tickLine={false}/>
                <YAxis tick={{fill:"#64748b",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>`${(v/1000).toFixed(0)}k`}/>
                <Tooltip formatter={v=>fmt(v)} contentStyle={{background:"#0d1220",border:"1px solid #1e2a3a",borderRadius:8,color:"#e2e8f0"}}/>
                <Bar dataKey="recebido" fill="#10b981" name="Recebido" radius={[4,4,0,0]}/>
                <Bar dataKey="despesas" fill="#f59e0b" name="Gasto" radius={[4,4,0,0]}/>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div style={S.card}>
            <h3 style={{...S.h3,marginBottom:4}}>Gastos por Categoria</h3>
            <div style={{fontSize:13,color:"#64748b",marginBottom:12}}>{filtroObra?obras.find(o=>o.id===filtroObra)?.nome:"Todas as obras"}</div>
            {catData.length===0?<div style={{color:"#334155",fontSize:13,paddingTop:20}}>Sem despesas registradas</div>:(
              <>
                <ResponsiveContainer width="100%" height={160}>
                  <PieChart>
                    <Pie data={catData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={70} innerRadius={35}>
                      {catData.map((_,i)=><Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>)}
                    </Pie>
                    <Tooltip formatter={v=>fmt(v)} contentStyle={{background:"#0d1220",border:"1px solid #1e2a3a",borderRadius:8,color:"#e2e8f0"}}/>
                  </PieChart>
                </ResponsiveContainer>
                <div style={{display:"flex",flexWrap:"wrap",gap:10,marginTop:8}}>
                  {catData.map((c,i)=>(
                    <div key={c.name} style={{display:"flex",alignItems:"center",gap:5,fontSize:12,color:"#94a3b8"}}>
                      <div style={{width:8,height:8,borderRadius:2,background:PIE_COLORS[i%PIE_COLORS.length]}}/>
                      {c.name}: <strong style={{color:"#f1f5f9"}}>{fmt(c.value)}</strong>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        </div>
      )}

      <div style={{...S.grid2,marginBottom:24}}>
        <div style={S.card}>
          <h3 style={{...S.h3,marginBottom:16}}>Ranking — Lucro Projetado</h3>
          {ranking.length===0?<div style={{color:"#334155",fontSize:13}}>Nenhuma obra cadastrada</div>:ranking.map((o,i)=>{
            const lucro=o.valorContrato-(obraStats[o.id]?.despesas||0);
            return(
              <div key={o.id} onClick={()=>goToObra(o.id)}
                style={{display:"flex",alignItems:"center",gap:12,padding:"10px 0",borderBottom:"1px solid #111827",cursor:"pointer"}}
                onMouseEnter={e=>e.currentTarget.style.opacity=".8"} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                <div style={{width:28,height:28,borderRadius:6,background:i===0?"#422006":i===1?"#1e3a5f":"#1a2235",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,color:i===0?"#f59e0b":i===1?"#93c5fd":"#64748b",flexShrink:0}}>{i+1}</div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:14,fontWeight:600,color:"#f1f5f9",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{o.nome}</div>
                  <div style={{fontSize:11,color:"#475569"}}>{o.cliente}</div>
                </div>
                <Tag s={o.status}/>
                <div style={{textAlign:"right",flexShrink:0}}>
                  <div style={{fontWeight:700,color:lucro>=0?"#4ade80":"#f87171",fontSize:14}}>{fmt(lucro)}</div>
                  <div style={{fontSize:11,color:"#475569"}}>lucro proj.</div>
                </div>
              </div>
            );
          })}
        </div>

        <div style={S.card}>
          <h3 style={{...S.h3,marginBottom:16}}>Status das Obras</h3>
          {obrasVisiveis.length===0?<div style={{color:"#334155",fontSize:13}}>Nenhuma obra encontrada</div>:obrasVisiveis.map(o=>{
            const st=obraStats[o.id]||{};
            const execPct=pct(st.despesas||0,o.valorContrato);
            return(
              <div key={o.id} onClick={()=>goToObra(o.id)}
                style={{marginBottom:16,cursor:"pointer",padding:8,borderRadius:8,transition:"background .15s"}}
                onMouseEnter={e=>e.currentTarget.style.background="#111827"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:6,gap:8}}>
                  <div style={{minWidth:0}}>
                    <span style={{fontSize:14,fontWeight:600,color:"#f1f5f9",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",display:"block"}}>{o.nome}</span>
                  </div>
                  <div style={{display:"flex",gap:6,alignItems:"center",flexShrink:0}}>
                    <Tag s={o.status} style={{fontSize:11}}/>
                    <span style={{fontSize:12,color:"#f59e0b",fontWeight:700}}>{execPct}%</span>
                  </div>
                </div>
                <div style={{background:"#1e2a3a",borderRadius:4,height:6,overflow:"hidden"}}>
                  <div style={{height:"100%",background:execPct>100?"#ef4444":"#f59e0b",width:`${Math.min(100,execPct)}%`,borderRadius:4}}/>
                </div>
                <div style={{display:"flex",justifyContent:"space-between",marginTop:4,fontSize:11,color:"#475569"}}>
                  <span>Gasto: {fmt(st.despesas||0)}</span>
                  <span>Contrato: {fmt(o.valorContrato)}</span>
                </div>
              </div>
            );
          })}
        </div>
      </div>

      <div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <h3 style={S.h3}>Últimos Lançamentos</h3>
          <button style={S.btnGhost} onClick={()=>setView("financeiro")}>Ver todos →</button>
        </div>
        {lans.length===0?<div style={{color:"#334155",textAlign:"center",padding:24}}>Nenhum lançamento ainda</div>:(
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{borderBottom:"1px solid #1e2a3a"}}>
              {["Data","Obra","Tipo","Categoria","Descrição","Valor"].map(h=><th key={h} style={{textAlign:"left",padding:"8px 12px",color:"#475569",fontWeight:500}}>{h}</th>)}
            </tr></thead>
            <tbody>
              {[...lans].sort((a,b)=>b.criadoEm>a.criadoEm?1:-1).slice(0,8).map(l=>{
                const obraNome=obras.find(o=>o.id===l.obraId)?.nome||"—";
                return(
                  <tr key={l.id} style={{borderBottom:"1px solid #111827"}}>
                    <td style={{padding:"10px 12px",color:"#64748b"}}>{fmtDate(l.data)}</td>
                    <td style={{padding:"10px 12px",color:"#94a3b8",maxWidth:140}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{obraNome}</div></td>
                    <td style={{padding:"10px 12px"}}><span style={{...S.tag,background:l.tipo==="Receita"?"#052e16":"#450a0a",color:l.tipo==="Receita"?"#4ade80":"#f87171"}}>{l.tipo}</span></td>
                    <td style={{padding:"10px 12px",color:"#94a3b8"}}>{l.categoria||"—"}</td>
                    <td style={{padding:"10px 12px",color:"#e2e8f0",maxWidth:200}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{l.descricao||"—"}</div></td>
                    <td style={{padding:"10px 12px",fontWeight:700,color:l.tipo==="Receita"?"#4ade80":"#f87171",whiteSpace:"nowrap"}}>{l.tipo==="Receita"?"+":"-"}{fmt(l.valor)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════
// PDF GENERATOR
// ═══════════════════════════════════════════════════════
async function gerarPDFObra(obra, lancamentos, obraStats){
  const jsPDFModule = await import("jspdf");
  const {jsPDF} = jsPDFModule;
  const doc = new jsPDF({orientation:"portrait",unit:"mm",format:"a4"});
  const pw = doc.internal.pageSize.getWidth();
  const st = obraStats[obra.id]||{};
  const aditivos=obra.aditivos||[];
  const totalAditivos=aditivos.reduce((s,a)=>s+(a.tipo==="Acréscimo"?a.valor:-a.valor),0);
  const contratoAtualizado=(obra.valorContrato||0)+totalAditivos;
  const lucro = contratoAtualizado-(st.despesas||0);
  const margem = contratoAtualizado>0?((lucro/contratoAtualizado)*100).toFixed(1):0;
  const dataHoje = new Date().toLocaleDateString("pt-BR");

  // Paleta de cores por categoria (RGB)
  const CAT_COLORS={
    "Material":{r:249,g:115,b:22,bg:[255,247,237],border:[254,215,170]},
    "Mão de Obra":{r:139,g:92,b:246,bg:[245,243,255],border:[221,214,254]},
    "Equipamento":{r:59,g:130,b:246,bg:[239,246,255],border:[191,219,254]},
    "Comissão":{r:234,g:179,b:8,bg:[254,252,232],border:[254,240,138]},
    "Outros":{r:16,g:185,b:129,bg:[236,253,245],border:[167,243,208]},
  };

  // ═══ HEADER ═══
  doc.setFillColor(13,18,32);
  doc.rect(0,0,pw,48,"F");
  // Barra lateral laranja
  doc.setFillColor(245,158,11);
  doc.rect(0,0,5,48,"F");
  // Detalhe decorativo
  doc.setFillColor(30,42,58);
  doc.rect(0,44,pw,1,"F");
  doc.setFillColor(245,158,11);
  doc.rect(0,44,60,1,"F");

  // Título
  doc.setFont("helvetica","bold");
  doc.setFontSize(20);
  doc.setTextColor(241,245,249);
  doc.text(obra.nome,12,17);
  
  // Modalidade tag
  doc.setFontSize(7);
  doc.setTextColor(245,158,11);
  doc.text(`◆ ${(obra.modalidade||"Empreitada").toUpperCase()}`,12,22);
  
  doc.setFontSize(9);
  doc.setTextColor(148,163,184);
  doc.text(`Cliente: ${obra.cliente}`,12,29);
  if(obra.endereco) doc.text(`Endereço: ${obra.endereco}`,12,34);
  
  // Info linha
  doc.setFontSize(8);
  doc.setTextColor(100,116,139);
  const infoParts=[`Status: ${obra.status}`,`Tipo: ${obra.tipo}`];
  if(obra.metrosQuadrados>0)infoParts.push(`Área: ${obra.metrosQuadrados}m²`);
  doc.text(infoParts.join("  |  "),12,40);

  // Datas
  if(obra.dataInicio||obra.dataFim){
    doc.setFontSize(8);
    doc.setTextColor(148,163,184);
    doc.text(`Início: ${fmtDate(obra.dataInicio)}`,pw-14,34,{align:"right"});
    doc.text(`Previsão: ${fmtDate(obra.dataFim)}`,pw-14,40,{align:"right"});
  }
  doc.setFontSize(7);
  doc.setTextColor(100,116,139);
  doc.text(`Emitido em ${dataHoje}`,pw-14,17,{align:"right"});

  // ═══ KPIs ═══
  let y=58;
  const kpis=[
    {label:"CONTRATO ATUAL",value:fmt(contratoAtualizado),cor:[245,158,11],bg:[255,247,237]},
    {label:"TOTAL RECEBIDO",value:fmt(st.recebido||0),cor:[16,185,129],bg:[236,253,245]},
    {label:"TOTAL GASTO",value:fmt(st.despesas||0),cor:[239,68,68],bg:[254,242,242]},
    {label:"SALDO ATUAL",value:fmt(st.saldo||0),cor:(st.saldo||0)>=0?[16,185,129]:[239,68,68],bg:(st.saldo||0)>=0?[236,253,245]:[254,242,242]},
    {label:"LUCRO PROJETADO",value:fmt(lucro),cor:lucro>=0?[16,185,129]:[239,68,68],bg:lucro>=0?[236,253,245]:[254,242,242]},
    {label:"MARGEM",value:`${margem}%`,cor:lucro>=0?[16,185,129]:[239,68,68],bg:lucro>=0?[236,253,245]:[254,242,242]},
  ];
  const kw=(pw-24-8)/3;
  kpis.forEach((k,i)=>{
    const col=i%3; const row=Math.floor(i/3);
    const kx=12+col*(kw+4); const ky=y+row*22;
    // Fundo colorido claro
    doc.setFillColor(...k.bg); doc.roundedRect(kx,ky,kw,19,2.5,2.5,"F");
    // Barra colorida à esquerda
    doc.setFillColor(...k.cor); doc.roundedRect(kx,ky,1.8,19,1,1,"F");
    // Label
    doc.setFontSize(7); doc.setTextColor(100,116,139); doc.setFont("helvetica","bold");
    doc.text(k.label,kx+5,ky+7);
    // Valor
    doc.setFontSize(12); doc.setFont("helvetica","bold"); doc.setTextColor(...k.cor);
    doc.text(k.value,kx+5,ky+15);
  });
  y+=48;

  // ═══ ADITIVOS (se houver) ═══
  if(aditivos.length>0){
    doc.setFillColor(239,246,255); doc.roundedRect(12,y,pw-24,7,2,2,"F");
    doc.setFillColor(59,130,246); doc.roundedRect(12,y,1.8,7,1,1,"F");
    doc.setFontSize(8); doc.setFont("helvetica","bold"); doc.setTextColor(59,130,246);
    doc.text(`ADITIVOS DO CONTRATO — ${aditivos.length} registro(s)`,17,y+4.6);
    doc.text(`${totalAditivos>=0?"+":""}${fmt(totalAditivos)}`,pw-14,y+4.6,{align:"right"});
    y+=10;
    aditivos.forEach(a=>{
      if(y>270){doc.addPage();y=20;}
      const isAcresc=a.tipo==="Acréscimo";
      doc.setFontSize(7); doc.setFont("helvetica","normal"); doc.setTextColor(100,116,139);
      doc.text(fmtDate(a.data),14,y);
      doc.setTextColor(51,65,85); doc.setFont("helvetica","bold");
      doc.text(a.tipo,38,y);
      doc.setTextColor(30,41,59); doc.setFont("helvetica","normal");
      doc.text((a.descricao||"").substring(0,55),58,y);
      doc.setTextColor(...(isAcresc?[16,185,129]:[239,68,68])); doc.setFont("helvetica","bold");
      doc.text(`${isAcresc?"+":"-"}${fmt(a.valor)}`,pw-14,y,{align:"right"});
      doc.setDrawColor(226,232,240); doc.line(12,y+1.5,pw-12,y+1.5);
      y+=5.5;
    });
    y+=3;
  }

  // ═══ ANÁLISE m² ═══
  if(obra.metrosQuadrados>0){
    const m2=obra.metrosQuadrados;
    const contratoM2=contratoAtualizado/m2;
    const custoM2=(st.despesas||0)/m2;
    const margemM2=contratoM2-custoM2;
    if(y>250){doc.addPage();y=20;}
    doc.setFillColor(239,246,255); doc.roundedRect(12,y,pw-24,20,2.5,2.5,"F");
    doc.setFillColor(59,130,246); doc.roundedRect(12,y,1.8,20,1,1,"F");
    doc.setFontSize(8); doc.setFont("helvetica","bold"); doc.setTextColor(59,130,246);
    doc.text(`ANÁLISE POR m²  —  ${m2} m²`,17,y+6);
    const w4=(pw-40)/4;
    const items=[
      {label:"CONTRATO/m²",val:fmt(contratoM2),cor:[245,158,11]},
      {label:"CUSTO/m²",val:custoM2>0?fmt(custoM2):"—",cor:[239,68,68]},
      {label:"MARGEM/m²",val:custoM2>0?fmt(margemM2):"—",cor:margemM2>=0?[16,185,129]:[239,68,68]},
      {label:"EXECUTADO",val:`${pct(custoM2,contratoM2)}%`,cor:[59,130,246]},
    ];
    items.forEach((it,i)=>{
      const x=17+i*w4;
      doc.setFontSize(6); doc.setTextColor(100,116,139); doc.setFont("helvetica","bold");
      doc.text(it.label,x,y+12);
      doc.setFontSize(9); doc.setTextColor(...it.cor);
      doc.text(it.val,x,y+17);
    });
    y+=26;
  }

  // ═══ SÓCIOS (Construção e Venda) ═══
  if(obra.modalidade==="Construção e Venda"&&obra.socios&&obra.socios.length>0){
    if(y>240){doc.addPage();y=20;}
    const totalDespesas=st.despesas||0;
    const valorVenda=obra.valorVendaReal||obra.valorVendaEstimado||0;
    const lucroVenda=valorVenda-totalDespesas;
    
    doc.setFillColor(245,243,255); doc.roundedRect(12,y,pw-24,8,2,2,"F");
    doc.setFillColor(139,92,246); doc.roundedRect(12,y,1.8,8,1,1,"F");
    doc.setFontSize(8); doc.setFont("helvetica","bold"); doc.setTextColor(139,92,246);
    doc.text(`ACERTO DE CONTAS — SÓCIOS`,17,y+5.5);
    if(valorVenda>0){
      doc.setTextColor(100,116,139); doc.setFontSize(7);
      doc.text(`Venda: ${fmt(valorVenda)}  |  Lucro: ${fmt(lucroVenda)}`,pw-14,y+5.5,{align:"right"});
    }
    y+=12;

    // Header tabela
    doc.setFillColor(250,250,252); doc.rect(12,y-3,pw-24,5,"F");
    doc.setFontSize(6); doc.setFont("helvetica","bold"); doc.setTextColor(139,92,246);
    doc.text("SÓCIO",14,y);
    doc.text("%",70,y);
    doc.text("DEVERIA",85,y);
    doc.text("PAGOU",115,y);
    doc.text("ACERTO",145,y);
    doc.text("LUCRO",pw-14,y,{align:"right"});
    y+=5;

    obra.socios.forEach(s=>{
      if(y>270){doc.addPage();y=20;}
      const pago=lancamentos.filter(l=>l.pagoPor===s.id&&l.tipo==="Despesa").reduce((sum,l)=>sum+l.valor,0);
      const deveria=totalDespesas*s.percentual/100;
      const acerto=pago-deveria;
      const cotaLucro=lucroVenda*s.percentual/100;
      doc.setFontSize(8); doc.setFont("helvetica","bold"); doc.setTextColor(30,41,59);
      doc.text(s.nome.substring(0,25),14,y);
      doc.setFont("helvetica","normal"); doc.setTextColor(139,92,246);
      doc.text(`${s.percentual}%`,70,y);
      doc.setTextColor(100,116,139);
      doc.text(fmt(deveria),85,y);
      doc.setTextColor(30,41,59); doc.setFont("helvetica","bold");
      doc.text(fmt(pago),115,y);
      doc.setTextColor(...(Math.abs(acerto)<1?[100,116,139]:acerto>0?[16,185,129]:[239,68,68]));
      doc.text(`${acerto>0?"+":""}${fmt(acerto)}`,145,y);
      doc.setTextColor(...(cotaLucro>=0?[16,185,129]:[239,68,68]));
      doc.text(valorVenda>0?fmt(cotaLucro):"—",pw-14,y,{align:"right"});
      doc.setDrawColor(226,232,240); doc.line(12,y+1.5,pw-12,y+1.5);
      y+=5.5;
    });
    y+=4;
  }

  const despesas=lancamentos.filter(l=>l.obraId===obra.id&&l.tipo==="Despesa");
  const receitas=lancamentos.filter(l=>l.obraId===obra.id&&l.tipo==="Receita");

  // ═══ RECEITAS ═══
  if(receitas.length>0){
    if(y>250){doc.addPage();y=20;}
    doc.setFillColor(236,253,245); doc.roundedRect(12,y,pw-24,8,2,2,"F");
    doc.setFillColor(16,185,129); doc.roundedRect(12,y,1.8,8,1,1,"F");
    doc.setFontSize(9); doc.setFont("helvetica","bold"); doc.setTextColor(16,185,129);
    doc.text(`RECEITAS — ${receitas.length} lançamento(s)`,17,y+5.5);
    doc.text(`+${fmt(receitas.reduce((s,l)=>s+l.valor,0))}`,pw-14,y+5.5,{align:"right"});
    y+=12;
    receitas.sort((a,b)=>b.data>a.data?1:-1).forEach(l=>{
      if(y>270){doc.addPage();y=20;}
      doc.setFontSize(7); doc.setFont("helvetica","normal"); doc.setTextColor(100,116,139);
      doc.text(fmtDate(l.data),14,y);
      doc.setTextColor(51,65,85); doc.setFont("helvetica","bold");
      doc.text((l.descricao||"—").substring(0,50),38,y);
      doc.setFont("helvetica","normal"); doc.setTextColor(100,116,139);
      doc.text((l.responsavel||"—").substring(0,20),120,y);
      doc.setTextColor(150,150,150); doc.setFontSize(6);
      doc.text(l.pagamento||"—",155,y);
      doc.setFontSize(8); doc.setTextColor(16,185,129); doc.setFont("helvetica","bold");
      doc.text(`+${fmt(l.valor)}`,pw-14,y,{align:"right"});
      doc.setDrawColor(230,238,244); doc.line(12,y+1.5,pw-12,y+1.5);
      y+=5.5;
    });
    y+=3;
  }

  // ═══ DESPESAS POR CATEGORIA ═══
  CATS.forEach(cat=>{
    const catLans=despesas.filter(l=>l.categoria===cat);
    if(catLans.length===0)return;
    const catTotal=catLans.reduce((s,l)=>s+l.valor,0);
    const cc=CAT_COLORS[cat]||{r:100,g:116,b:139,bg:[241,245,249],border:[203,213,225]};
    if(y>260){doc.addPage();y=20;}
    
    // Header da categoria com cor específica
    doc.setFillColor(...cc.bg); doc.roundedRect(12,y,pw-24,8,2,2,"F");
    doc.setFillColor(cc.r,cc.g,cc.b); doc.roundedRect(12,y,1.8,8,1,1,"F");
    doc.setFontSize(9); doc.setFont("helvetica","bold"); doc.setTextColor(cc.r,cc.g,cc.b);
    doc.text(`${cat.toUpperCase()} — ${catLans.length} lançamento(s)`,17,y+5.5);
    doc.text(`-${fmt(catTotal)}`,pw-14,y+5.5,{align:"right"});
    y+=12;
    
    catLans.sort((a,b)=>b.data>a.data?1:-1).forEach(l=>{
      if(y>270){doc.addPage();y=20;}
      doc.setFontSize(7); doc.setFont("helvetica","normal"); doc.setTextColor(100,116,139);
      doc.text(fmtDate(l.data),14,y);
      doc.setTextColor(cc.r,cc.g,cc.b); doc.setFont("helvetica","bold"); doc.setFontSize(6.5);
      doc.text((l.subcategoria||"—").substring(0,15),38,y);
      doc.setFont("helvetica","normal"); doc.setTextColor(30,41,59); doc.setFontSize(7);
      doc.text((l.descricao||"—").substring(0,40),68,y);
      doc.setTextColor(100,116,139);
      doc.text((l.responsavel||"—").substring(0,18),128,y);
      doc.setFontSize(6); doc.setTextColor(148,163,184);
      doc.text(l.pagamento||"—",158,y);
      doc.setFontSize(8); doc.setTextColor(cc.r,cc.g,cc.b); doc.setFont("helvetica","bold");
      doc.text(`-${fmt(l.valor)}`,pw-14,y,{align:"right"});
      doc.setDrawColor(...cc.border); doc.line(12,y+1.5,pw-12,y+1.5);
      y+=5.5;
    });
    y+=3;
  });

  // ═══ RODAPÉ ═══
  const pages=doc.internal.getNumberOfPages();
  for(let i=1;i<=pages;i++){
    doc.setPage(i);
    doc.setFillColor(13,18,32); doc.rect(0,283,pw,14,"F");
    doc.setFillColor(245,158,11); doc.rect(0,283,pw,0.5,"F");
    doc.setFontSize(7); doc.setFont("helvetica","bold"); doc.setTextColor(245,158,11);
    doc.text("OBRA GESTÃO",12,289);
    doc.setFont("helvetica","normal"); doc.setTextColor(148,163,184);
    doc.text(`${obra.nome}  ·  Gerado em ${dataHoje}`,12,293);
    doc.setTextColor(100,116,139);
    doc.text(`Página ${i} de ${pages}`,pw-12,291,{align:"right"});
  }

  doc.save(`Relatorio_${obra.nome.replace(/\s+/g,"_")}_${dataHoje.replace(/\//g,"-")}.pdf`);
}

// ═══════════════════════════════════════════════════════
// RELATÓRIOS VIEW
// ═══════════════════════════════════════════════════════
function RelatoriosView({obras,lancamentos,obraStats}){
  const [filterObra,setFilterObra]=useState("");
  const [gerandoPDF,setGerandoPDF]=useState(null);

  const consolidado=obras.map(o=>{
    const st=obraStats[o.id]||{};
    return{...o,...st,lucro:o.valorContrato-(st.despesas||0),margem:o.valorContrato>0?((o.valorContrato-(st.despesas||0))/o.valorContrato*100).toFixed(1):0};
  });

  const totalContrato=consolidado.reduce((s,o)=>s+o.valorContrato,0);
  const totalRec=consolidado.reduce((s,o)=>s+(o.recebido||0),0);
  const totalDesp=consolidado.reduce((s,o)=>s+(o.despesas||0),0);
  const totalLucro=consolidado.reduce((s,o)=>s+o.lucro,0);
  const margemMedia=totalContrato>0?(totalLucro/totalContrato*100).toFixed(1):0;

  const handlePDF=async(obra)=>{
    setGerandoPDF(obra.id);
    try{ await gerarPDFObra(obra,lancamentos,obraStats); }
    catch(e){ alert("Erro ao gerar PDF. Tente novamente."); }
    setGerandoPDF(null);
  };

  return(
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:24}}>
        <div><h1 style={S.h1}>Relatórios</h1><p style={S.muted}>Análise consolidada de todas as obras</p></div>
      </div>

      <div style={{...S.grid4,marginBottom:24}}>
        <KpiCard label="Contratos" value={fmt(totalContrato)} color="#f59e0b" icon="📋"/>
        <KpiCard label="Receitas" value={fmt(totalRec)} color="#4ade80" icon="↑"/>
        <KpiCard label="Despesas" value={fmt(totalDesp)} color="#f87171" icon="↓"/>
        <KpiCard label="Lucro Total" value={fmt(totalLucro)} color={totalLucro>=0?"#4ade80":"#f87171"} icon="★" sub={`Margem: ${margemMedia}%`}/>
      </div>

      {/* PDF por obra */}
      <div style={{...S.card,marginBottom:20}}>
        <div style={{marginBottom:16}}>
          <h3 style={S.h3}>Relatório PDF por Obra</h3>
          <p style={{...S.muted,marginTop:4}}>Gere um relatório completo em PDF para cada obra, separado por categoria de despesa.</p>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12}}>
          {obras.length===0&&<div style={{color:"#334155",fontSize:13}}>Nenhuma obra cadastrada</div>}
          {obras.map(o=>{
            const st=obraStats[o.id]||{};
            const lucro=o.valorContrato-(st.despesas||0);
            const lansObra=lancamentos.filter(l=>l.obraId===o.id);
            const isGerando=gerandoPDF===o.id;
            return(
              <div key={o.id} style={{background:"#111827",border:"1px solid #1e2a3a",borderRadius:10,padding:"16px"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                  <div style={{minWidth:0}}>
                    <div style={{fontWeight:700,color:"#f1f5f9",fontSize:14,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{o.nome}</div>
                    <div style={{fontSize:12,color:"#64748b",marginTop:2}}>{o.cliente}</div>
                  </div>
                  <Tag s={o.status}/>
                </div>
                <div style={{display:"flex",gap:12,marginBottom:12}}>
                  <div>
                    <div style={{fontSize:10,color:"#475569"}}>CONTRATO</div>
                    <div style={{fontSize:13,fontWeight:700,color:"#f59e0b"}}>{fmt(o.valorContrato)}</div>
                  </div>
                  <div>
                    <div style={{fontSize:10,color:"#475569"}}>GASTO</div>
                    <div style={{fontSize:13,fontWeight:700,color:"#f87171"}}>{fmt(st.despesas||0)}</div>
                  </div>
                  <div>
                    <div style={{fontSize:10,color:"#475569"}}>LUCRO</div>
                    <div style={{fontSize:13,fontWeight:700,color:lucro>=0?"#4ade80":"#f87171"}}>{fmt(lucro)}</div>
                  </div>
                </div>
                <div style={{fontSize:11,color:"#334155",marginBottom:10}}>
                  {lansObra.length} lançamento(s) · {lansObra.filter(l=>l.tipo==="Despesa").length} despesa(s) · {lansObra.filter(l=>l.tipo==="Receita").length} receita(s)
                </div>
                <button onClick={()=>handlePDF(o)} disabled={isGerando}
                  style={{width:"100%",padding:"9px 0",borderRadius:8,border:"none",cursor:isGerando?"not-allowed":"pointer",
                    background:isGerando?"#1e2a3a":"#450a0a",color:isGerando?"#475569":"#fca5a5",
                    fontSize:13,fontWeight:600,display:"flex",alignItems:"center",justifyContent:"center",gap:6,transition:"all .15s"}}
                  onMouseEnter={e=>{if(!isGerando)e.currentTarget.style.background="#7f1d1d"}}
                  onMouseLeave={e=>{if(!isGerando)e.currentTarget.style.background="#450a0a"}}>
                  {isGerando?"⏳ Gerando PDF...":"📄 Gerar Relatório PDF"}
                </button>
              </div>
            );
          })}
        </div>
      </div>

      <div style={{...S.card,marginBottom:20}}>
        <h3 style={{...S.h3,marginBottom:16}}>Consolidado por Obra</h3>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{borderBottom:"1px solid #1e2a3a"}}>
              {["Obra","Cliente","Status","Contrato","Recebido","Gasto","Saldo","Lucro Proj.","Margem","PDF"].map(h=>(
                <th key={h} style={{textAlign:"left",padding:"8px 12px",color:"#475569",fontWeight:500,whiteSpace:"nowrap"}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {consolidado.map(o=>(
                <tr key={o.id} style={{borderBottom:"1px solid #111827"}}>
                  <td style={{padding:"10px 12px",fontWeight:600,color:"#f1f5f9"}}>{o.nome}</td>
                  <td style={{padding:"10px 12px",color:"#94a3b8"}}>{o.cliente}</td>
                  <td style={{padding:"10px 12px"}}><Tag s={o.status}/></td>
                  <td style={{padding:"10px 12px",color:"#f59e0b",fontWeight:600}}>{fmt(o.valorContrato)}</td>
                  <td style={{padding:"10px 12px",color:"#4ade80"}}>{fmt(o.recebido||0)}</td>
                  <td style={{padding:"10px 12px",color:"#f87171"}}>{fmt(o.despesas||0)}</td>
                  <td style={{padding:"10px 12px",fontWeight:700,color:(o.saldo||0)>=0?"#4ade80":"#f87171"}}>{fmt(o.saldo||0)}</td>
                  <td style={{padding:"10px 12px",fontWeight:700,color:o.lucro>=0?"#4ade80":"#f87171"}}>{fmt(o.lucro)}</td>
                  <td style={{padding:"10px 12px",color:parseFloat(o.margem)>=0?"#4ade80":"#f87171"}}>{o.margem}%</td>
                  <td style={{padding:"10px 12px"}}>
                    <button onClick={()=>handlePDF(o)} disabled={gerandoPDF===o.id}
                      style={{padding:"5px 10px",borderRadius:6,border:"none",cursor:"pointer",background:"#450a0a",color:"#fca5a5",fontSize:11,fontWeight:600}}>
                      {gerandoPDF===o.id?"...":"📄 PDF"}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr style={{borderTop:"2px solid #1e2a3a",fontWeight:700}}>
                <td colSpan={3} style={{padding:"12px 12px",color:"#64748b"}}>TOTAIS</td>
                <td style={{padding:"12px 12px",color:"#f59e0b"}}>{fmt(totalContrato)}</td>
                <td style={{padding:"12px 12px",color:"#4ade80"}}>{fmt(totalRec)}</td>
                <td style={{padding:"12px 12px",color:"#f87171"}}>{fmt(totalDesp)}</td>
                <td style={{padding:"12px 12px",color:(totalRec-totalDesp)>=0?"#4ade80":"#f87171"}}>{fmt(totalRec-totalDesp)}</td>
                <td style={{padding:"12px 12px",color:totalLucro>=0?"#4ade80":"#f87171"}}>{fmt(totalLucro)}</td>
                <td style={{padding:"12px 12px",color:parseFloat(margemMedia)>=0?"#4ade80":"#f87171"}}>{margemMedia}%</td>
                <td/>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════
export default function App(){
  const [obras,setObras]=useState([]);
  const [lancamentos,setLancamentos]=useState([]);
  const [equipes,setEquipes]=useState([]);
  const [view,setView]=useState("dashboard");
  const [selectedObra,setSelectedObra]=useState(null);
  const [loading,setLoading]=useState(true);
  const [toast,setToast]=useState(null);
  const [sidebarOpen,setSidebarOpen]=useState(false);
  const [autenticado,setAutenticado]=useState(()=>{
    try{
      const saved=localStorage.getItem(PIN_STORAGE_KEY);
      if(!saved)return false;
      const diff=(Date.now()-parseInt(saved))/(1000*60*60);
      return diff < PIN_VALIDADE_HORAS;
    }catch{return false;}
  });

  const logout=()=>{
    localStorage.removeItem(PIN_STORAGE_KEY);
    setAutenticado(false);
  };

  const showToast=(ok,msg)=>{
    setToast({ok,msg});
    setTimeout(()=>setToast(null),4000);
  };

  useEffect(()=>{
    Promise.all([db.get("obras"),db.get("lancamentos"),db.get("equipes")]).then(([o,l,e])=>{
      if(o)setObras(o);
      if(l)setLancamentos(l);
      if(e)setEquipes(e);
      setLoading(false);
    });
  },[]);

  useEffect(()=>{if(!loading)db.set("obras",obras)},[obras,loading]);
  useEffect(()=>{if(!loading)db.set("lancamentos",lancamentos)},[lancamentos,loading]);
  useEffect(()=>{if(!loading)db.set("equipes",equipes)},[equipes,loading]);

  const obraStats=useMemo(()=>obras.reduce((acc,obra)=>{
    const lans=lancamentos.filter(l=>l.obraId===obra.id);
    const recebido=lans.filter(l=>l.tipo==="Receita").reduce((s,l)=>s+l.valor,0);
    const despesas=lans.filter(l=>l.tipo==="Despesa").reduce((s,l)=>s+l.valor,0);
    acc[obra.id]={recebido,despesas,saldo:recebido-despesas};
    return acc;
  },{}),[obras,lancamentos]);

  const addLancamento=useCallback(lan=>setLancamentos(p=>[...p,{...lan,id:uid(),criadoEm:new Date().toISOString()}]),[]);
  const deleteLancamento=useCallback(id=>setLancamentos(p=>p.filter(l=>l.id!==id)),[]);
  const updateObra=useCallback(obra=>setObras(p=>p.map(o=>o.id===obra.id?obra:o)),[]);
  const changeView=useCallback(v=>{setView(v);setSelectedObra(null)},[]);
  const goToObra=useCallback(id=>{setSelectedObra(id);setView("obra-detail")},[]);

  const handleExportar=useCallback(()=>{
    try{
      exportarExcel(obras,lancamentos,equipes);
      showToast(true,"Backup exportado com sucesso!");
    }catch{
      showToast(false,"Erro ao exportar. Tente novamente.");
    }
  },[obras,lancamentos,equipes]);

  const handleImportar=useCallback(file=>{
    importarExcel(file,setObras,setLancamentos,setEquipes,(ok,msg)=>showToast(ok,msg));
  },[]);

  if(!autenticado)return <LoginScreen onSuccess={()=>setAutenticado(true)}/>;

  if(loading)return(
    <div style={{background:"#060b14",height:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:16}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Syne:wght@800&display=swap');`}</style>
      <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:32,fontWeight:800,color:"#f59e0b"}}>OBRA GESTÃO</div>
      <div style={{color:"#334155",fontSize:14}}>Carregando sistema...</div>
    </div>
  );

  return(
    <div style={S.app}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        html,body{overflow-x:hidden}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-track{background:#0d1220}
        ::-webkit-scrollbar-thumb{background:#1e2a3a;border-radius:3px}
        input,select,textarea{outline:none;font-family:inherit;font-size:16px !important}
        @media(min-width:769px){input,select,textarea{font-size:14px !important}}
        button{font-family:inherit}
        table{font-size:12px}
        .mobile-menu-btn{display:none}
        @media(max-width:768px){
          body{-webkit-text-size-adjust:100%}
          .sidebar-desktop{position:fixed !important;left:-260px;top:0;bottom:0;z-index:200;transition:left .25s;height:100vh !important;width:240px !important}
          .sidebar-desktop.open{left:0;box-shadow:0 0 40px rgba(0,0,0,.6)}
          .mobile-overlay{position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:150;display:none}
          .mobile-overlay.open{display:block}
          .mobile-menu-btn{display:flex !important;align-items:center;justify-content:center;position:fixed;top:10px;left:10px;z-index:100;background:#f59e0b;color:#0d0d0d;border:none;width:42px;height:42px;padding:0;border-radius:8px;font-size:22px;font-weight:700;cursor:pointer;box-shadow:0 4px 12px rgba(0,0,0,.4)}
          main{padding:16px 10px !important;padding-top:60px !important;width:100vw !important;max-width:100vw !important;overflow-x:hidden}
          .mobile-hide{display:none !important}
          h1{font-size:20px !important;line-height:1.2}
          h2{font-size:16px !important}
          h3{font-size:14px !important}
          [style*="grid-template-columns"]{grid-template-columns:1fr !important}
          [style*="repeat(5,1fr)"]{grid-template-columns:repeat(2,1fr) !important}
          [style*="repeat(4,1fr)"]{grid-template-columns:repeat(2,1fr) !important}
          [style*="repeat(3,1fr)"]{grid-template-columns:repeat(2,1fr) !important}
          [style*="repeat(6,1fr)"]{grid-template-columns:repeat(2,1fr) !important}
          [style*="repeat(auto-fit"]{grid-template-columns:1fr !important}
          [style*="minmax"]{grid-template-columns:1fr !important}
          [style*="padding: \"20px 24px\""],[style*="padding:\"20px 24px\""]{padding:14px 14px !important}
          [style*="padding: \"16px 20px\""],[style*="padding:\"16px 20px\""]{padding:12px 14px !important}
          table{font-size:11px}
          th,td{padding:6px 6px !important;white-space:nowrap}
          .responsive-table-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
        }
      `}</style>
      {toast&&(
        <div style={{position:"fixed",top:20,right:20,left:20,zIndex:999,background:toast.ok?"#052e16":"#450a0a",border:`1px solid ${toast.ok?"#166534":"#7f1d1d"}`,borderRadius:10,padding:"12px 20px",color:toast.ok?"#4ade80":"#f87171",fontSize:14,fontWeight:600,boxShadow:"0 4px 24px rgba(0,0,0,.5)",display:"flex",alignItems:"center",gap:8,maxWidth:400,margin:"0 auto"}}>
          <span>{toast.ok?"✓":"✕"}</span>{toast.msg}
        </div>
      )}
      <button className="mobile-menu-btn" onClick={()=>setSidebarOpen(!sidebarOpen)}>☰</button>
      <div className={`mobile-overlay ${sidebarOpen?"open":""}`} onClick={()=>setSidebarOpen(false)}/>
      <div className={`sidebar-desktop ${sidebarOpen?"open":""}`} style={S.sidebar}>
        <Sidebar view={view} setView={(v)=>{changeView(v);setSidebarOpen(false);}} onExportar={handleExportar} onImportar={handleImportar} onLogout={logout}/>
      </div>
      <main style={S.main}>
        {view==="dashboard"&&<DashboardView obras={obras} obraStats={obraStats} lancamentos={lancamentos} setView={changeView} setSelectedObra={setSelectedObra} goToObra={goToObra}/>}
        {view==="obras"&&<ObrasView obras={obras} setObras={setObras} obraStats={obraStats} setView={changeView} setSelectedObra={setSelectedObra} goToObra={goToObra}/>}
        {view==="obra-detail"&&<ObraDetail obra={obras.find(o=>o.id===selectedObra)} stats={obraStats[selectedObra]} lancamentos={lancamentos.filter(l=>l.obraId===selectedObra)} setView={changeView} onUpdateObra={updateObra} addLancamento={addLancamento} allObras={obras}/>}
        {view==="financeiro"&&<FinanceiroView obras={obras} lancamentos={lancamentos} addLancamento={addLancamento} deleteLancamento={deleteLancamento}/>}
        {view==="mao-de-obra"&&<MaoDeObraView obras={obras} equipes={equipes} setEquipes={setEquipes} addLancamento={addLancamento} lancamentos={lancamentos}/>}
        {view==="relatorios"&&<RelatoriosView obras={obras} lancamentos={lancamentos} obraStats={obraStats}/>}
      </main>
    </div>
  );
}
