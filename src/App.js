import { useState, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";

const PIN_CORRETO = "GREMIO1903";

const G = {
  azul: "#0055A5", azulEscuro: "#003D7A", azulMedio: "#1A6BBF",
  preto: "#0D0D0D", branco: "#FFFFFF", cinzaClaro: "#F0F4F9",
  cinzaBorda: "#C8D8EC", cinzaTexto: "#5A7A9A", dourado: "#C9A84C",
};

const MOTIVOS_LIBERACAO = ["Criança","Sócio inadimplente","Transferência","Setor incorreto","Acesso já utilizado","Responsável já acessou","Catraca com defeito"];

const SETORES_PORTOES = [
  { setor: "Gramado Leste (K,M,O)", portoes: ["Portão K","Portão M","Portão O"] },
  { setor: "Gramado Oeste (B,D,Y)", portoes: ["Portão B","Portão D","Portão Y"] },
  { setor: "Gramado Sul (F,H)", portoes: ["Portão F","Portão H"] },
  { setor: "Superior Leste (L,P)", portoes: ["Portão L","Portão P"] },
  { setor: "Superior Sul (E)", portoes: ["Portão E"] },
  { setor: "Arquibancada Norte (Q,S,U,W)", portoes: ["Portão Q","Portão S","Portão U","Portão W"] },
  { setor: "Superior Oeste (C,X)", portoes: ["Portão C","Portão X"] },
  { setor: "Superior Norte (R,V)", portoes: ["Portão R","Portão V"] },
  { setor: "Gold Oeste (A)", portoes: ["Portão A"] },
  { setor: "Gold Leste (N)", portoes: ["Portão N"] },
  { setor: "Gold Premium Sul (G)", portoes: ["Portão G"] },
  { setor: "Visitante", portoes: ["Portão 6"] },
  { setor: "Totens", portoes: ["Toten A conselho","Toten A camarote","Toten N camarote","N-1 Tribuna (dir.)","N-1 Tribuna (esq.)","Toten 1/2","Toten 3/4","Toten 5","Toten 6","Toten 8","Toten 10/11","Toten 12/13","Toten 15/16/17"] },
];

const LIDERES = ["Guilherme","Keven","Sheron","Alexsandro","Ebert","Átila","Andrieli","Cristian","Fábio","Jesse","Franciely","Emily","Tainá"];

function formatCPF(v) {
  const d = v.replace(/\D/g,"").slice(0,11);
  if(d.length<=3) return d;
  if(d.length<=6) return d.slice(0,3)+"."+d.slice(3);
  if(d.length<=9) return d.slice(0,3)+"."+d.slice(3,6)+"."+d.slice(6);
  return d.slice(0,3)+"."+d.slice(3,6)+"."+d.slice(6,9)+"-"+d.slice(9);
}
function validateCPF(cpf){ return cpf.replace(/\D/g,"").length===11; }

function EscudoGremio({size=40}){
  return(
    <svg width={size} height={size} viewBox="0 0 100 110" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M50 4 L92 22 L92 60 C92 82 72 100 50 106 C28 100 8 82 8 60 L8 22 Z" fill="#0055A5" stroke="#FFFFFF" strokeWidth="2"/>
      <path d="M50 4 L92 22 L92 60 C92 82 72 100 50 106 C28 100 8 82 8 60 L8 22 Z" fill="none" stroke="#C9A84C" strokeWidth="1.5"/>
      <path d="M18 28 L82 28 L82 50 L18 50 Z" fill="#0D0D0D"/>
      <text x="50" y="46" textAnchor="middle" fill="#FFFFFF" fontSize="13" fontWeight="900" fontFamily="Arial Black, sans-serif">GFBPA</text>
      <path d="M8 52 L92 52 L92 62 L8 62 Z" fill="#FFFFFF"/>
      <path d="M8 62 L92 62 L92 60 C92 82 72 100 50 106 C28 100 8 82 8 60 Z" fill="#0055A5"/>
    </svg>
  );
}

function MotivoBadge({motivo}){
  const paleta={
    "Criança":"#E3F2FD,#1565C0","Sócio inadimplente":"#FFF3E0,#E65100",
    "Transferência":"#E8F5E9,#2E7D32","Setor incorreto":"#FFF8E1,#F57F17",
    "Acesso já utilizado":"#FCE4EC,#C62828","Responsável já acessou":"#F3E5F5,#6A1B9A",
    "Catraca com defeito":"#EFEBE9,#4E342E",
  };
  const [bg,cor]=(paleta[motivo]||"#F5F5F5,#424242").split(",");
  return <span style={{background:bg,color:cor,border:"1px solid "+cor+"44",padding:"2px 9px",borderRadius:99,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{motivo}</span>;
}

function RegistroCard({r,onDelete}){
  const abrev=r.portao.replace("Portão ","").replace("Toten ","T.").slice(0,4);
  return(
    <div style={{background:G.branco,borderRadius:10,border:"1.5px solid "+G.cinzaBorda,padding:"14px 16px",marginBottom:8,display:"flex",gap:12,alignItems:"flex-start",boxShadow:"0 1px 6px rgba(0,85,165,0.07)"}}>
      <div style={{minWidth:40,height:40,borderRadius:8,background:G.azul,display:"flex",alignItems:"center",justifyContent:"center",color:G.branco,fontWeight:900,fontSize:12,textAlign:"center",lineHeight:1.2,padding:2}}>{abrev}</div>
      <div style={{flex:1,minWidth:0}}>
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:5}}>
          <MotivoBadge motivo={r.motivo}/>
          <span style={{fontSize:11,color:G.cinzaTexto}}>{r.data} · {r.hora}</span>
        </div>
        <div style={{fontSize:13,color:"#1a1a1a",marginBottom:2}}>
          <span style={{fontWeight:700,color:G.azulEscuro}}>CPF:</span> <span style={{fontFamily:"monospace",letterSpacing:0.5}}>{r.cpf}</span>
          <span style={{margin:"0 7px",color:G.cinzaBorda}}>|</span>
          <span style={{fontWeight:700,color:G.azulEscuro}}>Portão:</span> {r.portao}
        </div>
        <div style={{fontSize:12,color:G.cinzaTexto}}><span style={{fontWeight:700}}>Líder:</span> {r.lider}</div>
      </div>
      <button onClick={()=>onDelete(r.id)} style={{background:"none",border:"none",color:G.cinzaBorda,cursor:"pointer",fontSize:20,lineHeight:1,padding:2}} onMouseEnter={e=>e.target.style.color="#DC2626"} onMouseLeave={e=>e.target.style.color=G.cinzaBorda}>×</button>
    </div>
  );
}

// ── TELA DE LOGIN ────────────────────────────────────────────────────────────
function TelaLogin({onLogin}){
  const [pin, setPin] = useState("");
  const [erro, setErro] = useState(false);
  const [tentando, setTentando] = useState(false);

  const handleLogin = () => {
    setTentando(true);
    setTimeout(()=>{
      if(pin.toUpperCase() === PIN_CORRETO){
        sessionStorage.setItem("gremio_auth","1");
        onLogin();
      } else {
        setErro(true);
        setPin("");
        setTimeout(()=>setErro(false), 2500);
      }
      setTentando(false);
    }, 400);
  };

  return(
    <div style={{minHeight:"100vh",background:"linear-gradient(160deg,"+G.preto+" 0%,"+G.azulEscuro+" 50%,"+G.azul+" 100%)",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24,fontFamily:"'Barlow','Segoe UI',sans-serif"}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;600;700;900&family=Barlow+Condensed:wght@700;900&family=DM+Mono:wght@500&display=swap');@keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-8px)}40%,80%{transform:translateX(8px)}}`}</style>

      {/* Faixa tricolor topo */}
      <div style={{position:"fixed",top:0,left:0,right:0,height:5,background:"linear-gradient(90deg,"+G.preto+" 33%,"+G.azul+" 33%,"+G.azul+" 66%,"+G.branco+" 66%)"}}/>

      <div style={{width:"100%",maxWidth:360,textAlign:"center"}}>
        {/* Escudo */}
        <div style={{marginBottom:24,filter:"drop-shadow(0 8px 24px rgba(0,0,0,0.5))"}}>
          <EscudoGremio size={90}/>
        </div>

        <div style={{fontFamily:"'Barlow Condensed',sans-serif",fontSize:28,fontWeight:900,color:G.branco,letterSpacing:2,textTransform:"uppercase",lineHeight:1,marginBottom:4}}>Grêmio FBPA</div>
        <div style={{fontSize:13,color:"#94BDDF",marginBottom:36,fontWeight:600,letterSpacing:0.5}}>Controle de Atendimentos</div>

        <div style={{background:"rgba(255,255,255,0.07)",borderRadius:16,border:"1px solid rgba(255,255,255,0.12)",padding:"28px 24px",backdropFilter:"blur(10px)"}}>
          <div style={{fontSize:13,fontWeight:800,color:"#94BDDF",marginBottom:14,textTransform:"uppercase",letterSpacing:1}}>🔒 Acesso Restrito</div>

          <input
            type="password"
            placeholder="Digite o PIN de acesso"
            value={pin}
            onChange={e=>setPin(e.target.value)}
            onKeyDown={e=>e.key==="Enter"&&handleLogin()}
            style={{
              width:"100%", padding:"13px 16px", borderRadius:9, fontSize:16,
              border:"1.5px solid "+(erro?"#F87171":"rgba(255,255,255,0.2)"),
              background:"rgba(255,255,255,0.1)", color:G.branco,
              fontFamily:"'DM Mono',monospace", letterSpacing:3, textAlign:"center",
              outline:"none", boxSizing:"border-box", marginBottom:14,
              animation: erro?"shake 0.4s ease":"none",
              transition:"border-color 0.2s",
            }}
            autoComplete="off"
          />

          {erro&&(
            <div style={{fontSize:12,color:"#F87171",marginBottom:12,fontWeight:700}}>
              ❌ PIN incorreto. Tente novamente.
            </div>
          )}

          <button
            onClick={handleLogin}
            disabled={tentando||pin.length===0}
            style={{
              width:"100%", padding:"13px", borderRadius:9,
              background:pin.length===0?"rgba(255,255,255,0.1)":"linear-gradient(135deg,"+G.azulEscuro+","+G.azulMedio+")",
              color:pin.length===0?"#64748b":G.branco,
              border:"none", fontSize:15, fontWeight:900, cursor:pin.length===0?"not-allowed":"pointer",
              fontFamily:"'Barlow Condensed',sans-serif", textTransform:"uppercase", letterSpacing:1,
              boxShadow:pin.length>0?"0 4px 16px rgba(0,85,165,0.4)":"none",
              transition:"all 0.2s",
            }}
          >
            {tentando?"Verificando...":"Entrar →"}
          </button>
        </div>

        <div style={{marginTop:20,fontSize:11,color:"#475569",fontWeight:600}}>
          Arena do Grêmio · Operação de Acesso
        </div>
      </div>
    </div>
  );
}

// ── APP PRINCIPAL ────────────────────────────────────────────────────────────
export default function App(){
  const [autenticado, setAutenticado] = useState(()=> sessionStorage.getItem("gremio_auth")==="1");
  const [tab, setTab] = useState("form");
  const [registros, setRegistros] = useState([]);
  const [form, setForm] = useState({motivo:"",portao:"",lider:"",cpf:""});
  const [errors, setErrors] = useState({});
  const [success, setSuccess] = useState(false);
  const [filtro, setFiltro] = useState("");
  const [exportando, setExportando] = useState(false);
  const [loading, setLoading] = useState(true);
  const [salvando, setSalvando] = useState(false);
  const [dbErro, setDbErro] = useState(null);

  const carregarRegistros = useCallback(async()=>{
    setLoading(true);
    const {data,error} = await supabase.from("atendimentos").select("*").order("created_at",{ascending:false});
    if(error){ setDbErro("Erro ao carregar: "+error.message); }
    else { setRegistros(data.map(r=>({id:r.id,cpf:r.cpf,motivo:r.motivo,portao:r.portao,lider:r.lider,hora:r.hora,data:r.data}))); }
    setLoading(false);
  },[]);

  useEffect(()=>{ if(autenticado) carregarRegistros(); },[autenticado,carregarRegistros]);

  useEffect(()=>{
    if(!autenticado) return;
    const channel = supabase.channel("atendimentos-rt")
      .on("postgres_changes",{event:"INSERT",schema:"public",table:"atendimentos"},payload=>{
        const r=payload.new;
        setRegistros(prev=>prev.find(x=>x.id===r.id)?prev:[{id:r.id,cpf:r.cpf,motivo:r.motivo,portao:r.portao,lider:r.lider,hora:r.hora,data:r.data},...prev]);
      })
      .on("postgres_changes",{event:"DELETE",schema:"public",table:"atendimentos"},payload=>{
        setRegistros(prev=>prev.filter(x=>x.id!==payload.old.id));
      })
      .subscribe();
    return ()=>supabase.removeChannel(channel);
  },[autenticado]);

  const validate=()=>{
    const e={};
    if(!form.motivo) e.motivo="Selecione o motivo";
    if(!form.portao) e.portao="Selecione o portão";
    if(!form.lider) e.lider="Selecione o líder";
    if(!validateCPF(form.cpf)) e.cpf="CPF inválido";
    return e;
  };

  const handleSubmit=async()=>{
    const e=validate(); setErrors(e);
    if(Object.keys(e).length>0) return;
    setSalvando(true);
    const now=new Date();
    const hora=now.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"});
    const data=now.toLocaleDateString("pt-BR");
    const {error}=await supabase.from("atendimentos").insert([{cpf:form.cpf,motivo:form.motivo,portao:form.portao,lider:form.lider,hora,data}]);
    setSalvando(false);
    if(error){setDbErro("Erro ao salvar: "+error.message);return;}
    setForm({motivo:"",portao:"",lider:"",cpf:""});
    setSuccess(true); setTimeout(()=>setSuccess(false),2800);
  };

  const handleDelete=async(id)=>{
    const {error}=await supabase.from("atendimentos").delete().eq("id",id);
    if(error) setDbErro("Erro ao excluir: "+error.message);
  };

  const filtered=useMemo(()=>{
    if(!filtro.trim()) return registros;
    const q=filtro.toLowerCase();
    return registros.filter(r=>r.cpf.includes(q)||r.motivo.toLowerCase().includes(q)||r.portao.toLowerCase().includes(q)||r.lider.toLowerCase().includes(q));
  },[registros,filtro]);

  const stats=useMemo(()=>{
    const hoje=new Date().toLocaleDateString("pt-BR");
    const hc=registros.filter(r=>r.data===hoje).length;
    const pp=registros.reduce((acc,r)=>{acc[r.portao]=(acc[r.portao]||0)+1;return acc;},{});
    const tp=Object.entries(pp).sort((a,b)=>b[1]-a[1])[0];
    return{total:registros.length,hoje:hc,top:tp?tp[0].replace("Portão ","P."):"—"};
  },[registros]);

  const resumo=useMemo(()=>{
    const porMotivo=registros.reduce((acc,r)=>{acc[r.motivo]=(acc[r.motivo]||0)+1;return acc;},{});
    const porPortao=registros.reduce((acc,r)=>{acc[r.portao]=(acc[r.portao]||0)+1;return acc;},{});
    const porLider=registros.reduce((acc,r)=>{acc[r.lider]=(acc[r.lider]||0)+1;return acc;},{});
    return{porMotivo,porPortao,porLider};
  },[registros]);

  const exportarExcel=()=>{
    if(registros.length===0) return;
    setExportando(true);
    setTimeout(()=>{
      try{
        const wb=XLSX.utils.book_new();
        const ws1=XLSX.utils.json_to_sheet(registros.map((r,i)=>({"#":i+1,"Data":r.data,"Hora":r.hora,"CPF do Torcedor":r.cpf,"Motivo da Liberação":r.motivo,"Portão":r.portao,"Líder Responsável":r.lider})));
        ws1["!cols"]=[{wch:4},{wch:12},{wch:8},{wch:18},{wch:25},{wch:22},{wch:20}];
        XLSX.utils.book_append_sheet(wb,ws1,"Registros");
        const ws2=XLSX.utils.json_to_sheet([...Object.entries(resumo.porMotivo).sort((a,b)=>b[1]-a[1]).map(([k,v])=>({"Motivo":k,"Quantidade":v})),{"Motivo":"TOTAL","Quantidade":registros.length}]);
        ws2["!cols"]=[{wch:28},{wch:12}]; XLSX.utils.book_append_sheet(wb,ws2,"Por Motivo");
        const ws3=XLSX.utils.json_to_sheet([...Object.entries(resumo.porPortao).sort((a,b)=>b[1]-a[1]).map(([k,v])=>({"Portão":k,"Quantidade":v})),{"Portão":"TOTAL","Quantidade":registros.length}]);
        ws3["!cols"]=[{wch:22},{wch:12}]; XLSX.utils.book_append_sheet(wb,ws3,"Por Portão");
        const ws4=XLSX.utils.json_to_sheet([...Object.entries(resumo.porLider).sort((a,b)=>b[1]-a[1]).map(([k,v])=>({"Líder Responsável":k,"Quantidade":v})),{"Líder Responsável":"TOTAL","Quantidade":registros.length}]);
        ws4["!cols"]=[{wch:24},{wch:12}]; XLSX.utils.book_append_sheet(wb,ws4,"Por Líder");
        const now=new Date();
        XLSX.writeFile(wb,"gremio_atendimentos_"+now.toLocaleDateString("pt-BR").replace(/\//g,"-")+".xlsx");
      }finally{setExportando(false);}
    },100);
  };

  if(!autenticado) return <TelaLogin onLogin={()=>setAutenticado(true)}/>;

  const inp=(err)=>({width:"100%",padding:"10px 12px",borderRadius:7,fontSize:14,border:"1.5px solid "+(err?"#DC2626":G.cinzaBorda),background:G.branco,color:G.preto,outline:"none",fontFamily:"inherit",appearance:"none",backgroundImage:"url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%230055A5' stroke-width='2' fill='none' stroke-linecap='round'/%3E%3C/svg%3E\")",backgroundRepeat:"no-repeat",backgroundPosition:"right 12px center"});
  const lbl={fontSize:11,fontWeight:800,color:G.azulEscuro,marginBottom:5,display:"block",letterSpacing:0.8,textTransform:"uppercase"};
  const err={fontSize:11,color:"#DC2626",marginTop:4};
  const TABS=[["form","✏️ Registrar"],["lista","📋 Registros ("+registros.length+")"],["relatorio","📊 Relatório"]];

  return(
    <div style={{minHeight:"100vh",background:G.cinzaClaro,fontFamily:"'Barlow','Segoe UI',sans-serif"}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;600;700;900&family=Barlow+Condensed:wght@700;900&family=DM+Mono:wght@500&display=swap');@keyframes fadeIn{from{opacity:0;transform:translateY(5px)}to{opacity:1;transform:none}}@keyframes slideDown{from{opacity:0;transform:translateY(-10px)}to{opacity:1;transform:none}}@keyframes spin{to{transform:rotate(360deg)}}select:focus,input:focus{border-color:${G.azul}!important;box-shadow:0 0 0 3px rgba(0,85,165,0.15)!important;outline:none}optgroup{font-weight:800;color:${G.azulEscuro};font-size:12px}::-webkit-scrollbar{width:6px}::-webkit-scrollbar-thumb{background:${G.azulMedio};border-radius:99px}`}</style>

      {/* HEADER */}
      <div style={{background:"linear-gradient(135deg,"+G.preto+" 0%,"+G.azulEscuro+" 40%,"+G.azul+" 100%)",padding:"0 20px",boxShadow:"0 4px 20px rgba(0,0,0,0.4)"}}>
        <div style={{height:4,background:"linear-gradient(90deg,"+G.preto+" 33%,"+G.azul+" 33%,"+G.azul+" 66%,"+G.branco+" 66%)"}}/>
        <div style={{maxWidth:680,margin:"0 auto"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"14px 0 10px"}}>
            <div style={{display:"flex",alignItems:"center",gap:14}}>
              <EscudoGremio size={48}/>
              <div>
                <div style={{fontSize:20,fontWeight:900,color:G.branco,fontFamily:"'Barlow Condensed',sans-serif",letterSpacing:1,textTransform:"uppercase",lineHeight:1}}>Grêmio FBPA</div>
                <div style={{fontSize:11,color:"#94BDDF",marginTop:2,fontWeight:600}}>Controle de Atendimentos · Arena do Grêmio</div>
              </div>
            </div>
            <button onClick={()=>{sessionStorage.removeItem("gremio_auth");setAutenticado(false);}} title="Sair" style={{background:"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.2)",color:"#94BDDF",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"inherit"}}>🔒 Sair</button>
          </div>

          <div style={{display:"flex",gap:10,paddingBottom:12}}>
            {[{label:"Total",value:stats.total,icon:"📋"},{label:"Hoje",value:stats.hoje,icon:"📅"},{label:"Top Portão",value:stats.top,icon:"🚪"}].map(s=>(
              <div key={s.label} style={{flex:1,background:"rgba(255,255,255,0.08)",borderRadius:9,padding:"9px 10px",textAlign:"center",border:"1px solid rgba(255,255,255,0.12)"}}>
                <div style={{fontSize:16}}>{s.icon}</div>
                <div style={{fontSize:20,fontWeight:900,color:G.branco,lineHeight:1.1,fontFamily:"'Barlow Condensed',sans-serif"}}>{s.value}</div>
                <div style={{fontSize:9,color:"#7AAED4",marginTop:1,textTransform:"uppercase",letterSpacing:0.6,fontWeight:700}}>{s.label}</div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:2}}>
            {TABS.map(([t,label])=>(
              <button key={t} onClick={()=>setTab(t)} style={{padding:"9px 16px",borderRadius:"8px 8px 0 0",border:"none",cursor:"pointer",fontFamily:"'Barlow',sans-serif",fontWeight:700,fontSize:13,background:tab===t?G.cinzaClaro:"transparent",color:tab===t?G.azulEscuro:"#94BDDF",transition:"all 0.2s"}}>{label}</button>
            ))}
          </div>
        </div>
      </div>

      {/* CONTENT */}
      <div style={{maxWidth:680,margin:"0 auto",padding:"20px 16px 50px"}}>
        {dbErro&&(<div style={{background:"#FEF2F2",border:"1.5px solid #FECACA",color:"#991B1B",borderRadius:9,padding:"11px 16px",marginBottom:16,fontWeight:600,fontSize:13,display:"flex",justifyContent:"space-between",alignItems:"center"}}>⚠️ {dbErro}<button onClick={()=>setDbErro(null)} style={{background:"none",border:"none",cursor:"pointer",color:"#991B1B",fontSize:18}}>×</button></div>)}
        {success&&(<div style={{background:"#EBF8EE",border:"1.5px solid #4ADE80",color:"#166534",borderRadius:9,padding:"11px 16px",marginBottom:16,fontWeight:700,fontSize:14,display:"flex",alignItems:"center",gap:9,animation:"slideDown 0.3s ease"}}><span style={{fontSize:18}}>⚽</span> Atendimento salvo com sucesso!</div>)}

        {/* FORM */}
        {tab==="form"&&(
          <div style={{background:G.branco,borderRadius:12,border:"1.5px solid "+G.cinzaBorda,padding:"22px 20px",boxShadow:"0 2px 12px rgba(0,85,165,0.09)",animation:"fadeIn 0.3s ease"}}>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:20,paddingBottom:14,borderBottom:"2px solid "+G.cinzaClaro}}>
              <div style={{width:4,height:22,background:G.azul,borderRadius:2}}/>
              <span style={{fontSize:15,fontWeight:900,color:G.azulEscuro,fontFamily:"'Barlow Condensed',sans-serif",textTransform:"uppercase",letterSpacing:1}}>Novo Atendimento</span>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              <div style={{gridColumn:"1 / -1"}}>
                <label style={lbl}>Motivo da Liberação *</label>
                <select style={inp(errors.motivo)} value={form.motivo} onChange={e=>setForm(f=>({...f,motivo:e.target.value}))}>
                  <option value="">Selecione o motivo...</option>
                  {MOTIVOS_LIBERACAO.map(m=><option key={m} value={m}>{m}</option>)}
                </select>
                {errors.motivo&&<div style={err}>{errors.motivo}</div>}
              </div>
              <div style={{gridColumn:"1 / -1"}}>
                <label style={lbl}>Portão / Setor *</label>
                <select style={inp(errors.portao)} value={form.portao} onChange={e=>setForm(f=>({...f,portao:e.target.value}))}>
                  <option value="">Selecione o portão...</option>
                  {SETORES_PORTOES.map(s=><optgroup key={s.setor} label={"— "+s.setor}>{s.portoes.map(p=><option key={p} value={p}>{p}</option>)}</optgroup>)}
                </select>
                {errors.portao&&<div style={err}>{errors.portao}</div>}
              </div>
              <div>
                <label style={lbl}>Líder Responsável *</label>
                <select style={inp(errors.lider)} value={form.lider} onChange={e=>setForm(f=>({...f,lider:e.target.value}))}>
                  <option value="">Selecione...</option>
                  {LIDERES.map(l=><option key={l} value={l}>{l}</option>)}
                </select>
                {errors.lider&&<div style={err}>{errors.lider}</div>}
              </div>
              <div>
                <label style={lbl}>CPF do Torcedor *</label>
                <input type="text" placeholder="000.000.000-00" value={form.cpf} onChange={e=>setForm(f=>({...f,cpf:formatCPF(e.target.value)}))} style={{...inp(errors.cpf),fontFamily:"'DM Mono',monospace",letterSpacing:1.2,backgroundImage:"none"}}/>
                {errors.cpf&&<div style={err}>{errors.cpf}</div>}
              </div>
            </div>
            <button onClick={handleSubmit} disabled={salvando} style={{width:"100%",marginTop:20,padding:"13px",background:salvando?"#94a3b8":"linear-gradient(135deg,"+G.azulEscuro+","+G.azul+")",color:G.branco,border:"none",borderRadius:9,fontSize:15,fontWeight:900,cursor:salvando?"not-allowed":"pointer",letterSpacing:0.8,fontFamily:"'Barlow Condensed',sans-serif",textTransform:"uppercase",boxShadow:"0 3px 12px rgba(0,85,165,0.35)"}}>
              {salvando?"⏳ Salvando...":"⚽ Registrar Atendimento"}
            </button>
          </div>
        )}

        {/* LISTA */}
        {tab==="lista"&&(
          <div style={{animation:"fadeIn 0.3s ease"}}>
            <div style={{display:"flex",gap:10,marginBottom:14,alignItems:"center"}}>
              <input type="text" placeholder="🔍  Buscar por CPF, motivo, portão ou líder..." value={filtro} onChange={e=>setFiltro(e.target.value)} style={{flex:1,padding:"11px 14px",borderRadius:9,fontSize:14,border:"1.5px solid "+G.cinzaBorda,background:G.branco,fontFamily:"inherit",outline:"none",boxSizing:"border-box"}}/>
              <button onClick={carregarRegistros} title="Atualizar" style={{padding:"11px 13px",borderRadius:9,border:"1.5px solid "+G.cinzaBorda,background:G.branco,cursor:"pointer",fontSize:16,color:G.azul}}>🔄</button>
            </div>
            {loading?(
              <div style={{textAlign:"center",padding:"40px",color:G.cinzaTexto}}>
                <div style={{fontSize:32,marginBottom:8,animation:"spin 1s linear infinite",display:"inline-block"}}>⚙️</div>
                <div style={{fontWeight:700}}>Carregando registros...</div>
              </div>
            ):filtered.length===0?(
              <div style={{textAlign:"center",padding:"44px 20px",color:G.cinzaTexto}}>
                <div style={{fontSize:42,marginBottom:10}}>📭</div>
                <div style={{fontWeight:800,fontSize:15,color:G.azulEscuro}}>{registros.length===0?"Nenhum registro ainda":"Sem resultados"}</div>
                <div style={{fontSize:13,marginTop:4}}>{registros.length===0?"Use o formulário para registrar atendimentos":"Tente outro termo de busca"}</div>
              </div>
            ):(
              <>
                <div style={{fontSize:12,color:G.cinzaTexto,marginBottom:10,fontWeight:700,textTransform:"uppercase",letterSpacing:0.5}}>{filtered.length} registro{filtered.length!==1?"s":""}</div>
                {filtered.map(r=><RegistroCard key={r.id} r={r} onDelete={handleDelete}/>)}
              </>
            )}
          </div>
        )}

        {/* RELATÓRIO */}
        {tab==="relatorio"&&(
          <div style={{animation:"fadeIn 0.3s ease"}}>
            <div style={{background:G.branco,borderRadius:12,border:"1.5px solid "+G.cinzaBorda,padding:"18px 20px",marginBottom:14,boxShadow:"0 2px 12px rgba(0,85,165,0.09)"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
                <div>
                  <div style={{fontSize:15,fontWeight:900,color:G.azulEscuro,fontFamily:"'Barlow Condensed',sans-serif",textTransform:"uppercase",letterSpacing:0.5}}>Exportar Relatório Excel</div>
                  <div style={{fontSize:12,color:G.cinzaTexto,marginTop:3}}>{registros.length} registro{registros.length!==1?"s":""} de todos os líderes · 4 abas</div>
                </div>
                <button onClick={exportarExcel} disabled={registros.length===0||exportando} style={{background:registros.length===0?G.cinzaClaro:"linear-gradient(135deg,"+G.azulEscuro+","+G.azul+")",color:registros.length===0?G.cinzaTexto:G.branco,border:"none",borderRadius:9,padding:"10px 18px",fontSize:13,fontWeight:800,cursor:registros.length===0?"not-allowed":"pointer",fontFamily:"'Barlow',sans-serif",whiteSpace:"nowrap",boxShadow:registros.length>0?"0 3px 10px rgba(0,85,165,0.3)":"none"}}>
                  {exportando?"⏳ Gerando...":"⬇️ Baixar Excel"}
                </button>
              </div>
              {registros.length===0&&<div style={{marginTop:12,fontSize:12,color:"#92400E",background:"#FEF3C7",border:"1px solid #FDE68A",borderRadius:7,padding:"8px 12px"}}>⚠️ Nenhum atendimento registrado ainda</div>}
            </div>
            {registros.length>0&&(
              <>
                {[
                  {titulo:"Por Motivo de Liberação",dados:resumo.porMotivo,cor:G.azul,renderKey:(k)=><MotivoBadge motivo={k}/>},
                  {titulo:"Por Portão / Setor",dados:resumo.porPortao,cor:G.dourado,renderKey:(k)=><span style={{fontSize:12,color:G.azulEscuro,minWidth:130,fontWeight:700}}>{k}</span>},
                  {titulo:"Por Líder Responsável",dados:resumo.porLider,cor:"#4F46E5",renderKey:(k)=><span style={{fontSize:12,color:G.azulEscuro,minWidth:140,fontWeight:700}}>{k}</span>},
                ].map(({titulo,dados,cor,renderKey})=>(
                  <div key={titulo} style={{background:G.branco,borderRadius:12,border:"1.5px solid "+G.cinzaBorda,padding:"16px 18px",marginBottom:12,boxShadow:"0 1px 6px rgba(0,85,165,0.06)"}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
                      <div style={{width:3,height:16,background:cor,borderRadius:2}}/>
                      <span style={{fontSize:12,fontWeight:800,color:G.azulEscuro,textTransform:"uppercase",letterSpacing:0.6}}>{titulo}</span>
                    </div>
                    {Object.entries(dados).sort((a,b)=>b[1]-a[1]).map(([k,qtd])=>(
                      <div key={k} style={{display:"flex",alignItems:"center",gap:10,marginBottom:9}}>
                        <div style={{minWidth:140}}>{renderKey(k)}</div>
                        <div style={{flex:1,height:6,background:G.cinzaClaro,borderRadius:99,overflow:"hidden"}}>
                          <div style={{height:"100%",background:cor,borderRadius:99,width:((qtd/registros.length)*100)+"%",transition:"width 0.6s ease"}}/>
                        </div>
                        <span style={{fontSize:13,fontWeight:900,color:G.azulEscuro,minWidth:22,textAlign:"right"}}>{qtd}</span>
                      </div>
                    ))}
                  </div>
                ))}
              </>
            )}
            {registros.length===0&&(
              <div style={{textAlign:"center",padding:"44px 20px",color:G.cinzaTexto}}>
                <div style={{fontSize:42,marginBottom:10}}>📊</div>
                <div style={{fontWeight:800,fontSize:15,color:G.azulEscuro}}>Nenhum dado para exibir</div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
