import React from "react";
import { useState, useEffect, useCallback } from "react";
import * as XLSX from 'xlsx';
const C = {
  coral:"#D85A30",coralL:"#FAECE7",coralD:"#993C1D",
  green:"#639922",greenL:"#EAF3DE",greenD:"#27500A",
  amber:"#BA7517",amberL:"#FAEEDA",amberD:"#633806",
  blue:"#185FA5",blueL:"#E6F1FB",blueD:"#0C447C",
  teal:"#1D9E75",tealL:"#E1F5EE",tealD:"#085041",
  purple:"#7F77DD",purpleL:"#EEEDFE",purpleD:"#3C3489",
};

const ETAPAS = [
  {id:"lead",label:"Lead",emoji:"🎯",cor:C.purple,corL:C.purpleL,corD:C.purpleD},
  {id:"primeiro_contato",label:"Primeiro Contato",emoji:"📞",cor:C.blue,corL:C.blueL,corD:C.blueD},
  {id:"em_conversa",label:"Em Conversa",emoji:"💬",cor:C.teal,corL:C.tealL,corD:C.tealD},
  {id:"proposta_feita",label:"Proposta Feita",emoji:"📋",cor:C.amber,corL:C.amberL,corD:C.amberD},
  {id:"convertido",label:"Convertido",emoji:"🏆",cor:C.green,corL:C.greenL,corD:C.greenD},
  {id:"encerrado",label:"Encerrado",emoji:"✗",cor:"#888",corL:"#f5f5f5",corD:"#555"},
];

const LISTAS = ["Triagem Universal","Páscoa — Falta Uma","Páscoa — Reativação","Outra"];

// ── SUPABASE CONFIG (salvo localmente em window.storage — só 2 strings) ────────
const _SB = {url:"", key:""};

const loadCfg = async () => {
  try { const v = localStorage.getItem("sb_cfg"); if(v) Object.assign(_SB, JSON.parse(v)); } catch(e) {}
  return _SB;
};
const saveCfg = async (cfg) => {
  Object.assign(_SB, cfg);
  try { localStorage.setItem("sb_cfg", JSON.stringify(cfg)); } catch(e) {}
};
const isConfigured = () => !!(_SB.url && _SB.key);

// ── SUPABASE REST API ─────────────────────────────────────────────────────────
const sb = async (path, opts={}) => {
  if (!isConfigured()) throw new Error("Supabase nao configurado.");
  const sep = path.includes("?") ? "&" : "?";
  const url = _SB.url + "/rest/v1" + path + sep + "apikey=" + encodeURIComponent(_SB.key);
  const resp = await fetch(url, {
    method: opts.method || "GET",
    headers: {
      "Authorization": "Bearer " + _SB.key,
      "Content-Type": "application/json",
      ...(opts.pref ? {"Prefer": opts.pref} : {}),
    },
    ...(opts.body !== undefined ? {body: JSON.stringify(opts.body)} : {})
  });
  if (!resp.ok) { const t = await resp.text(); throw new Error("Erro " + resp.status + ": " + t.slice(0,120)); }
  const t = await resp.text();
  return t ? JSON.parse(t) : null;
};

const dbGetAll  = async () => { const r = await sb("/clientes?select=dados&order=atualizado_em.desc"); return (r||[]).map(x=>x.dados).filter(Boolean); };
const dbGetAtivos = async () => { 
  // Load only non-closed leads for kanban performance
  const r = await sb("/clientes?select=dados&dados->>etapa=neq.encerrado&dados->>etapa=neq.convertido&order=atualizado_em.desc");
  // Supabase jsonb filter workaround — filter client-side if needed
  const all = (r||[]).map(x=>x.dados).filter(Boolean);
  return all;
};
const dbSave    = async (c)  => { await sb("/clientes", {method:"POST", body:{id:c.id, dados:c, atualizado_em:new Date().toISOString()}, pref:"resolution=merge-duplicates"}); };
const dbBulkSave= async (cs) => { if(!cs.length)return; await sb("/clientes", {method:"POST", body:cs.map(c=>({id:c.id,dados:c,atualizado_em:new Date().toISOString()})), pref:"resolution=merge-duplicates"}); };
const dbDelete  = async (id) => { await sb("/clientes?id=eq."+id, {method:"DELETE"}); };
const dbTest    = async ()   => { await sb("/clientes?limit=1"); return true; };
const dbGetConversoes = async () => {
  try { return await sb("/conversoes?select=resultado,registrado_em,ciclo_medio,pedidos,prob_estimada&order=registrado_em.desc") || []; }
  catch(e) { return []; }
};
const dbSaveConversao = async (cliente, resultado) => {
  const c = {
    id: "conv_" + Date.now(),
    cliente_id: cliente.id,
    resultado,
    objetivo: cliente.objetivo||"",
    ciclo_medio: cliente.cicloMedio||0,
    pedidos: cliente.p||0,
    gasto_total: cliente.gasto||0,
    fora_sp: cliente.fora||false,
    prob_estimada: cliente.prob||0,
    lista: cliente.lista||"",
  };
  try { await sb("/conversoes", {method:"POST", body:c, pref:"resolution=merge-duplicates"}); } catch(e) {}
};

const buildSeq = (obj, ciclo, p, fora, foraDaJanela, diasUnico) => {
  const steps = [];
  const sed = fora ? " (sem SEDEX)" : "";
  if (obj === "reativacao") {
    const d = diasUnico || 0;
    if (d > 30) steps.push({ label:"Antes — Reconexão", quem:"Time humano", cor:C.teal, copy:"Oi [Nome]! 😊 Aqui é o Lucas da Laricas.\n\nTudo bem? Vi que você experimentou a gente — o [produto comprado] acabou ou ainda tinha?", regra:"Não mencionar oferta. Só reconectar.", gatilho:"Positivo → T1. Sem resposta 48h → T1 direto." });
    steps.push({ label:d<=30?"T1 — Curadoria":"T2 — Curadoria", quem:"Time ou automação", cor:C.green, copy:"Oi [Nome]! 😊 Aqui é o Lucas da Laricas.\n\nVi que você experimentou o [produto comprado] — ótima escolha!\n\nTem um que combina muito com o seu perfil: o [próximo sabor]. Quer o link com um cupomzinho?", regra:"Sugestão baseada no produto comprado — nunca genérica.", gatilho:"Sim → T2/T3 com link + VOLTA10." });
    steps.push({ label:d<=30?"T2 — Link":"T3 — Link", quem:"Automação", cor:C.amber, copy:"Aqui está 🎁\n\n[link do produto]\n\nCupom VOLTA10 — válido 5 dias.", regra:"Cupom com prazo real. Link direto do produto.", gatilho:"Compra → ciclo natural. Não compra → encerramento." });
    steps.push({ label:"Encerramento", quem:"Automação", cor:C.blue, copy:"[Nome], o cupom expirou mas continua disponível 😊\n\nPosso te avisar de novidade?", regra:"Transformar em permissão de contato futuro.", gatilho:"Sim → lista. Não → encerrar." });
  } else if (obj === "falta_uma") {
    if (foraDaJanela) steps.push({ label:"Antes — Reconexão", quem:"Time humano", cor:C.amber, copy:"Oi [Nome]! 😊 Aqui é o Lucas da Laricas.\n\nComo tá? Ainda tem Laricas em casa? 😄", regra:"Confirmar engajamento antes de qualquer oferta.", gatilho:"Positivo → T1." });
    steps.push({ label:"T1 — Curadoria", quem:"Time humano", cor:C.green, copy:"Oi [Nome]! 😊 Aqui é o Lucas da Laricas.\n\nVocê já experimentou [produto 1] e [produto 2] — boas escolhas!\n\nQuase todo mundo que experimenta um terceiro encontra o favorito de vez 😄 Tem um que combina muito com você: o [sugestão]. Quer ver?", regra:"Produtos reais do Shopify. Não revelar intenção ainda.", gatilho:"Interesse → T2 com link. Sem resposta → 'pão de mel ou bolinho?'" });
    steps.push({ label:"T2 — Link + cupom", quem:"Automação", cor:C.amber, copy:"Aqui está 🎁\n\n[link direto do produto]\n\nCupom VOLTA10 — válido 5 dias!", regra:"Link direto. Prazo real.", gatilho:"Compra → 3° pedido → triagem Club. Não compra → T3." });
    steps.push({ label:"T3 — Última tentativa", quem:"Automação", cor:C.blue, copy:"[Nome], o cupom expirou — mas o [produto] continua disponível 😊\n\nQual foi seu favorito até agora?", regra:"Encerrar mantendo conversa viva.", gatilho:"Responde → sugestão + última tentativa. Sem resposta → encerrar." });
  } else if (obj === "habit_rebuild") {
    if (foraDaJanela) steps.push({ label:"Antes — Reconexão", quem:"Time humano", cor:C.coral, copy:"Oi [Nome]! 😊 Aqui é o Lucas da Laricas.\n\nVi que você já fez "+p+" pedidos! Faz um tempo do último, tá sem estoque?", regra:"NÃO mencionar Club. Objetivo: próxima compra.", gatilho:"Positivo → T1." });
    steps.push({ label:"T1 — Curadoria", quem:"Time humano", cor:C.amber, copy:"Oi [Nome]! 😊 Aqui é o Lucas da Laricas.\n\nVi que você já pediu "+p+" vezes — claramente gosta! 🙌\n\nQual foi seu favorito? Pergunto porque tem um que você ainda não experimentou.", regra:"Ciclo de "+ciclo+" dias — NÃO oferecer Club ainda.", gatilho:"Favorito → T2 personalizado." });
    steps.push({ label:"T2 — Sugestão", quem:"Automação", cor:C.amber, copy:"[Nome], baseado no que você já provou, o [próximo sabor] vai ser seu próximo favorito 😄\n\n[link] + VOLTA10 — válido 5 dias!", regra:"Objetivo: compra imediata, não Club.", gatilho:"Compra → próximo ciclo. Não compra → encerramento." });
    steps.push({ label:"Encerramento", quem:"Automação", cor:C.blue, copy:"[Nome], o cupom expirou — disponível quando quiser 😊\n\nPosso te avisar de novidade?", regra:"Nunca mencionar Club neste fluxo.", gatilho:"Aceita → avisos futuros." });
  } else {
    if (foraDaJanela && ciclo > 60) steps.push({ label:"Antes — Reativação", quem:"Time humano", cor:C.teal, copy:"[Nome], tudo bem? 😊 Aqui é o Lucas da Laricas.\n\nO estoque acabou ou ainda tinha?", regra:"Não mencionar Club ainda.", gatilho:"Positivo → T1." });
    const ang = p >= 7 ? "surpresa" : ciclo > 60 ? "morno" : p === 3 ? "emocional" : "financeiro";
    if (ang === "emocional") steps.push({ label:"T1 — Reconhecimento", quem:"Time humano", cor:C.green, copy:"Oi [Nome]! 😊 Aqui é o Lucas da Laricas.\n\nVi que você fez seu 3° pedido — fico muito feliz! 🙌\n\nQual foi o produto que você mais gostou?", regra:"Não mencionar Club. A resposta personaliza o T2.", gatilho:"Cita favorito → T2. Neutro → produto mais comprado." });
    else if (ang === "financeiro") steps.push({ label:"T1 — Curiosidade", quem:"Time humano", cor:C.amber, copy:"Oi [Nome]! 😊 Aqui é o Lucas da Laricas.\n\nVocê já fez "+p+" pedidos com a gente 🙌\n\nPosso te mostrar um número que talvez te surpreenda?", regra:"Pedir permissão antes do cálculo.", gatilho:"Sim → T2." });
    else if (ang === "surpresa") steps.push({ label:"T1 — Surpresa", quem:"Time humano", cor:C.purple, copy:"Oi [Nome]! 😊 Aqui é o Lucas da Laricas.\n\nVocê já fez "+p+" pedidos — é uma das nossas clientes mais fiéis 🏆\n\nFui checar e vi que você ainda não tem o Club. Tudo bem?\n\nAcho que ninguém te explicou direito ainda.", regra:"Responsabilidade na marca, não na cliente.", gatilho:"Curiosidade → T2." });
    else steps.push({ label:"T1 — Reconexão", quem:"Time humano", cor:C.amber, copy:"Oi [Nome]! 😊 Aqui é o Lucas da Laricas.\n\nVocê fez seu "+p+"° pedido — que bom te ver! 🙌\n\nO estoque acabou ou ainda tinha?", regra:"Ciclo de "+ciclo+" dias — confirmar engajamento primeiro.", gatilho:"Positivo → T2." });
    steps.push({ label:"T2 — O cálculo", quem:"Automação", cor:C.green, copy:"Nos seus "+p+" pedidos você investiu R$[total]"+sed+".\n\nSó de frete: R$[frete acumulado].\n\nNo Club: R$[preço]/mês · frete zero · [favorito] garantido · pausa quando quiser.\n\nEconomia de R$[diferença]/mês. Faz sentido?", regra:"Números reais do Shopify.", gatilho:"Sim → T3 fechamento. Objeção → T3 objeções." });
    steps.push({ label:"T3 — Objeções / fechamento", quem:"Automação", cor:C.blue, copy:"[Usar conforme resposta]\n\n💸 PREÇO: 'Você pagou R$[frete] só de frete. No Club some.'\n\n📦 ACÚMULO: 'Tem pausa — avisa e pulamos o envio.'\n\n⏳ PENSAR: 'Claro! Kit do próximo ciclo fecha [data].'\n\n✅ FECHAR: '[link do Club] — é rapidinho 🎉'", regra:"Urgência real. Nunca pressionar.", gatilho:"Converteu → boas-vindas. Não → aceitar e manter canal." });
  }
  return steps;
};

const calcProb = (obj, ciclo, p, fora, foraDaJanela, gasto, diasUnico) => {
  let base = 20, mP = 1, mR = 1, mL = fora ? 1.2 : 1, mG = 1;
  if (obj === "reativacao") { const d = diasUnico||0; base = d<=30?35:d<=60?22:12; }
  else if (obj === "falta_uma") { base = ciclo<=60?42:28; mR = foraDaJanela?0.75:1; }
  else if (obj === "habit_rebuild") { base = 18; mP = p>=7?1.3:p>=5?1.15:1; mR = foraDaJanela?0.7:1; }
  else { base = ciclo<=60?40:ciclo<=90?25:15; mP = p===3?1:p<=6?1.15:1.4; mR = foraDaJanela?0.72:1; }
  if (gasto>=1000) mG=1.35; else if (gasto>=500) mG=1.2; else if (gasto>=200) mG=1.05; else if (gasto>0) mG=0.85;
  const pct = Math.min(Math.round(base*mP*mR*mL*mG), 72);
  const cor = pct>=40?C.green:pct>=25?C.amber:C.coral;
  const corD = pct>=40?C.greenD:pct>=25?C.amberD:C.coralD;
  const corL = pct>=40?C.greenL:pct>=25?C.amberL:C.coralL;
  return { pct, label:pct>=40?"Alta":pct>=25?"Média":"Baixa", cor, corD, corL };
};

const runTriagem = (pedidos, dp, du, fora, gasto) => {
  const hoje = new Date();
  const p = parseInt(pedidos)||0;
  let ciclo=0, diasUlt=0, diasUnico=0, span=0, foraDaJanela=false;
  if (p===1) { const dt=new Date(dp); diasUnico=Math.round((hoje-dt)/86400000); foraDaJanela=diasUnico>30; }
  else { const dtP=new Date(dp),dtU=new Date(du); diasUlt=Math.round((hoje-dtU)/86400000); span=Math.round((dtU-dtP)/86400000); ciclo=p>1?Math.round(span/(p-1)):span; foraDaJanela=diasUlt>30; }
  let obj,label,cor,corD,alerta="";
  if (p===1) { obj="reativacao"; label="Reativação → 2ª compra"; cor=C.teal; corD=C.tealD; }
  else if (p===2) { if (ciclo<=60) { obj="falta_uma"; label="Falta Uma → 3ª compra"; cor=C.amber; corD=C.amberD; } else { obj="reativacao"; label="Reativar → 3ª compra"; cor=C.teal; corD=C.tealD; alerta="Ciclo "+ciclo+"d — reconectar antes de empurrar para o 3°."; } }
  else { if (ciclo<=90) { obj="club"; label=ciclo<=60?"Club — hábito formado":"Club — abordagem suave"; cor=C.green; corD=C.greenD; if(ciclo>60) alerta="Ciclo "+ciclo+"d — usar ângulo de conveniência."; } else { obj="habit_rebuild"; label="Reconstruir hábito → Club só depois"; cor=C.coral; corD=C.coralD; alerta="⛔ Ciclo "+ciclo+"d — NÃO oferecer Club agora."; } }
  const seq = buildSeq(obj, ciclo, p, fora, foraDaJanela, diasUnico);
  const prob = calcProb(obj, ciclo, p, fora, foraDaJanela, gasto, diasUnico);
  return { obj, label, cor, corD, alerta, ciclo, p, fora, foraDaJanela, diasUlt, diasUnico, span, seq, prob };
};

const inp = (ex) => ({ width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:14,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none",...ex });
const T = ({ label, active, color, onClick }) => ( <button onClick={onClick} style={{ padding:"8px 12px",fontSize:12,fontWeight:500,color:active?color:"var(--color-text-secondary)",borderBottom:active?"2px solid "+color:"2px solid transparent",marginBottom:-1,background:"transparent",border:"none",cursor:"pointer",whiteSpace:"nowrap" }}>{label}</button> );
const M = ({ label, value, sub, cor }) => ( <div style={{ background:"var(--color-background-secondary)",borderRadius:8,padding:"10px 12px" }}><div style={{ fontSize:10,color:"var(--color-text-tertiary)",marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em" }}>{label}</div><div style={{ fontSize:18,fontWeight:500,color:cor||"var(--color-text-primary)" }}>{value}</div>{sub&&<div style={{ fontSize:10,color:"var(--color-text-tertiary)",marginTop:2 }}>{sub}</div>}</div> );

const Steps = ({ steps, cur }) => {
  const [open, setOpen] = useState(cur||0);
  return (
    <div>
      {steps.map((s, i) => {
        const isCur = i === cur;
        return (
          <div key={i} style={{ marginBottom:8 }}>
            <button onClick={() => setOpen(open===i?-1:i)} style={{ width:"100%",textAlign:"left",background:isCur?C.tealL:open===i?"var(--color-background-secondary)":"var(--color-background-primary)",border:"0.5px solid "+(isCur?C.teal:open===i?s.cor:"var(--color-border-tertiary)"),borderLeft:"3px solid "+(isCur?C.teal:s.cor),borderRadius:open===i?"10px 10px 0 0":10,padding:"10px 14px",cursor:"pointer" }}>
              <div style={{ display:"flex",alignItems:"center",gap:8 }}>
                <div style={{ width:20,height:20,borderRadius:"50%",background:(isCur?C.teal:s.cor)+"22",color:isCur?C.teal:s.cor,fontSize:10,fontWeight:500,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0 }}>{i+1}</div>
                <div style={{ fontSize:13,fontWeight:500,color:isCur?C.tealD:"var(--color-text-primary)",flex:1 }}>{s.label}</div>
                {isCur&&<span style={{ fontSize:10,background:C.teal,color:"#fff",padding:"2px 7px",borderRadius:20 }}>Atual</span>}
                <span style={{ fontSize:11,background:"var(--color-background-secondary)",color:"var(--color-text-secondary)",padding:"2px 8px",borderRadius:20 }}>{s.quem}</span>
                <span style={{ fontSize:12,color:"var(--color-text-tertiary)" }}>{open===i?"▲":"▼"}</span>
              </div>
            </button>
            {open===i&&(
              <div style={{ border:"0.5px solid "+(isCur?C.teal:s.cor),borderTop:"none",borderRadius:"0 0 10px 10px",padding:"14px",background:"var(--color-background-primary)" }}>
                <div style={{ background:"var(--color-background-secondary)",borderRadius:10,padding:"13px 15px",marginBottom:10,fontSize:14,color:"var(--color-text-primary)",lineHeight:1.85,whiteSpace:"pre-line",fontFamily:"inherit",borderLeft:"3px solid "+(isCur?C.teal:s.cor) }}>{s.copy}</div>
                <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8 }}>
                  <div><div style={{ fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Regra</div><div style={{ fontSize:12,color:"var(--color-text-secondary)",lineHeight:1.5 }}>{s.regra}</div></div>
                  <div><div style={{ fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Próximo gatilho</div><div style={{ fontSize:12,color:"var(--color-text-secondary)",lineHeight:1.5 }}>{s.gatilho}</div></div>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
};

const ProbBar = ({ prob }) => (
  <div style={{ background:prob.corL,border:"1px solid "+prob.cor,borderRadius:12,padding:"12px 16px",marginBottom:12 }}>
    <div style={{ display:"flex",alignItems:"center",gap:12 }}>
      <div style={{ flex:1 }}>
        <div style={{ fontSize:10,color:prob.corD,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:4 }}>Probabilidade de conversão</div>
        <div style={{ height:4,background:prob.cor+"28",borderRadius:3,overflow:"hidden" }}><div style={{ width:prob.pct+"%",height:"100%",background:prob.cor,borderRadius:3 }} /></div>
      </div>
      <div style={{ textAlign:"right",flexShrink:0 }}>
        <div style={{ fontSize:30,fontWeight:500,color:prob.corD,lineHeight:1 }}>{prob.pct}%</div>
        <div style={{ fontSize:12,fontWeight:500,color:prob.cor }}>{prob.label}</div>
      </div>
    </div>
  </div>
);

const TriagemForm = ({ onSalvo, lista }) => {
  const [nome,setNome]=useState(""); const [tel,setTel]=useState(""); const [ped,setPed]=useState("");
  const [dp,setDp]=useState(""); const [du,setDu]=useState(""); const [fora,setFora]=useState(null);
  const [gasto,setGasto]=useState(""); const [res,setRes]=useState(null); const [salvo,setSalvo]=useState(false);
  const p = parseInt(ped)||0;
  const pronto = ped && dp && fora!==null && p>=1 && (p===1||(du&&new Date(du)>=new Date(dp)));
  const calcular = () => { setRes(runTriagem(ped, dp, p===1?dp:du, fora, parseFloat(gasto)||0)); };
  const salvar = () => {
    if (!res||!nome.trim()) return;
    const c = { id:"c_"+Date.now(), etapa:"lead", dataCriacao:new Date().toLocaleDateString("pt-BR"), notas:"", proximaAcao:"", dataProximoContato:"", lista:lista||"", nome:nome.trim(), telefone:tel.trim(), objetivo:res.obj, objetivoLabel:res.label, objetivoCor:res.cor, objetivoCorD:res.corD, prob:res.prob.pct, probLabel:res.prob.label, probCor:res.prob.cor, seq:res.seq, stepAtual:0, cicloMedio:res.ciclo, p:res.p, fora:res.fora, datasPreenchidas:true };
    dbSave(c).then(() => { setSalvo(true); setTimeout(() => { setNome("");setTel("");setPed("");setDp("");setDu("");setFora(null);setGasto("");setRes(null);setSalvo(false); onSalvo&&onSalvo(); }, 1500); }).catch(e => { setSalvo(false); alert("Erro ao salvar: " + e.message); });
  };
  return (
    <div>
      <div style={{ marginBottom:12 }}><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Nome <span style={{ color:C.coralD,fontSize:10 }}>*obrigatório</span></div><input style={inp()} type="text" placeholder="Ex: Maria Silva" value={nome} onChange={e=>setNome(e.target.value)} /></div>
      <div style={{ marginBottom:12 }}><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Telefone / WhatsApp</div><input style={inp()} type="text" placeholder="11 9XXXX-XXXX" value={tel} onChange={e=>setTel(e.target.value)} /></div>
      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12 }}>
        <div><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Total de pedidos</div><input style={inp()} type="number" min="1" placeholder="1, 2, 3..." value={ped} onChange={e=>{setPed(e.target.value);setRes(null);}} /></div>
        <div><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Total gasto R$ (opcional)</div><input style={inp()} type="number" min="0" placeholder="450..." value={gasto} onChange={e=>{setGasto(e.target.value);setRes(null);}} /></div>
      </div>
      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12 }}>
        <div><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Data do 1° pedido</div><input style={inp()} type="date" value={dp} onChange={e=>{setDp(e.target.value);setRes(null);}} /></div>
        <div><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Data do último {p===1&&<span style={{ fontWeight:400,fontSize:10 }}>(= 1°)</span>}</div><input style={inp({opacity:p===1?0.5:1})} type="date" value={p===1?dp:du} onChange={e=>{if(p>1){setDu(e.target.value);setRes(null);}}} disabled={p===1} /></div>
      </div>
      <div style={{ marginBottom:16 }}><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:6,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Localização</div><div style={{ display:"flex",gap:8 }}>{[{v:false,l:"SP capital / grande SP"},{v:true,l:"Fora de SP"}].map(op=>(<button key={String(op.v)} onClick={()=>{setFora(op.v);setRes(null);}} style={{ flex:1,padding:"9px 12px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",background:fora===op.v?C.tealL:"var(--color-background-secondary)",color:fora===op.v?C.tealD:"var(--color-text-secondary)",border:"0.5px solid "+(fora===op.v?C.teal:"var(--color-border-tertiary)") }}>{op.l}</button>))}</div></div>
      <button onClick={res ? salvar : calcular} disabled={(!pronto&&!res)||salvo} style={{ width:"100%",padding:"11px",borderRadius:10,fontSize:13,fontWeight:500,cursor:(pronto||res)&&!salvo?"pointer":"default",background:salvo?C.green:res||pronto?C.teal:"var(--color-background-secondary)",color:salvo||res||pronto?"#fff":"var(--color-text-tertiary)",border:"none",marginBottom:20 }}>
        {salvo?"✓ Salvo no CRM!":res?"Adicionar ao CRM →":"Calcular perfil →"}
      </button>
      {res&&(
        <div>
          <div style={{ display:"grid",gridTemplateColumns:p===1?"1fr 1fr":"1fr 1fr 1fr",gap:8,marginBottom:12 }}>
            {p===1?(<><M label="Desde o pedido" value={res.diasUnico+"d"} cor={res.diasUnico>30?C.coralD:undefined} sub={res.diasUnico<=30?"✓ Quente":res.diasUnico<=60?"⚠ Morno":"⛔ Frio"}/><M label="Situação" value="1 pedido" sub="Objetivo: 2ª compra"/></>):(<><M label="Ciclo médio" value={res.ciclo+"d"} sub="entre pedidos"/><M label="Desde o último" value={res.diasUlt+"d"} cor={res.foraDaJanela?C.coralD:undefined} sub={res.foraDaJanela?"⚠ Fora da janela":"✓ Janela ok"}/><M label="Cliente há" value={res.span+"d"}/></>)}
          </div>
          <div style={{ background:res.cor+"18",border:"1px solid "+res.cor,borderRadius:10,padding:"10px 14px",marginBottom:res.alerta?8:12 }}><div style={{ fontSize:10,color:res.corD,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:2 }}>Objetivo</div><div style={{ fontSize:14,fontWeight:500,color:res.corD }}>{res.label}</div></div>
          {res.alerta&&<div style={{ border:"0.5px solid "+res.cor,borderLeft:"3px solid "+res.cor,borderRadius:8,padding:"8px 12px",marginBottom:12,background:res.cor+"12" }}><div style={{ fontSize:12,color:res.corD,lineHeight:1.5 }}>{res.alerta}</div></div>}
          <ProbBar prob={res.prob}/>
          <div style={{ fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",marginBottom:10,textTransform:"uppercase",letterSpacing:"0.06em" }}>Sequência — {res.seq.length} etapas</div>
          <Steps steps={res.seq} cur={0}/>

        </div>
      )}
    </div>
  );
};

const Perfil = ({ clienteId, onVoltar }) => {
  const [c,setC]=useState(null); const [confirmDel,setConfirmDel]=useState(false); const [salvando,setSalvando]=useState(false); const [toast,setToast]=useState("");
  useEffect(() => { dbGetAll().then(lista => { const cl = lista.find(c=>c.id===clienteId); if(cl) setC(cl); }); }, [clienteId]);
  const save = async (updates) => { const novo={...c,...updates}; setC(novo); try { await dbSave(novo); } catch(e) {} };
  const mover = async (etapaId) => { setSalvando(true); await save({etapa:etapaId}); setSalvando(false); };
  const avancar = () => save({stepAtual:Math.min(c.stepAtual+1,c.seq.length-1)});
  const deletar = async () => { if(!confirmDel){setConfirmDel(true);setTimeout(()=>setConfirmDel(false),3000);return;} try { await dbDelete(clienteId); } catch(e) {} onVoltar(); };
  const calcularCiclo = async (dp2, du2, fora2) => {
    const r = runTriagem(c.p, dp2, c.p===1?dp2:du2, fora2, c.gasto||0);
    const atualizado = Object.assign({},c,{dataPrimeiro:dp2,dataUltimo:du2,datasPreenchidas:true,objetivo:r.obj,objetivoLabel:r.label,objetivoCor:r.cor,objetivoCorD:r.corD,objetivoAlerta:r.alerta,prob:r.prob.pct,probLabel:r.prob.label,probCor:r.prob.cor,seq:r.seq,stepAtual:0,cicloMedio:r.ciclo});
    setC(atualizado); try { await dbSave(atualizado); } catch(e) {}
  };
  if (!c) return <div style={{ textAlign:"center",padding:40,color:"var(--color-text-tertiary)" }}>Carregando...</div>;
  const etapa = ETAPAS.find(e=>e.id===c.etapa)||ETAPAS[0];
  const hoje = new Date().toISOString().split("T")[0];
  const amanha = new Date(Date.now()+86400000).toISOString().split("T")[0];
  const vencido = c.dataProximoContato&&c.dataProximoContato<hoje;
  const urgente = c.dataProximoContato&&c.dataProximoContato<=amanha&&!vencido;
  const semSeq = !c.datasPreenchidas&&c.seq&&c.seq.length===0;
  return (
    <div>
      <div style={{ display:"flex",alignItems:"center",gap:12,marginBottom:16 }}>
        <button onClick={onVoltar} style={{ background:"none",border:"none",color:C.teal,fontSize:13,fontWeight:500,cursor:"pointer",padding:0 }}>← Kanban</button>
        <div style={{ flex:1 }}><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:2 }}>{c.customerId?("ID "+c.customerId+" · "):""}{c.nome}</div><div style={{ fontSize:14,fontWeight:500,color:"var(--color-text-primary)",lineHeight:1.3 }}>{c.proximaAcao||"— sem próxima ação"}</div></div>
        <button onClick={deletar} style={{ background:confirmDel?C.coralL:"none",border:"0.5px solid "+(confirmDel?C.coral:"var(--color-border-tertiary)"),borderRadius:6,padding:"4px 10px",fontSize:11,color:confirmDel?C.coralD:"var(--color-text-tertiary)",cursor:"pointer" }}>{confirmDel?"Confirmar?":"Remover"}</button>
      </div>
      {semSeq&&(
        <div style={{ background:C.amberL,border:"1px solid "+C.amber,borderRadius:12,padding:"14px 16px",marginBottom:12 }}>
          <div style={{ fontSize:13,fontWeight:500,color:C.amberD,marginBottom:8 }}>⚠ Preencha as datas para gerar a sequência</div>
          <div style={{ marginBottom:10 }}><div style={{ fontSize:11,color:C.amberD,marginBottom:4,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Localização</div><div style={{ display:"flex",gap:8 }}>{[{v:false,l:"SP / grande SP"},{v:true,l:"Fora de SP"}].map(op=>(<button key={String(op.v)} onClick={()=>setC({...c,fora:op.v})} style={{ flex:1,padding:"7px 10px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",background:c.fora===op.v?C.amber:"var(--color-background-secondary)",color:c.fora===op.v?"#fff":"var(--color-text-secondary)",border:"0.5px solid "+(c.fora===op.v?C.amber:"var(--color-border-tertiary)") }}>{op.l}</button>))}</div></div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10 }}>
            <div><div style={{ fontSize:11,color:C.amberD,marginBottom:4,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Data do 1° pedido</div><input type="date" value={c.dataPrimeiro||""} onChange={e=>setC({...c,dataPrimeiro:e.target.value})} style={{ width:"100%",padding:"8px 10px",borderRadius:8,border:"0.5px solid "+C.amber,fontSize:13,color:C.amberD,background:"#fff",outline:"none" }}/></div>
            <div><div style={{ fontSize:11,color:C.amberD,marginBottom:4,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Data do último {c.p===1&&<span style={{ fontWeight:400 }}>(= 1°)</span>}</div><input type="date" value={c.p===1?(c.dataPrimeiro||""):(c.dataUltimo||"")} onChange={e=>{if(c.p>1)setC({...c,dataUltimo:e.target.value});}} disabled={c.p===1} style={{ width:"100%",padding:"8px 10px",borderRadius:8,border:"0.5px solid "+C.amber,fontSize:13,color:C.amberD,background:c.p===1?C.amberL:"#fff",outline:"none",opacity:c.p===1?0.6:1 }}/></div>
          </div>
          <button onClick={()=>calcularCiclo(c.dataPrimeiro,c.dataUltimo,c.fora||false)} disabled={!c.dataPrimeiro||(c.p>1&&!c.dataUltimo)} style={{ width:"100%",padding:"10px",borderRadius:10,fontSize:13,fontWeight:500,cursor:c.dataPrimeiro&&(c.p===1||c.dataUltimo)?"pointer":"default",background:c.dataPrimeiro&&(c.p===1||c.dataUltimo)?C.amber:"var(--color-background-secondary)",color:c.dataPrimeiro&&(c.p===1||c.dataUltimo)?"#fff":"var(--color-text-tertiary)",border:"none" }}>Calcular ciclo e gerar sequência →</button>
        </div>
      )}
      <div style={{ background:etapa.corL,border:"1px solid "+etapa.cor,borderRadius:12,padding:"12px 16px",marginBottom:12 }}>
        <div style={{ fontSize:11,color:etapa.corD,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:4 }}>Etapa atual</div>
        <div style={{ fontSize:14,fontWeight:500,color:etapa.corD,marginBottom:etapa.id==="lead"?6:10 }}>{etapa.emoji} {etapa.label}</div>
        {etapa.id==="lead"&&<div style={{ fontSize:12,color:etapa.corD,background:C.purpleL,borderRadius:6,padding:"5px 8px",marginBottom:10,lineHeight:1.4 }}>Mova para <strong>Primeiro Contato</strong> após enviar a primeira mensagem.</div>}
        <div style={{ fontSize:11,color:etapa.corD,marginBottom:6 }}>Mover para:</div>
        <div style={{ display:"flex",gap:6,flexWrap:"wrap" }}>{ETAPAS.filter(e=>e.id!==c.etapa).map(e=>(<button key={e.id} onClick={async()=>{ await mover(e.id); }} style={{ padding:"4px 10px",borderRadius:20,fontSize:11,fontWeight:500,cursor:salvando?"default":"pointer",background:e.corL,color:e.corD,border:"0.5px solid "+e.cor,opacity:salvando?0.6:1 }}>{salvando?"...":e.emoji+" "+e.label}</button>))}</div>
      </div>
      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12 }}>
        <M label="Objetivo" value={c.objetivoLabel} cor={c.objetivoCorD}/>
        <M label="Probabilidade" value={c.prob+"%"} sub={c.probLabel} cor={c.probCor}/>
        <M label="Pedidos · Ciclo" value={c.p+"p · "+(c.cicloMedio||"?")+"d"}/>
      </div>
      <div style={{ background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"16px",marginBottom:12 }}>
        <div style={{ fontSize:13,fontWeight:500,color:"var(--color-text-primary)",marginBottom:12 }}>Perfil do cliente</div>
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10 }}>
          <div><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Nome</div><input style={inp()} value={c.nome} onChange={e=>setC({...c,nome:e.target.value})} onBlur={()=>save({nome:c.nome})} /></div>
          <div><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Telefone / WhatsApp</div><input style={inp()} value={c.telefone||""} onChange={e=>setC({...c,telefone:e.target.value})} onBlur={()=>save({telefone:c.telefone})} placeholder="11 9XXXX-XXXX"/></div>
        </div>
        <div style={{ marginBottom:10 }}>
          <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Lista de origem (opcional)</div>
          <div style={{ display:"flex",gap:6,flexWrap:"wrap",marginBottom:6 }}>{LISTAS.map(l=>(<button key={l} onClick={()=>save({lista:l===c.lista?"":l})} style={{ padding:"4px 10px",borderRadius:20,fontSize:11,cursor:"pointer",background:c.lista===l?C.purpleL:"var(--color-background-secondary)",color:c.lista===l?C.purpleD:"var(--color-text-secondary)",border:"0.5px solid "+(c.lista===l?C.purple:"var(--color-border-tertiary)") }}>{l}</button>))}</div>
          <input style={inp({fontSize:12})} value={LISTAS.includes(c.lista||"")?"":c.lista||""} onChange={e=>setC({...c,lista:e.target.value})} onBlur={()=>save({lista:c.lista})} placeholder="Ou digite o nome da lista manualmente..."/>
        </div>
        <div style={{ marginBottom:10 }}><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Anotações</div><textarea value={c.notas} onChange={e=>setC({...c,notas:e.target.value})} onBlur={()=>save({notas:c.notas})} placeholder="Sabor favorito, objeções, contexto..." rows={3} style={{ width:"100%",padding:"8px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none",resize:"vertical",fontFamily:"inherit",lineHeight:1.5 }}/></div>
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
          <div><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Próxima ação</div><input style={inp()} value={c.proximaAcao||""} onChange={e=>setC({...c,proximaAcao:e.target.value})} onBlur={()=>save({proximaAcao:c.proximaAcao})} placeholder="Ex: Ligar após T1"/></div>
          <div><div style={{ fontSize:11,color:vencido?C.coralD:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Data próximo contato {vencido?"⚠ Vencida":""}</div><input type="date" value={c.dataProximoContato||""} onChange={e=>save({dataProximoContato:e.target.value})} style={{ width:"100%",padding:"8px 10px",borderRadius:8,border:"0.5px solid "+(vencido?C.coral:urgente?C.amber:"var(--color-border-tertiary)"),fontSize:13,color:vencido?C.coralD:urgente?C.amberD:"var(--color-text-primary)",background:vencido?C.coralL:urgente?C.amberL:"var(--color-background-secondary)",outline:"none" }}/></div>
        </div>
      </div>
      {c.seq&&c.seq.length>0&&(
        <div style={{ background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"16px",marginBottom:12 }}>
          <div style={{ display:"flex",alignItems:"center",gap:8,marginBottom:12 }}>
            <div style={{ fontSize:13,fontWeight:500,color:"var(--color-text-primary)",flex:1 }}>Sequência</div>
            <div style={{ display:"flex",gap:3 }}>{c.seq.map((_,i)=>(<div key={i} style={{ width:8,height:8,borderRadius:"50%",background:i<=c.stepAtual?C.teal:"var(--color-border-tertiary)" }}/>))}</div>
            <span style={{ fontSize:11,color:"var(--color-text-tertiary)" }}>{c.stepAtual+1}/{c.seq.length}</span>
          </div>
          <Steps steps={c.seq} cur={c.stepAtual}/>
          {c.stepAtual<c.seq.length-1&&<button onClick={avancar} style={{ width:"100%",marginTop:10,padding:"10px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",background:C.teal,color:"#fff",border:"none" }}>Cliente respondeu → avançar para passo {c.stepAtual+2} ↓</button>}
        </div>
      )}
      {toast&&<div style={{ position:"fixed",top:24,left:"50%",transform:"translateX(-50%)",background:C.green,color:"#fff",padding:"10px 24px",borderRadius:30,fontSize:14,fontWeight:500,zIndex:999,boxShadow:"0 4px 20px rgba(0,0,0,0.15)" }}>{toast}</div>}
      <div style={{ padding:"14px 16px",background:"var(--color-background-secondary)",borderRadius:12,border:"0.5px solid var(--color-border-tertiary)" }}>
        <div style={{ fontSize:13,fontWeight:500,color:"var(--color-text-primary)",marginBottom:10 }}>Encerrar atendimento</div>
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8 }}>{[
          {label:"✓ Club",resultado:"club",etapa:"convertido",cor:C.green},
          {label:"✓ Avulso",resultado:"avulso",etapa:"convertido",cor:C.amber},
          {label:"✗ Nao converteu",resultado:"nao_converteu",etapa:"encerrado",cor:C.coral}
        ].map(op=>{
          const handleEncerrar = async () => {
            setSalvando(true);
            await save({etapa:op.etapa});
            await dbSaveConversao(c, op.resultado);
            setSalvando(false);
            setToast("✓ " + op.label + " registrado!");
            setTimeout(() => { setToast(""); onVoltar(); }, 1500);
          };
          return (<button key={op.label} onClick={handleEncerrar} disabled={salvando} style={{ padding:"8px 6px",borderRadius:8,fontSize:11,fontWeight:500,cursor:salvando?"default":"pointer",background:op.cor+"22",color:op.cor,border:"0.5px solid "+op.cor,lineHeight:1.3,opacity:salvando?0.6:1 }}>{salvando?"...":op.label}</button>);
        })}</div>
      </div>
    </div>
  );
};


const METAS = {
  "2026-05":20,"2026-06":36,"2026-07":36,"2026-08":36,
  "2026-09":71,"2026-10":36,"2026-11":71,"2026-12":20,
};

const mesKey = (d) => d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0");
const mesLabel = (d) => d.toLocaleDateString("pt-BR",{month:"long",year:"numeric"});

const Dashboard = ({ clientes, conversoes }) => {
  const hoje = new Date();
  const mesAtual = mesKey(hoje);
  const meta = METAS[mesAtual] || 0;

  // Stage counts
  const stageCounts = ETAPAS.map(e=>({...e, count:clientes.filter(c=>c.etapa===e.id).length}));

  // Conversoes do mes atual
  const convMes = conversoes.filter(c=>{
    try { return mesKey(new Date(c.registrado_em))===mesAtual; } catch(e){ return false; }
  });
  const clubMes = convMes.filter(c=>c.resultado==="club").length;
  const totalEnc = conversoes.filter(c=>{
    try { return mesKey(new Date(c.registrado_em))===mesAtual; } catch(e){ return false; }
  }).length;
  const faltam = Math.max(0, meta - clubMes);
  const pct = meta > 0 ? Math.min(100, Math.round(clubMes/meta*100)) : 0;

  // Taxa de conversao: club / total leads no CRM
  // Logica: se taxa alta → precisa de menos leads. Se taxa baixa → precisa de mais.
  const totalLeads = clientes.length;
  const totalClubHist = conversoes.filter(c=>c.resultado==="club").length;
  const taxa = totalLeads > 0 ? totalClubHist/totalLeads : 0;
  const taxaPct = Math.round(taxa*100);

  // Leads ativos (nao encerrados, nao convertidos)
  const leadsAtivos = clientes.filter(c=>c.etapa!=="encerrado"&&c.etapa!=="convertido").length;
  const leadsNecessarios = taxa > 0 ? Math.ceil(faltam/taxa) : "—";

  return (
    <div style={{marginBottom:20}}>
      {/* Stage cards */}
      <div style={{display:"flex",gap:8,overflowX:"auto",paddingBottom:4,marginBottom:14}}>
        {stageCounts.map(e=>(
          <div key={e.id} style={{flexShrink:0,background:e.count>0?e.corL:"var(--color-background-secondary)",border:"0.5px solid "+(e.count>0?e.cor:"var(--color-border-tertiary)"),borderRadius:10,padding:"8px 12px",minWidth:90,textAlign:"center"}}>
            <div style={{fontSize:10,color:e.count>0?e.corD:"var(--color-text-tertiary)",fontWeight:500,marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em"}}>{e.emoji} {e.label}</div>
            <div style={{fontSize:22,fontWeight:500,color:e.count>0?e.corD:"var(--color-text-tertiary)"}}>{e.count}</div>
          </div>
        ))}
      </div>

      {/* Meta do mes */}
      <div style={{background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px",marginBottom:10}}>
        <div style={{display:"flex",alignItems:"flex-start",gap:16,marginBottom:10}}>
          <div style={{flex:1}}>
            <div style={{fontSize:11,color:"var(--color-text-tertiary)",fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:2}}>Meta Club — {mesLabel(hoje)}</div>
            <div style={{display:"flex",alignItems:"baseline",gap:6}}>
              <span style={{fontSize:28,fontWeight:500,color:clubMes>=meta&&meta>0?C.green:C.teal}}>{clubMes}</span>
              <span style={{fontSize:16,color:"var(--color-text-tertiary)"}}>/ {meta}</span>
              {faltam>0&&<span style={{fontSize:12,color:C.coralD,background:C.coralL,padding:"2px 8px",borderRadius:20}}>faltam {faltam}</span>}
              {clubMes>=meta&&meta>0&&<span style={{fontSize:12,color:C.greenD,background:C.greenL,padding:"2px 8px",borderRadius:20}}>Meta batida!</span>}
            </div>
          </div>
          <div style={{textAlign:"right",flexShrink:0}}>
            <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:2}}>Progresso</div>
            <div style={{fontSize:20,fontWeight:500,color:pct>=100?C.green:pct>=70?C.amber:C.coral}}>{pct}%</div>
          </div>
        </div>
        <div style={{height:6,background:"var(--color-border-tertiary)",borderRadius:3,overflow:"hidden"}}>
          <div style={{width:pct+"%",height:"100%",background:pct>=100?C.green:pct>=70?C.amber:C.coral,borderRadius:3,transition:"width 0.5s ease"}}/>
        </div>
      </div>

      {/* Perspectiva de leads */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
        <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 12px",borderLeft:"3px solid "+C.purple}}>
          <div style={{fontSize:10,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Taxa conversao Club</div>
          <div style={{fontSize:20,fontWeight:500,color:C.purpleD}}>{taxaPct}%</div>
          <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>{totalClubHist} club / {totalLeads} leads no CRM</div>
        </div>
        <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 12px",borderLeft:"3px solid "+C.amber}}>
          <div style={{fontSize:10,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Leads necessarios p/ meta</div>
          <div style={{fontSize:20,fontWeight:500,color:C.amberD}}>{faltam===0?"Meta batida":leadsNecessarios}</div>
          <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>{faltam===0?"":"para converter mais "+faltam+" assinaturas"}</div>
        </div>
        <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 12px",borderLeft:"3px solid "+(leadsAtivos>=(typeof leadsNecessarios==="number"?leadsNecessarios:0)?C.green:C.coral)}}>
          <div style={{fontSize:10,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Leads ativos no CRM</div>
          <div style={{fontSize:20,fontWeight:500,color:leadsAtivos>=(typeof leadsNecessarios==="number"?leadsNecessarios:0)?C.greenD:C.coralD}}>{leadsAtivos}</div>
          <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>{typeof leadsNecessarios==="number"&&leadsAtivos>=leadsNecessarios?"suficiente p/ meta":typeof leadsNecessarios==="number"&&faltam>0?"faltam "+(leadsNecessarios-leadsAtivos)+" leads":"adicione mais leads"}</div>
        </div>
      </div>
    </div>
  );
};

const Historico = () => {
  const [conv, setConv] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    dbGetConversoes().then(data=>{ setConv(data); setLoading(false); });
  }, []);

  if (loading) return <div style={{textAlign:"center",padding:40,color:"var(--color-text-tertiary)"}}>Carregando historico...</div>;
  if (conv.length===0) return (
    <div style={{textAlign:"center",padding:"48px 24px",background:"var(--color-background-secondary)",borderRadius:12,border:"0.5px dashed var(--color-border-tertiary)"}}>
      <div style={{fontSize:32,marginBottom:12}}>📊</div>
      <div style={{fontSize:14,fontWeight:500,color:"var(--color-text-primary)",marginBottom:6}}>Sem historico ainda</div>
      <div style={{fontSize:13,color:"var(--color-text-secondary)"}}>Encerre atendimentos no perfil do cliente para ver o historico aqui.</div>
    </div>
  );

  // Group by month
  const meses = {};
  conv.forEach(c=>{
    try {
      const d = new Date(c.registrado_em);
      const key = mesKey(d);
      const label = mesLabel(d);
      const meta = METAS[key]||0;
      if(!meses[key]) meses[key]={key,label,meta,club:0,avulso:0,nao:0,total:0};
      if(c.resultado==="club") meses[key].club++;
      else if(c.resultado==="avulso") meses[key].avulso++;
      else meses[key].nao++;
      meses[key].total++;
    } catch(e){}
  });

  const lista = Object.values(meses).sort((a,b)=>b.key.localeCompare(a.key));
  const totalClub = conv.filter(c=>c.resultado==="club").length;
  const totalAvulso = conv.filter(c=>c.resultado==="avulso").length;
  const totalNao = conv.filter(c=>c.resultado==="nao_converteu").length;

  return (
    <div>
      {/* Totais gerais */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:8,marginBottom:16}}>
        {[
          {label:"Total",value:conv.length,cor:C.teal},
          {label:"Club",value:totalClub,cor:C.green},
          {label:"Avulso",value:totalAvulso,cor:C.amber},
          {label:"Nao converteu",value:totalNao,cor:C.coral},
        ].map((s,i)=>(
          <div key={i} style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 14px",borderLeft:"3px solid "+s.cor}}>
            <div style={{fontSize:10,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>{s.label}</div>
            <div style={{fontSize:22,fontWeight:500,color:s.cor}}>{s.value}</div>
          </div>
        ))}
      </div>

      {/* Tabela por mes */}
      <div style={{background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px"}}>
        <div style={{fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:12}}>Historico por mes</div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:"var(--color-background-primary)"}}>
                {["Mes","Meta","Club","Avulso","Nao conv.","Total","vs Meta"].map(h=>(
                  <th key={h} style={{padding:"7px 10px",textAlign:h==="Mes"?"left":"center",fontWeight:500,color:"var(--color-text-tertiary)",fontSize:11,borderBottom:"0.5px solid var(--color-border-tertiary)",textTransform:"uppercase",letterSpacing:"0.05em"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {lista.map(m=>{
                const diff = m.club - m.meta;
                const corDiff = diff>=0?C.green:C.coral;
                return (
                  <tr key={m.key} style={{borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
                    <td style={{padding:"8px 10px",fontWeight:500,color:"var(--color-text-primary)",textTransform:"capitalize"}}>{m.label}</td>
                    <td style={{padding:"8px 10px",textAlign:"center",color:"var(--color-text-secondary)"}}>{m.meta||"—"}</td>
                    <td style={{padding:"8px 10px",textAlign:"center",color:m.club>0?C.greenD:"var(--color-text-tertiary)",fontWeight:m.club>0?500:400}}>{m.club}</td>
                    <td style={{padding:"8px 10px",textAlign:"center",color:m.avulso>0?C.amberD:"var(--color-text-tertiary)",fontWeight:m.avulso>0?500:400}}>{m.avulso}</td>
                    <td style={{padding:"8px 10px",textAlign:"center",color:m.nao>0?C.coralD:"var(--color-text-tertiary)",fontWeight:m.nao>0?500:400}}>{m.nao}</td>
                    <td style={{padding:"8px 10px",textAlign:"center",fontWeight:500}}>{m.total}</td>
                    <td style={{padding:"8px 10px",textAlign:"center"}}>
                      {m.meta>0?(
                        <span style={{background:corDiff+"22",color:corDiff,padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:500}}>
                          {diff>=0?"+":""}{diff}
                        </span>
                      ):"—"}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};


const Kanban = ({ onAbrir }) => {
  const [clientes,setClientes]=useState([]); const [loading,setLoading]=useState(true); const [conversoes,setConversoes]=useState([]);
  const [busca,setBusca]=useState("");
  const [abertos,setAbertos]=useState({});
  const toggleGrupo=(etapaId,grupo)=>{ const k=etapaId+"_"+grupo; setAbertos(a=>({...a,[k]:!a[k]})); };
  const isAberto=(etapaId,grupo)=>!!abertos[etapaId+"_"+grupo];

  const carregar = useCallback(async () => {
    setLoading(true);
    try {
      const [todos, conv] = await Promise.all([dbGetAll(), dbGetConversoes()]);
      setClientes(todos); setConversoes(conv);
    } catch(e) { setClientes([]); setConversoes([]); }
    setLoading(false);
  }, []);
  useEffect(()=>{carregar();},[carregar]);

  const hoje=new Date().toISOString().split("T")[0];
  const amanha=new Date(Date.now()+86400000).toISOString().split("T")[0];

  const filtrar=(lista)=>{
    if(!busca.trim()) return lista;
    const q=busca.toLowerCase();
    return lista.filter(c=>(c.nome||"").toLowerCase().includes(q)||(c.customerId||"").toLowerCase().includes(q)||(c.telefone||"").toLowerCase().includes(q));
  };

  const porEtapa=(id)=>{
    const grupo=filtrar(clientes.filter(c=>c.etapa===id));
    const vencidos=grupo.filter(c=>c.dataProximoContato&&c.dataProximoContato<hoje).sort((a,b)=>a.dataProximoContato>b.dataProximoContato?1:-1);
    const deHoje=grupo.filter(c=>c.dataProximoContato===hoje).sort((a,b)=>b.prob-a.prob);
    const deAmanha=grupo.filter(c=>c.dataProximoContato===amanha).sort((a,b)=>b.prob-a.prob);
    const depois=grupo.filter(c=>c.dataProximoContato&&c.dataProximoContato>amanha).sort((a,b)=>a.dataProximoContato>b.dataProximoContato?1:-1);
    const semData=grupo.filter(c=>!c.dataProximoContato).sort((a,b)=>b.prob-a.prob);
    return {vencidos,deHoje,deAmanha,depois,semData,total:grupo.length};
  };

  const Card=({cl})=>{
    const v=cl.dataProximoContato&&cl.dataProximoContato<hoje;
    const u=cl.dataProximoContato===hoje;
    const am=cl.dataProximoContato===amanha;
    return (
      <button onClick={()=>onAbrir(cl.id)} style={{ width:"100%",textAlign:"left",background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)",borderLeft:"3px solid "+cl.probCor,borderRadius:8,padding:"10px",marginBottom:6,cursor:"pointer" }}>
        <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:2,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis" }}>{cl.customerId?"#"+cl.customerId+" · ":""}{cl.nome}</div>
        <div style={{ fontSize:13,fontWeight:500,color:"var(--color-text-primary)",marginBottom:4,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis" }}>{cl.proximaAcao||"—"}</div>
        <div style={{ display:"flex",alignItems:"center",gap:4,marginBottom:cl.dataProximoContato?4:0 }}>
          <span style={{ fontSize:11,background:cl.probCor+"22",color:cl.probCor,padding:"1px 6px",borderRadius:20,fontWeight:500 }}>{cl.prob}%</span>
          <span style={{ fontSize:10,color:"var(--color-text-tertiary)" }}>{cl.p}p · {cl.cicloMedio||"?"}d</span>
        </div>
        {cl.dataProximoContato&&<div style={{ fontSize:10,color:v?C.coralD:u||am?C.amberD:"var(--color-text-tertiary)",background:v?C.coralL:u||am?C.amberL:"transparent",padding:v||u||am?"1px 5px":0,borderRadius:4 }}>{v?"⚠ Vencida":u?"⚡ Hoje":am?"📅 Amanhã":"📅"} {!u&&!am&&new Date(cl.dataProximoContato+"T12:00:00").toLocaleDateString("pt-BR")}</div>}
        {cl.lista&&<div style={{ fontSize:10,color:C.purpleD,marginTop:3 }}>{cl.lista}</div>}
      </button>
    );
  };

  const GrupoFixo=({label,cor,items})=>{
    if(items.length===0) return null;
    return (
      <div style={{ marginBottom:6 }}>
        <div style={{ fontSize:9,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.08em",color:cor,padding:"3px 4px",marginBottom:4 }}>{label} · {items.length}</div>
        {items.map(cl=><Card key={cl.id} cl={cl}/>)}
      </div>
    );
  };

  const MAX_CARDS = 50;

  const GrupoSanfona=({etapaId,grupo,label,cor,items})=>{
    if(items.length===0) return null;
    const aberto=isAberto(etapaId,grupo);
    return (
      <div style={{ marginBottom:4 }}>
        <button onClick={()=>toggleGrupo(etapaId,grupo)} style={{ width:"100%",display:"flex",alignItems:"center",gap:6,padding:"4px 4px",background:"none",border:"none",cursor:"pointer",borderRadius:6 }}>
          <span style={{ fontSize:9,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.08em",color:cor,flex:1,textAlign:"left" }}>{label} · {items.length}</span>
          <span style={{ fontSize:10,color:"var(--color-text-tertiary)" }}>{aberto?"▲":"▼"}</span>
        </button>
        {aberto&&<div style={{ marginTop:4 }}>{items.map(cl=><Card key={cl.id} cl={cl}/>)}</div>}
      </div>
    );
  };

  if (loading) return <div style={{ textAlign:"center",padding:40,color:"var(--color-text-tertiary)" }}>Carregando...</div>;
  if (clientes.length===0) return (
    <div style={{ textAlign:"center",padding:"48px 24px",background:"var(--color-background-secondary)",borderRadius:12,border:"0.5px dashed var(--color-border-tertiary)" }}>
      <div style={{ fontSize:36,marginBottom:12 }}>📋</div>
      <div style={{ fontSize:15,fontWeight:500,color:"var(--color-text-primary)",marginBottom:6 }}>CRM vazio</div>
      <div style={{ fontSize:13,color:"var(--color-text-secondary)",lineHeight:1.6 }}>Use a aba <strong>📥 Importar</strong> ou <strong>🎯 Triagem</strong> para adicionar clientes.</div>
    </div>
  );

  return (
    <div>
      <Dashboard clientes={clientes} conversoes={conversoes}/>
      <div style={{ display:"flex",alignItems:"center",gap:10,marginBottom:14 }}>
        <div style={{ position:"relative",flex:1 }}>
          <span style={{ position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",fontSize:14,pointerEvents:"none" }}>🔍</span>
          <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="Buscar por nome, ID ou telefone..." style={{ width:"100%",padding:"8px 12px 8px 32px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}/>
        </div>
        <div style={{ fontSize:13,color:"var(--color-text-tertiary)",whiteSpace:"nowrap" }}>{clientes.length} clientes</div>
        <button onClick={carregar} style={{ padding:"5px 12px",borderRadius:8,fontSize:12,background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-secondary)",cursor:"pointer",whiteSpace:"nowrap" }}>↺</button>
      </div>
      <div style={{ overflowX:"auto",paddingBottom:8 }}>
        <div style={{ display:"flex",gap:10,minWidth:"max-content" }}>
          {ETAPAS.map(etapa=>{
            const {vencidos,deHoje,deAmanha,depois,semData,total}=porEtapa(etapa.id);
            return (
              <div key={etapa.id} style={{ width:220,flexShrink:0 }}>
                <div style={{ display:"flex",alignItems:"center",gap:6,padding:"8px 10px",background:etapa.corL,borderRadius:"10px 10px 0 0",border:"0.5px solid "+etapa.cor,borderBottom:"none" }}>
                  <span style={{ fontSize:14 }}>{etapa.emoji}</span>
                  <span style={{ fontSize:12,fontWeight:500,color:etapa.corD,flex:1 }}>{etapa.label}</span>
                  <span style={{ fontSize:11,background:etapa.cor,color:"#fff",padding:"1px 7px",borderRadius:20 }}>{total}</span>
                </div>
                <div style={{ border:"0.5px solid "+etapa.cor,borderTop:"none",borderRadius:"0 0 10px 10px",padding:8,minHeight:80,background:"var(--color-background-primary)" }}>
                  <GrupoFixo label="⚠ Vencido" cor={C.coralD} items={vencidos}/>
                  <GrupoFixo label="⚡ Hoje" cor={C.amberD} items={deHoje}/>
                  {vencidos.length===0&&deHoje.length===0&&deAmanha.length===0&&depois.length===0&&semData.length===0&&(
                    <div style={{ fontSize:11,color:"var(--color-text-tertiary)",textAlign:"center",padding:"16px 0" }}>Vazio</div>
                  )}
                  <GrupoSanfona etapaId={etapa.id} grupo="amanha" label="📅 Amanhã" cor={C.blueD} items={deAmanha}/>
                  <GrupoSanfona etapaId={etapa.id} grupo="depois" label="🗓 Depois" cor="var(--color-text-tertiary)" items={depois}/>
                  <GrupoSanfona etapaId={etapa.id} grupo="semdata" label="— Sem data" cor="var(--color-text-tertiary)" items={semData}/>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};


const cepToFora = (cep) => {
  const c = (cep||"").replace(/\D/g,"");
  if (!c) return null;
  return !c.startsWith("0");
};

const excelDateToISO = (val) => {
  const n = parseInt(val);
  if (isNaN(n) || n < 40000 || n > 60000) return "";
  // Excel epoch: Jan 1 1900 = 1, with leap year bug (day 60 = Feb 29 1900, doesn't exist)
  const ms = (n - 25569) * 86400000;
  const d = new Date(ms);
  if (isNaN(d.getTime())) return "";
  return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
};

const parseShopifyDate = (str) => {
  if (!str && str !== 0) return "";
  const s = String(str).trim();
  if (!s) return "";
  // Excel serial number
  if (/^\d{4,6}$/.test(s)) return excelDateToISO(s);
  // ISO format
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0,10);
  // BR format DD/MM/AAAA
  if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) {
    const [d,m,y] = s.split("/");
    return y+"-"+m.padStart(2,"0")+"-"+d.padStart(2,"0");
  }
  return "";
};

const ImportarLista = ({ onSalvo }) => {
  const [txt,setTxt]=useState(""); const [prev,setPrev]=useState([]); const [imp,setImp]=useState(false);
  const [ok,setOk]=useState(null); const [erro,setErro]=useState("");
  const parse = (raw) => {
    const linhas = raw.split("\n").map(l=>l.trim()).filter(l=>l.length>0);
    const cls=[]; const errs=[];
    // Skip header row if first line looks like a header
    const start = /customer|nome|id/i.test(linhas[0]||"") ? 1 : 0;
    linhas.slice(start).forEach((linha,i)=>{
      const sep = linha.includes(";") ? ";" : ",";
      const pts = linha.split(sep).map(p=>p.trim());
      if(pts.length<4){errs.push("Linha "+(i+1+start)+": minimo 4 colunas");return;}
      const [customerId, nome, tel, gastoStr, pedStr, dp6, du7, cep8, lista9] = pts;
      const lista = (lista9||"").trim();
      const ped = parseInt(pedStr);
      const gasto = parseFloat((gastoStr||"0").replace(",","."))||0;
      const dp = parseShopifyDate(dp6||"");
      const du = parseShopifyDate(du7||"");
      const fora = cepToFora(cep8||"");
      if(!nome){errs.push("Linha "+(i+1+start)+": nome vazio");return;}
      if(isNaN(ped)||ped<1){errs.push("Linha "+(i+1+start)+": pedidos invalido");return;}
      cls.push({customerId:customerId||"",nome,tel:tel||"",ped,gasto,dp,du,fora,cep:cep8||"",lista});
    });
    return {cls,errs};
  };
  const atualizar = (val) => {
    setTxt(val); setOk(null); setErro("");
    if(!val.trim()){setPrev([]);return;}
    const {cls,errs}=parse(val); setPrev(cls);
    if(errs.length>0) setErro(errs.join("\n"));
  };
  const [prog, setProg] = useState(null); // {atual, total, inicio}

  const tempoRestante = (prog) => {
    if (!prog || prog.atual === 0) return null;
    const elapsed = Date.now() - prog.inicio;
    const rate = elapsed / prog.atual;
    const restMs = rate * (prog.total - prog.atual);
    if (restMs < 1500) return "menos de 1 segundo";
    if (restMs < 60000) return Math.ceil(restMs / 1000) + " segundos";
    return Math.ceil(restMs / 60000) + " minuto" + (restMs > 120000 ? "s" : "");
  };

  const importar = () => {
    if(prev.length===0) return;
    setImp(true);
    const inicio = Date.now();

    const run = async () => {
      try {
        setProg({ atual: 0, total: prev.length, inicio });

        // Checar duplicatas — 1 leitura no Supabase
        const existentes = await dbGetAll();
        const customerIdsExist = new Set(existentes.map(c=>c.customerId).filter(Boolean).map(String));

        const novos = prev
          .filter(cl => !cl.customerId || !customerIdsExist.has(String(cl.customerId).trim()))
          .map(cl => {
            const temDados = !!(cl.dp && cl.fora !== null && cl.ped >= 1);
            const tr = temDados ? runTriagem(cl.ped, cl.dp, cl.ped===1?cl.dp:(cl.du||cl.dp), cl.fora, cl.gasto||0) : null;
            return {
              id:"c_"+Date.now()+"_"+Math.random().toString(36).slice(2,8),
              etapa:"lead", dataCriacao:new Date().toLocaleDateString("pt-BR"),
              notas:"", proximaAcao:"", dataProximoContato:"",
              lista:cl.lista, customerId:cl.customerId, nome:cl.nome,
              telefone:cl.tel, p:cl.ped, gasto:cl.gasto,
              fora:cl.fora, cep:cl.cep||"",
              dataPrimeiro:cl.dp||"", dataUltimo:cl.du||cl.dp||"",
              datasPreenchidas:temDados,
              objetivo:tr?tr.obj:"",
              objetivoLabel:tr?tr.label:"⚠ Preencher datas para calcular",
              objetivoCor:tr?tr.cor:C.purple,
              objetivoCorD:tr?tr.corD:C.purpleD,
              objetivoAlerta:tr?tr.alerta:"",
              prob:tr?tr.prob.pct:0,
              probLabel:tr?tr.prob.label:"Pendente",
              probCor:tr?tr.prob.cor:C.purple,
              seq:tr?tr.seq:[], stepAtual:0,
              cicloMedio:tr?tr.ciclo:0,
            };
          });

        if (novos.length === 0) {
          setErro("Todos os clientes já existem no CRM (mesmo Customer ID).");
          setImp(false); setProg(null);
          return;
        }

        const pulados = prev.length - novos.length;
        setProg({ atual: Math.floor(novos.length * 0.5), total: novos.length, inicio });

        // Bulk insert em lotes de 200 para evitar payload gigante
        const LOTE_SIZE = 200;
        let salvos = 0;
        for (let i = 0; i < novos.length; i += LOTE_SIZE) {
          const lote = novos.slice(i, i + LOTE_SIZE);
          await dbBulkSave(lote);
          salvos += lote.length;
          setProg({ atual: Math.floor(novos.length * 0.1) + Math.floor(salvos/novos.length * 0.85 * novos.length), total: novos.length, inicio });
        }
        setProg({ atual: novos.length, total: novos.length, inicio });

        const comTriagem = novos.filter(c=>c.datasPreenchidas).length;
        const semTriagem = novos.length - comTriagem;
        const msg = (pulados > 0 ? pulados + " ignorados (ID duplic.) · " : "")
          + novos.length + " importados"
          + (comTriagem > 0 ? " · " + comTriagem + " com triagem completa" : "")
          + (semTriagem > 0 ? " · " + semTriagem + " aguardando datas" : "");

        setImp(false); setOk("✓ " + msg); setProg(null);
        setTimeout(()=>{ setTxt(""); setPrev([]); setOk(null); setErro(""); onSalvo&&onSalvo(); }, 2500);

      } catch(e) {
        setErro("Erro: " + (e.message||"tente novamente"));
        setImp(false); setProg(null);
      }
    };
    run();
  };
  const lerArquivo = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "xls" || ext === "xlsx") {
      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const wb = XLSX.read(ev.target.result, { type:"array" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
          // Skip header row if first cell looks like a label (not a number/ID)
          const startRow = isNaN(rows[0]&&rows[0][0]) && typeof rows[0][0] === "string" && rows[0][0].toLowerCase().includes("id") ? 1 : 0;
          const txt = rows.slice(startRow).map(r => r.join(",")).join("\n");
          atualizar(txt);
        } catch(err) {
          setErro("Erro ao ler o arquivo Excel. Verifique o formato.");
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = (ev) => { atualizar(ev.target.result); };
      reader.readAsText(file, "UTF-8");
    }
    e.target.value = "";
  };
  return (
    <div>
      <div style={{ background:C.purpleL,border:"0.5px solid "+C.purple,borderRadius:8,padding:"12px 16px",marginBottom:16 }}>
        <div style={{ fontSize:13,fontWeight:500,color:C.purpleD,marginBottom:6 }}>Importação em lote de leads</div>
        <div style={{ fontSize:12,color:C.purpleD,lineHeight:1.6 }}>Cole a lista abaixo ou faça upload de um CSV. Todos entram na coluna Lead. O operador preenche as datas no perfil para gerar a sequência.</div>
      </div>
      <div style={{ background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 16px",marginBottom:16 }}>
        <div style={{ fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:8 }}>Formato aceito</div>
        <div style={{ fontSize:12,color:"var(--color-text-secondary)",lineHeight:1.8,fontFamily:"monospace" }}>
          Customer ID, Nome, Telefone, Total Gasto, Nº Pedidos, Data 1° Pedido, Data Último Pedido, CEP, Lista<br/>
          <span style={{ color:C.tealD }}>1234, Maria Silva, 11 99999-1111, 380, 4, 2025-11-18, 2026-03-10, 04547-130, Pascoa Falta Uma</span><br/>
          <span style={{ color:"var(--color-text-tertiary)" }}>// Formato Shopify — triagem calculada automaticamente pelo CEP e datas</span>
        </div>
        <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginTop:8 }}>
          CEPs iniciados em 0 = SP/Grande SP · Demais = Fora de SP · Datas: AAAA-MM-DD ou DD/MM/AAAA · Colunas 6-9 opcionais
        </div>
      </div>
      <label style={{ display:"flex",alignItems:"center",gap:10,padding:"12px 16px",marginBottom:12,background:"var(--color-background-secondary)",borderRadius:10,border:"1.5px dashed "+C.purple,cursor:"pointer" }}>
        <span style={{ fontSize:24 }}>📄</span>
        <div>
          <div style={{ fontSize:13,fontWeight:500,color:C.purpleD }}>Upload de CSV</div>
          <div style={{ fontSize:11,color:C.purpleD,opacity:0.7 }}>Clique para selecionar .csv, .xls ou .xlsx</div>
        </div>
        <input type="file" accept=".csv,.txt,.xls,.xlsx" onChange={lerArquivo} style={{ display:"none" }}/>
      </label>
      <div style={{ marginBottom:12 }}>
        <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Ou cole a lista aqui</div>
        <textarea value={txt} onChange={e=>atualizar(e.target.value)} placeholder={"1234, Maria Silva, 11 99999-1111, 380, 4, Páscoa Falta Uma\n5678, André Santos, 11 98888-2222, 290, 3"} rows={6} style={{ width:"100%",padding:"10px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none",resize:"vertical",fontFamily:"monospace",lineHeight:1.6 }}/>
      </div>
      {erro&&<div style={{ background:C.coralL,border:"0.5px solid "+C.coral,borderRadius:8,padding:"10px 14px",marginBottom:12 }}><div style={{ fontSize:11,fontWeight:500,color:C.coralD,marginBottom:4 }}>⚠ Linhas com erro serão ignoradas</div><div style={{ fontSize:11,color:C.coralD,whiteSpace:"pre-line",lineHeight:1.6 }}>{erro}</div></div>}
      {prev.length>0&&(
        <div style={{ marginBottom:16 }}>
          <div style={{ fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:8 }}>Preview — {prev.length} cliente{prev.length!==1?"s":""}</div>
          <div style={{ border:"0.5px solid var(--color-border-tertiary)",borderRadius:10,overflow:"hidden" }}>
            <div style={{ display:"grid",gridTemplateColumns:"70px 1.2fr 90px 40px 70px 80px 80px 80px 1fr",gap:6,padding:"7px 10px",background:"var(--color-background-secondary)",fontSize:10,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.05em" }}><span>ID</span><span>Nome</span><span>Telefone</span><span>Ped.</span><span>Gasto</span><span>1° Pedido</span><span>Ult. Pedido</span><span>Local</span><span>Lista</span></div>
            {prev.map((cl,i)=>{
              const localLabel = cl.fora===null?"?":(cl.fora?"Fora SP":"SP");
              const localCor = cl.fora===null?"var(--color-text-tertiary)":cl.fora?C.amberD:C.tealD;
              return (<div key={i} style={{ display:"grid",gridTemplateColumns:"70px 1.2fr 90px 40px 70px 80px 80px 80px 1fr",gap:6,padding:"7px 10px",borderTop:"0.5px solid var(--color-border-tertiary)",fontSize:11,color:"var(--color-text-primary)",alignItems:"center" }}>
                <span style={{ color:"var(--color-text-tertiary)" }}>{cl.customerId||"—"}</span>
                <span style={{ fontWeight:500 }}>{cl.nome}</span>
                <span style={{ color:"var(--color-text-secondary)" }}>{cl.tel||"—"}</span>
                <span style={{ color:"var(--color-text-secondary)" }}>{cl.ped}</span>
                <span style={{ color:"var(--color-text-secondary)" }}>R${cl.gasto||"—"}</span>
                <span style={{ color:"var(--color-text-tertiary)",fontSize:10 }}>{cl.dp||"—"}</span>
                <span style={{ color:"var(--color-text-tertiary)",fontSize:10 }}>{cl.du||"—"}</span>
                <span style={{ fontSize:10,fontWeight:500,color:localCor }}>{localLabel}</span>
                <span style={{ fontSize:10,color:C.purpleD }}>{cl.lista||"—"}</span>
              </div>);
            })}
          </div>
        </div>
      )}
      {prog && (
        <div style={{ marginBottom:12, background:"var(--color-background-secondary)", borderRadius:10, padding:"14px 16px", border:"0.5px solid var(--color-border-tertiary)" }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:8 }}>
            <div style={{ fontSize:13, fontWeight:500, color:"var(--color-text-primary)" }}>
              Importando {prog.atual} de {prog.total}...
            </div>
            <div style={{ fontSize:12, color:"var(--color-text-tertiary)" }}>
              {Math.round(prog.atual / prog.total * 100)}%
            </div>
          </div>
          <div style={{ height:8, background:"var(--color-border-tertiary)", borderRadius:4, overflow:"hidden", marginBottom:6 }}>
            <div style={{ height:"100%", width:(prog.atual/prog.total*100)+"%", background:C.purple, borderRadius:4, transition:"width 0.15s ease" }}/>
          </div>
          <div style={{ fontSize:11, color:"var(--color-text-tertiary)" }}>
            {prog.atual === prog.total ? "Finalizando..." : tempoRestante(prog) ? "Tempo restante: " + tempoRestante(prog) : "Calculando..."}
          </div>
        </div>
      )}
      <button onClick={importar} disabled={prev.length===0||imp||ok!==null} style={{ width:"100%",padding:"12px",borderRadius:10,fontSize:14,fontWeight:500,cursor:prev.length>0&&!imp?"pointer":"default",border:"none",background:ok!==null?C.green:prev.length>0&&!imp?C.purple:"var(--color-background-secondary)",color:ok!==null||(prev.length>0&&!imp)?"#fff":"var(--color-text-tertiary)" }}>
        {ok!==null?ok:imp?"Aguarde...":prev.length>0?"Importar "+prev.length+" lead"+(prev.length!==1?"s":"")+" →":"Cole a lista acima para importar"}
      </button>
    </div>
  );
};

const PascoaTab = ({ tipo, onSalvo }) => (
  <div>
    <div style={{ background:tipo==="falta_uma"?C.amberL:C.tealL,border:"0.5px solid "+(tipo==="falta_uma"?C.amber:C.teal),borderRadius:8,padding:"12px 16px",marginBottom:16 }}>
      <div style={{ fontSize:13,fontWeight:500,color:tipo==="falta_uma"?C.amberD:C.tealD,marginBottom:3 }}>{tipo==="falta_uma"?"🐣 Páscoa — Falta Uma (42 clientes)":"🐣 Páscoa — Reativação (646 clientes)"}</div>
      <div style={{ fontSize:12,color:tipo==="falta_uma"?C.amberD:C.tealD,lineHeight:1.55 }}>{tipo==="falta_uma"?"Fizeram 2 pedidos na Páscoa. Objetivo: 3ª compra.":"Fizeram 1 pedido na Páscoa. Objetivo: 2ª compra. Último pedido há ~46–62 dias."}</div>
    </div>
    <TriagemForm lista={tipo==="falta_uma"?"Páscoa — Falta Uma":"Páscoa — Reativação"} onSalvo={onSalvo}/>
  </div>
);

const Backup = ({ onRestore }) => {
  const [restaurando, setRestaurando] = useState(false);
  const [status, setStatus] = useState("");
  const [gsUrl, setGsUrl] = useState("");
  const [gsSincStatus, setGsSincStatus] = useState("");
  const [sincronizando, setSincronizando] = useState(false);

  // ── OPÇÃO 1: Export / Import JSON ──────────────────────────────────────────
  const exportar = async () => {
    const clientes = await dbGetAll();
    const dados = { versao: 1, exportadoEm: new Date().toISOString(), clientes };
    const blob = new Blob([JSON.stringify(dados, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "laricas_crm_backup_" + new Date().toLocaleDateString("pt-BR").replace(/\//g,"-") + ".json";
    a.click(); URL.revokeObjectURL(url);
    setStatus("✓ Backup exportado com sucesso!");
    setTimeout(() => setStatus(""), 3000);
  };

  const restaurar = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setRestaurando(true);
    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const dados = JSON.parse(ev.target.result);
        if (!dados.clientes || !Array.isArray(dados.clientes)) throw new Error("Formato inválido");
        await dbBulkSave(dados.clientes);
        setStatus("✓ " + dados.clientes.length + " clientes restaurados!");
        setRestaurando(false);
        setTimeout(() => { setStatus(""); onRestore && onRestore(); }, 2000);
      } catch(err) {
        setStatus("⚠ Erro ao restaurar. Verifique o arquivo.");
        setRestaurando(false);
      }
    };
    reader.readAsText(file, "UTF-8");
    e.target.value = "";
  };

  // ── OPÇÃO 2: Google Sheets ─────────────────────────────────────────────────
  const jsonp = (url) => new Promise((resolve, reject) => {
    const cb = "crm_cb_" + Date.now();
    window[cb] = (data) => { resolve(data); delete window[cb]; if (script.parentNode) script.parentNode.removeChild(script); };
    const script = document.createElement("script");
    script.src = url + (url.includes("?")?"&":"?") + "callback=" + cb;
    script.onerror = () => reject(new Error("script error"));
    document.head.appendChild(script);
    setTimeout(() => reject(new Error("timeout")), 12000);
  });

  const sincGSheets = async (sentido) => {
    if (!gsUrl.trim()) { setGsSincStatus("⚠ Cole a URL do Apps Script primeiro."); return; }
    setSincronizando(true);
    setGsSincStatus("Conectando...");
    try {
      if (sentido === "enviar") {
        const clientes = await dbGetAll();
        await fetch(gsUrl.trim(), {
          method: "POST",
          mode: "no-cors",
          body: JSON.stringify({ acao: "salvar", clientes: clientes.filter(Boolean) }),
          headers: { "Content-Type": "text/plain" },
        });
        setGsSincStatus("✓ " + clientes.filter(Boolean).length + " clientes enviados! Verifique na planilha.");
      } else {
        const r = await jsonp(gsUrl.trim());
        if (!r.clientes) throw new Error("Resposta inválida");
        await dbBulkSave(r.clientes);
        setGsSincStatus("✓ " + r.clientes.length + " clientes carregados do Sheets!");
        setTimeout(() => { onRestore && onRestore(); }, 2000);
      }
    } catch(e) {
      setGsSincStatus("⚠ Erro de conexão. Verifique a URL e tente novamente.");
    }
    setSincronizando(false);
    setTimeout(() => setGsSincStatus(""), 5000);
  };

  return (
    <div>
      {/* OPÇÃO 1 */}
      <div style={{ background:"var(--color-background-secondary)", borderRadius:12, padding:"16px 20px", marginBottom:16, border:"0.5px solid var(--color-border-tertiary)" }}>
        <div style={{ fontSize:14, fontWeight:500, color:"var(--color-text-primary)", marginBottom:4 }}>💾 Opção 1 — Backup local (JSON)</div>
        <div style={{ fontSize:12, color:"var(--color-text-secondary)", marginBottom:14, lineHeight:1.5 }}>Exporte todos os dados do CRM para um arquivo no seu computador. Restaure quando precisar. <strong style={{ fontWeight:500 }}>Recomendado exportar ao fim de cada dia de trabalho.</strong></div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
          <button onClick={exportar} style={{ padding:"10px", borderRadius:10, fontSize:13, fontWeight:500, cursor:"pointer", background:C.teal, color:"#fff", border:"none" }}>
            ⬇ Exportar backup JSON
          </button>
          <label style={{ padding:"10px", borderRadius:10, fontSize:13, fontWeight:500, cursor:"pointer", background:"var(--color-background-primary)", color:"var(--color-text-primary)", border:"0.5px solid var(--color-border-tertiary)", textAlign:"center", display:"flex", alignItems:"center", justifyContent:"center" }}>
            ⬆ Restaurar backup
            <input type="file" accept=".json" onChange={restaurar} style={{ display:"none" }}/>
          </label>
        </div>
        {status && <div style={{ marginTop:10, fontSize:12, color:status.startsWith("✓")?C.greenD:C.coralD, background:status.startsWith("✓")?C.greenL:C.coralL, padding:"6px 10px", borderRadius:6 }}>{status}</div>}
        {restaurando && <div style={{ marginTop:8, fontSize:12, color:C.amberD }}>Restaurando dados...</div>}
      </div>

      {/* OPÇÃO 2 */}
      <div style={{ background:"var(--color-background-secondary)", borderRadius:12, padding:"16px 20px", border:"0.5px solid var(--color-border-tertiary)" }}>
        <div style={{ fontSize:14, fontWeight:500, color:"var(--color-text-primary)", marginBottom:4 }}>📊 Opção 2 — Google Sheets (sincronização)</div>
        <div style={{ fontSize:12, color:"var(--color-text-secondary)", marginBottom:14, lineHeight:1.5 }}>Sincroniza os dados do CRM com uma planilha Google. Dados persistem para sempre, acessíveis de qualquer dispositivo.</div>

        <div style={{ background:C.amberL, border:"0.5px solid "+C.amber, borderRadius:8, padding:"12px 14px", marginBottom:14 }}>
          <div style={{ fontSize:12, fontWeight:500, color:C.amberD, marginBottom:8 }}>Configure em 3 passos</div>
          {[
            { n:"1", t:"Criar planilha", d:'Abra sheets.google.com e crie uma planilha nova. Anote o ID da URL (é a parte entre /d/ e /edit).' },
            { n:"2", t:"Instalar o script", d:'Na planilha: Extensões → Apps Script. Apague tudo e cole o script abaixo. Salve com Ctrl+S.' },
            { n:"3", t:"Publicar como Web App", d:'Clique em Implantar → Nova implantação → Web App. Acesso: "Qualquer pessoa". Copie a URL e cole no campo abaixo.' },
          ].map((s,i) => (
            <div key={i} style={{ display:"flex", gap:10, marginBottom:i<2?10:0 }}>
              <div style={{ width:20, height:20, borderRadius:"50%", background:C.amber, color:"#fff", fontSize:11, fontWeight:500, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>{s.n}</div>
              <div>
                <div style={{ fontSize:12, fontWeight:500, color:C.amberD }}>{s.t}</div>
                <div style={{ fontSize:11, color:C.amberD, lineHeight:1.5, opacity:0.85 }}>{s.d}</div>
              </div>
            </div>
          ))}
        </div>

        <div style={{ background:"#1e1e1e", borderRadius:8, padding:"12px 14px", marginBottom:14, fontFamily:"monospace", fontSize:11, color:"#d4d4d4", lineHeight:1.7, overflowX:"auto" }}>
          <div style={{ color:"#608b4e" }}>// Cole este código no Apps Script</div>
          <div><span style={{ color:"#569cd6" }}>const</span> SHEET_NAME = <span style={{ color:"#ce9178" }}>"CRM"</span>;</div>
          <div><span style={{ color:"#569cd6" }}>function</span> <span style={{ color:"#dcdcaa" }}>doGet</span>(e) {"{"}</div>
          <div>&nbsp;&nbsp;<span style={{ color:"#569cd6" }}>const</span> sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);</div>
          <div>&nbsp;&nbsp;<span style={{ color:"#569cd6" }}>const</span> rows = sheet.getDataRange().getValues();</div>
          <div>&nbsp;&nbsp;<span style={{ color:"#569cd6" }}>const</span> clientes = rows.slice(<span style={{ color:"#b5cea8" }}>1</span>).map(r ={">"} JSON.parse(r[<span style={{ color:"#b5cea8" }}>1</span>] || <span style={{ color:"#ce9178" }}>'null'</span>)).filter(Boolean);</div>
          <div>&nbsp;&nbsp;<span style={{ color:"#c586c0" }}>return</span> ContentService.createTextOutput(JSON.stringify({"{"}<span style={{ color:"#9cdcfe" }}>clientes</span>{"}"})).setMimeType(ContentService.MimeType.JSON);</div>
          <div>{"}"}</div>
          <div><span style={{ color:"#569cd6" }}>function</span> <span style={{ color:"#dcdcaa" }}>doPost</span>(e) {"{"}</div>
          <div>&nbsp;&nbsp;<span style={{ color:"#569cd6" }}>const</span> data = JSON.parse(e.postData.contents);</div>
          <div>&nbsp;&nbsp;<span style={{ color:"#569cd6" }}>const</span> sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);</div>
          <div>&nbsp;&nbsp;sheet.clearContents();</div>
          <div>&nbsp;&nbsp;sheet.appendRow([<span style={{ color:"#ce9178" }}>"id"</span>, <span style={{ color:"#ce9178" }}>"dados"</span>, <span style={{ color:"#ce9178" }}>"nome"</span>, <span style={{ color:"#ce9178" }}>"etapa"</span>, <span style={{ color:"#ce9178" }}>"lista"</span>, <span style={{ color:"#ce9178" }}>"dataCriacao"</span>]);</div>
          <div>&nbsp;&nbsp;data.clientes.forEach(c ={">"} sheet.appendRow([c.id, JSON.stringify(c), c.nome, c.etapa, c.lista||<span style={{ color:"#ce9178" }}>""</span>, c.dataCriacao||<span style={{ color:"#ce9178" }}>""</span>]));</div>
          <div>&nbsp;&nbsp;<span style={{ color:"#c586c0" }}>return</span> ContentService.createTextOutput(JSON.stringify({"{"}<span style={{ color:"#9cdcfe" }}>salvos</span>: data.clientes.length{"}"})).setMimeType(ContentService.MimeType.JSON);</div>
          <div>{"}"}</div>
        </div>

        <div style={{ marginBottom:10 }}>
          <div style={{ fontSize:11, color:"var(--color-text-tertiary)", marginBottom:5, fontWeight:500, textTransform:"uppercase", letterSpacing:"0.06em" }}>URL do Web App (Apps Script)</div>
          <input value={gsUrl} onChange={e=>setGsUrl(e.target.value)} placeholder="https://script.google.com/macros/s/..." style={{ width:"100%", padding:"9px 12px", borderRadius:8, border:"0.5px solid var(--color-border-tertiary)", fontSize:13, color:"var(--color-text-primary)", background:"var(--color-background-primary)", outline:"none" }}/>
        </div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
          <button onClick={()=>sincGSheets("enviar")} disabled={sincronizando} style={{ padding:"10px", borderRadius:10, fontSize:13, fontWeight:500, cursor:sincronizando?"default":"pointer", background:C.green, color:"#fff", border:"none", opacity:sincronizando?0.6:1 }}>
            ⬆ Enviar para Sheets
          </button>
          <button onClick={()=>sincGSheets("carregar")} disabled={sincronizando} style={{ padding:"10px", borderRadius:10, fontSize:13, fontWeight:500, cursor:sincronizando?"default":"pointer", background:C.blue, color:"#fff", border:"none", opacity:sincronizando?0.6:1 }}>
            ⬇ Carregar do Sheets
          </button>
        </div>
        {gsSincStatus && <div style={{ marginTop:10, fontSize:12, color:gsSincStatus.startsWith("✓")?C.greenD:C.coralD, background:gsSincStatus.startsWith("✓")?C.greenL:C.coralL, padding:"6px 10px", borderRadius:6 }}>{gsSincStatus}</div>}
      </div>
    </div>
  );
};


// ── CONFIG SUPABASE ───────────────────────────────────────────────────────────
const ConfigSupabase = ({ onSalvo }) => {
  const [url, setUrl] = useState(_SB.url||"");
  const [key, setKey] = useState(_SB.key||"");
  const [status, setStatus] = useState("");
  const [testando, setTestando] = useState(false);

  const testar = async () => {
    if (!url.trim()||!key.trim()) { setStatus("⚠ Preencha os dois campos."); return; }
    setTestando(true); setStatus("Testando conexao...");
    Object.assign(_SB, {url:url.trim(), key:key.trim()});
    let timedOut = false;
    const timer = setTimeout(() => { timedOut = true; setTestando(false); setStatus("⚠ Timeout — Supabase demorou demais. Verifique o RLS e tente novamente."); Object.assign(_SB,{url:"",key:""}); }, 12000);
    try {
      await dbTest();
      clearTimeout(timer);
      if (timedOut) return;
      await saveCfg({url:url.trim(), key:key.trim()});
      setStatus("✓ Conectado! Redirecionando...");
      setTimeout(()=>{ onSalvo&&onSalvo(); }, 1500);
    } catch(e) {
      clearTimeout(timer);
      if (!timedOut) { Object.assign(_SB,{url:"",key:""}); setStatus("⚠ Erro: " + e.message); setTestando(false); }
    }
  };

  const pular = () => { onSalvo&&onSalvo(); };

  return (
    <div style={{ maxWidth:560, margin:"0 auto" }}>
      <div style={{ textAlign:"center", padding:"32px 0 24px" }}>
        <div style={{ fontSize:40, marginBottom:12 }}>🗄</div>
        <div style={{ fontSize:20, fontWeight:500, color:"var(--color-text-primary)", marginBottom:6 }}>Conectar ao Supabase</div>
        <div style={{ fontSize:13, color:"var(--color-text-secondary)", lineHeight:1.6 }}>
          Cole as credenciais do seu projeto Supabase para ativar o CRM.
        </div>
      </div>

      <div style={{ background:"var(--color-background-secondary)", borderRadius:12, padding:"20px", marginBottom:16, border:"0.5px solid var(--color-border-tertiary)" }}>
        <div style={{ fontSize:13, fontWeight:500, color:"var(--color-text-primary)", marginBottom:12 }}>Passo a passo</div>
        {[
          ["1", "Acesse supabase.com → New project → nome laricas-crm → regiao South America"],
          ["2", "SQL Editor → New query → cole e execute:\n\ncreate table clientes (\n  id text primary key,\n  dados jsonb not null,\n  atualizado_em timestamptz default now()\n);"],
          ["3", "Settings → API → copie o Project URL e o anon public key abaixo"],
        ].map(([n,t])=>(
          <div key={n} style={{ display:"flex", gap:10, marginBottom:10 }}>
            <div style={{ width:22,height:22,borderRadius:"50%",background:C.teal,color:"#fff",fontSize:11,fontWeight:500,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0 }}>{n}</div>
            <div style={{ fontSize:12, color:"var(--color-text-secondary)", lineHeight:1.6, whiteSpace:"pre-line" }}>{t}</div>
          </div>
        ))}
      </div>

      <div style={{ marginBottom:12 }}>
        <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Project URL</div>
        <input value={url} onChange={e=>setUrl(e.target.value)} placeholder="https://xxxxxxxxxxxx.supabase.co" style={{ width:"100%",padding:"10px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}/>
      </div>
      <div style={{ marginBottom:20 }}>
        <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Anon Public Key</div>
        <input value={key} onChange={e=>setKey(e.target.value)} placeholder="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..." style={{ width:"100%",padding:"10px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}/>
      </div>

      <button onClick={testar} disabled={testando} style={{ width:"100%",padding:"12px",borderRadius:10,fontSize:14,fontWeight:500,cursor:testando?"default":"pointer",background:C.teal,color:"#fff",border:"none",opacity:testando?0.7:1,marginBottom:10 }}>
        {testando?"Testando... (aguarde ate 12s)":"Conectar e salvar →"}
      </button>
      <button onClick={pular} style={{ width:"100%",padding:"10px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",background:"none",color:"var(--color-text-tertiary)",border:"0.5px solid var(--color-border-tertiary)" }}>
        Pular por enquanto (dados nao serao salvos)
      </button>
      {status&&<div style={{ marginTop:10,fontSize:12,color:status.startsWith("✓")?C.greenD:C.coralD,background:status.startsWith("✓")?C.greenL:C.coralL,padding:"8px 12px",borderRadius:8 }}>{status}</div>}
    </div>
  );
};

export default function App() {
  const [tab,setTab]=useState("kanban");
  const [clienteId,setClienteId]=useState(null);
  const [refresh,setRefresh]=useState(0);
  const [cfgOk,setCfgOk]=useState(false);
  const [cfgLoad,setCfgLoad]=useState(true);
  useEffect(()=>{ loadCfg().then(cfg=>{ setCfgOk(!!(cfg.url&&cfg.key)); setCfgLoad(false); }); },[]);
  const onSalvo = () => { setTab("kanban"); setRefresh(r=>r+1); };
  const onRestore = () => { setTab("kanban"); setRefresh(r=>r+1); };
  if (cfgLoad) return <div style={{ textAlign:"center",padding:"60px 0",color:"var(--color-text-tertiary)",fontFamily:"var(--font-sans)" }}>Carregando...</div>;
  if (!cfgOk) return (
    <div style={{ maxWidth:900,margin:"0 auto",padding:"0 20px 40px",fontFamily:"var(--font-sans)",color:"var(--color-text-primary)" }}>
      <ConfigSupabase onSalvo={()=>{ loadCfg().then(()=>setCfgOk(true)); }}/>
    </div>
  );
  return (
    <div style={{ maxWidth:900,margin:"0 auto",padding:"0 0 40px",fontFamily:"var(--font-sans)",color:"var(--color-text-primary)" }}>
      <div style={{ padding:"20px 0 8px" }}>
        <div style={{ fontSize:11,fontWeight:500,letterSpacing:"0.09em",textTransform:"uppercase",color:"var(--color-text-tertiary)",marginBottom:4 }}>Laricas Fitness</div>
        <div style={{ fontSize:22,fontWeight:500,lineHeight:1.3 }}>CRM de Conversão</div>
        <div style={{ fontSize:13,color:"var(--color-text-secondary)",marginTop:4 }}>Lead → Contato → Conversa → Proposta → Convertido</div>
      </div>
      <div style={{ display:"flex",borderBottom:"0.5px solid var(--color-border-tertiary)",marginBottom:24,overflowX:"auto" }}>
        <T label="📋 Kanban" active={tab==="kanban"} color={C.green} onClick={()=>{setClienteId(null);setTab("kanban");}}/>
        <T label="📥 Importar" active={tab==="import"} color={C.purple} onClick={()=>setTab("import")}/>
        <T label="🎯 Triagem" active={tab==="triagem"} color={C.teal} onClick={()=>setTab("triagem")}/>
        <T label="📊 Historico" active={tab==="historico"} color={C.teal} onClick={()=>setTab("historico")}/>
        <T label="💾 Backup" active={tab==="backup"} color={C.blue} onClick={()=>setTab("backup")}/>
        <T label="⚙ Config" active={tab==="config"} color="var(--color-text-tertiary)" onClick={()=>setTab("config")}/>
      </div>
      {tab==="kanban"&&(clienteId?<Perfil key={clienteId} clienteId={clienteId} onVoltar={()=>{setClienteId(null);setRefresh(r=>r+1);}}/>:<Kanban key={refresh} onAbrir={setClienteId}/>)}
      {tab==="import"&&<ImportarLista onSalvo={onSalvo}/>}
      {tab==="triagem"&&<TriagemForm onSalvo={onSalvo}/>}
      {tab==="historico"&&<Historico/>}
      {tab==="backup"&&<Backup onRestore={onRestore}/>}
      {tab==="config"&&<ConfigSupabase onSalvo={()=>{ loadCfg().then(()=>setCfgOk(true)); setTab("kanban"); }}/>}
    </div>
  );
}
