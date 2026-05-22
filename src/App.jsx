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
  {id:"experiencia",label:"Experiencia",emoji:"⭐",cor:C.purple,corL:C.purpleL,corD:C.purpleD},
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

const dbGetAll  = async () => {
  const PAGE = 1000;
  let all = [], offset = 0;
  while (true) {
    const r = await sb("/clientes?select=dados&order=atualizado_em.desc&limit="+PAGE+"&offset="+offset);
    const pg = (r||[]).map(x=>x.dados).filter(Boolean);
    all = [...all, ...pg];
    if (pg.length < PAGE) break;
    offset += PAGE;
  }
  return all;
};
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
const dbGetAssinantes = async () => {
  try {
    const r = await sb("/clientes?select=dados&dados->>etapa=eq.experiencia");
    return (r||[]).map(x=>x.dados).filter(Boolean);
  } catch(e) {
    // Fallback: get all and filter
    try { const all = await dbGetAll(); return all.filter(c=>c.etapa==="experiencia"); }
    catch(e2) { return []; }
  }
};
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


const CadastroRapido = ({ onSalvo }) => {
  const [modo, setModo] = useState(null); // "lead" | "club"
  const [nome, setNome] = useState("");
  const [tel, setTel] = useState("");
  const [lista, setLista] = useState("");
  const [listaCustom, setListaCustom] = useState("");
  const [fora, setFora] = useState(null);
  // Club fields
  const [tipoAssin, setTipoAssin] = useState("");
  const [dataInicio, setDataInicio] = useState("");
  const [valorMensal, setValorMensal] = useState("");
  const [salvo, setSalvo] = useState(false);

  const listaFinal = LISTAS.includes(lista) ? lista : listaCustom;

  const salvar = async () => {
    if (!nome.trim()) return;
    const base = {
      id: "c_"+Date.now()+"_"+Math.random().toString(36).slice(2,8),
      dataCriacao: new Date().toLocaleDateString("pt-BR"),
      notas: "", proximaAcao: "", dataProximoContato: "",
      nome: nome.trim(), telefone: tel.trim(),
      lista: listaFinal, customerId: "",
      p: 0, gasto: 0, fora: fora||false, cep: "",
      dataPrimeiro: "", dataUltimo: "",
      datasPreenchidas: false,
      objetivo: modo==="club" ? "club" : "reativacao",
      objetivoLabel: modo==="club" ? "Assinante direto" : "Primeiro contato — sem historico ecom",
      objetivoCor: modo==="club" ? C.purple : C.teal,
      objetivoCorD: modo==="club" ? C.purpleD : C.tealD,
      prob: modo==="club" ? 0 : 15,
      probLabel: modo==="club" ? "Assinante" : "Baixa",
      probCor: modo==="club" ? C.purple : C.coral,
      seq: [], stepAtual: 0, cicloMedio: 0,
      etapa: modo==="club" ? "experiencia" : "lead",
      historicoEtapas: [],
      ...(modo==="club" ? {
        tipoAssinatura: tipoAssin,
        dataInicioAssinatura: dataInicio,
        valorMensal: parseFloat(valorMensal)||0,
      } : {}),
    };
    try {
      await dbSave(base);
      setSalvo(true);
      setTimeout(() => {
        setModo(null); setNome(""); setTel(""); setLista(""); setListaCustom("");
        setFora(null); setTipoAssin(""); setDataInicio(""); setValorMensal(""); setSalvo(false);
        onSalvo && onSalvo();
      }, 1500);
    } catch(e) { alert("Erro ao salvar: " + e.message); }
  };

  const camposBase = (
    <div>
      <div style={{ marginBottom:12 }}>
        <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Nome <span style={{ color:C.coralD,fontSize:10 }}>*obrigatorio</span></div>
        <input value={nome} onChange={e=>setNome(e.target.value)} placeholder="Ex: Maria Silva"
          style={{ width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:14,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}/>
      </div>
      <div style={{ marginBottom:12 }}>
        <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Telefone / WhatsApp</div>
        <input value={tel} onChange={e=>setTel(e.target.value)} placeholder="11 9XXXX-XXXX"
          style={{ width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:14,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}/>
      </div>
      <div style={{ marginBottom:12 }}>
        <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Lista de origem (opcional)</div>
        <div style={{ display:"flex",gap:6,flexWrap:"wrap",marginBottom:6 }}>
          {LISTAS.map(l=>(
            <button key={l} onClick={()=>{setLista(l===lista?"":l);setListaCustom("");}}
              style={{ padding:"4px 10px",borderRadius:20,fontSize:11,cursor:"pointer",background:lista===l?C.purpleL:"var(--color-background-secondary)",color:lista===l?C.purpleD:"var(--color-text-secondary)",border:"0.5px solid "+(lista===l?C.purple:"var(--color-border-tertiary)") }}>
              {l}
            </button>
          ))}
        </div>
        <input value={LISTAS.includes(lista)?"":listaCustom} onChange={e=>{setListaCustom(e.target.value);setLista("");}}
          placeholder="Ou digite o nome da lista..." style={{ width:"100%",padding:"8px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}/>
      </div>
      <div style={{ marginBottom:16 }}>
        <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:6,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Localizacao (opcional)</div>
        <div style={{ display:"flex",gap:8 }}>
          {[{v:false,l:"SP / Grande SP"},{v:true,l:"Fora de SP"}].map(op=>(
            <button key={String(op.v)} onClick={()=>setFora(op.v)}
              style={{ flex:1,padding:"8px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",background:fora===op.v?C.tealL:"var(--color-background-secondary)",color:fora===op.v?C.tealD:"var(--color-text-secondary)",border:"0.5px solid "+(fora===op.v?C.teal:"var(--color-border-tertiary)") }}>
              {op.l}
            </button>
          ))}
        </div>
      </div>
    </div>
  );

  return (
    <div style={{ marginBottom:24 }}>
      <div style={{ background:C.amberL,border:"0.5px solid "+C.amber,borderRadius:8,padding:"12px 16px",marginBottom:16 }}>
        <div style={{ fontSize:13,fontWeight:500,color:C.amberD,marginBottom:3 }}>Cadastro rapido</div>
        <div style={{ fontSize:12,color:C.amberD,lineHeight:1.5 }}>Para clientes sem historico no ecom — indicacoes, assinantes diretos ou primeiros contatos.</div>
      </div>

      {!modo && (
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:12 }}>
          <button onClick={()=>setModo("lead")} style={{ padding:"16px",borderRadius:12,border:"1.5px solid "+C.teal,background:C.tealL,cursor:"pointer",textAlign:"left" }}>
            <div style={{ fontSize:20,marginBottom:6 }}>🎯</div>
            <div style={{ fontSize:13,fontWeight:500,color:C.tealD,marginBottom:4 }}>Novo Lead</div>
            <div style={{ fontSize:11,color:C.tealD,lineHeight:1.5 }}>Entrou em contato mas nunca comprou. Entra como Lead sem historico de compras.</div>
          </button>
          <button onClick={()=>setModo("club")} style={{ padding:"16px",borderRadius:12,border:"1.5px solid "+C.purple,background:C.purpleL,cursor:"pointer",textAlign:"left" }}>
            <div style={{ fontSize:20,marginBottom:6 }}>⭐</div>
            <div style={{ fontSize:13,fontWeight:500,color:C.purpleD,marginBottom:4 }}>Assinante Direto</div>
            <div style={{ fontSize:11,color:C.purpleD,lineHeight:1.5 }}>Assinou o Club sem passar pelo ecom. Entra diretamente em Experiencia.</div>
          </button>
        </div>
      )}

      {modo && (
        <div>
          <button onClick={()=>setModo(null)} style={{ background:"none",border:"none",color:C.teal,fontSize:12,fontWeight:500,cursor:"pointer",padding:0,marginBottom:14 }}>← Voltar</button>
          <div style={{ background:modo==="club"?C.purpleL:C.tealL,border:"0.5px solid "+(modo==="club"?C.purple:C.teal),borderRadius:8,padding:"10px 14px",marginBottom:14 }}>
            <div style={{ fontSize:12,fontWeight:500,color:modo==="club"?C.purpleD:C.tealD }}>{modo==="club"?"⭐ Assinante Direto — entra em Experiencia":"🎯 Novo Lead — entra como Lead"}</div>
          </div>
          {camposBase}
          {modo==="club" && (
            <div style={{ background:"var(--color-background-secondary)",borderRadius:10,padding:"14px",marginBottom:16,border:"0.5px solid var(--color-border-tertiary)" }}>
              <div style={{ fontSize:12,fontWeight:500,color:"var(--color-text-primary)",marginBottom:10 }}>Dados da assinatura</div>
              <div style={{ marginBottom:10 }}>
                <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:6,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Tipo de assinatura</div>
                <div style={{ display:"flex",gap:8 }}>
                  {TIPOS_ASSINATURA.map(t=>(
                    <button key={t.id} onClick={()=>setTipoAssin(t.id)}
                      style={{ flex:1,padding:"8px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",background:tipoAssin===t.id?C.purple:"var(--color-background-primary)",color:tipoAssin===t.id?"#fff":"var(--color-text-secondary)",border:"0.5px solid "+(tipoAssin===t.id?C.purple:"var(--color-border-tertiary)") }}>
                      {t.label}<br/><span style={{ fontSize:10,fontWeight:400 }}>{t.ciclosTotais} meses</span>
                    </button>
                  ))}
                </div>
              </div>
              <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
                <div>
                  <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Data de inicio</div>
                  <input type="date" value={dataInicio} onChange={e=>setDataInicio(e.target.value)}
                    style={{ width:"100%",padding:"8px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-primary)",outline:"none" }}/>
                </div>
                <div>
                  <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Valor medio mensal R$</div>
                  <input type="number" min="0" placeholder="Ex: 89.90" value={valorMensal} onChange={e=>setValorMensal(e.target.value)}
                    style={{ width:"100%",padding:"8px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-primary)",outline:"none" }}/>
                </div>
              </div>
            </div>
          )}
          <button onClick={salvar} disabled={!nome.trim()||salvo}
            style={{ width:"100%",padding:"11px",borderRadius:10,fontSize:13,fontWeight:500,cursor:nome.trim()&&!salvo?"pointer":"default",background:salvo?C.green:nome.trim()?(modo==="club"?C.purple:C.teal):"var(--color-background-secondary)",color:nome.trim()||salvo?"#fff":"var(--color-text-tertiary)",border:"none" }}>
            {salvo?"✓ Salvo!":modo==="club"?"Adicionar como Assinante Direto →":"Adicionar como Lead →"}
          </button>
        </div>
      )}
    </div>
  );
};


const NovoClienteForm = ({ onSalvo }) => {
  const [nome, setNome] = useState("");
  const [tel, setTel] = useState("");
  const [origem, setOrigem] = useState("");
  const [origemCustom, setOrigemCustom] = useState("");
  const [obs, setObs] = useState("");
  const [salvo, setSalvo] = useState(false);

  const ORIGENS = ["Indicacao", "Instagram", "TikTok", "Evento", "Presencial", "Outra"];

  const salvar = async () => {
    if (!nome.trim()) return;
    const origemFinal = origem === "Outra" ? origemCustom : origem;
    const seq = buildSeq("novo_cliente", 0, 0, false, false, 0);
    const prob = { pct: 25, label: "Media", cor: C.amber, corD: C.amberD, corL: C.amberL };
    const c = {
      id: "c_"+Date.now()+"_"+Math.random().toString(36).slice(2,8),
      etapa: "lead", dataCriacao: new Date().toLocaleDateString("pt-BR"),
      notas: obs.trim(), proximaAcao: "", dataProximoContato: "",
      lista: "Novo cliente — "+origemFinal,
      nome: nome.trim(), telefone: tel.trim(),
      customerId: "", p: 0, gasto: 0, fora: null,
      dataPrimeiro: "", dataUltimo: "", datasPreenchidas: false,
      objetivo: "novo_cliente", objetivoLabel: "Novo cliente — 1a compra",
      objetivoCor: C.teal, objetivoCorD: C.tealD, objetivoAlerta: "",
      prob: prob.pct, probLabel: prob.label, probCor: prob.cor,
      seq, stepAtual: 0, cicloMedio: 0,
    };
    await dbSave(c);
    setSalvo(true);
    setTimeout(() => {
      setNome(""); setTel(""); setOrigem(""); setOrigemCustom(""); setObs(""); setSalvo(false);
      onSalvo && onSalvo();
    }, 1500);
  };

  return (
    <div style={{ background:C.tealL, border:"0.5px solid "+C.teal, borderRadius:12, padding:"16px 20px" }}>
      <div style={{ fontSize:14, fontWeight:500, color:C.tealD, marginBottom:4 }}>Novo cliente sem historico de compra</div>
      <div style={{ fontSize:12, color:C.tealD, marginBottom:16, lineHeight:1.5 }}>
        Para clientes que chegaram por indicacao, redes sociais ou contato direto — sem compras no ecommerce ainda.
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:12 }}>
        <div>
          <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Nome <span style={{ color:C.coralD,fontSize:10 }}>*obrigatorio</span></div>
          <input style={{ width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:14,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}
            type="text" placeholder="Ex: Maria Silva" value={nome} onChange={e=>setNome(e.target.value)}/>
        </div>
        <div>
          <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Telefone / WhatsApp</div>
          <input style={{ width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:14,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}
            type="text" placeholder="11 9XXXX-XXXX" value={tel} onChange={e=>setTel(e.target.value)}/>
        </div>
      </div>
      <div style={{ marginBottom:12 }}>
        <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:6,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Como chegou ate a Laricas</div>
        <div style={{ display:"flex",gap:6,flexWrap:"wrap",marginBottom:origem==="Outra"?8:0 }}>
          {ORIGENS.map(o=>(
            <button key={o} onClick={()=>setOrigem(o===origem?"":o)}
              style={{ padding:"5px 12px",borderRadius:20,fontSize:12,fontWeight:500,cursor:"pointer",
                background:origem===o?C.teal:"var(--color-background-secondary)",
                color:origem===o?"#fff":"var(--color-text-secondary)",
                border:"0.5px solid "+(origem===o?C.teal:"var(--color-border-tertiary)") }}>
              {o}
            </button>
          ))}
        </div>
        {origem==="Outra"&&(
          <input style={{ width:"100%",padding:"8px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}
            type="text" placeholder="Como chegou?" value={origemCustom} onChange={e=>setOrigemCustom(e.target.value)}/>
        )}
      </div>
      <div style={{ marginBottom:16 }}>
        <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Observacoes <span style={{ fontWeight:400 }}>(opcional)</span></div>
        <textarea value={obs} onChange={e=>setObs(e.target.value)} placeholder="Contexto do contato, produto de interesse, etc..."
          rows={2} style={{ width:"100%",padding:"8px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none",resize:"none",fontFamily:"inherit" }}/>
      </div>
      {!nome.trim()&&<div style={{ fontSize:12,color:C.coralD,background:C.coralL,padding:"6px 10px",borderRadius:6,marginBottom:8 }}>Preencha o nome para habilitar.</div>}
      <button onClick={salvar} disabled={!nome.trim()||salvo}
        style={{ width:"100%",padding:"11px",borderRadius:10,fontSize:13,fontWeight:500,
          cursor:nome.trim()&&!salvo?"pointer":"default",
          background:salvo?C.green:nome.trim()?C.teal:"var(--color-background-secondary)",
          color:nome.trim()||salvo?"#fff":"var(--color-text-tertiary)",border:"none" }}>
        {salvo?"✓ Cliente adicionado ao CRM!":"Adicionar ao CRM como Lead →"}
      </button>
    </div>
  );
};

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
  const mover = async (etapaId) => {
    setSalvando(true);
    const hist = (c.historicoEtapas || []).slice(-9); // keep last 9, adding current = 10
    hist.push({ etapa: c.etapa, data: new Date().toLocaleDateString("pt-BR"), hora: new Date().toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"}) });
    await save({ etapa: etapaId, historicoEtapas: hist });
    setSalvando(false);
  };
  const desfazer = async () => {
    const hist = [...(c.historicoEtapas || [])];
    if (hist.length === 0) return;
    const anterior = hist.pop();
    setSalvando(true);
    await save({ etapa: anterior.etapa, historicoEtapas: hist });
    setSalvando(false);
  };
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
        {c.historicoEtapas&&c.historicoEtapas.length>0&&(
          <button onClick={desfazer} disabled={salvando} title={"Voltar para: "+ETAPAS.find(e=>e.id===(c.historicoEtapas[c.historicoEtapas.length-1]||{}).etapa)?.label}
            style={{ background:C.amberL,border:"0.5px solid "+C.amber,borderRadius:6,padding:"4px 10px",fontSize:11,color:C.amberD,cursor:salvando?"default":"pointer",fontWeight:500 }}>
            ↩ Desfazer
          </button>
        )}
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
      {c.seq&&c.seq.length>0&&c.etapa!=="experiencia"&&(
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
      {c.etapa==="experiencia"&&(()=>{
        const assin = calcAssinatura(c.tipoAssinatura, c.dataInicioAssinatura);
        const corCobranca = assin ? (assin.diasParaCobranca<=7?C.coral:assin.diasParaCobranca<=15?C.amber:C.green) : "var(--color-text-tertiary)";
        const corFim = assin ? (assin.diasParaFim<=30?C.coral:assin.diasParaFim<=60?C.amber:C.teal) : "var(--color-text-tertiary)";
        return (
          <div style={{ background:"var(--color-background-primary)",border:"0.5px solid "+C.purple,borderLeft:"3px solid "+C.purple,borderRadius:12,padding:"16px",marginBottom:12 }}>
            <div style={{ fontSize:13,fontWeight:500,color:C.purpleD,marginBottom:12 }}>⭐ Assinatura Club</div>
            <div style={{ marginBottom:10 }}>
              <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:6,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Tipo de assinatura</div>
              <div style={{ display:"flex",gap:8 }}>
                {TIPOS_ASSINATURA.map(t=>(
                  <button key={t.id} onClick={()=>save({tipoAssinatura:t.id})}
                    style={{ flex:1,padding:"8px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",background:c.tipoAssinatura===t.id?C.purple:"var(--color-background-secondary)",color:c.tipoAssinatura===t.id?"#fff":"var(--color-text-secondary)",border:"0.5px solid "+(c.tipoAssinatura===t.id?C.purple:"var(--color-border-tertiary)") }}>
                    {t.label}<br/><span style={{ fontSize:10,fontWeight:400 }}>{t.ciclosTotais} meses</span>
                  </button>
                ))}
              </div>
            </div>
            <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12 }}>
              <div>
                <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Data de inicio da assinatura</div>
                <input type="date" value={c.dataInicioAssinatura||""} onChange={e=>save({dataInicioAssinatura:e.target.value})}
                  style={{ width:"100%",padding:"8px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}/>
              </div>
              <div>
                <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:5,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Valor mensal R$</div>
                <input type="number" min="0" step="0.01" placeholder="Ex: 89.90" value={c.valorMensal||""} onChange={e=>save({valorMensal:parseFloat(e.target.value)||0})}
                  style={{ width:"100%",padding:"8px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}/>
              </div>
            </div>
            {assin&&(()=>{
              const vm = parseFloat(c.valorMensal)||0;
              const ciclosPagos = c.cancelado ? calcCiclosCancelado(c.dataInicioAssinatura, c.dataCancelamento) : assin.cicloAtual;
              const ltvPago = vm * ciclosPagos;
              const ciclosRestantes = c.cancelado ? 0 : assin.ciclosTotais - assin.cicloNoPeriodo;
              const ltvProjetado = ltvPago + vm * ciclosRestantes;
              const ltvTotal = ltvPago + (c.gasto||0);
              const ltvTotalProjetado = ltvProjetado + (c.gasto||0);
              return (vm > 0 && (
                <div style={{ background:C.tealL,border:"0.5px solid "+C.teal,borderRadius:10,padding:"12px 14px",marginBottom:12 }}>
                  <div style={{ fontSize:11,fontWeight:500,color:C.tealD,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:10 }}>LTV do cliente</div>
                  <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:8 }}>
                    <div style={{ background:"#fff",borderRadius:8,padding:"8px 10px" }}>
                      <div style={{ fontSize:9,color:"var(--color-text-tertiary)",marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em" }}>Pre-assinatura</div>
                      <div style={{ fontSize:14,fontWeight:500,color:"var(--color-text-primary)" }}>R${(c.gasto||0).toFixed(0)}</div>
                    </div>
                    <div style={{ background:"#fff",borderRadius:8,padding:"8px 10px" }}>
                      <div style={{ fontSize:9,color:"var(--color-text-tertiary)",marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em" }}>Club pago</div>
                      <div style={{ fontSize:14,fontWeight:500,color:C.tealD }}>R${ltvPago.toFixed(0)}</div>
                      <div style={{ fontSize:9,color:"var(--color-text-tertiary)",marginTop:1 }}>{assin.cicloAtual}x R${vm}</div>
                    </div>
                    <div style={{ background:"#fff",borderRadius:8,padding:"8px 10px" }}>
                      <div style={{ fontSize:9,color:"var(--color-text-tertiary)",marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em" }}>LTV atual</div>
                      <div style={{ fontSize:14,fontWeight:500,color:C.tealD }}>R${ltvTotal.toFixed(0)}</div>
                    </div>
                    <div style={{ background:C.tealL,borderRadius:8,padding:"8px 10px",border:"0.5px solid "+C.teal }}>
                      <div style={{ fontSize:9,color:C.tealD,marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em",fontWeight:500 }}>LTV projetado</div>
                      <div style={{ fontSize:14,fontWeight:500,color:C.tealD }}>R${ltvTotalProjetado.toFixed(0)}</div>
                      <div style={{ fontSize:9,color:C.tealD,marginTop:1 }}>+{ciclosRestantes}x R${vm}</div>
                    </div>
                  </div>
                </div>
              ));
            })()}
            {assin&&(
              <div>
                <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:8,marginBottom:8 }}>
                  <div style={{ background:C.purpleL,borderRadius:8,padding:"10px 12px" }}>
                    <div style={{ fontSize:10,color:C.purpleD,marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em",fontWeight:500 }}>Ciclo total</div>
                    <div style={{ fontSize:22,fontWeight:500,color:C.purpleD }}>{assin.cicloAtual}°<span style={{ fontSize:11,fontWeight:400,color:C.purple }}> mes</span></div>
                    <div style={{ fontSize:10,color:C.purpleD,marginTop:2 }}>{assin.periodoAtual}° renovacao · mes {assin.cicloNoPeriodo}/{assin.ciclosTotais} do periodo</div>
                  </div>
                  <div style={{ background:corCobranca+"18",borderRadius:8,padding:"10px 12px",border:"0.5px solid "+corCobranca }}>
                    <div style={{ fontSize:10,color:corCobranca,marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em",fontWeight:500 }}>Proxima cobranca</div>
                    <div style={{ fontSize:13,fontWeight:500,color:corCobranca }}>{assin.proximaCobranca}</div>
                    <div style={{ fontSize:10,color:corCobranca,marginTop:2 }}>em {assin.diasParaCobranca} dias</div>
                  </div>
                  <div style={{ background:corFim+"18",borderRadius:8,padding:"10px 12px",border:"0.5px solid "+corFim }}>
                    <div style={{ fontSize:10,color:corFim,marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em",fontWeight:500 }}>Fim da fidelidade</div>
                    <div style={{ fontSize:13,fontWeight:500,color:corFim }}>{assin.fimPeriodo}</div>
                    <div style={{ fontSize:10,color:corFim,marginTop:2 }}>em {assin.diasParaFim} dias</div>
                  </div>
                  <div style={{ background:"var(--color-background-secondary)",borderRadius:8,padding:"10px 12px" }}>
                    <div style={{ fontSize:10,color:"var(--color-text-tertiary)",marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em" }}>Status</div>
                    <div style={{ fontSize:12,fontWeight:500,color:c.cancelado?C.coralD:assin.ativa?C.greenD:"var(--color-text-tertiary)" }}>{c.cancelado?"Cancelado":assin.ativa?"Ativa":"Aguardando"}</div>
                  </div>
                </div>
                {!c.cancelado&&assin.diasParaFim<=30&&<div style={{ background:C.coralL,border:"0.5px solid "+C.coral,borderRadius:8,padding:"8px 12px",fontSize:12,color:C.coralD,fontWeight:500 }}>⚠ Fidelidade encerra em {assin.diasParaFim} dias — hora de trabalhar a renovacao!</div>}
                <div style={{ marginTop:10,padding:"10px 12px",background:c.cancelado?C.coralL:"var(--color-background-secondary)",borderRadius:8,border:"0.5px solid "+(c.cancelado?C.coral:"var(--color-border-tertiary)") }}>
                  <div style={{ display:"flex",alignItems:"center",gap:10 }}>
                    <div style={{ flex:1 }}>
                      <div style={{ fontSize:11,fontWeight:500,color:c.cancelado?C.coralD:"var(--color-text-primary)",marginBottom:2 }}>{c.cancelado?"✗ Assinatura cancelada":"Marcar como cancelada"}</div>
                      {c.cancelado&&<div style={{ fontSize:11,color:C.coralD }}>Data: {c.dataCancelamento||"—"} · LTV congelado em R${(calcCiclosCancelado(c.dataInicioAssinatura,c.dataCancelamento)*(parseFloat(c.valorMensal)||0)+(c.gasto||0)).toFixed(0)}</div>}
                    </div>
                    {!c.cancelado&&(
                      <button onClick={()=>{
                        const hoje2=new Date().toISOString().split("T")[0];
                        save({cancelado:true,dataCancelamento:hoje2});
                      }} style={{ padding:"6px 12px",borderRadius:8,fontSize:11,fontWeight:500,cursor:"pointer",background:C.coral,color:"#fff",border:"none" }}>
                        Registrar cancelamento
                      </button>
                    )}
                    {c.cancelado&&(
                      <button onClick={()=>save({cancelado:false,dataCancelamento:""})} style={{ padding:"6px 12px",borderRadius:8,fontSize:11,fontWeight:500,cursor:"pointer",background:"none",color:C.tealD,border:"0.5px solid "+C.teal }}>
                        Reativar
                      </button>
                    )}
                  </div>
                </div>
                {!c.cancelado&&(
                  <div style={{ marginTop:8,padding:"10px 12px",background:c.falhaRenovacao?C.amberL:"var(--color-background-secondary)",borderRadius:8,border:"0.5px solid "+(c.falhaRenovacao?C.amber:"var(--color-border-tertiary)") }}>
                    <div style={{ display:"flex",alignItems:"center",gap:10 }}>
                      <div style={{ flex:1 }}>
                        <div style={{ fontSize:11,fontWeight:500,color:c.falhaRenovacao?C.amberD:"var(--color-text-primary)",marginBottom:2 }}>{c.falhaRenovacao?"⚠ Falha na renovacao":"Marcar falha na renovacao"}</div>
                        {c.falhaRenovacao&&<div style={{ fontSize:11,color:C.amberD }}>Registrada em {c.dataFalhaRenovacao||"—"} · Nao contabilizado no MRR</div>}
                      </div>
                      {!c.falhaRenovacao&&(
                        <button onClick={()=>save({falhaRenovacao:true,dataFalhaRenovacao:new Date().toLocaleDateString("pt-BR")})}
                          style={{ padding:"6px 12px",borderRadius:8,fontSize:11,fontWeight:500,cursor:"pointer",background:C.amber,color:"#fff",border:"none" }}>
                          Registrar falha
                        </button>
                      )}
                      {c.falhaRenovacao&&(
                        <button onClick={()=>save({falhaRenovacao:false,dataFalhaRenovacao:""})}
                          style={{ padding:"6px 12px",borderRadius:8,fontSize:11,fontWeight:500,cursor:"pointer",background:"none",color:C.tealD,border:"0.5px solid "+C.teal }}>
                          Resolver
                        </button>
                      )}
                    </div>
                  </div>
                )}
              </div>
            )}
            {!c.tipoAssinatura&&<div style={{ fontSize:12,color:"var(--color-text-tertiary)",textAlign:"center",padding:"8px 0" }}>Selecione o tipo de assinatura para ver os calculos automaticos.</div>}
          </div>
        );
      })()}

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

  // Foco Club: clientes com objetivo club ou falta_uma (candidatos proximos ao club)
  const focoClub = clientes.filter(c=>c.objetivo==="club"||c.objetivo==="falta_uma");
  const focoClubAtivos = focoClub.filter(c=>c.etapa!=="encerrado"&&c.etapa!=="convertido").length;

  // Taxa de conversao: club encerrado / total Foco Club que passou pelo Primeiro Contato
  const focoContatados = focoClub.filter(c=>c.etapa!=="lead").length;
  const totalClubHist = conversoes.filter(c=>c.resultado==="club").length;
  const taxa = focoContatados > 0 ? totalClubHist/focoContatados : 0;
  const taxaPct = Math.round(taxa*100);

  // Leads Foco Club necessarios para bater a meta
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
          <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>{totalClubHist} club / {focoContatados} foco club contatados</div>
        </div>
        <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 12px",borderLeft:"3px solid "+C.amber}}>
          <div style={{fontSize:10,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Leads necessarios p/ meta</div>
          <div style={{fontSize:20,fontWeight:500,color:C.amberD}}>{faltam===0?"Meta batida":leadsNecessarios}</div>
          <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>{faltam===0?"":"foco club p/ converter +"+faltam}</div>
        </div>
        <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 12px",borderLeft:"3px solid "+(focoClubAtivos>=(typeof leadsNecessarios==="number"?leadsNecessarios:0)?C.green:C.coral)}}>
          <div style={{fontSize:10,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Foco Club ativos</div>
          <div style={{fontSize:20,fontWeight:500,color:focoClubAtivos>=(typeof leadsNecessarios==="number"?leadsNecessarios:0)?C.greenD:C.coralD}}>{focoClubAtivos}</div>
          <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>{typeof leadsNecessarios==="number"&&focoClubAtivos>=leadsNecessarios?"suficiente p/ meta":typeof leadsNecessarios==="number"&&faltam>0?"faltam "+(leadsNecessarios-focoClubAtivos)+" leads foco":"use filtro 🎯 Foco Club"}</div>
        </div>
      </div>
          );
        })()}
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




const TIPOS_ASSINATURA = [
  {id:"trimestral", label:"Trimestral", ciclosTotais:3},
  {id:"semestral",  label:"Semestral",  ciclosTotais:6},
  {id:"anual",      label:"Anual",      ciclosTotais:12},
];

const addMeses = (data, meses) => {
  const d = new Date(data);
  d.setMonth(d.getMonth() + meses);
  return d;
};

const calcAssinatura = (tipo, dataInicio) => {
  if (!tipo || !dataInicio) return null;
  const t = TIPOS_ASSINATURA.find(t=>t.id===tipo);
  if (!t) return null;
  const inicio = new Date(dataInicio + "T12:00:00");
  const hoje = new Date();

  // Ciclo mensal atual (1-indexed)
  // Quantos meses completos se passaram desde o inicio
  let cicloAtual = 0;
  let proxData = new Date(inicio);
  while (proxData <= hoje) {
    cicloAtual++;
    proxData = addMeses(inicio, cicloAtual);
  }
  cicloAtual = Math.max(1, cicloAtual);

  // Proxima cobranca mensal
  const proximaCobranca = addMeses(inicio, cicloAtual);

  // Fim do periodo de fidelidade atual
  // Quantos periodos completos ja se passaram
  const ciclosTotais = t.ciclosTotais;
  const periodosCompletos = Math.floor((cicloAtual - 1) / ciclosTotais);
  const fimPeriodoAtual = addMeses(inicio, (periodosCompletos + 1) * ciclosTotais);

  // Ciclo dentro do periodo (ex: mes 2 de 12 no anual)
  const cicloNoPeriodo = ((cicloAtual - 1) % ciclosTotais) + 1;

  const fmt = (d) => d.toLocaleDateString("pt-BR");
  const diasParaCobranca = Math.ceil((proximaCobranca.getTime() - hoje.getTime()) / 86400000);
  const diasParaFim = Math.ceil((fimPeriodoAtual.getTime() - hoje.getTime()) / 86400000);

  const periodoAtual = Math.ceil(cicloAtual / ciclosTotais); // qual renovacao (1=original, 2=1a renovacao...)
  return {
    cicloAtual,        // total de meses desde o inicio
    cicloNoPeriodo,    // mes dentro do periodo atual (1 a ciclosTotais)
    ciclosTotais,
    periodoAtual,      // numero da renovacao atual
    proximaCobranca: fmt(proximaCobranca),
    proximaCobrancaISO: proximaCobranca.toISOString().split("T")[0], // para ordenacao
    fimPeriodo: fmt(fimPeriodoAtual),
    diasParaCobranca,
    diasParaFim,
    ativa: hoje >= inicio,
  };
};

const calcCiclosCancelado = (dataInicio, dataCancelamento) => {
  // Quantos ciclos mensais foram cobrados ate o cancelamento
  if (!dataInicio || !dataCancelamento) return 0;
  const inicio = new Date(dataInicio + "T12:00:00");
  const cancel = new Date(dataCancelamento + "T12:00:00");
  let ciclos = 0;
  let prox = new Date(inicio);
  while (prox <= cancel) {
    ciclos++;
    prox = addMeses(inicio, ciclos);
  }
  return Math.max(0, ciclos - 1);
};

const prioScore = (c) => {
  // Stage bonus
  const stageBonus = {em_conversa:1000, proposta_feita:1000, primeiro_contato:200, lead:0, convertido:0, encerrado:0}[c.etapa] || 0;
  // Objective bonus
  const objBonus = {falta_uma:500, club:400, habit_rebuild:150, reativacao:200}[c.objetivo] || 0;
  // Ciclo adjustment: club with ciclo <= 60 gets full bonus, > 60 gets reduced
  const cicloAdj = (c.objetivo === "club" && c.cicloMedio > 60) ? -100 : 0;
  // Probability
  const prob = c.prob || 0;
  return stageBonus + objBonus + cicloAdj + prob;
};

const Kanban = ({ onAbrir }) => {
  const [clientes,setClientes]=useState([]); const [loading,setLoading]=useState(true); const [conversoes,setConversoes]=useState([]);
  const [pages,setPages]=useState({}); // {etapaId_grupo: pageIndex}
  const [filtroHoje,setFiltroHoje]=useState(false);
  const [filtroClub,setFiltroClub]=useState(false);
  const [draggedId,setDraggedId]=useState(null);
  const [dragOverEtapa,setDragOverEtapa]=useState(null);
  const [busca,setBusca]=useState("");
  const [menuAberto,setMenuAberto]=useState(null);
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

  // Auto-refresh a cada 60 segundos
  useEffect(()=>{
    const timer = setInterval(()=>{ carregar(); }, 60000);
    return ()=>clearInterval(timer);
  },[carregar]);

  const handleDrop = async (etapaId) => {
    if (!draggedId || draggedId === etapaId) { setDraggedId(null); setDragOverEtapa(null); return; }
    const cl = clientes.find(c=>c.id===draggedId);
    if (!cl || cl.etapa === etapaId) { setDraggedId(null); setDragOverEtapa(null); return; }
    const atualizado = {...cl, etapa:etapaId};
    setClientes(prev=>prev.map(c=>c.id===draggedId?atualizado:c));
    setDraggedId(null); setDragOverEtapa(null);
    try { await dbSave(atualizado); } catch(e) {}
  };

  const hoje=new Date().toISOString().split("T")[0];
  const amanha=new Date(Date.now()+86400000).toISOString().split("T")[0];

  const filtrar=(lista)=>{
    let r = lista;
    if(filtroClub) r = r.filter(c=>c.objetivo==="club"||c.objetivo==="falta_uma");
    if(filtroHoje) r = r.filter(c=>c.dataProximoContato===hoje);
    if(!busca.trim()) return r;
    const q=busca.toLowerCase();
    return r.filter(c=>(c.nome||"").toLowerCase().includes(q)||(c.customerId||"").toLowerCase().includes(q)||(c.telefone||"").toLowerCase().includes(q));
  };

  const porEtapa=(id)=>{
    const grupo=filtrar(clientes.filter(c=>c.etapa===id));
    // Sort by priority score (objective + stage + probability)
    // Experiencia: sort by next billing date ascending (closest first)
    const byPrio = (a,b) => {
      if (etapaId === "experiencia") {
        const assinA = calcAssinatura(a.tipoAssinatura, a.dataInicioAssinatura);
        const assinB = calcAssinatura(b.tipoAssinatura, b.dataInicioAssinatura);
        const dA = assinA ? assinA.proximaCobrancaISO : "9999";
        const dB = assinB ? assinB.proximaCobrancaISO : "9999";
        return dA > dB ? 1 : dA < dB ? -1 : 0;
      }
      return prioScore(b) - prioScore(a);
    };
    const byDateThenPrio = (a,b) => a.dataProximoContato > b.dataProximoContato ? 1 : a.dataProximoContato < b.dataProximoContato ? -1 : prioScore(b) - prioScore(a);
    const vencidos=grupo.filter(c=>c.dataProximoContato&&c.dataProximoContato<hoje).sort(byDateThenPrio);
    const deHoje=grupo.filter(c=>c.dataProximoContato===hoje).sort(byPrio);
    const deAmanha=grupo.filter(c=>c.dataProximoContato===amanha).sort(byPrio);
    const depois=grupo.filter(c=>c.dataProximoContato&&c.dataProximoContato>amanha).sort(byDateThenPrio);
    const semData=grupo.filter(c=>!c.dataProximoContato).sort(byPrio);
    return {vencidos,deHoje,deAmanha,depois,semData,total:grupo.length};
  };

  const Card=({cl})=>{
    const v=cl.dataProximoContato&&cl.dataProximoContato<hoje;
    const u=cl.dataProximoContato===hoje;
    const am=cl.dataProximoContato===amanha;
    return (
      <button
        onClick={()=>{ if(!draggedId) onAbrir(cl.id); }}
        draggable={true}
        onDragStart={(e)=>{ e.dataTransfer.effectAllowed="move"; setDraggedId(cl.id); }}
        onDragEnd={()=>{ setDraggedId(null); setDragOverEtapa(null); }}
        style={{ width:"100%",textAlign:"left",background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)",borderLeft:"3px solid "+cl.probCor,borderRadius:8,padding:"10px",marginBottom:6,cursor:draggedId===cl.id?"grabbing":"grab",opacity:draggedId===cl.id?0.4:1,transition:"opacity 0.15s" }}>
        <div style={{ display:"flex",alignItems:"flex-start",gap:4 }}>
          <div style={{ flex:1,minWidth:0 }}>
            <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:2,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis" }}>{cl.customerId?"#"+cl.customerId+" · ":""}{cl.nome}</div>
            <div style={{ fontSize:13,fontWeight:500,color:"var(--color-text-primary)",marginBottom:4,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis" }}>{cl.proximaAcao||"—"}</div>
          </div>
          <div style={{ position:"relative",flexShrink:0 }} onClick={e=>e.stopPropagation()}>
            <button onClick={e=>{e.stopPropagation();setMenuAberto(menuAberto===cl.id?null:cl.id);}}
              style={{ background:"none",border:"none",cursor:"pointer",padding:"0 4px",fontSize:14,color:"var(--color-text-tertiary)",lineHeight:1 }}>⋯</button>
            {menuAberto===cl.id&&(
              <div style={{ position:"absolute",right:0,top:"100%",background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:8,boxShadow:"0 4px 16px rgba(0,0,0,0.12)",zIndex:50,minWidth:160,overflow:"hidden" }}>
                <div style={{ fontSize:10,fontWeight:500,color:"var(--color-text-tertiary)",padding:"6px 10px 4px",textTransform:"uppercase",letterSpacing:"0.06em" }}>Mover para</div>
                {ETAPAS.filter(e=>e.id!==cl.etapa).map(e=>(
                  <button key={e.id} onClick={async()=>{ setMenuAberto(null); const atualizado={...cl,etapa:e.id}; setClientes(prev=>prev.map(c=>c.id===cl.id?atualizado:c)); try{await dbSave(atualizado);}catch(err){} }}
                    style={{ width:"100%",textAlign:"left",padding:"7px 10px",background:"none",border:"none",cursor:"pointer",fontSize:12,color:"var(--color-text-primary)",display:"flex",alignItems:"center",gap:6 }}>
                    <span>{e.emoji}</span><span>{e.label}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
        <div style={{ display:"flex",alignItems:"center",gap:4,marginBottom:cl.dataProximoContato?4:0 }}>
          <span style={{ fontSize:11,background:cl.probCor+"22",color:cl.probCor,padding:"1px 6px",borderRadius:20,fontWeight:500 }}>{cl.prob}%</span>
          <span style={{ fontSize:10,color:"var(--color-text-tertiary)" }}>{cl.p}p · {cl.cicloMedio||"?"}d</span>
        </div>
        {cl.dataProximoContato&&<div style={{ fontSize:10,color:v?C.coralD:u||am?C.amberD:"var(--color-text-tertiary)",background:v?C.coralL:u||am?C.amberL:"transparent",padding:v||u||am?"1px 5px":0,borderRadius:4 }}>{v?"⚠ Vencida":u?"⚡ Hoje":am?"📅 Amanhã":"📅"} {!u&&!am&&new Date(cl.dataProximoContato+"T12:00:00").toLocaleDateString("pt-BR")}</div>}
        {cl.lista&&<div style={{ fontSize:10,color:C.purpleD,marginTop:3 }}>{cl.lista}</div>}
        {cl.etapa==="experiencia"&&cl.tipoAssinatura&&cl.dataInicioAssinatura&&(()=>{
          const assin=calcAssinatura(cl.tipoAssinatura,cl.dataInicioAssinatura);
          if(!assin)return null;
          const d=assin.diasParaCobranca;
          if(d>15)return null;
          return <div style={{ fontSize:10,fontWeight:500,color:d<=7?C.coralD:C.amberD,background:d<=7?C.coralL:C.amberL,padding:"1px 5px",borderRadius:4,marginTop:3 }}>🔔 Renovacao em {d}d</div>;
        })()}
      </button>
    );
  };

  const PER_PAGE = 20;
  const getPage = (key) => pages[key]||0;
  const setPage = (key,p) => setPages(prev=>({...prev,[key]:p}));

  const PaginaCards = ({items, pageKey}) => {
    const pg = getPage(pageKey);
    const totalPages = Math.ceil(items.length/PER_PAGE);
    const slice = items.slice(pg*PER_PAGE, (pg+1)*PER_PAGE);
    return (
      <div>
        {slice.map(cl=><Card key={cl.id} cl={cl}/>)}
        {totalPages > 1 && (
          <div style={{ display:"flex",alignItems:"center",justifyContent:"space-between",padding:"4px 2px",marginTop:2 }}>
            <button onClick={()=>setPage(pageKey,pg-1)} disabled={pg===0}
              style={{ fontSize:11,padding:"3px 8px",borderRadius:6,border:"0.5px solid var(--color-border-tertiary)",background:"none",cursor:pg===0?"default":"pointer",opacity:pg===0?0.3:1 }}>←</button>
            <span style={{ fontSize:10,color:"var(--color-text-tertiary)" }}>{pg*PER_PAGE+1}–{Math.min((pg+1)*PER_PAGE,items.length)} de {items.length}</span>
            <button onClick={()=>setPage(pageKey,pg+1)} disabled={pg>=totalPages-1}
              style={{ fontSize:11,padding:"3px 8px",borderRadius:6,border:"0.5px solid var(--color-border-tertiary)",background:"none",cursor:pg>=totalPages-1?"default":"pointer",opacity:pg>=totalPages-1?0.3:1 }}>→</button>
          </div>
        )}
      </div>
    );
  };

  const GrupoFixo=({label,cor,items,pageKey})=>{
    if(items.length===0) return null;
    return (
      <div style={{ marginBottom:6 }}>
        <div style={{ fontSize:9,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.08em",color:cor,padding:"3px 4px",marginBottom:4 }}>{label} · {items.length}</div>
        <PaginaCards items={items} pageKey={pageKey}/>
      </div>
    );
  };


  const GrupoSanfona=({etapaId,grupo,label,cor,items})=>{
    if(items.length===0) return null;
    const aberto=isAberto(etapaId,grupo);
    const pageKey=etapaId+"_"+grupo;
    return (
      <div style={{ marginBottom:4 }}>
        <button onClick={()=>toggleGrupo(etapaId,grupo)} style={{ width:"100%",display:"flex",alignItems:"center",gap:6,padding:"4px 4px",background:"none",border:"none",cursor:"pointer",borderRadius:6 }}>
          <span style={{ fontSize:9,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.08em",color:cor,flex:1,textAlign:"left" }}>{label} · {items.length}</span>
          <span style={{ fontSize:10,color:"var(--color-text-tertiary)" }}>{aberto?"▲":"▼"}</span>
        </button>
        {aberto&&<div style={{ marginTop:4 }}><PaginaCards items={items} pageKey={pageKey}/></div>}
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
    <div onClick={()=>setMenuAberto(null)}>
      <Dashboard clientes={clientes} conversoes={conversoes}/>
      <div style={{ display:"flex",alignItems:"center",gap:10,marginBottom:14 }}>
        <div style={{ position:"relative",flex:1 }}>
          <span style={{ position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",fontSize:14,pointerEvents:"none" }}>🔍</span>
          <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="Buscar por nome, ID ou telefone..." style={{ width:"100%",padding:"8px 12px 8px 32px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}/>
        </div>
        <div style={{ fontSize:13,color:"var(--color-text-tertiary)",whiteSpace:"nowrap" }}>
          {clientes.length} clientes
          {!filtroClub && (() => { const fc=clientes.filter(c=>c.objetivo==="club"||c.objetivo==="falta_uma").length; return fc>0?<span style={{ marginLeft:6,fontSize:11,background:C.greenL,color:C.greenD,padding:"1px 7px",borderRadius:20,fontWeight:500 }}>{fc} foco club</span>:null; })()}
        </div>
        <button onClick={()=>setFiltroClub(f=>!f)} style={{ padding:"5px 14px",borderRadius:8,fontSize:12,fontWeight:500,background:filtroClub?C.green:"var(--color-background-secondary)",border:"0.5px solid "+(filtroClub?C.green:"var(--color-border-tertiary)"),color:filtroClub?C.greenD:"var(--color-text-secondary)",cursor:"pointer",whiteSpace:"nowrap" }}>
          {filtroClub?"🎯 Foco Club ×":"🎯 Foco Club"}
        </button>
        <button onClick={()=>setFiltroHoje(f=>!f)} style={{ padding:"5px 14px",borderRadius:8,fontSize:12,fontWeight:500,background:filtroHoje?C.amber:"var(--color-background-secondary)",border:"0.5px solid "+(filtroHoje?C.amber:"var(--color-border-tertiary)"),color:filtroHoje?C.amberD:"var(--color-text-secondary)",cursor:"pointer",whiteSpace:"nowrap" }}>
          {filtroHoje?"⚡ Hoje ×":"⚡ Hoje"}
        </button>
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
                <div
                  onDragOver={(e)=>{ e.preventDefault(); e.dataTransfer.dropEffect="move"; setDragOverEtapa(etapa.id); }}
                  onDragLeave={(e)=>{ if(!e.currentTarget.contains(e.relatedTarget)) setDragOverEtapa(null); }}
                  onDrop={(e)=>{ e.preventDefault(); handleDrop(etapa.id); }}
                  style={{ border:"0.5px solid "+etapa.cor,borderTop:"none",borderRadius:"0 0 10px 10px",padding:8,minHeight:80,background:dragOverEtapa===etapa.id?etapa.corL+"cc":"var(--color-background-primary)",transition:"background 0.15s",outline:dragOverEtapa===etapa.id?"2px dashed "+etapa.cor:"none",outlineOffset:-2 }}>
                  <GrupoFixo label="⚠ Vencido" cor={C.coralD} items={vencidos} pageKey={etapa.id+"_vencido"}/>
                  <GrupoFixo label="⚡ Hoje" cor={C.amberD} items={deHoje} pageKey={etapa.id+"_hoje"}/>
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
  // Excel may strip leading zero from SP CEPs (01xxx → 1xxx as number)
  // Pad to 8 digits to restore leading zero
  const raw = String(cep||"").replace(/\D/g,"");
  if (!raw) return null;
  const c = raw.padStart(8,"0");
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

        // Carregar existentes para checar duplicatas
        const existentes = await dbGetAll();
        const existentesPorId = {};
        existentes.forEach(c => { if (c.customerId) existentesPorId[String(c.customerId).trim()] = c; });

        const novos = [];
        const paraAtualizarLista = []; // clientes existentes que ganham nova lista

        prev.forEach(cl => {
          const cid = cl.customerId ? String(cl.customerId).trim() : null;
          const existente = cid ? existentesPorId[cid] : null;

          if (existente) {
            // Já existe — acrescentar lista se for nova e diferente
            if (cl.lista && cl.lista.trim()) {
              const listaAtual = existente.lista || "";
              const listas = listaAtual.split(" · ").map(l=>l.trim()).filter(Boolean);
              if (!listas.includes(cl.lista.trim())) {
                const novaLista = listaAtual ? listaAtual + " · " + cl.lista.trim() : cl.lista.trim();
                paraAtualizarLista.push({...existente, lista: novaLista});
              }
            }
          } else {
            // Novo cliente — criar com triagem
            const temDados = !!(cl.dp && cl.fora !== null && cl.ped >= 1);
            const tr = temDados ? runTriagem(cl.ped, cl.dp, cl.ped===1?cl.dp:(cl.du||cl.dp), cl.fora, cl.gasto||0) : null;
            novos.push({
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
            });
          }
        });

        if (novos.length === 0 && paraAtualizarLista.length === 0) {
          setErro("Todos os clientes já existem e nenhuma lista nova foi encontrada.");
          setImp(false); setProg(null);
          return;
        }

        const totalOps = novos.length + paraAtualizarLista.length;
        setProg({ atual: Math.floor(totalOps * 0.1), total: totalOps, inicio });

        // Inserir novos em lotes de 200
        const LOTE_SIZE = 200;
        let salvos = 0;
        for (let i = 0; i < novos.length; i += LOTE_SIZE) {
          await dbBulkSave(novos.slice(i, i + LOTE_SIZE));
          salvos += Math.min(LOTE_SIZE, novos.length - i);
          setProg({ atual: Math.floor(salvos/totalOps*85)+5, total: totalOps, inicio });
        }

        // Atualizar lista dos existentes em lotes de 200
        for (let i = 0; i < paraAtualizarLista.length; i += LOTE_SIZE) {
          await dbBulkSave(paraAtualizarLista.slice(i, i + LOTE_SIZE));
          salvos += Math.min(LOTE_SIZE, paraAtualizarLista.length - i);
          setProg({ atual: Math.floor(salvos/totalOps*85)+5, total: totalOps, inicio });
        }

        setProg({ atual: totalOps, total: totalOps, inicio });

        const comTriagem = novos.filter(c=>c.datasPreenchidas).length;
        const semTriagem = novos.length - comTriagem;
        const partes = [];
        if (novos.length > 0) partes.push(novos.length + " novos importados");
        if (comTriagem > 0) partes.push(comTriagem + " com triagem completa");
        if (semTriagem > 0) partes.push(semTriagem + " aguardando datas");
        if (paraAtualizarLista.length > 0) partes.push(paraAtualizarLista.length + " listas atualizadas");
        const msg = partes.join(" · ");

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
          // raw:true keeps numbers (including Excel dates as serials)
          // but we need to force CEP column (index 7) as string to preserve leading zero
          const rawRows = XLSX.utils.sheet_to_json(ws, { header:1, defval:"", raw:true });
          const rows = rawRows.map(row => row.map((cell, ci) => {
            // CEP column (index 7): if number with 7-8 digits, treat as CEP string
            if (ci === 7 && typeof cell === "number" && cell > 9999 && cell < 100000000) {
              return String(Math.round(cell)).padStart(8,"0");
            }
            return cell;
          }));
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


const Guia = () => {
  const Section = ({title, children}) => (
    <div style={{ marginBottom:24 }}>
      <div style={{ fontSize:14,fontWeight:500,color:"var(--color-text-primary)",marginBottom:10,paddingBottom:6,borderBottom:"0.5px solid var(--color-border-tertiary)" }}>{title}</div>
      {children}
    </div>
  );
  const Item = ({label, value}) => (
    <div style={{ display:"flex",gap:12,marginBottom:8,fontSize:13 }}>
      <div style={{ minWidth:180,fontWeight:500,color:"var(--color-text-primary)",flexShrink:0 }}>{label}</div>
      <div style={{ color:"var(--color-text-secondary)",lineHeight:1.5 }}>{value}</div>
    </div>
  );
  const Block = ({title,cor,children}) => (
    <div style={{ background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",marginBottom:10,borderLeft:"3px solid "+cor }}>
      <div style={{ fontSize:12,fontWeight:500,color:cor,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.06em" }}>{title}</div>
      {children}
    </div>
  );
  return (
    <div style={{ maxWidth:700 }}>
      <div style={{ background:C.tealL,border:"0.5px solid "+C.teal,borderRadius:10,padding:"14px 16px",marginBottom:24 }}>
        <div style={{ fontSize:14,fontWeight:500,color:C.tealD,marginBottom:4 }}>Guia de uso — Laricas CRM</div>
        <div style={{ fontSize:12,color:C.tealD,lineHeight:1.6 }}>Este guia documenta as premissas, logicas e regras do sistema para garantir consistencia entre os operadores.</div>
      </div>

      <Section title="🎯 Objetivos de conversao">
        <Block title="Reativacao → 2a compra" cor={C.teal}>
          <Item label="Quando" value="Cliente com 1 pedido, ou com 2 pedidos e ciclo > 60 dias"/>
          <Item label="Objetivo" value="Gerar a proxima compra — nao oferecer Club ainda"/>
          <Item label="Abordagem" value="Curadoria personalizada com base no produto comprado. Sugerir o proximo sabor."/>
          <Item label="Cupom" value="VOLTA10 — valido 5 dias. Nunca sem prazo."/>
        </Block>
        <Block title="Falta Uma → 3a compra (aha moment)" cor={C.amber}>
          <Item label="Quando" value="2 pedidos com ciclo ≤ 60 dias"/>
          <Item label="Objetivo" value="Gerar a 3a compra — e so entao oferecer Club"/>
          <Item label="Abordagem" value="Curadoria dos dois produtos ja comprados + sugestao do terceiro. Nao revelar intencao de venda ainda."/>
          <Item label="Prioridade" value="Maxima — janela fecha se ciclo passar de 90 dias sem contato"/>
        </Block>
        <Block title="Club — habito formado" cor={C.green}>
          <Item label="Quando" value="3+ pedidos com ciclo ≤ 90 dias"/>
          <Item label="Objetivo" value="Converter para assinatura Club"/>
          <Item label="Angulos" value="3o pedido: emocional (qual foi o favorito?). 4-6o: financeiro (calculo de economia). 7o+: surpresa (ainda nao tem o Club?)."/>
          <Item label="Calculo" value="Mostrar total gasto + frete acumulado vs preco do Club. Numeros reais do Shopify."/>
          <Item label="Preco" value="Desconto de 20% so em reuniao presencial — nunca no WhatsApp."/>
        </Block>
        <Block title="Reconstruir habito → Club so depois" cor={C.coral}>
          <Item label="Quando" value="3+ pedidos com ciclo > 90 dias"/>
          <Item label="Objetivo" value="Reativar a compra — Club so apos nova compra"/>
          <Item label="Atencao" value="Nao oferecer Club neste fluxo. Foco em gerar a proxima compra primeiro."/>
        </Block>
      </Section>

      <Section title="📊 Probabilidade de conversao">
        <Item label="Alta (≥ 40%)" value="Falta Uma com ciclo curto, ou cliente com muitos pedidos e ciclo regular"/>
        <Item label="Media (25–39%)" value="Club com ciclo longo, ou reativacao quente (< 30 dias)"/>
        <Item label="Baixa (< 25%)" value="Reativacao fria, ou ciclo muito longo"/>
        <Item label="Modificadores positivos" value="Fora de SP (+20%) — frete pesa mais na decisao pelo Club. Gasto total alto (+35%)."/>
        <Item label="Modificadores negativos" value="Fora da janela de 30 dias (-28%). Gasto total baixo (-15%)."/>
        <Item label="Teto" value="Maxima de 72% — nenhum lead e certeza de conversao."/>
      </Section>

      <Section title="📋 Prioridade no Kanban">
        <Item label="1° Em Conversa / Proposta Feita" value="Conversa ja iniciada — risco de esfriar. Prioridade maxima."/>
        <Item label="2° Falta Uma" value="Janela curta. Cada dia sem contato reduz chance."/>
        <Item label="3° Club habito" value="Candidata natural — abordagem tranquila mas eficiente."/>
        <Item label="4° Data vencida" value="Compromisso do operador que nao foi cumprido."/>
        <Item label="5° Alta probabilidade" value="Modelo indica boa chance independente do objetivo."/>
        <Item label="6° Reativacao / Reconstruir" value="Menor ROI de tempo. Trabalhar depois das prioridades acima."/>
        <Item label="Filtro Hoje" value="Botao ⚡ Ver hoje — exibe apenas clientes com data de contato = hoje, com a ordenacao de prioridade acima."/>
      </Section>

      <Section title="💬 Regras de abordagem">
        <Item label="Primeiro contato" value="Mover para 'Primeiro Contato' so apos enviar a primeira mensagem."/>
        <Item label="Cupom VOLTA10" value="Sempre com prazo de 5 dias. Nunca oferecer sem prazo."/>
        <Item label="Preco do Club" value="Usar preco cheio no WhatsApp. Desconto de 20% so como fechamento em reuniao."/>
        <Item label="Frase de impacto" value="Pacientes satisfeitos esquecem. Pacientes encantados indicam."/>
        <Item label="Sequencia" value="T1 abertura → T2 curadoria ou calculo → T3 objecoes → T4 fechamento. Nunca pular etapas."/>
        <Item label="Club apos compra" value="So oferecer Club apos a 3a compra (aha moment). Excecao: cliente com 7+ pedidos pode receber abordagem em qualquer etapa."/>
      </Section>

      <Section title="📥 Importacao de listas">
        <Item label="Formato" value="Customer ID, Nome, Telefone, Total Gasto, Nº Pedidos, Data 1° Pedido, Data Ultimo Pedido, CEP, Lista"/>
        <Item label="CEP" value="Iniciados em 0 (01xxx–09xxx) = SP / Grande SP. Demais = Fora de SP."/>
        <Item label="Datas" value="Aceita AAAA-MM-DD, DD/MM/AAAA ou serial do Excel."/>
        <Item label="Duplicata" value="Cliente com mesmo Customer ID: lista e acrescentada no perfil. Dados existentes nao sao alterados."/>
        <Item label="Triagem automatica" value="Se tiver datas + CEP, o sistema calcula ciclo, objetivo e gera sequencia automaticamente."/>
      </Section>

      <Section title="💾 Backup e versoes">
        <Item label="Dados" value="Exportar JSON antes de qualquer importacao grande (aba Backup → Exportar backup JSON)."/>
        <Item label="Codigo" value="Vercel guarda historico de todos os deploys. Em caso de problema, ir em Deployments → 3 pontinhos → Promote to Production."/>
        <Item label="Restaurar dados" value="Aba Backup → Restaurar backup → selecionar o arquivo JSON exportado."/>
      </Section>

      <Section title="🎯 Metas Club 2026">
        {[["Maio","20"],["Junho","36"],["Julho","36"],["Agosto","36"],["Setembro","71"],["Outubro","36"],["Novembro","71"],["Dezembro","20"]].map(([m,v])=>(
          <Item key={m} label={m+" 2026"} value={v+" novos assinantes Club"}/>
        ))}
      </Section>
    </div>
  );
};


const LTV = ({ onAbrir }) => {
  const [assinantes, setAssinantes] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    dbGetAssinantes().then(a => { setAssinantes(a); setLoading(false); });
  }, []);

  if (loading) return <div style={{textAlign:"center",padding:40,color:"var(--color-text-tertiary)"}}>Carregando...</div>;

  if (assinantes.length === 0) return (
    <div style={{textAlign:"center",padding:"48px 24px",background:"var(--color-background-secondary)",borderRadius:12,border:"0.5px dashed var(--color-border-tertiary)"}}>
      <div style={{fontSize:32,marginBottom:12}}>💰</div>
      <div style={{fontSize:14,fontWeight:500,color:"var(--color-text-primary)",marginBottom:6}}>Sem assinantes ainda</div>
      <div style={{fontSize:13,color:"var(--color-text-secondary)"}}>Mova clientes para a etapa ⭐ Experiencia para calcular o LTV.</div>
    </div>
  );

  const ativos = assinantes.filter(a=>!a.cancelado&&!a.falhaRenovacao);
  const comFalha = assinantes.filter(a=>a.falhaRenovacao&&!a.cancelado);
  const cancelados = assinantes.filter(a=>a.cancelado);
  const comValor = ativos.filter(a=>a.valorMensal>0);
  const ltvRealizadoAtivos = comValor.reduce((acc,a) => {
    const assin = calcAssinatura(a.tipoAssinatura, a.dataInicioAssinatura);
    return acc + (a.valorMensal||0) * (assin?assin.cicloAtual:0) + (a.gasto||0);
  }, 0);
  const ltvCancelados = cancelados.filter(a=>a.valorMensal>0).reduce((acc,a) => {
    const ciclos = calcCiclosCancelado(a.dataInicioAssinatura, a.dataCancelamento);
    return acc + (a.valorMensal||0) * ciclos + (a.gasto||0);
  }, 0);
  const ltvPagoTotal = ltvRealizadoAtivos + ltvCancelados;
  const ltvProjetadoTotal = comValor.reduce((acc,a) => {
    const assin = calcAssinatura(a.tipoAssinatura, a.dataInicioAssinatura);
    const vm = a.valorMensal||0;
    const ciclosRestantes = assin ? assin.ciclosTotais - assin.cicloNoPeriodo : 0;
    return acc + vm * (assin?assin.cicloAtual:0) + vm * ciclosRestantes + (a.gasto||0);
  }, 0);
  const mrr = comValor.reduce((acc,a) => acc + (a.valorMensal||0), 0);
  const semValor = ativos.length - comValor.length;

  // Evolucao do MRR mes a mes
  // Para cada mes desde o primeiro assinante ate hoje,
  // recalcula o MRR somando quem estava ativo naquele mes
  const mrrEvolucao = (() => {
    if (assinantes.length === 0) return [];
    const datas = assinantes
      .filter(a=>a.dataInicioAssinatura)
      .map(a=>new Date(a.dataInicioAssinatura+"T12:00:00"));
    if (datas.length === 0) return [];
    const minData = new Date(Math.min(...datas.map(d=>d.getTime())));
    const hoje2 = new Date();
    const meses = [];
    let cursor = new Date(minData.getFullYear(), minData.getMonth(), 1);
    while (cursor <= hoje2 && meses.length < 24) {
      const mesKey = cursor.toISOString().substring(0,7);
      const mesLabel = cursor.toLocaleDateString("pt-BR",{month:"short",year:"2-digit"});
      const mrr_mes = assinantes.filter(a => {
        if (!a.dataInicioAssinatura || !a.valorMensal) return false;
        const inicio = new Date(a.dataInicioAssinatura+"T12:00:00");
        if (inicio > cursor) return false;
        if (a.cancelado && a.dataCancelamento) {
          const cancel = new Date(a.dataCancelamento+"T12:00:00");
          const fimMes = new Date(cursor.getFullYear(), cursor.getMonth()+1, 0);
          if (cancel < cursor) return false;
        }
        return true;
      }).reduce((acc,a) => acc + (parseFloat(a.valorMensal)||0), 0);
      const novos = assinantes.filter(a => {
        if (!a.dataInicioAssinatura) return false;
        return new Date(a.dataInicioAssinatura+"T12:00:00").toISOString().substring(0,7) === mesKey;
      }).length;
      const canc = assinantes.filter(a => a.cancelado && a.dataCancelamento && a.dataCancelamento.substring(0,7) === mesKey).length;
      meses.push({ mesKey, mesLabel, mrr: Math.round(mrr_mes), novos, canc });
      cursor = new Date(cursor.getFullYear(), cursor.getMonth()+1, 1);
    }
    return meses;
  })();

  // Churn: cancelamentos por mes
  const churnPorMes = {};
  cancelados.forEach(a => {
    if (!a.dataCancelamento) return;
    const key = a.dataCancelamento.substring(0,7);
    churnPorMes[key] = (churnPorMes[key]||0) + 1;
  });
  const mesesComChurn = Object.keys(churnPorMes).sort().reverse().slice(0,3);
  const churnMesAtual = churnPorMes[new Date().toISOString().substring(0,7)]||0;
  const tempoMedioMeses = cancelados.length > 0
    ? Math.round(cancelados.filter(a=>a.dataInicioAssinatura&&a.dataCancelamento).reduce((acc,a)=>{
        return acc + calcCiclosCancelado(a.dataInicioAssinatura, a.dataCancelamento);
      },0) / cancelados.filter(a=>a.dataInicioAssinatura&&a.dataCancelamento).length)
    : 0;
  const churnRate = ativos.length > 0 ? Math.round(cancelados.length/(ativos.length+cancelados.length)*100) : 0;

  return (
    <div>
      <div style={{background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px",marginBottom:16}}>
        <div style={{fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:12}}>
          Dash Club — {ativos.length} ativos
          {comFalha.length>0&&<span style={{color:C.amberD,marginLeft:6}}>· {comFalha.length} falha renovacao</span>}
          {cancelados.length>0&&<span style={{color:C.coralD,marginLeft:6}}>· {cancelados.length} cancelados</span>}
          {semValor>0&&<span style={{color:C.amber,marginLeft:6}}>· {semValor} sem valor</span>}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:10}}>
          <div style={{background:C.purpleL,borderRadius:10,padding:"12px 14px",borderLeft:"3px solid "+C.purple}}>
            <div style={{fontSize:10,color:C.purpleD,marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em",fontWeight:500}}>Assinantes ativos</div>
            <div style={{fontSize:28,fontWeight:500,color:C.purpleD}}>{ativos.length}</div>
            <div style={{fontSize:10,color:C.purpleD,marginTop:2}}>{cancelados.length>0?cancelados.length+" cancelado"+(cancelados.length>1?"s":""):"nenhum cancelamento"}</div>
          </div>
          <div style={{background:"var(--color-background-primary)",borderRadius:10,padding:"12px 14px",borderLeft:"3px solid "+C.teal}}>
            <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>MRR — ativos</div>
            <div style={{fontSize:24,fontWeight:500,color:C.tealD}}>R${mrr.toLocaleString("pt-BR",{minimumFractionDigits:0})}</div>
            <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>{comValor.length} assinantes · media R${comValor.length>0?(mrr/comValor.length).toFixed(0):0}/mes</div>
          </div>
          <div style={{background:"var(--color-background-primary)",borderRadius:10,padding:"12px 14px",borderLeft:"3px solid "+C.green}}>
            <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>LTV realizado</div>
            <div style={{fontSize:24,fontWeight:500,color:C.greenD}}>R${ltvPagoTotal.toLocaleString("pt-BR",{minimumFractionDigits:0})}</div>
            <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>ativos + cancelados</div>
          </div>
          <div style={{background:"var(--color-background-primary)",borderRadius:10,padding:"12px 14px",borderLeft:"3px solid "+C.purple}}>
            <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>LTV projetado</div>
            <div style={{fontSize:24,fontWeight:500,color:C.purpleD}}>R${ltvProjetadoTotal.toLocaleString("pt-BR",{minimumFractionDigits:0})}</div>
            <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>ativos ate fim do plano</div>
          </div>
        </div>
      </div>

      {cancelados.length > 0 && (
        <div style={{background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px",marginBottom:16}}>
          <div style={{fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:12}}>Cancelamentos e churn</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
            <div style={{background:"var(--color-background-primary)",borderRadius:10,padding:"12px 14px",borderLeft:"3px solid "+C.coral}}>
              <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>Churn total</div>
              <div style={{fontSize:24,fontWeight:500,color:C.coralD}}>{churnRate}%</div>
              <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>{cancelados.length} de {ativos.length+cancelados.length} assinantes</div>
            </div>
            <div style={{background:"var(--color-background-primary)",borderRadius:10,padding:"12px 14px",borderLeft:"3px solid "+C.amber}}>
              <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>Tempo medio ate cancelar</div>
              <div style={{fontSize:24,fontWeight:500,color:C.amberD}}>{tempoMedioMeses || "—"}{tempoMedioMeses?"m":""}</div>
              <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>meses de assinatura</div>
            </div>
            <div style={{background:"var(--color-background-primary)",borderRadius:10,padding:"12px 14px",borderLeft:"3px solid "+C.purple}}>
              <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>Cancelamentos mes atual</div>
              <div style={{fontSize:24,fontWeight:500,color:churnMesAtual>0?C.coralD:"var(--color-text-primary)"}}>{churnMesAtual}</div>
              <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>{mesesComChurn.slice(1).map(m=><span key={m} style={{marginRight:6}}>{m}: {churnPorMes[m]}</span>)}</div>
            </div>
          </div>
        </div>
      )}

      {mrrEvolucao.length > 1 && (
        <div style={{background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px",marginBottom:16}}>
          <div style={{fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:14}}>Evolucao do MRR</div>
          {(()=>{
            const [mesSel, setMesSel] = useState(null);
            const maxMrr = Math.max(...mrrEvolucao.map(m=>m.mrr), 1);
            const barW = Math.max(32, Math.min(60, Math.floor(560/mrrEvolucao.length)));
            const mesDetalhes = mesSel ? assinantes.filter(a=>{
              if (!a.dataInicioAssinatura||!a.valorMensal) return false;
              const inicio = new Date(a.dataInicioAssinatura+"T12:00:00");
              const [y,mo] = mesSel.split("-").map(Number);
              const dMes = new Date(y,mo-1,1);
              if (inicio > dMes) return false;
              if ((a.cancelado||a.falhaRenovacao)&&a.dataCancelamento) {
                const cancel = new Date((a.dataCancelamento||a.dataFalhaRenovacao)+"T12:00:00");
                if (cancel < dMes) return false;
              }
              return true;
            }) : [];
            return (
              <div>
                <div style={{overflowX:"auto"}}>
                  <div style={{display:"flex",alignItems:"flex-end",gap:4,minWidth:"max-content",paddingTop:28,paddingBottom:8}}>
                    {mrrEvolucao.map((m,i)=>{
                      const pct=m.mrr/maxMrr;
                      const isLast=i===mrrEvolucao.length-1;
                      const isSel=mesSel===m.mesKey;
                      const prev=i>0?mrrEvolucao[i-1].mrr:m.mrr;
                      const cresceu=m.mrr>=prev;
                      const barColor=isSel?C.purple:isLast?C.teal:cresceu?C.green:C.coral;
                      return (
                        <div key={m.mesKey} onClick={()=>setMesSel(isSel?null:m.mesKey)}
                          style={{display:"flex",flexDirection:"column",alignItems:"center",gap:2,width:barW,cursor:"pointer",position:"relative"}}>
                          {(m.novos>0||m.canc>0)&&(
                            <div style={{position:"absolute",top:-22,display:"flex",gap:2,zIndex:1}}>
                              {m.novos>0&&<span style={{fontSize:7,color:C.greenD,background:C.greenL,padding:"1px 3px",borderRadius:3,whiteSpace:"nowrap"}}>+{m.novos}</span>}
                              {m.canc>0&&<span style={{fontSize:7,color:C.coralD,background:C.coralL,padding:"1px 3px",borderRadius:3,whiteSpace:"nowrap"}}>-{m.canc}</span>}
                            </div>
                          )}
                          <div style={{fontSize:9,fontWeight:500,color:isSel?C.purpleD:isLast?C.tealD:"var(--color-text-tertiary)",textAlign:"center",marginBottom:2}}>
                            {m.mrr>0?"R$"+m.mrr:"—"}
                          </div>
                          <div style={{width:"80%",height:Math.max(4,Math.round(pct*110)),background:barColor,borderRadius:"4px 4px 0 0",transition:"height 0.3s",border:isSel?"2px solid "+C.purpleD:"none"}}/>
                          <div style={{fontSize:9,color:isSel?C.purpleD:"var(--color-text-tertiary)",textAlign:"center",textTransform:"capitalize",fontWeight:isSel?500:400}}>{m.mesLabel}</div>
                        </div>
                      );
                    })}
                  </div>
                  <div style={{display:"flex",gap:12,marginTop:4,fontSize:10,color:"var(--color-text-tertiary)"}}>
                    <span style={{color:C.greenD}}>+N novos</span>
                    <span style={{color:C.coralD}}>-N cancelamentos</span>
                    <span style={{color:C.teal}}>● mes atual</span>
                    <span style={{color:C.purple}}>■ selecionado</span>
                    <span>Clique na barra para ver detalhes</span>
                  </div>
                </div>
                {mesSel&&(
                  <div style={{marginTop:12,background:"var(--color-background-primary)",borderRadius:10,padding:"12px 14px",border:"0.5px solid "+C.purple}}>
                    <div style={{fontSize:11,fontWeight:500,color:C.purpleD,marginBottom:8}}>
                      Assinantes ativos em {mrrEvolucao.find(m=>m.mesKey===mesSel)?.mesLabel} ({mesDetalhes.length})
                    </div>
                    {mesDetalhes.length===0?<div style={{fontSize:12,color:"var(--color-text-tertiary)"}}>Nenhum assinante com valor cadastrado neste mes.</div>:
                      mesDetalhes.map(a=>(
                        <div key={a.id} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 0",borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
                          <div style={{flex:1,fontSize:12,fontWeight:500,color:"var(--color-text-primary)"}}>{a.nome}</div>
                          <div style={{fontSize:11,color:"var(--color-text-secondary)"}}>{a.tipoAssinatura||"—"}</div>
                          <div style={{fontSize:12,fontWeight:500,color:C.tealD}}>R${a.valorMensal}/mes</div>
                        </div>
                      ))
                    }
                  </div>
                )}
              </div>
            );
          })()}
        </div>
      )}

      <div style={{background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px"}}>
        {(()=>{
          const [sortCol, setSortCol] = useState("cicloAtual");
          const [sortDir, setSortDir] = useState("desc");
          const toggleSort = (col) => { if(sortCol===col) setSortDir(d=>d==="asc"?"desc":"asc"); else {setSortCol(col);setSortDir("desc");} };
          const Th = ({col,label}) => (
            <th onClick={()=>toggleSort(col)} style={{padding:"7px 10px",textAlign:col==="nome"?"left":"center",fontWeight:500,color:sortCol===col?C.teal:"var(--color-text-tertiary)",fontSize:11,borderBottom:"0.5px solid var(--color-border-tertiary)",textTransform:"uppercase",letterSpacing:"0.05em",cursor:"pointer",userSelect:"none",whiteSpace:"nowrap"}}>
              {label} {sortCol===col?(sortDir==="asc"?"↑":"↓"):""}
            </th>
          );
          const sorted = [...assinantes].sort((a,b)=>{
            let vA,vB;
            if(sortCol==="nome"){vA=a.nome||"";vB=b.nome||""; return sortDir==="asc"?vA.localeCompare(vB):vB.localeCompare(vA);}
            const assinA=calcAssinatura(a.tipoAssinatura,a.dataInicioAssinatura);
            const assinB=calcAssinatura(b.tipoAssinatura,b.dataInicioAssinatura);
            if(sortCol==="cicloAtual"){vA=assinA?assinA.cicloAtual:0;vB=assinB?assinB.cicloAtual:0;}
            else if(sortCol==="valorMensal"){vA=a.valorMensal||0;vB=b.valorMensal||0;}
            else if(sortCol==="ltvAtual"){
              const cA=a.cancelado?calcCiclosCancelado(a.dataInicioAssinatura,a.dataCancelamento):(assinA?assinA.cicloAtual:0);
              const cB=b.cancelado?calcCiclosCancelado(b.dataInicioAssinatura,b.dataCancelamento):(assinB?assinB.cicloAtual:0);
              vA=(a.valorMensal||0)*cA+(a.gasto||0);vB=(b.valorMensal||0)*cB+(b.gasto||0);
            }
            else if(sortCol==="diasCobranca"){vA=assinA?assinA.diasParaCobranca:999;vB=assinB?assinB.diasParaCobranca:999;}
            else{vA=0;vB=0;}
            return sortDir==="asc"?vA-vB:vB-vA;
          });
          return (
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:12}}>
          <div style={{fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",flex:1}}>Assinantes individuais</div>
          <div style={{display:"flex",gap:8}}>
            {["trimestral","semestral","anual"].map(tipo=>{
              const grupo = comValor.filter(a=>a.tipoAssinatura===tipo);
              if(grupo.length===0) return null;
              const mediaVM = Math.round(grupo.reduce((acc,a)=>acc+(a.valorMensal||0),0)/grupo.length);
              const mediaLTV = Math.round(grupo.reduce((acc,a)=>{
                const assin=calcAssinatura(a.tipoAssinatura,a.dataInicioAssinatura);
                return acc+(a.valorMensal||0)*(assin?assin.ciclosTotais:0)+(a.gasto||0);
              },0)/grupo.length);
              return (<div key={tipo} style={{background:"var(--color-background-primary)",borderRadius:8,padding:"6px 10px",textAlign:"center"}}>
                <div style={{fontSize:9,color:"var(--color-text-tertiary)",textTransform:"capitalize",marginBottom:2}}>{tipo} ({grupo.length})</div>
                <div style={{fontSize:12,fontWeight:500,color:"var(--color-text-primary)"}}>R${mediaVM}/mes</div>
                <div style={{fontSize:10,color:C.purpleD}}>LTV ~R${mediaLTV}</div>
              </div>);
            })}
          </div>
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead>
              <tr style={{background:"var(--color-background-primary)"}}>
                  <Th col="nome" label="Nome"/>
                  <Th col="plano" label="Plano"/>
                  <Th col="cicloAtual" label="Ciclo total"/>
                  <Th col="valorMensal" label="R$/mes"/>
                  <Th col="ltvAtual" label="LTV atual"/>
                  <Th col="ltvAtual" label="LTV projetado"/>
                  <Th col="diasCobranca" label="Prox. cobr."/>
                  <th style={{padding:"7px 10px",textAlign:"center",fontWeight:500,color:"var(--color-text-tertiary)",fontSize:11,borderBottom:"0.5px solid var(--color-border-tertiary)",textTransform:"uppercase",letterSpacing:"0.05em"}}>Status</th>
              </tr>
            </thead>
            <tbody>
              {sorted.map(a=>{
                const assin = calcAssinatura(a.tipoAssinatura, a.dataInicioAssinatura);
                const vm = parseFloat(a.valorMensal)||0;
                const ciclosPagos = a.cancelado ? calcCiclosCancelado(a.dataInicioAssinatura,a.dataCancelamento) : (assin?assin.cicloAtual:0);
                const ltvAtual = vm*ciclosPagos+(a.gasto||0);
                const ciclosRest = a.cancelado?0:(assin?assin.ciclosTotais-assin.cicloNoPeriodo:0);
                const ltvProj = ltvAtual+vm*ciclosRest;
                return (
                  <tr key={a.id} style={{borderBottom:"0.5px solid var(--color-border-tertiary)",opacity:a.cancelado?0.6:1}}>
                    <td style={{padding:"7px 10px"}}><button onClick={()=>onAbrir&&onAbrir(a.id)} style={{background:"none",border:"none",cursor:onAbrir?"pointer":"default",fontWeight:500,color:onAbrir?C.teal:"var(--color-text-primary)",fontSize:12,padding:0,textAlign:"left"}}>{a.nome}</button></td>
                    <td style={{padding:"7px 10px",textAlign:"center",color:"var(--color-text-secondary)",textTransform:"capitalize"}}>{a.tipoAssinatura||"—"}</td>
                    <td style={{padding:"7px 10px",textAlign:"center",color:C.purpleD,fontWeight:500}}>{assin?assin.cicloNoPeriodo+"/"+assin.ciclosTotais:"—"}</td>
                    <td style={{padding:"7px 10px",textAlign:"center",color:"var(--color-text-secondary)"}}>{vm>0?"R$"+vm.toFixed(0):"—"}</td>
                    <td style={{padding:"7px 10px",textAlign:"center",fontWeight:500,color:C.greenD}}>{vm>0?"R$"+ltvAtual.toFixed(0):"—"}</td>
                    <td style={{padding:"7px 10px",textAlign:"center",fontWeight:500,color:C.purpleD}}>{vm>0&&!a.cancelado?"R$"+ltvProj.toFixed(0):a.cancelado?"Cancelado":"—"}</td>
                    <td style={{padding:"7px 10px",textAlign:"center"}}><span style={{fontSize:10,fontWeight:500,background:a.cancelado?C.coralL:C.greenL,color:a.cancelado?C.coralD:C.greenD,padding:"2px 8px",borderRadius:20}}>{a.cancelado?"Cancelado":"Ativa"}</span></td>
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


const GlobalSearch = ({ onAbrir }) => {
  const [q, setQ] = useState("");
  const [resultados, setResultados] = useState([]);
  const [buscando, setBuscando] = useState(false);
  const [aberto, setAberto] = useState(false);

  useEffect(() => {
    if (!q.trim() || q.length < 2) { setResultados([]); return; }
    setBuscando(true);
    const timer = setTimeout(async () => {
      const todos = await dbGetAll();
      const ql = q.toLowerCase();
      const res = todos.filter(c =>
        (c.nome||"").toLowerCase().includes(ql) ||
        (c.customerId||"").toLowerCase().includes(ql) ||
        (c.telefone||"").toLowerCase().includes(ql)
      ).slice(0, 8);
      setResultados(res);
      setBuscando(false);
    }, 300);
    return () => clearTimeout(timer);
  }, [q]);

  const etapaInfo = (id) => ETAPAS.find(e=>e.id===id) || ETAPAS[0];

  return (
    <div style={{ position:"relative", flex:1, maxWidth:400 }}>
      <div style={{ position:"relative" }}>
        <span style={{ position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",fontSize:14,pointerEvents:"none" }}>🔍</span>
        <input
          value={q} onChange={e=>{setQ(e.target.value);setAberto(true);}}
          onFocus={()=>setAberto(true)}
          onBlur={()=>setTimeout(()=>setAberto(false),200)}
          placeholder="Buscar cliente em qualquer etapa..."
          style={{ width:"100%",padding:"8px 12px 8px 32px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}
        />
        {buscando&&<span style={{ position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",fontSize:11,color:"var(--color-text-tertiary)" }}>...</span>}
      </div>
      {aberto&&q.length>=2&&(
        <div style={{ position:"absolute",top:"100%",left:0,right:0,background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:10,boxShadow:"0 4px 20px rgba(0,0,0,0.1)",zIndex:100,marginTop:4,overflow:"hidden" }}>
          {resultados.length===0&&!buscando&&<div style={{ padding:"12px 14px",fontSize:13,color:"var(--color-text-tertiary)" }}>Nenhum cliente encontrado</div>}
          {resultados.map(c => {
            const e = etapaInfo(c.etapa);
            return (
              <button key={c.id} onClick={()=>{onAbrir(c.id);setQ("");setAberto(false);}}
                style={{ width:"100%",textAlign:"left",padding:"10px 14px",background:"none",border:"none",borderBottom:"0.5px solid var(--color-border-tertiary)",cursor:"pointer",display:"flex",alignItems:"center",gap:10 }}>
                <div style={{ flex:1 }}>
                  <div style={{ fontSize:13,fontWeight:500,color:"var(--color-text-primary)" }}>{c.nome}</div>
                  <div style={{ fontSize:11,color:"var(--color-text-tertiary)" }}>{c.customerId?"#"+c.customerId+" · ":""}{c.telefone||""}</div>
                </div>
                <span style={{ fontSize:10,fontWeight:500,background:e.corL,color:e.corD,padding:"2px 8px",borderRadius:20,flexShrink:0 }}>{e.emoji} {e.label}</span>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
};


const TriagemTab = ({ onSalvo }) => {
  const [modo, setModo] = useState("historico"); // "historico" | "novo"
  return (
    <div>
      <div style={{ display:"flex",gap:8,marginBottom:20 }}>
        <button onClick={()=>setModo("historico")}
          style={{ flex:1,padding:"10px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",
            background:modo==="historico"?C.teal:"var(--color-background-secondary)",
            color:modo==="historico"?"#fff":"var(--color-text-secondary)",border:"none" }}>
          📊 Cliente com historico Shopify
        </button>
        <button onClick={()=>setModo("novo")}
          style={{ flex:1,padding:"10px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",
            background:modo==="novo"?C.teal:"var(--color-background-secondary)",
            color:modo==="novo"?"#fff":"var(--color-text-secondary)",border:"none" }}>
          ✨ Novo cliente sem compra
        </button>
      </div>
      {modo==="historico"&&<TriagemForm onSalvo={onSalvo}/>}
      {modo==="novo"&&<NovoClienteForm onSalvo={onSalvo}/>}
    </div>
  );
};

export default function App() {
  const [tab,setTab]=useState("kanban");
  const [clienteId,setClienteId]=useState(null);
  const [refresh,setRefresh]=useState(0);
  const abrirClienteGlobal = (id) => { setClienteId(id); setTab("kanban"); };
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
      <div style={{ display:"flex",alignItems:"center",gap:12,padding:"20px 0 8px" }}>
        <div style={{ flex:1 }}>
        <div style={{ fontSize:11,fontWeight:500,letterSpacing:"0.09em",textTransform:"uppercase",color:"var(--color-text-tertiary)",marginBottom:4 }}>Laricas Fitness</div>
        <div style={{ fontSize:22,fontWeight:500,lineHeight:1.3 }}>CRM de Conversão</div>
          <div style={{ fontSize:13,color:"var(--color-text-secondary)",marginTop:4 }}>Lead → Contato → Conversa → Proposta → Convertido</div>
        </div>
        <GlobalSearch onAbrir={abrirClienteGlobal}/>
      </div>
      <div style={{ display:"flex",borderBottom:"0.5px solid var(--color-border-tertiary)",marginBottom:24,overflowX:"auto" }}>
        <T label="📋 Kanban" active={tab==="kanban"} color={C.green} onClick={()=>{setClienteId(null);setTab("kanban");}}/>
        <T label="📥 Importar" active={tab==="import"} color={C.purple} onClick={()=>setTab("import")}/>
        <T label="🎯 Triagem" active={tab==="triagem"} color={C.teal} onClick={()=>setTab("triagem")}/>
        <T label="📊 Historico" active={tab==="historico"} color={C.teal} onClick={()=>setTab("historico")}/>
        <T label="📈 Dash Club" active={tab==="dashclub"} color={C.green} onClick={()=>setTab("dashclub")}/>
        <T label="📖 Guia" active={tab==="guia"} color={C.teal} onClick={()=>setTab("guia")}/>
        <T label="💾 Backup" active={tab==="backup"} color={C.blue} onClick={()=>setTab("backup")}/>
        <T label="⚙ Config" active={tab==="config"} color="var(--color-text-tertiary)" onClick={()=>setTab("config")}/>
      </div>
      {tab==="kanban"&&(clienteId?<Perfil key={clienteId} clienteId={clienteId} onVoltar={()=>{setClienteId(null);setRefresh(r=>r+1);}}/>:<Kanban key={refresh} onAbrir={setClienteId}/>)}
      {tab==="import"&&<ImportarLista onSalvo={onSalvo}/>}
      {tab==="triagem"&&(
        <div>
          <CadastroRapido onSalvo={onSalvo}/>
          <div style={{ display:"flex",alignItems:"center",gap:12,marginBottom:20 }}>
            <div style={{ flex:1,height:1,background:"var(--color-border-tertiary)" }}/>
            <span style={{ fontSize:11,color:"var(--color-text-tertiary)",fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Ou triagem com historico de compras</span>
            <div style={{ flex:1,height:1,background:"var(--color-border-tertiary)" }}/>
          </div>
          <TriagemForm onSalvo={onSalvo}/>
        </div>
      )}
      {tab==="historico"&&<Historico/>}
      {tab==="dashclub"&&<LTV onAbrir={(id)=>{abrirClienteGlobal(id);setTab("kanban");}}/>}
      {tab==="guia"&&<Guia/>}
      {tab==="backup"&&<Backup onRestore={onRestore}/>}
      {tab==="config"&&<ConfigSupabase onSalvo={()=>{ loadCfg().then(()=>setCfgOk(true)); setTab("kanban"); }}/>}
    </div>
  );
}
