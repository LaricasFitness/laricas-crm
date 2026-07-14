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
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 30000);
  try {
    const resp = await fetch(url, {
      method: opts.method || "GET",
      headers: {
        "Authorization": "Bearer " + _SB.key,
        "Content-Type": "application/json",
        ...(opts.pref ? {"Prefer": opts.pref} : {}),
      },
      ...(opts.body !== undefined ? {body: JSON.stringify(opts.body)} : {}),
      signal: controller.signal,
    });
    clearTimeout(timer);
    if (!resp.ok) { const t = await resp.text(); throw new Error("Erro " + resp.status + ": " + t.slice(0,120)); }
    const t = await resp.text();
    return t ? JSON.parse(t) : null;
  } catch(e) {
    clearTimeout(timer);
    if (e.name === "AbortError") throw new Error("Supabase timeout — projeto pode estar inativo. Aguarde 30s e tente novamente.");
    throw e;
  }
};

const dbGetAll  = async () => {
  const PAGE = 1000;
  let all = [], offset = 0;
  while (true) {
    const r = await sb("/clientes?select=id,dados&id=neq.__ultimo_import__&order=atualizado_em.desc&limit="+PAGE+"&offset="+offset);
    const pg = (r||[]).map(x=>x.dados).filter(Boolean).map(fixCliente);
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
const dbGetUltimoImport = async () => {
  try {
    const r = await sb("/clientes?id=eq.__ultimo_import__&select=dados");
    return r&&r[0] ? r[0].dados : null;
  } catch(e) { return null; }
};
const dbSaveUltimoImport = async (info) => {
  try {
    await sb("/clientes", { method:"POST", pref:"resolution=merge-duplicates",
      body:{ id:"__ultimo_import__", dados:info, atualizado_em:new Date().toISOString() } });
  } catch(e) {}
};
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
    if (foraDaJanela) steps.push({ label:"Antes — Reconexão", quem:"Time humano", cor:C.amber,
      copy:"Oi [Nome]! 😊 Aqui é o Lucas da Laricas.\n\nComo tá? Ainda tem Laricas em casa? 😄",
      regra:"Confirmar engajamento antes de qualquer oferta. Tom leve, sem pressão.",
      gatilho:"Resposta positiva → T1. Sem resposta em 48h → T1 direto." });
    steps.push({ label:"T1 — Curadoria personalizada", quem:"Time humano", cor:C.green,
      copy:"Oi [Nome]! 😊 Aqui é o Lucas da Laricas.\n\nVi que você já provou o [produto 1] e o [produto 2] — fico feliz demais que você voltou! 🥰\n\n[Se ciclo próximo: Suas Laricas já devem estar quase no fim por aí! 😄\n\n]Tenho uma sugestão que acho que vai ser seu próximo favorito:\n👉 [próximo sabor sugerido]\n\nQuer saber mais sobre ele?",
      regra:"Ver histórico no Shopify antes de enviar. Indicar UM produto específico, não lista. Incluir a frase de timing se diasUltimo estiver próximo do cicloMedio. Tom: caloroso e próximo, como indicação de amiga.",
      gatilho:"Interesse → T2 com link. Sem resposta em 48h → T3." });
    steps.push({ label:"T2 — Link direto", quem:"Time humano", cor:C.amber,
      copy:"[Nome], que bom! 😄\n\nAqui está o link direto pra você:\n👉 [link do produto]\n\nQualquer dúvida é só me chamar, tô por aqui! 💛",
      regra:"Só enviar após interesse confirmado. SEM cupom neste passo — guardar para T3.",
      gatilho:"Compra → encerrado. Sem compra em 72h → T3." });
    steps.push({ label:"T3 — Follow-up com cupom", quem:"Automação", cor:C.teal,
      copy:"[Nome], deixa eu facilitar pra você 😊\n\n🎁 VOLTA10 — 10% de desconto, válido por 5 dias.\n\n👉 [link do produto]\n\nQualquer coisa é só chamar!",
      regra:"Cupom com prazo real de 5 dias. Nunca prorrogar. Só revelar aqui, não antes.",
      gatilho:"Compra → encerrado. Sem compra → T4." });
    steps.push({ label:"T4 — Encerramento caloroso", quem:"Time humano", cor:C.purple,
      copy:"[Nome], sem problema nenhum! 😊\n\nQuando bater aquela vontade de Laricas, é só me chamar que a gente encontra o sabor certo pra você.\n\nA gente vai estar por aqui! 🍫💛",
      regra:"Sem novo cupom. Encerrar com leveza e porta aberta para contato futuro.",
      gatilho:"Qualquer contato futuro → retoma do T1." });
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


const buildRetencao = (assin, tipoLabel) => {
  const steps = [];

  // R1 — Onboarding (mes 1)
  steps.push({
    label: "R1 — Onboarding", quem: "Time humano", cor: C.green, fase: 1,
    copy: "[Nome]! 😊 Aqui é o Lucas da Laricas.\n\nSua primeira caixinha do Club deve ter chegado — tudo certo por aí?\n\nFico feliz de ter você no Club! 🥰 Me conta: qual foi o que você mais gostou?",
    regra: "Enviar até 3 dias após a data de início. Não mencionar próximo envio ainda. Registrar o favorito nas notas do perfil.",
    gatilho: "Responde com favorito → anotar. Não responde em 48h → aguardar próximo ciclo."
  });

  // R2 — Curadoria mensal (ciclo 2 em diante)
  steps.push({
    label: "R2 — Curadoria mensal", quem: "Time humano", cor: C.teal, fase: 2,
    copy: "[Nome]! 😊 Aqui é o Lucas.\n\nO envio do mês tá chegando! 🍫\n\nQuer manter a seleção atual ou prefere trocar algum produto esse mês?",
    regra: "Enviar ~5 dias antes da renovação mensal. Se não responder, manter seleção padrão e confirmar envio.",
    gatilho: "Quer trocar → personalizar e confirmar. Mantém → confirmar envio normalmente."
  });

  // R3 — Renovação (último mês da fidelidade)
  steps.push({
    label: "R3 — Renovação da fidelidade", quem: "Time humano", cor: C.amber, fase: "ultimo",
    copy: `[Nome]! 😊 Aqui é o Lucas da Laricas.\n\nSeu plano ${tipoLabel} encerra daqui a um mês — o tempo voou! 🥰\n\nQueria já garantir a continuidade pra você não ficar nenhum dia sem Laricas 😄\n\nPosso renovar por mais ${tipoLabel}?`,
    regra: "Abordar 30 dias antes do fim da fidelidade. Tom de cuidado, não de cobrança. Não mencionar preço primeiro.",
    gatilho: "Sim → confirmar renovação e nova data. Quer mudar plano → apresentar opções. Não → entender motivo antes de aceitar."
  });

  // R4 — Win-back (após cancelamento)
  steps.push({
    label: "R4 — Win-back pós cancelamento", quem: "Time humano", cor: C.coral, fase: "cancel",
    copy: "[Nome], tudo bem? 😊 Aqui é o Lucas.\n\nVi que sua assinatura encerrou — espero que tudo esteja bem com você!\n\nQuando bater saudade das Laricas, é só me chamar. A gente resolve 💛",
    regra: "Enviar até 7 dias após cancelamento. Sem oferta, sem desconto — apenas porta aberta. Uma tentativa só.",
    gatilho: "Responde → entender motivo e avaliar win-back personalizado. Silêncio → aguardar 30 dias e tentar uma última vez."
  });

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

  // Correção por tempo desde último pedido — independente do histórico
  // Se último pedido > 90 dias: hábito perdido → reativação
  // Se último pedido 45-90 dias e obj era club/falta_uma: reconstruir hábito primeiro
  if (p > 1 && diasUlt > 90) {
    obj="reativacao"; cor=C.teal; corD=C.tealD;
    label="Reativação — inativa +90d";
    alerta="⛔ Último pedido há "+diasUlt+"d — tratar como cliente nova antes de qualquer oferta de Club.";
  } else if (p > 1 && diasUlt > 45 && (obj==="club" || obj==="falta_uma")) {
    obj="habit_rebuild"; cor=C.coral; corD=C.coralD;
    label="Reconstruir hábito — inativa +45d";
    alerta="⚠ Último pedido há "+diasUlt+"d — reativar compra antes de oferecer Club.";
  }
  const seq = buildSeq(obj, ciclo, p, fora, foraDaJanela, diasUnico);
  const prob = calcProb(obj, ciclo, p, fora, foraDaJanela, gasto, diasUnico);
  return { obj, label, cor, corD, alerta, ciclo, p, fora, foraDaJanela, diasUlt, diasUnico, span, seq, prob };
};


// Aplica correção por tempo de inatividade em tempo real
// Usado para exibição — independente do valor salvo no Supabase
const corrigirObjetivo = (c) => {
  if (!c.dataUltimo || !c.p || c.p < 2) return c;
  const hoje = new Date();
  const dtU = new Date(c.dataUltimo + "T12:00:00");
  const diasUlt = Math.round((hoje - dtU) / 86400000);
  if (diasUlt > 90 && c.objetivo !== "reativacao") {
    return { ...c,
      objetivo: "reativacao",
      objetivoLabel: "Reativação — inativa +" + diasUlt + "d",
      objetivoCor: C.teal,
      objetivoCorD: C.tealD,
      objetivoAlerta: "⛔ Último pedido há " + diasUlt + "d — tratar como cliente nova antes de qualquer oferta de Club.",
    };
  }
  if (diasUlt > 45 && (c.objetivo === "club" || c.objetivo === "falta_uma")) {
    return { ...c,
      objetivo: "habit_rebuild",
      objetivoLabel: "Reconstruir hábito — inativa +" + diasUlt + "d",
      objetivoCor: C.coral,
      objetivoCorD: C.coralD,
      objetivoAlerta: "⚠ Último pedido há " + diasUlt + "d — reativar compra antes de oferecer Club.",
    };
  }
  return c;
};


// Lookup de cidade/estado por prefixo de CEP
const CEP_CIDADES = {
  "01":"São Paulo/SP","02":"São Paulo/SP","03":"São Paulo/SP","04":"São Paulo/SP","05":"São Paulo/SP",
  "06":"Osasco/SP","07":"Guarulhos/SP","08":"São Paulo/SP","09":"Santo André/SP",
  "10":"Santos/SP","11":"Santos/SP","12":"São José dos Campos/SP","13":"Campinas/SP",
  "14":"Ribeirão Preto/SP","15":"São José do Rio Preto/SP","16":"Araçatuba/SP","17":"Bauru/SP",
  "18":"Sorocaba/SP","19":"Presidente Prudente/SP",
  "20":"Rio de Janeiro/RJ","21":"Rio de Janeiro/RJ","22":"Rio de Janeiro/RJ","23":"Rio de Janeiro/RJ",
  "24":"Niterói/RJ","25":"Duque de Caxias/RJ","26":"Nova Iguaçu/RJ","27":"Campos/RJ",
  "28":"Campos/RJ","29":"Vitória/ES",
  "30":"Belo Horizonte/MG","31":"Belo Horizonte/MG","32":"Contagem/MG","33":"Belo Horizonte/MG",
  "34":"Belo Horizonte/MG","35":"Ipatinga/MG","36":"Juiz de Fora/MG","37":"Poços de Caldas/MG",
  "38":"Uberaba/MG","39":"Montes Claros/MG",
  "40":"Salvador/BA","41":"Salvador/BA","42":"Feira de Santana/BA","43":"Ilhéus/BA",
  "44":"Feira de Santana/BA","45":"Vitória da Conquista/BA","46":"Jequié/BA","47":"Barreiras/BA",
  "48":"Paulo Afonso/BA","49":"Aracaju/SE",
  "50":"Recife/PE","51":"Recife/PE","52":"Recife/PE","53":"Olinda/PE","54":"Caruaru/PE",
  "55":"Caruaru/PE","56":"Petrolina/PE","57":"Maceió/AL","58":"João Pessoa/PB","59":"Natal/RN",
  "60":"Fortaleza/CE","61":"Fortaleza/CE","62":"Fortaleza/CE","63":"Juazeiro do Norte/CE",
  "64":"Teresina/PI","65":"São Luís/MA","66":"Belém/PA","67":"Campo Grande/MS",
  "68":"Santarém/PA","69":"Manaus/AM",
  "70":"Brasília/DF","71":"Brasília/DF","72":"Brasília/DF","73":"Brasília/DF",
  "74":"Goiânia/GO","75":"Anápolis/GO","76":"Rio Verde/GO","77":"Palmas/TO",
  "78":"Cuiabá/MT","79":"Campo Grande/MS",
  "80":"Curitiba/PR","81":"Curitiba/PR","82":"Curitiba/PR","83":"Curitiba/PR",
  "84":"Ponta Grossa/PR","85":"Cascavel/PR","86":"Londrina/PR","87":"Maringá/PR",
  "88":"Florianópolis/SC","89":"Joinville/SC",
  "90":"Porto Alegre/RS","91":"Porto Alegre/RS","92":"Canoas/RS","93":"Porto Alegre/RS",
  "94":"Porto Alegre/RS","95":"Caxias do Sul/RS","96":"Pelotas/RS","97":"Santa Maria/RS",
  "98":"Santa Maria/RS","99":"Passo Fundo/RS",
};

const getCidade = (cep) => {
  if (!cep) return null;
  const num = (cep||"").replace(/\D/g,"");
  if (num.length < 2) return null;
  return CEP_CIDADES[num.substring(0,2)] || null;
};


const normalizarTelefone = (tel) => {
  if (!tel) return "";
  // Remove apóstrofo inicial, espaços, hífens, parênteses e outros não-dígitos exceto +
  const limpo = tel.replace(/^'+/, "").replace(/[^\d+]/g, "");
  // Se já tem +55, retorna como está
  if (limpo.startsWith("+55")) return limpo;
  // Se começa com 55 e tem 12-13 dígitos, adiciona +
  if (limpo.startsWith("55") && limpo.length >= 12) return "+" + limpo;
  // Se é número local (10-11 dígitos), adiciona +55
  if (limpo.length >= 10) return "+55" + limpo;
  return limpo;
};

const fixEncoding = (s) => {
  if (!s || typeof s !== "string") return s;
  try {
    // Detecta se tem sequências típicas de UTF-8 lido como Latin-1
    if (!/Ã|Â/.test(s)) return s;
    const bytes = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i);
    return new TextDecoder("utf-8").decode(bytes);
  } catch(e) { return s; }
};

const fixCliente = (c) => {
  if (!c) return c;
  return {
    ...c,
    nome: fixEncoding(c.nome),
    lista: fixEncoding(c.lista),
  };
};

const inp = (ex) => ({ width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:14,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none",...ex });

const T = ({ label, active, color, onClick }) => ( <button onClick={onClick} style={{ padding:"8px 12px",fontSize:12,fontWeight:500,color:active?color:"var(--color-text-secondary)",borderBottom:active?"2px solid "+color:"2px solid transparent",marginBottom:-1,background:"transparent",border:"none",cursor:"pointer",whiteSpace:"nowrap" }}>{label}</button> );
const M = ({ label, value, sub, cor }) => ( <div style={{ background:"var(--color-background-secondary)",borderRadius:8,padding:"10px 12px" }}><div style={{ fontSize:10,color:"var(--color-text-tertiary)",marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em" }}>{label}</div><div style={{ fontSize:18,fontWeight:500,color:cor||"var(--color-text-primary)" }}>{value}</div>{sub&&<div style={{ fontSize:10,color:"var(--color-text-tertiary)",marginTop:2 }}>{sub}</div>}</div> );

const Steps = ({ steps, cur, cliente }) => {
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
                {(()=>{
                  const textoPersonalizado = personalizarCopy(s.copy, cliente);
                  const temPlaceholder = /\[[A-Za-zÀ-ú][^\]]*\]/.test(textoPersonalizado);
                  return (
                    <div>
                      <div style={{ background:"var(--color-background-secondary)",borderRadius:10,padding:"13px 15px",marginBottom:6,fontSize:14,color:"var(--color-text-primary)",lineHeight:1.85,whiteSpace:"pre-line",fontFamily:"inherit",borderLeft:"3px solid "+(isCur?C.teal:s.cor) }}>{textoPersonalizado}</div>
                      {temPlaceholder&&<div style={{ fontSize:10,color:C.amber,marginBottom:6 }}>⚠ Campos em [colchetes] precisam ser preenchidos antes de enviar</div>}
                      <button onClick={()=>{
                        navigator.clipboard.writeText(textoPersonalizado).then(()=>{}).catch(()=>{});
                        const el=document.getElementById("cpbtn_"+i);
                        if(el){el.textContent="✓ Copiado!";el.style.background=C.green;setTimeout(()=>{el.textContent="📋 Copiar mensagem";el.style.background=C.tealL;},2000);}
                      }} id={"cpbtn_"+i}
                        style={{ marginBottom:10,padding:"5px 12px",borderRadius:6,fontSize:11,fontWeight:500,cursor:"pointer",background:C.tealL,color:C.tealD,border:"0.5px solid "+C.teal,transition:"background 0.2s" }}>
                        📋 Copiar mensagem
                      </button>
                    </div>
                  );
                })()}
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




// Detecção de gênero do operador para artigo correto (a/o) nos scripts
const GENERO_NOMES = {
  "lucas":"m","pedro":"m","joao":"m","gabriel":"m","gustavo":"m","rafael":"m","felipe":"m",
  "bruno":"m","diego":"m","thiago":"m","marcelo":"m","andre":"m","fernando":"m","ricardo":"m",
  "ceci":"f","cecilia":"f","luana":"f","maria":"f","ana":"f","julia":"f","marcia":"f",
  "camila":"f","fernanda":"f","beatriz":"f","carolina":"f","leticia":"f","patricia":"f","amanda":"f",
};
const detectarGenero = (nomeOperador) => {
  const n = (nomeOperador||"").toLowerCase().trim().split(" ")[0].normalize("NFD").replace(/[\u0300-\u036f]/g,"");
  if (!n) return "m";
  if (GENERO_NOMES[n]) return GENERO_NOMES[n];
  // Heurística: nomes terminados em "a" geralmente femininos (com exceções comuns)
  const excecoesMasculinas = ["luca","joshua","nathan","kauã","davi"];
  if (excecoesMasculinas.includes(n)) return "m";
  return n.endsWith("a") ? "f" : "m";
};

const personalizarCopy = (texto, cliente) => {
  if (!texto || !cliente) return texto;
  const primeiroNome = (cliente.nome||"").split(" ")[0] || "cliente";
  const nPedidos = cliente.p || 0;
  const gasto = (cliente.gasto||0).toLocaleString("pt-BR",{style:"currency",currency:"BRL",minimumFractionDigits:0});
  const responsavel = (cliente.responsavel||"").trim();
  const operador = responsavel.split(" ")[0] || "Lucas";
  const artigo = responsavel ? (detectarGenero(operador)==="f"?"a":"o") : "o";

  return texto
    .replace(/\[Nome\]/g, primeiroNome)
    .replace(/\[nome\]/g, primeiroNome)
    .replace(/\[N° pedidos\]/g, nPedidos)
    .replace(/\[numero de pedidos\]/g, nPedidos)
    .replace(/R\$\[frete\]/g, "R$XX")
    .replace(/\[gasto total\]/g, gasto)
    .replace(/Aqui é o Lucas/g, "Aqui é "+artigo+" "+operador)
    .replace(/Lucas da Laricas/g, operador+" da Laricas");
};

const LogAtividade = ({ c, save }) => {
  const [logTxt, setLogTxt] = useState("");
  const addLog = () => {
    if (!logTxt.trim()) return;
    const novo = {
      texto: logTxt.trim(),
      data: new Date().toLocaleDateString("pt-BR"),
      hora: new Date().toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"}),
      resp: c.responsavel||"",
    };
    const logs = [novo, ...(c.logAtividade||[])].slice(0,30);
    save({logAtividade: logs});
    setLogTxt("");
  };
  return (
    <div style={{ marginBottom:10 }}>
      <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:6,textTransform:"uppercase",letterSpacing:"0.06em" }}>Log de atividade</div>
      <div style={{ display:"flex",gap:8,marginBottom:8 }}>
        <input value={logTxt} onChange={e=>setLogTxt(e.target.value)}
          onKeyDown={e=>{ if(e.key==="Enter"){e.preventDefault();addLog();} }}
          placeholder="Ex: Tentei contato, nao atendeu. Enviou mensagem..."
          style={{ flex:1,padding:"7px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}/>
        <button onClick={addLog} disabled={!logTxt.trim()}
          style={{ padding:"7px 12px",borderRadius:8,fontSize:12,fontWeight:500,cursor:logTxt.trim()?"pointer":"default",background:logTxt.trim()?C.teal:"var(--color-background-secondary)",color:logTxt.trim()?"#fff":"var(--color-text-tertiary)",border:"none" }}>
          + Registrar
        </button>
      </div>
      {(c.logAtividade||[]).length > 0 && (
        <div style={{ maxHeight:160,overflowY:"auto",display:"flex",flexDirection:"column",gap:4 }}>
          {(c.logAtividade||[]).map((l,i) => (
            <div key={i} style={{ fontSize:11,padding:"5px 8px",background:"var(--color-background-primary)",borderRadius:6,border:"0.5px solid var(--color-border-tertiary)",display:"flex",gap:8 }}>
              <span style={{ color:"var(--color-text-primary)",flex:1 }}>{l.texto}</span>
              <span style={{ color:"var(--color-text-tertiary)",flexShrink:0 }}>{l.data} {l.hora}{l.resp?" · "+l.resp:""}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

const SeqRetencao = ({ c }) => {
  const assin = calcAssinatura(c.tipoAssinatura, c.dataInicioAssinatura, c.cicloAtualClub);
  const tipoLabel = (TIPOS_ASSINATURA.find(t=>t.id===c.tipoAssinatura)?.ciclosTotais||"")+"m";
  const retSteps = buildRetencao(assin, tipoLabel);
  const cicloAtual = assin ? assin.cicloAtual : 1;
  const diasParaFim = assin ? assin.diasParaFim : 999;
  const stepAtivo = c.cancelado ? 3 : diasParaFim <= 30 ? 2 : cicloAtual === 1 ? 0 : 1;
  const [openR, setOpenR] = useState(stepAtivo);

  return (
    <div style={{ background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px",marginBottom:12 }}>
      <div style={{ fontSize:13,fontWeight:500,color:C.purpleD,marginBottom:4 }}>Sequência de retenção</div>
      <div style={{ fontSize:11,color:C.purple,marginBottom:12 }}>
        {c.cancelado?"R4 ativo — win-back":diasParaFim<=30?"⚠ R3 ativo — renovação urgente":cicloAtual===1?"R1 ativo — onboarding":"R2 ativo — curadoria mensal"}
      </div>
      {retSteps.map((s,i)=>{
        const isAtivo = i === stepAtivo;
        const textoP = personalizarCopy(s.copy, c);
        const temPlaceholder = /\[[A-Za-zÀ-ú][^\]]*\]/.test(textoP);
        return (
          <div key={i} style={{ marginBottom:8,borderRadius:10,overflow:"hidden",border:"0.5px solid "+(isAtivo?s.cor:"var(--color-border-tertiary)"),opacity:isAtivo?1:0.6 }}>
            <button onClick={()=>setOpenR(openR===i?-1:i)} style={{ width:"100%",display:"flex",alignItems:"center",gap:10,padding:"10px 14px",background:isAtivo?s.cor+"18":"var(--color-background-secondary)",border:"none",cursor:"pointer",textAlign:"left" }}>
              <span style={{ fontSize:11,fontWeight:600,color:s.cor,flex:1 }}>{s.label}{isAtivo?" ← AGORA":""}</span>
              <span style={{ fontSize:10,color:"var(--color-text-tertiary)",background:"var(--color-background-primary)",padding:"1px 6px",borderRadius:10 }}>{s.quem}</span>
              <span style={{ fontSize:10,color:"var(--color-text-tertiary)" }}>{openR===i?"▲":"▼"}</span>
            </button>
            {openR===i&&(
              <div style={{ padding:"14px",background:"var(--color-background-primary)",borderTop:"0.5px solid var(--color-border-tertiary)" }}>
                <div style={{ background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",marginBottom:6,fontSize:14,color:"var(--color-text-primary)",lineHeight:1.85,whiteSpace:"pre-line",fontFamily:"inherit",borderLeft:"3px solid "+s.cor }}>{textoP}</div>
                {temPlaceholder&&<div style={{ fontSize:10,color:C.amber,marginBottom:6 }}>⚠ Campos em [colchetes] precisam ser preenchidos antes de enviar</div>}
                <button onClick={()=>{ navigator.clipboard.writeText(textoP).then(()=>{}).catch(()=>{}); const el=document.getElementById("rcpbtn_"+i); if(el){el.textContent="✓ Copiado!";el.style.background=C.green;setTimeout(()=>{el.textContent="📋 Copiar mensagem";el.style.background=C.tealL;},2000); }}} id={"rcpbtn_"+i}
                  style={{ marginBottom:10,padding:"5px 12px",borderRadius:6,fontSize:11,fontWeight:500,cursor:"pointer",background:C.tealL,color:C.tealD,border:"0.5px solid "+C.teal,transition:"background 0.2s" }}>
                  📋 Copiar mensagem
                </button>
                <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8 }}>
                  <div><div style={{ fontSize:10,color:"var(--color-text-tertiary)",marginBottom:3,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Regra</div><div style={{ fontSize:12,color:"var(--color-text-secondary)",lineHeight:1.5 }}>{s.regra}</div></div>
                  <div><div style={{ fontSize:10,color:"var(--color-text-tertiary)",marginBottom:3,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em" }}>Próximo gatilho</div><div style={{ fontSize:12,color:"var(--color-text-secondary)",lineHeight:1.5 }}>{s.gatilho}</div></div>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
};

const Perfil = ({ clienteId, onVoltar }) => {
  const [c,setC]=useState(null); const [confirmDel,setConfirmDel]=useState(false); const [salvando,setSalvando]=useState(false); const [toast,setToast]=useState("");
  useEffect(() => { dbGetAll().then(lista => { const cl = lista.find(c=>c.id===clienteId); if(cl) setC(corrigirObjetivo(fixCliente(cl))); }); }, [clienteId]);
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
        {c.telefone&&(
          <a href={"https://wa.me/55"+c.telefone.replace(/\D/g,"")} target="_blank" rel="noopener noreferrer"
            style={{ background:"#25D366",border:"none",borderRadius:6,padding:"4px 10px",fontSize:11,color:"#fff",cursor:"pointer",fontWeight:500,textDecoration:"none",display:"inline-flex",alignItems:"center",gap:4 }}>
            💬 WhatsApp
          </a>
        )}
        {c.datasPreenchidas&&(
          <button onClick={async()=>{
            const tr=runTriagem(c.p,c.dataPrimeiro,c.dataUltimo,c.fora,c.gasto);
            const atualizado={...c,
              objetivo:tr.obj,objetivoLabel:tr.label,objetivoCor:tr.cor,objetivoCorD:tr.corD,
              objetivoAlerta:tr.alerta,prob:tr.prob.pct,probLabel:tr.prob.label,probCor:tr.prob.cor,
              seq:tr.seq,cicloMedio:tr.ciclo,datasPreenchidas:true,
            };
            await save(atualizado);
            setC(corrigirObjetivo(atualizado));
          }} disabled={salvando}
            style={{ background:C.tealL,border:"0.5px solid "+C.teal,borderRadius:6,padding:"4px 10px",fontSize:11,color:C.tealD,cursor:salvando?"default":"pointer",fontWeight:500 }}>
            ↺ Recalcular
          </button>
        )}
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
        {c.statusClub&&c.statusClub!=="perdido"&&(
          <div style={{ display:"inline-flex",alignItems:"center",gap:6,background:C.tealL,borderRadius:6,padding:"3px 10px",marginBottom:6,fontSize:11,color:C.tealD,fontWeight:500 }}>
            🎯 Funil Club: {STATUS_CLUB.find(s=>s.id===c.statusClub)?.emoji} {STATUS_CLUB.find(s=>s.id===c.statusClub)?.label}
          </div>
        )}
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
      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12 }}>
        <div style={{ background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 12px" }}>
          <div style={{ fontSize:10,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4 }}>Nº pedidos</div>
          <div style={{ fontSize:20,fontWeight:500,color:"var(--color-text-primary)" }}>{c.p||0}</div>
        </div>
        <div style={{ background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 12px" }}>
          <div style={{ fontSize:10,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4 }}>Total gasto</div>
          <div style={{ fontSize:20,fontWeight:500,color:C.tealD }}>R${(c.gasto||0).toLocaleString("pt-BR",{minimumFractionDigits:0})}</div>
        </div>
        <div style={{ background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 12px" }}>
          <div style={{ fontSize:10,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4 }}>Ultimo pedido</div>
          <div style={{ fontSize:14,fontWeight:500,color:"var(--color-text-primary)" }}>
            {c.dataUltimo ? new Date(c.dataUltimo+"T12:00:00").toLocaleDateString("pt-BR",{day:"2-digit",month:"short",year:"numeric"}) : "—"}
          </div>
          {c.dataPrimeiro&&c.dataPrimeiro!==c.dataUltimo&&<div style={{ fontSize:10,color:"var(--color-text-tertiary)",marginTop:2 }}>
            1° em {new Date(c.dataPrimeiro+"T12:00:00").toLocaleDateString("pt-BR",{day:"2-digit",month:"short",year:"numeric"})}
          </div>}
        </div>
      </div>
      <div style={{ background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,padding:"16px",marginBottom:12 }}>
        <div style={{ fontSize:13,fontWeight:500,color:"var(--color-text-primary)",marginBottom:12 }}>Perfil do cliente</div>
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10 }}>
          <div><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Nome</div><input style={inp()} value={c.nome} onChange={e=>setC({...c,nome:e.target.value})} onBlur={()=>save({nome:c.nome})} /></div>
          <div><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Telefone / WhatsApp</div><input style={inp()} value={c.telefone||""} onChange={e=>setC({...c,telefone:e.target.value})} onBlur={()=>save({telefone:c.telefone})} placeholder="11 9XXXX-XXXX"/></div>
        </div>
        <div style={{ display:"flex",alignItems:"center",gap:8,marginBottom:10,padding:"7px 12px",background:"var(--color-background-secondary)",borderRadius:8,fontSize:12 }}>
          <span style={{ fontSize:16 }}>{c.fora?"🌎":"📍"}</span>
          <span style={{ color:"var(--color-text-primary)",fontWeight:500 }}>{getCidade(c.cep)||"Cidade não identificada"}</span>
          {c.fora&&<span style={{ fontSize:10,color:C.amberD,background:C.amberL,padding:"1px 6px",borderRadius:10,fontWeight:500 }}>Fora de SP</span>}
          {c.cep?<span style={{ fontSize:11,color:"var(--color-text-tertiary)",marginLeft:4 }}>CEP {c.cep}</span>:<span style={{ fontSize:11,color:"var(--color-text-tertiary)",marginLeft:4 }}>(sem CEP cadastrado)</span>}
        </div>
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10 }}>
          <div><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Email Shopify</div><input style={inp()} value={c.email||""} onChange={e=>setC({...c,email:e.target.value})} onBlur={()=>save({email:c.email})} placeholder="email da conta Shopify" type="email"/></div>
          <div><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Customer ID Shopify</div><input style={inp({fontSize:11,color:"var(--color-text-tertiary)"})} value={c.customerId||""} onChange={e=>setC({...c,customerId:e.target.value})} onBlur={()=>save({customerId:c.customerId})} placeholder="ID do cliente no Shopify"/></div>
        </div>
        <div style={{ marginBottom:10 }}>
          <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Responsável</div>
          <input style={inp()} value={c.responsavel||""} onChange={e=>setC({...c,responsavel:e.target.value})} onBlur={()=>save({responsavel:c.responsavel})} placeholder="Nome do operador"/>
        </div>
        <div style={{ marginBottom:10 }}>
          <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Email Club <span style={{ fontWeight:400,textTransform:"none",letterSpacing:0,color:"var(--color-text-tertiary)" }}>(preferencial para comunicação)</span></div>
          <input style={inp()} value={c.emailClub||""} onChange={e=>setC({...c,emailClub:e.target.value})} onBlur={()=>save({emailClub:c.emailClub})} placeholder="email preferencial para o Club" type="email"/>
        </div>
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10 }}>
          <div>
            <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Foi indicada por</div>
            <input style={inp()} value={c.indicadaPor||""} onChange={e=>setC({...c,indicadaPor:e.target.value})} onBlur={()=>save({indicadaPor:c.indicadaPor})} placeholder="Nome de quem indicou"/>
          </div>
          <div>
            <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Indicou quem</div>
            <input style={inp()} value={c.indicouQuem||""} onChange={e=>setC({...c,indicouQuem:e.target.value})} onBlur={()=>save({indicouQuem:c.indicouQuem})} placeholder="Nome de quem foi indicado"/>
          </div>
        </div>
        <div style={{ marginBottom:10 }}>
          <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Lista de origem (opcional)</div>
          <div style={{ display:"flex",gap:6,flexWrap:"wrap",marginBottom:6 }}>{LISTAS.map(l=>(<button key={l} onClick={()=>save({lista:l===c.lista?"":l})} style={{ padding:"4px 10px",borderRadius:20,fontSize:11,cursor:"pointer",background:c.lista===l?C.purpleL:"var(--color-background-secondary)",color:c.lista===l?C.purpleD:"var(--color-text-secondary)",border:"0.5px solid "+(c.lista===l?C.purple:"var(--color-border-tertiary)") }}>{l}</button>))}</div>
          <input style={inp({fontSize:12})} value={LISTAS.includes(c.lista||"")?"":c.lista||""} onChange={e=>setC({...c,lista:e.target.value})} onBlur={()=>save({lista:c.lista})} placeholder="Ou digite o nome da lista manualmente..."/>
        </div>
        <div style={{ marginBottom:10 }}><div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Anotações</div><textarea value={c.notas} onChange={e=>setC({...c,notas:e.target.value})} onBlur={()=>save({notas:c.notas})} placeholder="Sabor favorito, objeções, contexto..." rows={3} style={{ width:"100%",padding:"8px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none",resize:"vertical",fontFamily:"inherit",lineHeight:1.5 }}/></div>
        <LogAtividade c={c} save={save}/>
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
          <Steps steps={c.seq} cur={c.stepAtual} cliente={c}/>
          {c.stepAtual<c.seq.length-1&&<button onClick={avancar} style={{ width:"100%",marginTop:10,padding:"10px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",background:C.teal,color:"#fff",border:"none" }}>Cliente respondeu → avançar para passo {c.stepAtual+2} ↓</button>}
        </div>
      )}
      {toast&&<div style={{ position:"fixed",top:24,left:"50%",transform:"translateX(-50%)",background:C.green,color:"#fff",padding:"10px 24px",borderRadius:30,fontSize:14,fontWeight:500,zIndex:999,boxShadow:"0 4px 20px rgba(0,0,0,0.15)" }}>{toast}</div>}
      {c.etapa==="experiencia"&&(()=>{
        const assin = calcAssinatura(c.tipoAssinatura, c.dataInicioAssinatura, c.cicloAtualClub);
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
              // LTV da assinatura: usa dado real do RitsPay se disponível, senão estima
              const ltvAssinatura = c.ltvAssinatura
                ? parseFloat(c.ltvAssinatura)
                : vm * ciclosPagos;
              // Ticket médio: usa dado real se disponível
              const ticketMedio = c.ticketMedioClub
                ? parseFloat(c.ticketMedioClub)
                : vm;
              const ciclosRestantes = c.cancelado ? 0 : assin.ciclosTotais - assin.cicloNoPeriodo;
              const ltvProjetadoClub = ltvAssinatura + ticketMedio * ciclosRestantes;
              const ltvTotalAtual = ltvAssinatura + (c.gasto||0);
              const ltvTotalProjetado = ltvProjetadoClub + (c.gasto||0);
              return (vm > 0 && (
                <div style={{ background:C.tealL,border:"0.5px solid "+C.teal,borderRadius:10,padding:"12px 14px",marginBottom:12 }}>
                  <div style={{ fontSize:11,fontWeight:500,color:C.tealD,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:10 }}>LTV do cliente</div>
                  <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:8 }}>
                    <div style={{ background:"#fff",borderRadius:8,padding:"8px 10px" }}>
                      <div style={{ fontSize:9,color:"var(--color-text-tertiary)",marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em" }}>Pré-assinatura</div>
                      <div style={{ fontSize:14,fontWeight:500,color:"var(--color-text-primary)" }}>R${(c.gasto||0).toFixed(0)}</div>
                    </div>
                    <div style={{ background:"#fff",borderRadius:8,padding:"8px 10px" }}>
                      <div style={{ fontSize:9,color:"var(--color-text-tertiary)",marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em" }}>Club pago</div>
                      <div style={{ fontSize:14,fontWeight:500,color:C.tealD }}>R${ltvAssinatura.toFixed(0)}</div>
                      <div style={{ fontSize:9,color:"var(--color-text-tertiary)",marginTop:1 }}>
                        {c.ticketMedioClub ? `ticket médio R$${ticketMedio.toFixed(0)}` : `${assin.cicloAtual}x R$${vm.toFixed(0)}`}
                      </div>
                    </div>
                    <div style={{ background:"#fff",borderRadius:8,padding:"8px 10px" }}>
                      <div style={{ fontSize:9,color:"var(--color-text-tertiary)",marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em" }}>LTV atual (total)</div>
                      <div style={{ fontSize:14,fontWeight:500,color:C.tealD }}>R${ltvTotalAtual.toFixed(0)}</div>
                      <div style={{ fontSize:9,color:"var(--color-text-tertiary)",marginTop:1 }}>club + avulsos</div>
                    </div>
                    <div style={{ background:C.tealL,borderRadius:8,padding:"8px 10px",border:"0.5px solid "+C.teal }}>
                      <div style={{ fontSize:9,color:C.tealD,marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em",fontWeight:500 }}>LTV projetado</div>
                      <div style={{ fontSize:14,fontWeight:500,color:C.tealD }}>R${ltvTotalProjetado.toFixed(0)}</div>
                      <div style={{ fontSize:9,color:C.tealD,marginTop:1 }}>+{ciclosRestantes}x R${ticketMedio.toFixed(0)}</div>
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
                      {c.cancelado&&(
                        <div style={{ marginTop:8 }}>
                          <div style={{ fontSize:10,color:C.coralD,marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em" }}>Motivo do cancelamento</div>
                          <select value={c.motivoCancelamento||""} onChange={e=>save({motivoCancelamento:e.target.value})}
                            style={{ width:"100%",padding:"6px 10px",borderRadius:8,border:"0.5px solid "+C.coral,fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-primary)" }}>
                            <option value="">— Selecione o motivo —</option>
                            {["Preço","Não gostou do produto","Excesso de produto / estoque cheio","Protocolo médico / dieta restritiva","Mudança de rotina","Problema financeiro","Esqueceu / parou de usar","Trocou por concorrente","Problema com entrega","Outro"].map(m=><option key={m} value={m}>{m}</option>)}
                          </select>
                          {c.motivoCancelamento&&<div style={{ fontSize:10,color:C.coralD,marginTop:3 }}>💡 Registrado no Dash</div>}
                        </div>
                      )}
                    </div>
                    {!c.cancelado&&(
                      <button onClick={()=>{
                        const hoje2=new Date().toISOString().split("T")[0];
                        save({cancelado:true,dataCancelamento:hoje2,statusAssinatura:"cancelado"});
                      }} style={{ padding:"6px 12px",borderRadius:8,fontSize:11,fontWeight:500,cursor:"pointer",background:C.coral,color:"#fff",border:"none" }}>
                        Registrar cancelamento
                      </button>
                    )}
                    {c.cancelado&&(
                      <button onClick={()=>save({cancelado:false,dataCancelamento:"",statusAssinatura:"ativo"})} style={{ padding:"6px 12px",borderRadius:8,fontSize:11,fontWeight:500,cursor:"pointer",background:"none",color:C.tealD,border:"0.5px solid "+C.teal }}>
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
                        <button onClick={()=>save({falhaRenovacao:true,dataFalhaRenovacao:new Date().toLocaleDateString("pt-BR"),statusAssinatura:"atrasado"})}
                          style={{ padding:"6px 12px",borderRadius:8,fontSize:11,fontWeight:500,cursor:"pointer",background:C.amber,color:"#fff",border:"none" }}>
                          Registrar falha
                        </button>
                      )}
                      {c.falhaRenovacao&&(
                        <button onClick={()=>save({falhaRenovacao:false,dataFalhaRenovacao:"",statusAssinatura:"ativo"})}
                          style={{ padding:"6px 12px",borderRadius:8,fontSize:11,fontWeight:500,cursor:"pointer",background:"none",color:C.tealD,border:"0.5px solid "+C.teal }}>
                          Resolver
                        </button>
                      )}
                    </div>
                  </div>
                )}
                {/* PAUSA DE ASSINATURA */}
                {!c.cancelado&&(
                  <div style={{ marginTop:8,padding:"10px 12px",background:c.statusAssinatura==="pausado"?C.amberL:"var(--color-background-secondary)",borderRadius:8,border:"0.5px solid "+(c.statusAssinatura==="pausado"?C.amber:"var(--color-border-tertiary)") }}>
                    <div style={{ display:"flex",alignItems:"center",gap:10,marginBottom:c.statusAssinatura==="pausado"?10:0 }}>
                      <div style={{ flex:1 }}>
                        <div style={{ fontSize:11,fontWeight:500,color:c.statusAssinatura==="pausado"?C.amberD:"var(--color-text-primary)",marginBottom:2 }}>
                          {c.statusAssinatura==="pausado"?"⏸ Assinatura pausada":"Marcar como pausada"}
                        </div>
                        {c.statusAssinatura==="pausado"&&c.dataPausaFim&&(
                          <div style={{ fontSize:11,color:C.amberD }}>Retorno em {new Date(c.dataPausaFim+"T12:00:00").toLocaleDateString("pt-BR")}{c.motivoPausa?" · "+c.motivoPausa:""}</div>
                        )}
                      </div>
                      {c.statusAssinatura!=="pausado"&&(
                        <button onClick={()=>save({statusAssinatura:"pausado"})}
                          style={{ padding:"6px 12px",borderRadius:8,fontSize:11,fontWeight:500,cursor:"pointer",background:C.amber,color:"#fff",border:"none" }}>
                          Pausar
                        </button>
                      )}
                      {c.statusAssinatura==="pausado"&&(
                        <button onClick={()=>save({statusAssinatura:"ativo",dataPausaFim:"",motivoPausa:""})}
                          style={{ padding:"6px 12px",borderRadius:8,fontSize:11,fontWeight:500,cursor:"pointer",background:"none",color:C.tealD,border:"0.5px solid "+C.teal }}>
                          Reativar
                        </button>
                      )}
                    </div>
                    {c.statusAssinatura==="pausado"&&(
                      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8 }}>
                        <div>
                          <div style={{ fontSize:10,color:C.amberD,marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em" }}>Data de volta</div>
                          <input type="date" value={c.dataPausaFim||""} onChange={e=>{
                            save({ dataPausaFim: e.target.value, dataProximoContato: e.target.value });
                          }}
                            style={{ width:"100%",padding:"6px 10px",borderRadius:8,border:"0.5px solid "+C.amber,fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-primary)" }}/>
                          <div style={{ fontSize:9,color:C.amber,marginTop:2 }}>Vira follow-up automático</div>
                        </div>
                        <div>
                          <div style={{ fontSize:10,color:C.amberD,marginBottom:3,textTransform:"uppercase",letterSpacing:"0.06em" }}>Motivo</div>
                          <select value={c.motivoPausa||""} onChange={e=>save({motivoPausa:e.target.value})}
                            style={{ width:"100%",padding:"6px 10px",borderRadius:8,border:"0.5px solid "+C.amber,fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-primary)" }}>
                            <option value="">— Selecione —</option>
                            {["Viagem","Protocolo médico / dieta","Questão financeira","Mudança de rotina","Problema com entrega","Estoque cheio","Outro"].map(m=><option key={m} value={m}>{m}</option>)}
                          </select>
                        </div>
                        {c.dataPausaFim&&c.dataPausaFim<=new Date().toISOString().split("T")[0]&&(
                          <div style={{ gridColumn:"1/-1",background:C.greenL,border:"0.5px solid "+C.green,borderRadius:8,padding:"8px 10px",fontSize:11,color:C.greenD,fontWeight:500 }}>
                            ✅ Data de volta chegou — verificar cobrança e reativar!
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}
            {!c.tipoAssinatura&&<div style={{ fontSize:12,color:"var(--color-text-tertiary)",textAlign:"center",padding:"8px 0" }}>Selecione o tipo de assinatura para ver os calculos automaticos.</div>}
          </div>
        );
      })()}

      {c.etapa==="experiencia"&&c.tipoAssinatura&&(()=>{
        const assin = calcAssinatura(c.tipoAssinatura, c.dataInicioAssinatura, c.cicloAtualClub);
        const podeUpsell = c.tipoAssinatura!=="anual" && !c.cancelado && !c.falhaRenovacao && assin && assin.cicloAtual>=2 && assin.diasParaFim>30;
        const proximoPlano = c.tipoAssinatura==="trimestral"?"semestral":c.tipoAssinatura==="semestral"?"anual":null;
        const proximoCiclos = proximoPlano==="semestral"?6:proximoPlano==="anual"?12:0;
        return (
          <>
            {podeUpsell&&proximoPlano&&(
              <div style={{ background:C.greenL,border:"0.5px solid "+C.green,borderRadius:12,padding:"14px 16px",marginBottom:12 }}>
                <div style={{ display:"flex",alignItems:"center",gap:10,marginBottom:8 }}>
                  <div style={{ flex:1 }}>
                    <div style={{ fontSize:13,fontWeight:500,color:C.greenD,marginBottom:2 }}>🚀 Oportunidade de upsell</div>
                    <div style={{ fontSize:12,color:C.greenD }}>Assinante satisfeita no ciclo {assin.cicloAtual} — candidata a migrar para o plano {proximoPlano}</div>
                  </div>
                </div>
                <div style={{ background:"#fff",borderRadius:8,padding:"10px 12px",fontSize:12,color:"var(--color-text-primary)",lineHeight:1.7,whiteSpace:"pre-line",borderLeft:"3px solid "+C.green }}>
                  {`[Nome]! 😊 Aqui é o Lucas da Laricas.

Você já está no ${assin.cicloAtual}° mês do Club — fico tão feliz que está gostando! 🥰

Pensando em você, vi que o plano ${proximoPlano} (${proximoCiclos} meses) faria ainda mais sentido para o seu perfil — e o valor mensal fica menor.

Quer que eu te explique como funciona?`}
                </div>
                <button onClick={()=>{ navigator.clipboard.writeText(`[Nome]! 😊 Aqui é o Lucas da Laricas.

Você já está no ${assin.cicloAtual}° mês do Club — fico tão feliz que está gostando! 🥰

Pensando em você, vi que o plano ${proximoPlano} (${proximoCiclos} meses) faria ainda mais sentido para o seu perfil — e o valor mensal fica menor.

Quer que eu te explique como funciona?`).catch(()=>{}); }}
                  style={{ marginTop:8,padding:"5px 12px",borderRadius:6,fontSize:11,fontWeight:500,cursor:"pointer",background:C.greenL,color:C.greenD,border:"0.5px solid "+C.green }}>
                  📋 Copiar mensagem de upsell
                </button>
              </div>
            )}

                {/* PRODUTOS FAVORITOS */}
                <div style={{ marginTop:8,padding:"10px 12px",background:"var(--color-background-secondary)",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)" }}>
                  <div style={{ fontSize:11,fontWeight:500,color:"var(--color-text-primary)",marginBottom:8 }}>❤️ Produtos favoritos</div>
                  <div style={{ display:"flex",flexWrap:"wrap",gap:6 }}>
                    {["Pão de Mel Brigadeiro","Pão de Mel Beijinho","Pão de Mel Avelã Trufado","Pão de Mel Cookies'n Cream","Bolinho","Barra Recheada","Bombom","Potinho"].map(prod=>{
                      const favs = c.produtosFavoritos||[];
                      const ativo = favs.includes(prod);
                      return (
                        <button key={prod} onClick={()=>{
                          const novos = ativo ? favs.filter(f=>f!==prod) : [...favs,prod];
                          save({produtosFavoritos:novos});
                        }}
                          style={{ padding:"4px 10px",borderRadius:20,fontSize:11,cursor:"pointer",fontWeight:ativo?500:400,
                            background:ativo?C.purple:C.purpleL,color:ativo?"#fff":C.purpleD,
                            border:"0.5px solid "+(ativo?C.purple:C.purpleL) }}>
                          {prod}
                        </button>
                      );
                    })}
                  </div>
                  {(c.produtosFavoritos||[]).length===0&&<div style={{ fontSize:10,color:"var(--color-text-tertiary)",marginTop:4 }}>Clique para marcar os favoritos desta assinante</div>}
                </div>

                {/* CHECKLIST R1-R4 VISUAL */}
                {(()=>{
                  const etapasR = [
                    { id:"r1", label:"R1 — Onboarding", desc:"Até 3 dias após início", campo:"r1Feito", dataCampo:"r1Data", corAtivo:C.teal },
                    { id:"r2", label:"R2 — Como está indo?", desc:"7 a 15 dias", campo:"r2Feito", dataCampo:"r2Data", corAtivo:C.green },
                    { id:"r3", label:"R3 — Feedback da caixa", desc:"25 a 30 dias", campo:"r3Feito", dataCampo:"r3Data", corAtivo:C.purple },
                    { id:"r4", label:"R4 — Renovação / Upsell", desc:"45 a 60 dias", campo:"r4Feito", dataCampo:"r4Data", corAtivo:C.amber },
                  ];
                  const feitos = etapasR.filter(e=>c[e.campo]).length;
                  return (
                    <div style={{ marginTop:8,padding:"10px 12px",background:"var(--color-background-secondary)",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)" }}>
                      <div style={{ display:"flex",alignItems:"center",gap:8,marginBottom:10 }}>
                        <div style={{ fontSize:11,fontWeight:500,color:"var(--color-text-primary)" }}>🔄 Sequência de retenção</div>
                        <div style={{ flex:1,height:4,background:"var(--color-border-tertiary)",borderRadius:2,overflow:"hidden" }}>
                          <div style={{ width:(feitos/4*100)+"%",height:"100%",background:C.teal,borderRadius:2,transition:"width 0.3s" }}/>
                        </div>
                        <div style={{ fontSize:10,color:C.tealD,fontWeight:500 }}>{feitos}/4</div>
                      </div>
                      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:6 }}>
                        {etapasR.map(e=>{
                          const feito = !!c[e.campo];
                          return (
                            <button key={e.id} onClick={()=>{
                              const update = { [e.campo]: !feito };
                              if(!feito) update[e.dataCampo] = new Date().toLocaleDateString("pt-BR");
                              else update[e.dataCampo] = "";
                              save(update);
                            }}
                              style={{ padding:"8px 10px",borderRadius:8,cursor:"pointer",textAlign:"left",
                                background:feito?e.corAtivo+"18":"var(--color-background-primary)",
                                border:"0.5px solid "+(feito?e.corAtivo:"var(--color-border-tertiary)") }}>
                              <div style={{ display:"flex",alignItems:"center",gap:6,marginBottom:2 }}>
                                <span style={{ fontSize:14 }}>{feito?"✅":"○"}</span>
                                <span style={{ fontSize:11,fontWeight:500,color:feito?e.corAtivo:"var(--color-text-primary)" }}>{e.label}</span>
                              </div>
                              <div style={{ fontSize:10,color:"var(--color-text-tertiary)",paddingLeft:20 }}>
                                {feito&&c[e.dataCampo]?c[e.dataCampo]:e.desc}
                              </div>
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  );
                })()}
            <SeqRetencao c={c} save={save}/>
            {/* STATUS DE ASSINATURA */}
            {(()=>{
              const STATUS_ASSN = [
                { id:"ativo",     label:"Ativo",     cor:C.green,  corD:C.greenD,  corL:C.greenL,  emoji:"✅" },
                { id:"pausado",   label:"Pausado",   cor:C.amber,  corD:C.amberD,  corL:C.amberL,  emoji:"⏸" },
                { id:"atrasado",  label:"Atrasado",  cor:C.coral,  corD:C.coralD,  corL:C.coralL,  emoji:"⚠" },
                { id:"cancelado", label:"Cancelado", cor:"#888",   corD:"#555",    corL:"#f0f0f0", emoji:"✗" },
              ];
              const stAtual = STATUS_ASSN.find(s=>s.id===(c.statusAssinatura||"ativo"))||STATUS_ASSN[0];
              const hoje2 = new Date().toISOString().split("T")[0];
              return (
                <div style={{ background:stAtual.corL,border:"0.5px solid "+stAtual.cor,borderRadius:10,padding:"12px 14px",marginTop:12 }}>
                  <div style={{ fontSize:12,fontWeight:500,color:stAtual.corD,marginBottom:10 }}>
                    {stAtual.emoji} Status da assinatura
                  </div>
                  {/* Botões de status */}
                  <div style={{ display:"flex",gap:6,marginBottom:12 }}>
                    {STATUS_ASSN.map(s=>(
                      <button key={s.id} onClick={()=>{
                        const update = { statusAssinatura: s.id };
                        // Se pausado, manter data de volta; outros limpam
                        if(s.id !== "pausado") { update.dataPausaFim = ""; update.motivoPausa = ""; }
                        save(update);
                      }}
                        style={{ flex:1,padding:"7px 4px",borderRadius:8,cursor:"pointer",fontSize:11,fontWeight:500,textAlign:"center",
                          background:(c.statusAssinatura||"ativo")===s.id?s.cor:s.corL,
                          color:(c.statusAssinatura||"ativo")===s.id?"#fff":s.corD,
                          border:"0.5px solid "+s.cor }}>
                        {s.emoji} {s.label}
                      </button>
                    ))}
                  </div>
                  {/* Pausado — data de volta e motivo */}
                  {(c.statusAssinatura==="pausado")&&(
                    <div>
                      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8 }}>
                        <div>
                          <div style={{ fontSize:11,color:C.amberD,marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Data de volta</div>
                          <input type="date" value={c.dataPausaFim||""} onChange={e=>{
                            save({ dataPausaFim: e.target.value, proximoContato: e.target.value });
                          }}
                            style={{ width:"100%",padding:"7px 10px",borderRadius:8,border:"0.5px solid "+C.amber,fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-primary)" }}/>
                          <div style={{ fontSize:10,color:C.amber,marginTop:3 }}>Vira follow-up automático</div>
                        </div>
                        <div>
                          <div style={{ fontSize:11,color:C.amberD,marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Motivo</div>
                          <select value={c.motivoPausa||""} onChange={e=>save({motivoPausa:e.target.value})}
                            style={{ width:"100%",padding:"7px 10px",borderRadius:8,border:"0.5px solid "+C.amber,fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-primary)" }}>
                            <option value="">— Selecione —</option>
                            {["Viagem","Protocolo médico / dieta","Questão financeira","Mudança de rotina","Problema com entrega","Estoque cheio","Outro"].map(m=><option key={m} value={m}>{m}</option>)}
                          </select>
                        </div>
                      </div>
                      {c.dataPausaFim&&c.dataPausaFim<=hoje2&&(
                        <div style={{ background:C.greenL,border:"0.5px solid "+C.green,borderRadius:8,padding:"8px 12px",fontSize:12,color:C.greenD,fontWeight:500 }}>
                          ✅ Data de volta chegou — verificar cobrança e reativar!
                        </div>
                      )}
                    </div>
                  )}
                  {/* Atrasado — orientação */}
                  {c.statusAssinatura==="atrasado"&&(
                    <div style={{ fontSize:12,color:C.coralD,background:"#fff",borderRadius:8,padding:"8px 12px" }}>
                      ⚠ Cobrança com falha — entrar em contato para regularizar.
                    </div>
                  )}
                  {/* Cancelado — orientação */}
                  {c.statusAssinatura==="cancelado"&&(
                    <div style={{ fontSize:12,color:"#555",background:"#fff",borderRadius:8,padding:"8px 12px" }}>
                      ✗ Assinatura encerrada. Considerar reativação futura?
                    </div>
                  )}
                </div>
              );
            })()}
          </>
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

const calcAssinatura = (tipo, dataInicio, cicloRits) => {
  if (!tipo || !dataInicio) return null;
  const t = TIPOS_ASSINATURA.find(t=>t.id===tipo);
  if (!t) return null;
  const inicio = new Date(dataInicio + "T12:00:00");
  const hoje = new Date();

  // Ciclo mensal atual — usa dado real do RitsPay se disponível
  let cicloAtual = 0;
  if (cicloRits && cicloRits > 0) {
    cicloAtual = cicloRits;
  } else {
    let proxData = new Date(inicio);
    while (proxData <= hoje) {
      cicloAtual++;
      proxData = addMeses(inicio, cicloAtual);
    }
    cicloAtual = Math.max(1, cicloAtual);
  }

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

const Kanban = ({ onAbrir, reloadToken, filtroHoje, setFiltroHoje, filtroClub, setFiltroClub, filtroProb, setFiltroProb, filtroPedidos, setFiltroPedidos }) => {
  const [clientes,setClientes]=useState([]); const [loading,setLoading]=useState(true); const [conversoes,setConversoes]=useState([]);
  const [pages,setPages]=useState({});
  // filtroHoje, filtroClub, filtroProb, filtroPedidos — recebidos como props do App
  const [draggedId,setDraggedId]=useState(null);
  const [dragOverEtapa,setDragOverEtapa]=useState(null);
  const [busca,setBusca]=useState("");
  const [menuAberto,setMenuAberto]=useState(null);
  const [abertos,setAbertos]=useState({});
  const [erroCarregar,setErroCarregar]=useState("");
  const toggleGrupo=(etapaId,grupo)=>{ const k=etapaId+"_"+grupo; setAbertos(a=>({...a,[k]:!a[k]})); };
  const isAberto=(etapaId,grupo)=>!!abertos[etapaId+"_"+grupo];

  const carregar = useCallback(async () => {
    setLoading(true);
    try {
      const [todos, conv] = await Promise.all([dbGetAll(), dbGetConversoes()]);
      setClientes(todos); setConversoes(conv);
    } catch(e) {
      setClientes([]);
      setConversoes([]);
      console.error("Erro ao carregar:", e.message);
      // Show error in UI
      setErroCarregar(e.message||"Erro ao conectar com Supabase");
    }
    setLoading(false);
  }, []);
  useEffect(()=>{carregar();},[carregar]);
  useEffect(()=>{ if(reloadToken>0) carregar(); },[reloadToken]);

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
    if(filtroProb==="alta") r = r.filter(c=>(c.prob||0)>=40);
    else if(filtroProb==="media") r = r.filter(c=>(c.prob||0)>=25&&(c.prob||0)<40);
    else if(filtroProb==="baixa") r = r.filter(c=>(c.prob||0)<25);
    if(filtroPedidos==="1") r = r.filter(c=>(c.p||0)===1);
    else if(filtroPedidos==="2") r = r.filter(c=>(c.p||0)===2);
    else if(filtroPedidos==="3+") r = r.filter(c=>(c.p||0)>=3);
    if(!busca.trim()) return r;
    const q=busca.toLowerCase();
    return r.filter(c=>(c.nome||"").toLowerCase().includes(q)||(c.customerId||"").toLowerCase().includes(q)||(c.telefone||"").toLowerCase().includes(q)||(c.email||"").toLowerCase().includes(q)||(c.emailClub||"").toLowerCase().includes(q)||(c.responsavel||"").toLowerCase().includes(q));
  };

  const porEtapa=(id)=>{
    const grupo=filtrar(clientes.filter(c=>c.etapa===id).map(corrigirObjetivo));
    // Sort by priority score (objective + stage + probability)
    // Experiencia: sort by next billing date ascending (closest first)
    const byPrio = (etId) => (a,b) => {
      if (etId === "experiencia") {
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
    const deHoje=grupo.filter(c=>c.dataProximoContato===hoje).sort(byPrio(id));
    const deAmanha=grupo.filter(c=>c.dataProximoContato===amanha).sort(byPrio(id));
    const depois=grupo.filter(c=>c.dataProximoContato&&c.dataProximoContato>amanha).sort(byDateThenPrio);
    const semData=grupo.filter(c=>!c.dataProximoContato).sort(byPrio(id));
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
          {cl.statusClub&&cl.statusClub!=="perdido"&&(
            <span style={{ fontSize:9,fontWeight:500,color:C.tealD,background:C.tealL,padding:"1px 5px",borderRadius:10 }}>
              {STATUS_CLUB.find(s=>s.id===cl.statusClub)?.emoji} Club
            </span>
          )}
        </div>
        {cl.dataProximoContato&&<div style={{ fontSize:10,color:v?C.coralD:u||am?C.amberD:"var(--color-text-tertiary)",background:v?C.coralL:u||am?C.amberL:"transparent",padding:v||u||am?"1px 5px":0,borderRadius:4 }}>{v?"⚠ Vencida":u?"⚡ Hoje":am?"📅 Amanhã":"📅"} {!u&&!am&&new Date(cl.dataProximoContato+"T12:00:00").toLocaleDateString("pt-BR")}</div>}
        {cl.lista&&<div style={{ fontSize:10,color:C.purpleD,marginTop:3 }}>{cl.lista}</div>}
        {cl.etapa==="experiencia"&&cl.tipoAssinatura&&cl.dataInicioAssinatura&&(()=>{
          const assin=calcAssinatura(cl.tipoAssinatura,cl.dataInicioAssinatura);
          if(!assin)return null;
          const d=assin.diasParaCobranca;
          const df=assin.diasParaFim;
          const riscoChurn = df<=30&&!cl.cancelado&&!cl.falhaRenovacao;
          if(d>15&&!riscoChurn)return null;
          return (
            <div style={{display:"flex",flexDirection:"column",gap:2,marginTop:3}}>
              {d<=15&&<div style={{ fontSize:10,fontWeight:500,color:d<=7?C.coralD:C.amberD,background:d<=7?C.coralL:C.amberL,padding:"1px 5px",borderRadius:4 }}>🔔 Renovacao em {d}d</div>}
              {riscoChurn&&<div style={{ fontSize:10,fontWeight:500,color:C.coralD,background:C.coralL,padding:"1px 5px",borderRadius:4 }}>⚠ Fidelidade encerra em {df}d</div>}
            </div>
          );
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
  if (erroCarregar) return (
    <div style={{textAlign:"center",padding:48,background:"var(--color-background-secondary)",borderRadius:12,margin:"20px 0"}}>
      <div style={{fontSize:24,marginBottom:12}}>⚠️</div>
      <div style={{fontSize:14,fontWeight:500,color:"var(--color-text-primary)",marginBottom:8}}>Erro ao carregar</div>
      <div style={{fontSize:12,color:C.coralD,marginBottom:16}}>{erroCarregar}</div>
      <button onClick={()=>{setErroCarregar("");carregar();}} style={{padding:"8px 20px",borderRadius:8,fontSize:13,fontWeight:500,cursor:"pointer",background:C.teal,color:"#fff",border:"none"}}>
        Tentar novamente
      </button>
    </div>
  );
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
      {(()=>{
        const hoje2 = new Date().toISOString().split("T")[0];
        const vencidos2 = clientes.filter(c=>c.dataProximoContato&&c.dataProximoContato<hoje2&&c.etapa!=="encerrado"&&c.etapa!=="convertido"&&c.etapa!=="experiencia");
        const renovacoes7 = clientes.filter(c=>{
          if(c.etapa!=="experiencia"||!c.tipoAssinatura||!c.dataInicioAssinatura) return false;
          const assin=calcAssinatura(c.tipoAssinatura,c.dataInicioAssinatura);
          return assin&&assin.diasParaCobranca<=7&&!c.cancelado&&!c.falhaRenovacao;
        });
        const focoSemAcao = clientes.filter(c=>(c.objetivo==="club"||c.objetivo==="falta_uma")&&!c.dataProximoContato&&c.etapa!=="encerrado"&&c.etapa!=="convertido"&&c.etapa!=="experiencia");
        const expSemDados = clientes.filter(c=>c.etapa==="experiencia"&&(!c.tipoAssinatura||!c.valorMensal));
        const itens = [
          vencidos2.length>0&&{emoji:"🔴",label:`${vencidos2.length} contato${vencidos2.length>1?"s":""} vencido${vencidos2.length>1?"s":""}`,cor:C.coralD,bg:C.coralL,
            onClick:()=>{ setBusca(""); setFiltroHoje(true); setFiltroClub(false); }},
          renovacoes7.length>0&&{emoji:"🔔",label:`${renovacoes7.length} renovaç${renovacoes7.length>1?"ões":"ão"} em ≤7 dias`,cor:C.amberD,bg:C.amberL,
            onClick:()=>{ setBusca(""); setFiltroClub(false); setFiltroHoje(false); }},
          focoSemAcao.length>0&&{emoji:"🎯",label:`${focoSemAcao.length} Foco Club sem próxima ação`,cor:C.purpleD,bg:C.purpleL,
            onClick:()=>{ setBusca(""); setFiltroClub(true); setFiltroHoje(false); }},
          expSemDados.length>0&&{emoji:"⭐",label:`${expSemDados.length} Experiência sem dados`,cor:C.blueD,bg:C.blueL,
            onClick:()=>{ setBusca(""); setFiltroClub(false); setFiltroHoje(false); }},
        ].filter(Boolean);
        if(itens.length===0) return null;
        return (
          <div style={{ display:"flex",gap:8,flexWrap:"wrap",marginBottom:12,padding:"10px 14px",background:"var(--color-background-secondary)",borderRadius:10,border:"0.5px solid var(--color-border-tertiary)" }}>
            <div style={{ fontSize:10,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",alignSelf:"center",marginRight:4 }}>Ações urgentes</div>
            {itens.map((it,i)=>(
              <button key={i} onClick={it.onClick||undefined}
                style={{ fontSize:11,fontWeight:500,background:it.bg,color:it.cor,padding:"3px 10px",borderRadius:20,border:"0.5px solid "+it.cor,cursor:it.onClick?"pointer":"default",fontFamily:"inherit" }}>
                {it.emoji} {it.label} {it.onClick?"→":""}
              </button>
            ))}
          </div>
        );
      })()}
      <div style={{ display:"flex",alignItems:"center",gap:10,marginBottom:14 }}>
        <div style={{ position:"relative",flex:1 }}>
          <span style={{ position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",fontSize:14,pointerEvents:"none" }}>🔍</span>
          <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="Nome, ID, telefone ou email..." style={{ width:"100%",padding:"8px 12px 8px 32px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none" }}/>
        </div>
        <div style={{ fontSize:13,color:"var(--color-text-tertiary)",whiteSpace:"nowrap" }}>
          {clientes.length} clientes
          {!filtroClub && (() => { const fc=clientes.filter(c=>c.objetivo==="club"||c.objetivo==="falta_uma").length; return fc>0?<span style={{ marginLeft:6,fontSize:11,background:C.greenL,color:C.greenD,padding:"1px 7px",borderRadius:20,fontWeight:500 }}>{fc} foco club</span>:null; })()}
          {(filtroProb||filtroPedidos)&&<span style={{ marginLeft:6,fontSize:11,background:C.amberL,color:C.amberD,padding:"1px 7px",borderRadius:20,fontWeight:500 }}>
            {[filtroProb&&"prob:"+filtroProb,filtroPedidos&&filtroPedidos+"ped"].filter(Boolean).join(" · ")}
          </span>}
        </div>
        <button onClick={()=>setFiltroClub(f=>!f)} style={{ padding:"5px 14px",borderRadius:8,fontSize:12,fontWeight:500,background:filtroClub?C.green:"var(--color-background-secondary)",border:"0.5px solid "+(filtroClub?C.green:"var(--color-border-tertiary)"),color:filtroClub?C.greenD:"var(--color-text-secondary)",cursor:"pointer",whiteSpace:"nowrap" }}>
          {filtroClub?"🎯 Foco Club ×":"🎯 Foco Club"}
        </button>
        <button onClick={()=>setFiltroHoje(f=>!f)} style={{ padding:"5px 14px",borderRadius:8,fontSize:12,fontWeight:500,background:filtroHoje?C.amber:"var(--color-background-secondary)",border:"0.5px solid "+(filtroHoje?C.amber:"var(--color-border-tertiary)"),color:filtroHoje?C.amberD:"var(--color-text-secondary)",cursor:"pointer",whiteSpace:"nowrap" }}>
          {filtroHoje?"⚡ Hoje ×":"⚡ Hoje"}
        </button>
        {/* Filtro probabilidade */}
        <div style={{ display:"flex",gap:4 }}>
          {[["","Prob."],["alta","Alta ≥40%"],["media","Média"],["baixa","Baixa"]].map(([v,l])=>(
            <button key={v} onClick={()=>setFiltroProb(filtroProb===v?"":v)}
              style={{ padding:"4px 8px",borderRadius:6,fontSize:11,fontWeight:500,cursor:"pointer",whiteSpace:"nowrap",
                background:filtroProb===v?C.teal:"var(--color-background-secondary)",
                color:filtroProb===v?"#fff":"var(--color-text-secondary)",
                border:"0.5px solid "+(filtroProb===v?C.teal:"var(--color-border-tertiary)") }}>
              {l}
            </button>
          ))}
        </div>
        {/* Filtro pedidos */}
        <div style={{ display:"flex",gap:4 }}>
          {[["","Ped."],["1","1 ped."],["2","2 ped."],["3+","3+ ped."]].map(([v,l])=>(
            <button key={v} onClick={()=>setFiltroPedidos(filtroPedidos===v?"":v)}
              style={{ padding:"4px 8px",borderRadius:6,fontSize:11,fontWeight:500,cursor:"pointer",whiteSpace:"nowrap",
                background:filtroPedidos===v?C.purple:"var(--color-background-secondary)",
                color:filtroPedidos===v?"#fff":"var(--color-text-secondary)",
                border:"0.5px solid "+(filtroPedidos===v?C.purple:"var(--color-border-tertiary)") }}>
              {l}
            </button>
          ))}
        </div>
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


// ── SHOPIFY ORDERS CSV PARSER ─────────────────────────────────────────────────
const parseTagsParaPedidos = (tags) => {
  // "Novo Cliente" → 1, "Pedido 2" → 2, "Pedido > 10" → 10+
  if (!tags) return null;
  if (/novo cliente/i.test(tags)) return 1;
  const m = tags.match(/pedido (\d+)/i);
  if (m) return parseInt(m[1]);
  if (/pedido > (\d+)/i.test(tags)) {
    const m2 = tags.match(/pedido > (\d+)/i);
    return m2 ? parseInt(m2[1]) + 1 : null;
  }
  return null;
};

const isOrdersCSV = (headers) => {
  const h = headers.map(x=>x.toLowerCase().trim());
  return h.includes("financial status") && h.includes("billing name") && h.includes("billing zip");
};

const parseOrdersCSVContent = (raw) => {
  const linhas = raw.split("\n").map(l=>l.trim()).filter(l=>l.length>0);
  if (linhas.length < 2) return { pedidos: [], erros: ["Arquivo vazio"] };

  const sep = linhas[0].includes(";") ? ";" : ",";
  
  // Parse CSV properly handling quoted fields
  const parseCSVLine = (line) => {
    const result = [];
    let cur = "";
    let inQuote = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"' && !inQuote) { inQuote = true; }
      else if (ch === '"' && inQuote) {
        if (line[i+1] === '"') { cur += '"'; i++; }
        else { inQuote = false; }
      } else if (ch === sep && !inQuote) { result.push(cur.trim()); cur = ""; }
      else { cur += ch; }
    }
    result.push(cur.trim());
    return result;
  };

  const headers = parseCSVLine(linhas[0]);
  const idx = (name) => headers.findIndex(h=>h.toLowerCase().trim()===name.toLowerCase().trim());

  const iNome    = idx("Billing Name");
  const iEmail   = idx("Email");
  const iTotal   = idx("Total");
  const iData    = idx("Created at");
  const iCEP     = idx("Billing Zip");
  const iTel     = idx("Billing Phone");
  const iStatus  = idx("Financial Status");
  const iTags    = idx("Tags");
  const iName    = idx("Name");

  if (iEmail < 0 || iNome < 0) {
    return { pedidos: [], erros: ["Colunas Email ou Billing Name nao encontradas"] };
  }

  // Group rows by order number (Name column), take first row per order
  const pedidosPorNum = {};
  for (let i = 1; i < linhas.length; i++) {
    const cols = parseCSVLine(linhas[i]);
    const orderNum = iName >= 0 ? (cols[iName]||"").trim() : "";
    const email = (cols[iEmail]||"").trim().toLowerCase();
    const status = iStatus >= 0 ? (cols[iStatus]||"").trim() : "";
    
    // Only process first row of each order (has billing info) and skip cancelled/refunded
    if (!email || !orderNum) continue;
    if (status && status !== "paid" && status !== "partially_paid") continue;
    if (pedidosPorNum[orderNum]) continue; // already processed this order
    
    const nome   = iNome >= 0 ? (cols[iNome]||"").trim() : "";
    const total  = iTotal >= 0 ? parseFloat((cols[iTotal]||"0").replace(",",".")) || 0 : 0;
    const data   = iData >= 0 ? parseShopifyDate(cols[iData]||"") : "";
    const cep    = iCEP >= 0 ? (cols[iCEP]||"").replace(/\D/g,"").padStart(8,"0") : "";
    const tel    = iTel >= 0 ? (cols[iTel]||"").trim() : "";
    const tags   = iTags >= 0 ? (cols[iTags]||"").trim() : "";
    const nPed   = parseTagsParaPedidos(tags);
    const fora   = cep ? !cep.startsWith("0") : null;

    if (nome && email) {
      pedidosPorNum[orderNum] = { email, nome, tel, total, data, cep, fora, nPed, tags };
    }
  }

  // Group by EMAIL to get per-customer stats
  const porEmail = {};
  Object.values(pedidosPorNum).forEach(p => {
    const e = p.email;
    if (!porEmail[e]) {
      porEmail[e] = { ...p, totalGasto: 0, datas: [], pedidos: [] };
    }
    porEmail[e].totalGasto += p.total;
    if (p.data) porEmail[e].datas.push(p.data);
    porEmail[e].pedidos.push(p);
    // Use highest nPed found for this customer
    if (p.nPed && (!porEmail[e].nPed || p.nPed > porEmail[e].nPed)) {
      porEmail[e].nPed = p.nPed;
    }
  });

  const pedidos = Object.values(porEmail).map(c => {
    const datasOrdenadas = [...c.datas].sort();
    return {
      email: c.email,
      nome: c.nome,
      tel: c.tel,
      fora: c.fora,
      cep: c.cep,
      totalGasto: Math.round(c.totalGasto * 100) / 100,
      nPedidos: c.nPed || c.pedidos.length,
      dataUltimo: datasOrdenadas[datasOrdenadas.length - 1] || "",
      dataPrimeiro: datasOrdenadas[0] || "",
      numerosOrdem: Object.keys(pedidosPorNum).filter(n => pedidosPorNum[n].email === c.email),
    };
  });

  return { pedidos, erros: [] };
};

const ImportarLista = ({ onSalvo }) => {
  const [txt,setTxt]=useState(""); const [prev,setPrev]=useState([]); const [imp,setImp]=useState(false);
  const [ok,setOk]=useState(null); const [erro,setErro]=useState("");
  const [modoOrders,setModoOrders]=useState(false);
  const [pedidosPreview,setPedidosPreview]=useState([]);
  const [confirmacaoOrders,setConfirmacaoOrders]=useState(null);
  const [ultimoImport,setUltimoImport]=useState(null);
  useEffect(()=>{ dbGetUltimoImport().then(setUltimoImport); },[]);
  const parse = (raw) => {
    const linhas = raw.split("\n").map(l=>l.trim()).filter(l=>l.length>0);
    const cls=[]; const errs=[];
    // Skip header row if first line looks like a header
    const start = /customer|nome|id/i.test(linhas[0]||"") ? 1 : 0;
    linhas.slice(start).forEach((linha,i)=>{
      const sep = linha.includes(";") ? ";" : ",";
      const pts = linha.split(sep).map(p=>p.trim());
      if(pts.length<4){errs.push("Linha "+(i+1+start)+": minimo 4 colunas");return;}
      const [customerId, nome, tel, gastoStr, pedStr, dp6, du7, cep8, lista9, email10] = pts;
      const lista = (lista9||"").trim();
      const email = (email10||"").trim().toLowerCase();
      const ped = parseInt(pedStr);
      const gasto = parseFloat((gastoStr||"0").replace(",","."))||0;
      const dp = parseShopifyDate(dp6||"");
      const du = parseShopifyDate(du7||"");
      const fora = cepToFora(cep8||"");
      if(!nome){errs.push("Linha "+(i+1+start)+": nome vazio");return;}
      if(isNaN(ped)||ped<1){errs.push("Linha "+(i+1+start)+": pedidos invalido");return;}
      cls.push({customerId:customerId||"",nome,tel:tel||"",ped,gasto,dp,du,fora,cep:cep8||"",lista,email});
    });
    return {cls,errs};
  };
  const atualizar = (val) => {
    setTxt(val); setOk(null); setErro(""); setConfirmacaoOrders(null);
    if(!val.trim()){setPrev([]);setPedidosPreview([]);setModoOrders(false);return;}
    const primeiraLinha = val.split("\n")[0]||"";
    const headers = primeiraLinha.split(",");
    if(isOrdersCSV(headers)){
      setModoOrders(true); setPrev([]);
      const r = parseOrdersCSVContent(val);
      setPedidosPreview(r.pedidos);
      if(r.erros.length>0) setErro(r.erros.join("\n"));
    } else {
      setModoOrders(false); setPedidosPreview([]);
      const {cls,errs}=parse(val); setPrev(cls);
      if(errs.length>0) setErro(errs.join("\n"));
    }
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

  const executarOrders = async () => {
    const inicio = Date.now();
    try {
        const total = pedidosPreview.length;
        setProg({ atual: 0, total, inicio });

        const existentes = await dbGetAll();
        const porEmail = {};
        existentes.forEach(c => { if (c.email||c.telefone) porEmail[(c.email||"").toLowerCase()] = c; });

        const novos = [], atualizados = [];

        let pedidosIgnorados = 0;
        pedidosPreview.forEach(p => {
          const existente = porEmail[p.email.toLowerCase()];
          if (existente) {
            // Deduplication: check which orders are new
            const jaImportados = new Set(existente.pedidosImportados||[]);
            const novosOrdens = (p.numerosOrdem||[]).filter(n=>!jaImportados.has(n));
            if (novosOrdens.length === 0) { pedidosIgnorados++; return; }

            const fracaoNova = p.numerosOrdem?.length > 0 ? novosOrdens.length/p.numerosOrdem.length : 1;
            const deltaGasto = p.totalGasto * fracaoNova;
            const todosPedidos = [...Array.from(jaImportados), ...novosOrdens];

            // Update: merge data, recalculate triagem
            const novoTotal = (existente.gasto||0) + deltaGasto;
            const novoP = Math.max(existente.p||0, p.nPedidos||0);
            const novaDataU = p.dataUltimo > (existente.dataUltimo||"") ? p.dataUltimo : existente.dataUltimo;
            const dp = existente.dataPrimeiro||p.dataPrimeiro||"";
            const temDados = !!(dp && existente.fora !== null && novoP >= 1);
            const tr = temDados ? runTriagem(novoP, dp, novaDataU||dp, existente.fora, novoTotal) : null;
            // Se estava como "nao_agora" e fez nova compra → auto-reativar para follow_up
            const reativarClub = existente.statusClub === "nao_agora" && novosOrdens.length > 0;
            atualizados.push({
              ...existente,
              gasto: novoTotal,
              p: novoP,
              dataUltimo: novaDataU,
              dataPrimeiro: dp,
              email: existente.email || p.email,
              pedidosImportados: todosPedidos,
              ...(reativarClub ? {
                statusClub: "follow_up",
                proximoFollowup: new Date().toISOString().split("T")[0],
                obsClub: (existente.obsClub||"")+(existente.obsClub?"\n":"")+"⚡ Fez novo pedido avulso em "+novaDataU+" — abordar agora!",
              } : {}),
              ...(tr ? {
                objetivo: tr.obj, objetivoLabel: tr.label,
                objetivoCor: tr.cor, objetivoCorD: tr.corD, objetivoAlerta: tr.alerta,
                prob: tr.prob.pct, probLabel: tr.prob.label, probCor: tr.prob.cor,
                seq: tr.seq, cicloMedio: tr.ciclo, datasPreenchidas: true,
              } : {})
            });
          } else {
            // New client
            const temDados = !!(p.dataPrimeiro && p.fora !== null && p.nPedidos >= 1);
            const tr = temDados ? runTriagem(p.nPedidos, p.dataPrimeiro, p.dataUltimo||p.dataPrimeiro, p.fora, p.totalGasto) : null;
            novos.push({
              id: "c_"+Date.now()+"_"+Math.random().toString(36).slice(2,8),
              etapa: "lead", dataCriacao: new Date().toLocaleDateString("pt-BR"),
              notas: "", proximaAcao: "", dataProximoContato: "",
              lista: "Pedidos Shopify", email: p.email,
              pedidosImportados: p.numerosOrdem||[],
              customerId: "", nome: p.nome, telefone: p.tel,
              p: p.nPedidos||1, gasto: p.totalGasto, fora: p.fora, cep: p.cep||"",
              dataPrimeiro: p.dataPrimeiro, dataUltimo: p.dataUltimo,
              datasPreenchidas: temDados,
              objetivo: tr?tr.obj:"", objetivoLabel: tr?tr.label:"⚠ Preencher datas",
              objetivoCor: tr?tr.cor:C.purple, objetivoCorD: tr?tr.corD:C.purpleD,
              objetivoAlerta: tr?tr.alerta:"",
              prob: tr?tr.prob.pct:0, probLabel: tr?tr.prob.label:"Pendente", probCor: tr?tr.prob.cor:C.purple,
              seq: tr?tr.seq:[], stepAtual:0, cicloMedio: tr?tr.ciclo:0,
            });
          }
        });

        setProg({ atual: Math.floor(total*0.1), total, inicio });

        const LOTE = 200;
        let salvos = 0;
        for (let i = 0; i < novos.length; i += LOTE) {
          await dbBulkSave(novos.slice(i, i+LOTE));
          salvos += Math.min(LOTE, novos.length-i);
          setProg({ atual: Math.floor(salvos/total*85)+5, total, inicio });
        }
        for (let i = 0; i < atualizados.length; i += LOTE) {
          await dbBulkSave(atualizados.slice(i, i+LOTE));
          salvos += Math.min(LOTE, atualizados.length-i);
          setProg({ atual: Math.floor(salvos/total*85)+5, total, inicio });
        }

        setProg({ atual: total, total, inicio });
        const msg = [
          novos.length > 0 ? novos.length+" novos leads" : "",
          atualizados.length > 0 ? atualizados.length+" perfis atualizados" : "",
          novos.filter(c=>c.datasPreenchidas).length > 0 ? novos.filter(c=>c.datasPreenchidas).length+" com triagem" : "",
          pedidosIgnorados > 0 ? pedidosIgnorados+" pedidos ja importados (ignorados)" : "",
        ].filter(Boolean).join(" · ");
        // Salvar registro do ultimo import
        const datasImport = pedidosPreview.map(p=>p.dataUltimo).filter(Boolean).sort();
        const ultimaData = datasImport[datasImport.length-1]||"";
        const ultimoPedidoNum = pedidosPreview.reduce((acc,p)=>{
          const nums = (p.numerosOrdem||[]).map(n=>parseInt(n.replace(/\D/g,""))).filter(Boolean);
          const max = nums.length>0?Math.max(...nums):0;
          return max>acc?max:acc;
        },0);
        await dbSaveUltimoImport({
          data: new Date().toLocaleDateString("pt-BR"),
          hora: new Date().toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"}),
          dataUltimoPedido: ultimaData,
          ultimoPedidoNum: ultimoPedidoNum>0?"#"+ultimoPedidoNum:"",
          novos: novos.length,
          atualizados: atualizados.length,
          ignorados: pedidosIgnorados,
          total: pedidosPreview.length,
        });
        setOk("✓ "+msg);
        setImp(false); setProg(null);
        setTimeout(()=>{ setTxt(""); setPrev([]); setPedidosPreview([]); setModoOrders(false);
          setConfirmacaoOrders(null); setOk(null); setErro(""); onSalvo&&onSalvo(); }, 2500);
      } catch(e) {
        setErro("Erro: "+(e.message||"tente novamente"));
        setImp(false); setProg(null);
      }
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
            let atualizado = {...existente};
            let mudou = false;
            if (cl.lista && cl.lista.trim()) {
              const listaAtual = existente.lista || "";
              const listas = listaAtual.split(" · ").map(l=>l.trim()).filter(Boolean);
              if (!listas.includes(cl.lista.trim())) {
                atualizado.lista = listaAtual ? listaAtual + " · " + cl.lista.trim() : cl.lista.trim();
                mudou = true;
              }
            }
            if (cl.email) {
              atualizado.email = cl.email;
              mudou = true;
            }
            if (mudou) paraAtualizarLista.push(atualizado);
          } else {
            // Novo cliente — criar com triagem
            const temDados = !!(cl.dp && cl.fora !== null && cl.ped >= 1);
            const tr = temDados ? runTriagem(cl.ped, cl.dp, cl.ped===1?cl.dp:(cl.du||cl.dp), cl.fora, cl.gasto||0) : null;
            novos.push({
              id:"c_"+Date.now()+"_"+Math.random().toString(36).slice(2,8),
              etapa:"lead", dataCriacao:new Date().toLocaleDateString("pt-BR"),
              notas:"", proximaAcao:"", dataProximoContato:"",
              lista:cl.lista, customerId:cl.customerId, nome:cl.nome, email:cl.email||"",
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
      {ultimoImport&&(
        <div style={{ background:C.tealL,border:"0.5px solid "+C.teal,borderRadius:10,padding:"12px 14px",marginBottom:12 }}>
          <div style={{ fontSize:12,fontWeight:500,color:C.tealD,marginBottom:8 }}>📦 Último import de pedidos</div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:8,marginBottom:8 }}>
            <div>
              <div style={{ fontSize:10,color:C.teal,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:2 }}>Realizado em</div>
              <div style={{ fontSize:12,fontWeight:500,color:C.tealD }}>{ultimoImport.data} {ultimoImport.hora}</div>
            </div>
            <div>
              <div style={{ fontSize:10,color:C.teal,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:2 }}>Último pedido</div>
              <div style={{ fontSize:12,fontWeight:500,color:C.tealD }}>
                {ultimoImport.dataUltimoPedido ? new Date(ultimoImport.dataUltimoPedido+"T12:00:00").toLocaleDateString("pt-BR") : "—"}
                {ultimoImport.ultimoPedidoNum&&<span style={{ fontSize:11,color:C.teal,marginLeft:6 }}>{ultimoImport.ultimoPedidoNum}</span>}
              </div>
            </div>
            <div>
              <div style={{ fontSize:10,color:C.teal,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:2 }}>Novos leads</div>
              <div style={{ fontSize:12,fontWeight:500,color:C.tealD }}>{ultimoImport.novos||0}</div>
            </div>
            <div>
              <div style={{ fontSize:10,color:C.teal,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:2 }}>Atualizados</div>
              <div style={{ fontSize:12,fontWeight:500,color:C.tealD }}>{ultimoImport.atualizados||0}</div>
            </div>
          </div>
          <div style={{ fontSize:11,color:C.tealD,paddingTop:8,borderTop:"0.5px solid "+C.teal }}>
            💡 Exporte do Shopify os pedidos a partir de <strong>{ultimoImport.dataUltimoPedido ? new Date(ultimoImport.dataUltimoPedido+"T12:00:00").toLocaleDateString("pt-BR") : "ontem"}</strong>{ultimoImport.ultimoPedidoNum?" (após pedido "+ultimoImport.ultimoPedidoNum+")":""} para não importar duplicatas.
          </div>
        </div>
      )}
      <div style={{ background:C.purpleL,border:"0.5px solid "+C.purple,borderRadius:8,padding:"12px 16px",marginBottom:16 }}>
        <div style={{ fontSize:13,fontWeight:500,color:C.purpleD,marginBottom:6 }}>Importação em lote de leads</div>
        <div style={{ fontSize:12,color:C.purpleD,lineHeight:1.6 }}>Cole a lista abaixo ou faça upload de um CSV. Todos entram na coluna Lead. O operador preenche as datas no perfil para gerar a sequência.</div>
      </div>
      <div style={{ background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 16px",marginBottom:16 }}>
        <div style={{ fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:8 }}>Formato aceito</div>
        <div style={{ fontSize:12,color:"var(--color-text-secondary)",lineHeight:1.8,fontFamily:"monospace" }}>
          Customer ID, Nome, Telefone, Total Gasto, Nº Pedidos, Data 1° Pedido, Data Último Pedido, CEP, Lista, Email<br/>
          <span style={{ color:C.tealD }}>1234, Maria Silva, 11 99999-1111, 380, 4, 2025-11-18, 2026-03-10, 04547-130, Pascoa Falta Uma, maria@email.com</span><br/>
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
      {modoOrders && pedidosPreview.length > 0 && !prog && (
        <div style={{ marginBottom:16 }}>
          <div style={{ background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",marginBottom:12 }}>
            <div style={{ fontSize:13,fontWeight:500,color:"var(--color-text-primary)",marginBottom:8 }}>
              📦 Formato detectado: Pedidos Shopify
            </div>
            <div style={{ fontSize:12,color:"var(--color-text-secondary)",marginBottom:10 }}>
              {pedidosPreview.length} clientes únicos encontrados no arquivo
            </div>
            {confirmacaoOrders && (
              <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12 }}>
                <div style={{ background:C.greenL,borderRadius:8,padding:"10px 12px",border:"0.5px solid "+C.green }}>
                  <div style={{ fontSize:10,color:C.greenD,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4 }}>Novos leads</div>
                  <div style={{ fontSize:22,fontWeight:500,color:C.greenD }}>{confirmacaoOrders.novos}</div>
                  <div style={{ fontSize:11,color:C.greenD }}>serao criados no CRM</div>
                </div>
                <div style={{ background:C.amberL,borderRadius:8,padding:"10px 12px",border:"0.5px solid "+C.amber }}>
                  <div style={{ fontSize:10,color:C.amberD,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4 }}>Perfis existentes</div>
                  <div style={{ fontSize:22,fontWeight:500,color:C.amberD }}>{confirmacaoOrders.atualizados}</div>
                  <div style={{ fontSize:11,color:C.amberD }}>terao dados atualizados</div>
                </div>
              </div>
            )}
            {!confirmacaoOrders && <div style={{ fontSize:12,color:"var(--color-text-tertiary)" }}>Verificando base existente...</div>}
          </div>
          {prog && (
            <div style={{ marginBottom:12,background:"var(--color-background-secondary)",borderRadius:10,padding:"14px 16px",border:"0.5px solid var(--color-border-tertiary)" }}>
              <div style={{ display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8 }}>
                <div style={{ fontSize:13,fontWeight:500,color:"var(--color-text-primary)" }}>Importando {prog.atual} de {prog.total}...</div>
                <div style={{ fontSize:12,color:"var(--color-text-tertiary)" }}>{Math.round(prog.atual/prog.total*100)}%</div>
              </div>
              <div style={{ height:8,background:"var(--color-border-tertiary)",borderRadius:4,overflow:"hidden",marginBottom:6 }}>
                <div style={{ height:"100%",width:(prog.atual/prog.total*100)+"%",background:C.teal,borderRadius:4,transition:"width 0.15s ease" }}/>
              </div>
              <div style={{ fontSize:11,color:"var(--color-text-tertiary)" }}>
                {prog.atual===prog.total?"Finalizando...":tempoRestante(prog)?"Tempo restante: "+tempoRestante(prog):"Calculando..."}
              </div>
            </div>
          )}
          {ok&&<div style={{ background:C.greenL,border:"0.5px solid "+C.green,borderRadius:8,padding:"10px 14px",marginBottom:10,fontSize:13,fontWeight:500,color:C.greenD }}>{ok}</div>}
          {!prog&&!ok&&<div style={{ display:"flex",gap:8 }}>
            <button onClick={()=>{ setModoOrders(false); setPedidosPreview([]); setConfirmacaoOrders(null); setTxt(""); setProg(null); setOk(null); }}
              style={{ flex:1,padding:"10px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-secondary)" }}>
              Cancelar
            </button>
            <button onClick={()=>{ if(!confirmacaoOrders){ dbGetAll().then(ex=>{ const em=new Set(ex.map(c=>(c.email||"").toLowerCase()).filter(Boolean)); setConfirmacaoOrders({novos:pedidosPreview.filter(p=>!em.has(p.email.toLowerCase())).length,atualizados:pedidosPreview.filter(p=>em.has(p.email.toLowerCase())).length}); }); } else { setImp(true); executarOrders(); } }}
              disabled={imp}
              style={{ flex:2,padding:"10px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",background:confirmacaoOrders?C.teal:C.purple,color:"#fff",border:"none" }}>
              {confirmacaoOrders?"Confirmar e importar →":"Verificar base →"}
            </button>
          </div>}
        </div>
      )}
      {!modoOrders && prev.length>0&&(
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
      <button onClick={importar} disabled={(modoOrders?pedidosPreview.length===0:prev.length===0)||imp||ok!==null} style={{ width:"100%",padding:"12px",borderRadius:10,fontSize:14,fontWeight:500,cursor:((modoOrders?pedidosPreview.length>0:prev.length>0)&&!imp)?"pointer":"default",border:"none",background:ok!==null?C.green:(modoOrders?pedidosPreview.length>0:prev.length>0)&&!imp?C.purple:"var(--color-background-secondary)",color:ok!==null||((modoOrders?pedidosPreview.length>0:prev.length>0)&&!imp)?"#fff":"var(--color-text-tertiary)" }}>
        {ok!==null?ok:imp?"Aguarde...":(modoOrders&&pedidosPreview.length>0)?"Use o botao acima para importar":prev.length>0?"Importar "+prev.length+" lead"+(prev.length!==1?"s":"")+" →":"Cole a lista acima para importar"}
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
      {/* RELATÓRIO DIÁRIO */}
      <div style={{ background:C.tealL, border:"0.5px solid "+C.teal, borderRadius:12, padding:"16px 20px", marginBottom:16 }}>
        <div style={{ fontSize:14, fontWeight:500, color:C.tealD, marginBottom:4 }}>📋 Relatório diário</div>
        <div style={{ fontSize:12, color:C.teal, marginBottom:14, lineHeight:1.5 }}>
          Gera um PDF com todas as atividades do dia, resumo por operador e situação atual do funil.
        </div>
        <button onClick={gerarRelatorioDiario}
          style={{ padding:"10px 20px", borderRadius:10, fontSize:13, fontWeight:500, cursor:"pointer", background:C.teal, color:"#fff", border:"none" }}>
          📄 Gerar relatório do dia → PDF
        </button>
      </div>

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
      <div style={{ minWidth:200,fontWeight:500,color:"var(--color-text-primary)",flexShrink:0 }}>{label}</div>
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
    <div style={{ maxWidth:720 }}>
      <div style={{ background:C.tealL,border:"0.5px solid "+C.teal,borderRadius:10,padding:"14px 16px",marginBottom:24 }}>
        <div style={{ fontSize:14,fontWeight:500,color:C.tealD,marginBottom:4 }}>Guia de uso — Laricas CRM</div>
        <div style={{ fontSize:12,color:C.tealD,lineHeight:1.6 }}>Documenta as premissas, logicas e regras do sistema. Mantenha atualizado conforme o operacional evoluir.</div>
      </div>

      <Section title="🗺 Etapas do funil">
        <Item label="Lead 🎯" value="Cliente importada ou cadastrada. Ainda nao foi contatada."/>
        <Item label="Primeiro Contato 📞" value="Mover apenas apos enviar a primeira mensagem. Nao antes."/>
        <Item label="Em Conversa 💬" value="Houve resposta e a conversa esta ativa."/>
        <Item label="Proposta Feita 📋" value="Proposta de Club apresentada. Aguardando decisao."/>
        <Item label="Convertido 🏆" value="Fechou Club ou compra avulsa. Encerrar o atendimento pelo botao no perfil."/>
        <Item label="Experiencia ⭐" value="Assinante ativa do Club. Preencher tipo, data de inicio e valor mensal no perfil."/>
        <Item label="Encerrado ✗" value="Nao converteu ou foi desqualificado. Registrar o motivo nas notas."/>
      </Section>

      <Section title="🎯 Objetivos de conversao">
        <Block title="Reativacao → 2a compra" cor={C.teal}>
          <Item label="Quando" value="1 pedido, ou 2 pedidos com ciclo > 60 dias"/>
          <Item label="Objetivo" value="Gerar a proxima compra. Nao oferecer Club ainda."/>
          <Item label="Abordagem" value="Tom caloroso. Curadoria com base no produto comprado. Sugerir o proximo sabor especifico."/>
          <Item label="Cupom" value="VOLTA10 — so no T3 (follow-up), sempre com prazo de 5 dias. Nunca sem prazo."/>
        </Block>
        <Block title="Falta Uma → 3a compra (momento aha)" cor={C.amber}>
          <Item label="Quando" value="2 pedidos com ciclo ≤ 60 dias"/>
          <Item label="Objetivo" value="Gerar a 3a compra. So oferecer Club apos ela acontecer."/>
          <Item label="T1" value="Ver historico no Shopify. Citar os 2 produtos ja comprados. Indicar UM produto especifico. Tom: calorosa e proxima, como indicacao de amiga."/>
          <Item label="T2" value="Link direto apos interesse confirmado. SEM cupom neste passo."/>
          <Item label="T3" value="VOLTA10 como facilitador. Prazo real de 5 dias."/>
          <Item label="T4" value="Encerramento caloroso. Porta aberta. Sem novo cupom."/>
          <Item label="Prioridade" value="Maxima — janela fecha se ciclo passar de 90 dias sem contato."/>
        </Block>
        <Block title="Club — habito formado" cor={C.green}>
          <Item label="Quando" value="3+ pedidos com ciclo ≤ 90 dias E último pedido há ≤ 45 dias"/>
          <Item label="Objetivo" value="Converter para assinatura Club"/>
          <Item label="3o pedido" value="Tom emocional: qual foi o favorito? Nao revelar intencao ainda."/>
          <Item label="4-6o pedido" value="Tom financeiro: calcular total gasto + frete vs preco do Club com numeros reais."/>
          <Item label="7o+ pedido" value="Tom surpresa: ainda nao tem o Club?"/>
          <Item label="Preco" value="Preco cheio no WhatsApp. Desconto de 20% so como fechamento em reuniao presencial — nunca antes."/>
          <Item label="Club pos 3a compra" value="Oferecer Club so apos a 3a compra (aha moment). Excecao: 7+ pedidos pode abordar em qualquer etapa."/>
        </Block>
        <Block title="Reconstruir habito → Club só depois" cor={C.coral}>
          <Item label="Quando — ciclo longo" value="3+ pedidos com ciclo > 90 dias"/>
          <Item label="Quando — inativa 45-90d" value="3+ pedidos (qualquer ciclo) mas último pedido há 45–90 dias"/>
          <Item label="Quando — inativa +90d" value="Qualquer histórico com último pedido há mais de 90 dias → reclassificada como Reativação"/>
          <Item label="Objetivo" value="Reativar a compra primeiro. Club so apos nova compra."/>
          <Item label="Atencao" value="⛔ NÃO oferecer Club. Muitas clientes com histórico de Club ativo estão aqui por estarem inativas há meses — precisam reconectar primeiro."/>
        </Block>
        <Block title="Novo cliente — 1a compra" cor={C.blue}>
          <Item label="Quando" value="Chegou por indicacao, redes sociais, evento ou presencialmente. Sem compras ainda."/>
          <Item label="Objetivo" value="Gerar a primeira compra. Cadastrar pela aba Triagem → Novo cliente sem compra."/>
          <Item label="Cupom" value="BEMVINDO10 — 10% na primeira compra. Valido 7 dias."/>
        </Block>
      </Section>

      <Section title="⭐ Retenção Club (Experiencia)">
        <Item label="R1 — Onboarding (mes 1)" value="Enviar ate 3 dias apos inicio. Perguntar sobre a experiencia e o favorito. Registrar nas notas."/>
        <Item label="R2 — Curadoria mensal" value="~5 dias antes da renovacao mensal. Perguntar se quer personalizar ou manter a selecao."/>
        <Item label="R3 — Renovacao (30d antes)" value="Abordar 30 dias antes do fim da fidelidade. Tom de cuidado, nao de cobranca. Nao mencionar preco primeiro."/>
        <Item label="R4 — Win-back" value="Ate 7 dias apos cancelamento. Sem oferta, sem desconto. Porta aberta."/>
        <Item label="Upsell" value="Assinante trimestral/semestral satisfeita no ciclo 2+ → candidata a migrar para plano maior. Script disponivel no perfil."/>
        <Item label="Churn risk" value="Badge vermelho aparece no card quando fidelidade encerra em ≤30 dias. Abordar com R3 antes de vencer."/>
      </Section>

      <Section title="📊 Probabilidade de conversao">
        <Item label="Alta (≥ 40%)" value="Falta Uma com ciclo curto, ou 3+ pedidos com ciclo regular"/>
        <Item label="Media (25–39%)" value="Club com ciclo longo, ou reativacao recente (< 30 dias)"/>
        <Item label="Baixa (< 25%)" value="Reativacao fria, ciclo longo, ou novo cliente"/>
        <Item label="Modificador +20%" value="Fora de SP — frete pesa mais na decisao pelo Club"/>
        <Item label="Modificador +35%" value="Gasto total alto — cliente de alto valor"/>
        <Item label="Modificador -28%" value="Fora da janela de 30 dias desde ultimo pedido"/>
        <Item label="Teto" value="Maxima de 72% — nenhum lead e certeza absoluta"/>
      </Section>

      <Section title="📋 Prioridade no Kanban">
        <Item label="1° Em Conversa / Proposta" value="Conversa ja iniciada. Risco de esfriar. Prioridade maxima."/>
        <Item label="2° Falta Uma" value="Janela curta. Cada dia sem contato reduz a chance."/>
        <Item label="3° Club habito" value="Candidata natural. Abordagem tranquila mas eficiente."/>
        <Item label="4° Reativacao" value="Precisa de mais trabalho antes do Club."/>
        <Item label="Filtro 🎯 Foco Club" value="Exibe so Club + Falta Uma — as candidatas proximas ao Club. Usar para foco diario."/>
        <Item label="Filtro ⚡ Hoje" value="Exibe so clientes com data de contato = hoje. Combinavel com Foco Club."/>
        <Item label="Painel urgencias" value="Barra no topo: contatos vencidos / renovacoes em 7d / Foco Club sem acao / Experiencia sem dados. Clicar ativa o filtro correspondente."/>
      </Section>

      <Section title="💬 Regras de abordagem">
        <Item label="Tom da Laricas" value="Caloroso, proximo e confiante. Como indicacao de amiga, nao abordagem comercial. 92% do publico e feminino."/>
        <Item label="Lucas fala em 1a pessoa" value="Humaniza a marca. Manter mesmo quando outros operadores enviarem."/>
        <Item label="Scripts personalizados" value="[Nome] e substituido automaticamente pelo primeiro nome. Copiar pelo botao 📋 — ja vem com o nome preenchido."/>
        <Item label="Placeholders em [colchetes]" value="Campos que o operador deve preencher antes de enviar. O sistema avisa quando ainda existem."/>
        <Item label="Cupom VOLTA10" value="So no T3 (follow-up). Prazo de 5 dias. Nunca prorrogar. Nunca oferecer antes."/>
        <Item label="Cupom BEMVINDO10" value="So para novos clientes (1a compra). Prazo de 7 dias."/>
        <Item label="Preco do Club" value="Preco cheio no WhatsApp. Desconto de 20% so em reuniao presencial — nunca revelar antes."/>
      </Section>

      <Section title="📥 Importacao">
        <Item label="Formato clientes" value="Customer ID, Nome, Telefone, Total Gasto, Nº Pedidos, Data 1° Pedido, Data Ultimo Pedido, CEP, Lista, Email"/>
        <Item label="Formato pedidos" value="Export direto do Shopify (orders). Sistema detecta automaticamente e mostra preview com novos leads + perfis a atualizar."/>
        <Item label="CEP" value="Iniciados em 0 (01xxx–09xxx) = SP. Demais = Fora de SP."/>
        <Item label="Datas" value="Aceita AAAA-MM-DD, DD/MM/AAAA ou serial do Excel."/>
        <Item label="Duplicata por Customer ID" value="Lista e acrescentada. Dados existentes nao sao alterados."/>
        <Item label="Duplicata por email (pedidos)" value="Pedidos novos atualizam gasto, data e recalculam triagem. Pedidos ja importados sao ignorados."/>
        <Item label="Ultimo import" value="Exibido no topo da aba com data, numero do ultimo pedido e dica de quando exportar no Shopify."/>
      </Section>

      <Section title="🔗 Unificacao de perfis">
        <Item label="Quando usar" value="Mesma cliente com dois cadastros no CRM (ex: importada + cadastrada manualmente)."/>
        <Item label="Deteccao automatica" value="Sistema sugere pares com nome, email ou telefone similares."/>
        <Item label="Merge" value="Escolher qual perfil manter. O outro e removido. Listas, notas, logs e historico sao combinados. Pedidos e gasto sao somados."/>
      </Section>

      <Section title="💾 Backup e versoes">
        <Item label="Dados" value="Exportar JSON antes de qualquer importacao grande (aba Backup → Exportar backup JSON)."/>
        <Item label="Codigo" value="Vercel guarda historico de todos os deploys. Em caso de problema: Deployments → 3 pontinhos → Promote to Production."/>
        <Item label="Restaurar dados" value="Aba Backup → Restaurar backup → selecionar o arquivo JSON."/>
      </Section>

      <Section title="🎯 Metas Club 2026">
        {[["Maio","20"],["Junho","36"],["Julho","36"],["Agosto","36"],["Setembro","71"],["Outubro","36"],["Novembro","71"],["Dezembro","20"]].map(([m,v])=>(
          <Item key={m} label={m+" 2026"} value={v+" novos assinantes Club"}/>
        ))}
      </Section>
    </div>
  );
};



const GraficoMRR = ({ mrrEvolucao, assinantes }) => {
  const [mesSel, setMesSel] = useState(null);
  const maxMrr = Math.max(...mrrEvolucao.map(m=>m.mrr), 1);
  const barW = Math.max(32, Math.min(60, Math.floor(560/mrrEvolucao.length)));
  const mesDetalhes = mesSel ? assinantes.filter(a=>{
    if (!a.dataInicioAssinatura||!a.valorMensal) return false;
    const inicio = new Date(a.dataInicioAssinatura+"T12:00:00");
    const [y,mo] = mesSel.split("-").map(Number);
    const dMes = new Date(y,mo-1,1);
    if (inicio > dMes) return false;
    if ((a.cancelado||a.falhaRenovacao)&&(a.dataCancelamento||a.dataFalhaRenovacao)) {
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
                    {m.novos>0&&<span style={{fontSize:7,color:C.greenD,background:C.greenL,padding:"1px 3px",borderRadius:3}}>+{m.novos}</span>}
                    {m.canc>0&&<span style={{fontSize:7,color:C.coralD,background:C.coralL,padding:"1px 3px",borderRadius:3}}>-{m.canc}</span>}
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
          <span style={{color:C.purple}}>■ selecionado — clique para detalhar</span>
        </div>
      </div>
      {mesSel&&(
        <div style={{marginTop:12,background:"var(--color-background-primary)",borderRadius:10,padding:"12px 14px",border:"0.5px solid "+C.purple}}>
          <div style={{fontSize:11,fontWeight:500,color:C.purpleD,marginBottom:8}}>
            Assinantes ativos em {mrrEvolucao.find(m=>m.mesKey===mesSel)?.mesLabel} ({mesDetalhes.length})
          </div>
          {mesDetalhes.length===0
            ?<div style={{fontSize:12,color:"var(--color-text-tertiary)"}}>Nenhum assinante com valor cadastrado neste mes.</div>
            :mesDetalhes.map(a=>(
              <div key={a.id} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 0",borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
                <div style={{flex:1,fontSize:12,fontWeight:500,color:"var(--color-text-primary)"}}>{a.nome}</div>
                <div style={{fontSize:11,color:"var(--color-text-secondary)",textTransform:"capitalize"}}>{a.tipoAssinatura||"—"}</div>
                <div style={{fontSize:12,fontWeight:500,color:C.tealD}}>R${a.valorMensal}/mes</div>
              </div>
            ))
          }
        </div>
      )}
    </div>
  );
};

const TabelaAssinantes = ({ assinantes, onAbrir }) => {
  const [sortCol, setSortCol] = useState("cicloAtual");
  const [sortDir, setSortDir] = useState("desc");
  const toggleSort = (col) => {
    if (sortCol===col) setSortDir(d=>d==="asc"?"desc":"asc");
    else { setSortCol(col); setSortDir("desc"); }
  };
  const Th = ({col,label}) => (
    <th onClick={()=>toggleSort(col)} style={{padding:"7px 10px",textAlign:col==="nome"?"left":"center",fontWeight:500,color:sortCol===col?C.teal:"var(--color-text-tertiary)",fontSize:11,borderBottom:"0.5px solid var(--color-border-tertiary)",textTransform:"uppercase",letterSpacing:"0.05em",cursor:"pointer",userSelect:"none",whiteSpace:"nowrap"}}>
      {label}{sortCol===col?(sortDir==="asc"?" ↑":" ↓"):""}
    </th>
  );
  const sorted = [...assinantes].sort((a,b)=>{
    let vA,vB;
    if(sortCol==="nome"){vA=a.nome||"";vB=b.nome||"";return sortDir==="asc"?vA.localeCompare(vB):vB.localeCompare(vA);}
    const assinA=calcAssinatura(a.tipoAssinatura,a.dataInicioAssinatura);
    const assinB=calcAssinatura(b.tipoAssinatura,b.dataInicioAssinatura);
    if(sortCol==="cicloAtual"){vA=assinA?assinA.cicloAtual:0;vB=assinB?assinB.cicloAtual:0;}
    else if(sortCol==="valorMensal"){vA=parseFloat(a.valorMensal)||0;vB=parseFloat(b.valorMensal)||0;}
    else if(sortCol==="ltvAtual"){
      const cA=a.cancelado?calcCiclosCancelado(a.dataInicioAssinatura,a.dataCancelamento):(assinA?assinA.cicloAtual:0);
      const cB=b.cancelado?calcCiclosCancelado(b.dataInicioAssinatura,b.dataCancelamento):(assinB?assinB.cicloAtual:0);
      vA=(parseFloat(a.valorMensal)||0)*cA+(a.gasto||0);vB=(parseFloat(b.valorMensal)||0)*cB+(b.gasto||0);
    }
    else if(sortCol==="diasCobranca"){vA=assinA?assinA.diasParaCobranca:9999;vB=assinB?assinB.diasParaCobranca:9999;}
    else{vA=0;vB=0;}
    return sortDir==="asc"?vA-vB:vB-vA;
  });
  return (
    <div style={{background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px"}}>
      <div style={{fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:12}}>Assinantes individuais ({assinantes.length})</div>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead>
            <tr style={{background:"var(--color-background-primary)"}}>
              <Th col="nome" label="Nome"/>
              <Th col="plano" label="Plano"/>
              <Th col="cicloAtual" label="Ciclo total"/>
              <Th col="valorMensal" label="R$/mes"/>
              <Th col="ltvAtual" label="LTV atual"/>
              <Th col="diasCobranca" label="Prox. cobr."/>
              <th style={{padding:"7px 10px",textAlign:"center",fontWeight:500,color:"var(--color-text-tertiary)",fontSize:11,borderBottom:"0.5px solid var(--color-border-tertiary)",textTransform:"uppercase",letterSpacing:"0.05em"}}>Status</th>
            </tr>
          </thead>
          <tbody>
            {sorted.map(a=>{
              const assin=calcAssinatura(a.tipoAssinatura,a.dataInicioAssinatura);
              const vm=parseFloat(a.valorMensal)||0;
              const ciclosPagos=a.cancelado?calcCiclosCancelado(a.dataInicioAssinatura,a.dataCancelamento):(assin?assin.cicloAtual:0);
              const ltvAtual=vm*ciclosPagos+(a.gasto||0);
              const statusLabel=a.cancelado?"Cancelado":a.falhaRenovacao?"Falha renovacao":"Ativa";
              const statusCor=a.cancelado?C.coralD:a.falhaRenovacao?C.amberD:C.greenD;
              const statusBg=a.cancelado?C.coralL:a.falhaRenovacao?C.amberL:C.greenL;
              return (
                <tr key={a.id} style={{borderBottom:"0.5px solid var(--color-border-tertiary)",opacity:(a.cancelado||a.falhaRenovacao)?0.7:1}}>
                  <td style={{padding:"7px 10px"}}>
                    <button onClick={()=>onAbrir&&onAbrir(a.id)} style={{background:"none",border:"none",cursor:onAbrir?"pointer":"default",fontWeight:500,color:onAbrir?C.teal:"var(--color-text-primary)",fontSize:12,padding:0,textAlign:"left"}}>{a.nome}</button>
                  </td>
                  <td style={{padding:"7px 10px",textAlign:"center",color:"var(--color-text-secondary)",textTransform:"capitalize"}}>{a.tipoAssinatura||"—"}</td>
                  <td style={{padding:"7px 10px",textAlign:"center",color:C.purpleD,fontWeight:500}}>
                    {assin?<span>{assin.cicloAtual}°<span style={{fontSize:10,color:C.purple,fontWeight:400}}> ({assin.cicloNoPeriodo}/{assin.ciclosTotais})</span></span>:"—"}
                  </td>
                  <td style={{padding:"7px 10px",textAlign:"center",color:"var(--color-text-secondary)"}}>{vm>0?"R$"+vm.toFixed(0):"—"}</td>
                  <td style={{padding:"7px 10px",textAlign:"center",fontWeight:500,color:C.greenD}}>{vm>0?"R$"+ltvAtual.toFixed(0):"—"}</td>
                  <td style={{padding:"7px 10px",textAlign:"center",color:assin&&assin.diasParaCobranca<=7?C.coralD:assin&&assin.diasParaCobranca<=15?C.amberD:"var(--color-text-secondary)"}}>{assin&&!a.cancelado?assin.proximaCobranca:"—"}</td>
                  <td style={{padding:"7px 10px",textAlign:"center"}}><span style={{fontSize:10,fontWeight:500,background:statusBg,color:statusCor,padding:"2px 8px",borderRadius:20}}>{statusLabel}</span></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
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
  const ltvRealizadoAtivos = comValor.reduce((acc,a)=>{
    const assin=calcAssinatura(a.tipoAssinatura,a.dataInicioAssinatura);
    return acc+(a.valorMensal||0)*(assin?assin.cicloAtual:0)+(a.gasto||0);
  },0);
  const ltvCancelados = cancelados.filter(a=>a.valorMensal>0).reduce((acc,a)=>{
    const ciclos=calcCiclosCancelado(a.dataInicioAssinatura,a.dataCancelamento);
    return acc+(a.valorMensal||0)*ciclos+(a.gasto||0);
  },0);
  const ltvPagoTotal = ltvRealizadoAtivos + ltvCancelados;
  const ltvProjetadoTotal = comValor.reduce((acc,a)=>{
    const assin=calcAssinatura(a.tipoAssinatura,a.dataInicioAssinatura);
    const vm=a.valorMensal||0;
    const ciclosRestantes=assin?assin.ciclosTotais-assin.cicloNoPeriodo:0;
    return acc+vm*(assin?assin.cicloAtual:0)+vm*ciclosRestantes+(a.gasto||0);
  },0);
  const mrr = comValor.reduce((acc,a)=>acc+(a.valorMensal||0),0);
  const semValor = ativos.length - comValor.length;

  // MRR variação mês anterior
  const hoje3=new Date();
  const mesAnterior=new Date(hoje3.getFullYear(),hoje3.getMonth()-1,1);
  const mrrAnterior=ativos.reduce((acc,a)=>{
    if(!a.dataInicioAssinatura||!a.valorMensal) return acc;
    const ini=new Date(a.dataInicioAssinatura+"T12:00:00");
    if(ini>mesAnterior) return acc;
    return acc+(parseFloat(a.valorMensal)||0);
  },0);
  const varMRR=mrr-mrrAnterior;
  const varPct=mrrAnterior>0?Math.round(varMRR/mrrAnterior*100):0;

  // Churn
  const churnPorMes={};
  cancelados.forEach(a=>{
    if(!a.dataCancelamento) return;
    const key=a.dataCancelamento.substring(0,7);
    churnPorMes[key]=(churnPorMes[key]||0)+1;
  });
  const mesesComChurn=Object.keys(churnPorMes).sort().reverse().slice(0,3);
  const churnMesAtual=churnPorMes[new Date().toISOString().substring(0,7)]||0;
  const tempoMedioMeses=cancelados.length>0
    ?Math.round(cancelados.filter(a=>a.dataInicioAssinatura&&a.dataCancelamento).reduce((acc,a)=>acc+calcCiclosCancelado(a.dataInicioAssinatura,a.dataCancelamento),0)/cancelados.filter(a=>a.dataInicioAssinatura&&a.dataCancelamento).length)
    :0;
  const churnRate=ativos.length>0?Math.round(cancelados.length/(ativos.length+cancelados.length)*100):0;

  // MRR evolucao
  const mrrEvolucao=(()=>{
    if(assinantes.length===0) return [];
    const datas=assinantes.filter(a=>a.dataInicioAssinatura).map(a=>new Date(a.dataInicioAssinatura+"T12:00:00"));
    if(datas.length===0) return [];
    const minData=new Date(Math.min(...datas.map(d=>d.getTime())));
    const hoje2=new Date();
    const meses=[];
    let cursor=new Date(minData.getFullYear(),minData.getMonth(),1);
    while(cursor<=hoje2&&meses.length<24){
      const mesKey=cursor.toISOString().substring(0,7);
      const mesLabel=cursor.toLocaleDateString("pt-BR",{month:"short",year:"2-digit"});
      const mrr_mes=assinantes.filter(a=>{
        if(!a.dataInicioAssinatura||!a.valorMensal) return false;
        const inicio=new Date(a.dataInicioAssinatura+"T12:00:00");
        if(inicio>cursor) return false;
        if((a.cancelado||a.falhaRenovacao)&&(a.dataCancelamento||a.dataFalhaRenovacao)){
          const cancel=new Date((a.dataCancelamento||a.dataFalhaRenovacao)+"T12:00:00");
          const fimMes=new Date(cursor.getFullYear(),cursor.getMonth()+1,0);
          if(cancel<cursor) return false;
        }
        return true;
      }).reduce((acc,a)=>acc+(parseFloat(a.valorMensal)||0),0);
      const novos=assinantes.filter(a=>{
        if(!a.dataInicioAssinatura) return false;
        return new Date(a.dataInicioAssinatura+"T12:00:00").toISOString().substring(0,7)===mesKey;
      }).length;
      const canc=assinantes.filter(a=>a.cancelado&&a.dataCancelamento&&a.dataCancelamento.substring(0,7)===mesKey).length;
      meses.push({mesKey,mesLabel,mrr:Math.round(mrr_mes),novos,canc});
      cursor=new Date(cursor.getFullYear(),cursor.getMonth()+1,1);
    }
    return meses;
  })();

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
          {(()=>{
            return (
              <div style={{background:"var(--color-background-primary)",borderRadius:10,padding:"12px 14px",borderLeft:"3px solid "+C.teal}}>
                <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>MRR — ativos</div>
                <div style={{fontSize:24,fontWeight:500,color:C.tealD}}>R${mrr.toLocaleString("pt-BR",{minimumFractionDigits:0})}</div>
                <div style={{display:"flex",alignItems:"center",gap:6,marginTop:2}}>
                  <span style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{comValor.length} · média R${comValor.length>0?(mrr/comValor.length).toFixed(0):0}/mês</span>
                  {mrrAnterior>0&&<span style={{fontSize:10,fontWeight:500,color:varMRR>=0?C.greenD:C.coralD,background:varMRR>=0?C.greenL:C.coralL,padding:"1px 6px",borderRadius:10}}>{varMRR>=0?"+":""}R${varMRR.toFixed(0)} ({varPct>=0?"+":""}{varPct}%)</span>}
                </div>
              </div>
            );
          })()}
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

      {cancelados.length>0&&(
        <div style={{background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px",marginBottom:16}}>
          <div style={{fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:12}}>Cancelamentos e churn</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
            <div style={{background:"var(--color-background-primary)",borderRadius:10,padding:"12px 14px",borderLeft:"3px solid "+C.coral}}>
              <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>Churn total</div>
              <div style={{fontSize:24,fontWeight:500,color:C.coralD}}>{churnRate}%</div>
              <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:2}}>{cancelados.length} de {ativos.length+cancelados.length}</div>
            </div>
            <div style={{background:"var(--color-background-primary)",borderRadius:10,padding:"12px 14px",borderLeft:"3px solid "+C.amber}}>
              <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>Tempo medio ate cancelar</div>
              <div style={{fontSize:24,fontWeight:500,color:C.amberD}}>{tempoMedioMeses||"—"}{tempoMedioMeses?"m":""}</div>
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

      {mrrEvolucao.length>1&&(
        <div style={{background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px",marginBottom:16}}>
          <div style={{fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:14}}>Evolucao do MRR</div>
          <GraficoMRR mrrEvolucao={mrrEvolucao} assinantes={assinantes}/>
        </div>
      )}

      {(()=>{
        const grupos=comValor.length>0?["trimestral","semestral","anual"].map(tipo=>{
          const g=comValor.filter(a=>a.tipoAssinatura===tipo);
          if(g.length===0) return null;
          const mediaVM=Math.round(g.reduce((acc,a)=>acc+(a.valorMensal||0),0)/g.length);
          const mediaLTV=Math.round(g.reduce((acc,a)=>{
            const assin=calcAssinatura(a.tipoAssinatura,a.dataInicioAssinatura);
            return acc+(a.valorMensal||0)*(assin?assin.ciclosTotais:0)+(a.gasto||0);
          },0)/g.length);
          return {tipo,count:g.length,mediaVM,mediaLTV};
        }).filter(Boolean):[];
        if(grupos.length===0) return null;
        return (
          <div style={{display:"flex",gap:8,marginBottom:16}}>
            {grupos.map(g=>(
              <div key={g.tipo} style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 14px",flex:1,textAlign:"center"}}>
                <div style={{fontSize:10,color:"var(--color-text-tertiary)",textTransform:"capitalize",marginBottom:4}}>{g.tipo} ({g.count})</div>
                <div style={{fontSize:14,fontWeight:500,color:"var(--color-text-primary)"}}>R${g.mediaVM}/mes</div>
                <div style={{fontSize:11,color:C.purpleD}}>LTV ~R${g.mediaLTV}</div>
              </div>
            ))}
          </div>
        );
      })()}

      <TabelaAssinantes assinantes={assinantes} onAbrir={onAbrir}/>

      {(()=>{
        const indicadoras={};
        assinantes.forEach(a=>{
          if(a.indicadaPor&&a.indicadaPor.trim()){
            const k=a.indicadaPor.trim();
            if(!indicadoras[k]) indicadoras[k]={nome:k,count:0};
            indicadoras[k].count++;
          }
        });
        const ranking=Object.values(indicadoras).sort((a,b)=>b.count-a.count).slice(0,10);
        if(ranking.length===0) return null;
        return (
          <div style={{background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px",marginTop:16}}>
            <div style={{fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:12}}>🏆 Ranking de indicações</div>
            <div style={{display:"flex",flexDirection:"column",gap:6}}>
              {ranking.map((r,i)=>(
                <div key={r.nome} style={{display:"flex",alignItems:"center",gap:10,padding:"7px 10px",background:"var(--color-background-primary)",borderRadius:8}}>
                  <span style={{fontSize:12,fontWeight:500,color:C.amberD,minWidth:20}}>#{i+1}</span>
                  <span style={{fontSize:13,fontWeight:500,color:"var(--color-text-primary)",flex:1}}>{r.nome}</span>
                  <span style={{fontSize:11,background:C.greenL,color:C.greenD,padding:"1px 8px",borderRadius:20,fontWeight:500}}>{r.count} indicaç{r.count>1?"ões":"ão"}</span>
                </div>
              ))}
            </div>
          </div>
        );
      })()}
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
        (c.telefone||"").toLowerCase().includes(ql) ||
        (c.email||"").toLowerCase().includes(ql) ||
        (c.emailClub||"").toLowerCase().includes(ql) ||
        (c.responsavel||"").toLowerCase().includes(ql)
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
          placeholder="Buscar por nome, ID, telefone ou email..."
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
                  <div style={{ fontSize:11,color:"var(--color-text-tertiary)" }}>{c.customerId?"#"+c.customerId+" · ":""}{c.emailClub||c.email||c.telefone||""}</div>
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


const Unificar = ({ onSalvo }) => {
  const [clientes, setClientes] = useState([]);
  const [loading, setLoading] = useState(true);
  const [sugestoes, setSugestoes] = useState([]);
  const [busca, setBusca] = useState("");
  const [resultados, setResultados] = useState([]);
  const [selecionados, setSelecionados] = useState([]); // [idA, idB]
  const [preview, setPreview] = useState(null); // {a, b}
  const [salvando, setSalvando] = useState(false);
  const [ok, setOk] = useState("");

  useEffect(() => {
    dbGetAll().then(lista => {
      setClientes(lista);
      // Detectar duplicatas: mesmo nome (normalizado) ou mesmo email ou mesmo telefone
      const grupos = {};
      lista.forEach(c => {
        const nomeNorm = (c.nome||"").toLowerCase().trim().replace(/\s+/g," ");
        const email = (c.email||"").toLowerCase().trim();
        const tel = (c.telefone||"").replace(/\D/g,"").slice(-8);
        [nomeNorm, email&&"e:"+email, tel&&tel.length>=8&&"t:"+tel].filter(Boolean).forEach(k => {
          if(!grupos[k]) grupos[k]=[];
          grupos[k].push(c);
        });
      });
      const dupl = Object.values(grupos)
        .filter(g=>g.length>=2)
        .map(g=>({ a:g[0], b:g[1], score: calcScore(g[0],g[1]) }))
        .filter(g=>g.score>0)
        .sort((a,b)=>b.score-a.score)
        .slice(0,20);
      // Deduplicate sugestoes
      const seen = new Set();
      const uniq = dupl.filter(d=>{
        const k=[d.a.id,d.b.id].sort().join("|");
        if(seen.has(k)) return false;
        seen.add(k); return true;
      });
      setSugestoes(uniq);
      setLoading(false);
    });
  }, []);

  const calcScore = (a, b) => {
    let s = 0;
    const nA=(a.nome||"").toLowerCase().trim();
    const nB=(b.nome||"").toLowerCase().trim();
    if(nA&&nB&&nA===nB) s+=3;
    else if(nA&&nB&&(nA.includes(nB)||nB.includes(nA))) s+=2;
    const eA=(a.email||"").toLowerCase().trim();
    const eB=(b.email||"").toLowerCase().trim();
    if(eA&&eB&&eA===eB) s+=3;
    const tA=(a.telefone||"").replace(/\D/g,"").slice(-8);
    const tB=(b.telefone||"").replace(/\D/g,"").slice(-8);
    if(tA&&tB&&tA.length>=8&&tA===tB) s+=3;
    return s;
  };

  const buscar = () => {
    if(!busca.trim()) return;
    const q = busca.toLowerCase();
    const res = clientes.filter(c=>
      (c.nome||"").toLowerCase().includes(q)||
      (c.email||"").toLowerCase().includes(q)||
      (c.telefone||"").includes(q)||
      (c.customerId||"").includes(q)
    ).slice(0,10);
    setResultados(res);
    setSelecionados([]);
    setPreview(null);
  };

  const toggleSel = (c) => {
    setSelecionados(prev => {
      if(prev.find(x=>x.id===c.id)) return prev.filter(x=>x.id!==c.id);
      if(prev.length>=2) return [prev[1],c];
      return [...prev,c];
    });
  };

  useEffect(()=>{
    if(selecionados.length===2) setPreview({a:selecionados[0],b:selecionados[1]});
    else setPreview(null);
  },[selecionados]);

  const unificar = async (manter, remover) => {
    setSalvando(true);
    // Merge: keep manter, delete remover, combine listas and notas
    const listaA = manter.lista||"";
    const listaB = remover.lista||"";
    const listasMerge = [...new Set([...listaA.split(" · "),...listaB.split(" · ")].map(l=>l.trim()).filter(Boolean))].join(" · ");
    const notasMerge = [manter.notas, remover.notas].filter(Boolean).join("\n---\n");
    const logMerge = [...(manter.logAtividade||[]),...(remover.logAtividade||[])].sort((a,b)=>a.data>b.data?-1:1).slice(0,30);
    const histMerge = [...(manter.historicoEtapas||[]),...(remover.historicoEtapas||[])];
    const merged = {
      ...manter,
      lista: listasMerge,
      notas: notasMerge,
      logAtividade: logMerge,
      historicoEtapas: histMerge,
      email: manter.email||remover.email,
      emailClub: manter.emailClub||remover.emailClub,
      telefone: manter.telefone||remover.telefone,
      customerId: manter.customerId||remover.customerId,
      gasto: (manter.gasto||0) + (remover.gasto||0),
      p: (manter.p||0) + (remover.p||0),
      dataPrimeiro: [manter.dataPrimeiro,remover.dataPrimeiro].filter(Boolean).sort()[0]||"",
      dataUltimo: [manter.dataUltimo,remover.dataUltimo].filter(Boolean).sort().reverse()[0]||"",
    };
    try {
      await dbSave(merged);
      await dbDelete(remover.id);
      setOk("✓ Perfis unificados com sucesso!");
      setPreview(null); setSelecionados([]); setResultados([]);
      setBusca(""); setSugestoes(prev=>prev.filter(s=>s.a.id!==remover.id&&s.b.id!==remover.id));
      setClientes(prev=>prev.filter(c=>c.id!==remover.id).map(c=>c.id===manter.id?merged:c));
      setTimeout(()=>{setOk(""); onSalvo&&onSalvo();},2000);
    } catch(e) { setOk("Erro: "+e.message); }
    setSalvando(false);
  };

  const CardCliente = ({c, selecionado, onClick}) => {
    const etapa = ETAPAS.find(e=>e.id===c.etapa)||ETAPAS[0];
    return (
      <button onClick={onClick} style={{ width:"100%",textAlign:"left",padding:"10px 12px",borderRadius:10,border:"2px solid "+(selecionado?C.teal:"var(--color-border-tertiary)"),background:selecionado?C.tealL:"var(--color-background-secondary)",cursor:"pointer",marginBottom:6,transition:"all 0.15s" }}>
        <div style={{ display:"flex",alignItems:"center",gap:8,marginBottom:4 }}>
          <span style={{ fontSize:11,fontWeight:500,color:etapa.corD,background:etapa.corL,padding:"1px 6px",borderRadius:10 }}>{etapa.emoji} {etapa.label}</span>
          {selecionado&&<span style={{ fontSize:10,fontWeight:500,color:C.tealD }}>✓ Selecionado</span>}
        </div>
        <div style={{ fontSize:13,fontWeight:500,color:"var(--color-text-primary)",marginBottom:2 }}>{c.nome}</div>
        <div style={{ fontSize:11,color:"var(--color-text-tertiary)" }}>
          {c.email&&<span>{c.email} · </span>}{c.telefone&&<span>{c.telefone} · </span>}
          {c.p||0} pedidos · R${(c.gasto||0).toFixed(0)}
        </div>
        {c.notas&&<div style={{ fontSize:11,color:"var(--color-text-secondary)",marginTop:4,fontStyle:"italic" }}>"{c.notas.slice(0,60)}{c.notas.length>60?"...":""}"</div>}
      </button>
    );
  };

  if(loading) return <div style={{textAlign:"center",padding:40,color:"var(--color-text-tertiary)"}}>Analisando base...</div>;

  return (
    <div>
      {ok&&<div style={{ background:C.greenL,border:"0.5px solid "+C.green,borderRadius:8,padding:"10px 14px",marginBottom:12,fontSize:13,fontWeight:500,color:C.greenD }}>{ok}</div>}

      {/* Sugestões automáticas */}
      {sugestoes.length>0&&(
        <div style={{ background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px",marginBottom:16 }}>
          <div style={{ fontSize:13,fontWeight:500,color:"var(--color-text-primary)",marginBottom:4 }}>🔍 Possíveis duplicatas detectadas</div>
          <div style={{ fontSize:12,color:"var(--color-text-secondary)",marginBottom:12 }}>{sugestoes.length} pares com nome, email ou telefone similares</div>
          {sugestoes.map((s,i)=>(
            <div key={i} style={{ background:"var(--color-background-primary)",borderRadius:10,padding:"12px 14px",marginBottom:10,border:"0.5px solid var(--color-border-tertiary)" }}>
              <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:8 }}>
                <CardCliente c={s.a} selecionado={!!selecionados.find(x=>x.id===s.a.id)} onClick={()=>{setSelecionados([s.a,s.b]);}}/>
                <CardCliente c={s.b} selecionado={!!selecionados.find(x=>x.id===s.b.id)} onClick={()=>{setSelecionados([s.a,s.b]);}}/>
              </div>
              <button onClick={()=>setSelecionados([s.a,s.b])} style={{ width:"100%",padding:"7px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",background:C.teal,color:"#fff",border:"none" }}>
                Analisar e unificar →
              </button>
            </div>
          ))}
        </div>
      )}

      {/* Busca manual */}
      <div style={{ background:"var(--color-background-secondary)",borderRadius:12,padding:"14px 16px",marginBottom:16 }}>
        <div style={{ fontSize:13,fontWeight:500,color:"var(--color-text-primary)",marginBottom:10 }}>Busca manual</div>
        <div style={{ display:"flex",gap:8,marginBottom:10 }}>
          <input value={busca} onChange={e=>setBusca(e.target.value)} onKeyDown={e=>e.key==="Enter"&&buscar()}
            placeholder="Nome, email ou telefone..."
            style={{ flex:1,padding:"8px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-primary)",outline:"none" }}/>
          <button onClick={buscar} style={{ padding:"8px 16px",borderRadius:8,fontSize:13,fontWeight:500,cursor:"pointer",background:C.teal,color:"#fff",border:"none" }}>Buscar</button>
        </div>
        {resultados.length>0&&(
          <div>
            <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:8 }}>Selecione 2 perfis para unificar:</div>
            {resultados.map(c=><CardCliente key={c.id} c={c} selecionado={!!selecionados.find(x=>x.id===c.id)} onClick={()=>toggleSel(c)}/>)}
          </div>
        )}
      </div>

      {/* Preview unificação */}
      {preview&&(
        <div style={{ background:"var(--color-background-primary)",border:"0.5px solid "+C.teal,borderRadius:12,padding:"16px" }}>
          <div style={{ fontSize:13,fontWeight:500,color:C.tealD,marginBottom:12 }}>Escolha qual perfil manter</div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12 }}>
            {[preview.a, preview.b].map((c,i)=>(
              <div key={c.id} style={{ background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px" }}>
                <div style={{ fontSize:12,fontWeight:500,color:"var(--color-text-primary)",marginBottom:8 }}>{c.nome}</div>
                <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:6,lineHeight:1.6 }}>
                  {c.email&&<div>📧 {c.email}</div>}
                  {c.telefone&&<div>📱 {c.telefone}</div>}
                  <div>🛒 {c.p||0} pedidos · R${(c.gasto||0).toFixed(0)}</div>
                  <div>📋 {ETAPAS.find(e=>e.id===c.etapa)?.label||c.etapa}</div>
                  {c.notas&&<div>📝 Tem anotações</div>}
                  {(c.logAtividade||[]).length>0&&<div>📌 {c.logAtividade.length} logs</div>}
                </div>
                <button onClick={()=>unificar(c, i===0?preview.b:preview.a)} disabled={salvando}
                  style={{ width:"100%",padding:"8px",borderRadius:8,fontSize:12,fontWeight:500,cursor:salvando?"default":"pointer",background:C.teal,color:"#fff",border:"none",opacity:salvando?0.6:1 }}>
                  ✓ Manter este
                </button>
              </div>
            ))}
          </div>
          <div style={{ fontSize:11,color:"var(--color-text-tertiary)",textAlign:"center" }}>
            O perfil escolhido absorve listas, notas, logs e histórico do outro. O outro é removido.
          </div>
          {preview&&(()=>{
            // preview é {a, b} — desestruturação como objeto
            const pa = preview.a, pb = preview.b;
            const campos = ["nome","telefone","email","emailClub","responsavel"];
            const conflitos = campos.filter(campo=>pa[campo]&&pb[campo]&&pa[campo]!==pb[campo]);
            return (
              <div style={{marginTop:12,background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px"}}>
                <div style={{fontSize:12,color:C.tealD,marginBottom:6,fontWeight:500}}>
                  📊 Pedidos somados: {(pa.p||0)+(pb.p||0)} &nbsp;·&nbsp;
                  Gasto somado: R${((pa.gasto||0)+(pb.gasto||0)).toFixed(0)}
                </div>
                {conflitos.length>0&&(
                  <div style={{fontSize:11,color:C.coralD,marginTop:6}}>
                    <strong>⚠ Campos com valores diferentes — ao clicar "Manter este", o da esquerda prevalece:</strong>
                    {conflitos.map(campo=>(
                      <div key={campo} style={{marginTop:4,background:C.coralL,borderRadius:6,padding:"4px 8px"}}>
                        <strong>{campo}:</strong> "{pa[campo]}" ← vs → "{pb[campo]}"
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })()}
        </div>
      )}
    </div>
  );
};


const gerarRelatorioDiario = async () => {
  const todos = await dbGetAll();
  const hoje = new Date();
  const hojeStr = hoje.toLocaleDateString("pt-BR");
  const hojeISO = hoje.toISOString().split("T")[0];

  // ── Coletar atividades do dia ──────────────────────────────────────────────
  const atividades = []; // {hora, operador, cliente, tipo, detalhe}

  todos.forEach(c => {
    // Logs de atividade do dia
    (c.logAtividade||[]).forEach(l => {
      if (l.data === hojeStr) {
        atividades.push({ hora:l.hora||"--:--", operador:l.resp||"—", cliente:c.nome, tipo:"Contato", detalhe:l.texto, proximaAcao:c.proximaAcao||"", dataProximoContato:c.dataProximoContato||"" });
      }
    });
    // Movimentações de etapa do dia
    (c.historicoEtapas||[]).forEach(h => {
      const dataH = h.data||"";
      // data pode ser "DD/MM/AAAA" ou ISO
      const eHoje = dataH === hojeStr || dataH.startsWith(hojeISO);
      if (eHoje) {
        const etLabel = ETAPAS.find(e=>e.id===h.etapa)?.label || h.etapa;
        atividades.push({ hora:h.hora||"--:--", operador:h.resp||"—", cliente:c.nome, tipo:"Etapa", detalhe:"Movido para "+etLabel, proximaAcao:c.proximaAcao||"", dataProximoContato:c.dataProximoContato||"" });
      }
    });
    // Novos leads criados hoje
    if ((c.dataCriacao||"") === hojeStr) {
      atividades.push({ hora:"--:--", operador:c.responsavel||"—", cliente:c.nome, tipo:"Novo Lead", detalhe:"Lista: "+(c.lista||"—"), proximaAcao:c.proximaAcao||"", dataProximoContato:c.dataProximoContato||"" });
    }
  });

  atividades.sort((a,b) => a.hora > b.hora ? 1 : -1);

  // ── Resumo por operador ────────────────────────────────────────────────────
  const porOperador = {};
  atividades.forEach(a => {
    const op = a.operador||"—";
    if (!porOperador[op]) porOperador[op] = { contatos:0, etapas:0, leads:0 };
    if (a.tipo==="Contato") porOperador[op].contatos++;
    else if (a.tipo==="Etapa") porOperador[op].etapas++;
    else if (a.tipo==="Novo Lead") porOperador[op].leads++;
  });

  // ── Situação atual do funil ────────────────────────────────────────────────
  const porEtapaFunil = {};
  ETAPAS.forEach(e => { porEtapaFunil[e.id] = { label:e.label, emoji:e.emoji, count:0 }; });
  todos.forEach(c => { if (porEtapaFunil[c.etapa]) porEtapaFunil[c.etapa].count++; });

  const totalAtivos = todos.filter(c=>c.etapa!=="encerrado").length;
  const focoClub = todos.filter(c=>c.objetivo==="club"||c.objetivo==="falta_uma").length;
  const experiencia = todos.filter(c=>c.etapa==="experiencia").length;

  // ── Gerar HTML do relatório ────────────────────────────────────────────────
  const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Relatório Diário — Laricas CRM — ${hojeStr}</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family: 'Helvetica Neue', Arial, sans-serif; color:#1a1a1a; background:#fff; padding:32px; font-size:12px; }
  .header { border-bottom:3px solid #C9A84C; padding-bottom:16px; margin-bottom:24px; display:flex; justify-content:space-between; align-items:flex-end; }
  .header h1 { font-size:22px; font-weight:700; color:#1a1a1a; }
  .header .sub { font-size:13px; color:#666; margin-top:4px; }
  .header .data { font-size:14px; font-weight:600; color:#C9A84C; }
  .section { margin-bottom:24px; }
  .section h2 { font-size:13px; font-weight:600; text-transform:uppercase; letter-spacing:0.08em; color:#C9A84C; border-left:3px solid #C9A84C; padding-left:8px; margin-bottom:12px; }
  .cards { display:flex; gap:12px; margin-bottom:0; flex-wrap:wrap; }
  .card { background:#f8f8f8; border-radius:8px; padding:12px 16px; flex:1; min-width:120px; }
  .card .val { font-size:24px; font-weight:700; color:#1a1a1a; }
  .card .lbl { font-size:10px; color:#888; text-transform:uppercase; letter-spacing:0.06em; margin-top:3px; }
  table { width:100%; border-collapse:collapse; }
  th { background:#f0f0f0; padding:7px 10px; text-align:left; font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:0.06em; color:#666; }
  td { padding:7px 10px; border-bottom:1px solid #f0f0f0; vertical-align:top; }
  tr:last-child td { border-bottom:none; }
  .badge { display:inline-block; padding:2px 8px; border-radius:20px; font-size:10px; font-weight:600; }
  .badge-contato { background:#e0f5f5; color:#0a7070; }
  .badge-etapa { background:#fff4e0; color:#a05c00; }
  .badge-lead { background:#f0e8ff; color:#5a3d9e; }
  .funil-row { display:flex; align-items:center; gap:8px; padding:5px 0; border-bottom:1px solid #f0f0f0; }
  .funil-bar-bg { flex:1; background:#f0f0f0; border-radius:4px; height:10px; overflow:hidden; }
  .funil-bar { height:100%; background:#C9A84C; border-radius:4px; }
  .footer { margin-top:32px; padding-top:12px; border-top:1px solid #eee; font-size:10px; color:#bbb; text-align:center; }
  @media print {
    body { padding:16px; }
    .no-print { display:none; }
  }
</style>
</head>
<body>
  <div class="header">
    <div>
      <h1>Laricas Fitness — CRM</h1>
      <div class="sub">Relatório Diário de Atividades</div>
    </div>
    <div class="data">${hojeStr}</div>
  </div>

  <div class="section">
    <h2>Resumo do dia</h2>
    <div class="cards">
      <div class="card"><div class="val">${atividades.filter(a=>a.tipo==="Contato").length}</div><div class="lbl">Contatos realizados</div></div>
      <div class="card"><div class="val">${atividades.filter(a=>a.tipo==="Etapa").length}</div><div class="lbl">Etapas movidas</div></div>
      <div class="card"><div class="val">${atividades.filter(a=>a.tipo==="Novo Lead").length}</div><div class="lbl">Novos leads</div></div>
      <div class="card"><div class="val">${atividades.length}</div><div class="lbl">Total de ações</div></div>
    </div>
  </div>

  ${Object.keys(porOperador).length > 0 ? `
  <div class="section">
    <h2>Por operador</h2>
    <table>
      <thead><tr><th>Operador</th><th>Contatos</th><th>Etapas</th><th>Novos Leads</th><th>Total</th></tr></thead>
      <tbody>
        ${Object.entries(porOperador).map(([op,d])=>`
          <tr>
            <td style="font-weight:600">${op}</td>
            <td>${d.contatos}</td>
            <td>${d.etapas}</td>
            <td>${d.leads}</td>
            <td style="font-weight:600">${d.contatos+d.etapas+d.leads}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  </div>` : ""}

  ${atividades.length > 0 ? `
  <div class="section">
    <h2>Atividades detalhadas</h2>
    <table>
      <thead><tr><th>Hora</th><th>Operador</th><th>Cliente</th><th>Tipo</th><th>Detalhe</th><th>Próxima ação</th><th>Data contato</th></tr></thead>
      <tbody>
        ${atividades.map(a=>`
          <tr>
            <td style="color:#888;white-space:nowrap">${a.hora}</td>
            <td style="white-space:nowrap">${a.operador}</td>
            <td style="font-weight:500">${a.cliente}</td>
            <td><span class="badge badge-${a.tipo.toLowerCase().replace(" ","")}">${a.tipo}</span></td>
            <td style="color:#555">${a.detalhe}</td>
            <td style="color:#444">${a.proximaAcao||"—"}</td>
            <td style="color:#888;white-space:nowrap">${a.dataProximoContato ? new Date(a.dataProximoContato+"T12:00:00").toLocaleDateString("pt-BR") : "—"}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  </div>` : `<div class="section"><h2>Atividades detalhadas</h2><p style="color:#aaa;padding:16px 0">Nenhuma atividade registrada hoje.</p></div>`}

  <div class="section">
    <h2>Situação atual do funil</h2>
    <div style="margin-bottom:12px">
      <div class="cards">
        <div class="card"><div class="val">${totalAtivos}</div><div class="lbl">Leads ativos</div></div>
        <div class="card"><div class="val">${focoClub}</div><div class="lbl">Foco Club</div></div>
        <div class="card"><div class="val">${experiencia}</div><div class="lbl">Assinantes</div></div>
      </div>
    </div>
    <table>
      <thead><tr><th>Etapa</th><th>Clientes</th><th>Distribuição</th></tr></thead>
      <tbody>
        ${Object.values(porEtapaFunil).filter(e=>e.count>0).map(e=>`
          <tr>
            <td style="font-weight:500">${e.emoji} ${e.label}</td>
            <td style="font-weight:700">${e.count}</td>
            <td style="width:50%">
              <div class="funil-bar-bg"><div class="funil-bar" style="width:${Math.min(100,Math.round(e.count/totalAtivos*100))}%"></div></div>
            </td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  </div>

  <div class="footer">Laricas Fitness CRM · Gerado em ${hoje.toLocaleString("pt-BR")} · laricas-crm.vercel.app</div>
</body>
</html>`;

  // Abrir em nova aba para imprimir/salvar como PDF
  const win = window.open("", "_blank");
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 500);
};



// ── CALENDÁRIO DE FOLLOW-UPS ───────────────────────────────────────────────
const CalendarioFollowups = ({ clientes, onAbrirCliente }) => {
  const hoje = new Date();
  const [mesAtual, setMesAtual] = React.useState(new Date(hoje.getFullYear(), hoje.getMonth(), 1));
  const [diaSel, setDiaSel] = React.useState(null); // dia selecionado para ver todos

  const diasNoMes = new Date(mesAtual.getFullYear(), mesAtual.getMonth()+1, 0).getDate();
  const primeiroDia = mesAtual.getDay();

  const statusCor = (status) => {
    if (status === "interessado")  return {bg:C.greenL,  text:C.greenD,  border:C.green};
    if (status === "respondeu")    return {bg:C.blueL,   text:C.blueD,   border:C.blue};
    if (status === "link_enviado") return {bg:C.amberL,  text:C.amberD,  border:C.amber};
    if (status === "contatado")    return {bg:C.tealL,   text:C.tealD,   border:C.teal};
    if (status === "follow_up")    return {bg:C.purpleL, text:C.purpleD, border:C.purple};
    if (status === "nao_agora")    return {bg:"#f5f5f5", text:"#888",    border:"#ccc"};
    return {bg:"var(--color-background-secondary)", text:"var(--color-text-tertiary)", border:"var(--color-border-tertiary)"};
  };

  // Deduplicar por ID antes de agrupar
  const vistos = new Set();
  const clientesUnicos = clientes.filter(c => {
    if (!c.id || vistos.has(c.id)) return false;
    vistos.add(c.id);
    return true;
  });

  // Agrupar por data de follow-up
  const porDia = {};
  clientesUnicos.forEach(c => {
    if (!c.proximoFollowup || !c.statusClub) return;
    const [ano, mes, dia] = c.proximoFollowup.split("-").map(Number);
    if (ano === mesAtual.getFullYear() && mes === mesAtual.getMonth()+1) {
      if (!porDia[dia]) porDia[dia] = [];
      porDia[dia].push(c);
    }
  });

  const mesLabel = mesAtual.toLocaleDateString("pt-BR", {month:"long", year:"numeric"});
  const diasSemana = ["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"];
  const hojeISO = hoje.toISOString().split("T")[0];
  const diaSelecionadoISO = diaSel ? `${mesAtual.getFullYear()}-${String(mesAtual.getMonth()+1).padStart(2,"0")}-${String(diaSel).padStart(2,"0")}` : null;

  return (
    <div>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:16}}>
        <button onClick={()=>{setMesAtual(new Date(mesAtual.getFullYear(), mesAtual.getMonth()-1, 1));setDiaSel(null);}}
          style={{padding:"4px 12px",borderRadius:8,cursor:"pointer",background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)",fontSize:16,color:"var(--color-text-primary)"}}>‹</button>
        <div style={{flex:1,textAlign:"center",fontSize:14,fontWeight:500,color:"var(--color-text-primary)",textTransform:"capitalize"}}>{mesLabel}</div>
        <button onClick={()=>{setMesAtual(new Date(mesAtual.getFullYear(), mesAtual.getMonth()+1, 1));setDiaSel(null);}}
          style={{padding:"4px 12px",borderRadius:8,cursor:"pointer",background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)",fontSize:16,color:"var(--color-text-primary)"}}>›</button>
      </div>

      {/* Legenda */}
      <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:12}}>
        {[["Interessado",C.greenL],["Respondeu",C.blueL],["Link enviado",C.amberL],
          ["Contatado",C.tealL],["Follow-up",C.purpleL],["Não agora","#f5f5f5"]].map(([label,bg])=>(
          <div key={label} style={{display:"flex",alignItems:"center",gap:4,fontSize:10}}>
            <div style={{width:10,height:10,borderRadius:3,background:bg,border:"0.5px solid #ccc"}}/>
            <span style={{color:"var(--color-text-tertiary)"}}>{label}</span>
          </div>
        ))}
      </div>

      <div style={{display:"grid",gridTemplateColumns:diaSel?"1fr 260px":"1fr",gap:12}}>
        {/* Grade do calendário */}
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2,marginBottom:4}}>
            {diasSemana.map(d=>(
              <div key={d} style={{textAlign:"center",fontSize:10,fontWeight:500,color:"var(--color-text-tertiary)",padding:"4px 0"}}>{d}</div>
            ))}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2}}>
            {Array.from({length:primeiroDia}).map((_,i)=><div key={"e"+i}/>)}
            {Array.from({length:diasNoMes}).map((_,i)=>{
              const dia = i+1;
              const diaISO = `${mesAtual.getFullYear()}-${String(mesAtual.getMonth()+1).padStart(2,"0")}-${String(dia).padStart(2,"0")}`;
              const isHoje = diaISO === hojeISO;
              const isPassado = diaISO < hojeISO;
              const isSel = diaSel === dia;
              const clts = porDia[dia]||[];
              return (
                <div key={dia} onClick={()=>setDiaSel(isSel?null:dia)}
                  style={{minHeight:70,border:"0.5px solid "+(isSel?C.teal:isHoje?"#aaa":"var(--color-border-tertiary)"),
                    borderRadius:6,padding:"4px",cursor:clts.length>0?"pointer":"default",
                    background:isSel?C.tealL:isHoje?"var(--color-background-secondary)":"transparent",
                    opacity:isPassado&&clts.length===0?0.35:1,transition:"all 0.1s"}}>
                  <div style={{fontSize:10,fontWeight:isHoje||isSel?600:400,
                    color:isSel?C.tealD:isHoje?"#333":"var(--color-text-tertiary)",marginBottom:2}}>{dia}</div>
                  {clts.slice(0,4).map(c=>{
                    const cor = statusCor(c.statusClub);
                    return (
                      <div key={c.id}
                        style={{width:"100%",padding:"1px 4px",borderRadius:4,marginBottom:1,fontSize:9,fontWeight:500,
                          background:cor.bg,color:cor.text,border:"0.5px solid "+cor.border,
                          overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                        {(c.nome||"").split(" ")[0]}
                      </div>
                    );
                  })}
                  {clts.length>4&&(
                    <div style={{fontSize:8,color:C.tealD,fontWeight:500,textAlign:"center",
                      background:C.tealL,borderRadius:4,padding:"1px"}}>
                      +{clts.length-4} mais
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* Painel lateral — todos do dia selecionado */}
        {diaSel&&(
          <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",height:"fit-content",position:"sticky",top:0}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
              <div style={{flex:1,fontSize:13,fontWeight:500,color:"var(--color-text-primary)"}}>
                📅 {String(diaSel).padStart(2,"0")}/{String(mesAtual.getMonth()+1).padStart(2,"0")}
              </div>
              <button onClick={()=>setDiaSel(null)}
                style={{fontSize:12,color:"var(--color-text-tertiary)",background:"none",border:"none",cursor:"pointer"}}>✕</button>
            </div>
            {(porDia[diaSel]||[]).length===0&&(
              <div style={{fontSize:12,color:"var(--color-text-tertiary)"}}>Nenhum follow-up neste dia.</div>
            )}
            {(porDia[diaSel]||[]).map(c=>{
              const cor = statusCor(c.statusClub);
              const st = STATUS_CLUB.find(s=>s.id===c.statusClub);
              return (
                <button key={c.id} onClick={()=>onAbrirCliente&&onAbrirCliente(c)}
                  style={{width:"100%",textAlign:"left",padding:"8px 10px",borderRadius:8,marginBottom:6,
                    background:cor.bg,border:"0.5px solid "+cor.border,cursor:"pointer"}}>
                  <div style={{fontSize:12,fontWeight:500,color:cor.text}}>{c.nome}</div>
                  <div style={{fontSize:10,color:cor.text,opacity:0.8,marginTop:1}}>
                    {st?.emoji} {st?.label} · {c.p||0}p · R${(c.gasto||0).toFixed(0)}
                  </div>
                </button>
              );
            })}
          </div>
        )}
      </div>

      {/* Resumo do mês */}
      {Object.keys(porDia).length>0&&(
        <div style={{marginTop:16,background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px"}}>
          <div style={{fontSize:11,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:8}}>
            Follow-ups neste mês — {Object.values(porDia).flat().length} total
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:4}}>
            {Object.entries(porDia).sort(([a],[b])=>Number(a)-Number(b)).map(([dia,clts])=>(
              <button key={dia} onClick={()=>setDiaSel(Number(dia))}
                style={{display:"flex",alignItems:"center",gap:6,fontSize:11,textAlign:"left",
                  padding:"4px 6px",borderRadius:6,border:"none",cursor:"pointer",
                  background:diaSel===Number(dia)?C.tealL:"transparent",color:"var(--color-text-primary)"}}>
                <div style={{fontSize:10,color:"var(--color-text-tertiary)",minWidth:32}}>
                  {String(dia).padStart(2,"0")}/{String(mesAtual.getMonth()+1).padStart(2,"0")}
                </div>
                <div style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                  {clts.map(c=>(c.nome||"").split(" ")[0]).join(", ")}
                </div>
                <span style={{fontSize:10,fontWeight:600,color:C.tealD,flexShrink:0}}>{clts.length}</span>
              </button>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};


// ── RITSPAY INTEGRATION ────────────────────────────────────────────────────
const RITSPAY_STORAGE_KEY = "ritspay_cfg";
const RITSPAY_TOKEN_KEY   = "ritspay_token";

const ritspaySaveCfg = (cfg) => {
  try { localStorage.setItem(RITSPAY_STORAGE_KEY, JSON.stringify(cfg)); } catch(e) {}
};
const ritspayLoadCfg = () => {
  try { const r = localStorage.getItem(RITSPAY_STORAGE_KEY); return r ? JSON.parse(r) : {}; } catch(e) { return {}; }
};
const ritspayGetToken = () => {
  try { return localStorage.getItem(RITSPAY_TOKEN_KEY) || ""; } catch(e) { return ""; }
};
const ritspaySetToken = (t) => {
  try { localStorage.setItem(RITSPAY_TOKEN_KEY, t); } catch(e) {}
};

// Chama a API RitsPay com token
const ritspayFetch = async (path, token, opts = {}) => {
  const r = await fetch("https://api.ritspay.com" + path, {
    ...opts,
    headers: { "Content-Type": "application/json", "Authorization": "Bearer " + token, ...(opts.headers||{}) }
  });
  if (!r.ok) throw new Error("HTTP " + r.status);
  return r.json();
};

// Mapeia status RitsPay → statusAssinatura CRM
const mapRitsStatus = (sub) => {
  if (!sub) return "ativo";
  // canceled_at sempre prevalece — independente do status de cobrança
  if (sub.canceled_at || sub.deleted_at) return "cancelado";
  const s = (sub.status || "").toLowerCase();
  const cycle = typeof sub.cycle === "number" ? sub.cycle : 0;
  if (s === "canceled" || s === "cancelled" || s === "inactive") return "cancelado";
  if (s === "paused" || s === "suspended") return "pausado";
  // Falha na RENOVAÇÃO: só conta como atrasado se já teve pelo menos 1 ciclo pago
  if ((s === "past_due" || s === "unpaid" || s === "overdue" || sub.overdue_at) && cycle >= 1) return "atrasado";
  if ((s === "past_due" || s === "unpaid" || s === "overdue") && cycle === 0) return "nunca_ativado";
  return "ativo";
};

// Infere tipo de assinatura pelo período
const mapRitsPlan = (sub) => {
  if (!sub) return "";
  // Usa slug e name do plano (campos confirmados da API)
  const slug = (sub.plan?.slug || sub.plan?.code || "").toLowerCase();
  const name = (sub.plan?.name || "").toLowerCase();
  if (slug.includes("anual") || name.includes("anual") || name.includes("annual")) return "Anual";
  if (slug.includes("semestral") || name.includes("semestral")) return "Semestral";
  if (slug.includes("trimestral") || name.includes("trimestral")) return "Trimestral";
  return "";
};

const RitsPaySyncModal = ({ onClose, onSyncDone }) => {
  const [step, setStep] = React.useState("login"); // login | twofa | syncing | done | error
  const [email, setEmail] = React.useState(() => ritspayLoadCfg().email || "");
  const [senha, setSenha] = React.useState("");
  const [tenantId, setTenantId] = React.useState(() => ritspayLoadCfg().tenantId || "TEN-1G57I7LIVD8K0F8M");
  const [codigo2fa, setCodigo2fa] = React.useState("");
  const [loginToken, setLoginToken] = React.useState(""); // token intermediário antes do 2FA
  const [mensagem, setMensagem] = React.useState("");
  const [resultado, setResultado] = React.useState(null);
  const [loginResp, setLoginResp] = React.useState(null);

  // Ao abrir: tenta token salvo antes de pedir login
  React.useEffect(() => {
    const tokenSalvo = ritspayGetToken();
    if (!tokenSalvo) return;
    setStep("syncing");
    setMensagem("Usando sessão salva...");
    sincronizar(tokenSalvo).catch(() => {
      ritspaySetToken(""); // token expirou
      setStep("login");
      setMensagem("Sessão expirada. Faça login novamente.");
    });
  }, []);

  const fazerLogin = async () => {
    if (!email.trim() || !senha.trim()) { setMensagem("Preencha email e senha."); return; }
    setStep("syncing");
    setMensagem("Fazendo login no RitsPay...");
    try {
      const r = await fetch("https://api.ritspay.com/account/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email: email.trim(), password: senha })
      });
      const data = await r.json();
      if (!r.ok) throw new Error(data.detail || data.message || JSON.stringify(data).slice(0,200));
      setLoginResp(data); // Guarda resposta completa para debug
      const token = data.access_token || data.token || data.data?.access_token || data.data?.token || "";
      setLoginToken(token);
      ritspaySaveCfg({ email: email.trim(), tenantId: tenantId.trim() });
      setStep("twofa");
      setMensagem("Código enviado para o email. Digite abaixo:");
    } catch(e) {
      setStep("error");
      setMensagem("Erro no login: " + e.message);
    }
  };

  const confirmar2fa = async () => {
    if (!codigo2fa.trim()) { setMensagem("Digite o código."); return; }
    setStep("syncing");
    setMensagem("Autenticando 2FA...");
    try {
      // Campo correto confirmado pela API: two_factor_code
      const body2fa = {
        two_factor_code: codigo2fa.trim(),
        email: email.trim(),
      };
      const headers2fa = {
        "Content-Type": "application/json",
        ...(loginToken ? { "Authorization": "Bearer " + loginToken } : {}),
      };
      const r = await fetch("https://api.ritspay.com/account/auth/two_factor", {
        method: "POST",
        headers: headers2fa,
        body: JSON.stringify(body2fa)
      });
      const data = await r.json();
      if (!r.ok) {
        // Mostra resposta completa para debug
        const detalhe = JSON.stringify(data).slice(0,200);
        throw new Error(detalhe);
      }
      const token = data.access_token || data.token || data.data?.access_token || loginToken;
      ritspaySetToken(token);
      await sincronizar(token);
    } catch(e) {
      setStep("error");
      setMensagem("Erro no 2FA: " + e.message);
    }
  };

  const sincronizar = async (token) => {
    setStep("syncing");
    setMensagem("Buscando assinantes no RitsPay...");
    try {
      const tenant = tenantId.trim();

      // Helper para buscar todas as páginas — estrutura {data, links, meta}
      const fetchAll = async (path) => {
        const todos = [];
        let pagina = 1;
        while (pagina <= 50) {
          try {
            const sep = path.includes("?") ? "&" : "?";
            const resp = await ritspayFetch(`${path}${sep}page=${pagina}`, token);
            const items = Array.isArray(resp?.data) ? resp.data
              : Array.isArray(resp?.results) ? resp.results
              : Array.isArray(resp) ? resp : [];
            if (items.length === 0) break;
            todos.push(...items);
            const total = resp?.meta?.total ?? resp?.meta?.total_count ?? null;
            setMensagem(`${path.includes("subscription")?"Assinaturas":"Registros"}: ${todos.length}${total?"/"+total:""}...`);
            // Para quando não tem next ou já buscou o total
            const hasNext = resp?.links?.next != null;
            if (!hasNext) break;
            if (total && todos.length >= total) break;
            pagina++;
          } catch(pageErr) {
            // Se falhar uma página, para aqui mas não cancela tudo
            break;
          }
        }
        return todos;
      };

      // 1. Busca todas as subscriptions
      const subs = await fetchAll(`/sales/${tenant}/subscriptions`);
      if (subs.length === 0) {
        setStep("error");
        setMensagem("Sem assinaturas encontradas. Verifique o Tenant ID.");
        return;
      }
      setMensagem(`${subs.length} assinaturas. Buscando histórico de compras...`);

      // 2. Busca purchases para calcular ticket médio e LTV real da assinatura
      const purchases = await fetchAll(`/sales/${tenant}/purchases`);
      // Agrupa purchases por customer ID → { custId: [{cycle, total}...] }
      const purchasesByCust = {};
      purchases.forEach(p => {
        const cid = p.customer?.id || p.customer_id || "";
        if (!cid) return;
        if (!purchasesByCust[cid]) purchasesByCust[cid] = [];
        purchasesByCust[cid].push({
          total: parseFloat(p.total || p.amount || 0) / 100,
          date: (p.created_at || p.paid_at || "").split("T")[0],
        });
      });

      // 3. Busca customers para mapa de nomes (fallback de matching)
      setMensagem("Buscando clientes RitsPay...");
      const custs = await fetchAll(`/sales/${tenant}/customers`);
      const custMap = {}; // id → customer
      custs.forEach(c => { if (c.id) custMap[c.id] = c; });

      // 4. Ordena subscriptions: ativo primeiro
      const STATUS_PRIO = { "active": 0, "overdue": 1, "past_due": 1, "paused": 2, "canceled": 3, "inactive": 4 };
      const subsPriorizadas = [...subs].sort((a, b) =>
        (STATUS_PRIO[(a.status||"").toLowerCase()] ?? 9) - (STATUS_PRIO[(b.status||"").toLowerCase()] ?? 9)
      );

      // 5. Busca todos os clientes do CRM
      const crmClientes = await dbGetAll();
      let atualizados = 0;
      const detalhes = [];
      const processados = new Set();

      setMensagem(`Sincronizando ${subs.length} assinaturas com ${crmClientes.length} clientes CRM...`);

      for (const sub of subsPriorizadas) {
        // Email: vem direto em sub.customer.email (confirmado pela API)
        const custRef = sub.customer || {};
        const custId = custRef.id || "";
        let custEmail = custRef.email || "";
        const custNome = custRef.name || custMap[custId]?.name || "";

        if (processados.has(custEmail.toLowerCase())) continue;

        // Encontra no CRM: por email → por telefone (nome removido — falsos positivos)
        let crmCliente = crmClientes.find(c =>
          custEmail && (
            (c.email||"").toLowerCase().trim() === custEmail.toLowerCase().trim() ||
            (c.emailClub||"").toLowerCase().trim() === custEmail.toLowerCase().trim()
          )
        );
        // Fallback por telefone (9 últimos dígitos)
        if (!crmCliente) {
          const telefoneRits = (custRef.phone || custMap[custId]?.phone || "").replace(/\D/g,"").slice(-9);
          if (telefoneRits.length >= 8) {
            crmCliente = crmClientes.find(c => {
              const telCrm = (c.telefone||"").replace(/\D/g,"").slice(-9);
              return telCrm === telefoneRits && telCrm.length >= 8;
            });
          }
        }

        if (!crmCliente) {
          detalhes.push({ nome: custNome || custEmail || custId, status: "não encontrado" });
          continue;
        }

        if (custEmail) processados.add(custEmail.toLowerCase());

        const novoStatus = mapRitsStatus(sub);
        // Não atualizar CRM para quem nunca ativou (falha no primeiro pagamento)
        if (novoStatus === "nunca_ativado") {
          detalhes.push({ nome: crmCliente.nome, status: "nunca ativado — ignorado" });
          continue;
        }
        const novoPlano  = mapRitsPlan(sub);
        const proximaCob = (sub.next_billing_at || "").split("T")[0] || "";
        const dataInicio = (sub.start_at || sub.created_at || "").split("T")[0] || "";
        const cicloAtual = typeof sub.cycle === "number" ? sub.cycle : null;

        // Valor atual da assinatura (total em centavos)
        const valorAtual = parseFloat(sub.total || sub.subscription_price || 0) / 100;
        const valorMensalCalc = valorAtual > 0 ? valorAtual.toFixed(2) : "";

        // LTV da assinatura: soma dos purchases reais deste cliente no RitsPay
        const custPurchases = purchasesByCust[custId] || [];
        const ltvAssinaturaRits = custPurchases.reduce((s, p) => s + p.total, 0);
        const ticketMedioRits = custPurchases.length > 0
          ? (ltvAssinaturaRits / custPurchases.length).toFixed(2) : "";

        // Fim da fidelidade
        const mesesPlano = novoPlano==="Anual"?12:novoPlano==="Semestral"?6:novoPlano==="Trimestral"?3:0;
        const fimFidelidade = dataInicio && mesesPlano > 0 ? (() => {
          const d = new Date(dataInicio+"T12:00:00");
          d.setMonth(d.getMonth() + mesesPlano);
          return d.toISOString().split("T")[0];
        })() : "";

        const atualizado = {
          ...crmCliente,
          statusAssinatura: novoStatus,
          cancelado: novoStatus === "cancelado",
          falhaRenovacao: novoStatus === "atrasado",
          ...(novoPlano ? { tipoAssinatura: novoPlano } : {}),
          ...(valorMensalCalc ? { valorMensal: valorMensalCalc } : {}),
          ...(proximaCob ? { proximaCobranca: proximaCob } : {}),
          ...(dataInicio ? { dataInicioAssinatura: dataInicio } : {}),
          ...(cicloAtual !== null ? { cicloAtualClub: cicloAtual } : {}),
          ...(fimFidelidade ? { dataFimFidelidade: fimFidelidade } : {}),
          ...(ltvAssinaturaRits > 0 ? { ltvAssinatura: ltvAssinaturaRits.toFixed(2) } : {}),
          ...(ticketMedioRits ? { ticketMedioClub: ticketMedioRits } : {}),
          subscriptionIdRits: sub.id || "",
          customerIdRits: custId,
        };

        await dbSave(atualizado);
        atualizados++;
        detalhes.push({ nome: crmCliente.nome, status: novoStatus });
      }

      const naoEncontrados = detalhes.filter(d=>d.status==="não encontrado").length;
      setResultado({ total: subs.length, atualizados, naoEncontrados, detalhes });
      setStep("done");
      setMensagem(`${atualizados} atualizados · ${naoEncontrados} não encontrados no CRM`);
      if (onSyncDone) onSyncDone();
    } catch(e) {
      setStep("error");
      setMensagem("Erro na sincronização: " + e.message);
    }
  };

  const C2 = { coral: "#ef4444", coralL: "#fef2f2", green: "#10b981", greenL: "#ecfdf5", teal: "#0d9488", tealL: "#f0fdfa", tealD: "#0f766e", purple: "#7c3aed", purpleL: "#ede9fe", amber: "#f59e0b", amberL: "#fffbeb", amberD: "#b45309" };

  return (
    <div style={{ position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:9999 }}>
      <div style={{ width:440,background:"var(--color-background-primary)",borderRadius:16,padding:"24px 28px",boxShadow:"0 8px 40px rgba(0,0,0,0.15)" }}>
        <div style={{ display:"flex",alignItems:"center",gap:10,marginBottom:20 }}>
          <div style={{ fontSize:20 }}>🔄</div>
          <div style={{ flex:1 }}>
            <div style={{ fontSize:15,fontWeight:600,color:"var(--color-text-primary)" }}>Sincronizar com RitsPay</div>
            <div style={{ fontSize:12,color:"var(--color-text-tertiary)" }}>Status, próxima cobrança e valor mensal</div>
          </div>
          <button onClick={onClose} style={{ fontSize:18,cursor:"pointer",background:"none",border:"none",color:"var(--color-text-tertiary)" }}>✕</button>
        </div>

        {step === "login" && (
          <div>
            <div style={{ marginBottom:10 }}>
              <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Email RitsPay</div>
              <input value={email} onChange={e=>setEmail(e.target.value)} style={inp()} placeholder="seu@email.com"/>
            </div>
            <div style={{ marginBottom:10 }}>
              <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Senha</div>
              <input type="password" value={senha} onChange={e=>setSenha(e.target.value)} onKeyDown={e=>e.key==="Enter"&&fazerLogin()} style={inp()} placeholder="••••••••"/>
            </div>
            <div style={{ marginBottom:16 }}>
              <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Tenant ID</div>
              <input value={tenantId} onChange={e=>setTenantId(e.target.value)} style={inp()} placeholder="TEN-..."/>
            </div>
            {mensagem && <div style={{ fontSize:12,color:"#e53e3e",marginBottom:10 }}>{mensagem}</div>}
            <button onClick={fazerLogin}
              style={{ width:"100%",padding:"11px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",background:C.teal,color:"#fff",border:"none",marginBottom:8 }}>
              Entrar e buscar código 2FA
            </button>
            {ritspayGetToken()&&(
              <button onClick={()=>{ setStep("syncing"); setMensagem("Testando estrutura da API..."); ritspayFetch(`/sales/${tenantId.trim()}/subscriptions`, ritspayGetToken()).then(r=>{ const items=Array.isArray(r)?r:r?.results||r?.items||r?.data||[]; setResultado({debug:true,sub0:items[0],cust0:null,total:items.length}); setStep("done"); setMensagem("Estrutura carregada."); }).catch(e=>{ setStep("error"); setMensagem("Erro: "+e.message); }); }}
                style={{ width:"100%",padding:"8px",borderRadius:10,fontSize:12,cursor:"pointer",background:"none",border:"0.5px solid "+C.teal,color:C.tealD }}>
                🔍 Ver estrutura da API (sessão salva)
              </button>
            )}
          </div>
        )}

        {step === "twofa" && (
          <div>
            <div style={{ fontSize:13,color:"var(--color-text-primary)",marginBottom:16,lineHeight:1.5 }}>
              {mensagem}
            </div>
            {loginResp&&(
              <div style={{ fontSize:10,color:"var(--color-text-tertiary)",background:"var(--color-background-secondary)",borderRadius:6,padding:"8px",marginBottom:12,fontFamily:"monospace",wordBreak:"break-all" }}>
                Resposta do login: {JSON.stringify(loginResp)}
              </div>
            )}
            <div style={{ marginBottom:16 }}>
              <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Código recebido no email</div>
              <input value={codigo2fa} onChange={e=>setCodigo2fa(e.target.value)} onKeyDown={e=>e.key==="Enter"&&confirmar2fa()} style={inp()} placeholder="000000" autoFocus/>
            </div>
            <button onClick={confirmar2fa}
              style={{ width:"100%",padding:"11px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",background:C.teal,color:"#fff",border:"none" }}>
              Confirmar e sincronizar
            </button>
          </div>
        )}

        {step === "syncing" && (
          <div style={{ textAlign:"center",padding:"20px 0" }}>
            <div style={{ fontSize:32,marginBottom:12 }}>⏳</div>
            <div style={{ fontSize:13,color:"var(--color-text-primary)" }}>{mensagem}</div>
          </div>
        )}

        {step === "done" && resultado && (
          <div>
            {resultado.debug ? (
              <div>
                <div style={{fontSize:12,fontWeight:500,color:C.tealD,marginBottom:8}}>🔍 Estrutura da API ({resultado.total} assinaturas encontradas)</div>
                <div style={{marginBottom:8}}>
                  <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4}}>Assinatura [0]:</div>
                  <pre style={{fontSize:9,background:"var(--color-background-secondary)",borderRadius:6,padding:8,overflow:"auto",maxHeight:200,whiteSpace:"pre-wrap",wordBreak:"break-all"}}>
                    {JSON.stringify(resultado.sub0, null, 2)}
                  </pre>
                </div>
                {resultado.cust0&&<div>
                  <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:4}}>Customer [0]:</div>
                  <pre style={{fontSize:9,background:"var(--color-background-secondary)",borderRadius:6,padding:8,overflow:"auto",maxHeight:160,whiteSpace:"pre-wrap",wordBreak:"break-all"}}>
                    {JSON.stringify(resultado.cust0, null, 2)}
                  </pre>
                </div>}
              </div>
            ) : (
              <div>
                <div style={{ background:C.greenL,border:"0.5px solid "+C.green,borderRadius:10,padding:"12px 14px",marginBottom:12 }}>
                  <div style={{ fontSize:14,fontWeight:600,color:C.greenD,marginBottom:4 }}>✅ Sincronização concluída</div>
                  <div style={{ fontSize:12,color:C.greenD }}>{resultado.atualizados} de {resultado.total} atualizados</div>
                  {resultado.naoEncontrados>0&&<div style={{ fontSize:11,color:C.amberD,marginTop:4 }}>⚠ {resultado.naoEncontrados} não encontrados no CRM (email/tel/nome diferente)</div>}
                </div>
                <div style={{ maxHeight:200,overflowY:"auto",marginBottom:12 }}>
                  {resultado.detalhes.map((d,i)=>(
                    <div key={i} style={{ display:"flex",justifyContent:"space-between",fontSize:11,padding:"3px 0",borderBottom:"0.5px solid var(--color-border-tertiary)" }}>
                      <span style={{ color:d.status==="não encontrado"?"var(--color-text-tertiary)":"var(--color-text-primary)" }}>{d.nome}</span>
                      <span style={{ color:d.status==="ativo"?C.green:d.status==="cancelado"?C.coral:d.status==="atrasado"?C.amber:d.status==="pausado"?C.amber:"#aaa",fontWeight:500 }}>{d.status}</span>
                    </div>
                  ))}
                </div>
                <button onClick={onClose} style={{ width:"100%",padding:"10px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",background:C.teal,color:"#fff",border:"none" }}>
                  Fechar
                </button>
              </div>
            )}
          </div>
        )}

        {step === "error" && (
          <div>
            <div style={{ background:C.coralL,border:"0.5px solid "+C.coral,borderRadius:10,padding:"12px 14px",marginBottom:12,fontSize:13,color:C.coralD }}>
              ❌ {mensagem}
            </div>
            <button onClick={()=>setStep("login")} style={{ width:"100%",padding:"10px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",background:"none",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-secondary)" }}>
              Tentar novamente
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

// ── ANALYTICS RITSPAY — dados direto da API, sem sync ───────────────────────
const AnalyticsRitsPay = ({ onAbrirPerfil }) => {
  const [dados, setDados] = React.useState(null);
  const [loading, setLoading] = React.useState(false);
  const [erro, setErro] = React.useState("");
  const [ordenar, setOrdenar] = React.useState("nome");
  const [filtroStatus, setFiltroStatus] = React.useState("todos");

  const mapStatus = (s, sub) => {
    const st = (s||"").toLowerCase();
    if (st==="canceled"||st==="cancelled"||st==="inactive") return "cancelado";
    if (st==="paused"||st==="suspended") return "pausado";
    if (st==="past_due"||st==="unpaid"||st==="overdue"||sub?.overdue_at) return "atrasado";
    return "ativo";
  };

  const mapPlano = (sub) => {
    const slug = (sub?.plan?.slug||sub?.plan?.code||"").toLowerCase();
    const name = (sub?.plan?.name||"").toLowerCase();
    if (slug.includes("anual")||name.includes("anual")) return "Anual";
    if (slug.includes("semestral")||name.includes("semestral")) return "Semestral";
    if (slug.includes("trimestral")||name.includes("trimestral")) return "Trimestral";
    return "—";
  };

  const mesesPlano = p => p==="Anual"?12:p==="Semestral"?6:p==="Trimestral"?3:0;

  const carregar = async () => {
    const token = ritspayGetToken();
    if (!token) { setErro("Faça login no RitsPay primeiro (botão Sincronizar)."); return; }
    setLoading(true); setErro("");
    try {
      const cfg = ritspayLoadCfg();
      const tenant = cfg.tenantId || "TEN-1G57I7LIVD8K0F8M";

      const fetchAll = async (path) => {
        const todos = [];
        let pagina = 1;
        while (pagina <= 50) {
          try {
            const sep = path.includes("?")?"&":"?";
            const resp = await ritspayFetch(`${path}${sep}page=${pagina}`, token);
            const items = Array.isArray(resp?.data)?resp.data:Array.isArray(resp?.results)?resp.results:Array.isArray(resp)?resp:[];
            if (!items.length) break;
            todos.push(...items);
            if (!resp?.links?.next) break;
            if (resp?.meta?.total && todos.length >= resp.meta.total) break;
            pagina++;
          } catch(e) { break; }
        }
        return todos;
      };

      // Busca subscriptions e purchases em paralelo
      const [subs, purchases, crmClientes] = await Promise.all([
        fetchAll(`/sales/${tenant}/subscriptions`),
        fetchAll(`/sales/${tenant}/purchases`),
        dbGetAll(),
      ]);

      // Agrupa purchases por customer ID — apenas pagamentos confirmados
      const purchByCustomer = {};
      purchases.forEach(p => {
        const cid = p.customer?.id||p.customer_id||"";
        if (!cid) return;
        // Ignora tentativas com pagamento falho ou valor zero
        const st = (p.status||p.payment_status||"").toLowerCase();
        if (st==="failed"||st==="refused"||st==="canceled"||st==="declined"||st==="error"||st==="chargeback") return;
        const valor = parseFloat(p.total||p.amount||0)/100;
        if (valor <= 0) return;
        if (!purchByCustomer[cid]) purchByCustomer[cid] = [];
        purchByCustomer[cid].push(valor);
      });

      // Mapa email → gasto avulso do CRM
      const gastoPorEmail = {};
      crmClientes.forEach(c => {
        const emails = [(c.email||"").toLowerCase().trim(), (c.emailClub||"").toLowerCase().trim()].filter(Boolean);
        emails.forEach(e => { if(e) gastoPorEmail[e] = c.gasto||0; });
      });

      // Processa cada assinatura (ativa tem prioridade sobre cancelada)
      const STATUS_PRIO = {active:0,overdue:1,past_due:1,paused:2,canceled:3,inactive:4};
      const subsPrio = [...subs].sort((a,b)=>(STATUS_PRIO[(a.status||"").toLowerCase()]??9)-(STATUS_PRIO[(b.status||"").toLowerCase()]??9));

      const vistos = new Set();
      const rows = [];
      for (const sub of subsPrio) {
        const email = (sub.customer?.email||"").toLowerCase().trim();
        if (vistos.has(sub.customer?.id||email)) continue;
        vistos.add(sub.customer?.id||email);

        const status = mapStatus(sub.status, sub);
        const custId = sub.customer?.id || "";
        const custPurchases = purchByCustomer[custId] || [];
        const cycle = typeof sub.cycle === "number" ? sub.cycle : 0;

        // Regra definitiva: purchases são a fonte de verdade absoluta.
        // Sem purchases = nunca pagou, EXCETO se a assinatura foi criada há menos de 3 dias
        // (pagamento ainda pode estar em processamento)
        const temPurchase = custPurchases.length > 0;
        if (!temPurchase) {
          const criadaEm = new Date(sub.created_at || sub.start_at || "2000-01-01");
          const diasDesde = Math.round((new Date() - criadaEm) / 86400000);
          if (diasDesde > 3) continue; // mais de 3 dias sem purchase = nunca pagou
        }
        const plano = mapPlano(sub);
        const meses = mesesPlano(plano);
        const cicloAtual = typeof sub.cycle==="number" ? sub.cycle : null;
        const cicloDisplay = cicloAtual!==null && meses>0 ? `${cicloAtual}° (${cicloAtual}/${meses})` : cicloAtual!==null ? `${cicloAtual}°` : "—";

        // Ticket médio: média das purchases reais deste customer
        const ticketMedio = custPurchases.length>0
          ? (custPurchases.reduce((s,v)=>s+v,0)/custPurchases.length)
          : (parseFloat(sub.total||sub.subscription_price||0)/100);

        // LTV club: soma de todas as purchases reais
        const ltvClub = custPurchases.length>0
          ? custPurchases.reduce((s,v)=>s+v,0)
          : ticketMedio*(cicloAtual||1);

        // LTV total = club + avulsos do CRM
        const gastoAvulso = gastoPorEmail[email] || 0;
        const ltvTotal = ltvClub + gastoAvulso;

        const proximaCob = (sub.next_billing_at||"").split("T")[0]||"—";
        const nome = sub.customer?.name || email || "—";

        // Churn score: 0-100 (maior = mais risco)
        let churnScore = 0;
        if (status === "atrasado") churnScore += 40;
        if (status === "pausado") churnScore += 20;
        if (cicloAtual === 1) churnScore += 20; // primeiro mês, risco alto de não renovar
        if (custPurchases.length > 1) {
          const valores = custPurchases.slice(-3);
          const trend = valores[valores.length-1] - valores[0];
          if (trend < -20) churnScore += 15; // ticket caindo
        }
        if (cicloAtual !== null && meses > 0 && cicloAtual === meses) churnScore += 10; // renovação iminente

        // Aniversário de assinatura
        const dataIni = (sub.start_at || sub.created_at || "").split("T")[0];
        let aniversario = null;
        if (dataIni && cicloAtual !== null) {
          const marcos = [3, 6, 12];
          for (const m of marcos) {
            if (cicloAtual === m - 1) { // 1 mês antes do aniversário
              const d = new Date(dataIni+"T12:00:00");
              d.setMonth(d.getMonth() + m);
              aniversario = { meses: m, data: d.toLocaleDateString("pt-BR") };
              break;
            }
          }
        }

        // Candidato a upgrade
        const upgradeSugerido = plano === "Trimestral" && cicloAtual >= 2 && status === "ativo" && churnScore < 30
          ? "Semestral" : plano === "Semestral" && cicloAtual >= 4 && status === "ativo" && churnScore < 30
          ? "Anual" : null;

        rows.push({ nome, email, plano, cicloDisplay, cicloAtual, ticketMedio, ltvClub, ltvTotal, gastoAvulso, proximaCob, status, id: sub.id, churnScore, aniversario, upgradeSugerido });
      }

      // DEBUG TEMPORÁRIO — mostra dados brutos de cada assinatura
      setDados([...rows, { _debug: true, _rawSubs: subsPrio.map(s => ({
        nome: s.customer?.name,
        status: s.status,
        cycle: s.cycle,
        overdue_at: s.overdue_at,
        canceled_at: s.canceled_at,
        created_at: s.created_at?.split("T")[0],
        purchases: (purchByCustomer[s.customer?.id||""]||[]).length,
        id: s.id,
      })) }]);
      setLoading(false);
      return;
    } catch(e) {
      // Token expirado — limpar e pedir relogin
      if (e.message.includes("401")||e.message.includes("403")) {
        ritspaySetToken("");
        setErro("Sessão expirada. Clique em 'Sincronizar com RitsPay' para reconectar.");
      } else {
        setErro("Erro: " + e.message);
      }
    }
    setLoading(false);
  };

  React.useEffect(() => { carregar(); }, []);

  const COR_STATUS = {ativo:C.green,pausado:C.amber,atrasado:C.coral,cancelado:"#aaa"};
  const statusOpts = ["todos","ativo","pausado","atrasado","cancelado"];

  const filtrados = (dados||[]).filter(r => filtroStatus==="todos"||r.status===filtroStatus).filter(r=>!r._debug);
  const ordenados = [...filtrados].sort((a,b)=>{
    if (ordenar==="nome") return (a.nome||"").localeCompare(b.nome||"");
    if (ordenar==="ciclo") return (b.cicloAtual||0)-(a.cicloAtual||0);
    if (ordenar==="ticket") return b.ticketMedio-a.ticketMedio;
    if (ordenar==="ltvClub") return b.ltvClub-a.ltvClub;
    if (ordenar==="ltvTotal") return b.ltvTotal-a.ltvTotal;
    if (ordenar==="prox") return (a.proximaCob||"").localeCompare(b.proximaCob||"");
    if (ordenar==="churn") return (b.churnScore||0)-(a.churnScore||0);
    return 0;
  });

  // Totais
  const ativos = (dados||[]).filter(r=>r.status==="ativo"&&!r._debug);
  const mrrTotal = ativos.reduce((s,r)=>s+r.ticketMedio,0);
  const ltvMedio = ativos.length>0 ? ativos.reduce((s,r)=>s+r.ltvTotal,0)/ativos.length : 0;

  const Th = ({label, campo}) => (
    <th onClick={()=>setOrdenar(campo)} style={{padding:"8px 10px",fontSize:10,fontWeight:600,color:ordenar===campo?C.tealD:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",cursor:"pointer",whiteSpace:"nowrap",borderBottom:"1px solid var(--color-border-tertiary)",background:"var(--color-background-secondary)",userSelect:"none"}}>
      {label}{ordenar===campo?" ↓":""}
    </th>
  );

  return (
    <div>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16,flexWrap:"wrap"}}>
        <div style={{fontSize:14,fontWeight:600,color:"var(--color-text-primary)"}}>📈 Analytics de Assinantes</div>
        <div style={{flex:1}}/>
        <button onClick={carregar} disabled={loading}
          style={{padding:"6px 14px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",background:C.teal,color:"#fff",border:"none"}}>
          {loading?"⏳ Carregando...":"🔄 Atualizar"}
        </button>
      </div>

      {/* Cards resumo */}
      {dados&&(
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:16}}>
          {[
            ["Assinantes ativos", ativos.length, C.green, C.greenL],
            ["MRR total", "R$"+mrrTotal.toFixed(0)+"/mês", C.teal, C.tealL],
            ["LTV médio", "R$"+ltvMedio.toFixed(0), C.purple, C.purpleL],
            ["Total na base", (dados||[]).length+" clientes", "#666", "#f5f5f5"],
          ].map(([label,val,cor,bg])=>(
            <div key={label} style={{background:bg,border:"0.5px solid "+cor,borderRadius:10,padding:"10px 12px"}}>
              <div style={{fontSize:10,color:cor,marginBottom:2,textTransform:"uppercase",letterSpacing:"0.06em"}}>{label}</div>
              <div style={{fontSize:18,fontWeight:600,color:cor}}>{val}</div>
            </div>
          ))}
        </div>
      )}

      {dados&&dados.find(r=>r._debug)&&(
        <div style={{background:"#1a1a2e",borderRadius:10,padding:"14px",marginBottom:16,overflowX:"auto"}}>
          <div style={{fontSize:11,color:"#7fdbff",marginBottom:8,fontWeight:600}}>🔍 DEBUG — Todas as assinaturas (vermelho = 0 purchases)</div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:10,color:"#eee",fontFamily:"monospace"}}>
            <thead><tr style={{borderBottom:"1px solid #333"}}>
              {["Nome","Status","Cycle","Overdue_at","Canceled_at","Created","Purchases #"].map(h=><th key={h} style={{padding:"4px 8px",textAlign:"left",color:"#7fdbff"}}>{h}</th>)}
            </tr></thead>
            <tbody>
              {(dados.find(r=>r._debug)?._rawSubs||[]).map((s,i)=>(
                <tr key={i} style={{borderBottom:"0.5px solid #333",background:s.purchases===0?"rgba(200,0,0,0.3)":"transparent"}}>
                  <td style={{padding:"3px 8px",color:s.purchases===0?"#f87171":"#eee"}}>{s.nome}</td>
                  <td style={{padding:"3px 8px",color:s.status==="active"?"#4ade80":"#fbbf24"}}>{s.status}</td>
                  <td style={{padding:"3px 8px"}}>{s.cycle}</td>
                  <td style={{padding:"3px 8px",color:"#aaa",fontSize:9}}>{s.overdue_at||"—"}</td>
                  <td style={{padding:"3px 8px",color:s.canceled_at?"#f87171":"#aaa",fontSize:9}}>{s.canceled_at||"—"}</td>
                  <td style={{padding:"3px 8px",color:"#aaa"}}>{s.created_at}</td>
                  <td style={{padding:"3px 8px",fontWeight:600,color:s.purchases===0?"#f87171":"#4ade80"}}>{s.purchases}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {erro&&<div style={{background:C.coralL,border:"0.5px solid "+C.coral,borderRadius:8,padding:"10px 14px",fontSize:12,color:C.coralD,marginBottom:12}}>{erro}</div>}
      {loading&&<div style={{textAlign:"center",padding:32,color:"var(--color-text-tertiary)"}}>⏳ Buscando dados do RitsPay...</div>}

      {dados&&(
        <>
          {/* Alertas de ação */}
          {(()=>{
            const upgrades = (dados||[]).filter(r=>r.upgradeSugerido);
            const aniversarios = (dados||[]).filter(r=>r.aniversario);
            const churnAlto = (dados||[]).filter(r=>r.churnScore>=50&&r.status!=="cancelado");
            if (!upgrades.length && !aniversarios.length && !churnAlto.length) return null;
            return (
              <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
                {upgrades.length>0&&<div style={{background:C.purpleL,border:"0.5px solid "+C.purple,borderRadius:8,padding:"8px 12px",fontSize:11,color:C.purpleD}}>
                  ↑ <strong>{upgrades.length}</strong> candidatas a upgrade ({upgrades.map(r=>r.nome.split(" ")[0]).join(", ")})
                </div>}
                {aniversarios.length>0&&<div style={{background:C.tealL,border:"0.5px solid "+C.teal,borderRadius:8,padding:"8px 12px",fontSize:11,color:C.tealD}}>
                  🎂 <strong>{aniversarios.length}</strong> aniversários de assinatura em breve
                </div>}
                {churnAlto.length>0&&<div style={{background:C.coralL,border:"0.5px solid "+C.coral,borderRadius:8,padding:"8px 12px",fontSize:11,color:C.coralD}}>
                  ⚠ <strong>{churnAlto.length}</strong> com risco alto de churn
                </div>}
              </div>
            );
          })()}
          {/* Filtro por status */}
          <div style={{display:"flex",gap:6,marginBottom:10,flexWrap:"wrap"}}>
            {statusOpts.map(s=>(
              <button key={s} onClick={()=>setFiltroStatus(s)}
                style={{padding:"3px 12px",borderRadius:20,fontSize:11,cursor:"pointer",fontWeight:filtroStatus===s?500:400,
                  background:filtroStatus===s?(COR_STATUS[s]||C.teal):"var(--color-background-secondary)",
                  color:filtroStatus===s?"#fff":"var(--color-text-secondary)",
                  border:"0.5px solid "+(filtroStatus===s?(COR_STATUS[s]||C.teal):"var(--color-border-tertiary)")}}>
                {s==="todos"?"Todos":s.charAt(0).toUpperCase()+s.slice(1)} {s!=="todos"?`(${(dados||[]).filter(r=>r.status===s).length})`:`(${(dados||[]).length})`}
              </button>
            ))}
          </div>

          {/* Tabela */}
          <div style={{overflowX:"auto",borderRadius:10,border:"0.5px solid var(--color-border-tertiary)"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead>
                <tr>
                  <Th label="Nome" campo="nome"/>
                  <Th label="Plano" campo="plano"/>
                  <Th label="Ciclo" campo="ciclo"/>
                  <Th label="Ticket médio" campo="ticket"/>
                  <Th label="LTV Club" campo="ltvClub"/>
                  <Th label="LTV Total" campo="ltvTotal"/>
                  <Th label="Próx. cobr." campo="prox"/>
                  <Th label="Risco" campo="churn"/>
                  <th style={{padding:"8px 10px",fontSize:10,fontWeight:600,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",borderBottom:"1px solid var(--color-border-tertiary)",background:"var(--color-background-secondary)"}}>Status</th>
                </tr>
              </thead>
              <tbody>
                {ordenados.map((r,i)=>(
                  <tr key={r.id||i} style={{background:i%2===0?"var(--color-background-primary)":"var(--color-background-secondary)",cursor:"pointer"}}
                    onClick={()=>{ if(r.email) { const crm = null; onAbrirPerfil&&onAbrirPerfil(r.email); } }}>
                    <td style={{padding:"8px 10px",color:"var(--color-text-primary)",fontWeight:500}}>{r.nome}</td>
                    <td style={{padding:"8px 10px",color:"var(--color-text-secondary)"}}>{r.plano}</td>
                    <td style={{padding:"8px 10px",color:C.tealD,fontWeight:500}}>{r.cicloDisplay}</td>
                    <td style={{padding:"8px 10px",color:"var(--color-text-primary)"}}>R${r.ticketMedio.toFixed(0)}</td>
                    <td style={{padding:"8px 10px",color:C.tealD,fontWeight:500}}>R${r.ltvClub.toFixed(0)}</td>
                    <td style={{padding:"8px 10px",color:C.purple,fontWeight:500}}>R${r.ltvTotal.toFixed(0)}{r.gastoAvulso>0&&<span style={{fontSize:10,color:"var(--color-text-tertiary)",marginLeft:4}}>+R${r.gastoAvulso.toFixed(0)}</span>}</td>
                    <td style={{padding:"8px 10px",color:r.proximaCob!=="—"&&r.proximaCob<=new Date().toISOString().split("T")[0]?C.coralD:"var(--color-text-secondary)"}}>{r.proximaCob!=="—"?new Date(r.proximaCob+"T12:00:00").toLocaleDateString("pt-BR",{day:"2-digit",month:"2-digit"}):r.proximaCob}</td>
                    <td style={{padding:"8px 10px"}}>
                      {r.churnScore>0&&(
                        <div style={{display:"flex",alignItems:"center",gap:4}}>
                          <div style={{width:32,height:5,background:"#eee",borderRadius:3,overflow:"hidden"}}>
                            <div style={{width:r.churnScore+"%",height:"100%",background:r.churnScore>=50?C.coral:r.churnScore>=25?C.amber:C.green,borderRadius:3}}/>
                          </div>
                          <span style={{fontSize:9,color:"var(--color-text-tertiary)"}}>{r.churnScore}</span>
                        </div>
                      )}
                      {r.upgradeSugerido&&<div style={{fontSize:9,color:C.purple,marginTop:2}}>↑ {r.upgradeSugerido}</div>}
                      {r.aniversario&&<div style={{fontSize:9,color:C.tealD,marginTop:2}}>🎂 {r.aniversario.meses}m em {r.aniversario.data}</div>}
                    </td>
                    <td style={{padding:"8px 10px"}}>
                      <span style={{padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:500,background:(COR_STATUS[r.status]||"#eee")+"22",color:COR_STATUS[r.status]||"#666"}}>
                        {r.status}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
};


const NovoLeadModal = ({ onClose, onSalvo }) => {
  const [nome, setNome] = React.useState("");
  const [tel, setTel] = React.useState("");
  const [email, setEmail] = React.useState("");
  const [pedidos, setPedidos] = React.useState("0");
  const [gasto, setGasto] = React.useState("0");
  const [etapa, setEtapa] = React.useState("lead");
  const [origem, setOrigem] = React.useState("");
  const [salvando, setSalvando] = React.useState(false);
  const [ok, setOk] = React.useState(false);

  const salvar = async () => {
    if (!nome.trim()) return;
    setSalvando(true);
    const p = parseInt(pedidos)||0;
    const g = parseFloat(gasto)||0;
    const c = {
      id: "c_"+Date.now()+"_"+Math.random().toString(36).slice(2,8),
      nome: nome.trim(),
      telefone: normalizarTelefone(tel.trim()),
      email: email.trim(),
      p, gasto: g,
      etapa,
      dataCriacao: new Date().toLocaleDateString("pt-BR"),
      lista: origem ? "Lead — "+origem : "Lead manual",
      customerId: "", fora: null,
      dataPrimeiro: "", dataUltimo: "",
      datasPreenchidas: false,
      cicloMedio: 0,
      objetivo: p >= 2 ? "club" : "novo_cliente",
    };
    await dbSave(c);
    setSalvando(false);
    setOk(true);
    setTimeout(() => { onSalvo && onSalvo(); onClose(); }, 1200);
  };

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:9999}}>
      <div style={{width:420,background:"var(--color-background-primary)",borderRadius:16,padding:"24px 28px",boxShadow:"0 8px 40px rgba(0,0,0,0.15)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:20}}>
          <div style={{fontSize:20}}>➕</div>
          <div style={{flex:1,fontSize:15,fontWeight:600,color:"var(--color-text-primary)"}}>Novo Lead</div>
          <button onClick={onClose} style={{fontSize:18,cursor:"pointer",background:"none",border:"none",color:"var(--color-text-tertiary)"}}>✕</button>
        </div>
        {ok ? (
          <div style={{textAlign:"center",padding:"20px 0"}}>
            <div style={{fontSize:32,marginBottom:8}}>✅</div>
            <div style={{fontSize:14,color:C.greenD,fontWeight:500}}>Lead cadastrado!</div>
            <div style={{fontSize:12,color:"var(--color-text-tertiary)",marginTop:4}}>Aparece no Kanban e no Club</div>
          </div>
        ) : (
          <div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
              <div style={{gridColumn:"1/-1"}}>
                <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>Nome <span style={{color:C.coralD}}>*</span></div>
                <input autoFocus value={nome} onChange={e=>setNome(e.target.value)}
                  onKeyDown={e=>e.key==="Enter"&&salvar()}
                  style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none"}}
                  placeholder="Nome completo"/>
              </div>
              <div>
                <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>WhatsApp</div>
                <input value={tel} onChange={e=>setTel(e.target.value)}
                  style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none"}}
                  placeholder="11 9XXXX-XXXX"/>
              </div>
              <div>
                <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>Email</div>
                <input value={email} onChange={e=>setEmail(e.target.value)}
                  style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none"}}
                  placeholder="email@exemplo.com"/>
              </div>
              <div>
                <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>Pedidos anteriores</div>
                <input type="number" min="0" value={pedidos} onChange={e=>setPedidos(e.target.value)}
                  style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none"}}/>
              </div>
              <div>
                <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>Gasto total (R$)</div>
                <input type="number" min="0" value={gasto} onChange={e=>setGasto(e.target.value)}
                  style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none"}}/>
              </div>
              <div>
                <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>Etapa inicial</div>
                <select value={etapa} onChange={e=>setEtapa(e.target.value)}
                  style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)"}}>
                  <option value="lead">Lead</option>
                  <option value="primeiro_contato">Primeiro Contato</option>
                  <option value="em_conversa">Em Conversa</option>
                </select>
              </div>
              <div>
                <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em"}}>Origem</div>
                <select value={origem} onChange={e=>setOrigem(e.target.value)}
                  style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:13,color:"var(--color-text-primary)",background:"var(--color-background-secondary)"}}>
                  <option value="">— Selecione —</option>
                  {["Indicação","Instagram","TikTok","WhatsApp","Evento","Presencial","Outra"].map(o=><option key={o} value={o}>{o}</option>)}
                </select>
              </div>
            </div>
            {parseInt(pedidos)>=2&&(
              <div style={{background:C.tealL,border:"0.5px solid "+C.teal,borderRadius:8,padding:"8px 12px",fontSize:11,color:C.tealD,marginBottom:10}}>
                ⭐ Com {pedidos} pedidos, vai aparecer diretamente na lista do Club
              </div>
            )}
            <button onClick={salvar} disabled={!nome.trim()||salvando}
              style={{width:"100%",padding:"11px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",
                background:nome.trim()?C.teal:"#ccc",color:"#fff",border:"none",marginTop:4}}>
              {salvando?"Salvando...":"Cadastrar Lead"}
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

// ── FUNIL CLUB ─────────────────────────────────────────────────────────────
const STATUS_CLUB = [
  { id:"",            label:"Não abordado",   cor:C.purple,  emoji:"○" },
  { id:"contatado",   label:"Contatado",      cor:C.blue,    emoji:"📤" },
  { id:"respondeu",   label:"Respondeu",      cor:C.teal,    emoji:"💬" },
  { id:"interessado", label:"Interessado",    cor:C.green,   emoji:"🔥" },
  { id:"link_enviado",label:"Link enviado",   cor:C.amber,   emoji:"🔗" },
  { id:"fechou",      label:"Fechou ✓",       cor:C.green,   emoji:"🏆" },
  { id:"nao_agora",   label:"Não agora",      cor:C.coral,   emoji:"⏸" },
  { id:"follow_up",   label:"Follow-up",      cor:C.amber,   emoji:"🔔" },
  { id:"perdido",     label:"Perdido",        cor:C.coral,   emoji:"✗"  },
];

const OBJECTIONS = [
  "Caro","Não quer fidelidade","Não entendeu os planos","Não consome todo mês",
  "Medo de enjoar","Não mora em SP","Quer pensar","Prefere compra avulsa",
  "Problema com frete","Quer montar caixa diferente","Sem resposta","Outro"
];

const PLANOS = ["Trimestral","Semestral","Anual"];


const MOTIVOS_INATIVIDADE = [
  "Preço ficou alto",
  "Encontrou outro produto",
  "Mudou a dieta / protocolo",
  "Esqueceu da Laricas",
  "Teve problema com pedido ou entrega",
  "Estava viajando / mudança de rotina",
  "Questão financeira",
  "Não gostou de algum produto",
  "Não compra mais online",
  "Outro",
];

const MOTIVOS_NAO_PODE = [
  "Orçamento apertado no momento",
  "Acabou de fazer um pedido grande",
  "Está em dieta restritiva / protocolo",
  "Vai viajar / mudança de rotina",
  "Quer esperar o próximo salário",
  "Está testando outros produtos",
  "Momento pessoal difícil",
  "Muito caro",
  "Site é igual ou melhor",
  "Outro",
];


// Ao enviar um script, avança automaticamente o próximo follow-up
// sem precisar revisitar o perfil no final do dia
const SCRIPT_AUTO_FOLLOWUP = {
  "recorrente":   { dias: 2,  label: "48h — aguardando resposta" },
  "ticket_alto":  { dias: 2,  label: "48h — aguardando resposta" },
  "kit":          { dias: 2,  label: "48h — aguardando resposta" },
  "followup_48h": { dias: 5,  label: "Aguardando — F2 em 5 dias" },
  "followup_2":   { dias: 7,  label: "Aguardando — F3 em 7 dias" },
  "followup_3":   { dias: 10, label: "Aguardando — F4 em 10 dias" },
  "followup_4":   { dias: 14, label: "Última checagem em 14 dias" },
  "planos":       { dias: 2,  label: "Aguardando resposta — FP1 em 2 dias" },
  "fechamento":   { dias: 2,  label: "Aguardando resposta — FP2 em 2 dias" },
  "link_direto":  { dias: 2,  label: "Aguardando resposta — FP2 em 2 dias" },
  "fp1":          { dias: 2,  label: "Aguardando resposta — 2 dias" },
  "fp2":          { dias: 2,  label: "Aguardando resposta — 2 dias" },
  "obj_caro":     { dias: 2,  label: "Aguardando resposta — 2 dias" },
  "obj_enjoar":   { dias: 2,  label: "Aguardando resposta — 2 dias" },
  "obj_fidelidade":{ dias: 2, label: "Aguardando resposta — 2 dias" },
  "obj_pensar":   { dias: 3,  label: "Aguardando resposta — 3 dias" },
  "obj_nao_pode_agora": { dias: 3, label: "Aguardando motivo — 3 dias" },
  "condicao_especial":  { dias: 2, label: "Gatilho final enviado — 2 dias" },
  "followup_7d":  { dias: 3,  label: "Follow-up final — 3 dias" },
  "reativacao":   { dias: 3,  label: "Aguardando resposta — reativação em 3 dias" },
};
// Padrão para scripts não mapeados — toda cópia gera follow-up de 2 dias
const SCRIPT_AUTO_FOLLOWUP_DEFAULT = { dias: 2, label: "Aguardando resposta — 2 dias" };

const FOLLOW_UP_DAYS = {
  "contatado": 2,
  "respondeu": 1,
  "interessado": 1,
  "link_enviado": 2,
  "nao_agora": 30,
  "follow_up": 3,
};
const FOLLOW_UP_LABELS = {
  "contatado": "48h — Follow-up F",
  "respondeu": "Amanhã — está quente",
  "interessado": "Amanhã — enviar link",
  "link_enviado": "2 dias — aguardar decisão",
  "nao_agora": "30 dias — aguardar momento certo",
  "follow_up": "3 dias",
};


// Árvore de decisão: script enviado → respostas possíveis → próximo script + status
const ARVORE = {
  "recorrente": [
    { label:"Ficou curiosa / quer saber mais", emoji:"🤩", proximoScript:"planos", novoStatus:"respondeu" },
    { label:"Já pediu o link direto",          emoji:"🔗", proximoScript:"link_direto", novoStatus:"interessado" },
    { label:"Disse que é caro",                emoji:"💸", proximoScript:"obj_caro", novoStatus:"respondeu" },
    { label:"Não quer fidelidade",             emoji:"🔓", proximoScript:"obj_fidelidade", novoStatus:"respondeu" },
    { label:"Quer pensar",                     emoji:"🤔", proximoScript:"obj_pensar", novoStatus:"respondeu" },
    { label:"Não respondeu",                   emoji:"🔇", proximoScript:"followup_48h", novoStatus:"contatado" },
  ],
  "ticket_alto": [
    { label:"Ficou curiosa / quer saber mais", emoji:"🤩", proximoScript:"planos", novoStatus:"respondeu" },
    { label:"Já pediu o link direto",          emoji:"🔗", proximoScript:"link_direto", novoStatus:"interessado" },
    { label:"Disse que é caro",                emoji:"💸", proximoScript:"obj_caro", novoStatus:"respondeu" },
    { label:"Não quer fidelidade",             emoji:"🔓", proximoScript:"obj_fidelidade", novoStatus:"respondeu" },
    { label:"Quer pensar",                     emoji:"🤔", proximoScript:"obj_pensar", novoStatus:"respondeu" },
    { label:"Não respondeu",                   emoji:"🔇", proximoScript:"followup_48h", novoStatus:"contatado" },
  ],
  "kit": [
    { label:"Ficou curiosa / quer saber mais", emoji:"🤩", proximoScript:"planos", novoStatus:"respondeu" },
    { label:"Já pediu o link direto",          emoji:"🔗", proximoScript:"link_direto", novoStatus:"interessado" },
    { label:"Disse que é caro",                emoji:"💸", proximoScript:"obj_caro", novoStatus:"respondeu" },
    { label:"Quer pensar",                     emoji:"🤔", proximoScript:"obj_pensar", novoStatus:"respondeu" },
    { label:"Não respondeu",                   emoji:"🔇", proximoScript:"followup_48h", novoStatus:"contatado" },
  ],
  "planos": [
    { label:"Animada / quer o link",           emoji:"🔥", proximoScript:"fechamento", novoStatus:"interessado" },
    { label:"Medo de enjoar",                  emoji:"😰", proximoScript:"obj_enjoar", novoStatus:"respondeu" },
    { label:"Disse que é caro",                emoji:"💸", proximoScript:"obj_caro", novoStatus:"respondeu" },
    { label:"Não quer fidelidade",             emoji:"🔓", proximoScript:"obj_fidelidade", novoStatus:"respondeu" },
    { label:"Quer pensar",                     emoji:"🤔", proximoScript:"obj_pensar", novoStatus:"respondeu" },
  ],
  "fechamento": [
    { label:"Sim! Quero o link",               emoji:"✅", proximoScript:null, novoStatus:"link_enviado" },
    { label:"Ainda com dúvida de preço",       emoji:"💸", proximoScript:"obj_caro", novoStatus:"interessado" },
    { label:"Quer mais tempo",                 emoji:"🤔", proximoScript:"obj_pensar", novoStatus:"interessado" },
    { label:"Desistiu",                        emoji:"❌", proximoScript:null, novoStatus:"perdido" },
  ],
  "link_direto": [
    { label:"Confirmou o plano — enviar link", emoji:"✅", proximoScript:null, novoStatus:"link_enviado" },
    { label:"Quer outro plano",                emoji:"🔄", proximoScript:"planos", novoStatus:"interessado" },
    { label:"Desistiu",                        emoji:"❌", proximoScript:null, novoStatus:"perdido" },
  ],
  "obj_caro": [
    { label:"Entendeu e quer continuar",       emoji:"😊", proximoScript:"planos", novoStatus:"respondeu" },
    { label:"Ainda resistente",                emoji:"😐", proximoScript:"followup_7d", novoStatus:"follow_up" },
    { label:"Não quer mais",                   emoji:"❌", proximoScript:null, novoStatus:"perdido" },
  ],
  "obj_fidelidade": [
    { label:"Aceitou o trimestral",            emoji:"✅", proximoScript:"fechamento", novoStatus:"interessado" },
    { label:"Ainda indecisa",                  emoji:"🤔", proximoScript:"followup_7d", novoStatus:"follow_up" },
    { label:"Não quer mais",                   emoji:"❌", proximoScript:null, novoStatus:"perdido" },
  ],
  "obj_enjoar": [
    { label:"Ficou tranquila / quer continuar",emoji:"😊", proximoScript:"fechamento", novoStatus:"interessado" },
    { label:"Ainda com medo",                  emoji:"😰", proximoScript:"followup_7d", novoStatus:"follow_up" },
    { label:"Não quer mais",                   emoji:"❌", proximoScript:null, novoStatus:"perdido" },
  ],
  "obj_pensar": [
    { label:"Voltou interessada",              emoji:"🔥", proximoScript:"fechamento", novoStatus:"interessado" },
    { label:"Quer mais tempo",                 emoji:"⏳", proximoScript:"condicao_especial", novoStatus:"follow_up" },
    { label:"Não quer mais",                   emoji:"❌", proximoScript:null, novoStatus:"perdido" },
  ],
  "followup_48h": [
    { label:"Respondeu com interesse",         emoji:"🤩", proximoScript:"planos", novoStatus:"respondeu" },
    { label:"Disse que é caro",                emoji:"💸", proximoScript:"obj_caro", novoStatus:"respondeu" },
    { label:"Não pode agora",                  emoji:"⏳", proximoScript:"obj_nao_pode_agora", novoStatus:"respondeu" },
    { label:"Quer pensar",                     emoji:"🤔", proximoScript:"obj_pensar", novoStatus:"respondeu" },
    { label:"Não respondeu — F2",              emoji:"🔇", proximoScript:"followup_2", novoStatus:"follow_up" },
  ],
  "followup_2": [
    { label:"Respondeu com interesse",         emoji:"🤩", proximoScript:"planos", novoStatus:"respondeu" },
    { label:"Não pode agora",                  emoji:"⏳", proximoScript:"obj_nao_pode_agora", novoStatus:"respondeu" },
    { label:"Quer pensar",                     emoji:"🤔", proximoScript:"obj_pensar", novoStatus:"respondeu" },
    { label:"Não respondeu — F3",              emoji:"🔇", proximoScript:"followup_3", novoStatus:"follow_up" },
  ],
  "followup_3": [
    { label:"Respondeu — retomar conversa",    emoji:"🤩", proximoScript:"planos", novoStatus:"respondeu" },
    { label:"Não pode agora",                  emoji:"⏳", proximoScript:"obj_nao_pode_agora", novoStatus:"respondeu" },
    { label:"Não respondeu — F4 (última)",     emoji:"🔇", proximoScript:"followup_4", novoStatus:"follow_up" },
  ],
  "followup_4": [
    { label:"Respondeu — retomar conversa",    emoji:"🤩", proximoScript:"planos", novoStatus:"respondeu" },
    { label:"Não pode agora",                  emoji:"⏳", proximoScript:"obj_nao_pode_agora", novoStatus:"respondeu" },
    { label:"Não respondeu — encerrar",        emoji:"🔇", proximoScript:null, novoStatus:"perdido" },
  ],
  "followup_7d": [
    { label:"Respondeu interessada",           emoji:"🔥", proximoScript:"planos", novoStatus:"respondeu" },
    { label:"Ainda em dúvida",                 emoji:"🤔", proximoScript:"condicao_especial", novoStatus:"follow_up" },
    { label:"Não respondeu / ignorou",         emoji:"🔇", proximoScript:null, novoStatus:"perdido" },
  ],
  "condicao_especial": [
    { label:"Aceitou! Quer o link",            emoji:"🏆", proximoScript:null, novoStatus:"link_enviado" },
    { label:"Mesmo assim não quer",            emoji:"❌", proximoScript:null, novoStatus:"perdido" },
  ],
  "obj_nao_pode_agora": [
    { label:"Motivo financeiro — voltar em 30d", emoji:"💰", proximoScript:null, novoStatus:"nao_agora", followUpDias:30 },
    { label:"Mudança de rotina — voltar em 15d", emoji:"🔄", proximoScript:null, novoStatus:"nao_agora", followUpDias:15 },
    { label:"Outro motivo — definir data",       emoji:"📅", proximoScript:null, novoStatus:"nao_agora", followUpDias:null },
  ],
  "fp1": [
    { label:"Respondeu — tem interesse",         emoji:"🔥", proximoScript:"fechamento", novoStatus:"interessado" },
    { label:"Disse que é caro",                  emoji:"💸", proximoScript:"obj_caro", novoStatus:"respondeu" },
    { label:"Não pode agora",                    emoji:"⏳", proximoScript:"obj_nao_pode_agora", novoStatus:"respondeu" },
    { label:"Não respondeu — encerrar",          emoji:"🔇", proximoScript:null, novoStatus:"perdido" },
  ],
  "fp2": [
    { label:"Respondeu — enviar link",           emoji:"✅", proximoScript:null, novoStatus:"link_enviado" },
    { label:"Tem dúvida — retomar",              emoji:"🤔", proximoScript:"planos", novoStatus:"respondeu" },
    { label:"Não pode agora",                    emoji:"⏳", proximoScript:"obj_nao_pode_agora", novoStatus:"respondeu" },
    { label:"Não respondeu — encerrar",          emoji:"🔇", proximoScript:null, novoStatus:"perdido" },
  ],
};

const SCRIPTS_CLUB = [
  {
    id:"recorrente", label:"A — Cliente recorrente", tag:"recorrente",
    perfil:"3+ pedidos com ciclo regular",
    copy:`Oi [Nome]! 😊 Aqui é a [Operador] da Laricas.

Vi que você já é cliente de carteirinha [N° pedidos]x — e isso me deixou pensando em você.

Queria te apresentar o Laricas Club: é basicamente garantir que você sempre tenha aquele momento de prazer sem precisar pensar nisso. Sem culpa, sem ficar adiando.

Você escolhe seus produtos favoritos, recebe todo mês com desconto e frete grátis para SP.

Posso te mandar as opções de planos?`,
  },
  {
    id:"ticket_alto", label:"B — Ticket alto", tag:"ticket_alto",
    perfil:"Gasto acima de R$300 no último pedido",
    copy:`Oi [Nome]! 😊 Aqui é a [Operador] da Laricas.

Vi seu último pedido e pensei: você já decidiu que merece esse momento — só ainda não tem isso garantido todo mês.

O Laricas Club resolve exatamente isso: sua rotina de prazer já vem pronta, sem precisar lembrar de pedir, com desconto e frete grátis.

Quer que eu te explique rapidinho os planos?`,
  },
  {
    id:"kit", label:"C — Comprou kit", tag:"kit",
    perfil:"Comprou kit ou seleção ampla de produtos",
    copy:`Oi [Nome]! 😊 Aqui é a [Operador] da Laricas.

Vi que você já experimentou um pouco de tudo com a gente — e isso me fez pensar que talvez você já saiba o que gosta.

O Laricas Club é exatamente isso: sua seleção favorita, garantida todo mês, sem precisar escolher de novo toda vez.

Posso te mandar as opções?`,
  },
  {
    id:"planos", label:"D — Explicação dos planos", tag:"planos",
    perfil:"Após interesse confirmado",
    copy:`O Laricas Club existe para uma coisa simples: garantir que você sempre tenha o seu momento de prazer, sem culpa e sem depender de lembrar.

Você escolhe de 7 a 15 doces por mês, monta sua caixa do jeito que quiser, e pode trocar os sabores todo mês se quiser variar — com desconto em relação ao site e frete grátis para SP.

Hoje temos 3 opções de plano:

*Trimestral:* para quem quer testar com menor compromisso.
*Semestral:* bom meio-termo para manter a rotina.
*Anual:* melhor condição, com maior desconto.

Pelo seu perfil, eu recomendaria começar pelo plano [plano recomendado], porque [motivo]. Faz sentido pra você?`,
  },
  {
    id:"fechamento", label:"E — Fechamento", tag:"fechamento",
    perfil:"Cliente quente, pronta para fechar",
    copy:`Acho que para você faz todo sentido — você já decidiu que merece esse momento, agora é só deixar ele garantido todo mês, sem esforço.

Quer que eu te mande o link direto para montar sua caixa?`,
  },
  {
    id:"link_direto", label:"M — Link direto (cliente pediu)", tag:"fechamento",
    perfil:"Cliente pulou etapas e já pediu o link",
    copy:`Show, [Nome]! 😄

Vou te mandar o link certinho. Só confirma rapidinho: pode ser o plano [plano recomendado]?

Assim que confirmar, te mando o link na hora!`,
  },
  {
    id:"followup_48h", label:"F — Follow-up 48h", tag:"follow_up",
    perfil:"Sem resposta após 48h",
    copy:`Oi [Nome], passando só para não deixar pelo caminho.

Pensei em você porque seu perfil combina muito com o Club — ter aquele momento de prazer garantido, sem culpa e sem precisar pensar.

Quer que eu te explique rapidinho como funciona? 😊`,
  },
  {
    id:"followup_2", label:"F2 — Follow-up 2 (+5 dias)", tag:"follow_up",
    perfil:"Não respondeu o F — apresentar o Club de outro ângulo",
    copy:`Oi [Nome]! 😊

Deixa eu tentar de outro jeito.

O Laricas Club é basicamente isso: você escolhe seus doces favoritos, de 7 a 15 por mês, e recebe em casa com desconto e frete grátis — sem precisar lembrar de pedir.

Tem gente que chama de "meu estoque mensal de prazer sem culpa" 😄

Faz sentido pra você?`,
  },
  {
    id:"followup_3", label:"F3 — Follow-up 3 (+7 dias)", tag:"follow_up",
    perfil:"Não respondeu o F2 — transparência total",
    copy:`Oi [Nome], tô mandando essa mensagem porque genuinamente acho que o Club faz sentido pra você — mas entendo se o momento não for esse.

Se não for a hora, tudo bem mesmo. Só me fala e não te incomodo mais 😊`,
  },
  {
    id:"followup_4", label:"F4 — Follow-up 4 (+10 dias)", tag:"follow_up",
    perfil:"Não respondeu o F3 — identidade e pertencimento",
    copy:`Oi [Nome]! Uma última mensagem, prometo 😄

Pensei em você porque poucas pessoas entendem o que a Laricas representa — não é só doce, é aquele momento de prazer que não sabota nada.

Queria só saber: isso ainda faz sentido na sua rotina?`,
  },
  {
    id:"followup_7d", label:"G — Follow-up final", tag:"follow_up",
    perfil:"Última tentativa — condição especial reservada",
    copy:`Última mensagem sobre isso, prometo! 😄

Só queria deixar a porta aberta: quando você quiser garantir seu momento de prazer sem culpa todo mês, com desconto e frete grátis, me chama.

Posso deixar reservada uma condição especial pra você?`,
  },
  {
    id:"condicao_especial", label:"L — Condição especial (gatilho final)", tag:"fechamento",
    perfil:"Clientes em dúvida após follow-up — última oferta antes de perder",
    copy:`[Nome], não costumo fazer isso, mas queria te dar uma condição especial 😊

O plano Anual já tem 20% de desconto em relação ao site. Para você, vou garantir mais 10% — 30% de desconto no primeiro ano.

É a forma mais fácil de garantir seu momento de prazer sem culpa, com a melhor condição que tenho.

Quer que eu te mande o link?`,
  },
  {
    id:"obj_caro", label:"H — Objeção: caro", tag:"objecao",
    perfil:"Disse que é caro",
    copy:`Entendo totalmente!

Pensa assim: você já se permite esse prazer de vez em quando — o Club só garante que isso aconteça sem você precisar decidir toda vez, e ainda com desconto e frete grátis.

Para quem já tem esse hábito, normalmente sai mais em conta.

Faz sentido pra você nesse formato?`,
  },
  {
    id:"obj_enjoar", label:"I — Objeção: medo de enjoar", tag:"objecao",
    perfil:"Com medo de enjoar",
    copy:`Total, entendo! Por isso a ideia não é repetir sempre a mesma coisa.

Você escolhe de 7 a 15 doces por mês e pode trocar os sabores quando quiser — seu momento de prazer continua sendo surpresa, não repetição 😄

Isso muda sua visão sobre o Club?`,
  },
  {
    id:"obj_fidelidade", label:"J — Objeção: não quer fidelidade", tag:"objecao",
    perfil:"Não quer compromisso longo",
    copy:`Entendo! Nesse caso, o trimestral é o melhor jeito de começar.

Você testa esse momento de prazer garantido sem se comprometer por muito tempo 😊

Quer começar assim, sem compromisso longo?`,
  },
  {
    id:"obj_pensar", label:"K — Objeção: quer pensar", tag:"objecao",
    perfil:"Disse que vai pensar",
    copy:`Claro, sem pressão! 😊

Só uma reflexão: você já decidiu que se permite esse prazer de vez em quando. O Club só transforma isso em rotina garantida, sem culpa e sem esforço.

Posso te chamar de novo em alguns dias?`,
  },
  {
    id:"reativacao", label:"R — Reativação (inativa há muito tempo)", tag:"reativacao",
    perfil:"3+ pedidos no passado mas sem comprar há 90+ dias — entender o motivo antes de oferecer Club",
    copy:`Oi [Nome]! 😊 Aqui é a [Operador] da Laricas.

Vi que faz um tempinho que você não pede nada com a gente, e fiquei curiosa.

Tudo bem por aí? Teve alguma coisa que te afastou ou foi só correria mesmo?

Pergunto porque você era uma das clientes que mais gostava dos nossos produtos, e queria entender se tem algo que possa melhorar 😊`,
  },
  {
    id:"fp1", label:"FP1 — Não respondeu os planos (48h)", tag:"follow_up",
    perfil:"Recebeu os planos mas não respondeu",
    copy:`Oi [Nome]! 😊

Passando pra ver se ficou alguma dúvida sobre os planos.

Qual deles pareceu fazer mais sentido pra você?`,
  },
  {
    id:"fp2", label:"FP2 — Não respondeu o fechamento (48h)", tag:"follow_up",
    perfil:"Estava prestes a fechar mas sumiu",
    copy:`Oi [Nome]! 😊

Vi que não deu tempo de responder — sem problema.

Só queria saber: ficou alguma dúvida antes de montar sua caixa?`,
  },
  {
    id:"obj_nao_pode_agora", label:"N — Objeção: não pode agora", tag:"objecao",
    perfil:"Tem interesse genuíno mas não pode assinar no momento",
    copy:`Faz todo sentido, [Nome]! Sem pressão mesmo 😊

Só pra eu entender melhor — é mais uma questão de momento financeiro agora, ou tem alguma outra razão?`,
  },
];



// Score de calor Club (0-100)
const calcScoreClub = (c) => {
  const p = c.p || 0;
  const ciclo = c.cicloMedio || 999;
  const gasto = c.gasto || 0;
  const fora = c.fora || false;
  const hoje = new Date();
  const dtU = c.dataUltimo ? new Date(c.dataUltimo+"T12:00:00") : null;
  const diasUlt = dtU ? Math.round((hoje-dtU)/86400000) : 999;
  const inativa = diasUlt > 90;
  const muitoInativa = diasUlt > 180;
  // Ciclo válido só se cliente comprou recentemente (dentro de 3 ciclos)
  const cicloValido = ciclo < 120 && diasUlt < ciclo * 3;
  const diasParaProxima = dtU && cicloValido ? ciclo - diasUlt : null;

  let score = 0;
  // Pedidos (max 30)
  if (p >= 5) score += 30;
  else if (p >= 3) score += 22;
  else if (p === 2) score += 12;
  else score += 4;
  // Ciclo histórico (max 22)
  if (ciclo <= 20) score += 22;
  else if (ciclo <= 30) score += 18;
  else if (ciclo <= 45) score += 12;
  else if (ciclo <= 60) score += 6;
  else if (ciclo <= 90) score += 2;
  // Janela de compra — só conta se ativa (max 20)
  if (!inativa && diasParaProxima !== null) {
    if (diasParaProxima >= -3 && diasParaProxima <= 5) score += 20;
    else if (diasParaProxima <= 10) score += 14;
    else if (diasParaProxima <= 14) score += 8;
  }
  // Gasto total (max 15)
  if (gasto > 1500) score += 15;
  else if (gasto > 800) score += 10;
  else if (gasto > 400) score += 6;
  else if (gasto > 150) score += 3;
  // Fora SP (max 10)
  if (fora) score += 10;

  // Penalidade por inatividade — FATOR DETERMINANTE
  if (muitoInativa) score = Math.round(score * 0.05);        // >180d: score quase zero
  else if (diasUlt > 120) score = Math.round(score * 0.15);  // 120-180d: muito baixo
  else if (inativa) score = Math.round(score * 0.30);        // 90-120d: baixo
  else if (diasUlt > 60) score = Math.round(score * 0.60);   // 60-90d: moderado
  else if (diasUlt > 45) score = Math.round(score * 0.80);   // 45-60d: leve penalidade

  return { score: Math.min(100, score), diasUlt, diasParaProxima, inativa, muitoInativa };
};

const sugerirScript = (c) => {
  const gasto = c.gasto || 0;
  const p = c.p || 0;
  if (c.statusClub === "respondeu" || c.statusClub === "interessado") return "planos";
  if (c.statusClub === "link_enviado") return "fechamento";
  if (c.statusClub === "contatado") return "followup_48h";
  if (c.statusClub === "nao_agora" || c.statusClub === "follow_up") return "followup_7d";
  if (gasto / Math.max(p, 1) > 200) return "ticket_alto";
  if (p >= 3) return "recorrente";
  return "recorrente";
};

const razaoScript = (c) => {
  const gasto = c.gasto || 0;
  const p = c.p || 0;
  const ticketMedio = Math.round(gasto / Math.max(p, 1));
  if (c.statusClub === "respondeu") return "Respondeu — apresentar planos agora";
  if (c.statusClub === "interessado") return "Está interessada — fechar com planos";
  if (c.statusClub === "link_enviado") return "Link enviado — push final para fechar";
  if (c.statusClub === "contatado") return "Sem resposta em 48h — follow-up leve";
  if (c.statusClub === "nao_agora" || c.statusClub === "follow_up") return "Estava fria — última tentativa";
  if (gasto / Math.max(p, 1) > 200) return "Ticket médio R$"+ticketMedio+" — abordagem de volume";
  if (p >= 3) return p+"p · ciclo "+(c.cicloMedio||"?")+"d — hábito formado";
  return p+" pedidos · potencial de conversão";
};

const sugerirPlano = (c) => {
  const ciclo = c.cicloMedio || 999;
  const gasto = c.gasto || 0;
  const p = c.p || 0;
  if (p >= 6 || gasto > 1500 || ciclo <= 25) return "Anual";
  if (p >= 4 || gasto > 700 || ciclo <= 40) return "Semestral";
  return "Trimestral";
};

const addDays = (n) => {
  const d = new Date();
  d.setDate(d.getDate() + n);
  return d.toISOString().split("T")[0];
};

const ScoreBar = ({ score }) => {
  const cor = score >= 70 ? C.green : score >= 45 ? C.amber : C.coral;
  const corD = score >= 70 ? C.greenD : score >= 45 ? C.amberD : C.coralD;
  const corL = score >= 70 ? C.greenL : score >= 45 ? C.amberL : C.coralL;
  return (
    <div style={{ display:"flex",alignItems:"center",gap:6 }}>
      <div style={{ flex:1,height:6,background:"var(--color-border-tertiary)",borderRadius:3,overflow:"hidden" }}>
        <div style={{ width:score+"%",height:"100%",background:cor,borderRadius:3,transition:"width 0.3s" }}/>
      </div>
      <span style={{ fontSize:11,fontWeight:600,color:corD,background:corL,padding:"1px 6px",borderRadius:10,minWidth:32,textAlign:"center" }}>{score}</span>
    </div>
  );
};

const FunilClub = ({ onAbrirPerfil, onUrgencia }) => {
  const [clientes, setClientes] = useState([]);
  const [clientesDash, setClientesDash] = useState([]); // inclui fechou/perdido — só para estatísticas
  const [todosParaBusca, setTodosParaBusca] = useState([]); // todos os clientes para busca global
  const [loading, setLoading] = useState(true);
  const [sel, setSel] = useState(null);
  const [aba, setAba] = useState("lista");
  const [filtroStatus, setFiltroStatus] = useState("hoje");
  const [busca, setBusca] = useState("");
  const [sorts, setSorts] = useState([]);
  const [selecionadosLote, setSelecionadosLote] = useState(new Set());
  const [modoLote, setModoLote] = useState(false);
  const [showRitsSync, setShowRitsSync] = useState(false);
  const [ritsCount, setRitsCount] = React.useState(null); // contagem real de ativos do RitsPay

  // Busca contagem real de assinantes ativos do RitsPay ao montar
  React.useEffect(() => {
    const token = ritspayGetToken();
    if (!token) return;
    const cfg = ritspayLoadCfg();
    const tenant = cfg.tenantId || "TEN-1G57I7LIVD8K0F8M";
    ritspayFetch(`/sales/${tenant}/subscriptions?page=1`, token)
      .then(resp => {
        const items = Array.isArray(resp?.data) ? resp.data : [];
        // Meta: contar apenas ativos e pausados com pelo menos 1 purchase (excluir quem nunca pagou)
        const ativosEPausados = items.filter(s => {
          const st = (s.status||"").toLowerCase();
          const statusOk = st === "active" || st === "paused" || st === "suspended";
          // Será validado com purchases após buscar todas as páginas
          return statusOk && !s.overdue_at; // estimativa rápida na primeira página
        }).length;
        const total = resp?.meta?.total ?? null;
        // Se a primeira página tem menos que o total, precisamos do total real — usamos meta.total como base
        // mas subtraindo cancelados/atrasados proporcionalmente não é preciso, então buscamos todas as páginas
        if (total && total <= items.length) {
          setRitsCount(ativosEPausados);
        } else if (total) {
          // Busca todas as páginas para contar com precisão
          (async () => {
            let todos = [...items];
            let pagina = 2;
            while (todos.length < total && pagina <= 20) {
              try {
                const r2 = await ritspayFetch(`/sales/${tenant}/subscriptions?page=${pagina}`, token);
                const its = Array.isArray(r2?.data) ? r2.data : [];
                if (!its.length) break;
                todos = [...todos, ...its];
                pagina++;
              } catch(e) { break; }
            }
            const count = todos.filter(s => {
              const st = (s.status||"").toLowerCase();
              return st === "active" || st === "paused" || st === "suspended";
            }).length;
            setRitsCount(count);
          })();
        }
      })
      .catch(() => {});
  }, []);
  const toggleSort = (campo) => {
    setSorts(prev => {
      const existing = prev.find(s=>s.campo===campo);
      if (!existing) return [...prev, {campo, dir:"desc"}];
      if (existing.dir==="desc") return prev.map(s=>s.campo===campo?{campo,dir:"asc"}:s);
      return prev.filter(s=>s.campo!==campo); // remove on third click
    });
  };
  const [salvando, setSalvando] = useState(false);
  const [ok, setOk] = useState("");
  const [scriptSel, setScriptSel] = useState(null);
  const [campos, setCampos] = useState({});
  const [copiadoId, setCopiadoId] = useState("");
  const [followUpProposto, setFollowUpProposto] = useState(null);
  const [metaDiaria] = useState(12); // abordagens por dia
  const hoje = new Date().toISOString().split("T")[0];

  useEffect(() => {
    dbGetAll().then(lista => {
      const addScore = c => { const sc=calcScoreClub(c); return {...c,_score:sc.score,_diasUlt:sc.diasUlt,_diasParaProxima:sc.diasParaProxima,_inativa:sc.inativa,_muitoInativa:sc.muitoInativa}; };
      // Todos com score para busca global
      const todosSc = lista.map(c => fixCliente(addScore(c)));
      setTodosParaBusca(todosSc);
      const comScore = todosSc.filter(c => (c.p||0) >= 2);
      // Lista operacional: só candidatos ativos (exclui fechou/perdido/assinantes)
      const candidatos = comScore.filter(c =>
        c.etapa !== "experiencia" &&
        c.etapa !== "encerrado" &&
        c.statusClub !== "fechou" &&
        c.statusClub !== "perdido"
      ).sort((a,b) => b._score - a._score);
      setClientes(candidatos);
      // Emitir contagem de urgência para badge no tab
      if (onUrgencia) {
        const hoje0 = new Date().toISOString().split("T")[0];
        const urgente = candidatos.filter(c =>
          c.statusClub==="interessado" || c.statusClub==="respondeu" ||
          (c.proximoFollowup && c.proximoFollowup <= hoje0)
        ).length;
        onUrgencia(urgente);
      }
      // Lista para Dash: inclui TODOS que já entraram no funil (mesmo fechou/perdido/experiencia)
      const paraDash = comScore.filter(c => c.statusClub || c.etapa === "experiencia");
      setClientesDash(paraDash);
      setLoading(false);
    });
  }, []);

  const saveCliente = async (atualizado) => {
    setSalvando(true);
    await dbSave(atualizado);
    const sc0=calcScoreClub(atualizado);
    const comScore0 = {...atualizado,_score:sc0.score,_diasUlt:sc0.diasUlt,_diasParaProxima:sc0.diasParaProxima,_inativa:sc0.inativa,_muitoInativa:sc0.muitoInativa};
    setClientes(prev => {
      // Se virou fechou/perdido/experiencia, remove da lista operacional; senão atualiza/mantém
      if (comScore0.statusClub==="fechou" || comScore0.statusClub==="perdido" || comScore0.etapa==="experiencia") {
        return prev.filter(c=>c.id!==atualizado.id);
      }
      return prev.map(c => c.id!==atualizado.id ? c : comScore0);
    });
    setClientesDash(prev => {
      const existe = prev.find(c=>c.id===atualizado.id);
      if (existe) return prev.map(c=>c.id===atualizado.id?comScore0:c);
      // Se agora tem statusClub ou virou experiencia, inclui no Dash
      if (comScore0.statusClub || comScore0.etapa==="experiencia") return [...prev, comScore0];
      return prev;
    });
    const scSel=calcScoreClub(atualizado); setSel({...atualizado,_score:scSel.score,_diasUlt:scSel.diasUlt,_diasParaProxima:scSel.diasParaProxima,_inativa:scSel.inativa,_muitoInativa:scSel.muitoInativa});
    setOk("Salvo!");
    setTimeout(() => setOk(""), 1500);
    setSalvando(false);
  };

  const atualizarStatus = async (c, novoStatus) => {
    const followUpDias = FOLLOW_UP_DAYS[novoStatus];
    let followUpData = followUpDias ? addDays(followUpDias) : null;
    const tentativas = novoStatus === "contatado"
      ? (c.tentativasClub||0) + 1
      : (c.tentativasClub||0);

    // Conversão: mover para Experiência no Kanban + follow-up alinhado ao onboarding (R1, até 3 dias)
    let dadosConversao = {};
    if (novoStatus === "fechou") {
      const hist = (c.historicoEtapas || []).slice(-9);
      hist.push({ etapa: c.etapa, data: new Date().toLocaleDateString("pt-BR"), hora: new Date().toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"}) });
      followUpData = addDays(3); // R1 — Onboarding: até 3 dias após início
      dadosConversao = {
        etapa: "experiencia",
        historicoEtapas: hist,
        tipoAssinatura: c.planoFechado || c.planoRec || "trimestral",
        dataInicioAssinatura: c.dataInicioAssinatura || hoje,
        valorMensal: c.valorMensalClub || "",
      };
    }

    const atualizado = {
      ...c,
      statusClub: novoStatus,
      tentativasClub: tentativas,
      ...(novoStatus === "contatado" && !c.dataAbordagem ? { dataAbordagem: hoje } : {}),
      dataUltimoContato: hoje,
      ...(followUpData ? { proximoFollowup: followUpData } : {}),
      ...(novoStatus === "fechou" ? { dataConversao: hoje } : {}),
      ...dadosConversao,
    };
    await saveCliente(atualizado);
    // Atualizar script sugerido automaticamente
    setScriptSel(sugerirScript(atualizado));
    // Propor follow-up com opção de editar
    if (followUpData) {
      setFollowUpProposto({
        data: followUpData,
        label: novoStatus==="fechou" ? "Onboarding (R1) — até 3 dias após início" : (FOLLOW_UP_LABELS[novoStatus]||""),
        clienteId: c.id,
        status: novoStatus,
      });
      setTimeout(() => setFollowUpProposto(null), 8000);
    }
  };

  const personalizarScript = (copy, c) => {
    const nome = (c.nome||"").split(" ")[0] || "cliente";
    const operador = (c.responsavel||"Lucas").split(" ")[0];
    const artigo = c.responsavel ? (detectarGenero(operador)==="f"?"a":"o") : "o";
    const planoRec = c.planoRec || sugerirPlano(c);
    return copy
      .replace(/\[Nome\]/g, nome)
      .replace(/\[Operador\]/g, operador)
      .replace(/\[N° pedidos\]/g, c.p||0)
      .replace(/\[plano recomendado\]/g, planoRec)
      .replace(/Lucas da Laricas/g, operador+" da Laricas")
      .replace(/Aqui é o Lucas/g, "Aqui é "+artigo+" "+operador);
  };

  const registrarEnvio = async (c, scriptId) => {
    const s = SCRIPTS_CLUB.find(s=>s.id===scriptId);
    if (!s) return;
    // Avanço automático de follow-up — mapeado ou padrão de 2 dias
    const autoFU = SCRIPT_AUTO_FOLLOWUP[scriptId] || SCRIPT_AUTO_FOLLOWUP_DEFAULT;
    const logEntry = {
      texto:"Script enviado: "+s.label+" — próximo follow-up em "+autoFU.dias+"d",
      data:new Date().toLocaleDateString("pt-BR"),
      hora:new Date().toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"}),
      resp:c.responsavel||"",
    };
    const scriptsEnviados = [...(c.scriptsEnviados||[]), {id:scriptId, data:hoje, converteu:false}];
    const atualizado = {
      ...c,
      scriptsEnviados,
      scriptUsado: scriptId,
      logAtividade:[logEntry,...(c.logAtividade||[])].slice(0,30),
      proximoFollowup: addDays(autoFU.dias),
      dataUltimoContato: hoje,
      // Marcar data de abordagem se for o primeiro contato
      ...(c.dataAbordagem ? {} : { dataAbordagem: hoje }),
      // Marcar como contatado se ainda não estava no funil
      ...(c.statusClub ? {} : { statusClub: "contatado", tentativasClub: (c.tentativasClub||0)+1 }),
    };
    await saveCliente(atualizado);
    // Mostrar banner de confirmação do follow-up
    setFollowUpProposto({
      data: addDays(autoFU.dias),
      label: autoFU.label,
      clienteId: c.id,
      status: atualizado.statusClub,
    });
    setTimeout(() => setFollowUpProposto(null), 8000);
  };

  const abrirWhatsApp = (c, scriptId) => {
    const id = scriptId||sugerirScript(c);
    const s = SCRIPTS_CLUB.find(s=>s.id===id);
    if (!s || !c.telefone) return;
    const texto = personalizarScript(s.copy, c);
    const tel = (c.telefone||"").replace(/\D/g,"");
    const url = `https://wa.me/55${tel}?text=${encodeURIComponent(texto)}`;
    window.open(url, "_blank");
    registrarEnvio(c, id);
  };

  // ── Segmentos para o painel Hoje ──────────────────────────────────────────
  const vencidos = clientes.filter(c => c.proximoFollowup && c.proximoFollowup < hoje);
  const interessadosSemLink = clientes.filter(c => c.statusClub === "interessado");
  const linkSemFechamento = clientes.filter(c => c.statusClub === "link_enviado");
  const semResposta48h = clientes.filter(c => {
    if (c.statusClub !== "contatado") return false;
    if (!c.dataUltimoContato) return false;
    const dias = Math.round((new Date()-new Date(c.dataUltimoContato+"T12:00:00"))/86400000);
    return dias >= 2;
  });
  const janeslaAberta = clientes.filter(c => {
    if (c.statusClub) return false;
    if (c._inativa || c._muitoInativa) return false;
    return c._diasParaProxima !== null && c._diasParaProxima !== undefined && c._diasParaProxima >= -3 && c._diasParaProxima <= 7;
  });
  const quentesNaoAbordados = clientes.filter(c => !c.statusClub && (c._score||0) >= 65 && !c._inativa).slice(0, 10);
  const reativarPrimeiro = clientes.filter(c => c._inativa && !c.statusClub).slice(0, 5);

  // ── Filtro da lista ────────────────────────────────────────────────────────
  // Ordenação por prioridade — mesma lógica do Hoje
  const prioClub = (c) => {
    const dtU = c.dataUltimoContato ? new Date(c.dataUltimoContato+"T12:00:00") : null;
    const diasSemContato = dtU ? Math.round((new Date()-dtU)/86400000) : 999;
    let p = 0;
    if (c.statusClub === "interessado") p += 100;
    else if (c.statusClub === "respondeu") p += 80;
    else if (c.proximoFollowup && c.proximoFollowup <= hoje) p += 70;
    else if (c.statusClub === "link_enviado") p += 60;
    else if (c.statusClub === "contatado" && diasSemContato >= 2) p += 50;
    else if (!c.statusClub && c._diasParaProxima !== null && c._diasParaProxima <= 5) p += 90;
    p += (c._score||0) * 0.3;
    return p;
  };

  // Quando há busca, pesquisa em TODOS os clientes (incluindo leads manuais com p<2)
  const baseParaFiltro = busca ? todosParaBusca : clientes;
  const listaFiltrada = baseParaFiltro.filter(c => {
    if (filtroStatus === "hoje") {
      // Só aparece quem tem follow-up para HOJE ou vencido — sem exceção por status
      const followUpHoje = c.proximoFollowup && c.proximoFollowup <= hoje;
      const janelaAberta = c._diasParaProxima !== null && c._diasParaProxima !== undefined && c._diasParaProxima >= -2 && c._diasParaProxima <= 5 && !c.statusClub;
      if (!followUpHoje && !janelaAberta) return false;
    } else if (filtroStatus === "nao_abordado") {
      if (c.statusClub) return false; // qualquer status = já abordado
    } else if (filtroStatus) {
      if (c.statusClub !== filtroStatus) return false;
    }
    if (busca) {
      const q = busca.toLowerCase();
      if (!(c.nome||"").toLowerCase().includes(q) && !(c.telefone||"").includes(q)) return false;
    }
    return true;
  }).sort((a,b) => {
    // Sorts cumulativos se ativos
    if (sorts.length > 0) {
      for (const s of sorts) {
        let va, vb;
        if (s.campo==="p") { va=a.p||0; vb=b.p||0; }
        else if (s.campo==="ciclo") { va=a.cicloMedio||999; vb=b.cicloMedio||999; }
        else if (s.campo==="dias") { va=a._diasUlt||999; vb=b._diasUlt||999; }
        else if (s.campo==="score") { va=a._score||0; vb=b._score||0; }
        else continue;
        const diff = s.dir==="desc" ? vb-va : va-vb;
        if (diff!==0) return diff;
      }
      return 0;
    }
    // Padrão: prioClub + score
    const filtroEspecifico = filtroStatus && filtroStatus !== "hoje" && filtroStatus !== "nao_abordado";
    if (filtroEspecifico) return (b._score||0) - (a._score||0);
    const pa = prioClub(a), pb = prioClub(b);
    if (pb !== pa) return pb - pa;
    return (b._score||0) - (a._score||0);
  });

  if (loading) return <div style={{textAlign:"center",padding:60,color:"var(--color-text-tertiary)"}}>Carregando funil Club...</div>;

  const statusInfo = (id) => STATUS_CLUB.find(s=>s.id===id)||STATUS_CLUB[0];

  // ── Card de urgência clicável ──────────────────────────────────────────────
  const CardUrgencia = ({emoji, label, count, cor, corD, corL, lista}) => {
    if (count === 0) return null;
    return (
      <button onClick={()=>{ if(lista?.length>0){setSel(lista[0]);setAba("lista");} }}
        style={{ flex:1,minWidth:140,background:corL,border:"0.5px solid "+cor,borderRadius:10,padding:"10px 14px",cursor:"pointer",textAlign:"left" }}>
        <div style={{fontSize:22,marginBottom:4}}>{emoji}</div>
        <div style={{fontSize:20,fontWeight:600,color:corD}}>{count}</div>
        <div style={{fontSize:11,color:corD,lineHeight:1.3}}>{label}</div>
      </button>
    );
  };

  // ── Card cliente na lista ──────────────────────────────────────────────────
  const CardLista = ({c}) => {
    const st = statusInfo(c.statusClub);
    const ciclo = c.cicloMedio||0;
    const janelaAberta = c._diasParaProxima !== null && c._diasParaProxima !== undefined && c._diasParaProxima >= -3 && c._diasParaProxima <= 7 && !c._inativa;
    const isSelected = sel?.id === c.id;
    return (
      <div style={{display:"flex",alignItems:"flex-start",gap:6,marginBottom:6}}>
      {modoLote&&(
        <input type="checkbox" checked={selecionadosLote.has(c.id)}
          onChange={e=>{e.stopPropagation();setSelecionadosLote(prev=>{const n=new Set(prev);e.target.checked?n.add(c.id):n.delete(c.id);return n;});}}
          style={{marginTop:14,flexShrink:0,cursor:"pointer",width:14,height:14}}/>
      )}
      <button onClick={()=>{if(modoLote){setSelecionadosLote(prev=>{const n=new Set(prev);n.has(c.id)?n.delete(c.id):n.add(c.id);return n;});return;}setSel(c);setScriptSel(sugerirScript(c));setCampos({});}}
        style={{ flex:1,textAlign:"left",padding:"10px 12px",borderRadius:10,
          border:"1.5px solid "+(selecionadosLote.has(c.id)?C.purple:isSelected?C.teal:"var(--color-border-tertiary)"),
          background:selecionadosLote.has(c.id)?C.purpleL:isSelected?C.tealL:"var(--color-background-secondary)",cursor:"pointer",transition:"all 0.15s" }}>
        <div style={{display:"flex",alignItems:"flex-start",gap:8,marginBottom:6}}>
          <div style={{flex:1}}>
            <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
              <span style={{fontSize:13,fontWeight:500,color:"var(--color-text-primary)"}}>{c.nome}</span>
              {janelaAberta&&<span style={{fontSize:9,fontWeight:600,background:C.greenL,color:C.greenD,padding:"1px 6px",borderRadius:10}}>🛒 COMPRANDO AGORA</span>}
              {c.telefone&&(
                <button onClick={e=>{e.stopPropagation();const t=normalizarTelefone(c.telefone);navigator.clipboard.writeText(t).catch(()=>{});}}
                  title={"Copiar: "+normalizarTelefone(c.telefone)}
                  style={{fontSize:9,fontWeight:500,background:"var(--color-background-secondary)",color:"var(--color-text-tertiary)",padding:"1px 6px",borderRadius:10,border:"0.5px solid var(--color-border-tertiary)",cursor:"pointer"}}>
                  📋 Tel
                </button>
              )}
            </div>
            <div style={{fontSize:11,color:c._inativa?"var(--color-text-tertiary)":"var(--color-text-tertiary)",marginTop:2}}>
              {c.p||0}p · R${(c.gasto||0).toFixed(0)} · ciclo {ciclo||"?"}d
              {c._diasUlt&&<span style={{marginLeft:6,color:c._muitoInativa?C.coralD:c._inativa?C.amber:"var(--color-text-tertiary)"}}>
                · {c._diasUlt}d sem comprar{c._muitoInativa?" ⚠":""}
              </span>}
            </div>
          </div>
          <span style={{fontSize:10,fontWeight:500,color:st.cor,background:st.cor+"22",padding:"2px 7px",borderRadius:20,flexShrink:0,whiteSpace:"nowrap"}}>
            {st.emoji} {st.label}
          </span>
        </div>
        <ScoreBar score={c._score||0}/>
        <div style={{display:"flex",flexWrap:"wrap",gap:4,marginTop:4}}>
          {c._diasParaProxima!==null&&c._diasParaProxima!==undefined&&!c._inativa&&(
            <span style={{fontSize:10,fontWeight:500,
              color:c._diasParaProxima<=2?C.greenD:c._diasParaProxima<=7?C.amberD:"var(--color-text-tertiary)",
              background:c._diasParaProxima<=2?C.greenL:c._diasParaProxima<=7?C.amberL:"var(--color-background-primary)",
              padding:"1px 6px",borderRadius:10}}>
              {c._diasParaProxima<=0?"🛒 Comprando agora":c._diasParaProxima<=7?"📅 Em "+c._diasParaProxima+"d":"🗓 Em "+c._diasParaProxima+"d"}
            </span>
          )}
          {c.dataUltimoContato&&c.statusClub&&c.statusClub!=="fechou"&&(()=>{
            const dias=Math.round((new Date()-new Date(c.dataUltimoContato+"T12:00:00"))/86400000);
            return <span style={{fontSize:10,color:"var(--color-text-tertiary)"}}>Contatada há {dias}d</span>;
          })()}
          {(c.tentativasClub||0)>0&&c.statusClub!=="fechou"&&(
            <span style={{fontSize:10,color:(c.tentativasClub||0)>=3?C.coralD:C.amber,
              background:(c.tentativasClub||0)>=3?C.coralL:C.amberL,
              padding:"1px 6px",borderRadius:10,fontWeight:500}}>
              {c.tentativasClub}ª tentativa{(c.tentativasClub||0)>=3?" ⚠":""}
            </span>
          )}
          {c.proximoFollowup&&c.proximoFollowup<=hoje&&(
            <span style={{fontSize:10,color:C.coralD,background:C.coralL,padding:"1px 6px",borderRadius:10}}>⚠ Follow-up vencido</span>
          )}
        </div>
      </button>
    </div>
    );
  };

  // ── Painel de detalhes do cliente selecionado ──────────────────────────────
  const PainelDetalhe = () => {
    if (!sel) return (
      <div style={{textAlign:"center",padding:40,color:"var(--color-text-tertiary)"}}>
        <div style={{fontSize:32,marginBottom:12}}>👈</div>
        <div style={{fontSize:13}}>Selecione um cliente para ver detalhes e scripts</div>
      </div>
    );
    const c = sel;
    const st = statusInfo(c.statusClub);
    const planoRec = c.planoRec || sugerirPlano(c);
    const ciclo = c.cicloMedio||0;
    const janelaAberta = c._diasParaProxima !== null && c._diasParaProxima !== undefined && c._diasParaProxima >= -3 && c._diasParaProxima <= 7 && !c._inativa;
    // Local state para campos de texto — evita re-render a cada tecla (bug de scroll)
    const [obsLocal, setObsLocal] = React.useState(c.obsClub||"");
    React.useEffect(()=>setObsLocal(c.obsClub||""), [c.id]);

    const scriptAtual = SCRIPTS_CLUB.find(s=>s.id===scriptSel) || SCRIPTS_CLUB.find(s=>s.id===sugerirScript(c));
    const textoScript = scriptAtual ? personalizarScript(scriptAtual.copy, c) : "";

    const copiar = (txt, id) => {
      navigator.clipboard.writeText(txt).catch(()=>{});
      setCopiadoId(id);
      setTimeout(()=>setCopiadoId(""),2000);
    };

    return (
      <div>
        {/* Header cliente */}
        <div style={{background:C.tealL,border:"0.5px solid "+C.teal,borderRadius:10,padding:"12px 14px",marginBottom:12}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
            <div style={{flex:1}}>
              <div style={{fontSize:15,fontWeight:600,color:"var(--color-text-primary)"}}>{c.nome}</div>
              <div style={{fontSize:12,color:"var(--color-text-secondary)"}}>{c.telefone||"—"} · {c.email||"—"}</div>
              {(getCidade(c.cep)||c.fora)&&(
                <div style={{fontSize:11,color:C.tealD,marginTop:3}}>
                  {c.fora?"🌎":"📍"} {getCidade(c.cep)||"Cidade não identificada"}
                  {c.fora&&<span style={{marginLeft:4,fontSize:10,background:C.amberL,color:C.amberD,padding:"0px 5px",borderRadius:8,fontWeight:500}}>Fora de SP</span>}
                </div>
              )}
            </div>
            {c.telefone&&(
              <a href={"https://wa.me/55"+(c.telefone||"").replace(/\D/g,"")} target="_blank" rel="noopener noreferrer"
                style={{background:"#25D366",border:"none",borderRadius:8,padding:"6px 12px",fontSize:12,color:"#fff",fontWeight:500,textDecoration:"none"}}>
                💬 WhatsApp
              </a>
            )}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
            <div style={{textAlign:"center"}}>
              <div style={{fontSize:18,fontWeight:600,color:C.tealD}}>{c.p||0}</div>
              <div style={{fontSize:10,color:C.teal,textTransform:"uppercase"}}>Pedidos</div>
            </div>
            <div style={{textAlign:"center"}}>
              <div style={{fontSize:18,fontWeight:600,color:C.tealD}}>R${(c.gasto||0).toFixed(0)}</div>
              <div style={{fontSize:10,color:C.teal,textTransform:"uppercase"}}>Total gasto</div>
            </div>
            <div style={{textAlign:"center"}}>
              <div style={{fontSize:18,fontWeight:600,color:C.tealD}}>{ciclo||"?"}d</div>
              <div style={{fontSize:10,color:C.teal,textTransform:"uppercase"}}>Ciclo médio</div>
            </div>
          </div>
          {janelaAberta&&(
            <div style={{marginTop:8,background:C.green,borderRadius:8,padding:"6px 10px",fontSize:12,fontWeight:500,color:"#fff",textAlign:"center"}}>
              🛒 Janela de compra aberta — ela está pensando em Laricas agora!
            </div>
          )}
          {c._diasParaProxima!==null&&!janelaAberta&&!c._inativa&&(
            <div style={{marginTop:6,fontSize:11,color:C.teal,textAlign:"center"}}>
              {c._diasParaProxima>0
                ? "Próxima compra estimada em "+c._diasParaProxima+" dias"
                : "Próxima compra passou há "+(Math.abs(c._diasParaProxima))+" dias — momento ideal!"}
            </div>
          )}
          {c._inativa&&(
            <div style={{marginTop:6,fontSize:11,color:C.coralD,background:C.coralL,borderRadius:6,padding:"4px 8px",textAlign:"center"}}>
              ⚠ Inativa há {c._diasUlt} dias — reativar antes de oferecer Club
            </div>
          )}
        </div>

        {/* Status rápido */}
        <div style={{marginBottom:12}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
            <div style={{fontSize:11,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",flex:1}}>Status no funil</div>
            {(c.tentativasClub||0)>0&&<span style={{fontSize:10,fontWeight:500,
              color:(c.tentativasClub||0)>=3?C.coralD:C.amberD,
              background:(c.tentativasClub||0)>=3?C.coralL:C.amberL,
              padding:"2px 8px",borderRadius:20}}>
              {c.tentativasClub}ª tentativa{(c.tentativasClub||0)>=3?" — considerar encerrar":""}
            </span>}
          </div>
          {(c.tentativasClub||0)>=3&&c.statusClub!=="fechou"&&(
            <div style={{background:C.coralL,border:"0.5px solid "+C.coral,borderRadius:8,padding:"8px 12px",marginBottom:8,fontSize:12,color:C.coralD}}>
              ⚠ 3 tentativas sem resposta. Marcar como Perdido e focar em quem está mais quente?
            </div>
          )}
          <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
            {STATUS_CLUB.filter(s=>s.id!=="fechou"&&s.id!=="perdido").map(s=>(
              <button key={s.id} onClick={()=>atualizarStatus(c,s.id)} disabled={salvando}
                style={{padding:"4px 10px",borderRadius:20,fontSize:11,fontWeight:500,cursor:"pointer",
                  background:c.statusClub===s.id?s.cor:"var(--color-background-secondary)",
                  color:c.statusClub===s.id?"#fff":"var(--color-text-secondary)",
                  border:"0.5px solid "+(c.statusClub===s.id?s.cor:"var(--color-border-tertiary)")}}>
                {s.emoji} {s.label}
              </button>
            ))}
          </div>
          <div style={{display:"flex",gap:6,marginTop:6}}>
            <button onClick={()=>atualizarStatus(c,"fechou")} disabled={salvando}
              style={{flex:1,padding:"8px",borderRadius:8,fontSize:12,fontWeight:600,cursor:"pointer",background:C.green,color:"#fff",border:"none"}}>
              🏆 Fechou!
            </button>
            <button onClick={()=>atualizarStatus(c,"perdido")} disabled={salvando}
              style={{flex:1,padding:"8px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",background:"var(--color-background-secondary)",color:C.coralD,border:"0.5px solid "+C.coral}}>
              ✗ Perdido
            </button>
          </div>
          {ok&&<div style={{fontSize:12,color:C.greenD,marginTop:4,textAlign:"center"}}>✓ {ok}</div>}
          {followUpProposto&&followUpProposto.clienteId===c.id&&(
            <div style={{background:C.amberL,border:"0.5px solid "+C.amber,borderRadius:8,padding:"10px 12px",marginTop:8}}>
              <div style={{fontSize:11,color:C.amberD,fontWeight:500,marginBottom:6}}>
                🔔 Follow-up automático: {new Date(followUpProposto.data+"T12:00:00").toLocaleDateString("pt-BR")} — {followUpProposto.label}
              </div>
              <div style={{display:"flex",gap:6,alignItems:"center"}}>
                <input type="date" defaultValue={followUpProposto.data}
                  onChange={e=>{
                    const nova = e.target.value;
                    const atualizado = {...c, proximoFollowup: nova};
                    saveCliente(atualizado);
                    setFollowUpProposto(null);
                  }}
                  style={{flex:1,padding:"4px 8px",borderRadius:6,border:"0.5px solid "+C.amber,fontSize:11,color:C.amberD,background:"#fff"}}/>
                <button onClick={()=>setFollowUpProposto(null)}
                  style={{padding:"4px 10px",borderRadius:6,fontSize:11,cursor:"pointer",background:C.green,color:"#fff",border:"none",fontWeight:500}}>
                  ✓ OK
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Script sugerido */}
        <div style={{marginBottom:12}}>
          <div style={{fontSize:11,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>
            Script {scriptAtual&&<span style={{color:C.teal}}>— {scriptAtual.label}</span>}
          </div>
          {scriptAtual&&<div style={{fontSize:10,color:C.amber,marginBottom:6}}>💡 {razaoScript(c)}</div>}
          <div style={{display:"flex",flexWrap:"wrap",gap:4,marginBottom:8}}>
            {SCRIPTS_CLUB.map(s=>(
              <button key={s.id} onClick={()=>setScriptSel(s.id)}
                style={{padding:"3px 8px",borderRadius:6,fontSize:10,fontWeight:500,cursor:"pointer",
                  background:scriptSel===s.id?C.purple:"var(--color-background-secondary)",
                  color:scriptSel===s.id?"#fff":"var(--color-text-secondary)",
                  border:"0.5px solid "+(scriptSel===s.id?C.purple:"var(--color-border-tertiary)")}}>
                {s.label.split("—")[0].trim()}
              </button>
            ))}
          </div>
          {textoScript&&(
            <div>
              <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",fontSize:13,color:"var(--color-text-primary)",lineHeight:1.8,whiteSpace:"pre-line",fontFamily:"inherit",borderLeft:"3px solid "+C.teal,marginBottom:8}}>
                {textoScript}
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>{ copiar(textoScript,"script"); registrarEnvio(c, scriptSel||sugerirScript(c)); }}
                  style={{flex:1,padding:"8px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",
                    background:copiadoId==="script"?C.green:C.tealL,color:copiadoId==="script"?"#fff":C.tealD,
                    border:"0.5px solid "+C.teal,transition:"all 0.2s"}}>
                  {copiadoId==="script"?"✓ Copiado!":"📋 Copiar mensagem"}
                </button>
                {c.telefone&&(
                  <button onClick={()=>abrirWhatsApp(c,scriptSel)}
                    style={{flex:1,padding:"8px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",
                      background:"#25D366",color:"#fff",border:"none"}}>
                    💬 Enviar no WhatsApp
                  </button>
                )}
              </div>

              {/* Respostas possíveis */}
              {ARVORE[scriptSel||sugerirScript(c)]&&(
                <div style={{marginTop:10}}>
                  <div style={{fontSize:11,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:6}}>
                    Como ela respondeu?
                  </div>
                  <div style={{display:"flex",flexDirection:"column",gap:6}}>
                    {(ARVORE[scriptSel||sugerirScript(c)]||[]).map((op,i)=>(
                      <button key={i} onClick={async()=>{
                        // Registrar no log
                        const logEntry = {
                          texto:"Resposta: "+op.label+" → "+((op.proximoScript&&SCRIPTS_CLUB.find(s=>s.id===op.proximoScript)?.label)||op.novoStatus),
                          data:new Date().toLocaleDateString("pt-BR"),
                          hora:new Date().toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"}),
                          resp:c.responsavel||"",
                        };
                        // Se selecionou "não pode agora" com motivo específico, pré-marcar motivoNaoPode
                        const motivoPreenchido = op.novoStatus==="nao_agora" && op.followUpDias!==undefined && op.followUpDias!==null
                          ? MOTIVOS_NAO_PODE.find(m=>op.label.toLowerCase().includes(m.toLowerCase().split(" ")[0].toLowerCase())) || ""
                          : "";
                        const atualizado={
                          ...c,
                          statusClub:op.novoStatus,
                          ...(op.novoStatus==="contatado"&&!c.dataAbordagem?{dataAbordagem:hoje}:{}),
                          dataUltimoContato:hoje,
                          ...(( op.followUpDias !== undefined ? op.followUpDias : FOLLOW_UP_DAYS[op.novoStatus] ) ? {proximoFollowup:addDays(op.followUpDias !== undefined ? op.followUpDias : FOLLOW_UP_DAYS[op.novoStatus])} : {}),
                          ...(op.novoStatus==="fechou"?{dataConversao:hoje}:{}),
                          ...(op.novoStatus==="contatado"?{tentativasClub:(c.tentativasClub||0)+1}:{}),
                          logAtividade:[logEntry,...(c.logAtividade||[])].slice(0,30),
                        };
                        await saveCliente(atualizado);
                        if(op.proximoScript) setScriptSel(op.proximoScript);
                        else if(op.novoStatus==="perdido"||op.novoStatus==="link_enviado") setScriptSel(null);
                        const fuDias2 = op.followUpDias !== undefined ? op.followUpDias : FOLLOW_UP_DAYS[op.novoStatus];
                        if(fuDias2){
                          setFollowUpProposto({data:addDays(fuDias2),label:FOLLOW_UP_LABELS[op.novoStatus]||"",clienteId:c.id,status:op.novoStatus});
                          setTimeout(()=>setFollowUpProposto(null),8000);
                        } else if(op.followUpDias===null) {
                          // followUpDias explicitamente null = pedir ao usuário que defina manualmente
                          setFollowUpProposto({data:addDays(30),label:"Defina a data de retorno manualmente",clienteId:c.id,status:op.novoStatus});
                          setTimeout(()=>setFollowUpProposto(null),15000);
                        }
                      }}
                        style={{width:"100%",textAlign:"left",padding:"8px 12px",borderRadius:8,fontSize:12,cursor:"pointer",
                          background:op.novoStatus==="perdido"?C.coralL:op.novoStatus==="link_enviado"||op.novoStatus==="interessado"?C.greenL:"var(--color-background-primary)",
                          border:"0.5px solid "+(op.novoStatus==="perdido"?C.coral:op.novoStatus==="link_enviado"||op.novoStatus==="interessado"?C.green:"var(--color-border-tertiary)"),
                          color:op.novoStatus==="perdido"?C.coralD:op.novoStatus==="link_enviado"||op.novoStatus==="interessado"?C.greenD:"var(--color-text-primary)",
                          display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontSize:16,flexShrink:0}}>{op.emoji}</span>
                        <div style={{flex:1}}>
                          <div style={{fontWeight:500}}>{op.label}</div>
                          {op.proximoScript&&<div style={{fontSize:10,color:"var(--color-text-tertiary)",marginTop:1}}>
                            → {SCRIPTS_CLUB.find(s=>s.id===op.proximoScript)?.label||op.proximoScript}
                          </div>}
                        </div>
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Histórico de scripts enviados */}
        {c.logAtividade&&c.logAtividade.filter(l=>l.texto&&l.texto.startsWith("Script")).length>0&&(
          <div style={{marginBottom:12}}>
            <div style={{fontSize:11,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:6}}>Histórico de contatos</div>
            <div style={{maxHeight:120,overflowY:"auto",display:"flex",flexDirection:"column",gap:4}}>
              {c.logAtividade.filter(l=>l.texto&&(l.texto.startsWith("Script")||l.texto.startsWith("Resposta"))).map((l,i)=>(
                <div key={i} style={{fontSize:10,color:"var(--color-text-secondary)",background:"var(--color-background-primary)",borderRadius:6,padding:"4px 8px",borderLeft:"2px solid "+(l.texto.startsWith("Resposta")?C.green:C.teal)}}>
                  <span style={{color:"var(--color-text-tertiary)"}}>{l.data} {l.hora}</span> — {l.texto}
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Campos de registro */}
        <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",marginBottom:10}}>
          <div style={{fontSize:11,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:10}}>Registro da abordagem</div>

          <div style={{marginBottom:8}}>
            <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4}}>Plano recomendado</div>
            <div style={{display:"flex",gap:6}}>
              {PLANOS.map(p=>(
                <button key={p} onClick={()=>saveCliente({...c,planoRec:p})}
                  style={{flex:1,padding:"6px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",
                    background:c.planoRec===p?C.purple:"var(--color-background-primary)",
                    color:c.planoRec===p?"#fff":"var(--color-text-secondary)",
                    border:"0.5px solid "+(c.planoRec===p?C.purple:"var(--color-border-tertiary)")}}>
                  {p}
                </button>
              ))}
            </div>
            {!c.planoRec&&<div style={{fontSize:10,color:C.amber,marginTop:4}}>Sugestão automática: {planoRec}</div>}
          </div>

          <div style={{marginBottom:8}}>
            <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4}}>Principal objeção</div>
            <select value={c.objClub||""} onChange={e=>saveCliente({...c,objClub:e.target.value})}
              style={{width:"100%",padding:"7px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-primary)"}}>
              <option value="">— Nenhuma objeção ainda —</option>
              {OBJECTIONS.map(o=><option key={o} value={o}>{o}</option>)}
            </select>
          </div>

          <div style={{marginBottom:8}}>
            <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4}}>Próximo follow-up</div>
            <input type="date" value={c.proximoFollowup||""} onChange={e=>saveCliente({...c,proximoFollowup:e.target.value})}
              style={{width:"100%",padding:"7px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-primary)"}}/>
            <div style={{display:"flex",gap:4,marginTop:4}}>
              {[["Amanhã",1],["2d",2],["7d",7],["14d",14]].map(([l,d])=>(
                <button key={l} onClick={()=>saveCliente({...c,proximoFollowup:addDays(d)})}
                  style={{padding:"3px 8px",borderRadius:6,fontSize:10,cursor:"pointer",background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-secondary)"}}>
                  {l}
                </button>
              ))}
            </div>
          </div>

          {c.statusClub==="nao_agora"&&(
            <div style={{marginBottom:8}}>
              <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4}}>Motivo — não pode agora</div>
              <select value={c.motivoNaoPode||""} onChange={e=>saveCliente({...c,motivoNaoPode:e.target.value})}
                style={{width:"100%",padding:"7px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-primary)"}}>
                <option value="">— Selecione o motivo —</option>
                {MOTIVOS_NAO_PODE.map(m=><option key={m} value={m}>{m}</option>)}
              </select>
              {c.motivoNaoPode&&<div style={{fontSize:10,color:C.amberD,marginTop:4}}>
                💡 Registrado — vai aparecer no Dash e no relatório diário
              </div>}
            </div>
          )}
          {c.scriptUsado==="reativacao"&&(
            <div style={{marginBottom:8}}>
              <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4}}>Motivo da inatividade</div>
              <select value={c.motivoInatividade||""} onChange={e=>saveCliente({...c,motivoInatividade:e.target.value})}
                style={{width:"100%",padding:"7px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-primary)"}}>
                <option value="">— O que ela respondeu? —</option>
                {MOTIVOS_INATIVIDADE.map(m=><option key={m} value={m}>{m}</option>)}
              </select>
              {c.motivoInatividade&&<div style={{fontSize:10,color:C.tealD,marginTop:4}}>
                💡 Registrado — use para ajustar a abordagem de reativação
              </div>}
            </div>
          )}
          <div>
            <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4}}>Observações</div>
            <textarea value={obsLocal} onChange={e=>setObsLocal(e.target.value)}
              onBlur={()=>{ if(obsLocal!==c.obsClub) saveCliente({...c,obsClub:obsLocal}); }}
              rows={2} placeholder="Contexto, tom da conversa, o que ela disse..."
              style={{width:"100%",padding:"7px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-primary)",outline:"none",resize:"vertical",fontFamily:"inherit"}}/>
          </div>
        </div>

        {/* Conversão — só aparece se fechou */}
        {c.statusClub==="fechou"&&(
          <div style={{background:C.greenL,border:"0.5px solid "+C.green,borderRadius:10,padding:"12px 14px"}}>
            <div style={{fontSize:12,fontWeight:500,color:C.greenD,marginBottom:8}}>🏆 Conversão registrada</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              <div>
                <div style={{fontSize:11,color:C.greenD,marginBottom:4}}>Plano fechado</div>
                <div style={{display:"flex",gap:4}}>
                  {PLANOS.map(p=>(
                    <button key={p} onClick={()=>saveCliente({...c,planoFechado:p})}
                      style={{flex:1,padding:"5px",borderRadius:6,fontSize:11,cursor:"pointer",
                        background:c.planoFechado===p?C.green:"#fff",color:c.planoFechado===p?"#fff":C.greenD,
                        border:"0.5px solid "+C.green}}>
                      {p}
                    </button>
                  ))}
                </div>
              </div>
              <div>
                <div style={{fontSize:11,color:C.greenD,marginBottom:4}}>Valor mensal</div>
                <input type="number" value={c.valorMensalClub||""} onChange={e=>saveCliente({...c,valorMensalClub:e.target.value})}
                  placeholder="R$"
                  style={{width:"100%",padding:"5px 8px",borderRadius:6,border:"0.5px solid "+C.green,fontSize:12,color:C.greenD,background:"#fff"}}/>
              </div>
            </div>
          </div>
        )}

        {/* HISTÓRICO DE SCRIPTS */}
        {(c.logAtividade||[]).filter(l=>l.texto&&l.texto.includes("Script")).length>0&&(
          <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",marginTop:10}}>
            <div style={{fontSize:11,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:8}}>Histórico de conversas</div>
            <div style={{maxHeight:160,overflowY:"auto"}}>
              {(c.logAtividade||[]).filter(l=>l.texto).map((l,i)=>(
                <div key={i} style={{display:"flex",gap:8,marginBottom:6,alignItems:"flex-start"}}>
                  <div style={{fontSize:9,color:"var(--color-text-tertiary)",whiteSpace:"nowrap",marginTop:1,minWidth:30}}>{l.hora||""}</div>
                  <div style={{flex:1,fontSize:11,color:"var(--color-text-primary)",lineHeight:1.4,
                    background:l.texto.startsWith("Script")?C.tealL:l.texto.startsWith("Resposta")?C.greenL:"var(--color-background-primary)",
                    borderRadius:6,padding:"3px 8px",
                    borderLeft:"2px solid "+(l.texto.startsWith("Script")?C.teal:l.texto.startsWith("Resposta")?C.green:"var(--color-border-tertiary)")}}>
                    {l.texto}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        <button onClick={()=>onAbrirPerfil&&onAbrirPerfil(c.id)}
          style={{width:"100%",marginTop:10,padding:"8px",borderRadius:8,fontSize:12,cursor:"pointer",
            background:"none",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-secondary)"}}>
          Abrir perfil completo →
        </button>
      </div>
    );
  };

  // ── Dashboard simples ──────────────────────────────────────────────────────
  const DashClub = () => {
    const todos = clientesDash;
    const abordados = todos.filter(c=>c.statusClub&&c.statusClub!=="");
    const responderam = todos.filter(c=>["respondeu","interessado","link_enviado","fechou"].includes(c.statusClub));
    const interessados = todos.filter(c=>["interessado","link_enviado","fechou"].includes(c.statusClub));
    const convertidos = todos.filter(c=>c.statusClub==="fechou");
    const txResposta = abordados.length>0?Math.round(responderam.length/abordados.length*100):0;
    const txConversao = abordados.length>0?Math.round(convertidos.length/abordados.length*100):0;
    const objCounts = {};
    todos.forEach(c=>{ if(c.objClub){objCounts[c.objClub]=(objCounts[c.objClub]||0)+1;}});
    const objRanking = Object.entries(objCounts).sort((a,b)=>b[1]-a[1]);
    return (
      <div>
        <div style={{display:"flex",justifyContent:"flex-end",marginBottom:12}}>
          <button onClick={()=>setShowRitsSync(true)}
            style={{padding:"8px 16px",borderRadius:10,fontSize:12,fontWeight:500,cursor:"pointer",background:C.teal,color:"#fff",border:"none",display:"flex",alignItems:"center",gap:6}}>
            🔄 Sincronizar com RitsPay
          </button>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:16}}>
          {[
            ["Abordados",abordados.length,C.teal,C.tealD,C.tealL],
            ["Responderam",responderam.length,C.blue,C.blueD,C.blueL],
            ["Interessados",interessados.length,C.green,C.greenD,C.greenL],
            ["Convertidos",convertidos.length,C.green,C.greenD,C.greenL],
            ["Taxa resposta",txResposta+"%",C.amber,C.amberD,C.amberL],
            ["Taxa conversão",txConversao+"%",C.purple,C.purpleD,C.purpleL],
          ].map(([label,val,cor,corD,corL])=>(
            <div key={label} style={{background:corL,borderRadius:10,padding:"12px 14px",border:"0.5px solid "+cor}}>
              <div style={{fontSize:10,color:corD,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>{label}</div>
              <div style={{fontSize:24,fontWeight:600,color:corD}}>{val}</div>
            </div>
          ))}
        </div>
        {objRanking.length>0&&(
          <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px"}}>
            <div style={{fontSize:11,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:10}}>Objeções mais comuns</div>
            {objRanking.map(([obj,count])=>(
              <div key={obj} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                <div style={{flex:1,fontSize:12,color:"var(--color-text-primary)"}}>{obj}</div>
                <div style={{width:`${Math.round(count/objRanking[0][1]*100)}%`,maxWidth:120,height:8,background:C.coral,borderRadius:4,minWidth:20}}/>
                <span style={{fontSize:12,fontWeight:500,color:C.coralD,minWidth:20,textAlign:"right"}}>{count}</span>
              </div>
            ))}
          </div>
        )}
        {(()=>{
          // Tempo médio por etapa do funil Club
          const temposEtapa = {"contatado":[],"respondeu":[],"interessado":[],"link_enviado":[]};
          todos.forEach(c=>{
            const hist = c.logAtividade||[];
            const statusChanges = hist.filter(l=>l.texto&&(l.texto.includes("→"))).reverse();
            for(let i=0;i<statusChanges.length-1;i++){
              const d1=statusChanges[i].data, d2=statusChanges[i+1].data;
              if(!d1||!d2) continue;
              const [d1d,d1m,d1y]=d1.split("/"); const [d2d,d2m,d2y]=d2.split("/");
              const diff=Math.abs(Math.round((new Date(`${d1y}-${d1m}-${d1d}`)-new Date(`${d2y}-${d2m}-${d2d}`))/86400000));
              if(diff>=0&&diff<60){
                const statusAtual=c.statusClub;
                if(temposEtapa[statusAtual]) temposEtapa[statusAtual].push(diff);
              }
            }
          });
          const etapasComTempo = Object.entries(temposEtapa).filter(([_,v])=>v.length>0);
          if(!etapasComTempo.length) return null;
          return (
            <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",marginTop:12}}>
              <div style={{fontSize:11,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:10}}>⏱ Tempo médio por etapa</div>
              {etapasComTempo.map(([etapa,dias])=>{
                const media=Math.round(dias.reduce((a,b)=>a+b,0)/dias.length);
                const label=STATUS_CLUB.find(s=>s.id===etapa)?.label||etapa;
                return (
                  <div key={etapa} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                    <div style={{flex:1,fontSize:12}}>{label}</div>
                    <div style={{fontSize:12,fontWeight:500,color:media<=2?C.greenD:media<=5?C.amberD:C.coralD}}>{media}d em média</div>
                    <div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>({dias.length} casos)</div>
                  </div>
                );
              })}
            </div>
          );
        })()}
        {(()=>{
          const inativCounts = {};
          todos.forEach(c => { if(c.motivoInatividade) inativCounts[c.motivoInatividade]=(inativCounts[c.motivoInatividade]||0)+1; });
          const inativRanking = Object.entries(inativCounts).sort((a,b)=>b[1]-a[1]);
          if(!inativRanking.length) return null;
          return (
            <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",marginTop:12}}>
              <div style={{fontSize:11,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:10}}>
                🔄 Motivos de inatividade ({todos.filter(c=>c.scriptUsado==="reativacao").length} reativadas)
              </div>
              {inativRanking.map(([motivo,count])=>(
                <div key={motivo} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                  <div style={{flex:1,fontSize:12,color:"var(--color-text-primary)"}}>{motivo}</div>
                  <div style={{width:`${Math.round(count/inativRanking[0][1]*100)}%`,maxWidth:120,height:8,background:C.teal,borderRadius:4,minWidth:20}}/>
                  <span style={{fontSize:12,fontWeight:500,color:C.tealD,minWidth:20,textAlign:"right"}}>{count}</span>
                </div>
              ))}
            </div>
          );
        })()}
        {(()=>{
          const motivoCounts = {};
          todos.forEach(c => { if(c.motivoNaoPode) motivoCounts[c.motivoNaoPode]=(motivoCounts[c.motivoNaoPode]||0)+1; });
          const motivoRanking = Object.entries(motivoCounts).sort((a,b)=>b[1]-a[1]);
          if(!motivoRanking.length) return null;
          return (
            <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",marginTop:12}}>
              <div style={{fontSize:11,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:10}}>
                ⏳ Motivos — não pode agora ({todos.filter(c=>c.statusClub==="nao_agora").length} clientes)
              </div>
              {motivoRanking.map(([motivo,count])=>(
                <div key={motivo} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                  <div style={{flex:1,fontSize:12,color:"var(--color-text-primary)"}}>{motivo}</div>
                  <div style={{width:`${Math.round(count/motivoRanking[0][1]*100)}%`,maxWidth:120,height:8,background:C.amber,borderRadius:4,minWidth:20}}/>
                  <span style={{fontSize:12,fontWeight:500,color:C.amberD,minWidth:20,textAlign:"right"}}>{count}</span>
                </div>
              ))}
            </div>
          );
        })()}
        {(()=>{
          const scriptStats = {};
          const aberturas = ["recorrente","ticket_alto","kit"];
          todos.forEach(c => {
            (c.scriptsEnviados||[]).forEach(s => {
              if(!scriptStats[s.id]) scriptStats[s.id]={label:SCRIPTS_CLUB.find(sc=>sc.id===s.id)?.label||s.id,enviados:0,responderam:0,convertidos:0};
              scriptStats[s.id].enviados++;
              if(["respondeu","interessado","link_enviado","fechou"].includes(c.statusClub)) scriptStats[s.id].responderam++;
              if(c.statusClub==="fechou") scriptStats[s.id].convertidos++;
            });
          });
          const abStats = aberturas.map(id=>({id,...(scriptStats[id]||{label:SCRIPTS_CLUB.find(s=>s.id===id)?.label||id,enviados:0,responderam:0,convertidos:0})})).filter(s=>s.enviados>0);
          const outrosStats = Object.entries(scriptStats).filter(([id])=>!aberturas.includes(id)).map(([,s])=>s).filter(s=>s.enviados>0).sort((a,b)=>b.enviados-a.enviados);
          if(!Object.keys(scriptStats).length) return null;
          return (
            <div>
              {abStats.length>0&&(
                <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",marginTop:12}}>
                  <div style={{fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:10}}>🧪 A/B — Scripts de abertura</div>
                  {abStats.map(s=>{
                    const txResp=s.enviados>0?Math.round(s.responderam/s.enviados*100):0;
                    const txConv=s.enviados>0?Math.round(s.convertidos/s.enviados*100):0;
                    return (
                      <div key={s.label} style={{marginBottom:10}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                          <span style={{fontSize:11,fontWeight:500,color:"var(--color-text-primary)"}}>{s.label}</span>
                          <span style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{s.enviados} envios</span>
                        </div>
                        <div style={{height:8,background:"var(--color-border-tertiary)",borderRadius:4,overflow:"hidden",marginBottom:2}}>
                          <div style={{width:txResp+"%",height:"100%",background:C.teal,borderRadius:4}}/>
                        </div>
                        <div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>
                          Resposta: <span style={{color:C.tealD,fontWeight:500}}>{txResp}%</span>
                          {" · "}Conversão: <span style={{color:txConv>=10?C.greenD:C.amberD,fontWeight:500}}>{txConv}%</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
              {outrosStats.length>0&&(
                <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",marginTop:12}}>
                  <div style={{fontSize:11,fontWeight:500,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:10}}>📊 Conversão por script</div>
                  {outrosStats.map(s=>{
                    const tx=s.enviados>0?Math.round(s.convertidos/s.enviados*100):0;
                    return (
                      <div key={s.label} style={{marginBottom:8}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                          <span style={{fontSize:11,fontWeight:500,color:"var(--color-text-primary)"}}>{s.label}</span>
                          <span style={{fontSize:11,color:tx>=10?C.greenD:tx>=5?C.amberD:C.coralD,fontWeight:500}}>{tx}% ({s.convertidos}/{s.enviados})</span>
                        </div>
                        <div style={{height:6,background:"var(--color-border-tertiary)",borderRadius:3,overflow:"hidden"}}>
                          <div style={{width:tx+"%",height:"100%",background:tx>=10?C.green:tx>=5?C.amber:C.coral,borderRadius:3}}/>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })()}
      </div>
    );
  };

  // ── Biblioteca de scripts ──────────────────────────────────────────────────
  const Biblioteca = () => (
    <div>
      {SCRIPTS_CLUB.map(s=>(
        <div key={s.id} style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"12px 14px",marginBottom:10}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
            <div style={{flex:1}}>
              <div style={{fontSize:13,fontWeight:500,color:"var(--color-text-primary)"}}>{s.label}</div>
              <div style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{s.perfil}</div>
            </div>
          </div>
          <div style={{background:"var(--color-background-primary)",borderRadius:8,padding:"10px 12px",fontSize:12,color:"var(--color-text-primary)",lineHeight:1.8,whiteSpace:"pre-line",fontFamily:"inherit",borderLeft:"3px solid "+C.teal,marginBottom:8}}>
            {s.copy}
          </div>
          <button onClick={()=>{navigator.clipboard.writeText(s.copy).catch(()=>{});setCopiadoId(s.id);setTimeout(()=>setCopiadoId(""),2000);}}
            style={{padding:"5px 14px",borderRadius:6,fontSize:11,fontWeight:500,cursor:"pointer",
              background:copiadoId===s.id?C.green:C.tealL,color:copiadoId===s.id?"#fff":C.tealD,
              border:"0.5px solid "+C.teal,transition:"all 0.2s"}}>
            {copiadoId===s.id?"✓ Copiado!":"📋 Copiar"}
          </button>
        </div>
      ))}
    </div>
  );

  // ── RENDER PRINCIPAL ───────────────────────────────────────────────────────
  return (
    <div>
      {/* Abas internas */}
      <div style={{display:"flex",gap:0,marginBottom:16,borderBottom:"1px solid var(--color-border-tertiary)"}}>
        {(()=>{
        const assinantes = ritsCount !== null ? ritsCount
          : clientesDash.filter(c=>{
              if(!(c.statusClub==="fechou"||c.etapa==="experiencia")) return false;
              if(c.cancelado) return false;
              const st = c.statusAssinatura||"ativo";
              return st==="ativo"||st==="pausado";
            }).length;
        const meta100 = 100;
        const semanas = Math.max(1, Math.round((new Date("2026-12-31")-new Date())/604800000));
        const pct = Math.min(100, Math.round(assinantes/meta100*100));
        return assinantes>0&&(
          <div style={{background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 14px",marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
              <span style={{fontSize:12,fontWeight:500,color:"var(--color-text-primary)"}}>🏆 Meta 2026: {assinantes}/100 assinantes {ritsCount!==null&&<span style={{fontSize:10,color:C.tealD}}>(RitsPay)</span>}</span>
              <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>{Math.max(0,100-assinantes)} faltam · {semanas} semanas</span>
            </div>
            <div style={{height:6,background:"var(--color-border-tertiary)",borderRadius:3,overflow:"hidden"}}>
              <div style={{width:pct+"%",height:"100%",background:pct>=80?C.green:pct>=50?C.amber:C.teal,borderRadius:3,transition:"width 0.5s"}}/>
            </div>
          </div>
        );
      })()}
      {[["lista","📋 Lista"],["calendario","📅 Calendário"],["dash","📊 Dash"],["scripts","💬 Scripts"]].map(([id,label])=>(
          <button key={id} onClick={()=>setAba(id)}
            style={{padding:"8px 16px",fontSize:12,fontWeight:500,background:"none",border:"none",cursor:"pointer",
              color:aba===id?C.teal:"var(--color-text-secondary)",
              borderBottom:aba===id?"2px solid "+C.teal:"2px solid transparent",marginBottom:-1}}>
            {label}
          </button>
        ))}
      </div>

      {/* ABA LISTA */}
      {aba==="lista"&&(
        <div>
        {/* PAINEL DE URGÊNCIAS + PROGRESSO — substitui aba Hoje */}
        {(()=>{
          const prontas = clientes.filter(c=>(c.p||0)>=2&&(c._diasUlt||999)<90&&!c._muitoInativa);
          const nVencidos = prontas.filter(c=>c.statusClub&&c.proximoFollowup&&c.proximoFollowup<=hoje).length;
          const nImediata = prontas.filter(c=>["interessado","respondeu","link_enviado"].includes(c.statusClub)&&c.proximoFollowup&&c.proximoFollowup<=hoje).length;
          const nNovos = prontas.filter(c=>!c.statusClub&&(c._score||0)>=40).length;
          const hoje2 = new Date().toISOString().split("T")[0];
          const nRenovacao = clientesDash.filter(c=>c.etapa==="experiencia"&&c.dataInicioAssinatura&&(()=>{
            const d=new Date(c.dataInicioAssinatura+"T12:00:00");
            const dias=Math.round((new Date()-d)/86400000);
            return dias>=25&&dias<=35;
          })()).length;
          const nPausaVolta = clientesDash.filter(c=>c.statusAssinatura==="pausado"&&c.dataPausaFim&&c.dataPausaFim<=hoje2).length;
          // Meta anual
          const totalAssinantes = clientesDash.filter(c=>{
            if (!(c.statusClub==="fechou"||c.etapa==="experiencia")) return false;
            if (c.cancelado) return false;
            const st = c.statusAssinatura||"ativo";
            return st === "ativo" || st === "pausado"; // apenas ativos e pausados
          }).length;
          const meta = 100;
          const metaReal = ritsCount !== null ? ritsCount : totalAssinantes;
          const semanasFim = Math.max(1,Math.round((new Date("2026-12-31")-new Date())/604800000));
          const faltamReal = Math.max(0,meta-metaReal);
          const pctMeta = Math.min(100,Math.round(metaReal/meta*100));
          const novosContatadosHoje = prontas.filter(c=>!c.statusClub&&c.dataUltimoContato===hoje).length;
          const pctNovos = Math.min(100,Math.round(novosContatadosHoje/metaDiaria*100));
          const temAcao = nVencidos>0||nImediata>0||nNovos>0||nRenovacao>0;
          return (
            <div style={{marginBottom:12}}>
              {/* Meta + contatos do dia */}
              <div style={{display:"flex",gap:8,marginBottom:8}}>
                <div style={{flex:1,background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 12px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                    <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>Meta 2026 — Assinantes {ritsCount!==null?"(RitsPay)":"(CRM)"}</span>
                    <span style={{fontSize:11,fontWeight:600,color:pctMeta>=100?C.greenD:C.purple}}>{metaReal}/{meta}</span>
                  </div>
                  <div style={{height:5,background:"var(--color-border-tertiary)",borderRadius:3,overflow:"hidden",marginBottom:4}}>
                    <div style={{width:pctMeta+"%",height:"100%",background:pctMeta>=100?C.green:C.purple,borderRadius:3,transition:"width 0.3s"}}/>
                  </div>
                  <div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{faltamReal} faltam · {semanasFim} semanas</div>
                </div>
                <div style={{flex:1,background:"var(--color-background-secondary)",borderRadius:10,padding:"10px 12px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                    <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>Novos hoje</span>
                    <span style={{fontSize:11,fontWeight:600,color:pctNovos>=100?C.greenD:C.teal}}>{novosContatadosHoje}/{metaDiaria}</span>
                  </div>
                  <div style={{height:5,background:"var(--color-border-tertiary)",borderRadius:3,overflow:"hidden",marginBottom:4}}>
                    <div style={{width:pctNovos+"%",height:"100%",background:pctNovos>=100?C.green:C.teal,borderRadius:3,transition:"width 0.3s"}}/>
                  </div>
                  <div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{metaDiaria-novosContatadosHoje>0?metaDiaria-novosContatadosHoje+" restantes":""}</div>
                </div>
              </div>
              {/* Urgências clicáveis */}
              {temAcao&&<div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:8}}>
                {nImediata>0&&<button onClick={()=>{setFiltroStatus("interessado");}}
                  style={{flex:1,minWidth:90,background:C.coralL,border:"0.5px solid "+C.coral,borderRadius:8,padding:"8px 10px",cursor:"pointer",textAlign:"left"}}>
                  <div style={{fontSize:18,fontWeight:600,color:C.coralD}}>{nImediata}</div>
                  <div style={{fontSize:10,color:C.coralD}}>🔥 Ação imediata</div>
                </button>}
                {nVencidos>0&&<button onClick={()=>{setFiltroStatus("hoje");}}
                  style={{flex:1,minWidth:90,background:C.amberL,border:"0.5px solid "+C.amber,borderRadius:8,padding:"8px 10px",cursor:"pointer",textAlign:"left"}}>
                  <div style={{fontSize:18,fontWeight:600,color:C.amberD}}>{nVencidos}</div>
                  <div style={{fontSize:10,color:C.amberD}}>🔔 Follow-ups vencidos</div>
                </button>}
                {nRenovacao>0&&<button onClick={()=>setAba("dash")}
                  style={{flex:1,minWidth:90,background:C.greenL,border:"0.5px solid "+C.green,borderRadius:8,padding:"8px 10px",cursor:"pointer",textAlign:"left"}}>
                  <div style={{fontSize:18,fontWeight:600,color:C.greenD}}>{nRenovacao}</div>
                  <div style={{fontSize:10,color:C.greenD}}>🔄 Renovações próximas</div>
                </button>}
                {nNovos>0&&<button onClick={()=>setFiltroStatus("nao_abordado")}
                  style={{flex:1,minWidth:90,background:C.tealL,border:"0.5px solid "+C.teal,borderRadius:8,padding:"8px 10px",cursor:"pointer",textAlign:"left"}}>
                  <div style={{fontSize:18,fontWeight:600,color:C.tealD}}>{Math.min(nNovos,metaDiaria)}</div>
                  <div style={{fontSize:10,color:C.tealD}}>📤 Novos disponíveis</div>
                </button>}
                {nPausaVolta>0&&<button onClick={()=>setAba("dash")}
                  style={{flex:1,minWidth:90,background:C.greenL,border:"0.5px solid "+C.green,borderRadius:8,padding:"8px 10px",cursor:"pointer",textAlign:"left"}}>
                  <div style={{fontSize:18,fontWeight:600,color:C.greenD}}>{nPausaVolta}</div>
                  <div style={{fontSize:10,color:C.greenD}}>▶ Pausa terminou</div>
                </button>}
              </div>}
            </div>
          );
        })()}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          <div>
            <div style={{display:"flex",gap:8,marginBottom:8}}>
              <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="Buscar cliente..."
                style={{flex:1,padding:"7px 10px",borderRadius:8,border:"0.5px solid var(--color-border-tertiary)",fontSize:12,color:"var(--color-text-primary)",background:"var(--color-background-secondary)",outline:"none"}}/>
              <button onClick={()=>{setModoLote(!modoLote);setSelecionadosLote(new Set());}}
                style={{padding:"6px 10px",borderRadius:8,fontSize:11,cursor:"pointer",fontWeight:500,
                  background:modoLote?C.purple:"var(--color-background-secondary)",
                  color:modoLote?"#fff":"var(--color-text-secondary)",
                  border:"0.5px solid "+(modoLote?C.purple:"var(--color-border-tertiary)")}}>
                {modoLote?"✕ Cancelar":"☑ Lote"}
              </button>
            </div>
            {modoLote&&selecionadosLote.size>0&&(
              <div style={{background:C.purpleL,border:"0.5px solid "+C.purple,borderRadius:8,padding:"8px 12px",marginBottom:8,display:"flex",alignItems:"center",gap:8}}>
                <span style={{fontSize:12,color:C.purpleD,flex:1}}>{selecionadosLote.size} selecionado{selecionadosLote.size>1?"s":""}</span>
                <button onClick={async()=>{
                  const ids=[...selecionadosLote];
                  for(const id of ids){
                    const c=clientes.find(x=>x.id===id);
                    if(c&&!c.statusClub) await saveCliente({...c,statusClub:"contatado",dataAbordagem:hoje,dataUltimoContato:hoje,tentativasClub:(c.tentativasClub||0)+1,proximoFollowup:addDays(2)});
                  }
                  setSelecionadosLote(new Set()); setModoLote(false);
                }}
                  style={{padding:"5px 12px",borderRadius:6,fontSize:11,fontWeight:500,cursor:"pointer",background:C.purple,color:"#fff",border:"none"}}>
                  📤 Marcar como Contatado
                </button>
              </div>
            )}
            {/* Sorts cumulativos */}
            <div style={{display:"flex",gap:4,marginBottom:8,alignItems:"center"}}>
              <span style={{fontSize:10,color:"var(--color-text-tertiary)",textTransform:"uppercase",letterSpacing:"0.06em",marginRight:2}}>Ordenar:</span>
              {[
                {campo:"p", label:"Pedidos"},
                {campo:"ciclo", label:"Ciclo"},
                {campo:"dias", label:"Sem comprar"},
                {campo:"score", label:"Score"},
              ].map(({campo,label})=>{
                const active = sorts.find(s=>s.campo===campo);
                return (
                  <button key={campo} onClick={()=>toggleSort(campo)}
                    style={{padding:"3px 9px",borderRadius:20,fontSize:11,cursor:"pointer",fontWeight:active?500:400,
                      background:active?C.purple:"var(--color-background-secondary)",
                      color:active?"#fff":"var(--color-text-secondary)",
                      border:"0.5px solid "+(active?C.purple:"var(--color-border-tertiary)")}}>
                    {label}{active?(active.dir==="desc"?" ↓":" ↑"):""}
                  </button>
                );
              })}
              {sorts.length>0&&(
                <button onClick={()=>setSorts([])}
                  style={{padding:"3px 9px",borderRadius:20,fontSize:11,cursor:"pointer",
                    background:C.coralL,color:C.coralD,border:"0.5px solid "+C.coral}}>
                  ✕ Limpar
                </button>
              )}
            </div>
            <div style={{display:"flex",flexWrap:"wrap",gap:4,marginBottom:10}}>
              <button onClick={()=>setFiltroStatus("")}
                style={{padding:"3px 10px",borderRadius:20,fontSize:11,cursor:"pointer",
                  background:filtroStatus===""?C.teal:"var(--color-background-secondary)",
                  color:filtroStatus===""?"#fff":"var(--color-text-secondary)",
                  border:"0.5px solid "+(filtroStatus===""?C.teal:"var(--color-border-tertiary)")}}>
                Todos ({clientes.length})
              </button>
              <button onClick={()=>setFiltroStatus("hoje")}
                style={{padding:"3px 10px",borderRadius:20,fontSize:11,cursor:"pointer",
                  background:filtroStatus==="hoje"?C.coral:"var(--color-background-secondary)",
                  color:filtroStatus==="hoje"?"#fff":"var(--color-text-secondary)",
                  border:"0.5px solid "+(filtroStatus==="hoje"?C.coral:"var(--color-border-tertiary)")}}>
                ⚡ Hoje
              </button>
              <button onClick={()=>setFiltroStatus("nao_abordado")}
                style={{padding:"3px 10px",borderRadius:20,fontSize:11,cursor:"pointer",
                  background:filtroStatus==="nao_abordado"?C.teal:"var(--color-background-secondary)",
                  color:filtroStatus==="nao_abordado"?"#fff":"var(--color-text-secondary)",
                  border:"0.5px solid "+(filtroStatus==="nao_abordado"?C.teal:"var(--color-border-tertiary)")}}>
                ○ Não abordado
              </button>
              {STATUS_CLUB.filter(s=>s.id).map(s=>{
                const n=clientes.filter(c=>c.statusClub===s.id).length;
                if(n===0) return null;
                return (
                  <button key={s.id} onClick={()=>setFiltroStatus(s.id)}
                    style={{padding:"3px 10px",borderRadius:20,fontSize:11,cursor:"pointer",
                      background:filtroStatus===s.id?s.cor:"var(--color-background-secondary)",
                      color:filtroStatus===s.id?"#fff":"var(--color-text-secondary)",
                      border:"0.5px solid "+(filtroStatus===s.id?s.cor:"var(--color-border-tertiary)")}}>
                    {s.emoji} {s.label} ({n})
                  </button>
                );
              })}
            </div>
            <div style={{maxHeight:"65vh",overflowY:"auto"}}>
              {filtroStatus==="hoje"?(()=>{
                const prontas2 = clientes.filter(c=>(c.p||0)>=2&&(c._diasUlt||999)<90&&!c._muitoInativa);
                const acaoIm = prontas2.filter(c=>c.statusClub==="interessado"||c.statusClub==="respondeu").sort((a,b)=>a.statusClub==="interessado"?-1:1);
                const linksEnv = prontas2.filter(c=>c.statusClub==="link_enviado");
                const fuHoje = prontas2.filter(c=>c.statusClub&&c.statusClub!=="interessado"&&c.statusClub!=="respondeu"&&c.statusClub!=="link_enviado"&&c.statusClub!=="fechou"&&c.statusClub!=="perdido"&&c.proximoFollowup&&c.proximoFollowup<=hoje).sort((a,b)=>(a.proximoFollowup||"")>(b.proximoFollowup||"")?1:-1);
                const janelaH = prontas2.filter(c=>!c.statusClub&&c._diasParaProxima!=null&&c._diasParaProxima>=-2&&c._diasParaProxima<=5);
                const jIds = new Set(janelaH.map(c=>c.id));
                const novosH = prontas2.filter(c=>!c.statusClub&&!jIds.has(c.id)&&(c._score||0)>=40&&(c._diasUlt||999)<=60&&!(c._diasParaProxima!=null&&c._diasParaProxima>7)).sort((a,b)=>(b._score||0)-(a._score||0)).slice(0,Math.max(0,12-janelaH.length));
                const nada = acaoIm.length===0&&linksEnv.length===0&&fuHoje.length===0&&janelaH.length===0&&novosH.length===0;
                return (
                  <div>
                    {[
                      ["🔥 Ação imediata",acaoIm,C.coralD,C.coralL],
                      ["🔗 Link enviado",linksEnv,C.blueD,C.blueL],
                      ["📅 Follow-ups de hoje",fuHoje,C.amberD,C.amberL],
                      ["🛒 Janela de compra",janelaH,C.greenD,C.greenL],
                      ["📤 Novos contatos",novosH,C.tealD,C.tealL],
                    ].map(([titulo,lista,corD,corL])=>lista.length>0&&(
                      <div key={titulo} style={{marginBottom:12}}>
                        <div style={{fontSize:10,fontWeight:600,color:corD,background:corL,borderRadius:6,padding:"3px 8px",marginBottom:6,display:"inline-block"}}>{titulo} ({lista.length})</div>
                        {lista.map(c=><CardLista key={c.id} c={c}/>)}
                      </div>
                    ))}
                    {nada&&<div style={{textAlign:"center",padding:20,color:"var(--color-text-tertiary)",fontSize:12}}>✅ Nenhuma ação urgente hoje</div>}
                  </div>
                );
              })():(
                <>
                  {listaFiltrada.length===0&&<div style={{textAlign:"center",padding:20,color:"var(--color-text-tertiary)",fontSize:12}}>Nenhum cliente neste filtro</div>}
                  {listaFiltrada.map(c=><CardLista key={c.id} c={c}/>)}
                </>
              )}
            </div>
          </div>
          <div style={{maxHeight:"80vh",overflowY:"auto"}}>
            <PainelDetalhe/>
          </div>
        </div>
      </div>
      )}

      {/* ABA DASH */}
      {aba==="dash"&&(
        <div>
          <DashClub/>
          <div style={{marginTop:20,borderTop:"0.5px solid var(--color-border-tertiary)",paddingTop:16}}>
            <AnalyticsRitsPay onAbrirPerfil={onAbrirPerfil}/>
          </div>
        </div>
      )}

      {/* ABA CALENDÁRIO */}
      {aba==="calendario"&&<CalendarioFollowups clientes={clientesDash.length>0?clientesDash:[...clientes,...clientesDash].filter((c,i,arr)=>arr.findIndex(x=>x.id===c.id)===i)} onAbrirCliente={(c)=>{setSel(c);setScriptSel(sugerirScript(c));setAba("lista");}}/>}

      {/* ABA SCRIPTS */}
      {aba==="scripts"&&<Biblioteca/>}
      {showRitsSync&&<RitsPaySyncModal onClose={()=>setShowRitsSync(false)} onSyncDone={()=>{
        // Reload Club data after sync
        dbGetAll().then(lista => {
          const addScore = c => { const sc=calcScoreClub(c); return {...c,_score:sc.score,_diasUlt:sc.diasUlt,_diasParaProxima:sc.diasParaProxima,_inativa:sc.inativa,_muitoInativa:sc.muitoInativa}; };
          const todosSc = lista.map(c => fixCliente(addScore(c)));
          setTodosParaBusca(todosSc);
          const comScore = todosSc.filter(c => (c.p||0) >= 2);
          setClientes(comScore.filter(c=>c.etapa!=="experiencia"&&c.etapa!=="encerrado"&&c.statusClub!=="fechou"&&c.statusClub!=="perdido").sort((a,b)=>b._score-a._score));
          setClientesDash(comScore.filter(c=>c.statusClub||c.etapa==="experiencia"));
        });
      }}/>}
    </div>
  );
};


// ── AUTENTICAÇÃO ───────────────────────────────────────────────────────────
const hashSenha = async (senha) => {
  const enc = new TextEncoder().encode(senha);
  const buf = await crypto.subtle.digest("SHA-256", enc);
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,"0")).join("");
};

const dbGetUsuarios = async () => {
  try {
    const r = await sb("/usuarios?select=id,email,nome,perfil,senha_hash,ativo&order=email.asc");
    return r || [];
  } catch(e) { return []; }
};

const dbSaveUsuario = async (u) => {
  // Não envia abas_permitidas — campo de outro sistema, não mexer
  const payload = { id:u.id, email:u.email, nome:u.nome, perfil:u.perfil, senha_hash:u.senha_hash, ativo:u.ativo };
  await sb("/usuarios", { method:"POST", pref:"resolution=merge-duplicates", body:payload });
};

const dbDeleteUsuario = async (id) => {
  await sb("/usuarios?id=eq."+id, { method:"DELETE" });
};

const SESSION_KEY = "laricas_session";

const loadSession = () => {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch(e) { return null; }
};

const saveSession = (user) => {
  localStorage.setItem(SESSION_KEY, JSON.stringify(user));
};

const clearSession = () => {
  localStorage.removeItem(SESSION_KEY);
};

const Login = ({ onLogin }) => {
  const [usuario, setUsuario] = useState("");
  const [senha, setSenha] = useState("");
  const [erro, setErro] = useState("");
  const [carregando, setCarregando] = useState(false);

  const entrar = async () => {
    if (!usuario.trim() || !senha.trim()) { setErro("Preencha email e senha."); return; }
    setCarregando(true);
    setErro("");
    try {
      const usuarios = await dbGetUsuarios();
      const hash = await hashSenha(senha);
      const emailBusca = usuario.trim().toLowerCase();
      const u = usuarios.find(x => (x.email||"").toLowerCase() === emailBusca && x.ativo !== false);
      if (!u) {
        setErro("Usuário não encontrado ou inativo.");
        setCarregando(false);
        return;
      }
      if (u.senha_hash !== hash) {
        setErro("Senha incorreta. (Se o usuário foi criado em outro sistema, o algoritmo de hash pode ser diferente — avise o administrador.)");
        setCarregando(false);
        return;
      }
      const isAdminUser = (u.perfil||"").toLowerCase() === "admin";
      const sessao = { id:u.id, usuario:u.email, nome:u.nome, nivel: isAdminUser ? "admin" : "operador", perfilOriginal:u.perfil };
      saveSession(sessao);
      onLogin(sessao);
    } catch(e) {
      setErro("Erro ao conectar: "+(e.message||"tente novamente"));
    }
    setCarregando(false);
  };

  return (
    <div style={{ minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"var(--color-background-primary)" }}>
      <div style={{ width:340,padding:"32px 28px",background:"var(--color-background-secondary)",borderRadius:16,boxShadow:"0 4px 24px rgba(0,0,0,0.08)" }}>
        <div style={{ textAlign:"center",marginBottom:24 }}>
          <div style={{ fontSize:11,fontWeight:600,color:C.purple,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:4 }}>Laricas Fitness</div>
          <div style={{ fontSize:20,fontWeight:600,color:"var(--color-text-primary)" }}>CRM de Conversão</div>
        </div>
        <div style={{ marginBottom:12 }}>
          <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Email</div>
          <input value={usuario} onChange={e=>setUsuario(e.target.value)} onKeyDown={e=>e.key==="Enter"&&entrar()}
            style={inp()} placeholder="seu.email" autoFocus/>
        </div>
        <div style={{ marginBottom:16 }}>
          <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Senha</div>
          <input value={senha} onChange={e=>setSenha(e.target.value)} onKeyDown={e=>e.key==="Enter"&&entrar()}
            type="password" style={inp()} placeholder="••••••••"/>
        </div>
        {erro&&<div style={{ fontSize:12,color:C.coralD,background:C.coralL,borderRadius:8,padding:"8px 12px",marginBottom:12 }}>{erro}</div>}
        <button onClick={entrar} disabled={carregando}
          style={{ width:"100%",padding:"11px",borderRadius:10,fontSize:14,fontWeight:500,cursor:carregando?"default":"pointer",background:C.purple,color:"#fff",border:"none",opacity:carregando?0.6:1 }}>
          {carregando?"Entrando...":"Entrar"}
        </button>
      </div>
    </div>
  );
};

// Perfis do CRM Laricas — mapeiam para a coluna "perfil" compartilhada com outro sistema.
// Qualquer perfil diferente de "admin" é tratado como operador no CRM Laricas.
const PERFIS_CRM = [
  { id:"admin", label:"Admin", desc:"Acesso total ao CRM Laricas" },
  { id:"Operador", label:"Operador (CRM Laricas)", desc:"Sem acesso a Config, Backup e Importar" },
];
const NIVEIS = [
  { id:"admin", label:"Admin", desc:"Acesso total ao sistema" },
  { id:"operador", label:"Operador", desc:"Sem acesso a Config, Backup e Importar" },
];

const GerenciarUsuarios = () => {
  const [usuarios, setUsuarios] = useState([]);
  const [loading, setLoading] = useState(true);
  const [editando, setEditando] = useState(null);
  const [novaSenha, setNovaSenha] = useState("");
  const [salvando, setSalvando] = useState(false);
  const [erro, setErro] = useState("");
  const [ok, setOk] = useState("");

  const carregar = () => { dbGetUsuarios().then(u=>{setUsuarios(u);setLoading(false);}); };
  useEffect(carregar, []);

  const abrirNovo = () => { setEditando({ email:"", nome:"", perfil:"Operador", ativo:true }); setNovaSenha(""); setErro(""); };
  const abrirEdicao = (u) => { setEditando({...u}); setNovaSenha(""); setErro(""); };

  const salvar = async () => {
    if (!editando.email.trim() || !editando.nome.trim()) { setErro("Preencha email e nome."); return; }
    if (!editando.id && !novaSenha.trim()) { setErro("Defina uma senha para o novo usuário."); return; }
    setSalvando(true);
    setErro("");
    try {
      const payload = {
        id: editando.id || crypto.randomUUID(),
        email: editando.email.trim().toLowerCase(),
        nome: editando.nome.trim(),
        perfil: editando.perfil,
        ativo: editando.ativo !== false,
      };
      if (novaSenha.trim()) {
        payload.senha_hash = await hashSenha(novaSenha.trim());
      } else {
        payload.senha_hash = editando.senha_hash;
      }
      await dbSaveUsuario(payload);
      setOk("✓ Salvo!");
      setEditando(null);
      carregar();
      setTimeout(()=>setOk(""),2000);
    } catch(e) {
      setErro("Erro: "+(e.message||"tente novamente"));
    }
    setSalvando(false);
  };

  const remover = async (u) => {
    if (!window.confirm("Remover acesso de "+u.nome+"? Atenção: esta tabela é compartilhada com outro sistema — confirme que este usuário não é usado lá também.")) return;
    await dbDeleteUsuario(u.id);
    carregar();
  };

  if (loading) return <div style={{textAlign:"center",padding:40,color:"var(--color-text-tertiary)"}}>Carregando...</div>;

  return (
    <div style={{ maxWidth:600 }}>
      <div style={{ background:C.amberL,border:"0.5px solid "+C.amber,borderRadius:10,padding:"10px 14px",marginBottom:16,fontSize:12,color:C.amberD,lineHeight:1.5 }}>
        ⚠ Esta tabela de usuários é compartilhada com outro sistema (produção). Crie usuários novos para o CRM Laricas com cuidado — apenas o perfil "admin" dá acesso total aqui; qualquer outro perfil vira operador no CRM Laricas.
      </div>
      {ok&&<div style={{ fontSize:13,color:C.greenD,background:C.greenL,borderRadius:8,padding:"8px 12px",marginBottom:12 }}>{ok}</div>}

      {!editando&&(
        <div>
          <button onClick={abrirNovo}
            style={{ marginBottom:16,padding:"9px 18px",borderRadius:10,fontSize:13,fontWeight:500,cursor:"pointer",background:C.purple,color:"#fff",border:"none" }}>
            + Novo usuário
          </button>
          {usuarios.length===0&&<div style={{textAlign:"center",padding:30,color:"var(--color-text-tertiary)",fontSize:13}}>Nenhum usuário cadastrado ainda.</div>}
          {usuarios.map(u=>{
            const isAdminU = (u.perfil||"").toLowerCase()==="admin";
            return (
            <div key={u.id} style={{ display:"flex",alignItems:"center",gap:12,padding:"12px 14px",background:"var(--color-background-secondary)",borderRadius:10,marginBottom:8,opacity:u.ativo===false?0.5:1 }}>
              <div style={{ flex:1 }}>
                <div style={{ fontSize:14,fontWeight:500,color:"var(--color-text-primary)" }}>{u.nome}</div>
                <div style={{ fontSize:12,color:"var(--color-text-tertiary)" }}>{u.email}</div>
              </div>
              <span style={{ fontSize:11,fontWeight:500,color:isAdminU?C.purpleD:C.tealD,background:isAdminU?C.purpleL:C.tealL,padding:"2px 10px",borderRadius:20 }}>
                {u.perfil} {!isAdminU&&"(= operador no CRM)"}
              </span>
              {u.ativo===false&&<span style={{ fontSize:11,color:C.coralD,background:C.coralL,padding:"2px 10px",borderRadius:20 }}>Inativo</span>}
              <button onClick={()=>abrirEdicao(u)}
                style={{ padding:"5px 12px",borderRadius:8,fontSize:12,cursor:"pointer",background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-secondary)" }}>
                Editar
              </button>
              <button onClick={()=>remover(u)}
                style={{ padding:"5px 12px",borderRadius:8,fontSize:12,cursor:"pointer",background:"none",border:"0.5px solid "+C.coral,color:C.coralD }}>
                Remover
              </button>
            </div>
          );})}
        </div>
      )}

      {editando&&(
        <div style={{ background:"var(--color-background-secondary)",borderRadius:12,padding:"20px" }}>
          <div style={{ fontSize:14,fontWeight:500,color:"var(--color-text-primary)",marginBottom:16 }}>
            {editando.id?"Editar usuário":"Novo usuário"}
          </div>
          <div style={{ marginBottom:10 }}>
            <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Nome completo</div>
            <input value={editando.nome} onChange={e=>setEditando({...editando,nome:e.target.value})} style={inp()} placeholder="Maria Cecília"/>
          </div>
          <div style={{ marginBottom:10 }}>
            <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Email (login)</div>
            <input value={editando.email} onChange={e=>setEditando({...editando,email:e.target.value})} style={inp()} placeholder="ceci@laricas.com"/>
          </div>
          <div style={{ marginBottom:10 }}>
            <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>
              {editando.id?"Nova senha (deixe em branco para manter)":"Senha"}
            </div>
            <input value={novaSenha} onChange={e=>setNovaSenha(e.target.value)} type="password" style={inp()} placeholder="••••••••"/>
          </div>
          <div style={{ marginBottom:10 }}>
            <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.06em" }}>Perfil de acesso no CRM Laricas</div>
            <div style={{ display:"flex",gap:8 }}>
              {PERFIS_CRM.map(n=>(
                <button key={n.id} onClick={()=>setEditando({...editando,perfil:n.id})}
                  style={{ flex:1,padding:"10px",borderRadius:10,cursor:"pointer",textAlign:"left",
                    background:editando.perfil===n.id?C.purpleL:"var(--color-background-primary)",
                    border:"0.5px solid "+(editando.perfil===n.id?C.purple:"var(--color-border-tertiary)") }}>
                  <div style={{ fontSize:13,fontWeight:500,color:editando.perfil===n.id?C.purpleD:"var(--color-text-primary)" }}>{n.label}</div>
                  <div style={{ fontSize:11,color:"var(--color-text-tertiary)",marginTop:2 }}>{n.desc}</div>
                </button>
              ))}
            </div>
            <div style={{ fontSize:10,color:"var(--color-text-tertiary)",marginTop:6 }}>
              Atenção: este campo "perfil" é compartilhado com outro sistema. Use exatamente "admin" para acesso total, ou "Operador" para acesso restrito no CRM Laricas.
            </div>
          </div>
          {editando.id&&(
            <div style={{ marginBottom:14,display:"flex",alignItems:"center",gap:8 }}>
              <input type="checkbox" checked={editando.ativo!==false} onChange={e=>setEditando({...editando,ativo:e.target.checked})} id="ativoCheck"/>
              <label htmlFor="ativoCheck" style={{ fontSize:13,color:"var(--color-text-primary)",cursor:"pointer" }}>Usuário ativo</label>
            </div>
          )}
          {erro&&<div style={{ fontSize:12,color:C.coralD,background:C.coralL,borderRadius:8,padding:"8px 12px",marginBottom:12 }}>{erro}</div>}
          <div style={{ display:"flex",gap:8 }}>
            <button onClick={()=>setEditando(null)}
              style={{ flex:1,padding:"10px",borderRadius:10,fontSize:13,cursor:"pointer",background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-secondary)" }}>
              Cancelar
            </button>
            <button onClick={salvar} disabled={salvando}
              style={{ flex:2,padding:"10px",borderRadius:10,fontSize:13,fontWeight:500,cursor:salvando?"default":"pointer",background:C.purple,color:"#fff",border:"none",opacity:salvando?0.6:1 }}>
              {salvando?"Salvando...":"Salvar"}
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default function App() {
  const [tab,setTab]=useState("kanban");
  const [clienteId,setClienteId]=useState(null);
  const [refresh,setRefresh]=useState(0);
  const abrirClienteGlobal = (id) => { setClienteId(id); setTab("kanban"); };
  // Filtros persistidos fora do Kanban para sobreviver à navegação
  const [filtroHojeApp,setFiltroHojeApp]=useState(false);
  const [filtroClubApp,setFiltroClubApp]=useState(false);
  const [filtroProbApp,setFiltroProbApp]=useState("");
  const [filtroPedidosApp,setFiltroPedidosApp]=useState("");
  const [cfgOk,setCfgOk]=useState(false);
  const [cfgLoad,setCfgLoad]=useState(true);
  const [sessao,setSessao]=useState(null);
  const [sessaoLoad,setSessaoLoad]=useState(true);
  const [clubUrgencia,setClubUrgencia]=useState(0);
  useEffect(()=>{ loadCfg().then(cfg=>{ setCfgOk(!!(cfg.url&&cfg.key)); setCfgLoad(false); }); },[]);
  useEffect(()=>{ if(cfgOk){ setSessao(loadSession()); setSessaoLoad(false); } },[cfgOk]);
  const onSalvo = () => { setTab("kanban"); setRefresh(r=>r+1); };
  const onRestore = () => { setTab("kanban"); setRefresh(r=>r+1); };
  const sair = () => { clearSession(); setSessao(null); };
  const isAdmin = sessao?.nivel === "admin";

  if (cfgLoad) return <div style={{ textAlign:"center",padding:"60px 0",color:"var(--color-text-tertiary)",fontFamily:"var(--font-sans)" }}>Carregando...</div>;
  if (!cfgOk) return (
    <div style={{ maxWidth:900,margin:"0 auto",padding:"0 20px 40px",fontFamily:"var(--font-sans)",color:"var(--color-text-primary)" }}>
      <ConfigSupabase onSalvo={()=>{ loadCfg().then(()=>setCfgOk(true)); }}/>
    </div>
  );
  if (sessaoLoad) return <div style={{ textAlign:"center",padding:"60px 0",color:"var(--color-text-tertiary)",fontFamily:"var(--font-sans)" }}>Carregando...</div>;
  if (!sessao) return <Login onLogin={setSessao}/>;

  return (
    <div style={{ maxWidth:900,margin:"0 auto",padding:"0 0 40px",fontFamily:"var(--font-sans)",color:"var(--color-text-primary)" }}>
      <div style={{ display:"flex",alignItems:"center",gap:12,padding:"20px 0 8px" }}>
        <div style={{ flex:1 }}>
        <div style={{ fontSize:11,fontWeight:500,letterSpacing:"0.09em",textTransform:"uppercase",color:"var(--color-text-tertiary)",marginBottom:4 }}>Laricas Fitness</div>
        <div style={{ fontSize:22,fontWeight:500,lineHeight:1.3 }}>CRM de Conversão</div>
          <div style={{ fontSize:13,color:"var(--color-text-secondary)",marginTop:4 }}>Lead → Contato → Conversa → Proposta → Convertido</div>
        </div>
        <GlobalSearch onAbrir={abrirClienteGlobal}/>
        <div style={{ display:"flex",alignItems:"center",gap:8,flexShrink:0 }}>
          <div style={{ textAlign:"right" }}>
            <div style={{ fontSize:12,fontWeight:500,color:"var(--color-text-primary)" }}>{sessao.nome}</div>
            <div style={{ fontSize:10,color:"var(--color-text-tertiary)",textTransform:"uppercase" }}>{NIVEIS.find(n=>n.id===sessao.nivel)?.label}</div>
          </div>
          <button onClick={sair} title="Sair"
            style={{ padding:"6px 10px",borderRadius:8,fontSize:12,cursor:"pointer",background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-secondary)" }}>
            Sair
          </button>
        </div>
      </div>
      <div style={{ display:"flex",borderBottom:"0.5px solid var(--color-border-tertiary)",marginBottom:24,overflowX:"auto" }}>
        <T label="📋 Kanban" active={tab==="kanban"} color={C.green} onClick={()=>{setClienteId(null);setTab("kanban");}}/>
        <T label={"🎯 Club"+(clubUrgencia>0?" ("+clubUrgencia+")":"")} active={tab==="club"} color={C.teal} onClick={()=>setTab("club")}/>
        {isAdmin&&<T label="📥 Importar" active={tab==="import"} color={C.purple} onClick={()=>setTab("import")}/>}
        <button onClick={()=>setShowNovoLead(true)}
          style={{padding:"6px 14px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",
            background:C.teal,color:"#fff",border:"none",display:"flex",alignItems:"center",gap:5}}>
          ➕ Novo Lead
        </button>
        <T label="📊 Historico" active={tab==="historico"} color={C.teal} onClick={()=>setTab("historico")}/>
        <T label="🔗 Unificar" active={tab==="unificar"} color={C.amber} onClick={()=>setTab("unificar")}/>
        <T label="📖 Guia" active={tab==="guia"} color={C.teal} onClick={()=>setTab("guia")}/>
        {isAdmin&&<T label="💾 Backup" active={tab==="backup"} color={C.blue} onClick={()=>setTab("backup")}/>}
        {isAdmin&&<T label="👥 Usuários" active={tab==="usuarios"} color={C.purple} onClick={()=>setTab("usuarios")}/>}
        {isAdmin&&<T label="⚙ Config" active={tab==="config"} color="var(--color-text-tertiary)" onClick={()=>setTab("config")}/>}
      </div>
      {tab==="club"&&<FunilClub onAbrirPerfil={(id)=>{abrirClienteGlobal(id);setTab("kanban");}} onUrgencia={setClubUrgencia}/>}
      {tab==="kanban"&&(clienteId?<Perfil key={clienteId} clienteId={clienteId} onVoltar={()=>{setClienteId(null);setRefresh(r=>r+1);}}/>:
          <Kanban onAbrir={setClienteId} reloadToken={refresh}
            filtroHoje={filtroHojeApp} setFiltroHoje={setFiltroHojeApp}
            filtroClub={filtroClubApp} setFiltroClub={setFiltroClubApp}
            filtroProb={filtroProbApp} setFiltroProb={setFiltroProbApp}
            filtroPedidos={filtroPedidosApp} setFiltroPedidos={setFiltroPedidosApp}
          />)}
      {tab==="import"&&isAdmin&&<ImportarLista onSalvo={onSalvo}/>}
      {showNovoLead&&<NovoLeadModal onClose={()=>setShowNovoLead(false)} onSalvo={()=>{setShowNovoLead(false);}}/>}
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
      {tab==="unificar"&&<Unificar onSalvo={()=>setRefresh(r=>r+1)}/>}
      {tab==="guia"&&<Guia/>}
      {tab==="backup"&&isAdmin&&<Backup onRestore={onRestore}/>}
      {tab==="usuarios"&&isAdmin&&<GerenciarUsuarios/>}
      {tab==="config"&&isAdmin&&<ConfigSupabase onSalvo={()=>{ loadCfg().then(()=>setCfgOk(true)); setTab("kanban"); }}/>}
    </div>
  );
}

