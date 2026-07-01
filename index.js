const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const path = require("path");

// PALETA de cores (mantida)
const C = {
  navyDark:   "1A2B5E",
  navy:       "2D3D70",
  coral:      "F4614D",
  coralLight: "F9836F",
  white:      "FFFFFF",
  offWhite:   "F5F7FA",
  lightGray:  "E8ECF2",
  midGray:    "8A95A8",
  darkText:   "1E2235",
  teal:       "0D9488",
  amber:      "F59E0B",
  green:      "22C55E",
  purple:     "7C3AED",
};

// Ícones (sem sharp, SVG direto)
const { FaBrain, FaCogs, FaShieldAlt, FaChartLine, FaUsers,
        FaCheckCircle, FaDatabase, FaCode, FaExpandArrowsAlt } = require("react-icons/fa");

function renderSvg(IconComp, color = "#FFFFFF", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComp, { color, size: String(size) })
  );
}

async function icon64(IconComp, color = "#FFFFFF", size = 256) {
  const svg = renderSvg(IconComp, color, size);
  const base64 = Buffer.from(svg).toString("base64");
  return `data:image/svg+xml;base64,${base64}`;
}

// ASSETS (logos)
const AURA_LOGO   = path.join(__dirname, "rh", "aura_logo.png");
const INOVA_LOGO  = path.join(__dirname, "assets", "image10.png");

function addSlideFooter(slide) {
  slide.addShape("rect", {
    x: 0, y: 5.35, w: 10, h: 0.03,
    fill: { color: C.coral }, line: { color: C.coral }
  });
  slide.addImage({ path: AURA_LOGO,  x: 0.3,  y: 5.2, w: 0.9, h: 0.33 });
  slide.addImage({ path: INOVA_LOGO, x: 8.75, y: 5.17, w: 0.95, h: 0.37 });
}

function addSlideHeader(slide, label = "APOENA") {
  slide.addText(label, {
    x: 0.35, y: 0.13, w: 2, h: 0.28,
    fontSize: 9, bold: true, color: C.coral,
    fontFace: "Calibri", margin: 0, charSpacing: 2
  });
}

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title  = "Storytech: Extrator de Relatórios — Apoena";
  pres.author = "Déborah Giovana de Andrade Gonçalves";

  // ----------------------- SLIDE 1 - CAPA ----------------------- //
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDark };
    // decoração (círculos)
    s.addShape("ellipse", { x: 7.8, y: -1.2, w: 4.5, h: 4.5, fill: { color: C.coral, transparency: 82 } });
    s.addShape("ellipse", { x: 8.4, y: -0.5, w: 3.0, h: 3.0, fill: { color: C.coral, transparency: 70 } });
    s.addShape("ellipse", { x: -1.0, y: 3.8, w: 3.5, h: 3.5, fill: { color: C.navy, transparency: 40 } });
    
    s.addShape("roundRect", { x: 0.5, y: 0.42, w: 1.9, h: 0.32, fill: { color: C.coral }, rectRadius: 0.07 });
    s.addText("PROGRAMA DE INOVAÇÃO", { x: 0.5, y: 0.42, w: 1.9, h: 0.32, fontSize: 7.5, bold: true, color: C.white, align: "center", valign: "middle", fontFace: "Calibri" });

    s.addText("EXTRATOR UNIVERSAL\nDE DADOS", { x: 0.5, y: 1.0, w: 7.5, h: 1.85, fontSize: 46, bold: true, color: C.white, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.05 });
    s.addText("De horas de digitação manual a dados estruturados em 30 segundos.", { x: 0.5, y: 2.9, w: 7.2, h: 0.45, fontSize: 14, color: "B8C4DC", fontFace: "Calibri", italic: true });

    // Badges ajustados (sem alegação local/LGPD)
    const badges = [
      { label: "99,8% redução", sub: "de tempo" },
      { label: "Erro zero",     sub: "na extração" },
      { label: "Custo zero",    sub: "de licença" },
      { label: "Segurança",     sub: "corporativa" }
    ];
    const bx = [0.5, 2.8, 5.1, 7.4];
    for (let i = 0; i < badges.length; i++) {
      s.addShape("roundRect", { x: bx[i], y: 3.55, w: 2.1, h: 0.72, fill: { color: "FFFFFF", transparency: 88 }, line: { color: C.coral, width: 1.2 }, rectRadius: 0.08 });
      s.addText([
        { text: badges[i].label, options: { bold: true, fontSize: 13, color: C.white, breakLine: true } },
        { text: badges[i].sub,   options: { fontSize: 9,  color: "B8C4DC" } }
      ], { x: bx[i], y: 3.55, w: 2.1, h: 0.72, align: "center", valign: "middle", fontFace: "Calibri", margin: 4 });
    }

    s.addText("Déborah Gonçalves  ·  Célula de Gestão de Contratos  ·  Aura Minerals — Apoena  ·  GAFI  ·  Gestor: Cleisson Bento", { x: 0.5, y: 4.45, w: 9.0, h: 0.3, fontSize: 9, color: "7A8BAE", fontFace: "Calibri" });
    s.addImage({ path: AURA_LOGO, x: 0.4, y: 4.9, w: 1.1, h: 0.4 });
    s.addImage({ path: INOVA_LOGO, x: 8.7, y: 4.87, w: 0.9, h: 0.36 });
  }

  // ----------------------- SLIDE 2 - PROBLEMA x SOLUÇÃO ----------------------- //
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    addSlideHeader(s);
    addSlideFooter(s);
    s.addText("O Problema: Processo Manual Insustentável", { x: 0.35, y: 0.45, w: 9.3, h: 0.55, fontSize: 22, bold: true, color: C.navyDark });

    // ANTES
    s.addShape("rect", { x: 0.3, y: 1.1, w: 4.5, h: 3.85, fill: { color: C.navyDark } });
    s.addShape("rect", { x: 0.3, y: 1.1, w: 4.5, h: 0.45, fill: { color: "C0392B" } });
    s.addText("🔴  ANTES DA SOLUÇÃO", { x: 0.3, y: 1.1, w: 4.5, h: 0.45, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" });
    const before = [
      ["⏱", "8 a 16 horas por contrato"],
      ["📄", "3 contratos/mês para processar"],
      ["😫", "1 analista 100% dedicado"],
      ["❌", "Erros de digitação frequentes"],
      ["⚠️", "Risco de multas e retrabalho"],
      ["📉", "Processo manual e desgastante"],
    ];
    before.forEach(([ico, txt], i) => {
      s.addText(`${ico}  ${txt}`, { x: 0.5, y: 1.65 + i * 0.5, w: 4.1, h: 0.4, fontSize: 11, color: "D0D8E8", fontFace: "Calibri", valign: "middle" });
    });

    // DEPOIS (sem menção a local/LGPD)
    s.addShape("rect", { x: 5.2, y: 1.1, w: 4.5, h: 3.85, fill: { color: C.offWhite }, line: { color: C.lightGray, width: 1 } });
    s.addShape("rect", { x: 5.2, y: 1.1, w: 4.5, h: 0.45, fill: { color: C.teal } });
    s.addText("🟢  DEPOIS DA SOLUÇÃO", { x: 5.2, y: 1.1, w: 4.5, h: 0.45, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" });
    const after = [
      ["⚡", "30 segundos por contrato"],
      ["📊", "Dados estruturados automaticamente"],
      ["✅", "Zero erros de transcrição"],
      ["🔒", "Processamento seguro e auditável"],
      ["💰", "Custo zero de licenciamento"],
      ["🚀", "Analista liberado para análise estratégica"],
    ];
    after.forEach(([ico, txt], i) => {
      s.addText(`${ico}  ${txt}`, { x: 5.35, y: 1.65 + i * 0.5, w: 4.2, h: 0.4, fontSize: 11, color: C.darkText, fontFace: "Calibri", valign: "middle" });
    });
    s.addShape("ellipse", { x: 4.6, y: 2.55, w: 0.6, h: 0.6, fill: { color: C.coral } });
    s.addText("VS", { x: 4.6, y: 2.55, w: 0.6, h: 0.6, fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle" });
  }

  // ----------------------- SLIDE 3 - TRÊS PILARES ----------------------- //
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDark };
    addSlideHeader(s);
    addSlideFooter(s);
    s.addText("A Inovação: Motor Inteligente de Interpretação", { x: 0.4, y: 0.45, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.white });
    s.addText("Não é apenas um conversor de PDF. É um sistema que lê, interpreta e organiza dados com inteligência cirúrgica.", { x: 0.4, y: 1.0, w: 9.2, h: 0.4, fontSize: 11.5, color: "A0AABB", italic: true });
    const pillars = [
      { icon: FaBrain, color: C.coral, bg: "24304A", title: "INTELIGÊNCIA",   lines: ["Interpretação contextual", "baseada em regras", "adaptáveis por área"] },
      { icon: FaCogs,  color: C.teal,  bg: "1B3040", title: "AUTOMAÇÃO",      lines: ["Processamento em lote", "Arraste + clique =", "dados prontos no Excel"] },
      { icon: FaShieldAlt, color: C.amber, bg: "2E2A1A", title: "SEGURANÇA",    lines: ["Políticas de acesso", "Trilha de auditoria", "Governança de dados"] } // substituído
    ];
    for (let i = 0; i < pillars.length; i++) {
      const p = pillars[i];
      const x = 0.4 + i * 3.1;
      s.addShape("roundRect", { x, y: 1.55, w: 2.9, h: 3.35, fill: { color: p.bg }, line: { color: p.color, width: 1.5 }, rectRadius: 0.12 });
      const iconData = await icon64(p.icon, `#${p.color}`, 256);
      s.addImage({ data: iconData, x: x + 1.1, y: 1.72, w: 0.7, h: 0.7 });
      s.addText(p.title, { x, y: 2.52, w: 2.9, h: 0.38, fontSize: 13, bold: true, color: p.color, align: "center" });
      p.lines.forEach((line, li) => {
        s.addText(line, { x: x + 0.15, y: 2.98 + li * 0.38, w: 2.6, h: 0.36, fontSize: 11, color: "C8D0DC", align: "center" });
      });
    }
    s.addShape("roundRect", { x: 0.4, y: 5.0, w: 9.2, h: 0.22, fill: { color: C.coral, transparency: 85 }, line: { color: C.coral, width: 1 }, rectRadius: 0.04 });
    s.addText("Inteligência Cirúrgica: a inovação está em escolher exatamente o que absorver — não em apenas converter.", { x: 0.4, y: 5.0, w: 9.2, h: 0.22, fontSize: 8.5, bold: true, color: C.coral, align: "center", valign: "middle" });
  }

  // ----------------------- SLIDE 4 - EVIDÊNCIAS (KPIs + Arquitetura) ----------------------- //
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    addSlideHeader(s);
    addSlideFooter(s);
    s.addText("Evidências: Resultados Reais", { x: 0.35, y: 0.45, w: 9.3, h: 0.55, fontSize: 22, bold: true, color: C.navyDark });
    const kpis = [
      { num: "99,8%", label: "Redução de\ntempo",       color: C.coral  },
      { num: "47h",   label: "Economia\npor mês",       color: C.navy   },
      { num: "ZERO",  label: "Erros de\ntranscrição",   color: C.teal   },
      { num: "R$ 0",  label: "Custo de\nlicença",       color: C.amber  },
    ];
    kpis.forEach((k, i) => {
      const x = 0.35 + i * 2.35;
      s.addShape("rect", { x, y: 1.05, w: 0.12, h: 0.95, fill: { color: k.color } });
      s.addText(k.num, { x: x + 0.2, y: 1.05, w: 2.0, h: 0.55, fontSize: 28, bold: true, color: k.color, valign: "top" });
      s.addText(k.label, { x: x + 0.2, y: 1.6, w: 2.0, h: 0.4, fontSize: 10, color: C.midGray });
    });
    s.addShape("rect", { x: 0.35, y: 2.1, w: 9.3, h: 0.02, fill: { color: C.lightGray } });

    // Core tecnológico
    s.addShape("roundRect", { x: 0.35, y: 2.25, w: 4.4, h: 2.7, fill: { color: C.navyDark }, rectRadius: 0.1 });
    s.addText("O Motor de Inteligência (Core)", { x: 0.35, y: 2.3, w: 4.4, h: 0.38, fontSize: 12, bold: true, color: C.coral, align: "center" });
    const coreItems = [
      ["🔍", "pdfplumber", "OCR e extração"],
      ["⚙️", "Regex Engine", "Regras de negócio"],
      ["📊", "Pandas", "Estruturação de dados"],
    ];
    coreItems.forEach(([ico, lib, desc], i) => {
      s.addShape("roundRect", { x: 0.55, y: 2.78 + i * 0.62, w: 4.0, h: 0.52, fill: { color: "2A3B5E" }, line: { color: "3A5080", width: 1 }, rectRadius: 0.06 });
      s.addText(`${ico}  ${lib}`, { x: 0.7, y: 2.82 + i * 0.62, w: 1.7, h: 0.42, fontSize: 11, bold: true, color: C.white });
      s.addText(desc, { x: 2.5, y: 2.82 + i * 0.62, w: 2.0, h: 0.42, fontSize: 10, color: "8AAABB" });
    });

    // Workflow (sem menção a local)
    s.addShape("roundRect", { x: 5.1, y: 2.25, w: 4.55, h: 2.7, fill: { color: C.offWhite }, line: { color: C.lightGray, width: 1 }, rectRadius: 0.1 });
    s.addText("Workflow Automatizado", { x: 5.1, y: 2.3, w: 4.55, h: 0.38, fontSize: 12, bold: true, color: C.navyDark, align: "center" });
    const wfSteps = [
      { n: "1", label: "Ingestão em Lote", sub: "Arraste PDFs", color: C.coral },
      { n: "2", label: "Limpeza (Regex)",   sub: "Regras aplicadas", color: C.navy },
      { n: "3", label: "Exportação .xlsx",  sub: "Pronto para ERP", color: C.teal },
    ];
    wfSteps.forEach((w, i) => {
      const wx = 5.25 + i * 1.47;
      s.addShape("ellipse", { x: wx + 0.33, y: 2.78, w: 0.52, h: 0.52, fill: { color: w.color } });
      s.addText(w.n, { x: wx + 0.33, y: 2.78, w: 0.52, h: 0.52, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle" });
      s.addText(w.label, { x: wx, y: 3.4, w: 1.3, h: 0.32, fontSize: 9.5, bold: true, color: C.navyDark, align: "center" });
      s.addText(w.sub, { x: wx, y: 3.72, w: 1.3, h: 0.28, fontSize: 8.5, color: C.midGray, align: "center" });
      if (i < 2) {
        s.addShape("rect", { x: wx + 1.22, y: 2.99, w: 0.25, h: 0.08, fill: { color: C.midGray } });
      }
    });
    s.addShape("roundRect", { x: 5.25, y: 4.1, w: 4.3, h: 0.72, fill: { color: "E8F0FF" }, line: { color: C.navy, width: 1 }, rectRadius: 0.07 });
    s.addText([
      { text: "Interface: ", options: { bold: true, color: C.navy } },
      { text: "Streamlit (MVP de validação) — versão de demonstração.", options: { color: C.navyDark } },
      { text: "\nPronto para integração com sistemas corporativos.", options: { color: C.midGray, italic: true } },
    ], { x: 5.4, y: 4.13, w: 4.0, h: 0.66, fontSize: 9.5, fontFace: "Calibri", valign: "middle" });
    s.addText("📅  Implantado em 06/04/2026  ·  Célula de Gestão de Contratos  ·  Gestor: Cleisson Bento", { x: 0.35, y: 5.0, w: 9.3, h: 0.22, fontSize: 8.5, color: C.midGray, italic: true });
  }

  // ----------------------- SLIDE 5 - ESCALABILIDADE ----------------------- //
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDark };
    addSlideHeader(s);
    addSlideFooter(s);
    s.addText("Escalabilidade: Um Motor, Múltiplas Áreas", { x: 0.4, y: 0.45, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.white });
    s.addText("O mesmo motor de extração, adaptado para diferentes documentos em toda a companhia.", { x: 0.4, y: 1.02, w: 9.2, h: 0.35, fontSize: 11.5, color: "A0AABB", italic: true });
    const phases = [
      { fase: "FASE 1", status: "✅ IMPLANTADO", statusColor: C.teal, area: "Gestão de Contratos", docs: "3 contratos/mês", icon: FaCheckCircle, iconColor: C.teal, bg: "1E3B2A", border: C.teal },
      { fase: "FASE 2", status: "🔜 MAPEADO", statusColor: C.amber, area: "Jurídico", docs: "Petições, Sentenças", icon: FaDatabase, iconColor: C.amber, bg: "3A2E0A", border: C.amber },
      { fase: "FASE 3", status: "📋 PLANEJADO", statusColor: "7B8FF8", area: "RH", docs: "Currículos, Fichas", icon: FaUsers, iconColor: "7B8FF8", bg: "1A1E3A", border: "7B8FF8" },
      { fase: "FASE 4", status: "🚚 PLANEJADO", statusColor: C.coralLight, area: "Logística", docs: "Notas Fiscais", icon: FaDatabase, iconColor: C.coralLight, bg: "2E1A18", border: C.coralLight },
    ];
    for (let i = 0; i < phases.length; i++) {
      const p = phases[i];
      const x = 0.35 + i * 2.38;
      s.addShape("roundRect", { x, y: 1.5, w: 2.2, h: 3.35, fill: { color: p.bg }, line: { color: p.border, width: 1.5 }, rectRadius: 0.1 });
      s.addShape("roundRect", { x, y: 1.5, w: 2.2, h: 0.38, fill: { color: p.border }, rectRadius: 0.05 });
      s.addText(p.fase, { x, y: 1.5, w: 2.2, h: 0.38, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" });
      const iconData = await icon64(p.icon, `#${p.iconColor}`, 256);
      s.addImage({ data: iconData, x: x + 0.75, y: 2.0, w: 0.7, h: 0.7 });
      s.addText(p.area, { x: x + 0.1, y: 2.78, w: 2.0, h: 0.38, fontSize: 12, bold: true, color: C.white, align: "center" });
      s.addText(p.docs, { x: x + 0.1, y: 3.2, w: 2.0, h: 0.55, fontSize: 9.5, color: "A0B0C0", align: "center" });
      s.addShape("roundRect", { x: x + 0.25, y: 3.85, w: 1.7, h: 0.3, fill: { color: p.border, transparency: 75 }, line: { color: p.border, width: 1 }, rectRadius: 0.06 });
      s.addText(p.status, { x: x + 0.25, y: 3.85, w: 1.7, h: 0.3, fontSize: 8.5, bold: true, color: p.statusColor, align: "center", valign: "middle" });
    }
    s.addShape("roundRect", { x: 0.35, y: 5.0, w: 9.3, h: 0.3, fill: { color: C.coral, transparency: 80 }, line: { color: C.coral, width: 1 }, rectRadius: 0.06 });
    s.addText("🎯  Objetivo final: HUB de Automação Documental da Aura Minerals", { x: 0.35, y: 5.0, w: 9.3, h: 0.3, fontSize: 10, bold: true, color: C.coral, align: "center", valign: "middle" });
  }

  // ----------------------- SLIDE 6 - IMPACTO ESTRATÉGICO ----------------------- //
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    addSlideHeader(s);
    addSlideFooter(s);
    s.addText("Impacto Estratégico", { x: 0.35, y: 0.45, w: 9.3, h: 0.55, fontSize: 22, bold: true, color: C.navyDark });
    const quads = [
      { title: "💰 Economia Direta", color: C.teal, bg: C.white, items: ["47 horas economizadas/mês", "75% da jornada liberada", "Custo zero de licença", "ROI imediato"] },
      { title: "🎯 Qualidade & Governança", color: C.navy, bg: C.white, items: ["Erro zero em transcrições", "Dados padronizados", "Rastreabilidade total", "Auditoria automática"] },
      { title: "📈 Capacidade Operacional", color: C.coral, bg: C.white, items: ["Analista focado em análise", "Resposta ágil", "Engajamento da equipe", "Processo sem gargalo"] },
      { title: "🚀 Potencial de Crescimento", color: C.purple, bg: C.white, items: ["Motor adaptável por domínio", "Expansão para Jurídico, RH, Log.", "Plataforma reutilizável", "Tecnologia aberta"] },
    ];
    const positions = [{ x: 0.3, y: 1.1 }, { x: 5.1, y: 1.1 }, { x: 0.3, y: 3.0 }, { x: 5.1, y: 3.0 }];
    for (let i = 0; i < quads.length; i++) {
      const q = quads[i];
      const { x, y } = positions[i];
      s.addShape("roundRect", { x, y, w: 4.55, h: 1.75, fill: { color: q.bg }, line: { color: C.lightGray, width: 1 }, rectRadius: 0.1 });
      s.addShape("rect", { x, y: y + 0.12, w: 0.08, h: 1.5, fill: { color: q.color } });
      s.addText(q.title, { x: x + 0.2, y: y + 0.1, w: 4.2, h: 0.35, fontSize: 12, bold: true, color: q.color });
      q.items.forEach((item, ii) => {
        s.addText(`• ${item}`, { x: x + 0.2, y: y + 0.48 + ii * 0.3, w: 4.2, h: 0.28, fontSize: 10, color: C.darkText });
      });
    }
  }

  // ----------------------- SLIDE 7 - VISÃO DE FUTURO ----------------------- //
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDark };
    addSlideHeader(s);
    addSlideFooter(s);
    s.addText("Visão de Futuro: Expansão Multissetorial", { x: 0.4, y: 0.45, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.white });
    const steps = [
      { n: "01", title: "Transformação em API", desc: "Desacoplar o motor Python para integração com sistemas corporativos.", icon: FaCode, color: C.coral },
      { n: "02", title: "Integração Corporativa", desc: "Conectar o extrator a ferramentas homologadas (Power Apps, SAP).", icon: FaExpandArrowsAlt, color: C.teal },
      { n: "03", title: "Escala Operacional", desc: "Automatizar processos do RH, Jurídico e Logística com o mesmo motor.", icon: FaChartLine, color: C.amber },
    ];
    for (let i = 0; i < steps.length; i++) {
      const st = steps[i];
      const x = 0.4 + i * 3.1;
      s.addShape("roundRect", { x, y: 1.15, w: 2.9, h: 2.95, fill: { color: "1E2B48" }, line: { color: st.color, width: 1.5 }, rectRadius: 0.1 });
      s.addText(st.n, { x, y: 1.2, w: 2.9, h: 0.55, fontSize: 26, bold: true, color: `${st.color}`, align: "center" });
      const iconData = await icon64(st.icon, `#${st.color}`, 256);
      s.addImage({ data: iconData, x: x + 1.1, y: 1.85, w: 0.7, h: 0.7 });
      s.addText(st.title, { x: x + 0.1, y: 2.65, w: 2.7, h: 0.38, fontSize: 12, bold: true, color: C.white, align: "center" });
      s.addText(st.desc, { x: x + 0.15, y: 3.08, w: 2.6, h: 0.9, fontSize: 9.5, color: "A0AABB", align: "center" });
    }
    s.addShape("roundRect", { x: 0.4, y: 4.22, w: 9.2, h: 0.78, fill: { color: C.coral, transparency: 88 }, line: { color: C.coral, width: 1.5 }, rectRadius: 0.1 });
    s.addText([
      { text: "Projeto implantado em 06/04/2026. Resultados comprovados. Escalabilidade mapeada.\n", options: { bold: true, fontSize: 12, color: C.white } },
      { text: "Pronto para ser o próximo ativo estratégico da Aura Minerals.", options: { fontSize: 11, color: "D0C0FF", italic: true } }
    ], { x: 0.6, y: 4.28, w: 8.8, h: 0.66, align: "center", valign: "middle" });
  }

  // ----------------------- SLIDE 8 - ENCERRAMENTO ----------------------- //
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDark };
    s.addShape("ellipse", { x: 6.5, y: -1.5, w: 6.5, h: 6.5, fill: { color: C.coral, transparency: 88 } });
    s.addShape("ellipse", { x: 7.2, y: -0.8, w: 4.5, h: 4.5, fill: { color: C.coral, transparency: 78 } });
    s.addText("INOVE\nSEMPRE!", { x: 0.6, y: 0.7, w: 5.5, h: 2.0, fontSize: 56, bold: true, color: C.white, valign: "top", lineSpacingMultiple: 1.0 });
    s.addText("Inovar é Escalar", { x: 0.6, y: 2.75, w: 5.5, h: 0.5, fontSize: 18, color: C.coral, bold: true });
    s.addText("Vamos construir o futuro juntos?", { x: 0.6, y: 3.3, w: 5.5, h: 0.4, fontSize: 14, color: "A0B0C0", italic: true });
    s.addShape("roundRect", { x: 0.6, y: 3.9, w: 5.5, h: 1.2, fill: { color: "1E2B48" }, line: { color: C.coral, width: 1.2 }, rectRadius: 0.1 });
    s.addText([
      { text: "Déborah Giovana de Andrade Gonçalves\n", options: { bold: true, fontSize: 12, color: C.white } },
      { text: "Proponente · Extrator de Relatórios — Apoena\n", options: { fontSize: 10, color: "A0AABB" } },
      { text: "Célula de Gestão de Contratos  ·  GAFI  ·  Aura Minerals", options: { fontSize: 9.5, color: "6A7A8A" } }
    ], { x: 0.8, y: 3.98, w: 5.1, h: 1.05, valign: "middle" });
    s.addImage({ path: AURA_LOGO, x: 0.5, y: 5.15, w: 1.2, h: 0.44 });
    s.addImage({ path: INOVA_LOGO, x: 8.5, y: 5.12, w: 1.1, h: 0.44 });
    s.addShape("rect", { x: 0, y: 5.35, w: 10, h: 0.03, fill: { color: C.coral } });
  }

  const out = path.join(__dirname, "Storytech_Extrator_Apoena.pptx");
  await pres.writeFile({ fileName: out });
  console.log("Apresentação gerada com sucesso:", out);
}

// =====================================================================
//  APRESENTAÇÃO: Fluxo do Processo de Recrutamento e Seleção
//  (design corporativo azul / amarelo / cinza — 16:9 widescreen)
// =====================================================================
async function buildRecrutamento() {
  const FONT = "Calibri";
  // Paleta R&S (azul / amarelo / cinza)
  const R = {
    navy: "16305C", blue: "1F4E9B", blueMed: "2E6CC4", blueSoft: "DCE7F7",
    bluePale: "EEF3FB", yellow: "F5B301", yellowDk: "D99400", yellowSf: "FCE9BE",
    grayDk: "3A414C", gray: "6B7480", grayLt: "D7DCE4", grayBg: "F4F6F9",
    white: "FFFFFF", stop: "C0504D", stopSf: "F4DAD8", green: "2E7D32",
    greenBg: "EAF5EC",
  };

  const pres = new pptxgen();
  // Layout 16:9 widescreen padrão do PowerPoint (13,333" x 7,5")
  pres.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
  pres.layout = "WIDE";
  pres.title = "Fluxo do Processo de Recrutamento e Seleção";
  pres.author = "Déborah Giovana de Andrade Gonçalves";

  const sh = () => ({ type: "outer", color: "1B2A44", blur: 5, offset: 2.5, angle: 90, opacity: 0.30 });
  const noLine = { type: "none" };

  function arrow(s, x1, y1, x2, y2, opt = {}) {
    const { color = R.blueMed, width = 1.6, head = true } = opt;
    const x = Math.min(x1, x2), y = Math.min(y1, y2);
    const w = Math.abs(x2 - x1), h = Math.abs(y2 - y1);
    s.addShape("line", {
      x, y, w, h, flipH: x2 < x1, flipV: y2 < y1,
      line: { color, width, beginArrowType: "none", endArrowType: head ? "triangle" : "none" },
    });
  }

  function chip(s, cx, cy, d, glyph, opt = {}) {
    const { fill = R.yellow, gcolor = R.navy, gsize = 13 } = opt;
    s.addShape("ellipse", { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: fill }, line: noLine, shadow: sh() });
    s.addText(String(glyph), { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fontSize: gsize, bold: true, color: gcolor, align: "center", valign: "middle", fontFace: FONT, margin: 0 });
  }

  function diamond(s, cx, cy, w, h, text, opt = {}) {
    const { fill = R.yellow, line = R.yellowDk, tcolor = R.navy } = opt;
    s.addShape("diamond", { x: cx - w / 2, y: cy - h / 2, w, h, fill: { color: fill }, line: line ? { color: line, width: 1 } : noLine, shadow: sh() });
    if (text) s.addText(text, { x: cx - w / 2, y: cy - h / 2, w, h, fontSize: 9, bold: true, color: tcolor, align: "center", valign: "middle", fontFace: FONT, margin: 0 });
  }

  function stageBox(s, x, y, w, h, num, title, subs, opt = {}) {
    const { bar = R.blue, chipColor = R.blue, titleColor = R.navy, fill = R.white } = opt;
    s.addShape("roundRect", { x, y, w, h, fill: { color: fill }, line: noLine, rectRadius: 0.09, shadow: sh() });
    s.addShape("rect", { x, y, w: 0.1, h, fill: { color: bar } });
    chip(s, x + 0.42, y + 0.35, 0.44, num, { fill: chipColor, gcolor: R.white, gsize: 13 });
    const arr = [{ text: title, options: { bold: true, fontSize: 11.5, color: titleColor, breakLine: true } }];
    subs.forEach(sl => arr.push({ text: sl, options: { fontSize: 8, color: R.gray, breakLine: true } }));
    s.addText(arr, { x: x + 0.72, y: y + 0.03, w: w - 0.82, h: h - 0.06, align: "left", valign: "middle", fontFace: FONT, margin: 2, lineSpacingMultiple: 1.0 });
  }

  function stopNode(s, x, y, w, text) {
    s.addShape("roundRect", { x, y, w, h: 0.42, fill: { color: R.stopSf }, line: { color: R.stop, width: 1 }, rectRadius: 0.12, shadow: sh() });
    s.addText(text, { x, y, w, h: 0.42, fontSize: 8, bold: true, color: R.stop, align: "center", valign: "middle", fontFace: FONT, margin: 1 });
  }

  function ynLabel(s, x, y, text, color) {
    s.addText(text, { x, y, w: 0.55, h: 0.24, fontSize: 8.5, bold: true, color, align: "center", valign: "middle", fontFace: FONT, margin: 0 });
  }

  // ------------------------- SLIDE 1 — CAPA ------------------------- //
  {
    const s = pres.addSlide();
    s.background = { color: R.white };
    s.addShape("rect", { x: 0, y: 0, w: 4.9, h: 7.5, fill: { color: R.navy } });
    s.addShape("rect", { x: 4.9, y: 0, w: 0.12, h: 7.5, fill: { color: R.yellow } });
    s.addShape("ellipse", { x: 3.4, y: -1.5, w: 3.0, h: 3.0, fill: { color: R.blue }, line: noLine });
    s.addShape("ellipse", { x: -1.0, y: 5.7, w: 2.4, h: 2.4, fill: { color: R.blueMed }, line: noLine });

    // motivo gráfico de RH — rede de pessoas
    const person = (cx, cy, scale, color) => {
      const hd = 0.30 * scale, bw = 0.54 * scale, bh = 0.34 * scale;
      s.addShape("ellipse", { x: cx - hd / 2, y: cy - hd / 2, w: hd, h: hd, fill: { color }, line: noLine });
      s.addShape("roundRect", { x: cx - bw / 2, y: cy + hd / 2 - 0.02 * scale, w: bw, h: bh, fill: { color }, line: noLine, rectRadius: 0.14 * scale });
    };
    const nodes = [[2.45, 2.55, R.yellow, 1.25], [1.35, 3.65, R.white, 1.0], [3.55, 3.65, R.white, 1.0], [2.45, 4.75, R.blueMed, 1.05]];
    for (let i = 0; i < nodes.length; i++)
      for (let j = i + 1; j < nodes.length; j++)
        arrow(s, nodes[i][0], nodes[i][1] + 0.15, nodes[j][0], nodes[j][1] + 0.15, { color: "3E5C8A", width: 1.2, head: false });
    nodes.forEach(([cx, cy, col, sc]) => person(cx, cy, sc, col));

    s.addText("RECURSOS HUMANOS", { x: 0.55, y: 0.55, w: 4.0, h: 0.4, fontSize: 12, bold: true, color: R.yellow, align: "left", valign: "top", fontFace: FONT, charSpacing: 1 });
    s.addText([
      { text: "Fluxo do Processo de", options: { fontSize: 34, bold: true, color: R.navy, breakLine: true } },
      { text: "Recrutamento e Seleção", options: { fontSize: 34, bold: true, color: R.blue } },
    ], { x: 5.45, y: 2.25, w: 7.3, h: 1.9, align: "left", valign: "bottom", fontFace: FONT, lineSpacingMultiple: 1.02 });
    s.addShape("rect", { x: 5.5, y: 4.25, w: 1.7, h: 0.06, fill: { color: R.yellow } });
    s.addText("Etapas, aprovações, sistemas utilizados e pontos de decisão", { x: 5.45, y: 4.45, w: 7.2, h: 0.9, fontSize: 15, color: R.gray, align: "left", valign: "top", fontFace: FONT });
    s.addText("Apresentação executiva  ·  Processo de R&S", { x: 5.45, y: 6.75, w: 7.2, h: 0.4, fontSize: 10, bold: true, color: R.gray, align: "left", valign: "top", fontFace: FONT });
  }

  // ---------------- SLIDE 2 — FLUXO COMPLETO ---------------- //
  {
    const s = pres.addSlide();
    s.background = { color: R.grayBg };
    s.addShape("rect", { x: 0, y: 0, w: 13.333, h: 0.86, fill: { color: R.navy } });
    s.addShape("rect", { x: 0, y: 0.86, w: 13.333, h: 0.06, fill: { color: R.yellow } });
    s.addText("Fluxo Completo do Processo de Recrutamento e Seleção", { x: 0.4, y: 0, w: 9.0, h: 0.86, fontSize: 20, bold: true, color: R.white, align: "left", valign: "middle", fontFace: FONT });
    s.addText("Fluxograma executivo", { x: 9.4, y: 0, w: 3.6, h: 0.86, fontSize: 11, bold: true, color: R.yellow, align: "right", valign: "middle", fontFace: FONT });

    // LANE 1 (fluxo principal)
    const LX = 0.45, LW = 3.05, cx1 = LX + LW / 2;
    stageBox(s, LX, 1.10, LW, 0.90, 1, "Requisição de Vaga", ["Origem: Portal de Serviços (HUB RH)", "Gestor solicita abertura da vaga", "Cargo · Centro de custo · Escopo · Contratação"]);
    stageBox(s, LX, 2.18, LW, 0.86, 2, "Publicação da Vaga", ["Publicação na Pandapé", "Quintas-feiras · permanência de 7 dias", "Portal de Vagas · Aura Comunica (sexta)"]);
    stageBox(s, LX, 3.22, LW, 0.80, 3, "Triagem de Currículos", ["Avaliação técnica dos currículos", "Compatibilidade com os requisitos", "Seleção de candidatos aderentes"]);
    stageBox(s, LX, 4.20, LW, 0.86, 4, "Entrevista RH", ["Avaliação comportamental · histórico", "Motivação · alinhamento cultural", "Pretensão salarial"]);
    [[2.00, 2.18], [3.04, 3.22], [4.02, 4.20]].forEach(([ya, yb]) => arrow(s, cx1, ya, cx1, yb, { width: 1.8 }));

    const d1cy = 5.55;
    arrow(s, cx1, 5.06, cx1, d1cy - 0.42, { width: 1.8 });
    diamond(s, cx1, d1cy, 1.55, 0.85, "Aprovado?");
    ynLabel(s, LX - 0.02, d1cy - 0.32, "Não", R.stop);
    arrow(s, cx1 - 0.78, d1cy, LX + 0.25, d1cy, { color: R.stop, width: 1.5 });
    stopNode(s, LX - 0.35, d1cy + 0.30, 1.55, "Reprovado · processo encerrado");
    arrow(s, LX + 0.25, d1cy + 0.02, LX + 0.42, d1cy + 0.30, { color: R.stop, width: 1.4, head: false });
    ynLabel(s, cx1 + 0.42, d1cy - 0.02, "Sim", R.green);

    // BRANCH — Triagem Patrimonial
    const PX = 3.95, PW = 2.05, pcx = PX + PW / 2;
    s.addShape("roundRect", { x: PX, y: 1.10, w: PW, h: 0.44, fill: { color: R.yellow }, line: noLine, rectRadius: 0.10, shadow: sh() });
    s.addText("Triagem Patrimonial", { x: PX, y: 1.10, w: PW, h: 0.44, fontSize: 10.5, bold: true, color: R.navy, align: "center", valign: "middle", fontFace: FONT });
    s.addText("fluxo paralelo", { x: PX, y: 1.56, w: PW, h: 0.24, fontSize: 8, italic: true, color: R.gray, align: "center", valign: "top", fontFace: FONT });
    const mini = (y, text) => {
      s.addShape("roundRect", { x: PX, y, w: PW, h: 0.46, fill: { color: R.white }, line: { color: R.grayLt, width: 0.75 }, rectRadius: 0.08, shadow: sh() });
      s.addText(text, { x: PX, y, w: PW, h: 0.46, fontSize: 8.7, bold: true, color: R.grayDk, align: "center", valign: "middle", fontFace: FONT, margin: 1 });
    };
    mini(1.92, "Envio de Forms");
    mini(2.56, "Preenchimento pelo candidato");
    mini(3.20, "Pesquisa patrimonial");
    arrow(s, pcx, 1.80, pcx, 1.92, { color: R.yellowDk, width: 1.6 });
    [[2.38, 2.56], [3.02, 3.20]].forEach(([ya, yb]) => arrow(s, pcx, ya, pcx, yb, { color: R.yellowDk, width: 1.6 }));

    const pdcy = 4.25;
    arrow(s, pcx, 3.66, pcx, pdcy - 0.42, { color: R.yellowDk, width: 1.6 });
    diamond(s, pcx, pdcy, 1.55, 0.82, "Aprovado?", { fill: R.blueSoft, line: R.blueMed });
    ynLabel(s, pcx - 0.5, pdcy + 0.42, "Não", R.stop);
    arrow(s, pcx, pdcy + 0.41, pcx, pdcy + 0.62, { color: R.stop, width: 1.4 });
    stopNode(s, PX, pdcy + 0.62, PW, "Restrição · processo encerrado");

    // D1 "Sim" -> sobe pelo canal e entra no topo do branch
    const CH1 = 3.72;
    arrow(s, cx1 + 0.78, d1cy, CH1, d1cy, { color: R.green, width: 1.6, head: false });
    arrow(s, CH1, d1cy, CH1, 1.32, { color: R.green, width: 1.6, head: false });
    arrow(s, CH1, 1.32, PX, 1.32, { color: R.green, width: 1.6, head: true });
    ynLabel(s, pcx + 0.60, pdcy - 0.30, "Sim", R.green);

    // LANE 2 (continuação)
    const QX = 6.35, QW = 3.05, qcx = QX + QW / 2;
    stageBox(s, QX, 1.10, QW, 0.86, 5, "Entrevista com Gestor", ["Conhecimento técnico · experiência", "Competências · fit com a equipe"]);
    const CH2 = 6.12;
    arrow(s, pcx + 0.78, pdcy, CH2, pdcy, { color: R.green, width: 1.6, head: false });
    arrow(s, CH2, pdcy, CH2, 1.53, { color: R.green, width: 1.6, head: false });
    arrow(s, CH2, 1.53, QX, 1.53, { color: R.green, width: 1.6 });

    const d2cy = 2.55;
    arrow(s, qcx, 1.96, qcx, d2cy - 0.41, { width: 1.8 });
    diamond(s, qcx, d2cy, 1.7, 0.82, "Candidato\nFinalista?");
    const CH3 = 6.28;
    ynLabel(s, qcx - 1.05, d2cy - 0.02, "Não", R.stop);
    arrow(s, qcx - 0.85, d2cy, CH3, d2cy, { color: R.stop, width: 1.4, head: false });
    arrow(s, CH3, d2cy, CH3, 6.58, { color: R.stop, width: 1.4, head: false });
    arrow(s, CH3, 6.58, 6.62, 6.58, { color: R.stop, width: 1.4, head: true });
    s.addShape("roundRect", { x: 6.62, y: 6.30, w: 2.78, h: 0.56, fill: { color: R.stopSf }, line: { color: R.stop, width: 1 }, rectRadius: 0.10, shadow: sh() });
    s.addText("Não finalista  →  retorna à nova triagem ou nova publicação da vaga", { x: 6.62, y: 6.30, w: 2.78, h: 0.56, fontSize: 8, bold: true, color: R.stop, align: "center", valign: "middle", fontFace: FONT, margin: 1 });
    ynLabel(s, qcx + 0.10, d2cy + 0.42, "Sim", R.green);

    stageBox(s, QX, 3.22, QW, 0.80, 6, "Carta Proposta", ["Aprovação RH Custos  ·  Aprovação Gestor", "Formalização da oferta ao candidato"], { bar: R.yellowDk, chipColor: R.yellowDk });
    arrow(s, qcx, d2cy + 0.41, qcx, 3.22, { width: 1.8 });
    stageBox(s, QX, 4.18, QW, 0.86, 7, "Formalização & Assinatura", ["Envio por e-mail da proposta", "Assinatura eletrônica via GOV.BR", "Aceite do candidato"]);
    arrow(s, qcx, 4.02, qcx, 4.18, { width: 1.8 });
    stageBox(s, QX, 5.22, QW, 0.90, 8, "Seleção Concluída", ["Onboarding do novo colaborador", "Integração à equipe e à empresa"], { bar: R.green, chipColor: R.green, titleColor: "1B5E20", fill: R.greenBg });
    arrow(s, qcx, 5.04, qcx, 5.22, { color: R.green, width: 1.8 });

    // PAINÉIS INFORMATIVOS
    const panel = (x, y, w, h, title, items, titleBg) => {
      s.addShape("roundRect", { x, y, w, h, fill: { color: R.white }, line: noLine, rectRadius: 0.06, shadow: sh() });
      s.addShape("roundRect", { x, y, w, h: 0.40, fill: { color: titleBg }, line: noLine, rectRadius: 0.08 });
      s.addShape("rect", { x, y: y + 0.22, w, h: 0.20, fill: { color: titleBg } });
      s.addShape("rect", { x, y, w: 0.08, h: 0.40, fill: { color: R.yellow } });
      s.addText(title, { x: x + 0.05, y, w: w - 0.1, h: 0.40, fontSize: 10.5, bold: true, color: R.white, align: "center", valign: "middle", fontFace: FONT });
      const arr = items.map(it => ({ text: "•  " + it, options: { fontSize: 8.4, color: R.grayDk, breakLine: true, paraSpaceAfter: 3 } }));
      s.addText(arr, { x: x + 0.14, y: y + 0.48, w: w - 0.24, h: h - 0.54, align: "left", valign: "top", fontFace: FONT, lineSpacingMultiple: 1.0 });
    };
    const PNX = 9.72, PNW = 3.32;
    panel(PNX, 1.10, PNW, 2.00, "SISTEMAS UTILIZADOS", ["Portal de Serviços", "Teams", "Outlook", "Pandapé", "Microsoft Forms", "GOV.BR", "Aura Comunica"], R.blue);
    panel(PNX, 3.24, PNW, 1.86, "PONTOS DE DECISÃO", ["Aprovação da Requisição", "Aprovação RH Custos", "Aprovação Entrevista RH", "Aprovação Triagem Patrimonial", "Aprovação Gestor", "Aceite da Proposta"], R.blue);
    panel(PNX, 5.24, PNW, 1.72, "POSSÍVEIS ENCERRAMENTOS", ["Reprovação na Entrevista RH", "Restrição na Triagem Patrimonial", "Reprovação pelo Gestor", "Recusa da Proposta", "Reabertura da vaga"], R.stop);

    // Legenda
    const legY = 7.06;
    s.addText("Legenda:", { x: 0.45, y: legY, w: 1.0, h: 0.34, fontSize: 8.5, bold: true, color: R.grayDk, align: "left", valign: "middle", fontFace: FONT });
    const legItem = (x, color, label, dmd = false) => {
      if (dmd) s.addShape("diamond", { x: x - 0.13, y: legY + 0.06, w: 0.26, h: 0.22, fill: { color }, line: noLine });
      else s.addShape("roundRect", { x: x - 0.13, y: legY + 0.06, w: 0.26, h: 0.22, fill: { color }, line: noLine, rectRadius: 0.05 });
      s.addText(label, { x: x + 0.18, y: legY, w: 1.4, h: 0.34, fontSize: 8, color: R.gray, align: "left", valign: "middle", fontFace: FONT });
    };
    legItem(1.60, R.blue, "Atividade");
    legItem(3.55, R.yellow, "Decisão", true);
    legItem(5.50, R.stopSf, "Encerramento");
    legItem(7.55, R.greenBg, "Conclusão");
  }

  // ---------------- SLIDE 3 — VISÃO EXECUTIVA ---------------- //
  {
    const s = pres.addSlide();
    s.background = { color: R.white };
    s.addShape("rect", { x: 0, y: 0, w: 13.333, h: 0.86, fill: { color: R.navy } });
    s.addShape("rect", { x: 0, y: 0.86, w: 13.333, h: 0.06, fill: { color: R.yellow } });
    s.addText("Visão Executiva do Processo", { x: 0.4, y: 0, w: 9.5, h: 0.86, fontSize: 20, bold: true, color: R.white, align: "left", valign: "middle", fontFace: FONT });
    s.addText("Linha do tempo — 8 etapas", { x: 9.0, y: 0, w: 4.0, h: 0.86, fontSize: 11, bold: true, color: R.yellow, align: "right", valign: "middle", fontFace: FONT });

    const stages = [
      ["1", "Requisição", "Abertura da vaga pelo gestor"],
      ["2", "Divulgação", "Publicação na Pandapé"],
      ["3", "Triagem", "Análise de currículos"],
      ["4", "Entrevistas", "RH e gestor"],
      ["5", "Aprovação", "Decisões e validações"],
      ["6", "Proposta", "Carta e negociação"],
      ["7", "Admissão", "Assinatura via GOV.BR"],
      ["8", "Onboarding", "Integração do colaborador"],
    ];
    const lineY = 3.85, xStart = 0.95, xEnd = 12.4;
    const seg = (xEnd - xStart) / (stages.length - 1);
    s.addShape("rect", { x: xStart, y: lineY - 0.03, w: xEnd - xStart, h: 0.06, fill: { color: R.grayLt } });
    stages.forEach(([num, title, desc], i) => {
      const cx = xStart + seg * i, up = i % 2 === 0;
      const col = up ? R.blue : R.yellowDk, cardH = 1.32, cw = 1.42;
      const cardY = up ? lineY - 0.55 - cardH : lineY + 0.55;
      if (up) arrow(s, cx, lineY - 0.35, cx, cardY + cardH, { color: col, width: 1.4, head: false });
      else arrow(s, cx, lineY + 0.35, cx, cardY, { color: col, width: 1.4, head: false });
      s.addShape("roundRect", { x: cx - cw / 2, y: cardY, w: cw, h: cardH, fill: { color: R.white }, line: { color: R.grayLt, width: 0.75 }, rectRadius: 0.08, shadow: sh() });
      s.addShape("rect", { x: cx - cw / 2, y: cardY, w: cw, h: 0.09, fill: { color: col } });
      s.addText([
        { text: title, options: { bold: true, fontSize: 12, color: R.navy, breakLine: true, paraSpaceAfter: 4 } },
        { text: desc, options: { fontSize: 8.3, color: R.gray } },
      ], { x: cx - cw / 2 + 0.06, y: cardY + 0.12, w: cw - 0.12, h: cardH - 0.16, align: "center", valign: "top", fontFace: FONT, lineSpacingMultiple: 1.0 });
      chip(s, cx, lineY, 0.56, num, { fill: col, gcolor: R.white, gsize: 15 });
    });

    s.addText("Do pedido à integração — um fluxo estruturado, auditável e orientado a decisões.", { x: 0.4, y: 5.85, w: 12.5, h: 0.4, fontSize: 12, italic: true, color: R.gray, align: "center", valign: "middle", fontFace: FONT });

    const indY = 6.35;
    const indicators = [["8", "etapas principais", R.blue], ["6", "pontos de decisão", R.yellowDk], ["7", "sistemas integrados", R.blue], ["5", "possíveis encerramentos", R.stop]];
    const iw = 2.85, gap = (13.333 - iw * indicators.length) / (indicators.length + 1);
    indicators.forEach(([val, lbl, col], i) => {
      const x = gap + i * (iw + gap);
      s.addShape("roundRect", { x, y: indY, w: iw, h: 0.82, fill: { color: R.grayBg }, line: noLine, rectRadius: 0.10, shadow: sh() });
      s.addShape("rect", { x, y: indY, w: 0.10, h: 0.82, fill: { color: col } });
      s.addText(val, { x: x + 0.18, y: indY, w: 0.85, h: 0.82, fontSize: 26, bold: true, color: col, align: "center", valign: "middle", fontFace: FONT });
      s.addText(lbl, { x: x + 1.02, y: indY, w: iw - 1.1, h: 0.82, fontSize: 11, bold: true, color: R.grayDk, align: "left", valign: "middle", fontFace: FONT });
    });
  }

  const out = path.join(__dirname, "Fluxo_Recrutamento_Selecao.pptx");
  await pres.writeFile({ fileName: out });
  console.log("Apresentação gerada com sucesso:", out);
}

async function main() {
  await build();
  await buildRecrutamento();
}

main().catch(e => { console.error(e); process.exit(1); });