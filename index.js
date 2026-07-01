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
const AURA_LOGO   = path.join(__dirname, "assets", "aura_logo.png");
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

build().catch(e => { console.error(e); process.exit(1); });