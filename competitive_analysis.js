const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageBreak, Header, Footer, VerticalAlign
} = require('docx');
const fs = require('fs');

// Brand colors
const NAVY     = "1B3A6B";
const BLUE     = "1B4F8A";
const ACCENT   = "2E75B6";
const GREEN    = "1A7A4A";
const AMBER    = "B8670A";
const RED_DARK = "8B1A1A";
const LGRAY    = "F4F6FA";
const MGRAY    = "D9DDE6";
const DGRAY    = "444444";
const WHITE    = "FFFFFF";
const TEXT     = "1E1E1E";

const nb = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: nb, bottom: nb, left: nb, right: nb };
const thinBorder = (c = MGRAY) => ({ style: BorderStyle.SINGLE, size: 4, color: c });
const thinBorders = (c = MGRAY) => ({ top: thinBorder(c), bottom: thinBorder(c), left: thinBorder(c), right: thinBorder(c) });

function sp(before = 0, after = 0) { return { before, after }; }
function run(text, opts = {}) { return new TextRun({ text, font: "Calibri", size: 22, color: TEXT, ...opts }); }
function boldRun(text, opts = {}) { return run(text, { bold: true, ...opts }); }

function para(children, opts = {}) {
  return new Paragraph({ spacing: sp(60, 60), children: Array.isArray(children) ? children : [children], ...opts });
}

function h1(text) {
  return new Paragraph({
    spacing: sp(400, 120),
    children: [new TextRun({ text, font: "Calibri", size: 40, bold: true, color: BLUE })]
  });
}
function h2(text, color = ACCENT) {
  return new Paragraph({
    spacing: sp(280, 80),
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 2 } },
    children: [new TextRun({ text, font: "Calibri", size: 30, bold: true, color })]
  });
}
function h3(text, color = NAVY) {
  return new Paragraph({
    spacing: sp(200, 60),
    children: [new TextRun({ text, font: "Calibri", size: 24, bold: true, color })]
  });
}
function bodyText(text) {
  return para([run(text)]);
}
function spacer(n = 120) {
  return new Paragraph({ spacing: sp(n, 0), children: [] });
}
function pageBreak() {
  return new Paragraph({ children: [new TextRun({ break: 1 })] });
}

// Colored badge cell
function badgeCell(text, fill, textColor = WHITE, width = 1600) {
  return new TableCell({
    borders: noBorders,
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    width: { size: width, type: WidthType.DXA },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: sp(0,0),
      children: [new TextRun({ text, font: "Calibri", size: 18, bold: true, color: textColor })]
    })]
  });
}

// Regular data cell
function dataCell(children, fill = WHITE, width = null, align = AlignmentType.LEFT) {
  const cell = {
    borders: thinBorders(),
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: Array.isArray(children) ? children : [para([run(typeof children === 'string' ? children : '')], { spacing: sp(0,0), alignment: align })]
  };
  if (width) cell.width = { size: width, type: WidthType.DXA };
  return new TableCell(cell);
}

function headerCell(text, fill = BLUE, textColor = WHITE, width = null) {
  const cell = {
    borders: thinBorders(BLUE),
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: [new Paragraph({
      spacing: sp(0,0),
      children: [new TextRun({ text, font: "Calibri", size: 20, bold: true, color: textColor })]
    })]
  };
  if (width) cell.width = { size: width, type: WidthType.DXA };
  return new TableCell(cell);
}

// SWOT box
function swotBox(title, fill, textColor, items) {
  const rows = [
    new TableRow({ children: [new TableCell({
      borders: noBorders,
      shading: { fill, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 60, left: 140, right: 140 },
      children: [new Paragraph({ spacing: sp(0,0), children: [new TextRun({ text: title, font: "Calibri", size: 24, bold: true, color: textColor })] })]
    })]})
  ];
  items.forEach(item => {
    rows.push(new TableRow({ children: [new TableCell({
      borders: noBorders,
      shading: { fill, type: ShadingType.CLEAR },
      margins: { top: 30, bottom: 30, left: 140, right: 140 },
      children: [new Paragraph({
        spacing: sp(0,0),
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: item, font: "Calibri", size: 19, color: textColor })]
      })]
    })]});
  });
  return rows;
}

function swotTable(strengths, weaknesses, opportunities, threats) {
  const swSize = 4600;
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [swSize, swSize],
    rows: [
      new TableRow({ children: [
        new TableCell({
          borders: thinBorders(WHITE),
          shading: { fill: "1A4A2A", type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 120, left: 140, right: 140 },
          children: [
            new Paragraph({ spacing: sp(0,40), children: [new TextRun({ text: "STRENGTHS", font: "Calibri", size: 24, bold: true, color: WHITE })] }),
            ...strengths.map(s => new Paragraph({ spacing: sp(20,20), numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: s, font: "Calibri", size: 19, color: WHITE })] }))
          ]
        }),
        new TableCell({
          borders: thinBorders(WHITE),
          shading: { fill: "7A1A1A", type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 120, left: 140, right: 140 },
          children: [
            new Paragraph({ spacing: sp(0,40), children: [new TextRun({ text: "WEAKNESSES", font: "Calibri", size: 24, bold: true, color: WHITE })] }),
            ...weaknesses.map(s => new Paragraph({ spacing: sp(20,20), numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: s, font: "Calibri", size: 19, color: WHITE })] }))
          ]
        })
      ]}),
      new TableRow({ children: [
        new TableCell({
          borders: thinBorders(WHITE),
          shading: { fill: "1A3A6A", type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 120, left: 140, right: 140 },
          children: [
            new Paragraph({ spacing: sp(0,40), children: [new TextRun({ text: "OPPORTUNITIES", font: "Calibri", size: 24, bold: true, color: WHITE })] }),
            ...opportunities.map(s => new Paragraph({ spacing: sp(20,20), numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: s, font: "Calibri", size: 19, color: WHITE })] }))
          ]
        }),
        new TableCell({
          borders: thinBorders(WHITE),
          shading: { fill: "5A4A00", type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 120, left: 140, right: 140 },
          children: [
            new Paragraph({ spacing: sp(0,40), children: [new TextRun({ text: "THREATS", font: "Calibri", size: 24, bold: true, color: WHITE })] }),
            ...threats.map(s => new Paragraph({ spacing: sp(20,20), numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: s, font: "Calibri", size: 19, color: WHITE })] }))
          ]
        })
      ]})
    ]
  });
}

// Feature matrix check/dash
const CHECK = "✓";
const PART  = "◑";
const DASH  = "—";
const HOT   = "★";

function matrixCell(value, fill = WHITE) {
  const isCheck = value === CHECK || value === HOT;
  const color = value === CHECK ? GREEN : value === HOT ? AMBER : value === PART ? ACCENT : DGRAY;
  return new TableCell({
    borders: thinBorders(),
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    width: { size: 900, type: WidthType.DXA },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: sp(0,0),
      children: [new TextRun({ text: value, font: "Calibri", size: 20, bold: isCheck, color })]
    })]
  });
}

function featureRow(feature, values, rowFill = WHITE) {
  const cells = [dataCell(feature, rowFill, 2360)];
  values.forEach(v => cells.push(matrixCell(v, rowFill)));
  return new TableRow({ children: cells });
}

// ─── DOCUMENT ─────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 400, hanging: 200 } }, run: { color: ACCENT } } }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Calibri", size: 22, color: TEXT } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 40, bold: true, font: "Calibri", color: BLUE }, paragraph: { spacing: sp(400,120) } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: "Calibri", color: ACCENT }, paragraph: { spacing: sp(280,80) } }
    ]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1200, bottom: 1080, left: 1200 } } },
    headers: { default: new Header({ children: [
      new Table({ width: { size: 9840, type: WidthType.DXA }, columnWidths: [5800, 4040],
        rows: [new TableRow({ children: [
          new TableCell({ borders: noBorders, width: { size: 5800, type: WidthType.DXA }, margins: { top:40, bottom:40, left:0, right:0 },
            children: [new Paragraph({ children: [new TextRun({ text: "EPIC Addiction Recovery", font: "Calibri", size: 20, bold: true, color: BLUE })] })] }),
          new TableCell({ borders: noBorders, width: { size: 4040, type: WidthType.DXA }, margins: { top:40, bottom:40, left:0, right:0 },
            children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Sober Companion — Competitive Analysis | June 2026", font: "Calibri", size: 18, italics: true, color: "888888" })] })] })
        ]})]
      }),
      new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 } }, spacing: sp(0,0), children: [] })
    ]})},
    footers: { default: new Footer({ children: [
      new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 4, color: MGRAY, space: 1 } }, spacing: sp(60,0), alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Confidential — For internal planning purposes only  |  epicrecovery.ca", font: "Calibri", size: 18, color: "888888" })] })
    ]})},

    children: [

      // ═══════════════════════════════════════════════
      // COVER
      // ═══════════════════════════════════════════════
      spacer(180),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: sp(0,60),
        children: [new TextRun({ text: "COMPETITIVE LANDSCAPE ANALYSIS", font: "Calibri", size: 52, bold: true, color: BLUE })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: sp(0,60),
        children: [new TextRun({ text: "Sober Companion & Recovery Coaching Market", font: "Calibri", size: 32, color: ACCENT })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: sp(0,0),
        children: [new TextRun({ text: "London, Ontario  •  Province-Wide  •  National Canada  •  International", font: "Calibri", size: 24, italics: true, color: DGRAY })] }),
      spacer(60),
      new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: BLUE } }, spacing: sp(0,0), children: [] }),
      spacer(60),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: sp(0,0),
        children: [new TextRun({ text: "Prepared for EPIC Addiction Recovery  |  Pilot Program Planning  |  June 2026", font: "Calibri", size: 20, color: DGRAY })] }),
      spacer(240),

      // Executive Summary box
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
        rows: [new TableRow({ children: [new TableCell({
          borders: thinBorders(ACCENT),
          shading: { fill: LGRAY, type: ShadingType.CLEAR },
          margins: { top: 160, bottom: 160, left: 200, right: 200 },
          children: [
            new Paragraph({ spacing: sp(0,80), children: [new TextRun({ text: "Executive Summary", font: "Calibri", size: 26, bold: true, color: BLUE })] }),
            new Paragraph({ spacing: sp(0,60), children: [new TextRun({ text: "The sober companion and recovery coaching market is active at the national and international levels, but virtually absent in London, Ontario and across most of Southwestern Ontario. No dedicated provider with in-person, clinically credentialed sober companion services currently operates in the London market. This represents a significant first-mover opportunity for EPIC Addiction Recovery to establish a premium, locally grounded service with minimal direct local competition, strong regional referral demand, and differentiation that national virtual-only providers cannot replicate.", font: "Calibri", size: 22, color: TEXT })] }),
            new Paragraph({ spacing: sp(0,0), children: [new TextRun({ text: "This document profiles competitors across four geographic tiers, presents feature matrices and SWOT analyses for each tier, and closes with a hypothetical competitive position for EPIC at London-market scale.", font: "Calibri", size: 22, italics: true, color: DGRAY })] })
          ]
        })]})]}),

      spacer(120),

      // ═══════════════════════════════════════════════
      // SECTION 1: LONDON ONTARIO
      // ═══════════════════════════════════════════════
      h1("TIER 1 — London, Ontario Market"),
      bodyText("A thorough search for dedicated sober companion or recovery coaching services physically based in London, Ontario returns a stark result: there are none. The London addiction recovery landscape is dominated entirely by publicly funded clinical services and nonprofit organizations — none of which offer private-pay, in-person sober companion or intensive recovery coaching programs."),
      spacer(80),

      h2("Active Providers — London, Ontario"),
      spacer(40),

      // London providers table
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2800, 2200, 2200, 2160],
        rows: [
          new TableRow({ children: [
            headerCell("Organization", NAVY, WHITE, 2800),
            headerCell("Service Type", NAVY, WHITE, 2200),
            headerCell("Price Point", NAVY, WHITE, 2200),
            headerCell("Companion / Coach?", NAVY, WHITE, 2160)
          ]}),
          ...[
            ["CMHA Thames Valley", "Public mental health & addiction counselling", "Free (OHIP/provincial)", "No — no sober companion"],
            ["Turning Point Inc.", "Long-term recovery home — group programs & life skills", "Subsidized / low cost", "No — residential only"],
            ["London Clinic OAT Centre", "Opioid agonist therapy (MAT)", "OHIP-covered", "No — clinical/medical"],
            ["Carepoint Consumption & Treatment Services", "Supervised consumption, harm reduction", "Free / government-funded", "No"],
            ["Horizons Opioid Treatment Clinic", "Opiate dependence treatment", "OHIP/subsidized", "No"],
            ["AA / NA / Other 12-Step", "Peer recovery support groups", "Free", "Sponsor only — no professional"],
            ["EPIC Addiction Recovery", "Private-pay outpatient counselling & recovery maintenance", "Private-pay", "NOT YET — pilot opportunity"],
          ].map((r, i) => new TableRow({ children: [
            dataCell(r[0], i % 2 === 0 ? LGRAY : WHITE, 2800),
            dataCell(r[1], i % 2 === 0 ? LGRAY : WHITE, 2200),
            dataCell(r[2], i % 2 === 0 ? LGRAY : WHITE, 2200),
            dataCell(r[3], i % 2 === 0 ? LGRAY : WHITE, 2160),
          ]}))
        ]
      }),
      spacer(80),

      h3("Key Finding: London Is an Empty Field"),
      bodyText("There is no professional, private-pay sober companion service operating in London, Ontario. The entire private-pay recovery support market is unoccupied at the companion level. Individuals and families in London who can afford premium recovery support are currently forced to look to Toronto-based or national virtual providers — or go without. This is the defining competitive reality for EPIC's pilot: it would not be entering a crowded market. It would be creating one."),
      spacer(100),

      h2("London Market — SWOT Analysis"),
      spacer(40),
      swotTable(
        ["Zero direct local competition — genuine first-mover position", "EPIC already has brand presence, referral network, and private-pay clientele in London", "Western University, Fanshawe College, and business community generate professional-class clients", "Lower operating costs than Toronto — better margin on comparable pricing"],
        ["London is a smaller market than Toronto — premium client pool is smaller", "No established consumer awareness of sober companion services locally", "Client acquisition cost may be higher in early months before word-of-mouth builds", "Must build from scratch — no existing competitor models to reference locally"],
        ["Unmet demand from private rehab graduates with no transition support options", "Referral pipeline from EPIC's existing outpatient program is immediate and organic", "Proximity to Toronto allows eventual regional expansion", "Provincial HART hub investment signals growing public awareness of recovery support gaps"],
        ["National virtual providers (Michael Walsh, Evolve Sobriety) can serve London clients remotely", "Public system improvements could reduce urgency for private-pay clients over time", "Lack of regulation means competitors can enter easily once market is established", "Small market may limit ability to scale a large companion roster"]
      ),
      spacer(120),

      pageBreak(),

      // ═══════════════════════════════════════════════
      // SECTION 2: ONTARIO
      // ═══════════════════════════════════════════════
      h1("TIER 2 — Province of Ontario"),
      bodyText("At the Ontario level, a small number of private-pay recovery support providers operate — primarily concentrated in Toronto and the GTA. None have a dedicated in-person presence in Southwestern Ontario. Competition is thin, pricing is often opaque, and most operators are one-person practices rather than organized programs."),
      spacer(80),

      h2("Ontario Competitors — Profile Summary"),
      spacer(40),

      // Toronto Sober Coach
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2200, 7160],
        rows: [
          new TableRow({ children: [headerCell("COMPETITOR", NAVY, WHITE, 2200), headerCell("Toronto Sober Coach  (torontosobercoach.com)", ACCENT, WHITE, 7160)] }),
          ...[
            ["Location", "Toronto, ON — in-person and phone/virtual"],
            ["Services", "Sober coaching for addiction/alcoholism, post-residential support, individual sessions"],
            ["Pricing", "Pricing not publicly listed — consult required"],
            ["Credentials", "Minimal public information on qualifications"],
            ["Marketing", "Basic website, low SEO presence, no active social media found"],
            ["Positioning", "Positioned as private post-residential support for GTA clients — 'insure the investment in recovery'"],
            ["Gap", "No information on in-person availability, credentials, or service scope — low trust-building online presence"],
          ].map((r, i) => new TableRow({ children: [
            dataCell(r[0], i % 2 === 0 ? LGRAY : WHITE, 2200),
            dataCell(r[1], i % 2 === 0 ? LGRAY : WHITE, 7160)
          ]}))
        ]
      }),
      spacer(80),

      // Evolve Sobriety
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2200, 7160],
        rows: [
          new TableRow({ children: [headerCell("COMPETITOR", NAVY, WHITE, 2200), headerCell("Evolve Sobriety  (evolvesobriety.com)", ACCENT, WHITE, 7160)] }),
          ...[
            ["Location", "Ontario-based; operates nationally via online platform"],
            ["Services", "Online counselling & therapy, sober companions & coaching, case management, sober transport, in-home detox, interventions, assessments"],
            ["Pricing", "Not publicly listed — 'services not covered by insurance or provincial health plans'"],
            ["Credentials", "CACCF and CSAM member; clinical team with professional credentials"],
            ["Marketing", "Active Facebook, Twitter, Instagram, LinkedIn; decent SEO; content marketing via blog"],
            ["Positioning", "Concierge recovery care — discreet, confidential, national reach via online delivery"],
            ["Gap", "Primarily online; limited in-person physical presence; does not serve London specifically"],
          ].map((r, i) => new TableRow({ children: [
            dataCell(r[0], i % 2 === 0 ? LGRAY : WHITE, 2200),
            dataCell(r[1], i % 2 === 0 ? LGRAY : WHITE, 7160)
          ]}))
        ]
      }),
      spacer(80),

      // Clean Livin / Addiction Counsellor Edmonton
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2200, 7160],
        rows: [
          new TableRow({ children: [headerCell("COMPETITOR", NAVY, WHITE, 2200), headerCell("Clean Livin Recovery Coaching  (addiction-counsellor.com)", ACCENT, WHITE, 7160)] }),
          ...[
            ["Location", "Edmonton, AB — in-person + online Ontario"],
            ["Services", "Sober coaching / recovery coaching, 1-on-1 sessions, online delivery"],
            ["Pricing", "$175 + GST per 50-minute session (transparent, publicly listed)"],
            ["Credentials", "MA Counselling Psychology (Yorkville 2025), Canadian Certified Recovery Coach (CACCF), ASATT-Candidate"],
            ["Marketing", "Clean professional site; CACCF badge prominently displayed; clear credential-forward positioning"],
            ["Positioning", "Credential-led, judgment-free, moderate/harm-reduction friendly — not abstinence-only"],
            ["Relevance to London", "Online reach means they compete for Ontario clients; rate transparency is a benchmark"],
          ].map((r, i) => new TableRow({ children: [
            dataCell(r[0], i % 2 === 0 ? LGRAY : WHITE, 2200),
            dataCell(r[1], i % 2 === 0 ? LGRAY : WHITE, 7160)
          ]}))
        ]
      }),
      spacer(100),

      h2("Ontario-Level Feature Matrix"),
      spacer(40),
      new Table({ width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2360, 900, 900, 900, 900, 900, 900, 900, 900],
        rows: [
          new TableRow({ children: [
            headerCell("Feature / Service", NAVY, WHITE, 2360),
            headerCell("Toronto\nSober Coach", NAVY, WHITE, 900),
            headerCell("Evolve\nSobriety", NAVY, WHITE, 900),
            headerCell("Clean\nLivin", NAVY, WHITE, 900),
            headerCell("CMHA\nLondon", NAVY, WHITE, 900),
            headerCell("EPIC\n(Current)", NAVY, WHITE, 900),
            headerCell("EPIC\n(Pilot)", NAVY, WHITE, 900),
          ]}),
          ...([
            ["In-person local (London)", [DASH, DASH, DASH, CHECK, CHECK, HOT]],
            ["24/7 live-in companion", [PART, DASH, DASH, DASH, DASH, HOT]],
            ["Post-residential transition support", [CHECK, CHECK, DASH, DASH, DASH, CHECK]],
            ["Event / travel accompaniment", [DASH, PART, DASH, DASH, DASH, HOT]],
            ["Recovery coaching (scheduled)", [CHECK, CHECK, CHECK, DASH, DASH, CHECK]],
            ["Clinically credentialed staff", [DASH, CHECK, CHECK, CHECK, CHECK, CHECK]],
            ["Lived recovery experience", [PART, PART, CHECK, PART, CHECK, CHECK]],
            ["Transparent public pricing", [DASH, DASH, CHECK, DASH, DASH, CHECK]],
            ["Family support services", [DASH, CHECK, DASH, CHECK, PART, PART]],
            ["Case management", [DASH, CHECK, DASH, CHECK, DASH, PART]],
            ["Online / virtual delivery", [PART, CHECK, CHECK, CHECK, PART, PART]],
            ["Interventions", [DASH, CHECK, DASH, DASH, DASH, DASH]],
          ]).map((r, i) => featureRow(r[0], r[1], i % 2 === 0 ? LGRAY : WHITE))
        ]
      }),
      spacer(60),
      para([
        run("Key: "), boldRun(CHECK, { color: GREEN }), run("  = Available    "),
        boldRun(HOT, { color: AMBER }), run("  = Planned for pilot    "),
        run(PART + " = Partial    "),
        run(DASH + " = Not available")
      ]),
      spacer(100),

      h2("Ontario-Level SWOT Analysis"),
      spacer(40),
      swotTable(
        ["Only dedicated private-pay recovery coach with CACCF credentials in Ontario willing to serve London in-person", "EPIC's existing client base provides an immediate warm referral pipeline", "No in-person competitor in Southwestern Ontario"],
        ["No established brand recognition outside EPIC's existing network", "Single companion means capacity constraints limit scale initially", "Ontario operators with broader virtual reach can compete for clients online"],
        ["Significant unmet demand from post-residential clients across Southwest Ontario (Kitchener, Hamilton, Windsor corridor)", "EPIC's credentialing and clinical brand creates trust differential over uncredentialed Toronto operators", "Ability to set premium pricing with no comparable local benchmark"],
        ["Virtual-only Ontario providers (Evolve Sobriety) have lower overhead and can undercut on hourly coaching rates", "Evolve Sobriety's CACCF membership creates credential parity — differentiation must come from in-person depth", "Regulatory changes could create new licensing requirements"]
      ),
      spacer(120),

      pageBreak(),

      // ═══════════════════════════════════════════════
      // SECTION 3: NATIONAL CANADA
      // ═══════════════════════════════════════════════
      h1("TIER 3 — National Canada"),
      bodyText("At the national level, two operators stand out as the most credible and established Canadian sober companion / recovery coaching brands. Both are strong, but both have significant structural limitations that EPIC can exploit."),
      spacer(80),

      h2("National Competitors — Profile Summary"),
      spacer(40),

      // Michael Walsh
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2200, 7160],
        rows: [
          new TableRow({ children: [headerCell("COMPETITOR", NAVY, WHITE, 2200), headerCell("Michael Walsh — Sober Coach Canada  (michaelwalsh.com)", ACCENT, WHITE, 7160)] }),
          ...[
            ["Location", "Victoria, BC — virtual nationwide; in-person in Victoria/Vancouver/Duncan only"],
            ["Services", "Recovery coaching, sober companions, 1-on-1 therapy, interventions, family support (CRAFT model), online addiction treatment"],
            ["Pricing", "Not publicly listed — free consultation to discuss"],
            ["Credentials", "Certified Recovery Coach; team includes coaches, therapists, and sober companions across 8+ Canadian cities"],
            ["Marketing", "Strong SEO; active social media (Facebook, Twitter, Instagram, LinkedIn, WhatsApp); extensive content marketing; 1,000+ city-specific landing pages"],
            ["Positioning", "Canada's leading recovery coach — broad, national, evidence-based, inclusive (harm reduction AND abstinence)"],
            ["Gap vs EPIC", "No in-person presence in Southwestern Ontario; virtual delivery only for Ontario clients; no clinical outpatient anchor organization behind him"],
          ].map((r, i) => new TableRow({ children: [
            dataCell(r[0], i % 2 === 0 ? LGRAY : WHITE, 2200),
            dataCell(r[1], i % 2 === 0 ? LGRAY : WHITE, 7160)
          ]}))
        ]
      }),
      spacer(80),

      // Evolve Sobriety national profile
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2200, 7160],
        rows: [
          new TableRow({ children: [headerCell("COMPETITOR", NAVY, WHITE, 2200), headerCell("Evolve Sobriety — National  (evolvesobriety.com)", ACCENT, WHITE, 7160)] }),
          ...[
            ["Location", "National — operates across all provinces via online delivery; limited in-person"],
            ["Services", "Sober companions & coaching, online counselling, case management, sober transport, in-home detox, interventions"],
            ["Pricing", "Not publicly listed; private-pay only; 'not covered by insurance or provincial health plans'"],
            ["Credentials", "CACCF and CSAM member; clinical team"],
            ["Marketing", "Social media active; website SEO moderate; blog and content strategy in place"],
            ["Positioning", "Nationwide concierge recovery — 'gateway to network of treatment centers across Canada'"],
            ["Gap vs EPIC", "Primarily a virtual referral and matching service — limited physical presence; no local community anchoring"],
          ].map((r, i) => new TableRow({ children: [
            dataCell(r[0], i % 2 === 0 ? LGRAY : WHITE, 2200),
            dataCell(r[1], i % 2 === 0 ? LGRAY : WHITE, 7160)
          ]}))
        ]
      }),
      spacer(100),

      h2("National Canada — Feature Matrix"),
      spacer(40),
      new Table({ width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2360, 900, 900, 900, 900, 900],
        rows: [
          new TableRow({ children: [
            headerCell("Feature / Service", NAVY, WHITE, 2360),
            headerCell("Michael\nWalsh", NAVY, WHITE, 900),
            headerCell("Evolve\nSobriety", NAVY, WHITE, 900),
            headerCell("Sober Coach\nCanada (VIC)", NAVY, WHITE, 900),
            headerCell("EPIC\n(Current)", NAVY, WHITE, 900),
            headerCell("EPIC\n(Pilot)", NAVY, WHITE, 900),
          ]}),
          ...([
            ["In-person companion (SW Ontario)", [DASH, DASH, DASH, CHECK, HOT]],
            ["24/7 live-in companion (national)", [CHECK, CHECK, DASH, DASH, PART]],
            ["Virtual / online coaching", [CHECK, CHECK, CHECK, PART, PART]],
            ["Post-residential transition support", [CHECK, CHECK, CHECK, DASH, CHECK]],
            ["Clinically credentialed counsellor", [CHECK, CHECK, CHECK, CHECK, CHECK]],
            ["CACCF / national certification", [CHECK, CHECK, CHECK, DASH, CHECK]],
            ["Lived recovery experience on team", [CHECK, PART, CHECK, CHECK, CHECK]],
            ["Family/intervention services", [CHECK, CHECK, DASH, DASH, PART]],
            ["Case management", [PART, CHECK, DASH, DASH, PART]],
            ["Outpatient clinical anchor", [DASH, DASH, DASH, CHECK, CHECK]],
            ["Local community / referral network", [DASH, DASH, DASH, CHECK, CHECK]],
            ["Transparent pricing", [DASH, DASH, DASH, DASH, CHECK]],
          ]).map((r, i) => featureRow(r[0], r[1], i % 2 === 0 ? LGRAY : WHITE))
        ]
      }),
      spacer(100),

      h2("National Canada — SWOT Analysis"),
      spacer(40),
      swotTable(
        ["EPIC's in-person, community-embedded model cannot be replicated by virtual-only nationals", "CACCF credentialing is achievable and matches the strongest national operator (Michael Walsh)", "EPIC's outpatient clinical program is a structural differentiator no national virtual coach possesses", "Local pricing advantage — London cost of living allows competitive rates with stronger margins than Toronto/Vancouver operators"],
        ["National operators have far greater SEO authority and online visibility", "Michael Walsh's brand recognition in Canada is substantial — years of content investment", "Limited geographic reach in early pilot phase vs. national operators who can serve anyone online"],
        ["National operators confirm market exists and clients pay — validates business case entirely", "No national operator is actively building presence in London/SW Ontario — gap is explicit", "Partnership with Sober on Demand US could add cross-border referral volume nationals cannot access", "EPIC's position as Canada's first Recovery Maintenance Centre is a defensible brand claim"],
        ["Michael Walsh can aggressively target Ontario market via online delivery with minimal cost", "National operators have existing client relationships and testimonial libraries that build trust quickly", "A well-funded national operator could expand to London with reasonable investment"]
      ),
      spacer(120),

      pageBreak(),

      // ═══════════════════════════════════════════════
      // SECTION 4: INTERNATIONAL
      // ═══════════════════════════════════════════════
      h1("TIER 4 — International Market"),
      bodyText("The international market — particularly the US — is where the sober companion model is most developed, most professionalized, and most instructive for understanding where the Canadian market is heading. The four international operators profiled below represent the range from celebrity-focused luxury services to clinical concierge platforms."),
      spacer(80),

      h2("International Competitors — Profile Summary"),
      spacer(40),

      // Sober on Demand
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2200, 7160],
        rows: [
          new TableRow({ children: [headerCell("COMPETITOR", NAVY, WHITE, 2200), headerCell("Sober on Demand  (soberondemand.com)  — US National / International", ACCENT, WHITE, 7160)] }),
          ...[
            ["Founded by", "Dr. Cali Estes, Ph.D. — Vero Beach, FL; nationwide and international delivery"],
            ["Services", "Sober companions (24/7), mental health companions, concierge medical detox, 5-Day Executive Reset, luxury retreats, interventions, adolescent services"],
            ["Pricing", "Not publicly listed — premium tier; industry rates $800–$2,500+ USD/day for companions"],
            ["Companions", "1,000+ certified, bonded, insured contractors; matched by personality and background"],
            ["Credentials", "All companions certified, bonded, insured; min. 5 years sobriety; many former executives, physicians, athletes"],
            ["Marketing", "Strong brand presence; CNN, FOX, KTLA, Forbes, NBC coverage; celebrity endorsements (Steve Harvey, Nikki Sixx); YouTube channel; podcast; active social media"],
            ["Positioning", "'The Calm in Your Chaos' — discreet, premium, immediate; for those who 'cannot step away from life'"],
            ["Relevant to EPIC", "CEO contact is an active opportunity; affiliate/referral arrangement could bring US client volume to Canada"],
          ].map((r, i) => new TableRow({ children: [
            dataCell(r[0], i % 2 === 0 ? LGRAY : WHITE, 2200),
            dataCell(r[1], i % 2 === 0 ? LGRAY : WHITE, 7160)
          ]}))
        ]
      }),
      spacer(80),

      // ALYST Health
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2200, 7160],
        rows: [
          new TableRow({ children: [headerCell("COMPETITOR", NAVY, WHITE, 2200), headerCell("ALYST Health  (alysthealth.com)  — US National", ACCENT, WHITE, 7160)] }),
          ...[
            ["Location", "Los Angeles, CA — US national"],
            ["Services", "At-home addiction treatment, sober companions (Certified Recovery Agents), sober transport, interventions, clinical care integration"],
            ["Pricing", "$75–$100 USD/hour for hourly services; live-in priced as daily/weekly retainer (not per-hour calculation)"],
            ["Model", "Full continuum of care — clinical + companion integrated; positions companion (CRA) as part of clinical team"],
            ["Marketing", "Professional website; active content marketing; clinical-forward messaging"],
            ["Positioning", "'At-home rehab' — companion as clinical partner, not just support figure; highest clinical integration of any profiled competitor"],
            ["Relevance", "Their pricing transparency and clinical integration model is the best benchmark for premium pricing and scope-of-practice definition"],
          ].map((r, i) => new TableRow({ children: [
            dataCell(r[0], i % 2 === 0 ? LGRAY : WHITE, 2200),
            dataCell(r[1], i % 2 === 0 ? LGRAY : WHITE, 7160)
          ]}))
        ]
      }),
      spacer(80),

      // Active Recovery Companions
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2200, 7160],
        rows: [
          new TableRow({ children: [headerCell("COMPETITOR", NAVY, WHITE, 2200), headerCell("Active Recovery Companions  (activerecoverycompanions.com)  — US National", ACCENT, WHITE, 7160)] }),
          ...[
            ["Founded by", "Taylor Wilson, Certified Recovery Specialist; Los Angeles-based, US expansion"],
            ["Services", "Sober companions, recovery coaching, mental health companions, sober transport"],
            ["Pricing", "Hourly coaching $200–$400 USD; companion daily rates $1,000–$4,000 USD depending on intensity"],
            ["Model", "Peer-founded, non-clinical — explicitly disclaims being licensed medical/clinical professionals"],
            ["Marketing", "Active content marketing; city-specific landing pages (Miami, NYC, LA, London UK); strong SEO"],
            ["Positioning", "'Rehab isn't the finish line' — transition-focused; positions as the bridge between treatment and real life"],
            ["Relevance", "Their pricing data and service descriptions are the most detailed publicly available and provide the clearest market rate benchmarks"],
          ].map((r, i) => new TableRow({ children: [
            dataCell(r[0], i % 2 === 0 ? LGRAY : WHITE, 2200),
            dataCell(r[1], i % 2 === 0 ? LGRAY : WHITE, 7160)
          ]}))
        ]
      }),
      spacer(80),

      // Spearhead Health
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2200, 7160],
        rows: [
          new TableRow({ children: [headerCell("COMPETITOR", NAVY, WHITE, 2200), headerCell("Spearhead Health  (spearheadhealth.com)  — US Concierge / Family Office", ACCENT, WHITE, 7160)] }),
          ...[
            ["Location", "US — nationwide care management and companion coordination"],
            ["Services", "At-home detox, concierge home care, care management, companion coordination, recovery coaching, safe transport, mentoring, treatment placement"],
            ["Pricing", "$800–$2,500 USD/day for sober companion services (2025 published benchmark); hourly coaching available separately"],
            ["Model", "Care coordination — vets companions, coordinates with clinical team; designed for families and family offices managing HNW clients"],
            ["Marketing", "Professional site; trust-focused; targets legal, medical, and trust advisor referral networks — very sophisticated B2B positioning"],
            ["Positioning", "Premium vetting and coordination service — positions itself as the quality assurance layer over individual companions"],
            ["Relevance", "Spearhead's pricing guide is the most reliable and current public price benchmark in the industry; their B2B referral model (lawyers, physicians, trust advisors) is worth replicating for EPIC"],
          ].map((r, i) => new TableRow({ children: [
            dataCell(r[0], i % 2 === 0 ? LGRAY : WHITE, 2200),
            dataCell(r[1], i % 2 === 0 ? LGRAY : WHITE, 7160)
          ]}))
        ]
      }),
      spacer(100),

      h2("International — Pricing Benchmark Summary"),
      spacer(40),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2800, 2000, 2000, 2560],
        rows: [
          new TableRow({ children: [
            headerCell("Service Level", NAVY, WHITE, 2800),
            headerCell("US Rate (USD)", NAVY, WHITE, 2000),
            headerCell("CAD Equivalent", NAVY, WHITE, 2000),
            headerCell("Notes", NAVY, WHITE, 2560)
          ]}),
          ...[
            ["Hourly recovery coaching", "$75–$100 / hr", "~$105–$140 / hr CAD", "ALYST Health benchmark; Clean Livin CA rates $175 CAD/session"],
            ["Session-based coaching (50 min)", "$150–$250 / session", "$175–$350 / session CAD", "Clean Livin at $175+GST — Canadian public benchmark"],
            ["Monthly coaching package", "$1,500–$4,000 / mo", "$2,100–$5,600 / mo CAD", "Varies by frequency; most not publicly listed"],
            ["Companion — daily rate (standard)", "$800–$1,000 / day", "$1,100–$1,400 / day CAD", "Spearhead 2025 benchmark; Addictions Academy confirms $300–$2,500"],
            ["Companion — daily rate (premium)", "$1,500–$2,500 / day", "$2,100–$3,500 / day CAD", "High-risk, executive, celebrity tier; Sober on Demand range"],
            ["Companion — live-in (monthly)", "$25,000–$75,000 / mo", "$35,000–$105,000 / mo CAD", "30-day full-time; varies by market and risk profile"],
          ].map((r, i) => new TableRow({ children: [
            dataCell(r[0], i % 2 === 0 ? LGRAY : WHITE, 2800),
            dataCell(r[1], i % 2 === 0 ? LGRAY : WHITE, 2000),
            dataCell(r[2], i % 2 === 0 ? LGRAY : WHITE, 2000),
            dataCell(r[3], i % 2 === 0 ? LGRAY : WHITE, 2560)
          ]}))
        ]
      }),
      spacer(100),

      h2("International — SWOT Analysis (vs. International Operators)"),
      spacer(40),
      swotTable(
        ["EPIC operates at lower overhead than US operators, allowing competitive CAD pricing with strong margins", "No international operator has London Ontario in-person presence", "CACCF credentialing is internationally recognized and meets or exceeds US companion certification standards", "Active contact with Sober on Demand CEO is a real strategic asset — no other Canadian operator has this relationship"],
        ["International operators (especially Sober on Demand) have massive brand equity and media visibility EPIC cannot match short-term", "US operators can serve Canadian clients virtually with USD pricing that undercuts on conversion in some tiers", "EPIC's early-stage status vs. established operators with decades of testimonials and referral networks"],
        ["Sober on Demand affiliate/referral relationship could funnel US client volume directly to EPIC for Canadian assignments", "Canadian market is explicitly identified as underserved by US operators — organic referral from international networks possible", "EPIC could position as the premium Canadian alternative to US platforms for Canadian clients who prefer domestic service providers"],
        ["US platforms with strong SEO and content authority dominate Google results for 'sober companion Canada'", "Currency fluctuation makes USD-based pricing unpredictable for Canadian clients — EPIC's CAD pricing is a clarity advantage", "US operators could enter Canadian market directly with relatively low friction if the opportunity becomes obvious"]
      ),
      spacer(120),

      pageBreak(),

      // ═══════════════════════════════════════════════
      // SECTION 5: EPIC HYPOTHETICAL POSITION
      // ═══════════════════════════════════════════════
      h1("EPIC Addiction Recovery — Hypothetical Competitive Position"),
      bodyText("The following section presents EPIC's positioning as a sober companion provider at London, Ontario scale — based on what is known about EPIC's current services, the CACCF-certified counsellor leading the pilot, and the intelligence gathered across all four competitive tiers."),
      spacer(80),

      h2("Proposed Market Position Statement"),
      spacer(40),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
        rows: [new TableRow({ children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 10, color: BLUE }, bottom: { style: BorderStyle.SINGLE, size: 10, color: BLUE }, left: nb, right: nb },
          shading: { fill: LGRAY, type: ShadingType.CLEAR },
          margins: { top: 160, bottom: 160, left: 200, right: 200 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: sp(0,80), children: [new TextRun({ text: "EPIC is Southwestern Ontario's first and only in-person, clinically credentialed sober companion and recovery coaching service — the trusted bridge between structured treatment and sustainable daily life.", font: "Calibri", size: 26, bold: true, color: BLUE, italics: true })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: sp(0,0), children: [new TextRun({ text: "Not a sponsor. Not a therapist. Not a virtual platform. A credentialed professional who shows up in your life.", font: "Calibri", size: 22, color: DGRAY })] })
          ]
        })]})]}),
      spacer(100),

      h2("EPIC — Proposed Service & Pricing Structure"),
      spacer(40),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2400, 2000, 2000, 2960],
        rows: [
          new TableRow({ children: [
            headerCell("Service Tier", NAVY, WHITE, 2400),
            headerCell("Format", NAVY, WHITE, 2000),
            headerCell("Proposed Rate (CAD)", NAVY, WHITE, 2000),
            headerCell("Target Client", NAVY, WHITE, 2960)
          ]}),
          ...[
            ["Recovery Coaching Session", "50-min, in-person or virtual", "$175–$200 + tax / session", "EPIC existing clients; community referrals; harm-reduction or abstinence goals"],
            ["Monthly Coaching Package", "4–8 sessions/month + check-ins", "$1,200–$2,000 / month", "Post-residential clients; structured relapse prevention; regular accountability"],
            ["Transition Companion — Intensive", "Daily in-person, 4–8 hrs/day", "$400–$700 / day", "Clients leaving residential treatment; high-risk early recovery window (30–90 days)"],
            ["Event / Situational Companion", "Single event or specific period", "$300–$500 / event or day", "Weddings, family events, work functions, travel accompaniment"],
            ["Full Sober Companion — Live-In", "24/7 presence; 30–90 day min.", "$1,000–$1,500 / day", "Premium private-pay clients; executive / professional; post-relapse high-risk"],
          ].map((r, i) => new TableRow({ children: [
            dataCell(r[0], i % 2 === 0 ? LGRAY : WHITE, 2400),
            dataCell(r[1], i % 2 === 0 ? LGRAY : WHITE, 2000),
            dataCell(r[2], i % 2 === 0 ? LGRAY : WHITE, 2000),
            dataCell(r[3], i % 2 === 0 ? LGRAY : WHITE, 2960)
          ]}))
        ]
      }),
      spacer(60),
      para([run("Pricing is deliberately set below international rates to reflect London's market and cost-of-living realities while maintaining positioning at the premium end of the Canadian market. As volume and brand recognition build, rates can be revisited.")]),
      spacer(100),

      h2("EPIC — Full Competitive Feature Matrix"),
      spacer(40),
      new Table({ width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2360, 900, 900, 900, 900, 900],
        rows: [
          new TableRow({ children: [
            headerCell("Feature / Differentiator", NAVY, WHITE, 2360),
            headerCell("Michael\nWalsh (CA)", NAVY, WHITE, 900),
            headerCell("Sober on\nDemand (US)", NAVY, WHITE, 900),
            headerCell("Evolve\nSobriety (CA)", NAVY, WHITE, 900),
            headerCell("Active RC\n(US)", NAVY, WHITE, 900),
            headerCell("EPIC\nPilot", NAVY, WHITE, 900),
          ]}),
          ...([
            ["In-person London, Ontario", [DASH, DASH, DASH, DASH, HOT]],
            ["CACCF / national credential", [CHECK, PART, CHECK, DASH, HOT]],
            ["Lived recovery experience (18+ yrs)", [PART, PART, PART, CHECK, CHECK]],
            ["Outpatient clinical org behind it", [DASH, DASH, DASH, DASH, CHECK]],
            ["Residential referral pipeline", [DASH, DASH, DASH, DASH, CHECK]],
            ["Live-in sober companion available", [CHECK, CHECK, PART, CHECK, HOT]],
            ["Transition companion (part-time)", [CHECK, CHECK, CHECK, CHECK, CHECK]],
            ["Scheduled recovery coaching", [CHECK, PART, CHECK, CHECK, CHECK]],
            ["Event / situational companion", [CHECK, CHECK, DASH, CHECK, CHECK]],
            ["Mental health companion", [DASH, CHECK, PART, CHECK, PART]],
            ["Transparent public pricing", [DASH, DASH, DASH, PART, CHECK]],
            ["CAD pricing (no conversion needed)", [CHECK, DASH, CHECK, DASH, CHECK]],
            ["Family support services", [CHECK, PART, CHECK, DASH, PART]],
            ["US affiliate referral pathway", [DASH, HOT, DASH, DASH, HOT]],
            ["Shame-free / science-grounded model", [CHECK, PART, CHECK, PART, CHECK]],
          ]).map((r, i) => featureRow(r[0], r[1], i % 2 === 0 ? LGRAY : WHITE))
        ]
      }),
      spacer(100),

      h2("EPIC — Hypothetical SWOT Analysis (London Scale)"),
      spacer(40),
      swotTable(
        [
          "Only in-person, credentialed sober companion in London and Southwestern Ontario — uncontested local first-mover",
          "EPIC's existing outpatient program is a structural competitive moat no stand-alone operator can replicate",
          "Companion role is led by someone with nearly two decades of personal recovery AND clinical credentials — rare combination",
          "Immediate warm referral pipeline from EPIC's existing client base — no cold acquisition required to launch",
          "CACCF certification matches or exceeds credential standards of the best national competitors",
          "Transparent pricing will differentiate from every major Canadian and most US competitors",
          "Active relationship with Sober on Demand CEO creates a unique cross-border partnership pathway",
        ],
        [
          "Limited capacity at pilot stage — single companion creates bottleneck on volume",
          "No existing brand recognition in the sober companion category specifically",
          "Lower online content authority vs. established national operators (SEO is a gap to close)",
          "London's smaller private-pay market means premium client pool requires more active cultivation",
          "Program is new — no testimonials, case studies, or outcome data yet to anchor trust",
        ],
        [
          "Virtual and telephone coaching sessions extend EPIC's geographic reach across SW Ontario immediately",
          "Referral relationships with residential treatment centres (CAMH, private facilities) can funnel post-discharge clients",
          "Sober on Demand affiliate arrangement could provide Canadian client volume from US-generated leads",
          "Serving a single corporate or professional employer with EAP-style companion access could anchor recurring revenue",
          "Growing public awareness of recovery support gaps (HART hub investment, media coverage) creates tailwind",
          "Expanding to Kitchener/Waterloo, Windsor, or Hamilton is achievable once London model is proven",
        ],
        [
          "Michael Walsh's online presence dominates Canadian sober companion search results — SEO investment needed",
          "Lack of regulation means under-qualified operators can use identical language without accountability",
          "A single major client crisis or boundary incident early in the program could damage EPIC's broader reputation",
          "National virtual providers can serve London clients at lower overhead — must clearly articulate in-person value",
          "EPIC's outpatient brand and companion brand must be managed carefully to avoid client confusion about scope",
        ]
      ),
      spacer(120),

      // ═══════════════════════════════════════════════
      // SECTION 6: STRATEGIC IMPLICATIONS
      // ═══════════════════════════════════════════════
      h1("Strategic Implications — What This Means for the Pilot"),
      spacer(40),

      h3("1. The Opportunity Is Real and Uncontested Locally"),
      bodyText("No competitor analysis of four geographic tiers surfaces a single in-person, credentialed sober companion operating in London, Ontario. The competitive risk is entirely from virtual national providers — not local competitors. The first-mover advantage is substantial and available now."),
      spacer(80),

      h3("2. Credentialing Is the Moat"),
      bodyText("The strongest national Canadian operator (Michael Walsh) leads with CACCF certification. The strongest US operator (Sober on Demand) requires certification, bonding, and insurance of all companions. Completing CACCF credentialing before or during the pilot launch is the single highest-ROI preparation step — it matches the market leader's standard and immediately differentiates EPIC from uncredentialed operators."),
      spacer(80),

      h3("3. Transparent Pricing Is a Differentiator"),
      bodyText("Virtually every competitor in every tier hides pricing behind a consultation requirement. Transparent, published rate tiers — even approximate ranges — build client trust and reduce sales friction significantly. EPIC should consider publishing at minimum its coaching session rates and package structures. This alone will distinguish EPIC from every major competitor."),
      spacer(80),

      h3("4. The Sober on Demand Relationship Is Worth Pursuing Deliberately"),
      bodyText("No Canadian operator has a formal relationship with Sober on Demand. An affiliate or referral arrangement — even informal — would give EPIC access to US-generated Canadian client leads and the credibility of Dr. Estes' brand association. The CEO contact is a strategic asset that should be activated with a clear, prepared ask."),
      spacer(80),

      h3("5. SEO and Content Investment Is the Critical Long-Term Gap"),
      bodyText("Michael Walsh's national dominance is built almost entirely on content marketing and SEO — not service differentiation. EPIC should begin publishing addiction recovery content targeted to London and Southwestern Ontario immediately. A blog, local landing pages, and Google Business profile optimized for 'recovery coach London Ontario' and 'sober companion Ontario' will close this gap over 6–12 months. It is currently wide open."),
      spacer(120),

      // Closing box
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
        rows: [new TableRow({ children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 10, color: BLUE }, bottom: { style: BorderStyle.SINGLE, size: 10, color: BLUE }, left: nb, right: nb },
          shading: { fill: LGRAY, type: ShadingType.CLEAR },
          margins: { top: 160, bottom: 160, left: 200, right: 200 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: sp(0,80), children: [new TextRun({ text: "Bottom Line", font: "Calibri", size: 26, bold: true, color: BLUE })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: sp(0,0), children: [new TextRun({ text: "London, Ontario has no sober companion. Ontario has too few. Canada has some. The world has proven the model works. EPIC has the infrastructure, the credentials, the relationships, and the timing to step into a gap that nobody else in this market has filled — and to fill it well.", font: "Calibri", size: 22, italics: true, color: TEXT })] })
          ]
        })]})]}),
      spacer(60),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/EPIC_Competitive_Analysis_2026.docx', buf);
  console.log('Done');
});
