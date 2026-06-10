const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, Header, Footer
} = require('docx');
const fs = require('fs');

// Brand colors
const NAVY     = "1B3A6B";
const BLUE     = "1B4F8A";
const ACCENT   = "2E75B6";
const GREEN    = "1A7A4A";
const AMBER    = "B8670A";
const LGRAY    = "F4F6FA";
const MGRAY    = "D9DDE6";
const DGRAY    = "444444";
const WHITE    = "FFFFFF";
const TEXT     = "1E1E1E";

// Border utilities
const nb = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: nb, bottom: nb, left: nb, right: nb };
const thinBorder = (c = MGRAY) => ({ style: BorderStyle.SINGLE, size: 4, color: c });
const thinBorders = (c = MGRAY) => ({ top: thinBorder(c), bottom: thinBorder(c), left: thinBorder(c), right: thinBorder(c) });

// Helper functions
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

// Create document
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
            children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Sober Companion — Competitive Analysis | June 2026", font: "Calibri", size: 18, italics: true, color: DGRAY })] })] })
        ]})]
      }),
      new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 } }, spacing: sp(0,0), children: [] })
    ]})},
    footers: { default: new Footer({ children: [
      new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 4, color: MGRAY, space: 1 } }, spacing: sp(60,0), alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Confidential — For internal planning purposes only  |  epicrecovery.ca", font: "Calibri", size: 18, color: "888888" })] })
    ]})},

    children: [
      // COVER
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

      // Executive Summary
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
        rows: [new TableRow({ children: [new TableCell({
          borders: thinBorders(ACCENT),
          shading: { fill: LGRAY, type: ShadingType.CLEAR },
          margins: { top: 160, bottom: 160, left: 200, right: 200 },
          children: [
            new Paragraph({ spacing: sp(0,80), children: [new TextRun({ text: "Executive Summary", font: "Calibri", size: 26, bold: true, color: BLUE })] }),
            new Paragraph({ spacing: sp(0,60), children: [new TextRun({ text: "The sober companion and recovery coaching market is active at the national and international levels, but virtually absent in London, Ontario. This is an opportunity.", font: "Calibri", size: 22, color: TEXT })] }),
            new Paragraph({ spacing: sp(0,0), children: [new TextRun({ text: "This document profiles competitors across four geographic tiers, presents feature matrices and SWOT analyses for each tier, and positions EPIC's hypothetical sober companion service within the competitive landscape.", font: "Calibri", size: 22, color: TEXT })] })
          ]
        })]})]}),

      spacer(120),

      h1("TIER 1 — London, Ontario Market"),
      bodyText("A thorough search for dedicated sober companion or recovery coaching services physically based in London, Ontario returns a stark result: there are none. The London addiction recovery ecosystem is dominated by public system providers and peer support groups."),
      spacer(80),

      h2("Active Providers — London, Ontario"),
      spacer(40),

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
      spacer(100),

      h3("Key Finding: London Is an Empty Field"),
      bodyText("There is no professional, private-pay sober companion service operating in London, Ontario. This represents a genuine first-mover competitive opportunity."),
      spacer(120),

      h2("London Market — SWOT Analysis"),
      spacer(40),
      swotTable(
        [
          "Zero direct local competition — genuine first-mover position",
          "EPIC already has brand presence, referral network, and private-pay clientele in London",
          "Unmet demand from private rehab graduates with no transition support options"
        ],
        [
          "London is a smaller market than Toronto — premium client pool is smaller",
          "No established consumer awareness of sober companion services locally",
          "Limited capacity at pilot stage"
        ],
        [
          "Unmet demand from private rehab graduates with no transition support options",
          "Referral pipeline from EPIC's existing outpatient program is immediate and organic",
          "Growing public awareness of recovery support gaps"
        ],
        [
          "National virtual providers (Michael Walsh, Evolve Sobriety) can serve London clients remotely",
          "Public system improvements could reduce urgency for private-pay clients over time"
        ]
      ),
      spacer(120),

      h1("Key Competitive Insights"),
      bodyText("The analysis reveals that EPIC's greatest advantage is its in-person, clinically credentialed presence in an underserved market. Competitors operate either nationally (virtually) or in major metropolitan centers, leaving Southwestern Ontario largely uncontested."),
      spacer(60),

      h2("Competitive Positioning Summary"),
      spacer(40),
      para([
        boldRun("Strengths: "),
        run("Local presence, existing client base, clinical credentialing, and transparent pricing differentiate EPIC from all national and international competitors."),
        run("\n\n"),
        boldRun("Weaknesses: "),
        run("Limited brand recognition in the sober companion category and lower SEO authority than established operators."),
        run("\n\n"),
        boldRun("Opportunities: "),
        run("Virtual coaching extends geographic reach, residential treatment center partnerships generate referrals, and affiliate arrangements with US operators could create revenue pathways."),
        run("\n\n"),
        boldRun("Threats: "),
        run("Virtual-only providers have lower overhead, and early-stage operational risks could damage EPIC's broader reputation.")
      ]),
      spacer(100),

      h2("Strategic Recommendations"),
      spacer(40),
      para([
        run("1. "),
        boldRun("Invest in SEO and content marketing"),
        run(" to build authority in the sober companion category — Michael Walsh's national dominance is built almost entirely on this.\n\n"),
        run("2. "),
        boldRun("Pursue formal affiliate relationships"),
        run(" with Sober on Demand and other US platforms to funnel international client volume.\n\n"),
        run("3. "),
        boldRun("Maintain transparent pricing"),
        run(" as a core differentiator — every major competitor hides rates behind consultations.\n\n"),
        run("4. "),
        boldRun("Build a testimonials and case studies library"),
        run(" early to anchor trust and convert leads.\n\n"),
        run("5. "),
        boldRun("Carefully manage the distinction"),
        run(" between EPIC's outpatient clinical brand and the companion brand to avoid scope confusion.")
      ]),
      spacer(150),

      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
        rows: [new TableRow({ children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 10, color: BLUE }, bottom: { style: BorderStyle.SINGLE, size: 10, color: BLUE }, left: nb, right: nb },
          shading: { fill: LGRAY, type: ShadingType.CLEAR },
          margins: { top: 160, bottom: 160, left: 200, right: 200 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: sp(0,80), children: [new TextRun({ text: "Bottom Line", font: "Calibri", size: 26, bold: true, color: BLUE })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: sp(0,0), children: [new TextRun({ text: "London, Ontario has no sober companion. The market exists nationally and internationally—it's validated. The opportunity is real, uncontested locally, and achievable with disciplined execution.", font: "Calibri", size: 22, color: TEXT })] })
          ]
        })]})]}),
      spacer(60),
    ]
  }]
});

// Generate Word document
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('./EPIC_Competitive_Analysis_2026.docx', buf);
  console.log('✓ Document generated: EPIC_Competitive_Analysis_2026.docx');
});
