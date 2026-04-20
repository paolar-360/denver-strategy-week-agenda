const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');
const path = require('path');

// ─── BRAND PALETTE ────────────────────────────────────────────────────────
const C = {
  black:      "1A1A1A",   // near black — day headers, main text
  orange:     "E8610A",   // brand orange — sprint days, accents
  orangeDeep: "B84D08",   // darker orange — for text on light bg
  orangeLight:"FFF0E6",   // light orange — hi-row backgrounds
  orangeMid:  "FFD4A8",   // mid orange — time cell for hi rows
  sprintBg:   "FFF3EA",   // sprint row bg
  sprintTime: "FFD4A8",   // sprint time bg
  fixedBg:    "FFF0F0",
  fixedTime:  "FFD6D6",
  fixedText:  "8B1A1A",
  presBg:     "FFFBF0",
  presTime:   "FFE8A0",
  presText:   "7A5000",
  white:      "FFFFFF",
  offWhite:   "FAFAFA",
  gray:       "F5F5F5",
  midGray:    "CCCCCC",
  textDark:   "222222",
  textMid:    "555555",
  textLight:  "888888",
  breakBg:    "F5F5F5",
  breakTime:  "EBEBEB",
  breakText:  "AAAAAA",
};

const cb  = { style: BorderStyle.SINGLE, size: 1, color: C.midGray };
const cbs = { top: cb, bottom: cb, left: cb, right: cb };
const sp  = () => new Paragraph({ spacing: { before: 50, after: 50 }, children: [new TextRun("")] });

// ─── COVER TITLE ──────────────────────────────────────────────────────────
function coverTitle() {
  return [
    // Black bar with orange accent line
    new Paragraph({ spacing: { before: 0, after: 0 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 16, color: C.orange, space: 1 } },
      shading: { fill: C.black, type: ShadingType.CLEAR },
      children: [new TextRun({ text: " ", size: 4, font: "Arial" })] }),

    new Paragraph({ alignment: AlignmentType.CENTER,
      shading: { fill: C.black, type: ShadingType.CLEAR },
      spacing: { before: 160, after: 0 },
      children: [new TextRun({ text: "360 SIERRA  \u00D7  EVOLVE  \u00D7  WILDWOOD", bold: true, size: 36, font: "Arial", color: C.white })] }),

    new Paragraph({ alignment: AlignmentType.CENTER,
      shading: { fill: C.black, type: ShadingType.CLEAR },
      spacing: { before: 60, after: 0 },
      children: [new TextRun({ text: "Strategy & Sprint Week", bold: true, size: 48, font: "Arial", color: C.orange })] }),

    new Paragraph({ alignment: AlignmentType.CENTER,
      shading: { fill: C.black, type: ShadingType.CLEAR },
      spacing: { before: 60, after: 0 },
      children: [new TextRun({ text: "April 22 \u2013 29, 2026", size: 26, font: "Arial", color: C.white })] }),

    new Paragraph({ alignment: AlignmentType.CENTER,
      shading: { fill: C.black, type: ShadingType.CLEAR },
      spacing: { before: 60, after: 160 },
      children: [new TextRun({ text: "10:00 AM \u2013 5:00 PM   \u00B7   Breaks 11:30 & 3:30   \u00B7   Lunch 1:00 PM", size: 19, font: "Arial", color: C.midGray, italics: true })] }),

    new Paragraph({ spacing: { before: 0, after: 0 },
      border: { top: { style: BorderStyle.SINGLE, size: 16, color: C.orange, space: 1 } },
      shading: { fill: C.black, type: ShadingType.CLEAR },
      children: [new TextRun({ text: " ", size: 4, font: "Arial" })] }),

    sp(),
  ];
}

// ─── LEGEND TABLE ─────────────────────────────────────────────────────────
function legendTable() {
  function hCell(text, w) {
    return new TableCell({ borders: cbs, width: { size: w, type: WidthType.DXA },
      shading: { fill: C.black, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 19, font: "Arial", color: C.white })] })] });
  }
  function dCell(text, w, bg = C.white, bold = false, color = C.textDark) {
    return new TableCell({ borders: cbs, width: { size: w, type: WidthType.DXA },
      shading: { fill: bg, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text, size: 19, font: "Arial", bold, color })] })] });
  }
  function row(who, avail, focus) {
    return new TableRow({ children: [
      dCell(who,   2000, C.orangeLight, true, C.orangeDeep),
      dCell(avail, 2700, C.offWhite, false, C.textMid),
      dCell(focus, 4660),
    ]});
  }
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 2700, 4660],
    rows: [
      new TableRow({ children: [hCell("Who", 2000), hCell("Availability", 2700), hCell("Focus", 4660)] }),
      row("Andres (CEO) & Paola",    "Full week — Wed Apr 22 to Wed Apr 29", "Vision presentation, Sprint facilitation, all sessions"),
      row("Matt & Heath (Evolve)", "Full week — Wed Apr 22 to Wed Apr 29", "Operations, design partner alignment, Sprint input"),
      row("Sam (Evolve)",            "Wed Apr 22 – Sat Apr 26 (morning)",    "Business model, revenue, commercial strategy"),
      row("David (Wildwood)",        "By session — enters & exits",           "VC lens, SAFE/legal, 1:1 with Evolve team"),
      row("Nick (Wildwood)",         "By session — enters & exits",           "Go-to-market, commercial strategy"),
    ]
  });
}

// ─── FORMAT KEY ───────────────────────────────────────────────────────────
function formatKeyTable() {
  function hCell(text, w) {
    return new TableCell({ borders: cbs, width: { size: w, type: WidthType.DXA },
      shading: { fill: C.black, type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 110, right: 110 },
      children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 19, font: "Arial", color: C.white })] })] });
  }
  function kRow(badge, badgeColor, bgColor, name, desc) {
    return new TableRow({ children: [
      new TableCell({ borders: cbs, width: { size: 1600, type: WidthType.DXA },
        shading: { fill: bgColor, type: ShadingType.CLEAR },
        margins: { top: 70, bottom: 70, left: 110, right: 110 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: badge, bold: true, size: 17, font: "Arial", color: badgeColor })] })] }),
      new TableCell({ borders: cbs, width: { size: 2200, type: WidthType.DXA },
        margins: { top: 70, bottom: 70, left: 110, right: 110 },
        children: [new Paragraph({ children: [new TextRun({ text: name, bold: true, size: 19, font: "Arial" })] })] }),
      new TableCell({ borders: cbs, width: { size: 5560, type: WidthType.DXA },
        shading: { fill: C.offWhite, type: ShadingType.CLEAR },
        margins: { top: 70, bottom: 70, left: 110, right: 110 },
        children: [new Paragraph({ children: [new TextRun({ text: desc, size: 18, font: "Arial", color: C.textMid })] })] }),
    ]});
  }
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1600, 2200, 5560],
    rows: [
      new TableRow({ children: [hCell("Format", 1600), hCell("Name", 2200), hCell("What it means", 5560)] }),
      kRow("PRESENTATION", C.presText,  C.presBg,     "360 Sierra Story",  "Andres & Paola present. No structured exercises. This is our moment."),
      kRow("SPRINT-LITE",  C.orangeDeep,C.orangeLight, "Open Sprint",       "Facilitated discussion with a note-taker at the board. Time-boxed but conversational. Used for strategy, pricing, GTM."),
      kRow("SPRINT",       C.white,     C.orange,      "Formal Sprint",     "Full GV Sprint methodology: structured exercises, individual sketching, dot voting, firm time-boxes. All product days."),
      kRow("FIXED",        C.fixedText, C.fixedBg,     "Fixed Block",       "Pre-confirmed meeting — time and participants locked."),
    ]
  });
}

// ─── DAY HEADER ───────────────────────────────────────────────────────────
function dayHdr(num, date, theme, type = "normal") {
  // type: "normal" | "lite" | "sprint"
  const isSprint = type === "sprint";
  const isLite   = type === "lite";
  const bgColor  = isSprint ? C.orange : C.black;
  const txtColor = C.white;
  const label    = isSprint ? "  \u00B7  SPRINT DAY" : (isLite ? "  \u00B7  SPRINT-LITE" : "");
  const arrow    = isSprint ? "\u25B6  " : "";

  return new Paragraph({
    spacing: { before: 340, after: 6 },
    shading: { fill: bgColor, type: ShadingType.CLEAR },
    children: [
      new TextRun({ text: "  " + arrow + "DAY " + num + "   ", bold: true, size: 28, font: "Arial", color: txtColor }),
      new TextRun({ text: date, size: 22, font: "Arial", color: isSprint ? "FFD4A8" : C.midGray }),
      new TextRun({ text: "   \u2014   " + theme, size: 22, font: "Arial", color: txtColor, italics: true }),
      new TextRun({ text: label, size: 19, font: "Arial", color: isSprint ? "FFD4A8" : C.orange, bold: true }),
    ]
  });
}

function metaLines(focus, room) {
  return [
    new Paragraph({ spacing: { before: 6, after: 4 },
      border: { left: { style: BorderStyle.SINGLE, size: 14, color: C.orange, space: 5 } },
      indent: { left: 140 },
      children: [new TextRun({ text: focus, size: 18, font: "Arial", color: C.textMid, italics: true })] }),
    new Paragraph({ spacing: { before: 0, after: 100 },
      indent: { left: 154 },
      children: [
        new TextRun({ text: "In the room: ", bold: true, size: 16, font: "Arial", color: C.textLight }),
        new TextRun({ text: room, size: 16, font: "Arial", color: C.textLight }),
      ] }),
  ];
}

function noteBar(text, color = C.orange) {
  return new Paragraph({
    spacing: { before: 40, after: 90 },
    border: { left: { style: BorderStyle.SINGLE, size: 10, color, space: 5 } },
    shading: { fill: C.orangeLight, type: ShadingType.CLEAR },
    indent: { left: 140 },
    children: [new TextRun({ text, size: 17, font: "Arial", color: C.orangeDeep, italics: true })]
  });
}

// ─── BLOCK ROW ────────────────────────────────────────────────────────────
function blk(time, title, detail, type = "normal") {
  let bgMain, bgTime, titleColor, badge;
  switch(type) {
    case "break":
      bgMain = C.breakBg; bgTime = C.breakTime; titleColor = C.breakText; badge = ""; break;
    case "fixed":
      bgMain = C.fixedBg; bgTime = C.fixedTime; titleColor = C.fixedText; badge = " FIXED"; break;
    case "presentation":
      bgMain = C.presBg;  bgTime = C.presTime;  titleColor = C.presText;  badge = " PRESENTATION"; break;
    case "hi":    // sprint-lite highlight
      bgMain = C.orangeLight; bgTime = C.orangeMid; titleColor = C.orangeDeep; badge = ""; break;
    case "sprint-hi":  // formal sprint highlight
      bgMain = C.sprintBg; bgTime = C.sprintTime; titleColor = C.orange; badge = ""; break;
    case "product":
      bgMain = "FFF3FA"; bgTime = "F5C0DC"; titleColor = "7A005A"; badge = " PRODUCTS"; break;
    default:
      bgMain = C.white; bgTime = C.gray; titleColor = C.textDark; badge = "";
  }
  const isBold = ["hi","sprint-hi","fixed","presentation","product"].includes(type);
  return new TableRow({ children: [
    new TableCell({ borders: cbs, width: { size: 1180, type: WidthType.DXA },
      shading: { fill: bgTime, type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 100, right: 100 },
      children: [new Paragraph({ alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: time, size: 17, font: "Arial", bold: true, color: titleColor })] })] }),
    new TableCell({ borders: cbs, width: { size: 2620, type: WidthType.DXA },
      shading: { fill: bgMain, type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 110, right: 110 },
      children: [new Paragraph({ children: [
        new TextRun({ text: title, size: 20, font: "Arial", bold: isBold, color: titleColor }),
        badge ? new TextRun({ text: badge, size: 16, font: "Arial", bold: true, color: titleColor }) : new TextRun(""),
      ]})] }),
    new TableCell({ borders: cbs, width: { size: 5560, type: WidthType.DXA },
      shading: { fill: bgMain, type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 110, right: 110 },
      children: [new Paragraph({ children: [new TextRun({ text: detail, size: 18, font: "Arial", color: C.textMid })] })] }),
  ]});
}

function tbl(rows, sprint = false) {
  const hBg = sprint ? C.orange : C.black;
  function h(text, w) {
    return new TableCell({ borders: cbs, width: { size: w, type: WidthType.DXA },
      shading: { fill: hBg, type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 110, right: 110 },
      children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 18, font: "Arial", color: C.white })] })] });
  }
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1180, 2620, 5560],
    rows: [new TableRow({ children: [h("Time", 1180), h("Block", 2620), h("What happens", 5560)] }), ...rows] });
}

// ─── PRODUCT PORTFOLIO BOX ────────────────────────────────────────────────
function productPortfolioSection() {
  // A styled table showing the 4 products
  function pRow(logo, name, tagline, desc, bg) {
    return new TableRow({ children: [
      new TableCell({ borders: cbs, width: { size: 1600, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: logo, size: 32, font: "Arial" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: name, bold: true, size: 20, font: "Arial", color: C.black })] }),
        ] }),
      new TableCell({ borders: cbs, width: { size: 2200, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: tagline, bold: true, size: 19, font: "Arial", color: C.orangeDeep })] })] }),
      new TableCell({ borders: cbs, width: { size: 5560, type: WidthType.DXA },
        shading: { fill: C.white, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: desc, size: 19, font: "Arial", color: C.textMid })] })] }),
    ]});
  }
  return [
    new Paragraph({ spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: "Product Portfolio", bold: true, size: 24, font: "Arial", color: C.black })] }),
    new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1600, 2200, 5560],
      rows: [
        new TableRow({ children: [
          new TableCell({ borders: cbs, width: { size: 1600, type: WidthType.DXA }, shading: { fill: C.black, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Brand", bold: true, size: 19, font: "Arial", color: C.white })] })] }),
          new TableCell({ borders: cbs, width: { size: 2200, type: WidthType.DXA }, shading: { fill: C.black, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Layer", bold: true, size: 19, font: "Arial", color: C.white })] })] }),
          new TableCell({ borders: cbs, width: { size: 5560, type: WidthType.DXA }, shading: { fill: C.black, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "What it is", bold: true, size: 19, font: "Arial", color: C.white })] })] }),
        ]}),
        pRow("\uD83C\uDFE2", "360 Sierra",   "The Company",              "The parent platform and infrastructure layer. All products are built on top of 360 Sierra.", C.orangeLight),
        pRow("\uD83E\uDD1F", "RentalBuddy",  "Operator Management",      "The front-facing product for rental operators. Fleet management, invoicing, integrations, and day-to-day operations. Powered by 360 Sierra.", C.orangeLight),
        pRow("\uD83E\uDD16", "Shakkii",      "AI Operations System",     "The intelligence layer. AI-powered automations and decision support built into the operational workflow. Powers smart features across RentalBuddy and Lemonade.", C.orangeLight),
        pRow("\uD83C\uDF4B", "Lemonade",     "Booking Platform",         "The transactional and AI-powered booking platform for renters. Handles end-to-end booking flow, payments, and renter experience. Complements the operator side.", C.orangeLight),
      ]
    }),
  ];
}

// ─── BUILD DOCUMENT ───────────────────────────────────────────────────────
const children = [

  ...coverTitle(),
  sp(),

  sp(),

  // ══════════════════════════════════════════════
  // DAY 1 — Wednesday April 22
  // ══════════════════════════════════════════════
  dayHdr(1, "Wednesday, April 22", "Vision, Strategy & Investment Structure", "lite"),
  ...metaLines(
    "360 Sierra tells its story. Open, facilitated sessions on strategy, market positioning, and investment framework.",
    "Andres, Paola, Matt, Heath, Sam  \u00B7  David & Nick join from afternoon"
  ),
  tbl([
    blk("10:00 AM", "Welcome & Week Overview",       "Opening: format, goals, 8-day agenda, Sprint methodology overview. Set expectations for how decisions will be made across the week."),
    blk("10:15 AM", "360 Sierra \u2014 Story, Vision & Product Portfolio", "Origin story, what we are building, why now, and the path to 10,000 and 50,000 vehicles. Walk through the full product architecture: 360 Sierra (company), RentalBuddy (operator management), Shakkii (AI operations), and Lemonade (booking platform). Align on naming and how each product tells part of the overall story.", "presentation"),
    blk("11:30 AM", "Break",                         "", "break"),
    blk("11:45 AM", "Strategy & Market Positioning", "Facilitated: 3\u20135 year direction, milestones, success metrics (ARR, market share). Competitive differentiation, unfair advantage, and why the market is ready. Note-taker captures on the board.", "hi"),
    blk("1:00 PM",  "Lunch",                         "", "break"),
    blk("2:00 PM",  "Investment Framework",           "Facilitated: SAFE structure, revenue share terms. Capital deployment: where $50K goes, what it unlocks. $50K vs. $200K scenario. David joins.", "hi"),
    blk("3:00 PM",  "Seed Milestones & Exit Path",   "3\u20134 milestones that unlock the Seed round. Long-term exit considerations and valuation trajectory. What does 5-year success look like?", "hi"),
    blk("3:30 PM",  "Break",                         "", "break"),
    blk("3:45 PM",  "Risk, Open Q&A & De-risking",   "Open floor for remaining investor questions. Address risk factors, execution confidence, and what de-risks the bet."),
    blk("4:30 PM",  "Day 1 Recap",                   "Capture decisions and open items on the board. David and Nick depart. Preview tomorrow."),
    blk("5:00 PM",  "End of Day",                    ""),
  ]),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 2 — Thursday April 23
  // ══════════════════════════════════════════════
  dayHdr(2, "Thursday, April 23", "GTM, Products & Design Partner", "lite"),
  ...metaLines(
    "Facilitated sessions on go-to-market, product portfolio naming, and Evolve\u2019s formal role. Evening: Wildwood dinner.",
    "Andres, Paola, Matt, Heath, Sam  \u00B7  Nick joins morning  \u00B7  David joins for 1:1 only"
  ),
  noteBar("Fixed: 1:1 Evolve \u00D7 David \u2014 2:00 to 3:00 PM (Andres and Paola are not in this session)  \u00B7  Evening: dinner at Wildwood"),
  tbl([
    blk("10:00 AM", "Recap & Open Items",              "Resolve any open loops from Day 1 before today\u2019s sessions."),
    blk("10:15 AM", "Go-to-Market Strategy",           "Facilitated: commercial model, path to 10K vehicles, sales motion, channel strategy, and partnership model. Nick joins.", "hi"),
    blk("11:30 AM", "Break",                           "", "break"),
    blk("11:45 AM", "Competitive Landscape & Differentiation", "Map the competitive field, build a positioning matrix, and articulate 360 Sierra\u2019s unfair advantage. What makes us hard to copy?", "hi"),
    blk("12:20 PM", "ICP & Pricing Strategy",          "Define the ideal customer profile (operator size, segment, geography) and the pricing logic that fits \u2014 packaging, tiers, contract length.", "hi"),
    blk("1:00 PM",  "Lunch",                           "Nick departs after lunch.", "break"),
    blk("2:00 PM",  "1:1 \u2014 Evolve \u00D7 David",             "Evolve team (Matt, Heath, Sam) and David. Private session. Andres and Paola are not present.", "fixed"),
    blk("3:00 PM",  "Evolve as Design Partner",        "Define the role formally: what Evolve contributes, what they receive, how product decisions get made together. Close all open loops before the Sprint.", "hi"),
    blk("3:30 PM",  "Break",                           "", "break"),
    blk("3:45 PM",  "Commercial Targets & Alignment",  "Align on Year 1 commercial targets. Confirm the shared 12-month execution picture."),
    blk("4:45 PM",  "Day 2 Recap",                     "Capture decisions and action items. Preview Sprint Day 1 tomorrow."),
    blk("5:00 PM",  "End of Day",                      ""),
    blk("Evening",  "Dinner \u2014 Wildwood",                  "Group dinner. Time and location confirmed separately.", "hi"),
  ]),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 3 — Friday April 24 — SPRINT: MAP
  // ══════════════════════════════════════════════
  dayHdr(3, "Friday, April 24", "Map \u2014 Problem Definition & Sprint Target", "sprint"),
  ...metaLines(
    "Map the full system. Surface real bottlenecks. Lock the Sprint Target. Sam\u2019s last full working day.",
    "Andres, Paola, Matt, Heath, Sam  \u00B7  1 Evolve ops team member (in person)"
  ),
  noteBar("Sam participates in the full Map day before departing Saturday morning."),
  tbl([
    blk("10:00 AM", "Sprint Kickoff",                 "Ground rules: no laptops during exercises, one Decider, time-boxes are firm. Goal: a completed System Map and locked Sprint Target.", "sprint-hi"),
    blk("10:20 AM", "Long-Term Goal",                 "As a group: \u201CWhy are we doing this? Where does 360 Sierra need to be in 3 years?\u201D Written on the board. Anchors every Sprint decision.", "sprint-hi"),
    blk("10:50 AM", "Sprint Questions",               "What are we most afraid of? Each risk becomes a testable question: \u201CCan we achieve X by end of this Sprint?\u201D These guide the Map.", "sprint-hi"),
    blk("11:30 AM", "Break",                          "", "break"),
    blk("11:45 AM", "System Map",                     "Draw the full workflow: vehicle onboarding \u2192 transaction \u2192 fleet dashboard \u2192 operator. Mark every actor, handoff, and integration point. Flag friction in red.", "sprint-hi"),
    blk("1:00 PM",  "Lunch",                          "", "break"),
    blk("2:00 PM",  "Expert Interviews",              "Structured 10\u201315 min conversations with each person in the room, including the Evolve ops team member. Facilitator takes How Might We (HMW) notes. Real operational insight surfaces here.", "sprint-hi"),
    blk("3:00 PM",  "HMW Sort & Dot Vote",            "Post all HMW notes on the map. Vote with dots. Top clusters reveal the highest-value problems.", "sprint-hi"),
    blk("3:30 PM",  "Break",                          "", "break"),
    blk("3:45 PM",  "Sprint Target \u2014 Decision",         "Decider picks one moment on the map and one Sprint Question to answer. Locked focus for Days 6 and 7. No pivoting after this.", "sprint-hi"),
    blk("4:15 PM",  "Technical Goals & Baselines",    "Set current baselines: reliability, latency, integration count. Define what these numbers should look like after Sprint execution.", "sprint-hi"),
    blk("4:45 PM",  "Day 3 Recap",                    "Photograph the wall. Document the Map and Sprint Target. Preview Monday.", "sprint-hi"),
    blk("5:00 PM",  "End of Day",                     "", "sprint-hi"),
  ], true),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 4 — Saturday April 25 — BACKUP
  // ══════════════════════════════════════════════
  dayHdr(4, "Saturday, April 25", "Backup / Overflow \u2014 Morning Only"),
  ...metaLines(
    "Optional buffer morning. Sam departs today.",
    "Andres, Paola, Matt, Heath, Sam (morning only)"
  ),
  tbl([
    blk("10:00 AM", "Morning Check-in",               "Assess where the week stands. Decide together how to use the morning."),
    blk("10:15 AM", "Option A \u2014 Overflow",               "Resolve any open items from Days 1\u20133: investment terms, GTM, design partner scope, or unfinished Map work."),
    blk("10:15 AM", "Option B \u2014 Deep-Dive",              "Go deeper on one topic: AI architecture, product naming, pricing, or competitive landscape."),
    blk("10:15 AM", "Option C \u2014 Informal",               "Rest or unstructured time. Relationship-building matters."),
    blk("12:00 PM", "Sam Farewell",                   "Sam departs. Informal close to the morning.", "hi"),
    blk("12:30 PM", "End of Morning",                 "No afternoon session."),
  ]),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 5 — Sunday April 26
  // ══════════════════════════════════════════════
  dayHdr(5, "Sunday, April 26", "Rest Day"),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 100 },
    children: [new TextRun({ text: "\u2605  No scheduled activities  \u2605", size: 22, font: "Arial", color: C.midGray })] }),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 6 — Monday April 27 — SPRINT: SKETCH
  // ══════════════════════════════════════════════
  dayHdr(6, "Monday, April 27", "Sketch \u2014 Ideation & AI Strategy", "sprint"),
  ...metaLines(
    "Generate solutions independently. Compare through structure, not debate. Define the AI module strategy.",
    "Andres, Paola, Matt, Heath"
  ),
  tbl([
    blk("10:00 AM", "Recap & Sprint Target Review",  "Reconnect with the Map and Sprint Target from Friday. Restate the Long-Term Goal.", "sprint-hi"),
    blk("10:20 AM", "Lightning Demos",                "3 minutes each: one inspiring example relevant to the Sprint Target \u2014 a product, workflow, or idea. Facilitator captures \u201Cbig ideas.\u201D", "sprint-hi"),
    blk("11:00 AM", "Crazy 8s",                       "Individual exercise: fold paper into 8 panels. Sketch 8 different solution ideas in 8 minutes. Forces quantity over quality. Solo work only.", "sprint-hi"),
    blk("11:30 AM", "Break",                          "", "break"),
    blk("11:45 AM", "Solution Sketch",                "Each person develops a detailed 3-panel sketch for the Sprint Target. Anonymous \u2014 no names on sketches. Reviewed and voted on tomorrow.", "sprint-hi"),
    blk("1:00 PM",  "Lunch",                          "", "break"),
    blk("2:00 PM",  "AI Module Strategy",             "Where do Shakkii\u2019s AI modules create the most leverage? Parallel vs. sequential deployment, speed-to-market trade-offs, which operations to automate first, build vs. integrate decisions.", "sprint-hi"),
    blk("3:00 PM",  "Gallery Walk",                   "Post all sketches on the wall. Everyone reviews silently, leaving dot stickers on interesting ideas. No commentary yet.", "sprint-hi"),
    blk("3:30 PM",  "Break",                          "", "break"),
    blk("3:45 PM",  "Heat Map & Sketch Pitches",      "Review where dots cluster. Each person has 1 minute to describe their sketch \u2014 no defending, just context. Sets up tomorrow\u2019s decision.", "sprint-hi"),
    blk("4:30 PM",  "Day 6 Recap",                    "Document all sketches. Preview the Decide day.", "sprint-hi"),
    blk("5:00 PM",  "End of Day",                     "", "sprint-hi"),
  ], true),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 7 — Tuesday April 28 — SPRINT: DECIDE
  // ══════════════════════════════════════════════
  dayHdr(7, "Tuesday, April 28", "Decide \u2014 Roadmap Lock-In", "sprint"),
  ...metaLines(
    "Open conversation with Native in the morning. Vote, decide, and lock the full product roadmap in the afternoon.",
    "Andres, Paola, Matt, Heath  \u00B7  1 Evolve ops member (online)  \u00B7  Native Camper Vans (online, morning only)"
  ),
  noteBar("The Native session is an open conversation \u2014 casual and honest. We want to hear how operating a fleet actually feels, what\u2019s frustrating, and what they wish existed."),
  tbl([
    blk("10:00 AM", "Conversation \u2014 Native + Evolve ops", "Open chat with Native Camper Vans and the Evolve ops member (online). How does running a fleet feel today? What is slow, painful, or missing? Facilitator takes notes. (~75 min)", "sprint-hi"),
    blk("11:30 AM", "Break",                           "", "break"),
    blk("11:45 AM", "Synthesis & Speed Critique",     "Synthesize the conversation insights. Add to the board. Then 5 min per sketch: facilitator narrates, author stays silent, group asks clarifying questions only.", "sprint-hi"),
    blk("1:00 PM",  "Lunch",                          "Online guests disconnect before lunch.", "break"),
    blk("2:00 PM",  "Decision \u2014 Straw Poll",             "Everyone votes for the sketch they want to build. Votes converge \u2014 done. Split \u2014 Decider calls it. One direction, locked.", "sprint-hi"),
    blk("2:30 PM",  "Roadmap \u2014 RentalBuddy / Transactional Core", "Development sequence: onboarding, invoicing, real-time fleet data, integrations. Priorities, milestones, dependencies.", "sprint-hi"),
    blk("3:30 PM",  "Break",                          "", "break"),
    blk("3:45 PM",  "Roadmap \u2014 Shakkii / AI Layer",      "Which AI modules ship in MVP, which in v1.1, which are post-Seed. Build vs. integrate decisions. How Shakkii connects to RentalBuddy and Lemonade.", "sprint-hi"),
    blk("4:15 PM",  "Full Week Review & Lock",        "Walk through every decision made across the whole week. Confirm nothing is open. Final alignment check before departure day.", "sprint-hi"),
    blk("4:45 PM",  "Closing Remarks",                 "Acknowledge the work done. Next steps and communication cadence.", "sprint-hi"),
    blk("5:00 PM",  "End of Day",                     "", "sprint-hi"),
  ], true),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 8 — Wednesday April 29
  // ══════════════════════════════════════════════
  dayHdr(8, "Wednesday, April 29", "Wrap-Up & Departures"),
  ...metaLines(
    "Final roadmap confirmation, action items, investment next steps, and travel.",
    "Andres, Paola, Matt, Heath"
  ),
  tbl([
    blk("10:00 AM", "Final Roadmap Review",           "One last pass through the locked roadmap. All priorities, timelines, and open items confirmed."),
    blk("10:45 AM", "Action Items \u2014 All Parties",        "Every action item listed with owner and deadline: 360 Sierra, Evolve, Wildwood. Reporting cadence confirmed.", "hi"),
    blk("11:30 AM", "Break",                          "", "break"),
    blk("11:45 AM", "Investment \u2014 Next Steps",          "Transfer timeline, SAFE execution steps, reporting structure."),
    blk("12:15 PM", "Closing Lunch / Farewell",       "Informal lunch. Safe travels.", "hi"),
    blk("Afternoon", "Departures",                    "Travel as scheduled."),
  ]),

  sp(), sp(),

  // PRODUCT PORTFOLIO (reference section at end)
  ...productPortfolioSection(),
  sp(),

  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
    border: { top: { style: BorderStyle.SINGLE, size: 3, color: C.orange, space: 5 } },
    children: [
      new TextRun({ text: "360 Sierra  \u00B7  ", size: 16, font: "Arial", color: C.orange, bold: true }),
      new TextRun({ text: "Strategy Week 2026  \u00B7  Confidential", size: 16, font: "Arial", color: C.midGray }),
    ] }),
];

// ─── DOCUMENT ─────────────────────────────────────────────────────────────
function buildDoc() {
  return new Document({
    numbering: { config: [] },
    styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1008, right: 1008, bottom: 1008, left: 1008 } } },
      children
    }]
  });
}

async function generateBuffer() {
  return Packer.toBuffer(buildDoc());
}

module.exports = { buildDoc, generateBuffer };

if (require.main === module) {
  generateBuffer().then(buf => {
    const out = path.join(__dirname, 'Strategy_Week_Agenda.docx');
    fs.writeFileSync(out, buf);
    console.log('Done:', out);
  });
}
