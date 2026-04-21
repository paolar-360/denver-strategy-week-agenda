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
      children: [new TextRun({ text: "Strategy & Product Week", bold: true, size: 48, font: "Arial", color: C.orange })] }),

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
      row("Andres (CEO) & Paola",    "Full week — Wed Apr 22 to Wed Apr 29", "Vision presentation, facilitation, all sessions"),
      row("Matt & Heath (Evolve)", "Full week — Wed Apr 22 to Wed Apr 29", "Operations, design partner alignment, working-session input"),
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
      kRow("DISCUSSION",   C.orangeDeep,C.orangeLight, "Open Discussion",   "Conversational session with a note-taker at the board. Time-boxed but collaborative. Used for strategy, pricing, GTM."),
      kRow("WORKING",      C.white,     C.orange,      "Working Session",   "Full-day collaborative working session. One board, one note-taker, everyone contributes. Used for product, systems, and roadmap days."),
      kRow("FIXED",        C.fixedText, C.fixedBg,     "Fixed Block",       "Pre-confirmed meeting — time and participants locked."),
    ]
  });
}

// ─── DAY HEADER ───────────────────────────────────────────────────────────
function dayHdr(num, date, theme, type = "normal") {
  // type: "normal" | "lite" | "sprint"  (sprint now means full working-session day)
  const isSprint = type === "sprint";
  const isLite   = type === "lite";
  const bgColor  = isSprint ? C.orange : C.black;
  const txtColor = C.white;
  const label    = isSprint ? "  \u00B7  WORKING SESSION" : (isLite ? "  \u00B7  DISCUSSION" : "");
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
    "360 Sierra tells its story. Open sessions on strategy, market positioning, and investment framework.",
    "Andres, Paola, Matt, Heath, Sam  \u00B7  David & Nick join from afternoon"
  ),
  tbl([
    blk("10:00 AM", "Welcome & Week Overview",       "Opening: format, goals, 8-day agenda, and how the working sessions will run. Set expectations for how decisions will be made across the week."),
    blk("10:15 AM", "360 Sierra \u2014 Story, Vision & Product Portfolio", "Origin story, what we are building, why now, and the path to 10,000 and 50,000 vehicles. Walk through the full product architecture: 360 Sierra (company), RentalBuddy (operator management), Shakkii (AI operations), and Lemonade (booking platform). Align on naming and how each product tells part of the overall story.", "presentation"),
    blk("11:30 AM", "Break",                         "", "break"),
    blk("11:45 AM", "Strategy, Positioning & Competitive Differentiation", "3\u20135 year direction and success metrics (ARR, market share). Competitive landscape and 360 Sierra\u2019s unfair advantage \u2014 what makes us hard to copy.", "hi"),
    blk("1:00 PM",  "Lunch",                         "", "break"),
    blk("2:00 PM",  "Investment Framework",           "SAFE structure, revenue share terms. Capital deployment: where $50K goes, what it unlocks. $50K vs. $200K scenario. David joins.", "hi"),
    blk("3:00 PM",  "Seed Milestones & Exit Path",   "3\u20134 milestones that unlock the Seed round. Long-term exit considerations and valuation trajectory. What does 5-year success look like?", "hi"),
    blk("3:30 PM",  "Break",                         "", "break"),
    blk("3:45 PM",  "Open Items & Flex",             "Buffer for any open topics, investor questions, or items we did not get to earlier. If closed early, we end early."),
    blk("4:30 PM",  "Day 1 Recap",                   "Capture decisions and open items. David and Nick depart for the evening. Preview tomorrow."),
    blk("5:00 PM",  "End of Day",                    ""),
  ]),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 2 — Thursday April 23
  // ══════════════════════════════════════════════
  dayHdr(2, "Thursday, April 23", "Design Partner, GTM & Investor Alignment", "lite"),
  ...metaLines(
    "Close the Evolve design partner role, define ICP, pricing and GTM. Afternoon: 1:1 with David, open investor Q&A, and a social gathering with Wildwood to close the day.",
    "Andres, Paola, Matt, Heath, Sam  \u00B7  Nick joins for GTM  \u00B7  David joins afternoon (1:1 and Q&A)  \u00B7  Wildwood joins 4:00 PM (social)"
  ),
  tbl([
    blk("10:00 AM", "Recap & Open Items",              "Resolve any open loops from Day 1 before today\u2019s sessions."),
    blk("10:15 AM", "Evolve as Design Partner",        "Define the role formally: what Evolve contributes, what they receive, how product decisions get made together. Close all open loops before the working sessions begin.", "hi"),
    blk("11:30 AM", "Break",                           "", "break"),
    blk("11:45 AM", "ICP & Pricing Strategy",          "Define the ideal customer profile (operator size, segment, geography) and the pricing logic that fits \u2014 packaging, tiers, contract length.", "hi"),
    blk("12:20 PM", "Go-to-Market Strategy",           "Define together the commercial model, sales motion, channel strategy, and partnership model. Work through the path to 10K vehicles. Nick joins.", "hi"),
    blk("1:00 PM",  "Lunch",                           "", "break"),
    blk("2:00 PM",  "1:1 \u2014 Evolve \u00D7 David",             "Evolve team (Matt, Heath, Sam) and David. Private session. Andres and Paola are not present.", "fixed"),
    blk("3:00 PM",  "Risk, Open Q&A & De-risking",     "Open floor for remaining investor questions. Address risk factors, execution confidence, and what de-risks the bet.", "hi"),
    blk("4:00 PM",  "Wildwood \u00D7 Evolve \u00D7 360 Sierra \u2014 Social", "Drinks & dinner.", "hi"),
    blk("5:30 PM",  "End of Day",                      ""),
  ]),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 3 — Friday April 24 — LONG-TERM GOAL, SYSTEM MAP & BOOKING CORE
  // ══════════════════════════════════════════════
  dayHdr(3, "Friday, April 24", "Long-Term Goal, System Map & Booking Core", "sprint"),
  ...metaLines(
    "Set the three-year vision. Draw the full operator and customer flows. Move into the transactional core: booking rules, utilization, vehicle management, and OTA integrations.",
    "Andres, Paola, Matt, Heath, Sam  \u00B7  1 Evolve ops team member (in person)"
  ),
  tbl([
    blk("10:00 AM", "Kickoff & Ground Rules",         "Open the week as a team: how we will work, how decisions get made, how we capture notes. No solo work today \u2014 everything happens on the board.", "sprint-hi"),
    blk("10:15 AM", "Long-Term Goal",                 "As a group: where does 360 Sierra need to be in three years? The product, the operator base, the brand. Written on the board. Anchors every decision that follows.", "sprint-hi"),
    blk("11:00 AM", "What Are We Afraid Of?",         "Turn risks into testable questions: \u201CCan we achieve X by end of the rollout?\u201D These become the questions we answer across the week.", "sprint-hi"),
    blk("11:30 AM", "Break",                          "", "break"),
    blk("11:45 AM", "System Map \u2014 Operator Flow",      "Draw the full operator journey: onboarding \u2192 fleet setup \u2192 rate and category config \u2192 bookings \u2192 vehicle handover \u2192 returns \u2192 reporting. Every actor, handoff, and integration. Red-flag friction.", "sprint-hi"),
    blk("1:00 PM",  "Lunch",                          "", "break"),
    blk("2:00 PM",  "System Map \u2014 Customer Flow",      "Customer side of the map: search \u2192 quote \u2192 booking \u2192 communications \u2192 pickup \u2192 rental experience \u2192 return \u2192 post-trip. Where does it break today? What matters most?", "sprint-hi"),
    blk("2:45 PM",  "Booking Rules, Types & Utilization", "Re-calculation rules, change logs, extras (optional, mandatory, collision damage waiver, fixed/daily/percent fees). Signed rental agreement and what happens if a booking changes after signing. Booking types: revenue, maintenance, non-revenue. Utilization: automatic shuffling before pickup to maximize utilization, turnaround hours, inter-location sharing, controlled overbooking.", "sprint-hi"),
    blk("3:30 PM",  "Break",                          "", "break"),
    blk("3:45 PM",  "Vehicle Management",             "Preventive and manual servicing, damages, vehicle registration, road-user charges, GPS, activity logs. Category changes on vehicles that do not break historical reports. Mid-rental vehicle swaps for fines and tolls.", "sprint-hi"),
    blk("4:30 PM",  "OTA & Channel Integrations",     "OTA connectivity \u2014 Booking.com, Turo, Outdoorsy and other aggregators. Inbound bookings captured in real time, rate parity and availability pushed out, channel-manager architecture, channel-level commissions, customer data ownership, and how cancellations and modifications sync back.", "sprint-hi"),
    blk("5:00 PM",  "End of Day",                     "", "sprint-hi"),
  ], true),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 4 — Saturday April 25 — FOUNDATIONS & RENTALBUDDY ROADMAP (half-day)
  // ══════════════════════════════════════════════
  dayHdr(4, "Saturday, April 25", "Foundations & RentalBuddy Roadmap", "sprint"),
  ...metaLines(
    "Close out the transactional layer: access, security, payments, compliance, tenancy, rates, and fraud. Lock the RentalBuddy roadmap before Sam departs.",
    "Andres, Paola, Matt, Heath, Sam (morning only)"
  ),
  tbl([
    blk("10:00 AM", "Recap & Focus",                  "Reconnect with Friday\u2019s booking-core decisions. Frame Saturday: foundations (access, payments, tenancy, rates, fraud) and the RentalBuddy roadmap.", "sprint-hi"),
    blk("10:10 AM", "User Access, Security & Payments", "MFA required across the board, no shared users. Role- and location-based access levels. Logging and auditing of sensitive actions. PCI scope, payment gateway integrations. Who can issue refunds and modify payments, and under what controls.", "sprint-hi"),
    blk("10:50 AM", "System Setup, Tenancy & Compliance", "Single- vs multi-country. Tax rules: inclusive/exclusive, state and country level, tax-exempt bookings. Sensitive data retention and GDPR right-to-erasure flows. Franchise and multi-tenant branding \u2014 what is shared, what is isolated.", "sprint-hi"),
    blk("11:15 AM", "Break",                          "", "break"),
    blk("11:30 AM", "Rates Engine & Fraud Controls",   "Rental period calculation: 24-hour, calendar day, hourly, part-day. Rate types: Retail, Corporate, Agent, Long-term/subscription. Seasons, locations, categories. In-system vs 3rd-party rate aggregator. When rates re-calculate. Do-Not-Rent list, blacklists, risk checks, manual overrides.", "sprint-hi"),
    blk("12:15 PM", "Roadmap \u2014 RentalBuddy / Booking Core", "Lock the RentalBuddy roadmap. Sequence the transactional core: onboarding, booking engine, rates, payments, real-time fleet data, OTA and channel integrations, agent portal, check-in/return flow. MVP, v1.1, post-Seed. Dependencies and ownership.", "sprint-hi"),
    blk("1:00 PM",  "Close & Sam Farewell",           "Quick wrap-up of Saturday\u2019s foundations and the RentalBuddy roadmap. Sam departs.", "hi"),
  ], true),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 5 — Sunday April 26
  // ══════════════════════════════════════════════
  dayHdr(5, "Sunday, April 26", "Rest Day"),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 100 },
    children: [new TextRun({ text: "\u2605  No scheduled activities  \u2605", size: 22, font: "Arial", color: C.midGray })] }),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 6 — Monday April 27 — AI LAYER (FULL DAY)
  // ══════════════════════════════════════════════
  dayHdr(6, "Monday, April 27", "AI Layer \u2014 Modules, Orchestration & Roadmap", "sprint"),
  ...metaLines(
    "Full day on the AI layer. Revisit both system maps, decide where AI creates real leverage, work through MCP servers and agent orchestration, build vs. integrate, and lock the AI roadmap.",
    "Andres, Paola, Matt, Heath"
  ),
  tbl([
    blk("10:00 AM", "Recap & Focus",                  "Reconnect with Fri/Sat. Frame the AI day: where AI lives across the flow, how it is orchestrated, what we build vs. integrate, and the roadmap.", "sprint-hi"),
    blk("10:15 AM", "AI Across the Flow \u2014 Revisit the Maps", "Walk both system maps again and mark every touchpoint where AI meaningfully changes the experience or the unit economics. Customer-facing moments, operator flows, and fleet intelligence. What is truly high-leverage vs. nice-to-have.", "sprint-hi"),
    blk("11:15 AM", "Break",                          "", "break"),
    blk("11:30 AM", "AI \u2014 Customer-Facing Modules",     "Search, quote, booking assistance, journey rescue, support, upsell. How AI shows up to the customer without feeling artificial. Tone, fallbacks, human handoff, privacy boundaries.", "sprint-hi"),
    blk("12:30 PM", "AI \u2014 Operations & Fleet Intelligence", "Dispatch, pricing, utilization, predictive maintenance, anomaly detection, demand forecasting. Which ops to automate first (highest impact, lowest complexity).", "sprint-hi"),
    blk("1:00 PM",  "Lunch",                          "", "break"),
    blk("2:00 PM",  "AI \u2014 Operations & Fleet Intelligence (cont.)", "Deep-dive continued: operational agents, reporting, rebalancing logic, the data AI needs to work.", "sprint-hi"),
    blk("2:30 PM",  "MCP Servers, Agents & Orchestration", "How MCP servers and agents orchestrate the stack end-to-end. Data contracts, tool-use boundaries, memory and state, how agents read/write across RentalBuddy, Communications, and Lemonade.", "sprint-hi"),
    blk("3:30 PM",  "Break",                          "", "break"),
    blk("3:45 PM",  "Build vs. Integrate \u2014 Per Module", "For each AI module: do we build, wrap an existing provider, or integrate and own the orchestration? Speed-to-market vs. defensibility. Cost ceilings.", "sprint-hi"),
    blk("4:15 PM",  "Roadmap \u2014 AI Layer",                "Lock the AI roadmap. Which modules ship in MVP, which in v1.1, which post-Seed. Sequence and dependencies on RentalBuddy and the Communications layer.", "sprint-hi"),
    blk("5:00 PM",  "End of Day",                     "", "sprint-hi"),
  ], true),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 7 — Tuesday April 28 — COMMUNICATIONS PILLAR (FULL DAY)
  // ══════════════════════════════════════════════
  dayHdr(7, "Tuesday, April 28", "Communications Pillar", "sprint"),
  ...metaLines(
    "Full day on the Communications pillar \u2014 customer, internal, third-party, and proactive monitoring. Close with the Communications roadmap.",
    "Andres, Paola, Matt, Heath"
  ),
  tbl([
    blk("10:00 AM", "Recap & Focus",                  "Frame the day: customer comms, internal, partner, and proactive monitoring.", "sprint-hi"),
    blk("10:15 AM", "Customer Communications \u2014 Inbound", "Emails and messages coming in from partners and customers. Captured, classified, and routed so nothing is lost.", "sprint-hi"),
    blk("11:00 AM", "Customer Communications \u2014 Outbound", "Messages we send out: transactional, marketing, service-recovery. Tone, timing, channel, fallbacks.", "sprint-hi"),
    blk("11:30 AM", "Break",                          "", "break"),
    blk("11:45 AM", "Customer Communications \u2014 CRM & Unified Journey", "One inbox and one history per customer across email, call, SMS, and WhatsApp. Visible to everyone who needs it.", "sprint-hi"),
    blk("1:00 PM",  "Lunch",                          "", "break"),
    blk("2:00 PM",  "Internal Staff Communications",  "One platform for staff follow-ups, handoffs, and alerts. Threads tied to the booking, vehicle, or customer \u2014 not to a chat group.", "sprint-hi"),
    blk("2:45 PM",  "Third-Party Network & Repairer Portals", "Portals for repairers and service partners. They see what they need, nothing more. All partner comms in one place.", "sprint-hi"),
    blk("3:15 PM",  "Break",                          "", "break"),
    blk("3:30 PM",  "Proactive Monitoring \u2014 Errors & Missed Opportunities", "When something breaks, the system flags it and offers a fix before it escalates. When a search returns no availability, we log the missed demand.", "sprint-hi"),
    blk("4:15 PM",  "Roadmap \u2014 Communications Layer",    "Lock the Communications roadmap. What ships in MVP, v1.1, and post-Seed. Dependencies on RentalBuddy and AI.", "sprint-hi"),
    blk("5:00 PM",  "End of Day",                     "", "sprint-hi"),
  ], true),
  sp(),

  // ══════════════════════════════════════════════
  // DAY 8 — Wednesday April 29 — WRAP-UP, INVESTMENT & DEPARTURES
  // ══════════════════════════════════════════════
  dayHdr(8, "Wednesday, April 29", "Wrap-Up, Investment & Departures"),
  ...metaLines(
    "Short morning to consolidate the week, lock action items, confirm investment next steps, and close. Travel in the afternoon.",
    "Andres, Paola, Matt, Heath"
  ),
  tbl([
    blk("10:00 AM", "Week Recap & Lock",              "Walk through every decision made across the week: maps, AI roadmap, RentalBuddy roadmap, Communications roadmap, Foundations. Confirm nothing is open.", "hi"),
    blk("10:45 AM", "Action Items & Working Cadence",  "Every action item with owner and deadline between 360 Sierra and Evolve. Define the working format from here on: how we meet, how we share progress, how we raise blockers, and the reporting cadence.", "hi"),
    blk("11:15 AM", "Investment \u2014 Next Steps",          "Transfer timeline, SAFE execution steps, reporting structure.", "hi"),
    blk("11:45 AM", "Closing Remarks & Farewell",     "Acknowledge the work done. Communication cadence going forward.", "hi"),
    blk("12:00 PM", "End of Week",                    ""),
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
