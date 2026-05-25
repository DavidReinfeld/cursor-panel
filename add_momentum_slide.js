// add_momentum_slide.js
// Run with: node add_momentum_slide.js
// Requires: npm install pptxgenjs
// Adds the "90-Day Momentum Plan" slide to the existing Apple/Cursor AE presentation.
// Drop this file in the same directory as your presentation script and call addMomentumSlide(pres)
// before pres.writeFile(), or run standalone to generate a single-slide test file.

const PptxGenJS = require("pptxgenjs");

// ─── Style constants (match existing Apple presentation) ─────────────────────
const FONT = "Helvetica Neue";
const BG = "FFFFFF";
const CARD_FILL = "F5F5F7";
const TEXT_PRIMARY = "1D1D1F";
const TEXT_SECONDARY = "6E6E73";
const BORDER_COLOR = "D2D2D7";
const BORDER_WIDTH = 0.5;

// ─── Layout ──────────────────────────────────────────────────────────────────
// Slide: 10" × 5.625" (LAYOUT_16x9)
// Three equal columns side by side (horizontal layout)
const COL_W = 3.07;
const COL_GAP = 0.095;
const COL_X = [0.28, 0.28 + COL_W + COL_GAP, 0.28 + (COL_W + COL_GAP) * 2];
const CARD_H = 0.6;
const CARD_GAP = 0.07;
const CARDS_START_Y = 1.38;
const GOAL_Y = 4.95;

// ─── Phase data ──────────────────────────────────────────────────────────────
const phases = [
  {
    badgeFill: "EEEDFE",
    badgeText: "3C3489",
    badgeLabel: "Days 1–30  ·  Foundation",
    title: "Map the account, plant the flags",
    sub: "Intelligence gathering + first signals before/during WWDC (June 9)",
    goalLabel: "30-DAY GOAL",
    goal: "10+ booked discovery calls; 3+ follow-up conversations from WWDC events",
    motions: [
      {
        label: "Free user audit",
        detail: "Pull Cursor usage data for @apple.com emails. Warm outreach to active free users first.",
      },
      {
        label: "Community seeding",
        detail: "Swift Forums, r/swift, Mastodon/X dev circles. Share real Cursor + Xcode workflows; spark discourse.",
      },
      {
        label: "Cold outbound",
        detail: "Xcode Intelligence team leads + Siri/AI eng managers. Hook: Claude Agent in Xcode 26.3.",
      },
      {
        label: "WWDC activation",
        detail: "Fringe events, Apple dev labs, side dinners. Prospect Swift/Xcode eng in person.",
      },
    ],
  },
  {
    badgeFill: "E1F5EE",
    badgeText: "085041",
    badgeLabel: "Days 31–60  ·  Build presence",
    title: "Creating momentum through live events",
    sub: "Events, content, and experiences that pull Apple eng toward Cursor",
    goalLabel: "60-DAY GOAL",
    goal: "2+ hosted events with Apple eng attendance; champion tested and identified in ≥1 Tier 1 team; interest in active evaluation",
    motions: [
      {
        label: "Swift/Xcode webinar",
        detail: "Host a technical session on agentic coding with Cursor + Xcode 26.3. Target Apple dev community.",
      },
      {
        label: "Workshop series",
        detail: "Hands-on virtual sessions for Apple eng teams. Focused on real workflows, not product pitches.",
      },
      {
        label: "In-person events",
        detail: "Dinners, meetups, side events in Cupertino/SF. Build relationships outside the sales context.",
      },
      {
        label: "Proof-point content",
        detail: "Demos and write-ups on Cursor + Xcode agentic flow. Built to circulate in Apple dev circles.",
      },
    ],
  },
  {
    badgeFill: "FAEEDA",
    badgeText: "633806",
    badgeLabel: "Days 61–90  ·  Conversion",
    title: "Drive to pilot or commercial motion",
    sub: "Leverage proof points and momentum to engage senior leadership and close",
    goalLabel: "90-DAY GOAL",
    goal: "≥1 active pilot or commercial evaluation; exec sponsor engaged; multiple threads active across Tier 1",
    motions: [
      {
        label: "Pilot proposal",
        detail: "Structured pilot for Tier 1 team (internal eng or Siri rebuild). Time-boxed, clear success metrics.",
      },
      {
        label: "Executive alignment",
        detail: "VP/director-level intro via champion. Framed against Ternus HW/SW convergence strategy.",
      },
      {
        label: "Senior exec network",
        detail: "Use pilot proof points to engage network into Apple C-suite and senior leadership. Warm, not cold.",
      },
      {
        label: "Multi-thread the org",
        detail: "Expand contacts across Tier 1 teams. Don't let the deal ride on a single champion or thread.",
      },
    ],
  },
];

// ─── Main function ────────────────────────────────────────────────────────────
function addMomentumSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: BG };

  // Slide title
  slide.addText("90-Day Momentum Plan", {
    x: 0.28, y: 0.12, w: 9.44, h: 0.32,
    fontFace: FONT,
    fontSize: 13,
    bold: true,
    color: TEXT_PRIMARY,
    margin: 0,
  });

  // Thin divider under title
  slide.addShape(pres.ShapeType.rect, {
    x: 0.28, y: 0.46, w: 9.44, h: 0.01,
    fill: { color: BORDER_COLOR },
    line: { color: BORDER_COLOR, width: 0 },
  });

  phases.forEach((phase, pi) => {
    const x = COL_X[pi];

    // ── Badge ──
    slide.addShape(pres.ShapeType.roundRect, {
      x, y: 0.55, w: 2.2, h: 0.21,
      fill: { color: phase.badgeFill },
      line: { color: phase.badgeFill, width: 0 },
      rectRadius: 0.05,
    });
    slide.addText(phase.badgeLabel, {
      x, y: 0.55, w: 2.2, h: 0.21,
      fontFace: FONT,
      fontSize: 7.5,
      bold: true,
      color: phase.badgeText,
      align: "center",
      valign: "middle",
      margin: 0,
    });

    // ── Phase title ──
    slide.addText(phase.title, {
      x, y: 0.83, w: COL_W, h: 0.22,
      fontFace: FONT,
      fontSize: 9,
      bold: true,
      color: TEXT_PRIMARY,
      margin: 0,
    });

    // ── Phase subtitle ──
    slide.addText(phase.sub, {
      x, y: 1.05, w: COL_W, h: 0.28,
      fontFace: FONT,
      fontSize: 7.5,
      color: TEXT_SECONDARY,
      margin: 0,
    });

    // ── Motion cards (4 per column, stacked vertically) ──
    phase.motions.forEach((m, mi) => {
      const cardY = CARDS_START_Y + mi * (CARD_H + CARD_GAP);

      // Card background
      slide.addShape(pres.ShapeType.rect, {
        x, y: cardY, w: COL_W, h: CARD_H,
        fill: { color: CARD_FILL },
        line: { color: BORDER_COLOR, width: BORDER_WIDTH },
      });

      // Motion label
      slide.addText(m.label, {
        x: x + 0.1, y: cardY + 0.07, w: COL_W - 0.15, h: 0.17,
        fontFace: FONT,
        fontSize: 8.5,
        bold: true,
        color: TEXT_PRIMARY,
        margin: 0,
      });

      // Motion detail
      slide.addText(m.detail, {
        x: x + 0.1, y: cardY + 0.25, w: COL_W - 0.15, h: 0.3,
        fontFace: FONT,
        fontSize: 7.5,
        color: TEXT_SECONDARY,
        margin: 0,
      });
    });

    // ── Goal bar ──
    slide.addShape(pres.ShapeType.rect, {
      x, y: GOAL_Y, w: COL_W, h: 0.54,
      fill: { color: CARD_FILL },
      line: { color: BORDER_COLOR, width: BORDER_WIDTH },
    });

    slide.addText(phase.goalLabel, {
      x: x + 0.1, y: GOAL_Y + 0.06, w: COL_W - 0.15, h: 0.14,
      fontFace: FONT,
      fontSize: 6.5,
      bold: true,
      color: TEXT_SECONDARY,
      charSpacing: 1,
      margin: 0,
    });

    slide.addText(phase.goal, {
      x: x + 0.1, y: GOAL_Y + 0.2, w: COL_W - 0.15, h: 0.3,
      fontFace: FONT,
      fontSize: 7.5,
      bold: true,
      color: TEXT_PRIMARY,
      margin: 0,
    });
  });

  return slide;
}

// ─── Standalone: generate a test file ────────────────────────────────────────
// Remove this block when integrating into the main presentation script
if (require.main === module) {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  addMomentumSlide(pres);
  pres.writeFile({ fileName: "momentum_slide_test.pptx" })
    .then(() => console.log("✅  momentum_slide_test.pptx written"))
    .catch((e) => console.error("❌  Error:", e));
}

module.exports = { addMomentumSlide };
