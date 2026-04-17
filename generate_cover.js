const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icon rendering helpers
function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

async function main() {
  const { FaCloud, FaRobot, FaCogs, FaChartLine, FaNetworkWired } = require("react-icons/fa");


  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "TigerAI 虎智科技";
  pres.title = "提升商業服務業營運效能強化韌性平台";

  const slide = pres.addSlide();

  // === DARK NAVY BACKGROUND ===
  slide.background = { color: "0D1B2A" };

  // === DECORATIVE GEOMETRIC SHAPES (abstract tech feel) ===

  // Large circle (top-right, subtle)
  slide.addShape(pres.shapes.OVAL, {
    x: 7.5, y: -1.2, w: 4, h: 4,
    fill: { color: "1B3A5C", transparency: 60 },
    line: { color: "00A896", width: 1.5, transparency: 40 }
  });

  // Medium circle (bottom-left)
  slide.addShape(pres.shapes.OVAL, {
    x: -0.8, y: 3.5, w: 3, h: 3,
    fill: { color: "1B3A5C", transparency: 70 },
    line: { color: "028090", width: 1, transparency: 50 }
  });

  // Small circle (mid-right)
  slide.addShape(pres.shapes.OVAL, {
    x: 8.8, y: 2.8, w: 1.5, h: 1.5,
    fill: { color: "028090", transparency: 75 },
    line: { color: "00A896", width: 1, transparency: 30 }
  });

  // Tiny accent circle
  slide.addShape(pres.shapes.OVAL, {
    x: 1.2, y: 0.8, w: 0.6, h: 0.6,
    fill: { color: "00A896", transparency: 60 }
  });

  // Another tiny accent
  slide.addShape(pres.shapes.OVAL, {
    x: 6.5, y: 4.2, w: 0.4, h: 0.4,
    fill: { color: "02C39A", transparency: 50 }
  });

  // === HORIZONTAL ACCENT LINE ===
  slide.addShape(pres.shapes.LINE, {
    x: 0.8, y: 2.65, w: 2.2, h: 0,
    line: { color: "00A896", width: 3 }
  });

  // === TEAL LEFT ACCENT BAR ===
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.08, h: 5.625,
    fill: { color: "00A896" }
  });

  // === BOTTOM GRADIENT BAR ===
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.25, w: 10, h: 0.375,
    fill: { color: "0A1628" }
  });

  // === MAIN TITLE - Chinese ===
  slide.addText("提升商業服務業營運", {
    x: 0.8, y: 1.0, w: 7, h: 0.7,
    fontSize: 36, fontFace: "Microsoft JhengHei",
    color: "FFFFFF", bold: true, margin: 0
  });

  slide.addText("效能強化韌性平台", {
    x: 0.8, y: 1.65, w: 7, h: 0.7,
    fontSize: 36, fontFace: "Microsoft JhengHei",
    color: "FFFFFF", bold: true, margin: 0
  });

  // === ACCENT LINE UNDER TITLE ===
  // (using the horizontal line already placed at y: 2.65)

  // === SUBTITLE ===
  slide.addText("以低門檻、模組化方案，啟動您的 AI 轉型第一步", {
    x: 0.8, y: 2.85, w: 7, h: 0.5,
    fontSize: 16, fontFace: "Microsoft JhengHei",
    color: "8ECAE6", margin: 0
  });

  // === SOLUTION LABELS (pill-style boxes) ===
  // Pure Cloud label
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 3.55, w: 3.6, h: 0.75,
    fill: { color: "028090", transparency: 30 },
    line: { color: "00A896", width: 1 },
    rectRadius: 0.08
  });
  slide.addText([
    { text: "純雲端版", options: { bold: true, fontSize: 14, color: "00E5CC" } },
    { text: "  自動智能雲端管家 · AI Agent", options: { fontSize: 11, color: "B0D4D4" } }
  ], {
    x: 0.8, y: 3.55, w: 3.6, h: 0.75,
    fontFace: "Microsoft JhengHei",
    align: "center", valign: "middle", margin: 0
  });

  // Hybrid label
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 4.6, y: 3.55, w: 3.6, h: 0.75,
    fill: { color: "1B3A5C", transparency: 30 },
    line: { color: "028090", width: 1 },
    rectRadius: 0.08
  });
  slide.addText([
    { text: "雲地混和版", options: { bold: true, fontSize: 14, color: "00E5CC" } },
    { text: "  TigerAI OpenGenie", options: { fontSize: 11, color: "B0D4D4" } }
  ], {
    x: 4.6, y: 3.55, w: 3.6, h: 0.75,
    fontFace: "Microsoft JhengHei",
    align: "center", valign: "middle", margin: 0
  });

  // === ICONS ROW (bottom area) ===
  const iconColor = "#00A896";
  const [cloudIcon, robotIcon, cogsIcon, chartIcon, networkIcon] = await Promise.all([
    iconToBase64Png(FaCloud, iconColor, 256),
    iconToBase64Png(FaRobot, iconColor, 256),
    iconToBase64Png(FaCogs, iconColor, 256),
    iconToBase64Png(FaChartLine, iconColor, 256),
    iconToBase64Png(FaNetworkWired, iconColor, 256),
  ]);

  const iconY = 4.35;
  const iconSize = 0.35;
  const icons = [cloudIcon, robotIcon, cogsIcon, chartIcon, networkIcon];
  const iconLabels = ["雲端部署", "AI Agent", "自動化", "數據分析", "系統整合"];
  const startX = 0.8;
  const gap = 1.1;

  icons.forEach((icon, i) => {
    const x = startX + i * gap;
    // Icon circle background
    slide.addShape(pres.shapes.OVAL, {
      x: x, y: iconY, w: iconSize + 0.12, h: iconSize + 0.12,
      fill: { color: "0D2B3E", transparency: 20 },
      line: { color: "028090", width: 0.5, transparency: 40 }
    });
    // Icon
    slide.addImage({
      data: icon, x: x + 0.06, y: iconY + 0.06,
      w: iconSize, h: iconSize
    });
    // Label
    slide.addText(iconLabels[i], {
      x: x - 0.2, y: iconY + iconSize + 0.12, w: iconSize + 0.52, h: 0.25,
      fontSize: 8, fontFace: "Microsoft JhengHei",
      color: "6B9DAD", align: "center", margin: 0
    });
  });

  // === COMPANY INFO (bottom-right) ===
  slide.addText("虎智科技 TigerAI", {
    x: 6.8, y: 4.5, w: 2.7, h: 0.35,
    fontSize: 14, fontFace: "Microsoft JhengHei",
    color: "00A896", bold: true, align: "right", margin: 0
  });
  slide.addText("AI 解決方案", {
    x: 6.8, y: 4.8, w: 2.7, h: 0.3,
    fontSize: 10, fontFace: "Microsoft JhengHei",
    color: "5A8A9A", align: "right", margin: 0
  });

  // === SAVE ===
  const outputPath = "e:/Claude code/My Project/TigerAI/PPT製作/資料放置區/cover_slide.pptx";
  await pres.writeFile({ fileName: outputPath });
  console.log("Cover slide saved to:", outputPath);
}

main().catch(console.error);
