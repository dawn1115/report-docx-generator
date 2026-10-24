const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, PageNumber, NumberFormat, SectionType,
  AlignmentType, HeadingLevel, WidthType, BorderStyle, ShadingType,
  TableLayoutType, LevelFormat,
} = require("docx");
const fs = require("fs");

// ───────────────────────── palette (DS-1 Deep Sea) ─────────────────────────
const P = {
  bg: "0B1C2C", accent: "529286",
  titleColor: "FFFFFF", subtitleColor: "B0B8C0", metaColor: "90989F", footerColor: "687078",
  headingColor: "0B1C2C",
  table: { headerBg: "529286", headerText: "FFFFFF", accentLine: "529286", innerLine: "BECFCC", surface: "E8ECEB" },
};

const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB };
const allNoBorders = { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB };

// ───────────────────────── cover helpers (design-system) ─────────────────────────
function splitTitleLines(title, charsPerLine) {
  if (title.length <= charsPerLine) return [title];
  const breakAfter = new Set([..."\uff0c\u3002\u3001\uff1b\uff1a\uff01\uff1f", ..."\u7684\u4e0e\u548c\u53ca\u4e4b\u5728\u4e8e\u4e3a", ..."-_\u2014\u2013\u00b7/", ..." \t"]);
  const lines = [];
  let remaining = title;
  while (remaining.length > charsPerLine) {
    let breakAt = -1;
    for (let i = charsPerLine; i >= Math.floor(charsPerLine * 0.6); i--) {
      if (i < remaining.length && breakAfter.has(remaining[i - 1])) { breakAt = i; break; }
    }
    if (breakAt === -1) {
      const limit = Math.min(remaining.length, Math.ceil(charsPerLine * 1.3));
      for (let i = charsPerLine + 1; i < limit; i++) {
        if (breakAfter.has(remaining[i - 1])) { breakAt = i; break; }
      }
    }
    if (breakAt === -1) {
      breakAt = charsPerLine;
      const prevChar = remaining[breakAt - 1], nextChar = remaining[breakAt];
      if (prevChar && nextChar && !breakAfter.has(prevChar) && !breakAfter.has(nextChar) &&
          /[\u4e00-\u9fff]/.test(prevChar) && /[\u4e00-\u9fff]/.test(nextChar)) breakAt = breakAt - 1;
    }
    lines.push(remaining.slice(0, breakAt).trim());
    remaining = remaining.slice(breakAt).trim();
  }
  if (remaining) lines.push(remaining);
  if (lines.length > 1 && lines[lines.length - 1].length <= 2) {
    const last = lines.pop();
    lines[lines.length - 1] += last;
  }
  return lines;
}

function calcTitleLayout(title, maxWidthTwips, preferredPt = 40, minPt = 24) {
  const charWidth = (pt) => pt * 20;
  const charsPerLine = (pt) => Math.floor(maxWidthTwips / charWidth(pt));
  let titlePt = preferredPt, lines;
  while (titlePt >= minPt) {
    const cpl = charsPerLine(titlePt);
    if (cpl < 2) { titlePt -= 2; continue; }
    lines = splitTitleLines(title, cpl);
    if (lines.length <= 3) break;
    titlePt -= 2;
  }
  if (!lines || lines.length > 3) {
    lines = splitTitleLines(title, charsPerLine(minPt));
    titlePt = minPt;
  }
  return { titlePt, titleLines: lines };
}

function calcCoverSpacing(params) {
  const {
    titleLineCount = 1, titlePt = 36, hasSubtitle = false, hasEnglishLabel = false,
    metaLineCount = 0, fixedHeight = 800, pageHeight = 16838, marginTop = 0, marginBottom = 0,
  } = params;
  const SAFETY = 1200;
  const usableHeight = pageHeight - marginTop - marginBottom - SAFETY;
  const titleHeight = titleLineCount * (titlePt * 23 + 200);
  const subtitleHeight = hasSubtitle ? (12 * 23 + 600) : 0;
  const englishLabelHeight = hasEnglishLabel ? (9 * 23 + 600) : 0;
  const metaHeight = metaLineCount * (10 * 23 + 100);
  const implicitParaHeight = 3 * 300;
  const contentHeight = titleHeight + subtitleHeight + englishLabelHeight + metaHeight + fixedHeight + implicitParaHeight;
  const safeRemaining = Math.max(usableHeight - contentHeight, 400);
  const FOOTER_MIN = 800;
  const rawTop = Math.floor(safeRemaining * 0.45);
  const rawBottom = Math.floor(safeRemaining * 0.45);
  const bottomSpacing = Math.max(rawBottom, FOOTER_MIN);
  const topSpacing = Math.max(rawTop - Math.max(0, FOOTER_MIN - rawBottom), 400);
  const midSpacing = Math.max(safeRemaining - topSpacing - bottomSpacing, 0);
  return { topSpacing, midSpacing, bottomSpacing };
}

function buildCoverR1(config) {
  const C = config.palette;
  const padL = 1200, padR = 800;
  const availableWidth = 11906 - padL - padR - 300;
  const { titlePt, titleLines } = calcTitleLayout(config.title, availableWidth, 40, 24);
  const titleSize = titlePt * 2;
  const spacing = calcCoverSpacing({
    titleLineCount: titleLines.length, titlePt,
    hasSubtitle: !!config.subtitle, hasEnglishLabel: !!config.englishLabel,
    metaLineCount: (config.metaLines || []).length, fixedHeight: 400,
  });
  const accentLeft = { style: BorderStyle.SINGLE, size: 8, color: C.accent, space: 12 };
  const children = [];
  children.push(new Paragraph({ spacing: { before: spacing.topSpacing } }));
  if (config.englishLabel) {
    children.push(new Paragraph({
      indent: { left: padL, right: padR }, spacing: { after: 500 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.accent, space: 8 } },
      children: [new TextRun({ text: config.englishLabel.split("").join("  "),
        size: 18, color: C.accent, font: { ascii: "Calibri", eastAsia: "SimHei" }, characterSpacing: 40 })],
    }));
  }
  for (let i = 0; i < titleLines.length; i++) {
    children.push(new Paragraph({
      indent: { left: padL },
      spacing: { after: i < titleLines.length - 1 ? 100 : 300, line: Math.ceil(titlePt * 23), lineRule: "atLeast" },
      children: [new TextRun({ text: titleLines[i], size: titleSize, bold: true,
        color: C.titleColor, font: { eastAsia: "SimHei", ascii: "Arial" } })],
    }));
  }
  if (config.subtitle) {
    children.push(new Paragraph({
      indent: { left: padL }, spacing: { after: 800 },
      children: [new TextRun({ text: config.subtitle, size: 24, color: C.subtitleColor,
        font: { eastAsia: "Microsoft YaHei", ascii: "Arial" } })],
    }));
  }
  for (const line of (config.metaLines || [])) {
    children.push(new Paragraph({
      indent: { left: padL + 200 }, spacing: { after: 80 },
      border: { left: accentLeft },
      children: [new TextRun({ text: line, size: 24, color: C.metaColor,
        font: { eastAsia: "Microsoft YaHei", ascii: "Arial" } })],
    }));
  }
  children.push(new Paragraph({ spacing: { before: spacing.bottomSpacing } }));
  children.push(new Paragraph({
    indent: { left: padL, right: padR },
    border: { top: { style: BorderStyle.SINGLE, size: 2, color: C.accent, space: 8 } },
    spacing: { before: 200 },
    children: [
      new TextRun({ text: config.footerLeft || "", size: 16, color: C.footerColor, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } }),
      new TextRun({ text: "                                        " }),
      new TextRun({ text: config.footerRight || "", size: 16, color: C.footerColor, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } }),
    ],
  }));
  return [new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: allNoBorders,
    rows: [new TableRow({
      height: { value: 16838, rule: "exact" },
      children: [new TableCell({
        shading: { type: ShadingType.CLEAR, fill: C.bg }, borders: noBorders,
        children,
      })],
    })],
  })];
}

// ───────────────────────── body helpers ─────────────────────────
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 180, line: 312 },
    children: [new TextRun({ text, bold: true, color: P.headingColor, font: { ascii: "Calibri", eastAsia: "SimHei" } })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120, line: 312 },
    children: [new TextRun({ text, bold: true, color: P.headingColor, font: { ascii: "Calibri", eastAsia: "SimHei" } })],
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 480 },
    spacing: { line: 312, after: 80 },
    children: [new TextRun({ text, size: 24, color: "000000", bold: !!opts.bold })],
  });
}
// paragraph with a bold lead phrase + normal rest
function bodyLead(lead, rest) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 480 },
    spacing: { line: 312, after: 80 },
    children: [
      new TextRun({ text: lead, size: 24, color: "000000", bold: true }),
      new TextRun({ text: rest, size: 24, color: "000000" }),
    ],
  });
}
function note(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 480 },
    spacing: { line: 312, before: 40, after: 160 },
    children: [new TextRun({ text, size: 21, color: "606060" })],
  });
}
function numItem(text, ref) {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 312, after: 60 },
    children: [new TextRun({ text, size: 24, color: "000000" })],
  });
}
function tableTitle(text) {
  return new Paragraph({
    keepNext: true, alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100, line: 312 },
    children: [new TextRun({ text, bold: true, size: 21, color: P.headingColor, font: { eastAsia: "SimHei", ascii: "Calibri" } })],
  });
}
function makeTable(headers, rows, widths, opts = {}) {
  const dataSize = opts.dataSize || 20;
  const headRow = new TableRow({
    tableHeader: true, cantSplit: true,
    children: headers.map((text, i) => new TableCell({
      children: [new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { line: 276 },
        children: [new TextRun({ text, bold: true, size: dataSize, color: P.table.headerText })],
      })],
      shading: { type: ShadingType.CLEAR, fill: P.table.headerBg },
      margins: { top: 70, bottom: 70, left: 100, right: 100 },
      width: { size: widths[i], type: WidthType.PERCENTAGE },
    })),
  });
  const dataRows = rows.map((row, ri) => new TableRow({
    cantSplit: true,
    children: row.map((text, i) => new TableCell({
      children: [new Paragraph({
        alignment: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
        spacing: { line: 276 },
        children: [new TextRun({ text: String(text), size: dataSize, color: "000000", bold: i === 0 && !!opts.boldFirstCol })],
      })],
      shading: opts.zebra && ri % 2 === 1 ? { type: ShadingType.CLEAR, fill: P.table.surface } : undefined,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      width: { size: widths[i], type: WidthType.PERCENTAGE },
    })),
  }));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: P.table.accentLine },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: P.table.accentLine },
      left: NB, right: NB,
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "D0D0D0" },
      insideVertical: NB,
    },
    rows: [headRow, ...dataRows],
  });
}
function pngSize(buf) { return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) }; }
function fig(path, displayWidth = 540) {
  const buf = fs.readFileSync(path);
  const dim = pngSize(buf);
  const h = Math.round(displayWidth * dim.height / dim.width);
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 160, after: 200 },
    children: [new ImageRun({ data: buf, transformation: { width: displayWidth, height: h }, type: "png" })],
  });
}

// ───────────────────────── content ─────────────────────────
const Q = (s) => "\u201c" + s + "\u201d"; // curly-quote helper

const children = [];

// ── 要点速览 ──
children.push(h1("要点速览"));
children.push(bodyLead("总量：", "2024年湛江、茂名、阳江三市GDP合计9541.55亿元，逼近万亿量级；茂名（4072.04亿元）反超湛江（3839.93亿元）居首，阳江（1629.58亿元）约为前两市的四成。"));
children.push(bodyLead("增速：", "2010\u20142024年三市GDP年均名义增速为7.5%（湛江）、7.5%（茂名）、6.9%（阳江），增长呈\u201c十二五\u201d高位（10.5%\u201411.5%）、\u201c十三五\u201d以来回落（4%\u20146%）、2020年后徘徊于5%左右的三段式换挡。"));
children.push(bodyLead("投资：", "三市固定资产投资已分别于2017、2021、2023年见顶回落，2024年湛江、阳江较峰值收缩逾三成，粤西整体进入投资消化期。"));
children.push(bodyLead("开放：", "外贸出口长期徘徊，2024年三市出口合计约59亿美元、仅相当于GDP的4%左右；但湛江实际利用外资2021\u20142024年连续四年保持30亿美元量级以上，四年累计约154亿美元，为全国地级市中的外资高地。"));
children.push(bodyLead("创新：", "三市R&D经费十四年间扩张4\u201412倍，湛江投入强度率先逼近2%；但财政科技支出占比均不足0.6%，且年度波动大，稳定投入机制尚未形成。"));
children.push(bodyLead("提示：", "原表存在若干单位混用与疑似填报误差（详见第七部分），本稿已在口径说明中逐项处理，引用带疑点年份的数据前建议与统计公报核对。"));

// ── 一、数据来源与口径说明 ──
children.push(h1("一、数据来源与口径说明"));
children.push(body("本稿依据砥砺慢行整理的《原始表 湛江+阳江统计表》（含湛江、阳江、茂名三个分表，下称统计表）整理计算，覆盖2010\u20142025年地区生产总值（GDP）、人均GDP、固定资产投资、外贸出口总额、实际利用外资、R&D经费、财政科技支出及其占一般公共预算支出比重、实际利用外商直接投资、港口货物吞吐量、万人发明专利授权量等12项指标。各节分析按主题模块化撰写，可直接嵌入文章对应部分；图表均可单独取用。"));
children.push(body("文中趋势计算均基于2010\u20142024年完整年度数据；2025年为初步或不完整数据（详见第七部分），仅作当期参考、不纳入增速计算。其他口径说明如下："));
children.push(numItem("年均增速指2010\u20142024年复合增长率（CAGR），按当年价计算，未剔除价格因素；", "list-scope"));
children.push(numItem("出口统一以亿美元计：湛江2016\u20142024年出口数据原表以亿元计价，为保持可比，按国家外汇管理局公布的当年年均汇率折算为亿美元；", "list-scope"));
children.push(numItem("R&D经费投入强度＝R&D经费\u00f7GDP，由本稿按统计表数据计算（原表该列表头单位标注为亿元，实为R&D经费绝对额）；", "list-scope"));
children.push(numItem("引用标注建议统一为：数据来源：砥砺慢行。", "list-scope"));

// ── 二、总体格局 ──
children.push(h1("二、总体格局：三市合计已近万亿，增长中枢降至5%左右"));
children.push(bodyLead("三市体量合计逼近万亿，茂名反超湛江居首。", "2024年湛江、茂名、阳江三市GDP合计9541.55亿元，已逼近万亿量级。其中茂名以4072.04亿元居首，湛江3839.93亿元次之，两者体量相当；阳江1629.58亿元，约为前两市的四成。体量差距的背后是人口基数差异\u2014\u2014按GDP与人均GDP推算，湛江常住人口约710万、茂名约626万、阳江约262万，湛江的总量优势相当程度上来自人口规模。"));
children.push(bodyLead("增速三段式换挡。", "2010\u20142024年三市GDP年均名义增速分别为湛江7.5%、茂名7.5%、阳江6.9%，十四年间总量分别扩张1.76倍、1.75倍、1.54倍。但增长呈明显的阶段性放缓：2010\u20142015年三市年均名义增速均在10.5%\u201411.5%之间；2015\u20142020年回落至4.0%\u20146.1%；2020\u20142024年维持在4.8%\u20145.6%。粤西与全国同步进入增速换挡期，增长中枢由两位数降至5%左右。"));
children.push(bodyLead("人均水平：湛江总量最大、人均居末的格局十四年未变。", "2024年人均GDP茂名65020元、阳江62100元、湛江54087元。事实上2010年湛江人均GDP（19945元）即低于阳江、茂名约22%\u201428%，十四年后这一位次并未改变，说明湛江的追赶慢于总量扩张。同时也要看到，三市人均GDP与全国平均水平（2024年约9.5万元）仍有约三成差距，粤西整体仍处于追赶阶段。"));
children.push(tableTitle("表1  三市经济总量与人均水平（2010/2024年）"));
children.push(makeTable(
  ["指标", "湛江", "茂名", "阳江"],
  [
    ["GDP（亿元，2010年）", "1390.02", "1482.26", "641.99"],
    ["GDP（亿元，2024年）", "3839.93", "4072.04", "1629.58"],
    ["GDP年均名义增速（2010\u20142024）", "7.5%", "7.5%", "6.9%"],
    ["人均GDP（元，2010年）", "19945", "25328", "25640"],
    ["人均GDP（元，2024年）", "54087", "65020", "62100"],
    ["2024年GDP占三市合计比重", "40.2%", "42.7%", "17.1%"],
  ],
  [40, 20, 20, 20]
));
children.push(note("注：年均增速为名义复合增速；人口为按GDP\u00f7人均GDP推算的常住人口规模。数据来源：砥砺慢行。"));
children.push(fig("charts/gdp.png"));
children.push(fig("charts/pcgdp.png"));

// ── 三、基础支撑能力 ──
children.push(h1("三、基础支撑能力"));
children.push(h2("（一）固定资产投资：普遍见顶回落，进入投资消化期"));
children.push(body("三市固定资产投资分别在2017年（湛江，1641.53亿元）、2021年（阳江，1008.09亿元）、2023年（茂名，1515.20亿元）见顶，此后普遍回落。至2024年，湛江完成1136.99亿元，较峰值下降约31%；阳江673.90亿元，较峰值下降约33%；茂名1489.44亿元，与峰值基本持平（-1.7%）。投资高增长期整体结束，粤西进入以存量项目投产见效为主的投资消化期。"));
children.push(body("以固定资产投资率（固定资产投资\u00f7GDP）观察，三市走势明显分化：阳江由2010年的51.3%升至2021年66.7%的高点后回落至2024年的41.4%，仍为三市最高；湛江由41.2%持续降至29.6%；茂名则由16.5%一路升至36.6%。湛江投资强度的下降与其固定资产投资统计口径调整及重大项目投资周期有关，茂名投资强度的持续抬升则与近年绿色化工、港口建设等大项目集中推进相吻合。"));
children.push(tableTitle("表2  三市固定资产投资关键指标"));
children.push(makeTable(
  ["指标", "湛江", "茂名", "阳江"],
  [
    ["投资峰值年份", "2017", "2023", "2021"],
    ["峰值规模（亿元）", "1641.53", "1515.20", "1008.09"],
    ["2024年规模（亿元）", "1136.99", "1489.44", "673.90"],
    ["2024年较峰值变化", "-30.7%", "-1.7%", "-33.2%"],
    ["投资率2010年（固投/GDP）", "41.2%", "16.5%", "51.3%"],
    ["投资率2024年", "29.6%", "36.6%", "41.4%"],
  ],
  [40, 20, 20, 20]
));
children.push(note("注：湛江2018、2019年固定资产投资原表同为1260亿元，疑为重复填报，该两年数据未用于趋势判断。数据来源：砥砺慢行。"));
children.push(h2("（二）港口货物吞吐量：湛江规模领先，阳江增速最快"));
children.push(bodyLead("港口是基础支撑能力变化最直观的指标。", "湛江港货物吞吐量由2010年的1.36亿吨增至2024年的2.75亿吨，十四年翻番，规模稳居三市之首（其中2018\u20142019年数据出现一次台阶式回落，疑为统计口径调整，不影响长期趋势，见图3）；阳江港由779万吨增至4041.57万吨，累计增长4.2倍、年均增长12.5%，是三市中增长最快的；茂名港由2284万吨增至3367万吨，累计增长47%。2025年初步数据显示，阳江港（4460万吨）、茂名港（3949万吨）继续上行，湛江港小幅回落至2.65亿吨。"));
children.push(body("这一格局与三市产业结构相互印证：湛江依托钢铁、石化、造纸等重大临港产业布局，港口吞吐以大宗散货和液体化工品为主，总量优势明显；阳江港的增长则更多来自临港工业与粮油加工等产业的起步放量，基数低、增速高；茂名港体量居中，增长平缓。"));
children.push(fig("charts/port.png"));

// ── 四、开放发展能力 ──
children.push(h1("四、开放发展能力"));
children.push(h2("（一）外贸出口：规模长期徘徊，开放是突出短板"));
children.push(body("统一为亿美元口径后，湛江出口由2010年的16.87亿美元增至2024年的约27.3亿美元（其间2017年高点约32.2亿美元），阳江由16.06亿美元增至23.93亿美元，茂名由5.59亿美元增至8.04亿美元。十四年间三市名义GDP增长了1.5倍以上，而出口仅增长44%\u201462%，出口增速大幅落后于经济总量扩张。"));
children.push(body("尤其值得注意的是茂名出口的剧烈波动：2019年出口达24.69亿美元的高点，此后连续回落，2023年仅6.97亿美元，四年间缩水约三分之二，2024年小幅回升至8.04亿美元。以2024年计，三市出口合计约59亿美元、约合420亿元人民币，仅相当于三市GDP的4%左右，远低于全国平均外贸依存度。外贸规模与三市近万亿的经济体量明显不匹配，开放型经济是粤西与珠三角差距最直观的领域。"));
children.push(tableTitle("表3  三市外贸出口总额（亿美元）"));
children.push(makeTable(
  ["指标", "湛江", "茂名", "阳江"],
  [
    ["2010年", "16.87", "5.59", "16.06"],
    ["2024年", "27.3", "8.04", "23.93"],
    ["2010\u20142024年变化", "+62%", "+44%", "+49%"],
    ["期间高点（年份）", "32.2（2017）", "24.69（2019）", "28.92（2021）"],
  ],
  [34, 22, 22, 22]
));
children.push(note("注：湛江2016\u20142024年原表数据以亿元计价，已按当年年均汇率折算为亿美元。数据来源：砥砺慢行。"));
children.push(h2("（二）利用外资：湛江一骑绝尘，重大项目主导格局形成"));
children.push(bodyLead("湛江FDI跃升是统计表中最突出的结构性变化。", "湛江实际利用外商直接投资由2010年的3671万美元增至2024年的37.54亿美元，年均增速高达39%。特别是2021\u20142024年，湛江FDI连续四年保持在30亿美元量级以上（29.68亿、40.82亿、46.24亿、37.54亿美元），四年累计约154亿美元\u2014\u2014以巴斯夫（广东）一体化基地为代表的大型外资项目进入集中投资期，湛江由此成为全国地级市中吸引外资的高地。"));
children.push(body("相比之下，阳江FDI在2018年冲高至7.63亿美元后回落，2024年为4561万美元；茂名长期低位徘徊，2022年以来连续三年不足3000万美元。在出口徘徊不前的背景下，三市开放能力的差距正从贸易领域向资本流动领域进一步拉大，湛江以重大项目引外资的路径已与其他两市明显分层。"));
children.push(fig("charts/fdi_zj.png"));

// ── 五、创新支持水平 ──
children.push(h1("五、创新支持水平"));
children.push(h2("（一）R&D经费：基数小、增速快，湛江强度率先逼近2%"));
children.push(body("湛江R&D经费由2010年的10.7亿元增至2024年的75亿元，占GDP比重由0.77%升至1.95%，是三市中唯一接近全国平均水平（2024年约2.68%）的城市；阳江R&D经费由1.79亿元增至22.69亿元，年均增长19.9%、为三市最快，投入强度由0.28%升至1.39%；茂名由3.94亿元增至20.76亿元，强度仅由0.27%升至0.51%，创新投入强度与其经济体量明显不匹配。总体看，三市创新投入仍处于\u201c低基数、高增速\u201d的补课阶段。"));
children.push(body("需要提示的是，三市R&D数据的年度波动较大（湛江2024年数值骤增、阳江2021年骤降，详见第七部分），个别年份数值疑有填报误差，引用时建议以\u201c十四五\u201d以来的总体趋势为准，并与各市统计公报核对。"));
children.push(fig("charts/rd.png"));
children.push(h2("（二）财政科技支出：规模小、占比低、未形成稳定投入机制"));
children.push(body("2024年三市财政科技支出占一般公共预算支出的比重均不足0.6%：湛江0.54%、阳江0.41%、茂名0.36%；绝对规模上，三市财政科技支出均不超过3亿元（湛江2.89亿元、茂名1.84亿元、阳江1.04亿元）。从历史轨迹看，湛江（2016年1.21%）、阳江（2019年1.64%）的科技支出占比都曾达到1%以上，但均未能维持，说明粤西的财政科技投入尚未形成跨周期的稳定机制，科技投入的\u201c刚性\u201d不足。"));
children.push(tableTitle("表4  三市财政科技支出及其占比"));
children.push(makeTable(
  ["指标", "湛江", "茂名", "阳江"],
  [
    ["2024年科技支出（亿元）", "2.89", "1.84", "1.04"],
    ["2024年占一般公共预算支出比重", "0.54%", "0.36%", "0.41%"],
    ["2010年比重", "0.38%", "0.48%", "1.23%"],
    ["历史峰值比重（年份）", "1.21%（2016）", "0.79%（2018）", "1.64%（2019）"],
  ],
  [40, 20, 20, 20]
));
children.push(note("数据来源：砥砺慢行。"));
children.push(h2("（三）万人发明专利授权量：趋势向好，绝对水平仍低"));
children.push(body("以近5年累计口径计，万人发明专利授权量湛江由2010年的1.32件增至2024年的5.56件，阳江由0.71件增至3.28件，茂名由0.55件增至2.34件，三市年均增速均在11%左右，创新产出的改善趋势一致。但绝对水平仍然偏低，且茂名2020年（5.74件）至2021年（1.82件）出现断崖式回落，疑为统计口径调整所致，相关年份数据引用前需核实。"));
children.push(fig("charts/patent.png"));

// ── 六、总览表 ──
children.push(h1("六、三市核心指标总览（2010\u21922024年）"));
children.push(tableTitle("表5  三市核心指标总览"));
children.push(makeTable(
  ["指标", "湛江 2010\u21922024", "增速", "茂名 2010\u21922024", "增速", "阳江 2010\u21922024", "增速"],
  [
    ["GDP（亿元）", "1390\u21923840", "7.5%", "1482\u21924072", "7.5%", "642\u21921630", "6.9%"],
    ["人均GDP（元）", "19945\u219254087", "7.4%", "25328\u219265020", "7.0%", "25640\u219262100", "6.5%"],
    ["固定资产投资（亿元）", "573\u21921137", "5.0%", "245\u21921489", "13.8%", "329\u2192674", "5.3%"],
    ["港口吞吐量（万吨）", "13638\u219227500", "5.1%", "2284\u21923367", "2.8%", "779\u21924042", "12.5%"],
    ["外贸出口（亿美元）", "16.9\u219227.3", "3.5%", "5.6\u21928.0", "2.6%", "16.1\u219223.9", "2.9%"],
    ["实际利用FDI（万美元）", "3671\u2192375400", "39.2%", "3099\u21921175", "-6.7%", "20751\u21924561", "-10.3%"],
    ["R&D经费（亿元）", "10.7\u219275.0", "14.9%", "3.9\u219220.8", "12.6%", "1.8\u219222.7", "19.9%"],
    ["R&D投入强度", "0.77%\u21921.95%", "\u2014", "0.27%\u21920.51%", "\u2014", "0.28%\u21921.39%", "\u2014"],
    ["财政科技支出占比", "0.38%\u21920.54%", "\u2014", "0.48%\u21920.36%", "\u2014", "1.23%\u21920.41%", "\u2014"],
    ["万人发明专利（件）", "1.32\u21925.56", "10.8%", "0.55\u21922.34", "10.9%", "0.71\u21923.28", "11.6%"],
  ],
  [26, 15, 9.5, 15, 9.5, 15, 10],
  { dataSize: 18 }
));
children.push(note("注：出口为亿美元口径，湛江2016\u20142024年按当年年均汇率折算；年均增速为2010\u20142024年复合增速，\u2014\u2014表示比率类指标不适用；R&D强度、投资率由本稿按统计表计算。数据来源：砥砺慢行。"));

// ── 七、数据质量提示 ──
children.push(h1("七、数据质量提示（引用前建议核实）"));
children.push(body("整理过程中发现统计表存在以下疑点，本稿已作相应处理；文章引用相关年份数据前，建议与各市统计年鉴、统计公报交叉核对："));
children.push(numItem("湛江出口总额列存在单位混用：2016\u20142024年为亿元人民币，其余年份为亿美元（按汇率折算后序列连续可证）。本稿已统一折算为亿美元。", "list-quality"));
children.push(numItem("湛江分表中" + Q("实际利用外资额（亿元）") + "列实际填入的是亿美元数值（与" + Q("实际利用外商直接投资（万美元）") + "列为同一数据），该列缺少真正的亿元口径数据。", "list-quality"));
children.push(numItem("表头" + Q("R&D经费投入强度（亿元）") + "单位标注有矛盾：按数值判断该列实为R&D经费绝对额（亿元），本稿所称投入强度均另行计算。", "list-quality"));
children.push(numItem("湛江固定资产投资2018、2019年同为1260亿元，疑似重复填报或占位数据。", "list-quality"));
children.push(numItem("湛江R&D经费2024年（75亿元）较2023年（41.26亿元）骤增82%，且2025年数值（41.26亿元）与2023年完全相同，疑有填报误差。", "list-quality"));
children.push(numItem("阳江R&D经费2021年（0.4576亿元）较2020年（3.07亿元）骤降85%；茂名万人发明专利授权量2020年（5.74件）至2021年（1.82件）断崖回落\u2014\u2014两处疑似统计口径调整。", "list-quality"));
children.push(numItem("2025年数据完整性不足：三市实际利用外商直接投资较2024年骤降90%以上，明显为年内不完整数据；GDP、港口吞吐量等为初步数。本稿全部趋势计算仅使用2010\u20142024年数据。", "list-quality"));
children.push(numItem("湛江2024年外商直接投资原单元格为文本" + Q("3754 00") + "，本稿按375400万美元处理。", "list-quality"));
children.push(numItem("湛江港口货物吞吐量2018年（30185万吨）至2019年（21570万吨）骤降29%，与前后年份走势脱节，疑似统计口径调整，引用该两年数据前需核实。", "list-quality"));

// ───────────────────────── document ─────────────────────────
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: { ascii: "Times New Roman", eastAsia: "SimSun" }, size: 24, color: "000000" },
        paragraph: { spacing: { line: 312 } },
      },
      heading1: {
        run: { font: { ascii: "Calibri", eastAsia: "SimHei" }, size: 32, bold: true, color: P.headingColor },
        paragraph: { spacing: { before: 400, after: 180, line: 312 } },
      },
      heading2: {
        run: { font: { ascii: "Calibri", eastAsia: "SimHei" }, size: 28, bold: true, color: P.headingColor },
        paragraph: { spacing: { before: 280, after: 120, line: 312 } },
      },
    },
  },
  numbering: {
    config: [
      {
        reference: "list-scope",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
      {
        reference: "list-quality",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: { size: { width: 11906, height: 16838 }, margin: { top: 0, bottom: 0, left: 0, right: 0 } },
      },
      children: buildCoverR1({
        title: "粤西三市数据补充分析稿",
        subtitle: "湛江 \u00b7 阳江 \u00b7 茂名：基础支撑、开放发展与创新支持（2010\u20142025）",
        englishLabel: "DATA ANALYSIS SUPPLEMENT",
        metaLines: [
          "数据来源：砥砺慢行（湛江+阳江统计表）",
          "覆盖年份：2010\u20142025年（趋势分析基于2010\u20142024年）",
          "整理时间：2026年10月",
        ],
        footerLeft: "数据补充分析稿",
        footerRight: "2026年10月",
        palette: P,
      }),
    },
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, bottom: 1440, left: 1701, right: 1417 },
          pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL },
        },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "808080" })],
          })],
        }),
      },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("粤西三市数据补充分析稿.docx", buf);
  console.log("OK 粤西三市数据补充分析稿.docx", buf.length, "bytes");
});
