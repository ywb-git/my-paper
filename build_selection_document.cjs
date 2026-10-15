const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, HeadingLevel, AlignmentType, WidthType, BorderStyle,
  ShadingType, ExternalHyperlink, PageNumber, TableLayoutType,
} = require('./.document-tools/node_modules/docx');

const base = '大模型材料逆向设计_硕士选题比较_2026-10-15';
const source = fs.readFileSync(path.join(__dirname, base + '.md'), 'utf8');
const bodyFont = { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'Noto Serif CJK SC', cs: 'Arial' };
const headFont = { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'Noto Sans CJK SC', cs: 'Arial' };
const width = 9506;
const border = { style: BorderStyle.SINGLE, size: 4, color: 'D8DEE6' };
const borders = { top: border, bottom: border, left: border, right: border };

function runs(text, options = {}) {
  const pieces = text.split(/(https?:\/\/[^\s]+)/g).filter(Boolean);
  return pieces.map(piece => /^https?:\/\//.test(piece)
    ? new ExternalHyperlink({ link: piece, children: [new TextRun({ text: piece, color: '245B7A', underline: {}, size: 18, font: bodyFont })] })
    : new TextRun({ text: piece, font: bodyFont, ...options }));
}
function para(text, options = {}) {
  return new Paragraph({ spacing: { after: 100, line: 310 }, widowControl: true, children: runs(text), ...options });
}
function table(lines) {
  const rows = lines.map(line => line.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(cell => cell.trim()))
    .filter(row => !row.every(cell => /^:?-+:?$/.test(cell)));
  const cols = rows[0].length;
  const widths = cols === 4 ? [2850, 2050, 2400, 2206] : [1700, 3900, 3906];
  return new Table({
    width: { size: width, type: WidthType.DXA }, columnWidths: widths,
    layout: TableLayoutType.FIXED, margins: { top: 110, bottom: 110, left: 130, right: 130 },
    rows: rows.map((row, i) => new TableRow({ tableHeader: i === 0, cantSplit: true,
      children: row.map((text, j) => new TableCell({ width: { size: widths[j], type: WidthType.DXA }, borders,
        shading: { fill: i === 0 ? 'E8EEF4' : (i % 2 ? 'FFFFFF' : 'F5F7FA'), type: ShadingType.CLEAR },
        children: [new Paragraph({ spacing: { after: 20, line: 270 }, children: runs(text, { size: 19, bold: i === 0, font: headFont }) })],
      })),
    })),
  });
}

const children = [];
const lines = source.split(/\r?\n/);
for (let i = 0; i < lines.length; i++) {
  const line = lines[i].trim();
  if (!line) continue;
  if (line.startsWith('|')) {
    const group = [];
    while (i < lines.length && lines[i].trim().startsWith('|')) group.push(lines[i++]);
    i--;
    children.push(table(group));
    children.push(para(' ', { spacing: { after: 40 } }));
  } else if (line.startsWith('# ')) {
    children.push(new Paragraph({ heading: HeadingLevel.TITLE, spacing: { before: 120, after: 140 }, children: [new TextRun({ text: line.slice(2), font: headFont, size: 44, bold: true, color: '173349' })] }));
  } else if (i < 3 && line.startsWith('## ')) {
    children.push(para(line.slice(3), { spacing: { after: 160 }, children: [new TextRun({ text: line.slice(3), font: headFont, size: 30, color: '3C566B' })] }));
  } else if (line.startsWith('## ')) {
    const title = line.slice(3);
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: /^[3-8]\. 方向|^12\. 代表文献/.test(title), children: [new TextRun({ text: title, font: headFont })] }));
  } else if (line.startsWith('### ')) {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: line.slice(4), font: headFont })] }));
  } else {
    children.push(para(line));
  }
}
const doc = new Document({
  title: '大模型指导材料逆向设计：硕士论文选题比较与算法创新分析',
  subject: '选题比较、算法创新边界、难度和验证方案',
  creator: 'AI 辅助研究', description: '六类候选方向的初步比较，不构成全面查新结论。',
  styles: {
    default: { document: { run: { font: bodyFont, size: 21, color: '202A33' }, paragraph: { spacing: { line: 310, after: 100 } } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', run: { font: headFont, size: 44, bold: true }, paragraph: { keepNext: true } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: headFont, size: 29, bold: true, color: '173349' }, paragraph: { spacing: { before: 220, after: 140 }, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: headFont, size: 23, bold: true, color: '294E69' }, paragraph: { spacing: { before: 150, after: 85 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1100, bottom: 1100, left: 1200, right: 1200, header: 450, footer: 450 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: '材料逆向设计 · 硕士选题比较', font: headFont, size: 17, color: '6A7783' })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '2026-10-15    ·    ', size: 17, color: '6A7783' }), new TextRun({ children: [PageNumber.CURRENT], size: 17 }), new TextRun({ text: ' / ', size: 17 }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 17 })] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then(buffer => {
  const output = path.join(__dirname, base + '.docx');
  if (fs.existsSync(output)) throw new Error('拒绝覆盖已有输出：' + output);
  fs.writeFileSync(output, buffer);
  console.log(output);
}).catch(error => { console.error(error); process.exitCode = 1; });
