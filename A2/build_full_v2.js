const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, LevelFormat,
} = require('docx');

const FONT = { ascii: 'Calibri', hAnsi: 'Calibri', eastAsia: 'Microsoft JhengHei', cs: 'Calibri' };
const ACCENT = '4B2E83';

// A paragraph is a list of segments; [text, true] marks a revised segment (highlighted).
const para = (segs, opts = {}) => new Paragraph({
  spacing: { after: 160, line: 300 },
  children: segs.map(s => {
    const [text, hl] = Array.isArray(s) ? s : [s, false];
    return new TextRun({ text, font: FONT, size: 22, highlight: hl ? 'yellow' : undefined, ...opts });
  }),
});
const zh = segs => para(segs, { color: '555555' });
const label = t => new Paragraph({ spacing: { before: 280, after: 100 },
  children: [new TextRun({ text: t, font: FONT, size: 24, bold: true, color: ACCENT })] });
const plain = (t, o = {}) => new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: t, font: FONT, size: 21, ...o })] });
const num = t => new Paragraph({ numbering: { reference: 'nums', level: 0 }, spacing: { after: 60 },
  children: [new TextRun({ text: t, font: FONT, size: 21 })] });

function table(widths, rows) {
  const b = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
  return new Table({
    width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({ tableHeader: i === 0, children: r.map((c, j) => new TableCell({
      width: { size: widths[j], type: WidthType.DXA },
      borders: { top: b, bottom: b, left: b, right: b },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      shading: i === 0 ? { type: ShadingType.CLEAR, color: 'auto', fill: 'E8E1F3' } : undefined,
      children: [new Paragraph({ children: [new TextRun({ text: c, font: FONT, size: 20, bold: i === 0 })] })],
    })) })),
  });
}
const raw = fs.readFileSync(__dirname + '/full_text_v2_en.txt', 'utf8');
const ZH = require('./full_text_v2_zh.js');
const S = raw.split('\n# ').map(b => { const l = b.replace(/^# /, '').split('\n').filter(Boolean); return { h: l[0], ps: l.slice(1) }; });
const wc = t => t.split(/\s+/).filter(Boolean).length;
const counts = S.map(s => s.ps.reduce((a, p) => a + wc(p), 0));
const total = counts.reduce((a, b) => a + b, 0);
if (S.reduce((a, s) => a + s.ps.length, 0) !== ZH.length) throw new Error('paragraph count mismatch');
// Segments to highlight (wording changed beyond plain deletion)
const HL = ['First, ', 'Second, ', 'Third, ', 'To track all three', '(Goleman, 1998)', 'Appendix A, asks', 'Measured by'];
const hlSegs = t => { let segs = [t]; for (const k of HL) segs = segs.flatMap(x => typeof x !== 'string' ? [x] : x.split(k).flatMap((y, i) => i ? [[k, true], y] : [y])); return segs.filter(x => x !== ''); };

const c = [];
c.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: 'HOSP7053 Assessment 2 全文 刪減版', font: FONT, size: 32, bold: true, color: ACCENT })] }));
c.push(plain('以她的版本為準，只刪字和合併句子，沒有加回任何舊內容。每段先英文再中文，黃色是除了刪字以外有改到措辭的地方。'));
c.push(table([3600, 1800, 3626], [
  ['節', '原稿', '刪減後'],
  ...S.map((s, i) => [s.h, ['223', '1,032', '877', '732', '357'][i], String(counts[i])]),
  ['合計（不含標題、References、Appendix）', '3,221', String(total)],
]));
let k = 0;
for (const s of S) { c.push(label(s.h)); for (const p of s.ps) { c.push(para(hlSegs(p))); c.push(zh([ZH[k++]])); } }

const docA = new Document({ styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: c }] });
const e = [];
for (const s of S) {
  e.push(new Paragraph({ spacing: { before: 240, after: 120 }, children: [new TextRun({ text: s.h, font: FONT, size: 26, bold: true })] }));
  s.ps.forEach(p => e.push(new Paragraph({ spacing: { after: 160, line: 360 }, children: [new TextRun({ text: p, font: FONT, size: 22 })] })));
}
const docB = new Document({ styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: e }] });
Promise.all([Packer.toBuffer(docA), Packer.toBuffer(docB)]).then(([a, b]) => {
  fs.writeFileSync(__dirname + '/HOSP7053_A2_全文_刪減版_中英對照.docx', a);
  fs.writeFileSync(__dirname + '/HOSP7053_A2_全文_刪減版_英文.docx', b);
  console.log('written', counts, total);
});
