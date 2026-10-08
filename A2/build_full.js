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
const S = require('./full_text.js');
const txt = segs => segs.map(x => Array.isArray(x) ? x[0] : x).join('');
const wc = t => t.split(/\s+/).filter(Boolean).length;
const counts = S.map(s => s.ps.reduce((a, p) => a + wc(txt(p.en)), 0));
const total = counts.reduce((a, b) => a + b, 0);

// References: segments [text, italic]
const REFS = [
  [['Fleeson, W. (2001). Toward a structure- and process-integrated view of personality: Traits as density distributions of states. '], ['Journal of Personality and Social Psychology, 80', 1], ['(6), 1011–1027. https://doi.org/10.1037/0022-3514.80.6.1011']],
  [['George, B., Sims, P., McLean, A. N., & Mayer, D. (2007). Discovering your authentic leadership. '], ['Harvard Business Review, 85', 1], ['(2), 129–138.']],
  [['Goleman, D. (1998). What makes a leader? '], ['Harvard Business Review, 76', 1], ['(6), 93–102.']],
  [['Gross, J. J., & John, O. P. (2003). Individual differences in two emotion regulation processes: Implications for affect, relationships, and well-being. '], ['Journal of Personality and Social Psychology, 85', 1], ['(2), 348–362. https://doi.org/10.1037/0022-3514.85.2.348']],
  [['Hogreve, J., Iseke, A., & Derfuss, K. (2022). The service-profit chain: Reflections, revisions, and reimaginations. '], ['Journal of Service Research, 25', 1], ['(3), 460–477. https://doi.org/10.1177/10946705211052410']],
  [['Latemore, G. (2024). Leadership for service organizations. In J. Kandampully, D. Solnet, & A. Bilgihan (Eds.), '], ['Service management principles for hospitality and tourism in the age of digital technology', 1], [' (4th ed., pp. 263–298). Kendall Hunt.']],
  [['Northouse, P. G. (2026). '], ['Leadership: Theory and practice', 1], [' (10th ed.). SAGE.']],
  [['Wong, C.-S., & Law, K. S. (2002). The effects of leader and follower emotional intelligence on performance and attitude: An exploratory study. '], ['The Leadership Quarterly, 13', 1], ['(3), 243–274. https://doi.org/10.1016/S1048-9843(02)00099-1']],
];
const refPara = (r, size = 22) => new Paragraph({ spacing: { after: 120, line: 360 }, indent: { left: 720, hanging: 720 },
  children: r.map(([t, it]) => new TextRun({ text: t, font: FONT, size, italics: !!it })) });

// ---------- Bilingual review copy ----------
const c = [];
c.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: 'HOSP7053 Assessment 2 全文 2,500 字版', font: FONT, size: 32, bold: true, color: ACCENT })] }));
c.push(plain('每段先英文再中文。黃色螢光筆為內容有改動或補上的地方（純刪字不標）。中文對照與本頁說明交件前全部刪除，交件請用「英文交件版」。'));
c.push(label('字數'));
c.push(table([3200, 1800, 1800, 2226], [
  ['節', '原稿', '目標', '本版'],
  ...S.map((s, i) => [s.h, ['223', '1,032', '877', '732', '357 (缺第一項)'][i], ['~200', '~700', '~600', '~600', '~400'][i], String(counts[i])]),
  ['合計（不含標題、References、Appendix）', '3,221', '≤ 2,500', String(total)],
]));
for (const s of S) {
  c.push(label(s.h));
  for (const p of s.ps) { c.push(para(p.en)); c.push(zh(p.zh)); }
}
c.push(label('References'));
REFS.forEach(r => c.push(refPara(r)));
c.push(plain('George et al. (2007) 正文沒有引用，只在 Appendix 用到才留；Appendix 沒用就刪。Hersey & Blanchard (1969) 和 Preston (2001) 是轉引，不列，只列 Latemore (2024)。文獻標題裡的冒號是原標題，APA 必須保留。', { color: '555555' }));

c.push(label('主要修改'));
c.push(table([2000, 3800, 3226], [
  ['位置', '修改', '原因'],
  ['Intro', '刪掉 A reader could argue 的寫法、合併 Row 2/Row 4 兩句、路線圖改成 I then trace…', '省字，意思不變'],
  ['Repertoire', '刪掉 Mayer and Salovey (1997)，四個面向合成一句；刪掉 The rest of this section asks… 與 Their model treats these as abilities…（最後一段已講）', '省字，也避免 Mayer & Salovey 沒讀原文卻列年份的問題'],
  ['Repertoire', 'Fleeson 句改成你之前看得懂的版本，補回「和人與人之間的差異一樣大」', '原句漏了這個比較，Fleeson 的重點就不見了'],
  ['Repertoire', 'judgment 改 judgement', '全文英式拼法一致（behaviour, practise, apologise）'],
  ['Impact', '開頭改回 SPC 版（內部服務品質進入），並把「每一列的客人都得到需要的」改成「每一列我介入時，眼前的客人」', '老師要看 SPC；原句和 Ethics 裡「房內客人沒被告知」矛盾'],
  ['Impact', 'Row 2 補 Hogreve (2022, pp. 462, 464)、改回「服務是透過我而不是透過他完成」；刪掉 To him, it may also have felt…', '補頁碼；那句是推測，沒有證據支撐'],
  ['Impact', 'Row 4 補 p. 463，Row 4 開頭補「我不是在場資深者」', '補頁碼；開頭刪掉了位置對比，所以放回 Row 4'],
  ['Impact', '開頭段刪掉 Rows 1、2、5 和 Row 4 的對比，結尾段刪掉 Stepping back 那段重複', '兩段原本講同一件事，Repertoire 也講過'],
  ['Ethics', '補 Preston p. 284、Northouse pp. 465, 466 / p. 462 / p. 466；補「genuinely open」兩句；(Row 5) 改 (Appendix A, Row 5)', '這些是之前已確認的版本，貼上來的稿子沒有放進去'],
  ['Ethics', 'Northouse 的 honesty 兩句合成一句；Row 1 例子縮短；兩種解讀的段落合句', '省字'],
  ['Development', '補回第一項（情境領導 / Row 2）', '貼上來的稿子從 Second 開始，第一項不見了'],
  ['Development', '刪掉 in the future', '小組作業、打工、志工都是現在進行中的，第四項也寫 this semester'],
  ['Development', '第二項 Goleman 改成 (Goleman, 1998)，不再重講定義', 'Repertoire 已直接引用 p. 95 的定義，這裡不用重複'],
]));

c.push(label('需要她確認的事'));
[
  'Development 第一項是不是她刪掉的？如果是故意刪，全文會變成三項改變，仍符合老師說的 3 到 4 項，可以直接刪掉第一項，省約 55 字。',
  'Impact 開頭用 SPC 版（我建議），還是保留她原本「In every row that involved a guest…」的寫法？原寫法和 Ethics 的房內客人矛盾，所以我改成「我介入時眼前的客人」。',
  'Appendix 裡有沒有寫 Mayer and Salovey (1997)？有的話 References 要處理（改成轉引或刪年份）；正文已經不需要。',
  'Appendix 裡有沒有用到 George et al. (2007)？沒有就從 References 刪掉。',
  'Appendix 待修：Row 4 補上同事自己完成入住；Row 5、6、7 的 Goleman 原文加引號和 (p. 95)；Row 7 states as 改 defines as；References 的全形冒號；把 Hersey & Blanchard (1969) 換成 Latemore (2024)。',
  'AI 使用說明：如果課程要求，要寫出用 AI 檢查文法、刪減字數和核對引用頁碼。',
].forEach(t => c.push(num(t)));

const docA = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: 'nums', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 340 } } } }] }] },
  sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: c }],
});

// ---------- Clean English submission copy ----------
const e = [];
const h = t => new Paragraph({ spacing: { before: 240, after: 120 }, children: [new TextRun({ text: t, font: FONT, size: 26, bold: true })] });
const body = t => new Paragraph({ spacing: { after: 160, line: 360 }, children: [new TextRun({ text: t, font: FONT, size: 22 })] });
for (const s of S) { e.push(h(s.h)); s.ps.forEach(p => e.push(body(txt(p.en)))); }
e.push(h('References'));
REFS.forEach(r => e.push(refPara(r)));
const docB = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: e }],
});

Promise.all([Packer.toBuffer(docA), Packer.toBuffer(docB)]).then(([a, b]) => {
  fs.writeFileSync(__dirname + '/HOSP7053_A2_全文_2500字版.docx', a);
  fs.writeFileSync(__dirname + '/HOSP7053_A2_全文_英文交件版.docx', b);
  console.log('written', counts, total);
});
