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

const c = [];
c.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: 'HOSP7053 A2', font: FONT, size: 32, bold: true, color: ACCENT })] }));
c.push(new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: 'Your repertoire 修訂版（英中對照）', font: FONT, size: 28, bold: true })] }));
c.push(new Paragraph({ spacing: { after: 240 }, children: [
  new TextRun({ text: '標示說明 ', font: FONT, size: 21, bold: true }),
  new TextRun({ text: '螢光筆', font: FONT, size: 21, highlight: 'yellow' }),
  new TextRun({ text: ' 是這次修改或新增的地方，其餘保留原文。刪掉的句子無法標示，列在最後的修改表裡。中文對照交件前要全部刪除。', font: FONT, size: 21 }),
] }));

// ---------- P1 (unchanged) ----------
c.push(label('第 1 段'));
c.push(para(['Across my rows, when someone in front of me is clearly in difficulty, I step in to protect them (Rows 2, 4, 5 and 7), and in Rows 2, 5 and 7 I took the whole problem on myself. This is the approach I overuse, and it comes from being the senior or more experienced person present and knowing I can handle it. When no one is in difficulty, I listen and weigh other people\'s views before acting (Rows 1 and 3). The rest of this section asks what decides my use of EI once a problem is mine, and what makes it mine in the first place.']));
c.push(zh(['綜觀我的各列，當我面前有人明顯陷入困難時，我會介入保護他們（第 2、4、5、7 列），而在第 2、5、7 列，我把整個問題都攬到自己身上。這是我用得太多的做法，它來自我是在場較資深或較有經驗的人，也知道自己處理得來。當沒有人陷入困難時，我會在行動前傾聽並權衡別人的觀點（第 1、3 列）。這一節接下來要問的是，當問題成為我的問題之後，是什麼決定我會不會運用 EI，以及一開始是什麼讓問題變成我的。']));

// ---------- P2 ----------
c.push(label('第 2 段'));
c.push(para([
  'Read through emotional intelligence (EI), my rows suggest that having these abilities does not decide when I use them. Whether I use them depends on whether other people\'s feelings are part of what I judge matters most in the moment. I chose EI because many of my rows turn on how I appraise and respond to emotions, both my own and other people\'s. Wong and Law (2002) ',
  ['build on Mayer and Salovey\'s (1997) definition of EI and measure it as four dimensions, which are', true],
  ' appraising one\'s own emotions, appraising others\' emotions, using emotion and regulating emotion. Their model treats these as abilities that belong to the person, and places context in the demands of the job. I use their model as my lens because the WLEIS, the scale I completed, was built from it. My scores on the scale are only a starting picture, which my rows let me test against what I actually did. Appraising others\' emotions was my highest score, at 6.8 out of 7, and I used it clearly when people\'s feelings were part of my priority. In Row 4, what mattered most was getting the waiting guest checked in, and that depended on the people involved, so I read the impatient guest, my overwhelmed colleague and the supervisor\'s rising tone at the same time. When my priority did not include anyone\'s feelings, I did not use it. In Row 1, as the most senior person pushing my own proposal, what mattered most to me was showing that the change was necessary. ',
  ['I listened to my colleagues\' reasons, because they bore on that, but I missed that they', true],
  ' might be feeling pressured and not safe enough to keep pushing back. ',
  ['What changed across these rows was what I put first.', true],
  ' Because I often take the problem on as my own, I focus on what that problem contains, and when it does not include people\'s feelings, those feelings drop out of view.',
]));
c.push(zh([
  '用情緒智力（EI）來讀，我的各列顯示，擁有這些能力並不能決定我什麼時候會使用它們。我會不會使用，取決於別人的感受是否屬於我當下認為最重要的事。我選擇 EI，是因為我許多列的關鍵，在於我如何評估並回應情緒，包括我自己的和別人的。Wong 和 Law（2002）',
  ['以 Mayer 和 Salovey（1997）對 EI 的定義為基礎，把 EI 分成四個維度來測量，分別是', true],
  '評估自己的情緒、評估別人的情緒、運用情緒，以及調節情緒。他們的模型把這些視為屬於個人的能力，而把情境放在工作的要求上。我用他們的模型作為我的 lens，是因為我做的 WLEIS 量表就是根據這個模型建立的。我在量表上的分數只是一個初步的了解，我的各列讓我能拿實際行為來檢驗它。評估別人的情緒是我的最高分，滿分 7 分中拿到 6.8 分，而當別人的感受屬於我的優先事項時，我明顯用到了這個能力。在第 4 列，我認為最重要的是讓等待中的客人能入住，而這取決於在場的每個人，所以我同時讀出了不耐煩的客人、慌張的同事，以及主管越來越重的語氣。當我的優先事項不包含任何人的感受時，我就沒有用到這個能力。在第 1 列，身為在場最資深、又在推動自己提案的人，我認為最重要的是證明這個改變是必要的。',
  ['我聽了同事的理由，因為那些理由和這件事有關，但我沒察覺到他們', true],
  '可能感到有壓力，不夠安心繼續反對。',
  ['在這幾列之間，改變的是我把什麼放在第一位。', true],
  '因為我常常把問題攬成自己的，我會專注在那個問題包含了什麼，而當問題本身不包含別人的感受時，那些感受就從我的視野裡消失了。',
]));

// ---------- P3 ----------
c.push(label('第 3 段'));
c.push(para([
  'Two weekly concepts show most clearly where the lens falls short. In Row 5, Goleman\'s (1998) empathy overlaps with Wong and Law\'s appraisal of others\' emotions, and both show that I missed how my colleague felt. Read through empathy, as in Appendix A, Row 5 suggests that this ability failed under time pressure. But the lens treats reading others as an ability I carry with me. In Row 4, also under real time pressure during a peak check-in, I used it on everyone. So pressure alone cannot explain why it was missing in Row 5. Comparing the two rows, what differed was what I put first. In Row 4, getting the guest checked in depended on calming the people involved, so their feelings were part of the problem. In Row 5, the problem was stopping a guest before they reached an occupied room, so I needed only facts from her. Goleman, however, also defines empathy as ',
  ['"skill in treating people according to their emotional reactions" (p. 95)', true],
  ', which the lens leaves out. By that half, I also failed in how I treated her, because I pressed her to explain again while she was already rattled, and she looked close to tears. ',
  ['Her feelings were not part of the problem I was solving, so they shaped neither what I noticed nor how I spoke to her.', true],
]));
c.push(zh([
  '有兩個每週概念，最清楚地顯示出這個 lens 不足的地方。在第 5 列，Goleman（1998）的同理心和 Wong 與 Law 的「評估別人的情緒」有重疊，兩者都顯示我沒注意到同事的感受。用同理心來讀，就像 Appendix A 裡的寫法，第 5 列顯示這個能力在時間壓力下失效了。但這個 lens 把讀懂別人當成我隨身帶著的能力。在第 4 列，同樣是在尖峰入住時段、同樣有實際的時間壓力，我卻讀出了每一個人。所以光是壓力，解釋不了為什麼第 5 列少了這個能力。比較這兩列，不同的是我把什麼放在第一位。在第 4 列，讓客人入住取決於安撫在場的人，所以他們的感受是問題的一部分。在第 5 列，問題是在客人走到已有人住的房間前攔住他，所以我只需要從她那裡拿到事實。不過，Goleman 對同理心的定義還包括',
  ['「依照別人的情緒反應來對待他們的技巧」（p. 95）', true],
  '，而這是 lens 沒有涵蓋的。從這一半來看，我在對待她的方式上也失敗了，因為她已經很慌了，我還逼她再解釋一次，她看起來快哭了。',
  ['她的感受不屬於我正在解決的問題，所以既沒有影響我注意到什麼，也沒有影響我怎麼跟她說話。', true],
]));

// ---------- P4 ----------
c.push(label('第 4 段'));
c.push(para([
  'Row 2 shows a different gap. Through the EI lens, I did read my colleague, because I saw that he had frozen and stepped in to protect him. Situational leadership (Hersey & Blanchard, 1969, as cited in Latemore, 2024), the concept I first used for this row, asks for a further ',
  ['judgement', true],
  ', which is what the follower can still do, so that the leader takes on only what is needed. Looking back, he could do it, because freezing under a shouting guest is normal for someone new. But I did not make that ',
  ['judgement', true],
  ' in the moment. I saw that he was struggling and took over the whole complaint, because I knew I could handle it. EI shows that I noticed how he felt, while situational leadership shows that I skipped judging what he could still do before I acted.',
]));
c.push(zh(['第 2 列顯示了另一個缺口。透過 EI 的 lens，我確實讀懂了同事，因為我看到他僵住了，便介入保護他。情境領導（Hersey & Blanchard, 1969, as cited in Latemore, 2024）是我這一列原本使用的概念，它要求更進一步的判斷，也就是追隨者還能做到什麼，讓領導者只承擔必要的部分。回頭看，他其實做得到，因為新人在大吼的客人面前僵住是很正常的。但我當下沒有做這個判斷。我看到他很吃力，就把整個客訴接過來，因為我知道自己處理得來。EI 顯示我注意到了他的感受，而情境領導顯示，我在行動前跳過了判斷他還能做到什麼。']));

// ---------- P5 ----------
c.push(label('第 5 段'));
c.push(para([
  'Even with these concepts added, one thing stays unexplained, which is why a problem becomes mine in the first place. Wong and Law\'s (2002) model places EI in the person and context in the emotional demands of the job, and both weekly concepts assume someone is already leading. None of them has a place for where a person stands among others, and in my rows that is what made a problem mine. ',
  ['Fleeson (2001) offers a way to see this, finding that a person\'s behaviour varies across situations at least as much as people differ from one another, and that this variation follows situational cues. In my rows, my position is one of those cues.', true],
  ' In Rows 2, 5 and 7, I was not the supervisor, but I was the senior or more experienced person present. In Row 7, my position is also what kept my anger down, rather than an ability to regulate it, because the anger came out as soon as the journalist left. ',
  ['Gross and John (2003) describe this as suppression, which hides an emotion\'s expression while the emotion itself may "linger and accumulate unresolved" (p. 349).', true],
  ' With my partner in Row 6, after a draining meeting, there was no role holding me, and my tone slipped. So my position decides what becomes my problem, and what that problem contains decides whether I use EI. My analysis therefore has to follow what that position does to the people around me.',
]));
c.push(zh([
  '即使加上這些概念，還是有一件事沒有被解釋，那就是一開始問題為什麼會變成我的。Wong 和 Law（2002）的模型把 EI 放在個人身上，把情境放在工作的情緒要求上，而兩個每週概念都假設已經有人在領導。它們都沒有考慮一個人在群體中所處的位置，而在我的各列裡，正是這個位置讓問題變成了我的。',
  ['Fleeson（2001）提供了一個理解的角度。他發現，一個人在不同情境之間的行為差異，至少和人與人之間的差異一樣大，而且這些變化會跟著情境線索走。在我的各列裡，我的位置就是其中一個線索。', true],
  '在第 2、5、7 列，我不是主管，但我是在場較資深或較有經驗的人。在第 7 列，壓住我怒氣的也是我的位置，而不是調節情緒的能力，因為記者一離開，怒氣就出來了。',
  ['Gross 和 John（2003）把這稱為壓抑，這種方式只藏住了情緒的表達，情緒本身則可能「持續存在、累積而無法化解」（p. 349）。', true],
  '在第 6 列，開完一場耗盡心力的會議之後，面對伴侶時沒有任何角色撐著我，我的語氣就變了。所以，我的位置決定了什麼會變成我的問題，而那個問題包含了什麼，決定了我會不會運用 EI。因此，我的分析接下來必須追蹤這個位置對我身邊的人造成了什麼影響。',
]));

// ---------- Change list ----------
c.push(label('修改表'));
c.push(table([900, 4600, 3526], [
  ['段', '修改', '原因'],
  ['2', '改寫 Wong & Law 對 EI 的定義那句', '原句把 Wong & Law 自己的四個維度誤寫成 Mayer & Salovey 的定義（原文 p. 246）'],
  ['2', '刪掉 even when nothing was urgent，以及 not how urgent the situation was, but', '決定不談急迫，避免和第 1 段打架'],
  ['2', '新增 I listened to my colleagues\' reasons…', '解決第 1、2 段對 Row 1 的矛盾，也支撐優先事項的論點'],
  ['3', 'Goleman 原句加引號和頁碼', '直接引用原文必須標示，已核對原文在 p. 95'],
  ['3', '段末新增一句', '把段落接回優先事項'],
  ['4', 'judgment 改成 judgement', '和 Appendix A 拼法一致'],
  ['5', '新增 Fleeson (2001) 兩句', '自己找的學術文獻（Credit 以上必要），說法已核對原文'],
  ['5', '新增 Gross & John (2003) 一句', '同上，引文在 p. 349'],
]));

c.push(label('她還要自己處理的事'));
c.push(num('Appendix A 第 5、6、7 列直接引用了 Goleman 的定義，也要加上引號和 (p. 95)，第 7 列的 states as 改成 defines as'));
c.push(num('References 加上 Fleeson (2001) 和 Gross & John (2003)，Mayer & Salovey (1997) 沒讀過原文的話寫成 as cited in Wong & Law, 2002'));
c.push(num('中文對照在交件前全部刪掉'));
c.push(num('目前約 1,050 字，最後要刪到約 700 字'));
c.push(num('報告最後的 AI 使用說明要寫出這一節有用到 AI 修改'));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: 'nums', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 340 } } } }] }] },
  sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: c }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(__dirname + '/HOSP7053_A2_Repertoire_修訂版.docx', b); console.log('written'); });
