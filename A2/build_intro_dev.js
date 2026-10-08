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
c.push(new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: 'Introduction 與 Development plan 正式版（英中對照）', font: FONT, size: 28, bold: true })] }));
c.push(new Paragraph({ spacing: { after: 240 }, children: [
  new TextRun({ text: '標示說明 ', font: FONT, size: 21, bold: true }),
  new TextRun({ text: '螢光筆', font: FONT, size: 21, highlight: 'yellow' }),
  new TextRun({ text: ' 是修改或新增的地方，其餘保留她原稿的句子。刪掉的內容列在最後的修改表。中文對照交件前要全部刪除。', font: FONT, size: 21 }),
] }));

// ---------- Introduction ----------
c.push(label('1. Introduction'));
c.push(para([
  'My claim is that when I am the senior person present, I handle problems myself and decide for the people involved, and that my position, more than their needs, drives this. ',
  ['The guests in front of me are served well', true],
  ', but my colleagues bear the costs, ',
  ['in how they feel and what they practise, out of the guests\' sight.', true],
  ' A reader could argue that handling problems myself ',
  ['at a hotel front desk', true],
  ' was simply good service, and in the moment it often was. But in Rows 2 and 4, a new colleague was struggling in front of a waiting guest, and I took over ',
  ['in Row 2, where I was the most experienced person present, but', true],
  ' left the check-in with my colleague in Row 4, where the supervisor was present. The need was ',
  ['similar', true],
  ', so what changed my response was my position. Read through emotional intelligence (EI; Wong & Law, 2002), my repertoire shows that ',
  ['once a problem is mine, I notice people\'s feelings only when it includes them, while my position, which EI cannot explain, decides which problems become mine. The impact section traces these costs to guests', true],
  ', the ethics section shows that in Row 5 I also decided alone what guests were told, and the development plan sets out how I will leave more of each problem ',
  ['with others.', true],
]));
c.push(zh([
  '我的主張是，當我是在場資深的人時，我會自己處理問題，並替相關的人做決定；驅動這個行為的，主要是我的位置，而不是他們的需要。',
  ['眼前的客人得到了很好的服務', true],
  '，但同事承擔了代價，',
  ['包括他們的感受，以及他們能練習到什麼，而這些都在客人看不到的地方。', true],
  '讀者可能會認為，',
  ['在飯店櫃台', true],
  '自己處理問題本來就是好的服務；而在當下，它也常常是。但在第 2 列和第 4 列，都是新同事在等候的客人面前陷入困難。',
  ['在第 2 列，我是在場最有經驗的人，我把客訴接了過來；在第 4 列，主管在場，', true],
  '我則把入住留給同事自己完成。兩次的需要',
  ['相似', true],
  '，所以改變我做法的，是我的位置。用情緒智力（EI; Wong & Law, 2002）來看，我的 Repertoire 顯示，',
  ['一旦問題成為我的問題，只有在問題本身包含別人的感受時，我才會注意到他們的感受；而決定哪些問題會成為我的問題的，是 EI 無法解釋的位置。Impact 一節追蹤這些代價如何傳到客人', true],
  '；Ethics 一節顯示，在第 5 列，客人被告知什麼，也是由我一個人決定的；Development plan 則說明我會如何把每個問題更多地',
  ['留給別人處理。', true],
]));

// ---------- Development plan ----------
c.push(label('5. Development plan'));
c.push(label('開頭'));
c.push(para([
  'My three sections point to one area to develop. Stepping in is not the problem, but once I step in as the senior person present, I decide alone what others do, feel and are told. I no longer work at the hotel, but this habit follows my position rather than the workplace, so in my group assignments, part-time job and volunteer role, wherever I am the more experienced person present, I will put each change into practice and check whether it has worked.',
]));
c.push(zh([
  '我的三節分析都指向同一個需要發展的地方。問題不在於我會介入，而在於一旦我以在場資深者的身分介入，別人要做什麼、感受到什麼、被告知什麼，都由我一個人決定。我已經不在那間飯店工作了，但這個習慣跟著的是我的位置，而不是工作場所。所以在小組作業、打工和志工活動中，只要我是在場比較有經驗的人，我就會實際使用下面每一個改變，並確認它有沒有效。',
]));
c.push(label('第一項'));
c.push(para([
  'First, when a teammate falls behind or struggles with a task, I will ask which part they want me to take and support them on the rest, judging what they can still do, as situational leadership asks (Hersey & Blanchard, 1969, as cited in Latemore, 2024)',
  [', which I skipped in Row 2', true],
  '. It has worked if they finish their own part and handle the next similar task without asking me to take it.',
]));
c.push(zh([
  '第一，當隊友進度落後或在某個任務上卡住時，我會問他希望我接手哪個部分，其餘部分則從旁協助。這就是情境領導要求的，判斷對方還能做什麼（Hersey & Blanchard, 1969，引自 Latemore, 2024）',
  ['，而這正是我在第 2 列跳過的', true],
  '。如果他能完成自己的部分，下次遇到類似的任務也不需要請我接手，就代表這個改變有效。',
]));
c.push(label('第二項'));
c.push(para([
  'Second, when a co-worker brings me an urgent problem at my part-time job, I will include how they feel in the problem before asking for facts, by first telling them the priority and then asking what happened',
  [', as I failed to do in Row 5', true],
  '. This builds the half of empathy my lens left out, ',
  ['"skill in treating people according to their emotional reactions" (Goleman, 1998, p. 95)', true],
  '. It has worked if they can explain what happened without becoming more upset.',
]));
c.push(zh([
  '第二，當打工的同事帶著緊急問題來找我時，我會在詢問事實之前，先把他的感受納入問題裡，',
  ['也就是', true],
  '先告訴他現在的優先事項，再問發生了什麼事',
  ['，而這是我在第 5 列沒做到的', true],
  '。這補上了我的 lens 漏掉的同理心另一半，',
  ['也就是「依照別人的情緒反應來對待他們的技巧」（Goleman, 1998, p. 95）', true],
  '。如果他能把經過說清楚，而且情緒沒有變得更激動，就代表這個改變有效。',
]));
c.push(label('第三項'));
c.push(para([
  'Third, when a mistake in a group or at work affects people who are not present, I will give them a true reason, even a brief one, and ask before speaking for anyone else',
  [', unlike in Row 5', true],
  '. This keeps to Northouse\'s ',
  ['(2026, p. 465)', true],
  ' view of honesty as not misrepresenting reality. It has worked if I could repeat every explanation I gave in front of everyone involved without correcting it.',
]));
c.push(zh([
  '第三，當小組或工作中的錯誤影響到不在場的人時，我會給他們一個真實的理由，即使只是簡短的理由；在代替任何人發言之前，我會先問過對方',
  ['，不再像第 5 列那樣', true],
  '。這符合 Northouse',
  ['（2026, p. 465）', true],
  '對誠實的看法，也就是不扭曲事實。如果我說過的每一個解釋，都能當著所有相關的人再說一次而不需要更正，就代表這個改變有效。',
]));
c.push(label('第四項'));
c.push(para([
  'Fourth, because my regulation held mainly when my role held it (Rows 6 and 7), I will treat noticing that I have stopped listening as my cue to say I need a moment, practising this in my group meetings this semester. It has worked if I take that moment before my tone turns sharp.',
]));
c.push(zh([
  '第四，因為我的情緒調節主要是在角色撐著我的時候才守得住（第 6、7 列），我會把「發現自己沒在聽對方說話」當作信號，說出我需要一點時間，並在這學期的小組會議中練習。如果我能在語氣變尖銳之前先停下來，就代表這個改變有效。',
]));
c.push(label('結尾'));
c.push(para([
  'To track all four, I will keep a short note after each meeting or shift of who finished the problem, who was told what and whether I paused in time, and review it monthly. If the notes show others finishing more of the problems I step into, I will know that I am leading by handing the problem back rather than taking it on.',
]));
c.push(zh([
  '為了追蹤這四個改變，我會在每次會議或每一班結束後簡短記錄，包括問題最後是誰完成的、誰被告知了什麼，以及我有沒有及時停下來，並每個月回顧一次。如果紀錄顯示，在我介入的問題中，由別人完成的越來越多，我就知道自己是在把問題交還給別人，而不是把它攬到自己身上。',
]));

// ---------- Change table ----------
c.push(label('修改表'));
c.push(table([1700, 4000, 3326], [
  ['位置', '修改', '原因'],
  ['Intro', 'Guests 改成 The guests in front of me；刪掉 as a result', '和 Impact、Ethics 的「眼前的客人」一致，也避免和 Row 5 房內客人矛盾'],
  ['Intro', '刪掉 Six of my seven rows come from the front desk of a hotel in Taiwan，改成 at a hotel front desk', '省字數；情境已在 Appendix'],
  ['Intro', 'The need was the same 改成 similar', 'Row 2 和 Row 4 的情境不完全相同，說成一樣容易被反駁'],
  ['Intro', '縮短 Repertoire 預告和結尾的路線圖', '從 253 字縮到約 220 字'],
  ['Dev 第一至三項', '各補上一個列號', '老師 email 要求每個說法都連到某一列'],
  ['Dev 第二項', 'Goleman 加引號和 p. 95', '這句和原文一字不差，屬於直接引用'],
  ['Dev 第三項', 'Northouse 加上 p. 465', '已核對原文'],
]));

c.push(label('她還要處理的事'));
c.push(num('Intro 約 220 字、Development plan 約 425 字，最後和其他三節一起刪字數，全篇要在 2,500 字以內'));
c.push(num('全篇合併後，再確認 Intro 預告的內容和各節實際寫的一致'));
c.push(num('中文對照交件前全部刪除，AI 使用說明要寫出這兩節有用 AI 修改'));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: 'nums', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 340 } } } }] }] },
  sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: c }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(__dirname + '/HOSP7053_A2_Intro_Development_正式版.docx', b); console.log('written'); });
