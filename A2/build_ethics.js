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
c.push(new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: 'Your ethical responsibilities 修訂版（英中對照）', font: FONT, size: 28, bold: true })] }));
c.push(new Paragraph({ spacing: { after: 240 }, children: [
  new TextRun({ text: '標示說明 ', font: FONT, size: 21, bold: true }),
  new TextRun({ text: '螢光筆', font: FONT, size: 21, highlight: 'yellow' }),
  new TextRun({ text: ' 是修改或新增的地方，其餘保留她原稿的句子。刪掉的內容列在最後的修改表。Northouse 和 Preston 的說法與頁碼都已對照原文核對。中文對照交件前要全部刪除。', font: FONT, size: 21 }),
] }));

// P1
c.push(label('第 1 段'));
c.push(para([
  'Row 5 carries the clearest ethical weight in my rows. To stop an arriving guest from reaching an occupied room, I told them the supervisor had just upgraded them, without asking her and without mentioning the mistake. Working through Preston\'s (2001, as cited in Latemore, 2024',
  [', p. 284', true],
  ') five steps shows a decision that served the guest but rested on an invented reason, and the same pattern runs through how I usually lead as an ethical leader.',
]));
c.push(zh([
  '第 5 列是我所有列中倫理份量最重的一列。為了不讓剛抵達的客人走進一間已經有人住的房間，我告訴他主管剛幫他升等了，但我沒有先問主管，也沒有提到發錯房卡的事。用 Preston（2001，引自 Latemore, 2024',
  ['，p. 284', true],
  '）的五個步驟來分析，這個決定服務了客人，卻建立在一個編出來的理由上；而同樣的模式，也貫穿了我平常身為倫理領導者的做法。',
]));

// P2
c.push(label('第 2 段'));
c.push(para([
  'Assessing the situation shows that honesty with the guest pulled against protecting their stay, my colleague and the hotel\'s image, and that I had more options than I used. As senior staff could decide on upgrades, I could have offered the larger room and simply said their room had a problem. Instead, I credited the upgrade to the supervisor, because an upgrade from me would raise the question of why I had not given it at check-in. Against Northouse\'s (2026) principles, the upgrade was just, as it compensated the guest for our mistake, and it served them. Honesty is where it falls short. Northouse notes that honesty does not mean disclosing everything, but it does mean not misrepresenting reality',
  [' (pp. 465, 466)', true],
  '. So the ethical problem was not the upgrade or what I left unsaid, but the reason I invented when an honest one was available. ',
  ['Whether to tell either guest about the mistake was genuinely open, since telling them could have worried them about something that never reached them. Inventing a reason was not.', true],
]));
c.push(zh([
  '評估當時的情境會發現，對客人誠實，跟保護他的住宿體驗、我的同事以及飯店的形象彼此衝突，而且我其實有比實際用到更多的選項。資深員工本來就可以自行決定升等，所以我大可以直接提供較大的房間，只說他原本的房間有問題就好。但我把升等說成是主管給的，因為如果說是我升等的，客人會問為什麼入住時不直接升等。對照 Northouse（2026）的原則，升等是公正的，因為它補償了我們的錯誤給客人帶來的影響，也服務了客人。問題出在誠實。Northouse 指出，誠實不代表要揭露一切，但它的意思是不扭曲事實',
  ['（pp. 465, 466）', true],
  '。所以倫理上的問題不在升等，也不在我沒說出口的部分，而在於明明有誠實的理由可以用，我卻編了一個。',
  ['要不要告訴任何一位客人發錯卡的事，是一個真正開放的問題，因為告訴他們，可能會讓他們為一件根本沒有發生在他們身上的事擔心。但編造理由並不是。', true],
]));

// P3
c.push(label('第 3 段'));
c.push(para([
  'Weighing these factors, I put the guest\'s stay and the hotel\'s image ahead of honesty, because I believed the truth would only make us look unprofessional. Preston\'s final steps ask whether I can account for a decision and what it says about the kind of person I ought to be, and this reason is hard to account for, since as the guest I would not want to learn the upgrade story was made up. The interests served were the arriving guest\'s and the hotel\'s, while those quietly traded away belonged to people with no say. My colleague still faced the consequences described in the previous section, the supervisor\'s name was used before she was asked, and the guest in the occupied room was never told that a card for their room had been issued. ',
  ['At the time, I recorded that neither guest was affected (Appendix A, Row 5), and that judgement shows the risk I did not see, which was that I alone decided what each person needed to know.', true],
  ' So the decision served the guest in front of me, but its costs were carried by people who had no say in it.',
]));
c.push(zh([
  '權衡這些因素後，我把客人的住宿體驗和飯店的形象放在誠實之前，因為我認為說出真相只會讓我們顯得不專業。Preston 的最後幾步會問，我能不能為這個決定說明理由，以及這個決定說明了我應該成為什麼樣的人。這個理由很難說得過去，因為如果我是那位客人，我不會希望事後才知道升等的說法是編出來的。被顧到的是抵達的客人和飯店的利益，而被悄悄犧牲的，是那些沒有發言權的人。我的同事仍然承受了前一節所說的後果，主管的名字在她被問之前就被拿來用了，房間已有人住的那位客人，也一直不知道他房間的房卡曾經被發給別人。',
  ['我當時在紀錄裡寫下兩位客人都沒有受到影響（Appendix A，第 5 列），而這個判斷正顯示了我當時沒看到的風險，也就是每個人需要知道什麼，都是我一個人決定的。', true],
  '所以這個決定服務了眼前的客人，但它的代價，是由那些沒有參與決定的人承擔的。',
]));

// P4
c.push(label('第 4 段'));
c.push(para([
  'Row 5 is not a one-off lapse but a sharper version of how I usually lead as an ethical leader. ',
  ['Read through Northouse\'s (2026) principles, my rows show that I am strongest on service.', true],
  ' In Rows 2, 4, 5 and 7, I stepped in to protect the person in front of me. I am weaker on respect, which Northouse describes as ',
  ['treating other people\'s decisions and values with respect rather than as a means to my own ends (p. 462)', true],
  '. In Row 2, I took the complaint over without asking my colleague which part he could handle, and in Row 1, although I listened to every objection, I answered each one until the group agreed to my procedure. My honesty is uneven too. In Rows 5 and 7, I was open with my supervisor about the mistake, the upgrade and my need to calm down, yet in Row 5 I decided what the guests could be told.',
]));
c.push(zh([
  '第 5 列並不是一次性的失誤，而是我平常身為倫理領導者的做法中，比較尖銳的一個版本。',
  ['用 Northouse（2026）的原則來讀，我的各列顯示我最強的是服務。', true],
  '在第 2、4、5、7 列，我都挺身保護眼前的人。我比較弱的是尊重。Northouse 把尊重描述為',
  ['尊重別人的決定和價值，而不是把他們當成達成我自己目的的手段（p. 462）', true],
  '。在第 2 列，我沒問同事他能處理哪個部分，就把客訴接了過來；在第 1 列，雖然我聽了每一個反對意見，但我一一回應，直到大家同意我的流程。我的誠實也不一致。在第 5、7 列，我對主管很坦白，包括錯誤、升等，以及我需要時間冷靜；但在第 5 列，卻是我決定客人可以被告知什麼。',
]));

// P5
c.push(label('第 5 段'));
c.push(para([
  'Two readings follow. The generous one is right that ',
  ['no guest', true],
  ' was harmed, each decision stayed within my authority and, as Northouse (2026) allows, not everything must be disclosed. But it explains only this decision, while the stricter reading explains the pattern, that I am honest with those above me but decide on behalf of those around me. Northouse warns that lying implies we do not trust others ',
  ['to deal with the information we have and that we know what is best for them (p. 466)', true],
  ', and deciding for others carries the same distrust. The exception is Row 4, where someone more senior was present and I left the check-in with my colleague, which suggests that, like taking the problem on, deciding for others follows my position. So as an ethical leader, I put a smooth result for the people in front of me first, and when I am the senior person present and that result meets honesty or others\' right to decide, I have let the smooth result win.',
]));
c.push(zh([
  '由此可以有兩種解讀。寬容的解讀說對了幾件事，',
  ['沒有客人', true],
  '受到傷害，每個決定都在我的權限之內，而且就像 Northouse（2026）所說的，不是每件事都必須揭露。但它只解釋得了這一個決定；比較嚴格的解讀才解釋得了整個模式，也就是我對上面的人誠實，卻替身邊的人做決定。Northouse 提醒，說謊意味著我們不相信別人',
  ['能處理我們掌握的資訊，也意味著我們認為自己知道什麼對他們最好（p. 466）', true],
  '，而替別人做決定，也帶著同樣的不信任。例外是第 4 列，當時有更資深的人在場，我把入住留給同事自己完成。這顯示「替別人做決定」跟「把問題攬到自己身上」一樣，都取決於我的位置。所以身為倫理領導者，我把眼前的人能得到順利的結果放在第一位；而當我是在場資深的人，這個結果又遇上誠實或別人自己做決定的權利時，我讓順利的結果勝出了。',
]));

// Change table
c.push(label('修改表'));
c.push(table([1200, 4300, 3526], [
  ['段', '修改', '原因'],
  ['1', '刪掉 because the question was genuinely open；Preston 加上 p. 284', '開放的部分移到第 2 段寫清楚；已核對 Latemore p. 284 的 Table 9.3'],
  ['2', 'Northouse 加上頁碼；段末新增兩句，說明哪一部分是真正開放的問題', '回應 rubric 的 genuinely open；已核對 Northouse pp. 465, 466'],
  ['3', '刪掉 because I judged them unaffected，改成引用 Appendix 的原話來寫當時沒看到的風險', '照老師 email，不符合論點的列是有用的材料'],
  ['4', '修正懸垂修飾語；respect 的轉述改得更貼近原文並加頁碼', '原句文法錯誤；原文 p. 462 是 treat other people\'s decisions and values with respect'],
  ['5', 'no one 改成 no guest；說謊那句補上 we know what is best 並加頁碼', '和第 3 段同事的代價一致；已核對 Northouse p. 466'],
]));

c.push(label('她還要處理的事'));
c.push(num('第 3 段 the consequences described in the previous section 要搭配含 Row 5 的 Impact 版本，用不含 Row 5 的版本就要改成直接寫出同事的代價'));
c.push(num('第 2、3 段兩個當時的想法（說是主管升等的原因，以及認為說真話會顯得不專業）要符合事實'));
c.push(num('References 加上 Northouse, P. G. (2026). Leadership, theory and practice (10th ed.). SAGE. 正確格式見對話'));
c.push(num('目前約 770 字，目標 600，第 4、5 段連到平常模式的部分盡量保留'));
c.push(num('中文對照交件前全部刪除，AI 使用說明要寫出這一節有用 AI 修改'));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: 'nums', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 340 } } } }] }] },
  sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: c }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(__dirname + '/HOSP7053_A2_Ethics_修訂版.docx', b); console.log('written'); });
