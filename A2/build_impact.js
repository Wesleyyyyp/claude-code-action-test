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
c.push(new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: 'Impact on others 修訂版（英中對照）', font: FONT, size: 28, bold: true })] }));
c.push(new Paragraph({ spacing: { after: 240 }, children: [
  new TextRun({ text: '標示說明 ', font: FONT, size: 21, bold: true }),
  new TextRun({ text: '螢光筆', font: FONT, size: 21, highlight: 'yellow' }),
  new TextRun({ text: ' 是新寫或改寫的句子，沒有標的是保留她原稿的句子。這一版改用 service-profit chain 當主軸，每段追一條鏈並說明它在哪一環變弱。刪掉的內容列在最後的修改表。中文對照交件前要全部刪除。', font: FONT, size: 21 }),
] }));

// ---------- Opening ----------
c.push(label('開頭段'));
c.push(para([
  ['Read through the service-profit chain, my leadership enters at internal service quality, the support and treatment my colleagues receive, which shapes what they can deliver and in turn what guests receive. In every row where I stepped in, the guest in front of me received the service they needed, so judged by those guests alone, my leadership looks better than it is. The chain, however, runs on to guests I never meet, and the links that reach them are the ones my leadership tends to skip.', true],
]));
c.push(zh([
  ['用服務利潤鏈來讀，我的領導是從內部服務品質進入這條鏈的，也就是同事所得到的支持和對待，這會影響他們能提供什麼，進而影響客人得到什麼。在每一列我介入的情況裡，眼前的客人都得到了他們需要的服務，所以如果只用這些客人來評判，我的領導看起來比實際上好。然而，這條鏈會一路延伸到我從未見過的客人，而通往他們的那幾環，正是我的領導常常跳過的。', true],
]));

// ---------- R2 ----------
c.push(label('第 1 段（Row 2）'));
c.push(para([
  ['Row 2 shows that I can deliver service value to a guest while bypassing the link the chain depends on, which is employee capability. ', true],
  'I was not my colleagues\' trainer, but as the more experienced person on the desk, I often guided newer colleagues on the job. When my new colleague froze, I took the complaint over, so the guest\'s frustration ended sooner and the duty manager was likely spared having to apologise and decide on compensation. My colleague watched a full resolution, but lost the practice of handling the complaint himself, the very situation that had frozen him. Hogreve et al. (2022) place leadership and employee development among the internal practices that start the service-profit chain, noting that these shape what employees are able to do, and so the service guests receive. My takeover entered the chain there, ',
  ['but delivered the service through me rather than through him.', true],
  ' My evidence stops at my colleague, however, as I do not know how he later handled complaints without me.',
]));
c.push(zh([
  ['第 2 列顯示，我可以在繞過這條鏈所依賴的環節，也就是員工能力的情況下，為客人提供服務價值。', true],
  '我不是同事的培訓者，但身為櫃台上較有經驗的人，我常常在工作中帶領較資淺的同事。新同事僵住時，我把客訴接了過來，所以客人的不滿較快結束，值班經理也很可能因此不用親自道歉和決定補償。同事看到了完整的處理過程，卻失去了自己處理客訴的練習，而那正是讓他僵住的情境。Hogreve 等人（2022）把領導和員工培養列為啟動服務利潤鏈的內部做法，並指出它們會影響員工能做到什麼，進而影響客人得到的服務。我的接手就是從這裡進入這條鏈的，',
  ['但服務是經由我、而不是經由他提供的。', true],
  '不過我的證據只到同事為止，因為我不知道他之後在我不在時是怎麼處理客訴的。',
]));

// ---------- R4 ----------
c.push(label('第 2 段（Row 4）'));
c.push(para([
  ['Row 4 shows that customer satisfaction can rise while employee satisfaction falls, which the chain does not predict. ', true],
  'Here I was not the most senior person present, as the supervisor on duty was questioning my new colleague in front of a waiting guest, so ',
  ['what I stepped into was the questioning, not the check-in, because the questioning would not get the room ready any faster.', true],
  ' Since my colleague could already handle the check-in, she finished it herself once the questioning stopped. The guest, who had watched the questioning, wrote after check-out that although the room took a while, the staff were very proactive in helping. Behind that praise, being questioned in front of the guest had left my colleague overwhelmed, and when I checked on her afterwards, she told me she felt wronged. Hogreve et al. (2022) note that although the satisfaction mirror implies that employees\' feelings flow to customers, evidence finds employee satisfaction reaches customers mainly through the service delivered. So in Row 4, the guest\'s praise reflected the service my colleague delivered, not what she had paid to deliver it.',
]));
c.push(zh([
  ['第 4 列顯示，顧客滿意度可以在員工滿意度下降的同時上升，而這是服務利潤鏈沒有預測到的。', true],
  '這次我不是在場最資深的人，因為當班主管正在等候的客人面前質問我的新同事，所以',
  ['我介入的是那場質問，而不是入住手續，因為質問並不會讓房間更快準備好。', true],
  '因為同事本來就能處理入住，質問停了之後，她就自己把入住辦完了。目睹那場質問的客人，退房後寫道，雖然房間等了一陣子，但員工非常主動地幫忙。在那則好評背後，當著客人的面被質問，已經讓同事不堪負荷；事後我去關心她時，她告訴我她覺得很委屈。Hogreve 等人（2022）指出，雖然滿意度鏡像認為員工的感受會流向客人，但實證研究發現，員工滿意度主要是透過所提供的服務影響顧客。所以在第 4 列，客人的好評反映的是同事提供的服務，而不是她為了提供這份服務所付出的代價。',
]));

// ---------- R1 ----------
c.push(label('第 3 段（Row 1）'));
c.push(para([
  ['Row 1 shows the same link weakening over a longer period. ', true],
  'In Row 1, guests faced fewer parking disputes, but colleagues who did not fully agree had to ask questions they felt were intrusive, may not have felt safe to keep pushing back on me, and afterwards seemed reluctant to tell me what they thought. ',
  ['Their experience worsened while the guest outcome improved, because the improvement came from the procedure rather than from how my colleagues felt. This fits Hogreve et al.\'s (2022) conclusion that the chain\'s effects are not universal but depend on the conditions around them. A guest outcome can therefore improve while the employee link beneath it weakens, and no guest would notice.', true],
]));
c.push(zh([
  ['第 1 列顯示，同樣的環節在更長的時間裡變弱了。', true],
  '在第 1 列，客人的停車糾紛變少了，但不完全同意的同事得去問他們覺得冒犯的問題，可能也覺得不夠安心、不敢繼續反駁我，事後似乎也不太願意告訴我他們真正的想法。',
  ['他們的經驗變差了，客人的結果卻變好了，因為改善來自流程本身，而不是同事的感受。這符合 Hogreve 等人（2022）的結論，也就是服務利潤鏈的效果並非普遍成立，而是取決於周遭的條件。因此，客人的結果可以在底下的員工環節變弱的同時改善，而且沒有客人會察覺。', true],
]));

// ---------- Conclusion ----------
c.push(label('結尾段'));
c.push(para([
  ['The links I skip do not show in any single encounter, but the chain makes them depend on each other. In Rows 1, 2 and 4, the guest in front of me was served well, while my colleagues\' satisfaction or capability carried the cost. At this 73-room hotel, with two staff on each shift, there is often no one else to step in, so the desk depends on every staff member being able to handle a guest alone. What it needs from its senior staff is not someone who takes problems over, but someone who leaves colleagues able to handle them. So the guests my leadership affects most may be the ones I never meet. On the shifts I do not work, my colleagues serve them with how they feel and what they have practised, and that is what my leadership leaves behind.', true],
]));
c.push(zh([
  ['我跳過的那幾環，在任何一次接觸裡都看不出來，但服務利潤鏈讓它們彼此相依。在第 1、2、4 列，眼前的客人都被服務得很好，而同事的滿意度或能力卻承擔了代價。在這間有 73 間房、每班兩名員工的飯店，常常沒有其他人可以接手，所以櫃台依賴每一位員工都能獨自處理客人。它需要資深員工做的，不是把問題接過來，而是讓同事有能力自己處理。所以，我的領導影響最大的客人，可能是我從未見過的那些。在我沒有值班的班次，同事帶著他們的感受和練習過的東西去服務這些客人，而那就是我的領導所留下的。', true],
]));

// ---------- Change table ----------
c.push(label('修改表'));
c.push(table([1300, 4000, 3726], [
  ['段', '修改', '原因'],
  ['開頭', '整段重寫，改用 service-profit chain 框架，把「客人」改成「眼前的客人」', 'SPC 要當主軸；原句和 Appendix 的 R3、R5 矛盾'],
  ['Row 2', '首句改寫，點出被繞過的環節是 employee capability；刪掉「對他來說像前輩上了一課」和原本的末句', '用 SPC 的環節名稱分析；省字數'],
  ['Row 4', '首句改寫，點出 customer satisfaction 和 employee satisfaction 方向相反；修正「接手了質問」的語意；刪掉「我是想保護同事」和「主管的回饋沒有消失」', 'SPC 框架；語意更清楚；省字數'],
  ['Row 1', '從原本的跨列段落獨立出來，改寫成 SPC 反例', 'Row 1 是員工環節變弱、客人結果卻變好的最好證據'],
  ['原跨列段', 'Row 5 和 Row 7 刪除', 'R5 留給 Ethics；R7 不涉及 SPC 的員工到客人環節'],
  ['結尾', '整段重寫，加入 interdependence 和這間飯店的 leadership success factor，改寫原本那句不通順的末句', 'rubric HD 要求'],
]));

c.push(label('她還要處理的事'));
c.push(num('Hogreve 原文還沒收到。第 1、2、3 段的 Hogreve 說法請對照原文確認，並盡量附頁碼（Structure v2 標的是 p. 463 和 p. 469）'));
c.push(num('Appendix A 第 4 列第 2 欄補上半句，例如 who completed the guest\'s check-in herself，和正文一致'));
c.push(num('73 間房、每班兩人這兩個數字請再確認一次'));
c.push(num('目前約 700 字，目標 600，最後刪字數時可以優先刪第 2 段的情境描述'));
c.push(num('中文對照交件前全部刪除，AI 使用說明要寫出這一節有用 AI 改寫'));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: 'nums', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 340 } } } }] }] },
  sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: c }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(__dirname + '/HOSP7053_A2_Impact_修訂版.docx', b); console.log('written'); });
