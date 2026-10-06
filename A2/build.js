const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, LevelFormat, Footer, PageNumber,
} = require('docx');

const FONT = { ascii: 'Calibri', hAnsi: 'Calibri', eastAsia: 'Microsoft JhengHei', cs: 'Calibri' };
const CONTENT_W = 9026; // A4 width minus 1" margins
const ACCENT = '4B2E83';

// **bold** inline markup
function runs(text, opts = {}) {
  return text.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(s =>
    s.startsWith('**')
      ? new TextRun({ text: s.slice(2, -2), bold: true, font: FONT, ...opts })
      : new TextRun({ text: s, font: FONT, ...opts }));
}
const p = (t, o = {}) => new Paragraph({ children: runs(t), spacing: { after: 120 }, ...o });
const h1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs(t), pageBreakBefore: false });
const h2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs(t) });
const h3 = t => new Paragraph({ heading: HeadingLevel.HEADING_3, children: runs(t) });
const bullet = (t, level = 0) => new Paragraph({ numbering: { reference: 'bullets', level }, children: runs(t), spacing: { after: 60 } });
const num = (t, ref = 'nums') => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(t), spacing: { after: 60 } });
const quote = t => new Paragraph({
  children: runs(t, { italics: true, color: '333333' }),
  indent: { left: 567 }, spacing: { before: 60, after: 120 },
  border: { left: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 8 } },
});
const note = t => new Paragraph({
  children: runs(t, { color: '7A1F1F' }), spacing: { before: 60, after: 120 },
  shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'FBEFEF' },
});

function table(widths, rows) {
  const total = widths.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({
      tableHeader: i === 0,
      children: r.map((c, j) => new TableCell({
        width: { size: widths[j], type: WidthType.DXA },
        borders,
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        shading: i === 0 ? { type: ShadingType.CLEAR, color: 'auto', fill: 'E8E1F3' } : undefined,
        children: String(c).split('\n').map(line => new Paragraph({ children: runs(line, i === 0 ? { bold: true } : {}) })),
      })),
    })),
  });
}
const gap = () => new Paragraph({ children: [], spacing: { after: 60 } });

const c = [];

// ---------- Title ----------
c.push(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 60 },
  children: [new TextRun({ text: 'HOSP7053 Assessment 2', font: FONT, size: 36, bold: true, color: ACCENT })] }));
c.push(new Paragraph({ spacing: { after: 200 },
  children: [new TextRun({ text: '正文寫作指引：架構、順序與各節方向', font: FONT, size: 28, bold: true })] }));
c.push(p('**截止：** 10/9（五）5:00 pm，Turnitin（Learn.UQ）。正文 2,500 字；Appendix A、參考文獻、AI 說明不計字數。'));
c.push(p('**依據：** Rubric、Guideline、W8 和 W9 投影片、Week 10 workshop transcript、Structure v2、Appendix A（edited）。'));
c.push(note('英文句子只是示範方向，請一律改成自己的說法。matrix 裡的事實以 Appendix A 為準，正文不要加入 Appendix 沒有的事件。'));

// ---------- 1 ----------
c.push(h1('一、整體邏輯與寫作順序'));
c.push(p('**順序：Repertoire → Impact → Ethics → Development plan → Introduction**'));
c.push(p('每一節都要用到前一節的結論，所以照這個順序寫，整篇就是一條線往下推：'));
c.push(table([2200, 6826], [
  ['節', '這一節回答的問題'],
  ['Repertoire', '我在什麼條件下會失效？→ 扛起問題的時候'],
  ['Impact', '失效時，誰付出代價？→ 站在我旁邊、最資淺的人'],
  ['Ethics', '最嚴重的一次：保護身邊的人，變成了欺瞞（R5）'],
  ['Development plan', '在「扛起問題那一刻」要改什麼'],
  ['Introduction', '最後寫：前面都寫完了，才知道整篇真正證明了什麼'],
]));
c.push(gap());
c.push(p('**銜接：** 每一節的最後一句，要把讀者帶到下一節。例如 Repertoire 結尾說「我要改的是扛起問題那一刻的做法」，Impact 開頭就接著問「那一刻，誰付出了代價？」。'));

c.push(h2('老師在 workshop 說的寫法（每一節都適用）'));
c.push(bullet('**每一段：** 第一句是這段的主張 → 中間是證據（matrix 的列＋理論）→ 最後一句說明證據怎麼支持主張。**只讀第一句和最後一句，也要讀得通。**'));
c.push(bullet('**每一節：** 開頭一小段說明這節的主張 → 大約 3 段證據段落 → 結尾一小段，寫出「我從中學到什麼」。'));
c.push(bullet('**不要重述情境：** 寫 "In Row 5…" 加一句這代表什麼，閱卷者會自己翻 Appendix。'));
c.push(bullet('**用理論的詞描述自己的行為：** 不要寫「他還不太會」，要寫「我判斷他的 readiness…」。'));
c.push(bullet('**一定要批判理論：** 寫出理論解釋得了什麼，也寫出它解釋不了什麼。'));
c.push(bullet('**Dimension 欄不評分；** matrix 不一定要是工作情境（R6 可以用）。'));

c.push(h2('Appendix A 列號對照'));
c.push(table([900, 5726, 2400], [
  ['列', '情境', 'Dimension'],
  ['R1', '提出停車流程，逐一回應同事的反對', 'Leading others'],
  ['R2', '新同事在客人面前僵住，她接手客訴', 'Leading others'],
  ['R3', '大廳的長住客，聽主管的話但持續留意', 'Ethical'],
  ['R4', '轉移主管在客人面前對新同事的質問', 'Leading others'],
  ['R5', '同一位新同事發錯房卡；她用「升等」向客人解釋', 'Ethical'],
  ['R6', '開完長會議後對伴侶回話很衝', 'Leading self'],
  ['R7', '記者逼問名人房客資料', 'Leading self'],
]));
c.push(p('建議在 matrix 加一欄列號（Row 1–7），正文才能寫 "(Appendix A, Row 2)"。只是排版，不動內容。', { spacing: { before: 120, after: 120 } }));

// ---------- 2 Repertoire ----------
c.push(h1('二、Your repertoire（約 700 字）'));
c.push(p('**整篇主張（v2）：** My emotional intelligence works when I am not the one carrying the problem. When I take a problem on myself, as I tend to do as the most senior person, I stop reading the people beside me and hold my own feelings in instead of managing them.'));
c.push(p('**這一節的主張：** 我的 EI 不是固定的水準，而是取決於我在情境裡的位置；而且察覺到了，不等於運用了察覺。'));

c.push(h2('對照要求的檢查結果'));
c.push(table([4300, 2000, 2726], [
  ['要求', '出處', 'v2 狀態'],
  ['只用一個 overarching lens（EI）', 'W9 第 24 張、transcript', '✅'],
  ['和每週的概念比較，哪裡一致、各自漏掉什麼', 'W9 第 24 張', '✅ 第 3 段（R2）'],
  ['指出 lens 解釋不了什麼', 'W9、rubric HD', '✅ 結尾段'],
  ['自己找的文獻，做到指定讀物做不到的事', 'W9、rubric', '✅ Fleeson、Gross & John'],
  ['文獻互相比較，而不是並列', 'rubric HD', '✅'],
  ['說明切換條件', 'rubric', '✅ 旁觀 vs 扛著；有角色 vs 沒有'],
  ['各列涵蓋三個 dimension', 'rubric HD', '✅'],
  ['**列出常用的做法，指出哪一個用得太多**', 'W9 第 24 張第 1 點', '❌ 要補'],
  ['**分出習慣性和偶爾使用**', 'rubric Distinction', '❌ 要補'],
]));
c.push(gap());

c.push(h2('段落安排'));
c.push(table([1500, 3926, 1800, 1800], [
  ['段落（字數）', '內容', '列', '文獻'],
  ['開頭（~130）', '為什麼用 EI。**補上做法清單**：自己接手（R2、R5、R7，習慣性，也是用得太多的那一個）、用論證推動（R1）、替別人擋下來（R4）、服從但持續留意（R3，偶爾）。用 EI 來讀，看得出是什麼在切換：在旁邊看時能力都在，自己扛起問題時就失效。', '只點列號', 'Wong & Law；Goleman'],
  ['第 1 段（~170）', '讀人的能力（others\' emotion appraisal，6.8）取決於任務在誰身上。', '**R4 對 R5**（拿掉 R3）', 'Wong & Law p. 270；Fleeson'],
  ['第 2 段（~170）', '自我調節靠角色撐著，而且是壓抑，不是調節。', '**R7 對 R6**', 'Goleman；Gross & John p. 349；Wong & Law p. 271'],
  ['第 3 段（~150）', '察覺到了，卻沒有理解或運用。對照 situational leadership 和 EI 各自看到什麼。', '**R2**', 'Goleman；Latemore pp. 266–267'],
  ['結尾（~80）', 'EI 解釋不了「什麼讓能力失效」。Wong & Law 發現 EI 的效果取決於工作，她的各列顯示能力本身取決於角色。帶到 Impact。', '—', 'Wong & Law'],
]));
c.push(gap());
c.push(p('目前草稿約 730 字，加上做法清單後，第 1、2 段各要刪 20–30 字。'));

c.push(h2('讓段落不亂的寫法'));
c.push(p('不需要只挑一兩個場景。老師在 workshop 說的是「一個主張，底下 2–3 個重點，每個重點用 matrix 的例子」，而且最有力的證據本身就是對照（R4 對 R5、R7 對 R6）。看起來亂，通常是因為重述了情境，不是因為用了太多列。'));
c.push(bullet('**每段一組對照，最多兩列，** 每列最多一句話。'));
c.push(bullet('**第一句寫規律，不寫事件。**'));
c.push(quote('✗ In Row 4, I noticed my colleague was overwhelmed…'));
c.push(quote('✓ My reading of others depends on whether I am carrying the task.'));
c.push(bullet('**兩列放在同一句或相鄰的句子，** 讓對照一目了然：'));
c.push(quote('In Row 4 the pressure sat with the supervisor and I read everyone; in Row 5, with the same colleague, the pressure was mine and I read no one.'));
c.push(bullet('**每段的最後一句回到同一組關鍵詞**（例如 carrying the problem），讀者就會感覺整節在講同一件事。'));

c.push(h2('可以加強的地方'));
c.push(num('**第 3 段改用 W8 的 EI 四環節詞彙**（Perceive → Understand → Use → Regulate；Salovey & Mayer, 1990）。W8 投影片指出 WLEIS **沒有任何題目測量 Understand**。R2 她察覺到同事僵住（Perceive），卻判斷錯了原因（Understand），斷掉的正是 WLEIS 量不到的那一環。若引用 Salovey & Mayer，要讀過原文或用 secondary citation。', 're1'));
c.push(num('**R6 可以用 W8 情境 1 的說法：** "you cannot regulate a state you have not noticed"。R6 斷掉的是察覺自己的狀態；她的提示（發現自己沒在聽）正好是投影片的修正做法 "name your own state"。', 're1'));
c.push(num('**R7 加入 Gross & John 的另一面：** 原文說 "there are times when suppression is the best or even the only option… there may not be time to cognitively reevaluate a rapidly developing situation." 被記者錄影時，壓抑可能是對的，問題在於壓抑是她唯一的工具。這樣分析更公平，也符合 W9 要的「不只一種解讀」。', 're1'));
c.push(num('**用詞一致：** Appendix R2 寫的是 readiness（ability and willingness），Latemore 和 v2 正文寫的是 development level。正文第一次出現時交代兩者指的是同一件事。', 're1'));

// ---------- 3 Impact ----------
c.push(h1('三、Impact on others（約 600 字）'));
c.push(p('**這一節的主張（v2）：** My leadership tends to benefit the guest at the cost of the person standing next to me.'));
c.push(p('老師在 workshop 建議的切入角度：**你做決定時，比較會考慮誰、比較不會考慮誰？** 她的主張正好在回答這個問題。'));
c.push(table([1500, 4826, 900, 1800], [
  ['段落（字數）', '內容', '列', '重點'],
  ['開頭（~80）', '跨列來看誰總是贏、誰總是付代價：客人和當下的狀況是贏家；R1 的同事、R2 和 R5 的新同事付出代價。', '只點列號', 'W9 第 25 張第 3 點'],
  ['1. 鏈結成立（~150）', '同事被保護 → 客人順利入住 → 客人評論寫 "very proactive"。說明這條鏈為什麼成立：員工被保護和客人受益發生在同一時間、同一地點。', 'R4', '唯一有客人書面證據的鏈'],
  ['2. 鏈結比較弱（~130）', '客人當下被照顧好，但同事沒練到。**明確寫出證據停在哪裡**：「這一列的證據只到同事；之後的客人有沒有受影響，是我的推論。」', 'R2', 'W8："a short, honest trace is the better piece of work"'],
  ['3. 鏈結斷了，結果卻變好（~150）', '同事的體驗變差，客人的停車糾紛反而變少，和 service-profit chain 的預測相反。', 'R1', 'Hogreve et al. (2022) p. 463、p. 469'],
  ['結尾（~90）', '這間飯店（73 間房、每班兩人）需要資深員工做到「讓別人能自己處理」，而不是替別人處理。用這個需求回看，R2 的代價比表面上大。', '—', 'leadership success factors'],
]));
c.push(gap());
c.push(bullet('**每個人分開寫**（the new colleague、the supervisor、the waiting guest），不要寫成 "the team"。'));
c.push(bullet('R5 留到 Ethics 再談，這裡最多在開頭段點一下列號。'));
c.push(bullet('Hogreve 原文沒有給我，**引用的頁碼和句子請自己核對。**'));

// ---------- 4 Ethics ----------
c.push(h1('四、Ethical responsibilities（約 600 字）'));
c.push(p('**這一節的主張（v2）：** 在 R5，我對抵達的客人說了一個沒必要的謊，也在沒有權衡的情況下，決定不讓房內的客人知道。兩者都是為了保護身邊的人。'));
c.push(p('**Preston 的五個步驟融入段落裡寫，** 不要寫成五個小標題，否則 600 字會變成填表格。'));
c.push(table([1700, 4626, 1300, 1400], [
  ['段落（字數）', '內容', 'Preston 步驟', '原文'],
  ['1. 情境（~120）', 'R5 其實是**兩個決定**：對抵達客人說的話，和對房內客人的沉默。列出當時實際可選的做法：照實說、升等並說明原因、編理由升等、先問主管。', '評估情況', '—'],
  ['2. 說謊（~150）', '用定義和 middle axiom 檢驗：誠實的版本存在而且無害，所以這個謊站不住。**逐一檢查五個原則哪些做到、哪些沒做到。** 另外，她說的是「主管幫您升等了」，**借用了主管的名義**，主管也被牽連。', '對照價值', '定義 p. 76；舉證責任 p. 74；middle axiom p. 83'],
  ['3. 隱瞞：真正有爭議的部分（~170）', '房內客人的房間曾經暴露，但他不知道。reversibility test 也得不到答案：如果是她的房間，她也不確定想不想知道。**重點是她當時沒有權衡過這件事。**', '衡量因素、做出決定', 'p. 80（改寫，不直接引用）；p. 83'],
  ['4. 沒看到的風險＋跨列模式（~160）', '自我欺騙會從對同事的忠誠中長出來，她當時沒意識到。往外看：R4、R7 也是在替身邊的人擋，這是她的模式還是偶發事件？**至少寫出兩種解讀，誠實比較。**', '考慮品格', 'p. 78'],
]));
c.push(gap());
c.push(bullet('rubric 問的是 "Was the ethical question genuinely open?"。說謊的答案很明確，所以**重心放在第 3 段的隱瞞**，第 2 段可以短一點。'));
c.push(bullet('**五個原則（respect、service、justice、honesty、community）一定要附出處**，只寫 "from Week 9" 不符合 APA。'));
c.push(bullet('**Preston 的五步驟不在第 5 章**，請確認出自書中哪一頁；找不到的話引用 W9 投影片。'));
c.push(bullet('第 4 段是 HD 和 Distinction 的分界：W9 說 "stronger reports step back and ask what it reveals about your values across your rows."'));

// ---------- 5 Development ----------
c.push(h1('五、Development plan（約 400 字）'));
c.push(p('**開頭一句（~40 字）：** 前三節的發現都指向同一個時刻，也就是我扛起問題的那一刻，所以這些改變都針對那個時刻。'));
c.push(table([500, 2400, 1900, 2226, 2000], [
  ['#', '改變', '情境', '怎麼知道有效', '出處'],
  ['1', '**跨列的總原則：** 說出 "I\'ve got this" 之前，先問「哪部分你現在做不了，需要我來？」', '同事卡住、我準備接手時', '對方能留在現場，處理其中一部分', 'Repertoire 第 3 段＋Impact（R2、飯店規模）'],
  ['2', '先告訴同事優先順序，再問發生了什麼；向客人補救時照實說明原因', '同事慌亂、需要補救失誤時', '同事說得清楚、沒有更慌；我不需要編理由', 'Repertoire 第 1 段＋Ethics（R5）'],
  ['3', '先追問，再提出自己的看法', '向資淺同事提出改變時', '有人直接跟我說反對意見', 'Impact（R1）'],
  ['4', '察覺自己沒在聽時，先說「我需要一點時間」', '長時間的壓力或會議之後', '語氣變尖之前就先暫停', 'Repertoire 第 2 段（R6、R7）'],
]));
c.push(gap());
c.push(bullet('每一項約 80–90 字：做什麼、在什麼情境、怎麼知道有效，再加一句出自哪一節。'));
c.push(bullet('**第 1 項最重要：** 它超越了單一列，是讀完整個 matrix 才得出的結論。W9 要求 "the plan should go further than any single row"。'));
c.push(bullet('「怎麼知道有效」寫**別人看得到的改變**（同事、伴侶的反應），不是自己的感覺。這就是 rubric 說的 "answer to the people it affects"。'));

// ---------- 6 Intro ----------
c.push(h1('六、Introduction（約 200 字，最後寫）'));
c.push(table([1500, 6026, 1500], [
  ['句子', '內容', '字數'],
  ['1', '背景：台灣飯店的櫃台、她的角色、Appendix A 共七列', '~30'],
  ['2–3', '**整篇的主張**（v2 的那句，其他節寫完後再修）', '~50'],
  ['4–6', '後面各節分別怎麼支持這個主張，每節一句', '~90'],
  ['7', '一句話帶出 Repertoire', '~30'],
]));
c.push(gap());
c.push(bullet('**檢驗方法：** 讀過她 matrix 的人，能不能合理地反駁這個主張？不能的話，那只是一個標籤。'));
c.push(bullet('不要寫得太像 W9 投影片上的範例主張（"I protect the people in front of me by taking things on myself…"）。v2 的主張是從 EI 出發的，保持這樣就好。'));

// ---------- 7 Citations ----------
c.push(h1('七、引用核對結果'));
c.push(p('以下是拿 Structure v2 的說法對照原文的結果。'));
c.push(table([4500, 3026, 1500], [
  ['v2 的說法', '原文', '結果'],
  ['WLEIS 讀人題目沒有指定情境（p. 270）', 'items 5–8 在 p. 270', '✅'],
  ['4 題調節題有 3 題問「控制」（p. 271）', 'items 13、14、16 是 control；15 是 calm down', '✅'],
  ['Wong & Law 發現 EI 的效果取決於工作', 'emotional labour 的調節作用', '✅'],
  ['Fleeson：和自己的差異至少和人與人之間一樣大；平均值穩定；變化跟著情境線索', '都有', '✅'],
  ['Gross & John："linger and accumulate unresolved"（p. 349）', '有', '✅'],
  ['Latemore：風格配合 development level（pp. 266–267）', '有', '✅'],
  ['Preston：p. 76 定義、p. 74 舉證責任、p. 83 middle axiom、p. 78 loyalty', '都有，頁碼正確', '✅'],
  ['Preston p. 80：withholding 不等於 lying', '是 Preston 引用別人（Fried，經由 Davis）', '⚠️ 改寫成 Preston 自己的論點'],
  ['Wong & Law "built the WLEIS on" Gross 的模型', '原文是用 Gross 的模型詮釋 ROE、UOE', '⚠️ 改成 "drew on"'],
  ['Goleman p. 98、p. 100；Hogreve p. 463、p. 469', '沒有原文', '❓ 請自己核對'],
  ['Preston 五步驟框架', '不在第 5 章', '❓ 確認頁碼或引用投影片'],
  ['Wong & Law 卷期', '講義寫 13(2)，v2 寫 13(3)', '❓ 應為 13(3)，到資料庫確認'],
]));

// ---------- 8 Appendix ----------
c.push(h1('八、Appendix A 還剩的問題'));
c.push(bullet('**References：** 全形冒號「：」要改；刪掉 Hersey & Blanchard (1969)、加上 Latemore (2024)；最後和正文合併成一份清單，放在正文後、Appendix 前。'));
c.push(bullet('**R2 用詞：** Appendix 寫 readiness，正文寫 development level，正文第一次出現時要交代。'));
c.push(bullet('**R4 第 3 欄：** "that way" 沒有交代。W9 說各列保留原紀錄，所以可以不改，但讀者會看不懂。'));

// ---------- 9 Checklist ----------
c.push(h1('九、交件前檢查'));
c.push(num('**只讀每段的第一句和最後一句**，看整篇是否通順', 'chk'));
c.push(num('**正文的每個說法，Appendix 裡都有對應的列**', 'chk'));
c.push(num('**正文 ≤ 2,500 字**（不含 Appendix、參考文獻、AI 說明）', 'chk'));
c.push(num('**參考文獻**一份清單（APA 7），期刊名和書名斜體，直接引用都有頁碼', 'chk'));
c.push(num('**AI 使用說明**放在最後，例如：I used Claude to help structure the report, check my use of sources against the original articles, and revise wording.', 'chk'));

c.push(h2('時程建議'));
c.push(table([2500, 6526], [
  ['日期', '工作'],
  ['10/6（二）', '寫完 Repertoire'],
  ['10/7（三）', 'Impact＋Ethics'],
  ['10/8（四）', 'Development plan＋Introduction＋整理參考文獻'],
  ['10/9（五）上午', '最後檢查，**中午前交**，避開 5 點前 Turnitin 塞車'],
]));

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: ACCENT }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 25, bold: true, color: '222222' }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 23, bold: true }, paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 567, hanging: 283 } } } }] },
      ...['nums', 're1', 'chk'].map(ref => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 340 } } } }] })),
    ],
  },
  sections: [{
    properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: '888888' })] })] }) },
    children: c,
  }],
});

Packer.toBuffer(doc).then(b => { fs.writeFileSync('HOSP7053_A2_寫作指引.docx', b); console.log('written'); });
