const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, PageBreak,
  convertMillimetersToTwip,
} = require('docx');
const fs = require('fs');

const MIN = 'ＭＳ 明朝';
const GO  = 'ＭＳ ゴシック';
const W = 9600; // content width (dxa)

const NB = { top:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, bottom:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
             left:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, right:{style:BorderStyle.NONE,size:0,color:'FFFFFF'} };
const SB = (sz=4) => ({ top:{style:BorderStyle.SINGLE,size:sz,color:'000000'}, bottom:{style:BorderStyle.SINGLE,size:sz,color:'000000'},
             left:{style:BorderStyle.SINGLE,size:sz,color:'000000'}, right:{style:BorderStyle.SINGLE,size:sz,color:'000000'} });

function run(t, o={}) {
  return new TextRun({ text: t, font: { name: o.font || MIN, hint: 'eastAsia' }, size: o.size || 21, bold: !!o.bold });
}
function p(t, o={}) {
  return new Paragraph({
    children: (Array.isArray(t) ? t : [run(t, o)]),
    alignment: o.align,
    spacing: { before: o.before || 0, after: o.after === undefined ? 40 : o.after, line: o.line || 264 },
    indent: o.indent,
  });
}
// 問題文（ぶら下げインデント）
function q(t, o={}) {
  return p(t, Object.assign({ indent: { left: 420, hanging: 420 }, after: 40 }, o));
}
// 選択肢
function ch(t) { return p(t, { indent: { left: 700, hanging: 280 }, after: 20 }); }

function cell(children, o={}) {
  return new TableCell({
    children: (Array.isArray(children) ? children : [children]),
    width: { size: o.width, type: WidthType.DXA },
    columnSpan: o.span,
    shading: o.shade ? { type: ShadingType.CLEAR, fill: o.shade, color: 'auto' } : undefined,
    verticalAlign: 'center',
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
  });
}
function tcell(text, o={}) {
  return cell(p(text, { align: o.align || AlignmentType.CENTER, after: 0, font: o.font, bold: o.bold, size: o.size }), o);
}

// 資料ボックス（1セルの囲み）
function box(children) {
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [W],
    borders: SB(6),
    rows: [ new TableRow({ children: [ cell(children, { width: W }) ] }) ],
  });
}
function dataTable(colWidths, rows) {
  return new Table({
    width: { size: colWidths.reduce((a,b)=>a+b,0), type: WidthType.DXA },
    columnWidths: colWidths,
    borders: SB(4),
    rows: rows,
  });
}
const gap = (n=1) => Array.from({length:n}, () => p('', { after: 0, size: 12 }));

// ============================ 本文 ============================
const body = [];

body.push(p('4　江戸時代　（26点：問題19〜31）', { font: GO, bold: true, size: 24, after: 100 }));
body.push(p('次の会話文と資料G〜Qを見て、あとの問いに答えなさい。', { after: 100 }));

// ---- 会話文 ----
body.push(box([
  p('ゆうま：江戸幕府は、江戸・大阪・京都などの重要な都市や、主な鉱山を直接支配していたんだね。', { after: 20 }),
  p('あおい：それに、将軍から1万石以上の領地をあたえられた大名は、その領地と人々を独自に支配したけれど、幕府の定めた法によって厳しく統制されていたよ。', { after: 20 }),
  p('ゆうま：つまり、ⓐ幕府と、大名が支配する藩とが、全国の土地と人々を支配していたということだね。', { after: 20 }),
  p('あおい：そうだね。でも、ⓑ大名の中には、領地を取り上げられてしまった家もあるよ。ⓒ江戸と領地を往復する制度も、大名にとっては重い負担だったみたい。', { after: 0 }),
]));
body.push(...gap());

// ---- 資料G・H ----
body.push(p('資料G　大名の分類と数（1664年）　　　　　　　　資料H　主な大名の石高（万石）（1664年）', { font: GO, size: 20, after: 60 }));

const gTable = dataTable([1160, 780, 780, 780], [
  new TableRow({ children: [ tcell('', {width:1160, shade:'EFEFEF'}), tcell('親藩',{width:780,font:GO,shade:'EFEFEF'}), tcell('譜代',{width:780,font:GO,shade:'EFEFEF'}), tcell('外様',{width:780,font:GO,shade:'EFEFEF'}) ]}),
  new TableRow({ children: [ tcell('50万石以上',{width:1160}), tcell('2',{width:780}), tcell('0',{width:780}), tcell('4',{width:780}) ]}),
  new TableRow({ children: [ tcell('20万石以上',{width:1160}), tcell('6',{width:780}), tcell('1',{width:780}), tcell('10',{width:780}) ]}),
  new TableRow({ children: [ tcell('10万石以上',{width:1160}), tcell('4',{width:780}), tcell('14',{width:780}), tcell('9',{width:780}) ]}),
  new TableRow({ children: [ tcell('10万石未満',{width:1160}), tcell('1',{width:780}), tcell('97',{width:780}), tcell('77',{width:780}) ]}),
  new TableRow({ children: [ tcell('合　計',{width:1160,font:GO}), tcell('13',{width:780,font:GO}), tcell('112',{width:780,font:GO}), tcell('100',{width:780,font:GO}) ]}),
]);
const hRows = [
  ['徳川将軍家（幕府）', '680'],
  ['金沢の前田家（外様）', '103'],
  ['鹿児島の島津家（外様）', '73'],
  ['名古屋の徳川家（親藩）', '62'],
  ['仙台の伊達家（外様）', '56'],
  ['和歌山の徳川家（親藩）', '56'],
  ['熊本の細川家（外様）', '54'],
  ['福井の松平家（親藩）', '45'],
  ['福岡の黒田家（外様）', '43'],
  ['彦根の井伊家（譜代）', '30'],
];
const hTable = dataTable([2900, 900], hRows.map(r => new TableRow({ children: [
  tcell(r[0], { width: 2900, align: AlignmentType.LEFT, size: 19 }), tcell(r[1], { width: 900, size: 19 }) ]})));

body.push(new Table({
  width: { size: W, type: WidthType.DXA },
  columnWidths: [4200, 5400],
  borders: NB,
  rows: [ new TableRow({ children: [
    new TableCell({ width:{size:4200,type:WidthType.DXA}, borders: NB, margins:{top:0,bottom:0,left:0,right:200},
      children: [ gTable, p('（集英社「日本の歴史12」などをもとに作成）', { size: 17, after: 0 }) ] }),
    new TableCell({ width:{size:5400,type:WidthType.DXA}, borders: NB, margins:{top:0,bottom:0,left:0,right:0},
      children: [ hTable ] }),
  ]}) ],
}));
body.push(...gap());

body.push(q('問19　会話文中の下線部ⓐのような、幕府と藩が全国の土地と人々を支配したしくみを何といいますか。　〔記述・2点〕'));
body.push(...gap());
body.push(q('問20　資料G・Hから読み取れることとして適切でないものを、次のア〜エから1つ選び、記号で書きなさい。　〔選択・2点〕【思判表】'));
body.push(ch('ア　50万石以上の大名では、関ヶ原の戦いのころから徳川氏に従うようになった大名が最も多い。'));
body.push(ch('イ　古くから徳川氏に従ってきた大名は、10万石未満の大名が最も多い。'));
body.push(ch('ウ　大名全体の数を分類ごとに比べると、徳川氏の一族である大名が最も多い。'));
body.push(ch('エ　石高の大きい外様大名5家の石高を合計しても、徳川将軍家の石高には及ばない。'));
body.push(...gap());

// ---- 資料I ----
body.push(p('資料I　武家諸法度（一部）と大名の処分の例', { font: GO, size: 20, after: 60 }));
body.push(box([
  p('一、城を修理するときは、必ず幕府に届け出ること。城を新しくつくることは禁止する。', { after: 20 }),
  p('一、大名は、領地と江戸に交代で住むこと。　　　　　（「徳川禁令考」より、わかりやすく直したもの）', { after: 60 }),
  p('〈大名の処分の例〉', { font: GO, size: 20, after: 20 }),
  p('・小早川家（岡山　50万石）…小早川秀秋は、関ヶ原の戦いで徳川方を勝利に導いた人物だが、21歳で亡くなったとき、あと継ぎがいなかった。　→　領地取り上げ', { after: 20 }),
  p('・福島家（広島　50万石）…福島正則は、関ヶ原の戦いで徳川方として活躍した大名だが、城の修理を幕府の許可なく行った。　→　領地取り上げ', { after: 0 }),
]));
body.push(...gap());
body.push(q('問21　下線部ⓑについて、福島家が領地を取り上げられたのはなぜか。資料Iを参考にして、最も適するものを次のア〜エから1つ選び、記号で書きなさい。　〔選択・2点〕【思判表】'));
body.push(ch('ア　武家諸法度で、城は修理するのではなく、新しくつくり直すように定められていたから。'));
body.push(ch('イ　武家諸法度で、あと継ぎのいない大名は領地を返すように定められていたから。'));
body.push(ch('ウ　大名が城の修理に多くの費用を使うと、参勤交代の費用が不足してしまうと幕府が心配したから。'));
body.push(ch('エ　大名が幕府の許可なく城の守りを固めることは幕府にとって危険であり、武家諸法度で禁じられていたから。'));
body.push(...gap());

// ---- 資料J ----
body.push(p('資料J　参勤交代に関する二つの資料', { font: GO, size: 20, after: 60 }));
body.push(box([
  p('〈ある藩の支出の割合（1655年）〉', { font: GO, size: 20, after: 20 }),
  dataTable([2600, 2400, 2400], [
    new TableRow({ children: [ tcell('国元での費用など', {width:2600, size:19}), tcell('参勤交代の費用', {width:2400, size:19}), tcell('江戸屋敷の費用', {width:2400, size:19}) ]}),
    new TableRow({ children: [ tcell('52％', {width:2600, size:19}), tcell('20％', {width:2400, size:19}), tcell('28％', {width:2400, size:19}) ]}),
  ]),
  p('', { size: 12, after: 0 }),
  p('〈薩摩藩の参勤交代にかかった費用（1720年）〉', { font: GO, size: 20, after: 20 }),
  p('道中費　約4億1144万円　／　船賃　約5171万円　／　計　約4億6315万円', { after: 20 }),
  p('江戸に滞在する費用などを含めると…　約14億1750万円　　（現在の金額に置きかえたもの）', { after: 0 }),
]));
body.push(...gap());
body.push(q('問22　下線部ⓒの制度について、資料Jから読み取れる、この制度が大名（藩）にあたえた影響を、制度名を明らかにして簡潔に書きなさい。　〔記述・2点〕【思判表】'));
body.push(...gap(2));

// ---- 資料K ----
body.push(p('資料K　17世紀初めの海外との貿易', { font: GO, size: 20, after: 60 }));
body.push(box([
  p('　徳川家康は貿易の発展に努め、西日本の大名や、京都・堺・長崎の商人に、海外へ渡ることを許す証書をあたえた。この証書を持つ船による（　X　）がさかんに行われ、船はルソン・シャム（タイ）・安南（ベトナム）などにおもむいた。やがて東南アジア各地に日本人が移り住むようになり、アユタヤなどには（　Y　）とよばれる日本人の居住地がつくられた。', { after: 0 }),
]));
body.push(...gap());
body.push(q('問23　資料K中のX・Yに適する語句の組み合わせとして適するものを、次のア〜エから1つ選び、記号で書きなさい。　〔選択・2点〕'));
body.push(ch('ア　X…南蛮貿易　　　Y…日本町　　　　イ　X…南蛮貿易　　　Y…唐人屋敷'));
body.push(ch('ウ　X…朱印船貿易　　Y…日本町　　　　エ　X…朱印船貿易　　Y…唐人屋敷'));
body.push(...gap());
body.push(q('問24　次のア〜エは、幕府が禁教と貿易統制を進めていく過程で起こったできごとである。年代の古い順に並べたときの順序として正しいものを、あとの1〜4から1つ選び、番号で答えなさい。　〔整序・選択・2点〕【思判表】'));
body.push(ch('ア　幕府が、幕領にキリスト教の禁教令を出す。'));
body.push(ch('イ　島原・天草一揆が起こる。'));
body.push(ch('ウ　ポルトガル船の来航が禁止される。'));
body.push(ch('エ　平戸のオランダ商館が、長崎の出島に移される。'));
body.push(ch('1　ア→イ→ウ→エ　　2　ア→ウ→イ→エ　　3　イ→ア→ウ→エ　　4　ア→イ→エ→ウ'));
body.push(...gap());

// ---- 資料L ----
body.push(p('資料L　江戸幕府の交易相手と「四つの窓口」', { font: GO, size: 20, after: 60 }));
body.push(dataTable([2400, 3400, 3800], [
  new TableRow({ children: [ tcell('窓口', {width:2400, font:GO, shade:'EFEFEF'}), tcell('相手', {width:3400, font:GO, shade:'EFEFEF'}), tcell('おもな内容', {width:3800, font:GO, shade:'EFEFEF'}) ]}),
  new TableRow({ children: [ tcell('長崎（幕府の直轄）', {width:2400, size:19}), tcell('中国（清）・（　あ　）', {width:3400, size:19}), tcell('幕府が貿易を独占して統制した', {width:3800, size:19, align:AlignmentType.LEFT}) ]}),
  new TableRow({ children: [ tcell('対馬藩', {width:2400, size:19}), tcell('朝鮮', {width:3400, size:19}), tcell('国交が回復し、使節が来日した', {width:3800, size:19, align:AlignmentType.LEFT}) ]}),
  new TableRow({ children: [ tcell('薩摩藩', {width:2400, size:19}), tcell('琉球王国', {width:3400, size:19}), tcell('17世紀初めに征服され、支配を受けた', {width:3800, size:19, align:AlignmentType.LEFT}) ]}),
  new TableRow({ children: [ tcell('松前藩', {width:2400, size:19}), tcell('蝦夷地（アイヌ民族）', {width:3400, size:19}), tcell('幕府から交易の権利をあたえられた', {width:3800, size:19, align:AlignmentType.LEFT}) ]}),
]));
body.push(...gap());
body.push(q('問25　資料Lに関連して、「鎖国」下の日本の対外関係について述べた次のX〜Zの正誤の組み合わせとして適するものを、あとのア〜エから1つ選び、記号で書きなさい。　〔選択・2点〕【思判表】'));
body.push(ch('X　資料L中の（　あ　）にあたる国はオランダで、1641年に平戸から出島へ商館が移され、商館長は海外の情報を記した風説書を幕府に提出した。'));
body.push(ch('Y　朝鮮とは国交が回復し、将軍の代がわりごとなどに、祝賀の使節である朝鮮通信使が日本に派遣された。'));
body.push(ch('Z　蝦夷地では、松前藩が不利な条件の交易を強いたため、コシャマインを指導者としてアイヌの人々が戦いを起こした。'));
body.push(ch('ア　X…正　Y…正　Z…誤　　イ　X…正　Y…誤　Z…正'));
body.push(ch('ウ　X…誤　Y…正　Z…正　　エ　X…誤　Y…誤　Z…正'));
body.push(...gap());

// ---- 資料M ----
body.push(p('資料M　耕地面積の推移と、農業の進歩', { font: GO, size: 20, after: 60 }));
body.push(box([
  dataTable([3100, 2000, 2000, 2000], [
    new TableRow({ children: [ tcell('', {width:3100, shade:'EFEFEF'}), tcell('江戸時代の直前', {width:2000, size:19, shade:'EFEFEF'}), tcell('江戸時代の中ごろ', {width:2000, size:19, shade:'EFEFEF'}), tcell('明治時代の初め', {width:2000, size:19, shade:'EFEFEF'}) ]}),
    new TableRow({ children: [ tcell('耕地面積（万町歩）', {width:3100, size:19}), tcell('約164', {width:2000, size:19}), tcell('約297', {width:2000, size:19}), tcell('約305', {width:2000, size:19}) ]}),
  ]),
  p('', { size: 12, after: 0 }),
  p('　幕府や藩は、用水路をつくったり、海や沼を干拓したりして、耕地を広げることに力を入れた。また、備中ぐわ・千歯こきなどの農具や、干鰯などの肥料が広まり、綿・菜種・藍などの商品作物の栽培も各地に広がった。', { after: 0 }),
]));
body.push(...gap());
body.push(q('問26　資料Mから読み取れる江戸時代の農業の変化を、「新田」の語句を使って簡潔に書きなさい。　〔記述・2点〕【思判表】'));
body.push(...gap(2));

// ---- 資料N ----
body.push(p('資料N　江戸時代の政治についてまとめたカード', { font: GO, size: 20, after: 60 }));
body.push(dataTable([700, 8900], [
  new TableRow({ children: [ tcell('A', {width:700, font:GO}), tcell('極端な動物愛護を定めた命令を出した。財政が悪化したため、質を落とした貨幣を発行した。儒学のうち、主従関係を重んじる朱子学を広く学ばせた。', {width:8900, size:19, align:AlignmentType.LEFT}) ]}),
  new TableRow({ children: [ tcell('B', {width:700, font:GO}), tcell('新田開発を進め、裁判の基準となる法をまとめた。目安箱を設置した。また、キリスト教に関係のない、漢文に訳されたヨーロッパの書物の輸入制限をゆるめた。', {width:8900, size:19, align:AlignmentType.LEFT}) ]}),
  new TableRow({ children: [ tcell('C', {width:700, font:GO}), tcell('商工業者の同業者組合の結成を認め、営業税を納めさせた。長崎からの銅や俵物の輸出を活発にし、印旛沼の干拓にも取り組んだ。', {width:8900, size:19, align:AlignmentType.LEFT}) ]}),
  new TableRow({ children: [ tcell('D', {width:700, font:GO}), tcell('江戸に出かせぎに来ていた者を村に帰した。ききんに備えて村ごとに米を蓄えさせ、旗本・御家人の借金を帳消しにした。', {width:8900, size:19, align:AlignmentType.LEFT}) ]}),
]));
body.push(...gap());
body.push(q('問27　資料NのCで結成が認められた、商工業者の同業者組合を何といいますか。　〔記述・2点〕'));
body.push(...gap());
body.push(q('問28　資料NのA〜Dを、年代の古い順に並べたときの順序として正しいものを、次の1〜4から1つ選び、番号で答えなさい。　〔整序・選択・2点〕【思判表】'));
body.push(ch('1　B→A→C→D　　2　A→B→D→C　　3　A→B→C→D　　4　B→A→D→C'));
body.push(...gap());

// ---- 資料O ----
body.push(p('資料O　当時よまれた歌と、改革を行った人物の考え', { font: GO, size: 20, after: 60 }));
body.push(box([
  p('〈狂歌〉　白河の　清きに魚の　すみかねて　もとの濁りの　田沼こひしき', { after: 40 }),
  p('※「白河」は白河藩主であった松平定信を、「田沼」は田沼意次を指している。', { size: 18, after: 60 }),
  p('〈松平定信の自叙伝「宇下人言」〉', { font: GO, size: 20, after: 20 }),
  p('　老中（田沼意次）にわいろをおくって役職を得た者が多い。わいろは公然とわたされ、そのせいで政治が乱れている。だからわいろをやめさせることを同僚と相談し、今のように厳しくなったのである。', { after: 0 }),
]));
body.push(...gap());
body.push(q('問29　資料Oの狂歌が表している内容として最も適するものを、次のア〜エから1つ選び、記号で書きなさい。　〔選択・2点〕【思判表】'));
body.push(ch('ア　田沼意次の政治はわいろが横行して住みにくかったので、厳しくても松平定信の改革の方がよい。'));
body.push(ch('イ　松平定信の改革は厳しすぎて住みにくいので、わいろが横行していても田沼意次の政治の方がなつかしい。'));
body.push(ch('ウ　田沼意次の政治は厳しすぎて住みにくかったので、わいろが横行していても松平定信の改革の方がなつかしい。'));
body.push(ch('エ　松平定信の改革はわいろが横行して住みにくいので、厳しくても田沼意次の政治の方がよい。'));
body.push(...gap());

// ---- 資料P ----
body.push(p('資料P　江戸時代の二つの文化', { font: GO, size: 20, after: 60 }));
body.push(dataTable([2200, 7400], [
  new TableRow({ children: [ tcell('元禄文化', {width:2200, font:GO}), tcell('井原西鶴の浮世草子、近松門左衛門の人形浄瑠璃の脚本、松尾芭蕉の俳諧、尾形光琳の装飾画　など', {width:7400, size:19, align:AlignmentType.LEFT}) ]}),
  new TableRow({ children: [ tcell('化政文化', {width:2200, font:GO}), tcell('十返舎一九の小説、葛飾北斎の「富嶽三十六景」、歌川（安藤）広重の「東海道五十三次」　など', {width:7400, size:19, align:AlignmentType.LEFT}) ]}),
]));
body.push(...gap());
body.push(q('問30　資料Pの二つの文化について述べた文として適切でないものを、次のア〜エから1つ選び、記号で書きなさい。　〔選択・2点〕'));
body.push(ch('ア　元禄文化は、経済力をつけた上方（大阪・京都）の町人を担い手として栄えた文化である。'));
body.push(ch('イ　松尾芭蕉は「おくのほそ道」を著した人物で、化政文化を代表する俳諧の作者である。'));
body.push(ch('ウ　近松門左衛門は、義理と人情の板ばさみに悩む人々を題材とした脚本を書いた。'));
body.push(ch('エ　化政文化のころには多色刷りの版画の技術が広まり、風景をえがいた錦絵が安く売られた。'));
body.push(...gap());

// ---- 資料Q ----
body.push(p('資料Q　江戸時代後期の学問', { font: GO, size: 20, after: 60 }));
body.push(box([
  p('　前野良沢や杉田玄白らは、オランダ語で書かれた人体解剖書を苦心して翻訳し、1774年に「解体新書」として出版した。このころから、オランダ語を通してヨーロッパの学問や文化を研究する学問がさかんになり、その後、伊能忠敬による正確な日本地図の作成などにもつながっていった。', { after: 0 }),
]));
body.push(...gap());
body.push(q('問31　資料Qのような学問がさかんになった背景には、資料NのBの人物が行ったある政策がある。その政策の内容を、簡潔に書きなさい。　〔記述・2点〕【思判表】'));
body.push(...gap(2));

// ============================ 解答用紙 ============================
body.push(new Paragraph({ children: [ new PageBreak() ] }));
body.push(p('解　答　用　紙', { font: GO, bold: true, size: 28, align: AlignmentType.CENTER, after: 160 }));
body.push(dataTable([1400, 1400, 1400, 5400], [
  new TableRow({ children: [
    tcell('組', {width:1400, font:GO}), tcell('　　　組', {width:1400}),
    tcell('番号', {width:1400, font:GO}), tcell('　　　番　　名前：', {width:5400, align:AlignmentType.LEFT}) ]}),
]));
body.push(...gap());

const ansRows = [];
ansRows.push(new TableRow({ children: [
  tcell('番号', {width:1000, font:GO, shade:'EFEFEF'}), tcell('解　　答', {width:3800, font:GO, shade:'EFEFEF'}),
  tcell('番号', {width:1000, font:GO, shade:'EFEFEF'}), tcell('解　　答', {width:3800, font:GO, shade:'EFEFEF'}) ]}));
for (let i = 1; i <= 16; i++) {
  const right = i + 16;
  ansRows.push(new TableRow({ children: [
    tcell(String(i), {width:1000}),
    cell(p('', {after:0}), {width:3800}),
    tcell(right <= 31 ? String(right) : '', {width:1000}),
    cell(p('', {after:0}), {width:3800}),
  ]}));
}
body.push(dataTable([1000, 3800, 1000, 3800], ansRows));
body.push(...gap());
body.push(p('単元別得点　　地形図　　／10　　中世　　／14　　安土桃山　　／12　　江戸　　／26', { font: GO, size: 20, after: 60 }));
body.push(p('合　計　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　／62点', { font: GO, size: 20, after: 0 }));

// ============================ 教員用解答 ============================
body.push(new Paragraph({ children: [ new PageBreak() ] }));
body.push(p('【教員用】解答例', { font: GO, bold: true, size: 26, after: 120 }));

const answers = [
  ['1','三角点'], ['2','縮尺'], ['3','等高線'], ['4','イ（国道）'], ['5','2（X－正　Y－誤）'],
  ['6','十字軍'], ['7','ルネサンス'], ['8','宗教改革'], ['9','羅針盤'], ['10','バスコ・ダ・ガマ'],
  ['11','銀'], ['12','イ'], ['13','織田信長'], ['14','楽市・楽座（楽市楽座）'],
  ['15','（例）座による営業の独占という特権を失い、誰でも自由に商売ができるようになった。'],
  ['16','刀狩（刀狩令）'], ['17','太閤検地'], ['18','2（X－正　Y－誤）'],
  ['19','幕藩体制'],
  ['20','ウ'],
  ['21','エ'],
  ['22','（例）参勤交代の費用と江戸屋敷の費用が藩の支出の約半分を占め、大名（藩）の財政を圧迫した。'],
  ['23','ウ'],
  ['24','1（ア→イ→ウ→エ）'],
  ['25','ア（X－正　Y－正　Z－誤）'],
  ['26','（例）新田開発が進められ、江戸時代の初めから中ごろにかけて耕地面積が大きく増えた。'],
  ['27','株仲間'],
  ['28','3（A→B→C→D）'],
  ['29','イ'],
  ['30','イ'],
  ['31','（例）キリスト教に関係のない、漢文に訳されたヨーロッパの書物（漢訳洋書）の輸入制限をゆるめたこと。'],
];
body.push(dataTable([900, 8700], [
  new TableRow({ children: [ tcell('番号', {width:900, font:GO, shade:'EFEFEF'}), tcell('解　答　例', {width:8700, font:GO, shade:'EFEFEF'}) ]}),
  ...answers.map(a => new TableRow({ children: [
    tcell(a[0], {width:900, size:19}), tcell(a[1], {width:8700, size:19, align:AlignmentType.LEFT}) ]})),
]));
body.push(p('※15・22・26・31は記述式のため、上記は解答例です。生徒の解答が趣旨に合っていれば正答としてください。', { size: 19, before: 80 }));

// ---- 出題のねらい ----
body.push(new Paragraph({ children: [ new PageBreak() ] }));
body.push(p('【教員用】出題のねらい（江戸時代　問19〜31）', { font: GO, bold: true, size: 26, after: 100 }));
body.push(p('　「歴史　基本用語チェックテスト」で確認した基本用語に到達できるよう設計しています。解答そのものは既習の基本用語ですが、資料の読み取り・複数資料の関連づけ・年代の整序・正誤の判断を経なければ答えにたどり着けない形にしてあります。', { after: 100 }));
const aim = [
  ['19','記述','幕藩体制','会話文と資料G・Hから、幕府と藩による全国支配を読み取らせる。'],
  ['20','選択【思判表】','（幕藩体制の理解）','分類別の大名数と石高の2資料を照合し、誤った読み取りを判断させる。'],
  ['21','選択【思判表】','武家諸法度','法令の条文と処分事例を結びつけ、改易の理由を推論させる。'],
  ['22','記述【思判表】','参勤交代','支出の割合と実際の費用の2資料から、藩財政への影響を説明させる。'],
  ['23','選択','朱印船貿易／日本町','南蛮貿易・唐人屋敷との識別を求める組み合わせ選択。'],
  ['24','整序【思判表】','島原・天草一揆／鎖国／出島','禁教から「鎖国」完成までの流れを年代順に整理させる。'],
  ['25','選択【思判表】','出島／朝鮮通信使／シャクシャイン','コシャマイン（室町時代）との混同を判断させる正誤問題。'],
  ['26','記述【思判表】','新田開発','耕地面積の変化を数値から読み取り、原因と結びつけて説明させる。'],
  ['27','記述','株仲間','人物名を示さないカードから田沼意次の政策を特定させる。'],
  ['28','整序【思判表】','徳川綱吉／享保の改革／田沼意次／寛政の改革','カードの内容から4つの政治を特定し、年代順に並べさせる。'],
  ['29','選択【思判表】','（寛政の改革の評価）','狂歌と自叙伝を関連づけ、風刺の意味を解釈させる。'],
  ['30','選択','松尾芭蕉／近松門左衛門／化政文化','元禄文化と化政文化の担い手・時期の識別。'],
  ['31','記述【思判表】','蘭学／解体新書（享保の改革）','資料NのBと資料Qを関連づけ、洋書輸入制限の緩和を導かせる。'],
];
body.push(dataTable([700, 1700, 3000, 4200], [
  new TableRow({ children: [
    tcell('問', {width:700, font:GO, shade:'EFEFEF'}), tcell('形式', {width:1700, font:GO, shade:'EFEFEF'}),
    tcell('到達する基本用語', {width:3000, font:GO, shade:'EFEFEF'}), tcell('ねらい', {width:4200, font:GO, shade:'EFEFEF'}) ]}),
  ...aim.map(a => new TableRow({ children: [
    tcell(a[0], {width:700, size:18}), tcell(a[1], {width:1700, size:18}),
    tcell(a[2], {width:3000, size:18, align:AlignmentType.LEFT}), tcell(a[3], {width:4200, size:18, align:AlignmentType.LEFT}) ]})),
]));

// ============================ 出力 ============================
const doc = new Document({
  styles: { default: { document: { run: { font: { name: MIN, hint: 'eastAsia' }, size: 21 } } } },
  sections: [{
    properties: {
      page: {
        size: { width: convertMillimetersToTwip(210), height: convertMillimetersToTwip(297) },
        margin: { top: convertMillimetersToTwip(20), bottom: convertMillimetersToTwip(18),
                  left: convertMillimetersToTwip(18), right: convertMillimetersToTwip(18) },
      },
    },
    children: body,
  }],
});

Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2] || 'out.docx', b); console.log('written'); });
