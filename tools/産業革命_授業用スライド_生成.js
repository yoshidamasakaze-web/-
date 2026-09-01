const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const OUTDIR = process.env.OUTDIR || ".";
const IMGDIR = path.join(OUTDIR, "画像");

/* ============================================================
   使用する当時の絵（すべて著作権保護期間が満了したパブリックドメイン）
   画像/01.jpg … 12.jpg を置くと、自動的にスライドにはめ込まれます。
   ============================================================ */
const IMAGES = {
  1:  { t:"マンチェスターの遠景", y:"1852年", a:"ウィリアム・ワイルド",
        w:"Manchester from Kersal Moor",
        d:"手前は牧歌的な野原。その奥に、工場の煙突が林立して煙を吐いている。功と罪が一枚に写り込んだ絵。",
        q:"Manchester from Kersal Moor Wyld",
        f:"Manchester from Kersal Moor.jpg" },
  2:  { t:"糸車で糸を紡ぐ", y:"18世紀", a:"—",
        w:"Spinning wheel",
        d:"産業革命の前。糸は家で、一本ずつ手で紡いでいた。布を1着分織るのに何日もかかった。",
        q:"spinning wheel 18th century cottage industry",
        f:"（カテゴリ Spinning wheels から選ぶ）" },
  3:  { t:"ジェニー紡績機", y:"1760年代", a:"ジェームズ・ハーグリーブス",
        w:"Spinning jenny",
        d:"一人で何本もの糸を同時に紡げる機械。ここから全部が始まった。",
        q:"spinning jenny Hargreaves",
        f:"Spinning jenny.jpg" },
  4:  { t:"ワットの蒸気機関", y:"1780年代", a:"ジェームズ・ワット",
        w:"Watt steam engine",
        d:"風でも水でも筋肉でもない動力。天気にも場所にも左右されず、休まず動き続ける。",
        q:"Watt steam engine engraving 18th century",
        f:"（カテゴリ James Watt steam engines から選ぶ）" },
  5:  { t:"力織機のならぶ工場の中", y:"1835年", a:"エドワード・ベインズ『英国綿工業史』の図版",
        w:"Powerloom weaving",
        d:"広いホールに機械がずらりと並び、女性たちが働いている。家の仕事が「工場」に変わった瞬間。",
        q:"power loom weaving 1835 Baines",
        f:"Powerloom weaving in 1835.jpg" },
  6:  { t:"リヴァプール〜マンチェスター鉄道の開通", y:"1830〜31年", a:"T.T.ベリーの版画",
        w:"Liverpool and Manchester Railway",
        d:"世界初の本格的な旅客鉄道。人と物が、一日で遠くまで動くようになった。",
        q:"Liverpool and Manchester Railway 1831 Bury",
        f:"（カテゴリ Liverpool and Manchester Railway から選ぶ）" },
  7:  { t:"ロンドン万国博覧会・水晶宮", y:"1851年", a:"—",
        w:"The Crystal Palace",
        d:"鉄とガラスだけでできた巨大な建物に、世界中の工業製品が集められた。イギリスの絶頂。",
        q:"Crystal Palace Great Exhibition 1851",
        f:"（カテゴリ Great Exhibition 1851 から選ぶ）" },
  8:  { t:"「ロンドン、鉄道の上から」", y:"1872年", a:"ギュスターヴ・ドレ",
        w:"Over London by Rail",
        d:"高架鉄道から見下ろした労働者の住まい。裏庭に洗濯物が並び、煙突が空をふさぐ。スラムを描いた最も有名な絵。",
        q:"Gustave Dore Over London by Rail",
        f:"Dore London.jpg / Over London by Rail" },
  9:  { t:"炭鉱で石炭を引く子ども", y:"1842年", a:"英国王立委員会 児童労働調査報告書の挿絵",
        w:"Children in coal mines",
        d:"体に鎖を巻きつけ、四つんばいで石炭の車を引く。この報告書が世の中を動かした。",
        q:"1842 Royal Commission children mines report illustration",
        f:"（カテゴリ Child labour in the United Kingdom から選ぶ）" },
  10: { t:"紡績工場で働く子ども", y:"19世紀前半", a:"—",
        w:"Child labour in textile mills",
        d:"動いている機械の下にもぐって糸くずを取る。体が小さいからという理由で子どもの仕事にされた。",
        q:"child labour cotton mill 19th century engraving",
        f:"（カテゴリ Child labour から選ぶ）" },
  11: { t:"綿花を摘む奴隷", y:"19世紀", a:"—",
        w:"Slaves picking cotton",
        d:"イギリスの工場が飲み込んだ大量の綿花は、アメリカ南部の奴隷労働が支えていた。",
        q:"slaves picking cotton plantation 19th century engraving",
        f:"（カテゴリ Cotton plantations から選ぶ）" },
  12: { t:"インドの手織り職人", y:"19世紀", a:"—",
        w:"Indian handloom weaver",
        d:"イギリスの安い機械織りの布が流れ込み、何百年も続いたインドの手織物業がこわれた。",
        q:"Indian handloom weaver 19th century",
        f:"（カテゴリ Handloom weaving in India から選ぶ）" }
};

const searchURL = (n) =>
  "https://commons.wikimedia.org/w/index.php?search=" +
  encodeURIComponent(IMAGES[n].q) + "&title=Special:MediaSearch&type=image";

/* ============================================================ */

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "産業革命は私たち人類に何をもたらしたのか（授業用）";

const W = 13.333, H = 7.5, M = 0.62, CW = W - M * 2;

const INK="23252A", INK2="33363D", BRICK="A63D2E", BRASS="D9A441";
const PAPER="F1F1EF", WHITE="FFFFFF", MUTED="6E7178", MUTED_D="9DA1A9", LINE="DFDFDC";
const F = "Meiryo";

const sh = () => ({ type:"outer", color:"000000", blur:10, offset:2, angle:90, opacity:0.10 });

function slide(dark){ const s = pres.addSlide(); s.background = { color: dark?INK:WHITE }; return s; }

function phase(s, num, label, dark){
  s.addShape(pres.ShapeType.ellipse,{x:M,y:0.32,w:0.38,h:0.38,fill:{color:dark?BRASS:INK},line:{type:"none"}});
  s.addText(String(num),{x:M,y:0.32,w:0.38,h:0.38,isTextBox:true,margin:0,align:"center",valign:"middle",
    fontFace:F,fontSize:13,bold:true,color:dark?INK:WHITE});
  s.addText(label,{x:M+0.52,y:0.32,w:CW-0.52,h:0.38,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:12.5,bold:true,color:dark?MUTED_D:MUTED,charSpacing:1});
}
function title(s,t,dark,size,y,h){
  s.addText(t,{x:M,y:y===undefined?0.88:y,w:CW,h:h===undefined?0.68:h,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:size||30,bold:true,color:dark?WHITE:INK,lineSpacing:(size||30)*1.25});
}
function sub(s,t,dark,y){
  s.addText(t,{x:M,y:y===undefined?1.60:y,w:CW,h:0.40,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:13.5,color:dark?MUTED_D:MUTED});
}
function card(s,x,y,w,h,dark,tint){
  s.addShape(pres.ShapeType.roundRect,{x,y,w,h,rectRadius:0.06,
    fill:{color:tint||(dark?INK2:PAPER)},line:{type:"none"},shadow:sh()});
}

/* --- 画像枠：画像/NN.jpg があれば貼り、なければ差し込み用の枠を描く --- */
let missing = [];
function pic(s, n, x, y, w, h, dark){
  const p = path.join(IMGDIR, String(n).padStart(2,"0") + ".jpg");
  const p2 = path.join(IMGDIR, String(n).padStart(2,"0") + ".png");
  const real = fs.existsSync(p) ? p : (fs.existsSync(p2) ? p2 : null);
  if (real){
    s.addImage({ path: real, x, y, w, h, sizing:{ type:"cover", w, h } });
    return;
  }
  if (!missing.includes(n)) missing.push(n);
  const m = IMAGES[n];
  s.addShape(pres.ShapeType.roundRect,{x,y,w,h,rectRadius:0.04,
    fill:{color: dark ? "2E3138" : "EAE8E4"},
    line:{color: dark ? "4A4E57" : "CFCCC6", width:1.25, dashType:"dash"}});
  s.addShape(pres.ShapeType.roundRect,{x:x+0.16,y:y+0.16,w:1.02,h:0.32,rectRadius:0.05,
    fill:{color:BRASS},line:{type:"none"}});
  s.addText("画像 " + n,{x:x+0.16,y:y+0.16,w:1.02,h:0.32,isTextBox:true,margin:0,align:"center",valign:"middle",
    fontFace:F,fontSize:11,bold:true,color:INK});
  s.addText(m.t,{x:x+0.20,y:y+0.60,w:w-0.40,h:0.36,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:14,bold:true,color: dark?WHITE:INK});
  s.addText(m.d,{x:x+0.20,y:y+1.00,w:w-0.40,h:Math.max(0.4,h-1.50),isTextBox:true,margin:0,
    valign:"middle",fontFace:F,fontSize:11,color: dark?MUTED_D:MUTED,lineSpacing:16});
  s.addText("ここに画像をはめる（探し方は最終スライド）",{x:x+0.20,y:y+h-0.44,w:w-0.40,h:0.30,isTextBox:true,margin:0,
    valign:"middle",fontFace:F,fontSize:10,italic:true,color: dark?"7C818B":"A5A29B"});
}
function credit(s, n, x, y, w, dark){
  const m = IMAGES[n];
  s.addText("画像" + n + "：" + m.t + "（" + m.y + "／" + m.a + "）",
    {x,y,w,h:0.28,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:9.5,color: dark?"7C818B":"9A978F"});
}
const notes = (s,t) => s.addNotes(t);

/* ============================================================
   1　表紙
   ============================================================ */
{
  const s = slide(true);
  pic(s, 1, 0, 0, W, 3.60, true);
  s.addShape(pres.ShapeType.rect,{x:0,y:3.60,w:W,h:H-3.60,fill:{color:INK},line:{type:"none"}});
  s.addText("第2学年　社会科（歴史的分野）", {x:M,y:3.86,w:CW,h:0.34,isTextBox:true,margin:0,
    fontFace:F,fontSize:13,color:BRASS,charSpacing:2});
  s.addText("産業革命は私たち人類に\n何をもたらしたのか", {x:M,y:4.24,w:CW,h:1.90,isTextBox:true,margin:0,
    fontFace:F,fontSize:40,bold:true,color:WHITE,lineSpacing:56});
  credit(s, 1, M, 6.35, CW, true);
  notes(s,"【授業用スライド（全30枚）】そのまま投影して1時間の授業を流せます。設計の意図・活動形態・時間配分は、もう一つのスライド（指導用・全31枚）を参照してください。\n\n表紙の絵（ワイルド「マンチェスターの遠景」1852年）は、この授業のすべてが詰まった一枚です。手前は牧歌的な野原、奥は煙突の林。19枚目でもう一度、大きく提示します。\n\n授業前に「画像」フォルダに 01.jpg〜12.jpg を入れて再生成すると、枠が実際の絵に置き換わります。集め方は最終スライドに載せてあります。");
}

/* ============================================================
   2-6　導入（トンネルの入口）
   ============================================================ */
{
  const s = slide(true);
  phase(s,1,"導入　0-7分　｜　全体 → ペア",true);
  s.addText("今、着ている服。\nいくらだった？",{x:M,y:2.10,w:CW,h:2.20,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:50,bold:true,color:WHITE,lineSpacing:70});
  card(s,M,4.80,CW,1.00,true,INK2);
  s.addText([{text:"となりの人と 30秒　",options:{bold:true,color:BRASS}},
             {text:"わからなければ「高いと思う？安いと思う？」でよい",options:{color:"E3E4E6"}}],
    {x:M+0.32,y:4.80,w:CW-0.64,h:1.00,isTextBox:true,margin:0,valign:"middle",fontFace:F,fontSize:15});
  notes(s,"【ねらい】トンネルの入口。全員が当事者になる一点に潜らせる。\n\n「教科書を開く前に、一つだけ聞かせて。今きみが着ているその服、いくらだった？」\n制服なら「その制服」「昨日の私服」「今はいている靴下」で代替できます。\n\n【留意】家庭状況に触れる話題です。個人が特定される言わせ方は避け、相場感に留めてください。\n\n2分で切り上げます。「安い」という感覚が出れば十分。");
}
{
  const s = slide(false);
  phase(s,1,"導入　0-7分　｜　全体",false);
  title(s,"200年前は、こうだった",false,30);

  pic(s, 2, M, 1.68, 6.10, 4.30, false);
  credit(s, 2, M, 6.06, 6.10, false);

  const rx = M + 6.44, rw = CW - 6.44;
  card(s, rx, 1.68, rw, 2.00, false, PAPER);
  s.addText("18世紀なかばのイギリス",{x:rx+0.28,y:1.86,w:rw-0.56,h:0.36,isTextBox:true,margin:0,
    valign:"middle",fontFace:F,fontSize:15,bold:true,color:BRICK});
  s.addText("・綿の布は、インドからの輸入品。ぜいたく品\n・糸は家で、一本ずつ手で紡ぐ\n・1着つくるのに、何日もかかる\n・庶民の普段着は、毛や麻のおさがり",
    {x:rx+0.28,y:2.28,w:rw-0.56,h:1.28,isTextBox:true,margin:0,fontFace:F,fontSize:13,color:INK,lineSpacing:20});
  card(s, rx, 3.86, rw, 2.12, false, INK);
  s.addText("この200年ほどで、\n何が起きたのか。",{x:rx+0.30,y:3.86,w:rw-0.60,h:2.12,isTextBox:true,margin:0,
    valign:"middle",fontFace:F,fontSize:22,bold:true,color:WHITE,lineSpacing:34});
  notes(s,"【ねらい】「安い服」が歴史的にはごく最近の現象だと気づかせる。\n\n絵を見せながら「これ、何をしているところだと思う？」と問うと、「糸をつくってる」が出ます。「そう。200年前は、服を1着つくるのに、まずここから始めた」\n\n【留意】当時の値段を現代の円に換算するのは無理があります。聞かれたら「正確な換算は難しいけれど、庶民が気軽に買える物ではなかった」と答えてください。次のスライドで換算の要らない“倍率”を出します。\n\n2分。");
}
{
  const s = slide(false);
  phase(s,1,"導入　0-7分　｜　全体",false);
  title(s,"数字で見ると、こうなる",false,30);
  sub(s,"イギリスの綿工業に起きたこと");

  const cw = (CW-0.40)/2;
  card(s,M,2.14,cw,2.40,false,PAPER);
  s.addText("約 1/12",{x:M,y:2.44,w:cw,h:1.10,isTextBox:true,margin:0,align:"center",valign:"middle",
    fontFace:F,fontSize:52,bold:true,color:BRASS});
  s.addText("綿糸1ポンドの値段\n1786年 38シリング → 1832年 3シリング",
    {x:M+0.20,y:3.58,w:cw-0.40,h:0.72,isTextBox:true,margin:0,align:"center",
     fontFace:F,fontSize:12.5,color:MUTED,lineSpacing:18});
  card(s,M+cw+0.40,2.14,cw,2.40,false,PAPER);
  s.addText("約 85倍",{x:M+cw+0.40,y:2.44,w:cw,h:1.10,isTextBox:true,margin:0,align:"center",valign:"middle",
    fontFace:F,fontSize:52,bold:true,color:BRASS});
  s.addText("イギリスの綿花輸入量\n1780年 約690万ポンド → 1850年 約5億8,800万ポンド",
    {x:M+cw+0.60,y:3.58,w:cw-0.40,h:0.72,isTextBox:true,margin:0,align:"center",
     fontFace:F,fontSize:12.5,color:MUTED,lineSpacing:18});

  card(s,M,4.88,CW,1.10,false,INK);
  s.addText([{text:"なぜ？　",options:{bold:true,color:BRASS,fontSize:24}},
             {text:"値段が1/12になり、量が85倍になるようなことが、なぜ起きたのか。",options:{color:WHITE,fontSize:19}}],
    {x:M+0.32,y:4.88,w:CW-0.64,h:1.10,isTextBox:true,margin:0,valign:"middle",fontFace:F,bold:true});
  s.addText("※数値は概数。イギリス経済史の統計による。",
    {x:M,y:6.24,w:CW,h:0.30,isTextBox:true,margin:0,fontFace:F,fontSize:10,color:MUTED});
  notes(s,"【ねらい】トンネルの一番深いところ。ここで「なぜ」を子どもの内側から出させる。\n\n数字を読み上げたら一拍おいてください。「値段が12分の1。量が85倍。……これ、ふつうのことだと思う？」\n子どもが「機械？」「工場？」と口にすれば成功です。\n\n【留意】数値は概数です。使用中の教科書・資料集に別の数値があればそちらを優先してください。\n\n2分。");
}
{
  const s = slide(true);
  phase(s,1,"導入　0-7分　｜　全体",true);
  title(s,"答えは、この機械から始まった",true,30);

  pic(s, 3, M, 1.66, 6.30, 4.28, true);
  credit(s, 3, M, 6.02, 6.30, true);

  const rx = M+6.62, rw = CW-6.62;
  s.addText("ジェニー紡績機",{x:rx,y:1.90,w:rw,h:0.60,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:30,bold:true,color:BRASS});
  s.addText("1760年代　ハーグリーブス",{x:rx,y:2.52,w:rw,h:0.34,isTextBox:true,margin:0,
    fontFace:F,fontSize:13,color:MUTED_D});
  card(s,rx,3.06,rw,2.60,true,INK2);
  s.addText("一人で、何本もの糸を\n同時に紡げるようになった。\n\nこれまで一人が一本ずつ\n紡いでいたものが、\n一気に何倍にもなった。",
    {x:rx+0.28,y:3.06,w:rw-0.56,h:2.60,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:15,color:WHITE,lineSpacing:26});
  notes(s,"【ねらい】既習事項（第5・6時）の確認。説明せず、子どもに言わせる。\n\n「この機械、名前わかる人？」で十分です。出なければ教師が示します。\n\n1分。ここで説明を始めると導入が崩れます。確認だけして次へ。");
}
{
  const s = slide(true);
  phase(s,1,"導入　0-7分　｜　全体",true);
  title(s,"その正体が「産業革命」",true,30);
  sub(s,"18世紀後半、イギリスから　― 前の時間に学んだこと",true);

  const cw = (CW-0.40)/2;
  pic(s, 4, M, 2.10, cw, 3.06, true);
  s.addText("動力　― 蒸気機関",{x:M,y:5.22,w:cw,h:0.34,isTextBox:true,margin:0,
    fontFace:F,fontSize:15,bold:true,color:BRASS});
  s.addText("風・水・筋肉に代わる力。天気にも場所にも左右されず、休まず動く。",
    {x:M,y:5.56,w:cw,h:0.56,isTextBox:true,margin:0,fontFace:F,fontSize:12,color:"E3E4E6",lineSpacing:17});
  credit(s, 4, M, 6.16, cw, true);

  pic(s, 5, M+cw+0.40, 2.10, cw, 3.06, true);
  s.addText("工場　― 工場制機械工業",{x:M+cw+0.40,y:5.22,w:cw,h:0.34,isTextBox:true,margin:0,
    fontFace:F,fontSize:15,bold:true,color:BRASS});
  s.addText("家の仕事だった糸紡ぎと機織りが、人を1か所に集めて大量に生産する「工場」になった。",
    {x:M+cw+0.40,y:5.56,w:cw,h:0.56,isTextBox:true,margin:0,fontFace:F,fontSize:12,color:"E3E4E6",lineSpacing:17});
  credit(s, 5, M+cw+0.40, 6.16, cw, true);
  notes(s,"【押さえること】機械・動力・工場の3つ。前時の復習なので、子どもに言わせて確認するだけ。\n\n右の絵（力織機の工場・1835年）はよく見せてください。「働いているのは誰？」と問うと「女の人」と返ってきます。「そう。そして子どももいた。あとで出てくるよ」と伏線を張っておくと、影の資料が効きます。\n\n1分。");
}

/* ============================================================
   7　本時の問い
   ============================================================ */
{
  const s = slide(true);
  phase(s,2,"問いの設定　7-10分　｜　全体",true);
  s.addText("本時の問い",{x:M,y:1.50,w:CW,h:0.40,isTextBox:true,margin:0,
    fontFace:F,fontSize:14,bold:true,color:BRASS,charSpacing:3});
  s.addText("産業革命は\n私たち人類に\n何をもたらしたのか",{x:M,y:2.05,w:CW,h:3.20,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:48,bold:true,color:WHITE,lineSpacing:72});
  s.addText("「よかった／悪かった」の一言では終わらせない。光と影の両方を、資料で確かめる。",
    {x:M,y:5.56,w:CW,h:0.44,isTextBox:true,margin:0,valign:"middle",fontFace:F,fontSize:14,color:MUTED_D});
  notes(s,"【板書】この問いを黒板の中央上部に書き、授業の最後まで消さない。\n\n【ワークシート】問いを書き写させ、この時点での予想を一言だけ書かせます（30秒）。授業後に自分の変化を見取る起点になります。ここを飛ばすと、まとめの振り返りが浅くなります。\n\n3分。");
}

/* ============================================================
   8　活動指示
   ============================================================ */
{
  const s = slide(false);
  phase(s,3,"展開1　10-22分　｜　グループ（4人）",false);
  title(s,"これから見る資料を「光」と「影」に分けよう",false,29);
  sub(s,"12分　／　机の上を3つの場所に分けて置く");

  const cw = (CW-0.80)/3;
  const zones = [["光","人類にとって\nプラスだったこと",BRASS,INK],
                 ["迷い","どちらとも\n言い切れないもの","555961",WHITE],
                 ["影","人類にとって\nマイナスだったこと",BRICK,WHITE]];
  zones.forEach((z,i)=>{
    const x = M + i*(cw+0.40);
    card(s,x,2.24,cw,1.86,false,z[2]);
    s.addText(z[0],{x,y:2.42,w:cw,h:0.60,isTextBox:true,margin:0,align:"center",valign:"middle",
      fontFace:F,fontSize:30,bold:true,color:z[3]});
    s.addText(z[1],{x,y:3.06,w:cw,h:0.84,isTextBox:true,margin:0,align:"center",
      fontFace:F,fontSize:13,color:z[3],lineSpacing:20});
  });
  card(s,M,4.38,CW,1.70,false,INK);
  s.addText([{text:"いちばん大事なのは「迷い」の場所です。\n",options:{bold:true,color:BRASS,fontSize:19}},
             {text:"光にも影にも置けるものが必ずあります。迷ったら無理に決めず、真ん中に置いて「なぜ迷うのか」を班で話してください。あとで全体に聞きます。",options:{color:WHITE,fontSize:14}}],
    {x:M+0.32,y:4.38,w:CW-0.64,h:1.70,isTextBox:true,margin:0,valign:"middle",fontFace:F,lineSpacing:24});
  notes(s,"【活動形態】4人グループ。次の9〜18枚目の資料を順に投影しながら分類させます。カードを印刷して配る場合は、12・18枚目を印刷して切り分けてください。\n\n【手順】\n①9〜11枚目（光）を見せる → 班で話す\n②13〜17枚目（影）を見せる → 班で話す\n③「迷い」に置いたものについて、なぜ迷うのかを話す（3分）\n\n【机間指導で揺さぶる】\n・迷いが空の班 →「光4番（人口が3倍）を見て。増えた人は、みんな幸せに暮らせたと思う？」\n・全部影に寄せた班 →「じゃあ、今きみが着ている服は？」\n・全部光に寄せた班 → 14枚目の炭鉱の子どもを指す\n\n【評価】この場面で「思考・判断・表現」を見取ります。分類の結果ではなく、迷った理由を言えるかを見てください。");
}

/* ============================================================
   9-12　光の資料
   ============================================================ */
{
  const s = slide(false);
  phase(s,3,"展開1　｜　資料【光】",false);
  title(s,"人と物が、一日で遠くまで動くようになった",false,29);

  pic(s, 6, M, 1.66, 7.70, 4.30, false);
  credit(s, 6, M, 6.04, 7.70, false);

  const rx = M+8.02, rw = CW-8.02;
  card(s,rx,1.66,rw,4.30,false,PAPER);
  s.addText("1830年",{x:rx+0.26,y:1.88,w:rw-0.52,h:0.50,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:26,bold:true,color:BRASS});
  s.addText("リヴァプール〜\nマンチェスター\n鉄道が開通",{x:rx+0.26,y:2.40,w:rw-0.52,h:1.10,isTextBox:true,margin:0,
    fontFace:F,fontSize:15,bold:true,color:INK,lineSpacing:24});
  s.addText("それまで馬車で何日もかかった道のりが、数時間になった。\n\n運賃も下がり、ふつうの人が旅をするようになった。",
    {x:rx+0.26,y:3.62,w:rw-0.52,h:2.16,isTextBox:true,margin:0,
     fontFace:F,fontSize:13,color:MUTED,lineSpacing:21});
  notes(s,"【光の資料①】「速さ」\n\n「この絵、何が走ってる？」→「汽車」。「そう。世界で初めての、本格的な旅客鉄道。1830年」\n\n発問：「鉄道ができて、いちばん喜んだのは誰だと思う？」\n工場主（原料と製品を運べる）、商人、旅行する人……いろいろ出ます。ここで「じゃあ、困った人はいなかった？」と一言添えておくと、後半の「影」が効きます（馬車屋、運河の船頭）。\n\n1分半。");
}
{
  const s = slide(false);
  phase(s,3,"展開1　｜　資料【光】",false);
  title(s,"世界中の工業製品が、一か所に集まった",false,29);

  pic(s, 7, M, 1.66, 7.70, 4.30, false);
  credit(s, 7, M, 6.04, 7.70, false);

  const rx = M+8.02, rw = CW-8.02;
  card(s,rx,1.66,rw,4.30,false,PAPER);
  s.addText("1851年",{x:rx+0.26,y:1.88,w:rw-0.52,h:0.50,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:26,bold:true,color:BRASS});
  s.addText("ロンドン\n万国博覧会\n（水晶宮）",{x:rx+0.26,y:2.40,w:rw-0.52,h:1.10,isTextBox:true,margin:0,
    fontFace:F,fontSize:15,bold:true,color:INK,lineSpacing:24});
  s.addText("鉄とガラスだけでつくられた巨大な建物。\n\n世界中から工業製品が集まり、半年で600万人が見に来た。イギリスの絶頂。",
    {x:rx+0.26,y:3.62,w:rw-0.52,h:2.16,isTextBox:true,margin:0,
     fontFace:F,fontSize:13,color:MUTED,lineSpacing:21});
  notes(s,"【光の資料②】「豊かさの象徴」\n\n「この建物、何でできていると思う？」→ 鉄とガラス。「産業革命でつくれるようになった材料だけで建てた。当時の人は“水晶宮”と呼んだ」\n\n入場者数600万人は概数です（当時のイギリスの人口が約2,100万人。人口の4分の1以上が見に来た計算になる、という言い方が効きます）。\n\n1分。");
}
{
  const s = slide(false);
  phase(s,3,"展開1　｜　資料【光】",false);
  title(s,"人類は、はじめて「豊かになり続けた」",false,29);
  sub(s,"世界の一人あたりGDPの推移（1990年国際ドル・概数）");

  pres.ChartType && s.addChart(pres.ChartType.line,[{
    name:"世界の一人あたりGDP",
    labels:["1年","1000年","1500年","1700年","1820年","1870年","1913年","1950年","2000年"],
    values:[467,450,566,616,666,870,1524,2111,6039]
  }],{
    x:M,y:2.12,w:CW-4.30,h:3.50,
    chartColors:[BRASS],lineSize:4,lineSmooth:false,
    showLegend:false,showTitle:false,showValue:false,
    catAxisLabelColor:MUTED,valAxisLabelColor:MUTED,
    catAxisLabelFontFace:F,valAxisLabelFontFace:F,
    catAxisLabelFontSize:11,valAxisLabelFontSize:11,
    valGridLine:{color:LINE,size:1},catGridLine:{style:"none"},
    valAxisMinVal:0,lineDataSymbol:"circle",lineDataSymbolSize:7
  });
  const px = W-M-4.00;
  card(s,px,2.12,4.00,3.50,false,PAPER);
  s.addText("読み取ろう",{x:px+0.26,y:2.32,w:3.48,h:0.34,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:15,bold:true,color:BRICK});
  s.addText("・1年から1700年までの1700年間で、豊かさはどれだけ増えた？\n・グラフが折れ曲がるのは、いつごろ？\n・そこで何が起きていた？",
    {x:px+0.26,y:2.72,w:3.48,h:1.58,isTextBox:true,margin:0,
     fontFace:F,fontSize:13,color:INK,lineSpacing:20,paraSpaceAfter:7});
  card(s,px+0.20,4.46,3.60,0.96,false,INK);
  s.addText("人類の歴史の99%は、\n豊かさが横ばいだった。",{x:px+0.38,y:4.46,w:3.24,h:0.96,isTextBox:true,margin:0,
    valign:"middle",fontFace:F,fontSize:14,bold:true,color:WHITE,lineSpacing:21});
  s.addText("出典：マディソン・プロジェクト・データベースによる推計（概数）。※横軸の年の間隔は均等ではありません。",
    {x:M,y:H-0.62,w:CW,h:0.30,isTextBox:true,margin:0,fontFace:F,fontSize:10,color:MUTED});
  notes(s,"【光の資料③】「豊かさ」\n\n班の作業を一度止めて、45秒だけ全員でこのグラフを見ます。\n「手を止めて。これは世界の一人あたりの豊かさを、西暦1年から2000年まで並べたグラフ。何か気づく？」\n「ずっと平ら」「最後だけ急に上がる」が出れば十分。「折れ曲がるのは1800年代。産業革命だね」\n\n【留意】横軸は年が等間隔ではありません（1年→1000年と1950年→2000年が同じ幅）。上位の子から指摘が出たら褒めてください。資料を疑う態度そのものが社会科の学力です。\n\n【留意2】GDPは豊かさの一面にすぎず、格差・環境・健康を含みません。26枚目で効いてきます。ここでは深追いしない。");
}
{
  const s = slide(false);
  phase(s,3,"展開1　｜　資料【光】まとめ",false);
  title(s,"光　― 産業革命がもたらしたもの ①",false,29,0.88,0.55);

  const cards = [
    ["光1　値段","綿糸の値段が約1/12に。ぜいたく品だった綿の服を、ふつうの人が買えるようになった。"],
    ["光2　速さ","1830年、鉄道が開通。人と物が、一日で遠くまで運べるようになった。"],
    ["光3　力","蒸気機関が、風・水・筋力に代わる動力になった。天気や場所にしばられず、休みなく動かせる。"],
    ["光4　人の数","イギリスの人口は1801年 約1,060万人 → 1901年 約3,700万人へ。より多くの人が生きられるようになった。"],
    ["光5　豊かさ","世界の一人あたりGDPが、人類の歴史ではじめて、下がらずに上がり続けるようになった。"],
    ["光6　仕事","農村から都市へ人が移り、それまで存在しなかった職業と働き口が大量に生まれた。"]
  ];
  const cw=(CW-0.56)/3, ch=1.98;
  cards.forEach((c,i)=>{
    const col=i%3,row=Math.floor(i/3);
    const x=M+col*(cw+0.28), y=1.72+row*(ch+0.22);
    card(s,x,y,cw,ch,false,PAPER);
    s.addText(c[0],{x:x+0.24,y:y+0.16,w:cw-0.48,h:0.38,isTextBox:true,margin:0,valign:"middle",
      fontFace:F,fontSize:15,bold:true,color:BRASS});
    s.addText(c[1],{x:x+0.24,y:y+0.60,w:cw-0.48,h:ch-0.78,isTextBox:true,margin:0,
      fontFace:F,fontSize:12.5,color:INK,lineSpacing:19});
  });
  s.addText("※数値は概数。人口はイギリス国勢調査による。",
    {x:M,y:H-0.62,w:CW,h:0.30,isTextBox:true,margin:0,fontFace:F,fontSize:10,color:MUTED});
  notes(s,"【使い方】光の資料の整理。ここまで見てきたものを一覧にして、班の分類作業に使わせます。\n\n印刷して切り分ければ、そのまま資料カードになります。\n\n光4（人口が3倍）は、光にも影にも読めるカードとして意図的に入れてあります。「人が増えた＝生きられる人が増えた」なら光、「増えた人がスラムに押し込められた」なら影。迷いが出なければ、教師からこのカードを指して揺さぶってください。");
}

/* ============================================================
   13-18　影の資料
   ============================================================ */
{
  const s = slide(true);
  phase(s,3,"展開1　｜　資料【影】",true);
  title(s,"人が集まった街は、こうなった",true,29);

  pic(s, 8, M, 1.66, 7.70, 4.30, true);
  credit(s, 8, M, 6.04, 7.70, true);

  const rx = M+8.02, rw = CW-8.02;
  card(s,rx,1.66,rw,4.30,true,INK2);
  s.addText("1872年",{x:rx+0.26,y:1.88,w:rw-0.52,h:0.50,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:26,bold:true,color:BRICK});
  s.addText("ロンドンの\n労働者の住まい",{x:rx+0.26,y:2.40,w:rw-0.52,h:0.80,isTextBox:true,margin:0,
    fontFace:F,fontSize:15,bold:true,color:WHITE,lineSpacing:24});
  s.addText("工場に人が集まり、上下水道のない住宅が密集した。\n\nごみと汚水があふれ、コレラが何度も流行した。煙突の煙が空をふさいだ。",
    {x:rx+0.26,y:3.34,w:rw-0.52,h:2.44,isTextBox:true,margin:0,
     fontFace:F,fontSize:13,color:"C8CBD1",lineSpacing:21});
  notes(s,"【影の資料①】「まち」\n\nドレの「ロンドン、鉄道の上から」（1872年）。スラムを描いた最も有名な絵です。\n\n「この絵、何が見える？」と問うと、洗濯物・煙突・びっしり並んだ家、が出ます。「ここに、工場で働く人たちが住んでいた」\n\n発問：「さっきの水晶宮と、同じ街だと思う？」——同じロンドンです。ここで子どもが黙ります。\n\n1分半。");
}
{
  const s = slide(true);
  phase(s,3,"展開1　｜　資料【影】",true);
  title(s,"工場と炭鉱で働いたのは、大人だけではなかった",true,28);

  const cw = (CW-0.40)/2;
  pic(s, 9, M, 1.58, cw, 3.20, true);
  s.addText("炭鉱　― 石炭を引く子ども",{x:M,y:4.86,w:cw,h:0.34,isTextBox:true,margin:0,
    fontFace:F,fontSize:15,bold:true,color:BRICK});
  s.addText("体に鎖を巻きつけ、四つんばいで石炭の車を引いた。1842年の調査報告書が、この実態を世に知らせた。",
    {x:M,y:5.22,w:cw,h:0.62,isTextBox:true,margin:0,fontFace:F,fontSize:12,color:"C8CBD1",lineSpacing:18});
  credit(s, 9, M, 5.90, cw, true);

  pic(s, 10, M+cw+0.40, 1.58, cw, 3.20, true);
  s.addText("紡績工場　― 機械の下の子ども",{x:M+cw+0.40,y:4.86,w:cw,h:0.34,isTextBox:true,margin:0,
    fontFace:F,fontSize:15,bold:true,color:BRICK});
  s.addText("動いている機械の下にもぐって糸くずを取った。体が小さいから、という理由で子どもの仕事にされた。",
    {x:M+cw+0.40,y:5.22,w:cw,h:0.62,isTextBox:true,margin:0,fontFace:F,fontSize:12,color:"C8CBD1",lineSpacing:18});
  credit(s, 10, M+cw+0.40, 5.90, cw, true);

  card(s,M,6.24,CW,0.82,true,"2C2F36");
  s.addText("9歳の子どもが、1日12〜16時間。学校には行かない。",
    {x:M+0.32,y:6.24,w:CW-0.64,h:0.82,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:17,bold:true,color:WHITE});
  notes(s,"【影の資料②】「子ども」——この授業でいちばん重い2枚です。\n\n【提示の仕方】言葉を足さずに、まず10秒だまって見せてください。子どもが自分で気づきます。\nそのあと一言だけ。「働いているの、何歳くらいに見える？」\n\n「9歳」という数字は、1833年の工場法が9歳未満の雇用を禁止したことに対応します。逆に言えば、それまでは9歳未満も働いていたということです。\n\n【留意】労働時間や年齢は工場・時期によって幅があります。断定的に一つの数字だけを示さず、「12〜16時間」と幅で伝えてください。\n\n【つなぎ】「じゃあ、この子たちは何歳まで生きたと思う？」と問うて次のスライドへ。\n\n2分。");
}
{
  const s = slide(false);
  phase(s,3,"展開1　｜　資料【影】",false);
  title(s,"同じ街に住んで、これだけ違った",false,29);
  sub(s,"マンチェスターの階級別の平均死亡年齢（1842年・チャドウィック報告）");

  s.addChart(pres.ChartType.bar,[{
    name:"平均死亡年齢（歳）",
    labels:["上流・専門職の家","商人・農場主の家","労働者・職人の家"],
    values:[38,20,17]
  }],{
    x:M,y:2.14,w:CW-4.30,h:3.38,
    barDir:"col",chartColors:[BRICK],barGapWidthPct:120,
    showLegend:false,showTitle:false,
    showValue:true,dataLabelPosition:"outEnd",
    dataLabelColor:INK,dataLabelFontFace:F,dataLabelFontSize:16,
    dataLabelFormatCode:'0"歳"',
    catAxisLabelColor:MUTED,valAxisLabelColor:MUTED,
    catAxisLabelFontFace:F,valAxisLabelFontFace:F,
    catAxisLabelFontSize:12,valAxisLabelFontSize:11,
    valGridLine:{color:LINE,size:1},catGridLine:{style:"none"},
    valAxisMinVal:0,valAxisMaxVal:45
  });
  const px = W-M-4.00;
  card(s,px,2.14,4.00,1.70,false,INK);
  s.addText("同じ街の、同じ時代。\nちがうのは、\n生まれた家だけ。",{x:px+0.30,y:2.14,w:3.40,h:1.70,isTextBox:true,margin:0,
    valign:"middle",fontFace:F,fontSize:17,bold:true,color:WHITE,lineSpacing:27});
  card(s,px,3.98,4.00,1.54,false,PAPER);
  s.addText([{text:"この数字の読み方\n",options:{bold:true,color:BRICK}},
             {text:"「平均死亡年齢」は平均寿命とは違います。当時は赤ちゃんのうちに亡くなる子が非常に多く、その分だけ平均が下がっています。17歳まで生きた人がそこで死んだ、という意味ではありません。",options:{color:INK}}],
    {x:px+0.26,y:4.14,w:3.48,h:1.24,isTextBox:true,margin:0,fontFace:F,fontSize:11.5,lineSpacing:16});
  s.addText("出典：エドウィン・チャドウィック『大英帝国における労働者階級の衛生状態に関する報告』（1842年）。",
    {x:M,y:H-0.62,w:CW,h:0.30,isTextBox:true,margin:0,fontFace:F,fontSize:10,color:MUTED});
  notes(s,"【最重要の留意点】\n「平均死亡年齢17歳」を「平均寿命17歳」と説明しないでください。当時の乳幼児死亡率の高さを強く反映した数字です。右下の注記は口頭でも必ず触れてください。\n\n正しい押さえ方：「これは、生まれた子が何歳まで生きられたかの平均。労働者の家では赤ちゃんのうちに亡くなる子がとても多かった。だから17歳まで下がる。上流の家では38歳。同じ街に住んでいるのに、生まれた家でこれだけ違った」\n\n数字の扱いを丁寧にすることが、そのまま「資料を正しく読む」という知識・技能の指導になります。授業の見せ場にしてください。");
}
{
  const s = slide(true);
  phase(s,3,"展開1　｜　資料【影】",true);
  title(s,"影は、イギリスの中だけでは終わらなかった",true,28);

  const cw = (CW-0.40)/2;
  pic(s, 11, M, 1.58, cw, 3.20, true);
  s.addText("アメリカ南部　― 綿花と奴隷",{x:M,y:4.86,w:cw,h:0.34,isTextBox:true,margin:0,
    fontFace:F,fontSize:15,bold:true,color:BRICK});
  s.addText("イギリスの工場が飲み込んだ大量の綿花は、アメリカ南部の奴隷労働が支えていた。機械が増えるほど、奴隷は増えた。",
    {x:M,y:5.22,w:cw,h:0.62,isTextBox:true,margin:0,fontFace:F,fontSize:12,color:"C8CBD1",lineSpacing:18});
  credit(s, 11, M, 5.90, cw, true);

  pic(s, 12, M+cw+0.40, 1.58, cw, 3.20, true);
  s.addText("インド　― 手織り職人",{x:M+cw+0.40,y:4.86,w:cw,h:0.34,isTextBox:true,margin:0,
    fontFace:F,fontSize:15,bold:true,color:BRICK});
  s.addText("イギリスの安い機械織りの布が流れ込み、何百年も続いたインドの手織物業がこわれた。仕事を失った職人が大量に出た。",
    {x:M+cw+0.40,y:5.22,w:cw,h:0.62,isTextBox:true,margin:0,fontFace:F,fontSize:12,color:"C8CBD1",lineSpacing:18});
  credit(s, 12, M+cw+0.40, 5.90, cw, true);

  card(s,M,6.24,CW,0.82,true,"2C2F36");
  s.addText("イギリスの豊かさの「材料」は、どこから来たのか。",
    {x:M+0.32,y:6.24,w:CW-0.64,h:0.82,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:17,bold:true,color:WHITE});
  notes(s,"【影の資料③】「世界」\n\nここは第9時（欧米諸国のアジア進出）への布石です。深追いせず、事実を示すだけで十分。\n\n発問：「イギリスの工場は、原料の綿花をどこから持ってきたと思う？」\n「イギリスでは綿は育たない」ことに気づかせると効果的です。気候的に栽培できません。だから外から持ってくるしかなかった。\n\n最下段の一文を板書に残しておくと、第9時の導入にそのまま使えます。\n\n1分半。");
}
{
  const s = slide(false);
  phase(s,3,"展開1　｜　資料【影】まとめ",false);
  title(s,"影　― 産業革命がもたらしたもの ②",false,29,0.88,0.55);

  const cards = [
    ["影1　子ども","9歳の子どもが工場や炭鉱で働いた。1日12〜16時間。学校には行かなかった。"],
    ["影2　いのち","1842年の報告書では、マンチェスターの労働者の平均死亡年齢は17歳だった。"],
    ["影3　まち","工場に人が集まり、上下水道のないスラムができた。コレラが何度も流行した。"],
    ["影4　空気","石炭を燃やした煤煙が街をおおった。ロンドンは「霧の都」と呼ばれた。"],
    ["影5　世界","安い綿布はインドの手織物業をこわし、原料の綿花はアメリカ南部の奴隷労働が支えた。"],
    ["影6　地球","人類が石炭を大量に燃やし始めた。今の気候変動は、ここから始まっている。"]
  ];
  const cw=(CW-0.56)/3, ch=1.98;
  cards.forEach((c,i)=>{
    const col=i%3,row=Math.floor(i/3);
    const x=M+col*(cw+0.28), y=1.72+row*(ch+0.22);
    card(s,x,y,cw,ch,false,PAPER);
    s.addText(c[0],{x:x+0.24,y:y+0.16,w:cw-0.48,h:0.38,isTextBox:true,margin:0,valign:"middle",
      fontFace:F,fontSize:15,bold:true,color:BRICK});
    s.addText(c[1],{x:x+0.24,y:y+0.60,w:cw-0.48,h:ch-0.78,isTextBox:true,margin:0,
      fontFace:F,fontSize:12.5,color:INK,lineSpacing:19});
  });
  s.addText("※影1は工場法制定前後のイギリス議会報告による。影2はチャドウィック報告（1842年）。",
    {x:M,y:H-0.62,w:CW,h:0.30,isTextBox:true,margin:0,fontFace:F,fontSize:10,color:MUTED});
  notes(s,"【使い方】影の資料の整理。印刷して切り分ければ資料カードになります。\n\n影5と影6は、産業革命の影がイギリス国内の話で終わらないことを示すために必ず入れてください。影6は次時以降と単元のまとめへの布石です。");
}

/* ============================================================
   19　一枚の絵に両方
   ============================================================ */
{
  const s = slide(true);
  phase(s,3,"展開1のまとめ　｜　全体で共有",true);
  pic(s, 1, 0, 0, W, 4.30, true);
  s.addShape(pres.ShapeType.rect,{x:0,y:4.30,w:W,h:H-4.30,fill:{color:INK},line:{type:"none"}});
  s.addText("この一枚に、光と影の両方が写っている。",
    {x:M,y:4.48,w:CW,h:0.62,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:27,bold:true,color:WHITE});

  const cw=(CW-0.40)/2;
  s.addText([{text:"手前　",options:{bold:true,color:BRASS}},
             {text:"牧歌的な野原。人が寝そべり、羊がいる。産業革命の前からある風景。",options:{color:"D8DADE"}}],
    {x:M,y:5.24,w:cw,h:0.54,isTextBox:true,margin:0,fontFace:F,fontSize:13.5,lineSpacing:20});
  s.addText([{text:"奥　",options:{bold:true,color:BRICK}},
             {text:"林立する工場の煙突。空が煙でにごっている。これが新しい風景。",options:{color:"D8DADE"}}],
    {x:M+cw+0.40,y:5.24,w:cw,h:0.54,isTextBox:true,margin:0,fontFace:F,fontSize:13.5,lineSpacing:20});

  card(s,M,5.94,CW,0.86,true,"2C2F36");
  s.addText([{text:"発問　",options:{bold:true,color:BRASS}},
             {text:"「光の側にあるものを1つでも消したら、影も一緒に消える？」",options:{color:WHITE}}],
    {x:M+0.32,y:5.94,w:CW-0.64,h:0.86,isTextBox:true,margin:0,valign:"middle",fontFace:F,fontSize:17,bold:true});
  credit(s, 1, M, 6.92, CW, true);
  notes(s,"【この授業の山場です】\n\n表紙で一度見せた絵を、ここでもう一度、意味づけて見せます。\n\n【進め方】\n①「この絵、最初にも出したよね。もう一回よく見て。手前に何がある？　奥には？」\n②手前＝羊のいる野原、奥＝煙突の林。同じ一枚の絵の中に、産業革命の前と後が同時に写っています。\n③各班の「迷い」に置いたカードを2〜3班に聞く。\n\n【最後の発問の意図】\n光と影が同じ原因から出ていることに気づかせる問いです。「機械をなくせば子どもも働かなくていい。でも服も高くなる」という応答が出れば、展開2への準備が完全に整います。\n\n【時間管理】ここで延びやすい場面です。22分を超えたら切り上げてください。展開2と主発問のほうが本時の中心です。");
}

/* ============================================================
   20-22　展開2　課題の立ち上げ
   ============================================================ */
{
  const s = slide(true);
  phase(s,4,"展開2　22-30分　｜　全体（教師が板書する）",true);
  s.addText("この「影」を、\n当時の人たちは\nどう解決しようとした？",
    {x:M,y:1.86,w:CW,h:3.00,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:42,bold:true,color:WHITE,lineSpacing:64});
  card(s,M,5.16,CW,1.06,true,INK2);
  s.addText([{text:"自分が当時の人だったら、どうする？　",options:{bold:true,color:BRASS}},
             {text:"思いついたことを、そのまま言ってよい。正解はまだ言いません。",options:{color:"E3E4E6"}}],
    {x:M+0.32,y:5.16,w:CW-0.64,h:1.06,isTextBox:true,margin:0,valign:"middle",fontFace:F,fontSize:15});
  notes(s,"【本時でいちばん難しいところです】\n課題を子どもと立ち上げるには、子どもの発言を先に引き出し、あとから歴史的事実を重ねる順序が要ります。逆にすると教師の説明になります。\n\n【進め方】\n①この発問を出し、挙手で自由に出させる（3分）\n②出た案を黒板の「影」の側に箇条書きで残す。評価せず、全部書く\n③5〜6個たまったら次のスライドを出す\n\n【ほぼ必ず出る案】\n・働く時間を法律で決める　・子どもを働かせない　・給料を上げる\n・みんなで集まって文句を言う／ストライキ　・選挙で政治家を変える\n・国がお金を出して助ける　・工場をやめる／機械をこわす\n\n【出ないときの手立て】\n「じゃあ、今の日本では子どもが工場で働いてないよね。なんで働いてないんだと思う？」と現在から逆算させると必ず出ます。\n\n【言ってはいけないこと】\nここで「工場法というものがあってね」と先に言わないでください。子どもの発言を歴史と重ねる快感が、この授業の推進力です。");
}
{
  const s = slide(false);
  phase(s,4,"展開2　22-30分　｜　全体",false);
  title(s,"きみたちが今言ったこと、全部、実際に起きている",false,28);
  sub(s,"子どもの発言　→　歴史上、本当に行われたこと");

  const pairs = [
    ["働く時間を法律で決める","工場法（1833年）","9歳未満の雇用を禁止。13歳未満は1日9時間まで。工場を見回る監督官も置かれた。"],
    ["子どもを働かせない","教育の義務化へ","働かせない代わりに学校へ。のちの義務教育につながる。"],
    ["みんなで集まって声を上げる","労働組合","1824年に団結を禁じる法律が撤廃され、労働者が団結できるようになった。"],
    ["選挙で政治を変える","チャーティスト運動","1830〜40年代。労働者に選挙権を求める大運動が起きた。"],
    ["国が助ける","公衆衛生法（1848年）","上下水道の整備など、街の衛生を国と自治体の仕事にした。"],
    ["仕組みそのものを変える","社会主義","マルクスら。『共産党宣言』1848年。工場を個人のものにしない社会を構想した。"]
  ];
  const rh=0.64, rg=0.08;
  pairs.forEach((p,i)=>{
    const y = 2.10 + i*(rh+rg);
    card(s,M,y,CW,rh,false,PAPER);
    s.addText("「"+p[0]+"」",{x:M+0.24,y,w:3.30,h:rh,isTextBox:true,margin:0,valign:"middle",
      fontFace:F,fontSize:13,bold:true,color:MUTED});
    s.addShape(pres.ShapeType.rightArrow,{x:M+3.60,y:y+0.235,w:0.30,h:0.17,
      fill:{color:BRASS},line:{type:"none"}});
    s.addText(p[1],{x:M+4.04,y,w:2.50,h:rh,isTextBox:true,margin:0,valign:"middle",
      fontFace:F,fontSize:13.5,bold:true,color:BRICK});
    s.addText(p[2],{x:M+6.60,y,w:CW-6.86,h:rh,isTextBox:true,margin:0,valign:"middle",
      fontFace:F,fontSize:12,color:INK});
  });
  notes(s,"【課題を立ち上げるための切り札です】\n\n【出し方】子どもの案が黒板に5〜6個たまったところで出します。左の列を指しながら「これ、○○さんが言ったやつだよね」と、実際の発言者の名前を呼んでください。自分の考えが歴史と一致していた、という経験が課題への当事者意識を生みます。\n\n「実はね。きみたちが今言ったこと、全部、本当に起きてるんだ。しかも、ほとんど50年くらいの間に」\n\n【運用のコツ】\n・子どもから出なかった行だけを教師が補う。2〜3行は教師が足す前提で構いません\n・時間がなければ上から3行だけで十分機能します\n・年号は覚えさせません。「同じ時期に集中して起きた」ことだけ押さえます\n\n【次への発問】「じゃあ、これで問題は解決したのかな？」");
}
{
  const s = slide(true);
  phase(s,4,"課題の設定　｜　次の時間からの学習課題",true);
  s.addText("学習課題",{x:M,y:1.58,w:CW,h:0.40,isTextBox:true,margin:0,
    fontFace:F,fontSize:14,bold:true,color:BRASS,charSpacing:3});
  s.addText("産業革命が生んだ問題を、\n人類はどう解決しようとしたのか。\nそして、解決できたのか。",
    {x:M,y:2.16,w:CW,h:2.75,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:35,bold:true,color:WHITE,lineSpacing:57});
  card(s,M,5.24,CW,1.06,true,INK2);
  s.addText([{text:"この課題は次の時間（第8時）で本格的に追究します。",options:{color:"E3E4E6"}},
             {text:"　今日は、ここから見えてくる「あるもの」まで進みます。",options:{bold:true,color:BRASS}}],
    {x:M+0.32,y:5.24,w:CW-0.64,h:1.06,isTextBox:true,margin:0,valign:"middle",fontFace:F,fontSize:15});
  notes(s,"【課題の立ち上がり方】\nこの課題は教師が用意したものではなく、子どもの発言（20枚目）を歴史的事実（21枚目）と重ねた結果として立ち上がっています。この順序を守れば、課題設定は自然に成立します。\n\n【板書】本時の問いの下に書き加えます。第8時の冒頭でこの板書を再現して始めると、単元がつながります。\n\n【2時間扱いにする場合】ここで1時間目を終えます。振り返りは「今日出た解決策の中で、一番効きそうなのはどれか。理由も」。2時間目は次のスライドから。\n\n【1時間で通す場合】「じゃあ、この6つをよく見てほしい。実は全部に共通していることがある」と言って次へ。ここまでで30分が目安。");
}

/* ============================================================
   23　資本主義
   ============================================================ */
{
  const s = slide(false);
  phase(s,5,"展開3　30-35分　｜　全体",false);
  title(s,"6つの解決策に、共通していること",false,30);
  sub(s,"どれも「やらなかったこと」がある");

  const cw=(CW-0.40)/2;
  card(s,M,2.14,cw,2.44,false,PAPER);
  s.addText("やったこと",{x:M+0.26,y:2.32,w:cw-0.52,h:0.38,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:15,bold:true,color:BRASS});
  s.addText("・働く時間にルールをつくった\n・子どもを工場から学校へ移した\n・労働者が団結できるようにした\n・選挙権を広げようとした\n・街の衛生を国の仕事にした",
    {x:M+0.26,y:2.78,w:cw-0.52,h:1.62,isTextBox:true,margin:0,fontFace:F,fontSize:13,color:INK,lineSpacing:20});
  card(s,M+cw+0.40,2.14,cw,2.44,false,PAPER);
  s.addText("やらなかったこと",{x:M+cw+0.66,y:2.32,w:cw-0.52,h:0.38,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:15,bold:true,color:BRICK});
  s.addText("・工場をなくさなかった\n・機械を捨てなかった\n・「もうけを求めて競争する」\n　という仕組みは、やめなかった",
    {x:M+cw+0.66,y:2.78,w:cw-0.52,h:1.62,isTextBox:true,margin:0,fontFace:F,fontSize:13,color:INK,lineSpacing:20});

  card(s,M,4.86,CW,1.22,false,INK);
  s.addText("つまり ―　ルールを足しただけで、仕組みそのものは残した。",
    {x:M+0.32,y:4.86,w:CW-0.64,h:1.22,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:24,bold:true,color:WHITE});
  notes(s,"【ねらい】個別の解決策から、その背後にある社会の仕組みへ視点を引き上げる。\n\n「よく見て。6つとも、あることをやってない。何をやってないと思う？」\n子どもからは「工場をなくしてない」が出やすいです。出たら大きく褒めて、「そう。誰も工場をやめようとは言わなかった。機械も捨てなかった。じゃあ、その残した仕組みって、なんて名前だと思う？」\n\n【出ないときの手立て】「もし工場が全部なくなったら、きみの服はどうなる？」→「高くなる」「なくなる」→「そう。だから誰もやめられなかったんだ」\n\n2分。ここは速く通す。");
}
{
  const s = slide(true);
  phase(s,5,"展開3　30-35分　｜　全体",true);
  title(s,"その残した仕組みを、こう呼ぶ",true,29);
  s.addText("資本主義",{x:M,y:1.62,w:CW,h:0.88,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:44,bold:true,color:BRASS});

  const cw=(CW-0.80)/3;
  const defs=[["① 私有","工場や機械を、国ではなく個人や会社が持つ"],
              ["② 利潤","もうけを増やすことを目的に、生産する"],
              ["③ 競争","市場での自由な競争が、社会を動かす"]];
  defs.forEach((d,i)=>{
    const x=M+i*(cw+0.40);
    card(s,x,2.70,cw,1.88,true,INK2);
    s.addText(d[0],{x:x+0.26,y:2.88,w:cw-0.52,h:0.46,isTextBox:true,margin:0,valign:"middle",
      fontFace:F,fontSize:22,bold:true,color:BRASS});
    s.addText(d[1],{x:x+0.26,y:3.40,w:cw-0.52,h:1.00,isTextBox:true,margin:0,
      fontFace:F,fontSize:14,color:"E3E4E6",lineSpacing:22});
  });
  card(s,M,4.86,CW,1.22,true,"2C2F36");
  s.addText("産業革命は、この仕組みを世界中に広げた。私たちは今も、この中で暮らしている。",
    {x:M+0.32,y:4.86,w:CW-0.64,h:1.22,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:19,bold:true,color:WHITE});
  notes(s,"【押さえ方】定義を書き写させます（1分）。3つのキーワード（私有・利潤・競争）だけで十分。\n\n【中学2年への配慮】「資本主義」は公民的分野で本格的に扱う語です。ここでは厳密な経済学的定義ではなく、「工場を個人が持ち、もうけを求めて競争する仕組み」という素朴な理解で構いません。\n\n「今日はじめて出てきた言葉。3年生の公民でもう一回ちゃんとやるから、今日は3つだけ覚えて。私有・利潤・競争」\n\n※社会主義（マルクス／20世紀の試み／1991年のソ連解体）に触れる場合はここで1分。触れなくても本時は成立します。時間を見て判断してください。");
}

/* ============================================================
   24　主発問
   ============================================================ */
{
  const s = slide(true);
  phase(s,6,"主発問　35-45分　｜　個人 → 全体",true);
  s.addText("主発問",{x:M,y:1.16,w:CW,h:0.38,isTextBox:true,margin:0,
    fontFace:F,fontSize:14,bold:true,color:BRASS,charSpacing:3});
  s.addText("結局、資本主義は\n人類を幸せにしたのだろうか。",
    {x:M,y:1.70,w:CW,h:1.90,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:41,bold:true,color:WHITE,lineSpacing:60});

  const lx=M+0.60, lw=CW-1.20, ly=4.42;
  s.addShape(pres.ShapeType.line,{x:lx,y:ly,w:lw,h:0,line:{color:"5A5E67",width:2.5}});
  for(let i=0;i<=10;i++){
    const cx = lx + (lw/10)*i, big = (i%5===0);
    s.addShape(pres.ShapeType.ellipse,{x:cx-(big?0.11:0.06),y:ly-(big?0.11:0.06),
      w:big?0.22:0.12,h:big?0.22:0.12,fill:{color:big?BRASS:"5A5E67"},line:{type:"none"}});
    if(big) s.addText(String(i),{x:cx-0.30,y:ly+0.20,w:0.60,h:0.30,isTextBox:true,margin:0,
      align:"center",fontFace:F,fontSize:13,bold:true,color:BRASS});
  }
  s.addText("幸せにしていない",{x:lx-0.55,y:ly-0.62,w:2.40,h:0.34,isTextBox:true,margin:0,
    fontFace:F,fontSize:14,bold:true,color:BRICK});
  s.addText("幸せにした",{x:lx+lw-1.85,y:ly-0.62,w:2.40,h:0.34,isTextBox:true,margin:0,
    align:"right",fontFace:F,fontSize:14,bold:true,color:BRASS});
  s.addText("まず一人で　①自分の数字を決める　②理由を一言、ワークシートに書く　（2分）",
    {x:M,y:5.42,w:CW,h:0.44,isTextBox:true,margin:0,valign:"middle",fontFace:F,fontSize:15,color:"E3E4E6"});
  notes(s,"【進め方】\n①個人で数字と理由を書く（2分）。ここを飛ばさないこと。先に全体討論を始めると、声の大きい子の意見に引きずられます\n②黒板の数直線にネームプレートを貼らせ、学級の分布を可視化（1分）\n③分布の両端と中央から1人ずつ指名して理由を聞く（3分）\n④次の2枚のゆさぶり資料を出して、もう一度考えさせる（4分）\n\n【指名の順序】中間の子から当てると場が締まりません。端の子（0〜2、8〜10）から当てて対立軸を明確にし、そのあと中間の子に「両方わかるという人は？」と振ると、多面的な見方が言語化されます。\n\n【教師の立場】教師は自分の答えを言いません。最後まで言わないでください。まとめで示すのは「功罪の両面がある」という見方であって、幸せにしたかどうかの答えではありません。");
}

/* ============================================================
   25-27　ゆさぶり
   ============================================================ */
{
  const s = slide(false);
  phase(s,6,"討論のゆさぶり ①　｜「幸せにした」側の資料",false);
  title(s,"この200年で、世界はこう変わった",false,30);

  s.addText("極度の貧困のもとで\n暮らす人の割合",{x:M,y:1.62,w:CW-4.60,h:0.44,isTextBox:true,margin:0,
    fontFace:F,fontSize:12.5,color:MUTED});
  s.addChart(pres.ChartType.bar,[{
    name:"割合（%）",
    labels:["1820年","1910年","1950年","1990年","2019年"],
    values:[76,66,55,38,9]
  }],{
    x:M,y:2.14,w:CW-4.60,h:3.42,
    barDir:"col",chartColors:[BRASS],barGapWidthPct:90,
    showLegend:false,showTitle:false,
    showValue:true,dataLabelPosition:"outEnd",
    dataLabelColor:INK,dataLabelFontFace:F,dataLabelFontSize:14,
    dataLabelFormatCode:'0"%"',
    catAxisLabelColor:MUTED,valAxisLabelColor:MUTED,
    catAxisLabelFontFace:F,valAxisLabelFontFace:F,
    catAxisLabelFontSize:12,valAxisLabelFontSize:11,
    valGridLine:{color:LINE,size:1},catGridLine:{style:"none"},
    valAxisMinVal:0,valAxisMaxVal:90
  });
  const px=W-M-4.30;
  card(s,px,2.14,4.30,1.66,false,PAPER);
  s.addText("30歳 → 73歳",{x:px,y:2.28,w:4.30,h:0.74,isTextBox:true,margin:0,align:"center",valign:"middle",
    fontFace:F,fontSize:30,bold:true,color:INK});
  s.addText("世界の平均寿命\n1800年ごろ 約30歳 → 2019年 約73歳",
    {x:px+0.20,y:3.04,w:3.90,h:0.62,isTextBox:true,margin:0,align:"center",
     fontFace:F,fontSize:12,color:MUTED,lineSpacing:17});
  card(s,px,3.98,4.30,1.72,false,INK);
  s.addText([{text:"問い直し\n",options:{bold:true,color:BRASS}},
             {text:"産業革命がなかったら、\nこの200年で生まれた人の多くは、\n大人になれなかったかもしれない。",options:{color:WHITE}}],
    {x:px+0.30,y:3.98,w:3.70,h:1.72,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:14,lineSpacing:22});
  s.addText("出典：世界銀行／Our World in Data による推計（概数）。貧困率は2011年購買力平価で1日1.90ドル未満を基準とする。",
    {x:M,y:H-0.62,w:CW,h:0.30,isTextBox:true,margin:0,fontFace:F,fontSize:10,color:MUTED});
  notes(s,"【使い方】討論が「幸せにしていない」側に偏ったときに出します。偏っていなければ次のスライドと両方まとめて出して構いません。\n\n「ちょっと待って。この資料を見てから、もう一度考えて」\n\n【押さえどころ】貧困率76%→9%は、産業革命に始まる経済成長がもたらした最大の成果です。「豊かになったのは一部の金持ちだけ」という素朴な理解を揺さぶります。\n\n【留意】このグラフは「割合」です。世界人口が大きく増えているため、実数の減り方は割合ほど劇的ではありません。鋭い子が指摘したら褒めてください。\n\n【留意2】貧困率の低下すべてが資本主義の成果だと断定しないでください。医療・公衆衛生・国際協力など複数の要因があります。「産業革命が始めた経済成長が、大きな要因の一つ」が正確な言い方です。");
}
{
  const s = slide(false);
  phase(s,6,"討論のゆさぶり ②　｜「幸せにしていない」側の資料",false);
  title(s,"同じ200年で、こうもなった",false,30);

  const cw=(CW-0.56)/3;
  const facts=[["格差","世界の富の\n約半分","上位1%の人々が保有している"],
               ["地球","約 1.5℃","産業革命前とくらべた世界の平均気温の上昇"],
               ["子ども","約1億6千万人","今も世界で児童労働に従事している子どもの数"]];
  facts.forEach((f,i)=>{
    const x=M+i*(cw+0.28);
    card(s,x,2.14,cw,2.30,false,PAPER);
    s.addText(f[0],{x:x+0.26,y:2.32,w:cw-0.52,h:0.36,isTextBox:true,margin:0,valign:"middle",
      fontFace:F,fontSize:14,bold:true,color:BRICK});
    s.addText(f[1],{x:x+0.26,y:2.72,w:cw-0.52,h:0.90,isTextBox:true,margin:0,valign:"middle",
      fontFace:F,fontSize:28,bold:true,color:INK,lineSpacing:34});
    s.addText(f[2],{x:x+0.26,y:3.66,w:cw-0.52,h:0.62,isTextBox:true,margin:0,
      fontFace:F,fontSize:12,color:MUTED,lineSpacing:17});
  });
  card(s,M,4.62,CW,1.44,false,INK);
  s.addText([{text:"そして　",options:{bold:true,color:BRICK,fontSize:20}},
             {text:"きみが今着ているその服も、遠い国の、とても安い賃金の労働でつくられているかもしれない。",options:{color:WHITE,fontSize:20}}],
    {x:M+0.32,y:4.62,w:CW-0.64,h:1.44,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,bold:true,lineSpacing:30});
  s.addText("出典：格差＝クレディ・スイス／UBS グローバル・ウェルス・レポート、気温＝IPCC 第6次評価報告書、児童労働＝ILO・ユニセフ（2020年）。いずれも概数。",
    {x:M,y:H-0.62,w:CW,h:0.30,isTextBox:true,margin:0,fontFace:F,fontSize:10,color:MUTED});
  notes(s,"【最重要】最下段の一文が本時の設計の要です。導入で扱った「自分の服」がここで戻ってきます。トンネルの入口と出口がつながる瞬間です。必ず間をとって、静かに言ってください。\n\n「今日の最初に、着ている服の値段を聞いたよね。……安かったよね。なんで安いんだろう」\n\n【留意】特定の国・企業・ブランドを名指ししないでください。また「だから安い服を買うのは悪い」という結論に誘導しないこと。消費者個人の責任に落とすと、社会の仕組みを考える授業になりません。「仕組みの中に自分もいる」という気づきに留めます。\n\n【留意2】児童労働の1億6千万人はILO・ユニセフの2020年推計です。産業革命の直接の帰結ではなく、現代のグローバル経済の問題として扱ってください。");
}
{
  const s = slide(true);
  phase(s,6,"討論のゆさぶり ③　｜　ペア 30秒",true);
  s.addText("じゃあ、やめられる？",{x:M,y:1.66,w:CW,h:1.10,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:46,bold:true,color:WHITE});
  s.addText("今日ここに来るまでに、きみが使ったもの",{x:M,y:2.92,w:CW,h:0.38,isTextBox:true,margin:0,
    fontFace:F,fontSize:14,color:MUTED_D});

  const words=["電気","スマホ","電車・バス","服","給食","水道","教科書","薬"];
  const bw=(CW-0.28*3)/4, bh=0.72;
  words.forEach((w,i)=>{
    const col=i%4,row=Math.floor(i/4);
    const x=M+col*(bw+0.28), y=3.46+row*(bh+0.22);
    card(s,x,y,bw,bh,true,INK2);
    s.addText(w,{x,y,w:bw,h:bh,isTextBox:true,margin:0,align:"center",valign:"middle",
      fontFace:F,fontSize:19,bold:true,color:BRASS});
  });
  card(s,M,5.68,CW,0.86,true,"2C2F36");
  s.addText("この中に、産業革命なしで用意できるものは、いくつある？",
    {x:M+0.32,y:5.68,w:CW-0.64,h:0.86,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:18,bold:true,color:WHITE});
  notes(s,"【ねらい】結論への最後の一手。功罪の議論を「では、なくせるのか」という現実の地平に引き戻す。\n\nペアで30秒。「一つでもいい。産業革命なしで用意できるもの、ある？」\n……ほぼ出ません。出なかったという事実そのものが、まとめの説得力になります。\n\n【出た場合の受け方】「教科書は昔もあった」等の発言が出たら、「たしかに本はあった。でも、全員が1冊ずつ持てるようになったのはいつからだろう」と返します。「量」の問題に焦点化すると、産業革命の意味が際立ちます。\n\n1分半。ここまでで45分。");
}

/* ============================================================
   28-30　まとめ
   ============================================================ */
{
  const s = slide(false);
  phase(s,7,"まとめ　45-50分　｜　個人（振り返り）",false);
  title(s,"産業革命が、私たち人類にもたらしたもの",false,30);

  const cw=(CW-0.40)/2;
  card(s,M,2.10,cw,2.74,false,PAPER);
  s.addText("功",{x:M+0.28,y:2.30,w:1.00,h:0.56,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:32,bold:true,color:BRASS});
  s.addText("人類ははじめて、貧しさから\n抜け出す力を手に入れた。",
    {x:M+0.28,y:2.92,w:cw-0.56,h:0.80,isTextBox:true,margin:0,
     fontFace:F,fontSize:17,bold:true,color:INK,lineSpacing:26});
  s.addText("値段が下がり、量が増え、遠くまで運べるようになった。\nより多くの人が生まれ、生き延び、長く生きられるようになった。",
    {x:M+0.28,y:3.80,w:cw-0.56,h:0.86,isTextBox:true,margin:0,
     fontFace:F,fontSize:13,color:MUTED,lineSpacing:20});

  card(s,M+cw+0.40,2.10,cw,2.74,false,PAPER);
  s.addText("罪",{x:M+cw+0.68,y:2.30,w:1.00,h:0.56,isTextBox:true,margin:0,valign:"middle",
    fontFace:F,fontSize:32,bold:true,color:BRICK});
  s.addText("その代金を払ったのは、\n別の誰かだった。",
    {x:M+cw+0.68,y:2.92,w:cw-0.56,h:0.80,isTextBox:true,margin:0,
     fontFace:F,fontSize:17,bold:true,color:INK,lineSpacing:26});
  s.addText("工場と炭鉱で働いた子ども。17歳で亡くなった労働者。\n手仕事を奪われたインドの職人。綿花を摘んだ奴隷。\nそして、これから生きる人たちの地球。",
    {x:M+cw+0.68,y:3.80,w:cw-0.56,h:0.86,isTextBox:true,margin:0,
     fontFace:F,fontSize:13,color:MUTED,lineSpacing:20});

  card(s,M,5.12,CW,0.98,false,INK);
  s.addText("この二つは、別々に起きたことではない。同じ一つのことの、表と裏だった。",
    {x:M+0.32,y:5.12,w:CW-0.64,h:0.98,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:19,bold:true,color:WHITE});
  notes(s,"【まとめの原則】教師が「幸せにした／していない」の答えを言わないこと。示すのは「功と罪の両面がある」「その二つは切り離せない」という見方です。\n\n【最下段の一文が、本時の中心的な学習内容です】\n安い服（功）と、子どもの労働（罪）は、同じ「大量生産」という一つのことから出ています。だから片方だけを取ることができない。ここが多面的・多角的な考察の到達点です。\n\n【板書】黒板の左に功、右に罪を整理した状態で、中央下にこの一文を書きます。");
}
{
  const s = slide(true);
  phase(s,7,"まとめ　45-50分　｜　本時の結論",true);
  s.addText("産業革命は人類に、\n「豊かさ」と「請求書」を\n同時に渡した。",
    {x:M,y:1.16,w:CW,h:2.44,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:34,bold:true,color:WHITE,lineSpacing:54});
  card(s,M,3.76,CW,1.20,true,INK2);
  s.addText("それでも私たちは、もうこれなしでは生きられない。",
    {x:M+0.32,y:3.76,w:CW-0.64,h:1.20,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,fontSize:26,bold:true,color:BRASS});
  s.addText([{text:"だから、問いはこうなる。\n",options:{color:MUTED_D,fontSize:15}},
             {text:"「やめるか、続けるか」ではなく、「どう作り直すか」。",options:{color:WHITE,fontSize:27,bold:true}}],
    {x:M,y:5.12,w:CW,h:1.26,isTextBox:true,margin:0,valign:"middle",fontFace:F,lineSpacing:40});
  notes(s,"【本時の結論】ゆっくり、間をとって。\n\n「産業革命は、人類に豊かさをくれた。同時に、請求書も渡した。……その請求書は、まだ払い終わっていない。格差も、地球の温度も、今のきみたちの問題として残っている」\n「でもね。今日わかったとおり、私たちはもうこれなしでは生きられない。服も、電気も、薬も」\n「だから、考えるべきことは『やめるか続けるか』じゃない。『どう作り直すか』なんだ」\n\n【この結論の意義】\n功罪の両論併記で終わらせず、かといって断定もせず、子どもを当事者の位置に置いたまま次の学習へ渡します。単元を貫く問いへの最終的な答えは、単元のまとめ（第10時）で子ども自身が出します。本時はその材料を全部そろえた、という位置づけです。\n\n【時間】2分。ここは絶対に急がないでください。時間が足りなければ27枚目のペア活動を削ってでもここを確保します。");
}
{
  const s = slide(false);
  phase(s,7,"振り返り　｜　個人（ワークシート）",false);
  title(s,"振り返りを書こう",false,30);
  sub(s,"3分　／　どれか一つでよい");

  const cw=(CW-0.56)/3;
  const qs=[["視点1　自分の変化","授業の前と後で、自分の立場（数直線の数字）は動いた？　動いたなら、動かした資料はどれ？"],
            ["視点2　切り離せるか","産業革命の「光」と「影」は、切り離せると思う？　思う・思わない、どちらでもよい。理由を書こう。"],
            ["視点3　これから","「どう作り直すか」。今のきみに、何か一つ思いつくことはある？"]];
  qs.forEach((q,i)=>{
    const x=M+i*(cw+0.28);
    card(s,x,2.16,cw,2.12,false,PAPER);
    s.addText(q[0],{x:x+0.26,y:2.34,w:cw-0.52,h:0.38,isTextBox:true,margin:0,valign:"middle",
      fontFace:F,fontSize:15,bold:true,color:BRICK});
    s.addText(q[1],{x:x+0.26,y:2.78,w:cw-0.52,h:1.28,isTextBox:true,margin:0,
      fontFace:F,fontSize:13,color:INK,lineSpacing:20});
  });
  card(s,M,4.60,CW,1.48,false,INK);
  s.addText([{text:"次の時間　第8時　",options:{color:BRASS,bold:true,fontSize:14}},
             {text:"働く人を守ったのは、誰か　― 労働問題と社会主義\n",options:{color:WHITE,bold:true,fontSize:21}},
             {text:"今日みんなで立てた課題「産業革命が生んだ問題を、人類はどう解決しようとしたのか」を、腰を据えて追究します。",options:{color:MUTED_D,fontSize:13}}],
    {x:M+0.32,y:4.60,w:CW-0.64,h:1.48,isTextBox:true,margin:0,valign:"middle",
     fontFace:F,lineSpacing:28});
  notes(s,"【振り返りの設計】\n視点1が「主体的に学習に取り組む態度」の主たる評価材料になります。自分の考えの変容を、資料を根拠に説明できているかを見取ってください。「変わらなかった」も、理由が書けていれば十分に評価できます。\n\n視点2が「思考・判断・表現」の評価材料です。功と罪を関連付けて考察できているかを見ます。\n\n視点3は書ける子だけで構いません。単元のまとめ（第10時）への布石です。\n\n【時間が足りないとき】視点1だけを書かせて回収します。1分で書けます。");
}

/* ============================================================
   31　出典一覧（教師用）
   ============================================================ */
{
  const s = slide(false);
  phase(s,0,"教師用　｜　画像の出典と入手先",false);
  title(s,"使用した絵の一覧（すべてパブリックドメイン）",false,27,0.86,0.52);
  s.addText("いずれも著作権の保護期間が満了した作品。ウィキメディア・コモンズ（commons.wikimedia.org）で「検索語」を入れると見つかります。",
    {x:M,y:1.44,w:CW,h:0.34,isTextBox:true,margin:0,fontFace:F,fontSize:11.5,color:MUTED});

  const ns = Object.keys(IMAGES).map(Number).sort((a,b)=>a-b);
  const cw=(CW-0.32)/2, rh=0.50, rg=0.05;
  ns.forEach((n,i)=>{
    const m=IMAGES[n];
    const col = i<6?0:1, idx=i%6;
    const x = M + col*(cw+0.32), y = 1.92 + idx*(rh+rg);
    card(s,x,y,cw,rh,false,PAPER);
    s.addShape(pres.ShapeType.ellipse,{x:x+0.16,y:y+0.15,w:0.22,h:0.22,fill:{color:BRASS},line:{type:"none"}});
    s.addText(String(n),{x:x+0.16,y:y+0.15,w:0.22,h:0.22,isTextBox:true,margin:0,align:"center",valign:"middle",
      fontFace:F,fontSize:9,bold:true,color:INK});
    s.addText(m.t+"（"+m.y+"）",{x:x+0.48,y:y+0.06,w:cw-0.66,h:0.24,isTextBox:true,margin:0,valign:"middle",
      fontFace:F,fontSize:11.5,bold:true,color:INK});
    s.addText("検索語： "+m.q,{x:x+0.48,y:y+0.28,w:cw-0.66,h:0.22,isTextBox:true,margin:0,valign:"middle",
      fontFace:F,fontSize:9.5,color:MUTED});
  });
  card(s,M,5.40,CW,1.30,false,INK);
  s.addText([{text:"画像のはめ方　",options:{bold:true,color:BRASS,fontSize:14}},
             {text:"点線の枠に画像をドラッグ＆ドロップし、枠を右クリック →「最背面へ移動」または削除。\n",options:{color:WHITE,fontSize:13}},
             {text:"または　",options:{bold:true,color:BRASS,fontSize:14}},
             {text:"「画像」フォルダに 01.jpg 〜 12.jpg として保存し、生成スクリプトを再実行すると自動ではめ込まれます。",options:{color:WHITE,fontSize:13}}],
    {x:M+0.32,y:5.40,w:CW-0.64,h:1.30,isTextBox:true,margin:0,valign:"middle",fontFace:F,lineSpacing:22});
  notes(s,"【著作権について】\n掲載した12点はいずれも19世紀以前の作品で、著作権の保護期間が満了しています（パブリックドメイン）。授業での投影・配付・印刷いずれも問題ありません。\n\nただし、写真として撮り直されたものには撮影者の権利が及ぶ場合があります。ウィキメディア・コモンズで各ファイルのライセンス欄に「Public domain」または「PD-old」と表示されているものを選んでください。\n\n【入手先】\n・ウィキメディア・コモンズ　commons.wikimedia.org\n・アメリカ議会図書館　loc.gov（Free to Use のコレクション）\n・大英図書館 Flickr Commons\n\n【はめ方】\nPowerPoint では、点線の枠の上に画像ファイルをドラッグ＆ドロップし、枠を選んで削除するのが最も速いです。画像は「図の形式 → トリミング → 塗りつぶし」で枠に合わせられます。");
}

/* ============================================================
   出力 ＋ 画像リストと取得スクリプトを同時に生成
   ============================================================ */
const outfile = path.join(OUTDIR, "令和8年度_第2学年_社会科_産業革命_授業用スライド（絵入り）.pptx");
pres.writeFile({ fileName: outfile }).then(f => {
  console.log("written:", f);
  if (missing.length) console.log("画像未配置:", missing.join(", "));

  // --- 画像リスト.md ---
  let md = "# 産業革命の授業スライド　画像一覧\n\n";
  md += "授業用スライド（絵入り）で使う12点。**いずれも19世紀以前の作品で、著作権の保護期間が満了したパブリックドメイン**です。\n";
  md += "授業での投影・印刷・配付に制限はありません。\n\n";
  md += "## 入手のしかた\n\n";
  md += "1. 下の「検索」リンクを開く（ウィキメディア・コモンズの画像検索が開きます）\n";
  md += "2. 説明に合う絵を選ぶ。ライセンス欄が **Public domain / PD-old** のものを選ぶこと\n";
  md += "3. 「原寸大の画像を表示」から保存し、`画像/01.jpg` … `画像/12.jpg` という名前で入れる\n";
  md += "4. `node tools/産業革命_授業用スライド_生成.js` を実行すると、枠が絵に置き換わります\n\n";
  md += "PowerPoint で直接はめてもかまいません。点線の枠に画像をドラッグ＆ドロップし、枠を削除してください。\n\n";
  md += "## 一覧\n\n";
  ns_all().forEach(n => {
    const m = IMAGES[n];
    md += "### " + String(n).padStart(2,"0") + "　" + m.t + "（" + m.y + "）\n\n";
    md += "- **描いた人・出典**：" + m.a + "\n";
    md += "- **原題**：" + m.w + "\n";
    md += "- **何が写っているか**：" + m.d + "\n";
    md += "- **ファイル名の目安**：" + m.f + "\n";
    md += "- **検索**：[" + m.q + "](" + searchURL(n) + ")\n";
    md += "- **保存先**：`画像/" + String(n).padStart(2,"0") + ".jpg`\n\n";
  });
  md += "## 使いどころ\n\n";
  md += "| 画像 | スライド | 役割 |\n|---|---|---|\n";
  md += "| 1 | 1・19枚目 | 表紙と山場。一枚に功罪の両方が写っている決定的な絵 |\n";
  md += "| 2 | 3枚目 | 導入。産業革命の前の暮らし |\n";
  md += "| 3 | 5枚目 | 導入。すべての始まり |\n";
  md += "| 4・5 | 6枚目 | 産業革命の3要素（動力・工場） |\n";
  md += "| 6 | 9枚目 | 光。速さ |\n";
  md += "| 7 | 10枚目 | 光。豊かさの象徴 |\n";
  md += "| 8 | 13枚目 | 影。スラム |\n";
  md += "| 9・10 | 14枚目 | 影。子どもの労働（本時で最も重い2枚） |\n";
  md += "| 11・12 | 17枚目 | 影。世界への影響 |\n";
  fs.writeFileSync(path.join(OUTDIR, "画像リスト.md"), md);
  console.log("written: 画像リスト.md");

  // --- 取得用リンク（ブラウザで開くだけの HTML） ---
  let html = "<!doctype html><meta charset=\"utf-8\"><title>産業革命 授業スライド 画像さがし</title>";
  html += "<style>body{font-family:'Meiryo',sans-serif;max-width:900px;margin:40px auto;padding:0 20px;line-height:1.7;color:#23252A}";
  html += "h1{font-size:22px;border-bottom:2px solid #D9A441;padding-bottom:8px}";
  html += ".c{background:#F1F1EF;border-radius:8px;padding:14px 18px;margin:12px 0}";
  html += ".n{display:inline-block;background:#D9A441;color:#23252A;border-radius:50%;width:24px;height:24px;";
  html += "text-align:center;line-height:24px;font-weight:bold;font-size:13px;margin-right:8px}";
  html += ".t{font-weight:bold}.d{font-size:13px;color:#6E7178;margin:6px 0}";
  html += "a{color:#A63D2E}</style>";
  html += "<h1>産業革命の授業スライド　画像さがし</h1>";
  html += "<p>リンクを開いて絵を選び、<code>画像/01.jpg</code> … <code>画像/12.jpg</code> として保存してください。";
  html += "ライセンス欄が <b>Public domain / PD-old</b> のものを選ぶこと。</p>";
  ns_all().forEach(n => {
    const m = IMAGES[n];
    html += "<div class=\"c\"><span class=\"n\">"+n+"</span><span class=\"t\">"+m.t+"（"+m.y+"）</span>";
    html += "<div class=\"d\">"+m.d+"<br>出典："+m.a+"　／　原題："+m.w+"</div>";
    html += "<a href=\""+searchURL(n)+"\" target=\"_blank\">コモンズで探す →</a>　";
    html += "<span class=\"d\">保存先： 画像/"+String(n).padStart(2,"0")+".jpg</span></div>";
  });
  fs.writeFileSync(path.join(OUTDIR, "画像さがし.html"), html);
  console.log("written: 画像さがし.html");
});

function ns_all(){ return Object.keys(IMAGES).map(Number).sort((a,b)=>a-b); }
