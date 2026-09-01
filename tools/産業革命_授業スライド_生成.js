const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";            // 13.333 x 7.5
pres.author = "社会科";
pres.title  = "産業革命は、人類を幸せにしたのか";

const W = 13.333, H = 7.5, M = 0.62, CW = W - M * 2;

// ---- palette: 産業革命（すす・れんが・真鍮のガス灯） ----
const INK   = "23252A";   // すす色（主色）
const INK2  = "33363D";   // カード用の少し明るいすす色
const BRICK = "A63D2E";   // れんが（＝影／罪）
const BRASS = "D9A441";   // 真鍮・ガス灯（＝光／功）
const PAPER = "F1F1EF";
const WHITE = "FFFFFF";
const MUTED = "6E7178";
const MUTED_D = "9DA1A9";
const LINE  = "DFDFDC";

const F = "Meiryo";

const sh = () => ({ type: "outer", color: "000000", blur: 10, offset: 2, angle: 90, opacity: 0.10 });

function slide(dark) {
  const s = pres.addSlide();
  s.background = { color: dark ? INK : WHITE };
  return s;
}

// 意匠モチーフ：丸数字のチップ＋場面名（全スライド共通）
function phase(s, num, label, dark) {
  s.addShape(pres.ShapeType.ellipse, {
    x: M, y: 0.34, w: 0.40, h: 0.40,
    fill: { color: dark ? BRASS : INK }, line: { type: "none" }
  });
  s.addText(String(num), {
    x: M, y: 0.34, w: 0.40, h: 0.40, isTextBox: true, margin: 0,
    align: "center", valign: "middle",
    fontFace: F, fontSize: 14, bold: true, color: dark ? INK : WHITE
  });
  s.addText(label, {
    x: M + 0.54, y: 0.34, w: CW - 0.54, h: 0.40, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 13, bold: true,
    color: dark ? MUTED_D : MUTED, charSpacing: 1
  });
}

function title(s, text, dark, size, y, h) {
  s.addText(text, {
    x: M, y: y === undefined ? 0.92 : y, w: CW, h: h === undefined ? 0.72 : h,
    isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: size || 32, bold: true,
    color: dark ? WHITE : INK, lineSpacing: (size || 32) * 1.25
  });
}

function sub(s, text, dark, y) {
  s.addText(text, {
    x: M, y: y === undefined ? 1.70 : y, w: CW, h: 0.42,
    isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 14, color: dark ? MUTED_D : MUTED
  });
}

function source(s, text, dark) {
  s.addText(text, {
    x: M, y: H - 0.68, w: CW, h: 0.34, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 10.5,
    color: dark ? MUTED_D : MUTED
  });
}

// カード（角丸＋淡い地色＋やわらかい影。※縁のライン飾りは使わない）
function card(s, x, y, w, h, dark, tint) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: tint || (dark ? INK2 : PAPER) },
    line: { type: "none" }, shadow: sh()
  });
}

// 見出し＋本文入りカード
function infoCard(s, x, y, w, h, head, body, accent, dark, tint) {
  card(s, x, y, w, h, dark, tint);
  s.addText(head, {
    x: x + 0.24, y: y + 0.17, w: w - 0.48, h: 0.40, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 15, bold: true, color: accent
  });
  s.addText(body, {
    x: x + 0.24, y: y + 0.62, w: w - 0.48, h: h - 0.80, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, color: dark ? "E3E4E6" : INK, lineSpacing: 19
  });
}

// 大きな数字の見せ場
function stat(s, x, y, w, big, cap, color) {
  s.addText(big, {
    x, y, w, h: 1.15, isTextBox: true, margin: 0, align: "center", valign: "middle",
    fontFace: F, fontSize: 54, bold: true, color
  });
  s.addText(cap, {
    x, y: y + 1.15, w, h: 0.80, isTextBox: true, margin: 0, align: "center",
    fontFace: F, fontSize: 13, color: MUTED, lineSpacing: 18
  });
}

function notes(s, t) { s.addNotes(t); }

/* ============================================================
   PART 0　教師用：単元の位置づけと本時の設計
   ============================================================ */

// --- S1 表紙 ---
{
  const s = slide(true);
  s.addText("令和8年度　第2学年　社会科（歴史的分野）", {
    x: M, y: 1.25, w: CW, h: 0.36, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 14, color: BRASS, charSpacing: 2
  });
  s.addText("産業革命は、\n人類を幸せにしたのか", {
    x: M, y: 1.80, w: CW, h: 2.30, isTextBox: true, margin: 0, valign: "top",
    fontFace: F, fontSize: 46, bold: true, color: WHITE, lineSpacing: 62
  });
  card(s, M, 4.55, CW, 1.42, true, INK2);
  s.addText([
    { text: "【本時】", options: { bold: true, color: BRASS } },
    { text: "産業革命は私たち人類に何をもたらしたのか", options: { bold: true, color: WHITE } }
  ], {
    x: M + 0.34, y: 4.76, w: CW - 0.68, h: 0.42, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 20
  });
  s.addText("単元「欧米の近代化と産業革命」全10時間　第7時／功罪に着目して多面的・多角的に考察する", {
    x: M + 0.34, y: 5.26, w: CW - 0.68, h: 0.40, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 13, color: MUTED_D
  });
  notes(s, "【この資料の使い方】\nスライド1〜4は教師用（授業前の確認用）。児童生徒に投影するのはスライド5から。\n発表者ツールのノート欄に、各場面の時間配分・活動形態・予想される生徒の反応・切り返しを書いてあります。\n\n※数値はすべて概数です。授業前に必ず使用中の教科書・資料集の数値と照合してください。出典はスライド下部に明記しています。");
}

// --- S2 単元の全体像 ---
{
  const s = slide(false);
  phase(s, 0, "教師用｜単元の位置づけ", false);
  title(s, "単元の全体像 ― 本時はここ", false, 30);
  sub(s, "単元を貫く問い「産業革命は、人類を幸せにしたのか」に、毎時の問いで少しずつ迫る");

  const rows = [
    ["1", "単元の課題設定", "産業革命は、人類を幸せにしたのか", false],
    ["2", "市民革命①イギリス", "なぜ王の首をはねてまで議会をとったのか", false],
    ["3", "市民革命②アメリカ", "独立宣言の「平等」に、奴隷は入っていたか", false],
    ["4", "市民革命③フランス", "人権宣言は、誰の権利を守ったのか", false],
    ["5", "産業革命のはじまり", "なぜ「イギリスから」始まったのか", false],
    ["6", "機械と工場が変えた社会", "機械は、人を楽にしたのか", false],
    ["7", "【本時】功罪と資本主義", "産業革命は私たち人類に何をもたらしたのか", true],
    ["8", "労働問題と社会主義", "働く人を守ったのは、誰か", false],
    ["9", "欧米諸国のアジア進出", "豊かさの「材料」は、どこから来たのか", false],
    ["10", "単元のまとめ・振り返り", "私たちは、この仕組みとどう付き合うか", false]
  ];
  const colW = (CW - 0.34) / 2, rowH = 0.86, gap = 0.09;
  rows.forEach((r, i) => {
    const col = i < 5 ? 0 : 1, idx = i % 5;
    const x = M + col * (colW + 0.34), y = 2.24 + idx * (rowH + gap);
    const on = r[3];
    card(s, x, y, colW, rowH, false, on ? INK : PAPER);
    s.addShape(pres.ShapeType.ellipse, {
      x: x + 0.20, y: y + 0.235, w: 0.39, h: 0.39,
      fill: { color: on ? BRASS : "D8D8D4" }, line: { type: "none" }
    });
    s.addText(r[0], {
      x: x + 0.20, y: y + 0.235, w: 0.39, h: 0.39, isTextBox: true, margin: 0,
      align: "center", valign: "middle", fontFace: F, fontSize: 12, bold: true,
      color: on ? INK : "55575C"
    });
    s.addText(r[1], {
      x: x + 0.72, y: y + 0.13, w: colW - 0.94, h: 0.34, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 13.5, bold: true, color: on ? WHITE : INK
    });
    s.addText("「" + r[2] + "」", {
      x: x + 0.72, y: y + 0.45, w: colW - 0.94, h: 0.30, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 11, color: on ? BRASS : MUTED
    });
  });
  notes(s, "単元名は年間指導計画の様式に合わせ、対立軸を含む問いの形にしています（例：「江戸時代は『足かせ』か『土台』か」）。\n\n本時（第7時）は、第5・6時で得た「産業革命とは何か」という知識の上に立ち、価値判断へ踏み込む時間です。第5・6時で機械・蒸気機関・工場制機械工業の基本用語を必ず押さえておいてください。本時では用語の説明に時間を使いません。\n\n第8時（労働問題と社会主義）は、本時の展開2で子どもから出た解決策を、腰を据えて追究する時間として位置づきます。本時の板書を残しておくと接続できます。");
}

// --- S3 本時の設計図（トンネル方式） ---
{
  const s = slide(true);
  phase(s, 0, "教師用｜本時の設計", true);
  title(s, "文脈のつくり方 ―「トンネル方式」", true, 30);
  sub(s, "具体的な一点から入り、深く掘り、一般的な問いへ“抜ける”", true);

  const bw = (CW - 1.10) / 3;
  const blocks = [
    ["入口（せまく・具体）", "今日、自分が着ている\nシャツ 1枚", "一人残らず当事者になれる、\n目の前の“物”から入る", BRASS],
    ["トンネルの中（掘る）", "18世紀の綿布はぜいたく品\n綿糸の値段が約1/12に\n綿花の輸入が約85倍に", "「なぜ安くなった？」を\n数値で掘り下げる", "C9CCD2"],
    ["出口（ひろく・一般）", "産業革命は私たち人類に\n何をもたらしたのか", "具体を掘り切ったところで、\n本時の問いへ抜ける", BRICK]
  ];
  blocks.forEach((b, i) => {
    const x = M + i * (bw + 0.55);
    const y = 2.40, hh = 2.55;
    card(s, x, y, bw, hh, true, INK2);
    s.addText(b[0], {
      x: x + 0.24, y: y + 0.18, w: bw - 0.48, h: 0.34, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 12.5, bold: true, color: b[3]
    });
    s.addText(b[1], {
      x: x + 0.24, y: y + 0.60, w: bw - 0.48, h: 1.20, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 15, bold: true, color: WHITE, lineSpacing: 23
    });
    s.addText(b[2], {
      x: x + 0.24, y: y + 1.86, w: bw - 0.48, h: 0.56, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11.5, color: MUTED_D, lineSpacing: 16
    });
    if (i < 2) {
      s.addShape(pres.ShapeType.rightArrow, {
        x: x + bw + 0.11, y: 3.44, w: 0.34, h: 0.46,
        fill: { color: "4A4E57" }, line: { type: "none" }
      });
    }
  });
  card(s, M, 5.28, CW, 1.06, true, "2C2F36");
  s.addText([
    { text: "この方式のねらい　", options: { bold: true, color: BRASS } },
    { text: "「産業革命とは何か」を先に説明してから問うと、問いが他人事になる。先に“自分の服”という一点へ潜らせ、そこから出られなくしてから広い問いへ出す。出口の問いは、まとめ（スライド27）で再び入口の服に戻る。", options: { color: "E3E4E6" } }
  ], {
    x: M + 0.30, y: 5.44, w: CW - 0.60, h: 0.74, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 12.5, lineSpacing: 18
  });
  notes(s, "トンネル方式の要点は「入口を欲張らないこと」です。導入で産業革命の説明を足したくなりますが、足すほど問いがぼやけます。入口はシャツ1枚だけ。\n\n入口が効くのは、出口と再び結びつくときです。本時は終盤（スライド25「じゃあ、やめられる？」とスライド27のまとめ）で必ず服に戻ります。導入で「この服のことは、授業の最後にもう一度聞くよ」と予告しておくと、子どもが伏線として保持します。");
}

// --- S4 50分の流れ ---
{
  const s = slide(false);
  phase(s, 0, "教師用｜本時の展開と活動形態", false);
  title(s, "50分の流れ ― いつ、誰と考えるか", false, 30);
  sub(s, "「個 → グループ → 全体」を2回まわす。考えが動いた瞬間を全体に返すのが教師の仕事");

  const heads = ["時間", "場面", "発問・活動", "活動形態"];
  const colX = [M, M + 1.05, M + 3.10, M + 9.85];
  const colW2 = [1.05, 2.05, 6.75, CW - 9.23];
  heads.forEach((hd, i) => {
    s.addText(hd, {
      x: colX[i], y: 2.16, w: colW2[i], h: 0.34, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 11.5, bold: true, color: MUTED
    });
  });
  const rows = [
    ["0-7分", "導入", "「今、着ている服。いくらだった？」→ 200年前と比べる", "全体・ペア", BRASS],
    ["7-10分", "問いの設定", "【本時の問い】産業革命は私たち人類に何をもたらしたのか", "全体", INK],
    ["10-22分", "展開1", "資料カード12枚を「光」と「影」に分ける", "グループ（4人）", BRASS],
    ["22-30分", "展開2", "「この影を、当時の人はどう解決しようとした？」→ 課題を立ち上げる", "全体（教師が板書）", BRICK],
    ["30-35分", "展開3", "解決策の共通点 →「資本主義」／立場を決める", "個人 → 全体", INK],
    ["35-45分", "討論", "【主発問】結局、資本主義は人類を幸せにしたのだろうか", "全体（ゆさぶり資料）", BRICK],
    ["45-50分", "まとめ", "功罪の両面 ／「やめるか」ではなく「どう作り直すか」", "個人（振り返り）", INK]
  ];
  const rh = 0.56, rg = 0.06;
  rows.forEach((r, i) => {
    const y = 2.58 + i * (rh + rg);
    card(s, M, y, CW, rh, false, PAPER);
    s.addShape(pres.ShapeType.ellipse, {
      x: M + 0.20, y: y + 0.215, w: 0.13, h: 0.13,
      fill: { color: r[4] }, line: { type: "none" }
    });
    s.addText(r[0], {
      x: M + 0.42, y, w: 0.72, h: rh, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 11.5, bold: true, color: INK
    });
    s.addText(r[1], {
      x: colX[1] + 0.12, y, w: colW2[1], h: rh, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 12.5, bold: true, color: r[4]
    });
    s.addText(r[2], {
      x: colX[2], y, w: colW2[2], h: rh, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 12.5, color: INK
    });
    s.addText(r[3], {
      x: colX[3], y, w: colW2[3] - 0.20, h: rh, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 11.5, color: MUTED
    });
  });
  notes(s, "【活動形態の設計意図】\n展開1をグループにするのは、「光か影か迷う」経験を必ず全員にさせるためです。一人だと迷わずに分類して終わります。4人だと必ず割れます。割れたところが本時の学習内容そのものです。\n\n展開2を全体にするのは、課題を子どもと一緒に立ち上げるためです。ここだけはグループにしないでください。グループにすると出た案が班の中で閉じ、教師が歴史的事実と接続できません。\n\n【2時間扱いにする場合の切れ目】\nスライド18（課題の設定）で1時間目を終え、2時間目をスライド19から始めます。切れ目としてはここが最も自然です。");
}

/* ============================================================
   PART 1　導入（トンネルの入口）
   ============================================================ */

// --- S5 導入の投げかけ ---
{
  const s = slide(true);
  phase(s, 1, "導入　0-7分　｜　全体 → ペア", true);
  s.addText("今、着ている服。\nいくらだった？", {
    x: M, y: 2.05, w: CW, h: 2.20, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 52, bold: true, color: WHITE, lineSpacing: 72
  });
  card(s, M, 4.75, CW, 1.06, true, INK2);
  s.addText([
    { text: "となりの人と　30秒　", options: { bold: true, color: BRASS } },
    { text: "／　値段がわからなければ「高いと思う？安いと思う？」でよい", options: { color: "E3E4E6" } }
  ], {
    x: M + 0.30, y: 4.95, w: CW - 0.60, h: 0.66, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 15
  });
  notes(s, "【ねらい】トンネルの入口。全員が当事者になる一点に潜らせる。\n\n【言い方】「教科書を開く前に、一つ聞かせて。今きみが着ているその服、いくらだった？」\n制服の学校なら「その制服」または「昨日の私服」「今はいている靴下」で代替できます。\n\n【留意点】家庭状況に触れる話題です。個人が特定される言わせ方（「一番高い人は？」等）は避け、「だいたい何千円くらいか」の相場感に留めてください。答えられない子には「高いと思う？安いと思う？」で拾います。\n\n【時間】ここは2分。長引かせない。「安い」という感覚が出れば十分です。次のスライドへ。");
}

// --- S6 いま と 18世紀 ---
{
  const s = slide(false);
  phase(s, 1, "導入　0-7分　｜　全体", false);
  title(s, "その「安さ」は、当たり前ではなかった", false, 30);

  const cw2 = (CW - 0.40) / 2;
  infoCard(s, M, 2.05, cw2, 2.30,
    "いま（2026年）",
    "・Tシャツ1枚　約1,000〜2,000円\n・クローゼットに何着ある？\n・気に入らなければ、買いかえる",
    INK, false, PAPER);
  infoCard(s, M + cw2 + 0.40, 2.05, cw2, 2.30,
    "18世紀なかばのイギリス",
    "・綿の布は、インドからの輸入品。ぜいたく品\n・糸は手で紡ぎ、布は手で織る\n・1着つくるのに、何日もかかる\n・庶民の普段着は、毛や麻のおさがり",
    BRICK, false, PAPER);

  card(s, M, 4.70, CW, 1.42, false, INK);
  s.addText("この200年ほどのあいだに、何が起きたのか。", {
    x: M + 0.34, y: 4.70, w: CW - 0.68, h: 1.42, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 26, bold: true, color: WHITE
  });
  notes(s, "【ねらい】「安い服」が歴史的にはごく最近の現象だと気づかせる。\n\n【言い方】「今きみたちが当たり前だと思っている値段は、人類の歴史の中ではつい最近できたものなんだ。250年前のイギリスでは、綿の服は輸入品のぜいたく品だった」\n\n【留意点】具体的な円換算は示しません。当時の物価を現代の円に換算する作業には無理があり、聞かれたら「正確な換算は難しいけれど、庶民が気軽に買える物ではなかった」と答えてください。次のスライドで、換算の要らない“倍率”のデータを出します。\n\n【時間】2分。");
}

// --- S7 数値のインパクト ---
{
  const s = slide(false);
  phase(s, 1, "導入　0-7分　｜　全体", false);
  title(s, "数字で見ると、こうなる", false, 30);
  sub(s, "イギリスの綿工業に起きたこと");

  const cw3 = (CW - 0.40) / 2;
  card(s, M, 2.22, cw3, 2.55, false, PAPER);
  stat(s, M, 2.48, cw3, "約 1/12", "綿糸1ポンドの値段\n1786年 38シリング → 1832年 3シリング", BRASS);
  card(s, M + cw3 + 0.40, 2.22, cw3, 2.55, false, PAPER);
  stat(s, M + cw3 + 0.40, 2.48, cw3, "約 85倍", "イギリスの綿花輸入量\n1780年 約690万ポンド → 1850年 約5億8,800万ポンド", BRASS);

  card(s, M, 5.10, CW, 1.02, false, INK);
  s.addText([
    { text: "なぜ？　", options: { bold: true, color: BRASS, fontSize: 24 } },
    { text: "値段が1/12になり、量が85倍になるようなことが、なぜ起きたのか。", options: { color: WHITE, fontSize: 19 } }
  ], {
    x: M + 0.34, y: 5.10, w: CW - 0.68, h: 1.02, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 20, bold: true
  });
  source(s, "※数値は概数。イギリス経済史の統計（Deane & Cole ほか）による。授業前に使用中の資料集の数値と照合してください。", false);
  notes(s, "【ねらい】トンネルの一番深いところ。ここで「なぜ」を子どもの内側から出させる。\n\n【言い方】「値段が12分の1。量が85倍。……これ、ふつうのことだと思う？」\n数字を読み上げたら、一拍おいてください。子どもが「機械？」「工場？」と口にすれば成功です。前時（第5・6時）の内容なので、必ず出ます。\n\n【留意点】数値は概数です。使用中の教科書・資料集に別の数値があれば、そちらを優先してください。倍率のオーダー（10分の1程度・数十倍）が伝われば目的は達します。\n\n【時間】2分。");
}

// --- S8 産業革命 ---
{
  const s = slide(true);
  phase(s, 1, "導入　0-7分　｜　全体", true);
  title(s, "その正体が「産業革命」", true, 32);
  sub(s, "18世紀後半、イギリスから始まった　― 前の時間に学んだこと", true);

  const bw2 = (CW - 0.80) / 3;
  const items = [
    ["機械", "紡績機・力織機\n人の手より速く、休まず動く"],
    ["動力", "蒸気機関\n風・水・筋肉の力から解き放たれる"],
    ["工場", "工場制機械工業\n人を1か所に集め、大量に生産する"]
  ];
  items.forEach((it, i) => {
    const x = M + i * (bw2 + 0.40);
    card(s, x, 2.40, bw2, 2.10, true, INK2);
    s.addText(it[0], {
      x: x + 0.28, y: 2.60, w: bw2 - 0.56, h: 0.52, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 26, bold: true, color: BRASS
    });
    s.addText(it[1], {
      x: x + 0.28, y: 3.20, w: bw2 - 0.56, h: 1.06, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 14, color: "E3E4E6", lineSpacing: 22
    });
  });
  card(s, M, 4.86, CW, 1.14, true, "2C2F36");
  s.addText("人類が、はじめて「自分の筋力より大きな力」を、好きなだけ使えるようになった。", {
    x: M + 0.34, y: 4.86, w: CW - 0.68, h: 1.14, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 19, bold: true, color: WHITE
  });
  notes(s, "【ねらい】既習事項の確認。ここは説明せず、子どもに言わせて確認するだけ。\n\n【言い方】「そう、産業革命。前の時間にやったね。3つ、キーワードを言える人？」\n機械・蒸気機関・工場が出れば十分です。出なければ教師が示して確認します。\n\n【時間】1分。ここで説明を始めると導入が崩れます。確認だけして次へ。\n\n【次のスライドへのつなぎ】\n「じゃあ本題。この産業革命が、私たち人類に何をもたらしたのか。今日はこれを考える」と言ってスライド9を出します。");
}

// --- S9 本時の問い ---
{
  const s = slide(true);
  phase(s, 2, "問いの設定　7-10分　｜　全体", true);
  s.addText("本時の問い", {
    x: M, y: 1.55, w: CW, h: 0.42, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, bold: true, color: BRASS, charSpacing: 3
  });
  s.addText("産業革命は\n私たち人類に\n何をもたらしたのか", {
    x: M, y: 2.15, w: CW, h: 3.20, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 50, bold: true, color: WHITE, lineSpacing: 74
  });
  s.addText("「よかった／悪かった」の一言では終わらせない。光と影の両方を、資料で確かめる。", {
    x: M, y: 5.62, w: CW, h: 0.46, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 15, color: MUTED_D
  });
  notes(s, "【ねらい】トンネルの出口。ここまでの具体（服・数字）が、一般的な問いに変わる瞬間。\n\n【板書】この問いを黒板の中央上部に書き、授業の最後まで消さないでください。本時の板書はこの問いを軸に左右へ広がっていきます（板書計画＝スライド29）。\n\n【ワークシート】問いを書き写させ、この時点での予想を一言だけ書かせます（30秒）。授業後に自分の変化を見取るための起点になります。ここを飛ばすと、まとめの振り返りが浅くなります。\n\n【時間】3分。");
}

/* ============================================================
   PART 2　展開1　功罪の整理（グループ）
   ============================================================ */

// --- S10 活動指示 ---
{
  const s = slide(false);
  phase(s, 3, "展開1　10-22分　｜　グループ（4人）", false);
  title(s, "資料カードを「光」と「影」に分けよう", false, 30);
  sub(s, "12分　／　グループに1セット（12枚）　／　机の上を3つの場所に分けて置く");

  const cw4 = (CW - 0.80) / 3;
  const zones = [
    ["光", "人類にとって\nプラスだったこと", BRASS, INK],
    ["迷い", "どちらとも\n言い切れないもの", "B9BCC2", INK],
    ["影", "人類にとって\nマイナスだったこと", BRICK, WHITE]
  ];
  zones.forEach((z, i) => {
    const x = M + i * (cw4 + 0.40);
    card(s, x, 2.30, cw4, 1.90, false, z[2]);
    s.addText(z[0], {
      x, y: 2.48, w: cw4, h: 0.62, isTextBox: true, margin: 0,
      align: "center", valign: "middle", fontFace: F, fontSize: 30, bold: true, color: z[3]
    });
    s.addText(z[1], {
      x, y: 3.14, w: cw4, h: 0.86, isTextBox: true, margin: 0,
      align: "center", fontFace: F, fontSize: 13.5, color: z[3], lineSpacing: 21
    });
  });

  card(s, M, 4.46, CW, 1.66, false, INK);
  s.addText([
    { text: "いちばん大事なのは「迷い」の場所です。\n", options: { bold: true, color: BRASS, fontSize: 19 } },
    { text: "光にも影にも置けるカードが必ずあります。迷ったら無理に決めず、真ん中に置いて「なぜ迷うのか」を班で話してください。あとで全体に聞きます。", options: { color: WHITE, fontSize: 14 } }
  ], {
    x: M + 0.34, y: 4.46, w: CW - 0.68, h: 1.66, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, lineSpacing: 24
  });
  notes(s, "【活動形態】4人グループ。カードは印刷して切り、班に1セット配付（次の2枚のスライドがカードの内容です）。\n\n【手順の指示】\n①12枚を回し読みして、全員が全部読む（4分）\n②光・影・迷いに分ける（5分）\n③「迷い」に置いたカードについて、なぜ迷うのかを話す（3分）\n\n【机間指導で見取ること】\n・迷いの場所が空の班 → 「本当に全部きれいに分かれた？　光4番（人口が3倍）を見て。増えた人は、みんな幸せに暮らせたと思う？」と揺さぶる\n・全部影に寄せた班 → 光1番（服の値段）を指して「じゃあ、今きみが着ている服は？」\n・全部光に寄せた班 → 影2番（平均死亡年齢17歳）を指す\n\n【評価】この場面で「思考・判断・表現」を机間指導で見取ります。分類の結果ではなく、迷った理由を言えるかを見てください。");
}

// --- S11 光のカード ---
{
  const s = slide(false);
  phase(s, 3, "展開1　10-22分　｜　資料カード", false);
  title(s, "資料カード【光】", false, 30, 0.92, 0.55);
  sub(s, "産業革命が人類にもたらしたもの ①", false, 1.52);

  const cards = [
    ["光1　値段", "綿糸の値段が約1/12に。それまでぜいたく品だった綿の服を、ふつうの人が買えるようになった。"],
    ["光2　速さ", "1830年、リヴァプール〜マンチェスターに鉄道が開通。人と物が、一日で遠くまで運べるようになった。"],
    ["光3　力", "蒸気機関が、風・水・人や馬の筋力に代わる動力になった。天気や場所にしばられず、休みなく動かせる。"],
    ["光4　人の数", "イギリスの人口は1801年 約1,060万人 → 1901年 約3,700万人へ。より多くの人が生きられるようになった。"],
    ["光5　豊かさ", "世界の一人あたりGDPが、人類の歴史ではじめて、下がらずに上がり続けるようになった（次のスライド）。"],
    ["光6　仕事", "農村から都市へ人が移り、それまで存在しなかった職業と働き口が大量に生まれた。"]
  ];
  const cwc = (CW - 0.56) / 3, chc = 1.94;
  cards.forEach((c, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    infoCard(s, M + col * (cwc + 0.28), 2.06 + row * (chc + 0.22), cwc, chc,
      c[0], c[1], BRASS, false, PAPER);
  });
  source(s, "※数値は概数。人口はイギリス国勢調査による。印刷して切り離し、グループに配付してください。", false);
  notes(s, "【配付方法】このスライドとスライド13を印刷し、カードに切り分けて班に配ります。投影だけで進める場合は、この2枚を交互に映して読ませてください。\n\n【カードの設計意図】光は「安さ・速さ・力・人の数・豊かさ・仕事」の6つ。数値を伴うもの（光1・光4・光5）と、実感で分かるもの（光2・光3・光6）を混ぜてあります。\n\n光4（人口が3倍）は、光にも影にも読めるカードとして意図的に入れてあります。「人が増えた＝生きられる人が増えた」と読めば光ですが、「増えた人がスラムに押し込められた」と読めば影です。迷いが出なければ、教師からこのカードを指して揺さぶってください。");
}

// --- S12 一人あたりGDP ---
{
  const s = slide(false);
  phase(s, 3, "展開1　｜　資料【光5】", false);
  title(s, "人類は、はじめて「豊かになり続けた」", false, 30);
  sub(s, "世界の一人あたりGDPの推移（1990年国際ドル・概数）");

  s.addChart(pres.ChartType.line, [{
    name: "世界の一人あたりGDP",
    labels: ["1年", "1000年", "1500年", "1700年", "1820年", "1870年", "1913年", "1950年", "2000年"],
    values: [467, 450, 566, 616, 666, 870, 1524, 2111, 6039]
  }], {
    x: M, y: 2.20, w: CW - 4.30, h: 3.55,
    chartColors: [BRASS], lineSize: 4, lineSmooth: false,
    showLegend: false, showTitle: false,
    showValue: false,
    catAxisLabelColor: MUTED, valAxisLabelColor: MUTED,
    catAxisLabelFontFace: F, valAxisLabelFontFace: F,
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
    valGridLine: { color: LINE, size: 1 }, catGridLine: { style: "none" },
    valAxisMinVal: 0, lineDataSymbol: "circle", lineDataSymbolSize: 7
  });

  const px = W - M - 4.00;
  card(s, px, 2.20, 4.00, 3.55, false, PAPER);
  s.addText("読み取ろう", {
    x: px + 0.28, y: 2.40, w: 3.44, h: 0.34, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 15, bold: true, color: BRICK
  });
  s.addText("・1年から1700年までの1700年間で、豊かさはどれだけ増えた？\n・グラフが折れ曲がるのは、いつごろ？\n・そこで何が起きていた？", {
    x: px + 0.28, y: 2.80, w: 3.44, h: 1.60, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, color: INK, lineSpacing: 20, paraSpaceAfter: 7
  });
  card(s, px + 0.20, 4.54, 3.60, 0.96, false, INK);
  s.addText("人類の歴史の99%は、\n豊かさが横ばいだった。", {
    x: px + 0.40, y: 4.54, w: 3.20, h: 0.96, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 14, bold: true, color: WHITE, lineSpacing: 21
  });
  source(s, "出典：マディソン・プロジェクト・データベースによる推計（概数）。※横軸の年の間隔は均等ではありません。", false);
  notes(s, "【使い方】展開1の途中で全体に映します。班の作業を一度止めて、45秒だけ全員でこのグラフを見てください。\n\n【言い方】「手を止めて。これは世界の一人あたりの豊かさを、西暦1年から2000年まで並べたグラフ。……何か気づく？」\n\n「ずっと平ら」「最後だけ急に上がっている」が出れば十分です。「折れ曲がるのは1800年代。産業革命だね」と押さえます。\n\n【留意点】横軸は年が等間隔ではありません（1年→1000年と1950年→2000年が同じ幅）。厳密には立ち上がりを強調しすぎる表示です。上位の子から指摘が出たら、その指摘を褒めて認めてください。資料を疑う態度そのものが社会科の学力です。\n\n【留意点2】GDPは「豊かさ」の一面にすぎず、格差・環境・健康を含みません。これは主発問の討論（スライド24）で効いてきます。ここでは深追いしないでください。");
}

// --- S13 影のカード ---
{
  const s = slide(false);
  phase(s, 3, "展開1　10-22分　｜　資料カード", false);
  title(s, "資料カード【影】", false, 30, 0.92, 0.55);
  sub(s, "産業革命が人類にもたらしたもの ②", false, 1.52);

  const cards = [
    ["影1　子ども", "9歳の子どもが工場で働いた。1日12〜16時間。機械の下にもぐって糸くずを取る仕事は、子どもの体が小さいからと任された。"],
    ["影2　いのち", "1842年の報告書によると、マンチェスターの労働者の平均死亡年齢は17歳だった（次のスライド）。"],
    ["影3　まち", "工場に人が集まり、上下水道のないスラムができた。コレラが何度も流行した。"],
    ["影4　空気", "石炭を燃やした煤煙が街をおおった。ロンドンは「霧の都」と呼ばれた。"],
    ["影5　世界", "イギリスの安い綿布はインドの手織物業をこわした。その原料の綿花は、アメリカ南部の奴隷労働が支えていた。"],
    ["影6　地球", "人類が石炭を大量に燃やし始めた。今の気候変動は、ここから始まっている。"]
  ];
  const cwc = (CW - 0.56) / 3, chc = 1.94;
  cards.forEach((c, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    infoCard(s, M + col * (cwc + 0.28), 2.06 + row * (chc + 0.22), cwc, chc,
      c[0], c[1], BRICK, false, PAPER);
  });
  source(s, "※影1は工場法制定前後のイギリス議会報告による。影2はチャドウィック報告（1842年）。印刷して切り離し、グループに配付してください。", false);
  notes(s, "【カードの設計意図】影は「子ども・いのち・まち・空気・世界・地球」。手前（自分に近い）から遠く（地球規模）へ、また当時から現在へと並べてあります。\n\n影5と影6は、産業革命の影が“イギリス国内の話”で終わらないことを示すために必ず入れてください。豊かさの材料がどこから来たのかという視点は、第9時（欧米諸国のアジア進出）につながります。\n\n影6は、主発問の討論（スライド24）と、単元のまとめ（第10時）への布石です。\n\n【留意点】影1の労働時間や年齢は工場・時期により幅があります。「9歳」は1833年工場法が9歳未満の雇用を禁じたことに対応する数字として扱ってください。");
}

// --- S14 平均死亡年齢 ---
{
  const s = slide(false);
  phase(s, 3, "展開1　｜　資料【影2】", false);
  title(s, "同じ街に住んで、これだけ違った", false, 30);
  sub(s, "マンチェスターの階級別の平均死亡年齢（1842年・チャドウィック報告）");

  s.addChart(pres.ChartType.bar, [{
    name: "平均死亡年齢（歳）",
    labels: ["上流・専門職の家", "商人・農場主の家", "労働者・職人の家"],
    values: [38, 20, 17]
  }], {
    x: M, y: 2.20, w: CW - 4.30, h: 3.40,
    barDir: "col", chartColors: [BRICK], barGapWidthPct: 120,
    showLegend: false, showTitle: false,
    showValue: true, dataLabelPosition: "outEnd",
    dataLabelColor: INK, dataLabelFontFace: F, dataLabelFontSize: 16,
    dataLabelFormatCode: '0"歳"',
    catAxisLabelColor: MUTED, valAxisLabelColor: MUTED,
    catAxisLabelFontFace: F, valAxisLabelFontFace: F,
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
    valGridLine: { color: LINE, size: 1 }, catGridLine: { style: "none" },
    valAxisMinVal: 0, valAxisMaxVal: 45
  });

  const px = W - M - 4.00;
  card(s, px, 2.20, 4.00, 1.72, false, INK);
  s.addText("同じ街の、同じ時代。\nちがうのは、\n生まれた家だけ。", {
    x: px + 0.30, y: 2.20, w: 3.40, h: 1.72, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 17, bold: true, color: WHITE, lineSpacing: 27
  });
  card(s, px, 4.06, 4.00, 1.54, false, PAPER);
  s.addText([
    { text: "この数字の読み方\n", options: { bold: true, color: BRICK } },
    { text: "「平均死亡年齢」は平均寿命とは違います。当時は赤ちゃんのうちに亡くなる子が非常に多く、その分だけ平均が下がっています。17歳まで生きた人がそこで死んだ、という意味ではありません。", options: { color: INK } }
  ], {
    x: px + 0.26, y: 4.22, w: 3.48, h: 1.24, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, lineSpacing: 16
  });
  source(s, "出典：エドウィン・チャドウィック『大英帝国における労働者階級の衛生状態に関する報告』（1842年）。", false);
  notes(s, "【最重要の留意点】\n「平均死亡年齢17歳」を「平均寿命17歳」と説明しないでください。これは当時の乳幼児死亡率の高さを強く反映した数字です。右下の注記は必ず口頭でも触れてください。\n\n正しい押さえ方：「これは、生まれた子が何歳まで生きられたかの平均。労働者の家では、赤ちゃんのうちに亡くなる子がとても多かった。だから17歳まで下がる。上流の家では38歳。同じ街に住んでいるのに、生まれた家でこれだけ違った」\n\n数字の扱いを丁寧にすることが、そのまま「資料を正しく読む」という知識・技能の指導になります。むしろ授業の見せ場にしてください。\n\n【つなぎ】「この差を、産業革命は生んだ。……さっきのグラフでは、豊かさは上がっていたよね。この2枚は、どちらも本当のことなんだ」と言って、展開1のまとめ（スライド15）へ。");
}

// --- S15 全体共有 ---
{
  const s = slide(true);
  phase(s, 3, "展開1のまとめ　｜　全体で共有", true);
  title(s, "光と影は、切り離せるか", true, 30);
  sub(s, "各班の「迷い」に置かれたカードを聞く ― ここが本時の山場", true);

  const cw5 = (CW - 0.80) / 3;
  const cols = [
    ["光", BRASS, "値段が下がった\n速く運べる\n大きな力を使える\n人口が増えた\n豊かになり続けた\n仕事が生まれた", INK],
    ["迷い", "555961", "どちらにも置ける\nカードは？\n\nなぜ迷ったのか、\n班の言葉で説明する", WHITE],
    ["影", BRICK, "子どもが働いた\n17歳で亡くなった\nスラムができた\n空気が汚れた\nインド・奴隷制\n気候変動の始まり", WHITE]
  ];
  cols.forEach((c, i) => {
    const x = M + i * (cw5 + 0.40);
    card(s, x, 2.36, cw5, 2.82, true, c[1]);
    s.addText(c[0], {
      x, y: 2.52, w: cw5, h: 0.56, isTextBox: true, margin: 0,
      align: "center", valign: "middle", fontFace: F, fontSize: 26, bold: true, color: c[3]
    });
    s.addText(c[2], {
      x: x + 0.24, y: 3.14, w: cw5 - 0.48, h: 1.90, isTextBox: true, margin: 0,
      align: "center", fontFace: F, fontSize: 13.5, color: c[3], lineSpacing: 22
    });
  });
  card(s, M, 5.44, CW, 1.02, true, "2C2F36");
  s.addText([
    { text: "発問　", options: { bold: true, color: BRASS } },
    { text: "「光の側にあるカードを、1枚でも消したら、影も一緒に消える？」", options: { color: WHITE } }
  ], {
    x: M + 0.34, y: 5.44, w: CW - 0.68, h: 1.02, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 17, bold: true
  });
  notes(s, "【進め方】全体で3〜4分。すべての班に発表させる必要はありません。「迷い」に置いたカードを2〜3班に聞くだけで十分です。\n\n【聞き方】「迷いに置いたカード、どれ？　なんで迷ったの？」\n出やすいのは 光4（人口が増えた）と 光6（仕事が生まれた）です。「人は増えたけどスラムに住んでいる」「仕事はできたけど9歳の子どもの仕事」といった発言が出ます。これを板書に大きく残してください。\n\n【最後の発問の意図】\n「光を1枚消したら影も消える？」は、功と罪が同じ原因から出ていることに気づかせる問いです。「機械をなくせば子どもも働かなくていい。でも服も高くなる」という応答が出れば、展開2への準備が完全に整います。\n\n【時間管理】ここで延びやすい場面です。22分を超えたら切り上げて展開2へ進んでください。展開2と主発問のほうが本時の中心です。");
}

/* ============================================================
   PART 3　展開2　課題を子どもと立ち上げる
   ============================================================ */

// --- S16 展開2の発問 ---
{
  const s = slide(true);
  phase(s, 4, "展開2　22-30分　｜　全体（教師が板書する）", true);
  s.addText("この「影」を、\n当時の人たちは\nどう解決しようとした？", {
    x: M, y: 1.90, w: CW, h: 3.00, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 44, bold: true, color: WHITE, lineSpacing: 66
  });
  card(s, M, 5.20, CW, 1.10, true, INK2);
  s.addText([
    { text: "自分が当時の人だったら、どうする？　", options: { bold: true, color: BRASS } },
    { text: "思いついたことを、そのまま言ってよい。正解はまだ言いません。", options: { color: "E3E4E6" } }
  ], {
    x: M + 0.34, y: 5.20, w: CW - 0.68, h: 1.10, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 15
  });
  notes(s, "【この場面が本時でいちばん難しいところです】\n課題を子どもと立ち上げるには、子どもの発言を先に引き出し、あとから歴史的事実を重ねる順序が要ります。逆にすると教師の説明になります。\n\n【進め方】\n①この発問を出し、挙手で自由に出させる（3分）\n②出た案を、黒板の「影」の側に箇条書きで残す。評価せず、全部書く\n③5〜6個たまったら、次のスライド17を出す\n\n【必ず出る案（経験上ほぼ全部出ます）】\n・働く時間を法律で決める／制限する\n・子どもを働かせないようにする\n・給料を上げる\n・みんなで集まって文句を言う／ストライキ\n・選挙で政治家を変える\n・国がお金を出して助ける\n・工場をやめる／機械をこわす\n\n【出ないときの手立て】\n「じゃあ、今の日本では子どもが工場で働いてないよね。なんで働いてないんだと思う？」と現在から逆算させると必ず出ます。「法律」「学校があるから」が出ます。\n\n【言ってはいけないこと】\nここで「工場法というものがあってね」と先に言わないでください。子どもの発言を歴史と重ねる快感が、この授業の推進力です。");
}

// --- S17 対応表 ---
{
  const s = slide(false);
  phase(s, 4, "展開2　22-30分　｜　全体", false);
  title(s, "きみたちが今言ったこと、全部、実際に起きている", false, 29);
  sub(s, "子どもの発言　→　歴史上、本当に行われたこと");

  const pairs = [
    ["働く時間を法律で決める", "工場法（1833年）", "9歳未満の雇用を禁止。13歳未満は1日9時間まで。工場を見回る監督官も置かれた。"],
    ["子どもを働かせない", "教育の義務化へ", "働かせない代わりに学校へ。のちの義務教育につながる。"],
    ["みんなで集まって声を上げる", "労働組合", "1824年に団結を禁じる法律が撤廃され、労働者が団結できるようになった。"],
    ["選挙で政治を変える", "チャーティスト運動", "1830〜40年代。労働者に選挙権を求める大運動が起きた。"],
    ["国が助ける", "公衆衛生法（1848年）", "上下水道の整備など、街の衛生を国と自治体の仕事にした。"],
    ["仕組みそのものを変える", "社会主義", "マルクスら。『共産党宣言』1848年。工場を個人のものにしない社会を構想した。"]
  ];
  const rowH2 = 0.66, gap2 = 0.085;
  pairs.forEach((p, i) => {
    const y = 2.16 + i * (rowH2 + gap2);
    card(s, M, y, CW, rowH2, false, PAPER);
    s.addText("「" + p[0] + "」", {
      x: M + 0.26, y, w: 3.30, h: rowH2, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 13, bold: true, color: MUTED
    });
    s.addShape(pres.ShapeType.rightArrow, {
      x: M + 3.62, y: y + 0.245, w: 0.30, h: 0.17,
      fill: { color: BRASS }, line: { type: "none" }
    });
    s.addText(p[1], {
      x: M + 4.06, y, w: 2.50, h: rowH2, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 13.5, bold: true, color: BRICK
    });
    s.addText(p[2], {
      x: M + 6.62, y, w: CW - 6.88, h: rowH2, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 12, color: INK
    });
  });
  notes(s, "【このスライドが、課題を立ち上げるための切り札です】\n\n【出し方】子どもの案が黒板に5〜6個たまったところで出します。左の列を指しながら「これ、○○さんが言ったやつだよね」と、実際の発言者の名前を呼んでください。自分の考えが歴史と一致していた、という経験が課題への当事者意識を生みます。\n\n【言い方】「実はね。きみたちが今言ったこと、全部、本当に起きてるんだ。しかも、ほとんど50年くらいの間に」\n\n【運用のコツ】\n・子どもから出なかった行だけを教師が補います。全部出ることは稀なので、2〜3行は教師が足す前提で構いません\n・時間がなければ上から3行だけで十分機能します\n・年号は覚えさせません。「同じ時期に集中して起きた」ことだけ押さえます\n\n【次への発問】\n「じゃあ、これで問題は解決したのかな？」と問い、スライド18で課題を設定します。");
}

// --- S18 課題の設定 ---
{
  const s = slide(true);
  phase(s, 4, "課題の設定　｜　次の時間からの学習課題", true);
  s.addText("学習課題", {
    x: M, y: 1.60, w: CW, h: 0.42, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, bold: true, color: BRASS, charSpacing: 3
  });
  s.addText("産業革命が生んだ問題を、\n人類はどう解決しようとしたのか。\nそして、解決できたのか。", {
    x: M, y: 2.20, w: CW, h: 2.75, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 36, bold: true, color: WHITE, lineSpacing: 58
  });
  card(s, M, 5.28, CW, 1.10, true, INK2);
  s.addText([
    { text: "この課題は次の時間（第8時）で本格的に追究します。", options: { color: "E3E4E6" } },
    { text: "　今日は、ここから見えてくる「あるもの」まで進みます。", options: { bold: true, color: BRASS } }
  ], {
    x: M + 0.34, y: 5.28, w: CW - 0.68, h: 1.10, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 15
  });
  notes(s, "【課題の立ち上がり方】\nこの課題は教師が用意したものではなく、子どもの発言（スライド16）を歴史的事実（スライド17）と重ねた結果として立ち上がっています。この順序を守れば、課題設定は自然に成立します。\n\n【板書】この課題を、本時の問いの下に書き加えます。第8時の冒頭でこの板書を再現して始めると、単元がつながります。\n\n【2時間扱いにする場合】\nここで1時間目を終えます。振り返りは「今日出た解決策の中で、一番効きそうなのはどれか。理由も」と書かせます。2時間目はスライド19から始めてください。\n\n【1時間で通す場合】\n「じゃあ、この6つをよく見てほしい。実は全部に共通していることがある」と言ってスライド19へ。ここまでで30分が目安です。");
}

/* ============================================================
   PART 4　展開3　資本主義への到達 → 主発問
   ============================================================ */

// --- S19 共通点 ---
{
  const s = slide(false);
  phase(s, 5, "展開3　30-35分　｜　全体", false);
  title(s, "6つの解決策に、共通していること", false, 30);
  sub(s, "どれも「やらなかったこと」がある");

  const cw6 = (CW - 0.40) / 2;
  infoCard(s, M, 2.16, cw6, 2.42,
    "やったこと",
    "・働く時間にルールをつくった\n・子どもを工場から学校へ移した\n・労働者が団結できるようにした\n・選挙権を広げようとした\n・街の衛生を国の仕事にした",
    BRASS, false, PAPER);
  infoCard(s, M + cw6 + 0.40, 2.16, cw6, 2.42,
    "やらなかったこと",
    "・工場をなくさなかった\n・機械を捨てなかった\n・「もうけを求めて競争する」\n　という仕組みは、やめなかった",
    BRICK, false, PAPER);

  card(s, M, 4.86, CW, 1.28, false, INK);
  s.addText("つまり ―　ルールを足しただけで、仕組みそのものは残した。", {
    x: M + 0.34, y: 4.86, w: CW - 0.68, h: 1.28, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 24, bold: true, color: WHITE
  });
  notes(s, "【ねらい】個別の解決策から、その背後にある社会の仕組みへ視点を引き上げる。\n\n【言い方】「よく見て。6つとも、あることをやってない。何をやってないと思う？」\n\n子どもからは「工場をなくしてない」が出やすいです。出たら大きく褒めて、「そう。誰も工場をやめようとは言わなかった。機械も捨てなかった。じゃあ、その残した仕組みって、なんて名前だと思う？」とつなぎます。\n\n【出ないときの手立て】\n「もし工場が全部なくなったら、きみの服はどうなる？」と聞くと、「高くなる」「なくなる」が出ます。「そう。だから誰もやめられなかったんだ」\n\n【時間】2分。ここは速く通してください。");
}

// --- S20 資本主義の定義 ---
{
  const s = slide(true);
  phase(s, 5, "展開3　30-35分　｜　全体", true);
  title(s, "その残した仕組みを、こう呼ぶ", true, 30);
  s.addText("資本主義", {
    x: M, y: 1.66, w: CW, h: 0.90, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 46, bold: true, color: BRASS
  });

  const cw7 = (CW - 0.80) / 3;
  const defs = [
    ["① 私有", "工場や機械を、国ではなく個人や会社が持つ"],
    ["② 利潤", "もうけを増やすことを目的に、生産する"],
    ["③ 競争", "市場での自由な競争が、社会を動かす"]
  ];
  defs.forEach((d, i) => {
    const x = M + i * (cw7 + 0.40);
    card(s, x, 2.76, cw7, 1.86, true, INK2);
    s.addText(d[0], {
      x: x + 0.28, y: 2.94, w: cw7 - 0.56, h: 0.46, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 22, bold: true, color: BRASS
    });
    s.addText(d[1], {
      x: x + 0.28, y: 3.46, w: cw7 - 0.56, h: 0.98, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 14, color: "E3E4E6", lineSpacing: 22
    });
  });
  card(s, M, 4.90, CW, 1.22, true, "2C2F36");
  s.addText("産業革命は、この仕組みを世界中に広げた。私たちは今も、この中で暮らしている。", {
    x: M + 0.34, y: 4.90, w: CW - 0.68, h: 1.22, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 19, bold: true, color: WHITE
  });
  notes(s, "【押さえ方】定義を書き写させます（1分）。3つのキーワード（私有・利潤・競争）だけで十分です。\n\n【中学2年生への配慮】\n「資本主義」は公民的分野で本格的に扱う語です。ここでは厳密な経済学的定義ではなく、「工場を個人が持ち、もうけを求めて競争する仕組み」という素朴な理解で構いません。第3学年の公民でもう一度出会う語だと予告しておくと、系統性の意識が育ちます。\n\n【言い方】「今日はじめて出てきた言葉。3年生の公民でもう一回ちゃんとやるから、今日は3つだけ覚えて。私有・利潤・競争」");
}

// --- S21 社会主義 ---
{
  const s = slide(false);
  phase(s, 5, "展開3　30-35分　｜　全体", false);
  title(s, "ただし、仕組みごと変えようとした人たちがいた", false, 29);
  sub(s, "スライド17の6番目 ―「社会主義」のその後");

  const cw8 = (CW - 0.80) / 3;
  const flow = [
    ["1848年", "マルクスらが『共産党宣言』を発表。工場を個人のものにしない社会を構想した。", BRICK],
    ["20世紀", "ソ連や中国で、実際にその仕組みが国家の規模で試された。", MUTED],
    ["1991年", "ソ連が解体。多くの国が市場経済へ移った。", INK]
  ];
  flow.forEach((f, i) => {
    const x = M + i * (cw8 + 0.40);
    card(s, x, 2.20, cw8, 2.10, false, PAPER);
    s.addText(f[0], {
      x: x + 0.28, y: 2.40, w: cw8 - 0.56, h: 0.48, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 22, bold: true, color: f[2]
    });
    s.addText(f[1], {
      x: x + 0.28, y: 2.94, w: cw8 - 0.56, h: 1.18, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 21
    });
    if (i < 2) {
      s.addShape(pres.ShapeType.rightArrow, {
        x: x + cw8 + 0.06, y: 3.10, w: 0.28, h: 0.30,
        fill: { color: "C9CCD2" }, line: { type: "none" }
      });
    }
  });
  card(s, M, 4.58, CW, 1.54, false, INK);
  s.addText([
    { text: "では ―　資本主義が「勝った」ということだろうか。\n", options: { bold: true, color: WHITE, fontSize: 24 } },
    { text: "残った、ということと、正しかった、ということは、同じだろうか。", options: { color: BRASS, fontSize: 16 } }
  ], {
    x: M + 0.34, y: 4.58, w: CW - 0.68, h: 1.54, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, lineSpacing: 32
  });
  notes(s, "【ねらい】主発問への助走。「資本主義が残った」という事実と、「それでよかったのか」という価値判断を切り離す。\n\n【留意点】社会主義国の歴史的評価には、多様な立場があります。ここでは「試みがあり、多くは市場経済へ移行した」という事実の確認にとどめ、体制の善悪を教師が断定しないでください。断定すると、次の主発問で子どもが教師の顔色を読み始めます。本時の主発問は、子どもが自分の立場を決める時間です。\n\n【言い方】「残ったから正しい、って言えるかな？　……じゃあ、今日の一番の問いにいくよ」と言ってスライド22へ。\n\n【時間】2分。ここまでで35分。");
}

// --- S22 主発問 ---
{
  const s = slide(true);
  phase(s, 6, "主発問　35-45分　｜　個人 → 全体", true);
  s.addText("主発問", {
    x: M, y: 1.18, w: CW, h: 0.40, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, bold: true, color: BRASS, charSpacing: 3
  });
  s.addText("結局、資本主義は\n人類を幸せにしたのだろうか。", {
    x: M, y: 1.72, w: CW, h: 1.90, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 42, bold: true, color: WHITE, lineSpacing: 62
  });

  // 立場表明の数直線
  const lx = M + 0.60, lw = CW - 1.20, ly = 4.42;
  s.addShape(pres.ShapeType.line, {
    x: lx, y: ly, w: lw, h: 0,
    line: { color: "5A5E67", width: 2.5 }
  });
  for (let i = 0; i <= 10; i++) {
    const cxp = lx + (lw / 10) * i;
    const big = (i % 5 === 0);
    s.addShape(pres.ShapeType.ellipse, {
      x: cxp - (big ? 0.11 : 0.06), y: ly - (big ? 0.11 : 0.06),
      w: big ? 0.22 : 0.12, h: big ? 0.22 : 0.12,
      fill: { color: big ? BRASS : "5A5E67" }, line: { type: "none" }
    });
    if (big) {
      s.addText(String(i), {
        x: cxp - 0.30, y: ly + 0.20, w: 0.60, h: 0.30, isTextBox: true, margin: 0,
        align: "center", fontFace: F, fontSize: 13, bold: true, color: BRASS
      });
    }
  }
  s.addText("幸せにしていない", {
    x: lx - 0.55, y: ly - 0.62, w: 2.40, h: 0.34, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 14, bold: true, color: BRICK
  });
  s.addText("幸せにした", {
    x: lx + lw - 1.85, y: ly - 0.62, w: 2.40, h: 0.34, isTextBox: true, margin: 0,
    align: "right", fontFace: F, fontSize: 14, bold: true, color: BRASS
  });
  s.addText("まず一人で　①自分の数字を決める　②理由を一言、ワークシートに書く　（2分）", {
    x: M, y: 5.42, w: CW, h: 0.44, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 15, color: "E3E4E6"
  });
  notes(s, "【進め方】\n①個人で数字と理由を書く（2分）。ここを飛ばさないでください。先に全体討論を始めると、声の大きい子の意見に引きずられます\n②挙手または板書の数直線にネームプレートを貼って、学級の分布を可視化する（1分）\n③分布の両端と中央から1人ずつ指名して理由を聞く（3分）\n④次のスライド23・24の「ゆさぶり資料」を出して、もう一度考えさせる（4分）\n\n【指名の順序】\n中間の子から先に当てると、どっちつかずの発言で場が締まりません。端の子（0〜2、8〜10）から当てて対立軸を明確にし、そのあと中間の子に「両方わかるという人は？」と振ると、多面的な見方が言語化されます。\n\n【教師の立場】\n教師は自分の答えを言いません。最後まで言わないでください。まとめ（スライド26・27）で示すのは「功罪の両面がある」という見方であって、幸せにしたかどうかの答えではありません。\n\n【ICT】1人1台端末があれば、フォームで数字を集計して分布をその場でグラフ表示すると効果的です。");
}

/* ============================================================
   PART 5　討論のゆさぶり資料
   ============================================================ */

// --- S23 ゆさぶり① ---
{
  const s = slide(false);
  phase(s, 6, "討論のゆさぶり ①　｜「幸せにした」側の資料", false);
  title(s, "この200年で、世界はこう変わった", false, 30);

  s.addChart(pres.ChartType.bar, [{
    name: "極度の貧困のもとで暮らす人の割合（%）",
    labels: ["1820年", "1910年", "1950年", "1990年", "2019年"],
    values: [76, 66, 55, 38, 9]
  }], {
    x: M, y: 2.14, w: CW - 4.60, h: 3.42,
    barDir: "col", chartColors: [BRASS], barGapWidthPct: 90,
    showLegend: false, showTitle: false,
    showValue: true, dataLabelPosition: "outEnd",
    dataLabelColor: INK, dataLabelFontFace: F, dataLabelFontSize: 14,
    dataLabelFormatCode: '0"%"',
    catAxisLabelColor: MUTED, valAxisLabelColor: MUTED,
    catAxisLabelFontFace: F, valAxisLabelFontFace: F,
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
    valGridLine: { color: LINE, size: 1 }, catGridLine: { style: "none" },
    valAxisMinVal: 0, valAxisMaxVal: 90
  });

  const px = W - M - 4.30;
  s.addText("極度の貧困のもとで\n暮らす人の割合", {
    x: M, y: 1.66, w: CW - 4.60, h: 0.44, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, color: MUTED
  });
  card(s, px, 2.14, 4.30, 1.66, false, PAPER);
  s.addText("30歳 → 73歳", {
    x: px, y: 2.28, w: 4.30, h: 0.74, isTextBox: true, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 30, bold: true, color: INK
  });
  s.addText("世界の平均寿命\n1800年ごろ 約30歳 → 2019年 約73歳", {
    x: px + 0.20, y: 3.04, w: 3.90, h: 0.62, isTextBox: true, margin: 0,
    align: "center", fontFace: F, fontSize: 12, color: MUTED, lineSpacing: 17
  });
  card(s, px, 3.98, 4.30, 1.72, false, INK);
  s.addText([
    { text: "問い直し\n", options: { bold: true, color: BRASS } },
    { text: "産業革命がなかったら、\nこの200年で生まれた人の多くは、\n大人になれなかったかもしれない。", options: { color: WHITE } }
  ], {
    x: px + 0.30, y: 3.98, w: 3.70, h: 1.72, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 14, lineSpacing: 22
  });
  source(s, "出典：世界銀行／Our World in Data による推計（概数）。貧困率は2011年購買力平価で1日1.90ドル未満を基準とする。", false);
  notes(s, "【使い方】主発問の討論の途中、「幸せにしていない」側に意見が偏ったときに出します。偏っていなければ、スライド24と両方まとめて出して構いません。\n\n【言い方】「ちょっと待って。この資料を見てから、もう一度考えて」\n\n【押さえどころ】\n貧困率76%→9%は、産業革命に始まる経済成長がもたらした最大の成果です。「豊かになったのは一部の金持ちだけ」という素朴な理解を揺さぶります。\n\n【留意点】このグラフは「割合」です。世界人口が大きく増えているため、極度の貧困のもとで暮らす人の実数の減り方は割合ほど劇的ではありません。鋭い子が指摘したら褒めてください。\n\n【留意点2】貧困率の低下すべてが資本主義の成果だと断定しないでください。医療・公衆衛生・国際協力など複数の要因があります。「産業革命が始めた経済成長が、大きな要因の一つ」という言い方が正確です。");
}

// --- S24 ゆさぶり② ---
{
  const s = slide(false);
  phase(s, 6, "討論のゆさぶり ②　｜「幸せにしていない」側の資料", false);
  title(s, "同じ200年で、こうもなった", false, 30);

  const cw9 = (CW - 0.56) / 3;
  const facts = [
    ["格差", "世界の富の\n約半分", "上位1%の人々が保有している"],
    ["地球", "約 1.5℃", "産業革命前とくらべた世界の平均気温の上昇"],
    ["子ども", "約1億6千万人", "今も世界で児童労働に従事している子どもの数"]
  ];
  facts.forEach((f, i) => {
    const x = M + i * (cw9 + 0.28);
    card(s, x, 2.14, cw9, 2.30, false, PAPER);
    s.addText(f[0], {
      x: x + 0.26, y: 2.32, w: cw9 - 0.52, h: 0.36, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 14, bold: true, color: BRICK
    });
    s.addText(f[1], {
      x: x + 0.26, y: 2.72, w: cw9 - 0.52, h: 0.90, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 28, bold: true, color: INK, lineSpacing: 34
    });
    s.addText(f[2], {
      x: x + 0.26, y: 3.66, w: cw9 - 0.52, h: 0.62, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12, color: MUTED, lineSpacing: 17
    });
  });

  card(s, M, 4.62, CW, 1.44, false, INK);
  s.addText([
    { text: "そして　", options: { bold: true, color: BRICK, fontSize: 20 } },
    { text: "きみが今着ているその服も、遠い国の、とても安い賃金の労働でつくられているかもしれない。", options: { color: WHITE, fontSize: 20 } }
  ], {
    x: M + 0.34, y: 4.62, w: CW - 0.68, h: 1.44, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, bold: true, lineSpacing: 30
  });
  source(s, "出典：格差＝クレディ・スイス／UBS グローバル・ウェルス・レポート、気温＝IPCC 第6次評価報告書、児童労働＝ILO・ユニセフ（2020年）。いずれも概数。", false);
  notes(s, "【使い方】スライド23の直後に出します。「幸せにした」側に意見が傾いたときの揺さぶりでもあります。\n\n【最重要】最下段の一文が本時の設計の要です。導入（スライド5・6）で扱った「自分の服」がここで戻ってきます。トンネルの入口と出口がつながる瞬間です。ここは必ず間をとって、静かに言ってください。\n\n【言い方】「今日の最初に、着ている服の値段を聞いたよね。……安かったよね。なんで安いんだろう」\n\n【留意点】特定の国・企業・ブランドを名指ししないでください。また「だから安い服を買うのは悪い」という結論に誘導しないこと。消費者個人の責任に落とすと、社会の仕組みを考える授業にならなくなります。「仕組みの中に自分もいる」という気づきに留めます。\n\n【留意点2】児童労働の1億6千万人はILO・ユニセフの2020年推計です。産業革命の直接の帰結ではなく、現代のグローバル経済の問題として扱ってください。");
}

// --- S25 ゆさぶり③ ---
{
  const s = slide(true);
  phase(s, 6, "討論のゆさぶり ③　｜　ペア 30秒", true);
  s.addText("じゃあ、やめられる？", {
    x: M, y: 1.70, w: CW, h: 1.10, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 48, bold: true, color: WHITE
  });
  s.addText("今日ここに来るまでに、きみが使ったもの", {
    x: M, y: 2.96, w: CW, h: 0.40, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, color: MUTED_D
  });

  const words = ["電気", "スマホ", "電車・バス", "服", "給食", "水道", "教科書", "薬"];
  const bw3 = (CW - 0.28 * 3) / 4, bh3 = 0.72;
  words.forEach((w, i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = M + col * (bw3 + 0.28), y = 3.50 + row * (bh3 + 0.22);
    card(s, x, y, bw3, bh3, true, INK2);
    s.addText(w, {
      x, y, w: bw3, h: bh3, isTextBox: true, margin: 0,
      align: "center", valign: "middle", fontFace: F, fontSize: 19, bold: true, color: BRASS
    });
  });
  card(s, M, 5.72, CW, 0.86, true, "2C2F36");
  s.addText("この中に、産業革命なしで用意できるものは、いくつある？", {
    x: M + 0.34, y: 5.72, w: CW - 0.68, h: 0.86, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 18, bold: true, color: WHITE
  });
  notes(s, "【ねらい】結論への最後の一手。功罪の議論を「では、なくせるのか」という現実の地平に引き戻す。\n\n【進め方】ペアで30秒。「一つでも用意できるものある？」と聞くと、「水道は？」「薬は？」と考え始めます。どれも大量生産・動力・化学工業なしには成立しません。\n\n【言い方】「一つでもいい。産業革命なしで用意できるもの、ある？」\n……ほぼ出ません。出なかったという事実そのものが、まとめの説得力になります。\n\n【出た場合の受け方】「教科書は昔もあった」等の発言が出たら、「たしかに本はあった。でも、全員が1冊ずつ持てるようになったのはいつからだろう」と返します。「量」の問題に焦点化すると、産業革命の意味が際立ちます。\n\n【時間】1分半。ここまでで45分。");
}

/* ============================================================
   PART 6　まとめ
   ============================================================ */

// --- S26 まとめ 功罪 ---
{
  const s = slide(false);
  phase(s, 7, "まとめ　45-50分　｜　個人（振り返り）", false);
  title(s, "産業革命が、私たち人類にもたらしたもの", false, 30);

  const cwA = (CW - 0.40) / 2;
  card(s, M, 2.14, cwA, 2.72, false, PAPER);
  s.addText("功", {
    x: M + 0.30, y: 2.34, w: 1.00, h: 0.56, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 32, bold: true, color: BRASS
  });
  s.addText("人類ははじめて、貧しさから\n抜け出す力を手に入れた。", {
    x: M + 0.30, y: 2.96, w: cwA - 0.60, h: 0.80, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 17, bold: true, color: INK, lineSpacing: 26
  });
  s.addText("値段が下がり、量が増え、遠くまで運べるようになった。\nより多くの人が生まれ、生き延び、長く生きられるようになった。", {
    x: M + 0.30, y: 3.84, w: cwA - 0.60, h: 0.86, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, color: MUTED, lineSpacing: 20
  });

  card(s, M + cwA + 0.40, 2.14, cwA, 2.72, false, PAPER);
  s.addText("罪", {
    x: M + cwA + 0.70, y: 2.34, w: 1.00, h: 0.56, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 32, bold: true, color: BRICK
  });
  s.addText("その代金を払ったのは、\n別の誰かだった。", {
    x: M + cwA + 0.70, y: 2.96, w: cwA - 0.60, h: 0.80, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 17, bold: true, color: INK, lineSpacing: 26
  });
  s.addText("工場で働いた子ども。17歳で亡くなった労働者。\n手仕事を奪われたインドの職人。綿花を摘んだ奴隷。\nそして、これから生きる人たちの地球。", {
    x: M + cwA + 0.70, y: 3.84, w: cwA - 0.60, h: 0.86, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, color: MUTED, lineSpacing: 20
  });

  card(s, M, 5.14, CW, 0.98, false, INK);
  s.addText("この二つは、別々に起きたことではない。同じ一つのことの、表と裏だった。", {
    x: M + 0.34, y: 5.14, w: CW - 0.68, h: 0.98, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 19, bold: true, color: WHITE
  });
  notes(s, "【まとめの原則】\n教師が「幸せにした／していない」の答えを言わないこと。示すのは「功と罪の両面がある」「その二つは切り離せない」という見方です。\n\n【最下段の一文が、本時の中心的な学習内容です】\n安い服（功）と、子どもの労働（罪）は、同じ「大量生産」という一つのことから出ています。だから片方だけを取ることができない。ここが多面的・多角的な考察の到達点です。\n\n【板書】黒板の左に功、右に罪を整理した状態で、中央下にこの一文を書きます。");
}

// --- S27 結論 ---
{
  const s = slide(true);
  phase(s, 7, "まとめ　45-50分　｜　本時の結論", true);
  s.addText("産業革命は人類に、\n「豊かさ」と「請求書」を\n同時に渡した。", {
    x: M, y: 1.16, w: CW, h: 2.44, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 34, bold: true, color: WHITE, lineSpacing: 54
  });
  card(s, M, 3.76, CW, 1.20, true, INK2);
  s.addText("それでも私たちは、もうこれなしでは生きられない。", {
    x: M + 0.34, y: 3.76, w: CW - 0.68, h: 1.20, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 26, bold: true, color: BRASS
  });
  s.addText([
    { text: "だから、問いはこうなる。\n", options: { color: MUTED_D, fontSize: 15 } },
    { text: "「やめるか、続けるか」ではなく、「どう作り直すか」。", options: { color: WHITE, fontSize: 27, bold: true } }
  ], {
    x: M, y: 5.12, w: CW, h: 1.26, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, lineSpacing: 40
  });
  notes(s, "【本時の結論】これが、先生が子どもに持ち帰らせたい結論です。\n\n【言い方】ゆっくり、間をとって。\n「産業革命は、人類に豊かさをくれた。同時に、請求書も渡した。……その請求書は、まだ払い終わっていない。格差も、地球の温度も、今のきみたちの問題として残っている」\n「でもね。今日わかったとおり、私たちはもうこれなしでは生きられない。服も、電気も、薬も」\n「だから、考えるべきことは『やめるか続けるか』じゃない。『どう作り直すか』なんだ」\n\n【この結論の意義】\n功罪の両論併記で終わらせず、かといって「産業革命はよかった／悪かった」と断定もせず、子どもを当事者の位置に置いたまま次の学習へ渡します。単元を貫く問い「産業革命は、人類を幸せにしたのか」への最終的な答えは、単元のまとめ（第10時）で子ども自身が出します。本時はその材料を全部そろえた、という位置づけです。\n\n【時間】2分。ここは絶対に急がないでください。時間が足りなければ、スライド25のペア活動を削ってでもここを確保します。");
}

// --- S28 振り返りと次時 ---
{
  const s = slide(false);
  phase(s, 7, "振り返り　｜　個人（ワークシート）", false);
  title(s, "振り返りを書こう", false, 30);
  sub(s, "3分　／　どれか一つでよい");

  const cwB = (CW - 0.56) / 3;
  const qs = [
    ["視点1　自分の変化", "授業の前と後で、自分の立場（数直線の数字）は動いた？　動いたなら、動かした資料はどれ？"],
    ["視点2　切り離せるか", "産業革命の「光」と「影」は、切り離せると思う？　思う・思わない、どちらでもよい。理由を書こう。"],
    ["視点3　これから", "「どう作り直すか」。今のきみに、何か一つ思いつくことはある？"]
  ];
  qs.forEach((q, i) => {
    infoCard(s, M + i * (cwB + 0.28), 2.20, cwB, 2.12, q[0], q[1], BRICK, false, PAPER);
  });

  card(s, M, 4.64, CW, 1.48, false, INK);
  s.addText([
    { text: "次の時間　第8時　", options: { color: BRASS, bold: true, fontSize: 14 } },
    { text: "働く人を守ったのは、誰か　― 労働問題と社会主義\n", options: { color: WHITE, bold: true, fontSize: 21 } },
    { text: "今日みんなで立てた課題「産業革命が生んだ問題を、人類はどう解決しようとしたのか」を、腰を据えて追究します。", options: { color: MUTED_D, fontSize: 13 } }
  ], {
    x: M + 0.34, y: 4.64, w: CW - 0.68, h: 1.48, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, lineSpacing: 28
  });
  notes(s, "【振り返りの設計】\n視点1が「主体的に学習に取り組む態度」の主たる評価材料になります。自分の考えの変容を、資料を根拠に説明できているかを見取ってください。「変わらなかった」も、理由が書けていれば十分に評価できます。\n\n視点2が「思考・判断・表現」の評価材料です。功と罪を関連付けて考察できているかを見ます。\n\n視点3は書ける子だけで構いません。単元のまとめ（第10時）への布石です。\n\n【時間が足りないとき】\n視点1だけを書かせて回収します。1分で書けます。");
}

/* ============================================================
   PART 7　教師用付録
   ============================================================ */

// --- S29 板書計画 ---
{
  const s = slide(false);
  phase(s, 0, "教師用｜板書計画", false);
  title(s, "板書計画", false, 30);
  sub(s, "本時の問いを中央上部に置き、左右に功罪を広げる。課題と結論は中央下に残す");

  // 黒板
  const bx = M, by = 2.16, bw4 = CW, bh4 = 3.42;
  s.addShape(pres.ShapeType.rect, {
    x: bx, y: by, w: bw4, h: bh4,
    fill: { color: "1E3A2E" }, line: { color: "16281F", width: 2 }
  });
  s.addText("本時の問い　産業革命は私たち人類に何をもたらしたのか", {
    x: bx + 0.20, y: by + 0.16, w: bw4 - 0.40, h: 0.42, isTextBox: true, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 15, bold: true, color: WHITE
  });
  s.addText("【光】\n・値段が下がった\n・速く運べる\n・大きな力\n・人口が増えた\n・豊かになり続けた\n・仕事が生まれた", {
    x: bx + 0.28, y: by + 0.68, w: 3.20, h: 1.86, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12, color: "F2E6B8", lineSpacing: 17
  });
  s.addText("【影】\n・子どもが働いた\n・17歳で亡くなった\n・スラム／コレラ\n・空気の汚れ\n・インド／奴隷制\n・気候変動の始まり", {
    x: bx + bw4 - 3.48, y: by + 0.68, w: 3.20, h: 1.86, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12, color: "F6C7BC", lineSpacing: 17
  });
  s.addText("【子どもが出した解決策】\n工場法／教育／労働組合／\nチャーティスト／公衆衛生法／\n社会主義\n　↓　共通点は？\n《資本主義》私有・利潤・競争", {
    x: bx + 3.70, y: by + 0.68, w: 5.40, h: 1.86, isTextBox: true, margin: 0,
    align: "center", fontFace: F, fontSize: 12, color: WHITE, lineSpacing: 17
  });
  s.addText("【主発問】結局、資本主義は人類を幸せにしたのだろうか　→　0 ─── 5 ─── 10（ネームプレート）", {
    x: bx + 0.20, y: by + 2.62, w: bw4 - 0.40, h: 0.34, isTextBox: true, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 12.5, bold: true, color: "F2E6B8"
  });
  s.addText("光と影は、同じ一つのことの表と裏　／　問いは「やめるか」ではなく「どう作り直すか」", {
    x: bx + 0.20, y: by + 2.98, w: bw4 - 0.40, h: 0.32, isTextBox: true, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 12.5, bold: true, color: WHITE
  });

  s.addText("※ 中央の「子どもが出した解決策」は、展開2で子どもの発言を書き足していく場所。あらかじめ空けておく。", {
    x: M, y: 5.76, w: CW, h: 0.40, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 12, color: MUTED
  });
  notes(s, "【板書の要点】\n中央の空白がこの板書の生命線です。展開2で子どもが出した解決策を、その場で書き足していきます。授業開始時点で中央を空けておき、「ここは、あとできみたちの言葉で埋める」と予告しておくと、子どもが発言に向かいます。\n\n【ネームプレート】主発問の数直線は、黒板下部に横一線を引き、ネームプレートを貼らせます。学級の分布が一目で見え、討論後に動かしてよいことにすると、考えの変容がそのまま可視化されます。振り返り（視点1）の材料にもなります。\n\n【消さないもの】本時の問い、学習課題、最下段の結論。この3つは授業の最後まで残してください。写真に撮って第8時の冒頭で提示すると、単元がつながります。");
}

// --- S30 評価規準 ---
{
  const s = slide(false);
  phase(s, 0, "教師用｜評価", false);
  title(s, "本時の評価規準と見取りの場面", false, 30);
  sub(s, "年間指導計画・評価計画の様式に合わせた文言");

  const rows = [
    ["知識・技能", INK,
      "○産業革命によって工場制機械工業が広がり、資本主義社会が成立したことについて理解している。",
      "単元テスト／ワークシート"],
    ["思考・判断・表現", BRICK,
      "○産業革命がもたらした変化について、社会全体の豊かさと人々への負担を関連付けながら、多面的・多角的に考察している。",
      "単元シート／机間指導"],
    ["主体的に学習に取り組む態度", BRASS,
      "○産業革命と資本主義について、よりよい社会の実現を視野に、そこで見られる課題を主体的に追究しようとしている。",
      "振り返り（視点1）／数直線の変容"]
  ];
  rows.forEach((r, i) => {
    const y = 2.20 + i * 1.24;
    card(s, M, y, CW, 1.10, false, PAPER);
    s.addText(r[0], {
      x: M + 0.28, y: y + 0.14, w: 3.20, h: 0.36, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 14, bold: true, color: r[1]
    });
    s.addText("見取り：" + r[3], {
      x: M + 0.28, y: y + 0.56, w: 3.20, h: 0.36, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 11, color: MUTED
    });
    s.addText(r[2], {
      x: M + 3.70, y, w: CW - 4.00, h: 1.10, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 13.5, color: INK, lineSpacing: 21
    });
  });

  card(s, M, 5.98, CW, 0.86, false, INK);
  s.addText("本時の重点は「思考・判断・表現」。分類の正誤ではなく、功と罪を関連付けて説明できているかを見取る。", {
    x: M + 0.34, y: 5.98, w: CW - 0.68, h: 0.86, isTextBox: true, margin: 0,
    valign: "middle", fontFace: F, fontSize: 14, bold: true, color: WHITE
  });
  notes(s, "文言は年間指導計画・評価計画（単元・題材３の詳細シート）の書きぶりに合わせてあります。そのまま様式に転記できます。\n\n【本時の重点】\n1単位時間で3観点すべてを評価しようとしないでください。本時は「思考・判断・表現」が重点です。知識・技能は単元テストに、態度は振り返りの蓄積に委ねます。\n\n【机間指導で見取るポイント（思考・判断・表現）】\nB：光と影の両方を挙げて説明できている\nA：光と影が同じ原因から生じていることに触れて説明できている（例「安い服がほしいから工場ができて、工場ができたから子どもが働いた」）\nC：一方だけを挙げている → 反対側のカードを指して問い返す\n\n【留意点】主発問への答え（数直線の数字）そのものは評価対象にしません。どちらの立場でも、資料を根拠に説明できていればよいと、子どもにも明示してください。");
}

// --- S31 プランB ---
{
  const s = slide(false);
  phase(s, 0, "教師用｜想定外への備え", false);
  title(s, "うまくいかないときの手立て", false, 30);
  sub(s, "授業前に「何を削るか」を決めておく", false, 1.62);

  const cwC = (CW - 0.40) / 2;
  const bodyOpt = { fontSize: 12.5, lineSpacing: 19 };
  function planCard(x, y, w, h, head, lines) {
    card(s, x, y, w, h, false, PAPER);
    s.addText(head, {
      x: x + 0.26, y: y + 0.16, w: w - 0.52, h: 0.38, isTextBox: true, margin: 0,
      valign: "middle", fontFace: F, fontSize: 15, bold: true, color: BRICK
    });
    s.addText(lines, {
      x: x + 0.26, y: y + 0.62, w: w - 0.52, h: h - 0.80, isTextBox: true, margin: 0,
      fontFace: F, fontSize: bodyOpt.fontSize, color: INK, lineSpacing: bodyOpt.lineSpacing
    });
  }

  planCard(M, 2.12, cwC, 2.18, "時間が足りない",
    "・資料カードを12枚→6枚に（光1・4・5／影1・2・6）\n・スライド17の対応表は上3行だけ提示する\n・スライド25（じゃあ、やめられる？）を省く\n・最終手段：主発問を次時の冒頭へ回す");
  planCard(M + cwC + 0.40, 2.12, cwC, 2.18, "子どもの発言が出ない",
    "・展開2：「今の日本で子どもが工場で働かないのはなぜ？」\n・展開1：「迷い」に置いたカードから指名する\n・討論：「全部が光」の子に影2、「全部が影」の子に光1");
  planCard(M, 4.40, cwC, 2.18, "意見が一方に偏る",
    "・「幸せにした」に偏ったら → スライド24\n・「していない」に偏ったら → スライド23\n・どちらも動かないとき → スライド25");
  planCard(M + cwC + 0.40, 4.40, cwC, 2.18, "教師が気をつけること",
    "・主発問に自分の答えを言わない。最後まで言わない\n・数値は概数。教科書・資料集と照合してから使う\n・「平均死亡年齢」を「平均寿命」と言い換えない\n・安い服を買う消費者個人の責任に落とさない");
  notes(s, "【最優先で守るもの】\n時間が押したときに何を削るかを、授業前に決めておいてください。削ってよいのは展開1の資料枚数とスライド25です。削ってはいけないのは、スライド22（主発問の個人思考2分）とスライド27（結論）です。\n\nこの2つを確保できれば、本時のねらいは達成されます。逆に、展開1を完璧にやって結論が駆け足になると、本時は何も残りません。\n\n【最後に】\n本時は情報量が多い授業です。1時間で通すのは可能ですが、2時間扱い（切れ目はスライド18）のほうが、子どもの思考は確実に深まります。学級の実態に合わせて選んでください。");
}

pres.writeFile({ fileName: "令和8年度_第2学年_社会科_産業革命_授業スライド.pptx" })
  .then(f => console.log("written:", f));
