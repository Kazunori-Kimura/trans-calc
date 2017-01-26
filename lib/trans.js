const fs = require("fs");
const koyomi = require("koyomi");
const officegen = require("officegen");
const xlsx = officegen("xlsx");

xlsx.on("finalize", (written) => {
  console.log(`Finish to create an Excel file.`);
});
xlsx.on("error", (err) => {
  console.error(err);
});

/**
 * ルール定義
 */
const rules = [
  // A社 (2016/1/1 ~ 4/30)
  // 三宮 ~ 本町 片道 ¥570
  {
    /**
     * 判定
     * 
     * @param {Date} d
     * @return {boolean}
     */
    "is": (d) => {
      const d2 = new Date(2016, 4, 1); //2016-5-1
      if ( d >= d2) {
        return false;
      }
      return koyomi.isOpen(d); //平日
    },
    "mokuteki": "A",
    "keiro": "三宮〜本町",
    "kingaku": 1140
  },
  // B社 (2016/5/1 ~)
  // 三宮 ~ 京都 片道 ¥1,230
  {
    /**
     * 判定
     * 
     * @param {Date} d
     * @return {boolean}
     */
    "is": (d) => {
      const d2 = new Date(2016, 4, 1); //2016-5-1
      if (d < d2) {
        return false;
      }
      return d.getDay() === 5; //毎週金曜日
    },
    "mokuteki": "B社",
    "keiro": "三宮〜京都",
    "kingaku": 2460
  },
  // C社 (2016/1/1 ~)
  // 三宮 ~ 北浜 片道 ¥720
  {
    /**
     * 判定
     * 
     * @param {Date} d
     * @return {boolean}
     */
    "is": (d) => {
      return d.getDay() === 6; //毎週土曜日
    },
    "mokuteki": "C社",
    "keiro": "三宮〜北浜",
    "kingaku": 1440
  }
];



/**
 * 生成処理本体
 * 
 * @param {number} year - 年
 * @param {string} file - 出力ファイル
 */
function generater(year, file) {
  const d = new Date(year, 0, 1); //指定された年の1/1を取得

  let currentMonth = -1;
  let sheet = null;
  let amount = 0;
  let row = 0;

  while(d.getFullYear() === year) {
    if (d.getMonth() !== currentMonth) {
      // 月が変わった
      if (sheet) {
        sheet.data[row] = ["", "", "", "合計", amount];
      }

      // シート作成
      const title = koyomi.format(d, "YYYY年MM月");
      sheet = xlsx.makeNewSheet();
      row = 0;
      amount = 0;
      sheet.name = title;
      sheet.data[row++] = [title];
      sheet.data[row++] = ["日", "曜日", "目的", "経路", "費用"];

      currentMonth = d.getMonth();
    }

    const dd = koyomi.format(d, "DD");
    const ww = koyomi.format(d, "W").charAt(0);

    let existsData = false;
    rules.forEach((rule) => {
      if (rule.is(d)) {
        sheet.data[row++] = [dd, ww, rule.mokuteki, rule.keiro, rule.kingaku];
        amount += rule.kingaku;
        existsData = true;
      }
    });

    // データが無ければ空とする
    if (!existsData) {
      sheet.data[row++] = [dd, ww];
    }
    d.setDate(d.getDate() + 1); //翌日
  }

  // 最終月の合計を出力
  sheet.data[row] = ["", "", "", "合計", amount];

  // 出力
  const out = fs.createWriteStream(file);

  out.on("error", (err) => {
    console.error(err);
  });

  xlsx.generate(out);
}

module.exports = generater;

