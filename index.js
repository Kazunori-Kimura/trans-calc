#!/usr/bin/env node

/*
 * transportation cost calculation
 */
const commander = require("commander");
const trans = require("./lib/trans");

commander.version("1.0.0")
  .usage("[options]")
  .option("-y, --year <year>", "対象年を指定", parseInt)
  .option("-f, --file <file>", "出力ファイル名")
  .parse(process.argv);

let year, file;

if (commander.year) {
  year = commander.year;
} else {
  year = (new Date()).getFullYear();
}

if (commander.file) {
  file = commander.file;
} else {
  file = `交通費_${year}.xlsx`;
}

trans(year, file);
