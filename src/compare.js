const fs = require('fs');
const xlsx = require('node-xlsx');

const fileA = '../files/驿站员工数据.xlsx';
const fileB1 = '../files/1人员列表.csv';
const fileB2 = '../files/2人员列表.csv';
const fileB3 = '../files/3人员列表.csv';
const targetFile = '../files/人员信息对比.xlsx';

// 解析得到文档中的所有 sheet
const sheeetsA = xlsx.parse(fileA); // sheeetsA:
// 取第一个 sheet
const sheeetA = sheeetsA[0];
// 取 data
const sheeetDataA = sheeetA.data;
// 取所有id
let idListA = sheeetDataA.map((i) => i[0]);
idListA = idListA.slice(1);

console.log('读取表A数据完成');

const sheetsB1 = xlsx.parse(fileB1);
const sheetB1 = sheetsB1[0];
const sheetDataB1 = sheetB1.data;
const idListB1 = sheetDataB1.map((i) => i[0]);
const idMapB1 = new Map();
idListB1.forEach((id) => {
  idMapB1.set(id, 1);
});

console.log('读取表B1数据完成');

const sheetsB2 = xlsx.parse(fileB2);
const sheetB2 = sheetsB2[0];
const sheetDataB2 = sheetB2.data;
const idListB2 = sheetDataB2.map((i) => i[0]);
const idMapB2 = new Map();
idListB2.forEach((id) => {
  idMapB2.set(id, 1);
});

console.log('读取表B2数据完成');

const sheetsB3 = xlsx.parse(fileB3);
const sheetB3 = sheetsB3[0];
const sheetDataB3 = sheetB3.data;
const idListB3 = sheetDataB3.map((i) => i[0]);
const idMapB3 = new Map();
idListB3.forEach((id) => {
  idMapB3.set(id, 1);
});

console.log('读取表B3数据完成');

const idList = []; // 表 A 中包含表 B 不包含
idListA.forEach((i) => {
  if (!idMapB1.get(i) && !idMapB2.get(i) && !idMapB3.get(i)) {
    idList.push(i);
  }
});

console.log('取表 A 中包含表 B 不包含的id完成');

const idMap = new Map();
idList.forEach((id) => {
  idMap.set(id, 1);
});
const bufferArr = [];
sheeetDataA.forEach((i) => {
  if (idMap.get(i[0])) {
    bufferArr.push(i);
  }
});

const bufferData = [{
  name: '人员汇总',
  data: bufferArr,
}];
const buffer = xlsx.build(bufferData);

console.log('组装buffer数据完成');

fs.writeFile(targetFile, buffer, (e) => {
  if (e) {
    console.log(`写入失败：${e}`);
    return;
  }
  console.log('写入成功');
});
