const fs = require('fs');
const xlsx = require('node-xlsx');

// 解析得到文档中的所有 sheet
const sheets1 = xlsx.parse('../excel/A.xlsx');
// 取第一个 sheet
const sheet1 = sheets1[0];
// 取 data
const sheetData1 = sheet1.data;
// 取所有id
let idList1 = sheetData1.map((i) => i[0]);
idList1 = idList1.slice(1);

console.log('读取表A数据完成');

const sheets2 = xlsx.parse('../excel/B.xls');
const sheet2 = sheets2[0];
const sheetData2 = sheet2.data;
let idList2 = sheetData2.map((i) => i[0]);
idList2 = idList2.slice(2);

console.log('读取表B数据完成');

const idList = []; // 表 A 中包含表 B 不包含
idList1.forEach((i) => {
  if (!idList2.includes(i)) {
    idList.push(i);
  }
});

console.log('取表 A 中包含表 B 不包含的id完成');

const bufferArr = [];
sheetData1.forEach((i, index) => {
  if (index === 0) bufferArr.push(i);
  if (idList.includes(i[0])) {
    bufferArr.push(i);
  }
});
const bufferData = [{
  name: '人员汇总',
  data: bufferArr,
}];
const buffer = xlsx.build(bufferData);

console.log('组装buffer数据完成');

fs.writeFile('../excel/人员信息对比.xlsx', buffer, (e) => {
  if (e) {
    console.log(`写入失败：${e}`);
    return;
  }
  console.log('写入成功');
});
