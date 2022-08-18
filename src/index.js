const fs = require('fs');
const xlsx = require('node-xlsx');

// 解析得到文档中的所有 sheet
const sheets1 = xlsx.parse('../excel/人员信息一.xlsx');
const sheets2 = xlsx.parse('../excel/人员信息二.xlsx');

const sheet1 = sheets1[0];
const sheet2 = sheets2[0];

const sheetData1 = sheet1.data;
const sheetData2 = sheet2.data;

let idList1 = sheetData1.map((i) => i[1]);
idList1 = idList1.slice(1);

let idList2 = sheetData2.map((i) => i[1]);
idList2 = idList2.slice(1);

const repeatIdList = [];

idList1.forEach((i) => {
  if (!idList2.includes(i)) {
    repeatIdList.push(i);
  }
});

const bufferArr = [];

sheetData1.forEach((i, index) => {
  if (index === 0) bufferArr.push(i);
  if (repeatIdList.includes(i[1])) {
    bufferArr.push(i);
  }
});

const bufferData = [{
  name: '人员汇总',
  data: bufferArr,
}];

const buffer = xlsx.build(bufferData);

fs.writeFile('../excel/人员信息对比.xlsx', buffer, (e) => {
  if (e) {
    console.log(`写入失败：${e}`);
    return;
  }
  console.log('写入成功');
});
