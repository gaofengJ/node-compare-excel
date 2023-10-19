const fs = require('fs');
const xlsx = require('node-xlsx');

const filePathTrain = '../files/train.xlsx';
const targetFile = '../files/培训课时汇总.xlsx';

// 解析得到文档中所有sheet
const sheetsTrain = xlsx.parse(filePathTrain);

// 取第一个sheet
const sheetTrain = sheetsTrain[0];

// 取data
const sheetData = sheetTrain.data;

// 保存培训名称
const trainNameMap = new Map();

// 表头
const tableHeader = sheetData[0];

// 保存人员数据
const trainArr = sheetData.slice(1).map((i) => {
  const [tempNum, tempName, tempGender, tempTrainName, tempTotal, tempLevel] = i;
  if (!trainNameMap.has(tempTrainName)) {
    trainNameMap.set(tempTrainName, {
      tempTrainName,
      tempLevel,
    });
  }
  return {
    num: tempNum,
    name: tempName.trim(),
    gender: tempGender,
    trainName: tempTrainName,
    total: tempTotal,
    tempLevel,
  };
});

tableHeader.splice(3, 3, ...trainNameMap.keys(), '合计'); // 修改列名称

const combinedTrainArr = [];

for (let i = 0; i < trainArr.length; i++) {
  const tempTrainData = trainArr[i];
  const tempCombinedTrainData = combinedTrainArr.find((j) => j.name === tempTrainData.name);
  if (tempCombinedTrainData) {
    tempCombinedTrainData[tempTrainData.trainName] = tempTrainData.total;
  } else {
    combinedTrainArr.push({
      ...tempTrainData,
      [tempTrainData.trainName]: tempTrainData.total,
    });
  }
}

const bufferArr = [tableHeader]; // 保存最后的写入数据
const tempTrainNameArr = [...trainNameMap.keys()];

for (let i = 0; i < combinedTrainArr.length; i++) {
  const tempCombinedTrainData = combinedTrainArr[i];
  const tempTotalData = [];
  for (let j = 0; j < tempTrainNameArr.length; j++) {
    tempTotalData.push(tempCombinedTrainData[tempTrainNameArr[j]] || 0);
  }
  bufferArr.push([
    i + 1,
    tempCombinedTrainData.name,
    tempCombinedTrainData.gender,
    ...tempTotalData,
    tempTotalData.reduce((sum, total) => sum + total, 0),
  ]);
}

const bufferData = [
  {
    name: '培训课时汇总',
    data: bufferArr,
  },
];

const buffer = xlsx.build(bufferData);

console.log('组装buffer数据完成');

fs.writeFile(targetFile, buffer, (e) => {
  if (e) {
    console.log(`写入失败：${e}`);
    return;
  }
  console.log('写入成功');
});
