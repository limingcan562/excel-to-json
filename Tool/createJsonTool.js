// 把excel 生成json 文件
const xlsxrd = require('node-xlsx');
const path = require('path');
const fs = require('fs');
const colors = require('colors');
const ora = require('ora');


/**
 * @startRow 从第几行开始拿值(默认第2行)
 * @outPath 输出位置
 * @fileName 要转换的excel表格
 */


module.exports = ({
    startRow = 2,
    fileName,
    importPath = './doc',
    outPath,
    fixedKeyName,
    commonKey = 'category'
}) => {
    // 读取excel中所有工作表的数据
    const 
    xlsxList = xlsxrd.parse(path.join(__dirname, `${importPath}${fileName}.xlsx`)),
    finalResult = [];

    // console.log(xlsxList[0].data);

    xlsxList.forEach(item => {
        const 
        xlsxData = item.data,
        xlsxName = item.name,
        keyData = xlsxData[startRow - 1], // 每项数据里的key
        renderData = xlsxData.slice(2), // 每项数据里的key 里面对应的value
        fixedKeyIndex = keyData.findIndex(item => item === fixedKeyName),
        xlsxEveryData = {},
        categoryData = [];

        xlsxEveryData['xlsxName'] = xlsxName;
        // console.log(`${'[info]'.bold} 开始生成${xlsxName}.json文件`.yellow);
        const createAni = ora(`正在生成${xlsxName}.json文件`);
        // createAni.start();
        // console.log(keyData);
        // console.log(renderData);
        // console.log(fixedKeyIndex);
        // console.log(xlsxData);


        if (fixedKeyIndex === -1) {
            createAni.fail('请确定Excel表格有'+ fixedKeyName + '字段');
            return;
        }
        
        // 先把相同的 类别 归类
        renderData.forEach((renderDataItem) => {
            // console.log(renderDataItem);
            let obj = {};
            renderDataItem.forEach((item2, index) => {
                const key = keyData[index];
                obj[key] = item2;
            });
            categoryData.push(obj);
        });
        // console.log(categoryData);

        // 获得最后渲染的数据
        categoryData.forEach((item, index) => {
            const key = item[commonKey];
            if (!xlsxEveryData[key]) xlsxEveryData[key] = [];

            let obj = {};
            keyData.forEach(keyName => {
                obj[keyName] = item[keyName] ? item[keyName] : '';
            });
            xlsxEveryData[key].push(obj);
        });
        // console.log(xlsxEveryData);
        finalResult.push(xlsxEveryData);
    });

    // console.log(finalResult);

    finalResult.forEach(item => {
        const {xlsxName} = item;
        new Promise((resolve, reject) => {
            fs.writeFile(path.join(__dirname, `${outPath}${xlsxName}.json`), JSON.stringify(item), err => {
                if (err) {
                    // throw err;
                    reject(err)
                } else {
                    // console.log(`${'[success]'.bold} ${outName}.json生成成功`.green);
                    resolve();
                }
            });
        })
        .then(() => {
            console.log(`${'[success]'.bold} ${xlsxName}.json文件生成成功`.green);
        })
        .catch(err => {
            // console.err(JSON.stringify(`${xlsxName}.json文件生成错误：${JSON.stringify(err)}`));
            console.log(err);
        })
    });
}