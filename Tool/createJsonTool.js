// 把excel 生成json 文件
const xlsxrd = require('node-xlsx');
const path = require('path');
const fs = require('fs');
const colors = require('colors');
const ora = require('ora');


module.exports = ({
    fileName,
    importPath = './doc',
    outPath,
    commonKey = 'category'
}) => {
    // 读取excel中所有工作表的数据
    const 
    xlsxList = xlsxrd.parse(path.join(__dirname, `${importPath}${fileName}.xlsx`)),
    finalResult = [];
    
    xlsxList.forEach(item => {
        const 
        xlsxData = item.data,
        xlsxName = item.name,
        keyData = xlsxData[1],
        renderData = xlsxData.slice(2),
        xlsxEveryData = {},
        categoryData = [];

        xlsxEveryData['xlsxName'] = xlsxName;
        // console.log(`${'[info]'.bold} 开始生成${xlsxName}.json文件`.yellow);
        const createAni = ora(`正在生成${xlsxName}.json文件`);
        // console.log(renderData);
        
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