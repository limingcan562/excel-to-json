// 把excel 生成json 文件
const xlsxrd = require('node-xlsx');
const path = require('path');
const fs = require('fs');
const ora = require('ora');
const chalk = require('chalk');

/**
 * @startRow 从第几行开始拿值(默认第2行)
 * @outPath 输出位置
 * @fileName 要转换的excel表格
 */

module.exports = {
    async create({startRow = 2, fileName, importPath, outPath, outJsonName, fixedKeyName}) {
        // console.log(this);
        // 判断表格路径是否存在
        if (!this.judgeExist(importPath)) {
            ora(`请确认${chalk.red(importPath)}路径是否存在`).fail();
            return;
        }

        // 读取excel中所有工作表的数据
        const 
        xlsxList = xlsxrd.parse(path.join(__dirname, `${importPath}${fileName}.xlsx`)),
        finalResult = []; // 最后要输出的json数据，有多少个sheet 就有多少个json

        // 输出文件夹存在 --> 清空文件夹内容（不用unlinkSync是因为权限问题）
        if (this.judgeExist(outPath)) {
            fs.rmdirSync(path.join(__dirname, outPath), {recursive : true});
        }
        // 输出文件夹不存在，创建文件夹
        fs.mkdirSync(path.join(__dirname, outPath), true); 
        


        // console.log(xlsxList[0].data);
        // console.log(xlsxList);
        for (let index = 0; index < xlsxList.length; index++) {
            const 
            item = xlsxList[index]
            xlsxData = item.data.filter(item => item.length !== 0), // Excel 转成的一行行数据
            xlsxName = item.name ? item.name : `表格${index + 1}`, // Excel 表格里面的名字
            keyData = xlsxData[startRow - 1], // 每项数据里的key
            renderData = xlsxData.slice(startRow), // 每项数据里的key 里面对应的value
            fixedKeyIndex = keyData.findIndex(item => item === fixedKeyName), // Excel里面是不是有固定字段
            finalXlsxEveryData = {}, // Excel 最后得得到的每个 sheet 组装好的数据
            categoryData = [];  // 将每个值以key: value && 含有类别的数组


            finalXlsxEveryData['xlsxName'] = xlsxName;

            // console.log(fixedKeyIndex);
            const createAni = ora(`正在重组${chalk.yellow(xlsxName)} 表格文件`);
            if (fixedKeyIndex === -1) {
                createAni.fail(`请检查Excel表格有${chalk.red(fixedKeyName)}字段`);
                break;
            }

            // console.log(xlsxData);
            createAni.start();
            // 先将类别放进去自己组建的数组对象
            renderData.forEach(singleArr => {
                const prefixKey = singleArr[fixedKeyIndex];
                // console.log(prefixKey);
                // console.log(singleArr);
                let obj = {};
                singleArr.forEach((item2, index) => {
                    const key = keyData[index];
                    obj[key] = item2;
                });
                categoryData.push(obj);
            });
            // createAni.text = '获取到重组含类别数据';
            // console.log(categoryData);
            // console.log(finalXlsxEveryData);


            // 重组最后要的数据
            categoryData.forEach(item => {
                const commonKey = item[fixedKeyName];
                // console.log(commonKey);
                // ! 如果没有这个key的数组，则创建这个数组
                if (!finalXlsxEveryData[commonKey]) finalXlsxEveryData[commonKey] = [];

                // ! 如果有这个key的数组，则添加
                let 
                obj = {},
                defaultValue = '';
                keyData.forEach(keyName => {
                    if (keyName !== fixedKeyName) {
                        obj[keyName] = item[keyName] ? item[keyName] : defaultValue;
                    }
                });
                finalXlsxEveryData[commonKey].push(obj);
            });
            // console.log(finalXlsxEveryData);

            // 获取最后的数据
            // createAni.text = '获取最后的数据';
            finalResult.push(finalXlsxEveryData);

            // 开始输出json
            createAni.succeed(`${chalk.yellow(xlsxName)} 表格数据重组成功`);
            let exportAni = ora(`正在输出${chalk.yellow(xlsxName)} 表格json文件`).start();
            this.exportJson(finalXlsxEveryData, outPath, outJsonName[index] ? outJsonName[index] : xlsxName, exportAni);
        }
    },

    // 输出json
    exportJson(everyList, outPath, outJsonName, exportAni) {
        // console.log(everyList);
        // 开始输出json
        new Promise((resolve, reject) => {
            fs.writeFile(path.join(__dirname, `${outPath}${outJsonName}.json`), JSON.stringify(everyList), err => {
                if (err) {
                    // throw err;
                    reject(err)
                } else {
                    // console.log(`${'[success]'.bold} ${outJsonName}.json生成成功`.green);
                    resolve();
                }
            });
        })
        .then(() => {
            exportAni.succeed(`${chalk.yellow(`${outJsonName}.json`)} 文件生成成功`);
        })
        .catch(err => {
            // console.err(JSON.stringify(`${xlsxName}.json文件生成错误：${JSON.stringify(err)}`));
            console.log(err);
        });
    },

    // 判断文件存不存在
    judgeExist(outPath) {
        let flag = false;
        try {
            fs.accessSync(path.join(__dirname, outPath));
            // console.log('can read/write');
            flag = true;
            return flag;
        } catch (err) {
            // console.error('no access!');
            flag = false;
            return flag;
        }
    }
}