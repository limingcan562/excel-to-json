const Tool = require('./createJsonTool');

Tool.create({
    importPath: '../excel/',
    outPath: '../json/',
    outJsonName: ['data1', 'data2', 'data3'],
    excelFileName: '配置文件',
    fixedKeyName: 'fixedKey' 
});