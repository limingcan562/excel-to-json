const Tool = require('./createJsonTool');

Tool.create({
    importPath: '../excel/',
    outPath: '../json/',
    outJsonName: ['自定义名字1', 'confing2', '自定义名字3'],
    fileName: '配置文件',
    fixedKeyName: 'fixedKey' 
});