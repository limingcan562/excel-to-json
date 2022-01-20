const Tool = require('./createJsonTool');

Tool.create({
    importPath: '../excel/',
    outPath: '../json/',
    outJsonName: ['config1', 'config2', 'config3'],
    fileName: '配置文件',
    fixedKeyName: 'fixedKey' 
});