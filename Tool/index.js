const createJson = require('./createJsonTool');
createJson({
    importPath: '../excel/',
    outPath: '../json/',
    fileName: '配置文件',
    fixedKeyName: 'fixedKey' 
});