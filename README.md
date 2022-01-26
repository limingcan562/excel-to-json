# `excel-to-json`
将`Excel`数据配置表格转换成`Json`数据，方便数据维护

适用场景：
- 静态商场首页，可用表格每次更新不同商品数据，再生成最新`json`更新到服务器即可
- `H5`中一些静态数据的展示，例如每期更新`banner`位置广告等，可免去后台开发成本
- 可方便配置页面中每块区域要展示的数据

**特点：**  
生成的数据格式为：`{name: 'config', data1: [], data2: []}`
- 以键值对方式存在，可读性高，更方便取值
- 数据会提取表格内相同的字段，作为数据的`key`，这样可以更好的配置页面中每一块的数据

## 用法说明
1. `npm i` 先安装依赖
2. `npm run create`，可以输出`json`格式的表格数据

## 方法入参说明
|  参数名称  | 类型 | 默认值 | 说明  
|  :-:  | :-: | :-: | --- |
| `importPath`  |`String` | `../excel/` | 表格存放的路径 
| `excelFileName` |`String` | 配置文件 | 要编译的`Excel`表格文件名 
| `outPath` |`String` |`../json/` | 输出的`json`文件路径 |
| `outJsonName` |`Array` | `['data1', 'data2', 'data3']` | 输出名为`data1,data2,data3`的`json`，若值为`''`，则输出的`json`的名字以`Sheet`里的命名为主
| `fixedKeyName` | `String` | `fixedKey` |  `Excel`表格内**必须有此字段**，这样才能将相同`key`值的数据归类