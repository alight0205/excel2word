// 引入相关的库
var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');
var xlsx = require('node-xlsx');
var fs = require('fs');
var path = require('path');
var data = [];
var templatePath;
var outputDirname;

var config = require('./config.json');

process.argv.forEach((val, index) => {
    console.log(`${index}: ${val}`);
    data.push(val);
});
templatePath = data[2]; //第三个参数
outputDirname = path.dirname(templatePath);

// 读取文件,以二进制文件形式保存
var content = fs.readFileSync(templatePath, 'binary');
// 压缩数据
var zip = new PizZip(content);
// 生成模板文档
var doc = new Docxtemplater(zip);
// 设置填充数据
doc.setData(config);
//渲染数据生成文档
doc.render();
// 将文档转换文nodejs能使用的buf
var buf = doc.getZip().generate({ type: 'nodebuffer' });
// 输出文件
fs.writeFileSync(path.resolve(outputDirname, 'output.docx'), buf);
