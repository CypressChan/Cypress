//如果不是全局就得引入fs成员
const fs = require("fs");
const path = require("path");
const emlformat = require('eml-format');
const cheerio = require('cheerio');
const textract = require('textract');//对于docx文件，您可以使用textract，它将从.docx文件中提取文本。
const excelPort = require('excel-export');
const admZip = require('adm-zip');

//解析eml文件
async function readEml(path){
	var eml = fs.readFileSync(path, "utf-8");
	var info = null;
	await emlformat.read(eml, function(error, data) {
	  if (error) return console.log(error);
	  //console.log(data);
	  const $ = cheerio.load(data.html);
	  info = {
		  phone: $('td:contains("手机：")').next().eq(1).text(),
		  city: $('td:contains("地点：")').next().text(),
		  name: $('strong').eq(0).text()
	  }
	  //console.log(data.html);
	  //console.log(info);
	});
	return info;
}

//解析word
function readWord(path, filename) {
    return new Promise(function(resolve, reject){
		textract.fromFileWithPath(path, function (error, text) {
			let info = {};
			if (error) {
				console.log(error);
			} else {
				//console.log(text);
				if(filename.indexOf('智联招聘') > -1){
					let arr = filename.split('_');
					let phoneStartIndex = text.indexOf('手机：') > -1 ? text.indexOf('手机：')+3 : -1;
					let phoneEndIndex = text.indexOf(' ', phoneStartIndex);
					info = {
					  phone: phoneStartIndex !== -1 ? text.substring(phoneStartIndex, phoneEndIndex) : '',
					  city: arr[arr.length-2].match(/[\u4e00-\u9fa5]/g).join(''),
					  name: arr[1]
					}
				}
			}
			//console.log(info);
			resolve(info);
		})
	}).catch(e=>{
		console.log(e);
	});
 }

function mapDir(dir) {
	
  return new Promise(function(resolve, reject){
	  let promises = [];
	  fs.readdir(dir, function(err, files) {
		if (err) {
		  console.error(err)
		  return
		}
		promises = files.map(async (filename, index) => {
		  let pathname = path.join(dir, filename);
		  let extension = pathname.substring(pathname.lastIndexOf('.') + 1);
		  let info = {
			  fileName: filename,
			  phone: '',
			  city: '',
			  name: ''
		  };
		  if(extension === 'eml'){
			  Object.assign(info, await readEml(pathname));
		  }else if(['docx'].includes(extension)){
			  Object.assign(info, await readWord(pathname, filename));
		  }
		  return info;
		});
		 Promise.all(promises).then((infos)=>{
			 let info = infos.filter(n=>n);
			 resolve(info);
		 });
	  });
  });
}

function generateExcel(datas){
   /**
    * 定义一个空对象，来存放表头和内容
    * cols，rows为固定字段，不可修改
    */
   const excelConf = {
     cols: [], // 表头
     rows: [], // 内容
   };
   // 表头
   for(let key in datas[0]){
     excelConf.cols.push({
       caption: key,
       type: 'string', // 数据类型
       width: 200, // 宽度
     })
   }
   // 内容
   datas.forEach(item => {
     // 解构
     excelConf.rows.push(Object.values(item));
   })
   // 调用excelPort的方法，生成最终的数据
   const result = excelPort.execute(excelConf);
   // 写文件
   fs.writeFile('./Cypress.xlsx', result, 'binary', err => {
     if(!err){
       console.log('生成成功！')
     }
   })
}

function parseDocx(path){
	// 解压word文档
	const zip = new admZip(path);
	zip.extractAllTo('./output/2', true);
	// 提取内容
	let contentXml = zip.readAsText("word/document.xml");
	// 正则匹配文字
	let matchWT = contentXml.match(/(<w:t>.*?<\/w:t>)|(<w:t\s.[^>]*?>.*?<\/w:t>)/gi);
	matchWT = filter(matchWT);
	console.log(matchWT);
	/**
	 * 去除空白行
	 * @param matchWT
	 */
	function filter(matchWT) {
		let res = [];
	 
		matchWT.forEach(function(wtItem) {
			//如果不是<w:t xml:space="preserve">格式
			if (wtItem !== '<w:t xml:space="preserve"> </w:t>') { 
				wtItem = wtItem.split('>');
				wtItem = wtItem[1].split('<');
				res.push(wtItem[0]);
			}
		});
	 
		return res;
	}
}
//parseDocx();
//readWord();

mapDir('./files').then(function(arrData){
	if(arrData.length){
	  console.log(arrData);
	  generateExcel(arrData);		
	}
});

