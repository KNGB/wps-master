/**
 * 这是为了便于Demo能在开发者的本地快速运行起来，采用Nodejs模拟服务端，开发者可根据此文件的注释，自行在业务系统中实现对应的功能
 */
const express = require('express');
const fs = require('fs');
const path = require('path');
var urlencode = require('urlencode');
const formidable = require('formidable')
var ini = require('ini')
var regedit = require('regedit')
const os = require('os');
const app = express()
var cp = require('child_process');

//----开发者将WPS加载项集成到业务系统中时，需要实现的功能 Start--------
/** 
 * 支持jsplugins.xml中，在线模式下，WPS加载项的请求地址
 * 开发者可在业务系统部署时，将WPS加载项合并部署到服务端，提供jsplugins.xml中对应的请求地址即可
 */
app.all('*', function (req, res, next) {
	res.header('Access-Control-Allow-Origin', '*');
	console.log(getNow()+req.originalUrl)
	// res.setHeader('Content-Type','text/plain;charset=gbk');
	//Access-Control-Allow-Headers ,可根据浏览器的F12查看,把对应的粘贴在这里就行
	// res.header('Access-Control-Allow-Headers', 'Content-Type');
	// res.header('Access-Control-Allow-Methods', '*');
	// res.header('Content-Type', 'text/html;charset=utf-8');
	next();
});
app.use(express.static(path.join(__dirname, "wwwroot"))); //wwwroot代表http服务器根目录
app.use('/plugin/et', express.static(path.join(__dirname, "../EtOAAssist")));
app.use('/plugin/wps', express.static(path.join(__dirname, "../WpsOAAssist")));
app.use('/plugin/wpp', express.static(path.join(__dirname, "../WppOAAssist")));

/** 
 * 文件下载
 * 开发者可将此方法更换为自己业务系统的文件下载方法， 请求中必须的参数为： filename
 * 如果业务系统有特殊的要求， 此方法可与各加载项中的 js / common / common.js: DownloadFile 方法对应着修改
 */
app.use("/Download/:fileName", function (request, response) {
	var fileName = request.params.fileName;
	var filePath = path.join(__dirname, './wwwroot/file');
	filePath = path.join(filePath, fileName);
	var stats = fs.statSync(filePath);
	if (stats.isFile()) {
		let name = urlencode(fileName, "utf-8");
		response.set({
			'Content-Type': 'application/octet-stream',
			//'Content-Disposition': "attachment; filename* = UTF-8''" + name,
			'Content-Disposition': "attachment; filename=" + name,
			'Content-Length': stats.size
		});
		fs.createReadStream(filePath).pipe(response);
		console.log(getNow() + "下载文件接被调用，文件路径：" + filePath)
	} else {
		response.writeHead(200, "Failed", {
			"Content-Type": "text/html; charset=utf-8"
		});
		response.end("文件不存在");
	}
});
/** 
 * 文件上传
 * 开发者可将此方法更换为自己业务系统的接受文件上传的方法
 * 如果业务系统有特殊的要求， 此方法可与加载项中的 js / common / common.js: UploadFile 方法对应着修改
 */
app.post("/Upload", function (request, response) {
	const form = new formidable.IncomingForm();
	var uploadDir = path.join(__dirname, './wwwroot/uploaded/');
	form.encoding = 'utf-8';
	form.uploadDir = uploadDir;
	console.log(getNow() + "上传文件夹地址是：" + uploadDir);
	//判断上传文件夹地址是否存在，如果不存在就创建
	if (!fs.existsSync(form.uploadDir)) {
		fs.mkdirSync(form.uploadDir);
	}
	form.parse(request, function (error, fields, files) {
		for (let key in files) {
			let file = files[key]
			// 过滤空文件
			if (file.size == 0 && file.name == '') continue

			var fileName = file.name
			if (!fileName)
				fileName = request.headers.filename
			let oldPath = file.path
			let newPath = uploadDir + fileName

			fs.rename(oldPath, newPath, function (error) {
				console.log(getNow() + "上传文件成功，路径：" + newPath)
			})
		}
		response.writeHead(200, {
			"Content-Type": "text/html;charset=utf-8"
		})
		response.end("测试");
	})
});

/** 
 * 模拟填充到文件的服务端数据（ 加载本地的模拟json数据并提供请求）
 * 开发者可将此方法更换为自己业务系统的数据接口
 * json数据格式及解析， 可与各加载项中的 js / common / func_docProcess.js: GetServerTemplateData 方法对应着修改
 */
app.get('/getTemplateData', function (request, response) {
	var file = path.join(__dirname, './wwwroot/file/templateData.json');
	//读取json文件
	fs.readFile(file, 'utf-8', function (err, data) {
		if (err) {
			response.send('文件读取失败');
		} else {
			response.send(data);
		}
	});
});
//----开发者将WPS加载项集成到业务系统中时，需要实现的功能 End--------


//获取file目录下文件列表
app.use("/FileList", function (request, response) {
	var filePath = path.join(__dirname, './wwwroot/file');
	fs.readdir(filePath, function (err, results) {
		if (err) {
			response.writeHead(200, "OK", { "Content-Type": "text/html; charset=utf-8" });
			response.end("没有找到file文件夹");
			return;
		}
		if (results.length > 0) {
			var files = [];
			results.forEach(function (file) {
				if (fs.statSync(path.join(filePath, file)).isFile()) {
					files.push(file);
				}
			})
			response.writeHead(200, "OK", { "Content-Type": "text/html; charset=utf-8" });
			response.end(files.toString());
		} else {
			response.writeHead(200, "OK", { "Content-Type": "text/html; charset=utf-8" });
			response.end("当前目录下没有文件");
		}
	});
});

//wps安装包是否正确的检测
app.use("/WpsSetup", (request, response) => {
	response.writeHead(200, "OK", { "Content-Type": "text/html; charset=utf-8" })
	response.end("成功");
});

//wps加载项配置是否正确的检测
app.use("/OAAssistDeploy", (request, response) => {
	response.writeHead(200, "OK", { "Content-Type": "text/html; charset=utf-8" })
	response.end("成功");
});
//检测WPS客户端环境
app.use("/WpsSetupTest", function (request, response) {
	configOem(function (res) {
		response.writeHead(200, res.status, {
			"Content-Type": "text/html;charset=utf-8"
		});
		response.write('<head><meta charset="utf-8"/></head>');
		response.write("<br/>当前检测时间为： " + getNow() + "<br/>");
		response.end(res.msg);
	});
});
//定义node服务端口
var server = app.listen(3888, function () {
	console.log(getNow() + "启动本地web服务(http://127.0.0.1:3888)成功！");
	let url="http://127.0.0.1:3888/index.html";
	let exec=cp.exec;
	try{
		switch (process.platform) {
			//mac系统使用 一下命令打开url在浏览器
			case "darwin":
				exec(`open ${url}`);
				break;
			//win系统使用 一下命令打开url在浏览器
			case "win32":
				exec(`start ${url}`);
				break;
			//linux系统使用 一下命令打开url在浏览器
			case "linux":
				exec(`xdg-open ${url}`)
				break;
				// 默认linux系统
			default:
				exec(`xdg-open ${url}`)
				break;
		}
	}catch(e){
	}
});
//启动node服务
server.on('error', (e) => {
	if (e.code === 'EADDRINUSE') {
		console.log('地址正被使用，重试中...');
		setTimeout(() => {
			server.close();
			server.listen(3888);
		}, 2000);
	}
});
//获取当前时间
function getNow() {
	let nowDate = new Date()
	let year = nowDate.getFullYear()
	let month = nowDate.getMonth() + 1
	let day = nowDate.getDate()
	let hour = nowDate.getHours()
	let minute = nowDate.getMinutes()
	let second = nowDate.getSeconds()
	return year + '年' + month + '月' + day + '日 ' + hour + ':' + minute + ':' + second + "  "
}
//配置WPS客户端的WPS加载项的配置
//此功能开发者无需实现和考虑，在生产环境中，WPS客户端的配置文件可通过
//独立打包或开发者编写批处理命令实现修改，业务系统无法实现修改WPS客户端配置文件
function configOemFileInner(oemPath, callback) {
	var config = ini.parse(fs.readFileSync(oemPath, 'utf-8'))
	var sup = config.support || config.Support;
	var ser = config.server || config.Server;
	var needUpdate = false;
	if (!sup || !sup.JsApiPlugin || !sup.JsApiShowWebDebugger)
		needUpdate = true;
	if (!ser || !ser.JSPluginsServer || ser.JSPluginsServer != "http://127.0.0.1:3888/jsplugins.xml")
		needUpdate = true;
	if (!sup) {
		sup = {}
		config.Support = sup
	}
	if (!ser) {
		ser = {}
		config.Server = ser
	}
	sup.JsApiPlugin = true
	sup.JsApiShowWebDebugger = true
	ser.JSPluginsServer = "http://127.0.0.1:3888/jsplugins.xml"
	if (needUpdate) {
		fs.writeFileSync(oemPath, ini.stringify(config))
		if (os.platform() != 'win32')
			cp.exec("quickstartoffice restart");
	}

	callback({ status: 0, msg: "wps安装正常，" + oemPath + "文件设置正常。" })
}
//检测WPS客户端的安装情况
function configOem(callback) {
	let oemPath;
	try {
		if (os.platform() == 'win32') {
			cp.exec("REG QUERY HKEY_CLASSES_ROOT\\KWPS.Document.12\\shell\\open\\command /ve", function (error, stdout, stderr) {
				try {
					var val = stdout.split("    ")[3].split('"')[1];
					if (typeof (val) == "undefined" || val == null) {
						return callback({
							status: 1,
							msg: "WPS未安装。"
						})
					}
					fs.exists(val, function (exists) {
						if(!exists){
							return callback({
								status: 1,
								msg: "WSP安装异常，请确认有没有正确的安装WPS2019。"
							})
						}
						oemPath = path.dirname(val) + '\\cfgs\\oem.ini';
						configOemFileInner(oemPath, callback);
					});					
				} catch (e) {
					oemResult = "配置" + oemPath + "失败，请尝试以管理员重新运行！！";
					console.log(oemResult)
					console.log(e)
					return callback({ status: 1, msg: oemResult })
				}
			});
		} else {
			oemPath = "/opt/kingsoft/wps-office/office6/cfgs/oem.ini";
			configOemFileInner(oemPath, callback);
		}
	} catch (e) {
		oemResult = "配置" + oemPath + "失败，请尝试以管理员重新运行！！";
		console.log(oemResult)
		console.log(e)
		return callback({ status: 1, msg: oemResult })
	}
}
//获取服务端IP地址
function getServerIPAdress() {
	var interfaces = require('os').networkInterfaces();
	for (var devName in interfaces) {
		var iface = interfaces[devName];
		for (var i = 0; i < iface.length; i++) {
			var alias = iface[i];
			if (alias.family === 'IPv4' && alias.address !== '127.0.0.1' && !alias.internal) {
				return alias.address;
			}
		}
	}
}

//----模拟服务端的特有功能，开发者无需关心 End--------