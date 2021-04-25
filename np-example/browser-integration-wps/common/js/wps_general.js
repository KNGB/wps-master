/** 
 * @file common.js
 * @desc WPS一些通用的方法，各个平台通用
 * @author wps-developer
 */

/**
 * @desc WPS公用函数类的构造函数
 */
function WPSFrame() {
	this.elementObj = null;
	// 插件对象，包含一些封装的方法，以及电子公文标准接口
	this.pluginObj = null;
	// WPS的Application对象，调用WPS内部API入口点，和MicroSoft Office API通用
	this.Application = null;
	// 浏览器信息
	this.browserInfo = getBrowserInfo();
}

WPSFrame.prototype = {
	constructor: WPSFrame,
	/**
	 * @desc 获取插件元素对象
	 * @param { String } tagID 插件父布局div元素id
	 * @returns object元素对象，即WPS插件对象
	 */
	getElementObj: function (tagID) {
		var iframe;
		var codes = [];

		iframe = document.getElementById(tagID);
		if (isWindowsPlatform()) {
			codes.push('<object height="100%" width="100%" ');
			codes.push('classid=clsid:8E7DA7EC-07EC-4343-8141-88A2ADB63A5F></object> ');
			iframe.innerHTML = codes.join("");
			this.elementObj = iframe.children[0];
		} else {
			if (iframe.innerHTML.indexOf("application/x-wps") > -1) {
				iframe.innerHTML = "";
			}
			codes.push("<object type='application/x-wps'  width='100%'  height='100%'></object>");
			iframe.innerHTML = codes.join("");
			this.elementObj = iframe.children[0];
			window.onbeforeunload = function () {
				this.elementObj.Application.Quit();
			};
		}

		return this.elementObj;
	},

	/**
	 * @param { String } id 插件父布局div元素id
	 * @param { Function } callback 初始化完成之后的回调函数
	 * @desc 初始化WPS，并初始化插件对象和WPS的Application对象
	 */
	initPlugin: function (id, callback) {
		this.pluginObj = this.getElementObj(id);
		if (!isWindowsPlatform()) {
			if (undefined != this.browserInfo.chrome) {
				this.pluginObj = this.elementObj.Application;
				if (null == this.pluginObj) {
					this.elementObj.setAttribute('data', '../../../file/Normal.dotm');
					var Interval_control = setInterval(
						function () {
							this.pluginObj = this.elementObj.Application;
							if (this.pluginObj && this.pluginObj.IsLoad()) {
								clearInterval(Interval_control);
								this.Application = this.pluginObj;
								if (undefined != callback) {
									callback();
								}
							}
						}, 500);
				} else {
					this.Application = this.pluginObj;
				}
			} else {
				this.pluginObj = this.elementObj.Application;
				this.Application = this.pluginObj;
			}
		} else {
			this.pluginObj = this.elementObj;
			// WPS三个项目分别对应wps,et,wpp
			this.pluginObj.AppModeType = "wps";
			// this.Application = this.pluginObj.Application;
		}

		if (undefined != callback) {
			callback();
		}
	},

	/**
	 * @desc 将当前文档保存到指定路径的服务器上
	 * 		 注意: windows平台采用的上传不是表单上传，后台服务器需要参照upload_w.jsp来解析文件数据
	 * 			  linux平台则采用标准的表单上传，后台服务器可以采用标准的表单处理逻辑
	 * @param { String } uploadURL 保存到服务器的路径
	 * @param { String } fileName 保存服务器指定文档名
	 * @returns { Bool } Windows平台为true或false
	 * 			{ String } Linux平台为json类型字符串
	 */
	saveFileToSever: function (uploadURL, fileName) {
		var ret;

		if (undefined == fileName) {
			fileName = 'defaultSave.wps';
		}

		if (isWindowsPlatform()) {
			ret = this.pluginObj.saveURL(uploadURL, fileName);
		} else {
			var paramInfo = '{"fileName" :' + fileName + '}';
			ret = this.SaveDocumentToServerForLinux(uploadURL, paramInfo);
		}

		return ret;
	},

	/**
	 * @desc 打开本地指定路径的文档
	 * @param { String } fileName 要打开的文件名称
	 * @param { bool } readOnly 是否是只读打开文档，true只读，false可编辑
	 * @param { String } passWord 打开文档所需的密码
	 * @returns { Bool } 成功为true，失败为false
	 */
	openDocumentByLocal: function (filePath, readOnly, passWord) {
		var ret;

		if (undefined == readOnly) {
			readOnly = false;
		}

		if (undefined == passWord) {
			passWord = '';
		}

		ret = this.pluginObj.openDocument(filePath, readOnly, passWord);

		return ret;
	},

	/**
	 * @desc 打开服务器上指定路径的文档
	 * @param { String } fileName 要打开的文件名称
	 * @param { bool } readOnly 是否是只读打开文档，true只读，false可编辑
	 * @returns { Bool } Windows平台为true或false
	 * 			{ String } Linux平台为json类型字符串
	 */
	openDocumentBySever: function (fileURL, readOnly) {
		var ret;

		if (undefined == readOnly) {
			readOnly = false;
		}

		if (isWindowsPlatform()) {
			ret = this.pluginObj.openDocumentRemote(fileURL, readOnly);
		} else {
			var paramInfo = '{"readOnly" :' + readOnly + '}';
			ret = this.openDocumentFromSeverForLinux(fileURL, paramInfo);
		}

		return ret;
	},

	/**
	 * @desc 打开服务器指定文档，
	 * 		 支持http和https（注意：由于浏览器限制，不支持携带当前页面的sessionid以及httponly类型的cookie，
	 * 		 但可以携带其他的非httponly类型的cookie），请求头会带上根据文件大小计算的md5值
	 * @param { String } URL 服务器文件的URL或者返回文件数据的URL
	 * @param { String } paramInfo json格式字符串，默认可以不传该参数，json每一个键值对都是可选参数，不需要可以不传，若传值则格式如下
	 * 								{
	 *									"isNoTmpFile" : false 			{ Bool } 是否不落地保存，本地不产生临时文件为true，本地产生临时文件为false，默认值为false
	 *									"readOnly" : false				{ Bool } 是否只读打开文档，只读为true，可编辑为false，默认值为false
	 *									"password" : "123" 				{ String } 打开文档所需的密码，默认为空
	 *									"isGetResponseHead" : false		{ Bool } 是否返回响应头信息
	 *								}
	 *	@returns { String } json格式字符串，格式如下
	 *						{
	 * 							"result" : true					{ Bool } 打开服务器文档结果，成功为true，失败false
	 * 							"errMsg" : "错误信息"		 	 { String } 打开失败错误信息
	 * 							"fileName" : "打开文件名"		 { String } 要打开文件的文件名
	 * 							"responseHead" : {
	 *											"key1" : "value1",
	 * 											"key2" : "value2",
	 * 											"key3" : "value3",
	 * 												...
	 * 											 }				{ String } json类型字符串，返回响应头的所有信息
	 *						}
	 */
	openDocumentFromSeverForLinux: function (URL, paramInfo) {
		if (undefined == paramInfo || 0 == paramInfo.length) {
			paramInfo = "{}";
		}

		return this.pluginObj.OpenDocumentFromServer(URL, paramInfo);
	},

	/**
	 * @desc 保存当前文档到服务器，
	 * 		 支持http和https（注意：由于浏览器限制，不支持携带当前页面的sessionid以及httponly类型的cookie，
	 * 		 但可以携带其他的非httponly类型的cookie），请求头会带上根据文件大小计算的md5值，保存文档传输数据格式
	 * 		 采用标准的表单上传，后端服务器可采用表单上传流程进行解析数据
	 * @param { String } URL 服务器接收文件数据的接口URL地址
	 * @param { String } paramInfo json格式字符串，默认可以不传该参数，json每一个键值对都是可选参数，不需要可以不传，若传值则格式如下
	 * 						{
	 * 							"fileName" : "default.wps",		{ String } 上传文件的文件名
	 * 							"isGetBodyResult" : false,		{ Bool } 是否需要返回请求数据，如果需要返回http请求返回的内容，则值为true，否则为false，默认值false
	 * 							"isNoTmpFile" : false,			{ Bool } 是否不落地保存，本地不产生临时文件为true，本地产生临时文件为false，默认值为false
	 * 							"customFromData" : {
	 * 												"key1" : "value1"
	 * 												"key2" : "value2"
	 * 												"key3" : "value3"
	 * 												...
	 * 											},				{ String } json格式数组，用于表单上传的字段里面可自定义添加数据
	 * 							"isGetResponseHead" : false,	{ Bool } 是否返回响应头信息，true返回，false不返回
	 * 								
	 * 						}
	 * @returns { String } json格式字符串，格式如下
	 *						{
	 * 							"result" : true,					{ Bool } 保存服务器文档结果，成功为true，失败false
	 * 							"errMsg" : "错误信息",		 		 { String } 保存失败错误信息
	 * 							"responseBody" : "请求返回信息",	  { String } http或https请求返回信息
	 * 							"responseHead" : {
	 *											"key1" : "value1"
	 * 											"key2" : "value2"
	 * 											"key3" : "value3"
	 * 												...
	 * 											 }					{ String } json类型字符串，返回响应头的所有信息	
	 *						}
	 */
	SaveDocumentToServerForLinux: function (URL, paramInfo) {
		if (undefined == paramInfo || 0 == paramInfo.length) {
			paramInfo = "{}";
		}

		return this.pluginObj.SaveDocumentToServer(URL, paramInfo);
	}
}

/**
 * @desc 判断当前系统类型
 * @returns { bool } true代表windows系统，false代表linux系统
 */
function isWindowsPlatform() {
	var platform;

	platform = navigator.platform;
	if (platform == 'Win32' || platform == "Windows") {
		return true;
	} else {
		return false;
	}
}

/**
 * @desc 获取浏览器信息
 * @returns { object } 浏览器信息对象
 */
function getBrowserInfo() {
	var browserInfo = {},
		ua, s;

	ua = navigator.userAgent.toLowerCase();
	(s = ua.match(/msie ([\d.]+)/)) ? browserInfo.ie = s[1]:
		(s = ua.match(/firefox\/([\d.]+)/)) ? browserInfo.firefox = s[1] :
		(s = ua.match(/chrome\/([\d.]+)/)) ? browserInfo.chrome = s[1] :
		(s = ua.match(/opera.([\d.]+)/)) ? browserInfo.opera = s[1] :
		(s = ua.match(/version\/([\d.]+).*safari/)) ? browserInfo.safari = s[1] : 0;

	return browserInfo;
}

/**
 * @desc 检查当前环境是否满足wps插件运行条件
 */
function checkEnvironment() {
	var haveWPSPlugin = false;

	var browserInfo = getBrowserInfo();
	if (isWindowsPlatform()) {
		try {
			var wpsPlugin = new ActiveXObject('WPSDocFrame.Frame');
			haveWPSPlugin = true;
		} catch (error) {
			haveWPSPlugin = false;
		}
	} else {
		var plugins = navigator.plugins;
		for (var i = 0; i < plugins.length; ++i) {
			if (plugins[i].name.indexOf('Kingsoft WPS Plugin') > -1) {
				haveWPSPlugin = true;
			}
		}
	}

	if (!haveWPSPlugin) {
		if (undefined != browserInfo.chrome && ua.indexOf('qihu') > -1) {
			alert('当前Chrome内核浏览器不支持NPAPI协议，请更换其他支持该协议的浏览器！');
		} else if (undefined != browserInfo.ie && parseInt(browserInfo.ie) < 8) {
			alert('需要安装IE8或者IE8以上版本浏览器，低版本浏览器不支持ActiveX插件！');
		} else if (undefined != browserInfo.firefox && parseInt(browserInfo.firefox) > 52) {
			alert('当前FireFox版本不支持NPAPI协议，请更换52及以下版本再次尝试！');
		} else {
			alert('当前系统未安装WPS专业版，请安装之后重启浏览器再次尝试！');
		}
	}
}