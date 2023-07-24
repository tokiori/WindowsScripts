/*!
 * WindowsJScript
 * https://github.com/tokiori/WindowsScripts
 *
 * Copyright (c) 2022 tokiori
 * This software is Released under the MIT license.
 * see https://opensource.org/licenses/MIT
 *
 */

//----------------------------------------------
// WindowsJScript
//----------------------------------------------
var WindowsJScript = function(){ return this; };
WindowsJScript.prototype = {

	multiple : false,
	args : [],
	info : null,
	init : function(){
		try{
			this.initwsf();
		} catch(e) {
			this.inithta();
		}
		js.log.init();
	},
	inithta : function(){
		// HTA warn avoid.
		WScript = null;
		this.info = this.path.info(location.pathname);
		this.args = this.hta.href2arg(location.href);
	},
	initwsf : function(){
		this.info = this.path.info(WScript.ScriptFullName);
		for(var i=0; i<WScript.Arguments.length; i++){
			this.args.push(WScript.Arguments.Item(i));
		}
		if(!this.multiple && this.cmd.executed()){
			this.quit("duplicate execute.");
		}
	},
	quit : function(txt){
		var msg = [];
		msg.push("[js.quit]");
		if(this.isstring(txt)) msg.push("[cause] "+txt);
		msg.push(this.info.fullpath);
		this.echo(msg.join("\r\n"));
		this.log.add(msg.join("\r\n"));
		try{
			WScript.Quit();
		}catch(e){
		    window.close();
		    throw new Error('[js.quit]' + e);
		}
	},
	echo : function(arg){
		try{
			WScript.Echo(arg);
		}catch(e){
			alert(arg);
		}
	},

	isundefined   : function(argv){ return argv === undefined; },
	isnull        : function(argv){ return argv === null; },
	isnumber      : function(argv){ return typeof argv === 'number'; },
	isint         : function(argv){ return Number.isInteger(argv); },
	isfloat       : function(argv){ return Number.isFinite(argv); },
	isstring      : function(argv){ return typeof argv === 'string'; },
	isboolean     : function(argv){ return typeof argv === 'boolean'; },
	issymbol      : function(argv){ return typeof argv === 'symbol'; },
	isarray       : function(argv){ return Object.prototype.toString.call(argv) === '[object Array]'; },
	isobject      : function(argv){ return typeof argv === 'object' },
//	isobject      : function(argv){ return typeof argv === 'object' && !this.isnull(argv); },
	isnullorempty : function(argv){
		if(this.isnull(argv)) return true;
		if(this.isundefined(argv)) return true;
		if(this.isstring(argv) && argv.length == 0) return true;
		var size = this.size(argv);
		if(!this.isboolean(size) && size == 0) return true;
		return false;
	},

	istrue  : function(val){ return (this.isboolean(val) && !!val); },
	isfalse : function(val){ return (this.isboolean(val) && !val); },

	size : function(obj){
		if(!this.isobject(obj)) return false;
		return this.each(obj, function(){ return true; }).length;
	},

	sort : function(obj, opt){
		// binary tree sort.

		var defaultoptions = {
			order   : 1,     // order by [asc : 1, desc : -1]
			withval : false, // order by [key : false, value : true]
			func : null
		};
		if(this.isobject(opt.order)) opt = this.extend(defaultoptions, opt);
		opt.order = (opt.order == -1) ? -1 : 1;
		opt.withval = (!opt.withval) ? false : true;

		if(!obj) return obj;

		var len = this.size(obj);
		var keyarr = [];
		var valarr = [];
		var comparearr = (opt.withval) ? valarr : keyarr;
		for(var key in obj){
			var cursor = Math.floor(comparearr.length / 2);
			var maxloop = this.math.log2(comparearr.length);
			for(var pivot = cursor, loop = 0; cursor < comparearr.length; loop++){
				// [warn] end when the maximum number of loops is exceeded.
				if(loop > maxloop){
					this.echo(this.str.format("[js.sort][warn] exceed maxloop."));
					break;
				}
				var strcurr = String(comparearr[cursor]);
				var strprev = cursor > 0 ? String(comparearr[cursor - 1]) : strcurr;
				var strkey = String(key);
				if(strprev < strkey && strcurr >= strkey){
					break;
				}
				if(pivot > 1) pivot = Math.floor(pivot / 2);
				if(pivot < 1) pivot = 1;
				cursor = cursor + (strcurr < strkey) ? pivot : - pivot;
			}
			keyarr.splice(cursor, 0, key);
			valarr.splice(cursor, 0, obj[key]);
		}
		var ret = {};
		
		for(var cnt = 0; cnt < keyarr.length; cnt++){
			var curr = (opt.order == -1) ? keyarr.length - (cnt + 1) : cnt;
			var key = keyarr[curr];
			var val = valarr[curr];
			ret[key] = (this.isnull(opt.func)) ? val : opt.func(key, val);
		}
		return ret;
	},

	max : function(arr){ return arr.reduce(function(a, b){ return Math.max(a, b) }); },
	min : function(arr){ return arr.reduce(function(a, b){ return Math.min(a, b) }); },

	longestkey   : function(obj){ return this.longest(obj, false); },
	longestvalue : function(obj){ return this.longest(obj, true); },
	longest : function(obj, targetIsValue){
		var self = this;
		var str = "";
		this.each(obj, function(key, val){
			var target = !!targetIsValue ? val : key;
			if(self.isnullorempty(target) || self.isobject(target)) return;
			if(String(target).length > str.length) str = String(target);
		});
		return str;
	},

	// like a jQuery.extend() function
	extend : function(){
		var args = [].slice.call(arguments);
		var rtn = args.shift();
		if(!this.isobject(rtn)) rtn = {};
		var self = this;
		this.each(args, function(i, obj){
			if(!self.isobject(obj)) return;
			self.each(obj, function(key, item){
				rtn[key] = item;
			});
		});
		return rtn;
	},

	// like a jQuery.each() function
	each : function(obj, func, ownprop){
		var rtn = [];
		for(var key in obj){
			if(!ownprop && !obj.hasOwnProperty(key)){
				continue;
			}
			rtn.push(func(key, obj[key]));
		}
		return rtn;
	}
};
var js = new WindowsJScript();

//----------------------------------------------
// WindowsJScript.hta
//----------------------------------------------
var WrapHtaControl = function(){};
WrapHtaControl.prototype = {
//	sep : "|||",
	sep : "&",
	args : {},
	arg2str : function(arg){
		if(arg.length == 1 && js.isobject(arg[0])) arg = arg[0];
		var strarr = js.each(arg, function(i, val){
			return js.str.format("arg${0}=${1}",i, val);
		});
		return (strarr.length > 0) ? "?" + strarr.join(this.sep) : "";
	},
	href2arg : function(href){
		if(!href.match(/.+\?.*/)){
			return [];
		}
		return this.str2arg(href.replace(/^.*\?/, ""));
	},
	str2arg : function(argstr){
		var argarr = argstr.split(this.sep);
		var arr = [];
		for(var i = 0; i < argarr.length; i++){
			var param = argarr[i].split("=");
			var val = js.isundefined(param[1]) ? "" : param[1];
			this.args[param[0]] = val;
			arr.push(val);
		}
		return arr;
	},
	escape : function(str){
		return str
			.replace(/&/g,  "&amp;")
			.replace(/"/g,  "&quot;")
			.replace(/'/g,  "&#039;")
			.replace(/</g,  "&lt;")
			.replace(/>/g,  "&gt;")
			.replace(/\n/g, "<br>");
	},
	exec : function(){
		var arg = [].slice.call(arguments);
		if(arg.length == 0){
			js.quit("[js.cmd.hta] empty arguments.");
			return false;
		};
		var htapath = js.path.info(arg.shift()).fullpath;
		var argstr = this.arg2str(arg);
		if(!js.path.isfile(htapath)){
			js.quit(["[js.cmd.hta] hta file not found.", htapath].join("\r\n"));
			return false;
		}
		var cmd = js.str.format('mshta.exe "${0}${1}"', htapath, argstr);
		return js.cmd.exec(cmd);
//		return js.cmd.shell().run(cmd);
	}
};
WindowsJScript.prototype.hta = new WrapHtaControl();

//----------------------------------------------
// WindowsJScript.cmd
//----------------------------------------------
var WrapCommand = function(){};
WrapCommand.prototype = {
	// prettier-ignore
	HASHTYPE : {
		MD5    : "MD5",
		SHA1   : "SHA1",
		SHA256 : "SHA256"
	},
	shell : function(name){
		try{
			return WScript.CreateObject('WScript.Shell');
		}catch(e){
			return new ActiveXObject('WScript.Shell');
		}
	},
	exec : function(cmd){
		var sh = this.shell();
		return sh.Exec(cmd);
	},
	hash : function(path, type){
		if(!this.HASHTYPE[type]){
			js.echo(js.msg.get("NotExistKey", type));
			return false;
		}
		if(!js.path.isfile(path)){
			js.echo(js.msg.get("NotExistFile", path));
			return false;
		}
		var cmd = js.str.format('certutil -hashfile "${0}" ${1}', path, type);
		return js.cmd.exec(cmd);
	},
	createshortcut : function(frompath, topath, opt){
		var sh = this.shell();
		var file = sh.CreateShortcut(topath);
		var prefix = String(frompath).match(/\.lnk$/) ? "file:/" : "";
		file.TargetPath = prefix + String(frompath);
		file.Save();
	},
	executed : function(){
		var windir = this.env("SystemRoot", "Process");
		var wmiObj = GetObject("WinMgmts:Root\\Cimv2");
		var processes = wmiObj.ExecQuery("Select * From Win32_Process");
		var penum = new Enumerator(processes);
		var cnt = 0;
		var cmdnm = js.path.join(windir, "system32", "wscript.exe").toString();
		for(; !penum.atEnd(); penum.moveNext()){
			var cmdln = penum.item().CommandLine;
			if(cmdln != null && cmdln.toLowerCase().indexOf(cmdnm.toLowerCase()) == 1 && cmdln.indexOf(js.info.filename) > 0 ){
				cnt++;
			}
		}
		return cnt > 1;
	},
	env : function(name, env){
		if(js.isundefined(env)){
			return this.shell().ExpandEnvironmentStrings(name);
		}
		var sh = this.shell().Environment("Process");
		return sh(name);
	}
};
WindowsJScript.prototype.cmd = new WrapCommand();

//----------------------------------------------
// WindowsJScript.dialog
//----------------------------------------------
var WrapDialog = function(){ return this; };
WrapDialog.prototype = {
	// prettier-ignore
	BUTTON : {
		OK                : 0,
		OK_CANCEL         : 1,
		STOP_RETRY_IGNORE : 2,
		YES_NO_CANCEL     : 3,
		YES_NO            : 4,
		RETRY_CANCEL      : 5
	},
	// prettier-ignore
	ICON : {
		NONE        :  0,
		STOP        : 16,
		QUESTION    : 32,
		EXCLAMAITON : 48,
		INFO        : 64
	},
	// prettier-ignore
	CHOOSE : {
		OK     :  1,
		CANCEL :  2,
		STOP   :  3,
		RETRY  :  4,
		IGNORE :  5,
		YES    :  6,
		NO     :  7,
		NONE   : -1
	},
	popup : function(opt){
		if(!js.isobject(opt)) return false;
		if(js.isnullorempty(opt.txt))   opt.txt   = "js.dialog text";
		if(js.isnullorempty(opt.sec))   opt.sec   = 0;
		if(js.isnullorempty(opt.title)) opt.title = "js.dialog title";
		if(js.isnullorempty(opt.btn))   opt.btn   = this.BUTTON.OK;
		if(js.isnullorempty(opt.icon))  opt.icon  = this.ICON.NONE;
		var sh = new ActiveXObject( "WScript.Shell" );
		return sh.Popup(opt.txt, opt.sec, opt.title, opt.btn + opt.icon);
	},
	savepath : function(path){
		var sfs = new ActiveXObject("SAFRCFileDlg.FileSave");
		sfs.FileName = path;
		sfs.FileType = "テキストファイル";
		var ret = fsv.OpenFileSaveDlg();
		sfs = null;
		return ret;
	}
};
WindowsJScript.prototype.dialog = new WrapDialog();

//----------------------------------------------
// WindowsJScript.msg
//----------------------------------------------
var WrapMsg = function(){ return this; };
WrapMsg.prototype = {
	msg : {
		NotExistKey  : "key is not exist.(${0})",
		NotExistFile : "file is not exist.(${0})",
		NotExistDir  : "dir is not exist.(${0})",
		NotExistMsg  : "messageid is not exist.(${0})"
	},
	// get(id, str...)
	get : function(){
		var arr = [].slice.call(arguments);
		var id = arr.shift();
		if(js.isundefined(this.msg[id])){
			return js.str.format(this.msg.NotExistMsg, id);
		}
		return js.str.format(this.msg[id], arr);
	}
};
WindowsJScript.prototype.msg = new WrapMsg();

//----------------------------------------------
// WindowsJScript.dbg
//----------------------------------------------
var WrapDebug = function(){ return this; };
WrapDebug.prototype = {
	obj2str : function(obj, level){
		if(!js.isobject(obj) && !js.isarray(obj)) return js.str.format('"${0}"', obj);
		if(!js.isnumber(level)) level = 0;
		var self = this;
		var rtn = [];
		var waku = js.isarray(obj) ? ["[", "]"] : ["{", "}"];
		var prefix = "";
		for(var i=0;i<level * 2;i++){
			prefix = prefix + " ";
		}
		rtn.push(waku[0]);
		js.each(obj, function(key, val){
			var txt = js.str.format("\"${0}\" : ${1}", key, self.obj2str(val, level + 1));
			rtn.push(prefix + "  " + txt);
		})
		rtn.push(prefix + waku[1]);
		return rtn.join("\r\n");
	},
	arr : [],
	push : function(str){
		this.arr.push(str);
	},
	echo : function(){
		js.echo(this.arr.join("\r\n"));
		this.arr = [];
	}
};
WindowsJScript.prototype.dbg = new WrapDebug();

//----------------------------------------------
// WindowsJScript.date
//----------------------------------------------
var WrapDate = function(){ return this; };
WrapDate.prototype = {
	// prettier-ignore
	fmt : {
		aaa  : function(dt){ return ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][dt.getDay()]; },
		YYYY : function(dt){ return String(dt.getFullYear()).substr(-4, 4); },
		MM   : function(dt){ return String("0" + (dt.getMonth() + 1)).slice(-2); },
		DD   : function(dt){ return String("0" +  dt.getDate()).slice(-2); },
		hh   : function(dt){ return String("0" +  dt.getHours()).slice(-2); },
		mm   : function(dt){ return String("0" +  dt.getMinutes()).slice(-2); },
		ss   : function(dt){ return String("0" +  dt.getSeconds()).slice(-2); },
		ms   : function(dt){ return String("00" +  dt.getTime()).slice(-3); }
	},
	now : function(format){
		return this.format(new Date, format);
	},
	format : function(dt, format){
		if(!format || format.length == 0){
			return String(dt);
		}
		var fmt = js.each(this.fmt, function(key, val){ return key; });
		var func = this.fmt;
		var regex = new RegExp("(" + fmt.join("|") + ")", "g");
		return format.replace(regex, function(match,str1){
			return (func[str1])? func[str1](dt) : match;
		});
	}
};
WindowsJScript.prototype.date = new WrapDate();

//----------------------------------------------
// WindowsJScript.math
//----------------------------------------------
var WrapMath = function(){ return this; };
WrapMath.prototype = {
	log2 : function(num, cnt){
		if(!cnt) cnt = 0;
		var val = Math.floor(num / 2);
		return (val > 0) ? this.log2(val, cnt + 1) : cnt;
	}
};
WindowsJScript.prototype.math = new WrapMath();

//----------------------------------------------
// WindowsJScript.str
//----------------------------------------------
var WrapString = function(){ return this; };
WrapString.prototype = {
	format : function(){
		var args = [].slice.call(arguments);
		var rtn = args.shift();
		if(js.isobject(args[0])) args = args[0];
		return rtn.replace(/\$\{(\d+)\}/g, function(match, i){
			return (!js.isundefined(args[i])) ? String(args[i]) : match;
		});
	},
	mapping : function(txt, map){
		return txt.replace(/\$\{([^\}]+)\}/g, function(match, str){
			return (!js.isundefined(map[str])) ? String(map[str]) : match;
		});
	},
	fill : function(value, length, filler, alignright){
		if(js.isnullorempty(filler)) filler = " ";
		if(!alignright) alignright = false;
		if(js.isnullorempty(value)) return value;
		value = String(value);
		var len = value.length;
		if(len >= length) return value;
		var str = "";
		for(var loop = length - len; loop > 0; loop--){
			str += filler;
		}
		if(alignright){
			return str + value;
		} else {
			return value + str;
		}
	}
};
WindowsJScript.prototype.str = new WrapString();

//----------------------------------------------
// WindowsJScript.path
//----------------------------------------------
var WrapPath = function(){ return this; };
WrapPath.prototype = {
	// prettier-ignore
	TYPEIS : {
		DIR  : 1,
		FILE : 2,
		ALL  : 9
	},
	newfso : function(){
		return new ActiveXObject("Scripting.FileSystemObject");
	},
	wrapfso : function(func){
		var fso = this.newfso();
		var ret = func(fso);
		fso = null;
		return ret;
	},
	readfile   : function(path, opt){ return js.file.read(path, opt); },
	writefile  : function(path, txt, opt){ return js.file.write(path, txt, opt); },
	appendfile : function(path, txt, opt){ return js.file.append(path, txt, opt); },
	info : function(path){
		var fso = this.newfso();
		var rtn = {
			fullpath : fso.GetAbsolutePathName(path),
			filename : fso.GetFileName(path),
			basename : fso.GetBaseName(path),
			extname  : fso.GetExtensionName(path),
			parent   : this.parent(path),
			isdir    : this.isdir(path),
			isfile   : this.isfile(path)
		};
		if(rtn.isfile){
			var obj = fso.GetFile(path);
			rtn.attr = obj.Attributes;
			rtn.createdate = obj.DateCreated;
			rtn.updatedate = obj.DateLastModified;
			rtn.drive = obj.Drive;
			rtn.name = obj.Name;
			rtn.parent = obj.ParentFolder;
			rtn.path = obj.Path;
			rtn.type = obj.Type;
			rtn.size = obj.Size;
		} else if(rtn.isdir){
			var obj = fso.GetFolder(path);
			rtn.createdate = obj.DateCreated;
			rtn.updatedate = obj.DateLastModified;
			rtn.drive = obj.Drive;
			rtn.name = obj.Name;
			rtn.parent = obj.ParentFolder;
			rtn.path = obj.Path;
			rtn.type = obj.Type;
//			rtn.size = obj.Size;
			rtn.isroot = obj.IsRootFolder;
		}
		fso = null;
		return rtn;
	},
	each : function(path, func, opt) {
		var self = this;
		var rtn = [];
		if(js.isundefined(opt)) opt = {};
		if(js.isundefined(opt.curr)) opt.curr = 0;
		if(js.isundefined(opt.nest)) opt.nest = null;
		if(js.isundefined(opt.dir))  opt.dir  = false;
		if(js.isundefined(opt.file)) opt.file = true;

		if(this.isfile(path)){
			return [func(path)];
		} else if(this.isdir(path)){
			if(opt.dir){
				rtn.push(func(path));
			}
			var fso = this.newfso();
			var dir = fso.GetFolder(path);
			if(opt.file){
				for (var fp = new Enumerator(dir.Files); !fp.atEnd(); fp.moveNext()) {
					if(opt.filterfile && !String(fp.item()).match(opt.filterfile)){
						continue;
					}
					rtn.push.apply(rtn, this.each(fp.item(), func, opt));
				}
			}
			if(js.isnumber(opt.nest) && opt.curr >= opt.nest) return rtn;
			opt.curr++;
			for (var fp = new Enumerator(dir.SubFolders); !fp.atEnd(); fp.moveNext()) {
				if(opt.filterdir && !String(fp.item()).match(opt.filterdir)){
					continue;
				}
				rtn.push.apply(rtn, this.each(fp.item(), func, opt));
			}
			fso = null;
		}
		return rtn;
	},
	extchange : function(path, ext){ return js.file.extchange(path, ext); },
	remove: function(path){ return this.isdir(path) ? this.rmdir(path) : this.isfile(path) ? this.rm(path) : false; },
	copy  : function(from, to){ return this.isdir(from) ? this.cpdir(from, to) : this.isfile(from) ? this.cp(from, to) : false; },
	move  : function(from, to){ return this.isdir(from) ? this.mvdir(from, to) : this.isfile(from) ? this.mv(from, to) : false; },
	rm    : function(path){ return this.wrapfso(function(fso){ return fso.DeleteFile(path); }); },
	rmdir : function(path){ return this.wrapfso(function(fso){ return fso.DeleteFolder(path); }); },
	cp    : function(from, to){ return this.wrapfso(function(fso){ return fso.CopyFile(from, to); }); },
	cpdir : function(from, to){ return this.wrapfso(function(fso){ return fso.CopyFolder(from, to); }); },
	mv    : function(from, to){ return this.wrapfso(function(fso){ return fso.MoveFile(from, to); }); },
	mvdir : function(from, to){ return this.wrapfso(function(fso){ return fso.MoveFolder(from, to); }); },
	touch : function(path, ismkdir){ return js.file.touch(path, ismkdir); },
	mkdir : function(path){
		if(this.isdir(path)) return true;
		this.mkdir(this.parent(path));
		this.wrapfso(function(fso){ return fso.CreateFolder(path); });
		return true;
	},
	sla2dq : function(str){
		return str.replace(/\//g, "\\");
	},
	join : function(){
		var self = this;
		var args = [].slice.call(arguments);
		if(args.length == 1 && js.isobject(args[0])){
			args = args[0];
		}
		var path = this.info(String(args.shift())).fullpath;
		js.each(args, function(i, val){
			path = self.wrapfso(function(fso){ return fso.BuildPath(path, self.sla2dq(val)); });
		});
		return path;
	},
	tmpname : function(){ return this.wrapfso(function(fso){ return fso.GetTempName(); }); },
	isdrive : function(path){ return this.wrapfso(function(fso){ return fso.DriveExists(path); }); },
	drive   : function(path){ return this.wrapfso(function(fso){ return fso.GetDriveName(path); }); },
	parent  : function(path){ return this.wrapfso(function(fso){ return fso.GetParentFolderName(path); }); },
	isfile  : function(path){ return this.wrapfso(function(fso){ return fso.FileExists(path); }); },
	isdir   : function(path){ return this.wrapfso(function(fso){ return fso.FolderExists(path); }); },
	exist   : function(path){ return this.isdir(path) || this.isfile(path); }
};
WindowsJScript.prototype.path = new WrapPath();

//----------------------------------------------
// WindowsJScript.file
//----------------------------------------------
var WrapFile = function(){ return this; };
WrapFile.prototype = {
	// prettier-ignore
	OPENMODE : {
		FORREAD   : 1,     // [default]
		FORWRITE  : 2,
		FORAPPEND : 8
	},
	// prettier-ignore
	NOTHINGTHEN : {
		CREATE :  true,
		THROUGH : false    // [default]
	},
	// prettier-ignore
	CHARCODE : {
		UTF    : -1,
		SJIS   : 0,        // [default]
		SYSTEM : -2
	},
	// prettier-ignore
	NEWLINECODE : {
		CR   : "\r",
		LF   : "\n",
		CRLF : "\r\n"
	},
	defaultoptions : function(){
		return {
			OPENMODE    : this.OPENMODE.FORREAD,
			NOTHINGTHEN : this.NOTHINGTHEN.CREATE,
			CHARCODE    : this.CHARCODE.SYSTEM,
			NEWLINE     : false,
			NEWLINECODE : null
		}
	},
	open : function(path, txt, opt){
		try {
			var info = js.path.info(path);
			if(!js.path.isdir(info.parent)){
			    throw new Error(js.msg.get("NotExistDir", path));
			}
			if(js.isnullorempty(info.filename)){
			    throw new Error("path failed.(" + path + ")");
			}

			opt = js.extend(this.defaultoptions(), opt);

			var stream = js.path.wrapfso(function(fso){ return fso.OpenTextFile(path, opt.OPENMODE, opt.NOTHINGTHEN, opt.CHARCODE); });

			if(opt.OPENMODE === this.OPENMODE.FORREAD){
				// テキスト初期化
				var text = "";

//				// 読み取り方法① ファイルから全ての文字データを読み込む
//				text = stream.ReadAll();
//				// ファイルの末尾までループ
//				while (!stream.AtEndOfStream) {
//					// 読み取り方法② ファイルの文字データを一行ずつ表示する
//					text += stream.ReadLine();
//					// 読み取り方法③ 読み込みバッファ指定でループする
//					text += stream.Read(1024);
//				}
//				// 読み取り方法④ 全行を読み込む
//				text = stream.ReadText(-1);
				// 読み取り方法⑤ サイズ分を一括で読み込む
				text = stream.Read(stream.Size);

				stream.Close();
				return text;
			} else {
				if(txt != null && !opt.NEWLINE){
					// 文字列を書き込む(改行なし)
					stream.Write(txt);
				}else if(txt != null){
					// 文字列を書き込む(改行あり)
					stream.WriteLine(txt);
				}
//				//  ブランク行をN行書き込む
//				stream.WriteBlankLines( 2 );

				stream.Close();
				return true;
			}
		} catch(e) {
			if(stream) stream.Close();
		    throw new Error('[js.file.open] ' + e);
		}
	},
	read : function(path, opt){
		return this.open(path, null, js.extend(opt, { OPENMODE : this.OPENMODE.FORREAD }));
	},
	write : function(path, txt, opt){
		return this.open(path, txt, js.extend(opt, { OPENMODE : this.OPENMODE.FORWRITE }));
	},
	append : function(path, txt, opt){
		return this.open(path, txt, js.extend(opt, { OPENMODE : this.OPENMODE.FORAPPEND }));
	},
	extchange : function(path, ext){
		if(!js.path.isfile(path)) return path;
		return path.replace(/\.[^\.]+$/, ext);
	},
	touch : function(path, ismkdir){
		if(!!ismkdir){ js.path.mkdir(js.path.parent(path)); }
		return this.open(path, null, this.OPENMODE.FORWRITE, this.NOTHINGTHEN.CREATE, this.CHARCODE.SYSTEM);
	}
};
WindowsJScript.prototype.file = new WrapFile();

//----------------------------------------------
// WindowsJScript.log
//----------------------------------------------
var WrapLog = function(){ return this; };
WrapLog.prototype = {
	state : {
		enable : false,
		ready  : false,
		info   : true,
		warn   : true,
		err    : true,
		dbg    : false
	},
	path : null,
	enabled : function(){
		return this.state.enable;
	},
	on  : function(){
		this.ready();
		if(this.state.enable) return true;
		this.state.enable = true;
		this.add("[js.log.on] start log.")
	},
	off : function(){
		this.add("[js.log.off] stop log.")
		this.state.enable = false;
	},
	dologging :function(path){
		var txt = [];
			txt.push("[js.log] log folder is not exist.");
			txt.push("${0}");
			txt.push("");
			txt.push("create folder or no logging.");
			txt.push("");
			txt.push("[YES] create folder.");
			txt.push("[NO] no logging.");
			txt.push("[CANCEL] quit.");
		var opt = {
			title: "[js.log] confirm execute.",
			txt  : js.str.format(txt.join("\r\n"), path),
			btn  : js.dialog.BUTTON.YES_NO_CANCEL,
			icon : js.dialog.ICON.INFO
		}
		var flg = js.dialog.popup(opt);
		if(flg == js.dialog.CHOOSE.YES){
			js.path.mkdir(path);
			return true;
		}
		if(flg == js.dialog.CHOOSE.CANCEL){
			js.quit("cancel selected. (log folder not found.)");
		}
		return false;
	},
	ready : function(){
		var logdir = js.path.parent(this.path);
		if(js.path.isdir(logdir)){
	 		this.state.ready = true;
		} else {
	 		this.state.ready = this.dologging(logdir);
		}
 		return this.state.ready;
	},
	init : function(state, path){
		var currpath = (this.path) ? this.path : js.path.extchange(js.info.fullpath, js.str.format("_${0}.log", js.date.now("YYYYMMDDhhmmss")));
 		this.path = (path) ? path : currpath;
 		if(js.isobject(state)){
			js.extend(this.state, state);
 		}
 		return this.ready();
	},
	add : function(txt, type){
		type = !js.isstring(type) || this.state[type] ? true : false;
		if(type && this.state.enable && this.state.ready){
			js.path.appendfile(this.path, txt, {NEWLINE : true});
			return true;
		}
		return false;
	},
	info : function(txt){ return this.add(js.date.now("[YYYY/MM/DD hh:mm:ss.ms][INFO]") + txt, "info"); },
	warn : function(txt){ return this.add(js.date.now("[YYYY/MM/DD hh:mm:ss.ms][WARN]") + txt, "warn"); },
	err  : function(txt){ return this.add(js.date.now("[YYYY/MM/DD hh:mm:ss.ms][ERROR]") + txt, "err"); },
	dbg  : function(txt){ return this.add(js.date.now("[YYYY/MM/DD hh:mm:ss.ms][DEBUG]") + txt, "dbg"); }
};
WindowsJScript.prototype.log = new WrapLog();
WindowsJScript.prototype.Logger = WrapLog;

//----------------------------------------------
// WindowsJScript.book
//----------------------------------------------
var WrapBook = function(){ return this; };
WrapBook.prototype = {
	extregex : /\.xls[xms]?$/,
	VIEWMODE : {
		"xlNormalView"       : 1, // 標準
		"xlPageBreakPreview" : 2, // 改ページプレビュー
		"xlPageLayoutView"   : 3  // ページレイアウト
	},
	maxrow : function(sheet){ return sheet.Cells(1, 1).SpecialCells(11).Row },
	maxcol : function(sheet){ return sheet.Cells(1, 1).SpecialCells(11).Column },
	R1C1 : function(sheet){},
	A1 : function(sheet, a1){
		var r1c1 = this.A1toR1C1(a1);
		return sheet.Cells(r1c1.row, r1c1.col);
	},
	A1toR1C1 : function(a1){
		var arr = [];
		a1.toUpperCase().replace(/\:?\$?([A-Z]+)\$?([0-9]+)/g, function(match, col, row){
			var colarr = col.split('');
			var x = 0;
			var AtoZ = 26;
			var len = colarr.length;
			for(var i = 0; i < len; i++){
				var ix = String(colarr[len -(i + 1)]).toUpperCase().charCodeAt() - String("A").charCodeAt() + 1;
				x += ix * Math.pow(AtoZ, i);
			};
			arr.push({
				col : x,
				row : row,
				a1  : match
			});
		});
		return arr;
	},
	newexcel : function(){
		return new ActiveXObject("Excel.Application");
	},
	wrapexcel : function(func){
		var excel = this.newexcel();
		var ret = func(excel);
		excel = null;
		return ret;
	},
	isbook : function(path){
		if(!js.path.isfile(path)){
			return false;
		}
		if(!String(path).match(this.extregex)){
			return false;
		}
		return true;
	},
	each : function(book, func, opt){
		if(js.isundefined(opt)) opt = {};
		var ret = [];
		for(var id = 1; id <= book.WorkSheets.Count; id++){
			var sheet = book.WorkSheets(id);
			if(!js.isundefined(opt.filter) && String(sheet.Name).match(opt.filter)){
				ret.push(func(id, sheet));
			}
		}
		return ret;
	},
	exeptionlog : function(name, path, e){
		var msg = js.str.format("[${3}] error(${1}:${2}). path:${0}", path, (e.number & 0xFFFF), e.message, name);
		js.echo(msg);
		js.log.err(msg);
	},
	// prettier-ignore
	defaultoptions : function(){
		return {
			NewBook            : false,
			ReadOnly           : true,
			Visible            : false, // ブック表示
			DisplayAlerts      : false, // 実行中アラート表示
			CheckCompatibility : false, // 互換性チェック実行
			LogPrefix          : "js.book.open"
		}
	},
	read   : function(path, func, opt){ return this.open(path, func, js.extend(opt, { NewBook : false, ReadOnly : true,  LogPrefix : "js.book.read"   })); },
	update : function(path, func, opt){ return this.open(path, func, js.extend(opt, { NewBook : false, ReadOnly : false, LogPrefix : "js.book.update" })); },
	create : function(path, func, opt){ return this.open(path, func, js.extend(opt, { NewBook : true,  ReadOnly : false, LogPrefix : "js.book.create" })); },
	open : function(path, func, opt){
		var excel = this.newexcel();
		excel.Visible = opt.Visible;
		var book = null;
		var rtn = false;
		try{
			opt = js.extend(this.defaultoptions, opt);
			if(opt.NewBook){
				book = excel.Workbooks.Add();
			} else {
				book = excel.Workbooks.Open(path);
				if(!this.isbook(path)) return false;
			}
			if(!func(book, excel)) return false;
			if(opt.ReadOnly) return true;
			book.CheckCompatibility = opt.CheckCompatibility;
			excel.DisplayAlerts = opt.DisplayAlerts;
			book.SaveAs(path);
			rtn = true;
		} catch(e) {
			this.exeptionlog(opt.LogPrefix, path, e);
			rtn = false;
		} finally {
			if(!js.isnull(book)) {
				excel.CutCopyMode = false;
				book.Close(false);
				book = null;
			}
			excel = null;
			return rtn;
		}
	}
};
WindowsJScript.prototype.book = new WrapBook();

//----------------------------------------------
// WindowsJScript.json (Experimental implementation)
//----------------------------------------------
var WrapJson = function () { return this; };
WrapJson.prototype = {
    props: {
        indentStr: " ",
        indentLen: 4,
        newLineStr: "\r\n"
    },
    loopStr: function (str, loop) {
        var ret = [];
        for (var i = 0; i < loop; i++) {
            ret.push(str);
        }
        return ret.join("");
    },
    getIndent: function (nest) {
        return this.loopStr(this.props.indentStr, nest * this.props.indentLen);
    },
    anyEqual: function (str) {
        var args = [].slice.call(arguments);
        args.shift();
        for (var i = 0; i < args.length; i++) {
            if (str === args[i]) return true;
        }
        return false;
    },
    sharping: function (str) {
        str = str.replace(/(\n|\r)/g, "");
        var len = str.length;
        var ret = [];
        var nest = 0;
        var isInnerValue = false;
        for (var i = 0; i < len; i++) {
            var curr = str.charAt(i);
            var prev = str.charAt(i - 1);
            if (!isInnerValue && this.anyEqual(curr, " ", "\t")) {
                continue;
            } else if (!isInnerValue && curr === ",") {
                ret.push(curr + this.props.newLineStr + this.getIndent(nest));
            } else if (!isInnerValue && this.anyEqual(curr, "{", "[")) {
                nest++;
                ret.push(curr + this.props.newLineStr + this.getIndent(nest));
            } else if (!isInnerValue && this.anyEqual(curr, "}", "]")) {
                nest--;
                ret.push(this.props.newLineStr + this.getIndent(nest) + curr);
            } else if (!isInnerValue && curr === ":") {
                ret.push(curr + " ");
            } else if (curr === '"' && prev !== "\\") {
                isInnerValue = !isInnerValue;
                ret.push(curr);
            } else {
                ret.push(curr);
            }
        }
        return ret.join("");
    },
    filter: function (str, key, condition) {
        if (!str || str.replace(/ +/, "").length == 0) return "";
        return this.filterString(str, key, condition);
        // return this.filterObject(str, key);
    },
    filterObject: function (str, key) {
        if (!str || String(str).length == 0) return "";
        var jsonObj = JSON.parse(str);
        var findRegex = new RegExp(key, "i");
        var res = this.filterObjectRecursive(jsonObj, findRegex);
        return this.sharping(JSON.stringify(res.data));
    },
    filterObjectRecursive: function (value, keyword) {
        if (Object.prototype.toString.call(value) === "[object Array]") {
            var arr = [];
            for (var i in value) {
                var res = this.filterObjectRecursive(value[i], keyword);
                if (res.isMatch) arr.push(res.data);
            }
            return { isMatch: arr.length, data: arr };
        } else if (Object.prototype.toString.call(value) === "[object Object]") {
            var obj = {};
            for (var i in value) {
                var res = this.filterObjectRecursive(value[i], keyword);
                // contains keyword in value then append item
                if (res.isMatch) obj[i] = res.data;
                // contains keyword in key then append item
                if (String(i).match(keyword)) obj[i] = res.data;
            }
            return { isMatch: Object.keys(obj).length, data: obj };
        }
        return { isMatch: String(value).match(keyword), data: value };
    },
    filterString: function (str, key, condition) {
        var rows = this.sharping(str).split(this.props.newLineStr);
        var res = [];
        var printed = [];
        var current = [];
        var findRegex = new RegExp(key, "i");

        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var currNest = this.getNestLevel(row);

            if (current.length == currNest + 1) {
                current[currNest] = row;
            } else if (current.length < currNest + 1) {
                current.push(row);
                current[currNest] = row;
            } else if (current.length > currNest) {
                current.pop();
                current[currNest] = row;
            }
            if (printed.length > currNest) {
                this.setCloseNest(res, printed);
            }
            if (condition.looseMatcher) {
                row = row.replace(/^ +/, "").replace(/\"/g, "");
            }
            if (row.match(findRegex)) {
                this.setParentNest(res, printed, current);
            }
        }
        return this.trimLastChildComma(res).join(this.props.newLineStr);
    },
    setParentNest: function (res, printed, current) {
        for (var pushNest = 0; pushNest < current.length; pushNest++) {
            if (printed.length < pushNest || printed[pushNest] != current[pushNest]) {
                res.push(current[pushNest]);
                printed[pushNest] = current[pushNest];
            }
        }
    },
    setCloseNest: function (res, printed) {
        var poped = printed.pop();
        if (poped.match(/\[ *$/)) {
            res.push(this.getIndent(printed.length) + "],");
        } else if (poped.match(/\{ *$/)) {
            res.push(this.getIndent(printed.length) + "},");
        }
    },
    getNestLevel: function (row) {
        var indentRegex = new RegExp("^" + this.props.indentStr + "+");
        var matchIndent = row.match(indentRegex);
        return matchIndent ? matchIndent[0].length / this.props.indentLen : 0;
    },
    trimLastChildComma: function (rows) {
        if (rows.length == 0) return rows;
        var res = [];
        for (var i = 0; i < rows.length - 1; i++) {
            var curr = this.getNestLevel(rows[i]);
            var next = this.getNestLevel(rows[i + 1]);
            var row = curr != next ? rows[i].replace(/,? *$/, "") : rows[i];
            res.push(row);
        }
        res.push(rows[rows.length - 1].replace(/,? *$/, ""));
        return res;
    }
};
WindowsJScript.prototype.json = new WrapJson();

js.init();
