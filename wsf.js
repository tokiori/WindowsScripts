/*!
 * WindowsJScript
 *
 * Copyright (c) 2022 tokiori
 * This software is Released under the MIT license.
 * see https://opensource.org/licenses/MIT
 *
 */
var WindowsJScript = function(){ this.init(); return this; };
WindowsJScript.prototype = {

	execfile : null,
	args : [],
	init : function(){
		this.execfile = WScript.ScriptFullName;
		for(var i=0; i<WScript.Arguments.length; i++){
			this.args.push(WScript.Arguments.Item(i));
		}
	},

	echo:function(arg){ WScript.Echo(arg); },

	isundefined   : function(argv){ return argv === undefined; },
	isnull        : function(argv){ return argv === null; },
	isnumber      : function(argv){ return typeof argv === 'number'; },
	isint         : function(argv){ return Number.isInteger(argv); },
	isfloat       : function(argv){ return Number.isFinite(argv); },
	isstring      : function(argv){ return typeof argv === 'string'; },
	isboolean     : function(argv){ return typeof argv === 'boolean'; },
	issymbol      : function(argv){ return typeof argv === 'symbol'; },
	isobject      : function(argv){ return typeof argv === 'object' && !this.isnull(argv); },
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

	executed : function(){
		var process = GetObject("WinMgmts:").ExecQuery("Select * From Win32_Process");
		var ret = false;

		for(var e = new Enumerator(process); !e.atEnd(); e.moveNext()){
			if(e.item().ExecutablePath === WScript.FullName && String(e.item().CommandLine).indexOf(WScript.ScriptFullName) > -1){
				ret = true;
				break;
			}
		}

		if(ret){
			WScript.Echo(WScript.ScriptName + " is already running")
			WScript.Quit()
		}
		return ret;
	},

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
				// 挿入位置が見つかったら終わる
				if(strprev < strkey && strcurr >= strkey){
					break;
				}
				// ピボットが1より大きいなら中間値を取得
				if(pivot > 1) pivot = Math.floor(pivot / 2);
				// ピボットが0になったら1（隣）を指定
				if(pivot < 1) pivot = 1;
				// カーソル位置を移動
				cursor = cursor + (strcurr < strkey) ? pivot : - pivot;
			}
			// カーソル位置に値を追加
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
		var self = this;
		this.each(args, function(i, obj){
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
// WindowsJScript.cmd
//----------------------------------------------
var WrapCommand = function(){};
WrapCommand.prototype = {
	HASHTYPE : {
		MD5    : "MD5",
		SHA1   : "SHA1",
		SHA256 : "SHA256"
	},
	exec : function(cmd){
		var sh = WScript.CreateObject('WScript.Shell');
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
		var sh = WScript.CreateObject('WScript.Shell');
		var file = sh.CreateShortcut(topath);
		var prefix = String(frompath).match(/\.lnk$/) ? "file:/" : "";
		file.TargetPath = prefix + String(frompath);
		file.Save();
	},
	env : function(name){
		var sh = WScript.CreateObject('WScript.Shell');
		return sh.ExpandEnvironmentStrings(name);
	}
};
WindowsJScript.prototype.cmd = new WrapCommand();

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
// WindowsJScript.date
//----------------------------------------------
var WrapDate = function(){ return this; };
WrapDate.prototype = {
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
var WrapPath = function(){ return this.init(); };
WrapPath.prototype = {
	fso : null,
	OPENMODE : {
		FORREAD   : 1,     // [default]
		FORWRITE  : 2,
		FORAPPEND : 8
	},
	NOTHINGTHEN : {
		CREATE :  true,
		THROUGH : false    // [default]
	},
	CHARCODE : {
		UTF    : -1,
		SJIS   : 0,        // [default]
		SYSTEM : -2
	},
	TYPEIS : {
		DIR  : 1,
		FILE : 2,
		ALL  : 9
	},
	NEWLINECODE : {
		CR   : "\r",
		LF   : "\n",
		CRLF : "\r\n"
	},
	init : function(){
		this.fso = new ActiveXObject("Scripting.FileSystemObject");
		return this;
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
	openfile : function(path, txt, opt){
		opt = js.extend(this.defaultoptions(), opt);

		var stream = this.fso.OpenTextFile(path, opt.OPENMODE, opt.NOTHINGTHEN, opt.CHARCODE);

		if(opt.OPENMODE === this.OPENMODE.FORREAD){
			// テキスト初期化
			var text = "";

//			// 読み取り方法① ファイルから全ての文字データを読み込む
//			text = stream.ReadAll();
//			// ファイルの末尾までループ
//			while (!stream.AtEndOfStream) {
//				// 読み取り方法② ファイルの文字データを一行ずつ表示する
//				text += stream.ReadLine();
//				// 読み取り方法③ 読み込みバッファ指定でループする
//				text += stream.Read(1024);
//			}
//			// 読み取り方法④ 全行を読み込む
//			text = stream.ReadText(-1);
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
//			//  ブランク行を二行書き込む
//			stream.WriteBlankLines( 2 );

			stream.Close();
			return true;
		}
	},
	readfile : function(path, opt){
		if(!js.isobject(opt)) opt = {};
		opt.OPENMODE = this.OPENMODE.FORREAD;
		return this.openfile(path, null, opt);
	},
	writefile : function(path, txt, opt){
		if(!js.isobject(opt)) opt = {};
		opt.OPENMODE = this.OPENMODE.FORWRITE;
		return this.openfile(path, txt, opt);
	},
	appendfile : function(path, txt, opt){
		if(!js.isobject(opt)) opt = {};
		opt.OPENMODE = this.OPENMODE.FORAPPEND;
		return this.openfile(path, txt, opt);
	},
	info : function(path){
		var rtn = {
			fullpath : this.fso.GetAbsolutePathName(path),
			filename : this.fso.GetFileName(path),
			basename : this.fso.GetBaseName(path),
			extname  : this.fso.GetExtensionName(path),
			isdir    : this.isdir(path),
			isfile   : this.isfile(path)
		};
		if(rtn.isfile){
			var obj = this.fso.GetFile(path);
			rtn.attr = obj.Attributes;
			rtn.createdate = obj.DateCreated;
			rtn.updatedate = obj.DateLastModified;
			rtn.drive = obj.Drive;
			rtn.name = obj.Name;
			rtn.parent = obj.ParentFolder;
			rtn.path = obj.Path;
			rtn.type = obj.Type;
			rtn.size = obj.Size;
		}
		if(rtn.isdir){
			var obj = this.fso.GetFolder(path);
			rtn.createdate = obj.DateCreated;
			rtn.updatedate = obj.DateLastModified;
			rtn.drive = obj.Drive;
			rtn.name = obj.Name;
			rtn.parent = obj.ParentFolder;
			rtn.path = obj.Path;
			rtn.type = obj.Type;
			rtn.size = obj.Size;
			rtn.isroot = obj.IsRootFolder;
		}
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
		}
		if(this.isdir(path)){
			if(opt.dir){
				rtn.push(func(path));
			}
			var dir = this.fso.GetFolder(path);
			if(opt.file){
				for (var fp = new Enumerator(dir.Files); !fp.atEnd(); fp.moveNext()) {
					if(opt.filter && !String(fp.item()).match(opt.filter)){
						continue;
					}
					rtn.push.apply(rtn, this.each(fp.item(), func, opt));
				}
			}
			if(js.isnumber(opt.nest) && opt.curr >= opt.nest) return rtn;
			opt.curr++;
			for (var fp = new Enumerator(dir.SubFolders); !fp.atEnd(); fp.moveNext()) {
				if(opt.filter && !String(fp.item()).match(opt.filter)){
					continue;
				}
				rtn.push.apply(rtn, this.each(fp.item(), func, opt));
			}
		}
		return rtn;
	},
	remove: function(path){ return this.isdir(path) ? this.rmdir(path) : this.isfile(path) ? this.rm(path) : false; },
	copy  : function(from, to){ return this.isdir(from) ? this.cpdir(from, to) : this.isfile(from) ? this.cp(from, to) : false; },
	move  : function(from, to){ return this.isdir(from) ? this.mvdir(from, to) : this.isfile(from) ? this.mv(from, to) : false; },
	rm    : function(path){ this.fso.DeleteFile(path); },
	rmdir : function(path){ this.fso.DeleteFolder(path); },
	cp    : function(from, to){ this.fso.CopyFile(from, to); },
	cpdir : function(from, to){ this.fso.CopyFolder(from, to); },
	mv    : function(from, to){ this.fso.MoveFile(from, to); },
	mvdir : function(from, to){ this.fso.MoveFolder(from, to); },
	touch : function(path, ismkdir){
		if(!!ismkdir){ this.mkdir(this.parent(path)); }
		return this.openfile(path, null, this.OPENMODE.FORWRITE, this.NOTHINGTHEN.CREATE, this.CHARCODE.SYSTEM);
	},
	mkdir : function(path){
		if(this.isdir(path)) return true;
		this.mkdir(this.parent(path));
		this.fso.CreateFolder(path);
		return true;
	},
	join : function(){
		var self = this;
		var args = [].slice.call(arguments);
		if(js.isobject(args[0])){
			args = args[0];
		}
		var path = args.shift();
		js.each(args, function(i, val){
			path = self.fso.BuildPath(path, val);
		});
		return path;
	},
	tmpname : function(){ return this.fso.GetTempName(); },
	isdrive : function(path){ return this.fso.DriveExists(path); },
	drive  : function(path){ return this.fso.GetDriveName(path); },
	parent : function(path){ return this.fso.GetParentFolderName(path) },
	isfile : function(path){ return this.fso.FileExists(path); },
	isdir  : function(path){ return this.fso.FolderExists(path); },
	exist  : function(path){ return this.isdir || this.isfile; }
};
WindowsJScript.prototype.path = new WrapPath();

//----------------------------------------------
// WindowsJScript.log
//----------------------------------------------
var WrapLog = function(state, path){ this.init(state, path); return this; };
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
	ready : function(){
 		this.state.ready = js.path.isdir(js.path.parent(this.path)) ? true : false;
	},
	init : function(state, path){
		var currpath = (this.path) ? this.path : String(WScript.ScriptFullName).replace(/\.[^\.]+$/, js.str.format("_${0}.log", js.date.now("YYYYMMDDhhmmss")));
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
var WrapBook = function(){ this.init(); return this; };
WrapBook.prototype = {
	excel : null,
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
	init : function(){
		this.excel = new ActiveXObject("Excel.Application");
	},
	isbook : function(path){
		if(!js.path.isfile(path)){
			return false;
		}
		if(!String(path).match(/\.xls[ms]?$/)){
			return false;
		}
		return true;
	},
	read : function(path, func){
		if(!this.isbook) return false;
		var book = null;
		try{
			book = this.excel.Workbooks.Open(path);
			return func(book);
		} catch(e) {
			js.echo(js.str.format("[js.book.read] error(${1}:$2). path:${0}, ", path, (e.number & 0xFFFF), e.message));
		} finally {
			if(!js.isnull(book)) {
				book.Close(false);
				book = null;
			}
		}
	},
	write : function(path, func){
		if(!this.isbook) return false;
		this.excel.Visible = false;
		var book = null;
		try{
			book = this.excel.Workbooks.Add();
			func(book);
			book.CheckCompatibility = false;
			this.excel.DisplayAlerts = false;
			book.SaveAs(path);
		} catch(e) {
			js.echo(js.str.format("[js.book.write] error(${1}:$2). path:${0}, ", path, (e.number & 0xFFFF), e.message));
		} finally {
			if(!js.isnull(book)) {
				this.excel.CutCopyMode = false;
				book.Close(false);
				book = null;
			}
		}
	}
};
WindowsJScript.prototype.book = new WrapBook();
