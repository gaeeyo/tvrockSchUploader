// TvRock の tvrock.sch を変換してしょぼいカレンダーにアップロードするツール
//
//  cscript syobocalUploader.js <user> <pass> [epgurl] [slot]
//
//   <user>        しょぼいカレンダーのUserID
//   <pass>        しょぼいカレンダーのパスワード
//   [epgurl]      TvRockのURL
//                 例) http://recserver:8969/nobody/
//   [slot]        0～3の数値(未指定の場合0)
//

// 色の設定
//  "{COMPNAME}=#RRGGBB" か "#RRGGBB" のどちらかで指定
//  COMPNAMEはtvrock.schのCOMPNAMEの先頭と末尾の文字列になる
//  (例> "MAINPC"=>"MC", "SUBPC"=>"SC")
var _devColors = [
	//'MC:Cap=#887744',
	//'SC=#447788',
	'#777777'
];

var _userAgent = 'tvrockSchUploader/1.2.2';
var _uploadUrl = 'http://cal.syoboi.jp/sch_upload';
 
main(WScript.Arguments);

function main(args) {
	if (args.length < 2) {
		WScript.Echo('tvrockSchUploader.js <user> <pass> [epgurl]');
	}
	else {
		var items = loadSchFile('tvrock.sch');
		var sch_data = formatItems(items);
		var sch_epgurl = (args.length > 2 ? args(2) : '');
		var slot = (args.length > 3 ? args(3) : 0);
		
		upload(args(0), args(1), sch_data, sch_epgurl, slot);
	}
}

// アップロードするデータの形式に変換(tsvに)
function formatItems(items)
{
	var text = '';
	
	for (var j=0; j<items.length; j++) {
		var item = items[j];
		text += [
			item.START,
			item.END,
			item.DEV,
			tsvEscape(item.TITLE),
			tsvEscape(item.STATION),
			tsvEscape(
				[
					(item.NUMBER != '0' ? '#'+item.NUMBER : ''),
					(item.SUBTITLE != '未定' ? item.SUBTITLE : '')
				].join(' ')
			),
			item.OFFSET,
			item.UNIQID
		].join("\t")+"\n";
	}
	// WScript.Echo(text.replace(/\t/g,' ')); WScript.Quit(0);
	return text;
	
}

function tsvEscape(text)
{
	return text.replace("\t", " ");
}

// アップロード
function upload(user, pass, sch_data, sch_epgurl, slot)
{
	var http = new ActiveXObject('MSXML2.XMLHTTP');
	
	http.Open('POST', _uploadUrl+'?slot='+slot, false,
		encodeURIComponent(user), encodeURIComponent(pass)
	);
	http.setRequestHeader('Content-type', 'application/x-www-form-urlencoded');
	http.setRequestHeader('User-agent', _userAgent);
	http.onreadystatechange = function(){
		if (http.readyState == 4) {
			if (http.status == 200) {
				WScript.Echo(http.responseText);
			}
			else {
				WScript.Echo('UPLOAD ERROR: '+http.status);
			}
		}
	};
	http.send(''
		+'devcolors='+encodeURIComponent(_devColors.join("\t"))
		+'&epgurl='+encodeURIComponent(
			(sch_epgurl != '' ? sch_epgurl+'reg?i={XID}' : '')
		)
		+'&data='+encodeURIComponent(sch_data)
	);
}

// tvrock.sch を読み込んで配列を返す
function loadSchFile(path)
{
	var fso = new ActiveXObject('Scripting.FileSystemObject');
	var ts = fso.OpenTextFile(path, 1);

	var devMap = ['Cap','Ex1','Ex2','Ex3','Ex4','T1','T2','T3','T4','T5','T6','T7','T8','R'];
	
	var items = [];
	items.add = function(item) {
		if (item.VALIDATE != '0') {
			var compname = item.COMPNAME.substr(0,1) + item.COMPNAME.slice(-1) + ':';
			item.DEV = compname + devMap[item.DEVNO];
			item.OFFSET = item.START - item.STBK;
			//WScript.Echo([item.START, item.END, item.DEV, item.TITLE, item.STATION, item.SUBTITLE]);

			this.push(item);
		}
	}
	
	var number = null;
	var item = {};
	while (!ts.AtEndOfLine) {
		var cps = ts.ReadLine().match(/^(\d+) (\S+) (.*)$/);
		if (cps) {
			if (number != cps[1]) {
				if (number) {
					items.add(item);
				}
				number = cps[1];
				item = {idx:number};
			}
			if (cps[2].match(/^(TITLE|STATION|START|STBK|END|DEVNO|SUBTITLE|VALIDATE|UNIQID|COMPNAME|NUMBER)$/)) {
				item[cps[2]] = cps[3];
			}
		}
	}
	if (number) {
		items.add(item);
	}

	return items;
}
