#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# IDシートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import unohelper, os
from collections import OrderedDict
from . import commons, staticdialog, ichiran
from com.sun.star.awt import MouseButton, MessageBoxButtons, MessageBoxResults # 定数
from com.sun.star.awt.MessageBoxType import QUERYBOX, WARNINGBOX  # enum
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.sheet.CellDeleteMode import COLUMNS as delete_columns  # enum
from com.sun.star.table.CellHoriJustify import CENTER  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
class IDsheet():  # シート固有の値。
	def __init__(self):
		self.splittedrow = 2  # 分割行インデックス。
		self.mincolumn = 3  # 1日の最低点列のインデックス。
		self.daycolumn = 4  # 日付列インデックス。
		self.startcolumn = 5  # 開始列インデックス。	
		items = ("深さ", ("0: 皮膚損傷・発赤なし", "1: 持続する発赤", "2: 真皮までの損傷", "3: 皮下組織までの損傷", "4: 皮下組織を越える損傷", "5: 関節腔、体腔に至る損傷")),\
			("浸出液", ("0: なし", "1: 少量:毎日のドレッシング交換を要しない", "3: 中等量:1日1回のドレッシング交換を要する", "6: 多量:1日2回以上のドレッシング交換を要する")),\
			("大きさ", ("0: 皮膚損傷なし", "3: 4未満", "6: 4以上16未満", "8: 16以上36未満", "9: 36以上64未満", "12: 64以上100未満", "15: 100以上")),\
			("炎症・感染", ("0: 局所の炎症徴候なし", "1: 局所の炎症徴候あり(創周囲の発赤、腫脹、熱感、疼痛)", "3: 局所の明らかな感染徴候あり(炎症徴候、膿、悪臭など)", "9: 全身的影響あり(発熱など)")),\
			("肉芽形成", ("0: 治癒あるいは創が浅いため肉芽形成の評価ができない", "1: 良性肉芽が創面の90%以上を占める", "3: 良性肉芽が創面の50%以上90%未満を占める", "4: 良性肉芽が、創面の10%以上50%未満を占める", "5: 良性肉芽が、創面の10%未満を占める", "6: 良性肉芽が全く形成されていない")),\
			("壊死組織", ("0: 壊死組織なし", "3: 柔らかい壊死組織あり", "6: 硬く厚い密着した壊死組織あり")),\
			("ポケット", ("0: ポケットなし", "6: 4未満", "9: 4以上16未満", "12: 16以上36未満", "24: 36以上"))
		self.dic = OrderedDict([(k, v) for k, v in items])  # 通常のdictは順番が一定でないのでOrderedDictを使う。
	def setSheet(self, sheet):  # 逐次変化する値。
		self.sheet = sheet
		cellranges = sheet[:, self.daycolumn].queryContentCells(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME)  # ID列の文字列、数値、日付が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 日付列の最終行インデックス+1を取得。
		cellranges = sheet[self.splittedrow-1, :].queryContentCells(CellFlags.STRING)  # 文字列が入っているセルに限定して抽出。
		self.emptycolumn = cellranges.getRangeAddresses()[-1].EndColumn + 1  # 分割行の上行の最終列インデックス+1を取得。
VARS = IDsheet()
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
	VARS.setSheet(sheet)
	splittedrow = VARS.splittedrow
	emptycolumn = VARS.emptycolumn
	datarange = sheet[0, :emptycolumn+1]
	datarows = list(datarange.getDataArray()[0])  # 0行目をリストで取得。
	datarows[0] = "一覧へ"
	datarows[2] = "月更新"
	datarows[-1] = "部位追加"
	datarange.setDataArray((datarows,))  # ボタンになっているセルを修正した行をシートに戻す。
	if VARS.startcolumn<emptycolumn:  # 部位があるとき。
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
		functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。			
		startdatevalue = int(sheet[splittedrow, VARS.daycolumn].getValue())  # シートの開始日のシリアル値を取得。
		datevalues = [i for i in range(startdatevalue, startdatevalue+VARS.emptyrow-splittedrow)]  # シートの日付のシリアル値のリストを作成。
		todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。
		if todayvalue in datevalues:  # 今日がシリアル値のリストにある時。
			enddayrow = splittedrow + datevalues.index(todayvalue)  # 今日の行インデックスを取得。
		elif todayvalue>datevalues[-1]:  # シートの最終日がすでに過ぎた日の時は最終行までコピーする。
			enddayrow = splittedrow + len(datevalues) - 1
		else:  # 今日がシートの日付より前のときは何もしない。
			return	
		fillToEndDayRow(xscriptcontext.getDocument(), enddayrow)
		if enddayrow+1<VARS.emptyrow:  # 月末日以外の時。
			sheet[enddayrow+1:VARS.emptyrow, VARS.startcolumn:emptycolumn].clearContents(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME+CellFlags.FORMULA)  # 今日の行より下のセルの内容をクリア。
def fillToEndDayRow(doc, enddayrow):  # 各部位について最終行をenddayrowまでコピーして各行の最低点を算出。	
	sheet = VARS.sheet
	splittedrow = VARS.splittedrow
	emptycolumn = VARS.emptycolumn	
	for i in range(VARS.startcolumn, emptycolumn, 8):  # 部位別開始列をイテレート。
		cellranges = sheet[:, i+7].queryContentCells(CellFlags.STRING+CellFlags.VALUE+CellFlags.FORMULA)  # 部位別合計列の文字列、数値、式が入っているセルに限定して抽出。
		endrow = cellranges.getRangeAddresses()[-1].EndRow  # 列の最終行インデックスを取得。
		if splittedrow<=endrow<enddayrow:  # 部位別合計の最下行が今日の行より上にあるとき。
			if sheet[endrow+1, i+7].getPropertyValue("CellBackColor")>0:  # 最終行下行の部位別行合計に背景色があるとき。
				continue
			enddatarows = sheet[endrow, i:i+8].getDataArray()  # 最終行のタプルを取得。
			newdatarows = enddatarows*(enddayrow-endrow)  # 最終行を複製。
			sheet[endrow+1:enddayrow+1, i:i+8].setDataArray(newdatarows)  # 最終行を今日の行までコピー。
	sheet[splittedrow:VARS.emptyrow, VARS.mincolumn].setPropertyValue("CellBackColor", -1)	 # 最低点列の背景色をクリア。	
	datarows = sheet[splittedrow:enddayrow+1, :emptycolumn].getDataArray()  # 分割行から今日の行までの空列までのデータ行のタプルを取得。
	prevs = datarows[0][:VARS.mincolumn]  # 3月前の最低点のタプルを取得。
	cs = range(VARS.startcolumn+7, emptycolumn, 8)  # 部位別合計列インデックスのジェネレーター。
	mindatarows = [(min([d[i] for i in cs if not d[i]==""], default=""),) for d in datarows]  # 部位別合計列の最低点の行のタプルのリスト。
	highlightPenaltyDays(doc, prevs, mindatarows)
	sheet[splittedrow:splittedrow+len(mindatarows), VARS.mincolumn].setDataArray(mindatarows)  # 日の最低点のセル範囲に代入。		
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	if enhancedmouseevent.ClickCount==2 and enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ダブルクリックの時。まずselectionChanged()が発火している。
		selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			celladdress = selection.getCellAddress()
			r, c = celladdress.Row, celladdress.Column  # selectionの行インデックスと列インデックスを取得。
			sheet = selection.getSpreadsheet()
			if r==0:  # 行インデックス0の時。
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。					
				doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
				controller = doc.getCurrentController()  # コントローラの取得。	
				txt = selection.getString()
				if txt=="一覧へ":			
					controller.setActiveSheet(doc.getSheets()["一覧"])  # 一覧シートをアクティブにする。
				elif txt=="月更新":
					functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。		
					datarow = list(sheet[VARS.emptyrow-1, :VARS.emptycolumn].getDataArray()[0])  # 最終行を右端列までリストで取得。
					y, m = [int(functionaccess.callFunction(i, (datarow[VARS.daycolumn],))) for i in ("YEAR", "MONTH")]  # 最終行の日付セルの年と月を取得。	
					olddatetxt = "{}年{}月".format(y, m)  # 現在の年と月。
					ny = y
					nm = m + 1
					if nm>12:  # 年と月を更新。
						ny = y + 1
						nm = 1
					newdatetxt = "{}年{}月".format(ny, nm)  # 更新後の年と月。	
					msg = "{}を{}に更新します。\n古いシートは別ファイルに保存します。".format(olddatetxt, newdatetxt)
					if showWarningMessageBox(controller, msg)==MessageBoxResults.OK:					
						createCopySheet(xscriptcontext, y)(sheet.getName(), m)  # 現在のシートを年月名のファイルにコピーする。
						splittedrow = VARS.splittedrow
						prevs = sheet[splittedrow, :VARS.mincolumn].getDataArray()[0]  # 3月前分の最低点のタプルを取得。
						datarow[VARS.daycolumn] += 1  # シート最下行の日付シリアル値に1を加えて翌月初日にする。
						datarow[0] = prevs[1]  # 2月前の最低点を3月前の最低点に変更。
						datarow[1] = prevs[2]  # 1月前の最低点を2月前の最低点に変更。
						datarow[2] = min([i[0] for i in sheet[splittedrow:VARS.emptyrow, VARS.mincolumn].getDataArray() if i[0]], default="")  # 現シートの最低点を1月前の最低点にする。
						sheet[splittedrow, :VARS.emptycolumn].setDataArray((datarow,))  # 変更した最下行を1日目に代入。
						sheet[splittedrow+1:VARS.emptyrow, :VARS.emptycolumn].clearContents(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME+CellFlags.FORMULA)  # 2日目以降の文字列、数値、日付、式をクリア。
						sheet[splittedrow:VARS.emptyrow, :VARS.emptycolumn].setPropertyValue("CellBackColor", -1)  # 1日目以降の背景色をクリア。
						for i in range(VARS.startcolumn, VARS.emptycolumn, 8)[::-1]:  # 降順に部位別の開始列インデックスをイテレート。
							if not any(datarow[i:i+8]):  # 部位別のセルに空セルがある時。
								sheet.removeRange(sheet[0, i:i+8].getRangeAddress(), delete_columns)  # その部位の列を削除。	
						datevalue = int(datarow[VARS.daycolumn])  # 更新後の初日のシリアル値を整数で取得。
						daycount = int(functionaccess.callFunction("DAYSINMONTH", (datevalue,)))  # 更新後の月の日数を取得。
						sheet[splittedrow:splittedrow+daycount, VARS.daycolumn].setDataArray((i,) for i in range(datevalue, datevalue+daycount))  # 全日付を更新。
						sheet[splittedrow+1:splittedrow+daycount, :VARS.mincolumn].setPropertyValue("CellBackColor", commons.COLORS["silver"])  # 2日以降の前月の値のセルの背景色を付ける。	
						prevs = sheet[VARS.splittedrow, :VARS.daycolumn].getDataArray()[0]  # 3月前の最低点のタプルと行の日の最低点を取得。
						color = commons.COLORS["magenta3"] if not "" in prevs and prevs[0]<prevs[1]<prevs[2]<prevs[3] else -1  # ペナルティの時は最低点セルに背景色を付ける、そうでないなら背景色をクリアする。
						sheet[splittedrow, VARS.mincolumn].setPropertyValue("CellBackColor", color)		
						colorizeDays(doc, functionaccess, datevalue)
				elif txt=="部位追加":
					if (c-VARS.startcolumn)%8==0:  # 部位の先頭列であることを確認する。
						datarows = [("", "", "", "", "", "", "", "", "部位追加")]
						datarows.append(list(VARS.dic.keys()))
						datarows[-1].extend(("部位別合計", "")) 		
						endedge = c + len(datarows[0]) - 1
						sheet[:VARS.splittedrow, c:endedge+1].setDataArray(datarows)  # 分割行より上のセルを代入。
						sheet[0, c:endedge].getColumns().setPropertyValue("Width", 680)  # 列幅を設定。
						sheet[0, c:endedge].merge(True)  # 行インデックス0を結合。
						VARS.setSheet(sheet)  # 逐次変化する値を取得し直す。VARS.emptycolumnが変化する。
						idtxt = sheet[0, VARS.daycolumn].getString()
						ichiransheet = doc.getSheets()["一覧"]
						ichiranvars = ichiran.VARS
						ichiranvars.setSheet(ichiransheet)
						searchdescriptor = ichiransheet.createSearchDescriptor()
						searchdescriptor.setSearchString(idtxt)  # 戻り値はない。IDの文字列が入っているセルを探す。
						idcell = ichiransheet[ichiranvars.splittedrow:ichiranvars.emptyrow, ichiranvars.idcolumn].findFirst(searchdescriptor)  # 見つからなかった時はNoneが返る?。
						if idcell:
							startdatevalue = (ichiransheet[idcell.getCellAddress().Row, ichiranvars.startdaycolumn].getValue(),)  # 一覧シートにある開始日のシリアル値を行で取得。
							datevalues = sheet[VARS.splittedrow:VARS.emptyrow, VARS.daycolumn].getDataArray()
							if startdatevalue in datevalues:  # 一覧シートの開始日と一致する日付があるときはその行の上まで背景色をつける。
								startrow = VARS.splittedrow+datevalues.index(startdatevalue)
								if VARS.splittedrow<startrow:  # 1日は除く。
									sheet[VARS.splittedrow:startrow, c:c+8].setPropertyValue("CellBackColor", commons.COLORS["silver"])  # 背景色をつける
					else:  # 部位の先頭列でないときはエラーメッセージを出す。
						msg = "部位の先頭列ではありません。"
						commons.showErrorMessageBox(controller, msg)
				elif c==VARS.daycolumn:  # IDセルの時。IDをコピーする。
					systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
					systemclipboard.setContents(commons.TextTransferable(txt), None)  # クリップボードにIDをコピーする。
				elif (c-VARS.startcolumn)%8==0:  # 部位の先頭列の時。部位が入る。上下左右はコンテクストメニューで追加する。
					defaultrows = "肩", "腰椎部", "仙骨部", "坐骨部", "大転子部", "腓骨部", "腓腹部", "足関節外側", "踵", 
					staticdialog.createDialog(enhancedmouseevent, xscriptcontext, "部位選択", defaultrows)  # 定型句ダイアログを作成。
					selection.setPropertyValue("HoriJustify", CENTER)
			elif VARS.splittedrow<=r<VARS.emptyrow and VARS.startcolumn<=c<VARS.emptycolumn:  # 点数セルの時。
				if (c-VARS.startcolumn)%8!=7:  # 部位の最終行以外の時。
					headertxt = VARS.sheet[VARS.splittedrow-1, c].getString()  # ヘッダー文字列を取得。
					defaultrows = VARS.dic.get(headertxt, None)  # グリッドコントロールのデフォルト行を習得。
					gridcontrol1, datarows = staticdialog.createDialog(enhancedmouseevent, xscriptcontext, headertxt, defaultrows, callback=callback_wClickPointsCreator(xscriptcontext))  # 列ヘッダー毎に定型句ダイアログを作成。	
					selection = enhancedmouseevent.Target
					valtxt = selection.getString()  # セルの数値を文字列として取得。
					if not valtxt:  # 空セルのときは0にする。
						valtxt = "0"
						selection.setValue(0)
					txt = "{}:".format(enhancedmouseevent.Target.getString())  # セルの入っている数字を文字列で取得。
					for i in range(len(datarows)):
						if datarows[i][0].startswith(txt):
							gridcontrol1.selectRow(i)  # 先頭が一致するグリッドコントロールの行をハイライト。
							break	
		return False  # セル編集モードにしない。	
	return True  # セル編集モードにする。	シングルクリックは必ずTrueを返さないといけない。
def colorizeDays(doc, functionaccess, startdatevalue):
	y, m = [int(functionaccess.callFunction(i, (startdatevalue,))) for i in ("YEAR", "MONTH")]
	holidays = commons.HOLIDAYS	
	holidayindexes = set()
	if y in holidays:
		holidayindexes.update(VARS.splittedrow-1+i for i in holidays[y][m-1])  # 行インデックスに変換するして取得する。
	startweekday = int(functionaccess.callFunction("WEEKDAY", (startdatevalue, 3)))  # 開始日の曜日を取得。月=0。
	n = 6  # 日曜日の曜日番号。
	sunindexes = set(range(VARS.splittedrow+(n-startweekday)%7, VARS.emptyrow, 7))  # 日曜日の列インデックスの集合。祝日と重ならないようにあとで使用する。	
	holidayindexes.difference_update(sunindexes)  # 祝日インデックスから日曜日インデックスを除く。
	n = 5  # 土曜日の曜日番号。
	satindexes = set(range(VARS.splittedrow+(n-startweekday)%7, VARS.emptyrow, 7))  # 土曜日の列インデックスの集合。
	VARS.sheet[VARS.splittedrow:VARS.emptyrow, VARS.daycolumn].setPropertyValues(("CellBackColor", "CharColor"), (-1, -1))  # 日付列の背景色と字の色をリセットする。
	setRangesProperty = createSetRangesProperty(doc, VARS.daycolumn)
	setRangesProperty(holidayindexes, ("CellBackColor", commons.COLORS["red3"]))
	setRangesProperty(sunindexes, ("CharColor", commons.COLORS["red3"]))
	setRangesProperty(satindexes, ("CharColor", commons.COLORS["skyblue"]))
def createSetRangesProperty(doc, c): 
	def setRangesProperty(rowindexes, prop):  # c列のrowindexesの行のプロパティを変更。prop: プロパティ名とその値のリスト。
		cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
		if rowindexes:  
			cellranges.addRangeAddresses([VARS.sheet[i, c].getRangeAddress() for i in rowindexes], False)  # セル範囲コレクションを取得。rowindexesが空要素だとエラーになる。
			if len(cellranges):  # sheetcellrangesに要素がないときはsetPropertyValue()でエラーになるので要素の有無を確認する。
				cellranges.setPropertyValue(*prop)  # セル範囲コレクションのプロパティを変更。
	return setRangesProperty		
def createCopySheet(xscriptcontext, year):	
	desktop = xscriptcontext.getDesktop()
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
	controller = doc.getCurrentController()  # コントローラの取得。	
	sheets = doc.getSheets()
	dirpath = os.path.dirname(unohelper.fileUrlToSystemPath(doc.getURL()))  # このドキュメントのあるディレクトリのフルパスを取得。
	yeardirpath = os.path.join(dirpath, "{}年".format(year))  
	if not os.path.exists(yeardirpath):
		os.mkdir(yeardirpath) 
	def copySheet(sheetname, month):
		if sheetname in sheets:  # シートがある時。
			newdocpath = os.path.join(yeardirpath, "{}年{}月.ods".format(year, month))  # アーカイブファイルのフルパスを取得。
			fileurl = unohelper.systemPathToFileUrl(newdocpath)  # fileurlに変換。
			newfileurl = fileurl if os.path.exists(newdocpath) else "private:factory/scalc"
			newdoc = desktop.loadComponentFromURL(newfileurl, "_blank", 0, ())  # 新規ドキュメントの取得。隠し属性だと行と列の固定ができない。
			newsheets = newdoc.getSheets()  # 新規ドキュメントのシートコレクションを取得。
			if sheetname in newsheets:  # すでにシートが存在する時。
				msg = "シート{}はすでに保存済です。\n上書きしますか？".format(sheetname)
				componentwindow = controller.ComponentWindow
				msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "myRs", msg)
				if msgbox.execute()!=MessageBoxResults.YES:	 # YESではないときはこのまま終わる。
					return	
			newsheets.importSheet(doc, sheetname, len(newsheets))  # 新規ドキュメントのシートの最後にコピー。コピー先のシートの行と列の固定が解除されてしまう。
			newdoccontroller = newdoc.getCurrentController()  # コピー先のドキュメントのコントローラを取得。	
			newdoccontroller.setActiveSheet(newsheets[sheetname])  # コピーしたシートをアクティブにする。
			newdoccontroller.freezeAtPosition(VARS.startcolumn, VARS.splittedrow)  # 行と列の固定をする。
			if "Sheet1" in newsheets:
				del newsheets["Sheet1"]  # 新規ドキュメントのデフォルトシートを削除する。1枚しかシートがない時はエラーになる。	
			newdoc.storeAsURL(fileurl, ())  # アーカイブファイルを保存。  
			newdoc.close(True)  # アーカイブファイルを閉じる。
		else:
			msg = "シート{}が存在しません。".format(sheetname)	
			commons.showErrorMessageBox(controller, msg)	
	return copySheet
def callback_wClickPointsCreator(xscriptcontext):
	def callback_wClickPoints(gridcelldata):
		selection = xscriptcontext.getDocument().getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
		selection.setValue(int(gridcelldata.split(":", 1)[0]))  # 点数のみにして数値としてセルに代入し直す。
		celladdress = selection.getCellAddress()
		reCalc(celladdress.Row, celladdress.Column)  # 部位別合計点と日の最低点を計算。
	return callback_wClickPoints
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	selection = eventobject.Source.getSelection()
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
		VARS.setSheet(selection.getSpreadsheet())		
		drowBorders(selection)  # 枠線の作成。			
def drowBorders(selection):  # ターゲットを交点とする行列全体の外枠線を描く。
	rangeaddress = selection.getRangeAddress()  # 選択範囲のセル範囲アドレスを取得。	
	sheet = VARS.sheet
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = commons.createBorders()
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。	
	startrow = VARS.splittedrow if rangeaddress.StartRow<VARS.splittedrow else rangeaddress.StartRow
	edgerow = rangeaddress.EndRow+1 if rangeaddress.EndRow<VARS.emptyrow else VARS.emptyrow
	edgecolmun = rangeaddress.EndColumn+1 if rangeaddress.EndColumn<VARS.emptycolumn else VARS.emptycolumn
	if startrow<edgerow and rangeaddress.StartColumn<edgecolmun:
		sheet[startrow:edgerow, :VARS.emptycolumn].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く
		sheet[VARS.splittedrow-1:VARS.emptyrow, rangeaddress.StartColumn:edgecolmun].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。
		selection.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。		
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。	
	contextmenuname, addMenuentry, baseurl, selection = commons.contextmenuHelper(VARS, contextmenuexecuteevent, xscriptcontext)
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上角のセルのアドレスを取得。
	r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。		
	if contextmenuname=="cell":  # セルのとき	
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if r==0 and (c-VARS.startcolumn)%8==0:  # 部位セルの時。
				addMenuentry("ActionTrigger", {"Text": "左", "CommandURL": baseurl.format("entry2")}) 		
				addMenuentry("ActionTrigger", {"Text": "右", "CommandURL": baseurl.format("entry3")}) 		
				addMenuentry("ActionTrigger", {"Text": "左右なし", "CommandURL": baseurl.format("entry8")}) 		
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				addMenuentry("ActionTrigger", {"Text": "--上", "CommandURL": baseurl.format("entry4")}) 		
				addMenuentry("ActionTrigger", {"Text": "--下", "CommandURL": baseurl.format("entry5")}) 		
				addMenuentry("ActionTrigger", {"Text": "--左", "CommandURL": baseurl.format("entry6")}) 		
				addMenuentry("ActionTrigger", {"Text": "--右", "CommandURL": baseurl.format("entry7")}) 
				addMenuentry("ActionTrigger", {"Text": "--なし", "CommandURL": baseurl.format("entry9")}) 
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				commons.cutcopypasteMenuEntries(addMenuentry)					
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				addMenuentry("ActionTrigger", {"Text": "セル内容をクリア", "CommandURL": baseurl.format("entry1")}) 	
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				thisc = c-(c-VARS.startcolumn)%8	
				ptxt = VARS.sheet[0, thisc].getString()				
				addMenuentry("ActionTrigger", {"Text": "{} を削除".format(ptxt), "CommandURL": baseurl.format("entry12")}) 
			elif VARS.startcolumn<=c<VARS.emptycolumn or c in (VARS.mincolumn,):  # 点数セル、または、最低点列、の時。
				commons.cutcopypasteMenuEntries(addMenuentry)	
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				addMenuentry("ActionTrigger", {"Text": "セル内容をクリア", "CommandURL": baseurl.format("entry1")}) 	
				if not c in (VARS.mincolumn,): 
					addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
					thisc = c-(c-VARS.startcolumn)%8	
					ptxt = VARS.sheet[0, thisc].getString()
					if VARS.splittedrow<r<VARS.emptyrow:
						addMenuentry("ActionTrigger", {"Text": "{} 開始日にする".format(ptxt), "CommandURL": baseurl.format("entry10")})
						addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
					if VARS.splittedrow<=r<VARS.emptyrow-1:
						addMenuentry("ActionTrigger", {"Text": "{} 終了日にする".format(ptxt), "CommandURL": baseurl.format("entry11")})	
		elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # ターゲットがセル範囲の時。
			if VARS.startcolumn<=c<VARS.emptycolumn or c in (VARS.mincolumn,):  # 左上が点数セル、または、最低点列、の時。
				commons.cutcopypasteMenuEntries(addMenuentry)	
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				addMenuentry("ActionTrigger", {"Text": "セル内容をクリア", "CommandURL": baseurl.format("entry1")}) 						
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Remove"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:RenameTable"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。	
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。	
	controller = xscriptcontext.getDocument().getCurrentController()  # コントローラの取得。
	selection = controller.getSelection()  # 選択範囲を取得。
	if entrynum==1:  # セル内容をクリア。書式設定とオブジェクト以外を消去。
		selection.clearContents(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME+CellFlags.FORMULA)  # セル内容をクリアする。	
	elif entrynum in (2, 3, 8):  # 前に左右をつける。
		if entrynum==2:  # 左。
			prefix = "左"
		elif entrynum==3:  # 右。
			prefix = "右"			
		elif entrynum==8:  # 左右なし
			prefix = ""
		selection.setString("{}{}".format(prefix, selection.getString().lstrip("左右")))
	elif entrynum in (4, 5, 6, 7, 9):  # 後に上下左右をつける。	
		if entrynum==4:  # 上。
			suffix = "上"			
		elif entrynum==5:  # 下。
			suffix = "下"			
		if entrynum==6:  # 左。
			suffix = "左"			
		elif entrynum==7:  # 右。
			suffix = "右"	
		elif entrynum==9:  # --なし
			suffix = ""				
		selection.setString("{}{}".format(selection.getString().rstrip("上下左右"), suffix))
	elif entrynum in (10, 11):
		sheet = VARS.sheet
		celladdress = selection.getCellAddress()
		r, c = celladdress.Row, celladdress.Column  # selectionの行インデックスと列インデックスを取得。
		thisc = c - (c-VARS.startcolumn)%8  # 部位の開始列インデックスを取得。	
		ptxt = sheet[0, thisc].getString()  # 部位名文字列を取得。
		datevalue = sheet[r, VARS.daycolumn].getValue()  # 日付シリアル値を取得。	
		datetxt = ""
		if datevalue>0:
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。			
			functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。				
			datetxt = "{}月{}日".format(*[int(functionaccess.callFunction(i, (datevalue,))) for i in ("MONTH", "DAY")])				
		if entrynum==10:  # 開始日にする。これより上行をクリアする
			msg = "部位: {}\n{}より前の点数をクリアします。\n元には戻せません。".format(ptxt, datetxt)
			if showWarningMessageBox(controller, msg)==MessageBoxResults.OK:		
				datarange = sheet[VARS.splittedrow:r, thisc:thisc+8]
				datarange.setPropertyValue("CellBackColor", commons.COLORS["silver"])  # 背景色をつける
				datarange.clearContents(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME)  # セル内容をクリアする。	
		elif entrynum==11:  # 終了日にする。これより下の行をクリア。
			msg = "部位: {}\n{}より後の点数をクリアします。\n元には戻せません。".format(ptxt, datetxt)
			if showWarningMessageBox(controller, msg)==MessageBoxResults.OK:		
				datarange = sheet[r+1:VARS.emptyrow, thisc:thisc+8]
				datarange.setPropertyValue("CellBackColor", commons.COLORS["silver"])  # 背景色をつける	
				datarange.clearContents(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME)  # セル内容をクリアする。	
		clearCellBackColor(thisc)  # 列インデックスcのある部位の値のあるセルの背景色をクリアする。		
	elif entrynum==12:  # この部位を削除
		msg = "この部位({})をすべて削除します。\n元には戻せません。".format(selection.getString())
		if showWarningMessageBox(controller, msg)==MessageBoxResults.OK:		
			sheet = VARS.sheet	
			celladdress = selection.getCellAddress()
			c = celladdress.Column  # selectionの列インデックスを取得。
			VARS.sheet.removeRange(sheet[:, c:c+8].getRangeAddress(), delete_columns)  # 列を削除。	
			splittedrow = VARS.splittedrow
			emptycolumn = VARS.emptycolumn				
			sheet[splittedrow:VARS.emptyrow, VARS.mincolumn].setPropertyValue("CellBackColor", -1)	 # 最低点列の背景色をクリア。	
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
			functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。			
			startdatevalue = int(sheet[splittedrow, VARS.daycolumn].getValue())
			datevalues = [i for i in range(startdatevalue, startdatevalue+VARS.emptyrow-splittedrow)]
			todayvalue = int(functionaccess.callFunction("TODAY", ()))  # 今日のシリアル値を整数で取得。floatで返る。
			if todayvalue in datevalues:  # 今日の行が最終行より下にある時。
				todayrow = splittedrow + datevalues.index(todayvalue)			
				datarows = sheet[VARS.splittedrow:todayrow+1, :emptycolumn].getDataArray()  # 分割行から今日の行までの空列までのデータ行のタプルを取得。
				prevs = datarows[0][:VARS.mincolumn]  # 3月前の最低点のタプルを取得。
				cs = range(VARS.startcolumn+7, emptycolumn, 8)
				mindatarows = [(min([d[i] for i in cs if not d[i]==""], default=""),) for d in datarows]  # 部位別合計列の最低点の行のタプルのリスト。
				highlightPenaltyDays(xscriptcontext.getDocument(), prevs, mindatarows)
				sheet[splittedrow:splittedrow+len(mindatarows), VARS.mincolumn].setDataArray(mindatarows)  # 日の最低点のセル範囲に代入。			
def showWarningMessageBox(controller, msg):	
	componentwindow = controller.ComponentWindow
	msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, WARNINGBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_CANCEL, "myRs", msg)
	return msgbox.execute()
def clearCellBackColor(thisc):  # 部位開始列インデックスを渡してその部位の値のあるセルの背景色をクリアする。
	searchdescriptor = VARS.sheet.createSearchDescriptor()
	searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
	searchdescriptor.setSearchString("[:digit:]+")  # 戻り値はない。数値の入っているセルを検出。
	cellranges = VARS.sheet[VARS.splittedrow:VARS.emptyrow, thisc:thisc+8].findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
	if cellranges:
		cellranges.setPropertyValue("CellBackColor", -1)  # 数値の入っているセルの背景色をクリアする。
def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。	
	selection = None
	for change in changesevent.Changes:
		if change.Accessor=="cell-change":  # セルの値が変化した時。マクロで変更したときはセル範囲が入ってくる時がある。
			selection = change.ReplacedElement  # 値を変更したセルを取得。
			break
	if selection and selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
		celladdress = selection.getCellAddress()
		r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。	
		sheet = VARS.sheet
		if r==VARS.splittedrow and c<VARS.mincolumn:  # 3ヶ月前の最低点入力欄の時。
			prevs = sheet[VARS.splittedrow, :VARS.mincolumn].getDataArray()[0]  # 3ヶ月前の最低点のタプルを取得。
			datarange = sheet[VARS.splittedrow:VARS.emptyrow, VARS.mincolumn]
			datarange.setPropertyValue("CellBackColor", -1)  # 最低点列の背景色のクリア。
			datarows = datarange.getDataArray()  # 月のすべての日の最低点のタプルのタプルを取得。
			doc = xscriptcontext.getDocument()
			highlightPenaltyDays(doc, prevs, datarows)	
		elif VARS.splittedrow<=r<VARS.emptyrow and VARS.startcolumn<=c<VARS.emptycolumn:  # 点数セルの時。
			reCalc(r, c)  # 部位別合計点と日の最低点を計算。
def highlightPenaltyDays(doc, prevs, mindatarows):  # prevs: 3月前の最低点のタプル、mindatarows:  日の最低点の行のタプル。
	if not "" in prevs:  # 空文字がない時。
		if prevs[0]<prevs[1]<prevs[2]:  # 連続して数字が増加している時。
			penaltyrows = [VARS.splittedrow+i for i in range(len(mindatarows)) if mindatarows[i][0] and prevs[2]<mindatarows[i][0]]  # ペナルティ日の行インデックスのリストを取得。datarows[i][0]は空文字の時がある。
			if penaltyrows:  # ペナルティー日がある時。
				dataranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
				dataranges.addRangeAddresses((VARS.sheet[i, VARS.mincolumn].getRangeAddress() for i in penaltyrows), False)
				dataranges.setPropertyValue("CellBackColor", commons.COLORS["magenta3"])	
def reCalc(r, c):  # 部位別合計点と日の最低点を計算。r: 行インデックス、c: 列インデックス。
	sheet = VARS.sheet
	thisc = c - (c-VARS.startcolumn)%8  # 部位の開始列インデックスを取得。
	sheet[r, thisc:thisc+8].setPropertyValue("CellBackColor", -1)  # 同部位同日のセルの背景色をクリアする。
	datarange = sheet[r, :VARS.emptycolumn]  # インデックス0から1行すべてのデータのある範囲を取得。
	datarow = list(datarange.getDataArray()[0])  # リストにしてデータ行を取得。
	datarow[thisc+7] = "" if "" in datarow[thisc:thisc+7] else sum(datarow[thisc+1:thisc+7])  # 部位別合計を代入。「大きさ」は加算しない。部位の点数に空セルがあるときは部位別合計列をクリアする。
	psum = (datarow[i] for i in range(VARS.startcolumn+7, VARS.emptycolumn, 8) if datarow[i])  # 部位別合計のジェネレーター。
	datarow[VARS.mincolumn] = min(psum, default="")  # 部位別合計の日の最低点を代入。部位別合計が一つもなければ空セルにする。
	datarange.setDataArray((datarow,))  # データ行をシートに戻す。
	if datarow[VARS.mincolumn]:  # 日の最低点が空セルでない時。
		prevs = sheet[VARS.splittedrow, :VARS.daycolumn].getDataArray()[0]  # 3月前の最低点のタプルと行の日の最低点を取得。
		if not "" in prevs:  # 3月前の最低点が揃っていない時は何もしない。
			color = commons.COLORS["magenta3"] if prevs[0]<prevs[1]<prevs[2]<datarow[VARS.mincolumn] else -1  # ペナルティの時は最低点セルに背景色を付ける、そうでないなら背景色をクリアする。
			sheet[r, VARS.mincolumn].setPropertyValue("CellBackColor", color)
		