#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 振替伝票シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from . import commons, datedialog, dialogcommons, historydialog
import unohelper, os
from itertools import chain, compress, count, zip_longest
from datetime import datetime
from com.sun.star.awt import MouseButton, MessageBoxButtons, MessageBoxResults  # 定数
from com.sun.star.awt.MessageBoxType import QUERYBOX  # enum
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.table import BorderLine2, TableBorder2 # Struct
from com.sun.star.table import BorderLineStyle, CellVertJustify2  # 定数
from com.sun.star.table.CellHoriJustify import CENTER, LEFT  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
class Journal():  # シート固有の値。
	def __init__(self):
		self.kamokurow = 2  # 科目行インデックス。
		self.hojokamokurow = 3  # 補助科目行インデックス。
		self.subtotalrow = 4  # 科目毎計行インデックス。
		self.splittedrow = 5  # 固定行インデックス。
		self.sliptotalcolumn = 0  # 伝票内計列インデックス。
		self.slipnocolumn = 1  # 伝票番号列インデックス。
		self.daycolumn = 2  # 取引日列インデックス。
		self.tekiyocolumn = 3  # 摘要列インデックス。
		self.splittedcolumn = 4  # 固定列インデックス。
	def setSheet(self, sheet):  # 逐次変化する値。
		self.sheet = sheet
		cellranges = sheet[self.splittedrow:, self.slipnocolumn].queryContentCells(CellFlags.VALUE)  # 伝票番号列の日付列が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 伝票番号列の最終行インデックス+1を取得。
		columnedges = []
		cellranges = sheet[self.kamokurow, self.splittedcolumn:].queryContentCells(CellFlags.STRING) 
		if len(cellranges):
			columnedges.append(cellranges.getRangeAddresses()[-1].EndColumn+1)  # 科目行の右端+1インデックスを取得。
		cellranges = sheet[self.hojokamokurow, self.splittedcolumn:].queryContentCells(CellFlags.STRING) 
		if len(cellranges):
			columnedges.append(cellranges.getRangeAddresses()[-1].EndColumn+1)  # 補助科目行の右端+1インデックスを取得。
		self.emptycolumn = max(columnedges, default=self.splittedcolumn)  # 科目行または補助科目行の右端空列を取得。
VARS = Journal()
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。
	initSheet(activationevent.ActiveSheet, xscriptcontext)
def initSheet(sheet, xscriptcontext):	
	sheet["A1:A3"].setDataArray((("仕訳日記帳生成",), ("総勘定元帳生成",), ("全補助元帳生成",)))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	sheet["D1"].setDataArray((("次年度繰越",),))
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左クリックの時。
		selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==2:  # 左ダブルクリックの時。まずselectionChanged()が発火している。
				celladdress = selection.getCellAddress()
				r, c = celladdress.Row, celladdress.Column  # selectionの行と列インデックスを取得。	
				if r<VARS.splittedrow and c<VARS.splittedcolumn:
					txt = selection.getString()
					if txt=="仕訳日記帳生成":
						sortSlips(xscriptcontext)  # 伝票を日付順にソート。
						createJournalDayBook(xscriptcontext)
					elif txt=="総勘定元帳生成":
						newkingakucolumns = 3, 4, 5  # 金額書式にする列インデックスのタプル。
						newheadermergecolumns = 2, 4, 5  # セル結合するヘッダー行の列インデックスのタプル。
						newtekiyocolumn = 2  # 摘要列インデックス。
						newkamokucolumnidxes = 1,  # 科目列インデックスのタプル。
						newdatewidth = 1500  # 日付列幅。1/100mm。
						newkamokuwidth = 3500  # 科目列幅。
						newkingakuwidth = 2500  # 科目金額列幅。	
						splittedrow = VARS.splittedrow	
						slipnocolumn = VARS.slipnocolumn
						daycolumn = VARS.daycolumn	
						tekiyocolumn = VARS.tekiyocolumn		
						splittedcolumn = VARS.splittedcolumn		
						sheet = VARS.sheet
						ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
						smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
						doc = xscriptcontext.getDocument()
						controller = doc.getCurrentController()	
						indicator = controller.getStatusIndicator() 
						indicator.start("{}中".format(txt), 0)  # 新規ドキュメントを作成後はステータスバーを表示できない。
						controller.select(sheet[splittedrow:, :])  # ソートするセル範囲を選択。
						props = PropertyValue(Name="Col1", Value=daycolumn+1),  # Col1の番号は優先順位。Valueはインデックス+1。 
						dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
						dispatcher.executeDispatch(controller.getFrame(), ".uno:DataSort", "", 0, props)  # ディスパッチコマンドでソート。
						controller.select(selection)  # 元のセルを選択し直す。									
						datarows = sheet[:VARS.emptyrow, :VARS.emptycolumn].getDataArray()  # 全データ行を取得。
						for i in filter(lambda x: x, next(zip(*datarows[splittedrow:]))):  # 伝票内計が0から空セル以外の値をイテレート。
							commons.showErrorMessageBox(doc.getCurrentController(), "貸方と借方が一致しない行があるので\n処理を中止します。")	
							return
						if not datarows[VARS.kamokurow][splittedcolumn]:  # 科目行先頭列のセルがTrueでない時。
							commons.showErrorMessageBox(doc.getCurrentController(), "科目行の先頭セルには科目名が入っていないといけません。")	
							return							
						kamokus = []
						[kamokus.append(i if i else kamokus[-1]) for i in datarows[VARS.kamokurow][splittedcolumn:]]  # 科目行をすべて埋める。
						headerrows = range(splittedcolumn, VARS.emptycolumn), kamokus, datarows[VARS.hojokamokurow][splittedcolumn:]  # 列インデックス行, 科目行、補助科目行。
						newdoc = xscriptcontext.getDesktop().loadComponentFromURL("private:factory/scalc", "_blank", 0, ())  # 新規ドキュメントの取得。
						createFormatKey = commons.formatkeyCreator(newdoc)
						newsheets = newdoc.getSheets()  # 新規ドキュメントのシートコレクションを取得。
						borderline = BorderLine2(LineWidth=10, Color=commons.COLORS["black"])
						tableborder2 = TableBorder2(TopLine=borderline, LeftLine=borderline, RightLine=borderline, BottomLine=borderline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)	
						width, leftmargin, rightmargin = newdoc.getStyleFamilies()["PageStyles"]["Default"].getPropertyValues(("Width", "LeftMargin", "RightMargin"))
						pagewidth = width - leftmargin - rightmargin  # 印刷幅を1/100mmで取得。	
						for kozakamokuname in compress(*(datarows[VARS.kamokurow][splittedcolumn:],)*2):  # 口座科目名をイテレート。科目行の空セルでない値のみイテレート。
							kozacolumns = []  # 口座科目の列インデックスのリスト。
							i = 0
							while kozakamokuname in headerrows[1][i:]:
								i = headerrows[1].index(kozakamokuname, i)
								kozacolumns.append(headerrows[0][i])  # 補助科目の列インデックスを取得。
								i += 1
							newdatarows = [(kozakamokuname, "", "", "", "", ""),\
										(datarows[splittedrow][daycolumn], "", "", "", "", ""),\
										("日付", "相手勘定科目", "摘要", "補助科目", "貸方金額", "残高"),\
										("伝票番号", "相手補助科目", "", "借方金額", "", "")]  # 新規シートのヘッダー行。
							slipstartrows = []  # 新規シートの伝票開始行インデックスのリスト。
							datevalue = ""  # 伝票の日付シリアル値。
							zandaka = 0						
							for i, datarow in enumerate(datarows[splittedrow:], start=splittedrow):  # 伝票行を行インデックスと共にイテレート。
								if any([datarow[j] for j in kozacolumns]):  # 口座科目の列に値がある時のみ。
									slipstartrows.append(len(newdatarows))  # 新規シートの伝票開始行インデックスを取得。
									datevalue = "" if datevalue==datarow[daycolumn] else datarow[daycolumn]  # 前の伝票と日付が異なる時のみ日付を表示する。
									daycolumns = [datevalue, datarow[slipnocolumn]]  # 新規シートの日付列のデータのリスト。伝票の開始行に日付、その下行に伝票番号を表示。
									aitekamokus = []  # 相手科目列の行データリスト。。相手勘定科目と相手補助科目のペア。
									tekiyos = []  # この伝票の摘要列の行データリスト。
									karikatas = []  # この伝票の借方金額列の行データリスト。補助科目と借方金額のペア。
									kashikatas = []  # この伝票の貸方金額列の行データのリスト。
									zandakas = []  # この伝票の残高列の行のリスト。
									for j in compress(zip(*headerrows, datarow[splittedcolumn:]), datarow[splittedcolumn:]):  # 空文字や0でないセルが入っている列の行データを列インデックスとヘッダー行と共にイテレート。
										if j[1]==kozakamokuname:  # 科目名が口座科目名の時。
											annotation = sheet[i, j[0]].getAnnotation().getString().strip()  # 伝票行のこの列のセルのコメントを取得。空白文字を削除する。
											tekiyos = (annotation, "") if annotation else (datarow[tekiyocolumn], "")  # コメントがあるときはそれを摘要にする。
											if j[3]>0:  # 金額が正は借方。補助科目名も使う。
												karikatas.extend([j[2], j[3]])	
											else:  # 金額が負は貸方。
												kashikatas.extend(["", -j[3]])					
											zandaka += j[3]  # 残高を計算。
											zandakas = "", zandaka
										else:  # 口座科目でない時。
											if not aitekamokus:  # 相手科目がまだ未設定の時。
												aitekamokus = j[1:3]  # 相手勘定科目と相手補助科目を取得。					
											elif "諸口" not in aitekamokus: 
												aitekamokus = "諸口", ""  # 相手科目が複数の時の相手科目は諸口にする。
									newdatarows.extend(zip_longest(daycolumns, aitekamokus, tekiyos, karikatas, kashikatas, zandakas, fillvalue=""))  # 各列を1要素ずつイテレートして1行にする。	
							slipstartrows.append(len(newdatarows))  # 新規シートのデータ終了行の下行インデックスを取得。		
							if slipstartrows[0]==slipstartrows[-1]:  # 伝票がない時は何もしない。
								continue
							newsheets.insertNewByName(kozakamokuname, len(newsheets))  # 口座科目名のシートを新規ドキュメントに挿入。
							newsheet = newsheets[kozakamokuname]  # 新規シートを取得。
							newsheet[:len(newdatarows), :len(newdatarows[0])].setDataArray(newdatarows)  # 新規シートに代入。		
							columncount = len(newdatarows[0])  # 表の列数。	
							newsheet[0, :columncount].merge(True)  # 題名セルと結合。			
							newsheet[0, 0].setPropertyValue("HoriJustify", CENTER)	
							newsheet[1, 0].setPropertyValues(("NumberFormat", "HoriJustify"), (createFormatKey("YYYY年"), LEFT))  # 年表示セルのプロパティを設定。
							newsheet[1, :2].merge(True)  # 年表示セルを右横のセルと結合。
							cellranges = newdoc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
							cellranges.addRangeAddresses((newsheet[i, 0].getRangeAddress() for i in slipstartrows[:-1]), False)  # 伝票開始列の日付列セルのセル範囲コレクションを取得。  
							cellranges.setPropertyValues(("HoriJustify", "NumberFormat"), (LEFT, createFormatKey("M/D")))  # 日付書式設定。
							cellranges = newdoc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
							cellranges.addRangeAddresses((newsheet[j, i].getRangeAddress() for i in newkingakucolumns for j in range(slipstartrows[0]+1, slipstartrows[-1], 2)), False)  # 金額セルのセル範囲コレクションを取得。
							cellranges.setPropertyValue("NumberFormat", createFormatKey("#,##0"))  # 金額列の書式設定。
							for i in newheadermergecolumns:  # セル結合するヘッダー行。
								newsheet[2:4, i].merge(True)
								newsheet[2, i].setPropertyValue("VertJustify", CellVertJustify2.CENTER)
							rangeaddresses = []  # 摘要セルのセルアドレスを入れるリスト。
							for i in range(slipstartrows[0], slipstartrows[-1], 2):  # 1行おきに行インデックスをイテレート。
								newsheet[i:i+2, newtekiyocolumn].merge(True)  # 摘要列を2行ずつ結合。
								rangeaddresses.append(newsheet[i, newtekiyocolumn].getRangeAddress())
							cellranges = newdoc.createInstance("com.sun.star.sheet.SheetCellRanges") 
							cellranges.addRangeAddresses(rangeaddresses, False)				
							cellranges.setPropertyValues(("VertJustify", "IsTextWrapped"), (CellVertJustify2.CENTER, True))  # 摘要列を上下中央に、折り返し有効。
							newsheet[slipstartrows[0]:slipstartrows[-1], :].getRows().setPropertyValue("OptimalHeight", True)  # 内容を折り返した後の行の高さを調整。
							cellranges = newdoc.createInstance("com.sun.star.sheet.SheetCellRanges")  
							cellranges.addRangeAddresses((newsheet[i:i+2, j].getRangeAddress() for i in range(slipstartrows[0], slipstartrows[-1], 2) for j in range(1, columncount)), False)  # 1行目だけなぜか真ん中の罫線が引かれてしまう。 						
							cellranges.setPropertyValue("TableBorder2", tableborder2)  
							cellranges = newdoc.createInstance("com.sun.star.sheet.SheetCellRanges")  
							cellranges.addRangeAddresses((newsheet[2:4, j].getRangeAddress() for j in range(columncount)), False)  # 表の1行目の罫線を引く。				
							cellranges.setPropertyValue("TableBorder2", tableborder2)  						
							cellranges = newdoc.createInstance("com.sun.star.sheet.SheetCellRanges")  
							cellranges.addRangeAddresses((newsheet[i:j, 0].getRangeAddress() for i, j in zip(slipstartrows[:-1], slipstartrows[1:])), False)  # 1列目の罫線を伝票区切りで引く。				
							cellranges.setPropertyValue("TableBorder2", tableborder2)  		
							columns = newsheet.getColumns()  # 新規シートの列アクセスオブジェクト。
							for i, j in chain(zip(newkamokucolumnidxes, (newkamokuwidth,)*len(newkamokucolumnidxes)), zip(newkingakucolumns, (newkingakuwidth,)*len(newkingakucolumns))):
								columns[i].setPropertyValue("Width", j)  # 列幅を設定。
							columns[0].setPropertyValue("Width", newdatewidth)  # 日付列幅を設定。
							columns[newtekiyocolumn].setPropertyValue("Width", pagewidth-newdatewidth-newkamokuwidth*len(newkamokucolumnidxes)-newkingakuwidth*len(newkingakucolumns))  # 摘要列幅を設定。残った幅をすべて割り当てる。	
						indicator.end()  # reset()の前にend()しておかないと元に戻らない。
						indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。							
						newdocname = "総勘定元帳_{}.ods".format(datetime.now().strftime("%Y%m%d%H%M%S"))
						del newsheets["Sheet1"]  # 新規ドキュメントのデフォルトシートを削除する。 
						dirpath = os.path.dirname(unohelper.fileUrlToSystemPath(doc.getURL()))  # このドキュメントのあるディレクトリのフルパスを取得。
						systempath = os.path.join(dirpath, "帳簿", newdocname)  # 新規ドキュメントのフルパスを取得。
						if os.path.exists(systempath):  # すでにファイルが存在する時。
							msg = "{}はすでに存在します。\n上書きしますか？".format(newdocname)
							componentwindow = doc.getCurrentController().ComponentWindow
							msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "WEntryBook", msg)
							if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
								return
						newdoc.getStyleFamilies()["PageStyles"]["Default"].setPropertyValue("HeaderIsOn", False)  # 印刷時ヘッダーを付けない。
						newdoc.storeAsURL(unohelper.systemPathToFileUrl(systempath), ())  # 新規ドキュメントを保存。				
					elif txt=="全補助元帳生成":
						sortSlips(xscriptcontext)  # 伝票を日付順にソート。
						doc = xscriptcontext.getDocument()
						newdoc = xscriptcontext.getDesktop().loadComponentFromURL("private:factory/scalc", "_blank", 0, ())  # 新規ドキュメントの取得。
						newsheets = newdoc.getSheets()  # 新規ドキュメントのシートコレクションを取得。
						for i in range(VARS.splittedcolumn, VARS.emptycolumn):
							newsheetname = generateSubsidiaryLedger(xscriptcontext, i)
							if newsheetname:						
								newsheets.importSheet(doc, newsheetname, 0)  # 新規ドキュメントにシートをコピー。
								del doc.getSheets()[newsheetname]  # 切り出したカルテシートを削除する。 
						del newsheets["Sheet1"]  # 新規ドキュメントのデフォルトシートを削除する。 
						dirpath = os.path.dirname(unohelper.fileUrlToSystemPath(doc.getURL()))  # このドキュメントのあるディレクトリのフルパスを取得。
						newdocname = "{}_{}.ods".format("全補助元帳", datetime.now().strftime("%Y%m%d%H%M%S"))
						systempath = os.path.join(dirpath, "帳簿", newdocname)
						if os.path.exists(systempath):  # すでにファイルが存在する時。
							msg = "{}はすでに存在します。\n上書きしますか？".format(newdocname)
							componentwindow = doc.getCurrentController().ComponentWindow
							msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "WEntryBook", msg)
							if msgbox.execute()!=MessageBoxResults.YES:			
								return
						fileurl = unohelper.systemPathToFileUrl(systempath)
						newdoc.storeAsURL(fileurl, ()) 						
					elif txt=="次年度繰越":
						
						
						pass
					
					return False  # セル編集モードにしない。
				elif r>=VARS.splittedrow and c==VARS.daycolumn:  # 取引日列インデックスの時。
					datedialog.createDialog(enhancedmouseevent, xscriptcontext, "取引日", "YYYY-MM-DD")	
					return False  # セル編集モードにしない。
	return True  # セル編集モードにする。シングルクリックは必ずTrueを返さないといけない。
def sortSlips(xscriptcontext):  # 伝票を日付順にソート。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	controller = xscriptcontext.getDocument().getCurrentController()	
	controller.select(VARS.sheet[VARS.splittedrow:, :])  # ソートするセル範囲を選択。
	props = PropertyValue(Name="Col1", Value=VARS.daycolumn+1),  # Col1の番号は優先順位。Valueはインデックス+1。 
	dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
	dispatcher.executeDispatch(controller.getFrame(), ".uno:DataSort", "", 0, props)  # ディスパッチコマンドでソート。
	controller.select(VARS.sheet["A1"])  # A1セルだけ選択にしておく。
def generateSubsidiaryLedger(xscriptcontext, kamokuidx):  # 補助元帳作成。
	kamokucolumnidxes = 1,  # 科目列インデックスのタプル。
	kingakucolumnidxes = 3, 4, 5  # 金額の列インデックスのタプル。
	tekiyocolumnidxes = 2  # 摘要列インデックス。		
	datewidth = 1500  # 日付列幅。1/100mm。
	kamokuwidth = 3500  # 科目列幅。
	kingakuwidth = 2500  # 科目金額列幅。		
	sheet = VARS.sheet
	doc = xscriptcontext.getDocument()
	if not verifySlips(doc):
		return  # すべての伝票行の借方と貸方が一致していることを確認する。
	datarows = sheet[:VARS.emptyrow, :VARS.emptycolumn].getDataArray()  # データ範囲をすべて取得。
	headerrows = createHeaderRows(datarows)  # 列インデックス行, 科目行、補助科目行。
	newsheetname = "_".join([i for i in (headerrows[1][kamokuidx-VARS.splittedcolumn], headerrows[2][kamokuidx-VARS.splittedcolumn]) if i])
	newdatarows = [(datarows[VARS.splittedrow][VARS.daycolumn], "", "", "", "", ""),\
				("日付", "相手勘定科目", "摘要", "借方金額", "貸方金額", "残高"),\
				("伝票番号", "相手補助科目", "", "", "", "")]
	newsliprows = []
	slipstartrows = []  # 伝票開始行インデックスのリスト。
	createDataColumns2 = createDataColumns2Creator(slipstartrows, datarows, headerrows, kamokuidx)
	for i in range(VARS.splittedrow, VARS.emptyrow):  # 伝票行インデックスをイテレート。
		if datarows[i][kamokuidx]:  # 指定科目のセルが空セルや0でない時のみ。
			for k in zip_longest(*createDataColumns2(newdatarows, i), fillvalue=""):  # 各列を1要素ずつイテレートして1行にする。			
				newsliprows.append(k)
	if newsliprows:  # データ行が取得できた時のみ。取得できないのはその科目に伝票がない時。
		newdatarows.extend(newsliprows)			
		slipstartrows.append(len(newdatarows))  # 最終行下の行インデックスを取得。
		newsheet = createNewSheet(doc, newsheetname, newdatarows, slipstartrows, kamokucolumnidxes, kingakucolumnidxes, tekiyocolumnidxes)	
		if newsheet:
			columns = newsheet.getColumns()  # 列アクセスオブジェクト。
			for i, j in chain(zip(kamokucolumnidxes, (kamokuwidth,)*len(kamokucolumnidxes)), zip(kingakucolumnidxes, (kingakuwidth,)*len(kingakucolumnidxes))):
				columns[i].setPropertyValue("Width", j)  # 列幅を設定。
			width, leftmargin, rightmargin = doc.getStyleFamilies()["PageStyles"]["Default"].getPropertyValues(("Width", "LeftMargin", "RightMargin"))
			pagewidth = width - leftmargin - rightmargin  # 印刷幅を1/100mmで取得。
			columns[0].setPropertyValue("Width", datewidth)  # 日付列幅を設定。
			columns[tekiyocolumnidxes].setPropertyValue("Width", pagewidth-datewidth-kamokuwidth*len(kamokucolumnidxes)-kingakuwidth*len(kingakucolumnidxes))  # 摘要列幅を設定。残った幅をすべて割り当てる。
			return newsheetname
def createDataColumns2Creator(slipstartrows, datarows, headerrows, kamokuidx):
	datevalue = ""
	zandaka = 0
	def createDataColumns2(newdatarows, i):
		nonlocal datevalue, zandaka
		slipstartrows.append(len(newdatarows))  # 伝票開始行インデックスを取得。
		datarow = datarows[i]
		datevalue = "" if datevalue==datarow[VARS.daycolumn] else datarow[VARS.daycolumn]  # 前の伝票と日付が異なる時のみ日付を表示する。
		daycolumns = [datevalue, datarow[VARS.slipnocolumn]]  # 日付列のデータのリスト。
		aitekamokus = []
		karikatas = []
		kashikatas = []
		for j in compress(zip(*headerrows, datarow[VARS.splittedcolumn:]), datarow[VARS.splittedcolumn:]):  # 空文字や0でないセルが入っている列の要素のみイテレート。
			if j[0]==kamokuidx:  # 口座科目の時。
				annotation = VARS.sheet[i, j[0]].getAnnotation().getString()
				if annotation:
					tekiyos = annotation, ""
				else:
					tekiyos = datarow[VARS.tekiyocolumn], ""
				if j[3]>0:  # 金額が正は借方。
					karikatas.extend([j[3], ""])	
				else:  # 金額が負は貸方。
					kashikatas.extend([-j[3], ""])					
				zandaka += j[3]
				zandakas = zandaka, ""
			else:  # 口座科目でない時。
				if not aitekamokus:  # 相手科目がまだ未設定の時。
					aitekamokus = j[1:3]					
				elif "諸口" not in aitekamokus: 
					aitekamokus = "諸口", ""  # 相手科目が複数の時の相手科目は諸口にする。
		return daycolumns, aitekamokus, tekiyos, karikatas, kashikatas, zandakas
	return createDataColumns2
def createJournalDayBook(xscriptcontext):
	newsheetname = "仕訳日記帳"
	kamokucolumnidxes = 1, 3  # 科目列インデックスのタプル。
	kingakucolumnidxes = 2, 4  # 金額の列インデックスのタプル。
	tekiyocolumnidxes = 5  # 摘要列インデックス。	
	datewidth = 1500  # 日付列幅。1/100mm。
	kamokuwidth = 3500  # 科目列幅。
	kingakuwidth = 2500  # 科目金額列幅。	
	sheet = VARS.sheet
	doc = xscriptcontext.getDocument()
	if not verifySlips(doc):
		return  # すべての伝票行の借方と貸方が一致していることを確認する。
	datarows = sheet[:VARS.emptyrow, :VARS.emptycolumn].getDataArray()  # データ範囲をすべて取得。
	headerrows = createHeaderRows(datarows)  # 列インデックス行, 科目行、補助科目行、のタプルを取得。
	newdatarows = [(datarows[VARS.splittedrow][VARS.daycolumn], "", "", "", "", ""),\
				("日付", "借方科目", "借方金額", "貸方科目", "貸方金額", "摘要"),\
				("伝票番号", "借方補助科目", "", "貸方補助科目", "", "")]
	slipstartrows = []  # 伝票開始行インデックスのリスト。
	createDataColumns = createDataColumnsCreator(slipstartrows, datarows, headerrows)
	for i in range(VARS.splittedrow, VARS.emptyrow):  # 伝票行インデックスをイテレート。
		for k in zip_longest(*createDataColumns(newdatarows, i), fillvalue=""):  # 各列を1要素ずつイテレートして1行にする。			
			newdatarows.append([*k[:-3], "/".join([m for m in k[-3:] if m])])  # 摘要は/で結合する。
	slipstartrows.append(len(newdatarows))  # 最終行下の行インデックスを取得。
	newsheet = createNewSheet(doc, newsheetname, newdatarows, slipstartrows, kamokucolumnidxes, kingakucolumnidxes, tekiyocolumnidxes)	
	if newsheet:
		columns = newsheet.getColumns()  # 列アクセスオブジェクト。
		for i, j in chain(zip(kamokucolumnidxes, (kamokuwidth,)*len(kamokucolumnidxes)), zip(kingakucolumnidxes, (kingakuwidth,)*len(kingakucolumnidxes))):
			columns[i].setPropertyValue("Width", j)  # 列幅を設定。
		width, leftmargin, rightmargin = doc.getStyleFamilies()["PageStyles"]["Default"].getPropertyValues(("Width", "LeftMargin", "RightMargin"))
		pagewidth = width - leftmargin - rightmargin  # 印刷幅を1/100mmで取得。
		columns[0].setPropertyValue("Width", datewidth)  # 日付列幅を設定。
		columns[tekiyocolumnidxes].setPropertyValue("Width", pagewidth-datewidth-2*(kamokuwidth+kingakuwidth))  # 摘要列幅を設定。残った幅をすべて割り当てる。
		detachSheet(xscriptcontext, newsheetname, "{}_{}.ods".format(newsheetname, datetime.now().strftime("%Y%m%d%H%M%S")))  # シートをファイルに切り出す。
def createNewSheet(doc, newsheetname, newdatarows, slipstartrows, kamokucolumnidxes, kingakucolumnidxes, tekiyocolumnidxes):	
	columncount = len(newdatarows[0])
	sheets = doc.getSheets()
	if newsheetname in sheets:  # すでに同名シートがある時は削除する。
		msg = "{}シートはすでに存在します。\n上書きしますか？".format(newsheetname)
		componentwindow = doc.getCurrentController().ComponentWindow
		msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "WEntryBook", msg)
		if msgbox.execute()==MessageBoxResults.YES:				
			del sheets[newsheetname]
		else:
			return
	sheets.insertNewByName(newsheetname, len(sheets))
	newsheet = sheets[newsheetname]
	newsheet[:len(newdatarows), :len(newdatarows[0])].setDataArray(newdatarows)
	createFormatKey = commons.formatkeyCreator(doc)
	newsheet[0, 0].setPropertyValue("NumberFormat", createFormatKey("YYYY年"))  # 年表示セル。
	newsheet[0, :2].merge(True)  # 年表示セル。
	newsheet[0, 0].setPropertyValue("HoriJustify", LEFT)  # 年表示セルを左寄せ。	
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
	cellranges.addRangeAddresses((newsheet[i, 0].getRangeAddress() for i in slipstartrows), False)  
	cellranges.setPropertyValues(("HoriJustify", "NumberFormat"), (LEFT, createFormatKey("M/D")))  # 日付書式設定。
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
	cellranges.addRangeAddresses((newsheet[slipstartrows[0]:len(newdatarows), i].getRangeAddress() for i in kingakucolumnidxes), False)  # 金額列の書式設定。
	cellranges.setPropertyValue("NumberFormat", createFormatKey("#,##0"))
	for i in kingakucolumnidxes:  # 貸方金額と借方金額のヘッダー行。
		newsheet[1:3, i].merge(True)
		newsheet[1, i].setPropertyValue("VertJustify", CellVertJustify2.CENTER)
	rangeaddresses = []
	for i in range(1, len(newdatarows), 2):
		newsheet[i:i+2, tekiyocolumnidxes].merge(True)  # 摘要列について。
		rangeaddresses.append(newsheet[i, tekiyocolumnidxes].getRangeAddress())
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges") 
	cellranges.addRangeAddresses(rangeaddresses, False)				
	cellranges.setPropertyValues(("VertJustify", "IsTextWrapped"), (CellVertJustify2.CENTER, True))  # 摘要列を上下中央に、折り返し有効。
	newsheet[2:, tekiyocolumnidxes].getRows().setPropertyValue("OptimalHeight", True)  # 内容を折り返した後の行の高さを調整。
	borderline = BorderLine2(LineWidth=10, Color=commons.COLORS["black"])
	noneline = BorderLine2(LineStyle=BorderLineStyle.NONE)
	tableborder2 = TableBorder2(TopLine=borderline, LeftLine=borderline, RightLine=borderline, BottomLine=borderline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)
	nonetableborder2 = TableBorder2(TopLine=noneline, LeftLine=noneline, RightLine=noneline, BottomLine=noneline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  
	cellranges.addRangeAddresses((newsheet[i:i+2, j].getRangeAddress() for i in range(3, len(newdatarows), 2) for j in range(1, columncount)), False)  # 1行目だけなぜか真ん中の罫線が引かれてしまう。 						
	cellranges.setPropertyValue("TableBorder2", tableborder2)  
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges") 
	cellranges.addRangeAddresses((newsheet[i, j].getRangeAddress() for i in (3, 4) for j in kamokucolumnidxes), False)  # 1行目の勝手に引かれた罫線を消す。					
	cellranges.setPropertyValue("TableBorder2", nonetableborder2)  
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  
	cellranges.addRangeAddresses((newsheet[1:3, j].getRangeAddress() for j in range(columncount)), False)  # 1行目の罫線を引き直す。				
	cellranges.setPropertyValue("TableBorder2", tableborder2)  						
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  
	cellranges.addRangeAddresses((newsheet[i:j, 0].getRangeAddress() for i, j in zip(slipstartrows[:-1], slipstartrows[1:])), False)  # 1列目の罫線を伝票区切りで引く。				
	cellranges.setPropertyValue("TableBorder2", tableborder2)  		
	return newsheet
def createDataColumnsCreator(slipstartrows, datarows, headerrows):
	datevalue = ""
	def createDataColumns(newdatarows, i):
		nonlocal datevalue
		slipstartrows.append(len(newdatarows))  # 伝票開始行インデックスを取得。
		datarow = datarows[i]
		datevalue = "" if datevalue==datarow[VARS.daycolumn] else datarow[VARS.daycolumn]  # 前の伝票と日付が異なる時のみ日付を表示する。
		daycolumns = [datevalue, datarow[VARS.slipnocolumn]]  # 日付列のデータのリスト。
		karikatakamokus = []  # 借方科目列のデータのリスト。
		karikatas = []  # 借方金額列のデータのリスト。		
		karikatatekiyo = []  # 借方摘要列のデータのリスト。				
		kashikatakamokus = []  # 貸方科目列のデータのリスト。		
		kashikatas = []  # 貸方金額列のデータのリスト。		
		kashikatatekiyo = []  # 貸方摘要列のデータのリスト。		
		for j in compress(zip(*headerrows, datarow[VARS.splittedcolumn:]), datarow[VARS.splittedcolumn:]):  # 空文字や0でないセルが入っている列の要素のみイテレート。
			annotation = VARS.sheet[i, j[0]].getAnnotation().getString()
			if j[3]>0:  # 金額が正の科目は借方。
				karikatakamokus.extend(j[1:3])
				karikatas.extend([j[3], ""])	
				karikatatekiyo.extend([annotation, ""])		
			else:  # 金額が負の科目は貸方。
				kashikatakamokus.extend(j[1:3])
				kashikatas.extend([-j[3], ""])
				kashikatatekiyo.extend([annotation, ""])	
		return daycolumns, karikatakamokus, karikatas, kashikatakamokus, kashikatas, [datarow[VARS.tekiyocolumn]], karikatatekiyo, kashikatatekiyo
	return createDataColumns
def createHeaderRows(datarows):  # 科目行の空セルをすべて埋めたあと、列インデックス行, 科目行、補助科目行、を返す。
	kamokus = []
	buf = ""
	for i in datarows[VARS.kamokurow][VARS.splittedcolumn:]:
		if i:
			buf = i
		kamokus.append(buf)  # 科目行をすべて埋める。
	return range(VARS.splittedcolumn, VARS.emptycolumn), kamokus, datarows[VARS.hojokamokurow][VARS.splittedcolumn:]  # 列インデックス行, 科目行、補助科目行。
def verifySlips(doc):  # すべての伝票行の借方と貸方が一致していることを確認する。
	searchdescriptor = VARS.sheet.createSearchDescriptor()
	searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
	searchdescriptor.setSearchString("[^0]")  # 0以外のセルを取得。戻り値はない。	
	cellranges = VARS.sheet[VARS.splittedrow:VARS.emptyrow, VARS.sliptotalcolumn].queryContentCells(CellFlags.VALUE).findAll(searchdescriptor)  # 値のあるセルから0以外が入っているセル範囲コレクションを取得。見つからなかった時はNoneが返る。
	if cellranges:
		commons.showErrorMessageBox(doc.getCurrentController(), "貸方と借方が一致しない行があります。")	
		return False
	return True
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	selection = eventobject.Source.getSelection()	
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
		sheet = selection.getSpreadsheet()
		VARS.setSheet(sheet)		
		drowBorders(selection)  # 枠線の作成。
def drowBorders(selection):  # ターゲットを交点とする行列全体の外枠線を描く。
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上端のセルアドレスを取得。
	r = celladdress.Row  # selectionの行と列のインデックスを取得。	
	sheet = VARS.sheet
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = commons.createBorders()
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。	
	if r<VARS.splittedrow:  # 分割行より上の時。
		return  # 罫線を引き直さない。
	rangeaddress = selection.getRangeAddress()  # 選択範囲のセル範囲アドレスを取得。
	sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :VARS.emptycolumn].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く
	sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。
	selection.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。		
def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。マクロで変更した時は発火しない。	
	selection = None
	for change in changesevent.Changes:
		if change.Accessor=="cell-change":  # セルの値が変化した時。
			selection = change.ReplacedElement  # 値を変更したセルを取得。	
			break
		
	# セルに式を入れるとループする。		
		
	if selection:  # セルとは限らずセル範囲のときもある。シートからペーストしたときなど。テキストをペーストした時は発火しない。
		sheet = VARS.sheet
		cellranges = sheet[VARS.splittedrow:, :VARS.emptycolumn].queryIntersection(selection.getRangeAddress())  # 固定行以下と科目右列端との選択範囲と重なる部分のセル範囲コレクションを取得。
		if len(cellranges):  # 変化したセル範囲がある時。
			VARS.setSheet(sheet)  # 逐次変化する値を取得。伝票番号列の最終行を再取得したい。
			deadnogene = (j for j in count(1) if j not in list(chain.from_iterable(sheet[VARS.splittedrow:VARS.emptyrow, VARS.slipnocolumn].getDataArray())))
			createFormatKey = commons.formatkeyCreator(xscriptcontext.getDocument())
			for rangeaddress in cellranges.getRangeAddresses():  # セル範囲アドレスをイテレート。
				datarange = sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :VARS.emptycolumn]  # 行毎に処理するセル範囲を取得。
				datarange[:, VARS.daycolumn].setPropertyValue("NumberFormat", createFormatKey("YYYY-MM-DD"))  # 取引日列の書式を設定。
				newdatarows = []  # 処理後の伝票内計列と伝票番号列の行データを取得するリスト。
				for datarow in datarange.getDataArray():  # 各行をイテレート。
					sliptotal = sum(filter(lambda x: isinstance(x, float), datarow[VARS.splittedcolumn:]))  # 行の合計を取得。
					slipno = datarow[VARS.slipnocolumn] or next(deadnogene)  # 伝票番号を取得。
					newdatarows.append((sliptotal, slipno))
				datarange[:, :VARS.daycolumn].setDataArray(newdatarows)
				VARS.setSheet(sheet)  # 逐次変化する値を取得。伝票番号列の最終行を再取得したい。
				datarange = sheet[VARS.splittedrow:VARS.emptyrow, rangeaddress.StartColumn:rangeaddress.EndColumn+1]  # 列毎に処理するセル範囲を取得。
				sheet[VARS.subtotalrow, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setDataArray(([sum(filter(lambda x: isinstance(x, float), i)) for i in zip(*datarange.getDataArray())],))  # 列ごとの合計を取得。
			highlightDupeNo(xscriptcontext)  # 重複伝票番号セルをハイライトする。
			datarange = sheet[VARS.splittedrow:VARS.emptyrow, VARS.sliptotalcolumn]  # 伝票内計列のセル範囲を取得。
			highlightImBalance(xscriptcontext, datarange)  # 不均衡セルをハイライト。
			sheet[VARS.subtotalrow:VARS.emptyrow, VARS.splittedcolumn:VARS.emptycolumn].setPropertyValue("NumberFormat", createFormatKey("#,##0;[BLUE]-#,##0"))
def highlightDupeNo(xscriptcontext):  # 重複伝票番号セルをハイライトする。
	sheet = VARS.sheet
	splittedrow = VARS.splittedrow
	slipnocolumn = VARS.slipnocolumn
	datarange = sheet[splittedrow:VARS.emptyrow, slipnocolumn]  # 伝票番号列のセル範囲を取得。
	datarange.setPropertyValue("CellBackColor", -1)  # 伝票番号列の背景色をクリア。
	sliprows = datarange.getDataArray()  # 伝票番号列の行のタプルを取得。
	sliprowsset = set(sliprows)  # 重複行を削除した集合を取得。
	duperows = []  # 重複している伝票番号がある行インデックスを取得するリスト。
	if len(sliprows)>len(sliprowsset):  # 伝票番号列に重複行がある時。空文字も重複してはいけない。
		for i in sliprowsset:  # 重複は除いて伝票番号をイテレート。
			if sliprows.count(i)>1:  # 複数ある時。
				j = 0
				while i in sliprows[j:]:
					j = sliprows.index(i, j)
					duperows.append(j+splittedrow)  # 重複している伝票番号がある行インデックスを取得。
					j += 1
		cellranges = xscriptcontext.getDocument().createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
		cellranges.addRangeAddresses([sheet[i, slipnocolumn].getRangeAddress() for i in duperows], False)
		cellranges.setPropertyValue("CellBackColor", commons.COLORS["silver"])  # 重複伝票番号の背景色を変える。	
def highlightImBalance(xscriptcontext, datarange):  # 不均衡セルをハイライト。
	datarange.setPropertyValues(("CellBackColor", "NumberFormat"), (-1, commons.formatkeyCreator(xscriptcontext.getDocument())("#,##0;[BLUE]-#,##0")))  # 背景色をクリア, ついでに書式を設定。
	searchdescriptor = VARS.sheet.createSearchDescriptor()
	searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
	searchdescriptor.setSearchString("[^0]")  # 0以外のセルを取得。戻り値はない。	
	cellranges = datarange.queryContentCells(CellFlags.VALUE).findAll(searchdescriptor)  # 値のあるセルから0以外が入っているセル範囲コレクションを取得。見つからなかった時はNoneが返る。
	if cellranges:
		cellranges.setPropertyValue("CellBackColor", commons.COLORS["violet"])	
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。	
	contextmenuname, addMenuentry, baseurl, selection = commons.contextmenuHelper(VARS, contextmenuexecuteevent, xscriptcontext)
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上角のセルのアドレスを取得。
	r, c  = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。	
	sheet = VARS.sheet
	if contextmenuname=="cell":  # セルのとき。セル範囲も含む。
		if VARS.splittedcolumn<=c<VARS.emptycolumn:  # 科目行か補助科目行に値がある列の時。
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 単独セルの時のみ。
				datarows = sheet[VARS.kamokurow:VARS.hojokamokurow+1, c].getDataArray()  # 科目行と補助科目行を取得。
				kamoku = datarows[0][0]
				hojokamoku = datarows[1][0]
				if r==VARS.kamokurow and kamoku:  # 科目行かつ科目行に値があるとき。
					addMenuentry("ActionTrigger", {"Text": "{}の勘定元帳生成".format(kamoku), "CommandURL": baseurl.format("entry2")}) 
				elif r==VARS.hojokamokurow and hojokamoku:  # 補助科目行かつ補助科目行に値があるとき。:
					addMenuentry("ActionTrigger", {"Text": "{}の補助元帳生成".format(hojokamoku), "CommandURL": baseurl.format("entry3")}) 
				elif VARS.splittedrow<=r<=VARS.emptyrow:  # 取引日列が入力済で科目行か補助科目行に値がある列のセルの時。
					if sheet[r, VARS.sliptotalcolumn].getValue()!=0:  # 伝票内計が0でない時のみ。空セルも0として扱われる。
						txt = hojokamoku if hojokamoku else kamoku  # 補助科目行に値がある時は補助科目行、ないときは科目行の値を使う。
						if txt!="現金":  # 現金列でない時のみ。
							addMenuentry("ActionTrigger", {"Text": "現金で決済", "CommandURL": baseurl.format("entry4")}) 
						addMenuentry("ActionTrigger", {"Text": "{}で決済".format(txt), "CommandURL": baseurl.format("entry5")}) 
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertAnnotation"})	
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:EditAnnotation"})	
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteNote"})	
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:ShowNote"})			
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:HideNote"})							
		elif c==VARS.tekiyocolumn:  # 摘要列の時。
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 単独セルの時のみ。
				addMenuentry("ActionTrigger", {"Text": "伝票履歴", "CommandURL": baseurl.format("entry6")}) 
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			addMenuentry("ActionTrigger", {"Text": "伝票履歴に追加", "CommandURL": baseurl.format("entry7")}) 
		elif c==VARS.slipnocolumn:  # 伝票番号列の時。
			addMenuentry("ActionTrigger", {"Text": "空番号取得", "CommandURL": baseurl.format("entry8")}) 
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"Text": "クリア", "CommandURL": baseurl.format("entry1")}) 	
	elif contextmenuname=="rowheader" and len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 行ヘッダーのとき、かつ、選択範囲の列数がシートの列数が一致している時。	
		if r>=VARS.splittedrow:  # 固定行以下の時のみ。
			commons.cutcopypasteMenuEntries(addMenuentry)
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			commons.rowMenuEntries(addMenuentry)		
	elif contextmenuname=="colheader" and len(selection[:, 0].getRows())==len(sheet[:, 0].getRows()):  # 列ヘッダーの時、かつ、選択範囲の行数がシートの行数が一致している時。	
		if c>=VARS.splittedcolumn:
			commons.cutcopypasteMenuEntries(addMenuentry)
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})		
			commons.columnMenuEntries(addMenuentry)		
			if len(selection.getColumns())>1:  # 複数列を選択している時。
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})		
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:Group"})	
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:Ungroup"})	
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。	
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。		
	controller = xscriptcontext.getDocument().getCurrentController()  # コントローラの取得。
	selection = controller.getSelection()  # 選択範囲を取得。
	sheet = VARS.sheet
	if entrynum==1:  # クリア。書式設定とオブジェクト以外を消去。
		selection.clearContents(511)  # 範囲をすべてクリアする。
	elif entrynum==2:  # 勘定元帳生成
		
		
		
		pass
	elif entrynum==3:  # 補助元帳生成
		sortSlips(xscriptcontext)  # 伝票を日付順にソート。
		newsheetname = generateSubsidiaryLedger(xscriptcontext, selection.getCellAddress().Column)
		if newsheetname:
			detachSheet(xscriptcontext, newsheetname, "{}_{}.ods".format(newsheetname, datetime.now().strftime("%Y%m%d%H%M%S")))  # シートをファイルに切り出す。	
	elif entrynum==4:  # 現金で決済
		datarow = sheet[VARS.kamokurow, :VARS.emptycolumn].getDataArray()[0]
		settle(sheet[selection.getCellAddress().Row, datarow.index("現金", VARS.splittedcolumn)])
	elif entrynum==5:  # 決済
		settle(selection)
	elif entrynum==6:  # 伝票履歴
		historydialog.createDialog(xscriptcontext, "伝票履歴", callback=callback_sliphistoryCreator(xscriptcontext))
	elif entrynum==7:  # 伝票履歴に追加
		newgriddatarows = []  # グリッドコントロールに追加する行のリスト。
		datarows = sheet[:VARS.emptyrow, VARS.tekiyocolumn:VARS.emptycolumn].getDataArray()
		rangeaddress = selection.getRangeAddress()  # 選択範囲のアドレスを取得。
		for i in range(rangeaddress.StartRow, rangeaddress.EndRow+1):  # 行インデックスをイテレート。
			items = [datarows[i][0]]  # 摘要を取得。
			for j, val in enumerate(datarows[i][1:], start=1):  # リストのインデックスと値をイテレート。
				if val!="":  # 空セルでない時。つまり金額が入っている時。
					hojokamoku = datarows[VARS.hojokamokurow][j]  # 補助科目を取得。
					annotation = sheet[i, VARS.tekiyocolumn+j].getAnnotation().getString()  # セルコメントを取得。
					for k in range(j, 0, -1):  # 科目行を左にイテレート。
						kamoku = datarows[VARS.kamokurow][k]  # 科目を取得。
						if kamoku:  # 科目が取得できたらfor文を抜ける。
							break
					items.append("::".join([kamoku, hojokamoku, str(val), annotation]))
			newgriddatarows.append(("//".join(items),))
		doc = xscriptcontext.getDocument()
		dialogtitle = "伝票履歴"
		griddatarows = dialogcommons.getSavedData(doc, "GridDatarows_{}".format(dialogtitle))  # グリッドコントロールの行をconfigシートのragenameから取得する。	
		if griddatarows:  # 行のリストが取得出来た時。
			griddatarows.extend(newgriddatarows)
		else:
			griddatarows = newgriddatarows
		dialogcommons.saveData(doc, "GridDatarows_{}".format(dialogtitle), griddatarows)
	elif entrynum==8:  # 空番号取得。
		deadnogene = (j for j in count(1) if j not in list(chain.from_iterable(sheet[VARS.splittedrow:VARS.emptyrow, VARS.slipnocolumn].getDataArray())))
		selection.setValue(next(deadnogene))  # 空き番号を取得。
def settle(cell):		
	sheet = VARS.sheet
	r = cell.getCellAddress().Row
	sliptotal = sum(filter(lambda x: isinstance(x, float), sheet[r, VARS.splittedcolumn:VARS.emptycolumn].getDataArray()[0]))  # 行の合計を取得。
	cell.setValue(cell.getValue()-sliptotal)
	sheet[r, VARS.sliptotalcolumn].setValue(0)
	sheet[r, VARS.sliptotalcolumn].setPropertyValue("CellBackColor", -1) 		
def callback_sliphistoryCreator(xscriptcontext):		
	def callback_sliphistory(gridcelltxt):
		sheet = VARS.sheet
		headerrows = sheet[VARS.kamokurow:VARS.hojokamokurow+1, :VARS.emptycolumn].getDataArray()  # 科目行と補助科目行を取得。
		controller = xscriptcontext.getDocument().getCurrentController()  # コントローラの取得。
		selection = controller.getSelection()  # 選択範囲を取得。選択範囲はセルのみ。	
		r = selection.getCellAddress().Row
		datarange = sheet[r, :VARS.emptycolumn]  # 代入するセル範囲を取得。
		datarow = list(datarange.getDataArray()[0])  # 選択行をリストで取得。
		items = gridcelltxt.split("//")
		datarow[VARS.tekiyocolumn] = items[0]  # 摘要を取得。
		comments = []  # コメントのセルとコメントのタプルを取得するリスト。
		recalccols = []  # 再計算する列インデックスのリスト。
		for item in items[1:]:
			kamoku, hojokamoku, val, annotation = item.split("::")
			if headerrows[0][VARS.splittedcolumn:].count(kamoku)==1:  # 科目行に該当する科目が１つの時のみ。
				c = headerrows[0].index(kamoku, VARS.splittedcolumn)  # その科目の列インデックスを取得。
				if hojokamoku:  # 補助科目がある時。
					if headerrows[1][c:].count(hojokamoku)==1:  # 補助科目行にその補助科目が１つの時のみ。
						c = headerrows[1].index(hojokamoku, c)  # その補助科目の列インデックスを取得。
					else:
						commons.showErrorMessageBox(controller, "補助科目「{}」の列を同定できません。".format(hojokamoku))
						return	
				datarow[c] = float(val)  # セルに入れる数値。
				recalccols.append(c)
				if annotation:  # コメントがある時。
					comments.append((sheet[r, c], annotation))  # setDataArray()でコメントがクリアされるのでここでセルとコメントの文字列をタプルで取得しておく。
			else:
				commons.showErrorMessageBox(controller, "科目「{}」の列を同定できません。".format(kamoku))
				return
		deadnogene = (j for j in count(1) if j not in list(chain.from_iterable(sheet[VARS.splittedrow:VARS.emptyrow, VARS.slipnocolumn].getDataArray())))
		datarow[VARS.slipnocolumn] = datarow[VARS.slipnocolumn] or next(deadnogene)  # 伝票番号を取得。
		datarow[VARS.sliptotalcolumn] = sum(filter(lambda x: isinstance(x, float), datarow[VARS.splittedcolumn:]))  # 行の合計を取得。	
		datarange.setDataArray((datarow,))
		annotations = sheet.getAnnotations()  # コメントコレクションを取得。
		for i in comments:
			cell, annotation = i
			annotations.insertNew(cell.getCellAddress(), annotation)  # コメントを挿入。
			cell.getAnnotation().getAnnotationShape().setPropertyValue("Visible", False)  # これをしないとmousePressed()のTargetにAnnotationShapeが入ってしまう。		
		VARS.setSheet(sheet)  # 逐次変化する値を取得。伝票番号列の最終行を再取得したい。
		datarows = sheet[VARS.subtotalrow:VARS.emptyrow, min(recalccols):max(recalccols)+1].getDataArray()  # 個別の列だけ再計算するのは面倒なので、連続する列すべてを再計算する。
		sheet[VARS.subtotalrow, min(recalccols):max(recalccols)+1].setDataArray(([sum(filter(lambda x: isinstance(x, float), i)) for i in zip(*datarows[1:])],))  # 列ごとの合計を取得。			
	return callback_sliphistory	
def detachSheet(xscriptcontext, sheetname, newdocname):
	doc = xscriptcontext.getDocument()
	newdoc = xscriptcontext.getDesktop().loadComponentFromURL("private:factory/scalc", "_blank", 0, ())  # 新規ドキュメントの取得。
	newsheets = newdoc.getSheets()  # 新規ドキュメントのシートコレクションを取得。
	newsheets.importSheet(doc, sheetname, 0)  # 新規ドキュメントにシートをコピー。
	del newsheets["Sheet1"]  # 新規ドキュメントのデフォルトシートを削除する。 
	del doc.getSheets()[sheetname]  # 切り出したカルテシートを削除する。 
	dirpath = os.path.dirname(unohelper.fileUrlToSystemPath(doc.getURL()))  # このドキュメントのあるディレクトリのフルパスを取得。
	systempath = os.path.join(dirpath, "帳簿", newdocname)
	if os.path.exists(systempath):  # すでにファイルが存在する時。
		msg = "{}はすでに存在します。\n上書きしますか？".format(newdocname)
		componentwindow = doc.getCurrentController().ComponentWindow
		msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "WEntryBook", msg)
		if msgbox.execute()!=MessageBoxResults.YES:			
			return
	fileurl = unohelper.systemPathToFileUrl(systempath)
	newdoc.storeAsURL(fileurl, ())  
