#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 仕訳日誌シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
# import os, unohelper, glob
# from . import commons, datedialog, points, transientdialog
from . import commons, datedialog, dialogcommons, historydialog
from itertools import chain, compress, count, zip_longest
# from com.sun.star.accessibility import AccessibleRole  # 定数
from com.sun.star.awt import MouseButton  # 定数
# from com.sun.star.awt import MouseButton, MessageBoxButtons, MessageBoxResults, ScrollBarOrientation # 定数
# from com.sun.star.awt.MessageBoxType import INFOBOX, QUERYBOX  # enum
# from com.sun.star.beans import PropertyValue  # Struct
# from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH  # enum
# from com.sun.star.lang import Locale  # Struct
from com.sun.star.sheet import CellFlags  # 定数
# from com.sun.star.sheet.CellDeleteMode import ROWS as delete_rows  # enum
# from com.sun.star.table import BorderLine2  # Struct
# from com.sun.star.table import BorderLineStyle  # 定数
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
	sheet["A1:A3"].setDataArray((("仕訳帳生成",), ("総勘定元帳生成",), ("全補助元帳生成",)))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
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
					if txt=="仕訳帳生成":
						sheet = VARS.sheet
						datarows = sheet[:VARS.emptyrow, :VARS.emptycolumn].getDataArray()
						kamokus = []
						buf = ""
						for i in datarows[VARS.kamokurow][VARS.splittedcolumn:]:
							if i:
								buf = i
							kamokus.append(buf)
						headerrows = range(VARS.splittedcolumn, VARS.emptycolumn), kamokus, datarows[VARS.hojokamokurow][VARS.splittedcolumn:]  # 列インデックス行, 科目行、補助科目行。
						newdatarows = [(datarows[VARS.splittedrow][VARS.daycolumn], "", "", "", "", ""),\
										("日付", "借方科目", "借方金額", "貸方科目", "貸方金額", "摘要"),\
										("伝票番号", "借方補助科目", "", "貸方補助科目", "", "")]		
						datevalue = ""				
						for i in range(VARS.splittedrow, VARS.emptyrow):  # 伝票行インデックスをイテレート。
							datarow = datarows[i]
							datevalue = "" if datevalue==datarow[VARS.daycolumn] else datarow[VARS.daycolumn]
							daycolumns = [datevalue, datarow[VARS.slipnocolumn]]
							karikatakamokus = []
							karikatas = []		
							karikatatekiyo = []					
							kashikatakamokus = []
							kashikatas = []
							kashikatatekiyo = []
							for j in compress(zip(*headerrows, datarow[VARS.splittedcolumn:]), datarow[VARS.splittedcolumn:]):  # 空文字や0でないセルが入っている列の要素のみイテレート。
								annotation = sheet[i, j[0]].getAnnotation().getString()
								if j[3]>0:  # 金額が正の科目は借方。
									karikatakamokus.extend(j[1:3])
									karikatas.extend([j[3], ""])	
									karikatatekiyo.extend([annotation, ""])		
								else:  # 金額が負の科目は貸方。
									kashikatakamokus.extend(j[1:3])
									kashikatas.extend([-j[3], ""])
									kashikatatekiyo.extend([annotation, ""])		
							for k in zip_longest(daycolumns, karikatakamokus, karikatas, kashikatakamokus, kashikatas, [datarow[VARS.tekiyocolumn]], karikatatekiyo, kashikatatekiyo, fillvalue=""):
								newdatarows.append([*k[:-3], "/".join([m for m in k[-3:] if m])])
						

						
						newsheet 
						
						
							

							

							
							
						
						
						

						
						
						
						
						
						pass
					
					elif txt=="総勘定元帳生成":
						
						pass
					elif txt=="全補助元帳生成":
						
						pass
					elif txt=="次年度繰越":
						
						
						pass
					
					return False  # セル編集モードにしない。
				elif r>=VARS.splittedrow and c==VARS.daycolumn:  # 取引日列インデックスの時。
					datedialog.createDialog(enhancedmouseevent, xscriptcontext, "取引日", "YYYY-MM-DD")	
					return False  # セル編集モードにしない。
	return True  # セル編集モードにする。シングルクリックは必ずTrueを返さないといけない。
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
	datarange.setPropertyValues(("CellBackColor", "NumberFormat"), (-1, commons.formatkeyCreator(xscriptcontext.getDocument())("#,##0;[BLUE]-#,##0")))  # 背景色をクリア, 書式を設定。
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
		pass
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
