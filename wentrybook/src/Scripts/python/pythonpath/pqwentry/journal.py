#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 仕訳日誌シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
# import os, unohelper, glob
# from . import commons, datedialog, points, transientdialog
from . import commons, datedialog
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
# from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
class Journal():  # シート固有の値。
	def __init__(self):
		self.kamokurow = 2  # 科目行インデックス。
		self.hojokamokurow = 3  # 補助科目行インデックス。
		self.subtotalrow = 4  # 科目毎計行インデックス。
		self.splittedrow = 5  # 固定行インデックス。
		self.sliptotalcolumn = 0  # 伝票内計列インデックス。
		self.slipno = 1  # 伝票番号列インデックス。
		self.daycolumn = 2  # 取引日列インデックス。
		self.splittedcolumn = 4  # 固定列インデックス。
	def setSheet(self, sheet):  # 逐次変化する値。
		self.sheet = sheet
		cellranges = sheet[self.splittedrow:, self.daycolumn].queryContentCells(CellFlags.DATETIME)  # 取引日列の日付列が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 取引日列の最終行インデックス+1を取得。
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
	sheet["A1:A2"].setDataArray((("仕訳帳生成",), ("総勘定元帳生成",)))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左クリックの時。
		selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==2:  # 左ダブルクリックの時。まずselectionChanged()が発火している。
				celladdress = selection.getCellAddress()
				r, c = celladdress.Row, celladdress.Column  # selectionの行と列インデックスを取得。	
				if r>=VARS.splittedrow and c==VARS.daycolumn:  # 取引日列インデックスの時。
					datedialog.createDialog(enhancedmouseevent, xscriptcontext, "取引日", "YYYY-MM-DD")	
					return False
	return True  # セル編集モードにする。シングルクリックは必ずTrueを返さないといけない。		
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	selection = eventobject.Source.getSelection()	
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
		sheet = selection.getSpreadsheet()
		VARS.setSheet(sheet)		
		drowBorders(selection)  # 枠線の作成。
		splittedrow = VARS.splittedrow
		slipno = VARS.slipno
		datarange = sheet[splittedrow:VARS.emptyrow, slipno]
		datarange.setPropertyValue("CellBackColor", -1)  # 伝票番号列の背景色をクリア。
		sliprows = datarange.getDataArray()
		sliprowsset = set(sliprows)
		duperows = []
		if len(sliprows)>len(sliprowsset):  # 伝票番号列に重複行がある時。空文字も重複してはいけない。
			for i in sliprowsset:  # 重複は除いて伝票番号をイテレート。
				if sliprows.count(i)>1:  # 複数ある時。
					j = 0
					while i in sliprows[j:]:
						j = sliprows.index(i, j)
						duperows.append(j+splittedrow)  # 重複している伝票番号がある行インデックスを取得。
						j += 1
			cellranges = xscriptcontext.getDocument().createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
			cellranges.addRangeAddresses([sheet[i, slipno].getRangeAddress() for i in duperows], False)
			cellranges.setPropertyValue("CellBackColor", commons.COLORS["silver"])  # 重複伝票番号の背景色を返る。
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
			VARS.setSheet(sheet)  # 逐次変化する値を取得。取引日列の最終行を再取得したい。
			deadnos = sorted(set(range(1, VARS.emptyrow-VARS.splittedrow+1)).difference(i[0] for i in sheet[VARS.splittedrow:VARS.emptyrow, VARS.slipno].getDataArray()), reverse=True)  # 伝票番号の空き番号を取得して降順にする。
			createFormatKey = commons.formatkeyCreator(xscriptcontext.getDocument())
			for rangeaddress in cellranges.getRangeAddresses():  # セル範囲アドレスをイテレート。
				datarange = sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :VARS.emptycolumn]  # 行毎に処理するセル範囲を取得。
				datarange[:, VARS.daycolumn].setPropertyValue("NumberFormat", createFormatKey("YYYY-MM-DD"))  # 取引日列の書式を設定。
				newdatarows = []  # 処理後の行データを取得するリスト。
				for datarow in datarange.getDataArray():  # 各行をイテレート。
					if datarow[VARS.daycolumn]:  # 取引日列が入力されている時のみ。
						datarow = list(datarow)
						datarow[VARS.sliptotalcolumn] = sum(filter(lambda x: isinstance(x, float), datarow[VARS.splittedcolumn:]))  # 行の合計を取得。
						if not datarow[VARS.slipno]:  # 伝票番号列が空欄の時。
							datarow[VARS.slipno] = deadnos.pop()  # 空き番号を取得。
					newdatarows.append(datarow)
				datarange.setDataArray(newdatarows)
				datarange = sheet[VARS.splittedrow:VARS.emptyrow, rangeaddress.StartColumn:rangeaddress.EndColumn+1]  # 列毎に処理するセル範囲を取得。
				sheet[VARS.subtotalrow, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setDataArray(([sum(filter(lambda x: isinstance(x, float), i)) for i in zip(*datarange.getDataArray())],))  # 列ごとの合計を取得。
			datarange = sheet[VARS.splittedrow:VARS.emptyrow, VARS.sliptotalcolumn]  # 伝票内計列のセル範囲を取得。
			datarange.setPropertyValues(("CellBackColor", "NumberFormat"), (-1, createFormatKey("#,##0;[BLUE]-#,##0")))  # 背景色をクリア, 書式を設定。
			searchdescriptor = sheet.createSearchDescriptor()
			searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
			searchdescriptor.setSearchString("[^0]")  # 0以外のセルを取得。戻り値はない。	
			cellranges = datarange.queryContentCells(CellFlags.VALUE).findAll(searchdescriptor)  # 値のあるセルから0以外が入っているセル範囲コレクションを取得。見つからなかった時はNoneが返る。
			if cellranges:
				cellranges.setPropertyValue("CellBackColor", commons.COLORS["violet"])	
			sheet[VARS.subtotalrow:VARS.emptyrow, VARS.splittedcolumn:VARS.emptycolumn].setPropertyValue("NumberFormat", createFormatKey("#,##0;[BLUE]-#,##0"))
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。	
	contextmenuname, addMenuentry, baseurl, selection = commons.contextmenuHelper(VARS, contextmenuexecuteevent, xscriptcontext)
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上角のセルのアドレスを取得。
	r = celladdress.Row  # selectionの行と列のインデックスを取得。	
# 	if r<VARS.splittedrow or r==VARS.blackrow:  # 固定行より上、または黒行の時はコンテクストメニューを表示しない。
# 		return EXECUTE_MODIFIED
# 	elif contextmenuname=="cell":  # セルのとき。セル範囲も含む。
# 		commons.cutcopypasteMenuEntries(addMenuentry)
# 		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
# 		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
# 		addMenuentry("ActionTrigger", {"Text": "クリア", "CommandURL": baseurl.format("entry1")}) 
# 	elif contextmenuname=="rowheader" and len(selection[0, :].getColumns())==len(VARS.sheet[0, :].getColumns()):  # 行ヘッダーのとき、かつ、選択範囲の列数がシートの列数が一致している時。	
# 		if r>=VARS.splittedrow:
# 			if r<VARS.blackrow:
# 				addMenuentry("ActionTrigger", {"Text": "使用中最上行へ", "CommandURL": baseurl.format("entry15")})  # 黒行上から使用中最上行へ
# 				addMenuentry("ActionTrigger", {"Text": "使用中最下行へ", "CommandURL": baseurl.format("entry16")})  # 黒行上から使用中最下行へ
# 			elif r>VARS.blackrow:  # 黒行以外の時。
# 				addMenuentry("ActionTrigger", {"Text": "黒行上へ", "CommandURL": baseurl.format("entry17")})  
# 				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 				addMenuentry("ActionTrigger", {"Text": "使用中最上行へ", "CommandURL": baseurl.format("entry18")})  # 使用中から使用中最上行へ  
# 				addMenuentry("ActionTrigger", {"Text": "使用中最下行へ", "CommandURL": baseurl.format("entry19")})  # 使用中から使用中最下行へ		
# 			if r!=VARS.blackrow:
# 				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 				commons.cutcopypasteMenuEntries(addMenuentry)
# 				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 				commons.rowMenuEntries(addMenuentry)		
# 	elif contextmenuname=="colheader":  # 列ヘッダーの時。
# 		pass
# 	elif contextmenuname=="sheettab":  # シートタブの時。
# 		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。	
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。		
	controller = xscriptcontext.getDocument().getCurrentController()  # コントローラの取得。
	selection = controller.getSelection()  # 選択範囲を取得。
	if entrynum==1:  # クリア。書式設定とオブジェクト以外を消去。
		selection.clearContents(511)  # 範囲をすべてクリアする。
	elif 14<entrynum<20:
		sheet = controller.getActiveSheet()  # アクティブシートを取得。
		rangeaddress = selection.getRangeAddress()  # 選択範囲のアドレスを取得。
		if entrynum==15:  # 黒行上から使用中最上行へ
			commons.toOtherEntry(sheet, rangeaddress, VARS.blackrow, VARS.blackrow+1)
		elif entrynum==16:  # 黒行上から使用中最下行へ
			commons.toNewEntry(sheet, rangeaddress, VARS.blackrow, VARS.emptyrow) 
		elif entrynum==17:  # 黒行上へ
			commons.toOtherEntry(sheet, rangeaddress, VARS.emptyrow, VARS.blackrow)  
		elif entrynum==18:  # 使用中から使用中最上行へ 
			commons.toOtherEntry(sheet, rangeaddress, VARS.emptyrow, VARS.blackrow+1)
		elif entrynum==19:  # 使用中から使用中最下行へ		
			commons.toNewEntry(sheet, rangeaddress, VARS.emptyrow, VARS.emptyrow) 		
