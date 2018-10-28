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
		self.slipsubtotalcolumn = 0  # 伝票小計列インデックス。
		self.slipno = 1  # 伝票番号列インデックス。
		self.daycolumn = 2  # 取引日列インデックス。
		self.splittedcolumn = 4  # 固定列インデックス。
	def setSheet(self, sheet):  # 逐次変化する値。
		self.sheet = sheet
		cellranges = sheet[self.splittedrow:, self.daycolumn].queryContentCells(CellFlags.DATETIME)  # 取引日列の日付列が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 取引日列の最終行インデックス+1を取得。

		# 科目行または補助科目行の右端空列を取得。


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
		VARS.setSheet(selection.getSpreadsheet())		
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
	sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く
	sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。
	selection.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。		
def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。マクロで変更した時は発火しない。	
	selection = None
	for change in changesevent.Changes:
		if change.Accessor=="cell-change":  # セルの値が変化した時。
			selection = change.ReplacedElement  # 値を変更したセルを取得。	
			break
	if selection:  # セルとは限らずセル範囲のときもある。シートからペーストしたときなど。テキストをペーストした時は発火しない。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		
		
		sheet = VARS.sheet
# 		splittedrow = VARS.splittedrow
# 		idcolumn = VARS.idcolumn
# 		kanjicolumn = VARS.kanjicolumn
# 		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
# 		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
# 		rangeaddress = selection.getRangeAddress()
# 		transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
# 		transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))			
# 		for r in range(rangeaddress.StartRow, rangeaddress.EndRow+1):
# 			for c in range(rangeaddress.StartColumn, rangeaddress.EndColumn+1):
# 				if r>=splittedrow:  # 分割行以降の時。
# 					txt = sheet[r, c].getString()  # セルの文字列を取得。			
# 					if c==idcolumn:  # ID列の時。
# 						txt = transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換。
# 						if txt.isdigit():  # 数値の時のみ。空文字の時0で埋まってしまう。
# 							sheet[r, c].setString("{:0>8}".format(txt))  # 数値を8桁にして文字列として代入し直す。
# 					elif c==kanjicolumn:
# 						sheet[r, c].setString(txt.replace("　", " "))  # 全角スペースを半角スペースに置換。
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
