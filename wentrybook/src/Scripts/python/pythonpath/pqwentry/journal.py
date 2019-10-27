#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 振替一覧シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import json
import os
from collections import OrderedDict
from datetime import date, datetime, timedelta
from itertools import chain, compress, count, islice, zip_longest

import unohelper
from com.sun.star.awt import (MessageBoxButtons, MessageBoxResults,  # 定数
                              MouseButton)
from com.sun.star.awt.MessageBoxType import QUERYBOX, WARNINGBOX  # enum
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.sheet.CellInsertMode import ROWS as insert_rows  # enum
from com.sun.star.table import CellVertJustify2  # Struct; 定数
from com.sun.star.table import BorderLine2, TableBorder2
from com.sun.star.table.CellHoriJustify import CENTER, LEFT, RIGHT  # enum
from com.sun.star.table.CellOrientation import STACKED  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import \
    EXECUTE_MODIFIED  # enum
from com.sun.star.util import XModifyListener

from . import (commons, datedialog, dialogcommons, documentevent,
               historydialog, menudialog)


class Journal():  # シート固有の値。
	def __init__(self):
		self.kamokurow = 2  # 科目行インデックス。この上行は科目分類行、下行は補助科目行。
		self.splittedrow = 5  # 固定行インデックス。この上行は列別小計行。
		self.sliptotalcolumn = 0  # 伝票内計列インデックス。
		self.daycolumn = 2  # 取引日列インデックス。この左列は伝票番号列、右列が摘要列。
		self.splittedcolumn = 4  # 固定列インデックス。	
		self.settlingdayrows = 1, 3  # 期首日セルと期末日セルの行インデックスのタプル。
	def setSheet(self, sheet):  # シートの逐次変化する値。
		self.sheet = sheet
		rowedges = []
		cellranges = sheet[self.splittedrow:, self.daycolumn].queryContentCells(CellFlags.DATETIME+CellFlags.VALUE)  # 取引日列の日付が入っているセルに限定して抽出。書式設定される前のセルも取得する。
		if len(cellranges):
			rowedges.append(cellranges.getRangeAddresses()[-1].EndRow+1)  # 取引日列の最終行インデックス+1を取得。
		cellranges = sheet[self.splittedrow:, self.daycolumn+1].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # 摘要列の文字列か数値が入っているセルに限定して抽出。
		if len(cellranges):
			rowedges.append(cellranges.getRangeAddresses()[-1].EndRow+1)  # 摘要列の最終行インデックス+1を取得。		
		self.emptyrow = max(rowedges) if rowedges else VARS.splittedrow # 取引日列または摘要列の最終行インデックス+1を取得。		
		columnedges = []
		cellranges = sheet[self.kamokurow, self.splittedcolumn:].queryContentCells(CellFlags.STRING) 
		if len(cellranges):
			columnedges.append(cellranges.getRangeAddresses()[-1].EndColumn+1)  # 科目行の右端+1インデックスを取得。
		cellranges = sheet[self.kamokurow+1, self.splittedcolumn:].queryContentCells(CellFlags.STRING) 
		if len(cellranges):
			columnedges.append(cellranges.getRangeAddresses()[-1].EndColumn+1)  # 補助科目行の右端+1インデックスを取得。
		self.emptycolumn = max(columnedges) if columnedges else VARS.splittedcolumn  # 科目行または補助科目行の右端空列を取得。
VARS = Journal()
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。
	initSheet(activationevent.ActiveSheet, xscriptcontext)
def initSheet(sheet, xscriptcontext):	
	sheet["A1"].setString("メニュー")  # 入力間違いしやすいボタンセルの値を代入。
	VARS.setSheet(sheet)  # 逐次変化するシートの値を取得。
class SettlingDayModifyListener(unohelper.Base, XModifyListener):
	def __init__(self, xscriptcontext):	
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
		self.functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)
		doc = xscriptcontext.getDocument()
		self.setProperty = lambda x: x.setPropertyValue("NumberFormat", commons.formatkeyCreator(doc)("YYYY-MM-DD"))
		self.showErrorMessageBox = lambda x: commons.showErrorMessageBox(doc.getCurrentController()	, x)
	def modified(self, eventobject):  # 決算日セルが変化したら発火するメソッド。eventobject.Sourceには全シートの決算日セルのセル範囲コレクションが入っている。
		if VARS.sheet.getName().startswith("振替一覧"):
			sdaycell, edaycell = [VARS.sheet[i, VARS.daycolumn] for i in VARS.settlingdayrows]
			sdatevalue = sdaycell.getValue()  # 期首日セルの値を取得。空セルや文字のときは0.0が返る。
			edatevalue = edaycell.getValue()  # 期末日セルの値を取得。空セルや文字のときは0.0が返る。
			if sdatevalue>0 and edatevalue>0:  # 期首日も期末日の入力されている時。
				if sdatevalue<edatevalue:  # 期首日<期末日の時
					if edatevalue<self.functionaccess.callFunction("EDATE", (sdatevalue, 12)):  # 期末日が期首日の1年以内の時。
						pass
					else:
						self.showErrorMessageBox("期首日と期末日の間隔は1年以内にしてください。")
				else:
					self.showErrorMessageBox("期首日が期末日より古いので訂正してください。")	
				return
			elif sdatevalue>0:  # 期首日のみの時。
				eventobject.Source.removeModifyListener(self)  # WindowsではModifyListenerを外しておかないとクラッシュする。
				edaycell.setValue(self.functionaccess.callFunction("EDATE", (sdatevalue, 12))-1)  # 1年後を期末日にする。
				eventobject.Source.addModifyListener(self)  # ModifyListenerを付け直す。
			elif edatevalue>0:  # 期末日のみの時。
				eventobject.Source.removeModifyListener(self)  # WindowsではModifyListenerを外しておかないとクラッシュする。
				sdaycell.setValue(self.functionaccess.callFunction("EDATE", (edatevalue, -12))+1)  # 1年前を期首日にする。
				eventobject.Source.addModifyListener(self)  # ModifyListenerを付け直す。
			self.setProperty(sdaycell)
			self.setProperty(edaycell)
	def disposing(self, eventobject):
		eventobject.Source.removeModifyListener(self)
class ValueModifyListener(unohelper.Base, XModifyListener):  # WindowsではModifyListenerからModifyListenerを発火させるとクラッシュする。
	def __init__(self, xscriptcontext, slipnosubjectmodifylistener):
		self.formatkey = commons.formatkeyCreator(xscriptcontext.getDocument())("#,##0;[BLUE]-#,##0")
		self.slipnosubjectmodifylistener = slipnosubjectmodifylistener
	def modified(self, eventobject):  # 固定行以下固定列右のセルが変化すると発火するメソッド。サブジェクトのどこが変化したかはわからない。eventobject.Sourceは対象全シートのセル範囲コレクション。
		if VARS.sheet.getName().startswith("振替一覧"):
			sheet = VARS.sheet
			VARS.setSheet(sheet)  # 最終行と列を取得し直す。
			if VARS.splittedrow<VARS.emptyrow:  # 伝票行がある時のみ。
				if VARS.splittedcolumn<VARS.emptycolumn:  # 科目列がある時のみ。
					datarange = sheet[VARS.splittedrow:, VARS.sliptotalcolumn]
					datarange.clearContents(CellFlags.VALUE)
					datarange.setPropertyValue("CellBackColor", -1)
					datarows = sheet[VARS.splittedrow:VARS.emptyrow, VARS.splittedcolumn:VARS.emptycolumn].getDataArray()  # 伝票金額の全データ行を取得。
					sheet[VARS.splittedrow-1, VARS.splittedcolumn:VARS.emptycolumn].setDataArray(([sum(filter(None, i)) for i in zip(*datarows)],))  # 列ごとの合計を再計算。空セルの空文字を除いて合計する。
					datarange = sheet[VARS.splittedrow:VARS.emptyrow, VARS.sliptotalcolumn]  # 伝票内計列のセル範囲を取得。
					datarange.setDataArray((sum(filter(lambda x: isinstance(x, float), i)),) for i in datarows)  # 伝票内計列を再計算。
					datarange.setPropertyValue("NumberFormat", self.formatkey)  # 伝票内計列の書式を設定。
					searchdescriptor = sheet.createSearchDescriptor()
					searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
					searchdescriptor.setSearchString("[^0]")  # 0以外のセルを取得。戻り値はない。	
					cellranges = datarange.queryContentCells(CellFlags.VALUE).findAll(searchdescriptor)  # 値のあるセルから0以外が入っているセル範囲コレクションを取得。見つからなかった時はNoneが返る。
					if cellranges:
						cellranges.setPropertyValue("CellBackColor", commons.COLORS["violet"])  # 不均衡セルをハイライト。	
					sheet[VARS.splittedrow:VARS.emptyrow, VARS.splittedcolumn:VARS.emptycolumn].setPropertyValue("NumberFormat", self.formatkey)  # 伝票金額セルの書式を設定。	
				else:  # 科目列がない時。資産/現金、を先頭列に挿入。
					sheet[:VARS.kamokurow+1, VARS.splittedcolumn].setDataArray((("賃借対照表",), ("資産の部",), ("現金",)))
			else:  # 伝票行がない時。
				subj, listener = self.slipnosubjectmodifylistener
				subj.removeModifyListener(listener)  # WindowsではModifyListenerを外しておかないとクラッシュする。
				sheet[VARS.splittedrow, VARS.daycolumn:VARS.daycolumn+2].setDataArray(((sheet[VARS.settlingdayrows[0], VARS.daycolumn].getValue(), "前期より繰越"),))  # 繰越行を挿入。
				sheet[VARS.splittedrow-1, VARS.splittedcolumn:VARS.emptycolumn].setDataArray(((0,)*(VARS.emptycolumn-VARS.splittedcolumn),))  # 列毎小計を0にリセット。
				subj.addModifyListener(listener)  # ModifyListenerを付け直す。
	def disposing(self, eventobject):
		eventobject.Source.removeModifyListener(self)
class SlipNoModifyListener(unohelper.Base, XModifyListener):
	def __init__(self, xscriptcontext):	
		doc = xscriptcontext.getDocument()
		self.doc = doc
		self.formatkey = commons.formatkeyCreator(doc)("YYYY-MM-DD")
	def modified(self, eventobject):  # 伝票番号列や取引日列が変化した時に発火するメソッド。eventobject.Sourceは対象全シートのセル範囲コレクション。
		sheet = VARS.sheet
		if sheet.getName().startswith("振替一覧"):
			splittedrow = VARS.splittedrow
			VARS.setSheet(sheet)  # 最終行と列を取得し直す。
			if splittedrow<VARS.emptyrow:  # 伝票行がある時のみ。
				sheet[splittedrow:, VARS.daycolumn-1].setPropertyValue("CellBackColor", -1)  # 伝票番号列の背景色をクリア。
				datarange = sheet[splittedrow:VARS.emptyrow, VARS.daycolumn-1]  # 取引日の入力がある行までの伝票番号列のセル範囲を取得。
				sliprows = list(datarange.getDataArray())  # 伝票番号列の行をリストにして取得。
				i = ("",)  # 空セルの行。
				if i in sliprows:  # 空セルの行がある時。
					deadnogene = (j for j in count(1) if j not in list(chain.from_iterable(sliprows)))  # 空伝票番号のイテレーター。
					j = 0
					while i in sliprows[j:]:  # 空セルの行を空伝票番号を入れた行に置き換える。
						j = sliprows.index(i, j)
						sliprows[j] = next(deadnogene),
						j += 1
					eventobject.Source.removeModifyListener(self)  # WindowsではModifyListenerを外しておかないとクラッシュする。	
					datarange.setDataArray(sliprows)		
					eventobject.Source.addModifyListener(self)  # ModifyListenerを付け直す。
				sliprowsset = set(sliprows)  # 重複行を削除した集合を取得。		
				duperows = []  # 重複している伝票番号がある行インデックスを取得するリスト。
				if len(sliprows)>len(sliprowsset):  # 伝票番号列に重複行がある時。空文字の重複でもTrue。
					for i in sliprowsset:  # 重複は除いて伝票番号をイテレート。
						if sliprows.count(i)>1:  # 伝票番号が複数ある時。
							j = 0
							while i in sliprows[j:]:
								j = sliprows.index(i, j)
								duperows.append(j+splittedrow)  # 重複している伝票番号がある行インデックスを取得。
								j += 1		
				if duperows:  # 重複している伝票行がある時。
					cellranges = self.doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # com.sun.star.sheet.SheetCellRangesをインスタンス化。
					cellranges.addRangeAddresses([sheet[i, VARS.daycolumn-1].getRangeAddress() for i in duperows], False)
					cellranges.setPropertyValue("CellBackColor", commons.COLORS["silver"])  # 重複伝票番号の背景色を変える。	
				sheet[splittedrow:VARS.emptyrow, VARS.daycolumn].setPropertyValue("NumberFormat", self.formatkey)			
	def disposing(self, eventobject):
		eventobject.Source.removeModifyListener(self)		
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左クリックの時。
		selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==2:  # 左ダブルクリックの時。まずselectionChanged()が発火している。
				celladdress = selection.getCellAddress()
				r, c = celladdress.Row, celladdress.Column  # selectionの行と列インデックスを取得。		
				if r<VARS.splittedrow and c<VARS.splittedcolumn:  # 左上枠の時。
					if r in VARS.settlingdayrows and c==VARS.daycolumn:  # 期首日セルや期末日セルの時。
						datedialog.createDialog(enhancedmouseevent, xscriptcontext, "決算日")  # 書式はSettlingDayModifyListenerで設定する。	
					else:							
						txt = selection.getString()	
						if txt=="メニュー":
							defaultrows = "日付順に並び替え", "------", "仕訳日記帳生成", "総勘定元帳生成", "全補助元帳生成", "決算書生成", "------", "次年度繰越"
							menudialog.createDialog(xscriptcontext, txt, defaultrows, enhancedmouseevent=enhancedmouseevent, callback=callback_menuCreator(xscriptcontext))
					return False  # セル編集モードにしない。
				elif r>=VARS.splittedrow and c==VARS.daycolumn:  # 取引日列の時。
					datedialog.createDialog(enhancedmouseevent, xscriptcontext, "取引日")  # 書式はSlipNoModifyListenerで設定する。
					return False  # セル編集モードにしない。
	return True  # セル編集モードにする。シングルクリックは必ずTrueを返さないといけない。		
def callback_menuCreator(xscriptcontext):  # 内側のスコープでクロージャの変数を再定義するとクロージャの変数を参照できなくなる。	
	componentwindow = xscriptcontext.getDocument().getCurrentController().ComponentWindow
	querybox = lambda x: componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "WEntryBook", x)
	def callback_menu(gridcelltxt):			
		if gridcelltxt=="仕訳日記帳生成":	
			msgbox = querybox("{}します。".format(gridcelltxt))
			if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
				return	
			createShiwakeCho(xscriptcontext, gridcelltxt)
		elif gridcelltxt=="総勘定元帳生成":
			msgbox = querybox("{}します。".format(gridcelltxt))
			if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
				return	
			createMotoCho(xscriptcontext, gridcelltxt, "総勘定元帳", lambda x: compress(*(x[VARS.kamokurow][VARS.splittedcolumn:],)*2))
		elif gridcelltxt=="全補助元帳生成":
			msgbox = querybox("{}します。".format(gridcelltxt))
			if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
				return	
			createHojoMotoCho(xscriptcontext, gridcelltxt, "全補助元帳", lambda x: range(len(x[0])))	
		elif gridcelltxt=="決算書生成":
			msgbox = querybox("{}します。\n試算表、損益計算書、賃借対照表を生成します。".format(gridcelltxt))
			if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
				return	
			createFinancialStatements(xscriptcontext, gridcelltxt)
		elif gridcelltxt=="次年度繰越":
			startday, endday = [VARS.sheet[i, VARS.daycolumn].getString() for i in VARS.settlingdayrows]
			if not all([startday, endday]):  # 開始日と終了日、いずれかが空文字の時。
				commons.showErrorMessageBox(xscriptcontext.getDocument().getCurrentController(), "期首日と期末日を入力してください。\n処理を中止します。")	
				return					
			msgbox = querybox("期首日: {} 期末日: {}\nを{}します。".format(startday, endday, gridcelltxt))
			if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
				return			
			kurikoshi(xscriptcontext, querybox, gridcelltxt, startday, endday)
		elif gridcelltxt=="日付順に並び替え":
			msgbox = querybox("{}します。".format(gridcelltxt))
			if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
				return					
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
			dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)			
			sortSlips(dispatcher, xscriptcontext.getDocument())  # 日付順に伝票をソート。			
	return callback_menu
def kurikoshi(xscriptcontext, querybox, txt, startday, endday):
	doc = xscriptcontext.getDocument()
	controller = doc.getCurrentController()		
	indicator = controller.getFrame().createStatusIndicator()  # 現ドキュメントのステータスインディケーターを取得。				
	indicator.start("{}中".format(txt), 0)		
	sheet = VARS.sheet
	splittedrow = VARS.splittedrow	
	daycolumn = VARS.daycolumn
	slipnocolumn = daycolumn - 1
	tekiyocolumn = daycolumn + 1
	splittedcolumn = VARS.splittedcolumn
	headerrows, datarows = getDataRows(xscriptcontext)  # 科目ヘッダー行とすべてのデータ行を取得。
	if not headerrows:  # 伝票書式のエラーに引っかかった時ここで終わる。
		return
	indicator.setText("次期のシートを取得")
	sheetname = sheet.getName()  # 現シート名を取得。
	settledaytxt = "{}決算".format(endday.replace("-", ""))
	if not sheetname.endswith(settledaytxt):
		sheet.setName("_".join([sheetname, settledaytxt]))  # 決算日を最後につけた名前にする。		
	sheetname = sheet.getName()  # 現シート名を再取得。	
	sheets = doc.getSheets()	
	sheetnames = sorted(sheets.getElementNames())  # 全シート名のリストをソートして取得する。
	newi = sheetnames.index(sheetname) + 1	 # 現シートの次の位置を取得。
	newsheet = None
	if newi<len(sheetnames):  # 現シート名の次の位置にインデックスをある時。
		newsheetname = sheetnames[newi]  # 次の位置のシート名を取得。
		if newsheetname.startswith("振替一覧"):  # 次期の振替一覧がすでにある時。
			msgbox = querybox("{}はすでに存在します。\n金額のみ繰り越しますか?".format(newsheetname))
			if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
				return							
			newsheet = sheets[newsheetname]  # 既存の次期シートを取得。		
			VARS.setSheet(newsheet)	 # 新規シートに更新する。これをしないとこのシートにModifyListenerが影響しない。
			newdatarows = newsheet[:VARS.kamokurow+2, :VARS.emptycolumn].getDataArray()  # 補助科目行までのデータ行を取得。
			newheaderrowsgene = zip(*generateHeaderRows(newdatarows[:VARS.kamokurow+2])[1:])  # (区分行、科目行、補助科目行)をイテレートする。		
			if newsheet[splittedrow, daycolumn+1].getString()!="前期より繰越":  # 先頭行が繰越伝票でない時。
				newsheet.insertCells(newsheet[splittedrow, :].getRangeAddress(), insert_rows)  # 空行を挿入。
				slipnosubjectmodifylistener = documentevent.addModifyListener(doc, [newsheet[splittedrow, slipnocolumn:tekiyocolumn].getRangeAddress()], SlipNoModifyListener(xscriptcontext))  # 新規行にModifyListenerを付ける。
				documentevent.addModifyListener(doc, [newsheet[splittedrow, splittedcolumn:].getRangeAddress()], ValueModifyListener(xscriptcontext, slipnosubjectmodifylistener))  # 新規行にModifyListenerを付ける。  
	if not newsheet:  # まだ次期シートが取得できていない時。
		edate = date(*map(int, endday.split("-")))  # 現シートの期末日のdateオブジェクトを取得。
		newsdate = edate + timedelta(days=1)  # 次期期首日のdateオブジェクトを取得。		
		newedate = date(edate.year+1, edate.month, edate.day)  # 次期期末日のdateオブジェクトを取得。期間の差で取得するとうるう年を考慮する必要がある。
		newsheetname = "振替一覧_{}決算".format(newedate.isoformat().replace("-", ""))
		sheets.copyByName(sheetname, newsheetname, newi)  # 現シートをコピーして次期シートにする。
		newsheet = sheets[newsheetname]
		VARS.setSheet(newsheet)	 # 新規シートに更新する。これをしないとこのシートにModifyListenerが影響しない。
		newsheet[splittedrow:, :].clearContents(CellFlags.VALUE+CellFlags.DATETIME+CellFlags.STRING+CellFlags.ANNOTATION+CellFlags.FORMULA)  # 全伝票を全削除。
		newsdaycell, newedaycell = [newsheet[i, daycolumn] for i in VARS.settlingdayrows]  # 次期シートの期首日セルと期末日セルを取得。
		newsdaycell.setFormula(newsdate.isoformat())  # 新規期首日を代入。
		newedaycell.setFormula(newedate.isoformat())  # 新規期末日を代入。				
		documentevent.addModifyListener(doc, (i.getRangeAddress() for i in (newsdaycell, newedaycell)), SettlingDayModifyListener(xscriptcontext))  # 次期シートにModifyLsitenerの追加。
		slipnosubjectmodifylistener = documentevent.addModifyListener(doc, [newsheet[splittedrow:, slipnocolumn:tekiyocolumn].getRangeAddress()], SlipNoModifyListener(xscriptcontext))  # 次期シートにModifyLsitenerの追加。
		documentevent.addModifyListener(doc, [newsheet[splittedrow:, splittedcolumn:].getRangeAddress()], ValueModifyListener(xscriptcontext, slipnosubjectmodifylistener))  # 次期シートにModifyLsitenerの追加。
		newheaderrowsgene = zip(*headerrows[1:])  # (区分行、科目行、補助科目行)をイテレートする。			
	indicator.start("次期繰越金を算出", len(datarows[0]))		
	columnstotaldic = {i[:-1]: i[-1] for i in zip(*headerrows[1:], datarows[VARS.splittedrow-1][VARS.splittedcolumn:]) if i[-1]}  # キー: (区分、科目、補助科目)のタプル、値: 各列計、の辞書を取得。各列0が0や空セルのものは取得しない。
	newmotoire = sum(v for k, v in columnstotaldic.items() if (k[0] in ("経費", "収益")) or (k[1] in ("事業主貸", "事業主借", "元入金")))  # 事業主貸は正、事業主借は負、元入金は負、経費は正、収益は負、なのですべて合計すれば新元入金になる。
	carryovers = []  # 繰越行を取得するリスト。
	t = 1
	for i in newheaderrowsgene:  # 次期の(区分、科目、補助科目)をイテレート。
		indicator.setValue(t)	
		t += 1
		if i[1]=="元入金":  # 科目が元入金の時。
			val = newmotoire  # 新元入金を取得。		
			del columnstotaldic[i]  # 辞書から消去。		
		elif i in columnstotaldic:  # 前期の(区分、科目、補助科目)が一致するものがあるとき。
			if (i[0] in ("経費", "収益")) or (i[1] in ("事業主貸", "事業主借")):  # 区分が経費や収益の時、または、科目が事業主貸や事業主借の時。
				del columnstotaldic[i]  # 辞書から消去。
				val = ""  # 空セル。
			else:
				val = columnstotaldic.pop(i)  # 小計を取得、して辞書から消去する。。					
		else:
			val = ""
		carryovers.append(val or "")  # 0のときは空文字を返す。
	datarow = (newsheet[VARS.settlingdayrows[0], daycolumn].getValue(), "前期より繰越", *carryovers)  # 丸括弧がないと*で展開できない。
	controller.setActiveSheet(newsheet)  # 次期シートをアクティブにする。
	newsheet[splittedrow, daycolumn:daycolumn+len(datarow)].setDataArray((datarow,))  # 繰越金行を代入。
	indicator.end()  # reset()の前にend()しておかないと元に戻らない。
	indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。	
	if len(columnstotaldic)>0:  # 辞書に要素が残っている時。
		msg = "新規年度にはない科目がありました。\n\n{}".format("\n".join("{}: {}".format(*i) for i in columnstotaldic.items()))
		componentwindow = controller.ComponentWindow
		componentwindow.getToolkit().createMessageBox(componentwindow, WARNINGBOX, MessageBoxButtons.BUTTONS_OK, "WEntryBook", msg).execute()	
def createFinancialStatements(xscriptcontext, txt):  # 決算書作成。
	newdoc = xscriptcontext.getDesktop().loadComponentFromURL("private:factory/scalc", "_blank", 0, ())  # 新規ドキュメントの取得。	
	indicator = newdoc.getCurrentController().getFrame().createStatusIndicator()  # 新規ドキュメントのステータスインディケーターを取得。				
	indicator.start("{}中".format(txt), VARS.emptycolumn)  # 新規ドキュメントを作成後はステータスバーを表示できない。		
	doc = xscriptcontext.getDocument()	
	datetxtforfile, *datetxts = getDaytxts()
	headerrows, datarows = getDataRows(xscriptcontext)	
	if not headerrows:
		commons.showErrorMessageBox(doc.getCurrentController(), "シートのデータが取得できません。\n処理を中止します。")	
		return		
	addToTrialB, createTrialBalance = createTrialBalanceCreator(xscriptcontext, datetxts)	 # 試算表作成のための関数を取得。
	addPL, createPL = createProfitAndLossCreator(xscriptcontext, datetxts)	 # 損益計算書作成のための関数を取得。
	addToBS, createBalanceSheet = createBalanceSheetCreator(xscriptcontext, datetxts)	 # 賃借対照表作成のための関数を取得。
	bkarikata = []  # 各科目の期首借方金額を入れるリスト。
	bkashikata = []  # 各科目の期首貸方金額を入れるリスト。
	karikata = []  # 各科目の期中借方金額を入れるリスト。
	kashikata = []  # 各科目の期中貸方金額を入れるリスト。
	ekarikata = []  # 各科目の期末借方金額を入れるリスト。
	ekashikata = []  # 各科目の期末貸方金額を入れるリスト。
	kamoku = ""  # 科目のキャッシュ。
	kubun = ""  # 区分のキャッシュ。賃借対照表作成用。
	flg = True if "繰越" in datarows[VARS.splittedrow][VARS.daycolumn+1] else False  # 繰越フラグ。繰越行がないときは期首データはないということ。
	for i in zip(*headerrows, *[i[VARS.splittedcolumn:] for i in datarows[VARS.splittedrow-1:]]):  # 列インデックス、区分、科目、補助科目、列合計、固定列以下の列の要素、をイテレート。
		indicator.setText("{:　>10}を処理中".format(i[2]))  # 全角スペースで埋める。
		indicator.setValue(i[0])
		if kamoku!=i[2]:  # 科目が切り替わった時。
			sums = list(map(sum, (bkarikata, bkashikata, karikata, kashikata, ekarikata, ekashikata)))  # 各リストの合計のリストを取得。
			if sum(sums):  # 0でない要素がある時のみ。
				addToTrialB(kamoku, sums)  # 試算表データに追加。
				addPL(kubun, kamoku, sums)  # 損益計算書データを追加。			
				addToBS(kubun, kamoku, sums)  # 賃借対照表データに追加。
			bkarikata = []
			bkashikata = []
			karikata = []
			kashikata = []
			ekarikata = []
			ekashikata = []						
			kamoku = i[2]  # 科目のキャッシュを更新。
			kubun = i[1]  # 区分のキャッシュを更新。
			sign = -1 if i[1].startswith(("負債", "収益")) else 1  # 区分が負債または収益から始まっている時は残高は貸方を正とするため-1をかける。	
		startrow = 5  # 固定行以下の要素の開始インデックス。
		if flg:  # 繰越行がある時。
			if sign>0:  # 借方科目の時。
				bkarikata.append(i[startrow] or 0)
				bkashikata.append(0)
			else:  # 貸方科目の時。
				bkarikata.append(0)
				bkashikata.append((i[startrow] or 0)*sign)								
			startrow += 1  # 繰越行より下のみ加算するようにする。
		else:  # 繰越行がない時。期首データはなし。
			bkarikata.append(0)
			bkashikata.append(0)	
		karikata.append(sum(filter(lambda x: x and x>0, i[startrow:])))  # 列の借方合計を取得。空文字を除く。
		kashikata.append(-sum(filter(lambda x: x and x<0, i[startrow:])))  # 列の貸方合計を取得。空文字を除く。
		if sign>0:  # 借方科目の時。
			ekarikata.append(i[4] or 0)
			ekashikata.append(0)
		else:  # 貸方科目の時。	
			ekarikata.append(0)
			ekashikata.append((i[4] or 0)*sign)
	width, leftmargin, rightmargin = newdoc.getStyleFamilies()["PageStyles"]["Default"].getPropertyValues(("Width", "LeftMargin", "RightMargin"))
	pagewidth = width - leftmargin - rightmargin - 5  # 印刷幅を1/100mmで取得。なぜかはみ出るのでマージンを取る。	
	indicator.setText("試算表を描画中")
	createTrialBalance(newdoc, pagewidth)  # 試算表シートの作成。
	indicator.setText("損益計算書を描画中")
	createPL(newdoc, pagewidth)  # 損益計算書シートの作成。
	indicator.setText("賃借対照表を描画中")
	createBalanceSheet(newdoc, pagewidth)  # 賃借対照表シートの作成。
	newdocname = "決算書_{}_{}.ods".format(datetxtforfile, datetime.now().strftime("%Y%m%d%H%M%S"))
	indicator.setText("ファイルに保存中 {}".format(newdocname))	
	newdoc.getStyleFamilies()["PageStyles"]["Default"].setPropertyValue("HeaderIsOn", False)  # 印刷時ヘッダーを付けない。
	saveNewDoc(doc, newdoc, newdocname)	
	indicator.end()  # reset()の前にend()しておかないと元に戻らない。
	indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。	
def createProfitAndLossCreator(xscriptcontext, datetxts):	# 損益通算書の作成。
	expensesorder = "租税公課", "荷造運賃", "水道光熱費", "旅費交通費", "通信費", "広告宣伝費", "接待交際費",\
					"損害保険料", "修繕費", "消耗品費", "減価償却費", "福利厚生費", "給料賃金", "外注工賃",\
					"利子割引料", "地代家賃", "貸倒金", "雑費"  # 印字されている経費の科目名の順。
	expensesvaluedic = {}  # キー: 経費の科目名, 値: 金額。
	kamokuvaluedic = {}  # キー: 科目、値: 金額、の辞書。
	def addPL(kubun, kamoku, sums):
		if kubun=="経費":  # 借方科目。
			if kamoku in ("専従者給与", "貸倒引当金繰入", "期首商品棚卸高", "仕入金額"):
				kamokuvaluedic[kamoku] = sums[4]
			else:  # その他の経費。
				expensesvaluedic[kamoku] = sums[4]
		elif kubun=="収益":  # 貸方科目。"売上金額", "貸倒引当金繰戻", "期末商品棚卸高"。これ以外の収益は想定していない。
			kamokuvaluedic[kamoku] = sums[5]
	def createPL(newdoc, pagewidth):
		datetxtforsheet, presentdatetxt, dummy = datetxts
		newsheets = newdoc.getSheets()
		newsheetname = "損益計算書"		
		newsheets.insertNewByName(newsheetname, len(newsheets))
		newsheet = newsheets[newsheetname]			
		newdatarows = [("損益計算書", "", "", ""),\
					(datetxtforsheet, "", "", presentdatetxt),\
					("科目", "", "", "金額")]  # 新規シートのヘッダー行。
		newdatarows.append(("売上(収入)金額", "", "", kamokuvaluedic.get("売上金額", 0)))			
		newdatarows.append(("売上原価", "期首商品棚卸高", "", kamokuvaluedic.get("期首商品棚卸高", 0)))
		newdatarows.append(("", "仕入金額", "", kamokuvaluedic.get("仕入金額", 0)))
		newdatarows.append(("", "小計", "", newdatarows[-2][-1]+newdatarows[-1][-1]))
		newdatarows.append(("", "期末商品棚卸高", "", kamokuvaluedic.get("期末商品棚卸高", 0)))
		newdatarows.append(("", "差引原価", "", newdatarows[-2][-1]-newdatarows[-1][-1]))
		grossprofit = newdatarows[3][-1] - newdatarows[-1][-1]
		newdatarows.append(("差引金額", "", "", grossprofit))
		expensestotal = sum(expensesvaluedic.values())  # 経費の合計を取得。
		profit = grossprofit - expensestotal
		i = "租税公課"  # 経費の先頭科目名。
		expensesdatarows = [("経費", i, "", expensesvaluedic.pop(i, 0))]
		expensesdatarows.extend(("", i, "", expensesvaluedic.pop(i, 0)) for i in expensesorder[1:])
		lastexpensesrow = expensesdatarows.pop()  # 最後の「雑費」だけよけておく。
		if expensesvaluedic:  # まだ項目が残っている時。
			expensesdatarows.extend(("", k, "", v) for k, v in expensesvaluedic.items())
		expensesdatarows.append(lastexpensesrow)	
		newdatarows.extend(expensesdatarows)
		newdatarows.append(("", "計", "", expensestotal))
		newdatarows.append(("差引金額", "", "", profit))
		newdatarows.append(("各種引当金\n・準備金等", "繰戻額等", "貸倒引当金", kamokuvaluedic.get("貸倒引当金繰戻", 0)))
		fb = newdatarows[-1][-1]
		newdatarows.append(("", "", "計", fb))
		newdatarows.append(("", "繰入額等", "専従者給与", kamokuvaluedic.get("専従者給与", 0)))
		newdatarows.append(("", "", "貸倒引当金", kamokuvaluedic.get("貸倒引当金繰入", 0)))
		pb = newdatarows[-2][-1] + newdatarows[-1][-1]
		newdatarows.append(("", "", "計", pb))
		newdatarows.append(("青色申告特別控除前の所得金額", "", "", profit+fb-pb))
		columnscount = len(newdatarows[0])
		newsheet[0, :columnscount].merge(True)  # 題名セルの結合。	
		expensesendrow = len(expensesdatarows) + 11
		for i in (2, 3, 9, expensesendrow, expensesendrow+6):
			newsheet[i, :3].merge(True)		
		for i in chain(range(4, 9), range(10, expensesendrow)):
			newsheet[i, 1:3].merge(True)
		newsheet[4:9, 0].merge(True)	
		newsheet[10:expensesendrow, 0].merge(True)
		newsheet[expensesendrow+1:expensesendrow+6, 0].merge(True)				
		newsheet[expensesendrow+1:expensesendrow+3, 1].merge(True)		
		newsheet[expensesendrow+3:expensesendrow+6, 1].merge(True)			
		rowscount = len(newdatarows)
		newsheet[:rowscount, :columnscount].setDataArray(newdatarows)	
		newcontroller = newdoc.getCurrentController()
		selection = newdoc.getCurrentSelection()
		newcontroller.select(newsheet[2:rowscount, :columnscount])		
		drawTableBorders(xscriptcontext, newcontroller.getFrame())		
		newcontroller.select(selection)		
		datarange = newsheet[:rowscount, 3]	
		datarange.setPropertyValue("NumberFormat", commons.formatkeyCreator(newdoc)("#,##0;[BLUE]-#,##0"))	
		cellrangeobjects = newsheet[0, 0], newsheet[2, 0], newsheet[2, 3]  # 損益計算書、科目、金額、のセル。
		setCellRangeProperty(newdoc, (i.getRangeAddress() for i in cellrangeobjects), lambda x: x.setPropertyValue("HoriJustify", CENTER))
		cellrangeobjects = newsheet[4, 0], newsheet[10, 0]  # 売上原価、経費、のセル。
		setCellRangeProperty(newdoc, (i.getRangeAddress() for i in cellrangeobjects), lambda x: x.setPropertyValues(("VertJustify", "HoriJustify", "Orientation"), (CellVertJustify2.CENTER, CENTER, STACKED)))	
		cellrangeobjects = newsheet[expensesendrow+1, 0], newsheet[expensesendrow+1, 1], newsheet[expensesendrow+3, 1]  # 各種引当金・準備金等、繰戻額等、繰入額等、のセル。
		setCellRangeProperty(newdoc, (i.getRangeAddress() for i in cellrangeobjects), lambda x: x.setPropertyValue("VertJustify", CellVertJustify2.CENTER))
		cellrangeobjects = newsheet[1, 3], newsheet[6, 1], newsheet[8, 1], newsheet[9, 0], newsheet[expensesendrow-1, 1], newsheet[expensesendrow, 0], newsheet[expensesendrow+2, 2], newsheet[expensesendrow+5, 2]  # 現在日、差し引き、計、のセル。
		setCellRangeProperty(newdoc, (i.getRangeAddress() for i in cellrangeobjects), lambda x: x.setPropertyValue("HoriJustify", RIGHT))
		searchdescriptor = newsheet.createSearchDescriptor()
		searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
		searchdescriptor.setSearchString("^0")  # 0のセルを取得。戻り値はない。
		cellranges = datarange.findAll(searchdescriptor)  # 値のあるセルから0以外が入っているセル範囲コレクションを取得。見つからなかった時はNoneが返る。
		if cellranges:
			cellranges.clearContents(CellFlags.VALUE)  # 0のセルを空セルにする。	
		columns = newsheet.getColumns()
		columns[0].setPropertyValue("Width", 2000) 
		columns[1].setPropertyValue("Width", 2200) 
		columns[2].setPropertyValue("Width", 3500) 
		columns[3].setPropertyValue("Width", 5000)  # 金額列の列幅を設定。
	return addPL, createPL
def createTrialBalanceCreator(xscriptcontext, datetxts):  # 試算表の作成。
	datetxtforsheet, presentdatetxt, dummy = datetxts
	newdatarows = [("試算表", "", "", "", "", "", ""),\
				(datetxtforsheet, "", "", "", "", "", presentdatetxt),\
				("勘定科目", "期首残高", "", "期中取引", "", "期末残高", ""),\
				("", "借方", "貸方", "借方", "貸方", "借方", "貸方")]  # 新規シートのヘッダー行。	
	def addToTrialB(kamoku, sums):
		newdatarows.append((kamoku, *sums))
	def createTrialBalance(newdoc, pagewidth):
		newdatarows.append(("合計", *list(map(sum, islice(zip(*newdatarows[4:]), 1, None))),))  # 各列合計を取得。	
		newsheets = newdoc.getSheets()
		newsheet = newsheets[0]
		newsheet.setName("試算表")
		rowscount = len(newdatarows)
		columnscount = len(newdatarows[0])
		newsheet[:rowscount, :columnscount].setDataArray(newdatarows)
		horizontalmerges = 1, 3, 5  # 右隣のセルと結合するヘッダ行の列インデックス。	
		newsheet[0, :columnscount].merge(True)  # 題名セルの結合。	
		rangeaddresses = [newsheet[0, 0].getRangeAddress()]  # 中央揃えするセルのセルアドレスを入れるリストに題名セルを入れる。					。
		newsheet[2:4, 0].merge(True)  # 科目ヘッダーの結合。	
		newsheet[2, 0].setPropertyValue("VertJustify", CellVertJustify2.CENTER)  # 科目ヘッダーセルの縦中央揃え。
		[newsheet[2, i:i+2].merge(True) for i in horizontalmerges]  # 金額ヘッダーの結合。
		for i in horizontalmerges:  # 金額ヘッダーセルインデックスをイテレート。
			newsheet[2:4, i].merge(True)
			rangeaddresses.append(newsheet[2, i].getRangeAddress())
		setCellRangeProperty(newdoc, rangeaddresses, lambda x: x.setPropertyValue("HoriJustify", CENTER))
		datarange = newsheet[4:rowscount, 1:columnscount]		
		searchdescriptor = newsheet.createSearchDescriptor()
		searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
		searchdescriptor.setSearchString("^0")  # 0のセルを取得。戻り値はない。	
		cellranges = datarange.findAll(searchdescriptor)  # 値のあるセルから0が入っているセル範囲コレクションを取得。見つからなかった時はNoneが返る。
		if cellranges:
			cellranges.clearContents(CellFlags.VALUE)  # 0のセルを空セルにする。
		datarange.setPropertyValue("NumberFormat", commons.formatkeyCreator(newdoc)("#,##0;[BLUE]-#,##0"))	
		newsheet[1, columnscount-1].setPropertyValue("HoriJustify", RIGHT)
		newcontroller = newdoc.getCurrentController()	
		selection = newdoc.getCurrentSelection()
		newcontroller.select(newsheet[2:rowscount, :columnscount])		
		drawTableBorders(xscriptcontext, newcontroller.getFrame())	
		newcontroller.select(selection)	
		newkingakuwidth = 2000  # 科目金額列幅。
		newsheet[0, 1:columnscount].getColumns().setPropertyValue("Width", newkingakuwidth)  # 金額列の列幅を設定。
		newsheet.getColumns()[0].setPropertyValue("Width", pagewidth-newkingakuwidth*(columnscount-1))  # 科目列幅を設定。残った幅をすべて割り当てる。	
	return addToTrialB, createTrialBalance
def createBalanceSheetCreator(xscriptcontext, datetxts):  # 賃借対照表の作成。
	datetxtforsheet, presentdatetxt, datetxtsforBS = datetxts
	barancesheetrows = [("賃借対照表", "", "", "", "", ""),\
						(datetxtforsheet, "", "", "", "", presentdatetxt),\
						("資産の部", "", "", "負債・資本の部", "", ""),\
						("科目", *datetxtsforBS, "科目", *datetxtsforBS)]
	balancesheetkarikata = []  # 賃借対照表の借方の(科目, 期首金額, 期末金額)のタプルを入れるリスト。
	balancesheetkashikata = []  # 賃借対照表の貸方の(科目, 期首金額, 期末金額)のタプルを入れるリスト。
	shotoku = 0
	jigyonushikashirow = "事業主貸", 0, 0
	jigyonushikarirow = "事業主借", 0, 0
	motoirerow = "元入金", 0, 0
	def addToBS(kubun, kamoku, sums):
		nonlocal shotoku, jigyonushikashirow, jigyonushikarirow, motoirerow
		if kubun.startswith("資産"):
			if kamoku=="事業主貸":
				jigyonushikashirow = kamoku, 0, sums[4]
			else:
				balancesheetkarikata.append((kamoku, sums[0], sums[4])) # 賃借対照表の借方の(科目, 期首金額, 期末金額)を取得。
		elif kubun.startswith("負債"):
			if kamoku=="事業主借":
				jigyonushikarirow = kamoku, 0, sums[5]
			elif kamoku=="元入金":
				motoirerow = kamoku, sums[1], sums[5]  # 賃借対照表の元入金の期首と期末の額は同一。
			else:
				balancesheetkashikata.append((kamoku, sums[1], sums[5]))  # 賃借対照表の貸方の(科目, 期首金額, 期末金額)を取得。
		elif kubun=="経費":
			shotoku -= sums[4]
		elif kubun=="収益":
			shotoku += sums[5]	
	def createBalanceSheet(newdoc, pagewidth):	
		newcontroller = newdoc.getCurrentController()		
		newsheets = newdoc.getSheets()
		barancesheetrows.extend(i[0]+i[1] for i in zip_longest(balancesheetkarikata, balancesheetkashikata, fillvalue=("", 0, 0)))  # (借方科目, 期首金額, 期末金額, 貸方科目, 期首金額, 期末金額)をイテレート。chainやchain.from_iterableだとイテレーターのリストが返るのでさらに展開しないといけない。
		barancesheetrows.extend((("", 0, 0, *jigyonushikarirow),\
								("", 0, 0, *motoirerow),\
								(*jigyonushikashirow, "所得金額", 0, shotoku)))
		barancesheetrows.append(("合計", *list(map(sum, islice(zip(*barancesheetrows[4:]), 1, 3))), "合計", *list(map(sum, islice(zip(*barancesheetrows[4:]), 4, 6)))))
		newsheetname = "賃借対照表"
		newsheets.insertNewByName(newsheetname, len(newsheets))
		newsheet = newsheets[newsheetname]
		rowscount = len(barancesheetrows)
		columnscount = len(barancesheetrows[0])
		newsheet[:rowscount, :columnscount].setDataArray(barancesheetrows)	
		newsheet[0, :columnscount].merge(True)  # 題名セルの結合。	
		newsheet[2, :3].merge(True)
		newsheet[2, 3:6].merge(True)
		cellrangeobjects = newsheet[0, 0], newsheet[2, 0], newsheet[2, 3]
		setCellRangeProperty(newdoc, (i.getRangeAddress() for i in cellrangeobjects), lambda x: x.setPropertyValue("HoriJustify", CENTER))
		cellrangeobjects = newsheet[1, columnscount-1], newsheet[3, 1:3], newsheet[3, 4:6]
		setCellRangeProperty(newdoc, (i.getRangeAddress() for i in cellrangeobjects), lambda x: x.setPropertyValue("HoriJustify", RIGHT))
		selection = newdoc.getCurrentSelection()
		newcontroller.select(newsheet[2:rowscount, :columnscount])		
		drawTableBorders(xscriptcontext, newcontroller.getFrame())		
		newcontroller.select(selection)	
		newkingakuwidth = 2700
		newsheet[0, 1:3].getColumns().setPropertyValue("Width", newkingakuwidth)  # 金額列の列幅を設定。
		newsheet[0, 4:6].getColumns().setPropertyValue("Width", newkingakuwidth)  # 金額列の列幅を設定。
		kamokuwidth = (pagewidth-newkingakuwidth*4)//2
		columns = newsheet.getColumns()
		columns[0].setPropertyValue("Width", kamokuwidth)
		columns[3].setPropertyValue("Width", kamokuwidth)
		cellrangeobjects = 	newsheet[4:rowscount, 1:3], newsheet[4:rowscount, 4:6] 
		setCellRangeProperty(newdoc, (i.getRangeAddress() for i in cellrangeobjects), lambda x: x.setPropertyValue("NumberFormat", commons.formatkeyCreator(newdoc)("#,##0;[BLUE]-#,##0")))
		cellrangeobjects = newsheet[4:rowscount-2, 1:3], newsheet[4:rowscount-4, 4:6]  # 事業主貸、合計、事業主借、元入金、所得の金額を除いた金額欄は0は空セルにする。
		searchdescriptor = newsheet.createSearchDescriptor()
		searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
		searchdescriptor.setSearchString("^0")  # 0のセルを取得。戻り値はない。	
		cellranges = newdoc.createInstance("com.sun.star.sheet.SheetCellRanges")  
		cellranges.addRangeAddresses((i.getRangeAddress() for i in cellrangeobjects), False)			
		resulutcellranges = cellranges.queryContentCells(CellFlags.VALUE).findAll(searchdescriptor)  # 値のあるセルから0以外が入っているセル範囲コレクションを取得。見つからなかった時はNoneが返る。
		if resulutcellranges:
			resulutcellranges.clearContents(CellFlags.VALUE)  # 0のセルを空セルにする。	
		cellrangeobjects = newsheet[rowscount-2, 1], newsheet[rowscount-4, 4], newsheet[rowscount-2, 4]    # 期首事業主貸、期首事業主借、期首所得の金額のセル。
		setCellRangeProperty(newdoc, (i.getRangeAddress() for i in cellrangeobjects), lambda x: x.setPropertyValues(("DiagonalBLTR", "CharColor"), (BorderLine2(LineWidth=10, Color=commons.COLORS["black"]), commons.COLORS["white"])))
	return addToBS, createBalanceSheet
def createShiwakeCho(xscriptcontext, txt):
	newdoc = xscriptcontext.getDesktop().loadComponentFromURL("private:factory/scalc", "_blank", 0, ())  # 新規ドキュメントの取得。	
	indicator = newdoc.getCurrentController().getFrame().createStatusIndicator()  # 新規ドキュメントのステータスインディケーターを取得。
	indicator.start("{}中".format(txt), VARS.emptyrow)	
	doc = xscriptcontext.getDocument()	
	datetxtforfile, datetxtforsheet, presentdatetxt, dummy = getDaytxts()
	sheet = VARS.sheet
	daycolumn = VARS.daycolumn
	slipnocolumn = daycolumn - 1
	tekiyocolumn = daycolumn + 1
	splittedcolumn = VARS.splittedcolumn		
	kozakamokuname = "仕訳日記帳"
	newkingakucolumns = 2, 4  # 金額書式にする列インデックスのタプル。
	newtekiyocolumn = 5  # 摘要列インデックス。
	newkamokucolumnidxes = 1, 3  # 科目列インデックスのタプル。
	newheadermergecolumns = 2, 4, 5  # セル結合するヘッダー行の列インデックスのタプル。				
	headerrows, datarows = getDataRows(xscriptcontext)
	if not headerrows:
		newdoc.close(True)	
		return
	newdatarows = [(kozakamokuname, "", "", "", "", ""),\
				(datetxtforsheet, "", "", "", "", presentdatetxt),\
				("日付", "借方科目", "借方金額", "貸方科目", "貸方金額", "摘要"),\
				("伝票番号", "借方補助科目", "", "貸方補助科目", "", "")]  # 新規シートのヘッダー行。
	slipstartrows = []  # 新規シートの伝票開始行インデックスのリスト。
	datevalue = ""  # 伝票の日付シリアル値。	
	startrowi = VARS.splittedrow - 1
	totalslipcount = VARS.emptyrow - VARS.splittedrow
	for i, datarow in enumerate(datarows[VARS.splittedrow:], start=VARS.splittedrow):  # 伝票行を行インデックスと共にイテレート。
		indicator.setText("伝票処理 {:>4}/{}".format(i-startrowi, totalslipcount))	
		indicator.setValue(i)
		slipstartrows.append(len(newdatarows))  # 新規シートの伝票開始行インデックスを取得。
		datevalue = "" if datevalue==datarow[daycolumn] else datarow[daycolumn]  # 前の伝票と日付が異なる時のみ日付を表示する。
		daycolumns = [datevalue, datarow[slipnocolumn]]  # 新規シートの日付列のデータのリスト。伝票の開始行に日付、その下行に伝票番号を表示。
		karikatakamokus = []  # 借方科目列のデータのリスト。
		karikatas = []  # 借方金額列のデータのリスト。		
		karikatatekiyo = []  # 借方摘要列のデータのリスト。				
		kashikatakamokus = []  # 貸方科目列のデータのリスト。		
		kashikatas = []  # 貸方金額列のデータのリスト。		
		kashikatatekiyo = []  # 貸方摘要列のデータのリスト。
		for j in compress(zip(*headerrows, datarow[splittedcolumn:]), datarow[splittedcolumn:]):  # 空文字や0でないセルが入っている列の行データを列インデックスとヘッダー行と共にイテレート。
			annotation = sheet[i, j[0]].getAnnotation().getString().strip()  # 伝票行のこの列のセルのコメントを取得。空白文字を削除する。
			if j[4]>0:  # 金額が正の科目は借方。
				karikatakamokus.extend(j[2:4])
				karikatas.extend(["", j[4]])	
				karikatatekiyo.extend([annotation, ""])		
			else:  # 金額が負の科目は貸方。
				kashikatakamokus.extend(j[2:4])
				kashikatas.extend(["", -j[4]])
				kashikatatekiyo.extend([annotation, ""])									
		gene = zip_longest(daycolumns, karikatakamokus, karikatas, kashikatakamokus, kashikatas, [datarow[tekiyocolumn]], karikatatekiyo, kashikatatekiyo, fillvalue="")  # 各列を1要素ずつイテレートして1行にする。	
		for k in gene:
			newdatarows.append([*k[:-3], "/".join([str(m) for m in k[-3:] if m])])  # 摘要は/で結合する。
	slipstartrows.append(len(newdatarows))  # 新規シートのデータ終了行の下行インデックスを取得。		
	if slipstartrows[0]==slipstartrows[-1]:  # 伝票がない時は何もしない。
		commons.showErrorMessageBox(doc.getCurrentController(), "伝票が一つもありません。\n処理を中止します。")	
		newdoc.close(True)					
		return
	indicator.setText("仕訳日記帳を描画中")	
	newdocname = "仕訳日記帳_{}_{}.ods".format(datetxtforfile, datetime.now().strftime("%Y%m%d%H%M%S"))
	createNewSheetCreator(newdoc, newkamokucolumnidxes, newkingakucolumns, newheadermergecolumns, newtekiyocolumn)(kozakamokuname, newdatarows, slipstartrows)
	indicator.setText("ファイルに保存中 {}".format(newdocname))	
	saveNewDoc(doc, newdoc, newdocname)		
	indicator.end()  # reset()の前にend()しておかないと元に戻らない。
	indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。		
def createHojoMotoCho(xscriptcontext, txt, docname, hojokamokuindexgenefunc):
	newdoc = xscriptcontext.getDesktop().loadComponentFromURL("private:factory/scalc", "_blank", 0, ())  # 新規ドキュメントの取得。	
	indicator = newdoc.getCurrentController().getFrame().createStatusIndicator()  # 新規ドキュメントのステータスインディケーターを取得。			
	indicator.start("{}中".format(txt), VARS.emptycolumn)	
	doc = xscriptcontext.getDocument()	
	datetxtforfile, datetxtforsheet, *dummy = getDaytxts()
	newheadermergecolumns = 2, 3, 4, 5  # セル結合するヘッダー行の列インデックスのタプル。
	newkingakucolumns = 3, 4, 5  # 金額書式にする列インデックスのタプル。
	newtekiyocolumn = 2  # 摘要列インデックス。
	newkamokucolumnidxes = 1,  # 科目列インデックスのタプル。	
	headerrows, datarows = getDataRows(xscriptcontext)	
	if not headerrows:
		newdoc.close(True)	
		return
	createNewSheet = createNewSheetCreator(newdoc, newkamokucolumnidxes, newkingakucolumns, newheadermergecolumns, newtekiyocolumn)		
	createHojoSheet = createHojoSheetCreator(datetxtforsheet, headerrows, datarows, createNewSheet)	
	indicator.setText("各科目を処理中")
	for k in hojokamokuindexgenefunc(headerrows):
		indicator.setText("{:　>10}を処理中".format(headerrows[VARS.kamokurow][k]))
		indicator.setValue(k)
		createHojoSheet(k)
	if len(newdoc.getSheets())==1:  # シートが増えていない時。
		commons.showErrorMessageBox(doc.getCurrentController(), "伝票がある科目が一つもありませんでした。")	
		newdoc.close(True)				
		return										
	newdocname = "{}_{}_{}.ods".format(docname, datetxtforfile, datetime.now().strftime("%Y%m%d%H%M%S"))
	indicator.setText("ファイルに保存中 {}".format(newdocname))	
	saveNewDoc(doc, newdoc, newdocname)	
	indicator.end()  # reset()の前にend()しておかないと元に戻らない。
	indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。	
def getDaytxts():  # 帳簿に必要な日付文字列を取得。
	sdate, edate = getDateSection()  # 期首日と期末日を取得。
	if not sdate:  # 期首日と期末日が取得できていないときは空文字を返す。
		return "", "", ""
	startdaytxt, enddaytxt = sdate.isoformat(), edate.isoformat()
	datetxtforsheet = "自: {} 至: {}".format(startdaytxt, enddaytxt)
	datetxtforfile = "{}-{}".format(startdaytxt.replace("-", ""), enddaytxt.replace("-", ""))
	datetxtsforBS = "{}月{}日(期首)".format(sdate.month, sdate.day), "{}月{}日(期末)".format(edate.month, edate.day)
	todaydate = date.today()
	if todaydate<edate:  # 今日が期末日より前の時は今日にする。
		edate = todaydate
	presentdatetxt = "({}現在)".format(edate.isoformat())  # 期末日、または、今日がその前なら今日の日付。	
	return datetxtforfile, datetxtforsheet, presentdatetxt, datetxtsforBS
def createMotoCho(xscriptcontext, txt, docname, kozakamokunamegenefunc):  # xscriptcontext, ステータスバーの表示文字列、帳簿ファイル名の元、口座科目名のイテレーターを返す関数。
	newdoc = xscriptcontext.getDesktop().loadComponentFromURL("private:factory/scalc", "_blank", 0, ())  # 新規ドキュメントの取得。	
	indicator = newdoc.getCurrentController().getFrame().createStatusIndicator()  # 新規ドキュメントのステータスインディケーターを取得。			
	indicator.start("{}中".format(txt), VARS.emptycolumn) 
	doc = xscriptcontext.getDocument()	
	datetxtforfile, datetxtforsheet, *dummy = getDaytxts()	
	newkingakucolumns = 3, 4, 5  # 金額書式にする列インデックスのタプル。
	newtekiyocolumn = 2  # 摘要列インデックス。
	newkamokucolumnidxes = 1,  # 科目列インデックスのタプル。
	newheadermergecolumns = 2, 4, 5  # セル結合するヘッダー行の列インデックスのタプル。
	headerrows, datarows = getDataRows(xscriptcontext)	
	if not headerrows:
		newdoc.close(True)				
		return
	createNewSheet = createNewSheetCreator(newdoc, newkamokucolumnidxes, newkingakucolumns, newheadermergecolumns, newtekiyocolumn)
	createKamokuSheet = createKamokuSheetCreator(datetxtforsheet, headerrows, datarows, createNewSheet)
	for i, kozakamokuname in enumerate(kozakamokunamegenefunc(datarows), start=VARS.splittedcolumn):  # 口座科目名をイテレート。科目行の空セルでない値のみイテレート。
		indicator.setText("{:　>10}を処理中".format(kozakamokuname))
		indicator.setValue(i)
		createKamokuSheet(kozakamokuname)
	if len(newdoc.getSheets())==1:  # シートが増えていない時。
		commons.showErrorMessageBox(doc.getCurrentController(), "伝票がある科目が一つもありませんでした。")	
		newdoc.close(True)				
		return			
	newdocname = "{}_{}_{}.ods".format(docname, datetxtforfile, datetime.now().strftime("%Y%m%d%H%M%S"))
	indicator.setText("ファイルに保存中 {}".format(newdocname))	
	saveNewDoc(doc, newdoc, newdocname)	
	indicator.end()  # reset()の前にend()しておかないと元に戻らない。
	indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。		
def drawTableBorders(xscriptcontext, frame):  # 選択範囲内すべてに罫線を引く。UNO APIでやる方法がわからない。線種の設定方法も不明。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
	propertyvalues = PropertyValue(Name="OuterBorder.LeftBorder", Value=(0,0,2,0,0,2)),\
					PropertyValue(Name="OuterBorder.LeftDistance", Value=0),\
					PropertyValue(Name="OuterBorder.RightBorder", Value=(0,0,2,0,0,2)),\
					PropertyValue(Name="OuterBorder.RightDistance", Value=0),\
					PropertyValue(Name="OuterBorder.TopBorder", Value=(0,0,2,0,0,2)),\
					PropertyValue(Name="OuterBorder.TopDistance", Value=0),\
					PropertyValue(Name="OuterBorder.BottomBorder", Value=(0,0,2,0,0,2)),\
					PropertyValue(Name="OuterBorder.BottomDistance", Value=0),\
					PropertyValue(Name="InnerBorder.Horizontal", Value=(0,0,2,0,0,2)),\
					PropertyValue(Name="InnerBorder.Vertical", Value=(0,0,2,0,0,2)),\
					PropertyValue(Name="InnerBorder.Flags", Value=0),\
					PropertyValue(Name="InnerBorder.ValidFlags", Value=127),\
					PropertyValue(Name="InnerBorder.DefaultDistance", Value=0)  # InnerBorder.ValidFlagsが枠線の引く場所を指定と思われる。		
	dispatcher.executeDispatch(frame, ".uno:SetBorderStyle", "", 0, propertyvalues)  # ディスパッチコマンドで罫線を引く。
def createHojoSheetCreator(settlingdaytxt, headerrows, datarows, createNewSheet):
	splittedrow = VARS.splittedrow
	daycolumn = VARS.daycolumn
	slipnocolumn = daycolumn - 1
	tekiyocolumn = daycolumn + 1
	splittedcolumn = VARS.splittedcolumn
	sheet = VARS.sheet
	def createHojoSheet(k):
		kozakamokuname = "{}_{}".format(headerrows[2][k], headerrows[3][k]) if headerrows[3][k] else headerrows[2][k]
		kubun = headerrows[1][k]
		sign = -1 if kubun.startswith(("負債", "収益")) else 1  # 区分が負債または収益から始まっている時は残高は貸方を正とする。	
		newdatarows = [(kozakamokuname, "", "", "", "", ""),\
					(settlingdaytxt, "", "", "", "", kubun),\
					("日付", "相手勘定科目", "摘要", "借方金額", "貸方金額", "残高"),\
					("伝票番号", "相手補助科目", "", "", "", "")]  # 新規シートのヘッダー行。
		slipstartrows = []  # 新規シートの伝票開始行インデックスのリスト。
		datevalue = ""  # 伝票の日付シリアル値。
		zandaka = 0						
		for i, datarow in enumerate(datarows[splittedrow:], start=splittedrow):  # 伝票行を行インデックスと共にイテレート。
			if datarow[headerrows[0][k]]:  # 口座科目の列に値がある時のみ。
				slipstartrows.append(len(newdatarows))  # 新規シートの伝票開始行インデックスを取得。
				datevalue = "" if datevalue==datarow[daycolumn] else datarow[daycolumn]  # 前の伝票と日付が異なる時のみ日付を表示する。
				daycolumns = [datevalue, datarow[slipnocolumn]]  # 新規シートの日付列のデータのリスト。伝票の開始行に日付、その下行に伝票番号を表示。
				aitekamokus = []  # 相手科目列の行データリスト。。相手勘定科目と相手補助科目のペア。
				tekiyos = []  # この伝票の摘要列の行データリスト。
				karikatas = []  # この伝票の借方金額列の行データリスト。補助科目と借方金額のペア。
				kashikatas = []  # この伝票の貸方金額列の行データのリスト。
				zandakas = []  # この伝票の残高列の行のリスト。
				for j in compress(zip(*headerrows, datarow[splittedcolumn:]), datarow[splittedcolumn:]):  # 空文字や0でないセルが入っている列の行データを列インデックスとヘッダー行と共にイテレート。
					if j[0]==headerrows[0][k]:  # 帳簿にする科目名の時。
						annotation = sheet[i, j[0]].getAnnotation().getString().strip()  # 伝票行のこの列のセルのコメントを取得。空白文字を削除する。
						tekiyos = (annotation, "") if annotation else (datarow[tekiyocolumn], "")  # コメントがあるときはそれを摘要にする。
						if j[4]>0:  # 金額が正は借方。
							karikatas.extend(["", j[4]])	
						else:  # 金額が負は貸方。
							kashikatas.extend(["", -j[4]])					
						zandaka += j[4]*sign  # 残高を計算。
						zandakas = "", zandaka
					else:  # 口座科目でない時。
						if not aitekamokus:  # 相手科目がまだ未設定の時。
							aitekamokus = j[2:4]  # 相手勘定科目と相手補助科目を取得。					
						elif "諸口" not in aitekamokus: 
							aitekamokus = "諸口", ""  # 相手科目が複数の時の相手科目は諸口にする。
				newdatarows.extend(zip_longest(daycolumns, aitekamokus, tekiyos, karikatas, kashikatas, zandakas, fillvalue=""))  # 各列を1要素ずつイテレートして1行にする。	
		slipstartrows.append(len(newdatarows))  # 新規シートのデータ終了行の下行インデックスを取得。		
		if slipstartrows[0]==slipstartrows[-1]:  # 伝票がない時は何もしない。
			return
		createNewSheet(kozakamokuname, newdatarows, slipstartrows)		
	return createHojoSheet
def createKamokuSheetCreator(settlingdaytxt, headerrows, datarows, createNewSheet):
	splittedrow = VARS.splittedrow
	daycolumn = VARS.daycolumn
	slipnocolumn = daycolumn - 1
	tekiyocolumn = daycolumn + 1
	splittedcolumn = VARS.splittedcolumn
	sheet = VARS.sheet
	def createKamokuSheet(kozakamokuname):
		kozacolumns = []  # 口座科目の列インデックスのリスト。
		i = 0
		while kozakamokuname in headerrows[2][i:]:
			i = headerrows[2].index(kozakamokuname, i)
			kozacolumns.append(headerrows[0][i])  # 補助科目の列インデックスを取得。
			i += 1	
		kubun = headerrows[1][kozacolumns[0]-splittedcolumn]  # 区分を取得。	
		sign = -1 if kubun.startswith(("負債", "収益")) else 1  # 区分が負債または収益から始まっている時は残高は貸方を正とする。	
		newdatarows = [(kozakamokuname, "", "", "", "", ""),\
					(settlingdaytxt, "", "", "", "", kubun),\
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
					if j[2]==kozakamokuname:  # 科目名が口座科目名の時。
						annotation = sheet[i, j[0]].getAnnotation().getString().strip()  # 伝票行のこの列のセルのコメントを取得。空白文字を削除する。
						tekiyos = (annotation, "") if annotation else (str(datarow[tekiyocolumn]), "")  # コメントがあるときはそれを摘要にする。
						if j[4]>0:  # 金額が正は借方。補助科目名も使う。
							karikatas.extend([j[3], j[4]])	
						else:  # 金額が負は貸方。
							karikatas.extend([j[3], ""])	
							kashikatas.extend(["", -j[4]])					
						zandaka += j[4]*sign  # 残高を計算。
						zandakas = "", zandaka
					else:  # 口座科目でない時。
						if not aitekamokus:  # 相手科目がまだ未設定の時。
							aitekamokus = j[2:4]  # 相手勘定科目と相手補助科目を取得。					
						elif "諸口" not in aitekamokus: 
							aitekamokus = "諸口", ""  # 相手科目が複数の時の相手科目は諸口にする。
				newdatarows.extend(zip_longest(daycolumns, aitekamokus, tekiyos, karikatas, kashikatas, zandakas, fillvalue=""))  # 各列を1要素ずつイテレートして1行にする。	
		slipstartrows.append(len(newdatarows))  # 新規シートのデータ終了行の下行インデックスを取得。		
		if slipstartrows[0]==slipstartrows[-1]:  # 伝票がない時は何もしない。		
			return
		createNewSheet(kozakamokuname, newdatarows, slipstartrows)		
	return createKamokuSheet
def saveNewDoc(doc, newdoc, newdocname):
	sheets = newdoc.getSheets()
	if "Sheet1" in sheets:  # デフォルトシートが残っている時。
		if len(sheets)==1:  # デフォルトシート以外一つもシートが追加されていない時シートは保存せずに閉じる。
			newdoc.close(True)
			controller = doc.getCurrentController()	
			commons.showErrorMessageBox(controller, "生成されたシートがありませんでした。")	
			return		
		else:  # 複数シートが存在しSheet1が残っている時。
			del sheets["Sheet1"]  # 新規ドキュメントのデフォルトシートを削除する。 	
	dirpath = os.path.dirname(unohelper.fileUrlToSystemPath(doc.getURL()))  # このドキュメントのあるディレクトリのフルパスを取得。
	systempath = os.path.join(dirpath, "帳簿", newdocname)  # 新規ドキュメントのフルパスを取得。
	if os.path.exists(systempath):  # すでにファイルが存在する時。
		msg = "{}はすでに存在します。\n上書きしますか？".format(newdocname)
		componentwindow = doc.getCurrentController().ComponentWindow
		msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "WEntryBook", msg)
		if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
			return
	newdoc.storeAsURL(unohelper.systemPathToFileUrl(systempath), ())  # 新規ドキュメントを保存。	
def createNewSheetCreator(newdoc, newkamokucolumnidxes, newkingakucolumns, newheadermergecolumns, newtekiyocolumn):		
	newdatewidth = 1500  # 日付列幅。1/100mm。
	newkamokuwidth = 3500  # 科目列幅。
	newkingakuwidth = 2500  # 科目金額列幅。		
	borderline = BorderLine2(LineWidth=10, Color=commons.COLORS["black"])
	tableborder2 = TableBorder2(TopLine=borderline, LeftLine=borderline, RightLine=borderline, BottomLine=borderline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)	
	createFormatKey = commons.formatkeyCreator(newdoc)
	newsheets = newdoc.getSheets()  # 新規ドキュメントのシートコレクションを取得。	
	pagestyle = newdoc.getStyleFamilies()["PageStyles"]["Default"]
	width, leftmargin, rightmargin = pagestyle.getPropertyValues(("Width", "LeftMargin", "RightMargin"))
	pagewidth = width - leftmargin - rightmargin  # 印刷幅を1/100mmで取得。	
	insertPageBreaks = insertPageBreaksCreator(pagestyle)
	pagestyle.setPropertyValue("HeaderIsOn", False)  # 印刷時ヘッダーを付けない。
	def createNewSheet(kozakamokuname, newdatarows, slipstartrows):  # 新規シートを挿入してデータを代入して書式設定する。
		newsheets.insertNewByName(kozakamokuname, len(newsheets))  # 口座科目名のシートを新規ドキュメントに挿入。
		newsheet = newsheets[kozakamokuname]  # 新規シートを取得。
		newdatarange = newsheet[:len(newdatarows), :len(newdatarows[0])]
		newdatarange.setDataArray(newdatarows)  # 新規シートに代入。		
		columncount = len(newdatarows[0])  # 表の列数。	
		newsheet[0, :columncount].merge(True)  # 題名セルと結合。			
		newsheet[0, 0].setPropertyValue("HoriJustify", CENTER)  # 題名セルを中央揃え。
		newsheet[1, columncount-1].setPropertyValue("HoriJustify", RIGHT)  # 区分セルを右揃え。
		setCellRangeProperty(newdoc, (newsheet[i, 0].getRangeAddress() for i in slipstartrows[:-1]), lambda x: x.setPropertyValues(("HoriJustify", "NumberFormat"), (LEFT, createFormatKey("M/D"))))  # 伝票開始列の日付セルのプロパティ設定。
		setCellRangeProperty(newdoc, (newsheet[j, i].getRangeAddress() for i in newkingakucolumns for j in range(slipstartrows[0]+1, slipstartrows[-1], 2)), lambda x: x.setPropertyValue("NumberFormat", createFormatKey("#,##0;[BLUE]-#,##0")))  # 金額列の書式設定。
		for i in newheadermergecolumns:  # セル結合するヘッダー行。
			newsheet[2:4, i].merge(True)
			newsheet[2, i].setPropertyValue("VertJustify", CellVertJustify2.CENTER)
		rangeaddresses = []  # 摘要セルのセルアドレスを入れるリスト。
		for i in range(slipstartrows[0], slipstartrows[-1], 2):  # 1行おきに行インデックスをイテレート。
			newsheet[i:i+2, newtekiyocolumn].merge(True)  # 摘要列を2行ずつ結合。
			rangeaddresses.append(newsheet[i, newtekiyocolumn].getRangeAddress())
		setCellRangeProperty(newdoc, rangeaddresses, lambda x: x.setPropertyValues(("VertJustify", "IsTextWrapped"), (CellVertJustify2.CENTER, True)))  # 摘要列を上下中央に、折り返し有効。
		newsheet[slipstartrows[0]:slipstartrows[-1], :].getRows().setPropertyValue("OptimalHeight", True)  # 内容を折り返した後の行の高さを調整。
		setCellRangeProperty(newdoc, (newsheet[i:i+2, j].getRangeAddress() for i in range(slipstartrows[0], slipstartrows[-1], 2) for j in range(1, columncount)), lambda x: x.setPropertyValue("TableBorder2", tableborder2))  # 日付列の枠線を引く。			
		setCellRangeProperty(newdoc, (newsheet[2:4, j].getRangeAddress() for j in range(columncount)), lambda x: x.setPropertyValue("TableBorder2", tableborder2))  # 表の1行目の罫線を引く。					
		setCellRangeProperty(newdoc, (newsheet[i:j, 0].getRangeAddress() for i, j in zip(slipstartrows[:-1], slipstartrows[1:])), lambda x: x.setPropertyValue("TableBorder2", tableborder2))  # 1列目の罫線を伝票区切りで引く。	
		columns = newsheet.getColumns()  # 新規シートの列アクセスオブジェクト。
		for i, j in chain(zip(newkamokucolumnidxes, (newkamokuwidth,)*len(newkamokucolumnidxes)), zip(newkingakucolumns, (newkingakuwidth,)*len(newkingakucolumns))):
			columns[i].setPropertyValue("Width", j)  # 列幅を設定。
		columns[0].setPropertyValue("Width", newdatewidth)  # 日付列幅を設定。
		columns[newtekiyocolumn].setPropertyValue("Width", pagewidth-newdatewidth-newkamokuwidth*len(newkamokucolumnidxes)-newkingakuwidth*len(newkingakucolumns))  # 摘要列幅を設定。残った幅をすべて割り当てる。	
		insertPageBreaks(newsheet, newdatarange.getRows())
	return createNewSheet
def setCellRangeProperty(doc, rangeaddresses, setProperty):
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  
	cellranges.addRangeAddresses(rangeaddresses, False)	
	setProperty(cellranges)
def sortSlips(dispatcher, doc):  # 日付順に伝票をソート。
	controller = doc.getCurrentController()	
	selection = doc.getCurrentSelection()  # 選択範囲をここで保存しておく。
	controller.select(VARS.sheet[VARS.splittedrow:, :])  # ソートするセル範囲を選択。固定行以下すべてを選択。
	props = PropertyValue(Name="Col1", Value=VARS.daycolumn+1),  # 日付順にソート。Col1の番号は優先順位。Valueはインデックス+1。 
	dispatcher.executeDispatch(controller.getFrame(), ".uno:DataSort", "", 0, props)  # ディスパッチコマンドでソート。
	controller.select(selection)  # 元のセルを選択し直す。		
def getDataRows(xscriptcontext):
	if VARS.splittedrow<VARS.emptyrow:  # 伝票行がある時のみ。
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
		dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
		doc = xscriptcontext.getDocument()
		sortSlips(dispatcher, doc)  # 日付順に伝票をソート。			
		datarows = VARS.sheet[:VARS.emptyrow, :VARS.emptycolumn].getDataArray()  # 全データ行を取得。		
		msg = ""
		if not datarows[VARS.kamokurow][VARS.splittedcolumn]:  # 科目行先頭列のセルがTrueでない時。
			msg = "科目行の先頭セルには科目名が入っていないといけません。"	
		else:
			gene = zip(*datarows[VARS.splittedrow:])  # 固定列行以下の列のデータのイテレーター。
			if any(filter(None, next(gene))):  # 伝票内計が0か空セル以外の値をイテレート。
				msg = "貸方と借方が一致しない行があります。"
			elif "" in next(gene):  # 伝票番号列に空セルがある時。
				msg = "伝票番号のない行があります。"
			else:
				days = next(gene)  # 伝票の取引日列のタプルを取得。
				if "" in days:  # 取引日列に空セルにある時。
					msg = "取引日のない行があります。"
				else:
					dates = getDateSection()  # 期首日と期末日を取得。
					if all(dates):  # 決算日がある時。
						functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
						sday, eday = [functionaccess.callFunction("DATE", (i.year, i.month, i.day)) for i in dates]
						if days[0]<sday or eday<days[-1]:
							msg = "会計年度外の日付の行があります。"
	else:
		msg = "処理する伝票行がありません。"
	if msg:
		commons.showErrorMessageBox(doc.getCurrentController(), "{}\n処理を中止します。".format(msg))	
		return ("",)*2		
	headerrows = generateHeaderRows(datarows[:VARS.kamokurow+2])
	return headerrows, datarows
def generateHeaderRows(datarows):  # 列インデックス、区分、科目、補助科目、の行のタプルを空セルを埋めて返す。
	kubuns = []  # 科目行の上の区分行。
	[kubuns.append(i if i else kubuns[-1]) for i in datarows[VARS.kamokurow-1][VARS.splittedcolumn:]]  # 区分行をすべて埋める。				
	kamokus = []
	[kamokus.append(i if i else kamokus[-1]) for i in datarows[VARS.kamokurow][VARS.splittedcolumn:]]  # 科目行をすべて埋める。
	return range(VARS.splittedcolumn, VARS.emptycolumn), kubuns, kamokus, datarows[VARS.kamokurow+1][VARS.splittedcolumn:]  # 列インデックス行, 区分行、科目行、補助科目行。	
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	selection = eventobject.Source.getSelection()	
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
		sheet = selection.getSpreadsheet()
		VARS.setSheet(sheet)		
		drawBorders(selection)  # 枠線の作成。
def drawBorders(selection):  # ターゲットを交点とする行列全体の外枠線を描く。
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
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。	
	contextmenuname, addMenuentry, baseurl, selection = commons.contextmenuHelper(VARS, contextmenuexecuteevent, xscriptcontext)
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上角のセルのアドレスを取得。
	r, c  = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。	
	sheet = VARS.sheet
	if contextmenuname=="cell":  # セルのとき。セル範囲も含む。
		if VARS.splittedcolumn<=c<VARS.emptycolumn:  # 科目行か補助科目行に値がある列の時。
			datarows = sheet[VARS.kamokurow:VARS.kamokurow+2, c].getDataArray()  # 科目行と補助科目行を取得。
			kamoku = datarows[0][0]
			hojokamoku = datarows[1][0]			
			if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 単独セルの時のみ。
				if r==VARS.kamokurow and kamoku:  # 科目行かつ科目行に値があるとき。
					addMenuentry("ActionTrigger", {"Text": "{}の勘定元帳生成".format(kamoku), "CommandURL": baseurl.format("entry2")}) 
				elif r==VARS.kamokurow+1 and hojokamoku:  # 補助科目行かつ補助科目行に値があるとき。:
					addMenuentry("ActionTrigger", {"Text": "{}の補助元帳生成".format(hojokamoku), "CommandURL": baseurl.format("entry3")}) 	
				elif r>=VARS.splittedrow:  # 固定行より下の時。
					if r<=VARS.emptyrow:  # 取引日列が入力済の時。
						if sheet[r, VARS.sliptotalcolumn].getValue()!=0:  # 伝票内計が0でない時のみ。空セルや文字列は0が返る。
							txt = hojokamoku if hojokamoku else kamoku  # 補助科目行に値がある時は補助科目行、ないときは科目行の値を使う。
							addMenuentry("ActionTrigger", {"Text": "「{}」で決済".format(txt), "CommandURL": baseurl.format("entry5")}) 
							if txt!="現金":  # 現金列でない時のみ。
								addMenuentry("ActionTrigger", {"Text": "「現金」で決済", "CommandURL": baseurl.format("entry4")}) 
							addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})		
					if selection.getValue()!=0:  # セルの値が0でない時。
						addMenuentry("ActionTrigger", {"Text": "正負反転", "CommandURL": baseurl.format("entry10")}) 		
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertAnnotation"})	
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:EditAnnotation"})	
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteNote"})	
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:ShowNote"})			
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:HideNote"})		
			else:  # 複数セルを選択している時。
				if len(selection.getColumns())==1:  # 選択列が1つだけの時。
					if VARS.splittedrow<=r<=VARS.emptyrow:  # 左上セルの行の取引日列が入力済の時。
						txt = hojokamoku if hojokamoku else kamoku  # 補助科目行に値がある時は補助科目行、ないときは科目行の値を使う。
						addMenuentry("ActionTrigger", {"Text": "選択伝票を「{}」で決済".format(txt), "CommandURL": baseurl.format("entry8")}) 
						if txt!="現金":  # 現金列でない時のみ。
							addMenuentry("ActionTrigger", {"Text": "選択伝票を「現金」で決済", "CommandURL": baseurl.format("entry9")}) 		
						addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})			
				addMenuentry("ActionTrigger", {"Text": "正負反転", "CommandURL": baseurl.format("entry10")}) 								
		elif r>=VARS.splittedrow:  # 固定行より下。
			if c==VARS.daycolumn+1:  # 摘要列の時。
				if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 単独セルの時のみ。
					addMenuentry("ActionTrigger", {"Text": "伝票履歴", "CommandURL": baseurl.format("entry6")}) 
					addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
				if VARS.splittedrow<VARS.emptyrow:  # 伝票行がある時のみ。
					addMenuentry("ActionTrigger", {"Text": "伝票履歴に追加", "CommandURL": baseurl.format("entry7")}) 
			elif c==VARS.daycolumn:  # 取引日列の時。
				addMenuentry("ActionTrigger", {"Text": "文字列を日付に変換", "CommandURL": baseurl.format("entry11")}) 
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})					
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
			if r!=VARS.splittedrow:  # ModifyListenrが外れるので固定行の上に行の挿入はしない。
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsBefore"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsAfter"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteRows"}) 				
	elif contextmenuname=="colheader" and len(selection[:, 0].getRows())==len(sheet[:, 0].getRows()):  # 列ヘッダーの時、かつ、選択範囲の行数がシートの行数が一致している時。	
		if c>=VARS.splittedcolumn:
			commons.cutcopypasteMenuEntries(addMenuentry)
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})		
			if c!=VARS.splittedcolumn:  # ModifyListenrが外れるので固定列の左に行の挿入はしない。
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertColumnsBefore"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertColumnsAfter"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteColumns"}) 				
			if len(selection.getColumns())>1:  # 複数列を選択している時。
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})		
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:Group"})	
				addMenuentry("ActionTrigger", {"CommandURL": ".uno:Ungroup"})	
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:RenameTable"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。	
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。		
	controller = xscriptcontext.getDocument().getCurrentController()  # コントローラの取得。
	selection = controller.getSelection()  # 選択範囲を取得。
	sheet = VARS.sheet
	componentwindow = controller.ComponentWindow	
	querybox = lambda x: componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "WEntryBook", x)
	if entrynum==1:  # クリア。書式設定とオブジェクト以外を消去。
		selection.clearContents(511)  # 範囲をすべてクリアする。
	elif entrynum==2:  # 勘定元帳生成
		txt = selection.getString()
		basetxt = "{}の元帳生成".format(txt)
		msgbox = querybox("{}します。".format(basetxt))
		if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
			return			
		createMotoCho(xscriptcontext, basetxt, "勘定元帳_{}".format(txt), lambda x: [txt])			
	elif entrynum==3:  # 補助元帳生成
		c = selection.getCellAddress().Column
		kamokurows = sheet[VARS.kamokurow:VARS.kamokurow+2, c].getDataArray()
		txt = "{}_{}".format(kamokurows[0][0], kamokurows[1][0]) if kamokurows[1][0] else kamokurows[0][0]
		basetxt = "{}の補助元帳生成".format(txt)
		msgbox = querybox("{}します。".format(basetxt))
		if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
			return	
		createHojoMotoCho(xscriptcontext, basetxt, "補助元帳_{}".format(txt), lambda x: [c-VARS.splittedcolumn])			
	elif entrynum==4:  # 現金で決済
		datarow = sheet[VARS.kamokurow, :VARS.emptycolumn].getDataArray()[0]
		settle(sheet[selection.getCellAddress().Row, datarow.index("現金", VARS.splittedcolumn)])
	elif entrynum==5:  # 決済
		settle(selection)
	elif entrynum==6:  # 伝票履歴。単独セルの時のみ。
		datarow = sheet[selection.getCellAddress().Row, VARS.daycolumn+1:VARS.emptycolumn].getDataArray()[0]
		if any(filter(None, datarow)):
			msg = "すでに伝票データが存在する行です。\n上書きしますか？"
			componentwindow = controller.ComponentWindow
			msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "WEntryBook", msg)
			if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。			
				return
		historydialog.createDialog(xscriptcontext, "伝票履歴", callback=callback_sliphistoryCreator(xscriptcontext, selection))
	elif entrynum==7:  # 伝票履歴に追加。複数行選択の時もあり。
		newgriddatarows = []  # グリッドコントロールに追加する行のリスト。
		datarows = sheet[:VARS.emptyrow, :VARS.emptycolumn].getDataArray()
		headerrows = generateHeaderRows(datarows[:VARS.kamokurow+2])
		rangeaddress = selection.getRangeAddress()  # 選択範囲のアドレスを取得。
		tekiyocolumn = VARS.daycolumn + 1
		splittedcolumn = VARS.splittedcolumn
		for i in range(rangeaddress.StartRow, rangeaddress.EndRow+1):  # 行インデックスをイテレート。
			datarow = datarows[i]
			key = datarow[tekiyocolumn]  # 摘要を取得。
			if not key:
				commons.showErrorMessageBox(controller, "摘要がない行は履歴に追加できません。")	
				continue
			columnsgene = compress(zip(*headerrows, datarow[splittedcolumn:]), datarow[splittedcolumn:])  # 金額のある列のみ(列インデックス、区分、科目、補助科目、金額)をイテレートするジェネレーター。。
			kamokuvaldic = OrderedDict(("/".join(j[1:4]), (int(j[4]), sheet[i, j[0]].getAnnotation().getString().strip())) for j in columnsgene)  # キー: (区分,科目,補助科目)を結合した文字列、値: (金額、コメント)の辞書。jsonに変換するにはキーは文字列でないといけない。
			griddatarow = ["{}: {}".format(key, json.dumps(kamokuvaldic, ensure_ascii=False))]  # 辞書オブジェクトはJSONで文字列にする。タプルで渡してもjsonにするとリストで返ってくるのでリストにしておく。
			newgriddatarows.append(griddatarow)
		if newgriddatarows:
			doc = xscriptcontext.getDocument()
			dialogtitle = "伝票履歴"
			griddatarows = dialogcommons.getSavedData(doc, "GridDatarows_{}".format(dialogtitle))  # グリッドコントロールの行をconfigシートのragenameから取得する。	
			if griddatarows:  # 行のリストが取得出来た時。
				griddatarows = [i for i in griddatarows if i not in newgriddatarows]  # これから追加する行と同じ行はまず削除する。
				griddatarows.extend(newgriddatarows)
			else:
				griddatarows = newgriddatarows
			dialogcommons.saveData(doc, "GridDatarows_{}".format(dialogtitle), griddatarows)
	elif entrynum==8:  # 選択伝票を決済。選択列は1個のみ。選択範囲左上セルは固定行以下。
		rangeaddress = selection.getRangeAddress()
		settleMultipleSlips(rangeaddress, rangeaddress.StartColumn)
	elif entrynum==9:  # 選択伝票を現金で決済	。選択列は1個のみ。選択範囲左上セルは固定行以下。
		rangeaddress = selection.getRangeAddress()
		datarow = sheet[VARS.kamokurow, :VARS.emptycolumn].getDataArray()[0]  # 科目行を取得。
		settleMultipleSlips(rangeaddress, datarow.index("現金", VARS.splittedcolumn))
	elif entrynum==10:  # 正負反転。複数セル選択の時もあり。setDataArray()で処理するとコメントの処理が必要で時間がかかる。
		for i in selection.queryContentCells(CellFlags.VALUE).getCells():  # 数値が入っているセルを取得。
			val = i.getValue()
			if val:  # 値が0でない時。
				i.setValue(-val)	
	elif entrynum==11:  # 文字列を日付に変換。複数セル選択の時もあり。
		for i in selection[:, 0].queryContentCells(CellFlags.STRING).getCells():  # 文字列が入っているセルを取得。取引日列のみ対象にする。		
			s = i.getString()	
			if s[:4].isdigit():  # 先頭の4文字が数字の時は西暦と判断する。
				i.setFormula(s.replace(s[4], "-"))  # 区切り文字を-に変換して式として代入し直す。
		selection[:, 0].setPropertyValue("NumberFormat", commons.formatkeyCreator(xscriptcontext.getDocument())("YYYY-MM-DD"))
def settleMultipleSlips(rangeaddress, c):		
	sheet = VARS.sheet
	edgerow = rangeaddress.EndRow + 1
	newedgerow = VARS.emptyrow if edgerow>VARS.emptyrow else edgerow  # 最終行端を取得。
	startrow = rangeaddress.StartRow
	datarows = sheet[startrow:newedgerow, :VARS.emptycolumn].getDataArray()
	sliptotalcolumn = VARS.sliptotalcolumn
	for i, datarow in enumerate(datarows, start=startrow):
		if isinstance(datarow[sliptotalcolumn], float) and datarow[sliptotalcolumn]!=0:  # 伝票内計が数値かつ0でない時のみ。
			cellvalue = datarow[c] if isinstance(datarow[c], float) else 0  # 数値以外のときは0にする。
			val = (cellvalue-datarow[sliptotalcolumn]) or ""  # 0の時は空文字を代入。
			sheet[i, c].setDataArray(((val,),))  # 文字列でも数値でも代入できるのでsetDataArray()を使って代入。			
def settle(cell):
	celladdress = cell.getCellAddress()
	datarow = VARS.sheet[celladdress.Row, :VARS.emptycolumn].getDataArray()[0]
	cellvalue = datarow[celladdress.Column]
	cellvalue = cellvalue if isinstance(cellvalue, float) else 0  # 数値以外のときは0にする。
	val = (cellvalue-datarow[VARS.sliptotalcolumn]) or ""  # 0の時は空文字を代入。
	cell.setDataArray(((val,),))  # 文字列でも数値でも代入できるのでsetDataArray()を使って代入。			
def callback_sliphistoryCreator(xscriptcontext, selection):		
	def callback_sliphistory(gridcelltxt):
		tekiyo, jsondata = gridcelltxt.split(":", 1)  # 摘要、と、科目金額辞書の文字列を取得する。
		try:
			kamokuvaldic = json.loads(jsondata)  # 科目金額辞書の文字列を辞書オブジェクトに復元する。
		except json.JSONDecodeError as e:  # json構文にエラーがある時。
			dialogcommons.showJSONError(xscriptcontext.getDocument(), e)
			return
		kamokuvaldic = {tuple(k.split("/")): v for k, v in kamokuvaldic.items()}  # 科目金額辞書のキーをタプルに変換して再取得。	
		sheet = VARS.sheet
		datarows = sheet[:VARS.kamokurow+2, :VARS.emptycolumn].getDataArray()
		headerrows = generateHeaderRows(datarows)
		newdatarow = [tekiyo]	
		comments = []  # コメントのセルとコメントのタプルを取得するリスト。
		r = selection.getCellAddress().Row
		for i in zip(*headerrows):  # (列インデックス、区分、科目、補助科目)をイテレートする。	
			if i[1:] in kamokuvaldic:
				val, annotation = kamokuvaldic.pop(i[1:])
				if annotation:
					comments.append((sheet[r, i[0]], annotation))  # setDataArray()でコメントがクリアされるのでここでセルとコメントの文字列をタプルで取得しておく。
			else:
				val = ""		
			newdatarow.append(val)
		sheet[r, VARS.daycolumn+1:VARS.emptycolumn].setDataArray((newdatarow,))
		annotations = sheet.getAnnotations()  # コメントコレクションを取得。
		for i in comments:
			cell, annotation = i
			annotations.insertNew(cell.getCellAddress(), annotation)  # コメントを挿入。
			cell.getAnnotation().getAnnotationShape().setPropertyValue("Visible", False)  # これをしないとmousePressed()のTargetにAnnotationShapeが入ってしまう。				
		if len(kamokuvaldic)>0:  # 辞書に要素が残っている時。
			msg = "今年度には存在しない科目がありました。\n\n{}".format("\n".join("{}: {}".format(*i) for i in kamokuvaldic.items()))
			componentwindow = xscriptcontext.getDocument().getCurrentController().ComponentWindow
			componentwindow.getToolkit().createMessageBox(componentwindow, WARNINGBOX, MessageBoxButtons.BUTTONS_OK, "WEntryBook", msg).execute()					
	return callback_sliphistory	
def getDateSection():  # 期首日と期末日のdateオブジェクトのタプルを返す。
	dates = []
	for i in VARS.settlingdayrows:
		datecell = VARS.sheet[i, VARS.daycolumn]
		datevalue = datecell.getValue()  # 決算日セルから値を取得。
		if datevalue>0:  # 値が正の数の時はセルには日付が入っている。
			datetxt = datecell.getString()  # 日付を文字列で取得。
			dates.append(date(*tuple(map(int, datetxt.split(datetxt[4])))))
	if len(dates)==2:
		return dates
	return (None,)*2
def insertPageBreaksCreator(pagestyle):  # 行インデックスが偶数のところのみに改ページを挿入する。
	properties = "Height", "TopMargin", "BottomMargin", "HeaderIsOn", "HeaderHeight", "FooterIsOn", "FooterHeight"
	height, topmargin, bottommargin, headerison, headerheight, footerison, footerheight = pagestyle.getPropertyValues(properties)
	pageheight = height - topmargin - bottommargin  # 印刷高さを1/100mmで取得。
	if headerison:  # ヘッダーがあるときヘッダーの高さを除く。
		pageheight -= headerheight
	if footerison:  # フッターがあるときフッターの高さを除く。
		pageheight -= footerheight
	def insertPageBreaks(sheet, rows):
		rows.setPropertyValue("IsStartOfNewPage", False)  # すでにある改ページを消去。
		h = 0  # 行の高さの合計。
		for i in range(0, len(rows), 2):  # 偶数の行インデックスをイテレート。
			rowheight = rows[i].getPropertyValue("Height") + rows[i+1].getPropertyValue("Height")  # 行の高さを取得。
			h += rowheight  # 行の高さを加算する。
			if h>pageheight:  # 1ページあたりの高さを越えた時。
				rows[i].setPropertyValue("IsStartOfNewPage", True)  # 改ページを挿入。
				h = rowheight   # 行の高さをリセット。
	return insertPageBreaks
