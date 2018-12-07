#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper # import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from datetime import date, timedelta
from . import dialogcommons, exceptiondialog2, journal
from com.sun.star.awt import XMenuListener, XMouseListener, XTextListener
from com.sun.star.awt import MenuItemStyle, MouseButton, PopupMenuDirection, PosSize  # 定数
from com.sun.star.awt import Rectangle  # Struct
from com.sun.star.beans import NamedValue  # Struct
from com.sun.star.lang import Locale  # Struct
from com.sun.star.style.VerticalAlignment import MIDDLE  # enum
from com.sun.star.util import XCloseListener
from com.sun.star.view.SelectionType import SINGLE  # enum 
YEAR = None
def createDialog(enhancedmouseevent, xscriptcontext, dialogtitle, formatstring=None, outputcolumn=None, *, callback=None):  # dialogtitleはダイアログのデータ保存名に使うのでユニークでないといけない。formatstringは代入セルの書式。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	doc = xscriptcontext.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   
	dialogpoint = dialogcommons.getDialogPoint(doc, enhancedmouseevent)  # クリックした位置のメニューバーの高さ分下の位置を取得。単位ピクセル。一部しか表示されていないセルのときはNoneが返る。
	if not dialogpoint:  # クリックした位置が取得出来なかった時は何もしない。
		return
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	containerwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	maTopx = dialogcommons.createConverters(containerwindow)  # ma単位をピクセルに変換する関数を取得。
	m = 2  # コントロール間の間隔。
	h = 12  # コントロールの高さ
	gridprops = {"PositionX": 0, "PositionY": 0, "Width": 76, "Height": 73, "ShowRowHeader": False, "ShowColumnHeader": False, "VScroll": False, "SelectionModel": SINGLE}  # グリッドコントロールのプロパティ。
	controlcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": dialogcommons.XWidth(gridprops), "Height": dialogcommons.YHeight(gridprops), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
	optioncontrolcontainerprops = controlcontainerprops.copy()
	controlcontainer, addControl = dialogcommons.controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。		
	items = ("セル入力で閉じる", MenuItemStyle.CHECKABLE+MenuItemStyle.AUTOCHECK, {"checkItem": True}),  # グリッドコントロールのコンテクストメニュー。XMenuListenerのmenuevent.MenuIdでコードを実行する。
	menulistener = MenuListener()  # コンテクストメニューにつけるリスナー。
	gridpopupmenu = dialogcommons.menuCreator(ctx, smgr)("PopupMenu", items, {"addMenuListener": menulistener, "hideDisabledEntries": False})  # 右クリックでまず呼び出すポップアップメニュー。 
	mousemotionlistener = dialogcommons.MouseMotionListener()
	args = xscriptcontext, formatstring, outputcolumn, callback
	mouselistener = MouseListener(gridpopupmenu, args)
	gridcontrolwidth = gridprops["Width"]  # gridpropsは消費されるので、グリッドコントロールの幅を取得しておく。
	gridcontrol1 = addControl("Grid", gridprops, {"addMouseListener": mouselistener, "addMouseMotionListener": mousemotionlistener})  # グリッドコントロールの取得。
	gridcolumn = gridcontrol1.getModel().getPropertyValue("ColumnModel")  # DefaultGridColumnModel
	column0 = gridcolumn.createColumn() # 列の作成。
	column0.ColumnWidth = 25 # 列幅。
	gridcolumn.addColumn(column0)  # 1列目を追加。
	column1 = gridcolumn.createColumn() # 列の作成。
	column1.ColumnWidth = gridcontrolwidth - column0.ColumnWidth #  列幅。列の合計がグリッドコントロールの幅に一致するようにする。
	gridcolumn.addColumn(column1)  # 2列目を追加。
	numericfieldprops1 = {"PositionY": m, "Width": 24, "Height": h+2, "Spin": True, "StrictFormat": True, "Value": 0, "ValueStep": -1, "ShowThousandsSeparator": False, "DecimalAccuracy": 0}
	fixedtextprops1 = {"PositionY": m, "Width": 14, "Height": h, "Label": "週後", "VerticalAlign": MIDDLE}
	fixedtextprops1.update({"PositionX": gridcontrolwidth-fixedtextprops1["Width"]})
	numericfieldprops1.update({"PositionX": fixedtextprops1["PositionX"]-numericfieldprops1["Width"]})
	optioncontrolcontainerprops.update({"PositionY": dialogcommons.YHeight(optioncontrolcontainerprops), "Height": dialogcommons.YHeight(numericfieldprops1, m)})
	optioncontrolcontainer, optionaddControl = dialogcommons.controlcontainerMaCreator(ctx, smgr, maTopx, optioncontrolcontainerprops)  # コントロールコンテナの作成。
	textlistener = TextListener(gridcontrol1)
	numericfield1 = optionaddControl("NumericField", numericfieldprops1, {"addTextListener": textlistener})		
	optionaddControl("FixedText", fixedtextprops1)
	rectangle = controlcontainer.getPosSize()  # コントロールコンテナのRectangle Structを取得。px単位。
	rectangle.X, rectangle.Y = dialogpoint  # クリックした位置を取得。ウィンドウタイトルを含めない座標。
	rectangle.Height += optioncontrolcontainer.getSize().Height
	taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
	args = NamedValue("PosSize", rectangle), NamedValue("FrameName", "controldialog")  # , NamedValue("MakeVisible", True)  # TaskCreatorで作成するフレームのコンテナウィンドウのプロパティ。
	dialogframe = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。サイズ変更は想定しない。
	mouselistener.dialogframe = dialogframe
	dialogwindow = dialogframe.getContainerWindow()  # ダイアログのコンテナウィンドウを取得。
	dialogframe.setTitle(dialogtitle)  # フレームのタイトルを設定。
	docframe.getFrames().append(dialogframe) # 新しく作ったフレームを既存のフレームの階層に追加する。
	toolkit = dialogwindow.getToolkit()  # ピアからツールキットを取得。 	
	controlcontainer.createPeer(toolkit, dialogwindow) # ウィンドウにコントロールコンテナを描画。
	optioncontrolcontainer.createPeer(toolkit, dialogwindow) # ウィンドウにオプションコントロールコンテナを描画。Visibleにはしない。
	frameactionlistener = dialogcommons.FrameActionListener()  # FrameActionListener。フレームがアクティブでなくなった時に閉じるため。
	dialogframe.addFrameActionListener(frameactionlistener)  # FrameActionListenerをダイアログフレームに追加。
	controlcontainer.setVisible(True)  # コントロールの表示。
	optioncontrolcontainer.setVisible(True)
	dialogwindow.setVisible(True) # ウィンドウの表示。これ以降WindowListenerが発火する。
	numericfield1.setFocus()
	todayindex = 7//2  # 今日の日付の位置を決定。切り下げ。
	col0 = [""]*7  # 全てに空文字を挿入。
	selection = enhancedmouseevent.Target  # 選択セルを取得。
	datevalue = selection.getValue()  # セルの値を取得。
	if not datevalue>0:  # セルに日付が入っていない時。
		centerdate = date.today()  # 今日の日付を中央にする。
		col0[todayindex-1:todayindex+2] = "昨日", "今日", "明日"  # 列インデックス0に入れる文字列を取得。
	else:  # セルに日付が入っている時。
		datetxt = selection.getString()  # 日付文字列を取得。2018-8-5などを想定。
		centerdate = date(*map(int, datetxt.split(datetxt[4])))  # 日付文字列をdateオブジェクトにして中央にする。
		col0[todayindex] = "セル値"
	sdate, edate = None, None
	lowerlimit, col0min, upperlimit, col0max = None, None, None, None	
	if selection.getCellAddress().Row>=journal.VARS.splittedrow:  # 選択セルが固定行以下の時のみ上限と下限を指定する。	
		sdate, edate = journal.getDateSection()  # 期首日と期末日のdateオブジェクトを取得。
		if sdate:  # 期首日と期末日が取得出来ている問。
			if not sdate<=centerdate<=edate:  # 中央値が会計期間内でない時。centerdateを変更する。
				if centerdate<sdate:  # 期首日より新しい取得日の時。
					centerdate = sdate + timedelta(days=todayindex)  # 期首日が１番上に来るようにする。
					lowerlimit = 0  # 最小週数を設定。
					col0 = ("期首日", *[""]*6)		
				else:  # 期末日より古い取得日の時。
					centerdate = edate - timedelta(days=todayindex)  # 期末日が一番下に来るようにする。
					upperlimit = 0  # 最大週数を設定。
					numericfield1.setMax(0)
					col0 = (*[""]*6, "期末日")  		
			if lowerlimit is None:
				diffmindays = (centerdate-timedelta(days=todayindex)-sdate).days  # centerdateから期首日までの日数差。
				lowerlimit = diffmindays//-7  # 期首日までの週数差。負数が返る。
				indexmin = (7-diffmindays%7)%7  # 最小週数での期首日の位置。
				col0min = (*[""]*indexmin, "期首日", *[""]*(6-indexmin))		
			if upperlimit is None:	
				diffmaxdays = (edate-timedelta(days=todayindex)-centerdate).days  # centerdateから期末日までの日数差。
				upperlimit = -(diffmaxdays//-7)   # 期末日までの週数差。
				indexmax = (diffmaxdays%7-1)%7  # 最大週数での期末日の位置。
				col0max = (*[""]*indexmax, "期末日", *[""]*(6-indexmax))		
			numericfield1.setMin(lowerlimit)  # 最小週数を設定。		
			numericfield1.setMax(upperlimit)  # 最大週数を設定。				
	textlistener.colargs = col0, lowerlimit, col0min, upperlimit, col0max		
	addDays = addDaysCreator(selection, gridcontrol1, sdate, edate)
	textlistener.addDays = addDays
	addDays(centerdate, col0)  # グリッドコントロールに行を入れる。	
	menulistener.args = controlcontainer, mouselistener, mousemotionlistener
	dialogstate = dialogcommons.getSavedData(doc, "dialogstate_{}".format(dialogtitle))  # 保存データを取得。optioncontrolcontainerの表示状態は常にFalseなので保存されていない。
	if dialogstate is not None:  # 保存してあるダイアログの状態がある時。
		for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
			itemtext = gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
			if itemtext.startswith("セル入力で閉じる"):
				closecheck = dialogstate.get("CloseCheck")  # セル入力で閉じる、のチェックがある時。
				if closecheck is not None:
					gridpopupmenu.checkItem(menuid, closecheck)	
	args = doc, mouselistener, controlcontainer, mousemotionlistener, menulistener
	dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。	
def addDaysCreator(selection, gridcontrol, sdate, edate):
	todayindex = 7//2  # 今日の日付の位置を決定。切り下げ。
	weekdays = "月", "火", "水", "木", "金", "土", "日"		
	if selection.getCellAddress().Row<journal.VARS.splittedrow or not all([sdate, edate]):  # 選択セルが固定行より上、または期首と期末が指定されていない時。
		datetxtgene = lambda x: ("{}({})".format(i.isoformat(), weekdays[i.weekday()]) for i in x) 	
	elif all([sdate, edate]):  # 期首日と期末日ともに指定されている時。
		datetxtgene = lambda x: ("{}({})".format(i.isoformat(), weekdays[i.weekday()]) if sdate<=i<=edate else "" for i in x)  # 年度開始日と終了日以外には空文字を入れる。	
	elif sdate:  # 期首日のみ指定の時。
		datetxtgene = lambda x: ("{}({})".format(i.isoformat(), weekdays[i.weekday()]) if sdate<=i else "" for i in x)  	
	elif edate:  # 期末日のみ指定の時。
		datetxtgene = lambda x: ("{}({})".format(i.isoformat(), weekdays[i.weekday()]) if i<=edate else "" for i in x)
	def addDays(centerdate, col0):
		startdate = centerdate - timedelta(days=1)*todayindex  # 開始dateを取得。
		dategene = (startdate+timedelta(days=i) for i in range(7))  # dateオブジェクトのジェネレーターを取得。
		datarows = tuple(zip(col0, datetxtgene(dategene)))  # 列インデックス0に語句、列インデックス1に日付を入れる。
		griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModel
		griddatamodel.removeAllRows()  # グリッドコントロールの行を全削除。
		griddatamodel.addRows(("",)*len(datarows), datarows)  # グリッドに行を追加。	
	return addDays
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログのリスナー削除用。
	def __init__(self, args):
		self.args = args
	def queryClosing(self, eventobject, getsownership):  # ノンモダルダイアログを閉じる時に発火。
		dialogframe = eventobject.Source
		doc, mouselistener, controlcontainer, mousemotionlistener, menulistener = self.args
		gridpopupmenu = mouselistener.gridpopupmenu
		for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
			itemtext = gridpopupmenu.getItemText(menuid)
			if itemtext.startswith("セル入力で閉じる"):
				dialogstate = {"CloseCheck": gridpopupmenu.isItemChecked(menuid)}
		dialogtitle = dialogframe.getTitle()  # コンテナウィンドウタイトルを取得。データ保存のIDに使う。
		dialogcommons.saveData(doc, "dialogstate_{}".format(dialogtitle), dialogstate)  # ダイアログの状態を保存。
		gridpopupmenu.removeMenuListener(menulistener)
		gridcontrol1 = controlcontainer.getControl("Grid1")
		gridcontrol1.removeMouseListener(mouselistener)
		gridcontrol1.removeMouseMotionListener(mousemotionlistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		pass
class TextListener(unohelper.Base, XTextListener):
	def __init__(self, gridcontrol):
		self.gridcontrol = gridcontrol
		self.colargs = None  # 一列目に関する引数のタプル。
		self.val = 0  # 変更前の値。
		self.addDays = None  # グリッドコントロールに日付を追加する関数。
	def textChanged(self, textevent):
		numericfield = textevent.Source
		todayindex = 7//2  # 本日と同じインデックスを取得。
		griddatamodel = self.gridcontrol.getModel().getPropertyValue("GridDataModel")
		datetxt = griddatamodel.getCellData(1, 0)  # 先頭行の日付文字列を取得。
		if datetxt:  # 先頭行に日付文字列がある時。
			centerdate = date(*map(int, datetxt.split("(")[0].split("-"))) + timedelta(days=todayindex)
		else:  # 先頭行に日付文字列がない時は最終行から日付文字列を取得する。
			datetxt = griddatamodel.getCellData(1, griddatamodel.RowCount-1)  # 最終行の日付文字列を取得。
			centerdate = date(*map(int, datetxt.split("(")[0].split("-"))) - timedelta(days=todayindex)
		val = numericfield.getValue()  # 数値フィールドの値を取得。		
		diff = val - self.val  # 前値との差を取得。
		centerdate += timedelta(days=7*diff)  # 週を移動。
		col0 = [""]*7
		col0init, lowerlimit, col0min, upperlimit, col0max = self.colargs	
		if val==0:  # 開始日の時。
			if col0init is not None:
				col0 = col0init
		elif lowerlimit is not None and val==lowerlimit:
			col0 = col0min	
		elif upperlimit is not None and val==upperlimit:
			col0 = col0max	
		else:
			txt = "{}週後" if val>0 else "{}週前" 
			col0[todayindex] = txt.format(int(abs(val)))  # valはfloatなので小数点が入ってくる。	
		self.addDays(centerdate, col0)  # グリッドコントロールに行を入れる。
		self.val = val  # 変更後の値を前値として取得。
	def disposing(self, eventobject):
		pass
class MouseListener(unohelper.Base, XMouseListener):  
	def __init__(self, gridpopupmenu, args): 	
		self.args = args
		self.gridpopupmenu = gridpopupmenu  # CloseListenerでも使用する。
		self.dialogframe = None
		self.flg = True  # リスナーをつけると発火してしまうのを抑制するフラグ。
	def mousePressed(self, mouseevent):  # グリッドコントロールをクリックした時。コントロールモデルにはNameプロパティはない。
		xscriptcontext, formatstring, outputcolumn, callback = self.args
		gridcontrol = mouseevent.Source  # グリッドコントロールを取得。
		if mouseevent.Buttons==MouseButton.LEFT:
			if mouseevent.ClickCount==1:  # シングルクリックでセルに入力する。
				if self.flg:
					closeflg = False  # ダイアログを閉じた時に立てるフラグ。		
					doc = xscriptcontext.getDocument()
					selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
					if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
						sheet = selection.getSpreadsheet()
						rowindexes = dialogcommons.getSelectedRowIndexes(gridcontrol)  # グリッドコントロールの選択行インデックスを返す。昇順で返す。負数のインデックスがある時は要素をクリアする。
						if rowindexes:
							for menuid in range(1, self.gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
								itemtext = self.gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
								if itemtext.startswith("セル入力で閉じる"):
									if self.gridpopupmenu.isItemChecked(menuid):  # 選択項目にチェックが入っている時。
										self.dialogframe.close(True)  # 用が終わったらさっさと閉じないとその前にブレークや例外がでるとマウスが使えなくなる。
										closeflg = True							
							datetxt = gridcontrol.getModel().getPropertyValue("GridDataModel").getCellData(1, rowindexes[0])  # 選択行の日付文字列を取得。
							if outputcolumn is not None:  # 出力する列が指定されている時。
								selection = sheet[selection.getCellAddress().Row, outputcolumn]  # 同じ行の指定された列のセルを取得。						
							if formatstring is not None:  # 書式が与えられている時。
								numberformats = doc.getNumberFormats()  # ドキュメントのフォーマット一覧を取得。デフォルトのフォーマット一覧はCalcの書式→セル→数値でみれる。
								localestruct = Locale(Language="ja", Country="JP")  # フォーマット一覧をくくる言語と国を設定。インストールしていないUIの言語でもよい。
								formatkey = numberformats.queryKey(formatstring, localestruct, True)  # formatstringが既存のフォーマット一覧にあるか調べて取得。第3引数のブーリアンは意味はないはず。 
								if formatkey == -1:  # デフォルトのフォーマットにformatstringがないとき。
									formatkey = numberformats.addNew(formatstring, localestruct)  # フォーマット一覧に追加する。保存はドキュメントごと。
								selection.setPropertyValue("NumberFormat", formatkey)  # セルの書式を設定。 
							datetxt = datetxt.split("(")[0]  # 2018-8-7という書式にする。
							selection.setFormula(datetxt)  # 2018-8-7の書式で式としてセルに代入。
							if callback is not None:  # コールバック関数が与えられている時。
								try:
									callback(datetxt)	
								except:  # これをしないとエラーダイアログが出てこない。
									exceptiondialog2.createDialog(xscriptcontext)  # XSCRIPTCONTEXTを渡す。	
							if not closeflg:  # ダイアログが閉じられていない時。
								celladdress = selection.getCellAddress()
								nextcell = sheet[celladdress.Row+1, celladdress.Column]  # 下のセルを取得。
								doc.getCurrentController().select(nextcell)  # 下のセルを選択。							
				else:
					self.flg = True		
		elif mouseevent.Buttons==MouseButton.RIGHT:  # 右ボタンクリックの時。mouseevent.PopupTriggerではサブジェクトによってはTrueにならないので使わない。
			pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
			self.gridpopupmenu.execute(gridcontrol.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向							
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		pass
	def mouseExited(self, mouseevent):  
		pass
	def disposing(self, eventobject):
		pass
class MenuListener(unohelper.Base, XMenuListener):
	def __init__(self):
		self.args = None
	def itemHighlighted(self, menuevent):
		pass
	def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。どこのコントロールのメニューかを知る方法はない。
		controlcontainer, mouselistener, mousemotionlistener = self.args
		gridcontrol1 = controlcontainer.getControl("Grid1")		
		gridcontrol1.addMouseMotionListener(mousemotionlistener)	
		mouselistener.flg = False  # mouselistenerをつけたときに発火しても動作せないフラグ。
		gridcontrol1.addMouseListener(mouselistener)	
	def itemActivated(self, menuevent):  # メニュー項目を有効にした時。チェックボックスをオンにした時ではない。ポップアップメニューをexecute()した時も発火する。
		controlcontainer, mouselistener, mousemotionlistener = self.args
		gridcontrol1 = controlcontainer.getControl("Grid1")
		gridcontrol1.removeMouseListener(mouselistener)  # ポップアップメニュー上でもMouseListenerが発火するの外しておく。MouseListnerをつけたままダイアログを閉じるとLibreOfficeがクラッシュする。
		gridcontrol1.removeMouseMotionListener(mousemotionlistener)
	def itemDeactivated(self, menuevent):  # メニュー項目が無効になった時。ポップアップメニュー項目を選択せずに閉じた時も発火する。
		self.itemSelected(menuevent)
	def disposing(self, eventobject):
		pass
