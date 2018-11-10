#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper # import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from datetime import date, timedelta
from . import dialogcommons, journal
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
	cellvalue = enhancedmouseevent.Target.getValue()  # セルの値を取得。
	centerday = None
	if cellvalue>0:  # セルの値が0より大きい時、日付シリアル値と断定する。文字列のときは0.0が返る。
		functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
		if cellvalue!=functionaccess.callFunction("TODAY", ()):  # セルの数値が今日でない時。
			centerday = date(*[int(functionaccess.callFunction(i, (cellvalue,))) for i in ("YEAR", "MONTH", "DAY")])  # シリアル値をシート関数で年、月、日に変換してdateオブジェクトにする。
			col0[todayindex] = "セル値"
	if centerday is None:
		centerday = date.today()
		col0[todayindex-1:todayindex+2] = "昨日", "今日", "明日"  # 列インデックス0に入れる文字列を取得。
		
		
	settlingdatedigits = journal.VARS.settlingdatedigits
	if settlingdatedigits:  # シートの年度が取得できた時。
		y, m, d = settlingdatedigits
		sdate = date(y-1, m, d) + timedelta(days=1)  # 年度開始日。
		edate = date(*settlingdatedigits)  # 年度終了日。			
		numericfield1.setMin((sdate-centerday).days//7)  # 最小週数を設定。
		numericfield1.setMax((edate-centerday).days//7)  # 最大週数を設定。
		
		
		
	addDays(gridcontrol1, centerday, col0)  # グリッドコントロールに行を入れる。	
	if cellvalue>0: # セルに値があった時。
		gridcontrol1.selectRow(todayindex)  # セル値の行を選択する。
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
def addDays(gridcontrol, centerday, col0, daycount=7):
	todayindex = 7//2  # 今日の日付の位置を決定。切り下げ。
	startday = centerday - timedelta(days=1)*todayindex  # 開始dateを取得。
	
	
	
	settlingdatedigits = journal.VARS.settlingdatedigits
	if settlingdatedigits:  # シートの年度が取得できた時。
		y, m, d = settlingdatedigits
		sdate = date(y-1, m, d) + timedelta(days=1)  # 年度開始日。
		edate = date(*settlingdatedigits)  # 年度終了日。

			
			
	dategene = (startday+timedelta(days=i) for i in range(daycount))  # daycount分のdateオブジェクトのジェネレーターを取得。
	weekdays = "月", "火", "水", "木", "金", "土", "日"
	datarows = tuple(zip(col0, ("{}-{}-{}({})".format(i.year, i.month, i.day, weekdays[i.weekday()]) if sdate<=i<=edate else "" for i in dategene)))  # 列インデックス0に語句、列インデックス1に日付を入れる。
	griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModel
	griddatamodel.removeAllRows()  # グリッドコントロールの行を全削除。
	griddatamodel.addRows(("",)*len(datarows), datarows)  # グリッドに行を追加。	
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
	def __init__(self, *args):
		self.args = args
		self.val = 0  # 変更前の値。
	def textChanged(self, textevent):
		numericfield = textevent.Source
		gridcontrol, = self.args
		
		
		todayindex = 7//2  # 本日と同じインデックスを取得。
		
		
		datetxt = gridcontrol.getModel().getPropertyValue("GridDataModel").getCellData(1, todayindex)  # 中央行の日付文字列を取得。
		
		if datetxt:
			centerday = date(*map(int, datetxt.split("(")[0].split("-")))
		else:
			
			
			
		
		
		val = numericfield.getValue()  # 数値フィールドの値を取得。		
		diff = val - self.val  # 前値との差を取得。
		centerday += timedelta(days=7*diff)  # 週を移動。
		col0 = [""]*7
		if val==0:
			if centerday==date.today():
				col0[todayindex-1:todayindex+2] = "昨日", "今日", "明日"  # 列インデックス0に入れる文字列を取得。
			else:	
				col0[todayindex] = "基準日"
		else:
			txt = "{}週後" if val>0 else "{}週前" 
			col0[todayindex] = txt.format(int(abs(val)))  # valはfloatなので小数点が入ってくる。		
		addDays(gridcontrol, centerday, col0)  # グリッドコントロールに行を入れる。
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
					doc = xscriptcontext.getDocument()
					selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
					if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
						rowindexes = dialogcommons.getSelectedRowIndexes(gridcontrol)  # グリッドコントロールの選択行インデックスを返す。昇順で返す。負数のインデックスがある時は要素をクリアする。
						if rowindexes:
							datetxt = gridcontrol.getModel().getPropertyValue("GridDataModel").getCellData(1, rowindexes[0])  # 選択行の日付文字列を取得。
							if outputcolumn is not None:  # 出力する列が指定されている時。
								sheet = selection.getSpreadsheet()
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
								callback(datetxt)
					for menuid in range(1, self.gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
						itemtext = self.gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
						if itemtext.startswith("セル入力で閉じる"):
							if self.gridpopupmenu.isItemChecked(menuid):  # 選択項目にチェックが入っている時。
								self.dialogframe.close(True)
							else:
								controller = doc.getCurrentController()  # 現在のコントローラを取得。	
								sheet = controller.getActiveSheet()
								celladdress = selection.getCellAddress()
								nextcell = sheet[celladdress.Row+1, celladdress.Column]  # 下のセルを取得。
								controller.select(nextcell)  # 下のセルを選択。							
							break
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
		