#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from . import dialogcommons, exceptiondialog2
from com.sun.star.accessibility import AccessibleRole  # 定数
from com.sun.star.awt import XActionListener, XMenuListener, XMouseListener, XWindowListener, XTextListener, XItemListener
from com.sun.star.awt import MenuItemStyle, MessageBoxButtons, MessageBoxResults, MouseButton, PopupMenuDirection, PosSize, ScrollBarOrientation  # 定数
from com.sun.star.awt import MenuEvent, Rectangle, Selection  # Struct
from com.sun.star.awt.MessageBoxType import QUERYBOX  # enum
from com.sun.star.beans import NamedValue  # Struct
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.util import XCloseListener
from com.sun.star.view.SelectionType import MULTI  # enum 
from com.sun.star.lang import Locale  # Struct
DATAROWS = []  # グリッドコントロールのデータ行、タプルのタプルやリストのタプルやリストのリスト、の可能性がある。複数クラスからアクセスするのでグローバルにしないといけない。
def createDialog(xscriptcontext, dialogtitle, defaultrows=None, outputcolumn=None, *, enhancedmouseevent=None, callback=None):  # dialogtitleはダイアログのデータ保存名に使うのでユニークでないといけない。defaultrowsはグリッドコントロールのデフォルトデータ。
	items = ("選択行を削除", 0, {"setCommand": "delete"}),\
			("全行を削除", 0, {"setCommand": "deleteall"}),\
			(),\
			("セル入力で閉じる", MenuItemStyle.CHECKABLE+MenuItemStyle.AUTOCHECK, {"checkItem": True}),\
			("オプション表示", MenuItemStyle.CHECKABLE+MenuItemStyle.AUTOCHECK, {"checkItem": False})  # グリッドコントロールにつける右クリックメニュー。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	doc = xscriptcontext.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。  
	controller = doc.getCurrentController()
	docframe = controller.getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	containerwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	maTopx = dialogcommons.createConverters(containerwindow)  # ma単位をピクセルに変換する関数を取得。	
	m = 2  # コントロール間の間隔。
	h = 12  # コントロール間の高さ。	
	txt = doc.getCurrentSelection().getString()  # 選択セルの文字列を取得。
	XWidth, YHeight = dialogcommons.XWidth, dialogcommons.YHeight
	gridprops = {"PositionX": 0, "PositionY": 0, "Width": 50, "Height": 50, "ShowRowHeader": False, "ShowColumnHeader": False, "SelectionModel": MULTI}  # グリッドコントロールのプロパティ。
	controlcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(gridprops), "Height": YHeight(gridprops), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。		
	controlcontainer, addControl = dialogcommons.controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。		
	mousemotionlistener = dialogcommons.MouseMotionListener()  # グリッドコントロールにつけるマウスが動くと発火するリスナー。
	menulistener = MenuListener(mousemotionlistener)  # コンテクストメニューにつけるリスナー。mousemotionlistenerはグリッドコントロールにつけるもの。
	gridpopupmenu = dialogcommons.menuCreator(ctx, smgr)("PopupMenu", items, {"addMenuListener": menulistener, "hideDisabledEntries": False})  # 右クリックでまず呼び出すポップアップメニュー。hideDisabledEntries()が反応しない。 
	args = gridpopupmenu, xscriptcontext, outputcolumn, callback  # gridpopupmenuは先頭でないといけない。
	mouselistener = MouseListener(args)
	gridcontrol1 = addControl("Grid", gridprops, {"addMouseListener": mouselistener, "addMouseMotionListener": mousemotionlistener})  # グリッドコントロールの取得。
	gridmodel = gridcontrol1.getModel()  # グリッドコントロールモデルの取得。
	gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
	gridcolumn.addColumn(gridcolumn.createColumn())  # 列を追加。
	griddatamodel = gridmodel.getPropertyValue("GridDataModel")  # GridDataModel
	datarows = dialogcommons.getSavedData(doc, "GridDatarows_{}".format(dialogtitle))  # グリッドコントロールの行をconfigシートのragenameから取得する。	
	if datarows is None and defaultrows is not None:  # 履歴がなくデフォルトdatarowsがあるときデフォルトデータを使用。
		datarows = [i if isinstance(i, (list, tuple)) else (i,) for i in defaultrows]  # defaultrowsの要素をリストかタプルでなければタプルに変換する。
	if datarows:  # 行のリストが取得出来た時。
		griddatamodel.addRows(("",)*len(datarows), datarows)  # グリッドに行を追加。	
	else:
		datarows = []  # Noneのままではあとで処理できないので空リストを入れる。
	global DATAROWS
	DATAROWS = datarows  # マクロの起動中グローバル変数は保持されるので毎回リセットしないといけない。
	controlcontainerwindowlistener = ControlContainerWindowListener(controlcontainer)		
	controlcontainer.addWindowListener(controlcontainerwindowlistener)  # コントロールコンテナの大きさを変更するとグリッドコントロールの大きさも変更するようにする。
	textboxprops = {"PositionX": 0, "PositionY": m, "Height": h, "Text": txt}  # テクストボックスコントロールのプロパティ。
	checkboxprops = {"PositionY": YHeight(textboxprops, m), "Height": h, "Tabstop": False}  # チェックボックスコントロールのプロパティ。
	checkboxprops1, checkboxprops2 = [checkboxprops.copy() for dummy in range(2)]
	checkboxprops1.update({"PositionX": 0, "Width": 42, "Label": "~サイズ復元", "State": 1})  # サイズ復元はデフォルトでは有効。
	checkboxprops2.update({"PositionX": XWidth(checkboxprops1), "Width": 38, "Label": "~逐次検索", "State": 0})  # 逐次検索はデフォルトでは無効。
	buttonprops = {"PositionX": XWidth(checkboxprops2), "PositionY": YHeight(textboxprops, 4), "Width": 30, "Height": h+2, "Label": "New"}  # ボタンのプロパティ。PushButtonTypeの値はEnumではエラーになる。VerticalAlignではtextboxと高さが揃わない。
	textboxprops.update({"Width": XWidth(buttonprops)})
	optioncontrolcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(textboxprops), "Height": YHeight(buttonprops, 2), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
	optioncontrolcontainer, optionaddControl = dialogcommons.controlcontainerMaCreator(ctx, smgr, maTopx, optioncontrolcontainerprops)  # コントロールコンテナの作成。		
	textlistener = TextListener(xscriptcontext, controlcontainer)	
	optionaddControl("Edit", textboxprops, {"addTextListener": textlistener})  
	checkboxcontrol1 = optionaddControl("CheckBox", checkboxprops1)  
	itemlistener = ItemListener(controlcontainer)
	checkboxcontrol2 = optionaddControl("CheckBox", checkboxprops2, {"addItemListener": itemlistener}) 
	args = xscriptcontext, controlcontainer
	actionlistener = ActionListener(args)  # ボタンコントロールにつけるリスナー。	
	optionaddControl("Button", buttonprops, {"addActionListener": actionlistener, "setActionCommand": "enter"})  
	optioncontrolcontainerwindowlistener = OptionControlContainerWindowListener(optioncontrolcontainer)		
	optioncontrolcontainer.addWindowListener(optioncontrolcontainerwindowlistener)  # コントロールコンテナの大きさを変更するとグリッドコントロールの大きさも変更するようにする。
	mouselistener.optioncontrolcontainer = optioncontrolcontainer
	rectangle = controlcontainer.getPosSize()  # コントロールコンテナのRectangle Structを取得。px単位。
	if enhancedmouseevent is None:
		visibleareaonscreen = controller.getPropertyValue("VisibleAreaOnScreen")
		rectangle.X, rectangle.Y = visibleareaonscreen.X, visibleareaonscreen.Y
	else:
		dialogpoint = dialogcommons.getDialogPoint(doc, enhancedmouseevent)  # クリックした位置のメニューバーの高さ分下の位置を取得。単位ピクセル。一部しか表示されていないセルのときはNoneが返る。
		if not dialogpoint:  # クリックした位置が取得出来なかった時は何もしない。
			return
		rectangle.X, rectangle.Y = dialogpoint  # クリックした位置を取得。ウィンドウタイトルを含めない座標。
	taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
	args = NamedValue("PosSize", rectangle), NamedValue("FrameName", "controldialog")  # , NamedValue("MakeVisible", True)  # TaskCreatorで作成するフレームのコンテナウィンドウのプロパティ。
	dialogframe = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。
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
	dialogwindow.setVisible(True) # ウィンドウの表示。これ以降WindowListenerが発火する。
	windowlistener = WindowListener(controlcontainer, optioncontrolcontainer) # コンテナウィンドウからコントロールコンテナを取得する方法はないはずなので、ここで渡す。WindowListenerはsetVisible(True)で呼び出される。
	dialogwindow.addWindowListener(windowlistener) # コンテナウィンドウにリスナーを追加する。
	menulistener.args = dialogwindow, windowlistener, mouselistener
	dialogstate = dialogcommons.getSavedData(doc, "dialogstate_{}".format(dialogtitle))  # 保存データを取得。optioncontrolcontainerの表示状態は常にFalseなので保存されていない。
	if dialogstate is not None:  # 保存してあるダイアログの状態がある時。
		for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
			itemtext = gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
			if itemtext.startswith("セル入力で閉じる"):
				closecheck = dialogstate.get("CloseCheck")  # 保存してある、セル入力で閉じる、のチェックの状態を取得。
				if closecheck is not None:
					gridpopupmenu.checkItem(menuid, closecheck)
			elif itemtext.startswith("オプション表示"):
				optioncheck = dialogstate.get("OptionCheck")  # 保存してある、オプション表示、のチェックの状態を取得。
				if optioncheck is not None:
					gridpopupmenu.checkItem(menuid, optioncheck)  # ItemIDは1から始まる。これでMenuListenerは発火しない。
					if optioncheck:  # チェックが付いている時MenuListenerを発火させる。
						gridcontrol1.removeMouseMotionListener(mousemotionlistener)  # オプション表示している時は行のハイライトはしない。
						menulistener.itemSelected(MenuEvent(MenuId=menuid, Source=mouselistener.gridpopupmenu))
		checkbox1sate = dialogstate.get("CheckBox1sate")  # サイズ復元、チェックボックス。キーがなければNoneが返る。	
		if checkbox1sate is not None:  # サイズ復元、が保存されている時。
			if checkbox1sate:  # サイズ復元がチェックされている時。
				dialogwindow.setPosSize(0, 0, dialogstate["Width"], dialogstate["Height"], PosSize.SIZE)  # ウィンドウサイズを復元。
			checkboxcontrol1.setState(checkbox1sate)  # 状態を復元。	
		checkbox2sate = dialogstate.get("CheckBox2sate")  # 逐語検索、チェックボックス。			
		if checkbox2sate is not None:  # 逐語検索、が保存されている時。
			if checkbox2sate:  # チェックされている時逐次検索を有効にする。	
				refreshRows(gridcontrol1, [i for i in DATAROWS if i[0].startswith(txt)])  # txtで始まっている行だけに絞る。txtが空文字の時はすべてTrueになる。
			checkboxcontrol2.setState(checkbox2sate)  # itemlistenerは発火しない。	
	args = doc, actionlistener, dialogwindow, windowlistener, mouselistener, menulistener, textlistener, itemlistener, controlcontainerwindowlistener, optioncontrolcontainerwindowlistener, mousemotionlistener
	dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。		
	scrollDown(gridcontrol1)		
class ItemListener(unohelper.Base, XItemListener):
	def __init__(self, controlcontainer):
		self.controlcontainer = controlcontainer
	def itemStateChanged(self, itemevent):  
		checkboxcontrol2 = itemevent.Source
		gridcontrol1 = self.controlcontainer.getControl("Grid1")
		if checkboxcontrol2.getState():
			txt = checkboxcontrol2.getContext().getControl("Edit1").getText()
			refreshRows(gridcontrol1, [i for i in DATAROWS if i[0].startswith(txt)])  # txtで始まっている行だけに絞る。txtが空文字の時はすべてTrueになる。
		else:
			refreshRows(gridcontrol1, DATAROWS)
			scrollDown(gridcontrol1)	
	def disposing(self, eventobject):
		pass
def refreshRows(gridcontrol, datarows):	
	griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。	
	griddatamodel.removeAllRows()  # グリッドコントロールの行を全削除。
	if datarows:  # データ行がある時。
		griddatamodel.addRows(("",)*len(datarows), datarows)  # グリッドに行を追加。
class TextListener(unohelper.Base, XTextListener):
	def __init__(self, xscriptcontext, controlcontainer):
		self.controlcontainer = controlcontainer
		self.transliteration = fullwidth_halfwidth(xscriptcontext)
		self.history = ""  # 前値を保存する。
	def textChanged(self, textevent):  # 複数回呼ばれるので前値との比較が必要。
		editcontrol1 = textevent.Source
		txt = editcontrol1.getText()
		positionfromback = len(txt) - editcontrol1.getSelection().Max  # テキストカーソルの後ろからの位置を取得しておく。前からの位置は濁点、半濁点の分増えるので。
		if txt!=self.history:  # 前値から変化する時のみ。
			txt = self.transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換
			editcontrol1.removeTextListener(self)
			editcontrol1.setText(txt)  # 永久ループになるのでTextListenerを発火しないようにしておかないといけない。
			editcontrol1.addTextListener(self)
			if editcontrol1.getContext().getControl("CheckBox2").getState():  # 逐次検索が有効になっている時。
				gridcontrol1 = self.controlcontainer.getControl("Grid1")
				datarows = [i for i in DATAROWS if i[0].startswith(txt)]  # 逐語抽出した行のリスト。
				if len(datarows)==1:  # 行が一行だけになる時。	
					selectedrowindex = gridcontrol1.getCurrentRow()  # 選択行インデックスを取得。
					if selectedrowindex>0:  # 選択行インデックスが1以上の時。
						griddatamodel = gridcontrol1.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
						if selectedrowindex+1<griddatamodel.RowCount:  # 選択行より後に行がある時。
							[griddatamodel.removeRow(i) for i in range(selectedrowindex+1, griddatamodel.RowCount)[::-1]]  # 選択行の次から最後までを最後から削除する。
						return  # 選択行より前の行の削除は諦める。選択行より上の行を削除するとグリッドコントロール以外マウスに反応しなくなるので。ソートして一番上に持ってきてもダメ。
				refreshRows(gridcontrol1, datarows)
			self.history = txt	
			position = len(txt) - positionfromback  # 後ろの位置からカーソルの新しい位置を算出。
			edit1selection = Selection(Min=position, Max=position)  # カーソルの位置を取得。
			editcontrol1.setSelection(edit1selection)  # テクストボックスコントロールのカーソルの位置を変更。ピア作成後でないと反映されない。		
	def disposing(self, eventobject):
		pass
def fullwidth_halfwidth(xscriptcontext):
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。					
	transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
	transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))  # 全角を半角に変換するモジュール。
	return transliteration
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログのリスナー削除用。
	def __init__(self, args):
		self.args = args
	def queryClosing(self, eventobject, getsownership):  # ノンモダルダイアログを閉じる時に発火。
		dialogframe = eventobject.Source
		doc, actionlistener, dialogwindow, windowlistener, mouselistener, menulistener, textlistener, itemlistener, controlcontainerwindowlistener, optioncontrolcontainerwindowlistener, mousemotionlistener = self.args
		controlcontainer, optioncontrolcontainer = windowlistener.args
		dialogwindowsize = dialogwindow.getSize()
		checkboxcontrol2 = optioncontrolcontainer.getControl("CheckBox2")
		checkboxcontrol2.removeItemListener(itemlistener)			
		dialogstate = {"CheckBox1sate": optioncontrolcontainer.getControl("CheckBox1").getState(),\
					"CheckBox2sate": checkboxcontrol2.getState(),\
					"Width": dialogwindowsize.Width,\
					"Height": dialogwindowsize.Height}  # チェックボックスの状態と大きさを取得。
		gridpopupmenu = mouselistener.gridpopupmenu
		for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
			itemtext = gridpopupmenu.getItemText(menuid)
			if itemtext.startswith("セル入力で閉じる"):
				dialogstate.update({"CloseCheck": gridpopupmenu.isItemChecked(menuid)})
			elif itemtext.startswith("オプション表示"):
				dialogstate.update({"OptionCheck": gridpopupmenu.isItemChecked(menuid)})		
		dialogtitle = dialogframe.getTitle()
		dialogcommons.saveData(doc, "dialogstate_{}".format(dialogtitle), dialogstate)
		dialogcommons.saveData(doc, "GridDatarows_{}".format(dialogtitle), DATAROWS)
		gridpopupmenu.removeMenuListener(menulistener)
		gridcontrol1 = controlcontainer.getControl("Grid1")
		gridcontrol1.removeMouseListener(mouselistener)
		gridcontrol1.removeMouseMotionListener(mousemotionlistener)
		controlcontainer.removeWindowListener(controlcontainerwindowlistener)
		optioncontrolcontainer.getControl("Button1").removeActionListener(actionlistener)
		optioncontrolcontainer.getControl("Edit1").removeTextListener(textlistener)
		optioncontrolcontainer.removeWindowListener(optioncontrolcontainerwindowlistener)		
		dialogwindow.removeWindowListener(windowlistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		pass
class ActionListener(unohelper.Base, XActionListener):
	def __init__(self, args):
		xscriptcontext, self.controlcontainer = args
		self.transliteration = fullwidth_halfwidth(xscriptcontext)  # xscriptcontextを渡す。
	def actionPerformed(self, actionevent):  
		cmd = actionevent.ActionCommand
		if cmd=="enter":
			optioncontrolcontainer = actionevent.Source.getContext()			
			edit1 = optioncontrolcontainer.getControl("Edit1")  # テキストボックスコントロールを取得。
			txt = edit1.getText()  # テキストボックスコントロールの文字列を取得。
			if txt:  # テキストボックスコントロールに文字列がある時。
				global DATAROWS
				txt = self.transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換
				datarows = DATAROWS
				if datarows:  # すでにグリッドコントロールにデータがある時。
					lastindex = len(datarows) - 1  # 最終インデックスを取得。
					[datarows.pop(lastindex-i) for i, datarow in enumerate(datarows[::-1]) if txt in datarow]  # txtがある行は後ろから削除する。
				datarows.append((txt,))  # txtの行を追加。
				gridcontrol1 = self.controlcontainer.getControl("Grid1")
				refreshRows(gridcontrol1, datarows)
				scrollDown(gridcontrol1)  # グリッドコントロールを下までスクロール。
				DATAROWS = datarows				
	def disposing(self, eventobject):
		pass
class MouseListener(unohelper.Base, XMouseListener):  
	def __init__(self, args): 	
		self.gridpopupmenu, *self.args = args  # gridpopupmenuはCloseListenerで使うので、別にする。
		self.optioncontrolcontainer = None
		self.dialogframe = None
		self.flg = True  # 付け直した時にmousePressed()が発火しないようにするフラグ。		
	def mousePressed(self, mouseevent):  # グリッドコントロールをクリックした時。コントロールモデルにはNameプロパティはない。
		gridcontrol = mouseevent.Source  # グリッドコントロールを取得。
		if mouseevent.Buttons==MouseButton.LEFT:  # オプション表示しているときはダブルクリック、そうでない時はシングルクリックでセル入力する。
			selectedrowindexes = dialogcommons.getSelectedRowIndexes(gridcontrol)
			if mouseevent.ClickCount==1:  # シングルクリックの時。
				if self.flg:
					for menuid in range(1, self.gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
						itemtext = self.gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
						if itemtext.startswith("オプション表示"):
							if not self.gridpopupmenu.isItemChecked(menuid):  # 選択項目にチェックが入っていない時。
								if selectedrowindexes:  # 選択行がある時。
									self._toCell(gridcontrol, selectedrowindexes)  # オプション表示していない時はシングルクリックでセルに入力する。
									break						
					else:  # 選択項目にチェックが入っている時、オプション表示部分の設定。
						if selectedrowindexes:  # 選択行がある時。
							griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。
							rowdata = griddatamodel.getRowData(selectedrowindexes[0])  # 選択行の最初の行のデータを取得。
							txt = rowdata[0]
						else:
							txt = ""  # 選択行がない時は空文字にする。			
						self.optioncontrolcontainer.getControl("Edit1").setText(txt)  # テキストボックスに選択行の初行の文字列を代入。
				else:
					self.flg = True
			elif mouseevent.ClickCount==2:  # ダブルクリックの時。
				self._toCell(gridcontrol, selectedrowindexes)
		elif mouseevent.Buttons==MouseButton.RIGHT:  # 右ボタンクリックの時。mouseevent.PopupTriggerではサブジェクトによってはTrueにならないので使わない。
			rowindex = gridcontrol.getRowAtPoint(mouseevent.X, mouseevent.Y)  # クリックした位置の行インデックスを取得。該当行がない時は-1が返ってくる。
			if rowindex>-1:  # クリックした位置に行が存在する時。
				if not gridcontrol.isRowSelected(rowindex):  # クリックした位置の行が選択状態でない時。
					gridcontrol.deselectAllRows()  # 行の選択状態をすべて解除する。
					gridcontrol.selectRow(rowindex)  # 右クリックしたところの行を選択する。		
			pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
			self.gridpopupmenu.execute(gridcontrol.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向					
	def _toCell(self, gridcontrol, selectedrowindexes):  # callback関数で指定した行をマウスで選択し直さないとgetCurrentRow()では0が返ってしまうのでselectedrowindexesも受け取る。
		xscriptcontext, outputcolumn, callback = self.args
		doc = xscriptcontext.getDocument()
		selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
			if len(selectedrowindexes)==1 and selectedrowindexes[0]>-1:  # グリッドコントロールの選択行インデックスが1つ、かつ、0以上の時のみ。
				j = selectedrowindexes[0]  # グリッドコントロールの選択行インデックスを取得。
				griddata = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。グリッドコントロールは1列と決めつけて処理する。
				rowdata = griddata.getRowData(j)  # グリッドコントロールで選択している行のすべての列をタプルで取得。
				controller = doc.getCurrentController()  # 現在のコントローラを取得。			
				sheet = controller.getActiveSheet()
				celladdress = selection.getCellAddress()
				r, c = celladdress.Row, celladdress.Column
				if outputcolumn is not None:  # 出力する列が指定されている時。
					c = outputcolumn  # 同じ行の指定された列のセルに入力するようにする。
				sheet[r, c].setString(rowdata[0])  # セルに代入。
				if callback is not None:  # コールバック関数が与えられている時。
					try:
						callback(rowdata[0])		
					except:  # これをしないとエラーダイアログが出てこない。
						exceptiondialog2.createDialog(xscriptcontext)  # XSCRIPTCONTEXTを渡す。					
				global DATAROWS
				datarows = DATAROWS
				if datarows:  # すでにグリッドコントロールにデータがある時。
					lastindex = len(datarows) - 1  # 最終インデックスを取得。
					[datarows.pop(lastindex-i) for i, datarow in enumerate(datarows[::-1]) if rowdata[0] in datarow]  # 同じデータがある行は後ろから削除する。
				datarows.append((rowdata[0],))  # 最後に追加し直す。
				refreshRows(gridcontrol, datarows)
				scrollDown(gridcontrol)  # グリッドコントロールを下までスクロール。
				DATAROWS = datarows			
				gridpopupmenu = self.gridpopupmenu	
				for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
					itemtext = gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
					if itemtext.startswith("セル入力で閉じる"):
						if gridpopupmenu.isItemChecked(menuid):  # 選択項目にチェックが入っている時。
							self.dialogframe.close(True)  # gridcontrolのMouseListenerを外しておかないとクラッシュする。
							break						
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		pass
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):
		pass
class MenuListener(unohelper.Base, XMenuListener):
	def __init__(self, mousemotionlistener):
		self.mousemotionlistener = mousemotionlistener
		self.args = None
	def itemHighlighted(self, menuevent):
		pass
	def itemSelected(self, menuevent):  # PopupMenuの項目がクリックされた時。どこのコントロールのメニューかを知る方法はない。
		menuid = menuevent.MenuId  # メニューIDを取得。1から始まる。
		dialogwindow, windowlistener, mouselistener = self.args
		controlcontainer, optioncontrolcontainer = windowlistener.args
		mousemotionlistener = self.mousemotionlistener
		gridpopupmenu = menuevent.Source
		itemtext = gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
		gridcontrol1 = controlcontainer.getControl("Grid1")
		if itemtext.startswith("オプション表示"):	
			dialogwindowsize = dialogwindow.getSize()
			optioncontrolcontainersize = optioncontrolcontainer.getSize()		
			if gridpopupmenu.isItemChecked(menuid):  # 選択項目にチェックが入った時。
				windowlistener.option = True  # オプションコントロールダイアログを表示させるフラグを立てる。
				diff_width = optioncontrolcontainersize.Width - dialogwindowsize.Width  # オプションコントロールコンテナ幅とコンテナウィンドウ幅の差。
				diff_width = 0 if diff_width<0 else diff_width  # オプションコントロールコンテナ幅よりコンテナウィンドウ幅が大きい時は幅の調整をしない。
				diff_height = optioncontrolcontainersize.Height  # オプションコントロールコンテナの高さを追加する。
				dialogcommons.createApplyDiff(diff_width, diff_height)(dialogwindow, PosSize.SIZE)  # コンテナウィンドウの大きさを変更。
			else:  # 選択項目にチェックが外れた時。
				windowlistener.option = False  # オプションコントロールダイアログを表示させるフラグを倒す。
				diff_height = -optioncontrolcontainersize.Height  # オプションコントロールコンテナの高さを減らす。
				dialogcommons.createApplyDiff(0, diff_height)(dialogwindow, PosSize.HEIGHT)  # コンテナウィンドウの大きさを変更。	
				gridcontrol1.addMouseMotionListener(mousemotionlistener)
		else:
			cmd = menuevent.Source.getCommand(menuevent.MenuId)
			global DATAROWS
			datarows = list(DATAROWS)
			peer = controlcontainer.getPeer()  # ピアを取得。	
			gridcontrol = controlcontainer.getControl("Grid1")  # グリッドコントロールを取得。
			griddatamodel = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。		
			selectedrowindexes = dialogcommons.getSelectedRowIndexes(gridcontrol)	 # 選択行のリストを取得。
			if not selectedrowindexes:
				return  # 選択行がない時何もしない。
			if cmd=="delete":  # 選択行を削除する。  
				msg = "選択行を削除しますか?"
				msgbox = peer.getToolkit().createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "履歴", msg)
				if msgbox.execute()==MessageBoxResults.YES:		
					if griddatamodel.RowCount==len(datarows):  # グリッドコントロールとDATAROWSの行数が一致している時。
						[datarows.pop(i) for i in selectedrowindexes[::-1]]  # 後ろから選択行を削除。
						refreshRows(gridcontrol, datarows)
					else:
						for i in selectedrowindexes[::-1]:  # 選択した行インデックスを後ろから取得。逐次検索のときはグリッドコントロールとDATAROWSが一致しないので別に処理する。
							d = griddatamodel.getRowData(i)[0]  # タプルが返るのでその先頭の要素を取得。
							datarows = [j for j in datarows if not d in j]  # dが要素にある行を除いて取得。
							griddatamodel.removeRow(i)  # グリッドコントロールから選択行を削除。
			elif cmd=="deleteall":  # 全行を削除する。  	
				msg = "表示しているすべての行を削除しますか?"
				msgbox = peer.getToolkit().createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "履歴", msg)
				if msgbox.execute()==MessageBoxResults.YES:		
					msg = "本当に表示しているすべての行を削除しますか？\n削除したデータは取り戻せません。"
					msgbox = peer.getToolkit().createMessageBox(peer, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO, "履歴", msg)				
					if msgbox.execute()==MessageBoxResults.YES:	
						if griddatamodel.RowCount==len(datarows):  # グリッドコントロールとDATAROWSの行数が一致している時。
							griddatamodel.removeAllRows()  # グリッドコントロールの行を全削除。
							datarows.clear()  # 全データ行をクリア。	
						else:
							gridcontrol.selectAllRows()  # すべての行を選択。
							for i in selectedrowindexes[::-1]:  # 選択した行インデックスを後ろから取得。逐次検索のときはグリッドコントロールとDATAROWSが一致しないので別に処理する。
								d = griddatamodel.getRowData(i)[0]  # タプルが返るのでその先頭の要素を取得。
								datarows = [j for j in datarows if not d in j]  # dが要素にある行を除いて取得。
							griddatamodel.removeAllRows()  # グリッドコントロールの行を全削除。						
			DATAROWS = datarows					
		mouselistener.flg = False		
		gridcontrol1.addMouseListener(mouselistener)  # ポップアップメニューを表示する時に外したMouseListenerを付け直す。つけた時点でmousePressed()が発火する。
	def itemActivated(self, menuevent):  # メニュー項目を有効にした時。チェックボックスをオンにした時ではない。ポップアップメニューをexecute()した時に発火する。
		dummy, windowlistener, mouselistener = self.args
		controlcontainer, dummy = windowlistener.args
		gridcontrol1 = controlcontainer.getControl("Grid1")
		gridcontrol1.removeMouseMotionListener(self.mousemotionlistener)  # ポップアップメニュー上で動くとMouseMotionLitenerも発火するので外しておく。
		gridcontrol1.removeMouseListener(mouselistener)
	def itemDeactivated(self, menuevent):  # メニュー項目が無効になった時。ポップアップメニュー項目を選択せずに閉じた時も発火する。
		dummy, windowlistener, mouselistener = self.args
		controlcontainer, dummy = windowlistener.args
		gridcontrol1 = controlcontainer.getControl("Grid1")
		mouselistener.flg = False
		gridcontrol1.addMouseListener(mouselistener)  # ポップアップメニューを表示する時に外したMouseListenerを付け直す。つけた時点でmousePressed()が発火する。
		gridpopupmenu = menuevent.Source
		for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
			itemtext = gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
			if itemtext.startswith("オプション表示"):
				if not gridpopupmenu.isItemChecked(menuid):  # 選択項目にチェックが入っていない時。
					gridcontrol1.addMouseMotionListener(self.mousemotionlistener)
					break							 
	def disposing(self, eventobject):
		pass		
class WindowListener(unohelper.Base, XWindowListener):
	def __init__(self, *args):
		self.args = args
		self.option = False  # optioncontrolcontainerを表示しているかのフラグ。
	def windowResized(self, windowevent):
		controlcontainer, optioncontrolcontainer = self.args
		if self.option:  # optioncontrolcontainerを表示している時。
			optioncontrolcontainer.setVisible(True)
			newwidth, newheight = windowevent.Width, windowevent.Height
			controlcontainerheight = newheight - optioncontrolcontainer.getSize().Height  # オプションコントロールコンテナの高さを除いた高さを取得。
			optioncontrolcontainer.setPosSize(0, controlcontainerheight, newwidth, 0, PosSize.Y+PosSize.WIDTH)
			controlcontainer.setPosSize(0, 0, newwidth, controlcontainerheight, PosSize.SIZE)
		else:
			optioncontrolcontainer.setVisible(False)
			controlcontainer.setPosSize(0, 0, windowevent.Width, windowevent.Height, PosSize.SIZE)
	def windowMoved(self, windowevent):
		pass
	def windowShown(self, eventobject):
		pass
	def windowHidden(self, eventobject):
		pass
	def disposing(self, eventobject):
		pass
class ControlContainerWindowListener(unohelper.Base, XWindowListener):
	def __init__(self, controlcontainer):
		size = controlcontainer.getSize()
		self.oldwidth, self.oldheight = size.Width, size.Height  # 次の変更前の値として取得。		
		self.controlcontainer = controlcontainer
	def windowResized(self, windowevent):
		newwidth, newheight = windowevent.Width, windowevent.Height
		gridcontrol1 = self.controlcontainer.getControl("Grid1")
		diff_width = newwidth - self.oldwidth  # 幅変化分
		diff_height = newheight - self.oldheight  # 高さ変化分		
		dialogcommons.createApplyDiff(diff_width, diff_height)(gridcontrol1, PosSize.SIZE)  # コントロールの位置と大きさを変更		
		self.oldwidth, self.oldheight = newwidth, newheight  # 次の変更前の値として取得。
		scrollDown(gridcontrol1)		
	def windowMoved(self, windowevent):
		pass
	def windowShown(self, eventobject):
		pass
	def windowHidden(self, eventobject):
		pass
	def disposing(self, eventobject):
		pass
class OptionControlContainerWindowListener(unohelper.Base, XWindowListener):
	def __init__(self, optioncontrolcontainer):
		self.oldwidth = optioncontrolcontainer.getSize().Width  # 次の変更前の値として取得。		
		self.optioncontrolcontainer = optioncontrolcontainer
	def windowResized(self, windowevent): # ウィンドウの大きさの変更に合わせてコントロールの位置と大きさを変更。Yと幅のみ変更。
		optioncontrolcontainer = self.optioncontrolcontainer
		newwidth = windowevent.Width
		checkboxcontrol1 = optioncontrolcontainer.getControl("CheckBox1")
		checkboxcontrol2 = optioncontrolcontainer.getControl("CheckBox2")
		buttoncontrol1 = optioncontrolcontainer.getControl("Button1")
		checkbox1rect = checkboxcontrol1.getPosSize()  # hをHeightから取得。
		minwidth = checkbox1rect.Width + checkboxcontrol2.getPosSize().Width + buttoncontrol1.getSize().Width  # 幅下限を取得。
		if newwidth<minwidth:  # 変更後のコントロールコンテナの幅を取得。サイズ下限より小さい時は下限値とする。
			newwidth = minwidth
		diff_width = newwidth - self.oldwidth  # 幅変化分
		applyDiff = dialogcommons.createApplyDiff(diff_width, 0)  # コントロールの位置と大きさを変更する関数を取得。
		applyDiff(optioncontrolcontainer.getControl("Edit1"), PosSize.Y+PosSize.WIDTH)
		applyDiff(checkboxcontrol1, PosSize.Y)
		applyDiff(checkboxcontrol2, PosSize.Y)
		applyDiff(buttoncontrol1, PosSize.POS)		
		self.oldwidth = newwidth  # 次の変更前の値として取得。
	def windowMoved(self, windowevent):
		pass
	def windowShown(self, eventobject):
		pass
	def windowHidden(self, eventobject):
		pass
	def disposing(self, eventobject):
		pass	
def scrollDown(gridcontrol):  # グリッドコントロールを下までスクロールする。		
	accessiblecontext = gridcontrol.getAccessibleContext()  # グリッドコントロールのAccessibleContextを取得。
	for i in range(accessiblecontext.getAccessibleChildCount()):  # 子要素のインデックスを走査する。
		child = accessiblecontext.getAccessibleChild(i)  # 子要素を取得。
		if child.getAccessibleContext().getAccessibleRole()==AccessibleRole.SCROLL_BAR:  # スクロールバーの時。
			if child.getOrientation()==ScrollBarOrientation.VERTICAL:  # 縦スクロールバーの時。
				child.setValue(0)  # 一旦0にしないといけない？
				child.setValue(child.getMaximum())  # 最大値にスクロールさせる。
				break				
