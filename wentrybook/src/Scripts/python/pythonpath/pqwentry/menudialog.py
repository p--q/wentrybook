#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from . import dialogcommons, staticdialog  # staticdialogのオブジェクトを複数流用している。
from com.sun.star.awt import MenuItemStyle, MouseButton, PopupMenuDirection, PosSize  # 定数
from com.sun.star.awt import MenuEvent, Rectangle  # Struct
from com.sun.star.beans import NamedValue  # Struct
def createDialog(xscriptcontext, dialogtitle, defaultrows, *, enhancedmouseevent=None, callback=None):  # dialogtitleはダイアログのデータ保存名に使うのでユニークでないといけない。defaultrowsはグリッドコントロールのデフォルトデータ。
# 	# 一番最初のダイアログのオプション設定。
	items = ("オプション表示", MenuItemStyle.CHECKABLE+MenuItemStyle.AUTOCHECK, {"checkItem": False}),  # グリッドコントロールのコンテクストメニュー。XMenuListenerのmenuevent.MenuIdでコードを実行する。	
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	doc = xscriptcontext.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。  
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	containerwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	maTopx = dialogcommons.createConverters(containerwindow)  # ma単位をピクセルに変換する関数を取得。
	m = 2  # コントロール間の間隔。
	h = 12  # コントロールの高さ
	XWidth, YHeight = dialogcommons.XWidth, dialogcommons.YHeight
	gridprops = {"PositionX": 0, "PositionY": 0, "Width": 50, "Height": 50, "ShowRowHeader": False, "ShowColumnHeader": False}  # グリッドコントロールのプロパティ。
	controlcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(gridprops), "Height": YHeight(gridprops), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
	controlcontainer, addControl = dialogcommons.controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。		
	mousemotionlistener = dialogcommons.MouseMotionListener()
	menulistener = staticdialog.MenuListener(mousemotionlistener)  # コンテクストメニューにつけるリスナー。
	gridpopupmenu = dialogcommons.menuCreator(ctx, smgr)("PopupMenu", items, {"addMenuListener": menulistener, "hideDisabledEntries": False})  # 右クリックでまず呼び出すポップアップメニュー。hideDisabledEntries()が反応しない。  
	args = gridpopupmenu, xscriptcontext, callback  # gridpopupmenuは先頭でないといけない。
	mouselistener = MouseListener(args)
	gridcontrol1 = addControl("Grid", gridprops, {"addMouseListener": mouselistener, "addMouseMotionListener": mousemotionlistener})  # グリッドコントロールの取得。
	gridmodel = gridcontrol1.getModel()  # グリッドコントロールモデルの取得。
	gridcolumn = gridmodel.getPropertyValue("ColumnModel")  # DefaultGridColumnModel
	gridcolumn.addColumn(gridcolumn.createColumn())  # 列を追加。
	griddatamodel = gridmodel.getPropertyValue("GridDataModel")  # GridDataModel
	if defaultrows is not None:  # デフォルトdatarowsがあるときデフォルトデータを使用。	
		datarows = [i if isinstance(i, (list, tuple)) else (i,) for i in defaultrows]  # defaultrowsの要素をリストかタプルでなければタプルに変換する。
		griddatamodel.addRows(("",)*len(datarows), datarows)  # グリッドに行を追加。	
	else:
		datarows = []  # Noneのままではあとで処理できないので空リストを入れる。
	controlcontainerwindowlistener = staticdialog.ControlContainerWindowListener(controlcontainer)		
	controlcontainer.addWindowListener(controlcontainerwindowlistener)  # コントロールコンテナの大きさを変更するとグリッドコントロールの大きさも変更するようにする。
	checkboxprops1 = {"PositionX": 0, "PositionY": m, "Width": 46, "Height": h, "Label": "~サイズ復元", "State": 1}  # サイズ復元はデフォルトでは有効。		
	optioncontrolcontainerprops = {"PositionX": 0, "PositionY": 0, "Width": XWidth(checkboxprops1), "Height": YHeight(checkboxprops1, 2), "BackgroundColor": 0xF0F0F0}  # コントロールコンテナの基本プロパティ。幅は右端のコントロールから取得。高さはコントロール追加後に最後に設定し直す。		
	optioncontrolcontainer, optionaddControl = dialogcommons.controlcontainerMaCreator(ctx, smgr, maTopx, optioncontrolcontainerprops)  # コントロールコンテナの作成。		
	checkboxcontrol2 = optionaddControl("CheckBox", checkboxprops1)  
	mouselistener.optioncontrolcontainer = optioncontrolcontainer
	rectangle = controlcontainer.getPosSize()  # コントロールコンテナのRectangle Structを取得。px単位。
	controller = doc.getCurrentController()  # 現在のコントローラを取得。
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
	windowlistener = staticdialog.WindowListener(controlcontainer, optioncontrolcontainer) # コンテナウィンドウからコントロールコンテナを取得する方法はないはずなので、ここで渡す。WindowListenerはsetVisible(True)で呼び出される。
	dialogwindow.addWindowListener(windowlistener) # コンテナウィンドウにリスナーを追加する。
	menulistener.args = dialogwindow, windowlistener, mouselistener
	dialogstate = dialogcommons.getSavedData(doc, "dialogstate_{}".format(dialogtitle))  # 保存データを取得。optioncontrolcontainerの表示状態は常にFalseなので保存されていない。
	if dialogstate is not None:  # 保存してあるダイアログの状態がある時。
		for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
			itemtext = gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
			if itemtext.startswith("オプション表示"):
				optioncheck = dialogstate.get("OptionCheck")  # オプション表示、のチェックがある時。
				if optioncheck is not None:
					gridpopupmenu.checkItem(menuid, optioncheck)  # ItemIDは1から始まる。これでMenuListenerは発火しない。
					if optioncheck:  # チェックが付いている時MenuListenerを発火させる。
						menulistener.itemSelected(MenuEvent(MenuId=menuid, Source=mouselistener.gridpopupmenu))
		checkbox1sate = dialogstate.get("CheckBox1sate")  # サイズ復元、チェックボックス。	
		if checkbox1sate is not None:  # サイズ復元、が保存されている時。
			checkboxcontrol2.setState(checkbox1sate)  # 状態を復元。	
			if checkbox1sate:  # サイズ復元がチェックされている時。
				dialogwindow.setPosSize(0, 0, dialogstate["Width"], dialogstate["Height"], PosSize.SIZE)  # ウィンドウサイズを復元。WindowListenerが発火する。
	args = doc, dialogwindow, windowlistener, mouselistener, menulistener, controlcontainerwindowlistener, mousemotionlistener
	dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。	
class CloseListener(staticdialog.CloseListener):  # ノンモダルダイアログのリスナー削除用。	
	def queryClosing(self, eventobject, getsownership):  # ノンモダルダイアログを閉じる時に発火。
		dialogframe = eventobject.Source
		doc, dialogwindow, windowlistener, mouselistener, menulistener, controlcontainerwindowlistener, mousemotionlistener = self.args
		controlcontainer, optioncontrolcontainer = windowlistener.args
		dialogwindowsize = dialogwindow.getSize()
		dialogstate = {"CheckBox1sate": optioncontrolcontainer.getControl("CheckBox1").getState(),\
					"Width": dialogwindowsize.Width,\
					"Height": dialogwindowsize.Height}  # チェックボックスコントロールの状態とコンテナウィンドウの大きさを取得。
		gridpopupmenu = mouselistener.gridpopupmenu
		for menuid in range(1, gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
			itemtext = gridpopupmenu.getItemText(menuid)
			if itemtext.startswith("オプション表示"):
				dialogstate.update({"OptionCheck": gridpopupmenu.isItemChecked(menuid)})
		dialogtitle = dialogframe.getTitle()  # コンテナウィンドウタイトルを取得。データ保存のIDに使う。
		dialogcommons.saveData(doc, "dialogstate_{}".format(dialogtitle), dialogstate)  # ダイアログの状態を保存。
		gridpopupmenu.removeMenuListener(menulistener)
		gridcontrol1 = controlcontainer.getControl("Grid1")
		gridcontrol1.removeMouseListener(mouselistener)
		gridcontrol1.removeMouseMotionListener(mousemotionlistener)
		controlcontainer.removeWindowListener(controlcontainerwindowlistener)
		dialogwindow.removeWindowListener(windowlistener)
		eventobject.Source.removeCloseListener(self)
class MouseListener(staticdialog.MouseListener):  	
	def mousePressed(self, mouseevent):  # グリッドコントロールをクリックした時。コントロールモデルにはNameプロパティはない。
		gridcontrol = mouseevent.Source  # グリッドコントロールを取得。
		if mouseevent.Buttons==MouseButton.LEFT:
			selectedrowindexes = dialogcommons.getSelectedRowIndexes(gridcontrol)
			if not selectedrowindexes:  # 選択行がない時(選択行を削除した時)。
				return  # 何もしない					
			if mouseevent.ClickCount==1:  # シングルクリックの時。
				if self.flg:
					for menuid in range(1, self.gridpopupmenu.getItemCount()+1):  # ポップアップメニューを走査する。
						itemtext = self.gridpopupmenu.getItemText(menuid)  # 文字列にはショートカットキーがついてくる。
						if itemtext.startswith("オプション表示"):
							if not self.gridpopupmenu.isItemChecked(menuid):  # 選択項目にチェックが入っていない時。
								self._toCell(gridcontrol, selectedrowindexes)  # オプション表示していない時はシングルクリックでセルに入力する。
								break	
				else:
					self.flg = True				
			elif mouseevent.ClickCount==2:  # ダブルクリックの時。
				self._toCell(gridcontrol, selectedrowindexes)						
		elif mouseevent.Buttons==MouseButton.RIGHT:  # 右ボタンクリックの時。mouseevent.PopupTriggerではサブジェクトによってはTrueにならないので使わない。
			pos = Rectangle(mouseevent.X, mouseevent.Y, 0, 0)  # ポップアップメニューを表示させる起点。
			self.gridpopupmenu.execute(gridcontrol.getPeer(), pos, PopupMenuDirection.EXECUTE_DEFAULT)  # ポップアップメニューを表示させる。引数は親ピア、位置、方向							
	def _toCell(self, gridcontrol, selectedrowindexes):  # callback関数で指定した行をマウスで選択し直さないとgetCurrentRow()では0が返ってしまうのでselectedrowindexesも受け取る。
		self.dialogframe.close(True)  # ダイアログを先に閉じてしまう。そうしないとコールバック関数でメッセージボックスを使うとダイアログが残ってしまう。	
		xscriptcontext, callback = self.args
		doc = xscriptcontext.getDocument()
		selection = doc.getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択オブジェクトがセルの時。
			if len(selectedrowindexes)==1 and selectedrowindexes[0]>-1:  # グリッドコントロールの選択行インデックスが1つ、かつ、0以上の時のみ。
				j = selectedrowindexes[0]  # グリッドコントロールの選択行インデックスを取得。
				griddata = gridcontrol.getModel().getPropertyValue("GridDataModel")  # GridDataModelを取得。グリッドコントロールは1列と決めつけて処理する。
				rowdata = griddata.getRowData(j)  # グリッドコントロールで選択している行のすべての列をタプルで取得。
				if callback is not None:  # コールバック関数が与えられている時。
					callback(rowdata[0])			
				
	