#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# embeddedmacro.pyから呼び出した関数ではXSCRIPTCONTEXTは使えない。デコレーターも使えない。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)でブレークする。
import unohelper  # オートメーションには必須(必須なのはuno)。
from . import commons, exceptiondialog2
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.document import XDocumentEventListener
from com.sun.star.sheet import XActivationEventListener
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui.ContextMenuInterceptorAction import IGNORED  # enum
from com.sun.star.util import XChangesListener
from com.sun.star.view import XSelectionChangeListener
def invokeModuleMethod(name, methodname, *args):  # commons.getModle()でモジュールを振り分けてそのモジュールのmethodnameのメソッドを引数argsで呼び出す。
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # ここでブレークするとすべてのイベントでブレークすることになる。
	try:
		m = commons.getModule(name)  # モジュールを取得。
		if hasattr(m, methodname):  # モジュールにmethodnameの関数が存在する時。	
			return getattr(m, methodname)(*args)  # その関数を実行。
		return None  # メソッドが見つからなかった時はNoneを返す。ハンドラやインターセプターは戻り値の処理が必ず必要。
	except:  # UNO APIのメソッド以外のエラーはダイアログがでないのでここで捉える。
		exceptiondialog2.createDialog(args[-1])  # XSCRIPTCONTEXTを渡す。
def addLinsteners(tdocimport, modulefolderpath, xscriptcontext):  # 引数は文書のイベント駆動用。
	invokeModuleMethod(None, "documentOnLoad", xscriptcontext)  # ドキュメントを開いた時に実行するメソッド。リスナー追加前（リスナー追加後であってもリスナーは発火しない模様)。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	changeslistener = ChangesListener(xscriptcontext)  # ChangesListener。セルの変化の感知に利用。列の挿入も感知。
	selectionchangelistener = SelectionChangeListener(xscriptcontext)  # SelectionChangeListener。選択範囲の変更の感知に利用。
	activationeventlistener = ActivationEventListener(xscriptcontext, selectionchangelistener)  # ActivationEventListener。シートの切替の感知に利用。selectionchangelistenerを無効にするために渡す。
	enhancedmouseclickhandler = EnhancedMouseClickHandler(xscriptcontext)  # EnhancedMouseClickHandler。マウスの左クリックの感知に利用。enhancedmouseeventのSourceはNone。
	contextmenuinterceptor = ContextMenuInterceptor(xscriptcontext)  # ContextMenuInterceptor。右クリックメニューの変更に利用。
	doc.addChangesListener(changeslistener)
	controller.addSelectionChangeListener(selectionchangelistener)
	controller.addActivationEventListener(activationeventlistener)
	controller.addEnhancedMouseClickHandler(enhancedmouseclickhandler)
	controller.registerContextMenuInterceptor(contextmenuinterceptor)
	listeners = changeslistener, selectionchangelistener, activationeventlistener, enhancedmouseclickhandler, contextmenuinterceptor
	doc.addDocumentEventListener(DocumentEventListener(xscriptcontext, tdocimport, modulefolderpath, controller, *listeners))  # DocumentEventListener。ドキュメントとコントローラに追加したリスナーの除去に利用。
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, xscriptcontext, *args):
		self.xscriptcontext = xscriptcontext
		self.args = args
	def documentEventOccured(self, documentevent):
		eventname = documentevent.EventName
		if eventname=="OnUnload":  # ドキュメントを閉じる時。リスナーを除去する。
			tdocimport, modulefolderpath, controller, changeslistener, selectionchangelistener, activationeventlistener, enhancedmouseclickhandler, contextmenuinterceptor = self.args
			tdocimport.remove_meta(modulefolderpath)  # modulefolderpathをメタパスから除去する。
			documentevent.Source.removeChangesListener(changeslistener)
			controller.removeSelectionChangeListener(selectionchangelistener)
			controller.removeActivationEventListener(activationeventlistener)
			controller.removeEnhancedMouseClickHandler(enhancedmouseclickhandler)
			controller.releaseContextMenuInterceptor(contextmenuinterceptor)
			invokeModuleMethod(None, "documentUnLoad", self.xscriptcontext)  # ドキュメントを閉じた時に実行するメソッド。
	def disposing(self, eventobject):  # ドキュメントを閉じるときに発火する。	
		eventobject.Source.removeDocumentEventListener(self)
class ActivationEventListener(unohelper.Base, XActivationEventListener):
	def __init__(self, xscriptcontext, selectionchangelistener):
		self.xscriptcontext = xscriptcontext
		self.selectionchangelistener = selectionchangelistener
	def activeSpreadsheetChanged(self, activationevent):  # アクティブシートが変化した時に発火。
		controller = activationevent.Source
		controller.removeSelectionChangeListener(self.selectionchangelistener)  # シートを切り替えた時はselectionchangelistenerが発火しないようにSelectionChangeListenerをはずす。
		invokeModuleMethod(activationevent.ActiveSheet.getName(), "activeSpreadsheetChanged", activationevent, self.xscriptcontext)
		controller.addSelectionChangeListener(self.selectionchangelistener)  # SelectionChangeListenerを付け直す。
	def disposing(self, eventobject):
		eventobject.Source.removeActivationEventListener(self)	
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):  # enhancedmouseeventのSourceはNoneなので、このリスナーのメソッドの引数からコントローラーを直接取得する方法はない。
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def mousePressed(self, enhancedmouseevent):  # セルをクリックした時に発火する。固定行列の最初のクリックは同じ相対位置の固定していないセルが返ってくる(表示されている自由行の先頭行に背景色がる時のみ）。
		target = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if target.supportsService("com.sun.star.sheet.SheetCellRange"):  # targetがチャートの時がありうる?
			b = invokeModuleMethod(target.getSpreadsheet().getName(), "mousePressed", enhancedmouseevent, self.xscriptcontext)  # 正しく実行されれば、ブーリアンが返ってくるはず。
			if b is not None:
				return b
		return True  # Falseを返すと右クリックメニューがでてこなくなる。		
	def mouseReleased(self, enhancedmouseevent):
		return True  # シングルクリックでFalseを返すとセル選択範囲の決定の状態になってどうしようもなくなる。
	def disposing(self, eventobject):  # eventobject.SourceはNone。
		self.xscriptcontext.getDocument().getCurrentController().removeEnhancedMouseClickHandler(self)
class SelectionChangeListener(unohelper.Base, XSelectionChangeListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
		self.selectionrangeaddress = None  # selectionChanged()メソッドが何回も無駄に発火するので選択範囲アドレスのStructをキャッシュして比較する。
	def selectionChanged(self, eventobject):  # マウスから呼び出した時の反応が遅い。このメソッドでエラーがでるとショートカットキーでの操作が必要。
		selection = eventobject.Source.getSelection()
		if hasattr(selection, "getRangeAddress"):  # 選択範囲がセル範囲とは限らないのでgetRangeAddress()メソッドがあるか確認する。
			selectionrangeaddress = selection.getRangeAddress()
			if selectionrangeaddress==self.selectionrangeaddress:  # キャッシュのセル範囲アドレスと一致する時。Structで比較。セル範囲では比較できない。
				return  # 何もしない。
			else:  # キャッシュのセル範囲と一致しない時。
				self.selectionrangeaddress = selectionrangeaddress  # キャッシュを更新。
		invokeModuleMethod(eventobject.Source.getActiveSheet().getName(), "selectionChanged", eventobject, self.xscriptcontext)	
	def disposing(self, eventobject):
		eventobject.Source.removeSelectionChangeListener(self)		
class ChangesListener(unohelper.Base, XChangesListener):
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def changesOccurred(self, changesevent):  # Sourceにはドキュメントが入る。
		invokeModuleMethod(changesevent.Source.getCurrentController().getActiveSheet().getName(), "changesOccurred", changesevent, self.xscriptcontext)							
	def disposing(self, eventobject):
		eventobject.Source.removeChangesListener(self)			
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, xscriptcontext):
		self.xscriptcontext = xscriptcontext
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 右クリックで呼ばれる関数。contextmenuexecuteevent.ActionTriggerContainerを操作しないとコンテクストメニューが表示されない。:
		contextmenuinterceptoraction = invokeModuleMethod(contextmenuexecuteevent.Selection.getActiveSheet().getName(), "notifyContextMenuExecute", contextmenuexecuteevent, self.xscriptcontext)  # 正しく実行されれば、enumのcom.sun.star.ui.ContextMenuInterceptorActionのいずれかが返るはず。	
		if contextmenuinterceptoraction is not None:
			return contextmenuinterceptoraction
		return IGNORED  # コンテクストメニューのカスタマイズをしない。
