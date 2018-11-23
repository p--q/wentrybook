#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import os
# from datetime import date, timedelta
from pqwentry import documentevent, journal  # Contextmenuの呼び出しは相対インポートではエラーになる。
from com.sun.star.awt import MessageBoxButtons  # 定数
from com.sun.star.awt.MessageBoxType import ERRORBOX  # enum
from com.sun.star.table import BorderLine2, TableBorder2 # Struct
from com.sun.star.lang import Locale  # Struct
from com.sun.star.table import BorderLineStyle  # 定数
COLORS = {\
		"magenta3": 0xFF00FF,\
		"black": 0x000000,\
		"silver": 0xC0C0C0,\
		"white": 0xFFFFFF,\
		"violet": 0x9999FF}  # 色の16進数。	
def getModule(sheetname):  # シート名に応じてモジュールを振り分ける関数。
	if sheetname is None:  # シート名でNoneが返ってきた時はドキュメントイベントとする。
		return documentevent
	elif sheetname.startswith("振替伝票"):
		return journal
	return None  # モジュールが見つからなかった時はNoneを返す。
def formatkeyCreator(doc):  # ドキュメントを引数にする。
	def createFormatKey(formatstring):  # formatstringの書式はLocalによって異なる。 
		numberformats = doc.getNumberFormats()  # ドキュメントのフォーマット一覧を取得。デフォルトのフォーマット一覧はCalcの書式→セル→数値でみれる。
		locale = Locale(Language="ja", Country="JP")  # フォーマット一覧をくくる言語と国を設定。インストールしていないUIの言語でもよい。。 
		formatkey = numberformats.queryKey(formatstring, locale, True)  # formatstringが既存のフォーマット一覧にあるか調べて取得。第3引数のブーリアンは意味はないはず。 
		if formatkey == -1:  # デフォルトのフォーマットにformatstringがないとき。
			formatkey = numberformats.addNew(formatstring, locale)  # フォーマット一覧に追加する。保存はドキュメントごと。 
		return formatkey
	return createFormatKey
def createBorders():# 枠線の作成。
	noneline = BorderLine2(LineStyle=BorderLineStyle.NONE)
	firstline = BorderLine2(LineStyle=BorderLineStyle.DASHED, LineWidth=45, Color=COLORS["violet"])
	secondline =  BorderLine2(LineStyle=BorderLineStyle.DASHED, LineWidth=45, Color=COLORS["magenta3"])	
	tableborder2 = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)
	topbottomtableborder = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=False, IsRightLineValid=False)
	leftrighttableborder = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=False, IsBottomLineValid=False, IsLeftLineValid=True, IsRightLineValid=True)
	return noneline, tableborder2, topbottomtableborder, leftrighttableborder  # 作成した枠線をまとめたタプル。
def showErrorMessageBox(controller, msg):
	componentwindow = controller.ComponentWindow
	componentwindow.getToolkit().createMessageBox(componentwindow, ERRORBOX, MessageBoxButtons.BUTTONS_OK, "WEntryBook", msg).execute()	
def contextmenuHelper(sheetvars, contextmenuexecuteevent, xscriptcontext):	
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。。
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
	del contextmenu[:]  # contextmenu.clear()は不可。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
	dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)		
	dispatcher.executeDispatch(controller.getFrame(), ".uno:TableDeselectAll", "", 0, ())  # すべてのシートの選択を解除。
	sheetvars.setSheet(controller.getActiveSheet())  # 変数を取得し直す。
	selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
	return contextmenuname, addMenuentry, baseurl, selection
# 	
# 	
# 	
# 	
# 以下コンテクストメニュー
def menuentryCreator(menucontainer):  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	i = 0  # インデックスを初期化する。
	def addMenuentry(menutype, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
		nonlocal i
		menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
		[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
		menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
		i += 1  # インデックスを増やす。
	return addMenuentry
def cutcopypasteMenuEntries(addMenuentry):  # コンテクストメニュー追加。
	addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
	addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
	addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
def getBaseURL(xscriptcontext):	 # 埋め込みマクロのScriptingURLのbaseurlを返す。__file__はvnd.sun.star.tdoc:/4/Scripts/python/filename.pyというように返ってくる。
	modulepath = __file__  # ScriptingURLにするマクロがあるモジュールのパスを取得。ファイルのパスで場合分け。sys.path[0]は__main__の位置が返るので不可。
	ucp = "vnd.sun.star.tdoc:"  # 埋め込みマクロのucp。
	filepath = modulepath.replace(ucp, "")  #  ucpを除去。ドキュメントを一旦閉じて開き直してもContentIdentifierが更新されない。
	filepath = os.path.join(*filepath.split("/")[2:])  # Scripts/python/pythonpath/indoc/commons.py。ContentIdentifierを除く。
	macrofolder = "Scripts/python"
	location = "document"  # マクロの場所。	
	relpath = os.path.relpath(filepath, start=macrofolder)  # マクロフォルダからの相対パスを取得。パス区切りがOS依存で返ってくる。
	return "vnd.sun.star.script:{}${}?language=Python&location={}".format(relpath.replace(os.sep, "|"), "{}", location)  # ScriptingURLのbaseurlを取得。Windowsのためにos.sepでパス区切りを置換。	
def invokeMenuEntry(entrynum):  # コンテクストメニュー項目から呼び出された処理をシートごとに振り分ける。コンテクストメニューから呼び出しているこの関数ではXSCRIPTCONTEXTが使える。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	selection = doc.getCurrentSelection()  # セル(セル範囲)またはセル範囲、セル範囲コレクションが入るはず。
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # セル範囲コレクション以外の時。
		m = getModule(doc.getCurrentController().getActiveSheet().getName())
		if hasattr(m, "contextMenuEntries"):
			getattr(m, "contextMenuEntries")(entrynum, XSCRIPTCONTEXT)	
# ContextMenuInterceptorのnotifyContextMenuExecute()メソッドで設定したメニュー項目から呼び出される関数。関数名変更不可。動的生成も不可。
def entry1():
	invokeMenuEntry(1)
def entry2():
	invokeMenuEntry(2)	
def entry3():
	invokeMenuEntry(3)	
def entry4():
	invokeMenuEntry(4)
def entry5():
	invokeMenuEntry(5)
def entry6():
	invokeMenuEntry(6)
def entry7():
	invokeMenuEntry(7)
def entry8():
	invokeMenuEntry(8)
def entry9():
	invokeMenuEntry(9)	
def entry10():
	invokeMenuEntry(10)	
def entry11():
	invokeMenuEntry(11)	
def entry12():
	invokeMenuEntry(12)	
def entry13():
	invokeMenuEntry(13)	
def entry14():
	invokeMenuEntry(14)	
def entry15():
	invokeMenuEntry(15)	
def entry16():
	invokeMenuEntry(16)	
def entry17():
	invokeMenuEntry(17)	
def entry18():
	invokeMenuEntry(18)	
def entry19():
	invokeMenuEntry(19)	
def entry20():
	invokeMenuEntry(20)	
def entry21():
	invokeMenuEntry(21)	
def entry22():
	invokeMenuEntry(22)	
def entry23():
	invokeMenuEntry(23)	
def entry24():
	invokeMenuEntry(24)	
def entry25():
	invokeMenuEntry(25)	
def entry26():
	invokeMenuEntry(26)	
def entry27():
	invokeMenuEntry(27)	
def entry28():
	invokeMenuEntry(28)	
def entry29():
	invokeMenuEntry(29)	
def entry30():
	invokeMenuEntry(30)	
	