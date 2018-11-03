#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import os, unohelper
from pqwentry import documentevent, journal  # Contextmenuの呼び出しは相対インポートではエラーになる。
from com.sun.star.awt import MessageBoxButtons  # 定数
from com.sun.star.awt.MessageBoxType import ERRORBOX  # enum
from com.sun.star.datatransfer import XTransferable
from com.sun.star.datatransfer import DataFlavor  # Struct
from com.sun.star.datatransfer import UnsupportedFlavorException  # 例外
from com.sun.star.i18n.TransliterationModulesNew import HALFWIDTH_FULLWIDTH  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.sheet.CellDeleteMode import ROWS as delete_rows  # enum
from com.sun.star.sheet.CellInsertMode import ROWS as insert_rows  # enum
from com.sun.star.table import BorderLine2, TableBorder2 # Struct
from com.sun.star.table import BorderLineStyle  # 定数
COLORS = {\
		"magenta3": 0xFF00FF,\
		"black": 0x000000,\
		"skyblue": 0x00CCFF,\
		"silver": 0xC0C0C0,\
		"red3": 0xFF0000,\
		"violet": 0x9999FF}  # 色の16進数。	
HOLIDAYS = {\
		2018:[[1,2,3,8],[11,12],[21],[29,30],[3,4,5],[],[16],[11],[17,23,24],[8],[3,23],[23,24,28,29,30,31]],\
		2019:[[1,2,3,14],[11],[21],[29],[3,4,5,6],[],[15],[11,12],[16,23],[14],[3,4,23],[23,28,29,30,31]],\
		2020:[[1,2,3,13],[11],[20],[29],[3,4,5,6],[],[23,24],[10],[21,22],[],[3,23],[23,28,29,30,31]],\
		2021:[[1,2,3,11],[11],[20],[29],[3,4,5],[],[19],[11],[20,23],[11],[3,23],[23,28,29,30,31]],\
		2022:[[1,2,3,10],[11],[21],[29],[3,4,5],[],[18],[11],[19,23],[10],[3,23],[23,28,29,30,31]],\
		2023:[[1,2,3,9],[11],[21],[29],[3,4,5],[],[17],[11],[18,23],[9],[3,23],[23,28,29,30,31]],\
		2024:[[1,2,3,8],[11,12],[20],[29],[3,4,5,6],[],[15],[11,12],[16,22,23],[14],[3,4,23],[23,28,29,30,31]],\
		2025:[[1,2,3,13],[11],[20],[29],[3,4,5,6],[],[21],[11],[15,23],[13],[3,23,24],[23,28,29,30,31]],\
		2026:[[1,2,3,12],[11],[20],[29],[3,4,5,6],[],[20],[11],[21,22,23],[12],[3,23],[23,28,29,30,31]],\
		2027:[[1,2,3,11],[11],[21,22],[29],[3,4,5],[],[19],[11],[20,23],[11],[3,23],[23,28,29,30,31]],\
		2028:[[1,2,3,10],[11],[20],[29],[3,4,5],[],[17],[11],[18,22],[9],[3,23],[23,28,29,30,31]],\
		2029:[[1,2,3,8],[11,12],[20],[29,30],[3,4,5],[],[16],[11],[17,23,24],[8],[3,23],[23,24,28,29,30,31]],\
		2030:[[1,2,3,14],[11],[20],[29],[3,4,5,6],[],[15],[11,12],[16,23],[14],[3,4,23],[23,28,29,30,31]]}  # 祝日JSON。HOLIDAYS[年][月-1]で祝日の日のタプルが返る。日曜日の祝日も含まれる。
def getModule(sheetname):  # シート名に応じてモジュールを振り分ける関数。
	if sheetname is None:  # シート名でNoneが返ってきた時はドキュメントイベントとする。
		return documentevent
	elif sheetname=="振替伝票":
		return journal
	return None  # モジュールが見つからなかった時はNoneを返す。
class TextTransferable(unohelper.Base, XTransferable):
	def __init__(self, txt):  # クリップボードに渡す文字列を受け取る。
		self.txt = txt
		self.unicode_content_type = "text/plain;charset=utf-16"
	def getTransferData(self, flavor):
		if flavor.MimeType.lower()!=self.unicode_content_type:
			raise UnsupportedFlavorException()
		return self.txt
	def getTransferDataFlavors(self):
		return DataFlavor(MimeType=self.unicode_content_type, HumanPresentableName="Unicode Text"),  # DataTypeの設定方法は不明。
	def isDataFlavorSupported(self, flavor):
		return flavor.MimeType.lower()==self.unicode_content_type
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
def convertKanaFULLWIDTH(transliteration, kanatxt):  # カナ名を半角からスペースを削除して全角にして返す。
	transliteration.loadModuleNew((HALFWIDTH_FULLWIDTH,), Locale(Language = "ja", Country = "JP"))
	kanatxt = kanatxt.replace(" ", "")  # 半角空白を除去してカナ名を取得。
	return transliteration.transliterate(kanatxt, 0, len(kanatxt), [])[0]  # ｶﾅを全角に変換。	
def showErrorMessageBox(controller, msg):
	componentwindow = controller.ComponentWindow
	msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, ERRORBOX, MessageBoxButtons.BUTTONS_OK, "myRs", msg)
	msgbox.execute()
def toNewEntry(sheet, rangeaddress, edgerow, dest_row):  # 使用中最下行へ。新規行挿入は不要。
	startrow, endrowbelow = rangeaddress.StartRow, rangeaddress.EndRow+1  # 選択範囲の開始行と終了行の取得。
	if endrowbelow>edgerow:
		endrowbelow = edgerow
	sourcerangeaddress = sheet[startrow:endrowbelow, :].getRangeAddress()  # コピー元セル範囲アドレスを取得。
	sheet.moveRange(sheet[dest_row, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。	
	sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動したソース行を削除。
def toOtherEntry(sheet, rangeaddress, edgerow, dest_row):  # 新規行挿入が必要な移動。
	startrow, endrowbelow = rangeaddress.StartRow, rangeaddress.EndRow+1  # 選択範囲の開始行と終了行の取得。
	if endrowbelow>edgerow:
		endrowbelow = edgerow
	sourcerange = sheet[startrow:endrowbelow, :]  # 行挿入前にソースのセル範囲を取得しておく。
	dest_rangeaddress = sheet[dest_row:dest_row+(endrowbelow-startrow), :].getRangeAddress()  # 挿入前にセル範囲アドレスを取得しておく。
	sheet.insertCells(dest_rangeaddress, insert_rows)  # 空行を挿入。	
	sheet.queryIntersection(dest_rangeaddress).clearContents(511)  # 挿入した行の内容をすべてを削除。挿入セルは挿入した行の上のプロパティを引き継いでいるのでリセットしないといけない。
	sourcerangeaddress = sourcerange.getRangeAddress()  # コピー元セル範囲アドレスを取得。行挿入後にアドレスを取得しないといけない。
	sheet.moveRange(sheet[dest_row, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。			
	sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動したソース行を削除。	
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
def rowMenuEntries(addMenuentry):  # コンテクストメニュー追加。
	addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsBefore"})
	addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsAfter"})
	addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteRows"}) 
def columnMenuEntries(addMenuentry):  # コンテクストメニュー追加。
	addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertColumnsBefore"})
	addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertColumnsAfter"})
	addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteColumns"}) 	
def getBaseURL(xscriptcontext):	 # 埋め込みマクロのScriptingURLのbaseurlを返す。__file__はvnd.sun.star.tdoc:/4/Scripts/python/filename.pyというように返ってくる。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	modulepath = __file__  # ScriptingURLにするマクロがあるモジュールのパスを取得。ファイルのパスで場合分け。sys.path[0]は__main__の位置が返るので不可。
	ucp = "vnd.sun.star.tdoc:"  # 埋め込みマクロのucp。
	filepath = modulepath.replace(ucp, "")  #  ucpを除去。
	transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
	transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(doc)
	contentidentifierstring = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # __file__の数値部分に該当。
	macrofolder = "{}/Scripts/python".format(contentidentifierstring.replace(ucp, ""))  #埋め込みマクロフォルダへのパス。	
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
	