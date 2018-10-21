#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import glob
import os, sys
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.document import MacroExecMode  # 定数
def main():  
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	os.chdir("..")  # 一つ上のディレクトリに移動。
	simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)  # SimpleFileAccess
	source_path = os.path.join(os.getcwd(), "src", "Scripts", "python")  # コピー元フォルダのパスを取得。	
	source_fileurl = unohelper.systemPathToFileUrl(source_path)  # fileurlに変換。	
	if not simplefileaccess.exists(source_fileurl):  # ソースにするフォルダがないときは終了する。
		print("fileurl: {}\nThe source macro folder does not exist.".format(source_fileurl))	
		return	
	ods = glob.glob("*.ods")[0]  # odsファイルを取得。最初の一つのみ取得。
	systempath = os.path.join(os.getcwd(), ods)  # odsファイルのフルパス。
	doc_fileurl = unohelper.systemPathToFileUrl(systempath)  # fileurlに変換。
	desktop = ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # デスクトップの取得。
	flg = isComponentLoaded(desktop, doc_fileurl)  # ドキュメントが開いていたら保存して閉じる。
	python_pkgurl = getVndSunStarPkgUrl(ctx, smgr, doc_fileurl)  # pkgurlの取得。
	if simplefileaccess.exists(python_pkgurl):  # 埋め込みマクロフォルダがすでに存在する時。simplefileaccess.kill(pkgurl)では削除できない。
		package = smgr.createInstanceWithArgumentsAndContext("com.sun.star.packages.Package", (doc_fileurl,), ctx)  # Package。第2引数はinitialize()メソッドで後でも渡せる。
		docroot = package.getByHierarchicalName("/")  # /Scripts/pythonは不可。
		for name in docroot["Scripts"]["python"].getElementNames(): # すでに存在する埋め込みマクロフォルダの各要素を削除。
			del docroot["Scripts"]["python"][name]
		package.commitChanges()  # ファイルにパッケージの変更を書き込む。manifest.xmlも編集される。	
	else:  # 埋め込みマクロフォルダが存在しない時。
		propertyvalues = PropertyValue(Name="Hidden",Value=True),
		doc = desktop.loadComponentFromURL(doc_fileurl, "_blank", 0, propertyvalues)  # ドキュメントをバックグラウンドで開く。
		if doc is None:  # ドキュメントが壊れているときなどはNoneになる。
			print("{} may be corrupted.".format(ods), file=sys.stderr)
			sys.exit()
		createEmbeddedMacroFolder(ctx, smgr, simplefileaccess, doc)  # 埋め込みマクロフォルダを新規作成。開いているドキュメントにしか作成できない。
		doc.store()  # ドキュメントを保存する。
		doc.close(True)  # ドキュメントを閉じる。
	simplefileaccess.copy(source_fileurl, python_pkgurl)  # 埋め込みマクロフォルダにコピーする。開いているドキュメントでは書き込みが反映されない時があるので閉じたドキュメントにする。
	print("Replaced the embedded macro folder in {} with {}.".format(ods, source_path))
	prop = PropertyValue(Name="Hidden",Value=True)
	desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (prop,))  # バックグラウンドでWriterのドキュメントを開く。そうでないとsoffice.binが終了しないときがある。
	terminated = desktop.terminate()  # LibreOfficeを終了しないとリスナーの変更が反映されない。
	if terminated:
		print("\nThe LibreOffice has been terminated.")  # 未保存のドキュメントがなくうまく終了出来た時。
	else:
		print("\nThe LibreOffice is still running. Someone else prevents termination.\nListener changes will not be reflected unless LibreOffice has been terminated.")  # 未保存のドキュメントがあってキャンセルボタンが押された時。
	sys.exit()  # これがないとsoffice.binが終わらないときもある。なぜか1回目の起動後の終了は2分程かかる。とりあえず1回は待たないと次も待たされる。soffice.binというものが1回目終了時に時間がかかる。1回LibreOfficeを起動して終了するとすぐ終わる?
def getVndSunStarPkgUrl(ctx, smgr, doc_fileurl):  # pkgurlの取得。
	urireferencefactory = smgr.createInstanceWithContext("com.sun.star.uri.UriReferenceFactory", ctx)  # UriReferenceFactory
	urireference = urireferencefactory.parse(doc_fileurl)  # ドキュメントのUriReferenceを取得。
	vndsunstarpkgurlreferencefactory = smgr.createInstanceWithContext("com.sun.star.uri.VndSunStarPkgUrlReferenceFactory", ctx)  # VndSunStarPkgUrlReferenceFactory
	vndsunstarpkgurlreference = vndsunstarpkgurlreferencefactory.createVndSunStarPkgUrlReference(urireference)  # ドキュメントのvnd.sun.star.pkgプロトコールにUriReferenceを変換。
	pkgurl = vndsunstarpkgurlreference.getUriReference()  # UriReferenceから文字列のURIを取得。
	return "/".join((pkgurl, "Scripts/python"))  # 開いていないドキュメントの埋め込みマクロフォルダへのフルパスを取得。	
def isComponentLoaded(desktop, doc_fileurl):  # ドキュメントが開いていたら保存して閉じる。
	components = desktop.getComponents()  # ロードしているコンポーネントコレクションを取得。
	for component in components:  # 各コンポーネントについて。
		if hasattr(component, "getURL"):  # スタートモジュールではgetURL()はないためチェックする。
			if component.getURL()==doc_fileurl:  # fileurlが一致するとき、ドキュメントが開いているということ。
				component.store()  # ドキュメントを保存する。
				component.close(True)  # ドキュメントを閉じる。
				return True
	else:
		return False
def createEmbeddedMacroFolder(ctx, smgr, simplefileaccess, component):  # 埋め込みマクロフォルダを作成。	
	transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
	transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(component)
	tdocurl = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # ex. vnd.sun.star.tdoc:/1
	python_tdocurl = "/".join((tdocurl, "Scripts/python"))  # 開いているドキュメントの埋め込みマクロフォルダへのフルパスを取得。	
	simplefileaccess.createFolder(python_tdocurl)  # 埋め込みマクロフォルダを作成。
if __name__ == "__main__":  # オートメーションで実行するとき
	def automation():  # オートメーションのためにglobalに出すのはこの関数のみにする。
		import officehelper
		from functools import wraps
		import sys
		from com.sun.star.beans import PropertyValue  # Struct
		from com.sun.star.script.provider import XScriptContext  
		def connectOffice(func):  # funcの前後でOffice接続の処理
			@wraps(func)
			def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
				try:
					ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
				except:
					print("Could not establish a connection with a running office.", file=sys.stderr)
					sys.exit()
				print("Connected to a running office ...")
				smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
				print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
				return func(ctx, smgr)  # 引数の関数の実行。
			def _getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
				cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
				node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
				ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
				return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
			return wrapper
		@connectOffice  # createXSCRIPTCONTEXTの引数にctxとsmgrを渡すデコレータ。
		def createXSCRIPTCONTEXT(ctx, smgr):  # XSCRIPTCONTEXTを生成。
			class ScriptContext(unohelper.Base, XScriptContext):
				def __init__(self, ctx):
					self.ctx = ctx
				def getComponentContext(self):
					return self.ctx
				def getDesktop(self):
					return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
				def getDocument(self):
					return self.getDesktop().getCurrentComponent()
			return ScriptContext(ctx)  
		return createXSCRIPTCONTEXT()  # XSCRIPTCONTEXTの取得。
	XSCRIPTCONTEXT = automation()  # XSCRIPTCONTEXTを取得。	
	main()  