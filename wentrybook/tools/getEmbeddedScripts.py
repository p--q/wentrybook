#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import glob
import os, sys
def main():  
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)  # SimpleFileAccess
	os.chdir("..")  # 一つ上のディレクトリに移動。
	ods = glob.glob("*.ods")[0]  # odsファイルを取得。最初の一つのみ取得。
	systempath = os.path.join(os.getcwd(), ods)  # odsファイルのフルパス。
	doc_fileurl = unohelper.systemPathToFileUrl(systempath)  # fileurlに変換。
	desktop = ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # デスクトップの取得。
	components = desktop.getComponents()  # ロードしているコンポーネントコレクションを取得。
	transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)  # TransientDocumentsDocumentContentFactory
	for component in components:  # 各コンポーネントについて。
		if hasattr(component, "getURL"):  # スタートモジュールではgetURL()はないためチェックする。
			if component.getURL()==doc_fileurl:  # fileurlが一致するとき、ドキュメントが開いているということ。
				transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(component)
				pkgurl = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # ex. vnd.sun.star.tdoc:/1
				python_fileurl = "/".join((pkgurl, "Scripts/python"))  # ドキュメント内フォルダへのフルパスを取得。
				if simplefileaccess.exists(python_fileurl):  # 埋め込みマクロフォルダが存在する時。
					dest_dir = createDest(simplefileaccess)  # 出力先フォルダのfileurlを取得。
					simplefileaccess.copy(python_fileurl, dest_dir)  # 埋め込みマクロフォルダを出力先フォルダにコピーする。
					print("The embedded Macro folder in '{}' has been exported to '{}'.".format(python_fileurl, dest_dir))
					return	# 関数を出る。		
	else:  # ドキュメントを開いていない時。
		package = smgr. createInstanceWithArgumentsAndContext("com.sun.star.packages.Package", (doc_fileurl,), ctx)  # Package。第2引数はinitialize()メソッドで後でも渡せる。
		docroot = package.getByHierarchicalName("/")  # /Scripts/pythonは不可。
		if ("Scripts" in docroot and "python" in docroot["Scripts"]):  # 埋め込みマクロフォルダが存在する時。
			dest_dir = createDest(simplefileaccess)  # 出力先フォルダのfileurlを取得。
			getContents(simplefileaccess, docroot["Scripts"]["python"], dest_dir) 
			print("The embedded Macro folder in '{}' has been exported to '{}'.".format(ods, dest_dir))
			return	# 関数を出る。	
	print("The embedded macro folder does not exist in {}.".format(ods))  # 埋め込みマクロフォルダが存在しない時。
def getContents(simplefileaccess, folder, pwd):
	for sub in folder:  # 子要素のオブジェクトについて。
		name = sub.getName()  # 子要素のオブジェクトの名前を取得。
		fileurl = "/".join((pwd, name))  # 出力先のfileurlを取得。
		if sub.supportsService("com.sun.star.packages.PackageFolder"):  # PackageFolderの時はフォルダとして出力。
			if not simplefileaccess.exists(fileurl):
				simplefileaccess.createFolder(fileurl)
			getContents(simplefileaccess, sub, fileurl)  # 子要素のオブジェクトについて同様に出力。
		elif sub.supportsService("com.sun.star.packages.PackageStream"):  # PackageStreamのときはファイルとして出力。
			simplefileaccess.writeFile(fileurl, sub.getInputStream())  # ファイルが存在しなければ新規作成してくれる。			
def createDest(simplefileaccess):  # 出力先フォルダのfileurlを取得する。
	src_path = os.path.join(os.getcwd(), "src")  # srcフォルダのパスを取得。
	src_fileurl = unohelper.systemPathToFileUrl(src_path)  # fileurlに変換。
	destdir = "/".join((src_fileurl, "Scripts/python"))
	if simplefileaccess.exists(destdir):  # pythonフォルダがすでにあるとき
		s = input("Delete the existing src/Scripts/python?[y/N]").lower()
		if s=="y":
			simplefileaccess.kill(destdir)  # すでにあるpythonフォルダを削除。	
		else:
			print("Exit")
			sys.exit(0)
	simplefileaccess.createFolder(destdir)  # pythonフォルダを作成。
	return destdir			
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
	