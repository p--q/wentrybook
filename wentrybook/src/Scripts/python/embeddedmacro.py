#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import sys
from types import ModuleType
def macro(documentevent=None):  # 引数は文書のイベント駆動用。  
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)  # SimpleFileAccess
	modulefolderpath = getModuleFolderPath(ctx, smgr, doc)  # 埋め込みpythonpathフォルダのパスを取得。
	tdocimport = load_module(simplefileaccess, "/".join((modulefolderpath, "tdocimport.py")))  # import hooks
	tdocimport.install_meta(simplefileaccess, modulefolderpath)
	from pqdesignr2 import listeners  # ここでインポートしたモジュールの関数だけなぜかXSCRIPTCONTEXTが使えない。デコレーターも不可。
	listeners.addLinsteners(tdocimport, modulefolderpath, XSCRIPTCONTEXT)  # tdocimportとmodulefolderpathは最後にremoveするために渡す。
def load_module(simplefileaccess, modulepath):
	inputstream = simplefileaccess.openFileRead(modulepath)
	dummy, b = inputstream.readBytes([], inputstream.available())  # simplefileaccess.getSize(module_tdocurl)は0が返る。
	source = bytes(b).decode("utf-8")  # モジュールのソースをテキストで取得。
	mod = sys.modules.setdefault(modulepath, ModuleType(modulepath))  # 新規モジュールをsys.modulesに挿入。
	code = compile(source, modulepath, 'exec')  # urlを呼び出し元としてソースコードをコンパイルする。
	mod.__file__ = modulepath  # モジュールの__file__を設定。
	mod.__package__ = ''  # モジュールの__package__を設定。
	exec(code, mod.__dict__)  # モジュールの名前空間を設定する。
	return mod
def getModuleFolderPath(ctx, smgr, doc):
	transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
	transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(doc)
	tdocurl = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # ex. vnd.sun.star.tdoc:/1	
	return "/".join((tdocurl, "Scripts/python/pythonpath"))  # 開いているドキュメント内の埋め込みマクロフォルダへのパス。
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。	
