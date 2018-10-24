#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# サブサブモジュール以遠のエラーは検出できない。
# ノンモダルダイアログ。UnoControlDialogとかメッセージボックスだとマウスをクリックした状態のままになってしまうことがある。
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import os, platform, subprocess, traceback, unohelper
from . import dialogcommons
from com.sun.star.awt import XActionListener, XMouseListener
from com.sun.star.awt import MessageBoxButtons, MessageBoxResults, PosSize, SystemPointer  # 定数
from com.sun.star.awt.MessageBoxType import ERRORBOX, QUERYBOX  # enum
from com.sun.star.beans import NamedValue  # Struct
from com.sun.star.util import URL  # Struct
from com.sun.star.util import XCloseListener
from com.sun.star.style.VerticalAlignment import MIDDLE  # enum
def createDialog(xscriptcontext):  # 選択範囲を削除して、フレームを変更するとマウスボタン押してドラッグしている状態になったままになっている。
	traceback.print_exc()  # PyDevのコンソールにトレースバックを表示。stderrToServer=Trueが必須。
	#  ダイアログに表示する。raiseだとPythonの構文エラーはエラーダイアログがでてこないので。
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。	
	doc = xscriptcontext.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。   	
	docframe = doc.getCurrentController().getFrame()  # モデル→コントローラ→フレーム、でドキュメントのフレームを取得。
	containerwindow = docframe.getContainerWindow()  # ドキュメントのウィンドウ(コンテナウィンドウ=ピア)を取得。
	maTopx = dialogcommons.createConverters(containerwindow)  # ma単位をピクセルに変換する関数を取得。	
	txt = traceback.format_exc()
	lines = txt.split("\n")  # トレースバックを改行で分割。
	h = 20  # FixedTextコントロールの高さ。ma単位。2行分。	
	dialogwidth = 380  # ウィンドウの幅。ma単位。
	controlcontainerprops = {"PositionX": 20, "PositionY": 120, "Width": dialogwidth, "Height": 10, "BackgroundColor": 0xF0F0F0}  # Heightは後で設定し直す。PositionXとPositionYはTaskCreatorに渡したら0にする。
	controlcontainer, addControl = dialogcommons.controlcontainerMaCreator(ctx, smgr, maTopx, controlcontainerprops)  # コントロールコンテナの作成。		
	mouselistener = MouseListener(xscriptcontext)
	controlheight = 0  # コントロールの高さ。ma単位。
	subjs = []  # マウスリスナーのサブジェクトのリスト。
	for i in lines[1:]:  # 2行目以降イテレート。
		if i:  # 空行は除外。
			fixedtextprops = [{"PositionX": 0, "PositionY": controlheight, "Width": dialogwidth, "Height": h, "Label": i, "MultiLine": True, "NoLabel": True, "VerticalAlign": MIDDLE}]
			if i.lstrip().startswith("File "):  # File から始まる行の時。	
				fixedtextprops[0]["TextColor"] = 0x0000FF  # 文字色をblue3にする。
				fixedtextprops.append({"addMouseListener": mouselistener})
				fixedtextcontrol = addControl("FixedText", *fixedtextprops)
				subjs.append(fixedtextcontrol)  # マウスリスナーをつけたコントロールに追加する。
			elif not i.startswith(" "):  # スペース以外から始まる時。
				fixedtextprops[0]["TextColor"] = 0xFF0000  # 文字色をred3にする。
				fixedtextcontrol = addControl("FixedText", *fixedtextprops)
			else:
				fixedtextcontrol = addControl("FixedText", *fixedtextprops)
			controlheight += h
	buttonprops1 = {"PositionX": 10, "PositionY": controlheight, "Width": 60, "Height": 14, "Label": "to ClipBoard"}  # ボタンのプロパティ。PushButtonTypeの値はEnumではエラーになる。VerticalAlignではtextboxと高さが揃わない。
	actionlistener = ActionListener(xscriptcontext, txt)  # ボタンコントロールにつけるリスナー。		
	button1 = addControl("Button", buttonprops1, {"addActionListener": actionlistener, "setActionCommand": "copy"})  
	controlrectangle = button1.getPosSize()  # コントロール間の間隔を幅はX、高さはYから取得。最後に追加したコントロールから取得。
	controlcontainer.setPosSize(0, 0, 0, controlrectangle.Y+controlrectangle.Height+10, PosSize.HEIGHT)  # 最後の行からダイアログの高さを再設定。
	rectangle = controlcontainer.getPosSize()  # コントロールコンテナのRectangle Structを取得。px単位。
	controlcontainer.setPosSize(0, 0, 0, 0, PosSize.POS)  # コントロールコンテナの位置をTaskCreatorのウィンドウの原点にする。
	taskcreator = smgr.createInstanceWithContext('com.sun.star.frame.TaskCreator', ctx)
	args = NamedValue("PosSize", rectangle), NamedValue("FrameName", "exceptiondialog")  # , NamedValue("MakeVisible", True)  # TaskCreatorで作成するフレームのコンテナウィンドウのプロパティ。
	dialogframe = taskcreator.createInstanceWithArguments(args)  # コンテナウィンドウ付きの新しいフレームの取得。サイズ変更は想定しない。
	dialogframe.setTitle(lines[0])  # フレームのタイトルを設定。
	docframe.getFrames().append(dialogframe) # 新しく作ったフレームを既存のフレームの階層に追加する。
	dialogwindow = dialogframe.getContainerWindow()  # ダイアログのコンテナウィンドウを取得。
	toolkit = dialogwindow.getToolkit()  # ピアからツールキットを取得。 	
	controlcontainer.createPeer(toolkit, dialogwindow) # ウィンドウにコントロールコンテナを描画。
	controlcontainer.setVisible(True)  # コントロールの表示。
	dialogwindow.setVisible(True) # ウィンドウの表示。これ以降WindowListenerが発火する。
	args = mouselistener, actionlistener, button1, subjs
	dialogframe.addCloseListener(CloseListener(args))  # CloseListener。ノンモダルダイアログのリスナー削除用。		
class CloseListener(unohelper.Base, XCloseListener):  # ノンモダルダイアログのリスナー削除用。
	def __init__(self, args):
		self.args = args
	def queryClosing(self, eventobject, getsownership):  # ノンモダルダイアログを閉じる時に発火。
		mouselistener, actionlistener, button1, subjs = self.args
		for i in subjs:
			i.removeMouseListener(mouselistener)
		button1.removeActionListener(actionlistener)
		eventobject.Source.removeCloseListener(self)
	def notifyClosing(self, eventobject):
		pass
	def disposing(self, eventobject):  
		pass
class MouseListener(unohelper.Base, XMouseListener):  # Editコントロールではうまく動かない。
	def __init__(self, xscriptcontext):
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。			
		self.pointer = smgr.createInstanceWithContext("com.sun.star.awt.Pointer", ctx)  # ポインタのインスタンスを取得。
		self.args = ctx, smgr, xscriptcontext.getDocument()
	def mousePressed(self, mouseevent):
		ctx, smgr, doc = self.args
		txt = mouseevent.Source.getText()
		fileurl = txt.split('"')[1]  # エラー箇所のfileurlを取得。
		lineno = txt.split(',')[1].split(" ")[2]  # エラー箇所の行番号を取得。		
		# Geayでエラー箇所を開く。
		if all([fileurl, lineno]):  # ファイル名と行番号が取得出来ている時。
			flg = (platform.system()=="Windows")  # Windowsかのフラグ。
			if flg:  # Windowsの時
				geanypath = "C:\\Program Files (x86)\\Geany\\bin\\geany.exe"  # 64bitでのパス。パス区切りは\\にしないとエスケープ文字に反応してしまう。
				if not os.path.exists(geanypath):  # binフォルダはなぜかos.path.exists()は常にFalseになるので使えない。
					geanypath = "C:\\Program Files\\Geany\\bin\\geany.exe"  # 32bitでのパス。
					if not os.path.exists(geanypath):
						geanypath = ""
			else:  # Linuxの時。
				p = subprocess.run(["which", "geany"], universal_newlines=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)  # which geanyの結果をuniversal_newlines=Trueで文字列で取得。
				geanypath = p.stdout.strip()  # /usr/bin/geany が返る。
			componentwindow = doc.getCurrentController().ComponentWindow
			toolkit = componentwindow.getToolkit()			
			if geanypath:  # geanyがインストールされている時。
				msg = "Geanyでソースのエラー箇所を一時ファイルで表示しますか?"
				msgbox = toolkit.createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "myRs", msg)
				if msgbox.execute()==MessageBoxResults.YES:			
					simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)					
					tempfile = smgr.createInstanceWithContext("com.sun.star.io.TempFile", ctx)  # 一時ファイルを取得。一時フォルダを利用するため。
					urltransformer = smgr.createInstanceWithContext("com.sun.star.util.URLTransformer", ctx)
					dummy, tempfileURL = urltransformer.parseStrict(URL(Complete=tempfile.Uri))
					dummy, fileURL = urltransformer.parseStrict(URL(Complete=fileurl))
					destfileurl = "".join([tempfileURL.Protocol, tempfileURL.Path, fileURL.Name])
					simplefileaccess.copy(fileurl, destfileurl)  # マクロファイルを一時フォルダにコピー。
					filepath =  unohelper.fileUrlToSystemPath(destfileurl)  # 一時フォルダのシステムパスを取得。
					if flg:  # Windowsの時。Windowsではなぜか一時ファイルが残る。削除してもLibreOffice6.0.5を終了すると復活して残る。C:\Users\pq\AppData\Local\Temp\
						os.system('start "" "{}" "{}:{}"'.format(geanypath, filepath, lineno))  # バックグランドでGeanyでカーソルの行番号を指定して開く。第一引数の""はウィンドウタイトル。
					else:
						os.system("{} {}:{} &".format(geanypath, filepath, lineno))  # バックグランドでGeanyでカーソルの行番号を指定して開く。
			else:
				msg = "Geanyがインストールされていません。"
				msgbox = toolkit.createMessageBox(componentwindow, ERRORBOX, MessageBoxButtons.BUTTONS_OK, "myRs", msg)
				msgbox.execute()
	def mouseReleased(self, mouseevent):
		pass
	def mouseEntered(self, mouseevent):
		self.pointer.setType(SystemPointer.REFHAND)  # マウスポインタの種類を設定。
		mouseevent.Source.getPeer().setPointer(self.pointer)  # マウスポインタを変更。コントロールからマウスがでるとポインタは元に戻る。
	def mouseExited(self, mouseevent):
		pass
	def disposing(self, eventobject):
		eventobject.Source.removeMouseListener(self)	
class ActionListener(unohelper.Base, XActionListener):
	def __init__(self, *args):
		self.args = args
	def actionPerformed(self, actionevent):
		cmd = actionevent.ActionCommand
		if cmd=="copy":  
			xscriptcontext, txt = self.args
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。			
			systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
			systemclipboard.setContents(dialogcommons.TextTransferable(txt), None)  # クリップボードにコピーする。
	def disposing(self, eventobject):
		pass
