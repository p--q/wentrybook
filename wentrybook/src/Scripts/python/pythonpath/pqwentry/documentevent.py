#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import platform
from . import journal
# ドキュメントイベントについて。
def documentOnLoad(xscriptcontext):  # ドキュメントを開いた時。リスナー追加後。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	sheets = doc.getSheets()
	if platform.system()=="Windows":  # Windowsの時はすべてのシートのフォントをMS Pゴシックにする。
		[i.setPropertyValues(("CharFontName", "CharFontNameAsian"), ("ＭＳ Ｐゴシック", "ＭＳ Ｐゴシック")) for i in sheets]
# 	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。
# 	for i in namedranges.getElementNames():  # namedrangesをイテレートするとfor文中でnamedrangesを操作してはいけない。
# 		if not namedranges[i].getReferredCells():
# 			namedranges.removeByName(i)  # 参照範囲がエラーの名前を削除する。	
	sheet = sheets["振替伝票"]		
	doc.getCurrentController().setActiveSheet(sheet)  # 仕訳日誌シートをアクティブにする。	
	journal.initSheet(sheet, xscriptcontext)
def documentUnLoad(xscriptcontext):  # ドキュメントを閉じた時。リスナー削除後。
	pass
