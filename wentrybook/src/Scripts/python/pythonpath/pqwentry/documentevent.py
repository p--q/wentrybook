#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import platform
from . import journal
# ドキュメントイベントについて。
def documentOnLoad(xscriptcontext):  # ドキュメントを開いた時。リスナー追加後。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	sheets = doc.getSheets()
	charheight = 12  # フォントの大きさ。
	if platform.system()=="Windows":  # Windowsの時
		[i.setPropertyValues(("CharFontName", "CharFontNameAsian", "CharHeight"), ("ＭＳ Ｐゴシック", "ＭＳ Ｐゴシック", charheight)) for i in sheets]
	else:
		[i.setPropertyValue("CharHeight", charheight) for i in sheets]
	sheetname = next(i for i in sorted(sheets.getElementNames(), reverse=True) if i.startswith("振替伝票"))  # 最新年度の振替伝票シート名を取得。
	sheet = sheets[sheetname]			
	doc.getCurrentController().setActiveSheet(sheet)
	journal.initSheet(sheet, xscriptcontext)
def documentUnLoad(xscriptcontext):  # ドキュメントを閉じた時。リスナー削除後。
	pass
