#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import platform
from . import journal
# ドキュメントイベントについて。
MODIFYLISTENERS = []
def documentOnLoad(xscriptcontext):  # ドキュメントを開いた時。リスナー追加後。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	sheets = doc.getSheets()
	charheight = 12  # フォントの大きさ。
	if platform.system()=="Windows":  # Windowsの時
		setSheetProps = lambda x: x.setPropertyValues(("CharFontName", "CharFontNameAsian", "CharHeight"), ("ＭＳ Ｐゴシック", "ＭＳ Ｐゴシック", charheight))
	else:
		setSheetProps = lambda x: x.setPropertyValue("CharHeight", charheight)
	journalvars = journal.VARS
	splittedrow = journalvars.splittedrow
	slipnocolumn = journalvars.daycolumn - 1
	splittedcolumn = journalvars.splittedcolumn	
	settrlingdaycelladdress = journalvars.settrlingdaycelladdress
	settlingdayrangeaddresses = []  # 各シートの決算日のセル範囲アドレスを取得するリスト。
	slipnorangeaddresses = []
	valuerangeaddresses = []
	sheetnames = []
	for i in sheets:
		sheetname = i.getName()
		if sheetname.startswith("振替伝票"):
			sheetnames.append(sheetname)
			setSheetProps(i)
			settlingdayrangeaddresses.append(i[settrlingdaycelladdress].getRangeAddress())
			slipnorangeaddresses.append(i[splittedrow:, slipnocolumn].getRangeAddress())
			valuerangeaddresses.append(i[splittedrow:, splittedcolumn:].getRangeAddress())
	global MODIFYLISTENERS			
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
	cellranges.addRangeAddresses(settlingdayrangeaddresses, False)	
	settlingdaymodifylistener = journal.SettlingDayModifyListener(xscriptcontext)
	cellranges.addModifyListener(settlingdaymodifylistener)
	MODIFYLISTENERS.append((cellranges, settlingdaymodifylistener))	
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
	cellranges.addRangeAddresses(slipnorangeaddresses, False)
	slipnomodifylistener = journal.SlipNoModifyListener(xscriptcontext)
	cellranges.addModifyListener(slipnomodifylistener)
	MODIFYLISTENERS.append((cellranges, slipnomodifylistener))
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
	cellranges.addRangeAddresses(valuerangeaddresses, False)
	valuemodifylistener = journal.ValueModifyListener(xscriptcontext)  # 伝票の金額につけるリスナー。	
	cellranges.addModifyListener(valuemodifylistener)
	MODIFYLISTENERS.append((cellranges, valuemodifylistener))
	sheet = sheets[sorted(sheetnames)[-1]]  # 最新年度の振替伝票シートを取得。			
	doc.getCurrentController().setActiveSheet(sheet)
	journal.initSheet(sheet, xscriptcontext)
def documentUnLoad(xscriptcontext):  # ドキュメントを閉じた時。リスナー削除後。
	for subject, modifylistener in MODIFYLISTENERS:
		subject.removeModifyListener(modifylistener)
