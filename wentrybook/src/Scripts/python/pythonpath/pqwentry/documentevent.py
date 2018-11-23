#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# ドキュメントイベントについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from . import journal
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.sheet.CellInsertMode import ROWS as insert_rows  # enum
MODIFYLISTENERS = []  # ModifyListenerのサブジェクトとリスナーのタプルのリスト。
def documentOnLoad(xscriptcontext):  # ドキュメントを開いた時。リスナー追加後。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。
	for i in namedranges.getElementNames():  # namedrangesをイテレートするとfor文中でnamedrangesを操作してはいけない。
		if not namedranges[i].getReferredCells():
			namedranges.removeByName(i)  # 参照範囲がエラーの名前を削除する。	
	sheets = doc.getSheets()
	journalvars = journal.VARS  # 振替伝票シート固有値。
	beginningdayrow, enddayrow = journalvars.settlingdayrows  # 期首日セルと期末日セルの行インデックスを取得。
	splittedrow = journalvars.splittedrow  # 固定行インデックス。
	daycolumn = journalvars.daycolumn  # 取引日列インデックス。
	tekiyocolumn = daycolumn + 1  # 提要列インデックス。
	slipnocolumn = daycolumn - 1  # 伝票番号列インデックス。
	splittedcolumn = journalvars.splittedcolumn  # 固定列インデックス。
	settlingdayrangeaddresses = []  # 全振替伝票シートの決算日のセル範囲アドレスを取得するリスト。
	slipnorangeaddresses = []  # 全振替伝票シートの伝票番号列と取引日列のセル範囲アドレスを取得するリスト。
	valuerangeaddresses = []  # 全振替伝票シートの金額セルのセル範囲アドレスを取得するリスト。
	sheetnames = []  # 全振替伝票シート名を取得するリスト。
	commetcellstrings = "資産の部", "事業主貸", "負債・資本の部", "事業主借", "元入金", "経費", "専従者給与", "貸倒引当金繰入", "期首商品棚卸高", "仕入金額", "収益", "売上金額", "貸倒引当金戻入", "期末商品棚卸高"  # ハードコーディングしているので変更してはいけないセルの文字列。
	regexpattern = "|".join(commetcellstrings)
	carryovers = []  # 繰越行を挿入するシートとその期主日のタプルのリスト。
	for i in sheets:
		sheetname = i.getName()
		if sheetname.startswith("振替伝票"):  # 振替伝票、から始まるシート名の時。
			journalvars.setSheet(i)
			headerrange = i[:splittedrow, splittedcolumn:journalvars.emptycolumn]
			headerrange.clearContents(CellFlags.ANNOTATION)
			searchdescriptor = i.createSearchDescriptor()
			searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
			searchdescriptor.setSearchString(regexpattern)  # 戻り値はない。				
			cellranges = headerrange.findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
			if cellranges:
				annotations = i.getAnnotations()  # コメントコレクションを取得。
				for j in cellranges.getCells():
					annotations.insertNew(j.getCellAddress(), "名前変更不可")  # コメントを挿入。
					j.getAnnotation().getAnnotationShape().setPropertyValue("Visible", False)  # これをしないとmousePressed()のTargetにAnnotationShapeが入ってしまう。				
			if "繰越" not in i[splittedrow, tekiyocolumn].getString():  # 最初の伝票行の摘要に「繰越」がない時。
				i.insertCells(i[splittedrow, :].getRangeAddress(), insert_rows)  # 最初の伝票行に空行を挿入。
				carryovers.append((i, i[journalvars.settlingdayrows[0], daycolumn].getValue()))  # 繰越行を挿入したシートとその期首日の日付シリアル値のタプルを取得。
			sheetnames.append(sheetname)  # シート名を取得。
			settlingdayrangeaddresses.append(i[beginningdayrow, daycolumn].getRangeAddress())  # 期首日のセル範囲アドレスを取得。
			settlingdayrangeaddresses.append(i[enddayrow, daycolumn].getRangeAddress())  # 期末日のセル範囲アドレスを取得。
			slipnorangeaddresses.append(i[splittedrow:, slipnocolumn:tekiyocolumn].getRangeAddress())  # 伝票番号列と取引日列のセル範囲アドレスを取得。固定行に行挿入でも反応できるように固定行の上行から付ける。
			valuerangeaddresses.append(i[splittedrow:, splittedcolumn:].getRangeAddress())  # 固定列右のセル範囲アドレスを取得。固定行に行挿入でも反応できるように固定行の上行から付ける。
	addModifyListener(doc, settlingdayrangeaddresses, journal.SettlingDayModifyListener(xscriptcontext))  # 決算日の変更を検知するリスナー。
	addModifyListener(doc, slipnorangeaddresses, journal.SlipNoModifyListener(xscriptcontext))  # 伝票番号と取引日の変更を検知するリスナー。	
	addModifyListener(doc, valuerangeaddresses, journal.ValueModifyListener(xscriptcontext))  # 伝票の金額の変更を検知するリスナー。	
	[i[0][splittedrow, daycolumn:tekiyocolumn+1].setDataArray(((i[1], "前記より繰越"),)) for i in carryovers]  # ModifyListenerを追加してから繰越伝票に代入する。
	sheet = sheets[sorted(sheetnames)[-1]]  # 最新年度の振替伝票シートを取得。			
	doc.getCurrentController().setActiveSheet(sheet)
	journal.initSheet(sheet, xscriptcontext)
def addModifyListener(doc, rangeaddresses, modifylistener):	
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
	cellranges.addRangeAddresses(rangeaddresses, False)
	cellranges.addModifyListener(modifylistener)
	MODIFYLISTENERS.append((cellranges, modifylistener))	
def documentUnLoad(xscriptcontext):  # ドキュメントを閉じた時。リスナー削除後。
	for subject, modifylistener in MODIFYLISTENERS:
		subject.removeModifyListener(modifylistener)
