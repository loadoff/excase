package excase

import (
	"fmt"
	"strconv"
	"strings"
	"time"

	"github.com/loadoff/excl"
)

type exStyles struct {
	style1    *excl.Style
	style1New *excl.Style
	style2    *excl.Style
	style2New *excl.Style
	style3    *excl.Style
	style3New *excl.Style
	style4    *excl.Style
	style4New *excl.Style
	style5    *excl.Style
	style5New *excl.Style
	style6    *excl.Style
	style6New *excl.Style
	style7    *excl.Style
	style8    *excl.Style
	style9    *excl.Style
	style10   *excl.Style
	style11   *excl.Style
	style12   *excl.Style
	style13   *excl.Style
	style14   *excl.Style
	style15   *excl.Style
	style16   *excl.Style
	style17   *excl.Style
}

// ExCase テストケース
type ExCase struct {
	FilePath string
	caseBook *excl.Workbook
	dir      string
	styles   *exStyles
	sections []*ExSection
}

// ExSection セクション情報
type ExSection struct {
	testCount   int
	largeCount  int
	middleCount int
	smallCount  int
	styles      *exStyles
	caseSheet   *excl.Sheet
	name        string
	large       string
	middle      string
	small       string
}

// ExTest テストの出力行
type ExTest struct {
	row *excl.Row
}

// InitExCase Excelテストケースを作成する
func InitExCase() *ExCase {
	var err error
	ex := &ExCase{}
	ex.FilePath = strings.Replace(time.Now().Format("20060102030405"), ".", "", 1) + ".xlsx"

	if ex.caseBook, err = excl.Create(); err != nil {
		fmt.Println(err.Error())
		return nil
	}
	// スタイル作成
	style := &excl.Style{Wrap: 1, Vertical: "top"}
	border := excl.Border{Left: &excl.BorderSetting{Style: "thin"}}
	style.BorderID = ex.caseBook.Styles.SetBorder(border)
	style.Wrap = 1
	ex.styles = &exStyles{}
	ex.styles.style1 = style
	ex.styles.style3 = style
	ex.styles.style5 = style

	style = &excl.Style{Wrap: 1, Vertical: "top"}
	border = excl.Border{Left: &excl.BorderSetting{Style: "hair"}}
	style.BorderID = ex.caseBook.Styles.SetBorder(border)
	style.Wrap = 1
	ex.styles.style2 = style
	ex.styles.style4 = style
	ex.styles.style6 = style

	style = &excl.Style{Wrap: 1, Vertical: "top"}
	border = excl.Border{Left: &excl.BorderSetting{Style: "thin"}, Top: &excl.BorderSetting{Style: "thin"}}
	style.BorderID = ex.caseBook.Styles.SetBorder(border)
	style.Wrap = 1
	ex.styles.style1New = style
	ex.styles.style3New = style
	ex.styles.style5New = style
	ex.styles.style7 = style
	ex.styles.style9 = style
	ex.styles.style10 = style
	ex.styles.style11 = style
	ex.styles.style12 = style
	ex.styles.style13 = style
	ex.styles.style14 = style
	ex.styles.style15 = style
	ex.styles.style16 = style

	style = &excl.Style{Wrap: 1, Vertical: "top"}
	border = excl.Border{Left: &excl.BorderSetting{Style: "thin"}, Top: &excl.BorderSetting{Style: "thin"}, Right: &excl.BorderSetting{Style: "thin"}}
	style.BorderID = ex.caseBook.Styles.SetBorder(border)
	ex.styles.style17 = style

	style = &excl.Style{Wrap: 1, Vertical: "top"}
	border = excl.Border{Left: &excl.BorderSetting{Style: "hair"}, Top: &excl.BorderSetting{Style: "thin"}}
	style.BorderID = ex.caseBook.Styles.SetBorder(border)
	ex.styles.style2New = style
	ex.styles.style4New = style
	ex.styles.style6New = style
	ex.styles.style8 = style
	return ex
}

// OpenSection 新しいシートにテストを出力する
func (ex *ExCase) OpenSection(name string) *ExSection {
	for _, sec := range ex.sections {
		if sec.name == name {
			return sec
		}
	}
	sec := &ExSection{name: name, styles: ex.styles}
	sec.caseSheet, _ = ex.caseBook.OpenSheet(name)
	sec.caseSheet.ShowGridlines(false)
	caseRow := sec.caseSheet.GetRow(4)
	borderSetting := &excl.BorderSetting{Style: "thin"}
	border := excl.Border{Left: borderSetting, Right: borderSetting, Top: borderSetting, Bottom: borderSetting}
	font := excl.Font{Color: "FFFFFF"}
	style := &excl.Style{}
	style.FontID = ex.caseBook.Styles.SetFont(font)
	style.FillID = ex.caseBook.Styles.SetBackgroundColor("361e6d")
	style.BorderID = ex.caseBook.Styles.SetBorder(border)
	caseRow.SetString("No.", 1).SetStyle(style)
	caseRow.SetString("大項目名", 2).SetStyle(style)
	caseRow.SetString("No.", 3).SetStyle(style)
	caseRow.SetString("中項目名", 4).SetStyle(style)
	caseRow.SetString("No.", 5).SetStyle(style)
	caseRow.SetString("小項目名", 6).SetStyle(style)
	caseRow.SetString("No.", 7).SetStyle(style)
	caseRow.SetString("実施内容", 8).SetStyle(style)
	caseRow.SetString("合格条件", 9).SetStyle(style)
	caseRow.SetString("実施日", 10).SetStyle(style)
	caseRow.SetString("実施者", 11).SetStyle(style)
	caseRow.SetString("結果", 12).SetStyle(style)
	caseRow.SetString("補足", 13).SetStyle(style)
	caseRow.SetString("エビデンス", 14).SetStyle(style)
	caseRow.SetString("検証日", 15).SetStyle(style)
	caseRow.SetString("検証者", 16).SetStyle(style)
	caseRow.SetString("結果", 17).SetStyle(style)
	ex.sections = append(ex.sections, sec)

	sec.caseSheet.SetColWidth(4, 1)    // A
	sec.caseSheet.SetColWidth(25, 2)   // B
	sec.caseSheet.SetColWidth(4, 3)    // C
	sec.caseSheet.SetColWidth(25, 4)   // D
	sec.caseSheet.SetColWidth(4, 5)    // E
	sec.caseSheet.SetColWidth(25, 6)   // F
	sec.caseSheet.SetColWidth(4, 7)    // G
	sec.caseSheet.SetColWidth(30, 8)   // H
	sec.caseSheet.SetColWidth(30, 9)   // I
	sec.caseSheet.SetColWidth(6.5, 10) // J
	sec.caseSheet.SetColWidth(6.5, 11) // K
	sec.caseSheet.SetColWidth(6.5, 12) // L
	sec.caseSheet.SetColWidth(25, 13)  // M
	return sec
}

// CloseSection セクションを閉じる
func (ex *ExSection) CloseSection() {
	if ex.caseSheet == nil {
		return
	}
	caseRow := ex.caseSheet.GetRow(ex.testCount + 5)
	border := excl.Border{Top: &excl.BorderSetting{Style: "thin"}}
	style := &excl.Style{}
	style.BorderID = ex.caseSheet.Styles.SetBorder(border)
	caseRow.GetCell(1).SetStyle(style)
	caseRow.GetCell(2).SetStyle(style)
	caseRow.GetCell(3).SetStyle(style)
	caseRow.GetCell(4).SetStyle(style)
	caseRow.GetCell(5).SetStyle(style)
	caseRow.GetCell(6).SetStyle(style)
	caseRow.GetCell(7).SetStyle(style)
	caseRow.GetCell(8).SetStyle(style)
	caseRow.GetCell(9).SetStyle(style)
	caseRow.GetCell(10).SetStyle(style)
	caseRow.GetCell(11).SetStyle(style)
	caseRow.GetCell(12).SetStyle(style)
	caseRow.GetCell(13).SetStyle(style)
	caseRow.GetCell(14).SetStyle(style)
	caseRow.GetCell(15).SetStyle(style)
	caseRow.GetCell(16).SetStyle(style)
	caseRow.GetCell(17).SetStyle(style)
	ex.caseSheet.Close()
	ex.caseSheet = nil
}

// Close 閉じる
func (ex *ExCase) Close() {
	for _, sec := range ex.sections {
		sec.CloseSection()
	}
	ex.caseBook.Save(ex.FilePath)
}

// Large 大項目をセット
func (ex *ExSection) Large(name string) *ExSection {
	ex.largeCount++
	ex.middleCount = 0
	ex.smallCount = 0
	ex.large = name
	return ex
}

// Middle 中項目をセット
func (ex *ExSection) Middle(name string) *ExSection {
	ex.smallCount = 0
	ex.middleCount++
	ex.middle = name
	/*	*/
	return ex
}

// Small 小項目をセット
func (ex *ExSection) Small(name string) *ExSection {
	ex.smallCount++
	ex.small = name
	return ex
}

// Test テストの内容と合格条件をセットする
func (ex *ExSection) Test(content string, pass string) *ExTest {
	ex.testCount++
	test := &ExTest{}
	test.row = ex.caseSheet.GetRow(ex.testCount + 4)

	test.row.GetCell(1).SetStyle(ex.styles.style1)
	test.row.GetCell(2).SetStyle(ex.styles.style2)
	test.row.GetCell(3).SetStyle(ex.styles.style3)
	test.row.GetCell(4).SetStyle(ex.styles.style4)
	test.row.GetCell(5).SetStyle(ex.styles.style5)
	test.row.GetCell(6).SetStyle(ex.styles.style6)
	test.row.GetCell(7).SetStyle(ex.styles.style7)
	test.row.GetCell(8).SetStyle(ex.styles.style8)
	test.row.GetCell(9).SetStyle(ex.styles.style9)
	test.row.GetCell(10).SetStyle(ex.styles.style10)
	test.row.GetCell(11).SetStyle(ex.styles.style11)
	test.row.GetCell(12).SetStyle(ex.styles.style12)
	test.row.GetCell(13).SetStyle(ex.styles.style13)
	test.row.GetCell(14).SetStyle(ex.styles.style14)
	test.row.GetCell(15).SetStyle(ex.styles.style15)
	test.row.GetCell(16).SetStyle(ex.styles.style16)
	test.row.GetCell(17).SetStyle(ex.styles.style17)

	// 大項目処理
	if ex.large != "" {
		test.row.SetNumber(strconv.Itoa(ex.largeCount), 1).SetStyle(ex.styles.style1New)
		test.row.SetString(ex.large, 2).SetStyle(ex.styles.style2New)
		test.row.GetCell(3).SetStyle(ex.styles.style3New)
		test.row.GetCell(4).SetStyle(ex.styles.style4New)
		test.row.GetCell(5).SetStyle(ex.styles.style5New)
		test.row.GetCell(6).SetStyle(ex.styles.style6New)
		ex.large = ""
	}
	// 中項目処理
	if ex.middle != "" {
		test.row.SetNumber(strconv.Itoa(ex.middleCount), 3).SetStyle(ex.styles.style3New)
		test.row.SetString(ex.middle, 4).SetStyle(ex.styles.style4New)
		test.row.GetCell(5).SetStyle(ex.styles.style5New)
		test.row.GetCell(6).SetStyle(ex.styles.style6New)
		ex.middle = ""
	}
	// 小項目処理
	if ex.small != "" {
		test.row.SetNumber(strconv.Itoa(ex.smallCount), 5).SetStyle(ex.styles.style5New)
		test.row.SetString(ex.small, 6).SetStyle(ex.styles.style6New)
		ex.small = ""
	}
	test.row.SetNumber(strconv.Itoa(ex.testCount), 7)
	test.row.SetString(content, 8)
	test.row.SetString(pass, 9)
	return test
}

// Passed 合格をセット
func (test *ExTest) Passed() *ExTest {
	test.row.SetString(time.Now().Format("01/02"), 10)
	test.row.SetString("合格", 12)
	return test
}

// Failed 不合格をセット
func (test *ExTest) Failed() *ExTest {
	test.row.SetString(time.Now().Format("01/02"), 10)
	test.row.SetString("不合格", 12)

	test.row.GetCell(1).SetBackgroundColor("fb0a2a")
	test.row.GetCell(2).SetBackgroundColor("fb0a2a")
	test.row.GetCell(3).SetBackgroundColor("fb0a2a")
	test.row.GetCell(4).SetBackgroundColor("fb0a2a")
	test.row.GetCell(5).SetBackgroundColor("fb0a2a")
	test.row.GetCell(6).SetBackgroundColor("fb0a2a")
	test.row.GetCell(7).SetBackgroundColor("fb0a2a")
	test.row.GetCell(8).SetBackgroundColor("fb0a2a")
	test.row.GetCell(9).SetBackgroundColor("fb0a2a")
	test.row.GetCell(10).SetBackgroundColor("fb0a2a")
	test.row.GetCell(11).SetBackgroundColor("fb0a2a")
	test.row.GetCell(12).SetBackgroundColor("fb0a2a")
	test.row.GetCell(13).SetBackgroundColor("fb0a2a")
	test.row.GetCell(14).SetBackgroundColor("fb0a2a")
	test.row.GetCell(15).SetBackgroundColor("fb0a2a")
	test.row.GetCell(16).SetBackgroundColor("fb0a2a")
	test.row.GetCell(17).SetBackgroundColor("fb0a2a")
	return test
}

// Unexecuted 未実施をセット
func (test *ExTest) Unexecuted() *ExTest {
	test.row.SetString("未実施", 12)

	test.row.GetCell(1).SetBackgroundColor("d9d9d9")
	test.row.GetCell(2).SetBackgroundColor("d9d9d9")
	test.row.GetCell(3).SetBackgroundColor("d9d9d9")
	test.row.GetCell(4).SetBackgroundColor("d9d9d9")
	test.row.GetCell(5).SetBackgroundColor("d9d9d9")
	test.row.GetCell(6).SetBackgroundColor("d9d9d9")
	test.row.GetCell(7).SetBackgroundColor("d9d9d9")
	test.row.GetCell(8).SetBackgroundColor("d9d9d9")
	test.row.GetCell(9).SetBackgroundColor("d9d9d9")
	test.row.GetCell(10).SetBackgroundColor("d9d9d9")
	test.row.GetCell(11).SetBackgroundColor("d9d9d9")
	test.row.GetCell(12).SetBackgroundColor("d9d9d9")
	test.row.GetCell(13).SetBackgroundColor("d9d9d9")
	test.row.GetCell(14).SetBackgroundColor("d9d9d9")
	test.row.GetCell(15).SetBackgroundColor("d9d9d9")
	test.row.GetCell(16).SetBackgroundColor("d9d9d9")
	test.row.GetCell(17).SetBackgroundColor("d9d9d9")
	return test
}

// Info 補足情報をセットする
func (test *ExTest) Info(info string) *ExTest {
	test.row.SetString(time.Now().Format("01/02"), 10)
	test.row.SetString(info, 13)
	return test
}
