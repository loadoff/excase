package excase

import (
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"

	"io/ioutil"

	"github.com/loadoff/excl"
)

// ExCase テストケース
type ExCase struct {
	FilePath    string
	caseBook    *excl.Workbook
	caseSheet   *excl.Sheet
	caseRow     *excl.Row
	dir         string
	testCount   int
	largeCount  int
	middleCount int
	smallCount  int
	style1      *excl.Style
	style1New   *excl.Style
	style2      *excl.Style
	style2New   *excl.Style
	style3      *excl.Style
	style3New   *excl.Style
	style4      *excl.Style
	style4New   *excl.Style
	style5      *excl.Style
	style5New   *excl.Style
	style6      *excl.Style
	style6New   *excl.Style
	style7      *excl.Style
	style8      *excl.Style
	style9      *excl.Style
	style10     *excl.Style
	style11     *excl.Style
	style12     *excl.Style
	style13     *excl.Style
	style14     *excl.Style
	style15     *excl.Style
	style16     *excl.Style
	style17     *excl.Style
}

// InitExCase Excelテストケースを作成する
func InitExCase() *ExCase {
	var err error
	ex := &ExCase{testCount: 0, largeCount: 0, middleCount: 0, smallCount: 0}
	ex.FilePath = strings.Replace(time.Now().Format("20060102030405"), ".", "", 1) + ".xlsx"
	ex.dir, err = ioutil.TempDir("", "expand"+strings.Replace(time.Now().Format("20060102030405"), ".", "", 1))

	if ex.caseBook, err = excl.CreateWorkbook(ex.dir, ex.FilePath); err != nil {
		fmt.Println(err.Error())
		return nil
	}
	ex.caseSheet, _ = ex.caseBook.OpenSheet("libexcl")
	ex.caseSheet.ShowGridlines(false)
	caseRow := ex.caseSheet.GetRow(4)
	borderSetting := &excl.BorderSetting{Style: "thin"}
	border := excl.Border{Left: borderSetting, Right: borderSetting, Top: borderSetting, Bottom: borderSetting}
	font := excl.Font{Color: "FFFFFF"}
	style := &excl.Style{}
	style.FontID = ex.caseSheet.Styles.SetFont(font)
	style.FillID = ex.caseSheet.Styles.SetBackgroundColor("361e6d")
	style.BorderID = ex.caseSheet.Styles.SetBorder(border)
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
	ex.caseRow = ex.caseSheet.GetRow(ex.testCount + 5)
	// スタイル作成
	style = &excl.Style{}
	border = excl.Border{Left: &excl.BorderSetting{Style: "thin"}}
	style.BorderID = ex.caseSheet.Styles.SetBorder(border)
	ex.style1 = style
	ex.style3 = style
	ex.style5 = style

	style = &excl.Style{}
	border = excl.Border{Left: &excl.BorderSetting{Style: "hair"}}
	style.BorderID = ex.caseSheet.Styles.SetBorder(border)
	ex.style2 = style
	ex.style4 = style
	ex.style6 = style

	style = &excl.Style{}
	border = excl.Border{Left: &excl.BorderSetting{Style: "thin"}, Top: &excl.BorderSetting{Style: "thin"}}
	style.BorderID = ex.caseSheet.Styles.SetBorder(border)
	ex.style1New = style
	ex.style3New = style
	ex.style5New = style
	ex.style7 = style
	ex.style9 = style
	ex.style10 = style
	ex.style11 = style
	ex.style12 = style
	ex.style13 = style
	ex.style14 = style
	ex.style15 = style
	ex.style16 = style

	style = &excl.Style{}
	border = excl.Border{Left: &excl.BorderSetting{Style: "thin"}, Top: &excl.BorderSetting{Style: "thin"}, Right: &excl.BorderSetting{Style: "thin"}}
	style.BorderID = ex.caseSheet.Styles.SetBorder(border)
	ex.style17 = style

	style = &excl.Style{}
	border = excl.Border{Left: &excl.BorderSetting{Style: "hair"}, Top: &excl.BorderSetting{Style: "thin"}}
	style.BorderID = ex.caseSheet.Styles.SetBorder(border)
	ex.style2New = style
	ex.style4New = style
	ex.style6New = style
	ex.style8 = style
	return ex
}

// Close 閉じる
func (ex *ExCase) Close() {
	ex.caseRow = ex.caseSheet.GetRow(ex.testCount + 5)
	border := excl.Border{Top: &excl.BorderSetting{Style: "thin"}}
	style := &excl.Style{}
	style.BorderID = ex.caseSheet.Styles.SetBorder(border)
	ex.caseRow.GetCell(1).SetStyle(style)
	ex.caseRow.GetCell(2).SetStyle(style)
	ex.caseRow.GetCell(3).SetStyle(style)
	ex.caseRow.GetCell(4).SetStyle(style)
	ex.caseRow.GetCell(5).SetStyle(style)
	ex.caseRow.GetCell(6).SetStyle(style)
	ex.caseRow.GetCell(7).SetStyle(style)
	ex.caseRow.GetCell(8).SetStyle(style)
	ex.caseRow.GetCell(9).SetStyle(style)
	ex.caseRow.GetCell(10).SetStyle(style)
	ex.caseRow.GetCell(11).SetStyle(style)
	ex.caseRow.GetCell(12).SetStyle(style)
	ex.caseRow.GetCell(13).SetStyle(style)
	ex.caseRow.GetCell(14).SetStyle(style)
	ex.caseRow.GetCell(15).SetStyle(style)
	ex.caseRow.GetCell(16).SetStyle(style)
	ex.caseRow.GetCell(17).SetStyle(style)
	ex.caseSheet.Close()
	ex.caseBook.Close()
	os.RemoveAll(ex.dir)
}

// Case ケースの作成
func (ex *ExCase) Case() *ExCase {
	ex.testCount++
	ex.caseRow = ex.caseSheet.GetRow(ex.testCount + 4)
	ex.caseRow.GetCell(1).SetStyle(ex.style1)
	ex.caseRow.GetCell(2).SetStyle(ex.style2)
	ex.caseRow.GetCell(3).SetStyle(ex.style3)
	ex.caseRow.GetCell(4).SetStyle(ex.style4)
	ex.caseRow.GetCell(5).SetStyle(ex.style5)
	ex.caseRow.GetCell(6).SetStyle(ex.style6)
	ex.caseRow.GetCell(7).SetStyle(ex.style7)
	ex.caseRow.GetCell(8).SetStyle(ex.style8)
	ex.caseRow.GetCell(9).SetStyle(ex.style9)
	ex.caseRow.GetCell(10).SetStyle(ex.style10)
	ex.caseRow.GetCell(11).SetStyle(ex.style11)
	ex.caseRow.GetCell(12).SetStyle(ex.style12)
	ex.caseRow.GetCell(13).SetStyle(ex.style13)
	ex.caseRow.GetCell(14).SetStyle(ex.style14)
	ex.caseRow.GetCell(15).SetStyle(ex.style15)
	ex.caseRow.GetCell(16).SetStyle(ex.style16)
	ex.caseRow.GetCell(17).SetStyle(ex.style17)
	return ex
}

// Large 大項目をセット
func (ex *ExCase) Large(name string) *ExCase {
	ex.middleCount = 0
	ex.smallCount = 0
	ex.largeCount++
	ex.caseRow.SetString(strconv.Itoa(ex.largeCount), 1).SetStyle(ex.style1New)
	ex.caseRow.SetString(name, 2).SetStyle(ex.style2New)
	return ex
}

// Middle 中項目をセット
func (ex *ExCase) Middle(name string) *ExCase {
	ex.smallCount = 0
	ex.middleCount++
	ex.caseRow.SetNumber(strconv.Itoa(ex.middleCount), 3).SetStyle(ex.style3New)
	ex.caseRow.SetString(name, 4).SetStyle(ex.style4New)
	return ex
}

// Small 小項目をセット
func (ex *ExCase) Small(name string) *ExCase {
	ex.smallCount++
	ex.caseRow.SetNumber(strconv.Itoa(ex.smallCount), 5).SetStyle(ex.style5New)
	ex.caseRow.SetString(name, 6).SetStyle(ex.style6New)
	return ex
}

// Test テストの内容と合格条件をセットする
func (ex *ExCase) Test(content string, pass string) *ExCase {
	ex.caseRow.SetNumber(strconv.Itoa(ex.testCount), 7)
	ex.caseRow.SetString(content, 8)
	ex.caseRow.SetString(pass, 9)
	return ex
}

// Passed 合格をセット
func (ex *ExCase) Passed() *ExCase {
	ex.caseRow.SetString(time.Now().Format("01/02"), 10)
	ex.caseRow.SetString("合格", 12)
	return ex
}

// Failed 不合格をセット
func (ex *ExCase) Failed() *ExCase {
	ex.caseRow.SetString(time.Now().Format("01/02"), 10)
	ex.caseRow.SetString("不合格", 12).SetBackgroundColor("fb0a2a")
	return ex
}
