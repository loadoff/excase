package excase

import (
	"os"
	"testing"
)

func TestExCase(t *testing.T) {
	c := InitExCase()
	ex := c.OpenSection("ExCaseのテスト")
	test := ex.Large("大項目1").Middle("中項目1-1").Small("小項目1-1-1").Test("テスト内容1-1-1-1", "合格条件1-1-1-1")
	test.Passed()
	test = ex.Small("小項目1-1-2").Test("テスト内容1-1-1-2", "合格条件1-1-1-2")
	test.Passed()
	test = ex.Middle("中項目1-2").Small("小項目1-2-1").Test("テスト内容1-2-1-1", "合格条件1-2-1-1")
	test.Passed()
	test = ex.Small("小項目1-2-2").Test("テスト内容1-2-2-1", "合格条件1-2-2-1")
	test.Failed()
	test = ex.Test("テスト内容1-2-2-2", "合格条件1-2-2-2")
	test = ex.Large("大項目2").Middle("中項目2-1").Small("小項目2-1-1").Test("テスト内容2-1-1-1", "合格条件2-1-1-1")
	test = ex.Small("小項目2-1-2").Test("テスト内容2-1-2-1", "合格条件2-1-2-1")
	c.Close()
	if stat, err := os.Stat(c.FilePath); err != nil {
		t.Error("Excel file should be exist.")
	} else if stat.IsDir() {
		t.Error("Excel file should not be directory.")
	}
	os.Remove(c.FilePath)
}
