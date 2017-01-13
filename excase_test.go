package excase

import (
	"os"
	"testing"
)

func TestExCase(t *testing.T) {
	ex := InitExCase()
	ex.Case().Large("大項目1").Middle("中項目1-1").Small("小項目1-1-1").Test("テスト内容1-1-1-1", "合格条件1-1-1-1")
	ex.Passed()
	ex.Case().Small("小項目1-1-2").Test("テスト内容1-1-1-2", "合格条件1-1-1-2")
	ex.Passed()
	ex.Case().Middle("中項目1-2").Small("小項目1-2-1").Test("テスト内容1-2-1-1", "合格条件1-2-1-1")
	ex.Passed()
	ex.Case().Small("小項目1-2-2").Test("テスト内容1-2-2-1", "合格条件1-2-2-1")
	ex.Failed()
	ex.Case().Test("テスト内容1-2-2-2", "合格条件1-2-2-2")
	ex.Case().Large("大項目2").Middle("中項目2-1").Small("小項目2-1-1").Test("テスト内容2-1-1-1", "合格条件2-1-1-1")
	ex.Case().Small("小項目2-1-2").Test("テスト内容2-1-2-1", "合格条件2-1-2-1")
	ex.Close()

	if stat, err := os.Stat(ex.FilePath); err != nil {
		t.Error("Excel file should be exist.")
	} else if stat.IsDir() {
		t.Error("Excel file should not be directory.")
	}
	os.Remove(ex.FilePath)
}
