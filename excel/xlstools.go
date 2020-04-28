package excel

import (
	"../strutils"
	"fmt"
	"github.com/0x5f81/xls"
	"strings"
)

type XlsTools struct {
	FilePath  string
	StartStr  string
	EndStr    string
	CellStart string
	CellEnd   string
	Separator string
	ColNum    int
	StartRow  int
	SheetNum  int
}

func (xt XlsTools) ParseToSql() (string, error) {
	//起始行数不能小于1
	if xt.StartRow < 1 {
		return "", fmt.Errorf("输入的起始行数%d不能小于1", xt.StartRow)
	}
	//表格位置不能小于1
	if xt.SheetNum < 1 {
		return "", fmt.Errorf("输入的表格位置%d不能小于1", xt.SheetNum)
	}
	//数据列列数不能小于1
	if xt.ColNum < 1 {
		return "", fmt.Errorf("输入的数据列列数%d不能小于1", xt.ColNum)
	}
	//读取excel
	xf, closer, err := xls.OpenWithCloser(xt.FilePath, "utf-8")
	if err != nil {
		return "", err
	}
	//输入的表格位置不能超过excel中Sheet数量
	if xt.SheetNum > xf.NumSheets() {
		return "", fmt.Errorf("输入的表格位置%d超过excel中Sheet数量%d", xt.SheetNum, xf.NumSheets())
	}
	sheet := xf.GetSheet(xt.SheetNum - 1)
	maxRow := int(sheet.MaxRow)
	if xt.StartRow > maxRow {
		return "", fmt.Errorf("输入的起始行数%d超过excel最大行数%d", xt.StartRow, maxRow)
	}
	maxCol := sheet.Row(xt.StartRow - 1).LastCol()
	sb := strutils.NewStringBuilder(xt.StartStr)
	//起始行数不能大于excel的最大行数

	//数据列列数不能大于excel的最大列数
	if xt.ColNum > maxCol {
		return "", fmt.Errorf("输入的数据列列数%d超过excel最大列数%d", xt.ColNum, maxCol)
	}
	//遍历sheet
	for i := xt.StartRow - 1; i <= maxRow; i++ {
		row := sheet.Row(i)
		identity := strings.TrimSpace(row.Col(xt.ColNum - 1))
		if identity != "" {
			sb.Append(xt.CellStart).Append(identity).Append(xt.CellEnd).Append(xt.Separator)
		}

	}
	_ = closer.Close()
	//若有效遍历的行数大于1则需要删除最后一个分割器
	if len(xt.StartStr) < sb.Len() {
		sb.SetLen(sb.Len() - len(xt.Separator))
	}
	sqlStr := sb.Append(xt.EndStr).ToString()
	return sqlStr, nil
}
