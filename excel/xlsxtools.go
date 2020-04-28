package excel

import (
	"../strutils"
	"fmt"
	"github.com/tealeg/xlsx"
	"strings"
)

type XlsxTools struct {
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

func (xxt XlsxTools) ParseToSql() (string, error) {
	//起始行数不能小于1
	if xxt.StartRow < 1 {
		return "", fmt.Errorf("输入的起始行数%d不能小于1", xxt.StartRow)
	}
	//表格位置不能小于1
	if xxt.SheetNum < 1 {
		return "", fmt.Errorf("输入的表格位置%d不能小于1", xxt.SheetNum)
	}
	//数据列列数不能小于1
	if xxt.ColNum < 1 {
		return "", fmt.Errorf("输入的数据列列数%d不能小于1", xxt.ColNum)
	}
	//读取excel
	xf, err := xlsx.OpenFile(xxt.FilePath)
	if err != nil {
		return "", err
	}
	//输入的表格位置不能超过excel中Sheet数量
	if xxt.SheetNum > len(xf.Sheets) {
		return "", fmt.Errorf("输入的表格位置%d超过excel中Sheet数量%d", xxt.SheetNum, len(xf.Sheets))
	}
	sheet := xf.Sheets[xxt.SheetNum-1]
	maxRow := sheet.MaxRow
	maxCol := sheet.MaxCol
	sb := strutils.NewStringBuilder(xxt.StartStr)
	//起始行数不能大于excel的最大行数
	if xxt.StartRow > maxRow {
		return "", fmt.Errorf("输入的起始行数%d超过excel最大行数%d", xxt.StartRow, maxRow)
	}
	//数据列列数不能大于excel的最大列数
	if xxt.ColNum > maxCol {
		return "", fmt.Errorf("输入的数据列列数%d超过excel最大列数%d", xxt.ColNum, maxCol)
	}
	//遍历sheet
	for i := xxt.StartRow - 1; i <= maxRow; i++ {
		row, err := sheet.Row(i)
		if err != nil {
			continue
		}
		identity := strings.TrimSpace(row.GetCell(xxt.ColNum - 1).String())
		if identity != "" {
			sb.Append(xxt.CellStart).Append(identity).Append(xxt.CellEnd).Append(xxt.Separator)
		}

	}
	//若有效遍历的行数大于1则需要删除最后一个分割器
	if len(xxt.StartStr) < sb.Len() {
		sb.SetLen(sb.Len() - len(xxt.Separator))
	}
	sqlStr := sb.Append(xxt.EndStr).ToString()
	return sqlStr, nil
}
