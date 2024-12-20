// ------------------------------------------------------------------------
// -------------------           Author：符华            -------------------
// -------------------           Gitee：寒霜剑客          -------------------
// ------------------------------------------------------------------------

package excelx

import (
	"fmt"
	"github.com/morrisxyang/xreflect"
	"github.com/pkg/errors"
	"github.com/ruomm/gox-excelx/excelx/model"
	"github.com/ruomm/goxframework/gox/refx"
	"github.com/xuri/excelize/v2"
	"html/template"
	"os"
	"reflect"
	"sort"
	"strconv"
	"strings"
)

// GetExcelColumnName 根据列数生成 Excel 列名
func GetExcelColumnName(columnNumber int) string {
	columnName := ""
	for columnNumber > 0 {
		columnNumber--
		columnName = string(rune('A'+columnNumber%26)) + columnName
		columnNumber /= 26
	}
	return columnName
}

// sheet sheet名称
// title 标题
// fileName 下载的文件名
// isGhbj 是否设置隔行背景色（true 设置 false 不设置）
// isIgnore 是否忽略指定字段（true 要忽略的字段 false 要导出的字段）
// fields 选择的字段，多个字段用逗号隔开，最后一个字段后面也要加逗号，如：字段1,字段2,字段3,
// changeHead 要改变表头的字段，格式是{"字段1":"更改的表头1","字段2":"更改的表头2"}
// heads 表头内容
// list 数据内容
// startColName 开始的列名
// endColName 结束的列名

// 构建标题
func buildTitle(e *model.Excel, sheet, title, endColName string) (dataRow int) {
	dataRow = 2 // 开始的数据行号，默认为1表示一定有一行表头，数据行从第二行开始
	// 标题默认在第一行
	if title != "" {
		dataRow = 3 // 为3表示有一行标题和一行表头，数据行从第三行开始
		e.F.SetCellValue(sheet, "A1", title)
		e.F.MergeCell(sheet, "A1", endColName+"1") // 合并标题单元格
		e.F.SetCellStyle(sheet, "A1", endColName+"1", e.TitleStyle)
		e.F.SetRowHeight(sheet, 1, float64(30)) // 第一行行高
	}
	return
}

// 构建结尾行
func ExportMergeRow(config *model.ExportConfig, mergeStr string) error {
	// 标题默认在第一行
	if mergeStr != "" {
		rowNum := strconv.Itoa(config.RowBuildReq.StartRow)
		err := config.E.F.SetCellValue(config.SheetName, "A"+rowNum, mergeStr)
		if err != nil {
			return err
		}
		err = config.E.F.MergeCell(config.SheetName, "A"+rowNum, config.RowBuildReq.EndColName+rowNum) // 合并单元格
		if err != nil {
			return err
		}
		err = config.E.F.SetCellStyle(config.SheetName, "A"+rowNum, config.RowBuildReq.EndColName+rowNum, config.E.MergeRowStyle)
		if err != nil {
			return err
		}
		err = config.E.F.SetRowHeight(config.SheetName, config.RowBuildReq.StartRow, float64(30)) // 行高
		if err != nil {
			return err
		}
		config.RowBuildReq.StartRow++
		return nil
	} else {
		return nil
	}
}

// 构建结尾行
func ExportMergeRowByStyle(config *model.ExportConfig, mergeStr string, horizontal string) error {
	if mergeStr != "" {
		rowNum := strconv.Itoa(config.RowBuildReq.StartRow)
		err := config.E.F.SetCellValue(config.SheetName, "A"+rowNum, mergeStr)
		if err != nil {
			return err
		}
		err = config.E.F.MergeCell(config.SheetName, "A"+rowNum, config.RowBuildReq.EndColName+rowNum) // 合并单元格
		if err != nil {
			return err
		}
		err = config.E.F.SetCellStyle(config.SheetName, "A"+rowNum, config.RowBuildReq.EndColName+rowNum, config.E.ParseMergeRowStyle(horizontal))
		if err != nil {
			return err
		}
		err = config.E.F.SetRowHeight(config.SheetName, config.RowBuildReq.StartRow, float64(30)) // 行高
		if err != nil {
			return err
		}
		config.RowBuildReq.StartRow++
		return nil
	} else {
		return nil
	}
}

// 构建表头：headerRowNum 当前表头行行号
func buildHeader(e *model.Excel, sheet, endColName string, headerRowNum int, heads *[]string) (err error) {
	row := fmt.Sprintf("%d", headerRowNum)
	e.F.SetRowHeight(sheet, headerRowNum, float64(30))
	e.F.SetCellStyle(sheet, "A"+row, endColName+row, e.HeadStyle)
	return e.F.SetSheetRow(sheet, "A"+row, heads)
}

// 构建标题和表头：headerRowNum 当前表头行行号
func buildTitleHeader(e *model.Excel, sheet, title, endColName string, heads *[]string) (dataRow int, err error) {
	dataRow = buildTitle(e, sheet, title, endColName) // 构建标题，获取第一行数据所在的行号
	// dataRow-1：表头行所在的行号
	err = buildHeader(e, sheet, endColName, dataRow-1, heads)
	return
}

// 构建自定义复杂表头
func buildCustomHeader(heads interface{}, sheet, title string) (*model.Excel, []string, string, int, error) {
	rowsHead := [][]string{}  // 存储多行表头
	lastRowHead := []string{} // 最后一行表头
	// 类型断言，判断是单行表头还是多行表头
	switch heads.(type) {
	case []string: // 单行表头
		lastRowHead = heads.([]string)
	case [][]string: // 复杂表头
		// 复杂表头规定：从一级表头开始到二级、三级...最后一级，每一级表头必须按从上往下顺序存储，最后一级表头必须放在数组最后
		// 每一级表头的列数必须一致，也就是说所有表头的列数必须以最后一级表头的列数为准，如果列不够填相同的内容即可，后续会将相同的内容的列合并。
		// 例如下面这组数据，有四级表头，最后一级有6列，所以一、二、三级表头也需要有6列。然后每一行相同内容的列会合并。
		/*header := [][]string{
			{"一级表头1", "一级表头1", "一级表头1", "一级表头1", "一级表头2", "一级表头2"},
			{"二级表头1", "二级表头1", "二级表头2", "二级表头2", "二级表头3", "二级表头3"},
			{"三级表头1", "三级表头1", "三级表头2", "三级表头2", "三级表头3", "三级表头4"},
			{"四级表头1", "四级表头2", "四级表头3", "四级表头4", "四级表头5", "四级表头6"},
		}*/
		rowsHead = heads.([][]string)
		lastRowHead = rowsHead[len(rowsHead)-1] // 在多行表头中，获取最后一行表头
	default:
		return nil, nil, "", 0, errors.New("表头格式错误")
	}
	e := model.NewExcel()
	e.SetDefaultStyle()
	if sheet != "Sheet1" {
		errRename := e.F.SetSheetName("Sheet1", sheet)
		if errRename != nil {
			return nil, nil, "", 0, errors.New("命名Sheet1为" + sheet + "错误")
		}
	}

	index, _ := e.F.GetSheetIndex(sheet)
	if index < 0 { // 如果sheet名称不存在
		e.F.NewSheet(sheet)
	}
	endColName := GetExcelColumnName(len(lastRowHead)) // 根据列数生成 Excel 列名
	dataRow := 0                                       // 数据行开始的行号，有title时，默认为3（1 为title行，2 为表头行，3 开始就是数据行，包括了3），无title时默认为2（1 为表头行 2 开始就是数据行，包括了2）
	if len(rowsHead) > 0 {
		headRowNum := 1 // 第一行表头行号
		if title != "" {
			dataRow = 1
			headRowNum = 2                          // 有标题是，为2
			buildTitle(e, sheet, title, endColName) // 构建标题
		}
		// 当有多行表头时，数据行号就是 表头数量+1，
		dataRow = dataRow + len(rowsHead) + 1
		for i, items := range rowsHead {
			err := buildHeader(e, sheet, endColName, i+headRowNum, &items) // 构建表头
			if err != nil {
				return nil, nil, "", 0, err
			}
		}
	} else {
		dataRow, _ = buildTitleHeader(e, sheet, title, endColName, &lastRowHead) // 构建标题和表头
	}
	e.F.SetColWidth(sheet, "A", endColName, float64(20)) // 设置列宽
	return e, lastRowHead, endColName, dataRow, nil
}

// ExportExcel excel导出，获取表头、内容数据
func ExportExcel(config *model.ExportConfig, exportData interface{}) (err error) {
	index, _ := config.E.F.GetSheetIndex(config.SheetName)
	if index < 0 { // 如果sheet名称不存在
		_, err = config.E.F.NewSheet(config.SheetName)
		if err != nil {
			return err
		}
	}
	// 构造excel表格
	// 取目标对象的元素类型、字段类型和 tag
	dataValue := reflect.ValueOf(exportData)
	// 判断数据的类型
	if dataValue.Kind() != reflect.Slice {
		err = errors.New("invalid data type")
		return
	}
	// 构造表头
	endColName, dataRow, err := normalBuildTitle(config.E, config, dataValue)
	if err != nil {
		return
	}
	config.RowBuildReq = &model.DataRowBuildReq{
		Excel:          config.E,
		SheetName:      config.SheetName,
		StringMode:     config.StringMode,
		EndColName:     endColName,
		StartRow:       dataRow,
		IsDefaultStyle: config.IsDefaultStyle,
		DataValue:      dataValue,
		DynamicColConfig: model.DynamicColConfig{
			Fields:   config.Fields,
			IsIgnore: config.IsIgnore,
		},
	}
	// 构造数据行
	err = normalBuildDataRow(config.RowBuildReq)
	return
}

func ExportRowData(config *model.ExportConfig, exportData interface{}) (err error) {
	// 取目标对象的元素类型、字段类型和 tag
	dataValue := reflect.ValueOf(exportData)
	// 判断数据的类型
	if dataValue.Kind() != reflect.Slice {
		err = errors.New("invalid data type")
		return
	}
	config.RowBuildReq.DataValue = dataValue
	// 构造数据行
	err = normalBuildDataRow(config.RowBuildReq)
	return
}

// ================================= 普通导出 =================================

// NormalDynamicExport 导出excel
func NormalDynamicExport(config *model.ExportConfig) (file *excelize.File, err error) {
	if config.E == nil {
		config.E = model.NewExcel()
		config.E.SetDefaultStyle()
		if config.SheetName != "Sheet1" {
			errRename := config.E.F.SetSheetName("Sheet1", config.SheetName)
			if errRename != nil {
				return nil, errors.New("命名Sheet1为" + config.SheetName + "错误")
			}
		}
	}
	err = ExportExcel(config, config.ExportData)
	return config.E.F, err
}

// CustomHeaderExport 自定义表头导出
func CustomHeaderExport(sheet, title string, isGhbj bool, heads interface{}, list interface{}) (file *excelize.File, err error) {
	e, _, endColName, dataRow, err := buildCustomHeader(heads, sheet, title)
	if err != nil {
		return
	}
	dataValue := reflect.ValueOf(list)
	// 判断数据的类型
	if dataValue.Kind() != reflect.Slice {
		err = errors.New("invalid data type")
		return
	}
	// 构造数据行
	err = normalBuildDataRow(&model.DataRowBuildReq{
		Excel:          e,
		SheetName:      sheet,
		StringMode:     false,
		EndColName:     endColName,
		StartRow:       dataRow,
		IsDefaultStyle: isGhbj,
		DataValue:      dataValue,
		DynamicColConfig: model.DynamicColConfig{
			Fields:   "",
			IsIgnore: false,
		},
	})
	return e.F, err
}

// 构造表头（endColName 最后一列的列名 dataRow 数据行开始的行号）
func normalBuildTitle(e *model.Excel, config *model.ExportConfig, dataValue reflect.Value) (endColName string, dataRow int, err error) {
	dataType := dataValue.Type().Elem() // 获取导入目标对象的类型信息
	if dataType.Kind() == reflect.Ptr {
		// 获取指针指向的类型
		dataType = dataType.Elem()
	}
	var exportTitle []model.ExcelTag // 遍历目标对象的字段
	for j := 0; j < dataType.NumField(); j++ {
		var dataCol model.ExcelTag
		field := dataType.Field(j) // 获取字段信息和tag
		tag := field.Tag.Get(model.ExcelTagKey)
		if tag == "" { // 如果非导出则跳过
			continue
		}
		if config.Fields != "" { // 选择要导出或要忽略的字段
			if config.IsIgnore && strings.Contains(config.Fields, field.Name) { // 忽略指定字段
				continue
			}
			if !config.IsIgnore && !strings.Contains(config.Fields, field.Name) { // 导出指定字段
				continue
			}
		}
		err = dataCol.GetTag(tag)
		if err != nil {
			return
		}
		// 更改指定字段的表头标题
		if config.ChangeHead != nil {
			tmpName, tmpOk := config.ChangeHead[field.Name]
			if tmpOk && tmpName != "" {
				dataCol.Name = tmpName
			}
		}
		if config.ReplaceHead != nil {
			for k, v := range config.ReplaceHead {
				dataCol.Name = strings.ReplaceAll(dataCol.Name, "${"+k+"}", v)
			}
		}
		dataCol.FieldIndex = j
		exportTitle = append(exportTitle, dataCol)
	}
	// 排序
	sort.Slice(exportTitle, func(i, j int) bool {
		if exportTitle[i].Index == exportTitle[j].Index {
			return exportTitle[i].FieldIndex < exportTitle[j].FieldIndex
		}
		return exportTitle[i].Index < exportTitle[j].Index
	})
	var titleRowData []string // 列头行
	for i, colTitle := range exportTitle {
		endColName := GetExcelColumnName(i + 1)
		if colTitle.Width > 0 { // 根据给定的宽度设置列宽
			_ = e.F.SetColWidth(config.SheetName, endColName, endColName, float64(colTitle.Width))
		} else {
			_ = e.F.SetColWidth(config.SheetName, endColName, endColName, float64(20)) // 默认宽度为20
		}
		titleRowData = append(titleRowData, colTitle.Name)
	}
	endColName = GetExcelColumnName(len(titleRowData)) // 根据列数生成 Excel 列名
	dataRow, err = buildTitleHeader(e, config.SheetName, config.Title, endColName, &titleRowData)
	return
}
func normalBuildDataRow(buildReq *model.DataRowBuildReq) (err error) {
	if buildReq.StringMode {
		return normalBuildDataRowStringMode(buildReq)
	} else {
		return normalBuildDataRowObjMode(buildReq)
	}
}

// 构造数据行
func normalBuildDataRowObjMode(buildReq *model.DataRowBuildReq) (err error) {
	//实时写入数据
	for i := 0; i < buildReq.DataValue.Len(); i++ {
		startCol := fmt.Sprintf("A%d", buildReq.StartRow)
		endCol := fmt.Sprintf("%s%d", buildReq.EndColName, buildReq.StartRow)
		item := buildReq.DataValue.Index(i)
		if item.Kind() == reflect.Ptr {
			item = item.Elem()
		}
		typ := item.Type()
		num := item.NumField()
		var exportRow []model.ExcelTag
		maxLen := 0 // 记录这一行中，数据最多的单元格的值的长度
		//遍历结构体的所有字段
		for j := 0; j < num; j++ {
			dataField := typ.Field(j) //获取到struct标签，需要通过reflect.Type来获取tag标签的值
			tagVal := dataField.Tag.Get(model.ExcelTagKey)
			if tagVal == "" { // 如果非导出则跳过
				continue
			}
			if buildReq.Fields != "" { // 选择要导出或要忽略的字段
				if buildReq.IsIgnore && strings.Contains(buildReq.Fields, dataField.Name+",") { // 忽略指定字段
					continue
				}
				if !buildReq.IsIgnore && !strings.Contains(buildReq.Fields, dataField.Name+",") { // 导出指定字段
					continue
				}
			}
			var dataCol model.ExcelTag
			err = dataCol.GetTag(tagVal)
			if err != nil {
				return
			}
			fieldData := item.FieldByName(dataField.Name) // 取字段值
			if len(dataCol.Convert) > 0 {
				origValueByMethod, errByMethod := xParseOrigValueByMethod(dataCol.Convert, fieldData.Interface(), item.Interface())
				if errByMethod != nil {
					return
				} else {
					fieldData = reflect.ValueOf(origValueByMethod)
				}
			}
			// 替换
			if dataCol.Replace != "" {
				split := strings.Split(dataCol.Replace, ",")
				for j := range split {
					s := strings.Split(split[j], "_") // 根据下划线进行分割，格式：需要替换的内容_替换后的内容
					value := fieldData.String()
					// 判断当前字段类型，都转为string类型
					if strings.Contains(fieldData.Type().String(), "int") {
						value = strconv.Itoa(int(fieldData.Int()))
					} else if fieldData.Type().String() == "bool" {
						value = strconv.FormatBool(fieldData.Bool())
					} else if strings.Contains(fieldData.Type().String(), "float") {
						value = strconv.FormatFloat(fieldData.Float(), 'f', -1, 64)
					}
					if s[0] == value {
						dataCol.Value = s[1]
					}
				}
			} else {
				dataCol.Value = fieldData
			}
			// 转换函数
			//if dataCol.Convert != "" && convert.BaseConvertMap[dataCol.Convert] != nil {
			//	dataCol.Value = convert.BaseConvertMap[dataCol.Convert](fieldData)
			//}
			if fieldData.Type().String() == "string" { // string类型的才计算长度
				rwsTemp := fieldData.Len() // 当前单元格内容的长度
				if rwsTemp > maxLen {      //这里取每一行中的每一列字符长度最大的那一列的字符
					maxLen = rwsTemp
				}
			}
			dataCol.FieldIndex = j
			exportRow = append(exportRow, dataCol)
		}
		// 排序
		sort.Slice(exportRow, func(i, j int) bool {
			if exportRow[i].Index == exportRow[j].Index {
				return exportRow[i].FieldIndex < exportRow[j].FieldIndex
			}
			return exportRow[i].Index < exportRow[j].Index
		})
		var rowData []interface{} // 数据列
		for _, colTitle := range exportRow {
			rowData = append(rowData, colTitle.Value)
		}
		//if buildReq.IsDefaultStyle && buildReq.StartRow%2 == 0 {
		//	_ = buildReq.Excel.F.SetCellStyle(buildReq.SheetName, startCol, endCol, buildReq.Excel.ContentStyle2)
		//} else {
		//	_ = buildReq.Excel.F.SetCellStyle(buildReq.SheetName, startCol, endCol, buildReq.Excel.ContentStyle1)
		//}
		//_ = buildReq.Excel.F.SetCellStyle(buildReq.SheetName, startCol, endCol, buildReq.Excel.ContentStyleCenter)
		for colIndex, dataCol := range exportRow {
			tmpColName := GetExcelColumnName(colIndex + 1)
			tmpCol := fmt.Sprintf("%s%d", tmpColName, buildReq.StartRow)
			if dataCol.Align == "left" {
				_ = buildReq.Excel.F.SetCellStyle(buildReq.SheetName, tmpCol, tmpCol, buildReq.Excel.ContentStyleLeft)
			} else if dataCol.Align == "right" {
				_ = buildReq.Excel.F.SetCellStyle(buildReq.SheetName, tmpCol, endCol, buildReq.Excel.ContentStyleRight)
			} else {
				_ = buildReq.Excel.F.SetCellStyle(buildReq.SheetName, tmpCol, tmpCol, buildReq.Excel.ContentStyleCenter)
			}
		}
		if maxLen > 25 { // 自适应行高
			d := maxLen / 25
			f := 25 * d
			_ = buildReq.Excel.F.SetRowHeight(buildReq.SheetName, buildReq.StartRow, float64(f))
		} else {
			_ = buildReq.Excel.F.SetRowHeight(buildReq.SheetName, buildReq.StartRow, float64(25)) // 默认行高25
		}
		if err = buildReq.Excel.F.SetSheetRow(buildReq.SheetName, startCol, &rowData); err != nil {
			return
		}
		buildReq.StartRow++
	}
	return
}

// 构造数据行
func normalBuildDataRowStringMode(buildReq *model.DataRowBuildReq) (err error) {
	//实时写入数据
	for i := 0; i < buildReq.DataValue.Len(); i++ {
		startCol := fmt.Sprintf("A%d", buildReq.StartRow)
		endCol := fmt.Sprintf("%s%d", buildReq.EndColName, buildReq.StartRow)
		item := buildReq.DataValue.Index(i)
		if item.Kind() == reflect.Ptr {
			item = item.Elem()
		}
		typ := item.Type()
		num := item.NumField()
		var exportRow []model.ExcelTag
		maxLen := 0 // 记录这一行中，数据最多的单元格的值的长度
		//遍历结构体的所有字段
		for j := 0; j < num; j++ {
			dataField := typ.Field(j) //获取到struct标签，需要通过reflect.Type来获取tag标签的值
			tagVal := dataField.Tag.Get(model.ExcelTagKey)
			if tagVal == "" { // 如果非导出则跳过
				continue
			}
			if buildReq.Fields != "" { // 选择要导出或要忽略的字段
				if buildReq.IsIgnore && strings.Contains(buildReq.Fields, dataField.Name+",") { // 忽略指定字段
					continue
				}
				if !buildReq.IsIgnore && !strings.Contains(buildReq.Fields, dataField.Name+",") { // 导出指定字段
					continue
				}
			}
			var dataCol model.ExcelTag
			err = dataCol.GetTag(tagVal)
			if err != nil {
				return
			}
			origValue, _ := xreflect.EmbedFieldValue(item, dataField.Name)
			if len(dataCol.Convert) > 0 {
				origValueByMethod, errByMethod := xParseOrigValueByMethod(dataCol.Convert, origValue, item.Interface())
				if errByMethod != nil {
					return
				} else {
					origValue = origValueByMethod
				}
			}
			vi := refx.ParseToString(origValue, dataCol.Cpoption)
			fieldDataStr := ""
			if vi == nil {
				fieldDataStr = ""
			} else {
				fieldDataStr = vi.(string)
			}
			// 替换
			if dataCol.Replace != "" {
				split := strings.Split(dataCol.Replace, ",")
				for j := range split {
					s := strings.Split(split[j], "_") // 根据下划线进行分割，格式：需要替换的内容_替换后的内容
					value := fieldDataStr
					if s[0] == value {
						dataCol.Value = s[1]
					}
				}
			} else {
				dataCol.Value = fieldDataStr
			}
			rwsTemp := len(fieldDataStr) // 当前单元格内容的长度
			if rwsTemp > maxLen {        //这里取每一行中的每一列字符长度最大的那一列的字符
				maxLen = rwsTemp
			}
			dataCol.FieldIndex = j
			exportRow = append(exportRow, dataCol)
		}
		// 排序
		sort.Slice(exportRow, func(i, j int) bool {
			if exportRow[i].Index == exportRow[j].Index {
				return exportRow[i].FieldIndex < exportRow[j].FieldIndex
			}
			return exportRow[i].Index < exportRow[j].Index
		})
		var rowData []interface{} // 数据列
		for _, colTitle := range exportRow {
			rowData = append(rowData, colTitle.Value)
		}
		//if buildReq.IsDefaultStyle && buildReq.StartRow%2 == 0 {
		//	_ = buildReq.Excel.F.SetCellStyle(buildReq.SheetName, startCol, endCol, buildReq.Excel.ContentStyle2)
		//} else {
		//	_ = buildReq.Excel.F.SetCellStyle(buildReq.SheetName, startCol, endCol, buildReq.Excel.ContentStyle1)
		//}
		//_ = buildReq.Excel.F.SetCellStyle(buildReq.SheetName, startCol, endCol, buildReq.Excel.ContentStyleCenter)
		for colIndex, dataCol := range exportRow {
			tmpColName := GetExcelColumnName(colIndex + 1)
			tmpCol := fmt.Sprintf("%s%d", tmpColName, buildReq.StartRow)
			if dataCol.Align == "left" {
				_ = buildReq.Excel.F.SetCellStyle(buildReq.SheetName, tmpCol, tmpCol, buildReq.Excel.ContentStyleLeft)
			} else if dataCol.Align == "right" {
				_ = buildReq.Excel.F.SetCellStyle(buildReq.SheetName, tmpCol, endCol, buildReq.Excel.ContentStyleRight)
			} else {
				_ = buildReq.Excel.F.SetCellStyle(buildReq.SheetName, tmpCol, tmpCol, buildReq.Excel.ContentStyleCenter)
			}
		}
		if maxLen > 25 { // 自适应行高
			d := maxLen / 25
			f := 25 * d
			_ = buildReq.Excel.F.SetRowHeight(buildReq.SheetName, buildReq.StartRow, float64(f))
		} else {
			_ = buildReq.Excel.F.SetRowHeight(buildReq.SheetName, buildReq.StartRow, float64(25)) // 默认行高25
		}
		if err = buildReq.Excel.F.SetSheetRow(buildReq.SheetName, startCol, &rowData); err != nil {
			return
		}
		buildReq.StartRow++
	}
	return
}

// ================================= 基于map导出 =================================

// MapExport map导出
func MapExport(req *model.MapExportConfig) (file *excelize.File, err error) {
	e, lastRowHead, endColName, dataRow, err := buildCustomHeader(req.Heads, req.SheetName, req.Title)
	if err != nil {
		return nil, err
	}
	// 构建数据行
	for _, rowData := range req.ExportData {
		startCol := fmt.Sprintf("A%d", dataRow)
		endCol := fmt.Sprintf("%s%d", endColName, dataRow)
		row := make([]interface{}, 0)
		for _, v := range lastRowHead {
			if val, ok := rowData[v]; ok {
				row = append(row, val)
			}
		}
		//if req.IsDefaultStyle && dataRow%2 == 0 {
		//	_ = e.F.SetCellStyle(req.SheetName, startCol, endCol, e.ContentStyle2)
		//} else {
		//	_ = e.F.SetCellStyle(req.SheetName, startCol, endCol, e.ContentStyle1)
		//}
		_ = e.F.SetCellStyle(req.SheetName, startCol, endCol, e.ContentStyleCenter)
		_ = e.F.SetRowHeight(req.SheetName, dataRow, float64(25)) // 默认行高25
		if err := e.F.SetSheetRow(req.SheetName, fmt.Sprintf("A%d", dataRow), &row); err != nil {
			return nil, err
		}
		dataRow++
	}
	return e.F, nil
}

// ================================== 合并单元格 ==================================

// 横向合并单元格：startRowNum（开始合并的行号，注意是行号从1开始，不是索引）endRowNum（停止合并的行号，从1开始，从endRowNum开始，包括这一行不进行合并；不需要停止合并的话传-1）
func HorizontalMerge(f *excelize.File, sheet string, startRowNum, endRowNum int) {
	// startRowNum：比如第一行是标题，第二行是表头，所以从第三行开始合并，startRowNum = 3
	rows, _ := f.GetRows(sheet) // 获取sheet的所有行，包括 标题、表头行（如果有标题和表头的话）
	// row 行号，从1开始
	for row := 1; row <= len(rows); row++ {
		if row < startRowNum { // 如果当前行号，小于开始合并的行号，则跳过
			continue
		}
		if endRowNum > 0 && row >= endRowNum { // 如果当前行号，大于等于结束合并的行号，退出合并
			break
		}
		prevValue := ""     // 上一单元格的值
		mergeStartCol := 0  // 开始合并的单元格列索引
		cols := rows[row-1] // 当前行的列数据（当前行每个单元格的数据）
		// 遍历单元格时，判断当前单元格和上一单元格的值是否相同，相同继续，不同则判断合并，并且将当前单元格的值和索引，赋值给对应的变量。
		/** 比如：a,b,b,b,c,c 这六个单元格的值，
		第一个值 a != ""，进入判断， i-mergeStartCol = 0-0，不进行合并，prevValue=a，mergeStartCol=0
		第二个值 b != a，进入判断，i-mergeStartCol = 1-0，不进行合并，prevValue=b，mergeStartCol=1
		第三、四个值 不进入 cellValue != prevValue 的判断，i分别为2、3
		第五个值 c != b，进入判断，i-mergeStartCol = 4-1，合并 B1:D1，prevValue=c，mergeStartCol=4
		第六个值不进入 cellValue != prevValue 的判断，也结束了for循环，会在 len(cols)-mergeStartCol > 0 这个判断里面进行合并
		*/
		for i, col := range cols {
			cellValue := col // 当前单元格的值
			// 如果当前单元格的值和上一个单元格的值不相等
			if cellValue != prevValue {
				// 当前单元格的列索引 - 开始合并的单元格列索引 大于1，则进行合并
				if i-mergeStartCol > 1 {
					// 获取开始合并的单元格
					startCell := GetExcelColumnName(mergeStartCol+1) + fmt.Sprintf("%d", row)
					// 获取结束合并的单元格
					endCell := GetExcelColumnName(i) + fmt.Sprintf("%d", row)
					f.MergeCell(sheet, startCell, endCell)
				}
				prevValue = cellValue
				mergeStartCol = i
			}
		}
		// 如果最后一个值和上一个值不同，则肯定会合并前面的单元格；如果最后一个值和上一个值相同，则会在这个判断里面进行合并
		if len(cols)-mergeStartCol > 0 {
			startCell := GetExcelColumnName(mergeStartCol+1) + fmt.Sprintf("%d", row)
			endCell := GetExcelColumnName(len(cols)) + fmt.Sprintf("%d", row)
			//fmt.Println(startCell, ":", endCell)
			f.MergeCell(sheet, startCell, endCell)
		}
	}
}

// 纵向合并单元格：headIndex 表头所在索引（一般情况下，不管表头有多少行，只要有标题headIndex都传1，无标题传0）
// needColIndex 需要合并的列号（列号从1开始，如全部列都需合并，传nil就行）
func VerticalMerge(f *excelize.File, sheet string, headIndex int, needColIndex []int) {
	rows, _ := f.GetRows(sheet) // 获取sheet的所以行，包括 标题、表头行（如果有标题和表头的话）
	// 遍历每一列
	for colIndex := 1; colIndex <= len(rows[headIndex]); colIndex++ {
		if len(needColIndex) > 0 && !model.IsContain(needColIndex, colIndex) {
			continue
		}
		startRow := headIndex + 1 // 开始合并的行号
		endRow := headIndex + 1   // 结束结束的行号
		prevValue := rows[headIndex][colIndex-1]
		// 遍历每一行
		for rowIndex := headIndex; rowIndex < len(rows); rowIndex++ {
			row := rows[rowIndex]
			// 因为获取rows时，会忽略空单元格，如果存在空单元格，那每一行的列数并不是相同的，所以需要判断列号是否大于当前行的列数
			if colIndex <= len(row) {
				// 判断当前单元格的值和上一个单元格的值是否相同
				if row[colIndex-1] == prevValue {
					endRow = rowIndex + 1 // 相同，则更新结束合并的行号
				} else {
					if startRow != endRow {
						colName := GetExcelColumnName(colIndex)
						f.MergeCell(sheet, colName+fmt.Sprintf("%d", startRow), colName+fmt.Sprintf("%d", endRow))
					}
					startRow = rowIndex + 1
					endRow = rowIndex + 1
					prevValue = row[colIndex-1]
				}
			}
		}
		// 处理最后一组相同内容的单元格
		if startRow != endRow {
			colName := GetExcelColumnName(colIndex)
			f.MergeCell(sheet, colName+fmt.Sprintf("%d", startRow), colName+fmt.Sprintf("%d", endRow))
		}
	}
}

// ================================== 基于模板导出 ==================================

// TemplateExport 基于excel的模板导出
func TemplateExport(templatePath, outPath string, data map[string]interface{}) error {
	// 解析模板
	tmpl, err := template.ParseFiles(templatePath)
	if err != nil {
		return errors.New("模板解析失败：" + err.Error())
	}
	// 创建输出文件
	file, err := os.Create(outPath)
	if err != nil {
		return errors.New("创建输出文件失败：" + err.Error())
	}
	defer file.Close()
	// 渲染模板并输出结果
	err = tmpl.Execute(file, data)
	if err != nil {
		return errors.New("模板数据渲染失败：" + err.Error())
	}
	return nil
}

// 来源方法转换赋值
func xParseOrigValueByMethod(method_trans string, origVal interface{}, destO interface{}) (interface{}, error) {
	actualValue := reflect.ValueOf(origVal)
	callResultValues, err := xreflect.CallMethod(destO, method_trans, actualValue.Interface())
	if err != nil {
		return nil, errors.New("字段无法赋值，来源字段转换方法执行错误。")
	}
	if callResultValues == nil || len(callResultValues) <= 0 {
		return nil, errors.New("字段无需赋值，来源字段转换方法无有效值返回。")
	}
	actualDestValue := callResultValues[0]
	if actualDestValue.Kind() == reflect.Pointer || actualDestValue.Kind() == reflect.Interface {
		if actualDestValue.IsNil() {
			return nil, errors.New("字段无需赋值，来源字段值为nil，来源字段转换方法无法执行。")
		}
		actualDestValue = actualDestValue.Elem()
	}
	return actualDestValue.Interface(), nil
}
