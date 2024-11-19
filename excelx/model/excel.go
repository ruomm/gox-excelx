// ------------------------------------------------------------------------
// -------------------           Author：符华            -------------------
// -------------------           Gitee：寒霜剑客          -------------------
// ------------------------------------------------------------------------

package model

import (
	"github.com/pkg/errors"
	"github.com/ruomm/goxframework/gox/corex"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
)

// 定义正则表达式模式
const (
	ExcelTagKey = "excel"
	//Pattern     = "name:(.*?);|index:(.*?);|width:(.*?);|replace:(.*?);|convert:(.*?);|cpoption:(.*?);"
)

// ExcelTag 自定义一个tag结构体
type ExcelTag struct {
	Value    interface{}
	Name     string // 表头标题
	Index    int    // 列下标(从0开始)
	Width    int    // 列宽
	Replace  string // 替换（需要替换的内容_替换后的内容。比如：1_未开始 ==> 表示1替换为未开始）
	Convert  string // 转换方法名
	Cpoption string //方法优化名称
	Align    string // 对齐方式 left、center、right
}

// NewExcelTag 构造函数，返回一个带有默认值的 ExcelTag 实例
func NewExcelTag() ExcelTag {
	return ExcelTag{
		// 导入时会根据这个下标来拿单元格的值，当目标结构体字段没有设置index时，
		// 解析字段tag值时Index没读到就一直默认为0，拿单元格的值时，就始终拿的是第一列的值
		Index: -1, // 设置 Index 的默认值为 -1
	}
}

// GetTag 读取字段tag值
func (e *ExcelTag) GetTag(tag string) (err error) {

	subTags := corex.ParseToSubTag(tag)
	for _, subTag := range subTags {
		e.setValue("name", parseSubTagValue(subTag, "name"))
		e.setValue("index", parseSubTagValue(subTag, "index"))
		e.setValue("width", parseSubTagValue(subTag, "width"))
		e.setValue("replace", parseSubTagValue(subTag, "replace"))
		e.setValue("convert", parseSubTagValue(subTag, "convert"))
		e.setValue("cpoption", parseSubTagValue(subTag, "cpoption"))
		e.setValue("align", parseSubTagValue(subTag, "align"))
	}
	if len(e.Name) <= 0 {
		err = errors.New("未匹配到值")
		return
	}
	return
}
func parseSubTagValue(subTag, optionKey string) string {
	keyLen := len(optionKey)
	if keyLen <= 0 {
		return ""
	}
	if strings.HasPrefix(subTag, optionKey+"=") || strings.HasPrefix(subTag, optionKey+":") || strings.HasPrefix(subTag, optionKey+".") {
		return subTag[keyLen+1:]
	} else {
		return ""
	}
}

// setValue 设置ExcelTag 对应字段的值
func (e *ExcelTag) setValue(tag string, value string) {
	if len(value) <= 0 {
		return
	}
	if strings.Contains(tag, "name") {
		e.Name = value
	}
	if strings.Contains(tag, "index") {
		v, _ := strconv.ParseInt(value, 10, 8)
		e.Index = int(v)
	}
	if strings.Contains(tag, "width") {
		v, _ := strconv.ParseInt(value, 10, 8)
		e.Width = int(v)
	}
	if strings.Contains(tag, "replace") {
		e.Replace = value
	}
	if strings.Contains(tag, "convert") {
		e.Convert = value
	}
	if strings.Contains(tag, "cpoption") {
		e.Cpoption = value
	}
	if strings.Contains(tag, "align") {
		e.Align = value
	}
}

// Excel 自定义一个excel对象结构体
type Excel struct {
	F          *excelize.File // excel 对象
	TitleStyle int            // 表头样式
	HeadStyle  int            // 表头样式
	//ContentStyle1 int            // 主体样式1，无背景色
	//ContentStyle2 int            // 主体样式2，有背景色
	ContentStyleCenter int // 内容主体样式-中间对齐
	ContentStyleLeft   int // 内容主体样式-左对齐
	ContentStyleRight  int // 内容主体样式-右对齐

	MergeRowStyle int // 主体样式2，有背景色
}

// NewExcel 初始化
func NewExcel() (e *Excel) {
	e = &Excel{}
	// excel构建
	e.F = excelize.NewFile()
	return e
}

// SetDefaultStyle
//
//	@Description:  设置默认样式
//	@receiver e
func (e *Excel) SetDefaultStyle() {
	e.SetTitleRowStyle()
	e.SetHeadRowStyle()
	e.SetDataRowStyle()
	e.SetMergeRowStyle()
}

// ===================================== 设置样式 =====================================

// 获取边框样式
func getBorder() []excelize.Border {
	return []excelize.Border{ // 边框
		{Type: "top", Color: "000000", Style: 1},
		{Type: "bottom", Color: "000000", Style: 1},
		{Type: "left", Color: "000000", Style: 1},
		{Type: "right", Color: "000000", Style: 1},
	}
}

// SetTitleRowStyle 标题样式
func (e *Excel) SetTitleRowStyle() {
	e.TitleStyle, _ = e.F.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{ // 对齐方式
			Horizontal: "center", // 水平对齐居中
			Vertical:   "center", // 垂直对齐居中
		},
		Fill: excelize.Fill{ // 背景颜色
			Type:    "pattern",
			Color:   []string{"#fff2cc"},
			Pattern: 1,
		},
		Font: &excelize.Font{ // 字体
			Bold: true,
			Size: 16,
		},
		Border: getBorder(),
	})
}

// SetEndStyle 标题样式
func (e *Excel) SetMergeRowStyle() {
	e.MergeRowStyle, _ = e.F.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{ // 对齐方式
			Horizontal: "right",  // 水平对齐居中
			Vertical:   "center", // 垂直对齐居中
		},
		Fill: excelize.Fill{ // 背景颜色
			Type:    "pattern",
			Color:   []string{"#fff2cc"},
			Pattern: 1,
		},
		Font: &excelize.Font{ // 字体
			Bold: true,
			Size: 16,
		},
		Border: getBorder(),
	})
}

// SetEndStyle 标题样式
func (e *Excel) ParseMergeRowStyle(horizontal string) int {
	mergeRowStyle, _ := e.F.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{ // 对齐方式
			Horizontal: horizontal, // 水平对齐居中
			Vertical:   "center",   // 垂直对齐居中
		},
		Fill: excelize.Fill{ // 背景颜色
			Type:    "pattern",
			Color:   []string{"#fff2cc"},
			Pattern: 1,
		},
		Font: &excelize.Font{ // 字体
			Bold: true,
			Size: 16,
		},
		Border: getBorder(),
	})
	return mergeRowStyle
}

// SetHeadRowStyle 列头行样式
func (e *Excel) SetHeadRowStyle() {
	e.HeadStyle, _ = e.F.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{ // 对齐方式
			Horizontal: "center", // 水平对齐居中
			Vertical:   "center", // 垂直对齐居中
			WrapText:   true,     // 自动换行
		},
		Fill: excelize.Fill{ // 背景颜色
			Type:    "pattern",
			Color:   []string{"#FDE9D8"},
			Pattern: 1,
		},
		Font: &excelize.Font{ // 字体
			Bold: true,
			Size: 14,
		},
		Border: getBorder(),
	})
}

// SetDataRowStyle 数据行样式
func (e *Excel) SetDataRowStyle() {
	style := excelize.Style{}
	style.Border = getBorder()
	style.Alignment = &excelize.Alignment{
		Horizontal: "center", // 水平对齐居中
		Vertical:   "center", // 垂直对齐居中
		WrapText:   true,     // 自动换行
	}
	style.Font = &excelize.Font{
		Size: 12,
	}
	//e.ContentStyle1, _ = e.F.NewStyle(&style)
	/*	style.Fill = excelize.Fill{ // 背景颜色
		Type:    "pattern",
		Color:   []string{"#cce7f5"},
		Pattern: 1,
	}*/
	//e.ContentStyle2, _ = e.F.NewStyle(&style)
	e.ContentStyleCenter, _ = e.F.NewStyle(&style)
	style.Alignment = &excelize.Alignment{
		Horizontal: "left",   // 水平对齐左侧
		Vertical:   "center", // 垂直对齐居中
		WrapText:   true,     // 自动换行
	}
	e.ContentStyleLeft, _ = e.F.NewStyle(&style)
	style.Alignment = &excelize.Alignment{
		Horizontal: "right",  // 水平对齐右侧
		Vertical:   "center", // 垂直对齐居中
		WrapText:   true,     // 自动换行
	}
	e.ContentStyleRight, _ = e.F.NewStyle(&style)
}

// IsContain 判断数组中是否包含指定元素
func IsContain(items interface{}, item interface{}) bool {
	switch items.(type) {
	case []int:
		intArr := items.([]int)
		for _, value := range intArr {
			if value == item.(int) {
				return true
			}
		}
	case []string:
		strArr := items.([]string)
		for _, value := range strArr {
			if value == item.(string) {
				return true
			}
		}
	default:
		return false
	}
	return false
}
