package main

import (
	"errors"
	"fmt"
	"io"
	"net/http"
	"regexp"
	"strings"

	"github.com/PuerkitoBio/goquery"
	"github.com/spf13/cast"
	"github.com/xuri/excelize/v2"
)

type Row struct {
	Name           string  //合约品种
	Code           string  //合约代码
	NameCode       string  //品种代码
	Price          float64 //现价
	CS             float64 //交易乘数
	YSSZ           float64 //一手市值
	OpenBuyMargin  string  //买开保证金
	OpenSellMargin string  //卖开保证金
	FeeUnit        string  //手续费单位

	BuyFeeStr           string  //开仓手续费
	BuyFee              float64 //开仓手续费
	SellTodayFeeStr     string  //平今手续费
	SellTodayFee        float64 //平今手续费
	SellYesterdayFeeStr string  //平昨手续费
	SellYesterdayFee    float64 //平昨手续费

	OneDotProfile    float64 //每跳毛利
	BuySellFee       float64 //手续费(开+平今)
	OneDotNetProfile float64 //每跳净利/元
}

// 将数字转换为Excel列的标识（A-Z）
var ExcelChar = []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

func ConvertNumToChar(num int) (string, error) {
	if num < 27 {
		return ExcelChar[num], nil
	}
	k := num % 26
	if k == 0 {
		k = 26
	}
	v := (num - k) / 26
	col, err := ConvertNumToChar(v)
	if err != nil {
		return "", err
	}
	cols := col + ExcelChar[k]
	return cols, nil
}

func getContent() (string, *[]string, *map[int]*[]*Row, int, error) {
	res, err := http.Get("https://www.iweiai.com/")
	if err != nil {
		return "", nil, nil, 0, err
	}
	defer res.Body.Close()

	if res.StatusCode != 200 {
		return "", nil, nil, 0, fmt.Errorf("status code error: %d %s", res.StatusCode, res.Status)
	}

	html_bytes, err := io.ReadAll(res.Body)
	if err != nil {
		return "", nil, nil, 0, err
	}

	html := string(html_bytes)
	html = strings.ReplaceAll(html, "\n", "")
	html = strings.ReplaceAll(html, "/元", "")
	html = strings.ReplaceAll(html, "元", "")

	// Load the HTML document
	doc, err := goquery.NewDocumentFromReader(strings.NewReader(html))
	if err != nil {
		return "", nil, nil, 0, err
	}

	//root
	container := doc.Find(".container")
	container_sub := container.Children()

	//2024年最新期货保证金和手续费查询 最近更新：2024-11-19 20:33:51
	header := container_sub.Filter(".page-header").Find("h1 small").Text()
	headers := strings.Split(header, "：")
	if len(headers) <= 1 {
		return "", nil, nil, 0, errors.New("header error")
	}
	header = "期货手续费和保证金" + headers[1]

	//data
	dataes := make(map[int]*[]*Row)
	exchange_names := make([]string, 0)
	data_size := 0

	//exchange
	container_sub.Filter(".panel-primary").Each(func(i int, s *goquery.Selection) {
		i_sub := s.Children()

		//上期所主力合约
		exchange_name := strings.Trim(i_sub.Filter(".panel-heading").Text(), " ")
		exchange_names = append(exchange_names, exchange_name)

		table := i_sub.Find("table")
		rows := make([]*Row, 0)
		table.Children().Filter("tbody").Children().Filter("[class !=active]").Each(func(_ int, tr *goquery.Selection) {
			str_sub := tr.Children()
			row := new(Row)

			//合约品种
			reg := regexp.MustCompile(`\d+$`)
			row.Name = reg.ReplaceAllString(strings.Trim(str_sub.Eq(0).Text(), " "), "")

			if row.Name != "中证1000股指" && row.Name != "沪深300股指" && row.Name != "上证50股指" && row.Name != "中证500股指" &&
				row.Name != "债十" && row.Name != "债五" && row.Name != "债三十" && row.Name != "债二" {
				//合约代码
				row.Code = strings.Trim(str_sub.Eq(1).Text(), " ")
				row.NameCode = row.Name + "(" + row.Code + ")"

				//现价
				row.Price = cast.ToFloat64(strings.Trim(str_sub.Eq(2).Text(), " "))
				//交易乘数
				row.CS = cast.ToFloat64(strings.Trim(str_sub.Eq(3).Text(), " "))
				//一手市值
				row.YSSZ = cast.ToFloat64(strings.Trim(str_sub.Eq(4).Text(), " "))

				//买开保证金
				row.OpenBuyMargin = strings.Trim(str_sub.Eq(5).Text(), " ")
				//卖开保证金
				row.OpenSellMargin = strings.Trim(str_sub.Eq(6).Text(), " ")
				//手续费单位
				row.FeeUnit = strings.Trim(str_sub.Eq(7).Text(), " ")

				//开仓手续费
				row.BuyFeeStr = strings.Trim(str_sub.Eq(8).Text(), " ")
				buyFeeStrs := strings.Split(row.BuyFeeStr, "≈")
				if len(buyFeeStrs) <= 1 {
					row.BuyFee = cast.ToFloat64(row.BuyFeeStr)
				} else {
					row.BuyFee = cast.ToFloat64(strings.Trim(buyFeeStrs[1], " "))
				}

				//平今手续费
				row.SellTodayFeeStr = strings.Trim(str_sub.Eq(9).Text(), " ")
				sellTodayFeeStrs := strings.Split(row.SellTodayFeeStr, "≈")
				if len(sellTodayFeeStrs) <= 1 {
					row.SellTodayFee = cast.ToFloat64(row.SellTodayFeeStr)
				} else {
					row.SellTodayFee = cast.ToFloat64(strings.Trim(sellTodayFeeStrs[1], " "))
				}

				//平昨手续费
				row.SellYesterdayFeeStr = strings.Trim(str_sub.Eq(10).Text(), " ")
				sellYesterdayFeeStrs := strings.Split(row.SellYesterdayFeeStr, "≈")
				if len(sellYesterdayFeeStrs) <= 1 {
					row.SellYesterdayFee = cast.ToFloat64(row.SellYesterdayFeeStr)
				} else {
					row.SellYesterdayFee = cast.ToFloat64(strings.Trim(sellYesterdayFeeStrs[1], " "))
				}

				//每跳毛利
				row.OneDotProfile = cast.ToFloat64(strings.Trim(str_sub.Eq(11).Text(), " "))
				//手续费(开+平今)
				row.BuySellFee = cast.ToFloat64(strings.Trim(str_sub.Eq(12).Text(), " "))
				//每跳净利/元
				row.OneDotNetProfile = cast.ToFloat64(strings.Trim(str_sub.Eq(13).Text(), " "))
				rows = append(rows, row)

				data_size++
			}
		})

		dataes[i] = &rows
	})

	return header, &exchange_names, &dataes, data_size, nil
}

func GetColumnName(c, row_index int) string {
	letter, _ := ConvertNumToChar(c)
	return fmt.Sprintf("%s%d", letter, row_index)
}

func toExcel(header string, exchange_names *[]string, dataes *map[int]*[]*Row, data_size int) error {
	f := excelize.NewFile()
	defer f.Close()

	red_font, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			WrapText:   true,
			Horizontal: "center",
			Vertical:   "center",
		},
		Font: &excelize.Font{
			Color:  "FF0000",
			Family: "微软雅黑",
			Bold:   true,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})

	//
	default_style, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			WrapText:   true,
			Horizontal: "center",
			Vertical:   "center",
		},
		Font: &excelize.Font{
			Family: "微软雅黑",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	//
	bold_font, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			WrapText:   true,
			Horizontal: "center",
			Vertical:   "center",
		},
		Font: &excelize.Font{
			Family: "微软雅黑",
			Bold:   true,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	//
	sheet_name := "Sheet1"

	// 创建一个工作表
	index, err := f.NewSheet(sheet_name)
	if err != nil {
		return err
	}

	//行高
	rc := 2
	f.SetRowHeight(sheet_name, 1, 8)
	for ; rc <= data_size+2; rc++ {
		f.SetRowHeight(sheet_name, rc, 28)
	}

	//列宽
	f.SetColWidth(sheet_name, "A", "A", 1.5)
	f.SetColWidth(sheet_name, "B", "K", 18)

	if err := f.SetCellStyle(sheet_name, "B2", fmt.Sprintf("K%d", data_size+2), default_style); err != nil {
		return err
	}

	if err := f.SetCellStyle(sheet_name, "B2", "K2", bold_font); err != nil {
		return err
	}

	////
	var row_index = 2
	c := 0
	column_titles := []string{"", "品种代码", "现价", "买开保证金",
		"卖开保证金", "开仓手续费", "平今手续费", "平昨手续费", "每跳毛利", "手续费(开+平今)", "每跳净利",
	}
	for ; c < 11; c++ {
		f.SetCellValue(sheet_name, GetColumnName(c+1, row_index), column_titles[c])
	}

	for _, rows := range *dataes {
		for _, row := range *rows {
			row_index += 1
			name_set_red := false
			//
			c := 1
			f.SetCellValue(sheet_name, GetColumnName(c, row_index), "")
			c++
			//
			f.SetCellValue(sheet_name, GetColumnName(c, row_index), row.NameCode)
			c++
			//
			f.SetCellValue(sheet_name, GetColumnName(c, row_index), row.Price)
			c++
			//
			f.SetCellValue(sheet_name, GetColumnName(c, row_index), row.OpenBuyMargin)
			c++
			//
			f.SetCellValue(sheet_name, GetColumnName(c, row_index), row.OpenSellMargin)
			c++
			//
			f.SetCellValue(sheet_name, GetColumnName(c, row_index), row.BuyFee)
			if row.BuyFee > 30 {
				f.SetCellStyle(sheet_name, GetColumnName(c, row_index), GetColumnName(c, row_index), red_font)
				name_set_red = true
			}
			c++
			//
			f.SetCellValue(sheet_name, GetColumnName(c, row_index), row.SellTodayFee)
			if row.SellTodayFee > 30 {
				f.SetCellStyle(sheet_name, GetColumnName(c, row_index), GetColumnName(c, row_index), red_font)
				name_set_red = true
			}
			c++
			//
			f.SetCellValue(sheet_name, GetColumnName(c, row_index), row.SellYesterdayFee)
			if row.SellYesterdayFee > 30 {
				f.SetCellStyle(sheet_name, GetColumnName(c, row_index), GetColumnName(c, row_index), red_font)
				name_set_red = true
			}
			c++
			//
			f.SetCellValue(sheet_name, GetColumnName(c, row_index), row.OneDotProfile)
			c++
			//
			f.SetCellValue(sheet_name, GetColumnName(c, row_index), row.BuySellFee)
			if row.BuySellFee > 20 {
				f.SetCellStyle(sheet_name, GetColumnName(c, row_index), GetColumnName(c, row_index), red_font)
				name_set_red = true
			}
			c++

			//
			f.SetCellValue(sheet_name, GetColumnName(c, row_index), row.OneDotNetProfile)
			if row.OneDotNetProfile < 0 {
				f.SetCellStyle(sheet_name, GetColumnName(c, row_index), GetColumnName(c, row_index), red_font)
			}

			if name_set_red {
				f.SetCellStyle(sheet_name, GetColumnName(2, row_index), GetColumnName(2, row_index), red_font)
			}
		}
	}

	//
	f.SetActiveSheet(index)

	// 根据指定路径保存文件
	header = strings.ReplaceAll(header, ":", "")
	if err := f.SaveAs(header + ".xlsx"); err != nil {
		return err
	}

	return nil
}

func main() {
	header, exchange_names, dataes, data_size, err := getContent()
	if err != nil {
		panic(err.Error())
	}
	if err := toExcel(header, exchange_names, dataes, data_size); err != nil {
		panic(err.Error())
	}
}
