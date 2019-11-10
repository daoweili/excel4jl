package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"math"
	"os"
	"strconv"
	"strings"
)

var usage = func() {
	fmt.Println("USAGE: excel4jl command [arguments] ...")
	fmt.Println("example: excel4jl a.xlsx weight.xlsx ")
}

var heads = []string{"内部订单号", "订单类型", "线上订单号", "店铺", "买家账号",
	"买家留言", "卖家备注", "发货日期", "收件人姓名", "省", "市", "区县", "快递公司",
	"快递单号", "商品编码", "款式编码", "实发数量", "单价", "价款", "重量", "运费支出", "小计"}

var cellNames = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "G", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V"}

const firstWeight = 3
const firstWeightPrice = 5.2

var renewalFeePrice = map[string]int{
	"广东":  1,
	"江苏":  4,
	"浙江":  4,
	"上海":  4,
	"广西":  4,
	"江西":  4,
	"福建":  4,
	"湖南":  4,
	"湖北":  4,
	"安徽":  4,
	"河南":  5,
	"河北":  5,
	"海南":  5,
	"山东":  5,
	"北京":  5,
	"天津":  5,
	"四川":  6,
	"重庆":  6,
	"贵州":  6,
	"云南":  6,
	"山西":  6,
	"陕西":  6,
	"黑龙江": 6,
	"辽宁":  6,
	"吉林":  6,
	"甘肃":  10,
	"青海":  10,
	"宁夏":  10,
	"内蒙古": 12,
	"西藏":  12,
	"新疆":  12,
}

type JlExcel struct {
	innerOrder     string  //内部订单号
	orderType      string  //订单类型
	onlineOrder    string  //线上订单号
	shop           string  //店铺
	buyer          string  //买家账户
	buyerMsg       string  //买家留言
	sellerRemark   string  //卖家备注
	shipTime       string  //发货日期
	recipient      string  //收件人
	province       string  //省
	city           string  //市
	district       string  //区县
	courierCompany string  //快递公司
	trackingNo     string  //快递单号
	commodityCode  string  //商品编码
	styleCode      string  //款式编码
	realWage       int     //实发数量
	unitPrice      float64 //单价
	price          float64 //价款
	weight         float64 //重量
	freightExpense float64 //运费支出
	subTotal       float64 //小计
}

func readExcel(file *excelize.File) map[string][]*JlExcel {
	resultMap := make(map[string][]*JlExcel, 256)
	for index, name := range file.GetSheetMap() {
		fmt.Printf("handler index : %d, name : %s \n", index, name)
		rows, err := file.GetRows(name)
		jlExcels := make([]*JlExcel, 0)
		if err != nil {
			fmt.Println("get rows err", err.Error())
			continue
		}
		for index, row := range rows {
			//首行标题不处理
			if index == 0 {
				continue
			}
			realWage, _ := strconv.Atoi(row[16])
			unitPrice, _ := strconv.ParseFloat(row[17], 64)
			price, _ := strconv.ParseFloat(row[18], 64)
			weight, _ := strconv.ParseFloat(row[19], 64)
			freightExpense, _ := strconv.ParseFloat(row[20], 64)
			subTotal, _ := strconv.ParseFloat(row[21], 64)
			jlExcels = append(jlExcels, &JlExcel{
				innerOrder:     row[0],         //内部订单号
				orderType:      row[1],         //订单类型
				onlineOrder:    row[2],         //线上订单号
				shop:           row[3],         //店铺
				buyer:          row[4],         //买家账户
				buyerMsg:       row[5],         //买家留言
				sellerRemark:   row[6],         //卖家备注
				shipTime:       row[7],         //发货日期
				recipient:      row[8],         //收件人
				province:       row[9],         //省
				city:           row[10],        //市
				district:       row[11],        //区县
				courierCompany: row[12],        //快递公司
				trackingNo:     row[13],        //快递单号
				commodityCode:  row[14],        //商品编码
				styleCode:      row[15],        //款式编码
				realWage:       realWage,       //实发数量
				unitPrice:      unitPrice,      //单价
				price:          price,          //价款
				weight:         weight,         //重量
				freightExpense: freightExpense, //运费支出
				subTotal:       subTotal,       //小计
			})
		}
		resultMap[name] = jlExcels
	}
	return resultMap
}

func printErr(err error) {
	if err == nil {
		return
	}
	fmt.Println(err)
}

func setCell(val interface{}, index, sheet string, file *excelize.File) {
	switch val.(type) {
	case int:
		if v, ok := val.(int); ok {
			err := file.SetCellInt(sheet, index, v)
			printErr(err)
		}
	case string:
		if v, ok := val.(string); ok {
			err := file.SetCellStr(sheet, index, v)
			printErr(err)
		}
	case float64:
		if v, ok := val.(float64); ok {
			err := file.SetCellFloat(sheet, index, v, 2, 64)
			printErr(err)
		}
	default:
		err := file.SetCellDefault(sheet, index, val.(string))
		printErr(err)
	}
}

func setStrRow(str string, index, sheet string, file *excelize.File) {
	for i, n := range cellNames {
		if strconv.Itoa(i+1) == index {

			setCell(str, n+strconv.Itoa(1), sheet, file)
		}
	}
}

func setJlExcelRow(jlExcel *JlExcel, index, sheet string, file *excelize.File) {
	setCell(jlExcel.innerOrder, "A"+index, sheet, file)
	setCell(jlExcel.orderType, "B"+index, sheet, file)
	setCell(jlExcel.onlineOrder, "C"+index, sheet, file)
	setCell(jlExcel.shop, "D"+index, sheet, file)
	setCell(jlExcel.buyer, "E"+index, sheet, file)
	setCell(jlExcel.buyerMsg, "F"+index, sheet, file)
	setCell(jlExcel.sellerRemark, "G"+index, sheet, file)
	setCell(jlExcel.shipTime, "H"+index, sheet, file)
	setCell(jlExcel.recipient, "I"+index, sheet, file)
	setCell(jlExcel.province, "J"+index, sheet, file)
	setCell(jlExcel.city, "K"+index, sheet, file)
	setCell(jlExcel.district, "L"+index, sheet, file)
	setCell(jlExcel.courierCompany, "M"+index, sheet, file)
	setCell(jlExcel.trackingNo, "N"+index, sheet, file)
	setCell(jlExcel.commodityCode, "O"+index, sheet, file)
	setCell(jlExcel.styleCode, "P"+index, sheet, file)
	setCell(jlExcel.realWage, "Q"+index, sheet, file)
	setCell(jlExcel.unitPrice, "R"+index, sheet, file)
	setCell(jlExcel.price, "S"+index, sheet, file)
	setCell(jlExcel.weight, "T"+index, sheet, file)
	setCell(jlExcel.freightExpense, "U"+index, sheet, file)
	setCell(jlExcel.subTotal, "V"+index, sheet, file)

}

func IsFileExist(fileName string) bool {
	if _, err := os.Stat(fileName); os.IsNotExist(err) {
		return false
	}
	return true
}

/**
	读取重量表
**/
func readWeight(fileName string) map[string]float64 {
	resultMap := make(map[string]float64, 256)
	file, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		return resultMap
	}
	firstSheetName := file.GetSheetName(1)
	rows, err := file.GetRows(firstSheetName)
	printErr(err)
	for _, row := range rows {
		model := strings.TrimSpace(row[0])
		model = strings.ReplaceAll(model, "【", "")
		model = strings.ReplaceAll(model, "】", "")
		resultMap[model], err = strconv.ParseFloat(row[1], 64)
		printErr(err)
	}
	return resultMap
}

func checkCalcWeight(weightMap map[string]float64) bool {
	if len(weightMap) > 0 {
		return true
	}
	return false
}

func main() {
	args := os.Args[1:]
	fmt.Println("args: ", args)
	if len(args) < 1 {
		usage()
		os.Exit(0)
	}
	var file = args[0]
	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return
	}
	//读取计算的excel内容
	resultMap := readExcel(f)
	//读取型号重量表
	var weightPath string
	if len(args) > 1 {
		weightPath = args[1]
	}
	var weightMap map[string]float64
	if weightPath != "" {
		weightMap = readWeight(weightPath)
	}
	for k, sheetContext := range resultMap {
		fmt.Println("sheet name : ", k)
		for i := 0; i < len(sheetContext)-1; i++ {
			jlExcel := sheetContext[i]
			if checkCalcWeight(weightMap) {
				weight := weightMap[jlExcel.styleCode]
				jlExcel.weight = float64(jlExcel.realWage) * weight
			}
			for j := 0; j < len(sheetContext)-1; j++ {
				cmJlExcel := sheetContext[j]
				if i == j {
					continue
				}
				if jlExcel.trackingNo == cmJlExcel.trackingNo {
					if checkCalcWeight(weightMap) {
						weight := weightMap[cmJlExcel.styleCode]
						cmJlExcel.weight = float64(cmJlExcel.realWage) * weight
					}
					jlExcel.weight += cmJlExcel.weight
					cmJlExcel.weight = 0
					if jlExcel.freightExpense == 0 {
						jlExcel.freightExpense = cmJlExcel.freightExpense
					}
					cmJlExcel.freightExpense = 0
				}
			}
			if jlExcel.weight > firstWeight {
				other := jlExcel.weight - firstWeight
				for province, rfp := range renewalFeePrice {
					if strings.HasPrefix(jlExcel.province, province) {
						jlExcel.freightExpense = firstWeightPrice + float64(rfp)*math.Ceil(other)
					}
				}
			} else {
				jlExcel.freightExpense = firstWeightPrice
			}
		}
	}

	f = excelize.NewFile()

	for sheet, sheetContext := range resultMap {
		f.NewSheet(sheet)
		for k, j := range heads {
			setStrRow(j, strconv.Itoa(k+1), sheet, f)
		}
		for i := 0; i < len(sheetContext); i++ {
			jlExcel := sheetContext[i]
			s := strconv.Itoa(i + 2)
			setJlExcelRow(jlExcel, s, sheet, f)
		}
	}
	f.SetActiveSheet(0)
	paths := strings.Split(file, ".")
	newPath := strings.TrimSpace(paths[0]) + "_new." + paths[1]
	if IsFileExist(newPath) {
		err = os.Remove(newPath)
		if err != nil {
			fmt.Println(err)
		}
	}
	err = f.SaveAs(newPath)
	if err != nil {
		fmt.Println(err)
		os.Exit(0)
	}
	fmt.Println("execute success : ", newPath)
}
