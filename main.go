package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"strings"
)

type Item struct {
	Name        string
	Spec        string
	Price       string
	Count       int
	BuyerNote   string
	SellerNote  string
	Address     string
	Province    string
	City        string
	District    string
	PhoneNumber string
}

func processFile(filePath string) {
	// 打开Excel文件
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println("无法打开文件:", err)
		return
	}

	// 获取第一个工作表
	sheet := f.GetSheetName(0)

	// 读取数据
	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println("读取数据出错:", err)
		return
	}

	// 存储统计结果
	itemMap := make(map[string]map[string]*Item)
	// 存储备注订单
	noteOrders := make([]Item, 0)

	// 遍历数据,从第二行开始
	for i, row := range rows {
		if i == 0 {
			// 跳过表头行
			continue
		}

		if len(row) >= 32 {
			// 获取商品信息
			itemName := row[29]
			itemSpec := strings.ReplaceAll(row[30], " ", "")
			itemPrice := row[31]
			buyerNote := row[10]
			sellerNote := row[11]
			address := row[5]
			province := row[6]
			city := row[7]
			district := row[8]
			phoneNumber := row[9]

			// 生成唯一键
			key := itemName

			// 统计数量
			if !(buyerNote != "" || sellerNote != "") {
				if itemSpecs, ok := itemMap[key]; ok {
					if item, ok := itemSpecs[itemSpec]; ok {
						item.Count++
					} else {
						itemSpecs[itemSpec] = &Item{
							Name:  itemName,
							Spec:  itemSpec,
							Price: itemPrice,
							Count: 1,
						}
					}
				} else {
					itemMap[key] = make(map[string]*Item)
					itemMap[key][itemSpec] = &Item{
						Name:  itemName,
						Spec:  itemSpec,
						Price: itemPrice,
						Count: 1,
					}
				}
			}

			// 处理备注订单
			if buyerNote != "" || sellerNote != "" {
				noteOrder := Item{
					Name:        itemName,
					Spec:        itemSpec,
					Price:       itemPrice,
					BuyerNote:   buyerNote,
					SellerNote:  sellerNote,
					Address:     address,
					Province:    province,
					City:        city,
					District:    district,
					PhoneNumber: phoneNumber,
				}
				noteOrders = append(noteOrders, noteOrder)
			}
		}
	}

	// 生成输出字符串
	var output strings.Builder
	count := 1
	for key, itemSpecs := range itemMap {
		output.WriteString(fmt.Sprintf("%d.%s 商品价格:%s\n\n", count, key, itemSpecs[getFirstSpec(itemSpecs)].Price))
		for spec, item := range itemSpecs {
			output.WriteString(fmt.Sprintf("SKU:%s 数量:%d\n\n", spec, item.Count))
		}
		output.WriteString("\n")
		count++
	}

	if len(noteOrders) > 0 {
		output.WriteString("额外的备注订单(未统计在上方):\n\n")
		for i, order := range noteOrders {
			if order.BuyerNote != "" || order.SellerNote != "" {
				output.WriteString(fmt.Sprintf("%d.%s 原始SKU:%s 单价:%s", i+1, order.Name, order.Spec, order.Price))
				if order.BuyerNote != "" {
					output.WriteString(fmt.Sprintf(" 买家备注:%s", order.BuyerNote))
				}
				if order.SellerNote != "" {
					output.WriteString(fmt.Sprintf(" 商家备注:%s", order.SellerNote))
				}
				output.WriteString(fmt.Sprintf(" 收件人地址:%s 省:%s 市:%s 区:%s 收件人手机:%s\n\n", order.Address, order.Province, order.City, order.District, order.PhoneNumber))
			}
		}
	}

	// 获取可执行文件所在目录
	execDir, err := os.Executable()
	if err != nil {
		fmt.Println("获取可执行文件路径出错:", err)
		return
	}
	execDir = filepath.Dir(execDir)

	// 获取文件名(不包括扩展名)
	fileNameWithoutExt := strings.TrimSuffix(filepath.Base(filePath), filepath.Ext(filePath))

	// 生成txt文件路径
	txtFilePath := filepath.Join(execDir, fileNameWithoutExt+".txt")

	// 将输出字符串写入txt文件
	err = os.WriteFile(txtFilePath, []byte(output.String()), 0644)
	if err != nil {
		fmt.Println("生成txt文件出错:", err)
		return
	}

	fmt.Println("处理完成,已在程序所在目录下生成同名的txt文件:", txtFilePath)
}

func getFirstSpec(itemSpecs map[string]*Item) string {
	for spec := range itemSpecs {
		return spec
	}
	return ""
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("请将Excel文件拖放到程序上")
		fmt.Println("按任意键退出...")
		fmt.Scanln()
		return
	}

	filePath := os.Args[1]
	processFile(filePath)

	fmt.Println("按任意键退出...")
	fmt.Scanln()
}
