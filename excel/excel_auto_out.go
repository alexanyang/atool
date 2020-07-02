package excel

import (
	"fmt"
	excelize "github.com/360EntSecGroup-Skylar/excelize/v2"
	"io"
	"reflect"
)

/**
 * @author anyang
 * Email: 1300378587@qq.com
 * Created Date:2020-07-02 15:28
 */

//传入类型必须是切片
func ExcelOutput(data interface{}, writer io.Writer) (err error) {
	defer func() {
		if err1 := recover(); err1 != nil {
			err = fmt.Errorf(fmt.Sprint(err1))
		}
	}()
	typeOf := reflect.TypeOf(data)
	valueOf := reflect.ValueOf(data)
	kind := typeOf.Kind()
	//判断传入的结构体是否是切片,不是切片报错
	switch kind {
	default:
		return fmt.Errorf(" %v is not a Slice ", data)
	case reflect.Slice:

	}
	vlen := valueOf.Len()
	if vlen <= 0 {
		return
	}
	index := valueOf.Index(0)
	fieldsMap := GetFieldsMap(index.Type())

	//新建excel文件
	newFile := excelize.NewFile()
	//用字段名构建文件第一列
	for k, v := range fieldsMap {
		ax := fmt.Sprintf("%s1", GetRow(k))
		newFile.SetCellValue("Sheet1", ax, v)
	}

	for i := 0; i < vlen; i++ {
		//取出切片元素
		value := valueOf.Index(i)
		numField := value.NumField()
		for j := 0; j < numField; j++ {
			field := value.Field(j)
			axis := fmt.Sprintf("%s%d", GetRow(j), i+2)
			newFile.SetCellValue("Sheet1", axis, field.Interface())
		}
	}
	//是切片,循环切片元素
	err = newFile.Write(writer)
	return
}

func GetRow(index int) string {
	var (
		head    int    = index / 26
		line    int    = index % 26
		headStr string = ""
	)

	switch head {
	case 1:
		headStr = "A"
	case 2:
		headStr = "B"
	case 3:
		headStr = "C"
	case 4:
		headStr = "D"
	case 5:
		headStr = "E"
	case 6:
		headStr = "F"
	case 7:
		headStr = "G"
	case 8:
		headStr = "H"
	case 9:
		headStr = "I"
	case 10:
		headStr = "J"
	case 11:
		headStr = "K"
	case 12:
		headStr = "L"
	case 13:
		headStr = "M"
	case 14:
		headStr = "N"
	case 15:
		headStr = "O"
	case 16:
		headStr = "P"
	case 17:
		headStr = "Q"
	case 18:
		headStr = "R"
	case 19:
		headStr = "S"
	case 20:
		headStr = "T"
	case 21:
		headStr = "U"
	case 22:
		headStr = "V"
	case 23:
		headStr = "W"
	case 24:
		headStr = "X"
	case 25:
		headStr = "Y"
	case 26:
		headStr = "Z"
	}

	switch line {
	case 0:
		return headStr + "A"
	case 1:
		return headStr + "B"
	case 2:
		return headStr + "C"
	case 3:
		return headStr + "D"
	case 4:
		return headStr + "E"
	case 5:
		return headStr + "F"
	case 6:
		return headStr + "G"
	case 7:
		return headStr + "H"
	case 8:
		return headStr + "I"
	case 9:
		return headStr + "J"
	case 10:
		return headStr + "K"
	case 11:
		return headStr + "L"
	case 12:
		return headStr + "M"
	case 13:
		return headStr + "N"
	case 14:
		return headStr + "O"
	case 15:
		return headStr + "P"
	case 16:
		return headStr + "Q"
	case 17:
		return headStr + "R"
	case 18:
		return headStr + "S"
	case 19:
		return headStr + "T"
	case 20:
		return headStr + "U"
	case 21:
		return headStr + "V"
	case 22:
		return headStr + "W"
	case 23:
		return headStr + "X"
	case 24:
		return headStr + "Y"
	case 25:
		return headStr + "Z"
	}
	panic("out of range A-Z")
}

func GetFieldsMap(typeOf reflect.Type) map[int]string {
	var m = make(map[int]string)
	//解析结构体,获取字段名
	numField := typeOf.NumField()
	for i := 0; i < numField; i++ {
		field := typeOf.Field(i)
		excelName := field.Tag.Get("excel")
		if len(excelName) > 0 {
			m[i] = excelName
			continue
		}
		m[i] = field.Name
	}
	return m
}
