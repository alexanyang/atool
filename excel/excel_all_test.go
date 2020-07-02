package excel

import (
	"fmt"
	"os"
	"testing"
	"time"
)

type User struct {
	Name  string  `excel:"姓名"`
	Age   uint    `excel:"年龄"`
	Money float64 `excel:"财富值"`
	Tm    time.Time
	U     Key
}

type Key struct {
	X string
}

func TestExcelOutput(t *testing.T) {
	user1 := User{
		Name:  "赵1",
		Age:   13,
		Money: 100.2,
		U:     Key{},
	}
	user2 := User{
		Name:  "钱1",
		Age:   18,
		Money: 333.3,
		U:     Key{X: "xxx"},
	}
	users := []User{user1, user2}
	fileName := "./testII.xlsx"
	file, err := os.OpenFile(fileName, os.O_CREATE|os.O_TRUNC|os.O_WRONLY, 0666)
	if err != nil {
		file, err = os.Create(fileName)
		if err != nil {
			fmt.Println(err)
			return
		}
	}
	defer file.Close()
	file.Chmod(os.ModeNamedPipe)
	file.SetWriteDeadline(time.Now().Add(5 * time.Second))
	err = ExcelOutput(users, file)
	if err != nil {
		fmt.Println(err)
	}
}
