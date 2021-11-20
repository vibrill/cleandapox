package Proses

import (
	"dapofiles" // "github.com/vibrill/dapofiles"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

var (
	listSGT [3]string
)

//column generator
func genCoord() []string {
	alfabet1 := []string{"", "A", "B"}
	alfabet2 := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	var hasil []string
	for _, x := range alfabet1 {
		for _, y := range alfabet2 {
			hasil = append(hasil, (x + y))
		}
	}
	return hasil
}

func proGT(path string) {
	//fmt.Println("cleaning file : " + path)
	f, err := excelize.OpenFile(path)
	if err != nil {
		log.Fatal("ERROR", err.Error())
	}
	f.SetActiveSheet(0)
	sheetName := f.GetSheetName(0)
	kolom := genCoord()

	for _, x := range kolom {
		nilai, err := f.GetCellValue(sheetName, x+"5")
		if err == nil {
			nilai = strings.ReplaceAll(nilai, " ", "_")
			nilai = strings.ToUpper(nilai)
			f.SetCellValue(sheetName, x+"5", nilai)
			//fmt.Println(x+"5 ", nilai)
		}
	}
	f.SetCellValue(sheetName, "AZ5", "YA")
	f.RemoveRow(sheetName, 1)
	f.RemoveRow(sheetName, 1)
	f.RemoveRow(sheetName, 1)
	f.RemoveRow(sheetName, 1)
	f.Save()
	//fmt.Println("file GTK telah dirapikan")
}

func proSis(path string) {
	//fmt.Println("cleaning file : " + path)
	f, err := excelize.OpenFile(path)
	if err != nil {
		log.Fatal("ERROR", err.Error())
	}
	f.SetActiveSheet(0)
	sheetName := f.GetSheetName(0)

	//insert row
	f.UnmergeCell(sheetName, "A1", "BN6")
	f.InsertRow(sheetName, 7)
	//fmt.Println("row 7 inserted")

	//set column name
	kolom := genCoord()
	for _, x := range kolom {
		nilai, err := f.GetCellValue(sheetName, x+"5")
		if err == nil {
			nilai = strings.ReplaceAll(nilai, " ", "_")
			nilai = strings.ToUpper(nilai)
			f.SetCellValue(sheetName, x+"7", nilai)
			//fmt.Println(x+"7 ", nilai)
		}
	}
	f.SetCellValue(sheetName, "Y7", "Y6"+"_AYAH")
	f.SetCellValue(sheetName, "Z7", "Z6"+"_AYAH")
	f.SetCellValue(sheetName, "AA7", "AA6"+"_AYAH")
	f.SetCellValue(sheetName, "AB7", "AB6"+"_AYAH")
	f.SetCellValue(sheetName, "AC7", "AC6"+"_AYAH")
	f.SetCellValue(sheetName, "AD7", "AD6"+"_AYAH")
	f.SetCellValue(sheetName, "AE7", "AE6"+"_IBU")
	f.SetCellValue(sheetName, "AF7", "AF6"+"_IBU")
	f.SetCellValue(sheetName, "AG7", "AG6"+"_IBU")
	f.SetCellValue(sheetName, "AH7", "AH6"+"_IBU")
	f.SetCellValue(sheetName, "AI7", "AI6"+"_IBU")
	f.SetCellValue(sheetName, "AJ7", "AJ6"+"_IBU")
	f.SetCellValue(sheetName, "AK7", "AK6"+"_WALI")
	f.SetCellValue(sheetName, "AL7", "AL6"+"_WALI")
	f.SetCellValue(sheetName, "AM7", "AM6"+"_WALI")
	f.SetCellValue(sheetName, "AN7", "AN6"+"_WALI")
	f.SetCellValue(sheetName, "AO7", "AO6"+"_WALI")
	f.SetCellValue(sheetName, "AP7", "AP6"+"_WALI")
	//fmt.Println("row 7 named")

	f.RemoveRow(sheetName, 1)
	f.RemoveRow(sheetName, 1)
	f.RemoveRow(sheetName, 1)
	f.RemoveRow(sheetName, 1)
	f.RemoveRow(sheetName, 1)
	f.RemoveRow(sheetName, 1)

	f.Save()
	//fmt.Println("file siswa telah dirapikan")
}

func Proses() {
	//ini program utama
	u, _ := os.UserHomeDir()
	listSGT[0], listSGT[1], listSGT[2] = dapofiles.Cek()
	//mendapatkan path desktop
	_, err := ioutil.ReadDir(string(u) + `/Desktop`)
	ud := u + `/Desktop`
	if err != nil {
		_, err = ioutil.ReadDir(`E:/Desktop`)
		ud = `E:/Desktop`
		if err != nil {
			_, err = ioutil.ReadDir(`D:/Desktop`)
			ud = `D:/Desktop`
			if err != nil {
				fmt.Println(err)
			}
		}
	}

	//membuat directory di desktop untuk mengkopi file dapodik
	if err == nil {
		err = os.Mkdir(ud+"/DapoSniff", 0755)
		if err != nil {
			fmt.Println("Folder ", ud+"/DapoSniff telah siap")
		} else {
			fmt.Println("Membuat Folder " + u + "/DapoSniff")
		}
	}

	//mendapatkan path file dapodik di folder download
	for _, x := range listSGT {
		_, err := os.Open(u + "/Downloads/" + x)
		ua := u + "/Downloads/" + x
		if err != nil {
			_, err = os.Open("D:/Downloads/" + x)
			ua = "D:/Downloads/" + x
			if err != nil {
				_, err = os.Open("E:/Downloads/" + x)
				ua = "E:/Downloads/" + x
				if err != nil {
					fmt.Println(err)
				}
			}
		}

		//read file
		original, err := os.Open(ua)
		if err != nil {
			log.Fatal(err)
		}
		defer original.Close()

		//make file
		new, err := os.Create(ud + `/DapoSniff/` + x)
		if err != nil {
			log.Fatal(err)
		}
		defer new.Close()

		//copy file
		_, err = io.Copy(new, original) // _ = bytesWritten
		if err != nil {
			log.Fatal(err)
		}
		//fmt.Printf("Bytes Written: %d\n", bytesWritten)
	}
	proSis(ud + `/DapoSniff/` + listSGT[0])
	proGT(ud + `/DapoSniff/` + listSGT[1])
	proGT(ud + `/DapoSniff/` + listSGT[2])
}
