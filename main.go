package Proses

import (
	"dapofiles" // "github.com/vibrill/dapofiles"
	"fmt"
	"io"
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
	if err == nil {
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

}

func proSis(path string) {
	//fmt.Println("cleaning file : " + path)
	f, err := excelize.OpenFile(path)
	if err == nil {
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
		f.SetCellValue(sheetName, "Y7", "NAMA"+"_AYAH")
		f.SetCellValue(sheetName, "Z7", "TL"+"_AYAH")
		f.SetCellValue(sheetName, "AA7", "PEND"+"_AYAH")
		f.SetCellValue(sheetName, "AB7", "KERJA"+"_AYAH")
		f.SetCellValue(sheetName, "AC7", "HASIL"+"_AYAH")
		f.SetCellValue(sheetName, "AD7", "NIK"+"_AYAH")
		f.SetCellValue(sheetName, "AE7", "NAMA"+"_IBU")
		f.SetCellValue(sheetName, "AF7", "TL"+"_IBU")
		f.SetCellValue(sheetName, "AG7", "PEND"+"_IBU")
		f.SetCellValue(sheetName, "AH7", "KERJA"+"_IBU")
		f.SetCellValue(sheetName, "AI7", "HASIL"+"_IBU")
		f.SetCellValue(sheetName, "AJ7", "NIK"+"_IBU")
		f.SetCellValue(sheetName, "AK7", "NAMA"+"_WALI")
		f.SetCellValue(sheetName, "AL7", "TL"+"_WALI")
		f.SetCellValue(sheetName, "AM7", "PEND"+"_WALI")
		f.SetCellValue(sheetName, "AN7", "KERJA"+"_WALI")
		f.SetCellValue(sheetName, "AO7", "HASIL"+"_WALI")
		f.SetCellValue(sheetName, "AP7", "NIK"+"_WALI")
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

}

func Proses() {
	//ini program utama
	u, _ := os.UserHomeDir()
	listSGT[0], listSGT[1], listSGT[2] = dapofiles.Cek()
	if listSGT[0] != "empty" || listSGT[1] != "empty" || listSGT[2] != "empty" {
		_ = os.Mkdir("/Temp", 0755)
	}

	//mendapatkan path file dapodik di folder download
	for _, x := range listSGT {
		if x != "empty" {
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
			original, err1 := os.Open(ua)
			if err1 == nil {
				//make file
				new, err2 := os.Create(`/Temp/` + x)
				if err2 == nil {
					//copy file
					_, err = io.Copy(new, original) // _ = bytesWritten
					//fmt.Printf("Bytes Written: %d\n", bytesWritten)
					if err != nil {
						log.Fatal(err, " error diufhb")
					}
				}
				defer new.Close()
			}
			defer original.Close()
		}
	}
	if listSGT[0] != "empty" {
		proSis(`/Temp/` + listSGT[0])
	}
	if listSGT[1] != "empty" {
		proGT(`/Temp/` + listSGT[1])
	}
	if listSGT[2] != "empty" {
		proGT(`/Temp/` + listSGT[2])
	}
}
