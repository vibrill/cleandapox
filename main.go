package main

import (
	"dapofiles" // "github.com/vibrill/dapofiles"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"os"

	excelize "github.com/360EntSecGroup-Skylar/excelize/v2"
)

var (
	listSGT [3]string
)

func proses(file string) {
	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
	}

	f.save()
}

func main() {
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
			fmt.Println(err)
		} else {
			fmt.Println("Folder " + u + "/DapoSniff created")
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
		bytesWritten, err := io.Copy(new, original)
		if err != nil {
			log.Fatal(err)
		}
		fmt.Printf("Bytes Written: %d\n", bytesWritten)
	}
	proses(ud + `/DapoSniff/` + listSGT[0]) //siswa

}
