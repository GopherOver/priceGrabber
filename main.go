package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"runtime"
	"strconv"
	"sync"
	"time"

	"golang.org/x/text/encoding/charmap"

	"github.com/PuerkitoBio/goquery"
	"github.com/go-toast/toast"
	"github.com/tealeg/xlsx"
)

// Company struct
type Company struct {
	Title    string         `json:"title"`
	Selector string         `json:"selector"`
	Attr     string         `json:"attr"`
	Color    string         `json:"color"`
	Links    []string       `json:"links"`
	Price    []int          `json:"price"`
	Style    *xlsx.Style    `json:"-"`
	Wg       sync.WaitGroup `json:"-"`
}

// Object model for json file
type Object struct {
	Companies []*Company     `json:"companies"`
	Models    []string       `json:"models"`
	Wg        sync.WaitGroup `json:"-"`
}

// Warning ...
var Warning bool

// Start ...
func (o *Object) Start() {
	o.Wg.Add(len(o.Companies))
	for _, c := range o.Companies {
		go func(c *Company) {
			c.Parse(o.Models)
			o.Wg.Done()
		}(c)
	}
	o.Wg.Wait()
}

// Parse ...
func (c *Company) Parse(m []string) {
	c.Wg.Add(len(c.Links))

	for key, link := range c.Links {
		go func(c *Company, link string, key int) {
			repeats := 30
			for {
				if link != "" {
					if repeats < 1 {
						break
					}
					log.Println(c.Title, m[key], " - осталось попыток: ", repeats)
					repeats--
					if doc, err := goquery.NewDocument(link); err == nil {
						if price := doc.Find(c.Selector).AttrOr(c.Attr, ""); price != "" {
							c.Price[key], _ = strconv.Atoi(price)
							fmt.Println(c.Title, m[key], price)
							break
						}
					}
				} else {
					break
				}
			}
			c.Wg.Done()
		}(c, link, key)

	}

	c.Wg.Wait()
}

func main() {

	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell

	data, err := ioutil.ReadFile("config.json")

	if err != nil {
		fmt.Println("Read file error: ", err)
		return
	}

	var c Object

	if err = json.Unmarshal(data, &c); err != nil {
		fmt.Println("Unmarshal error: ", err)
		return
	}

	for _, s := range c.Companies {
		if len(s.Price) == 0 {
			s.Price = make([]int, len(s.Links))
		}
		s.Style = &xlsx.Style{
			Fill:      *xlsx.NewFill("solid", s.Color, ""),
			Alignment: xlsx.Alignment{Horizontal: "center", Vertical: "center"},
			Border:    *xlsx.NewBorder("thin", "thin", "thin", "thin"),
		}
	}

	c.Start()

	path := "./Цены.xlsx"

	xlFile, err := xlsx.OpenFile(path)

	if err != nil {
		log.Fatalln(err.Error())
		return
	}

	imarketPrice := Company{
		Title: "iMarket",
		Style: &xlsx.Style{
			Fill:      *xlsx.NewFill("solid", "CEFF00", ""),
			Alignment: xlsx.Alignment{Horizontal: "center", Vertical: "center"},
			Border:    *xlsx.NewBorder("thin", "thin", "thin", "thin"),
		},
		Price: make([]int, len(c.Companies[0].Links)),
	}

	// magic
	for _, sheet := range xlFile.Sheets {
		for kr, row := range sheet.Rows {
			if kr > 1 {
				for kc, cell := range row.Cells {
					if kc == 1 {
						val, notInt := cell.Int()

						if notInt != nil {
							fmt.Println(notInt.Error())
							return
						}
						imarketPrice.Price[kr-2] = val
					}
				}
			}
		}
	}

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Цены")

	if err != nil {
		log.Fatalln(err.Error())
		return
	}

	sheet.SetColWidth(0, 0, 30)
	xlsx.SetDefaultFont(12, "Verdana")

	border := *xlsx.NewBorder("thin", "thin", "thin", "thin")

	warning := &xlsx.Style{
		Fill:      *xlsx.NewFill("solid", "FF0000", ""),
		Alignment: xlsx.Alignment{Horizontal: "center", Vertical: "center"},
		Border:    *xlsx.NewBorder("thin", "thin", "thin", "thin"),
	}

	row = sheet.AddRow()
	row.SetHeight(25)

	cell = row.AddCell()
	cell.Merge(5, 0)
	cell.SetValue("Актуально на: " + time.Now().Format("02-01-2006"))
	cell.SetStyle(&xlsx.Style{
		Alignment: xlsx.Alignment{Vertical: "center", Horizontal: "center"},
		Border:    border,
	})

	for i := 0; i < len(c.Models); i++ {

		if i == 0 {
			row = sheet.AddRow()
			row.SetHeight(25)
			cell = row.AddCell()
			cell.SetStyle(&xlsx.Style{
				Border: border,
			})

			cell = row.AddCell()
			cell.SetString(imarketPrice.Title)
			cell.SetStyle(imarketPrice.Style)

			for _, v := range c.Companies {
				cell = row.AddCell()
				cell.SetString(v.Title)
				cell.SetStyle(v.Style)
			}
		}
		row = sheet.AddRow()
		row.SetHeight(25)
		cell = row.AddCell()
		cell.SetString(c.Models[i])
		cell.SetStyle(&xlsx.Style{
			Font:      xlsx.Font{Bold: true},
			Alignment: xlsx.Alignment{Vertical: "center"},
			Border:    border,
		})

		cell = row.AddCell()
		cell.SetInt(imarketPrice.Price[i])
		cell.SetStyle(imarketPrice.Style)

		for _, v := range c.Companies {
			cell = row.AddCell()
			cell.SetInt(v.Price[i])
			cell.SetStyle(v.Style)
			if v.Price[i] > 0 && imarketPrice.Price[i] > v.Price[i] {
				Warning = true
				cell.SetStyle(warning)
			} else {
				cell.SetStyle(v.Style)
			}
		}
	}

	file.Save(path)

	println("All Done!")

	if Warning && runtime.GOOS == "windows" {
		b := charmap.Windows1251.NewEncoder()
		title, err := b.Bytes([]byte("Граббер Цен"))
		message, err := b.Bytes([]byte("Есть изменения!"))

		if err != nil {
			log.Fatalln(err.Error())
		}

		notification := toast.Notification{
			AppID:               "",
			Title:               string(title),
			Message:             string(message),
			ActivationArguments: "",
			Audio:               toast.SMS,
			Duration:            toast.Long,
		}

		err = notification.Push()

		if err != nil {
			log.Fatalln(err.Error())
		}
	}

}
