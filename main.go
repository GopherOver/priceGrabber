package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"runtime"
	"strconv"
	"sync"
	"time"

	"github.com/PuerkitoBio/goquery"
	"github.com/briandowns/spinner"
	"github.com/go-toast/toast"
	"github.com/tealeg/xlsx"
	"golang.org/x/text/encoding/charmap"
)

var (
	models = []string{
		"Apple iPhone X 64Gb",
		"Apple iPhone X 256Gb",
		"Apple iPhone 8 64Gb",
		"Apple iPhone 8 256Gb",
		"Apple iPhone 8 Plus 64Gb",
		"Apple iPhone 8 Plus 256Gb",
		"Apple iPhone 7 32Gb",
		"Apple iPhone 7 128Gb",
		"Apple iPhone 7 256Gb",
		"Apple iPhone 7 Plus 32Gb",
		"Apple iPhone 7 Plus 128Gb",
		"Apple iPhone 7 Plus 256Gb",
		"Apple iPhone 6s 32Gb",
		"Apple iPhone 6s 128Gb",
		"Apple iPhone 6s Plus 32Gb",
		"Apple iPhone 6s Plus 128Gb",
		"Apple iPhone SE 32Gb",
		"Apple iPhone SE 128Gb",
		"Apple AirPods",
		"Apple TV 4 32Gb",
		"Apple TV 4 64Gb",
		"Apple TV 4K 32Gb",
		"Apple TV 4K 64Gb",
		"Apple iPad Pro 10.5 64Gb",
		"Apple iPad Pro 10.5 256Gb",
		"Apple iPad Pro 10.5 512Gb",
		"Apple iPad Pro 10.5 Cellular 64Gb",
		"Apple iPad Pro 10.5 Cellular 256Gb",
		"Apple iPad Pro 10.5 Cellular 512Gb",
	}
	config      *string
	wg          sync.WaitGroup
	warningFlag bool
	filePrice   = "./Цены.xlsx"
	spin        = spinner.New(spinner.CharSets[40], 300*time.Millisecond)
)

type (
	company struct {
		Title     string  `json:"title"`
		Selector  string  `json:"selector"`
		Attribute string  `json:"attribute"`
		Color     string  `json:"color"`
		Models    []model `json:"models"`
	}
	model struct {
		Link  string `json:"link"`
		Price int    `json:"-"`
	}
)

func init() {
	flag.Parse()
	config = flag.String("с", "./config.json", "Конфигурационный файл (по умолчанию: config.json)")
}

func main() {

	start := time.Now().UTC()

	data, err := ioutil.ReadFile(*config)

	if err != nil {
		log.Fatalln("Read file error: ", err)
	}

	var companies = new([]*company)

	if err = json.Unmarshal(data, &companies); err != nil {
		log.Fatalln("Unmarshal error: ", err)
	}

	fmt.Println("Файл конфигурации успешно загружен")

	wg.Add(len(*companies))

	spin.Suffix = "  загружаю цены..."
	// Запускаем спиннер
	spin.Start()

	// Загружаем цены наших компаний
	for _, c := range *companies {
		tm := make([]model, len(models), cap(models))
		copy(tm, c.Models)
		c.Models = tm
		go process(c)
	}

	// Ожидаем завершения всех горутин
	wg.Wait()
	// Создаём новую таблицу
	if err = makeNewSheet(*companies); err != nil {
		log.Fatalln(err.Error())
	}
	// Уведовляем об окончании операции
	notify()
	// Останавливаем спиннер
	spin.Stop()

	fmt.Printf("Файл `%s` обновлён\n", filePrice)
	fmt.Println("Времени потрачено: ", time.Since(start))
}

// process Загружает цены указанной компании
func process(c *company) {
	defer wg.Done()
	for i := range c.Models {
		if c.Models[i].Link != "" {
			if doc, err := goquery.NewDocument(c.Models[i].Link); err == nil {
				if price, ok := doc.Find(c.Selector).Attr(c.Attribute); ok {
					c.Models[i].Price, _ = strconv.Atoi(price)
					// fmt.Printf("%s - %s : %d\n", c.Title, models[i], model.Price)
				} else {
					fmt.Printf("! -> %s - %s - не найден нужный тег: %s %s, возможно, ссылка устарела\n", c.Title, models[i], c.Selector, c.Attribute)
				}
			}
		}
		// } else {
		// 	fmt.Printf("! -> %s - Не задана ссылка для модели %s\n", c.Title, models[i])
		// }
	}
}

// makeNewSheet Создаёт новыую таблицу
func makeNewSheet(c []*company) error {
	var (
		newFile *xlsx.File
		sheet   *xlsx.Sheet
		row     *xlsx.Row
		cell    *xlsx.Cell
	)

	// Пытаемся открыть нашу таблицу (наличие файла обязательно!)
	currentFile, err := xlsx.OpenFile(filePrice)

	if err != nil {
		return err
	}

	// Создаём экземляр нашей компании
	imarketCompany := &company{
		Title:  "iMarket",
		Color:  "CEFF00",
		Models: make([]model, len(models), cap(models)),
	}

	// Заполняем срез цен нашей компании данными из таблицы
	for _, sheet := range currentFile.Sheets {
		for kr, row := range sheet.Rows {
			// Ряд с названиями компаний нам не нужен, пропускаем его
			if kr > 1 {
				// Нужная нам ячейка 2я слева
				val, ok := row.Cells[1].Int()
				// Запоминаем значение ячейки
				if ok == nil {
					imarketCompany.Models[kr-2].Price = val
				}
			}
		}
	}

	// Создаём новую таблицу
	newFile = xlsx.NewFile()
	// Добавляем страницу
	sheet, err = newFile.AddSheet("Цены")

	if err != nil {
		return err
	}

	// Устанавливаем ширину столбца для названий моделей
	sheet.SetColWidth(0, 0, 30)
	// Устанавливаем шрифт по умолчанию "Verdana" с размером 12
	xlsx.SetDefaultFont(12, "Verdana")

	// Переменная для выделенной рамки ячейки
	border := *xlsx.NewBorder("thin", "thin", "thin", "thin")
	// Переменная для расположения данных по центру яцейки
	centred := xlsx.Alignment{Horizontal: "center", Vertical: "center"}

	// Переенная для обозначения более низкой цены на конкретную модель
	// относительно цены нашей компании (красный фон)
	warning := &xlsx.Style{
		Fill:      *xlsx.NewFill("solid", "FF0000", ""),
		Alignment: centred,
		Border:    border,
	}

	// Добавляем ряд
	row = sheet.AddRow()
	// Устанавливаем высоту в 25 пунктов
	row.SetHeight(25)

	// Добавляем ячейку с текущей датой, центрируем и делаем её смежной
	cell = row.AddCell()
	cell.SetValue("Актуально на: " + time.Now().Format("02-01-2006"))
	cell.SetStyle(&xlsx.Style{
		Alignment: centred,
		Border:    border,
	})
	cell.Merge(len(c)+1, 0)

	// Первый ряд - названия компаний
	row = sheet.AddRow()
	row.SetHeight(25)
	cell = row.AddCell()
	cell.SetStyle(&xlsx.Style{
		Border: border,
	})

	// Заносим в ячейку название нашей компании
	cell = row.AddCell()
	cell.SetString(imarketCompany.Title)
	cell.SetStyle(&xlsx.Style{
		Fill:      *xlsx.NewFill("solid", imarketCompany.Color, ""),
		Alignment: centred,
		Border:    border,
	})

	// Заносим в ячейки названия других компаний
	for _, v := range c {
		cell = row.AddCell()
		cell.SetString(v.Title)
		cell.SetStyle(&xlsx.Style{
			Fill:      *xlsx.NewFill("solid", v.Color, ""),
			Alignment: centred,
			Border:    border,
		})
	}

	for i, model := range models {
		// Записываем в ячейку названием модели
		row = sheet.AddRow()
		row.SetHeight(25)
		cell = row.AddCell()
		cell.SetString(model)
		cell.SetStyle(&xlsx.Style{
			Font:      xlsx.Font{Bold: true},
			Alignment: xlsx.Alignment{Vertical: "center"},
			Border:    border,
		})

		// Записываем в ячейку цену на модель нашей компании
		cell = row.AddCell()
		cell.SetInt(imarketCompany.Models[i].Price)
		cell.SetStyle(&xlsx.Style{
			Fill:      *xlsx.NewFill("solid", imarketCompany.Color, ""),
			Alignment: centred,
			Border:    border,
		})

		// Записываем в ячейку цену на модель текущей компании
		for _, v := range c {
			// fmt.Println(v.Title, modelsNames[ii], v.Models[i].Price)
			cell = row.AddCell()
			cell.SetInt(v.Models[i].Price)
			cell.SetStyle(&xlsx.Style{
				Fill:      *xlsx.NewFill("solid", v.Color, ""),
				Alignment: centred,
				Border:    border,
			})
			// Проверям, является ли цена у конкурента меньше нашей
			if v.Models[i].Price > 0 && imarketCompany.Models[i].Price >= v.Models[i].Price {
				warningFlag = true
				cell.SetStyle(warning)
			} else {
				cell.SetStyle(&xlsx.Style{
					Fill:      *xlsx.NewFill("solid", v.Color, ""),
					Alignment: centred,
					Border:    border,
				})
			}
		}
	}

	// Заменяем таблицу новым файлом
	if err = newFile.Save(filePrice); err != nil {
		return err
	}

	return nil
}

// notify Выводит оповещение в ОС Windows
func notify() {
	if warningFlag && runtime.GOOS == "windows" {
		b := charmap.Windows1251.NewEncoder()

		title, err := b.Bytes([]byte("Цены конкурентов загружены!"))

		if err != nil {
			fmt.Println(err.Error())
		}

		message, err := b.Bytes([]byte("Имеются более низкие цены!"))

		if err != nil {
			fmt.Println(err.Error())
		}

		notification := toast.Notification{
			AppID:               "",
			Title:               string(title),
			Message:             string(message),
			ActivationArguments: "",
			Audio:               toast.SMS,
			Duration:            toast.Long,
		}

		if err = notification.Push(); err != nil {
			fmt.Println(err.Error())
		}
	}
}
