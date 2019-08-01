package main

import (
	"log"
	"os"
	"strings"
	"time"

	"github.com/gentlemanautomaton/stathat"
	"github.com/tealeg/xlsx"
	kingpin "gopkg.in/alecthomas/kingpin.v2"
)

var (
	token      = kingpin.Arg("token", "StatHat access token (https://www.stathat.com/access)").Required().String()
	filename   = kingpin.Arg("filename", "Output filename; needs to be something.xlsx").Required().String()
	stats      = kingpin.Arg("stats", "Stat IDs to fetch").Required().Strings()
	period     = kingpin.Flag("period", "Data period (like 1M, 2w, etc)").Short('p').Default("1w").String()
	timezone   = kingpin.Flag("timezone", "Override timezone").Short('z').String()
	datetype   = kingpin.Flag("datetype", "Date Type").Short('t').Default("date").Enum("date", "string", "epoch")
	dateformat = kingpin.Flag("dateformat", "Date Format (https://golang.org/pkg/time/#Time.Format)").Short('f').Default("2006/01/02").String()
)

var l = log.New(os.Stderr, "", 0)

func init() {
	kingpin.Parse()
	stathat.UserAgent = `https://github.com/scjalliance/stathat-to-xlsx`
	if *timezone != "" {
		time.Local, _ = time.LoadLocation(*timezone)
	}
	if !strings.HasSuffix(*filename, ".xlsx") {
		l.Println("Filename must end in .xlsx")
		os.Exit(1)
	}
}

func main() {
	s := stathat.New().Token(*token)

	x := xlsx.NewFile()
	titleStyle := xlsx.NewStyle()
	titleStyle.Font.Bold = true

	for _, id := range *stats {
		data, err := s.Get(stathat.GetOptions{
			Period:  *period,
			Summary: true,
		}, id)
		if err != nil {
			l.Println(err)
			os.Exit(1)
		}

		for _, stat := range data {
			sheet, err := x.AddSheet(id)
			if err != nil {
				l.Println(err)
				os.Exit(1)
			}
			row := sheet.AddRow()
			title := row.AddCell()
			title.SetString(stat.Name)
			title.SetStyle(titleStyle)
			for i := len(stat.Points) - 1; i >= 0; i-- {
				point := stat.Points[i]
				row := sheet.AddRow()
				if *datetype == "date" {
					row.AddCell().SetDate(point.Time)
				} else if *datetype == "epoch" {
					row.AddCell().SetInt64(point.Time.Unix())
				} else if *datetype == "string" {
					row.AddCell().SetString(point.Time.Format(*dateformat))
				}
				row.AddCell().SetFloat(point.Value)
			}
		}
	}

	x.Save(*filename)
}
