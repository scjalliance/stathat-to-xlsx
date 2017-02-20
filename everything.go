package main

import (
	"log"
	"os"
	"strings"
	"time"

	"github.com/dustywilson/stathat"
	"github.com/tealeg/xlsx"
	kingpin "gopkg.in/alecthomas/kingpin.v2"
)

var (
	token    = kingpin.Arg("token", "StatHat access token (https://www.stathat.com/access)").Required().String()
	filename = kingpin.Arg("filename", "Output filename; needs to be something.xlsx").Required().String()
	period   = kingpin.Arg("period", "Data period (like 1M, 2w, etc)").Default("1w").String()
	stats    = kingpin.Arg("stats", "Stat IDs to fetch").Required().Strings()
	timezone = kingpin.Flag("timezone", "Override timezone").Short('z').String()
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
			for _, point := range stat.Points {
				row := sheet.AddRow()
				row.AddCell().SetDate(point.Time)
				row.AddCell().SetFloat(point.Value)
			}
		}
	}

	x.Save(*filename)
}
