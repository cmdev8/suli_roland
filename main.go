package main

import (
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"strconv"

	"github.com/pkg/errors"
	"github.com/tealeg/xlsx/v3"
)

type studentResult struct {
	ItemName string
	Score    int
}

func main() {
	inputDir := flag.String("input_dir", "", "az xlsx fileokat tartalmazo konyvtar neve")
	outputFilePath := flag.String("out", "./eredmeny.xlsx", "az eredmeny xlsx teljes eleresi utja filenevvel egyutt")

	flag.Parse()

	baseDir := *inputDir

	summary := map[string]int{}

	files, err := ioutil.ReadDir(baseDir)
	if err != nil {
		fmt.Println("konyvtar beolvasasi hiba", err.Error())
	}

	for _, f := range files {
		filePath := filepath.Join(baseDir, f.Name())

		if r, err := parse(filePath); err == nil {
			for _, row := range r {
				summary[row.ItemName] += row.Score
			}
		} else {
			fmt.Println("HIBA", f.Name(), err.Error())
		}
	}

	if writeErr := writeResults(*outputFilePath, summary); writeErr != nil {
		fmt.Println("eredmeny kiirasi hiba:", writeErr.Error())
	}
}

func parse(filePath string) ([]studentResult, error) {
	wb, openErr := xlsx.OpenFile(filePath)
	if openErr != nil {
		return nil, errors.Wrapf(openErr, "nem lehet megnyitni a filet: %s", filePath)
	}

	if len(wb.Sheets) < 2 {
		return nil, errors.New("kevesebb mint ketto sheet talalhato")
	}

	sh, sheetOk := wb.Sheet[wb.Sheets[1].Name]
	if !sheetOk {
		return nil, errors.New("nem lehet megnyitni a sheetet")
	}

	r := []studentResult{}

	for i := 0; i < 8; i++ {
		score, scoreErr := sh.Cell(i, 11)
		if scoreErr != nil {
			return nil, errors.Wrapf(scoreErr, "pontszam beolvasasi hiba, ebben a sorban: %d", i)
		}
		scoreInt, _ := strconv.Atoi(score.String())

		itemName, itemNameErr := sh.Cell(i, 12)
		if itemNameErr != nil {
			return nil, errors.Wrapf(itemNameErr, "pontszamnev beolvasasi hiba, ebben a sorban: %d", i)
		}

		r = append(r, studentResult{
			ItemName: itemName.String(),
			Score:    scoreInt,
		})
	}

	return r, nil
}

func writeResults(outFilePath string, data map[string]int) error {
	xlsxFile := xlsx.NewFile()
	sh, _ := xlsxFile.AddSheet("Osszegzes")

	for k, v := range data {
		row := sh.AddRow()
		row.AddCell().SetString(k)
		row.AddCell().SetInt(v)
	}

	out, err := os.Create(outFilePath)

	if err != nil {
		return err
	}

	defer out.Close()
	xlsxFile.Write(out)

	return nil
}
