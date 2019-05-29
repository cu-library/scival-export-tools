// Copyright 2019 Carleton University Library All rights reserved.
// Use of this source code is governed by the MIT
// license that can be found in the LICENSE file.

package main

import (
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"os"
	"strings"
)

var (
	// A version flag, which should be overwritten when building using ldflags.
	version = "devel"
)

func main() {

	// The subcommand for creating a per-researcher publication listing.
	perresearcher := flag.NewFlagSet("perresearcher", flag.ExitOnError)
	perresearcherhelp := fmt.Sprintf("%v    Create a new Excel spreadsheet with per-researcher publication listing.\n", perresearcher.Name())
	publications := perresearcher.String("publications", "Publications-export.xlsx", "Publications input file.")
	researchers := perresearcher.String("researchers", "mySciVal_Researchers_Export.xlsx", "Researchers input file.")
	output := perresearcher.String("output", "PerResearcher.xlsx", "Per researcher output xlsx file.")

	// Create the usage message which is printed on error.
	flag.Usage = func() {
		fmt.Fprintf(flag.CommandLine.Output(), "svet: Tools to transform and enhance SciVal exports.\nVersion %v\n\n", version)
		fmt.Fprintf(flag.CommandLine.Output(), "Subcommands:\n")
		fmt.Fprintf(flag.CommandLine.Output(), "    %v", perresearcherhelp)
	}

	// flag.Parse() lets us catch --help and -h.
	flag.Parse()
	// If no subcommand was given, print usage and exit.
	if len(os.Args) == 1 {
		flag.Usage()
		os.Exit(2)
	}

	switch os.Args[1] {
	case perresearcher.Name():
		perresearcher.Parse(os.Args[2:])
	default:
		fmt.Fprintf(flag.CommandLine.Output(), "%s is not a valid subcommand. Available subcommands are:\n\n%v", os.Args[1], perresearcherhelp)
		os.Exit(2)
	}

	if perresearcher.Parsed() {
		err := PerResearcher(*publications, *researchers, *output)
		if err != nil {
			fmt.Println("Error!")
			os.Exit(2)
		}
	}

}

// PerResearcher processes the list of publications and the list of researchers to
// produce a list of per-researcher publications.
func PerResearcher(publications, researchers, output string) error {

	// Open the publications excel file.
	pubs, err := excelize.OpenFile(publications)
	if err != nil {
		fmt.Println("Unable to open publications file.")
		fmt.Println(err)
		return err
	}

	// Open the researchers excel file.
	res, err := excelize.OpenFile(researchers)
	if err != nil {
		fmt.Println("Unable to open researchers file.")
		fmt.Println(err)
		return err
	}

	// Output
	out := excelize.NewFile()

	// Set up the headers

	// Create the header style.
	headerstyle, err := out.NewStyle(`{"fill":{"type":"pattern","color":["#C0C0C0"],"pattern":1}, "alignment":{"horizontal":"center"}}`)
	if err != nil {
		fmt.Println("Unable to create header style.")
		fmt.Println(err)
		return err
	}

	// Headers for researcher section.
	out.SetCellValue("Sheet1", "A1", "Researcher")
	err = out.MergeCell("Sheet1", "A1", "C1")
	if err != nil {
		fmt.Println("Unable to merge cells.")
		fmt.Println(err)
		return err
	}
	err = out.SetCellStyle("Sheet1", "A1", "C2", headerstyle)
	if err != nil {
		fmt.Println("Unable to set header style.")
		fmt.Println(err)
		return err
	}
	out.SetCellValue("Sheet1", "A2", "Author")
	out.SetCellValue("Sheet1", "B2", "Level 1")
	out.SetCellValue("Sheet1", "C2", "Level 2")

	// Headers for Publication section.
	out.SetCellValue("Sheet1", "D1", "Publication")
	err = out.MergeCell("Sheet1", "D1", "G1")
	if err != nil {
		fmt.Println("Unable to merge cells.")
		fmt.Println(err)
		return err
	}
	err = out.SetCellStyle("Sheet1", "D1", "F2", headerstyle)
	if err != nil {
		fmt.Println("Unable to set header style.")
		fmt.Println(err)
		return err
	}

	// Copy and process data.

	// First, create a map of Scopus Author IDs to publication row.
	pubrows, err := pubs.GetRows(pubs.GetSheetMap()[1])
	if err != nil {
		fmt.Println("Unable to get rows.")
		fmt.Println(err)
		return err
	}

	// The publication listing header, string value to index.
	pubsHeader := map[string]int{}

	// A map of scopus IDs to publication rows.
	scopusIDToPubRow := map[string][]int{}

	// Go through the headers in the publication listing.
	for colindex, cell := range pubrows[0] {
		pubsHeader[cell] = colindex
		if colindex == 0 {
			continue
		}
		headerCellName, err := excelize.CoordinatesToCellName(colindex+3, 2)
		if err != nil {
			fmt.Println("Unable to get cell name.")
			fmt.Println(err)
			return err
		}
		out.SetCellValue("Sheet1", headerCellName, cell)
		err = out.SetCellStyle("Sheet1", headerCellName, headerCellName, headerstyle)
		if err != nil {
			fmt.Println("Unable to set header style.")
			fmt.Println(err)
			return err
		}
	}

	// Find the column index for the scopus author IDs field.
	scopusIDColIndex, ok := pubsHeader["Scopus Author Ids"]
	if !ok {
		fmt.Println("Couldn't find 'Scopus Author Ids' in publication export header.")
		return fmt.Errorf("")
	}

	// Build the map of scopus IDs to publication rows.
	for rowindex, row := range pubrows {
		if rowindex == 0 {
			continue
		}
		for _, scopusID := range strings.Split(row[scopusIDColIndex], ",") {
			scopusID = strings.TrimSpace(scopusID)
			_, exists := scopusIDToPubRow[scopusID]
			if exists {
				scopusIDToPubRow[scopusID] = append(scopusIDToPubRow[scopusID], rowindex)
			} else {
				scopusIDToPubRow[scopusID] = []int{rowindex}
			}
		}
	}

	// Go through researcher listing
	resrows, err := res.GetRows("Sheet0")
	if err != nil {
		fmt.Println("Unable to get rows.")
		fmt.Println(err)
		return err
	}
	// At the beginning, we need to move "down" two rows to accommodate the header.
	// We increment this every time we add more than one publication per researcher.
	extrarows := 2

	for rowindex, row := range resrows {
		if rowindex == 0 {
			continue
		}
		scopusID := strings.TrimSpace(row[1])

		// If no listings are found, still print the researcher info
		// TODO: Make this a function.
		if len(scopusIDToPubRow[scopusID]) == 0 {
			authorCellName, err := excelize.CoordinatesToCellName(1, rowindex+extrarows)
			if err != nil {
				fmt.Println("Unable to get cell name.")
				fmt.Println(err)
				return err
			}
			level1CellName, err := excelize.CoordinatesToCellName(2, rowindex+extrarows)
			if err != nil {
				fmt.Println("Unable to get cell name.")
				fmt.Println(err)
				return err
			}
			level2CellName, err := excelize.CoordinatesToCellName(3, rowindex+extrarows)
			if err != nil {
				fmt.Println("Unable to get cell name.")
				fmt.Println(err)
				return err
			}

			out.SetCellValue("Sheet1", authorCellName, row[0])
			out.SetCellValue("Sheet1", level1CellName, row[2])
			out.SetCellValue("Sheet1", level2CellName, row[3])
		}

		for numpubs, pubrow := range scopusIDToPubRow[scopusID] {

			// If we are adding more than one publication per researcher...
			if numpubs > 0 {
				extrarows++
			}

			authorCellName, err := excelize.CoordinatesToCellName(1, rowindex+extrarows)
			if err != nil {
				fmt.Println("Unable to get cell name.")
				fmt.Println(err)
				return err
			}
			level1CellName, err := excelize.CoordinatesToCellName(2, rowindex+extrarows)
			if err != nil {
				fmt.Println("Unable to get cell name.")
				fmt.Println(err)
				return err
			}
			level2CellName, err := excelize.CoordinatesToCellName(3, rowindex+extrarows)
			if err != nil {
				fmt.Println("Unable to get cell name.")
				fmt.Println(err)
				return err
			}

			out.SetCellValue("Sheet1", authorCellName, row[0])
			out.SetCellValue("Sheet1", level1CellName, row[2])
			out.SetCellValue("Sheet1", level2CellName, row[3])

			for colindex, cell := range pubrows[pubrow] {
				if colindex == 0 {
					continue
				}
				writeTo, err := excelize.CoordinatesToCellName(colindex+3, rowindex+extrarows)
				if err != nil {
					fmt.Println("Unable to get cell name.")
					fmt.Println(err)
					return err
				}
				out.SetCellValue("Sheet1", writeTo, cell)
			}
		}
	}

	// Save xlsx file by the given path.
	err = out.SaveAs(output)
	if err != nil {
		fmt.Printf("Unable to save output file to %v.\n", output)
		fmt.Println(err)
		return err
	}

	return nil
}
