// Copyright 2019 Carleton University Library All rights reserved.
// Use of this source code is governed by the MIT
// license that can be found in the LICENSE file.

package main

import (
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"os"
)

var (
	// A version flag, which should be overwritten when building using ldflags.
	version = "devel"
)

func main() {

	// The subcommand for creating a per-author publication listing.
	perauthor := flag.NewFlagSet("perauthor", flag.ExitOnError)
	perauthorhelp := fmt.Sprintf("%v    Create a new Excel spreadsheet with per-author publication listing.\n", perauthor.Name())
	publications := perauthor.String("publications", "Publications-export.xlsx", "Publications input file.")
	researchers := perauthor.String("researchers", "mySciVal_Researchers_Export.xlsx", "Researchers input file.")

	// Create the usage message which is printed on error.
	flag.Usage = func() {
		fmt.Fprintf(flag.CommandLine.Output(), "svet: Tools to transform and enhance SciVal exports.\nVersion %v\n\n", version)
		fmt.Fprintf(flag.CommandLine.Output(), "Subcommands:\n")
		fmt.Fprintf(flag.CommandLine.Output(), "    %v", perauthorhelp)
	}

	// flag.Parse() lets us catch --help and -h.
	flag.Parse()
	// If no subcommand was given, print usage and exit.
	if len(os.Args) == 1 {
		flag.Usage()
		os.Exit(2)
	}

	switch os.Args[1] {
	case perauthor.Name():
		perauthor.Parse(os.Args[2:])
	default:
		fmt.Fprintf(flag.CommandLine.Output(), "%s is not a valid subcommand. Available subcommands are:\n\n%v", os.Args[1], perauthorhelp)
		os.Exit(2)
	}

	if perauthor.Parsed() {
		err := PerAuthor(*publications, *researchers)
		if err != nil {
			os.Exit(2)
		}
	}

}

// PerAuthor processes the list of publications and the list of researchers to
// produce a list of per-author publications.
func PerAuthor(publications, researchers string) error {
	_, err := excelize.OpenFile(publications)
	if err != nil {
		fmt.Println("Unable to open publications file.")
		fmt.Println(err)
		return err
	}

	res, err := excelize.OpenFile(researchers)
	if err != nil {
		fmt.Println("Unable to open researchers file.")
		fmt.Println(err)
		return err
	}

	f := excelize.NewFile()

	// Set up the headers
	f.SetCellValue("Sheet1", "A1", "Researcher")
	err = f.MergeCell("Sheet1", "A1", "C1")
	if err != nil {
		return err
	}
	f.SetCellValue("Sheet1", "A2", "Author")
	f.SetCellValue("Sheet1", "B2", "Position Title")
	f.SetCellValue("Sheet1", "C2", "Level 2")

	f.SetCellValue("Sheet1", "D1", "Publication")
	err = f.MergeCell("Sheet1", "D1", "G1")
	if err != nil {
		return err
	}
	f.SetCellValue("Sheet1", "D2", "Title")

	resrows, err := res.GetRows("Sheet0")
	if err != nil {
		return err
	}
	for rowindex, row := range resrows {
		for colindex, cell := range row {
			fmt.Println(rowindex, colindex)
			writeTo, err := excelize.CoordinatesToCellName(colindex, rowindex+2)
			if err != nil {
				return nil
			}
			f.SetCellValue("Sheet1", writeTo, cell)
		}
	}

	// Save xlsx file by the given path.
	err = f.SaveAs("./PerAuthor.xlsx")
	if err != nil {
		fmt.Println("Unable to save PerAuthor output file.")
		fmt.Println(err)
		return err
	}

	return nil

}
