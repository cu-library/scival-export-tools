[![Go Report Card](https://goreportcard.com/badge/cu-library/scival-export-tools)](https://goreportcard.com/report/cu-library/scival-export-tools)

# scival-export-tools
SciVal Export Tools (SVET) is a command line tool developed to work on SciVal exports.

This tool currently has one subcommand, with room to grow later.

## Per Researcher

Taking a publications export and a list of researchers, the tool will build output with a line per publication,
with the author information as the first cells.

```
Researcher			Publication
Author	Level 1	Level 2	Title	Authors	Number of Authors	Scopus Author Ids	Year	Scopus Source title	Volume	Issue	Pages	ISSN	Source ID	Source-type	Field-Weighted View Impact	Citations	Field-Weighted Citation Impact	Outputs in Top Citation Percentiles, per percentile	Field-Weighted Outputs in Top Citation Percentiles, per percentile	Reference	DOI	Publication-type	EID	Institutions	Scopus affiliation names	Country
```

```
./svet perresearcher -h
Usage of perresearcher:
  -output string
        Per researcher output xlsx file. (default "PerResearcher.xlsx")
  -publications string
        Publications input file. (default "Publications-export.xlsx")
  -researchers string
        Researchers input file. (default "mySciVal_Researchers_Export.xlsx")
```

