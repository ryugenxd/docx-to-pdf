package main

import (
	"log"

	docxtopdf "github.com/ryugenxd/docx-to-pdf"
)

func main() {
	err := docxtopdf.ConvertFile("./TestDocument.docx", "./out.pdf")
	if err != nil {
		log.Fatal(err)
	}
}
