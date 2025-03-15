package docxtopdf

import (
	"archive/zip"
	"encoding/xml"
	"io/ioutil"
	"strings"

	"github.com/jung-kurt/gofpdf"
)

type Text struct {
	Text string `xml:",chardata"`
}

type Run struct {
	Texts []Text `xml:"t"`
}

type Paragraph struct {
	Runs []Run `xml:"r"`
}

type Document struct {
	Paragraphs []Paragraph `xml:"body>p"`
}

func extractTextFromDocx(docxPath string) (string, error) {
	reader, err := zip.OpenReader(docxPath)
	if err != nil {
		return "", err
	}
	defer reader.Close()

	var documentXML string
	for _, file := range reader.File {
		if file.Name == "word/document.xml" {
			rc, err := file.Open()
			if err != nil {
				return "", err
			}
			defer rc.Close()

			bytes, err := ioutil.ReadAll(rc)
			if err != nil {
				return "", err
			}
			documentXML = string(bytes)
			break
		}
	}

	var doc Document
	err = xml.Unmarshal([]byte(documentXML), &doc)
	if err != nil {
		return "", err
	}

	var textBuilder strings.Builder
	for _, para := range doc.Paragraphs {
		for _, run := range para.Runs {
			for _, text := range run.Texts {
				textBuilder.WriteString(text.Text + " ")
			}
		}
		textBuilder.WriteString("\n")
	}

	return textBuilder.String(), nil
}

func createPDF(text string, outputPath string) error {
	pdf := gofpdf.New("P", "mm", "A4", "")
	pdf.SetFont("Arial", "", 12)
	pdf.AddPage()
	pdf.MultiCell(190, 10, text, "", "L", false)

	err := pdf.OutputFileAndClose(outputPath)
	return err
}

func ConvertFile(inputFile string, outputFile string) error {
	text, err := extractTextFromDocx(inputFile)
	if err != nil {
		return err
	}

	err = createPDF(text, outputFile)
	if err != nil {
		return err
	}

	return nil
}
