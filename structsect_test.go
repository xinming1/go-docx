package docx

import (
	"os"
	"testing"
)

func TestPgBorders(t *testing.T) {
	w := New().WithDefaultTheme().WithA4WithPgBorders()
	//w.AddParagraph().AddText("test")
	f, err := os.Create("testBorder.docx")
	if err != nil {
		panic(err)
	}
	_, err = w.WriteTo(f)
}
