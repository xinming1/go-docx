/*
   Copyright (c) 2024 mabiao0525 (马飚)

   This program is free software: you can redistribute it and/or modify
   it under the terms of the GNU Affero General Public License as published
   by the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   This program is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU Affero General Public License for more details.

   You should have received a copy of the GNU Affero General Public License
   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

package docx

import (
	"encoding/xml"
	"io"
	"strconv"
	"strings"
)

// SectPr show the properties of the document, like paper size
type SectPr struct {
	XMLName   xml.Name   `xml:"w:sectPr,omitempty"` // properties of the document, including paper size
	PgSz      *PgSz      `xml:"w:pgSz,omitempty"`
	PgMar     *PgMar     `xml:"w:pgMar,omitempty"`
	PgBorders *PgBorders `xml:"w:pgBorders,omitempty"`
	Cols      *Cols      `xml:"w:cols,omitempty"`
	DocGrid   *DocGrid   `xml:"w:docGrid,omitempty"`
}

// PgBorders show the page border
type PgBorders struct {
	XMLName    xml.Name    `xml:"w:pgBorders"`
	OffsetFrom string      `xml:"w:offsetFrom,attr,omitempty"` // 对应 w:offsetFrom="page"
	Top        *BorderLine `xml:"w:top"`
	Left       *BorderLine `xml:"w:left"`
	Bottom     *BorderLine `xml:"w:bottom"`
	Right      *BorderLine `xml:"w:right"`
}

type BorderLine struct {
	XMLName xml.Name `xml:""`             // 动态设置标签名（由父结构决定） Dynamically set tag names (determined by the parent structure)
	Val     string   `xml:"w:val,attr"`   // 边框类型（single/double等）Border type (single/double, etc.)
	Sz      string   `xml:"w:sz,attr"`    // 边框粗细（8=0.5pt，48=3pt）Border thickness (8=0.5pt, 48=3pt)
	Space   string   `xml:"w:space,attr"` // 边框与内容的间距 The spacing between the border and the content
	Color   string   `xml:"w:color,attr"` // 颜色值（十六进制或 auto）Color value (hexadecimal or auto)
}

func CreateBlackPageBorders() *PgBorders {
	borderLine := &BorderLine{
		Val:   "single",
		Sz:    "4",      // 对应 0.25pt
		Space: "24",     // 间距 24/8=3pt
		Color: "000000", // 黑色（注意不是 auto）
	}

	return &PgBorders{
		OffsetFrom: "page",
		Top:        borderLine,
		Left:       borderLine,
		Bottom:     borderLine,
		Right:      borderLine,
	}
}

// PgSz show the paper size
type PgSz struct {
	W int `xml:"w:w,attr"` // width of paper
	H int `xml:"w:h,attr"` // high of paper
}

// PgMar show the page margin
type PgMar struct {
	Top    int `xml:"w:top,attr"`
	Left   int `xml:"w:left,attr"`
	Bottom int `xml:"w:bottom,attr"`
	Right  int `xml:"w:right,attr"`
	Header int `xml:"w:header,attr"`
	Footer int `xml:"w:footer,attr"`
	Gutter int `xml:"w:gutter,attr"`
}

// Cols show the number of columns
type Cols struct {
	Space int `xml:"w:space,attr"`
}

// DocGrid show the document grid
type DocGrid struct {
	Type      string `xml:"w:type,attr"`
	LinePitch int    `xml:"w:linePitch,attr"`
}

// UnmarshalXML ...
func (sect *SectPr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "pgSz":
				var value PgSz
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				sect.PgSz = &value
			case "pgMar":
				var value PgMar
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				sect.PgMar = &value
			case "cols":
				var value Cols
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				sect.Cols = &value
			case "docGrid":
				var value DocGrid
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				sect.DocGrid = &value
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
			}
		}
	}
	return nil
}

// UnmarshalXML ...
func (pgsz *PgSz) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error

	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "w":
			pgsz.W, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "h":
			pgsz.H, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		default:
			// ignore other attributes now
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// UnmarshalXML ...
func (pgmar *PgMar) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error

	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "top":
			pgmar.Top, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "left":
			pgmar.Left, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "bottom":
			pgmar.Bottom, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "right":
			pgmar.Right, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "header":
			pgmar.Header, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "footer":
			pgmar.Footer, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "gutter":
			pgmar.Gutter, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		default:
			// ignore other attributes now
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// UnmarshalXML ...
func (cols *Cols) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error

	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "space":
			cols.Space, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		default:
			// ignore other attributes now
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// UnmarshalXML ...
func (dg *DocGrid) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error

	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "linePitch":
			dg.LinePitch, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "type":
			dg.Type = attr.Value
		default:
			// ignore other attributes now
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}
