package style

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

type Maker struct {
	f *excelize.File
}

func NewMaker(f *excelize.File) *Maker {
	return &Maker{f}
}

func (m *Maker) MustNewStyleID(styles ...any) int {
	s := &excelize.Style{}
	for _, style := range styles {
		switch val := style.(type) {
		case *excelize.Font:
			s.Font = val
		case []excelize.Border:
			s.Border = val
		case *excelize.Alignment:
			s.Alignment = val
		case *excelize.Protection:
			s.Protection = val
		case int:
			s.NumFmt = val
		case *int:
			s.DecimalPlaces = val
		case *string:
			s.CustomNumFmt = val
		case bool:
			s.NegRed = val
		default:
			panic(fmt.Sprintf("unsupport type: %T", val))
		}
	}
	styleID, err := m.f.NewStyle(s)
	if err != nil {
		panic(err)
	}
	return styleID
}
