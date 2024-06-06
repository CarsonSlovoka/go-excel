package style

import "github.com/xuri/excelize/v2"

var AlignmentCenter = &excelize.Alignment{
	Horizontal: "center",
	Vertical:   "center",
}

var BorderBold = []excelize.Border{
	{Type: "left", Color: "000000", Style: 5},
	{Type: "top", Color: "000000", Style: 5},
	{Type: "bottom", Color: "000000", Style: 5},
	{Type: "right", Color: "000000", Style: 5},
	// {Type: "diagonalDown", Color: "000000", Style: 5},
	// {Type: "diagonalUp", Color: "000000", Style: 5},
}
