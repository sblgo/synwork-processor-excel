package excel

import (
	"context"
	"fmt"
	"strings"

	excelize "github.com/xuri/excelize/v2"
	"sbl.systems/go/synwork/plugin-sdk/schema"
)

// https://xuri.me/excelize/en/cell.html#SetCellStyle
var (
	excel_cell = map[string]*schema.Schema{
		"name":         {Type: schema.TypeString, Required: true, DefaultValue: ""},
		"value":        {Type: schema.TypeString, Optional: true},
		"double_value": {Type: schema.TypeFloat, Optional: true},
		"int_value":    {Type: schema.TypeInt, Optional: true},
		"style":        {Type: schema.TypeString, Optional: true, DefaultValue: ""},
	}
	excel_sheet = map[string]*schema.Schema{
		"name": {Type: schema.TypeString, Required: true, DefaultValue: "Sheet1"},
		"cell": {Type: schema.TypeList, Required: true, Elem: excel_cell},
	}
	excel_style = map[string]*schema.Schema{
		"name":           {Type: schema.TypeString, Required: true},
		"lang":           {Type: schema.TypeString, Optional: true, DefaultValue: ""},
		"neg-red":        {Type: schema.TypeBool, Optional: true, DefaultValue: false},
		"decimal-places": {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
		"num-fmt":        {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
		"fill": {
			Type:     schema.TypeMap,
			Optional: true,
			Elem: map[string]*schema.Schema{
				"color":   {Type: schema.TypeString, Optional: true, DefaultValue: ""},
				"pattern": {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
				"shading": {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
				"type":    {Type: schema.TypeString, Optional: true, DefaultValue: ""},
			},
		},
		"alignment": {
			Type:     schema.TypeMap,
			Optional: true,
			Elem: map[string]*schema.Schema{
				"horizontal":        {Type: schema.TypeString, Optional: true, DefaultValue: ""},
				"indent":            {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
				"justify-last-line": {Type: schema.TypeBool, Optional: true, DefaultValue: false},
				"reading-order":     {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
				"relative-indent":   {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
				"text-rotation":     {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
				"vertical":          {Type: schema.TypeString, Optional: true, DefaultValue: ""},
				"shrink-to-fit":     {Type: schema.TypeBool, Optional: true, DefaultValue: false},
				"wrap-text":         {Type: schema.TypeBool, Optional: true, DefaultValue: false},
			},
		},
		"border": {
			Type:     schema.TypeList,
			Optional: true,
			Elem: map[string]*schema.Schema{
				"style": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
				"type":  {Type: schema.TypeString, Optional: true, DefaultValue: ""},
				"color": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
			},
		},
		"font": {
			Type:     schema.TypeMap,
			Optional: true,
			Elem: map[string]*schema.Schema{
				"bold":      {Type: schema.TypeBool, Optional: true, DefaultValue: false},
				"italic":    {Type: schema.TypeBool, Optional: true, DefaultValue: false},
				"strike":    {Type: schema.TypeBool, Optional: true, DefaultValue: false},
				"underline": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
				"family":    {Type: schema.TypeString, Optional: true, DefaultValue: ""},
				"color":     {Type: schema.TypeString, Optional: true, DefaultValue: ""},
				"size":      {Type: schema.TypeFloat, Optional: true, DefaultValue: 8.0},
			},
		},
	}
	excel_file = map[string]*schema.Schema{
		"file-name": {Type: schema.TypeString, Required: true, DefaultValue: "file1.xlsx"},
		"sheet":     {Type: schema.TypeList, Optional: true, Elem: excel_sheet},
		"style":     {Type: schema.TypeList, Optional: true, Elem: excel_style},
	}
)

var Method_write_file = &schema.Method{
	Schema:   excel_file,
	Result:   map[string]*schema.Schema{},
	ExecFunc: excel_write_file,
	Description: `Method write_excel_file provides a way to create excel file based on a configuration.

	method "write_excel_file" "processor-instance" "method-instance" {
		file-name = "test01.xlsx"
		sheet {
			name = "sheet01"
			cell { 
				name = "A1"
				value = "content of field A1"
			}
			cell { 
				name = "A2:C3"
				double_value = 1.22
				style = "grey"
			}
		}
		style {
			name = "grey"
			fill {
				color   = "#888888,#FFFFFF"
				type    = "gradient"
				shading = 1
			}
		}
	}
	
	`,
}

func excel_write_file(ctx context.Context, data *schema.MethodData, client interface{}) error {
	f := excelize.NewFile()
	fileName := data.GetConfig("file-name").(string)
	sheets := data.GetConfig("sheet").([]interface{})
	styles, err := excel_define_styles(ctx, f, data)
	if err != nil {
		return err
	}
	sheetsToRemove := map[string]bool{}
	for count := f.SheetCount; count > 0; count-- {
		sheetName := f.GetSheetName(count - 1)
		sheetsToRemove[sheetName] = true
	}
	for _, s := range sheets {
		sheet := s.(map[string]interface{})
		sheetName := sheet["name"].(string)
		if _, ok := sheetsToRemove[sheetName]; ok {
			delete(sheetsToRemove, sheetName)
		}
		f.NewSheet(sheetName)
		for _, c := range sheet["cell"].([]interface{}) {
			cell := c.(map[string]interface{})
			cellName := cell["name"].(string)
			cellEnd := cellName
			if strings.Contains(cellName, ":") {
				parts := strings.Split(cellName, ":")
				if len(parts) != 2 {
					return fmt.Errorf("error merging cell %s", cellName)
				}
				if err := f.MergeCell(sheetName, parts[0], parts[1]); err != nil {
					return err
				} else {
					cellName = parts[0]
					cellEnd = parts[1]
				}
			}
			if style, existStyle := styles[cell["style"].(string)]; existStyle {
				f.SetCellStyle(sheetName, cellName, cellEnd, style)
			}
			for _, k := range []string{"value", "int_value", "double_value"} {
				if v := cell[k]; v != nil {

					if err := f.SetCellValue(sheetName, cellName, v); err != nil {
						return err
					} else {
						goto next_cell
					}
				}
			}
		next_cell:
		}
	}
	for k, _ := range sheetsToRemove {
		f.DeleteSheet(k)
	}

	if err := f.SaveAs(fileName); err != nil {
		return err
	}
	return nil
}

func excel_define_styles(ctx context.Context, f *excelize.File, data *schema.MethodData) (map[string]int, error) {
	styles := data.GetConfig("style")
	styleDefs := make(map[string]int)
	if styles == nil {
		return styleDefs, nil
	}
	for _, s := range styles.([]interface{}) {
		style := s.(map[string]interface{})
		s := &excelize.Style{
			Alignment:     excel_style_alignment(style["alignment"]),
			Border:        excel_style_borders(styleDefs, style["border"]),
			Fill:          excel_style_fill(style["fill"]),
			Lang:          style["lang"].(string),
			NegRed:        style["neg-red"].(bool),
			DecimalPlaces: style["decimal-places"].(int),
			NumFmt:        style["num-fmt"].(int),
			Font:          excel_style_font(style["font"]),
		}

		if newStyle, err := f.NewStyle(s); err != nil {
			return nil, err
		} else {
			styleDefs[style["name"].(string)] = newStyle
		}
	}

	return styleDefs, nil
}

func excel_style_font(v interface{}) *excelize.Font {
	if v == nil {
		return nil
	}
	font := v.(map[string]interface{})
	return &excelize.Font{
		Bold:      font["bold"].(bool),
		Italic:    font["italic"].(bool),
		Strike:    font["strike"].(bool),
		Underline: font["underline"].(string),
		Family:    font["family"].(string),
		Color:     font["color"].(string),
		Size:      font["size"].(float64),
	}
}

func excel_style_fill(v interface{}) excelize.Fill {
	f := excelize.Fill{}
	if v == nil {
		return f
	}
	fill := v.(map[string]interface{})
	f.Color = []string{}
	if c := fill["color"]; c != nil {
		switch t := c.(type) {
		case string:
			p := strings.Split(t, ",")
			f.Color = append(f.Color, p...)
		case []interface{}:
		}
	}
	f = excelize.Fill{
		Pattern: fill["pattern"].(int),
		Shading: fill["shading"].(int),
		Type:    fill["type"].(string),
		Color:   f.Color,
	}
	return f
}

func excel_style_alignment(v interface{}) *excelize.Alignment {
	if v == nil {
		return nil
	}
	align := v.(map[string]interface{})
	a := &excelize.Alignment{
		Horizontal:      align["horizontal"].(string),
		Indent:          align["indent"].(int),
		JustifyLastLine: align["justify-last-line"].(bool),
		ReadingOrder:    uint64(align["reading-order"].(int)),
		RelativeIndent:  align["relative-indent"].(int),
		ShrinkToFit:     align["shrink-to-fit"].(bool),
		TextRotation:    align["text-rotation"].(int),
		Vertical:        align["vertical"].(string),
		WrapText:        align["wrap-text"].(bool),
	}

	return a
}

func excel_style_borders(styles map[string]int, v interface{}) []excelize.Border {
	borders := []excelize.Border{}
	if v == nil {
		return borders
	}
	bDefs := v.([]interface{})
	for _, bRaw := range bDefs {
		var style int
		border := bRaw.(map[string]interface{})
		style = styles[border["style"].(string)]
		if style == 0 {
			style = 1
		}
		borders = append(borders, excelize.Border{
			Type:  border["type"].(string),
			Color: border["color"].(string),
			Style: style,
		})
	}
	return borders
}
