package excel

import (
	"context"
	"fmt"
	"regexp"
	"strings"

	excelize "github.com/xuri/excelize/v2"
	"sbl.systems/go/synwork/plugin-sdk/schema"
	"sbl.systems/go/synwork/plugin-sdk/utils"
)

// https://xuri.me/excelize/en/cell.html#SetCellStyle
var (
	excel_file_read = map[string]*schema.Schema{
		"file-name": {Type: schema.TypeString, Required: true, DefaultValue: "file1.xlsx"},
		"sheet": {
			Type: schema.TypeList,
			Elem: map[string]*schema.Schema{
				"when": {
					Type: schema.TypeList,
					Elem: map[string]*schema.Schema{
						"pattern": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
						"low":     {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
						"high":    {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
					},
				},
				"row": {Type: schema.TypeString, Required: true, DefaultValue: ""},
			},
		},
		"row": {
			Type: schema.TypeList,
			Elem: map[string]*schema.Schema{
				"name": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
				"when": {
					Type: schema.TypeList,
					Elem: map[string]*schema.Schema{
						"col":     {Type: schema.TypeString, Optional: true, DefaultValue: ""},
						"pattern": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
					},
				},
				"cell": {
					Type: schema.TypeList,
					Elem: map[string]*schema.Schema{
						"col": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
						"tag": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
					},
				},
				"children": {Type: schema.TypeString, Required: true, DefaultValue: ""},
			},
		},
	}
)

var Method_read_file = &schema.Method{
	Schema: excel_file_read,
	Result: map[string]*schema.Schema{
		"sheets": {
			Type: schema.TypeList,
			Elem: map[string]*schema.Schema{
				"name":  {Type: schema.TypeString, Required: true},
				"index": {Type: schema.TypeInt, Required: true},
				"rows": {
					Type: schema.TypeList,
					Elem: _excel_row_element(),
				},
			},
		},
	},
	ExecFunc: excel_read_file,
	Description: `Method read_excel_file provides a way to read excel file based on a configuration.

	method "read_excel_file" "processor-instance" "method-instance" {
		file-name = "test01.xlsx"
		sheet {
			when {
				name = "pattern"
			}
			when {
				low = 1
				high = 2
			}
			row = "standard"

		}

		row {
			name = "standard"
			when {
				col = "A"
				pattern = "^AD$" // regulare expression
			}
			cell {
				col = "A"
				tag = "lineType"
			}
			children = "rowType1,rowType2"
		}

	}
	
	result has following structure:

	sheets:[
		{
			name : "sheet01",
			index : 1,
			rows : [
				{
					index : 1,
					name : "standard",
					cols : [
						{
							col : "A",
							tag : "tag01",
							value : "",
						}
					],
					children : [
						{
							name : "child01",
							rows : []
						}
					]
				}

			]
		}

	]

	`,
}

type (
	ExcelReadSheet struct {
		Conditions ExcelReadSheetConditions
		Row        string
	}
	ExcelReadSheetCondition interface {
		Test(name string, index int) (bool, error)
	}
	ExcelReadSheetCondPattern struct {
		Pattern string
	}
	ExcelReadSheetCondRange struct {
		Low, High int
	}
	ExcelReadSheetConditions []ExcelReadSheetCondition
	ExcelReadRow             struct {
		Name       string
		Children   []string
		Conditions ExcelReadRowConditions
		Cells      ExcelReadRowCells
	}
	ExcelReadRowCondition struct {
		Col     string
		Pattern string
	}
	ExcelReadRowConditions []*ExcelReadRowCondition
	ExcelReadRowCell       struct {
		Col string
		Tag string
	}
	ExcelReadRowCells   []*ExcelReadRowCell
	ExcelReadRepository struct {
		Sheets []*ExcelReadSheet
		Rows   map[string]*ExcelReadRow
	}
	readStack struct {
		parent     *readStack
		row        *ExcelReadRow
		childIndex int
		sheet      *ExcelReadSheet
		currentRow *ExcelDataRow
		allRows    ExcelDataRows
	}

	ExcelDataSheet struct {
		Name  string        `json:"name"`
		Index int           `json:"index"`
		Rows  ExcelDataRows `json:"rows"`
	}
	ExcelDataSheets []*ExcelDataSheet
	ExcelDataRows   []*ExcelDataRow
	ExcelDataRow    struct {
		Name     string            `json:"name"`
		Index    int               `json:"index"`
		Cols     ExcelDataCols     `json:"cols"`
		Children ExcelDataChildren `json:"child"`
	}
	ExcelDataCols []*ExcelDataCol
	ExcelDataCol  struct {
		Col   string      `json:"col"`
		Tag   string      `json:"tag"`
		Value interface{} `json:"value"`
	}
	ExcelDataChildren []*ExcelDataChild
	ExcelDataChild    struct {
		Name string
		Rows ExcelDataRows
	}
)

func (c ExcelReadRowConditions) Test(values []string) (bool, error) {
	for _, cond := range c {
		if idx, err := excelize.ColumnNameToNumber(cond.Col); err != nil {
			return false, err
		} else if idx-1 >= len(values) {
			return false, nil
		} else if expr, err := regexp.Compile(cond.Pattern); err != nil {
			return false, err
		} else if !expr.MatchString(values[idx-1]) {
			return false, nil
		}
	}
	return true, nil
}

func (c ExcelReadRowCells) Apply(values []string) (ExcelDataCols, error) {
	cols := make(ExcelDataCols, 0)
	for _, col := range c {
		if idx, err := excelize.ColumnNameToNumber(col.Col); err != nil {
			return nil, err
		} else if idx-1 < len(values) {
			cols = append(cols, &ExcelDataCol{
				Col:   col.Col,
				Tag:   col.Tag,
				Value: values[idx-1],
			})
		}
	}
	return cols, nil
}

func (c *ExcelReadSheetCondPattern) Test(name string, index int) (bool, error) {
	pattern, err := regexp.Compile(c.Pattern)
	if err != nil {
		return false, err
	}
	return pattern.MatchString(name), nil
}

func (c *ExcelReadSheetCondRange) Test(name string, index int) (bool, error) {
	return c.Low <= index && index <= c.High, nil
}

func excel_read_file(ctx context.Context, data *schema.MethodData, client interface{}) error {
	fileName := data.GetConfig("file-name").(string)
	repository, err := excel_read_file_configure(ctx, data)
	if err != nil {
		return err
	}
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return err
	}
	resultSheets := make(ExcelDataSheets, 0)
	for sheetIdx := 0; sheetIdx < f.SheetCount; sheetIdx++ {
		sheetName := f.GetSheetName(sheetIdx)
		cfgSheetIdx := 0
	select_cfg_sheet:
		for cfgSheetIdx < len(repository.Sheets) {
			cfgSheet := repository.Sheets[cfgSheetIdx]
			for _, cfgSheetCond := range cfgSheet.Conditions {
				if ok, err := cfgSheetCond.Test(sheetName, sheetIdx+1); err != nil {
					return err
				} else if !ok {
					cfgSheetIdx++
					goto select_cfg_sheet
				}
			}
			resultSheet := &ExcelDataSheet{
				Name:  sheetName,
				Index: sheetIdx + 1,
				Rows:  make(ExcelDataRows, 0),
			}
			sheetTop := &readStack{
				sheet:      cfgSheet,
				childIndex: -1,
				row:        repository.Rows[cfgSheet.Row],
			}
			stack := sheetTop
			rows, _ := f.Rows(sheetName)
			rowIdx := 0
			for rows.Next() {
				rowIdx++
				cells, _ := rows.Columns()
			read_cells:
				if stack.row == nil {
					return fmt.Errorf("there is no definition for %s used in sheet %s", cfgSheet.Row, sheetName)
				}
				if ok, err := stack.row.Conditions.Test(cells); err != nil {
					return err
				} else if ok {
					if eCells, err := stack.row.Cells.Apply(cells); err != nil {
						return err
					} else {
						stack.currentRow = &ExcelDataRow{
							Name:     stack.row.Name,
							Index:    rowIdx,
							Cols:     eCells,
							Children: make(ExcelDataChildren, 0),
						}
						stack.allRows = append(stack.allRows, stack.currentRow)
					}
				} else if stack.currentRow != nil && stack.childIndex+1 < len(stack.row.Children) {
					stack.childIndex++
					child := &readStack{
						parent:     stack,
						childIndex: -1,
						row:        repository.Rows[stack.row.Children[stack.childIndex]],
					}
					stack = child
					goto read_cells
				} else if stack.parent != nil {
					if stack.currentRow != nil {
						stack.parent.currentRow.Children = append(stack.parent.currentRow.Children, &ExcelDataChild{
							Name: stack.currentRow.Name,
							Rows: stack.allRows,
						})
					}
					stack = stack.parent
					goto read_cells
				}
			}
			cfgSheetIdx++
			resultSheet.Rows = stack.allRows
			resultSheets = append(resultSheets, resultSheet)
		}

	}

	sheets := []interface{}{}
	for _, item := range resultSheets {
		enc := utils.NewEncoder()
		if encItem, err := enc.Encode(item); err != nil {
			return err
		} else {
			sheets = append(sheets, encItem)
		}
	}
	data.SetResult("sheets", sheets)
	return nil
}

func excel_read_file_configure(ctx context.Context, data *schema.MethodData) (*ExcelReadRepository, error) {
	repository := &ExcelReadRepository{
		Sheets: make([]*ExcelReadSheet, 0),
		Rows:   make(map[string]*ExcelReadRow),
	}
	sheetRaw := data.GetConfig("sheet")
	if sheets, ok := sheetRaw.([]interface{}); ok {
		for _, sheetJson := range sheets {
			if sheet, err := excel_read_file_config_sheet(ctx, sheetJson); err != nil {
				return nil, err
			} else {
				repository.Sheets = append(repository.Sheets, sheet)
			}
		}
	}
	rowRaw := data.GetConfig("row")
	if rows, ok := rowRaw.([]interface{}); ok {
		for _, rowJson := range rows {
			if row, err := excel_read_file_config_row(ctx, rowJson); err != nil {
				return nil, err
			} else {
				repository.Rows[row.Name] = row
			}
		}
	}
	return repository, nil
}

func excel_read_file_config_row(ctx context.Context, rowRaw interface{}) (*ExcelReadRow, error) {
	rd := rowRaw.(map[string]interface{})
	conds, err := excel_read_file_config_row_conds(ctx, rd["when"])
	if err != nil {
		return nil, err
	}
	cells, err := excel_read_file_config_row_cells(ctx, rd["cell"])
	if err != nil {
		return nil, err
	}
	readRow := &ExcelReadRow{
		Name:       rd["name"].(string),
		Children:   []string{},
		Conditions: conds,
		Cells:      cells,
	}
	for _, s := range strings.Split(rd["children"].(string), ",") {
		cleanStr := strings.Trim(s, "\n\t\r ")
		if cleanStr != "" {
			readRow.Children = append(readRow.Children, cleanStr)
		}
	}
	return readRow, nil
}

func excel_read_file_config_row_cells(ctx context.Context, cellRaw interface{}) (ExcelReadRowCells, error) {
	cells := make(ExcelReadRowCells, 0)
	cellRawArr := cellRaw.([]interface{})
	for _, item := range cellRawArr {
		cri := item.(map[string]interface{})
		cells = append(cells, &ExcelReadRowCell{
			Col: cri["col"].(string),
			Tag: cri["tag"].(string),
		})
	}
	return cells, nil
}

func excel_read_file_config_row_conds(ctx context.Context, condRaw interface{}) (ExcelReadRowConditions, error) {
	conds := make(ExcelReadRowConditions, 0)
	if condRaw == nil {
		return conds, nil
	}
	condRawArr := condRaw.([]interface{})
	for _, itemRaw := range condRawArr {
		rj := itemRaw.(map[string]interface{})
		conds = append(conds, &ExcelReadRowCondition{
			Col:     rj["col"].(string),
			Pattern: rj["pattern"].(string),
		})
	}
	return conds, nil
}

func excel_read_file_config_sheet_conds(ctx context.Context, condRaw interface{}) (ExcelReadSheetConditions, error) {
	conds := make(ExcelReadSheetConditions, 0)
	if condRaw == nil {
		return conds, nil
	}
	condRawArr := condRaw.([]interface{})
	for _, itemRaw := range condRawArr {
		rj := itemRaw.(map[string]interface{})
		if rj["low"].(int) == 0 {
			conds = append(conds, &ExcelReadSheetCondPattern{
				Pattern: rj["pattern"].(string),
			})
		} else if rj["low"].(int) != 0 {
			conds = append(conds, &ExcelReadSheetCondRange{
				Low:  rj["low"].(int),
				High: rj["high"].(int),
			})
		}

	}
	return conds, nil
}

func excel_read_file_config_sheet(ctx context.Context, sheetRaw interface{}) (*ExcelReadSheet, error) {
	if sheetRaw == nil {
		return nil, fmt.Errorf("required data")
	}
	sr := sheetRaw.(map[string]interface{})
	conds, err := excel_read_file_config_sheet_conds(ctx, sr["when"])
	if err != nil {
		return nil, err
	}
	sheet := &ExcelReadSheet{
		Row:        sr["row"].(string),
		Conditions: conds,
	}

	return sheet, nil
}

// defines recursive structure
func _excel_row_element() map[string]*schema.Schema {
	rowType := map[string]*schema.Schema{
		"name":  {Type: schema.TypeString, Required: true},
		"index": {Type: schema.TypeInt, Required: true},
		"cols": {
			Type: schema.TypeList,
			Elem: map[string]*schema.Schema{
				"col":   {Type: schema.TypeString, Required: true},
				"tag":   {Type: schema.TypeString, Optional: true, DefaultValue: ""},
				"value": {Type: schema.TypeGeneric, Required: true},
			},
		},
		"children": {
			Type: schema.TypeList,
			Elem: map[string]*schema.Schema{
				"name": {Type: schema.TypeString, Required: true},
			},
		},
	}
	rowType["children"].Elem["rows"] = &schema.Schema{
		Type: schema.TypeList,
		Elem: rowType,
	}
	return rowType
}
