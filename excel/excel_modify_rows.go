package excel

import (
	"context"
	"fmt"
	"regexp"
	"strings"

	"sbl.systems/go/synwork/plugin-sdk/schema"
	"sbl.systems/go/synwork/plugin-sdk/utils"
)

// https://xuri.me/excelize/en/cell.html#SetCellStyle
var (
	excel_modify = map[string]*schema.Schema{
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
		"sheet": {
			Type: schema.TypeList,
			Elem: map[string]*schema.Schema{
				"apply-rules": {Type: schema.TypeString, Required: true},
				"when": {
					Type: schema.TypeList,
					Elem: map[string]*schema.Schema{
						"name": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
						"low":  {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
						"high": {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
					},
				},
			},
		},
		"add-cell": {
			Type: schema.TypeList,
			Elem: map[string]*schema.Schema{
				"rule-name": {Type: schema.TypeString, Required: true},
				"row-type":  {Type: schema.TypeString, Optional: true, DefaultValue: ".*"},
				"when": {
					Type: schema.TypeList,
					Elem: map[string]*schema.Schema{
						"col":     {Type: schema.TypeString, Optional: true, DefaultValue: ""},
						"tag":     {Type: schema.TypeString, Optional: true, DefaultValue: ""},
						"pattern": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
					},
				},
				"new-cell": {
					Type: schema.TypeMap,
					Elem: map[string]*schema.Schema{
						"col": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
						"tag": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
					},
				},
				"from-cell": {
					Type: schema.TypeList,
					Elem: map[string]*schema.Schema{
						"col": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
						"tag": {Type: schema.TypeString, Optional: true, DefaultValue: ""},
					},
				},
				"expr": {
					Type: schema.TypeMap,
					Elem: map[string]*schema.Schema{
						"pattern-extract": {
							Type: schema.TypeMap,
							Elem: map[string]*schema.Schema{
								"pattern": {Type: schema.TypeString, Optional: true, DefaultValue: "(.*)"},
								"group":   {Type: schema.TypeInt, Optional: true, DefaultValue: 0},
							},
						},
					},
				},
			},
		},
	}
)

type (
	ExcelModConfiguration struct {
		Sheets ExcelDataSheets
		Sheet  ExcelModifySheets
	}
)

var Method_modify_rows = &schema.Method{
	Schema:   excel_modify,
	Result:   Method_read_file.Result,
	ExecFunc: excel_modify_rows,
	Description: `Method excel_modify_rows provides a way to change row structures.

	method "excel_modify_rows" "processor-instance" "method-instance" {
		sheets = $method.excel_read.sheets
		sheet {
			when {
				name = "pattern"
			}
			when {
				low = 1
				high = 2
			}
			apply-rules = "rule1,rule2"
		}
		add-cell {
			rule-name = "rule1"
			row-type = "standard"
			when {
				col = "D" // tag = "tag1"
				pattern = ""
			}
			new-cell {
				col = "AA"
				tag = "new"
			}
			from-cell {
				col = "A"
				tag = "tag1"
				index = 1
			}

			pattern-extract {
				pattern = ""
				group = 1
			}
			fix-value {
				value = ""
			}
		}
	}

	
	`,
}

type (
	ExcelModifySheet struct {
		ApplyRules string
		When       ExcelModifySheetConditions `snw:"when"`
	}
	ExcelModifySheets  []*ExcelModifySheet
	ExcelModifyAddCell struct {
		RuleName string
		RowType  string
		When     ExcelModifyCellConds
		NewCell  ExcelModifyCell
		FromCell ExcelModifyCells
		Expr     ExcelModifyExprStruct
	}
	ExcelModifyCellCond struct {
		Col     string
		Tag     string
		Pattern string
	}
	ExcelModifyCellConds []*ExcelModifyCellCond
	ExcelModifyCell      struct {
		Col   string
		Tag   string
		Index int
	}
	ExcelModifyCells      []*ExcelModifyCell
	ExcelModifyExpression interface {
		Eval([]interface{}) (interface{}, error)
	}
	ExcelModifyExprStruct struct {
		Expression ExcelModifyExpression
	}
	ExcelModifyConfiguration struct {
		Sheets ExcelModifySheets
		Rules  map[string]*ExcelModifyAddCell
		Data   ExcelDataSheets
	}
	ExcelModifySheetConditions []ExcelModifySheetCondition

	ExcelModifySheetCondition struct {
		Name      string
		Low, High int
	}
	ExcelModifyPatternExtract struct {
		Pattern string
		Group   int
		pRegexp *regexp.Regexp
	}
)

var exprFactories = map[string]func(v map[string]interface{}) (ExcelModifyExpression, error){
	"pattern-extract": func(v map[string]interface{}) (ExcelModifyExpression, error) {
		e := &ExcelModifyPatternExtract{
			Pattern: v["pattern"].(string),
			Group:   v["group"].(int),
		}
		if p, err := regexp.Compile(e.Pattern); err != nil {
			return nil, err
		} else {
			e.pRegexp = p
		}
		return e, nil
	},
}

func (e *ExcelModifyPatternExtract) Eval(v []interface{}) (interface{}, error) {
	if v == nil || len(v) == 0 {
		return nil, fmt.Errorf("pattern-extract.Eval() missing required parameter")
	}
	if str, ok := v[0].(string); !ok {
		return nil, fmt.Errorf("pattern-extract.Eval() missing string parameter")
	} else if parts := e.pRegexp.FindAllStringSubmatch(str, -1); len(parts) == 0 {
		return "", nil
	} else if 0 < e.Group && e.Group < len(parts[0]) {
		return parts[0][e.Group], nil
	} else {
		return nil, fmt.Errorf("pattern-extract.Eval() invalid group %d", e.Group)
	}
}

func (e *ExcelModifyExprStruct) UnmarshallStruct(v interface{}) error {
	j := v.(map[string]interface{})
	for key, factory := range exprFactories {
		if kv, ok := j[key]; ok {
			if jkv, ok := kv.(map[string]interface{}); ok {
				if expr, err := factory(jkv); err != nil {
					return err
				} else {
					e.Expression = expr
					return nil
				}
			}
		}
	}
	return fmt.Errorf("no expr found")
}

func (e ExcelModifySheetConditions) Test(name string, index int) (bool, error) {
	for _, cond := range e {
		if ok, err := cond.Test(name, index); err != nil {
			return false, err
		} else if !ok {
			return false, nil
		}
	}
	return true, nil
}

func (e *ExcelModifySheetCondition) Test(name string, index int) (bool, error) {
	if e.Name != "" {
		if pattern, err := regexp.Compile(e.Name); err != nil {
			return false, err
		} else {
			return pattern.MatchString(name), nil
		}
	} else {
		return e.Low > 0 && e.Low <= index+1 && index+1 <= e.High, nil
	}
}

func excel_modify_rows(ctx context.Context, data *schema.MethodData, client interface{}) error {
	config, err := excel_modify_rows_config(ctx, data)
	if err != nil {
		return err
	}
	result := ExcelDataSheets{}
	for _, sheet := range config.Data {
		for _, sheetCond := range config.Sheets {
			if ok, err := sheetCond.When.Test(sheet.Name, sheet.Index); err != nil {
				return err
			} else if ok {
				if newRows, err := excel_modify_rows_apply(config, sheet.Rows, strings.Split(sheetCond.ApplyRules, ",")); err != nil {
					return err
				} else if newRows != nil {
					result = append(result, &ExcelDataSheet{
						Name:  sheet.Name,
						Index: sheet.Index,
						Rows:  newRows,
					})
				}
				goto goto_next
			}
		}
		result = append(result, sheet)
	goto_next:
	}
	sheets := []interface{}{}
	for _, item := range result {
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

func excel_modify_rows_apply(config *ExcelModifyConfiguration, rows ExcelDataRows, rules []string) (ExcelDataRows, error) {
	result := ExcelDataRows{}
	for _, row := range rows {
		if newRow, err := excel_modify_rows_apply_row(config, row, rules); err != nil {
			return nil, err
		} else {
			result = append(result, newRow)
		}
	}
	return result, nil
}
func excel_modify_rows_apply_row(config *ExcelModifyConfiguration, row *ExcelDataRow, rules []string) (*ExcelDataRow, error) {
	newRow := &ExcelDataRow{
		Name:     row.Name,
		Index:    row.Index,
		Cols:     make(ExcelDataCols, 0),
		Children: make(ExcelDataChildren, 0),
	}
	for _, child := range row.Children {
		newChild := &ExcelDataChild{
			Name: child.Name,
		}
		if childRows, err := excel_modify_rows_apply(config, child.Rows, rules); err != nil {
			return nil, err
		} else {
			newChild.Rows = childRows
			newRow.Children = append(newRow.Children, newChild)
		}
		newRow.Children = append(newRow.Children, newChild)
	}
	newRow.Cols = append(newRow.Cols, row.Cols...)
	for _, ruleName := range utils.MapArray[string, string](rules, []string{}, strings.TrimSpace) {
		if rule, ok := config.Rules[ruleName]; ok {
			if p, err := regexp.Compile(rule.RowType); err != nil {
				return nil, err
			} else if !p.MatchString(newRow.Name) {
				continue
			}
			values := []interface{}{}
			for _, sourceField := range rule.FromCell {
				for _, col := range newRow.Cols {
					if (col.Col == sourceField.Col && col.Col != "") || (col.Tag == sourceField.Tag && col.Tag != "") {
						values = append(values, col.Value)
					}
				}
			}
			if rule.Expr.Expression == nil {
				return nil, fmt.Errorf("expr missed in rule %s", ruleName)
			}
			if newValue, err := rule.Expr.Expression.Eval(values); err != nil {
				return nil, err
			} else {
				newRow.Cols = append(newRow.Cols, &ExcelDataCol{
					Col:   rule.NewCell.Col,
					Tag:   rule.NewCell.Tag,
					Value: newValue,
				})
			}
		}
	}
	return newRow, nil
}

func excel_modify_rows_config(ctx context.Context, data *schema.MethodData) (*ExcelModifyConfiguration, error) {
	config := &ExcelModifyConfiguration{
		Rules: make(map[string]*ExcelModifyAddCell),
	}
	err := utils.NewDecoder().Decode(&config.Data, data.GetConfig("sheets"))
	if err != nil {
		return nil, err
	}
	err = utils.NewDecoder().Decode(&config.Sheets, data.GetConfig("sheet"))
	if err != nil {
		return nil, err
	}
	cells := make([]ExcelModifyAddCell, 0)
	err = utils.NewDecoder().Decode(&cells, data.GetConfig("add-cell"))
	if err != nil {
		return nil, err
	}
	for _, cell := range cells {
		config.Rules[cell.RuleName] = &cell
	}
	return config, nil
}
