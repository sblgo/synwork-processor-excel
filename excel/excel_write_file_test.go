package excel

import (
	"testing"

	"sbl.systems/go/synwork/plugin-sdk/tunit"
)

func TestWriteExcelFile01(t *testing.T) {
	_defs := `
	method "write_excel_file" "dum" "join01" {
		file-name = "test01.xlsx"
		sheet {
			name = "sheet01"
			cell { 
				name = "A1"
				value = "Das ist ein Field"
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
	`
	mm := tunit.MethodMock{
		ProcessorDef: Opts.Provider,
		InstanceMock: nil,
		ExecFunc:     excel_write_file,
		References:   map[string]interface{}{},
	}
	result := tunit.CallMockMethod(t, mm, _defs)
	if len(result) != 0 {
		t.Fatal()
	}
}

func TestWriteExcelFile02(t *testing.T) {
	_defs := `
	method "write_excel_file" "dum" "join01" {
		file-name = "test02.xlsx"
		sheet {
			name = "sheet01"
			cell = $method.cells
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
	`
	mm := tunit.MethodMock{
		ProcessorDef: Opts.Provider,
		InstanceMock: nil,
		ExecFunc:     excel_write_file,
		References: map[string]interface{}{
			"method": map[string]interface{}{

				"cells": []interface{}{
					map[string]interface{}{
						"name":  "A1",
						"value": "Professor",
					},
					map[string]interface{}{
						"name":  "B2",
						"value": "Professor",
						"style": "grey",
					},
					map[string]interface{}{
						"name":  "D3",
						"value": "Professor",
					},
					map[string]interface{}{
						"name":  "K2",
						"value": "Professor",
					},
				},
			},
		},
	}
	result := tunit.CallMockMethod(t, mm, _defs)
	if len(result) != 0 {
		t.Fatal()
	}
}
