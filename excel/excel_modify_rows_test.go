package excel

import (
	"bytes"
	"encoding/json"
	"testing"

	"sbl.systems/go/synwork/plugin-sdk/tunit"
	"sbl.systems/go/synwork/plugin-sdk/utils"
)

func TestModifyRows01(t *testing.T) {
	_json := `{
			"read": {
				"sheets": [
					{
						"name":"sheet01",
						"index": 1,
						"rows": [
							{
								"name": "std",
								"index": 1,
								"cols": [
									{
										"col": "A",
										"tag": "sign",
										"value": "€ Bücher"
									},
									{
										"col": "B",
										"tag": "text",
										"value": "Kenne ich "
									},
									{
										"col": "C",
										"tag": "net",
										"value": 10
									},
									{
										"col": "D",
										"tag": "vat",
										"value": 19
									}
								]
							}
						]
					}
				]
			}
	}`
	type Read struct {
		Sheets ExcelDataSheets `json:"sheets"`
	}
	type Method struct {
		Read *Read `json:"read"`
	}

	d := json.NewDecoder(bytes.NewReader([]byte(_json)))
	method := &Method{}
	if err := d.Decode(method); err != nil {
		t.Fatal(err)
	}
	_jsonData, err := utils.NewEncoder().Encode(method)
	if err != nil {
		t.Fatal(err)
	}

	_defs := `
	method "modify_rows" "dum" "join01" {
		sheets = $method.read.sheets
		sheet {
			when {
				name = "sheet01"
			}
			apply-rules = "rule1"
		}

		add-cell {
			rule-name = "rule1"
			row-type = "std"
			when {
				col = "A"
				pattern = "€.*"
			}
			new-cell {
				col = "AA"
				tag = "woeuro"
			}
			from-cell {
				col = "A"
			}
			expr {
				pattern-extract {
					pattern = "€ (.*)$"
					group = 1
				}
			}
		}

	}
	`
	mm := tunit.MethodMock{
		ProcessorDef: Opts.Provider,
		InstanceMock: nil,
		ExecFunc:     excel_modify_rows,
		References: map[string]interface{}{
			"method": _jsonData,
		},
	}
	result := tunit.CallMockMethod(t, mm, _defs)
	if len(result) != 1 {
		t.Fatal()
	}
	resultSheets := &Read{}
	err = utils.NewDecoder().Decode(resultSheets, result)
	if err != nil {
		t.Fatal(err)
	}
	if len(resultSheets.Sheets[0].Rows[0].Cols) != 5 {
		t.Fatal("invalid result cols count")
	}
}
