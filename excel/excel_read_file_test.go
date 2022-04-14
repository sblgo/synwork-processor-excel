package excel

import (
	"testing"

	"sbl.systems/go/synwork/plugin-sdk/tunit"
)

func TestReadExcelFile01(t *testing.T) {
	_defs := `
	method "read_excel_file" "dum" "join01" {
		file-name = "read01.xlsx"
		sheet {
			when {
				low = 1
				high = 1
			}
			row = "standard"
		}

		row {
			name = "standard"
			when {
				col = "A"
				pattern = "AA"
			}
			cell {
				col = "A"
				tag = "sign"
			}
			cell {
				col = "B"
				tag = "name"
			}
			cell {
				col = "C"
				tag = "addr"
			}
			cell {
				col = "D"
				tag = "cost"
			}
			cell {
				col = "E"
				tag = "vat"
			}
		}

	}
	`
	mm := tunit.MethodMock{
		ProcessorDef: Opts.Provider,
		InstanceMock: nil,
		ExecFunc:     excel_read_file,
		References:   map[string]interface{}{},
	}
	result := tunit.CallMockMethod(t, mm, _defs)
	if len(result) != 1 {
		t.Fatal()
	}
}
