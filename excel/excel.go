package excel

import (
	"sbl.systems/go/synwork/plugin-sdk/plugin"
	"sbl.systems/go/synwork/plugin-sdk/schema"
)

var Opts = plugin.PluginOptions{
	Provider: func() schema.Processor {
		return schema.Processor{
			Schema: map[string]*schema.Schema{},
			MethodMap: map[string]*schema.Method{
				"write_excel_file": Method_write_file,
				"read_excel_file":  Method_read_file,
				"modify_rows":      Method_modify_rows,
			},
			InitFunc: excel_initfunc,
			Description: `Processor sbl.systems/synwork/excel provides methods to handle excel files.
			
			actual 
			  - write_excel_file
			  - read_excel_file
			`,
		}
	},
}
