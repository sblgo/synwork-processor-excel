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
			},
			InitFunc: excel_initfunc,
		}
	},
}
