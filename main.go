package main

import (
	"sbl.system/synwork/synwork-processor-excel/excel"
	"sbl.systems/go/synwork/plugin-sdk/plugin"
)

func main() {
	plugin.Serve(excel.Opts)
}
