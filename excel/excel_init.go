package excel

import (
	"context"

	"sbl.systems/go/synwork/plugin-sdk/schema"
)

func excel_initfunc(ctx context.Context, ma *schema.ObjectData, client interface{}) (interface{}, error) {
	return nil, nil
}

func toString(v interface{}) string {
	return v.(string)
}
