# synwork-processor-excel

This project provides a [synwork.io](http://www.synwork.io) processor for creating excel files.


```
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

```

Of course, this configuration is extendable.