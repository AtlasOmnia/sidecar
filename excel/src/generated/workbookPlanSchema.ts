/* Generated from the repository-authored WorkbookPlan V1 contract. Do not edit. */
export const workbookPlanSchema = {
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "title": "WorkbookPlan V1",
  "description": "Static model-facing contract for bounded, non-destructive Excel plans.",
  "type": "object",
  "required": [
    "version",
    "outcome"
  ],
  "additionalProperties": false,
  "properties": {
    "version": {
      "type": "integer",
      "const": 1,
      "description": "Plan schema version. Only version 1 is valid."
    },
    "outcome": {
      "type": "string",
      "enum": [
        "inspect",
        "execute",
        "ask",
        "refuse"
      ],
      "description": "The plan outcome discriminator."
    },
    "reason": {
      "type": "string",
      "minLength": 1,
      "maxLength": 1024
    },
    "summary": {
      "type": "string",
      "minLength": 1,
      "maxLength": 2048
    },
    "message": {
      "type": "string",
      "minLength": 1,
      "maxLength": 2048
    },
    "reads": {
      "type": "array",
      "minItems": 1,
      "maxItems": 50,
      "items": {
        "$ref": "#/$defs/InspectionOperation"
      }
    },
    "operations": {
      "type": "array",
      "minItems": 1,
      "maxItems": 30,
      "items": {
        "$ref": "#/$defs/WorkbookOperation"
      }
    },
    "verification": {
      "type": "array",
      "minItems": 1,
      "maxItems": 30,
      "items": {
        "$ref": "#/$defs/VerificationRule"
      }
    }
  },
  "oneOf": [
    {
      "type": "object",
      "required": [
        "version",
        "outcome",
        "reason",
        "reads"
      ],
      "additionalProperties": false,
      "properties": {
        "version": {
          "type": "integer",
          "const": 1
        },
        "outcome": {
          "const": "inspect"
        },
        "reason": {
          "type": "string",
          "minLength": 1,
          "maxLength": 1024
        },
        "reads": {
          "type": "array",
          "minItems": 1,
          "maxItems": 50,
          "items": {
            "$ref": "#/$defs/InspectionOperation"
          }
        }
      }
    },
    {
      "type": "object",
      "required": [
        "version",
        "outcome",
        "summary",
        "operations",
        "verification"
      ],
      "additionalProperties": false,
      "properties": {
        "version": {
          "type": "integer",
          "const": 1
        },
        "outcome": {
          "const": "execute"
        },
        "summary": {
          "type": "string",
          "minLength": 1,
          "maxLength": 2048
        },
        "operations": {
          "type": "array",
          "minItems": 1,
          "maxItems": 30,
          "items": {
            "$ref": "#/$defs/WorkbookOperation"
          }
        },
        "verification": {
          "type": "array",
          "minItems": 1,
          "maxItems": 30,
          "items": {
            "$ref": "#/$defs/VerificationRule"
          }
        }
      }
    },
    {
      "type": "object",
      "required": [
        "version",
        "outcome",
        "message"
      ],
      "additionalProperties": false,
      "properties": {
        "version": {
          "type": "integer",
          "const": 1
        },
        "outcome": {
          "type": "string",
          "enum": [
            "ask",
            "refuse"
          ]
        },
        "message": {
          "type": "string",
          "minLength": 1,
          "maxLength": 2048
        }
      }
    }
  ],
  "$defs": {
    "ScalarValue": {
      "oneOf": [
        {
          "type": "string",
          "maxLength": 32767
        },
        {
          "type": "number"
        },
        {
          "type": "boolean"
        },
        {
          "type": "null"
        }
      ]
    },
    "ValueRow": {
      "type": "array",
      "minItems": 1,
      "maxItems": 1000,
      "items": {
        "$ref": "#/$defs/ScalarValue"
      }
    },
    "ValueMatrix": {
      "type": "array",
      "minItems": 1,
      "maxItems": 1000,
      "items": {
        "$ref": "#/$defs/ValueRow"
      }
    },
    "FormulaRow": {
      "type": "array",
      "minItems": 1,
      "maxItems": 1000,
      "items": {
        "type": "string",
        "maxLength": 4096
      }
    },
    "FormulaMatrix": {
      "type": "array",
      "minItems": 1,
      "maxItems": 1000,
      "items": {
        "$ref": "#/$defs/FormulaRow"
      }
    },
    "FormatSpec": {
      "type": "object",
      "required": [],
      "additionalProperties": false,
      "properties": {
        "bold": {
          "type": "boolean"
        },
        "italic": {
          "type": "boolean"
        },
        "fontColor": {
          "type": "string",
          "minLength": 1,
          "maxLength": 64
        },
        "fillColor": {
          "type": "string",
          "minLength": 1,
          "maxLength": 64
        },
        "numberFormat": {
          "type": "string",
          "minLength": 1,
          "maxLength": 128
        },
        "borders": {
          "type": "boolean"
        },
        "columnAutofit": {
          "type": "boolean"
        },
        "rowAutofit": {
          "type": "boolean"
        },
        "horizontalAlignment": {
          "type": "string",
          "enum": [
            "General",
            "Left",
            "Center",
            "Right",
            "Fill",
            "Justify",
            "CenterAcrossSelection",
            "Distributed"
          ]
        },
        "verticalAlignment": {
          "type": "string",
          "minLength": 1,
          "maxLength": 32
        },
        "fontName": {
          "type": "string",
          "minLength": 1,
          "maxLength": 128
        },
        "fontSize": {
          "type": "number",
          "exclusiveMinimum": 0
        },
        "underline": {
          "type": "string",
          "enum": [
            "None",
            "Single",
            "Double"
          ]
        },
        "wrapText": {
          "type": "boolean"
        },
        "indentLevel": {
          "type": "integer",
          "minimum": 0
        },
        "textOrientation": {
          "type": "integer",
          "minimum": -90,
          "maximum": 90
        },
        "borderColor": {
          "type": "string",
          "minLength": 1,
          "maxLength": 64
        },
        "borderWeight": {
          "type": "string",
          "enum": [
            "Hairline",
            "Thin",
            "Medium",
            "Thick"
          ]
        }
      }
    },
    "InspectionOperation": {
      "type": "object",
      "required": [
        "kind"
      ],
      "additionalProperties": false,
      "properties": {
        "kind": {
          "type": "string",
          "enum": [
            "read_range",
            "summarize_range_stats"
          ]
        },
        "sheet": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255
        },
        "address": {
          "type": "string",
          "minLength": 1,
          "maxLength": 256
        }
      },
      "oneOf": [
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "read_range"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "summarize_range_stats"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            }
          }
        }
      ]
    },
    "WorkbookOperation": {
      "type": "object",
      "required": [
        "kind"
      ],
      "additionalProperties": false,
      "properties": {
        "kind": {
          "type": "string",
          "enum": [
            "write_values",
            "write_formula",
            "fill_formula",
            "format_range",
            "create_table",
            "sort_table",
            "filter_table",
            "reset_table_filter",
            "create_chart",
            "set_validation_list",
            "conditional_format",
            "freeze_panes",
            "create_sheet",
            "rename_sheet",
            "copy_sheet",
            "add_comment",
            "calculate"
          ]
        },
        "sheet": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255
        },
        "address": {
          "type": "string",
          "minLength": 1,
          "maxLength": 256
        },
        "values": {
          "$ref": "#/$defs/ValueMatrix"
        },
        "formula": {
          "type": "string",
          "minLength": 1,
          "maxLength": 4096
        },
        "sourceAddress": {
          "type": "string",
          "minLength": 1,
          "maxLength": 256
        },
        "targetAddress": {
          "type": "string",
          "minLength": 1,
          "maxLength": 256
        },
        "format": {
          "$ref": "#/$defs/FormatSpec"
        },
        "hasHeaders": {
          "type": "boolean"
        },
        "tableName": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255
        },
        "columnIndex": {
          "type": "integer",
          "minimum": 0
        },
        "ascending": {
          "type": "boolean"
        },
        "criterion1": {
          "type": "string",
          "maxLength": 512
        },
        "criterion2": {
          "type": "string",
          "maxLength": 512
        },
        "operator": {
          "type": "string",
          "enum": [
            "And",
            "Or"
          ]
        },
        "dataRange": {
          "type": "string",
          "minLength": 1,
          "maxLength": 256
        },
        "chartType": {
          "type": "string",
          "minLength": 1,
          "maxLength": 128
        },
        "title": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255
        },
        "source": {
          "type": "string",
          "minLength": 1,
          "maxLength": 4096
        },
        "promptTitle": {
          "type": "string",
          "maxLength": 255
        },
        "promptMessage": {
          "type": "string",
          "maxLength": 1024
        },
        "ruleType": {
          "type": "string",
          "enum": [
            "cellValue",
            "containsText"
          ]
        },
        "formula1": {
          "type": "string",
          "maxLength": 4096
        },
        "formula2": {
          "type": "string",
          "maxLength": 4096
        },
        "text": {
          "type": "string",
          "maxLength": 1024
        },
        "fillColor": {
          "type": "string",
          "minLength": 1,
          "maxLength": 64
        },
        "fontColor": {
          "type": "string",
          "minLength": 1,
          "maxLength": 64
        },
        "rows": {
          "type": "integer",
          "minimum": 0
        },
        "columns": {
          "type": "integer",
          "minimum": 0
        },
        "name": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255
        },
        "oldName": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255
        },
        "newName": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255
        },
        "sourceName": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255
        }
      },
      "oneOf": [
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "values"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "write_values"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "values": {
              "$ref": "#/$defs/ValueMatrix"
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "formula"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "write_formula"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "formula": {
              "type": "string",
              "minLength": 1,
              "maxLength": 4096
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "sourceAddress",
            "targetAddress"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "fill_formula"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "sourceAddress": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "targetAddress": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "format"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "format_range"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "format": {
              "$ref": "#/$defs/FormatSpec"
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "hasHeaders"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "create_table"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "hasHeaders": {
              "type": "boolean"
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "tableName",
            "columnIndex"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "sort_table"
            },
            "tableName": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "columnIndex": {
              "type": "integer",
              "minimum": 0
            },
            "ascending": {
              "type": "boolean"
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "tableName",
            "columnIndex"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "filter_table"
            },
            "tableName": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "columnIndex": {
              "type": "integer",
              "minimum": 0
            },
            "values": {
              "type": "array",
              "minItems": 1,
              "maxItems": 200,
              "items": {
                "type": "string",
                "minLength": 1,
                "maxLength": 512
              }
            },
            "criterion1": {
              "type": "string",
              "maxLength": 512
            },
            "criterion2": {
              "type": "string",
              "maxLength": 512
            },
            "operator": {
              "type": "string",
              "enum": [
                "And",
                "Or"
              ]
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "tableName",
            "columnIndex"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "reset_table_filter"
            },
            "tableName": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "columnIndex": {
              "type": "integer",
              "minimum": 0
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "dataRange",
            "chartType",
            "title"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "create_chart"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "dataRange": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "chartType": {
              "type": "string",
              "minLength": 1,
              "maxLength": 128
            },
            "title": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "source"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "set_validation_list"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "source": {
              "type": "string",
              "minLength": 1,
              "maxLength": 4096
            },
            "promptTitle": {
              "type": "string",
              "maxLength": 255
            },
            "promptMessage": {
              "type": "string",
              "maxLength": 1024
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "ruleType"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "conditional_format"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "ruleType": {
              "type": "string",
              "enum": [
                "cellValue",
                "containsText"
              ]
            },
            "operator": {
              "type": "string",
              "maxLength": 64
            },
            "formula1": {
              "type": "string",
              "maxLength": 4096
            },
            "formula2": {
              "type": "string",
              "maxLength": 4096
            },
            "text": {
              "type": "string",
              "maxLength": 1024
            },
            "fillColor": {
              "type": "string",
              "minLength": 1,
              "maxLength": 64
            },
            "fontColor": {
              "type": "string",
              "minLength": 1,
              "maxLength": 64
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "rows",
            "columns"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "freeze_panes"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "rows": {
              "type": "integer",
              "minimum": 0
            },
            "columns": {
              "type": "integer",
              "minimum": 0
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "name"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "create_sheet"
            },
            "name": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "oldName",
            "newName"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "rename_sheet"
            },
            "oldName": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "newName": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sourceName",
            "newName"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "copy_sheet"
            },
            "sourceName": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "newName": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "text"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "add_comment"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "text": {
              "type": "string",
              "minLength": 1,
              "maxLength": 1024
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "calculate"
            }
          }
        }
      ]
    },
    "VerificationRule": {
      "type": "object",
      "required": [
        "kind"
      ],
      "additionalProperties": false,
      "properties": {
        "kind": {
          "type": "string",
          "enum": [
            "range_values_equal",
            "range_formulas_equal",
            "range_formula_pattern",
            "range_number_format_equal",
            "table_exists",
            "table_row_order",
            "table_filter_state",
            "sheet_exists",
            "chart_exists",
            "comment_exists",
            "sentinel_unchanged",
            "calculation_completed"
          ]
        },
        "sheet": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255
        },
        "address": {
          "type": "string",
          "minLength": 1,
          "maxLength": 256
        },
        "expectedValues": {
          "$ref": "#/$defs/ValueMatrix"
        },
        "expectedFormulas": {
          "$ref": "#/$defs/FormulaMatrix"
        },
        "pattern": {
          "type": "string",
          "minLength": 1,
          "maxLength": 4096
        },
        "expectedNumberFormat": {
          "type": "string",
          "minLength": 1,
          "maxLength": 128
        },
        "tableName": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255
        },
        "expectedOrder": {
          "type": "array",
          "maxItems": 1000,
          "items": {
            "type": "string",
            "minLength": 1,
            "maxLength": 255
          }
        },
        "columnIndex": {
          "type": "integer",
          "minimum": 0
        },
        "state": {
          "type": "string",
          "enum": [
            "filtered",
            "unfiltered"
          ]
        },
        "values": {
          "type": "array",
          "maxItems": 200,
          "items": {
            "type": "string",
            "minLength": 1,
            "maxLength": 512
          }
        },
        "criterion1": {
          "type": "string",
          "maxLength": 512
        },
        "criterion2": {
          "type": "string",
          "maxLength": 512
        },
        "operator": {
          "type": "string",
          "enum": [
            "And",
            "Or"
          ]
        },
        "name": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255
        },
        "title": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255
        },
        "text": {
          "type": "string",
          "minLength": 1,
          "maxLength": 1024
        },
        "expected": {
          "$ref": "#/$defs/ScalarValue"
        }
      },
      "oneOf": [
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "expectedValues"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "range_values_equal"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "expectedValues": {
              "$ref": "#/$defs/ValueMatrix"
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "expectedFormulas"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "range_formulas_equal"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "expectedFormulas": {
              "$ref": "#/$defs/FormulaMatrix"
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "pattern"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "range_formula_pattern"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "pattern": {
              "type": "string",
              "minLength": 1,
              "maxLength": 4096
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "expectedNumberFormat"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "range_number_format_equal"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "expectedNumberFormat": {
              "type": "string",
              "minLength": 1,
              "maxLength": 128
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "tableName"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "table_exists"
            },
            "tableName": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "tableName",
            "expectedOrder"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "table_row_order"
            },
            "tableName": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "expectedOrder": {
              "type": "array",
              "minItems": 1,
              "maxItems": 1000,
              "items": {
                "type": "string",
                "minLength": 1,
                "maxLength": 255
              }
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "tableName",
            "columnIndex",
            "state"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "table_filter_state"
            },
            "tableName": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "columnIndex": {
              "type": "integer",
              "minimum": 0
            },
            "state": {
              "type": "string",
              "enum": [
                "filtered",
                "unfiltered"
              ]
            },
            "values": {
              "type": "array",
              "maxItems": 200,
              "items": {
                "type": "string",
                "minLength": 1,
                "maxLength": 512
              }
            },
            "criterion1": {
              "type": "string",
              "maxLength": 512
            },
            "criterion2": {
              "type": "string",
              "maxLength": 512
            },
            "operator": {
              "type": "string",
              "enum": [
                "And",
                "Or"
              ]
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "name"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "sheet_exists"
            },
            "name": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "title"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "chart_exists"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "title": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "text"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "comment_exists"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "text": {
              "type": "string",
              "minLength": 1,
              "maxLength": 1024
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind",
            "sheet",
            "address",
            "expected"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "sentinel_unchanged"
            },
            "sheet": {
              "type": "string",
              "minLength": 1,
              "maxLength": 255
            },
            "address": {
              "type": "string",
              "minLength": 1,
              "maxLength": 256
            },
            "expected": {
              "$ref": "#/$defs/ScalarValue"
            }
          }
        },
        {
          "type": "object",
          "required": [
            "kind"
          ],
          "additionalProperties": false,
          "properties": {
            "kind": {
              "const": "calculation_completed"
            }
          }
        }
      ]
    }
  }
} as const;

export const workbookPlanFingerprint = "5fcab038f26bef8423d0ec2271b283947c87256bfe69f250205ea348cbfc83c1" as const;
export const WORKBOOK_PLAN_FINGERPRINT = workbookPlanFingerprint;

export type ScalarValue = string | number | boolean | null;
export type ValueRow = ReadonlyArray<ScalarValue>;
export type ValueMatrix = ReadonlyArray<ValueRow>;
export type FormulaRow = ReadonlyArray<string>;
export type FormulaMatrix = ReadonlyArray<FormulaRow>;
export type FormatSpec = { readonly bold?: boolean; readonly italic?: boolean; readonly fontColor?: string; readonly fillColor?: string; readonly numberFormat?: string; readonly borders?: boolean; readonly columnAutofit?: boolean; readonly rowAutofit?: boolean; readonly horizontalAlignment?: "General" | "Left" | "Center" | "Right" | "Fill" | "Justify" | "CenterAcrossSelection" | "Distributed"; readonly verticalAlignment?: string; readonly fontName?: string; readonly fontSize?: number; readonly underline?: "None" | "Single" | "Double"; readonly wrapText?: boolean; readonly indentLevel?: number; readonly textOrientation?: number; readonly borderColor?: string; readonly borderWeight?: "Hairline" | "Thin" | "Medium" | "Thick"; };
export type InspectionOperation = { readonly kind: "read_range"; readonly sheet: string; readonly address: string; } | { readonly kind: "summarize_range_stats"; readonly sheet: string; readonly address: string; };
export type WorkbookOperation = { readonly kind: "write_values"; readonly sheet: string; readonly address: string; readonly values: ValueMatrix; } | { readonly kind: "write_formula"; readonly sheet: string; readonly address: string; readonly formula: string; } | { readonly kind: "fill_formula"; readonly sheet: string; readonly sourceAddress: string; readonly targetAddress: string; } | { readonly kind: "format_range"; readonly sheet: string; readonly address: string; readonly format: FormatSpec; } | { readonly kind: "create_table"; readonly sheet: string; readonly address: string; readonly hasHeaders: boolean; } | { readonly kind: "sort_table"; readonly tableName: string; readonly columnIndex: number; readonly ascending?: boolean; } | { readonly kind: "filter_table"; readonly tableName: string; readonly columnIndex: number; readonly values?: ReadonlyArray<string>; readonly criterion1?: string; readonly criterion2?: string; readonly operator?: "And" | "Or"; } | { readonly kind: "reset_table_filter"; readonly tableName: string; readonly columnIndex: number; } | { readonly kind: "create_chart"; readonly sheet: string; readonly dataRange: string; readonly chartType: string; readonly title: string; } | { readonly kind: "set_validation_list"; readonly sheet: string; readonly address: string; readonly source: string; readonly promptTitle?: string; readonly promptMessage?: string; } | { readonly kind: "conditional_format"; readonly sheet: string; readonly address: string; readonly ruleType: "cellValue" | "containsText"; readonly operator?: string; readonly formula1?: string; readonly formula2?: string; readonly text?: string; readonly fillColor?: string; readonly fontColor?: string; } | { readonly kind: "freeze_panes"; readonly sheet: string; readonly rows: number; readonly columns: number; } | { readonly kind: "create_sheet"; readonly name: string; } | { readonly kind: "rename_sheet"; readonly oldName: string; readonly newName: string; } | { readonly kind: "copy_sheet"; readonly sourceName: string; readonly newName: string; } | { readonly kind: "add_comment"; readonly sheet: string; readonly address: string; readonly text: string; } | { readonly kind: "calculate"; };
export type VerificationRule = { readonly kind: "range_values_equal"; readonly sheet: string; readonly address: string; readonly expectedValues: ValueMatrix; } | { readonly kind: "range_formulas_equal"; readonly sheet: string; readonly address: string; readonly expectedFormulas: FormulaMatrix; } | { readonly kind: "range_formula_pattern"; readonly sheet: string; readonly address: string; readonly pattern: string; } | { readonly kind: "range_number_format_equal"; readonly sheet: string; readonly address: string; readonly expectedNumberFormat: string; } | { readonly kind: "table_exists"; readonly tableName: string; } | { readonly kind: "table_row_order"; readonly tableName: string; readonly expectedOrder: ReadonlyArray<string>; } | { readonly kind: "table_filter_state"; readonly tableName: string; readonly columnIndex: number; readonly state: "filtered" | "unfiltered"; readonly values?: ReadonlyArray<string>; readonly criterion1?: string; readonly criterion2?: string; readonly operator?: "And" | "Or"; } | { readonly kind: "sheet_exists"; readonly name: string; } | { readonly kind: "chart_exists"; readonly sheet: string; readonly title: string; } | { readonly kind: "comment_exists"; readonly sheet: string; readonly address: string; readonly text: string; } | { readonly kind: "sentinel_unchanged"; readonly sheet: string; readonly address: string; readonly expected: ScalarValue; } | { readonly kind: "calculation_completed"; };

export type WorkbookPlan = { readonly version: 1; readonly outcome: "inspect"; readonly reason: string; readonly reads: ReadonlyArray<InspectionOperation>; } | { readonly version: 1; readonly outcome: "execute"; readonly summary: string; readonly operations: ReadonlyArray<WorkbookOperation>; readonly verification: ReadonlyArray<VerificationRule>; } | { readonly version: 1; readonly outcome: "ask" | "refuse"; readonly message: string; };
