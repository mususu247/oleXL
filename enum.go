package oleXL

func EnumToStrings(enum map[string]int32) []string {
	var results []string

	for k := range enum {
		results = append(results, k)
	}
	return results
}

// FileFormat
func EnumFileFormat() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlAddIn"] = 18
	enum["xlAddIn8"] = 18
	enum["xlCSV"] = 6
	enum["xlCSVMac"] = 22
	enum["xlCSVMSDOS"] = 24
	enum["xlCSVUTF8"] = 62
	enum["xlCSVWindows"] = 23
	enum["xlCurrentPlatformText"] = -4158
	enum["xlDBF2"] = 7
	enum["xlDBF3"] = 8
	enum["xlDBF4"] = 11
	enum["xlDIF"] = 9
	enum["xlExcel12"] = 50
	enum["xlExcel2"] = 16
	enum["xlExcel2FarEast"] = 27
	enum["xlExcel3"] = 29
	enum["xlExcel4"] = 33
	enum["xlExcel4Workbook"] = 35
	enum["xlExcel5"] = 39
	enum["xlExcel7"] = 39
	enum["xlExcel8"] = 56
	enum["xlExcel9795"] = 43
	enum["xlHtml"] = 44
	enum["xlIntlAddIn"] = 26
	enum["xlIntlMacro"] = 25
	enum["xlOpenDocumentSpreadsheet"] = 60
	enum["xlOpenXMLAddIn"] = 55
	enum["xlOpenXMLStrictWorkbook"] = 61
	enum["xlOpenXMLTemplate"] = 54
	enum["xlOpenXMLTemplateMacroEnabled"] = 53
	enum["xlOpenXMLWorkbook"] = 51
	enum["xlOpenXMLWorkbookMacroEnabled"] = 52
	enum["xlSYLK"] = 2
	enum["xlTemplate"] = 17
	enum["xlTemplate8"] = 17
	enum["xlTextMac"] = 19
	enum["xlTextMSDOS"] = 21
	enum["xlTextPrinter"] = 36
	enum["xlTextWindows"] = 20
	enum["xlUnicodeText"] = 42
	enum["xlWebArchive"] = 45
	enum["xlWJ2WD1"] = 14
	enum["xlWJ3"] = 40
	enum["xlWJ3FJ3"] = 41
	enum["xlWK1"] = 5
	enum["xlWK1ALL"] = 31
	enum["xlWK1FMT"] = 30
	enum["xlWK3"] = 15
	enum["xlWK3FM3"] = 32
	enum["xlWK4"] = 38
	enum["xlWKS"] = 4
	enum["xlWorkbookDefault"] = 51 //default
	enum["xlWorkbookNormal"] = -4143
	enum["xlWorks2FarEast"] = 28
	enum["xlWQ1"] = 34
	enum["xlXMLSpreadsheet"] = 46

	return enum
}

func GetEnumFileFormatNum(enumType string) int32 {
	var result int32
	enum := EnumFileFormat()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlWorkbookDefault"]
	}
	return result
}

func GetEnumFileFormatStr(enumNum int32) string {
	var result string
	enum := EnumFileFormat()
	result = "xlWorkbookDefault"

	for k, v := range enum {
		if v == enumNum {
			result = k
			break
		}
	}
	return result
}

func SetEnumFileFormat(enumNum int32) int32 {
	var result int32
	enum := EnumFileFormat()
	result = enum["xlWorkbookDefault"]

	for _, v := range enum {
		if v == enumNum {
			result = v
			break
		}
	}
	return result
}

// XlPlatform
func EnumPlatform() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlMacintosh"] = 1
	enum["xlMSDOS"] = 2
	enum["xlWindows"] = 3 //Default

	return enum
}

func GetEnumPlatformNum(enumType string) int32 {
	var result int32
	enum := EnumPlatform()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlWindows"]
	}
	return result
}

func GetEnumPlatformStr(enumNum int32) string {
	var result string
	enum := EnumPlatform()
	result = "xlWindows"

	for k, v := range enum {
		if v == enumNum {
			result = k
			break
		}
	}
	return result
}

func SetEnumPlatform(enumNum int32) int32 {
	var result int32
	enum := EnumPlatform()
	result = enum["xlWindows"]

	for _, v := range enum {
		if v == enumNum {
			result = v
			break
		}
	}
	return result
}

// XlCorruptLoad
func EnumCorruptLoad() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlExtractData"] = 2
	enum["xlNormalLoad"] = 0 //Default
	enum["xlRepairFile"] = 1

	return enum
}

func GetEnumCorruptLoadNum(enumType string) int32 {
	var result int32
	enum := EnumCorruptLoad()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlNormalLoad"]
	}
	return result
}

func GetEnumCorruptLoadStr(enumNum int32) string {
	var result string
	enum := EnumCorruptLoad()
	result = "xlNormalLoad"

	for k, v := range enum {
		if v == enumNum {
			result = k
			break
		}
	}
	return result
}

func SetEnumCorruptLoad(enumNum int32) int32 {
	var result int32
	enum := EnumCorruptLoad()
	result = enum["xlNormalLoad"]

	for _, v := range enum {
		if v == enumNum {
			result = v
			break
		}
	}
	return result
}

// Paste
func EnumPaste() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlPasteAll"] = -4104
	enum["xlPasteAllExceptBorders"] = 7
	enum["xlPasteAllMergingConditionalFormats"] = 14
	enum["xlPasteAllUsingSourceTheme"] = 13
	enum["xlPasteColumnWidths"] = 8
	enum["xlPasteComments"] = -4144
	enum["xlPasteFormats"] = -4122
	enum["xlPasteFormulas"] = -4123
	enum["xlPasteFormulasAndNumberFormats"] = 11
	enum["xlPasteValidation"] = 6
	enum["xlPasteValues"] = -4163
	enum["xlPasteValuesAndNumberFormats"] = 12

	return enum
}

func GetEnumPasteNum(enumType string) int32 {
	var result int32
	enum := EnumPaste()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlPasteAll"]
	}
	return result
}

func GetEnumPasteStr(enumNum int32) string {
	var result string
	enum := EnumPaste()
	result = "xlPasteAll"

	for k, v := range enum {
		if v == enumNum {
			result = k
			break
		}
	}
	return result
}

func SetEnumPaste(enumNum int32) int32 {
	var result int32
	enum := EnumPaste()
	result = enum["xlPasteAll"]

	for _, v := range enum {
		if v == enumNum {
			result = v
			break
		}
	}
	return result
}

// PasteSpecial
func EnumPasteOperation() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlPasteSpecialOperationAdd"] = 2
	enum["xlPasteSpecialOperationDivide"] = 5
	enum["xlPasteSpecialOperationMultiply"] = 4
	enum["xlPasteSpecialOperationNone"] = -4142
	enum["xlPasteSpecialOperationSubtract"] = 3

	return enum
}

func GetEnumPasteOperationNum(enumType string) int32 {
	var result int32
	enum := EnumPasteOperation()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlPasteSpecialOperationNone"]
	}
	return result
}

func GetEnumPasteOperationStr(enumType int32) string {
	var result string
	enum := EnumPasteOperation()
	result = "xlPasteAll"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumPasteOperation(enumType int32) int32 {
	var result int32
	enum := EnumPasteOperation()
	result = enum["xlPasteSpecialOperationNone"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Direction
func EnumDirection() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlDown"] = -4121
	enum["xlToLeft"] = -4159
	enum["xlToRight"] = -4161
	enum["xlUp"] = -4162

	return enum
}

func GetEnumDirectionNum(enumType string) int32 {
	var result int32
	enum := EnumDirection()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlShiftDown"]
	}
	return result
}

func GetEnumDirectionStr(enumType int32) string {
	var result string
	enum := EnumDirection()
	result = "xlShiftDown"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumDirection(enumType int32) int32 {
	var result int32
	enum := EnumDirection()
	result = enum["xlShiftDown"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// InsertShift
func EnumInsertShift() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlShiftDown"] = -4121
	enum["xlShiftToRight"] = -4161

	return enum
}

func GetEnumInsertShiftNum(enumType string) int32 {
	var result int32
	enum := EnumInsertShift()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlShiftDown"]
	}
	return result
}

func GetEnumInsertShiftStr(enumType int32) string {
	var result string
	enum := EnumInsertShift()
	result = "xlShiftDown"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumInsertShift(enumType int32) int32 {
	var result int32
	enum := EnumInsertShift()
	result = enum["xlShiftDown"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// WindowState
func EnumWindowState() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlMaximized"] = -4137
	enum["xlMinimized"] = -4140
	enum["xlNormal"] = -4143

	return enum
}

func GetEnumWindowStateNum(enumType string) int32 {
	var result int32
	enum := EnumWindowState()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlNormal"]
	}
	return result
}

func GetEnumWindowStateStr(enumType int32) string {
	var result string
	enum := EnumWindowState()
	result = "xlNormal"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumWindowState(enumType int32) int32 {
	var result int32
	enum := EnumWindowState()
	result = enum["xlNormal"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Calculation
func EnumCalculation() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlCalculationAutomatic"] = -4105
	enum["xlCalculationManual"] = -4135
	enum["xlCalculationSemiautomatic"] = 2

	return enum
}

func GetEnumCalculationNum(enumType string) int32 {
	var result int32
	enum := EnumCalculation()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlCalculationAutomatic"]
	}
	return result
}

func GetEnumCalculationStr(enumType int32) string {
	var result string
	enum := EnumCalculation()
	result = "xlCalculationAutomatic"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumCalculation(enumType int32) int32 {
	var result int32
	enum := EnumCalculation()
	result = enum["xlCalculationAutomatic"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Borders
func EnumBorders() map[string]int32 {
	enum := make(map[string]int32)

	enum["default"] = 0
	enum["xlDiagonalDown"] = 5
	enum["xlDiagonalUp"] = 6
	enum["xlEdgeBottom"] = 9
	enum["xlEdgeLeft"] = 7
	enum["xlEdgeRight"] = 10
	enum["xlEdgeTop"] = 8
	enum["xlInsideHorizontal"] = 12
	enum["xlInsideVertical"] = 11

	return enum
}

func GetEnumBordersNum(enumType string) int32 {
	var result int32
	enum := EnumBorders()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["default"]
	}
	return result
}

func GetEnumBordersStr(enumType int32) string {
	var result string
	enum := EnumBorders()
	result = "default"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumBorders(enumType int32) int32 {
	var result int32
	enum := EnumBorders()
	result = enum["default"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// LineStyle
func EnumLineStyle() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlContinuous"] = 1 //Def
	enum["xlDash"] = -4115
	enum["xlDashDot"] = 4
	enum["xlDashDotDot"] = 5
	enum["xlDot"] = -4118
	enum["xlDouble"] = -4119
	enum["xlLineStyleNone"] = -4142
	enum["xlSlantDashDot"] = 13

	return enum
}

func GetEnumLineStyleNum(enumType string) int32 {
	var result int32
	enum := EnumLineStyle()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlContinuous"]
	}
	return result
}

func GetEnumLineStyleStr(enumType int32) string {
	var result string
	enum := EnumLineStyle()
	result = "xlContinuous"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumLineStyle(enumType int32) int32 {
	var result int32
	enum := EnumLineStyle()
	result = enum["xlContinuous"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Weight
func EnumWeight() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlHairline"] = 1
	enum["xlMedium"] = -4138
	enum["xlThick"] = 4
	enum["xlThin"] = 2

	return enum
}

func GetEnumWeightNum(enumType string) int32 {
	var result int32
	enum := EnumWeight()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlMedium"]
	}
	return result
}

func GetEnumWeightStr(enumType int32) string {
	var result string
	enum := EnumWeight()
	result = "xlMedium"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumWeight(enumType int32) int32 {
	var result int32
	enum := EnumWeight()
	result = enum["xlMedium"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Pattern
func EnumPattern() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlPatternAutomatic"] = -4105
	enum["xlPatternChecker"] = 9
	enum["xlPatternCrissCross"] = 16
	enum["xlPatternDown"] = -4121
	enum["xlPatternGray16"] = 17
	enum["xlPatternGray25"] = -4124
	enum["xlPatternGray50"] = -4125
	enum["xlPatternGray75"] = -4126
	enum["xlPatternGray8"] = 18
	enum["xlPatternGrid"] = 15
	enum["xlPatternHorizontal"] = -4128
	enum["xlPatternLightDown"] = 13
	enum["xlPatternLightHorizontal"] = 11
	enum["xlPatternLightUp"] = 14
	enum["xlPatternLightVertical"] = 12
	enum["xlPatternNone"] = -4142
	enum["xlPatternSemiGray75"] = 10
	enum["xlPatternSolid"] = 1
	enum["xlPatternUp"] = -4162
	enum["xlPatternVertical"] = -4166

	return enum
}

func GetEnumPatternNum(enumType string) int32 {
	var result int32
	enum := EnumPattern()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlPatternAutomatic"]
	}
	return result
}

func GetEnumPatternStr(enumType int32) string {
	var result string
	enum := EnumPattern()
	result = "xlPatternAutomatic"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumPattern(enumType int32) int32 {
	var result int32
	enum := EnumPattern()
	result = enum["xlPatternAutomatic"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Shape
func EnumShapeType() map[string]int32 {
	enum := make(map[string]int32)

	enum["mso3DModel"] = 30
	enum["msoAutoShape"] = 1
	enum["msoCallout"] = 2
	enum["msoCanvas"] = 20
	enum["msoChart"] = 3
	enum["msoComment"] = 4
	enum["msoContentApp"] = 27
	enum["msoDiagram"] = 21
	enum["msoEmbeddedOLEObject"] = 7
	enum["msoFormControl"] = 8
	enum["msoFreeform"] = 5
	enum["msoGraphic"] = 28
	enum["msoGroup"] = 6
	enum["msoIgxGraphic"] = 24
	enum["msoInk"] = 22
	enum["msoInkComment"] = 23
	enum["msoLine"] = 9
	enum["msoLinked3DModel"] = 31
	enum["msoLinkedGraphic"] = 29
	enum["msoLinkedOLEObject"] = 10
	enum["msoLinkedPicture"] = 11
	enum["msoMedia"] = 16
	enum["msoOLEControlObject"] = 12
	enum["msoPicture"] = 13
	enum["msoPlaceholder"] = 14
	enum["msoScriptAnchor"] = 18
	enum["msoShapeTypeMixed"] = -2
	enum["msoSlicer"] = 25
	enum["msoTable"] = 19
	enum["msoTextBox"] = 17
	enum["msoTextEffect"] = 15
	enum["msoWebVideo"] = 26

	return enum
}

func GetEnumShapeTypeNum(enumType string) int32 {
	var result int32
	enum := EnumShapeType()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["msoAutoShape"]
	}
	return result
}

func GetEnumShapeTypeStr(enumType int32) string {
	var result string
	enum := EnumShapeType()
	result = "msoAutoShape"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumShapeType(enumType int32) int32 {
	var result int32
	enum := EnumShapeType()
	result = enum["msoAutoShape"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// LineDash
func EnumLineDash() map[string]int32 {
	enum := make(map[string]int32)

	//enum["msoLineDashStyleMixed"] = -2
	enum["msoLineSolid"] = 1
	enum["msoLineSquareDot"] = 2
	enum["msoLineRoundDot"] = 3
	enum["msoLineDash"] = 4
	enum["msoLineDashDot"] = 5
	enum["msoLineDashDotDot"] = 6
	enum["msoLineLongDash"] = 7
	enum["msoLineLongDashDot"] = 8
	enum["msoLineLongDashDotDot"] = 9
	enum["msoLineSysDash"] = 10
	enum["msoLineSysDot"] = 11
	enum["msoLineSysDashDot"] = 12

	return enum
}

func GetEnumLineDashNum(enumType string) int32 {
	var result int32
	enum := EnumLineDash()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["msoLineSolid"]
	}
	return result
}

func GetEnumLineDashStr(enumType int32) string {
	var result string
	enum := EnumLineDash()
	result = "msoLineSolid"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumLineDash(enumType int32) int32 {
	var result int32
	enum := EnumLineDash()
	result = enum["msoLineSolid"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Align
func EnumAlignCmd() map[string]int32 {
	enum := make(map[string]int32)

	enum["msoAlignBottoms"] = 5
	enum["msoAlignCenters"] = 1
	enum["msoAlignLefts"] = 0
	enum["msoAlignMiddles"] = 4
	enum["msoAlignRights"] = 2
	enum["msoAlignTops"] = 3

	return enum
}

func GetEnumAlignCmdNum(enumType string) int32 {
	var result int32
	enum := EnumAlignCmd()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["msoAlignLefts"]
	}
	return result
}

func GetEnumAlignCmdStr(enumType int32) string {
	var result string
	enum := EnumAlignCmd()
	result = "msoAlignLefts"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumAlignCmd(enumType int32) int32 {
	var result int32
	enum := EnumAlignCmd()
	result = enum["msoAlignLefts"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Flip
func EnumFlipCmd() map[string]int32 {
	enum := make(map[string]int32)

	enum["msoFlipHorizontal"] = 0
	enum["msoFlipVertical"] = 1

	return enum
}

func GetEnumFlipCmdNum(enumType string) int32 {
	var result int32
	enum := EnumAlignCmd()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["msoFlipHorizontal"]
	}
	return result
}

func GetEnumFlipCmdStr(enumType int32) string {
	var result string
	enum := EnumAlignCmd()
	result = "msoFlipHorizontal"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumFlipCmd(enumType int32) int32 {
	var result int32
	enum := EnumAlignCmd()
	result = enum["msoFlipHorizontal"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// ZOder
func EnumZOrderCmd() map[string]int32 {
	enum := make(map[string]int32)

	enum["msoBringForward"] = 2
	enum["msoBringInFrontOfText"] = 4
	enum["msoBringToFront"] = 0
	enum["msoSendBackward"] = 3
	enum["msoSendBehindText"] = 5
	enum["msoSendToBack"] = 1

	return enum
}

func GetEnumZOrderCmdNum(enumType string) int32 {
	var result int32
	enum := EnumZOrderCmd()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["msoBringToFront"]
	}
	return result
}

func GetEnumZOrderCmdStr(enumType int32) string {
	var result string
	enum := EnumZOrderCmd()
	result = "msoBringToFront"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumZOrderCmd(enumType int32) int32 {
	var result int32
	enum := EnumZOrderCmd()
	result = enum["msoBringToFront"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// AutoShape
func EnumAutoShape() map[string]int32 {
	enum := make(map[string]int32)

	enum["msoShape10pointStar"] = 149
	enum["msoShape12pointStar"] = 150
	enum["msoShape16pointStar"] = 94
	enum["msoShape24pointStar"] = 95
	enum["msoShape32pointStar"] = 96
	enum["msoShape4pointStar"] = 91
	enum["msoShape5pointStar"] = 92
	enum["msoShape6pointStar"] = 147
	enum["msoShape7pointStar"] = 148
	enum["msoShape8pointStar"] = 93
	enum["msoShapeActionButtonBackorPrevious"] = 129
	enum["msoShapeActionButtonBeginning"] = 131
	enum["msoShapeActionButtonCustom"] = 125
	enum["msoShapeActionButtonDocument"] = 134
	enum["msoShapeActionButtonEnd"] = 132
	enum["msoShapeActionButtonForwardorNext"] = 130
	enum["msoShapeActionButtonHelp"] = 127
	enum["msoShapeActionButtonHome"] = 126
	enum["msoShapeActionButtonInformation"] = 128
	enum["msoShapeActionButtonMovie"] = 136
	enum["msoShapeActionButtonReturn"] = 133
	enum["msoShapeActionButtonSound"] = 135
	enum["msoShapeArc"] = 25
	enum["msoShapeBalloon"] = 137
	enum["msoShapeBentArrow"] = 41
	enum["msoShapeBentUpArrow"] = 44
	enum["msoShapeBevel"] = 15
	enum["msoShapeBlockArc"] = 20
	enum["msoShapeCan"] = 13
	enum["msoShapeChartPlus"] = 182
	enum["msoShapeChartStar"] = 181
	enum["msoShapeChartX"] = 180
	enum["msoShapeChevron"] = 52
	enum["msoShapeChord"] = 161
	enum["msoShapeCircularArrow"] = 60
	enum["msoShapeCloud"] = 179
	enum["msoShapeCloudCallout"] = 108
	enum["msoShapeCorner"] = 162
	enum["msoShapeCornerTabs"] = 169
	enum["msoShapeCross"] = 11
	enum["msoShapeCube"] = 14
	enum["msoShapeCurvedDownArrow"] = 48
	enum["msoShapeCurvedDownRibbon"] = 100
	enum["msoShapeCurvedLeftArrow"] = 46
	enum["msoShapeCurvedRightArrow"] = 45
	enum["msoShapeCurvedUpArrow"] = 47
	enum["msoShapeCurvedUpRibbon"] = 99
	enum["msoShapeDecagon"] = 144
	enum["msoShapeDiagonalStripe"] = 141
	enum["msoShapeDiamond"] = 4
	enum["msoShapeDodecagon"] = 146
	enum["msoShapeDonut"] = 18
	enum["msoShapeDoubleBrace"] = 27
	enum["msoShapeDoubleBracket"] = 26
	enum["msoShapeDoubleWave"] = 104
	enum["msoShapeDownArrow"] = 36
	enum["msoShapeDownArrowCallout"] = 56
	enum["msoShapeDownRibbon"] = 98
	enum["msoShapeExplosion1"] = 89
	enum["msoShapeExplosion2"] = 90
	enum["msoShapeFlowchartAlternateProcess"] = 62
	enum["msoShapeFlowchartCard"] = 75
	enum["msoShapeFlowchartCollate"] = 79
	enum["msoShapeFlowchartConnector"] = 73
	enum["msoShapeFlowchartData"] = 64
	enum["msoShapeFlowchartDecision"] = 63
	enum["msoShapeFlowchartDelay"] = 84
	enum["msoShapeFlowchartDirectAccessStorage"] = 87
	enum["msoShapeFlowchartDisplay"] = 88
	enum["msoShapeFlowchartDocument"] = 67
	enum["msoShapeFlowchartExtract"] = 81
	enum["msoShapeFlowchartInternalStorage"] = 66
	enum["msoShapeFlowchartMagneticDisk"] = 86
	enum["msoShapeFlowchartManualInput"] = 71
	enum["msoShapeFlowchartManualOperation"] = 72
	enum["msoShapeFlowchartMerge"] = 82
	enum["msoShapeFlowchartMultidocument"] = 68
	enum["msoShapeFlowchartOfflineStorage"] = 139
	enum["msoShapeFlowchartOffpageConnector"] = 74
	enum["msoShapeFlowchartOr"] = 78
	enum["msoShapeFlowchartPredefinedProcess"] = 65
	enum["msoShapeFlowchartPreparation"] = 70
	enum["msoShapeFlowchartProcess"] = 61
	enum["msoShapeFlowchartPunchedTape"] = 76
	enum["msoShapeFlowchartSequentialAccessStorage"] = 85
	enum["msoShapeFlowchartSort"] = 80
	enum["msoShapeFlowchartStoredData"] = 83
	enum["msoShapeFlowchartSummingJunction"] = 77
	enum["msoShapeFlowchartTerminator"] = 69
	enum["msoShapeFoldedCorner"] = 16
	enum["msoShapeFrame"] = 158
	enum["msoShapeFunnel"] = 174
	enum["msoShapeGear6"] = 172
	enum["msoShapeGear9"] = 173
	enum["msoShapeHalfFrame"] = 159
	enum["msoShapeHeart"] = 21
	enum["msoShapeHeptagon"] = 145
	enum["msoShapeHexagon"] = 10
	enum["msoShapeHorizontalScroll"] = 102
	enum["msoShapeIsoscelesTriangle"] = 7
	enum["msoShapeLeftArrow"] = 34
	enum["msoShapeLeftArrowCallout"] = 54
	enum["msoShapeLeftBrace"] = 31
	enum["msoShapeLeftBracket"] = 29
	enum["msoShapeLeftCircularArrow"] = 176
	enum["msoShapeLeftRightArrow"] = 37
	enum["msoShapeLeftRightArrowCallout"] = 57
	enum["msoShapeLeftRightCircularArrow"] = 177
	enum["msoShapeLeftRightRibbon"] = 140
	enum["msoShapeLeftRightUpArrow"] = 40
	enum["msoShapeLeftUpArrow"] = 43
	enum["msoShapeLightningBolt"] = 22
	enum["msoShapeLineCallout1"] = 109
	enum["msoShapeLineCallout1AccentBar"] = 113
	enum["msoShapeLineCallout1BorderandAccentBar"] = 121
	enum["msoShapeLineCallout1NoBorder"] = 117
	enum["msoShapeLineCallout2"] = 110
	enum["msoShapeLineCallout2AccentBar"] = 114
	enum["msoShapeLineCallout2BorderandAccentBar"] = 122
	enum["msoShapeLineCallout2NoBorder"] = 118
	enum["msoShapeLineCallout3"] = 111
	enum["msoShapeLineCallout3AccentBar"] = 115
	enum["msoShapeLineCallout3BorderandAccentBar"] = 123
	enum["msoShapeLineCallout3NoBorder"] = 119
	enum["msoShapeLineCallout4"] = 112
	enum["msoShapeLineCallout4AccentBar"] = 116
	enum["msoShapeLineCallout4BorderandAccentBar"] = 124
	enum["msoShapeLineCallout4NoBorder"] = 120
	enum["msoShapeLineInverse"] = 183
	enum["msoShapeMathDivide"] = 166
	enum["msoShapeMathEqual"] = 167
	enum["msoShapeMathMinus"] = 164
	enum["msoShapeMathMultiply"] = 165
	enum["msoShapeMathNotEqual"] = 168
	enum["msoShapeMathPlus"] = 163
	//enum["msoShapeMixed"] = -2
	enum["msoShapeMoon"] = 24
	enum["msoShapeNonIsoscelesTrapezoid"] = 143
	enum["msoShapeNoSymbol"] = 19
	enum["msoShapeNotchedRightArrow"] = 50
	//enum["msoShapeNotPrimitive"] = 138
	enum["msoShapeOctagon"] = 6
	enum["msoShapeOval"] = 9
	enum["msoShapeOvalCallout"] = 107
	enum["msoShapeParallelogram"] = 2
	enum["msoShapePentagon"] = 51
	enum["msoShapePie"] = 142
	enum["msoShapePieWedge"] = 175
	enum["msoShapePlaque"] = 28
	enum["msoShapePlaqueTabs"] = 171
	enum["msoShapeQuadArrow"] = 39
	enum["msoShapeQuadArrowCallout"] = 59
	enum["msoShapeRectangle"] = 1
	enum["msoShapeRectangularCallout"] = 105
	enum["msoShapeRegularPentagon"] = 12
	enum["msoShapeRightArrow"] = 33
	enum["msoShapeRightArrowCallout"] = 53
	enum["msoShapeRightBrace"] = 32
	enum["msoShapeRightBracket"] = 30
	enum["msoShapeRightTriangle"] = 8
	enum["msoShapeRound1Rectangle"] = 151
	enum["msoShapeRound2DiagRectangle"] = 157
	enum["msoShapeRound2SameRectangle"] = 152
	enum["msoShapeRoundedRectangle"] = 5
	enum["msoShapeRoundedRectangularCallout"] = 106
	enum["msoShapeSmileyFace"] = 17
	enum["msoShapeSnip1Rectangle"] = 155
	enum["msoShapeSnip2DiagRectangle"] = 157
	enum["msoShapeSnip2SameRectangle"] = 156
	enum["msoShapeSnipRoundRectangle"] = 154
	enum["msoShapeSquareTabs"] = 170
	enum["msoShapeStripedRightArrow"] = 49
	enum["msoShapeSun"] = 23
	enum["msoShapeSwooshArrow"] = 178
	enum["msoShapeTear"] = 160
	enum["msoShapeTrapezoid"] = 3
	enum["msoShapeUpArrow"] = 35
	enum["msoShapeUpArrowCallout"] = 55
	enum["msoShapeUpDownArrow"] = 38
	enum["msoShapeUpDownArrowCallout"] = 58
	enum["msoShapeUpRibbon"] = 97
	enum["msoShapeUTurnArrow"] = 42
	enum["msoShapeVerticalScroll"] = 101
	enum["msoShapeWave"] = 103

	return enum
}

func GetEnumAutoShapeNum(enumType string) int32 {
	var result int32
	enum := EnumAutoShape()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["msoShapeRectangle"]
	}
	return result
}

func GetEnumAutoShapeStr(enumType int32) string {
	var result string
	enum := EnumAutoShape()
	result = "msoShapeRectangle"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumAutoShape(enumType int32) int32 {
	var result int32
	enum := EnumAutoShape()
	result = enum["msoShapeRectangle"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// UpdateLinks
func EnumUpdateLinks() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlUpdateLinks"] = 0 //original
	enum["xlUpdateLinksUserSetting"] = 1
	enum["xlUpdateLinksNever"] = 2
	enum["xlUpdateLinksAlways"] = 3

	return enum
}

func GetEnumUpdateLinksNum(enumType string) int32 {
	var result int32
	enum := EnumAutoShape()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlUpdateLinks"]
	}
	return result
}

func GetEnumUpdateLinksStr(enumType int32) string {
	var result string
	enum := EnumAutoShape()
	result = "xlUpdateLinks"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumUpdateLinks(enumType int32) int32 {
	var result int32
	enum := EnumAutoShape()
	result = enum["xlUpdateLinks"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Alignment XlHAlign
func EnumHAlign() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlHAlignCenter"] = -4108 //Default
	enum["xlHAlignCenterAcrossSelection"] = 7
	enum["xlHAlignDistributed"] = -4117
	enum["xlHAlignFill"] = 5
	enum["xlHAlignGeneral"] = 1
	enum["xlHAlignJustify"] = -4130
	enum["xlHAlignLeft"] = -4131
	enum["xlHAlignRight"] = -4152

	return enum
}

func GetEnumHAlignNum(enumType string) int32 {
	var result int32
	enum := EnumHAlign()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlHAlignCenter"]
	}
	return result
}

func GetEnumHAlignStr(enumType int32) string {
	var result string
	enum := EnumHAlign()
	result = "xlHAlignCenter"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumHAlign(enumType int32) int32 {
	var result int32
	enum := EnumHAlign()
	result = enum["xlHAlignCenter"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Alignment XlVAlign
func EnumVAlign() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlVAlignBottom"] = -4107 //Default
	enum["xlVAlignCenter"] = -4108
	enum["xlVAlignDistributed"] = -4117
	enum["xlVAlignJustify"] = -4130
	enum["xlVAlignTop"] = -4160

	return enum
}

func GetEnumVAlignNum(enumType string) int32 {
	var result int32
	enum := EnumVAlign()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlVAlignBottom"]
	}
	return result
}

func GetEnumVAlignStr(enumType int32) string {
	var result string
	enum := EnumVAlign()
	result = "xlVAlignBottom"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumVAlign(enumType int32) int32 {
	var result int32
	enum := EnumVAlign()
	result = enum["xlVAlignBottom"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Alignment XlOrientation
func EnumOrientation() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlDownward"] = -4170
	enum["xlHorizontal"] = -4128 //Default
	enum["xlUpward"] = -4171
	enum["xlVertical"] = -4166

	return enum
}

func GetEnumOrientationNum(enumType string) int32 {
	var result int32
	enum := EnumOrientation()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlHorizontal"]
	}
	return result
}

func GetEnumOrientationStr(enumType int32) string {
	var result string
	enum := EnumOrientation()
	result = "xlHorizontal"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumOrientation(enumType int32) int32 {
	var result int32
	enum := EnumOrientation()
	result = enum["xlHorizontal"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Alignment XlReadingOrder
func EnumReadingOrder() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlContext"] = -5002 //Default
	enum["xlLTR"] = -5003
	enum["xlRTL"] = -5004

	return enum
}

func GetEnumReadingOrderNum(enumType string) int32 {
	var result int32
	enum := EnumReadingOrder()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlContext"]
	}
	return result
}

func GetEnumReadingOrderStr(enumType int32) string {
	var result string
	enum := EnumReadingOrder()
	result = "xlContext"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumReadingOrder(enumType int32) int32 {
	var result int32
	enum := EnumReadingOrder()
	result = enum["xlContext"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Alignment XlUnderlineStyle
func EnumUnderlineStyle() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlUnderlineStyleDouble"] = -4119
	enum["xlUnderlineStyleDoubleAccounting"] = 5
	enum["xlUnderlineStyleNone"] = -4142 //Default
	enum["xlUnderlineStyleSingle"] = 2
	enum["xlUnderlineStyleSingleAccounting"] = 4

	return enum
}

func GetEnumUnderlineStyleNum(enumType string) int32 {
	var result int32
	enum := EnumUnderlineStyle()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlUnderlineStyleNone"]
	}
	return result
}

func GetEnumUnderlineStyleStr(enumType int32) string {
	var result string
	enum := EnumUnderlineStyle()
	result = "xlUnderlineStyleNone"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumUnderlineStyle(enumType int32) int32 {
	var result int32
	enum := EnumUnderlineStyle()
	result = enum["xlUnderlineStyleNone"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// Color XlRgbColor
func EnumRgbColor() map[string]float64 {
	enum := make(map[string]float64)

	enum["rgbAliceBlue"] = 16775408
	enum["rgbAntiqueWhite"] = 14150650
	enum["rgbAqua"] = 16776960
	enum["rgbAquamarine"] = 13959039
	enum["rgbAzure"] = 16777200
	enum["rgbBeige"] = 14480885
	enum["rgbBisque"] = 12903679
	enum["rgbBlack"] = 0 //Default
	enum["rgbBlanchedAlmond"] = 13495295
	enum["rgbBlue"] = 16711680
	enum["rgbBlueViolet"] = 14822282
	enum["rgbBrown"] = 2763429
	enum["rgbBurlyWood"] = 8894686
	enum["rgbCadetBlue"] = 10526303
	enum["rgbChartreuse"] = 65407
	enum["rgbCoral"] = 5275647
	enum["rgbCornflowerBlue"] = 15570276
	enum["rgbCornsilk"] = 14481663
	enum["rgbCrimson"] = 3937500
	enum["rgbDarkBlue"] = 9109504
	enum["rgbDarkCyan"] = 9145088
	enum["rgbDarkGoldenrod"] = 755384
	enum["rgbDarkGray"] = 11119017
	enum["rgbDarkGreen"] = 25600
	enum["rgbDarkGrey"] = 11119017
	enum["rgbDarkKhaki"] = 7059389
	enum["rgbDarkMagenta"] = 9109643
	enum["rgbDarkOliveGreen"] = 3107669
	enum["rgbDarkOrange"] = 36095
	enum["rgbDarkOrchid"] = 13382297
	enum["rgbDarkRed"] = 139
	enum["rgbDarkSalmon"] = 8034025
	enum["rgbDarkSeaGreen"] = 9419919
	enum["rgbDarkSlateBlue"] = 9125192
	enum["rgbDarkSlateGray"] = 5197615
	enum["rgbDarkSlateGrey"] = 5197615
	enum["rgbDarkTurquoise"] = 13749760
	enum["rgbDarkViolet"] = 13828244
	enum["rgbDeepPink"] = 9639167
	enum["rgbDeepSkyBlue"] = 16760576
	enum["rgbDimGray"] = 6908265
	enum["rgbDimGrey"] = 6908265
	enum["rgbDodgerBlue"] = 16748574
	enum["rgbFireBrick"] = 2237106
	enum["rgbFloralWhite"] = 15792895
	enum["rgbForestGreen"] = 2263842
	enum["rgbFuchsia"] = 16711935
	enum["rgbGainsboro"] = 14474460
	enum["rgbGhostWhite"] = 16775416
	enum["rgbGold"] = 55295
	enum["rgbGoldenrod"] = 2139610
	enum["rgbGray"] = 8421504
	enum["rgbGreen"] = 32768
	enum["rgbGreenYellow"] = 3145645
	enum["rgbGrey"] = 8421504
	enum["rgbHoneydew"] = 15794160
	enum["rgbHotPink"] = 11823615
	enum["rgbIndianRed"] = 6053069
	enum["rgbIndigo"] = 8519755
	enum["rgbIvory"] = 15794175
	enum["rgbKhaki"] = 9234160
	enum["rgbLavender"] = 16443110
	enum["rgbLavenderBlush"] = 16118015
	enum["rgbLawnGreen"] = 64636
	enum["rgbLemonChiffon"] = 13499135
	enum["rgbLightBlue"] = 15128749
	enum["rgbLightCoral"] = 8421616
	enum["rgbLightCyan"] = 9145088
	enum["rgbLightGoldenrodYellow"] = 13826810
	enum["rgbLightGray"] = 13882323
	enum["rgbLightGreen"] = 9498256
	enum["rgbLightGrey"] = 13882323
	enum["rgbLightPink"] = 12695295
	enum["rgbLightSalmon"] = 8036607
	enum["rgbLightSeaGreen"] = 11186720
	enum["rgbLightSkyBlue"] = 16436871
	enum["rgbLightSlateGray"] = 10061943
	enum["rgbLightSteelBlue"] = 14599344
	enum["rgbLightYellow"] = 14745599
	enum["rgbLime"] = 65280
	enum["rgbLimeGreen"] = 3329330
	enum["rgbLinen"] = 15134970
	enum["rgbMaroon"] = 128
	enum["rgbMediumAquamarine"] = 11206502
	enum["rgbMediumBlue"] = 13434880
	enum["rgbMediumOrchid"] = 13850042
	enum["rgbMediumPurple"] = 14381203
	enum["rgbMediumSeaGreen"] = 7451452
	enum["rgbMediumSlateBlue"] = 15624315
	enum["rgbMediumSpringGreen"] = 10156544
	enum["rgbMediumTurquoise"] = 13422920
	enum["rgbMediumVioletRed"] = 8721863
	enum["rgbMidnightBlue"] = 7346457
	enum["rgbMintCream"] = 16449525
	enum["rgbMistyRose"] = 14804223
	enum["rgbMoccasin"] = 11920639
	enum["rgbNavajoWhite"] = 11394815
	enum["rgbNavy"] = 8388608
	enum["rgbNavyBlue"] = 8388608
	enum["rgbOldLace"] = 15136253
	enum["rgbOlive"] = 32896
	enum["rgbOliveDrab"] = 2330219
	enum["rgbOrange"] = 42495
	enum["rgbOrangeRed"] = 17919
	enum["rgbOrchid"] = 14053594
	enum["rgbPaleGoldenrod"] = 7071982
	enum["rgbPaleGreen"] = 10025880
	enum["rgbPaleTurquoise"] = 15658671
	enum["rgbPaleVioletRed"] = 9662683
	enum["rgbPapayaWhip"] = 14020607
	enum["rgbPeachPuff"] = 12180223
	enum["rgbPeru"] = 4163021
	enum["rgbPink"] = 13353215
	enum["rgbPlum"] = 14524637
	enum["rgbPowderBlue"] = 15130800
	enum["rgbPurple"] = 8388736
	enum["rgbRed"] = 255
	enum["rgbRosyBrown"] = 9408444
	enum["rgbRoyalBlue"] = 14772545
	enum["rgbSalmon"] = 7504122
	enum["rgbSandyBrown"] = 6333684
	enum["rgbSeaGreen"] = 5737262
	enum["rgbSeashell"] = 15660543
	enum["rgbSienna"] = 2970272
	enum["rgbSilver"] = 12632256
	enum["rgbSkyBlue"] = 15453831
	enum["rgbSlateBlue"] = 13458026
	enum["rgbSlateGray"] = 9470064
	enum["rgbSnow"] = 16448255
	enum["rgbSpringGreen"] = 8388352
	enum["rgbSteelBlue"] = 11829830
	enum["rgbTan"] = 9221330
	enum["rgbTeal"] = 8421376
	enum["rgbThistle"] = 14204888
	enum["rgbTomato"] = 4678655
	enum["rgbTurquoise"] = 13688896
	enum["rgbViolet"] = 15631086
	enum["rgbWheat"] = 11788021
	enum["rgbWhite"] = 16777215
	enum["rgbWhiteSmoke"] = 16119285
	enum["rgbYellow"] = 65535
	enum["rgbYellowGreen"] = 3329434

	return enum
}

func GetEnumRgbColorNum(enumType string) float64 {
	var result float64
	enum := EnumRgbColor()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["rgbBlack"]
	}
	return result
}

// Color XlThemeColor
func EnumThemeColor() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlThemeColorAccent1"] = 5
	enum["xlThemeColorAccent2"] = 6
	enum["xlThemeColorAccent3"] = 7
	enum["xlThemeColorAccent4"] = 8
	enum["xlThemeColorAccent5"] = 9
	enum["xlThemeColorAccent6"] = 10
	enum["xlThemeColorDark1"] = 1
	enum["xlThemeColorDark2"] = 3
	enum["xlThemeColorFollowedHyperlink"] = 12
	enum["xlThemeColorHyperlink"] = 11
	enum["xlThemeColorLight1"] = 2 //Def
	enum["xlThemeColorLight2"] = 4

	return enum
}

func GetEnumUThemeColorNum(enumType string) int32 {
	var result int32
	enum := EnumThemeColor()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlThemeColorLight1"]
	}
	return result
}

func GetEnumThemeColorStr(enumType int32) string {
	var result string
	enum := EnumThemeColor()
	result = "xlThemeColorLight1"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumThemeColor(enumType int32) int32 {
	var result int32
	enum := EnumThemeColor()
	result = enum["xlThemeColorLight1"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// XlChartType
func EnumChartType() map[string]int32 {
	enum := make(map[string]int32)

	enum["xl3DArea"] = -4098
	enum["xl3DAreaStacked"] = 78
	enum["xl3DAreaStacked100"] = 79
	enum["xl3DBarClustered"] = 60
	enum["xl3DBarStacked"] = 61
	enum["xl3DBarStacked100"] = 62
	enum["xl3DColumn"] = -4100
	enum["xl3DColumnClustered"] = 54
	enum["xl3DColumnStacked"] = 55
	enum["xl3DColumnStacked100"] = 56
	enum["xl3DLine"] = -4101
	enum["xl3DPie"] = -4102
	enum["xl3DPieExploded"] = 70
	enum["xlArea"] = 1
	enum["xlAreaEx"] = 135
	enum["xlAreaStacked"] = 76
	enum["xlAreaStacked100"] = 77
	enum["xlAreaStacked100Ex"] = 137
	enum["xlAreaStackedEx"] = 136
	enum["xlBarClustered"] = 57
	enum["xlBarClusteredEx"] = 132
	enum["xlBarOfPie"] = 71
	enum["xlBarStacked"] = 58
	enum["xlBarStacked100"] = 59
	enum["xlBarStacked100Ex"] = 134
	enum["xlBarStackedEx"] = 133
	enum["xlBoxwhisker"] = 121
	enum["xlBubble"] = 15
	enum["xlBubble3DEffect"] = 87
	enum["xlBubbleEx"] = 139
	enum["xlColumnClustered"] = 51
	enum["xlColumnClusteredEx"] = 124
	enum["xlColumnStacked"] = 52
	enum["xlColumnStacked100"] = 53
	enum["xlColumnStacked100Ex"] = 126
	enum["xlColumnStackedEx"] = 125
	enum["xlCombo"] = -4152
	enum["xlComboAreaStackedColumnClustered"] = 115
	enum["xlComboColumnClusteredLine"] = 113
	enum["xlComboColumnClusteredLineSecondaryAxis"] = 114
	enum["xlConeBarClustered"] = 102
	enum["xlConeBarStacked"] = 103
	enum["xlConeBarStacked100"] = 104
	enum["xlConeCol"] = 105
	enum["xlConeColClustered"] = 99
	enum["xlConeColStacked"] = 100
	enum["xlConeColStacked100"] = 101
	enum["xlCylinderBarClustered"] = 95
	enum["xlCylinderBarStacked"] = 96
	enum["xlCylinderBarStacked100"] = 97
	enum["xlCylinderCol"] = 98
	enum["xlCylinderColClustered"] = 92
	enum["xlCylinderColStacked"] = 93
	enum["xlCylinderColStacked100"] = 94
	enum["xlDoughnut"] = -4120
	enum["xlDoughnutEx"] = 131
	enum["xlDoughnutExploded"] = 80
	enum["xlFunnel"] = 123
	enum["xlHistogram"] = 118
	enum["xlLine"] = 4 //Def
	enum["xlLineEx"] = 127
	enum["xlLineMarkers"] = 65
	enum["xlLineMarkersStacked"] = 66
	enum["xlLineMarkersStacked100"] = 67
	enum["xlLineStacked"] = 63
	enum["xlLineStacked100"] = 64
	enum["xlLineStacked100Ex"] = 129
	enum["xlLineStackedEx"] = 128
	enum["xlOtherCombinations"] = 116
	enum["xlPareto"] = 122
	enum["xlPie"] = 5
	enum["xlPieEx"] = 130
	enum["xlPieExploded"] = 69
	enum["xlPieOfPie"] = 68
	enum["xlPyramidBarClustered"] = 109
	enum["xlPyramidBarStacked"] = 110
	enum["xlPyramidBarStacked100"] = 111
	enum["xlPyramidCol"] = 112
	enum["xlPyramidColClustered"] = 106
	enum["xlPyramidColStacked"] = 107
	enum["xlPyramidColStacked100"] = 108
	enum["xlRadar"] = -4151
	enum["xlRadarFilled"] = 82
	enum["xlRadarMarkers"] = 81
	enum["xlRegionMap"] = 140
	enum["xlStockHLC"] = 88
	enum["xlStockOHLC"] = 89
	enum["xlStockVHLC"] = 90
	enum["xlStockVOHLC"] = 91
	enum["xlSuggestedChart"] = -2
	enum["xlSunburst"] = 120
	enum["xlSurface"] = 83
	enum["xlSurfaceTopView"] = 85
	enum["xlSurfaceTopViewWireframe"] = 86
	enum["xlSurfaceWireframe"] = 84
	enum["xlTreemap"] = 117
	enum["xlWaterfall"] = 119
	enum["xlXYScatter"] = -4169
	enum["xlXYScatterEx"] = 138
	enum["xlXYScatterLines"] = 74
	enum["xlXYScatterLinesNoMarkers"] = 75
	enum["xlXYScatterSmooth"] = 72
	enum["xlXYScatterSmoothNoMarkers"] = 73

	return enum
}

func GetEnumChartTypeNum(enumType string) int32 {
	var result int32
	enum := EnumChartType()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlLine"]
	}
	return result
}

func GetEnumChartTypeStr(enumType int32) string {
	var result string
	enum := EnumChartType()
	result = "xlLine"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumChartType(enumType int32) int32 {
	var result int32
	enum := EnumChartType()
	result = enum["xlLine"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// XlRowCol
func EnumRowCol() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlColumns"] = 2
	enum["xlRows"] = 1 //Def

	return enum
}

func GetEnumRowColNum(enumType string) int32 {
	var result int32
	enum := EnumRowCol()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlRows"]
	}
	return result
}

func GetEnumRowColStr(enumType int32) string {
	var result string
	enum := EnumRowCol()
	result = "xlRows"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumRowCol(enumType int32) int32 {
	var result int32
	enum := EnumRowCol()
	result = enum["xlRows"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// XlReferenceStyle
func EnumReferenceStyle() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlA1"] = 1 //Default
	enum["xlR1C1"] = -4150

	return enum
}

func GetEnumReferenceStyleNum(enumType string) int32 {
	var result int32
	enum := EnumReferenceStyle()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlA1"]
	}
	return result
}

func GetEnumReferenceStyleStr(enumType int32) string {
	var result string
	enum := EnumReferenceStyle()
	result = "xlA1"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumReferenceStyle(enumType int32) int32 {
	var result int32
	enum := EnumReferenceStyle()
	result = enum["xlA1"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// XlLegendPosition
func EnumLegendPosition() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlLegendPositionBottom"] = -4107
	enum["xlLegendPositionCorner"] = 2
	enum["xlLegendPositionCustom"] = -4161
	enum["xlLegendPositionLeft"] = -4131
	enum["xlLegendPositionRight"] = -4152
	enum["xlLegendPositionTop"] = -4160 //Default

	return enum
}

func GetEnumLegendPositionNum(enumType string) int32 {
	var result int32
	enum := EnumLegendPosition()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlLegendPositionTop"]
	}
	return result
}

func GetEnumLegendPositionStr(enumType int32) string {
	var result string
	enum := EnumLegendPosition()
	result = "xlLegendPositionTop"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumLegendPosition(enumType int32) int32 {
	var result int32
	enum := EnumLegendPosition()
	result = enum["xlLegendPositionTop"]

	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// XlAxisGroup
func EnumAxisGroup() map[string]int32 {
	enum := make(map[string]int32)
	enum["xlPrimary"] = 1 //Default
	enum["xlSecondary"] = 2
	return enum
}

func GetEnumAxisGroupNum(enumType string) int32 {
	var result int32
	enum := EnumAxisGroup()
	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlPrimary"]
	}
	return result
}

func GetEnumAxisGroupStr(enumType int32) string {
	var result string
	enum := EnumAxisGroup()
	result = "xlPrimary"
	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumAxisGroup(enumType int32) int32 {
	var result int32
	enum := EnumAxisGroup()
	result = enum["xlPrimary"]
	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// XlAxisType
func EnumAxisType() map[string]int32 {
	enum := make(map[string]int32)
	enum["xlCategory"] = 1 //Default
	enum["xlSeriesAxis"] = 3
	enum["xlValue"] = 2
	return enum
}

func GetEnumAxisTypeNum(enumType string) int32 {
	var result int32
	enum := EnumAxisType()
	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlCategory"]
	}
	return result
}

func GetEnumAxisTypeStr(enumType int32) string {
	var result string
	enum := EnumAxisType()
	result = "xlCategory"
	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumAxisType(enumType int32) int32 {
	var result int32
	enum := EnumAxisType()
	result = enum["xlCategory"]
	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// XlTickLabelPosition
func EnumTickLabelPosition() map[string]int32 {
	enum := make(map[string]int32)
	enum["xlTickLabelPositionHigh"] = -4127
	enum["xlTickLabelPositionLow"] = -4134
	enum["xlTickLabelPositionNextToAxis"] = 4 //Default
	enum["xlTickLabelPositionNone"] = -4142
	return enum
}

func GetEnumTickLabelPositionNum(enumType string) int32 {
	var result int32
	enum := EnumTickLabelPosition()
	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlTickLabelPositionNextToAxis"]
	}
	return result
}

func GetEnumTickLabelPositionStr(enumType int32) string {
	var result string
	enum := EnumTickLabelPosition()
	result = "xlTickLabelPositionNextToAxis"
	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumTickLabelPosition(enumType int32) int32 {
	var result int32
	enum := EnumTickLabelPosition()
	result = enum["xlTickLabelPositionNextToAxis"]
	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// msoChartElementType
func EnumChartElementType() map[string]int32 {
	enum := make(map[string]int32)
	enum["msoElementChartFloorNone"] = 1200
	enum["msoElementChartFloorShow"] = 1201
	enum["msoElementChartTitleAboveChart"] = 2
	enum["msoElementChartTitleCenteredOverlay"] = 1
	enum["msoElementChartTitleNone"] = 0 //Def
	enum["msoElementChartWallNone"] = 1100
	enum["msoElementChartWallShow"] = 1101
	enum["msoElementDataLabelBestFit"] = 210
	enum["msoElementDataLabelBottom"] = 209
	enum["msoElementDataLabelCallout"] = 211
	enum["msoElementDataLabelCenter"] = 202
	enum["msoElementDataLabelInsideBase"] = 204
	enum["msoElementDataLabelInsideEnd"] = 203
	enum["msoElementDataLabelLeft"] = 206
	enum["msoElementDataLabelNone"] = 200
	enum["msoElementDataLabelOutSideEnd"] = 205
	enum["msoElementDataLabelRight"] = 207
	enum["msoElementDataLabelShow"] = 201
	enum["msoElementDataLabelTop"] = 208
	enum["msoElementDataTableNone"] = 500
	enum["msoElementDataTableShow"] = 501
	enum["msoElementDataTableWithLegendKeys"] = 502
	enum["msoElementErrorBarNone"] = 700
	enum["msoElementErrorBarPercentage"] = 702
	enum["msoElementErrorBarStandardDeviation"] = 703
	enum["msoElementErrorBarStandardError"] = 701
	enum["msoElementLegendBottom"] = 104
	enum["msoElementLegendLeft"] = 103
	enum["msoElementLegendLeftOverlay"] = 106
	enum["msoElementLegendNone"] = 100
	enum["msoElementLegendRight"] = 101
	enum["msoElementLegendRightOverlay"] = 105
	enum["msoElementLegendTop"] = 102
	enum["msoElementLineDropHiLoLine"] = 804
	enum["msoElementLineDropLine"] = 801
	enum["msoElementLineHiLoLine"] = 802
	enum["msoElementLineNone"] = 800
	enum["msoElementLineSeriesLine"] = 803
	enum["msoElementPlotAreaNone"] = 1000
	enum["msoElementPlotAreaShow"] = 1001
	enum["msoElementPrimaryCategoryAxisBillions"] = 374
	enum["msoElementPrimaryCategoryAxisLogScale"] = 375
	enum["msoElementPrimaryCategoryAxisMillions"] = 373
	enum["msoElementPrimaryCategoryAxisNone"] = 348
	enum["msoElementPrimaryCategoryAxisReverse"] = 351
	enum["msoElementPrimaryCategoryAxisShow"] = 349
	enum["msoElementPrimaryCategoryAxisThousands"] = 372
	enum["msoElementPrimaryCategoryAxisTitleAdjacentToAxis"] = 301
	enum["msoElementPrimaryCategoryAxisTitleBelowAxis"] = 302
	enum["msoElementPrimaryCategoryAxisTitleHorizontal"] = 305
	enum["msoElementPrimaryCategoryAxisTitleNone"] = 300
	enum["msoElementPrimaryCategoryAxisTitleRotated"] = 303
	enum["msoElementPrimaryCategoryAxisTitleVertical"] = 304
	enum["msoElementPrimaryCategoryAxisWithoutLabels"] = 350
	enum["msoElementPrimaryCategoryGridLinesMajor"] = 334
	enum["msoElementPrimaryCategoryGridLinesMinor"] = 333
	enum["msoElementPrimaryCategoryGridLinesMinorMajor"] = 335
	enum["msoElementPrimaryCategoryGridLinesNone"] = 332
	enum["msoElementPrimaryValueAxisBillions"] = 356
	enum["msoElementPrimaryValueAxisLogScale"] = 357
	enum["msoElementPrimaryValueAxisMillions"] = 355
	enum["msoElementPrimaryValueAxisNone"] = 352
	enum["msoElementPrimaryValueAxisShow"] = 353
	enum["msoElementPrimaryValueAxisThousands"] = 354
	enum["msoElementPrimaryValueAxisTitleAdjacentToAxis"] = 307
	enum["msoElementPrimaryValueAxisTitleBelowAxis"] = 308
	enum["msoElementPrimaryValueAxisTitleHorizontal"] = 311
	enum["msoElementPrimaryValueAxisTitleNone"] = 306
	enum["msoElementPrimaryValueAxisTitleRotated"] = 309
	enum["msoElementPrimaryValueAxisTitleVertical"] = 310
	enum["msoElementPrimaryValueGridLinesMajor"] = 330
	enum["msoElementPrimaryValueGridLinesMinor"] = 329
	enum["msoElementPrimaryValueGridLinesMinorMajor"] = 331
	enum["msoElementPrimaryValueGridLinesNone"] = 328
	enum["msoElementSecondaryCategoryAxisBillions"] = 378
	enum["msoElementSecondaryCategoryAxisLogScale"] = 379
	enum["msoElementSecondaryCategoryAxisMillions"] = 377
	enum["msoElementSecondaryCategoryAxisNone"] = 358
	enum["msoElementSecondaryCategoryAxisReverse"] = 361
	enum["msoElementSecondaryCategoryAxisShow"] = 359
	enum["msoElementSecondaryCategoryAxisThousands"] = 376
	enum["msoElementSecondaryCategoryAxisTitleAdjacentToAxis"] = 313
	enum["msoElementSecondaryCategoryAxisTitleBelowAxis"] = 314
	enum["msoElementSecondaryCategoryAxisTitleHorizontal"] = 317
	enum["msoElementSecondaryCategoryAxisTitleNone"] = 312
	enum["msoElementSecondaryCategoryAxisTitleRotated"] = 315
	enum["msoElementSecondaryCategoryAxisTitleVertical"] = 316
	enum["msoElementSecondaryCategoryAxisWithoutLabels"] = 360
	enum["msoElementSecondaryCategoryGridLinesMajor"] = 342
	enum["msoElementSecondaryCategoryGridLinesMinor"] = 341
	enum["msoElementSecondaryCategoryGridLinesMinorMajor"] = 343
	enum["msoElementSecondaryCategoryGridLinesNone"] = 340
	enum["msoElementSecondaryValueAxisBillions"] = 366
	enum["msoElementSecondaryValueAxisLogScale"] = 367
	enum["msoElementSecondaryValueAxisMillions"] = 365
	enum["msoElementSecondaryValueAxisNone"] = 362
	enum["msoElementSecondaryValueAxisShow"] = 363
	enum["msoElementSecondaryValueAxisThousands"] = 364
	enum["msoElementSecondaryValueAxisTitleAdjacentToAxis"] = 319
	enum["msoElementSecondaryValueAxisTitleBelowAxis"] = 320
	enum["msoElementSecondaryValueAxisTitleHorizontal"] = 323
	enum["msoElementSecondaryValueAxisTitleNone"] = 318
	enum["msoElementSecondaryValueAxisTitleRotated"] = 321
	enum["msoElementSecondaryValueAxisTitleVertical"] = 322
	enum["msoElementSecondaryValueGridLinesMajor"] = 338
	enum["msoElementSecondaryValueGridLinesMinor"] = 337
	enum["msoElementSecondaryValueGridLinesMinorMajor"] = 339
	enum["msoElementSecondaryValueGridLinesNone"] = 336
	enum["msoElementSeriesAxisGridLinesMajor"] = 346
	enum["msoElementSeriesAxisGridLinesMinor"] = 345
	enum["msoElementSeriesAxisGridLinesMinorMajor"] = 347
	enum["msoElementSeriesAxisGridLinesNone"] = 344
	enum["msoElementSeriesAxisNone"] = 368
	enum["msoElementSeriesAxisReverse"] = 371
	enum["msoElementSeriesAxisShow"] = 369
	enum["msoElementSeriesAxisTitleHorizontal"] = 327
	enum["msoElementSeriesAxisTitleNone"] = 324
	enum["msoElementSeriesAxisTitleRotated"] = 325
	enum["msoElementSeriesAxisTitleVertical"] = 326
	enum["msoElementSeriesAxisWithoutLabeling"] = 370
	enum["msoElementTrendlineAddExponential"] = 602
	enum["msoElementTrendlineAddLinear"] = 601
	enum["msoElementTrendlineAddLinearForecast"] = 603
	enum["msoElementTrendlineAddTwoPeriodMovingAverage"] = 604
	enum["msoElementTrendlineNone"] = 600
	enum["msoElementUpDownBarsNone"] = 900
	enum["msoElementUpDownBarsShow"] = 901

	return enum
}

func GetEnumChartElementTypeNum(enumType string) int32 {
	var result int32
	enum := EnumChartElementType()
	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["msoElementChartTitleNone"]
	}
	return result
}

func GetEnumChartElementTypeStr(enumType int32) string {
	var result string
	enum := EnumChartElementType()
	result = "msoElementChartTitleNone"
	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumChartElementType(enumType int32) int32 {
	var result int32
	enum := EnumChartElementType()
	result = enum["msoElementChartTitleNone"]
	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// XlChartLocation
func EnumChartLocation() map[string]int32 {
	enum := make(map[string]int32)
	enum["xlLocationAsNewSheet"] = 1
	enum["xlLocationAsObject"] = 2
	enum["xlLocationAutomatic"] = 3 //Default
	return enum
}

func GetEnumChartLocationNum(enumType string) int32 {
	var result int32
	enum := EnumChartLocation()
	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlLocationAutomatic"]
	}
	return result
}

func GetEnumChartLocationStr(enumType int32) string {
	var result string
	enum := EnumChartLocation()
	result = "xlLocationAutomatic"
	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumChartLocation(enumType int32) int32 {
	var result int32
	enum := EnumChartLocation()
	result = enum["xlLocationAutomatic"]
	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}

// XlMarkerStyle
func EnumMarkerStyle() map[string]int32 {
	enum := make(map[string]int32)
	enum["xlMarkerStyleAutomatic"] = -4105
	enum["xlMarkerStyleCircle"] = 8
	enum["xlMarkerStyleDash"] = -4115
	enum["xlMarkerStyleDiamond"] = 2
	enum["xlMarkerStyleDot"] = -4118
	enum["xlMarkerStyleNone"] = -4142 //Def
	enum["xlMarkerStylePicture"] = -4147
	enum["xlMarkerStylePlus"] = 9
	enum["xlMarkerStyleSquare"] = 1
	enum["xlMarkerStyleStar"] = 5
	enum["xlMarkerStyleTriangle"] = 3
	enum["xlMarkerStyleX"] = -4168
	return enum
}

func GetEnumMarkerStyleNum(enumType string) int32 {
	var result int32
	enum := EnumMarkerStyle()
	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlMarkerStyleNone"]
	}
	return result
}

func GetEnumMarkerStyleStr(enumType int32) string {
	var result string
	enum := EnumMarkerStyle()
	result = "xlMarkerStyleNone"
	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumMarkerStyle(enumType int32) int32 {
	var result int32
	enum := EnumMarkerStyle()
	result = enum["xlMarkerStyleNone"]
	for _, v := range enum {
		if v == enumType {
			result = v
			break
		}
	}
	return result
}
