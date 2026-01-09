package oleXL

import "log"

// version 2026-01-09

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
	enum["xlWorkbookDefault"] = 51
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
	result = "xlWorkbookNormal"

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
	result = enum["xlWorkbookNormal"]

	for k, v := range enum {
		if v == enumNum {
			log.Printf("check: %v %v\n", enumNum, k)
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

	for k, v := range enum {
		if v == enumNum {
			log.Printf("check: %v %v\n", enumNum, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
			result = v
			break
		}
	}
	return result
}

// LineStyle
func EnumLineStyle() map[string]int32 {
	enum := make(map[string]int32)

	enum["xlContinuous"] = 1
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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
	enum := EnumFlipCmd()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["msoFlipHorizontal"]
	}
	return result
}

func GetEnumFlipCmdStr(enumType int32) string {
	var result string
	enum := EnumFlipCmd()
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
	enum := EnumFlipCmd()
	result = enum["msoFlipHorizontal"]

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
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
	enum := EnumUpdateLinks()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlUpdateLinks"]
	}
	return result
}

func GetEnumUpdateLinksStr(enumType int32) string {
	var result string
	enum := EnumUpdateLinks()
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
	enum := EnumUpdateLinks()
	result = enum["xlUpdateLinks"]

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
			result = v
			break
		}
	}
	return result
}

// ChartType
func EnumChart() map[string]int32 {
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
	enum["xlLine"] = 4
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

func GetEnumChartNum(enumType string) int32 {
	var result int32
	enum := EnumChart()

	if v, ok := enum[enumType]; ok {
		result = v
	} else {
		result = enum["xlLine"]
	}
	return result
}

func GetEnumChartStr(enumType int32) string {
	var result string
	enum := EnumChart()
	result = "xlLine"

	for k, v := range enum {
		if v == enumType {
			result = k
			break
		}
	}
	return result
}

func SetEnumChart(enumType int32) int32 {
	var result int32
	enum := EnumChart()
	result = enum["xlLine"]

	for k, v := range enum {
		if v == enumType {
			log.Printf("check: %v %v\n", enumType, k)
			result = v
			break
		}
	}
	return result
}
