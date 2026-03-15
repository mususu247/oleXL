package oleXL

import (
	"log"
	"time"

	"github.com/go-ole/go-ole"
)

type workRange struct {
	app    *Excel
	parent *workSheet
	num    int
}

func (ws *workSheet) Range(cell ...any) *workRange {
	var wr workRange
	xl := ws.app

	name := "Range"
	core, num := xl.cores.FindAdd(name, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		for i := range cell {
			switch x := cell[i].(type) {
			case string:
				opt = append(opt, x)
			case *workRange:
				opt = append(opt, xl.cores.getCore(x.num).disp)
			}
		}

		ans, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	wr.app = xl
	wr.num = num
	wr.parent = ws
	return &wr
}

func (ws *workSheet) Cells(cell ...any) *workRange {
	var wr workRange
	xl := ws.app

	kind := "Range"
	name := "Cells"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		if len(cell) > 0 {
			opt = append(opt, cell[0])
		}
		if len(cell) > 1 {
			opt = append(opt, cell[1])
		}

		ans, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	wr.app = xl
	wr.num = num
	wr.parent = ws
	return &wr
}

func (wr *workRange) Cells(cell ...any) *workRange {
	var xr workRange
	xl := wr.app
	ws := wr.parent

	kind := "Range"
	name := "Cells"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		if len(cell) > 0 {
			opt = append(opt, cell[0])
		}
		if len(cell) > 1 {
			opt = append(opt, cell[1])
		}

		ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xr.app = xl
	xr.num = num
	xr.parent = ws
	return &xr
}

func (ws *workSheet) Rows(cell any) *workRange {
	var wr workRange
	xl := ws.app

	kind := "Range"
	name := "Rows"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		opt = append(opt, cell)

		ans, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	wr.app = xl
	wr.num = num
	wr.parent = ws
	return &wr
}

func (wr *workRange) Rows(cell any) *workRange {
	var xr workRange
	xl := wr.app
	ws := wr.parent

	kind := "Range"
	name := "Rows"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		opt = append(opt, cell)

		ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xr.app = xl
	xr.num = num
	xr.parent = ws
	return &xr
}

func (ws *workSheet) Columns(cell any) *workRange {
	var wr workRange
	xl := ws.app

	kind := "Range"
	name := "Columns"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		opt = append(opt, cell)

		ans, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	wr.app = xl
	wr.num = num
	wr.parent = ws
	return &wr
}

func (wr *workRange) Columns(cell any) *workRange {
	var xr workRange
	xl := wr.app
	ws := wr.parent

	kind := "Range"
	name := "Columns"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		opt = append(opt, cell)

		ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xr.app = xl
	xr.num = num
	xr.parent = ws
	return &xr
}

func (wr *workRange) End(shift any) *workRange {
	var xr workRange
	xl := wr.app
	ws := wr.parent

	kind := "Range"
	name := "End"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any

		var z int32
		switch x := shift.(type) {
		case int:
			z = SetEnumDirection(int32(x))
		case int32:
			z = SetEnumDirection(x)
		case string:
			z = GetEnumDirectionNum(x)
		}
		opt = append(opt, z)

		ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xr.app = xl
	xr.num = num
	xr.parent = ws
	return &xr
}

func (wr *workRange) Delete(shift ...any) *workRange {
	var xr workRange
	xl := wr.app
	ws := wr.parent

	kind := "Range"
	name := "Delete"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any

		if len(shift) > 0 {
			var z int32
			switch x := shift[0].(type) {
			case int:
				z = SetEnumDirection(int32(x))
			case int32:
				z = SetEnumDirection(x)
			case string:
				z = GetEnumDirectionNum(x)
			}
			opt = append(opt, z)
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xr.app = xl
	xr.num = num
	xr.parent = ws
	return &xr
}

func (wr *workRange) Insert(shift ...any) *workRange {
	var xr workRange
	xl := wr.app
	ws := wr.parent

	kind := "Range"
	name := "Insert"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any

		if len(shift) > 0 {
			var z int32
			switch x := shift[0].(type) {
			case int:
				z = SetEnumDirection(int32(x))
			case int32:
				z = SetEnumDirection(x)
			case string:
				z = GetEnumDirectionNum(x)
			}
			opt = append(opt, z)
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xr.app = xl
	xr.num = num
	xr.parent = ws
	return &xr
}

func (wr *workRange) CurrentRegion() *workRange {
	var xr workRange
	xl := wr.app
	ws := wr.parent

	kind := "Range"
	name := "CurrentRegion"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xr.app = xl
	xr.num = num
	xr.parent = ws
	return &xr
}

func (wr *workRange) MergeArea() *workRange {
	var xr workRange
	xl := wr.app
	ws := wr.parent

	kind := "Range"
	name := "MergeArea"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xr.app = xl
	xr.num = num
	xr.parent = ws
	return &xr
}

func (wr *workRange) Offset(RowOffset int32, ColumnOffset int32) *workRange {
	var xr workRange
	xl := wr.app
	ws := wr.parent

	kind := "Range"
	name := "Offset"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		opt = append(opt, RowOffset)
		opt = append(opt, ColumnOffset)

		ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xr.app = xl
	xr.num = num
	xr.parent = ws
	return &xr
}

func (wr *workRange) Resize(RowSize int32, ColumnSize int32) *workRange {
	var xr workRange
	xl := wr.app
	ws := wr.parent

	kind := "Range"
	name := "Resize"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		opt = append(opt, RowSize)
		opt = append(opt, ColumnSize)

		ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xr.app = xl
	xr.num = num
	xr.parent = ws
	return &xr
}

func (wr *workRange) Address(options ...map[string]any) string {
	xl := wr.app

	name := "Address"
	cmd := "Get"
	var opt []any

	//Default: RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=True, RelativeTo:=Nothing
	opt = append(opt, true)                             //RowAbsolute
	opt = append(opt, true)                             //ColumnAbsolute
	opt = append(opt, GetEnumReferenceStyleNum("xlA1")) //ReferenceStyle
	opt = append(opt, false)                            //External
	opt = append(opt, nil)                              //RelativeTo

	if len(options) > 0 {
		for k, v := range options[0] {
			switch k {
			case "RowAbsolute":
				switch x := v.(type) {
				case bool:
					opt[0] = x
				}
			case "ColumnAbsolute":
				switch x := v.(type) {
				case bool:
					opt[1] = x
				}
			case "ReferenceStyle":
				switch x := v.(type) {
				case string:
					opt[2] = GetEnumReferenceStyleNum(x)
				}
			case "External":
				switch x := v.(type) {
				case bool:
					opt[3] = x
				}
			case "RelativeTo":
				switch x := v.(type) {
				case *workRange:
					opt[4] = xl.cores.getCore(x.num).disp
				}
			}
		}
	}

	ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
	if err != nil {
		log.Printf("(Error) %v", err)
		return ""
	}
	switch x := ans.(type) {
	case string:
		return x
	}

	return ""
}

func (wr *workRange) EntireRow() *workRange {
	var xr workRange
	xl := wr.app
	ws := wr.parent

	kind := "Range"
	name := "EntireRow"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xr.app = xl
	xr.num = num
	xr.parent = ws
	return &xr
}

func (wr *workRange) EntireColumn() *workRange {
	var xr workRange
	xl := wr.app
	ws := wr.parent

	kind := "Range"
	name := "EntireColumn"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xr.app = xl
	xr.num = num
	xr.parent = ws
	return &xr
}

func (xl *Excel) ActiveCell() *workRange {
	var wr workRange
	ws := xl.ActiveSheet()

	kind := "Range"
	name := "ActiveCell"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	wr.app = xl
	wr.num = num
	wr.parent = ws
	ws.Release()
	return &wr
}

func (xl *Excel) Selection() *workRange {
	var wr workRange
	ws := xl.ActiveSheet()

	kind := "Range"
	name := "Selection"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 1 //Lock on
		}
	}
	wr.app = xl
	wr.num = num
	wr.parent = ws
	ws.Release()
	return &wr
}

func (wr *workRange) Release() error {
	xl := wr.app
	xl.cores.Release(wr.num, false)
	return nil
}

func (wr *workRange) Nothing() error {
	xl := wr.app
	xl.cores.releaseChild(wr.num)

	xl.cores.Unlock(wr.num)
	err := wr.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(wr.num)
	wr = nil
	return nil
}

func (wr *workRange) Set() *workRange {
	if wr == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := wr.app
	xl.cores.Lock(wr.num)
	return wr
}

func (wr *workRange) Value(value ...any) any {
	xl := wr.app

	name := "Value"
	if len(value) > 0 {
		var f64 float64

		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case int:
			f64 = float64(x)
			opt = append(opt, f64)
		case int32:
			f64 = float64(x)
			opt = append(opt, f64)
		case int64:
			f64 = float64(x)
			opt = append(opt, f64)
		case float32:
			f64 = float64(x)
			opt = append(opt, f64)
		case float64:
			opt = append(opt, x)
		case string:
			opt = append(opt, x)
		case bool:
			opt = append(opt, x)
		case time.Time:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case float64:
			return x
		case string:
			return x
		case bool:
			return x
		case time.Time:
			return x
		case *ole.VARIANT:
			switch x.Val {
			case 2148141008:
				return "#NULL!"
			case 2148141015:
				return "#DIV/0!"
			case 2148141023:
				return "#VALUE!"
			case 2148141031:
				return "#REF!"
			case 2148141037:
				return "#NAME?"
			case 2148141044:
				return "#NUM!"
			case 2148141050:
				return "#N/A"
			}
			return x.Val
		default:
			return x
		}
	}
	return nil
}

func (wr *workRange) Value2(value ...any) any {
	xl := wr.app

	name := "Value2"
	if len(value) > 0 {
		var f64 float64

		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case int:
			f64 = float64(x)
			opt = append(opt, f64)
		case int32:
			f64 = float64(x)
			opt = append(opt, f64)
		case int64:
			f64 = float64(x)
			opt = append(opt, f64)
		case float32:
			f64 = float64(x)
			opt = append(opt, f64)
		case float64:
			opt = append(opt, x)
		case string:
			opt = append(opt, x)
		case bool:
			opt = append(opt, x)
		case time.Time:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case float64:
			return x
		case string:
			return x
		case bool:
			return x
		case time.Time:
			return x
		case *ole.VARIANT:
			switch x.Val {
			case 2148141008:
				return "#NULL!"
			case 2148141015:
				return "#DIV/0!"
			case 2148141023:
				return "#VALUE!"
			case 2148141031:
				return "#REF!"
			case 2148141037:
				return "#NAME?"
			case 2148141044:
				return "#NUM!"
			case 2148141050:
				return "#N/A"
			}
			return x.Val
		default:
			return x
		}
	}
	return nil
}

func (wr *workRange) Formula(value ...string) string {
	xl := wr.app

	name := "Formula"
	if len(value) > 0 {

		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
		switch x := ans.(type) {
		case string:
			return x
		}
	}
	return ""
}

func (wr *workRange) Formula2(value ...string) string {
	xl := wr.app

	name := "Formula2"
	if len(value) > 0 {

		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
		switch x := ans.(type) {
		case string:
			return x
		}
	}
	return ""
}

func (wr *workRange) FormulaR1C1(value ...string) string {
	xl := wr.app

	name := "FormulaR1C1"
	if len(value) > 0 {

		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
		switch x := ans.(type) {
		case string:
			return x
		}
	}
	return ""
}

func (wr *workRange) Formula2R1C1(value ...string) string {
	xl := wr.app

	name := "Formula2R1C1"
	if len(value) > 0 {

		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
		switch x := ans.(type) {
		case string:
			return x
		}
	}
	return ""
}

func (wr *workRange) Activate() error {
	xl := wr.app

	cmd := "Method"
	name := "Activate"

	_, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (wr *workRange) Select() error {
	if wr == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := wr.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (wr *workRange) Copy(value ...any) bool {
	xl := wr.app

	cmd := "Method"
	name := "Copy"
	var opt []any
	if len(value) > 0 {
		switch x := value[0].(type) {
		case *workRange:
			core := xl.cores.getCore(x.num)
			if core.disp != nil {
				opt = append(opt, core.disp)
			}
		}
	} else {
		opt = nil
	}

	ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
	if err != nil {
		return false
	}
	switch x := ans.(type) {
	case bool:
		return x
	}
	return false
}

func (wr *workRange) Cut() bool {
	xl := wr.app

	cmd := "Method"
	name := "Cut"

	ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		return false
	}
	switch x := ans.(type) {
	case bool:
		return x
	}
	return false
}

func (wr *workRange) PasteSpecial(Paste any, Operation any, SkipBlanks any, Transpose any) bool {
	xl := wr.app

	cmd := "Method"
	name := "PasteSpecial"
	var opt []any
	var z int32

	switch x := Paste.(type) {
	case int:
		z = SetEnumPaste(int32(x))
	case int32:
		z = SetEnumPaste(x)
	case string:
		z = GetEnumPasteNum(x)
	default:
		opt = append(opt, nil)
	}
	opt = append(opt, z)

	switch x := Operation.(type) {
	case int:
		z = SetEnumPasteOperation(int32(x))
	case int32:
		z = SetEnumPasteOperation(x)
	case string:
		z = GetEnumPasteOperationNum(x)
	default:
		opt = append(opt, nil)
	}
	opt = append(opt, z)

	switch x := SkipBlanks.(type) {
	case bool:
		opt = append(opt, x)
	default:
		opt = append(opt, nil)
	}

	switch x := Transpose.(type) {
	case bool:
		opt = append(opt, x)
	default:
		opt = append(opt, nil)
	}

	ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
	if err != nil {
		return false
	}
	switch x := ans.(type) {
	case bool:
		return x
	}
	return false
}

func (wr *workRange) Paste() bool {
	return wr.PasteSpecial("xlPasteAll", "xlPasteSpecialOperationNone", false, false)
}

func (wr *workRange) Clear() error {
	xl := wr.app

	cmd := "Method"
	name := "Clear"

	_, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (wr *workRange) ClearComments() error {
	xl := wr.app

	cmd := "Method"
	name := "ClearComments"

	_, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (wr *workRange) ClearContents() error {
	xl := wr.app

	cmd := "Method"
	name := "ClearContents"

	_, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (wr *workRange) ClearFormats() error {
	xl := wr.app

	cmd := "Method"
	name := "ClearFormats"

	_, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (wr *workRange) ClearHyperlinks() error {
	xl := wr.app

	cmd := "Method"
	name := "ClearHyperlinks"

	_, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (wr *workRange) Count() int32 {
	var result int32
	xl := wr.app

	cmd := "Get"
	name := "Count"
	ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return result
	}

	switch x := ans.(type) {
	case int32:
		result = x
	}
	return result
}

func (wr *workRange) Row() int32 {
	var result int32
	xl := wr.app

	cmd := "Get"
	name := "Row"
	ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return result
	}

	switch x := ans.(type) {
	case int32:
		result = x
	}
	return result
}

func (wr *workRange) Column() int32 {
	var result int32
	xl := wr.app

	cmd := "Get"
	name := "Column"
	ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return result
	}

	switch x := ans.(type) {
	case int32:
		result = x
	}
	return result
}

func (wr *workRange) NumberFormatLocal(value ...string) string {
	xl := wr.app

	name := "NumberFormatLocal"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
		switch x := ans.(type) {
		case string:
			return x
		}
	}
	return ""
}

func (wr *workRange) HorizontalAlignment(value ...any) int32 {
	xl := wr.app

	name := "HorizontalAlignment"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumHAlign(int32(x))
		case int32:
			z = SetEnumHAlign(x)
		case string:
			z = GetEnumHAlignNum(x)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		_, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	}
	return 0
}

func (wr *workRange) VerticalAlignment(value ...any) int32 {
	xl := wr.app

	name := "VerticalAlignment"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumVAlign(int32(x))
		case int32:
			z = SetEnumVAlign(x)
		case string:
			z = GetEnumVAlignNum(x)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		_, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	}
	return 0
}

func (wr *workRange) WrapText(value ...bool) bool {
	xl := wr.app

	name := "WrapText"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
		switch x := ans.(type) {
		case bool:
			return x
		}
	}

	return false
}

func (wr *workRange) Orientation(value ...float64) float64 {
	xl := wr.app

	name := "Orientation"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case float64:
			return x
		}
	}

	return 0
}

func (wr *workRange) AddIndent(value ...bool) bool {
	xl := wr.app

	name := "AddIndent"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
		switch x := ans.(type) {
		case bool:
			return x
		}
	}

	return false
}

func (wr *workRange) IndentLevel(value ...int32) int32 {
	xl := wr.app

	name := "IndentLevel"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case int32:
			return x
		}
	}

	return 0
}

func (wr *workRange) ShrinkToFit(value ...bool) bool {
	xl := wr.app

	name := "ShrinkToFit"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
		switch x := ans.(type) {
		case bool:
			return x
		}
	}

	return false
}

func (wr *workRange) ReadingOrder(value ...any) int32 {
	xl := wr.app

	name := "ReadingOrder"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumReadingOrder(int32(x))
		case int32:
			z = SetEnumReadingOrder(x)
		case string:
			z = GetEnumReadingOrderNum(x)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		_, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	}
	return 0
}

func (wr *workRange) MergeCells(value ...bool) bool {
	xl := wr.app

	name := "MergeCells"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
		switch x := ans.(type) {
		case bool:
			return x
		}
	}

	return false
}

func (wr *workRange) RowHeight(value ...float64) float64 {
	xl := wr.app

	name := "RowHeight"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case float64:
			return x
		}
	}

	return 0
}

func (wr *workRange) ColumnWidth(value ...float64) float64 {
	xl := wr.app

	name := "ColumnWidth"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case float64:
			return x
		}
	}

	return 0
}

func (wr *workRange) Height() float64 {
	xl := wr.app

	name := "Height"
	cmd := "Get"

	ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return 0
	}
	switch x := ans.(type) {
	case float64:
		return x
	}

	return 0
}

func (wr *workRange) Width() float64 {
	xl := wr.app

	name := "Width"
	cmd := "Get"

	ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return 0
	}
	switch x := ans.(type) {
	case float64:
		return x
	}

	return 0
}

func (wr *workRange) AutoFit() error {
	xl := wr.app

	cmd := "Method"
	name := "AutoFit"

	_, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		return err
	}
	return nil
}
