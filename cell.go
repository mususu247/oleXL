package oleXL

import (
	"fmt"
	"log"
	"strconv"
	"strings"
	"time"

	"github.com/atotto/clipboard"
	"github.com/go-ole/go-ole"
)

type workRange struct {
	app    *Excel
	parent any
	num    int
}

func getSheet(v *workRange) *workSheet {
	var w any
	w = v

	for {
		switch x := w.(type) {
		case *workRange:
			w = x.parent
		case *workSheet:
			return x
		}
	}
}

func (Q *workSheet) Range(cell ...any) *workRange {
	var body workRange
	xl := Q.app

	name := "Range"
	core, num := xl.cores.FindAdd(name, Q.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		for i := range cell {
			switch x := cell[i].(type) {
			case string:
				opt = append(opt, x)
			case int:
				opt = append(opt, int32(x))
			case int32:
				opt = append(opt, x)
			case *workRange:
				opt = append(opt, xl.cores.getCore(x.num).disp)
			}
		}

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = Q
	return &body
}

func (Q *workSheet) Cells(cell ...any) *workRange {
	var body workRange
	xl := Q.app

	kind := "Range"
	name := "Cells"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		if len(cell) > 0 {
			opt = append(opt, cell[0])
		}
		if len(cell) > 1 {
			opt = append(opt, cell[1])
		}

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = Q
	return &body
}

func (Q *workRange) Cells(cell ...any) *workRange {
	var body workRange
	var ws *workSheet
	xl := Q.app

	sw := true
	for sw {
		var w *workRange
		w = Q

		switch x := w.parent.(type) {
		case *workSheet:
			ws = x
			sw = false
		case *workRange:
			w = x
		default:
			sw = false
		}
	}

	if ws == nil {
		ws = xl.ActiveSheet()
	}

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

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (ws *workSheet) Rows(cell ...any) *workRange {
	var body workRange
	xl := ws.app

	kind := "Range"
	name := "Rows"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		if len(cell) > 0 {
			opt = append(opt, cell[0])
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (Q *workRange) Rows(cell ...any) *workRange {
	var body workRange
	xl := Q.app
	ws := getSheet(Q)

	kind := "Range"
	name := "Rows"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		if len(cell) > 0 {
			opt = append(opt, cell[0])
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (ws *workSheet) Columns(cell ...any) *workRange {
	var body workRange
	xl := ws.app

	kind := "Range"
	name := "Columns"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		if len(cell) > 0 {
			opt = append(opt, cell[0])
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (Q *workRange) Columns(cell ...any) *workRange {
	var body workRange
	xl := Q.app
	ws := getSheet(Q)

	kind := "Range"
	name := "Columns"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		if len(cell) > 0 {
			opt = append(opt, cell[0])
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (Q *workRange) End(shift any) *workRange {
	var body workRange
	xl := Q.app
	ws := getSheet(Q)

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

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (Q *workRange) Delete(shift ...any) *workRange {
	var body workRange
	xl := Q.app
	ws := getSheet(Q)

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

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (Q *workRange) Insert(shift ...any) *workRange {
	var body workRange
	xl := Q.app
	ws := getSheet(Q)

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

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (Q *workRange) CurrentRegion() *workRange {
	var body workRange
	xl := Q.app
	ws := getSheet(Q)

	kind := "Range"
	name := "CurrentRegion"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (Q *workRange) MergeArea() *workRange {
	var body workRange
	xl := Q.app
	ws := getSheet(Q)

	kind := "Range"
	name := "MergeArea"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (Q *workRange) Offset(RowOffset int32, ColumnOffset int32) *workRange {
	var body workRange
	xl := Q.app
	ws := getSheet(Q)

	kind := "Range"
	name := "Offset"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		opt = append(opt, RowOffset)
		opt = append(opt, ColumnOffset)

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (Q *workRange) Resize(RowSize int32, ColumnSize int32) *workRange {
	var body workRange
	xl := Q.app
	ws := getSheet(Q)

	kind := "Range"
	name := "Resize"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		opt = append(opt, RowSize)
		opt = append(opt, ColumnSize)

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (Q *workRange) Address(options ...map[string]any) string {
	xl := Q.app

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

	ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
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

func (Q *workRange) EntireRow() *workRange {
	var body workRange
	xl := Q.app
	ws := getSheet(Q)

	kind := "Range"
	name := "EntireRow"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (Q *workRange) EntireColumn() *workRange {
	var body workRange
	xl := Q.app
	ws := getSheet(Q)

	kind := "Range"
	name := "EntireColumn"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	return &body
}

func (Q *Excel) ActiveCell() *workRange {
	var body workRange
	xl := Q
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
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	ws.Release()
	return &body
}

func (Q *Excel) Selection() *workRange {
	var body workRange
	xl := Q
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
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = ws
	ws.Release()
	return &body
}

func (Q *workRange) Release() error {
	xl := Q.app
	xl.cores.Release(Q.num, false)
	return nil
}

func (Q *workRange) Nothing() error {
	xl := Q.app
	xl.cores.releaseChild(Q.num)

	xl.cores.Unlock(Q.num)
	err := Q.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(Q.num)
	Q = nil
	return nil
}

func (Q *workRange) Set() (*workRange, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workRange) Value(value ...any) any {
	xl := Q.app

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

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) Value2(value ...any) any {
	xl := Q.app

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

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) Values() [][]any {
	var result [][]any
	//xl := Q.app

	rows := int(Q.Rows().Count())

	var nfs []string
	for c := range Q.Columns().Count() {
		v := Q.Columns(c + 1).NumberFormatLocal()
		v = NumberFoarmat2Layout(v)
		nfs = append(nfs, v)
	}

	Q.Copy()
	text, _ := clipboard.ReadAll()
	clipboard.WriteAll("") //clipboard.Claer

	line := strings.Split(text, "\r\n")
	for r := range line {
		cell := strings.Split(line[r], "\t")
		var record []any
		for c := range cell {
			//value null
			if len(cell[c]) == 0 {
				record = append(record, nil)
				continue
			}

			//value number
			f64, err := strconv.ParseFloat(cell[c], 64)
			if err == nil {
				record = append(record, f64)
				continue
			}

			//value bool
			bl, err := strconv.ParseBool(cell[c])
			if err == nil {
				record = append(record, bl)
				continue
			}

			//value time
			dt, err := time.Parse(nfs[c], cell[c])
			if err == nil {
				record = append(record, dt)
				continue
			}

			//value string
			record = append(record, cell[c])
		}

		if r < rows {
			result = append(result, record)
			//log.Printf("[%v] %v\n", r, record)
		}
	}

	return result
}

func (Q *workRange) Formula(value ...string) string {
	xl := Q.app

	name := "Formula"
	if len(value) > 0 {

		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) Formula2(value ...string) string {
	xl := Q.app

	name := "Formula2"
	if len(value) > 0 {

		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) FormulaR1C1(value ...string) string {
	xl := Q.app

	name := "FormulaR1C1"
	if len(value) > 0 {

		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) Formula2R1C1(value ...string) string {
	xl := Q.app

	name := "Formula2R1C1"
	if len(value) > 0 {

		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) Activate() error {
	xl := Q.app

	cmd := "Method"
	name := "Activate"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workRange) Select() error {
	if Q == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := Q.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workRange) Copy(value ...any) bool {
	xl := Q.app

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

	ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
	if err != nil {
		return false
	}
	switch x := ans.(type) {
	case bool:
		return x
	}
	return false
}

func (Q *workRange) Cut() bool {
	xl := Q.app

	cmd := "Method"
	name := "Cut"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return false
	}
	switch x := ans.(type) {
	case bool:
		return x
	}
	return false
}

func (Q *workRange) PasteSpecial(Paste any, Operation any, SkipBlanks any, Transpose any) bool {
	xl := Q.app

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

	ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
	if err != nil {
		return false
	}
	switch x := ans.(type) {
	case bool:
		return x
	}
	return false
}

func (Q *workRange) Paste() bool {
	return Q.PasteSpecial("xlPasteAll", "xlPasteSpecialOperationNone", false, false)
}

func (Q *workRange) Clear() error {
	xl := Q.app

	cmd := "Method"
	name := "Clear"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workRange) ClearComments() error {
	xl := Q.app

	cmd := "Method"
	name := "ClearComments"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workRange) ClearContents() error {
	xl := Q.app

	cmd := "Method"
	name := "ClearContents"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workRange) ClearFormats() error {
	xl := Q.app

	cmd := "Method"
	name := "ClearFormats"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workRange) ClearHyperlinks() error {
	xl := Q.app

	cmd := "Method"
	name := "ClearHyperlinks"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workRange) Count() int32 {
	var result int32
	xl := Q.app

	cmd := "Get"
	name := "Count"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) Row() int32 {
	var result int32
	xl := Q.app

	cmd := "Get"
	name := "Row"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) Column() int32 {
	var result int32
	xl := Q.app

	cmd := "Get"
	name := "Column"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) NumberFormatLocal(value ...string) string {
	xl := Q.app

	name := "NumberFormatLocal"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) HorizontalAlignment(value ...any) int32 {
	xl := Q.app

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

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	}
	return 0
}

func (Q *workRange) VerticalAlignment(value ...any) int32 {
	xl := Q.app

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

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	}
	return 0
}

func (Q *workRange) WrapText(value ...bool) bool {
	xl := Q.app

	name := "WrapText"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) Orientation(value ...float64) float64 {
	xl := Q.app

	name := "Orientation"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) AddIndent(value ...bool) bool {
	xl := Q.app

	name := "AddIndent"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) IndentLevel(value ...int32) int32 {
	xl := Q.app

	name := "IndentLevel"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) ShrinkToFit(value ...bool) bool {
	xl := Q.app

	name := "ShrinkToFit"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) ReadingOrder(value ...any) int32 {
	xl := Q.app

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

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	}
	return 0
}

func (Q *workRange) MergeCells(value ...bool) bool {
	xl := Q.app

	name := "MergeCells"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) RowHeight(value ...float64) float64 {
	xl := Q.app

	name := "RowHeight"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) ColumnWidth(value ...float64) float64 {
	xl := Q.app

	name := "ColumnWidth"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) Left() float64 {
	xl := Q.app

	name := "Left"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) Top() float64 {
	xl := Q.app

	name := "Top"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) Height() float64 {
	xl := Q.app

	name := "Height"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) Width() float64 {
	xl := Q.app

	name := "Width"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workRange) AutoFit() error {
	xl := Q.app

	cmd := "Method"
	name := "AutoFit"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}
