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

func (ws *workSheet) Range(cell ...string) *workRange {
	var wr workRange
	xl := ws.app

	name := "Range"
	core, num := xl.cores.FindAdd(name, ws.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		for i := range cell {
			opt = append(opt, cell[i])
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

func (ws *workSheet) Cells(cell ...int32) *workRange {
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

func (wr *workRange) Cells(cell ...int32) *workRange {
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
			core.lock = 0
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
	xl := wr.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, wr.num, nil)
	if err != nil {
		return err
	}
	return nil
}
