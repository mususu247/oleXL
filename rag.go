package oleXL

import (
	"fmt"
	"log"
	"time"

	"github.com/go-ole/go-ole"
)

// version 2026-01-04
// VBA style like

type workRag struct {
	app *Excel
	mx  int
}

func (rg *workRag) Nothing() error {
	xl := rg.app
	_, err := xl.getCore(rg.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	// wb.child.RelaseAll
	for k, v := range xl.WorkCores.cores {
		if v.px == rg.mx {
			xl.Release(k)
		}
	}
	return nil
}

func (ws *workSheet) Cells(row, col int) *workRag {
	var rg workRag
	rg.app = ws.app
	xl := ws.app

	_core, err := xl.getCore(ws.mx)
	if err != nil {
		log.Printf("not found: %v", ws.mx)
		return nil
	}

	const cmd = "Get"
	const name = "Cells"
	var opt []any
	opt = append(opt, int32(row))
	opt = append(opt, int32(col))

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			rg.mx, _core = xl.addCore(ws.mx, x, "Range", 0)
			_core.values["Row"] = row
			_core.values["Column"] = col
			return &rg
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (ws *workSheet) Range(value ...any) *workRag {
	var rg workRag
	rg.app = ws.app
	xl := ws.app

	_core, err := xl.getCore(ws.mx)
	if err != nil {
		log.Printf("not found: %v", ws.mx)
		return nil
	}

	const cmd = "Get"
	const name = "Range"
	var opt []any

	for i := range value {
		switch x := value[i].(type) {
		case string:
			opt = append(opt, x)
		case *workRag:
			_rag, err := xl.getCore(x.mx)
			if err == nil {
				opt = append(opt, _rag.disp)
			}
		}
	}

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

			rg.mx, _core = xl.addCore(ws.mx, x, "Range", 0)
			return &rg
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}

	return nil
}

func (rg *workRag) CurrentRegion() *workRag {
	var cur workRag
	xl := rg.app

	_core, err := xl.getCore(rg.mx)
	if err != nil {
		log.Printf("not found: %v", rg.mx)
		return nil
	}

	var ws workSheet
	ws.app = rg.app
	ws.mx = _core.px

	const cmd = "Get"
	const name = "CurrentRegion"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			cur.mx, _core = xl.addCore(ws.mx, x, "Range", 0)
			return &cur
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (rg *workRag) EntireColumn() *workRag {
	var cur workRag
	cur.app = rg.app
	xl := rg.app

	_core, err := xl.getCore(rg.mx)
	if err != nil {
		log.Printf("not found: %v", rg.mx)
		return nil
	}

	var ws workSheet
	ws.app = rg.app
	ws.mx = _core.px

	const cmd = "Get"
	const name = "EntireColumn"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			cur.mx, _core = xl.addCore(ws.mx, x, "Range", 0)
			return &cur
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (rg *workRag) EntireRow() *workRag {
	var cur workRag
	cur.app = rg.app
	xl := rg.app

	_core, err := xl.getCore(rg.mx)
	if err != nil {
		log.Printf("not found: %v", rg.mx)
		return nil
	}

	var ws workSheet
	ws.app = rg.app
	ws.mx = _core.px

	const cmd = "Get"
	const name = "EntireRow"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			cur.mx, _core = xl.addCore(ws.mx, x, "Range", 0)
			return &cur
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (rg *workRag) Columns() *workRag {
	var cur workRag
	cur.app = rg.app
	xl := rg.app

	_core, err := xl.getCore(rg.mx)
	if err != nil {
		log.Printf("not found: %v", rg.mx)
		return nil
	}

	var ws workSheet
	ws.app = rg.app
	ws.mx = _core.px

	const cmd = "Get"
	const name = "Columns"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			cur.mx, _core = xl.addCore(ws.mx, x, "Range", 0)
			return &cur
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (rg *workRag) Rows() *workRag {
	var cur workRag
	cur.app = rg.app
	xl := rg.app

	_core, err := xl.getCore(rg.mx)
	if err != nil {
		log.Printf("not found: %v", rg.mx)
		return nil
	}

	var ws workSheet
	ws.app = rg.app
	ws.mx = _core.px

	const cmd = "Get"
	const name = "Rows"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			cur.mx, _core = xl.addCore(ws.mx, x, "Range", 0)
			return &cur
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (rg *workRag) Activate() error {
	xl := rg.app
	_core, err := xl.getCore(rg.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	const cmd = "Get"
	const name = "Activate"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			return x
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			if x {
				xl.setCoreValue(xl.mx, "ActiveCell", rg.mx)
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}

func (rg *workRag) Select() error {
	xl := rg.app
	_core, err := xl.getCore(rg.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	const cmd = "Get"
	const name = "Select"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			return x
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			if x {
				xl.setCoreValue(xl.mx, "Selection", rg.mx)
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}

func (rg *workRag) Value(value ...any) any {
	var result any
	xl := rg.app
	_core, err := xl.getCore(rg.mx)
	if err != nil {
		return -1
	}
	var cmd string
	const name = "Value"
	var opt []any

	if len(value) > 0 {
		switch x := value[0].(type) {
		case int:
			opt = append(opt, float64(x))
		case float64:
			opt = append(opt, x)
		case string:
			opt = append(opt, x)
		case bool:
			opt = append(opt, x)
		case time.Time:
			opt = append(opt, x)
		default:
			opt = nil
		}
	} else {
		opt = nil
	}

	if opt != nil {
		cmd = "Put"
		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case nil:
				log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	result = nil
	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case float64:
			log.Printf("%v ans (float64) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["Value"] = result
		case string:
			log.Printf("%v ans (string) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["Value"] = result
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["Value"] = result
		case time.Time:
			log.Printf("%v ans (time) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["Value"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return result
}

func (rg *workRag) Formula(value ...string) string {
	var result string
	xl := rg.app
	_core, err := xl.getCore(rg.mx)
	if err != nil {
		return ""
	}
	var cmd string
	const name = "Formula"
	var opt []any

	if len(value) > 0 {
		opt = append(opt, value[0])
	} else {
		opt = nil
	}

	if opt != nil {
		cmd = "Put"
		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case nil:
				log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case string:
			log.Printf("%v ans (string) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["Formula"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return result
}

func (rg *workRag) FormulaR1C1(value ...string) string {
	var result string
	xl := rg.app
	_core, err := xl.getCore(rg.mx)
	if err != nil {
		return ""
	}
	var cmd string
	const name = "FormulaR1C1"
	var opt []any

	if len(value) > 0 {
		opt = append(opt, value[0])
	} else {
		opt = nil
	}

	if opt != nil {
		cmd = "Put"
		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case nil:
				log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case string:
			log.Printf("%v ans (string) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["FormulaR1C1"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return result
}

func (rg *workRag) End(value any) *workRag {
	var area workRag
	area.app = rg.app
	xl := rg.app
	_core, err := xl.getCore(rg.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	const cmd = "Get"
	const name = "End"
	var opt []any

	var z int32
	switch v := value.(type) {
	case int:
		z = SetEnumDirection(int32(v))
	case int32:
		z = SetEnumDirection(v)
	case string:
		z = GetEnumDirectionNum(v)
	}
	opt = append(opt, z)

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			ws := xl.ActiveSheet()
			area.mx, _core = xl.addCore(ws.mx, x, "Range", 0)
			return &area
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}

func (rg *workRag) Copy() error {
	xl := rg.app
	_core, err := xl.getCore(rg.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	const cmd = "Method"
	const name = "Copy"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			return x
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}

func (rg *workRag) PasteSpecial(Paste, Operation, SkipBlanks, Transpose any) error {
	xl := rg.app
	_core, err := xl.getCore(rg.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	const cmd = "Method"
	const name = "PasteSpecial"
	var opt []any

	var z1 int32
	switch v := Paste.(type) {
	case int:
		z1 = SetEnumPaste(int32(v))
	case int32:
		z1 = SetEnumPaste(v)
	case string:
		z1 = GetEnumPasteNum(v)
	default:
		z1 = GetEnumPasteNum("")
	}
	opt = append(opt, z1)

	var z2 int32
	switch v := Operation.(type) {
	case int:
		z2 = SetEnumPasteOperation(int32(v))
	case int32:
		z2 = SetEnumPasteOperation(v)
	case string:
		z2 = GetEnumPasteOperationNum(v)
	}
	opt = append(opt, z2)

	switch v := SkipBlanks.(type) {
	case bool:
		opt = append(opt, v)
	default:
		opt = append(opt, false)
	}

	switch v := Transpose.(type) {
	case bool:
		opt = append(opt, v)
	default:
		opt = append(opt, false)
	}

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			return x
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return nil
}

func (rg *workRag) AutoFit() error {
	xl := rg.app
	_core, err := xl.getCore(rg.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	const cmd = "Get"
	const name = "AutoFit"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			return x
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}
