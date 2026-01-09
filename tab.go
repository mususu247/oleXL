package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

// version 2026-01-05
// VBA style like

type workTab struct {
	app *Excel
	mx  int
}

func (ws *workSheet) Tab() *workTab {
	var tb workTab
	tb.app = ws.app
	xl := ws.app

	_core, err := xl.getCore(ws.mx)
	if err != nil {
		log.Printf("not found: %v", ws.mx)
		return nil
	}

	const cmd = "Get"
	const name = "Tab"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			tb.mx, _ = xl.addCore(ws.mx, x, "Tab", 0)
			return &tb
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (tb *workTab) ColorIndex(value ...any) int32 {
	var result int32
	xl := tb.app
	_core, err := xl.getCore(tb.mx)
	if err != nil {
		return -1
	}
	var cmd string
	const name = "ColorIndex"
	var opt []any

	if len(value) > 0 {
		switch x := value[0].(type) {
		case int:
			opt = append(opt, int32(x))
		case int32:
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

	result = -1
	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case int32:
			log.Printf("%v ans (int32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["ColorIndex"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return result
}
