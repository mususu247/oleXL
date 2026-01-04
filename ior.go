package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

// version 2026-01-04
// VBA style like

type workInterior struct {
	app *Excel
	mx  int
}

func (rg *workRag) Interior() *workInterior {
	var ig workInterior
	ig.app = rg.app
	xl := rg.app

	_core, err := xl.getCore(rg.mx)
	if err != nil {
		log.Printf("not found: %v", rg.mx)
		return nil
	}

	const cmd = "Get"
	const name = "Interior"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			ig.mx, _ = xl.addCore(rg.mx, x, "Interior", 0)
			return &ig
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (ior *workInterior) ColorIndex(value ...any) int32 {
	var result int32
	xl := ior.app
	_core, err := xl.getCore(ior.mx)
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

func (ior *workInterior) Color(value ...float64) float64 {
	xl := ior.app
	_core, err := xl.getCore(ior.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	var cmd string
	const name = "Color"
	var opt []any

	if len(value) > 0 {
		cmd = "Put"
		opt = append(opt, value[0])
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
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case float64:
			log.Printf("%v ans (float64) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return -1
}
