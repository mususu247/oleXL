package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

// version 2026-01-05
// VBA style like

type workBorder struct {
	app *Excel
	mx  int
}

func (rg *workRag) Borders(value ...any) *workBorder {
	var bd workBorder
	bd.app = rg.app
	xl := rg.app

	_core, err := xl.getCore(rg.mx)
	if err != nil {
		log.Printf("not found: %v", rg.mx)
		return nil
	}

	const cmd = "Get"
	const name = "Borders"
	var opt []any

	if len(value) > 0 {
		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumBorders(int32(x))
		case int32:
			z = SetEnumBorders(x)
		case string:
			z = GetEnumBordersNum(x)
		}
		opt = append(opt, z)
	}

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

			bd.mx, _ = xl.addCore(rg.mx, x, "Borders", 0)
			return &bd
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}

	return nil
}

func (bd *workBorder) ColorIndex(value ...any) int32 {
	if bd == nil {
		return -1
	}
	var result int32
	xl := bd.app
	_core, err := xl.getCore(bd.mx)
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

func (bd *workBorder) LineStyle(value ...any) int32 {
	var result int32
	xl := bd.app
	_core, err := xl.getCore(bd.mx)
	if err != nil {
		return -1
	}
	var cmd string
	const name = "LineStyle"
	var opt []any

	if len(value) > 0 {
		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumLineStyle(int32(x))
		case int32:
			z = SetEnumLineStyle(x)
		case string:
			z = GetEnumLineStyleNum(x)
		default:
			opt = nil
		}
		opt = append(opt, z)
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
			_core.values["LineStyle"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return result
}

func (bd *workBorder) Weight(value ...any) int32 {
	var result int32
	xl := bd.app
	_core, err := xl.getCore(bd.mx)
	if err != nil {
		return -1
	}
	var cmd string
	const name = "Weight"
	var opt []any

	if len(value) > 0 {
		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumWeight(int32(x))
		case int32:
			z = SetEnumWeight(x)
		case string:
			z = GetEnumWeightNum(x)
		default:
			opt = nil
		}
		opt = append(opt, z)
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
			_core.values["Weight"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return result
}
