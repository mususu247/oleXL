package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

// version 2026-01-05
// VBA style like

type listObjects struct {
	app *Excel
	mx  int
}

type listObject struct {
	app *Excel
	mx  int
}

func (lo *listObject) Nothing() error {
	xl := lo.app
	_, err := xl.getCore(lo.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	// .child.RelaseAll
	for k, v := range xl.WorkCores.cores {
		if v.px == lo.mx {
			xl.Release(k)
		}
	}
	return nil
}

func (ws *workSheet) ListObjects() *listObjects {
	var los listObjects
	los.app = ws.app
	xl := ws.app

	los.mx, _ = xl.findCore(ws.mx, "ListObjects", 0)
	if los.mx >= 0 {
		return &los
	}

	_core, _ := xl.getCore(ws.mx)

	const cmd = "Get"
	const name = "ListObjects"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			los.mx, _ = xl.addCore(ws.mx, x, name, 0)
			los.Count()
			return &los
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (los *listObjects) Count(lock ...bool) int32 {
	xl := los.app
	_core, err := xl.getCore(los.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	const cmd = "Get"
	const name = "Count"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case int32:
			log.Printf("%v ans (int32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return -1
}

func (los *listObjects) List() []*Core {
	var loz []*Core

	xl := los.app
	_core, err := xl.getCore(los.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}

	// delete.index = -1
	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == los.mx {
			xl.WorkCores.cores[i].index = -1
		}
	}

	count := los.Count()

	const cmd = "Method"
	const name = "Item"
	var opt []any
	opt = append(opt, int32(0))

	for j := int32(1); j <= count; j++ {
		opt[0] = j
		args := xl.worker.Send(cmd, _core.disp, name, opt)
		var lo listObject
		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case *ole.IDispatch:
				log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

				var _lo *Core
				lo.app = los.app
				lo.mx, _lo = xl.addCore(los.mx, x, "ListObject", j)
				lo.Name()
				loz = append(loz, _lo)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	// delete.index = -1
	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == los.mx {
			if xl.WorkCores.cores[i].index == -1 {
				delete(xl.WorkCores.cores, i)
			}
		}
	}

	return loz
}

func (lo *listObject) Name(value ...string) string {
	xl := lo.app
	_core, err := xl.getCore(lo.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return ""
	}
	var cmd string
	const name = "Name"
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
		case string:
			log.Printf("%v ans (string) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return ""
}
