package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

// version 2026-01-04
// VBA style like

type workCharts struct {
	app *Excel
	mx  int
}

type workChart struct {
	app *Excel
	mx  int
}

type workText struct {
	app *Excel
	mx  int
}

func (ws *workSheet) ChartObjects() *workCharts {
	var cos workCharts
	cos.app = ws.app
	xl := ws.app

	cos.mx, _ = xl.findCore(ws.mx, "ChartObjects", 0)
	if cos.mx >= 0 {
		return &cos
	}

	_core, _ := xl.getCore(ws.mx)

	const cmd = "Method"
	const name = "ChartObjects"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			cos.mx, _ = xl.addCore(ws.mx, x, name, 0)
			cos.Count()
			return &cos
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (ws *workSheet) ChartObjectz(value any) *workChart {
	var co workChart
	xl := ws.app
	cos := ws.ChartObjects()
	cos.List()
	wsz := xl.findChild(cos.mx, "ChartObjects")

	switch x := value.(type) {
	case int:
		for i := range wsz {
			j := wsz[i]
			if xl.WorkCores.cores[j].index == int32(x) {
				co.app = cos.app
				co.mx = j
				return &co
			}
		}
	case int32:
		for i := range wsz {
			j := wsz[i]
			if xl.WorkCores.cores[j].index == int32(x) {
				co.app = cos.app
				co.mx = j
				return &co
			}
		}
	case string:
		for i := range wsz {
			j := wsz[i]
			if v, ok := xl.WorkCores.cores[j].values["Name"]; ok {
				if v.(string) == x {
					co.app = cos.app
					co.mx = j
					return &co
				}
			}
		}
	}
	return nil
}

func (cos *workCharts) Count(lock ...bool) int32 {
	xl := cos.app
	_core, err := xl.getCore(cos.mx)
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

func (cos *workCharts) List() []*Core {
	var coz []*Core

	xl := cos.app
	_core, err := xl.getCore(cos.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}

	// delete.index = -1
	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == cos.mx {
			xl.WorkCores.cores[i].index = -1
		}
	}

	count := cos.Count()

	const cmd = "Method"
	const name = "Item"
	var opt []any
	opt = append(opt, int32(0))

	for j := int32(1); j <= count; j++ {
		opt[0] = j
		args := xl.worker.Send(cmd, _core.disp, name, opt)
		var co workChart
		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case *ole.IDispatch:
				log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

				var _co *Core
				co.app = cos.app
				co.mx, _co = xl.addCore(cos.mx, x, "Chart", j)
				co.Name()
				co.Index()
				coz = append(coz, _co)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	// delete.index = -1
	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == cos.mx {
			if xl.WorkCores.cores[i].index == -1 {
				delete(xl.WorkCores.cores, i)
			}
		}
	}

	return coz
}

func (co *workChart) Name(value ...string) string {
	xl := co.app
	_core, err := xl.getCore(co.mx)
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

func (co *workChart) Index() int32 {
	xl := co.app
	_core, err := xl.getCore(co.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	const cmd = "Get"
	const name = "Index"

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

func (co *workChart) HasTitle(value ...bool) bool {
	xl := co.app
	_core, err := xl.getCore(co.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return false
	}
	var cmd string
	const name = "HasTitle"

	if len(value) > 0 {
		var opt []any
		opt = append(opt, value[0])

		cmd = "Put"
		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			}
		}
	}

	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return false
}

func (co *workChart) ChartTitle() *workText {
	var tx workText
	tx.app = co.app
	xl := co.app
	_core, err := xl.getCore(co.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	const cmd = "Get"
	const name = "ChartTitle"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			tx.mx, _ = xl.addCore(co.mx, x, "ChartTitle", 0)
			return &tx
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}

func (tx *workText) Text(value ...string) string {
	xl := tx.app
	_core, err := xl.getCore(tx.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return ""
	}
	var cmd string
	const name = "Text"

	if len(value) > 0 {
		var opt []any
		opt = append(opt, value[0])
		cmd = "Put"
		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
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
