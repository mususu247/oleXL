package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

// version 2026-01-05
// VBA style like

type workNames struct {
	app *Excel
	mx  int
}

type workName struct {
	app *Excel
	mx  int
}

func (nm *workName) Nothing() error {
	xl := nm.app
	_, err := xl.getCore(nm.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	// .child.RelaseAll
	for k, v := range xl.WorkCores.cores {
		if v.px == nm.mx {
			xl.Release(k)
		}
	}
	return nil
}

func (wb *workBook) Names() *workNames {
	var nms workNames
	nms.app = wb.app
	xl := wb.app

	nms.mx, _ = xl.findCore(wb.mx, "Names", 0)
	if nms.mx >= 0 {
		return &nms
	}

	_core, _ := xl.getCore(wb.mx)

	const cmd = "Get"
	const name = "Names"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			nms.mx, _ = xl.addCore(wb.mx, x, name, 0)
			nms.Count()
			return &nms
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (wb *workBook) Namez(value any) *workName {
	var nm workName
	xl := wb.app
	nms := wb.Names()
	nms.List()
	wsz := xl.findChild(nms.mx, "Name")

	switch x := value.(type) {
	case int:
		for i := range wsz {
			j := wsz[i]
			if xl.WorkCores.cores[j].index == int32(x) {
				nm.app = nms.app
				nm.mx = j
				return &nm
			}
		}
	case int32:
		for i := range wsz {
			j := wsz[i]
			if xl.WorkCores.cores[j].index == int32(x) {
				nm.app = nms.app
				nm.mx = j
				return &nm
			}
		}
	case string:
		for i := range wsz {
			j := wsz[i]
			if v, ok := xl.WorkCores.cores[j].values["Name"]; ok {
				if v.(string) == x {
					nm.app = nms.app
					nm.mx = j
					return &nm
				}
			}
		}
	}
	return nil
}

func (nms *workNames) Count(lock ...bool) int32 {
	xl := nms.app
	_core, err := xl.getCore(nms.mx)
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

func (nm *workName) Name() string {
	xl := nm.app
	_core, err := xl.getCore(nm.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return ""
	}
	const cmd = "Get"
	const name = "Name"

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

func (nms *workNames) List() []*Core {
	var nmz []*Core

	xl := nms.app
	_core, err := xl.getCore(nms.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}

	// delete.index = -1
	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == nms.mx {
			xl.WorkCores.cores[i].index = -1
		}
	}

	count := nms.Count()

	const cmd = "Method"
	const name = "Item"
	var opt []any
	opt = append(opt, int32(0))

	for j := int32(1); j <= count; j++ {
		opt[0] = j
		args := xl.worker.Send(cmd, _core.disp, name, opt)
		var nm workName
		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case *ole.IDispatch:
				log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

				var _nm *Core
				nm.app = nms.app
				nm.mx, _nm = xl.addCore(nms.mx, x, "Name", j)
				nm.Name()
				nmz = append(nmz, _nm)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	// delete.index = -1
	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == nms.mx {
			if xl.WorkCores.cores[i].index == -1 {
				delete(xl.WorkCores.cores, i)
			}
		}
	}

	return nmz
}

func (nms *workNames) Add(Name string, RefersTo string) *workName {
	var nm workName
	nm.app = nms.app
	xl := nms.app
	_core, err := xl.getCore(nms.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	const cmd = "Method"
	const name = "Add"
	var opt []any

	opt = append(opt, Name)
	opt = append(opt, RefersTo)

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

			nm.mx, _ = xl.addCore(nms.mx, x, "Name", 0)
			nms.List()
			return &nm
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return nil
}

func (nm *workName) Delete() error {
	xl := nm.app

	_core, err := xl.getCore(nm.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}

	var nms workNames
	nms.app = nm.app
	nms.mx = _core.px

	const cmd = "Method"
	const name = "Delete"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			return x
		case nil:
			log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			xl.Release(nm.mx)
			nms.List()
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}
