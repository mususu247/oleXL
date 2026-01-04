package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

// version 2026-01-04
// VBA style like

type workCores struct {
	app   *Excel
	last  int
	cores map[int]*Core
}

type Core struct {
	px    int
	disp  *ole.IDispatch
	kind  string
	index int32

	values map[string]any
}

func (xl *Excel) addCore(px int, disp *ole.IDispatch, kind string, index int32) (int, *Core) {
	var wc Core

	mx, _core := xl.findDisp(disp)
	if mx >= 0 {
		_core.index = index
		return mx, _core
	}

	wc.px = px
	wc.disp = disp
	wc.kind = kind
	wc.index = index
	wc.values = make(map[string]any)

	xl.WorkCores.last++
	last := xl.WorkCores.last
	xl.WorkCores.cores[last] = &wc
	return last, &wc
}

func (xl *Excel) getCore(mx int) (*Core, error) {
	if _, ok := xl.WorkCores.cores[mx]; ok {
		return xl.WorkCores.cores[mx], nil
	}
	return nil, fmt.Errorf("not found: Cores[%v]", mx)
}

func (xl *Excel) setCoreValue(mx int, name string, value any) error {
	if _, ok := xl.WorkCores.cores[mx]; ok {
		xl.WorkCores.cores[mx].values[name] = value
		return nil
	} else {
		return fmt.Errorf("not found: Cores[%v]", mx)
	}
}

func (xl *Excel) findCore(px int, kind string, index int32) (int, *Core) {
	for mx, v := range xl.WorkCores.cores {
		if v.px == px {
			if (v.kind == kind) && (v.index == index) {
				return mx, xl.WorkCores.cores[mx]
			}
		}
	}
	return -1, nil
}

func (xl *Excel) findDisp(disp *ole.IDispatch) (int, *Core) {
	for mx, v := range xl.WorkCores.cores {
		if v.disp == disp {
			return mx, xl.WorkCores.cores[mx]
		}

	}
	return -1, nil
}

func (xl *Excel) findChild(px int, kind string) []int {
	var results []int

	if kind == "" {
		// kind=All
		for mx, v := range xl.WorkCores.cores {
			if v.px == px {
				results = append(results, mx)
			}
		}
	} else {
		for mx, v := range xl.WorkCores.cores {
			if (v.px == px) && (v.kind == kind) {
				results = append(results, mx)
			}
		}
	}
	return results
}

func (xl *Excel) Release(mx int) error {
	// Childs.Release
	for k, v := range xl.WorkCores.cores {
		if v.px == mx {
			xl.Release(k)
		}
	}

	_core, err := xl.getCore(mx)
	if err != nil {
		return err
	}

	const cmd = "Release"
	var name = _core.kind

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, name, x, nil, _core.values)
			return x
		case int32:
			log.Printf("%v ans (int32) %v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, name, x, nil, _core.values)
		default:
			log.Printf("%v def %v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, name, x, nil, _core.values)
		}
	}

	//log.Printf("Release.kind: %v", _core.kind)
	_core.px = -1
	_core.disp = nil
	_core.kind = ""
	_core.values = nil
	_core.index = -1
	delete(xl.WorkCores.cores, mx)
	return nil
}

func (xl *Excel) ReleaseAll() error {
	for mx := range xl.WorkCores.cores {
		xl.Release(mx)
	}
	return nil
}
