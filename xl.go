package oleXL

import (
	"fmt"
	"log"
	"strings"
	"time"

	"github.com/go-ole/go-ole"
)

// version 2026-01-05
// VBA style like

//type any = interface{}

type Excel struct {
	worker    Worker
	WorkCores workCores
	mx        int
	hWnd      int32
}

func (xl *Excel) Init() error {
	xl.WorkCores.app = xl
	xl.WorkCores.cores = make(map[int]*Core)

	err := xl.worker.Start()
	if err != nil {
		return err
	}
	return nil
}

func (xl *Excel) Nothing() error {
	xl.Release(xl.mx)

	err := xl.worker.Stop()
	if err != nil {
		log.Printf("%v", err)
		return err
	}

	for {
		if xl.worker.IsOpened() {
			time.Sleep(1 * time.Millisecond)
		} else {
			log.Printf("worker.IsOpened: false")
			time.Sleep(1 * time.Millisecond)
			return nil
		}
	}

	//return nil
}

func (xl *Excel) CreateObject() error {
	const cmd = "Create"
	const name = "Excel.Application"

	args := xl.worker.Send(cmd, nil, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, nil, name, x, nil, nil)
			xl.Nothing()
			return x
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, "", name, x, nil, nil)
			xl.mx, _ = xl.addCore(-1, x, name, 0)
			xl.hand()
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, "", name, x, nil, nil)
		}
	}
	return nil
}

func (xl *Excel) hand() int32 {
	var result int32
	_core, err := xl.getCore(xl.mx)
	if err != nil {
		return -1
	}
	const cmd = "Get"
	const name = "hWnd"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case int32:
			log.Printf("%v ans (int32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			result = x
			xl.hWnd = x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return result
}

func (xl *Excel) Visible(value bool) error {
	_core, err := xl.getCore(xl.mx)
	if err != nil {
		return err
	}
	const cmd = "Put"
	const name = "Visible"
	var opt []any
	opt = append(opt, value)

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			return x
		case nil:
			log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return nil
}

func (xl *Excel) Quit() error {
	_core, err := xl.getCore(xl.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	const cmd = "Method"
	const name = "Quit"

	// all Book.Close
	var wb workBook
	for mx, v := range xl.WorkCores.cores {
		if v.kind == "Workbook" {
			wb.app = xl
			wb.mx = mx
			wb.Close(false)
		}
	}

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			return x
		case nil:
			log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			xl.ReleaseAll()
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}

func (xl *Excel) subList(mx int, level int) {
	lv := level + 1
	h := strings.Repeat(" ", lv)

	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == mx {
			core := xl.WorkCores.cores[i]
			fmt.Printf("%v kind:%v index:%v values:%v mx:%v\n", h, core.kind, core.index, core.values, i)
			xl.subList(i, lv)
		}
	}
}

func (xl *Excel) List() {
	var level int = 1
	h := strings.Repeat(" ", level)

	fmt.Println("[Excel]")
	if _, ok := xl.WorkCores.cores[xl.mx]; ok {
		core := xl.WorkCores.cores[xl.mx]
		fmt.Printf("%v kind:%v index:%v values:%v mx:%v\n", h, core.kind, core.index, core.values, xl.mx)
		xl.subList(xl.mx, level)
	}
}

func (xl *Excel) ActiveWorkbook(value ...bool) *workBook {
	var wb workBook
	wb.app = xl
	_core, _ := xl.getCore(xl.mx)

	if mz, ok := _core.values["ActiveWorkbook"]; ok {
		switch z := mz.(type) {
		case int:
			if _, ok := xl.WorkCores.cores[z]; ok {
				wb.mx = z
				return &wb
			} else {
				delete(_core.values, "ActiveWorkbook")
			}
		}
	}

	const cmd = "Get"
	const name = "ActiveWorkbook"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			for j := range xl.WorkCores.cores {
				if xl.WorkCores.cores[j].kind == "Workbook" {
					if xl.WorkCores.cores[j].disp == x {
						wb.mx = j
						return &wb
					}
				}
			}

			var wbs workBooks
			wbs.app = xl
			wbs.mx = xl.Workbooks().mx

			wb.mx, _ = xl.addCore(wbs.mx, x, "Workbook", 0)
			wb.Name()
			wbs.List()
			xl.setCoreValue(xl.mx, "ActiveWorkbook", wb.mx)
			return &wb
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return &wb
}

func (xl *Excel) ActiveSheet(value ...bool) *workSheet {
	var ws workSheet
	ws.app = xl
	_core, _ := xl.getCore(xl.mx)

	if len(value) == 0 {
		if mz, ok := _core.values["ActiveSheet"]; ok {
			switch z := mz.(type) {
			case int:
				if _, ok := xl.WorkCores.cores[z]; ok {
					ws.mx = z
					return &ws
				} else {
					delete(_core.values, "ActiveSheet")
				}
			}
		}
	}

	const cmd = "Get"
	const name = "ActiveSheet"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			ws.mx, _ = xl.findDisp(x)
			if ws.mx >= 0 {
				ws.app = xl
				_core.values["ActiveSheet"] = ws.mx
				return &ws
			} else {
				wss := xl.ActiveWorkbook().Worksheets()
				ws.mx, _ = xl.addCore(wss.mx, x, "Worksheet", 0)
				wss.List()
				xl.setCoreValue(xl.mx, "ActiveSheet", ws.mx)
				return &ws
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return &ws
}

func (xl *Excel) ActiveCell(value ...bool) *workRag {
	var rg workRag
	rg.app = xl
	_core, _ := xl.getCore(xl.mx)

	if len(value) == 0 {
		if mz, ok := _core.values["ActiveCell"]; ok {
			switch z := mz.(type) {
			case int:
				if _, ok := xl.WorkCores.cores[z]; ok {
					rg.mx = z
					return &rg
				} else {
					delete(_core.values, "ActiveCell")
				}
			}
		}
	}

	const cmd = "Get"
	const name = "ActiveCell"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			rg.mx, _ = xl.findDisp(x)
			if rg.mx >= 0 {
				//rg.app = xl
				_core.values["ActiveCell"] = rg.mx
				return &rg
			} else {
				ws := xl.ActiveSheet()
				rg.mx, _ = xl.addCore(ws.mx, x, "Range", 0)
				xl.setCoreValue(xl.mx, "ActiveCell", rg.mx)
				return &rg
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return &rg
}

func (xl *Excel) ActiveChart(value ...bool) *workChart {
	var co workChart
	co.app = xl
	_core, _ := xl.getCore(xl.mx)

	if len(value) > 0 {
		if mz, ok := _core.values["ActiveChart"]; ok {
			switch z := mz.(type) {
			case int:
				if _, ok := xl.WorkCores.cores[z]; ok {
					co.mx = z
					return &co
				} else {
					delete(_core.values, "ActiveChart")
				}
			}
		}
	}

	const cmd = "Get"
	const name = "ActiveChart"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case nil:
			log.Printf("%v ans %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			co.mx, _ = xl.findDisp(x)
			if co.mx >= 0 {
				//co.app = xl
				_core.values["ActiveChart"] = co.mx
				return &co
			} else {
				cos := xl.ActiveSheet().ChartObjects()
				co.mx, _ = xl.addCore(cos.mx, x, "ChartObject", 0)
				xl.setCoreValue(xl.mx, "ActiveChart", co.mx)
				return &co
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return &co
}

func (xl *Excel) Selection(value ...bool) *workRag {
	var rg workRag
	rg.app = xl
	_core, _ := xl.getCore(xl.mx)

	if len(value) == 0 {
		if mz, ok := _core.values["Selection"]; ok {
			switch z := mz.(type) {
			case int:
				if _, ok := xl.WorkCores.cores[z]; ok {
					rg.mx = z
					return &rg
				} else {
					delete(_core.values, "Selection")
				}
			}
		}
	}

	const cmd = "Get"
	const name = "Selection"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			rg.mx, _ = xl.findDisp(x)
			if rg.mx >= 0 {
				rg.app = xl
				_core.values["Selection"] = rg.mx
				return &rg
			} else {
				ws := xl.ActiveSheet()
				rg.mx, _ = xl.addCore(ws.mx, x, "Range", 0)
				xl.setCoreValue(xl.mx, "Selection", rg.mx)
				return &rg
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return &rg
}
