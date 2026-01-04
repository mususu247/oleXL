package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

// version 2026-01-04
// VBA style like

type workSheets struct {
	app *Excel
	mx  int
}

type workSheet struct {
	app *Excel
	mx  int
}

func (ws *workSheet) Nothing() error {
	xl := ws.app
	_, err := xl.getCore(ws.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	// wb.child.RelaseAll
	for k, v := range xl.WorkCores.cores {
		if v.px == ws.mx {
			xl.Release(k)
		}
	}
	return nil
}

func (wb *workBook) Worksheets() *workSheets {
	var wss workSheets
	wss.app = wb.app
	xl := wb.app

	wss.mx, _ = xl.findCore(wb.mx, "Worksheets", 0)
	if wss.mx >= 0 {
		return &wss
	}

	_core, _ := xl.getCore(xl.mx)

	const cmd = "Get"
	const name = "Worksheets"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			wss.mx, _ = xl.addCore(wb.mx, x, name, 0)
			wss.Count()
			return &wss
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (wb *workBook) Worksheetz(value any) *workSheet {
	var ws workSheet
	xl := wb.app
	wss := wb.Worksheets()
	wss.List()
	wsz := xl.findChild(wss.mx, "Worksheet")

	switch x := value.(type) {
	case int:
		for i := range wsz {
			j := wsz[i]
			if xl.WorkCores.cores[j].index == int32(x) {
				ws.app = wss.app
				ws.mx = j
				return &ws
			}
		}
	case int32:
		for i := range wsz {
			j := wsz[i]
			if xl.WorkCores.cores[j].index == int32(x) {
				ws.app = wss.app
				ws.mx = j
				return &ws
			}
		}
	case string:
		for i := range wsz {
			j := wsz[i]
			if v, ok := xl.WorkCores.cores[j].values["Name"]; ok {
				if v.(string) == x {
					ws.app = wss.app
					ws.mx = j
					return &ws
				}
			}
		}
	}
	return nil
}

func (wss *workSheets) Count(lock ...bool) int32 {
	xl := wss.app
	_core, err := xl.getCore(wss.mx)
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

func (wss *workSheets) List() []*Core {
	var wsz []*Core

	xl := wss.app
	_wss, err := xl.getCore(wss.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	_core, err := xl.getCore(_wss.px)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}

	// delete.index = -1
	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == wss.mx {
			xl.WorkCores.cores[i].index = -1
		}
	}

	count := wss.Count()

	const cmd = "Get"
	const name = "Worksheets"
	var opt []any
	opt = append(opt, int32(0))

	for j := int32(1); j <= count; j++ {
		opt[0] = j
		args := xl.worker.Send(cmd, _core.disp, name, opt)
		var ws workSheet
		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case *ole.IDispatch:
				log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

				var _ws *Core
				ws.app = wss.app
				ws.mx, _ws = xl.addCore(wss.mx, x, "Worksheet", j)
				ws.Name()
				wsz = append(wsz, _ws)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	// delete.index = -1
	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == wss.mx {
			if xl.WorkCores.cores[i].index == -1 {
				delete(xl.WorkCores.cores, i)
			}
		}
	}

	return wsz
}

func (ws *workSheet) Name(value ...string) string {
	xl := ws.app
	_core, err := xl.getCore(ws.mx)
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

func (ws *workSheet) Activate() error {
	xl := ws.app
	_core, err := xl.getCore(ws.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	const cmd = "Method"
	const name = "Activate"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			return x
		case bool:
			log.Printf("%v abs (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			if x {
				_xl, _ := xl.getCore(xl.mx)
				_xl.values["ActiveSheet"] = ws.mx
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}

func (ws *workSheet) Select(value ...bool) error {
	xl := ws.app
	_core, err := xl.getCore(ws.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	const cmd = "Method"
	const name = "Select"
	var opt []any
	if len(value) > 0 {
		opt = append(opt, value[0])
	} else {
		opt = append(opt, true)
	}

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			return x
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			if x {
				xl.setCoreValue(xl.mx, "Selection", ws.mx)
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return nil
}

func (wss *workSheets) Add(value ...any) *workSheet {
	var ws workSheet
	ws.app = wss.app
	xl := wss.app
	_core, err := xl.getCore(wss.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	const cmd = "Method"
	const name = "Add"
	var opt []any

	if len(value) > 0 {
		for i := range value {
			switch x := value[i].(type) {
			case string:
				cz := wss.List()
				for j := range cz {
					if cz[j].values["Name"] == x {
						opt = append(opt, cz[j].disp)
					}
				}
			case int32:
				cz := wss.List()
				for j := range cz {
					if cz[j].index == x {
						opt = append(opt, cz[j].disp)
					}
				}
			case nil:
				opt = append(opt, nil)
			}
		}
	} else {
		opt = nil
	}

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

			ws.mx, _ = xl.addCore(wss.mx, x, "Worksheet", 0)
			ws.Name()
			wss.List()
			return &ws
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return nil
}

func (ws *workSheet) Move(value ...any) *workSheet {
	xl := ws.app
	_core, err := xl.getCore(ws.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	var wss workSheets
	wss.app = ws.app
	wss.mx = _core.px

	const cmd = "Method"
	const name = "Move"
	var opt []any

	if len(value) > 0 {
		for i := range value {
			switch x := value[i].(type) {
			case string:
				cz := wss.List()
				for j := range cz {
					if cz[j].values["Name"] == x {
						opt = append(opt, cz[j].disp)
					}
				}
			case int32:
				cz := wss.List()
				for j := range cz {
					if cz[j].index == x {
						opt = append(opt, cz[j].disp)
					}
				}
			case nil:
				opt = append(opt, nil)
				_xl, _ := xl.getCore(xl.mx)
				delete(_xl.values, "ActiveWorkbook")
			}
		}
	} else {
		opt = nil
	}

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			if x {
				sheets := xl.ActiveWorkbook().Worksheets()
				sheet := xl.ActiveSheet()
				_sheet, err := xl.getCore(sheet.mx)
				if err == nil {
					_sheet.px = sheets.mx
				}

				if sheet.mx != ws.mx {
					childs := xl.findChild(ws.mx, "")
					for j := range childs {
						k := childs[j]
						xl.WorkCores.cores[k].px = sheet.mx
					}
					xl.Release(ws.mx)
					wss.List()

					ws.mx = sheet.mx
					return ws
				}
				sheets.List()
				return ws
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return nil
}

func (ws *workSheet) Copy(value ...any) *workSheet {
	xl := ws.app
	_core, err := xl.getCore(ws.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	var wss workSheets
	wss.app = ws.app
	wss.mx = _core.px

	const cmd = "Method"
	const name = "Copy"
	var opt []any

	if len(value) > 0 {
		for i := range value {
			switch x := value[i].(type) {
			case string:
				cz := wss.List()
				for j := range cz {
					if cz[j].values["Name"] == x {
						opt = append(opt, cz[j].disp)
					}
				}
			case int32:
				cz := wss.List()
				for j := range cz {
					if cz[j].index == x {
						opt = append(opt, cz[j].disp)
					}
				}
			case nil:
				opt = append(opt, nil)
				_xl, _ := xl.getCore(xl.mx)
				delete(_xl.values, "ActiveWorkbook")
			}
		}
	} else {
		opt = nil
	}

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			if x {
				sheets := xl.ActiveWorkbook().Worksheets()
				sheet := xl.ActiveSheet()
				_sheet, err := xl.getCore(sheet.mx)
				if err == nil {
					_sheet.px = sheets.mx
				}

				if sheet.mx != ws.mx {
					childs := xl.findChild(ws.mx, "")
					for j := range childs {
						k := childs[j]
						xl.WorkCores.cores[k].px = sheet.mx
					}
					xl.Release(ws.mx)
					wss.List()

					ws.mx = sheet.mx
					return ws
				}
				sheets.List()
				return ws
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return nil
}

func (ws *workSheet) Delete() error {
	xl := ws.app
	_core, err := xl.getCore(ws.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	var wss workSheets
	wss.app = ws.app
	wss.mx = _core.px

	const cmd = "Method"
	const name = "Delete"

	sw := xl.Application().DisplayAlerts()
	xl.Application().DisplayAlerts(false)

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			return x
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			if x {
				xl.Release(ws.mx)
				wss.List()
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	xl.Application().DisplayAlerts(sw)
	return nil
}
