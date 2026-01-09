package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

// version 2026-01-09
// VBA style like

type workBooks struct {
	app *Excel
	mx  int
}

type workBook struct {
	app *Excel
	mx  int
}

func (wb *workBook) Nothing() error {
	xl := wb.app
	_, err := xl.getCore(wb.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	// wb.child.RelaseAll
	for k, v := range xl.WorkCores.cores {
		if v.px == wb.mx {
			xl.Release(k)
		}
	}
	return nil
}

func (xl *Excel) Workbooks() *workBooks {
	var wbs workBooks
	wbs.app = xl

	wbs.mx, _ = xl.findCore(xl.mx, "Workbooks", 0)
	if wbs.mx >= 0 {
		return &wbs
	}

	_core, _ := xl.getCore(xl.mx)

	const cmd = "Get"
	const name = "Workbooks"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			wbs.mx, _ = xl.addCore(xl.mx, x, name, 0)
			wbs.Count()
			return &wbs
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (xl *Excel) Workbookz(value any) *workBook {
	var wb workBook

	var wbs workBooks
	wbs.app = xl
	wbs.mx, _ = xl.findCore(xl.mx, "Workbooks", 0)
	if wbs.mx >= 0 {
		log.Printf("(Error) not found: %v\n", "Workbooks")
		return nil
	}
	wbs.List()
	wbz := xl.findChild(xl.mx, "Workbook")

	switch x := value.(type) {
	case int:
		for i := range wbz {
			j := wbz[i]
			if xl.WorkCores.cores[j].index == int32(x) {
				wb.app = xl
				wb.mx = j
				return &wb
			}
		}
	case int32:
		for i := range wbz {
			j := wbz[i]
			if xl.WorkCores.cores[j].index == int32(x) {
				wb.app = xl
				wb.mx = j
				return &wb
			}
		}
	case string:
		for i := range wbz {
			j := wbz[i]
			if v, ok := xl.WorkCores.cores[j].values["Name"]; ok {
				if v.(string) == x {
					wb.app = xl
					wb.mx = j
					return &wb
				}
			}
		}
	}
	return nil
}

func (wbs *workBooks) Count(lock ...bool) int32 {
	xl := wbs.app
	_core, err := xl.getCore(wbs.mx)
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

func (wbs *workBooks) List() []*Core {
	var wbz []*Core

	xl := wbs.app
	_wbs, err := xl.getCore(wbs.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	_core, err := xl.getCore(_wbs.px)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}

	// delete.index = -1
	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == wbs.mx {
			xl.WorkCores.cores[i].index = -1
		}
	}

	count := wbs.Count()

	const cmd = "Get"
	const name = "Workbooks"
	var opt []any
	opt = append(opt, int32(0))

	for j := int32(1); j <= count; j++ {
		opt[0] = j
		args := xl.worker.Send(cmd, _core.disp, name, opt)
		var wb workBook
		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case *ole.IDispatch:
				log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

				var _wb *Core
				wb.app = wbs.app
				wb.mx, _wb = xl.addCore(wbs.mx, x, "Workbook", j)
				wb.Name()
				wb.Worksheets().List()
				wbz = append(wbz, _wb)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	// delete.index = -1
	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == wbs.mx {
			if xl.WorkCores.cores[i].index == -1 {
				delete(xl.WorkCores.cores, i)
			}
		}
	}

	return wbz
}

func (wb *workBook) Name() string {
	xl := wb.app
	_core, err := xl.getCore(wb.mx)
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

func (wb *workBook) Activate() error {
	xl := wb.app
	_core, err := xl.getCore(wb.mx)
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
				xl.setCoreValue(xl.mx, "ActiveWorkbook", wb.mx)
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}

func (wbs *workBooks) Add() *workBook {
	var wb workBook
	wb.app = wbs.app
	xl := wbs.app
	_core, err := xl.getCore(wbs.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	const cmd = "Method"
	const name = "Add"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			wb.mx, _ = xl.addCore(wbs.mx, x, "Workbook", 0)
			wb.Name()
			wbs.List()

			_xl, _ := xl.getCore(xl.mx)
			_xl.values["ActiveWorkbook"] = wb.mx
			return &wb
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}

func (wb *workBook) Close(value ...bool) error {
	xl := wb.app
	_core, err := xl.getCore(wb.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}

	const cmd = "Method"
	const name = "Close"
	var opt []any

	if len(value) > 0 {
		opt = append(opt, value[0])
	} else {
		opt = append(opt, false)
	}

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			if x {
				// wb.child.RelaseAll
				for k, v := range xl.WorkCores.cores {
					if v.px == wb.mx {
						xl.Release(k)
					}
				}
			} else {
				return nil
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}

	// Release Book.Worksheets...
	var wbs workBooks
	wbs.app = wb.app
	wbs.mx = _core.px
	xl.Release(wb.mx)
	wbs.List()
	return nil
}

func (wb *workBook) RefreshAll() error {
	xl := wb.app
	_core, err := xl.getCore(wb.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}

	const cmd = "Method"
	const name = "RefreshAll"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case nil:
			log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (wbs *workBooks) Open(fileName string) (*workBook, error) {
	var wb workBook
	wb.app = wbs.app
	xl := wbs.app
	_core, err := xl.getCore(wbs.mx)
	if err != nil {
		return nil, fmt.Errorf("(Error) %v", err)
	}
	const cmd = "Method"
	const name = "Open"
	var opt []any

	fn, err := GetAbsolutePathName(fileName)
	if err != nil {
		log.Printf("(Error) %v GetAbsolutePathName:%v", err, fn)
	}

	if FileExists(fn) {
		opt = append(opt, fn)
	}

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			wb.mx, _ = xl.addCore(wbs.mx, x, "Workbook", 0)
			wb.Name()
			wbs.List()

			_xl, _ := xl.getCore(xl.mx)
			_xl.values["ActiveWorkbook"] = wb.mx
			return &wb, nil
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil, nil
}

func (wb *workBook) Save() error {
	xl := wb.app
	_core, err := xl.getCore(wb.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}

	const cmd = "Method"
	const name = "Save"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (wb *workBook) SaveAs(fileName string, option ...any) error {
	xl := wb.app
	_core, err := xl.getCore(wb.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}

	const cmd = "Method"
	const name = "SaveAs"
	var opt []any
	var z int32

	fn, err := GetAbsolutePathName(fileName)
	if err != nil {
		log.Printf("(Error) %v GetAbsolutePathName:%v", err, fn)
	}

	if FileExists(fn) {
		DeleteFile(fn)
	}
	opt = append(opt, fn)

	if len(option) > 0 {
		switch x := option[0].(type) {
		case int:
			z = SetEnumFileFormat(int32(x))
		case int32:
			z = SetEnumFileFormat(x)
		case string:
			z = GetEnumFileFormatNum(x)
		}
	} else {
		z = GetEnumFileFormatNum("Default")
	}
	opt = append(opt, z)

	sw := xl.Application().DisplayAlerts()
	if sw {
		xl.Application().DisplayAlerts(false)
	}

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	xl.Application().DisplayAlerts(sw)
	return nil
}

func (wb *workBook) SaveCopyAs(fileName string) error {
	xl := wb.app
	_core, err := xl.getCore(wb.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}

	const cmd = "Method"
	const name = "SaveCopyAs"
	var opt []any

	fn, _ := GetAbsolutePathName(fileName)
	if FileExists(fn) {
		DeleteFile(fn)
	}
	opt = append(opt, fn)

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}

	return nil
}

func (wb *workBook) ReadOnly() bool {
	xl := wb.app
	_core, err := xl.getCore(wb.mx)
	if err != nil {
		log.Panicf("(Error) %v", err)
		return true
	}
	const cmd = "Get"
	const name = "ReadOnly"

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
	return true
}
