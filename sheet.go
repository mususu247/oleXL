package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workSheets struct {
	app    *Excel
	parent *workBook
	num    int
}

type workSheet struct {
	app    *Excel
	parent *workBook
	num    int
}

func (wb *workBook) Worksheets() *workSheets {
	var wss workSheets
	xl := wb.app

	name := "Worksheets"
	core, num := xl.cores.FindAdd(name, wb.num)
	if core.disp == nil {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wb.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 1 //Lock.on
		}
	}

	wss.app = xl
	wss.num = num
	wss.parent = wb
	return &wss
}

func (wb *workBook) Worksheetz(value any) *workSheet {
	var ws workSheet
	xl := wb.app

	kind := "Worksheet"
	core, num := xl.cores.FindAdd(kind, wb.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Worksheets"
		var opt []any
		switch x := value.(type) {
		case int:
			if x > 0 {
				opt = append(opt, int32(x))
			}
		case int32:
			if x > 0 {
				opt = append(opt, x)
			}
		case string:
			opt = append(opt, x)
		}

		ans, err := xl.cores.SendNum(cmd, name, wb.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	ws.app = xl
	ws.num = num
	ws.parent = wb
	return &ws
}

func (xl *Excel) ActiveSheet() *workSheet {
	var ws workSheet
	wb := xl.ActiveWorkbook()

	kind := "Worksheet"
	core, num := xl.cores.FindAdd(kind, wb.num)
	if core.disp == nil {
		cmd := "Get"
		name := "ActiveSheet"
		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	ws.app = xl
	ws.num = num
	ws.parent = wb
	wb.Release()
	return &ws
}

func (wss *workSheets) Release() error {
	xl := wss.app
	return xl.cores.Release(wss.num, false)
}

func (wss *workSheets) Nothing() error {
	xl := wss.app
	xl.cores.releaseChild(wss.num)

	xl.cores.Unlock(wss.num)
	err := wss.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(wss.num)
	wss = nil
	return nil
}

func (wss *workSheets) Count() int32 {
	var result int32
	xl := wss.app

	cmd := "Get"
	name := "Count"
	ans, err := xl.cores.SendNum(cmd, name, wss.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return result
	}

	switch x := ans.(type) {
	case int32:
		result = x
	}
	return result
}

func (wss *workSheets) Add(value ...any) *workSheet {
	var ws workSheet
	xl := wss.app
	wb := xl.ActiveWorkbook()

	kind := "Worksheet"
	core, num := xl.cores.FindAdd(kind, wb.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Add"

		var opt []any
		if len(value) > 0 {
			for i := range value {
				switch x := value[i].(type) {
				case int, int32, string:
					ws := wb.Worksheetz(x)
					core := xl.cores.getCore(ws.num)
					opt = append(opt, core.disp)
				case nil:
					opt = append(opt, nil)
				}
			}
		} else {
			opt = append(opt, nil)
		}

		ans, err := xl.cores.SendNum(cmd, name, wss.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	ws.app = xl
	ws.num = num
	ws.parent = wb
	return &ws
}

func (wss *workSheets) Set() *workSheets {
	if wss == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := wss.app
	xl.cores.Lock(wss.num)
	return wss
}

func (ws *workSheet) Release() error {
	xl := ws.app
	xl.cores.Release(ws.num, true)
	return nil
}

func (ws *workSheet) Set() *workSheet {
	if ws == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := ws.app
	xl.cores.Lock(ws.num)
	return ws
}

func (ws *workSheet) Nothing() error {
	xl := ws.app
	xl.cores.releaseChild(ws.num)

	xl.cores.Unlock(ws.num)
	err := ws.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(ws.num)
	ws = nil
	return nil
}

func (ws *workSheet) Name(value ...any) string {
	xl := ws.app

	if len(value) > 0 {
		cmd := "Put"
		name := "Name"
		var opt []any
		switch x := value[0].(type) {
		case string:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		name := "Name"
		ans, err := xl.cores.SendNum(cmd, name, ws.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}

		switch x := ans.(type) {
		case string:
			return x
		}
	}
	return ""
}

func (ws *workSheet) Activate() error {
	xl := ws.app

	cmd := "Method"
	name := "Activate"

	_, err := xl.cores.SendNum(cmd, name, ws.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (ws *workSheet) Parent() *workBook {
	wb := ws.parent
	xl := ws.app

	core := xl.cores.getCore(wb.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Parent"
		ans, err := xl.cores.SendNum(cmd, name, ws.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	return wb
}

func (ws *workSheet) Copy(value ...any) *workSheet {
	var xs workSheet
	xl := ws.app
	wb := ws.parent
	_wb := xl.cores.getCore(wb.num)
	if _wb.disp == nil {
		ws.Parent()
	}

	kind := "Worksheet"
	core, num := xl.cores.FindAdd(kind, wb.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Copy"
		var opt []any
		if len(value) > 0 {
			for i := range value {
				switch x := value[i].(type) {
				case int, int32, string:
					_ws := wb.Worksheetz(x)
					core := xl.cores.getCore(_ws.num)
					opt = append(opt, core.disp)
				case nil:
					opt = append(opt, nil)
				}
			}
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xs.app = xl
	xs.num = num
	if len(value) > 0 {
		xs.parent = wb
	} else {
		wb := xl.ActiveWorkbook()
		xs.parent = wb
		core.parent = wb.num
	}
	return &xs
}

func (ws *workSheet) Move(value ...any) *workSheet {
	var xs workSheet
	xl := ws.app
	wb := ws.parent
	_wb := xl.cores.getCore(wb.num)
	if _wb.disp == nil {
		ws.Parent()
	}

	kind := "Worksheet"
	core, num := xl.cores.FindAdd(kind, wb.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Move"
		var opt []any
		if len(value) > 0 {
			for i := range value {
				switch x := value[i].(type) {
				case int, int32, string:
					_ws := wb.Worksheetz(x)
					core := xl.cores.getCore(_ws.num)
					opt = append(opt, core.disp)
				case nil:
					opt = append(opt, nil)
				}
			}
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 0
		}
	}
	xs.app = xl
	xs.num = num
	if len(value) > 0 {
		xs.parent = wb
	} else {
		wb := xl.ActiveWorkbook()
		xs.parent = wb
		core.parent = wb.num
	}
	return &xs
}
