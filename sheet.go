package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type workSheets struct {
	app    *Excel
	parent any
	num    int
}

type workSheet struct {
	app    *Excel
	parent any
	num    int
}

func getBook(ws *workSheet) *workBook {
	var wb *workBook

	switch x := ws.parent.(type) {
	case *workBook:
		wb = x
	default:
		xl := ws.app
		ws = xl.ActiveSheet()
	}

	return wb
}

func (Q *workBook) Worksheets() *workSheets {
	var body workSheets
	xl := Q.app

	name := "Worksheets"
	core, num := xl.cores.FindAdd(name, Q.num)
	if core.disp == nil {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = Q
	return &body
}

func (Q *workBook) Worksheetz(value any) *workSheet {
	var body workSheet
	xl := Q.app

	kind := "Worksheet"
	core, num := xl.cores.FindAdd(kind, Q.num)
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

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = Q
	return &body
}

func (Q *Excel) ActiveSheet() *workSheet {
	var body workSheet
	xl := Q
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
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = wb
	wb.Release()
	return &body
}

func (Q *workSheets) Release() error {
	xl := Q.app
	return xl.cores.Release(Q.num, false)
}

func (Q *workSheets) Nothing() error {
	xl := Q.app
	xl.cores.releaseChild(Q.num)

	xl.cores.Unlock(Q.num)
	err := Q.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(Q.num)
	Q = nil
	return nil
}

func (Q *workSheets) Count() int32 {
	xl := Q.app

	cmd := "Get"
	name := "Count"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return 0
	}

	switch x := ans.(type) {
	case int32:
		return x
	}
	return 0
}

func (Q *workSheets) Add(value ...any) *workSheet {
	var body workSheet
	xl := Q.app
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

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = wb
	return &body
}

func (Q *workSheets) Set() (*workSheets, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workSheet) Release() error {
	xl := Q.app
	xl.cores.Release(Q.num, true)
	return nil
}

func (Q *workSheet) Set() (*workSheet, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workSheet) Nothing() error {
	xl := Q.app
	xl.cores.releaseChild(Q.num)

	xl.cores.Unlock(Q.num)
	err := Q.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(Q.num)
	Q = nil
	return nil
}

func (Q *workSheet) Name(value ...any) string {
	xl := Q.app

	name := "Name"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case string:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workSheet) Activate() error {
	xl := Q.app

	cmd := "Method"
	name := "Activate"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workSheet) Select() error {
	xl := Q.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workSheet) Parent() *workBook {
	wb := getBook(Q)
	xl := Q.app

	core := xl.cores.getCore(wb.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Parent"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	return wb
}

func (Q *workSheet) Copy(value ...any) *workSheet {
	var body workSheet
	xl := Q.app
	wb := getBook(Q)
	_wb := xl.cores.getCore(wb.num)
	if _wb.disp == nil {
		Q.Parent()
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

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	if len(value) > 0 {
		body.parent = wb
	} else {
		wb := xl.ActiveWorkbook()
		body.parent = wb
		core.parent = wb.num
	}
	return &body
}

func (Q *workSheet) Move(value ...any) *workSheet {
	var body workSheet
	xl := Q.app
	wb := getBook(Q)
	_wb := xl.cores.getCore(wb.num)
	if _wb.disp == nil {
		Q.Parent()
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

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	if len(value) > 0 {
		body.parent = wb
	} else {
		wb := xl.ActiveWorkbook()
		body.parent = wb
		core.parent = wb.num
	}
	return &body
}

func (Q *workSheet) Delete() error {
	xl := Q.app
	cmd := "Method"
	name := "Delete"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		log.Printf("(Error) cmd:%v name:%v ", cmd, name)
	}
	return nil
}

func (Q *workSheet) Visible(value bool) error {
	xl := Q.app
	cmd := "Put"
	name := "Visible"
	var opt []any
	opt = append(opt, value)

	_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
	if err != nil {
		log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
	}
	return nil
}

func (Q *workSheet) Paste(option ...any) bool {
	xl := Q.app

	//Destination *workRange, Link bool

	cmd := "Method"
	name := "Paste"
	var opt []any
	if len(option) > 0 {
		for range 2 {
			opt = append(opt, nil)
		}

		for i := range option {
			switch x := option[i].(type) {
			case *workRange:
				core := xl.cores.getCore(x.num)
				opt[0] = core.disp
				opt[1] = false
			case bool:
				opt[1] = x
			}
		}
	} else {
		opt = nil
	}

	ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
	if err != nil {
		return false
	}
	switch x := ans.(type) {
	case bool:
		return x
	}
	return false
}
