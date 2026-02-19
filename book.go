package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workBooks struct {
	app *Excel
	num int
}

type workBook struct {
	app *Excel
	num int
}

func (xl *Excel) Workbooks() *workBooks {
	var wbs workBooks

	kind := "Workbooks"
	core, num := xl.cores.FindAdd(kind, xl.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Workbooks"
		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
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

	wbs.app = xl
	wbs.num = num
	return &wbs
}

func (xl *Excel) ActiveWorkbook() *workBook {
	var wb workBook
	//wbs := xl.Workbooks()

	kind := "Workbook"
	core, num := xl.cores.FindAdd(kind, xl.num)
	if core.disp == nil {
		cmd := "Get"
		name := "ActiveWorkbook"
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
	wb.app = xl
	wb.num = num
	//wbs.Release()
	return &wb
}

func (xl *Excel) Workbookz(value any) *workBook {
	var wb workBook
	wbs := xl.Workbooks()

	kind := "Workbook"
	core, num := xl.cores.FindAdd(kind, wbs.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Workbooks"
		var opt []any

		switch x := value.(type) {
		case int:
			opt = append(opt, int32(x))
		case int32:
			opt = append(opt, x)
		case string:
			opt = append(opt, x)
		}

		ans, err := xl.cores.SendNum(cmd, name, xl.num, opt)
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
	wb.app = xl
	wb.num = num
	return &wb
}

func (wbs *workBooks) Release() error {
	xl := wbs.app
	xl.cores.Release(wbs.num, true)
	return nil
}

func (wbs *workBooks) Nothing() error {
	xl := wbs.app
	xl.cores.releaseChild(wbs.num)

	xl.cores.Unlock(wbs.num)
	err := wbs.Release()
	if err != nil {
		return err
	}

	xl.cores.Remove(wbs.num)
	wbs = nil
	return nil
}

func (wbs *workBooks) Count() int32 {
	var result int32
	xl := wbs.app

	cmd := "Get"
	name := "Count"
	ans, err := xl.cores.SendNum(cmd, name, wbs.num, nil)
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

func (wbs *workBooks) Add() *workBook {
	var wb workBook
	xl := wbs.app

	kind := "Workbook"
	core, num := xl.cores.FindAdd(kind, xl.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Add"
		ans, err := xl.cores.SendNum(cmd, name, wbs.num, nil)
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
	wb.app = xl
	wb.num = num
	return &wb
}

func (wbs *workBooks) Set() *workBooks {
	if wbs == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := wbs.app
	xl.cores.Lock(wbs.num)
	return wbs
}

func (wb *workBook) Release() error {
	xl := wb.app
	xl.cores.Release(wb.num, true)
	return nil
}

func (wb *workBook) Set() *workBook {
	if wb == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := wb.app
	xl.cores.Lock(wb.num)
	return wb
}

func (wb *workBook) Nothing() error {
	xl := wb.app
	xl.cores.releaseChild(wb.num)

	xl.cores.Unlock(wb.num)
	err := wb.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(wb.num)
	wb = nil
	return nil
}

func (wb *workBook) Close(SaveChanges ...bool) error {
	xl := wb.app
	xl.cores.releaseChild(wb.num)

	cmd := "Method"
	name := "Close"

	var opt []any
	if len(SaveChanges) > 0 {
		opt = append(opt, SaveChanges[0])
	}

	_, err := xl.cores.SendNum(cmd, name, wb.num, opt)
	if err != nil {
		return err
	}

	return nil
}

func (wb *workBook) Name() string {
	var result string
	xl := wb.app

	cmd := "Get"
	name := "Name"
	ans, err := xl.cores.SendNum(cmd, name, wb.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return result
	}

	switch x := ans.(type) {
	case string:
		result = x
	}
	return result
}

func (wb *workBook) RefreshAll() error {
	xl := wb.app

	cmd := "Method"
	name := "RefreshAll"
	_, err := xl.cores.SendNum(cmd, name, wb.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (wbs *workBooks) Open(fileName string) *workBook {
	var wb workBook
	xl := wbs.app

	fn, err := GetAbsolutePathName(fileName)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}

	if !FileExists(fn) {
		log.Printf("(Error) File not found: %v", fn)
		return nil
	}

	kind := "Workbook"
	core, num := xl.cores.FindAdd(kind, xl.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Open"
		var opt []any
		opt = append(opt, fn)

		ans, err := xl.cores.SendNum(cmd, name, wbs.num, opt)
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
	wb.app = xl
	wb.num = num
	return &wb
}

func (wb *workBook) SaveAs(fileName string, fileFormat ...any) error {
	xl := wb.app

	fn, err := GetAbsolutePathName(fileName)
	if err != nil {
		return err
	}

	if FileExists(fn) {
		DeleteFile(fn)
	}

	cmd := "Method"
	name := "SaveAs"
	var opt []any
	opt = append(opt, fn)

	var ff int32 = -4143 // xlWorkbookDefault
	if len(fileFormat) > 0 {
		switch x := fileFormat[0].(type) {
		case int:
			ff = SetEnumFileFormat(int32(x))
		case int32:
			ff = SetEnumFileFormat(x)
		case string:
			ff = GetEnumFileFormatNum(x)
		}
	}
	opt = append(opt, ff)

	_, err = xl.cores.SendNum(cmd, name, wb.num, opt)
	if err != nil {
		return err
	}
	return nil
}

func (wb *workBook) Save() error {
	xl := wb.app

	cmd := "Method"
	name := "Save"

	_, err := xl.cores.SendNum(cmd, name, wb.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (wb *workBook) SaveCopyAs(fileName string) error {
	xl := wb.app

	fn, err := GetAbsolutePathName(fileName)
	if err != nil {
		return err
	}

	if FileExists(fn) {
		DeleteFile(fn)
	}

	cmd := "Method"
	name := "SaveCopyAs"
	var opt []any
	opt = append(opt, fn)

	_, err = xl.cores.SendNum(cmd, name, wb.num, opt)
	if err != nil {
		return err
	}
	return nil
}

func (wb *workBook) Activate() error {
	xl := wb.app

	cmd := "Method"
	name := "Activate"

	_, err := xl.cores.SendNum(cmd, name, wb.num, nil)
	if err != nil {
		return err
	}
	return nil
}
