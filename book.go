package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workBooks struct {
	app    *Excel
	parent any
	num    int
}

type workBook struct {
	app    *Excel
	parent any
	num    int
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
	wbs := xl.Workbooks()

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
	wb.parent = wbs
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

func (wb *workBook) Path() string {
	var result string
	xl := wb.app

	cmd := "Get"
	name := "Path"
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

func (wbs *workBooks) Open(fileName string, options ...map[string]any) *workBook {
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

		if len(options) > 0 {
			for range 14 {
				opt = append(opt, nil)
			}

			for k, v := range options[0] {
				switch k {
				case "Filename":
					switch x := v.(type) {
					case string:
						opt[0] = x
					}
				case "UpdateLinks":
					switch x := v.(type) {
					case int32:
						switch x {
						case 0, 1, 2, 3:
							opt[1] = x
						default:
							opt[1] = int32(0) // Default value if not in range
						}
					default:
						opt[1] = int32(0)
					}
				case "ReadOnly":
					switch x := v.(type) {
					case bool:
						opt[2] = x
					default:
						opt[2] = nil
					}
				case "Format":
					switch x := v.(type) {
					case int32:
						switch x {
						case 1, 2, 3, 4, 5, 6:
							opt[3] = x
						default:
							opt[3] = int32(1) // Default value if not in range
						}
					default:
						opt[3] = int32(1) // Default value if not in range
					}
				case "Password":
					switch x := v.(type) {
					case string:
						opt[4] = x
					default:
						opt[4] = nil
					}
				case "WriteResPassword":
					switch x := v.(type) {
					case bool:
						opt[5] = x
					default:
						opt[5] = nil
					}
				case "IgnoreReadOnlyRecommended":
					switch x := v.(type) {
					case bool:
						opt[6] = x
					default:
						opt[6] = nil
					}
				case "Origin":
					var z int32
					switch x := v.(type) {
					case int32:
						z = SetEnumPlatform(x)
					case int:
						z = SetEnumPlatform(int32(x))
					case string:
						z = GetEnumPlatformNum(x)
					}
					opt[7] = z
				case "Delimiter":
					switch x := v.(type) {
					case string:
						opt[8] = x
					default:
						opt[8] = ","
					}
				case "Editable":
					switch x := v.(type) {
					case bool:
						opt[9] = x
					default:
						opt[9] = false
					}
				case "Notify":
					switch x := v.(type) {
					case bool:
						opt[10] = x
					default:
						opt[10] = false
					}
				case "Converter":
					switch x := v.(type) {
					case int32:
						opt[11] = x
					default:
						opt[11] = int32(0)
					}
				case "AddToMru":
					switch x := v.(type) {
					case bool:
						opt[12] = x
					default:
						opt[12] = false
					}
				case "Local":
					switch x := v.(type) {
					case bool:
						opt[13] = x
					default:
						opt[13] = false
					}
				case "CorruptLoad":
					opt[14] = v
					switch x := v.(type) {
					case int32:
						opt[14] = SetEnumCorruptLoad(x)
					case int:
						opt[14] = SetEnumCorruptLoad(int32(x))
					case string:
						opt[14] = GetEnumCorruptLoadNum(x)
					default:
						opt[14] = nil
					}
				}
			}
		}

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

func (wb *workBook) SaveAs(fileName string, options ...map[string]any) error {
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

	if len(options) > 0 {
		for range 11 {
			opt = append(opt, nil)
		}

		for k, v := range options[0] {
			switch k {
			case "FileFormat":
				var z int32
				switch x := v.(type) {
				case int:
					z = SetEnumFileFormat(int32(x))
				case int32:
					z = SetEnumFileFormat(x)
				case string:
					z = GetEnumFileFormatNum(x)
				}
				opt[1] = z
			case "Password":
				switch x := v.(type) {
				case string:
					opt[2] = x
				default:
					opt[2] = nil
				}
			case "WriteResPassword":
				switch x := v.(type) {
				case string:
					opt[3] = x
				default:
					opt[3] = nil
				}
			case "ReadOnlyRecommended":
				switch x := v.(type) {
				case bool:
					opt[4] = x
				default:
					opt[4] = nil
				}
			case "CreateBackup":
				switch x := v.(type) {
				case bool:
					opt[5] = x
				default:
					opt[5] = nil
				}
			case "AccessMode":
				switch x := v.(type) {
				case int32:
					switch x {
					case 1, 2, 3:
						opt[6] = x
					default:
						opt[6] = int32(1) // Default value if not in range
					}
				default:
					opt[6] = int32(1) // Default value if not in range
				}
			case "ConflictResolution":
				switch x := v.(type) {
				case int32:
					switch x {
					case 1, 2, 3:
						opt[7] = x
					default:
						opt[7] = int32(1) // Default value if not in range
					}
				default:
					opt[7] = int32(1) // Default value if not in range
				}
			case "AddToMru":
				switch x := v.(type) {
				case bool:
					opt[8] = x
				default:
					opt[8] = nil
				}
			case "TextCodepage":
				opt[9] = v
			case "TextVisualLayout":
				opt[10] = v
			case "Local":
				switch x := v.(type) {
				case bool:
					opt[11] = x
				default:
					opt[11] = nil
				}
			}
		}
	}

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

func (wb *workBook) ReadOnly() bool {
	xl := wb.app

	cmd := "Get"
	name := "ReadOnly"
	ans, _ := xl.cores.SendNum(cmd, name, wb.num, nil)
	switch x := ans.(type) {
	case bool:
		return x
	}
	return false
}
