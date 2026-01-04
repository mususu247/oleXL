package oleXL

import (
	"log"
)

// version 2025-12-19
// VBA style like

type any = interface{}

type application struct {
	app *Excel
	mx  int
}

func (xl *Excel) Application() *application {
	var app application
	app.app = xl
	app.mx = xl.mx

	return &app
}

func (ap *application) Hand() int32 {
	xl := ap.app
	return xl.hand()
}

func (ap *application) WindowState(value any) int32 {
	var result int32
	xl := ap.app
	_core, err := xl.getCore(ap.mx)
	if err != nil {
		return -1
	}
	var cmd = "Get"
	const name = "WindowState"
	var opt []any

	if value != nil {
		switch x := value.(type) {
		case int:
			result = SetEnumWindowState(int32(x))
		case int32:
			result = SetEnumWindowState(x)
		case string:
			result = GetEnumWindowStateNum(x)
		}
		opt = append(opt, result)

		cmd = "Put"
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

	result = -1
	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case int32:
			log.Printf("%v ans (int32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			result = x
			_core.values["WindowState"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return result
}

func (ap *application) Left(value any) float64 {
	var result float64
	xl := ap.app
	_core, err := xl.getCore(ap.mx)
	if err != nil {
		return -1
	}
	var cmd string
	const name = "Left"
	var opt []any

	switch x := value.(type) {
	case float64:
		opt = append(opt, x)
	default:
		opt = nil
	}

	if opt != nil {
		cmd = "Put"
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

	result = -1
	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case float64:
			log.Printf("%v ans (float64) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["Left"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return result
}

func (ap *application) Top(value any) float64 {
	var result float64
	xl := ap.app
	_core, err := xl.getCore(ap.mx)
	if err != nil {
		return -1
	}
	var cmd string
	const name = "Top"
	var opt []any

	switch x := value.(type) {
	case float64:
		opt = append(opt, x)
	default:
		opt = nil
	}

	if opt != nil {
		cmd = "Put"
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

	result = -1
	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case float64:
			log.Printf("%v ans (float64) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["Top"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return result
}

func (ap *application) Width(value any) float64 {
	var result float64
	xl := ap.app
	_core, err := xl.getCore(ap.mx)
	if err != nil {
		return -1
	}
	var cmd string
	const name = "Width"
	var opt []any

	switch x := value.(type) {
	case float64:
		opt = append(opt, x)
	default:
		opt = nil
	}

	if opt != nil {
		cmd = "Put"
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

	result = -1
	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case float64:
			log.Printf("%v ans (float64) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["Width"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return result
}

func (ap *application) Height(value any) float64 {
	var result float64
	xl := ap.app
	_core, err := xl.getCore(ap.mx)
	if err != nil {
		return -1
	}
	var cmd string
	const name = "Height"
	var opt []any

	switch x := value.(type) {
	case float64:
		opt = append(opt, x)
	default:
		opt = nil
	}

	if opt != nil {
		cmd = "Put"
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

	result = -1
	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case float64:
			log.Printf("%v ans (float64) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["Height"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return result
}

func (ap *application) SetWindowRect(left, top, width, height float64) {
	ap.Left(left)
	ap.Top(top)
	ap.Width(width)
	ap.Height(height)
}

func (ap *application) ScreenUpdating(value ...bool) bool {
	var result bool
	xl := ap.app
	_core, err := xl.getCore(ap.mx)
	if err != nil {
		return false
	}
	var cmd string
	const name = "ScreenUpdating"
	var opt []any

	if len(value) > 0 {
		opt = append(opt, value)
	} else {
		opt = nil
	}

	if opt != nil {
		cmd = "Put"
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

	result = false
	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["ScreenUpdating"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return result
}

func (ap *application) DisplayAlerts(value ...bool) bool {
	var result bool
	xl := ap.app
	_core, err := xl.getCore(ap.mx)
	if err != nil {
		return false
	}
	var cmd string
	const name = "DisplayAlerts"
	var opt []any

	if len(value) > 0 {
		opt = append(opt, value[0])
	} else {
		opt = nil
	}

	if opt != nil {
		cmd = "Put"
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

	result = false
	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			result = x
			_core.values["DisplayAlerts"] = result
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return result
}

func (ap *application) Run(macroName string, macroArgs ...any) any {
	xl := ap.app
	_core, err := xl.getCore(ap.mx)
	if err != nil {
		return err
	}
	const cmd = "Method"
	const name = "Run"
	var opt []any

	opt = append(opt, macroName)
	for i := range macroArgs {
		if i > 30 {
			break
		}
		opt = append(opt, macroArgs[i])
	}

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case any:
			log.Printf("%v ans (any) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			return x
		case nil:
			log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			return x
		}
	}
	return nil
}
