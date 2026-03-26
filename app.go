package oleXL

import (
	"log"
	"time"

	"github.com/go-ole/go-ole"
)

type workApp struct {
	app    *Excel
	parent any
	num    int
}

type workWindow struct {
	app    *Excel
	parent any
	num    int
}

func (Q *Excel) Application() *workApp {
	var body workApp
	xl := Q

	kind := "Application"
	_, num := xl.cores.FindAdd(kind, xl.num)
	body.app = xl
	body.num = num
	return &body
}

func (Q *workApp) Left(value ...float64) float64 {
	xl := Q.app

	name := "Left"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case float64:
			return x
		}
	}
	return 0
}

func (Q *workApp) Top(value ...float64) float64 {
	xl := Q.app

	name := "Top"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case float64:
			return x
		}
	}
	return 0
}

func (Q *workApp) Width(value ...float64) float64 {
	xl := Q.app

	name := "Width"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case float64:
			return x
		}
	}
	return 0
}

func (Q *workApp) Height(value ...float64) float64 {
	xl := Q.app

	name := "Height"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case float64:
			return x
		}
	}
	return 0
}

func (Q *workApp) WindowState(value ...any) int32 {
	xl := Q.app

	name := "WindowState"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any

		var v int32
		switch x := value[0].(type) {
		case int:
			v = SetEnumWindowState(int32(x))
		case int32:
			v = SetEnumWindowState(x)
		case string:
			v = GetEnumWindowStateNum(x)
		}
		opt = append(opt, v)

		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case int32:
			return x
		}
	}
	return 0
}

func (Q *workApp) SetWindowRect(left, top, height, width float64) error {
	xl := Q.app
	Q.WindowState("xlNormal")
	Q.Left(left)
	Q.Top(top)
	Q.Height(height)
	Q.Width(width)

	if xl.cores.debug {
		log.Printf(".WindowRect(%v,%v,%v,%v)", Q.Left(), Q.Top(), Q.Height(), Q.Width())
	}
	return nil
}

func (Q *workApp) Run(Macro string, Args ...any) any {
	xl := Q.app

	cmd := "Method"
	name := "Run"
	var opt []any
	opt = append(opt, Macro)

	if len(Args) > 0 {
		for i := range Args {
			if i > 29 {
				break
			}
			switch x := Args[i].(type) {
			case int:
				opt = append(opt, int32(x))
			case int32:
				opt = append(opt, x)
			case float64:
				opt = append(opt, x)
			case string:
				opt = append(opt, x)
			case bool:
				opt = append(opt, x)
			case time.Time:
				opt = append(opt, x)
			case nil:
				opt = append(opt, x)
			}
		}
	}
	ans, err := xl.cores.SendNum(cmd, name, xl.num, opt)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	switch x := ans.(type) {
	case int32:
		return x
	case float64:
		return x
	case string:
		return x
	case bool:
		return x
	case time.Time:
		return x
	case nil:
		return x
	case *ole.VARIANT:
		switch x.Val {
		case 2148141008:
			return "#NULL!"
		case 2148141015:
			return "#DIV/0!"
		case 2148141023:
			return "#VALUE!"
		case 2148141031:
			return "#REF!"
		case 2148141037:
			return "#NAME?"
		case 2148141044:
			return "#NUM!"
		case 2148141050:
			return "#N/A"
		}
		return x.Val
	default:
		return x
	}
}

func (Q *Excel) ActiveWindow() *workWindow {
	var body workWindow
	xl := Q

	kind := "Window"
	core, num := xl.cores.FindAdd(kind, xl.num)
	if core.disp == nil {
		cmd := "Get"
		name := "ActiveWindow"
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
	body.parent = xl
	return &body
}

func (Q *workWindow) FreezePanes(value ...bool) bool {
	xl := Q.app

	name := "FreezePanes"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
		switch x := ans.(type) {
		case bool:
			return x
		}
	}
	return false
}

func (Q *workWindow) Zoom(value ...float64) float64 {
	xl := Q.app

	name := "Zoom"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case float64:
			return x
		}
	}
	return 0
}

func (Q *workApp) CutCopyMode(value ...bool) bool {
	xl := Q.app

	name := "CutCopyMode"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
		switch x := ans.(type) {
		case bool:
			return x
		}
	}
	return false
}

func (Q *workApp) StatusBar(value string) bool {
	xl := Q.app

	name := "StatusBar"
	cmd := "Put"
	var opt []any
	if len(value) > 0 {
		opt = append(opt, value)
	} else {
		opt = append(opt, false)
	}
	ans, _ := xl.cores.SendNum(cmd, name, xl.num, opt)
	switch x := ans.(type) {
	case bool:
		return x
	}
	return false
}

func (Q *workApp) ReferenceStyle(value ...any) int32 {
	xl := Q.app

	name := "ReferenceStyle"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any

		var v int32
		switch x := value[0].(type) {
		case int:
			v = SetEnumReferenceStyle(int32(x))
		case int32:
			v = SetEnumReferenceStyle(x)
		case string:
			v = GetEnumReferenceStyleNum(x)
		}
		opt = append(opt, v)

		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case int32:
			return x
		}
	}
	return 0
}
