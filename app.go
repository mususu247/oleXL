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

func (xl *Excel) Application() *workApp {
	var wa workApp

	kind := "Application"
	_, num := xl.cores.FindAdd(kind, xl.num)
	wa.app = xl
	wa.num = num
	return &wa
}

func (wa *workApp) Left(value ...float64) float64 {
	xl := wa.app

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

func (wa *workApp) Top(value ...float64) float64 {
	xl := wa.app

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

func (wa *workApp) Width(value ...float64) float64 {
	xl := wa.app

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

func (wa *workApp) Height(value ...float64) float64 {
	xl := wa.app

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

func (wa *workApp) WindowState(value ...any) int32 {
	xl := wa.app

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

		ans, err := xl.cores.SendNum(cmd, name, wa.num, nil)
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

func (wa *workApp) SetWindowRect(left, top, height, width float64) error {
	xl := wa.app
	wa.WindowState("xlNormal")
	wa.Left(left)
	wa.Top(top)
	wa.Height(height)
	wa.Width(width)

	if xl.cores.debug {
		log.Printf(".WindowRect(%v,%v,%v,%v)", wa.Left(), wa.Top(), wa.Height(), wa.Width())
	}
	return nil
}

func (wa *workApp) Run(Macro string, Args ...any) any {
	xl := wa.app

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

func (xl *Excel) ActiveWindow() *workWindow {
	var ww workWindow

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
	ww.app = xl
	ww.num = num
	ww.parent = xl
	return &ww
}

func (ww *workWindow) FreezePanes(value ...bool) bool {
	xl := ww.app

	name := "FreezePanes"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, ww.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ww.num, nil)
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

func (ww *workWindow) Zoom(value ...float64) float64 {
	xl := ww.app

	name := "Zoom"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, ww.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ww.num, nil)
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

func (wa *workApp) CutCopyMode(value ...bool) bool {
	xl := wa.app

	name := "CutCopyMode"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wa.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wa.num, nil)
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

func (wa *workApp) StatusBar(value string) bool {
	xl := wa.app

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

func (wa *workApp) Union(rag ...any) *workRange {
	var wr workRange
	xl := wa.app
	ws := xl.ActiveSheet()

	kind := "Range"
	core, num := xl.cores.FindAdd(kind, wa.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Union"
		var opt []any
		for i := range rag {
			switch x := rag[i].(type) {
			case string:
				opt = append(opt, x)
			case *workRange:
				disp := xl.cores.getCore(x.num).disp
				opt = append(opt, disp)
			}
		}

		ans, err := xl.cores.SendNum(cmd, name, xl.num, opt)
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
		default:
			log.Printf("%T %v", ans, ans)
		}
	}
	wr.app = xl
	wr.num = num
	wr.parent = ws
	return &wr
}
