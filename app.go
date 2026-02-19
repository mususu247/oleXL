package oleXL

import (
	"log"
	"time"

	"github.com/go-ole/go-ole"
)

type workApp struct {
	app *Excel
	num int
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

	if len(value) > 0 {
		cmd := "Put"
		name := "Left"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		name := "Left"
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

	if len(value) > 0 {
		cmd := "Put"
		name := "Top"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		name := "Top"
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

	if len(value) > 0 {
		cmd := "Put"
		name := "Width"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		name := "Width"
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

	if len(value) > 0 {
		cmd := "Put"
		name := "Height"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		name := "Height"
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

	if len(value) > 0 {
		cmd := "Put"
		name := "WindowState"
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
		name := "WindowState"
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
