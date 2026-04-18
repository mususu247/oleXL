package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type workWindow struct {
	app    *Excel
	parent any
	num    int
}

type workWindows struct {
	app    *Excel
	parent any
	num    int
}

func (Q *workApp) Windows() *workWindows {
	var body workWindows
	xl := Q.app

	kind := "Window"
	core, num := xl.cores.FindAdd(kind, xl.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Windows"
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
		default:
			log.Printf("(Error) %v", x)
			return nil
		}
	}
	body.app = xl
	body.num = num
	body.parent = Q
	return &body
}

func (Q *workApp) Windowz(value any) *workWindow {
	var body workWindow
	xl := Q.app

	kind := "Window"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Windows"
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

func (Q *workBook) Windows() *workWindows {
	var body workWindows
	xl := Q.app

	kind := "Window"
	core, num := xl.cores.FindAdd(kind, xl.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Windows"
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

func (Q *workBook) Windowz(value any) *workWindow {
	var body workWindow
	xl := Q.app

	kind := "Window"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Windows"
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

func (Q *workWindows) Release() error {
	xl := Q.app
	xl.cores.Release(Q.num, true)
	return nil
}

func (Q *workWindows) Nothing() error {
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

func (Q *workWindows) Count() int32 {
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

func (Q *workBook) NewWindow() *workWindow {
	var body workWindow
	xl := Q.app

	kind := "Window"
	core, num := xl.cores.FindAdd(kind, xl.num)
	if core.disp == nil {
		cmd := "Method"
		name := "NewWindow"
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

func (Q *workWindow) NewWindow() *workWindow {
	var body workWindow
	xl := Q.app

	kind := "Window"
	core, num := xl.cores.FindAdd(kind, xl.num)
	if core.disp == nil {
		cmd := "Method"
		name := "NewWindow"
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

func (Q *workWindow) Release() error {
	xl := Q.app
	xl.cores.Release(Q.num, true)
	return nil
}

func (Q *workWindow) Set() (*workWindow, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workWindow) Nothing() error {
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

func (Q *workWindows) Arrange(value ...any) bool {
	xl := Q.app

	name := "Arrange"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any

		for i := range value {
			switch i {
			case 0:
				var z int32
				switch x := value[0].(type) {
				case int:
					z = SetEnumArrangeStyle(int32(x))
				case int32:
					z = SetEnumArrangeStyle(x)
				case string:
					z = GetEnumArrangeStyleNum(x)
				}
				opt = append(opt, z)
			case 1:
				// ActiveWorkbook
				switch x := value[1].(type) {
				case bool:
					opt = append(opt, x)
				}
			case 2:
				// SyncHorizontal
				switch x := value[2].(type) {
				case bool:
					opt = append(opt, x)
				}
			case 3:
				// SyncVertical
				switch x := value[3].(type) {
				case bool:
					opt = append(opt, x)
				}
			}
		}

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

func (Q *workWindow) Close(SaveChanges ...bool) error {
	xl := Q.app
	xl.cores.releaseChild(Q.num)

	cmd := "Method"
	name := "Close"

	var opt []any
	if len(SaveChanges) > 0 {
		opt = append(opt, SaveChanges[0])
	}

	_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
	if err != nil {
		return err
	}

	return nil
}
