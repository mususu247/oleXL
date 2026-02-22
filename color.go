package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workColor struct {
	app *Excel
	num int
}

func (wl *workFill) ForeColor() *workColor {
	var wc workColor
	xl := wl.app

	name := "ForeColor"
	core, num := xl.cores.FindAdd(name, wl.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wl.num, nil)
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
	wc.app = xl
	wc.num = num
	return &wc
}

func (wl *workFill) BackColor() *workColor {
	var wc workColor
	xl := wl.app

	name := "BackColor"
	core, num := xl.cores.FindAdd(name, wl.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wl.num, nil)
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
	wc.app = xl
	wc.num = num
	return &wc
}

func (wl *workLine) ForeColor() *workColor {
	var wc workColor
	xl := wl.app

	name := "ForeColor"
	core, num := xl.cores.FindAdd(name, wl.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wl.num, nil)
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
	wc.app = xl
	wc.num = num
	return &wc
}

func (wc *workColor) Release() error {
	xl := wc.app
	return xl.cores.Release(wc.num, false)
}

func (wc *workColor) Nothing() error {
	xl := wc.app
	xl.cores.releaseChild(wc.num)

	xl.cores.Unlock(wc.num)
	err := wc.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(wc.num)
	wc = nil
	return nil
}

func (wc *workColor) RGB(value ...any) float64 {
	xl := wc.app

	name := "RGB"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z float64
		switch x := value[0].(type) {
		case float64:
			z = x
		case string:
			z = GetEnumRgbColorNum(x)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, wc.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wc.num, nil)
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

// Range().Font().Color()
func (wf *workFont) ColorIndex(value ...int32) int32 {
	xl := wf.app

	name := "ColorIndex"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFont) Color(value ...any) float64 {
	xl := wf.app

	name := "Color"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z float64
		switch x := value[0].(type) {
		case float64:
			z = x
		case string:
			z = GetEnumRgbColorNum(x)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

// Range().Interior().Color()
func (wi *workInterior) ColorIndex(value ...int32) int32 {
	xl := wi.app

	name := "ColorIndex"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wi.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wi.num, nil)
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

func (wi *workInterior) Color(value ...any) float64 {
	xl := wi.app

	name := "Color"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z float64
		switch x := value[0].(type) {
		case float64:
			z = x
		case string:
			z = GetEnumRgbColorNum(x)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, wi.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wi.num, nil)
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

// Range().Border().Color()
func (br *workBorder) ColorIndex(value ...int32) int32 {
	xl := br.app

	name := "ColorIndex"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, br.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, br.num, nil)
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

func (br *workBorder) Color(value ...any) float64 {
	xl := br.app

	name := "Color"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z float64
		switch x := value[0].(type) {
		case float64:
			z = x
		case string:
			z = GetEnumRgbColorNum(x)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, br.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, br.num, nil)
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

// Shape().Color()
func (sp *workShape) ColorIndex(value ...int32) int32 {
	xl := sp.app

	name := "ColorIndex"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, sp.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, sp.num, nil)
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

func (sp *workShape) Color(value ...any) float64 {
	xl := sp.app

	name := "Color"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z float64
		switch x := value[0].(type) {
		case float64:
			z = x
		case string:
			z = GetEnumRgbColorNum(x)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, sp.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, sp.num, nil)
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

// RGB(red,green,blue)
func RGB(red, green, blue int) float64 {
	r := uint8(red)
	g := uint8(green)
	b := uint8(blue)

	var color float64
	color = float64(b)
	color = color * 256
	color = color + float64(g)
	color = color * 256
	color = color + float64(r)
	return color
}
