package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workFont struct {
	app *Excel
	num int
}

func (wr *workRange) Font() *workFont {
	var wf workFont
	xl := wr.app

	name := "Font"
	core, num := xl.cores.FindAdd(name, wr.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wr.num, nil)
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
	wf.app = xl
	wf.num = num
	return &wf
}

func (ch *workChar) Font() *workFont {
	var wf workFont
	xl := ch.app

	name := "Font"
	core, num := xl.cores.FindAdd(name, ch.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, ch.num, nil)
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
	wf.app = xl
	wf.num = num
	return &wf
}

func (wf *workFont) Release() error {
	xl := wf.app
	return xl.cores.Release(wf.num, false)
}

func (wf *workFont) Nothing() error {
	xl := wf.app
	xl.cores.releaseChild(wf.num)

	xl.cores.Unlock(wf.num)
	err := wf.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(wf.num)
	wf = nil
	return nil
}

func (wf *workFont) Set() *workFont {
	if wf == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := wf.app
	xl.cores.Lock(wf.num)
	return wf
}

func (wf *workFont) Name(value ...string) string {
	xl := wf.app

	name := "Name"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFont) Bold(value ...bool) bool {
	xl := wf.app

	name := "Bold"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFont) Italic(value ...bool) bool {
	xl := wf.app

	name := "Italic"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFont) Size(value ...float64) float64 {
	xl := wf.app

	name := "Size"
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
		case float64:
			return x
		}
	}

	return 0
}

func (wf *workFont) Strikethrough(value ...bool) bool {
	xl := wf.app

	name := "Strikethrough"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFont) Superscript(value ...bool) bool {
	xl := wf.app

	name := "Superscript"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFont) Subscript(value ...bool) bool {
	xl := wf.app

	name := "Subscript"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFont) OutlineFont(value ...bool) bool {
	xl := wf.app

	name := "OutlineFont"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFont) Shadow(value ...bool) bool {
	xl := wf.app

	name := "Shadow"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value)

		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFont) Underline(value ...any) int32 {
	xl := wf.app

	name := "Underline"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumUnderlineStyle(int32(x))
		case int32:
			z = SetEnumUnderlineStyle(x)
		case string:
			z = GetEnumUnderlineStyleNum(x)
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
		case int32:
			return x
		}
	}
	return 0
}
