package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workChar struct {
	app    *Excel
	parent any
	num    int
}

func (wf *workFrame) Characterz(value ...any) *workChar {
	var ch workChar
	xl := wf.app

	name := "Characters"
	core, num := xl.cores.FindAdd(name, wf.num)
	if core.disp == nil {
		cmd := "Method"
		var opt []any
		if len(value) > 0 {
			for i := range value {
				switch x := value[i].(type) {
				case int:
					opt = append(opt, int32(x))
				case int32:
					opt = append(opt, x)
				}
			}
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, wf.num, opt)
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
	ch.app = xl
	ch.num = num
	ch.parent = wf
	return &ch
}

func (wf *workFrame) Characters() *workChar {
	var ch workChar
	xl := wf.app

	name := "Characters"
	core, num := xl.cores.FindAdd(name, wf.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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
	ch.app = xl
	ch.num = num
	ch.parent = wf
	return &ch
}

func (tr *workTextRange) Characterz(value ...any) *workChar {
	var ch workChar
	xl := tr.app

	name := "Characters"
	core, num := xl.cores.FindAdd(name, tr.num)
	if core.disp == nil {
		cmd := "Method"
		var opt []any
		if len(value) > 0 {
			for i := range value {
				switch x := value[i].(type) {
				case int:
					opt = append(opt, int32(x))
				case int32:
					opt = append(opt, x)
				}
			}
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, tr.num, opt)
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
	ch.app = xl
	ch.num = num
	ch.parent = tr
	return &ch
}

func (tr *workTextRange) Characters() *workChar {
	var ch workChar
	xl := tr.app

	name := "Characters"
	core, num := xl.cores.FindAdd(name, tr.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, tr.num, nil)
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
	ch.app = xl
	ch.num = num
	ch.parent = tr
	return &ch
}

func (ch *workChar) Release() error {
	xl := ch.app
	return xl.cores.Release(ch.num, false)
}

func (ch *workChar) Nothing() error {
	xl := ch.app
	xl.cores.releaseChild(ch.num)

	xl.cores.Unlock(ch.num)
	err := ch.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(ch.num)
	ch = nil
	return nil
}

func (ch *workChar) Text(value ...string) string {
	xl := ch.app

	name := "Text"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, ch.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ch.num, nil)
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

func (ch *workChar) Count() int32 {
	xl := ch.app

	name := "Count"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, ch.num, nil)
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

func (ch *workChar) Insert(value string) error {
	xl := ch.app

	name := "Text"
	cmd := "Put"
	var opt []any
	opt = append(opt, value)
	_, err := xl.cores.SendNum(cmd, name, ch.num, opt)
	if err != nil {
		log.Printf("(Error) %v", err)
		return err
	}
	return nil
}
