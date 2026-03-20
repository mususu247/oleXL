package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workLine struct {
	app    *Excel
	parent any
	num    int
}

func (sp *workShape) Line() *workLine {
	var wl workLine
	xl := sp.app

	name := "Line"
	core, num := xl.cores.FindAdd(name, sp.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, sp.num, nil)
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
	wl.app = xl
	wl.num = num
	wl.parent = sp
	return &wl
}

func (wf *workFormat) Line() *workLine {
	var wl workLine
	xl := wf.app

	name := "Line"
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
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	wl.app = xl
	wl.num = num
	wl.parent = wf
	return &wl
}

func (wl *workLine) Release() error {
	xl := wl.app
	return xl.cores.Release(wl.num, false)
}

func (wl *workLine) Nothing() error {
	xl := wl.app
	xl.cores.releaseChild(wl.num)

	xl.cores.Unlock(wl.num)
	err := wl.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(wl.num)
	wl = nil
	return nil
}

func (wl *workLine) Weight(value ...float64) float64 {
	xl := wl.app

	name := "Weight"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wl.num, nil)
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

func (wl *workLine) Visible(value ...bool) bool {
	xl := wl.app

	name := "Visible"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, wl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wl.num, nil)
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
