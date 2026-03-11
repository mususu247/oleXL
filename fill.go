package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workFill struct {
	app    *Excel
	parent any
	num    int
}

func (sp *workShape) Fill() *workFill {
	var wl workFill
	xl := sp.app

	name := "Fill"
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
			core.disp = x
			core.lock = 1 //Lock.on
		}
	}
	wl.app = xl
	wl.num = num
	wl.parent = sp
	return &wl
}

func (wf *workFont) Fill() *workFill {
	var wl workFill
	xl := wf.app

	name := "Fill"
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
			core.lock = 1 //Lock.on
		}
	}
	wl.app = xl
	wl.num = num
	wl.parent = wf
	return &wl
}

func (wl *workFill) Release() error {
	xl := wl.app
	return xl.cores.Release(wl.num, false)
}

func (wl *workFill) Nothing() error {
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

func (wl *workFill) Visible(value bool) error {
	xl := wl.app
	cmd := "Put"
	name := "Visible"
	var opt []any
	opt = append(opt, value)

	_, err := xl.cores.SendNum(cmd, name, wl.num, opt)
	if err != nil {
		log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
	}
	return nil
}

func (wf *workFill) Transparency(value ...float64) float64 {
	xl := wf.app

	name := "Transparency"
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
