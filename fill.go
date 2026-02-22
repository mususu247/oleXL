package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workFill struct {
	app    *Excel
	parent *workShape
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
