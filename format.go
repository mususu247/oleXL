package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workFormat struct {
	app    *Excel
	parent *workTitle
	num    int
}

func (wt *workTitle) Format() *workFormat {
	var wf workFormat
	xl := wt.app

	name := "TextFrame"
	core, num := xl.cores.FindAdd(name, wt.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wt.num, nil)
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
	wf.parent = wt
	return &wf
}

func (wf *workFormat) Release() error {
	xl := wf.app
	return xl.cores.Release(wf.num, false)
}

func (wf *workFormat) Nothing() error {
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
