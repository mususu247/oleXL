package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workFrame struct {
	app    *Excel
	parent *workShape
	num    int
}

func (sp *workShape) TextFrame() *workFrame {
	var wf workFrame
	xl := sp.app

	name := "TextFrame"
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
			core.lock = 0
		}
	}
	wf.app = xl
	wf.num = num
	wf.parent = sp
	return &wf
}

func (wf *workFrame) Release() error {
	xl := wf.app
	return xl.cores.Release(wf.num, false)
}

func (wf *workFrame) Nothing() error {
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
