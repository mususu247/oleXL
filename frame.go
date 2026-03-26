package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workFrame struct {
	app    *Excel
	parent any
	num    int
}

func (Q *workShape) TextFrame() *workFrame {
	var body workFrame
	xl := Q.app

	name := "TextFrame"
	core, num := xl.cores.FindAdd(name, Q.num)
	if core.disp == nil {
		cmd := "Get"
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

func (Q *workFrame) Release() error {
	xl := Q.app
	return xl.cores.Release(Q.num, false)
}

func (Q *workFrame) Nothing() error {
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
