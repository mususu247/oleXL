package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type workInterior struct {
	app    *Excel
	parent any
	num    int
}

func (wr *workRange) Interior() *workInterior {
	var wi workInterior
	xl := wr.app

	name := "Interior"
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
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	wi.app = xl
	wi.num = num
	wi.parent = wr
	return &wi
}

func (wi *workInterior) Release() error {
	xl := wi.app
	return xl.cores.Release(wi.num, false)
}

func (wi *workInterior) Nothing() error {
	xl := wi.app
	xl.cores.releaseChild(wi.num)

	xl.cores.Unlock(wi.num)
	err := wi.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(wi.num)
	wi = nil
	return nil
}

func (wi *workInterior) Set() (*workInterior, error) {
	if wi == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := wi.app
	xl.cores.Lock(wi.num)
	return wi, nil
}
