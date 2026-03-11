package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workTitle struct {
	app    *Excel
	parent any
	num    int
}

func (ct *workChart) ChartTitle() *workTitle {
	var wt workTitle
	xl := ct.app

	kind := "ChartTitle"
	core, num := xl.cores.FindAdd(kind, ct.num)
	if core.disp == nil {
		cmd := "Get"
		name := "ChartTitle"

		ans, err := xl.cores.SendNum(cmd, name, ct.num, nil)
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
	wt.app = xl
	wt.num = num
	return &wt
}

func (wt *workTitle) Release() error {
	xl := wt.app
	xl.cores.Release(wt.num, false)
	return nil
}

func (wt *workTitle) Nothing() error {
	xl := wt.app
	xl.cores.releaseChild(wt.num)

	xl.cores.Unlock(wt.num)
	err := wt.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(wt.num)
	wt = nil
	return nil
}

func (wt *workTitle) Set() *workTitle {
	if wt == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := wt.app
	xl.cores.Lock(wt.num)
	return wt
}

func (wt *workTitle) Select() error {
	xl := wt.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, wt.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (wt *workTitle) Delete() error {
	xl := wt.app

	cmd := "Method"
	name := "Delete"

	_, err := xl.cores.SendNum(cmd, name, wt.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (wt *workTitle) Text(value ...string) string {
	xl := wt.app

	name := "Text"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, wt.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wt.num, nil)
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
