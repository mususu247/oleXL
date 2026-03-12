package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type chartObjects struct {
	app    *Excel
	parent any
	num    int
}

type chartObject struct {
	app    *Excel
	parent any
	num    int
}

func (ws *workSheet) ChartObjects() *chartObjects {
	var cos chartObjects
	xl := ws.app

	kind := "ChartObjects"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		name := "ChartObjects"
		ans, err := xl.cores.SendNum(cmd, name, ws.num, nil)
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
	cos.app = xl
	cos.num = num
	cos.parent = ws
	return &cos
}

func (cos *chartObjects) Release() error {
	xl := cos.app
	xl.cores.Release(cos.num, false)
	return nil
}

func (cos *chartObjects) Nothing() error {
	xl := cos.app
	xl.cores.releaseChild(cos.num)

	xl.cores.Unlock(cos.num)
	err := cos.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(cos.num)
	cos = nil
	return nil
}

func (cos *chartObjects) Set() *chartObjects {
	if cos == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := cos.app
	xl.cores.Lock(cos.num)
	return cos
}

func (cos *chartObjects) ChartObjectz(value any) *chartObject {
	var co chartObject
	xl := cos.app

	kind := "ChartObject"
	core, num := xl.cores.FindAdd(kind, cos.num)
	if core.disp == nil {
		cmd := "Get"
		name := "ChartObjects"
		var opt []any
		switch x := value.(type) {
		case int:
			if x > 0 {
				opt = append(opt, int32(x))
			}
		case int32:
			if x > 0 {
				opt = append(opt, x)
			}
		case string:
			opt = append(opt, x)
		}

		ans, err := xl.cores.SendNum(cmd, name, cos.num, opt)
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
	co.app = xl
	co.num = num
	co.parent = cos
	return &co
}

func (co *chartObject) Release() error {
	xl := co.app
	xl.cores.Release(co.num, false)
	return nil
}

func (co *chartObject) Nothing() error {
	xl := co.app
	xl.cores.releaseChild(co.num)

	xl.cores.Unlock(co.num)
	err := co.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(co.num)
	co = nil
	return nil
}

func (co *chartObject) Set() *chartObject {
	if co == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := co.app
	xl.cores.Lock(co.num)
	return co
}

func (co *chartObject) Activate() error {
	xl := co.app

	cmd := "Method"
	name := "Activate"

	_, err := xl.cores.SendNum(cmd, name, co.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (co *chartObject) Select() error {
	xl := co.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, co.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (co *chartObject) Chart() *workChart {
	var ct workChart
	xl := co.app

	kind := "Chart"
	core, num := xl.cores.FindAdd(kind, co.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Chart"
		ans, err := xl.cores.SendNum(cmd, name, co.num, nil)
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
	ct.app = xl
	ct.num = num
	ct.parent = co
	return &ct
}

func (co *chartObject) Copy() error {
	xl := co.app

	cmd := "Method"
	name := "Copy"

	_, err := xl.cores.SendNum(cmd, name, co.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (co *chartObject) Cut() error {
	xl := co.app

	cmd := "Method"
	name := "Cut"

	_, err := xl.cores.SendNum(cmd, name, co.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (co *chartObject) Delete() error {
	xl := co.app

	cmd := "Method"
	name := "Delete"

	_, err := xl.cores.SendNum(cmd, name, co.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (co *chartObject) Duplicate() *chartObject {
	var xo chartObject
	xl := co.app

	kind := "ChartObject"
	core, num := xl.cores.FindAdd(kind, co.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Duplicate"

		ans, err := xl.cores.SendNum(cmd, name, co.num, nil)
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
	xo.app = xl
	xo.num = num
	xo.parent = co.parent
	return &xo
}
