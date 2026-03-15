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

func (sps *workShapes) AddChart2(style int32, ChartType any, left, top, width, height float64, newLayout bool) *chartObject {
	var co chartObject
	xl := sps.app

	kind := "Shape"
	core, num := xl.cores.FindAdd(kind, sps.num)
	if core.disp == nil {
		cmd := "Method"
		name := "AddChart2"
		var opt []any

		var z int32
		opt = append(opt, style)

		switch x := ChartType.(type) {
		case int:
			z = SetEnumChartType(int32(x))
		case int32:
			z = SetEnumChartType(x)
		case string:
			z = GetEnumChartTypeNum(x)
		}
		opt = append(opt, z)

		opt = append(opt, left)
		opt = append(opt, top)
		opt = append(opt, width)
		opt = append(opt, height)
		opt = append(opt, newLayout)

		ans, err := xl.cores.SendNum(cmd, name, sps.num, opt)
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
	co.parent = sps
	return &co
}

func (ct *workChart) ChartObjects() *chartObjects {
	var cos chartObjects
	xl := ct.app

	kind := "ChartObjects"
	core, num := xl.cores.FindAdd(kind, ct.num)
	if core.disp == nil {
		cmd := "Method"
		name := "ChartObjects"
		ans, err := xl.cores.SendNum(cmd, name, ct.num, nil)
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
	cos.parent = ct
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

func (cos *chartObject) Select() error {
	xl := cos.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, cos.num, nil)
	if err != nil {
		return err
	}
	return nil
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

func (ct *workChart) ChartObjectz(value any) *chartObject {
	var co chartObject
	xl := ct.app

	kind := "ChartObject"
	core, num := xl.cores.FindAdd(kind, ct.num)
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

		ans, err := xl.cores.SendNum(cmd, name, ct.num, opt)
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
	co.parent = ct
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

func (cos *chartObjects) Count() int32 {
	xl := cos.app

	name := "Count"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, cos.num, nil)
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

func (co *chartObject) Name(value ...string) string {
	xl := co.app

	name := "Name"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, co.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, co.num, nil)
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
