package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workAxes struct {
	app    *Excel
	parent any
	num    int
}

func (ct *workChart) HasAxis(AxisType any, AxisGroup any, value ...any) bool {
	var opt []any
	xl := ct.app

	var z int32
	// set AxisType
	switch x := AxisType.(type) {
	case int:
		z = SetEnumAxisType(int32(x))
	case int32:
		z = SetEnumAxisType(x)
	case string:
		z = GetEnumAxisTypeNum(x)
	}
	opt = append(opt, z)

	// AxisGroup
	switch x := AxisGroup.(type) {
	case int:
		z = SetEnumAxisGroup(int32(x))
	case int32:
		z = SetEnumAxisGroup(x)
	case string:
		z = GetEnumAxisGroupNum(x)
	}
	opt = append(opt, z)

	name := "HasAxis"
	if len(value) > 0 {
		cmd := "Put"

		// HasAxis
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, ct.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ct.num, nil)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}

		switch x := ans.(type) {
		case bool:
			return x
		}
	}

	return false
}

func (ct *workChart) Axes(value ...any) *workAxes {
	var ax workAxes
	xl := ct.app

	name := "Axes"
	core, num := xl.cores.FindAdd(name, ct.num)
	if core.disp == nil {
		cmd := "Method"
		var opt []any
		if len(value) > 0 {
			var z int32
			switch x := value[0].(type) {
			case int:
				z = SetEnumAxisType(int32(x))
			case int32:
				z = SetEnumAxisType(x)
			case string:
				z = GetEnumAxisTypeNum(x)
			}
			opt = append(opt, z)
		}
		if len(value) > 1 {
			var z int32
			switch x := value[1].(type) {
			case int:
				z = SetEnumAxisGroup(int32(x))
			case int32:
				z = SetEnumAxisGroup(x)
			case string:
				z = GetEnumAxisGroupNum(x)
			}
			opt = append(opt, z)
		} else {
			opt = nil
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
	ax.app = xl
	ax.num = num
	ax.parent = ct
	return &ax
}

func (ax *workAxes) Release() error {
	xl := ax.app
	xl.cores.Release(ax.num, false)
	return nil
}

func (ax *workAxes) Nothing() error {
	xl := ax.app
	xl.cores.releaseChild(ax.num)

	xl.cores.Unlock(ax.num)
	err := ax.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(ax.num)
	ax = nil
	return nil
}

func (ax *workAxes) Set() *workAxes {
	if ax == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := ax.app
	xl.cores.Lock(ax.num)
	return ax
}

func (ax *workAxes) HasMajorGridlines(value ...bool) bool {
	var opt []any
	xl := ax.app

	name := "HasMajorGridlines"
	if len(value) > 0 {
		//Set
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, ax.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ax.num, nil)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}

		switch x := ans.(type) {
		case bool:
			return x
		}
	}

	return false
}

func (ax *workAxes) HasMinorGridlines(value ...bool) bool {
	var opt []any
	xl := ax.app

	name := "HasMinorGridlines"
	if len(value) > 0 {
		//Set
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, ax.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ax.num, nil)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}

		switch x := ans.(type) {
		case bool:
			return x
		}
	}

	return false
}

func (ax *workAxes) HasTitle(value ...bool) bool {
	var opt []any
	xl := ax.app

	name := "HasTitle"
	if len(value) > 0 {
		//Set
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, ax.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ax.num, nil)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}

		switch x := ans.(type) {
		case bool:
			return x
		}
	}

	return false
}

func (ax *workAxes) TickLabelPosition(value ...any) int32 {
	var opt []any
	xl := ax.app

	name := "TickLabelPosition"
	if len(value) > 0 {
		//Set
		cmd := "Put"

		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumTickLabelPosition(int32(x))
		case int32:
			z = SetEnumTickLabelPosition(x)
		case string:
			z = GetEnumTickLabelPositionNum(x)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, ax.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ax.num, nil)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}

		switch x := ans.(type) {
		case int32:
			return x
		}
	}
	return 0
}

func (ax *workAxes) AxisTitle() *workTitle {
	var wt workTitle
	xl := ax.app

	kind := "AxisTitle"
	core, num := xl.cores.FindAdd(kind, ax.num)
	if core.disp == nil {
		cmd := "Get"
		name := "AxisTitle"

		ans, err := xl.cores.SendNum(cmd, name, ax.num, nil)
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
	wt.parent = ax
	return &wt
}

func (ax *workAxes) MinimumScale(value ...float64) float64 {
	var opt []any
	xl := ax.app

	name := "MinimumScale"
	if len(value) > 0 {
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, ax.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ax.num, nil)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}

		switch x := ans.(type) {
		case float64:
			return x
		}
	}

	return 0
}

func (ax *workAxes) MaximumScale(value ...float64) float64 {
	var opt []any
	xl := ax.app

	name := "MaximumScale"
	if len(value) > 0 {
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, ax.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ax.num, nil)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}

		switch x := ans.(type) {
		case float64:
			return x
		}
	}

	return 0
}

func (ax *workAxes) MinimumScaleIsAuto(value ...bool) bool {
	var opt []any
	xl := ax.app

	name := "MinimumScaleIsAuto"
	if len(value) > 0 {
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, ax.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ax.num, nil)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}

		switch x := ans.(type) {
		case bool:
			return x
		}
	}

	return false
}

func (ax *workAxes) MaximumScaleIsAuto(value ...bool) bool {
	var opt []any
	xl := ax.app

	name := "MaximumScaleIsAuto"
	if len(value) > 0 {
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, ax.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ax.num, nil)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}

		switch x := ans.(type) {
		case bool:
			return x
		}
	}

	return false
}
