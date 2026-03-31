package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type workAxes struct {
	app    *Excel
	parent any
	num    int
}

func (Q *workChart) HasAxis(AxisType any, AxisGroup any, value ...any) bool {
	var opt []any
	xl := Q.app

	var z int32
	// set AxisType
	switch x := AxisType.(type) {
	case int:
		z = SetEnumAxisType(int32(x))
	case int32:
		z = SetEnumAxisType(x)
	case string:
		z = GetEnumAxisTypeNum(x)
	default:
		z = SetEnumAxisType(0)
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
	default:
		z = SetEnumAxisGroup(0)
	}
	opt = append(opt, z)

	name := "HasAxis"
	if len(value) > 0 {
		cmd := "Put"

		// HasAxis
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workChart) Axes(value ...any) *workAxes {
	var body workAxes
	xl := Q.app

	name := "Axes"
	core, num := xl.cores.FindAdd(name, Q.num)
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
			default:
				z = SetEnumAxisType(0)
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
			default:
				z = SetEnumAxisGroup(0)
			}
			opt = append(opt, z)
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			if x != nil {
				core.disp = x
				core.lock = 1 //Lock on
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

func (Q *workAxes) Release() error {
	xl := Q.app
	xl.cores.Release(Q.num, false)
	return nil
}

func (Q *workAxes) Nothing() error {
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

func (Q *workAxes) Set() (*workAxes, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workAxes) Select() error {
	xl := Q.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workAxes) HasMajorGridlines(value ...bool) bool {
	var opt []any
	xl := Q.app

	name := "HasMajorGridlines"
	if len(value) > 0 {
		//Set
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workAxes) HasMinorGridlines(value ...bool) bool {
	var opt []any
	xl := Q.app

	name := "HasMinorGridlines"
	if len(value) > 0 {
		//Set
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workAxes) HasTitle(value ...bool) bool {
	var opt []any
	xl := Q.app

	name := "HasTitle"
	if len(value) > 0 {
		//Set
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workAxes) TickLabelPosition(value ...any) int32 {
	var opt []any
	xl := Q.app

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
		default:
			z = SetEnumTickLabelPosition(0)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workAxes) AxisTitle() *workTitle {
	var body workTitle
	xl := Q.app

	kind := "AxisTitle"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "AxisTitle"
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

func (Q *workAxes) MinimumScale(value ...float64) float64 {
	var opt []any
	xl := Q.app

	name := "MinimumScale"
	if len(value) > 0 {
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workAxes) MaximumScale(value ...float64) float64 {
	var opt []any
	xl := Q.app

	name := "MaximumScale"
	if len(value) > 0 {
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workAxes) MinimumScaleIsAuto(value ...bool) bool {
	var opt []any
	xl := Q.app

	name := "MinimumScaleIsAuto"
	if len(value) > 0 {
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workAxes) MaximumScaleIsAuto(value ...bool) bool {
	var opt []any
	xl := Q.app

	name := "MaximumScaleIsAuto"
	if len(value) > 0 {
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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
