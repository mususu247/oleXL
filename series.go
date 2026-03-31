package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type workSeries struct {
	app    *Excel
	parent any
	num    int
}

type seriesCollection struct {
	app    *Excel
	parent any
	num    int
}

func (Q *workChart) SeriesCollection() *seriesCollection {
	var body seriesCollection
	xl := Q.app

	name := "SeriesCollection"
	core, num := xl.cores.FindAdd(name, Q.num)
	if core.disp == nil {
		cmd := "Method"
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

func (Q *workChart) FullSeriesCollection() *seriesCollection {
	var body seriesCollection
	xl := Q.app

	name := "FullSeriesCollection"
	core, num := xl.cores.FindAdd(name, Q.num)
	if core.disp == nil {
		cmd := "Method"
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

func (Q *seriesCollection) Release() error {
	xl := Q.app
	xl.cores.Release(Q.num, false)
	return nil
}

func (Q *seriesCollection) Nothing() error {
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

func (Q *seriesCollection) Set() (*seriesCollection, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *seriesCollection) Count() int32 {
	xl := Q.app
	name := "Count"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *seriesCollection) Item(value any) *workSeries {
	var body workSeries
	xl := Q.app

	kind := "Series"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Item"
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

		ans, err := xl.cores.SendNum(cmd, name, Q.num, opt)
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
	body.parent = Q.parent //wrokChart
	return &body
}

func (Q *seriesCollection) Extend(value *workRange, option ...any) *workSeries {
	var body workSeries
	xl := Q.app

	kind := "Series"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Extend"
		var opt []any
		core := xl.cores.getCore(Q.num)
		opt = append(opt, core.disp)

		if len(option) > 0 {
			var z int32
			switch x := option[0].(type) {
			case int:
				z = SetEnumRowCol(int32(x))
			case int32:
				z = SetEnumRowCol(x)
			case string:
				z = GetEnumRowColNum(x)
			default:
				z = SetEnumRowCol(0)
			}
			opt = append(opt, z)
			opt = append(opt, true)
		}

		if len(option) > 1 {
			switch x := option[1].(type) {
			case bool:
				opt = append(opt, x)
			}
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
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = Q.parent //wrokChart
	return &body
}

func (Q *seriesCollection) NewSeries() *workSeries {
	var body workSeries
	xl := Q.app

	kind := "Series"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Method"
		name := "NewSeries"
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
	body.parent = Q.parent //wrokChart
	return &body
}

func (Q *seriesCollection) Add(Source *workRange, option ...any) *workSeries {
	var body workSeries
	xl := Q.app

	kind := "Series"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Add"
		var opt []any
		for range 6 {
			opt = append(opt, true)
		}

		core := xl.cores.getCore(Source.num)
		opt[0] = core.disp

		for i := range option {
			switch i {
			case 0:
				var z int32
				switch x := option[i].(type) {
				case int:
					z = SetEnumRowCol(int32(x))
				case int32:
					z = SetEnumRowCol(x)
				case string:
					z = GetEnumRowColNum(x)
				default:
					z = SetEnumRowCol(0)
				}
				opt[1] = z
			default:
				switch x := option[i].(type) {
				case bool:
					opt[i+1] = x
				}
			}
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
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = Q.parent //wrokChart
	return &body
}

func (Q *workSeries) Release() error {
	xl := Q.app
	xl.cores.Release(Q.num, false)
	return nil
}

func (Q *workSeries) Nothing() error {
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

func (Q *workSeries) Set() (*workSeries, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workSeries) Select() error {
	if Q == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := Q.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workSeries) AxisGroup(value ...any) int32 {
	xl := Q.app

	name := "AxisGroup"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumAxisGroup(int32(x))
		case int32:
			z = SetEnumAxisGroup(x)
		case string:
			z = GetEnumAlignCmdNum(x)
		default:
			z = SetEnumAxisGroup(0)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case int32:
			return x
		}
	}
	return 0
}

func (Q *workSeries) Delete() error {
	xl := Q.app

	name := "Delete"
	cmd := "Method"
	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return err
	}
	return nil
}

func (Q *workSeries) Name(value ...any) string {
	xl := Q.app

	name := "Name"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case string:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workSeries) MarkerStyle(value ...any) string {
	xl := Q.app

	name := "MarkerStyle"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumMarkerStyle(int32(x))
		case int32:
			z = SetEnumMarkerStyle(x)
		case string:
			z = GetEnumMarkerStyleNum(x)
		default:
			z = SetEnumMarkerStyle(0)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
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

func (Q *workSeries) MarkerSize(value ...int32) int32 {
	xl := Q.app

	name := "MarkerSize"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case int32:
			return x
		}
	}

	return 0
}
