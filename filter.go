package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type autoFilter struct {
	app    *Excel
	parent any
	num    int
}

type workFilter struct {
	app    *Excel
	parent any
	num    int
}

type workFilters struct {
	app    *Excel
	parent any
	num    int
}

func (Q *workSheet) AutoFilterMode() bool {
	xl := Q.app

	name := "AutoFilterMode"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return false
	}
	switch x := ans.(type) {
	case bool:
		return x
	}

	return false
}

func (Q *workSheet) AutoFilter() *autoFilter {
	var body autoFilter
	xl := Q.app

	if !Q.AutoFilterMode() {
		log.Printf("(Error) AutoFilter is null")
		return nil
	}

	kind := "AutoFilter"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "AutoFilter"

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
				log.Printf("(Error) AutoFilter is null")
				return nil
			}
		}
	}
	body.app = xl
	body.num = num
	body.parent = Q
	return &body
}

func (Q *workRange) AutoFilter(Field int32, Criteria1 string, Operator any, Criteria2 string, SubField int32, VisibleDropDown bool) *workFilter {
	var body workFilter
	xl := Q.app

	kind := "AutoFilter"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "AutoFilter"

		var opt []any
		opt = append(opt, Field)
		opt = append(opt, Criteria1)

		var z int32
		switch x := Operator.(type) {
		case int:
			z = SetEnumAutoFilterOperator(int32(x))
		case int32:
			z = SetEnumAutoFilterOperator(x)
		case string:
			z = GetEnumAutoFilterOperatorNum(x)
		default:
			z = SetEnumAutoFilterOperator(0)
		}
		opt = append(opt, z)

		if len(Criteria2) > 0 {
			opt = append(opt, Criteria2)

			if SubField > 0 {
				opt = append(opt, SubField)
			} else {
				opt = append(opt, nil)
			}
			opt = append(opt, VisibleDropDown)
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
	body.parent = Q
	return &body
}

func (Q *autoFilter) Release() error {
	xl := Q.app
	return xl.cores.Release(Q.num, false)
}

func (Q *autoFilter) Nothing() error {
	if Q == nil {
		return fmt.Errorf("(Error) Object is NULL.")
	}
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

func (Q *autoFilter) Set() (*autoFilter, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *autoFilter) Filters() *workFilters {
	var body workFilters
	xl := Q.app

	kind := "Filters"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Filters"

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

func (Q *autoFilter) ShowAllData() error {
	xl := Q.app
	cmd := "Method"
	name := "ShowAllData"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		log.Printf("(Error) cmd:%v name:%v", cmd, name)
	}
	return nil
}

func (Q *workFilters) Release() error {
	xl := Q.app
	return xl.cores.Release(Q.num, false)
}

func (Q *workFilters) Nothing() error {
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

func (Q *workFilters) Set() (*workFilters, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workFilters) Count() int32 {
	xl := Q.app

	cmd := "Get"
	name := "Count"
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

func (Q *autoFilter) Filterz(value int32) *workFilter {
	var body workFilter
	xl := Q.app

	Qf, _ := Q.Filters().Set()
	kind := "Filter"
	core, num := xl.cores.FindAdd(kind, Qf.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Item"

		ans, err := xl.cores.SendNum(cmd, name, Qf.num, nil)
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

func (Q *workFilter) Release() error {
	xl := Q.app
	return xl.cores.Release(Q.num, false)
}

func (Q *workFilter) Nothing() error {
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

func (Q *workFilter) Set() (*workFilter, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workFilter) On() bool {
	xl := Q.app

	name := "On"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return false
	}
	switch x := ans.(type) {
	case bool:
		return x
	}

	return false
}
