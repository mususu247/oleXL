package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type workFilter struct {
	app    *Excel
	parent any
	num    int
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
