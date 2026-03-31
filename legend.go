package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type workLegend struct {
	app    *Excel
	parent any
	num    int
}

func (Q *workChart) Legend() *workLegend {
	var body workLegend
	xl := Q.app

	name := "Legend"
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

func (Q *workLegend) Release() error {
	xl := Q.app
	xl.cores.Release(Q.num, false)
	return nil
}

func (Q *workLegend) Nothing() error {
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

func (Q *workLegend) Set() (*workLegend, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workLegend) Select() error {
	xl := Q.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workLegend) Delete() error {
	xl := Q.app

	cmd := "Method"
	name := "Delete"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workLegend) Clear() error {
	xl := Q.app

	cmd := "Method"
	name := "Clear"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workLegend) Position(value ...any) int32 {
	var opt []any
	xl := Q.app

	name := "Position"
	if len(value) > 0 {
		cmd := "Put"

		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumLegendPosition(int32(x))
		case int32:
			z = SetEnumLegendPosition(x)
		case string:
			z = GetEnumLegendPositionNum(x)
		default:
			z = SetEnumLegendPosition(0)
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
