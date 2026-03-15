package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workLegend struct {
	app    *Excel
	parent any
	num    int
}

func (ct *workChart) Legend() *workLegend {
	var lg workLegend
	xl := ct.app

	name := "Legend"
	core, num := xl.cores.FindAdd(name, ct.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, ct.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
		switch x := ans.(type) {
		case *ole.IDispatch:
			core.disp = x
			core.lock = 1 //Lock on
		}
	}
	lg.app = xl
	lg.num = num
	lg.parent = ct
	return &lg
}

func (lg *workLegend) Release() error {
	xl := lg.app
	xl.cores.Release(lg.num, false)
	return nil
}

func (lg *workLegend) Nothing() error {
	xl := lg.app
	xl.cores.releaseChild(lg.num)

	xl.cores.Unlock(lg.num)
	err := lg.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(lg.num)
	lg = nil
	return nil
}

func (lg *workLegend) Set() *workLegend {
	if lg == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := lg.app
	xl.cores.Lock(lg.num)
	return lg
}

func (lg *workLegend) Select() error {
	xl := lg.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, lg.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (lg *workLegend) Delete() error {
	xl := lg.app

	cmd := "Method"
	name := "Delete"

	_, err := xl.cores.SendNum(cmd, name, lg.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (lg *workLegend) Clear() error {
	xl := lg.app

	cmd := "Method"
	name := "Clear"

	_, err := xl.cores.SendNum(cmd, name, lg.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (lg *workLegend) Position(value ...any) int32 {
	var opt []any
	xl := lg.app

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
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, lg.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, lg.num, nil)
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
