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

func (ct *workChart) SeriesCollection() *seriesCollection {
	var sc seriesCollection
	xl := ct.app

	name := "SeriesCollection"
	core, num := xl.cores.FindAdd(name, ct.num)
	if core.disp == nil {
		cmd := "Method"

		ans, err := xl.cores.SendNum(cmd, name, ct.num, nil)
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
	sc.app = xl
	sc.num = num
	sc.parent = ct
	return &sc
}

func (ct *workChart) FullSeriesCollection() *seriesCollection {
	var sc seriesCollection
	xl := ct.app

	name := "FullSeriesCollection"
	core, num := xl.cores.FindAdd(name, ct.num)
	if core.disp == nil {
		cmd := "Method"

		ans, err := xl.cores.SendNum(cmd, name, ct.num, nil)
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
	sc.app = xl
	sc.num = num
	sc.parent = ct
	return &sc
}

func (sc *seriesCollection) Release() error {
	xl := sc.app
	xl.cores.Release(sc.num, false)
	return nil
}

func (sc *seriesCollection) Nothing() error {
	xl := sc.app
	xl.cores.releaseChild(sc.num)

	xl.cores.Unlock(sc.num)
	err := sc.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(sc.num)
	sc = nil
	return nil
}

func (sc *seriesCollection) Set() (*seriesCollection, error) {
	if sc == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := sc.app
	xl.cores.Lock(sc.num)
	return sc, nil
}

func (sc *seriesCollection) Count() int32 {
	xl := sc.app
	name := "Count"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, sc.num, nil)
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

func (sc *seriesCollection) Item(value any) *workSeries {
	var ws workSeries
	xl := sc.app

	kind := "Series"
	core, num := xl.cores.FindAdd(kind, sc.num)
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

		ans, err := xl.cores.SendNum(cmd, name, sc.num, opt)
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
	ws.app = xl
	ws.num = num
	ws.parent = sc.parent //wrokChart
	return &ws
}

func (sc *seriesCollection) Extend(value *workRange, option ...any) *workSeries {
	var ws workSeries
	xl := sc.app

	kind := "Series"
	core, num := xl.cores.FindAdd(kind, sc.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Extend"
		var opt []any
		core := xl.cores.getCore(sc.num)
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

		ans, err := xl.cores.SendNum(cmd, name, sc.num, opt)
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
	ws.app = xl
	ws.num = num
	ws.parent = sc.parent //wrokChart
	return &ws
}

func (ws *workSeries) Release() error {
	xl := ws.app
	xl.cores.Release(ws.num, false)
	return nil
}

func (ws *workSeries) Nothing() error {
	xl := ws.app
	xl.cores.releaseChild(ws.num)

	xl.cores.Unlock(ws.num)
	err := ws.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(ws.num)
	ws = nil
	return nil
}

func (ws *workSeries) Set() (*workSeries, error) {
	if ws == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := ws.app
	xl.cores.Lock(ws.num)
	return ws, nil
}

func (ws *workSeries) Select() error {
	if ws == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := ws.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, ws.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (ws *workSeries) AxisGroup(value ...any) string {
	xl := ws.app

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
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ws.num, nil)
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

func (ws *workSeries) Delete() error {
	xl := ws.app

	name := "Delete"
	cmd := "Method"
	_, err := xl.cores.SendNum(cmd, name, ws.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return err
	}
	return nil
}

func (ws *workSeries) Name(value ...any) string {
	xl := ws.app

	name := "Name"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case string:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ws.num, nil)
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

func (ws *workSeries) MarkerStyle(value ...any) string {
	xl := ws.app

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
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ws.num, nil)
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

func (ws *workSeries) MarkerSize(value ...int32) int32 {
	xl := ws.app

	name := "MarkerSize"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, ws.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, ws.num, nil)
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
