package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workCharts struct {
	app    *Excel
	parent any
	num    int
}

type workChart struct {
	app    *Excel
	parent any
	num    int
}

func (sps *workShapes) AddChart2(style int32, ChartType any, left, top, width, height float64, newLayout bool) *workChart {
	var ct workChart
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
	ct.app = xl
	ct.num = num
	ct.parent = sps
	return &ct
}

func (ct *workChart) Release() error {
	xl := ct.app
	xl.cores.Release(ct.num, false)
	return nil
}

func (ct *workChart) Nothing() error {
	xl := ct.app
	xl.cores.releaseChild(ct.num)

	xl.cores.Unlock(ct.num)
	err := ct.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(ct.num)
	ct = nil
	return nil
}

func (ct *workChart) Set() *workChart {
	if ct == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := ct.app
	xl.cores.Lock(ct.num)
	return ct
}

func (ct *workChart) SetSourceData(Source *workRange, RowCol ...any) error {
	xl := ct.app

	name := "SetSourceData"
	cmd := "Method"
	var opt []any

	core := xl.cores.getCore(Source.num)
	opt = append(opt, core.disp)

	var z int32
	if len(RowCol) > 0 {
		switch x := RowCol[0].(type) {
		case int:
			z = SetEnumRowCol(int32(x))
		case int32:
			z = SetEnumRowCol(x)
		case string:
			z = GetEnumRowColNum(x)
		}
	} else {
		z = GetEnumRowColNum("xlRows")
	}
	opt = append(opt, z)

	_, err := xl.cores.SendNum(cmd, name, ct.num, opt)
	if err != nil {
		log.Printf("(Error) %v", err)
		return err
	}

	return nil
}

func (xl *Excel) ActiveChart() *workChart {
	var ct workChart
	ws := xl.ActiveSheet()

	kind := "Chart"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		name := "ActiveChart"
		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
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
	ct.parent = ws
	ws.Release()
	return &ct
}

func (ct *workChart) Select() error {
	xl := ct.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, ct.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (ct *workChart) HasTitle(value ...bool) bool {
	xl := ct.app

	name := "HasTitle"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, ct.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ct.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
		switch x := ans.(type) {
		case bool:
			return x
		}
	}
	return false
}

func (ct *workChart) Parent() *chartObject {
	var co chartObject
	xl := ct.app

	kind := "Chart"
	core, num := xl.cores.FindAdd(kind, ct.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Parent"
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
	co.app = xl
	co.num = num
	return &co
}
