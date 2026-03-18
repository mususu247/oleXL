package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type chartGroups struct {
	app    *Excel
	parent any
	num    int
}

type workChartGroup struct {
	app    *Excel
	parent any
	num    int
}

type workChart struct {
	app    *Excel
	parent any
	num    int
}

func (co *chartObject) Chart() *workChart {
	var ct workChart
	xl := co.app

	kind := "Chart"
	core, num := xl.cores.FindAdd(kind, co.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Chart"
		ans, err := xl.cores.SendNum(cmd, name, co.num, nil)
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
	ct.app = xl
	ct.num = num
	ct.parent = co
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
			core.lock = 1 //Lock on
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

func (ct *workChart) HasLegend(value ...bool) bool {
	xl := ct.app

	name := "HasLegend"
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

func (ct *workChart) Position(value ...any) int32 {
	xl := ct.app

	name := "Position"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any

		var v int32
		switch x := value[0].(type) {
		case int:
			v = SetEnumLegendPosition(int32(x))
		case int32:
			v = SetEnumLegendPosition(x)
		case string:
			v = GetEnumLegendPositionNum(x)
		}
		opt = append(opt, v)

		_, err := xl.cores.SendNum(cmd, name, ct.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, ct.num, nil)
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

func (ct *workChart) Name(value ...string) string {
	xl := ct.app

	name := "Name"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, ct.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, ct.num, nil)
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

func (ct *workChart) SetElement(value any) error {
	xl := ct.app

	cmd := "Method"
	name := "SetElement"
	var opt []any
	var z int32
	switch x := value.(type) {
	case int:
		z = SetEnumChartElementType(int32(x))
	case int32:
		z = SetEnumChartElementType(x)
	case string:
		z = GetEnumChartElementTypeNum(x)
	}
	opt = append(opt, z)

	_, err := xl.cores.SendNum(cmd, name, ct.num, opt)
	if err != nil {
		return err
	}
	return nil
}

func (ct *workChart) ChartGroups() *chartGroups {
	var cgs chartGroups
	xl := ct.app

	name := "ChartGroups"
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
			core.disp = x
			core.lock = 0
		}
	}
	cgs.app = xl
	cgs.num = num
	cgs.parent = ct
	return &cgs
}

func (cgs *chartGroups) Count() int32 {
	xl := cgs.app

	name := "Count"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, cgs.num, nil)
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

func (ct *workChart) ChartGroupz(value int32) *workChartGroup {
	var cg workChartGroup
	xl := ct.app

	name := "ChartGroups"
	core, num := xl.cores.FindAdd(name, ct.num)
	if core.disp == nil {
		cmd := "Method"
		var opt []any
		opt = append(opt, value)

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
	cg.app = xl
	cg.num = num
	cg.parent = ct
	return &cg
}

func (ct *workChart) Location(value any, option ...string) error {
	xl := ct.app

	cmd := "Method"
	name := "Location"
	var opt []any
	var z int32
	switch x := value.(type) {
	case int:
		z = SetEnumChartLocation(int32(x))
	case int32:
		z = SetEnumChartLocation(x)
	case string:
		z = GetEnumChartLocationNum(x)
	}
	opt = append(opt, z)

	if len(option) > 0 {
		opt = append(opt, option[0])
	}

	_, err := xl.cores.SendNum(cmd, name, ct.num, opt)
	if err != nil {
		return err
	}
	return nil
}
