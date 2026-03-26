package oleXL

import (
	"fmt"
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

func (Q *chartObject) Chart() *workChart {
	var body workChart
	xl := Q.app

	kind := "Chart"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Chart"
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

func (Q *workChart) Release() error {
	xl := Q.app
	xl.cores.Release(Q.num, false)
	return nil
}

func (Q *workChart) Nothing() error {
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

func (Q *workChart) Set() (*workChart, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workChart) SetSourceData(Source *workRange, RowCol ...any) error {
	xl := Q.app

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

	_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
	if err != nil {
		log.Printf("(Error) %v", err)
		return err
	}

	return nil
}

func (xl *Excel) ActiveChart() *workChart {
	var body workChart
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
	body.parent = ws
	ws.Release()
	return &body
}

func (Q *workChart) Select() error {
	xl := Q.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workChart) HasTitle(value ...bool) bool {
	xl := Q.app

	name := "HasTitle"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
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
	}
	return false
}

func (Q *workChart) HasLegend(value ...bool) bool {
	xl := Q.app

	name := "HasLegend"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
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
	}
	return false
}

func (Q *workChart) Parent() *chartObject {
	var body chartObject
	xl := Q.app

	kind := "Chart"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Parent"
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

func (Q *workChart) Position(value ...any) int32 {
	xl := Q.app

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

func (Q *workChart) Name(value ...string) string {
	xl := Q.app

	name := "Name"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

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

func (Q *workChart) SetElement(value any) error {
	xl := Q.app

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

	_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workChart) ChartGroups() *chartGroups {
	var body chartGroups
	xl := Q.app

	name := "ChartGroups"
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

func (Q *workChart) ChartGroupz(value int32) *workChartGroup {
	var body workChartGroup
	xl := Q.app

	name := "ChartGroups"
	core, num := xl.cores.FindAdd(name, Q.num)
	if core.disp == nil {
		cmd := "Method"
		var opt []any
		opt = append(opt, value)

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

func (Q *workChart) Location(value any, option ...string) error {
	xl := Q.app

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

	_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workChart) ChartType(value ...any) int32 {
	xl := Q.app

	name := "ChartType"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any

		var v int32
		switch x := value[0].(type) {
		case int:
			v = SetEnumChartType(int32(x))
		case int32:
			v = SetEnumChartType(x)
		case string:
			v = GetEnumChartTypeNum(x)
		}
		opt = append(opt, v)

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

func (Q *workChart) PlotBy(value ...any) int32 {
	xl := Q.app

	name := "PlotBy"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any

		var v int32
		switch x := value[0].(type) {
		case int:
			v = SetEnumRowCol(int32(x))
		case int32:
			v = SetEnumRowCol(x)
		case string:
			v = GetEnumRowColNum(x)
		}
		opt = append(opt, v)

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

func (Q *chartGroups) Count() int32 {
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

func (Q *chartGroups) AxisGroup(value ...any) int32 {
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
