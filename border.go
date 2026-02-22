package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workBorder struct {
	app *Excel
	num int
}

func (wr *workRange) Borders(value ...any) *workBorder {
	var br workBorder
	xl := wr.app

	name := "Borders"
	core, num := xl.cores.FindAdd(name, wr.num)
	if core.disp == nil {
		cmd := "Get"
		var opt []any
		var z int32

		if len(value) > 0 {
			switch x := value[0].(type) {
			case int:
				z = SetEnumBorders(int32(x))
			case int32:
				z = SetEnumBorders(x)
			case string:
				z = GetEnumBordersNum(x)
			}
			opt = append(opt, z)
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
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
	br.app = xl
	br.num = num
	return &br
}

func (wr *workRange) BorderAround(LineStyle any, Weight any, ColorIndex any, Color any, ThemeColor any) bool {
	xl := wr.app

	name := "BorderAround"
	cmd := "Method"
	var opt []any
	var z int32

	switch x := LineStyle.(type) {
	case int:
		z = SetEnumLineStyle(int32(x))
		opt = append(opt, z)
	case int32:
		z = SetEnumLineStyle(x)
		opt = append(opt, z)
	case string:
		z = GetEnumLineStyleNum(x)
		opt = append(opt, z)
	default:
		opt = append(opt, nil)
	}

	switch x := Weight.(type) {
	case int:
		z = SetEnumWeight(int32(x))
		opt = append(opt, z)
	case int32:
		z = SetEnumWeight(x)
		opt = append(opt, z)
	case string:
		z = GetEnumWeightNum(x)
		opt = append(opt, z)
	default:
		opt = append(opt, nil)
	}

	switch x := ColorIndex.(type) {
	case int:
		opt = append(opt, int32(x))
	case int32:
		opt = append(opt, x)
	default:
		opt = append(opt, nil)
	}

	switch x := Color.(type) {
	case float64:
		opt = append(opt, x)
	case string:
		f64 := GetEnumRgbColorNum(x)
		opt = append(opt, f64)
	default:
		opt = append(opt, nil)
	}

	switch x := ThemeColor.(type) {
	case int:
		z = SetEnumThemeColor(int32(x))
		opt = append(opt, z)
	case int32:
		z = SetEnumThemeColor(x)
		opt = append(opt, z)
	case string:
		z := GetEnumShapeTypeNum(x)
		opt = append(opt, z)
	default:
		opt = append(opt, nil)
	}

	ans, err := xl.cores.SendNum(cmd, name, wr.num, opt)
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

func (br *workBorder) Release() error {
	xl := br.app
	return xl.cores.Release(br.num, false)
}

func (br *workBorder) Nothing() error {
	xl := br.app
	xl.cores.releaseChild(br.num)

	xl.cores.Unlock(br.num)
	err := br.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(br.num)
	br = nil
	return nil
}

func (br *workBorder) Set() *workBorder {
	if br == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := br.app
	xl.cores.Lock(br.num)
	return br
}

func (br *workBorder) LineStyle(value ...any) int32 {
	xl := br.app

	name := "LineStyle"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumLineStyle(int32(x))
		case int32:
			z = SetEnumLineStyle(x)
		case string:
			z = GetEnumLineStyleNum(x)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, br.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, br.num, nil)
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

func (br *workBorder) Weight(value ...any) int32 {
	xl := br.app

	name := "Weight"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumWeight(int32(x))
		case int32:
			z = SetEnumWeight(x)
		case string:
			z = GetEnumWeightNum(x)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, br.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, br.num, nil)
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
