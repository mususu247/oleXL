package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workArea struct {
	app    *Excel
	parent any
	num    int
}

func (ct *workChart) ChartArea() *workArea {
	var wa workArea
	xl := ct.app

	name := "ChartArea"
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
			core.lock = 1 //Lock.on
		}
	}
	wa.app = xl
	wa.num = num
	wa.parent = ct
	return &wa
}

func (wa *workArea) Release() error {
	xl := wa.app
	return xl.cores.Release(wa.num, false)
}

func (wa *workArea) Nothing() error {
	xl := wa.app
	xl.cores.releaseChild(wa.num)

	xl.cores.Unlock(wa.num)
	err := wa.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(wa.num)
	wa = nil
	return nil
}

func (wa *workArea) Set() *workArea {
	if wa == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := wa.app
	xl.cores.Lock(wa.num)
	return wa
}

func (wa *workArea) Copy() error {
	xl := wa.app

	kind := "ChartArea"
	core, _ := xl.cores.FindAdd(kind, wa.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Copy"

		_, err := xl.cores.SendNum(cmd, name, wa.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return nil
		}
	}

	return nil
}

func (wa *workArea) Select() error {
	if wa == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := wa.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, wa.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (wa *workArea) Name(value ...any) string {
	xl := wa.app

	name := "Name"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case string:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, wa.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wa.num, nil)
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

func (wa *workArea) Left(value ...float64) float64 {
	xl := wa.app

	name := "Left"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case float64:
			return x
		}
	}
	return 0
}

func (wa *workArea) Top(value ...float64) float64 {
	xl := wa.app

	name := "Top"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case float64:
			return x
		}
	}
	return 0
}

func (wa *workArea) Width(value ...float64) float64 {
	xl := wa.app

	name := "Width"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case float64:
			return x
		}
	}
	return 0
}

func (wa *workArea) Height(value ...float64) float64 {
	xl := wa.app

	name := "Height"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
		switch x := ans.(type) {
		case float64:
			return x
		}
	}
	return 0
}
