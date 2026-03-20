package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type chartObjects struct {
	app    *Excel
	parent any
	num    int
}

type chartObject struct {
	app    *Excel
	parent any
	num    int
}

func (ws *workSheet) ChartObjects() *chartObjects {
	var cos chartObjects
	xl := ws.app

	kind := "ChartObjects"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Method"
		name := "ChartObjects"
		ans, err := xl.cores.SendNum(cmd, name, ws.num, nil)
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
	cos.app = xl
	cos.num = num
	cos.parent = ws
	return &cos
}

func (ws *workSheet) ChartObjectz(value any) *chartObject {
	return ws.ChartObjects().Item(value)
}

func (cos *chartObjects) Release() error {
	xl := cos.app
	xl.cores.Release(cos.num, false)
	return nil
}

func (cos *chartObjects) Nothing() error {
	xl := cos.app
	xl.cores.releaseChild(cos.num)

	xl.cores.Unlock(cos.num)
	err := cos.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(cos.num)
	cos = nil
	return nil
}

func (cos *chartObjects) Set() (*chartObjects, error) {
	if cos == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := cos.app
	xl.cores.Lock(cos.num)
	return cos, nil
}

func (cos *chartObjects) Select() error {
	xl := cos.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, cos.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (cos *chartObjects) Item(value any) *chartObject {
	var co chartObject
	xl := cos.app

	kind := "ChartObject"
	core, num := xl.cores.FindAdd(kind, cos.num)
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

		ans, err := xl.cores.SendNum(cmd, name, cos.num, opt)
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
	co.app = xl
	co.num = num
	co.parent = cos
	return &co
}

func (co *chartObject) Release() error {
	xl := co.app
	xl.cores.Release(co.num, false)
	return nil
}

func (co *chartObject) Nothing() error {
	xl := co.app
	xl.cores.releaseChild(co.num)

	xl.cores.Unlock(co.num)
	err := co.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(co.num)
	co = nil
	return nil
}

func (co *chartObject) Set() (*chartObject, error) {
	if co == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := co.app
	xl.cores.Lock(co.num)
	return co, nil
}

func (co *chartObject) Select() error {
	xl := co.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, co.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (co *chartObject) Activate() error {
	xl := co.app

	cmd := "Method"
	name := "Activate"

	_, err := xl.cores.SendNum(cmd, name, co.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (cos *chartObjects) Count() int32 {
	xl := cos.app

	name := "Count"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, cos.num, nil)
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

func (co *chartObject) Copy() error {
	xl := co.app

	cmd := "Method"
	name := "Copy"

	_, err := xl.cores.SendNum(cmd, name, co.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (co *chartObject) Cut() error {
	xl := co.app

	cmd := "Method"
	name := "Cut"

	_, err := xl.cores.SendNum(cmd, name, co.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (co *chartObject) Delete() error {
	xl := co.app

	cmd := "Method"
	name := "Delete"

	_, err := xl.cores.SendNum(cmd, name, co.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (co *chartObject) Duplicate() *chartObject {
	var xo chartObject
	xl := co.app

	kind := "ChartObject"
	core, num := xl.cores.FindAdd(kind, co.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Duplicate"

		ans, err := xl.cores.SendNum(cmd, name, co.num, nil)
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
	xo.app = xl
	xo.num = num
	xo.parent = co.parent
	return &xo
}

func (co *chartObject) Name(value ...string) string {
	xl := co.app

	name := "Name"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, co.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, co.num, nil)
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

func (co *chartObject) Left(value ...float64) float64 {
	xl := co.app

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

func (co *chartObject) Top(value ...float64) float64 {
	xl := co.app

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

func (co *chartObject) Width(value ...float64) float64 {
	xl := co.app

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

func (co *chartObject) Height(value ...float64) float64 {
	xl := co.app

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
