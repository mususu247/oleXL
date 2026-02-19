package oleXL

// version 2026-01-26

import (
	"log"
	"time"

	"github.com/go-ole/go-ole"
)

type Excel struct {
	cores *Cores
	num   int
}

func (xl *Excel) Init(debug ...bool) error {
	var cores Cores
	if len(debug) > 0 {
		cores.Init(debug[0])
	} else {
		cores.Init(false)
	}

	xl.num = 1
	xl.cores = &cores
	xl.cores.Start()
	return nil
}

func (xl *Excel) CreateObject() error {
	cmd := "Create"
	name := "Excel.Application"

	core, num := xl.cores.Add(name, 0)
	core.lock = 1 //lockOn
	ans, err := xl.cores.SendNum(cmd, name, num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	switch x := ans.(type) {
	case *ole.IDispatch:
		core.disp = x
		core.lock = 1 //lockOn
	}

	xl.num = num
	return nil
}

func (xl *Excel) Quit() error {
	count := xl.Workbooks().Count()
	for i := count; i > 0; i-- {
		xl.Workbookz(i).Close(false)
	}
	xl.cores.releaseChild(xl.num)

	cmd := "Method"
	name := "Quit"
	ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
	if err != nil {
		return err
	}

	switch x := ans.(type) {
	case nil:
		//ok
	default:
		log.Printf("Quit(): %v", x)
	}

	xl.cores.Release(xl.num, true)
	return nil
}

func (xl *Excel) Nothing() error {
	xl.cores.Release(xl.num, true)

	err := xl.cores.worker.Stop()
	if err != nil {
		log.Printf("%v", err)
		return err
	}

	for {
		if xl.cores.worker.IsOpened() {
			time.Sleep(1 * time.Millisecond)
		} else {
			if xl.cores.debug {
				log.Printf("worker.IsOpened: false")
			}
			time.Sleep(1 * time.Millisecond)
			break
		}
	}

	return nil
}

func (xl *Excel) Hand() int32 {
	cmd := "Get"
	name := "hWnd"

	ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
	if err != nil {
		log.Printf("(Error) .Hand:%v", err)
		return -1 //err
	}
	switch x := ans.(type) {
	case int32:
		return x
	}
	return -1 //err
}

func (xl *Excel) Visible(value bool) error {
	cmd := "Put"
	name := "Visible"
	var opt []any
	opt = append(opt, value)

	_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
	if err != nil {
		log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
	}
	return nil
}

func (xl *Excel) DisplayAlerts(value ...bool) bool {
	var opt []any

	name := "DisplayAlerts"
	if len(value) > 0 {
		//Set
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	}

	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
	if err != nil {
		log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
	}

	switch x := ans.(type) {
	case bool:
		return x
	}
	return false
}

func (xl *Excel) ScreenUpdating(value ...bool) bool {
	var opt []any

	name := "ScreenUpdating"
	if len(value) > 0 {
		//Set
		cmd := "Put"
		opt = append(opt, value[0])

		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	}

	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
	if err != nil {
		log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
	}

	switch x := ans.(type) {
	case bool:
		return x
	}
	return false
}

func (xl *Excel) Calculation(value ...any) int32 {
	var opt []any

	name := "Calculation"
	if len(value) > 0 {
		//Set
		cmd := "Put"

		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumCalculation(int32(x))
		case int32:
			z = SetEnumCalculation(x)
		case string:
			z = GetEnumCalculationNum(x)
		}
		opt = append(opt, z)

		_, err := xl.cores.SendNum(cmd, name, xl.num, opt)
		if err != nil {
			log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
		}
	}

	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, xl.num, nil)
	if err != nil {
		log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
	}

	switch x := ans.(type) {
	case int32:
		return x
	}
	return -1
}
