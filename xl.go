package oleXL

import (
	"fmt"
	"log"
	"runtime/debug"
	"time"

	"github.com/go-ole/go-ole"
)

type Excel struct {
	cores *Cores
	num   int
}

func (Q *Excel) Init(debug ...bool) error {
	xl := Q
	Q.getVersion()
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

func (Q *Excel) CreateObject() error {
	xl := Q
	name := "Excel.Application"
	cmd := "Create"
	core, num := xl.cores.Add(name, 0)
	core.lock = 1 //Lock on
	ans, err := xl.cores.SendNum(cmd, name, num, nil)
	if err != nil {
		return fmt.Errorf("(Error) CreateObject:%v\n", err)
	}
	switch x := ans.(type) {
	case *ole.IDispatch:
		if x != nil {
			core.disp = x
			core.lock = 1 //Lock on
		} else {
			return nil
		}
	case nil:
		return fmt.Errorf("(Error) CreateObject\n")
	}

	xl.num = num
	return nil
}

func (Q *Excel) Quit() error {
	xl := Q
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

func (Q *Excel) Nothing() error {
	xl := Q
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

func (Q *Excel) Hand() int32 {
	xl := Q
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

func (Q *Excel) Visible(value bool) error {
	xl := Q
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

func (Q *Excel) DisplayAlerts(value ...bool) bool {
	xl := Q
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

func (Q *Excel) ScreenUpdating(value ...bool) bool {
	xl := Q
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

func (Q *Excel) Calculation(value ...any) int32 {
	xl := Q
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
		default:
			z = SetEnumCalculation(0)
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

func (Q *Excel) getVersion() {
	info, _ := debug.ReadBuildInfo()

	targetModule := "github.com/go-ole/go-ole"
	for _, dep := range info.Deps {
		if dep.Path == targetModule {
			fmt.Printf("Module .Name: %s .Version: %s\n", dep.Path, dep.Version)
		}
	}

	targetModule = "github.com/mususu247/oleXL"
	for _, dep := range info.Deps {
		if dep.Path == targetModule {
			fmt.Printf("Module .Name: %s .Version: %s\n", dep.Path, dep.Version)
		}
	}
}
