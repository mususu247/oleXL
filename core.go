package oleXL

// version 2026-01-26

import (
	"fmt"
	"log"
	"time"

	"github.com/go-ole/go-ole"
)

type Core struct {
	parent int
	kind   string
	disp   *ole.IDispatch
	lock   int
}

type Cores struct {
	count  int
	worker Worker
	list   map[int]*Core
	debug  bool
}

func (cs *Cores) Init(debug bool) error {
	cs.list = make(map[int]*Core)
	cs.debug = debug
	return nil
}

func (cs *Cores) Start() error {
	cs.worker.parent = cs
	err := cs.worker.Start()
	if err != nil {
		return err
	}

	return nil
}

func (cs *Cores) Stop() error {
	err := cs.worker.Stop()
	if err != nil {
		return err
	}

	return nil
}

func (cs *Cores) Add(kind string, parent int) (*Core, int) {
	var core Core
	var pos int

	core.parent = parent
	core.kind = kind
	core.disp = nil
	core.lock = -1

	cs.count++
	pos = cs.count
	cs.list[pos] = &core

	return &core, pos
}

func (cs *Cores) FindAdd(kind string, parent int) (*Core, int) {
	for i := range cs.list {
		if cs.list[i].parent == parent {
			if cs.list[i].kind == kind {
				if cs.list[i].disp == nil {
					return cs.list[i], i
				}
			}
		}
	}
	return cs.Add(kind, parent)
}

func (cs *Cores) getCore(num int) *Core {
	if _, ok := cs.list[num]; ok {
		return cs.list[num]
	}
	return nil
}

func (cs *Cores) SendDisp(cmd, name string, disp *ole.IDispatch, opt []any) (any, error) {
	var ans any

	args := cs.worker.Send(cmd, disp, name, opt)
	for i := range args {
		switch x := args[i].(type) {
		case error:
			if cs.debug {
				log.Printf("err %v.%v: %v", cmd, name, x)
			}
			return nil, x
		case *ole.IDispatch:
			if cs.debug {
				log.Printf("ans (object) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case int16: //vba integer
			if cs.debug {
				log.Printf("ans (int16) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case int32: //vba long
			if cs.debug {
				log.Printf("ans (int32) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case int64: //vba longlong
			if cs.debug {
				log.Printf("ans (int64) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case float32: //single
			if cs.debug {
				log.Printf("ans (float32) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case float64: //vba diuble
			if cs.debug {
				log.Printf("ans (float64) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case time.Time: //vba date
			if cs.debug {
				log.Printf("ans (date.time) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case string: //vba string
			if cs.debug {
				log.Printf("ans (string) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case bool: //vba boolen
			if cs.debug {
				log.Printf("ans (bool) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case nil: //vba null
			if cs.debug {
				log.Printf("ans (nil) %v.%v: %v", cmd, name, x)
			}
			ans = x
		default:
			if cs.debug {
				log.Printf("def %v.%v: %v", cmd, name, x)
			}
			ans = x
		}
	}

	return ans, nil
}

func (cs *Cores) SendNum(cmd, name string, num int, opt []any) (any, error) {
	var ans any
	core := cs.getCore(num)
	if core == nil {
		return nil, fmt.Errorf("not found.getCore: %v\n", num)
	}

	args := cs.worker.Send(cmd, core.disp, name, opt)
	for i := range args {
		switch x := args[i].(type) {
		case error:
			if cs.debug {
				log.Printf("err %v.%v: %v", cmd, name, x)
			}
			return nil, x
		case *ole.IDispatch:
			if cs.debug {
				log.Printf("ans (object) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case int16: //vba integer
			if cs.debug {
				log.Printf("ans (int16) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case int32: //vba long
			if cs.debug {
				log.Printf("ans (int32) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case int64: //vba longlong
			if cs.debug {
				log.Printf("ans (int64) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case float32: //single
			if cs.debug {
				log.Printf("ans (float32) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case float64: //vba diuble
			if cs.debug {
				log.Printf("ans (float64) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case time.Time: //vba date
			if cs.debug {
				log.Printf("ans (date.time) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case string: //vba string
			if cs.debug {
				log.Printf("ans (string) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case bool: //vba boolen
			if cs.debug {
				log.Printf("ans (bool) %v.%v: %v", cmd, name, x)
			}
			ans = x
		case nil: //vba null
			if cs.debug {
				log.Printf("ans (nil) %v.%v: %v", cmd, name, x)
			}
			ans = x
		default:
			if cs.debug {
				log.Printf("def %v.%v: %v", cmd, name, x)
			}
			ans = x
		}
		cs.Release(num, false)
	}

	return ans, nil
}

func (cs *Cores) Lock(num int) error {
	if _, ok := cs.list[num]; ok {
		if cs.list[num].disp != nil {
			cs.list[num].lock = 1
		} else {
			cs.list[num].lock = -1
		}
	}
	return nil
}

func (cs *Cores) Unlock(num int) error {
	if _, ok := cs.list[num]; ok {
		if cs.list[num].disp != nil {
			cs.list[num].lock = 0
		} else {
			cs.list[num].lock = -1
		}
	}
	return nil
}

func (cs *Cores) Release(num int, lockOn bool) error {
	const cmd = "Release"
	core := cs.getCore(num)
	if core == nil {
		return fmt.Errorf("not found. core.num:%v", num)
	}
	if core.disp == nil {
		return nil
	}
	if !lockOn {
		if core.lock > 0 {
			return fmt.Errorf("is Locked kind:%v", core.kind)
		}
	}

	if core.disp != nil {
		cmd := "Release"
		name := core.kind
		ans, err := cs.SendDisp(cmd, name, core.disp, nil)
		if err != nil {
			return err
		}
		if cs.debug {
			log.Printf("(Release) kind:%v num:%v", core.kind, num)
		}
		switch x := ans.(type) {
		case int32:
			_ = x
		}
		core.disp = nil
		core.lock = -1
	}
	return nil
}

func (cs *Cores) Remove(num int) {
	delete(cs.list, num)
}

func (cs *Cores) releaseChild(num int) error {
	for i := range cs.list {
		if _, ok := cs.list[i]; ok {
			if cs.list[i].parent == num {
				cs.releaseChild(i)
				if cs.list[i].disp != nil {
					cs.Release(i, true)
				}
				delete(cs.list, i)
			}
		}
	}
	return nil
}
