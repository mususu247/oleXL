package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type workNames struct {
	app    *Excel
	parent any
	num    int
}

type workName struct {
	app    *Excel
	parent any
	num    int
}

func (Q *workBook) Names() *workNames {
	var body workNames
	xl := Q.app

	name := "Names"
	core, num := xl.cores.FindAdd(name, Q.num)
	if core.disp == nil {
		cmd := "Get"
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

func (Q *workNames) Release() error {
	xl := Q.app
	return xl.cores.Release(Q.num, false)
}

func (Q *workNames) Nothing() error {
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

func (Q *workNames) Set() (*workNames, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workNames) GetAddOptions() map[string]any {
	var m map[string]any
	m = make(map[string]any)
	m["Name"] = nil
	m["RefersTo"] = nil
	m["Visible"] = nil
	m["MacroType"] = nil
	m["ShortcutKey"] = nil
	m["Category"] = nil
	m["NameLocal"] = nil
	m["RefersToLocal"] = nil
	m["CategoryLocal"] = nil
	m["RefersToR1C1"] = nil
	m["RefersToR1C1Local"] = nil
	return m
}

func (Q *workNames) Add(value map[string]any) *workName {
	var body workName
	xl := Q.app

	name := "Add"
	core, num := xl.cores.FindAdd(name, Q.num)
	if core.disp == nil {
		cmd := "Method"

		var opt []any
		for range 2 {
			opt = append(opt, nil)
		}

		for k, v := range value {
			switch k {
			case "Name":
				switch x := v.(type) {
				case string:
					opt[0] = x
				}
			case "RefersTo":
				switch x := v.(type) {
				case string:
					opt[1] = x
				}
				//case "Visible":
				//case "MacroType":
				//case "ShortcutKey":
				//case "Category":
				//case "NameLocal":
				//case "RefersToLocal":
				//case "CategoryLocal":
				//case "RefersToR1C1":
				//case "RefersToR1C1Local":
			}
		}

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

func (Q *workName) Release() error {
	xl := Q.app
	return xl.cores.Release(Q.num, false)
}

func (Q *workName) Nothing() error {
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

func (Q *workName) Set() (*workName, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workNames) Namez(value any) *workName {
	var body workName
	xl := Q.app

	name := "Names"
	core, num := xl.cores.FindAdd(name, Q.num)
	if core.disp == nil {
		cmd := "Get"

		var opt []any
		switch x := value.(type) {
		case int:
			opt = append(opt, int32(x))
		case int32:
			opt = append(opt, x)
		case string:
			opt = append(opt, x)
		}

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

func (Q *workName) Name(value ...any) string {
	xl := Q.app

	name := "Name"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case string:
			opt = append(opt, x)
		}

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

func (Q *workName) RefersTo(value ...any) string {
	xl := Q.app

	name := "RefersTo"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case string:
			opt = append(opt, x)
		}

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

func (Q *workName) RefersToR1C1(value ...any) string {
	xl := Q.app

	name := "RefersToR1C1"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case string:
			opt = append(opt, x)
		}

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

func (Q *workName) Value() string {
	xl := Q.app

	cmd := "Get"
	name := "Value"
	ans, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return ""
	}
	switch x := ans.(type) {
	case string:
		return x
	}
	return ""
}

func (Q *workName) Visible(value ...any) bool {
	xl := Q.app

	name := "Visible"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case bool:
			opt = append(opt, x)
		}

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

func (Q *workName) Delete() error {
	xl := Q.app
	cmd := "Method"
	name := "Delete"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		log.Printf("(Error) cmd:%v name:%v ", cmd, name)
	}
	return nil
}
