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
	m["Visible"] = true
	m["MacroType"] = int32(-4142)
	return m
}

func (Q *workNames) Count() int32 {
	xl := Q.app

	cmd := "Get"
	name := "Count"
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

func (Q *workNames) Add(value map[string]any) *workName {
	var body workName
	xl := Q.app

	name := "Add"
	core, num := xl.cores.FindAdd(name, Q.num)
	if core.disp == nil {
		cmd := "Method"

		var opt []any
		for range 4 {
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
			case "Visible":
				switch x := v.(type) {
				case bool:
					opt[2] = x
				}
			case "MacroType":
				switch x := v.(type) {
				case int32:
					opt[3] = x
				}
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

func (Q *workBook) Namez(value any) *workName {
	var body workName
	xl := Q.app
	parent := Q.Names()

	name := "Item"
	core, num := xl.cores.FindAdd(name, parent.num)
	if core.disp == nil {
		cmd := "Method"

		var opt []any
		switch x := value.(type) {
		case int:
			opt = append(opt, int32(x))
		case int32:
			opt = append(opt, x)
		case string:
			opt = append(opt, x)
		}

		ans, err := xl.cores.SendNum(cmd, name, parent.num, opt)
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

func (Q *workName) Name(value ...string) string {
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

func (Q *workName) RefersTo(value ...string) string {
	xl := Q.app

	name := "RefersTo"
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

func (Q *workName) Visible(value ...bool) bool {
	xl := Q.app

	name := "Visible"
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

func (Q *workName) MacroType(value ...any) int32 {
	xl := Q.app

	name := "MacroType"
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

func (Q *workName) Value(value ...string) string {
	xl := Q.app

	name := "Value"
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

func (Q *workName) RefersToRange() *workRange {
	var body workRange
	xl := Q.app

	name := "RefersToRange"
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

func (Q *workName) Index() int32 {
	xl := Q.app

	name := "Index"
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

func (Q *workName) Comment(value ...string) string {
	xl := Q.app

	name := "Comment"
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

func (Q *workName) WorkbookParameter() bool {
	xl := Q.app

	name := "WorkbookParameter"
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
	return false
}

func (Q *workName) ValidWorkbookParameter() bool {
	xl := Q.app

	name := "ValidWorkbookParameter"
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
	return false
}
