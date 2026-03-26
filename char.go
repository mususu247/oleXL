package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workChar struct {
	app    *Excel
	parent any
	num    int
}

func (Q *workFrame) Characterz(value ...any) *workChar {
	var body workChar
	xl := Q.app

	name := "Characters"
	core, num := xl.cores.FindAdd(name, Q.num)
	if core.disp == nil {
		cmd := "Method"
		var opt []any
		if len(value) > 0 {
			for i := range value {
				switch x := value[i].(type) {
				case int:
					opt = append(opt, int32(x))
				case int32:
					opt = append(opt, x)
				}
			}
		} else {
			opt = nil
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

func (Q *workFrame) Characters() *workChar {
	var body workChar
	xl := Q.app

	name := "Characters"
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

func (Q *workTextRange) Characterz(value ...any) *workChar {
	var body workChar
	xl := Q.app

	name := "Characters"
	core, num := xl.cores.FindAdd(name, Q.num)
	if core.disp == nil {
		cmd := "Method"
		var opt []any
		if len(value) > 0 {
			for i := range value {
				switch x := value[i].(type) {
				case int:
					opt = append(opt, int32(x))
				case int32:
					opt = append(opt, x)
				}
			}
		} else {
			opt = nil
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

func (Q *workTextRange) Characters() *workChar {
	var body workChar
	xl := Q.app

	name := "Characters"
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

func (Q *workChar) Release() error {
	xl := Q.app
	return xl.cores.Release(Q.num, false)
}

func (Q *workChar) Nothing() error {
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

func (Q *workChar) Text(value ...string) string {
	xl := Q.app

	name := "Text"
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

func (Q *workChar) Count() int32 {
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

func (Q *workChar) Insert(value string) error {
	xl := Q.app

	name := "Insert"
	cmd := "Method"
	var opt []any
	opt = append(opt, value)
	_, err := xl.cores.SendNum(cmd, name, Q.num, opt)
	if err != nil {
		log.Printf("(Error) %v", err)
		return err
	}
	return nil
}
