package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workNotes struct {
	app    *Excel
	parent any
	num    int
}

type workNote struct {
	app    *Excel
	parent any
	num    int
}

func (ws *workSheet) Comments() *workNotes {
	var nts workNotes
	xl := ws.app

	name := "Comments"
	core, num := xl.cores.FindAdd(name, ws.num)
	if core.disp == nil {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, ws.num, nil)
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

	nts.app = xl
	nts.num = num
	nts.parent = ws
	return &nts
}

func (ws *workSheet) Commentz(value any) *workNote {
	var nt workNote
	xl := ws.app

	kind := "Comment"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Comments"
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

		ans, err := xl.cores.SendNum(cmd, name, ws.num, opt)
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
	nt.app = xl
	nt.num = num
	nt.parent = ws
	return &nt
}

func (nts *workNotes) Release() error {
	xl := nts.app
	return xl.cores.Release(nts.num, false)
}

func (nts *workNotes) Nothing() error {
	xl := nts.app
	xl.cores.releaseChild(nts.num)

	xl.cores.Unlock(nts.num)
	err := nts.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(nts.num)
	nts = nil
	return nil
}

func (nts *workNotes) Count() int32 {
	var result int32
	xl := nts.app

	cmd := "Get"
	name := "Count"
	ans, err := xl.cores.SendNum(cmd, name, nts.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return result
	}

	switch x := ans.(type) {
	case int32:
		result = x
	}
	return result
}

func (rg *workRange) Comment() *workNote {
	var nt workNote
	xl := rg.app

	name := "Comment"
	core, num := xl.cores.FindAdd(name, rg.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, rg.num, nil)
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
	nt.app = xl
	nt.num = num
	nt.parent = rg
	return &nt
}

func (rg *workRange) AddComment(text string) *workNote {
	var nt workNote
	xl := rg.app

	name := "AddComment"
	core, num := xl.cores.FindAdd(name, rg.num)
	if core.disp == nil {
		cmd := "Method"
		var opt []any
		opt = append(opt, text)

		ans, err := xl.cores.SendNum(cmd, name, rg.num, opt)
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
	nt.app = xl
	nt.num = num
	nt.parent = rg
	return &nt
}

func (nt *workNote) Release() error {
	xl := nt.app
	return xl.cores.Release(nt.num, false)
}

func (nt *workNote) Nothing() error {
	xl := nt.app
	xl.cores.releaseChild(nt.num)

	xl.cores.Unlock(nt.num)
	err := nt.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(nt.num)
	nt = nil
	return nil
}

func (nt *workNote) Visible(value bool) error {
	xl := nt.app
	cmd := "Put"
	name := "Visible"
	var opt []any
	opt = append(opt, value)

	_, err := xl.cores.SendNum(cmd, name, nt.num, opt)
	if err != nil {
		log.Printf("(Error) cmd:%v name:%v %v", cmd, name, value)
	}
	return nil
}

func (nt *workNote) Delete() error {
	xl := nt.app

	name := "Delete"
	cmd := "Method"

	_, err := xl.cores.SendNum(cmd, name, nt.num, nil)
	if err != nil {
		return err
	}

	return nil
}

func (nt *workNote) Text(value ...any) string {
	xl := nt.app

	name := "Text"
	if len(value) > 0 {
		cmd := "Method"
		var opt []any
		if len(value) > 0 {
			for _, v := range value {
				opt = append(opt, v)
			}
		}

		_, err := xl.cores.SendNum(cmd, name, nt.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Method"

		ans, err := xl.cores.SendNum(cmd, name, nt.num, nil)
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
