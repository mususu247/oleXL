package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workFrame2 struct {
	app    *Excel
	parent any
	num    int
}

type workTextRange struct {
	app    *Excel
	parent any
	num    int
}

func (wf *workFormat) TextFrame2() *workFrame2 {
	var tf workFrame2
	xl := wf.app

	name := "TextFrame2"
	core, num := xl.cores.FindAdd(name, wf.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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
	tf.app = xl
	tf.num = num
	tf.parent = wf
	return &tf
}

func (sp *workShape) TextFrame2() *workFrame2 {
	var tf workFrame2
	xl := sp.app

	name := "TextFrame2"
	core, num := xl.cores.FindAdd(name, sp.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, sp.num, nil)
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
	tf.app = xl
	tf.num = num
	tf.parent = sp
	return &tf
}

func (wf *workFrame2) Release() error {
	xl := wf.app
	return xl.cores.Release(wf.num, false)
}

func (wf *workFrame2) Nothing() error {
	xl := wf.app
	xl.cores.releaseChild(wf.num)

	xl.cores.Unlock(wf.num)
	err := wf.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(wf.num)
	wf = nil
	return nil
}

func (wf *workFrame2) HasText() bool {
	xl := wf.app

	name := "HasText"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFrame2) DeleteText() bool {
	xl := wf.app

	name := "DeleteText"
	cmd := "Method"
	ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFrame2) MarginBottom(value ...int32) int32 {
	xl := wf.app

	name := "MarginBottom"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFrame2) MarginLeft(value ...int32) int32 {
	xl := wf.app

	name := "MarginLeft"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFrame2) MarginRight(value ...int32) int32 {
	xl := wf.app

	name := "MarginRight"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

func (wf *workFrame2) MarginTop(value ...int32) int32 {
	xl := wf.app

	name := "MarginTop"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, wf.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return 0
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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

// TextFrame2.TextRane
func (wf *workFrame2) TextRange() *workTextRange {
	var tr workTextRange
	xl := wf.app

	name := "TextRange"
	core, num := xl.cores.FindAdd(name, wf.num)
	if core.disp == nil {
		cmd := "Get"

		ans, err := xl.cores.SendNum(cmd, name, wf.num, nil)
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
	tr.app = xl
	tr.num = num
	tr.parent = wf
	return &tr
}

func (tr *workTextRange) Release() error {
	xl := tr.app
	return xl.cores.Release(tr.num, false)
}

func (tr *workTextRange) Nothing() error {
	xl := tr.app
	xl.cores.releaseChild(tr.num)

	xl.cores.Unlock(tr.num)
	err := tr.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(tr.num)
	tr = nil
	return nil
}

func (tr *workTextRange) Text(value ...string) string {
	xl := tr.app

	name := "Text"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		opt = append(opt, value[0])
		_, err := xl.cores.SendNum(cmd, name, tr.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, tr.num, nil)
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
