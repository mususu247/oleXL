package oleXL

import (
	"log"

	"github.com/go-ole/go-ole"
)

type workShapes struct {
	app    *Excel
	parent *workSheet
	num    int
}

type workShape struct {
	app    *Excel
	parent *workSheet
	num    int
}

func (ws *workSheet) Shapes() *workShapes {
	var sps workShapes
	xl := ws.app

	kind := "Shapes"
	core, num := xl.cores.FindAdd(kind, ws.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Shapes"

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
	sps.app = xl
	sps.num = num
	sps.parent = ws
	return &sps
}

func (ws *workSheet) Shapez(value any) *workShape {
	var sp workShape
	xl := ws.app
	sps := ws.Shapes()

	kind := "Shape"
	core, num := xl.cores.FindAdd(kind, ws.num)
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

		ans, err := xl.cores.SendNum(cmd, name, sps.num, opt)
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
	sp.app = xl
	sp.num = num
	sp.parent = ws
	return &sp
}

func (sps *workShapes) Release() error {
	xl := sps.app
	return xl.cores.Release(sps.num, false)
}

func (sps *workShapes) Nothing() error {
	xl := sps.app
	xl.cores.releaseChild(sps.num)

	xl.cores.Unlock(sps.num)
	err := sps.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(sps.num)
	sps = nil
	return nil
}

func (sps *workShapes) Count() int32 {
	xl := sps.app
	name := "Count"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, sps.num, nil)
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

func (sps *workShapes) AddShape(Type any, left, top, width, height float64) *workShape {
	var sp workShape
	xl := sps.app

	kind := "Shape"
	core, num := xl.cores.FindAdd(kind, sps.num)
	if core.disp == nil {
		cmd := "Method"
		name := "AddShape"
		var opt []any

		var z int32
		switch x := Type.(type) {
		case int:
			z = SetEnumShapeType(int32(x))
		case int32:
			z = SetEnumShapeType(x)
		case string:
			z = GetEnumShapeTypeNum(x)
		}
		opt = append(opt, z)

		opt = append(opt, left)
		opt = append(opt, top)
		opt = append(opt, width)
		opt = append(opt, height)
		ans, err := xl.cores.SendNum(cmd, name, sps.num, opt)
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
	sp.app = xl
	sp.num = num
	sp.parent = sps.parent
	return &sp
}

func (sp *workShape) Release() error {
	xl := sp.app
	xl.cores.Release(sp.num, false)
	return nil
}

func (sp *workShape) Nothing() error {
	xl := sp.app
	xl.cores.releaseChild(sp.num)

	xl.cores.Unlock(sp.num)
	err := sp.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(sp.num)
	sp = nil
	return nil
}

func (sp *workShape) Set() *workShape {
	if sp == nil {
		log.Printf("(Error) Object is NULL.")
		return nil
	}
	xl := sp.app
	xl.cores.Lock(sp.num)
	return sp
}

func (sp *workShape) Name(value ...any) string {
	xl := sp.app

	if len(value) > 0 {
		cmd := "Put"
		name := "Name"
		var opt []any
		switch x := value[0].(type) {
		case string:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, sp.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		name := "Name"
		ans, err := xl.cores.SendNum(cmd, name, sp.num, nil)
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
