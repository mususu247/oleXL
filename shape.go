package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type workShapes struct {
	app    *Excel
	parent any
	num    int
}

type workShape struct {
	app    *Excel
	parent any
	num    int
}

func (Q *workNote) Shape() *workShape {
	var body workShape
	xl := Q.app

	kind := "Shape"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Shape"
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

func (Q *workSheet) Shapes() *workShapes {
	var body workShapes
	xl := Q.app

	kind := "Shapes"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Shapes"
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

func (Q *workSheet) Shapez(value any) *workShape {
	var body workShape
	xl := Q.app
	sps := Q.Shapes()

	kind := "Shape"
	core, num := xl.cores.FindAdd(kind, Q.num)
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

func (Q *workShapes) Release() error {
	xl := Q.app
	return xl.cores.Release(Q.num, false)
}

func (Q *workShapes) Nothing() error {
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

func (Q *workShapes) Count() int32 {
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

func (Q *workShapes) AddShape(Type any, left, top, width, height float64) *workShape {
	var body workShape
	xl := Q.app

	kind := "Shape"
	core, num := xl.cores.FindAdd(kind, Q.num)
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
		default:
			z = SetEnumShapeType(0)
		}
		opt = append(opt, z)
		opt = append(opt, left)
		opt = append(opt, top)
		opt = append(opt, width)
		opt = append(opt, height)

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
	body.parent = Q.parent
	return &body
}

func (Q *workShape) Release() error {
	xl := Q.app
	xl.cores.Release(Q.num, false)
	return nil
}

func (Q *workShape) Nothing() error {
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

func (Q *workShape) Set() (*workShape, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workShape) Select() error {
	xl := Q.app

	cmd := "Method"
	name := "Select"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *workShape) Name(value ...any) string {
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

func (Q *workShape) Delete() error {
	xl := Q.app

	name := "Delete"
	cmd := "Method"
	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return err
	}
	return nil
}

func (Q *workShapes) AddChart2(style int32, ChartType any, option ...any) *workShape {
	var body workShape
	xl := Q.app

	//style int32, ChartType any, left, top, width, height float64, newLayout bool

	kind := "Shape"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Method"
		name := "AddChart2"
		var opt []any
		for range 6 {
			opt = append(opt, nil)
		}

		opt[0] = style

		var z int32
		switch x := ChartType.(type) {
		case int:
			z = SetEnumChartType(int32(x))
		case int32:
			z = SetEnumChartType(x)
		case string:
			z = GetEnumChartTypeNum(x)
		default:
			z = SetEnumChartType(0)
		}
		opt[1] = z

		for i := range option {
			switch x := option[i].(type) {
			case int:
				opt[i+2] = float64(x)
			case float64:
				opt[i+2] = x
			case bool:
				opt[5] = x
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

func (Q *workShapes) AddPicture(fileName string, LinkToFile bool, SaveWithDocument bool, left, top, width, height float64) *workShape {
	var body workShape
	xl := Q.app

	if !FileExists(fileName) {
		log.Printf("(Error) not found: %v", fileName)
		return nil
	}

	kind := "Shape"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Method"
		name := "AddPicture"
		var opt []any
		opt = append(opt, fileName)

		var z int32
		z = 0 //msoFalse:0
		if LinkToFile {
			z = -1 //msoTrue:-1
		}
		opt = append(opt, z)

		z = 0 //msoFalse:0
		if SaveWithDocument {
			z = -1 //msoTrue:-1
		}
		opt = append(opt, z)

		opt = append(opt, left)
		opt = append(opt, top)
		opt = append(opt, width)
		opt = append(opt, height)

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
	body.parent = Q.parent
	return &body
}
