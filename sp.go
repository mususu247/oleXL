package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

// version 2026-01-05
// VBA style like

type workShapes struct {
	app *Excel
	mx  int
}

type workShape struct {
	app *Excel
	mx  int
}

type workLine struct {
	app *Excel
	mx  int
}

type workColor struct {
	app *Excel
	mx  int
}

func (sp *workShape) Nothing() error {
	xl := sp.app
	_, err := xl.getCore(sp.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	// wb.child.RelaseAll
	for k, v := range xl.WorkCores.cores {
		if v.px == sp.mx {
			xl.Release(k)
		}
	}
	return nil
}

func (ws *workSheet) Shapes() *workShapes {
	var sps workShapes
	sps.app = ws.app
	xl := ws.app

	sps.mx, _ = xl.findCore(ws.mx, "Shapes", 0)
	if sps.mx >= 0 {
		return &sps
	}

	_core, _ := xl.getCore(ws.mx)

	const cmd = "Get"
	const name = "Shapes"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			sps.mx, _ = xl.addCore(ws.mx, x, name, 0)
			sps.Count()
			return &sps
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (ws *workSheet) Shapez(value any) *workShape {
	var sp workShape
	xl := ws.app
	sps := ws.Shapes()
	sps.List()
	wsz := xl.findChild(sps.mx, "Shape")

	switch x := value.(type) {
	case int:
		for i := range wsz {
			j := wsz[i]
			if xl.WorkCores.cores[j].index == int32(x) {
				sp.app = sps.app
				sp.mx = j
				return &sp
			}
		}
	case int32:
		for i := range wsz {
			j := wsz[i]
			if xl.WorkCores.cores[j].index == int32(x) {
				sp.app = sps.app
				sp.mx = j
				return &sp
			}
		}
	case string:
		for i := range wsz {
			j := wsz[i]
			if v, ok := xl.WorkCores.cores[j].values["Name"]; ok {
				if v.(string) == x {
					sp.app = sps.app
					sp.mx = j
					return &sp
				}
			}
		}
	}
	return nil
}

func (sps *workShapes) Count(lock ...bool) int32 {
	xl := sps.app
	_core, err := xl.getCore(sps.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	const cmd = "Get"
	const name = "Count"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case int32:
			log.Printf("%v ans (int32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return -1
}

func (sps *workShapes) List() []*Core {
	var spz []*Core

	xl := sps.app
	_core, err := xl.getCore(sps.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}

	// delete.index = -1
	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == sps.mx {
			xl.WorkCores.cores[i].index = -1
		}
	}

	count := sps.Count()

	const cmd = "Method"
	const name = "Item"
	var opt []any
	opt = append(opt, int32(0))

	for j := int32(1); j <= count; j++ {
		opt[0] = j
		args := xl.worker.Send(cmd, _core.disp, name, opt)
		var sp workShape
		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case *ole.IDispatch:
				log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

				var _sp *Core
				sp.app = sps.app
				sp.mx, _sp = xl.addCore(sps.mx, x, "Shape", j)
				sp.ID()
				sp.Type()
				sp.Name()
				spz = append(spz, _sp)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	// delete.index = -1
	for i := range xl.WorkCores.cores {
		if xl.WorkCores.cores[i].px == sps.mx {
			if xl.WorkCores.cores[i].index == -1 {
				delete(xl.WorkCores.cores, i)
			}
		}
	}

	return spz
}

func (sp *workShape) ID() int32 {
	xl := sp.app
	_core, err := xl.getCore(sp.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	const cmd = "Get"
	const name = "ID"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case int32:
			log.Printf("%v ans (int32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return -1
}

func (sp *workShape) Type() int32 {
	xl := sp.app
	_core, err := xl.getCore(sp.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	const cmd = "Get"
	const name = "Type"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case int32:
			log.Printf("%v ans (int32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return -1
}

func (sp *workShape) Name(value ...string) string {
	xl := sp.app
	_core, err := xl.getCore(sp.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return ""
	}
	var cmd string
	const name = "Name"
	var opt []any

	if len(value) > 0 {
		cmd = "Put"
		opt = append(opt, value[0])
		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case nil:
				log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case string:
			log.Printf("%v ans (string) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return ""
}

func (sp *workShape) Left(value ...float32) float32 {
	xl := sp.app
	_core, err := xl.getCore(sp.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	var cmd string
	const name = "Left"
	var opt []any

	if len(value) > 0 {
		cmd = "Put"
		opt = append(opt, value[0])
		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case nil:
				log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case float32:
			log.Printf("%v ans (float32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return -1
}

func (sp *workShape) Top(value ...float32) float32 {
	xl := sp.app
	_core, err := xl.getCore(sp.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	var cmd string
	const name = "Top"
	var opt []any

	if len(value) > 0 {
		cmd = "Put"
		opt = append(opt, value[0])
		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case nil:
				log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case float32:
			log.Printf("%v ans (float32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return -1
}

func (sp *workShape) Width(value ...float32) float32 {
	xl := sp.app
	_core, err := xl.getCore(sp.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	var cmd string
	const name = "Width"
	var opt []any

	if len(value) > 0 {
		cmd = "Put"
		opt = append(opt, value[0])
		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case nil:
				log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case float32:
			log.Printf("%v ans (float32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return -1
}

func (sp *workShape) Height(value ...float32) float32 {
	xl := sp.app
	_core, err := xl.getCore(sp.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	var cmd string
	const name = "Height"
	var opt []any

	if len(value) > 0 {
		cmd = "Put"
		opt = append(opt, value[0])
		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case nil:
				log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case float32:
			log.Printf("%v ans (float32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return -1
}

func (sp *workShape) Delete() error {
	xl := sp.app
	_core, err := xl.getCore(sp.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	var sps workShapes
	sps.app = sp.app
	sps.mx = _core.px

	const cmd = "Method"
	const name = "Delete"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			return x
		case bool:
			log.Printf("%v ans (bool) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			if x {
				xl.Release(sp.mx)
				sps.List()
			}
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}

func (sp *workShape) Select() error {
	xl := sp.app
	_core, err := xl.getCore(sp.mx)
	if err != nil {
		return fmt.Errorf("(Error) %v", err)
	}
	var sps workShapes
	sps.app = sp.app
	sps.mx = _core.px

	const cmd = "Method"
	const name = "Select"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			return x
		case nil:
			log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return nil
}

func (sps *workShapes) AddShape(shapeStyle any, left, top, width, height float64) *workShape {
	var sp workShape
	sp.app = sps.app
	xl := sps.app
	_core, err := xl.getCore(sps.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	const cmd = "Method"
	const name = "AddShape"
	var opt []any

	var z int32
	switch x := shapeStyle.(type) {
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

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

			sp.mx, _ = xl.addCore(sps.mx, x, "Shape", 0)
			sp.ID()
			sp.Type()
			sp.Name()
			sps.List()
			return &sp
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return nil
}

func (sps *workShapes) AddPicture(fileName string, linkToFile, saveWithDocument bool, left, top, width, height float32) *workShape {
	var sp workShape
	sp.app = sps.app
	xl := sps.app
	_core, err := xl.getCore(sps.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	const cmd = "Method"
	const name = "AddPicture"
	var opt []any

	fn, err := GetAbsolutePathName(fileName)
	if err != nil {
		log.Printf("(Error) AddPicture file path error: %v", err)
		return nil
	}

	if !FileExists(fn) {
		log.Printf("(Error) AddPicture file not found: %v", fileName)
		return nil
	}

	opt = append(opt, fn)
	opt = append(opt, linkToFile)
	opt = append(opt, saveWithDocument)
	opt = append(opt, left)
	opt = append(opt, top)
	opt = append(opt, width)
	opt = append(opt, height)

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

			sp.mx, _ = xl.addCore(sps.mx, x, "Picture", 0)
			sp.ID()
			sp.Type()
			sp.Name()
			sps.List()
			return &sp
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return nil
}

func (sps *workShapes) AddChart2(style int32, chartStyle any, left, top, width, height float64) *workShape {
	var sp workShape
	sp.app = sps.app
	xl := sps.app
	_core, err := xl.getCore(sps.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return nil
	}
	const cmd = "Method"
	const name = "AddChart2"
	var opt []any

	var z int32
	switch x := chartStyle.(type) {
	case int:
		z = SetEnumChart(int32(x))
	case int32:
		z = SetEnumChart(x)
	case string:
		z = GetEnumChartNum(x)
	}
	opt = append(opt, style)
	opt = append(opt, z)
	opt = append(opt, left)
	opt = append(opt, top)
	opt = append(opt, width)
	opt = append(opt, height)

	args := xl.worker.Send(cmd, _core.disp, name, opt)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)

			sp.mx, _ = xl.addCore(sps.mx, x, "Shape", 0)
			sp.ID()
			sp.Type()
			sp.Name()
			sps.List()
			return &sp
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
		}
	}
	return nil
}

func (sp *workShape) Line() *workLine {
	var ln workLine
	ln.app = sp.app
	xl := sp.app

	ln.mx, _ = xl.findCore(sp.mx, "Line", 0)
	if ln.mx >= 0 {
		return &ln
	}

	_core, _ := xl.getCore(sp.mx)

	const cmd = "Get"
	const name = "Line"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			ln.mx, _ = xl.addCore(sp.mx, x, name, 0)
			return &ln
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (sp *workShape) Fill() *workLine {
	var ln workLine
	ln.app = sp.app
	xl := sp.app

	ln.mx, _ = xl.findCore(sp.mx, "Fill", 0)
	if ln.mx >= 0 {
		return &ln
	}

	_core, _ := xl.getCore(sp.mx)

	const cmd = "Get"
	const name = "Fill"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			ln.mx, _ = xl.addCore(sp.mx, x, name, 0)
			return &ln
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (ln *workLine) Weight(value ...float32) float32 {
	xl := ln.app
	_core, err := xl.getCore(ln.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	var cmd string
	const name = "Weight"
	var opt []any

	if len(value) > 0 {
		cmd = "Put"
		opt = append(opt, value[0])
		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case nil:
				log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case float32:
			log.Printf("%v ans (float32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return -1
}

func (ln *workLine) ForeColor() *workColor {
	var wc workColor
	wc.app = ln.app
	xl := ln.app

	wc.mx, _ = xl.findCore(ln.mx, "ForeColor", 0)
	if wc.mx >= 0 {
		return &wc
	}

	_core, _ := xl.getCore(ln.mx)

	const cmd = "Get"
	const name = "ForeColor"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			wc.mx, _ = xl.addCore(ln.mx, x, name, 0)
			return &wc
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (ln *workLine) BackColor() *workColor {
	var wc workColor
	wc.app = ln.app
	xl := ln.app

	wc.mx, _ = xl.findCore(ln.mx, "BackColor", 0)
	if ln.mx >= 0 {
		return &wc
	}

	_core, _ := xl.getCore(wc.mx)

	const cmd = "Get"
	const name = "BackColor"

	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case *ole.IDispatch:
			log.Printf("%v ans (object) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)

			wc.mx, _ = xl.addCore(ln.mx, x, name, 0)
			return &wc
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}

	return nil
}

func (ln *workLine) DashStyle(value ...any) int32 {
	xl := ln.app
	_core, err := xl.getCore(ln.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	var cmd string
	const name = "DashStyle"
	var opt []any

	if len(value) > 0 {
		cmd = "Put"

		var z int32
		switch x := value[0].(type) {
		case int:
			z = SetEnumLineDash(int32(x))
		case int32:
			z = SetEnumLineDash(x)
		case string:
			z = GetEnumLineDashNum(x)
		}
		opt = append(opt, z)

		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case nil:
				log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case int32:
			log.Printf("%v ans (int32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return -1
}

func (wc *workColor) RGB(value ...int32) int32 {
	xl := wc.app
	_core, err := xl.getCore(wc.mx)
	if err != nil {
		log.Printf("(Error) %v", err)
		return -1
	}
	var cmd string
	const name = "RGB"
	var opt []any

	if len(value) > 0 {
		cmd = "Put"
		opt = append(opt, value[0])
		args := xl.worker.Send(cmd, _core.disp, name, opt)

		for i := range args {
			switch x := args[i].(type) {
			case error:
				log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			case nil:
				log.Printf("%v ans (nil) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			default:
				log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, opt, _core.values)
			}
		}
	}

	cmd = "Get"
	args := xl.worker.Send(cmd, _core.disp, name, nil)

	for i := range args {
		switch x := args[i].(type) {
		case error:
			log.Printf("%v err %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		case int32:
			log.Printf("%v ans (int32) %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
			_core.values[name] = x
			return x
		default:
			log.Printf("%v def %v.%v.%v: %v v:%v, opt: %v\n", xl.hWnd, cmd, _core.kind, name, x, nil, _core.values)
		}
	}
	return -1
}
