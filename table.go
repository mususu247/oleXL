package oleXL

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
)

type workTable struct {
	app    *Excel
	parent any
	num    int
}

type workTables struct {
	app    *Excel
	parent any
	num    int
}

type listRow struct {
	app    *Excel
	parent any
	num    int
}

type listRows struct {
	app    *Excel
	parent any
	num    int
}

type listColumn struct {
	app    *Excel
	parent any
	num    int
}

type listColumns struct {
	app    *Excel
	parent any
	num    int
}

func (Q *workSheet) ListObjects() *workTables {
	var body workTables
	xl := Q.app

	name := "ListObjects"
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

func (Q *workSheet) ListObjectz(value any) *workTable {
	var body workTable
	xl := Q.app
	tbs := Q.ListObjects()

	kind := "ListObject"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
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

		ans, err := xl.cores.SendNum(cmd, name, tbs.num, opt)
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
	body.parent = tbs
	return &body
}

func (Q *workTables) Release() error {
	xl := Q.app
	return xl.cores.Release(Q.num, false)
}

func (Q *workTables) Nothing() error {
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

func (Q *workTables) Set() (*workTables, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workTables) Add(SourceType any, Source *workRange, option ...any) *workTable {
	var body workTable
	xl := Q.app

	kind := "ListObject"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Method"
		name := "Add"
		var opt []any
		for range 4 {
			opt = append(opt, nil)
		}

		var z int32
		switch x := SourceType.(type) {
		case int:
			z = SetEnumListObjectSourceType(int32(x))
		case int32:
			z = SetEnumListObjectSourceType(x)
		case string:
			z = GetEnumListObjectSourceTypeNum(x)
		default:
			z = SetEnumListObjectSourceType(0)
		}
		opt[0] = z

		xcore := xl.cores.getCore(Source.num)
		opt[1] = xcore.disp
		opt[2] = true

		switch x := option[1].(type) {
		case int:
			z = SetEnumYesNoGuess(int32(x))
		case int32:
			z = SetEnumYesNoGuess(x)
		case string:
			z = GetEnumYesNoGuessNum(x)
		default:
			z = SetEnumYesNoGuess(0)
		}
		opt[3] = z

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

func (Q *workTables) Count() int32 {
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

func (Q *workTable) Release() error {
	xl := Q.app
	return xl.cores.Release(Q.num, false)
}

func (Q *workTable) Nothing() error {
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

func (Q *workTable) Set() (*workTable, error) {
	if Q == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := Q.app
	xl.cores.Lock(Q.num)
	return Q, nil
}

func (Q *workTable) Range(value ...any) *workRange {
	var body workRange
	xl := Q.app

	kind := "Range"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Range"
		var opt []any
		if len(value) > 0 {
			opt = append(opt, value[0])
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

func (Q *workTable) Name(value ...any) string {
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

func (Q *workTable) TableStyle(value ...any) string {
	xl := Q.app

	name := "TableStyle"
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

func (Q *workTable) ShowHeaders(value ...any) bool {
	xl := Q.app

	name := "ShowHeaders"
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

func (Q *workTable) ShowTableStyleRowStripes(value ...any) bool {
	xl := Q.app

	name := "ShowTableStyleRowStripes"
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

func (Q *workTable) ShowTotals(value ...any) bool {
	xl := Q.app

	name := "ShowTotals"
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

func (Q *workTable) ShowTableStyleColumnStripes(value ...any) bool {
	xl := Q.app

	name := "ShowTableStyleColumnStripes"
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

func (Q *workTable) ShowTableStyleLastColumn(value ...any) bool {
	xl := Q.app

	name := "ShowTableStyleLastColumn"
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

func (Q *workTable) ShowTableStyleFirstColumn(value ...any) bool {
	xl := Q.app

	name := "ShowTableStyleFirstColumn"
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

func (Q *workTable) ShowAutoFilterDropDown(value ...any) bool {
	xl := Q.app

	name := "ShowAutoFilterDropDown"
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

func (Q *workTable) HeaderRowRange() *workRange {
	var body workRange
	xl := Q.app

	kind := "Range"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "HeaderRowRange"
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

func (Q *workTable) DataBodyRange() *workRange {
	var body workRange
	xl := Q.app

	kind := "Range"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "DataBodyRange"
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

func (Q *workTable) ListRows() *listRows {
	var body listRows
	xl := Q.app

	name := "ListRows"
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

func (Q *workTable) ListRowz(value any) *listRow {
	var body listRow
	xl := Q.app
	lrs := Q.ListRows()

	kind := "ListColumn"
	core, num := xl.cores.FindAdd(kind, lrs.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Item"
		var opt []any
		switch x := value.(type) {
		case int:
			opt = append(opt, int32(x))
		case int32:
			opt = append(opt, x)
		case string:
			opt = append(opt, x)
		}

		ans, err := xl.cores.SendNum(cmd, name, lrs.num, opt)
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
	body.parent = lrs
	return &body
}

func (Q *workTable) Activate() error {
	xl := Q.app

	cmd := "Method"
	name := "Activate"

	_, err := xl.cores.SendNum(cmd, name, Q.num, nil)
	if err != nil {
		return err
	}
	return nil
}

func (Q *listRow) Range(value ...any) *workRange {
	var body workRange
	xl := Q.app

	kind := "Range"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Range"
		var opt []any
		if len(value) > 0 {
			opt = append(opt, value[0])
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

func (Q *listRows) Count() int32 {
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

func (Q *listRow) Index() int32 {
	xl := Q.app

	cmd := "Get"
	name := "Index"
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

func (Q *workTable) ListColumns() *listColumns {
	var body listColumns
	xl := Q.app

	name := "ListColumns"
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

func (Q *workTable) ListColumnz(value any) *listColumn {
	var body listColumn
	xl := Q.app
	lcs := Q.ListColumns()

	kind := "ListColumn"
	core, num := xl.cores.FindAdd(kind, lcs.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Item"
		var opt []any
		switch x := value.(type) {
		case int:
			opt = append(opt, int32(x))
		case int32:
			opt = append(opt, x)
		case string:
			opt = append(opt, x)
		}

		ans, err := xl.cores.SendNum(cmd, name, lcs.num, opt)
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
	body.parent = lcs
	return &body
}

func (Q *listColumn) Range(value ...any) *workRange {
	var body workRange
	xl := Q.app

	kind := "Range"
	core, num := xl.cores.FindAdd(kind, Q.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Range"
		var opt []any
		if len(value) > 0 {
			opt = append(opt, value[0])
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

func (Q *listColumns) Count() int32 {
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

func (Q *listColumn) Name() string {
	xl := Q.app

	cmd := "Get"
	name := "Name"
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

func (Q *workTable) Comment(value ...any) string {
	xl := Q.app

	name := "Comment"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		if len(value) > 0 {
			for _, v := range value {
				opt = append(opt, v)
			}
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
