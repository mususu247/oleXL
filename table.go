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

func table2sheet(v any) *workSheet {
	var ws *workSheet

	w := v
	for {
		switch x := w.(type) {
		case *workTable:
			w = x.parent
		case *workTables:
			w = x.parent
		case *listColumn:
			w = x.parent
		case *listColumns:
			w = x.parent
		case *listRow:
			w = x.parent
		case *listRows:
			w = x.parent
		case *workSheet:
			ws = x
			return ws
		}
	}
}

func (ws *workSheet) ListObjects() *workTables {
	var tbs workTables
	xl := ws.app

	name := "ListObjects"
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
			if x != nil {
				core.disp = x
				core.lock = 0
			} else {
				return nil
			}
		}
	}
	tbs.app = xl
	tbs.num = num
	tbs.parent = ws
	return &tbs
}

func (ws *workSheet) ListObjectz(value any) *workTable {
	var tb workTable
	xl := ws.app
	tbs := ws.ListObjects()

	kind := "ListObject"
	core, num := xl.cores.FindAdd(kind, ws.num)
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
	tb.app = xl
	tb.num = num
	tb.parent = tbs
	return &tb
}

func (tbs *workTables) Release() error {
	xl := tbs.app
	return xl.cores.Release(tbs.num, false)
}

func (tbs *workTables) Nothing() error {
	xl := tbs.app
	xl.cores.releaseChild(tbs.num)

	xl.cores.Unlock(tbs.num)
	err := tbs.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(tbs.num)
	tbs = nil
	return nil
}

func (tbs *workTables) Set() (*workTables, error) {
	if tbs == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := tbs.app
	xl.cores.Lock(tbs.num)
	return tbs, nil
}

func (tbs *workTables) Add(SourceType any, Source *workRange, option ...any) *workTable {
	var tb workTable
	xl := tbs.app

	kind := "ListObject"
	core, num := xl.cores.FindAdd(kind, tbs.num)
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
		}
		opt[3] = z

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
	tb.app = xl
	tb.num = num
	tb.parent = tbs
	return &tb
}

func (tbs *workTables) Count() int32 {
	xl := tbs.app

	name := "Count"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, tbs.num, nil)
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

func (tb *workTable) Release() error {
	xl := tb.app
	return xl.cores.Release(tb.num, false)
}

func (tb *workTable) Nothing() error {
	xl := tb.app
	xl.cores.releaseChild(tb.num)

	xl.cores.Unlock(tb.num)
	err := tb.Release()
	if err != nil {
		return err
	}
	xl.cores.Remove(tb.num)
	tb = nil
	return nil
}

func (tb *workTable) Set() (*workTable, error) {
	if tb == nil {
		return nil, fmt.Errorf("(Error) Object is NULL.")
	}
	xl := tb.app
	xl.cores.Lock(tb.num)
	return tb, nil
}

func (tb *workTable) Range(value ...any) *workRange {
	var rg workRange
	xl := tb.app

	kind := "Range"
	core, num := xl.cores.FindAdd(kind, tb.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Range"
		var opt []any
		if len(value) > 0 {
			opt = append(opt, value[0])
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, tb.num, opt)
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
	rg.app = xl
	rg.num = num
	rg.parent = tb
	return &rg
}

func (tb *workTable) Name(value ...any) string {
	xl := tb.app

	name := "Name"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case string:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, tb.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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

func (tb *workTable) TableStyle(value ...any) string {
	xl := tb.app

	name := "TableStyle"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case string:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, tb.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return ""
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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

func (tb *workTable) ShowHeaders(value ...any) bool {
	xl := tb.app

	name := "ShowHeaders"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case bool:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, tb.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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

func (tb *workTable) ShowTableStyleRowStripes(value ...any) bool {
	xl := tb.app

	name := "ShowTableStyleRowStripes"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case bool:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, tb.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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

func (tb *workTable) ShowTotals(value ...any) bool {
	xl := tb.app

	name := "ShowTotals"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case bool:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, tb.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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

func (tb *workTable) ShowTableStyleColumnStripes(value ...any) bool {
	xl := tb.app

	name := "ShowTableStyleColumnStripes"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case bool:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, tb.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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

func (tb *workTable) ShowTableStyleLastColumn(value ...any) bool {
	xl := tb.app

	name := "ShowTableStyleLastColumn"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case bool:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, tb.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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

func (tb *workTable) ShowTableStyleFirstColumn(value ...any) bool {
	xl := tb.app

	name := "ShowTableStyleFirstColumn"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case bool:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, tb.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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

func (tb *workTable) ShowAutoFilterDropDown(value ...any) bool {
	xl := tb.app

	name := "ShowAutoFilterDropDown"
	if len(value) > 0 {
		cmd := "Put"
		var opt []any
		switch x := value[0].(type) {
		case bool:
			opt = append(opt, x)
		}

		_, err := xl.cores.SendNum(cmd, name, tb.num, opt)
		if err != nil {
			log.Printf("(Error) %v", err)
			return false
		}
	} else {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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

func (tb *workTable) HeaderRowRange() *workRange {
	var wr workRange
	xl := tb.app

	kind := "Range"
	core, num := xl.cores.FindAdd(kind, tb.num)
	if core.disp == nil {
		cmd := "Get"
		name := "HeaderRowRange"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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
	wr.app = xl
	wr.num = num
	wr.parent = tb
	return &wr
}

func (tb *workTable) DataBodyRange() *workRange {
	var wr workRange
	xl := tb.app

	kind := "Range"
	core, num := xl.cores.FindAdd(kind, tb.num)
	if core.disp == nil {
		cmd := "Get"
		name := "DataBodyRange"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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
	wr.app = xl
	wr.num = num
	wr.parent = tb
	return &wr
}

func (tb *workTable) ListRows() *listRows {
	var lr listRows
	xl := tb.app

	name := "ListRows"
	core, num := xl.cores.FindAdd(name, tb.num)
	if core.disp == nil {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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
	lr.app = xl
	lr.num = num
	lr.parent = tb
	return &lr
}

func (tb *workTable) ListRowz(value any) *listRow {
	var lr listRow
	xl := tb.app
	lrs := tb.ListRows()

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
	lr.app = xl
	lr.num = num
	lr.parent = lrs
	return &lr
}

func (lr *listRow) Range(value ...any) *workRange {
	var rg workRange
	xl := lr.app

	kind := "Range"
	core, num := xl.cores.FindAdd(kind, lr.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Range"
		var opt []any
		if len(value) > 0 {
			opt = append(opt, value[0])
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, lr.num, opt)
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
	rg.app = xl
	rg.num = num
	rg.parent = lr
	return &rg
}

func (lrs *listRows) Count() int32 {
	xl := lrs.app

	name := "Count"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, lrs.num, nil)
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

func (lr *listRow) Index() int32 {
	var result int32
	xl := lr.app

	cmd := "Get"
	name := "Index"
	ans, err := xl.cores.SendNum(cmd, name, lr.num, nil)
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

func (tb *workTable) ListColumns() *listColumns {
	var lcs listColumns
	xl := tb.app

	name := "ListColumns"
	core, num := xl.cores.FindAdd(name, tb.num)
	if core.disp == nil {
		cmd := "Get"
		ans, err := xl.cores.SendNum(cmd, name, tb.num, nil)
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
	lcs.app = xl
	lcs.num = num
	lcs.parent = tb
	return &lcs
}

func (tb *workTable) ListColumnz(value any) *listColumn {
	var lc listColumn
	xl := tb.app
	lcs := tb.ListColumns()

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
	lc.app = xl
	lc.num = num
	lc.parent = lcs
	return &lc
}

func (lc *listColumn) Range(value ...any) *workRange {
	var rg workRange
	xl := lc.app

	kind := "Range"
	core, num := xl.cores.FindAdd(kind, lc.num)
	if core.disp == nil {
		cmd := "Get"
		name := "Range"
		var opt []any
		if len(value) > 0 {
			opt = append(opt, value[0])
		} else {
			opt = nil
		}

		ans, err := xl.cores.SendNum(cmd, name, lc.num, opt)
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
	rg.app = xl
	rg.num = num
	rg.parent = lc
	return &rg
}

func (lcs *listColumns) Count() int32 {
	xl := lcs.app

	name := "Count"
	cmd := "Get"
	ans, err := xl.cores.SendNum(cmd, name, lcs.num, nil)
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

func (lc *listColumn) Name() string {
	var result string
	xl := lc.app

	cmd := "Get"
	name := "Name"
	ans, err := xl.cores.SendNum(cmd, name, lc.num, nil)
	if err != nil {
		log.Printf("(Error) %v", err)
		return result
	}
	switch x := ans.(type) {
	case string:
		result = x
	}
	return result
}
