package oleXL

import (
	"strings"
)

func NumberFoarmat2Layout(format string) string {
	layout := format

	if strings.Contains(layout, "yyyy") {
		layout = strings.ReplaceAll(layout, "yyyy", "2006")
	}
	if strings.Contains(layout, "yy") {
		layout = strings.ReplaceAll(layout, "yyyy", "06")
	}

	if strings.Contains(layout, "/mm") {
		layout = strings.ReplaceAll(layout, "/mm", "/01")
	}
	if strings.Contains(layout, "/m") {
		layout = strings.ReplaceAll(layout, "/m", "/1")
	}

	if strings.Contains(layout, "mm/") {
		layout = strings.ReplaceAll(layout, "mm/", "01/")
	}
	if strings.Contains(layout, "m/") {
		layout = strings.ReplaceAll(layout, "m/", "1/")
	}

	if strings.Contains(layout, "dd") {
		layout = strings.ReplaceAll(layout, "dd", "02")
	}
	if strings.Contains(layout, "d") {
		layout = strings.ReplaceAll(layout, "d", "2")
	}

	if strings.Contains(layout, "hh") {
		layout = strings.ReplaceAll(layout, "hh", "15")
	}
	if strings.Contains(layout, "h") {
		layout = strings.ReplaceAll(layout, "h", "15")
	}

	if strings.Contains(layout, ":mm") {
		layout = strings.ReplaceAll(layout, ":mm", ":04")
	}
	if strings.Contains(layout, ":m") {
		layout = strings.ReplaceAll(layout, ":m", ":4")
	}

	if strings.Contains(layout, "mm:") {
		layout = strings.ReplaceAll(layout, "mm:", "04:")
	}
	if strings.Contains(layout, "m:") {
		layout = strings.ReplaceAll(layout, "m:", "4:")
	}

	if strings.Contains(layout, "ss") {
		layout = strings.ReplaceAll(layout, "ss", "05")
	}
	if strings.Contains(layout, "s") {
		layout = strings.ReplaceAll(layout, "s", "5")
	}
	return layout
}
