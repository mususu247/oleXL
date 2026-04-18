package oleXL

import (
	"fmt"
	"math"
	"time"
)

func FromOADate(value float64) (time.Time, error) {
	var result time.Time

	if value <= 0 {
		return result, fmt.Errorf("invalid OLE date value: %f", value)
	}
	days := math.Floor(value)
	hns := (value - days) * 24
	hh := math.Floor(hns)
	hns = (hns - hh) * 60
	nn := math.Floor(hns)
	hns = (hns - nn) * 60
	ss := math.Floor(hns)

	result = time.Date(1900, 1, int(days)-1, int(hh), int(nn), int(ss), 0, time.UTC)
	return result, nil
}
