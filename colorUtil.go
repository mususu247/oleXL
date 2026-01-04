package oleXL

import "strconv"

// version 2026-01-04

func Color2RGB(value float64) map[string]int {
	var rgb map[string]int
	rgb = make(map[string]int)

	v0 := int(value)
	v1 := v0 % 256
	v0 = v0 - v1
	v0 = v0 / 256
	v2 := v0 % 256
	v0 = v0 - v2
	v0 = v0 / 256
	v3 := v0 % 256
	v0 = v0 - v3
	v4 := v0 / 256

	rgb["Red"] = v1
	rgb["Green"] = v2
	rgb["Blue"] = v3
	rgb["Alpha"] = v4
	return rgb
}

func RGB(red, green, blue int) int32 {
	r := uint8(red)
	g := uint8(green)
	b := uint8(blue)

	var color int32
	color = int32(b)
	color = color * 256
	color = color + int32(g)
	color = color * 256
	color = color + int32(r)
	return color
}

func Code2Color(value string) float64 {
	v0, _ := strconv.ParseInt(value, 16, 64)
	return float64(v0)
}

func RGB2Gray(red, green, blue int) float64 {
	r := uint8(red)
	g := uint8(green)
	b := uint8(blue)

	rr := float64(r) * 0.299
	gg := float64(g) * 0.587
	bb := float64(b) * 0.114
	xx := rr + gg + bb
	x := uint8(xx)
	xx = float64(x)

	var color float64
	color = xx
	color = color * 256
	color = color + xx
	color = color * 256
	color = color + xx
	return color
}
