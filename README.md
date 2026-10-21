# oleXL

```go
package main

import (
  "fmt"
  "github.com/mususu247/oleXL"
)

func main() {
  var xl oleXL.Excel

  xl.Init()
  err := xl.CreateObject()
  if err != nil {
    fmt.Printf("(Error) %v", err)
    return
  }
  xl.Visible(true)
	
  xl.Quit()
  xl.Nothing()
}
```
