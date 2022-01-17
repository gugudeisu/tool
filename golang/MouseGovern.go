package golang

import (
	"github.com/go-vgo/robotgo"
	"time"
)

// MouseGovern 每五分钟模拟鼠标移动,一天后结束
func MouseGovern() {

	var x int
	var y int
	x, y = robotgo.GetMousePos()
	println(`x：`, x, ` y：`, y)
	robotgo.MoveMouseSmooth(x+10, y, 1)
	robotgo.MoveMouseSmooth(x-20, y, 1)
	robotgo.MoveMouseSmooth(x+10, y, 1)
	println(`x：`, x, ` y：`, y)

	for i := 1; i < 288; i++ {
		time.Sleep(300 * time.Second)
		x, y = robotgo.GetMousePos()
		println(`x：`, x, ` y：`, y)
		robotgo.MoveMouseSmooth(x+10, y, 1)
		robotgo.MoveMouseSmooth(x-20, y, 1)
		robotgo.MoveMouseSmooth(x+10, y, 1)
		println(`x：`, x, ` y：`, y)

	}

}
