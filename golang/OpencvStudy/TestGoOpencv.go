package OpencvStudy

import (
	"fmt"
	"gocv.io/x/gocv"
)

func TestGoOpencv() {
	fmt.Print("test")
	webcam, _ := gocv.OpenVideoCapture(0)
	window := gocv.NewWindow("Hello")
	img := gocv.NewMat()

	for {
		webcam.Read(&img)
		window.IMShow(img)
		window.WaitKey(1)
	}
}
