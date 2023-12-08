import cv2 as cv


# 学习Opencv
def opencvtest():
    # imread写入图像
    img = cv.imread('data/1.jpg')
    cv.imshow('img', img)
    # 绑定时间 按下任意键继续
    cv.waitKey(0)
    cv.destroyAllWindows()
    # waitKey 保存


