import python.GetSheetNameToExcel
import python.OpencvTest
import python.PDFToIMG


def main():
    # python.GetSheetNameToExcel.GetSheetNameToFile()
    # opencv 学习
    # python.OpencvTest.opencvtest()
    # pdf转图片
    python.PDFToIMG.pdf2img(filename=r'D:/临时文件/gwt流程图.pdf')
    # imgtopdf(input_paths="./", outputpath="./docimg.pdf", file_type=".png")
    print("\n转换完成！")


if __name__ == '__main__':
    main()
