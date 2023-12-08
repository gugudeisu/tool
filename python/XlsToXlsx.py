import os
import win32com.client




# 递归该目录及其下所有目录将其转为xlsx
def toXlsx():
    # a = input("输入1将当前及其下级目录中xls文件保存为xlsx\n")
    a = 1
    if a == 1 :
        # i = os.getcwd()
        i = r"C:\栈\cs"
        # C:\栈\cs
        # XlsToXlsx(f"{i}")
        # 当前目录（.）作为根目录开始递归处理
        # current_directory = f"{i}"
        print("当前目录: C:\栈\cs")
        # XlsToXlsx(f"{i}")
        XlsToXlsx(r"C:\栈\cs")
        # current_directory = f"{i}"
        current_directory = r"C:\栈\cs"
        recursive_directory_processing(current_directory)
        # os.system("pause")
    else:
        return



# 定义你要执行的函数，例如处理文件的函数
def process_directory(file_path):
    # 这里可以执行你的处理逻辑
    print(f"当前目录: {file_path}")
    XlsToXlsx(f"{file_path}")

# 定义递归函数来遍历目录并执行处理函数
def recursive_directory_processing(directory):
    for root, dirs, files in os.walk(directory):
        for dir_name in dirs:
            dir_path = os.path.join(root, dir_name)
            process_directory(dir_path)





def xls_to_xlsx(file_path, excel):
    """
    将指定的xls文件转化为xlsx格式
    file_path: 文件路径
    excel: 代表Excel应用程序
    """
    # 打开原始文档
    workbook = excel.Workbooks.Open(file_path)

    # 将文档另存为xlsx格式
    new_file_path = os.path.splitext(file_path)[0] + ".xlsx"
    workbook.SaveAs(new_file_path, FileFormat=51)

    # 关闭文档
    workbook.Close()

    # 删除原始文件
    os.remove(file_path)
    # 打印操作过程
    print(f"{file_path}已经被成功转换为{new_file_path}")

def XlsToXlsx(p):
    # 定义文件夹路径和Excel应用程序对象
    folder_path = p
    excel = win32com.client.Dispatch("Excel.Application")

    # 遍历文件夹中所有的.xls文件，并将其转换为.xlsx格式
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xls"):
                file_path = os.path.join(root, file)
                xls_to_xlsx(file_path, excel)

    # 关闭Excel应用程序
    excel.Quit()

    print("全部xls文件已经全部转换为xlsx!")



