import pandas
import os


def GetSheetNameToFile():
    FileName = file_name('./')
    data = GetExcelSheetName(FileName)
    ExcelIO(data)


def ExcelIO(data):
    try:
        xl = pandas.read_excel(r'sheetnames.xlsx', 0)
    except Exception as e:
        print('文件读取出错,err:ExcelIO')
    else:
        pass
    finally:
        pass
    # print(xl)
    # xl = pandas.DataFrame.from_dict(data, orient = 'index')
    xl = pandas.DataFrame(dict([(k, pandas.Series(v)) for k, v in data.items()]))
    xl.to_excel('sheetnames.xlsx')


# 获取文件名
def FileName(file_dir):
    for root, dirs, files in os.walk(file_dir):
        # print(root) #当前目录路径
        # print(dirs) #当前路径下所有子目录
        # print(files) #当前路径下所有非目录子文件
        return files


# 获取该目录下所有文件名,拆分为文件名+扩展名
def file_name(file_dir):
    L = []
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            # os.path.splitext(file)[1] 函数将路径拆分为文件名+扩展名
            if os.path.splitext(file)[1] == '.xlsx':
                if os.path.splitext(file)[0] == 'sheetnames':
                    continue
                L.append(os.path.join(root, file))
    return L


# 获取文件表格中sheet名称
def GetExcelSheetName(data):
    excelData = {}

    for i in range(len(data)):
        excelDataList = []
        try:
            xl = pandas.ExcelFile(data[i])
        except Exception as e:
            continue
        else:
            pass
        finally:
            pass

        sheet_names = xl.sheet_names  # 所有的sheet名
        # df = xl.parse(sheet_name)#读取ExceL中sheet

        for x in sheet_names:
            excelDataList.append(x)
        excelData[data[i]] = excelDataList

    return excelData
    # print(excelData)


if __name__ == '__main__':
    GetSheetNameToFile()
