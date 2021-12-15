import os
from PyPDF2 import PdfFileMerger
from xlrd import open_workbook  # 对xls、xlsx文件进行读操作
import pdfplumber


def getXczpDict(xczpPath):
    # 读取现场照片字典
    xczpDict = {}
    with pdfplumber.open(xczpPath) as xczpPdf:
        pageNum = len(xczpPdf.pages)
        for j in range(0, pageNum):
            text = xczpPdf.pages[j].extract_text()
            if xczpDict.__contains__(text):
                xczpDict[text].append(j)
            else:
                xczpDict[text] = []
                xczpDict[text].append(j)
    return xczpDict


def getDirPaths(xlsPath):
    # 获取文件夹名称
    dirPaths = []
    workbook = open_workbook(xlsPath)
    worksheet = workbook.sheet_by_name("分组准备")
    for i in range(1, worksheet.nrows):
        dirName = worksheet.cell(i, 0).value + worksheet.cell(i, 1).value
        if dirName is not None:
            # print(dirName)
            dirPath = os.path.join(root_path, dirName)
            dirPaths.append(dirPath)
    return dirPaths


def merge0810XCZP(dirPaths, xczpPath, xczpDict):
    k = 0
    pdf_merger = PdfFileMerger()
    pdf_merger_all = PdfFileMerger()

    for dirPath in dirPaths:
        dirName = dirPath.split("\\")[-1]
        if os.path.exists(dirPath):
            # 合并08&10
            pdf08 = 0
            pdf10 = 0
            for file in os.listdir(dirPath):
                # print(file[0:2])
                if file.endswith(".pdf") and file.startswith("08"):
                    pdf08 = 1
                    pdfPath = os.path.join(dirPath, file)
                    pdf_merger.append(pdfPath)
                elif file.endswith(".pdf") and file.startswith("10"):
                    pdf10 = 1
                    pdfPath = os.path.join(dirPath, file)
                    pdf_merger.append(pdfPath)
            if pdf08 == 0:
                print(dirName + "\\08pdf" + "_找不到")
            if pdf10 == 0:
                print(dirName + "\\10pdf" + "_找不到")

            # 获取现场照片pdf对应页
            if xczpDict.__contains__(dirName):
                pageIndex1 = xczpDict[dirName][0]
                pageIndex2 = xczpDict[dirName][-1]
                if pageIndex1 == pageIndex2:
                    pdf_merger.append(fileobj=xczpPath, pages=(pageIndex1, pageIndex1+1))
                else:
                    pdf_merger.append(fileobj=xczpPath, pages=(pageIndex1, pageIndex2+1))
            else:
                print(dirName + "_现场照片_找不到")
            # pdf_merger_all = pdf_merger
            k = k + 1
            if k % 100 == 0:
                output_pdf = xczpPath[0:-10] + "08_10_现场照片_" + str(k) + "_.pdf"
                pdf_merger.write(output_pdf)
                # pdf_merger_all.append(fileobj=pdf_merger, pages=(0,pdf_merger.pages._len_))
                pdf_merger = PdfFileMerger()
        else:
            print(dirName + "_文件夹_找不到")
    output_pdf = xczpPath[0:-10] + "08_10_现场照片_" + str(k) + "_.pdf"
    pdf_merger.write(output_pdf)
    # return pdf_merger_all


if __name__ == '__main__':
    # 路径
    root_path = os.getcwd()
    xczpPath = ""
    xlsPath = ""
    pdf_merger_all = PdfFileMerger()

    for f in os.listdir(root_path):
        if f.endswith("现场照片分组.pdf"):
            xczpPath = os.path.join(root_path, f)
        if f.endswith("现场照片分组.xls"):
            xlsPath = os.path.join(root_path, f)

    if xlsPath == "" or xczpPath == "":
        input("请先将 XX村现场照片分组.xls 和 XX村现场照片分组.pdf 复制到当前文件夹中。")
    else:
        print("开始合并，请稍候。。。")
        xczpDict = getXczpDict(xczpPath)
        dirPaths = getDirPaths(xlsPath)
        merge0810XCZP(dirPaths, xczpPath, xczpDict)
        '''
        pdf_merger_all = merge0810XCZP(dirPaths, xczpPath, xczpDict)
        output_pdf = xczpPath[0:-10] + "08_10_现场照片" + "_全.pdf"
        pdf_merger_all.write(output_pdf)
        '''
        input("注意找不到的文件夹和文件...")
