
import os
import fitz
from PIL import Image
from matplotlib import pyplot as plt
import numpy as np
import cv2
import pytesseract
import pandas as pd
import re
import time

# PDF转图片
def pdf_to_img(file_pdf, file_img):
    print("正在将PDF文件转换成图片......")
    cons = []
    pdfs = os.listdir(file_pdf)
    for pdf in pdfs:
        pdf_doc = fitz.open(file_pdf + "\\" + pdf)
        page = pdf_doc[0]
        zoom_x = 2
        zoom_y = 2
        mat = fitz.Matrix(zoom_x, zoom_y)
        pix = page.getPixmap(matrix=mat, alpha=False)
        Name = pdf[pdf.rfind("\\") + 1 : pdf.rfind(".")]
        image_ori = file_img + "\\" + Name + ".png"
        pix.writePNG(image_ori)
        cons.append(Name)
    dir_1 = {"文件名":cons}
    return dir_1

# 图片灰度化
def img_gary(file_img):
    print("正在进行图片灰度化......")
    imgs_ori = os.listdir(file_img)
    for img_ori in imgs_ori:
        im = Image.open(file_img + "\\" + img_ori)
        im_gray = im.convert("L")
        im_arr = np.array(im_gray)
        im_1 = 255.0 * (im_arr / 255.0) ** 2
        plt.axis("off")
        plt.imshow(Image.fromarray(im_1), cmap='gray')
        Name = img_ori[img_ori.rfind("\\") + 1: img_ori.rfind(".")]
        img_gray = file_img + "\\" + Name + ".png"
        plt.savefig(img_gray, dpi=400)

# 裁剪灰度图特定区域
def img_cropping(file_img):
    print("正在进行特定区域裁剪......")
    imgs_gray = os.listdir(file_img)
    for img_gray in imgs_gray:
        img = cv2.imdecode(np.fromfile(file_img + "\\" + img_gray, dtype=np.uint8), -1)
        # 特定区域1
        cropImg_1 = img[500:600, 1100:1700] #[y start:y end, x start:x end]，参数需要自定义，可以通过PS查看具体坐标
        Name_1 = img_gray[img_gray.rfind("\\") + 1: img_gray.rfind(".")] + "_1"
        img_cropped_1 = file_img + "\\" + Name_1 + ".png"
        cv2.imwrite(img_cropped_1, cropImg_1)
        # 特定区域2
        cropImg_2 = img[700:800, 1200:1800]
        Name_2 = img_gray[img_gray.rfind("\\") + 1 : img_gray.rfind(".")] + "_2"
        img_cropped_2 = file_img + "\\" + Name_2 + ".png"
        cv2.imwrite(img_cropped_2, cropImg_2)

# 文字识别，并将所识别文字添加进DataFrame中
def ocr(file_img, dir_1, path_excel, excelName):
    print("正在进行文字识别......")
    ocr_1 = []
    ocr_2 = []
    rstr = r"[\=\(\)\,\/\\\:\*\?\"\<\>\|\' '\\\n\\\x0c]"
    imgs_cropped = os.listdir(file_img)
    for img_cropped in imgs_cropped:
        if img_cropped.endswith("_1.png"):
            # 指定tesseract中文语言包，否则无法识别中文。同时去除所识别文字中\n、\0xc等冗余字符
            text_1 = pytesseract.image_to_string(Image.open(file_img + "\\" + img_cropped), lang='chi_sim')
            text_1 = re.sub(rstr, "", text_1)
            ocr_1.append(text_1)
        elif img_cropped.endswith("_2.png"):
            text_2 = pytesseract.image_to_string(Image.open(file_img + "\\" + img_cropped))
            text_2 = re.sub(rstr, "", text_2)
            ocr_2.append(text_2)

    dir_2 = {"特定区域1":ocr_1, "特定区域2":ocr_2}
    dir_1.update(dir_2)
    pd.DataFrame(dir_1).to_excel(path_excel + "\\" + excelName)

if __name__ == '__main__':
    file_pdf = r"C:\Users\ASUS\Desktop\ocr\TEST\1_pdf" #存放PDF文件的文件夹
    file_img= r"C:\Users\ASUS\Desktop\ocr\TEST\2_img" #存放各种图片的文件夹
    path_excel = r"C:\Users\ASUS\Desktop\ocr\TEST" #存放各种Excle的文件夹
    excelName = r"ocr_res.xlsx"
    dir_1 = pdf_to_img(file_pdf, file_img)
    print("文件格式已转换成功")
    print("="*130)
    img_gary(file_img)
    print("图片已灰度化成功")
    print("="*130)
    img_cropping(file_img)
    print("灰度图已裁剪成功")
    print("="*130)
    ocr(file_img, dir_1, path_excel, excelName)
    print("文字已成功识别，并已将内容填充进Excel表中")
    print("=" * 130)
    print("\n")
    print("程序运行时间为：", time.process_time())


# 注：
# 这里的 pdf_to_img() 只提取第一页内容，若想全部提取，则代码为：
# def pdf_to_img(file_pdf, file_img_ori):
#     pdfs = os.listdir(file_pdf)
#     for pdf in pdfs:
#         Name = pdf[pdf.rfind("\\") + 1 : pdf.rfind(".")]
#         pdf_doc = fitz.open(file_pdf + "\\" + pdf)
#         for pg in range(pdf_doc.pageCount):
#             page = pdf_doc[pg]
#             zoom_x = 2
#             zoom_y = 2
#             mat = fitz.Matrix(zoom_x, zoom_y)
#             pix = page.getPixmap(matrix=mat, alpha=False)
#             image_ori = file_img_ori + "\\" + Name + "-%i.png" % pg
#             pix.writePNG(image_ori)
