
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
start = time.time()

def ocr(file_pdf, file_img, path_excel, excelName):
    cons = []
    ocr_1 = []
    ocr_2 = []
    pdfs = os.listdir(file_pdf)
    for pdf in pdfs:
        ser = pdfs.index(pdf) + 1
        # PDF文件转图片
        Name = pdf[pdf.rfind("\\") + 1: pdf.rfind(".")]
        pdf_doc = fitz.open(file_pdf + "\\" + pdf)
        page = pdf_doc[0]
        zoom_x = 2
        zoom_y = 2
        mat = fitz.Matrix(zoom_x, zoom_y)
        pix = page.getPixmap(matrix=mat, alpha=False)
        img_ori = file_img + "\\" + Name + ".png"
        pix.writePNG(img_ori)
        print("第%s个PDF文件 ————"%ser, pdf, "已完成文件格式转换")

        # 图片灰度化
        im = Image.open(img_ori)
        im_gray = im.convert("L")
        im_arr = np.array(im_gray)
        im_1 = 255.0 * (im_arr / 255.0)
        plt.axis("off")
        plt.imshow(Image.fromarray(im_1), cmap='gray')
        img_gray = file_img + "\\" + Name + ".png"
        plt.savefig(img_gray, dpi=400)
        print("第%s个PDF文件 ————"%ser, pdf, "已完成图片灰度化")

        # 裁剪灰度图特定区域
        img = cv2.imdecode(np.fromfile(img_gray, dtype=np.uint8), -1)
        # 特定区域1
        cropImg_1 = img[500:600, 1100:1700]
        img_cropped_1 = file_img + "\\" + Name + "_1" + ".png"
        cv2.imwrite(img_cropped_1, cropImg_1)
        # 特定区域2
        cropImg_2 = img[700:800, 1200:1800]
        img_cropped_2 = file_img + "\\" + Name + "_2" + ".png"
        cv2.imwrite(img_cropped_2, cropImg_2)
        print("第%s个PDF文件 ————"%ser, pdf, "已完成指定区域裁剪")

        # 文字识别
        text_1 = pytesseract.image_to_string(Image.open(img_cropped_1), lang='chi_sim')
        text_2 = pytesseract.image_to_string(Image.open(img_cropped_2))
        rstr = r"[\=\(\)\,\/\\\:\*\?\"\<\>\|\' '\\\n\\\x0c]"
        i = re.sub(rstr, "", text_1)
        ocr_1.append(i)
        j = re.sub(rstr, "", text_2)
        ocr_2.append(j)

        # 添加文件名，并去除.pdf后缀
        cons.append(Name)
        print("第%s个PDF文件 ————" % ser, pdf, "已完成文字识别，并成功添加进字典")
        print("=" * 130)

    dir = {"文件名": cons, "特定区域1": ocr_1, "特定区域2": ocr_2}
    pd.DataFrame(dir).to_excel(path_excel + "\\" + excelName)

if __name__ == '__main__':
    file_pdf = r"C:\Users\ASUS\Desktop\ocr\TEST\1_pdf"
    file_img = r"C:\Users\ASUS\Desktop\ocr\TEST\2_img"
    path_excel = r"C:\Users\ASUS\Desktop\ocr\TEST"
    excelName = r"ocr_res.xlsx"
    ocr(file_pdf, file_img, path_excel, excelName)

    print("\n")
    end = time.time()
    print("程序结束！运行时间为：", end-start, "秒")
