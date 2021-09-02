# OCR_python

## 实现逻辑
- 通过PyMuPDF库的fitz模块将图片版PDF转为图片格式；
- 通过PIL库实现对图片文件灰度化（因为我们需要用Tesseract进行文字识别，而Tesseract对彩色图片识别的支持不是很友好.......）；
- 通过OpenCV对灰度图进行坐标定位，从而实现特定区域裁剪；
- 最后通过Tesseract对裁剪好的灰度图进行文字识别；
- 通过Pandas对所识别内容与文件名称所匹配，并填充进Excel中。

## 版本一
此版本程序逻辑为，面对所有pdf文件分步执行程序流程。

## 版本二
此版本程序逻辑为，对每个pdf文件轮流执行全部流程。

## 注
由于在我的工作内容中，所需信息都在PDF文件首页，故pdf_to_img()只需提取第一页内容，若想将pdf各页面全部提取，则代码为：
```python
def pdf_to_img(file_pdf, file_img_ori):
    pdfs = os.listdir(file_pdf)
    for pdf in pdfs:
        Name = pdf[pdf.rfind("\\") + 1 : pdf.rfind(".")]
        pdf_doc = fitz.open(file_pdf + "\\" + pdf)
        for pg in range(pdf_doc.pageCount):
			page = pdf_doc[pg]
			zoom_x = 2
            zoom_y = 2
            mat = fitz.Matrix(zoom_x, zoom_y)
            pix = page.getPixmap(matrix=mat, alpha=False)
            image_ori = file_img + "\\" + Name + "-%i.png" % pg
            pix.writePNG(image_ori)
```
