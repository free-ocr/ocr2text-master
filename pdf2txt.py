from PIL import Image
import os
import pytesseract
import cv2 as cv
import fitz

def pdf2txt(pdfPath,zoom_x,zoom_y,rotation_angle):
    # 打开PDF文件
    pdf = fitz.open(pdfPath)
    text = ''
    imgPath = './test_files/image'
    # 逐页读取PDF
    for pg in range(0, pdf.pageCount):
        page = pdf[pg]
        # 设置缩放和旋转系数
        trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotation_angle)
        pm = page.getPixmap(matrix=trans, alpha=False)
        # 开始写图像
        pm.writePNG(imgPath+str(pg)+".png")
        img = cv.imread(imgPath+str(pg)+".png")
        text1 = pytesseract.image_to_string(Image.fromarray(img), lang='chi_sim')
        os.remove(imgPath+str(pg)+".png")
        text = text + text1
        #pm.writePNG(imgPath)
    pdf.close()
    return text

# def pdf2txt(pdfPath):
#     img_path = './test_files/image.png'
#     pdf_image(pdfPath, img_path, 5, 5, 0)
#     # 依赖opencv
#     img = cv.imread(img_path)
#     text = pytesseract.image_to_string(Image.fromarray(img), lang='chi_sim')
#     # 不依赖opencv写法
#     # text=pytesseract.image_to_string(Image.open(img_path))
#     #print(text)
#     return text


if __name__ == '__main__':
    pdf_path = './test_files/QAR-ZM17-0057.pdf'

    print(pdf2txt(pdf_path, 5, 5, 0))