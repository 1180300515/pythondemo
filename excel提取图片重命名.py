import openpyxl
from openpyxl_image_loader import SheetImageLoader
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# 制作获取文件的弹窗
root = tk.Tk()
root.wm_attributes('-topmost',1) # 弹窗置顶
root.withdraw()
file_path = filedialog.askopenfilename()  # 获取 excel 路径
# 加载excel表和图片
pxl_doc = openpyxl.load_workbook(file_path) 
sheet = pxl_doc["Sheet1"]  # excel的Sheet名
image_loader = SheetImageLoader(sheet)

# 用pd获取图片所在列的起止行号list——ls, 此处省略代码
ls = []
for i in range(2,262):
    ls.append(i)
print(ls)
# # 用pd获取图片名称所在列list——image_name，此处省略代码
image_name = []
data = pd.read_excel(file_path)
data_li = data.values.tolist()
for i in data_li:
    image_name.append(i[0])
print(image_name)


for i in range(len(ls)):
    #get the image (put the cell you need instead of 'A1')
    image = image_loader.get("B"+str(ls[i]))  # 假设图片在C列
    image.save('/照片/'+str(image_name[i])+'.png')