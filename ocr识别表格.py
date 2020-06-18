# -*- coding:utf-8 -*-
import cv2
import pytesseract
import numpy as np
import win32com.client as win32
import os


class Pdf2xlsx(object):
	def __init__(self, path):
		self.path = path
		self.image = cv2.imread(path)
		self.x, self.y = [], []
		self.ocr_result = []

	def chooseBlackArea(self, img):
		return cv2.inRange(img, np.array([0, 0, 0], np.uint8), np.array([180, 255, 46], np.uint8)) + \
			   cv2.inRange(img, np.array([0, 0, 46], np.uint8), np.array([180, 43, 220], np.uint8))  # 选取黑色及黑色的范围， 排除如盖章之类的元素干扰

	def extractLines(self):
		img = cv2.cvtColor(self.image, cv2.COLOR_BGR2HSV)
		ROI = self.chooseBlackArea(img)
		# 提取表格横线
		horizon_kernel = np.array([[0] * 100, [1] * 100], np.uint8)
		horizon = cv2.bitwise_or(cv2.erode(ROI, horizon_kernel, anchor=(0, 0)),cv2.erode(ROI, horizon_kernel, anchor=(-1, 0)))
		horizon = cv2.dilate(horizon, np.array([[1, 1, 0]], np.uint8), iterations=300)
		cv2.imwrite('h.jpg', horizon)

		# 提取表格竖线
		vertical_kernel = np.array([[0, 1]] * 100, np.uint8)
		# vertical = cv2.erode(255-img, vertical_kernel, anchor=(-1, 1))
		vertical = cv2.bitwise_or(cv2.erode(ROI, vertical_kernel, anchor=(0, 0)),
								  cv2.erode(ROI, vertical_kernel, anchor=(0, -1)))
		vertical = cv2.dilate(vertical, np.array([1, 1, 0], np.uint8), iterations=300)
		cv2.imwrite('w.jpg', vertical)

		# 提取表格顶点
		crosspoints = np.bitwise_and(horizon, vertical)  # 横竖线交点
		y, x = np.where(crosspoints == 255)  # 所有交点的x, y值
		x = sorted(list(set(x)))
		self.x = [x1 for x1, x2 in zip(x, x[1:]) if abs(x2 - x1) > 10]
		y = sorted(list(set(y)))
		self.y = [y1 for y1, y2 in zip(y, y[1:]) if abs(y2 - y1) > 10]

	def ocr(self):

		# 裁剪单元格黑边
		def removeLines(img):
			img2gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
			t1, img2mono = cv2.threshold(img2gray, 230, 255, cv2.THRESH_BINARY)
			h, w = img2gray.shape
			# 检测表格上横线的下缘部分
			for i in range(h // 2):
				if np.mean(img2mono[i + 1, :]) - np.mean(img2mono[i, :]) >= 60:
					hor_top = i + 1
					break
			else:
				hor_top = None

			for j in range(h - 1, h // 2, -1):
				if np.mean(img2mono[j - 1, :]) - np.mean(img2mono[j, :]) >= 60:
					hor_bottom = j - 1
					break
			else:
				hor_bottom = None

			for l in range(w // 2):
				if np.mean(img2mono[:, l + 1]) - np.mean(img2mono[:, l]) >= 60:
					ver_left = l + 1
					break
			else:
				ver_left = None

			for r in range(w - 1, w // 2, -1):
				if np.mean(img2mono[:, r - 1]) - np.mean(img2mono[:, r]) >= 60:
					ver_right = r - 1
					break
			else:
				ver_right = None
			# print(hor_top, hor_bottom, ver_left, ver_right)
			return img[hor_top:hor_bottom, ver_left:ver_right]

		cnt = 1
		for y1, y2 in zip(self.y[:], self.y[1:]):
			line = []
			for x1, x2 in zip(self.x, self.x[1:]):
				# 裁剪单元格黑边
				to_ocr = removeLines(self.image[y1:y2, x1-10:x2])# , cv2.COLOR_BGR2HSV
				a = pytesseract.image_to_string(to_ocr, lang='chi_sim')
				# cv2.imwrite(f"{cnt}.jpg", to_ocr)
				# cnt += 1
				line.append(a)
			self.ocr_result.append(line)

	def createExcel(self):
		app = win32.gencache.EnsureDispatch('Excel.Application')
		ss = app.Workbooks.Add()
		sh = ss.ActiveSheet
		app.Visible = True
		for line, i in enumerate(self.ocr_result, 1):
			for column, j in enumerate(i, 1):
				sh.Cells(line, column).Value = j
		ss.Save()
		ss.Close(True)
		app.Application.Quit()

	def markCorners(self):
		for i in self.y:
			for j in self.x:
				cv2.circle(self.image, (j, i), 20, 0, 5)
		cv2.imwrite('result.jpg', self.image)

def main():
	trans = Pdf2xlsx(input("The path of pdf: "))
	print("\rExtracting Lines of Tablets...", end="")
	trans.extractLines()
	print("\rOcring...", end="")
	trans.ocr()
	print("\rWriting Excel...", end="")
	trans.createExcel()
	print("\rDone!")
	# trans.markCorners()

main()
