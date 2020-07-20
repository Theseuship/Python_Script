import tkinter as Tk
from win32com.client import Dispatch
from tkinter.simpledialog import askstring
from tkinter.messagebox import showerror
from tkinter.filedialog import askdirectory
import itertools
import random
from collections import OrderedDict
import json



def generateDict(d):
	entry = OrderedDict(**{key: {i: False for i in l} for key, l in d.items()})
	for k in d:
		entry[k].update({"记录": ""})
		entry[k].update({False: False})
	return entry

def sort_type(k):
	if k in test_data1.keys():
		return ("test_data1", k)
	elif k in test_data2.keys():
		return ("test_data2", k)
	else:
		return k

test_data1 = {j: [chr(96+i)*i for i in range(1, random.randrange(2, 8))] for j in range(random.randrange(1, 10))}
test_data2 = {chr(64+i): [f"{j}"*j for j in range(1, random.randrange(2, 8))] for i in range(random.randrange(1, 10))}

class Layout:
	def show(self, d, name, entry, frame_list, frame_no, basic_info):
		for i in frame_list[frame_no].grid_slaves():
			i.destroy()
		if False in d[name][entry].keys():
			l = Tk.Label(frame_list[frame_no], text=f"{name} - {entry}")
			l.grid()
			a = Tk.Checkbutton(frame_list[frame_no], text=f"有无{entry}资料")
			def click(event):
				d[name][entry].__delitem__(False)
				l.destroy()
				a.destroy()
				self.show(d, name, entry, frame_list, frame_no, basic_info)
			a.bind("<Button-1>", click)
			a.grid()
		else:
			ticked_list = []
			Tk.Label(frame_list[frame_no], text=f"{name} - {entry}").grid()

			def modify(key, index):
				def inner():
					d[name][entry][key] = bool(ticked_list[index].get())
				return inner

			def update_record(event):
				d[name][entry][key] = t.get("0.0", "end")

			def reset():
				d[name][entry].clear()
				d[name].update(generateDict({entry: test_data1[entry]} if name == "test_data1" else generateDict({entry: test_data2[entry]})))
				self.show(d, name, entry, frame_list, frame_no, basic_info)

			for index, key in enumerate(d[name][entry].keys()):
				ticked_list.append(Tk.IntVar())
				if key != "记录":
					ticked_list[index].set(d[name][entry][key])
					Tk.Checkbutton(frame_list[frame_no], text=key, variable=ticked_list[index], command=modify(key, index)).grid()
				else:
					t = Tk.Text(frame_list[frame_no], height=3)
					t.insert("0.0", d[name][entry][key])
					t.bind("<KeyRelease>", update_record)
					t.grid()
			a = Tk.Button(frame_list[frame_no], text="Reset")
			a.configure(command=reset)
			a.grid()

class P(Layout):
	def show(self, d, name, entry, frame_list, frame_no, basic_info):
		super().show(d, name, entry, frame_list, frame_no, basic_info)
		if not False in d[name][entry].keys():
			Tk.Label(frame_list[frame_no], text="计算公式").grid()
			e2 = Tk.Entry(frame_list[frame_no])
			e2.grid()
			def cal(e, e2):
				def inner():
					try:
						scope = {}
						result = eval(e2.get(), globals=scope)
					except:
						showerror("输入错误", message="输入公式错误，请重新输入！")
						e2.set()
					Tk.Label(frame_list[frame_no], text=f"{result}").grid()
				return inner
			Tk.Button(frame_list[frame_no], text="公式运行结果", command=cal(e, e2)).grid()

class P2(Layout):
	def show(self, d, name, entry, frame_list, frame_no, basic_info):
		super().show(d, name, entry, frame_list, frame_no, basic_info)
		path = askdirectory("Excel保存位置") + "test.xlsx"
		if not False in d[name][entry].keys():
			def openExcel():
				try:
					app.Workbooks.Open(path)
				except:
					app = Dispatch("Excel.Application")
					app.Workbooks.Open(path)
				app.Visible = True
				sheet = app.ActiveSheet
				for index, i in enumerate(basic_info["name"]):
					sheet.Cells(index, "A").Value = i

			def readExcel():
				try:
					try:
						worksheet = app.Workbooks.Open(path)
					except:
						app = Dispatch("Excel.Application")
						worksheet = app.Workbooks.Open(path)
					sheet = app.ActiveSheet
					t = tk.Text(frame_list[frame_no])
					t.grid()
					for index, i in enumerate(basic_info["name"]):
						t.insert("end", f"{{{str(i)}: {sheet.Cells(index, 'A').Value}}}\n")
				finally:
					worksheet.Close(False)
					app.Quit()
			Tk.Button(frame_list[frame_no], text="Open Excel", command=openExcel).grid()
			Tk.Button(frame_list[frame_no], text="Read Data", command=readExcel).grid()

class GUI:
	def __init__(self):
		self.window = Tk.Tk()
		self.window.geometry("800x800")
		self.navigator = Tk.Frame(self.window, bg="cyan")
		self.frame = [Tk.Frame(self.window)]
		self.confirm = Tk.Frame(self.window, height=2)
		self.window.grid_columnconfigure(1, weight=1)
		self.window.grid_rowconfigure(1, weight=1)
		self.navigator.grid(row=0, column=0, sticky="ns")
		self.confirm.grid(row=1, column=1, rowspan=1)
		self.frame[0].grid(row=0, column=1)
		self.basic_info = {"name": [],
						   "test1": {k: 0 for k in (f"Test{i}" for i in range(5))},
						   "test2": {k: 0 for k in (f"Test{i}" for i in range(5))}}
		self.entry = {"test_data1": generateDict(test_data1)}

	def refresh(self, n):
		for i in self.frame:
			i.destroy()
		self.frame = [Tk.Frame(self.window) for i in range(n)]
		self.window.grid_columnconfigure(n, weight=1)
		for index, i in enumerate(self.frame, 1):
			i.grid(row=0, column=index)
		self.confirm.destroy()
		self.confirm = Tk.Frame(self.window, height=2)
		self.confirm.grid(row=1, column=1, rowspan=1)

	def basic(self):
		self.refresh(1)
		Tk.Label(self.frame[0], text="名字").grid(row=0, column=0)
		name = Tk.Entry(self.frame[0])
		name.grid(row=0, column=1, columnspan=4, sticky="ew")
		name.insert(0, " & ".join(self.basic_info["name"]))

		def display_checkbutton(entry, key, lst, index):
			def inner():
				self.basic_info[entry][key] = lst[index].get()
			return inner()

		for row, entry in enumerate(("test1", "test2")):
			Tk.Label(self.frame[0], text=entry).grid(row=1+row, column=0)
			tmp = [Tk.BooleanVar() for i in self.basic_info[entry]]
			for index, k in enumerate(self.basic_info[entry].keys()):
				tmp[index].set(self.basic_info[entry][k])
				Tk.Checkbutton(self.frame[0], text=k, var=tmp[index], command=display_checkbutton(entry, k, tmp, index)).grid(row=1+row, column=1+index)

		def basic_info_update():
			self.basic_info["name"] = [i.title().strip() for i in name.get().split("&")]
			for i in self.basic_info["name"]:
				self.entry[i] = generateDict(test_data2)
			if self.basic_info["test1"]["Test3"]:
				key = random.choice(test_data1.keys())
				self.entry["test_data1"][key].setdefault("Bonus", False)
				self.entry["test_data1"][key].move_to_end("记录")
			self.displayCheckList()

		Tk.Button(self.confirm, text="Confirm", command=basic_info_update).grid()

	def displayCheckList(self):
		self.refresh(1)
		order = list(itertools.chain(test_data1, test_data2))
		random.shuffle(order)
		checklist = (sort_type(i) for i in order)
		k = next(checklist)
		def test(k):
			if k[0] == "test_data1":
				self.refresh(1)
				if k[1] == 1:
					P().show(self.entry, k[0], k[1], self.frame, 0, self.basic_info)
				else:
					Layout().show(self.entry, k[0], k[1], self.frame, 0, self.basic_info)
			else:
				self.refresh(len(self.basic_info["name"]))
				for index, name in enumerate(self.basic_info["name"]):
					if k[1] == "A":
						P2().show(self.entry, name, k[1], self.frame, index, self.basic_info)
					else:
						Layout().show(self.entry, name, k[1], self.frame, index, self.basic_info)

			def tmp():
				for i in self.confirm.grid_slaves():
					i.destroy()
				try:
					test(next(checklist))
				except StopIteration:
					self.summmary()
			Tk.Button(self.confirm, text="Next", command=tmp).grid()

		for i in (sort_type(i) for i in ["Basic Info", *order, "汇总"]):
			def d(i):
				def inner():
					test(i)
				return inner
			if i == "Basic Info":
				Tk.Button(self.navigator, text="{:\u3000^5}".format("基本信息"), command=self.basic).grid()
			elif i == "汇总":
				Tk.Button(self.navigator, text="{:\u3000^5}".format("汇总"), command=self.summary).grid()
			else:
				Tk.Button(self.navigator, text=f"{i[1]:\u3000^5}", command=d(i)).grid()

		test(k)

	def summary(self):
		self.refresh(1)
		t = Tk.Text(self.frame[0])
		t.grid()
		t.insert("end", json.dumps(self.basic_info) + "\n")
		t.insert("end", json.dumps(self.entry) + "\n")

		def copy():
			self.window.clipboard_clear()
			self.window.clipboard_append(t.get(0.0, "end"))
			self.window.update()

		Tk.Button(self.confirm, text="Copy", command=copy).grid()

m = GUI()
m.basic()
m.window.mainloop()