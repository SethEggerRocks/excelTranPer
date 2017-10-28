import tkinter as tk
import datetime
from excelPlayin import *
import time

now = datetime.datetime.now()
today = (str(now.month) + "/" + str(now.day) + "/" + str(now.year))


class Application(tk.Frame):


# This is the widget!!
    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.hi_there = tk.Button(self)
        self.hi_there["text"] = "\n Press Here To Get Stats!\n"
        self.hi_there["command"] = self.say_hi
        self.hi_there.pack(side="top")
        #self.quit = tk.Button(self, text="QUIT", fg="red",
        #                      command=root.destroy)
        #self.quit.pack(side="bottom")   # This closes window when pressed

    def say_hi(self):
        glcNumbers()
        droegeNumbers()
        detoxNumbers()
        wrapNumbers()
        doughertyNumbers()
        print(today)
        time.sleep(3)
        app.destroy()
        exit()



root = tk.Tk()
app = Application(master=root)
app.mainloop()
