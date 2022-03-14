from _typeshed import Self
import tkinter as tk
import excel_cell_to_list

root = tk.Tk()
root.geometry("500x600")
Button = tk.Button(root, text="let's excel", command=excel_cell_to_list())

class GUItime():
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("300x400")
        Button = tk.Button(self.root, text="in house",command=self.dakoku)
        Button2 = tk.Button(self.root, text="in office", command=self.quit)


        Button.pack()
        Button2.pack()
        self.root.mainloop()


    def quit(self):
        self.root.destroy()

    def dakoku(self):
        self.root.destroy()
        
root.mainloop()
