from tkinter import *
from docxcompose.composer import Composer
from docx import Document
from tkinter import filedialog
import ntpath

# globals
master = 0
documents = []
ui = []
root = Tk()
r = IntVar()
files = filedialog.askopenfilenames(initialdir="/", title="search for file",
                                    filetypes=(("docx", "*.docx"), ("doc", "*.doc"), ("all", "*")))

def master_selected(index):
    global master
    master = index
    print_debug()

def up_clicked(index):
    if index == 0:
        return
    else:
        documents[index], documents[index-1] = documents[index-1], documents[index]
        updateui()
        draw()
        print_debug()


def down_clicked(index):
    if index == len(documents)-1:
        return
    else:
        documents[index], documents[index+1] = documents[index+1], documents[index]
        updateui()
        draw()
        print_debug()

def print_debug():
    print("master: " + str(master))
    for doc in documents:
        print(doc["path"])

def updateui():
    for i in range(len(documents)):
        doc = documents[i]
        uiline = ui[i]
        uiline["path"]["text"] = doc["path"]
        uiline["position"]["text"] = str(i)


def draw():
    for uiline in ui:
        position = ui.index(uiline)
        uiline["path"].grid(row=position, column=0)
        uiline["ismaster"].grid(row=position, column=1)
        uiline["position"].grid(row=position, column=2)
        uiline["up"].grid(row=position, column=3)
        uiline["down"].grid(row=position, column=4)

# main

# init documents
for file in files:
    documents.append({ "path": file })


# init UI
for i in range(len(documents)):
    ui.append({
        "path": Label(root, text=""),
        "ismaster": Radiobutton(root, variable=r, value=i, command=lambda: master_selected(r.get())),
        "position": Label(root, text=""),
        "up": Button(root, text="up", command=lambda index=i: up_clicked(index)),
        "down": Button(root, text="down", command=lambda index=i: down_clicked(index))
    })

#updateui()
draw()
root.mainloop()