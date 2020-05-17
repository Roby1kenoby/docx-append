from tkinter import *
from docxcompose.composer import Composer
from docx import Document
from tkinter import filedialog
import ntpath

root = Tk()
files = filedialog.askopenfilenames(initialdir="/", title="search for file",
                                    filetypes=(("docx", "*.docx"), ("doc", "*.doc"), ("all", "*")))


def clicked(value):
    print(value)


documents = []


def up(u):
    documents[u]["index"] -= 1
    documents[u-1]["index"] += 1
    documents.sort(key=lambda k: k["index"])
    print(str(documents[u]["index"]) + ' ' + str(documents[u-1]["index"]))

    draw()
    return


def down(d):
    print(str(d))
    return


def draw():
    for document in documents:
        document["label"].grid(row=documents.index(document), column=0)
        document["radio"].grid(row=documents.index(document), column=1)
        document["label2"].grid(row=documents.index(document), column=2)
        document["button"].grid(row=documents.index(document), column=3)
        document["button2"].grid(row=documents.index(document), column=4)


m = IntVar()

for file in files:
    documents.append({"path": file, "index": files.index(file), "master": (1 if files.index(file) == 0 else 0),
                      "label": Label(root, text=ntpath.basename(file)),
                      "radio": Radiobutton(root, variable=m, value=files.index(file), command=lambda: clicked(m.get())),
                      "label2": Label(root, text=files.index(file)),
                      "button": Button(root, text="up", command=lambda u=files.index(file): up(u),
                                       state=(DISABLED if files.index(file) == 0 else NORMAL)),
                      "button2": Button(root, text="down",
                                        state=(DISABLED if files.index(file) == len(files) - 1 else NORMAL),
                                        command=lambda d=files.index(file): down(d))
                      })

draw()

root.mainloop()
