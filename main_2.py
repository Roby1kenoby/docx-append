from tkinter import *
from docxcompose.composer import Composer
from docx import Document
from tkinter import filedialog
import os
import ntpath

#debug
DEBUG_MODE = True

# globals
master = 0
documents = []
ui = []
root = Tk()
root.geometry("290x400")
# r = IntVar()

if DEBUG_MODE:
    files = []
    for letter in [chr(i) for i in range(ord('a'),ord('g')+1)]:
        files.append(os.path.join("debug",str(letter)+".docx"))
else:
    files = filedialog.askopenfilenames(initialdir="/", title="search for file",filetypes=(("docx", "*.docx"), ("doc", "*.doc"), ("all", "*")))

'''def master_selected(index):
    global master
    master = index
    # print_debug()
'''

def up_clicked(index):
    if index == 0:
        return
    else:
        documents[index], documents[index-1] = documents[index-1], documents[index]
        updateui()
        draw()
        # print_debug()


def down_clicked(index):
    if index == len(documents)-1:
        return
    else:
        documents[index], documents[index+1] = documents[index+1], documents[index]
        updateui()
        draw()
        # print_debug()


def print_debug():
    print("master: " + str(master))
    for doc in documents:
        print(doc["path"])


def updateui():
    for i in range(len(documents)):
        doc = documents[i]
        uiline = ui[i]
        uiline["filename"]["text"] = ntpath.basename(doc["path"])
        uiline["position"]["text"] = str(i)


def draw():
    for uiline in ui:
        position = ui.index(uiline)
        uiline["filename"].grid(row=position, column=0, sticky=NSEW)
        uiline["position"].grid(row=position, column=2, sticky=NSEW)
        uiline["up"].grid(row=position, column=3)
        uiline["down"].grid(row=position, column=4)
    append_button.pack()
    canvas.pack(fill="both", expand=True, side="left")
    sb.pack(fill="y", side="right")


def append():
    masterDoc = Document(documents[0]["path"])
    masterDoc.add_page_break()
    composer = Composer(masterDoc)
    for doc in range(1, len(documents)):
        docu = Document(documents[doc]["path"])
        docu.add_page_break()
        composer.append(docu)
        doc += 1
    composer.save("combined.docx")


# main

# init documents
for file in files:
    documents.append({"path": file})


# init UI
append_button = Button(root, text="Append", command=append)
canvas = Canvas(root)
sb = Scrollbar(canvas, orient="vertical", command=canvas.yview)
frame = Frame(canvas)

for i in range(len(documents)):
    ui.append({
        "filename": Label(frame, text=ntpath.basename(documents[i]["path"])),
        "position": Label(frame, text=str(i)),
        "up": Button(frame, text="up", command=lambda index=i: up_clicked(index)),
        "down": Button(frame, text="down", command=lambda index=i: down_clicked(index))
    })
canvas.create_window(0, 0, anchor="nw", window=frame)
canvas.update_idletasks()
canvas.configure(scrollregion=canvas.bbox("all"), yscrollcommand=sb.set)


draw()


root.mainloop()
