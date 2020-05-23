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
# root.geometry("290x400")
# r = IntVar()
files = filedialog.askopenfilenames(initialdir="/", title="search for file",
                                    filetypes=(("docx", "*.docx"), ("doc", "*.doc"), ("all", "*")))

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
    container.pack()
    canvas.pack(side="left", fill="both", expand=TRUE)
    sb.pack(side="right", fill="y")


def append():
    masterDoc = Document(documents[0]["path"])
    composer = Composer(masterDoc)
    for doc in range(1, len(documents)):
        docu = Document(documents[doc]["path"])
        composer.append(docu)
        doc += 1
    composer.save("combined.docx")


# main

# init documents
for file in files:
    documents.append({"path": file})


# init UI
containerPadre = Frame(root)
canvas = Canvas(containerPadre)
scrollable_frame = Frame(canvas)
sb = Scrollbar(containerPadre, orient="vertical", command=canvas.yview)
scrollable_frame.bind("configure", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.configure(yscrollcommand=sb.set)
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

# doc_frame = Listbox(root, yscrollcommand=sb.set)
append_button = Button(root, text="Append", command=append)

for i in range(len(documents)):
    ui.append({
        "filename": Label(scrollable_frame, text=ntpath.basename(documents[i]["path"])),
        # "ismaster": Radiobutton(root, variable=r, value=i, command=lambda: master_selected(r.get())),
        "position": Label(scrollable_frame, text=str(i)),
        "up": Button(scrollable_frame, text="up", command=lambda index=i: up_clicked(index)),
        "down": Button(scrollable_frame, text="down", command=lambda index=i: down_clicked(index))
    })


#sb.config(command=doc_frame.yview())



draw()


root.mainloop()