from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from docxcompose.composer import Composer
from docx import Document
import os
import ntpath


#debug
DEBUG_MODE = False

# globals
master = 0
documents = []
ui = []
path = ""
root = Tk()
root.geometry("330x400")
# r = IntVar()

if DEBUG_MODE:
    files = []
    for letter in [chr(i) for i in range(ord('a'), ord('g')+1)]:
        files.append(os.path.join("debug", str(letter)+".docx"))
else:
    files = filedialog.askopenfilenames(initialdir="/", title="search for file",
                                        filetypes=(("docx", "*.docx"), ("doc", "*.doc"), ("all", "*")))
    path = os.path.dirname(os.path.realpath(files[0]))


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
        name = ntpath.basename(documents[i]["path"])
        uiline = ui[i]
        uiline["filename"]["text"] = name[:40] if len(name) > 40 else name
        uiline["position"]["text"] = str(i)


def draw():
    for uiline in ui:
        position = ui.index(uiline)
        uiline["filename"].grid(row=position, column=0, sticky=NSEW)
        uiline["position"].grid(row=position, column=2, sticky=NSEW)
        uiline["up"].grid(row=position, column=3)
        uiline["down"].grid(row=position, column=4)


def append():
    masterDoc = Document(documents[0]["path"])
    masterDoc.add_page_break()
    composer = Composer(masterDoc)
    for doc in range(1, len(documents)):
        docu = Document(documents[doc]["path"])
        docu.add_page_break()
        composer.append(docu)
        doc += 1
    composer.save(os.path.join(path, "combined.docx"))
    messagebox.showinfo("POTEITOES!", "Ho finito!")


# main

# init documents
for file in files:
    documents.append({"path": file})


# init UI

main_frame = Frame(root, width=290, height=400)
main_frame.place(x=10, y=10)

append_button = Button(main_frame, text="Append", command=append)

canvas = Canvas(main_frame)
sb = Scrollbar(main_frame, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=sb.set)
content_frame = Frame(canvas)

for i in range(len(documents)):
    name = ntpath.basename(documents[i]["path"])
    ui.append({
        "filename": Label(content_frame, text=(name[:40] if len(name) > 40 else name)),
        "position": Label(content_frame, text=str(i)),
        "up": Button(content_frame, text="up", command=lambda index=i: up_clicked(index)),
        "down": Button(content_frame, text="down", command=lambda index=i: down_clicked(index))
    })

canvas.grid(row=1, column=0, sticky='nsew')
sb.grid(row=1, column=1, sticky='ns')
append_button.grid(row=0, column=0, sticky='nesw', columnspan=2)
canvas.create_window((0, 0), anchor="nw", window=content_frame)
content_frame.bind("<Configure>", lambda a: canvas.configure(scrollregion=canvas.bbox("all"), width=290, height=400))

draw()


root.mainloop()
