# imports for basic app function
from logging import exception
from sys import exit

# tkinter imports
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

# imports for PDF handling
from PyPDF2 import PdfFileReader, PdfFileWriter
from pathlib import Path

# Function to rebuild ListBox
# Called when filenames list is altered
# by remove, add, move up/down, or reset
def rebuild_lb():
    global lb, filenames, lb_num

    # Clear existing ListBox items
    lb.delete(0, 'end')

    # Update number of files (for label)
    # and refill ListBox with updated filenames
    num = 0
    if filenames != None:
        for name in filenames:
            lb.insert('end', name)
        num = len(filenames)
    lb_num.configure(text=str(num) + " files loaded")


def add_files():
    global filenames

    if filenames == None:
        filenames = filedialog.askopenfilenames(
            title="Select files to merge",
            filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
    else:
        f = filedialog.askopenfilenames(
            title="Select files to merge",
            filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
        filenames += f

    rebuild_lb()


def rem_files():
    global filenames, lb
    selected = lb.curselection()

    revised = []
    for i in range(len(filenames)):
        if i < selected[0] or i > selected[len(selected) - 1]:
            revised.append(filenames[i])

    if len(revised) > 0:
        filenames = tuple(revised)
    else:
        filenames = None

    rebuild_lb()


def move_up():
    global filenames, lb

    selected = lb.curselection()
    revised = []

    if filenames == None or len(selected) == 0:
        return

    for i in range(len(filenames)):
        if i != selected[0] - 1:
            revised.append(filenames[i])

    if selected[0] != 0:
        revised.insert(selected[len(selected) - 1], filenames[selected[0] - 1])

    if len(revised) > 0:
        filenames = tuple(revised)
    else:
        filenames = None

    rebuild_lb()


def move_down():
    global filenames, lb

    selected = lb.curselection()
    revised = []

    if filenames == None or len(selected) == 0:
        return

    for i in range(len(filenames)):
        if i != selected[len(selected) - 1] + 1:
            revised.append(filenames[i])

    if selected[len(selected) - 1] != len(filenames) - 1:
        revised.insert(selected[0], filenames[selected[len(selected) - 1] + 1])

    if len(revised) > 0:
        filenames = tuple(revised)
    else:
        filenames = None

    rebuild_lb()


def merge_files():
    global filenames

    if filenames == None or len(filenames) < 2:
        messagebox.showerror(
            title="PDF Utility", 
            message="You must have at least 2 files loaded in order to merge")
        return

    fn = filedialog.asksaveasfilename(title="Select file", filetypes=(
        ("PDF files", "*.pdf"), ("all files", "*.*")))
    pdf_writer = PdfFileWriter()

    try:
        for pdf_path in filenames:
            input_pdf = PdfFileReader(str(pdf_path))

            if input_pdf.getNumPages() == 0:
                raise Exception(
                    "You can't merge empty PDFs ('" + str(pdf_path) + "' is empty)")

            pdf_writer.appendPagesFromReader(input_pdf)

        with Path(fn).open(mode="wb") as output_file:
            pdf_writer.write(output_file)

        messagebox.showinfo(title="PDF Utility", message=str(
            len(filenames)) + " files were merged into '" + fn + "' successfully!")

    except Exception as e:
        messagebox.showerror(title="PDF Utility",
                             message="Merge Failed: " + str(e))


def reset():
    global filenames
    filenames = None

    rebuild_lb()


def main():
    global filenames, lb, lb_num
    filenames = None

    dark_bg = "#192734"
    dark_gr = "#22303C"

    # Initialise window properties
    window = Tk()
    window.geometry("400x480")
    window.resizable(False, False)
    window.configure(bg=dark_bg)
    window.title('PDF Merger')

    # Main heading text
    title = Label(window, text="Select PDF files to merge", 
                  font=("Helvetica", 24, "bold"), bg=dark_bg, 
                  fg="white", padx=30, pady=25)

    # Listbox for pdf files names
    lb = Listbox(window, listvariable=filenames, selectmode="extended",
                 width=40, bg=dark_gr, bd=0, borderwidth=0)

    # Listbox buttons
    lb_add = Button(window, text="+", command=add_files, width=2,
                    height=1, highlightbackground=dark_bg)
    lb_sub = Button(window, text="-", command=rem_files, width=2,
                    height=1, highlightbackground=dark_bg)
    lb_up = Button(window, text="▲", command=move_up, width=2,
                   height=1, highlightbackground=dark_bg)
    lb_down = Button(window, text="▼", command=move_down,
                     width=2, height=1, highlightbackground=dark_bg)

    num = 0
    if filenames != None:
        num = len(filenames)

    # Show the number of files queued for merging
    lb_num = Label(window, text=str(num) + " files loaded", font=("Helvetica", 14),
                   bg=dark_bg, fg="white", height=2)

    # Bottom 3 buttons
    button_merge = Button(window, text="MERGE FILES", command=merge_files,
                          height=2, font=("Helvetica", 14, "bold"), 
                          highlightbackground=dark_bg)
    button_exit = Button(window, text="Exit", command=exit, width=5,
                          highlightbackground=dark_bg, font=("Helvetica", 14))
    button_reset = Button(window, text="Reset", command=reset,
                          width=5, highlightbackground=dark_bg, 
                          font=("Helvetica", 14))

    # Put components on the grid
    title.grid(column=1, columnspan=6, row=0, pady=10)

    lb_add.grid(column=1, row=1, sticky="e")
    lb_sub.grid(column=2, row=1, sticky="w")
    lb_up.grid(column=5, row=1, sticky="e")
    lb_down.grid(column=6, row=1, sticky="w")

    lb.grid(column=1, columnspan=6, row=2, pady=10)
    lb_num.grid(column=1, columnspan=6, row=3, sticky="s")

    button_merge.grid(column=2, columnspan=4, row=4, sticky="ew", pady=10)

    button_exit.grid(column=2, columnspan=2, row=5, sticky="w")
    button_reset.grid(column=4, columnspan=2, row=5, sticky="e")

    # Assign columns equal weight
    window.grid_columnconfigure((1, 2, 3, 4, 5, 6), weight=1)

    window.mainloop()


if __name__ == "__main__":
    main()
