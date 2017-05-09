import tkFileDialog
from Tkinter import *
import logic
import os

def gui():
    """make the GUI version of this command that is run if no options are
    provided on the command line"""

    def button_go_callback():
        """ what to do when the "Go" button is pressed """
        ipblock1 = ipblock.get()
        if ipblock1.endswith(".xlsx"):
            pass
        else:
            statusText.set("Filename must end in `.xlsx'")
            message.configure(fg="red")
            return
        zip_files_path1 = zip_files_path.get()


        reportspath = zip_files_path1 + "\\Company PDF Reports"
        if not os.path.exists(reportspath):
            os.makedirs(reportspath)

        #call for the logic function
        if logic.generate_pdf_reports(ipblock1, zip_files_path1, reportspath) is None:
            statusText.set("Check " + reportspath + " for the generated reports")
            message.configure(fg="black")

        else:
            statusText.set("Errors in generating reports, try again")
            message.configure(fg="black")
            return

    def button_browse_callback():
        """ What to do when the Browse button is pressed """
        filename = tkFileDialog.askopenfilename()
        ipblock.delete(0, END)
        ipblock.insert(0, filename)

    def button_browse_callback2():
        """ What to do when the Browse button2 is pressed """
        filename2 = tkFileDialog.askdirectory()
        zip_files_path.delete(0, END)
        zip_files_path.insert(0, filename2)



    root = Tk()
    root.wm_title("Generate Reports")
    frame = Frame(root)
    frame.pack()

    statusText = StringVar(root)
    statusText.set("Press Browse button or enter the file name in the entry, "
                   "then press the Generate Reports Button")

    label = Label(root, text="IP Allocation, Excel file")
    label.pack()
    ipblock = Entry(root, width=60)
    ipblock.pack()
    
    # Button to browse for the file
    button_browse = Button(root,
                           text="Browse",
                           command=button_browse_callback)
    button_browse.pack(pady=8)


    label = Label(root, text="Folder for the zipped files: ")
    label.pack()
    zip_files_path = Entry(root, width=60)
    zip_files_path.pack()

    # Button to browse the folder
    button_browse2 = Button(root,
                           text="Browse",
                           command=button_browse_callback2)


    button_browse2.pack(pady=8)


    separator = Frame(root, height=2, bd=1, relief=SUNKEN)
    separator.pack(fill=X, padx=5, pady=5)

    # Button for generating reports
    button_go = Button(root,
                       text="Generate Reports",
                       command=button_go_callback)
    button_exit = Button(root,
                         text="Exit",
                        command=sys.exit)


    button_go.pack(pady=8)
    button_exit.pack(pady=8)

    separator = Frame(root, height=2, bd=1, relief=SUNKEN)
    separator.pack(fill=X, padx=5, pady=5)

    message = Label(root, textvariable=statusText)
    message.pack(pady=10)

    mainloop()

gui()