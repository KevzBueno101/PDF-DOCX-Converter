from pdf2docx import Converter
import customtkinter as ctk
from PIL import Image, ImageTk
from tkinter import PhotoImage
from tkinter.filedialog import askopenfilename, asksaveasfilename
from CustomTkinterMessagebox import CTkMessagebox
import os
import sys
import subprocess

# Initialize the app
app = ctk.CTk()
app.geometry("500x300")
app.resizable(False, False)
app.title("PDF-DOCX Converter")
app.iconbitmap(r"assets/logo.ico")

def resource_path(rel):
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base, rel)
app.iconbitmap(resource_path("assets/logo.ico"))

pdf_file_path = None

# Function to import file
def import_file():
    global pdf_file_path
    pdf_file_path = askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if pdf_file_path:
        header.configure(text=f"Selected:\n{os.path.basename(pdf_file_path)}")

# Function to convert PDF to DOCX
def convert_pdf():
    if not pdf_file_path:
        CTkMessagebox.messagebox(
            title='Error',
            text='No file imported.\nPlease import your PDF file first.',
            sound='on',
            button_text='OK'
        )
        return

    docx_file = asksaveasfilename(defaultextension=".docx",
                                   filetypes=[("Word Documents", "*.docx")],
                                   initialfile=os.path.basename(pdf_file_path).replace(".pdf", ".docx"))
    if not docx_file:
        return

    try:
        cv = Converter(pdf_file_path)
        cv.convert(docx_file, start=0, end=None)
        cv.close()

        view = CTkMessagebox.messagebox(
            title='Finished',
            text='File converted successfully!',
            sound='on',
            button_text='Ok'
        )

    except Exception as e:
        CTkMessagebox.messagebox(title='Error', text=str(e), sound='on', button_text='OK')

# Load background image
bg_image = Image.open("assets/bg.png")
bg_image = bg_image.resize((500, 300))
bg_photo = ImageTk.PhotoImage(bg_image)

# Background label
bg_label = ctk.CTkLabel(app, image=bg_photo, text="")
bg_label.place(x=0, y=0, relwidth=1, relheight=1)

# Header label
header = ctk.CTkLabel(app, text='PDF -> DOCX Converter', font=('Arial', 16, 'bold'))
header.pack(pady=10)

# Import button
pdf_btn = ctk.CTkButton(app, text='Import File', width=70, command=import_file)
pdf_btn.place(x=100, y=240)

# Convert button
convert_btn = ctk.CTkButton(app, text='Convert', width=70, command=convert_pdf)
convert_btn.place(x=300, y=240)

# Run app
app.mainloop()
