import pandas as pd
from PyPDF2 import PdfReader
from tkinter import Tk, Label, Button, Entry, Text, END
from tkinter.filedialog import askopenfilename
import webbrowser

def select_xlsx():
    # Ask the user to select the xlsx file
    xlsx_file = askopenfilename(title='Select the xlsx file', filetypes=[('Excel Files', '*.xlsx')])

    # Update the xlsx label with the selected file name
    xlsx_label['text'] = f'Selected xlsx file: {xlsx_file}'

    # Read the xlsx file
    df = pd.read_excel(xlsx_file)

    # Get the data from the 'quantity' column
    global quantity_data
    quantity_data = df['Order Quantity'].tolist()

def select_pdf():
    # Ask the user to select the PDF file
    pdf_file = askopenfilename(title='Select the PDF file', filetypes=[('PDF Files', '*.pdf')])

    # Open the PDF file
    with open(pdf_file, 'rb') as f:
        pdf = PdfReader(f)
        # Convert the PDF to plain text
        text = ''
        for page in pdf.pages:
            text += page.extract_text()

    # Search for the data from the 'quantity' column within the text
    search_text(text)

def enter_text():
    # Get the text entered by the user
    text = text_entry.get('1.0', END)

    # Search for the data from the 'quantity' column within the text
    search_text(text)

def search_text(text):
    # Clear the output text box
    output_text.delete('1.0', END)

    # Search for the data from the 'quantity' column within the text
    missing_values = []
    for data in quantity_data:
        a = ' '+str(data)
        b = str(data)+'EA'
        c = str(data)+'ea'
        d = str(data)+'FT'
        e = str(data)+'ft'
        if a not in text and b not in text and c not in text and d not in text and e not in text:
            missing_values.append(data)

    if not missing_values:
        output_text.insert(END, 'All values from the quantity column are present in the text')
    else:
        output_text.insert(END, f'The following values from the quantity column are not present in the text: {missing_values}')

def open_onlineocr():
    webbrowser.open('https://www.onlineocr.net/')

# Create a Tkinter root window
root = Tk()
root.title('PDF or Text')

# Create a label and button for selecting the xlsx file
xlsx_label = Label(root, text='Select the xlsx file:')
xlsx_label.pack()
xlsx_button = Button(root, text='Select xlsx', command=select_xlsx)
xlsx_button.pack()

# Create a label and buttons for selecting a PDF file or entering text manually
pdf_text_label = Label(root, text='Select a PDF file or enter text manually:')
pdf_text_label.pack()
pdf_button = Button(root, text='Select PDF', command=select_pdf)
pdf_button.pack()
text_button = Button(root, text='Enter Text', command=enter_text)
text_button.pack()

# Create a text box for entering text manually
text_entry = Text(root)
text_entry.pack()

# Create a label and text box for displaying the output
output_label = Label(root, text='Output:')
output_label.pack()
output_text = Text(root)
output_text.pack()

# Create a button for opening the OnlineOCR website
onlineocr_button = Button(root, text='OCR Text Converter', command=open_onlineocr)
onlineocr_button.pack()

# Run the main loop
root.mainloop()