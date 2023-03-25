import pandas as pd
from PyPDF2 import PdfReader
from tkinter import Tk, simpledialog
from tkinter.filedialog import askopenfilename

# Create a Tkinter root window (it will not be shown)
root = Tk()
root.withdraw()

# Ask the user to select the xlsx file
xlsx_file = askopenfilename(title='Select the xlsx file', filetypes=[('Excel Files', '*.xlsx')])

# Read the xlsx file
df = pd.read_excel(xlsx_file)

# Get the data from the 'quantity' column
quantity_data = df['Order Quantity'].tolist()

# Ask the user if they want to select a PDF file or enter text manually
pdf_or_text = simpledialog.askstring('PDF or Text', 'Enter "PDF" to select a PDF file or "Text" to enter text manually:')

if pdf_or_text.lower() == 'pdf':
    # Ask the user to select the PDF file
    pdf_file = askopenfilename(title='Select the PDF file', filetypes=[('PDF Files', '*.pdf')])

    # Open the PDF file
    with open(pdf_file, 'rb') as f:
        pdf = PdfReader(f)
        # Convert the PDF to plain text
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
else:
    # Ask the user to enter text manually
    text = simpledialog.askstring('Enter Text', 'Enter the text:')

# Search for the data from the 'quantity' column within the text
missing_values = []
for data in quantity_data:
    a = ' '+str(data)
    b = ' '+str(data)+'EA'
    c = ' '+str(data)+'FT'
    if a not in text and b not in text and c not in text:
        missing_values.append(data)

if not missing_values:
    print('All values from the quantity column are present in the text')
else:
    print(f'The following values from the quantity column are not present in the text: {missing_values}')