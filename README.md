# Mini Proiect Banca Transilvania

## About
RPA project for Banca Transilvania using Python, Tesseract, OpenCV. Automate reading of invoices and comparing with values from Excel. Send email when process is complete.

## Requirements

- Read specific values from columns and rows of an Excel workbook;
- Based on a numeric value, find the right invoice and read region(s)(ROI) of interest.
- For each invoice:
  - get data from ROI and compare with specific values from row and columns of Baza_clienti.xlsx;
  - if comparison is false, write that in Rezultat_VerificareFacturi_10.10.10;
  - get product number(s) from a field and add it to Cod Produs column in Baza_clienti.xlsx;
  - check if signature is present.
- Send an email that the process is completed.

## Installation

- [Pytesseract](https://pypi.org/project/pytesseract/) 
- [Install Tesseract for Windows](https://github.com/UB-Mannheim/tesseract/wiki)
- [Install Poppler for PDF files - Windows](https://blog.alivate.com.au/poppler-windows/)

##### Local installation
i. First clone the project
   https://github.com/Walachul/MiniProiect_Banca_Transilvania.git 

ii. Make sure you have Python 3.9.5 installed

iii. Create a virtual environment in Windows:
- Navigate to where the project folder is and run:

        python -m venv venv 
    Activate the venv:
- Navigate to venv and inside run:

        C:\Python\Example\venv>Scripts\activate

iv. If successful, you should see the name of the virtual environment in curly braces in the front of the path:

    (venv) C:\Python\Example\venv>
v. To install packages:

 Navigate to the home folder and type:

     pip install requirements.txt

**Please note that for sending email, environment variables need to be set.**

To setup the following variables in Windows environment variables can be done like this:

    Navigate: Control Panel > System > Advanced system settings > Environment Variables > 
    add new > EMAIL_USER; EMAIL_PASSWORD
