# Invoice-Generator-Template

This is an Excel-based Invoice Template automated using VBA macros. It allows users to generate and export invoices as PDF files with dynamic file names and specified paths.

## Features

- Auto-fetches **Invoice Number** and **Client Name** from the worksheet
- Generates a **PDF invoice** with a custom name (e.g., `12345_ClientName.pdf`)
- Saves the file to a predefined location on your computer
- Easy to use with a single macro button click
- Built-in compatibility for reuse with different invoice data

## How It Works

1. Open the Excel file `Invoice_Template1.xlsm`
2. Fill in:
   - **Cell G5** with the Invoice Number
   - **Cell D5** with the Client Name
3. Click the **"Export PDF"** button (or run the `PDF()` macro)
4. A PDF will be saved to the specified file path (you can customize this in the code)

## Requirements

- Microsoft Excel with **macros enabled**
- VBA support (built-in with Excel for Windows)

## File Structure

- `Invoice_Template1.xlsm` – The macro-enabled Excel file
- `Module1` – Contains the `PDF()` macro code

## Sample Code

Sub PDF()

' Declare variables
Dim Invoice_number As Long ' Stores the invoice number from cell G5
Dim Name As String         ' Stores the name from cell D5
Dim File_path As String    ' Stores the file path where the PDF will be saved
Dim File_name As String    ' Stores the final file name for the PDF

' Assign values from the worksheet to variables
Invoice_number = Range("G5") ' Get invoice number from cell G5
Name = Range("D5")           ' Get name from cell D5

' Set the file path where the PDF will be saved
File_path = "C:\Users\Public\"

' Set the file name for the PDF using the name from cell D5
File_name = Name

' Export the active sheet as a PDF to the specified path and file name
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, ignoreprintareas:=False, Filename:=File_path & File_name

End Sub


