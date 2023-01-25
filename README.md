# Issue-of-receipt
This is a project for issuing payment receipts for a charity organization based in Brazil. The excel spreadsheet already existed and was used as the basis for the treasurer to copy the payment data of each person and fill out each receipt by hand, on paper intended for this purpose. I volunteered to develop an automation that would facilitate the work, using the same existing spreadsheet (I uploaded it). The code is written in Python and automates the process of generating and sending the receipts in the form of a PDF document through WhatsApp. It uses several libraries, including time, openpyxl, pyautogui, pyperclip, reportlab, os, and PIL. Selenium was not used to manipulate WhatsApp in this script as it would require QR code validation, which would make the automation process less efficient.

The script starts by defining four functions:

- formatar(num): Formats a number as a string with two decimal places and a thousand separator, specific to the Brazilian locale (pt_BR.UTF-8).
- enviar_msg(): Uses the pyautogui library to simulate mouse clicks and keyboard presses in order to send a message on WhatsApp. It clicks on a contact field, writes a phone number, and then attaches a PDF document to the message before sending it.
- date(data_plan): Takes a date as a parameter and assigns the year, month and day to variables.
- merge_duplicates(lst): Takes a list as parameter and groups the lists with the same element in the third position (phone number). If it's the same phone number, even with different people, it means that they are members of the same family (father, mother, children, etc.) and that, as a result, all the receipts grouped will be sent to the same person, who is responsible for the payment of the group.

The script then opens WhatsApp Web using the os.startfile() function, loads an excel sheet using the openpyxl library and opens an image file for a logo and signature. The script resizes the images using the PIL library and generates the PDF document using the reportlab library. Reportlab is used to format and position the text on the pdf, create the pdf and set the layout of the pdf. Finally, it calls the enviar_msg() function to send the message to the correspondent person.

This script can be useful for automating repetitive tasks and sending bulk messages on WhatsApp. Note that the script uses the pyautogui library to automate mouse clicks and keyboard presses, so it must be run on a computer with a GUI. Additionally, you have to have the correct path of the files and sheet to make the script work.

NOTE: The script was created for use in Brazil and includes Portuguese language in some texts, which may need to be adapted for use in other countries.

