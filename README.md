# PdfLetterBulkMailer

PdfLetterBulkMailer is a program for bulk creation and sending of formatted PDF files in emails.

It is written in C# targetting Microsoft's .NET (v4.0) and can be built using Microsoft's Visual Studio (Community or Professional).

It calls a Latex binary (pdflatex) to create a PDF file from a .tex formatted markup file. The example latex file demonstrates a formatted letter that includes cross-referenced name and address. The letter also includes a logo in the header and a picture that is part of the letter's contents.

The cross referenced information such as name and address is included in the .tex file as placeholders which are substituted by actual values by the program. The program preprocess the .tex file for each record in a database of names, addresses and other information creating a final .tex file that can be compiled into a postscript and eventually a PDF file by the latex processor.

The example database is implemented as a simple comma separated values (.csv) file.

The email is sent using the SmtpClient class in Microsoft's .NET platform. The email sender's SMTP address and username are passed as arguments to the program while the sender's email password is input on the console when prompted by the program.

To successfully run the program, the pdflatex binary must be in the PATH. The program must be executed from the directory that contains the .csv database file, the .tex file, a text file containing the email message and other resources. These resources are contained in the Resources directory in the root directory of this project.
