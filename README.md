DV Email to PDF Converter
Description
This is a Python application that converts Outlook MSG files to PDF format. It provides a graphical user interface (GUI) for easy file selection and conversion. The application uses the Tkinter library for the GUI, the outlookmsgfile library for parsing MSG files, and the xhtml2pdf library for converting HTML to PDF.

Dependencies
Python 3.x
Tkinter
tkinterdnd2
PyPDF2
outlookmsgfile
xhtml2pdf
Installation
Clone the repository:
bash
Copy code
git clone https://github.com/abhishek98as/Msg-to-pdf-convertor-python/tree/master
Install the dependencies:
bash
Copy code
pip install tkinter tkinterdnd2 PyPDF2 outlookmsgfile xhtml2pdf
Run the application:
bash
Copy code
python app.py
Usage
Drag and drop MSG files into the application or click the "Select Files" button to manually select MSG files.
Click the "Get PDF" button to convert the selected MSG files to PDF format.
The progress bar will show the conversion progress, and the status messages will indicate the success or failure of the conversion.
Once the conversion is complete, the combined PDF file will be saved.
Creating an Executable
Install pyinstaller:
bash
Copy code
pip install pyinstaller
Navigate to the project directory and run:
bash
Copy code
pyinstaller --onefile --noconsole --collect-all reportlab --collect-all tkinterdnd2 final.py
The executable file will be created in the dist directory.
Note
Due to the limitations of the reportlab library for creating graphics, an alternative method was used for handling graphics in the PDF files to avoid issues when creating the executable.
The application deletes temporary files after the conversion process is complete to clean up the system.
Credits
This application was created by Abhishek Singh as a part of Personal. It is open-source and available on GitHub.

License
This project is licensed under the GNU License - see the LICENSE.md file for details.
