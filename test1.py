import os
import tempfile
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
import email
from xhtml2pdf import pisa
from PyPDF2 import PdfMerger
# from outlookmsgfile import load
from datetime import datetime, timedelta
from threading import Thread
from concurrent.futures import ThreadPoolExecutor




def convert_msg_to_eml(msg_file, temp_dir):
    msg = load(msg_file)
    eml_file = os.path.join(temp_dir, os.path.basename(msg_file).replace('.msg', '.eml'))
    with open(eml_file, 'w') as f:
        f.write(msg.as_string())
    return eml_file

def convert_eml_to_pdf(eml_file, temp_dir):
    with open(eml_file, 'r') as f:
        msg = email.message_from_file(f)
    
    html_file = os.path.join(temp_dir, os.path.basename(eml_file).replace('.eml', '.html'))
    pdf_file = os.path.join(temp_dir, os.path.basename(eml_file).replace('.eml', '.pdf'))

    email_metadata = f"""
    <p><strong>From:</strong> {msg['From']}</p>
    <p><strong>Sent:</strong> {msg['Date']}</p>
    <p><strong>To:</strong> {msg['To']}</p>
    <p><strong>Subject:</strong> {msg['Subject']}</p>
    <hr>
    """

    with open(html_file, 'w') as f:
        f.write(email_metadata)
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                f.write(part.get_payload(decode=True).decode())

    with open(html_file, 'r') as f:
        html_content = f.read()

    with open(pdf_file, 'wb') as f:
        pisa.CreatePDF(html_content, dest=f)
    
    return pdf_file

def combine_pdfs(pdf_files, output_file):
    merger = PdfMerger()
    for pdf in pdf_files:
        merger.append(pdf)
    merger.write(output_file)
    merger.close()

def process_file(file, temp_dir):
    try:
        eml_file = convert_msg_to_eml(file, temp_dir)
        pdf_file = convert_eml_to_pdf(eml_file, temp_dir)
        return pdf_file
    except Exception as e:
        print(f"Error converting {file}: {e}")
        return None

def process_files(files):
    global start_time
    start_time = datetime.now()
    temp_dir = tempfile.mkdtemp()
    pdf_files = []

    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_file = {executor.submit(process_file, file, temp_dir): file for file in files}
        for future in future_to_file:
            result = future.result()
            if result:
                pdf_files.append(result)
                update_progress()

    if pdf_files:
        output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_file:
            combine_pdfs(pdf_files, output_file)
            status_message.set(f"Combined PDF saved as '{output_file}'")
    else:
        status_message.set("No PDF files were created.")

    # Clean up temporary files
    for file in os.listdir(temp_dir):
        os.remove(os.path.join(temp_dir, file))
    os.rmdir(temp_dir)

    end_time = datetime.now()
    elapsed_time = end_time - start_time
    timer_message.set(f"Time taken: {str(elapsed_time)}")

    reset_ui()

def update_progress():
    global processed_files
    processed_files += 1
    progress_var.set(f"{processed_files}/{total_files} files processed")
    progress_bar['value'] = (processed_files / total_files) * 100

def on_drop(event):
    files = root.tk.splitlist(event.data)
    add_files(files)

def add_files(files):
    for file in files:
        if file.endswith('.msg'):
            file_listbox.insert(tk.END, file)

def remove_selected_file():
    selected_items = file_listbox.curselection()
    for index in selected_items[::-1]:
        file_listbox.delete(index)

def reset_ui():
    global processed_files, total_files
    file_listbox.delete(0, tk.END)
    processed_files = 0
    total_files = 0
    progress_var.set("0/0 files processed")
    progress_bar['value'] = 0

def start_conversion():
    global total_files, processed_files
    files = file_listbox.get(0, tk.END)
    total_files = len(files)
    processed_files = 0
    if files:
        progress_var.set(f"0/{total_files} files processed")
        progress_bar['value'] = 0
        thread = Thread(target=lambda: process_files(files))
        thread.start()
    else:
        status_message.set("Please add some files to process.")


root = TkinterDnD.Tk()
root.title("DV Email to PDF")
root.geometry("600x400")



frame = ttk.Frame(root)
frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

label = ttk.Label(frame, text="Drag and drop MSG files here or click to select")
label.pack(padx=10, pady=10)

button_frame = ttk.Frame(frame)
button_frame.pack(padx=10, pady=10, fill=tk.X)

select_button = ttk.Button(button_frame, text="Select Files", command=lambda: add_files(filedialog.askopenfilenames(filetypes=[("MSG files", "*.msg")])))
select_button.pack(side=tk.LEFT, padx=5)

remove_button = ttk.Button(button_frame, text="Remove Selected", command=remove_selected_file)
remove_button.pack(side=tk.LEFT, padx=5)

reset_button = ttk.Button(button_frame, text="Reset", command=reset_ui)
reset_button.pack(side=tk.LEFT, padx=5)

file_listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE)
file_listbox.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

get_pdf_button = ttk.Button(frame, text="Get PDF", command=start_conversion)
get_pdf_button.pack(padx=10, pady=10)

progress_var = tk.StringVar(value="0/0 files processed")
progress_label = ttk.Label(frame, textvariable=progress_var)
progress_label.pack(padx=10, pady=5)

progress_bar = ttk.Progressbar(frame, mode='determinate')
progress_bar.pack(padx=10, pady=10, fill=tk.X)

status_message = tk.StringVar()
status_label = ttk.Label(frame, textvariable=status_message)
status_label.pack(padx=10, pady=5)

timer_message = tk.StringVar()
timer_label = ttk.Label(frame, textvariable=timer_message)
timer_label.pack(padx=10, pady=5)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)

processed_files = 0
total_files = 0

root.mainloop()
