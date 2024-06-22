import os
import tempfile
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
import email
from xhtml2pdf import pisa
from PyPDF2 import PdfMerger
from outlookmsgfile import load
from datetime import datetime, timedelta
from threading import Thread
from concurrent.futures import ThreadPoolExecutor