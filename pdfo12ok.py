import os
import re
import fnmatch
import shutil
from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import json
from datetime import datetime


STATE_FILE_PATH = "app_state.json"

def save_state(app):
    """Zapisuje stan programu do pliku JSON."""
    state = {
        "source_directory": app.directory_entry.get(),
        "target_directory": app.target_directory_entry.get(),
        "search_string_1": app.search_string_entry_1.get("1.0", tk.END),
        "search_string_2": app.search_string_entry_2.get("1.0", tk.END),
        "operator": app.operator_var.get(),
        "recursive": app.recursive_search_var.get(),
        "case_sensitive": app.case_sensitive_var.get(),
        "outlook_folder": app.outlook_folder_entry.get(),
        "search_outlook": app.search_outlook_var.get(),
        "prefix": app.prefix_entry.get()
    }
    with open(STATE_FILE_PATH, "w") as file:
        json.dump(state, file)

def load_state(app):
    """Odczytuje stan programu z pliku JSON i uzupełnia odpowiednie pola."""
    if os.path.exists(STATE_FILE_PATH):
        with open(STATE_FILE_PATH, "r") as file:
            state = json.load(file)
            app.directory_entry.insert(0, state.get("source_directory", ""))
            app.target_directory_entry.insert(0, state.get("target_directory", ""))
            app.search_string_entry_1.insert(tk.END, state.get("search_string_1", ""))
            app.search_string_entry_2.insert(tk.END, state.get("search_string_2", ""))
            app.operator_var.set(state.get("operator", "AND"))
            app.recursive_search_var.set(state.get("recursive", False))
            app.case_sensitive_var.set(state.get("case_sensitive", False))
            app.outlook_folder_entry.insert(0, state.get("outlook_folder", "Skrzynka odbiorcza"))
            app.search_outlook_var.set(state.get("search_outlook", False))
            app.prefix_entry.insert(0, state.get("prefix", ""))

def search_pdf_for_string(directory, search_patterns, case_sensitive, operator, exclude_patterns=None, recursive=False):
    matching_files = []

    if recursive:
        for dirpath, dirnames, filenames in os.walk(directory):
            for filename in filenames:
                if exclude_patterns and any(fnmatch.fnmatch(filename, pattern) for pattern in exclude_patterns):
                    continue
                if filename.endswith(".pdf"):
                    full_path = os.path.join(dirpath, filename)
                    with open(full_path, 'rb') as file:
                        reader = PdfReader(file)
                        total_pages = len(reader.pages)
                        for page_num in range(total_pages):
                            text = reader.pages[page_num].extract_text()
                            if not case_sensitive:
                                text = text.lower()
                            matches = []
                            for pattern in search_patterns:
                                if not case_sensitive:
                                    pattern = pattern.lower()
                                if re.search(pattern, text):
                                    matches.append(True)
                                else:
                                    matches.append(False)
                            if operator == "AND" and all(matches):
                                matching_files.append(full_path)
                                break
                            elif operator == "OR" and any(matches):
                                matching_files.append(full_path)
                                break
                            elif operator == "NOT" and not any(matches):
                                matching_files.append(full_path)
                                break
    else:
        for filename in os.listdir(directory):
            if exclude_patterns and any(fnmatch.fnmatch(filename, pattern) for pattern in exclude_patterns):
                continue
            if filename.endswith(".pdf"):
                full_path = os.path.join(directory, filename)
                with open(full_path, 'rb') as file:
                    reader = PdfReader(file)
                    total_pages = len(reader.pages)
                    for page_num in range(total_pages):
                        text = reader.pages[page_num].extract_text()
                        if not case_sensitive:
                            text = text.lower()
                        matches = []
                        for pattern in search_patterns:
                            if not case_sensitive:
                                pattern = pattern.lower()
                            if re.search(pattern, text):
                                matches.append(True)
                            else:
                                matches.append(False)
                        if operator == "AND" and all(matches):
                            matching_files.append(full_path)
                            break
                        elif operator == "OR" and any(matches):
                            matching_files.append(full_path)
                            break
                        elif operator == "NOT" and not any(matches):
                            matching_files.append(full_path)
                            break
    return matching_files

def search_local_outlook(search_string, folder_name="inbox"):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Recursive function to search through all folders
    def find_folder(target_name, current_folder):
        for folder in current_folder.Folders:
            if folder.Name == target_name:
                return folder
            found_subfolder = find_folder(target_name, folder)
            if found_subfolder:
                return found_subfolder
        return None

    # Start from the root
    root_folder = outlook.Folders.Item(1)
    desired_folder = find_folder(folder_name, root_folder)

    if not desired_folder:
        print(f"Folder '{folder_name}' not found.")
        return []

    messages = desired_folder.Items
    matching_messages = messages.Restrict(f"[Subject] = '{search_string}' OR [Body] = '{search_string}'")
    results = []
    for message in matching_messages:
        results.append({
            'Subject': message.Subject,
            'Body': message.Body,
            'ReceivedTime': message.ReceivedTime
        })

    return results


# ... [Pozostały kod, który pozostaje bez zmian] ...

def copy_matching_files_to_new_folder(matching_files, search_string, target_directory, prefix=""):
    new_folder_path = os.path.join(target_directory, search_string)
    if not os.path.exists(new_folder_path):
        os.makedirs(new_folder_path)
    for idx, file_path in enumerate(matching_files, start=1):
        with open(file_path, 'rb') as file:
            reader = PdfReader(file)
            creation_date_str = reader.metadata.get("/CreationDate", '')[2:10]  # Format: D:YYYYMMDDHHMMSS
        
        creation_date_formatted = f"{creation_date_str[6:8]}{creation_date_str[4:6]}{creation_date_str[2:4]}"
        
        # Usuwanie niedozwolonych znaków z nazwy pliku
        safe_search_string = re.sub(r'[\/:*?"<>|]', '_', search_string)
        
        new_filename = f"{idx}_{prefix}_{safe_search_string}_{creation_date_formatted}.pdf"
        shutil.copy(file_path, os.path.join(new_folder_path, new_filename))






# ... [Pozostały kod, który pozostaje bez zmian] ...



class PDFSearchAppFull(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Przeszukiwarka PDF i Outlook")
        self.geometry("900x600")

        # Panel główny
        main_frame = tk.Frame(self)
        main_frame.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)

        # Ścieżka do folderu źródłowego
        dir_frame = tk.Frame(main_frame)
        dir_frame.grid(row=0, column=0, sticky="w")

        tk.Label(dir_frame, text="Ścieżka do folderu źródłowego:").grid(row=0, column=0, padx=5, pady=5)
        self.directory_entry = tk.Entry(dir_frame, width=40)
        self.directory_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(dir_frame, text="Wybierz", command=self.select_directory).grid(row=0, column=2, padx=5, pady=5)

        # Ścieżka do folderu docelowego
        target_dir_frame = tk.Frame(main_frame)
        target_dir_frame.grid(row=1, column=0, sticky="w")

        tk.Label(target_dir_frame, text="Ścieżka do folderu docelowego:").grid(row=0, column=0, padx=5, pady=5)
        self.target_directory_entry = tk.Entry(target_dir_frame, width=40)
        self.target_directory_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(target_dir_frame, text="Wybierz", command=self.select_target_directory).grid(row=0, column=2, padx=5, pady=5)

        # Szukane frazy
        search_frame = tk.Frame(main_frame)
        search_frame.grid(row=2, column=0, sticky="w")

        tk.Label(search_frame, text="Szukana fraza 1:").grid(row=0, column=0, padx=5, pady=5)
        self.search_string_entry_1 = tk.Text(search_frame, height=5, width=40)
        self.search_string_entry_1.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(search_frame, text="Szukana fraza 2:").grid(row=1, column=0, padx=5, pady=5)
        self.search_string_entry_2 = tk.Text(search_frame, height=5, width=40)
        self.search_string_entry_2.grid(row=1, column=1, padx=5, pady=5)

        # Operator
        operator_frame = tk.Frame(main_frame)
        operator_frame.grid(row=3, column=0, sticky="w")

        tk.Label(operator_frame, text="Operator:").grid(row=0, column=0, padx=5, pady=5)
        self.operator_var = tk.StringVar(value="AND")
        self.operator_options = ["AND", "OR", "NOT"]
        self.operator_dropdown = ttk.Combobox(operator_frame, textvariable=self.operator_var, values=self.operator_options, width=37)
        self.operator_dropdown.grid(row=0, column=1, padx=5, pady=5)

        # Prefix
        prefix_frame = tk.Frame(main_frame)
        prefix_frame.grid(row=4, column=0, sticky="w")

        tk.Label(prefix_frame, text="Prefix dla skopiowanych plików:").grid(row=0, column=0, padx=5, pady=5)
        self.prefix_entry = tk.Entry(prefix_frame, width=40)
        self.prefix_entry.grid(row=0, column=1, padx=5, pady=5)

        # Dodatkowe opcje
        options_frame = tk.Frame(main_frame)
        options_frame.grid(row=5, column=0, sticky="w")

        self.recursive_search_var = tk.BooleanVar(value=False)
        tk.Checkbutton(options_frame, text="Przeszukaj foldery rekurencyjnie", variable=self.recursive_search_var).grid(row=0, column=0, sticky="w", padx=5, pady=5)

        self.case_sensitive_var = tk.BooleanVar(value=False)
        tk.Checkbutton(options_frame, text="Uwzględnij wielkość liter", variable=self.case_sensitive_var).grid(row=1, column=0, sticky="w", padx=5, pady=5)

        # Nazwa folderu w Outlook
        outlook_frame = tk.Frame(main_frame)
        outlook_frame.grid(row=6, column=0, sticky="w")

        tk.Label(outlook_frame, text="Nazwa folderu w Outlook:").grid(row=0, column=0, padx=5, pady=5)
        self.outlook_folder_entry = tk.Entry(outlook_frame, width=40)
        self.outlook_folder_entry.grid(row=0, column=1, padx=5, pady=5)

        self.search_outlook_var = tk.BooleanVar(value=False)
        tk.Checkbutton(outlook_frame, text="Przeszukaj Outlook", variable=self.search_outlook_var).grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=5)

        # Przycisk do wyszukiwania
        self.search_button = tk.Button(main_frame, text="Szukaj", command=self.on_search_click)
        self.search_button.grid(row=7, column=0, pady=20)

    def select_directory(self):
        folder_selected = filedialog.askdirectory()
        self.directory_entry.delete(0, tk.END)
        self.directory_entry.insert(0, folder_selected)

    def select_target_directory(self):
        folder_selected = filedialog.askdirectory()
        self.target_directory_entry.delete(0, tk.END)
        self.target_directory_entry.insert(0, folder_selected)

    def on_search_click(self):
        directory = self.directory_entry.get()
        target_directory = self.target_directory_entry.get()
        search_strings_raw_1 = self.search_string_entry_1.get("1.0", tk.END).splitlines()
        search_strings_1 = [s.strip() for s in search_strings_raw_1 if s.strip()]
        search_strings_raw_2 = self.search_string_entry_2.get("1.0", tk.END).splitlines()
        search_strings_2 = [s.strip() for s in search_strings_raw_2 if s.strip()]
        case_sensitive = self.case_sensitive_var.get()
        operator = self.operator_var.get()

        outlook_folder_name = self.outlook_folder_entry.get()
        if not outlook_folder_name:
           outlook_folder_name = "Skrzynka odbiorcza"
        recursive = self.recursive_search_var.get()
        prefix = self.prefix_entry.get() if self.prefix_entry.get() else None
        for search_string_1 in search_strings_1:
            for search_string_2 in search_strings_2:
                if self.operator_var.get() == "AND":
                    search_strings_combined = [search_string_1, search_string_2]
                else:  # OR or NOT
                    search_strings_combined = [search_string_1, search_string_2]
                pdf_results = search_pdf_for_string(directory, search_strings_combined, case_sensitive, operator, recursive=recursive)
                folder_name = f"{prefix}_{search_string_1}" if prefix else search_string_1
                copy_matching_files_to_new_folder(pdf_results, folder_name, target_directory)
                if self.search_outlook_var.get():
                    search_local_outlook(f"{search_string_1} {self.operator_var.get()} {search_string_2}", outlook_folder_name)

app = PDFSearchAppFull()
load_state(app)  # Wczytanie stanu
app.protocol("WM_DELETE_WINDOW", lambda: (save_state(app), app.destroy()))  # Zapis stanu przed zamknięciem programu
app.mainloop()
