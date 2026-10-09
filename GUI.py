
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from tempfile import TemporaryDirectory
import shutil
import threading
import queue

import main


class PDFExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF & Excel Automation")
        self.root.geometry("720x540")

        self.files = []
        self.output_folder = tk.StringVar()
        self.status = tk.StringVar(value="Ready")
        self.results = queue.Queue()

        self.create_ui()

    def create_ui(self):
        frame = ttk.Frame(self.root, padding=25)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="PDF & Excel Automation",
            font=("Segoe UI", 22, "bold")
        ).pack(pady=10)

        ttk.Label(
            frame,
            text="Extract PDF data and clean CSV files"
        ).pack(pady=(0, 20))

        ttk.Button(
            frame,
            text="Add PDF / CSV Files",
            command=self.add_files
        ).pack(fill="x", pady=5)

        self.file_list = tk.Listbox(frame, height=10)
        self.file_list.pack(
            fill="both", expand=True, pady=10
        )

        ttk.Button(
            frame,
            text="Remove Selected File",
            command=self.remove_file
        ).pack(pady=5)

        ttk.Button(
            frame,
            text="Choose Output Folder",
            command=self.choose_output
        ).pack(fill="x", pady=5)

        ttk.Label(
            frame,
            textvariable=self.output_folder,
            wraplength=600
        ).pack(pady=5)

        self.generate_button = ttk.Button(
            frame,
            text="Generate Excel Report",
            command=self.generate
        )
        self.generate_button.pack(
            fill="x", pady=15
        )

        ttk.Label(
            frame,
            textvariable=self.status
        ).pack(pady=5)

    def add_files(self):
        selected = filedialog.askopenfilenames(
            filetypes=[
                ("Supported Files", "*.pdf *.csv"),
                ("PDF Files", "*.pdf"),
                ("CSV Files", "*.csv")
            ]
        )

        for file_path in selected:
            if file_path not in self.files:
                self.files.append(file_path)
                self.file_list.insert(
                    tk.END, Path(file_path).name
                )

    def remove_file(self):
        selected = self.file_list.curselection()

        for index in reversed(selected):
            del self.files[index]
            self.file_list.delete(index)

    def choose_output(self):
        folder = filedialog.askdirectory()

        if folder:
            self.output_folder.set(folder)

    def generate(self):
        if not self.files:
            messagebox.showwarning(
                "Missing Files",
                "Please select PDF or CSV files."
            )
            return

        if not self.output_folder.get():
            messagebox.showwarning(
                "Missing Folder",
                "Please select an output folder."
            )
            return

        self.status.set("Processing files...")
        self.generate_button.config(state="disabled")

        threading.Thread(
            target=self.process_files,
            args=(list(self.files), self.output_folder.get()),
            daemon=True
        ).start()

        self.root.after(150, self.check_result)

    def process_files(self, files, output_path):
        try:
            output = Path(output_path)

            with TemporaryDirectory() as temp:
                input_folder = Path(temp)

                for file_path in files:
                    source = Path(file_path)
                    destination = input_folder / source.name
                    counter = 1

                    while destination.exists():
                        destination = input_folder / (
                            f"{source.stem}_{counter}{source.suffix}"
                        )
                        counter += 1

                    shutil.copy2(source, destination)

                old_input = main.INPUT_FOLDER
                old_output = main.OUTPUT_FOLDER

                try:
                    main.INPUT_FOLDER = input_folder
                    main.OUTPUT_FOLDER = output

                    main.generate_excel()
                finally:
                    main.INPUT_FOLDER = old_input
                    main.OUTPUT_FOLDER = old_output

            self.results.put(("success", str(output)))

        except Exception as error:
            self.results.put(("error", str(error)))

    def check_result(self):
        try:
            result, detail = self.results.get_nowait()
        except queue.Empty:
            self.root.after(150, self.check_result)
            return

        self.generate_button.config(state="normal")

        if result == "success":
            self.status.set("Processing finished")
            messagebox.showinfo(
                "Complete",
                f"Check your Excel report in:\n{detail}"
            )
        else:
            self.status.set("Processing failed")
            messagebox.showerror("Error", detail)


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExcelApp(root)
    root.mainloop()
