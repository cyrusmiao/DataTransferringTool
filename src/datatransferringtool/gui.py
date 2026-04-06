import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
from .config import load_config
from .core import DataTransfer
from pathlib import Path

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Data Transferring Tool")
        self.geometry("500x300")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=0)

        self.label = ctk.CTkLabel(self.main_frame, text="Select YAML Configuration File:", font=("Arial", 14, "bold"))
        self.label.grid(row=0, column=0, columnspan=2, pady=10, sticky="w")

        self.file_path_var = ctk.StringVar()
        self.entry = ctk.CTkEntry(self.main_frame, textvariable=self.file_path_var, placeholder_text="Path to YAML file...")
        self.entry.grid(row=1, column=0, padx=(0, 10), pady=10, sticky="ew")

        self.browse_btn = ctk.CTkButton(self.main_frame, text="Browse", command=self.browse_file)
        self.browse_btn.grid(row=1, column=1, pady=10, sticky="e")

        self.run_btn = ctk.CTkButton(self.main_frame, text="Run Data Transfer", command=self.run_transfer, height=40)
        self.run_btn.grid(row=2, column=0, columnspan=2, pady=20, sticky="ew")

        self.status_label = ctk.CTkLabel(self.main_frame, text="", text_color="green")
        self.status_label.grid(row=3, column=0, columnspan=2, pady=10, sticky="w")

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select YAML config",
            filetypes=(("YAML files", "*.yaml"), ("YML files", "*.yml"), ("All files", "*.*"))
        )
        if filename:
            self.file_path_var.set(filename)

    def run_transfer(self):
        config_path = self.file_path_var.get()
        if not config_path or not Path(config_path).exists():
            messagebox.showerror("Error", "Please select a valid YAML configuration file.")
            return

        self.status_label.configure(text="Running transfer...", text_color="blue")
        self.run_btn.configure(state="disabled")
        
        # Run in thread to prevent GUI freeze
        thread = threading.Thread(target=self._execute_transfer, args=(config_path,))
        thread.start()

    def _execute_transfer(self, config_path):
        try:
            config = load_config(config_path)
            transfer = DataTransfer(config)
            transfer.run()
            self.status_label.configure(text=f"Success! Output saved to: {config.output_file}", text_color="green")
            success_message = f"Data transfer completed successfully!\nOutput file: {config.output_file}"
            if config.generate_transfer_report:
                success_message += "\nReport: transfer_report.xlsx"
            if config.generate_reference_report:
                success_message += "\nReference report: reference_report.md"
            messagebox.showinfo("Success", success_message)
        except Exception as e:
            self.status_label.configure(text=f"Error occurred.", text_color="red")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        finally:
            self.run_btn.configure(state="normal")

def run_gui():
    app = App()
    app.mainloop()
