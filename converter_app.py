import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import os
from openpyxl.utils import get_column_letter

# Create a class that inherits from both TkinterDnD.Tk and ttk.Window
# Note: ttkbootstrap.Window internally creates a tk.Tk, so we need to be careful with inheritance.
# The safest way to combine ttkbootstrap themes and DnD is to use TkinterDnD.Tk as root and apply style manually, 
# OR hijack the root of ttk.Window. Let's try wrapping.

class DnDWindow(TkinterDnD.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

class CSVConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CtE Converter v.1.0.0")
        self.root.geometry("450x450")
        self.root.resizable(False, False) # Fixed Width
        
        # Apply theme manually since we are using DnD root
        self.style = ttk.Style(theme="flatly")
        self.style.master = self.root
        
        # Variables
        self.files_path = [] 
        self.output_folder_path = tk.StringVar()
        self.use_source_folder = tk.BooleanVar(value=True)
        self.has_header = tk.BooleanVar(value=True)
        self.auto_width = tk.BooleanVar(value=True)
        self.auto_convert = tk.BooleanVar(value=False) # Auto convert checklist
        self.always_on_top = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Siap (Seret file ke sini)")

        self.create_widgets()
        self.setup_dnd()

    def create_widgets(self):
        # Top Controls (Always on top Checkbox)
        top_bar = ttk.Frame(self.root, padding=(10, 5))
        top_bar.pack(fill=X)
        ttk.Checkbutton(top_bar, text="Selalu di Atas", variable=self.always_on_top, command=self.toggle_topmost, bootstyle="secondary-round-toggle").pack(side=RIGHT)

        # Footer (Packed early to ensure visibility at bottom)
        footer_frame = ttk.Frame(self.root, padding=5)
        footer_frame.pack(side=BOTTOM, fill=X)
        footer_label = ttk.Label(footer_frame, text="CtE Converter v.1.0.0 oleh Lalu Habib Sasiwimbe", font=("Helvetica", 7), foreground="gray", anchor=CENTER)
        footer_label.pack()

        # Main Container
        main_frame = ttk.Frame(self.root, padding="15 5 15 10") # Reduced padding
        main_frame.pack(fill=BOTH, expand=True)

        # Title
        title_label = ttk.Label(main_frame, text="CSV to Excel Converter", font=("Helvetica", 16, "bold"), bootstyle="primary")
        title_label.pack(pady=(0, 10))

        # --- Drop Zone / File Selection ---
        self.drop_frame = ttk.Labelframe(main_frame, text="Pilih File", padding=10, bootstyle="info")
        self.drop_frame.pack(fill=X, pady=(0, 10))
        
        self.drop_label = ttk.Label(self.drop_frame, text="Seret File ke sini / Pilih Manual", font=("Helvetica", 9), justify=CENTER)
        self.drop_label.pack(pady=(0, 5))

        btn_container = ttk.Frame(self.drop_frame)
        btn_container.pack(fill=X)
        
        browse_btn = ttk.Button(btn_container, text="Pilih File", command=self.browse_files, bootstyle="info-outline", width=15)
        browse_btn.pack(side=TOP) # Compact button
        
        self.file_list_label = ttk.Label(self.drop_frame, text="0 file dipilih", font=("Helvetica", 8, "italic"))
        self.file_list_label.pack(pady=(5, 0))

        # --- Settings ---
        settings_frame = ttk.Frame(main_frame) # Removed Labelframe border to save space
        settings_frame.pack(fill=X, pady=(0, 10))

        # Output Path Row
        out_row = ttk.Frame(settings_frame)
        out_row.pack(fill=X, pady=(0, 5))
        ttk.Checkbutton(out_row, text="Simpan di Sumber", variable=self.use_source_folder, command=self.toggle_output_browse, bootstyle="success-round-toggle").pack(side=LEFT)
        self.browse_out_btn = ttk.Button(out_row, text="Folder...", command=self.browse_output_folder, state=DISABLED, bootstyle="secondary-outline", width=8)
        self.browse_out_btn.pack(side=RIGHT)
        
        # Options Row
        opts_row = ttk.Frame(settings_frame)
        opts_row.pack(fill=X)
        ttk.Checkbutton(opts_row, text="Header", variable=self.has_header, bootstyle="success-round-toggle").pack(side=LEFT, padx=(0,10))
        ttk.Checkbutton(opts_row, text="Auto-Width", variable=self.auto_width, bootstyle="success-round-toggle").pack(side=LEFT)
        ttk.Checkbutton(opts_row, text="Auto-Run", variable=self.auto_convert, bootstyle="warning-round-toggle").pack(side=RIGHT)

        # --- Action & Tools ---
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=X, pady=(5, 5))

        self.convert_btn = ttk.Button(action_frame, text="MULAI KONVERSI", command=self.convert_files, bootstyle="primary", width=20)
        self.convert_btn.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        
        # Quick Open Output Button
        self.open_folder_btn = ttk.Button(action_frame, text="📂", command=self.open_current_output_folder, bootstyle="secondary-outline", width=3, state=DISABLED)
        self.open_folder_btn.pack(side=RIGHT)

        # --- Status ---
        self.progress = ttk.Progressbar(main_frame, bootstyle="success-striped", mode='determinate')
        self.progress.pack(fill=X, pady=(5, 5))
        
        self.toast_frame = ttk.Frame(main_frame, style="secondary.TFrame", padding=5)
        self.toast_label = ttk.Label(self.toast_frame, text="", font=("Helvetica", 9), foreground="white", background="#95a5a6")
        self.toast_label.pack()

    def open_current_output_folder(self):
        target = self.output_folder_path.get()
        if self.use_source_folder.get():
             # Try to guess based on first file if available
             if self.files_path:
                 target = os.path.dirname(self.files_path[0])
             else:
                 self.show_toast("Belum ada file dipilih", is_error=True)
                 return
        
        if target and os.path.exists(target):
            os.startfile(target)
        else:
             self.show_toast("Folder tidak ditemukan", is_error=True)

    def setup_dnd(self):
        # Register the drop target
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)

    def toggle_topmost(self):
        self.root.attributes('-topmost', self.always_on_top.get())

    def toggle_output_browse(self):
        if self.use_source_folder.get():
            self.browse_out_btn.config(state=DISABLED)
        else:
            self.browse_out_btn.config(state=NORMAL)

    def browse_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder_path.set(folder)

    def browse_files(self):
        filenames = filedialog.askopenfilenames(
            title="Pilih File CSV",
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))
        )
        if filenames:
            self.handle_files_input(filenames)

    def on_drop(self, event):
        raw_files = self.root.tk.splitlist(event.data)
        self.handle_files_input(raw_files)

    def handle_files_input(self, filenames):
        valid_files = [f for f in filenames if f.lower().endswith('.csv')]
        
        if not valid_files:
            self.show_toast("File tidak valid!", is_error=True)
            return

        self.files_path = valid_files
        count = len(valid_files)
        self.file_list_label.config(text=f"{count} file CSV dipilih")
        self.status_var.set("Siap")
        
        # Unlock open folder btn since we have a source now
        self.open_folder_btn.config(state=NORMAL)
        
        if self.auto_convert.get():
            self.convert_files()

    def show_toast(self, message, is_error=False):
        if is_error:
            bg_color = "#e74c3c"
        else:
            bg_color = "#2ecc71"
        
        self.toast_label.config(text=message, background=bg_color)
        self.toast_frame.config(bootstyle="danger" if is_error else "success") 
        self.toast_frame.pack(pady=5, fill=X)
        self.root.update()
        self.root.after(3000, lambda: self.toast_frame.pack_forget())

    def convert_files(self):
        if not self.files_path:
            self.show_toast("Pilih file dulu!", is_error=True)
            return

        # Enable Open Folder button if not already
        self.open_folder_btn.config(state=NORMAL)

        target_dir = None
        if not self.use_source_folder.get():
            target_dir = self.output_folder_path.get()
            if not target_dir:
                self.show_toast("Folder output kosong!", is_error=True)
                return

        self.convert_btn.config(state=DISABLED)
        total_files = len(self.files_path)
        self.progress['maximum'] = total_files
        self.progress['value'] = 0
        
        success_count = 0
        errors = []

        for i, input_file in enumerate(self.files_path, 1):
            self.status_var.set(f"Processing...")
            self.root.update()
            
            try:
                header_arg = 0 if self.has_header.get() else None
                try:
                    df = pd.read_csv(input_file, header=header_arg)
                except Exception:
                     df = pd.read_csv(input_file, sep=None, engine='python', header=header_arg)

                directory, filename = os.path.split(input_file)
                name, _ = os.path.splitext(filename)
                
                final_dir = directory if self.use_source_folder.get() else target_dir
                output_file = os.path.join(final_dir, f"{name}.xlsx")

                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, header=self.has_header.get(), sheet_name='Sheet1')
                    if self.auto_width.get():
                        worksheet = writer.sheets['Sheet1']
                        for column in worksheet.columns:
                            max_length = 0
                            try:
                                column_letter = get_column_letter(column[0].column)
                            except:
                                continue # Skip if column issue
                            
                            for cell in column:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            worksheet.column_dimensions[column_letter].width = (max_length + 2)
                
                success_count += 1
                
            except Exception as e:
                errors.append(f"{os.path.basename(input_file)}: {str(e)}")
            
            self.progress['value'] = i
        
        self.convert_btn.config(state=NORMAL)
        self.status_var.set("Selesai")
        
        if errors:
            self.show_toast(f"Selesai dengan error ({len(errors)})", is_error=True)
            messagebox.showwarning("Log Error", "\n".join(errors))
        else:
            self.show_toast(f"Sukses! {success_count} file OK.")

if __name__ == "__main__":
    # Initialize DnD Root
    root = DnDWindow()
    app = CSVConverterApp(root)
    root.mainloop()
