import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
from analyzer import analyze_vba_code, generate_documentation
from vba_parser import parse_vba_code

class ExcelReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Reader")

        self.label = tk.Label(root, text="Select an Excel file to read VBA macros")
        self.label.pack(pady=10)

        self.button = tk.Button(root, text="Browse", command=self.browse_file)
        self.button.pack(pady=10)

        self.text_area = tk.Text(root, wrap='word', height=20, width=80)
        self.text_area.pack(pady=10)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsm *.xlsb *.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            try:
                vba_code = self.extract_vba_code(file_path)
                if vba_code:  # Check if vba_code is not None and contains data
                    analysis = analyze_vba_code(vba_code)  # Call the analysis function
                    documentation = generate_documentation(analysis)  # Generate documentation
                    self.display_documentation(documentation)
                else:
                    messagebox.showerror("Error", "No VBA macros found in the selected file or an error occurred.")
            except Exception as e:
                messagebox.showerror("Error", f"Error extracting VBA code: {e}")

    def extract_vba_code(self, file_path):
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(file_path, ReadOnly=True)
            vba_code = {}
            if wb.HasVBProject:
                for vbcomponent in wb.VBProject.VBComponents:
                    if vbcomponent.Type == 1:  # vbext_ct_StdModule (Standard Module)
                        module_name = vbcomponent.Name
                        code = vbcomponent.CodeModule.Lines(1, vbcomponent.CodeModule.CountOfLines)
                        vba_code[module_name] = code
            wb.Close(SaveChanges=False)
            excel.Quit()
            return vba_code if vba_code else None
        except Exception as e:
            excel.Quit()  # Ensure Excel application is closed in case of an error
            print(f"Error extracting VBA code: {e}")
            return None

    def display_documentation(self, documentation):
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, documentation)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelReaderApp(root)
    root.mainloop()
