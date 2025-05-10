import pandas as pd
from transformers import pipeline
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tqdm import tqdm
import torch
import warnings
warnings.filterwarnings('ignore')

class BookSummarizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Book Preface Summarizer")
        self.root.geometry("500x400")
        
        # Variables
        self.file_path = tk.StringVar()
        self.column_name = tk.StringVar()
        self.min_words = tk.IntVar(value=100)
        self.max_words = tk.IntVar(value=200)
        
        # GUI Elements
        self.create_widgets()
    
    def create_widgets(self):
        # File Upload Section
        tk.Label(self.root, text="Upload Excel File", font=('Arial', 12, 'bold')).pack(pady=10)
        
        tk.Button(
            self.root,
            text="Browse",
            command=self.browse_file,
            bg="#4CAF50",
            fg="white"
        ).pack(pady=5)
        
        tk.Label(self.root, textvariable=self.file_path, wraplength=400).pack()
        
        # Column Selection
        tk.Label(self.root, text="Select Column with Prefaces", font=('Arial', 12, 'bold')).pack(pady=10)
        
        self.column_dropdown = ttk.Combobox(self.root, textvariable=self.column_name, state="readonly")
        self.column_dropdown.pack()
        
        # Summary Length Settings
        tk.Label(self.root, text="Summary Length (Words)", font=('Arial', 12, 'bold')).pack(pady=10)
        
        tk.Label(self.root, text="Min Words:").pack()
        tk.Entry(self.root, textvariable=self.min_words).pack()
        
        tk.Label(self.root, text="Max Words:").pack()
        tk.Entry(self.root, textvariable=self.max_words).pack()
        
        # Process Button
        tk.Button(
            self.root,
            text="Generate Summaries",
            command=self.process_file,
            bg="#2196F3",
            fg="white",
            font=('Arial', 10, 'bold')
        ).pack(pady=20)
        
        # Status Label
        self.status_label = tk.Label(self.root, text="", fg="green")
        self.status_label.pack()
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            self.file_path.set(file_path)
            self.load_columns(file_path)
    
    def load_columns(self, file_path):
        try:
            df = pd.read_excel(file_path)
            self.column_dropdown['values'] = df.columns.tolist()
            if len(df.columns) > 0:
                self.column_name.set(df.columns[0])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {e}")
    
    def process_file(self):
        if not self.file_path.get():
            messagebox.showerror("Error", "Please select an Excel file first!")
            return
        
        if not self.column_name.get():
            messagebox.showerror("Error", "Please select a column!")
            return
        
        try:
            # Initialize summarizer
            self.status_label.config(text="Loading model... (this may take a minute)", fg="blue")
            self.root.update()
            
            summarizer = pipeline(
                "summarization",
                model="facebook/bart-large-cnn",
                device=0 if torch.cuda.is_available() else -1,
                truncation=True
            )
            
            # Read Excel file
            df = pd.read_excel(self.file_path.get())
            
            # Summarize each entry
            summaries = []
            for text in tqdm(df[self.column_name.get()], desc="Summarizing"):
                if pd.isna(text) or len(str(text).split()) < 10:
                    summaries.append("")
                    continue
                
                summary = summarizer(
                    str(text),
                    max_length=self.max_words.get(),
                    min_length=self.min_words.get()
                )[0]['summary_text']
                summaries.append(summary)
            
            # Add summaries to DataFrame
            df["Summary"] = summaries
            
            # Save to new file
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")]
            )
            
            if output_path:
                df.to_excel(output_path, index=False)
                self.status_label.config(text="Summarization complete! File saved.", fg="green")
                messagebox.showinfo("Success", "Summarized Excel file saved successfully!")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            self.status_label.config(text="Error occurred!", fg="red")

# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = BookSummarizerApp(root)
    root.mainloop()