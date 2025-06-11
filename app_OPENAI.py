import os
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import openai
import dotenv


dotenv.load_dotenv()

# Initialize the OpenAI client
client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY")) 

class ChatGPTSummarizer:
    def __init__(self):
        pass

    def summarize(self, text, min_words=40, max_words=50, max_retries=3):
        prompt = (
            f"|||NON-NEGOTIABLE INSTRUCTIONS|||\n"
            f"Create a {min_words}-{max_words} word summary that:\n"
            f"1. Is EXACTLY {max_words} words (ABSOLUTE LIMIT)\n"
            f"2. Uses COMPLETE SENTENCES only\n"
            f"3. Preserves ALL KEY INFORMATION from this 300+ word text:\n"
            f"----------------\n"
            f"{text}\n"
            f"----------------\n"
            f"FORMAT YOUR RESPONSE AS:\n"
            f"[Your summary here]"
        )
        messages = [
            {"role": "system", "content": "You are a ruthless summarization engine that never exceeds word limits."},
            {"role": "user", "content": prompt}
        ]
        
        for attempt in range(max_retries):
            try:
                response = client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=messages,
                    temperature=0.5,
                    max_tokens=150
                )
                summary = response.choices[0].message.content.strip()
                return summary
            except Exception as e:
                print(f"Error during summarization (attempt {attempt+1}): {e}")
                time.sleep(2)
        return "[SUMMARY FAILED]"

class SmartSummarizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart Summarizer")
        self.root.geometry("600x500")
        
        # Variables for file and settings
        self.file_path = tk.StringVar()
        self.column_name = tk.StringVar()
        self.min_words = tk.IntVar(value=40)
        self.max_words = tk.IntVar(value=50)
        
        self.summarizer = ChatGPTSummarizer()
        
        self.create_widgets()
        
    def create_widgets(self):
        # Title label
        tk.Label(self.root, text="Smart Summarizer", font=("Arial", 16, "bold")).pack(pady=10)
        
        # File upload section
        tk.Label(self.root, text="1. Upload Excel File", font=("Arial", 12, "bold")).pack(anchor="w", padx=20)
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=5)
        tk.Button(file_frame, text="Browse File", command=self.browse_file, bg="#4CAF50", fg="white").pack(side=tk.LEFT)
        tk.Label(file_frame, textvariable=self.file_path, wraplength=400).pack(side=tk.LEFT, padx=10)
        
        # Column selection
        tk.Label(self.root, text="2. Select Text Column", font=("Arial", 12, "bold")).pack(anchor="w", padx=20, pady=(10,0))
        self.column_dropdown = ttk.Combobox(self.root, textvariable=self.column_name, state="readonly", width=50)
        self.column_dropdown.pack(pady=5)
        
        # Summary length settings
        tk.Label(self.root, text="3. Set Summary Length (words)", font=("Arial", 12, "bold")).pack(anchor="w", padx=20, pady=(10,0))
        length_frame = tk.Frame(self.root)
        length_frame.pack()
        tk.Label(length_frame, text="Min Words:").grid(row=0, column=0, padx=5, pady=5)
        tk.Entry(length_frame, textvariable=self.min_words, width=5).grid(row=0, column=1, padx=5, pady=5)
        tk.Label(length_frame, text="Max Words:").grid(row=0, column=2, padx=5, pady=5)
        tk.Entry(length_frame, textvariable=self.max_words, width=5).grid(row=0, column=3, padx=5, pady=5)
        
        # Process button
        tk.Button(self.root, text="Generate Summaries", command=self.process_file,
                  bg="#2196F3", fg="white", font=("Arial", 12, "bold")).pack(pady=20, fill=tk.X, padx=50)
        
        # Status label
        self.status_label = tk.Label(self.root, text="", fg="green")
        self.status_label.pack()
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            self.file_path.set(file_path)
            self.load_columns(file_path)
    
    def load_columns(self, file_path):
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            self.column_dropdown['values'] = df.columns.tolist()
            if df.columns.tolist():
                self.column_name.set(df.columns.tolist()[0])
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
            self.status_label.config(text="Processing file...", fg="blue")
            self.root.update()
            
            # Read the Excel file
            df = pd.read_excel(self.file_path.get(), engine="openpyxl")
            col = self.column_name.get()
            
            summaries = []
            total_rows = len(df)
            for i, text in enumerate(df[col], start=1):
                if pd.isna(text) or len(str(text).strip().split()) < 10:
                    summaries.append("")
                else:
                    summary = self.summarizer.summarize(
                        str(text),
                        min_words=self.min_words.get(),
                        max_words=self.max_words.get()
                    )
                    summaries.append(summary)
                self.status_label.config(text=f"Processed {i}/{total_rows} rows")
                self.root.update()
            
            # Add the summaries to a new column and save to a new Excel file
            df["Catalogue Summary"] = summaries
            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                       filetypes=[("Excel Files", "*.xlsx")],
                                                       title="Save Summarized File As")
            if output_path:
                df.to_excel(output_path, index=False)
                self.status_label.config(text="Summarization complete!", fg="green")
                messagebox.showinfo("Success", "Smart summaries saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            self.status_label.config(text="Error occurred!", fg="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = SmartSummarizerApp(root)
    root.mainloop()
