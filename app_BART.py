import pandas as pd
import torch
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tqdm import tqdm
from transformers import pipeline, BartTokenizer, BartForConditionalGeneration
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
from keybert import KeyBERT
import numpy as np
import warnings
warnings.filterwarnings('ignore')

class SmartBookSummarizer:
    def __init__(self):
        # INITIALIZE ALL MODELS
        self.device = 0 if torch.cuda.is_available() else -1
        self.sentence_model = SentenceTransformer('all-MiniLM-L6-v2')
        self.keyword_model = KeyBERT()
        self.tokenizer = BartTokenizer.from_pretrained('facebook/bart-large-cnn')
        self.model = BartForConditionalGeneration.from_pretrained('facebook/bart-large-cnn').to(
            torch.device("cuda" if torch.cuda.is_available() else "cpu")
        )

    def extract_key_concepts(self, text):
        #IDENTIFY CORE CONCEPTS USING KEYBERT
        return [kw[0] for kw in 
                self.keyword_model.extract_keywords(
                    text,
                    keyphrase_ngram_range=(1, 2),
                    stop_words='english',
                    top_n=5
                )]

    def select_relevant_sentences(self, text, concepts):
        """Semantic filtering of important sentences"""
        sentences = [s.strip() for s in text.split('. ') if len(s) > 10]
        if len(sentences) < 3:
            return text
            
        # ENCODE AND COMPARE
        sentence_embeddings = self.sentence_model.encode(sentences)
        concept_embeddings = self.sentence_model.encode(concepts)
        similarity_scores = cosine_similarity(sentence_embeddings, concept_embeddings)
        
        # SELECT TOP 3 MOST RELEVANT SENTENCES
        top_indices = np.argsort(np.max(similarity_scores, axis=1))[-3:][::-1]
        return '. '.join([sentences[i] for i in top_indices if similarity_scores[i].max() > 0.3])

    def generate_summary(self, text, max_length=200, min_length=100):
        """End-to-end smart summarization"""
        try:
            if pd.isna(text) or len(str(text).strip()) < 50:
                return ""
                
            text = str(text)
            concepts = self.extract_key_concepts(text)
            important_content = self.select_relevant_sentences(text, concepts) or text
            
            # GENERATE SUMMARY WITH PROPER PARAMETERES
            inputs = self.tokenizer(
                important_content,
                max_length=1024,
                truncation=True,
                return_tensors="pt"
            ).to(self.model.device)
            
            summary_ids = self.model.generate(
                inputs['input_ids'],
                num_beams=4,
                length_penalty=2.0,
                max_length=max_length,
                min_length=min_length,
                no_repeat_ngram_size=3,
                early_stopping=True
            )
            
            return self.tokenizer.decode(summary_ids[0], skip_special_tokens=True)
            
        except Exception as e:
            print(f"⚠️ Error processing one entry: {str(e)[:100]}...")
            return "[SUMMARY FAILED]"

class BookSummarizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart Book Preface Summarizer")
        self.root.geometry("550x450")
        
        # VARIABLES
        self.file_path = tk.StringVar()
        self.column_name = tk.StringVar()
        self.min_words = tk.IntVar(value=100)
        self.max_words = tk.IntVar(value=200)
        self.summarizer = None
        
        # GUI ELEMENTS
        self.create_widgets()
    
    def create_widgets(self):
        # Title
        tk.Label(self.root, text="Smart Book Summarizer", 
                font=('Arial', 14, 'bold')).pack(pady=10)
        
        # FILE UPLOAD SECTION
        tk.Label(self.root, text="1. Upload Excel File", 
               font=('Arial', 10, 'bold')).pack(anchor='w', padx=20)
        
        upload_frame = tk.Frame(self.root)
        upload_frame.pack(pady=5)
        
        tk.Button(
            upload_frame,
            text="Browse File",
            command=self.browse_file,
            bg="#4CAF50",
            fg="white"
        ).pack(side=tk.LEFT)
        
        tk.Label(upload_frame, textvariable=self.file_path, 
                wraplength=400).pack(side=tk.LEFT, padx=10)
        
        # COLUMN SELECTION
        tk.Label(self.root, text="2. Select Text Column", 
               font=('Arial', 10, 'bold')).pack(anchor='w', padx=20, pady=(10,0))
        
        self.column_dropdown = ttk.Combobox(self.root, textvariable=self.column_name, 
                                          state="readonly", width=50)
        self.column_dropdown.pack(pady=5)
        
        # SUMMARY LENGTH SETTINGS
        tk.Label(self.root, text="3. Set Summary Length", 
               font=('Arial', 10, 'bold')).pack(anchor='w', padx=20, pady=(10,0))
        
        length_frame = tk.Frame(self.root)
        length_frame.pack()
        
        tk.Label(length_frame, text="Min Words:").grid(row=0, column=0, sticky='e')
        tk.Entry(length_frame, textvariable=self.min_words, width=5).grid(row=0, column=1, padx=5)
        
        tk.Label(length_frame, text="Max Words:").grid(row=0, column=2, sticky='e')
        tk.Entry(length_frame, textvariable=self.max_words, width=5).grid(row=0, column=3, padx=5)
        
        # PROCESS BUTTON
        tk.Button(
            self.root,
            text="Generate Smart Summaries",
            command=self.process_file,
            bg="#2196F3",
            fg="white",
            font=('Arial', 10, 'bold'),
            pady=10
        ).pack(pady=20, fill=tk.X, padx=50)
        
        # STATUS LABEL
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
            # INITIALIZE SUMMARIZER
            self.status_label.config(text="Loading AI models... (please wait)", fg="blue")
            self.root.update()
            
            self.summarizer = SmartBookSummarizer()
            
            # READ EXCEL FILE
            df = pd.read_excel(self.file_path.get())
            
            # SUMMARIZE EACH ENTRY
            summaries = []
            for text in tqdm(df[self.column_name.get()], desc="Generating smart summaries"):
                if pd.isna(text) or len(str(text).split()) < 10:
                    summaries.append("")
                    continue
                
                summary = self.summarizer.generate_summary(
                    str(text),
                    max_length=self.max_words.get(),
                    min_length=self.min_words.get()
                )
                summaries.append(summary)
            
            # ADD SUMMARIES TO DATAFRAME
            df["Summary"] = summaries
            
            # SAVE TO NEW FILE
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title="Save Summarized File As"
            )
            
            if output_path:
                df.to_excel(output_path, index=False)
                self.status_label.config(text="Smart summarization complete!", fg="green")
                messagebox.showinfo("Success", "Book summaries saved successfully!")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            self.status_label.config(text="Error occurred!", fg="red")

# RUN THE APPLICATION
if __name__ == "__main__":
    root = tk.Tk()
    app = BookSummarizerApp(root)
    root.mainloop()