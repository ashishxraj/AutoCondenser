# AutoCondenser

AI-Powered Text Summarization for Catalogs, Reports, and Long-Form Content

## 🚀Features 
  - Strict Word Limits: Convert 300+ word texts → exact 40-50 word summaries.
  - Batch Processing: Handle 8,000+ Excel rows in one go
  - Cost Tracking: ~$5 per 8,000 summaries (GPT-3.5 turbo) or switch to BART Model for completely free solution.
  - Multilingual Support: Preserves Hindi/Devanagari and other non-English text
  - No Code Needed: Simple GUI for non-technical users

## 🛠️ Installation
  - Prerequisites
    - Python 3.8+
    - OpenAI API key ([Get yours here](https://openai.com/api/))

  - Steps
    - 1. clone the repo:
      ```bash
      git clone https://github.com/yourusername/AutoCondenser.git   
    - 2 Navigate to the project directory
         ```bash
         cd AutoCondenser
    - 3 Install dependencies:
         ``bash
         pip install -r requirements.txt
    - 4 Add your OpenAI API key:
         - Open main.py and replace:
           ```bash
           client = OpenAI(api_key="your-api-key-here")  # ← Paste your key here


