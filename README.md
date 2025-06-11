# AutoCondenser
![GUI Screenshot](https://github.com/user-attachments/assets/33f024ef-2b8e-4b8d-9b55-6f956c007dbc)
### Book Preface Summarizer Toolkit


A collection of AI-powered tools for generating high-quality summaries of book prefaces from Excel files, using different NLP approaches.

## Table of Contents
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Comparison of Approaches](#comparison-of-approaches)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)
- [License](#license)

## Features

### Core Functionality
- Processes Excel files containing book prefaces
- Generates concise summaries with configurable length
- Preserves key information while reducing text length
- Adds summaries as a new column in the output Excel file

### Available Implementations:
1. **Basic BART Summarizer** (`app.py`)
   - Uses Facebook's BART-large-cnn model
   - Simple interface with file selection
   - Fast processing for basic summarization needs

2. **Smart BART Summarizer** (`app_BART.py`)
   - Enhanced version with semantic analysis
   - Key concept extraction using KeyBERT
   - Sentence relevance scoring with Sentence-Transformers
   - More coherent and focused summaries

3. **OpenAI GPT Summarizer** (`app_OPENAI.py`)
   - Uses GPT-3.5-turbo via OpenAI API
   - Strict word limit enforcement
   - Retry mechanism for API failures
   - Highest quality summaries (requires API key)

## Installation

### Prerequisites
- Python 3.8+
- pip package manager

### Steps
1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/book-summarizer.git
   cd book-summarizer

2. Install dependencies:
   ```bash
   pip install -r requirements.txt

3. For OpenAI version, create a ```.env``` file:
   ```bash
   OPENAI_API_KEY=your_api_key_here

## Usage
### Running the application
    # Basic BART version
    python app.py

    # Enhanced BART version
    python app_BART.py

    # OpenAI version
    python app_OPENAI.py

## Comparison of Approaches

| Feature               | Basic BART | Smart BART | OpenAI GPT |
|-----------------------|------------|------------|------------|
| Summary Quality       | Good       | Very Good  | Excellent  |
| Processing Speed      | Fast       | Moderate   | Slow       |
| Offline Capability    | Yes        | Yes        | No         |
| Hardware Requirements | Low        | Moderate   | Low        |
| Customization         | Limited    | High       | Medium     |
| API Key Required      | No         | No         | Yes        |

### Workflow
1. Click "Browse" to select your Excel file

2. Select the column containing book prefaces

3. Set desired summary length (min/max words)

4. Click "Generate Summaries"

5. Choose output location for the processed file

## Configuration
### Common Settings (All Versions)
- Input file path (via file dialog)
- Text column selection
- Minimum/Maximum summary length

### Smart BART Specific
- Keyphrase n-gram range (in code)
- Sentence relevance threshold (in code)
- Beam search parameters (in code)

### OpenAI Specific
- API key in `.env` file
- Model selection (`gpt-3.5-turbo`)
- Temperature setting (in code)
- Max retry attempts (in code)

## Troubleshooting
### Common Issues
1. File Loading Errors
- Ensure file is not open in Excel
- Verify file format (.xlsx or .xls)

2. Model Loading Issues
- Check internet connection (for first run)
- Ensure sufficient disk space (~1.5GB for BART models)
- Verify CUDA availability for GPU acceleration

3. OpenAI API Errors
- Validate API key in .env
- Check account quota
- Retry during low-traffic periods

4. Memory Errors
- Reduce batch size (in code)
- Process smaller files
- Use CPU-only mode if GPU memory limited

## License
This project is licensed under the MIT License. 

## Acknowledgments
- Facebook AI for BART models
- OpenAI for GPT models
- HuggingFace for transformers library
- KeyBERT for keyword extraction

