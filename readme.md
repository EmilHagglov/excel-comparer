# Emil's Excel Comparer

This application allows you to compare two Excel files and identify differences between them.

## Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package installer)

### Windows

1. Install Python from [python.org](https://www.python.org/downloads/windows/)
2. Open Command Prompt and navigate to the project directory
3. Create a virtual environment:
   ```
   python -m venv venv
   ```
4. Activate the virtual environment:
   ```
   venv\Scripts\activate
   ```
5. Install required packages:
   ```
   pip install -r requirements.txt
   ```

### macOS

1. Install Python using Homebrew:
   ```
   brew install python
   ```
2. Open Terminal and navigate to the project directory
3. Create a virtual environment:
   ```
   python3 -m venv venv
   ```
4. Activate the virtual environment:
   ```
   source venv/bin/activate
   ```
5. Install required packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Activate the virtual environment (if not already activated)
2. Run the application:
   ```
   python excel_compare.py
   ```
3. Use the GUI to select two Excel files and compare them

## Development

- The main application code is in `excel_compare.py`
- Requirements are listed in `requirements.txt`