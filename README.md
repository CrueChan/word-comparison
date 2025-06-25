# Word Document Comparison Tool

A Python tool for comparing content differences between multiple Word documents.

## Features

- Compare multiple Word documents (.docx format)
- Calculate similarity matrix between documents
- Identify content differences
- Provide detailed comparison reports

## Requirements

- Python 3.12+
- Dependencies: python-docx

## Installation

### Using uv (Recommended)
```bash
# Clone the repository
git clone https://github.com/CrueChan/word-comparison.git
cd word-comparison

# Install dependencies
uv sync
```

### Using pip
```bash
pip install python-docx
```

## Usage

1. Place Word documents to compare in the project directory
2. Modify file paths in `main.py`:
```python
files = [
    "document1.docx",
    "document2.docx"
]
```
3. Run the program:
```bash
python main.py
```

## Example Output

```
Comparison Results:
Documents have differences:
Similarity between document 1 and document 2: 95.67%

Similarity Matrix:
Document 1: ['100.00%', '95.67%']
Document 2: ['95.67%', '100.00%']
```

## Project Structure

```
word-comparison/
├── main.py              # Main program file
├── pyproject.toml       # Project configuration
├── uv.lock             # Dependency lock file
├── .gitignore          # Git ignore file
├── .python-version     # Python version
└── README.md           # Project documentation
```

## Contributing

Issues and Pull Requests are welcome to improve this project.

## License

[MIT License](LICENSE)