# DOCX to Markdown Converter

A modern tool to convert Microsoft Word (DOCX) files to Markdown format using Pandoc, with a clean PyQt5 interface.

![DOCX to Markdown Converter](screenshot.png)

## Features

- Convert single or multiple DOCX files to Markdown
- Batch conversion of entire directories
- Modern PyQt5 GUI interface
- Command-line interface for automation
- Real-time progress tracking for batch conversions
- Customizable Pandoc arguments
- Cross-platform compatibility (Windows, macOS, Linux)

## Requirements

- Python 3.6+
- Pandoc (must be installed separately)
- PyQt5

## Installation

### Install Pandoc

First, install Pandoc following the instructions for your operating system:

- **macOS**: `brew install pandoc`
- **Windows**: Download from [pandoc.org/installing.html](https://pandoc.org/installing.html)
- **Linux**: Use your package manager, e.g., `sudo apt install pandoc`

### Install the Converter

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/docx-to-markdown-converter.git
   cd docx-to-markdown-converter
   ```

2. Create and activate a virtual environment (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install the required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Graphical User Interface (GUI)

To launch the GUI:

```bash
python docx_to_markdown.py --gui
```

Or simply:

```bash
python docx_to_markdown.py
```

### Command Line Interface

Convert specific DOCX files:

```bash
python docx_to_markdown.py file1.docx file2.docx
```

Convert all DOCX files in a directory:

```bash
python docx_to_markdown.py directory_path
```

Recursively search for DOCX files in directories:

```bash
python docx_to_markdown.py directory_path -r
```

Specify an output directory:

```bash
python docx_to_markdown.py file.docx -o output_directory
```

Pass additional arguments to Pandoc:

```bash
python docx_to_markdown.py file.docx --pandoc-args --toc --standalone
```

## Advanced Pandoc Options

The converter supports all Pandoc options. Some useful ones include:

- `--toc`: Include a table of contents
- `--standalone`: Produce a standalone document
- `--extract-media=DIR`: Extract media files to the specified directory
- `--wrap=none`: Disable text wrapping (useful for version control)
- `--reference-links`: Use reference-style links instead of inline links

## Development

### Project Structure

```
docx-to-markdown-converter/
├── docx_to_markdown.py  # Main application file
├── requirements.txt     # Python dependencies
├── LICENSE             # MIT License
├── README.md           # This file
└── screenshot.png      # Screenshot for README
```

### Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
