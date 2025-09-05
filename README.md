# OFT to EML Converter

A Python application for converting Microsoft Outlook Template (.oft) files to standard EML format with embedded image support.

## Features

- **Format Support**: Converts .oft (Outlook Template) files to .eml (standard email format)
- **Embedded Images**: Properly handles inline images with Content-ID references
- **Batch Processing**: Convert multiple files at once
- **GUI Interface**: Easy-to-use graphical interface
- **Command Line**: Also available as a command-line tool
- **Cross-Platform**: Works on Windows, macOS, and Linux

## Installation

### Prerequisites

- Python 3.7 or higher
- tkinter (usually included with Python)

### Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/oft-eml-converter.git
   cd oft-eml-converter
   ```

2. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install extract-msg
   ```

## Usage

### GUI Application (Recommended)

Launch the graphical interface:
```bash
./run_gui.sh
```

Or directly:
```bash
python oft_to_eml_gui.py
```

**Features:**
- Click to browse and select OFT files
- Choose output directory (remembers last location)
- Real-time conversion progress
- Batch processing support
- Error handling and results display

### Command Line

Convert a single file:
```bash
python oft_to_eml_converter.py input.oft [output.eml]
```

**Examples:**
```bash
# Convert with automatic output naming
python oft_to_eml_converter.py template.oft

# Convert with specific output name
python oft_to_eml_converter.py template.oft converted.eml
```

## How It Works

The converter:

1. **Parses OFT files** using the `extract-msg` library
2. **Extracts email components**: headers, plain text, HTML body, and attachments
3. **Handles embedded images**: Converts attachments with Content-IDs to inline images
4. **Creates EML files**: Uses Python's `email` library to generate RFC-compliant MIME messages
5. **Preserves formatting**: Maintains original styling and embedded content

## Technical Details

### Supported Content

- **Headers**: From, To, CC, Subject, Date
- **Body**: Plain text and HTML content
- **Attachments**: Regular file attachments
- **Embedded Images**: Inline images with proper Content-ID mapping

### File Structure

```
oft-eml-converter/
├── oft_to_eml_converter.py    # Core conversion logic
├── oft_to_eml_gui.py          # GUI application
├── run_gui.sh                 # GUI launcher script
├── requirements.txt           # Python dependencies
└── README.md                  # This file
```

## Requirements

### Python Packages

- `extract-msg`: For parsing OFT/MSG files
- `tkinter`: For GUI interface (usually included)

### System Requirements

- Python 3.7+
- 50MB free disk space
- Internet connection for initial setup

## Troubleshooting

### Common Issues

**"No module named '_tkinter'"**
- Install tkinter: `brew install python-tk` (macOS) or `sudo apt-get install python3-tk` (Ubuntu)

**"Unable to load tkdnd library"**
- This is expected - the application gracefully falls back to browse-only mode

**Conversion errors**
- Ensure the .oft file isn't corrupted
- Check that you have read permissions for the input file
- Verify the output directory is writable

### Getting Help

1. Check that all dependencies are installed
2. Verify your Python version: `python --version`
3. Run with verbose output for debugging
4. Open an issue if problems persist

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Setup

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test with sample OFT files
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Built with Python's standard email and tkinter libraries
- Uses the excellent `extract-msg` library for OFT file parsing
- Inspired by the need for reliable OFT to EML conversion

## Changelog

### v1.0.0
- Initial release
- GUI and command-line interfaces
- Embedded image support
- Batch processing
- Cross-platform compatibility