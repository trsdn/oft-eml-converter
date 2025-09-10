# OFT to EML Converter

[![GitHub release (latest by date)](https://img.shields.io/github/v/release/trsdn/oft-eml-converter)](https://github.com/trsdn/oft-eml-converter/releases)
[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)](https://www.python.org/downloads/)
[![Test Suite](https://github.com/trsdn/oft-eml-converter/workflows/Test%20Suite/badge.svg)](https://github.com/trsdn/oft-eml-converter/actions)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)](#system-requirements)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

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

- Python 3.8 or higher (3.8-3.12 officially supported)
- tkinter (usually included with Python)

### Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/trsdn/oft-eml-converter.git
   cd oft-eml-converter
   ```

2. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
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

- Python 3.8+ (officially tested on 3.8-3.12)
- 50MB free disk space
- Internet connection for initial setup
- Supported OS: Windows 10+, macOS 10.15+, Ubuntu 20.04+

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

1. Check that all dependencies are installed: `pip list`
2. Verify your Python version: `python --version` (should be 3.8+)
3. Run with verbose output for debugging
4. Check our [CI/CD status](https://github.com/trsdn/oft-eml-converter/actions) to ensure the latest build is working
5. Review [existing issues](https://github.com/trsdn/oft-eml-converter/issues) for similar problems
6. [Open a new issue](https://github.com/trsdn/oft-eml-converter/issues/new) if your problem persists

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Setup

1. Fork the repository
2. Clone your fork: `git clone https://github.com/yourusername/oft-eml-converter.git`
3. Create a feature branch: `git checkout -b feature-name`
4. Set up development environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   pip install -r requirements.txt
   ```
5. Make your changes
6. Run tests: `pytest tests/ -v` or `python -m tests.test_converter`
7. Ensure CI passes locally
8. Submit a pull request with a clear description

### Code Style

- Follow PEP 8 guidelines
- Use meaningful variable names
- Add docstrings for functions and classes
- Include tests for new functionality

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Built with Python's standard email and tkinter libraries
- Uses the excellent `extract-msg` library for OFT file parsing
- Inspired by the need for reliable OFT to EML conversion

## Changelog

### v1.1.0 *(Latest)*
- Enhanced CI/CD with comprehensive cross-platform testing
- Advanced CLI converter error handling validation
- Virtual environment creation and management testing
- Headless-safe GUI import testing for CI environments
- Cross-platform file operations validation
- Package requirement verification system
- Professional release management with GitHub Actions

### v1.0.0
- Initial release
- GUI and command-line interfaces
- Embedded image support
- Batch processing
- Cross-platform compatibility