# OFT to EML Converter

[![License](https://img.shields.io/github/license/trsdn/oft-eml-converter)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.10%2B-blue)](#system-requirements)
[![Test Suite](https://github.com/trsdn/oft-eml-converter/actions/workflows/test.yml/badge.svg)](https://github.com/trsdn/oft-eml-converter/actions/workflows/test.yml)
[![Release](https://img.shields.io/github/v/release/trsdn/oft-eml-converter)](https://github.com/trsdn/oft-eml-converter/releases)
[![Conformance](.github/badges/conformance.svg)](docs/self-assessment.md)

A Python application for converting Microsoft Outlook Template (.oft) files to standard EML format with embedded image support.

Its language is English, and it ships no localized content. See `L01` and `L03`.

**This tool runs entirely on your machine.** It reads the files you point it at, writes the results next to them, and contacts no network service. See [Data handling](#data-handling).

## Features

- **Format Support**: Converts .oft (Outlook Template) files to .eml (standard email format)
- **Embedded Images**: Properly handles inline images with Content-ID references
- **Batch Processing**: Convert multiple files at once
- **GUI Interface**: Easy-to-use graphical interface
- **Command Line**: Also available as a command-line tool
- **Cross-Platform**: Works on Windows, macOS, and Linux

## Installation

### Prerequisites

- Python 3.10 or higher (3.10-3.12 officially supported)
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

For how the pieces fit together and why, see [Architecture](docs/ARCHITECTURE.md) and the [decision records](docs/decisions/).

## Requirements

### Python Packages

- `extract-msg`: For parsing OFT/MSG files
- `tkinter`: For GUI interface (usually included)

### System Requirements

- Python 3.10+ (officially tested on 3.10-3.12)
- 50MB free disk space
- Supported OS: Windows 10+, macOS 10.15+, Ubuntu 20.04+
- Internet access is needed **once**, to install dependencies. The converter itself never uses the network.

## Data handling

- **What is collected:** nothing. There is no telemetry, no analytics, and no crash reporting.
- **What is transmitted:** nothing. Conversion is entirely local, and the tool does not resolve remote references even when the HTML body it copies contains them.
- **What is stored:** the `.eml` files you asked for, in the output directory you chose. The GUI additionally writes `converter_config.json` in the directory it was started from, holding only your last output directory. Delete that file to reset it; nothing else persists.
- **Outbound network destinations:** PyPI, and only during `pip install -r requirements.txt`. At runtime there are none.

## Accessibility

- The command-line interface prints plain ASCII. It uses no colour, no Unicode box drawing, and no emoji, so it stays readable on a terminal without colour support and through a screen reader.
- **Known limitation:** the Tkinter GUI has not been audited with a screen reader, and its accessible names come from whatever Tkinter exposes by default. The command-line interface is the accessible path, and it can do everything the GUI can do.

## Versioning

This project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html). The public surface is the CLI invocation and the `convert_oft_to_eml` function; a breaking change to either raises the major version. Dropping a Python version from the supported range is a breaking change.

Releases are tagged `vMAJOR.MINOR.PATCH`. See the [changelog](CHANGELOG.md).

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
2. Verify your Python version: `python --version` (should be 3.10+)
3. Run with verbose output for debugging
4. Check our [CI/CD status](https://github.com/trsdn/oft-eml-converter/actions) to ensure the latest build is working
5. Review [existing issues](https://github.com/trsdn/oft-eml-converter/issues) for similar problems
6. [Open a new issue](https://github.com/trsdn/oft-eml-converter/issues/new) if your problem persists

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Setup

1. Fork the repository
2. Clone your fork: `git clone https://github.com/<your-account>/oft-eml-converter.git`
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

## Repository activity

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/trsdn/oft-eml-converter/stats/.github/stats/repo-card-dark.svg">
  <img alt="Repository statistics" src="https://raw.githubusercontent.com/trsdn/oft-eml-converter/stats/.github/stats/repo-card.svg">
</picture>

## License

MIT — see [LICENSE](LICENSE).

This project depends on [`extract-msg`](https://github.com/TeamMsgExtractor/msg-extractor) (GPL-3.0) as a runtime dependency installed separately via pip. It is neither vendored nor redistributed here, so the MIT licence applies to all source in this repository.

## Security

Report vulnerabilities privately. See [SECURITY.md](SECURITY.md).

## Support status

Actively maintained by [@trsdn](https://github.com/trsdn). Issues and pull requests are welcome; response times are best-effort.

## Acknowledgments

- Built with Python's standard email and tkinter libraries
- Uses the excellent `extract-msg` library for OFT file parsing
- Inspired by the need for reliable OFT to EML conversion

## Changelog

See [CHANGELOG.md](CHANGELOG.md).
