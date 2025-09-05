# Sample Test Files

This directory contains sample files for testing the OFT to EML converter.

## File Types
- `.oft` files: Sample Outlook Template files for testing conversion
- `.eml` files: Expected output files for comparison testing

## Usage
These files are used by the automated test suite to verify:
- Correct conversion of various OFT formats
- Proper handling of embedded images
- Unicode content preservation
- Attachment processing

## Adding New Test Files
When adding new test files:
1. Place `.oft` files in this directory
2. Generate expected `.eml` output using the converter
3. Review the output manually to ensure correctness
4. Add corresponding test cases in the test suite

## Security Note
These are test files only and should not contain any sensitive information.