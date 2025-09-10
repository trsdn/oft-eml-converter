# Contributing to OFT to EML Converter

Thank you for considering contributing to this project! We welcome contributions from everyone.

## Code of Conduct

This project adheres to a code of conduct. By participating, you are expected to uphold this code:
- Use welcoming and inclusive language
- Be respectful of differing viewpoints and experiences
- Gracefully accept constructive criticism
- Focus on what is best for the community

## How to Contribute

### Reporting Bugs

Before submitting a bug report:
1. Check the [existing issues](https://github.com/trsdn/oft-eml-converter/issues) to avoid duplicates
2. Use the latest version of the software
3. Test with multiple OFT files to verify the issue

When submitting a bug report, include:
- Operating system and version
- Python version
- Full error message and stack trace
- Steps to reproduce the issue
- Sample OFT file (if possible and not confidential)

### Suggesting Features

Feature requests are welcome! When suggesting a feature:
1. Check existing issues and discussions
2. Explain the use case and benefit
3. Consider the scope - keep it focused
4. Be prepared to help implement it

### Development Process

1. **Fork and Clone**
   ```bash
   git clone https://github.com/yourusername/oft-eml-converter.git
   cd oft-eml-converter
   ```

2. **Set Up Environment**
   ```bash
   python -m venv venv
   source venv/bin/activate  # Windows: venv\Scripts\activate
   pip install -r requirements.txt
   ```

3. **Create a Branch**
   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b bugfix/issue-description
   ```

4. **Make Changes**
   - Write clean, documented code
   - Follow existing code style
   - Add tests for new functionality
   - Update documentation as needed

5. **Test Your Changes**
   ```bash
   # Run the test suite
   pytest tests/ -v --cov=.
   
   # Run specific test modules
   python -m tests.test_converter
   
   # Test the GUI (if applicable)
   python oft_to_eml_gui.py
   
   # Test CLI functionality
   python oft_to_eml_converter.py sample.oft
   ```

6. **Check Code Quality**
   ```bash
   # Run linting (if available)
   flake8 . --count --select=E9,F63,F7,F82 --show-source --statistics
   
   # Check formatting
   black --check .
   
   # Check imports
   isort --check-only .
   ```

7. **Commit and Push**
   ```bash
   git add .
   git commit -m "feat: add support for XYZ format"
   git push origin feature/your-feature-name
   ```

8. **Submit Pull Request**
   - Create a pull request from your branch
   - Use a clear, descriptive title
   - Explain what the PR does and why
   - Reference any related issues

## Code Style Guidelines

### Python Code Style

- Follow [PEP 8](https://pep8.org/) style guidelines
- Use 4 spaces for indentation
- Maximum line length of 88 characters (Black formatter standard)
- Use meaningful variable and function names

### Docstring Style

```python
def convert_oft_to_eml(oft_file_path, eml_file_path=None):
    """
    Convert an OFT file to EML format.
    
    Args:
        oft_file_path (str): Path to the input OFT file
        eml_file_path (str): Path to the output EML file (optional)
        
    Returns:
        str: Path to the created EML file
        
    Raises:
        FileNotFoundError: If the input file doesn't exist
        ValueError: If the file format is invalid
    """
```

### Commit Message Format

Use conventional commit format:
- `feat:` new feature
- `fix:` bug fix
- `docs:` documentation changes
- `test:` adding or updating tests
- `refactor:` code refactoring
- `style:` formatting changes
- `chore:` maintenance tasks

Examples:
- `feat: add support for RTF attachments`
- `fix: handle unicode characters in subject lines`
- `docs: update installation instructions`

## Testing Guidelines

### Writing Tests

- Add tests for all new functionality
- Use descriptive test names: `test_conversion_with_unicode_content`
- Include both positive and negative test cases
- Mock external dependencies when appropriate

### Test Structure

```python
def test_feature_description(self):
    """Test that feature works correctly under specific conditions."""
    # Arrange
    input_data = create_test_data()
    
    # Act
    result = function_under_test(input_data)
    
    # Assert
    self.assertEqual(expected_result, result)
```

### Running Tests

Our CI/CD pipeline runs tests on:
- **Operating Systems**: Ubuntu, Windows, macOS
- **Python Versions**: 3.8, 3.9, 3.10, 3.11, 3.12

Make sure your changes work across all supported platforms.

## Documentation

### README Updates

When adding features, update the README.md:
- Add new features to the Features section
- Update usage examples if needed
- Update requirements if dependencies change

### Code Documentation

- Add docstrings to all public functions and classes
- Comment complex logic within functions
- Update type hints where applicable

## Release Process

Releases are handled by maintainers:

1. Version bumping follows [Semantic Versioning](https://semver.org/)
2. Changes are documented in the README changelog
3. GitHub releases include detailed release notes
4. Tags are created for each release

## Getting Help

If you need help:
1. Check the [documentation](README.md)
2. Look through [existing issues](https://github.com/trsdn/oft-eml-converter/issues)
3. Ask questions in a new issue with the "question" label
4. Reach out to maintainers

## Recognition

Contributors will be recognized in:
- GitHub contributor list
- Release notes for significant contributions
- Special thanks in documentation for major features

Thank you for contributing! 🎉