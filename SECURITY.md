# Security Policy

## Supported Versions

We provide security updates for the following versions:

| Version | Supported          |
| ------- | ------------------ |
| 1.1.x   | :white_check_mark: |
| 1.0.x   | :white_check_mark: |
| < 1.0   | :x:                |

## Reporting a Vulnerability

We take security seriously. If you discover a security vulnerability, please follow these steps:

### For Security Issues

**DO NOT** create a public GitHub issue for security vulnerabilities.

Instead, please report security issues by:

1. **Email**: Send details to the repository maintainer
2. **GitHub Security Advisory**: Use GitHub's private vulnerability reporting feature
3. **Include**: 
   - Description of the vulnerability
   - Steps to reproduce
   - Potential impact
   - Suggested fix (if any)

### What We Commit To

- **Acknowledgment**: We'll acknowledge receipt within 48 hours
- **Assessment**: Initial assessment within 5 business days
- **Updates**: Regular updates on investigation progress
- **Fix Timeline**: Security fixes will be prioritized and released as soon as possible
- **Credit**: Security researchers will be credited (unless they prefer anonymity)

## Security Considerations

### File Handling

- The application processes OFT files which could potentially contain malicious content
- Always scan files with antivirus before processing
- Be cautious when processing files from untrusted sources
- The application extracts embedded content - ensure your system is protected

### Data Privacy

- OFT files may contain sensitive email content
- Converted EML files preserve all original content including attachments
- Consider the security implications of file storage and transmission
- No data is transmitted to external servers during conversion

### Dependencies

- We regularly update dependencies to address security vulnerabilities
- The main dependency `extract-msg` is actively maintained
- Monitor our CI/CD pipeline for dependency security checks

### Best Practices

When using this tool:

1. **Validate Input**: Only process OFT files from trusted sources
2. **Scan Files**: Run antivirus scans on input files
3. **Secure Storage**: Store converted files securely
4. **Access Control**: Limit access to sensitive converted content
5. **Updates**: Keep the tool updated to the latest version

## Known Security Considerations

### File Parsing Risks

- OFT files use Microsoft's compound document format
- Malformed files could potentially cause parsing errors
- We use the trusted `extract-msg` library for parsing
- Error handling prevents crashes from malformed input

### Attachment Handling

- Embedded attachments are preserved in EML output
- Malicious attachments remain malicious after conversion
- Consider scanning attachments separately
- Be aware of file type restrictions in your email system

## Vulnerability History

No security vulnerabilities have been reported to date.

## Security Contact

For security-related questions or concerns, please contact the maintainers through:
- GitHub Security Advisory (preferred)
- Repository issues (for non-sensitive security questions)

## Security Tools

Our CI/CD pipeline includes:
- Dependency vulnerability scanning
- Static code analysis
- Cross-platform testing to identify platform-specific issues

Thank you for helping keep our project secure! 🔒