#!/usr/bin/env python3
"""
OFT to EML Converter

This script converts Outlook Template (.oft) files to standard EML format.
OFT files use the same MSG format as regular Outlook messages but with a different CLSID.

Usage:
    python oft_to_eml_converter.py <input_oft_file> [output_eml_file]
"""

import sys
import os
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import extract_msg


def convert_oft_to_eml(oft_file_path, eml_file_path=None):
    """
    Convert an OFT file to EML format.
    
    Args:
        oft_file_path (str): Path to the input OFT file
        eml_file_path (str): Path to the output EML file (optional)
        
    Returns:
        str: Path to the created EML file
    """
    
    # Validate input file
    if not os.path.exists(oft_file_path):
        raise FileNotFoundError(f"Input file not found: {oft_file_path}")
    
    # Generate output filename if not provided
    if eml_file_path is None:
        base_name = Path(oft_file_path).stem
        eml_file_path = f"{base_name}.eml"
    
    try:
        # Extract message from OFT file using extract_msg
        print(f"Reading OFT file: {oft_file_path}")
        msg = extract_msg.Message(oft_file_path)
        
        # Create MIME message - use 'related' to support inline images
        mime_msg = MIMEMultipart('related')
        
        # Set headers
        if msg.sender:
            mime_msg['From'] = msg.sender
        if msg.to:
            mime_msg['To'] = msg.to
        if msg.cc:
            mime_msg['Cc'] = msg.cc
        if msg.subject:
            mime_msg['Subject'] = msg.subject
        if msg.date:
            mime_msg['Date'] = msg.date.strftime('%a, %d %b %Y %H:%M:%S %z') if hasattr(msg.date, 'strftime') else str(msg.date)
        
        # Create alternative container for text/html content
        msg_alternative = MIMEMultipart('alternative')
        
        # Add message body
        if msg.body:
            text_part = MIMEText(msg.body, 'plain', 'utf-8')
            msg_alternative.attach(text_part)
        
        if msg.htmlBody:
            html_part = MIMEText(msg.htmlBody, 'html', 'utf-8')
            msg_alternative.attach(html_part)
        
        # Add the alternative part to the main message
        mime_msg.attach(msg_alternative)
        
        # Add attachments if any
        if msg.attachments:
            print(f"Found {len(msg.attachments)} attachments")
            for attachment in msg.attachments:
                if hasattr(attachment, 'data') and attachment.data:
                    filename = attachment.longFilename or attachment.shortFilename or "attachment"
                    
                    # Check if this is an embedded image (has Content-ID)
                    content_id = getattr(attachment, 'contentId', None)
                    
                    if content_id and filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                        # Handle as inline image
                        image_type = filename.split('.')[-1].lower()
                        if image_type == 'jpg':
                            image_type = 'jpeg'
                        
                        part = MIMEBase('image', image_type)
                        part.set_payload(attachment.data)
                        encoders.encode_base64(part)
                        
                        # Set Content-ID for inline images
                        part.add_header('Content-ID', f'<{content_id}>')
                        part.add_header('Content-Disposition', 'inline', filename=filename)
                        print(f"  Added inline image: {filename} (Content-ID: {content_id})")
                    else:
                        # Handle as regular attachment
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.data)
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                        print(f"  Added attachment: {filename}")
                    
                    mime_msg.attach(part)
        
        # Write EML file
        print(f"Writing EML file: {eml_file_path}")
        with open(eml_file_path, 'w', encoding='utf-8') as f:
            f.write(mime_msg.as_string())
        
        print(f"Conversion completed successfully!")
        print(f"Output: {eml_file_path}")
        
        # Print some info about the converted message
        print("\n--- Message Info ---")
        try:
            print(f"From: {msg.sender or 'N/A'}")
            print(f"To: {msg.to or 'N/A'}")
            print(f"Subject: {msg.subject or 'N/A'}")
            print(f"Date: {msg.date or 'N/A'}")
        except UnicodeEncodeError:
            # Handle Windows console encoding issues
            print("From:", repr(msg.sender) if msg.sender else 'N/A')
            print("To:", repr(msg.to) if msg.to else 'N/A')  
            print("Subject:", repr(msg.subject) if msg.subject else 'N/A')
            print("Date:", repr(msg.date) if msg.date else 'N/A')
        
        print(f"Body length: {len(msg.body) if msg.body else 0} chars")
        print(f"HTML body length: {len(msg.htmlBody) if msg.htmlBody else 0} chars")
        print(f"Attachments: {len(msg.attachments) if msg.attachments else 0}")
        
        return eml_file_path
        
    except Exception as e:
        print(f"Error during conversion: {str(e)}")
        raise


def main():
    """Main entry point for the script."""
    
    if len(sys.argv) < 2:
        print("Usage: python oft_to_eml_converter.py <input_oft_file> [output_eml_file]")
        print("Example: python oft_to_eml_converter.py template.oft output.eml")
        sys.exit(1)
    
    oft_file = sys.argv[1]
    eml_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        result_file = convert_oft_to_eml(oft_file, eml_file)
        print(f"\nSuccess! EML file created: {result_file}")
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()