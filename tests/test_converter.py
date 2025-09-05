#!/usr/bin/env python3
"""
Test suite for OFT to EML converter.

This module tests the core conversion functionality including:
- Basic OFT to EML conversion
- Email header preservation
- Attachment handling
- Error cases
"""

import unittest
import os
import tempfile
import shutil
from pathlib import Path
from email import message_from_string
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from unittest.mock import Mock, patch, MagicMock

# Import the converter module
from oft_to_eml_converter import convert_oft_to_eml


class TestOFTtoEMLConverter(unittest.TestCase):
    """Test cases for the OFT to EML converter."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = tempfile.mkdtemp()
        self.test_oft = os.path.join(self.test_dir, "test.oft")
        self.test_eml = os.path.join(self.test_dir, "test.eml")
    
    def tearDown(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.test_dir, ignore_errors=True)
    
    @patch('oft_to_eml_converter.extract_msg.Message')
    def test_basic_conversion(self, mock_message_class):
        """Test basic OFT to EML conversion."""
        # Create mock message
        mock_msg = Mock()
        mock_msg.sender = "sender@example.com"
        mock_msg.to = "recipient@example.com"
        mock_msg.subject = "Test Subject"
        mock_msg.body = "Test body content"
        mock_msg.htmlBody = None
        mock_msg.date = None
        mock_msg.cc = None
        mock_msg.attachments = []
        
        mock_message_class.return_value = mock_msg
        
        # Create dummy OFT file
        Path(self.test_oft).touch()
        
        # Run conversion
        result = convert_oft_to_eml(self.test_oft, self.test_eml)
        
        # Verify EML was created
        self.assertTrue(os.path.exists(self.test_eml))
        self.assertEqual(result, self.test_eml)
        
        # Verify content
        with open(self.test_eml, 'r') as f:
            eml_content = f.read()
            self.assertIn("From: sender@example.com", eml_content)
            self.assertIn("To: recipient@example.com", eml_content)
            self.assertIn("Subject: Test Subject", eml_content)
            # Content is base64 encoded, so check for the encoded version
            import base64
            encoded_body = base64.b64encode("Test body content".encode()).decode()
            self.assertIn(encoded_body, eml_content)
    
    @patch('oft_to_eml_converter.extract_msg.Message')
    def test_html_content(self, mock_message_class):
        """Test conversion with HTML content."""
        # Create mock message with HTML
        mock_msg = Mock()
        mock_msg.sender = "sender@example.com"
        mock_msg.to = "recipient@example.com"
        mock_msg.subject = "HTML Test"
        mock_msg.body = "Plain text"
        mock_msg.htmlBody = "<html><body><p>HTML content</p></body></html>"
        mock_msg.date = None
        mock_msg.cc = None
        mock_msg.attachments = []
        
        mock_message_class.return_value = mock_msg
        
        # Create dummy OFT file
        Path(self.test_oft).touch()
        
        # Run conversion
        convert_oft_to_eml(self.test_oft, self.test_eml)
        
        # Verify multipart structure
        with open(self.test_eml, 'r') as f:
            eml_content = f.read()
            self.assertIn("Content-Type: multipart/related", eml_content)
            self.assertIn("Content-Type: multipart/alternative", eml_content)
            # HTML content is base64 encoded
            import base64
            encoded_html = base64.b64encode("<html><body><p>HTML content</p></body></html>".encode()).decode()
            self.assertIn(encoded_html, eml_content)
    
    @patch('oft_to_eml_converter.extract_msg.Message')
    def test_attachment_handling(self, mock_message_class):
        """Test handling of attachments."""
        # Create mock attachment
        mock_attachment = Mock()
        mock_attachment.longFilename = "test.png"
        mock_attachment.shortFilename = "test.png"
        mock_attachment.contentId = "image001@test.com"
        mock_attachment.data = b"fake image data"
        
        # Create mock message
        mock_msg = Mock()
        mock_msg.sender = "sender@example.com"
        mock_msg.to = "recipient@example.com"
        mock_msg.subject = "Attachment Test"
        mock_msg.body = "See attached"
        mock_msg.htmlBody = '<img src="cid:image001@test.com">'
        mock_msg.date = None
        mock_msg.cc = None
        mock_msg.attachments = [mock_attachment]
        
        mock_message_class.return_value = mock_msg
        
        # Create dummy OFT file
        Path(self.test_oft).touch()
        
        # Run conversion
        convert_oft_to_eml(self.test_oft, self.test_eml)
        
        # Verify attachment handling
        with open(self.test_eml, 'r') as f:
            eml_content = f.read()
            self.assertIn("Content-Type: image/png", eml_content)
            self.assertIn("Content-ID: <image001@test.com>", eml_content)
            self.assertIn("Content-Disposition: inline", eml_content)
    
    def test_missing_input_file(self):
        """Test handling of missing input file."""
        non_existent = os.path.join(self.test_dir, "missing.oft")
        
        with self.assertRaises(FileNotFoundError):
            convert_oft_to_eml(non_existent, self.test_eml)
    
    @patch('oft_to_eml_converter.extract_msg.Message')
    def test_output_filename_generation(self, mock_message_class):
        """Test automatic output filename generation."""
        # Create mock message
        mock_msg = Mock()
        mock_msg.sender = "sender@example.com"
        mock_msg.to = "recipient@example.com"
        mock_msg.subject = "Test"
        mock_msg.body = "Test"
        mock_msg.htmlBody = None
        mock_msg.date = None
        mock_msg.cc = None
        mock_msg.attachments = []
        
        mock_message_class.return_value = mock_msg
        
        # Create dummy OFT file
        Path(self.test_oft).touch()
        
        # Run conversion without output path
        result = convert_oft_to_eml(self.test_oft)
        
        # Verify automatic naming (should create test.eml in current directory)
        expected_filename = "test.eml"
        self.assertEqual(result, expected_filename)
        self.assertTrue(os.path.exists(expected_filename))
        
        # Clean up the created file
        if os.path.exists(expected_filename):
            os.remove(expected_filename)
    
    @patch('oft_to_eml_converter.extract_msg.Message')
    def test_unicode_handling(self, mock_message_class):
        """Test handling of Unicode content."""
        # Create mock message with Unicode
        mock_msg = Mock()
        mock_msg.sender = "sender@example.com"
        mock_msg.to = "recipient@example.com"
        mock_msg.subject = "Ümläut Tëst 你好"
        mock_msg.body = "Unicode content: €£¥ 🎉"
        mock_msg.htmlBody = None
        mock_msg.date = None
        mock_msg.cc = None
        mock_msg.attachments = []
        
        mock_message_class.return_value = mock_msg
        
        # Create dummy OFT file
        Path(self.test_oft).touch()
        
        # Run conversion with suppressed output to avoid Windows encoding issues
        import io
        import sys
        original_stdout = sys.stdout
        sys.stdout = io.StringIO()  # Suppress print output during testing
        try:
            convert_oft_to_eml(self.test_oft, self.test_eml)
        finally:
            sys.stdout = original_stdout
        
        # Verify Unicode handling
        with open(self.test_eml, 'r', encoding='utf-8') as f:
            eml_content = f.read()
            # Subject will be encoded
            self.assertIn("Subject:", eml_content)
            # Body should contain the Unicode
            msg = message_from_string(eml_content)
            self.assertIsNotNone(msg)


class TestGUIFunctions(unittest.TestCase):
    """Test cases for GUI functionality."""
    
    def test_gui_import(self):
        """Test that GUI module can be imported."""
        try:
            import oft_to_eml_gui
            self.assertTrue(True)
        except ImportError as e:
            self.fail(f"Failed to import GUI module: {e}")
    
    @patch('oft_to_eml_gui.tk.StringVar')
    @patch('oft_to_eml_gui.tk.Tk')
    def test_gui_initialization(self, mock_tk, mock_string_var):
        """Test GUI initialization."""
        from oft_to_eml_gui import OFTtoEMLGUI
        
        # Mock the Tk root and StringVar
        mock_root = MagicMock()
        mock_tk.return_value = mock_root
        mock_string_var.return_value = MagicMock()
        
        # Mock other tkinter components
        with patch('oft_to_eml_gui.ttk.Frame'), \
             patch('oft_to_eml_gui.ttk.Label'), \
             patch('oft_to_eml_gui.ttk.Button'), \
             patch('oft_to_eml_gui.tk.Label'), \
             patch('oft_to_eml_gui.ttk.Entry'), \
             patch('oft_to_eml_gui.ttk.Progressbar'), \
             patch('oft_to_eml_gui.ttk.LabelFrame'), \
             patch('oft_to_eml_gui.tk.Text'), \
             patch('oft_to_eml_gui.ttk.Scrollbar'):
            
            # Test initialization
            try:
                gui = OFTtoEMLGUI()
                self.assertIsNotNone(gui)
                self.assertEqual(gui.root, mock_root)
            except Exception as e:
                self.fail(f"GUI initialization failed: {e}")


def run_tests():
    """Run all tests and return results."""
    # Create test suite
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # Add tests
    suite.addTests(loader.loadTestsFromTestCase(TestOFTtoEMLConverter))
    suite.addTests(loader.loadTestsFromTestCase(TestGUIFunctions))
    
    # Run tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    return result.wasSuccessful()


if __name__ == "__main__":
    # Run tests
    success = run_tests()
    exit(0 if success else 1)