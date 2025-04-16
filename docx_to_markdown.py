#!/usr/bin/env python3
"""
DOCX to Markdown Converter
--------------------------
A tool to convert DOCX files to Markdown format using Pandoc.
"""

import os
import sys
import argparse
import subprocess
from pathlib import Path
import threading
from tqdm import tqdm
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QPushButton, QFileDialog, QListWidget, QProgressBar,
                             QLineEdit, QFormLayout, QGroupBox, QMessageBox, QSizePolicy,
                             QAbstractItemView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize

class DocxToMarkdownConverter:
    """Handles conversion of DOCX files to Markdown using Pandoc."""
    
    def __init__(self):
        """Initialize the converter."""
        self._check_pandoc()
    
    def _check_pandoc(self):
        """Check if Pandoc is installed and accessible."""
        try:
            result = subprocess.run(['pandoc', '--version'], 
                                   stdout=subprocess.PIPE, 
                                   stderr=subprocess.PIPE, 
                                   text=True)
            if result.returncode != 0:
                raise Exception("Pandoc is installed but not working properly.")
        except FileNotFoundError:
            raise Exception("Pandoc not found. Please install Pandoc: https://pandoc.org/installing.html")
    
    def convert_file(self, input_file, output_file=None, extra_args=None):
        """
        Convert a single DOCX file to Markdown.
        
        Args:
            input_file (str): Path to the input DOCX file
            output_file (str, optional): Path to the output Markdown file
            extra_args (list, optional): Additional arguments to pass to Pandoc
            
        Returns:
            bool: True if conversion was successful, False otherwise
        """
        if not os.path.exists(input_file):
            print(f"Error: File not found: {input_file}")
            return False
            
        if not output_file:
            # Create output filename by replacing .docx with .md
            output_file = os.path.splitext(input_file)[0] + '.md'
        
        cmd = ['pandoc', input_file, '-o', output_file, '--wrap=none']
        
        if extra_args:
            cmd.extend(extra_args)
        
        try:
            result = subprocess.run(cmd, 
                                   stdout=subprocess.PIPE, 
                                   stderr=subprocess.PIPE, 
                                   text=True)
            
            if result.returncode != 0:
                print(f"Error converting {input_file}: {result.stderr}")
                return False
                
            return True
        except Exception as e:
            print(f"Exception while converting {input_file}: {str(e)}")
            return False
    
    def batch_convert(self, input_files, output_dir=None, extra_args=None):
        """
        Convert multiple DOCX files to Markdown.
        
        Args:
            input_files (list): List of paths to input DOCX files
            output_dir (str, optional): Directory to save output files
            extra_args (list, optional): Additional arguments to pass to Pandoc
            
        Returns:
            tuple: (successful_conversions, failed_conversions)
        """
        successful = []
        failed = []
        
        for input_file in tqdm(input_files, desc="Converting files"):
            if not input_file.lower().endswith('.docx'):
                print(f"Skipping non-DOCX file: {input_file}")
                continue
                
            if output_dir:
                # Create output filename in the specified directory
                filename = os.path.basename(input_file)
                output_file = os.path.join(output_dir, os.path.splitext(filename)[0] + '.md')
            else:
                output_file = None  # Let convert_file determine the output path
            
            if self.convert_file(input_file, output_file, extra_args):
                successful.append(input_file)
            else:
                failed.append(input_file)
        
        return successful, failed


# Worker thread for conversion process
class ConversionWorker(QThread):
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    conversion_complete = pyqtSignal(list, list)
    error_occurred = pyqtSignal(str)
    
    def __init__(self, converter, files, output_dir=None, extra_args=None):
        super().__init__()
        self.converter = converter
        self.files = files
        self.output_dir = output_dir
        self.extra_args = extra_args
        
    def run(self):
        try:
            self.status_updated.emit("Converting files...")
            total_files = len(self.files)
            
            successful = []
            failed = []
            
            for i, input_file in enumerate(self.files):
                # Update progress
                progress = int((i / total_files) * 100)
                self.progress_updated.emit(progress)
                
                if self.output_dir:
                    # Create output filename in the specified directory
                    filename = os.path.basename(input_file)
                    output_file = os.path.join(self.output_dir, 
                                             os.path.splitext(filename)[0] + '.md')
                else:
                    output_file = None  # Let convert_file determine the output path
                
                if self.converter.convert_file(input_file, output_file, self.extra_args):
                    successful.append(input_file)
                else:
                    failed.append(input_file)
            
            # Set progress to 100% when done
            self.progress_updated.emit(100)
            self.conversion_complete.emit(successful, failed)
            
        except Exception as e:
            self.error_occurred.emit(str(e))


class DocxToMarkdownGUI(QMainWindow):
    """GUI interface for the DOCX to Markdown converter using PyQt5."""
    
    def __init__(self):
        super().__init__()
        self.converter = DocxToMarkdownConverter()
        self.selected_files = []
        self.output_directory = None
        self.init_ui()
    
    def init_ui(self):
        """Initialize the user interface."""
        # Set window properties
        self.setWindowTitle("DOCX to Markdown Converter")
        self.setMinimumSize(700, 500)
        self.resize(800, 600)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Title label
        title_label = QLabel("DOCX to Markdown Converter")
        title_label.setStyleSheet("font-size: 18pt; font-weight: bold; margin: 10px;")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # Button row
        button_layout = QHBoxLayout()
        
        self.select_files_btn = QPushButton("Select Files")
        self.select_files_btn.clicked.connect(self.select_files)
        button_layout.addWidget(self.select_files_btn)
        
        self.select_folder_btn = QPushButton("Select Folder")
        self.select_folder_btn.clicked.connect(self.select_folder)
        button_layout.addWidget(self.select_folder_btn)
        
        self.clear_btn = QPushButton("Clear Selection")
        self.clear_btn.clicked.connect(self.clear_selection)
        button_layout.addWidget(self.clear_btn)
        
        button_layout.addStretch(1)  # Push buttons to the left
        main_layout.addLayout(button_layout)
        
        # Files list
        files_group = QGroupBox("Selected Files")
        files_layout = QVBoxLayout(files_group)
        
        self.files_list = QListWidget()
        self.files_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.files_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        files_layout.addWidget(self.files_list)
        
        main_layout.addWidget(files_group)
        
        # Options group
        options_group = QGroupBox("Options")
        options_layout = QFormLayout(options_group)
        
        # Output directory
        output_dir_layout = QHBoxLayout()
        self.output_dir_edit = QLineEdit("Same as input files")
        output_dir_layout.addWidget(self.output_dir_edit)
        
        self.browse_btn = QPushButton("Browse")
        self.browse_btn.clicked.connect(self.select_output_dir)
        output_dir_layout.addWidget(self.browse_btn)
        
        options_layout.addRow("Output Directory:", output_dir_layout)
        
        # Extra arguments
        self.extra_args_edit = QLineEdit()
        options_layout.addRow("Extra Pandoc Args:", self.extra_args_edit)
        
        main_layout.addWidget(options_group)
        
        # Convert button
        self.convert_btn = QPushButton("Convert Files")
        self.convert_btn.clicked.connect(self.convert_files)
        self.convert_btn.setMinimumHeight(40)  # Make button taller
        main_layout.addWidget(self.convert_btn)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)
        
        # Status bar
        self.statusBar().showMessage("Ready")
    
    def select_files(self):
        """Open file dialog to select DOCX files."""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select DOCX Files",
            "",
            "Word Documents (*.docx);;All Files (*.*)"
        )
        
        if files:
            self.add_files(files)
    
    def select_folder(self):
        """Open folder dialog to select a folder containing DOCX files."""
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Folder Containing DOCX Files"
        )
        
        if folder:
            docx_files = []
            for root, _, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith('.docx'):
                        docx_files.append(os.path.join(root, file))
            
            if docx_files:
                self.add_files(docx_files)
            else:
                QMessageBox.information(
                    self,
                    "No DOCX Files",
                    f"No DOCX files found in {folder}"
                )
    
    def add_files(self, files):
        """Add files to the listbox and selected_files list."""
        for file in files:
            if file not in self.selected_files:
                self.selected_files.append(file)
                self.files_list.addItem(os.path.basename(file))
        
        self.statusBar().showMessage(f"{len(self.selected_files)} files selected")
    
    def clear_selection(self):
        """Clear the selected files."""
        self.selected_files = []
        self.files_list.clear()
        self.statusBar().showMessage("Ready")
    
    def select_output_dir(self):
        """Open folder dialog to select output directory."""
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Output Directory"
        )
        
        if folder:
            self.output_directory = folder
            self.output_dir_edit.setText(folder)
    
    def convert_files(self):
        """Convert the selected files."""
        if not self.selected_files:
            QMessageBox.information(
                self,
                "No Files Selected",
                "Please select at least one DOCX file to convert."
            )
            return
        
        output_dir = None
        if self.output_dir_edit.text() != "Same as input files":
            output_dir = self.output_dir_edit.text()
            
            # Create the output directory if it doesn't exist
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
        
        extra_args = None
        if self.extra_args_edit.text().strip():
            extra_args = self.extra_args_edit.text().strip().split()
        
        # Disable UI elements during conversion
        self.set_ui_enabled(False)
        
        # Start conversion in a worker thread
        self.worker = ConversionWorker(
            self.converter,
            self.selected_files,
            output_dir,
            extra_args
        )
        
        # Connect signals
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.status_updated.connect(self.update_status)
        self.worker.conversion_complete.connect(self.conversion_finished)
        self.worker.error_occurred.connect(self.handle_error)
        
        # Start the worker
        self.worker.start()
    
    def update_progress(self, value):
        """Update the progress bar."""
        self.progress_bar.setValue(value)
    
    def update_status(self, message):
        """Update the status bar."""
        self.statusBar().showMessage(message)
    
    def conversion_finished(self, successful, failed):
        """Handle completion of the conversion process."""
        # Re-enable UI elements
        self.set_ui_enabled(True)
        
        # Show results
        if failed:
            QMessageBox.warning(
                self,
                "Conversion Results",
                f"Converted {len(successful)} files successfully.\n"
                f"Failed to convert {len(failed)} files."
            )
        else:
            QMessageBox.information(
                self,
                "Conversion Complete",
                f"Successfully converted all {len(successful)} files."
            )
        
        self.statusBar().showMessage(f"Converted {len(successful)} files, {len(failed)} failed")
    
    def handle_error(self, error_message):
        """Handle errors during conversion."""
        # Re-enable UI elements
        self.set_ui_enabled(True)
        
        QMessageBox.critical(
            self,
            "Error",
            f"An error occurred: {error_message}"
        )
        
        self.statusBar().showMessage("Error during conversion")
    
    def set_ui_enabled(self, enabled):
        """Enable or disable UI elements."""
        self.select_files_btn.setEnabled(enabled)
        self.select_folder_btn.setEnabled(enabled)
        self.clear_btn.setEnabled(enabled)
        self.browse_btn.setEnabled(enabled)
        self.convert_btn.setEnabled(enabled)
        self.files_list.setEnabled(enabled)
        self.output_dir_edit.setEnabled(enabled)
        self.extra_args_edit.setEnabled(enabled)


def parse_arguments():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Convert DOCX files to Markdown using Pandoc."
    )
    
    parser.add_argument(
        "files", nargs="*",
        help="DOCX files to convert"
    )
    
    parser.add_argument(
        "-o", "--output-dir",
        help="Directory to save the converted Markdown files"
    )
    
    parser.add_argument(
        "-r", "--recursive", action="store_true",
        help="Recursively search for DOCX files in directories"
    )
    
    parser.add_argument(
        "-g", "--gui", action="store_true",
        help="Launch the graphical user interface"
    )
    
    parser.add_argument(
        "--pandoc-args", nargs=argparse.REMAINDER,
        help="Additional arguments to pass to Pandoc"
    )
    
    return parser.parse_args()


def find_docx_files(paths, recursive=False):
    """
    Find DOCX files in the given paths.
    
    Args:
        paths (list): List of file or directory paths
        recursive (bool): Whether to search directories recursively
        
    Returns:
        list: List of DOCX file paths
    """
    docx_files = []
    
    for path in paths:
        path = os.path.abspath(path)
        
        if os.path.isfile(path):
            if path.lower().endswith('.docx'):
                docx_files.append(path)
        elif os.path.isdir(path):
            if recursive:
                for root, _, files in os.walk(path):
                    for file in files:
                        if file.lower().endswith('.docx'):
                            docx_files.append(os.path.join(root, file))
            else:
                for file in os.listdir(path):
                    file_path = os.path.join(path, file)
                    if os.path.isfile(file_path) and file.lower().endswith('.docx'):
                        docx_files.append(file_path)
    
    return docx_files


def main():
    """Main entry point for the script."""
    args = parse_arguments()
    
    # Launch GUI if requested or if no files specified
    if args.gui or (not args.files and len(sys.argv) == 1):
        app = QApplication(sys.argv)
        window = DocxToMarkdownGUI()
        window.show()
        sys.exit(app.exec_())
        return
    
    try:
        converter = DocxToMarkdownConverter()
        
        # Find DOCX files
        if args.files:
            docx_files = find_docx_files(args.files, args.recursive)
        else:
            print("No files specified. Use --gui to launch the GUI or specify files to convert.")
            return
        
        if not docx_files:
            print("No DOCX files found.")
            return
        
        print(f"Found {len(docx_files)} DOCX files to convert.")
        
        # Convert files
        successful, failed = converter.batch_convert(
            docx_files, args.output_dir, args.pandoc_args
        )
        
        # Print results
        print(f"\nConversion complete:")
        print(f"  - Successfully converted: {len(successful)} files")
        
        if failed:
            print(f"  - Failed to convert: {len(failed)} files")
            for file in failed:
                print(f"    - {file}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
