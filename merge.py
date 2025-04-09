import sys
import os
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                            QVBoxLayout, QHBoxLayout, QFileDialog, QListWidget, 
                            QWidget, QProgressBar, QMessageBox, QAbstractItemView,
                            QDialog, QScrollArea, QGroupBox, QLineEdit, QGridLayout)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyPDF2 import PdfReader, PdfWriter


class PdfPageSelectionDialog(QDialog):
    """Dialog for selecting which pages to include/exclude from each PDF"""
    def __init__(self, pdf_files, parent=None):
        super().__init__(parent)
        self.pdf_files = pdf_files
        self.page_selections = {}  # Dictionary to store selections for each file
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Page Selection")
        self.setMinimumWidth(600)
        self.setMinimumHeight(500)
        
        layout = QVBoxLayout(self)
        
        # Explanation label
        help_label = QLabel(
            "Specify which pages to include or exclude from each PDF.\n"
            "Use commas for multiple pages and hyphens for ranges.\n"
            "Examples:\n"
            "• Include pages 1, 3, 5-7: '1,3,5-7'\n"
            "• Exclude first page: '-1'\n"
            "• Exclude pages 1-2: '-1,-2' or '-1-2'\n"
            "• Include all pages: leave blank or 'all'"
        )
        help_label.setWordWrap(True)
        layout.addWidget(help_label)
        
        # Add a default option for all files
        default_layout = QHBoxLayout()
        default_label = QLabel("Default for all files:")
        self.default_input = QLineEdit("-1")  # Default to removing first page
        self.apply_default_button = QPushButton("Apply to All")
        self.apply_default_button.clicked.connect(self.apply_default_to_all)
        
        default_layout.addWidget(default_label)
        default_layout.addWidget(self.default_input, 1)
        default_layout.addWidget(self.apply_default_button)
        layout.addLayout(default_layout)
        
        # Scrollable area for file list
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(scroll_content)
        
        # Add each file
        for pdf_file in self.pdf_files:
            self.add_file_entry(pdf_file)
            
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area, 1)
        
        # Buttons
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.ok_button)
        layout.addLayout(button_layout)
        
    def add_file_entry(self, pdf_file):
        filename = os.path.basename(pdf_file)
        
        # Create a group for this file
        file_group = QGroupBox(filename)
        file_layout = QVBoxLayout(file_group)
        
        # Get total pages
        try:
            pdf_reader = PdfReader(pdf_file)
            total_pages = len(pdf_reader.pages)
            
            info_layout = QHBoxLayout()
            info_label = QLabel(f"Total pages: {total_pages}")
            preview_button = QPushButton("Preview Pages")
            preview_button.clicked.connect(lambda checked, f=pdf_file: self.preview_pages(f))
            
            info_layout.addWidget(info_label)
            info_layout.addWidget(preview_button)
            info_layout.addStretch()
            
            file_layout.addLayout(info_layout)
            
            # Page selection input
            selection_layout = QHBoxLayout()
            selection_label = QLabel("Pages to include/exclude:")
            selection_input = QLineEdit("-1")  # Default to excluding first page
            
            selection_layout.addWidget(selection_label)
            selection_layout.addWidget(selection_input, 1)
            
            file_layout.addLayout(selection_layout)
            
            # Store the input field reference
            self.page_selections[pdf_file] = selection_input
            
        except Exception as e:
            error_label = QLabel(f"Error reading PDF: {str(e)}")
            file_layout.addWidget(error_label)
        
        self.scroll_layout.addWidget(file_group)
    
    def preview_pages(self, pdf_file):
        """Show a preview dialog with thumbnails of each page"""
        try:
            preview_dialog = QDialog(self)
            preview_dialog.setWindowTitle(f"Preview: {os.path.basename(pdf_file)}")
            preview_dialog.setMinimumWidth(800)
            preview_dialog.setMinimumHeight(600)
            
            layout = QVBoxLayout(preview_dialog)
            
            # Add a scrollable area for thumbnails
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll_content = QWidget()
            thumbnail_layout = QGridLayout(scroll_content)
            
            # Load the PDF
            pdf_reader = PdfReader(pdf_file)
            total_pages = len(pdf_reader.pages)
            
            # Simple message - can be enhanced with actual thumbnails in a production app
            message = QLabel(f"This PDF contains {total_pages} pages.\n\n"
                             f"In a full implementation, this dialog would show actual thumbnails of each page.\n"
                             f"For now, please use the page numbers to make your selection.")
            message.setWordWrap(True)
            message.setAlignment(Qt.AlignCenter)
            
            thumbnail_layout.addWidget(message, 0, 0, 1, 4)
            
            # Add page number labels in a grid (4 columns)
            for i in range(total_pages):
                page_label = QLabel(f"Page {i+1}")
                page_label.setAlignment(Qt.AlignCenter)
                page_label.setStyleSheet("border: 1px solid gray; padding: 8px; margin: 2px;")
                
                row = (i // 4) + 1  # +1 because row 0 has the message
                col = i % 4
                thumbnail_layout.addWidget(page_label, row, col)
            
            scroll.setWidget(scroll_content)
            layout.addWidget(scroll)
            
            # Close button
            close_button = QPushButton("Close")
            close_button.clicked.connect(preview_dialog.accept)
            layout.addWidget(close_button)
            
            preview_dialog.exec_()
            
        except Exception as e:
            QMessageBox.warning(self, "Preview Error", f"Could not preview PDF: {str(e)}")
    
    def apply_default_to_all(self):
        """Apply the default page selection to all files"""
        default_value = self.default_input.text()
        for input_field in self.page_selections.values():
            input_field.setText(default_value)
    
    def get_selections(self):
        """Return the page selection for each file"""
        selections = {}
        for pdf_file, input_field in self.page_selections.items():
            selections[pdf_file] = input_field.text()
        return selections


class PdfMergerThread(QThread):
    """Thread for PDF merging operations to keep the GUI responsive"""
    update_progress = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, pdf_files, output_file, page_selections=None):
        super().__init__()
        # Make a deep copy of the pdf_files list to prevent any issues with the main thread
        self.pdf_files = pdf_files.copy()
        self.output_file = output_file
        self.page_selections = page_selections or {}  # Dictionary of page selections
        
        # Log the file order for debugging
        file_order = [os.path.basename(f) for f in self.pdf_files]
        print(f"Merging PDFs in order: {file_order}")
        
    def run(self):
        try:
            # Create PDF writer
            pdf_writer = PdfWriter()
            
            # Process each PDF file
            for i, pdf_file in enumerate(self.pdf_files):
                # Update progress
                progress = int((i / len(self.pdf_files)) * 100)
                self.update_progress.emit(progress)
                
                try:
                    # Create PDF reader
                    pdf_reader = PdfReader(pdf_file)
                    total_pages = len(pdf_reader.pages)
                    
                    # Get page selection for this file (default: exclude first page)
                    selection_str = self.page_selections.get(pdf_file, "-1")
                    
                    # Parse the page selection string
                    pages_to_include = self.parse_page_selection(selection_str, total_pages)
                    
                    # Add the selected pages
                    for page_num in pages_to_include:
                        pdf_writer.add_page(pdf_reader.pages[page_num])
                    
                except Exception as e:
                    self.finished_signal.emit(False, f"Error processing {os.path.basename(pdf_file)}: {str(e)}")
                    return
            
            # Write the merged PDF
            if len(pdf_writer.pages) > 0:
                with open(self.output_file, 'wb') as out_file:
                    pdf_writer.write(out_file)
                self.update_progress.emit(100)
                self.finished_signal.emit(True, f"Successfully merged {len(self.pdf_files)} PDFs into {self.output_file}")
            else:
                self.finished_signal.emit(False, "No pages to merge after applying page selections")
        except Exception as e:
            self.finished_signal.emit(False, f"Error during merge: {str(e)}")
    
    def parse_page_selection(self, selection_str, total_pages):
        """
        Parse a page selection string and return a list of page indices (0-based)
        
        Format:
        - Blank or 'all': All pages
        - '1,3,5-7': Include pages 1, 3, and 5 through 7
        - '-1,-3': Exclude pages 1 and 3
        - '-1-3': Exclude pages 1 through 3
        """
        # If empty or "all", include all pages
        if not selection_str or selection_str.lower() == "all":
            return list(range(total_pages))
        
        # Start with all pages
        all_pages = set(range(total_pages))
        pages_to_include = set()
        pages_to_exclude = set()
        
        # Parse each part of the selection (comma-separated)
        for part in selection_str.split(','):
            part = part.strip()
            if not part:
                continue
                
            # Check if it's an exclusion
            is_exclude = part.startswith('-')
            if is_exclude:
                part = part[1:]  # Remove the '-'
            
            # Check if it's a range (contains a hyphen)
            if '-' in part:
                start, end = part.split('-', 1)
                try:
                    # Convert to 0-based indices
                    start_idx = int(start) - 1 if start else 0
                    end_idx = int(end) - 1 if end else total_pages - 1
                    
                    # Ensure valid page range
                    start_idx = max(0, start_idx)
                    end_idx = min(total_pages - 1, end_idx)
                    
                    page_range = set(range(start_idx, end_idx + 1))
                    
                    if is_exclude:
                        pages_to_exclude.update(page_range)
                    else:
                        pages_to_include.update(page_range)
                except ValueError:
                    # Invalid range, ignore
                    continue
            else:
                # Single page
                try:
                    # Convert to 0-based index
                    page_idx = int(part) - 1
                    
                    # Ensure valid page
                    if 0 <= page_idx < total_pages:
                        if is_exclude:
                            pages_to_exclude.add(page_idx)
                        else:
                            pages_to_include.add(page_idx)
                except ValueError:
                    # Invalid page, ignore
                    continue
        
        # If specific includes were given, start with those
        if pages_to_include:
            result = pages_to_include
        else:
            # Otherwise, start with all pages
            result = all_pages
        
        # Remove excluded pages
        result -= pages_to_exclude
        
        # Sort and return as a list
        return sorted(result)


class PDFMergerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pdf_files = []
        self.drag_enabled = False
        self.init_ui()
        
    def init_ui(self):
        # Set window properties
        self.setWindowTitle('PDF Merger - Advanced Page Selection')
        self.setGeometry(300, 300, 700, 500)
        
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Title and description
        title_label = QLabel('PDF Merger Tool')
        title_label.setStyleSheet('font-size: 18pt; font-weight: bold;')
        title_label.setAlignment(Qt.AlignCenter)
        
        desc_label = QLabel('Merge multiple PDFs and customize which pages to include or exclude from each file.')
        desc_label.setWordWrap(True)
        desc_label.setAlignment(Qt.AlignCenter)
        
        # File selection area
        file_layout = QHBoxLayout()
        
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)
        
        button_layout = QVBoxLayout()
        
        self.add_button = QPushButton('Add PDFs')
        self.add_button.clicked.connect(self.add_pdf_files)
        
        self.remove_button = QPushButton('Remove Selected')
        self.remove_button.clicked.connect(self.remove_selected_files)
        
        self.clear_button = QPushButton('Clear All')
        self.clear_button.clicked.connect(self.clear_files)
        
        self.move_up_button = QPushButton('Move Up')
        self.move_up_button.clicked.connect(self.move_file_up)
        
        self.move_down_button = QPushButton('Move Down')
        self.move_down_button.clicked.connect(self.move_file_down)
        
        # Add sorting options
        self.sort_alpha_button = QPushButton('Sort Alphabetically')
        self.sort_alpha_button.clicked.connect(self.sort_alphabetically)
        
        self.sort_num_button = QPushButton('Sort Numerically')
        self.sort_num_button.clicked.connect(self.sort_numerically)
        
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.remove_button)
        button_layout.addWidget(self.clear_button)
        button_layout.addWidget(self.move_up_button)
        button_layout.addWidget(self.move_down_button)
        button_layout.addWidget(self.sort_alpha_button)
        button_layout.addWidget(self.sort_num_button)
        button_layout.addStretch()
        
        file_layout.addWidget(self.file_list, 3)
        file_layout.addLayout(button_layout, 1)
        
        # Custom order section
        order_layout = QHBoxLayout()
        order_label = QLabel('Custom Order:')
        
        self.drag_button = QPushButton('Enable Drag & Drop Ordering')
        self.drag_button.setCheckable(True)
        self.drag_button.toggled.connect(self.toggle_drag_drop)
        
        order_layout.addWidget(order_label)
        order_layout.addWidget(self.drag_button)
        order_layout.addStretch()
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        
        # Merge button
        self.merge_button = QPushButton('Merge PDFs')
        self.merge_button.setStyleSheet('font-size: 14pt; padding: 10px;')
        self.merge_button.clicked.connect(self.merge_pdfs)
        
        # Status label
        self.status_label = QLabel('Ready')
        self.status_label.setAlignment(Qt.AlignCenter)
        
        # Add widgets to main layout
        main_layout.addWidget(title_label)
        main_layout.addWidget(desc_label)
        main_layout.addLayout(file_layout)
        main_layout.addLayout(order_layout)
        main_layout.addWidget(self.progress_bar)
        main_layout.addWidget(self.merge_button)
        main_layout.addWidget(self.status_label)
        
        # Update UI state
        self.update_ui_state()
        
    def update_ui_state(self):
        has_files = len(self.pdf_files) > 0
        has_selection = len(self.file_list.selectedItems()) > 0
        
        self.remove_button.setEnabled(has_selection)
        self.clear_button.setEnabled(has_files)
        self.merge_button.setEnabled(has_files)
        self.move_up_button.setEnabled(has_selection)
        self.move_down_button.setEnabled(has_selection)
        self.sort_alpha_button.setEnabled(has_files)
        self.sort_num_button.setEnabled(has_files)
        
    def add_pdf_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "Select PDF Files", 
            "", 
            "PDF Files (*.pdf)"
        )
        
        if files:
            for file in files:
                if file not in self.pdf_files:
                    self.pdf_files.append(file)
                    self.file_list.addItem(os.path.basename(file))
            
            self.update_ui_state()
            self.status_label.setText(f"{len(self.pdf_files)} PDFs selected")
    
    def remove_selected_files(self):
        selected_rows = [self.file_list.row(item) for item in self.file_list.selectedItems()]
        
        # Remove from bottom to top to avoid index changes
        for row in sorted(selected_rows, reverse=True):
            self.file_list.takeItem(row)
            self.pdf_files.pop(row)
        
        self.update_ui_state()
        self.status_label.setText(f"{len(self.pdf_files)} PDFs selected")
    
    def clear_files(self):
        self.file_list.clear()
        self.pdf_files.clear()
        self.update_ui_state()
        self.status_label.setText("Ready")
    
    def move_file_up(self):
        current_row = self.file_list.currentRow()
        if current_row > 0:
            # Swap items in both the list widget and the files list
            self.file_list.insertItem(current_row - 1, self.file_list.takeItem(current_row))
            self.pdf_files[current_row], self.pdf_files[current_row - 1] = self.pdf_files[current_row - 1], self.pdf_files[current_row]
            self.file_list.setCurrentRow(current_row - 1)
    
    def move_file_down(self):
        current_row = self.file_list.currentRow()
        if current_row < self.file_list.count() - 1:
            # Swap items in both the list widget and the files list
            self.file_list.insertItem(current_row + 1, self.file_list.takeItem(current_row))
            self.pdf_files[current_row], self.pdf_files[current_row + 1] = self.pdf_files[current_row + 1], self.pdf_files[current_row]
            self.file_list.setCurrentRow(current_row + 1)
    
    def sort_alphabetically(self):
        """Sort PDF files alphabetically by filename"""
        if not self.pdf_files:
            return
            
        # Create pairs of (filename, filepath)
        file_pairs = [(os.path.basename(path), path) for path in self.pdf_files]
        
        # Sort by filename
        file_pairs.sort(key=lambda x: x[0].lower())
        
        # Update lists with a new list
        self.pdf_files = [pair[1] for pair in file_pairs]
        
        # Clear and repopulate the list widget
        self.refresh_file_list()
        
        # Update status
        self.status_label.setText("Files sorted alphabetically")
        print(f"Sorted files alphabetically. New order: {[os.path.basename(f) for f in self.pdf_files]}")
    
    def sort_numerically(self):
        """Sort PDF files numerically based on numbers in filenames"""
        if not self.pdf_files:
            return
            
        import re
        
        def extract_number(filename):
            # Extract numbers from filename for numerical sorting
            numbers = re.findall(r'\d+', filename)
            return int(numbers[0]) if numbers else 0
            
        # Create pairs of (filename, filepath)
        file_pairs = [(os.path.basename(path), path) for path in self.pdf_files]
        
        # Try to sort numerically by the first number in each filename
        try:
            file_pairs.sort(key=lambda x: extract_number(x[0]))
            
            # Store the old order for debugging
            old_order = [os.path.basename(f) for f in self.pdf_files]
            
            # Create a completely new list
            self.pdf_files = [pair[1] for pair in file_pairs]
            
            # Clear and repopulate the list widget
            self.refresh_file_list()
            
            # Log the change for debugging
            new_order = [os.path.basename(f) for f in self.pdf_files]
            print(f"Numeric sort: Changed order from {old_order} to {new_order}")
            
            self.status_label.setText("Files sorted numerically")
        except Exception as e:
            QMessageBox.warning(self, "Sorting Error", 
                               f"Could not sort files numerically: {str(e)}")
    
    def refresh_file_list(self):
        """Refresh the file list widget to match the current pdf_files list"""
        # Temporarily block signals to prevent unnecessary updates
        self.file_list.blockSignals(True)
        
        # Clear the list
        self.file_list.clear()
        
        # Add all items in the current order
        for file_path in self.pdf_files:
            self.file_list.addItem(os.path.basename(file_path))
        
        # Re-enable signals
        self.file_list.blockSignals(False)
    
    def toggle_drag_drop(self, enabled):
        """Toggle drag and drop functionality for manual ordering"""
        self.drag_enabled = enabled
        
        if enabled:
            self.file_list.setDragDropMode(QAbstractItemView.InternalMove)
            self.file_list.setSelectionMode(QListWidget.SingleSelection)
            self.status_label.setText("Drag and drop mode enabled. Drag files to reorder.")
            
            # Connect the model's row moved signal to update the files list
            self.file_list.model().rowsMoved.connect(self.update_pdf_files_from_list)
            
            # Disable other ordering buttons
            self.move_up_button.setEnabled(False)
            self.move_down_button.setEnabled(False)
            self.sort_alpha_button.setEnabled(False)
            self.sort_num_button.setEnabled(False)
        else:
            self.file_list.setDragDropMode(QAbstractItemView.NoDragDrop)
            self.file_list.setSelectionMode(QListWidget.ExtendedSelection)
            self.status_label.setText("Ready")
            
            # Disconnect the signal to prevent unnecessary updates
            try:
                self.file_list.model().rowsMoved.disconnect(self.update_pdf_files_from_list)
            except:
                pass
            
            # Make sure to update the PDF files list one last time
            self.update_pdf_files_from_list()
            
            # Re-enable other ordering buttons
            has_files = len(self.pdf_files) > 0
            has_selection = len(self.file_list.selectedItems()) > 0
            
            self.move_up_button.setEnabled(has_selection)
            self.move_down_button.setEnabled(has_selection)
            self.sort_alpha_button.setEnabled(has_files)
            self.sort_num_button.setEnabled(has_files)
    
    def update_pdf_files_from_list(self):
        """Update the pdf_files list to match the current order in the list widget"""
        # Create a mapping from filenames to full paths
        filename_to_path = {os.path.basename(path): path for path in self.pdf_files}
        
        # Create new ordered list based on current list widget order
        new_order = []
        for i in range(self.file_list.count()):
            filename = self.file_list.item(i).text()
            if filename in filename_to_path:
                new_order.append(filename_to_path[filename])
        
        # Update pdf_files list
        if len(new_order) == len(self.pdf_files):
            self.pdf_files = new_order.copy()  # Make a copy to ensure a new list is created
            self.status_label.setText("File order updated")
    
    def merge_pdfs(self):
        if not self.pdf_files:
            QMessageBox.warning(self, "Warning", "No PDF files selected.")
            return
        
        # If in drag-drop mode, make sure to update the file order first
        if self.drag_enabled:
            self.update_pdf_files_from_list()
        
        # Open the page selection dialog
        page_dialog = PdfPageSelectionDialog(self.pdf_files, self)
        if page_dialog.exec_() != QDialog.Accepted:
            return
            
        # Get the page selections
        page_selections = page_dialog.get_selections()
        
        # Show the current order for confirmation
        file_order = "\n".join([
            f"{i+1}. {os.path.basename(path)} - Pages: {page_selections.get(path, '-1')}" 
            for i, path in enumerate(self.pdf_files)
        ])
        
        confirm = QMessageBox.question(
            self, 
            "Confirm Order and Page Selection",
            f"Files will be merged in this order with the selected pages:\n\n{file_order}\n\nProceed?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if confirm != QMessageBox.Yes:
            return
        
        # Get output file path
        output_file, _ = QFileDialog.getSaveFileName(
            self, 
            "Save Merged PDF", 
            "", 
            "PDF Files (*.pdf)"
        )
        
        if not output_file:
            return
            
        # Ensure .pdf extension
        if not output_file.lower().endswith('.pdf'):
            output_file += '.pdf'
        
        # Disable UI during processing
        self.set_ui_enabled(False)
        self.status_label.setText("Merging PDFs...")
        
        # Start merge thread
        self.merge_thread = PdfMergerThread(
            self.pdf_files.copy(), 
            output_file,
            page_selections
        )
        self.merge_thread.update_progress.connect(self.progress_bar.setValue)
        self.merge_thread.finished_signal.connect(self.merge_completed)
        self.merge_thread.start()
    
    def merge_completed(self, success, message):
        # Re-enable UI
        self.set_ui_enabled(True)
        
        # Show result message
        if success:
            QMessageBox.information(self, "Success", message)
            self.status_label.setText("Merge completed successfully")
        else:
            QMessageBox.critical(self, "Error", message)
            self.status_label.setText("Merge failed")
    
    def set_ui_enabled(self, enabled):
        # Enable/disable UI elements during processing
        self.add_button.setEnabled(enabled)
        self.remove_button.setEnabled(enabled)
        self.clear_button.setEnabled(enabled)
        self.move_up_button.setEnabled(enabled and not self.drag_enabled)
        self.move_down_button.setEnabled(enabled and not self.drag_enabled)
        self.sort_alpha_button.setEnabled(enabled and not self.drag_enabled)
        self.sort_num_button.setEnabled(enabled and not self.drag_enabled)
        self.drag_button.setEnabled(enabled)
        self.merge_button.setEnabled(enabled)
        self.file_list.setEnabled(enabled)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFMergerApp()
    window.show()
    sys.exit(app.exec_())