📄 PDF Maestro

AI-powered PDF manipulation made effortless

PDF Maestro is a powerful, intuitive GUI application that makes advanced PDF operations simple. Born from a collaboration with Claude AI, this tool transforms hours of manual PDF manipulation into a task that takes just seconds.
✨ Features

Smart Page Selection - Precisely control which pages to include or exclude using an intuitive syntax
Flexible Ordering - Arrange PDFs through drag-and-drop, alphabetical, or numerical sorting
Page Preview - View page counts and structure before processing
Batch Processing - Apply the same page selection pattern to multiple PDFs simultaneously
Multi-threaded Engine - Process large PDFs without freezing the interface

🚀 Created with AI
This entire application was developed in collaboration with Claude AI from Anthropic. What would have been days of manual coding was accomplished in minutes through natural language conversation.
🔧 Installation
bash# Clone the repository
git clone https://github.com/yourusername/pdf-maestro.git
cd pdf-maestro

# Install dependencies
pip install PyQt5 PyPDF2
💻 Usage
bashpython pdf_merger_gui.py
Page Selection Syntax

Include specific pages: 1,3,5-7 (includes pages 1, 3, and 5 through 7)
Exclude specific pages: -1,-3 (excludes pages 1 and 3)
Exclude page ranges: -1-3 (excludes pages 1 through 3)
Include all pages: leave blank or type all


📝 License
This project is licensed under the MIT License - see the LICENSE file for details.
🙏 Acknowledgements

Special thanks to Claude AI from Anthropic for assisting in the development
PyQt5 for the GUI framework
PyPDF2 for PDF manipulation capabilities