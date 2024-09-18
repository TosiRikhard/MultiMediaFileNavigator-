# MultiMedia File Navigator

## Description
MultiMedia File Navigator is a PyQt5-based application designed to help users efficiently navigate, preview, and organize large collections of files across various formats. This tool was originally created to assist in sorting through personal files, making it particularly useful for tasks like managing inherited digital content or decluttering personal archives.

## Features
- Preview support for multiple file types:
  - Images (JPG, JPEG, PNG, GIF, BMP)
  - Documents (PDF, DOC, DOCX, ODT)
  - Spreadsheets (XLS, XLSX)
  - Presentations (PPT, PPTX)
  - Audio (MP3, WAV, OGG, FLAC)
  - Video (MP4, AVI, MOV)
- Simple file management: Delete, Move, or Skip files
- Built-in media player for audio and video files
- Keyboard shortcuts for quick navigation
- Recursive file scanning from a selected source folder

## Installation

### Prerequisites
- Python 3.6 or higher
- pip (Python package installer)

### Steps
1. Clone the repository or download the source code.
2. Navigate to the project directory in your terminal.
3. Install the required dependencies:

```
pip install PyQt5 PyMuPDF docx2txt openpyxl python-pptx odfpy
```

## Usage
1. Run the script:
```
python file_navigator.py
```
2. Select the source folder containing the files you want to navigate.
3. Choose a destination folder for moved files.
4. Use the interface buttons or keyboard shortcuts to manage files:
   - Delete (D): Remove the current file
   - Move (M): Move the file to the destination folder
   - Skip (S): Go to the next file
   - Open (O): Open the file with the default application

## Potential Improvements
1. Add support for more file types (e.g., CAD files, specific image formats)
2. Implement a tagging system for better organization
3. Add a search functionality within the navigator
4. Improve error handling and logging
5. Create a more detailed preview for text-based files (e.g., code files)
6. Add an option to customize keyboard shortcuts
7. Implement a thumbnail view for quicker visual scanning
8. Add multi-file selection for batch operations

## Contributing
Contributions to improve MultiMedia File Navigator are welcome. Please feel free to fork the repository, make changes, and submit pull requests.

## License
[MIT License](https://opensource.org/licenses/MIT)

## Acknowledgments
This project was inspired by the need to efficiently sort through inherited digital files, making the process of "keep or discard" decisions much easier and faster.