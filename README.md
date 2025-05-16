# Block Search for macOS

A specialized document search tool for debaters to efficiently search, manage, and organize debate blocks and evidence files.

## Features

- Search through Word documents (.docx) for specific content
- Convert between JSON and Word document formats
- Process and organize debate files into structured sections
- Intuitive GUI interface built with Tkinter
- macOS optimized with native support

## Requirements

- Python 3.6+
- macOS operating system
- Pandoc (version 2.11 or later recommended)
- Python packages listed in requirements.txt

## Installation

1. Clone this repository
2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Install Pandoc:
   - When running the Python script directly, Pandoc must be installed on your system
   - Install Pandoc via Homebrew: `brew install pandoc`
   - Or download from the [official Pandoc website](https://pandoc.org/installing.html)
   - Ensure Pandoc is available in your system PATH
   
   Note: When built as a macOS application with PyInstaller, Pandoc is typically bundled with the application

## Usage

Run the application:

```
python BlockSearch-Mac.py
```

The application provides a GUI interface for:
- Selecting documents to search through
- Converting between document formats
- Managing debate blocks and evidence

## Building a macOS Application

This repository includes the necessary files for packaging the application as a macOS app:

- `tk_runtime_hook.py`: Fixes Tkinter GUI issues specific to macOS packaged applications
- `block_sender_icon.ico`: Application icon file