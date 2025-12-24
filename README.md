# Image to PDF Converter Pro

A powerful, modern, and lossless Image to PDF converter for Windows. Built with Python and ttkbootstrap.

![App Icon](main.ico)

## Features

- **Professional UI**: Modern Dark Mode interface (ttkbootstrap).
- **Lossless Conversion**: Validated `Quality=100` engine.
- **Archive Support**: Supports `.zip`, `.cbz` natively. Supports `.rar`, `.cbr` via WinRAR integration.
- **Smart Logic**:
  - Auto-detects WinRAR installation.
  - Locked configuration paths (portable-friendly).
  - Parallel processing for high performance.
- **Safety First**: "Delete Source" is opt-in to prevent data loss.

## Installation

### Method 1: Standalone Exe (Recommended)

Download the latest release from the [Releases] page.

1. Download `Image2PDF_Pro.exe`.
2. Move it to a folder (e.g., in Documents).
3. Run it!

### Method 2: Run from Source

1. Clone the repo:
   ```bash
   git clone https://github.com/YOUR_USERNAME/image2pdf.git
   cd image2pdf
   ```
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run:
   ```bash
   python main.py
   ```

## Usage

1. **Drag & Drop**: Drag a folder of images (or a ZIP/CBZ file) onto the executable/script.
2. **GUI Mode**: Double-click to open the dashboard.
   - Configure **Quality** (Default: 100).
   - Set **Landscape Mode** (None, Letterbox, Split, Rotate).
   - Click **Start Conversion**.

## Requirements

- Python 3.10+ (for source)
- Windows 10/11 (for exe)
- WinRAR (Optional, for RAR/CBR support)

## License

MIT License
