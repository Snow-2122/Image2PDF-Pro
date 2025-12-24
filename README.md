# 📄 Image2PDF Pro

**The Professional, Lossless Image to PDF Converter for Windows.**

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge&logo=python)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

**Image2PDF Pro** is a robust desktop application designed for comic book readers, archivists, and power users. Unlike basic online converters, it ensures **100% pixel fidelity**, handles complex archive formats (ZIP/RAR/CBZ/CBR) natively, and offers a sleek, dark-mode GUI.

---

## ✨ Key Features

### 🚀 Professional Performance

- **Lossless Engine**: Uses `Quality=100` and `Subsampling=0` to wrap images into PDF without degrading visual quality.
- **Parallel Processing**: Multi-threaded engine utilizes all CPU cores for blazing fast conversions.
- **Smart Queue**: Drag-and-drop folders, subfolders, or archives to batch process them instantly.

### 🖼️ Intelligent Image Handling

- **Landscape Modes**:
  - **None**: Keeps original aspect ratio (Mixed orientation).
  - **Split**: Automatically cuts wide double-page spreads into two single pages (Perfect for Manga/Comics).
  - **Rotate**: Rotates wide images 90° to fit standard screens.
  - **Letterbox**: Adds white bars to maintain a consistent page size.
- **Format Support**: JPG, PNG, WEBP, BMP, JPEG.

### 🧠 Smart System Integration

- **Auto-Detect WinRAR**: The app automatically detects if WinRAR is installed and uses its `UnRAR.exe` engine to handle `.rar` and `.cbr` files without complex setup.
- **Path Locking**: Fully portable. Configuration and logs are locked to the application folder, preventing file clutter on your system.
- **Safety First**: "Delete Source" is opt-in only. Your original files are never touched unless you explicitly enable deletions.

---

## 📥 Installation

### Option 1: Standalone Executable (Recommended)

No Python installed? No problem.

1. Go to the [**Releases**](https://github.com/Snow-2122/Image2PDF-Pro/releases) page.
2. Download `Image2PDF_Pro.exe`.
3. Move the file to a permanent folder (e.g., `Documents/Image2PDF Pro`).
4. **Run it!**
   - _Note: If you want RAR support, the app will guide you to copy `UnRAR.exe` to the same folder._

### Option 2: Run from Source

If you are a developer or want to modify the code:

1. **Clone the Repository**:

   ```bash
   git clone https://github.com/Snow-2122/Image2PDF-Pro.git
   cd Image2PDF-Pro
   ```

2. **Install Dependencies**:

   ```bash
   pip install -r requirements.txt
   ```

3. **Run the App**:
   ```bash
   python main.py
   ```

---

## 🛠️ Usage Guide

### 1. GUI Mode (Dashboard)

Double-click the application to open the modern Dark Mode dashboard.

- **Source**: Select a folder containing images or archives.
- **Output**: (Optional) Choose where to save the PDFs. Default is the source folder.
- **Settings**:
  - **Quality**: Keep at **100** for lossless. Lower it to reduce file size.
  - **Threads**: Number of CPU cores to use.
  - **Landscape**: Choose how to handle wide images.
- **Start**: Click to begin. The progress bar will show real-time status.

### 2. Drag & Drop (Batch Mode)

Simply **drag a folder** (or multiple folders) directly onto the `.exe` file or the `main.py` script.

- The app will instantly start converting using your **last saved settings**.
- It respects your decision on "Delete Source" and "Landscape Mode".

---

## ⚙️ Configuration (Advanced)

The application creates a `config.ini` file in its root directory. You can edit this manually if needed:

```ini
[General]
pdf_quality = 100
parallel_processing = True
thread_count = 8
enable_gui = True
enable_rar = True
theme = solar
landscape_mode = none
delete_source = False
```

---

## 📦 Building from Source (Create .exe)

Want to build your own executable? We use **PyInstaller**.

1. **Install PyInstaller**:
   ```bash
   pip install pyinstaller
   ```
2. **Run the Build Command**:
   (This specific command ensures all theme files are collected)
   ```bash
   python -m PyInstaller --noconsole --onefile --icon=icon.ico --name="Image2PDF_Pro" --hidden-import=ttkbootstrap --collect-all=ttkbootstrap --hidden-import=plyer --collect-all=plyer main.py
   ```
3. The new executable will appear in the `dist/` folder.

---

## 📋 Requirements

- **OS**: Windows 10 / 11
- **Python**: 3.10+ (Only for Source users)
- **WinRAR**: Installed (Optional, for .cbr support)

## 📄 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

---

_Built with ❤️ using Python, Pillow, and ttkbootstrap._
