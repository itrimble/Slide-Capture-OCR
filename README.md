# Enhanced Universal Slide Capture & OCR Renamer

![Banner showing slide capture functionality](https://example.com/banner-image.png)

[![Version](https://img.shields.io/badge/Version-2.2.0-brightgreen.svg)](https://github.com/yourusername/slide-capture-ocr)
[![Release Date](https://img.shields.io/badge/Released-May%206%2C%202025-blue.svg)](https://github.com/yourusername/slide-capture-ocr/releases)
[![Platform](https://img.shields.io/badge/Platform-macOS%2011%2B-orange.svg)](https://github.com/yourusername/slide-capture-ocr)
[![License](https://img.shields.io/badge/License-MIT-lightgrey.svg)](https://github.com/yourusername/slide-capture-ocr/blob/main/LICENSE)

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
  - [Capture Mode](#capture-mode)
  - [Single Screenshot Mode](#single-screenshot-mode)
  - [Settings Mode](#settings-mode)
- [Keyboard Shortcuts](#keyboard-shortcuts)
- [Slide Types Recognized](#slide-types-recognized)
- [Adaptive Screen Resolution Support](#adaptive-screen-resolution-support)
- [Configuration Options](#configuration-options)
- [Troubleshooting](#troubleshooting)
- [Version History](#version-history)
- [License](#license)

## Overview

**Enhanced Universal Slide Capture & OCR Renamer** is a powerful macOS AppleScript utility designed to automate the process of capturing presentation slides and intelligently renaming them based on content analysis. The tool is particularly useful for educators, trainers, and students who need to save web-based presentations (like Google Slides, PowerPoint Online, etc.) with meaningful filenames that reflect the slide content.

This script uses Optical Character Recognition (OCR) to extract text from captured slides, analyzes the content to determine slide type and subject, and applies a consistent naming convention that makes organizing and finding slides easier.

## Features

- **Multi-browser support**: Works with Chrome, Safari, Firefox, and Microsoft Edge
- **Intelligent slide classification**: Automatically detects different slide types
- **Adaptive geometry scaling**: Adjusts to different screen resolutions
- **Keyboard shortcuts**: Pause, resume, or cancel capture with simple key combinations
- **Resume capability**: Pick up where you left off if capture is interrupted
- **Progress tracking**: Shows progress bar with time estimates
- **Module context awareness**: Tracks presentation modules for better naming
- **Batch processing**: Capture and process entire presentations
- **Single file processing**: Rename individual screenshots
- **Configurable settings**: Adjust timing, file paths, and OCR parameters
- **Comprehensive logging**: Track all operations with customizable log levels

## Requirements

- **macOS 11 (Big Sur)** or newer
- **Tesseract OCR** for text extraction
- **ImageMagick** for image processing

## Installation

### 1. Install Required Dependencies

If you don't have Homebrew installed, install it first:

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

Then install the required dependencies:

```bash
brew install tesseract
brew install imagemagick
```

### 2. Install the Script

1. Download the script from [GitHub](https://github.com/yourusername/slide-capture-ocr/releases/latest)
2. Save it to a location of your choice (e.g., `~/Applications/Scripts/`)
3. Make it executable:
   ```bash
   chmod +x ~/Applications/Scripts/slide_capture.scpt
   ```

### 3. Verify Installation

Run the script and select "Help & Documentation" from the menu. If the help screen appears, the script is installed correctly.

## Usage

The script offers three main operating modes:

### Capture Mode

This mode automates the capture of multiple slides from a presentation:

1. Open your presentation in a supported browser (Chrome, Safari, Firefox, Edge)
2. Run the script and select "Capture Slides from Presentation"
3. Select the browser containing your presentation (or confirm the auto-detected one)
4. Enter the number of slides to capture
5. Select an output folder
6. Let the script run - it will:
   - Enter full-screen presentation mode
   - Capture each slide
   - Extract text via OCR
   - Intelligently name and save each slide
   - Advance to the next slide automatically

### Single Screenshot Mode

Process and rename an existing screenshot:

1. Run the script and select "Rename Single Screenshot"
2. Select a PNG or JPG file
3. The script will analyze the screenshot, extract the title, and rename it according to the detected slide type

### Settings Mode

Configure script options:

1. Run the script and select "Settings & Configuration"
2. Adjust parameters like:
   - Log level
   - Delay between slides
   - Capture delay
   - Maximum title length
3. Save your settings or reset to defaults

## Keyboard Shortcuts

During capture mode, you can use these keyboard shortcuts:

- **Option+P**: Pause capture
- **Option+R**: Resume capture
- **Option+C**: Cancel capture

## Slide Types Recognized

The script can identify and specially name these slide types:

| Slide Type | Detection Method | Example Filename |
|------------|------------------|------------------|
| Regular Content | Default | `01_Module_Title_Content.png` |
| Cyber Lab | "Cyber Lab" in content | `02_Cyber_Lab_Configuring_Firewall.png` |
| Knowledge Check | "Knowledge Check" in content | `03_Knowledge_Check_1.png` |
| Knowledge Check Answer | "Knowledge Check Answer" | `04_Knowledge_Check_Answer_1.png` |
| Pulse Check | "Pulse Check" in content | `05_Pulse_Check_Module.png` |
| Summary | "Summary" in content | `06_Summary_Module.png` |
| Break | "Break" in content | `07_Break.png` |
| Real World Scenario | "Real World Scenario" in sidebar | `08_Real_World_Scenario_Module.png` |

## Adaptive Screen Resolution Support

The script automatically detects your screen resolution and scales the capture regions accordingly. This ensures optimal performance on various display configurations:

- Standard resolution displays (1920×1080)
- Retina displays
- 4K/5K displays
- Ultrawide monitors

The scaling algorithm uses a base resolution of 3840×2160 to calculate proportions for different screen sizes.

## Configuration Options

The script stores configuration in a plist file at `~/Library/Preferences/com.slidecapture.config.plist`.

Configurable parameters include:

| Parameter | Default | Description |
|-----------|---------|-------------|
| logLevel | 1 | 0=Debug, 1=Info, 2=Warning, 3=Error |
| delayBetweenSlides | 1.5 | Seconds to wait between slide advances |
| captureDelay | 0.5 | Seconds to wait before capturing screenshot |
| maxTitleLength | 60 | Maximum length for slide titles |
| defaultNumSlides | 20 | Default number of slides when starting new capture |
| defaultOutputFolder | ~/Downloads/Slides | Default output location |

## Troubleshooting

### Common Issues

#### OCR Quality Problems

**Symptoms**: Slides are named "Untitled" or text is incorrectly extracted.

**Solutions**:
- Ensure your screen brightness is adequate
- Use presentations with good contrast
- Try adjusting capture delay to ensure slides are fully loaded

#### Slide Advance Issues

**Symptoms**: Script fails to advance slides or captures the same slide multiple times.

**Solutions**:
- Increase the delay between slides
- Ensure the presentation is in full-screen mode
- Check if keyboard focus is on the browser window

#### Dependency Issues

**Symptoms**: "Required tools not found" error message.

**Solutions**:
- Verify Tesseract and ImageMagick are installed: `which tesseract` and `which convert`
- Try reinstalling the dependencies: `brew reinstall tesseract imagemagick`

### Logs

Check the log file for detailed troubleshooting information:
```
~/Library/Logs/slide_capture.log
```

Increase log verbosity by setting logLevel to 0 (Debug) in Settings.

## Version History

### v2.2.0 (2025-05-06)
- Implemented functional keyboard handlers using Carbon framework
- Fixed configuration loading/saving functionality
- Enhanced error recovery with specific strategies for different error types
- Added Settings & Configuration mode
- Added image capture verification
- Improved slide advance recovery
- Added comprehensive inline documentation

### v2.1.0 (2025-04-15)
- Added multi-browser support
- Implemented resume capability
- Added progress tracking with time estimation
- Enhanced OCR processing with multiple regions

### v1.0.0 (2025-03-01)
- Initial release with basic slide capture and OCR

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

© 2025 Your Organization. All rights reserved.