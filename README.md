# PowerPoint Presentation Creator

This project provides two Python scripts for automatically creating PowerPoint presentations from images and music files. It's designed to make it easy to create professional-looking presentations with minimal effort.

## Features

### 1. Basic Image Presentation (`presetaion_only_with_images.py`)
- Creates a PowerPoint presentation from images in a directory
- Supports multiple image formats (PNG, JPG, JPEG, GIF, BMP, TIFF)
- Automatically centers and resizes images to fit slides while maintaining aspect ratio
- Organizes images from subdirectories with title slides
- Maintains alphabetical order of images and folders

### 2. Enhanced Presentation with Music (`presetaion_with_images_and_music.py`)
- All features from the basic version
- Adds background music support (MP3, WAV, WMA)
- Automatic slide transitions
- 16:9 widescreen format
- Better error handling and logging
- Improved title slide formatting

## Requirements

- Python 3.x
- Required Python packages:
  - python-pptx
  - Pillow (for image processing)

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/pdf-presetaion-with-images-and-music.git
cd pdf-presetaion-with-images-and-music
```

2. Install required packages:
```bash
pip install python-pptx Pillow
```

## Directory Structure

Create the following directory structure in your project folder:

```
project_folder/
├── images/              # Main directory for images
│   ├── image1.jpg      # Images in root directory
│   ├── image2.png
│   └── subfolder1/     # Subdirectories for organizing images
│       ├── image3.jpg
│       └── image4.png
├── music/              # Directory for background music (optional)
│   └── background.mp3
└── *.py               # Python scripts
```

## Usage

### Basic Image Presentation

1. Place your images in the `images` directory
2. Run the script:
```bash
python presetaion_only_with_images.py
```

### Enhanced Presentation with Music

1. Place your images in the `images` directory
2. (Optional) Place your background music in the `music` directory
3. Run the script:
```bash
python presetaion_with_images_and_music.py
```

## Output

- The script will create a PowerPoint presentation file named `presentation.pptx`
- Images will be automatically sized and centered on slides
- Subdirectories will be represented with title slides
- If music is provided, it will be embedded in the first slide

## Notes

- For the music-enhanced version, you may need to manually configure some settings in PowerPoint:
  - Set "Play Across Slides" for the audio
  - Configure "Loop Until Stopped" if desired
- The script maintains image aspect ratios while maximizing size
- All images and folders are processed in alphabetical order
- Supported image formats: PNG, JPG, JPEG, GIF, BMP, TIFF
- Supported audio formats: MP3, WAV, WMA

## Troubleshooting

1. If the presentation fails to open:
   - Try changing the output filename to `.pptx` instead of `.ppsx`
   - Check if the file is not open in another program
   - Verify you have write permissions in the directory

2. If images aren't showing up:
   - Verify the image files are in supported formats
   - Check if the images directory exists and contains files
   - Ensure you have read permissions for the image files

3. If music isn't playing:
   - Verify the music file is in a supported format
   - Check if the music file is properly placed in the music directory
   - Try manually configuring the audio settings in PowerPoint

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.