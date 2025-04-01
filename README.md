# Auto Watermarker V2.5 (ImageMagick Path Config)

This Python application automates the process of adding watermarks to images. It supports batch processing for multiple image files and folders, with options for customizing the watermark placement and output.

## Features

* **Automatic Watermarking:** Adds watermarks to images at specified intervals.
* **Batch Processing:** Processes multiple image files and folders.
* **Configurable Watermark Placement:** Adjust watermark frequency, search step, uniformity threshold, and maximum search steps.
* **Image Format Support:**
    * Processes PNG and JPEG files directly.
    * Supports PSD and PSB files via ImageMagick (requires ImageMagick installation).
* **Automatic Uniform Area Detection:** Detects and places watermarks in uniform areas of images.
* **Output Options:**
    * Saves processed images in a specified output directory.
    * Option to create ZIP archives for each processed folder.
* **GUI Interface:** User-friendly interface for configuration and processing.
* **Configuration Saving:** Saves and loads settings from a JSON file.
* **ImageMagick Path Configuration:** Allows users to specify the path to the ImageMagick executable.
* **Detailed Logging:** Provides detailed processing logs in the application's text area.

## Requirements

* Python 3.x
* `tkinter`
* `customtkinter`
* `Pillow (PIL)`
* `ImageMagick` (for PSD/PSB support)

## Installation

1.  **Clone the repository:**

    ```bash
    git clone [repository_url]
    cd Auto-Watermarker
    ```

2.  **Install dependencies:**

    ```bash
    pip install customtkinter pillow
    ```

3.  **(Optional) Install ImageMagick:**
    * For PSD and PSB support, install ImageMagick and configure the path in the application.

4.  **Run the application:**

    ```bash
    python AutoWatermarker.py
    ```

## Usage

1.  **Select the main folder** containing the images.
2.  **Select the watermark file** (PNG, JPG, or JPEG).
3.  **Configure the settings:**
    * Frequency: Interval between watermarks.
    * Search Step: Step size for searching watermark placement.
    * Uniformity Threshold: Threshold for detecting uniform areas.
    * Max Steps: Maximum search steps.
    * Create ZIP: Option to create ZIP archives.
    * ImageMagick path: path to magick.exe if you want to process psd and psb files.
4.  **Click "Start Processing"** to begin the watermarking process.
5.  **Monitor the progress** in the status log and progress bar.

## Configuration file

The program saves all the settings in a json file located in the user home directory, inside of .config/AutoWatermarker or AutoWatermarker, the name of the file is .AutoWatermarkerConfig.json.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## License

[MIT License]

## Author

Vlad


