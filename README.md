# MP3 Metadata Extractor

This Python script extracts ID3 metadata from MP3 files and saves it to:

* A CSV file.
* An Excel spreadsheet with separate sheets for each album and an "All Songs" sheet.

## Features

* Extracts artist, title, duration, album, track number, genre, filename, and comments.
* Formats the duration in MM:SS format.
* Sorts the "All Songs" sheet by album and track number.
* Sorts individual album sheets by track number.
* Applies formatting to the Excel spreadsheet (headers, gridlines, column widths).

## Requirements

* Python 3
* pandas
* tinytag
* openpyxl

## How to use

1.  **Install the required libraries:**

    ```bash
    pip install -r requirements.txt
    ```

2.  **Update the script:**

    *   Open the `mp3-metadata-extractor.py` script in a text editor.
    *   Update the `mp3_folder`, `csv_output_path`, and `excel_output_path` variables with the paths to your music folder and where you want to save the output files.

3.  **Run the script:**

    ```bash
    python mp3-metadata-extractor.py
    ```

## Contributing

If you'd like to contribute to this project, please follow these steps:

1.  Fork the repository.
2.  Create a new branch for your changes.
3.  Make your changes and commit them.
4.  Push your changes to your fork.
5.  Submit a pull request.

## Project Structure

* `mp3-metadata-extractor.py`: The main Python script.
* `README.md`: This file.
* `requirements.txt`: Lists the required Python libraries.
=======
# mp3_metadata_extractor
Python script to extract metadata from a directory of MP3s, sort it by album and track, and organize it into an Excel spreadsheet.