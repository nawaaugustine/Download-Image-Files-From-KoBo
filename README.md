# KoBO Image Downloader

## Description

This project contains a Python script that downloads an Excel file from a KoBo project, extracts URLs of images from specific columns, and downloads those images into folders named after the Excel tabs. The images are then zipped into respective folders.

## Prerequisites

- Python 3.x
- Requests library
- Openpyxl library

To install the required libraries, run:

```bash
pip install requests openpyxl
```

## Configuration

Create a config.json file in the project root directory with the following content:

```json
{
    "api_key": "<YOUR_API_KEY>",
    "project": "LINK TO YOUR KOBO DOWNLOAD SETTING PROJECT"
}
```

Replace the above with your actual details.

## Usage

You can run the script using the command:

```bash
python main.py
```

## How It Works

1. The script downloads an Excel file from KoBo.
2. It reads the Excel file, looking for columns with "_URL" in the title, and extracts the URLs.
3. The images from the URLs are downloaded into folders named after the Excel tabs.
4. The folders with images are zipped.

## Contributing

Feel free to fork the repository and submit pull requests for any improvements or feature additions.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
