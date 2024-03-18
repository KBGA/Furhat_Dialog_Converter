
# Furhat Dialog Converter

## Description

The Furhat Dialog Converter is an open-source Python utility designed to transform dialog data from Furhat robots, stored in JSON format, into more accessible PDF and Excel files. This tool is ideal for researchers, developers, and users looking to analyze or archive conversational interactions with Furhat robots.

## Features

- Reads dialog data from JSON files.
- Generates a comprehensive PDF report with customizable fonts and colors for participants.
- Creates an Excel workbook with detailed dialog entries for easy analysis.
- Supports multiple dialog sessions and participants.
- Automatic detection and processing of dialog files within a specified directory.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/KBGA/Furhat_Dialog_Converter.git
   ```

2. Navigate to the project directory:
   ```bash
   cd Furhat_Dialog_Converter
   ```

3. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

To convert dialog data from a specified directory containing `dialog.json` files, follow these steps:

1. Ensure the `main_directory` variable in the script points to your directory with the dialog JSON files.
2. Run the script:
   ```bash
   python furhat_dialog_converter.py
   ```
3. The script will process each `dialog.json` file, converting it into corresponding PDF and Excel files in the same directory.

## Contributing

Contributions to the Furhat Dialog Converter are welcome. If you have suggestions for improvement or encounter any issues, please feel free to fork the repository, make your changes, and submit a pull request for review.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contact

For further information, support, or inquiries, please [click here to send a mail](mailto:kouayim@kouayim.com).

