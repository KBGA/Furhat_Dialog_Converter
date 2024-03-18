"""
Furhat Dialog Converter

This script converts dialog data from Furhat robots, stored in JSON format, into PDF and Excel files.
The purpose is to facilitate the analysis and archival of conversational interactions with Furhat robots.

Usage:
- Ensure that the 'DejaVuSans.ttf' and 'DejaVuSans-Bold.ttf' fonts are in your working directory or specify their path.
- Adjust the 'main_directory' to point to your directory containing the dialog JSON files.
- Run the script to generate PDF and Excel files in the same directory as your JSON files.

Requirements:
- Python 3.x
- openpyxl
- reportlab

Author: Georges A. K. Bonga
Date: 2024-03-18
"""

import os
import json
from openpyxl import Workbook
from openpyxl.styles import Alignment
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color, black
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Register fonts for PDF generation
pdfmetrics.registerFont(TTFont('DejaVuSans', 'DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', 'DejaVuSans-Bold.ttf'))

# Color palette for up to 10 participants (RGB)
color_palette = [
    colors.blue,  # Blau
    colors.magenta,  # Magenta
    colors.orange,  # Orange
    colors.green,  # Grün
    colors.purple,  # Lila
    Color(165/255, 42/255, 42/255),  # Braun
    colors.pink,  # Rosa
    Color(64/255, 224/255, 208/255),  # Türkis
    Color(0, 0, 128/255),  # Dunkelblau
    Color(128/255, 128/255, 0)  # Oliv
]


def convert_json_to_pdf_and_excel(json_file_path, pdf_output_path, excel_output_path):
    """
    Converts dialog data from a JSON file to PDF and Excel formats.

    Parameters:
    - json_file_path: The file path to the source JSON file.
    - pdf_output_path: The file path to output the PDF file.
    - excel_output_path: The file path to output the Excel file.
    """

    with open(json_file_path, 'r', encoding='utf-8') as inputFile:
        json_text = inputFile.read()

    json_text = json_text.strip().strip(',')

    try:
        data = json.loads(f'[{json_text}]')
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON file: {e}")
        return

    if data and len(data) > 1:
        session_name = data[0].get('sessionName', 'Unknown Title')
        session_start_time = data[0].get('startTime', 'Unknown').split('.')[0]
        session_end_time = data[-1].get('endTime', 'Unknown').split('.')[0]

        # Initialize PDF document
        c = canvas.Canvas(pdf_output_path, pagesize=A4)
        c.setFont("DejaVuSans-Bold", 16)
        c.setFillColor(black)
        c.drawCentredString(A4[0] / 2, 800, f"Titel: {session_name}")
        c.setFont("DejaVuSans", 12)
        c.drawString(50, 780, f"Session Start Time: {session_start_time}")
        c.drawString(50, 765, f"Session End Time: {session_end_time}")

        # Insert empty line
        y = 745 - 20  # 20 points for the height of a line

        participants = {"Robot": color_palette[0]}
        next_color = 1
        last_participant = None

        x = 50
        right_margin = 25
        page_width = A4[0]
        max_width = page_width - right_margin

        # Create a new Excel workbook and select the active worksheet
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = 'Dialog Data'

        # Write title, start and end time in the first three lines and merge the first two cells
        for row, value in enumerate([f'Titel: {session_name}', f'Session Start Time: {session_start_time}',
                                     f'Session End Time: {session_end_time}'], 1):
            worksheet.append([value, ''])  # Zwei Zellen für den Merge
            worksheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            for cell in worksheet[row]:
                cell.alignment = Alignment(horizontal='center')

        # Empty line after the information
        worksheet.append([])

        # Write the column headers for the dialog data
        worksheet.append(['Zeit', 'Benutzer', 'Text'])

        for entry in data[1:]:
            if 'startTime' in entry and 'type' in entry:
                if y < 50:
                    c.showPage()
                    c.setFont("DejaVuSans", 12)
                    x = 50
                    y = 800

                timestamp = entry['startTime'].split('.')[0]
                entry_type = entry['type']

                if entry_type == 'robot.speech':
                    participant = "Robot"
                elif 'user' in entry and (entry_type == 'user.speech'):
                    participant = entry['user']
                    last_participant = participant
                elif entry_type == 'user.response':
                    participant = last_participant if last_participant else "Unknown"
                    continue
                else:
                    participant = "Unknown"

                message = entry.get('text', '')
                # Remove <mark name="action"/> at the end of the text
                message = message.replace('<mark name="action"/>', '')

                if participant and participant not in participants:
                    participants[participant] = color_palette[next_color]
                    next_color = (next_color + 1) % len(color_palette)

                c.setFont("DejaVuSans", 12)
                color = participants.get(participant, black)
                c.setFillColor(color)
                speaker_info = f"{timestamp} {participant}: "
                c.drawString(x, y, speaker_info)
                current_width = c.stringWidth(speaker_info, "DejaVuSans", 12)
                x += current_width

                c.setFont("DejaVuSans", 12)
                c.setFillColor(black)
                words = message.split()
                for word in words:
                    word_width = c.stringWidth(word + ' ', "DejaVuSans", 12)
                    if x + word_width > max_width:
                        x = 50
                        y -= 20
                        if y < 50:
                            c.showPage()
                            c.setFont("DejaVuSans", 12)
                            y = 800
                    c.drawString(x, y, word + ' ')
                    x += word_width
                x = 50
                y -= 20

                # Insert the data into the Excel table
                worksheet.append([timestamp, participant, message.replace('\n', ' ')])

        c.save()
        print(f"Die PDF-Datei wurde erfolgreich unter {pdf_output_path} gespeichert.")

        # Save the Excel workbook
        workbook.save(filename=excel_output_path)
        print(f"Die Excel-Tabelle wurde erfolgreich unter {excel_output_path} gespeichert.")
    else:
        print(f"Die JSON-Datei in {json_file_path} scheint leer oder nicht ausreichend zu sein.")


# Main directory
main_directory = 'C:/Users/admin/.furhat/logs'

# Go through all subfolders in the main directory
for root, dirs, files in os.walk(main_directory):
    for file in files:
        if file == 'dialog.json':
            jsonFilePath = os.path.join(root, file)
            pdfOutputPath = os.path.join(root, 'dialog.pdf')
            excelOutputPath = os.path.join(root, 'dialog.xlsx')
            convert_json_to_pdf_and_excel(jsonFilePath, pdfOutputPath, excelOutputPath)
