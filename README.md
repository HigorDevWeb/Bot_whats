WhatsApp Bot README
Description

This Python script automates the process of sending WhatsApp messages to clients listed in an Excel spreadsheet. It opens WhatsApp Web, extracts client information from the Excel file, and sends personalized messages regarding invoice due dates.
Features

    Loads client data from an Excel spreadsheet.
    Sends personalized messages to clients via WhatsApp Web.
    Uses pyautogui to automate clicking the send button.
    Handles exceptions and logs errors to a CSV file.

Requirements

    Python 3.x
    openpyxl for handling Excel files
    pyautogui for screen automation
    An image of the send button saved as image.png in the bot_whats directory

Installation

    Clone the repository or download the script.
    Ensure you have Python 3.x installed.
    Install the required Python libraries:

    sh

    pip install openpyxl pyautogui

    Save an image of the WhatsApp send button as image.png in a folder named bot_whats.

Usage

    Open WhatsApp Web and log in.
    Ensure the Excel file (clientes.xlsx) is in the bot_whats directory with the following columns: Name, Phone, Due Date.
    Run the script:

    sh

    python app.py

    The script will:
        Open WhatsApp Web.
        Load client data from the Excel file.
        Send personalized messages to each client regarding their invoice due date.
        Log any errors encountered during the process to erros.csv.

File Structure

bash

/path-to-your-project/
│
├── bot_whats/
│   ├── clientes.xlsx  # Excel file with client data
│   ├── image.png      # Image of the WhatsApp send button
│   ├── erros.csv      # Log file for errors
│
└── app.py             # Python script

Error Handling

    The script waits for 30 seconds to load WhatsApp Web initially.
    It logs any errors encountered while sending messages to erros.csv in the format: Name,Phone,Due Date.

Note

    Ensure the Excel file and the image file are correctly placed in the bot_whats directory.
    Adjust the sleep times (time.sleep()) if necessary to match your internet speed and WhatsApp Web loading times.
