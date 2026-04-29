# IT3040 ITPM - Assignment 1 Option 1
## Singlish to Sinhala Transliteration Testing

This repository contains automated test scripts for testing the Chat Sinhala transliteration function at [pixelssuite.com/chat-translator](https://www.pixelssuite.com/chat-translator).

## Prerequisites
- Python 3.11 or higher
- Google Chrome (recommended)

## Installation

```bash
pip3 install -U pip
pip3 install playwright openpyxl
playwright install
```

## Running the Tests

```bash
python3 test_automation.py --excel "Assignment 1 - Test cases.xlsx" --url "https://www.pixelssuite.com/chat-translator" --wait-ms 5000 --type-delay-ms 80 --slow-mo-ms 200 --save-every 1 --keep-open
```

## Results
- Total test cases: 50
- All test cases cover failures in Singlish to Sinhala transliteration
- Results are saved automatically in the Excel file
