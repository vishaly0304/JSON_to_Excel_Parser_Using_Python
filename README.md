# JSON to Excel Parser: Exam Results

## Overview
This project parses a JSON file containing exam results from different states, colleges, and standards, then writes the data into a well-structured Excel workbook. The Python code is designed to dynamically extract unique subjects and organize the data for easy readability in the Excel sheet.

## Features
- Parses JSON data with a hierarchical structure (states -> colleges -> standards -> students -> subjects).
- Dynamically collects all unique subjects across the dataset for comprehensive representation.
- Styles the Excel file with headings, borders, and background colors for better visualization.
- Saves the parsed data into an Excel file named `parsed_data.xlsx`.

## Prerequisites
- Python 3.x
- `pandas` library (Install using `pip install pandas`)
- `openpyxl` library (Install using `pip install openpyxl`)

## Instructions
1. Place the JSON data file in the same directory as the Python script and name it `data.json`.
2. Ensure the JSON file follows the structure detailed below:
    ```json
    {
      "states": [
        {
          "name": "State Name",
          "colleges": [
            {
              "name": "College Name",
              "standards": [
                {
                  "standard": "Grade Level",
                  "students": [
                    {
                      "name": "Student Name",
                      "subjects": {
                        "Subject1": Score,
                        "Subject2": Score
                      }
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    }
    ```
3. Run the Python script by executing the following command in your terminal:
    ```
    python JSON_to_Excel_Parser_Using_Python.py
    ```
4. Check the directory for the generated `parsed_data.xlsx` file.

## Output Structure in Excel
- **Heading**: Displays the title "Exam Results for Different States" styled with bold text, background color, and centered alignment.
- **State and College Sections**: Each state's data is grouped, with a unique background color applied to the state name.
- **Subjects and Marks**: Dynamically lists all subjects per student, even if some students do not have scores for certain subjects.
- **Center Alignment**: All subject marks are center-aligned for better readability.

## Error Handling
- If the JSON file is missing or contains errors, the script gracefully exits with a relevant error message.
- The script handles missing marks for subjects by leaving the respective cells empty.
