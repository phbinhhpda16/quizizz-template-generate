# Quizizz-Template-Generate

This repository contains a Python script that automates the conversion of text copied from Word into an Excel template format for Quizizz questions. The script takes input from the "Source.xlsx" file, starting from cell A2, and generates the converted output in the "Result.xlsx" file.

## Folder Structure

The repository has the following folder structure:

```
- Source.xlsx         # Excel file for users to input text copied from Word (starting from A2)
- Result.xlsx         # Excel file containing the converted output after running the Python script
- quizizz-template-generate.py    # Python script for automating the conversion process
- run_quizizz.bat    # Batch script to run the Python file without opening a code editor
- README.md           # This file, providing an overview of the project
```

## Usage

To use the Quizizz-Template-Generate, follow these steps:

1. Update the Python Script: Open the `quizizz-template-generate.py` file and modify the variables within the script to suit your requirements. Adjust any necessary settings or formatting options as needed.

2. Run the Python Script: Execute the `quizizz-template-generate.py` file to initiate the conversion process. This will read the text from `Source.xlsx`, convert it into the Quizizz question template format, and save the result in the `Result.xlsx` file.

3. Review the Output: Open the `Result.xlsx` file to see the converted output. Verify that the questions are correctly formatted and ready to be imported into Quizizz.

4. Customize the Excel Templates: If desired, modify the `Source.xlsx` and `Result.xlsx` files to customize the input and output templates according to your specific needs. Make sure to adjust the Python script accordingly to accommodate any changes made to the template structure.

5. Batch Script (Optional): If you prefer to run the Python script without opening a code editor, you can use the `run_quizizz.bat` file. Open the file in a text editor, and modify the paths to the `Python_script.py` and the Python interpreter (anaconda or standard Python) as per your system configuration. Then, run the `run_quizizz.bat` file to execute the Python script.

Please ensure that you have the necessary permissions to read from and write to the files and folders involved in the conversion process.

## License

This project is licensed under the [MIT License](LICENSE), allowing you to modify and distribute the code as per the terms stated.

Feel free to explore and adapt the Word-to-Quizizz-Excel-Converter to meet your specific needs. If you encounter any issues or have suggestions for improvement, please feel free to contribute to the repository or open an issue.
