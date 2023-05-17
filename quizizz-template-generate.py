import pandas as pd

# Path and sheet settings
path_data_source = 'E:\Code\quizizz-template-generate\Source.xlsx'  # Specify the path to the source Excel file
sheet_data_source = "Sheet1"  # Specify the name of the sheet in the source Excel file
usecols_data_source = "A:H"  # Specify the columns to be read from the source Excel file
result_path = 'E:\Code\quizizz-template-generate\Result.xlsx'  # Specify the path for the result Excel file
sheet_result = 'Sheet1'  # Specify the name of the sheet in the result Excel file
default_question_type = "Multiple Choice"
default_duration = 30   # Specify the default duration

# Read the data from the source Excel file
data = pd.read_excel(path_data_source, sheet_name = sheet_data_source, usecols = usecols_data_source)

# Create a writer to save the result to an Excel file
writer = pd.ExcelWriter(result_path, engine = 'xlsxwriter')

# Iterate over each row in the data
for index, row in data.iterrows():
    # Check if the "Question Text" does not start with "A.", "B.", "C.", or "D."
    if row["Question Text"][:2] not in ["A.", "B.", "C.", "D."]:
        # Save the index of question
        index_question = index
        # Set the question type as "Multiple Choice"
        data.at[index_question, "Question Type"] = default_question_type
        # Set the time in seconds for the question
        data.at[index_question, "Time in seconds"] = default_duration
    else:
        # Iterate over each column in the row not containing the question
        for col_name, col_value in row.iteritems():
            # Check the options starting with "A.", "B.", "C.", or "D."
            if str(col_value).startswith("A."):
                # Set the option 1 value of the question row based on index_question
                data.at[index_question, "Option 1"] = col_value.replace("A.", "").lstrip()
            elif str(col_value).startswith("B."):
                # Set the option 2 value of the question row based on index_question
                data.at[index_question, "Option 2"] = col_value.replace("B.", "").lstrip()
            elif str(col_value).startswith("C."):
                # Set the option 3 value of the question row based on index_question
                data.at[index_question, "Option 3"] = col_value.replace("C.", "").lstrip()
            elif str(col_value).startswith("D."):
                # Set the option 4 value of the question row based on index_question
                data.at[index_question, "Option 4"] = col_value.replace("D.", "").lstrip()
        # Drop the row after processing the options
        data = data.drop(index = index)

# Write the modified data to the result Excel file
data.to_excel(writer, sheet_name=sheet_result, index=False)
writer.save()
