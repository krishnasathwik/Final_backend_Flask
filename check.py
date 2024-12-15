import os
from flask import Flask, request, jsonify
import pandas as pd

# Initialize the Flask app
app = Flask(__name__)

# Folder to store uploaded files
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the upload folder exists (if not, it will be created)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Route to handle file uploads
@app.route('/upload', methods=['POST'])
def upload_file():
    # Check if the 'file' part is in the incoming request
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400  # Return error if no file is uploaded
    
    file = request.files['file']  # Get the file from the request
    
    # Check if the file has a filename (i.e., it's not an empty file upload)
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400  # Return error if file is empty
    
    # Save the uploaded file to the specified upload folder
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    # Call the validate_file function to validate the contents of the uploaded file
    result = validate_file(file_path)
    return jsonify({"message": result})  # Return the result of the validation


# Function to validate the uploaded Excel file
def validate_file(file_path):
    try:
        # Load the Excel file using pandas ExcelFile
        df = pd.ExcelFile(file_path)
        
        # Define the required sheet names
        required_sheets = ['Course', 'Topic', 'Resource', 'Learner']
        sheet_names = df.sheet_names  # Get the list of sheet names in the uploaded Excel file
        
        # Check if the number of sheets is exactly 4
        if len(sheet_names) != 4:
            return "Invalid file format: The file must contain exactly 4 sheets."
        
        # Check if all required sheets are present in the file
        for sheet_name in required_sheets:
            if sheet_name not in sheet_names:
                return f"Invalid file format: Missing required sheet '{sheet_name}'."

        # Define field validations for each sheet
        field_validations = {
            'Course': ['Course ID', 'Course Name'],
            'Topic': ['Topic ID', 'Topic Name', 'Description'],
            'Resource': ['Resource ID', 'Resource Name', 'Resource Content', 'Module ID', 'Module Name', 'Sub Module ID'],
            'Learner': ['Learner ID', 'Name', 'Essay', 'Module ID', 'Submodule ID']
        }

        # Loop through each required sheet and validate its fields
        for sheet_name in required_sheets:
            sheet_df = df.parse(sheet_name)  # Read the sheet into a pandas DataFrame
            
            # Check if all required fields exist in the sheet (by checking column names)
            missing_fields = [field for field in field_validations[sheet_name] if field not in sheet_df.columns]
            
            # If any fields are missing, return an error message
            if missing_fields:
                return f"Invalid file format: The sheet '{sheet_name}' is missing the following fields: {', '.join(missing_fields)}"
            
            # Check if the sheet contains any data (i.e., it should not be empty)
            if sheet_df.empty:
                return f"Invalid file format: The sheet '{sheet_name}' is empty."
            
            # Check if any of the required fields contain empty values (null values)
            for field in field_validations[sheet_name]:
                if sheet_df[field].isnull().any():
                    return f"Invalid file format: The field '{field}' in sheet '{sheet_name}' contains empty values."
        
        # If all checks pass, return a success message
        return "File is valid."
        
    except Exception as e:
        # If there was an error processing the file, return the error message
        return f"Error processing the file: {e}"


# Run the Flask app when the script is executed
if __name__ == '__main__':
    app.run(debug=True)
