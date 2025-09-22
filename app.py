from flask import Flask, request, render_template, send_file, url_for, flash, redirect
import pandas as pd
import os
import re

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.abspath('Uploads')  # Absolute path
app.config['OUTPUT_FOLDER'] = os.path.abspath('outputs')  # Absolute path
app.secret_key = 'your-secret-key'  # Required for flash messages

# Create directories if they don't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if files are uploaded
        if 'old_to_new' not in request.files or 'data_2025' not in request.files or 'data_2002' not in request.files:
            flash('Please upload all three files.')
            return redirect(request.url)

        old_to_new_file = request.files['old_to_new']
        data_2025_file = request.files['data_2025']
        data_2002_file = request.files['data_2002']

        # Validate file extensions
        if not all(file.filename.endswith('.xlsx') for file in [old_to_new_file, data_2025_file, data_2002_file]):
            flash('All files must be .xlsx format.')
            return redirect(request.url)

        # Save uploaded files
        old_to_new_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Mallisala Old to New.xlsx')
        data_2025_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Mallisala.xlsx')
        data_2002_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Mallisala2002.xlsx')

        try:
            old_to_new_file.save(old_to_new_path)
            data_2025_file.save(data_2025_path)
            data_2002_file.save(data_2002_path)

            # Process files
            process_files(old_to_new_path, data_2025_path, data_2002_path)
            print("Processing completed, attempting to render downloads.html")
            try:
                return render_template('downloads.html')
            except Exception as e:
                print(f"Error rendering downloads.html: {str(e)}")
                raise Exception(f"Failed to render downloads.html: {str(e)}")
        except Exception as e:
            print(f"Error in index route: {str(e)}")
            flash(f'Error processing files: {str(e)}')
            return redirect(request.url)

    return render_template('upload.html')

def process_files(old_to_new_path, data_2025_path, data_2002_path):
    # Read Excel files with debugging
    try:
        xl = pd.ExcelFile(old_to_new_path)
        print(f"Sheets in Mallisala Old to New.xlsx: {xl.sheet_names}")
        old_to_new = pd.read_excel(old_to_new_path, sheet_name=0)
        data_2025 = pd.read_excel(data_2025_path, sheet_name=0)
        data_2002 = pd.read_excel(data_2002_path, sheet_name=0)
    except Exception as e:
        raise Exception(f"Failed to read Excel files: {str(e)}")

    # Debug: Print raw column names and first few rows
    print("Raw columns in Mallisala Old to New.xlsx:", old_to_new.columns.tolist())
    print("First 2 rows in Mallisala Old to New.xlsx:\n", old_to_new.head(2).to_string())
    print("Raw columns in Mallisala.xlsx:", data_2025.columns.tolist())
    print("Raw columns in Mallisala2002.xlsx:", data_2002.columns.tolist())

    # Sanitize column names aggressively
    def sanitize_column_name(col):
        col = str(col).strip()
        col = re.sub(r'[^\x00-\x7F]+', '', col)  # Remove non-ASCII characters
        col = re.sub(r'\s+', ' ', col)  # Replace multiple spaces with single space
        return col.upper()

    old_to_new.columns = [sanitize_column_name(col) for col in old_to_new.columns]
    data_2025.columns = [sanitize_column_name(col) for col in data_2025.columns]
    data_2002.columns = [sanitize_column_name(col) for col in data_2002.columns]

    # Debug: Print sanitized column names
    print("Sanitized columns in Mallisala Old to New.xlsx:", old_to_new.columns.tolist())
    print("Sanitized columns in Mallisala.xlsx:", data_2025.columns.tolist())
    print("Sanitized columns in Mallisala2002.xlsx:", data_2002.columns.tolist())

    # Verify required columns
    required_cols_old_to_new = ['EPIC_NO', 'PREV_EPIC_NO']
    required_cols_2025 = ['EPIC NO.', "ELECTOR'S NAME", 'RELATIVE NAME', 'AC NO.', 'PART NO.', 'SERIAL NO.']
    required_cols_2002 = ['EPIC', 'NAME_ENG', 'RELATION NAME', 'OLD AC NO', 'OLD PART_NO', 'OLD PART SL.NO']

    for df, cols, name in [
        (old_to_new, required_cols_old_to_new, 'Mallisala Old to New.xlsx'),
        (data_2025, required_cols_2025, 'Mallisala.xlsx'),
        (data_2002, required_cols_2002, 'Mallisala2002.xlsx')
    ]:
        missing_cols = [col for col in cols if col not in df.columns]
        if missing_cols:
            raise Exception(f"Missing required columns in {name}: {missing_cols}")

    # Output column mapping for data mapping.xlsx and data_mapping_name_match.xlsx
    output_cols = {
        'EPIC_NO': 'EPIC NO (2025)',
        'AC NO.': 'AC NO (2025)',
        'PART NO.': 'PART NO (2025)',
        'SERIAL NO.': 'SERIAL NO IN PART (2025)',
        "ELECTOR'S NAME": 'Name (2025)',
        'RELATIVE NAME': 'Relation Name (2025)',
        'NAME_ENG': 'Name (2002)',
        'RELATION NAME': 'Relation Name (2002)',
        'OLD AC NO': 'OLD AC NO',
        'OLD PART_NO': 'OLD PART NO',
        'OLD PART SL.NO': 'OLD PART SERIAL NO',
        'PREV_EPIC_NO': 'OLD EPIC NO'
    }
    output_cols_direct = {
        'EPIC NO.': 'EPIC NO (2025)',
        'AC NO.': 'AC NO (2025)',
        'PART NO.': 'PART NO (2025)',
        'SERIAL NO.': 'SERIAL NO IN PART (2025)',
        "ELECTOR'S NAME": 'Name (2025)',
        'RELATIVE NAME': 'Relation Name (2025)',
        'NAME_ENG': 'Name (2002)',
        'RELATION NAME': 'Relation Name (2002)',
        'OLD AC NO': 'OLD AC NO',
        'OLD PART_NO': 'OLD PART NO',
        'OLD PART SL.NO': 'OLD PART SERIAL NO',
        'EPIC': 'OLD EPIC NO'
    }

    # 1. Generate data mapping.xlsx
    print("Performing first merge: old_to_new with data_2002 on PREV_EPIC_NO and EPIC")
    merged1 = pd.merge(old_to_new, data_2002, how='left', left_on='PREV_EPIC_NO', right_on='EPIC')
    print("Columns after first merge:", merged1.columns.tolist())

    print("Performing second merge: merged1 with data_2025 on EPIC_NO and EPIC NO.")
    final_merge = pd.merge(merged1, data_2025, how='left', left_on='EPIC_NO', right_on='EPIC NO.')
    print("Columns after second merge:", final_merge.columns.tolist())

    # Debug: Print exact column names as strings
    print("Exact column names in final_merge:", [str(col) for col in final_merge.columns])

    # Reset index to avoid any indexing issues
    final_merge = final_merge.reset_index(drop=True)

    # Select specific columns to avoid conflicts
    select_cols = [
        'EPIC_NO',
        'AC NO.',
        'PART NO.',
        'SERIAL NO.',
        "ELECTOR'S NAME",
        'RELATIVE NAME',
        'NAME_ENG',
        'RELATION NAME',
        'OLD AC NO',
        'OLD PART_NO',
        'OLD PART SL.NO',
        'PREV_EPIC_NO'
    ]

    # Debug: Print columns being selected
    print("Columns being selected for final_df1:", select_cols)

    # Check if all required columns are present
    missing_output_cols = [col for col in select_cols if col not in final_merge.columns]
    if missing_output_cols:
        raise Exception(f"Missing columns in final_merge: {missing_output_cols}")

    # Debug: Test column access
    for col in select_cols:
        try:
            _ = final_merge[col]
            print(f"Successfully accessed column: {col}")
        except Exception as e:
            print(f"Failed to access column {col}: {str(e)}")

    # Manually copy columns to avoid indexing issues
    final_df1 = pd.DataFrame()
    for col in select_cols:
        try:
            final_df1[col] = final_merge[col]
        except Exception as e:
            raise Exception(f"Error accessing column {col}: {str(e)}")

    # Debug: Print columns before renaming
    print("Columns in final_df1 before renaming:", final_df1.columns.tolist())

    # Manually rename columns
    new_columns = []
    for col in final_df1.columns:
        new_columns.append(output_cols.get(col, col))
    final_df1.columns = new_columns

    # Save to Excel
    try:
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'data mapping.xlsx')
        output_path = os.path.abspath(output_path)
        print(f"Saving data mapping.xlsx to: {output_path}")
        final_df1.insert(0, 'Sl.No', range(1, len(final_df1) + 1))
        final_df1.to_excel(output_path, index=False)
        if os.path.exists(output_path):
            print(f"Successfully saved data mapping.xlsx to: {output_path}")
        else:
            raise Exception(f"Failed to verify existence of data mapping.xlsx at: {output_path}")
    except Exception as e:
        raise Exception(f"Error saving data mapping.xlsx: {str(e)}")

    # 2. Generate data_mapping_name_match.xlsx
    print("Generating data_mapping_name_match.xlsx")
    final_merge["ELECTOR'S NAME"] = final_merge["ELECTOR'S NAME"].str.strip().str.upper()
    final_merge['NAME_ENG'] = final_merge['NAME_ENG'].str.strip().str.upper()
    name_match_df = final_merge[final_merge["ELECTOR'S NAME"] == final_merge['NAME_ENG']]
    name_match_df = name_match_df.reset_index(drop=True)

    # Debug: Print columns being selected
    print("Columns being selected for final_df2:", select_cols)

    # Debug: Test column access
    for col in select_cols:
        try:
            _ = name_match_df[col]
            print(f"Successfully accessed column in name_match_df: {col}")
        except Exception as e:
            print(f"Failed to access column in name_match_df {col}: {str(e)}")

    # Manually copy columns to avoid indexing issues
    final_df2 = pd.DataFrame()
    for col in select_cols:
        try:
            final_df2[col] = name_match_df[col]
        except Exception as e:
            raise Exception(f"Error accessing column {col} in name_match_df: {str(e)}")

    # Debug: Print columns before renaming
    print("Columns in final_df2 before renaming:", final_df2.columns.tolist())

    # Manually rename columns
    new_columns = []
    for col in final_df2.columns:
        new_columns.append(output_cols.get(col, col))
    final_df2.columns = new_columns

    # Save to Excel
    try:
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'data_mapping_name_match.xlsx')
        output_path = os.path.abspath(output_path)
        print(f"Saving data_mapping_name_match.xlsx to: {output_path}")
        final_df2.insert(0, 'Sl.No', range(1, len(final_df2) + 1))
        final_df2.to_excel(output_path, index=False)
        if os.path.exists(output_path):
            print(f"Successfully saved data_mapping_name_match.xlsx to: {output_path}")
        else:
            raise Exception(f"Failed to verify existence of data_mapping_name_match.xlsx at: {output_path}")
    except Exception as e:
        raise Exception(f"Error saving data_mapping_name_match.xlsx: {str(e)}")

    # 3. Generate data_mapping_name_match_direct.xlsx
    print("Generating data_mapping_name_match_direct.xlsx")
    try:
        data_2025["ELECTOR'S NAME"] = data_2025["ELECTOR'S NAME"].str.strip().str.upper()
        data_2002['NAME_ENG'] = data_2002['NAME_ENG'].str.strip().str.upper()
        merged_direct = pd.merge(data_2025, data_2002, how='inner', left_on="ELECTOR'S NAME", right_on='NAME_ENG')
        final_df3 = merged_direct[list(output_cols_direct.keys())].rename(columns=output_cols_direct)
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'data_mapping_name_match_direct.xlsx')
        output_path = os.path.abspath(output_path)
        print(f"Saving data_mapping_name_match_direct.xlsx to: {output_path}")
        final_df3.insert(0, 'Sl.No', range(1, len(final_df3) + 1))
        final_df3.to_excel(output_path, index=False)
        if os.path.exists(output_path):
            print(f"Successfully saved data_mapping_name_match_direct.xlsx to: {output_path}")
        else:
            raise Exception(f"Failed to verify existence of data_mapping_name_match_direct.xlsx at: {output_path}")
    except Exception as e:
        raise Exception(f"Error saving data_mapping_name_match_direct.xlsx: {str(e)}")

    # 4. Generate data_mapping_merged.xlsx
    print("Generating data_mapping_merged.xlsx")
    try:
        df_all = final_df1.drop(columns=['Sl.No'])
        df_name_match_direct = final_df3.drop(columns=['Sl.No'])
        merged = pd.concat([df_all, df_name_match_direct], ignore_index=True)
        string_cols = ['EPIC NO (2025)', 'Name (2025)', 'Relation Name (2025)', 'Name (2002)', 'Relation Name (2002)', 'OLD EPIC NO']
        for col in string_cols:
            if col in merged.columns:
                merged[col] = merged[col].str.strip().str.upper()
        merged_unique = merged.drop_duplicates()
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'data_mapping_merged.xlsx')
        output_path = os.path.abspath(output_path)
        print(f"Saving data_mapping_merged.xlsx to: {output_path}")
        merged_unique.insert(0, 'Sl.No', range(1, len(merged_unique) + 1))
        merged_unique.to_excel(output_path, index=False)
        if os.path.exists(output_path):
            print(f"Successfully saved data_mapping_merged.xlsx to: {output_path}")
        else:
            raise Exception(f"Failed to verify existence of data_mapping_merged.xlsx at: {output_path}")
    except Exception as e:
        raise Exception(f"Error saving data_mapping_merged.xlsx: {str(e)}")

    # Debug: Print columns in merged_unique
    print("Columns in merged_unique:", merged_unique.columns.tolist())

    # 5. Generate data_mapping_merged_cleaned.xlsx
    print("Generating data_mapping_merged_cleaned.xlsx")
    try:
        if 'Name (2025)' not in merged_unique.columns:
            raise Exception("Column 'Name (2025)' not found in merged_unique")
        if 'Name (2002)' not in merged_unique.columns:
            raise Exception("Column 'Name (2002)' not found in merged_unique")
        
        # Create a copy to avoid SettingWithCopyWarning
        merged_unique = merged_unique.copy()
        merged_unique.loc[:, 'Name (2025)'] = merged_unique['Name (2025)'].replace(r'^\s*$', pd.NA, regex=True)
        merged_unique.loc[:, 'Name (2002)'] = merged_unique['Name (2002)'].replace(r'^\s*$', pd.NA, regex=True)
        cleaned = merged_unique.dropna(subset=['Name (2025)', 'Name (2002)']).copy()
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'data_mapping_merged_cleaned.xlsx')
        output_path = os.path.abspath(output_path)
        print(f"Saving data_mapping_merged_cleaned.xlsx to: {output_path}")
        cleaned.loc[:, 'Sl.No'] = range(1, len(cleaned) + 1)
        cleaned.to_excel(output_path, index=False)
        if os.path.exists(output_path):
            print(f"Successfully saved data_mapping_merged_cleaned.xlsx to: {output_path}")
        else:
            raise Exception(f"Failed to verify existence of data_mapping_merged_cleaned.xlsx at: {output_path}")
    except Exception as e:
        raise Exception(f"Error saving data_mapping_merged_cleaned.xlsx: {str(e)}")

    # 6. Generate data_mapping_unique_part_serial.xlsx
    print("Generating data_mapping_unique_part_serial.xlsx")
    try:
        unique_part = cleaned.drop_duplicates(subset=['PART NO (2025)', 'SERIAL NO IN PART (2025)'], keep='first').copy()
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'data_mapping_unique_part_serial.xlsx')
        output_path = os.path.abspath(output_path)
        print(f"Saving data_mapping_unique_part_serial.xlsx to: {output_path}")
        unique_part.loc[:, 'Sl.No'] = range(1, len(unique_part) + 1)
        unique_part.to_excel(output_path, index=False)
        if os.path.exists(output_path):
            print(f"Successfully saved data_mapping_unique_part_serial.xlsx to: {output_path}")
        else:
            raise Exception(f"Failed to verify existence of data_mapping_unique_part_serial.xlsx at: {output_path}")
    except Exception as e:
        raise Exception(f"Error saving data_mapping_unique_part_serial.xlsx: {str(e)}")

    # 7. Generate data_mapping_unique_relation_name.xlsx
    print("Generating data_mapping_unique_relation_name.xlsx")
    try:
        cleaned = cleaned.copy()
        cleaned.loc[:, 'Relation Name (2025)'] = cleaned['Relation Name (2025)'].str.strip().str.upper()
        unique_relation = cleaned.drop_duplicates(subset=['Relation Name (2025)'], keep='first').copy()
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'data_mapping_unique_relation_name.xlsx')
        output_path = os.path.abspath(output_path)
        print(f"Saving data_mapping_unique_relation_name.xlsx to: {output_path}")
        unique_relation.loc[:, 'Sl.No'] = range(1, len(unique_relation) + 1)
        unique_relation.to_excel(output_path, index=False)
        if os.path.exists(output_path):
            print(f"Successfully saved data_mapping_unique_relation_name.xlsx to: {output_path}")
        else:
            raise Exception(f"Failed to verify existence of data_mapping_unique_relation_name.xlsx at: {output_path}")
    except Exception as e:
        raise Exception(f"Error saving data_mapping_unique_relation_name.xlsx: {str(e)}")

@app.route('/download/<filename>')
def download(filename):
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        file_path = os.path.abspath(file_path)
        print(f"Attempting to download file: {file_path}")
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        print(f"Error downloading {filename}: {str(e)}")
        flash(f'File {filename} not found: {str(e)}')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)