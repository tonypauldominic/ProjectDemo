# ProjectDemo
A collection of my sample SQL files

# Global variables
new_file = None
file_type_var = None
directory_path = None
messages = []
teilkonzern_directory_path = r'U:\rlbnas1_rlb_finrep\KRW_GRM\2023\3. Quartal\Konzernabgleich\Teilkonzerne'






def get_date_from_filename(filename: str):
    try:
        return datetime.strptime(filename.split('_')[0], '%Y-%m-%d')
    except ValueError:
        print(f'Skipping file {filename} because its name does not start with a date')
        return None

def convert_to_numeric(value):
    try:
        return pd.to_numeric(value)
    except ValueError:
        return value



def is_valid_date(date_string, date_format='%Y%m%d'):
    try:
        datetime.strptime(date_string, date_format)
        return True
    except ValueError:
        return False



def execute_sql(condition, new_filepath):
    conn = pyodbc.connect(f'driver={driver_name};server={server_name};database={database_name};trusted_connection=yes')

    today = datetime.today()

    # Calculate the quarter
    month = today.month
    if 1 <= month <= 3:
        quartal = 4
        stichtag = f'{today.year - 1}1231'
    elif 4 <= month <= 6:
        quartal = 1
        stichtag = f'{today.year}0331'
    elif 7 <= month <= 9:
        quartal = 2
        stichtag = f'{today.year}0630'
    elif 10 <= month <= 12:
        quartal = 3
        stichtag = f'{today.year}0930'
        
      


    # Show a dialog box to the user asking if they want to change the 'stichtag'
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    user_response = messagebox.askyesno("Stichtag Ändern", f"Der Stichtag ist {stichtag}. Möchtest du den Stichtag ändern?")
    root.destroy()  # Destroy the root window


    # If the user chooses to change the 'stichtag', show a dialog box to get the new 'stichtag'
    if user_response:
        while True:
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            stichtag = simpledialog.askstring("Input", "Bitte den Neuen Stichtag eingeben (YYYY-MM-DD):", parent=root)
            root.destroy()  # Destroy the root window

            if is_valid_date(stichtag):
                break
            else:
                messagebox.showerror("Ungültiges Datum", "Das eingegebene Datum ist ungültig. Bitte das Datum in 'YYYY-MM-DD' format eingeben.")
     

    sql = sql_script.format(stichtag=stichtag) + condition + " order by KonsolidierungskreisIDL, PartnerID, KonzernkontoNr"




 

    print("Executing SQL query: ", sql)

    df = pd.read_sql(sql, conn)
    print(f"Number of rows obtained from SQL Server: {len(df)}")

   # Check if 'KonzernkontoNr' and 'PartnerCodeIDL' are in df.columns before creating 'key'
    if 'KonzernkontoNr' in df.columns and 'PartnerCodeIDL' in df.columns and 'KonsolidierungskreisIDL' in df.columns:
     df['key'] = (df['KonzernkontoNr'].astype(str).str.zfill(3) + "_" + df['PartnerCodeIDL'].astype(str).str.lstrip("0") + "_" + df['KonsolidierungskreisIDL'].astype(str).str.zfill(3))
     df['key'] = df['key'].apply(lambda x: x.strip().lower())



    if 'KonzernkontoNr' in df.columns:
        df['KonzernkontoNr'] = df['KonzernkontoNr'].apply(convert_to_numeric)
    else:
        print("Column 'KonzernkontoNr' is not present in the DataFrame")

    if 'PartnerCodeIDL' in df.columns:
        df['PartnerCodeIDL'] = df['PartnerCodeIDL'].apply(convert_to_numeric)
    else:
        print("Column 'PartnerCodeIDL' is not present in the DataFrame")

    cols = df.columns.tolist()
    cols = cols[-1:] + cols[:-1]
    df = df[cols]

    df.columns = [col.lower() for col in df.columns]

    print(df.head())
    print("Number of rows in dataframe: ", len(df))

    # Load the Excel workbook and clear existing data
    try:
        wb = load_workbook(new_filepath)
        ws = wb['SQL-SPOT']
    except Exception as e:
        print(f"Unable to load workbook due to error: {str(e)}")
        return

    # Convert DataFrame column names to lowercase
    df.columns = df.columns.str.lower()

    # Clear the contents of the columns A to O from row 11 onwards before writing new data
    for row in ws.iter_rows(min_row=11, min_col=1, max_col=15):  # Column A to O corresponds to 1 to 15
        for cell in row:
            cell.value = None
    # Write the DataFrame to the worksheet starting from the 11th row
    for i, row in enumerate(df.values, start=11):
        for j, value in enumerate(row, start=1):
            # Only write the value if the cell is empty
            if ws.cell(row=i, column=j).value is None:
                ws.cell(row=i, column=j, value=value)

    if file_selection_var.get() == 'Teilkonzern':
        # Load workbook with data_only=False for copying formulas
        latest_file = find_most_recent_file(teilkonzern_directory_path, 'Konzernabgleich_Teilkonzern.xlsx')
        src_book_formula = load_workbook(latest_file, data_only=False)
        src_sheet_formula = src_book_formula['SQL-SPOT']

        # Copy columns with formulas
        for col_index in range(15, 22):  # Column P to V corresponds to indices 15 to 21
            for src_row, tgt_row in zip(src_sheet_formula.iter_rows(min_row=11), ws.iter_rows(min_row=11)):
                src_cell = src_row[col_index]
                tgt_cell = tgt_row[col_index]
                tgt_cell.value = src_cell.value
    else:
        for row in ws.iter_rows(min_row=11, min_col=16, max_col=22):  # Column P to V corresponds to 16 to 22
            for cell in row:
                cell.value = None

    # Save the workbook after writing all the data
    wb.save(new_filepath)

    print(f"Data copied to {new_filepath}")

    return df



def find_most_recent_ifrs_and_crr(directory):
    # This is a placeholder function. 
    return ['IFRS_file.xlsx', 'CRR_file.xlsx']
def add_anmerkung_column():
    column_names = [['Bearbeiter', 'Bearbeiter:in'], 'Anmerkung', 'Bleibt Differenz bestehen?', 'Werte die bestehen bleiben', 'Wann wird Korrigiert?']
    column_letters_to_copy = ['P', 'Q', 'R', 'S', 'T', 'U', 'V']

    file_selection = file_selection_var.get()

    # Avoid repeated function call
    recent_ifrs, recent_crr = find_most_recent_ifrs_and_crr(directory_path)

    for file_type in ['IFRS', 'CRR']:
        if file_type not in file_selection:
            continue

        logging.info(f"Processing file type: {file_type}")

        if file_type == 'IFRS':
            src_filepath = os.path.join(directory_path, recent_ifrs)
        else:  
            src_filepath = os.path.join(directory_path, recent_crr)

        global new_file
        tgt_filepath = new_file

        try:
            src_book_data = load_workbook(src_filepath, data_only=True)
            src_sheet_data = src_book_data['SQL-SPOT']

            src_book_formula = load_workbook(src_filepath, data_only=False)
            src_sheet_formula = src_book_formula['SQL-SPOT']

            logging.info(f"Source file loaded: {src_filepath}")
        except Exception as e:
            logging.error(f"Unable to load source file due to error: {str(e)}")
            continue

        try:
            tgt_book = load_workbook(tgt_filepath, data_only=False)
            tgt_sheet = tgt_book['SQL-SPOT']

            logging.info(f"Loaded target file: {tgt_filepath}")
        except Exception as e:
            logging.error(f"Unable to load target workbook due to error: {str(e)}")
            continue

        # Use list comprehension
        src_key_to_rows = {row[0].value.strip().lower(): row for row in src_sheet_data.iter_rows(min_row=11) if row[0].value is not None}
        tgt_key_to_rows = {row[0].value.strip().lower(): row for row in tgt_sheet.iter_rows(min_row=11) if row[0].value is not None}

        # Create column name to letter mapping
        col_name_to_letter = {cell.value: cell.column_letter for cell in src_sheet_data[10]}

        key_count = 0
        for key, src_row in src_key_to_rows.items():
            if key in tgt_key_to_rows:
                tgt_row = tgt_key_to_rows[key]
                for col_names in column_names:
                    if isinstance(col_names, str):
                        col_names = [col_names]
                    for col_name in col_names:
                        # Use mapping to get column letter
                        col_letter = col_name_to_letter.get(col_name)

                        if col_letter is not None:
                            col_index = column_index_from_string(col_letter) - 1  
                            src_cell = src_row[col_index]
                            tgt_cell = tgt_row[col_index]
                            tgt_cell.value = src_cell.value

                            logging.debug(f"Source cell value: {src_cell.value}, Target cell value: {tgt_cell.value}")

                key_count += 1
            else:
                logging.debug(f"Key {key} not found in tgt_key_to_rows")  

        logging.info(f"Successfully mapped columns for {key_count} keys")

        for col_letter in column_letters_to_copy:
            col_index = column_index_from_string(col_letter) - 1  
            for src_row, tgt_row in zip(src_sheet_formula.iter_rows(min_row=11), tgt_sheet.iter_rows(min_row=11)):
               

                src_cell = src_row[col_index]
                tgt_cell = tgt_row[col_index]
                tgt_cell.value = src_cell.value

        try:
            tgt_book.save(tgt_filepath)
            logging.info(f"Target workbook saved successfully at: {tgt_filepath}")

            root = tk.Tk()
            root.withdraw()  
            messagebox.showinfo("Success", f"The operation completed successfully for {file_type}.")
            root.destroy()  
        except Exception as e:
            logging.error(f"Unable to save workbook due to error: {str(e)}")
            continue
def execute_and_check_sql(sql_query, new_filepath):
    # Get the user's file selection
    file_selection = file_selection_var.get()

    # Depending on the user's selection, create the new file(s)
    if file_selection == 'IFRS':
        # Create the new IFRS file
        latest_ifrs_file = find_most_recent_file(directory_path, 'IFRS.xlsx')
        new_filename, new_filepath = create_new_file_with_today_date(directory_path, latest_ifrs_file)
    elif file_selection == 'CRR':
        # Create the new CRR file
        latest_crr_file = find_most_recent_file(directory_path, 'CRR.xlsx')
        new_filename, new_filepath = create_new_file_with_today_date(directory_path, latest_crr_file)
    elif file_selection == 'Teilkonzern':
        # Create the new Teilkonzern file
        latest_teilkonzern_file = find_most_recent_file(teilkonzern_directory_path, 'Konzernabgleich_Teilkonzern.xlsx')
        new_filename, new_filepath = create_new_file_with_today_date(teilkonzern_directory_path, latest_teilkonzern_file)

    # Execute SQL and get data
    df = execute_sql(sql_query, new_filepath)
    return df




def create_new_file_with_today_date(directory: str, filename_pattern: str):
    # Get a list of all files that match the pattern
    files = glob.glob(os.path.join(directory, filename_pattern))

    # Sort the files by date
    files.sort(key=os.path.getmtime, reverse=True)

    # Get the most recent file
    filename = files[0] if files else None

    if filename is None:
        print(f"No files found that match the pattern '{filename_pattern}'")
        return None

    # Get today's date in the format 'YYYY-MM-DD'
    today = datetime.today().strftime('%Y-%m-%d')
    
    # Replace the date in the filename with today's date
    new_filename = re.sub(r'\d{4}-\d{2}-\d{2}', today, filename)
    new_filepath = os.path.join(directory, new_filename)
    print(f"Trying to create new file: {new_filename}") 

    # If the new filename is the same as the original filename, return without creating a new file
    if new_filename == filename:
        print(f"The file with today's date already exists: {new_filename}")
        return filename, new_filepath

    # Create a copy of the existing .xlsx file
    try:
        shutil.copyfile(filename, new_filepath)
    except FileNotFoundError:
        print(f"The file {filename} does not exist.")
        return None

    # Load the new .xlsx file
    wb_new = load_workbook(new_filepath)

    # Modify the new workbook as needed
    # This is where you would add any modifications you need to make to the new workbook

    # Save the new workbook
    #wb_new.save(new_filepath)

    # Save file's name to a global variable
    global new_file
    new_file = new_filepath
    print(f"Target file: {new_file}")

    print(f'New file {new_filename} created')
    return new_filename, new_filepath


def check_data(df_original, filepath, sheet_name='SQL-SPOT'):
    # Check if df_original is None
    if df_original is None:
        print('df_original is None')
        return

    # Attempt to load the workbook and worksheet
    try:
        wb = load_workbook(filepath)
        ws = wb[sheet_name]
    except FileNotFoundError:
        print(f"The file {filepath} does not exist")
        return
    except PermissionError:
        print(f"Permission denied: could not open file {filepath}")
        return
    except InvalidFileException:
        print(f"The file {filepath} is not a valid .xlsx file")
        return
    except KeyError:
        print(f"The worksheet {sheet_name} does not exist in the file {filepath}")
        return

    # Specify the columns you want to read (A-O corresponds to indices 0-14)
    cols = list(range(0, 15))

    # Read the column names from the Excel file
    try:
        df_excel = pd.read_excel(filepath, sheet_name=sheet_name, header=9, usecols=cols)
    except Exception as e:
        print(f"Error reading column names from the Excel file: {e}")
        return

    # Print column names
        print('df_original column names:')
        print(df_original.columns.tolist())
        print('df_excel column names:')
        print(df_excel.columns.tolist())
    # Check if the column names are identical
    if set([col.lower() for col in df_original.columns.tolist()]) == set([col.lower() for col in df_excel.columns.tolist()]):
        print('Column names are identical')
    else:
        print('Column names are not identical')





def find_most_recent_ifrs_and_crr(directory: str):
    latest_ifrs_file = None
    latest_ifrs_date = None
    latest_crr_file = None
    latest_crr_date = None
    today = datetime.today().date()  # Get today's date

    for filename in os.listdir(directory):
        date = get_date_from_filename(filename)
        if date and date.date() < today:  # Only consider files before today
            if 'IFRS.' in filename and (latest_ifrs_date is None or date > latest_ifrs_date):
                latest_ifrs_file = filename
                latest_ifrs_date = date
            elif 'CRR.' in filename and (latest_crr_date is None or date > latest_crr_date):
                latest_crr_file = filename
                latest_crr_date = date

    if latest_ifrs_file is not None:
        print(f'Most recent IFRS file is {latest_ifrs_file}')
    else:
        print(f'No valid IFRS file found in directory {directory}')

    if latest_crr_file is not None:
        print(f'Most recent CRR file is {latest_crr_file}')
    else:
        print(f'No valid CRR file found in directory {directory}')

    return latest_ifrs_file, latest_crr_file

def create_new_teilkonzern_file():
    # Define the directory for 'Teilkonzern' files
    tk_directory_path = r'U:\rlbnas1_rlb_finrep\KRW_GRM\2023\3. Quartal\Konzernabgleich\Teilkonzerne'

    # Get today's date in the format 'YYYY-MM-DD'
    today = datetime.today().strftime('%Y-%m-%d')

    # Create the filename using the current date
    filename = f"{today}_Konzernabgleich_Teilkonzern.xlsx"

    # Create the full path to the new file
    new_filepath = os.path.join(tk_directory_path, filename)

    # Create a new Excel file at the new file path
    wb = Workbook()
    wb.save(new_filepath)

    return new_filepath


def on_btn_click(i):
    # Assuming `file_selection` is a variable containing the selected file type
    # Assuming `directory_path` is the path where the new file should be created
    # Assuming `latest_crr_file` and `latest_ifrs_file` are the latest file names for CRR and IFRS respectively

    if file_selection in ['CRR', 'Both']:
        # Create the new CRR file and get its filename and path
        new_filename, new_filepath = create_new_file_with_today_date(directory_path, latest_crr_file)

        # Load the newly created file using its full path
        wb = openpyxl.load_workbook(new_filepath)

        # Assume `df` is a DataFrame containing the data to be written to the new file
        ws = wb.active
        for r in dataframe_to_rows(df, index=False, header=False):
            ws.append(r)

        # Save the changes to the new file
        wb.save(new_filepath)

    if file_selection in ['IFRS', 'Both']:
        # Create the new IFRS file and get its filename and path
        new_filename, new_filepath = create_new_file_with_today_date(directory_path, latest_ifrs_file)

        # Load the newly created file using its full path
        wb = openpyxl.load_workbook(new_filepath)

        # Assume `df` is a DataFrame containing the data to be written to the new file
        ws = wb.active
        for r in dataframe_to_rows(df, index=False, header=False):
            ws.append(r)

        # Save the changes to the new file
        wb.save(new_filepath)




# Assume you have a list of PartnerCodeIDL codes
partnercode_list = ['001', '013', '030', '031', '054', '057', '058', '071', '076', '078', '083', '090', '093', '097', '113', '114', '119', '127', '150', '156',
'157', '165', '169', '176', '195', '200', '205', '208', '210', '211', '212', '237', '240', '242', '245', '248', '500', '501', '502', '503',
'509', '510', '513', '514', '520', '521', '524', '528', '529', '543', '550', '552', '553', '554', '556', '557', '558', '564',
'565', '566', '567', '568', '575', '576', '583', '586', '590', '592', '596', '597', '602', '603', '606', '607', '608', '610', '611', '613',
'617', '619', '624', '625', '628', '633', '634', '636', '637', '640', '641', '645', '647', '648', '649', '650', '652', '653', '654', '656', '660', '829', '851', '852', '853', '859', '960', 'B01', 'B03', 'B04', 'B13', 'B14', 'B20', 'B24', 'B25', 'B31', 'B33', 'B44', 'B45', 'B47', 'B50',
'B54', 'B55', 'B58', 'B60', 'B63', 'B65', 'B66', 'B73', 'B74', 'B80', 'B81', 'B82', 'B91', 'B92', 'B93', 'B94', 'B95', 'E27', 'E30', 'E31', 'E32',
'E34', 'E37', 'E41', 'E42', 'E44', 'E45', 'E46', 'E47', 'E50', 'E57', 'E61', 'E64', 'E77', 'E89', 'E90', 'E91', 'E92', 'E93', 'E94', 'E95',
'E96', 'E99', 'F01', 'F02', 'F03', 'F04', 'F05', 'F06', 'F07', 'F08']

# Create the main window
window = tk.Tk()
window.title('Konzernabgleich')

file_type_var = tk.StringVar()

directory_path = r'U:\rlbnas1_rlb_finrep\KRW_GRM\2023\3. Quartal\Konzernabgleich\ARCHIV'


def find_most_recent_file(directory: str, filename_pattern: str):
    files = glob.glob(os.path.join(directory, '*' + filename_pattern))
    files.sort(key=os.path.getmtime, reverse=True)
    return files[0] if files else None




def add_teilkonzern_columns(src_filepath, tgt_filepath):
    try:
        src_book_formula = load_workbook(src_filepath, data_only=False)
        src_sheet_formula = src_book_formula['SQL-SPOT']

        tgt_book = load_workbook(tgt_filepath, data_only=False)
        tgt_sheet = tgt_book['SQL-SPOT']
    except Exception as e:
        logging.error(f"Unable to load workbook due to error: {str(e)}")
        return

    column_letters_to_copy = ['P', 'Q', 'R', 'S', 'T', 'U', 'V']
    # Copy columns with formulas
    for col_letter in column_letters_to_copy:
        for row in range(1, src_sheet_formula.max_row + 1):
            cell_formula = src_sheet_formula['{}{}'.format(col_letter, row)].value
            cell_coordinate = '{}{}'.format(col_letter, row)
            add_value_to_merged_cell(tgt_sheet, cell_coordinate, cell_formula)

    tgt_book.save(tgt_filepath)
    print(f"Columns with formulas copied to {tgt_filepath}")
    
    messagebox.showinfo("Success", "Operations were successfully completed.")


def create_new_files():
    # Get the user's file selection
    file_selection = file_selection_var.get()

    # Depending on the user's selection, create the new file(s)
    if file_selection == 'Teilkonzern':
        # Create the new Teilkonzern file
        new_filepath = create_new_teilkonzern_file()
    else:
        directory_path = r'U:\rlbnas1_rlb_finrep\KRW_GRM\2023\3. Quartal\Konzernabgleich\ARCHIV'

        # Find the most recent IFRS and CRR files before creating the new file
        latest_ifrs_file, latest_crr_file = find_most_recent_ifrs_and_crr(directory_path)

        if file_selection == 'IFRS':
            # Create the new IFRS file
            new_filename, new_filepath = create_new_file_with_today_date(directory_path, latest_ifrs_file)
        elif file_selection == 'CRR':
            # Create the new CRR file
            new_filename, new_filepath = create_new_file_with_today_date(directory_path, latest_crr_file)


# Attach create_new_files function to a button.
btn = tk.Button(window, text='Neue Datei erstellen', command=create_new_files)
btn.pack(expand=True)

# Create a Combobox for the file selection
# Create a Combobox for the file selection
file_selection_var = tk.StringVar()
file_selection_combobox = ttk.Combobox(window, textvariable=file_selection_var)
file_selection_combobox['values'] = ('CRR', 'IFRS', 'Teilkonzern')
file_selection_combobox.current(0)  # set 'CRR' as the default value
file_selection_combobox.pack()


stichtag_label = tk.Label(window, text="Stichtag eingeben:")
stichtag_label.pack()
stichtag_entry = tk.Entry(window)
stichtag_entry.pack()

# Create a new frame for SQL selection
sql_selection_frame = tk.Frame(window)
sql_selection_frame.pack()

# Create a combobox for the SQL selection with the parent as sql_selection_frame
sql_option_var = tk.StringVar()
sql_option_names = ['ILG-VIVATI-RR-RTM-BWOHN-EFKO', 'KIFI', 'RLBOOE', 'E45']
sql_option_combobox = ttk.Combobox(sql_selection_frame, textvariable=sql_option_var)
sql_option_combobox['values'] = sql_option_names
sql_option_combobox.current(0)  # Set the default value to the first sql_option_name
sql_option_combobox.pack(side=tk.LEFT)

# Create "Execute SQL Command" button with the parent as sql_selection_frame
def execute_sql_condition():
    selected_option = sql_option_var.get()

    execute_and_check_sql(condition[sql_option_names.index(selected_option)])

execute_sql_button = tk.Button(sql_selection_frame, text='SQL-Befehl ausführen', command=execute_sql_condition)
execute_sql_button.pack(side=tk.LEFT, padx=10)  # Added some padding for aesthetics

# Create a new frame for the PartnerCodeIDL section
partnercode_frame = tk.Frame(window)
partnercode_frame.pack()

# Create a label and entry field for the code
code_label = tk.Label(partnercode_frame, text="PartnerCodeIDL eingeben:")
code_label.pack(side=tk.LEFT)
code_entry = tk.Entry(partnercode_frame)
code_entry.pack(side=tk.LEFT)

# Create radio buttons for the add/delete option
add_delete_var = tk.StringVar()
add_delete_var.set('add')  # set 'add' as the default value
add_radio = tk.Radiobutton(partnercode_frame, text='Hinzufügen', variable=add_delete_var, value='add')
delete_radio = tk.Radiobutton(partnercode_frame, text='Löschen', variable=add_delete_var, value='delete')
add_radio.pack(side=tk.LEFT)
delete_radio.pack(side=tk.LEFT)

def submit_changes():
    codes = code_entry.get().split(',')
    # Strip leading/trailing whitespace and filter out empty strings
    codes = [code.strip() for code in codes if code.strip()]

    # If codes is empty, return from the function
    if not codes:
        return

    action = add_delete_var.get()
    if action == 'add':
        for code in codes:
            if code not in partnercode_list:
                partnercode_list.append(code)
                messagebox.showinfo("Success", f"Code {code} added successfully!")
            else:
                messagebox.showerror("Error", f"Code {code} is already in the list!")
    elif action == 'delete':
        for code in codes:
            if code in partnercode_list:
                partnercode_list.remove(code)
                messagebox.showinfo("Success", f"Code {code} deleted successfully!")
            else:
                messagebox.showerror("Error", f"Code {code} is not in the list!")


submit_btn = tk.Button(partnercode_frame, text='Änderungen übermitteln', command=submit_changes)
submit_btn.pack(side=tk.LEFT)

btn = tk.Button(window, text='Spalten vom Vortag hinzufügen', command=add_anmerkung_column)
btn.pack()


def add_value_to_merged_cell(sheet, cell_coordinate, value):
    range_ = None

    # Unmerge cells
    for merged_cell in sheet.merged_cells.ranges:
        if cell_coordinate in merged_cell:
            range_ = merged_cell
            sheet.unmerge_cells(str(range_))
            break

    # Assign value
    sheet[cell_coordinate].value = value

    # Merge cells back if they were originally merged
    if range_:
        sheet.merge_cells(str(range_))

 

 
def execute_all_tasks():
    # Get the user's file selection
    file_selection = file_selection_var.get()

    # Get the selected SQL query
    selected_sql = sql_option_var.get()
    sql_query = condition[sql_option_names.index(selected_sql)]

    # Depending on the user's selection, create the new file(s)
    if file_selection == 'IFRS':
        # Create the new IFRS file
        latest_ifrs_file = find_most_recent_file(directory_path, 'IFRS.xlsx')
        new_filename, new_filepath = create_new_file_with_today_date(directory_path, latest_ifrs_file)
    elif file_selection == 'CRR':
        # Create the new CRR file
        latest_crr_file = find_most_recent_file(directory_path, 'CRR.xlsx')
        new_filename, new_filepath = create_new_file_with_today_date(directory_path, latest_crr_file)
    elif file_selection == 'Teilkonzern':
        # Create the new Teilkonzern file
        latest_teilkonzern_file = find_most_recent_file(teilkonzern_directory_path, 'Konzernabgleich_Teilkonzern.xlsx')
      

        new_filename, new_filepath = create_new_file_with_today_date(teilkonzern_directory_path, latest_teilkonzern_file)

    # Execute SQL and get data
    df = execute_and_check_sql(sql_query, new_filepath)

    # Submit changes
    submit_changes()

    # Add columns from the previous day for IFRS and CRR files
    if file_selection != 'Teilkonzern':
        add_anmerkung_column()
    else:
        # Add columns with formulas for Teilkonzern files
        add_teilkonzern_columns(latest_teilkonzern_file, new_filepath)



execute_all_btn = tk.Button(window, text='Alles ausführen', command=execute_all_tasks)
execute_all_btn.pack()

# Run the main loop
window.mainloop()


