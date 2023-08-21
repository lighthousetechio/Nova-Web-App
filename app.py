'''
Nova Payroll Processor
filename: app.py
This file implements the frontend of the payroll-processing web app.
'''
#import pacakges
import os
from flask import Flask, render_template, request, send_file, redirect
from helpers import *

#initialize app
app = Flask(__name__)

# Configure the app to store files
app.config['SHIFT_RECORD_FOLDER'] = 'shift_record'
app.config['OLD_TRACKER_FOLDER'] = 'old_tracker'
app.config['PROCESSED_FILES_FOLDER'] = 'processed_files'

#allowed extension
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

#check if the file names has the extension required
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# render index.html
@app.route('/')
def index():
    return render_template('index.html')

#upload shift record
@app.route('/shift_record', methods=['POST'])
def upload_shift_record():
    delete_files_in_folder("./processed_files")
    file = request.files['shift_record']
    if file and allowed_file(file.filename):
        filename = 'shift_record.xlsx'
        file.save(os.path.join(app.config['SHIFT_RECORD_FOLDER'], filename))
        return "Shift record uploaded successfully."

    # Handle invalid file or file type
    return "Invalid file or file type."

#upload old tracker
@app.route('/old_tracker', methods=['POST'])
def upload_tracker():
    delete_files_in_folder("./processed_files")
    file = request.files['old_tracker']
    if file and allowed_file(file.filename):
        filename = 'old_tracker.xlsx'
        file.save(os.path.join(app.config['OLD_TRACKER_FOLDER'], filename))
        return "Old tracker uploaded successfully."

    # Handle invalid file or file type
    return "Invalid file or file type."

#process file
@app.route('/process', methods=['POST'])
def process_files():
    # Implement file processing logic here
    try:
        #folder paths
        shift_record_path = './shift_record/shift_record.xlsx'
        tracker_path = './old_tracker/old_tracker.xlsx'
        save_path = './processed_files'

        #read shift record and check for errors
        df, PAY_PERIOD, start_date, end_date = read_shift_record(shift_record_path)

        #read old tracker adn check for errors
        manager_rates, non_manager_rates, accrued_hrs, bonus_df, bonus, original_bonus_df, staff_info, prepaid_last_time, unpaid_last_time = read_old_tracker(tracker_path, start_date)
        
        #merge shift record with pay rates from the tracker
        df_shift_merged = merge_shifts(df, staff_info, manager_rates, non_manager_rates)

        #calculate vacation and sick times
        df_shift_merged, time_off, time_off_as_shifts = calc_time_off(df_shift_merged)

        #merge overight shifts unpaid last time with df_shift_merged
        df_shift_merged = pd.concat([unpaid_last_time, df_shift_merged], ignore_index=True)

        #crop the shift record based on pay cycle
        df_shift_merged, df_after_pay_period, prepaid_hours, week_order, PREPAY = crop_shifts(df_shift_merged, start_date, end_date)

        #generate payroll outputs
        non_mgr_pr, mgr_pr, non_mgr_bkd, mgr_bkd, new_accrued_hrs = generate_payroll(df_shift_merged, accrued_hrs, bonus_df, bonus, time_off, manager_rates, staff_info, prepaid_last_time, PAY_PERIOD, week_order, PREPAY)
        
        #output payroll files
        output_payroll_files(save_path, df_shift_merged, staff_info, non_mgr_pr, mgr_pr, non_mgr_bkd, mgr_bkd, new_accrued_hrs, original_bonus_df, time_off_as_shifts, non_manager_rates, manager_rates, prepaid_hours, df_after_pay_period, PAY_PERIOD)
        
        #generate invoice outputs
        shift_list, output, mgr_benefits, df_benefits, total_mgr = generate_invoice(df_shift_merged, manager_rates, non_manager_rates, staff_info, non_mgr_pr, mgr_pr)
        
        #output invoice file
        output_invoice(save_path, shift_list, output, mgr_benefits, df_benefits, total_mgr, df_shift_merged, PAY_PERIOD)
        
        #output machine_readable payroll
        output_underlying(mgr_pr, non_mgr_pr, save_path, PAY_PERIOD)

        #flag success
        return "Files Processed Successfully!"
    except Exception as e:
        # display error
        return str(e)
    
#save files
@app.route('/save', methods=['POST'])
def save_files():
    save_path = app.config['PROCESSED_FILES_FOLDER']
    file_links = []
    # Generate download links for each processed file
    for root, dirs, files in os.walk(save_path):
        for file in files:
            file_path = os.path.join(root, file)
            download_link = f'<a href="/download/{file}" target="_blank">{file}</a>'
            file_links.append(download_link)

    # Create HTML response with download links
    download_links = "<br>".join(file_links)
    return f"Download links:<br>{download_links}"

#download files
@app.route('/download/<filename>')
def download_file(filename):
    # Get the path to the processed file
    processed_folder = app.config['PROCESSED_FILES_FOLDER']
    file_path = os.path.join(processed_folder, filename)
    # Send the file for download
    return send_file(file_path, as_attachment=True)

#refresh the processor for a new session.
@app.route('/refresh')
def refresh_page():
    # Extract the URL from the Referer header or default to the index page
    delete_files_in_folder("./shift_record")
    delete_files_in_folder("./old_tracker")
    delete_files_in_folder("./processed_files")
    referer_url = request.headers.get('Referer', '/')
    return redirect(referer_url)


if __name__ == '__main__':
    app.run(debug=True)
