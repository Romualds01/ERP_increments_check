import sys
import os
from pathlib import Path

rest_api_path = Path(r'\\10.89.140.140\Scripts\rest_api')
if os.path.exists(rest_api_path):
    sys.path.insert(0, str(rest_api_path))
else:
    sys.path.insert(0, str(Path(r'C:\Scripts\rest_api')))

from rest_api_module import *

from datetime import datetime
from dateutil.relativedelta import relativedelta
import pytz

# Current month minus 1 calendar month:
one_month_ago = datetime.now() - relativedelta(months=1)

# closing_month_Mmm = one_month_ago.strftime("%b")  # returns Mmm like "Jan", "Apr"
closing_year_yyyy = one_month_ago.strftime("%Y")  # returns yyyy-string like "2024"
closing_month_mm = one_month_ago.strftime("%m")  # returns mm-string like "01", "04"
# current_year_yy = datetime.now().strftime('%y')  # returns yy-string like "24"
# current_month_mm = datetime.now().strftime("%m")  # returns mm-string like "01", "04"
# current_month_Mmm = datetime.now().strftime("%b")  # returns Mmm like "Jan", "Apr"
today_swe_format = datetime.now().strftime("%Y-%m-%d")  # returns date in Swedish format like "2024-04-25"
now_stockholm = datetime.now(pytz.timezone('Europe/Stockholm')) # datetime object with current date and time in Stockholm timezone

SCRIPT_LOG_FOLDER_PATH = r'\\fdm.se.telenor.net\FDM\PowerBI\ERP_Files\Oracle\_GL_export_check\logs\\'
SCRIPT_LOG_FILE_PATH = os.path.join(SCRIPT_LOG_FOLDER_PATH, now_stockholm.strftime("%y%m%d_%H%M%S") + 'gl_export_check_log_' + '.txt')
SIZE_CHECK_TXT = r'\\fdm.se.telenor.net\FDM\PowerBI\ERP_Files\Oracle\_GL_export_check\GL_export_size_check.txt'
FDM_CALENDAR_PATH = r'\\fdm.se.telenor.net\FDM\PowerBI\Planning_files\FDM_Import_Calendar\FDM_Import_Calendar.xlsx'
GL_EXPORTS_FOLDER_PATH = r'\\fdm.se.telenor.net\FDM\PowerBI\ERP_Files\Oracle\O_GL'
PROCESS_STATUS_FILE_PATH = r'\\fdm.se.telenor.net\FDM\PowerBI\Process_statuses\ERP_export_file_checks.txt'
MAIL_RECIPIENTS = ["romualds.ivanovs@telenor.se",
                    "emerentia.nyholm@telenor.se"]

def write_log(log_file_path: str, message_type: str, message: str):
    try:
        with open(log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write(f"{datetime.now(pytz.timezone('Europe/Stockholm')).strftime('%Y-%m-%d %H:%M:%S')}\t{message_type}\t{message}\n")
    except Exception as e:
        print(f"Error writing to log file: {e}")

write_log(SCRIPT_LOG_FILE_PATH, "INFO", f"Script started.")

# ------ CHECKING Month closing calendar dates ------
try:
    calendar = pd.read_excel(FDM_CALENDAR_PATH, dtype='str').fillna('')
    write_log(SCRIPT_LOG_FILE_PATH, "INFO", f"Found calendar {FDM_CALENDAR_PATH}")
except:
    write_log(SCRIPT_LOG_FILE_PATH, "ERROR", f"Calendar not found / not accessible. {FDM_CALENDAR_PATH}")
    mail_status = send_mail(MAIL_RECIPIENTS, 'ERROR! FDM actuals not loaded | Calendar error', f'Check calendar availability {FDM_CALENDAR_PATH}', [SCRIPT_LOG_FILE_PATH])
    sys.exit(f"ERROR! Calendar not found / not accessible. {FDM_CALENDAR_PATH}")


if len(calendar.loc[calendar['Date']==today_swe_format]) != 1:
    print('No defined tasks for today, {today_swe_format}')
    write_log(SCRIPT_LOG_FILE_PATH, "INFO", f"No checks scheduled for today, {today_swe_format}")
    sys.exit()
else:
    day_type = calendar.loc[calendar['Date']==today_swe_format, "Date_type"].item()
    day_number = int(day_type[1:]) if day_type.startswith('D') else 32
    if day_number > 3:
        write_log(  PROCESS_STATUS_FILE_PATH,
                    'OK',
                    f'ERP export file size check')
        write_log(SCRIPT_LOG_FILE_PATH, "INFO", f"Today is {today_swe_format}, {day_type}. No ERP export checks scheduled. Status set to OK")

        sys.exit(f"Today is {today_swe_format}, {day_type}. No ERP export checks scheduled. Status set to OK")
    else:
        print(f'Today is {today_swe_format}, {day_type}. ERP export checks are scheduled')
        write_log(SCRIPT_LOG_FILE_PATH, "INFO", f"Today is {today_swe_format}, {day_type}. ERP export checks are scheduled")

gl_check_filename = f'GL_{closing_year_yyyy}{closing_month_mm}.txt'
gl_file_to_check = Path(GL_EXPORTS_FOLDER_PATH,
                        gl_check_filename)
print(gl_file_to_check)
write_log(SCRIPT_LOG_FILE_PATH,
          'INFO',
          f'GL file to check: {gl_file_to_check}')

if os.path.exists(gl_file_to_check) and os.path.getsize(gl_file_to_check) > 0:
    print(f"Found non-empty export file {gl_file_to_check}")
    current_filesize = os.path.getsize(gl_file_to_check)
else:
    print(f"Export file NOT found or is empty: {gl_file_to_check}")
    current_filesize = 0

# --- Get previous file size  ---

size_df = pd.read_csv(  SIZE_CHECK_TXT,
                        sep='\t',
                        dtype={'Timestamp': str, 'File': str, 'File_size': int}) # fields: Timestamp	File	File_size

latest_entry = size_df.loc[size_df['File'] == gl_check_filename].tail(1)
if not latest_entry.empty:
    previous_filesize = latest_entry['File_size'].values[0]
    previous_timestamp = latest_entry['Timestamp'].values[0]
    write_log(SCRIPT_LOG_FILE_PATH, 'INFO', f"Previous file size for {gl_check_filename}: {previous_filesize} bytes as of {previous_timestamp}")
else:
    previous_filesize = 0
    previous_timestamp = 'N/A'
    write_log(SCRIPT_LOG_FILE_PATH, 'INFO', f"No previous file size entry found for {gl_check_filename}. Assuming 0 bytes.")

write_log(SIZE_CHECK_TXT, gl_check_filename, current_filesize)

if current_filesize > previous_filesize:
    print(f"File size check PASSED for {gl_check_filename}. Current size: {current_filesize} bytes, Previous size: {previous_filesize} bytes as of {previous_timestamp}.")
    write_log(SCRIPT_LOG_FILE_PATH, 'INFO', f"File size check PASSED for {gl_check_filename}. Current size: {current_filesize} bytes, Previous size: {previous_filesize} bytes as of {previous_timestamp}.")
    write_log(  PROCESS_STATUS_FILE_PATH,
                'OK',
                f'ERP export file size check')
else:
    print(f"File size check FAILED for {gl_check_filename}. Current size: {current_filesize} bytes, Previous size: {previous_filesize} bytes as of {previous_timestamp}.")
    write_log(SCRIPT_LOG_FILE_PATH, 'ERROR', f"File size check FAILED for {gl_check_filename}. Current size: {current_filesize} bytes, Previous size: {previous_filesize} bytes as of {previous_timestamp}.")
    write_log(  PROCESS_STATUS_FILE_PATH,
                'Not OK',
                f'ERP export file size check')
    attachments = [SCRIPT_LOG_FILE_PATH, SIZE_CHECK_TXT]
    mail_status = send_mail(MAIL_RECIPIENTS, 
                            f'WARNING! Check ERP export file size - no additional data since last check | {gl_check_filename}', 
                            f'Check file size of {gl_check_filename}. Current size: {current_filesize} bytes, Previous size: {previous_filesize} bytes as of {previous_timestamp}.',
                            attachments)

# ----- REMOVING OLD LOG FILES ON FDM SERVER -----
from datetime import timedelta
cutoff_date_logs_on_fdm = datetime.now() - timedelta(days=65) # define cutoff date for log files on FDM server

for filename in os.listdir(SCRIPT_LOG_FOLDER_PATH):
    file_path = os.path.join(SCRIPT_LOG_FOLDER_PATH, filename)

    if os.path.isfile(file_path):
        try:
            modified_time = datetime.fromtimestamp(os.path.getmtime(file_path))

            if modified_time < cutoff_date_logs_on_fdm:
                os.remove(file_path)
                write_log(SCRIPT_LOG_FILE_PATH, "INFO", f"Deleted old log file on FDM server: {file_path} with modified time {modified_time}")
                print(f"Deleted: {file_path}")
        except Exception as e:
            print(f"Could not delete {file_path}: {e}")

# ----- keeping latest 200 rows in status file and size check csv-file -----

df = pd.read_csv(SIZE_CHECK_TXT, sep='\t', dtype={'Timestamp': str, 'File': str, 'File_size': int})
df.tail(200).to_csv(SIZE_CHECK_TXT, index=False)
write_log(SCRIPT_LOG_FILE_PATH, "INFO", f"Size check TXT trimmed to latest 200 rows.")

df = pd.read_csv(PROCESS_STATUS_FILE_PATH, sep='\t', dtype='str')
df.tail(200).to_csv(PROCESS_STATUS_FILE_PATH, index=False)
write_log(SCRIPT_LOG_FILE_PATH, "INFO", f"{PROCESS_STATUS_FILE_PATH} trimmed to latest 200 rows.")