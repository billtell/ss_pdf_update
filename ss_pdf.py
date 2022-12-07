import smartsheet
import fitz
import sys
import os
import requests
import io
import time

def get_column_ids(id_):
    column_id_dict = {}
    for col in ss.Sheets.get_columns(id_).data:
        column_id_dict[col.title] = col.id
        column_id_dict[col.id] = col.title
    return column_id_dict

SHEET_ID = 6253132649916292

#access token
key = os.environ.get('SMRT_API')

if key is None:
    print('Environment variable SMRT_API not found.')
    sys.exit(1)

# Initialize client
ss = smartsheet.Smartsheet(key)
# Get the sheet
sheet = ss.Sheets.get_sheet(SHEET_ID)
# Get the column ids
sheet_columns_dict = get_column_ids(SHEET_ID)
# print(sheet_columns_dict)
#get the column 
column = ss.Sheets.get_sheet(SHEET_ID, column_ids = [sheet_columns_dict.get('Results File')])


#set the file directory
file_dir = '/Users/billteller/Desktop/project_lee/ss_automation/bioanalyzer_pdf_files'
#get the file names from file_dir
file_names = os.listdir(file_dir)
#remove .DS_Store from the list
file_names.remove('.DS_Store')
#get the full paths from file_dir
files_fullpath = [os.path.join(file_dir, file) for file in file_names]
#file name corrected
file_names_corrected = [file.replace('pdf_files_', '') for file in file_names]
# print(file_names_corrected)

#does the file exist in the smartsheet already?
def files_to_add(file_names_corrected, column):
    
    files_to_upload = []
    results_files = []

    for row in column.rows:

        col = row.get_column(sheet_columns_dict.get('Results File'))
        #remove all None values from the list
        if col.value is not None:
            results_files.append(col.value)
            
    for file in file_names_corrected:
        if file not in results_files:
            files_to_upload.append(file)
              
    # print(results_files)
    print(files_to_upload)
    
    if files_to_upload == []:
        exit('Exiting: No files to upload')

    return files_to_upload

timestamp = time.strftime('%m-%d-%Y %H:%M %p')

#add rows to the sheet with files to upload
def row_addition(SHEET_ID, ss, sheet_columns_dict, column, file_dir, file_names_corrected, files_to_add):
    
    rows_updated = []
    
    for file in files_to_add(file_names_corrected, column):

        row = smartsheet.models.Row()
        row.to_top = True
        row.cells.append(smartsheet.models.Cell({'column_id': sheet_columns_dict.get('Results File'), 'value': file}))
        # row.cells.append(smartsheet.models.Cell({'column_id': sheet_columns_dict.get('Created'), 'value': timestamp}))
        #copy the format from the first row
        row.format = column.rows[0].format
        #add the row to the sheet
        ss.Sheets.add_rows(SHEET_ID, [row])
        #get the row id of the row just added 
        row_id = ss.Sheets.get_sheet(SHEET_ID, column_ids = [sheet_columns_dict.get('Results File')]).rows[0].id
    
        #upload the pdf files to smartsheet row
        file_path = os.path.join(file_dir, 'pdf_files_'+ file)
        add_pdf = ss.Attachments.attach_file_to_row(SHEET_ID,row_id, (file, open(file_path, 'rb'),'application/pdf'))

        #add the row id to the list
        rows_updated.append(row_id)

    print('Row updated: ', rows_updated)

    return rows_updated   

new_rows = row_addition(SHEET_ID, ss, sheet_columns_dict, column, file_dir, file_names_corrected, files_to_add)
print('New rows: ', new_rows)

# open the pdf files and extract the data using the row id
try:
    for row in new_rows:
        print('row =', row)
        #get the row id
        row_id = row
        #get the row
        row = ss.Sheets.get_sheet(SHEET_ID, row_ids = [row_id])
        #get the attachment 
        attachment = ss.Attachments.list_row_attachments(SHEET_ID, row_id, include_all=True)
        #get the attachment name
        file_name = attachment.data[0].name
        #get the file path
        file_path = os.path.join(file_dir, 'pdf_files_'+ file_name)
        #open the pdf
        pdf = fitz.open(file_path)
        #get all pages from the pdf
        for page in pdf:
            pages = page.get_text("text")
            print(pages) 
        #close the pdf
        pdf.close()
    print()
    print('hello!')

except Exception as e:
    print('exception: ', e)
  
#function to grab metrics from pages 
def get_metrics(pages):
#     #get the metrics from the pages
#     #grab 11 fields

  
        
# create child rows to the row id and add the metrics to the child rows
def child_rows():
    



















#get the attachment if it were uploaded to the row 
#not part of the project.  need to upload pdf file to row.  need to add to top 

# #get the attachment id from the row
# response = ss.Attachments.list_row_attachments(SHEET_ID, row_id, include_all=True)
# attachments = response.data
# #get the attachment id
# attachment_id = attachments[0].id
# #get the attachment name
# attachment_name = attachments[0].name.replace('pdf_files_', '')
# #get the attachment
# get_attachment = ss.Attachments.get_attachment(
#         SHEET_ID,       
#         attachment_id)          
# #get temp url from get_attachment
# temp_url = get_attachment.url
# #open the attachment
# r = requests.get(temp_url)
# #open the pdf
# filestream = io.BytesIO(r.content)
# pdf = fitz.open(stream=filestream, filetype="pdf")
#get the text from the pdf
)
