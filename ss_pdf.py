import smartsheet
import fitz
import sys
import os
from io import BytesIO
from PIL import Image

SHEET_ID = 6253132649916292

#access token
key = os.environ.get('SMRT_API')

if key is None:
    print('Environment variable SMRT_API not found.')
    sys.exit(1)

def get_column_ids(id_):
    column_id_dict = {}
    for col in ss.Sheets.get_columns(id_).data:
        column_id_dict[col.title] = col.id
        column_id_dict[col.id] = col.title
    return column_id_dict

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
    
    files_to_upload = ['20221114_Friess_QC.pdf']
    results_files = []

    # for row in column.rows:

    #     col = row.get_column(sheet_columns_dict.get('Results File'))
    #     #remove all None values from the list
    #     if col.value is not None:
    #         results_files.append(col.value)
            
    # for file in file_names_corrected:
    #     if file not in results_files:
    #         files_to_upload.append(file)
              
    # # print(results_files)
    # print(files_to_upload)
    
    # if files_to_upload == []:
    #     exit('Exiting: No files to upload')

    return files_to_upload

#add rows to the sheet with files to upload
def row_addition(SHEET_ID, ss, sheet_columns_dict, column, file_dir, file_names_corrected, files_to_add):
    
    rows_updated = []
    
    for file in files_to_add(file_names_corrected, column):

        row = smartsheet.models.Row()
        row.to_top = True
        row.cells.append(smartsheet.models.Cell({'column_id': sheet_columns_dict.get('Results File'), 'value': file}))
        
        #copy the format from the first row
        row.format = column.rows[0].format
        #add the row to the sheet
        ss.Sheets.add_rows(SHEET_ID, [row])
        #get the row id of the row just added 
        row_id = ss.Sheets.get_sheet(SHEET_ID, column_ids = [sheet_columns_dict.get('Results File')]).rows[0].id
    
        #upload the pdf files to smartsheet row
        file_path = os.path.join(file_dir, 'pdf_files_'+ file)
        ss.Attachments.attach_file_to_row(SHEET_ID, row_id, (file, open(file_path, 'rb'),'application/pdf'))

        #add the row id to the list
        rows_updated.append(row_id)

    return rows_updated   

new_rows = row_addition(SHEET_ID, ss, sheet_columns_dict, column, file_dir, file_names_corrected, files_to_add)
print('New rows: ', new_rows)

def individual_image_process(img):

    d = pdf.extract_image(img[0])
    image_data = BytesIO(d['image'])
    image = Image.open(image_data)
    return image

    # with open(f'/Users/billteller/Desktop/{img[7]}.{d["ext"]}', 'wb') as imgout:
    #     imgout.write(d['image'])
    # return f'{img[7]}.{d["ext"]}'

#results dictionary
pdf_results = {}

# open the pdf files and extract the data using the row id
try:
    for row in new_rows:
        print('gathering from the attachment on row: ', row)
        #get the row id
        row_id = row
        #add the row id to the pdf_results dict
        pdf_results[row_id] = {}
        #get the row
        row = ss.Sheets.get_sheet(SHEET_ID, row_ids = [row_id])
        #get the attachment 
        attachment = ss.Attachments.list_row_attachments(SHEET_ID, row_id, include_all=True)
        #get the attachment name
        file_name = attachment.data[0].name
        #get the file path
        file_path = os.path.join(file_dir, 'pdf_files_'+ file_name)

        #open the pdf
        with fitz.open(file_path) as pdf:
            
            for i in range(len(pdf)):
                images = pdf[i].get_images()
                image_list = [individual_image_process(img) for img in images]
                #use info inside the bytes object to get the image name at index 7
                image_filenames = [img[7] for img in images]

            #get all pages from the pdf
            for pages in pdf:
                
                image_idx = 0
                sample_idx = 0
                #convert string to list
                page = pages.get_text().split('\n')

                for i, line in enumerate(page):
                    
                    #use index position
                    if 'Overall Results for' in line:
                        sample = page[i+1].strip()
                        if sample == 'RNA Area:':
                            sample = page[i-1].strip()
                        #update the pdf_results, row_id is the key to a new dict
                        pdf_results[row_id][sample] = {}
                        
                        #do image work
                        sample_jpg_list = image_list[sample_idx:sample_idx+2]
                        filename_index = image_filenames[sample_idx:sample_idx+2]

                        img1 = sample_jpg_list[0]
                        img1.save(f'/Users/billteller/Desktop/project_lee/ss_automation/image_results/{sample}_{filename_index[0]}.jpg')
                    
                        img2 = sample_jpg_list[1]
                        img2.save(f'/Users/billteller/Desktop/project_lee/ss_automation/image_results/{sample}_{filename_index[1]}.jpg')

                        pdf_results[row_id][sample]['image1'] = f'/Users/billteller/Desktop/project_lee/ss_automation/image_results/{sample}_{filename_index[0]}.jpg'
                        pdf_results[row_id][sample]['image2'] = f'/Users/billteller/Desktop/project_lee/ss_automation/image_results/{sample}_{filename_index[1]}.jpg'
                        
                        #merge the images   
                        images = [x for x in [img1, img2]]
                        widths, heights = zip(*(i.size for i in images))
                        total_width = sum(widths)
                        max_height = max(heights)
                        new_im = Image.new('RGB', (total_width, max_height))
                        x_offset = 0
                        for im in images:
                            new_im.paste(im, (x_offset,0))
                            x_offset += im.size[0]
                        #save the merged image
                        new_im.save(f'/Users/billteller/Desktop/project_lee/ss_automation/image_results/{sample}_merged.jpg')
                        #add the merged image to the pdf_results dict
                        pdf_results[row_id][sample]['merged_image'] = f'/Users/billteller/Desktop/project_lee/ss_automation/image_results/{sample}_merged.jpg'
                        
                        #increment counter by 2
                        sample_idx += 2
                    
                    if 'RNA Concentration' in line:
                        RNA_conc = page[i+1]
                        pdf_results[row_id][sample]['RNA Concentration'] = RNA_conc
                    
                    if 'RNA Integrity Number' in line:
                        RNA_integrity = page[i+1]
                        pdf_results[row_id][sample]['RNA Integrity Number (RIN)'] = RNA_integrity

                    #% of Total

                    if 'RNA Area' in line:
                        RNA_area = page[i+1]
                        pdf_results[row_id][sample]['RNA Area'] = RNA_area

                    if 'rRNA Ratio [28s / 18s]' in line:
                        rRNA_ratio = page[i+1]
                        pdf_results[row_id][sample]['rRNA Ratio [28s / 18s]'] = rRNA_ratio
                                        
                    if 'Result Flagging Label' in line:
                        result_flagging = page[i+1]
                        pdf_results[row_id][sample]['Result Flagging Label'] = result_flagging

                    # Corr. Area 1

                    if 'of total Area' in line:
                        Eighteen_s_total_area = page[i+5]
                        pdf_results[row_id][sample]['18S % of total Area'] = Eighteen_s_total_area
                        Twenty_eight_s_total_area = page[i+10]
                        pdf_results[row_id][sample]['28S % of total Area'] = Twenty_eight_s_total_area

except Exception as e:
    #make sure to catch any errors
    print('exception: ', e)
    #show which line the error occurred on
    print('line: ', sys.exc_info()[-1].tb_lineno)

# for key, value in pdf_results.items():
#     print(key)
#     for k, v in value.items():
#         print(k)
#         for k1, v1 in v.items():
#             print(k1, v1)

#add child rows to the sheet with the extracted data
def child_row_addition(SHEET_ID, ss, sheet_columns_dict, column, pdf_results, new_rows):

    # create a list to store the rows to add
    rows_to_add = []

    for row in new_rows:
        
        #if the row id matches the key in the pdf_results dict
        row_id = row

        for sample in pdf_results[row_id]:
            #create a row object
            row = smartsheet.models.Row()
            #set the parent id
            row.parent_id = row_id
            row.to_top = True
            #add the cells to the row        
            row.cells.append(smartsheet.models.Cell({'column_id': sheet_columns_dict.get('Sample'), 'value': sample}))
            row.cells.append(smartsheet.models.Cell({'column_id': sheet_columns_dict.get('RNA Concentration'), 'value': pdf_results[row_id][sample]['RNA Concentration']}))
            row.cells.append(smartsheet.models.Cell({'column_id': sheet_columns_dict.get('RNA Integrity Number (RIN)'), 'value': pdf_results[row_id][sample]['RNA Integrity Number (RIN)']}))
            row.cells.append(smartsheet.models.Cell({'column_id': sheet_columns_dict.get('RNA Area'), 'value': pdf_results[row_id][sample]['RNA Area']}))
            row.cells.append(smartsheet.models.Cell({'column_id': sheet_columns_dict.get('rRNA Ratio [28s / 18s]'), 'value': pdf_results[row_id][sample]['rRNA Ratio [28s / 18s]']}))
            row.cells.append(smartsheet.models.Cell({'column_id': sheet_columns_dict.get('Result Flagging Label'), 'value': pdf_results[row_id][sample]['Result Flagging Label']}))
            row.cells.append(smartsheet.models.Cell({'column_id': sheet_columns_dict.get('18S % of total Area'), 'value': pdf_results[row_id][sample]['18S % of total Area']}))
            row.cells.append(smartsheet.models.Cell({'column_id': sheet_columns_dict.get('28S % of total Area'), 'value': pdf_results[row_id][sample]['28S % of total Area']}))
            
            #copy the format from the first row
            row.format = column.rows[0].format
            
            #add the row to the list of rows to add
            rows_to_add.append(row)

    #add the row to the sheet
    ss.Sheets.add_rows(SHEET_ID, rows_to_add)
    return row_id

# #add the images to the child rows
def image_addition(SHEET_ID, ss, sheet_columns_dict, pdf_results, parent_id):
    
    #create a list to store the images to add
    found = []
    #look for rows that have the parent id of interest
    sheet = ss.Sheets.get_sheet(SHEET_ID)
    rows = sheet.rows
    
    sample_count = 0
    #index the samples in the pdf_results dict by the sample_count
    sample_list = list(pdf_results[parent_id].keys())  
    
    for row in rows:
        
        #if the row has the parent id of interest
        if row.parent_id == parent_id:
            found.append(row.id)
            print('found: ', found)
            
            for rows in found:
     
                #set the image parameters
                sheet_id = SHEET_ID
                column_id = sheet_columns_dict.get('Electropherogram')
                row_id = rows
                pic1 = pdf_results[parent_id][sample_list[sample_count]]['image1']
                pic2 = pdf_results[parent_id][sample_list[sample_count]]['image2']
                piclist = [pic1, pic2]
                merged_pic = pdf_results[parent_id][sample_list[sample_count]]['merged_image']
                file_type = "jpg"

            sample_count += 1 
            
            #add the images to the sheet
            ss.Cells.add_image_to_cell(sheet_id, row_id, column_id, merged_pic, file_type) 
            
            for pic in piclist:
                
                ss.Attachments.attach_file_to_row(SHEET_ID, row_id, (pic.split('/')[-1], open(pic, 'rb'),'image/jpg'))
                        
    
#add the child rows to the sheet
parent_id = child_row_addition(SHEET_ID, ss, sheet_columns_dict, column, pdf_results, new_rows)
print('parent_id: ', parent_id)

#add the images to the sheet
image_addition(SHEET_ID, ss, sheet_columns_dict, pdf_results, parent_id)