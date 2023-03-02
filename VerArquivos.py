import os
import xlsxwriter
from moviepy.video.io.VideoFileClip import VideoFileClip
from datetime import timedelta


errados = ['Apps','Jobs']

# Define the root folder where you want to start searching for .mp4 files
root_folder = 'E:\\BACKUP\\Lucas'

# Define the output Excel file where you want to store the information
output_file = 'C:\\Users\\Henrique\\Desktop\\file.xlsx'

# Create a new Excel workbook and worksheet
workbook = xlsxwriter.Workbook(output_file)
worksheet = workbook.add_worksheet()

# Write the headers to the worksheet
worksheet.write('A1', 'File Name')
worksheet.write('B1', 'Folder Name')
worksheet.write('C1', 'Duration')

# Define the row number to start writing the data
row = 2

# Loop through all folders and subfolders under the root folder
for root, dirs, files in os.walk(root_folder):
    # Skip folders ending with '1'
    if os.path.basename(root).endswith('1') or 'Apps' in root or 'Jobs' in root:
        print('\033[91m' + 'Pasta Pulada: ' + os.path.basename(root) + '\033[0m')

        continue
    else:
        print('\033[92m' + 'Pasta Lida: ' + os.path.basename(root) + '\033[0m')

        
    
    
     # Loop through all files in the current folder
    for file in files:
        # Check if the file is an .mp4 file
        if file.endswith('.mp4'):
            # Construct the output values with the file name and folder name
            file_name = file
            folder_name = os.path.basename(root)
            file_path = os.path.join(root, file)
            
            # Get the duration of the video
            video_path = os.path.join(root, file)
            video = VideoFileClip(video_path)
            duration = video.duration
            hours, remainder = divmod(duration, 3600)
            minutes, seconds = divmod(remainder, 60)
            if hours == 0:
                duration_str = f'{int(minutes):02d}m {int(seconds):02d}s'
            else:
                duration_str = f'{int(hours):02d}h {int(minutes):02d}m {int(seconds):02d}s'
            video.close()
            # Format the duration as hours, minutes, and seconds
            
            
            # Check if the folder ends with '2'
            if folder_name.endswith('2'):
                # Write the values to the Excel file with red font
                worksheet.write(f'A{row}', file_name, workbook.add_format({'font_color': 'red'}))
                worksheet.write(f'B{row}', folder_name, workbook.add_format({'font_color': 'red'}))
                worksheet.write(f'C{row}', duration_str, workbook.add_format({'font_color': 'red'}))
            else:
                # Write the values to the Excel file
                worksheet.write(f'A{row}', file_name)
                worksheet.write(f'B{row}', folder_name)
                worksheet.write(f'C{row}', duration_str)
            
            # Increment the row number for the next data
            row += 1

# Close the Excel workbook
print('Feito')
workbook.close()