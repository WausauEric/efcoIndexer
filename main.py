import os
import pandas as pd

# Define the folder we're starting work in, which probably won't be where this program lives.
rootFolderPath = '\\\\srvnas01.efco.com\\CSG\\Engineering\\Structural\\STRL Projects'

# Switch the active directory to that folder, using the 'os' package
os.chdir(rootFolderPath)

# Set up the dataframe we'll save all the info into
data = []

# Scan the folder to grab all files and folders, then check which ones are folder & have names that start with "20"
for entry in os.scandir():
    if entry.is_dir and entry.name[:2] == '20':
        if entry.name[:2] == '20':
            # Define the project year variable, then loop through that folder
            yearProjectStart = entry.name[:4]
            for projectFolder in os.scandir(entry.path):
                folderName = projectFolder.name

                # Discard any folder names shorter than 11 characters
                if len(folderName) >= 11:
                    # Remove dashes from project numbers
                    if folderName[1] == '-':
                        folderName = folderName[:1] + folderName[2:]
                    if folderName[8] == '-':
                        folderName = folderName[:8] + folderName[10:]

                    # Attempt to remove any random folders from the list by only keeping those with space for 7th char
                    if folderName[7] == ' ':
                        # Split out the project name and number from the folder name
                        projectNumber = folderName[:6]
                        projectName = folderName[8:]
                        projectPath = rootFolderPath + projectFolder.path[1:]

                        # Recast the project path as a hyperlink for Excel
                        projectHyperlink = '=HYPERLINK("' + projectPath + '", "' + projectPath + '")'

                        # Append the parsed-out project info to the pre-allocated numpy array
                        data.append([yearProjectStart, projectNumber, projectName, projectHyperlink])

# Convert our array to a Pandas DataFrame
df = pd.DataFrame(data, columns=['FY Start', 'Project Number', 'Project Name', 'Folder Link'])

# Reverse the row index to make newer projects show first
data.reverse()
rowData = len(data) + 1

# Write DataFrame to Excel
writer = pd.ExcelWriter('ProjectIndex.xlsx')
df.to_excel(writer, sheet_name='Index', index=False)

# Auto-adjust columns' width
for column in df:
    column_width = max(df[column].astype(str).map(len).max(), len(column))
    col_idx = df.columns.get_loc(column)
    writer.sheets['Index'].set_column(col_idx, col_idx, column_width + 2)

# Format as Table from the start for search / filter functions
writer.sheets['Index'].add_table('A1:D' + str(rowData), {'data': data,
                                                         'columns': [{'header': 'FY Start'},
                                                                     {'header': 'Project Number'},
                                                                     {'header': 'Project Name'},
                                                                     {'header': 'Folder Link'}
                                                                     ]})
# Set the zoom to 130% so we can see
writer.sheets['Index'].set_zoom(130)

# Publish the .xlsx file
writer.save()
