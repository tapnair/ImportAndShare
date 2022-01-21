# Author Patrick Rainsberry
# Description Sample Script

import adsk.core
import adsk.fusion
import traceback
import csv
import os

csv_file_name = os.path.join(os.path.dirname(__file__), 'output.csv')
EXTENSION_TYPES = ['.step', '.stp']
imported_filenames = []
results = []

# Local list of event handlers used to maintain a reference so
# they are not released and garbage collected.
local_handlers = []

app = adsk.core.Application.get()
ui = app.userInterface


def run(context):

    try:
        onDataFileComplete = MyDataFileCompleteHandler()
        app.dataFileComplete.add(onDataFileComplete)
        local_handlers.append(onDataFileComplete)

        folder_dialog = ui.createFolderDialog()
        folder_dialog.title = "Select Folder"
        dialog_result = folder_dialog.showDialog()
        if dialog_result == adsk.core.DialogResults.DialogOK:
            folder = folder_dialog.folder
        else:
            return

        with open(csv_file_name, mode='w') as csv_file:
            fieldnames = ['Name', 'URN', 'Link']
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writeheader()

        target_data_folder = app.data.activeProject.rootFolder
        for full_file_name in os.listdir(folder):
            file_path = os.path.join(folder, full_file_name)
            # checking if it is a file
            if os.path.isfile(file_path):
                split_tup = os.path.splitext(full_file_name)
                file_name = split_tup[0]
                file_extension = split_tup[1]
                if file_extension in EXTENSION_TYPES:
                    import_manager = app.importManager
                    step_options = import_manager.createSTEPImportOptions(file_path)
                    new_document = import_manager.importToNewDocument(step_options)
                    new_document.saveAs(file_name, target_data_folder, 'Imported from script', 'tag')
                    imported_filenames.append(new_document.name)
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


class MyDataFileCompleteHandler(adsk.core.DataEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.DataEventArgs):
        if args.filename in imported_filenames:
            data_file: adsk.core.DataFile = args.file
            document = app.documents.open(data_file)
            public_link = data_file.publicLink
            with open(csv_file_name, mode='a') as csv_file:
                fieldnames = ['Name', 'URN', 'link']
                writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                writer.writerow({
                    'Name': data_file.name,
                    'URN': data_file.versionId,
                    'Link': public_link
                })
            imported_filenames.remove(args.filename)
            document.close(False)




