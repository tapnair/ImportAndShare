# Author Patrick Rainsberry
# Description Sample Script to import and share with event handlers
import adsk.core
import adsk.fusion
import traceback
import csv
import os

# Output csv file to record results.
csv_file_name = os.path.join(os.path.dirname(__file__), 'output.csv')

# Extension types that will be processed for import
EXTENSION_TYPES = ['.step', '.stp']

# Keep track of imported files
imported_filenames = []

# Flag to determine when it is safe to terminate the script
DONE_IMPORTING = False

# Local list of event handlers used to maintain a reference so
# they are not released and garbage collected.
local_handlers = []

# Global references to main Application and User Interface objects.
app = adsk.core.Application.get()
ui = app.userInterface


def run(context):
    global DONE_IMPORTING
    try:

        # Have User select a folder containing STEP Files to be imported
        folder_dialog = ui.createFolderDialog()
        folder_dialog.title = "Select Folder"
        dialog_result = folder_dialog.showDialog()
        if dialog_result == adsk.core.DialogResults.DialogOK:
            folder = folder_dialog.folder
        else:
            return

        # For this script we are simply choosing the root folder of the currently active project
        # You can see this in the in the data panel
        # Here you could do something more complete to determine the proper location
        try:
            target_data_folder = app.data.activeProject.rootFolder
        except:
            ui.messageBox("You probably are navigated to a project in the data panel. "
                          "For example, can't be in recent documents.")
            return

        # Create a csv file for results
        with open(csv_file_name, mode='w') as csv_file:
            fieldnames = ['Name', 'URN', 'Link']
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writeheader()

        # Setup Event handler to process files after initial save is complete
        onDataFileComplete = MyDataFileCompleteHandler()
        app.dataFileComplete.add(onDataFileComplete)
        local_handlers.append(onDataFileComplete)

        # Iterate over all STEP files in user selected directory
        for full_file_name in os.listdir(folder):
            file_path = os.path.join(folder, full_file_name)
            if os.path.isfile(file_path):
                split_tup = os.path.splitext(full_file_name)
                file_name = split_tup[0]
                file_extension = split_tup[1]
                if file_extension in EXTENSION_TYPES:

                    # Execute the Fusion 360 import into a new document
                    import_manager = app.importManager
                    step_options = import_manager.createSTEPImportOptions(file_path)
                    new_document = import_manager.importToNewDocument(step_options)

                    # Save the new document and record name
                    new_document.saveAs(file_name, target_data_folder, 'Imported from script', 'tag')
                    imported_filenames.append(file_name)

        # We have imported all of the files
        DONE_IMPORTING = True

        # Normally a script will terminate upon completion.
        # In this case we want it to stay active until all data files are processed.
        adsk.autoTerminate(False)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# We create a subclass to handle data events
# This will be called when the data files are processed and ready on the cloud.
class MyDataFileCompleteHandler(adsk.core.DataEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.DataEventArgs):

        # Make sure we are processing a file imported from this script
        if args.filename in imported_filenames:
            data_file: adsk.core.DataFile = args.file

            # In this case it is really just "activating" the window.
            document = app.documents.open(data_file, True)

            # Create the public link for the data file
            public_link = data_file.publicLink

            # Write results to the output csv file
            with open(csv_file_name, mode='a') as csv_file:
                fieldnames = ['Name', 'URN', 'Link']
                writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                writer.writerow({
                    'Name': data_file.name,
                    'URN': data_file.versionId,
                    'Link': public_link
                })

            # remove this from the list and close
            imported_filenames.remove(args.filename)
            document.close(False)

        # Terminate the script after the last file is processed
        if len(imported_filenames) == 0 and DONE_IMPORTING:
            adsk.terminate()


