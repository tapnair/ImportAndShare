# Author Patrick Rainsberry
# Description Sample Script to import STEP files from a directory
import adsk.core
import adsk.fusion
import traceback
import os

# Extension types that will be processed for import
EXTENSION_TYPES = ['.step', '.stp']

# Global references to main Application and User Interface objects.
app = adsk.core.Application.get()
ui = app.userInterface


def run(context):
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

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
