import csv
import json

import adsk.core

from ... import config
from ...lib import fusion360utils as futil

app = adsk.core.Application.get()
ui = app.userInterface

NAME1 = 'Data_Handler'
NAME2 = "Custom Import Event"
NAME3 = "Custom Save Event"
NAME4 = "Custom Close Event"
# Local list of event handlers used to maintain a reference so
# they are not released and garbage collected.
local_handlers = []
my_data_handlers = []
my_custom_handlers = []


# Executed when add-in is run.  Create custom events so we don't disrupt the main application loop.
def start():
    app.unregisterCustomEvent(config.custom_event_id_import)
    custom_event_import = app.registerCustomEvent(config.custom_event_id_import)
    custom_event_handler_import = futil.add_handler(custom_event_import, custom_event_import_execute, name=NAME2)
    my_custom_handlers.append({
        'custom_event_id': config.custom_event_id_import,
        'custom_event': custom_event_import,
        'custom_event_handler': custom_event_handler_import
    })

    app.unregisterCustomEvent(config.custom_event_id_save)
    custom_event_save = app.registerCustomEvent(config.custom_event_id_save)
    custom_event_handler_save = futil.add_handler(custom_event_save, custom_event_save_execute, name=NAME3)
    my_custom_handlers.append({
        'custom_event_id': config.custom_event_id_save,
        'custom_event': custom_event_save,
        'custom_event_handler': custom_event_handler_save
    })

    app.unregisterCustomEvent(config.custom_event_id_close)
    custom_event_close = app.registerCustomEvent(config.custom_event_id_close)
    custom_event_handler_close = futil.add_handler(custom_event_close, custom_event_close_execute, name=NAME4)
    my_custom_handlers.append({
        'custom_event_id': config.custom_event_id_close,
        'custom_event': custom_event_close,
        'custom_event_handler': custom_event_handler_close
    })

    # Create the event handler for when data files are complete.
    my_data_handlers.append(
        futil.add_handler(app.dataFileComplete, application_data_file_complete, local_handlers=local_handlers,
                          name=NAME1))


# Executed when add-in is stopped.  Remove events.
def stop():
    for custom_item in my_custom_handlers:
        custom_item['custom_event'].remove(custom_item['custom_event_handler'])
        app.unregisterCustomEvent(custom_item['custom_event_id'])

    for data_handler in my_data_handlers:
        app.dataFileComplete.remove(data_handler)


# Import a document from the list
def custom_event_import_execute(args: adsk.core.CustomEventArgs):
    event_data = json.loads(args.additionalInfo)
    file_name = event_data['file_name']
    file_path = event_data['file_path']

    futil.log(f'**********Importing: {file_name}')

    # Execute the Fusion 360 import into a new document
    import_manager = app.importManager
    step_options = import_manager.createSTEPImportOptions(file_path)
    new_document = import_manager.importToNewDocument(step_options)

    # Keep track of imported files
    config.imported_documents[file_name] = new_document
    config.imported_filenames.append(file_name)

    # Fire event to save the document
    event_data = {
        'file_name': file_name,
        'file_path': file_path
    }
    additional_info = json.dumps(event_data)
    app.fireCustomEvent(config.custom_event_id_save, additional_info)


# Save a specific Document
def custom_event_save_execute(args: adsk.core.CustomEventArgs):
    event_data = json.loads(args.additionalInfo)
    file_name = event_data['file_name']

    futil.log(f'**********Saving: {file_name}')

    new_document = config.imported_documents[file_name]
    new_document.saveAs(file_name, config.target_data_folder, 'Imported from script', 'tag')


# Close a specific document
def custom_event_close_execute(args: adsk.core.CustomEventArgs):
    event_data = json.loads(args.additionalInfo)
    file_name = event_data['file_name']

    futil.log(f'**********Closing: {file_name}')

    new_document = config.imported_documents[file_name]
    new_document.close(False)


# Function to be executed by the dataFileComplete event.
def application_data_file_complete(args: adsk.core.DataEventArgs):
    futil.log(f'***In application_data_file_complete event handler for: {args.file.name}')

    # Get the dataFile and process it
    data_file: adsk.core.DataFile = args.file
    process_data_file(data_file)


def process_data_file(data_file: adsk.core.DataFile):
    # Make sure we are processing a file imported from this script
    if data_file.name in config.imported_filenames:

        # Create the public link for the data file
        public_link = data_file.publicLink
        futil.log(f"**********Created public link: {public_link}")

        # Store the result of this file
        config.results.append({
            'Name': data_file.name,
            'URN': data_file.versionId,
            'Link': public_link
        })
        config.imported_filenames.remove(data_file.name)

        # Fire close event for this Document
        event_data = {
            'file_name': data_file.name,
        }
        additional_info = json.dumps(event_data)
        app.fireCustomEvent(config.custom_event_id_close, additional_info)

        # If all documents have been processed finalize results
        if len(config.imported_filenames) == 0:
            if not config.run_finished:
                config.run_finished = True
                finished()
    else:
        futil.log(f"**********Already processed: {data_file.name}")


# After all files are processed write the results
def finished():
    futil.log(f"Writing CSV")
    with open(config.csv_file_name, mode='w') as csv_file:
        fieldnames = ['Name', 'URN', 'Link']
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        for row in config.results:
            writer.writerow(row)
