import adsk.core
from ...lib import fusion360utils as futil
from ... import config
import csv


app = adsk.core.Application.get()
ui = app.userInterface


# Local list of event handlers used to maintain a reference so
# they are not released and garbage collected.
local_handlers = []


# Executed when add-in is run.
def start():
    futil.add_handler(app.dataFileComplete, application_data_file_complete, local_handlers=local_handlers)
    create_csv()


# Executed when add-in is stopped.
def stop():
    pass


# Event handler for the dataFileComplete event.
def application_data_file_complete(args: adsk.core.DataEventArgs):
    # Code to react to the event.
    futil.log(f'In application_data_file_complete event handler for: {args.file.name}')

    if args.file.name in config.imported_filenames:
        data_file: adsk.core.DataFile = args.file
        document = app.documents.open(data_file)
        public_link = data_file.publicLink
        with open(config.csv_file_name, mode='a') as csv_file:
            fieldnames = ['Name', 'URN', 'Link']
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writerow({
                'Name': data_file.name,
                'URN': data_file.versionId,
                'Link': public_link
            })
        config.imported_filenames.remove(args.file.name)
        document.close(False)


def create_csv():
    with open(config.csv_file_name, mode='w') as csv_file:
        fieldnames = ['Name', 'URN', 'Link']
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()


