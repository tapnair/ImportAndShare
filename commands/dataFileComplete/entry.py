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

    # Create the event handler for when data files are complete.
    futil.add_handler(app.dataFileComplete, application_data_file_complete, local_handlers=local_handlers)

    # Create a csv file for the results.
    create_csv()


# Executed when add-in is stopped.
def stop():
    pass


# Function to be executed by the dataFileComplete event.
def application_data_file_complete(args: adsk.core.DataEventArgs):
    futil.log(f'In application_data_file_complete event handler for: {args.file.name}')

    # Make sure we are processing a file imported from this script
    if args.file.name in config.imported_filenames:
        data_file: adsk.core.DataFile = args.file

        # Create the public link for the data file
        public_link = data_file.publicLink

        # Write results to the output csv file
        with open(config.csv_file_name, mode='a') as csv_file:
            fieldnames = ['Name', 'URN', 'Link']
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writerow({
                'Name': data_file.name,
                'URN': data_file.versionId,
                'Link': public_link
            })

        # remove this from the list and close
        config.imported_filenames.remove(args.file.name)
        this_document = config.imported_documents.pop(args.file.name)
        this_document.close(False)


# Initial creation of csv for results
def create_csv():
    with open(config.csv_file_name, mode='w') as csv_file:
        fieldnames = ['Name', 'URN', 'Link']
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()


