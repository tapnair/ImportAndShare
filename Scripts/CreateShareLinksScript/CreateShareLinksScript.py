# Author Patrick Rainsberry
# Description Sample Script to import STEP files from a directory
import adsk.core
import adsk.fusion
import traceback
import csv
import os

# Output csv file to record results.
csv_file_name = os.path.join(os.path.dirname(__file__), 'output.csv')

# Global references to main Application and User Interface objects.
app = adsk.core.Application.get()
ui = app.userInterface


def run(context):
    try:

        # Create a csv file for results
        with open(csv_file_name, mode='w') as csv_file:
            fieldnames = ['Name', 'URN', 'Link']
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writeheader()

        # We don't want to mess up the documents collection by closing them while we are still iterating
        documents_to_close = []

        # Iterate over all open documents
        document: adsk.core.Document
        for document in app.documents:
            document.activate()

            # Check if this document is saved and assume it is one we imported.
            if document.isSaved:

                # Get the data file (Cloud representation) and check that it is done processing and available
                data_file = document.dataFile
                if data_file.isComplete:

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

                    # Add this to list of documents to close
                    documents_to_close.append(document)

        # Close the documents
        for document in documents_to_close:
            document.close(False)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
