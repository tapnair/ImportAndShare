# Application Global Variables
# This module serves as a way to share variables across different
# modules (global variables).

import os

# Flag that indicates to run in Debug mode or not. When running in Debug mode
# more information is written to the Text Command window. Generally, it's useful
# to set this to True while developing an add-in and set it to False when you
# are ready to distribute it.
DEBUG = True

# Gets the name of the add-in from the name of the folder the py file is in.
# This is used when defining unique internal names for various UI elements 
# that need a unique name. It's also recommended to use a company name as 
# part of the ID to better ensure the ID is unique.
ADDIN_NAME = os.path.basename(os.path.dirname(__file__))
COMPANY_NAME = 'Autodesk'


# *********** Global Variables Unique to this Add-in **************

# Keep track of imported files
imported_filenames = []

# Output csv file to record results.
csv_file_name = os.path.join(os.path.dirname(__file__), 'output.csv')

# Extension types that will be processed for import
EXTENSION_TYPES = ['.step', '.stp']

# Keep track of imported files
imported_documents = {}

custom_event_id_import = 'custom_event_import'
custom_event_id_save = 'custom_event_id_save'
custom_event_id_close = 'custom_event_id_close'

target_data_folder = None

results = []
run_finished = False

