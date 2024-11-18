import os
import globalvars

def get_path(directory):
    folder_names = []
    if globalvars.settings_batch == 0:
        for entry in os.scandir(directory):
            if entry.is_dir():
                folder_names.append(entry.name)
        return folder_names