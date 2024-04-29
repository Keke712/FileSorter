
import os
import shutil
import time
from datetime import datetime
import win32file, win32con, pywintypes

directory = os.getcwd()
# Directory where files are copied and sorted in folders
save_to = directory+"\\Sorted"

# Check if the directory to sort images is created
def check_dir(dr):
    if not os.path.exists(dr):
        os.mkdir(dr)

check_dir(save_to)

def main():
    for root, dirs, files in os.walk(directory):
        for name in files:
            # For all files ending by..
            if name.lower().endswith(('.png', '.jpg', '.jpeg', '.mp4')):
                created_time = os.path.getctime(name)
                modified_time = os.path.getmtime(name)

                # Format file date
                month_year = datetime.fromtimestamp(created_time).strftime('%Y %B')
                path_to_save = save_to+"\\"+month_year


                check_dir(path_to_save)
                file_to_save = path_to_save+"\\"+name

                if not os.path.exists(file_to_save):
                    # Copy file with metadata
                    shutil.copy2(name, file_to_save)

                    # Initialize Win32 objects
                    creation_time_obj = pywintypes.Time(created_time)
                    modified_time_obj = pywintypes.Time(modified_time)

                    # Get a handle to the copied file (open)
                    handle = win32file.CreateFile(
                        file_to_save,
                        win32con.GENERIC_WRITE,
                        0,
                        None,
                        win32con.OPEN_EXISTING,
                        0,
                        None
                    )

                    # Set the creation time of the file to the modification time of the source file
                    win32file.SetFileTime(
                        handle,
                        CreationTime=creation_time_obj,
                        LastWriteTime=modified_time_obj
                    )

                    # Close the file handle
                    win32file.CloseHandle(handle)

if __name__ == "__main__":
    main()
