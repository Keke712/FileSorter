
import os, sys
import shutil
import time
from datetime import datetime
import win32file, win32con, pywintypes
from tqdm import tqdm

choice = input("Sort by:\n1. Creation date\n2. Modification date (recommended)\n: ")
choice2 = input("\nBy day or month (D/M) ?")
choice3 = input("\nSave old files (Y/N): ")

directory = os.getcwd()

# Directory where files are copied and sorted in folders
rollback_folder = directory+"\\Rollback"

# Check if dir is created else: create it
def check_dir(dr):
    if not os.path.exists(dr):
        os.mkdir(dr)

def main():
    total_files = sum(len(files) for _, _, files in os.walk(directory))

    with tqdm(total=total_files, desc="Progression") as pbar:
        for root, dirs, files in os.walk(directory):
            for name in files:
                # For all files ending by..
                if name.lower().endswith(('.png', '.jpg', '.jpeg', '.mp4')):
                    try:
                        created_time = os.path.getctime(name)
                    except Exception:
                        print("\nFolders cannot be moved")
                        sys.exit()

                    modified_time = os.path.getmtime(name)

                    # Format file date
                    # Choice with the first input

                    if choice2 == "M":
                        if choice == "1":
                            month_year = datetime.fromtimestamp(created_time).strftime('%Y %B')
                        elif choice == "2":
                            month_year = datetime.fromtimestamp(modified_time).strftime('%Y %B')
                        else:
                            print("\nSort option is not valid")
                            sys.exit()
                        
                        path_to_save = directory+"\\"+month_year
                        
                        check_dir(path_to_save)
                        file_to_save = path_to_save+"\\"+name
                    elif choice2 == "D":
                        if choice == "1":
                            month_day = datetime.fromtimestamp(created_time).strftime('%B %d')
                        elif choice == "2":
                            month_day = datetime.fromtimestamp(modified_time).strftime('%B %d')
                        else:
                            print("\nSort option is not valid")
                            sys.exit()
                        
                        path_to_save = directory+"\\"+month_day
                        
                        check_dir(path_to_save)
                        file_to_save = path_to_save+"\\"+name
                    else:
                        print("\nDate option is not valid")

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
                            # /!\ Modify the code here if you need
                            # Setting the creation time on modified time for precision
                            # Example: Copying on photos from phone to your PC will replace the creationTime with the current day

                            CreationTime=modified_time_obj,
                            LastWriteTime=modified_time_obj
                        )

                        # Close the file handle
                        win32file.CloseHandle(handle)

                        if choice3 == "Y":
                            # Y: Move old file in rollback folder
                            check_dir(rollback_folder)
                            shutil.move(name, rollback_folder)
                        elif choice3 == "N":
                            # N: Delete files
                            os.remove(name)
                        else:
                            print("\nSave old file option is not valid")
                            sys.exit()

                        # Progression bar updated
                        pbar.update(1)

if __name__ == "__main__":
    main()
