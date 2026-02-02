import os, shutil, tempfile, sys, subprocess, ctypes, win32com.client # type: ignore

def ask(phrase):
    if confirmation:
        return True

    question = input(phrase).lower()
    return question != 'n'


def run_all():
    question = input("Run each recommended script? [Y/n] ")
    return True if question.lower() == 'y' else False

def check_admin_privileges():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def request_admin():
    try:
        result = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
        return result > 32
    except Exception as e:
        print("Error: ", str(e))
        return False

def check_os():
    if os.name != 'nt':
        sys.exit(f"{"-" * 50}\nNot running Windows\n{"-" * 50}")

def get_temp_path():
    return tempfile.gettempdir()

def unit(size):
    if size < 1000:
        return f"{size} bytes"
    if size >= 1000 and size < 10 ** 6:
        return f"{round(size/1000)} KB"
    if size >= 10 ** 6 and size < 10 ** 9:
        return f"{round(size/(10 ** 6), 1)} MB"
    else:
        return f"{round(size/(10 ** 9), 2)} GB"

def clear_temp(PATH):
    if not ask("Clear temporary files? [Y/n] "):
        return

    deleted_files = 0
    deleted_directories = 0
    deleted_size = 0
    failed = 0

    with os.scandir(PATH) as temp:
        for item in temp:
            file_path = item.path
            try:
                if item.is_dir():
                    size = os.stat(file_path).st_size
                    shutil.rmtree(file_path)
                    print(f"Deleted directory: {file_path}")
                    deleted_directories += 1
                    deleted_size += size
                else:
                    size = os.stat(file_path).st_size
                    os.remove(file_path)
                    print(f"Deleted file: {file_path}")
                    deleted_files += 1
                    deleted_size += size
            except Exception as e:
                print(f"Couldn't delete {file_path}. Error: {e}")
                failed += 1
    
    print(f"{"-" * 50}"
          f"\nDeleted {deleted_files} {"files" if deleted_files != 1 else "file"} and {deleted_directories} {"directories" if deleted_directories != 1 else "directory"}"
          f"\nFailed to delete {failed} items"
          f"\nFreed up {unit(deleted_size)} of storage\n"
          f"{"-" * 50}")
    
    if not confirmation:
        os.system('pause')

def run_sfc():
    if not ask("Run search for corrupted? [Y/n] "):
        return

    print("\nProceeding to fix corrupted files.")
    try:
        subprocess.run("sfc /scannow", check=False)
        if not confirmation:
            os.system('pause')
    except Exception as e:
        print("Something went wrong. Error: {0}".format(e))

def dism():
    if not ask("Run DISM tool? [Y/n] "):
        return
    
    print("\nProceeding to start the Deployment Image Servicing and Management tool")
    try:
        subprocess.run("DISM /Online /Cleanup-Image /RestoreHealth", check=False)
    except Exception as e:
        print("Something went wrong. Error: {0}".format(e))

def check_windows_updates():
    if not ask("Check for Windows Updates? [Y/n] "):
        return
        
    print("\nChecking for Windows Updates...")
    session = win32com.client.Dispatch("Microsoft.Update.Session")
    searcher = session.CreateUpdateSearcher()

    result = searcher.Search("IsInstalled=0 and Type='Software'")

    if result.Updates.Count != 0:
        print(f"You have {result.Updates.Count} pending {"Windows updates." if result.Updates.Count != 1 else "Windows update."}")
    else:
        print("Your system is up-to-date!")

def main():
    check_os()
    print(f"{"-" * 50}" "\nWelcome to Helpy! The Ultimate Windows Maintenance tool!\nIt does everything what Microsoft Support would do, except \nyou don't have to wait in line for ages!\n"
          f"{"-" * 50}") 

    if not check_admin_privileges():
        if not request_admin():
            sys.exit("Administrator privileges were not granted. Please launch as Administrator.")

    global confirmation
    confirmation = run_all()
    temp_path = get_temp_path()
    clear_temp(temp_path)
    run_sfc()
    dism()
    check_windows_updates()
    print("\nAll done!")
    os.system('pause')


if __name__ == '__main__':
    main()