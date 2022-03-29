import ctypes
import errno
import os.path
import subprocess


def ms_store():
    launch_dir = os.getcwd()
    try:
        print(launch_dir)
        store_apps = open(f"{launch_dir}/store_apps.txt")

    except IOError as I:
        if I.errno == errno.ENOENT:
            print("This file does not exist")
            msgbox("The Paths.txt file is not in the root, Exiting!", "Error")
            raise SystemExit(0)
        elif I.errno == errno.EACCES:
            print("You do not have read access to this file.")
            msgbox("This program does not have read access to Paths.txt, Exiting!", "Error")
            raise SystemExit(0)
        else:
            print("An unexpected error has occurred")
            print()
            print(I.errno)
            msgbox("An unexpected error has occurred while reading the file, Exiting!", "Error")
            raise SystemExit(0)

    for line in store_apps:
        without_space = line.strip()
        remove_cmd = run(f'get-appxpackage -allusers -name "{without_space}" | remove-appxpackage')
        if remove_cmd.returncode != 0:
            print(f"An error has occurred: \n {remove_cmd.stderr}")
        else:
            print(line + " Has been removed successfully")
            print(remove_cmd.stderr)

    # msgbox("Microsoft Apps have been removed.", "Apps Removed")


def shortcut_func():
    launch_dir = os.getcwd()
    try:
        print(launch_dir)
        shortcuts = open(f"{launch_dir}/shortcuts.txt")

    except IOError as I:
        if I.errno == errno.ENOENT:
            print("This file does not exist")
            msgbox("The shortcuts.txt file is not in the working directory, Exiting!", "Error")
            raise SystemExit(0)
        elif I.errno == errno.EACCES:
            print("You do not have read access to this file.")
            msgbox("This program does not have read access to the working directory, Exiting!", "Error")
            raise SystemExit(0)
        else:
            print("An unexpected error has occurred")
            print()
            print(I.errno)
            msgbox("An unexpected error has occurred while reading the file, Exiting!", "Error")
            raise SystemExit(0)

    for line in shortcuts:
        os.remove(line)


def msgbox(message, title):
    ctypes.windll.user32.MessageBoxW(0, message, title, 48)


def run(cmd):
    completed = subprocess.run(["powershell", "-Command", cmd], capture_output=True)
    return completed


def main():
    ms_store()
    shortcut_func()


if __name__ == '__main__':
    main()
