from tkinter import *
from tkinter import ttk
import ctypes
import requests
import time
import os
import sys
from requests.models import Response
import winshell
import zipfile
import threading
import pythoncom

pythoncom.CoInitialize()


def stored_in_program_files_check(
    storedInProgramFiles, storedInProgramFilesCheckboxStatus, storage_location_entry
):
    if storedInProgramFilesCheckboxStatus.get() == 1:
        storage_location_entry.delete(0, END)
        storage_location_entry.configure(state=DISABLED)
        storedInProgramFiles = True
    else:
        storage_location_entry.configure(state=NORMAL)
        storedInProgramFiles = False


def create_shortcut_check(createShortcut, createShortcutCheckboxStatus):
    if createShortcutCheckboxStatus.get() == 1:
        createShortcut = True
    else:
        createShortcut = False


def submit(
    url_entry,
    storage_location_entry,
    stored_in_program_files_checkbox,
    create_shortcut_checkbox,
    submit_button,
    storedInProgramFiles,
    createShortcut,
    runTheFile,
    storageLocation,
    fileName,
    file,
    type,
    isExe,
    isZip,
    isMsi,
    isOther,
    NORM_FONT,
    extractedDir,
    executable,
    frm,
):
    pythoncom.CoInitialize()
    url_entry.configure(state=DISABLED)
    storage_location_entry.configure(state=DISABLED)
    stored_in_program_files_checkbox.configure(state=DISABLED)
    create_shortcut_checkbox.configure(state=DISABLED)
    submit_button.configure(state=DISABLED)
    if storedInProgramFiles == True:
        storageLocation = "C:\\Program Files\\"
    else:
        storageLocation = storage_location_entry.get()
    url = url_entry.get()
    fileName = url.split("/")[-1]
    print(f"url: {url}")
    print(f"fileName: {fileName}")
    print(f"storage location: {storageLocation}")
    print(f"stored in program files: {storedInProgramFiles}")
    print(f"create shortcut: {createShortcut}")
    print(f"run the file: {runTheFile}")

    outputBox = Text(frm, width=50, height=10, font=NORM_FONT, state=DISABLED)
    outputBox.grid(column=1, row=7, sticky=W, padx=10, pady=10)

    # append to output box
    def append_to_output_box(text):
        outputBox.configure(state=NORMAL)
        outputBox.insert(END, text)
        outputBox.configure(state=DISABLED)
        # auto scroll to bottom
        outputBox.see(END)

    append_to_output_box(
        f"Preparing to download file from {url} to {storageLocation}" + "\n"
    )
    file = requests.get(url, allow_redirects=True)

    if file.status_code != 200:
        append_to_output_box("Error: " + file.status_code + "\n")
        append_to_output_box("Please check the URL and try again." + "\n")
        time.sleep(5)
        sys.exit()
    fileName = file.url.split("/")[-1]

    if "." not in fileName:
        append_to_output_box("Error: " + "File must have an extension." + "\n")
        append_to_output_box("Please check the URL and try again." + "\n")
        time.sleep(5)
        sys.exit()
    append_to_output_box(f"File found, download {fileName} to {storageLocation}" + "\n")
    storageLocation = storageLocation.replace("\\", "/")

    if storageLocation[-1] != "/":
        storageLocation += "/"

    with open(storageLocation + fileName, "wb") as f:
        append_to_output_box("Downloading %s" % fileName + "\n")
        response = requests.get(url, stream=True)
        total_length = response.headers.get("content-length")
        if total_length is None:  # no content length header
            f.write(response.content)
        else:
            dl = 0
            total_length = int(total_length)
            for data in response.iter_content(chunk_size=4096):
                dl += len(data)
                f.write(data)
                if dl == total_length:
                    append_to_output_box("Download complete." + "\n")
                    f.close()
                    break

    append_to_output_box("Checking file type..." + "\n")
    type = fileName.split(".")[-1]
    append_to_output_box(f"File type: {type}" + "\n")
    if type == "exe":
        isExe = True
        append_to_output_box("File is an executable." + "\n")
    elif type == "zip":
        isZip = True
        append_to_output_box("File is a zip file." + "\n")
    elif type == "rar":
        isRar = True
        append_to_output_box("File is a rar file." + "\n")
    elif type == "msi":
        isMsi = True
        append_to_output_box("File is an MSI file." + "\n")
    else:
        isOther = True
        append_to_output_box("File is not an executable, zip, rar, or MSI file." + "\n")
        time.sleep(5)
        sys.exit()

    if isExe == True:
        append_to_output_box("Running file..." + "\n")
        os.system(storageLocation + fileName)
        time.sleep(5)
        sys.exit()
    elif isZip == True:
        append_to_output_box("Extracting zip file..." + "\n")
        with zipfile.ZipFile(storageLocation + fileName, "r") as zip_ref:
            zip_ref.extractall(storageLocation + fileName.split(".")[0])
        os.remove(storageLocation + fileName)
        append_to_output_box("Extraction complete." + "\n")
        extractedDir = storageLocation + fileName.split(".")[0]
        append_to_output_box(f"Finding executable in {extractedDir}" + "\n")
        for root, dirs, files in os.walk(extractedDir):
            for file in files:
                if file.endswith(".exe"):
                    # path of the file
                    fullPath = os.path.join(root, file)
                    executable = file
                    append_to_output_box(f"Found executable: {executable}" + "\n")
        if createShortcut == True:
            append_to_output_box("Creating shortcut..." + "\n")
            desktop = winshell.desktop()
            shorthcut = desktop + "\\" + executable.split(".")[0] + ".lnk"
            winshell.CreateShortcut(
                Path=shorthcut, Target=fullPath, Description=executable.split(".")[0]
            )
            append_to_output_box("Shortcut created." + "\n")
        elif createShortcut == False:
            append_to_output_box("Shortcut not created due to user preferences" + "\n")
        append_to_output_box("Installation complete" + "\n")
        time.sleep(5)
        sys.exit()
    elif isMsi == True:
        append_to_output_box("MSI file detected." + "\n")
        append_to_output_box("Installing MSI file..." + "\n")
        os.system(f"msiexec /i {storageLocation + fileName}")
        time.sleep(5)
        sys.exit()
    elif isRar == True:
        append_to_output_box("RAR file detected." + "\n")
        append_to_output_box("RAR is not currently supported." + "\n")
    elif isOther == True:
        append_to_output_box("Other file detected." + "\n")
        append_to_output_box("Other file types are not currently supported." + "\n")
    else:
        append_to_output_box("File type not recognized." + "\n")
        time.sleep(5)
        sys.exit()
