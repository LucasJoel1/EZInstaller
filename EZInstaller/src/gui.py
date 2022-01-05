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
from guimodules import *
from globalmodules import *
import tarfile

# don't open a console window
ctypes.windll.kernel32.SetConsoleTitleW("EzInstaller")

fullPath = None
url = None
fileName = None
file = None
type = None
isExe = False
isZip = False
isRar = False
isMsi = False
isWinget = False
isTar = False
isOther = False
storedInProgramFiles = True
storageLocation = None
fileToCreate = None
executable = None
extractedDir = None
runTheFile = True
createShortcut = True
install = False
warning = True
run = False


myappid = "lucasjoel1.EZInstaller.GUI.1.0.0"  # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

root = Tk()
frm = ttk.Frame(root, padding="5 5 5 5")
frm.pack()

# styling
EXTRA_LARGE_FONT = ("Verdana", 14, "bold")
LARGE_FONT = ("Verdana", 12)
NORM_FONT = ("Verdana", 10)
SMALL_FONT = ("Verdana", 8)
EXTRA_SMALL_FONT = ("Verdana", 6)
style = ttk.Style(root)
style.theme_names()
current_theme = style.theme_use()


ttk.Label(frm, text="EZInstaller GUI", font=EXTRA_LARGE_FONT).grid(
    column=1, row=0, pady=10, padx=10
)
# a textbox where the user can enter the url with a label that says URL:
url_label = ttk.Label(frm, text="URL:", font=NORM_FONT)
url_label.grid(column=0, row=1, sticky=W, padx=10, pady=10)

url_entry = ttk.Entry(frm, width=50, font=NORM_FONT)
url_entry.grid(column=1, row=1, sticky=W, padx=10, pady=10)

storage_location_label = ttk.Label(frm, text="Storage Location: ", font=NORM_FONT)
storage_location_label.grid(column=0, row=3, sticky=W, padx=10, pady=10)

storage_location_entry = ttk.Entry(frm, width=50, font=NORM_FONT)
storage_location_entry.grid(column=1, row=3, sticky=W, padx=10, pady=10)

storedInProgramFilesCheckboxStatus = IntVar()
createShortcutCheckboxStatus = IntVar()
runTheFileCheckboxStatus = IntVar()
createShortcutCheckboxStatus = IntVar()
storage_location_entry.configure(state=DISABLED)



def unzipper():
    outputBox = Text(frm, width=50, height=10, font=NORM_FONT, state=DISABLED)
    outputBox.grid(column=1, row=8, sticky=W, padx=10, pady=10)
    def append_to_output_box(text):
        outputBox.configure(state=NORMAL)
        outputBox.insert(END, text)
        outputBox.configure(state=DISABLED)
        # auto scroll to bottom
        outputBox.see(END)
    append_to_output_box("Beginning Extraction Process...\n")
    url_entry.configure(state=DISABLED)
    storage_location_entry.configure(state=DISABLED)
    stored_in_program_files_checkbox.configure(state=DISABLED)
    create_shortcut_checkbox.configure(state=DISABLED)
    submit_button.configure(state=DISABLED)
    storedInProgramFiles = storedInProgramFilesCheckboxStatus.get()
    if storedInProgramFiles == True:
        storageLocation = "C:\\Program Files\\"
        append_to_output_box("Storing in Program Files\n")
    else:
        storageLocation = storage_location_entry.get()
    url = url_entry.get()
    files = url.split(", " or ",")
    for i in range(len(files)):
        file[i] = requests.get(files[i])
        if file[i].status_code == 200:
            append_to_output_box("File found, preparing to download...\n")
            append_to_output_box("Downloading file...\n")
            with open(storageLocation + files[i].split("/")[-1], "wb") as f:
                f.write(file[i].content)
            append_to_output_box("File downloaded.\n")
            append_to_output_box("Extracting file...\n")
            with zipfile.ZipFile(storageLocation + files[i].split("/")[-1], "r") as zip_ref:
                zip_ref.extractall(storageLocation)
            append_to_output_box("File extracted.\n")
            append_to_output_box("Deleting file...\n")
            os.remove(storageLocation + files[i].split("/")[-1])
            append_to_output_box("File deleted.\n")
            if i != len(files) - 1:
                append_to_output_box("Moving To Next Zip...\n")
        append_to_output_box("All files extracted.\n")

def submit():
    pythoncom.CoInitialize()
    global fullPath
    global url
    global fileName
    global file
    global type
    global isExe
    global isZip
    global isRar
    global isMsi
    global isWinget
    global isTar
    global isOther
    global storedInProgramFiles
    global storageLocation
    global fileToCreate
    global executable
    global extractedDir
    global runTheFile
    global createShortcut
    global install
    global warning
    global run
    url_entry.configure(state=DISABLED)
    storage_location_entry.configure(state=DISABLED)
    stored_in_program_files_checkbox.configure(state=DISABLED)
    create_shortcut_checkbox.configure(state=DISABLED)
    submit_button.configure(state=DISABLED)
    unzipper_button.configure(state=DISABLED)
    if storedInProgramFiles == True:
        storageLocation = "C:\\Program Files\\"
    else:
        storageLocation = storage_location_entry.get()
    url = url_entry.get()
    print(f"url: {url}")
    print(f"storage location: {storageLocation}")
    print(f"stored in program files: {storedInProgramFiles}")
    print(f"create shortcut: {createShortcut}")
    print(f"run the file: {runTheFile}")

    outputBox = Text(frm, width=50, height=10, font=NORM_FONT, state=DISABLED)
    outputBox.grid(column=1, row=8, sticky=W, padx=10, pady=10)

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

    fileName = file.headers["Content-Disposition"].split("filename=")[1]
    print(f"fileName: {fileName}")

    append_to_output_box("Checking file type..." + "\n")
    if fileName.endswith(".exe"):
        isExe = True
        append_to_output_box("File is an executable." + "\n")
    elif fileName.endswith(".zip"):
        isZip = True
        append_to_output_box("File is a zip file." + "\n")
    elif fileName.endswith(".rar"):
        isRar = True
        append_to_output_box("File is a rar file." + "\n")
    elif fileName.endswith(".msi"):
        isMsi = True
        append_to_output_box("File is an MSI file." + "\n")
    elif "tar" in fileName:
        isTar = True
        append_to_output_box("File is a tar file." + "\n")
    else:
        isOther = True
        append_to_output_box("File is not an executable, zip, rar, or MSI file." + "\n")
        time.sleep(5)
        sys.exit()

    if isOther == True:
        append_to_output_box("Error: File is invalid \n")
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
        append_to_output_box("RAR is not currently supported" + "\n")
    elif isTar == True:
        append_to_output_box("Tar file detected." + "\n")
        append_to_output_box("Extracting tar file..." + "\n")
        with tarfile.open(storageLocation + fileName) as tar:
            tar.extractall(storageLocation + fileName.split(".")[0])
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
    else:
        append_to_output_box("File type not recognized." + "\n")
        time.sleep(5)
        sys.exit()


# a check box with a label install to program files and if it is unchecked it will ask for a location
stored_in_program_files_label = ttk.Label(
    frm, text="(Not Recommended For Unzipper) Install to Program Files: ", font=NORM_FONT
)
stored_in_program_files_label.grid(column=0, row=2, sticky=W, padx=10, pady=10)

stored_in_program_files_checkbox = ttk.Checkbutton(
    frm,
    variable=storedInProgramFilesCheckboxStatus,
    onvalue=True,
    offvalue=False,
    command=stored_in_program_files_check(
        storedInProgramFiles=storedInProgramFiles,
        storedInProgramFilesCheckboxStatus=storedInProgramFilesCheckboxStatus,
        storage_location_entry=storage_location_entry,
    ),
)

storedInProgramFilesCheckboxStatus.set(1)
stored_in_program_files_checkbox.grid(column=1, row=2, sticky=W, padx=10, pady=10)

# a check box with a label create a shortcut and if it is unchecked it will not create a shortcut
create_shortcut_label = ttk.Label(frm, text="(Does Not Apply To Unzipper)Create Shortcut: ", font=NORM_FONT)
create_shortcut_label.grid(column=0, row=4, sticky=W, padx=10, pady=10)

create_shortcut_checkbox = ttk.Checkbutton(
    frm,
    variable=createShortcutCheckboxStatus,
    onvalue=True,
    offvalue=False,
    command=create_shortcut_check(
        createShortcut=createShortcut,
        createShortcutCheckboxStatus=createShortcutCheckboxStatus,
    ),
)
createShortcutCheckboxStatus.set(1)
create_shortcut_checkbox.grid(column=1, row=4, sticky=W, padx=10, pady=10)

# submit button
submit_button = ttk.Button(
    frm, text="Submit", command=threading.Thread(target=submit).start
)
submit_button.grid(column=1, row=6, sticky=W, padx=10, pady=10)

unzipper_button = ttk.Button(
    frm, text="Unzipper", command=unzipper
)
unzipper_button.grid(column=1, row=7, sticky=W, padx=10, pady=10)

root.title("EzInstaller")
root.geometry("1200x600")
root.iconbitmap("C:\\Program Files\\EZInstaller\\EZInstaller\\assets\\pictures\\logo.ico")
root.mainloop()
