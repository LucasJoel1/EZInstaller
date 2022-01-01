# download: C:\other\coding\ezinstaller\downloads

import requests
import time
import os
import sys
import winshell
import zipfile



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
isOther = False
storedInProgramFiles = False
storageLocation = None
fileToCreate = None
executable = None
extractedDir = None
runTheFile = False
createShortcut = False
install = False

url = input("Enter a link to the file (for winget input the package name): ")
storedInProgramFiles = input("Should the file be stored in Program Files? (y/n): ")
if storedInProgramFiles == "y":
    storedInProgramFiles = True
    storageLocation = "C:\\Program Files\\"

elif storedInProgramFiles == "n":
    storedInProgramFiles = False
    storageLocation = input("Enter the location to store the file (full path): ")


print(f"Preparing to download file from {url} to {storageLocation}")
file = requests.get(url, allow_redirects=True)

if file.status_code != 200:
    print("Error: File not found")
    time.sleep(5)
    sys.exit(1)
fileName = file.url.split("/")[-1]

if "." not in fileName:
    print("Error: File must have an extension")
    time.sleep(5)
    sys.exit(1)
print(f"File found, downloading {fileName} to {storageLocation}")

storageLocation = storageLocation.replace("\\", "/")
# if storage location does not end with "/" then add it
if storageLocation[-1] != "/":
    storageLocation += "/"

# download file
with open(storageLocation + fileName, "wb") as f:
    print("Downloading %s" % fileName)
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
            done = int(50 * dl / total_length)
            sys.stdout.write("\r[%r%s]" % ("=" * done, " " * (50 - done)))
            sys.stdout.flush()

print("\n")
print("Download complete")

type = fileName.split(".")[-1]

if type == "exe":
    isExe = True
elif type == "zip":
    isZip = True
elif type == "rar":
    isRar = True
elif type == "msi":
    isMsi = True
else:
    isOther = True
    type = input("Would you like to use winget instead? (y/n): ")
    if type == "y":
        isWinget = True
    elif type == "n":
        isWinget = False
    else:
        print("Error: Invalid input")
        time.sleep(5)
        sys.exit(1)

if isOther == True:
    print("Error: File type not supported")
    time.sleep(5)
    sys.exit(1)

if isExe == True:
    print("File is an executable")
    print("Running executable")
    os.system(storageLocation + fileName)
    time.sleep(5)
    sys.exit(0)
elif isZip == True or isRar == True:
    print("File is a zip file")
    print("Extracting zip file")
    with zipfile.ZipFile(storageLocation + fileName, "r") as zip_ref:
        zip_ref.extractall(storageLocation + fileName.split(".")[0])
    os.remove(storageLocation + fileName)
    print("Extraction complete")
    extractedDir = storageLocation + fileName.split(".")[0]
    print(f"finding executable in {extractedDir}")
    for root, dirs, files in os.walk(extractedDir):
        for file in files:
            if file.endswith(".exe"):
                # path of the file
                fullPath = os.path.join(root, file)
                executable = file
                print(f"Found executable: {executable}")
    createShortcut = input("Create shortcut? (y/n): ")
    if createShortcut == "y":
        createShortcut = True
    elif createShortcut == "n":
        createShortcut = False
    else:
        print("Error: Invalid input")
        time.sleep(5)
        sys.exit(1)
    print(fullPath)
    if createShortcut == True:
        # create a desktop shortcut
        print("Creating shortcut")
        desktop = winshell.desktop()
        shorthcut = desktop + "\\" + executable.split(".")[0] + ".lnk"
        winshell.CreateShortcut(
            Path=shorthcut, Target=fullPath, Description=executable.split(".")[0]
        )
        print("Shortcut created")
    runTheFile = input("Run the file? (y/n): ")
    if runTheFile == "y":
        runTheFile = True
    elif runTheFile == "n":
        runTheFile = False
    else:
        print("Error: Invalid input")
        time.sleep(5)
        sys.exit(1)
    if runTheFile == True:
        print("Running the file")
        os.system(fullPath)
    print("Done")
    time.sleep(5)
    sys.exit(0)
elif isMsi == True:
    print("File is an MSI")
    print("Installing MSI")
    os.system(storageLocation + fileName)
    time.sleep(5)
    sys.exit(0)
elif isWinget == True:
    print("Using winget to install")
    os.system(f"winget show {url}")
    install = input("Install? (y/n): ")
    if install == "y":
        install = True
    elif install == "n":
        install = False
    if install == "y":
        os.system(f"winget install {url}")
    print("Done")
    time.sleep(5)
    sys.exit(0)
else:
    print("Error: File type not supported")
    time.sleep(5)
    sys.exit(1)
