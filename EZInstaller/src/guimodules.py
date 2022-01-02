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
