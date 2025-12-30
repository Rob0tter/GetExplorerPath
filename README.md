![Logo](./src/icon.ico)
 
# GetExplorerPath

Get the full path of the currently active tab of an Explorer Window

## Installation
No installation necessary, just copy the executable and the ini file in the same folder and run.

## Description
If program is run with a parameter (preferably from a command line), it will either return the path for further use on std out or copy it into clipboard and then terminates with exit code 0.
Command line parameter could be *CopyToClipboard* or *CopyToStdOut*.

If program is run without a parameter (e.g. by double-clicking it), it will silently sit in the tray and wait for a hotkey press.

The INI file allows you to define three hotkeys:

- To copy the path to the clipboard only (overwriting previous content)
- To write the path to the currently active text input field of any program
- To perform both actions simultaneously

You can also define substitutions for any path, for example, to replace the internal identifier "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" with a more readable "Network".

If multiple Explorer windows are open, the most recently used one will be considered.
If multiple tabs are open in Explorer, the active one will be used.

## Tested environments
This program is tested to work with at least:

- Windows 11 Pro 24H2
- Windows 11 Pro 25H2

## Compile
AutoHotkey v2 is necessary to run and compile this script.
All compiler settings are in the script file, no further settings needed in the GUI of AHK2EXE.
The script is designed to be run on the command line as well as resident in the tray.
