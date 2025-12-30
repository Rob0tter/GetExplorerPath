#Requires AutoHotkey v2
#SingleInstance

;@Ahk2Exe-ConsoleApp
;@Ahk2Exe-SetMainIcon .\icon.ico
;@Ahk2Exe-SetName Get Explorer Path
;@Ahk2Exe-SetFileVersion 1.0
;@Ahk2Exe-SetCopyright 2025 Dirk Schwarzmann
;@Ahk2Exe-SetDescription Gets the full path of the currently active tab of an Explorer Window

; Let the ini file have the same name as this script:
iniName := SubStr(A_ScriptName, 1, StrLen(A_ScriptName)-4) . ".ini"

; Get path mappings (if any)
pathMap := Map()
pathMap.Default := ""
pathMap.CaseSense := False
pathMap := IniReadKeys(iniName, "PathMapping")

; Read & set hotkeys:
; Defaults: Ctrl-Alt -F5, -F6, -F7
hk_CopyToClipboard := IniRead(iniName, "Hotkeys", "CopyToClipboard", "^!F5")
hk_InsertOnTextfield := IniRead(iniName, "Hotkeys", "InsertOnTextfield", "^!F6")
hk_BothHkFunctions := IniRead(iniName, "Hotkeys", "BothFunctions", "^!F7")

; If executed with command line parameters, only the requested function is called and the program terminates right after.
If (A_Args.Length > 0) {
	Switch A_Args[1], False {
		Case "CopyToClipboard":
			SetExplorerPath("CopyToClipboard")
		Case "CopyToStdOut":
			SetExplorerPath("CopyToStdOut")
		Default:
			FileAppend "Syntax: GetExplorerPath [CopyToClipboard | CopyToStdOut]", "*"
	}
	
	ExitApp 0
} Else {
	; No parameter given: register Hotkeys and run in background
	If (hk_CopyToClipboard != "") {
		Hotkey hk_CopyToClipboard, SetExplorerPath
	}
	If (hk_InsertOnTextfield != "") {
		Hotkey hk_InsertOnTextfield, SetExplorerPath
	}
	If (hk_BothHkFunctions != "") {
		Hotkey hk_BothHkFunctions, SetExplorerPath
	}
	
	; Hide the ugly console window (but stay in tray)
	WinHide "A"
}

/*****************************************************************************
	Functions
*****************************************************************************/

/*
	Get a map object containing all non-empty Key-Value-Pairs of a given section
	
	Input:
	- String:iniName - path & file name of the INI file to read from
	- String:iniSection - name of the section without brackets
	
	Return:
	- Map object - may be empty or contains Key-Value-Pairs
*/
IniReadKeys(iniName, iniSection) {
	sectionFound := False
	pathMap := Map()
	
	Loop Read, iniName {
		If (A_LoopReadLine = "[" . iniSection . "]") {
			sectionFound := True
		} Else {
			If (sectionFound = True And Trim(A_LoopReadLine) != "") {
				If (SubStr(A_LoopReadLine, 1, 1) != "[") {
					; Found an entry within the proper section
					kvp := StrSplit(A_LoopReadLine, "=")
					If (kvp.Length > 1)
						pathMap[Trim(kvp[1])] := Trim(kvp[2])
				} Else {
					; Next section begins, we are finished
					sectionFound := False
					Continue
				}
			}
		}
	}
	
	Return pathMap
}

/*
	Get the full path of the currently active tab of the explorer window.
	If no window handle of an Explorer is provided, the program gets the last active Explorer Window itself.
	In case the path string is found in the path mapping object, the string will be replaced before being returned.
	
	Input:
	- HWND:hwnd - (optional) Handle to Explorer Window
	
	Return:
	- String - contains a full (absolute) path or a substitute or a raw ID of a special folder. May also be empty if no explorer window was found.
*/
GetExplorerActualPath(hwnd := WinExist("ahk_class CabinetWClass")) {
	; Do something only if Windows Explorer instance is found
	If (hwnd = 0) {
		Return ""
	}
	
	activeTab := 0
	activeTab := ControlGetHwnd("ShellTabWindowClass1", hwnd)
	
	For (wndw in ComObject("Shell.Application").Windows) {
		If (wndw.hwnd != hwnd)
			Continue
		If (activeTab > 0) {
			; The window has tabs, so make sure this is the right one.
			Static IID_IShellBrowser := "{000214E2-0000-0000-C000-000000000046}"
			shellBrowser := ComObjQuery(wndw, IID_IShellBrowser, IID_IShellBrowser)
			ComCall(3, shellBrowser, "uint*", &thisTab:=0)
			If (thisTab != activeTab)
				Continue
		}
		
		path := wndw.Document.Folder.Self.Path
		
		; Set path mapping if any:
		If (pathMap.Has(path)) {
			path := pathMap[path]
		}

		Return path
	}
}

/*
	Set / output the path to the desired destination.
	
	Input:
	- String:thisHK - not necessarily the hotkey identifier but could also be another predefined verb to identify the target
	
	Return:
	- Nothing
*/
SetExplorerPath(thisHK) {
	path := GetExplorerActualPath()
	
	If (path != "") {
		Switch (thisHK) {
			Case hk_CopyToClipboard, "CopyToClipboard":
				A_Clipboard := path
			Case hk_InsertOnTextfield:
				SendText path
			Case hk_BothHkFunctions:
				A_Clipboard := path
				SendText path
			Case "CopyToStdOut":
				FileAppend path, "*"
			Default:
		}
	}
}
