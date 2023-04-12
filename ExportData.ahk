#Include C:\Users\%A_UserName%\Dropbox\Include\Windows\functions.ahk
global include := "C:\Users\" A_UserName "\Dropbox\Include\Windows"

#NoEnv
SetWorkingDir %A_ScriptDir%
CoordMode, Mouse, Window
SendMode Input
#SingleInstance Force
SetTitleMatchMode 2
DetectHiddenWindows On
#WinActivateForce
SetControlDelay 1
SetWinDelay 0
SetKeyDelay 50, 50
SetMouseDelay, 50
SetBatchLines -1
#Persistent  ; Prevent the script from exiting automatically.

; Made for 1920x1080 at 100% scaling
; Right parameter window pushed most to as far right as possible
Sleep, 5000

; ------ input params -----
; dir to copy iPiMotion files from for analysis
inputDir = ..\LW tracked files
; ------ input params -----

FileDelete, .\*.iPiMotion
FileCopy, %inputDir%\*.iPiMotion, ., 1

; Dump nTakes to common csv
nTakesFile := FileOpen(".\nPasses.csv", "w")
nTakesFile.write("file`,nPasses`n")

Loop Files, .\*.iPiMotion
{
	SplitPath, A_LoopFileLongPath, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
	OutDir := A_WorkingDir
	;OutDir = %inputDir%\%OutNameNoExt%
	;IfNotExist, %OutDir%
	;	FileCreateDir, %OutDir%
	
    Run, %A_LoopFileName%,,, iPiPID
	WinWaitActive, ahk_pid %iPiPID%
	WinWaitActive, License Dialog,, 5
	Sleep, 1000
	ControlSend,, {Tab 2}{Enter}, License Dialog
	; skip needing file
	WinWaitActive, Video not found
	Sleep, 1000
	ControlSend,, {Tab}{Enter}, Video not found
	Sleep, 5000
	
	; Select Biomech tab
	Click, 1865, 97
	Sleep, 1000
	; Select Load Button
	Click, 1700, 158
	Sleep, 3000
	Send, C:\Users\%A_UserName%\Dropbox\Active\infantGait\export.iPiBiomech{Enter}
	Sleep, 3000
	
	; Initialize Take Menu to max take
	Click, 700, 939
	Send, {End}{Enter}{Tab}
	MouseMove, 100, 100
	Sleep, 2000
	
	; Grab number of Takes
	cmd = %include%\Capture2Text\Capture2Text_CLI.exe -d -s "656 925 724 939" ; Use Screen coords from Window Spy
	nTakes := returnCmd(cmd)
	StringTrimLeft, nTakes, nTakes, 4
	
	; Common Takes File
	nTakesFile.write( OutNameNoExt "`," nTakes "`n")
	
	; Initialize Master Takes file for subject session
	outFile := FileOpen(OutDir "\" OutNameNoExt ".csv", "w")
	headerDone := false
	Loop, %nTakes%
	{
		nTake := A_Index
		
		Click, 700, 939
		Sleep, 1000
		nDown := nTake - 1
		Send, {Home}
		Sleep, 500
		Send, {down %nDown%}

		Send, {Enter}
		Sleep, 500
		Click, 804, 941
		
		; ---------------------
		; Export Data
		Click, 1826, 866
		Sleep, 1000
		Send, {down}{Enter}
		Sleep, 3000
		
		; Enter new name with Take
		tempName = %OutDir%\%OutNameNoExt%_%A_Index%
		Send, %tempName%
		Send, {Enter}
		
		SLeep, 2000
		ControlSend,, {Tab}{Enter}, Confirm Save As
		Sleep, 5000
		; ---------------------
		
		; Fix text file
		Loop, Read, %tempName%.txt
		{
			out := StrReplace(A_loopReadLine, A_Tab, "`,")
			
			If (A_Index = 8) and (!headerDone)
			{
				outFile.write("Frame" out "`,Pass`n")
				headerDone := true
			}
			
			If (A_Index > 8)
				outFile.write(out "`," nTake "`n")
		}
		FileDelete, %tempName%.txt
	}
	outFile.close()
	
	WinClose, ahk_pid %iPiPID%
	Sleep, 2000
	ControlSend,, {Tab}{Enter}, Unsaved changes
	WinWaitClose
}

nTakesFile.close()

ExitApp
