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

; ------ input params -----
; dir to copy iPiMotion files from for analysis
inputDir = ..\LW tracked files
; ------ input params -----

FileDelete, ./*.iPiMotion
FileCopy, %inputDir%\*.iPiMotion, ., 1

FileInstall, .\slider.png, %A_Temp%\slider.png
FileInstall, .\bubble.png, %A_Temp%\bubble.png

sliderValue(X1, Y1, X2, Y2, firstTick, lastTick)
{
	ImageSearch, X, Y, X1, Y1, X2, Y2, *5 %A_Temp%\slider.png
	if (ErrorLevel != 0)
		MsgBox Error
	a := X+5-firstTick ; add 5 to get icon center
	b := lastTick-firstTick
	If (a > 0)
		value :=  Round(a/b*100)
	else
		value := 0
	
	Return value
}

Sleep, 5000

FileDelete, .\ActorParams.csv
file := FileOpen(".\ActorParams.csv", "w")
file.write("File`,Gender`,Height`,Body Mass Index`,Feet Size`,Chest`,Bust`,Waist`,Hips`,Belly`,Legs Length`,Arms Length`,Shoulder Width`,Head Size`n")

Loop Files, .\*.iPiMotion
{
	file.write(A_LoopFileName "`,")
	Sleep, 250
	
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
	
	; Use Window Coordinates throughout
	;iconWidth := 10
	
	; Select Actor tab
	Click, 1709, 123
	Sleep, 1000
	
	;Gender
	ImageSearch, X, Y, 1630, 177, 1704, 198, *5 %A_Temp%\bubble.png
	if (ErrorLevel != 0)
		MsgBox Error
	diff := 1693-X
	If (diff < 20)
		gender = f
	else
		gender = m
	file.write(gender "`,")
	Sleep, 500
	
	; Click to set to Female to ungrey Slider to detect bust
	Click, 1693, 186
	
	; Height
	value := sliderValue(1571, 220, 1920, 248, 1615, 1890)
	file.write(value "`,")
	
	; Body Mass Index
	value := sliderValue(1571, 280, 1920, 309, 1690, 1907)
	file.write(value "`,")
	
	; Feet Size
	value := sliderValue(1571, 309, 1920, 336, 1690, 1907)
	file.write(value "`,")
	
	; Chest
	value := sliderValue(1571, 336, 1920, 364, 1690, 1907)
	file.write(value "`,")
	
	; Bust
	value := sliderValue(1571, 364, 1920, 395, 1690, 1907)
	file.write(value "`,")
	
	; Waist
	value := sliderValue(1571, 395, 1920, 422, 1690, 1907)
	file.write(value "`,")
	
	; Hips
	value := sliderValue(1571, 422, 1920, 450, 1690, 1907)
	file.write(value "`,")
	
	; Belly
	value := sliderValue(1571, 450, 1920, 480, 1690, 1907)
	file.write(value "`,")
	
	; Expand Body Dimensions
	Click, 1597, 854
	Sleep, 1000
	
	; Legs Length
	value := sliderValue(1571, 683, 1920, 713, 1690, 1907)
	file.write(value "`,")
	
	; Arms Length
	value := sliderValue(1571, 713, 1920, 740, 1690, 1907)
	file.write(value "`,")
	
	; Shoulders Width
	value := sliderValue(1571, 740, 1920, 766, 1690, 1907)
	file.write(value "`,")
	
	; Head Size
	value := sliderValue(1571, 766, 1920, 796, 1690, 1907)
	file.write(value "`,")
	
	file.write("`n")
	
	WinClose, ahk_pid %iPiPID%
	
	; If set to female slider when male (to detect bust) will need to deny saving changes
	WinWaitActive, Unsaved changes,, 3
	ControlSend,, {Tab}{Enter}, Unsaved changes
	
	WinWaitClose, ahk_pid %iPiPID%
}
file.close()
ExitApp