#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include function.ahk

F1::

	FoundPos = 1

	No_of_Boxes = 12 33
	
	MsgBox, % No_of_Boxes

	HowManyBoxes(No_of_Boxes)
	
	
	l = 1


	
	;박스 담긴 변수 받아서 각 배열에 넣기
	HowManyBoxes(No_of_Boxes){
		
		while(FoundPos := RegExMatch(No_of_Boxes, "([0-9].*)", BoxWeight, FoundPos + strLen(BoxWeight))){
			
			;MsgBox, % BoxWeight

			Box_arr[l] := BoxWeight
			l += 1
			;MsgBox, % Box_arr[l]
		}
		
		
	}



Exitapp

Esc::
 Exitapp