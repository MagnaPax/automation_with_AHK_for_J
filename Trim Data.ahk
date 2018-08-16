#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

/*
Data := "ab,,, Cdef                asdf   "

Data := Clipboard

MsgBox, % "Origin Data`n"Data "'"
;~ MsgBox, % 

;~ StringReplace, clipboard, clipboard, `r`n, , All
StringReplace, Data, Data, `,, , All

MsgBox, % ", exculsive`n"Data "'"

Data := Trim(Data)

MsgBox, % "Trimed`n"Data "'"

StringUpper, Data, Data ; Staff only notes 대문자로 바꾸기


MsgBox, % Data "'"
*/


^1::
;~ SetKeyDelay, 300
;~ SetKeyDelay 50,200
SetKeyDelay, 1000
;~ SetKeyDelay 300,200

Data = %Clipboard%

StringReplace, Data, Data, ', , All
StringReplace, Data, Data, -, , All
StringReplace, Data, Data, (, , All
StringReplace, Data, Data, ), , All
Data := Trim(Data)
StringUpper, Data, Data ; 대문자로 바꾸기

;~ StringLeft, Data, Data, 20  ; 왼쪽부터 20개 읽어서 저장하기

Send, %Data%
return




^2::
SetKeyDelay, 1000
Data = %Clipboard%

;~ RegExMatch(Data, "imU)(\d*)\.", SubPat)
;~ Data := SubPat1

Data := Trim(Data)
Send, %Data%
return




^3::
SetKeyDelay, 1000

Data = %Clipboard%

Data := RegExReplace(Data, "[^0-9]", "") ;숫자만 저장

StringReplace, Data, Data, ', , All
StringReplace, Data, Data, -, , All
StringReplace, Data, Data, (, , All
StringReplace, Data, Data, ), , All
StringReplace, Data, Data, %A_SPACE%, , All
StringReplace, Data, Data, `n, , All
StringReplace, Data, Data, `r, , All
StringUpper, Data, Data ; 대문자로 바꾸기
Data := Trim(Data)


;~ StringLeft, Data, Data, 20  ; 왼쪽부터 20개 읽어서 저장하기

Send, %Data%
return




Exitapp


^Esc::
 Exitapp