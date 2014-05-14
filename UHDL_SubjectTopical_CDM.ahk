#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force

; UHDL_SubjectTopical_CDM.ahk
; 0.1 beta
; 0.2 gui boxes; existing lcsh term logic; open reports in notepad++; refine match condition; authority uri

; variables ================================
reportfile = path\to\report\UHDL_SubjectTopical_Report_%A_YYYY%%A_MM%%A_DD%.txt
uhdlsubjectfile = path\to\authority\file\UHDL_SubjectTopical.txt
appname = UHDL Subject.Topical Authorities 0.2

userinfo =
userinfolen = 0
timestamplen = 17
divider =
previousLCSH =
runcount = 0

TW = 560
BW = 590
Yvalue = 10
Ycounter = 30
Ydiff = 50
Ytext = 33

Y1 := Yvalue
Y2 := Yvalue + Ydiff
Y3 := Yvalue + 2 * Ydiff
Y4 := Yvalue + 3 * Ydiff
Y5 := Yvalue + 4 * Ydiff
Y6 := Yvalue + 5 * Ydiff

YT1 = 25
YT2 := Ytext + Ydiff
YT3 := Ytext + 2 * Ydiff
YT4 := Ytext + 3 * Ydiff
YT5 := Ytext + 4 * Ydiff
YT6 := Ytext + 5 * Ydiff
; variables ================================

Gui, 1:Color, d0d0d0, 912206
Gui, 1:Show, h0 w0, %appname%

Menu, FileMenu, Add, &Reload, Reload
Menu, FileMenu, Add, E&xit, Exit
Menu, EditMenu, Add, &Authority File    (Ctrl+Alt+A), Authorities
Menu, EditMenu, Add, &Report File         (Ctrl+Alt+R), Report
Menu, MenuBar, Add, &File, :FileMenu
Menu, MenuBar, Add, &Edit, :EditMenu
Gui, Menu, MenuBar

; labels ================================
Gui, Font,, Arial

Gui, Add, Text, x10 y%Y1%, PREVIOUS LCSH
Gui, Add, Text, x10 y%Y2% w%TW% h20, LCSH
Gui, Add, Text, x10 y%Y3% w%TW% h20, TGM
Gui, Add, Text, x10 y%Y4% w%TW% h20, AAT
Gui, Add, Text, x10 y%Y5% w%TW% h20, SAA
Gui, Add, Text, x10 y%Y6% w%TW% h20, UHDL

; STATIC 7-12
Gui, Add, Text, x30 y%YT1% w530 h20,
Gui, Add, Text, x30 y%YT2% w550 h20,
Gui, Add, Text, x30 y%YT3% w500 h20,
Gui, Add, Text, x30 y%YT4% w500 h20,
Gui, Add, Text, x30 y%YT5% w500 h20,
Gui, Add, Text, x30 y%YT6% w500 h20,

Gui, Add, GroupBox, x5 y0 w%BW% h50,
Gui, Add, GroupBox, x5 y50 w%BW% h260,
; labels ================================



WinGetPos, winX, winY, winWidth, winHeight, %appname%
winX+=%winWidth%
Gui, 1:Show, x%winX% y%winY% h315 w600, %appname%
WinActivate, %appname%


; hotkeys ================================
^!a::
	Gosub, Authorities
Return

^!r::
	Gosub, Report
Return

^!s::
	WinGet, cdmwindow, ID, CONTENTdm ; CDM window id
	clipsave = %clipboard% ; save clipboard contents

	LCSHvar =
	ControlSetText, Static8,, %appname%
	ControlSetText, Static9,, %appname%
	ControlSetText, Static10,, %appname%
	ControlSetText, Static11,, %appname%
	ControlSetText, Static12,, %appname%

	tgmreport =
	aatreport =
	saareport =
	uhdlreport =

	if (runcount == 0)
	{
		InputBox, input,, Please Enter Your Initials,, 180, 125,,,,,
		if ErrorLevel
			Return
		else
		{
			username = %input%
			InputBox, input,, Please Enter The Project Name,, 380, 125,,,,,
			if ErrorLevel
				Return
			else
			{
				projectname = %input%
				userinfo = %projectname% %username%
				StringLen, userinfolen, userinfo
				dividerlen := userinfolen + timestamplen + 2
				Loop, %dividerlen%
				{
					divider .= "_"
				}
			}
		}
	}

	if (runcount > 0)
	{
		ControlSetText, Static7, %previousLCSH%, %appname%
	}

	; read in local authority list
	FileRead, UHDL_Subjects, %uhdlsubjectfile%

; TGM ===============
	WinActivate, ahk_id %cdmwindow%	; activate cdm window

	; copy field contents
	Clipboard =
	Gosub, CopyField			
	
	; process field contents
	subjectstring = %clipboard%
	StringReplace, subjectstring, subjectstring, `n,, All
	StringReplace, subjectstring, subjectstring, `r,, All

	; map to LCSH if field content exists
	StringLen, length, subjectstring
	if (length > 0)
	{
		LCSHvar := SubjectMap("TGM", subjectstring, reportfile, LCSHvar, UHDL_Subjects, uhdlsubjectfile, 9, appname, username)
		tgmreport = TGM: %subjectstring%`n`n
	}

; AAT ===============
	WinActivate, ahk_id %cdmwindow%	; activate cdm window

	; copy field contents
	Clipboard =
	Sleep, 500
	Gosub, CopyField			
	
	; process field contents
	subjectstring = %clipboard%		
	StringReplace, subjectstring, subjectstring, `n,, All
	StringReplace, subjectstring, subjectstring, `r,, All

	; map to LCSH if field content exists
	StringLen, length, subjectstring
	if (length > 0)
	{
		LCSHvar := SubjectMap("AAT", subjectstring, reportfile, LCSHvar, UHDL_Subjects, uhdlsubjectfile, 10, appname, username)
		aatreport = AAT: %subjectstring%`n`n
	}

; SAA ===============
	WinActivate, ahk_id %cdmwindow%	; activate cdm window

	; copy field contents
	Clipboard =
	Sleep, 500
	Gosub, CopyField			
	
	; process field contents
	subjectstring = %clipboard%		
	StringReplace, subjectstring, subjectstring, `n,, All
	StringReplace, subjectstring, subjectstring, `r,, All

	; map to LCSH if field content exists
	StringLen, length, subjectstring
	if (length > 0)
	{
		LCSHvar := SubjectMap("SAA", subjectstring, reportfile, LCSHvar, UHDL_Subjects, uhdlsubjectfile, 11, appname, username)
		saareport = SAA: %subjectstring%`n`n
	}

; UHDL ===============
	WinActivate, ahk_id %cdmwindow%	; activate cdm window

	; copy field contents
	Clipboard =
	Sleep, 500
	Gosub, CopyField			
	
	; process field contents
	subjectstring = %clipboard%		
	StringReplace, subjectstring, subjectstring, `n,, All
	StringReplace, subjectstring, subjectstring, `r,, All

	; map to LCSH if field content exists
	subjectstring = %clipboard%		
	StringReplace, subjectstring, subjectstring, `n,, All
	StringReplace, subjectstring, subjectstring, `r,, All

	StringLen, length, subjectstring
	if (length > 0)
	{
		LCSHvar := SubjectMap("UHDL", subjectstring, reportfile, LCSHvar, UHDL_Subjects, uhdlsubjectfile, 12, appname, username)
		uhdlreport = UHDL: %subjectstring%`n`n
	}

	StringTrimRight, LCSHvar, LCSHvar, 2 ; remove trailing semicolon and space

	; format record entry
	FileAppend, %divider%`n%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min% %userinfo%`n`n, %reportfile%
	FileAppend, %tgmreport%%aatreport%%saareport%%uhdlreport%, %reportfile%
	FileAppend, LCSH: %LCSHvar%`n`n, %reportfile%
	ControlSetText, Static8, %LCSHvar%, %appname%

	; activate cdm window
	WinActivate, ahk_id %cdmwindow%

	; populate LCSH field with mapped values
	Send, {Left}
	Sleep, 50
	Send, {Left}
	Sleep, 50
	Send, {Left}
	Sleep, 50
	Send, {Left}
	Sleep, 50
	Send, {Left}
	Sleep, 50
	StringLen, length, LCSHvar
	if (length > 0)
	{
		Clipboard =
		Gosub, CopyField
		Sleep, 200
		Send, {Left}
		StringLen, length, Clipboard
		if (length > 0)
		{
			originalLCSH = %Clipboard%
			StringReplace, originalLCSH, originalLCSH, `n,, All
			StringReplace, originalLCSH, originalLCSH, `r,, All
			StringRight, rightchar, originalLCSH, 1
			if (rightchar == ";")
			{
				LCSHvar := A_Space . LCSHvar
			}
			Else
			{
				LCSHvar := "; " . LCSHvar
			}
		}	
		Send, {F2}
		Sleep, 100
		Send, %LCSHvar%
		Sleep, 500
	}
	Send, {Tab}

	previousLCSH = %LCSHvar%
	runcount++
	clipboard = %clipsave% ; restore clipboard

	Sleep, 200
	WinActivate, %appname%
	Sleep, 200
	WinActivate, ahk_id %cdmwindow%
Return
; hotkeys ================================

; helper functions ================================
CopyField:
	Send, {F2}
	Sleep, 50
	Send, ^a
	Sleep, 50
	Send, ^c
	Sleep, 50
	Send, {Tab}
Return

SubjectMap(authority, subjectstring, reportfile, LCSHvar, UHDL_Subjects, uhdlsubjectfile, controlnum, appname, username)
{
	ControlSetText, Static%controlnum%, %subjectstring%, %appname%
	count = 0
	Loop, parse, subjectstring, `;
	{
		StringLen, length, A_LoopField
		if (length > 0)
		{
			match = 0
			subject = %A_LoopField%
			Loop, parse, UHDL_Subjects, `n
			{
				UHDL_ListEntry = %A_LoopField%
				StringReplace, UHDL_ListEntry, UHDL_ListEntry, `n,, All
				StringReplace, UHDL_ListEntry, UHDL_ListEntry, `r,, All

				Loop, parse, UHDL_ListEntry, `t
				{
					count++
					if (count == 1)
					{
						if (A_LoopField == authority)
							Continue
						Else
						{
							count = 0
							Break
						}
					}

					if (count == 2)
					{
						if (A_LoopField == subject)
							Continue
						Else
						{
							count = 0
							Break
						}
					}

					if (count == 3)
					{
						IfNotInString, LCSHvar, %A_LoopField%
						{
							match = 1
							LCSHvar := LCSHvar . A_LoopField . "; "
						}
						count = 0
						Break
					}
				}
			}
			if (match == 0)
			{
				Run, http://id.loc.gov/search/?q=%subject%&q=cs`%3Ahttp`%3A`%2F`%2Fid.loc.gov`%2Fauthorities`%2Fsubjects
				InputBox, LCSHauthorized, Authorized Term, `n%subject%`n`nEnter an authorized term:,, 550, 175,,,,,
				if ErrorLevel
					Return
				else
				{
					InputBox, LCSHuri, Authorized Term URI, `n%LCSHauthorized%`n`nEnter the URI:,, 550, 175,,,,,
					if ErrorLevel
						Return
					else
					{
						IfNotInString, LCSHvar, %LCSHauthorized%
						{
							StringReplace, LCSHauthorized, LCSHauthorized, `n,, All
							StringReplace, LCSHauthorized, LCSHauthorized, `r,, All

							LCSHvar := LCSHvar . LCSHauthorized . "; "
						}
						FileAppend, %authority%`t%subject%`t%LCSHauthorized%`t%LCSHuri%`t%username%`t%A_YYYY%%A_MM%%A_DD%`n, %uhdlsubjectfile%
					}
				}
			}
		}
	}

Return LCSHvar
}
; helper functions ================================


; menu functions ================================
Authorities:
	IfExist, C:\Program Files (x86)\Notepad++
	{
		Run, "C:\Program Files (x86)\Notepad++\notepad++.exe" %uhdlsubjectfile%
	}
	Else
	{
		Run, "C:\Program Files\Notepad++\notepad++.exe" %uhdlsubjectfile%
	}
Return

Report:
	IfExist, C:\Program Files (x86)\Notepad++
	{
		Run, "C:\Program Files (x86)\Notepad++\notepad++.exe" %reportfile%
	}
	Else
	{
		Run, "C:\Program Files\Notepad++\notepad++.exe" %reportfile%
	}
Return


Reload:
Reload

Exit:
ExitApp

GuiClose:
ExitApp
