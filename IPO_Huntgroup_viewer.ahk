{ 		; Application init for: IP Office Hunt Group viewer
AppName = IP Office Hunt Group viewer
Menu, Tray, Icon, imageres.dll, 195
#NoEnv
#SingleInstance force
SendMode Input
#InstallKeybdHook
#Persistent
SetTitleMatchMode, 1
SetWorkingdir %A_ScriptDir%
MoveMx = 0
MoveMp = 0
GuiToggle = 0
FileEncoding, CP28591 ; = iso-8859-1 = Avaya IP Office TFTP-filer. Se https://docs.microsoft.com/sv-se/windows/win32/intl/code-page-identifiers
If !FileExist(A_Workingdir . "/IPO_Huntgroup_viewer.ini")
	gosub CreateInifile
If !FileExist(IPO_Files)
	FileCreateDir, IPO_Files
FileDelete IPO_Files\_sysinfo.txt
ini = %A_Workingdir%/IPO_Huntgroup_viewer.ini
iniread, ipo, %ini%, init, currentipo
iniread, iplist, %ini%, init, iplist
iniread, X, %ini%, init, Xpos 
iniread, Y, %ini%, init, Ypos
iniread, language, %ini%, init, language

Gosub iniReadLanguage
SysGet, Mp, MonitorWorkArea , 1
SysGet, Mc, MonitorCount
if Mc > 1
	goto Monitor2Check
	goto Monitor1Check
}
TheGUI:
	{
	{ ; Main GUI
iniread, ipo, %ini%, init, currentipo
iniread, X, %ini%, init, Xpos 
iniread, Y, %ini%, init, Ypos
Gui Destroy
Gui Color, 181820, 303038
Gui font, s12 c80A0F4, Webdings
Gui Add, Text, x4 y8 gSelect, i
Gui font, bold s12 cC4C4D0, Microsoft Sans Serif
Gui Add, Text, y8 x+4 vNam backgroundtrans gSelect, % tit_

Gui font, norm s8 cE0E0E4, Microsoft Sans Serif
Gui Add, DDL , y7 xp w202 vchipo gChangeAddr, %hLoa%|%ipo%||%hAdd%|%iplist%
Gui Add, Text, y-3 x527 gReload_, ●
Gui font, cE0E0E4, Microsoft Sans Serif
if ipo = %hAdd%
	GuiControl, , Nam, %iniEdit%
Gui font, s12 norm cB0B0B4

Gui Add, Text, y8 x252 backgroundtrans, %hGrp%
Gui Add, Text, yp+0 xp+65 backgroundtrans, %hQue%
Gui Add, Text, yp+0 xp+35 backgroundtrans, Status
Gui Add, Text, yp+0 xp+125 backgroundtrans, %hMsg%
Gui font, s14 c707078
Gui Add, Text, x12 yp+7 backgroundtrans, ___________________________________________________
Gui font, s14 cC4C4C8
}
loop Parse, dList, `n, `r
	{
	row := A_LoopField
	pos_ := StrSplit(row, ",")
	NAMN := pos_[1]
	NUMR := pos_[2]
	VMSG := pos_[8]
	IF VMSG = 0
		VMSG := A_Space
	DNO := pos_[5]
	Gui Add, Text, w260 x12 yp+24 backgroundtrans vNam%A_Index%, % NAMN
	Gui font, s12 norm 
	Gui Add, Text, yp+2 x252 backgroundtrans vNum%A_Index%, % NUMR
	Gui font, s14 
	If (DNO = 2 )
		{ ; Group in Night service
		Gui font, cFF8888
		DNO := Nite
		QUE := "⛔"
		}
	If (DNO = 0 )
		{ ; Group in Out of service/fallback
		Gui font, cEE77BB
		DNO := Oos
		QUE := "❌"
		}
	If (DNO = 1 )
		{ ; Group in day mode
		DNO =
		Gui font, cFFFF00
		if pos_[4] = 0
			QUE =
		if pos_[4] > 0
			{
			QUE := pos_[4]
			DNO := inQ
			}
		}
	If NAMN =
		{
		QUE =
		DNO =
		VMSG =
		}
	Gui Add, Text, yp-2 xp+65 w50 backgroundtrans vQ%A_Index% , % QUE
	Gui Add, Text, yp+0 xp+30  w170 backgroundtrans vDNO%A_Index%, % DNO
	Gui font, s12 cEEEEA0
	If VMSG > 0
		Gui Add, Text, yp+2 xp+136 backgroundtrans vMs%A_Index%, % "🗨" VMSG
	; Gui font, s12 cC4C4C8
	if NUMR > 0
		{
		Gui font, s14 c606068
		Gui Add, Text, x12 yp+6 backgroundtrans, ___________________________________________________
		Gui font, cC4C4C8
		}
	} ; loop done
	GuiControl, Hide , chipo
Gui, Show, x%X% y%Y% w535, %AppName% : %tit_% @ %ipo%
SetTimer, GUIposition, 5000
return
}
ChangeAddr:
	{ 	; DropDown-menyn aktiverar denna rutin.
	GuiToggle = 0
	Gui 3: Destroy
	Gui, submit, NoHide
	GuiControl, , Nam, %connTry%
	If chipo = %hAdd%
		{
		SetTimer, TFTPget, Off
		GuiControl, , Nam, %iniEdit%
		loop 20
			{
			GuiControl, ,   Nam%A_Index%,
			GuiControl, ,   Num%A_Index%, 
			GuiControl, ,   Q%A_Index%, 
			GuiControl, ,   DNO%A_Index%, 
			GuiControl, ,   Ms%A_Index%, 
			}
		InfoHead = %Add0%
		i1 := A_Space
		i2 := Add1 " IPO_Huntgroup_viewer.ini"
		i3 := Add2
		i4 := A_Space
		i5 := A_Space
		Goto InfoGUI
		}
	If chipo = %hLoa%
		{
		loop 20
			{
			GuiControl, ,   Nam%A_Index%,
			GuiControl, ,   Num%A_Index%, 
			GuiControl, ,   Q%A_Index%, 
			GuiControl, ,   DNO%A_Index%, 
			GuiControl, ,   Ms%A_Index%, 
			}
		Gui 2: Destroy
		iniread, ipo, %ini%, init, currentipo
		iniread, iplist, %ini%, init, iplist
		GuiControl, , Nam, % ipo
		gosub iniReadLanguage
		FileDelete IPO_Files\_sysinfo.txt
		FileDelete IPO_Files\%tit_%GroupsList.txt
		gosub Whois
		Goto TheGUI
		}
	;iniwrite, %ipo%, %ini%, init, currentipo
	ipo = %chipo%
	Gui 2: Destroy
	gosub Whois
	gosub TheGUI
	Return
}
Select:
	{ 	; Klickade på texten.
	
	GuiControlGet, linktext, , Nam
	if linktext = No response
		{
		GuiControl, , Nam, Edit ini-file
		If !WinExist("IPO_Huntgroup_viewer.ini")
		run %A_Workingdir%/IPO_Huntgroup_viewer.ini
		Winactivate IPO_Huntgroup_viewer.ini
		Return
		}
	if linktext := tit_
		{
		GuiToggle := !GuiToggle
			if (GuiToggle = 0) {
			GuiControl, Hide , chipo
			Gui 3: Destroy
			Return
			}
		gosub InfoPage
		Return
		}
	iniread, ipo, %ini%, init, currentipo
	if ipo = %hAdd%
		goto ChangeAddr
goto TheGUI
}
Whois:
	{ 	; Kolla om IPO svarar på denna adress. TFTP GET who_is & licence_list. 
	iniread, language, %ini%, init, language
	gosub iniReadLanguage	
	if ipo = %hAdd%
		goto TheGUI
	FileDelete IPO_Files\_sysinfo.txt
	FileDelete IPO_Files\GroupsList.txt
	RunWait, %comspec% /c tftp.exe %ipo% GET nasystem/who_is3 IPO_Files\_sysinfo.txt,,Hide
	sleep 1000
	If !FileExist(A_Workingdir . "\IPO_Files\_sysinfo.txt")
	{
	SetTimer, TFTPget, Off
	Gosub NoAnswer
	FileDelete IPO_Files\GroupsList.txt
	Return
	}

fileread, sysinfo, IPO_Files\_sysinfo.txt
	si := StrSplit(sysinfo, "`""","")  ; Split by " and commas}.
	mac_ := si[1] . " " . si[2] 
	typ_ := si[3] . "" . si[4] 
		typ := StrSplit(typ_ , " ")
		if typ[3] = 500
		{
		Webmgr = https://%ipo%:8443/WebMgmtEE/WebManagement.html
		} else {
		Webmgr = https://%ipo%:7070/WebManagement/WebManagement.html
		}
	ver_ := si[9] . "" . si[10]
	nam_ := si[11] . " " . si[12]
	tit_ := si[12]
	RunWait, %comspec% /c tftp.exe %ipo% GET nasystem/licence_list "IPO_Files\%tit_%_Licence.txt",,Hide
	;fileappend, % si[4], sddasd.txt
	; if no mac then show text reason.
	If nam_ = 
	{
	SetTimer, TFTPget, Off
	FileDelete IPO_Files\%tit_%_GroupsList.txt
	Gosub NoAnswer
	Return
	} else {
	iniwrite, %ipo%, %ini%, init, currentipo
	
	goto TFTPget
	}
	goto TFTPget
Return
}
TFTPget:
	{	; Hämta grupplistan. Uppdateringsfrekvens, se "SetTimer, TFTPget, xxxx" längst ner i denna rutin
	GuiControl, , Nam, % tit_
	RunWait, %comspec% /c tftp.exe %ipo% GET nasystem/hunt_list "IPO_Files\%tit_%_GroupsList.txt",,Hide
	FileRead, dList , IPO_Files\%tit_%_GroupsList.txt
	if StrLen(dList) < 8
	{
	SetTimer, TFTPget, off
	GuiControl, , Nam, % tit_
	dList =
	(
	%NoHG%
	)
	Return
	}
	Sort, dList, D`n
loop Parse, dList, `n, `r
{
row := A_LoopField
pos_ := StrSplit(row, ",")
NAMN := pos_[1]
VMSG := pos_[8]
IF VMSG = 0
	VMSG := A_Space
DNO := pos_[5]
	If (DNO = 2 )
		{ ; Group in Night service
		GuiControl, +cFF8888 +Redraw, Q%A_Index%
		GuiControl, +cFF8888 +Redraw, DNO%A_Index%
		DNO := Nite
		QUE := "⛔"
		}
	If (DNO = 0 )
		{ ; Group in Out of service/fallback
		GuiControl, +cEE77BB +Redraw, Q%A_Index%
		GuiControl, +cEE77BB +Redraw, DNO%A_Index%
		DNO := Oos
		QUE := "❌"
		}
	If (DNO = 1 )
		{ ; Group in day mode
		DNO =
		GuiControl, +cFFFF00 +Redraw, Q%A_Index%
		GuiControl, +cFFFF00 +Redraw, DNO%A_Index%
		if pos_[4] = 0
			QUE =
		if pos_[4] > 0
			{
			QUE := pos_[4]
			DNO := inQ 
			}
		}
If NAMN =
	{
	QUE := A_Space
	DNO := A_Space
	VMSG := A_Space
	}	
GuiControl, ,   Q%A_Index%, % QUE
GuiControl, , DNO%A_Index%, % DNO
If VMSG > 0
	GuiControl, , Ms%A_Index%, % "🗨" VMSG
	

}
GuiControl, , Nam, % tit_
SetTimer, TFTPget, 2000
Return
}
Monitor2Check:
	{ 	; Check if the Xpos and Ypos in ini-file fits in Monitor2. If not, check if pos fits Monitor1.
	SysGet, Mx, MonitorWorkArea , 2
	if X between %MxLeft% and %MxRight%
	{
	MoveX = 0
	} else {
	MoveX = 1
	}
	if Y between %MxTop% and %MxBottom%
	{
	MoveY = 0
	} else {
	MoveY = 1
	}
	MoveMx := MoveX + MoveY
if MoveMx > 0
	{
	;msgbox, 0, Position: ,ini file settings did not fit in Monitor2 - check Monitor1, 1
	goto Monitor1Check
	} else {
	;msgbox, 0, Position: ,Yes open GUI on Monitor2 , 1
	Gosub Whois
	goto TheGUI
	}}
Monitor1Check:
	{ 	; Check if the Xpos and Ypos in ini-file fits in Monitor1. If not, reset position to x10,y10.
	if X between %MpLeft% and %MpRight%
	{
	MoveX = 0
	} else {
	MoveX = 1
	}
	if Y between %MpTop% and %MpBottom%
	{
	MoveY = 0
	} else {
	MoveY = 1
	}
	MoveMp := MoveX + MoveY
if MoveMp > 0
	{ ; msgbox, 0, Position: ,did not fit in Monitor1 - store new X & Y positions, 1
	X = 10
	iniwrite, %X%, %ini%, init, Xpos
	Y = 10
	iniwrite, %Y%, %ini%, init, Ypos
	} 
	Gosub Whois
	goto TheGUI
}
GetUsers:
	{	; TFTP GET user_list8 Spara med citationstecken annars fungerar inte kundnamn med blanksteg
	uXList =
	uList =
	rece = 0
	powr = 0
	twrk = 0
	offi = 0
	mobi = 0
	WinGetPos, sX, sY, Width, Height, %AppName%
	sY := sY + 100
	SplashImage, , x%sX% y%sY% w%Width% CWFFFF00 b fs18, %dlUsr%
	RunWait, %comspec% /c tftp.exe %ipo% GET nasystem/user_list8 "IPO_Files\%tit_%_Users.txt",,Hide
	FileRead, users , IPO_Files\%tit_%_Users.txt
	FileDelete, IPO_Files\%tit_%_Users.csv
	loop, parse, users, `n, `r
		maxU := A_Index - 1
	uTabl =
	t_ := "<tr><td><a href='tel:"
	m_ := "<a href='mailto:"
	a := "'>"
	t := "</a></td><td>"
	D = </td><td>
	R = </td></tr>
	c = `,
	s = `;		
	Loop Parse, users,`n
		{
		Apps =
		WpM =
		WpD =
		RemW =
		1xTc =
		1xWc =
		OneX =
		Teams =
		Old =		
		IF A_Index > %maxU%
			Break
		uCol := StrSplit(A_LoopField, ",")
		uId := uCol[1]
		uXt := uCol[2]
		uNa := uCol[3]
		uLc := uCol[4]
		{
		if uCol[4] = 1
			uLc =
		if uCol[4] = 2
			{
			uLc = Teleworker
			twrk += 1
			}
		if uCol[4] = 3
			{
			uLc = Mobile worker
			mobi += 1
			}
		if uCol[4] = 4
			{
			uLc = Power
			powr += 1
			}
		if uCol[4] = 5
			{
			uLc = Office
			offi += 1
			}
		if uCol[4] = 8
			uLc = Non-Licenced
		}
		uFea := uCol[5]
		Rec =
		If Mod(uFea, 2) != 0 ; if Odd number; Receptionist is activated. Remove 1 from applications sum.
			{
			Rec := "Rec"
			rece += 1
			uFea := uFea - 1
			} else {
			}
		{ ;					list features and licences per User
			if uFea > 512 
				{
				Teams := ". T."
				uFea := uFea - 512
				}
			if uFea > 62
				{
				1xWc := " +Conf" 
				uFea := uFea - 64
				}
			if uFea > 30
				{
				WpM := uCol[4]= 4 ? "+m" : ""
				uFea := uFea - 32
				}
			if uFea > 14
				{
				WpD := uCol[4]= 4 or 5 ? "Workplace" : ""
				uFea := uFea - 16
				}
			if uFea > 6 
				{
				RemW := uCol[4]= 1 or 2 or 3 or 5 ? ". rX" : ""
				uFea := uFea - 8
				}
			if uFea > 2 
				{
				1xTc := " +comm"
				uFea := uFea - 4
				}
			if uFea = 2
				{
				Old := ". pc"
				uFea := uFea - 2
				}
			Apps := WpD . WpM . RemW . Old . Teams
			OneX := uCol[6]=1 ? "OneX" . 1xWc . 1xTc : ""
			uEm := uCol[8]
			}
		uXList :=uXList uXt s uId s uNa s uEm s Rec s uLc s Apps s OneX "`n"
		uList  := uList uXt c uId c uNa c uEm c Rec c uLc c Apps c OneX "`n"
		uTabl  := uTabl t_ uXt a uXt t uNa D m_ uEm a uEm t uId D Rec D uLc D Apps D OneX R "`n"
		}
		
Goto CreateUpage
}
CreateUpage:
	{
	SplashImage, Off
	p = `%
	CSVhead = % "Extension" s "User ID" sp s "Full Name" sp sp s "eMail Address" sp sp s "rConsole" s "Licence" sp s "Application access (enabled):" sp sp s "One-X portal" s "Press Ctrl+T to format as a table  `n"
	FileAppend, %CSVhead%%uXList% , IPO_Files\%tit_%_Users.csv
	FileDelete, IPO_Files\%tit_%_Users.html
uHTML :=
(
"<HTML>
<head><style>
	body { margin: 0; font-family: Arial, sans-serif; font-size: calc(8px + .5vw); background-color:#112; color:#AAB; cursor:default; }
	table { top: 30px; margin-left: 8px; border-collapse: collapse; width:98.5%; text-align: left;}
	tr:hover {background-color: #335;}
	th { background-color:#AAB; color:#112;}
	td, th {padding-left: 2px; padding-right: 2px;  border-top:solid 1px grey; border-left:solid 1px #223; border-right:solid 1px #223; border-top:solid 1px grey;}
	td:nth-child(3) { font-size: calc(7px + .4vw); }
	td:nth-child(4) { font-size: calc(7px + .4vw); }
	td:nth-child(7) { font-size: calc(7px + .4vw); }
	td:nth-child(8) { font-size: calc(7px + .4vw); }
	a {text-decoration: none; color:#AAE; cursor:pointer;}
	.navbar { overflow: hidden; background-color: #f0f1f2; position: fixed; top: 0; }
	.navbar a { float: left; display: block; text-decoration: underline; color: #335; text-align: center; padding: 4px 6px; font-size: 17px; }
	.navbar a:hover { background: #ddd; color: black; }
</style></head>
<body><div class='navbar'>"
)

;uHead = <tr><thhref="#"> ( Used licences: Power: %powr% Office worker: %offi% Teleworker: %twrk%. Mobile worker: %mobi%. Receptionist: %rece% )</tr>

uHead = <tr><th>Nr</th> <th>Full Name</th> <th>eMail</th> <th>ID</th> <th>Rec.</th> <th>Lic</th> <th>Appl</th> <th>OneX P</th> </tr>

a_WMgr = <a href=" %Webmgr% " target="_blank">IP Office Web Manager</a>
a_Ucsv = <a href="file:///%A_Workingdir%/IPO_Files/%tit_%_Users.csv  " target="_blank">Excel(CSV)</a>
a_Utxt = <a href="file:///%A_Workingdir%/IPO_Files/%tit_%_Users.txt  " target="_blank">`[View as text`]</a>
	
	
	FileAppend, % uHTML a_WMgr a_Ucsv a_Utxt "</div><p>User extension list</p><table>`n" uHead uTabl "`n</table></body></html>" , IPO_Files\%tit_%_Users.html
	
	
	
	Return
	}
WebUlist:
	{
	run,  IPO_Files\%tit_%_Users.html
	Return	
	}
ExcelUlist:	
	{
	run,  IPO_Files\%tit_%_Users.csv
	return
	}
DirList:
	{ 	; TFTP GET dir_list.  Spara med citationstecken annars fungerar inte kundnamn med blanksteg
	if ipo = %hAdd%
		Return
	dLi =
	WinGetPos, sX, sY, Width, Height, %AppName%
	sY := sY + 100
	SplashImage, , x%sX% y%sY% w%Width% CWFFFF00 b fs18, %dlDir%
	dHead = OPEN IN EXCEL IF THIS IS TOO LONG `n################################`n
	dTail := "########### END OF LIST ###########`n"
	RunWait, %comspec% /c tftp.exe %ipo% GET nasystem/dir_list "IPO_Files\%tit_%_SysDir.txt",,Hide
	FileRead, directory , IPO_Files\%tit_%_SysDir.txt
	FileDelete, IPO_Files\%tit_%_SysDir.csv
		loop, parse, directory, `n, `r
		maxD := A_Index - 1
		Loop Parse, directory,`n
		{
		IF A_Index > %maxD%
			Break
		dCol := StrSplit(A_LoopField, ",")
		dNa := dCol[1]
		dNu := dCol[2]
		dLi := dLi dNa "`;" dNu "`n"
		}	
	TheDir = Name/Feature %sp% `; Number/Code `; Press Ctrl+T to format as a table `n%dLi%
	FileAppend, %TheDir% , IPO_Files\%tit_%_SysDir.csv
	SplashImage, Off
	
;	Gui 5: Color, 808088, C0C0C8
;	Gui 5: font, s10 ce0e8FF bold Underline
;	Gui 5: Add, Text, gExcelUlist, Open in Excel
;	Gui 5: font, s10 c000000 norm
;	Gui 5: Add, Edit, R20 w1100, % lHead uHead uLine uList lTail
;	Gui 5: Show, , %tit_% - System Directory phone book
	
	
	
	
	Msgbox, 4, Phonebook directory, Open the `n`nSystem phone book `n`nin Excel?, 0
	IfMsgBox, Yes
	run,  IPO_Files\%tit_%_SysDir.csv
	IfMsgBox, No
	Msgbox, 0, System phone book,  % lHead dLi lTail, 0
	Return
	}
InfoGUI:
	{	; info popup. Add IPO (Edit inifile)
	gosub 2GuiClose
SetTimer, TFTPget, Off
If !WinExist("IP Office Hunt Group viewer")
	{
	X=10
	Y=10
	Yb=83
	Wb=584
	} else {
	WinGetPos, X, Y, Width, Height, %AppName%
	Yb := Y + 62
	Ye := Y + 278
	Wb := Width - 6
	}
	Gui 2: +ToolWindow +AlwaysOnTop
	Gui 2: Color, 303038
	Gui 2: font, s16 norm cF0B0B4
	Gui 2: Add, Text, x10 backgroundtrans vIG1, % i1
	Gui 2: Add, Text, x10 backgroundtrans vIG2, % i2 
	Gui 2: Add, Text, x10 backgroundtrans vIG3, % i3
	Gui 2: Add, Text, x10 backgroundtrans vIG4, % i4
	Gui 2: Add, Text, x10 backgroundtrans vIG5, % i5
	Gui 2: Show,x%X% y%Yb% w%Wb%, %A_Space% %InfoHead% ; 
	If !WinExist("IPO_Huntgroup_viewer.ini")
	run %A_Workingdir%/IPO_Huntgroup_viewer.ini
	WinWait, IPO_Huntgroup_viewer.ini - Notepad
	WinMove, %X%, %Ye%
	Winactivate IPO_Huntgroup_viewer.ini
Return
}
InfoPage:
	{	; IPO type software version & installed licences
	gosub GetUsers
	maxL =
	Licences =
	Fileread, lili, %A_Workingdir%\IPO_Files\%tit_%_Licence.txt
	sort, lili, N
	loop, parse, lili, `n, `r
	maxL := A_Index - 1
	Gui 3: Destroy
	Gui 3: -Caption Border
	Gui 3: Color, 303038
	Gui 3: font, s12 norm cB0B0B4
	Gui 3: add, text, x6 y9, System information
	Gui 3: font, s10
	Gui 3: add, text, x+8 y11 , %typ_% %ver_%
	Gui 3: font, s8 c80A0F4
	Gui 3: Add, Text, x427 y7 gUserList, Users
	Gui 3: Add, Text, x+4 yp gDirList, SysDir
	Gui 3: font, s10 c80A0F4, Webdings
	Gui 3: Add, Text, x435 yp+10 gWebUlist BackgroundTrans, i
	Gui 3: Add, Text, x+18 yp gDirList BackgroundTrans, i
	Gui 3: font, s10 cB4B050
	Gui 3: Add, Text, X+18 yp gFiles BackgroundTrans, Ì
	Gui 3: font, s8 , Microsoft Sans Serif
	Gui 3: Add, Text, xp-2 y7 gFiles BackgroundTrans, files
	Gui 3: font, s9 c8098F0 Underline
	Gui 3: Add, Text, x10 yp+30 gBrowse BackgroundTrans, Web Manager
	Gui 3: Add, Text, x+10 yp gPlatfo BackgroundTrans, Platform adm
	Gui 3: Add, Text, x+10 yp gUserPortal BackgroundTrans, User Portal
	Gui 3: Add, Text, x+10 yp g46xx BackgroundTrans, 46xxsettings
	Gui 3: Add, Text, x+10 yp gOneXPortal BackgroundTrans, OneX Portal
	Gui 3: Add, Text, x+10 yp gWebCollab BackgroundTrans, Web Collab
	Gui 3: font, s9 cB0B0B4 norm
	Gui 3: add, text, x10 yp+20 BackgroundTrans, % sp sp sp sp sp sp sp sp sp "          max  /  (users)"
	Gui 3: font, s11 bold, Courier New 
	Gui 3: add, text, x6 yp BackgroundTrans, __________________________________________________________
	Loop Parse, lili,`n
		{
		Lr := StrSplit(A_LoopField, ",")
		LiID := Lr[1]
		LiAM := Lr[4]
		IF A_Index > %maxL%
		Break
			Loop, parse, Ltypes, `n
			{
			Lt := StrSplit(A_LoopField, "=")
			LtID := Lt[1]
			Ltxt := Lt[2]
				If LiID = %LtID%
				{
				Gui 3: font, s11 bold c90E090, Courier New 
				Gui 3: add, text, x10 y+6, % Ltxt
				Gui 3: add, text, x450 yp, % LiAM
				
				if Lt[1] = 16
					Gui 3: add, text, x485 yp, (%rece%)
				if Lt[1] = 65
					Gui 3: add, text, x485 yp, (%twrk%)			
				if Lt[1] = 65
					Gui 3: add, text, x485 yp, (%twrk%)
				if Lt[1] = 66
					Gui 3: add, text, x485 yp, (%mobi%)
				if Lt[1] = 67
					Gui 3: add, text, x485 yp, (%powr%)
				if Lt[1] = 69
					Gui 3: add, text, x485 yp, (%offi%)
				; skanna extensions list?  21=3rd pty endpoints, 62=Avaya endpoints 
				Gui 3: font, c808084,
				Gui 3: add, text, x6 yp+8 BackgroundTrans, __________________________________________________________
				}
			}	
		}
	Gui 3: font, s10 c808084
	Gui 3: add, text, y+5 ,
	Gui 3: add, text, x10 yp, (Old obsolete licences will show here if not deleted in the IPO)
	Gui 3: add, text, yp+8 BackgroundTrans,	
	WinGetPos, X, Y, Width, Height, %AppName%
	Yc := Y + 62
	Xc := X +2
	Wc := Width - 6
	GuiControl, Show, chipo
	Gui 3: Show,x%Xc% y%Yc% w%Wc%, %A_Space% System- & Licence info for %tit_%:
	Return
}
UserList:
	{ 	; 
	if ipo = %hAdd%
		Return
	uHead := "( Used licences: Power:" powr ". Office worker:" offi ". Teleworker:" twrk ". Mobile worker:" mobi ". Receptionist:" rece " )`n"
	uLine = ―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――`n
	Sort, uList, D`n
	Gui 4: Color, 808088, C0C0C8
	Gui 4: font, s10 ce0e8FF bold 
	Gui 4: Add, Text, x8 , Open in: 
	Gui 4: font, Underline
	Gui 4: Add, Text, yp x+10 gExcelUlist, Excel
	Gui 4: Add, Text, yp x+10 gWebUlist, Web browser
	Gui 4: font, s10 c000000 norm, Arial
	Gui 4: Add, Edit, x8 R20 w1100 , % lHead uHead uLine uList lTail
	WinGetPos, X, Y, Width, Height, %AppName%
	Yc := Y + 100
	Xc := X +7
	Gui 4: Show, x%Xc% y%Yc%, User Extensions list
	return
	}
Browse:
	{	; browse to IP addr. link in Infopage
	Run, %Webmgr%
	; msgbox, %Webmgr%
	Return
	}
Platfo:	
	{	; browse to system platform port 7071. link in Infopage
	run, https://%ipo%:7071
	Return
	}
46xx:
	{
	run, https:///%ipo%/46xxspecials.txt
	run, https:///%ipo%/46xxsettings.txt
	Return
	}
UserPortal:
	{
	run, https://%ipo%:7444/userportal/index.html
	Return	
	}
OneXPortal:
	{
	run, https://%ipo%:9443/onexportal.html
	Return	
	}
WebCollab:
	{
	run, https://%ipo%:9443/meeting/login.jsp
	Return
	}
NoAnswer:
	{
	GuiControl, , Nam, No response!
	InfoHead := NoR0 " " ipo
	if currentipo =
	{
	i1 := NoIP
	} else {
	i1 := NoR1
	}
	i2 := NoR2 " " ipo
	i3 := NoR3
	i4 := NoR4
	i5 := NoR5
	iniwrite, %hAdd%, %ini%, init, currentipo
	FileDelete IPO_Files\%tit_%GroupsList.txt
	iniread, iplist, %ini%, init, iplist
Goto InfoGUI
}
CreateInifile:
	{ 	; Create ini-file first time.
	FileDelete IPO_Huntgroup_viewer.ini
	msgbox, 0, There was no ini-file found., First time use? `n `nCreating default ini-file..., 2
	FileAppend,
(
###                  Inifile for IPO_Huntgroup_viewer:                ###
###    Select (Reload) in the application after saving changes here   ###
### Hunt groups configured "Exclude From Directory" will not be shown ###

[init]
# currentipo = the latest selected IP address. language = EN or SE.
currentipo =
Xpos = 10
Ypos = 10
language = SE

### iplist: add selectable IP Office addresses here, separate with pipe '|' character.
iplist = 

[labels]
### Information texts, also shown in popups ###
seAdd0 = Lägga till IP Office-adress(er):
seAdd1 = Redigera och spara filen
seAdd2 = Välj sedan (Uppdatera lista) i menyn.
seNoIP = Ange adress till IP Office i 'currentipo'
seNoGr = Växeln har inga ringgrupper.
seNoR0 = Inget svar från vald adress:
seNoR1 = Felsökning: currentipo i ini-filen. Rätt IP-adress?
seNoR2 = - Avaya IPO-adressen ska svara på ping.
seNoR3 = - TFTP ska vara aktiverat i IPO/security settings.
seNoR4 = - TFTP-klient ska vara aktiverad i din PC.
seNoR5 = - TFTP(port69) ej blockat i brandvägg mellan PC-IPO.
seNamE = Redigera ini-fil
seNamC = Ansluter...
seNamT = kontrollerar IP...

enAdd0 = Adding IP Office address(es):
enAdd1 = Edit and save the file
enAdd2 = Then select (Reload) from the menu to update.
enNoIP = Set an IP Office address in 'currentipo'.
enNoGr = No hunt groups in this PBX.
enNoR0 = No response from selected IP address:
enNoR1 = Troubleshooting: Correct currentipo in ini file?
enNoR2 = - The IPO address must be ping-able.
enNoR3 = - TFTP must be activated in the IPO/security.
enNoR4 = - TFTP-client must be activated on this PC.
enNoR5 = - TFTP(p69) not be blocked if firewall between PC-IPO.
enNamE = Edit ini-file
enNamC = connecting...
enNamT = Trying IP address..

### Hunt group Service status texts ###
seinQ = samtal i kö
seNite= Natt/Lunch
seOos = Stängd/omstyrd

eninQ = call(s) in queue
enNite= Night/lunch
enOos = Closed/Diverted

### Column headers: "Msg"=Voicemail messages ###
seHGrp = Grupp
seHQue = i kö
seHMsg = Medd
seHAdd = ( Lägg till IPO )
seHLoa = ( Uppdatera lista )

enHGrp = Group
enHQue = inQ
enHMsg = Msg
enHAdd = ( Add IPO address )
enHLoa = ( Update list )
),	IPO_Huntgroup_viewer.ini
	Return
}

iniReadLanguage:	
	{	; ini-file check key; language (EN or SE)
	sp =
	loop, 15
	sp := sp . A_Space
if language = SE
{
iniread, inQ,  %ini%, labels, seinQ
iniread, Nite, %ini%, labels, seNite
iniread, Oos,  %ini%, labels, seOos
iniread, hGrp, %ini%, labels, seHGrp
iniread, hQue, %ini%, labels, seHQue
iniread, hMsg, %ini%, labels, seHMsg
iniread, hAdd, %ini%, labels, seHAdd
iniread, hLoa, %ini%, labels, seHLoa
iniread, Add0, %ini%, labels, seAdd0
iniread, Add1, %ini%, labels, seAdd1
iniread, Add2, %ini%, labels, seAdd2
iniread, NoIP, %ini%, labels, seNoIP
iniread, NoR0, %ini%, labels, seNoR0
iniread, NoHG, %ini%, labels, seNoGr
Loop 5
	iniread, NoR%A_Index%, %ini%, labels, seNoR%A_Index%
iniread, iniEdit, %ini%, labels, seNamE
iniread, conn, %ini%, labels, seNamC
iniread, connTry, %ini%, labels, seNamT	
lHead =  Välj Open in Excel för att få sorterade kolumner (Tryck Ctrl+T+Enter i excel).      Välj Web Browser (nummer och epost klickbart).  `n
lTail := "########## SLUT PÅ LISTAN ##########`n"
dlDir = Hämtar system-telefonkatalog...
dlUsr = Läser Anknytningslista och licenser...

} else {
iniread, inQ,  %ini%, labels, eninQ
iniread, Nite, %ini%, labels, enNite
iniread, Oos,  %ini%, labels, enOos
iniread, hGrp, %ini%, labels, enHGrp
iniread, hQue, %ini%, labels, enHQue
iniread, hMsg, %ini%, labels, enHMsg
iniread, hAdd, %ini%, labels, enHAdd
iniread, hLoa, %ini%, labels, enHLoa
iniread, Add0, %ini%, labels, enAdd0
iniread, Add1, %ini%, labels, enAdd1
iniread, Add2, %ini%, labels, enAdd2
iniread, NoIP, %ini%, labels, enNoIP
iniread, NoR0, %ini%, labels, enNoR0
iniread, NoHG, %ini%, labels, enNoGr
Loop 5
	iniread, NoR%A_Index%, %ini%, labels, enNoR%A_Index%
iniread, iniEdit, %ini%, labels, enNamE
iniread, conn, %ini%, labels, enNamC
iniread, connTry, %ini%, labels, enNamT
lHead = Select Open in Excel for column sorting (Press Ctrl+T+Enter in excel).     Web Browser (numbers and emails clickable). `n
lTail := "########### END OF LIST ###########`n"
dlDir = Downloading System phone directory...
dlUsr = Reading User Extension list and licenses...
}
Ltypes =
(
2=CTI link pro       (3rd party CTI)
3=Wave User          (3rd party CTI)

16=Receptionist       (operator Softconsole)
18=VMPro channels     (VoiceMail Pro)
21=3rd pty endpoints  (for ATAbox, DECT etc)
31=VMPro TTS          (Scansoft)
36=IPsec Tunneling

41=Voice Networking   (IPO-to-IPO network)
45=SIP Trunk channels (Lines, integration)
62=Avaya endpoints    (Deskphones)
65=TeleWorker         (IP500 only)
66=Mobile Worker      (IP500 only)
67=Power User         (Avaya Softphones)
69=Office Worker      (Avaya Softphones)

86=Essential Edition  (Embedded VoiceMail)
91=VMPro TTS Pro      (Text-To-Speech)
92=Preferred Edition  (VoiceMail Pro)
97=Server Edition
98=Server Ed.Upgrade

108=Web Collaboration  (Conferencing Manager)
115=ACCS               (Avaya Contact Center Select)
116=Server Ed R10
123=Server Ed.Upgrade
126=Media Manager      (Call Recording Manager)
)

Return	
}
3GuiClose:
	{ ; Reset Gui toggle counter to 0
	GuiToggle = 0
	GuiControl, Hide, chipo
	Gui, 3:destroy
	Return
	}
2GuiClose:
	{ 	; closes the second Gui window
	GuiControl, Hide, chipo
	Gui, 2:destroy
	Return
	}
GuiClose:
	{	; Exitapp
	WinGetPos, X, Y, Width, Height, %AppName%
	iniwrite, %X%, %ini%, init, Xpos 
	iniwrite, %Y%, %ini%, init, Ypos 
	ExitApp
	}
GUIposition:
	{	; Get GUIposition. (settimer in GUI)
	WinGetPos, X, Y, Width, Height, %AppName%
	iniwrite, %X%, %ini%, init, Xpos 
	iniwrite, %Y%, %ini%, init, Ypos
	Return
	}
Files:
Run, IPO_Files\
Return
Reload_:
Reload