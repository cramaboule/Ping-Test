#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=green.ico
#AutoIt3Wrapper_Res_Description=Ping Test By Cramaboule
#AutoIt3Wrapper_Res_Fileversion=1.0.6.3
#AutoIt3Wrapper_Res_ProductVersion=1.0.6.3
#AutoIt3Wrapper_Res_Icon_Add=green.ico
#AutoIt3Wrapper_Res_Icon_Add=red.ico
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/reel
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region    ;Timestamp =====================
#    Last complie at : 2025/01/08 07:58:41
#EndRegion ;Timestamp =====================

#comments-start

	AutoIt Version: 3.3.16.1
	Author:         Cramaboule

	History:
	Ping Test 1.0.6.3	08.01.2025
						Changed: Improve the _GetMachine, fixing bugs simplify making Gui.
						Put the code up to date
	Ping Test 1.6.2	Changed: GUI for W11 16.05.2022
	Ping Test 1.6.1	Changed: Tray improvement
					Changed: Initial display
	Ping Test 1.6	Added: ini file for the setting (read and write) : Idea from GreenCan
					Added: 2 pings for the Outside Network if we get an error: Idea from GreenCan
	Ping Test 1.5	Removed: grip.DLL but:
					Add: icons into exe file using AutoIt3Wrapper
					Add: Settings
					Changed: Improved code
					Changed: Display (Thanks to Melba23)
	Ping Test 1.4   Fixed: To work with AutoIt v3.3.0.0
	Ping Test 1.4   Fixed: The TraySetIcon due to the change of Autoit
					Fixed: Not refreching when a machine came back
					online on "hide good ping" mode.
	Ping Test 1.3	Rewrite from scratch (almost !)
					Add refresh button !
	Ping Test 1.2   Add tray icons (from grip.dll)
	Ping Test 1.1	Add switch for file for batch file:
					Pingtest1.1.exe [file]
					Open witout prompting
	Ping Test 1.0	First realese


	Options: 		Command line: open host file

	Settings: 		Ping Time: 		Set the 'Timer' when all machine will be pinged min 10 Sec
	Inside Network:	Set the IP of the Inside Network i.e. 192.168. This, to set
	timer differtly for the Inside and the Outside
	Time out:		Set the time out for the inside network


#comments-end

#include <GUIConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Constants.au3>
#include <Misc.au3>
#include <File.au3>

Global $Nb_of_machines = 1, $button_hide = 1, $button_show = 1, $ShowAllPing = 1, $Nb_of_ping = 0, $f = 0
Global $machine[60][5], $machine_not_good[60][5], $Input[60][5], $seq[100], $Winposping[4]
Global $noping, $head, $file, $exititem, $showgooditem, $gui, $winstate, $state
Global $pPing, $Ttimer, $Tflag, $Tdiff, $RedoGui, $NewMachine, $OldMachine, $mgg, $oldnoping, $messtray
Global $trayrefresh, $showframe, $winstatetray

; head--------------------------------------------------
$head = "Ping Test 1.0.6.3"

#Region Ini File
Global $TWait = IniRead(@ScriptDir & "\pingtest.ini", "Settings", "PingTime", 30) * 1000
Global $IP = IniRead(@ScriptDir & "\pingtest.ini", "Settings", "InsideNetworkIP", "")
Global $Len = StringLen($IP), $button_refresh
Global $TOut_In = IniRead(@ScriptDir & "\pingtest.ini", "Settings", "INW_TimeOut", 200)
Global $TOut_Out = IniRead(@ScriptDir & "\pingtest.ini", "Settings", "ONW_TimeOut", 10) * 1000

If Not FileExists(@ScriptDir & "\pingtest.ini") Then
	IniWrite(@ScriptDir & "\pingtest.ini", "Settings", "PingTime", ($TWait / 1000))
	IniWrite(@ScriptDir & "\pingtest.ini", "Settings", "InsideNetworkIP", $IP)
	IniWrite(@ScriptDir & "\pingtest.ini", "Settings", "INW_TimeOut", $TOut_In)
	IniWrite(@ScriptDir & "\pingtest.ini", "Settings", "ONW_TimeOut", ($TOut_Out / 1000))
EndIf
#EndRegion Ini File

#Region Tray Menu
Opt("TrayMenuMode", 1) ; Default tray menu items (Script Paused/Exit) will not be shown.
$showframe = TrayCreateItem("Show/Hide " & $head)
$trayrefresh = TrayCreateItem("Refresh")
$showgooditem = TrayCreateItem("Show/Hide Good Pings")
$TraySettings = TrayCreateItem("Settings")
$exititem = TrayCreateItem("Exit")
TrayItemSetState($showframe, $TRAY_DEFAULT)
If @Compiled Then
	TraySetIcon(@ScriptFullPath, -5)
Else
	TraySetIcon(@ScriptDir & "\pingtest1.6.3.exe", -5)
EndIf
TraySetState()
TraySetClick(16)
#EndRegion Tray Menu

If $CmdLine[0] And $CmdLine[1] <> "" Then
	$file = $CmdLine[1]
Else
;~ 	$file = FileOpenDialog($head, "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "All (*.*)", 1, "hosts")
	$file = @ScriptDir & '\hosts.txt'
	If $file = "" Then Exit
EndIf

_GetMachine()

_CreateGui(1)
GUISetState(@SW_SHOWNORMAL, $gui)
$Winposping = WinGetPos($head)
$winstatetray = 1
$RedoGui = 0
_Ping() ; for the first time

While 1
	$mgg = GUIGetMsg()
	Select
		Case $mgg = $GUI_EVENT_CLOSE
			Exit
		Case $mgg = $button_refresh
			$RedoGui = 1
		Case $mgg = $button_show Or $mgg = $button_hide
			$RedoGui = 1
			$ShowAllPing = Not ($ShowAllPing)
		Case $mgg = $GUI_EVENT_MINIMIZE
			$winstatetray = 0
			GUISetState(@SW_HIDE, $gui)
	EndSelect
	$traymsg = TrayGetMsg()
	Select
		Case $traymsg = $exititem
			Exit
		Case $traymsg = $showframe
			If $winstatetray = 1 Then
				GUISetState(@SW_HIDE, $gui)
				$winstatetray = Not ($winstatetray)
			Else
				GUISetState(@SW_SHOWNORMAL, $gui)
				$winstatetray = Not ($winstatetray)
			EndIf
		Case $traymsg = $showgooditem
			$RedoGui = 1
			$ShowAllPing = Not ($ShowAllPing)
		Case $traymsg = $trayrefresh
			$RedoGui = 1
		Case $traymsg = $TraySettings
			Call("GUISettings")
	EndSelect
	;
	If $pPing >= 1 Then
		If $ShowAllPing Then
			GUICtrlSetBkColor($Input[$Nb_of_ping][1], 0x00ff00) ; green
			GUICtrlSetBkColor($Input[$Nb_of_ping][2], 0x00ff00) ; green
		EndIf
	Else
		If $f = 0 Then
			$noping += 1
			If @Compiled Then
				TraySetIcon(@ScriptFullPath, -6)
			Else
				TraySetIcon(@ScriptDir & "\pingtest1.6.3.exe", -6)
			EndIf
			$f = 1
			$machine_not_good[$noping][1] = $machine[$Nb_of_ping][1]
			$machine_not_good[$noping][2] = $machine[$Nb_of_ping][2]
			If $ShowAllPing Then
				GUICtrlSetBkColor($Input[$Nb_of_ping][1], 0xff0000) ; red
				GUICtrlSetBkColor($Input[$Nb_of_ping][2], 0xff0000) ; red
			Else
				GUICtrlSetBkColor($Input[$noping][1], 0xff0000) ; red
				GUICtrlSetBkColor($Input[$noping][2], 0xff0000) ; red
			EndIf
		EndIf
		If $noping > 1 Then
			$messtray &= @CRLF & $machine_not_good[$noping][1] & " " & $machine_not_good[$noping][2]
		Else
			$messtray &= $machine_not_good[$noping][1] & " " & $machine_not_good[$noping][2]
		EndIf
	EndIf
	$Tdiff = TimerDiff($Ttimer)
	If $RedoGui = 1 Then _CreateGui()
	If $Tdiff >= $TWait Or $RedoGui = 1 Or $Nb_of_ping > 0 Then
		_Ping()
	EndIf
	$state = WinGetState($head)
	If BitAND($state, 2) And BitAND($state, 8) Then ; 2 = visible  8 = active
		$Winposping = WinGetPos($head)
	EndIf
WEnd

Func _GetMachine()
	$NewMachine = ""
	$Nb_of_machines = 1
	Local $aLines
	If _FileReadToArray($file, $aLines) Then
		For $i = 1 To $aLines[0]
			;--------------- Get ride of the 'SPACE' and '#'
			If StringLeft($aLines[$i], 1) = "#" Then ContinueLoop
			$aLines[$i] = StringReplace($aLines[$i], Chr(9), " ")
			$aLines[$i] = StringStripWS($aLines[$i], 7)
			$seg = StringSplit($aLines[$i], " ")
			If $seg[0] >= 2 Then
				$machine[$Nb_of_machines][1] = $seg[1]
				$machine[$Nb_of_machines][2] = StringStripWS(StringReplace($aLines[$i], $seg[1], '', 1), 7)
				$NewMachine &= $machine[$Nb_of_machines][1] & " " & $machine[$Nb_of_machines][2] & " "     ;to record any changes in file
				$Nb_of_machines += 1
			EndIf
		Next
	EndIf
	If $NewMachine <> $OldMachine Then $RedoGui = 1
	$OldMachine = $NewMachine
EndFunc   ;==>_GetMachine

Func _Ping()
	$Nb_of_ping += 1
	;
	If $Nb_of_ping = 1 Then
		_GetMachine()
		If $RedoGui = 1 Then
			$oldnoping = $noping
			_CreateGui()
		EndIf
		$noping = 0
		$messtray = ""
	EndIf
	;
	$TOut = $TOut_Out
	If StringLeft($machine[$Nb_of_ping][1], $Len) = $IP Then $TOut = $TOut_In
	If $ShowAllPing Then
		GUICtrlSetBkColor($Input[$Nb_of_ping][1], 0xffff00) ; yellow
		GUICtrlSetBkColor($Input[$Nb_of_ping][2], 0xffff00) ; yellow
	Else
		GUICtrlSetBkColor($Input[$noping][1] + 2, 0xffff00) ; yellow
		GUICtrlSetBkColor($Input[$noping][2] + 2, 0xffff00) ; yellow
	EndIf
	;
	$pPing = Ping($machine[$Nb_of_ping][1], 10)
	If $pPing = 0 And StringLeft($machine[$Nb_of_ping][1], $Len) <> $IP Then
		$pPing = Ping($machine[$Nb_of_ping][1], $TOut * 2)
	EndIf
	;
	If $Nb_of_ping >= $Nb_of_machines Then
		If (Not ($ShowAllPing) And $noping <> $oldnoping) = 1 Then
			$oldnoping = $noping
			$RedoGui = 1
		EndIf
		If $noping = 0 Then
			If @Compiled Then
				TraySetIcon(@ScriptFullPath, -5)
			Else
				TraySetIcon(@ScriptDir & "\pingtest1.0.6.3.exe", -5)
			EndIf
			$messtray = $head
			$f = 1
		Else
			$f = 1
		EndIf
		;
		TraySetToolTip($messtray)
		$Nb_of_ping = 0
		$Ttimer = TimerInit()
		Return
	EndIf
	$f = 0
EndFunc   ;==>_Ping

Func _CreateGui($bFirstTime = 0)
	If Not ($bFirstTime) Then
		$state = WinGetState($head)
		If Not BitAND($state, 2) Then ; 2 = visible
			$winstate = 0
		ElseIf Not BitAND($state, 8) Then ; 8 = active
			$winstate = 1
		Else
			$winstate = 2
			$Winposping = WinGetPos($head)
		EndIf
		GUIDelete($gui)

		If $Nb_of_ping > 1 Then $Nb_of_ping = 0
		$Tdiff = $TWait

	EndIf

	If $ShowAllPing Or $bFirstTime Then
		$iCount = $Nb_of_machines - 1
		$aCount = $machine
		$iPos1 = @DesktopWidth - 320
		$iPos2 = -1
	Else
		$oldnoping = $noping
		$iCount = $noping
		$aCount = $machine_not_good
		$iPos1 = $Winposping[0]
		$iPos2 = $Winposping[1]
	EndIf

	$gui = GUICreate($head, 300, (15 * $iCount) + 22, $iPos1, $iPos2)
	$la = GUICtrlCreateLabel("", 0, 0, 300, (15 * $iCount) + 1)
	GUICtrlSetState($la, $GUI_DISABLE)
	GUICtrlSetBkColor($la, 0x000000)
	If $ShowAllPing Or $bFirstTime Then
		$button_hide = GUICtrlCreateButton("<< Hide", 180, (15 * $iCount) + 1, 100, 20, $BS_FLAT)
	Else
		$button_show = GUICtrlCreateButton("Show >>", 180, (15 * $iCount) + 1, 100, 20, $BS_FLAT)
	EndIf
	$button_refresh = GUICtrlCreateButton("Refresh", 20, (15 * $iCount) + 1, 100, 20, $BS_FLAT)
	For $i = 1 To $iCount
		$Input[$i][1] = GUICtrlCreateInput('  ' & $aCount[$i][1], 1, (15 * ($i - 1)) + 1, 98, 14, $ES_READONLY, $WS_EX_NOINHERITLAYOUT)
		$Input[$i][2] = GUICtrlCreateInput('  ' & $aCount[$i][2], 100, (15 * ($i - 1)) + 1, 199, 14, $ES_READONLY, $WS_EX_NOINHERITLAYOUT)
		GUICtrlSetBkColor($Input[$i][1], 0xffff00)     ; yellow
		GUICtrlSetBkColor($Input[$i][2], 0xffff00)     ; yellow
	Next
	If $bFirstTime Then Return

	Select
		Case $winstate = 2 ; 2 -> active
			GUISetState(@SW_SHOWNORMAL, $gui)
		Case $winstate = 1 ; 1 -> not active
			GUISetState(@SW_SHOWNOACTIVATE, $gui)
		Case $winstate = 0 ; 0 -> on the tray
			GUISetState(@SW_HIDE, $gui)
	EndSelect
	$RedoGui = 0
EndFunc   ;==>_CreateGui

Func GUISettings()
	$F_Settings = GUICreate("Settings", 274, 222, -1, -1, $WS_CAPTION)
	GUICtrlCreateGroup(" Ping Time ", 5, 5, 261, 46)
	$Wait = GUICtrlCreateInput(($TWait / 1000), 129, 20, 50, 20)
	GUICtrlCreateLabel("Ping every: [sec]", 10, 20, 117, 20, $SS_CENTERIMAGE)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	GUICtrlCreateGroup(" Inside Network ", 5, 55, 261, 81)
	$InWait = GUICtrlCreateInput($TOut_In, 95, 105, 50, 20)
	GUICtrlCreateLabel("Time out: [msec]", 10, 105, 82, 20, $SS_CENTERIMAGE)
	GUICtrlCreateLabel("Set the Inside Network:", 10, 75, 115, 20, $SS_CENTERIMAGE)
	$IPAddress1 = GUICtrlCreateInput($IP, 125, 75, 130, 20)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	GUICtrlCreateGroup(" Outside Network ", 5, 140, 261, 46)
	$OutWait = GUICtrlCreateInput(($TOut_Out / 1000), 165, 155, 50, 20)
	GUICtrlCreateLabel("Time out ouside Network: [sec]", 15, 155, 148, 20, $SS_CENTERIMAGE)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$Apply = GUICtrlCreateButton("Apply and Close", 87, 192, 100, 25, BitOR($BS_CENTER, $BS_DEFPUSHBUTTON, $BS_VCENTER))
	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Select
			Case $nMsg = $Apply
				$state = WinGetState($head, "")
				$TWait = (GUICtrlRead($Wait) * 1000)
				If $TWait < 10000 Then $TWait = 10000
				$TOut_In = GUICtrlRead($InWait)
				$TOut_Out = (GUICtrlRead($OutWait) * 1000)
				$Len = StringLen(GUICtrlRead($IPAddress1))
				$IP = GUICtrlRead($IPAddress1) ;192.168.
				GUIDelete($F_Settings)
				IniWrite(@ScriptDir & "\pingtest.ini", "Settings", "PingTime", ($TWait / 1000))
				IniWrite(@ScriptDir & "\pingtest.ini", "Settings", "InsideNetworkIP", $IP)
				IniWrite(@ScriptDir & "\pingtest.ini", "Settings", "INW_TimeOut", $TOut_In)
				IniWrite(@ScriptDir & "\pingtest.ini", "Settings", "ONW_TimeOut", ($TOut_Out / 1000))
				$RedoGui = 1
				Return
		EndSelect
	WEnd
EndFunc   ;==>GUISettings
