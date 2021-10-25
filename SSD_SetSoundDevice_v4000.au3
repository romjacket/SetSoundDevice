#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ssd_logo.ico
#AutoIt3Wrapper_Outfile=SSD.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Comment=© SSD - SetSoundDevice, 2010-2017 by Karsten Funk. All rights reserved. http://www.funk.eu
#AutoIt3Wrapper_Res_Description=SetSoundDevice
#AutoIt3Wrapper_Res_Fileversion=4.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=© Karsten Funk under Creative Commons "by-nc-sa 3.0"
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Field=Made By|Karsten Funk
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile Date|%date% %time%
#AutoIt3Wrapper_Res_Field=ProductName|SMF - Search my Files
#AutoIt3Wrapper_Res_Field=ProductVersion|%AutoItVer%
#AutoIt3Wrapper_Res_Field=CompanyName|Karsten Funk. All rights reserved. http://www.funk.eu
#AutoIt3Wrapper_Res_Field=LegalTrademarks|by-nc-nd 3.0
#AutoIt3Wrapper_Res_Field=InternalName|%scriptfile%
#AutoIt3Wrapper_Res_Field=Platform|Vista,Win7,Win8,Win81,Win10
#AutoIt3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_Tidy_Stop_OnError=n
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/pe /sf /sv /rm /rsln
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Global Const $s_versionnumber = "v4 - (2017-Sep-16)"
Global Const $sGUITitle = "SSD - SetSoundDevice - " & $s_versionnumber
Global Const $iBuildnumber = 4000

Opt("WinDetectHiddenText", 1)
Opt("WinWaitDelay", 100) ; default = 250
Opt("SendKeyDelay", 10) ; default = 5
Opt("SendKeyDownDelay", 10) ; default = 5

; _SetSoundDeviceSet_Speaker_Config.au3
; take a look at speaker config fucntion
; Add option for permanent tray icon / desktop contextmenu integration?

; #  PROGRAM HEADER  # ==========================================================================================================
;
; Name...........: SSD - SetSoundDevice on Vista+
;
; Description ...: Windows XP allows you to output sound to more than one audio device. Win Vista and Windows 7 do not.
;                  To overcome this “Failure by Design”, (IMHO somehow related to DRM, preventing user to make digital
;                  copies of analog sources), I’ve coded SSD – Set Sound Device.
;
; Author ........: Karsten Funk (KaFu)
;
; Buildnumber....: 4000, 2017 Sep 16
; AutoIt Version.: 3.3.14.2
;
; Website........: http://funk.eu/ssd/
;
; Comments.......: For the "invisible" mode SSD opens the Control Panel on a separate Desktop (see "CreateDesktop" call).
;
; ===============================================================================================================================

; #  CHANGELOG  # ===============================================================================================================
; v3 > v4
; 4.0.0.0: - 170915  -   Updated - If more then one "Sound" window seems to be open, try to close wrong ones
; 4.0.0.0: - 170915  -   Updated - to use AutoIt Version 3.3.14.2


; v2.1 > v3
; 3.0.0.1: - Updated -   If more then one "Sound" window seems to be open, try to close wrong ones
; 3.0.0.1: - Updated -   Run program if at least one hwnd_sound is detected, see "If $a_EnumChildWindows[0][0] > 0 Then"

; v2 > v2.1
; 2.0.0.1: - Updated -   Removed error msgbox on first startup, when no recording devices are detected
; 2.0.0.1: - Updated -   Added _MsgBox_SHEx function
; 2.0.0.2: - Updated -   Changed mode for SSD.ini file initialisation from 1024 to 32 > Use Unicode UTF16 Little Endian reading and writing mode
; 2.0.0.3: - Updated -   Changed CreateDesktop flags to $DESKTOP_ALL_ACCESS
; 2.0.0.6: - Updated -   Tested for Win10 compatibility
; 2.0.0.7: - Updated -   Removed @AppDataDir save file location again, .ini file to be saved along SSD.exe only

; v1.3 > v2
; 1.3.0.8: - Updated -   Added function _Merge_Two_Shortcuts()
; 1.3.0.9: - Updated -   Added _WinAPI_SetWindowLong($hWnd_Sound, $GWL_HWNDPARENT, $h_GUI)
; 1.3.0.9: - Updated -   Added toggle mode functionality, see > If StringInStr($iDeviceNo, 999, 0) Then ; toggle mode
; 1.3.2.3: - Updated -   Added option to change recording device too
; 1.3.2.3: - Updated -   SSD.ini required to make toggle work
; 1.3.2.3: - Updated -   SSD Hopefully Win10 compatible now
; 1.3.2.4: - Updated -   Added option to perform setting based on "Device Name"
; 1.3.3.0: - Updated -   Optionally save SSD.ini into @AppDataDir
; 1.3.3.1: - Updated -   Added "Select shortcut location"
; 1.3.3.6: - Updated -   Added "Save and re-apply Volume"
; 1.3.3.9: - Updated -   Added _First_Startup()
; 1.3.4.0: - Updated -   Ask User before deleting existing SSD.ini files / Ask User before creating new SSD.ini file at ANY location
; 1.3.4.3: - Updated -   Add optional tooltip @ MousePos on succesful device switch
; 1.3.4.6: - Updated -   Enhanced tooltip and tray messages on switch
; 1.3.4.8: - Updated -   Minor layout changes

#Region DllOpen_PostProcessor START
Global $h_DLL_Kernel32 = DllOpen("kernel32.dll")
Global $h_DLL_User32 = DllOpen("user32.dll")
Global $h_DLL_advapi32 = DllOpen("advapi32.dll")
Global $h_DLL_gdi32 = DllOpen("gdi32.dll")
Global $h_DLL_shell32 = DllOpen("shell32.dll")
Global $h_DLL_ntdll = DllOpen("ntdll.dll")
Global $h_DLL_ole32 = DllOpen("ole32.dll")
Global $h_DLL_oleaut32 = DllOpen("oleaut32.dll")
Global $h_DLL_Crypt32 = DllOpen("Crypt32.dll")
OnAutoItExitRegister("_OnAutoitExit_CloseDlls")
#EndRegion DllOpen_PostProcessor START

Global Const $g_Macro_CR = @CR, $g_Macro_CRLF = @CRLF, $g_Macro_LF = @LF, $g_Macro_TAB = @TAB
Global Const $g_Macro_Compiled = @Compiled, $g_Macro_SW_DISABLE = @SW_DISABLE, $g_Macro_SW_ENABLE = @SW_ENABLE, $g_Macro_SW_HIDE = @SW_HIDE, $g_Macro_SW_LOCK = @SW_LOCK, $g_Macro_SW_MAXIMIZE = @SW_MAXIMIZE, $g_Macro_SW_MINIMIZE = @SW_MINIMIZE, $g_Macro_SW_RESTORE = @SW_RESTORE, $g_Macro_SW_SHOW = @SW_SHOW, $g_Macro_SW_SHOWDEFAULT = @SW_SHOWDEFAULT, $g_Macro_SW_SHOWMAXIMIZED = @SW_SHOWMAXIMIZED, $g_Macro_SW_SHOWMINIMIZED = @SW_SHOWMINIMIZED, $g_Macro_SW_SHOWMINNOACTIVE = @SW_SHOWMINNOACTIVE, $g_Macro_SW_SHOWNA = @SW_SHOWNA, $g_Macro_SW_SHOWNOACTIVATE = @SW_SHOWNOACTIVATE, $g_Macro_SW_SHOWNORMAL = @SW_SHOWNORMAL, $g_Macro_SW_UNLOCK = @SW_UNLOCK
Global Const $g_Macro_AutoItExe = @AutoItExe, $g_Macro_AutoItPID = @AutoItPID, $g_Macro_AutoItVersion = @AutoItVersion, $g_Macro_AutoItX64 = @AutoItX64, $g_Macro_CPUArch = @CPUArch, $g_Macro_OSArch = @OSArch, $g_Macro_OSBuild = @OSBuild, $g_Macro_OSVersion = @OSVersion, $g_Macro_OSType = @OSType, $g_Macro_ComputerName = @ComputerName, $g_Macro_SystemDir = @SystemDir, $g_Macro_AppDataDir = @AppDataDir, $g_Macro_WindowsDir = @WindowsDir, $g_Macro_TempDir = @TempDir, $g_Macro_ProgramFilesDir = @ProgramFilesDir, $g_Macro_ComSpec = @ComSpec, $g_Macro_ScriptDir = @ScriptDir, $g_Macro_ScriptFullPath = @ScriptFullPath, $g_Macro_ScriptName = @ScriptName
ConsoleWrite("-Obfuscation Prevention" & $g_Macro_AutoItExe & $g_Macro_AutoItPID & $g_Macro_AutoItVersion & $g_Macro_AutoItX64 & $g_Macro_Compiled & $g_Macro_CPUArch & $g_Macro_CR & $g_Macro_CRLF & $g_Macro_LF & $g_Macro_OSArch & $g_Macro_OSBuild & $g_Macro_ScriptDir & $g_Macro_ScriptFullPath & $g_Macro_ScriptName & $g_Macro_SW_DISABLE & $g_Macro_SW_ENABLE & $g_Macro_SW_HIDE & $g_Macro_SW_LOCK & $g_Macro_SW_MAXIMIZE & $g_Macro_SW_MINIMIZE & $g_Macro_SW_RESTORE & $g_Macro_SW_SHOW & $g_Macro_SW_SHOWDEFAULT & $g_Macro_SW_SHOWMAXIMIZED & $g_Macro_SW_SHOWMINIMIZED & $g_Macro_SW_SHOWMINNOACTIVE & $g_Macro_SW_SHOWNA & $g_Macro_SW_SHOWNOACTIVATE & $g_Macro_SW_SHOWNORMAL & $g_Macro_SW_UNLOCK & $g_Macro_TAB & @CRLF & $g_Macro_OSVersion & $g_Macro_OSType & $g_Macro_ComputerName & $g_Macro_SystemDir & $g_Macro_AppDataDir & $g_Macro_WindowsDir & $g_Macro_TempDir & $g_Macro_ProgramFilesDir & $g_Macro_ComSpec & @CRLF & "-------" & @CRLF & @CRLF)

Global $__CmdLineRaw = $CmdLineRaw
$__CmdLineRaw = StringReplace($__CmdLineRaw, "/ErrorStdOut ", "")
$__CmdLineRaw = StringReplace($__CmdLineRaw, """" & @ScriptFullPath & """", "")
$__CmdLineRaw = StringStripWS($__CmdLineRaw, 3)

#Region ;----- Includes -----
#include <StaticConstants.au3>
#include <GDIPlus.au3>
; Global Const $STM_SETIMAGE = 0x0172

#include <GuiListView.au3>
#include <ComboConstants.au3>
#include <GuiComboBox.au3>
#include <GUIConstantsEx.au3>
#include <Misc.au3>
#include <Winapi.au3>
#include <WinAPIEx.au3>
#include <WindowsConstants.au3>
#include <SliderConstants.au3>

; #include <WinAPIEx_3.8_3380\APIConstants.au3>
; #include <WinAPIEx_3.8_3380\WinAPIEx.au3>

#include <SSD_z_Func_EnumChildWindows.au3>
#include <SSD_z_Func_AudioEndpointVolume_v04.au3>
#EndRegion ;----- Includes -----

Global $__a_Global_MsgBox_SHEx[11]
; $__a_Global_MsgBox_SHEx[0] = SHELLHOOK registered
; $__a_Global_MsgBox_SHEx[1] = hWnd for Hook
; $__a_Global_MsgBox_SHEx[2] = MsgBox Icon
; $__a_Global_MsgBox_SHEx[3] = MsgBox Title
; $__a_Global_MsgBox_SHEx[4] = MsgBox Text
; $__a_Global_MsgBox_SHEx[5] = Button 1 - Text
; $__a_Global_MsgBox_SHEx[6] = Button 2 - Text
; $__a_Global_MsgBox_SHEx[7] = Button 3 - Text
; $__a_Global_MsgBox_SHEx[8] = Apply $WS_EX_TOOLWINDOW to MsgBox
; $__a_Global_MsgBox_SHEx[9] = Transparency > A number in the range 0 - 255. The lower the number, the more transparent the window will become. 255 = Solid, 0 = Invisible.
; $__a_Global_MsgBox_SHEx[10] = hWnd for "Parent" GUI

Global $__WINVER_RtlGetVersion = __WINVER_RtlGetVersion()

#include <Crypt.au3>
Global $s_Computer_ID = StringLower(@ComputerName & "_" & Hex($__WINVER_RtlGetVersion, 4) & "_" & @OSArch & @CPUArch)
If @OSArch = "x64" Then
	$s_Computer_ID &= RegRead("HKLM64\SOFTWARE\Microsoft\Cryptography", "MachineGuid")
Else
	$s_Computer_ID &= RegRead("HKLM\SOFTWARE\Microsoft\Cryptography", "MachineGuid")
EndIf
$s_Computer_ID = StringTrimLeft(_Crypt_HashData($s_Computer_ID, $CALG_MD5), 2)

Global $hWnd_Sound, $iPID_Sound, $timer_timeout, $iDeviceNo = ""

Global $sGUID = "12eef221-387e-4f5a-a9d2-fd68799f1438"

Global $h_Taskbar = WinGetHandle("[CLASS:Shell_TrayWnd;]")
Global $b_SSD_Is_Running_on_Hidden_Desktop = False, $i_Taskbar_PID

If _WinAPI_GetThreadDesktop(_WinAPI_GetCurrentThreadId()) <> _WinAPI_GetThreadDesktop(_WinAPI_GetWindowThreadProcessId($h_Taskbar, $i_Taskbar_PID)) Then
	$b_SSD_Is_Running_on_Hidden_Desktop = True
	$sGUID = "12eef221-387e-4f5a-a9d2-Hidden"
EndIf

Global $a_Global_MousePos_Buffer

; Debug
#cs
	$b_SSD_Is_Running_on_Hidden_Desktop = True
	$__CmdLineRaw = "7776"
	$g_Macro_Compiled = 1
#ce

If $__WINVER_RtlGetVersion < 0x0600 Then ; Windows Vista or later
	If $b_SSD_Is_Running_on_Hidden_Desktop = False Then _MsgBox_SHEx(16, $sGUITitle & " - Info", "SSD works on Win7+ only (maybe Vista, but untested).", 10)
	Exit 11
EndIf

Global $b_Change_Recording_Device = False
If StringInStr($__CmdLineRaw, "_Recording", 2) Then
	$__CmdLineRaw = StringReplace($__CmdLineRaw, "_Recording", "")
	$b_Change_Recording_Device = True
	$eRender = 1 ; Switch to recording device
	$sGUID &= "_Recording"

	#cs
		If $b_Change_Recording_Device = True Then $eRender = 1 ; Switch to recording device
		> $eConsole = 2 for Communication Device switch only?
		0 = eConsole = Games, system notification sounds, and voice commands.
		1 = eMultimedia = Music, movies, narration, and live music recording.
		2 = eCommunications = Voice communications (talking to another person).
	#ce
EndIf

; For Debugging Recording Mode
; $b_Change_Recording_Device = True

Global $h_Global_Tooltip
Global $hwnd_WM_COPYDATA_Receiver = GUICreate("SSD_WM_COPYDATA_" & $b_SSD_Is_Running_on_Hidden_Desktop)
Global $hwnd_AutoIt = WinGetHandle(AutoItWinGetTitle())
WinSetOnTop($hwnd_AutoIt, "", 1)
ControlSetText($hwnd_AutoIt, '', ControlGetHandle($hwnd_AutoIt, '', 'Edit1'), $hwnd_WM_COPYDATA_Receiver) ; to pass hWnd of main GUI to AutoIt default GUI

_EnforceSingleInstance($sGUID) ; any 'unique' string; created with http://www.guidgen.com/Index.aspx

GUIRegisterMsg($WM_COPYDATA, "WM_COPYDATA")

If $b_SSD_Is_Running_on_Hidden_Desktop = False And $b_Change_Recording_Device = False Then
	If FileExists(@ScriptDir & "\SSD.ini") Then
		If IniRead(@ScriptDir & "\SSD.ini", "Settings", "SSD_Version", "") <> String($sGUITitle & "-" & $iBuildnumber) Then
			If _MsgBox_SHEx(4 + 64, $sGUITitle & ' - SSD.ini initialization', "SSD found an existing SSD.ini file here:" & @CRLF _
					 & @ScriptDir & "\SSD.ini" & @CRLF & @CRLF _
					 & "The build number indicates a different SSD version." & @CRLF & @CRLF _
					 & "SSD.ini Build Number = " & @CRLF & """" & IniRead(@ScriptDir & "\SSD.ini", "Settings", "SSD_Version", "") & """" & @CRLF & @CRLF _
					 & "This version Build Number = " & @CRLF & """" & String($sGUITitle & "-" & $iBuildnumber) & """" & @CRLF & @CRLF _
					 & "For SSD to work properly, this SSD.ini file has to be deleted." & @CRLF & @CRLF & "Go ahead?", 0, $hwnd_AutoIt) <> 6 Then
				Exit 12
			EndIf
			FileDelete(@ScriptDir & "\SSD.ini")
		EndIf
	EndIf

EndIf

Global $s_INI_File_Location = @ScriptDir & "\SSD.ini"

Local $b_First_Startup_Success = False

If $b_SSD_Is_Running_on_Hidden_Desktop = False And $b_Change_Recording_Device = False Then

	If Not FileExists(@ScriptDir & "\SSD.ini") Then

		Local $h_file = -1

		; 3 Yes, No, and Cancel
		Switch _MsgBox_SHEx(1 + 64, $sGUITitle & ' - SSD.ini initialization', "SSD has to create a SSD.ini file to work properly." & @CRLF & @CRLF _
				 & "It will be created at:" & @ScriptDir & "\SSD.ini" & @CRLF & @CRLF _
				 & "Please make sure that the directory is writeable, as SSD provides portable installation only.", 0, $hwnd_AutoIt)

			Case 1 ; ok
				$h_file = FileOpen(@ScriptDir & "\SSD.ini", 2 + 32) ; 32 = Use Unicode UTF16 Little Endian reading and writing mode.

			Case Else
				Exit 14

		EndSwitch

		If $h_file = -1 Then

			$s_Error_Message = "SSD works only, if the SSD.ini file is writeable." & @CRLF & @CRLF & "The SSD.ini file is created in the same directory as SSD.exe is in. Move SSD.exe to a writeable location." & @CRLF & @CRLF & @ScriptDir & "\SSD.ini"
			_MsgBox_SHEx(16, $sGUITitle & ' - Error', $s_Error_Message, 0, $hwnd_AutoIt)
			Exit 15

		Else
			$b_First_Startup_Success = True
		EndIf

		FileWrite($h_file, "[Settings]" & @CRLF) ; you need to write something
		FileClose($h_file)

	EndIf

	If Not IniWrite(@ScriptDir & "\SSD.ini", "Settings", "SSD_Version", $sGUITitle & "-" & $iBuildnumber) Then

		$s_Error_Message = "SSD works only, if the SSD.ini file is writeable." & @CRLF & @CRLF & "The SSD.ini file is created in the same directory as SSD.exe resides in. Move SSD.exe to a writeable location." & @CRLF & @CRLF & @ScriptDir & "\SSD.ini"
		_MsgBox_SHEx(16, $sGUITitle & ' - Error', $s_Error_Message, 0, $hwnd_AutoIt)
		Exit 17

		If $b_Change_Recording_Device = True Then
			IniDelete($s_INI_File_Location, "Settings", "Last_Activated_Recording_Device")
		Else
			IniDelete($s_INI_File_Location, "Settings", "Last_Activated_Playback_Device")
		EndIf
	EndIf

EndIf

If $b_First_Startup_Success = True Then
	; _MsgBox_SHEx(64, $sGUITitle & ' - Info', "SSD has created a new SSD.ini file here:" & @CRLF & @CRLF & $s_INI_File_Location & @CRLF & @CRLF & "It might be necessary to re-create any existing shortcuts created with previous versions of SSD for the program to work properly.", 0, $hwnd_AutoIt)
	_SSD_First_Startup()
	IniReadSection($s_INI_File_Location, "Names_of_Recording_Devices_" & $s_Computer_ID)
	If @error Then ConsoleWrite("! First Startup - Error - 'Names_of_Recording_Devices_' could not be aquired." & @CRLF)
	; If @error Then _MsgBox_SHEx(16, $sGUITitle & ' - Error', "First Startup" & @CRLF & @CRLF & """Names of Recording Devices"" could not be aquired.", 0, $hwnd_AutoIt)
EndIf

Func _SSD_First_Startup()
	ConsoleWrite("+ First Startup" & @CRLF)
	AutoItWinSetTitle("")
	Local $iPID = Run(@ScriptFullPath & " _Recording_First_Startup_hidden", @ScriptDir, @SW_HIDE)
	Local $iTimer = TimerInit()

	While ProcessExists($iPID)
		Sleep(250)
		; check for error msgbox here

		If WinExists("[CLASS:#32770;REGEXPTITLE:(?i)(.* First Startup - No recording devices detected.*)]") Then
			If TimerDiff($iTimer) > 60000 Then ExitLoop
		Else
			If TimerDiff($iTimer) > 10000 Then ExitLoop
		EndIf

	WEnd

	; _MsgBox_SHEx(0, "First Startup", $iPID & @CRLF & _Get_ExitCode($iPID) & @crlf & TimerDiff($iTimer))
	ProcessClose($iPID)
	AutoItWinSetTitle($hwnd_AutoIt)
	$__CmdLineRaw = ""
EndFunc   ;==>_SSD_First_Startup

Global $b_NoChange_Com = False
If StringInStr($__CmdLineRaw, "_NoChange_Com", 2) Then
	$__CmdLineRaw = StringReplace($__CmdLineRaw, "_NoChange_Com", "")
	$b_NoChange_Com = True
EndIf

Global $b_NoChange_Audio = False
If StringInStr($__CmdLineRaw, "_NoChange_Audio", 2) Then
	$__CmdLineRaw = StringReplace($__CmdLineRaw, "_NoChange_Audio", "")
	$b_NoChange_Audio = True

	; Set Sound Level for this only:
	$eConsole = 2 ; 2 = eCommunications = Voice communications (talking to another person).

EndIf

Global $s_Global_Traytip_Message, $s_Global_Traytip_Title, $i_Global_Traytip_Icon, $i_Global_Traytip_Options

If $g_Macro_Compiled And $__CmdLineRaw And $b_SSD_Is_Running_on_Hidden_Desktop = False And StringInStr($__CmdLineRaw, "hidden") Then

	$iDeviceNo = StringReplace($__CmdLineRaw, "hidden", "")

	If $b_NoChange_Com = True Then $iDeviceNo &= "_NoChange_Com"
	If $b_NoChange_Audio = True Then $iDeviceNo &= "_NoChange_Audio"

	; $h_Desktop_New = _WinAPI_CreateDesktop('SSD_Desktop', BitOR($DESKTOP_CREATEWINDOW, $DESKTOP_SWITCHDESKTOP),0x0001) ; $GENERIC_ALL

	Global $h_Desktop_New = _WinAPI_CreateDesktop('SSD_Desktop', $DESKTOP_ALL_ACCESS, $DF_ALLOWOTHERACCOUNTHOOK)
	If Not $h_Desktop_New Then
		_MsgBox_SHEx(16, $sGUITitle & ' - Error', 'Unable to create desktop, SSD will not work hidden.', 0, $hwnd_AutoIt)
		Exit 18
	EndIf

	; _WinAPI_SwitchDesktop($h_Desktop_New)

	Global $tProcess = DllStructCreate($tagPROCESS_INFORMATION)
	Global $tStartup = DllStructCreate($tagSTARTUPINFO)
	DllStructSetData($tStartup, 'Size', DllStructGetSize($tStartup))

	Global $tDesktop
	DllStructSetData($tStartup, 'Desktop', _WinAPI_CreateString('SSD_Desktop', $tDesktop))

	Local $s_Add_Recording_Switch
	If $b_Change_Recording_Device = True Then $s_Add_Recording_Switch = "_Recording"

	If StringInStr($__CmdLineRaw, "_First_Startup", 2) Then $s_Add_Recording_Switch &= "_First_Startup"

	If _WinAPI_CreateProcess('', '"' & @ScriptFullPath & '" ' & $iDeviceNo & $s_Add_Recording_Switch, 0, 0, 0, $CREATE_NEW_PROCESS_GROUP, 0, 0, DllStructGetPtr($tStartup), DllStructGetPtr($tProcess)) Then

		$i_RunTimer = TimerInit()
		While Sleep(10)

			If _Get_ExitCode(DllStructGetData($tProcess, 'ProcessID')) <> 259 Then ExitLoop ; STILL_ACTIVE (259)

			If TimerDiff($i_RunTimer) > 10000 Then
				ProcessClose(DllStructGetData($tProcess, 'ProcessID'))
				ExitLoop
			EndIf

		WEnd

		ProcessClose(DllStructGetData($tProcess, 'ProcessID'))

		Local $b_Skip_Exit_Code_of_Hidden_Process = False
		Local $s_Error_Message
		Local $i_Get_ExitCode = _Get_ExitCode(DllStructGetData($tProcess, 'ProcessID'))
		Switch $i_Get_ExitCode

			Case 6741 ; toggle mode ssd.ini file not writeable
				$s_Error_Message = "SSD Toggle Mode works only, if the SSD.ini file is writeable." & @CRLF & @CRLF & "The SSD.ini file is created in the same directory as SSD.exe resides in, move SSD.exe to a writeable location." & @CRLF & @CRLF & $s_INI_File_Location

			Case 6821 ; _Exit(682)
				$s_Error_Message = "SSD was started hidden with no device provided to switch too."

			Case 581 ; > _Exit(58)
				; $s_Error_Message = "First Startup of Recording successful"
				$b_Skip_Exit_Code_of_Hidden_Process = True

			Case 3331 ; > _Exit(333)
				$s_Error_Message = "SSD found volume change trigger in shortcut, but could not extract information."

			Case 9311, 9321 ; _Exit(931) & _Exit(932)
				$s_Error_Message = "SSD found volume change trigger in shortcut, but could not extract information."

			Case 11 ; > _Exit(1)
				$s_Error_Message = "Sound control panel window not found in time."

			Case 231 ; > _Exit(23)
				$s_Error_Message = "Running hidden without device to switch found in command line argument > error"

			Case 21 ; > _Exit(2)
				If $b_Change_Recording_Device = True Then
					$s_Error_Message = "No rows found in SysListView32 control of the 'Sound' control panel dialog, no recording devices installed?"
				Else
					$s_Error_Message = "No rows found in SysListView32 control of the 'Sound' control panel dialog, no output devices installed?"
				EndIf

			Case 51 ; > _Exit(5)
				$s_Error_Message = "SysListView32 control in 'Sound' control panel dialog not found."

			Case 771 ; > _Exit(77)
				$s_Error_Message = "SSD could not terminate existing rundll32.exe process in time."

			Case 7711 ; > _Exit(771)
				$s_Error_Message = "SSD shortcut indicates that device name is to be used, but device name #1 was not found in SSD.ini." & @CRLF & $iDeviceNo

			Case 7721 ; > _Exit(772)
				$s_Error_Message = "SSD shortcut indicates that device name is to be used, but device name #2 was not found in SSD.ini." & @CRLF & $iDeviceNo

			Case 7731 ; > _Exit(773)
				$s_Error_Message = "SSD shortcut indicates that device name is to be used, but device name was not found in SSD.ini." & @CRLF & $iDeviceNo

			Case 41
				; no error, hidden process ran through smoothly

			Case 0 ; ?
				$s_Error_Message = "" ; "Error 0"

			Case Else
				$s_Error_Message = "Hidden Process ended with error code"
				; e.g. _Exit((8 & $iDeviceNo & WinExists($hWnd_Sound)))

		EndSwitch

		If $s_Error_Message Then

			Run("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0")

			If $i_Get_ExitCode = 21 And $b_Change_Recording_Device And StringInStr($__CmdLineRaw, "_First_Startup", 2) Then

				_MsgBox_SHEx(32, $sGUITitle & " - First Startup - No recording devices detected", $s_Error_Message, 0, $hwnd_AutoIt)

			Else

				_MsgBox_SHEx(16, $sGUITitle & " - Error", $s_Error_Message & @CRLF & @CRLF _
						 & "Error code: " & $i_Get_ExitCode & @CRLF & @CRLF _
						 & "- Does the ""Sound"" dialog pop-up when started from control panel?" & @CRLF & "- Did you install a new version of Windows or SSD?" & @CRLF & @CRLF & "Try to re-create the shortcuts." & @CRLF & @CRLF _
						 & $__CmdLineRaw & @CRLF & @CRLF _
						 & $CmdLineRaw, 0, $hwnd_AutoIt)

			EndIf


		ElseIf $b_Skip_Exit_Code_of_Hidden_Process = True Then
			; Do nothing

		Else

			_Check_for_ToolTip_and_TrayTip_Notifications()

		EndIf

	EndIf

	; _WinAPI_SwitchDesktop($h_Desktop_Current)

	_WinAPI_CloseDesktop($h_Desktop_New) ; do not close desktop to improve performance of succeding calls?

	Exit 19

ElseIf $g_Macro_Compiled And $__CmdLineRaw Then

	$iDeviceNo = $__CmdLineRaw

EndIf

Func _Check_for_ToolTip_and_TrayTip_Notifications()

	If $b_Change_Recording_Device = True Then
		$s_Global_Traytip_Message = IniRead($s_INI_File_Location, "Settings", "Last_Activated_Recording_Device", "")
	Else
		$s_Global_Traytip_Message = IniRead($s_INI_File_Location, "Settings", "Last_Activated_Playback_Device", "")
	EndIf

	Local $b_Activate_Sleep = False

	If IniRead($s_INI_File_Location, "GUI", "c_checkbox_show_tray_notification_on_change", $GUI_UNCHECKED) = $GUI_CHECKED Then
		$b_Activate_Sleep = True
		TraySetState()
		If $s_Global_Traytip_Message Then ; all seems good
			If $b_Change_Recording_Device = False Then
				$s_Global_Traytip_Title = "SSD changed PLAYBACK device to"
				TrayTip($s_Global_Traytip_Title, $s_Global_Traytip_Message, 10, 16)
			Else
				$s_Global_Traytip_Title = "SSD changed RECORDING device to"
				TrayTip($s_Global_Traytip_Title, $s_Global_Traytip_Message, 10, 16)
			EndIf
		Else ; error
			$s_Global_Traytip_Message = " " & @CRLF & $CmdLineRaw
			If $b_Change_Recording_Device = False Then
				$s_Global_Traytip_Title = "SSD failed to change PLAYBACK device"
				TrayTip($s_Global_Traytip_Title, $s_Global_Traytip_Message, 10, 19)
			Else
				$s_Global_Traytip_Title = "SSD failed to change RECORDING device"
				TrayTip($s_Global_Traytip_Title, $s_Global_Traytip_Message, 10, 19)
			EndIf
		EndIf
	EndIf

	If IniRead($s_INI_File_Location, "GUI", "c_checkbox_show_tooltip_at_mousepos_on_change", $GUI_CHECKED) = $GUI_CHECKED Then

		ToolTip("1992f793-4307-4f9a-bab0-3bba40940241")
		$h_Global_Tooltip = WinGetHandle("[TITLE:1992f793-4307-4f9a-bab0-3bba40940241; CLASS:tooltips_class32]")

		$b_Activate_Sleep = True
		Local $a_MousePos = MouseGetPos()
		$a_MousePos[0] += 15
		$a_MousePos[1] += 15
		$a_Global_MousePos_Buffer = $a_MousePos
		If $s_Global_Traytip_Message Then ; all seems good
			$s_Global_Traytip_Message = StringReplace($s_Global_Traytip_Message, ", ", @CRLF)
			If $b_Change_Recording_Device = False Then
				$s_Global_Traytip_Title = "SSD changed PLAYBACK device to"
				$i_Global_Traytip_Icon = 1
				$i_Global_Traytip_Options = 0
				ToolTip($s_Global_Traytip_Message, $a_MousePos[0], $a_MousePos[1], $s_Global_Traytip_Title, $i_Global_Traytip_Icon, $i_Global_Traytip_Options)
			Else
				$s_Global_Traytip_Title = "SSD changed RECORDING device to"
				$i_Global_Traytip_Icon = 1
				$i_Global_Traytip_Options = 0
				ToolTip($s_Global_Traytip_Message, $a_MousePos[0], $a_MousePos[1], $s_Global_Traytip_Title, $i_Global_Traytip_Icon, $i_Global_Traytip_Options)
			EndIf
		Else ; error
			$s_Global_Traytip_Message = $CmdLineRaw
			If $b_Change_Recording_Device = False Then
				$s_Global_Traytip_Title = "SSD failed to change PLAYBACK device"
				$i_Global_Traytip_Icon = 4
				$i_Global_Traytip_Options = 0
				ToolTip($s_Global_Traytip_Message, $a_MousePos[0], $a_MousePos[1], $s_Global_Traytip_Title, $i_Global_Traytip_Icon, $i_Global_Traytip_Options)
			Else
				$s_Global_Traytip_Title = "SSD failed to change RECORDING device"
				$i_Global_Traytip_Icon = 4
				$i_Global_Traytip_Options = 0
				ToolTip($s_Global_Traytip_Message, $a_MousePos[0], $a_MousePos[1], $s_Global_Traytip_Title, $i_Global_Traytip_Icon, $i_Global_Traytip_Options)
			EndIf
		EndIf

		AdlibRegister("_ToolTip_Trace", 10)

	EndIf

	If $b_Activate_Sleep = True Then

		Local $i_Timer_Sleep = TimerInit()
		While Sleep(10)
			If TimerDiff($i_Timer_Sleep) > 2500 Then ExitLoop
		WEnd

		AdlibUnRegister("_ToolTip_Trace")
		TrayTip("SSD", "", 3, 1) ; clear tip
		ToolTip("") ; clear tip

	EndIf

EndFunc   ;==>_Check_for_ToolTip_and_TrayTip_Notifications

Func _ToolTip_Trace()
	Local $a_MousePos = MouseGetPos()
	$a_MousePos[0] += 15
	$a_MousePos[1] += 15
	If $a_MousePos[0] <> $a_Global_MousePos_Buffer[0] Or $a_MousePos[1] <> $a_Global_MousePos_Buffer[1] Then
		$a_Global_MousePos_Buffer = $a_MousePos
		WinMove($h_Global_Tooltip, "", $a_MousePos[0], $a_MousePos[1])
	EndIf
EndFunc   ;==>_ToolTip_Trace

Func _Get_ExitCode($hPID)
	Local $hProc = _WinAPI_OpenProcess(__Iif($__WINVER_RtlGetVersion < 0x0600, 0x00000500, 0x00001100), 0, $hPID)
	Local $aReturn = DllCall("kernel32.dll", "hwnd", "GetExitCodeProcess", "handle", $hProc, "dword*", 0)
	If @error Or Not IsArray($aReturn) Then
		_WinAPI_CloseHandle($hProc)
		_MsgBox_SHEx(16, "ERROR", "Could not get ExitCode", 0, $hwnd_AutoIt)
		Return 999
	EndIf
	_WinAPI_CloseHandle($hProc)
	Return $aReturn[2]
EndFunc   ;==>_Get_ExitCode

_Close_Existing_RunDll32_Processes()
If _Start_new_RunDll32_Process() = 1 Then
	_Close_Existing_RunDll32_Processes()
	_Start_new_RunDll32_Process(True) ; retried one more time, enough, exit with error now
EndIf

Func _Close_Existing_RunDll32_Processes()
	; Check for open "Sound" dialogs and close them gracefully
	Local $timer = TimerInit()
	While ProcessExists("rundll32.exe")
		$hWnd_Sound = _Detect_Sound_dialog_hWnd(ProcessExists("rundll32.exe"))
		If Not IsHWnd($hWnd_Sound) Then ExitLoop
		Local $timer_sub = TimerInit()
		While IsHWnd($hWnd_Sound)
			WinClose($hWnd_Sound)
			$hWnd_Sound = _Detect_Sound_dialog_hWnd(ProcessExists("rundll32.exe"))
			If TimerDiff($timer_sub) > 1000 Then
				; _Exit(7)
				ProcessClose(ProcessExists("rundll32.exe"))
				ExitLoop
			EndIf
			If Not IsHWnd($hWnd_Sound) Then ExitLoop
			Sleep(10)
		WEnd
		If TimerDiff($timer) > 10000 Then _Exit(77)
	WEnd
EndFunc   ;==>_Close_Existing_RunDll32_Processes

Func _Start_new_RunDll32_Process($B_Exit = False)
	If $b_Change_Recording_Device = False Then
		; Sound
		$iPID_Sound = Run("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0", @ScriptDir, @SW_HIDE)
	Else
		; Recording
		$iPID_Sound = Run("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,1", @ScriptDir, @SW_HIDE)
	EndIf
	$timer_timeout = TimerInit()
	While Not IsHWnd($hWnd_Sound)
		$hWnd_Sound = _Detect_Sound_dialog_hWnd($iPID_Sound)
		If TimerDiff($timer_timeout) > 10000 Or Not ProcessExists($iPID_Sound) Then
			If $B_Exit = True Then
				_Exit(1)
			Else
				Return 1 ; retry
			EndIf
		EndIf
		Sleep(10)
	WEnd
EndFunc   ;==>_Start_new_RunDll32_Process

If $b_SSD_Is_Running_on_Hidden_Desktop = True And ($b_NoChange_Audio = True Or $b_NoChange_Com = True) Then
	WinMove($hWnd_Sound, "", 10000, 10000) ; !! Clicking on a Split Button on a different Desktop still shows that button flash on visible Destkop !!!
Else
	WinMove($hWnd_Sound, "", 100, Default)
EndIf

If $b_SSD_Is_Running_on_Hidden_Desktop = True And StringLen($iDeviceNo) = 0 Then _Exit(23) ; running hidden without device to switch > error

Global $a_EnumChildWindows = _EnumChildWindows($hWnd_Sound, 0, 0, "SysListView32")
If $a_EnumChildWindows[0][0] > 1 Then
	; If more then one "Sound" window seems to be open, try to close wrong ones
	For $i = 1 To $a_EnumChildWindows[0][0]
		If _WinAPI_GetAncestor($a_EnumChildWindows[$i][0], $GA_ROOT) <> $hWnd_Sound Then WinClose(_WinAPI_GetAncestor($a_EnumChildWindows[$i][0], $GA_ROOT))
	Next
	Global $a_EnumChildWindows = _EnumChildWindows($hWnd_Sound, 0, 0, "SysListView32")
EndIf


Global $aSoundDevices

; WinClose(_WinAPI_GetAncestor($a_EnumChildWindows[1][0], $GA_ROOT))
#cs
	MsgBox(0, "", $a_EnumChildWindows[1][0] & @CRLF _
	& _WinAPI_GetParent($a_EnumChildWindows[1][0]) & @CRLF _
	& _WinAPI_GetAncestor($a_EnumChildWindows[1][0], $GA_PARENT) & @CRLF _
	& _WinAPI_GetAncestor($a_EnumChildWindows[1][0], $GA_ROOT) & @CRLF _
	& _WinAPI_GetAncestor($a_EnumChildWindows[1][0], $GA_ROOTOWNER))
#ce

Func _Device_Names_Get_Name_From_INI($i_Ref)
	Local $s_Ini_Device_Name, $a_Saved_Names_of_Devices
	If $b_Change_Recording_Device = True Then
		$a_Saved_Names_of_Devices = IniReadSection($s_INI_File_Location, "Names_of_Recording_Devices_" & $s_Computer_ID)
	Else
		$a_Saved_Names_of_Devices = IniReadSection($s_INI_File_Location, "Names_of_Playback_Devices_" & $s_Computer_ID)
	EndIf
	For $y = 1 To UBound($a_Saved_Names_of_Devices) - 1
		If $i_Ref = $a_Saved_Names_of_Devices[$y][0] Then
			$s_Ini_Device_Name = $a_Saved_Names_of_Devices[$y][1]
			ExitLoop
		EndIf
	Next
	If $s_Ini_Device_Name Then
		$s_Ini_Device_Name = StringTrimRight(StringTrimLeft($s_Ini_Device_Name, 1), 1)
		For $i = 0 To UBound($aSoundDevices) - 1
			If $s_Ini_Device_Name = $aSoundDevices[$i] Then Return $i + 1
		Next
	EndIf
	Return 0
EndFunc   ;==>_Device_Names_Get_Name_From_INI

Func _Device_Names_Get_Ref_From_INI($sName)
	Local $a_Saved_Names_of_Devices
	If $b_Change_Recording_Device = True Then
		$a_Saved_Names_of_Devices = IniReadSection($s_INI_File_Location, "Names_of_Recording_Devices_" & $s_Computer_ID)
	Else
		$a_Saved_Names_of_Devices = IniReadSection($s_INI_File_Location, "Names_of_Playback_Devices_" & $s_Computer_ID)
	EndIf
	For $y = 1 To UBound($a_Saved_Names_of_Devices) - 1
		$a_Saved_Names_of_Devices[$y][1] = StringTrimRight(StringTrimLeft($a_Saved_Names_of_Devices[$y][1], 1), 1)
		If $sName = $a_Saved_Names_of_Devices[$y][1] Then
			Return $a_Saved_Names_of_Devices[$y][0] ; return ref number
		EndIf
	Next
	Return 0
EndFunc   ;==>_Device_Names_Get_Ref_From_INI

Func _Device_Names_Add_to_INI()
	Local $a_Saved_Names_of_Devices

	If $b_Change_Recording_Device = True Then
		$a_Saved_Names_of_Devices = IniReadSection($s_INI_File_Location, "Names_of_Recording_Devices_" & $s_Computer_ID)
	Else
		$a_Saved_Names_of_Devices = IniReadSection($s_INI_File_Location, "Names_of_Playback_Devices_" & $s_Computer_ID)
	EndIf

	Local $b_Match_Device_Name_found_in_INI = False, $i_Match_Device_Name_found_in_INI_max_Enum = UBound($a_Saved_Names_of_Devices)

	For $i = 0 To UBound($aSoundDevices) - 1
		For $y = 1 To UBound($a_Saved_Names_of_Devices) - 1
			$a_Saved_Names_of_Devices[$y][1] = StringTrimRight(StringTrimLeft($a_Saved_Names_of_Devices[$y][1], 1), 1)
			If $aSoundDevices[$i] = $a_Saved_Names_of_Devices[$y][1] Then
				$b_Match_Device_Name_found_in_INI = True
				ExitLoop
			EndIf
		Next
		If $b_Match_Device_Name_found_in_INI = False Then
			$i_Match_Device_Name_found_in_INI_max_Enum += 1
			If $b_Change_Recording_Device = True Then
				IniWrite($s_INI_File_Location, "Names_of_Recording_Devices_" & $s_Computer_ID, 777 & $i_Match_Device_Name_found_in_INI_max_Enum, '"' & $aSoundDevices[$i] & '"')
			Else
				IniWrite($s_INI_File_Location, "Names_of_Playback_Devices_" & $s_Computer_ID, 777 & $i_Match_Device_Name_found_in_INI_max_Enum, '"' & $aSoundDevices[$i] & '"')
			EndIf
		EndIf
		$b_Match_Device_Name_found_in_INI = False
	Next

	; _ArrayDisplay($aSoundDevices)
	; _ArrayDisplay($a_Saved_Names_of_Devices)

EndFunc   ;==>_Device_Names_Add_to_INI

Global $h_GUI

If $a_EnumChildWindows[0][0] > 0 Then

	$iListviewSize = ControlListView($a_EnumChildWindows[1][0], "", 0, "GetItemCount")
	If Not $iListviewSize > 0 Then

		If $b_SSD_Is_Running_on_Hidden_Desktop = False Then
			If $b_Change_Recording_Device = True Then
				_MsgBox_SHEx(16, $sGUITitle & " - Error", "Error, no rows found in SysListView32 control of the 'Sound' control panel dialog, no recording devices installed?", 0, $hwnd_AutoIt)
			Else
				_MsgBox_SHEx(16, $sGUITitle & " - Error", "Error, no rows found in SysListView32 control of the 'Sound' control panel dialog, no output devices installed?", 0, $hwnd_AutoIt)
			EndIf
		EndIf

		_Exit(2)

	EndIf
	Dim $aSoundDevices[$iListviewSize]
	For $i = 0 To $iListviewSize - 1
		For $y = 0 To 1
			$aSoundDevices[$i] &= ControlListView($a_EnumChildWindows[1][0], "", 0, "GetText", $i, $y)
			If $y = 0 Then $aSoundDevices[$i] &= ", "
		Next
	Next

	If StringInStr($__CmdLineRaw, "_First_Startup", 2) Then
		_Device_Names_Add_to_INI()
		_Exit(58)
	EndIf

	If StringLen($iDeviceNo) = 0 Then ; not set by command-line > show GUI

		If $b_SSD_Is_Running_on_Hidden_Desktop = True Then
			_Exit(682)
		EndIf

		_Device_Names_Add_to_INI()

		$h_GUI = GUICreate($sGUITitle, 450, 390 + 25)

		_WinAPI_SetWindowLong($hWnd_Sound, $GWL_HWNDPARENT, $h_GUI)

		If @Compiled Then
			$c_icon_Hyperlink_FunkEu = GUICtrlCreateIcon(@ScriptName, 0, 6, 5, 32, 32)
		Else
			$c_icon_Hyperlink_FunkEu = GUICtrlCreateIcon(@ScriptDir & "\ssd_logo.ico", 0, 6, 5, 32, 32)
		EndIf
		GUICtrlSetCursor(-1, 0)

		$c_Picture_Title = GUICtrlCreatePic('', 50, 6, 395, 26)
		_DrawShadowText_Title($c_Picture_Title)
		GUICtrlSetCursor(-1, 0)

		$c_label_slider_volume_titel = GUICtrlCreateLabel("Volume", 401, 42, 40, 20, $SS_CENTER)
		GUICtrlSetFont(-1, 7, 800, 0, "Arial")
		$c_slider_volume = GUICtrlCreateSlider(400, 58, 40, 173, BitOR($TBS_AUTOTICKS, $TBS_VERT, $TBS_BOTH))
		GUICtrlSetLimit(-1, 100, 0) ; change min/max value
		GUICtrlSetData(-1, Abs(100 - IniRead($s_INI_File_Location, "GUI", "c_slider_volume", 100)))
		$c_label_slider_volume = GUICtrlCreateLabel("", 400, 231, 40, 20, $SS_CENTER)
		GUICtrlSetFont(-1, 7, 400, 0, "Arial")
		GUICtrlSetData($c_label_slider_volume, Abs(100 - GUICtrlRead($c_slider_volume)) & "%")
		$c_checkbox_volume_mute = GUICtrlCreateCheckbox("Mute", 400, 248, 40, 20)
		GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")
		GUICtrlSetState(-1, IniRead($s_INI_File_Location, "GUI", "c_checkbox_volume_mute", $GUI_UNCHECKED))

		GUICtrlCreateLabel("Visit:", 6, 42, 25, 17)
		GUICtrlSetFont(-1, 8, 400, 0, "Arial")
		GUICtrlSetState(-1, $GUI_DISABLE)

		$c_label_Hyperlink_FunkEu = GUICtrlCreateLabel("http://www.funk.eu", 33, 42, 98, 17)
		GUICtrlSetColor($c_label_Hyperlink_FunkEu, 0xA7A6AA)
		GUICtrlSetFont($c_label_Hyperlink_FunkEu, 8, 400, 4, "Arial")
		GUICtrlSetColor($c_label_Hyperlink_FunkEu, 0x518CB8)
		GUICtrlSetCursor($c_label_Hyperlink_FunkEu, 0)

		$c_label_Hyperlink_FunkEu2 = GUICtrlCreateLabel("", 7, 66, 386, 28 + 6)
		GUICtrlSetBkColor(-1, 0xFCFCFC)
		GUICtrlSetCursor(-1, 0)

		GUICtrlCreateLabel("SSD enables you to set the default Playback or Recording Sound Devices via shortcuts or the command line. For any feedback please visit my website.", 10, 69, 380, 28)
		GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
		GUICtrlSetState(-1, $GUI_DISABLE)
		GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")

		$c_combo_devices = GUICtrlCreateCombo("", 10, 135, 380, 20, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
		For $i = 0 To UBound($aSoundDevices) - 1
			GUICtrlSetData($c_combo_devices, $aSoundDevices[$i])
		Next

		Local $b_Standard_Found = False
		ControlSend($hWnd_Sound, "", "", "{HOME}")
		For $i = 0 To UBound($aSoundDevices) - 1
			; Determine "Default Device"
			Sleep(50)
			If ControlCommand($hWnd_Sound, "", "Button2", "IsEnabled", "") = 0 And ControlCommand($hWnd_Sound, "", "Button1", "IsEnabled", "") = 1 Then ; returns 0 if device is active and 1 if device is not active
				GUICtrlSetData($c_combo_devices, $aSoundDevices[$i])
				$b_Standard_Found = True
				ExitLoop
			EndIf
			ControlSend($hWnd_Sound, "", "", "{DOWN}")
		Next
		If $b_Standard_Found = False Then GUICtrlSetData($c_combo_devices, $aSoundDevices[0])

		Global $c_Button_Refresh = GUICtrlCreateButton("Refresh", 320, 114, 70, 17)
		GUICtrlSetFont(-1, 8.5, 800, 0, "Arial")

		Global $c_checkbox_create_shortcut_hidden = GUICtrlCreateCheckbox("Perform Device Switch Hidden", 11, 188, 190)
		GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")
		GUICtrlSetState(-1, IniRead($s_INI_File_Location, "GUI", "c_checkbox_create_shortcut_hidden", $GUI_CHECKED))
		GUICtrlSetTip(-1, "The Device Switch is performed by simulating mouse clicks in the Sound control panel. When checked," & @CRLF & "the process will create a secondary hidden desktop and perform the actual switch there (invisible to the user).")

		Global $c_checkbox_select_custom_icon = GUICtrlCreateCheckbox("Select Shortcut Icon", 220, 188, 170)
		GUICtrlSetState(-1, IniRead($s_INI_File_Location, "GUI", "c_checkbox_select_custom_icon", $GUI_CHECKED))
		GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")
		GUICtrlSetTip(-1, "When checked SSD will ask you to select a custom icon," & @CRLF & "otherwise the SSD will use the default icons.")

		Global $c_checkbox_select_shortcut_location = GUICtrlCreateCheckbox("Select Shortcut Location", 220, 208, 170)
		GUICtrlSetState(-1, IniRead($s_INI_File_Location, "GUI", "c_checkbox_select_shortcut_location", $GUI_CHECKED))
		GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")
		GUICtrlSetTip(-1, "When checked SSD will ask for a location to create the shortcut at," & @CRLF & "otherwise the shortcut is created in the program directory.")

		Global $c_checkbox_use_names_for_switch = GUICtrlCreateCheckbox("Use Device Names for Switching", 10, 208, 205)
		GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")
		GUICtrlSetState(-1, IniRead($s_INI_File_Location, "GUI", "c_checkbox_use_names_for_switch", $GUI_CHECKED))
		GUICtrlSetTip(-1, "SSD etiher uses the number of the device in the device list (e.g. third from top)," & @CRLF & "or check this box and SSD will write a reference to SSD.ini and try to identify" & @CRLF & "the device by name (useful when your device list changes frequently).")

		Global $c_checkbox_show_tray_notification_on_change = GUICtrlCreateCheckbox("On Switch - Show Tray Notification", 10, 228, 205)
		GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")
		GUICtrlSetState(-1, IniRead($s_INI_File_Location, "GUI", "c_checkbox_show_tray_notification_on_change", $GUI_UNCHECKED))
		GUICtrlSetTip(-1, "SSD will show a 3 seconds notification bubble in the tray on change of a device")

		Global $c_checkbox_show_tooltip_at_mousepos_on_change = GUICtrlCreateCheckbox("On Switch - Show ToolTip", 10, 248, 205)
		GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")
		GUICtrlSetState(-1, IniRead($s_INI_File_Location, "GUI", "c_checkbox_show_tooltip_at_mousepos_on_change", $GUI_CHECKED))
		GUICtrlSetTip(-1, "SSD will show a 3 seconds tooltip at the current mouse position on change of a device")

		$c_checkbox_Close_Control_Panel = GUICtrlCreateCheckbox("Close Sound Control Panel", 220, 228, 170)
		GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")
		GUICtrlSetState(-1, IniRead($s_INI_File_Location, "GUI", "c_checkbox_Close_Control_Panel", $GUI_CHECKED))
		GUICtrlSetTip(-1, "When checked the Sound Control Panel will be closed on exit of SSD (ony in GUI mode).")

		Global $c_checkbox_save_and_reapply_volume = GUICtrlCreateCheckbox("Set Specific Target Volume", 220, 248, 170)
		GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")
		GUICtrlSetTip(-1, "1) Click button to start shortcut creation" & @CRLF _
				 & "2) Select Volume for device" & @CRLF _
				 & "3) Select Device in Dropdown list" & @CRLF & @CRLF & "For toggle shortcuts you can activate / deactivate this feature separately for each device.")
		GUICtrlSetState(-1, IniRead($s_INI_File_Location, "GUI", "c_checkbox_save_and_reapply_volume", $GUI_UNCHECKED))
		If GUICtrlRead($c_checkbox_save_and_reapply_volume) = $GUI_UNCHECKED Then
			GUICtrlSetState($c_slider_volume, $GUI_DISABLE)
			GUICtrlSetState($c_label_slider_volume_titel, $GUI_DISABLE)
			GUICtrlSetState($c_label_slider_volume, $GUI_DISABLE)
			GUICtrlSetState($c_checkbox_volume_mute, $GUI_DISABLE)
		EndIf


		If $b_Change_Recording_Device = True Then

			$c_button_switch = GUICtrlCreateButton("Switch to Playback Devices", 192, 38, 200, 20)
			GUICtrlSetFont(-1, 8, 600, 0, "Arial")

			GUICtrlCreateLabel("Select Recording Device to be set:", 10, 113, 290, 20)
			GUICtrlSetFont(-1, 11, 800, 4, "Arial")

			$c_button_start = GUICtrlCreateButton("a) Create Shortcut to change to this Recording Device", 10, 279, 430, 30)
			GUICtrlSetFont(-1, 8, 600, 0, "Arial")

			$c_checkbox_change_audio_Playback_device = GUICtrlCreateCheckbox("Change Recording Device", 10, 156, 170)
			GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")
			GUICtrlSetState(-1, IniRead($s_INI_File_Location, "GUI", "c_checkbox_change_audio_Playback_device", $GUI_CHECKED))
			GUICtrlSetTip(-1, "When checked the Recording Device will be changed to a new default.")

			$c_button_start_toggle = GUICtrlCreateButton("b) Create Shortcut to toggle between two Recording Devices", 10, 312, 430, 30) ;, 0x2000); $BS_MULTILINE = 0x2000
			GUICtrlSetFont(-1, 8, 600, 0, "Arial")

		Else

			$c_button_switch = GUICtrlCreateButton("Switch to Recording Devices", 192, 38, 200, 20)
			GUICtrlSetFont(-1, 8, 600, 0, "Arial")

			GUICtrlCreateLabel("Select Playback Device to be set:", 10, 113, 290, 20)
			GUICtrlSetFont(-1, 11, 800, 4, "Arial")

			$c_button_start = GUICtrlCreateButton("a) Create Shortcut to change to this Playback Device", 10, 279, 430, 30)
			GUICtrlSetFont(-1, 8, 600, 0, "Arial")

			$c_checkbox_change_audio_Playback_device = GUICtrlCreateCheckbox("Change Audio Device", 10, 156, 170)
			GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")
			GUICtrlSetState(-1, IniRead($s_INI_File_Location, "GUI", "c_checkbox_change_audio_Playback_device", $GUI_CHECKED))
			GUICtrlSetTip(-1, "When checked the Audio Device will be changed to a new default")

			$c_button_start_toggle = GUICtrlCreateButton("b) Create Shortcut to toggle between two Playback Devices", 10, 312, 430, 30) ;, 0x2000); $BS_MULTILINE = 0x2000
			GUICtrlSetFont(-1, 8, 600, 0, "Arial")

		EndIf

		If $iListviewSize < 2 Then GUICtrlSetState($c_button_start_toggle, $GUI_DISABLE)

		$c_checkbox_change_communication_device = GUICtrlCreateCheckbox("Change Communication Device", 220, 156, 170)
		GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")
		GUICtrlSetState(-1, IniRead($s_INI_File_Location, "GUI", "c_checkbox_change_communication_device", $GUI_CHECKED))
		GUICtrlSetTip(-1, "When checked the Communication Device will be changed to a new default")

		$c_button_merge = GUICtrlCreateButton("c) Merge two existing Playback and Recording Device Switch Shortcuts", 10, 360, 430, 25)
		GUICtrlSetTip(-1, "The command line arguments of both shortcuts will be merged" & @CRLF & "into a "".bat"" file and a shortcut pointing to that file will be created." & @CRLF & @CRLF & "Through this SSD will be started twice, applying both changes in separate program runs.")
		GUICtrlSetFont(-1, 8.5, 400, 0, "Arial")

		$c_Hyperlink_Donate_Picture = GUICtrlCreatePic("", 10, 392, 100, 20, $SS_NOTIFY)
		GUICtrlSetCursor($c_Hyperlink_Donate_Picture, 0)

		$c_Hyperlink_CC = GUICtrlCreatePic("", 360, 395, 80, 15, $SS_NOTIFY)
		GUICtrlSetCursor($c_Hyperlink_CC, 0)

		_GDIPlus_Startup()
		$hBitmap_License = _Load_BMP_From_Mem(_BinaryString_Picture_License(), True)
		_WinAPI_DeleteObject(GUICtrlSendMsg($c_Hyperlink_CC, $STM_SETIMAGE, $IMAGE_BITMAP, $hBitmap_License))
		_WinAPI_DeleteObject($hBitmap_License)
		$hBitmap_Donate = _Load_BMP_From_Mem(_BinaryString_Picture_Donate(), True)
		_WinAPI_DeleteObject(GUICtrlSendMsg($c_Hyperlink_Donate_Picture, $STM_SETIMAGE, $IMAGE_BITMAP, $hBitmap_Donate))
		_WinAPI_DeleteObject($hBitmap_Donate)
		_GDIPlus_Shutdown()

		GUIRegisterMsg($WM_COMMAND, "My_WM_COMMAND")
		Global $s_ToolTip_Text, $s_ToolTip_Title
		GUIRegisterMsg($WM_ACTIVATE, "WM_ACTIVATE")
		GUISetState(@SW_SHOW)

		Local $iDeviceNo_1, $iDeviceNo_2
		Local $iDeviceNo_1_Volume, $iDeviceNo_2_Volume
		Local $i_GUI_Msg

		While 1
			$i_GUI_Msg = GUIGetMsg()
			Switch $i_GUI_Msg
				Case $c_checkbox_save_and_reapply_volume
					If GUICtrlRead($c_checkbox_save_and_reapply_volume) = $GUI_UNCHECKED Then
						GUICtrlSetState($c_slider_volume, $GUI_DISABLE)
						GUICtrlSetState($c_label_slider_volume, $GUI_DISABLE)
						GUICtrlSetState($c_label_slider_volume_titel, $GUI_DISABLE)
						GUICtrlSetState($c_checkbox_volume_mute, $GUI_DISABLE)
					Else
						GUICtrlSetState($c_slider_volume, $GUI_ENABLE)
						GUICtrlSetState($c_label_slider_volume, $GUI_ENABLE)
						GUICtrlSetState($c_label_slider_volume_titel, $GUI_ENABLE)
						GUICtrlSetState($c_checkbox_volume_mute, $GUI_ENABLE)
					EndIf

				Case $c_slider_volume
					GUICtrlSetData($c_label_slider_volume, Abs(100 - GUICtrlRead($c_slider_volume)) & "%")

				Case $c_Hyperlink_CC
					ShellExecute("http://creativecommons.org/licenses/by-nc-nd/3.0/us/")

				Case $c_Hyperlink_Donate_Picture
					ShellExecute("https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=smf%40funk%2eeu&item_name=Thank%20you%20for%20your%20donation%20for%20SSD!&no_shipping=0&no_note=1&tax=0&currency_code=EUR&lc=US&bn=PP%2dDonationsBF&charset=UTF%2d8")

				Case $c_button_merge
					_Merge_Two_Shortcuts()

				Case $c_Button_Refresh
					_GUI_IniWrite_All()
					WinClose($hWnd_Sound)
					AutoItWinSetTitle("")
					If $b_Change_Recording_Device = True Then
						Run(@ScriptFullPath & " _Recording", @ScriptDir)
					Else
						Run(@ScriptFullPath, @ScriptDir)
					EndIf

					_Exit(58)

				Case $GUI_EVENT_CLOSE
					_GUI_IniWrite_All()
					GUIRegisterMsg($WM_ACTIVATE, "")
					Local $i_checkbox_Close_Control_Panel = GUICtrlRead($c_checkbox_Close_Control_Panel)
					GUIDelete()
					_Exit(0, $i_checkbox_Close_Control_Panel)

				Case $c_button_switch
					_GUI_IniWrite_All()
					WinClose($hWnd_Sound)
					AutoItWinSetTitle("")
					If $b_Change_Recording_Device = True Then
						Run(@ScriptFullPath, @ScriptDir)
					Else
						Run(@ScriptFullPath & " _Recording", @ScriptDir)
					EndIf

					_Exit(57)

				Case $c_checkbox_change_communication_device
					If GUICtrlRead($c_checkbox_change_communication_device) = $GUI_UNCHECKED Then
						GUICtrlSetState($c_checkbox_change_audio_Playback_device, $GUI_CHECKED)
					EndIf

				Case $c_checkbox_change_audio_Playback_device
					If GUICtrlRead($c_checkbox_change_audio_Playback_device) = $GUI_UNCHECKED Then
						GUICtrlSetState($c_checkbox_change_communication_device, $GUI_CHECKED)
					EndIf

				Case $c_button_start_toggle, $c_combo_devices, $c_button_start

					If GUICtrlRead($c_checkbox_change_communication_device) = $GUI_UNCHECKED And GUICtrlRead($c_checkbox_change_audio_Playback_device) = $GUI_UNCHECKED Then
						_MsgBox_SHEx(16, $sGUITitle & " - Error", "Either Audio or Communication Device has to be checked to create shortcut.", 0, $h_GUI)
					Else

						If $i_GUI_Msg = $c_button_start_toggle And (StringInStr(GUICtrlRead($c_button_start_toggle), "Select first") Or StringInStr(GUICtrlRead($c_button_start_toggle), "Select second")) Then

							If $b_Change_Recording_Device = True Then
								GUICtrlSetData($c_button_start_toggle, "b) Create Shortcut to toggle between two Recording Devices")
							Else
								GUICtrlSetData($c_button_start_toggle, "b) Create Shortcut to toggle between two Playback Devices")
							EndIf
							GUICtrlSetState($c_checkbox_select_custom_icon, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_create_shortcut_hidden, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_change_audio_Playback_device, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_change_communication_device, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_show_tray_notification_on_change, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_show_tooltip_at_mousepos_on_change, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_use_names_for_switch, $GUI_ENABLE)
							GUICtrlSetState($c_button_merge, $GUI_ENABLE)
							GUICtrlSetState($c_button_start, $GUI_ENABLE)
							GUICtrlSetState($c_button_switch, $GUI_ENABLE)
							GUICtrlSetState($c_Button_Refresh, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_select_shortcut_location, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_Close_Control_Panel, $GUI_ENABLE)
							$s_ToolTip_Title = ""
							ToolTip("")

						ElseIf $i_GUI_Msg = $c_button_start And StringInStr(GUICtrlRead($c_button_start), "Select") Then

							If $b_Change_Recording_Device = True Then
								GUICtrlSetData($c_button_start, "a) Create Shortcut to change to this Recording Device")
							Else
								GUICtrlSetData($c_button_start, "a) Create Shortcut to change to this Playback Device")
							EndIf

							GUICtrlSetState($c_checkbox_select_custom_icon, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_create_shortcut_hidden, $GUI_ENABLE)

							If $iListviewSize > 1 Then GUICtrlSetState($c_button_start_toggle, $GUI_ENABLE)

							GUICtrlSetState($c_button_switch, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_change_audio_Playback_device, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_change_communication_device, $GUI_ENABLE)
							GUICtrlSetState($c_button_merge, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_show_tray_notification_on_change, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_show_tooltip_at_mousepos_on_change, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_use_names_for_switch, $GUI_ENABLE)
							GUICtrlSetState($c_Button_Refresh, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_select_shortcut_location, $GUI_ENABLE)
							GUICtrlSetState($c_checkbox_Close_Control_Panel, $GUI_ENABLE)
							$s_ToolTip_Title = ""
							ToolTip("")

						Else

							If $i_GUI_Msg = $c_combo_devices Then

								If Not (StringInStr(GUICtrlRead($c_button_start_toggle), "Select") Or StringInStr(GUICtrlRead($c_button_start), "Select")) Then ContinueLoop

								If StringInStr(GUICtrlRead($c_button_start_toggle), "Select first") Then
									$iDeviceNo_1 = 0
									$iDeviceNo_1_Volume = ""
									For $i = 0 To UBound($aSoundDevices) - 1
										If GUICtrlRead($c_combo_devices) = $aSoundDevices[$i] Then
											$iDeviceNo_1 = $i + 1
											ExitLoop
										EndIf
									Next

									If $iDeviceNo_1 Then
										Local $a_WinPos = WinGetPos($h_GUI)
										$s_ToolTip_Title = "First Device selected:"
										$s_ToolTip_Text = " " & GUICtrlRead($c_combo_devices)
										ToolTip($s_ToolTip_Text, $a_WinPos[0], $a_WinPos[1], $s_ToolTip_Title, 1)

										If GUICtrlRead($c_checkbox_save_and_reapply_volume) = $GUI_CHECKED Then
											If GUICtrlRead($c_checkbox_volume_mute) = $GUI_CHECKED Then ; SetMute
												$iDeviceNo_1_Volume = 3331
											Else
												$iDeviceNo_1_Volume = 3330
											EndIf
											$iDeviceNo_1_Volume = Int($iDeviceNo_1_Volume & Abs(100 - GUICtrlRead($c_slider_volume)))
										EndIf

									EndIf

								ElseIf StringInStr(GUICtrlRead($c_button_start_toggle), "Select second") Then
									$iDeviceNo_2 = 0
									$iDeviceNo_2_Volume = ""
									For $i = 0 To UBound($aSoundDevices) - 1
										If GUICtrlRead($c_combo_devices) = $aSoundDevices[$i] Then
											$iDeviceNo_2 = $i + 1
											ExitLoop
										EndIf
									Next

									If GUICtrlRead($c_checkbox_save_and_reapply_volume) = $GUI_CHECKED Then
										If GUICtrlRead($c_checkbox_volume_mute) = $GUI_CHECKED Then ; SetMute
											$iDeviceNo_2_Volume = 3331
										Else
											$iDeviceNo_2_Volume = 3330
										EndIf
										$iDeviceNo_2_Volume = Int($iDeviceNo_2_Volume & Abs(100 - GUICtrlRead($c_slider_volume)))
									EndIf

								ElseIf StringInStr(GUICtrlRead($c_button_start), "Select") Then

									For $i = 0 To UBound($aSoundDevices) - 1
										If GUICtrlRead($c_combo_devices) = $aSoundDevices[$i] Then
											$iDeviceNo = $i + 1
											If Not $g_Macro_Compiled Then
												_MsgBox_SHEx(16, $sGUITitle & " - Error", "Shortcuts are only created if SSD is compiled. #1", 0, $h_GUI)
											Else

												If GUICtrlRead($c_checkbox_save_and_reapply_volume) = $GUI_CHECKED Then
													If GUICtrlRead($c_checkbox_volume_mute) = $GUI_CHECKED Then ; SetMute
														Local $i_Volume_Info = 1
													Else
														Local $i_Volume_Info = 0
													EndIf
													$i_Volume_Info = Int(333 & $i_Volume_Info & Abs(100 - GUICtrlRead($c_slider_volume)))
													_func_create_shortcut_standard($iDeviceNo, $aSoundDevices[$i], $i_Volume_Info)
												Else
													_func_create_shortcut_standard($iDeviceNo, $aSoundDevices[$i], "")
												EndIf

											EndIf
										EndIf
									Next

								EndIf

							EndIf

							If $i_GUI_Msg = $c_button_start And Not StringInStr(GUICtrlRead($c_button_start), "Select") Then

								GUICtrlSetState($c_checkbox_select_custom_icon, $GUI_DISABLE)
								GUICtrlSetState($c_checkbox_create_shortcut_hidden, $GUI_DISABLE)
								GUICtrlSetState($c_button_start_toggle, $GUI_DISABLE)
								GUICtrlSetState($c_button_switch, $GUI_DISABLE)
								GUICtrlSetState($c_checkbox_change_audio_Playback_device, $GUI_DISABLE)
								GUICtrlSetState($c_checkbox_change_communication_device, $GUI_DISABLE)
								GUICtrlSetState($c_button_merge, $GUI_DISABLE)
								GUICtrlSetState($c_checkbox_show_tray_notification_on_change, $GUI_DISABLE)
								GUICtrlSetState($c_checkbox_show_tooltip_at_mousepos_on_change, $GUI_DISABLE)
								GUICtrlSetState($c_checkbox_use_names_for_switch, $GUI_DISABLE)
								GUICtrlSetState($c_Button_Refresh, $GUI_DISABLE)
								GUICtrlSetState($c_checkbox_select_shortcut_location, $GUI_DISABLE)
								GUICtrlSetState($c_checkbox_Close_Control_Panel, $GUI_DISABLE)

								Local $a_WinPos = WinGetPos($h_GUI)
								If $b_Change_Recording_Device = True Then
									GUICtrlSetData($c_button_start, "Select Recording Device to activate by shortcut")
									$s_ToolTip_Title = "Select Recording Device"
									$s_ToolTip_Text = " to activate by shortcut"
									ToolTip($s_ToolTip_Text, $a_WinPos[0], $a_WinPos[1], $s_ToolTip_Title, 1)
								Else
									GUICtrlSetData($c_button_start, "Select Playback Device to activate by shortcut")
									$s_ToolTip_Title = "Select Playback Device"
									$s_ToolTip_Text = " to activate by shortcut"
									ToolTip($s_ToolTip_Text, $a_WinPos[0], $a_WinPos[1], $s_ToolTip_Title, 1)
								EndIf

								_GUICtrlComboBox_ShowDropDown($c_combo_devices, True)

							Else

								If StringInStr(GUICtrlRead($c_button_start), "Select") Then
									If $b_Change_Recording_Device = True Then
										GUICtrlSetData($c_button_start, "a) Create Shortcut to change to this Recording Device")
									Else
										GUICtrlSetData($c_button_start, "a) Create Shortcut to change to this Playback Device")
									EndIf

									GUICtrlSetState($c_checkbox_select_custom_icon, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_create_shortcut_hidden, $GUI_ENABLE)

									If $iListviewSize > 1 Then GUICtrlSetState($c_button_start_toggle, $GUI_ENABLE)

									GUICtrlSetState($c_button_switch, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_change_audio_Playback_device, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_change_communication_device, $GUI_ENABLE)
									GUICtrlSetState($c_button_merge, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_show_tray_notification_on_change, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_show_tooltip_at_mousepos_on_change, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_use_names_for_switch, $GUI_ENABLE)
									GUICtrlSetState($c_Button_Refresh, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_select_shortcut_location, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_Close_Control_Panel, $GUI_ENABLE)
									$s_ToolTip_Title = ""
									ToolTip("")

								ElseIf Not (StringInStr(GUICtrlRead($c_button_start_toggle), "Select first") Or StringInStr(GUICtrlRead($c_button_start_toggle), "Select second")) Then
									; Select first

									GUICtrlSetState($c_checkbox_select_custom_icon, $GUI_DISABLE)
									GUICtrlSetState($c_checkbox_create_shortcut_hidden, $GUI_DISABLE)
									GUICtrlSetState($c_button_start, $GUI_DISABLE)
									GUICtrlSetState($c_button_switch, $GUI_DISABLE)
									GUICtrlSetState($c_checkbox_change_audio_Playback_device, $GUI_DISABLE)
									GUICtrlSetState($c_checkbox_change_communication_device, $GUI_DISABLE)
									GUICtrlSetState($c_button_merge, $GUI_DISABLE)
									GUICtrlSetState($c_checkbox_show_tray_notification_on_change, $GUI_DISABLE)
									GUICtrlSetState($c_checkbox_show_tooltip_at_mousepos_on_change, $GUI_DISABLE)
									GUICtrlSetState($c_checkbox_use_names_for_switch, $GUI_DISABLE)
									GUICtrlSetState($c_Button_Refresh, $GUI_DISABLE)
									GUICtrlSetState($c_checkbox_select_shortcut_location, $GUI_DISABLE)
									GUICtrlSetState($c_checkbox_Close_Control_Panel, $GUI_DISABLE)

									Local $a_WinPos = WinGetPos($h_GUI)
									If $b_Change_Recording_Device = True Then
										GUICtrlSetData($c_button_start_toggle, "Select first Recording Device to toggle to")
										$s_ToolTip_Title = "Select first Recording Device"
										$s_ToolTip_Text = " to toggle to"
										ToolTip($s_ToolTip_Text, $a_WinPos[0], $a_WinPos[1], $s_ToolTip_Title, 1)
									Else
										GUICtrlSetData($c_button_start_toggle, "Select first Playback Device to toggle to")
										$s_ToolTip_Title = "Select first Playback Device"
										$s_ToolTip_Text = " to toggle to"
										ToolTip($s_ToolTip_Text, $a_WinPos[0], $a_WinPos[1], $s_ToolTip_Title, 1)
									EndIf

									_GUICtrlComboBox_ShowDropDown($c_combo_devices, True)

								ElseIf Not StringInStr(GUICtrlRead($c_button_start_toggle), "Select second") Then
									; Select second

									Local $a_WinPos = WinGetPos($h_GUI)
									If $b_Change_Recording_Device = True Then
										GUICtrlSetData($c_button_start_toggle, "Select second Recording Device to toggle to")
										$s_ToolTip_Title = "Select second Recording Device"
										$s_ToolTip_Text = " to toggle to"
										ToolTip($s_ToolTip_Text, $a_WinPos[0], $a_WinPos[1], $s_ToolTip_Title, 1)
									Else
										GUICtrlSetData($c_button_start_toggle, "Select second Playback Device to toggle to")
										$s_ToolTip_Title = "Select second Playback Device"
										$s_ToolTip_Text = " to toggle to"
										ToolTip($s_ToolTip_Text, $a_WinPos[0], $a_WinPos[1], $s_ToolTip_Title, 1)
									EndIf

									_GUICtrlComboBox_ShowDropDown($c_combo_devices, True)

								Else

									If $iDeviceNo_1 = $iDeviceNo_2 Or $iDeviceNo_1 = 0 Or $iDeviceNo_2 = 0 Then
										_MsgBox_SHEx(16, $sGUITitle & " - Error", "Select two different devices to create toggle shortcut." & @CRLF & @CRLF & $iDeviceNo_1 & "-" & $iDeviceNo_2, 0, $h_GUI)
									Else
										_func_create_shortcut_toogle($iDeviceNo_1, $aSoundDevices[$iDeviceNo_1 - 1], $iDeviceNo_1_Volume, $iDeviceNo_2, $aSoundDevices[$iDeviceNo_2 - 1], $iDeviceNo_2_Volume)
									EndIf

									If $b_Change_Recording_Device = True Then
										GUICtrlSetData($c_button_start_toggle, "b) Create Shortcut to toggle between two Recording Devices")
									Else
										GUICtrlSetData($c_button_start_toggle, "b) Create Shortcut to toggle between two Playback Devices")
									EndIf

									GUICtrlSetState($c_checkbox_select_custom_icon, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_create_shortcut_hidden, $GUI_ENABLE)
									GUICtrlSetState($c_button_start, $GUI_ENABLE)
									GUICtrlSetState($c_button_switch, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_change_audio_Playback_device, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_change_communication_device, $GUI_ENABLE)
									GUICtrlSetState($c_button_merge, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_show_tray_notification_on_change, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_show_tooltip_at_mousepos_on_change, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_use_names_for_switch, $GUI_ENABLE)
									GUICtrlSetState($c_Button_Refresh, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_select_shortcut_location, $GUI_ENABLE)
									GUICtrlSetState($c_checkbox_Close_Control_Panel, $GUI_ENABLE)

									$s_ToolTip_Title = ""
									ToolTip("")

								EndIf
							EndIf
						EndIf
					EndIf

			EndSwitch

			If Not WinExists($hWnd_Sound) Then
				_GUI_IniWrite_All()
				_Exit(3)
			EndIf

		WEnd

	EndIf

	; Perform the actual switch !!!

	If $iDeviceNo > 0 And WinExists($hWnd_Sound) Then

		WinActivate($hWnd_Sound)

		Local $iDeviceNo_Volume[1][2] = [[-1, -1]]

		If StringInStr($iDeviceNo, 999) Then ; toggle mode

			If Not IniWrite(@ScriptDir & "\SSD.ini", "Settings", "SSD_Version", $sGUITitle & "-" & $iBuildnumber) Then
				If $b_SSD_Is_Running_on_Hidden_Desktop Then
					_Exit(674)
				Else
					_MsgBox_SHEx(16, $sGUITitle & ' - Error', "SSD Toggle Mode works only, if the SSD.ini file is writeable." & @CRLF & @CRLF & "The SSD.ini file is created in the same directory as SSD.exe resides in, move SSD.exe to a writeable location." & @CRLF & @CRLF & $s_INI_File_Location, 0, $hwnd_AutoIt)
					Exit 43
				EndIf
			EndIf

			Local $aDeviceNo_Toggle = StringSplit($iDeviceNo, "999", 1)
			If $aDeviceNo_Toggle[0] <> 2 Then _Exit(99)
			; _ArrayDisplay($aDeviceNo_Toggle)

			Local $iDeviceNo_Toggle[2]

			$iDeviceNo_Toggle[0] = $aDeviceNo_Toggle[1]
			$iDeviceNo_Toggle[1] = $aDeviceNo_Toggle[2]

			Local $iDeviceNo_Toggle_Volume[2][2] = [[-1, -1], [-1, -1]]

			If StringInStr($iDeviceNo_Toggle[0], 333) Then ; Extract Volume info
				Local $aDeviceNo_Toggle = StringSplit($iDeviceNo_Toggle[0], "333", 1)
				If $aDeviceNo_Toggle[0] <> 2 Then _Exit(931)
				$iDeviceNo_Toggle[0] = $aDeviceNo_Toggle[1]
				$iDeviceNo_Toggle_Volume[0][0] = StringLeft($aDeviceNo_Toggle[2], 1)
				$iDeviceNo_Toggle_Volume[0][1] = StringTrimLeft($aDeviceNo_Toggle[2], 1)
			EndIf

			If StringInStr($iDeviceNo_Toggle[1], 333) Then ; Extract Volume info
				Local $aDeviceNo_Toggle = StringSplit($iDeviceNo_Toggle[1], "333", 1)
				If $aDeviceNo_Toggle[0] <> 2 Then _Exit(932)
				; _ArrayDisplay($aDeviceNo_Toggle)
				$iDeviceNo_Toggle[1] = $aDeviceNo_Toggle[1]
				$iDeviceNo_Toggle_Volume[1][0] = StringLeft($aDeviceNo_Toggle[2], 1)
				$iDeviceNo_Toggle_Volume[1][1] = StringTrimLeft($aDeviceNo_Toggle[2], 1)
			EndIf

			; ConsoleWrite($iDeviceNo_Toggle[0] & @TAB & $iDeviceNo_Toggle_Volume[0][0] & @TAB & $iDeviceNo_Toggle_Volume[0][1] & @CRLF)
			; ConsoleWrite($iDeviceNo_Toggle[1] & @TAB & $iDeviceNo_Toggle_Volume[1][0] & @TAB & $iDeviceNo_Toggle_Volume[1][1] & @CRLF)

			If StringInStr($iDeviceNo_Toggle[0], 777) Then ; get by name
				; ConsoleWrite($iDeviceNo_Toggle[0] & @CRLF)
				$iDeviceNo_Toggle[0] = _Device_Names_Get_Name_From_INI($iDeviceNo_Toggle[0])
				If $iDeviceNo_Toggle[0] = 0 Then
					If $b_SSD_Is_Running_on_Hidden_Desktop = False Then
						_Exit(771)
					Else
						_MsgBox_SHEx(16, $sGUITitle & " - Error", "SSD shortcut indicates that device name is to be used, but device name #1 was not found in SSD.ini." & @CRLF & @CRLF & $iDeviceNo, 0, $hwnd_AutoIt)
					EndIf
				EndIf
			EndIf

			If StringInStr($iDeviceNo_Toggle[1], 777) Then ; get by name
				$iDeviceNo_Toggle[1] = _Device_Names_Get_Name_From_INI($iDeviceNo_Toggle[1])
				If $iDeviceNo_Toggle[1] = 0 Then ; EXIT ERROR, name ref not found
					If $b_SSD_Is_Running_on_Hidden_Desktop = False Then
						_Exit(772)
					Else
						_MsgBox_SHEx(16, $sGUITitle & " - Error", "SSD shortcut indicates that device name is to be used, but device name #2 was not found in SSD.ini." & @CRLF & @CRLF & $iDeviceNo, 0, $hwnd_AutoIt)
					EndIf
				EndIf
				$iDeviceNo = StringReplace($iDeviceNo, "777", "")
			EndIf

			If $b_Change_Recording_Device = True Then
				If IniRead($s_INI_File_Location, "Toggle_Recording", StringReplace($iDeviceNo, "999", "_"), "") = $iDeviceNo_Toggle[0] Then
					IniWrite($s_INI_File_Location, "Toggle_Recording", StringReplace($iDeviceNo, "999", "_"), $iDeviceNo_Toggle[1])
					$iDeviceNo = $iDeviceNo_Toggle[1]
					$iDeviceNo_Volume[0][0] = Int($iDeviceNo_Toggle_Volume[1][0])
					$iDeviceNo_Volume[0][1] = Int($iDeviceNo_Toggle_Volume[1][1])
				Else
					IniWrite($s_INI_File_Location, "Toggle_Recording", StringReplace($iDeviceNo, "999", "_"), $iDeviceNo_Toggle[0])
					$iDeviceNo = $iDeviceNo_Toggle[0]
					$iDeviceNo_Volume[0][0] = Int($iDeviceNo_Toggle_Volume[0][0])
					$iDeviceNo_Volume[0][1] = Int($iDeviceNo_Toggle_Volume[0][1])
				EndIf
			Else
				If IniRead($s_INI_File_Location, "Toggle_Playback", StringReplace($iDeviceNo, "999", "_"), "") = $iDeviceNo_Toggle[0] Then
					IniWrite($s_INI_File_Location, "Toggle_Playback", StringReplace($iDeviceNo, "999", "_"), $iDeviceNo_Toggle[1])
					$iDeviceNo = $iDeviceNo_Toggle[1]
					$iDeviceNo_Volume[0][0] = Int($iDeviceNo_Toggle_Volume[1][0])
					$iDeviceNo_Volume[0][1] = Int($iDeviceNo_Toggle_Volume[1][1])
				Else
					IniWrite($s_INI_File_Location, "Toggle_Playback", StringReplace($iDeviceNo, "999", "_"), $iDeviceNo_Toggle[0])
					$iDeviceNo = $iDeviceNo_Toggle[0]
					$iDeviceNo_Volume[0][0] = Int($iDeviceNo_Toggle_Volume[0][0])
					$iDeviceNo_Volume[0][1] = Int($iDeviceNo_Toggle_Volume[0][1])
				EndIf
			EndIf

			; ConsoleWrite($iDeviceNo & @TAB & $iDeviceNo_Volume[0][0] & @TAB & $iDeviceNo_Volume[0][1] & @CRLF)

		Else

			If StringInStr($iDeviceNo, 333) Then ; Extract Volume info
				Local $a_DeviceNo = StringSplit($iDeviceNo, "333", 1)
				If $a_DeviceNo[0] <> 2 Then
					_Exit(333)
				EndIf
				$iDeviceNo = Int($a_DeviceNo[1])
				$iDeviceNo_Volume[0][0] = Int(StringLeft($a_DeviceNo[2], 1))
				$iDeviceNo_Volume[0][1] = Int(StringTrimLeft($a_DeviceNo[2], 1))
			EndIf

			If StringInStr($iDeviceNo, 777) Then ; get by name single
				Local $iDeviceNo_Org
				$iDeviceNo = _Device_Names_Get_Name_From_INI($iDeviceNo)
				If $iDeviceNo = 0 Then
					If $b_SSD_Is_Running_on_Hidden_Desktop = False Then
						_Exit(773)
					Else
						_MsgBox_SHEx(16, $sGUITitle & " - Error", "SSD shortcut indicates that device name is to be used, but device name was not found in SSD.ini." & @CRLF & @CRLF & $iDeviceNo_Org, 0, $hwnd_AutoIt)
					EndIf
				EndIf
			EndIf

		EndIf

		; ConsoleWrite($iDeviceNo & @TAB & $iDeviceNo_Volume[0][0] & @TAB & $iDeviceNo_Volume[0][1] & @CRLF)

		If $b_NoChange_Audio Then
			; Set Communication Device only
			ControlSend($hWnd_Sound, "", "", "{DOWN " & $iDeviceNo & "}")
			$aPos = ControlGetPos($hWnd_Sound, "", "Button2")
			ControlClick($hWnd_Sound, "", "Button2", "left", 1, $aPos[2] - 1, $aPos[3] - 1)
			ControlSend($hWnd_Sound, "", "Button2", "{DOWN}{DOWN}")
			ControlSend($hWnd_Sound, "", "Button2", "{ENTER}")

		ElseIf $b_NoChange_Com Then
			; Set Sound Device Only
			ControlSend($hWnd_Sound, "", "", "{DOWN " & $iDeviceNo & "}")
			$aPos = ControlGetPos($hWnd_Sound, "", "Button2")
			ControlClick($hWnd_Sound, "", "Button2", "left", 1, $aPos[2] - 1, $aPos[3] - 1)
			ControlSend($hWnd_Sound, "", "Button2", "{DOWN}")
			ControlSend($hWnd_Sound, "", "Button2", "{ENTER}")

		Else
			; Set Sound AND Communication Device
			ControlSend($hWnd_Sound, "", "", "{DOWN " & $iDeviceNo & "}")
			ControlClick($hWnd_Sound, "", "Button2")
			ControlClick($hWnd_Sound, "", "Button4")

		EndIf

		; Adjust Volume
		If $iDeviceNo_Volume[0][0] <> -1 Or $iDeviceNo_Volume[0][1] <> -1 Then
			Sleep(250)
			If $iDeviceNo_Volume[0][0] <> -1 Then
				If $iDeviceNo_Volume[0][0] = 0 Or $iDeviceNo_Volume[0][0] = 1 Then
					; ConsoleWrite("_SetMute" & @TAB & $iDeviceNo_Volume[0][0] & @CRLF)
					_SetMute($iDeviceNo_Volume[0][0])
				EndIf
			EndIf
			If $iDeviceNo_Volume[0][1] <> -1 Then
				If $iDeviceNo_Volume[0][1] > -1 And $iDeviceNo_Volume[0][1] < 101 Then
					; ConsoleWrite("_SetMasterVolumeLevelScalar" & @TAB & $iDeviceNo_Volume[0][1] & @CRLF)
					_SetMasterVolumeLevelScalar($iDeviceNo_Volume[0][1])
				EndIf
			EndIf
		EndIf

		If UBound($aSoundDevices) > $iDeviceNo - 1 Then

			If $b_Change_Recording_Device = False Then
				IniWrite($s_INI_File_Location, "Settings", "Last_Activated_Playback_Device", '"' & StringReplace($aSoundDevices[$iDeviceNo - 1], '"', "'") & '"')
			Else
				IniWrite($s_INI_File_Location, "Settings", "Last_Activated_Recording_Device", '"' & StringReplace($aSoundDevices[$iDeviceNo - 1], '"', "'") & '"')
			EndIf

		EndIf

		If $b_SSD_Is_Running_on_Hidden_Desktop = False Then _Check_for_ToolTip_and_TrayTip_Notifications()

		_Exit(4)

	EndIf

	_Exit((8 & $iDeviceNo & WinExists($hWnd_Sound)))

Else
	If $b_SSD_Is_Running_on_Hidden_Desktop = False Then _MsgBox_SHEx(16, $sGUITitle & " - Error", "SysListView32 control in 'Sound' control panel dialog not found.", 0, $hwnd_AutoIt)
	_Exit(5)
EndIf

Func _GUI_IniWrite_All()
	IniWrite($s_INI_File_Location, "GUI", "c_checkbox_create_shortcut_hidden", GUICtrlRead($c_checkbox_create_shortcut_hidden))
	IniWrite($s_INI_File_Location, "GUI", "c_checkbox_change_audio_Playback_device", GUICtrlRead($c_checkbox_change_audio_Playback_device))
	IniWrite($s_INI_File_Location, "GUI", "c_checkbox_change_communication_device", GUICtrlRead($c_checkbox_change_communication_device))
	IniWrite($s_INI_File_Location, "GUI", "c_checkbox_select_custom_icon", GUICtrlRead($c_checkbox_select_custom_icon))
	IniWrite($s_INI_File_Location, "GUI", "c_checkbox_show_tray_notification_on_change", GUICtrlRead($c_checkbox_show_tray_notification_on_change))
	IniWrite($s_INI_File_Location, "GUI", "c_checkbox_show_tooltip_at_mousepos_on_change", GUICtrlRead($c_checkbox_show_tooltip_at_mousepos_on_change))
	IniWrite($s_INI_File_Location, "GUI", "c_checkbox_use_names_for_switch", GUICtrlRead($c_checkbox_use_names_for_switch))
	IniWrite($s_INI_File_Location, "GUI", "c_checkbox_select_shortcut_location", GUICtrlRead($c_checkbox_select_shortcut_location))
	IniWrite($s_INI_File_Location, "GUI", "c_checkbox_save_and_reapply_volume", GUICtrlRead($c_checkbox_save_and_reapply_volume))
	IniWrite($s_INI_File_Location, "GUI", "c_checkbox_volume_mute", GUICtrlRead($c_checkbox_volume_mute))
	IniWrite($s_INI_File_Location, "GUI", "c_slider_volume", Abs(100 - GUICtrlRead($c_slider_volume)))
	IniWrite($s_INI_File_Location, "GUI", "c_checkbox_Close_Control_Panel", GUICtrlRead($c_checkbox_Close_Control_Panel))
EndFunc   ;==>_GUI_IniWrite_All

Func WM_ACTIVATE($hWnd, $Msg, $wParam, $lParam)
	; ConsoleWrite(TimerInit() & @tab & $hWnd & @tab & $Msg & @crlf)
	Switch $hWnd
		Case $h_GUI
			; Window de-activated
			If Not $wParam Then
				ToolTip("")
			ElseIf $s_ToolTip_Title Then
				Local $a_WinPos = WinGetPos($h_GUI)
				ToolTip($s_ToolTip_Text, $a_WinPos[0], $a_WinPos[1], $s_ToolTip_Title, 1)
			EndIf
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_ACTIVATE

Func _func_create_shortcut_toogle($iDeviceNo_1, $sDeviceNo_1_Name, $iDeviceNo_1_Volume, $iDeviceNo_2, $sDeviceNo_2_Name, $iDeviceNo_2_Volume)

	If Not $g_Macro_Compiled Then
		_MsgBox_SHEx(16, $sGUITitle & " - Error", "Shortcuts are only created if SSD is compiled. #2", 0, $h_GUI)
	Else

		Local $iDeviceNo

		If GUICtrlRead($c_checkbox_use_names_for_switch) = $GUI_CHECKED Then
			$iDeviceNo_1 = _Device_Names_Get_Ref_From_INI($sDeviceNo_1_Name)
			If $iDeviceNo_1 = 0 Then
				_MsgBox_SHEx(16, $sGUITitle & " - Error", "Shortcut could not be created, Name reference for device #1 not found in SSD.ini.", 0, $h_GUI)
				Return
			EndIf
			$iDeviceNo_2 = _Device_Names_Get_Ref_From_INI($sDeviceNo_2_Name)
			If $iDeviceNo_2 = 0 Then
				_MsgBox_SHEx(16, $sGUITitle & " - Error", "Shortcut could not be created, Name reference for device #2 not found in SSD.ini.", 0, $h_GUI)
				Return
			EndIf
		EndIf

		If $iDeviceNo_1 > $iDeviceNo_2 Then
			Local $iDeviceNo_Buffer = $iDeviceNo_2
			$iDeviceNo_2 = $iDeviceNo_1
			$iDeviceNo_1 = $iDeviceNo_Buffer
			Local $sDeviceName_Buffer = $sDeviceNo_2_Name
			$sDeviceNo_2_Name = $sDeviceNo_1_Name
			$sDeviceNo_1_Name = $sDeviceName_Buffer
		EndIf

		Local $sDevice_Names = StringReplace($iDeviceNo_1 & "-" & $iDeviceNo_2, "777", "")

		$iDeviceNo = $iDeviceNo_1 & $iDeviceNo_1_Volume & "999" & $iDeviceNo_2 & $iDeviceNo_2_Volume

		; ConsoleWrite($iDeviceNo & @crlf)

		Local $iSuccess = 0
		Local $s_FileCreateShortcut_Filename = StringRegExpReplace(StringLeft($sDeviceNo_1_Name, StringInStr($sDeviceNo_1_Name, ", ") - 1) & " & " & StringLeft($sDeviceNo_2_Name, StringInStr($sDeviceNo_2_Name, ", ") - 1) & ".lnk", '[\\/:*?"<>|]', '') ; Remove invalid characters from filename

		Local $s_Add_Recording_Switch, $s_LinkFilename
		Local $s_Link_Device_Type = "Playback"
		If $b_Change_Recording_Device = True Then
			$s_Add_Recording_Switch = "_Recording"
			$s_Link_Device_Type = "Recording"
		EndIf

		Local $a_Shortcut_Icons = _Select_Custom_Icon()

		Local $s_LNK_Target_Location = @ScriptDir

		If GUICtrlRead($c_checkbox_create_shortcut_hidden) = $GUI_CHECKED Then

			If GUICtrlRead($c_checkbox_change_communication_device) = $GUI_CHECKED And GUICtrlRead($c_checkbox_change_audio_Playback_device) = $GUI_CHECKED Then
				$s_LinkFilename = "SSD - " & $s_Link_Device_Type & " - Toggle-#" & $sDevice_Names & " - Audio&Com - Hidden - " & $s_FileCreateShortcut_Filename
				If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
					$s_LNK_Target_Location = _FileSaveDialog_Ex("Save Shortcut to", @ScriptDir, "Shortcut (*.lnk)", 2, $s_LinkFilename, $h_GUI)
					If @error Then Return
				Else
					$s_LNK_Target_Location = $s_LNK_Target_Location & "\" & $s_LinkFilename
				EndIf
				If StringRight($s_LNK_Target_Location, 4) <> ".lnk" Then $s_LNK_Target_Location &= ".lnk"

				$iSuccess = FileCreateShortcut(@ScriptDir & "\" & @ScriptName, $s_LNK_Target_Location, @ScriptDir, $iDeviceNo & "hidden" & $s_Add_Recording_Switch, "SSD - " & $s_Link_Device_Type & " - Audio&Com - Hidden - " & StringTrimRight($s_FileCreateShortcut_Filename, 4), $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)

			ElseIf GUICtrlRead($c_checkbox_change_communication_device) = $GUI_CHECKED Then
				$s_LinkFilename = "SSD - " & $s_Link_Device_Type & " - Toggle-#" & $sDevice_Names & " - Com - Hidden - " & $s_FileCreateShortcut_Filename
				If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
					$s_LNK_Target_Location = _FileSaveDialog_Ex("Save Shortcut to", @ScriptDir, "Shortcut (*.lnk)", 2, $s_LinkFilename, $h_GUI)
					If @error Then Return
				Else
					$s_LNK_Target_Location = $s_LNK_Target_Location & "\" & $s_LinkFilename
				EndIf
				If StringRight($s_LNK_Target_Location, 4) <> ".lnk" Then $s_LNK_Target_Location &= ".lnk"

				$iSuccess = FileCreateShortcut(@ScriptDir & "\" & @ScriptName, $s_LNK_Target_Location, @ScriptDir, $iDeviceNo & "hidden_NoChange_Audio" & $s_Add_Recording_Switch, "SSD - " & $s_Link_Device_Type & " - Com - Hidden - " & StringTrimRight($s_FileCreateShortcut_Filename, 4), $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)

			Else
				$s_LinkFilename = "SSD - " & $s_Link_Device_Type & " - Toggle-#" & $sDevice_Names & " - Audio - Hidden - " & $s_FileCreateShortcut_Filename
				If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
					$s_LNK_Target_Location = _FileSaveDialog_Ex("Save Shortcut to", @ScriptDir, "Shortcut (*.lnk)", 2, $s_LinkFilename, $h_GUI)
					If @error Then Return
				Else
					$s_LNK_Target_Location = $s_LNK_Target_Location & "\" & $s_LinkFilename
				EndIf
				If StringRight($s_LNK_Target_Location, 4) <> ".lnk" Then $s_LNK_Target_Location &= ".lnk"

				$iSuccess = FileCreateShortcut(@ScriptDir & "\" & @ScriptName, $s_LNK_Target_Location, @ScriptDir, $iDeviceNo & "hidden_NoChange_Com" & $s_Add_Recording_Switch, "SSD - " & $s_Link_Device_Type & " - Audio - Hidden - " & StringTrimRight($s_FileCreateShortcut_Filename, 4), $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)

			EndIf

		Else

			If GUICtrlRead($c_checkbox_change_communication_device) = $GUI_CHECKED And GUICtrlRead($c_checkbox_change_audio_Playback_device) = $GUI_CHECKED Then
				$s_LinkFilename = "SSD - " & $s_Link_Device_Type & " - Toggle-#" & $sDevice_Names & " - Audio&Com - Visible - " & $s_FileCreateShortcut_Filename
				If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
					$s_LNK_Target_Location = _FileSaveDialog_Ex("Save Shortcut to", @ScriptDir, "Shortcut (*.lnk)", 2, $s_LinkFilename, $h_GUI)
					If @error Then Return
				Else
					$s_LNK_Target_Location = $s_LNK_Target_Location & "\" & $s_LinkFilename
				EndIf
				If StringRight($s_LNK_Target_Location, 4) <> ".lnk" Then $s_LNK_Target_Location &= ".lnk"

				$iSuccess = FileCreateShortcut(@ScriptDir & "\" & @ScriptName, $s_LNK_Target_Location, @ScriptDir, $iDeviceNo & $s_Add_Recording_Switch, "SSD - " & $s_Link_Device_Type & " - Audio&Com - Visible - " & StringTrimRight($s_FileCreateShortcut_Filename, 4), $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)

			ElseIf GUICtrlRead($c_checkbox_change_communication_device) = $GUI_CHECKED Then
				$s_LinkFilename = "SSD - " & $s_Link_Device_Type & " - Toggle-#" & $sDevice_Names & " - Com - Visible - " & $s_FileCreateShortcut_Filename
				If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
					$s_LNK_Target_Location = _FileSaveDialog_Ex("Save Shortcut to", @ScriptDir, "Shortcut (*.lnk)", 2, $s_LinkFilename, $h_GUI)
					If @error Then Return
				Else
					$s_LNK_Target_Location = $s_LNK_Target_Location & "\" & $s_LinkFilename
				EndIf
				If StringRight($s_LNK_Target_Location, 4) <> ".lnk" Then $s_LNK_Target_Location &= ".lnk"

				$iSuccess = FileCreateShortcut(@ScriptDir & "\" & @ScriptName, $s_LNK_Target_Location, @ScriptDir, $iDeviceNo & "_NoChange_Audio" & $s_Add_Recording_Switch, "SSD - " & $s_Link_Device_Type & " - Com - Visible - " & StringTrimRight($s_FileCreateShortcut_Filename, 4), $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)

			Else
				$s_LinkFilename = "SSD - " & $s_Link_Device_Type & " - Toggle-#" & $sDevice_Names & " - Audio - Visible - " & $s_FileCreateShortcut_Filename
				If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
					$s_LNK_Target_Location = _FileSaveDialog_Ex("Save Shortcut to", @ScriptDir, "Shortcut (*.lnk)", 2, $s_LinkFilename, $h_GUI)
					If @error Then Return
				Else
					$s_LNK_Target_Location = $s_LNK_Target_Location & "\" & $s_LinkFilename
				EndIf
				If StringRight($s_LNK_Target_Location, 4) <> ".lnk" Then $s_LNK_Target_Location &= ".lnk"

				$iSuccess = FileCreateShortcut(@ScriptDir & "\" & @ScriptName, $s_LNK_Target_Location, @ScriptDir, $iDeviceNo & "_NoChange_Com" & $s_Add_Recording_Switch, "SSD - " & $s_Link_Device_Type & " - Audio - Visible - " & StringTrimRight($s_FileCreateShortcut_Filename, 4), $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)

			EndIf

		EndIf

		If $iSuccess = 1 Then

			If _MsgBox_SHEx(4 + 64 + 256, $sGUITitle, "Shortcut has been created successfully:" & @CRLF & @CRLF _
					 & StringReplace($s_LinkFilename, $s_LNK_Target_Location & "\", "", 0, 2) & @CRLF & @CRLF _
					 & "Pressing the shortcut SSD will toggle betweeen these two " & $s_Link_Device_Type & " Sound Devices:" & @CRLF & @CRLF _
					 & "1) " & $sDeviceNo_1_Name & @CRLF _
					 & "2) " & $sDeviceNo_2_Name & @CRLF & @CRLF & "Do you want SSD to open the target location?", 0, $h_GUI) = 6 Then

				If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
					ShellExecute(StringTrimRight($s_LNK_Target_Location, StringLen($s_LNK_Target_Location) - StringInStr($s_LNK_Target_Location, "\", 2, -1) + 1), "", @ScriptDir, "open", @SW_MAXIMIZE)
				Else
					ShellExecute($s_LNK_Target_Location, "", @ScriptDir, "open", @SW_MAXIMIZE)
				EndIf

			EndIf
		Else
			_MsgBox_SHEx(16, $sGUITitle & " - Error", "Shortcut could not be created.", 0, $h_GUI)
		EndIf

	EndIf

EndFunc   ;==>_func_create_shortcut_toogle

Func _Select_Custom_Icon()
	Local $a_Shortcut_Icons[2]
	If GUICtrlRead($c_checkbox_select_custom_icon) = $GUI_CHECKED Then
		If $b_Change_Recording_Device = True Then
			$a_Shortcut_Icons = _PickIconDlg("%windir%\system32\mmres.dll", 12, WinGetHandle(""))
			; If Not @error Then _MsgBox_SHEx(64, "Results", StringFormat("IconFile Selected: %s\nIconID Selected: %i", $aRet[0], $aRet[1]))
		Else
			$a_Shortcut_Icons = _PickIconDlg("%windir%\system32\mmres.dll", 0, WinGetHandle(""))
		EndIf
	EndIf
	If @error Then
		Local $a_Shortcut_Icons[2]
	EndIf
	If Not $a_Shortcut_Icons[0] Then
		Local $a_Shortcut_Icons[2]
		$a_Shortcut_Icons[0] = "%windir%\system32\mmres.dll"
		If $b_Change_Recording_Device = True Then
			$a_Shortcut_Icons[1] = 12
		Else
			$a_Shortcut_Icons[1] = 0
		EndIf
	EndIf
	Return $a_Shortcut_Icons
EndFunc   ;==>_Select_Custom_Icon

Func _PickIconDlg($sFileName, $nIconIndex = 0, $hWnd = 0)
	; by MrCreatoR
	; https://www.autoitscript.com/forum/topic/80978-icon-selector/?do=findComment&comment=581943
	Local $nRet, $aRetArr[2]

	$nRet = DllCall("shell32.dll", "int", "PickIconDlg", _
			"hwnd", $hWnd, _
			"wstr", $sFileName, "int", 1000, "int*", $nIconIndex)

	If Not $nRet[0] Then Return SetError(1, 0, -1)

	$aRetArr[0] = $nRet[2]
	$aRetArr[1] = $nRet[4]

	Return $aRetArr
EndFunc   ;==>_PickIconDlg


Func _func_create_shortcut_standard($iDeviceNo, $s_SoundDevice_Name, $i_Volume_Info = "")

	; _MsgBox_SHEx(0, "", $iDeviceNo & @CRLF & $s_SoundDevice_Name & @CRLF & $i_Volume_Info)

	If GUICtrlRead($c_checkbox_use_names_for_switch) = $GUI_CHECKED Then
		$iDeviceNo = _Device_Names_Get_Ref_From_INI($s_SoundDevice_Name)
		If $iDeviceNo = 0 Then
			_MsgBox_SHEx(16, $sGUITitle & " - Error", "Shortcut could not be created, Name reference not found in SSD.ini.", 0, $h_GUI)
			Return
		EndIf
	EndIf

	Local $iSuccess = 0
	Local $s_FileCreateShortcut_Filename = StringRegExpReplace(StringLeft($s_SoundDevice_Name, StringInStr($s_SoundDevice_Name, ", ") - 1) & ".lnk", '[\\/:*?"<>|]', '') ; Remove invalid characters from filename

	Local $s_Add_Recording_Switch, $s_LinkFilename
	Local $s_Link_Device_Type = "Playback"
	If $b_Change_Recording_Device = True Then
		$s_Add_Recording_Switch = "_Recording"
		$s_Link_Device_Type = "Recording"
	EndIf

	Local $a_Shortcut_Icons = _Select_Custom_Icon()

	Local $s_LNK_Target_Location = @ScriptDir

	If GUICtrlRead($c_checkbox_create_shortcut_hidden) = $GUI_CHECKED Then

		If GUICtrlRead($c_checkbox_change_communication_device) = $GUI_CHECKED And GUICtrlRead($c_checkbox_change_audio_Playback_device) = $GUI_CHECKED Then
			$s_LinkFilename = "SSD - " & $s_Link_Device_Type & " - #" & StringReplace($iDeviceNo, "777", "") & " - Audio&Com - Hidden - " & $s_FileCreateShortcut_Filename
			If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
				$s_LNK_Target_Location = _FileSaveDialog_Ex("Save Shortcut to", @ScriptDir, "Shortcut (*.lnk)", 2, $s_LinkFilename, $h_GUI)
				If @error Then Return
			Else
				$s_LNK_Target_Location = $s_LNK_Target_Location & "\" & $s_LinkFilename
			EndIf
			If StringRight($s_LNK_Target_Location, 4) <> ".lnk" Then $s_LNK_Target_Location &= ".lnk"

			$iSuccess = FileCreateShortcut(@ScriptDir & "\" & @ScriptName, $s_LNK_Target_Location, @ScriptDir, $iDeviceNo & $i_Volume_Info & "hidden" & $s_Add_Recording_Switch, "SSD - " & $s_Link_Device_Type & " - Audio&Com - Hidden - " & StringTrimRight($s_FileCreateShortcut_Filename, 4), $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)

		ElseIf GUICtrlRead($c_checkbox_change_communication_device) = $GUI_CHECKED Then
			$s_LinkFilename = "SSD - " & $s_Link_Device_Type & " - #" & StringReplace($iDeviceNo, "777", "") & " - Com - Hidden - " & $s_FileCreateShortcut_Filename
			If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
				$s_LNK_Target_Location = _FileSaveDialog_Ex("Save Shortcut to", @ScriptDir, "Shortcut (*.lnk)", 2, $s_LinkFilename, $h_GUI)
				If @error Then Return
			Else
				$s_LNK_Target_Location = $s_LNK_Target_Location & "\" & $s_LinkFilename
			EndIf
			If StringRight($s_LNK_Target_Location, 4) <> ".lnk" Then $s_LNK_Target_Location &= ".lnk"

			$iSuccess = FileCreateShortcut(@ScriptDir & "\" & @ScriptName, $s_LNK_Target_Location, @ScriptDir, $iDeviceNo & $i_Volume_Info & "hidden_NoChange_Audio" & $s_Add_Recording_Switch, "SSD - " & $s_Link_Device_Type & " - Com - Hidden - " & StringTrimRight($s_FileCreateShortcut_Filename, 4), $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)

		Else
			$s_LinkFilename = "SSD - " & $s_Link_Device_Type & " - #" & StringReplace($iDeviceNo, "777", "") & " - Audio - Hidden - " & $s_FileCreateShortcut_Filename
			If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
				$s_LNK_Target_Location = _FileSaveDialog_Ex("Save Shortcut to", @ScriptDir, "Shortcut (*.lnk)", 2, $s_LinkFilename, $h_GUI)
				If @error Then Return
			Else
				$s_LNK_Target_Location = $s_LNK_Target_Location & "\" & $s_LinkFilename
			EndIf
			If StringRight($s_LNK_Target_Location, 4) <> ".lnk" Then $s_LNK_Target_Location &= ".lnk"

			$iSuccess = FileCreateShortcut(@ScriptDir & "\" & @ScriptName, $s_LNK_Target_Location, @ScriptDir, $iDeviceNo & $i_Volume_Info & "hidden_NoChange_Com" & $s_Add_Recording_Switch, "SSD - " & $s_Link_Device_Type & " - Audio - Hidden - " & StringTrimRight($s_FileCreateShortcut_Filename, 4), $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)

		EndIf

	Else

		If GUICtrlRead($c_checkbox_change_communication_device) = $GUI_CHECKED And GUICtrlRead($c_checkbox_change_audio_Playback_device) = $GUI_CHECKED Then
			$s_LinkFilename = "SSD - " & $s_Link_Device_Type & " - #" & StringReplace($iDeviceNo, "777", "") & " - Audio&Com - Visible - " & $s_FileCreateShortcut_Filename
			If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
				$s_LNK_Target_Location = _FileSaveDialog_Ex("Save Shortcut to", @ScriptDir, "Shortcut (*.lnk)", 2, $s_LinkFilename, $h_GUI)
				If @error Then Return
			Else
				$s_LNK_Target_Location = $s_LNK_Target_Location & "\" & $s_LinkFilename
			EndIf
			If StringRight($s_LNK_Target_Location, 4) <> ".lnk" Then $s_LNK_Target_Location &= ".lnk"

			$iSuccess = FileCreateShortcut(@ScriptDir & "\" & @ScriptName, $s_LNK_Target_Location, @ScriptDir, $iDeviceNo & $i_Volume_Info & $s_Add_Recording_Switch, "SSD - " & $s_Link_Device_Type & " - Audio&Com - Visible - " & StringTrimRight($s_FileCreateShortcut_Filename, 4), $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)

		ElseIf GUICtrlRead($c_checkbox_change_communication_device) = $GUI_CHECKED Then
			$s_LinkFilename = "SSD - " & $s_Link_Device_Type & " - #" & StringReplace($iDeviceNo, "777", "") & " - Com - Visible - " & $s_FileCreateShortcut_Filename
			If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
				$s_LNK_Target_Location = _FileSaveDialog_Ex("Save Shortcut to", @ScriptDir, "Shortcut (*.lnk)", 2, $s_LinkFilename, $h_GUI)
				If @error Then Return
			Else
				$s_LNK_Target_Location = $s_LNK_Target_Location & "\" & $s_LinkFilename
			EndIf
			If StringRight($s_LNK_Target_Location, 4) <> ".lnk" Then $s_LNK_Target_Location &= ".lnk"

			$iSuccess = FileCreateShortcut(@ScriptDir & "\" & @ScriptName, $s_LNK_Target_Location, @ScriptDir, $iDeviceNo & $i_Volume_Info & "_NoChange_Audio" & $s_Add_Recording_Switch, "SSD - " & $s_Link_Device_Type & " - Com - Visible - " & StringTrimRight($s_FileCreateShortcut_Filename, 4), $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)

		Else
			$s_LinkFilename = "SSD - " & $s_Link_Device_Type & " - #" & StringReplace($iDeviceNo, "777", "") & " - Audio - Visible - " & $s_FileCreateShortcut_Filename
			If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
				$s_LNK_Target_Location = _FileSaveDialog_Ex("Save Shortcut to", @ScriptDir, "Shortcut (*.lnk)", 2, $s_LinkFilename, $h_GUI)
				If @error Then Return
			Else
				$s_LNK_Target_Location = $s_LNK_Target_Location & "\" & $s_LinkFilename
			EndIf
			If StringRight($s_LNK_Target_Location, 4) <> ".lnk" Then $s_LNK_Target_Location &= ".lnk"

			$iSuccess = FileCreateShortcut(@ScriptDir & "\" & @ScriptName, $s_LNK_Target_Location, @ScriptDir, $iDeviceNo & $i_Volume_Info & "_NoChange_Com" & $s_Add_Recording_Switch, "SSD - " & $s_Link_Device_Type & " - Audio - Visible - " & StringTrimRight($s_FileCreateShortcut_Filename, 4), $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)

		EndIf

	EndIf

	If $iSuccess = 1 Then

		If _MsgBox_SHEx(4 + 64 + 256, $sGUITitle, "Shortcut has been created successfully:" & @CRLF & @CRLF & $s_LNK_Target_Location & @CRLF & @CRLF & "Do you want SSD to open the target location?", 0, $h_GUI) = 6 Then
			If GUICtrlRead($c_checkbox_select_shortcut_location) = $GUI_CHECKED Then
				ShellExecute(StringTrimRight($s_LNK_Target_Location, StringLen($s_LNK_Target_Location) - StringInStr($s_LNK_Target_Location, "\", 2, -1) + 1), "", @ScriptDir, "open", @SW_MAXIMIZE)
			Else
				ShellExecute($s_LNK_Target_Location, "", @ScriptDir, "open", @SW_MAXIMIZE)
			EndIf
		EndIf

	Else
		_MsgBox_SHEx(16, $sGUITitle & " - Error", "Shortcut could not be created.", 0, $h_GUI)
	EndIf

EndFunc   ;==>_func_create_shortcut_standard


Func _Detect_Sound_dialog_hWnd($iPID)
	Local $hWnds = WinList("[CLASS:#32770;]", "")
	For $i = 1 To $hWnds[0][0]
		If BitAND(WinGetState($hWnds[$i][1], ''), 2) Then
			If WinGetProcess(_WinAPI_GetAncestor($hWnds[$i][1], $GA_ROOTOWNER)) = ProcessExists($iPID) Then
				Return $hWnds[$i][1]
			EndIf
		EndIf
	Next
	Return 0
EndFunc   ;==>_Detect_Sound_dialog_hWnd

Func _Merge_Two_Shortcuts()

	Local $s_Data_to_write

	Local $s_File_A = FileOpenDialog("Select Shortcut #1", @ScriptDir, "Shortcut (*.lnk)", 3, "", $h_GUI)
	If @error Then Return

	Local $a_Data = FileGetShortcut($s_File_A)
	If @error Then
		_MsgBox_SHEx(16, $sGUITitle & " - Error", "Shortcut info could not be extracted for:" & @CRLF & $s_File_A, 0, $h_GUI)
		Return
	EndIf

	$s_Data_to_write = ':' & $a_Data[3] & @CRLF & '"' & $a_Data[0] & '" ' & $a_Data[2] & @CRLF & @CRLF

	Local $s_File_B = FileOpenDialog("Select Shortcut #2", @ScriptDir, "Shortcut (*.lnk)", 3, "", $h_GUI)
	If @error Then Return

	$a_Data = FileGetShortcut($s_File_B)
	If @error Then
		_MsgBox_SHEx(16, $sGUITitle & " - Error", "Shortcut info could not be extracted for:" & @CRLF & $s_File_B, 0, $h_GUI)
		Return
	EndIf

	$s_Data_to_write &= ':' & $a_Data[3] & @CRLF & '"' & $a_Data[0] & '" ' & $a_Data[2]

	Local $s_File_New = FileSaveDialog("Enter name for merged Shortcut to be created", StringTrimRight($a_Data[0], StringLen($a_Data[0]) - StringInStr($a_Data[0], "\", 2, -1) + 1), "Shortcut (*.lnk)", 2 + 16, "SSD - Merged Shortcut - ", $h_GUI)
	If @error Then Return

	If StringRight($s_File_New, 4) = ".lnk" Then $s_File_New = StringTrimRight($s_File_New, 4)

	Local $iSuccess
	Local $h_file = FileOpen($s_File_New & ".bat", 2 + 128)
	If $h_file <> -1 Then $iSuccess = FileWrite($h_file, $s_Data_to_write)
	FileClose($h_file)

	If $iSuccess = 1 Then
		Local $a_Shortcut_Icons = _Select_Custom_Icon()
		$iSuccess = FileCreateShortcut($s_File_New & ".bat", $s_File_New & ".lnk", @ScriptDir, "", $s_File_New, $a_Shortcut_Icons[0], "", $a_Shortcut_Icons[1], @SW_SHOWMINNOACTIVE)
	Else
		FileDelete($s_File_New & ".bat")
	EndIf

	If $iSuccess = 1 Then

		If _MsgBox_SHEx(4 + 64, $sGUITitle, "Shortcut has been created successfully:" & @CRLF & @CRLF & $s_File_New & @CRLF & @CRLF & "Do you want SSD to open the target location?", 0, $h_GUI) = 6 Then
			ShellExecute(StringTrimRight($s_File_New & ".lnk", StringLen($s_File_New & ".lnk") - StringInStr($s_File_New & ".lnk", "\", 2, -1) + 1), "", @ScriptDir, "open", @SW_MAXIMIZE)
		EndIf

	Else
		FileDelete($s_File_New & ".bat")
		_MsgBox_SHEx(16, $sGUITitle & " - Error", "Shortcut could not be created.", 0, $h_GUI)
	EndIf

EndFunc   ;==>_Merge_Two_Shortcuts

Func _Exit($iExitCode = 0, $i_checkbox_Close_Control_Panel = 0)

	If $i_checkbox_Close_Control_Panel <> $GUI_UNCHECKED Then

		$hWnd_Sound = _Detect_Sound_dialog_hWnd($iPID_Sound)

		Local $timer_timeout = TimerInit()
		While ProcessExists($iPID_Sound)
			WinClose($hWnd_Sound)
			If TimerDiff($timer_timeout) > 2000 Then ExitLoop
			If Not WinExists($hWnd_Sound) Then ExitLoop
			Sleep(10)
		WEnd

		$timer_timeout = TimerInit()
		While ProcessExists($iPID_Sound) And WinExists($hWnd_Sound)
			ProcessClose($iPID_Sound)
			If TimerDiff($timer_timeout) > 2000 Then
				; _Exit_Error()
				If $b_SSD_Is_Running_on_Hidden_Desktop = False Then _MsgBox_SHEx(16, $sGUITitle & " - Error", "The process rundll32.exe with the PID " & $iPID_Sound & "could not be closed by SSD." & @CRLF & @CRLF & "SSD will exit now, close the process via Taskmanager.", 0, $hwnd_AutoIt)
				Exit 3
			EndIf
			Sleep(10)
		WEnd

	EndIf

	If $b_SSD_Is_Running_on_Hidden_Desktop = True Then
		Exit Int($iExitCode & "1")
	Else
		If $iExitCode = 2 Then
			AutoItWinSetTitle("")
			Run(@ScriptFullPath, @ScriptDir)
		EndIf
		Exit Int($iExitCode & "0")
	EndIf

EndFunc   ;==>_Exit

Func _OnAutoitExit_CloseDlls()
	DllClose($h_DLL_Kernel32)
	DllClose($h_DLL_User32)
	DllClose($h_DLL_advapi32)
	DllClose($h_DLL_gdi32)
	DllClose($h_DLL_shell32)
	DllClose($h_DLL_ntdll)
	DllClose($h_DLL_ole32)
	DllClose($h_DLL_oleaut32)
	DllClose($h_DLL_Crypt32)
	TraySetState(2) ; Hide
EndFunc   ;==>_OnAutoitExit_CloseDlls

Func MY_WM_COMMAND($hWnd, $Msg, $wParam, $lParam)
	Local $iIDFrom = BitAND($wParam, 0xFFFF) ;LoWord
	Local $iCode = BitShift($wParam, 16) ;HiWord
	Switch $iIDFrom
		Case $c_label_Hyperlink_FunkEu, $c_label_Hyperlink_FunkEu2, $c_icon_Hyperlink_FunkEu, $c_Picture_Title
			GUICtrlSetColor($c_label_Hyperlink_FunkEu, 0xFF0000)
			While _IsPressed("01", $h_DLL_User32)
				Sleep(10)
			WEnd
			GUICtrlSetColor($c_label_Hyperlink_FunkEu, 0x518CB8)
			ShellExecute('http://www.funk.eu/')
	EndSwitch
EndFunc   ;==>MY_WM_COMMAND

Func _EnforceSingleInstance($GUID_Program = "")
	Local $hWnd = WinGetHandle($GUID_Program), $hwnd_Target, $timer
	If IsHWnd($hWnd) Then
		$timer = TimerInit()
		While IsHWnd($hWnd)
			$hWnd = WinGetHandle($GUID_Program)
			$hwnd_Target = HWnd(ControlGetText($hWnd, '', ControlGetHandle($hWnd, '', 'Edit1')))
			WM_COPYDATA_SendData($hwnd_Target, "Exit")
			Sleep(10)
			If TimerDiff($timer) > 3000 Then ExitLoop
		WEnd
		$hWnd = WinGetHandle($GUID_Program)
		If IsHWnd($hWnd) Then ProcessClose(WinGetProcess($hWnd))
	EndIf
	AutoItWinSetTitle($GUID_Program)
EndFunc   ;==>_EnforceSingleInstance

Func WM_COPYDATA($hWnd, $MsgID, $wParam, $lParam)
	; http://www.autoitscript.com/forum/index.php?showtopic=105861&view=findpost&p=747887
	; Melba23, based on code from Yashied
	Local $tCOPYDATA = DllStructCreate("ulong_ptr;dword;ptr", $lParam)
	Local $tMsg = DllStructCreate("char[" & DllStructGetData($tCOPYDATA, 2) & "]", DllStructGetData($tCOPYDATA, 3))

	If DllStructGetData($tMsg, 1) = "Exit" Then
		TraySetState(2) ; hide tray icon
		Exit 20
	EndIf

	Return 0
EndFunc   ;==>WM_COPYDATA

Func WM_COPYDATA_SendData($hWnd, $sData)
	If Not IsHWnd($hWnd) Then Return 0
	; If $sData = "" Then $sData = " "
	; Local $tCOPYDATA, $tMsg
	Local $tMsg = DllStructCreate("char[" & StringLen($sData) + 1 & "]")
	DllStructSetData($tMsg, 1, $sData)
	Local $tCOPYDATA = DllStructCreate("ulong_ptr;dword;ptr")
	DllStructSetData($tCOPYDATA, 2, StringLen($sData) + 1)
	DllStructSetData($tCOPYDATA, 3, DllStructGetPtr($tMsg))
	Local $Ret = DllCall("user32.dll", "lparam", "SendMessage", "hwnd", $hWnd, "int", $WM_COPYDATA, "wparam", 0, "lparam", DllStructGetPtr($tCOPYDATA))
	If (@error) Or ($Ret[0] = -1) Then Return 0
	Return 1
EndFunc   ;==>WM_COPYDATA_SendData

Func __WINVER_RtlGetVersion()
	; GetVersionEx
	; https://msdn.microsoft.com/de-de/library/windows/desktop/ms724451(v=vs.85).aspx
	; With the release of Windows 8.1, the behavior of the GetVersionEx API has changed in the value it will return for the operating system version.
	; The value returned by the GetVersionEx function now depends on how the application is manifested.

	; If you don't want to depend on manifests and reply on this deprecated API, use kernel-mode RtlGetVersion:
	; https://msdn.microsoft.com/en-us/library/windows/hardware/ff561910(v=vs.85).aspx

	#cs
		typedef struct _OSVERSIONINFOW {
		ULONG dwOSVersionInfoSize;
		ULONG dwMajorVersion;
		ULONG dwMinorVersion;
		ULONG dwBuildNumber;
		ULONG dwPlatformId;
		WCHAR szCSDVersion[128];
		} RTL_OSVERSIONINFOW, *PRTL_OSVERSIONINFOW;
	#ce

	Local $tOSVI = DllStructCreate('dword;dword;dword;dword;dword;wchar[128]')
	DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
	Local $Ret = DllCall("ntdll.dll", "int", "RtlGetVersion", "ptr", DllStructGetPtr($tOSVI))
	If @error Or $Ret[0] <> 0 Then Return SetError(1, 0, 0) ; RtlGetVersion returns STATUS_SUCCESS = 0

	; 0x0501 = Win XP
	; 0x0502 = Win Server 2003
	; 0x0600 = Win Vista
	; 0x0601 = Win7 / Major Version = 6, Minor Version = 1
	; 0x0602 = Win8
	; 0x0603 = Win8.1
	; 0x0604 = Win10 "Technical Preview"
	; 0x0A00 = Win10 RTM (build 10240 or later) / Major Version = 10, Minor Version = 0

	; Return "0x" & Hex(BitOR(BitShift(10, -8), 0), 4)
	Return "0x" & Hex(BitOR(BitShift(DllStructGetData($tOSVI, 2), -8), DllStructGetData($tOSVI, 3)), 4)
EndFunc   ;==>__WINVER_RtlGetVersion









; Based on File to Base64 String Code Generator
; by UEZ
; http://www.autoitscript.com/forum/topic/134350-file-to-base64-string-code-generator-v103-build-2011-11-21/

;======================================================================================
; Function Name:        Load_BMP_From_Mem
; Description:          Loads an image which is saved as a binary string and converts it to a bitmap or hbitmap
;
; Parameters:           $bImage:    the binary string which contains any valid image which is supported by GDI+
; Optional:             $hHBITMAP:  if false a bitmap will be created, if true a hbitmap will be created
;
; Remark:               hbitmap format is used generally for GUI internal images, $bitmap is more a GDI+ image format
;                       Don't forget _GDIPlus_Startup() and _GDIPlus_Shutdown()
;
; Requirement(s):       GDIPlus.au3, Memory.au3 and _GDIPlus_BitmapCreateDIBFromBitmap() from WinAPIEx.au3
; Return Value(s):      Success: handle to bitmap (GDI+ bitmap format) or hbitmap (WinAPI bitmap format),
;                       Error: 0
; Error codes:          1: $bImage is not a binary string
;                       2: unable to create stream on HGlobal
;                       3: unable to create bitmap from stream
;
; Author(s):            UEZ
; Additional Code:      thanks to progandy for the MemGlobalAlloc and tVARIANT lines and
;                       Yashied for _GDIPlus_BitmapCreateDIBFromBitmap() from WinAPIEx.au3
; Version:              v0.97 Build 2012-01-04 Beta
;=======================================================================================
Func _Load_BMP_From_Mem($bImage, $hHBITMAP = False)
	If Not IsBinary($bImage) Then Return SetError(1, 0, 0)
	Local $aResult
	Local Const $memBitmap = Binary($bImage) ;load image  saved in variable (memory) and convert it to binary
	Local Const $len = BinaryLen($memBitmap) ;get length of image
	Local Const $hData = _MemGlobalAlloc($len, $GMEM_MOVEABLE) ;allocates movable memory  ($GMEM_MOVEABLE = 0x0002)
	Local Const $pData = _MemGlobalLock($hData) ;translate the handle into a pointer
	Local $tMem = DllStructCreate("byte[" & $len & "]", $pData) ;create struct
	DllStructSetData($tMem, 1, $memBitmap) ;fill struct with image data
	_MemGlobalUnlock($hData) ;decrements the lock count  associated with a memory object that was allocated with GMEM_MOVEABLE
	$aResult = DllCall("ole32.dll", "int", "CreateStreamOnHGlobal", "handle", $pData, "int", True, "ptr*", 0) ;Creates a stream object that uses an HGLOBAL memory handle to store the stream contents
	If @error Then SetError(2, 0, 0)
	Local Const $hStream = $aResult[3]
	$aResult = DllCall($__g_hGDIPDll, "uint", "GdipCreateBitmapFromStream", "ptr", $hStream, "int*", 0) ;Creates a Bitmap object based on an IStream COM interface
	If @error Then SetError(3, 0, 0)
	Local Const $hBitmap = $aResult[2]
	Local $tVARIANT = DllStructCreate("word vt;word r1;word r2;word r3;ptr data; ptr")
	DllCall("oleaut32.dll", "long", "DispCallFunc", "ptr", $hStream, "dword", 8 + 8 * @AutoItX64, _
			"dword", 4, "dword", 23, "dword", 0, "ptr", 0, "ptr", 0, "ptr", DllStructGetPtr($tVARIANT)) ;release memory from $hStream to avoid memory leak
	$tMem = 0
	$tVARIANT = 0
	If $hHBITMAP Then
		Local Const $hHBmp = _GDIPlus_BitmapCreateDIBFromBitmap($hBitmap)
		_GDIPlus_BitmapDispose($hBitmap)
		Return $hHBmp
	EndIf
	Return $hBitmap
EndFunc   ;==>_Load_BMP_From_Mem

#cs
	Func _GDIPlus_BitmapCreateDIBFromBitmap($hBitmap)
	Local $tBIHDR, $Ret, $tData, $pBits, $hResult = 0
	$Ret = DllCall($__g_hGDIPDll, 'uint', 'GdipGetImageDimension', 'ptr', $hBitmap, 'float*', 0, 'float*', 0)
	If (@error) Or ($Ret[0]) Then Return 0
	$tData = _GDIPlus_BitmapLockBits($hBitmap, 0, 0, $Ret[2], $Ret[3], $GDIP_ILMREAD, $GDIP_PXF32ARGB)
	$pBits = DllStructGetData($tData, 'Scan0')
	If Not $pBits Then Return 0
	$tBIHDR = DllStructCreate('dword;long;long;ushort;ushort;dword;dword;long;long;dword;dword')
	DllStructSetData($tBIHDR, 1, DllStructGetSize($tBIHDR))
	DllStructSetData($tBIHDR, 2, $Ret[2])
	DllStructSetData($tBIHDR, 3, $Ret[3])
	DllStructSetData($tBIHDR, 4, 1)
	DllStructSetData($tBIHDR, 5, 32)
	DllStructSetData($tBIHDR, 6, 0)
	$hResult = DllCall('gdi32.dll', 'ptr', 'CreateDIBSection', 'hwnd', 0, 'ptr', DllStructGetPtr($tBIHDR), 'uint', 0, 'ptr*', 0, 'ptr', 0, 'dword', 0)
	If (Not @error) And ($hResult[0]) Then
	DllCall('gdi32.dll', 'dword', 'SetBitmapBits', 'ptr', $hResult[0], 'dword', $Ret[2] * $Ret[3] * 4, 'ptr', DllStructGetData($tData, 'Scan0'))
	$hResult = $hResult[0]
	Else
	$hResult = 0
	EndIf
	_GDIPlus_BitmapUnlockBits($hBitmap, $tData)
	Return $hResult
	EndFunc   ;==>_GDIPlus_BitmapCreateDIBFromBitmap
#ce

;Code was generated by File to Base64 String Code Generator

Func _Base64Decode($input_string)
	Local $struct = DllStructCreate("int")
	Local $a_Call = DllCall("Crypt32.dll", "int", "CryptStringToBinary", "str", $input_string, "int", 0, "int", 1, "ptr", 0, "ptr", DllStructGetPtr($struct, 1), "ptr", 0, "ptr", 0)
	If @error Or Not $a_Call[0] Then Return SetError(1, 0, "")
	Local $a = DllStructCreate("byte[" & DllStructGetData($struct, 1) & "]")
	$a_Call = DllCall("Crypt32.dll", "int", "CryptStringToBinary", "str", $input_string, "int", 0, "int", 1, "ptr", DllStructGetPtr($a), "ptr", DllStructGetPtr($struct, 1), "ptr", 0, "ptr", 0)
	If @error Or Not $a_Call[0] Then Return SetError(2, 0, "")
	Return DllStructGetData($a, 1)
EndFunc   ;==>_Base64Decode

Func _Decompress_Binary_String_to_Bitmap($Base64String)
	$Base64String = Binary($Base64String)
	Local $iSize_Source = BinaryLen($Base64String)
	Local $pBuffer_Source = _WinAPI_CreateBuffer($iSize_Source)
	DllStructSetData(DllStructCreate('byte[' & $iSize_Source & ']', $pBuffer_Source), 1, $Base64String)
	Local $pBuffer_Decompress = _WinAPI_CreateBuffer(8388608)
	Local $Size_Decompressed = _WinAPI_DecompressBuffer($pBuffer_Decompress, 8388608, $pBuffer_Source, $iSize_Source)
	Local $b_Result = Binary(DllStructGetData(DllStructCreate('byte[' & $Size_Decompressed & ']', $pBuffer_Decompress), 1))
	_WinAPI_FreeMemory($pBuffer_Source)
	_WinAPI_FreeMemory($pBuffer_Decompress)
	Return $b_Result
EndFunc   ;==>_Decompress_Binary_String_to_Bitmap

Func _BinaryString_Picture_License()
	Local $Base64String
	$Base64String &= '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAAPAFADAREAAhEBAxEB/8QAGQABAAMBAQAAAAAAAAAAAAAACAYHCQoA/8QAJRAAAgIDAAICAgIDAAAAAAAABgcFCAMECQECFBUWFwAKEyQo/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAP/xAAjEQACAgICAQQDAAAAAAAAAAAAAQIRITESQQMTUYHwYZHR/9oADAMBAAIRAxEAPwCRb62rspOddPZsFpRQsndDVqhzjBQOXZNKqvGm+SvuzyXr+M6hmfFRKnCEnJ9jOfsPdPS+RnZPckZz48njzb3na3Mfn+Qcp+rwUqVrpa42+rzkg5Tfl4KVLHS1xt7TNNBjT45qcT2QQi5Sg1qJZcuzFSrQslpc7ecYxD2uuZBsMVVJepgCKhMipHQwt0zclkPH2jCB0sms8EEHO6MsQpxB8tJbNy4VmHW9Ai9wydCZ6WVLXYBZyseWzqAwNHm3z+inhVHb2Sr3TjFVJVARNfZkNMfRVmUwIni5liOQO/M3pyG7AExMyofUxTu5LySlDjJPF1JUv5er+7l5ZShxkni6axnvdWryvueX2vkLfyy1p7TVPAR/mxHm1S41kzTAkd7jFThm55uNWb2BK+y2EJXNWuZNjnGVzGwYsSBlfbTgV3JRsOI6ZGTz01GxEDtbXvUrvRSWdu3F1AQuZ21n5AeV+tf2KPtMsh+e/JkxhwBuhUwRQwcismyC0uJ/2KwLCfTR5BW8lSnqz0m2QTfN2PFNvVX9Zrkk1cALADYrrYUmLZWM7TuoCfcCqr/G2QwJl+8cObySabWBJt7LSuo7DpUQYtF4HdYjAMmaz4wfXARH+2vMNYjhZlcK3GZODdEVyUgeO5i7qyW3syz/AHuQAhpTq/Az9SQBFzV5qQ0w9fySp1e7rHoUupWSoBqgsMwEohbRpaSLhxumSu/aR2WYFPVTdsUzfXGKZgI+qpfqEzzRTB0pWCkCN1n2n3c8kgxbM8ieZtfU01l6hUER2PL58OcDMpBArvcH9xdwOjljjvcIdFZxOcvD5g6Nw0Jldgu0AEACqXo+XhYqQz8dygUhswyBZrNaoxt8m6QDDlNrDuJ+3grSuayZITQ5kzgstnBONagLth5LcdhisFIE4JQCysBpCuzJE+mIAD+pVtT6yh82lE3VLSCWCZakHSIsyYxPm7z0VZVGlSr56Wfa67JhliKisASxBAgEGIEipVDTIqVQ0jryMNrePOz763vsa+YDfzN0S5eNag9bExMdCANMORc1QpBBaOxL1/uIRy6jsXWFOJbDHb2xnF6wlYnO64m3Vl6+m5sDRFKxZDA4drHGy2TVkvXN/IuE/U5xcesNv2p9Mi4S9TmnHrDb9qfT/JokMdPf6w7ChJI+sY5RqBsA0pjC0W9rJTD1k2VoGP8AkCkKOy5sVskcyjV0ilTooO1wElky0lQDqVnkG+PxWMqn57Joe29u2LBLnOofK/zZAzeoZeivYwJqJB69ZaQJfZVvRufh4VcDsvIMTyRvI8nKVSJdsmzlavmB3GNnjvzvZFBeC1fXCRM4l99yY3Zzg5uKxxTt7t/FVon5IOfFYUU7eXfwqawtX7/vm6BJPXWzdfLrGOvHL3fMLNaZXGO6OZ1SryPwAPYkya4m8JeNmlpYHkq0gLa9MTUAw0yipDMP5ZuHlxyOzx0pr+fGx/noUKrm1oMk4q2A4q7SUgLIh6OANfzf3CxfdRigqYblAo5wxY0wiY6n+ZkibbxBj1n83cszm8kPrgKpEy2Zgq15qXi4LeiwI+GohYL8cbImI9fqARA+'
	$Base64String &= '8V/Gq1pR/wCpemu/+UAkQ01o64+C+VJ8w93dhPjs1Prom+zHdmImMv479NnkMo/LzsVJgJ85bLQZ2Bk+GP3dpAey7YH4MZMTQxTHRUlZnrHRaaCa7kOyHtWY5V7jJWpA40SuAdRWZLFuVCZVa9ejUcK2YmWxEeuXWygGBaIhYJ8jkixddfqADpBLr9sq2QkP1L01l/kAjxVhklGlBfFneYcnpYvyhZMAuGfs8GtimIT7f7kdkIgg0IyV0gF+NP5xCwIPrrS7q0AmB8MX6kWi8kDitl7GQdqkcr/MPWdQcknGkwuSROzU6wEVu2UcmBLNhXlwk01TDksUOgJkPj4IvowWAP8AXBdVdq+WtFtzPSaoDQ/5AvsrR9eq1b9EvzswO7AUdsMgVxBQWZnUOWS/0vmsBmjOOTkyw9GYeKh/sJLZkP8AU9dfMB//2Q=='
	Return Binary(_Base64Decode($Base64String))
EndFunc   ;==>_BinaryString_Picture_License

Func _BinaryString_Picture_Donate()
	Local $Base64String
	$Base64String &= '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAIBAQEBAQIBAQECAgICAgQDAgICAgUEBAMEBgUGBgYFBgYGBwkIBgcJBwYGCAsICQoKCgoKBggLDAsKDAkKCgr/2wBDAQICAgICAgUDAwUKBwYHCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgr/wgARCAAUAGQDAREAAhEBAxEB/8QAHAAAAQUBAQEAAAAAAAAAAAAAAAEEBQYHAwII/8QAGQEBAQEBAQEAAAAAAAAAAAAAAAMCAQQF/9oADAMBAAIQAxAAAAH7OtHMvb5ZTGuHULSdjjSr3ltvzffI40ABVLxrFcU70wg6zQ686y1lHNM+H9iCGAg/NPv5+XXrhprkHWdnhUAAAAGW8f/EACAQAAICAgIDAQEAAAAAAAAAAAQFAwYAFQIHARQWERL/2gAIAQEAAQUCaGL0dYHe3Lkcs7NVMUo/b9MI8c721LiBuxY5z+/Ts5FA/wCqvWBz1gc9YHPWBx7HFGV2TP5+Oc/aWwCA6rVUNDx8Kuv6LE0JrwQQClgvLksLG+XcquvRe0XxTJd3G+0RHYPY4qIu/wBmUdiO/P8AU/FtLx4beXDyYmYSJKhrUu3lzby5t5c28ubeXNvLm3lzby4UVzL5/wD/xAApEQACAQQBAgUEAwAAAAAAAAABAgADERIhEyJRBDEyQXEgIyRhgZGx/9oACAEDAQE/AWKpSLdhA9e9um/axi+KRkywNviDxdE26Tv9Tnc9WgL22ItcglGTY7SpXyvhrEb1EXoGXnMVmKzFZisqAAzxJ/Ht8RuesrOmv42ZlSpA8Y9ux/2IcaC29W7fHeUM3RUXVtm4gCq33PVfeif6inlZhY9TD2PkPqq+qCoROVoz5riZTVKRuonK05WnK05WnK05WnK05WjMWn//xAAqEQACAQIEBQIHAAAAAAAAAAABAgADEQQSEyEiMTJBUXGRFCAkgaGx0f/aAAgBAgEBPwFA1SsE8mGnh8uazZb2vcfyPg3Wpk1Bf1hwGIF+Ibc9+U+GRSE3Y2ubEW/UbDKVDo9gfPOUsKEy6m+Y2Fj+ZUbjOXlMzTM0zNMzSmSRMGPqb+L+8Q4eg6U6m/nfYQrWrsuqe+/ELfYSpx4ly3Ttf18e8xJRKj1X3vsAG7Qs1RLU+i224G/e99440UQ3HAp7jqPzUumGmDNJYq5GDLzEq1KlYWczSWaSzSWaSzSWaSzSWaSxVCz/xAA2EAABAwICCAQEBAcAAAAAAAABAgMEBREAEgYTISIxNKLRFCNBURUWMpNCYXKBB0NiY3Ghwf/aAAgBAQAGPwJ+tvw2leFgKeVdA22RfDVGXO0WVVF08TDSfh7yFZPbPnI47OGG68n+HtY8MtnWLebp6VJT7+t1Ae4GImXRuoBVQbz01Bp6byxcDy9u3jf/ABh+tIkUOj0/4r4CGirUtxby3QBe+RwAb2YftiXo9WtBXpVQgLTr10aMFsqSpOZKt8jL+nbwxPOiy26cijUpyTU251Ku4Hb7jRBtbYCbi/EYjGqQ2PE+HR4jK0AM9t7/AHjk2vtjHJtfbGOTa+2Mcm19sYSGW0pGr4JFvU4bgFh9bcp+K1J8PHU6Us5klw5Ugn6QR++K1pFo44mnBSHGYITR8k2S0lI/G5tFzewticrRClSk2opYghVCla8O2/mPLFgPy4DFDjUxsuVkNzI1DllF8kUr3pWX9ATb3KgPXFD0QokQMGDFXKqEysUNxZbkKWCAjPl395e8PbCpOlKJMmtfGi5UJcqiSJOuYSbNqZ1e4jcCdvptxWmVUmoIer+kjF9dTXkAQGcpCipSbC4Sdn9WKbSkaUUWjsy4kl1yZWmipJU2pgJQnzm9p1ijxP04iNyYsSA88uGn5ckoV455LyGi48g5k7jWsXm8o8s5tTtyRarWNHG0OfJ8msvobKtW/kSwpGrc9AdYsKSQVJIHFOVS6+qb'
	$Base64String &= 'BiQ59DpqaiDMpwyyGil/y8jUtzKbs/XrPX6NlzA0Fq9foqA7EYddkKhhvXqW68nIhK5YUDZtCRlS7vKubCww2f7X/TgI1DZypAvt745Zrq74ep06A04y+0pt1F1bySLEccF+jURppamw3nU66shA/CMyjYfkMcs11d8cs11d8cs11d8cs11d8cs11d8cs11d8cs11d8cs11d8Ba0JFk2ATj/xAAhEAEBAAIBBAMBAQAAAAAAAAABEQAhMUFRwfAQYYEwof/aAAgBAQABPyEjt9Sm/s/3N+Tk+N6dyqu8mReAF6bkuL1MpnVzvVlo7AqyBemOUCl5I/rA1zcI+CNmEtSUVQ1UcJJoLXlVWoB1XFfjR6R6E0baz3zxnvnjPfPGe+eMZyioRfoyh7AlLRBxx0Y9fYp5BYPjlNbM3zC89jUAzgcl1jFNRJVIbnTaCMoSXOMHNplE/ebvkZqiniGi6pQMOwXSyBsltvCZrjmxfqQOYTU5cYHHXXboKYjzWwqMDtY+1CcxjDHwUapL28PdaZMZtBUj17bWB0rfAr2lMK3QJ0+BKVrGNKgsovGWlRShRN77Bo7H8SSSSSSSSeQQ0SVeq98//9oADAMBAAIAAwAAABDAku1tuYBXOASefb//AP8A/wD/xAAgEQEAAgMAAwADAQAAAAAAAAABABEhMWFBUXEggZHw/9oACAEDAQE/EGaGz+EoW4bwH5dpAusuwJ3zn7UA+J7ZrEQo1yxVPj7v5UQk7q2U5HNV8jtsKNsr8FeJZQMC8efM4E4E4E4EI0eIgKKKGhaLLwZ1DgZkDAivLnOaxCzhob76iivWv5Eprxv0m2u9VXuz3AL4FFaV0XWcueREhdpuA4aY158foJSxO4Ha0owa/LR8gwUTgRkERKf9cXDHW1x6ys4E4E4E4E4E4E4E4EQtn//EACARAQACAwADAQADAAAAAAAAAAEAESExYUFRcYEgkaH/2gAIAQIBAT8QPi4D+2pbmlshfyj3f7F2yqlDzxRfpZ7Caye28cr7BwNXQHLTqvtwltur1YaTA39xAAVmmGvNvO9NagBpyaz4vH+Tozozozoxjb5gKuCEWhdEMqG0fyFpfDtZXwYaKvMpBd/BcG2/eVmbKuz7Ji/6v0D6j2ZQEWDbV4wYfcKUmEFZGGxm8efcGwjSTZgBvCmefy2fYsts6MvwiWONn5AyUu9Bn3gLZ0Z0Z0Z0Z0Z0Z0Z0YBRP/8QAHxABAQEBAAEEAwAAAAAAAAAAAREAIUEQMVFhIDCB/9oACAEBAAE/EA7IDlu0yqo+V+Tmt+pO4G+dFuUcCWCR+2mANAWCYKrl4BIJSsFTgNojpCgkkEl9M/VLL492YIQQnADPN1OCQCmEfnpz/mV+Bz8a1atWW6/IRWA7A79ZqmiliF92jvmTJYuvyADJaRCNKNQSxgjJSMKADB6m3aWsmWIMbGxEELhbhBEuNQ5gAwW0qo35faywmYAHZlpg/JwoBjx9RwQCmMeLBHdYu/DRlHPCo+zcdS3zApEBFyyKnpPfN/JUlwjiSD7mGoKUCQYWJYHsHpILSTWoQ4JUJaI9y6S1Rh6CGiKI8P0SSSSSSSSTAMD5k7LX53//2Q=='
	Return Binary(_Base64Decode($Base64String))
EndFunc   ;==>_BinaryString_Picture_Donate

Func _FileSaveDialog_Ex($sTitle, $sInitDir, $sFilter, $iOptions = 0, $sDefaultName = "", $hWnd = 0, $bSaveSelection_SQL = True)
	If Not $sInitDir Then $sInitDir = @ScriptDir ; "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" ; CLSID for "My Computer"
	Local $sWorkingDir = @WorkingDir
	Local $sFileName = FileSaveDialog($sTitle, $sInitDir, $sFilter, $iOptions, $sDefaultName, $hWnd)
	Local $iError = @error
	FileChangeDir($sWorkingDir)
	Return SetError($iError, 0, $sFileName)
EndFunc   ;==>_FileSaveDialog_Ex

Func _DrawShadowText_Title($idPic)
	Local $hPic = GUICtrlGetHandle($idPic)
	; Create bitmap
	Local $tRECT = _WinAPI_GetClientRect($hPic)
	Local $iWidth = DllStructGetData($tRECT, 3) - DllStructGetData($tRECT, 1)
	Local $iHeight = DllStructGetData($tRECT, 4) - DllStructGetData($tRECT, 2)
	Local $hDC = _WinAPI_GetDC($hPic)
	Local $hDestDC = _WinAPI_CreateCompatibleDC($hDC)
	Local $hBitmap = _WinAPI_CreateCompatibleBitmap($hDC, $iWidth, $iHeight)
	Local $hDestSv = _WinAPI_SelectObject($hDestDC, $hBitmap)
	Local $hSrcDC = _WinAPI_CreateCompatibleDC($hDC)
	Local $hSource = _WinAPI_CreateCompatibleBitmapEx($hDC, $iWidth, $iHeight, _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_3DFACE)))
	Local $hSrcSv = _WinAPI_SelectObject($hSrcDC, $hSource)
	Local $hFont = _WinAPI_CreateFont(32, 21, 0, 0, 800, 1, 0, 0, $DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, $ANTIALIASED_QUALITY, $DEFAULT_PITCH, 'Calibri')
	_WinAPI_SelectObject($hSrcDC, $hFont)
	_WinAPI_DrawShadowText($hSrcDC, "SSD - SetSoundDevice", 0xFCFCFC, 0x606060, 2, 2, $tRECT, BitOR($DT_SINGLELINE, $DT_VCENTER)) ; BitOR($DT_CENTER, $DT_SINGLELINE, $DT_VCENTER))
	_WinAPI_BitBlt($hDestDC, 0, 0, $iWidth, $iHeight, $hSrcDC, 0, 0, $MERGECOPY)
	_WinAPI_ReleaseDC($hPic, $hDC)
	_WinAPI_SelectObject($hDestDC, $hDestSv)
	_WinAPI_DeleteDC($hDestDC)
	_WinAPI_SelectObject($hSrcDC, $hSrcSv)
	_WinAPI_DeleteDC($hSrcDC)
	_WinAPI_DeleteObject($hSource)
	_WinAPI_DeleteObject($hFont)
	; Set bitmap to control
	_SendMessage($hPic, $STM_SETIMAGE, 0, $hBitmap)
	Local $hObj = _SendMessage($hPic, 0x0173) ; $STM_GETIMAGE = 0x0173
	If $hObj <> $hBitmap Then _WinAPI_DeleteObject($hBitmap)
EndFunc   ;==>_DrawShadowText_Title


Func _MsgBox_SHEx($iFlag = 0, $sTitle = "", $sText = "", $iTimeout = 0, $hWnd = "", $s_Button1 = "", $s_Button2 = "", $s_Button3 = "", $b_WS_EX_TOOLWINDOW = False, $i_Transparency = 255, $hIcon = 0)

	Local $b_Add_Icon = False, $iReturn, $i_hWndState

	; SHMessageBoxCheck
	; https://msdn.microsoft.com/en-us/library/windows/desktop/bb773836(v=vs.85).aspx

	$__a_Global_MsgBox_SHEx[3] = $sTitle
	$__a_Global_MsgBox_SHEx[4] = $sText
	$__a_Global_MsgBox_SHEx[5] = $s_Button1
	$__a_Global_MsgBox_SHEx[6] = $s_Button2
	$__a_Global_MsgBox_SHEx[7] = $s_Button3
	$__a_Global_MsgBox_SHEx[8] = $b_WS_EX_TOOLWINDOW

	$__a_Global_MsgBox_SHEx[9] = Int($i_Transparency)
	If $__a_Global_MsgBox_SHEx[9] < 0 Or $__a_Global_MsgBox_SHEx[9] > 255 Then $__a_Global_MsgBox_SHEx[9] = 255

	Local $i_Misc_Settings = 0
	If $iFlag >= 2097152 And BitAND($iFlag, 2097152) Then ; MB_SERVICE_NOTIFICATION = 0x00200000 = 2097152
		$i_Misc_Settings = 2097152
		$iFlag -= 2097152
	EndIf
	If $iFlag >= 524288 And BitAND($iFlag, 524288) Then ; title and text are right-justified
		$i_Misc_Settings = 524288
		$iFlag -= 524288
	EndIf
	If $iFlag >= 262144 And BitAND($iFlag, 262144) Then ; MsgBox has top-most attribute set
		$i_Misc_Settings += 262144
		$iFlag -= 262144
	EndIf
	If $iFlag >= 16384 And BitAND($iFlag, 16384) Then ; SHow Help Button ; MB_HELP = 0x00004000 = 16384
		$i_Misc_Settings += 16384
		$iFlag -= 16384
	EndIf

	Local $i_Modality = 0
	If $iFlag >= 8192 And BitAND($iFlag, 8192) Then ; Task modal
		$i_Modality = 8192
		$iFlag -= 8192
	EndIf
	If $iFlag >= 4096 And BitAND($iFlag, 4096) Then ; System modal (dialog has an icon)
		$i_Modality = 4096
		$iFlag -= 4096
	EndIf

	Local $i_Default_Button = 0
	If $iFlag >= 512 And BitAND($iFlag, 512) Then ; Third button is default button
		$i_Default_Button = 512
		$iFlag -= 512
	EndIf
	If $iFlag >= 256 And BitAND($iFlag, 256) Then ; Second button is default button
		$i_Default_Button = 256
		$iFlag -= 256
	EndIf

	$__a_Global_MsgBox_SHEx[2] = 0
	If $iFlag >= 64 And BitAND($iFlag, 64) Then ; Information-sign icon consisting of an 'i' in a circle
		$__a_Global_MsgBox_SHEx[2] = 64
		$iFlag -= 64
		$b_Add_Icon = True
	EndIf
	If $iFlag >= 48 And BitAND($iFlag, 48) Then ; Exclamation-point icon
		$__a_Global_MsgBox_SHEx[2] = 48
		$iFlag -= 48
		$b_Add_Icon = True
	EndIf
	If $iFlag >= 32 And BitAND($iFlag, 32) Then ; Question-mark icon
		$__a_Global_MsgBox_SHEx[2] = 32
		$iFlag -= 32
		$b_Add_Icon = True
	EndIf
	If $iFlag >= 16 And BitAND($iFlag, 16) Then ; Stop-sign icon
		$__a_Global_MsgBox_SHEx[2] = 16
		$iFlag -= 16
		$b_Add_Icon = True
	EndIf

	Local $iButton = 0
	If $iFlag >= 6 And BitAND($iFlag, 6) Then ; Cancel, Try Again, Continue
		$iButton = 6
	ElseIf $iFlag >= 5 And BitAND($iFlag, 5) Then ; Retry and Cancel
		$iButton = 5
	ElseIf $iFlag >= 4 And BitAND($iFlag, 4) Then ; Yes and No
		$iButton = 4
	ElseIf $iFlag >= 3 And BitAND($iFlag, 3) Then ; Yes, No, and Cancel
		$iButton = 3
	ElseIf $iFlag >= 2 And BitAND($iFlag, 2) Then ; Abort, Retry, and Ignore
		$iButton = 2
	ElseIf $iFlag >= 1 And BitAND($iFlag, 1) Then ; OK and Cancel
		$iButton = 1
		; Else ; OK button
	EndIf

	If IsPtr($hIcon) Then
		$__a_Global_MsgBox_SHEx[2] = $hIcon
		$b_Add_Icon = True
	EndIf

	If Not $__a_Global_MsgBox_SHEx[0] Then
		$__a_Global_MsgBox_SHEx[0] = GUIRegisterMsg(_MsgBox_SHEx_RegisterWindowMessage("SHELLHOOK"), "_MsgBox_SHEx_HShellWndProc")
		Local $old_GUISwitch = GUISwitch(0)
		$__a_Global_MsgBox_SHEx[1] = GUICreate("SHELLHOOK GUI - " & @ScriptName)
		OnAutoItExitRegister("_MsgBox_SHEx_OnAutoItExitRegister_CleanUp")
		GUISwitch($old_GUISwitch)
	EndIf

	$__a_Global_MsgBox_SHEx[10] = HWnd($hWnd)
	If IsHWnd($__a_Global_MsgBox_SHEx[10]) Then
		If $__a_Global_MsgBox_SHEx[10] <> WinGetHandle(AutoItWinGetTitle()) Then
			$i_hWndState = WinGetState($__a_Global_MsgBox_SHEx[10])
			WinSetState($__a_Global_MsgBox_SHEx[10], "", @SW_DISABLE)
		EndIf
	EndIf

	_MsgBox_SHEx_ShellHookWindow($__a_Global_MsgBox_SHEx[1], 1) ; activate hook

	If $b_Add_Icon = True Then
		; Question mark symbol does not seem to raise sound / beep
		$iReturn = MsgBox($iButton + 32 + $i_Default_Button + $i_Modality + $i_Misc_Settings, $sTitle, $sText, $iTimeout)
	Else
		$iReturn = MsgBox($iButton + $i_Default_Button + $i_Modality + $i_Misc_Settings, $sTitle, $sText, $iTimeout)
	EndIf

	_MsgBox_SHEx_ShellHookWindow($__a_Global_MsgBox_SHEx[1], 0) ; de-activate hook

	If $i_hWndState Then WinSetState($__a_Global_MsgBox_SHEx[10], "", $i_hWndState)

	Return $iReturn

EndFunc   ;==>_MsgBox_SHEx

Func _MsgBox_SHEx_OnAutoItExitRegister_CleanUp()
	GUIDelete($__a_Global_MsgBox_SHEx[1])
EndFunc   ;==>_MsgBox_SHEx_OnAutoItExitRegister_CleanUp

Func _MsgBox_SHEx_RegisterWindowMessage($sText)
	Local $aRet = DllCall('user32.dll', 'int', 'RegisterWindowMessage', 'str', $sText)
	Return $aRet[0]
EndFunc   ;==>_MsgBox_SHEx_RegisterWindowMessage

Func _MsgBox_SHEx_ShellHookWindow($hWnd, $bFlag)
	Local $sFunc = 'DeregisterShellHookWindow'
	If $bFlag Then $sFunc = 'RegisterShellHookWindow'
	Local $aRet = DllCall('user32.dll', 'int', $sFunc, 'hwnd', $hWnd)
	; ConsoleWrite("_MsgBox_SHEx_ShellHookWindow: " & $aRet[0] & @crlf)
	Return $aRet[0]
EndFunc   ;==>_MsgBox_SHEx_ShellHookWindow

Func _MsgBox_SHEx_HShellWndProc($hWnd, $Msg, $wParam, $lParam)

	; ConsoleWrite($hwnd & @tab & $Msg & @tab & $wParam & @tab & $lParam & @crlf)

	If $wParam = 1 Then ; 1 = $HSHELL_WINDOWCREATED

		If $lParam <> WinGetHandle("[TITLE:" & $__a_Global_MsgBox_SHEx[3] & "; CLASS:#32770]", $__a_Global_MsgBox_SHEx[4]) Or WinGetProcess($lParam) <> @AutoItPID Then Return

		Local $hIcon, $hCtrl, $aRet

		If IsPtr($__a_Global_MsgBox_SHEx[2]) Then
			$hIcon = $__a_Global_MsgBox_SHEx[2]
		Else
			Switch $__a_Global_MsgBox_SHEx[2]
				Case 16 ; Stop-sign icon
					$hIcon = _WinAPI_LoadIcon(0, $IDI_ERROR)
				Case 32 ; Question-mark icon
					$hIcon = _WinAPI_LoadIcon(0, $IDI_QUESTION)
				Case 48 ; Exclamation-point icon
					$hIcon = _WinAPI_LoadIcon(0, $IDI_EXCLAMATION)
				Case 64 ; Information-sign icon consisting of an 'i' in a circle
					$hIcon = _WinAPI_LoadIcon(0, $IDI_INFORMATION)
			EndSwitch

		EndIf

		If $hIcon Then
			$hCtrl = ControlGetHandle($lParam, "", "Static1")
			; Local Const $STM_SETIMAGE = 0x0172
			; Local Const $IMAGE_BITMAP = 0
			; Local Const $IMAGE_ICON = 1
			$aRet = DllCall('user32.dll', 'ptr', 'SendMessage', 'hwnd', $hCtrl, 'uint', 0x0172, 'wparam', 1, 'lparam', $hIcon)
			_WinAPI_DestroyIcon($aRet[0])
		EndIf

		; WinSetTitle($lParam, "", "New_Loooooooooong_Title")
		; WinMove($lParam, "", 100, 100, 400, 400)

		If $__a_Global_MsgBox_SHEx[5] Then ControlSetText($lParam, "", "Button1", $__a_Global_MsgBox_SHEx[5])
		If $__a_Global_MsgBox_SHEx[6] Then ControlSetText($lParam, "", "Button2", $__a_Global_MsgBox_SHEx[6])
		If $__a_Global_MsgBox_SHEx[7] Then ControlSetText($lParam, "", "Button3", $__a_Global_MsgBox_SHEx[7])

		; Set optional transparency
		If $__a_Global_MsgBox_SHEx[9] <> 255 Then WinSetTrans($lParam, "", $__a_Global_MsgBox_SHEx[9])

		If $__a_Global_MsgBox_SHEx[8] Then ; Set $WS_EX_TOOLWINDOW style
			_WinAPI_SetWindowLong($lParam, $GWL_EXSTYLE, BitOR(_WinAPI_GetWindowLong($lParam, $GWL_EXSTYLE), BitOR($WS_EX_APPWINDOW, $WS_EX_TOOLWINDOW)))
			_WinAPI_SetWindowPos($lParam, $HWND_TOP, 0, 0, 0, 0, BitOR($SWP_FRAMECHANGED, $SWP_NOACTIVATE, $SWP_NOMOVE, $SWP_NOSIZE))
		EndIf

		_MsgBox_SHEx_ShellHookWindow($__a_Global_MsgBox_SHEx[1], 0) ; de-activate hook

		If IsHWnd($__a_Global_MsgBox_SHEx[10]) Then
			_WinAPI_SetWindowLong($lParam, $GWL_HWNDPARENT, $__a_Global_MsgBox_SHEx[10])
			_WinAPI_SetWindowPos($lParam, $__a_Global_MsgBox_SHEx[10], 0, 0, 0, 0, BitOR($SWP_FRAMECHANGED, $SWP_NOACTIVATE, $SWP_NOMOVE, $SWP_NOSIZE))
		EndIf

	EndIf

EndFunc   ;==>_MsgBox_SHEx_HShellWndProc

