#include <Array.au3>

$sFile = @ScriptDir & "\SSD_SetSoundDevice_v4000_stripped.au3"

If Not FileExists($sFile) Then Exit 9

Global $aDllOpen_Calls[1][2]

Global $bFirstLineFound = False, $sline
$hfile = FileOpen($sFile)
Global $iLineNumber = 0

While 1
	$iLineNumber += 1
	$sline = FileReadLine($hfile)
	If @error = -1 Then ExitLoop
	If $iLineNumber > 200 Then ExitLoop
	#cs
		MsgBox(4096, "Exit", "DllOpen() not found within the first 300 lines of the script")
		Exit
		EndIf
	#ce

	; If $bFirstLineFound And Not StringInStr($sline, "dllopen(") Then ExitLoop
	; If StringInStr($sline, "dllopen(") Then $bFirstLineFound = True

	; If $bFirstLineFound Then
	If StringInStr($sline, "dllopen(") Then
		ReDim $aDllOpen_Calls[UBound($aDllOpen_Calls) + 1][2]
		$aDllOpen_Calls[0][0] += 1
		$aDllOpen_Calls[UBound($aDllOpen_Calls) - 1][0] = $sline
		$aDllOpen_Calls[UBound($aDllOpen_Calls) - 1][1] = $sline
	EndIf

	#cs
		If StringInStr($sline, '"Address SendMessageW"') Then
		$iLineNumber += 1
		$sline = FileReadLine($hfile)
		If @error = -1 Then ExitLoop
		$sline_SendMessageW = StringTrimLeft(StringLeft($sline, StringInStr($sline, " ", 0, 2)-1), 7)
		; MsgBox(0, "", $sline_SendMessageW & @CRLF & StringLen($sline_SendMessageW))

		DllCall("user32.dll", $gk, "SendMessageW",:


		Exit
		EndIf
	#ce

WEnd
FileClose($hfile)

If Not $aDllOpen_Calls[0][0] Then
	MsgBox(4096, "Exit", "No DllOpen calls found")
	Exit
EndIf

For $i = 1 To $aDllOpen_Calls[0][0]
	$aDllOpen_Calls[$i][0] = StringTrimLeft($aDllOpen_Calls[$i][0], 7)
	$aDllOpen_Calls[$i][0] = StringLeft($aDllOpen_Calls[$i][0], StringInStr($aDllOpen_Calls[$i][0], " ") - 1)
	$aDllOpen_Calls[$i][1] = StringTrimLeft($aDllOpen_Calls[$i][1], StringInStr($aDllOpen_Calls[$i][1], '"', 0, -2))
	$aDllOpen_Calls[$i][1] = StringTrimRight($aDllOpen_Calls[$i][1], 2)
Next

$hfile = FileOpen($sFile)
$sBuffer = FileRead($hfile)
FileClose($hfile)


$hfile = FileOpen(StringReplace($sFile, "_stripped.au3", ".au3"))
Global $s_Header_Org
$iLineNumber = 0
While 1
	$iLineNumber += 1
	$sline = FileReadLine($hfile)
	If @error = -1 Then ExitLoop

	If StringLeft($sline, 1) <> "#" Then ExitLoop

	$s_Header_Org &= $sline & @CRLF

	If $iLineNumber > 300 Then
		MsgBox(4096, "Exit", "# found within the first 300 lines of the ORIGINAL script")
		Exit
	EndIf
WEnd
FileClose($hfile)

$s_Header_Org = StringReplace($s_Header_Org, "#AutoIt3Wrapper_Run_Au3Stripper=y", "#AutoIt3Wrapper_Run_Au3Stripper=n")
$s_Header_Org = StringTrimRight($s_Header_Org, 2)

$sBuffer = StringReplace($sBuffer, "#NoTrayIcon", "")

$sBuffer = $s_Header_Org & $sBuffer

For $i = 1 To $aDllOpen_Calls[0][0]
	ConsoleWrite(TimerInit() & @TAB & $i & " / " & $aDllOpen_Calls[0][0] & @TAB & $aDllOpen_Calls[$i][1] & @CRLF)
	$sBuffer = StringReplace($sBuffer, 'dllcall("' & $aDllOpen_Calls[$i][1] & '"', 'dllcall(' & $aDllOpen_Calls[$i][0])
	$sBuffer = StringReplace($sBuffer, "dllcall('" & $aDllOpen_Calls[$i][1] & "'", "dllcall(" & $aDllOpen_Calls[$i][0])
Next

FileDelete(StringReplace($sFile, ".au3", "_pp.au3"))
FileWrite(StringReplace($sFile, ".au3", "_pp.au3"), $sBuffer)





$hfile = FileOpen(StringReplace($sFile, ".au3", "_pp.au3"))
; $hfile = FileOpen(StringReplace($sFile, "_Obfuscated.au3", ".au3"))

#cs
	@SystemDir
	@WindowsDir
	@AppDataDir
	@ComSpec

	@DesktopWidth
	@DesktopHeight
#ce

$sBuffer = ""
Global $a_Macros[40][3]
; 1 = macro name
; 2 = var name
; 3 = macro name found (only replace once)
$a_Macros[0][0] = "ScriptDir"
$a_Macros[1][0] = "AutoItPID"
$a_Macros[2][0] = "CRLF"
$a_Macros[3][0] = "CR"
$a_Macros[4][0] = "LF"
$a_Macros[5][0] = "TAB"
$a_Macros[6][0] = "SW_HIDE"
$a_Macros[7][0] = "SW_SHOWDEFAULT"
$a_Macros[8][0] = "SW_SHOWMAXIMIZED"
$a_Macros[9][0] = "SW_SHOWMINIMIZED"
$a_Macros[10][0] = "SW_SHOWMINNOACTIVE"
$a_Macros[11][0] = "SW_SHOWNOACTIVATE"
$a_Macros[12][0] = "SW_SHOWNORMAL"
$a_Macros[13][0] = "SW_SHOWNA"
$a_Macros[14][0] = "SW_SHOW"
$a_Macros[15][0] = "SW_DISABLE"
$a_Macros[16][0] = "SW_ENABLE"
$a_Macros[17][0] = "SW_LOCK"
$a_Macros[18][0] = "SW_MAXIMIZE"
$a_Macros[19][0] = "SW_MINIMIZE"
$a_Macros[20][0] = "SW_RESTORE"
$a_Macros[21][0] = "SW_UNLOCK"
$a_Macros[22][0] = "ScriptFullPath"
$a_Macros[23][0] = "ScriptName"
$a_Macros[24][0] = "AutoItExe"
$a_Macros[25][0] = "AutoItVersion"
$a_Macros[26][0] = "AutoItX64"
$a_Macros[27][0] = "Compiled"
$a_Macros[28][0] = "CPUArch"
$a_Macros[29][0] = "OSArch"
$a_Macros[30][0] = "OSBuild"
$a_Macros[31][0] = "OSVersion"
$a_Macros[32][0] = "OSType"
$a_Macros[33][0] = "ComputerName"
$a_Macros[34][0] = "SystemDir"
$a_Macros[35][0] = "AppDataDir"
$a_Macros[36][0] = "WindowsDir"
$a_Macros[37][0] = "TempDir"
$a_Macros[38][0] = "ProgramFilesDir"
$a_Macros[39][0] = "ComSpec"

; Global $g_Macro_AutoItExe = @AutoItExe, $g_Macro_AutoItPID = @AutoItPID, $g_Macro_AutoItVersion = @AutoItVersion, $g_Macro_AutoItX64 = @AutoItX64, $g_Macro_Compiled = @Compiled, $g_Macro_CPUArch = @CPUArch, $g_Macro_CR = @CR, $g_Macro_CRLF = @CRLF, $g_Macro_LF = @LF, $g_Macro_OSArch = @OSArch, $g_Macro_OSBuild = @OSBuild, $g_Macro_ScriptDir = @ScriptDir, $g_Macro_ScriptFullPath = @ScriptFullPath, $g_Macro_ScriptName = @ScriptName, $g_Macro_SW_DISABLE = @SW_DISABLE, $g_Macro_SW_ENABLE = @SW_ENABLE, $g_Macro_SW_HIDE = @SW_HIDE, $g_Macro_SW_LOCK = @SW_LOCK, $g_Macro_SW_MAXIMIZE = @SW_MAXIMIZE, $g_Macro_SW_MINIMIZE = @SW_MINIMIZE, $g_Macro_SW_RESTORE = @SW_RESTORE, $g_Macro_SW_SHOW = @SW_SHOW, $g_Macro_SW_SHOWDEFAULT = @SW_SHOWDEFAULT, $g_Macro_SW_SHOWMAXIMIZED = @SW_SHOWMAXIMIZED, $g_Macro_SW_SHOWMINIMIZED = @SW_SHOWMINIMIZED, $g_Macro_SW_SHOWMINNOACTIVE = @SW_SHOWMINNOACTIVE, $g_Macro_SW_SHOWNA = @SW_SHOWNA, $g_Macro_SW_SHOWNOACTIVATE = @SW_SHOWNOACTIVATE, $g_Macro_SW_SHOWNORMAL = @SW_SHOWNORMAL, $g_Macro_SW_UNLOCK = @SW_UNLOCK, $g_Macro_TAB = @TAB, $g_Macro_OSVersion = @OSVersion, $g_Macro_OSType = @OSType
; ConsoleWrite("Obfuscation Prevention" & $g_Macro_AutoItExe & $g_Macro_AutoItPID & $g_Macro_AutoItVersion & $g_Macro_AutoItX64 & $g_Macro_Compiled & $g_Macro_CPUArch & $g_Macro_CR & $g_Macro_CRLF & $g_Macro_LF & $g_Macro_OSArch & $g_Macro_OSBuild & $g_Macro_ScriptDir & $g_Macro_ScriptFullPath & $g_Macro_ScriptName & $g_Macro_SW_DISABLE & $g_Macro_SW_ENABLE & $g_Macro_SW_HIDE & $g_Macro_SW_LOCK & $g_Macro_SW_MAXIMIZE & $g_Macro_SW_MINIMIZE & $g_Macro_SW_RESTORE & $g_Macro_SW_SHOW & $g_Macro_SW_SHOWDEFAULT & $g_Macro_SW_SHOWMAXIMIZED & $g_Macro_SW_SHOWMINIMIZED & $g_Macro_SW_SHOWMINNOACTIVE & $g_Macro_SW_SHOWNA & $g_Macro_SW_SHOWNOACTIVATE & $g_Macro_SW_SHOWNORMAL & $g_Macro_SW_UNLOCK & $g_Macro_TAB & @CRLF & $g_Macro_OSVersion & $g_Macro_OSType)

Global $i_MacroLinesFound = 0

ConsoleWrite("-----------" & @CRLF & "$g_Macro_" & @CRLF)

$timer = TimerInit()
$iLineNumber = 0
While 1
	$sline = FileReadLine($hfile)
	If @error = -1 Then ExitLoop
	$iLineNumber += 1

	If $i_MacroLinesFound < 4 Then

		If (StringInStr($sline, "@CR", 2) And StringInStr($sline, "@CRLF", 2) And StringInStr($sline, "@LF", 2) And StringInStr($sline, "@TAB", 2)) Or _
				(StringInStr($sline, "@Compiled", 2) And StringInStr($sline, "@SW_DISABLE", 2) And StringInStr($sline, "@SW_ENABLE", 2) And StringInStr($sline, "@SW_HIDE", 2)) Or _
				(StringInStr($sline, "@AutoItExe", 2) And StringInStr($sline, "@AutoItPID", 2) And StringInStr($sline, "@AutoItVersion", 2) And StringInStr($sline, "@AutoItX64", 2)) Then

			$i_MacroLinesFound += 1

			$sBuffer &= $sline & @CRLF

			$sline = StringReplace($sline, "Global Const ", "")
			$a_Line_Macros = StringSplit($sline, ",")
			For $i = 1 To $a_Line_Macros[0]
				$a_Line_Macros_Split = StringSplit($a_Line_Macros[$i], " = ", 1)
				For $y = 0 To UBound($a_Macros) - 1
					If StringInStr($a_Line_Macros_Split[2], $a_Macros[$y][0]) And Not $a_Macros[$y][2] Then
						$a_Macros[$y][1] = StringStripWS($a_Line_Macros_Split[1], 8)
						$a_Macros[$y][2] = 1
						$a_Line_Macros[$i] = ""
						$a_Line_Macros_Split[1] = ""
						$a_Line_Macros_Split[2] = ""
					EndIf
				Next
			Next
			; _ArrayDisplay($a_Macros)
			ContinueLoop
		EndIf
	EndIf

	For $i = 0 To UBound($a_Macros) - 1
		$sline = StringReplace($sline, "@" & $a_Macros[$i][0], $a_Macros[$i][1], 0, 2)
	Next

	$sBuffer &= $sline & @CRLF

	If Not Mod($iLineNumber, 100) Then
		ConsoleWrite($iLineNumber & @TAB & TimerDiff($timer) & @CRLF)
	EndIf

WEnd
FileClose($hfile)

FileDelete(StringReplace($sFile, ".au3", "_pp.au3"))
FileWrite(StringReplace($sFile, ".au3", "_pp.au3"), $sBuffer)

ShellExecute(StringReplace($sFile, ".au3", "_pp.au3"), "", "", "open")