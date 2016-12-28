#include-Once
#include <IE.au3>
#include <Excel.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>


Global $ResolucaoTela = @DesktopWidth & " x " & @DesktopHeight
Local $xP,$yP

Func execMSAccess ($dbname, $tblname,$CAMPOS,$VALORES)
   LOCAL $adoCon = ObjCreate("ADODB.Connection")
   LOCAL $SQL = "Insert into " & $tblname & " (["& _ArrayToString($CAMPOS,"],[") &"]) values ("& _ArrayToString($VALORES,", ") & ");"

   ;$adoCon.Open("Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & $dbname) ;Use this line if using MS Access 2003 and lower
   $adoCon.Open ("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $dbname) ;Use this line if using MS Access 2007 and using the .accdb file extension

   if $adoCon.state = 1 Then
	  $adoCon.Execute($SQL)
	  Return 1
   Else
	  Return 0
   EndIf

   $adoCon.Close

EndFunc

Func _SelecionarArquivo($sExtensao)
    ; Create a constant variable in Local scope of the message to display in FileOpenDialog.
    Local Const $sMessage = "Selecione um único arquivo."

    ; Display an open dialog to select a file.
    Local $sFileOpenDialog = FileOpenDialog($sMessage, @ScriptDir & "\", "All (" & $sExtensao & ")", $FD_FILEMUSTEXIST)
    If @error Then
        ; Display the error message.
        MsgBox($MB_SYSTEMMODAL, "", "Nenhum arquivo foi selecionado.")

        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)
    Else
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)

        ; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
        $sFileOpenDialog = StringReplace($sFileOpenDialog, "|", @CRLF)

        ; Display the selected file.
;~         MsgBox($MB_SYSTEMMODAL, "", "Você escolheu o seguinte arquivo:" & @CRLF & $sFileOpenDialog)

		Return $sFileOpenDialog

    EndIf
EndFunc

Func SendWait($text,$sleep,$repeat=1,$format="%s",$flag=0)
   for $j = 1 to $repeat
	 Send(StringFormat($format,$text),$flag)
	 Sleep($sleep/$repeat)
  Next
EndFunc

Func _GetFile($sFile, $iFormat = 0)
    Local $hFileOpen = FileOpen($sFile, $iFormat)
    If $hFileOpen = -1 Then
        Return SetError(1, 0, "")
    EndIf
    Local $sData = FileRead($hFileOpen)
    FileClose($hFileOpen)
    Return $sData
EndFunc

Func _GetXML($sString, $sData)
    Local $aError[2] = [1, $sString], $aReturn
    $aReturn = StringRegExp('<' & $sData & '></' & $sData & '>' & $sString, '(?s)(?i)<' & $sData & '>(.*?)</' & $sData & '>', 3)
    If @error Then
        Return SetError(1, 0, $aError)
    EndIf
    $aReturn[0] = UBound($aReturn, 1) - 1
    Return SetError(0, 0, $aReturn)
 EndFunc

func _ALLTRIM($sString, $sTrimChars=' ')

	;  Trim from left first, then right

	$sTrimChars = StringReplace( $sTrimChars, "%%whs%%", " " & chr(9) & chr(11) & chr(12) & @CRLF )
	local $sStringWork = ""

	$sStringWork = _LTRIM($sString, $sTrimChars)
	if $sStringWork <> "" then
		$sStringWork = _RTRIM($sStringWork, $sTrimChars)
	endif
	return $sStringWork

endfunc

func _LTRIM($sString, $sTrimChars=' ')

	$sTrimChars = StringReplace( $sTrimChars, "%%whs%%", " " & chr(9) & chr(11) & chr(12) & @CRLF )
	local $nCount, $nFoundChar
	local $aStringArray = StringSplit($sString, "")
	local $aCharsArray = StringSplit($sTrimChars, "")

	for $nCount = 1 to $aStringArray[0]
		$nFoundChar = 0
		for $i = 1 to $aCharsArray[0]
			if $aCharsArray[$i] = $aStringArray[$nCount] then
				$nFoundChar = 1
			EndIf
		next
		if $nFoundChar = 0 then return StringTrimLeft( $sString, ($nCount-1) )
	next
endfunc

func _RTRIM($sString, $sTrimChars=' ')

	$sTrimChars = StringReplace( $sTrimChars, "%%whs%%", " " & chr(9) & chr(11) & chr(12) & @CRLF )
	local $nCount, $nFoundChar
	local $aStringArray = StringSplit($sString, "")
	local $aCharsArray = StringSplit($sTrimChars, "")

	for $nCount = $aStringArray[0] to 1 step -1
		$nFoundChar = 0
		for $i = 1 to $aCharsArray[0]
			if $aCharsArray[$i] = $aStringArray[$nCount] then
				$nFoundChar = 1
			EndIf
		next
		if $nFoundChar = 0 then return StringTrimRight( $sString, ($aStringArray[0] - $nCount) )
	next
 endfunc

Func _ShowStatus($Texto, $Titulo = "")
   IF $Texto <> "" THEN
;~ 	  ToolTip($Texto,(@DesktopWidth - 100),(@DesktopHeight - 75),$Titulo,1,4)
	  ToolTip($Texto,(@DesktopWidth),(@DesktopHeight),$Titulo,1,4)
   Else
	  ToolTip("")
   EndIf
EndFunc

Func _Retorno ($sFilePath, $sLin, $sCol, $sCmp)

   Local $hFileOpen  = FileOpen($sFilePath, $FO_READ)
   Local $sFileRead  = FileReadLine($hFileOpen, $sLin)
   local $sRetorno = StringMid($sFileRead,$sCol,$sCmp)

   Return $sRetorno

EndFunc

Func _CheckActivation($x,$y,$SleepTime); 16777215 = branco | 0 = preto
   Sleep($SleepTime)
   MouseMove($x,$y)
   $sCollor = PixelGetColor($x,$y)
   LOCAL $bIsOk = ""

   If $sCollor == "16777215" then
	  $bIsOk = "ok"
   EndIf

   Return $bIsOk

EndFunc

Func _CaracteresEspeciais($caract)
	local $codiA = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
	local $codiB = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
	local $temp = $caract
	local $result = ""

		For $i = 1 To StringLen($temp)
			$p = StringInStr($codiA, StringMid($temp, $i, 1))
			If $p > 0 Then
				$result = $result & StringMid($codiB, $p, 1)
			else
				$result = $result & StringMid($temp, $i, 1)
			endif
		Next

	Return $result
EndFunc

Func _ATIVAR_JANELA($WindowTitle)
   ;Faz ativação da janela
   WinActivate("[REGEXPTITLE:" & $WindowTitle & ".+]")

   ;Aguarda ativação da janela por 10 segundos
   $nSapAtv = WinWaitActive("[REGEXPTITLE:" & $WindowTitle & ".+]", "",10)

EndFunc

Func SheetExists($sWorkbook,$sht)
	Local $aWShts = _Excel_SheetList($sWorkbook)
	For $i=0 to UBound($aWShts)-1
	   If $aWShts[$i][0] = $sht Then
		   Return 1
	   Else
		   Return 0
	   EndIf
	Next
EndFunc

Func _ProgressBar($aList)
	Local $iIndex,$iStatus

	for $iIndex = 0 to Ubound($aList)-1
		ProgressOn("Carregando Registros da Planilha", "Por Favor Aguarde!", "0%")
		$iStatus = ($iIndex / (Ubound($aList)-1)) * 100
		ProgressSet($iStatus, StringLeft($iStatus,4) & "%")
		Sleep(100)
	Next

	ProgressSet(100, "Carregamento Completo", "Completo")
	ProgressOff()

EndFunc

Func _getFileName($sFileIn)
	local $i

	For $i = StringLen($sFileIn) To 1 Step -1
	 If StringInStr("\", StringMid($sFileIn, $i, 1)) Then
		 ExitLoop
	 EndIf
	Next

	Return StringLeft(StringMid($sFileIn, $i + 1, StringLen($sFileIn) - $i), StringLen(StringMid($sFileIn, $i + 1, StringLen($sFileIn) - $i)) - 4)

EndFunc

Func connectApp($oIE,$eForm,$eUser,$ePws,$User,$Pws)
	Sleep(2000)

	_IELoadWait($oIE)

	local $oForm	=	_IEFormGetObjByName($oIE, $eForm)
	local $oUser	=	_IEFormElementGetObjByName($oForm, $eUser)
	local $oPwd		=	_IEFormElementGetObjByName($oForm, $ePws)

	local $HWND = _IEPropertyGet($oIE, "hwnd")
	WinSetState($HWND, "", @SW_MAXIMIZE)

   Sleep(2000)

	_IEFormElementSetValue($oUser, $User)
	_IEFormElementSetValue($oPwd, $Pws)

   Sleep(2000)

	_IEFormSubmit($oForm)

EndFunc

func PosicionamentoXY($oIE)
   local $hWnd = _IEPropertyGet($oIE, "hwnd")
   WinActivate($hWnd)

   While WinExists($hWnd)
	  ;Captura o tamanho da tela
	  local $aClientSize = WinGetClientSize($hWnd)

	  local $aPos = MouseGetPos()
	  ToolTip($aClientSize[0] & "x" & $aClientSize[1] & @CR & "x: " & $aPos[0] & @CR & "y: " & $aPos[1], Default, Default, Default, Default, 4)
	  $xP = $aPos[0]
	  $yP = $aPos[1]
	  Sleep(100)
   WEnd
EndFunc

Func _MARCACAO_CORINGA($xPosicao,$yPosicao,$SleepTime)
	Sleep($SleepTime)
	MouseMove($xPosicao,$yPosicao)
	MouseClick("left")
	EndFunc

Func CaptureEND()
   Switch @HotKeyPressed ; The last hotkey pressed.
	   Case "^{end}" ; String is the {end} hotkey.
			Exit
	   Case "^{p}" ; String is the {end} hotkey.
		   createLog()
   EndSwitch
EndFunc

func createLog()
	;Open the file temp.txt in overwrite mode. If the folder C:\AutomationDevelopers does not exist, it will be created.
	Local $hFileOpen = FileOpen(@DesktopDir&"\log.txt", $FO_APPEND + $FO_CREATEPATH)

	;Display a message box in case of any errors.
	If $hFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "", "An error occurred when opening the file.")
	EndIf

	Local $sValue = InputBox("Marcação", "Informe o nome da marcação: ", "", "")

	FileWrite($hFileOpen, $xP & "," & $yP & @TAB &"|" & @TAB & $sValue & @CRLF)

	;Close the handle returned by FileOpen.
	FileClose($hFileOpen)
EndFunc

Func compactFile($fileName,$fileZip)
Local $fileApp = "\7-Zip\7z.exe  a " & $fileZip & " "
Local $sEnvVar = EnvGet("ProgramFiles")

Run($sEnvVar & $fileApp & $fileName  , "", @SW_SHOWMINIMIZED)

EndFunc