 ; ****************************************************************************
 ;
 ;	����� ������� ��� ���� 1�:������� 2
 ;	� 2013-2015 Liris 
 ;	mailto:liris@ngs.ru
 ;
 ;*****************************************************************************

#include <GUIConstantsEx.au3>
#include <ProgressConstants.au3>
#include <GUIConstantsEx.au3>
#include <ProgressConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#NoTrayIcon

Global	$sLogfileName	; ��� log-�����
Global	$sRetailIBConn	; ������ ����������� � ��
Global	$sV8exePath		; ���� � ������������ ����� 1cv8.exe
Global	$sArcLogfileName; ��� ������ log-�����
Global	$iLogMaxSize	; ������������ ������ log-�����
Global	$cPIDFileExt	; ���������� pid-�����
Global	$bShowProgress	; ��������/�� ���������� ���� ���������
Global	$bShowTrayTip	; ��������/�� ���������� ��������� � ����
Global	$bUpdateRS		; ��������� ������� �������� ���������������������������������������
Global	$bApplyCfg		; ��������� ��������� ������������
Global	$sComConnectorObj	; ��� COM-������� 
Global	$v8ComConnector, $connRetail

; ���������� ��� �������� ����������
Global	$ProgressBar, $LabText

; ������ �������� �� ini-����� � ������� ���������� ��������� ������
Func ReadParamsFromIni()
	
	$sIniFileName	= StringRegExpReplace(@ScriptFullPath, '^(?:.*\\)([^\\]*?)(?:\.[^.]+)?$', '\1')
	$sIniFileName	= @ScriptDir & "\" & $sIniFileName & ".ini"
	; ������ �� INI-����� �������� 'Key' � ������ 'EXCHANGE'.
	$sLogfileName		= IniRead($sIniFileName,	"EXCHANGE", "LogfileName", "retail_exchange.log")
	$sRetailIBConn		= IniRead($sIniFileName,	"EXCHANGE", "RetailIBConn", "File=D:\Retail;Usr=Admin;Pwd=Admin;")
	$sV8exePath			= IniRead($sIniFileName,	"EXCHANGE", "V8exePath", """C:\Program Files\1cv82\common\1cestart.exe""" )
	$sComConnectorObj	= IniRead($sIniFileName,	"EXCHANGE", "ComConnectorObj", "V83.COMConnector")
	$sArcLogfileName	= IniRead($sIniFileName,	"EXCHANGE", "ArcLogfileName", "retail_exchange_DDMMYYYY_log.old")
	$cPIDFileExt		= IniRead($sIniFileName,	"EXCHANGE", "PIDFileExt", "pid")	
	$iLogMaxSize		= Int(IniRead($sIniFileName,	"EXCHANGE", "LogMaxSize", 512000))
	$bShowTrayTip		= Int(IniRead($sIniFileName,	"EXCHANGE", "ShowTrayTip", 1))
	$bShowProgress		= Int(IniRead($sIniFileName,	"EXCHANGE", "ShowProgress", 0))
	$bUpdateRS			= Int(IniRead($sIniFileName,	"EXCHANGE", "UpdateRS", 0))
	$bApplyCfg			= Int(IniRead($sIniFileName,	"EXCHANGE", "ApplyCfg", 0))

	; ������ ���������� ��������� ������
	; ���������, ���������� � ��������� ������ ����� ��������� ��� ����������� ini-�����
	$bShowProgress	= False
	$iParamCount	= $CmdLine[0]
	
	For $iCurrParam = 1 To $iParamCount
		Select
			Case StringLower($CmdLine[$iCurrParam]) = StringLower("ShowProgress")
				$bShowProgress	= True
			Case StringLower($CmdLine[$iCurrParam]) = StringLower("ShowTrayTip")
				$bShowTrayTip	= True
			Case StringLower($CmdLine[$iCurrParam]) = StringLower("UpdateRS")
				$bUpdateRS		= True
			Case StringLower($CmdLine[$iCurrParam]) = StringLower("ApplyCfg")
				$bApplyCfg		= True
		EndSelect
	Next

EndFunc

; ������� ��� ������ � ��� ������
; *****************************************************************************

Func AddToLog($sMsg)
	
	PrepareLogFile()
	$mCurrFolder	=	@ScriptDir
	$mLogfilePath	=	$mCurrFolder & "\" & $sLogfileName
	$fFileLog		=	FileOpen($mLogfilePath, 1)
	$sWriteToFile	=	GetTimestampString() & $sMsg
	FileWriteLine($fFileLog, $sWriteToFile)
	FileClose($fFileLog)
	
EndFunc

Func GetTimestampString()
	Local $sDT
	$sDT	= "[" & @MDAY & "." & @MON & "." & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC & "." & @MSEC & "] ";
	Return $sDT
EndFunc

Func PrepareLogFile()
	$mCurrFolder	=	@ScriptDir
	$mLogfilePath	=	$mCurrFolder & "\" & $sLogfileName
	If (FileExists($mLogfilePath)) Then
	
		$fLogfileSize	=	FileGetSize($mLogfilePath)
		
		If ($fLogfileSize >= $iLogMaxSize) Then
		
			;	��������� ������ ��� ���������
			$sDT	= @YEAR & "_" & @MON & "_" & @MDAY
			;	�������� ������ � ����� �����
			$sNewFileName	=	StringReplace($sArcLogfileName,"DDMMYYYY", $sDT)
			$sNewFileName	=	$mCurrFolder & "\" & $sNewFileName
			;	���������� ���� � �����
			FileMove($mLogfilePath, $sNewFileName)
			;	������� ����� ���� ����
			$fFileLog		=	FileOpen($mLogfilePath, 1)
			$sWriteToFile	=	GetTimestampString() & "���������� ��� ��������� � �����: " & $sNewFileName
			FileWriteLine($fFileLog, $sWriteToFile)
			FileClose($fFileLog)
		EndIf

	Else
		$fFileLog		=	FileOpen($mLogfilePath, 1)
		$sWriteToFile	=	GetTimestampString() & "����������� ����� ���-���� "
		FileWriteLine($fFileLog, $sWriteToFile)
		FileClose($fFileLog)
	EndIf

EndFunc

; *****************************************************************************

; ������� ��� ������ � ����������
; *****************************************************************************

; ���������� PID ����������� ��������. ���� � ������� ������ �������� �������� �� ��������, ���������� 0
Func GetLastPID()
	
	$mReturn		=	0;
	$mCurrFolder	=	@ScriptDir;
	$hSearch		=	FileFindFirstFile($mCurrFolder & "\*." & $cPIDFileExt)
	; ��������, �������� �� ����� ��������
	If $hSearch = -1 Then
		; ������ ��� ������ ������
		; ����� ������� 0
		Return 0
		Exit
	EndIf	
	
	While 1
		$sFile = FileFindNextFile($hSearch) ; ���������� ��� ���������� �����, ������� �� ������� �� ����������
		If @error Then ExitLoop
		
		AddToLog("������ PID-����: " & $sFile)
		$sPID	=	StringReplace($sFile, "." & $cPIDFileExt, "")
		;AddToLog("������������� ��������: " & $sPID)
		$mReturn=	Int($sPID)
	WEnd
	
	FileClose($hSearch)
	
	return $mReturn
	
EndFunc

; ������� ��� PID-����� � ����� �������
Func DeletePIDFile()
	
	UpdateProgress(100, "��������� ������ ��������� ������")
	AddToLog("������� ������� PID-����� � �����: " & @ScriptDir)
	$mCurrFolder	=	@ScriptDir;
	$iResult		=	FileDelete($mCurrFolder & "\*." & $cPIDFileExt)
	If $iResult > 0 Then
		AddToLog("PID-����� ������� �������")
	Else
		AddToLog("��� �������� ������ ��������� ������, ���� ��� ������ ��� ��������")
	EndIf
	
EndFunc

; ������� ����� PID-����
Func CreatePIDFile()
	
	$mCurrFolder	=	@ScriptDir;
	$iCurrentPID	=	@AutoItPID;
	$sPIDFileName	=	$mCurrFolder & "\" & $iCurrentPID & "." & $cPIDFileExt;
	
	$fFileOut		=	FileOpen($sPIDFileName, 1)
	FileWrite($fFileOut, String($iCurrentPID))
	FileClose($fFileOut)
	
	$sMsg = "������ ����� PID-����: " & String($iCurrentPID)
	AddToLog($sMsg)
	UpdateProgress(20, $sMsg)
	
EndFunc	

; ������� ������ � ������
; *****************************************************************************

Func CreateProgressForm()

	$_DH	= Ceiling(@DesktopHeight / 20)
	$_DW	= Ceiling(@DesktopWidth / 30)

	GUICreate("�������� ��������", 12*$_DW, 6*$_DH)
	$ProgressBar	= GUICtrlCreateProgress(2*$_DW, $_DH, 8*$_DW, 2*$_DH)
	$LabText		= GUICtrlCreateLabel("", 2*$_DW, 4*$_DH, 8*$_DW, 2*$_DH, $SS_CENTER)
	GUICtrlSetFont($LabText, 20)
	GUISetState()
	Opt("GUICloseOnESC", 0)
	
EndFunc

Func UpdateProgress($iProgress, $sStatusText)
	
	; �������� ������������� ��������
	If GUICtrlGetHandle($Progressbar) <> 0 Then
		; ��������� ��������� ������������
		GUICtrlSetData($Progressbar, $iProgress)
	EndIf
	
	If GUICtrlGetHandle($LabText) <> 0 Then
		; ���������� ��������� ����������
		GUICtrlSetData($LabText, $sStatusText)
	EndIf
	
	If $bShowTrayTip Then
		ToolTip($sStatusText, @DesktopWidth - 260, @DesktopHeight - 90, "����� �������", 1)
	EndIf
	
EndFunc

; *****************************************************************************

; ��������� ������ ����������� �� ������������
Func SplitConnectionString($sIBConn)
	
	Local $aResult[3]
	$aIBParams = StringSplit($sIBConn, ";")
	$iParamCount	= $aIBParams[0]

	For $iCurrParam = 1 To $iParamCount

		$iLength	= StringLen($aIBParams[$iCurrParam])
		If $iLength = 0 Then ContinueLoop

		$sParamName	= StringLeft($aIBParams[$iCurrParam], StringInStr($aIBParams[$iCurrParam], "=") -1)
		
		Select
			Case StringLower($sParamName) = StringLower("File")
				$aResult[0]	= StringReplace($aIBParams[$iCurrParam], $sParamName & "=", "")
			Case StringLower($sParamName) = StringLower("Usr")
				$aResult[1]	= StringReplace($aIBParams[$iCurrParam], $sParamName & "=", "")
			Case StringLower($sParamName) = StringLower("Pwd")
				$aResult[2]= StringReplace($aIBParams[$iCurrParam], $sParamName & "=", "")
		EndSelect

	Next
	
	Return $aResult
	
EndFunc

; ������� �������� ��������� ��������� 
Func RunUpdateCfg()

	;v8exe & " DESIGNER /F" & $sIBPath  & " /N" & IBAdminName & " /P" & IBAdminPwd & " /WA- /UpdateDBCfg /Out" & $ServiceFileName & " -NoTruncate /DisableStartupMessages"
	Local $aConnParams[3]
	Local $sUpdCmdLine, $sRunClientCmdLine
	Local $sIBPath, $sIBAdmin, $sIBAdminPwd, $ServiceFileName

	; ��� ������ ����� ����������� �� ���� ������
	While IsObj($connRetail) OR IsObj($v8ComConnector)
		
		$connRetail	= 0
		$v8ComConnector = 0
		UpdateProgress(50, "�������� ������������ ������," & @CRLF & " ������� COM-���������")
		Sleep(500)
		AddToLog("�������� ������������ ������ �� ��������")

	WEnd
	
	$sQuestion	=	"��������� ������������!" & @CRLF & "�������� ��������� ������������ ���� ������." & @CRLF
	$sQuestion	=	$sQuestion & "����������, �������� ��� � ������� ������ ��" & @CRLF & "����� ��������� �� �������� ��������� ������� ������"
	$bQResult	=	MsgBox(1 + 32, "��������� ���������� ���������", $sQuestion, 60)	
	
	If $bQResult = 1 Then
	
		$sServiceFileName	=	@ScriptDir & "\1c_update.txt" 
		; ������� ��������� ����������� �� ������ �����������
		; !���������� ������ �� �������� �������� ��
		$aConnParams= SplitConnectionString($sRetailIBConn)
		$sIBPath	=	$aConnParams[0]
		$sIBAdmin	=	$aConnParams[1]
		$sIBAdminPwd=	$aConnParams[2]
		
		; ���� �������� ���������, ������� �������� �������������
		
		$sUpdCmdLine	=	$sV8exePath & " DESIGNER /F" & $sIBPath  & " /N" & $sIBAdminPwd & " /P" & $sIBAdminPwd 
		$sUpdCmdLine	=	$sUpdCmdLine & " /WA- /UpdateDBCfg /Out""" & $sServiceFileName & """ -NoTruncate /DisableStartupMessages"
		; �������� ���������
		UpdateProgress(50, "����������� �������� ���������")
		AddToLog("����������� �������� ���������")
		$PIDUpdCfg	= Run($sUpdCmdLine)
		ProcessWaitClose($PIDUpdCfg)
		UpdateProgress(50, "������� �� ��������" & @CRLF & "��������� ���������")
		AddToLog("������� �� �������� ��������� ���������")
		
	Else
		
		; ������� ��� ������������ ��������� �� ���������� ����������
		UpdateProgress(50, "����� �������� ���������")
		AddToLog("������ �������� ���������. ��������� �������: " & String($bQResult) )
		
	EndIf
		
EndFunc

Func RunApplyCfg()

	;v8exe & " DESIGNER /F" & $sIBPath  & " /N" & IBAdminName & " /P" & IBAdminPwd & " /WA- /UpdateDBCfg /Out" & $ServiceFileName & " -NoTruncate /DisableStartupMessages"
	Local $aConnParams[3]
	Local $sUpdCmdLine, $sRunClientCmdLine
	Local $sIBPath, $sIBAdmin, $sIBAdminPwd, $ServiceFileName
	
	; ��� ������ ����� ����������� �� ���� ������
	While IsObj($connRetail) OR IsObj($v8ComConnector)
		$connRetail	= 0
		$v8ComConnector = 0
		Sleep(500)
		AddToLog("�������� ������������ ������ �� ��������")
	WEnd
	
	$sQuestion	=	"��������� ������������!" & @CRLF & "���� ������ ������������� ��� ����������" & @CRLF
	$sQuestion	=	$sQuestion & "����������, �������� ��� � ������� ������ ��" & @CRLF & "����� �������� ��� �������� ������� ������"
	$bQResult	=	MsgBox(1 + 32, "�� ������������� ��� ����������", $sQuestion, 60)
	
	If $bQResult = 1 Then
	
		; ������� ��������� ����������� �� ������ �����������
		; ���������� ������ �� �������� �������� ���� ������
		$aConnParams= SplitConnectionString($sRetailIBConn)
		$sIBPath	=	$aConnParams[0]
		$sIBAdmin	=	$aConnParams[1]
		$sIBAdminPwd=	$aConnParams[2]
		
		; ������ ������� ��� ���������� ���������
		$sRunClientCmdLine	=	$sV8exePath & " ENTERPRISE /F""" & $sIBPath  & """ /N" & $sIBAdminPwd & " /P" & $sIBAdminPwd 
		$sRunClientCmdLine	=	$sRunClientCmdLine & " /WA-"
		UpdateProgress(60, "����������� ������ ����������" & @CRLF & "��� �������� ���������")
		AddToLog("����������� ������ ����������� ����������")
		$PIDApplyCfg		=	Run($sRunClientCmdLine)
		ProcessWaitClose($PIDApplyCfg)
		UpdateProgress(60, "���������� ���������" & @CRLF & "��������� �������")
		AddToLog("���������� ���� ������ ���������")
		
	Else
		
		; ������� ��� ������������ ��������� �� �������� ����������
		UpdateProgress(60, "������ �������� ���������")
		AddToLog("������ �������� ���������. ��������� �������: " & String($bQResult) )
		
	EndIf
		
EndFunc
 
; ������� ��������� ��������� ������ �������
Func RunExchange()

	$v8ComConnector = ObjCreate($sComConnectorObj)
	If @error Then
		AddToLog("������ ��� �������� COM-������� " & $sComConnectorObj)
		UpdateProgress(20, "������ ��� �������� COM-������� " & $sComConnectorObj)
		Exit
	EndIf
	
	AddToLog("������� ����������� � ��");
	UpdateProgress(25, "������� ����������� � ��")
	
	$connRetail	=	$v8ComConnector.Connect($sRetailIBConn);

	If Not IsObj($connRetail) Then
		
		UpdateProgress(30, "��� ����������� � �� ��������� ������")
		AddToLog("��� ����������� � �� ��������� ������")
		AddToLog("������ ��������� ������ ��-�� ������ ���������� � ��")
		; ����� �������� �������� alarm'a? �������� �� SMS ��� email
		Exit
		
	Else
		AddToLog("����������� � �� �����������")
		UpdateProgress(30, "������������ � ��")
	EndIf
	
	If $connRetail.��������������������() Then
		UpdateProgress(40, "������������ �� ��������."& @CRLF & "������� ��������� ���������")
		AddToLog("������������ �� ��������. ��������� ���������� ���������")
		RunUpdateCfg()
		DeletePIDFile()
		Exit
	EndIf
	
	; �� ���������� ��� ������, ��� �� ������� ������� ���������.
	; ������� �������� ��������� ���������� ������������� ��� ������� ���������� (� ����������� ������)
	; ��� ���� ������� �������� ���������� ������ ApplyCfg
	If $bApplyCfg Then

		UpdateProgress(60, "�� ������������� ��� ����������." & @CRLF & "������� ���������� ����������")
		AddToLog("�� ������������� ��� ����������. ����������� ������� ���������� ����������")
		RunApplyCfg()
		DeletePIDFile()
		Exit
		
	EndIf
	
	If $bUpdateRS Then
		UpdateProgress(60, "���������� �� �������������")
		AddToLog("���������� ��������� �����������������������������������������������������������������")
		$connRetail.������������.�����������������������������������������������������������������()
	EndIf
	$NodeList	=	$connRetail.�����������.����������.�������()
	$ThisNode	=	$connRetail.�����������.����������.��������()
	UpdateProgress(70, "���������� ��������� ������")
	AddToLog("���������� ��������� ������. ���� ����: " & $ThisNode.Description & " (" & $ThisNode.Code & ")")
	while ($NodeList.Next())
	
		if ($NodeList.Code <> $ThisNode.Code) Then
		
			AddToLog("����� ��������� ����������������������������������������������. ���� " & $NodeList.Description & " (" & $NodeList.Code & ")")
			$connRetail.������������������.����������������������������������������������(False, $NodeList.Ref)
			
		EndIf
		
	WEnd
	
	UpdateProgress(90, "����� ������� ��������")
	AddToLog("������������ ������� ������, ���������� ��� v8ComConnector")
	While IsObj($connRetail) OR IsObj($v8ComConnector)
		$connRetail	= 0
		$v8ComConnector = 0
		Sleep(500)
		AddToLog("�������� ������������ ������...")
	WEnd
	
EndFunc

; �������� ���������
; *****************************************************************************

; ���������� ����, ���� ������ ���������� ���������� �������
Func CanIContinue()
	$mResult	=	False
	$iLastPID	=	GetLastPID()
	
	If ($iLastPID > 0) Then
		If (ProcessExists($iLastPID)) Then
			AddToLog("������� " & String($iLastPID) & " ��������")
			$mResult	= False
		Else
			DeletePIDFile()
			$mResult	= True
		EndIf
	Else
		$mResult = True
	EndIf
	
	Return $mResult
	
EndFunc

ReadParamsFromIni()
If $bShowProgress Then
	CreateProgressForm()
EndIf
AddToLog("=> ����� ���������")

If ( CanIContinue() ) Then
	
	CreatePIDFile()	; 20 %
	RunExchange()	; 90 %
	DeletePIDFile()	; 100%
	AddToLog("<= ��������� ���������");
	
else

	UpdateProgress(100, "������ �� ����� ���������� ������." & @CRLF & "������� CanIContinue �� �����������")
	AddToLog("������ �� ����� ���������� ������. ������� CanIContinue �� �����������");
	AddToLog("<= ������ ��������� ������");
	
EndIf