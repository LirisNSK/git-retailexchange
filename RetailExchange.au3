 ; ****************************************************************************
 ;
 ;	Обмен данными для узла 1С:Розница 2
 ;	© 2013-2015 Liris 
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

Global	$sLogfileName	; Имя log-файла
Global	$sRetailIBConn	; Строка подключения к ИБ
Global	$sV8exePath		; Путь к исполняемому файлу 1cv8.exe
Global	$sArcLogfileName; Имя архива log-файла
Global	$iLogMaxSize	; Максимальный размер log-файла
Global	$cPIDFileExt	; Расширение pid-файла
Global	$bShowProgress	; Показать/не показывать окно прогресса
Global	$bShowTrayTip	; Показать/не показывать сообщение в трее
Global	$bUpdateRS		; Обновлять регистр сведений ИнформативныеОстаткиТоваровДляМагазинов
Global	$bApplyCfg		; Применить изменения конфигурации
Global	$sComConnectorObj	; Имя COM-объекта 
Global	$v8ComConnector, $connRetail

; Переменные для оконного интерфейса
Global	$ProgressBar, $LabText

; Чтение настроек из ini-файла и разбора параметров командной строки
Func ReadParamsFromIni()
	
	$sIniFileName	= StringRegExpReplace(@ScriptFullPath, '^(?:.*\\)([^\\]*?)(?:\.[^.]+)?$', '\1')
	$sIniFileName	= @ScriptDir & "\" & $sIniFileName & ".ini"
	; Читает из INI-файла параметр 'Key' в секции 'EXCHANGE'.
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

	; Чтение параметров командной строки
	; Параметры, переданные в командной строке имеют приоритет над параметрами ini-файла
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

; Функции для работы с лог файлом
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
		
			;	Формируем строку для замещения
			$sDT	= @YEAR & "_" & @MON & "_" & @MDAY
			;	Замещаем строку в имени файла
			$sNewFileName	=	StringReplace($sArcLogfileName,"DDMMYYYY", $sDT)
			$sNewFileName	=	$mCurrFolder & "\" & $sNewFileName
			;	Перемещаем файл в архив
			FileMove($mLogfilePath, $sNewFileName)
			;	Создаем новый файл лога
			$fFileLog		=	FileOpen($mLogfilePath, 1)
			$sWriteToFile	=	GetTimestampString() & "Предыдущий лог перемещен в архив: " & $sNewFileName
			FileWriteLine($fFileLog, $sWriteToFile)
			FileClose($fFileLog)
		EndIf

	Else
		$fFileLog		=	FileOpen($mLogfilePath, 1)
		$sWriteToFile	=	GetTimestampString() & "Сформирован новый лог-файл "
		FileWriteLine($fFileLog, $sWriteToFile)
		FileClose($fFileLog)
	EndIf

EndFunc

; *****************************************************************************

; Функции для работы с процессами
; *****************************************************************************

; Возвращает PID работающего процесса. Если в текущий момент никакого процесса не запущено, возвращает 0
Func GetLastPID()
	
	$mReturn		=	0;
	$mCurrFolder	=	@ScriptDir;
	$hSearch		=	FileFindFirstFile($mCurrFolder & "\*." & $cPIDFileExt)
	; Проверка, является ли поиск успешным
	If $hSearch = -1 Then
		; Ошибка при поиске файлов
		; Нужно вернуть 0
		Return 0
		Exit
	EndIf	
	
	While 1
		$sFile = FileFindNextFile($hSearch) ; возвращает имя следующего файла, начиная от первого до последнего
		If @error Then ExitLoop
		
		AddToLog("Найден PID-файл: " & $sFile)
		$sPID	=	StringReplace($sFile, "." & $cPIDFileExt, "")
		;AddToLog("Идентификатор процесса: " & $sPID)
		$mReturn=	Int($sPID)
	WEnd
	
	FileClose($hSearch)
	
	return $mReturn
	
EndFunc

; Удаляет все PID-файлы в папке скрипта
Func DeletePIDFile()
	
	UpdateProgress(100, "Процедура обмена завершает работу")
	AddToLog("Попытка удалить PID-файлы в папке: " & @ScriptDir)
	$mCurrFolder	=	@ScriptDir;
	$iResult		=	FileDelete($mCurrFolder & "\*." & $cPIDFileExt)
	If $iResult > 0 Then
		AddToLog("PID-файлы удалены успешно")
	Else
		AddToLog("При удалении файлов произошла ошибка, либо нет файлов для удаления")
	EndIf
	
EndFunc

; Создает новый PID-файл
Func CreatePIDFile()
	
	$mCurrFolder	=	@ScriptDir;
	$iCurrentPID	=	@AutoItPID;
	$sPIDFileName	=	$mCurrFolder & "\" & $iCurrentPID & "." & $cPIDFileExt;
	
	$fFileOut		=	FileOpen($sPIDFileName, 1)
	FileWrite($fFileOut, String($iCurrentPID))
	FileClose($fFileOut)
	
	$sMsg = "Создан новый PID-файл: " & String($iCurrentPID)
	AddToLog($sMsg)
	UpdateProgress(20, $sMsg)
	
EndFunc	

; Функции работы с окнами
; *****************************************************************************

Func CreateProgressForm()

	$_DH	= Ceiling(@DesktopHeight / 20)
	$_DW	= Ceiling(@DesktopWidth / 30)

	GUICreate("Прогресс операции", 12*$_DW, 6*$_DH)
	$ProgressBar	= GUICtrlCreateProgress(2*$_DW, $_DH, 8*$_DW, 2*$_DH)
	$LabText		= GUICtrlCreateLabel("", 2*$_DW, 4*$_DH, 8*$_DW, 2*$_DH, $SS_CENTER)
	GUICtrlSetFont($LabText, 20)
	GUISetState()
	Opt("GUICloseOnESC", 0)
	
EndFunc

Func UpdateProgress($iProgress, $sStatusText)
	
	; Проверка существования объектов
	If GUICtrlGetHandle($Progressbar) <> 0 Then
		; Изменение положения прогрессбара
		GUICtrlSetData($Progressbar, $iProgress)
	EndIf
	
	If GUICtrlGetHandle($LabText) <> 0 Then
		; Обновление статусной информации
		GUICtrlSetData($LabText, $sStatusText)
	EndIf
	
	If $bShowTrayTip Then
		ToolTip($sStatusText, @DesktopWidth - 260, @DesktopHeight - 90, "Обмен данными", 1)
	EndIf
	
EndFunc

; *****************************************************************************

; Разбирает строку подключения на составляющие
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

; Функция пытается применить изменения 
Func RunUpdateCfg()

	;v8exe & " DESIGNER /F" & $sIBPath  & " /N" & IBAdminName & " /P" & IBAdminPwd & " /WA- /UpdateDBCfg /Out" & $ServiceFileName & " -NoTruncate /DisableStartupMessages"
	Local $aConnParams[3]
	Local $sUpdCmdLine, $sRunClientCmdLine
	Local $sIBPath, $sIBAdmin, $sIBAdminPwd, $ServiceFileName

	; Для начала нужно отключиться от базы данных
	While IsObj($connRetail) OR IsObj($v8ComConnector)
		
		$connRetail	= 0
		$v8ComConnector = 0
		UpdateProgress(50, "Ожидание освобождения памяти," & @CRLF & " занятой COM-объектами")
		Sleep(500)
		AddToLog("Ожидание освобождения памяти от объектов")

	WEnd
	
	$sQuestion	=	"Уважаемый пользователь!" & @CRLF & "Получены изменения конфигурации базы данных." & @CRLF
	$sQuestion	=	$sQuestion & "Пожалуйста, закройте РМК и нажмите кнопку ОК" & @CRLF & "Чтобы откааться от принятия изменений нажмите ОТМЕНА"
	$bQResult	=	MsgBox(1 + 32, "Требуется обновление программы", $sQuestion, 60)	
	
	If $bQResult = 1 Then
	
		$sServiceFileName	=	@ScriptDir & "\1c_update.txt" 
		; Получаю параметры подключения из строки подключения
		; !Опробовано только на файловом варианте ИБ
		$aConnParams= SplitConnectionString($sRetailIBConn)
		$sIBPath	=	$aConnParams[0]
		$sIBAdmin	=	$aConnParams[1]
		$sIBAdminPwd=	$aConnParams[2]
		
		; Сюда добавить обработки, которые выгоняют пользователей
		
		$sUpdCmdLine	=	$sV8exePath & " DESIGNER /F" & $sIBPath  & " /N" & $sIBAdminPwd & " /P" & $sIBAdminPwd 
		$sUpdCmdLine	=	$sUpdCmdLine & " /WA- /UpdateDBCfg /Out""" & $sServiceFileName & """ -NoTruncate /DisableStartupMessages"
		; Принятие изменений
		UpdateProgress(50, "Выполняется принятие изменений")
		AddToLog("Выполняется принятие изменений")
		$PIDUpdCfg	= Run($sUpdCmdLine)
		ProcessWaitClose($PIDUpdCfg)
		UpdateProgress(50, "Команда на принятие" & @CRLF & "изменений выполнена")
		AddToLog("Команда на принятие изменений выполнена")
		
	Else
		
		; Таймаут или пользователь отказался от применения обновления
		UpdateProgress(50, "Отказ принятия изменений")
		AddToLog("Отмена принятия изменений. Результат диалога: " & String($bQResult) )
		
	EndIf
		
EndFunc

Func RunApplyCfg()

	;v8exe & " DESIGNER /F" & $sIBPath  & " /N" & IBAdminName & " /P" & IBAdminPwd & " /WA- /UpdateDBCfg /Out" & $ServiceFileName & " -NoTruncate /DisableStartupMessages"
	Local $aConnParams[3]
	Local $sUpdCmdLine, $sRunClientCmdLine
	Local $sIBPath, $sIBAdmin, $sIBAdminPwd, $ServiceFileName
	
	; Для начала нужно отключиться от базы данных
	While IsObj($connRetail) OR IsObj($v8ComConnector)
		$connRetail	= 0
		$v8ComConnector = 0
		Sleep(500)
		AddToLog("Ожидание освобождения памяти от объектов")
	WEnd
	
	$sQuestion	=	"Уважаемый пользователь!" & @CRLF & "База данных заблокирована для обновления" & @CRLF
	$sQuestion	=	$sQuestion & "Пожалуйста, закройте РМК и нажмите кнопку ОК" & @CRLF & "Чтобы отложить это действие нажмите ОТМЕНА"
	$bQResult	=	MsgBox(1 + 32, "ИБ заблокирована для обновления", $sQuestion, 60)
	
	If $bQResult = 1 Then
	
		; Получаю параметры подключения из строки подключения
		; Опробовано только на файловом варианте базы данных
		$aConnParams= SplitConnectionString($sRetailIBConn)
		$sIBPath	=	$aConnParams[0]
		$sIBAdmin	=	$aConnParams[1]
		$sIBAdminPwd=	$aConnParams[2]
		
		; Запуск клиента для применения изменений
		$sRunClientCmdLine	=	$sV8exePath & " ENTERPRISE /F""" & $sIBPath  & """ /N" & $sIBAdminPwd & " /P" & $sIBAdminPwd 
		$sRunClientCmdLine	=	$sRunClientCmdLine & " /WA-"
		UpdateProgress(60, "Выполняется запуск приложения" & @CRLF & "для принятия изменений")
		AddToLog("Выполняется запуск клиентского приложения")
		$PIDApplyCfg		=	Run($sRunClientCmdLine)
		ProcessWaitClose($PIDApplyCfg)
		UpdateProgress(60, "Применение изменений" & @CRLF & "выполнено успешно")
		AddToLog("Обновление базы данных выполнено")
		
	Else
		
		; Таймаут или пользователь отказался от принятия обновления
		UpdateProgress(60, "Отмена принятия изменений")
		AddToLog("Отмена принятия изменений. Результат диалога: " & String($bQResult) )
		
	EndIf
		
EndFunc
 
; Функция выполняет процедуры обмена данными
Func RunExchange()

	$v8ComConnector = ObjCreate($sComConnectorObj)
	If @error Then
		AddToLog("Ошибка при создании COM-объекта " & $sComConnectorObj)
		UpdateProgress(20, "Ошибка при создании COM-объекта " & $sComConnectorObj)
		Exit
	EndIf
	
	AddToLog("Попытка подключения к ИБ");
	UpdateProgress(25, "Попытка подключения к ИБ")
	
	$connRetail	=	$v8ComConnector.Connect($sRetailIBConn);

	If Not IsObj($connRetail) Then
		
		UpdateProgress(30, "При подключении к ИБ произошла ошибка")
		AddToLog("При подключении к ИБ произошла ошибка")
		AddToLog("Скрипт завершает работу из-за ошибки соединения с ИБ")
		; Может вставить отправку alarm'a? Например на SMS или email
		Exit
		
	Else
		AddToLog("Подключение к ИБ установлено")
		UpdateProgress(30, "Подключились к ИБ")
	EndIf
	
	If $connRetail.КонфигурацияИзменена() Then
		UpdateProgress(40, "Конфигурация ИБ изменена."& @CRLF & "Попытка применить изменения")
		AddToLog("Конфигурация ИБ изменена. Требуется применение изменений")
		RunUpdateCfg()
		DeletePIDFile()
		Exit
	EndIf
	
	; Не разобрался как узнать, что ИБ требует принять изменения.
	; Поэтому принятие изменений выполнится автоматически при запуске приложения (в управляемом режиме)
	; или если передан параметр коммандной строки ApplyCfg
	If $bApplyCfg Then

		UpdateProgress(60, "ИБ заблокирована для обновления." & @CRLF & "Попытка применения обновления")
		AddToLog("ИБ заблокирована для обновления. Выполняется попытка применения обновления")
		RunApplyCfg()
		DeletePIDFile()
		Exit
		
	EndIf
	
	If $bUpdateRS Then
		UpdateProgress(60, "Обновление РС ИнфОстТоваров")
		AddToLog("Выполнение процедуры ОбновлениеРегистраСведенийИнформативныеОстаткиТоваровДляМагазинов")
		$connRetail.ЗапасыСервер.ОбновлениеРегистраСведенийИнформативныеОстаткиТоваровДляМагазинов()
	EndIf
	$NodeList	=	$connRetail.ПланыОбмена.ПоМагазину.Выбрать()
	$ThisNode	=	$connRetail.ПланыОбмена.ПоМагазину.ЭтотУзел()
	UpdateProgress(70, "Выполнение процедуры обмена")
	AddToLog("Выполнение процедуры обмена. Этот узел: " & $ThisNode.Description & " (" & $ThisNode.Code & ")")
	while ($NodeList.Next())
	
		if ($NodeList.Code <> $ThisNode.Code) Then
		
			AddToLog("Вызов процедуры ВыполнитьОбменДаннымиДляУзлаИнформационнойБазы. Узел " & $NodeList.Description & " (" & $NodeList.Code & ")")
			$connRetail.ОбменДаннымиСервер.ВыполнитьОбменДаннымиДляУзлаИнформационнойБазы(False, $NodeList.Ref)
			
		EndIf
		
	WEnd
	
	UpdateProgress(90, "Обмен данными завершен")
	AddToLog("Освобождение области памяти, отведенной под v8ComConnector")
	While IsObj($connRetail) OR IsObj($v8ComConnector)
		$connRetail	= 0
		$v8ComConnector = 0
		Sleep(500)
		AddToLog("Ожидание освобождения памяти...")
	WEnd
	
EndFunc

; Основная программа
; *****************************************************************************

; Возвращает ЛОЖЬ, если нельзя продолжать выполнение скрипта
Func CanIContinue()
	$mResult	=	False
	$iLastPID	=	GetLastPID()
	
	If ($iLastPID > 0) Then
		If (ProcessExists($iLastPID)) Then
			AddToLog("Процесс " & String($iLastPID) & " работает")
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
AddToLog("=> Старт обработки")

If ( CanIContinue() ) Then
	
	CreatePIDFile()	; 20 %
	RunExchange()	; 90 %
	DeletePIDFile()	; 100%
	AddToLog("<= Обработка завершена");
	
else

	UpdateProgress(100, "Скрипт не может продолжить работу." & @CRLF & "Условие CanIContinue не выполнилось")
	AddToLog("Скрипт не может продолжить работу. Условие CanIContinue не выполнилось");
	AddToLog("<= Скрипт завершает работу");
	
EndIf