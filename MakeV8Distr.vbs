' Utility to change 1C:Enterprise Windows Installer database
' to use package on group polices in Active directory domains
' Version: 1.5

Const msiOpenDatabaseModeTransact = 1
Const msiOpenDatabaseModeDirect = 2
Const msiTransformErrorAddExistingRow = &h0001
Const msiTransformErrorDeleteNonExistingRow = &h0002
Const msiTransformErrorAddExistingTable = &h0004
Const msiTransformErrorDeleteNonExistingTable = &h0008
Const msiTransformErrorUpdateNonExistingRow = &h0010
Const msiTransformErrorChangeCodePage = &h0020
Const msiTransformErrorViewTransform = &h0100
Const msiViewModifyUpdate = 2
Const SummaryProperty_PID_COMMENTS = 6

Dim argCount:argCount = Wscript.Arguments.Count
If argCount > 0 Then If InStr(1, Wscript.Arguments(0), "?", vbTextCompare) > 0 Then argCount = 0
If argCount = 0 Then
	Wscript.Echo "Utility to change 1C:Enterprise Windows Installer database." &_
		vbLf & " Syntax: cscript.exe MakeV8Distr.vbs <1CV8 installer database> [/l<Language>] [/t<transform file>] [/h<On|Off>] [/v<On|Off>] [/d<On|Off>] [/c<On|Off>] [/e<On|Off>]" &_
		vbLf & "  <1CV8 installer database> - path to V8 MSI database (installer package);" &_
		vbLf & "  /l - turn off all additional interface accept specifed;" &_
		vbLf & "   <Language> - short language interface name or the list of names through a comma, AUTO or it is not specified - it is defined by language of the user at installation;" &_
		vbLf & "  /t - apply specifed transform file to database;" &_
		vbLf & "   <transform file> - path to localization transform file;" &_
		vbLf & "  /h - turn off or on to install HASP driver;" &_
		vbLf & "  /v - turn off or on to install V7 database converter;" &_
		vbLf & "  /d - turn off or on to install configuration repository server;" &_
		vbLf & "  /c - turn off or on to install server client;" &_
		vbLf & "  /e - turn off or on to install additional shortcut on desktop." &_
		vbLf & " Example: cscript.exe MakeV8Distr.vbs ""1CEnterprise 8.1 for clients.msi"" /lRU /t1049.mst /hOff /vOff /dOff /eOn" &_
		vbLf & "" &_
		vbNewLine & "Copyright (c) IT - otdel mailto:it@igb.ru tel:22222222"
	Wscript.Quit 1
End If

Dim installer, database
Dim DatabasePath:DatabasePath = Wscript.Arguments(0)
Dim nArg:nArg = 1
Dim  Language, Transform, HaspDrv, Converter, Repository, ServerClient, ShortcutOnDesktop
Do While nArg < argCount
    Select Case UCase(Left(Wscript.Arguments(nArg), 2))
        Case "/L"
            Language = UCase(Mid(Wscript.Arguments(nArg), 3))
            If Language=Empty Then
                Language = "AUTO"
            End If
        Case "/T"
            Transform = Mid(Wscript.Arguments(nArg), 3)
        Case "/H"
            Select Case UCase(Mid(Wscript.Arguments(nArg), 3))
                Case "ON"
                    HaspDrv = "yes"
                Case "OFF"
                    HaspDrv = "no"
                Case ""
                    HaspDrv = Empty
                Case Else
                    Wscript.Echo "Wrong option value"
                    Wscript.Quit 3
            End Select
        Case "/V"
            Converter = GetFutureLevel(Mid(Wscript.Arguments(nArg), 3))
        Case "/D"
            Repository = GetFutureLevel(Mid(Wscript.Arguments(nArg), 3))
        Case "/C"
            ServerClient = GetFutureLevel(Mid(Wscript.Arguments(nArg), 3))
        Case "/E"
            ShortcutOnDesktop = GetFutureLevel(Mid(Wscript.Arguments(nArg), 3))
        Case Else
            Wscript.Echo "Wrong option"
            Wscript.Quit 3
    End Select
    nArg = nArg + 1
Loop

Set FSO = Wscript.CreateObject("Scripting.FileSystemObject") : CheckError
If Not FSO.FileExists(DatabasePath) Then
	Wscript.Echo  "Error: """ & DatabasePath & """ file is not exists."
	Wscript.Quit 4
End If

Set file = FSO.GetFile(DatabasePath) : CheckError
If file.Attributes And 1 Then
	Wscript.Echo  "Error: """ & DatabasePath & """ is read only."
	Wscript.Quit 4
End If

Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError
Set database = installer.OpenDatabase(DatabasePath, msiOpenDatabaseModeTransact) : CheckError
Dim view, recordData

If Transform<>Empty Then
    database.ApplyTransform Transform, msiTransformErrorDeleteNonExistingRow : CheckError
    Dim recordName
    Set view = database.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = 'ProductNameLocal'")
    view.Execute : CheckError
    Set recordName = view.Fetch
    If Not recordName Is Nothing Then
        UpdateQuery database, "UPDATE `Property` SET `Property`.`Value` = ? WHERE `Property` = 'ProductName'", recordName
    End If
End If

If Language<>Empty Then
    If Language="AUTO" Then
	    Set view = database.OpenView("UPDATE `Feature` SET `Feature`.`Level` = 1 WHERE `Feature_Parent` = 'Languages'") : CheckError
	    view.Execute : CheckError
	    
        Set view = database.OpenView("SELECT `Action` FROM `InstallExecuteSequence` WHERE `Action` = 'customPresetDefLang'") : CheckError
        view.Execute : CheckError
        Set recordData = view.Fetch
        If recordData Is Nothing Then
	        Set view = database.OpenView("INSERT INTO `InstallExecuteSequence` (`Action`, `Condition`, `Sequence`) VALUES ('customPresetDefLang', 'Not Installed And Not PATCH', 701)") : CheckError
	        view.Execute : CheckError
	    End If
	    
        Set view = database.OpenView("SELECT `Action` FROM `InstallUISequence` WHERE `Action` = 'customPresetDefLang'") : CheckError
        view.Execute : CheckError
        Set recordData = view.Fetch
        If recordData Is Nothing Then
	        Set view = database.OpenView("INSERT INTO `InstallUISequence` (`Action`, `Condition`, `Sequence`) VALUES ('customPresetDefLang', 'Not Installed And Not PATCH', 643)") : CheckError
	        view.Execute : CheckError
	    End If
    Else
	    Set view = database.OpenView("UPDATE `Feature` SET `Feature`.`Level` = 200 WHERE `Feature_Parent` = 'Languages'") : CheckError
	    view.Execute : CheckError

	    Dim Languages, paramLanguage
        Languages = Split(Language, ",")
	    Set paramLanguage = installer.CreateRecord(1)
	    Set view = database.OpenView("UPDATE `Feature` SET `Feature`.`Level` = 1 WHERE `Feature_Parent` = 'Languages' AND `Feature` = ?") : CheckError
        Dim CurrentLanguage
	    For Each CurrentLanguage in Languages
            paramLanguage.StringData(1) = CurrentLanguage
            view.Execute paramLanguage : CheckError
        Next
        
	    Set view = database.OpenView("DELETE FROM `InstallExecuteSequence` WHERE `Action` = 'customPresetDefLang'") : CheckError
	    view.Execute : CheckError
	    Set view = database.OpenView("DELETE FROM `InstallUISequence` WHERE `Action` = 'customPresetDefLang'") : CheckError
	    view.Execute : CheckError
    End If
End If

If HaspDrv<>Empty Then
	Dim paramHasp
	Set paramHasp = installer.CreateRecord(1)
	paramHasp.StringData(1) = HaspDrv
	UpdateQuery database, "UPDATE `Property` SET `Property`.`Value` = ? WHERE `Property` = 'HASPInstall'", paramHasp
End If

If Converter<>Empty Then
	Dim paramConverter
	Set paramConverter = installer.CreateRecord(1)
	paramConverter.IntegerData(1) = Converter
	UpdateQuery database, "UPDATE `Feature` SET `Feature`.`Level` = ? WHERE `Feature` = 'Convertor'", paramConverter
End If

If Repository<>Empty Then
	Dim paramRepository
	Set paramRepository = installer.CreateRecord(1)
	paramRepository.IntegerData(1) = Repository
	UpdateQuery database, "UPDATE `Feature` SET `Feature`.`Level` = ? WHERE `Feature` = 'DepotServer'", paramRepository
End If

If ServerClient<>Empty Then
	Dim paramServerClient
	Set paramServerClient = installer.CreateRecord(1)
	paramServerClient.IntegerData(1) = ServerClient
	UpdateQuery database, "UPDATE `Feature` SET `Feature`.`Level` = ? WHERE `Feature` = 'CSClient'", paramServerClient
End If

If ShortcutOnDesktop<>Empty Then
    If ShortcutOnDesktop=1 Then
        Dim viewSelect
        Set viewSelect = database.OpenView("SELECT `Shortcut` FROM `Shortcut` WHERE `Shortcut` = 'ShortcutEnterprise'") : CheckError
        viewSelect.Execute : CheckError
        Set recordData = viewSelect.Fetch
        If recordData Is Nothing Then
            Dim viewUpdate
            Set viewSelect = database.OpenView("SELECT `Name`, `Component_`, `Target`, `Arguments`, `Description`, `Icon_`, `IconIndex`, `ShowCmd`, `WkDir` FROM `Shortcut` WHERE `Shortcut` = 'ShortCut_Enterpr'") : CheckError
            viewSelect.Execute : CheckError
            Set recordData = viewSelect.Fetch
            Set viewUpdate = database.OpenView(_
            "INSERT INTO `Shortcut` (`Shortcut`, `Directory_`, `Name`, `Component_`, `Target`, `Arguments`, `Description`, `Icon_`, `IconIndex`, `ShowCmd`, `WkDir`)" &_
            "VALUES ('ShortcutEnterprise', 'DesktopFolder', ?, ?, ?, ?, ?, ?, ?, ?, ?)") : CheckError
            viewUpdate.Execute recordData : CheckError
        End If
    Else
        Set view = database.OpenView("DELETE FROM `Shortcut` WHERE `Shortcut`.`Shortcut` = 'ShortcutEnterprise'") : CheckError
        view.Execute : CheckError
    End If
End If

CorrectError_00097458 database : CheckError
ChangeSummaryInfo database : CheckError
database.Commit : CheckError
Set database = Nothing
Set installer = Nothing
Wscript.Echo "Database succesfuly changed"

Sub CheckError
	Dim message, errRec
	If Err = 0 Then Exit Sub
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbNewLine & errRec.FormatText
	End If
	Wscript.Echo message
	Wscript.Quit 2
End Sub

Sub UpdateQuery(database, QueryText, param)
	Dim view
	Set view = database.OpenView(QueryText) : CheckError
	view.Execute param : CheckError
End Sub

Function GetFutureLevel(ValueOnOff)
    Select Case UCase(ValueOnOff)
        Case "ON"
            GetFutureLevel = 1
        Case "OFF"
            GetFutureLevel = 200
        Case ""
            GetFutureLevel = Empty
        Case Else
            Wscript.Echo "Wrong option value"
            Wscript.Quit 3
    End Select
End Function

Sub CorrectError_00097458(database)
    Dim view, record, ScriptText
    Set view = database.OpenView("SELECT `Target` FROM `CustomAction` WHERE `Action` = 'customDetectPrevVersion'") : CheckError
    view.Execute : CheckError
    Set record = view.Fetch
    ScriptText = record.StringData(1)
    If InStr(ScriptText, "if thatInstalled <> -1 then")>0 Then
        ScriptText = Replace(ScriptText, "if thatInstalled <> -1 then", "if thatInstalled <> -1 And thatInstalled <> 1 then")
        record.StringData(1) = ScriptText
        view.Modify msiViewModifyUpdate, record : CheckError
    End If
End Sub

Sub ChangeSummaryInfo(database)
    Dim SI
    Set SI = database.SummaryInformation(1) : CheckError
    SI.Property(SummaryProperty_PID_COMMENTS) = "1C:Enterprise 8.1 changed by MakeV8Distr.vbs v1.5"
    SI.Persist
End Sub
