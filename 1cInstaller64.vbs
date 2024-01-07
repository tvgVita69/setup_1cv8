Option Explicit
Const msiUILevelNoChange = 0        'Не изменять интерфейс пользователя
Const msiUILevelDefault = 1         'Использовать интерфейс пользователя, заданный по умолчанию
Const msiUILevelNone = 2            'Не отображать интерфейс пользователя (молчаливая установка)
Const msiUILevelBasic = 3           'Только индикатор прогресса и отображение ошибок
Const msiUILevelReduced = 4         'Интерфейс пользователя без диалоговых сообщений
Const msiUILevelFull = 5            'Полный интерфейс пользователя
Const msiUILevelHideCancel = 32     'Если используется с msiUILevelBasic, то отображается индикатор прогресса без кнопки Cancel
Const msiUILevelProgressOnly = 64   'Если используется с msiUILevelBasic, то отображается индикатор прогресса без тображения любых диалогов, в т.ч. и ошибок.
Const msiUILevelEndDialog = 128     'Если используется с любым из перечисленных значений, интсталятор выводит сообщение в конце установки о конечном результате.

Const DistrFolder = "C:\Program Files\1cv8" rem //укажем путь для установки
Const shortcutName = "1C Предприятия"
Dim shortcutTarget : shortcutTarget = DistrFolder & "1cestart.exe"
Const requiredInstall = 1
Const requiredUninstall = 0
Const InstallUID= "{B81E11C4-D21B-46BC-BF35-E1799439FF9F}"  rem //это значение нужно взять из дистрибутива из файла setup.ini - ProductCode
installOrUninstall InstallUID, DistrFolder + "8.3.15.1747\1CEnterprise 8 (x86-64).msi", "1049.mst", "adminstallrestart.mst", requiredInstall   rem //8.2.16.368 укажем название папки, в которую будет установлена платформа
Sub installOrUninstall (ByVal productCode, ByVal msiPackage, ByVal mstTransform, ByVal mstinstall, ByVal requiredAction)
productCode = "{B81E11C4-D21B-46BC-BF35-E1799439FF9F}"  rem //это значение нужно взять из дистрибутива из файла setup.ini - ProductCode
msiPackage = "\\192.168.17.7\netlogon\1c_new_platphorma\64.8.3.15.1747\1CEnterprise 8 (x86-64).msi"   rem //полный путь к дистрибутиву, а именно к файлу 1CEnterprise 8.2.msi

Dim cmdLine
On Error Resume Next
Dim installer, session
Set installer = Nothing
Set session = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : processError
installer.UILevel = msiUILevelBasic 'msiUILevelNone
Set session = installer.OpenProduct(productCode)
If session Is Nothing AND requiredAction = requiredInstall Then

cmdLine = "TRANSFORMS=adminstallrestart.mst; "
If Not mstTransform Is Empty Then

cmdLine = cmdLine & mstTransform
rem //указываем какие компоненты устанавливаем
cmdLine = cmdLine &  "THICKCLIENT=1 THINCLIENT=1 WEBSERVEREXT=1 SERVER=0 CONFREPOSSERVER=0 CONVERTER77=0 SERVERCLIENT=0 LANGUAGES=RU"
End If
rem //установка платформы
'MsgBox("Устанавливаем платформу" & "cmd=" & cmdLine & " msi=" & msiPackage)
Set session = installer.InstallProduct(msiPackage, cmdLine) : processError

createShurtcut()
ElseIf Not session Is Nothing AND requiredAction = requiredUninstall Then


Set session = Nothing
cmdLine = "REMOVE=ALL"

Set session = installer.InstallProduct(msiPackage, cmdLine) : processError
End If
Set session = Nothing
Set installer = Nothing
End Sub

Sub processError
Dim msg
If Err = 0 Then Exit Sub
msg = Str(Err.Number) & Err.Source & " " & Hex(Err) & ": " & Err.Description
Wscript.Echo msg
MsgBox(msg)
End Sub
'создание ярлыка
Sub createShurtcut
Dim WshShell, oShellLink
Set WshShell = WScript.CreateObject("WScript.Shell")
Dim strDesktop : strDesktop = WshShell.SpecialFolders("Desktop")
Set oShellLink = WshShell.CreateShortcut(strDesktop & "\" & shortcutName & ".lnk")
oShellLink.TargetPath = shortcutTarget
oShellLink.WindowStyle = 1
oShellLink.Description = shortcutName
oShellLink.Save
Set oShellLink = Nothing
Set WshShell = Nothing
End Sub