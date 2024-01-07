@echo off
rem set _1Clog=%temp%\install1c.log
rem echo run 1C installation > %_1Clog%
rem Путь к дистрибутиву 1С (без завершающего обратного слеша):
rem set _1Cpath=\\192.168.17.7\netlogon\1c_new_platphorma\windows8_3_15_1747
rem подключаем версию 1С:
rem call %_1Cpath%\1cver.bat
rem echo 1cversion=%_1Cversion% >> %_1Clog%

rem goto %PROCESSOR_ARCHITECTURE%
rem goto exit
rem :x86
rem IF EXIST "%ProgramFiles%\1cv8\8.3.15.1747\" goto installed
rem goto not_installed
rem :AMD64
rem IF EXIST "%ProgramFiles(x86)%\1cv8\8.3.15.1747\" goto installed
rem goto not_installed

rem :installed
rem echo installed >> %_1Clog%
rem Удаление 1С 8.3.8.2167:
rem start /wait msiexec /passive /uninstall {3A97FC22-6940-4119-B3DE-323EDB63721D}
rem start /wait msiexec /qb /i "\\192.168.17.7\netlogon\1c_new_platphorma\windows8_3_15_1747\1CEnterprise 8 (x86-64).msi" /qn TRANSFORMS=adminstallrelogon.mst;1049.mst DESIGNERALLCLIENTS=1 THINCLIENTFILE=1 THINCLIENT=1 WEBSERVEREXT=1 SERVER=0 CONFREPOSSERVER=0 CONVERTER77=0 SERVERCLIENT=0 LANGUAGES=RU
rem start /wait msiexec /qn "\\192.168.17.7\netlogon\1c_new_platphorma\windows8_3_15_1747\1CEnterprise 8 (x86-64).msi"  TRANSFORMS=adminstallrelogon.mst;1049.mst DESIGNERALLCLIENTS=1 THICKCLIENT=1 THINCLIENTFILE=1 THINCLIENT=1 WEBSERVEREXT=0 SERVER=0 CONFREPOSSERVER=0 
msiexec /i "\\192.168.17.7\netlogon\1c_new_platphorma\windows8_3_15_1747\1CEnterprise 8 (x86-64).msi" /quiet /passive TRANSFORMS=adminstallrelogon.mst;1049.mst DESIGNERALLCLIENTS=1 THINCLIENT=1 THINCLIENTFILE=1 SERVER=0 WEBSERVEREXT=1 CONFREPOSSERVER=0 SERVERCLIENT=0 CONVERTER77=0 LANGUAGES=RU

rem goto exit

rem :not_installed
rem echo not installed >> %_1Clog%
rem echo installing... >> %_1Clog%
rem start  msiexec "\\192.168.17.7\netlogon\1c_new_platphorma\windows8_3_15_1747\1CEnterprise 8 (x86-64).msi" /qr TRANSFORMS=adminstallrelogon.mst;1049.mst DESIGNERALLCLIENTS=1 THICKCLIENT=1 THINCLIENTFILE=1 THINCLIENT=1 WEBSERVEREXT=0 SERVER=0 CONFREPOSSERVER=0 CONVERTER77=0 SERVERCLIENT=0 LANGUAGES=RU

rem echo msiexec /qb /i "\\192.168.17.7\netlogon\1c_new_platphorma\windows8_3_15_1747\1CEnterprise 8 (x86-64).msi" TRANSFORMS=adminstallrelogon.mst;1049.mst DESIGNERALLCLIENTS=1 THINCLIENTFILE=1 THINCLIENT=1 WEBSERVEREXT=1 SERVER=0 CONFREPOSSERVER=0 CONVERTER77=0 SERVERCLIENT=0 LANGUAGES=RU >> %_1Clog%

rem start /wait msiexec /qb /i "\\192.168.17.7\netlogon\1c_new_platphorma\windows8_3_15_1747\1CEnterprise 8 (x86-64).msi" TRANSFORMS=adminstallrelogon.mst;1049.mst DESIGNERALLCLIENTS=1 THINCLIENTFILE=1 THINCLIENT=1 WEBSERVEREXT=1 SERVER=0 CONFREPOSSERVER=0 CONVERTER77=0 SERVERCLIENT=0 LANGUAGES=RU

rem Здесь может быть патч 1С: например, копирование файла патченной DLL в папку с установленной 1С
rem ...
rem goto exit

rem :exit
rem echo exiting >> %_1Clog%
rem echo. >> %_1Clog%
pause  