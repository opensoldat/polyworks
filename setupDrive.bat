:: Workaround for vbp file cosistency issue when saving
::
:: Map project folder to "Z:" drive (if not already mapped) and
:: open VB6 project (with option to skip if VB6.exe is running)
::
:: Detail: In the *.vbp project file the "Register=*..." lines use relative
::         paths if the project is opened on the same drive as registered
::         dlls/ocx files. This is problematic as the checkout folder might be
::         at a different location for every developer. To avoid the issue this
::         script maps the current folder to the Z: drive letter and runs VB6
::         with the project from Z:\* forcing VB6 to use absolue paths instead.
::
:: Note: Make sure Visual Basic 6 can be used as a normal user.
::       Allow full read/write acces for the vb6-exe folder:
::       "C:\Program Files (x86)\Microsoft Visual Studio\VB98" via the
::       context menu: Properties -> Security -> Edit -> Users (PC\Users) ->
::       Full control -> [x] Allow (tick the checkbox for Allow column) -> OK
::       and verify that opening VB6 as normal user doesn't make it freeze.
::
:: Note: Normally this script doesn't need admin rights (see note above)
::
:: Note: If Register=* paths are still different from the one in the repository
::       when saving the project, for example:
::       "C:\Windows\SysWow64\dx8vb.dll" instead of
::       "C:\Windows\SysWOW64\dx8vb.dll" (Wow isntead of WOW),
::       you can modify the path in the windows registry under:
::       "Computer\HKEY_CLASSES_ROOT\TypeLib\{E1211242-8E94-11D1-8808-00C04FC2C603}\1.0\0\win32"
::       where {E1211242-8E94-11D1-8808-00C04FC2C603} is the Register id.
::       You can find the Register ids in the *.vbp file at the start of the line.

@ECHO OFF
PUSHD %~dp0
SETLOCAL

:: user config
SET "MAPPED_DRIVE=Z:"
SET "PROJECT_FILE=src\prjOpenSoldatMapEditor.vbp"
:: user config

SET "MAPPED_PROJECT=%MAPPED_DRIVE%\%PROJECT_FILE%"
IF EXIST %MAPPED_PROJECT% GOTO:CHECK_RUN_VB

:: remove previous Z:
SUBST %MAPPED_DRIVE% /d > NUL
:: register current folder as new Z:
SUBST %MAPPED_DRIVE% .

IF NOT EXIST %MAPPED_PROJECT% GOTO:DRIVE_ERROR

:CHECK_RUN_VB
IF NOT EXIST "%WINDIR%\System32\TASKLIST.EXE" CALL:REQUIREMENT_ERROR "%WINDIR%\System32\TASKLIST.EXE"
IF NOT EXIST "%WINDIR%\System32\FIND.EXE"     CALL:REQUIREMENT_ERROR "%WINDIR%\System32\FIND.EXE"
IF %ERRORLEVEL% NEQ 0 GOTO:WAIT_QUIT

"%WINDIR%\System32\TASKLIST.EXE" /fi "ImageName eq VB6.exe" /fo csv 2> NUL | "%WINDIR%\System32\FIND.EXE" /I "VB6.exe">NUL
IF NOT "%ERRORLEVEL%"=="0" GOTO:RUN_VB

:ASK_VB_RUN
SET /P "CHOICE=Visual Basic 6 already running. Start another instance[y/N]?"
IF /I "%CHOICE%" == "Y" GOTO:RUN_VB
IF /I "%CHOICE%" == "N" GOTO:END
IF /I "%CHOICE%" == ""  GOTO:END
GOTO:ASK_VB_RUN

:: run vb6 as normal user with our project
:RUN_VB
START %MAPPED_PROJECT%
GOTO:END

:: functions
:DRIVE_ERROR
ECHO.ERROR: Cannot map %MAPPED_DRIVE% drive: %MAPPED_PROJECT% project file not found
GOTO:WAIT_QUIT

:WAIT_QUIT
PAUSE>NULL
GOTO:END

:REQUIREMENT_ERROR
ECHO.ERROR: Cannot find %1
SET "ERRORLEVEL=1"
GOTO:END

:END
POPD
ENDLOCAL
