:: Registers ocx/dll files from pwinstall folder

@ECHO OFF
PUSHD .

CALL:CHECK_IF_RUN_AS_ADMIN

SET REG_SUCCESS=1

IF [%1]==[/v] (
  :: verbose mode
  SET EXTRA_ARGS=
) ELSE IF [%1]==[] (
  :: silent mode
  SET EXTRA_ARGS=/s
) ELSE (
  GOTO OUTPUT_USAGE
)

CALL:REGISTER_ACTIVEX_COMPONENT "%~dp0pwinstall\MBMouse.ocx"  %EXTRA_ARGS%
CALL:REGISTER_ACTIVEX_COMPONENT "%~dp0pwinstall\mscomctl.ocx" %EXTRA_ARGS%
CALL:REGISTER_ACTIVEX_COMPONENT "%~dp0pwinstall\COMDLG32.ocx" %EXTRA_ARGS%
CALL:REGISTER_ACTIVEX_COMPONENT "%~dp0pwinstall\dx8vb.dll"    %EXTRA_ARGS%

IF NOT %REG_SUCCESS% == 1 (
  ECHO.
  ECHO Registration FAILED
) ELSE (
  ECHO.
  ECHO Registration SUCCESSFUL
)

GOTO END


:SHOW_NOADMIN_ERROR
ECHO ERROR: Run as admin!
ECHO ERROR: Registering OCX files requires Admin rights.

GOTO END


:OUTPUT_USAGE
ECHO %1: Registers ocx files from pwinstall folder.
ECHO %1  [/v] <absolute path to ocx/dll file>
ECHO   /v  - verbose mode: shows popup for each registration result
ECHO   /h  - display usage

GOTO END


:: functions
:CHECK_IF_RUN_AS_ADMIN
NET SESSION > NUL 2>&1
IF NOT %errorLevel% == 0 GOTO SHOW_NOADMIN_ERROR

GOTO:EOF


:REGISTER_ACTIVEX_COMPONENT
:: unregister first
%systemroot%\SysWOW64\regsvr32.exe /u %2 %1
:: register
%systemroot%\SysWOW64\regsvr32.exe %2 %1
IF NOT %errorLevel% == 0 (
  ECHO ERROR:   Failed Registering %1 ErrorLevel: %errorLevel%
  SET REG_SUCCESS=0
) ELSE (
  ECHO SUCCESS: Registered %1
)

GOTO:EOF


:END
PAUSE > NUL
POPD
GOTO:EOF
