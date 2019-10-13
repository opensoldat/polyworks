:: Bump batch script
::
:: Updates the version numbers for this project
:: Requires sed to work

@ECHO OFF
PUSHD .

sed >NUL 2>&1
IF %ERRORLEVEL% == 9009 (
  ECHO ERROR: Cannot find sed
  GOTO END
)

IF "%1" == "" GOTO USAGE
SET BUMP_VERSION=%1

ECHO _START_%BUMP_VERSION%_END_ | sed -nb "/^_START_[0-9]\+\(\.[0-9]\+\)*_END_/!{q100}"
IF ERRORLEVEL 1 GOTO USAGE

CD /D "%~dp0"

:: Add more matches here
sed -bi "s/^!define PRODUCT_VERSION \".*\"/!define PRODUCT_VERSION \"%BUMP_VERSION%\"/g" pwinstall/pw.nsi
sed -bi "s/^Soldat Polyworks [0-9]\+\(\.[0-9]\+\)*/Soldat Polyworks %BUMP_VERSION%/g" pwinstall/readme.txt

ECHO DONE!

GOTO END

:USAGE
ECHO %0: Updates the version numbers in the project
ECHO Usage: %0 1.2.3.4

:END
POPD
