:: Workaround for vbp file cosistency issue when saving
::
:: Detail: In the *.vbp project file the "Register=*..." lines use relative
::         paths if the project is opened on the same drive as registered
::         dlls/ocx files. This is problematic as the checkout folder might be
::         at a different location for every developer. To avoid this issue this
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
IF [%CD:~0,2%]==[Z:] GOTO:DRIVE_ERROR

PUSHD %~dp0
:: remove previous Z:
SUBST Z: /d
:: register current folder as new Z:
SUBST Z: .
:: run vb6 as normal user with our project
START Z:\prjSoldatMapEditor.vbp
POPD
GOTO:END

:: functions
:DRIVE_ERROR
ECHO ERROR: Please run the script from real project location instead of Z:
PAUSE>NUL

:END
