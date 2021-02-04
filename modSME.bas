Attribute VB_Name = "modSME"
Option Explicit

Global Const pi As Single = 3.14159265358979  'mmm... pi

Global gfxDir As String

Global appPath As String
Global bgClr As Long
Global lblBackClr As Long
Global lblTextClr As Long
Global txtBackClr As Long
Global txtTextClr As Long
Global frameClr As Long

Global font1 As String, font2 As String

Public Const BUTTON_WIDTH = 64
Public Const BUTTON_HEIGHT = 24

Public Const MENU_WIDTH = 64
Public Const MENU_HEIGHT = 16

Public Const BUTTON_SMALL = 0
Public Const BUTTON_LARGE = 1
Public Const BUTTON_MENU = 2
Public Const BUTTON_TOOL = 3

Public Const BUTTON_X = 48
Public Const BUTTON_Y = 0

Public Const MENU_X = 48
Public Const MENU_Y = 96

Public Const BUTTON_UP = 0
Public Const BUTTON_MOVE = 1
Public Const BUTTON_DOWN = 2

'bitblt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'stretchblit
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
        ByVal dwRop As Long) As Long

'mouse over
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'dragging window
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_NCLBUTTONDOWN = &HA1

'taskbar
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

'get pixel
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

'browse
Private Type BROWSEINFO
    hOwner            As Long
    pidlRoot          As Long
    pszDisplayName    As String
    lpszTitle         As String
    ulFlags           As Long
    lpfn              As Long
    lParam            As Long
    iImage            As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const MAX_PATH = 260

Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" _
        (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" _
        (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)


'registry
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Const STANDARD_RIGHTS_READ As Long = &H20000
Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const SYNCHRONIZE As Long = &H100000

Private Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
        KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
        ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long


'file time
Public Const OFS_MAXPATHNAME = 128
Public Const OF_READWRITE = &H2

Public Type OFSTRUCT
    cBytes      As Byte
    fFixedDisk  As Byte
    nErrCode    As Integer
    Reserved1   As Integer
    Reserved2   As Integer
    szPathName(0 To OFS_MAXPATHNAME - 1) As Byte '0-based
End Type

Public Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type

Public Type SYSTEMTIME
    wYear          As Integer
    wMonth         As Integer
    wDayOfWeek     As Integer
    wDay           As Integer
    wHour          As Integer
    wMinute        As Integer
    wSecond        As Integer
    wMilliseconds  As Integer
End Type


Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, _
        lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long

Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
        lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hFile As Long) As Long

Public Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, _
        ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long

Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, _
        lpLocalFileTime As FILETIME) As Long

Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

'ini file
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" _
        (ByVal sSectionName As String, ByVal sKeyName As String, _
        ByVal lDefault As Long, ByVal sFileName As String) As Long

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" _
        (ByVal sSectionName As String, ByVal sReturnedString As String, _
        ByVal lSize As Long, ByVal sFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal sSectionName As String, ByVal sKeyName As String, ByVal sDefault As String, _
        ByVal sReturnedString As String, ByVal lSize As Long, ByVal sFileName As String) As Long

Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" _
        (ByVal sSectionName As String, ByVal sString As String, ByVal sFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal sSectionName As String, ByVal sKeyName As String, _
        ByVal sString As String, ByVal sFileName As String) As Long

'ShellExecute
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'key mapping
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" _
        (ByVal wCode As Long, ByVal wMapType As Long) As Long

'gdi+
Private Type GUID
   Data1    As Long
   Data2    As Integer
   Data3    As Integer
   Data4(7) As Byte
End Type

Private Type PICTDESC
   Size     As Long
   Type     As Long
   hBmp     As Long
   hpal     As Long
   Reserved As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type ImageCodecInfo
    Clsid As GUID
    FormatID As GUID
    CodecNamePtr As Long
    DllNamePtr As Long
    FormatDescriptionPtr As Long
    FilenameExtensionPtr As Long
    MimeTypePtr As Long
    Flags As Long
    Version As Long
    SigCount As Long
    SigSize As Long
    SigPatternPtr As Long
    SigMaskPtr As Long
End Type

'GDI Functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'GDI+ functions
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal fileName As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hBmp As Long, ByVal hpal As Long, GpBitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal image As Long, Height As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)

'functions for gif loading
Private Declare Function GdipSaveImageToFile Lib "gdiplus.dll" (ByVal image As Long, ByVal fileName As Long, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus.dll" (ByVal fileName As Long, ByRef Bitmap As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus.dll" (ByVal Bitmap As Long, ByRef hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus.dll" (ByRef numEncoders As Long, ByRef Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus.dll" (ByVal numEncoders As Long, ByVal Size As Long, ByRef Encoders As Any) As Long
Private Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long


'GDI and GDI+ constants
Private Const PLANES = 14            'Number of planes
Private Const BITSPIXEL = 12         'Number of bits per pixel
Private Const PATCOPY = &HF00021     '(DWORD) dest = pattern
Private Const PICTYPE_BITMAP = 1     'Bitmap type
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2

Public Sub SelectAllText(tb As TextBox)

    tb.SelStart = 0
    tb.SelLength = Len(tb.Text)

End Sub

Private Function GetEncoderClsid(mimeType As String, pClsid As GUID) As Boolean

    Dim num As Long
    Dim Size As Long
    Dim pImageCodecInfo() As ImageCodecInfo
    Dim j As Long
    Dim buffer As String

    Call GdipGetImageEncodersSize(num, Size)
    If (Size = 0) Then
        GetEncoderClsid = False
        Exit Function
    End If

    ReDim pImageCodecInfo(0 To Size \ Len(pImageCodecInfo(0)) - 1)

    Call GdipGetImageEncoders(num, Size, pImageCodecInfo(0))

    For j = 0 To num - 1

        buffer = Space$(lstrlenW(ByVal pImageCodecInfo(j).MimeTypePtr))

        Call lstrcpyW(ByVal StrPtr(buffer), _
        ByVal pImageCodecInfo(j).MimeTypePtr)

        If (StrComp(buffer, mimeType, vbTextCompare) = 0) Then
            pClsid = pImageCodecInfo(j).Clsid
            Erase pImageCodecInfo

            GetEncoderClsid = True
            Exit Function
        End If
    Next j

    Erase pImageCodecInfo

    GetEncoderClsid = False
End Function

Private Function SaveImageAsPNG(ByVal sFileName As String, ByVal sDestFileName As String) As Boolean

    Dim lBitmap As Long
    Dim hBitmap As Long
    Dim Results As Long
    Dim tPicEncoder As GUID

    If GdipCreateBitmapFromFile(StrPtr(sFileName), lBitmap) = 0 Then
        If GdipCreateHBITMAPFromBitmap(lBitmap, hBitmap, 0) = 0 Then
            If GetEncoderClsid("image/png", tPicEncoder) Then
                SaveImageAsPNG = (GdipSaveImageToFile(lBitmap, StrPtr(sDestFileName), tPicEncoder, ByVal 0) = 0)
            Else
                SaveImageAsPNG = False
            End If
            GdipDisposeImage lBitmap
        End If
    End If

End Function

Public Function GifToPng(ByVal src As String, ByVal dest As String) As Long

    Dim Token As Long

    Token = InitGDIPlus

    If SaveImageAsPNG(src, dest) Then
      GifToPng = -1
    Else
      GifToPng = 5
    End If

    FreeGDIPlus Token

End Function

Public Function GifToBmp(ByVal src As String, ByVal dest As String) As Long

    GifToBmp = GifToPng(src, dest)

End Function

'mouse event
Public Function mouseEvent(ByRef pic As PictureBox, ByVal xVal As Integer, ByVal yVal As Integer, xSrc As Integer, ySrc As Integer, Width As Integer, Height As Integer) As Boolean

    If (xVal < 0) Or (xVal > Width) Or (yVal < 0) Or (yVal > Height) Then 'the MOUSELEAVE pseudo-event
        ReleaseCapture
        BitBlt pic.hDC, 0, 0, Width, Height, frmSoldatMapEditor.picGfx.hDC, xSrc, ySrc, vbSrcCopy
        pic.Refresh
        mouseEvent = True
    ElseIf GetCapture() <> pic.hWnd Then 'the MOUSEENTER pseudo-event
        SetCapture pic.hWnd
        BitBlt pic.hDC, 0, 0, Width, Height, frmSoldatMapEditor.picGfx.hDC, xSrc + Width, ySrc, vbSrcCopy
        pic.Refresh
        mouseEvent = True
    End If

End Function

'mouse event
Public Function mouseEvent2(ByRef pic As PictureBox, ByVal xVal As Integer, ByVal yVal As Integer, ByVal buttonType As Byte, ByVal active As Byte, ByVal action As Byte, Optional exWidth As Integer) As Boolean

    Dim xSrc As Integer, ySrc As Integer
    Dim Width As Integer, Height As Integer

    On Error GoTo ErrorHandler

    If buttonType = BUTTON_SMALL Then
        Width = 16
        Height = 16
        xSrc = 0
        ySrc = Int(pic.Tag) * Height
    ElseIf buttonType = BUTTON_LARGE Then
        Width = BUTTON_WIDTH
        Height = BUTTON_HEIGHT
        xSrc = BUTTON_X
        ySrc = BUTTON_Y + Int(pic.Tag) * Height
    ElseIf buttonType = BUTTON_MENU Then
        Width = MENU_WIDTH
        Height = MENU_HEIGHT
        xSrc = MENU_X
        ySrc = MENU_Y + Int(pic.Index) * Height
    End If

    active = active / 255

    If exWidth = 0 Then exWidth = Width

    If action = BUTTON_UP Or action = BUTTON_DOWN Then
        mouseEvent2 = True
    ElseIf (xVal < 0) Or (xVal > exWidth) Or (yVal < 0) Or (yVal > Height) Then 'the MOUSELEAVE pseudo-event
        ReleaseCapture
        mouseEvent2 = True
        action = BUTTON_UP
    ElseIf GetCapture() <> pic.hWnd Then 'the MOUSEENTER pseudo-event
        SetCapture pic.hWnd
        mouseEvent2 = True
        action = BUTTON_MOVE
    End If

    If mouseEvent2 = True Then
        BitBlt pic.hDC, 0, 0, Width, Height, frmSoldatMapEditor.picButtonGfx.hDC, xSrc + Width * action, ySrc + active * Height, vbSrcCopy
        pic.Refresh
    End If

    Exit Function

ErrorHandler:

    MsgBox Error$

End Function


'browse
Public Function SelectFolder(ownerForm As Form) As String

    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim path As String
    Dim pos As Long

    With bi
        .hOwner = ownerForm.hWnd
        .pidlRoot = 0&
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    pidl = SHBrowseForFolder(bi)
    path = Space$(MAX_PATH)

    If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
        pos = InStr(path, Chr$(0))
        SelectFolder = LCase$(left(path, pos - 1))
    End If

    Call CoTaskMemFree(pidl)

End Function



Public Function snapForm(currentForm As Form, otherForm As Form) As String

    snapForm = ""

    'snap bottom to bottom
    If Abs(currentForm.Top + currentForm.Height - otherForm.Top - otherForm.Height) <= 8 * Screen.TwipsPerPixelY Then
        If (currentForm.left + currentForm.Width + 8 * Screen.TwipsPerPixelX) >= otherForm.left And currentForm.left <= (otherForm.left + otherForm.Width + 8 * Screen.TwipsPerPixelX) Then
            currentForm.Top = otherForm.Top + otherForm.Height - currentForm.Height
            snapForm = "snap"
        End If
    'snap bottom to top
    ElseIf Abs(currentForm.Top + currentForm.Height - otherForm.Top) <= 8 * Screen.TwipsPerPixelY Then
        If (currentForm.left + currentForm.Width + 8 * Screen.TwipsPerPixelX) >= otherForm.left And currentForm.left <= (otherForm.left + otherForm.Width + 8 * Screen.TwipsPerPixelX) Then
            currentForm.Top = otherForm.Top - currentForm.Height + Screen.TwipsPerPixelY
            snapForm = "snap"
        End If
    End If
    'snap right to right
    If Abs(currentForm.left + currentForm.Width - otherForm.left - otherForm.Width) <= 8 * Screen.TwipsPerPixelX Then
        If (currentForm.Top + currentForm.Height + 8 * Screen.TwipsPerPixelY) >= otherForm.Top And currentForm.Top <= (otherForm.Top + otherForm.Height + 8 * Screen.TwipsPerPixelY) Then
            currentForm.left = otherForm.left + otherForm.Width - currentForm.Width
            snapForm = "snap"
        End If
    'snap right to left
    ElseIf Abs(currentForm.left + currentForm.Width - otherForm.left) <= 8 * Screen.TwipsPerPixelX Then
        If (currentForm.Top + currentForm.Height + 8 * Screen.TwipsPerPixelY) >= otherForm.Top And currentForm.Top <= (otherForm.Top + otherForm.Height + 8 * Screen.TwipsPerPixelY) Then
            currentForm.left = otherForm.left - currentForm.Width + Screen.TwipsPerPixelX
            snapForm = "snap"
        End If
    End If


    'snap top to top
    If Abs(currentForm.Top - otherForm.Top) <= 8 * Screen.TwipsPerPixelY Then
        If (currentForm.left + currentForm.Width + 8 * Screen.TwipsPerPixelX) >= otherForm.left And currentForm.left <= (otherForm.left + otherForm.Width + 8 * Screen.TwipsPerPixelX) Then
            currentForm.Top = otherForm.Top
            snapForm = "snap"
        End If
    'snap top to bottom
    ElseIf Abs(currentForm.Top - otherForm.Top - otherForm.Height) <= 8 * Screen.TwipsPerPixelY Then
        If (currentForm.left + currentForm.Width + 8 * Screen.TwipsPerPixelX) >= otherForm.left And currentForm.left <= (otherForm.left + otherForm.Width + 8 * Screen.TwipsPerPixelX) Then
            currentForm.Top = otherForm.Top + otherForm.Height - Screen.TwipsPerPixelY
            snapForm = "snap"
        End If
    End If
    'snap left to left
    If Abs(currentForm.left - otherForm.left) <= 8 * Screen.TwipsPerPixelX Then
        If (currentForm.Top + currentForm.Height + 8 * Screen.TwipsPerPixelY) >= otherForm.Top And currentForm.Top <= (otherForm.Top + otherForm.Height + 8 * Screen.TwipsPerPixelY) Then
            currentForm.left = otherForm.left
           snapForm = "snap"
        End If
    'snap left to right
    ElseIf Abs(currentForm.left - otherForm.left - otherForm.Width) <= 8 * Screen.TwipsPerPixelX Then
        If (currentForm.Top + currentForm.Height + 8 * Screen.TwipsPerPixelY) >= otherForm.Top And currentForm.Top <= (otherForm.Top + otherForm.Height + 8 * Screen.TwipsPerPixelY) Then
            currentForm.left = otherForm.left + otherForm.Width - Screen.TwipsPerPixelX
            snapForm = "snap"
        End If
    End If

End Function


Public Function GetSoldatDir() As String

    On Error GoTo ErrorHandler

    'HKEY_CLASSES_ROOT\Soldat\DefaultIcon

    Dim hKey As Long
    Dim sKey As String

    sKey = "Soldat\DefaultIcon"
    hKey = OpenRegKey(HKEY_CLASSES_ROOT, sKey)

    If hKey <> 0 Then

        GetSoldatDir = GetRegValue(hKey, "")
        RegCloseKey hKey

    Else
        'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Soldat_is1\Inno Setup: App Path

        sKey = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Soldat_is1"
        hKey = OpenRegKey(HKEY_LOCAL_MACHINE, sKey)

        If hKey <> 0 Then

            GetSoldatDir = GetRegValue(hKey, "Inno Setup: App Path")
            RegCloseKey hKey

        Else

            GetSoldatDir = "C:\Soldat"

        End If

    End If

    If Not DirExists(GetSoldatDir) Then

        MsgBox "Could not locate the Soldat directory. (" & GetSoldatDir & ")" & vbNewLine & "Please configure the Soldat path, otherwise PolyWorks will not work properly." & vbNewLine & "See: Edit -> Preferences"

    End If

    Exit Function

ErrorHandler:

    MsgBox "Error getting soldat directory from registry" & vbNewLine & Error$

End Function


Private Function DirExists(DirName As String) As Boolean

    On Error GoTo ErrorHandler
    DirExists = GetAttr(DirName) And vbDirectory

ErrorHandler:

End Function


Private Function OpenRegKey(ByVal hKey As Long, ByVal lpSubKey As String) As Long

    Dim hSubKey As Long

    If RegOpenKeyEx(hKey, lpSubKey, 0, KEY_READ, hSubKey) = 0 Then

        OpenRegKey = hSubKey

    End If

End Function

Private Function GetRegValue(hSubKey As Long, sKeyName As String) As String

    Dim lpValue As String 'name of the value to retrieve
    Dim lpcbData As Long  'length of the retrieved value
    Dim Result As Long

    'if valid
    If hSubKey <> 0 Then

        lpValue = Space$(260)
        lpcbData = Len(lpValue)

        'find the passed value if present
        If RegQueryValueEx(hSubKey, sKeyName, 0&, 0&, ByVal lpValue, lpcbData) = 0 Then

            GetRegValue = left$(lpValue, lstrlenW(StrPtr(lpValue)))

        End If

    End If

End Function


Public Function getFileDate(fileName As String) As Long

    On Error GoTo ErrorHandler

    Dim hFile As Long

    Dim OFS As OFSTRUCT
    Dim FT_CREATE As FILETIME
    Dim FT_ACCESS As FILETIME
    Dim FT_WRITE As FILETIME

    Dim dosDate As Integer, dosTime As Integer
    Dim timeString As String
    Dim localFT As FILETIME
    Dim sysTime As SYSTEMTIME

    hFile = OpenFile(frmSoldatMapEditor.soldatDir & "Scenery-gfx\" + fileName, OFS, OF_READWRITE)
    Call GetFileTime(hFile, FT_CREATE, FT_ACCESS, FT_WRITE)
    Call CloseHandle(hFile)

    Call FileTimeToLocalFileTime(FT_WRITE, localFT)
    FT_WRITE = localFT
    Call FileTimeToDosDateTime(FT_WRITE, VarPtr(dosDate), VarPtr(dosTime))
    timeString = Hex$(dosTime)
    If Len(timeString) < 4 Then
        timeString = String$(4 - Len(timeString), "0") & timeString
    End If

    getFileDate = CLng("&H" & Hex$(dosDate) & timeString)

    Exit Function

ErrorHandler:

    MsgBox "get file date" & vbNewLine & Error$

End Function

Public Sub saveSection(sectionName As String, sectionData As String, Optional fileName As String)

    Dim lReturn  As Long

    If fileName = "" Then
        fileName = appPath & "\polyworks.ini"
    End If

    lReturn = WritePrivateProfileSection(sectionName, sectionData, fileName)

End Sub

Public Function loadString(section As String, Entry As String, Optional fileName As String, Optional length As Integer) As String

    Dim sString  As String
    Dim lSize    As Long
    Dim lReturn  As Long

    If fileName = "" Then
        fileName = appPath & "\polyworks.ini"
    End If

    If length = 0 Then length = 10

    sString = String$(length, "*")
    lSize = Len(sString)
    lReturn = GetPrivateProfileString(section, Entry, "", sString, lSize, fileName)

    loadString = left(sString, lReturn)

End Function

Public Function loadInt(section As String, Entry As String, Optional fileName As String) As Long

    Dim lReturn As Long

    If fileName = "" Then
        fileName = appPath & "\polyworks.ini"
    End If

    lReturn = GetPrivateProfileInt(section, Entry, -1, fileName)

    loadInt = lReturn

End Function

Public Function loadSection(section As String, ByRef lReturn As String, length As Integer, Optional fileName As String) As String

    If fileName = "" Then
        fileName = appPath & "\polyworks.ini"
    End If

    GetPrivateProfileSection section, lReturn, length, fileName

    loadSection = lReturn

End Function

Public Function RGBtoHex(DecValue As Long) As String

    Dim hexValue As String

    hexValue = Hex$(Val(DecValue))

    If Len(hexValue) < 6 Then
        hexValue = String$(6 - Len(hexValue), "0") + hexValue
    End If

    RGBtoHex = hexValue

End Function

Public Function HexToLong(hexValue As String) As Long

    On Error GoTo ErrorHandler

    If Len(hexValue) > 8 Then
        hexValue = right$(hexValue, 8)
    End If

    HexToLong = CLng("&H" & hexValue)

    Exit Function

ErrorHandler:

    HexToLong = -1

End Function

Public Sub RunSoldat()

    frmSoldatMapEditor.picMinimize_MouseUp 1, 0, 0, 0

    ShellExecute 0&, vbNullString, frmSoldatMapEditor.soldatDir & "Soldat.exe", "-start", vbNullString, vbNormalFocus

End Sub

Public Sub RunHelp()

    Dim iReturn As Long

    iReturn = ShellExecute(frmSoldatMapEditor.hWnd, "Open", appPath & "\PolyWorks Help.html", vbNullString, vbNullString, vbNormalFocus) 'SW_ShowNormal)

End Sub

Public Sub SetGameMode(fileName As String)

    Dim lReturn As Long
    Dim gameMode As Integer

    If LCase(left(fileName, 4)) = "ctf_" Then
        gameMode = 3
    ElseIf LCase(left(fileName, 4)) = "inf_" Then
        gameMode = 5
    ElseIf LCase(left(fileName, 4)) = "htf_" Then
        gameMode = 6
    Else
        gameMode = 0
    End If

    lReturn = WritePrivateProfileString("GAME", "GameStyle", gameMode, frmSoldatMapEditor.soldatDir & "soldat.ini")

End Sub

Public Sub SetColors()

    frmSoldatMapEditor.picMenuBar.BackColor = bgClr
    frmSoldatMapEditor.picStatus.BackColor = bgClr
    frmPreferences.BackColor = bgClr
    frmColor.BackColor = bgClr
    frmDisplay.BackColor = bgClr
    frmInfo.BackColor = bgClr
    frmMap.BackColor = bgClr

    frmScenery.BackColor = bgClr
    frmTools.BackColor = bgClr
    frmWaypoints.BackColor = bgClr

End Sub

'Initialises GDI Plus
Public Function InitGDIPlus() As Long
    Dim Token    As Long
    Dim gdipInit As GdiplusStartupInput

    gdipInit.GdiplusVersion = 1
    GdiplusStartup Token, gdipInit, ByVal 0&
    InitGDIPlus = Token
End Function

'Frees GDI Plus
Public Sub FreeGDIPlus(Token As Long)
    GdiplusShutdown Token
End Sub

'Loads the picture (optionally resized)
Public Function LoadPictureGDIPlus(PicFile As String, Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal BackColor As Long = vbWhite) As IPicture

    On Error GoTo ErrorHandler

    Dim hDC     As Long
    Dim hBitmap As Long
    Dim Img     As Long
    Dim hBrush As Long
    Dim Graphics   As Long      'Graphics Object Pointer

    Dim IID_IDispatch As GUID
    Dim pic           As PICTDESC
    Dim IPic          As IPicture

    'Load the image
    If Len(Dir$(PicFile)) <> 0 Then
        If GdipLoadImageFromFile(StrPtr(PicFile), Img) <> 0 Then
            Exit Function
        End If
    End If
    'Calculate picture's width and height if not specified
    If Width = -1 Or Height = -1 Then
        GdipGetImageWidth Img, Width
        GdipGetImageHeight Img, Height
    End If
    'Initialise the hDC
    'Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hDC = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hDC, hBitmap)
    hBrush = CreateSolidBrush(BackColor)
    hBrush = SelectObject(hDC, hBrush)
    PatBlt hDC, 0, 0, Width, Height, PATCOPY
    DeleteObject SelectObject(hDC, hBrush)
    'Resize the picture
    GdipCreateFromHDC hDC, Graphics
    GdipDrawImageRectI Graphics, Img, 0, 0, Width, Height
    GdipDeleteGraphics Graphics
    GdipDisposeImage Img
    'Get the bitmap back
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
    'Create the picture
    'Fill in OLE IDispatch Interface ID
    IID_IDispatch.Data1 = &H20400
    IID_IDispatch.Data4(0) = &HC0
    IID_IDispatch.Data4(7) = &H46
    'Fill Pic with necessary parts
    pic.Size = Len(pic)       'Length of structure
    pic.Type = PICTYPE_BITMAP 'Type of Picture (bitmap)
    pic.hBmp = hBitmap        'Handle to bitmap
    'Create the picture
    OleCreatePictureIndirect pic, IID_IDispatch, True, IPic
    Set LoadPictureGDIPlus = IPic

    Exit Function

ErrorHandler:

    MsgBox Error$ & vbNewLine & "Error loading picture"

End Function
