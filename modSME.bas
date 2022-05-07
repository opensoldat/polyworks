Attribute VB_Name = "modSME"
Option Explicit

' misc stuff


' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
#End If


' vars - public

Public Const PI As Single = 3.14159265358979  ' mmm... PI

Public gfxDir As String

Public appPath As String
Public bgColor As Long
Public lblBackColor As Long
Public lblTextColor As Long
Public txtBackColor As Long
Public txtTextColor As Long
Public frameColor As Long

Public font1 As String
Public font2 As String

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

' bitblt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' stretchblit
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
        ByVal dwRop As Long) As Long

' mouse over
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
' dragging window
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_NCLBUTTONDOWN = &HA1

' taskbar
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

' get pixel
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long


' vars - private

' browse
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


' registry
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


' file time
Public Const OFS_MAXPATHNAME = 128
Public Const OF_READWRITE = &H2

Public Type OFSTRUCT
    cBytes      As Byte
    fFixedDisk  As Byte
    nErrCode    As Integer
    Reserved1   As Integer
    Reserved2   As Integer
    szPathName(0 To OFS_MAXPATHNAME - 1) As Byte  ' 0-based
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

' ini file
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

' ShellExecute
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' key mapping
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" _
        (ByVal wCode As Long, ByVal wMapType As Long) As Long

' GDI+
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

' GDI Functions
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

' GDI+ functions
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hBmp As Long, ByVal hpal As Long, GpBitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal image As Long, Height As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal token As Long)

' functions for gif loading
Private Declare Function GdipSaveImageToFile Lib "gdiplus.dll" (ByVal image As Long, ByVal FileName As Long, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus.dll" (ByVal FileName As Long, ByRef Bitmap As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus.dll" (ByVal Bitmap As Long, ByRef hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus.dll" (ByRef numEncoders As Long, ByRef Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus.dll" (ByVal numEncoders As Long, ByVal Size As Long, ByRef Encoders As Any) As Long
Private Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long


' GDI and GDI+ constants
Private Const PLANES = 14            ' Number of planes
Private Const BITSPIXEL = 12         ' Number of bits per pixel
Private Const PATCOPY = &HF00021     ' (DWORD) dest = pattern
Private Const PICTYPE_BITMAP = 1     ' Bitmap type
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7

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
        buffer = Space(lstrlenW(ByVal pImageCodecInfo(j).MimeTypePtr))

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

    Dim token As Long

    On Error GoTo ErrorHandler

    token = InitGDIPlus

    If SaveImageAsPNG(src, dest) Then
        GifToPng = -1
    Else
        GifToPng = 5
    End If

ErrorHandler:

    FreeGDIPlus token

End Function

Public Function GifToBmp(ByVal src As String, ByVal dest As String) As Long

    GifToBmp = GifToPng(src, dest)

End Function

' mouse event
Public Function MouseEvent(ByRef pic As PictureBox, ByVal xVal As Integer, ByVal yVal As Integer, xSrc As Integer, ySrc As Integer, Width As Integer, Height As Integer) As Boolean

    If (xVal < 0) Or (xVal > Width) Or (yVal < 0) Or (yVal > Height) Then  ' the MOUSELEAVE pseudo-event
        ReleaseCapture
        BitBlt pic.hDC, 0, 0, Width, Height, frmSoldatMapEditor.picGfx.hDC, xSrc, ySrc, vbSrcCopy
        pic.Refresh
        MouseEvent = True
    ElseIf GetCapture() <> pic.hWnd Then  ' the MOUSEENTER pseudo-event
        SetCapture pic.hWnd
        BitBlt pic.hDC, 0, 0, Width, Height, frmSoldatMapEditor.picGfx.hDC, xSrc + Width, ySrc, vbSrcCopy
        pic.Refresh
        MouseEvent = True
    End If

End Function

' mouse event
Public Function MouseEvent2(ByRef pic As PictureBox, ByVal xVal As Integer, ByVal yVal As Integer, ByVal buttonType As Byte, ByVal active As Byte, ByVal action As Byte, Optional exWidth As Integer) As Boolean

    Dim xSrc As Integer
    Dim ySrc As Integer
    Dim Width As Integer
    Dim Height As Integer

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
        MouseEvent2 = True
    ElseIf (xVal < 0) Or (xVal > exWidth) Or (yVal < 0) Or (yVal > Height) Then  ' the MOUSELEAVE pseudo-event
        ReleaseCapture
        MouseEvent2 = True
        action = BUTTON_UP
    ElseIf GetCapture() <> pic.hWnd Then  ' the MOUSEENTER pseudo-event
        SetCapture pic.hWnd
        MouseEvent2 = True
        action = BUTTON_MOVE
    End If

    If MouseEvent2 = True Then
        BitBlt pic.hDC, 0, 0, Width, Height, frmSoldatMapEditor.picButtonGfx.hDC, xSrc + Width * action, ySrc + active * Height, vbSrcCopy
        pic.Refresh
    End If

    Exit Function

ErrorHandler:

    MsgBox "Error in mouse event (2)" & vbNewLine & Error

End Function


' browse
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
    path = Space(MAX_PATH)

    If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
        pos = InStr(path, Chr(0))
        SelectFolder = Left(path, pos - 1)

        If Right(SelectFolder, 1) <> "\" Then
            SelectFolder = SelectFolder & "\"
        End If
    End If

    Call CoTaskMemFree(pidl)

End Function



Public Function SnapForm(currentForm As Form, otherForm As Form) As String

    Const SNAP_DELTA = 8
    Dim SNAP_DELTA_X As Single
    Dim SNAP_DELTA_Y As Single

    Dim currentFormBottom As Single
    Dim currentFormRight As Single
    Dim otherFormBottom As Single
    Dim otherFormRight As Single

    SNAP_DELTA_X = SNAP_DELTA * Screen.TwipsPerPixelX
    SNAP_DELTA_Y = SNAP_DELTA * Screen.TwipsPerPixelY

    currentFormBottom = currentForm.Top + currentForm.Height
    currentFormRight = currentForm.Left + currentForm.Width
    otherFormBottom = otherForm.Top + otherForm.Height
    otherFormRight = otherForm.Left + otherForm.Width

    SnapForm = ""

    ' snap bottom to bottom
    If Abs(currentFormBottom - otherFormBottom) <= SNAP_DELTA_Y Then
        If (currentFormRight + SNAP_DELTA_X) >= otherForm.Left And currentForm.Left <= (otherFormRight + SNAP_DELTA_X) Then
            currentForm.Top = otherFormBottom - currentForm.Height
            SnapForm = "snap"
        End If
    ' snap bottom to top
    ElseIf Abs(currentFormBottom - otherForm.Top) <= SNAP_DELTA_Y Then
        If (currentFormRight + SNAP_DELTA_X) >= otherForm.Left And currentForm.Left <= (otherFormRight + SNAP_DELTA_X) Then
            currentForm.Top = otherForm.Top - currentForm.Height + Screen.TwipsPerPixelY
            SnapForm = "snap"
        End If
    End If

    ' snap right to right
    If Abs(currentFormRight - otherFormRight) <= SNAP_DELTA_X Then
        If (currentFormBottom + SNAP_DELTA_Y) >= otherForm.Top And currentForm.Top <= (otherFormBottom + SNAP_DELTA_Y) Then
            currentForm.Left = otherFormRight - currentForm.Width
            SnapForm = "snap"
        End If
    ' snap right to left
    ElseIf Abs(currentFormRight - otherForm.Left) <= SNAP_DELTA_X Then
        If (currentFormBottom + SNAP_DELTA_Y) >= otherForm.Top And currentForm.Top <= (otherFormBottom + SNAP_DELTA_Y) Then
            currentForm.Left = otherForm.Left - currentForm.Width + Screen.TwipsPerPixelX
            SnapForm = "snap"
        End If
    End If


    currentFormBottom = currentForm.Top + currentForm.Height
    currentFormRight = currentForm.Left + currentForm.Width


    ' snap top to top
    If Abs(currentForm.Top - otherForm.Top) <= SNAP_DELTA_Y Then
        If (currentFormRight + SNAP_DELTA_X) >= otherForm.Left And currentForm.Left <= (otherFormRight + SNAP_DELTA_X) Then
            currentForm.Top = otherForm.Top
            SnapForm = "snap"
        End If
    ' snap top to bottom
    ElseIf Abs(currentForm.Top - otherFormBottom) <= SNAP_DELTA_Y Then
        If (currentFormRight + SNAP_DELTA_X) >= otherForm.Left And currentForm.Left <= (otherFormRight + SNAP_DELTA_X) Then
            currentForm.Top = otherFormBottom - Screen.TwipsPerPixelY
            SnapForm = "snap"
        End If
    End If

    ' snap left to left
    If Abs(currentForm.Left - otherForm.Left) <= SNAP_DELTA_X Then
        If (currentFormBottom + SNAP_DELTA_Y) >= otherForm.Top And currentForm.Top <= (otherFormBottom + SNAP_DELTA_Y) Then
            currentForm.Left = otherForm.Left
            SnapForm = "snap"
        End If
    ' snap left to right
    ElseIf Abs(currentForm.Left - otherFormRight) <= SNAP_DELTA_X Then
        If (currentFormBottom + SNAP_DELTA_Y) >= otherForm.Top And currentForm.Top <= (otherFormBottom + SNAP_DELTA_Y) Then
            currentForm.Left = otherFormRight - Screen.TwipsPerPixelX
            SnapForm = "snap"
        End If
    End If

End Function


Public Function GetSoldatDir() As String

    On Error GoTo ErrorHandler

    Dim hKey As Long
    Dim sKey As String

    ' HKEY_CLASSES_ROOT\Soldat\DefaultIcon
    sKey = "Soldat\DefaultIcon"
    hKey = OpenRegKey(HKEY_CLASSES_ROOT, sKey)

    If hKey <> 0 Then
        GetSoldatDir = GetRegValue(hKey, "")
        RegCloseKey hKey
    Else
        ' HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Soldat_is1\Inno Setup: App Path
        sKey = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Soldat_is1"
        hKey = OpenRegKey(HKEY_LOCAL_MACHINE, sKey)

        If hKey <> 0 Then
            GetSoldatDir = GetRegValue(hKey, "Inno Setup: App Path")
            RegCloseKey hKey
        Else
            GetSoldatDir = "C:\Soldat"
        End If
    End If

    ' Fix soldat installer sets invalid soldat dir path
    GetSoldatDir = Replace(GetSoldatDir, Chr(34), "")
    ' Fix other possible paths too
    GetSoldatDir = Replace(GetSoldatDir, "'", "")
    GetSoldatDir = Replace(GetSoldatDir, "/", "\")

    If Not DirExists(GetSoldatDir) And FileExists(GetSoldatDir) Then
        GetSoldatDir = Left(GetSoldatDir, InStrRev(GetSoldatDir, "\"))
    End If

    If Not DirExists(GetSoldatDir) Then
        MsgBox "Could not locate the Soldat directory. (" & GetSoldatDir & ")" & vbNewLine & "Please configure the Soldat path, otherwise PolyWorks will not work properly." & vbNewLine & "See: Edit -> Preferences"
    End If

    Exit Function

ErrorHandler:

    MsgBox "Error getting soldat directory from registry" & vbNewLine & Error

End Function


Private Function OpenRegKey(ByVal hKey As Long, ByVal lpSubKey As String) As Long

    Dim hSubKey As Long

    If RegOpenKeyEx(hKey, lpSubKey, 0, KEY_READ, hSubKey) = 0 Then
        OpenRegKey = hSubKey
    End If

End Function

Private Function GetRegValue(hSubKey As Long, sKeyName As String) As String

    Dim lpValue As String  ' name of the value to retrieve
    Dim lpcbData As Long   ' length of the retrieved value
    Dim Result As Long

    ' if valid
    If hSubKey <> 0 Then
        lpValue = Space(260)
        lpcbData = Len(lpValue)

        ' find the passed value if present
        If RegQueryValueEx(hSubKey, sKeyName, 0&, 0&, ByVal lpValue, lpcbData) = 0 Then
            GetRegValue = Left(lpValue, lstrlenW(StrPtr(lpValue)))
        End If
    End If

End Function


Public Function GetFileDate(FileName As String) As Long

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

    hFile = OpenFile(frmSoldatMapEditor.soldatDir & "Scenery-gfx\" + FileName, OFS, OF_READWRITE)
    Call GetFileTime(hFile, FT_CREATE, FT_ACCESS, FT_WRITE)
    Call CloseHandle(hFile)

    Call FileTimeToLocalFileTime(FT_WRITE, localFT)
    FT_WRITE = localFT
    Call FileTimeToDosDateTime(FT_WRITE, VarPtr(dosDate), VarPtr(dosTime))
    timeString = Hex(dosTime)
    If Len(timeString) < 4 Then
        timeString = String(4 - Len(timeString), "0") & timeString
    End If

    GetFileDate = CLng("&H" & Hex(dosDate) & timeString)

    Exit Function

ErrorHandler:

    MsgBox "Error getting file date" & vbNewLine & Error

End Function

Public Sub SaveSection(sectionName As String, sectionData As String, Optional FileName As String)

    Dim lReturn  As Long

    If FileName = "" Then
        FileName = appPath & "\polyworks.ini"
    End If

    lReturn = WritePrivateProfileSection(sectionName, sectionData, FileName)

End Sub

Public Function LoadString(section As String, Entry As String, Optional FileName As String, Optional length As Integer, Optional DefaultValue As String = "") As String

    Dim sString  As String
    Dim lSize    As Long
    Dim lReturn  As Long

    If FileName = "" Then
        FileName = appPath & "\polyworks.ini"
    End If

    If length = 0 Then length = 10

    sString = String(length, "*")
    lSize = Len(sString)
    lReturn = GetPrivateProfileString(section, Entry, DefaultValue, sString, lSize, FileName)

    LoadString = Left(sString, lReturn)

End Function

Public Function LoadInt(section As String, Entry As String, Optional FileName As String, Optional DefaultValue As Long = -1) As Long

    Dim lReturn As Long

    If FileName = "" Then
        FileName = appPath & "\polyworks.ini"
    End If

    lReturn = GetPrivateProfileInt(section, Entry, DefaultValue, FileName)

    LoadInt = lReturn

End Function

Public Function LoadByte(section As String, Entry As String, Optional FileName As String, Optional DefaultValue As Byte = 0) As Byte

    Dim lReturn As Byte

    If FileName = "" Then
        FileName = appPath & "\polyworks.ini"
    End If

    lReturn = GetPrivateProfileInt(section, Entry, DefaultValue, FileName)

    LoadByte = lReturn

End Function

Public Function LoadSection(section As String, ByRef lReturn As String, length As Integer, Optional FileName As String) As String  ' unused?

    If FileName = "" Then
        FileName = appPath & "\polyworks.ini"
    End If

    GetPrivateProfileSection section, lReturn, length, FileName

    LoadSection = lReturn

End Function

Public Function RGBtoHex(DecValue As Long) As String

    Dim hexValue As String

    hexValue = Hex(Val(DecValue))

    If Len(hexValue) < 6 Then
        hexValue = String(6 - Len(hexValue), "0") + hexValue
    End If

    RGBtoHex = hexValue

End Function

Public Function HexToLong(hexValue As String, Optional DefaultValue As Long = -1) As Long

    On Error GoTo ErrorHandler

    If Len(hexValue) > 8 Then
        hexValue = Right(hexValue, 8)
    ElseIf Len(hexValue) = 0 Then
        hexValue = DefaultValue
        Exit Function
    End If

    HexToLong = CLng("&H" & hexValue)

    Exit Function

ErrorHandler:

    HexToLong = DefaultValue

End Function

Public Sub RunSoldat()

    frmSoldatMapEditor.picMinimize_MouseUp 1, 0, 0, 0

    ShellExecute 0&, vbNullString, frmSoldatMapEditor.soldatDir & "Soldat.exe", "-start", vbNullString, vbNormalFocus

End Sub

Public Sub RunHelp()

    Dim iReturn As Long

    iReturn = ShellExecute(frmSoldatMapEditor.hWnd, "Open", appPath & "\PolyWorks Help.html", vbNullString, vbNullString, vbNormalFocus) 'SW_ShowNormal)

End Sub

Public Sub SetGameMode(FileName As String)

    Dim lReturn As Long
    Dim gameMode As Integer

    If LCase(Left(FileName, 4)) = "ctf_" Then
        gameMode = 3
    ElseIf LCase(Left(FileName, 4)) = "inf_" Then
        gameMode = 5
    ElseIf LCase(Left(FileName, 4)) = "htf_" Then
        gameMode = 6
    Else
        gameMode = 0
    End If

    lReturn = WritePrivateProfileString("GAME", "GameStyle", gameMode, frmSoldatMapEditor.soldatDir & "soldat.ini")

End Sub

Public Sub SetColors()

    frmSoldatMapEditor.picMenuBar.BackColor = bgColor
    frmSoldatMapEditor.picStatus.BackColor = bgColor
    frmSoldatMapEditor.picResize.BackColor = bgColor
    frmPreferences.BackColor = bgColor
    frmColor.BackColor = bgColor
    frmDisplay.BackColor = bgColor
    frmInfo.BackColor = bgColor
    frmMap.BackColor = bgColor

    frmScenery.BackColor = bgColor
    frmTools.BackColor = bgColor
    frmWaypoints.BackColor = bgColor

End Sub

' Initializes GDI+
Public Function InitGDIPlus() As Long

    Dim token    As Long
    Dim gdipInit As GdiplusStartupInput

    gdipInit.GdiplusVersion = 1
    GdiplusStartup token, gdipInit, ByVal 0&
    InitGDIPlus = token

End Function

' Frees GDI Plus
Public Sub FreeGDIPlus(token As Long)

    GdiplusShutdown token

End Sub

' Loads the picture (optionally resized)
Public Function LoadPictureGDIPlus(PicFile As String, Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal BackColor As Long = vbWhite) As IPicture

    On Error GoTo ErrorHandler

    Dim hDC     As Long
    Dim hBitmap As Long
    Dim Img     As Long
    Dim hBrush As Long
    Dim Graphics   As Long  ' Graphics Object Pointer

    Dim IID_IDispatch As GUID
    Dim pic           As PICTDESC
    Dim IPic          As IPicture

    ' Load the image
    If Len(Dir(PicFile)) <> 0 Then
        If GdipLoadImageFromFile(StrPtr(PicFile), Img) <> 0 Then
            Exit Function
        End If
    End If
    ' Calculate picture's width and height if not specified
    If Width = -1 Or Height = -1 Then
        GdipGetImageWidth Img, Width
        GdipGetImageHeight Img, Height
    End If
    ' Initialize the hDC
    ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hDC = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hDC, hBitmap)
    hBrush = CreateSolidBrush(BackColor)
    hBrush = SelectObject(hDC, hBrush)
    PatBlt hDC, 0, 0, Width, Height, PATCOPY
    DeleteObject SelectObject(hDC, hBrush)
    ' Resize the picture
    GdipCreateFromHDC hDC, Graphics
    GdipDrawImageRectI Graphics, Img, 0, 0, Width, Height
    GdipDeleteGraphics Graphics
    GdipDisposeImage Img
    ' Get the bitmap back
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
    ' Create the picture
    ' Fill in OLE IDispatch Interface ID
    IID_IDispatch.Data1 = &H20400
    IID_IDispatch.Data4(0) = &HC0
    IID_IDispatch.Data4(7) = &H46
    ' Fill Pic with necessary parts
    pic.Size = Len(pic)        ' Length of structure
    pic.Type = PICTYPE_BITMAP  ' Type of Picture (bitmap)
    pic.hBmp = hBitmap         ' Handle to bitmap
    ' Create the picture
    OleCreatePictureIndirect pic, IID_IDispatch, True, IPic
    Set LoadPictureGDIPlus = IPic

    Exit Function

ErrorHandler:

    MsgBox "Error loading picture" & vbNewLine & Error

End Function
