Attribute VB_Name = "modRender"
Option Explicit

' rendering maps (directx8)


' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
#End If


' vars - public

Public noRedraw As Boolean

Public initialized As Boolean
Public initialized2 As Boolean

Public selectionChanged As Boolean


Public DX As DirectX8
Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8

Public D3DX As D3DX8


Public mapTexture As Direct3DTexture8
Public particleTexture As Direct3DTexture8
Public patternTexture As Direct3DTexture8
Public objectsTexture As Direct3DTexture8
Public lineTexture As Direct3DTexture8
Public pathTexture As Direct3DTexture8
Public rCenterTexture As Direct3DTexture8
Public sketchTexture As Direct3DTexture8

Public renderTarget As Direct3DTexture8
Public renderSurface As Direct3DSurface8
Public backBuffer As Direct3DSurface8


Public scenerySprite As D3DXSprite

Public objTexSize As D3DVECTOR2
Public SceneryTextures() As TextureData
Public imageInfo As TImageInfo
Public textureDesc As D3DSURFACE_DESC

Public particleSize As Single


Public Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE


' vars - private


' functions - public

Public Sub InitDX8()

    On Error GoTo ErrorHandler

    initialized = False
    noRedraw = False
    selectionChanged = False

    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim debugVal As String


    debugVal = "Error creating Direct3D objects"

    If Not initialized2 Then
        Set D3DX = New D3DX8
        Set DX = New DirectX8
        Set D3D = DX.Direct3DCreate()
        initialized2 = True
    End If


    debugVal = "Error getting display mode"

    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3DWindow.Windowed = 1
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    D3DWindow.BackBufferFormat = D3DFMT_A8R8G8B8


    debugVal = "Error creating D3D device"

    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmOpenSoldatMapEditor.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow) ' Main screen turn on


    debugVal = "Error setting render states"

    D3DDevice.SetVertexShader FVF
    D3DDevice.SetRenderState D3DRS_LIGHTING, False

    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE  ' polys that are ccw

    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    D3DDevice.SetRenderState D3DRS_POINTSIZE, FtoDW(particleSize)

    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE

    Set renderTarget = D3DX.CreateTexture(D3DDevice, 256, 256, D3DX_DEFAULT, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
    Set renderSurface = renderTarget.GetSurfaceLevel(0)
    Set backBuffer = D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)


    debugVal = "Error creating pattern texture"

    Set patternTexture = D3DX.CreateTextureFromFile(D3DDevice, appPath & "\skins\" & gfxDir & "\pattern.bmp")


    debugVal = "Error creating objects texture"

    Set objectsTexture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\objects.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)

    objectsTexture.GetLevelDesc 0, textureDesc

    objTexSize.X = textureDesc.Width
    objTexSize.Y = textureDesc.Height


    debugVal = "Error creating scenery not found texture"

    Set SceneryTextures(0).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\notfound.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)

    SceneryTextures(0).Texture.GetLevelDesc 0, textureDesc

    SceneryTextures(0).Width = imageInfo.Width
    SceneryTextures(0).Height = imageInfo.Height

    SceneryTextures(0).reScale.X = SceneryTextures(0).Width / textureDesc.Width
    SceneryTextures(0).reScale.Y = SceneryTextures(0).Height / textureDesc.Height

    If SceneryTextures(0).reScale.X = 0 Or SceneryTextures(0).reScale.Y = 0 Then
        SceneryTextures(0).reScale.X = 1
        SceneryTextures(0).reScale.Y = 1
    End If


    debugVal = "Error creating line texture"

    Set lineTexture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\lines.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)


    debugVal = "Error creating path texture"

    Set pathTexture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\path.png", D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)


    debugVal = "Error creating rotation center texture"

    Set rCenterTexture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\rcenter.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)


    debugVal = "Error creating sketch texture"

    Set sketchTexture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\sketch.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)


    debugVal = "Error creating scenery sprite"

    Set scenerySprite = D3DX.CreateSprite(D3DDevice)


    debugVal = "Error creating particle texture"

    Set particleTexture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\vertex8x8.bmp", 8, 8, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)

    initialized = True

    Exit Sub

ErrorHandler:

    If D3DX Is Nothing Then
        MsgBox "Error initializing Direct3D" & vbNewLine & debugVal & vbNewLine & Error
    Else
        MsgBox "Error initializing Direct3D" & vbNewLine & D3DX.GetErrorString(err.Number) & vbNewLine & debugVal
    End If

End Sub


' functions - private

Private Function FtoDW(f As Single) As Long

    Dim buf As D3DXBuffer
    Dim l As Long

    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, l
    FtoDW = l

End Function
