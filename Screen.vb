Module aScreen

    'all code is implented

    Public Const ColorBlue As Long = &HFF0000FF
    Public Const ColorRed As Long = &HFFFF0000
    Public Const ColorPink As Long = &HFFFF0FF0
    Public Const ColorLightGreen As Long = &HFF00FF00
    Public Const ColorYellow As Long = &HFFFFF000
    Public Const ColorWhite As Long = &HFFFFFFF0
    Public Const ColorBlack As Long = &HFF000000
    Public Const ColorBlueGreen As Long = &HFF0FFFF0
    Public Const FVF_TEX As Long = (D3DFVF_XYZ Or D3DFVF_TEX1)
    Public Const PI As Single = 3.14159275180032
    Public Const Rad As Single = PI / 180
    Public CurrentWidth As Integer
    Public CurrentHeight As Integer
    Public NewWidth As Double 
    Public NewHeight As Double 
    Public FixWidth As Double 
    Public FixHeight As Double 
    Public vFontSize As String
    Public vDistanceFromCamera As Single
    Public vAnimationSpeed As Long
    Public v3DModels As New a3DWorld
    Public vBackgroundTexture As Direct3DTexture8
    Public vForegroundTexture As Direct3DTexture8
    Public vWorldBackground(35) As TextureVertexType
    Public vVertexBuffer As Direct3DVertexBuffer8
    Public vPlanetChangeTime As Long
    Public vAngFromSun As Single
    Public vPlanetModelIndex As Integer
    Public vAngFromEarth As Single
    Public vDistanceFromSun As Single
    Public vAngMoon As Single
    Public vCamPlace As D3DVECTOR
    Public vCamRotate As D3DVECTOR
    Public TextColor(122) As Long
    Public ColorButtonNormal As Long
    Public ColorButtonPressed As Long
    Public ColorTextNormal As Long
    Public ColorTextSelected As Long
    Public ColorListSelected As Long
    Public ColorListNormal As Long
    Public ColorListFocus As Long
    Public ColorLabelNormal As Long
    Public ColorLabelPressed As Long
    Public SkinThema As String
    Public ImageSpace As Integer
    Public DX As New DirectX8
    Public D3DX As New D3DX8
    Public D3DFont As D3DXFont
    Public D3D As Direct3D8
    Public D3DDevice As Direct3DDevice8
    Public D3DTexture() As Direct3DTexture8
    Public D3DTexture2() As String
    Public BackBufferSurf As Direct3DSurface8
    Public OffScreenTexture As Direct3DTexture8
    Public OffScreenSurf As Direct3DSurface8
    Public TextCursor As RECT
    Public TextList As String
    Public TextCursorPosition As Long
    Public TextCursorString As String
    Public TextSelectedStart As Long
    Public TextSelectedCount As Integer
    Public ColorVertex(0 To 3) As ColorVertexT
    Public TextSelected() As SelectText
    Structure ColorVertexT
        Public X As Single
        Public y As Single
        Public z As Single
        Public r As Single
        Public c As Long
        Public S As Long
        Public u As Single
        Public v As Single
    End Structure
    Structure TextureVertexStructure
        Public X As Single
        Public y As Single
        Public z As Single
        Public tu As Single
        Public tv As Single
    End Structure
    Structure SelectText
        Public Bottom As Long
        Public Right As Long
        Public Top As Long
        Public Left As Long
        Public Visible As Boolean
        Public Text As String
    End Structure
    Structure DStateMouseProxyT
        Public Buttons(1) As Byte
        Public lZ As Long
    End Structure

    Function ColorVertexMake(X As Single, y As Single, z As Single, r As Single, c As Long, S As Long, u As Single, v As Single) As ColorVertexT
        ColorVertexMake.X = X
        ColorVertexMake.y = y
        ColorVertexMake.z = z
        ColorVertexMake.r = r
        ColorVertexMake.c = c
        ColorVertexMake.S = S
        ColorVertexMake.u = u
        ColorVertexMake.v = v
    End Function

    Sub FormResize()
        Dim iA As Integer, iB As Integer, iC As Integer, oA As D3DDISPLAYMODE, MainFont As IFont, TempFont As New StdFont
        Static iY As Long, iZ As Long
        oA = GetScreenResolution()

        If FixWidth = CurrentWidth / 1024 Or oA.Width = 0 Then
            Exit Sub
        End If
        CurrentWidth = oA.Width
        CurrentHeight = oA.Height
        If iZ = 0 Then
            FixWidth = CurrentWidth / 1024
            FixHeight = CurrentHeight / 768
        Else
            FixWidth = CurrentWidth / iZ
            FixHeight = CurrentHeight / iY
        End If
        If CurrentWidth > 800 Then
            TempFont.Name = "Arial"
            TempFont.Bold = True
        Else
            TempFont.Name = "Small Fonts"
            TempFont.Bold = False
        End If
        Select Case CurrentWidth
            Case Is < 800
                TempFont.Size = 8
            Case 800, 960
                TempFont.Size = 9
            Case 1024
                TempFont.Size = 10
            Case 1152
                TempFont.Size = 11
            Case 1280
                TempFont.Size = 12
            Case 1360, 1366
                TempFont.Size = 10
            Case Is >= 1600
                TempFont.Size = 10
            Case 1792
                TempFont.Size = 14
            Case 1800
                TempFont.Size = 15
            Case Is > 1800
                TempFont.Size = 16
            Case Else
                Call Message("Resolution To Low " & CurrentWidth & "x" & CurrentHeight, vbCritical)
                End
        End Select
        vFontSize = TempFont.Size + 8
        If FixWidth = 1 And FixHeight = 1 Then
            FixWidth = CurrentWidth / 1024
            FixHeight = CurrentHeight / 768
        Else
            For iB = 2 To BFormCount
                For iA = 0 To DXForm(iB).ControlsCount
                    For iC = 0 To DXForm(iB).Controls(iA).PositionCount
                        Call xForm(iB).Controls(iC).Move(xForm(iB).Controls(iC).Left * FixWidth, xForm(iB).Controls(iC).Top * FixHeight, xForm(iB).Controls(iC).Width * FixWidth, xForm(iB).Controls(iC).Height * FixHeight)
                    Next iC
                Next iA
                DXForm(iB).Width = DXForm(iB).Width * FixWidth
                DXForm(iB).Height = DXForm(iB).Height * FixHeight
            Next iB
            iZ = CurrentWidth
            iY = CurrentHeight
            FixWidth = CurrentWidth / 1024
            FixHeight = CurrentHeight / 768
        End If
        MainFont = TempFont
        D3DFont = D3DX.CreateFont(D3DDevice, MainFont.hFont)
    End Sub

    Function GetScreenResolution() As D3DDISPLAYMODE
        ''On Local Error Resume Next
        Call D3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, GetScreenResolution)
    End Function

    Function GetTextColor(sA As String) As Long
        Select Case sA
            Case "Blue"
                GetTextColor = ColorBlue
            Case "Red"
                GetTextColor = ColorRed
            Case "Pink"
                GetTextColor = ColorPink
            Case "Light Green"
                GetTextColor = ColorLightGreen
            Case "Yellow"
                GetTextColor = ColorYellow
            Case "White"
                GetTextColor = ColorWhite
            Case "Black"
                GetTextColor = ColorBlack
            Case "Blue / Green"
                GetTextColor = ColorBlueGreen
        End Select
    End Function

    Sub LoadScreen()
        Call aScreen.LoadSkin(Path(App.Path))
        FixWidth = 0
        Call aScreen.FormResize()
        Call aScreen.RenderScreen()
    End Sub

    Private Sub LoadSkin(sB As String, Optional OnlyImageLoad As Boolean = False)
        Dim D3DDispMode As D3DDISPLAYMODE, D3DPresent As D3DPRESENT_PARAMETERS, sA As String, vModelIndex As Integer
        ImageSpace = 122
        ReDim D3DTexture(ImageSpace)
        If OnlyImageLoad = False Then
            D3D = DX.Direct3DCreate()
            Call D3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, D3DDispMode)
            D3DPresent.Windowed = 1
            If Screen.Width & Screen.Height > 72007199 Then
                Call AForm1.Move(0, 0, D3DDispMode.Width * 15, D3DDispMode.Height * 15)
            Else
                Call AForm1.Move(0, 0, D3DDispMode.Width, D3DDispMode.Height)
            End If
            CurrentWidth = D3DDispMode.Width
            CurrentHeight = D3DDispMode.Height
            D3DPresent.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
            D3DPresent.AutoDepthStencilFormat = D3DFMT_D16
            D3DPresent.EnableAutoDepthStencil = 1
            D3DPresent.BackBufferFormat = D3DDispMode.Format
            D3DPresent.BackBufferHeight = AForm1.ScaleHeight
            D3DPresent.BackBufferWidth = AForm1.ScaleWidth
            D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, AForm1.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPresent)
            Call D3DDevice.SetRenderState(D3DRS_LIGHTING, 0)
            Call D3DDevice.SetRenderState(D3DRS_ZENABLE, 1)
            Call D3DDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, True)
            Call D3DDevice.SetVertexShader(D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR)
            OffScreenTexture = D3DX.CreateTexture(D3DDevice, CurrentWidth, CurrentHeight, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
            OffScreenSurf = OffScreenTexture.GetSurfaceLevel(0)
            BackBufferSurf = D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
        End If
        sA = sB
        vBackgroundTexture = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Skins\Background\!Default\Background.jpg", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        vForegroundTexture = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Skins\Address Book\!Default\No Picture.JPG", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\Planets\Sun.X", 139.2)
        Call v3DModels.Rotate(vModelIndex, 0, 0.01, 0)
        Call v3DModels.Move(vModelIndex, 0, -2000, 0)
        Call v3DModels.DistanceFromCamera(vModelIndex, 1)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\Planets\Mercury.X", 0.4878)
        Call v3DModels.Rotate(vModelIndex, 0, 0.01, 0)
        Call v3DModels.Move(vModelIndex, 0, -5, 5890)
        Call v3DModels.DistanceFromCamera(vModelIndex, 0.4)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\Planets\Venus.X", 1.2104)
        Call v3DModels.Rotate(vModelIndex, 0, 0.01, 0)
        Call v3DModels.Move(vModelIndex, 0, -15, 10910)
        Call v3DModels.DistanceFromCamera(vModelIndex, 0.5)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\Planets\Earth Night.X", 1.2756)
        Call v3DModels.Rotate(vModelIndex, 0, 0.01, 0)
        Call v3DModels.Move(vModelIndex, 0, -20, 15050)
        Call v3DModels.DistanceFromCamera(vModelIndex, 0.5)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\Planets\Moon.X", 0.3476)
        Call v3DModels.Rotate(vModelIndex, 0, 0.01, 0)
        Call v3DModels.Move(vModelIndex, 0, -5, 15088)
        Call v3DModels.DistanceFromCamera(vModelIndex, 0.08)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\Ships\spstob.x", 0.005)
        Call v3DModels.Rotate(vModelIndex, -1.6, -1.6, 0)
        Call v3DModels.Move(vModelIndex, 0, 0, 15110)
        Call v3DModels.DistanceFromCamera(vModelIndex, 0.04)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\People\astnt1.x", 0.05)
        Call v3DModels.Rotate(vModelIndex, -1.6, -1.6, 0)
        Call v3DModels.Move(vModelIndex, 0, 0, 15188)
        Call v3DModels.DistanceFromCamera(vModelIndex, 0.1)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\Planets\Mars.X", 0.6788)
        Call v3DModels.Rotate(vModelIndex, 0, 0.01, 0)
        Call v3DModels.Move(vModelIndex, 0, -10, 22880)
        Call v3DModels.DistanceFromCamera(vModelIndex, 0.09)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\Planets\Jupiter.X", 14.2893)
        Call v3DModels.Rotate(vModelIndex, 0, 0.01, 0)
        Call v3DModels.Move(vModelIndex, 0, -210, 77930)
        Call v3DModels.DistanceFromCamera(vModelIndex, 1)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\Planets\Saturnus.X", 12.0#)
        Call v3DModels.Rotate(vModelIndex, 0, 0.01, 0)
        Call v3DModels.Move(vModelIndex, 0, -190, 142800)
        Call v3DModels.DistanceFromCamera(vModelIndex, 0.5)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\Planets\Uranus.X", 5.08)
        Call v3DModels.Rotate(vModelIndex, 0, 0.01, 0)
        Call v3DModels.Move(vModelIndex, 0, -80, 287000)
        Call v3DModels.DistanceFromCamera(vModelIndex, 0.1)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\Planets\Neptune.X", 4.86)
        Call v3DModels.Rotate(vModelIndex, 0, 0.01, 0)
        Call v3DModels.Move(vModelIndex, 0, -80, 449800)
        Call v3DModels.DistanceFromCamera(vModelIndex, 0.04)
        vModelIndex = v3DModels.Add(sA & "Skins\Background\!Default\Planets\Pluto.X", 0.828)
        Call v3DModels.Rotate(vModelIndex, 0, 0.01, 0)
        Call v3DModels.Move(vModelIndex, 0, -5, 590100)
        Call v3DModels.DistanceFromCamera(vModelIndex, 0.02)
        Call v3DModels.CreateWorld
        SkinThema = "Space Ship"
        sA = sB & "Skins\System\" & SkinThema & "\"
        D3DTexture(7) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Border V.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(8) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Border H.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(9) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Button Normal.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(10) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Button Click.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(11) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Scroll V 1.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(12) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Scroll V 2.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(13) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Scroll V 3.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(14) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Scroll V 4.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(15) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Scroll V 5.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(16) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Scroll V 6.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(17) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Scroll H 1.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(18) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Scroll H 2.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(19) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Scroll H 3.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(20) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Scroll H 4.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        sA = sB & "Skins\Address Book\!Default\"
        D3DTexture(28) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "No Picture.jpg", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        sA = sB & "Skins\System\" & SkinThema & "\"
        D3DTexture(21) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Scroll H 5.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(22) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Scroll H 6.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(60) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Label Left.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(61) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Label Normal.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(62) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Label Click.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(63) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Label Over.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(64) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Label Right.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(70) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Text Box 1.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(71) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Text Box 2.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(72) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Text Box 3.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(73) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Text Box 4.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(74) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Text Box 5.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(75) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Text Box 6.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(76) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Text Box 7.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(77) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Text Box 8.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(78) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Text Box 9.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(79) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 10.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(80) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 01.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(81) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 02.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(82) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 03.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(83) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 04.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(84) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 05.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(85) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 06.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(86) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 07.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(87) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 08.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(88) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 09.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(89) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 10.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(90) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 11.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(91) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 11.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(92) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "List Box 12.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(93) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Text Box Small 1.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(94) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Text Box Small 2.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(95) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Text Box Small 3.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        sA = sB & "Skins\System\" & SkinThema & "\"
        D3DTexture(111) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 01.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(112) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 02.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(113) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 03.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(114) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 04.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(115) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 05.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(116) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 06.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(117) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 07.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(118) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 08.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(119) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 09.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(120) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 10.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(121) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 11.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
        D3DTexture(122) = D3DX.CreateTextureFromFileEx(D3DDevice, sA & "Combo Box 05.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFFC0FF, 0, 0)
    End Sub

    Sub RenderBackground()
        Dim vModelIndex As Integer, vSizeTextureVertexType As TextureVertexType, vVertexSize As Single, vSpeedForwardDistanceFromCamera As Single
        Dim vecPlanet As D3DVECTOR, vecMoon As D3DVECTOR, vDistanceFromEarth As Single, VSpeedForward As Single
        Call v3DModels.matrixTransforms
        Call D3DDevice.SetVertexShader(FVF_TEX)
        Call D3DDevice.SetRenderState(D3DRS_LIGHTING, 0)
        Call D3DDevice.SetRenderState(D3DRS_ZENABLE, D3DZB_TRUE)
        Call D3DDevice.SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        vVertexSize = Len(vSizeTextureVertexType)
        Call D3DDevice.SetTexture(0, vBackgroundTexture)
        Call D3DDevice.SetStreamSource(0, vVertexBuffer, vVertexSize)
        Call D3DDevice.DrawPrimitive(D3DPT_TRIANGLELIST, 0, 12)
        Call v3DModels.RenderState
        If vAnimationSpeed < GetTickCount Then
            vAngFromSun = vAngFromSun + 0.01
            vAngMoon = vAngMoon + 0.2
            If vAngMoon > 360 Then
                vAngMoon = vAngMoon - 360
            End If
            If vAngFromSun > 360 Then
                vAngFromSun = vAngFromSun - 360
            End If
            For vModelIndex = 0 To v3DModels.vModelCount - 1
                If vModelIndex > 0 Then
                    vecPlanet = v3DModels.Position(vModelIndex)
                    vecPlanet.X = (v3DModels.DistanceFromSun(vModelIndex) * Sin(vAngFromSun * Rad))
                    vecPlanet.z = (v3DModels.DistanceFromSun(vModelIndex) * Cos(vAngFromSun * Rad))
                    Call v3DModels.SetPosition(vModelIndex, vecPlanet)
                End If
                If vModelIndex <> 5 Then
                    Call v3DModels.Rotate(vModelIndex, 0, -0.001, 0)
                End If
            Next vModelIndex
            vModelIndex = v3DModels.ModelIndex("earth night.x")
            vecPlanet = v3DModels.Position(vModelIndex)
            vModelIndex = v3DModels.ModelIndex("moon.x")
            vDistanceFromEarth = 38
            vecMoon = v3DModels.Position(vModelIndex)
            vecMoon.X = vecPlanet.X + (vDistanceFromEarth * Sin(vAngMoon * Rad))
            vecMoon.z = vecPlanet.z + (vDistanceFromEarth * Cos(vAngMoon * Rad))
            Call v3DModels.SetPosition(vModelIndex, vecMoon)
            If vPlanetModelIndex <> 0 Then
                If vDistanceFromSun > v3DModels.DistanceFromSun(vPlanetModelIndex) Then
                    If vDistanceFromSun - v3DModels.DistanceFromSun(vPlanetModelIndex) > 200000 Then
                        VSpeedForward = 150
                    ElseIf vDistanceFromSun - v3DModels.DistanceFromSun(vPlanetModelIndex) > 100000 Then
                        VSpeedForward = 50
                    ElseIf vDistanceFromSun - v3DModels.DistanceFromSun(vPlanetModelIndex) > 10000 Then
                        VSpeedForward = 25
                    ElseIf vDistanceFromSun - v3DModels.DistanceFromSun(vPlanetModelIndex) > 2000 Then
                        VSpeedForward = 10
                    ElseIf vDistanceFromSun - v3DModels.DistanceFromSun(vPlanetModelIndex) > 100 Then
                        VSpeedForward = 0.2
                    Else
                        VSpeedForward = 0.02
                    End If
                ElseIf vDistanceFromSun < v3DModels.DistanceFromSun(vPlanetModelIndex) Then
                    If vDistanceFromSun - v3DModels.DistanceFromSun(vPlanetModelIndex) > -100 Then
                        VSpeedForward = 0.2
                    ElseIf vDistanceFromSun - v3DModels.DistanceFromSun(vPlanetModelIndex) > -2000 Then
                        VSpeedForward = 10
                    ElseIf vDistanceFromSun - v3DModels.DistanceFromSun(vPlanetModelIndex) > -10000 Then
                        VSpeedForward = 25
                    ElseIf vDistanceFromSun - v3DModels.DistanceFromSun(vPlanetModelIndex) > -100000 Then
                        VSpeedForward = 50
                    ElseIf vDistanceFromSun - v3DModels.DistanceFromSun(vPlanetModelIndex) > -200000 Then
                        VSpeedForward = 150
                    Else
                        VSpeedForward = 0.02
                    End If
                End If
            End If
            If vDistanceFromSun + 100 > 5890 And vDistanceFromSun <= 5890 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun + 100 > 10910 And vDistanceFromSun <= 10910 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun + 100 > 15050 And vDistanceFromSun <= 15050 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun + 100 > 15088 And vDistanceFromSun <= 15088 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun + 100 > 15110 And vDistanceFromSun <= 15110 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun + 100 > 15188 And vDistanceFromSun <= 15188 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun + 100 > 22880 And vDistanceFromSun <= 22880 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun + 800 > 77930 And vDistanceFromSun <= 77930 Then
                VSpeedForward = 0.2
            ElseIf vDistanceFromSun + 1200 > 142800 And vDistanceFromSun <= 142800 Then
                VSpeedForward = 1
            ElseIf vDistanceFromSun + 2500 > 287000 And vDistanceFromSun <= 287000 Then
                VSpeedForward = 1
            ElseIf vDistanceFromSun + 1500 > 449800 And vDistanceFromSun <= 449800 Then
                VSpeedForward = 1
            ElseIf vDistanceFromSun + 100 > 590100 And vDistanceFromSun <= 590100 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun - 100 < 5890 And vDistanceFromSun >= 5890 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun - 100 < 10910 And vDistanceFromSun >= 10910 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun - 100 < 15050 And vDistanceFromSun >= 15050 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun - 100 < 15088 And vDistanceFromSun >= 15088 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun - 100 < 15110 And vDistanceFromSun >= 15110 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun - 100 < 15188 And vDistanceFromSun >= 15188 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun - 100 < 22880 And vDistanceFromSun >= 22880 Then
                VSpeedForward = 0.02
            ElseIf vDistanceFromSun - 800 < 77930 And vDistanceFromSun >= 77930 Then
                VSpeedForward = 0.2
            ElseIf vDistanceFromSun - 1200 < 142800 And vDistanceFromSun >= 142800 Then
                VSpeedForward = 1
            ElseIf vDistanceFromSun - 2500 < 287000 And vDistanceFromSun >= 287000 Then
                VSpeedForward = 1
            ElseIf vDistanceFromSun - 1500 < 449800 And vDistanceFromSun >= 449800 Then
                VSpeedForward = 1
            ElseIf vDistanceFromSun - 100 < 590100 And vDistanceFromSun >= 590100 Then
                VSpeedForward = 0.02
            End If
            If vDistanceFromSun + VSpeedForward < v3DModels.DistanceFromSun(vPlanetModelIndex) Then
                vDistanceFromSun = vDistanceFromSun + VSpeedForward
                If vDistanceFromCamera < 1 Then
                    vDistanceFromCamera = vDistanceFromCamera + 0.001
                End If
            ElseIf vDistanceFromSun - VSpeedForward > v3DModels.DistanceFromSun(vPlanetModelIndex) Then
                vDistanceFromSun = vDistanceFromSun - VSpeedForward
                If vDistanceFromCamera < 1 Then
                    vDistanceFromCamera = vDistanceFromCamera + 0.001
                End If
            ElseIf vDistanceFromSun <> v3DModels.DistanceFromSun(vPlanetModelIndex) Then
                vDistanceFromSun = v3DModels.DistanceFromSun(vPlanetModelIndex)
                vPlanetChangeTime = GetTickCount + 600000
            ElseIf vPlanetChangeTime < GetTickCount Then
                If vPlanetModelIndex = 0 Then
                    vDistanceFromSun = 580100 * Rnd(100) + 3000
                    vDistanceFromCamera = 1
                End If
                vPlanetModelIndex = 11 * Rnd(100) + 1
            Else
                vSpeedForwardDistanceFromCamera = 0.001
                If vDistanceFromCamera + vSpeedForwardDistanceFromCamera < v3DModels.DistanceFromCamera(vPlanetModelIndex) Then
                    vDistanceFromCamera = vDistanceFromCamera + vSpeedForwardDistanceFromCamera
                ElseIf vDistanceFromCamera - vSpeedForwardDistanceFromCamera > v3DModels.DistanceFromCamera(vPlanetModelIndex) Then
                    vDistanceFromCamera = vDistanceFromCamera - vSpeedForwardDistanceFromCamera
                ElseIf vDistanceFromCamera <> v3DModels.DistanceFromCamera(vPlanetModelIndex) Then
                    vDistanceFromCamera = v3DModels.DistanceFromCamera(vPlanetModelIndex)
                End If
            End If
            If vPlanetModelIndex = 4 Then
                vCamPlace.X = (vDistanceFromSun - vDistanceFromEarth) * Sin((vAngFromSun + vDistanceFromCamera) * Rad) + vDistanceFromEarth * Sin(vAngMoon * Rad)
                vCamPlace.y = 0
                vCamPlace.z = (vDistanceFromSun - vDistanceFromEarth) * Cos((vAngFromSun + vDistanceFromCamera) * Rad) + vDistanceFromEarth * Cos(vAngMoon * Rad)
                vCamRotate.X = (vDistanceFromSun - vDistanceFromEarth) * Sin(vAngFromSun * Rad) + vDistanceFromEarth * Sin(vAngMoon * Rad)
                vCamRotate.y = 0
                vCamRotate.z = (vDistanceFromSun - vDistanceFromEarth) * Cos(vAngFromSun * Rad) + vDistanceFromEarth * Cos(vAngMoon * Rad)
            Else
                vCamPlace.X = vDistanceFromSun * Sin((vAngFromSun + vDistanceFromCamera) * Rad)
                vCamPlace.y = 0
                vCamPlace.z = vDistanceFromSun * Cos((vAngFromSun + vDistanceFromCamera) * Rad)
                vCamRotate.X = vDistanceFromSun * Sin(vAngFromSun * Rad)
                vCamRotate.y = 0
                vCamRotate.z = vDistanceFromSun * Cos(vAngFromSun * Rad)
            End If
            vAnimationSpeed = GetTickCount + 1
            Call v3DModels.CameraPos(vCamPlace, vCamRotate)
        End If
        For vModelIndex = 0 To v3DModels.vModelCount - 1
            Call v3DModels.Render(vModelIndex)
        Next vModelIndex
    End Sub

    Sub RenderControls()
        Dim iA As Integer, lA As Long, oA As RECT, iB As Byte, iC As Byte, iD As Integer
        If BFormIndex = 0 Then
            Exit Sub
        End If
        Call v3DModels.matrixTransforms
        Call D3DDevice.SetVertexShader(D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR)
        Call D3DDevice.SetRenderState(D3DRS_ZENABLE, False)
        Call D3DDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        For iC = 2 To BFormCount
            If DXForm(iC).Visible = True Then
                For iA = 0 To DXForm(iC).ControlsCount
                    If DXForm(iC).Controls(iA).Visible = True Then
                        If DXForm(iC).Controls(iA).Tag = "VBA.Time$" Then
                            DXForm(iC).Controls(iA).Text(0) = "Time: " & VBA.Time$
                        End If
                        If iFlashTimer < GetTickCount Then
                            iB = BFormIndex
                            BFormIndex = iC
                            If DXForm(iC).Controls(iA).Tag = "#Flash:0#" Then
                                DXForm(iC).Controls(iA).Tag = "#Flash:1#"
                                xForm(iC).Controls(iA).Value = 1
                            ElseIf DXForm(iC).Controls(iA).Tag = "#Flash:1#" Then
                                DXForm(iC).Controls(iA).Tag = "#Flash:0#"
                                xForm(iC).Controls(iA).Value = 0
                            End If
                            BFormIndex = iB
                        End If
                        If DXForm(iC).Controls(iA).Position(0).Left < CurrentWidth And DXForm(iC).Controls(iA).Position(0).Top < CurrentHeight Then
                            For iB = 0 To DXForm(iC).Controls(iA).PositionCount
                                If DXForm(iC).Controls(iA).Name = "List1" Then
                                    If iB = 13 Then
                                        Exit For
                                    End If
                                End If
                                If DXForm(iC).Controls(iA).Image(iB) <> 0 Then
                                    If DXForm(iC).Controls(iA).Image(iB) = 92 And iC = BFormIndex And iA = iFocus And iB = iFocus2 Then
                                        Call D3DDevice.SetTexture(0, D3DTexture(86))
                                    Else
                                        Call D3DDevice.SetTexture(0, D3DTexture(DXForm(iC).Controls(iA).Image(iB)))
                                    End If
                                    If DXForm(iC).Controls(iA).Image(iB) = 85 Then
                                        If Query(DXForm(iC).Controls(iA).DataBase).ListCount >= (DXForm(iC).Controls(iA).ScrollIndex + iB - 17) Then
                                            If Query(DXForm(iC).Controls(iA).DataBase).Selected(DXForm(iC).Controls(iA).ScrollIndex + iB - 17) = True Then
                                                lA = TextColor(92)
                                            Else
                                                lA = TextColor(DXForm(iC).Controls(iA).Image(iB))
                                            End If
                                        Else
                                            lA = TextColor(DXForm(iC).Controls(iA).Image(iB))
                                        End If
                                    Else
                                        lA = TextColor(DXForm(iC).Controls(iA).Image(iB))
                                    End If
                                    If Not D3DDevice.GetTexture(0) Is Nothing Then
                                        ColorVertex(0) = ColorVertexMake(CSng(DXForm(iC).Controls(iA).Position(iB).Left), CSng(DXForm(iC).Controls(iA).Position(iB).Top), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 0, 0)
                                        ColorVertex(1) = ColorVertexMake(CSng(DXForm(iC).Controls(iA).Position(iB).Right), CSng(DXForm(iC).Controls(iA).Position(iB).Top), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 1, 0)
                                        ColorVertex(2) = ColorVertexMake(CSng(DXForm(iC).Controls(iA).Position(iB).Left), CSng(DXForm(iC).Controls(iA).Position(iB).Bottom), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 0, 1)
                                        ColorVertex(3) = ColorVertexMake(CSng(DXForm(iC).Controls(iA).Position(iB).Right), CSng(DXForm(iC).Controls(iA).Position(iB).Bottom), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 1, 1)
                                        Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, ColorVertex(0), Len(ColorVertex(0)))
                                    End If
                                    Select Case DXForm(iC).Controls(iA).Image(iB)
                                        Case 74
                                            Call ScreenCursor(iC, iA, lA, iB)
                                        Case 9, 10, 61, 62, 63
                                            Call D3DX.DrawText(D3DFont, lA, DXForm(iC).Controls(iA).Text(0), DXForm(iC).Controls(iA).Position(iB), DT_VCENTER Or DT_CENTER)
                                        Case 84, 85, 86, 92, 93, 94, 95, 115, 116, 117
                                            Call D3DX.DrawText(D3DFont, lA, ScreenDataBase(iC, iA, "-1", iB), DXForm(iC).Controls(iA).Position(iB), DT_VCENTER Or DT_LEFT)
                                        Case Else
                                            If DXForm(iC).Controls(iA).Image(iB) = 48 Then
                                                Call D3DX.DrawText(D3DFont, lA, ScreenDataBase(iC, iA), DXForm(iC).Controls(iA).Position(iB), DT_TOP Or DT_CENTER)
                                            End If
                                    End Select
                                End If
                            Next iB
                        End If
                    End If
                Next iA
            End If
        Next iC
        If iFlashTimer < GetTickCount Then
            iFlashTimer = GetTickCount + 500
        End If
        Call D3DDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, False)
        If DXForm(1).Visible = True Then
            If DXForm(1).Controls(0).Visible = True Then
                If DXForm(1).Controls(0).Position(0).Left < CurrentWidth And DXForm(1).Controls(0).Position(0).Top < CurrentHeight Then
                    If DXForm(1).Controls(0).Image(0) <> 0 Then
                        Call D3DDevice.SetTexture(0, D3DTexture(DXForm(1).Controls(0).Image(0)))
                        If Not D3DDevice.GetTexture(0) Is Nothing Then
                            ColorVertex(0) = ColorVertexMake(CSng(DXForm(1).Controls(0).Position(0).Left), CSng(DXForm(1).Controls(0).Position(0).Top), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 0, 0)
                            ColorVertex(1) = ColorVertexMake(CSng(DXForm(1).Controls(0).Position(0).Right), CSng(DXForm(1).Controls(0).Position(0).Top), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 1, 0)
                            ColorVertex(2) = ColorVertexMake(CSng(DXForm(1).Controls(0).Position(0).Left), CSng(DXForm(1).Controls(0).Position(0).Bottom), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 0, 1)
                            ColorVertex(3) = ColorVertexMake(CSng(DXForm(1).Controls(0).Position(0).Right), CSng(DXForm(1).Controls(0).Position(0).Bottom), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 1, 1)
                            Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, ColorVertex(0), Len(ColorVertex(0)))
                        End If
                    End If
                End If
            End If
        End If
        For iC = 2 To BFormCount
            If DXForm(iC).Visible = True Then
                For iA = 0 To DXForm(iC).ControlsCount
                    If DXForm(iC).Controls(iA).Visible = True Then
                        If DXForm(iC).Controls(iA).Position(0).Left < CurrentWidth And DXForm(iC).Controls(iA).Position(0).Top < CurrentHeight Then
                            If DXForm(iC).Controls(iA).Name = "List1" Then
                                For iB = 13 To DXForm(iC).Controls(iA).PositionCount
                                    If DXForm(iC).Controls(iA).Image(iB) <> 0 Then
                                        If DXForm(iC).Controls(iA).Image(iB) = 92 And iC = BFormIndex And iA = iFocus And iB = iFocus2 Then
                                            Call D3DDevice.SetTexture(0, D3DTexture(86))
                                        Else
                                            Call D3DDevice.SetTexture(0, D3DTexture(DXForm(iC).Controls(iA).Image(iB)))
                                        End If
                                        If DXForm(iC).Controls(iA).Image(iB) = 85 Then
                                            If Query(DXForm(iC).Controls(iA).DataBase).ListCount >= (DXForm(iC).Controls(iA).ScrollIndex + iB - 17) Then
                                                If Query(DXForm(iC).Controls(iA).DataBase).Selected(DXForm(iC).Controls(iA).ScrollIndex + iB - 17) = True Then
                                                    lA = TextColor(92)
                                                Else
                                                    lA = TextColor(DXForm(iC).Controls(iA).Image(iB))
                                                End If
                                            Else
                                                lA = TextColor(DXForm(iC).Controls(iA).Image(iB))
                                            End If
                                        Else
                                            lA = TextColor(DXForm(iC).Controls(iA).Image(iB))
                                        End If
                                        If DXForm(iC).Controls(iA).Image(iB) <> 84 Then
                                            If Not D3DDevice.GetTexture(0) Is Nothing Then
                                                ColorVertex(0) = ColorVertexMake(CSng(DXForm(iC).Controls(iA).Position(iB).Left), CSng(DXForm(iC).Controls(iA).Position(iB).Top), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 0, 0)
                                                ColorVertex(1) = ColorVertexMake(CSng(DXForm(iC).Controls(iA).Position(iB).Right), CSng(DXForm(iC).Controls(iA).Position(iB).Top), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 1, 0)
                                                ColorVertex(2) = ColorVertexMake(CSng(DXForm(iC).Controls(iA).Position(iB).Left), CSng(DXForm(iC).Controls(iA).Position(iB).Bottom), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 0, 1)
                                                ColorVertex(3) = ColorVertexMake(CSng(DXForm(iC).Controls(iA).Position(iB).Right), CSng(DXForm(iC).Controls(iA).Position(iB).Bottom), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 1, 1)
                                                Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, ColorVertex(0), Len(ColorVertex(0)))
                                            End If
                                        End If
                                        Select Case DXForm(iC).Controls(iA).Image(iB)
                                            Case 84, 85, 86, 92, 93, 94, 95, 115, 116, 117
                                                Call D3DX.DrawText(D3DFont, lA, ScreenDataBase(iC, iA, "-1", iB), DXForm(iC).Controls(iA).Position(iB), DT_TOP Or DT_LEFT)
                                        End Select
                                    End If
                                Next iB
                            End If
                        End If
                    End If
                Next iA
            End If
        Next iC
        For iA = 0 To TextSelectedCount
            If TextSelected(iA).Visible = True Then
                Call D3DDevice.SetTexture(0, D3DTexture(85))
                If Not D3DDevice.GetTexture(0) Is Nothing Then
                    ColorVertex(0) = ColorVertexMake(CSng(TextSelected(iA).Left), CSng(TextSelected(iA).Top), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 0, 0)
                    ColorVertex(1) = ColorVertexMake(CSng(TextSelected(iA).Right - 1), CSng(TextSelected(iA).Top), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 1, 0)
                    ColorVertex(2) = ColorVertexMake(CSng(TextSelected(iA).Left), CSng(TextSelected(iA).Bottom - 1), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 0, 1)
                    ColorVertex(3) = ColorVertexMake(CSng(TextSelected(iA).Right - 1), CSng(TextSelected(iA).Bottom - 1), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 1, 1)
                    Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, ColorVertex(0), Len(ColorVertex(0)))
                End If
                oA.Left = TextSelected(iA).Left
                oA.Top = TextSelected(iA).Top
                oA.Right = TextSelected(iA).Right
                oA.Bottom = TextSelected(iA).Bottom
                Call D3DX.DrawText(D3DFont, &HFFFFF000, TextSelected(iA).Text, oA, DT_TOP Or DT_LEFT)
            End If
        Next iA
        For iC = 1 To BFormCount
            If DXForm(iC).Visible = True Then
                For iA = 0 To DXForm(iC).ControlsCount
                    If DXForm(iC).Controls(iA).Visible = True Then
                        If DXForm(iC).Controls(iA).Position(0).Left < CurrentWidth And DXForm(iC).Controls(iA).Position(0).Top < CurrentHeight Then
                            For iD = 1 To DXForm(iC).Controls(iA).EmosCount
                                If DXForm(iC).Controls(iA).Emos(iD).Visible = True Then
                                    Call D3DDevice.SetTexture(0, D3DTexture(DXForm(iC).Controls(iA).Emos(iD).Image))
                                    If Not D3DDevice.GetTexture(0) Is Nothing Then
                                        ColorVertex(0) = ColorVertexMake(CSng(DXForm(iC).Controls(iA).Emos(iD).Position.Left), CSng(DXForm(iC).Controls(iA).Emos(iD).Position.Top), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 0, 0)
                                        ColorVertex(1) = ColorVertexMake(CSng(DXForm(iC).Controls(iA).Emos(iD).Position.Right), CSng(DXForm(iC).Controls(iA).Emos(iD).Position.Top), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 1, 0)
                                        ColorVertex(2) = ColorVertexMake(CSng(DXForm(iC).Controls(iA).Emos(iD).Position.Left), CSng(DXForm(iC).Controls(iA).Emos(iD).Position.Bottom), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 0, 1)
                                        ColorVertex(3) = ColorVertexMake(CSng(DXForm(iC).Controls(iA).Emos(iD).Position.Right), CSng(DXForm(iC).Controls(iA).Emos(iD).Position.Bottom), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 1, 1)
                                        Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, ColorVertex(0), Len(ColorVertex(0)))
                                    End If
                                End If
                            Next iD
                        End If
                    End If
                Next iA
            End If
        Next iC
        If BComboList.Visible = True Then
            For iB = 0 To BComboList.PositionCount
                If BComboList.Image(iB) <> 0 Then
                    Call D3DDevice.SetTexture(0, D3DTexture(BComboList.Image(iB)))
                    lA = TextColor(BComboList.Image(iB))
                    If Not D3DDevice.GetTexture(0) Is Nothing Then
                        ColorVertex(0) = ColorVertexMake(CSng(BComboList.Position(iB).Left), CSng(BComboList.Position(iB).Top), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 0, 0)
                        ColorVertex(1) = ColorVertexMake(CSng(BComboList.Position(iB).Right), CSng(BComboList.Position(iB).Top), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 1, 0)
                        ColorVertex(2) = ColorVertexMake(CSng(BComboList.Position(iB).Left), CSng(BComboList.Position(iB).Bottom), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 0, 1)
                        ColorVertex(3) = ColorVertexMake(CSng(BComboList.Position(iB).Right), CSng(BComboList.Position(iB).Bottom), 0, 1, D3DColorRGBA(255, 255, 255, 255), 0, 1, 1)
                        Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, ColorVertex(0), Len(ColorVertex(0)))
                    End If
                    Select Case BComboList.Image(iB)
                        Case 84, 85, 86, 92, 93, 94, 95, 115, 116, 117
                            Call D3DX.DrawText(D3DFont, lA, ScreenDataBase(0, 0, "-1", iB), BComboList.Position(iB), DT_TOP Or DT_LEFT)
                    End Select
                End If
            Next iB
        End If
    End Sub

    Sub RenderScreen()
        Dim Result As Long
        Static bSurface As Direct3DSurface8
        On Local Error GoTo DeviceProblem
        Call D3DDevice.Clear(0, 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, RGB(0, 0, 0), 1.0#, 0)
        Call D3DDevice.BeginScene
        Call RenderBackground()
        Call RenderControls()
        On Local Error GoTo NoEnding
        Call D3DDevice.EndScene
        On Local Error GoTo DeviceProblem
        Call D3DDevice.Present(0, 0, 0, 0)
        Exit Sub
DeviceProblem:
        Select Case Err.Number
            Case D3D_OK
            Case D3DERR_DEVICELOST
                Result = 0
                Do Until Result = D3DERR_DEVICENOTRESET
                    DoEvents
                    Result = D3DDevice.TestCooperativeLevel
                Loop
                Call aScreen.LoadScreen()
            Case D3DERR_DEVICENOTRESET
                Call aScreen.LoadScreen()
        End Select
NoEnding:
        Exit Sub
    End Sub

    Sub ScreenComboList()
        Dim iA As Double
        If Query(BComboList.DataBase).ListCount >= 0 Then
            If Query(BComboList.DataBase).ListCount > 0 Then
                BComboList.Position(11).Top = BComboList.Position(10).Top + ((BComboList.Position(10).Bottom - BComboList.Position(10).Top - (10 * FixHeight)) / Query(BComboList.DataBase).ListCount) * BComboList.ListIndex
                BComboList.Position(11).Bottom = BComboList.Position(11).Top + (10 * FixHeight)
            End If
            If (BComboList.ListIndex - BComboList.ScrollIndex) > CInt(BComboList.PositionCount - 17) Then
                iA = BComboList.ListIndex
                Do Until iA <= CInt(BComboList.PositionCount - 17)
                    If CInt(BComboList.PositionCount - 17) < 0 Then
                        Exit Do
                    End If
                    iA = iA - 1
                    DoEvents
                Loop
                BComboList.ScrollIndex = BComboList.ListIndex - iA
            ElseIf (BComboList.ListIndex - BComboList.ScrollIndex) <= 0 Then
                BComboList.ScrollIndex = BComboList.ListIndex
            End If
            If BComboList.ScrollIndex > (Query(BComboList.DataBase).ListCount - CInt(BComboList.PositionCount - 17) + 1) Then
                BComboList.ScrollIndex = (Query(BComboList.DataBase).ListCount - CInt(BComboList.PositionCount - 17))
            ElseIf BComboList.ScrollIndex < 0 Then
                BComboList.ScrollIndex = 0
            End If
        End If
        For iA = 17 To BComboList.PositionCount
            If BComboList.Image(iA) = 84 Or BComboList.Image(iA) = 85 Or BComboList.Image(iA) = 86 Or BComboList.Image(iA) = 92 Then
                If BComboList.ScrollIndex + iA - 17 > Query(BComboList.DataBase).ListCount Then
                    BComboList.Image(iA) = 84
                ElseIf iA - 17 + BComboList.ScrollIndex = BComboList.ListIndex Then
                    BComboList.Image(iA) = 86
                    iFocus2 = iA
                Else
                    BComboList.Image(iA) = 84
                End If
            End If
        Next iA
        DXForm(BFormIndex).Controls(iFocus).Text(0) = Query(BComboList.DataBase).List(BComboList.ListIndex)
    End Sub

    Function ScreenDataBase(iC As Byte, iA As Integer, Optional sA As String = "-1", Optional iB As Byte = 0) As String
        If iC = 0 And iA = 0 Then
            If iB > 16 And BComboList.ScrollIndex + iB - 17 <= Query(BComboList.DataBase).ListCount And BComboList.ScrollIndex + iB - 17 > -1 And Query(BComboList.DataBase).ListCount > 0 Then
                ScreenDataBase = AndDouble(Query(BComboList.DataBase).List(BComboList.ScrollIndex + iB - 17))
            End If
        ElseIf iA >= 0 Then
            If sA = "-1" Then
                Select Case DXForm(iC).Controls(iA).Name
                    Case "List1"
                        If iB > 16 And DXForm(iC).Controls(iA).ScrollIndex + iB - 17 <= Query(DXForm(iC).Controls(iA).DataBase).ListCount And DXForm(iC).Controls(iA).ScrollIndex + iB - 17 > -1 And Query(DXForm(iC).Controls(iA).DataBase).ListCount > 0 Then
                            ScreenDataBase = AndDouble(Query(DXForm(iC).Controls(iA).DataBase).List(DXForm(iC).Controls(iA).ScrollIndex + iB - 17))
                        End If
                    Case "Combo1"
                        ScreenDataBase = AndDouble(DXForm(iC).Controls(iA).Text(0))
                    Case "Data1"
                        Select Case DXForm(iC).Controls(iA).Image(iB)
                            Case 9, 74, 116, 117
                                If DXForm(iC).Controls(iA).PositionCount = 2 Then
                                    ''If DXForm(iC).Controls(iA).EmosText = Empty Then
                                    'DXForm(iC).Controls(iA).EmosText = ScrollText(DXForm(iC).Controls(iA).Text(0), DXForm(iC).Controls(iA))
                                    ''End If
                                    ScreenDataBase = DXForm(iC).Controls(iA).EmosText
                                Else
                                    If DXForm(iC).Controls(iA).ScrollIndex + iB <= UBound(DXForm(iC).Controls(iA).Text()) Then
                                        If (DXForm(iC).Controls(iA).ScrollIndex + iB) = 0 Then
                                            ScreenDataBase = DXForm(iC).Controls(iA).EmosText
                                        Else
                                            ScreenDataBase = DXForm(iC).Controls(iA).Text(DXForm(iC).Controls(iA).ScrollIndex + iB)
                                        End If
                                    Else
                                        ScreenDataBase = Empty
                                    End If
                                End If
                            Case 93
                                If Val(DXForm(iC).Controls(iA).DataField) >= 0 And DXForm(iC).Controls(iA).DataBase <> 0 Then
                                    ScreenDataBase = DXForm(iC).Controls(iA).Text(iB)
                                Else
                                    ScreenDataBase = DXForm(iC).Controls(iA).Tag
                                End If
                        End Select
                    Case "Command1"
                        ScreenDataBase = DXForm(iC).Controls(iA).Text(0)
                    Case "Text1"
                        If DXForm(iC).Controls(iA).PasswordChar = "" Then
                            If DXForm(iC).Controls(iA).EmosText = Empty Then
                                DXForm(iC).Controls(iA).EmosText = ScrollText(DXForm(iC).Controls(iA).Text(0), DXForm(iC).Controls(iA))
                            End If
                            ScreenDataBase = DXForm(iC).Controls(iA).EmosText
                        Else
                            ScreenDataBase = Replace(Space(Len(DXForm(iC).Controls(iA).Text(0))), " ", DXForm(iC).Controls(iA).PasswordChar)
                        End If
                    Case Else
                        Select Case DXForm(iC).Controls(iA).FillStyle
                            Case 0, 5
                                ScreenDataBase = DXForm(iC).Controls(iA).Text(0)
                            Case 1
                                ScreenDataBase = DXForm(iC).Controls(iA).Text(0)
                            Case 2
                                If DXForm(iC).Controls(iA).Name = "List1" Then
                                    If iB > 16 And DXForm(iC).Controls(iA).ScrollIndex + iB - 17 <= Query(DXForm(iC).Controls(iA).DataBase).ListCount And DXForm(iC).Controls(iA).ScrollIndex + iB - 17 > -1 Then
                                        ScreenDataBase = Query(DXForm(iC).Controls(iA).DataBase).List(DXForm(iC).Controls(iA).ScrollIndex + iB - 17)
                                    End If
                                Else
                                    ScreenDataBase = Query(DXForm(iC).Controls(iA).DataSource).List(DXForm(iC).Controls(iA).ScrollIndex)
                                End If
                            Case 4
                                ScreenDataBase = DXForm(iC).Controls(iA).Text(0)
                        End Select
                End Select
            Else
                Select Case DXForm(iC).Controls(iA).FillStyle
                    Case 0, 5
                        DXForm(iC).Controls(iA).Text(0) = sA
                    Case 1
                        DXForm(iC).Controls(iA).Text(0) = sA
                    Case 2
                        Query(DXForm(iC).Controls(iA).DataBase).List(DXForm(iC).Controls(iA).ScrollIndex + DXForm(iC).Controls(iA).Index) = sA
                    Case 4
                        DXForm(iC).Controls(iA).Text(0) = sA
                End Select
            End If
        End If
    End Function

    Sub ScreenListBox()
        Dim iA As Double
        If Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount >= 0 Then
            If Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount > 0 Then
                DXForm(BFormIndex).Controls(iFocus).Position(11).Top = DXForm(BFormIndex).Controls(iFocus).Position(10).Top + ((DXForm(BFormIndex).Controls(iFocus).Position(10).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(10).Top - (10 * FixHeight)) / Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount) * DXForm(BFormIndex).Controls(iFocus).ListIndex
                DXForm(BFormIndex).Controls(iFocus).Position(11).Bottom = DXForm(BFormIndex).Controls(iFocus).Position(11).Top + (10 * FixHeight)
            End If
            If (DXForm(BFormIndex).Controls(iFocus).ListIndex - DXForm(BFormIndex).Controls(iFocus).ScrollIndex) > CInt(DXForm(BFormIndex).Controls(iFocus).PositionCount - 17) Then
                iA = DXForm(BFormIndex).Controls(iFocus).ListIndex
                Do Until iA <= CInt(DXForm(BFormIndex).Controls(iFocus).PositionCount - 17)
                    If CInt(DXForm(BFormIndex).Controls(iFocus).PositionCount - 17) < 0 Then
                        Exit Do
                    End If
                    iA = iA - 1
                    DoEvents
                Loop
                DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex - iA
            ElseIf (DXForm(BFormIndex).Controls(iFocus).ListIndex - DXForm(BFormIndex).Controls(iFocus).ScrollIndex) <= 0 Then
                DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex
            End If
            If (Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount - (DXForm(BFormIndex).Controls(iFocus).PositionCount - 17)) < 0 Then
                DXForm(BFormIndex).Controls(iFocus).ScrollIndex = 0
            ElseIf DXForm(BFormIndex).Controls(iFocus).ScrollIndex > (Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount - (DXForm(BFormIndex).Controls(iFocus).PositionCount - 17)) Then
                DXForm(BFormIndex).Controls(iFocus).ScrollIndex = (Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount - CInt(DXForm(BFormIndex).Controls(iFocus).PositionCount - 17))
            ElseIf DXForm(BFormIndex).Controls(iFocus).ScrollIndex < 0 Then
                DXForm(BFormIndex).Controls(iFocus).ScrollIndex = 0
            End If
        End If
        For iA = 17 To DXForm(BFormIndex).Controls(iFocus).PositionCount
            If DXForm(BFormIndex).Controls(iFocus).Image(iA) = 84 Or DXForm(BFormIndex).Controls(iFocus).Image(iA) = 85 Or DXForm(BFormIndex).Controls(iFocus).Image(iA) = 86 Or DXForm(BFormIndex).Controls(iFocus).Image(iA) = 92 Then
                If DXForm(BFormIndex).Controls(iFocus).ScrollIndex + iA - 17 > Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount Then
                    DXForm(BFormIndex).Controls(iFocus).Image(iA) = 84
                ElseIf iA - 17 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex Then
                    If Query(DXForm(BFormIndex).Controls(iFocus).DataBase).Selected(DXForm(BFormIndex).Controls(iFocus).ScrollIndex + iA - 17) = True Then
                        DXForm(BFormIndex).Controls(iFocus).Image(iA) = 92
                    Else
                        DXForm(BFormIndex).Controls(iFocus).Image(iA) = 86
                    End If
                    iFocus2 = iA
                ElseIf Query(DXForm(BFormIndex).Controls(iFocus).DataBase).Selected(DXForm(BFormIndex).Controls(iFocus).ScrollIndex + iA - 17) = True Then
                    DXForm(BFormIndex).Controls(iFocus).Image(iA) = 92
                Else
                    DXForm(BFormIndex).Controls(iFocus).Image(iA) = 84
                End If
            End If
        Next iA
    End Sub

    Function ScrollText(vText As String, BControl As JControl) As String
        Dim sSplitIndex As Long, sTijdelijkeTekst As String
        'verdeel de tekst
        If BControl.MultiLine = True Then
            sTijdelijkeTekst = TextVerdelen(vText, xForm(BFormIndex).Font, CLng((BControl.Position(0).Right - BControl.Position(0).Left)), BControl.MultiLine)
            'sla de tekst op voor de cursor om de positie te bepalen
            If TextCursorString <> sTijdelijkeTekst And TextCursorString <> Empty Then
                sTijdelijkeTekst = TextCursorString
            Else
                TextCursorString = sTijdelijkeTekst
            End If
            If BControl.ListCount = 0 And sTijdelijkeTekst <> "" Then
                BControl.ListIndex = 0
                BControl.ListCount = UBound(Split(sTijdelijkeTekst, vbCrLf)) + 1
                BControl.ScrollIndex = 0
                BControl.ScrollCount = Val((BControl.Position(0).Bottom - BControl.Position(0).Top) / (vFontSize)) + 1
            End If
            If BControl.ListCount <> UBound(Split(sTijdelijkeTekst, vbCrLf)) + 1 Then
                BControl.ListCount = UBound(Split(sTijdelijkeTekst, vbCrLf)) + 1
            End If
            If BControl.ListCount < BControl.ScrollCount Then
                BControl.ScrollCount = BControl.ListCount
            End If
            If BControl.ListCount > 0 Then
                'verdeel de text over het textveld en zorg dat je kan scrollen van boven naar beneden.
                For sSplitIndex = BControl.ListIndex To BControl.ListIndex + BControl.ScrollCount Step 1
                    If BControl.ListCount >= BControl.ScrollCount And sSplitIndex < BControl.ListCount Then
                        ScrollText = ScrollText & Split(sTijdelijkeTekst, vbCrLf)(sSplitIndex) & vbCrLf
                    ElseIf sSplitIndex <= BControl.ScrollCount - BControl.ListCount Then
                        ScrollText = ScrollText & Split(sTijdelijkeTekst, vbCrLf)(sSplitIndex) & vbCrLf
                    ElseIf sSplitIndex = BControl.ListCount And ScrollText = "" Then
                        ScrollText = sTijdelijkeTekst
                        Exit For
                    End If
                Next sSplitIndex
            End If
        End If
    End Function

    Sub SetTextColor()
        TextColor(9) = ColorRed
        TextColor(10) = ColorLightGreen
        TextColor(74) = ColorYellow
        TextColor(84) = ColorWhite
        TextColor(85) = ColorYellow
        TextColor(86) = ColorRed
        TextColor(92) = ColorBlue
        TextColor(93) = ColorWhite
        TextColor(94) = ColorWhite
        TextColor(95) = ColorBlack
        TextColor(61) = ColorBlack
        TextColor(62) = ColorBlack
        TextColor(63) = ColorBlack
        TextColor(48) = ColorYellow
        TextColor(115) = ColorWhite
        TextColor(116) = ColorYellow
        TextColor(117) = ColorRed
    End Sub

End Module
