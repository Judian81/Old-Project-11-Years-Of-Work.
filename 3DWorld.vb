Module a3DWorld

    Private Structure modelType
        Public ModelMesh As D3DXBaseMesh
        Public nMaterials As Long
        Public MeshTextures() As Direct3DTexture8
        Public MeshMaterials() As D3DMATERIAL8
        Public matScale As D3DMATRIX
        Public matRotation As D3DMATRIX
        Public matPosition As D3DMATRIX
        Public vPosition As D3DVECTOR
        Public vRotation As D3DVECTOR
        Public vModelName As String
        Public vDistanceFromSun As Single
        Public vDistanceFromCamera As Single
        Public vRotateSpeed As Single
    End Structure

    Private vModel() As modelType

    Private Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

    Public vModelCount As Integer

    Private Function D3DVec(ByVal X As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR
        D3DVec.X = X
        D3DVec.y = y
        D3DVec.z = z
    End Function

    Private Function CreateTexVert(ByVal X As Single, ByVal y As Single, ByVal z As Single, ByVal tv As Single, ByVal tu As Single) As TextureVertexType
        CreateTexVert.X = X
        CreateTexVert.y = y
        CreateTexVert.z = z
        CreateTexVert.tu = tu
        CreateTexVert.tv = tv
    End Function

    Sub CameraPos(vCamPlace As D3DVECTOR, vCamRotate As D3DVECTOR)
        Dim matView As D3DMATRIX
        Call D3DXMatrixLookAtLH(matView, vCamPlace, vCamRotate, D3DVec(0#, Cos(0), 0#))
        Call D3DDevice.SetTransform(D3DTS_VIEW, matView)
    End Sub

    Sub matrixTransforms()
        Dim matWorld As D3DMATRIX
        Call D3DXMatrixIdentity(matWorld)
        Call D3DDevice.SetTransform(D3DTS_WORLD, matWorld)
    End Sub

    Function DistanceFromSun(vModelIndex As Integer) As Single
        DistanceFromSun = vModel(vModelIndex).vDistanceFromSun
    End Function

    Function RotateSpeed(vModelIndex As Integer) As Single
        RotateSpeed = vModel(vModelIndex).vRotateSpeed
    End Function

    Function DistanceFromCamera(vModelIndex As Integer, Optional vDistance As Single = 0) As Single
        If vDistance = 0 Then
            DistanceFromCamera = vModel(vModelIndex).vDistanceFromCamera
        Else
            vModel(vModelIndex).vDistanceFromCamera = vDistance
        End If
    End Function

    Function ModelIndex(vModelName As String) As Integer
        Dim vA As Integer
        For vA = 0 To vModelCount - 1
            If vModelName = vModel(vA).vModelName Then
                ModelIndex = vA
                Exit For
            End If
        Next vA
    End Function

    Function Position(vModelIndex As Integer) As D3DVECTOR
        Position = vModel(vModelIndex).vPosition
    End Function

    Sub SetPosition(vModelIndex As Integer, vPosition As D3DVECTOR)
        vModel(vModelIndex).vPosition.X = vPosition.X
        vModel(vModelIndex).vPosition.y = vPosition.y
        vModel(vModelIndex).vPosition.z = vPosition.z
        Call D3DXMatrixIdentity(vModel(vModelIndex).matPosition)
        Call D3DXMatrixTranslation(vModel(vModelIndex).matPosition, vModel(vModelIndex).vPosition.X, vModel(vModelIndex).vPosition.y, vModel(vModelIndex).vPosition.z)
    End Sub

    Sub MoveCamera(vRotation As D3DVECTOR, vLocation As D3DVECTOR)
        Dim matRotate As D3DMATRIX, matView As D3DMATRIX, matLocation As D3DMATRIX
        Call D3DXMatrixTranslation(matLocation, vLocation.X, vLocation.y, vLocation.z)
        Call D3DXMatrixRotationYawPitchRoll(matRotate, vRotation.y, vRotation.X, vRotation.z)
        Call D3DXMatrixMultiply(matView, matLocation, matRotate)
        Call D3DDevice.SetTransform(D3DTS_VIEW, matView)
    End Sub

    Public Sub CreateWorld()
        Dim matView As D3DMATRIX, matProj As D3DMATRIX, Material As D3DMATERIAL8, Light As D3DLIGHT8
        Call D3DXMatrixLookAtLH(matView, D3DVec(0, 0, 0), D3DVec(0#, 0#, 1), D3DVec(0#, 1.0#, 0#))
        vCamPlace = D3DVec(0, 0, 0)
        vCamRotate = D3DVec(0, 0, 0)
        Call D3DDevice.SetTransform(D3DTS_VIEW, matView)
        Call D3DXMatrixPerspectiveFovLH(matProj, PI / 4, 1, 6, 3000000)
        Call D3DDevice.SetTransform(D3DTS_PROJECTION, matProj)

        vWorldBackground(0) = CreateTexVert(-1000000, 1000000, -1000000, 0, 0)
        vWorldBackground(1) = CreateTexVert(1000000, 1000000, -1000000, 1, 0)
        vWorldBackground(2) = CreateTexVert(-1000000, 1000000, 1000000, 0, 1)

        vWorldBackground(3) = CreateTexVert(1000000, 1000000, -1000000, 1, 0)
        vWorldBackground(4) = CreateTexVert(1000000, 1000000, 1000000, 1, 1)
        vWorldBackground(5) = CreateTexVert(-1000000, 1000000, 1000000, 0, 1)

        vWorldBackground(6) = CreateTexVert(-1000000, -1000000, -1000000, 0, 0)
        vWorldBackground(7) = CreateTexVert(1000000, -1000000, -1000000, 1, 0)
        vWorldBackground(8) = CreateTexVert(-1000000, -1000000, 1000000, 0, 1)

        vWorldBackground(9) = CreateTexVert(1000000, -1000000, -1000000, 1, 0)
        vWorldBackground(10) = CreateTexVert(1000000, -1000000, 1000000, 1, 1)
        vWorldBackground(11) = CreateTexVert(-1000000, -1000000, 1000000, 0, 1)

        vWorldBackground(12) = CreateTexVert(-1000000, 1000000, -1000000, 0, 0)
        vWorldBackground(13) = CreateTexVert(-1000000, 1000000, 1000000, 1, 0)
        vWorldBackground(14) = CreateTexVert(-1000000, -1000000, -1000000, 0, 1)

        vWorldBackground(15) = CreateTexVert(-1000000, 1000000, 1000000, 1, 0)
        vWorldBackground(16) = CreateTexVert(-1000000, -1000000, 1000000, 1, 1)
        vWorldBackground(17) = CreateTexVert(-1000000, -1000000, -1000000, 0, 1)

        vWorldBackground(18) = CreateTexVert(1000000, 1000000, -1000000, 0, 0)
        vWorldBackground(19) = CreateTexVert(1000000, 1000000, 1000000, 1, 0)
        vWorldBackground(20) = CreateTexVert(1000000, -1000000, -1000000, 0, 1)

        vWorldBackground(21) = CreateTexVert(1000000, 1000000, 1000000, 1, 0)
        vWorldBackground(22) = CreateTexVert(1000000, -1000000, 1000000, 1, 1)
        vWorldBackground(23) = CreateTexVert(1000000, -1000000, -1000000, 0, 1)

        vWorldBackground(24) = CreateTexVert(-1000000, 1000000, 1000000, 0, 0)
        vWorldBackground(25) = CreateTexVert(1000000, 1000000, 1000000, 1, 0)
        vWorldBackground(26) = CreateTexVert(-1000000, -1000000, 1000000, 0, 1)

        vWorldBackground(27) = CreateTexVert(1000000, 1000000, 1000000, 1, 0)
        vWorldBackground(28) = CreateTexVert(1000000, -1000000, 1000000, 1, 1)
        vWorldBackground(29) = CreateTexVert(-1000000, -1000000, 1000000, 0, 1)

        vWorldBackground(30) = CreateTexVert(-1000000, 1000000, -1000000, 0, 0)
        vWorldBackground(31) = CreateTexVert(1000000, 1000000, -1000000, 1, 0)
        vWorldBackground(32) = CreateTexVert(-1000000, -1000000, -1000000, 0, 1)

        vWorldBackground(33) = CreateTexVert(1000000, 1000000, -1000000, 1, 0)
        vWorldBackground(34) = CreateTexVert(1000000, -1000000, -1000000, 1, 1)
        vWorldBackground(35) = CreateTexVert(-1000000, -1000000, -1000000, 0, 1)

        D3DDevice.SetVertexShader FVF_TEX

    D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    D3DDevice.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

     vVertexBuffer = D3DDevice.CreateVertexBuffer(Len(vWorldBackground(0)) * 36, 0, FVF_TEX, D3DPOOL_DEFAULT)
        Call D3DVertexBuffer8SetData(vVertexBuffer, 0, Len(vWorldBackground(0)) * 36, 0, vWorldBackground(0))
    End Sub

    Public Sub Render(vModelIndex)
        On Local Error Resume Next
        Dim vA As Integer, matObject As D3DMATRIX
        If vModel(vModelIndex).nMaterials = 0 Then
            Call D3DXMatrixIdentity(matObject)
            Call D3DXMatrixMultiply(matObject, vModel(vModelIndex).matRotation, vModel(vModelIndex).matPosition)
            Call D3DXMatrixMultiply(matObject, matObject, vModel(vModelIndex).matScale)
            Call D3DDevice.SetTransform(D3DTS_WORLD, matObject)
            Call vModel(vModelIndex).ModelMesh.DrawSubset(0)
        Else
            For vA = 0 To vModel(vModelIndex).nMaterials
                Call D3DXMatrixIdentity(matObject)
                Call D3DXMatrixMultiply(matObject, vModel(vModelIndex).matRotation, vModel(vModelIndex).matPosition)
                Call D3DXMatrixMultiply(matObject, vModel(vModelIndex).matScale, matObject)
                Call D3DDevice.SetTransform(D3DTS_WORLD, matObject)
                Call D3DDevice.SetTexture(0, vModel(vModelIndex).MeshTextures(vA))
                Call D3DDevice.SetMaterial(vModel(vModelIndex).MeshMaterials(vA))
                Call vModel(vModelIndex).ModelMesh.DrawSubset(vA)
            Next vA
        End If
    End Sub

    Public Sub RenderState()
        Call D3DDevice.SetVertexShader(FVF_VERTEX)
        Call D3DDevice.SetRenderState(D3DRS_LIGHTING, 0)
        Call D3DDevice.SetRenderState(D3DRS_ZENABLE, D3DZB_TRUE)
        Call D3DDevice.SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        Call D3DDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, False)
    End Sub

    Public Function Add(vModelName As String, vScale As Single) As Integer
        On Local Error Resume Next
        Dim vMaterialsBuffer As D3DXBuffer, vTextureFolder As String, vTextureName As String, vA As Long
        vModelCount += 1
        vTextureFolder = VBA.Left(vModelName, InStrRev(vModelName, "\"))
        ReDim Preserve vModel(vModelCount - 1)

        vModel(vModelCount - 1).ModelMesh = D3DX.LoadMeshFromX(vModelName, D3DXMESH_MANAGED, D3DDevice, Nothing, vMaterialsBuffer, vModel(vModelCount - 1).nMaterials)
        ReDim vModel(vModelCount - 1).MeshMaterials(vModel(vModelCount - 1).nMaterials)
        ReDim vModel(vModelCount - 1).MeshTextures(vModel(vModelCount - 1).nMaterials)
        vModel(vModelCount - 1).vModelName = LCase(VBA.Mid(vModelName, InStrRev(vModelName, "\") + 1))
        Call D3DXMatrixScaling(vModel(vModelCount - 1).matScale, vScale, vScale, vScale)
        For vA = 0 To vModel(vModelCount - 1).nMaterials - 1
            Call D3DX.BufferGetMaterial(vMaterialsBuffer, vA, vModel(vModelCount - 1).MeshMaterials(vA))
            vTextureName = D3DX.BufferGetTextureName(vMaterialsBuffer, vA)
            vModel(vModelCount - 1).MeshMaterials(vA).Ambient = vModel(vModelCount - 1).MeshMaterials(vA).diffuse
            If vTextureName <> "" Then
                vModel(vModelCount - 1).MeshTextures(vA) = D3DX.CreateTextureFromFileEx(D3DDevice, vTextureFolder & vTextureName, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, 0, 0)
            End If
        Next vA
        Add = vModelCount - 1
    End Function

    Public Sub Move(vModelIndex As Integer, ByVal X As Single, ByVal y As Single, ByVal z As Single)
        vModel(vModelIndex).vPosition.X = vModel(vModelIndex).vPosition.X + X
        vModel(vModelIndex).vPosition.y = vModel(vModelIndex).vPosition.y + y
        vModel(vModelIndex).vPosition.z = vModel(vModelIndex).vPosition.z + z
        vModel(vModelIndex).vDistanceFromSun = z
        Call D3DXMatrixIdentity(vModel(vModelIndex).matPosition)
        Call D3DXMatrixTranslation(vModel(vModelIndex).matPosition, vModel(vModelIndex).vPosition.X, vModel(vModelIndex).vPosition.y, vModel(vModelIndex).vPosition.z)
    End Sub

    Public Sub Rotate(vModelIndex As Integer, ByVal AngleX As Single, ByVal AngleY As Single, ByVal AngleZ As Single)
        Dim vMatrixCalcAdd As D3DMATRIX, vMatrixBuffer As D3DMATRIX
        Call D3DXMatrixIdentity(vMatrixBuffer)
        vModel(vModelIndex).vRotation.X = vModel(vModelIndex).vRotation.X + AngleX
        Call D3DXMatrixRotationX(vMatrixCalcAdd, vModel(vModelIndex).vRotation.X)
        Call D3DXMatrixMultiply(vMatrixBuffer, vMatrixBuffer, vMatrixCalcAdd)
        vModel(vModelIndex).vRotation.y = vModel(vModelIndex).vRotation.y + AngleY
        Call D3DXMatrixRotationY(vMatrixCalcAdd, vModel(vModelIndex).vRotation.y)
        Call D3DXMatrixMultiply(vMatrixBuffer, vMatrixBuffer, vMatrixCalcAdd)
        vModel(vModelIndex).vRotation.z = vModel(vModelIndex).vRotation.z + AngleZ
        Call D3DXMatrixRotationZ(vMatrixCalcAdd, vModel(vModelIndex).vRotation.z)
        Call D3DXMatrixMultiply(vMatrixBuffer, vMatrixBuffer, vMatrixCalcAdd)
        vModel(vModelIndex).matRotation = vMatrixBuffer
    End Sub

End Module
