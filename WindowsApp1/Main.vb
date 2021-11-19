Module Main
    ' all code are implented. YEAH!!!
    Declare Function GetTickCount Lib "kernel32.dll" () As Long

    Public ProgramState As String

    Sub Main()
        Call System_Start()

        Do Until ProgramState = "Exit"
            Call RenderMain()
        Loop
        Call System_Unload()
    End Sub

    Sub RenderMain()
        DoEvents
        If AForm1.WindowState = 0 Or AForm1.WindowState = 2 Then
            Select Case ProgramState
                Case "Ready"
                    Call FormResize()
                    Call RenderKeys()
                    Call RenderMouse()
                    Call RenderScreen()
            End Select
        End If
    End Sub

    Private Sub System_Start()
        ProgramState = "Ready"
        Call SetProgram()
        Call VB.Load(AForm1)
        Call aMouseKeyboard.LoadDirectXInput
        Call aMouseKeyboard.LoadKeyboardKeys
        Call LoadQuery()
        Call aScreen.SetTextColor
        Call LoadScreen
        AForm1.Visible = True
        Call aProgram.Program("System")
    End Sub

    Private Sub System_Unload()
        Dim iA As Integer
        ''On Local Error Resume Next
        D3DFont = Nothing
        For iA = 0 To ImageSpace
            D3DTexture(iA) = Nothing
        Next iA
        D3DDevice = Nothing
        D3D = Nothing
        Call DInputDeviceKeyBoard.Unacquire
        Call DInputDeviceMouse.Unacquire
        DInputDeviceMouse = Nothing
        DInputDeviceKeyBoard = Nothing
        DInput = Nothing
        Call VB.Unload(AForm1)
    End Sub

End Module
