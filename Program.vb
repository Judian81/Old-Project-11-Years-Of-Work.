Module aProgram

    'all code implanted
    Public Structure Emoticons
        Public Position As RECT
        Public Image As Integer
        Public Visible As Boolean
    End Structure

    '
    Public Structure JFont
        Public Bold As Boolean
        Public Charset As Integer
        Public Italic As Boolean
        Public Name As String
        Public Size As Double
        Public Strikethrough As Boolean
        Public Underline As Boolean
        Public Weight As Integer
    End Structure

    Public Structure JControl
        Public Accesskeys As String
        Public Alignment As Integer
        Public Appearance As Integer
        Public Archive As Boolean
        Public BackColor As Long
        Public BackStyle As Integer
        Public BorderStyle As Integer
        Public BorderWidth As Long
        Public BorderColor As Long
        Public Columns As Integer
        Public Container As Object
        Public DataChanged As Boolean
        Public DataField As String
        Public DataBase As Integer
        Public DataFormat As String
        Public DataMember As String
        Public DataSource As Integer
        Public Drive As String
        Public Enabled As Boolean
        Public FileName As String
        Public FillStyle As Integer
        Public Font As JFont
        Public ForeColor As Long
        Public HelpContextID As Long
        Public Image() As Integer
        Public Index As Integer
        Public IntegralHeight As Boolean
        Public Interval As Long
        Public LargeChange As Integer
        Public ListCount As Integer
        Public ListIndex As Integer
        Public Locked As Boolean
        Public Max As Integer
        Public Min As Integer
        Public MouseIcon As IPictureDisp
        Public MousePointer As Integer
        Public MultiLine As Boolean
        Public MultiSelect As Integer
        Public Name As String
        Public NewIndex As Integer
        Public Normal As Boolean
        Public Path As String
        Public Parent As String
        Public PasswordChar As String
        Public Position() As RECT
        Public PositionCount As Integer
        Public sReadOnly As Boolean
        Public RightToLeft As Boolean
        Public ScrollCount As Double
        Public ScrollIndex As Double
        Public SelCount As Integer
        Public SelLength As Integer
        Public SelStart As Integer
        Public SelText As Integer
        Public SetFocus As Boolean
        Public SmallChange As Integer
        Public Sorted As Boolean
        Public Style As Integer
        Public System As Boolean
        Public TabIndex As Integer
        Public TabStop As Boolean
        Public Tag As String
        Public Text() As String
        Public Value As Boolean
        Public ToolTipText As String
        Public TopIndex As Integer
        Public Visible As Boolean
        Public WhatsThisHelpID As Long
        Public WordWrap As Boolean
        Public X1 As Single       ' Left
        Public X2 As Single       ' Width, Right
        Public Y1 As Single       ' Top
        Public Y2 As Single       ' Height, Bottom
        Public Undo As String
        Public ZOrder As Integer
        Public Row As Byte
        Public Colom As Byte
        Public Emos() As Emoticons
        Public EmosGif() As Object
        Public EmosCount As Integer
        Public EmosGifCount As Integer
        Public EmosText As String
    End Structure

    Public Structure JForm
        Public Plugin As Object
        Public PluginForm As String
        Public ControlIndex As Integer
        Public ControlStructure As Integer
        Public Controls() As JControl
        Public ControlsCount As Integer
        Public Appearance As Integer
        Public BackColor As Long
        Public BorderStyle As Integer
        Public Caption As String
        Public ControlBox As Boolean
        Public Count As Integer
        Public CurrentX As Single
        Public CurrentY As Single
        Public DragMode As Integer
        Public DrawMode As Integer
        Public DrawStyle As Integer
        Public DrawWidth As Integer
        Public Enabled As Boolean
        Public FillColor As Long
        Public FillStyle As Integer
        Public Font As JFont
        Public ForeColor As Long
        Public Height As Single
        Public HelpContextID As Long
        Public Icon As IPictureDisp
        Public Image As IPictureDisp
        Public KeyPreview As Boolean
        Public Left As Single
        Public MouseIcon As IPictureDisp
        Public MousePointer As Integer
        Public Name As String
        Public Picture As IPictureDisp
        Public RightToLeft As Boolean
        Public SetFocus As Boolean
        Public Tag As String
        Public TabIndexCount As Integer
        Public Top As Single
        Public WindowState As Integer 'Maximized 2, Minimized 1, Normal 0
        Public Visible As Boolean
        Public WhatsThisHelpID As Long
        Public Width As Single
        Public ZOrder As Integer
    End Structure

    Public DXForm() As JForm
    Public xForm As New AForm()
    Public vForm As New AForm1()

    Public BFormCount As Byte
    Public BFormIndex As Byte
    Public BComboList As JControl
    Public BControl As JControl
    Public BControlIndex As Integer
    Public BControlStructure As String
    Public iFocus As Integer
    Public iFocus2 As Integer
    Public iFocus3 As Integer
    Public vMessageBox As String
    Public vMsgBoxEnabled As Boolean
    Public iFlashTimer As Double


    Sub CreateComboList()
        BComboList.Name = "ComboList"
        BComboList.FillStyle = 2
        BComboList.PositionCount = 17
        ReDim BComboList.Image(17)
        ReDim BComboList.Text(0)
        ReDim BComboList.Position(17)
        BComboList.Image(0) = 80
        BComboList.Image(1) = 81
        BComboList.Image(2) = 82
        BComboList.Image(3) = 83
        BComboList.Image(4) = 84
        BComboList.Image(5) = 87
        BComboList.Image(6) = 88
        BComboList.Image(7) = 89
        BComboList.Image(8) = 90
        BComboList.Image(9) = 15
        BComboList.Image(10) = 13
        BComboList.Image(11) = 14
        BComboList.Image(12) = 11
        BComboList.Image(13) = 18
        BComboList.Image(14) = 19
        BComboList.Image(15) = 20
        BComboList.Image(16) = 0
        BComboList.Image(17) = 84
    End Sub


    Sub Program(Name As String, Optional Name0 As String = "Load", Optional Index As Integer = 0, Optional KeyCode As Integer = 0, Optional Shift As Integer = 0, Optional KeyAscii As Integer = 0, Optional Button As Integer = 0, Optional X As Single = 0, Optional y As Single = 0)
        Dim vA As Integer
        Select Case Name0
            Case "Load"
                Call Load(Name)
                If Name = "Object Already Loaded" Then
                    Exit Sub
                End If
        End Select
        Select Case Name
            Case "System"
                Call aProgram.System(xForm(BFormIndex), Name0, Index, KeyCode, Shift, KeyAscii, Button, X, y)
            Case "MsgBox"
                Call aProgram.MsgBoxClick(xForm(BFormIndex), Name0, Index, KeyCode, Shift, KeyAscii, Button, X, y)
        End Select
        If Name0 = "Load" Then
            For vA = 0 To DXForm(BFormIndex).ControlsCount
                If DXForm(BFormIndex).Controls(vA).TabStop = True Then
                    If DXForm(BFormIndex).Controls(vA).TabIndex = 0 Then
                        Call FocusSet(BFormIndex, vA)
                        Exit For
                    End If
                End If
            Next vA
        End If
    End Sub

    Sub System(vForm As AForm, Optional Name As String = "Load", Optional Index As Integer = 0, Optional KeyCode As Integer = 0, Optional Shift As Integer = 0, Optional KeyAscii As Integer = 0, Optional Button As Integer = 0, Optional X As Single = 0, Optional y As Single = 0)
        Dim vA As Integer
        Select Case Name
            Case "Label1_Click()"
                Select Case MsgBox("Do you want to update the time?", vbYesNo)
                    Case vbYes
                        xForm(BFormIndex).Label1(1).Caption = "Time: " & VBA.Time$
                    Case vbNo
                End Select
            Case "Check1_Click()"
                Select Case xForm(BFormIndex).Check1(Index).Value
                    Case 1
                        xForm(BFormIndex).List1(0).Visible = True
                    Case 0
                        xForm(BFormIndex).List1(0).Visible = False
                End Select
            Case "Command1_Click()"
                ProgramState = "Exit"
            Case "Load"
                For vA = 0 To 3
                    Call Load(vForm.Label1(vA))
                Next vA
                xForm(BFormIndex).Label1(0).Appearance = 1
                xForm(BFormIndex).Label1(0).Caption = "&Date: " & VBA.Date$
                Call xForm(BFormIndex).Label1(0).Move(4, 54, 164, 22)
                xForm(BFormIndex).Label1(1).Caption = "Time: " & VBA.Time$
                xForm(BFormIndex).Label1(1).Tag = "VBA.Time$"
                Call xForm(BFormIndex).Label1(1).Move(167, 54, 159, 22)
                xForm(BFormIndex).Label1(2).Caption = Empty
                Call xForm(BFormIndex).Label1(2).Move(4, 76, 164, 22)
                xForm(BFormIndex).Label1(3).Caption = Empty
                Call xForm(BFormIndex).Label1(3).Move(167, 76, 159, 22)
                For vA = 0 To 0
                    Call Load(vForm.List1(vA))
                Next vA
                Call xForm(BFormIndex).List1(0).Move(424, 8, 100, 87)
                xForm(BFormIndex).List1(0).DataBase = DataNR("|Time")
                xForm(BFormIndex).List1(0).Visible = True
                For vA = 0 To 0
                    Call Load(vForm.Text1(vA))
                Next vA
                Call xForm(BFormIndex).Text1(0).Move(564, 8, 449, 87)
                xForm(BFormIndex).Text1(0).MultiLine = True
                xForm(BFormIndex).Text1(0).Text = "1 this is a textbox" & vbCrLf & "2 this line is going much to far in lenght of one line so it has to be broken in one why so you can still see the rest of the text on the next line." & vbCrLf & "3 this line should be shown on screen as normal." & vbCrLf & "4 this should be shown when you scrolle the page." & vbCrLf & "1 this is a textbox" & vbCrLf & "2 this line is going much to far in lenght of one line so it has to be broken in one why so you can still see the rest of the text on the next line." & vbCrLf & "3 this line should be shown on screen as normal." & vbCrLf & "4 this should be shown when you scrolle the page."
                xForm(BFormIndex).Text1(0).Locked = False
                For vA = 0 To 2
                    Call Load(vForm.BorderV(vA))
                Next vA
                Call xForm(BFormIndex).BorderV(0).Move(326, 4, 4, 96)
                Call xForm(BFormIndex).BorderV(1).Move(0, 0, 4, 768)
                Call xForm(BFormIndex).BorderV(2).Move(1020, 0, 4, 768)
                For vA = 0 To 3
                    Call Load(vForm.BorderH(vA))
                Next vA
                Call xForm(BFormIndex).BorderH(0).Move(0, 0, 1024, 4)
                Call xForm(BFormIndex).BorderH(1).Move(0, 100, 1024, 4)
                Call xForm(BFormIndex).BorderH(2).Move(0, 432, 1024, 4)
                Call xForm(BFormIndex).BorderH(3).Move(0, 764, 1024, 4)
                For vA = 0 To 0
                    Call Load(vForm.Check1(vA))
                Next vA
                xForm(BFormIndex).Check1(0).Caption = "&Checkbox"
                xForm(BFormIndex).Check1(0).Value = 1
                Call xForm(BFormIndex).Check1(0).Move(340, 4, 77, 26)
                For vA = 0 To 0
                    Call Load(vForm.Command1(vA))
                Next vA
                xForm(BFormIndex).Command1(0).Caption = "&Quit"
                Call xForm(BFormIndex).Command1(0).Move(15, 8, 70, 39)
                For vA = 0 To 0
                    Call Load(vForm.Combo1(0))
                Next vA
                xForm(BFormIndex).Combo1(0).DataSource = DataNR("|Time")
                Call xForm(BFormIndex).Combo1(0).Move(99, 8, 200, 39)
                xForm(BFormIndex).Combo1(0).ListIndex = 0
                xForm(BFormIndex).Combo1(0).TabIndex = 0
        End Select
    End Sub



    '--------------------------------------------------------------CODE IN EEN KEER GEPLAKT-----------------------------------------------------


    Sub MsgBoxClick(vForm As AForm, Optional Name As String = "Load", Optional Index As Integer = 0, Optional KeyCode As Integer = 0, Optional Shift As Integer = 0, Optional KeyAscii As Integer = 0, Optional Button As Integer = 0, Optional X As Single = 0, Optional y As Single = 0)
        Select Case Name
            Case "Command1_Click()"
                Select Case xForm(BFormIndex).Command1(Index).Caption
                    Case "&Yes"
                        vMessageBox = "Yes"
                    Case "&No"
                        vMessageBox = "No"
                    Case "&Cancel"
                        vMessageBox = "Cancel"
                    Case "&Ok"
                        vMessageBox = "Ok"
                End Select
        End Select
    End Sub

    Function MsgBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "DirectX 8 Controls") As VbMsgBoxResult
        Dim vA As Integer
        Call Load("MsgBox")
        For vA = 0 To 1
            Call Load(xForm(BFormIndex).Label1(vA))
        Next vA
        xForm(BFormIndex).Label1(0).Caption = Title
        Call xForm(BFormIndex).Label1(0).Move(CurrentWidth / FixWidth / 2 - 165, CurrentHeight / FixHeight / 2 - 75, 330, 31)
        xForm(BFormIndex).Label1(1).Caption = Prompt
        Call xForm(BFormIndex).Label1(1).Move(CurrentWidth / FixWidth / 2 - 165, CurrentHeight / FixHeight / 2 - 40, 330, 62)
        Select Case Buttons
            Case vbOKOnly
                Stop
            Case vbOKCancel
                For vA = 0 To 1
                    Call Load(xForm(BFormIndex).Command1(vA))
                Next vA
                xForm(BFormIndex).Command1(0).Caption = "&Ok"
                Call xForm(BFormIndex).Command1(0).Move(CurrentWidth / FixWidth / 2 - 140, CurrentHeight / FixHeight / 2 + 35, 75, 39)
                xForm(BFormIndex).Command1(1).Caption = "&Cancel"
                Call xForm(BFormIndex).Command1(1).Move(CurrentWidth / FixWidth / 2 + 50, CurrentHeight / FixHeight / 2 + 35, 75, 39)
            Case vbYesNo
                For vA = 0 To 1
                    Call Load(xForm(BFormIndex).Command1(vA))
                Next vA
                xForm(BFormIndex).Command1(0).Caption = "&Yes"
                Call xForm(BFormIndex).Command1(0).Move(CurrentWidth / FixWidth / 2 - 140, CurrentHeight / FixHeight / 2 + 35, 75, 39)
                xForm(BFormIndex).Command1(1).Caption = "&No"
                Call xForm(BFormIndex).Command1(1).Move(CurrentWidth / FixWidth / 2 + 50, CurrentHeight / FixHeight / 2 + 35, 75, 39)
            Case vbYesNoCancel
                For vA = 0 To 2
                    Call Load(xForm(BFormIndex).Command1(vA))
                Next vA
                xForm(BFormIndex).Command1(0).Caption = "&Yes"
                Call xForm(BFormIndex).Command1(0).Move(CurrentWidth / FixWidth / 2 - 140, CurrentHeight / FixHeight / 2 + 35, 75, 39)
                xForm(BFormIndex).Command1(1).Caption = "&No"
                Call xForm(BFormIndex).Command1(1).Move(CurrentWidth / FixWidth / 2 - 45, CurrentHeight / FixHeight / 2 + 35, 75, 39)
                xForm(BFormIndex).Command1(2).Caption = "&Cancel"
                Call xForm(BFormIndex).Command1(2).Move(CurrentWidth / FixWidth / 2 + 50, CurrentHeight / FixHeight / 2 + 35, 75, 39)
        End Select
        xForm(BFormIndex).Height = 150
        xForm(BFormIndex).Width = 330
        vMsgBoxEnabled = True
        Do Until vMessageBox <> "" Or ProgramState = "Exit"
            Call RenderMain()
        Loop
        Select Case vMessageBox
            Case "Yes"
                MsgBox = vbYes
            Case "No"
                MsgBox = vbNo
            Case "Cancel"
                MsgBox = vbCancel
            Case "Ok"
                MsgBox = vbOK
        End Select
        vMessageBox = ""
        Call ProgramClose("MsgBox")
        vMsgBoxEnabled = False
    End Function

    Sub ProgramClose(sA As String)
        Dim iA As Integer, iC As Integer
        For iA = 3 To BFormCount
            If DXForm(iA - 1).Name = sA Then
                DXForm(iA - 1) = DXForm(iA)
            End If
        Next iA
        BFormCount = BFormCount - 1
        BFormIndex = BFormCount
        ReDim Preserve DXForm(BFormCount)
        Call ScreenPosition()

        For iA = 3 To BFormCount
            If DXForm(iA).Height > 328 Then
                Exit For
            End If
        Next iA
        If iA > BFormCount Then
            Call ScreenDeskTop(False)
        Else
            Call ScreenDeskTop(True)
        End If
        iC = DataNR("|Forms Active")
        Call QueryClear(iC)
        For iA = 3 To BFormCount
            Call QueryAddItem(iC, DXForm(iA).Name)
        Next iA
    End Sub

    Sub ProgramStart(sA As String)
        Dim iA As Integer, iC As Integer
        If sA = Empty Then
            Exit Sub
        End If
        For iA = 1 To BFormCount
            If sA = DXForm(iA).Name Then
                If DXForm(iA).Height > 328 Then
                    Call ScreenDeskTop(True, sA)
                Else
                    Call ScreenDeskTop(False)
                End If
                Exit Sub
            End If
        Next iA
        Call Program(sA)
        Select Case sA
            Case "System"
            Case Else
                If DXForm(BFormIndex).Height > 328 Then
                    Call ScreenDeskTop(True, sA)
                Else
                    Call ScreenDeskTop(False)
                End If
        End Select
        Call ScreenPosition()
    End Sub

    Function Unload(Name As String) As String
        Dim vA As Integer
        If Name = "" Then
        ElseIf TypeOf Name Is AForm Then
            vForm(BFormCount) = Name
        Else
            For vA = 1 To BFormCount
                If DXForm(vA).Name = Name Then
                    BFormIndex = vA
                    Exit For
                End If
            Next vA
            If BFormIndex <> BFormCount Then
                For vA = BFormIndex To BFormCount - 1
                    vForm(vA) = vForm(vA + 1)
                    DXForm(vA) = DXForm(vA + 1)
                    xForm(vA) = xForm(vA + 1)
                Next vA
            End If
            vForm(BFormCount) = Nothing
            BFormCount = BFormCount - 1
            BFormIndex = BFormCount
            ReDim Preserve DXForm(BFormCount)
            ReDim Preserve vForm(BFormCount)
            ReDim Preserve xForm(BFormCount)
        End If
    End Function

    Sub ScreenDeskTop(bA As Boolean, Optional sA As String = "")
        Dim iA As Integer, bB As Boolean, iB As Byte, iC As Integer
        Static bC As Boolean
        iC = DataNR("|Forms Active")
        Call QueryClear(iC)
        For iA = 3 To BFormCount
            Call QueryAddItem(iC, DXForm(iA).Name)
        Next iA
        If bC = bA Then
            Exit Sub
        Else
            bC = bA
        End If
        bB = CBool(Xorr(CByte(bA)))
        iA = BFormIndex
        BFormIndex = 2
        xForm(iA).BorderH(2).Visible = bB
        BFormIndex = iA
        For iA = 3 To BFormCount
            If DXForm(iA).Height = 0 Then
                DXForm(iA).Visible = bB
            Else
                DXForm(iA).Visible = bA
            End If
        Next iA
    End Sub

    Sub ScreenPosition()
        Dim iA As Integer, iB As Integer, iC As Byte, iL As Long, iT As Long, iW As Integer, iH As Integer
        For iA = BFormCount To 2 Step -1
            If DXForm(iA).Visible = True Then
                If DXForm(iA).Name <> "System" And DXForm(iA).Name <> Empty Then
                    If iL = 0 Then
                        If DXForm(iA).Height = (514 * FixHeight) Then
                            If DXForm(iA).Top = (100 * FixHeight) Then
                                Exit Sub
                            End If
                            DXForm(iA).Top = (100 * FixHeight)
                            For iB = 0 To DXForm(iA).ControlsCount
                                DXForm(iA).Controls(iB).Y1 = DXForm(iA).Controls(iB).Y1 + DXForm(iA).Top
                                For iC = 0 To DXForm(iA).Controls(iB).PositionCount
                                    DXForm(iA).Controls(iB).Position(iC).Top = DXForm(iA).Controls(iB).Position(iC).Top + DXForm(iA).Top
                                    DXForm(iA).Controls(iB).Position(iC).Bottom = DXForm(iA).Controls(iB).Position(iC).Bottom + DXForm(iA).Top
                                Next iC
                            Next iB
                            Exit Sub
                        Else
                            iL = (4 * FixWidth)
                            iT = (100 * FixHeight)
                            iW = iL - DXForm(iA).Left
                            iH = iT - DXForm(iA).Top
                            For iB = 0 To DXForm(iA).ControlsCount
                                If iW >= 0 Then
                                    DXForm(iA).Controls(iB).Y1 = DXForm(iA).Controls(iB).Y1 + iH
                                    DXForm(iA).Controls(iB).X1 = DXForm(iA).Controls(iB).X1 + iW
                                End If
                                For iC = 0 To DXForm(iA).Controls(iB).PositionCount
                                    DXForm(iA).Controls(iB).Position(iC).Top = DXForm(iA).Controls(iB).Position(iC).Top + iH
                                    DXForm(iA).Controls(iB).Position(iC).Bottom = DXForm(iA).Controls(iB).Position(iC).Bottom + iH
                                    DXForm(iA).Controls(iB).Position(iC).Left = DXForm(iA).Controls(iB).Position(iC).Left + iW
                                    DXForm(iA).Controls(iB).Position(iC).Right = DXForm(iA).Controls(iB).Position(iC).Right + iW
                                Next iC
                            Next iB
                            DXForm(iA).Left = iL
                            DXForm(iA).Top = iT
                            iL = DXForm(iA).Left + DXForm(iA).Width
                        End If
                    ElseIf iL <> 0 Then
                        iW = iL - DXForm(iA).Left
                        iH = iT - DXForm(iA).Top
                        If iL + DXForm(iA).Width > 1024 * FixWidth Then
                            iL = (4 * FixWidth)
                            iT = iT + (332 * FixHeight)
                            iW = iL - DXForm(iA).Left
                            iH = iT - DXForm(iA).Top
                        End If
                        DXForm(iA).Left = iL
                        DXForm(iA).Top = iT
                        For iB = 0 To DXForm(iA).ControlsCount
                            For iC = 0 To DXForm(iA).Controls(iB).PositionCount
                                DXForm(iA).Controls(iB).Position(iC).Top = DXForm(iA).Controls(iB).Position(iC).Top + iH
                                DXForm(iA).Controls(iB).Position(iC).Bottom = DXForm(iA).Controls(iB).Position(iC).Bottom + iH
                                DXForm(iA).Controls(iB).Position(iC).Left = DXForm(iA).Controls(iB).Position(iC).Left + iW
                                DXForm(iA).Controls(iB).Position(iC).Right = DXForm(iA).Controls(iB).Position(iC).Right + iW
                            Next iC
                        Next iB
                        iL = DXForm(iA).Left + DXForm(iA).Width
                    End If
                End If
            End If
        Next iA
    End Sub

    Sub Message(sA As String, Optional yA As Byte = 144)
        Dim iA As Integer
        Static sMessage As String
        iA = BFormIndex
        BFormIndex = 2
        If BFormIndex > BFormCount Then
            sMessage = sMessage & sA & vbCrLf
            BFormIndex = iA
            Exit Sub
        ElseIf sMessage <> "" Then
            xForm(BFormIndex).Text1(0).Text = xForm(BFormIndex).Text1(0).Text & TextVerdelen(sMessage, xForm(BFormIndex).Text1(0).Font, xForm(BFormIndex).Text1(0).Width, xForm(BFormIndex).Text1(0).MultiLine) & vbCrLf
            sMessage = ""
        End If
        If sA = Empty Then
            xForm(BFormIndex).Text1(0).Text = Empty
        Else
            If sA & vbCrLf = Right(xForm(BFormIndex).Text1(0).Text, Len(sA) + 2) Then
                xForm(BFormIndex).Text1(0).Text = Left(xForm(BFormIndex).Text1(0).Text, Len(xForm(BFormIndex).Text1(0).Text) - Len(sA) - 2) & TextVerdelen(sA & " " & Right("000" & CStr(Val(Mid(xForm(BFormIndex).Text1(0).Text, Len(sA) + 2, 3)) + 1), 3), xForm(BFormIndex).Text1(0).Font, xForm(BFormIndex).Text1(0).Width, xForm(BFormIndex).Text1(0).MultiLine) & vbCrLf
            ElseIf sA = Left(Right(xForm(BFormIndex).Text1(0).Text, Len(sA) + 6), Len(sA)) Then
                xForm(BFormIndex).Text1(0).Text = Left(xForm(BFormIndex).Text1(0).Text, Len(xForm(BFormIndex).Text1(0).Text) - Len(sA) - 6) & TextVerdelen(sA & " " & Right("000" & CStr(Val(Left(Right(xForm(BFormIndex).Text1(0).Text, 5), 3)) + 1), 3), xForm(BFormIndex).Text1(0).Font, xForm(BFormIndex).Text1(0).Width, xForm(BFormIndex).Text1(0).MultiLine) & vbCrLf
            Else
                xForm(BFormIndex).Text1(0).Text = xForm(BFormIndex).Text1(0).Text & TextVerdelen(sA, xForm(BFormIndex).Text1(0).Font, xForm(BFormIndex).Text1(0).Width, xForm(BFormIndex).Text1(0).MultiLine) & vbCrLf
            End If
        End If
        BFormIndex = iA
    End Sub

    Sub SetProgram()
        Call CreateComboList()
        BFormCount = 1
        BFormIndex = BFormCount
        ReDim Preserve xForm(BFormCount)
        ReDim Preserve DXForm(BFormCount)
        DXForm(BFormCount).Visible = True
        ReDim DXForm(BFormCount).Controls(0)
        ReDim DXForm(BFormCount).Controls(0).Image(0)
        ReDim DXForm(BFormCount).Controls(0).Position(0)
        DXForm(BFormIndex).ControlsCount = 0
        DXForm(BFormIndex).Controls(0).PositionCount = 0
        DXForm(BFormIndex).Controls(0).Position(0).Right = CurrentWidth
        DXForm(BFormIndex).Controls(0).Position(0).Bottom = CurrentHeight
        DXForm(BFormIndex).Controls(0).Position(0).Left = 0
        DXForm(BFormIndex).Controls(0).Position(0).Top = 0
        DXForm(BFormIndex).Controls(0).Image(0) = 2
        DXForm(BFormIndex).Controls(0).Enabled = True
        DXForm(BFormIndex).Controls(0).Visible = True
        DXForm(BFormIndex).Controls(0).Visible = False
        DXForm(BFormIndex).Controls(0).Position(0).Right = 0
        DXForm(BFormIndex).Controls(0).Position(0).Bottom = 0
        DXForm(BFormIndex).Controls(0).Position(0).Left = 0
        DXForm(BFormIndex).Controls(0).Position(0).Top = 0
        TextSelectedCount = 0
        ReDim TextSelected(TextSelectedCount)
        ReDim DXForm(1)
        DXForm(1).ControlsCount = 0
    End Sub

    Sub Load(Name As String)
        Dim vA As Integer
        If TypeOf Name Is cControl Then
            Select Case Name.Name
                Case "Check1"
                    xForm(BFormIndex).Check1(Name.Index).Visible = True
                    xForm(BFormIndex).Check1(Name.Index).Enabled = True
                    xForm(BFormIndex).Check1(Name.Index).TabIndex = -1
                Case "List1"
                    xForm(BFormIndex).List1(Name.Index).Visible = True
                    xForm(BFormIndex).List1(Name.Index).Enabled = True
                    xForm(BFormIndex).List1(Name.Index).TabStop = True
                    xForm(BFormIndex).List1(Name.Index).TabIndex = DXForm(BFormIndex).TabIndexCount
                    DXForm(BFormIndex).TabIndexCount = DXForm(BFormIndex).TabIndexCount + 1
                Case "BorderH"
                    xForm(BFormIndex).BorderH(Name.Index).Visible = True
                    xForm(BFormIndex).BorderH(Name.Index).Enabled = True
                    xForm(BFormIndex).BorderH(Name.Index).TabIndex = -1
                Case "BorderV"
                    xForm(BFormIndex).BorderV(Name.Index).Visible = True
                    xForm(BFormIndex).BorderV(Name.Index).Enabled = True
                    xForm(BFormIndex).BorderV(Name.Index).TabIndex = -1
                Case "Image1"
                    xForm(BFormIndex).Image1(Name.Index).Visible = True
                    xForm(BFormIndex).Image1(Name.Index).Enabled = True
                    xForm(BFormIndex).Image1(Name.Index).TabIndex = -1
                Case "Command1"
                    xForm(BFormIndex).Command1(Name.Index).BackColor = -2147483633
                    xForm(BFormIndex).Command1(Name.Index).FontBold = False
                    xForm(BFormIndex).Command1(Name.Index).FontItalic = False
                    xForm(BFormIndex).Command1(Name.Index).FontName = "MS Sans Serif"
                    xForm(BFormIndex).Command1(Name.Index).FontSize = 8
                    xForm(BFormIndex).Command1(Name.Index).TabStop = False
                    xForm(BFormIndex).Command1(Name.Index).TabIndex = -1
                    xForm(BFormIndex).Command1(Name.Index).Visible = True
                    xForm(BFormIndex).Command1(Name.Index).Enabled = True
                Case "Text1"
                    xForm(BFormIndex).Text1(Name.Index).Visible = True
                    xForm(BFormIndex).Text1(Name.Index).Alignment = 0
                    xForm(BFormIndex).Text1(Name.Index).Appearance = 1
                    xForm(BFormIndex).Text1(Name.Index).BackColor = &HC0FFC0
                    xForm(BFormIndex).Text1(Name.Index).ForeColor = vbBlack
                    xForm(BFormIndex).Text1(Name.Index).BorderStyle = 1
                    xForm(BFormIndex).Text1(Name.Index).FontBold = True
                    xForm(BFormIndex).Text1(Name.Index).FontItalic = False
                    xForm(BFormIndex).Text1(Name.Index).FontName = "Arial"
                    xForm(BFormIndex).Text1(Name.Index).FontSize = 12
                    xForm(BFormIndex).Text1(Name.Index).TabStop = True
                    xForm(BFormIndex).Text1(Name.Index).TabIndex = DXForm(BFormIndex).TabIndexCount
                    DXForm(BFormIndex).TabIndexCount = DXForm(BFormIndex).TabIndexCount + 1
                    Call xForm(BFormIndex).Text1(Name.Index).ZOrder(0)
                    xForm(BFormIndex).Text1(Name.Index).Enabled = True
                    'XForm(BFormIndex).Text1(Name.Index).ListCount = 0
                    xForm(BFormIndex).Text1(Name.Index).ListIndex = 0
                Case "Combo1"
                    xForm(BFormIndex).Combo1(Name.Index).Visible = True
                    xForm(BFormIndex).Combo1(Name.Index).Appearance = 1
                    xForm(BFormIndex).Combo1(Name.Index).BackColor = &HC0FFC0
                    xForm(BFormIndex).Combo1(Name.Index).BorderStyle = 1
                    xForm(BFormIndex).Combo1(Name.Index).TabStop = True
                    xForm(BFormIndex).Combo1(Name.Index).TabIndex = DXForm(BFormIndex).TabIndexCount
                    DXForm(BFormIndex).TabIndexCount = DXForm(BFormIndex).TabIndexCount + 1
                    xForm(BFormIndex).Combo1(Name.Index).FontName = "Arial"
                    xForm(BFormIndex).Combo1(Name.Index).FontBold = True
                    xForm(BFormIndex).Combo1(Name.Index).FontItalic = False
                    xForm(BFormIndex).Combo1(Name.Index).FontSize = 12
                    xForm(BFormIndex).Combo1(Name.Index).Enabled = True
                Case "Data1"
                    xForm(BFormIndex).Data1(Name.Index).Visible = True
                    xForm(BFormIndex).Data1(Name.Index).TabStop = True
                    xForm(BFormIndex).Data1(Name.Index).TabIndex = DXForm(BFormIndex).TabIndexCount
                    DXForm(BFormIndex).TabIndexCount = DXForm(BFormIndex).TabIndexCount + 1
                    xForm(BFormIndex).Data1(Name.Index).DataChanged = False
                    xForm(BFormIndex).Data1(Name.Index).Enabled = True
                Case "Label1"
                    xForm(BFormIndex).Label1(Name.Index).Visible = True
                    xForm(BFormIndex).Label1(Name.Index).Alignment = 0
                    xForm(BFormIndex).Label1(Name.Index).Appearance = 0
                    xForm(BFormIndex).Label1(Name.Index).BackStyle = 1
                    xForm(BFormIndex).Label1(Name.Index).BorderStyle = 1
                    xForm(BFormIndex).Label1(Name.Index).BackColor = 16777215
                    xForm(BFormIndex).Label1(Name.Index).ForeColor = 0
                    xForm(BFormIndex).Label1(Name.Index).FontBold = True
                    xForm(BFormIndex).Label1(Name.Index).FontItalic = False
                    xForm(BFormIndex).Label1(Name.Index).FontName = "Arial"
                    xForm(BFormIndex).Label1(Name.Index).FontSize = 8
                    xForm(BFormIndex).Label1(Name.Index).Enabled = True
                    xForm(BFormIndex).Label1(Name.Index).TabIndex = -1
                Case "Option1"
                    xForm(BFormIndex).Option1(Name.Index).Visible = True
                    xForm(BFormIndex).Option1(Name.Index).Enabled = True
                    xForm(BFormIndex).Option1(Name.Index).TabIndex = -1
            End Select
        ElseIf TypeOf Name Is AForm Then
            vForm(BFormCount) = Name
        Else
            For vA = 1 To BFormCount
                If DXForm(vA).Name = Name Then
                    BFormIndex = vA
                    Exit For
                End If
            Next vA
            If vA > BFormCount Then
                BFormCount = BFormCount + 1
                BFormIndex = BFormCount
                ReDim Preserve xForm(BFormCount)
                ReDim Preserve DXForm(BFormCount)
                DXForm(BFormCount).Name = Name
                DXForm(BFormCount).Visible = True
                DXForm(BFormCount).Top = 0
                DXForm(BFormCount).ControlsCount = -1
                ReDim DXForm(BFormCount).Controls(0)
                ReDim DXForm(BFormCount).Controls(0).Image(0)
                ReDim DXForm(BFormCount).Controls(0).Position(0)
                ReDim DXForm(BFormCount).Controls(0).Text(0)
                DXForm(BFormCount).Controls(0).Index = -1
            Else
                Name = "Object Already Loaded"
            End If
        End If
    End Sub

    Function ControlGet(sType As String) As JControl
        Dim vA As Integer
        If sType = "Controls" Then
            ControlGet = DXForm(BFormIndex).Controls(DXForm(BFormIndex).ControlIndex)
            Exit Function
        End If
        For vA = 0 To DXForm(BFormIndex).ControlsCount
            If sType = DXForm(BFormIndex).Controls(vA).Name Then
                If DXForm(BFormIndex).Controls(vA).Index = DXForm(BFormIndex).ControlIndex Then
                    ControlGet = DXForm(BFormIndex).Controls(vA)
                    Exit Function
                End If
            End If
        Next vA
        DXForm(BFormIndex).ControlsCount = DXForm(BFormIndex).ControlsCount + 1
        vA = DXForm(BFormIndex).ControlsCount
        If vA = 1 Then
            ReDim Preserve DXForm(BFormIndex).Controls(150)
        End If
        DXForm(BFormIndex).Controls(vA).Name = sType
        DXForm(BFormIndex).Controls(vA).Index = DXForm(BFormIndex).ControlIndex
        Select Case sType
            Case "BorderH"
                ReDim DXForm(BFormIndex).Controls(vA).Image(0)
                ReDim DXForm(BFormIndex).Controls(vA).Text(0)
                ReDim DXForm(BFormIndex).Controls(vA).Position(0)
                DXForm(BFormIndex).Controls(vA).Image(0) = 8
            Case "BorderV"
                ReDim DXForm(BFormIndex).Controls(vA).Image(0)
                ReDim DXForm(BFormIndex).Controls(vA).Text(0)
                ReDim DXForm(BFormIndex).Controls(vA).Position(0)
                DXForm(BFormIndex).Controls(vA).Image(0) = 7
            Case "Combo1"
                DXForm(BFormIndex).Controls(vA).FillStyle = 2
                DXForm(BFormIndex).Controls(vA).PositionCount = 9
                ReDim DXForm(BFormIndex).Controls(vA).Image(9)
                ReDim DXForm(BFormIndex).Controls(vA).Text(0)
                ReDim DXForm(BFormIndex).Controls(vA).Position(9)
                DXForm(BFormIndex).Controls(vA).Image(0) = 111
                DXForm(BFormIndex).Controls(vA).Image(1) = 112
                DXForm(BFormIndex).Controls(vA).Image(2) = 113
                DXForm(BFormIndex).Controls(vA).Image(3) = 114
                DXForm(BFormIndex).Controls(vA).Image(4) = 116
                DXForm(BFormIndex).Controls(vA).Image(5) = 118
                DXForm(BFormIndex).Controls(vA).Image(6) = 119
                DXForm(BFormIndex).Controls(vA).Image(7) = 120
                DXForm(BFormIndex).Controls(vA).Image(8) = 121
                DXForm(BFormIndex).Controls(vA).Image(9) = 11
            Case "Command1"
                ReDim DXForm(BFormIndex).Controls(vA).Text(0)
                ReDim DXForm(BFormIndex).Controls(vA).Image(0)
                ReDim DXForm(BFormIndex).Controls(vA).Position(0)
                DXForm(BFormIndex).Controls(vA).Image(0) = 9
            Case "Data1"
                DXForm(BFormIndex).Controls(vA).PositionCount = 2
                ReDim DXForm(BFormIndex).Controls(vA).Text(0)
                ReDim DXForm(BFormIndex).Controls(vA).Image(2)
                ReDim DXForm(BFormIndex).Controls(vA).Position(2)
                DXForm(BFormIndex).Controls(vA).Image(1) = 93
                DXForm(BFormIndex).Controls(vA).Image(0) = 74
            Case "Dir1"
            Case "Drive1"
            Case "File1"
            Case "Frame1"
                ReDim DXForm(BFormIndex).Controls(vA).Image(0)
                ReDim DXForm(BFormIndex).Controls(vA).Text(0)
                ReDim DXForm(BFormIndex).Controls(vA).Position(0)
                DXForm(BFormIndex).Controls(vA).Image(0) = 0
            Case "Image1"
                DXForm(BFormIndex).Controls(vA).PositionCount = 0
                ReDim DXForm(BFormIndex).Controls(vA).Image(0)
                ReDim DXForm(BFormIndex).Controls(vA).Text(0)
                ReDim DXForm(BFormIndex).Controls(vA).Position(0)
                DXForm(BFormIndex).Controls(vA).Image(0) = 0
            Case "Label1"
                DXForm(BFormIndex).Controls(vA).PositionCount = 2
                ReDim DXForm(BFormIndex).Controls(vA).Image(2)
                ReDim DXForm(BFormIndex).Controls(vA).Text(0)
                ReDim DXForm(BFormIndex).Controls(vA).Position(2)
                DXForm(BFormIndex).Controls(vA).Image(0) = 60
                DXForm(BFormIndex).Controls(vA).Image(1) = 64
                DXForm(BFormIndex).Controls(vA).Image(2) = 61
            Case "Line1"
            Case "List1"
                DXForm(BFormIndex).Controls(vA).FillStyle = 2
                DXForm(BFormIndex).Controls(vA).PositionCount = 17
                ReDim DXForm(BFormIndex).Controls(vA).Image(17)
                ReDim DXForm(BFormIndex).Controls(vA).Text(0)
                ReDim DXForm(BFormIndex).Controls(vA).Position(17)
                DXForm(BFormIndex).Controls(vA).Image(0) = 80
                DXForm(BFormIndex).Controls(vA).Image(1) = 81
                DXForm(BFormIndex).Controls(vA).Image(2) = 82
                DXForm(BFormIndex).Controls(vA).Image(3) = 83
                DXForm(BFormIndex).Controls(vA).Image(4) = 84
                DXForm(BFormIndex).Controls(vA).Image(5) = 87
                DXForm(BFormIndex).Controls(vA).Image(6) = 88
                DXForm(BFormIndex).Controls(vA).Image(7) = 89
                DXForm(BFormIndex).Controls(vA).Image(8) = 90
                DXForm(BFormIndex).Controls(vA).Image(9) = 15
                DXForm(BFormIndex).Controls(vA).Image(10) = 13
                DXForm(BFormIndex).Controls(vA).Image(11) = 14
                DXForm(BFormIndex).Controls(vA).Image(12) = 11
                DXForm(BFormIndex).Controls(vA).Image(13) = 18
                DXForm(BFormIndex).Controls(vA).Image(14) = 19
                DXForm(BFormIndex).Controls(vA).Image(15) = 20
                DXForm(BFormIndex).Controls(vA).Image(16) = 0
                DXForm(BFormIndex).Controls(vA).Image(17) = 84
            Case "Option1", "Check1", "ButtonPress"
                ReDim DXForm(BFormIndex).Controls(vA).Image(0)
                ReDim DXForm(BFormIndex).Controls(vA).Text(0)
                ReDim DXForm(BFormIndex).Controls(vA).Position(0)
                DXForm(BFormIndex).Controls(vA).Image(0) = 9
            Case "Picture1"
            Case "HScroll1"
            Case "VScroll1"
                DXForm(BFormIndex).Controls(vA).PositionCount = 3
                ReDim DXForm(BFormIndex).Controls(vA).Image(3)
                ReDim DXForm(BFormIndex).Controls(vA).Text(0)
                ReDim DXForm(BFormIndex).Controls(vA).Position(3)
                DXForm(BFormIndex).Controls(vA).Image(0) = 15
                DXForm(BFormIndex).Controls(vA).Image(1) = 0
                DXForm(BFormIndex).Controls(vA).Image(2) = 13
                DXForm(BFormIndex).Controls(vA).Image(3) = 11
            Case "Shape1"
            Case "Text1"
                DXForm(BFormIndex).Controls(vA).PositionCount = 16
                ReDim DXForm(BFormIndex).Controls(vA).Image(16)
                ReDim DXForm(BFormIndex).Controls(vA).Text(0)
                ReDim DXForm(BFormIndex).Controls(vA).Position(16)
                DXForm(BFormIndex).Controls(vA).Image(0) = 74
                DXForm(BFormIndex).Controls(vA).Image(1) = 71
                DXForm(BFormIndex).Controls(vA).Image(2) = 72
                DXForm(BFormIndex).Controls(vA).Image(3) = 73
                DXForm(BFormIndex).Controls(vA).Image(4) = 70
                DXForm(BFormIndex).Controls(vA).Image(5) = 75
                DXForm(BFormIndex).Controls(vA).Image(6) = 76
                DXForm(BFormIndex).Controls(vA).Image(7) = 77
                DXForm(BFormIndex).Controls(vA).Image(8) = 78
                DXForm(BFormIndex).Controls(vA).Image(9) = 0
                DXForm(BFormIndex).Controls(vA).Image(10) = 0
                DXForm(BFormIndex).Controls(vA).Image(11) = 0
                DXForm(BFormIndex).Controls(vA).Image(12) = 0
                DXForm(BFormIndex).Controls(vA).Image(13) = 0
                DXForm(BFormIndex).Controls(vA).Image(14) = 0
                DXForm(BFormIndex).Controls(vA).Image(15) = 0
                DXForm(BFormIndex).Controls(vA).Image(16) = 0
            Case "Timer1"
        End Select
        ControlGet = DXForm(BFormIndex).Controls(vA)
    End Function

    Sub ControlSet(Control As JControl)
        Dim vA As Integer
        If Control.Name = "ComboList" Then
            BComboList = Control
        Else
            For vA = 0 To DXForm(BFormIndex).ControlsCount
                If Control.Name = DXForm(BFormIndex).Controls(vA).Name Then
                    If Control.Index = DXForm(BFormIndex).Controls(vA).Index Then
                        DXForm(BFormIndex).Controls(vA) = Control
                        Exit Sub
                    End If
                End If
            Next vA
        End If
    End Sub

    Sub ControlUse(iE As Byte, vA As Integer, iK As Integer)
        Dim iA As Integer, iC As Byte, iD As Byte
        Dim vC As Integer, bA As Boolean
        iC = iE
        iD = BFormIndex
        BFormIndex = iC
        If (vA < 0 And vA > DXForm(iC).ControlsCount) Then
            Exit Sub
        End If
        If DXForm(iC).Controls(vA).Name = "Command1" Then
            Call xForm(iC).Controls(vA).Click
        ElseIf DXForm(iC).Controls(vA).Name = "List1" And iK = 28 Then
            Call xForm(iC).Controls(vA).DblClick
        ElseIf DXForm(iC).Controls(vA).Name = "List1" Then
            Call xForm(iC).Controls(vA).Click
        ElseIf DXForm(iC).Controls(vA).Name = "Label1" Then
            If DXForm(iC).Controls(vA).Text(0) = "&" & DXForm(iC).Name Then
                Call ProgramClose(DXForm(iC).Name)
                If Not DXForm(iC).Plugin Is Nothing Then
                    DXForm(iC).Plugin = Nothing
                End If
                If BFormIndex + 1 > BFormCount Then
                    BFormIndex = 2
                Else
                    BFormIndex = BFormIndex + 1
                End If
                iA = 0
                For vC = 0 To DXForm(BFormIndex).ControlsCount
                    If DXForm(BFormIndex).Controls(vC).TabIndex = iA Then
                        If DXForm(BFormIndex).Controls(vC).TabStop = False Then
                            iA = iA + 1
                            vC = -1
                        ElseIf DXForm(BFormIndex).Controls(vC).Visible = False Then
                            iA = iA + 1
                            vC = -1
                        Else
                            If DXForm(BFormIndex).Controls(vC).Colom > 1 Then
                                Call FocusSet(BFormIndex, vC, CByte(DXForm(BFormIndex).Controls(vC).Colom))
                            Else
                                Call FocusSet(BFormIndex, vC)
                            End If
                            Exit Sub
                        End If
                    End If
                    If iA = DXForm(BFormIndex).TabIndexCount Then
                        iA = 0
                        If bA = False Then
                            bA = True
                        Else
                            Exit Sub
                        End If
                    ElseIf vC = DXForm(BFormIndex).ControlsCount Then
                        If bA = False Then
                            bA = True
                        Else
                            Exit Sub
                        End If
                        vC = -1
                    End If
                Next vC
                Exit Sub
            Else
                Call xForm(iC).Controls(vA).Click
            End If
        ElseIf DXForm(iC).Controls(vA).Name = "Image1" Then
            Call xForm(iC).Controls(vA).Click
        ElseIf DXForm(iC).Controls(vA).Name = "Combo1" And iK = 28 Then
            Call xForm(iC).Controls(vA).DblClick
        ElseIf DXForm(iC).Controls(vA).Name = "Combo1" Then
            Call xForm(iC).Controls(vA).Click
        ElseIf DXForm(iC).Controls(vA).Name = "Check1" Then
            Call xForm(iC).Controls(vA).Click
        ElseIf DXForm(iC).Controls(vA).Name = "Option1" Then
            Call xForm(iC).Controls(vA).Click
        ElseIf DXForm(iC).Controls(vA).Name = "Data1" Then
            Call xForm(iC).Controls(vA).Change
        End If
        If BFormIndex = iC Then
            BFormIndex = iD
        End If
    End Sub

End Module
