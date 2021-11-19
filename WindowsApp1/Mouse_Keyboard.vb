Module aMouseKeyboard

    'all code implented
    Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long

    Structure POINTAPI
        Public l As Long
        Public t As Long
    End Structure
    Public MousePos As POINTAPI
    Public MouseButton(3) As Boolean
    Public KeyNames(355) As String
    Public KeyPressed(255) As Boolean
    'Public DInput As DirectInput8
    'Public DInputDeviceMouse As DirectInputDevice8
    'Public DInputDeviceKeyBoard As DirectInputDevice8
    'Public DStateKeyBoard As DIKEYBOARDSTATE
    'Public DStateMouse As DIMOUSESTATE
    Public MouseDblClickTime As Double
    Public MouseDblClickFocus As Integer
    Public MouseDblClickFocus2 As Integer
    Public iCheckOption As Byte

    Function KeyAltAction(iK As Integer) As Boolean
        Dim iA As Integer, iB As Integer, iC As Integer, iD As Integer, vKeyAltAction As Boolean
        vKeyAltAction = False
        If ((KeyPressed(56) Or KeyPressed(184)) = True) Then
            'this is set to ture
            vKeyAltAction = True
            If (iK > 1 And iK < 12) Or (iK > 15 And iK < 26) Or (iK > 29 And iK < 39) Or (iK > 43 And iK < 51) Then
                For iA = 0 To DXForm(BFormIndex).ControlsCount
                    If DXForm(BFormIndex).Controls(iA).Visible = True Then
                        If DXForm(BFormIndex).Controls(iA).Image(0) = 9 Or DXForm(BFormIndex).Controls(iA).Image(0) = 10 Or DXForm(BFormIndex).Controls(iA).Image(0) = 60 Or DXForm(BFormIndex).Controls(iA).Image(0) = 61 Or DXForm(BFormIndex).Controls(iA).Image(0) = 62 Then
                            For iB = 1 To Len(DXForm(BFormIndex).Controls(iA).Text(0))
                                If Mid(DXForm(BFormIndex).Controls(iA).Text(0), iB, 1) = "&" Then
                                    If UCase(Mid(DXForm(BFormIndex).Controls(iA).Text(0), iB + 1, 1)) = UCase(KeyNames(iK)) Then
                                        If DXForm(BFormIndex).Controls(iA).Name = "Check1" Then
                                            If DXForm(BFormIndex).Controls(iA).Value = False Then
                                                xForm(BFormIndex).Controls(iA).Value = 1
                                            Else
                                                xForm(BFormIndex).Controls(iA).Value = 0
                                            End If
                                        End If
                                        Call ControlUse(BFormIndex, iA, iK)
                                        Return vKeyAltAction
                                    Else
                                        Exit For
                                    End If
                                End If
                            Next iB
                        End If
                    End If
                Next iA
                For iA = 3 To BFormCount
                    If DXForm(iA).Name = "MsgBox" Then
                        'do i set a true or a false. i have a cluw but i need the be shure
                        Exit Function
                    End If
                Next iA
                For iC = 2 To UBound(DXForm)
                    If iC <> BFormIndex Then
                        For iA = 0 To DXForm(iC).ControlsCount
                            If DXForm(iC).Controls(iA).Visible = True Then
                                If DXForm(iC).Controls(iA).Image(0) = 9 Or DXForm(iC).Controls(iA).Image(0) = 10 Or DXForm(iC).Controls(iA).Image(0) = 60 Or DXForm(iC).Controls(iA).Image(0) = 61 Or DXForm(iC).Controls(iA).Image(0) = 62 Then
                                    For iB = 1 To Len(DXForm(iC).Controls(iA).Text(0))
                                        If Mid(DXForm(iC).Controls(iA).Text(0), iB, 1) = "&" Then
                                            If UCase(Mid(DXForm(iC).Controls(iA).Text(0), iB + 1, 1)) = UCase(KeyNames(iK)) Then
                                                If DXForm(iC).Controls(iA).Name = "Check1" Then
                                                    iD = BFormIndex
                                                    BFormIndex = iC
                                                    If DXForm(iC).Controls(iA).Value = False Then
                                                        xForm(iC).Controls(iA).Value = 1
                                                    Else
                                                        xForm(iC).Controls(iA).Value = 0
                                                    End If
                                                    BFormIndex = iD
                                                End If
                                                Call ControlUse(CByte(iC), iA, iK)
                                                Return vKeyAltAction
                                            Else
                                                Exit For
                                            End If
                                        End If
                                    Next iB
                                End If
                            End If
                        Next iA
                    End If
                Next iC
            End If
        Else
            vKeyAltAction = False
        End If
        Return vKeyAltAction
    End Function

    Sub KeyLink(iK As Integer)
        Dim vA As Integer
        For vA = 3 To BFormCount
            If vA <> BFormIndex Then
                If DXForm(vA).Name = "MsgBox" Then
                    Exit Sub
                End If
            End If
        Next vA
        Select Case True
            Case ((KeyPressed(29) = True Or iK = 29) And (KeyPressed(56) = True Or iK = 56) And (KeyPressed(211) = True Or iK = 211)), ((KeyPressed(29) = True Or iK = 29) And (KeyPressed(1) = True Or iK = 1)), ((KeyPressed(56) = True Or iK = 56) And (KeyPressed(15) = True Or iK = 15)), ((KeyPressed(56) = True Or iK = 56) And (KeyPressed(1) = True Or iK = 1)), (KeyPressed(219) Or iK = 219), (KeyPressed(221) Or iK = 221), (KeyPressed(220) Or iK = 220)
                'this has something to do about code i do not have yet ??????
                If ProgramState = "Ready" Then
                    AForm1.WindowState = 1
                End If
                Exit Sub
        End Select
        Select Case True
            Case KeyPressed(210), KeyPressed(58), KeyPressed(69), KeyPressed(70), KeyPressed(183), KeyPressed(59), KeyPressed(60), KeyPressed(61), KeyPressed(62), KeyPressed(63), KeyPressed(64), KeyPressed(65), KeyPressed(66), KeyPressed(67), KeyPressed(68), KeyPressed(87), KeyPressed(88)
                Exit Sub
        End Select
        Select Case iK
            Case 210, 58, 69, 70, 183, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 87, 88, 56, 42, 184, 157, 54, 29
                Exit Sub
        End Select
        If KeyTab(iK) = True Then
            BComboList.Visible = False
            Exit Sub
        ElseIf KeyAltAction(iK) = True Then
            Exit Sub
        End If
        If BComboList.Visible = True Then
            Call KeyListBoxCombo(iK)
            Exit Sub
        End If
        '??????????????????????????
        Select Case DXForm(BFormIndex).Controls(iFocus).Name
            Case "Text1"
                ''Call KeyTextBox(iK)
            Case "List1"
                Call KeyListBox(iK)
            Case "Combo1"
                ''Call KeyComboBox(iK)
            Case "Data1"
                If DXForm(BFormIndex).Controls(iFocus).Image(0) = 74 Or DXForm(BFormIndex).Controls(iFocus).Colom > 1 Or DXForm(BFormIndex).Controls(iFocus).Row > 1 Then
                    ''Call KeyTextBox(iK)
                ElseIf DXForm(BFormIndex).Controls(iFocus).Image(0) = 117 Then
                    ''Call KeyComboBox(iK)
                End If
        End Select
    End Sub

    Sub KeyListBox(iK As Integer)
        Dim iA As Double
        Select Case iK
            Case 208
                If Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount = DXForm(BFormIndex).Controls(iFocus).ListIndex Then
                    Exit Sub
                Else
                    DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex + 1
                End If
            Case 200
                If DXForm(BFormIndex).Controls(iFocus).ListIndex = 0 Then
                    Exit Sub
                Else
                    DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex - 1
                End If
            Case 28, 156
                Call ControlUse(BFormIndex, iFocus, iK)
                TextList = Empty
                Exit Sub
            Case 199
                DXForm(BFormIndex).Controls(iFocus).ListIndex = 0
            Case 201
                If DXForm(BFormIndex).Controls(iFocus).ListIndex = 0 Then
                    Exit Sub
                ElseIf 1 > DXForm(BFormIndex).Controls(iFocus).ListIndex - (DXForm(BFormIndex).Controls(iFocus).PositionCount - 17) - 1 Then
                    DXForm(BFormIndex).Controls(iFocus).ListIndex = 0
                Else
                    DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex - (DXForm(BFormIndex).Controls(iFocus).PositionCount - 17) - 1
                End If
            Case 207
                DXForm(BFormIndex).Controls(iFocus).ListIndex = Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount
            Case 209
                If Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount = DXForm(BFormIndex).Controls(iFocus).ListIndex Then
                    Exit Sub
                ElseIf Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount < DXForm(BFormIndex).Controls(iFocus).ListIndex + (DXForm(BFormIndex).Controls(iFocus).PositionCount - 17) + 1 Then
                    DXForm(BFormIndex).Controls(iFocus).ListIndex = Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount
                Else
                    DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex + (DXForm(BFormIndex).Controls(iFocus).PositionCount - 17) + 1
                End If
            Case 1
                For iA = 1 To Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount
                    Query(DXForm(BFormIndex).Controls(iFocus).DataBase).Selected(iA) = False
                Next iA
                Query(DXForm(BFormIndex).Controls(iFocus).DataBase).SelCount = 0
            Case Else
                If (KeyPressed(29) Or KeyPressed(157)) = True And DXForm(BFormIndex).Controls(iFocus).Image(0) <> 117 Then
                    Select Case iK
                        Case 211
                            'dit doet dus niks.
                            Call QueryRemoveSelected(DXForm(BFormIndex).Controls(iFocus).DataBase)
                            ''''Call ScreenListBox
                            'TextList = ""
                            Exit Sub
                        Case 44
                            Call QueryUndoRemove(DXForm(BFormIndex).Controls(iFocus).DataBase)
                            Call ScreenListBox()
                            TextList = ""
                            Exit Sub
                        Case 30
                            If Query(DXForm(BFormIndex).Controls(iFocus).DataBase).SelCount <> Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount Then
                                Query(DXForm(BFormIndex).Controls(iFocus).DataBase).SelCount = 0
                                For iA = 1 To Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount
                                    Query(DXForm(BFormIndex).Controls(iFocus).DataBase).Selected(iA) = True
                                    Query(DXForm(BFormIndex).Controls(iFocus).DataBase).SelCount = Query(DXForm(BFormIndex).Controls(iFocus).DataBase).SelCount + 1
                                Next iA
                            End If
                            'TextList = Empty
                            ''Call ScreenListBox
                            Exit Sub
                        Case 57
                            If DXForm(BFormIndex).Controls(iFocus).ListIndex > 0 Then
                                If Query(DXForm(BFormIndex).Controls(iFocus).DataBase).Selected(DXForm(BFormIndex).Controls(iFocus).ListIndex) = Empty Then
                                    Query(DXForm(BFormIndex).Controls(iFocus).DataBase).Selected(DXForm(BFormIndex).Controls(iFocus).ListIndex) = 1
                                    Query(DXForm(BFormIndex).Controls(iFocus).DataBase).SelCount = Query(DXForm(BFormIndex).Controls(iFocus).DataBase).SelCount + 1
                                Else
                                    Query(DXForm(BFormIndex).Controls(iFocus).DataBase).Selected(DXForm(BFormIndex).Controls(iFocus).ListIndex) = Empty
                                    Query(DXForm(BFormIndex).Controls(iFocus).DataBase).SelCount = Query(DXForm(BFormIndex).Controls(iFocus).DataBase).SelCount - 1
                                End If
                            End If
                            'TextList = ""
                            'Call ScreenListBox
                            Exit Sub
                    End Select
                End If
        End Select
        Select Case iK
            Case 199, 200, 201, 207, 208, 209
                If Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount > 0 Then
                    ''''DXForm(BFormIndex).Controls(iFocus).Position(11).Top = DXForm(BFormIndex).Controls(iFocus).Position(10).Top + ((DXForm(BFormIndex).Controls(iFocus).Position(10).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(10).Top - (10 * FixHeight)) / Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount) * DXForm(BFormIndex).Controls(iFocus).ListIndex
                    ''''DXForm(BFormIndex).Controls(iFocus).Position(11).Bottom = DXForm(BFormIndex).Controls(iFocus).Position(11).Top + (10 * FixHeight)
                End If
                'TextList = ""
            Case Else
                If (KeyPressed(42) Or KeyPressed(54)) = True Then
                    If iK + 300 > 355 Then
                        Exit Sub
                    End If
                    ''TextList = TextList & KeyNames(iK + 300)
                Else
                    ''TextList = TextList & KeyNames(iK)
                End If
                For iA = 0 To Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount
                    If UCase(Left(Query(DXForm(BFormIndex).Controls(iFocus).DataBase).List(iA), Len(TextList))) = UCase(TextList) And UCase(TextList) <> Empty Then
                        DXForm(BFormIndex).Controls(iFocus).ListIndex = iA
                        Exit For
                    End If
                Next iA
        End Select
        ''Call ScreenListBox
        ''Call aProgram.Program(DXForm(BFormIndex).Name, DXForm(BFormIndex).Controls(iFocus).Name & "_" & "Click()", DXForm(BFormIndex).Controls(iFocus).Index)
    End Sub

    Sub KeyListBoxCombo(iK As Integer)
        Dim iA As Double
        Select Case iK
            Case 208
                If Query(BComboList.DataBase).ListCount = BComboList.ListIndex Then
                    Exit Sub
                Else
                    BComboList.ListIndex += 1
                End If
            Case 200
                If BComboList.ListIndex = 0 Then
                    Exit Sub
                Else
                    BComboList.ListIndex -= 1
                End If
            Case 28, 156
                BComboList.Visible = False
                Call ControlUse(BFormIndex, iFocus, iK)
                'TextList = ""
                Exit Sub
            Case 199
                BComboList.ListIndex = 0
            Case 201
                If BComboList.ListIndex = 0 Then
                    Exit Sub
                ElseIf 1 > BComboList.ListIndex - (BComboList.PositionCount - 17) - 1 Then
                    BComboList.ListIndex = 0
                Else
                    BComboList.ListIndex = BComboList.ListIndex - (BComboList.PositionCount - 17) - 1
                End If
            Case 207
                BComboList.ListIndex = Query(BComboList.DataBase).ListCount
            Case 209
                If Query(BComboList.DataBase).ListCount = BComboList.ListIndex Then
                    Exit Sub
                ElseIf Query(BComboList.DataBase).ListCount < BComboList.ListIndex + (BComboList.PositionCount - 17) + 1 Then
                    BComboList.ListIndex = Query(BComboList.DataBase).ListCount
                Else
                    BComboList.ListIndex = BComboList.ListIndex + (BComboList.PositionCount - 17) + 1
                End If
            Case 1
                BComboList.ScrollIndex = Val(BComboList.Tag)
                BComboList.Visible = False
        End Select
        Select Case iK
            Case 199, 200, 201, 207, 208, 209
                If Query(BComboList.DataBase).ListCount > 0 Then
                    ''BComboList.Position(11).Top = BComboList.Position(10).Top + ((BComboList.Position(10).Bottom - BComboList.Position(10).Top - (10 * FixHeight)) / Query(BComboList.DataBase).ListCount) * BComboList.ListIndex
                    ''BComboList.Position(11).Bottom = BComboList.Position(11).Top + (10 * FixHeight)
                End If
                TextList = Empty
            Case Else
                If (KeyPressed(42) Or KeyPressed(54)) = True Then
                    If iK + 300 > 355 Then
                        Exit Sub
                    End If
                    TextList &= KeyNames(iK + 300)
                Else
                    TextList &= KeyNames(iK)
                End If
                For iA = 0 To Query(BComboList.DataBase).ListCount
                    If UCase(Left(Query(BComboList.DataBase).List(iA), Len(TextList))) = UCase(TextList) And UCase(TextList) <> Empty Then
                        BComboList.ListIndex = iA
                        Exit For
                    End If
                Next iA
        End Select
        '''Call ScreenComboList
    End Sub

    Function KeyTab(iK As Integer) As Boolean
        Dim vA As Integer, iA As Integer, bA As Boolean
        Select Case True
        'Ctrl & Shift & Tab ----------------------------COMPLETED---------------------------
            Case KeyPressed(42) And KeyPressed(29) And (iK = 15 Or KeyPressed(15) = True)
                If DXForm(BFormIndex).Controls(iFocus).Name = "Data1" And (DXForm(BFormIndex).Controls(iFocus).Colom Or DXForm(BFormIndex).Controls(iFocus).Row) > 1 Then
                    iA = iFocus2 - 1
                    If iA > DXForm(BFormIndex).Controls(iFocus).Colom Then
                        KeyTab = True
                        ''Call FocusSet(BFormIndex, iFocus, CByte(iA))
                        Exit Function
                    End If
                End If
                iA = DXForm(BFormIndex).Controls(iFocus).TabIndex - 1
                For vA = DXForm(BFormIndex).ControlsCount To 0 Step -1
                    If DXForm(BFormIndex).Controls(vA).TabIndex = iA Then
                        If DXForm(BFormIndex).Controls(vA).TabStop = False Then
                            iA -= 1
                            vA = DXForm(BFormIndex).ControlsCount + 1
                        ElseIf DXForm(BFormIndex).Controls(vA).Visible = False Then
                            iA -= 1
                            vA = DXForm(BFormIndex).ControlsCount + 1
                        ElseIf DXForm(BFormIndex).Controls(iFocus).TabIndex = iA Then
                            iA -= 1
                            vA = DXForm(BFormIndex).ControlsCount + 1
                        Else
                            KeyTab = True
                            If DXForm(BFormIndex).Controls(vA).Colom > 1 Then
                                ''Call FocusSet(BFormIndex, vA, CByte(DXForm(BFormIndex).Controls(vA).Colom))
                            Else
                                ''Call FocusSet(BFormIndex, vA)
                            End If
                            Exit Function
                        End If
                    End If
                    If iA < 0 Then
                        iA = DXForm(BFormIndex).TabIndexCount - 1
                        If bA = False Then
                            bA = True
                        Else
                            Exit Function
                        End If
                    ElseIf vA = 0 Then
                        If KeyTab = True Then
                            Exit Function
                        Else
                            KeyTab = True
                        End If
                        vA = DXForm(BFormIndex).ControlsCount + 1
                    End If
                Next vA
        ' Tab ----------------------------COMPLETED---------------------------
            Case Not KeyPressed(157) And KeyPressed(29) = False And (iK = 15 Or KeyPressed(15) = True)
                KeyTab = False
        ' Ctrl & Tab ----------------------------COMPLETED---------------------------
            Case Not KeyPressed(157) And KeyPressed(29) = True And (iK = 15 Or KeyPressed(15))
                If DXForm(BFormIndex).Controls(iFocus).Name = "Data1" And (DXForm(BFormIndex).Controls(iFocus).Colom Or DXForm(BFormIndex).Controls(iFocus).Row) > 1 Then
                    iA = iFocus2 + 1
                    If iA <= DXForm(BFormIndex).Controls(iFocus).PositionCount Then
                        If iA <= UBound(DXForm(BFormIndex).Controls(iFocus).Text) Then
                            KeyTab = True
                            ''Call FocusSet(BFormIndex, iFocus, CByte(iA))
                            Exit Function
                        End If
                    End If
                End If
                iA = DXForm(BFormIndex).Controls(iFocus).TabIndex + 1
                For vA = 0 To DXForm(BFormIndex).ControlsCount
                    If DXForm(BFormIndex).Controls(vA).TabIndex = iA Then
                        If DXForm(BFormIndex).Controls(vA).TabStop = False Then
                            iA += 1
                            vA = -1
                        ElseIf DXForm(BFormIndex).Controls(vA).Visible = False Then
                            iA += 1
                            vA = -1
                        ElseIf DXForm(BFormIndex).Controls(iFocus).TabIndex = iA Then
                            iA += 1
                            vA = -1
                        Else
                            KeyTab = True
                            If DXForm(BFormIndex).Controls(vA).Colom > 1 Then
                                ''Call FocusSet(BFormIndex, vA, CByte(DXForm(BFormIndex).Controls(vA).Colom))
                            Else
                                ''Call FocusSet(BFormIndex, vA)
                            End If
                            Exit Function
                        End If
                    End If
                    If iA = DXForm(BFormIndex).TabIndexCount Then
                        iA = 0
                        If bA = False Then
                            bA = True
                        Else
                            Exit Function
                        End If
                    ElseIf vA = DXForm(BFormIndex).ControlsCount Then
                        If KeyTab = True Then
                            Exit Function
                        Else
                            KeyTab = True
                        End If
                        vA = -1
                    End If
                Next vA
        End Select
    End Function

    Sub LoadDirectXInput()
        DInput = DX.DirectInputCreate()
        DInputDeviceMouse = DInput.CreateDevice("GUID_SysMouse")
        Call DInputDeviceMouse.SetCommonDataFormat(DIFORMAT_MOUSE)
        Call DInputDeviceMouse.SetCooperativeLevel(AForm1.hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND)
        Call DInputDeviceMouse.Acquire
        DInputDeviceKeyBoard = DInput.CreateDevice("GUID_SysKeyboard")
        Call DInputDeviceKeyBoard.SetCommonDataFormat(DIFORMAT_KEYBOARD)
        Call DInputDeviceKeyBoard.SetCooperativeLevel(AForm1.hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND)
        Call DInputDeviceKeyBoard.Acquire
    End Sub

    Sub LoadKeyboardKeys()
        KeyNames(2) = "1"
        KeyNames(3) = "2"
        KeyNames(4) = "3"
        KeyNames(5) = "4"
        KeyNames(6) = "5"
        KeyNames(7) = "6"
        KeyNames(8) = "7"
        KeyNames(9) = "8"
        KeyNames(10) = "9"
        KeyNames(11) = "0"
        KeyNames(12) = "-"
        KeyNames(13) = "="
        KeyNames(15) = "    "
        KeyNames(16) = "q"
        KeyNames(17) = "w"
        KeyNames(18) = "e"
        KeyNames(19) = "r"
        KeyNames(20) = "t"
        KeyNames(21) = "y"
        KeyNames(22) = "u"
        KeyNames(23) = "i"
        KeyNames(24) = "o"
        KeyNames(25) = "p"
        KeyNames(26) = "["
        KeyNames(27) = "]"
        KeyNames(28) = vbCrLf
        KeyNames(30) = "a"
        KeyNames(31) = "s"
        KeyNames(32) = "d"
        KeyNames(33) = "f"
        KeyNames(34) = "g"
        KeyNames(35) = "h"
        KeyNames(36) = "j"
        KeyNames(37) = "k"
        KeyNames(38) = "l"
        KeyNames(39) = ";"
        KeyNames(40) = "'"
        KeyNames(41) = "`"
        KeyNames(43) = "\"
        KeyNames(44) = "z"
        KeyNames(45) = "x"
        KeyNames(46) = "c"
        KeyNames(47) = "v"
        KeyNames(48) = "b"
        KeyNames(49) = "n"
        KeyNames(50) = "m"
        KeyNames(51) = ","
        KeyNames(52) = "."
        KeyNames(53) = "/"
        KeyNames(71) = "7"
        KeyNames(72) = "8"
        KeyNames(73) = "9"
        KeyNames(75) = "4"
        KeyNames(76) = "5"
        KeyNames(77) = "6"
        KeyNames(79) = "1"
        KeyNames(80) = "2"
        KeyNames(81) = "3"
        KeyNames(82) = "0"
        KeyNames(83) = "."
        KeyNames(55) = "*"
        KeyNames(57) = " "
        KeyNames(74) = "-"
        KeyNames(78) = "+"
        KeyNames(156) = vbCrLf
        KeyNames(181) = "/"
        KeyNames(302) = "!"
        KeyNames(303) = "@"
        KeyNames(304) = "#"
        KeyNames(305) = "$"
        KeyNames(306) = "%"
        KeyNames(307) = "^"
        KeyNames(308) = "&&"
        KeyNames(309) = "*"
        KeyNames(310) = "("
        KeyNames(311) = ")"
        KeyNames(312) = "_"
        KeyNames(313) = "+"
        KeyNames(316) = "Q"
        KeyNames(317) = "W"
        KeyNames(318) = "E"
        KeyNames(319) = "R"
        KeyNames(320) = "T"
        KeyNames(321) = "Y"
        KeyNames(322) = "U"
        KeyNames(323) = "I"
        KeyNames(324) = "O"
        KeyNames(325) = "P"
        KeyNames(326) = "{"
        KeyNames(327) = "}"
        KeyNames(330) = "A"
        KeyNames(331) = "S"
        KeyNames(332) = "D"
        KeyNames(333) = "F"
        KeyNames(334) = "G"
        KeyNames(335) = "H"
        KeyNames(336) = "J"
        KeyNames(337) = "K"
        KeyNames(338) = "L"
        KeyNames(339) = ":"
        KeyNames(340) = ""
        KeyNames(341) = "~"
        KeyNames(343) = "|"
        KeyNames(344) = "Z"
        KeyNames(345) = "X"
        KeyNames(346) = "C"
        KeyNames(347) = "V"
        KeyNames(348) = "B"
        KeyNames(349) = "N"
        KeyNames(350) = "M"
        KeyNames(351) = "<"
        KeyNames(352) = ">"
        KeyNames(353) = "?"
    End Sub

    Sub MouseDown(iForm As Byte, iControl As Integer, iImage As Byte, iComboListActive As Boolean)
        ''Dim vA As Double
        ''If BComboList.Visible = True And iComboListActive = False Then
        ''    If iForm = 0 Then
        ''        BComboList.Visible = False
        ''    ElseIf (DXForm(iForm).Controls(iControl).Name = "Combo1" And (DXForm(iForm).Controls(iControl).Image(iImage) = 11 Or DXForm(iForm).Controls(iControl).Image(iImage) = 12)) = False Then
        ''        BComboList.Visible = False
        ''    End If
        ''End If
        ''If iComboListActive = True Then
        ''    Select Case BComboList.Image(iImage)
        ''        Case 84, 85, 86, 92
        ''            If Query(BComboList.DataBase).ListCount >= BComboList.ScrollIndex + iImage - 17 Then
        ''                BComboList.ListIndex = BComboList.ScrollIndex + iImage - 17
        ''                Call ScreenComboList
        ''            End If
        ''        Case 13, 17
        ''            If MousePos.t - BComboList.Position(10).Top < 0 Then
        ''                Exit Sub
        ''            ElseIf BComboList.Position(10).Bottom - MousePos.t < 0 Then
        ''                Exit Sub
        ''            End If
        ''            BComboList.Position(11).Top = MousePos.t - (5 * FixHeight)
        ''            BComboList.Position(11).Bottom = (5 * FixHeight) + MousePos.t
        ''            If BComboList.Position(11).Top < BComboList.Position(10).Top Then
        ''                BComboList.ListIndex = 0
        ''                BComboList.Position(11).Top = BComboList.Position(10).Top
        ''                BComboList.Position(11).Bottom = BComboList.Position(10).Top + (10 * FixHeight)
        ''            ElseIf BComboList.Position(11).Bottom > BComboList.Position(10).Bottom Then
        ''                BComboList.ListIndex = Query(BComboList.DataBase).ListCount
        ''                BComboList.Position(11).Top = BComboList.Position(10).Bottom - (10 * FixHeight)
        ''                BComboList.Position(11).Bottom = BComboList.Position(10).Bottom
        ''            Else
        ''                BComboList.ListIndex = CInt((Query(BComboList.DataBase).ListCount / (BComboList.Position(10).Bottom - BComboList.Position(10).Top - (10 * FixHeight))) * (BComboList.Position(11).Top - BComboList.Position(10).Top))
        ''            End If
        ''            Call ScreenComboList
        ''    End Select
        ''ElseIf iForm <> 0 Then
        ''    Select Case DXForm(iForm).Controls(iControl).Image(iImage)
        ''        Case 9 'button
        ''            DXForm(iForm).Controls(iControl).Image(iImage) = 10
        ''            Call FocusSet(iForm, iControl, iImage)
        ''        Case 10 'button
        ''            DXForm(iForm).Controls(iControl).Image(iImage) = 9
        ''            Call FocusSet(iForm, iControl, iImage)
        ''        Case 61, 62, 63 'label
        ''            If DXForm(iForm).Controls(iControl).Appearance = 1 Then
        ''                DXForm(iForm).Controls(iControl).Image(iImage) = 62
        ''                Call FocusSet(iForm, iControl, iImage)
        ''            End If
        ''        Case 74 'textbox
        ''            Call FocusSet(iForm, iControl, iImage)
        ''        Case 84, 85, 86, 92, 116, 117 'listbox and combobox
        ''            Call FocusSet(iForm, iControl, iImage)
        ''            Call ScreenListBox
        ''        Case 11, 12 'list down button
        ''            If DXForm(iForm).Controls(iControl).Name = "List1" Then
        ''                Call FocusSet(iForm, iControl)
        ''                Call ScreenListBox
        ''            ElseIf DXForm(iForm).Controls(iControl).Name = "Combo1" Then
        ''                Call FocusSet(iForm, iControl, iImage)
        ''            End If
        ''        Case 13, 17 'list scrollbutton
        ''            If DXForm(iForm).Controls(iControl).Name = "List1" Then
        ''                If MousePos.t - DXForm(iForm).Controls(iControl).Position(10).Top < 0 Then
        ''                    Exit Sub
        ''                ElseIf DXForm(iForm).Controls(iControl).Position(10).Bottom - MousePos.t < 0 Then
        ''                    Exit Sub
        ''                End If
        ''                DXForm(iForm).Controls(iControl).Position(11).Top = MousePos.t - (5 * FixHeight)
        ''                DXForm(iForm).Controls(iControl).Position(11).Bottom = (5 * FixHeight) + MousePos.t
        ''                If DXForm(iForm).Controls(iControl).Position(11).Top < DXForm(iForm).Controls(iControl).Position(10).Top Then
        ''                    DXForm(iForm).Controls(iControl).ListIndex = 0
        ''                    DXForm(iForm).Controls(iControl).Position(11).Top = DXForm(iForm).Controls(iControl).Position(10).Top
        ''                    DXForm(iForm).Controls(iControl).Position(11).Bottom = DXForm(iForm).Controls(iControl).Position(10).Top + (10 * FixHeight)
        ''                ElseIf DXForm(iForm).Controls(iControl).Position(11).Bottom > DXForm(iForm).Controls(iControl).Position(10).Bottom Then
        ''                    DXForm(iForm).Controls(iControl).ListIndex = Query(DXForm(iForm).Controls(iControl).DataBase).ListCount
        ''                    DXForm(iForm).Controls(iControl).Position(11).Top = DXForm(iForm).Controls(iControl).Position(10).Bottom - (10 * FixHeight)
        ''                    DXForm(iForm).Controls(iControl).Position(11).Bottom = DXForm(iForm).Controls(iControl).Position(10).Bottom
        ''                Else
        ''                    DXForm(iForm).Controls(iControl).ListIndex = CInt((Query(DXForm(iForm).Controls(iControl).DataBase).ListCount / (DXForm(iForm).Controls(iControl).Position(10).Bottom - DXForm(iForm).Controls(iControl).Position(10).Top - (10 * FixHeight))) * (DXForm(iForm).Controls(iControl).Position(11).Top - DXForm(iForm).Controls(iControl).Position(10).Top))
        ''                End If
        ''                Call FocusSet(iForm, iControl)
        ''                Call ScreenListBox
        ''            End If
        ''        Case 15, 16 'list up button
        ''            If DXForm(iForm).Controls(iControl).Name = "List1" Then
        ''                Call FocusSet(iForm, iControl)
        ''                Call ScreenListBox
        ''            End If
        ''    End Select
        ''End If
    End Sub

    Sub MouseOver(iForm As Byte, iControl As Integer, iImage As Byte, iComboListActive As Boolean)

    End Sub

    '---------------------------------------------here i was-------------------------------------????
    Sub MousePress(iForm As Byte, iControl As Integer, iImage As Byte, iComboListActive As Boolean)
        Dim vA As Integer
        Static iTimer As Double, iImageUpDown As Byte
        If iTimer < GetTickCount Then
            iTimer = GetTickCount + 75
            If BComboList.Visible = True Then
                If DXForm(BFormIndex).Controls(iFocus).Name = "Combo1" Then
                    If UBound(DXForm(BFormIndex).Controls(iFocus).Image) >= iImageUpDown Then
                        If iForm <> 0 And iForm = BFormIndex And iControl = iFocus And iImage = iImageUpDown Then
                            Select Case DXForm(BFormIndex).Controls(iFocus).Image(iImageUpDown)
                                Case 11, 12 'list down button
                                    DXForm(BFormIndex).Controls(iFocus).Image(iImageUpDown) = 12
                                Case 15, 16 'list up button
                                    DXForm(BFormIndex).Controls(iFocus).Image(iImageUpDown) = 16
                            End Select
                        ElseIf BFormIndex <> 0 Then
                            Select Case DXForm(BFormIndex).Controls(iFocus).Image(iImageUpDown)
                                Case 11, 12 'list down button
                                    DXForm(BFormIndex).Controls(iFocus).Image(iImageUpDown) = 11
                                Case 15, 16 'list up button
                                    DXForm(BFormIndex).Controls(iFocus).Image(iImageUpDown) = 15
                            End Select
                        End If
                    End If
                End If
                Select Case BComboList.Image(iImageUpDown)
                    Case 11, 12 'list down button
                        BComboList.Image(iImageUpDown) = 11
                    Case 15, 16 'list up button
                        BComboList.Image(iImageUpDown) = 15
                End Select
                If iComboListActive = True Then
                    Select Case BComboList.Image(iImage)
                        Case 84, 85, 86, 92
                            If Query(BComboList.DataBase).ListCount >= BComboList.ScrollIndex + iImage - 17 Then
                                BComboList.ListIndex = BComboList.ScrollIndex + iImage - 17
                                Call ScreenComboList()
                            End If
                        Case 11, 12
                            BComboList.Image(iImage) = 12
                            iImageUpDown = iImage
                            vA = BComboList.PositionCount - 17
                            If Query(BComboList.DataBase).ListCount = BComboList.ListIndex Then
                                Exit Sub
                            ElseIf Query(BComboList.DataBase).ListCount < BComboList.ListIndex + vA + 1 Then
                                BComboList.ListIndex = Query(BComboList.DataBase).ListCount
                            Else
                                BComboList.ListIndex = BComboList.ListIndex + vA + 1
                            End If
                            Call ScreenComboList()
                        Case 13, 17
                            If MousePos.t - BComboList.Position(10).Top < 0 Then
                                Exit Sub
                            ElseIf BComboList.Position(10).Bottom - MousePos.t < 0 Then
                                Exit Sub
                            End If
                            BComboList.Position(11).Top = MousePos.t - (5 * FixHeight)
                            BComboList.Position(11).Bottom = (5 * FixHeight) + MousePos.t
                            If BComboList.Position(11).Top < BComboList.Position(10).Top Then
                                BComboList.ListIndex = 0
                                BComboList.Position(11).Top = BComboList.Position(10).Top
                                BComboList.Position(11).Bottom = BComboList.Position(10).Top + (10 * FixHeight)
                            ElseIf BComboList.Position(11).Bottom > BComboList.Position(10).Bottom Then
                                BComboList.ListIndex = Query(BComboList.DataBase).ListCount
                                BComboList.Position(11).Top = BComboList.Position(10).Bottom - (10 * FixHeight)
                                BComboList.Position(11).Bottom = BComboList.Position(10).Bottom
                            Else
                                BComboList.ListIndex = CInt((Query(BComboList.DataBase).ListCount / (BComboList.Position(10).Bottom - BComboList.Position(10).Top - (10 * FixHeight))) * (BComboList.Position(11).Top - BComboList.Position(10).Top))
                            End If
                            Call ScreenComboList()
                        Case 15, 16
                            BComboList.Image(iImage) = 16
                            iImageUpDown = iImage
                            vA = BComboList.PositionCount - 17
                            If 0 = BComboList.ListIndex Then
                                Exit Sub
                            ElseIf 0 > BComboList.ListIndex - vA - 1 Then
                                BComboList.ListIndex = 0
                            Else
                                BComboList.ListIndex = BComboList.ListIndex - vA - 1
                            End If
                            Call ScreenComboList()
                    End Select
                End If
            ElseIf iForm <> 0 And iForm = BFormIndex And iControl = iFocus Then
                If DXForm(iForm).Controls(iControl).Name = "List1" Then
                    Select Case DXForm(iForm).Controls(iControl).Image(iImageUpDown)
                        Case 11, 12 'list down button
                            DXForm(iForm).Controls(iControl).Image(iImageUpDown) = 11
                        Case 15, 16 'list up button
                            DXForm(iForm).Controls(iControl).Image(iImageUpDown) = 15
                    End Select
                End If
                Select Case DXForm(iForm).Controls(iControl).Image(iImage)
                    Case 84, 85, 86, 92, 116, 117 'listbox and combobox
                        Call FocusSet(iForm, iControl, iImage)
                        Call ScreenListBox()
                    Case 61, 62, 63 'label
                        If DXForm(iForm).Controls(iControl).Appearance = 1 Then
                            DXForm(iForm).Controls(iControl).Image(iFocus2) = 62
                        End If
                    Case 9, 10 'button
                        If DXForm(iForm).Controls(iControl).Name <> "Option1" And DXForm(iForm).Controls(iControl).Name <> "Check1" Then
                            DXForm(iForm).Controls(iControl).Image(iFocus2) = 10
                        ElseIf iCheckOption = 0 Then
                            iCheckOption = DXForm(iForm).Controls(iControl).Image(iFocus2)
                        Else
                            DXForm(iForm).Controls(iControl).Image(iFocus2) = iCheckOption
                        End If
                    Case 11, 12 'list down button
                        DXForm(iForm).Controls(iControl).Image(iImage) = 12
                        iImageUpDown = iImage
                        If DXForm(iForm).Controls(iControl).Name = "List1" Then
                            vA = DXForm(iForm).Controls(iControl).PositionCount - 17
                            If Query(DXForm(iForm).Controls(iControl).DataBase).ListCount = DXForm(iForm).Controls(iControl).ListIndex Then
                                Call FocusSet(iForm, iControl)
                                Exit Sub
                            ElseIf Query(DXForm(iForm).Controls(iControl).DataBase).ListCount < DXForm(iForm).Controls(iControl).ListIndex + vA + 1 Then
                                DXForm(iForm).Controls(iControl).ListIndex = Query(DXForm(iForm).Controls(iControl).DataBase).ListCount
                            Else
                                DXForm(iForm).Controls(iControl).ListIndex = DXForm(iForm).Controls(iControl).ListIndex + vA + 1
                            End If
                            Call ScreenListBox()
                        ElseIf DXForm(iForm).Controls(iControl).Name = "Combo1" Then
                            BComboList.Visible = True
                            BComboList.Tag = DXForm(iForm).Controls(iControl).ScrollIndex
                            Call xForm(0).ComboList(0).Move(CSng(DXForm(iForm).Controls(iControl).Position(0).Left), CSng(DXForm(iForm).Controls(iControl).Position(6).Bottom), DXForm(iForm).Controls(iControl).X2, 150 * FixHeight, False)
                            xForm(0).ComboList(0).DataBase = DXForm(iForm).Controls(iControl).DataSource
                            Call ScreenComboList()
                        End If
                    Case 13, 17 'list scrollbutton
                        If DXForm(iForm).Controls(iControl).Name = "List1" Then
                            If MousePos.t - DXForm(iForm).Controls(iControl).Position(10).Top < 0 Then
                                Exit Sub
                            ElseIf DXForm(iForm).Controls(iControl).Position(10).Bottom - MousePos.t < 0 Then
                                Exit Sub
                            End If
                            DXForm(iForm).Controls(iControl).Position(11).Top = MousePos.t - (5 * FixHeight)
                            DXForm(iForm).Controls(iControl).Position(11).Bottom = (5 * FixHeight) + MousePos.t
                            If DXForm(iForm).Controls(iControl).Position(11).Top < DXForm(iForm).Controls(iControl).Position(10).Top Then
                                DXForm(iForm).Controls(iControl).ListIndex = 0
                                DXForm(iForm).Controls(iControl).Position(11).Top = DXForm(iForm).Controls(iControl).Position(10).Top
                                DXForm(iForm).Controls(iControl).Position(11).Bottom = DXForm(iForm).Controls(iControl).Position(10).Top + (10 * FixHeight)
                            ElseIf DXForm(iForm).Controls(iControl).Position(11).Bottom > DXForm(iForm).Controls(iControl).Position(10).Bottom Then
                                DXForm(iForm).Controls(iControl).ListIndex = Query(DXForm(iForm).Controls(iControl).DataBase).ListCount
                                DXForm(iForm).Controls(iControl).Position(11).Top = DXForm(iForm).Controls(iControl).Position(10).Bottom - (10 * FixHeight)
                                DXForm(iForm).Controls(iControl).Position(11).Bottom = DXForm(iForm).Controls(iControl).Position(10).Bottom
                            Else
                                DXForm(iForm).Controls(iControl).ListIndex = CInt((Query(DXForm(iForm).Controls(iControl).DataBase).ListCount / (DXForm(iForm).Controls(iControl).Position(10).Bottom - DXForm(iForm).Controls(iControl).Position(10).Top - (10 * FixHeight))) * (DXForm(iForm).Controls(iControl).Position(11).Top - DXForm(iForm).Controls(iControl).Position(10).Top))
                            End If
                            Call ScreenListBox()
                        End If
                    Case 15, 16 'list up button
                        DXForm(iForm).Controls(iControl).Image(iImage) = 16
                        iImageUpDown = iImage
                        If DXForm(iForm).Controls(iControl).Name = "List1" Then
                            vA = CInt(DXForm(iForm).Controls(iControl).PositionCount - 17)
                            If 1 > DXForm(iForm).Controls(iControl).ListIndex - vA - 1 Then
                                DXForm(iForm).Controls(iControl).ListIndex = 0
                            Else
                                DXForm(iForm).Controls(iControl).ListIndex = DXForm(iForm).Controls(iControl).ListIndex - vA - 1
                            End If
                            Call ScreenListBox()
                        End If
                End Select
            ElseIf BFormIndex <> 0 Then
                If DXForm(BFormIndex).Controls(iFocus).Name = "List1" Then
                    Select Case DXForm(BFormIndex).Controls(iFocus).Image(iImageUpDown)
                        Case 11, 12 'list down button
                            DXForm(BFormIndex).Controls(iFocus).Image(iImageUpDown) = 11
                        Case 15, 16 'list up button
                            DXForm(BFormIndex).Controls(iFocus).Image(iImageUpDown) = 15
                    End Select
                End If
                If DXForm(BFormIndex).Controls(iFocus).Name = "Command1" Or DXForm(BFormIndex).Controls(iFocus).Name = "Check1" Or DXForm(BFormIndex).Controls(iFocus).Name = "Option1" Or DXForm(BFormIndex).Controls(iFocus).Name = "Label1" Then
                    Select Case DXForm(BFormIndex).Controls(iFocus).Image(iFocus2)
                        Case 61, 62, 63 'label
                            If DXForm(BFormIndex).Controls(iFocus).Appearance = 1 Then
                                DXForm(BFormIndex).Controls(iFocus).Image(iFocus2) = 63
                            End If
                        Case 9, 10 'button
                            If DXForm(BFormIndex).Controls(iFocus).Name <> "Option1" And DXForm(BFormIndex).Controls(iFocus).Name <> "Check1" Then
                                DXForm(BFormIndex).Controls(iFocus).Image(iFocus2) = 9
                            ElseIf iCheckOption = 9 Then
                                DXForm(BFormIndex).Controls(iFocus).Image(iFocus2) = 10
                            ElseIf iCheckOption = 10 Then
                                DXForm(BFormIndex).Controls(iFocus).Image(iFocus2) = 9
                            End If
                    End Select
                End If
            End If
        End If
    End Sub

    Sub MouseUp(iForm As Byte, iControl As Integer, iImage As Byte, iComboListActive As Boolean)
        Dim vA As Integer, iControlOld As Integer
        If BComboList.Visible = True Then
            BComboList.Image(12) = 11
            BComboList.Image(9) = 15
            If iComboListActive = True Then
                Select Case BComboList.Image(iImage)
                    Case 86
                        BComboList.Visible = False
                        Call ControlUse(BFormIndex, iFocus, 28)
                End Select
                Exit Sub
            End If
        End If
        If iForm = BFormIndex And iControl = iFocus Then
            Select Case DXForm(iForm).Controls(iControl).Image(iImage)
                Case 9, 10 'button
                    If DXForm(iForm).Controls(iControl).Name = "Command1" Then
                        DXForm(iForm).Controls(iControl).Image(iImage) = 9
                        Call ControlUse(iForm, iControl, 28)
                    ElseIf DXForm(iForm).Controls(iControl).Name = "Option1" Then
                        For vA = 0 To DXForm(iForm).ControlsCount
                            If vA <> iControl Then
                                If DXForm(iForm).Controls(vA).Name = "Option1" Then
                                    If DXForm(iForm).Controls(vA).DataMember = DXForm(iForm).Controls(iControl).DataMember Then
                                        DXForm(iForm).Controls(vA).Value = False
                                        DXForm(iForm).Controls(vA).Image(iImage) = 9
                                    End If
                                End If
                            End If
                        Next vA
                        If DXForm(iForm).Controls(iControl).Image(iImage) = 9 Then
                            DXForm(iForm).Controls(iControl).Value = False
                            If DXForm(iForm).Controls(iControl).Text(0) = "Yes" Or DXForm(iForm).Controls(iControl).Text(0) = "No" Or DXForm(iForm).Controls(iControl).Text(0) = Empty Then
                                DXForm(iForm).Controls(iControl).Text(0) = "No"
                            End If
                        Else
                            DXForm(iForm).Controls(iControl).Value = True
                            If DXForm(iForm).Controls(iControl).Text(0) = "Yes" Or DXForm(iForm).Controls(iControl).Text(0) = "No" Or DXForm(iForm).Controls(iControl).Text(0) = Empty Then
                                DXForm(iForm).Controls(iControl).Text(0) = "Yes"
                            End If
                        End If
                        Call ControlUse(iForm, iControl, -1)
                    ElseIf DXForm(iForm).Controls(iControl).Name = "Check1" Then
                        If DXForm(iForm).Controls(iControl).Image(iImage) = 9 Then
                            DXForm(iForm).Controls(iControl).Value = False
                            If DXForm(iForm).Controls(iControl).Text(0) = "Yes" Or DXForm(iForm).Controls(iControl).Text(0) = "No" Or DXForm(iForm).Controls(iControl).Text(0) = Empty Then
                                DXForm(iForm).Controls(iControl).Text(0) = "No"
                            End If
                        Else
                            DXForm(iForm).Controls(iControl).Value = True
                            If DXForm(iForm).Controls(iControl).Text(0) = "Yes" Or DXForm(iForm).Controls(iControl).Text(0) = "No" Or DXForm(iForm).Controls(iControl).Text(0) = Empty Then
                                DXForm(iForm).Controls(iControl).Text(0) = "Yes"
                            End If
                        End If
                        Call ControlUse(iForm, iControl, -1)
                    End If
                Case 61, 62, 63
                    If DXForm(iForm).Controls(iControl).Appearance = 1 Then
                        DXForm(iForm).Controls(iControl).Image(iImage) = 63
                        Call ControlUse(iForm, iControl, -1)
                    End If
                Case 93, 94, 95
                    If DXForm(iForm).Controls(iControl).Appearance = 1 Then
                        DXForm(iForm).Controls(iControl).Image(iImage) = 93
                        Call ControlUse(iForm, iControl, -1)
                    End If
                Case 84, 85, 86, 92, 116, 117
                    Call FocusSet(iForm, iControl, iImage)
                    If MouseDblClickFocus2 = iImage And MouseDblClickFocus = iControl And MouseDblClickTime > GetTickCount Then
                        Call ControlUse(iForm, iControl, 28)
                    Else
                        Call ControlUse(iForm, iControl, 82)
                    End If
                    If MouseDblClickTime < GetTickCount Then
                        MouseDblClickTime = 500 + GetTickCount
                        MouseDblClickFocus = iControl
                        MouseDblClickFocus2 = iImage
                    End If
                Case 11, 12 'list down button
                    DXForm(BFormIndex).Controls(iFocus).Image(iImage) = 11
                Case 15, 16 'list up button
                    DXForm(BFormIndex).Controls(iFocus).Image(iImage) = 15
            End Select
        End If
        iCheckOption = 0
    End Sub

    Sub RenderKeys()
        Dim iA As Integer, iB As Integer, yA As Byte
        Static vY As Double, vZ As Double
        Call DInputDeviceKeyBoard.GetDeviceStateKeyboard(DStateKeyBoard)
        For iA = 0 To 255
            If DStateKeyBoard.Key(iA) <> 0 Then
                Select Case iA
                    Case 1, 210, 58, 69, 70, 183, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 87, 88, 56, 42, 183, 184, 157, 54, 29
                        iB = iB + 1
                End Select
                If iFocus <> -1 Then
                    If KeyPressed(iA) <> True Then
                        Call KeyLink(iA)
                    ElseIf KeyPressed(iA) = True And yA < 12 And iA <> 42 And iA <> 54 Then
                        yA = yA + 1
                    ElseIf yA > 11 Then
                        Call KeyLink(iA)
                    End If
                End If
                KeyPressed(iA) = True
            Else
                iB = iB + 1
                If KeyPressed(iA) <> False Then
                    yA = 0
                End If
                KeyPressed(iA) = False
            End If
        Next iA
        If iB = 256 Then
            vZ = Empty
        ElseIf vZ = Empty Then
            vZ = GetTickCount + 500
        End If
        If vZ <> Empty And vZ < GetTickCount Then
            If vY < GetTickCount Then
                vY = GetTickCount + 50
                For iA = 0 To 41
                    KeyPressed(iA) = False
                Next iA
                For iA = 43 To 255
                    KeyPressed(iA) = False
                Next iA
            End If
        End If
    End Sub

    Sub RenderMouse()
        Dim iForm As Byte, iControl As Integer, iImage As Byte, iButton As Byte
        Dim iFormActive As Byte, iControlActive As Integer, iImageActive As Byte
        Dim iComboListActive As Boolean
        'muis knoppen opvragen
        Call DInputDeviceMouse.GetDeviceStateMouse(DStateMouse)
        'muis positie opvragen
        Call GetCursorPos(MousePos)
        'controleer of er binnen een combolist gewerkt wordt en kijk dan welke element er gebruikt wordt
        If BComboList.Visible = True Then
            If MousePos.l >= BComboList.X1 And MousePos.t >= BComboList.Y1 And MousePos.l <= (BComboList.X2 + BComboList.X1) And MousePos.t <= (BComboList.Y2 + BComboList.Y1) Then
                iComboListActive = True
            End If
            If iComboListActive = True Then
                For iImage = 0 To BComboList.PositionCount
                    If MousePos.l >= BComboList.Position(iImage).Left And MousePos.t >= BComboList.Position(iImage).Top And MousePos.l <= BComboList.Position(iImage).Right And MousePos.t <= BComboList.Position(iImage).Bottom And iImage <> 4 Then
                        Select Case BComboList.Image(iImage)
                            Case 11, 12, 13, 15, 16, 17, 84, 85, 86, 92
                                iImageActive = iImage
                                Exit For
                        End Select
                    End If
                Next iImage
            End If
        End If
        'Als de message box er is zorg dan dat de rest niet te besturen is
        If vMsgBoxEnabled = True Then
            For iForm = 1 To BFormCount
                If DXForm(iForm).Visible = True And DXForm(iForm).Name = "MsgBox" Then
                    For iControl = DXForm(iForm).ControlsCount To 0 Step -1
                        If DXForm(iForm).Controls(iControl).Visible = True Then
                            For iImage = 0 To DXForm(iForm).Controls(iControl).PositionCount
                                If MousePos.l >= DXForm(iForm).Controls(iControl).Position(iImage).Left And MousePos.t >= DXForm(iForm).Controls(iControl).Position(iImage).Top And MousePos.l <= DXForm(iForm).Controls(iControl).Position(iImage).Right And MousePos.t <= DXForm(iForm).Controls(iControl).Position(iImage).Bottom Then
                                    Select Case DXForm(iForm).Controls(iControl).Image(iImage)
                                        Case 9, 10, 11, 12, 13, 15, 16, 17, 61, 62, 63, 74, 84, 85, 86, 92, 116, 117
                                            If (DXForm(iForm).Controls(iControl).Name = "List1" And iImage = 4) = False Then
                                                iFormActive = iForm
                                                iControlActive = iControl
                                                iImageActive = iImage
                                                Exit For
                                            End If
                                    End Select
                                End If
                            Next iImage
                            If iFormActive <> 0 Then
                                Exit For
                            End If
                        End If
                    Next iControl
                    If iFormActive <> 0 Then
                        Exit For
                    End If
                End If
            Next iForm
            'controleer boven welke element de muis actie gebeurt
        ElseIf iComboListActive = False Then
            For iForm = 1 To BFormCount
                If DXForm(iForm).Visible = True Then
                    For iControl = DXForm(iForm).ControlsCount To 0 Step -1
                        If DXForm(iForm).Controls(iControl).Visible = True Then
                            For iImage = 0 To DXForm(iForm).Controls(iControl).PositionCount
                                If MousePos.l >= DXForm(iForm).Controls(iControl).Position(iImage).Left And MousePos.t >= DXForm(iForm).Controls(iControl).Position(iImage).Top And MousePos.l <= DXForm(iForm).Controls(iControl).Position(iImage).Right And MousePos.t <= DXForm(iForm).Controls(iControl).Position(iImage).Bottom Then
                                    Select Case DXForm(iForm).Controls(iControl).Image(iImage)
                                        Case 9, 10, 11, 12, 13, 15, 16, 17, 61, 62, 63, 74, 84, 85, 86, 92, 116, 117
                                            If (DXForm(iForm).Controls(iControl).Name = "List1" And iImage = 4) = False Then
                                                iFormActive = iForm
                                                iControlActive = iControl
                                                iImageActive = iImage
                                                Exit For
                                            End If
                                    End Select
                                End If
                            Next iImage
                            If iFormActive <> 0 Then
                                Exit For
                            End If
                        End If
                    Next iControl
                    If iFormActive <> 0 Then
                        Exit For
                    End If
                End If
            Next iForm
        End If
        'controleer of de control wel aan staat om te kunnen gebruiken
        If DXForm(BFormIndex).Controls(iFocus).Enabled = False Then
            Exit Sub
        End If
        'controleer wat de scrolwiel eventueel doet en roep de bijpassende functie aan
        If DStateMouse.lZ <> 0 Then
            If Left(DStateMouse.lZ, 1) = "-" Then
                If DXForm(BFormIndex).Controls(iFocus).Name = "List1" Then
                    Call KeyListBox(209)
                ElseIf DXForm(BFormIndex).Controls(iFocus).Name = "Combo1" Then
                    Call KeyComboBox(208)
                ElseIf BComboList.Visible = True Then
                    Call KeyListBoxCombo(209)
                End If
            Else
                If DXForm(BFormIndex).Controls(iFocus).Name = "List1" Then
                    Call KeyListBox(201)
                ElseIf DXForm(BFormIndex).Controls(iFocus).Name = "Combo1" Then
                    Call KeyComboBox(200)
                ElseIf BComboList.Visible = True Then
                    Call KeyListBoxCombo(201)
                End If
            End If
        End If
        'controleer of de control wel aan staat om te kunnen gebruiken
        If iFormActive <> 0 Then
            If DXForm(iFormActive).Controls(iControlActive).Enabled = False Then
                Exit Sub
            End If
        End If
        'controleer de eventuele veranderingen van de muis en roep de bijpassende functie aan
        For iButton = 0 To 1
            If DStateMouse.Buttons(iButton) <> 0 Then
                If MouseButton(iButton) = False Then
                    MouseButton(iButton) = True
                    'muis indrukken
                    Call MouseDown(iFormActive, iControlActive, iImageActive, iComboListActive)
                Else
                    'muis blijft ingedrukt
                    Call MousePress(iFormActive, iControlActive, iImageActive, iComboListActive)
                End If
            Else
                If MouseButton(iButton) = True Then
                    MouseButton(iButton) = False
                    'muis loslaten
                    Call MouseUp(iFormActive, iControlActive, iImageActive, iComboListActive)
                Else
                    'muis over
                    Call MouseOver(iFormActive, iControlActive, iImageActive, iComboListActive)
                End If
            End If
        Next iButton
    End Sub

End Module
