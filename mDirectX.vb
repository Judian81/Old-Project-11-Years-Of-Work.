Module mDirectX

    'all code is implented

    'this is a control. in this case it is about a combobox. it have to react the same like the combobox in windows do.
    'it is only more optimized then the default visual basic controles. it do not need as much memmory as the official way
    'it was a long time ago when i made this. back then it was one of the best
    'but this is in directx and it is build up from the ground to create windows like controls

    'this is a combobox that have the same controle like the default windows combobox does.
    'the big adventure is that they way it controls with a database attached to it
    Sub KeyComboBox(iK As Integer)
        Dim iA As Integer, iB As Integer, iC As Integer
        If DXForm(BFormIndex).Controls(iFocus).FillStyle = 5 Then
            DXForm(BFormIndex).Controls(iFocus).DataSource = DXForm(BFormIndex).Controls(iFocus).DataBase
        End If
        Select Case iK
            Case 199, 200, 201, 207, 208, 209
                TextList = ""
                Select Case iK
                    Case 208
                        If BComboList.Visible = False Then
                            BComboList.Visible = True
                            BComboList.ListIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex
                            BComboList.Tag = DXForm(BFormIndex).Controls(iFocus).ScrollIndex
                            If DXForm(BFormIndex).Controls(iFocus).Name = "Combo1" Then
                                Call xForm(0).ComboList(0).Move(CSng(DXForm(BFormIndex).Controls(iFocus).Position(0).Left), CSng(DXForm(BFormIndex).Controls(iFocus).Position(6).Bottom), DXForm(BFormIndex).Controls(iFocus).X2, 150 * FixHeight, False)
                            ElseIf DXForm(BFormIndex).Controls(iFocus).PositionCount = 2 Then
                                Call xForm(0).ComboList(0).Move(CSng(DXForm(BFormIndex).Controls(iFocus).Position(1).Left), CSng(DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom), DXForm(BFormIndex).Controls(iFocus).X2, 150 * FixHeight, False)
                            Else
                                Call xForm(0).ComboList(0).Move(CSng(DXForm(BFormIndex).Controls(iFocus).Position(3).Left), CSng(DXForm(BFormIndex).Controls(iFocus).Position(6).Bottom), DXForm(BFormIndex).Controls(iFocus).X2, 150 * FixHeight, False)
                            End If
                            xForm(0).ComboList(0).DataBase = DXForm(BFormIndex).Controls(iFocus).DataSource
                            Call ScreenComboList()
                        ElseIf Query(DXForm(BFormIndex).Controls(iFocus).DataSource).ListCount = DXForm(BFormIndex).Controls(iFocus).ListIndex Then
                            Exit Sub
                        Else
                            DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex + 1
                        End If
                    Case 200
                        If BComboList.Visible = False Then
                            BComboList.Visible = True
                            BComboList.ListIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex
                            BComboList.Tag = DXForm(BFormIndex).Controls(iFocus).ScrollIndex
                            If DXForm(BFormIndex).Controls(iFocus).Name = "Combo1" Then
                                Call xForm(0).ComboList(0).Move(CSng(DXForm(BFormIndex).Controls(iFocus).Position(0).Left), CSng(DXForm(BFormIndex).Controls(iFocus).Position(6).Bottom), DXForm(BFormIndex).Controls(iFocus).X2, 150 * FixHeight, False)
                            ElseIf DXForm(BFormIndex).Controls(iFocus).PositionCount = 2 Then
                                Call xForm(0).ComboList(0).Move(CSng(DXForm(BFormIndex).Controls(iFocus).Position(1).Left), CSng(DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom), DXForm(BFormIndex).Controls(iFocus).X2, 150 * FixHeight, False)
                            Else
                                Call xForm(0).ComboList(0).Move(CSng(DXForm(BFormIndex).Controls(iFocus).Position(3).Left), CSng(DXForm(BFormIndex).Controls(iFocus).Position(6).Bottom), DXForm(BFormIndex).Controls(iFocus).X2, 150 * FixHeight, False)
                            End If
                            xForm(0).ComboList(0).DataBase = DXForm(BFormIndex).Controls(iFocus).DataSource
                            Call ScreenComboList()
                        ElseIf DXForm(BFormIndex).Controls(iFocus).ListIndex = 0 Then
                            Exit Sub
                        Else
                            DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex - 1
                        End If
                    Case 199
                        DXForm(BFormIndex).Controls(iFocus).ListIndex = 0
                    Case 201
                        If DXForm(BFormIndex).Controls(iFocus).ListIndex = 0 Then
                            Exit Sub
                        ElseIf DXForm(BFormIndex).Controls(iFocus).ListIndex = 1 Then
                            DXForm(BFormIndex).Controls(iFocus).ListIndex = 0
                        Else
                            DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex - 2
                        End If
                    Case 207
                        DXForm(BFormIndex).Controls(iFocus).ListIndex = Query(DXForm(BFormIndex).Controls(iFocus).DataSource).ListCount
                    Case 209
                        If DXForm(BFormIndex).Controls(iFocus).ListIndex = Query(DXForm(BFormIndex).Controls(iFocus).DataSource).ListCount Then
                            Exit Sub
                        ElseIf DXForm(BFormIndex).Controls(iFocus).ListIndex = Query(DXForm(BFormIndex).Controls(iFocus).DataSource).ListCount - 1 Then
                            DXForm(BFormIndex).Controls(iFocus).ListIndex = Query(DXForm(BFormIndex).Controls(iFocus).DataSource).ListCount
                        Else
                            DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex + 2
                        End If
                End Select
            Case 28, 156
                TextList = ""
                If DXForm(BFormIndex).Controls(iFocus).Appearance = 999 Then
                    iB = DataNR(ScreenDataBase(BFormIndex, iFocus))
                    For iC = DXForm(BFormIndex).ControlsCount To 0 Step -1
                        If DXForm(BFormIndex).Controls(iC).Visible <> 0 Then
                            If DXForm(BFormIndex).Controls(iC).Parent = DXForm(BFormIndex).Controls(iFocus).Parent Then
                                If DXForm(BFormIndex).Controls(iC).FillStyle <> 0 Then
                                    Select Case DXForm(BFormIndex).Controls(iC).Image(0)
                                        Case 9, 10, 11, 13, 15, 74, 84, 85, 86, 92, 93, 116
                                            If DXForm(BFormIndex).Controls(iC).DataBase = DXForm(BFormIndex).Controls(iFocus + 13 - DXForm(BFormIndex).Controls(iFocus + 13).Index).DataBase Then
                                                If DXForm(BFormIndex).Controls(iC).Image(0) = 13 Then
                                                    If Query(iB).ListCount = "" Then
                                                        DXForm(BFormIndex).Controls(iC).Position(0).Top = DXForm(BFormIndex).Controls(iC - 1).Position(0).Top
                                                        DXForm(BFormIndex).Controls(iC).Position(0).Bottom = DXForm(BFormIndex).Controls(iC).Position(0).Top + 10
                                                    Else
                                                        DXForm(BFormIndex).Controls(iC).Position(0).Top = DXForm(BFormIndex).Controls(iC - 1).Position(0).Top + ((DXForm(BFormIndex).Controls(iC - 1).Position(0).Bottom - DXForm(BFormIndex).Controls(iC - 1).Position(0).Top - 10) / Query(iB).ListCount) * DXForm(BFormIndex).Controls(iC).ListIndex
                                                        DXForm(BFormIndex).Controls(iC).Position(0).Bottom = DXForm(BFormIndex).Controls(iC).Position(0).Top + 10
                                                    End If
                                                ElseIf DXForm(BFormIndex).Controls(iC).Image(0) = 84 Or DXForm(BFormIndex).Controls(iC).Image(0) = 85 Or DXForm(BFormIndex).Controls(iC).Image(0) = 86 Then
                                                    If DXForm(BFormIndex).Controls(iC).ListIndex - DXForm(BFormIndex).Controls(iC).ScrollIndex = DXForm(BFormIndex).Controls(iC).Index Then
                                                        If DXForm(BFormIndex).Controls(iFocus).DataBase = iB Then
                                                            DXForm(BFormIndex).Controls(iC).Image(0) = 86
                                                        Else
                                                            DXForm(BFormIndex).Controls(iC).Image(0) = 85
                                                        End If
                                                    Else
                                                        DXForm(BFormIndex).Controls(iC).Image(0) = 84
                                                    End If
                                                End If
                                                If Val(iFocus + 13 - DXForm(BFormIndex).Controls(iFocus + 13).Index) <> iC Then
                                                    DXForm(BFormIndex).Controls(iC).DataBase = iB
                                                End If
                                            End If
                                    End Select
                                End If
                            End If
                        End If
                    Next iC
                    DXForm(BFormIndex).Controls(iFocus + 13 - DXForm(BFormIndex).Controls(iFocus + 13).Index).DataSource = Left(DXForm(BFormIndex).Controls(iFocus + 13 - DXForm(BFormIndex).Controls(iFocus + 13).Index).DataSource, 2) & Right("00000" & iB, 5)
                    DXForm(BFormIndex).Controls(iFocus + 13 - DXForm(BFormIndex).Controls(iFocus + 13).Index).DataBase = iB
                End If
                Call ControlUse(BFormIndex, iFocus, iK)
            Case 44 And (KeyPressed(29) Or KeyPressed(157))
                TextList = ""
        ' Tab Ctrl ----------------------------COMPLETED---------------------------
            Case 15 And (KeyPressed(29) Or KeyPressed(157))
                'KEYTAB
                'EINDE
                Exit Sub
            Case Else
                If DXForm(BFormIndex).Controls(iFocus).Name = "Combo1" Then
                    Exit Sub
                End If
                If DXForm(BFormIndex).Controls(iFocus).Name = "Data1" Then
                    Call KeyTextBox(iK)
                    Exit Sub
                End If
                If (KeyPressed(42) Or KeyPressed(54)) = True Then
                    TextList = TextList & KeyNames(iK + 300)
                Else
                    TextList = TextList & KeyNames(iK)
                End If
                For iA = 0 To Query(DXForm(BFormIndex).Controls(iFocus).DataSource).ListCount
                    If UCase(Left(Query(DXForm(BFormIndex).Controls(iFocus).DataSource).List(iA), Len(TextList))) = UCase(TextList) And UCase(TextList) <> "" Then
                        DXForm(BFormIndex).Controls(iFocus).ListIndex = iA
                        Exit For
                    End If
                Next iA
        End Select
        If iK <> 28 Then
            DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex
            If DXForm(BFormIndex).Controls(iFocus).FillStyle = 5 Then
                DXForm(BFormIndex).Controls(iFocus).DataSource = Query(DXForm(BFormIndex).Controls(iFocus).DataBase).List(DXForm(BFormIndex).Controls(iFocus).ListIndex)
            End If
            xForm(BFormIndex).Controls(iFocus).Text = Query(DXForm(BFormIndex).Controls(iFocus).DataSource).List(DXForm(BFormIndex).Controls(iFocus).ScrollIndex)
            If DXForm(BFormIndex).Controls(iFocus).Name = "Data1" Then
                Call xForm(BFormIndex).Data1(DXForm(BFormIndex).Controls(iFocus).Index).Change
            Else
                Call xForm(BFormIndex).Combo1(DXForm(BFormIndex).Controls(iFocus).Index).Change
            End If
        End If
    End Sub



    Sub KeyTextBox(iK As Integer)
        Dim iA As Long, iB As Long, iC As Long, sA As String, sB As String
        'stel alles terug op nul als er geen tekst is
        If TextCursorString = "" Then
            TextSelectedCount = 0
            TextCursorPosition = 0
        End If
        'regel hier de verticale scroll mee met een tekstbox
        If DXForm(BFormIndex).Controls(iFocus).MultiLine = True And DXForm(BFormIndex).Controls(iFocus).ListCount <> 0 Then
            ' Home & Shift & Ctrl ----------------------------COMPLETED---------------------------
            If iK = 199 And (KeyPressed(42) Or KeyPressed(54)) = True And (KeyPressed(29) Or KeyPressed(157)) = True Then
                'SCROLLCOUNT
                DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                'LISTINDEX
                DXForm(BFormIndex).Controls(iFocus).ListIndex = 0
                DXForm(BFormIndex).Controls(iFocus).ScrollIndex = 0
                'TOTAL
                DXForm(BFormIndex).Controls(iFocus).ScrollCount = DXForm(BFormIndex).Controls(iFocus).ScrollCount
                'EMOSTEXT
                DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                ' Home & Ctrl ----------------------------COMPLETED---------------------------
            ElseIf iK = 199 And (KeyPressed(29) Or KeyPressed(157)) = True Then
                'SCROLLCOUNT
                DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                'LISTINDEX
                DXForm(BFormIndex).Controls(iFocus).ListIndex = 0
                'EMOSTEXT
                DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                ' End & Shift & Ctrl ----------------------------COMPLETED---------------------------
            ElseIf iK = 207 And (KeyPressed(42) Or KeyPressed(54)) = True And (KeyPressed(29) Or KeyPressed(157)) = True Then
                'SCROLLCOUNT
                DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                'LISTINDEX
                DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListCount - 1 - DXForm(BFormIndex).Controls(iFocus).ScrollCount
                'SCROLLINDEX
                DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollCount
                'EMOSTEXT
                DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                ' End & Ctrl ----------------------------COMPLETED---------------------------
            ElseIf iK = 207 And (KeyPressed(29) Or KeyPressed(157)) = True Then
                'SCROLLCOUNT
                DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                'LISTINDEX
                DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListCount - DXForm(BFormIndex).Controls(iFocus).ScrollCount
                'EMOSTEKST
                DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
            End If
        End If
        Select Case iK
            Case 28 'enter
                If DXForm(BFormIndex).Controls(iFocus).PasswordChar <> "" Or DXForm(BFormIndex).Controls(iFocus).MultiLine = False Then
                    Exit Sub
                End If
            Case 1 'escape
                If DXForm(BFormIndex).Controls(iFocus).DataBase <> 0 Then
                    If DXForm(BFormIndex).Controls(iFocus).DataField <> "" Then
                        For iA = 0 To DXForm(BFormIndex).ControlsCount
                            If DXForm(BFormIndex).Controls(iFocus).DataBase = DXForm(BFormIndex).Controls(iA).DataBase Then
                                If DXForm(BFormIndex).Controls(iA).DataField <> "" And Val(DXForm(BFormIndex).Controls(iA).DataField) > -1 Then
                                    DXForm(BFormIndex).Controls(iA).Text(0) = ""
                                    DXForm(BFormIndex).Controls(iA).DataChanged = False
                                End If
                            End If
                        Next iA
                    End If
                End If
            Case 30 ' A & Select
                If (KeyPressed(29) Or KeyPressed(157)) Then ' Ctrl & Select
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    'UNDO
                    DXForm(BFormIndex).Controls(iFocus).Undo = TextCursorString
                End If
            Case 199, 200, 203, 205, 207, 208 ' Movement
            'doe niks
            Case 44 ' Undo
                If (KeyPressed(29) Or KeyPressed(157)) = False Then
                    DXForm(BFormIndex).Controls(iFocus).Undo = TextCursorString
                End If
            Case 15 ' Tab ----------------------------COMPLETED---------------------------
                'LOCKED
                If DXForm(BFormIndex).Controls(iFocus).Locked = False Then
                    If Not KeyPressed(42) And Not KeyPressed(29) And (iK = 15 Or KeyPressed(15) = True) Then
                        'UNDO
                        DXForm(BFormIndex).Controls(iFocus).Undo = TextCursorString
                        TextCursorString = Left(TextCursorString, TextCursorPosition) & KeyNames(15) & Right(TextCursorString, Len(TextCursorString) - TextCursorPosition)
                        'TEXTCURSORSTRING
                        TextCursorPosition = Len(Left(TextCursorString, TextCursorPosition) & KeyNames(15))
                        'EMOSTEXT
                        DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                        'CHANGE
                        Call xForm(BFormIndex).Controls(iFocus).Change
                    End If
                End If
                Exit Sub
            Case Else
                'zorg dat je iets ongedaan kunt maken
                DXForm(BFormIndex).Controls(iFocus).Undo = DXForm(BFormIndex).Controls(iFocus).Text(0)
        End Select
        If (KeyPressed(42) Or KeyPressed(54)) = True Then 'shift
            If TextSelectedCount = -1 And TextCursorPosition = Len(TextCursorString) And iK = 207 Then
                Exit Sub
            ElseIf TextSelectedCount = -1 Then
                TextSelectedStart = TextCursorPosition
                TextSelectedCount = 0
                ReDim TextSelected(TextSelectedCount)
                TextSelected(TextSelectedCount).Visible = True
            ElseIf TextSelectedCount = 0 And TextSelected(TextSelectedCount).Visible = False Then
                TextSelected(TextSelectedCount).Visible = True
            End If
            If (KeyPressed(29) Or KeyPressed(157)) = True Then
                Select Case iK
                    Case 203 ' Left Step Selection
                        If TextCursorPosition = 1 Then
                            TextSelected(TextSelectedCount).Text = Mid$(TextCursorString, TextCursorPosition, 1) & TextSelected(TextSelectedCount).Text
                            TextCursorPosition = TextCursorPosition - 1
                        ElseIf TextCursorPosition > 0 Then
                            If Mid$(TextCursorString, TextCursorPosition - 1, 1) = vbLf Then

                            ElseIf Mid$(TextCursorString, TextCursorPosition - 1, 1) = vbCr Then

                            End If
                            If TextSelectedStart < TextCursorPosition Then
                                For iA = TextCursorPosition - 1 To 1 Step -1
                                    If Mid(TextCursorString, iA, 1) = " " Then
                                        If TextSelected(TextSelectedCount).Text <> "" Then
                                            If TextSelectedStart >= iA Then
                                                TextSelectedCount = -1
                                                TextSelected(0).Visible = False
                                                TextCursorPosition = TextSelectedStart
                                                Exit Sub
                                            Else
                                                TextSelected(TextSelectedCount).Text = Split(Mid(TextCursorString, TextSelectedStart + 1, iA - TextSelectedStart), vbCrLf)(TextSelectedCount)
                                                TextCursorPosition = iA
                                                Exit For
                                            End If
                                        ElseIf TextSelectedCount = 0 Then
                                            TextSelectedCount = -1
                                            TextSelected(0).Visible = False
                                            Exit Sub
                                        End If
                                    ElseIf Mid(TextCursorString, iA, 2) = vbCrLf Then
                                        If iA = TextCursorPosition - 1 Then
                                            TextSelectedCount = TextSelectedCount - 1
                                            ReDim Preserve TextSelected(TextSelectedCount)
                                            TextCursorPosition = TextCursorPosition - 2
                                        Else
                                            TextSelected(TextSelectedCount).Text = ""
                                            TextCursorPosition = iA + 1
                                            Exit For
                                        End If
                                    End If
                                Next iA
                                If iA = 0 Then
                                    TextSelected(TextSelectedCount).Text = ""
                                    TextSelected(TextSelectedCount).Visible = False
                                    TextCursorPosition = 0
                                End If
                            Else
                                For iA = TextCursorPosition - 1 To 1 Step -1
                                    If Mid(TextCursorString, iA, 1) = " " Then
                                        If TextSelected(TextSelectedCount).Text <> "" Then
                                            TextCursorPosition = iA
                                            TextSelected(TextSelectedCount).Text = Mid(TextCursorString, TextCursorPosition + 1, TextSelectedStart - TextCursorPosition)
                                        ElseIf TextSelectedCount = -1 Then
                                            TextSelectedCount = TextSelectedCount + 1
                                            ReDim Preserve TextSelected(TextSelectedCount)
                                            TextSelectedStart = TextCursorPosition
                                            TextCursorPosition = iA
                                            TextSelected(TextSelectedCount).Text = Mid(TextCursorString, TextCursorPosition + 1, TextSelectedStart - TextCursorPosition)
                                        Else
                                            TextCursorPosition = iA
                                            TextSelected(TextSelectedCount).Text = Mid(TextCursorString, TextCursorPosition + 1, TextSelectedStart - TextCursorPosition)
                                        End If
                                        Exit For
                                    ElseIf Mid(TextCursorString, iA, 2) = vbCrLf Then
                                        TextCursorPosition = iA + 2
                                        TextSelected(TextSelectedCount).Text = Mid(TextCursorString, TextCursorPosition, TextSelectedStart - TextCursorPosition + 1)
                                        'TextSelectedCount = TextSelectedCount + 1
                                        'ReDim Preserve TextSelected(TextSelectedCount)
                                        Exit For
                                    End If
                                Next iA
                                If iA = 0 Then
                                    TextCursorPosition = 0
                                    TextSelected(TextSelectedCount).Text = Mid(TextCursorString, TextCursorPosition + 1, TextSelectedStart - TextCursorPosition)
                                End If
                            End If
                        End If
                    Case 205 ' Right Step Selection ----------------------------COMPLETED---------------------------
                        'SCROLL COUNT
                        DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                        If TextCursorPosition < Len(TextCursorString) Then
                            If Mid$(TextCursorString, TextCursorPosition + 1, 1) = vbLf Or Mid$(TextCursorString, TextCursorPosition + 1, 1) = vbCr Then
                                'SCROLLINDEX
                                If DXForm(BFormIndex).Controls(iFocus).ScrollIndex < DXForm(BFormIndex).Controls(iFocus).ScrollCount Then
                                    DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex + 1
                                    'LISTINDEX
                                ElseIf DXForm(BFormIndex).Controls(iFocus).ListIndex < DXForm(BFormIndex).Controls(iFocus).ListCount Then
                                    DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex + 1
                                End If
                            End If
                        End If
                        'EMOSTEXT
                        DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                        If TextCursorPosition >= TextSelectedStart Then
                            For iA = TextCursorPosition + 1 To Len(TextCursorString)
                                If Mid(TextCursorString, iA, 1) = " " Then
                                    TextSelected(TextSelectedCount).Text = TextSelected(TextSelectedCount).Text & Mid(TextCursorString, TextCursorPosition + 1, iA - TextCursorPosition)
                                    TextCursorPosition = iA
                                    Exit For
                                ElseIf Mid(TextCursorString, iA, 1) = vbCr Or Mid(TextCursorString, iA, 1) = vbLf Then
                                    If iA = (TextCursorPosition + 1) Then
                                        TextSelectedCount = TextSelectedCount + 1
                                        ReDim Preserve TextSelected(TextSelectedCount)
                                        TextSelected(TextSelectedCount).Visible = True
                                        If Mid(TextCursorString, iA, 1) = vbCr Then
                                            iA = iA + 1
                                            TextCursorPosition = TextCursorPosition + 2
                                        Else
                                            TextCursorPosition = TextCursorPosition + 1
                                        End If
                                        Exit For
                                    Else
                                        If Mid(TextCursorString, iA, 1) = vbLf Then
                                            TextCursorPosition = TextCursorPosition - 1
                                        End If
                                        TextSelected(TextSelectedCount).Text = TextSelected(TextSelectedCount).Text & Mid(TextCursorString, TextCursorPosition + 1, iA - TextCursorPosition - 1)
                                        TextCursorPosition = iA - 1
                                        Exit For
                                    End If
                                End If
                            Next iA
                            If iA > Len(TextCursorString) Then
                                TextSelected(TextSelectedCount).Text = TextSelected(TextSelectedCount).Text & Mid(TextCursorString, TextCursorPosition + 1, iA - TextCursorPosition - 1)
                                TextCursorPosition = Len(TextCursorString)
                            End If
                        Else
                            For iA = TextCursorPosition + 1 To Len(TextCursorString)
                                If Mid(TextCursorString, iA, 1) = " " Then
                                    If TextSelectedStart <= iA Then
                                        TextSelectedCount = -1
                                        TextSelected(0).Visible = False
                                        TextCursorPosition = TextSelectedStart
                                        Exit Sub
                                    Else
                                        TextSelected(TextSelectedCount).Text = Mid(TextCursorString, iA + 1, TextSelectedStart - iA)
                                        TextCursorPosition = iA + 1
                                        Exit For
                                    End If
                                ElseIf Mid(TextCursorString, iA, 2) = vbCrLf Then
                                    If iA = TextCursorPosition + 1 Then
                                        TextSelectedCount = TextSelectedCount + 1
                                        ReDim Preserve TextSelected(TextSelectedCount)
                                        TextSelected(TextSelectedCount).Visible = True
                                        TextCursorPosition = TextCursorPosition + 2
                                    Else
                                        TextSelected(TextSelectedCount).Text = TextSelected(TextSelectedCount).Text & Mid(TextCursorString, TextCursorPosition + 1, iA - TextCursorPosition - 1)
                                        TextCursorPosition = iA - 1
                                        Exit For
                                    End If
                                End If
                            Next iA
                            If iA > Len(TextCursorString) Then
                                TextSelected(TextSelectedCount).Text = TextSelected(TextSelectedCount).Text & Mid(TextCursorString, TextCursorPosition + 1, iA - TextCursorPosition - 1)
                                TextCursorPosition = Len(TextCursorString)
                            End If
                        End If
                        'CURSOR POSITION
                        If Mid$(TextCursorString, TextCursorPosition - 1, 1) = vbLf Or Mid$(TextCursorString, TextCursorPosition - 1, 1) = vbCr Then
                            For iA = 0 To TextSelectedCount
                                If iA < (TextSelectedCount - DXForm(BFormIndex).Controls(iFocus).ScrollCount) Then
                                    TextSelected(iA).Visible = False
                                Else
                                    TextSelected(iA).Visible = True
                                End If
                                If DXForm(BFormIndex).Controls(iFocus).ScrollCount < DXForm(BFormIndex).Controls(iFocus).ScrollIndex + DXForm(BFormIndex).Controls(iFocus).ListIndex Then
                                    TextSelected(iA).Top = TextSelected(iA).Top - vFontSize
                                    TextSelected(iA).Bottom = TextSelected(iA).Bottom - vFontSize
                                End If
                            Next iA
                        End If
                    Case 207 ' End & Ctrl & Shift
                        TextSelectedCount = -1
                        ReDim TextSelected(0)
                        iB = 1 + TextSelectedStart
                        sB = TextCursorString
                        For iA = iB To Len(TextCursorString)
                            If Mid(TextCursorString, iA, 2) = vbCrLf Then
                                TextSelectedCount = TextSelectedCount + 1
                                ReDim Preserve TextSelected(TextSelectedCount)
                                TextSelected(TextSelectedCount).Text = Mid(TextCursorString, iB, iA - iB)
                                If TextSelectedCount = 0 Then
                                    If CountReturns(0, Left(sB, TextCursorPosition)) = 0 Then
                                        TextSelected(TextSelectedCount).Left = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Left + 2 + TextWidth(Left(sB, TextSelectedStart))
                                    Else
                                        If TextWidth(CStr(Split(sB, vbCrLf)(CountReturns(0, Left(sB, TextCursorPosition))))) <> "" Then
                                            TextSelected(TextSelectedCount).Left = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Left + 2 + TextWidth(CStr(Split(Left(sB, TextSelectedStart), vbCrLf)(CountReturns(0, Left(sB, TextSelectedStart)))))
                                        End If
                                    End If
                                Else
                                    TextSelected(TextSelectedCount).Left = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Left + 2
                                End If
                                TextSelected(TextSelectedCount).Right = TextSelected(TextSelectedCount).Left + TextWidth(TextSelected(TextSelectedCount).Text)
                                TextSelected(TextSelectedCount).Top = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Top + (TextSelectedCount + CountReturns(0, Left(sB, TextSelectedStart))) * (vFontSize - 2) - (DXForm(BFormIndex).Controls(iFocus).ListIndex * (vFontSize - 2))
                                TextSelected(TextSelectedCount).Bottom = TextSelected(TextSelectedCount).Top + vFontSize
                                iB = iA + 2
                                TextCursorPosition = iB
                            End If
                        Next iA
                        TextSelectedCount = TextSelectedCount + 1
                        ReDim Preserve TextSelected(TextSelectedCount)
                        TextSelected(TextSelectedCount).Text = Mid(sB, iB)
                        TextSelected(TextSelectedCount).Left = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Left + 2
                        TextSelected(TextSelectedCount).Right = TextSelected(TextSelectedCount).Left + TextWidth(TextSelected(TextSelectedCount).Text)
                        TextSelected(TextSelectedCount).Top = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Top + (TextSelectedCount + CountReturns(0, Left(sB, TextSelectedStart))) * (vFontSize - 2) - (DXForm(BFormIndex).Controls(iFocus).ListIndex * (vFontSize - 2))
                        TextSelected(TextSelectedCount).Bottom = TextSelected(TextSelectedCount).Top + vFontSize
                        TextSelected(TextSelectedCount).Visible = True
                        'TEXTCURSORPOSITION
                        TextCursorPosition = Len(sB)
                        'SELECTION VISIBLE
                        For iA = 0 To TextSelectedCount
                            If iA >= DXForm(BFormIndex).Controls(iFocus).ListIndex Then
                                TextSelected(iA).Visible = True
                            Else
                                TextSelected(iA).Visible = False
                            End If
                        Next iA
                        'EINDE
                        Exit Sub
                    Case 199 ' Home & Ctrl & Shift To Select ----------------------------COMPLETED---------------------------
                        'REFRESH
                        TextCursorPosition = 0
                        TextSelectedCount = -1
                        ReDim TextSelected(0)
                        TextSelected(0).Visible = False
                        iB = 1
                        sB = TextCursorString
                        'PLACE
                        For iA = iB To TextSelectedStart
                            If Mid(TextCursorString, iA, 2) = vbCrLf Then
                                TextSelectedCount = TextSelectedCount + 1
                                ReDim Preserve TextSelected(TextSelectedCount)
                                TextSelected(TextSelectedCount).Text = Mid(TextCursorString, iB, iA - iB)
                                If TextSelectedCount = 0 Then
                                    TextSelected(TextSelectedCount).Left = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Left + 2
                                Else
                                    TextSelected(TextSelectedCount).Left = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Left + 2
                                End If
                                TextSelected(TextSelectedCount).Right = TextSelected(TextSelectedCount).Left + TextWidth(TextSelected(TextSelectedCount).Text)
                                TextSelected(TextSelectedCount).Top = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Top + TextSelectedCount * (vFontSize - 2) - (DXForm(BFormIndex).Controls(iFocus).ListIndex * (vFontSize - 2))
                                TextSelected(TextSelectedCount).Bottom = TextSelected(TextSelectedCount).Top + vFontSize
                                TextSelected(TextSelectedCount).Visible = True
                                iB = iA + 2
                            End If
                        Next iA
                        'TOTAL
                        TextSelectedCount = TextSelectedCount + 1
                        ReDim Preserve TextSelected(TextSelectedCount)
                        'POSTITION
                        TextSelected(TextSelectedCount).Text = Mid(sB, iB, TextSelectedStart - iB + 1)
                        TextSelected(TextSelectedCount).Left = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Left + 2
                        TextSelected(TextSelectedCount).Right = TextSelected(TextSelectedCount).Left + TextWidth(TextSelected(TextSelectedCount).Text)
                        TextSelected(TextSelectedCount).Top = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Top + TextSelectedCount * (vFontSize - 2) - (DXForm(BFormIndex).Controls(iFocus).ListIndex * (vFontSize - 2))
                        TextSelected(TextSelectedCount).Bottom = TextSelected(TextSelectedCount).Top + vFontSize
                        'VISIBLE
                        For iA = 0 To TextSelectedCount
                            If iA <= DXForm(BFormIndex).Controls(iFocus).ListIndex + DXForm(BFormIndex).Controls(iFocus).ScrollCount + DXForm(BFormIndex).Controls(iFocus).ScrollIndex Then
                                TextSelected(iA).Visible = True
                            Else
                                TextSelected(iA).Visible = False
                            End If
                        Next iA
                        'EINDE
                        Exit Sub
                End Select
            Else
                Select Case iK
                    Case 15 ' Tab
                    'KEYTAB
                    Case 203 ' Left Selection
                        If TextCursorPosition = 1 Then
                            TextSelected(TextSelectedCount).Text = Mid$(TextCursorString, TextCursorPosition, 1) & TextSelected(TextSelectedCount).Text
                            TextCursorPosition = 0
                        ElseIf TextCursorPosition > 0 Then
                            'SCROLL COUNT
                            DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                            If Mid$(TextCursorString, TextCursorPosition - 1, 1) = vbCr Then
                                'LISTINDEX
                                If DXForm(BFormIndex).Controls(iFocus).ListIndex > 0 Then
                                    DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex - 1
                                    'SCROLLINDEX
                                ElseIf DXForm(BFormIndex).Controls(iFocus).ScrollIndex > 0 Then
                                    DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex - 1
                                End If
                            End If
                            'EMOSTEXT
                            DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                            'TEXTCURSORPOSITION
                            If TextCursorPosition <= TextSelectedStart Then
                                If Mid$(TextCursorString, TextCursorPosition - 1, 1) = vbCr Then
                                    'TEXT SELECTION
                                    TextSelectedCount = TextSelectedCount + 1
                                    ReDim Preserve TextSelected(TextSelectedCount)
                                    TextSelected(TextSelectedCount).Visible = True
                                    TextSelected(TextSelectedCount).Text = Mid(TextCursorString, TextCursorPosition - 1, 1)
                                    'TEXTCURSORPOSITION
                                    TextCursorPosition = TextCursorPosition - 2
                                    'TEXTSELECTED START
                                Else
                                    'TEXT SELECTION
                                    TextSelected(TextSelectedCount).Text = Mid(TextCursorString, TextCursorPosition, 1) & TextSelected(TextSelectedCount).Text
                                    'TEXTCURSORPOSITION
                                    TextCursorPosition = TextCursorPosition - 1
                                End If
                            Else
                                If TextCursorPosition > 0 Then
                                    If Len(TextSelected(TextSelectedCount).Text) = 1 Then
                                        'TEXT SELECTION
                                        ReDim Preserve TextSelected(TextSelectedCount)
                                        TextSelectedCount = TextSelectedCount - 1
                                        TextSelected(0).Text = ""
                                        TextCursorPosition = TextCursorPosition - 1
                                        TextSelected(0).Visible = False
                                        Exit Sub
                                    Else
                                        'TEXT SELECTION
                                        TextSelected(TextSelectedCount).Text = Left(TextSelected(TextSelectedCount).Text, Len(TextSelected(TextSelectedCount).Text) - 1)
                                        'TEXTCURSORPOSITION
                                        TextCursorPosition = TextCursorPosition - 1
                                    End If
                                End If
                            End If
                            'CURSOR POSITION
                            If TextCursorPosition - 1 > 0 Then
                                If Mid$(TextCursorString, TextCursorPosition + 1, 1) = vbLf Or Mid$(TextCursorString, TextCursorPosition + 1, 1) = vbCr Then
                                    For iA = TextSelectedCount To 0 Step -1
                                        If iA < (TextSelectedCount - DXForm(BFormIndex).Controls(iFocus).ScrollCount) Then
                                            TextSelected(iA).Visible = False
                                        Else
                                            TextSelected(iA).Visible = True
                                        End If
                                        If DXForm(BFormIndex).Controls(iFocus).ScrollCount > DXForm(BFormIndex).Controls(iFocus).ScrollIndex + DXForm(BFormIndex).Controls(iFocus).ListIndex Then
                                            TextSelected(iA).Top = TextSelected(iA).Top - vFontSize
                                            TextSelected(iA).Bottom = TextSelected(iA).Bottom - vFontSize
                                        End If
                                    Next iA
                                End If
                            End If
                        End If
                    Case 205 ' Right Selection ----------------------------COMPLETED---------------------------
                        'SCROLL COUNT
                        DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                        If TextCursorPosition < Len(TextCursorString) Then
                            If Mid$(TextCursorString, TextCursorPosition + 1, 1) = vbLf Then
                                'SCROLLINDEX
                                If DXForm(BFormIndex).Controls(iFocus).ScrollIndex < DXForm(BFormIndex).Controls(iFocus).ScrollCount Then
                                    DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex + 1
                                    'LISTINDEX
                                ElseIf DXForm(BFormIndex).Controls(iFocus).ListIndex < DXForm(BFormIndex).Controls(iFocus).ListCount Then
                                    DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex + 1
                                End If
                            End If
                        End If
                        'EMOSTEXT
                        DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                        'TEXTCURSORPOSITION
                        If TextCursorPosition < Len(TextCursorString) Then
                            If TextCursorPosition >= TextSelectedStart Then
                                If Mid$(TextCursorString, TextCursorPosition + 1, 1) = vbLf Then
                                    'TEXT SELECTION
                                    TextSelectedCount = TextSelectedCount + 1
                                    ReDim Preserve TextSelected(TextSelectedCount)
                                    TextSelected(TextSelectedCount).Visible = True
                                    TextSelected(TextSelectedCount).Text = Mid(TextCursorString, TextCursorPosition + 2, 1)
                                    'TEXTCURSORPOSITION
                                    TextCursorPosition = TextCursorPosition + 2
                                    'TEXTSELECTED START
                                Else
                                    'TEXT SELECTION
                                    TextSelected(TextSelectedCount).Text = TextSelected(TextSelectedCount).Text & Mid(TextCursorString, TextCursorPosition + 1, 1)
                                    'TEXTCURSORPOSITION
                                    TextCursorPosition = TextCursorPosition + 1
                                End If
                            Else
                                If TextCursorPosition >= 0 Then
                                    If TextSelected(TextSelectedCount).Text = "" Then
                                        'TEXT SELECTION
                                        TextSelectedCount = TextSelectedCount - 1
                                        ReDim Preserve TextSelected(TextSelectedCount)
                                        TextCursorPosition = TextCursorPosition - 2
                                        If TextSelected(TextSelectedCount).Text <> "" Then
                                            TextSelected(TextSelectedCount).Text = Left(TextSelected(TextSelectedCount).Text, Len(TextSelected(TextSelectedCount).Text) - 1)
                                            TextCursorPosition = TextCursorPosition - 1
                                        ElseIf TextSelectedCount = 0 Then
                                            TextSelectedCount = -1
                                            TextSelected(0).Visible = False
                                            Exit Sub
                                        End If
                                    Else
                                        'TEXT SELECTION
                                        TextSelected(TextSelectedCount).Text = Right(TextSelected(TextSelectedCount).Text, Len(TextSelected(TextSelectedCount).Text) - 1)
                                        'TEXTCURSORPOSITION
                                        TextCursorPosition = TextCursorPosition + 1
                                    End If
                                End If
                            End If
                        ElseIf TextCursorPosition < Len(TextCursorString) Then
                            'TEXT SELECTION
                            TextSelected(TextSelectedCount).Text = TextSelected(TextSelectedCount).Text & Mid(TextCursorString, TextCursorPosition + 1, 1)
                            'TEXTCURSORPOSITION
                            TextCursorPosition = TextCursorPosition + 1
                        End If
                        'CURSOR POSITION
                        If TextCursorPosition - 1 > 0 Then
                            If Mid$(TextCursorString, TextCursorPosition - 1, 1) = vbLf Or Mid$(TextCursorString, TextCursorPosition - 1, 1) = vbCr Then
                                For iA = 0 To TextSelectedCount
                                    If iA < (TextSelectedCount - DXForm(BFormIndex).Controls(iFocus).ScrollCount) Then
                                        TextSelected(iA).Visible = False
                                    Else
                                        TextSelected(iA).Visible = True
                                    End If
                                    If DXForm(BFormIndex).Controls(iFocus).ScrollCount < DXForm(BFormIndex).Controls(iFocus).ScrollIndex + DXForm(BFormIndex).Controls(iFocus).ListIndex Then
                                        TextSelected(iA).Top = TextSelected(iA).Top - vFontSize
                                        TextSelected(iA).Bottom = TextSelected(iA).Bottom - vFontSize
                                    End If
                                Next iA
                            End If
                        End If
                    Case 199 ' Home Selection ----------------------------COMPLETED---------------------------
                        If TextSelectedStart >= TextCursorPosition Then
                            For iA = TextSelectedStart - 1 To 1 Step -1
                                If Mid(TextCursorString, iA, 2) = vbCrLf Then
                                    If (TextSelectedStart - iA - 1) = 0 Then
                                        Exit Sub
                                    Else
                                        TextSelected(TextSelectedCount).Text = Mid(TextCursorString, iA + 2, TextSelectedStart - iA - 1)
                                        TextCursorPosition = iA + 1
                                    End If
                                    Exit For
                                End If
                            Next iA
                            If iA = 0 Then
                                If (iA - TextCursorPosition - 1) > 0 Then
                                    TextSelected(TextSelectedCount).Text = TextSelected(TextSelectedCount).Text & Mid(TextCursorString, TextCursorPosition + 1, iA - TextCursorPosition - 1)
                                    TextCursorPosition = Len(TextCursorString)
                                Else
                                    TextSelected(TextSelectedCount).Text = Left(TextCursorString, TextCursorPosition)
                                    TextCursorPosition = iA
                                End If
                            End If
                        Else
                            If TextSelectedStart = TextCursorPosition - Len(TextSelected(TextSelectedCount).Text) Then
                                If TextCursorPosition = Len(TextSelected(TextSelectedCount).Text) Then
                                    TextSelected(TextSelectedCount).Text = Left(TextCursorString, TextCursorPosition)
                                    TextCursorPosition = 0
                                Else
                                    TextSelected(TextSelectedCount).Text = Mid$(TextCursorString, TextSelectedStart + 1, Len(TextSelected(TextSelectedCount).Text))
                                    TextCursorPosition = TextSelectedStart
                                End If
                            Else
                                TextCursorPosition = TextCursorPosition - Len(TextSelected(TextSelectedCount).Text)
                                TextSelected(TextSelectedCount).Text = ""
                            End If
                        End If
                        If TextSelectedStart = 0 Then
                            Exit Sub
                        End If
                    Case 207 ' End Selection ----------------------------COMPLETED---------------------------
                        For iA = TextCursorPosition + 1 To Len(TextCursorString)
                            If Mid(TextCursorString, iA, 2) = vbCrLf Then
                                If CountReturns(0, Mid(TextCursorString, TextSelectedStart + 1, iA - TextSelectedStart)) > TextSelectedCount Then
                                    Exit Sub
                                Else
                                    If iA - TextSelectedStart = 0 Then
                                        Exit Sub
                                    ElseIf Split(Mid(TextCursorString, TextSelectedStart + 1, iA - TextSelectedStart), vbCrLf)(TextSelectedCount) = vbCr Then
                                        TextSelected(TextSelectedCount).Text = Mid$(TextCursorString, TextCursorPosition + 1, TextSelectedStart - TextCursorPosition)
                                        TextCursorPosition = TextSelectedStart
                                        Exit For
                                    Else
                                        TextCursorPosition = iA
                                        TextSelected(TextSelectedCount).Text = Split(Mid(TextCursorString, TextSelectedStart + 1, iA - TextSelectedStart), vbCrLf)(TextSelectedCount)
                                    End If
                                    Exit For
                                End If
                            End If
                        Next iA
                        If iA > Len(TextCursorString) Then
                            If TextSelectedStart > TextCursorPosition Then
                                TextSelected(TextSelectedCount).Text = Mid(TextCursorString, TextSelectedStart + 1, iA - TextSelectedStart)
                            ElseIf TextCursorPosition = Len(TextCursorString) Then
                                'do nothing
                            Else
                                TextSelected(TextSelectedCount).Text = TextSelected(TextSelectedCount).Text & Mid(TextCursorString, TextCursorPosition + 1, iA - TextCursorPosition - 1)
                            End If
                            TextCursorPosition = Len(TextCursorString)
                        End If
                    Case 211 ' Clear All
                        'dit doet dus niks
                        If DXForm(BFormIndex).Controls(iFocus).Locked = False Then
                            DXForm(BFormIndex).Controls(iFocus).Text(iFocus2) = ""
                            DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(iFocus2), DXForm(BFormIndex).Controls(iFocus))
                            TextCursorPosition = 0
                        End If
                    Case Else ' Type Over Selection - With Shift
                        If DXForm(BFormIndex).Controls(iFocus).Locked = False Then
                            If KeyNames(iK) <> "" Then
                                If TextSelected(0).Visible = True Then
                                    For iA = 0 To TextSelectedCount - 1
                                        sA = sA & TextSelected(iA).Text & vbCrLf
                                    Next iA
                                    sA = sA & TextSelected(TextSelectedCount).Text
                                    If sA = "" Then
                                        ' Do Not A Thing
                                    ElseIf Len(TextCursorString) = Len(sA) Then
                                        TextSelectedCount = -1
                                        ReDim TextSelected(0)
                                        TextSelected(0).Visible = False
                                        xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = ""
                                        TextCursorPosition = 0
                                    Else
                                        For iA = 1 To Len(TextCursorString)
                                            If Mid(TextCursorString, iA, Len(sA)) = sA Then
                                                xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = Left(TextCursorString, iA - 1) & Mid(TextCursorString, iA + Len(sA))
                                                TextSelectedCount = -1
                                                ReDim TextSelected(0)
                                                TextSelected(0).Visible = False
                                                TextCursorPosition = iA - 1
                                                Exit For
                                            End If
                                        Next iA
                                    End If
                                End If
                                If iK + 300 > 355 Then
                                    Exit Sub
                                End If
                                If Len(TextCursorString) = TextCursorPosition Then
                                    xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = TextCursorString & KeyNames(iK + 300)
                                    TextCursorPosition = TextCursorPosition + Len(KeyNames(iK + 300))
                                Else
                                    xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = Left(TextCursorString, TextCursorPosition) & KeyNames(iK + 300) & Right(TextCursorString, Len(TextCursorString) - TextCursorPosition)
                                    TextCursorPosition = TextCursorPosition + Len(KeyNames(iK + 300))
                                End If
                            End If
                            Call xForm(BFormIndex).Text1(DXForm(BFormIndex).Controls(iFocus).Index).Change
                        End If
                        Exit Sub
                End Select
            End If
            If TextCursorString <> "" Then
                If TextCursorPosition > TextSelectedStart Then
                    If Left(TextCursorString, TextCursorPosition) = "" Then
                        iB = 0
                    ElseIf TextSelectedCount = -1 Then
                        iB = 0
                    ElseIf TextCursorPosition < Len(TextSelected(TextSelectedCount).Text) Then
                        iB = 0
                    ElseIf Left(TextCursorString, TextCursorPosition - Len(TextSelected(TextSelectedCount).Text)) = "" Then
                        iB = 0
                    ElseIf CountReturns(0, Left(TextCursorString, TextCursorPosition)) > UBound(Split(Left(TextCursorString, TextCursorPosition - Len(TextSelected(TextSelectedCount).Text)), vbCrLf)) Then
                        iB = 0
                    ElseIf TextCursorPosition > TextSelectedStart Then
                        iB = TextWidth(CStr(Split(Left(TextCursorString, TextCursorPosition - Len(TextSelected(TextSelectedCount).Text)), vbCrLf)(CountReturns(0, Left(TextCursorString, TextCursorPosition)))))
                    Else
                        iB = TextWidth(CStr(Split(Left(TextCursorString, TextSelectedStart - Len(TextSelected(TextSelectedCount).Text)), vbCrLf)(CountReturns(0, Left(TextCursorString, TextSelectedStart)))))
                    End If
                Else
                    If Left(TextCursorString, TextSelectedStart) = "" Then
                        iB = 0
                    ElseIf TextSelectedCount = -1 Then
                        iB = 0
                    ElseIf TextSelectedStart < Len(TextSelected(TextSelectedCount).Text) Then
                        iB = 0
                    ElseIf Left(TextCursorString, TextSelectedStart - Len(TextSelected(TextSelectedCount).Text)) = "" Then
                        iB = 0
                    ElseIf CountReturns(0, Left(TextCursorString, TextSelectedStart)) > UBound(Split(Left(TextCursorString, TextSelectedStart - Len(TextSelected(TextSelectedCount).Text)), vbCrLf)) Then
                        iB = 0
                    Else
                        iB = TextWidth(CStr(Split(Left(TextCursorString, TextSelectedStart - Len(TextSelected(TextSelectedCount).Text)), vbCrLf)(CountReturns(0, Left(TextCursorString, TextSelectedStart)))))
                    End If
                End If
            Else
                iB = 0
            End If
            TextCursor.Left = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Left + iB
            If CountReturns(0, Left(TextCursorString, TextCursorPosition)) - 1 >= DXForm(BFormIndex).Controls(iFocus).ScrollCount Then
                TextCursor.Top = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Top + DXForm(BFormIndex).Controls(iFocus).ScrollIndex * (vFontSize - 2)
            Else
                TextCursor.Top = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Top + CountReturns(0, Left(TextCursorString, TextCursorPosition)) * (vFontSize - 2)
            End If
            TextSelected(TextSelectedCount).Top = TextCursor.Top
            TextSelected(TextSelectedCount).Bottom = TextSelected(TextSelectedCount).Top + vFontSize
            iA = TextWidth(TextSelected(TextSelectedCount).Text)
            'If TextSelectedCount > 0 Then
            'For iB = 0 To TextSelectedCount - 1
            '                iA = iA - TextWidth(TextSelected(iB).Text & vbCrLf)
            'Next iB
            'End If
            If DXForm(BFormIndex).Controls(iFocus).Name = "Data1" And DXForm(BFormIndex).Controls(iFocus).DataField = "2" Then
                TextSelected(TextSelectedCount).Left = TextCursor.Left
            Else
                TextSelected(TextSelectedCount).Left = TextCursor.Left + 2
            End If
            TextSelected(TextSelectedCount).Right = TextSelected(TextSelectedCount).Left + iA
            If iK <> 30 Then
                Exit Sub
            End If
        End If
        If TextSelected(0).Visible = True Then
            If (KeyPressed(29) Or KeyPressed(157)) = True Then
                Select Case iK
                    Case 45 ' Cut
                        For iA = 0 To TextSelectedCount - 1
                            sA = sA & TextSelected(iA).Text & vbCrLf
                        Next iA
                        sA = sA & TextSelected(TextSelectedCount).Text
                        Call Clipboard.SetText(sA)
                        If DXForm(BFormIndex).Controls(iFocus).Locked = False Then
                            If Len(TextCursorString) = Len(sA) Then
                                TextSelectedCount = -1
                                ReDim TextSelected(0)
                                TextSelected(0).Visible = False
                                xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = ""
                                TextCursorPosition = 0
                            Else
                                For iA = 1 To Len(TextCursorString)
                                    If Mid(TextCursorString, iA, Len(sA)) = sA Then
                                        xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = Left(TextCursorString, iA - 1) & Right(TextCursorString, Len(TextCursorString) - Len(sA) - iA + 1)
                                        TextSelectedCount = -1
                                        ReDim TextSelected(0)
                                        TextSelected(0).Visible = False
                                        TextCursorPosition = iA - 1
                                        Exit Sub
                                    End If
                                Next iA
                            End If
                        End If
                        Exit Sub
                    Case 46 ' Copy
                        For iA = 0 To TextSelectedCount - 1
                            sA = sA & TextSelected(iA).Text & vbCrLf
                        Next iA
                        sA = sA & TextSelected(TextSelectedCount).Text
                        Call Clipboard.SetText(sA)
                        Exit Sub
                End Select
            Else
                Select Case iK
                    Case 15
                        'KEYTAB
                        'EINDE
                        Exit Sub
                    Case 1, 199, 200, 203, 205, 207, 208
                    ' Do Not A Thing
                    Case 14, 211 ' Remove
                        'LOCKED
                        If DXForm(BFormIndex).Controls(iFocus).Locked = False Then
                            'TEXTSELECTED
                            If TextSelected(TextSelectedCount).Left > TextCursor.Left Then
                                TextCursorString = Left(TextCursorString, TextCursorPosition) & Mid(TextCursorString, TextCursorPosition + Len(TextSelected(TextSelectedCount).Text) + 1)
                            Else
                                TextCursorString = Left(TextCursorString, TextCursorPosition - Len(TextSelected(TextSelectedCount).Text)) & Mid(TextCursorString, TextCursorPosition + 1)
                                TextCursorPosition = TextCursorPosition - Len(TextSelected(TextSelectedCount).Text)
                            End If
                            'SELECTED OPHEFFEN
                            TextSelected(TextSelectedCount).Text = ""
                            TextSelected(TextSelectedCount).Visible = False
                            TextSelected(TextSelectedCount).Left = 0
                            TextSelected(TextSelectedCount).Right = 0
                            'EMOSTEKST
                            DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(iFocus2), DXForm(BFormIndex).Controls(iFocus))
                        End If
                        'EINDE
                        Exit Sub
                    Case Else ' Type Over Selection - No Shift
                        If DXForm(BFormIndex).Controls(iFocus).Locked = False Then
                            For iA = 0 To TextSelectedCount - 1
                                sA = sA & TextSelected(iA).Text & vbCrLf
                            Next iA
                            sA = sA & TextSelected(TextSelectedCount).Text
                            If Len(TextCursorString) = Len(sA) Then
                                TextSelectedCount = -1
                                ReDim TextSelected(0)
                                TextSelected(0).Visible = False
                                xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = ""
                                TextCursorPosition = 0
                            ElseIf Len(sA) = 0 Then
                                TextSelectedCount = -1
                                ReDim TextSelected(0)
                                TextSelected(0).Visible = False
                            Else
                                For iA = 1 To Len(TextCursorString)
                                    If Mid(TextCursorString, iA, Len(sA)) = sA Then
                                        xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = Left(TextCursorString, iA - 1) & Mid(TextCursorString, iA + Len(sA))
                                        TextSelectedCount = -1
                                        ReDim TextSelected(0)
                                        TextSelected(0).Visible = False
                                        TextCursorPosition = iA - 1
                                        Exit For
                                    End If
                                Next iA
                            End If
                            If Len(TextCursorString) = TextCursorPosition Then
                                xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = TextCursorString & KeyNames(iK)
                                TextCursorPosition = TextCursorPosition + Len(KeyNames(iK))
                            Else
                                xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = Left(TextCursorString, TextCursorPosition) & KeyNames(iK) & Right(TextCursorString, Len(TextCursorString) - TextCursorPosition)
                                TextCursorPosition = TextCursorPosition + Len(KeyNames(iK))
                            End If
                            Call xForm(BFormIndex).Text1(DXForm(BFormIndex).Controls(iFocus).Index).Change
                        End If
                        Exit Sub
                End Select
            End If
        End If
        TextSelectedStart = 0
        TextSelectedCount = -1
        ReDim TextSelected(0)
        TextSelected(0).Visible = False
        If (KeyPressed(29) Or KeyPressed(157)) Then ' Ctrl
            Select Case iK
                Case 44 ' Undo
                    If DXForm(BFormIndex).Controls(iFocus).Locked = False Then
                        sA = TextCursorString
                        xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = DXForm(BFormIndex).Controls(iFocus).Undo
                        DXForm(BFormIndex).Controls(iFocus).Undo = sA
                    End If
                Case 15 ' Tab
                    'KEYTAB
                    Exit Sub
                Case 30 ' Ctrl & A = Select All
                    iB = 1
                    For iA = 1 To Len(TextCursorString)
                        If Mid(TextCursorString, iA, 2) = vbCrLf Then
                            TextSelectedCount = TextSelectedCount + 1
                            ReDim Preserve TextSelected(TextSelectedCount)
                            TextSelected(TextSelectedCount).Text = Mid(TextCursorString, iB, iA - iB)
                            TextSelected(TextSelectedCount).Left = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Left + 2
                            TextSelected(TextSelectedCount).Right = TextSelected(TextSelectedCount).Left + TextWidth(TextSelected(TextSelectedCount).Text)
                            TextSelected(TextSelectedCount).Top = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Top + (TextSelectedCount + CountReturns(0, Left(sB, TextSelectedStart))) * (vFontSize - 2) - (DXForm(BFormIndex).Controls(iFocus).ListIndex * (vFontSize - 2))
                            TextSelected(TextSelectedCount).Bottom = TextSelected(TextSelectedCount).Top + vFontSize
                            If TextSelectedCount < DXForm(BFormIndex).Controls(iFocus).ListIndex Then
                                TextSelected(TextSelectedCount).Visible = False
                            ElseIf TextSelectedCount <= (DXForm(BFormIndex).Controls(iFocus).ListIndex + DXForm(BFormIndex).Controls(iFocus).ScrollIndex + DXForm(BFormIndex).Controls(iFocus).ScrollCount) Then
                                TextSelected(TextSelectedCount).Visible = True
                            Else
                                TextSelected(TextSelectedCount).Visible = False
                            End If
                            iB = iA + 2
                        End If
                    Next iA
                    TextSelectedCount = TextSelectedCount + 1
                    ReDim Preserve TextSelected(TextSelectedCount)
                    TextSelected(TextSelectedCount).Text = Mid(TextCursorString, iB)
                    TextSelected(TextSelectedCount).Left = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Left + 2
                    TextSelected(TextSelectedCount).Right = TextSelected(TextSelectedCount).Left + TextWidth(TextSelected(TextSelectedCount).Text)
                    TextSelected(TextSelectedCount).Top = DXForm(BFormIndex).Controls(iFocus).Position(iFocus2).Top + (TextSelectedCount + CountReturns(0, Left(sB, TextSelectedStart))) * (vFontSize - 2) - (DXForm(BFormIndex).Controls(iFocus).ListIndex * (vFontSize - 2))
                    TextSelected(TextSelectedCount).Bottom = TextSelected(TextSelectedCount).Top + vFontSize
                    If TextSelectedCount <= (DXForm(BFormIndex).Controls(iFocus).ListIndex + DXForm(BFormIndex).Controls(iFocus).ScrollIndex + DXForm(BFormIndex).Controls(iFocus).ScrollCount) Then
                        TextSelected(TextSelectedCount).Visible = True
                    Else
                        TextSelected(TextSelectedCount).Visible = False
                    End If
                    'EINDE
                    Exit Sub
                Case 47 ' Paste
                    If DXForm(BFormIndex).Controls(iFocus).Locked = False Then
                        xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = Left(TextCursorString, TextCursorPosition) & Clipboard.GetText & Right(TextCursorString, Len(TextCursorString) - TextCursorPosition)
                        TextCursorPosition = TextCursorPosition + Len(Clipboard.GetText)
                    End If
                Case 203 ' Step Left ----------------------------COMPLETED---------------------------
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    'TEXTCURSORPOSITION
                    For iA = (TextCursorPosition - 1) To 1 Step -1
                        If Mid(TextCursorString, iA, 1) = " " Then
                            TextCursorPosition = iA
                            'EMOSTEXT
                            DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                            Exit Sub
                        ElseIf Mid(TextCursorString, iA, 1) = Right(vbCrLf, 1) Then
                            TextCursorPosition = iA - 1
                            'SCROLLINDEX
                            If DXForm(BFormIndex).Controls(iFocus).ScrollIndex <> 0 And DXForm(BFormIndex).Controls(iFocus).ScrollIndex <= DXForm(BFormIndex).Controls(iFocus).ScrollCount Then
                                DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex - 1
                                'LISTINDEX
                            ElseIf DXForm(BFormIndex).Controls(iFocus).ListIndex <> 0 And DXForm(BFormIndex).Controls(iFocus).ListIndex <= DXForm(BFormIndex).Controls(iFocus).ListCount - DXForm(BFormIndex).Controls(iFocus).ScrollCount - 1 Then
                                DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex - 1
                            End If
                            'EMOSTEXT
                            DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                            Exit Sub
                        End If
                    Next iA
                    TextCursorPosition = 0
                    'EMOSTEXT
                    DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    'EINDE
                    Exit Sub
                Case 205 ' Step Right ----------------------------COMPLETED---------------------------
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    'TEXTCURSORPOSITION
                    For iA = TextCursorPosition + 1 To Len(TextCursorString)
                        If Mid(TextCursorString, iA, 1) = " " Then
                            TextCursorPosition = iA
                            'EMOSTEXT
                            DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                            Exit Sub
                        ElseIf Mid(TextCursorString, iA, 2) = vbCrLf Then
                            TextCursorPosition = iA + 1
                            'SCROLLINDEX
                            If DXForm(BFormIndex).Controls(iFocus).ScrollIndex < DXForm(BFormIndex).Controls(iFocus).ScrollCount Then
                                DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex + 1
                                'LISTINDEX
                            ElseIf DXForm(BFormIndex).Controls(iFocus).ListIndex < DXForm(BFormIndex).Controls(iFocus).ListCount - DXForm(BFormIndex).Controls(iFocus).ScrollCount - 1 Then
                                DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex + 1
                            End If
                            'EMOSTEXT
                            DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                            Exit Sub
                        End If
                    Next iA
                    TextCursorPosition = Len(TextCursorString)
                    'EMOSTEXT
                    DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    'EINDE
                    Exit Sub
                Case 199 ' Step Home ----------------------------COMPLETED---------------------------
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    'SCROLLINDEX
                    DXForm(BFormIndex).Controls(iFocus).ScrollIndex = 0
                    'LISTINDEX
                    DXForm(BFormIndex).Controls(iFocus).ListIndex = 0
                    'EMOSTEXT
                    DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    'TEXTCURSORPOSITION
                    TextCursorPosition = 0
                    'EINDE
                    Exit Sub
                Case 207 ' Step End ----------------------------COMPLETED---------------------------
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    'SCROLLINDEX
                    DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollCount
                    'LISTINDEX
                    DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListCount - DXForm(BFormIndex).Controls(iFocus).ScrollCount - 1
                    'EMOSTEXT
                    DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    'TEXTCURSORPOSITION
                    TextCursorPosition = Len(TextCursorString)
                    'EINDE
                    Exit Sub
                Case 211 ' Clear All
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    'TEXTCURSORPOSITION
                    iC = TextCursorPosition
                    For iA = TextCursorPosition + 1 To Len(TextCursorString)
                        If Mid(TextCursorString, iA, 1) = " " Then
                            TextCursorPosition = iA
                            'EMOSTEXT
                            DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                            Exit Sub
                        ElseIf Mid(TextCursorString, iA, 2) = vbCrLf Then
                            TextCursorPosition = iA + 1
                            'SCROLLINDEX
                            If DXForm(BFormIndex).Controls(iFocus).ScrollIndex < DXForm(BFormIndex).Controls(iFocus).ScrollCount Then
                                DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex + 1
                                'LISTINDEX
                            ElseIf DXForm(BFormIndex).Controls(iFocus).ListIndex < DXForm(BFormIndex).Controls(iFocus).ListCount - DXForm(BFormIndex).Controls(iFocus).ScrollCount - 1 Then
                                DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex + 1
                            End If
                            'EMOSTEXT
                            DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                            Exit Sub
                        End If
                    Next iA
                    TextCursorPosition = Len(TextCursorString)
                    'EMOSTEXT
                    DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    'EINDE
                    Exit Sub
            End Select
        Else
            Select Case iK
                Case 15 ' Tab ----------------------------COMPLETED---------------------------
                ' Do Not A Thing
                Case 14 ' Back Space ----------------------------COMPLETED---------------------------
                    'LOCKED
                    If DXForm(BFormIndex).Controls(iFocus).Locked = False Then
                        'TEXTCURSORPOSTION
                        If Len(TextCursorString) = 0 Or TextCursorPosition = 0 Then
                            Exit Sub
                        ElseIf Len(TextCursorString) = 1 Or TextCursorPosition = 1 Then
                            iA = 1
                        ElseIf Mid(TextCursorString, TextCursorPosition - 1, 2) = vbCrLf Then
                            iA = 2
                        Else
                            iA = 1
                        End If
                        If Len(TextCursorString) = TextCursorPosition Then
                            TextCursorString = Left(TextCursorString, Len(TextCursorString) - iA)
                            TextCursorPosition = TextCursorPosition - iA
                        ElseIf Len(TextCursorString) < TextCursorPosition Then
                            TextCursorPosition = Len(TextCursorString)
                        ElseIf TextCursorPosition > 0 Then
                            TextCursorString = Left(TextCursorString, TextCursorPosition - iA) & Right(TextCursorString, Len(TextCursorString) - TextCursorPosition)
                            TextCursorPosition = TextCursorPosition - iA
                        End If
                        'EMOSTEXT
                        DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(TextCursorString, DXForm(BFormIndex).Controls(iFocus))
                        'TEXT CHANGE
                        Call xForm(BFormIndex).Text1(DXForm(BFormIndex).Controls(iFocus).Index).Change
                        'EXIT
                        Exit Sub
                    End If
                Case 199 ' Home ----------------------------COMPLETED---------------------------
                    'RESET TEXT SELECTION START
                    TextSelectedStart = 0
                    'TEXTCURSORPOSITION
                    For iA = TextCursorPosition - 1 To 1 Step -1
                        If Mid(TextCursorString, iA, 1) = Left(vbCrLf, 1) Or iA = 1 Then
                            If iA = 1 Then
                                iA = -1
                            End If
                            TextCursorPosition = iA + 1
                            Exit For
                        End If
                    Next iA
                    If TextCursorPosition = 1 Then
                        TextCursorPosition = 0
                    End If
                    Exit Sub
                Case 207 ' End ----------------------------COMPLETED---------------------------
                    'RESET TEXT SELECTION START
                    TextSelectedStart = 0
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    'EMOSTEXT
                    DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    'TEXTCURSORPOSITION
                    For iA = TextCursorPosition To Len(TextCursorString)
                        If iA = 0 Then
                            iA = 1
                        End If
                        If Mid(TextCursorString, iA, 1) = vbCr Then
                            TextCursorPosition = iA - 1
                            Exit Sub
                        End If
                    Next iA
                    TextCursorPosition = Len(TextCursorString)
                    'EINDE
                    Exit Sub
                Case 203 ' Left ----------------------------COMPLETED---------------------------
                    'RESET TEXT SELECTION START
                    TextSelectedStart = 0
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    If TextCursorPosition > 1 Then
                        If Mid$(TextCursorString, TextCursorPosition - 1, 1) = vbCr Then
                            'LISTINDEX
                            If DXForm(BFormIndex).Controls(iFocus).ListIndex > 0 Then
                                DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex - 1
                                'SCROLLINDEX
                            ElseIf DXForm(BFormIndex).Controls(iFocus).ScrollIndex > 0 Then
                                DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex - 1
                            End If
                        End If
                    End If
                    'EMOSTEXT
                    DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    'TEXTCURSORPOSITION
                    If TextCursorPosition > 1 Then
                        If Mid(TextCursorString, TextCursorPosition - 1, 2) = vbCrLf Then
                            TextCursorPosition = TextCursorPosition - 2
                        Else
                            TextCursorPosition = TextCursorPosition - 1
                        End If
                    ElseIf TextCursorPosition > 0 Then
                        TextCursorPosition = TextCursorPosition - 1
                    End If
                    'EINDE
                    Exit Sub
                Case 205 ' Right ----------------------------COMPLETED---------------------------
                    'RESET TEXT SELECTION START
                    TextSelectedStart = 0
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    If TextCursorPosition < Len(TextCursorString) Then
                        If Mid$(TextCursorString, TextCursorPosition + 2, 1) = vbLf Then
                            'SCROLLINDEX
                            If DXForm(BFormIndex).Controls(iFocus).ScrollIndex < DXForm(BFormIndex).Controls(iFocus).ScrollCount Then
                                DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex + 1
                                'LISTINDEX
                            ElseIf DXForm(BFormIndex).Controls(iFocus).ListIndex < DXForm(BFormIndex).Controls(iFocus).ListCount Then
                                DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex + 1
                            End If
                        End If
                    End If
                    'EMOSTEXT
                    DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    'TEXTCURSORPOSITION
                    If TextCursorPosition < (Len(TextCursorString) - 1) Then
                        If Mid(TextCursorString, TextCursorPosition + 1, 2) = vbCrLf Then
                            TextCursorPosition = TextCursorPosition + 2
                        Else
                            TextCursorPosition = TextCursorPosition + 1
                        End If
                    ElseIf TextCursorPosition < Len(TextCursorString) Then
                        TextCursorPosition = TextCursorPosition + 1
                    End If
                    'EINDE
                    Exit Sub
                Case 200 ' Up  ----------------------------COMPLETED---------------------------
                    'RESET TEXT SELECTION START
                    TextSelectedStart = 0
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    'LISTINDEX
                    If DXForm(BFormIndex).Controls(iFocus).ListIndex > 0 Then
                        DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex - 1
                    ElseIf DXForm(BFormIndex).Controls(iFocus).ScrollIndex > 0 Then
                        'SCROLLINDEX
                        DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex - 1
                    End If
                    'EMOSTEXT
                    DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    'TEXTCURSORPOSITION
                    For iA = TextCursorPosition To 1 Step -1
                        If Mid(TextCursorString, iA, 2) = vbCrLf And iB = 0 Then
                            iB = TextCursorPosition - iA
                        ElseIf Mid(TextCursorString, iA, 2) = vbCrLf And iC = 0 Then
                            iC = TextCursorPosition - iA - iB
                            Exit For
                        End If
                    Next iA
                    If iB = 0 Then
                        TextCursorPosition = 0
                    ElseIf iC = 0 Then
                        iA = TextCursorPosition - iB
                        If iA > iB Then
                            TextCursorPosition = iB - 1
                        Else
                            TextCursorPosition = TextCursorPosition - iB - 1
                        End If
                    ElseIf iC >= iB Then
                        TextCursorPosition = TextCursorPosition - iC
                    ElseIf iC < iB Then
                        TextCursorPosition = TextCursorPosition - iB - 1
                    End If
                    'EINDE
                    Exit Sub
                Case 208 ' Down ----------------------------COMPLETED---------------------------
                    'RESET TEXT SELECTION START
                    TextSelectedStart = 0
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    'SCROLLINDEX
                    If DXForm(BFormIndex).Controls(iFocus).ScrollIndex < DXForm(BFormIndex).Controls(iFocus).ScrollCount Then
                        DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex + 1
                        'LISTINDEX
                    ElseIf DXForm(BFormIndex).Controls(iFocus).ListIndex < DXForm(BFormIndex).Controls(iFocus).ListCount - DXForm(BFormIndex).Controls(iFocus).ScrollCount - 1 Then
                        DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex + 1
                    End If
                    'EMOSTEXT
                    DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    'TEXTCURSORPOSITION
                    If TextCursorPosition = 0 Then
                        TextCursorPosition = 1
                    End If
                    For iA = TextCursorPosition To Len(TextCursorString) - 1
                        If Mid(TextCursorString, iA, 2) = vbCrLf Then
                            TextCursorPosition = iA + 1
                            Exit For
                        End If
                    Next iA
                    If iA = Len(TextCursorString) Then
                        TextCursorPosition = Len(TextCursorString)
                    End If
                    'EINDE
                    Exit Sub
                Case 201 ' Page Up ----------------------------COMPLETED---------------------------
                    'RESET TEXT SELECTION START
                    TextSelectedStart = 0
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    'SCROLLINDEX
                    If DXForm(BFormIndex).Controls(iFocus).ScrollIndex > 0 Then
                        DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex - DXForm(BFormIndex).Controls(iFocus).ScrollCount
                        If DXForm(BFormIndex).Controls(iFocus).ScrollIndex < 0 Then
                            DXForm(BFormIndex).Controls(iFocus).ScrollIndex = 0
                        End If
                        'LISTINDEX
                    ElseIf DXForm(BFormIndex).Controls(iFocus).ListIndex > 0 Then
                        DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex - DXForm(BFormIndex).Controls(iFocus).ScrollCount
                        If DXForm(BFormIndex).Controls(iFocus).ListIndex < 0 Then
                            DXForm(BFormIndex).Controls(iFocus).ListIndex = 0
                        End If
                    End If
                    'EMOSTEXT
                    DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    'TEXTCURSORPOSITION
                    iC = DXForm(BFormIndex).Controls(iFocus).ListCount
                    For iA = Len(TextCursorString) To 1 Step -1
                        If Mid(TextCursorString, iA, 2) = vbCrLf Then
                            iC = iC - 1
                            If iC = DXForm(BFormIndex).Controls(iFocus).ListIndex Then
                                TextCursorPosition = iA + 1
                                Exit Sub
                            End If
                        End If
                    Next iA
                    TextCursorPosition = 0
                    'EINDE
                    Exit Sub
                Case 209  ' Page Down ----------------------------COMPLETED---------------------------
                    'RESET TEXT SELECTION START
                    TextSelectedStart = 0
                    'SCROLLCOUNT
                    DXForm(BFormIndex).Controls(iFocus).ScrollCount = Val((DXForm(BFormIndex).Controls(iFocus).Position(0).Bottom - DXForm(BFormIndex).Controls(iFocus).Position(0).Top) / (vFontSize))
                    'SCROLLINDEX
                    If DXForm(BFormIndex).Controls(iFocus).ScrollIndex < DXForm(BFormIndex).Controls(iFocus).ScrollCount Then
                        DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollIndex + DXForm(BFormIndex).Controls(iFocus).ScrollCount
                        If DXForm(BFormIndex).Controls(iFocus).ScrollIndex > DXForm(BFormIndex).Controls(iFocus).ScrollCount Then
                            DXForm(BFormIndex).Controls(iFocus).ScrollIndex = DXForm(BFormIndex).Controls(iFocus).ScrollCount
                        End If
                        'LISTINDEX
                    ElseIf DXForm(BFormIndex).Controls(iFocus).ListIndex < DXForm(BFormIndex).Controls(iFocus).ListCount - DXForm(BFormIndex).Controls(iFocus).ScrollCount - 1 Then
                        DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListIndex + DXForm(BFormIndex).Controls(iFocus).ScrollCount + 1
                        If DXForm(BFormIndex).Controls(iFocus).ListIndex > DXForm(BFormIndex).Controls(iFocus).ListCount - DXForm(BFormIndex).Controls(iFocus).ScrollCount - 1 Then
                            DXForm(BFormIndex).Controls(iFocus).ListIndex = DXForm(BFormIndex).Controls(iFocus).ListCount - DXForm(BFormIndex).Controls(iFocus).ScrollCount - 1
                        End If
                    End If
                    'EMOSTEXT
                    DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    'TEXTCURSORPOSITION
                    iC = 0
                    For iA = 1 To Len(TextCursorString)
                        If Mid(TextCursorString, iA, 2) = vbCrLf Then
                            iC = iC + 1
                            If iC = DXForm(BFormIndex).Controls(iFocus).ScrollCount + DXForm(BFormIndex).Controls(iFocus).ListIndex Then
                                TextCursorPosition = iA + 1
                                Exit Sub
                            End If
                        End If
                    Next iA
                    TextCursorPosition = Len(TextCursorString)
                    'EINDE
                    Exit Sub
                Case 211 ' Delete ----------------------------COMPLETED---------------------------
                    'LOCKED
                    If DXForm(BFormIndex).Controls(iFocus).Locked = False Then
                        'TEXTCURSORPOS
                        If Len(TextCursorString) > TextCursorPosition Then
                            If (Len(TextCursorString) - TextCursorPosition) > 1 Then
                                If Mid(TextCursorString, TextCursorPosition + 1, 2) = vbCrLf Then
                                    iA = 2
                                Else
                                    iA = 1
                                End If
                            Else
                                iA = 1
                            End If
                            TextCursorString = Left(TextCursorString, TextCursorPosition) & Right(TextCursorString, Len(TextCursorString) - TextCursorPosition - iA)
                            'EMOSTEXT
                            DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(iFocus2), DXForm(BFormIndex).Controls(iFocus))
                            'CHANGE
                            Call xForm(BFormIndex).Text1(DXForm(BFormIndex).Controls(iFocus).Index).Change
                        End If
                    End If
                    'EINDE
                    Exit Sub
                Case Else ' Type
                    'de cursor positie loopt achter omdat je werkt met tekst verdeling
                    'en de woorden die langer zijn dan 1 regel zorgen voor 2 stappen
                    'te veel achter dat woordt waardoor het op een verkeerde plek
                    'ingevoegt wordt
                    If DXForm(BFormIndex).Controls(iFocus).Locked = False Then
                        If KeyNames(iK) <> "" Then
                            If ((KeyPressed(42) Or KeyPressed(54)) And iK <> 57) = True Then
                                If Len(TextCursorString) = TextCursorPosition Then
                                    xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = TextCursorString & KeyNames(iK + 300)
                                    TextCursorPosition = TextCursorPosition + Len(KeyNames(iK + 300))
                                Else
                                    xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = Left(TextCursorString, TextCursorPosition) & KeyNames(iK + 300) & Right(TextCursorString, Len(TextCursorString) - TextCursorPosition)
                                    TextCursorPosition = TextCursorPosition + Len(KeyNames(iK + 300))
                                End If
                            Else
                                If Len(TextCursorString) = TextCursorPosition Then
                                    xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = TextCursorString & KeyNames(iK)
                                    TextCursorPosition = TextCursorPosition + Len(KeyNames(iK))
                                Else
                                    xForm(BFormIndex).Controls(iFocus).Text(iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex) = Left(TextCursorString, TextCursorPosition) & KeyNames(iK) & Right(TextCursorString, Len(TextCursorString) - TextCursorPosition)
                                    TextCursorPosition = TextCursorPosition + Len(KeyNames(iK))
                                End If
                            End If
                        End If
                        Call xForm(BFormIndex).Text1(DXForm(BFormIndex).Controls(iFocus).Index).Change
                        DXForm(BFormIndex).Controls(iFocus).EmosText = ScrollText(DXForm(BFormIndex).Controls(iFocus).Text(0), DXForm(BFormIndex).Controls(iFocus))
                    End If
            End Select
        End If
        If DXForm(BFormIndex).Controls(iFocus).Locked = False Then
            If DXForm(BFormIndex).Controls(iFocus).ScrollIndex <> 0 Then
                If UBound(DXForm(BFormIndex).Controls(iFocus).Text()) = iFocus2 + DXForm(BFormIndex).Controls(iFocus).ScrollIndex Then
                    ReDim Preserve DXForm(BFormIndex).Controls(iFocus).Text(UBound(DXForm(BFormIndex).Controls(iFocus).Text) + DXForm(BFormIndex).Controls(iFocus).Colom)
                    DXForm(BFormIndex).Controls(iFocus).ScrollIndex = UBound(DXForm(BFormIndex).Controls(iFocus).Text()) - iFocus2
                    iFocus2 = iFocus2 - DXForm(BFormIndex).Controls(iFocus).Colom + 1
                End If
            End If
        End If
    End Sub

    Sub ScreenCursor(iC As Byte, iA As Integer, lA As Long, iZ As Byte)
        Dim oA As RECT, LB As Long, sA As String
        Static bA As Boolean, vA As Double, iB As Double
        sA = ScreenDataBase(iC, iA, "-1", iZ)
        If iFocus = iA And iC = BFormIndex Then
            If vA < GetTickCount Then
                If bA = False Then
                    vA = 500 + GetTickCount
                    bA = True
                Else
                    vA = 250 + GetTickCount
                    bA = False
                End If
            End If
            If iB <> iA Then
                iB = iA
                bA = True
            End If
            If bA = True And iZ = iFocus2 Then
                If sA <> "" Then
                    If Left(sA, TextCursorPosition) = "" Then
                        LB = 0
                    ElseIf DXForm(iC).Controls(iA).PasswordChar <> "" Then
                        LB = TextWidth(Replace(Space(Len(Left(DXForm(iC).Controls(iA).Text(0), TextCursorPosition))), " ", DXForm(iC).Controls(iA).PasswordChar))
                    Else
                        LB = TextWidth(CStr(Split(Left(TextCursorString, TextCursorPosition), vbCrLf)(CountReturns(0, Left(TextCursorString, TextCursorPosition)))))
                    End If
                Else
                    LB = 0
                End If
                TextCursor.Left = DXForm(iC).Controls(iA).Position(iFocus2).Left + LB
                If DXForm(iC).Controls(iA).Text(iFocus2) <> "" Then
                    TextCursor.Top = DXForm(iC).Controls(iA).Position(iFocus2).Top + (CountReturns(0, Left(TextCursorString, TextCursorPosition)) - (DXForm(iC).Controls(iA).ListIndex)) * ((vFontSize - 2))
                Else
                    TextCursor.Top = DXForm(iC).Controls(iA).Position(iFocus2).Top
                End If
                TextCursor.Right = (4 * FixHeight) + TextCursor.Left
                TextCursor.Bottom = (vFontSize) + TextCursor.Top
                Call D3DX.DrawText(D3DFont, lA, "I", TextCursor, DT_TOP Or DT_LEFT)
            End If
        End If
        oA = DXForm(iC).Controls(iA).Position(iZ)
        oA.Left = (oA.Left + 2)
        Call D3DX.DrawText(D3DFont, lA, sA, oA, DT_TOP Or DT_LEFT)
    End Sub

End Module
