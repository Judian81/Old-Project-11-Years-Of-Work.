Module mTranslate

    Sub FocusSet(iC As Byte, iA As Integer, Optional iB As Byte = 0)
        Dim vA As Byte
        Static yFocusOldForm As Byte
        TextCursorPosition = 0
        TextSelectedCount = 0
        TextSelectedStart = 0
        ReDim TextSelected(TextSelectedCount)
        If iC <> BFormIndex Then
            BFormIndex = iC
            TextList = Empty
        End If
        If iFocus <> iA Then
            TextList = Empty
            If iFocus > -1 And yFocusOldForm <> 0 And yFocusOldForm <= BFormCount Then
                If iFocus2 <= UBound(DXForm(yFocusOldForm).Controls(iFocus).Image()) Then
                    If DXForm(yFocusOldForm).Controls(iFocus).Name = "List1" Then
                        If iFocus2 > 16 Then
                            DXForm(yFocusOldForm).Controls(iFocus).Image(iFocus2) = 85
                        End If
                    ElseIf DXForm(yFocusOldForm).Controls(iFocus).Name = "Combo1" Then
                        DXForm(yFocusOldForm).Controls(iFocus).Image(4) = 116
                    ElseIf DXForm(yFocusOldForm).Controls(iFocus).Name = "Data1" And (DXForm(yFocusOldForm).Controls(iFocus).Image(0) = 116 Or DXForm(yFocusOldForm).Controls(iFocus).Image(0) = 117) Then
                        DXForm(yFocusOldForm).Controls(iFocus).Image(iFocus2) = 116
                    End If
                End If
            End If
            yFocusOldForm = iC
            iFocus = iA
            If DXForm(iC).Controls(iFocus).Name = "List1" And iB = 0 Then
                iFocus2 = 0
                For vA = 17 To DXForm(iC).Controls(iA).PositionCount
                    If DXForm(iC).Controls(iFocus).Image(vA) = 85 Then
                        DXForm(iC).Controls(iFocus).Image(vA) = 86
                        iFocus2 = vA
                        Exit For
                    End If
                Next vA
                If iFocus2 = 0 Then
                    iFocus2 = 17
                End If
            ElseIf DXForm(iC).Controls(iFocus).Name = "Data1" And (DXForm(iC).Controls(iFocus).Colom > 1 Or DXForm(iC).Controls(iFocus).Row > 1) Then
                If UBound(DXForm(iC).Controls(iFocus).Text) >= iB Then
                    iFocus2 = iB
                Else
                    iFocus2 = UBound(DXForm(iC).Controls(iFocus).Text)
                End If
            Else
                iFocus2 = iB
            End If
            If DXForm(iC).Controls(iFocus).Name = "List1" Then
                For vA = 17 To DXForm(iC).Controls(iFocus).PositionCount
                    DXForm(iC).Controls(iFocus).Image(vA) = 84
                Next vA
                DXForm(iC).Controls(iFocus).Image(iFocus2) = 86
                xForm(iC).Controls(iFocus).ListIndex = DXForm(iC).Controls(iFocus).ScrollIndex + iFocus2 - 17
            ElseIf DXForm(iC).Controls(iFocus).Name = "Combo1" Then
                iFocus2 = 4
                DXForm(iC).Controls(iFocus).Image(iFocus2) = 117
            ElseIf (DXForm(yFocusOldForm).Controls(iFocus).Name = "Data1" And (DXForm(yFocusOldForm).Controls(iFocus).Image(0) = 116 Or DXForm(yFocusOldForm).Controls(iFocus).Image(0) = 117)) Then
                DXForm(iC).Controls(iFocus).Image(iFocus2) = 117
            End If
        ElseIf iFocus2 <> iB Then
            If DXForm(iC).Controls(iFocus).Name = "List1" And iB <> 0 Then
                If DXForm(iC).Controls(iFocus).ScrollIndex + iB - 17 <= Query(DXForm(iC).Controls(iFocus).DataBase).ListCount Then
                    If iFocus2 > 16 Then
                        DXForm(iC).Controls(iFocus).Image(iFocus2) = 84
                    End If
                    iFocus2 = iB
                    If iFocus2 > 16 Then
                        DXForm(iC).Controls(iFocus).Image(iFocus2) = 86
                        xForm(iC).Controls(iFocus).ListIndex = DXForm(iC).Controls(iFocus).ScrollIndex + iFocus2 - 17
                    End If
                End If
            ElseIf DXForm(iC).Controls(iFocus).Name = "Data1" And (DXForm(iC).Controls(iFocus).Colom > 1 Or DXForm(iC).Controls(iFocus).Row > 1) Then
                If UBound(DXForm(iC).Controls(iFocus).Text) >= iB Then
                    iFocus2 = iB
                Else
                    iFocus2 = UBound(DXForm(iC).Controls(iFocus).Text)
                End If
            Else
                iFocus2 = iB
            End If
        End If
        '    If DXForm(iC).Controls(iFocus).Name = "Data1" Or DXForm(iC).Controls(iFocus).Name = "Text1" Then
        '        TextCursorString = TextVerdelen(DXForm(iC).Controls(iFocus).Text(iFocus2), XForm(BFormIndex).Font, CLng((DXForm(iC).Controls(iFocus).Position(0).Right - DXForm(iC).Controls(iFocus).Position(0).Left) / FixWidth), DXForm(BFormIndex).Controls(iFocus).MultiLine)
        '    End If
    End Sub

End Module
