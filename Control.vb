Public Class Control

    ' this is the code that could work. the old way is deprecated
    'this is not do
    Property Accesskeys() As String
        Get
            Return BControl.Accesskeys
        End Get
        ' i have removed ByVal because i do not know what it is for ???
        Set(b As String)
            Call ControlSet(BControl)
        End Set
    End Property


    Property Alignment() As Integer
        Get
            Return BControl.Alignment
        End Get
        Set(b As Integer)
            BControl.Alignment = b
            Call ControlSet(BControl)
            'what is shit for?!?!
            'NodeLabelEditEventArgs
        End Set
    End Property

    Property Appearance() As Integer
        Get
            Return BControl.Appearance
        End Get
        Set(b As Integer)
            If BControl.Name = "Label1" Then
                If b = 0 Then
                    BControl.Image(2) = 61
                Else
                    BControl.Image(2) = 63
                End If
            End If
            BControl.Appearance = b
            Call ControlSet(BControl)
        End Set
    End Property


    Property Archive() As Boolean
        Get
            Return BControl.Archive
        End Get
        Set(b As Boolean)
            BControl.Archive = b
            Call ControlSet(BControl)
        End Set
    End Property


    Property BackColor() As Long
        Get
            Return BControl.BackColor
        End Get
        Set(b As Long)
            BControl.BackColor = b
            Call ControlSet(BControl)
        End Set
    End Property


    Property BackStyle() As Integer
        Get
            Return BControl.BackStyle
        End Get
        Set(b As Integer)
            BControl.BackStyle = b
            Call ControlSet(BControl)
        End Set
    End Property


    Property BorderColor() As Integer
        Get
            Return BControl.BorderColor
        End Get
        Set(b As Integer)
            BControl.BorderColor = b
            Call ControlSet(BControl)
        End Set
    End Property


    Property BorderStyle() As Integer
        Get
            Return BControl.BorderStyle
        End Get
        Set(b As Integer)
            BControl.BorderStyle = b
            Call ControlSet(BControl)
        End Set
    End Property


    Property BorderWidth() As Integer
        Get
            Return BControl.BorderWidth
        End Get
        Set(b As Integer)
            BControl.BorderWidth = b
            Call ControlSet(BControl)
        End Set
    End Property


    Property Caption() As String
        Get
            Return BControl.Text(0)
        End Get
        Set(b As String)
            BControl.Text(0) = b
            Call ControlSet(BControl)
        End Set
    End Property


    Sub Change()
        Dim sB As String
        sB = DXForm(BFormIndex).Name
        Select Case DXForm(BFormIndex).Controls(iFocus).Name
            Case "List1"
            Case "Text1"
                'Call aProgram.Program(sB, BControl.Name & "_" & "Change()", BControl.Index)
            Case "Combo1"
                'Call aProgram.Program(sB, BControl.Name & "_" & "Change()", BControl.Index)
            Case "Data1"
                If Val(BControl.DataField) = 2 And BComboList.DataBase <> 0 Then
                    DXForm(BFormIndex).Controls(iFocus).Text(0) = Query(BComboList.DataBase).List(BComboList.ListIndex)
                    DXForm(BFormIndex).Controls(iFocus).EmosText = DXForm(BFormIndex).Controls(iFocus).Text(0)
                End If
                'Call aProgram.Program(sB, BControl.Name & "_" & "Change()", BControl.Index)
            Case "Option1"
        End Select
    End Sub


    Sub Click()
        Dim sB As String
        sB = DXForm(BFormIndex).Name
        'Call aProgram.Program(sB, BControl.Name & "_" & "Click()", BControl.Index)
    End Sub


    Property Columns() As Integer
        Get
            Return BControl.Columns
        End Get
        Set(b As Integer)
            Call ControlSet(BControl)
            BControl.Columns = b
        End Set
    End Property


    Property DataBase() As Integer
        Get
            Return BControl.DataBase
        End Get
        Set(b As Integer)
            Dim vA As Integer, vC As Double, vE As Integer
            If b < 0 Or b > QueryCount Then
                Exit Property
            Else
                If Val(BControl.DataField) = -1 Then
                    vE = BControl.DataBase
                    BControl.DataBase = b
                End If
                BControl.DataBase = b
                If BControl.Name = "" Then
                    Exit Property
                ElseIf BControl.Name = "List1" Or BControl.Name = "ComboList" Then
                    vC = CInt((BControl.Position(4).Bottom - BControl.Position(4).Top) / (17 * FixHeight) - 1)
                    BControl.PositionCount = vC + 17
                    ReDim Preserve BControl.Position(BControl.PositionCount)
                    ReDim Preserve BControl.Image(BControl.PositionCount)
                    For vA = 0 To vC
                        If vA = 0 Then
                            BControl.Image(vA + 17) = 85
                        Else
                            BControl.Image(vA + 17) = 84
                        End If
                        BControl.Position(vA + 17).Left = BControl.Position(4).Left
                        BControl.Position(vA + 17).Top = BControl.Position(4).Top + 17 * FixHeight * vA
                        BControl.Position(vA + 17).Right = BControl.Position(4).Right
                        BControl.Position(vA + 17).Bottom = BControl.Position(vA + 17).Top + 17 * FixHeight - 1
                    Next vA
                    BControl.Position(11).Top = BControl.Position(10).Top
                    BControl.Position(11).Bottom = BControl.Position(10).Top + (10 * FixHeight)
                End If
                Call ControlSet(BControl)
            End If
        End Set
    End Property


    Property DataChanged() As Boolean
        Get
            Return BControl.DataChanged
        End Get
        Set(value As Boolean)
            BControl.DataChanged = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property DataField() As String
        Get
            Return BControl.DataField
            Call ControlSet(BControl)
        End Get
        Set(value As String)
            BControl.DataField = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property DataFormat() As String
        Get
            Return BControl.DataFormat
        End Get
        Set(value As String)
            BControl.DataFormat = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property DataMember() As String
        Get
            Return BControl.DataMember
        End Get
        Set(value As String)
            BControl.DataMember = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property DataSource() As Integer
        Get
            Return BControl.DataSource
        End Get
        Set(value As Integer)
            If BControl.DataSource <> value Then
                BControl.DataSource = value
                If Query(value).ListCount = 0 Then
                    BControl.Text(0) = ""
                Else
                    BControl.Text(0) = Query(value).List(BControl.ScrollIndex)
                End If
                Call ControlSet(BControl)
            End If
        End Set
    End Property


    Sub DblClick()
        Dim sB As String
        sB = DXForm(BFormIndex).Name
        ''''Call aProgram.Program(sB, BControl.Name & "_" & "DblClick()", BControl.Index)
    End Sub


    Property Drive() As String
        Get
            Return BControl.Drive
        End Get
        Set(value As String)
            BControl.Drive = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property Enabled() As Boolean
        Get
            Return BControl.Enabled
        End Get
        Set(value As Boolean)
            BControl.Enabled = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property FileName() As String
        Get
            Return BControl.FileName
        End Get
        Set(value As String)
            BControl.FileName = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property FillStyle() As Integer
        Get
            Return BControl.FillStyle
        End Get
        Set(value As Integer)
            BControl.FillStyle = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property Font() As IFontDisp
        Get
            Font.Bold = BControl.Font.Bold
            Font.Charset = BControl.Font.Charset
            Font.Italic = BControl.Font.Italic
            Font.Name = BControl.Font.Name
            Font.Size = BControl.Font.Size
            Font.Strikethrough = BControl.Font.Strikethrough
            Font.Underline = BControl.Font.Underline
            Font.Weight = BControl.Font.Weight
        End Get
        Set(value As IFontDisp)
            ''''''''Set Font = New StdFont  ?????
            BControl.Font.Bold = value.Bold
            BControl.Font.Charset = value.Charset
            BControl.Font.Italic = value.Italic
            BControl.Font.Name = value.Name
            BControl.Font.Size = value.Size
            BControl.Font.Strikethrough = value.Strikethrough
            BControl.Font.Underline = value.Underline
            BControl.Font.Weight = value.Weight
            Call ControlSet(BControl)
        End Set
    End Property


    Property FontBold() As Boolean
        Get
            Return BControl.Font.Bold
        End Get
        Set(value As Boolean)
            BControl.Font.Bold = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property FontItalic() As Boolean
        Get
            Return BControl.Font.Italic
        End Get
        Set(value As Boolean)
            BControl.Font.Italic = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property FontName() As String
        Get
            Return BControl.Font.Name
        End Get
        Set(value As String)
            BControl.Font.Name = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property FontSize() As Single
        Get
            Return BControl.Font.Size
        End Get
        Set(value As Single)
            BControl.Font.Size = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property FontStrikethru() As Boolean
        Get
            FontStrikethru = BControl.Font.Strikethrough
        End Get
        Set(value As Boolean)
            BControl.Font.Strikethrough = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property FontUnderline() As Boolean
        Get
            FontUnderline = BControl.Font.Underline
        End Get
        Set(value As Boolean)
            BControl.Font.Underline = value
            Call ControlSet(BControl)
        End Set
    End Property

    '------------------------------------lets start with huge amount of code------------------------------------------????
    'the whole code was messed up. i had the same property pages. some times 3 times the samen
    Property ForeColor() As Long
        Get
            ForeColor = BControl.ForeColor
        End Get
        Set(value As Long)
            BControl.ForeColor = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property Height() As Single
        Get
            Height = (BControl.Position(0).Bottom - BControl.Position(0).Top) / FixHeight
        End Get
        Set(value As Single)
            Dim vA As Integer, vD As Integer
            vD = value * FixHeight - BControl.Y2
            BControl.Y2 = value * FixHeight
            For vA = 0 To BControl.PositionCount
                BControl.Position(0).Bottom = BControl.Position(0).Bottom + vD
            Next vA
            Call ControlSet(BControl)
        End Set
    End Property

    Property HelpContextID() As Long
        Get
            HelpContextID = BControl.HelpContextID
        End Get
        Set(value As Long)
            BControl.HelpContextID = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property Image(c As Integer) As Integer
        Get
            Image = BControl.Image(c)
        End Get
        Set(value As Integer)
            BControl.Image(c) = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property Index() As Integer
        Get

            Index = BControl.Index
        End Get
        Set(value As Integer)
            ''''does this have to do something???? lets check this better in future
        End Set
    End Property

    Property IntegralHeight() As Boolean
        Get
            IntegralHeight = BControl.IntegralHeight
        End Get
        Set(value As Boolean)
            ''''does nothing? is it needed???
        End Set
    End Property

    Property Interval() As Long
        Get
            Interval = BControl.Interval
        End Get
        Set(value As Long)
            BControl.Interval = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property ItemData(lnteger As Integer) As Long
        Get
            ItemData = Query(BControl.DataBase).ItemData(lnteger)
        End Get
        Set(value As Long)
            Query(BControl.DataBase).ItemData(lnteger) = value
        End Set
    End Property


    Property LargeChange() As Integer
        Get
            LargeChange = BControl.LargeChange
        End Get
        Set(value As Integer)
            BControl.LargeChange = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property Left() As Single
        Get
            Left = BControl.Position(0).Left / FixWidth
        End Get
        Set(value As Single)
            Dim vA As Integer, vD As Integer
            BControl.X1 = value * FixWidth
            vD = (value * FixWidth) - BControl.Position(0).Left
            For vA = 0 To BControl.PositionCount
                BControl.Position(0).Left = BControl.Position(0).Left + vD
                BControl.Position(0).Right = BControl.Position(0).Right + vD
            Next vA
            Call ControlSet(BControl)
        End Set
    End Property


    Property List(lnteger As Integer) As String
        Get
            List = Query(BControl.DataBase).List(lnteger)
        End Get
        Set(value As String)
            Query(BControl.DataBase).List(lnteger) = value
        End Set
    End Property


    Property ListCount() As Integer
        Get
            ListCount = Query(BControl.DataBase).ListCount
        End Get
        Set(value As Integer)

        End Set
    End Property


    Property ListIndex() As Integer
        Get
            ListIndex = BControl.ListIndex
        End Get
        Set(value As Integer)
            BControl.ListIndex = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property Locked() As Boolean
        Get
            Locked = BControl.Locked
        End Get
        Set(value As Boolean)
            BControl.Locked = value
            Call ControlSet(BControl)
        End Set
    End Property


    Sub LostFocus()
        'this do not do much do they?!
        'Dim sB As String
        'sB = DXForm(BFormIndex).Name
        Select Case BControl.Name
            Case "Combo1", "Text1", "Data1"
                BControl.SelLength = 0
        End Select
    End Sub

    Property Max() As Integer
        Get
            Max = BControl.Max
        End Get
        Set(value As Integer)
            BControl.Max = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property Min() As Integer
        Get
            Min = BControl.Min
        End Get
        Set(value As Integer)
            BControl.Min = value
            Call ControlSet(BControl)
        End Set
    End Property


    'Property MouseIcon() As IPictureDisp
    '    Get
    '        MouseIcon = BControl.MouseIcon
    '    End Get
    '    Set(value As IPictureDisp)
    '        BControl.MouseIcon = value
    '        Call ControlSet(BControl)
    '    End Set
    'End Property

    Property MousePointer() As Integer
        Get
            MousePointer = BControl.MousePointer
        End Get
        Set(value As Integer)
            BControl.MousePointer = value
            Call ControlSet(BControl)
        End Set
    End Property


    ''''Sub Move(Left As Single, Optional Top As Single = -1, Optional Width As Single = -1, Optional Height As Single = -1, Optional Resize As Boolean = True)
    ''''    Dim vA As Integer, vD As Integer
    ''''    Dim b As Integer, iQ As Integer, iB As Integer
    ''''    Dim iLeft As Single, iTop As Single, iWidth As Single, iHeight As Single
    ''''    If Resize = True Then
    ''''        iLeft = Left * FixWidth
    ''''        iTop = Top * FixHeight
    ''''        iWidth = Width * FixWidth
    ''''        iHeight = Height * FixHeight
    ''''    Else
    ''''        iLeft = Left
    ''''        iTop = Top
    ''''        iWidth = Width
    ''''        iHeight = Height
    ''''    End If
    ''''    BControl.X1 = iLeft
    ''''    BControl.X2 = iWidth
    ''''    BControl.Y1 = iTop
    ''''    BControl.Y2 = iHeight
    ''''    Select Case BControl.Name
    ''''        Case "Combo1", "Drive1"
    ''''            BControl.Position(0).Left = iLeft
    ''''            BControl.Position(0).Top = iTop
    ''''            BControl.Position(0).Right = BControl.Position(0).Left + (17 * FixWidth)
    ''''            BControl.Position(0).Bottom = BControl.Position(0).Top + (10 * FixHeight)
    ''''            BControl.Position(2).Left = iLeft + iWidth - (17 * FixWidth)
    ''''            BControl.Position(2).Top = iTop
    ''''            BControl.Position(2).Right = iLeft + iWidth
    ''''            BControl.Position(2).Bottom = BControl.Position(2).Top + (10 * FixHeight)
    ''''            BControl.Position(1).Left = BControl.Position(0).Right
    ''''            BControl.Position(1).Top = iTop
    ''''            BControl.Position(1).Right = BControl.Position(2).Left
    ''''            BControl.Position(1).Bottom = BControl.Position(1).Top + (10 * FixHeight)
    ''''            BControl.Position(3).Left = iLeft
    ''''            BControl.Position(3).Top = BControl.Position(0).Bottom
    ''''            BControl.Position(3).Right = BControl.Position(3).Left + (9 * FixWidth)
    ''''            BControl.Position(3).Bottom = BControl.Position(3).Top + iHeight - (22 * FixHeight)
    ''''            BControl.Position(4).Left = BControl.Position(3).Right
    ''''            BControl.Position(4).Top = BControl.Position(0).Bottom
    ''''            BControl.Position(4).Right = BControl.Position(4).Left + iWidth - (34 * FixWidth)
    ''''            BControl.Position(4).Bottom = BControl.Position(4).Top + (17 * FixHeight)
    ''''            BControl.Position(9).Left = BControl.Position(4).Right + 1
    ''''            BControl.Position(9).Top = BControl.Position(0).Bottom
    ''''            BControl.Position(9).Right = BControl.Position(9).Left + (17 * FixWidth)
    ''''            BControl.Position(9).Bottom = BControl.Position(4).Bottom
    ''''            BControl.Position(5).Left = (iLeft + iWidth) - (7 * FixWidth)
    ''''            BControl.Position(5).Top = BControl.Position(0).Bottom
    ''''            BControl.Position(5).Right = iLeft + iWidth
    ''''            BControl.Position(5).Bottom = BControl.Position(4).Bottom
    ''''            BControl.Position(6).Left = iLeft
    ''''            BControl.Position(6).Top = BControl.Position(4).Bottom
    ''''            BControl.Position(6).Right = BControl.Position(0).Right
    ''''            BControl.Position(6).Bottom = BControl.Position(6).Top + (12 * FixHeight)
    ''''            BControl.Position(7).Left = BControl.Position(6).Right
    ''''            BControl.Position(7).Top = BControl.Position(4).Bottom
    ''''            BControl.Position(7).Right = BControl.Position(7).Left + iWidth - (34 * FixWidth)
    ''''            BControl.Position(7).Bottom = BControl.Position(6).Bottom
    ''''            BControl.Position(8).Left = BControl.Position(7).Right
    ''''            BControl.Position(8).Top = BControl.Position(4).Bottom
    ''''            BControl.Position(8).Right = iLeft + iWidth
    ''''            BControl.Position(8).Bottom = BControl.Position(6).Bottom
    ''''        Case "BorderH", "BorderV", "Command1", "Image1", "Option1", "Check1", "ButtonPress", "Picture1", "Frame1", "Line1", "Shape1"
    ''''            If iLeft <> -1 Then
    ''''                vD = (iLeft) - BControl.Position(0).Left
    ''''                For vA = 0 To BControl.PositionCount
    ''''                    BControl.Position(0).Left = BControl.Position(0).Left + vD
    ''''                    BControl.Position(0).Right = BControl.Position(0).Right + vD
    ''''                Next vA
    ''''            End If
    ''''            If iTop <> -1 Then
    ''''                vD = (iTop) - BControl.Position(0).Top
    ''''                For vA = 0 To BControl.PositionCount
    ''''                    BControl.Position(0).Top = BControl.Position(0).Top + vD
    ''''                    BControl.Position(0).Bottom = BControl.Position(0).Bottom + vD
    ''''                Next vA
    ''''            End If
    ''''            If iWidth <> -1 Then
    ''''                vD = ((iWidth) + BControl.Position(0).Left)
    ''''                For vA = 0 To BControl.PositionCount
    ''''                    BControl.Position(0).Right = vD
    ''''                Next vA
    ''''            End If
    ''''            If iHeight <> -1 Then
    ''''                vD = ((iHeight) + BControl.Position(0).Top)
    ''''                For vA = 0 To BControl.PositionCount
    ''''                    BControl.Position(0).Bottom = vD
    ''''                Next vA
    ''''            End If
    ''''        Case "Data1"
    ''''            If BControl.BorderStyle = 1 Then
    ''''                If BControl.Image(0) = 111 Or BControl.Image(0) = 116 Then
    ''''                    BControl.PositionCount = 9
    ''''                    ReDim Preserve BControl.Image(9)
    ''''                    ReDim Preserve BControl.Position(9)
    ''''                    BControl.Image(0) = 116
    ''''                    BControl.Image(1) = 112
    ''''                    BControl.Image(2) = 113
    ''''                    BControl.Image(3) = 114
    ''''                    BControl.Image(4) = 111
    ''''                    BControl.Image(5) = 118
    ''''                    BControl.Image(6) = 119
    ''''                    BControl.Image(7) = 120
    ''''                    BControl.Image(8) = 121
    ''''                    BControl.Image(9) = 11
    ''''                Else
    ''''                    BControl.PositionCount = 8
    ''''                    ReDim Preserve BControl.Image(8)
    ''''                    ReDim Preserve BControl.Position(8)
    ''''                    BControl.Image(1) = 71
    ''''                    BControl.Image(2) = 72
    ''''                    BControl.Image(3) = 73
    ''''                    BControl.Image(4) = 70
    ''''                    BControl.Image(5) = 75
    ''''                    BControl.Image(6) = 76
    ''''                    BControl.Image(7) = 77
    ''''                    BControl.Image(8) = 78
    ''''                End If
    ''''                BControl.Position(4).Left = iLeft
    ''''                BControl.Position(4).Top = iTop
    ''''                BControl.Position(4).Right = iLeft + (17 * FixWidth)
    ''''                BControl.Position(4).Bottom = iTop + (10 * FixHeight)
    ''''                BControl.Position(1).Left = BControl.Position(4).Right 'iLeft + (17 * FixWidth)
    ''''                BControl.Position(1).Top = iTop
    ''''                BControl.Position(1).Right = BControl.Position(1).Left + iWidth - (34 * FixWidth)
    ''''                BControl.Position(1).Bottom = BControl.Position(4).Bottom 'BControl.Position(1).top + (10 * FixHeight)
    ''''                BControl.Position(2).Left = BControl.Position(1).Right  'iLeft + iWidth - (17 * FixWidth)
    ''''                BControl.Position(2).Top = iTop
    ''''                BControl.Position(2).Right = BControl.Position(2).Left + (17 * FixWidth)
    ''''                BControl.Position(2).Bottom = BControl.Position(4).Bottom 'BControl.Position(2).top + (10 * FixHeight)
    ''''                BControl.Position(3).Left = iLeft
    ''''                BControl.Position(3).Top = BControl.Position(4).Bottom  'iTop + (10 * FixHeight)
    ''''                BControl.Position(3).Right = BControl.Position(3).Left + (9 * FixWidth)
    ''''                BControl.Position(3).Bottom = BControl.Position(3).Top + iHeight - (22 * FixHeight)
    ''''                BControl.Position(5).Right = BControl.Position(2).Right 'BControl.Position(5).left + (7 * FixWidth)
    ''''                BControl.Position(5).Left = BControl.Position(2).Right - (7 * FixWidth) 'iLeft + iWidth - (7 * FixWidth)
    ''''                BControl.Position(5).Top = BControl.Position(3).Top 'iTop + (10 * FixHeight)
    ''''                BControl.Position(5).Bottom = BControl.Position(3).Bottom 'BControl.Position(5).top + iHeight - (22 * FixHeight)
    ''''                Select Case b
    ''''                    Case 0
    ''''                        BControl.Position(0).Left = BControl.Position(3).Right 'iLeft + (9 * FixWidth)
    ''''                        BControl.Position(0).Top = iTop + (10 * FixHeight)
    ''''                        BControl.Position(0).Right = BControl.Position(5).Left 'BControl.Position(0).left + iWidth - (16 * FixWidth)
    ''''                        BControl.Position(0).Bottom = BControl.Position(3).Bottom 'BControl.Position(0).top + iHeight - (22 * FixHeight)
    ''''                    Case 1
    ''''                        BControl.Position(0).Left = iLeft + (11 * FixWidth)
    ''''                        BControl.Position(0).Top = iTop + (10 * FixHeight)
    ''''                        BControl.Position(0).Right = BControl.Position(0).Left + iWidth - (20 * FixWidth)
    ''''                        BControl.Position(0).Bottom = BControl.Position(0).Top + iHeight - (22 * FixHeight) - (17 * FixHeight)
    ''''                    REM Call ControlAdd("ScrollH", o, p, v, f, B, ileft + (11 * FixWidth), itop + (10 * FixHeight), iwidth - (20 * FixWidth), iheight - (22 * FixHeight), i, t, d)
    ''''                    Case 2
    ''''                        BControl.Position(0).Left = iLeft + (11 * FixWidth)
    ''''                        BControl.Position(0).Top = iTop + (10 * FixHeight)
    ''''                        BControl.Position(0).Right = BControl.Position(0).Left + iWidth - (20 * FixWidth) - (17 * FixWidth)
    ''''                        BControl.Position(0).Bottom = BControl.Position(0).Top + iHeight - (22 * FixHeight)
    ''''                    REM Call ControlAdd("ScrollV", o, p, v, f, B, ileft + (11 * FixWidth), itop + (10 * FixHeight), iwidth - (20 * FixWidth), iheight - (22 * FixHeight), i, t, d)
    ''''                    Case 3
    ''''                        BControl.Position(0).Left = iLeft + (11 * FixWidth)
    ''''                        BControl.Position(0).Top = iTop + (10 * FixHeight)
    ''''                        BControl.Position(0).Right = BControl.Position(0).Left + iWidth - (20 * FixWidth) - (17 * FixWidth)
    ''''                        BControl.Position(0).Bottom = BControl.Position(0).Top + iHeight - (22 * FixHeight) - (17 * FixHeight)
    ''''                        REM Call ControlAdd("ScrollH", o, p, v, f, B, ileft + (11 * FixWidth), itop + (10 * FixHeight), iwidth - (37 * FixWidth), iheight - (22 * FixHeight), i, t, d)
    ''''                        REM Call ControlAdd("ScrollV", o, p, v, f, B, ileft + (11 * FixWidth), itop + (10 * FixHeight), iwidth - (20 * FixWidth), iheight - (39 * FixHeight), i, t, d)
    ''''                End Select
    ''''                BControl.Position(6).Left = iLeft
    ''''                BControl.Position(6).Top = BControl.Position(3).Bottom 'iTop + iHeight - (12 * FixHeight)
    ''''                BControl.Position(6).Right = BControl.Position(4).Right 'BControl.Position(6).left + (17 * FixWidth)
    ''''                BControl.Position(6).Bottom = BControl.Position(6).Top + (12 * FixHeight)
    ''''                BControl.Position(7).Left = iLeft + (17 * FixWidth)
    ''''                BControl.Position(7).Top = BControl.Position(0).Bottom 'iTop + iHeight - (12 * FixHeight)
    ''''                BControl.Position(7).Right = BControl.Position(2).Left 'BControl.Position(7).left + iWidth - (34 * FixWidth)
    ''''                BControl.Position(7).Bottom = BControl.Position(6).Bottom 'BControl.Position(7).top + (12 * FixHeight)
    ''''                BControl.Position(8).Left = BControl.Position(2).Left 'iLeft + iWidth - (17 * FixWidth)
    ''''                BControl.Position(8).Top = BControl.Position(5).Bottom 'iTop + iHeight - (12 * FixHeight)
    ''''                BControl.Position(8).Right = BControl.Position(5).Right 'BControl.Position(8).left + (17 * FixWidth)
    ''''                BControl.Position(8).Bottom = BControl.Position(6).Bottom 'BControl.Position(8).top + (12 * FixHeight)
    ''''                If BControl.Image(0) = 116 Then
    ''''                    'BControl.Position(0).left = iLeft + (9 * FixWidth)
    ''''                    'BControl.Position(0).top = iTop + (10 * FixHeight)
    ''''                    'BControl.Position(0).Bottom = BControl.Position(0).top + (17 * FixHeight)
    ''''                    BControl.Position(9).Left = BControl.Position(5).Left - (17 * FixWidth) 'BControl.Position(0).Right + (1 * FixWidth)
    ''''                    BControl.Position(9).Top = BControl.Position(0).Top  'iTop + (10 * FixHeight)
    ''''                    BControl.Position(9).Right = BControl.Position(5).Left 'BControl.Position(9).left + (17 * FixWidth)
    ''''                    BControl.Position(9).Bottom = BControl.Position(0).Bottom 'BControl.Position(9).top + (17 * FixHeight)
    ''''                    BControl.Position(0).Right = BControl.Position(9).Left 'BControl.Position(0).left + iWidth - (34 * FixWidth)
    ''''                End If
    ''''            ElseIf BControl.Tag = "-1" Then
    ''''                BControl.Position(1).Left = 0
    ''''                BControl.Position(1).Right = 0
    ''''                BControl.Position(1).Top = 0
    ''''                BControl.Position(1).Bottom = 0
    ''''                BControl.Position(0).Left = iLeft
    ''''                BControl.Position(0).Right = iLeft + iWidth
    ''''                BControl.Position(0).Top = iTop
    ''''                BControl.Position(0).Bottom = iTop + iHeight
    ''''            Else
    ''''                BControl.Position(1).Left = iLeft
    ''''                BControl.Position(1).Right = BControl.Position(1).Left + iWidth / 3
    ''''                BControl.Position(1).Top = iTop
    ''''                BControl.Position(1).Bottom = BControl.Position(1).Top + iHeight
    ''''                BControl.Position(0).Left = BControl.Position(1).Right 'iLeft + iWidth / 3
    ''''                BControl.Position(0).Right = iLeft + iWidth 'BControl.Position(0).left + iWidth / 3 * 2
    ''''                BControl.Position(0).Top = iTop
    ''''                BControl.Position(0).Bottom = BControl.Position(1).Bottom 'BControl.Position(0).top + iHeight
    ''''            End If
    ''''        Case "Label1"
    ''''            If BControl.BorderStyle = 0 Then
    ''''                BControl.PositionCount = 0
    ''''                ReDim BControl.Image(0)
    ''''                ReDim BControl.Position(0)
    ''''                BControl.Image(0) = 93
    ''''                BControl.Position(0).Left = iLeft
    ''''                BControl.Position(0).Right = BControl.Position(0).Left + iWidth
    ''''                BControl.Position(0).Top = iTop
    ''''                BControl.Position(0).Bottom = BControl.Position(0).Top + iHeight
    ''''            Else
    ''''                BControl.PositionCount = 2
    ''''                ReDim BControl.Image(2)
    ''''                ReDim BControl.Position(2)
    ''''                BControl.Image(0) = 60
    ''''                BControl.Image(1) = 64
    ''''                If BControl.Appearance = 1 Then
    ''''                    BControl.Image(2) = 63
    ''''                Else
    ''''                    BControl.Image(2) = 61
    ''''                End If
    ''''                BControl.Position(0).Left = iLeft
    ''''                BControl.Position(0).Right = BControl.Position(0).Left + (25 * FixWidth)
    ''''                BControl.Position(0).Top = iTop
    ''''                BControl.Position(0).Bottom = BControl.Position(0).Top + iHeight
    ''''                BControl.Position(2).Left = BControl.Position(0).Right 'iLeft + (25 * FixWidth)
    ''''                BControl.Position(2).Right = BControl.Position(2).Left + iWidth - (50 * FixWidth)
    ''''                BControl.Position(2).Top = iTop
    ''''                BControl.Position(2).Bottom = BControl.Position(0).Bottom 'BControl.Position(2).top + iHeight
    ''''                BControl.Position(1).Left = BControl.Position(2).Right 'iLeft + iWidth - (25 * FixWidth)
    ''''                BControl.Position(1).Right = iLeft + iWidth 'BControl.Position(1).left + (25 * FixWidth)
    ''''                BControl.Position(1).Top = iTop
    ''''                BControl.Position(1).Bottom = BControl.Position(0).Bottom 'BControl.Position(1).top + iHeight
    ''''            End If
    ''''        Case "List1", "Dir1", "File1", "ComboList"
    ''''            BControl.Position(0).Left = iLeft
    ''''            BControl.Position(0).Top = iTop
    ''''            BControl.Position(0).Right = BControl.Position(0).Left + (17 * FixWidth)
    ''''            BControl.Position(0).Bottom = BControl.Position(0).Top + (10 * FixHeight)
    ''''            BControl.Position(2).Left = iLeft + iWidth - (17 * FixWidth)
    ''''            BControl.Position(2).Top = iTop
    ''''            BControl.Position(2).Right = iLeft + iWidth
    ''''            BControl.Position(2).Bottom = BControl.Position(2).Top + (10 * FixHeight)
    ''''            BControl.Position(1).Left = BControl.Position(0).Right
    ''''            BControl.Position(1).Top = iTop
    ''''            BControl.Position(1).Right = BControl.Position(2).Left
    ''''            BControl.Position(1).Bottom = BControl.Position(1).Top + (10 * FixHeight)
    ''''            BControl.Position(3).Left = iLeft
    ''''            BControl.Position(3).Top = BControl.Position(0).Bottom
    ''''            BControl.Position(3).Right = BControl.Position(3).Left + (9 * FixWidth)
    ''''            BControl.Position(3).Bottom = BControl.Position(3).Top + iHeight - (22 * FixHeight)
    ''''            Select Case b
    ''''                Case 0
    ''''                    iQ = iHeight - (22 * FixHeight)
    ''''                    iB = iWidth - (34 * FixWidth)
    ''''                    BControl.Position(9).Left = iLeft + (11 * FixWidth) + iWidth - (35 * FixWidth)
    ''''                    BControl.Position(9).Top = iTop + (10 * FixHeight)
    ''''                    BControl.Position(9).Right = BControl.Position(9).Left + (17 * FixWidth)
    ''''                    BControl.Position(9).Bottom = BControl.Position(9).Top + (17 * FixHeight)
    ''''                    BControl.Position(12).Left = iLeft + (11 * FixWidth) + iWidth - (35 * FixWidth)
    ''''                    BControl.Position(12).Top = iTop + (10 * FixHeight) + iHeight - (22 * FixHeight) - (17 * FixHeight)
    ''''                    BControl.Position(12).Right = BControl.Position(12).Left + (17 * FixWidth)
    ''''                    BControl.Position(12).Bottom = BControl.Position(12).Top + (17 * FixHeight)
    ''''                    BControl.Position(10).Left = iLeft + (11 * FixWidth) + iWidth - (35 * FixWidth)
    ''''                    BControl.Position(10).Top = BControl.Position(9).Bottom + (2 * FixWidth)
    ''''                    BControl.Position(10).Right = BControl.Position(10).Left + (17 * FixWidth)
    ''''                    BControl.Position(10).Bottom = BControl.Position(12).Top - (2 * FixWidth)
    ''''                    BControl.Position(11).Left = iLeft + (11 * FixWidth) + iWidth - (35 * FixWidth)
    ''''                    BControl.Position(11).Right = BControl.Position(11).Left + (17 * FixWidth)
    ''''                    BControl.Position(11).Top = BControl.Position(10).Top
    ''''                    BControl.Position(11).Bottom = BControl.Position(10).Top + (10 * FixHeight)
    ''''                    BControl.Position(4).Left = BControl.Position(3).Right
    ''''                    BControl.Position(4).Top = BControl.Position(0).Bottom
    ''''                    BControl.Position(4).Right = BControl.Position(4).Left + iB
    ''''                    BControl.Position(4).Bottom = BControl.Position(4).Top + iQ
    ''''                Case 2
    ''''                    iQ = iHeight - (22 * FixHeight)
    ''''                    iB = iWidth - (20 * FixWidth)
    ''''                    BControl.Position(4).Left = iLeft + (11 * FixWidth)
    ''''                    BControl.Position(4).Top = iTop + (10 * FixHeight)
    ''''                    BControl.Position(4).Right = BControl.Position(4).Left + iB
    ''''                    BControl.Position(4).Bottom = BControl.Position(4).Top + iQ
    ''''            End Select
    ''''            BControl.Position(5).Left = (iLeft + iWidth) - (7 * FixWidth)
    ''''            BControl.Position(5).Top = BControl.Position(0).Bottom
    ''''            BControl.Position(5).Right = iLeft + iWidth
    ''''            BControl.Position(5).Bottom = BControl.Position(4).Bottom
    ''''            BControl.Position(6).Left = iLeft
    ''''            BControl.Position(6).Top = BControl.Position(4).Bottom
    ''''            BControl.Position(6).Right = BControl.Position(0).Right
    ''''            BControl.Position(6).Bottom = BControl.Position(6).Top + (12 * FixHeight)
    ''''            BControl.Position(7).Left = BControl.Position(6).Right
    ''''            BControl.Position(7).Top = BControl.Position(4).Bottom
    ''''            BControl.Position(7).Right = BControl.Position(7).Left + iWidth - (34 * FixWidth)
    ''''            BControl.Position(7).Bottom = BControl.Position(6).Bottom
    ''''            BControl.Position(8).Left = BControl.Position(7).Right
    ''''            BControl.Position(8).Top = BControl.Position(4).Bottom
    ''''            BControl.Position(8).Right = iLeft + iWidth
    ''''            BControl.Position(8).Bottom = BControl.Position(6).Bottom
    ''''        Case "VScroll1"
    ''''            BControl.Position(0).Left = iLeft + iWidth - (17 * FixWidth)
    ''''            BControl.Position(0).Top = iTop
    ''''            BControl.Position(0).Right = BControl.Position(0).Left + (17 * FixWidth)
    ''''            BControl.Position(0).Bottom = BControl.Position(0).Top + (17 * FixHeight)
    ''''            BControl.Position(1).Left = iLeft + iWidth - (17 * FixWidth)
    ''''            BControl.Position(1).Top = iTop + (25 * FixHeight)
    ''''            BControl.Position(1).Right = BControl.Position(1).Left + (17 * FixWidth)
    ''''            BControl.Position(1).Bottom = BControl.Position(1).Top + iHeight - 50 * FixHeight
    ''''            BControl.Position(2).Left = iLeft + iWidth - (17 * FixWidth)
    ''''            BControl.Position(2).Top = iTop + (25 * FixHeight)
    ''''            BControl.Position(2).Right = BControl.Position(2).Left + (17 * FixWidth)
    ''''            BControl.Position(2).Bottom = BControl.Position(2).Top + 10 * FixHeight
    ''''            BControl.Position(3).Left = iLeft + iWidth - (17 * FixWidth)
    ''''            BControl.Position(3).Top = iTop + iHeight - (17 * FixHeight)
    ''''            BControl.Position(3).Right = BControl.Position(3).Left + (17 * FixWidth)
    ''''            BControl.Position(3).Bottom = BControl.Position(3).Top + (17 * FixHeight)
    ''''        Case "HScroll1"
    ''''        Case "Text1"
    ''''            If BControl.BorderStyle = 0 Then
    ''''                BControl.PositionCount = 0
    ''''                ReDim BControl.Image(0)
    ''''                ReDim BControl.Position(0)
    ''''                BControl.Image(0) = 74
    ''''                BControl.Position(0).Left = iLeft
    ''''                BControl.Position(0).Right = BControl.Position(0).Left + iWidth
    ''''                BControl.Position(0).Top = iTop
    ''''                BControl.Position(0).Bottom = BControl.Position(0).Top + iHeight
    ''''            Else
    ''''                BControl.PositionCount = 16
    ''''                ReDim BControl.Image(16)
    ''''                ReDim BControl.Position(16)
    ''''                BControl.Image(0) = 74
    ''''                BControl.Image(1) = 71
    ''''                BControl.Image(2) = 72
    ''''                BControl.Image(3) = 73
    ''''                BControl.Image(4) = 70
    ''''                BControl.Image(5) = 75
    ''''                BControl.Image(6) = 76
    ''''                BControl.Image(7) = 77
    ''''                BControl.Image(8) = 78
    ''''                BControl.Image(9) = 0
    ''''                BControl.Image(10) = 0
    ''''                BControl.Image(11) = 0
    ''''                BControl.Image(12) = 0
    ''''                BControl.Image(13) = 0
    ''''                BControl.Image(14) = 0
    ''''                BControl.Image(15) = 0
    ''''                BControl.Image(16) = 0
    ''''                BControl.Position(4).Left = iLeft
    ''''                BControl.Position(4).Top = iTop
    ''''                BControl.Position(4).Right = BControl.Position(4).Left + (17 * FixWidth)
    ''''                BControl.Position(4).Bottom = BControl.Position(4).Top + (10 * FixHeight)
    ''''                BControl.Position(1).Left = BControl.Position(4).Right
    ''''                BControl.Position(1).Top = iTop
    ''''                BControl.Position(1).Right = BControl.Position(1).Left + iWidth - (34 * FixWidth)
    ''''                BControl.Position(1).Bottom = BControl.Position(4).Bottom
    ''''                BControl.Position(2).Left = BControl.Position(1).Right
    ''''                BControl.Position(2).Top = iTop
    ''''                BControl.Position(2).Right = BControl.Position(2).Left + (17 * FixWidth)
    ''''                BControl.Position(2).Bottom = BControl.Position(4).Bottom
    ''''                BControl.Position(3).Left = iLeft
    ''''                BControl.Position(3).Top = BControl.Position(4).Bottom
    ''''                BControl.Position(3).Right = BControl.Position(3).Left + (9 * FixWidth)
    ''''                BControl.Position(3).Bottom = BControl.Position(3).Top + iHeight - (22 * FixHeight)
    ''''                Select Case b
    ''''                    Case 0
    ''''                        BControl.Position(0).Left = BControl.Position(3).Right
    ''''                        BControl.Position(0).Top = BControl.Position(2).Bottom
    ''''                        BControl.Position(0).Right = BControl.Position(0).Left + iWidth - (16 * FixWidth)
    ''''                        BControl.Position(0).Bottom = BControl.Position(0).Top + iHeight - (22 * FixHeight)
    ''''                    Case 1
    ''''                        Stop
    ''''                        BControl.Position(0).Left = iLeft + (11 * FixWidth)
    ''''                        BControl.Position(0).Top = iTop + (10 * FixHeight)
    ''''                        BControl.Position(0).Right = BControl.Position(0).Left + iWidth - (20 * FixWidth)
    ''''                        BControl.Position(0).Bottom = BControl.Position(0).Top + iHeight - (22 * FixHeight) - (17 * FixHeight)
    ''''                    REM Call ControlAdd("ScrollH", o, p, v, f, B, ileft + (11 * FixWidth), itop + (10 * FixHeight), iwidth - (20 * FixWidth), iheight - (22 * FixHeight), i, t, d)
    ''''                    Case 2
    ''''                        Stop
    ''''                        BControl.Position(0).Left = iLeft + (11 * FixWidth)
    ''''                        BControl.Position(0).Top = iTop + (10 * FixHeight)
    ''''                        BControl.Position(0).Right = BControl.Position(0).Left + iWidth - (20 * FixWidth) - (17 * FixWidth)
    ''''                        BControl.Position(0).Bottom = BControl.Position(0).Top + iHeight - (22 * FixHeight)
    ''''                    REM Call ControlAdd("ScrollV", o, p, v, f, B, ileft + (11 * FixWidth), itop + (10 * FixHeight), iwidth - (20 * FixWidth), iheight - (22 * FixHeight), i, t, d)
    ''''                    Case 3
    ''''                        Stop
    ''''                        BControl.Position(0).Left = iLeft + (11 * FixWidth)
    ''''                        BControl.Position(0).Top = iTop + (10 * FixHeight)
    ''''                        BControl.Position(0).Right = BControl.Position(0).Left + iWidth - (20 * FixWidth) - (17 * FixWidth)
    ''''                        BControl.Position(0).Bottom = BControl.Position(0).Top + iHeight - (22 * FixHeight) - (17 * FixHeight)
    ''''                        REM Call ControlAdd("ScrollH", o, p, v, f, B, ileft + (11 * FixWidth), itop + (10 * FixHeight), iwidth - (37 * FixWidth), iheight - (22 * FixHeight), i, t, d)
    ''''                        REM Call ControlAdd("ScrollV", o, p, v, f, B, ileft + (11 * FixWidth), itop + (10 * FixHeight), iwidth - (20 * FixWidth), iheight - (39 * FixHeight), i, t, d)
    ''''                End Select
    ''''                BControl.Position(5).Left = BControl.Position(0).Right
    ''''                BControl.Position(5).Top = BControl.Position(3).Top
    ''''                BControl.Position(5).Right = BControl.Position(2).Right
    ''''                BControl.Position(5).Bottom = BControl.Position(3).Bottom
    ''''                BControl.Position(6).Left = iLeft
    ''''                BControl.Position(6).Top = BControl.Position(5).Bottom
    ''''                BControl.Position(6).Right = BControl.Position(4).Right
    ''''                BControl.Position(6).Bottom = BControl.Position(6).Top + (12 * FixHeight)
    ''''                BControl.Position(7).Left = BControl.Position(6).Right
    ''''                BControl.Position(7).Top = BControl.Position(0).Bottom
    ''''                BControl.Position(7).Right = BControl.Position(1).Right
    ''''                BControl.Position(7).Bottom = BControl.Position(6).Bottom
    ''''                BControl.Position(8).Left = BControl.Position(7).Right
    ''''                BControl.Position(8).Top = BControl.Position(6).Top
    ''''                BControl.Position(8).Right = BControl.Position(8).Left + (17 * FixWidth)
    ''''                BControl.Position(8).Bottom = BControl.Position(6).Bottom
    ''''            End If
    ''''    End Select
    ''''    Call ControlSet(BControl)
    ''''End Sub


    Property MultiLine() As Boolean
        Get
            MultiLine = BControl.MultiLine
        End Get
        Set(value As Boolean)
            BControl.MultiLine = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property MultiSelect() As Integer
        Get
            MultiSelect = BControl.MultiSelect
        End Get
        Set(value As Integer)
            'do nothing??? why???
        End Set
    End Property

    Property Name() As String
        Get
            Name = BControl.Name
        End Get
        Set(value As String)
            BControl.Name = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property NewIndex() As Integer
        Get
            NewIndex = BControl.NewIndex
        End Get
        Set(value As Integer)

        End Set
    End Property

    Property Normal() As Boolean
        Get
            Normal = BControl.Normal
        End Get
        Set(value As Boolean)
            BControl.Normal = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property Path() As String
        Get
            Path = BControl.Path
        End Get
        Set(value As String)
            BControl.Path = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property PasswordChar() As String
        Get
            PasswordChar = BControl.PasswordChar
        End Get
        Set(value As String)
            BControl.PasswordChar = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property sReadOnly() As Boolean
        Get
            Return BControl.sReadOnly
        End Get
        Set(value As Boolean)
            BControl.sReadOnly = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property RightToLeft() As Boolean
        Get
            RightToLeft = BControl.RightToLeft
        End Get
        Set(value As Boolean)
            BControl.RightToLeft = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property ScrollCount() As Long
        Get
            ScrollCount = BControl.ScrollCount
        End Get
        Set(value As Long)
            'it is just get
        End Set
    End Property

    Property ScrollIndex() As Long
        Get
            ScrollIndex = BControl.ScrollIndex
        End Get
        Set(value As Long)
            BControl.ScrollIndex = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property SelCount() As Integer
        Get
            SelCount = BControl.SelCount
        End Get
        Set(value As Integer)

        End Set
    End Property

    Property SelLength() As Integer
        Get
            SelLength = BControl.SelLength
        End Get
        Set(value As Integer)
            BControl.SelLength = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property SelStart() As Integer
        Get
            SelStart = BControl.SelStart

        End Get
        Set(value As Integer)
            BControl.SelStart = value
            Call ControlSet(BControl)

        End Set
    End Property

    Property SelText() As String
        Get
            SelText = BControl.SelText
        End Get
        Set(value As String)
            BControl.SelText = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property Selected(lnteger As Integer) As Boolean
        Get
            Selected = Query(BControl.DataBase).Selected(lnteger)
        End Get
        Set(value As Boolean)
            Query(BControl.DataBase).Selected(lnteger) = value
        End Set
    End Property


    Property SmallChange() As Integer
        Get
            SmallChange = BControl.SmallChange
        End Get
        Set(value As Integer)
            BControl.SmallChange = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property Sorted() As Boolean
        Get
            Sorted = BControl.Sorted
        End Get
        Set(value As Boolean)
            BControl.Sorted = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property Style() As Integer
        Get
            Style = BControl.Style
        End Get
        Set(value As Integer)
            BControl.Style = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property System() As Boolean
        Get
            System = BControl.System
        End Get
        Set(value As Boolean)
            BControl.System = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property TabIndex() As Integer
        Get
            TabIndex = BControl.TabIndex
        End Get
        Set(value As Integer)
            Dim vA As Integer
            If value > -1 And value < DXForm(BFormIndex).TavalueIndexCount Then
                For vA = 0 To DXForm(BFormIndex).ControlsCount
                    If DXForm(BFormIndex).Controls(vA).TabStop = True Then
                        If DXForm(BFormIndex).Controls(vA).TabIndex < BControl.TabIndex And DXForm(BFormIndex).Controls(vA).TabIndex >= value Then
                            DXForm(BFormIndex).Controls(vA).TabIndex = DXForm(BFormIndex).Controls(vA).TabIndex + 1
                        End If
                    End If
                Next vA
            End If
            BControl.TabIndex = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property TabStop() As Boolean
        Get
            TabStop = BControl.TabStop
        End Get
        Set(value As Boolean)
            Dim vA As Integer
            If value = BControl.TabStop Then
                Exit Property
            ElseIf value = False Then
                DXForm(BFormIndex).TabIndexCount = DXForm(BFormIndex).TabIndexCount - 1
                For vA = 0 To DXForm(BFormIndex).ControlsCount
                    If DXForm(BFormIndex).Controls(vA).TabIndex > BControl.TabIndex Then
                        DXForm(BFormIndex).Controls(vA).TabIndex = DXForm(BFormIndex).Controls(vA).TabIndex - 1
                    End If
                Next vA
                BControl.TabIndex = -1
            Else
                If BControl.TabIndex = -1 Then
                    DXForm(BFormIndex).TabIndexCount = DXForm(BFormIndex).TabIndexCount + 1
                    BControl.TabIndex = DXForm(BFormIndex).TabIndexCount
                End If
            End If
            BControl.TabStop = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property Tag() As String
        Get
            Tag = BControl.Tag
        End Get
        Set(value As String)
            BControl.Tag = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property Text(Optional ListItem As Integer = 0) As String
        Get
            Text = BControl.Text(ListItem)
        End Get
        Set(value As String)

        End Set
    End Property


    Property ToolTipText() As String
        Get
            ToolTipText = BControl.ToolTipText
        End Get
        Set(value As String)
            BControl.ToolTipText = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property Top() As Single
        Get
            Top = BControl.Position(0).Top / FixHeight
        End Get
        Set(value As Single)
            Dim vA As Integer, vD As Integer
            BControl.Y1 = value * FixHeight
            vD = (value * FixHeight) - BControl.Position(0).Top
            For vA = 0 To BControl.PositionCount
                BControl.Position(vA).Top = BControl.Position(vA).Top + vD
                BControl.Position(vA).Bottom = BControl.Position(vA).Bottom + vD
            Next vA
            Call ControlSet(BControl)
        End Set
    End Property


    Property TopIndex() As Integer
        Get
            TopIndex = BControl.TopIndex
        End Get
        Set(value As Integer)
            BControl.TopIndex = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property Value() As Integer
        Get
            If BControl.Value = True Then
                Value = 1
            Else
                Value = 0
            End If
        End Get
        Set(value As Integer)
            If value = 1 Then
                BControl.Value = True
                BControl.Image(0) = 10
                If BControl.Text(0) = "Yes" Or BControl.Text(0) = "No" Or BControl.Text(0) = "" Then
                    BControl.Text(0) = "Yes"
                End If
            Else
                BControl.Value = False
                BControl.Image(0) = 9
                If BControl.Text(0) = "Yes" Or BControl.Text(0) = "No" Or BControl.Text(0) = "" Then
                    BControl.Text(0) = "No"
                End If
            End If
            Call ControlSet(BControl)
        End Set
    End Property


    Property Visible() As Boolean
        Get
            Visible = BControl.Visible
        End Get
        Set(value As Boolean)
            Dim vA As Integer
            BControl.Visible = value
            For vA = 1 To BControl.EmosGifCount
                BControl.EmosGif(vA).Visible = value
            Next vA
            Call ControlSet(BControl)
        End Set
    End Property


    Property WhatsThisHelpID() As Long
        Get
            WhatsThisHelpID = BControl.WhatsThisHelpID
        End Get
        Set(value As Long)
            BControl.WhatsThisHelpID = value
            Call ControlSet(BControl)
        End Set
    End Property


    Property Width() As Single
        Get
            Width = (BControl.Position(0).Right - BControl.Position(0).Left) / FixWidth
        End Get
        Set(value As Single)
            Dim vA As Integer, vD As Integer
            BControl.X2 = value * FixWidth
            vD = ((value * FixWidth) + BControl.Position(0).Left) - BControl.Position(0).Right
            For vA = 0 To BControl.PositionCount
                BControl.Position(0).Right = BControl.Position(0).Right + vD
            Next vA
            Call ControlSet(BControl)
        End Set
    End Property


    Property WordWrap() As Boolean
        Get
            WordWrap = BControl.WordWrap
        End Get
        Set(value As Boolean)
            BControl.WordWrap = value
            Call ControlSet(BControl)
        End Set
    End Property

    Property X1() As Single
        Get
            X1 = BControl.X1 / FixWidth
        End Get
        Set(value As Single)
            BControl.X1 = value * FixWidth
            Call ControlSet(BControl)
        End Set
    End Property


    Property X2() As Single
        Get
            X2 = BControl.X2 / FixWidth
        End Get
        Set(value As Single)
            BControl.X2 = value * FixWidth
            Call ControlSet(BControl)
        End Set
    End Property


    Property Y1() As Single
        Get
            Y1 = BControl.Y1 / FixHeight
        End Get
        Set(value As Single)
            BControl.Y1 = value * FixHeight
            Call ControlSet(BControl)
        End Set

    End Property

    Property Y2() As Single
        Get
            Y2 = BControl.Y2 / FixHeight
        End Get
        Set(value As Single)
            BControl.Y2 = value * FixHeight
            Call ControlSet(BControl)
        End Set

    End Property


    Sub ZOrder(Position As Integer)
        BControl.ZOrder = Position
        Call ControlSet(BControl)
    End Sub

End Class