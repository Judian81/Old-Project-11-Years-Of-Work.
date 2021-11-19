Public Class AForm

    'all code implented
    Private Linking As New cControl

    Property Timer1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Timer1"
            BControlIndex = Index
            BControl = ControlGet("Timer")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property


    Property BorderH(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "BorderH"
            BControlIndex = Index
            BControl = ControlGet("BorderH")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property BorderV(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "BorderV"
            BControlIndex = Index
            BControl = ControlGet("BorderV")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property Check1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Check1"
            BControlIndex = Index
            BControl = ControlGet("Check1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property Combo1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Combo1"
            BControlIndex = Index
            BControl = ControlGet("Combo1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property ComboList(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "ComboList"
            BControlIndex = Index
            BControl = BComboList
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property Command1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Command1"
            BControlIndex = Index
            BControl = ControlGet("Command1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property Data1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Data1"
            BControlIndex = Index
            BControl = ControlGet("Data1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property Frame1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Frame1"
            BControlIndex = Index
            BControl = ControlGet("Frame1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property Image1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Image1"
            BControlIndex = Index
            BControl = ControlGet("Image1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property Label1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Label1"
            BControlIndex = Index
            BControl = ControlGet("Label1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property Line1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Line1"
            BControlIndex = Index
            BControl = ControlGet("Line1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property List1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "List1"
            BControlIndex = Index
            BControl = ControlGet("List1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property Option1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Option1"
            BControlIndex = Index
            BControl = ControlGet("Option1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property Picture1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Picture1"
            BControlIndex = Index
            BControl = ControlGet("Picture1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property HScroll1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "HScroll1"
            BControlIndex = Index
            BControl = ControlGet("HScroll1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property VScroll1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "VScroll1"
            BControlIndex = Index
            BControl = ControlGet("VScroll1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property Shape1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Shape1"
            BControlIndex = Index
            BControl = ControlGet("Shape1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property Text1(Index As Integer) As cControl
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Text1"
            BControlIndex = Index
            BControl = ControlGet("Text1")
            Return Linking
        End Get
        Set(value As cControl)

        End Set
    End Property

    Property ActiveControl() As Control
        'do nothing
        Get

        End Get
        Set(value As Control)

        End Set
    End Property

    Property Appearance() As Integer
        Get
            Return DXForm(BFormIndex).Appearance
        End Get
        Set(value As Integer)
            DXForm(BFormIndex).Appearance = value
        End Set
    End Property

    Property AutoRedraw() As Boolean
        'do nothing
        Get
            Return False
        End Get
        Set(value As Boolean)

        End Set
    End Property

    Property BackColor() As Long
        Get
            Return DXForm(BFormIndex).BackColor
        End Get
        Set(value As Long)
            DXForm(BFormIndex).BackColor = value
        End Set
    End Property


    Property BorderStyle() As Integer
        Get
            Return DXForm(BFormIndex).BorderStyle
        End Get
        Set(value As Integer)
            DXForm(BFormIndex).BorderStyle = value
        End Set
    End Property

    Property Caption() As String
        Get
            Return DXForm(BFormIndex).Caption
        End Get
        Set(value As String)
            DXForm(BFormIndex).Caption = value
        End Set

    End Property


    Property Controls(Index As Integer) As Object
        Get
            DXForm(BFormIndex).ControlIndex = Index
            BControlType = "Controls"
            BControl = ControlGet("Controls")
            Return Linking
        End Get
        Set(value As Object)

        End Set
    End Property

    Property Count() As Integer
        Get
            Return DXForm(BFormIndex).ControlsCount
        End Get
        Set(value As Integer)

        End Set
    End Property

    Property Enabled() As Boolean
        Get
            Return DXForm(BFormIndex).Enabled
        End Get
        Set(value As Boolean)
            DXForm(BFormIndex).Enabled = value
        End Set

    End Property


    Property FillStyle() As Integer
        Get
            Return DXForm(BFormIndex).FillStyle
        End Get
        Set(value As Integer)
            DXForm(BFormIndex).FillStyle = value
        End Set
    End Property



    Property Font() As IFontDisp
        Get
            Dim vFont As stdFont
            vFont = New StdFont
            vFont.Bold = DXForm(BFormIndex).Font.Bold
            vFont.Char = DXForm(BFormIndex).Font.Charset
            vFont.Italic = DXForm(BFormIndex).Font.Italic
            If DXForm(BFormIndex).Font.Name <> "" Then
                vFont.Name = DXForm(BFormIndex).Font.Name
            End If
            If DXForm(BFormIndex).Font.Size <> 0 Then
                vFont.Size = DXForm(BFormIndex).Font.Size
            End If
            vFont.Strikethrough = DXForm(BFormIndex).Font.Strikethrough
            vFont.Underline = DXForm(BFormIndex).Font.Underline
            vFont.Weight = DXForm(BFormIndex).Font.Weight
            Return vFont
        End Get
        Set(value As IFontDisp)
            ''''vFont = New StdFont do not need this
            DXForm(BFormIndex).Font.Bold = value.Bold
            DXForm(BFormIndex).Font.Char = value.Charset
            DXForm(BFormIndex).Font.Italic = value.Italic
            DXForm(BFormIndex).Font.Name = value.Name
            DXForm(BFormIndex).Font.Size = value.Size
            DXForm(BFormIndex).Font.Strikethrough = value.FontStrikethrough
            DXForm(BFormIndex).Font.Underline = value.Underline
            DXForm(BFormIndex).Font.Weight = value.Weight
        End Set
    End Property

    Property FontBold() As Boolean
        Get
            Return DXForm(BFormIndex).Font.Bold
        End Get
        Set(value As Boolean)
            DXForm(BFormIndex).Font.Bold = value
        End Set
    End Property

    Property FontItalic() As Boolean
        Get
            Return DXForm(BFormIndex).Font.Italic
        End Get
        Set(value As Boolean)
            DXForm(BFormIndex).Font.Italic = value
        End Set
    End Property


    Property FontName() As String
        Get
            Return DXForm(BFormIndex).Font.Name
        End Get
        Set(value As String)
            DXForm(BFormIndex).Font.Name = value
        End Set
    End Property

    Property FontSize() As Single
        Get
            Return DXForm(BFormIndex).Font.Size
        End Get
        Set(value As Single)
            DXForm(BFormIndex).Font.Size = value
        End Set
    End Property

    Property FontStrikethrough() As Boolean
        Get
            Return DXForm(BFormIndex).Font.Strikethrough
        End Get
        Set(value As Boolean)
            DXForm(BFormIndex).Font.Strikethrough = value
        End Set
    End Property


    Property FontUnderline() As Boolean
        Get
            Return DXForm(BFormIndex).Font.Underline
        End Get
        Set(value As Boolean)
            DXForm(BFormIndex).Font.Underline = value
        End Set
    End Property

    Property ForeColor() As Long
        Get
            Return DXForm(BFormIndex).ForeColor
        End Get
        Set(value As Long)
            DXForm(BFormIndex).ForeColor = value
        End Set
    End Property


    Property Height() As Single
        Get
            Return DXForm(BFormIndex).Height / FixHeight
        End Get
        Set(value As Single)
            DXForm(BFormIndex).Height = value * FixHeight
        End Set
    End Property


    Property HelpContextID() As Long
        Get
            HelpContextID = DXForm(BFormIndex).HelpContextID
        End Get
        Set(value As Long)
            DXForm(BFormIndex).HelpContextID = value
        End Set
    End Property


    Property Image() As IPictureDisp
        Get
            Return DXForm(BFormIndex).Image
        End Get
        Set(value As IPictureDisp)
            DXForm(BFormIndex).Image = value
        End Set
    End Property

    Property Left() As Single
        Get
            Return DXForm(BFormIndex).Left / FixWidth
        End Get
        Set(value As Single)
            DXForm(BFormIndex).Left = value * FixWidth
        End Set
    End Property

    Property MouseIcon() As IPictureDisp
        Get
            Return DXForm(BFormIndex).MouseIcon
        End Get
        Set(value As IPictureDisp)
            DXForm(BFormIndex).MouseIcon = value
        End Set
    End Property


    Property MousePointer() As Integer
        Get
            MousePointer = DXForm(BFormIndex).MousePointer
        End Get
        Set(value As Integer)
            DXForm(BFormIndex).MousePointer = value
        End Set
    End Property


    Sub Move(Left As Single, Optional Top As Single = -1, Optional Width As Single = -1, Optional Height As Single = -1)
        DXForm(BFormIndex).Left = Left * FixWidth
        If Top <> -1 Then
            DXForm(BFormIndex).Top = Top * FixHeight
        End If
        If Width <> -1 Then
            DXForm(BFormIndex).Width = Width * FixWidth
        End If
        If Height <> -1 Then
            DXForm(BFormIndex).Height = Height * FixHeight
        End If
    End Sub

    Property Name() As String
        Get
            Return DXForm(BFormIndex).Name
        End Get
        Set(value As String)

        End Set
    End Property

    Property RightToLeft() As Boolean
        Get
            Return DXForm(BFormIndex).RightToLeft
        End Get
        Set(value As Boolean)
            DXForm(BFormIndex).RightToLeft = value
        End Set
    End Property

    Sub SetFocus()
        DXForm(BFormIndex).SetFocus = True
    End Sub

    Sub Show() 'Optional Modal As Byte = False) ', Optional OwnerForm As Form = Form)
        DXForm(BFormIndex).Visible = True
    End Sub

    Property Tag() As String
        Get
            Return DXForm(BFormIndex).Tag
        End Get
        Set(value As String)
            DXForm(BFormIndex).Tag = value
        End Set
    End Property

    Property Top() As Single
        Get
            Return DXForm(BFormIndex).Top / FixHeight
        End Get
        Set(value As Single)
            DXForm(BFormIndex).Top = value * FixHeight
        End Set
    End Property


    Property Visible() As Boolean
        Get
            Return DXForm(BFormIndex).Visible
        End Get
        Set(value As Boolean)
            DXForm(BFormIndex).Visible = value
        End Set
    End Property


    Property WhatsThisHelpID() As Long
        Get
            Return DXForm(BFormIndex).WhatsThisHelpID
        End Get
        Set(value As Long)
            DXForm(BFormIndex).WhatsThisHelpID = value
        End Set
    End Property


    Property Width() As Single
        Get
            Return DXForm(BFormIndex).Width / FixWidth
        End Get
        Set(value As Single)
            DXForm(BFormIndex).Width = value * FixWidth
        End Set
    End Property

    Property WindowState() As Integer
        Get
            Return DXForm(BFormIndex).WindowState
        End Get
        Set(value As Integer)
            DXForm(BFormIndex).WindowState = value
        End Set
    End Property


    Sub ZOrder(value As Integer)
        DXForm(BFormIndex).ZOrder = value
    End Sub

End Class
