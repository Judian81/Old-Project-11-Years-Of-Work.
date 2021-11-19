Module List

    'code totaly implented
    Structure DQuery
        Public ItemData() As Double
        Public List() As String
        Public ListCount As Double
        Public UndoNumber() As Double
        Public UndoCount As Double
    End Structure

    Structure JQuery
        Public DataBase As Integer
        Public ItemData() As Double
        Public List() As String
        Public ListCount As Double
        Public SelCount As Double
        Public Selected() As Boolean
        Public Sorted As Integer
        Public ListName As String
        Public Refresh As Integer
        Public ListChange As Integer
        Public ScrollChange As Integer
        Public SortStructure As Byte
        Public Record As Integer
        Public Field As Integer
        Public Path As String
    End Structure

    Public Query() As JQuery
    Public QueryUndelete() As DQuery
    Public QueryCount As Integer

    'the parts that are visible on the screen uses the info from the date base so there is not much more loaded into the screen
    Sub ComboListIndex(sList As String, iDataBase As Integer)
        Dim vA As Double
        For vA = 1 To Query(iDataBase).ListCount
            If sList = Query(iDataBase).List(vA) Then
                BComboList.ListIndex = vA
                Exit Sub
            End If
        Next vA
    End Sub

    'sorting all the records from one database
    Sub RecordSorting(zDataBase As Integer)
        Dim vA As Integer, vC As Integer, vD As Integer, vE As Integer, vF As Integer
        Dim sA() As String, iA() As Double, bA As Boolean
        ReDim Preserve sA(Query(zDataBase).ListCount)
        ReDim Preserve iA(Query(zDataBase).ListCount)
        If Query(zDataBase).ListCount = 0 Then
            Exit Sub
        End If
        If Query(zDataBase).List(0) = "" Then
            vC = 1
            vD = 1
            vE = 1
        Else
            vC = 0
            vD = 0
            vE = 0
        End If
        For vA = vE To Query(zDataBase).ListCount
            sA(vA) = Query(zDataBase).List(vA)
            iA(vA) = Query(zDataBase).ItemData(vA)
        Next vA
        Do Until vD > Query(zDataBase).ListCount
            For vA = vE To Query(zDataBase).ListCount
                If sA(vA) <> "" Then
                    If UCase$(sA(vC)) > UCase$(sA(vA)) Or sA(vC) = "" Then
                        vC = vA
                    End If
                End If
            Next vA
            Query(zDataBase).List(vD) = sA(vC)
            Query(zDataBase).ItemData(vD) = iA(vC)
            vD = (vD + 1)
            If vD > Query(zDataBase).ListCount Then
                Exit Do
            End If
            sA(vC) = ""
            iA(vC) = ""
            For vA = vE To Query(zDataBase).ListCount
                If sA(vA) <> "" Then
                    vC = vA
                    Exit For
                End If
            Next vA
            For vA = vE To Query(zDataBase).ListCount
                If sA(vA) <> "" Then
                    If UCase$(sA(vC)) > UCase$(sA(vA)) Then
                        Exit For
                    Else
                        vC = vA
                    End If
                End If
            Next vA
            For vA = vC To Query(zDataBase).ListCount
                If sA(vA) <> "" Then
                    vF = vA
                    Exit For
                End If
            Next vA
            For vA = (vC + 1) To Query(zDataBase).ListCount
                If sA(vA) <> "" Then
                    If UCase$(sA(vF)) > UCase$(sA(vA)) Then
                        vF = vA
                    End If
                End If
            Next vA
            For vA = vE To (vC + 1)
                If Query(zDataBase).ListCount >= vA Then
                    If sA(vA) <> "" Then
                        If UCase$(sA(vF)) > UCase$(sA(vA)) Then
                            Query(zDataBase).List(vD) = sA(vA)
                            Query(zDataBase).ItemData(vD) = iA(vA)
                            vD += 1
                            sA(vA) = ""
                            iA(vA) = ""
                            vE = vA
                        Else
                            Query(zDataBase).List(vD) = sA(vF)
                            Query(zDataBase).ItemData(vD) = iA(vF)
                            vD += 1
                            sA(vF) = ""
                            iA(vF) = ""
                            Exit For
                        End If
                    End If
                End If
            Next vA
            DoEvents
        Loop
        If Query(zDataBase).SortStructure = 2 Then
            ReDim Preserve sA(Query(zDataBase).ListCount)
            ReDim Preserve iA(Query(zDataBase).ListCount)
            For vA = 0 To Query(zDataBase).ListCount - 1
                sA(vA + 1) = Query(zDataBase).List(Query(zDataBase).ListCount - vA)
                iA(vA + 1) = Query(zDataBase).ItemData(Query(zDataBase).ListCount - vA)
            Next vA
            For vA = 0 To Query(zDataBase).ListCount
                Query(zDataBase).List(vA) = sA(vA)
                Query(zDataBase).ItemData(vA) = iA(vA)
            Next vA
        End If
    End Sub

    'this function search if the database exist. and what database is in that number
    Function DataNR(vZ As String, Optional vX As String = "") As Integer
        Dim vA As Integer, iB As Integer
        If Left$(vZ, 1) = "|" Then
            For vA = 0 To QueryCount
                If LCase$(vZ) = LCase$(Query(vA).ListName) Then
                    Return vA
                    Exit For
                End If
            Next vA
        End If
        Return 0
    End Function

    'with use of database number we could find what name the database represents
    Function DataName(vZ As Integer) As String
        Dim vA As Integer
        For vA = 0 To QueryCount
            If vZ = vA Then
                If Query(vA).ListName <> "" Then
                    'this is the name from the database
                    Return LCase$(Query(vA).ListName)
                End If
            End If
        Next vA
        'we could not find the database so we also does not get the name
        Return "we could not find the database so we also does not get the name"
    End Function

    Function QueryAdd(sDataBase As String) As String
        QueryAdd = DataNR(sDataBase)
        If QueryAdd = "0" Then
            QueryCount += 1
            ReDim Preserve Query(QueryCount)
            ReDim Preserve Query(QueryCount).List(0)
            ReDim Preserve Query(QueryCount).Selected(0)
            ReDim Preserve Query(QueryCount).ItemData(0)
            Query(QueryCount).DataBase = QueryCount
            Query(QueryCount).ListName = sDataBase
            ReDim Preserve QueryUndelete(QueryCount)
            ReDim Preserve QueryUndelete(QueryCount).List(0)
            ReDim Preserve QueryUndelete(QueryCount).UndoNumber(0)
            ReDim Preserve QueryUndelete(QueryCount).ItemData(0)
            Return QueryCount
        End If
        Return ""
    End Function

    Sub QueryAddItem(vDatabase As Integer, sList As String, Optional vItemData As Double = 0, Optional vSelected As Boolean = False, Optional vIndex As Integer = 0)
        Dim vQueryIndex As Double
        Query(vDatabase).ListCount = Query(vDatabase).ListCount + 1
        ReDim Preserve Query(vDatabase).List(Query(vDatabase).ListCount)
        ReDim Preserve Query(vDatabase).ItemData(Query(vDatabase).ListCount)
        ReDim Preserve Query(vDatabase).Selected(Query(vDatabase).ListCount)
        If vIndex < 1 Or vIndex > Query(vDatabase).ListCount Then
            Query(vDatabase).List(Query(vDatabase).ListCount) = sList
            Query(vDatabase).ItemData(Query(vDatabase).ListCount) = vItemData
            Query(vDatabase).Selected(Query(vDatabase).ListCount) = vSelected
        Else
            For vQueryIndex = Query(vDatabase).ListCount - 1 To vIndex + 1 Step -1
                Query(vDatabase).List(vQueryIndex) = Query(vDatabase).List(vQueryIndex - 1)
                Query(vDatabase).ItemData(vQueryIndex) = Query(vDatabase).ItemData(vQueryIndex - 1)
                Query(vDatabase).Selected(vQueryIndex) = Query(vDatabase).Selected(vQueryIndex - 1)
            Next vQueryIndex
            Query(vDatabase).List(vIndex) = sList
            Query(vDatabase).ItemData(vIndex) = vItemData
            Query(vDatabase).Selected(vIndex) = vSelected
        End If
    End Sub

    Sub QueryClear(vDatabase As Integer)
        ReDim Preserve Query(vDatabase).List(0)
        ReDim Preserve Query(vDatabase).ItemData(0)
        ReDim Preserve Query(vDatabase).Selected(0)
        Query(vDatabase).ListCount = 0
        Query(vDatabase).SelCount = 0
    End Sub

    Sub QueryRemoveSelected(zDataBase As Integer, Optional bMessage As Boolean = True)
        Dim iItem As Double, iItemReplace As Double
        If Query(zDataBase).SelCount <> "" Then
            QueryUndelete(zDataBase).UndoCount = QueryUndelete(zDataBase).UndoCount + 1
            For iItem = Query(zDataBase).ListCount To 1 Step -1
                Select Case ""
                    Case Query(zDataBase).Selected(iItem), Query(zDataBase).List(iItem)
                    Case Else
                        If Query(zDataBase).Selected(iItem) = True Then
                            Query(zDataBase).SelCount = Query(zDataBase).SelCount - 1
                        End If
                        Call QueryUndoAddItem(zDataBase, Query(zDataBase).List(iItem), Query(zDataBase).ItemData(iItem))
                        For iItemReplace = iItem To Query(zDataBase).ListCount - 1
                            Query(zDataBase).List(iItemReplace) = Query(zDataBase).List(iItemReplace + 1)
                            Query(zDataBase).ItemData(iItemReplace) = Query(zDataBase).ItemData(iItemReplace + 1)
                            Query(zDataBase).Selected(iItemReplace) = Query(zDataBase).Selected(iItemReplace + 1)
                        Next iItemReplace
                        Query(zDataBase).ListCount = Query(zDataBase).ListCount - 1
                        ReDim Preserve Query(zDataBase).List(Query(zDataBase).ListCount)
                        ReDim Preserve Query(zDataBase).ItemData(Query(zDataBase).ListCount)
                        ReDim Preserve Query(zDataBase).Selected(Query(zDataBase).ListCount)
                End Select
            Next iItem
            If DXForm(BFormIndex).Controls(iFocus).ListIndex > Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount Then
                DXForm(BFormIndex).Controls(iFocus).ListIndex = Query(DXForm(BFormIndex).Controls(iFocus).DataBase).ListCount
            End If
        End If
    End Sub

    Sub QueryUndoAddItem(vDatabase As Integer, sList As String, vItemData As Double)
        QueryUndelete(vDatabase).ListCount = QueryUndelete(vDatabase).ListCount + 1
        ReDim Preserve QueryUndelete(vDatabase).List(QueryUndelete(vDatabase).ListCount)
        ReDim Preserve QueryUndelete(vDatabase).ItemData(QueryUndelete(vDatabase).ListCount)
        ReDim Preserve QueryUndelete(vDatabase).UndoNumber(QueryUndelete(vDatabase).ListCount)
        QueryUndelete(vDatabase).List(QueryUndelete(vDatabase).ListCount) = sList
        QueryUndelete(vDatabase).ItemData(QueryUndelete(vDatabase).ListCount) = vItemData
        QueryUndelete(vDatabase).UndoNumber(QueryUndelete(vDatabase).ListCount) = QueryUndelete(vDatabase).UndoCount
    End Sub

    Sub QueryUndoRemove(vDatabase As Integer)
        Dim iItem As Double, iItemUndo As Double
        If 0 < QueryUndelete(vDatabase).UndoCount Then
            For iItem = QueryUndelete(vDatabase).ListCount To 1 Step -1
                If QueryUndelete(vDatabase).UndoNumber(iItem) = QueryUndelete(vDatabase).UndoCount Then
                    Query(vDatabase).ListCount = Query(vDatabase).ListCount + 1
                    ReDim Preserve Query(vDatabase).List(Query(vDatabase).ListCount)
                    ReDim Preserve Query(vDatabase).ItemData(Query(vDatabase).ListCount)
                    ReDim Preserve Query(vDatabase).Selected(Query(vDatabase).ListCount)
                    Query(vDatabase).List(Query(vDatabase).ListCount) = QueryUndelete(vDatabase).List(iItem)
                    Query(vDatabase).ItemData(Query(vDatabase).ListCount) = QueryUndelete(vDatabase).ItemData(iItem)
                    Query(vDatabase).Selected(Query(vDatabase).ListCount) = 0
                    For iItemUndo = iItem To QueryUndelete(vDatabase).ListCount - 1
                        QueryUndelete(vDatabase).List(iItemUndo) = QueryUndelete(vDatabase).List(iItemUndo + 1)
                        QueryUndelete(vDatabase).ItemData(iItemUndo) = QueryUndelete(vDatabase).ItemData(iItemUndo + 1)
                        QueryUndelete(vDatabase).UndoNumber(iItemUndo) = QueryUndelete(vDatabase).UndoNumber(iItemUndo + 1)
                    Next iItemUndo
                    QueryUndelete(vDatabase).ListCount = QueryUndelete(vDatabase).ListCount - 1
                    ReDim Preserve QueryUndelete(vDatabase).List(QueryUndelete(vDatabase).ListCount)
                    ReDim Preserve QueryUndelete(vDatabase).ItemData(QueryUndelete(vDatabase).ListCount)
                    ReDim Preserve QueryUndelete(vDatabase).UndoNumber(QueryUndelete(vDatabase).ListCount)
                End If
            Next iItem
            QueryUndelete(vDatabase).UndoCount = QueryUndelete(vDatabase).UndoCount - 1
            Call RecordSorting(vDatabase)
        End If
    End Sub

    Private Sub SetQuery()
        ReDim Query(0)
        ReDim Query(0).List(0)
        ReDim Query(0).ItemData(0)
        ReDim Query(0).Selected(0)
        Query(0).ListCount = 0
        Query(0).SelCount = 0
        ReDim QueryUndelete(0)
        ReDim QueryUndelete(0).List(0)
        ReDim QueryUndelete(0).ItemData(0)
        ReDim QueryUndelete(0).UndoNumber(0)
        QueryUndelete(0).ListCount = 0
    End Sub

    'some thing does not working. i do not see yet where it is used for. have a good day :)
    Sub LoadQuery()
        Call SetQuery()
        Call Create(6, "MaxNr", , 2, , , "1")
        Call Create(6, "MaxRows", 1, , 1)
        Call Create(1999, "MiliSecond", 4, , , "000")
        Call Color()
        Call Hour()
        Call Month()
        Call Minute()
        Call Calculate()
        Call Skins()
        Call FoodManager()
        Call FormsActive()
        Call GTM()
        Call MediaPlayer()
        Call Agenda()
        Call Programs()
        '''''''Call Operator
        Call Special()
        Call TimeList()
        Call RecordSelection()
        Call Attachment()
    End Sub

    Sub Add(sListName As String)
        Call QueryAdd("|" & sListName)
    End Sub

    Private Sub Create(vTo As Integer, sListName As String, Optional vTextLength As Integer = 0, Optional vOption As Byte = 0, Optional vFrom As Integer = 0, Optional sTextFill As String = "", Optional sStartText As String = "")
        Dim vNumber As Integer
        Call Add(sListName)
        Select Case vOption
            Case 0
                For vNumber = vFrom To vTo
                    Call QueryAddItem(QueryCount, Right(sTextFill & vNumber, vTextLength))
                Next vNumber
            Case 1
                For vNumber = vFrom To vTo
                    Call QueryAddItem(QueryCount, sStartText + vNumber)
                Next vNumber
            Case 2
                For vNumber = vFrom To vTo
                    Call QueryAddItem(QueryCount, CStr(sStartText))
                    sStartText = sStartText & "0"
                Next vNumber
        End Select
    End Sub

    Private Sub Attachment()
        Call Add("Attachment")
    End Sub

    Private Sub Calculate()
        Call Add("Calculate")
        Call QueryAddItem(QueryCount, "+")
        Call QueryAddItem(QueryCount, "-")
        Call QueryAddItem(QueryCount, "*")
        Call QueryAddItem(QueryCount, "/")
        Call QueryAddItem(QueryCount, "%")
        Call QueryAddItem(QueryCount, "^")
        Call QueryAddItem(QueryCount, "v")
    End Sub

    Private Sub Color()
        Call Add("Color")
        Call QueryAddItem(QueryCount, "Red")
        Call QueryAddItem(QueryCount, "Pink")
        Call QueryAddItem(QueryCount, "LightGreen")
        Call QueryAddItem(QueryCount, "Yellow")
        Call QueryAddItem(QueryCount, "White")
        Call QueryAddItem(QueryCount, "Black")
        Call QueryAddItem(QueryCount, "Blue / Green")
        Call QueryAddItem(QueryCount, "Blue")
    End Sub

    Private Sub FoodManager()
        Call Add("Food Manager")
        Call QueryAddItem(QueryCount, "Food All")
        Call QueryAddItem(QueryCount, "Food Products")
        Call QueryAddItem(QueryCount, "Food Stok")
    End Sub

    Private Sub FormsActive()
        Call Add("Forms Active")
    End Sub

    Private Sub GTM()
        Dim vHour As Byte
        Call Add("GTM")
        For vHour = 1 To 12
            Call QueryAddItem(QueryCount, "-" & Right("0" & (13 - vHour), 2))
        Next vHour
        Call QueryAddItem(QueryCount, "+00")
        For vHour = 1 To 12
            Call QueryAddItem(QueryCount, "+" & Right("0" & vHour, 2))
        Next vHour
    End Sub

    Private Sub Hour()
        Dim vHour As Byte
        Call Add("Hour")
        For vHour = 0 To 23
            Call QueryAddItem(QueryCount, Right("0" & vHour, 2))
        Next vHour
    End Sub

    Private Sub Month()
        Dim vMonth As Byte
        Call Add("Month")
        For vMonth = 1 To 12
            Call QueryAddItem(QueryCount, Right("0" & vMonth, 2))
        Next vMonth
    End Sub

    Private Sub Minute()
        Dim vMinute As Byte
        Call Add("Minute")
        For vMinute = 0 To 59
            Call QueryAddItem(QueryCount, Right("0" & vMinute, 2))
        Next vMinute
    End Sub

    Private Sub MediaPlayer()
        Call Add("Media Player")
        Call QueryAddItem(QueryCount, "DVD")
        Call QueryAddItem(QueryCount, "Movie")
        Call QueryAddItem(QueryCount, "Albums")
        Call QueryAddItem(QueryCount, "Play List")
        Call QueryAddItem(QueryCount, "Television")
        Call QueryAddItem(QueryCount, "Series")
    End Sub

    Private Sub Agenda()
        Call Add("Agenda")
        Call QueryAddItem(QueryCount, "History")
        Call QueryAddItem(QueryCount, "Inbox")
        Call QueryAddItem(QueryCount, "Outbox")
    End Sub

    Private Sub Skins()
        Call Add("Skins")
    End Sub

    Private Sub Programs()
        Call Add("Programs")
        Call QueryAddItem(QueryCount, "Controls")
        Query(QueryCount).Sorted = True
        Call RecordSorting(CInt(QueryCount))
    End Sub

    Private Sub RecordSelection()
        Call Add("Record Selection")
    End Sub

    Private Sub Special()
        Call Add("From/For/With")
        Call QueryAddItem(QueryCount, "From")
        Call QueryAddItem(QueryCount, "For")
        Call QueryAddItem(QueryCount, "With")
    End Sub

    ' i have to rename Operator to VOperator, yeah still awake. bleh
    Private Sub VOperator()
        Call Add("Operator")
        Call QueryAddItem(QueryCount, "=")
        Call QueryAddItem(QueryCount, "<>")
        Call QueryAddItem(QueryCount, ">")
        Call QueryAddItem(QueryCount, "<")
        Call QueryAddItem(QueryCount, ">=")
        Call QueryAddItem(QueryCount, "=<")
    End Sub

    Private Sub TimeList()
        Dim vHour As Byte, vMinute As Byte, sTime As String
        Call Add("Time")
        For vHour = 0 To 23
            For vMinute = 0 To 59
                sTime = vHour & ":" & vMinute
                If Len(sTime) = 3 Then
                    sTime = "0" & Left(sTime, 1) & ":0" & Right(sTime, 1)
                End If
                If Len(sTime) = 4 And Mid(sTime, 2, 1) = ":" Then
                    sTime = "0" & sTime
                ElseIf Len(sTime) = 4 And Mid(sTime, 3, 1) = ":" Then
                    sTime = Left(sTime, 3) & "0" & Right(sTime, 1)
                End If
                Call QueryAddItem(QueryCount, sTime)
            Next vMinute
        Next vHour
    End Sub

End Module
