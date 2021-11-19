Module mString

    '--------------------------THIS HAS TO DO WITH TIME------------------------------------
    'all code inplented
    'i know what it is doing. but why is the big question
    Function TimeCompile(vZ As String) As String
        'it looks at the minutes. could also be secondes. every quarter hour it change to something that could be 1/4 of 100. stead of 1/4 of 60
        Select Case CStr(Right$(vZ, 2))
            Case "15", "25"
                Return Left$(vZ, Len(vZ) - 2) & "25"
            Case "30", "50"
                Return Left$(vZ, Len(vZ) - 2) & "50"
            Case "45", "75"
                Return Left$(vZ, Len(vZ) - 2) & "75"
            Case "00"
                Return vZ
            Case Else
                'realy in hours you have a "," ???? this is very weard ?!?!?! is it not ????
                Return "0,00"
        End Select
    End Function

    'i do not know if this is calculated correctly
    Function SunSign(vZ As Integer) As String
        'the numbers are ??? wauw that is a very good question ???
        If vZ >= 120 And vZ < 219 Then
            Return "Waterman"
        ElseIf vZ >= 219 And vZ < 320 Then
            Return "Vissen"
        ElseIf vZ >= 320 And vZ < 420 Then
            Return "Ram"
        ElseIf vZ >= 420 And vZ < 521 Then
            Return "Stier"
        ElseIf vZ >= 521 And vZ < 621 Then
            Return "Tweelingen"
        ElseIf vZ >= 621 And vZ < 723 Then
            Return "Kreeft"
        ElseIf vZ >= 723 And vZ < 823 Then
            Return "Leeuw"
        ElseIf vZ >= 823 And vZ < 923 Then
            Return "Maagd"
        ElseIf vZ >= 923 And vZ < 1023 Then
            Return "Weegschaal"
        ElseIf vZ >= 1023 And vZ < 1122 Then
            Return "Schorpioen"
        ElseIf vZ >= 1122 And vZ < 1222 Then
            Return "Boogschutter"
        ElseIf (vZ >= 1222 And vZ <= 1231) Or (vZ < 120 And vZ > 11) Then
            Return "Steenbok"
        Else
            Return "?"
        End If
    End Function

    'first look what this does
    'vZ minus vX is the difference in time
    Function TimeMin(vZ As String, vX As String) As String
        Dim vA As Integer, iB As Integer, vC As Integer
        If Len(vZ) = 5 Or Len(vX) = 5 Then
            vA = 0
        Else
            vA = Val(Right$(vZ, 2)) - Val(Right$(vX, 2))
        End If
        If vA < 0 Then
            vA = 60 + vA
            iB = -1
        End If
        iB = Val(Mid$(vZ, 4, 2)) - Val(Mid$(vX, 4, 2)) + iB
        If iB < 0 Then
            iB = 60 + iB
            vC = -1
        End If
        vC = Val(Left$(vZ, 2)) - Val(Left$(vX, 2)) + vC
        If vC < 0 Then
            vC = 24 - vC
        End If
        Return Right$("0" & vC, 2) & ":" & Right$("0" & iB, 2) & ":" & Right$("0" & vA, 2)
    End Function

    'how can you win if it is in all of us
    'to go on. the two time's given in this function are added to eachother
    Function TimePlus(vZ As String, vX As String, Optional bZ As Boolean = False) As String
        Dim vA As Integer, vC As Integer
        vA = Val(Mid$(vZ, 4, 2)) + Val(Mid$(vX, 4, 2))
        If vA >= 60 Then
            vA = vA - 60
            vC = 1
        End If
        vC = Val(Left$(vZ, 2)) + Val(Left$(vX, 2)) + vC
        If bZ = False Then
            If vC > 23 Then
                vC = vC - 24
                If vC = 0 And Right(vZ, 2) = "PM" Then
                    vC = 12
                End If
            End If
        End If
        Return Right$("0" & vC, 2) & ":" & Right$("0" & vA, 2)
    End Function

    'i feel relaxt now. i do programming. but i take rest in the whole thing
    'i do not know realy why i have to use it. but i look if the minutes in total are more then 24 hours. if so minus a day in minutes and when then stop
    Function TimeMinutes2HourFix(iB As Integer) As Integer
        Dim iA As Integer
        iA = iB
        If iA < 0 Then
            Do Until iA > 0
                'add up minutes until it is positive and less then 24 hours in minutes
                iA += 1440
            Loop
            Return iA
        ElseIf iA > 1440 Then
            Do Until iA <= 1440
                'minus a day in minutes from the total minutes
                iA -= 1440
            Loop
            Return iA
        Else
            Return iA
        End If
    End Function

    ' i am feeling not so well. i think i have a depression

    'to calculatie with dates. i realy do not know what this is for??? I HAVE TO COME BACK LATER
    'if i do not know what script functions are i can not find some anwsers do i??? look back when script code is there
    'what is "e" for? "Daily" ? "," ?? who knows ??!!
    Function DateCalc(vScript As String, Optional vDate As String = "") As String
        Dim vA As Integer, bA As Boolean, vC As Integer, vD As String, vE As Byte, sDate As String
        sDate = ""
        If vScript = "Daily" Then
            Return Date.Today
        ElseIf vDate = "" Then
            sDate = Date.Today
        End If
        For vA = 0 To UBound(Split(vScript, ",")) - 1
            If bA = True Then
                If Right(Split(vScript, ",")(vA), 1) = "e" Then
                    vC = 1
                    For vE = 0 To 32
                        vD = DateAdd("d", vE, DateSerial(DatePart("yyyy", sDate), DatePart("mm", sDate), 1))
                        If Weekday(vD, vbSunday) = 1 Then
                            vC += 1
                        End If
                        If vC = Val(Left(Split(vScript, ",")(vA), 1)) And sDate = vD Then
                            Return sDate
                        End If
                    Next vE
                End If
            ElseIf Split(vScript, ",")(vA) = CStr(Weekday(Date.Today)) Then
                bA = True
            End If
        Next vA
        Return "what will become of this"
    End Function

    'zzzzzzzzzz. yawn. i have to relax for a bit. a shower helpt me much, celebrating shower day. yeahhh. hmm did it again. en while i was there i thought about code solving
    Function DateSort(vDate As String, Optional bDutch As Boolean = False) As String
        Dim vC As String
        vC = vDate
        If bDutch = True Then
            'this transfer the dutch date to a date you can use for sorting things out
            '10-20-2020
            'Month Day year
            Return Right(vC, 4) & "-" & Left(vC, 2) & "-" & Mid(vC, 4, 2)
        Else
            'this is when you give an amerikan date
            '20-10-2020
            'Day Month year
            Return Right(vC, 4) & "-" & Mid(vC, 4, 2) & "-" & Left(vC, 2)
        End If
    End Function

    'this is to get the number of weeks we are in a year. this code is simple and can do much better. oke it is much better because it calculate back to the 1800's
    Function WeekNumber(vDate As String) As String
        'to tired to go on. now trying to sleep. good bye. huggels
        'removed 100 lines to 2 lines. yay better now?
        Return DatePart("ww", CDate(vDate))
    End Function

    'look for the time and discoverd if it is morning or night. 24 time stamp is used
    Function TimeName24clock(vZ As Byte) As String
        Dim sTimeName As String
        sTimeName = ""
        'it is like what it is, use only hours
        If vZ >= 0 And vZ < 6 Then
            sTimeName = "night"
        ElseIf vZ >= 6 And vZ < 12 Then
            sTimeName = "morning"
        ElseIf vZ >= 12 And vZ < 18 Then
            sTimeName = "afternoon"
        ElseIf vZ >= 18 And vZ < 24 Then
            sTimeName = "evening"
        End If
        Return sTimeName
    End Function

    'going to search for total hours. the minutes left will be added to the time string
    Function TimeMinutes2Hour(iA As Integer) As String
        Dim iB As Integer, iC As Integer
        iB = iA
        'make hours by counting max amount of minutes there could be
        Do Until iB < 60
            'minutes left
            iB -= -60
            'hours left
            iC += +1
        Loop
        'put the magic together
        Return Right("0" & iC, 2) & ":" & Right("0" & iB, 2)
    End Function

    'calculate age by the given date like year-month-day like this 2000-12-31
    Function Age(vDate As String) As String
        Dim vA As Integer
        Dim OrigionalDate As String
        'first year, then month, then day
        OrigionalDate = Mid(Today.ToString, 7, 4) & "-" & Left(Today.ToString, 2) & "-" & Mid(Today.ToString, 4, 2)

        If Len(vDate) = 10 Then
            vA = Left(OrigionalDate, 4) - Left(vDate, 4)
            'when the month is bigger then the month that is your bird day
            If Mid(OrigionalDate, 5, 2) > Mid(vDate, 5, 2) Then
                Return vA
                'in this case i would like to know it is the same month as your bird day in it
            ElseIf Mid(OrigionalDate, 5, 2) = Mid(vDate, 5, 2) Then
                'now we only have to take a look if it is your bird day yet calculated by day of month
                If Right(OrigionalDate, 2) >= Right(vDate, 2) Then
                    'it is your bird day
                    Return vA
                Else
                    'sorry your bird day is still not there. nou you only have to do is count the night you have to sleep
                    Return (vA - 1)
                End If
            Else
                'you bird day is not there yet
                Return (vA - 1)
            End If
        Else
            'help someone to realize what is the problem
            Return "ERROR: sting has to be like 2000-12-31, year-month-day"
        End If
    End Function

    'this is if you have the date of anytime, you can ask for what day it is, it must be given in year-month-day
    Function DayName(vDate As String) As String
        'the given date is year-month-day
        Dim daterearange As String
        'make the date compatible month-day-year
        daterearange = Mid(vDate, 6, 2) & "/" & Right(vDate, 2) & "/" & Left(vDate, 4)
        'ask kindly what day of the week it is
        DayName = WeekdayName(Weekday(daterearange, vbSunday), False, vbSunday)
    End Function

    'this function wants to know that it realy is a string with time, this is about the hours and minutes. so secondes does not come in this place
    Function TimeCheck(vA As String) As Boolean
        'this is trying to discover it is exactly needed to see it is about time
        If Len(vA) = 5 And Mid(vA, 3, 1) = ":" And Left(vA, 2) < 24 And Left(vA, 2) >= 0 And Right(vA, 2) < 60 And Right(vA, 2) >= 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    'if you bring "01:01" to the table. this function is going to count hoy many minutes it is
    Function TimeHour2Minutes(sA As String) As String
        'first change hours to minutes and then add the minutes it already had
        Return Val(Left(sA, 2)) * 60 + Val(Right(sA, 2))
    End Function

    '-----------------------------------THIS HAS TO DO WITH FILES AND DIRECTORY'S--------------------------------------------------------------

    'removes a folder if it does exist
    Sub DirRemove(sPath As String)
        If DirReal(sPath) = True Then
            Call RmDir(sPath)
        End If
    End Sub

    'i you give a path name without a file it takes care it ends with a "\"
    Function DirReal(sZ As String) As Boolean
        Dim MyName As String, cPath As String, vA As Integer, vToMake As String
        'the path needs to be closed as a "\"
        cPath = Path(sZ)
        'now we remove the "\", in this order it will be right when you close or not close with a "\"
        cPath = Left(cPath, Len(cPath) - 1)
        vToMake = cPath
        For vA = Len(cPath) - 1 To 1 Step -1
            If Mid(cPath, vA, 1) = "\" Then
                cPath = Left(vToMake, vA)
                vToMake = Right(vToMake, Len(vToMake) - vA)
                Exit For
            End If
        Next
        MyName = Dir(cPath, vbDirectory)   ' Retrieve the first entry.
        Do While MyName <> ""   ' Start the loop.
            ' Use bitwise comparison to make sure MyName is a directory.
            If (GetAttr(cPath & MyName) And vbDirectory) = vbDirectory Then
                ' Display entry only if it's a directory.
                If MyName = vToMake Then
                    Return True
                End If
            End If
            MyName = Dir()   ' Get next entry.
        Loop
        Return False
    End Function

    'this function makes a new directory if it does not exist
    Function DirMake(sPath As String, Optional bInfo As Boolean = False) As String
        If DirReal(sPath) = False Then
            Call MkDir(sPath)
            Return "A directory is created"
        Else
            Return "Directory is already there"
        End If
    End Function

    'this is meant to remove the file that is selected like this "C:\like\you\file.exe" and give you the pathname in return
    Function PathName(vZ As String) As String
        Dim vA As Integer, stemppath As String
        stemppath = ""
        'loop through every char to find when the folder is all that is left
        For vA = Len(vZ) To 1 Step -1
            'it looks for char \
            If Mid(vZ, vA, 1) = "\" Then
                'this is what we wanted "C:\like\you\", ended up with the path without the file
                stemppath = Left(vZ, vA)
                Exit For
            End If
        Next vA
        Return stemppath
        'warning: if you is a directory and you do not use a "\" at the end than i gots messy and then that is what you get "C:\like\" missing the folder "you"
    End Function

    'this is when you need to know in what folder you are, do not forget if you pass a directory it ends with the "\"
    Function DirName(vZ As String) As String
        Dim vA As Integer, vC As Integer, sDirName As String
        sDirName = Path(vZ)
        'search for the second directory
        For vA = Len(sDirName) To 1 Step -1
            If Mid(sDirName, vA, 1) = "\" And vC = 0 Then
                vC = vA - 1
            ElseIf Mid(sDirName, vA, 1) = "\" Then
                'oke now we know what folder we are
                sDirName = Mid(sDirName, vA + 1, vC - vA)
                Exit For
            End If
        Next vA
        Return sDirName
    End Function

    'this is when we want to know the file extension. you can give the full directory if you want like this "c:\cat\mouse.exe" or just the file like "cat.exe"
    Function FileExtension(vZ As String) As String
        Dim vA As Integer, sFileExtension As String
        sFileExtension = ""
        For vA = Len(vZ) To 1 Step -1
            If Mid(vZ, vA, 1) = "." Then
                'going from right to left while seeking for a dot to learn when it founds the extension
                sFileExtension = Mid(vZ, vA + 1)
                Exit For
            End If
        Next vA
        'if we reach this there is no file or the file does not have an extension
        Return sFileExtension
    End Function

    'if you want the file name from a directory string you can use it like "c:\bravo.com"
    Function FileName(vZ As String) As String
        Dim vA As Integer, sFileName As Integer
        For vA = Len(vZ) To 1 Step -1
            'first we have to find the extension and remove it from the string
            If Mid(vZ, vA, 1) = "." Then
                sFileName = Left(vZ, vA - 1)
                Exit For
            End If
        Next vA
        If sFileName = "" Then
            sFileName = vZ
        End If
        'so now the filename and directory is left. so we have to figur out when the path is beginning and remove that too.
        For vA = Len(sFileName) To 1 Step -1
            If Mid(sFileName, vA, 1) = "\" Then
                'path removed, file extension removed. al we have left is the name of the file
                sFileName = Right(sFileName, Len(sFileName) - vA)
                Exit For
            End If
        Next vA
        Return sFileName
    End Function

    'in some functions you have to have a good path that closes with "\" so you can do "C:\maps\folder" and it ends up like this "c:\maps\folder\"
    Function Path(sPath As String) As String
        Dim sPathTemp As String
        'is there anything??
        If sPath <> "" Then
            'this is strange but it should do any good i quess
            If Left(sPath, 2) = "//" Then
                'if the wrong pahtlike stuff is happing correct it
                If Right(sPath, 1) <> "/" Then
                    sPathTemp = sPath & "\"
                Else
                    sPathTemp = sPath
                End If
                'see if the end is correct. if there is no "\" make one
            ElseIf Right(sPath, 1) <> "\" Then
                sPathTemp = sPath & "\"
            Else
                'it looks like it is good the way it was
                sPathTemp = sPath
            End If
        Else
            sPathTemp = ""
        End If
        Return sPathTemp
    End Function

    '------------------------------------------THIS HAS TO DO WITH DATA----------------------------------------------------

    'this function is made so that the text is width enough to fit in a textbox or other control it uses
    'if it is to width it makes the text on such a way it will fit in a new line
    'Function TextVerdelen(vText As String, vFont As Variant, vWidth As Long, vMultiLine As Boolean) As String
    'vFont is set but there is nothing done with it. so for now i have removed it
    'vWidth??? what number do i have to give. without the code that uses this i could not recover what it is need for
    'multie line is set to true then run the code. if it is single line no height's have te be counted
    Function TextVerdelen(vText As String, vWidth As Long, vMultiLine As Boolean) As String
        Dim vPlekInText As Double, vBeginNieuweRegel As Double, sText As String, vEventueleRegelBreek As Double
        If vMultiLine = False Then
            Return vText
        End If
        vBeginNieuweRegel = 1
        vEventueleRegelBreek = 1
        sText = ""
        For vPlekInText = 1 To Len(vText)
            'kijk hier of de regel echt verbroken wordt
            If Mid(vText, vPlekInText, 2) = vbCrLf Then
                sText &= Mid(vText, vBeginNieuweRegel, vPlekInText - vBeginNieuweRegel + 2)
                'de regel is gebroken dus maak een een nieuwe start zodat de rest behandeld kan worden
                vBeginNieuweRegel = vPlekInText + 2 'plus de enter
                vPlekInText += 2
                'controleer hier of de functie afgebroken moet worden
                If vBeginNieuweRegel > Len(vText) Then
                    Return sText
                End If
                vEventueleRegelBreek = vBeginNieuweRegel
            ElseIf Mid(vText, vPlekInText, 1) = " " Then
                'sla hier de regel breek op voor als de tekst gebroken moet worden in een later stadium
                vEventueleRegelBreek = vPlekInText + 1
            End If
            'controleer hier of de geselecteerde tekst breeder is dan de regel
            If TextWidth(Mid(vText, vBeginNieuweRegel, vPlekInText - vBeginNieuweRegel)) > Val(vWidth - 30) Then
                If vBeginNieuweRegel = vEventueleRegelBreek Then
                    'hier kom je een woord tegen die langer is dan de regel breek dat woord
                    sText = sText & Mid(vText, vBeginNieuweRegel, vPlekInText - vBeginNieuweRegel) & vbCrLf
                    vEventueleRegelBreek = vPlekInText
                    vBeginNieuweRegel = vEventueleRegelBreek
                Else
                    'hier zoek je de eventuele regel breek op om te gebruiken om de regel te breken
                    sText = sText & Mid(vText, vBeginNieuweRegel, vEventueleRegelBreek - vBeginNieuweRegel) & vbCrLf
                    vBeginNieuweRegel = vEventueleRegelBreek
                End If
            End If
        Next vPlekInText
        If (vPlekInText - vBeginNieuweRegel - 1) > 0 Then
            If TextWidth(Mid(vText, vBeginNieuweRegel, vPlekInText - vBeginNieuweRegel - 1)) > Val(vWidth - 30) Then
                'als de regel groter is dan de control zet het laatst deel dan hier op
                sText = sText & Mid(vText, vBeginNieuweRegel, vEventueleRegelBreek - vBeginNieuweRegel) & vbCrLf & Mid(vText, vEventueleRegelBreek)
            Else
                'zet de rest van de regel erop
                sText &= Mid(vText, vBeginNieuweRegel)
            End If
        Else
            'zet de rest van de regel erop
            sText &= Mid(vText, vBeginNieuweRegel)
        End If
        Return sText
    End Function

    'this counts the lines of text is used. and give us the height of the text
    Function TextHeight(vText As String, vSize As Integer) As Long
        Dim vA As Integer, vHeight As Integer
        vHeight = 1
        For vA = 1 To Len(vText)
            If Mid(vText, vA, 2) = vbCrLf Then
                vHeight += 1
            End If
        Next vA
        vHeight *= vSize
        Return vHeight
    End Function

    'this is code that want to know how width the text is. so it could be used to select the text on the right size
    Function TextWidth(sB As String) As Integer
        'sB is text that is been selected
        Dim iA As Integer, iB As Integer
        'every letter or number or what ever is been checked to see how width it is
        For iA = Len(sB) To 1 Step -1
            Select Case Mid(sB, iA, 1)
                Case "|", "'"
                    iB += 3
                Case "i", "l", "t", "f", "j", "I", " ", "!", "(", ")", "-", "[", "]", "\", ":", ",", ".", "`", "/", ";"
                    iB += 4
                Case "r", "*", "{", "}"
                    iB += 5
                Case "s"
                    iB += 6
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "c", "k", "v", "y", "J", "Z", "#", "$", "_"
                    iB += 7
                Case "a", "b", "d", "e", "g", "h", "o", "n", "p", "q", "u", "x", "z", "E", "F", "L", "T", "Y", "~", "^", "+", "=", "<", ">", "?"
                    iB += 8
                Case "A", "B", "C", "D", "H", "K", "N", "P", "R", "S", "U", "V", "X", "&"
                    iB += 9
                Case "G", "O", "Q", "%"
                    iB += 10
                Case "w", "M"
                    iB += 11
                Case "m"
                    iB += 12
                Case "W", "@"
                    iB += 13
                Case vbCr
                    iB += 5
            End Select
        Next iA
        'it returns how width the total text is
        Return iB
    End Function

    'hmm have to see what this is doing right?!
    'i have read on the internet, that the first 31 asc are non printeble codes. so if we have one only the text before will be returnd
    Function TextOnly(sZ As String) As String
        Dim vA As Integer, sA As String
        sA = sZ
        For vA = 1 To Len(sZ)
            Select Case Asc(Mid(sZ, vA, 1))
                Case 0 To 31
                    sA = Left(sZ, vA - 1)
                    Exit For
            End Select
        Next vA
        Return sA
    End Function

    'what does this do??? byval? where do you need byval for??? I DO NOT UNDERSTAND WHAT USE THIS COULD BE. for the moment
    Function Bytes2Len(ByVal TwoBytes As String) As Long
        Dim FirstByteVal As Long
        Return Asc(Left(TwoBytes, 1))
        Exit Function
        FirstByteVal = Asc(Left(TwoBytes, 1))
        If FirstByteVal >= 128 Then
            FirstByteVal -= 128
        End If
        Return 256 * FirstByteVal + Asc(Right(TwoBytes, 1))
    End Function

    'this checks if data in database is deleted. i want the data still be there. so it could be recoverd
    Function Deleted(vZ As String) As Boolean
        If Len(vZ) > 7 Then
            If LCase(Left(vZ, 7)) = "deleted" Then
                Return True
            End If
        End If
        Return False
    End Function

    'this is for base 64 encryption. it uses this to check all the charakters one by one
    Function CharCodeB64(strB64Char As String) As Integer
        'if the char does not exist in this line it wil return us a zero
        Const BASE64_TABLE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
        'returns you a number of or given char
        Return InStr(BASE64_TABLE, strB64Char)
    End Function

    'it says what it says right. i am getting tired from al the working
    Function CountReturns(iX As Integer, sA As String) As Integer
        Dim iA As Long, iB As Integer
        For iA = 1 To Len(sA)
            Select Case Mid(sA, iA, 2)
                Case vbCrLf
                    'yay we found a new return in the text. add the number bro
                    iB += 1
                    If iX = iB Then
                        Exit For
                    End If
            End Select
        Next iA
        If iX > 0 Then
            Return iA
        Else
            Return iB
        End If
    End Function

    'this is to make false or true. depending on witch numbers you give in
    Function CCBool(sZ As String) As Boolean
        If sZ = "" Then
            'by using "" you could counter a error
            Return False
        Else
            'must be a number or it will be false
            Return CBool(NumberOnly(sZ, False))
        End If
        'why the fuck did i ever made this up, because it is logic
    End Function

    'this code converts a true or a false to a 1 or 0
    Function Boolean2Value(vZ As Boolean) As Byte
        If vZ = True Then
            Return 1
        Else
            Return 0
        End If
    End Function

    'this code makes it possible that it cant go wrong, if you realy want a byte, why the fuck did i ever make this
    Function ByteAlways(vZ As Integer) As Byte
        If vZ > 255 Then
            'the max size of a byte
            Return 255
        ElseIf vZ < 0 Then
            'a byte can not be negative
            Return 0
        Else
            'everything else make it a number
            Return vZ
        End If
    End Function

    'this function replace the "¿" karakter to "/", more information about why "¿" is here comes later on
    Function BlockEndRemove(vZ As String) As String
        Dim vA As Integer, sBlockEndRemove As String
        sBlockEndRemove = vZ
        For vA = 1 To Len(sBlockEndRemove)
            If Mid(sBlockEndRemove, vA, 1) = "¿" Then
                If vA = Len(sBlockEndRemove) Then
                    sBlockEndRemove = Left(sBlockEndRemove, vA - 1)
                    Exit For
                Else
                    sBlockEndRemove = Left(sBlockEndRemove, vA - 1) & "/" & Right(sBlockEndRemove, Len(sBlockEndRemove) - vA)
                End If
            End If
        Next vA
        Return sBlockEndRemove
    End Function

    'use the symbol "&" in this case there must be two every were in the string
    Function AndDouble(vZ As String) As String
        Dim vA As Integer, sAndDouble As String
        sAndDouble = ""
        For vA = 1 To Len(vZ)
            If Mid(vZ, vA, 1) = "&" And Mid(vZ, vA, 2) <> "&&" Then
                'this is done so we can write a visible "&" to the screen
                sAndDouble = BlockEndRemove(Left(vZ, vA) & "&" & Right(vZ, Len(vZ) - vA))
                Exit For
            End If
        Next vA
        If vZ <> "?" Then
            sAndDouble = BlockEndRemove(vZ)
        End If
        Return sAndDouble
    End Function

    'this time there must only be one "&" next to eachother, so it will be a shortcut whitin a button we can control with alt key
    Function AndSingel(vZ As String) As String
        Dim vA As Integer
        For vA = 1 To Len(vZ)
            If Mid(vZ, vA, 2) = "&&" Then
                'we just use one shortcut so this code does not have to be more advenced
                Return Left(vZ, vA) & Right(vZ, Len(vZ) - vA - 1)
            End If
        Next vA
        Return vZ
    End Function

    'replaces parts of text
    Function TextReplace(sText As String, sRemove As String, Optional sReplace As String = "") As String
        Dim iA As Single, sC As String
        If sText = "" Or sRemove = "" Then
            Return sText
        End If
        sC = sText
        For iA = Len(sC) To 1 Step -1
            If Mid(sC, iA, Len(sRemove)) = sRemove Then
                sC = Left(sC, iA - 1) & sReplace & Right(sC, Len(sC) - ((iA - 1) + Len(sRemove)))
            End If
        Next iA
        If Len(sRemove) <= Len(sC) Then
            If Right(sC, Len(sRemove)) = sRemove Then
                sC = Left(sC, Len(sC) - Len(sRemove)) & sReplace
            End If
            If Left(sC, Len(sRemove)) = sRemove Then
                sC = Right(sC, Len(sC) - Len(sRemove)) & sReplace
            End If
        End If
        Return sC
    End Function

    'in this part of code all the beginning text are set to upper case
    Function TextUppercases(vZ As String) As String
        Dim vA As Integer
        If vZ <> "" Then
            vZ = UCase(Left(vZ, 1) & Right(vZ, Len(vZ) - 1))
            For vA = 2 To Len(vZ)
                If vA > Len(vZ) Then
                    Exit For
                End If
                If Mid(vZ, vA, 1) = " " Then
                    If vA = Len(vZ) Then
                        Exit For
                    End If
                    If Mid(vZ, vA + 1, 1) <> " " Then
                        vZ = Left(vZ, vA) & UCase(Mid(vZ, (vA + 1), 1)) & Right(vZ, Len(vZ) - vA - 1)
                        vA += 1
                    End If
                Else
                    vZ = Left(vZ, vA - 1) & LCase(Mid(vZ, vA, 1)) & Right(vZ, Len(vZ) - vA)
                End If
            Next vA
        End If
        Return vZ
    End Function

    'this code pimps some things
    Function TextRemoveDouble(vZ As String) As String
        Dim vA As Integer, sZ As String, sX As String
        sZ = ""
        'search for "-" and "-" to be replaced with space
        For vA = 1 To Len(vZ)
            If Mid(vZ, vA, 1) = "-" Or Mid(vZ, vA, 1) = "_" Then
                sZ += " "
            Else
                sZ += Mid(vZ, vA, 1)
            End If
        Next
        sX = ""
        'find all the spaces to exclude two spaces next to eachother
        For vA = Len(sZ) To 1 Step -1
            If Mid(sZ, vA, 2) = "  " Then
                'do nothing
            Else
                sX = Mid(sZ, vA, 1) + sX
            End If
        Next
        Return sX
    End Function

    'this replaces all the spaces to something a browser need to be running from
    Function URLEncode(FileName As String) As String
        Dim vA As Single, sURLEncode As String
        sURLEncode = FileName
        For vA = Len(sURLEncode) To 1 Step -1
            If Mid(sURLEncode, vA, 1) = " " Then
                '%20 is the space, a web browser need to operate
                sURLEncode = Left(sURLEncode, vA - 1) & "%20" & Mid(sURLEncode, vA + 1)
            End If
        Next vA
        Return sURLEncode
    End Function

    'this get rid of all the browser stuff that is needed to have in a url in your browser
    Function URLDecode(FileName As String) As String
        Dim vA As Single, vURLDecode As String
        vURLDecode = FileName
        For vA = 1 To Len(vURLDecode)
            If Mid(vURLDecode, vA, 1) = "%" Then
                vURLDecode = Left(vURLDecode, vA - 1) & " " & Mid(vURLDecode, vA + 3)
            End If
        Next vA
        Return vURLDecode
    End Function

    'all it does is reversing the string
    Function ReverseString(sZ As String) As String
        Dim sA As String, vA As Integer
        sA = ""
        'flips all the tekst by making small steps
        For vA = Len(sZ) To 1 Step -1
            sA = sA & Mid(sZ, vA, 1)
        Next vA
        Return sA
    End Function

    'all it does is replacing "_" for a space like this " "
    Function Score2Space(vZ As String) As String
        Dim vA As Integer
        If vZ <> "" Then
            For vA = 1 To Len(vZ)
                If vA > Len(vZ) Then
                    Exit For
                End If
                If Mid(vZ, vA, 1) = "_" Then
                    vZ = Left(vZ, vA - 1) & " " & Right(vZ, Len(vZ) - (vA))
                End If
            Next vA
        End If
        Return vZ
    End Function

    'this does replace space " " with "_"
    Function Space2Score(vZ As String) As String
        Dim vA As Integer
        If vZ <> "" Then
            'searching for the big work
            For vA = 1 To Len(vZ)
                ' no nothing left
                If vA > Len(vZ) Then
                    Exit For
                End If
                'found it
                If Mid(vZ, vA, 1) = " " Then
                    'replace it, yay
                    vZ = Left(vZ, vA - 1) & "_" & Right(vZ, Len(vZ) - (vA))
                End If
            Next vA
        End If
        Return vZ
    End Function

    'looks for 1 and 0. and if that is true then you will tell it is true or false
    Function Value2Boolean(vZ As Byte) As Boolean
        If vZ = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    '----------------------------------------THIS IS FOR KEYBOARDS AND MOUSE ONLY-------------------------------------------------

    ' again a translation from one number to another
    Function KeyAsciiToDX(vAsccii As Integer) As Byte
        If vAsccii = 13 Then
            Return 28 'return
        ElseIf vAsccii = 112 Then
            Return 59 'F1
        ElseIf vAsccii = 113 Then
            Return 60 'F2
        ElseIf vAsccii = 114 Then
            Return 61 'F3
        ElseIf vAsccii = 115 Then
            Return 62 'F4
        ElseIf vAsccii = 116 Then
            Return 63 'F5
        ElseIf vAsccii = 117 Then
            Return 64 'F6
        ElseIf vAsccii = 118 Then
            Return 65 'F7
        ElseIf vAsccii = 119 Then
            Return 66 'F8
        ElseIf vAsccii = 120 Then
            Return 67 'F9
        ElseIf vAsccii = 121 Then
            Return 68 'F10
        ElseIf vAsccii = 122 Then
            Return 87 'F11
        ElseIf vAsccii = 123 Then
            Return 88 'F12
        ElseIf vAsccii = 9 Then
            Return 15 'Tab
        ElseIf vAsccii = 38 Then
            Return 200 'Up
        ElseIf vAsccii = 40 Then
            Return 208 'Down
        ElseIf vAsccii = 27 Then
            Return 1 'Escape
        ElseIf vAsccii = 37 Then
            Return 203 'Left
        ElseIf vAsccii = 39 Then
            Return 205 'Right
        ElseIf vAsccii = 35 Then
            Return 207 'End
        ElseIf vAsccii = 36 Then
            Return 199 'Home
        ElseIf vAsccii = 8 Then
            Return 14 'Back Space
        ElseIf vAsccii = 46 Then
            Return 211 'Delete
        ElseIf vAsccii = 27 Then
            Return 1 'Escape
        Else
            'could not make a chouse
            Return 0
        End If
    End Function

    'translating. from one dedicated number to another dedicated number. if you ask me. i do not like this shit
    Function KeyAsciiF(KeyCode As Byte, Shift As Byte) As Byte
        If Shift = 1 Then
            Select Case KeyCode
                Case 188 '<
                    Return 60 '<
                Case 190 '>
                    Return 62 '>
                Case 49 '!
                    Return 33 '!
                Case 50 '@
                    Return 64 '@
                Case 51 '#
                    Return 35 '#
                Case 52 '$
                    Return 36 '$
                Case 53 '%
                    Return 37 '%
                Case 54 '^
                    Return 94 '^
                Case 55 '&
                    Return 38 '&
                Case 57 '(
                    Return 40 '(
                Case 48 ')
                    Return 41 ')
                Case 189 '_
                    Return 95 '_
                Case 187 '+
                    Return 43 '+
                Case 56 '*
                    Return 42 '*
                Case 192 '~
                    Return 126 '~
                Case 220 '|
                    Return 124 '|
                Case 191 '/
                    Return 47 '/
                Case 219 '{
                    Return 123 '{
                Case 221 '}
                    Return 125 '}
                Case 222 '"
                    Return 34 '"
                Case 186 ':
                    Return 58 ':
            End Select
        Else
            Select Case KeyCode
                Case 191 '?
                    Return 63 '?
                Case 65 To 90 'a to z
                    Return KeyCode
                Case 48 To 57 '0 to 9
                    Return KeyCode
                Case 96 To 105 '0 to 9
                    Return KeyCode - 48
                Case 111 '/
                    Return 47 '/
                Case 107 '+
                    Return 43 '+
                Case 109, 189  '-
                    Return 45 '-
                Case 106  '*
                    Return 42 '*
                Case 192 '`
                    Return 96 '`
                Case 187 '=
                    Return 61 '=
                Case 220 '\
                    Return 92 '\
                Case 221 ']
                    Return 93 ']
                Case 219 '[
                    Return 91 '[
                Case 186 ';
                    Return 59 ';
                Case 222 ''
                    Return 39 ''
                Case 190 '.
                    Return 46 '.
                Case 188 ',
                    Return 44 ',
            End Select
        End If
        Return 0
    End Function

    ' i love to code and this is what becomes of it. all references to another keyboard calling way. i do not understand why there are different in the first place
    'Function VKToDIK(vVK As Integer) As Integer
    '    Dim iA As Integer
    '    Select Case vVK
    '        Case 27  'esc
    '            Return 1
    '        Case 179 'ply pause
    '            Return 162
    '        Case 178 'stop
    '            Return 164
    '        Case 177 'prev
    '            Return 144
    '        Case 176 'next
    '            Return 153
    '        Case 173 'mute
    '            Return 160
    '        Case 174 'min sound
    '            Return 174
    '        Case 175 'plus sound
    '            Return 176
    '        Case 180 'mail
    '            Return 236
    '        Case 172 'home web
    '            Return 178
    '        Case 183 'calc
    '            Return 161
    '        Case 133 'close
    '            Return 109
    '        Case 166 'back virtual
    '            Return 234
    '        Case 192 '`
    '            Return 41
    '        Case 189 '-
    '            Return 12
    '        Case 187 '=
    '            Return 13
    '        Case 219 '[
    '            Return 26
    '        Case 221 ']
    '            Return 27
    '        Case 186 ';
    '            Return 39
    '        Case 222 ''
    '            Return 40
    '        Case 220 '\
    '            Return 43
    '        Case 188 ',
    '            Return 51
    '        Case 190 '.
    '            Return 52
    '        Case 191 '/
    '            Return 53
    '        Case 160 'lshift
    '            Return DIK_LSHIFT
    '        Case 161 'rshit
    '            Return DIK_RSHIFT
    '        Case 8 'backspace
    '            Return DIK_BACKSPACE
    '        Case 9 'tab
    '            Return DIK_TAB
    '        Case 20 'CAPSLOCK
    '            Return DIK_CAPSLOCK
    '        Case 162 'lcrtl
    '            Return DIK_LCONTROL
    '        Case 163 'rctrl
    '            Return DIK_RCONTROL
    '        Case 164 'lalt ralt
    '            Return DIK_LALT
    '        Case 226 '\
    '            Return 86
    '        Case 91 'winkey
    '            Return 219
    '        Case 116 '???
    '            Return -1 'DIK_LALT
    '        Case 32 'space
    '            Return DIK_SPACE
    '        Case 93 'menu
    '            Return 221
    '        Case 13 'lreturn rreturn
    '            Return DIK_RETURN
    '        Case 255 'f mode
    '            Return 218
    '        Case 45 'insert
    '            Return DIK_INSERT
    '        Case 44 'printscreen
    '            Return 183
    '        Case 144 'num lock
    '            Return DIK_NUMLOCK
    '        Case 19 'pause break
    '            Return DIK_PAUSE
    '        Case 36 'home
    '            Return DIK_HOME
    '        Case 35 'end
    '            Return DIK_END
    '        Case 46 'delete
    '            Return DIK_DELETE
    '        Case 33 'page up
    '            Return DIK_PGUP
    '        Case 34 'page down
    '            Return DIK_PGDN
    '        Case 38 'up
    '            Return DIK_UP
    '        Case 40 'down
    '            Return DIK_DOWN
    '        Case 37 'left
    '            Return DIK_LEFT
    '        Case 39 'right
    '            Return DIK_RIGHT
    '        Case 111 '/
    '            Return 181
    '        Case 106 '*
    '            Return 55
    '        Case 109 '-
    '            Return 74
    '        Case 107 '+
    '            Return 78
    '        Case 110 '.
    '            Return 83
    '        Case 96 To 105 '0 '9
    '            Return 82 + 14
    '        Case 96
    '            Return DIK_NUMPAD0
    '        Case 97
    '            Return DIK_NUMPAD1
    '        Case 98
    '            Return DIK_NUMPAD2
    '        Case 99
    '            Return DIK_NUMPAD3
    '        Case 100
    '            Return DIK_NUMPAD4
    '        Case 101
    '            Return DIK_NUMPAD5
    '        Case 102
    '            Return DIK_NUMPAD6
    '        Case 103
    '            Return DIK_NUMPAD7
    '        Case 104
    '            Return DIK_NUMPAD8
    '        Case 105
    '            Return DIK_NUMPAD9
    '        Case 112
    '            Return DIK_F1
    '        Case 113
    '            Return DIK_F2
    '        Case 114
    '            Return DIK_F3
    '        Case 115
    '            Return DIK_F4
    '        Case 116
    '            Return DIK_F5
    '        Case 117
    '            Return DIK_F6
    '        Case 118
    '            Return DIK_F7
    '        Case 119
    '            Return DIK_F8
    '        Case 120
    '            Return DIK_F9
    '        Case 121
    '            Return DIK_F10
    '        Case 122
    '            Return DIK_F11
    '        Case 123
    '            Return DIK_F12
    '        Case 124
    '            Return DIK_F13
    '        Case 125
    '            Return DIK_F14
    '        Case 126
    '            Return DIK_F15
    '        Case 108, 59, 92, 47, 43, 42, 61 ' disable
    '            Return -1
    '        Case Else
    '            For iA = 0 To 255
    '                If LCase(KeyNames(iA)) = LCase(Chr(vVK)) Then
    '                    Return iA
    '                End If
    '            Next iA
    '            'it has to return something
    '            Return -1
    '    End Select
    'End Function

    'this must be great is you understand why i made this. even i do not understand this
    'Function VKToDIKPress(vVK As Integer) As Integer
    '    Select Case vVK
    '        Case 160 'lshift
    '            Return DIK_LSHIFT
    '        Case 161 'rshit
    '            Return DIK_RSHIFT
    '        Case 162 'lcrtl
    '            Return DIK_LCONTROL
    '        Case 163 'rctrl
    '            Return DIK_RCONTROL
    '        Case 164 'lalt ralt
    '            Return DIK_LALT
    '        Case Else
    '            Return -1
    '    End Select
    'End Function

    'this has to do something with phone numbers. it returns a key code, for now it is good enough but could be better
    Function PhoneNumber(vKeyAscii As Integer, vSelLength As Integer, vText As String) As Byte
        'if you enter a number, and the lenght could not be longer then 10 digits
        If vKeyAscii > 47 And vKeyAscii < 58 And Len(vText) - vSelLength < 10 Then
            Return vKeyAscii
        Else
            'gives a reaction when things go wrong
            Return 0
        End If
    End Function

    'this is needed when you want to type in a postal code. it checks if it match to postal code format. (in dutch terms)
    Function PostalCode(vKeyAscii As Integer, vSelStart As Integer, vText As String) As String
        Dim vA As Integer, iB As Integer, vPostalCode As String
        vPostalCode = vKeyAscii
        If vKeyAscii > 47 And vKeyAscii < 58 Then
            If vSelStart < 4 Then
                iB = 0
                For vA = 1 To Len(vText)
                    If Asc(Mid(vText, vA, 1)) > 47 And Asc(Mid(vText, vA, 1)) < 58 Then
                        iB += 1
                    End If
                Next vA
                If iB = 4 Then
                    vPostalCode = 0
                End If
            Else
                vPostalCode = 0
            End If
        ElseIf vKeyAscii > 64 And vKeyAscii < 91 Then
            If vSelStart = 4 Then
                vText = Left(vText, 4) & " " & Right(vText, Len(vText) - 4)
                vSelStart = 5
            ElseIf vSelStart > 3 And vSelStart < 7 Then
                iB = 0
                For vA = 1 To Len(vText)
                    If Asc(Mid(vText, vA, 1)) > 64 And Asc(Mid(vText, vA, 1)) < 91 Then
                        iB += 1
                    End If
                Next vA
                If iB = 2 Then
                    vPostalCode = 0
                End If
            Else
                vPostalCode = 0
            End If
        ElseIf vKeyAscii > 96 And vKeyAscii < 123 Then
            vPostalCode = (vKeyAscii - 32)
            If vSelStart = 4 Then
                vText = Left(vText, 4) & " " & Right(vText, Len(vText) - 4)
                vSelStart = 5
            ElseIf vSelStart > 3 And vSelStart < 7 Then
                iB = 0
                For vA = 1 To Len(vText)
                    If Asc(Mid(vText, vA, 1)) > 64 And Asc(Mid(vText, vA, 1)) < 91 Then
                        iB += 1
                    End If
                Next vA
                If iB = 2 Then
                    vPostalCode = 0
                End If
            Else
                vPostalCode = 0
            End If
        Else
            vPostalCode = 0
        End If
        Return vPostalCode
    End Function

    '--------------------------------------------THIS IS FOR NUMBERS---------------------------------------------------

    'simply reverse positive to negative and the other way around
    Function Xorr(vZ As Byte) As Byte
        If vZ = 0 Then
            Return 1
        Else
            Return 0
        End If
    End Function

    'this is for auto correction. when you deal data controles. and it always have to be a number.
    Function CurDecimal(vZ As String, Optional vDecimal As Integer = 2, Optional Cut As Integer = False) As String
        Dim vA As Integer, sA As String, vC As Integer, vCurDecimal As String
        vCurDecimal = ""
        If NumberOnly(vZ) = False Or Len(vZ) > 255 Then
            If vDecimal = 2 Then
                Return "0,00"
            Else
                Return "0,000"
            End If
        End If
        sA = CStr(vZ)
        For vA = 1 To Len(sA)
            Select Case Mid(sA, vA, 1)
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "-"
                Case ","
                    If vC <> 0 Then
                        If vDecimal = 2 Then
                            Return "0,00"
                        Else
                            Return "0,000"
                        End If
                    Else
                        vC = 1
                    End If
                Case "."
                    If vC <> 0 Then
                        If vDecimal = 2 Then
                            Return "0,00"
                        Else
                            Return "0,000"
                        End If
                    Else
                        vC = 2
                    End If
                Case Else
                    If vDecimal = 2 Then
                        Return "0,00"
                    Else
                        Return "0,000"
                    End If
            End Select
        Next vA
        Select Case vC
            Case 0
                If sA = "" Then
                    sA = "0"
                End If
                If vDecimal = 2 Then
                    Return (sA & ",00")
                Else
                    Return (sA & ",000")
                End If
            Case 2
                For vA = 1 To Len(sA)
                    If Mid(sA, vA, 1) = "." Then
                        If vA = 1 Then
                            sA = "," & Right(sA, Len(sA) - 1)
                        ElseIf vA = Len(sA) Then
                            sA = Left(sA, Len(sA) - 1) & ","
                        Else
                            sA = Left(sA, vA - 1) & "," & Right(sA, Len(sA) - vA)
                        End If
                        Exit For
                    End If
                Next vA
        End Select
        For vA = Len(sA) To 1 Step -1
            If Mid(sA, vA, 1) = "," Then
                Select Case vA
                    Case Len(sA)
                        If vDecimal = 2 Then
                            vCurDecimal = sA & "00"
                        Else
                            vCurDecimal = sA & "000"
                        End If
                        Exit For
                    Case Len(sA) - 1
                        If vDecimal = 2 Then
                            vCurDecimal = sA & "0"
                        Else
                            vCurDecimal = sA & "00"
                        End If
                        Exit For
                    Case Len(sA) - 2
                        If vDecimal = 2 Then
                            vCurDecimal = sA
                        Else
                            vCurDecimal = sA & "0"
                        End If
                        Exit For
                    Case Else
                        If Cut = True Then
                            If vDecimal = 2 Then
                                vCurDecimal = Left(sA, vA + 2)
                            Else
                                vCurDecimal = Left(sA, vA + 3)
                            End If
                        ElseIf vDecimal = 2 Then
                            For vC = 1 To Len(sA)
                                If Mid(sA, vC, 1) = "," Then
                                    If Mid(sA, vC + 3, 1) > "4" Then
                                        If Len(Mid(sA, vC + 2, 1) + 1) = 2 Then
                                            If Len(Val(Val(Mid(sA, vC + 1, 1)) + 1) & "0") > 2 Then
                                                vCurDecimal = (Left(sA, vC - 1) + 1) & "," & Right(Val(Mid(sA, vC + 1, 1) + 1) & "0", 2)
                                            Else
                                                vCurDecimal = Left(sA, vC) & Val(Val(Mid(sA, vC + 1, 1)) + 1) & "0"
                                            End If
                                        Else
                                            vCurDecimal = Left(sA, vC + 1) & Val(Val(Mid(sA, vC + 2, 1)) + 1)
                                        End If
                                    Else
                                        vCurDecimal = Left(sA, vC + 2)
                                    End If
                                    Exit For
                                End If
                            Next vC
                        End If
                        Exit For
                End Select
            End If
        Next vA
        If Left(vCurDecimal, 1) = "," Then
            Return ("0" & vCurDecimal)
        Else
            Return vCurDecimal
        End If
    End Function

    'this is to look on the bright side of life
    Function PositiveAlways(vA As Single) As Integer
        Dim sA As Single
        sA = vA
        If sA <> 0 Then
            If Left(sA, 1) = "-" Then
                sA /= sA - 1
            End If
            Return sA
        Else
            Return sA
        End If
    End Function

    'check this if there are only numbers in this string, or comma's and points, or minus char
    Function NumberOnly(vZ As String, Optional vDec As Boolean = True) As Boolean
        Dim vA As Single
        Const NUMBERS_TABLE As String = "0123456789"
        Const CALC_TABLE As String = ",.-"
        For vA = 1 To Len(vZ)
            If InStr(NUMBERS_TABLE, Mid(vZ, vA, 1)) <> 0 Then
                'this char is a number. so it is good
            ElseIf InStr(CALC_TABLE, Mid(vZ, vA, 1)) <> 0 Then
                'exception for currencies
                If vDec = False Then
                    Return False
                End If
            Else
                Return False
            End If
        Next vA
        'every char has been checked if it is only numbers in the string. so this is correct
        Return True
    End Function

    '------------------------------------------------THIS IS FOR NETWORKING---------------------------------------------------

    'this is a sort of translation from numberic winsock codes to tekst descriptio
    Function WinSockState(vZ As Byte) As String
        Select Case vZ
            Case 0
                Return "Closed"
            Case 1
                Return "Open"
            Case 2
                Return "Listening"
            Case 3
                Return "Connection Pending"
            Case 4
                Return "Resolving Host"
            Case 5
                Return "Host Resolved"
            Case 6
                Return "Connecting"
            Case 7
                Return "Connected"
            Case 8
                Return "Closing"
            Case 9
                Return "Error"
            Case Else
                Return "unknown terretorie"
        End Select
    End Function

    'this could be for networking. but i realy do not know, what it should have to be doing???
    'i think this is for cheaking the bytes we have. so we could go one where we left with downloading a file
    Function RecvSequence(dZ As Double) As String
        Dim TempString As String, TempSequence As Double, TempByte As Double, vA As Integer
        TempSequence = dZ
        TempString = ""
        For vA = 1 To 4
            TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
            TempSequence = Int(TempSequence / 256)
            TempString = Chr(TempByte) & TempString
        Next
        Return TempString
    End Function

    'this is the same. i have used this in my code to do something. but what it is for is hard to know if i do not have the code that use this
    'i now see there both doing the same
    Function SendSequence(dZ As Double) As String
        Dim TempString As String, TempSequence As Double, TempByte As Double, vA As Byte
        TempSequence = dZ
        TempString = ""
        For vA = 1 To 4
            TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
            TempSequence = Int(TempSequence / 256)
            TempString = Chr(TempByte) & TempString
        Next
        Return TempString
    End Function

    '-----------------------------------THIS IS FOR EMOTICONS-------------------------------------------------------------

    'this is used when combining a text color and a transparenty level
    Function GetTextColour(ByVal p_lngTextColour As Long, ByVal p_intTransparency As Integer) As Long
        Dim intTransparency As Integer, intRed As Integer, intGreen As Integer, intBlue As Integer, strColour As String
        intTransparency = 255 - ((p_intTransparency / 100) * 255)
        intRed = (p_lngTextColour And &HFF&)
        intGreen = (p_lngTextColour And &HFF00&) \ &H100
        intBlue = (p_lngTextColour And &HFF0000) \ &H10000
        strColour = "&H"
        strColour = strColour & IIf(Len(Hex(intTransparency)) = 1, "0", "") & Hex(intTransparency)
        strColour = strColour & IIf(Len(Hex(intRed)) = 1, "0", "") & Hex(intRed)
        strColour = strColour & IIf(Len(Hex(intGreen)) = 1, "0", "") & Hex(intGreen)
        strColour = strColour & IIf(Len(Hex(intBlue)) = 1, "0", "") & Hex(intBlue)
        Return CLng(strColour)
    End Function

    'here we asign the picture number that is used in memory. we return it so we could print it on screen
    Function EmoticonsImage(sA As String) As Integer
        Select Case LCase(Left(sA, 2))
            Case ":)"
                Return 292
            Case ":d"
                Return 293
            Case ":p"
                Return 296
            Case ";)"
                Return 294
            Case ":@"
                Return 298
            Case ":S"
                Return 300
            Case ":$"
                Return 299
            Case ":("
                Return 301
            Case Else
                Select Case LCase(Left(sA, 3))
                    Case ":^)"
                        Return 363
                    Case ":-o"
                        Return 295
                    Case "(h)"
                        Return 297
                    Case ":'("
                        Return 302
                    Case ":|"
                        Return 303
                    Case "(6)"
                        Return 304
                    Case "(a)"
                        Return 305
                    Case "8o|"
                        Return 340
                    Case "8-|"
                        Return 341
                    Case "+o("
                        Return 344
                    Case "|-)"
                        Return 369
                    Case "*-)"
                        Return 364
                    Case ":-#"
                        Return 339
                    Case ":-*"
                        Return 343
                    Case "^o)"
                        Return 342
                    Case "8-)"
                        Return 367
                    Case "(l)"
                        Return 306
                    Case "(u)"
                        Return 307
                    Case "(m)"
                        Return 308
                    Case "(@)"
                        Return 309
                    Case "(&)"
                        Return 310
                    Case "(s)"
                        Return 311
                    Case "(*)"
                        Return 312
                    Case "(~)"
                        Return 313
                    Case "(8)"
                        Return 314
                    Case "(e)"
                        Return 315
                    Case "(b)"
                        Return 328
                    Case "(d)"
                        Return 329
                    Case "(z)"
                        Return 330
                    Case "(x)"
                        Return 331
                    Case "(y)"
                        Return 332
                    Case ":["
                        Return 334
                    Case "(n)"
                        Return 333
                    Case "(#)"
                        Return 337
                    Case "(r)"
                        Return 338
                    Case "({)"
                        Return 326
                    Case "(})"
                        Return 327
                    Case "(t)"
                        Return 325
                    Case "(c)"
                        Return 324
                    Case "(i)"
                        Return 323
                    Case "(p)"
                        Return 322
                    Case "(^)"
                        Return 321
                    Case "(g)"
                        Return 320
                    Case "(k)"
                        Return 319
                    Case "(f)"
                        Return 316
                    Case "(w)"
                        Return 317
                    Case "(o)"
                        Return 318
                    Case ":-)"
                        Return 292
                    Case "brb"
                        Return 296
                    Case Else
                        Select Case LCase(Left(sA, 4))
                            Case "<:o)"
                                Return 366
                            Case "(li)"
                                Return 365
                            Case "(sn)"
                                Return 345
                            Case "(tu)"
                                Return 346
                            Case "(pl)"
                                Return 347
                            Case "(||)"
                                Return 348
                            Case "(pi)"
                                Return 349
                            Case "(so)"
                                Return 350
                            Case "(au)"
                                Return 351
                            Case "(um)"
                                Return 353
                            Case "(ip)"
                                Return 354
                            Case "(co)"
                                Return 355
                            Case "(mp)"
                                Return 356
                            Case "(st)"
                                Return 358
                            Case "(h5)"
                                Return 360
                            Case "(mo)"
                                Return 361
                            Case "(ap)"
                                Return 352
                            Case Else
                                Select Case LCase(Left(sA, 5))
                                    Case "(brb)"
                                        Return 357
                                    Case "(bah)"
                                        Return 362
                                End Select
                        End Select
                End Select
        End Select
    End Function

    'this if for getting the right size for an emoticon
    'the emoticon is build from letters and replaced by pictures. so it has to have the right space so we could see it between the words on one line
    Function EmoticonsSize(sA As String) As String
        Select Case LCase(Left(sA, 2))
            Case ":)", ":d", ":p", ";)", ":@", ":S", ":$", ":("
                'here we give the right size, only i see that for all the emoticons is used the same size. that could not be correct
                Return "19*19"
            Case Else
                Select Case LCase(Left(sA, 3))
                    Case ":^)", ":-o", "(h)", ":'(", ":|", "(6)", "(a)", "8o|", "8-|", "+o(", "|-)", "*-)", ":-#", ":-*", "^o)", "8-)", "(l)", "(u)"
                        Return "19*19"
                    Case "(m)", "(@)", "(&)", "(s)", "(*)", "(~)", "(8)", "(e)", "(b)", "(d)", "(z)", "(x)", "(y)", ":[", "(n)", "(#)", "(r)", "({)"
                        Return "19*19"
                    Case "(})", "(t)", "(c)", "(i)", "(p)", "(^)", "(g)", "(k)", "(f)", "(w)", "(o)", ":-)"
                        Return "19*19"
                    Case Else
                        Select Case LCase(Left(sA, 4))
                            Case "<:o)", "(li)", "(sn)", "(tu)", "(pl)", "(||)", "(pi)", "(so)", "(au)", "(um)", "(ip)", "(co)", "(mp)", "(st)"
                                Return "19*19"
                            Case "(h5)", "(mo)", "(ap)"
                                Return "19*19"
                            Case Else
                                Select Case LCase(Left(sA, 5))
                                    Case "(brb)", "(bah)"
                                        Return "19*19"
                                    Case Else
                                        Return ""
                                End Select
                        End Select
                End Select
        End Select
    End Function

    'this looks if the text is used for emoticons, and if so put some code before the string
    Function EmoticonsText(sA As String) As String
        'this search for word for word
        Select Case LCase(Left(sA, 2))
            Case ":)", ":d", ":p", ";)", ":@", ":S", ":$", ":("
                ' in alle cases there is "ws" writen before the emoticon starts. i do not know if it is for the weight or for reconission.
                'if it is for reconission i have to use an unwriten char in stead of "ws"
                Return "ws" & Left(sA, 2)
            Case Else
                Select Case LCase(Left(sA, 3))
                    Case ":^)", ":-o", "(h)", ":'(", ":|", "(6)", "(a)", "8o|", "8-|", "+o(", "|-)", "*-)", ":-#", ":-*", "^o)", "8-)", "(l)", "(u)"
                        Return "ws" & Left(sA, 3)
                    Case "(m)", "(@)", "(&)", "(s)", "(*)", "(~)", "(8)", "(e)", "(b)", "(d)", "(z)", "(x)", "(y)", ":[", "(n)", "(#)", "(r)", "({)"
                        Return "ws" & Left(sA, 3)
                    Case "(})", "(t)", "(c)", "(i)", "(p)", "(^)", "(g)", "(k)", "(f)", "(w)", "(o)", ":-)"
                        Return "ws" & Left(sA, 3)
                    Case Else
                        Select Case LCase(Left(sA, 4))
                            Case "<:o)", "(li)", "(sn)", "(tu)", "(pl)", "(||)", "(pi)", "(so)", "(au)", "(um)", "(ip)", "(co)", "(mp)", "(st)"
                                Return "ws" & Left(sA, 4)
                            Case "(h5)", "(mo)", "(ap)"
                                Return "ws" & Left(sA, 4)
                            Case Else
                                Select Case LCase(Left(sA, 5))
                                    Case "(brb)", "(bah)"
                                        Return "ws" & Left(sA, 5)
                                    Case Else
                                        Return sA
                                End Select
                        End Select
                End Select
        End Select
    End Function
End Module