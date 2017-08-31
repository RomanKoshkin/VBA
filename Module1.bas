Attribute VB_Name = "Module1"
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'VIEW READMEs IN THE KOSH_EEG2 FOLDER FOR THE WHOLE SEQUENCE
' #######################################################
Public coun As Integer
Public beginning As String
Public ending As String
Public a, b, c, d As Long
Private Sub Entry()
Call ComputeArrowOnAndOffSets
Call RedChk
End Sub
Function DetLastRow()
'don't delete this function, it returns the size of data
Dim maxval As Double

maxval1 = Application.WorksheetFunction.Max(Range("C:C"))
maxval2 = Application.WorksheetFunction.Max(Range("I:I"))
maxval = WorksheetFunction.Max(maxval1, maxval2)

If maxval = maxval1 Then
    DetLastRow = Application.WorksheetFunction.Match(maxval1, Range("C:C"))
Else
    DetLastRow = Application.WorksheetFunction.Match(maxval2, Range("I:I"))
End If

End Function

Sub ComputeArrowOnAndOffSets()
lastrow = DetLastRow()

Set MyDocument = ActiveSheet
MyDocument.Shapes(1).Line.DashStyle = msoLineDashDot
ReDim aArray(lastrow, 1) As Integer
For i = 1 To MyDocument.Shapes.count
MyDocument.Shapes(i).Line.DashStyle = msoLineSolid
aArray(i, 0) = MyDocument.Shapes(i).TopLeftCell.row - 1
aArray(i, 1) = MyDocument.Shapes(i).BottomRightCell.row
Cells(aArray(i, 0), 11) = aArray(i, 0)
Cells(aArray(i, 0), 12) = aArray(i, 1)
Next i

'Sub Proc2()
Application.ScreenUpdating = True
For i = 1 To lastrow
        Application.StatusBar = i
        
        If Cells(i, 4).Font.Bold = False And Cells(i, 4).Font.ColorIndex = 1 And Cells(i, 4).Font.Underline = 2 Then
            Cells(i, 13) = 1
        Else
            Cells(i, 13) = 0
        End If
        
        If Cells(i, 10).Font.Bold = False And Cells(i, 10).Font.ColorIndex = 1 And Cells(i, 10).Font.Underline = 2 Then
            Cells(i, 14) = 1
        Else
            Cells(i, 14) = 0
        End If

Next i
Application.StatusBar = ""
Call OPQ
End Sub
Private Sub OPQ()
Set MyDocument = ActiveSheet
w = MyDocument.Shapes.count
w = MyDocument.Shapes(w).BottomRightCell.row
r = Range(Cells(1, 15).Address, Cells(w, 15).Address).Address
Range(r).FormulaR1C1 = "=SUM(R1C[-2]:RC[-2])"
r = Range(Cells(1, 16).Address, Cells(w, 16).Address).Address
Range(r).FormulaR1C1 = "=SUM(R1C[-2]:RC[-2])"
r = Range(Cells(1, 17).Address, Cells(w, 17).Address).Address
Range(r).FormulaR1C1 = "=RC[-2] - RC[-1]"

End Sub
Private Sub VW()
Set MyDocument = ActiveSheet
w = MyDocument.Shapes.count
w = MyDocument.Shapes(w).BottomRightCell.row

r = Range(Cells(1, 22).Address, Cells(w, 22).Address).Address
Range(r).FormulaR1C1 = "=SUM(R2C[-2]:RC[-2]) + SUM(R2C[-1]:RC[-1])"
'=SUM(T2:T$2)+SUM(U$2:U2)
r = Range(Cells(1, 23).Address, Cells(w, 23).Address).Address
Range(r).FormulaR1C1 = "=RC[-6]+RC[-1]"

End Sub
Sub FinalReformat()

'MAKE SURE THERE'S NO HEADER ROW!!!! THE FIRST WORD SHOULD BE IN ROW 1, NOT 2!!!!

leng = DetLastRow()
Range(Cells(1, 3), Cells(leng, 4)).Copy Range(Cells(1, 25), Cells(leng, 26))

For Each cell In Range(Cells(1, 27), Cells(leng, 27))
t1 = cell.row
t2 = cell.Column
cell.Value = Cells(t1, t2 - 14).Value + Cells(t1, t2 - 7).Value
Next cell

Range(Cells(1, 9), Cells(leng, 10)).Copy Range(Cells(1, 28), Cells(leng, 29))

For Each cell In Range(Cells(1, 30), Cells(leng, 30))
t1 = cell.row
t2 = cell.Column
cell.Value = -1 * Cells(t1, t2 - 16).Value + Cells(t1, t2 - 9).Value
Next cell


Range(Cells(1, 28), Cells(leng, 28)).NumberFormat = "hh:mm:ss.000"
Set rang = Range(Cells(1, 31), Cells(leng, 31))
rang.NumberFormat = "hh:mm:ss.000"
rang.FormulaR1C1 = "=IF(RC[-6]="""", RC[-3], RC[-6])"

Set rang = Range(Cells(1, 33), Cells(leng, 33))
rang.NumberFormat = "Text"
rang.FormulaR1C1 = "=RC[-7]&RC[-4]"

Set rang = Range(Cells(1, 34), Cells(leng, 34))
rang.FormulaR1C1 = "=SUM(R1C[-7]:RC[-7]) + SUM(R1C[-4]:RC[-4])"
rang.NumberFormat = "General"

Range(Cells(2, 31), Cells(leng + 1, 31)).Copy
Range(Cells(1, 32), Cells(leng, 32)).PasteSpecial xlPasteValues

Range(Cells(1, 31), Cells(leng, 34)).Copy
Range(Cells(1, 25), Cells(leng, 28)).PasteSpecial xlPasteValuesAndNumberFormats

Range(Cells(1, 29), Cells(leng, 34)).Clear

Set rang = Range(Cells(1, 25), Cells(leng, 26))
rang.NumberFormat = "hh:mm:ss.000"



'Set rang = Range(Cells(1, 31), Cells(leng, 31))
'Range(Cells(1, 28), Cells(leng, 30)).Copy Range(Cells(leng + 1, 25), Cells(2 * leng, 27)) '!!!!!!!!!
'Range(Cells(1, 28), Cells(leng, 30)).Clear
'Range(Cells(1, 25), Cells(1 + leng * 2, 27)).Sort Key1:=Range("Y1"), Order1:=xlAscending, Header:=xlYes
'n = Range("Y:Y").Cells.SpecialCells(xlCellTypeConstants).count
'r = Range(Cells(1, 28), Cells(n, 28)).Address
'Range(r).FormulaR1C1 = "=SUM(R1C[-1]:RC[-1])"
'Range("Z:Z").EntireColumn.Insert
'r = Range(Cells(1, 26), Cells(n, 26)).Address
'Range(r).FormulaR1C1 = "=R[+1]C[-1]"

'Set r = Range(Cells(1, 29), Cells(n, 29))
'Set l = Range(Cells(1, 28), Cells(n, 28))
'l.Value = r.Value
'Range(Cells(1, 29), Cells(n, 29)).Clear

For Each cell In Range(Cells(1, 29), Cells(leng, 29))
t1 = cell.row
t2 = cell.Column
x = Cells(t1, t2 - 4).Value * 86400
cell.Value = Cells(t1, t2 - 4).Value * 86400
Next cell

For Each cell In Range(Cells(1, 30), Cells(leng, 30))
t1 = cell.row
t2 = cell.Column
cell.Value = Cells(t1, t2 - 4).Value * 86400
Next cell


End Sub
Sub NewProc() 'This procedure and (insert_time_shifting_down2)
                        'time-syncs two time-coded colums
                        'ORIGINAL IN COL A-D, TRANSLATION IN COL G-J
                        
Application.ScreenUpdating = False
lim = DetLastRow()
Application.StatusBar = j


i = 1

While Cells(i, 3) <> "" And Z < 30 'Z < 30 is the stop condition
If Cells(i + Z, 3) < Cells(i, 9) Then
        Z = Z + 1
    Else
        f = insert_cells_shifting_down_2(7, i, Z)
        i = i + Z
        f = insert_cells_shifting_down_2(1, i, 1)
        i = i + 1
        Z = 0
    End If
Wend

Application.ScreenUpdating = True
End Sub
Function insert_cells_shifting_down_2(bgn, j, i)
For n = 1 To i
    Range(Cells(j, bgn), Cells(j, bgn + 5)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Next n
End Function
Sub RemoveGrammaticalWords()

'DECLARATIONS
Const ColNo = 6 ' sets the number of columns in WrdArray
Dim ColLen(ColNo) As Integer

' get number of words in each column in WrdArray
For i = 0 To ColNo
Set r = Sheets("WrdArray").Columns(i + 1)
ColLen(i) = WorksheetFunction.CountA(r) - 1
Next i

'word array loading
Dim WrdArray(ColNo, 19) As String
For i = 0 To ColNo
For j = 0 To 19
WrdArray(i, j) = Sheets("WrdArray").Cells(j + 2, i + 1).Text
Next j
Next i

'weeding out shit
For j = 0 To ColNo 'get column number in WrdArray
For k = 0 To ColLen(j) - 1 'get row number in WrdArray
For i = 1 To 2000 'get row number in target text
If Cells(i, 4) = WrdArray(j, k) Then
    Cells(i, 4).Activate
    Cells(i, 4) = Empty
    End If
If Cells(i, 10) = WrdArray(j, k) Then
    Cells(i, 10).Activate
    Cells(i, 10) = Empty
    End If

Next i

Next k
Next j
End Sub
Sub CompressionOmissionAdditionCheck()
For i = 2 To 1850
If Cells(i, 4).Font.ColorIndex = 3 And Not Cells(i, 4) = Empty Then Cells(i, 16).Value = Cells(i, 16).Value + 1
If Cells(i, 10).Font.ColorIndex = 4 And Not Cells(i, 10) = Empty Then Cells(i, 16).Value = Cells(i, 16).Value - 1
Next i
End Sub
Sub ArrowTracker()
    
    Set v = ActiveSheet.Shapes
    
    For c = 1 To v.count
    br = v.Item(c).BottomRightCell.row
    bc = v.Item(c).BottomRightCell.Column
    tr = v.Item(c).TopLeftCell.row - 1
    tc = v.Item(c).TopLeftCell.Column - 1
    Cells(br, bc).Select
    Cells(tr, tc).Select
    For i = tr To br
    acc = acc + Cells(i, 6)
    Next i
    Cells(br, 85) = acc
    acc = 0
    Next c
End Sub
Sub CountRedOrphansInSentence()
a = 1
For i = 1 To 2000 'set the colum length here
Set v = ActiveSheet.Cells(i, 10)
If v.Interior.ColorIndex = 41 Then
    For q = a + 1 To i
    Cells(q, 4).Select
    If Cells(q, 4).Font.ColorIndex = 3 And Not Cells(q, 4).Text = "" Then count = count + 1
    Next q
    Cells(i, 21) = count
    a = i
    count = 0
End If
Next i
End Sub
Sub CorrectWrongRedColorInColumnD()
Set v = Cells(796, 4)
For i = 1 To 2000
If Not Cells(i, 4).Font.ColorIndex = 3 And Not Cells(i, 4).Text = "" And Not Cells(i, 4).Font.Underline = 2 Then
Cells(i, 4).Select
End If
Next i
End Sub
Sub Find_First() 'this code finds lost non-meaningful words

For n = 1 To 2000
If Cells(n, 6) = "" And Not Cells(n, 5) = "" Then
m = Cells(n, 5)
   
        With Sheets("Sheet1").Range("A:A")
            Set Rng = .Find(What:=m, _
                            After:=.Cells(.Cells.count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                a = Cells(Rng.row, 4)
            Else
                MsgBox "Nothing found"
            End If
        End With
Cells(n, 6) = a
End If
Next n
End Sub
Sub CalcRealDur() 'this code is a workaround for an already processed piece of data
For i = 2 To 1814
If Not Cells(i, 1) = "" Then
x = Cells(i, 1)
    If Not Cells(i + 1, 1) = "" Then
        Y = Cells(i + 1, 1)
        Cells(i, 5) = Y - x
        Else
        For o = 1 To 10
            If Not Cells(i + 1 + o, 1) = "" Then
                Y = Cells(i + 1 + o, 1)
                Cells(i, 5) = Y - x
                Exit For
            End If
        Next o
    End If
End If
Next i
End Sub
'Set the target cell's value to 1 if the font in the referenced cell is UNDERLINED
'Function ChUn(Rng1 As Range)
'Application.ScreenUpdating = False
 '       If Rng1.Font.Bold = False And Rng1.Font.ColorIndex = 1 And Rng1.Font.Underline = 2 Then
   '         ChUn = 1
  '      Else
    '        ChUn = 0
     '   End If
'End Function


'set cell's value to 1 if the font color in the referenced cell is RED
Sub RedChk()
'MAKE SURE THERE'S NO HEADER ROW!!!! THE FIRST WORD SHOULD BE IN ROW 1, NOT 2!!!!
Application.StatusBar = False
'first, let's find out the length of the column to work with:
lastrow = DetLastRow()
'now do the important part:
For i = 1 To lastrow
If Cells(i, 4).Font.Bold = False And Cells(i, 4).Font.ColorIndex = 3 And Not Cells(i, 4).Text = "" Then
    Cells(i, 20) = 1
    For q = 1 To 20
        If Cells(i + q, 4).Font.Underline = 2 Then
            t = Cells(i + q, 12)
            Cells(t, 21) = Cells(t, 21) - 1
            Exit For
        End If
    Next q
Else
    Red = 0
End If
Next i
Application.StatusBar = False
Call VW
Call FinalReformat
End Sub
Sub Replace()

'The word document must be closed. The MUST BE NO CAPS in the document and NO PARAGRAPHS.

Dim pathh As String
Dim pathhi As String
Dim oCell  As Integer
Dim from_text As String, to_text As String
Dim WA As Object

pathh = "Users:RomanKoshkin:Documents:MATLAB:KOSH_EEG2:test.docx"

Set WA = CreateObject("Word.Application")
WA.Documents.Open (pathh)
WA.Visible = True

For oCell = 2 To 31
    from_text = Sheets("WrdArray").Range("K" & oCell).Value
    from_text1 = " " & from_text & " "
    to_text = " " & from_text & "¤"
    With WA.ActiveDocument
        Set myRange = .Content
        With myRange.Find
            .Execute FindText:=from_text1, ReplaceWith:=to_text, Replace:=2
        End With
    End With
Next oCell

With WA.ActiveDocument
       Set myRange = .Content
       With myRange.Find
            .Execute FindText:=" ", ReplaceWith:="^p!^p", Replace:=2
       End With
End With
    
With WA.ActiveDocument
        Set myRange = .Content
        With myRange.Find
            .Execute FindText:="¤", ReplaceWith:=" ", Replace:=2
        End With
End With

End Sub
Sub FindEmptyUnderlined()
'choose row 4 or 10
longi = DetLastRow()
For i = 1 To longi
Set wer = Cells(i, 10)
If wer.Font.Underline = 2 And wer.Text = "" Then
MsgBox (i)
End If
Next i
End Sub
Sub RS() 'Count syllable load
ReDim a(1 To Application.WorksheetFunction.count(Range("T:T"))) As Integer
ReDim b(1 To Application.WorksheetFunction.count(Range("U:U"))) As Integer
lim = DetLastRow()
ax = 1
bx = 1
For i = 1 To lim
    If Cells(i, 20) = 1 Then
        a(ax) = i
        ax = ax + 1
    End If
    If Cells(i, 21) < 0 Then
        b(bx) = i
        bx = bx + 1
    End If
Next i

ReDim aCS(1 To UBound(a))
For i = 1 To UBound(a)
aCS(i) = CS(Cells(a(i), 4))
Cells(a(i), 18) = aCS(i)
Cells(a(i), 18).Font.ColorIndex = 4
Next i

ReDim bCS(1 To UBound(b))
i = 0
disp = 0
disp2 = 0
While i < UBound(b)
i = i + 1

    disp = -Cells(b(i), 21) - 1
    temp1 = aCS(i + disp2)
    temp2 = aCS(i + disp2 + disp)
    temp3 = a(i + disp2)
    temp4 = a(i + disp2 + disp)
    sigma = Application.WorksheetFunction.Sum(Range(Cells(temp3, 18), Cells(temp4, 18)))
    Cells(b(i), 19) = -1 * (sigma)
    Cells(b(i), 19).Font.ColorIndex = 4
    Cells(b(i), 19).Select
    disp2 = disp2 + disp

Wend

Jump1:
i = 1
While i <= lim
        
        If Cells(i, 4) <> "" And Cells(i, 4).Font.Underline = xlUnderlineStyleSingle Then
                ent = Cells(i, 11)
                ext = Cells(i, 12)
                SC = CS(Cells(i, 4))
                Cells(ent, 31) = SC
                Cells(ext, 32) = -SC
                i = i + 1
        Else:
                i = i + 1
        End If
Wend
Call SylC
'Call CommaToDot 'if you need to replace commas with dots in numbers
Call Enter_Values
End Sub
Private Function CS(ByVal daTxt As String) As Long
    Dim x As Long
    Dim l As Long
    CS = 0
    'Vowels
    qqq = Len(daTxt)
    For x = 1 To Len(daTxt)
        Select Case UCase(Mid(daTxt, x, 1))
            Case "A", "E", "I", "O", "U", "Y"
                CS = CS + 1
            Case Else
        End Select
    Next
    'Dipthongs
    For l = 1 To Len(daTxt)
        Select Case UCase(Mid(daTxt, l, 2))
            Case _
            "AA", "AE", "AI", "AO", "AU", "AY", _
            "EA", "EE", "EI", "EO", "EU", "EY", _
            "IA", "IE", "II", "IO", "IU", "IY", _
            "OA", "OE", "OI", "OO", "OU", "OY", _
            "UA", "UE", "UI", "UO", "UU", "UY", _
            "YA", "YE", "YI", "YO", "YU", "YY"
            
                CS = CS - 1
            Case Else
        End Select
    Next
    If CS > 1 And UCase(Right(daTxt, 1)) = "E" Then
        Select Case UCase(Mid(daTxt, Len(daTxt) - 1, 1))
            '// Check if the second to last letter is a vowel
            Case "A", "E", "I", "O", "U", "Y"
                Exit Function
            '// Not a vowel silent E
            Case Else
                CS = CS - 1
        End Select
    End If
    
      'count Russian syllables
  
  For i = 1 To Len(daTxt)
If Mid(daTxt, i, 1) = ChrW(1105) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1091) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1077) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1099) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1072) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1086) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1101) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1103) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1080) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1102) Then CS = CS + 1

If Mid(daTxt, i, 1) = ChrW(1025) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1059) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1045) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1067) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1040) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1054) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1069) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1071) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1048) Then CS = CS + 1
If Mid(daTxt, i, 1) = ChrW(1070) Then CS = CS + 1
Next i
End Function

Private Sub SylC()
lim = DetLastRow()
r = Range(Cells(1, 33).Address, Cells(lim, 33).Address).Address
Range(r).FormulaR1C1 = "=RC[-15] + RC[-2]"
r = Range(Cells(1, 34).Address, Cells(lim, 34).Address).Address
Range(r).FormulaR1C1 = "=RC[-15] + RC[-2]"
r = Range(Cells(1, 35).Address, Cells(lim, 35).Address).Address
Range(r).FormulaR1C1 = "=SUM(R1C[-2]:RC[-2]) + SUM(R1C[-1]:RC[-1])"
End Sub
Private Sub CommaToDot()
For i = 1 To 8

Application.Sheets(i).Select
Range("AC:AD").Select
Selection.Replace What:=",", Replacement:=".", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
Next i
End Sub
Sub Enter_Values()
    'For i = 1 To 8
    'Application.Sheets(i).Select
    Set ww = Range("AC1:AD2000")
    ww.NumberFormat = "0.000"
    For Each xCell In ww
        xCell.Value = xCell.Value
    Next xCell
    'Next i

'If Application.ActiveSheet.Index = 8 Then
'    ActiveWorkbook.SaveAs Filename:="1.xlsx", FileFormat:=51
'End If

End Sub

Function LoadArrayOfWordFrequencies() As Variant
x = ActiveSheet.Name
Worksheets("weighted dictionary").Activate
TotalRows = Rows(Rows.count).End(xlUp).row

Dim token() As Variant
Dim Frequency() As Double
ReDim token(TotalRows, 1)
For i = 2 To TotalRows
    token(i, 0) = Cells(i, 1).Value
    token(i, 1) = Cells(i, 5).Value
Next i
LoadArrayOfWordFrequencies = token
Worksheets(x).Activate
End Function
Function GetFrequency(ByVal Valu As String, ByVal f As Variant, ByVal w As Variant)
b = Split(Valu, " ")
    frAcc = 0
    For Each Item In b
        pos = Application.Match(Item, w, False)
            If Not IsError(pos) Then
                a = f(pos, 1)
                frAcc = frAcc + f(pos, 1)
            Else
                errorCount = errorCount + 1
            End If
    Next
    GetFrequency = frAcc
End Function

Sub GetCumFrequencyOfLine()
' COLUMNS AK AND AL
Dim x As Variant
x = LoadArrayOfWordFrequencies()
w = Application.Index(x, 0, 1)
f = Application.Index(x, 0, 2)
errorCount = 0
For i = 1 To DetLastRow
    If Cells(i, 4).Font.Underline = 2 Then
        frAcc = GetFrequency(Cells(i, 4).Value, f, w)
        Cells(i, 37) = frAcc                                    'AK
        Cells(Cells(i, 12).Value, 38) = -frAcc         'AL
    Else
    End If
Next i
MsgBox (errorCount)
RedSEEK
End Sub

Sub RedSEEK()
Range("AM:AM").Cells.Value = Empty
Dim x As Variant
x = LoadArrayOfWordFrequencies()
w = Application.Index(x, 0, 1)
f = Application.Index(x, 0, 2)

For i = 1 To DetLastRow
If Cells(i, 20) = 1 Then
r = GetFrequency(Cells(i, 4).Value, f, w)
Cells(i, 39) = r
End If
Next i
RedSEEK2
End Sub

Sub RedSEEK2()

Range("AN:AN").Cells.Value = Empty
lastrow = 1
For i = 2 To DetLastRow
If Cells(i, 21) < 0 Then
    count = -Cells(i, 21)
    For m = 1 To count
    'Range(Cells(lastrow, 20), Cells(i, 20)).Select
    
    With ActiveSheet.Range(Cells(lastrow, 20), Cells(i, 20))
    Set tobj = .Find(1, LookIn:=xlValues, SearchOrder:=xlByColumns)
        If Not tobj Is Nothing Then
            lastrow = tobj.row
        End If
     End With
     Cells(i, 40) = Cells(i, 40) - Cells(lastrow, 39)
     Next m
End If
Next i
r = Range(Cells(1, 42).Address, Cells(DetLastRow, 42).Address).Address
Range(r).FormulaR1C1 = "=SUM(R1C[-5]:RC[-4])"
r = Range(Cells(1, 43).Address, Cells(DetLastRow, 43).Address).Address
Range(r).FormulaR1C1 = "=SUM(R1C[-6]:RC[-3])"
End Sub

Sub CorrectCLred()
Dim rngData As Range
Dim er As Variant
Dim RE As Variant
Dim REidx() As Integer
Dim eridx() As Integer
Application.Calculation = xlCalculationManual

ShNames = Array("Colombia", "CostaRica", "Chile", "Morocco", "Peru", "Uruguay", "France", "Honduras")
lang = Array("er", "er", "er", "er", "RE", "RE", "RE", "RE")
e = 0
r = 0
c = 0
d = 0

For i = 1 To 8
    b = Split(ActiveWorkbook.Sheets(i).Name, "#")
    pos = Application.Match(b(0), ShNames, False) - 1
    language = lang(pos)
    If language = "er" Then
        e = e + Application.WorksheetFunction.Sum(Worksheets(i).Range("AK:AK, AM:AM"))
        ReDim Preserve eridx(d)
        eridx(d) = i
        d = d + 1
    Else:
        r = r + Application.WorksheetFunction.Sum(Worksheets(i).Range("AK:AK, AM:AM"))
        ReDim Preserve REidx(c)
        REidx(c) = i
        c = c + 1
    End If
              
Next
multiplier = e / r

For Each i In REidx
Set Rng = ActiveWorkbook.Sheets(i).Range("AK1:AN4000")
    For Each c In Rng.Cells
        c.Value = c.Value * multiplier
    Next
    Worksheets(i).Activate
    ActiveWorkbook.Sheets(i).Range("AK1:AN4000").Select
    Selection.Replace What:=0, Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False
Next
Application.Calculation = xlCalculationAutomatic
End Sub



