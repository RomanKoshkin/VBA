VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)

If Target.Column = 4 Then
beginning = Target.Cells.Offset(rowOffset:=1, columnOffset:=1).Address
a = Range(beginning).row
b = Range(beginning).Column
coun = 1
End If


If Target.Column = 10 And coun = 1 Then
Application.ScreenUpdating = False

ending = Target.Cells.Address
c = Range(ending).row
d = Range(ending).Column

Cells(c, d).Font.Underline = xlUnderlineStyleSingle
Cells(c, d).Font.ColorIndex = 1

Cells(a - 1, b - 1).Font.Underline = xlUnderlineStyleSingle
Cells(a - 1, b - 1).Font.ColorIndex = 1

ActiveSheet.Shapes.AddConnector(msoConnectorStraight, Range(beginning).Left, Range(beginning).Top, Range(ending).Left, Range(ending).Top).Select
Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadOpen
coun = 0


For Each cell In Range(Cells(2, 10), Cells(c, 10))
If Not cell.Font.Underline = xlUnderlineStyleSingle Then cell.Font.ColorIndex = 4
Next cell

For Each cell In Range(Cells(2, 4), Cells(a - 1, 4))
If Not cell.Font.Underline = xlUnderlineStyleSingle Then cell.Font.ColorIndex = 3
Next cell
Application.ScreenUpdating = True

'populate arrow matrix
Cells((Cells.Range(beginning).row - 1), 11) = Cells.Range(beginning).row - 1
Cells((Cells.Range(beginning).row - 1), 12) = Cells.Range(ending).row

End If

End Sub
