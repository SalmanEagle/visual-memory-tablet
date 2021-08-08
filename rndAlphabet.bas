Attribute VB_Name = "rndAlphabet"
Sub randomAlphabet()
Dim val As Integer
Randomize

Dim iter As Integer

Dim wb As Workbook
Set wb = ThisWorkbook

Dim iterWS As Worksheet

Dim wsPRAC As Worksheet
Set wsPRAC = ThisWorkbook.Worksheets("Practice")

Dim rngA As Range
Set rngA = wsPRAC.Range("A2:A27")
Dim rngB As Range
Set rngB = wsPRAC.Range("B2:B27")

Dim wsSOL As Worksheet
Set wsSOL = ThisWorkbook.Worksheets("Solutions")

Dim rngX As Range
Set rngX = wsSOL.Range("A2:A27")
Dim rngY As Range
Set rngY = wsSOL.Range("C2:C27")

With wb.Sheets("Practice").Cells
    .ClearFormats
    .Clear
End With

With wb.Sheets("Solutions").Cells
    .ClearFormats
    .Clear
End With

With wsPRAC.Range("A1")
    .Font.Color = vbRed
    .Font.Bold = True
    .Value = "Take printouts, and repeatedly challenge yourself to figure out the VM Notation for random letters:"
End With

For a = 1 To 26
val = Int((90 - 65 + 1) * Rnd + 65)    'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    With wb.Sheets("Practice").Range(Cells(a + 1, 1), Cells(a + 1, 1))
        .Font.Bold = True
        .Font.Size = 20.5
        .Value = Chr(val) + ": "
        .Columns.ColumnWidth = 45
    End With
Next a

Dim aa As Integer
For aa = 1 To 26
val = Int((90 - 65 + 1) * Rnd + 65)    'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    With wb.Sheets("Practice").Range(Cells(aa + 1, 2), Cells(aa + 1, 2))
        .Font.Bold = True
        .Font.Size = 20.5
        .Value = Chr(val) + ": "
        .Columns.ColumnWidth = 45
    End With
Next aa


With wb.Sheets("Solutions").Range("A1")
    .Value = "Answers to Letters in PRACTICE:"
    .Font.ColorIndex = 10
    .Font.Bold = True
    .Columns.EntireColumn.AutoFit
End With

wb.Sheets("Practice").Range("A2:A27").Copy Destination:=wb.Sheets("Solutions").Range("A2")
wb.Sheets("Solutions").Range("A2:A27").Columns.ColumnWidth = 8

wb.Sheets("Practice").Range("B2:B27").Copy Destination:=wb.Sheets("Solutions").Range("C2")
wb.Sheets("Solutions").Range("C2:C27").Columns.ColumnWidth = 8

Dim b As Integer
Dim baseLtrChnk2 As String
Dim converted2VMN As String

For b = 1 To 26
baseLtrChnk2 = Left(rngX(b, 1).Value, 1)
converted2VMN = conv2VMN(baseLtrChnk2)
rngX(b, 2).Value = converted2VMN
rngX(b, 2).Font.Size = 13
rngX(b, 2).Font.Bold = True
Next b
wb.Sheets("Solutions").Range("B2:B27").Columns.ColumnWidth = 37

Dim bb As Integer
Dim baseLtrChnk3 As String
Dim converted2VMN2

For bb = 1 To 26
baseLtrChnk3 = Left(rngY(bb, 1).Value, 1)
converted2VMN2 = conv2VMN(baseLtrChnk3)
rngY(bb, 2).Value = converted2VMN2
rngY(bb, 2).Font.Size = 13
rngY(bb, 2).Font.Bold = True
Next bb
wb.Sheets("Solutions").Range("D2:D27").Columns.ColumnWidth = 37
wb.Sheets("Solutions").Range("A1:D1").Merge

End Sub




