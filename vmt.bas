Attribute VB_Name = "vmt"
Option Explicit
Sub vmtM()
Dim e As Integer
Dim f As Integer
Dim c As Integer
Dim myEntry As String
Dim baseLtr As String
Dim baseLtrChnk As String
Dim chunkyStr As String
Dim cL As String
Dim wb As Workbook
Set wb = ThisWorkbook
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Visual Memory Tablet")
Dim arrRev() As String
Dim inpt As String

With wb.Sheets("Visual Memory Tablet").Cells
    .ClearFormats
    .Clear
End With


myEntry = largeEntry()


'myEntry = InputBox("Enter string for rendition as VMT Notation", "Visual Memory Tablet")

Dim arr() As String
    arr = Split(myEntry)


Let f = 1
Let c = 1
Let e = 0


For e = 0 To UBound(arr)
    c = c + 1
    baseLtrChnk = UCase(Left(arr(e), 1))
    
    f = nextLine(e + 1, f)
    c = resetCols(e + 1, c)
    
   
    With wb.Sheets("Visual Memory Tablet").Range(Cells(f, c), Cells(f, c))
         .BorderAround xlContinuous, xlThick
         '.Offset(0, 0).Select
         .NumberFormat = "@"      'Text format
         .Value = arr(e) + ": " + convBaseLtr(baseLtrChnk) + strLengthdPLUS(arr(e), baseLtrChnk)
         .Columns.EntireColumn.AutoFit
    End With
Next e

Call borderRemoval

cL = wb.Sheets("Symbolization").Range("C32")
'cL = getComponentLevel("Letters")
wb.Sheets("Visual Memory Tablet").Range(Cells(f + 2, 1), Cells(f + 2, 1)).Value = cL


End Sub

Function largeEntry() As String
'    MsgBox (data)
    'TextBox.Caption = data
'    UserForm1.Show
    largeEntry = UserForm1.TextBox1.Value
    'largeEntry = TextBox.txtstuff.Text
End Function



Function strLengthdPLUS(entry As Variant, firstChar As String) As String
Dim chars As Long
Dim strLengthd As String
Dim alphaNumOnly As String
Dim regx As New RegExp

alphaNumOnly = replRegx(entry, "", "[^a-zA-Z\d]")    'regex for any character other than alphanumeric
Debug.Print alphaNumOnly

chars = Len(alphaNumOnly)

regx.Pattern = "[^a-zA-Z\d]"   'accounting for the special case where first character is non-alphanumeric
If regx.Test(firstChar) Then
chars = chars + 1 'as this function in the above, unfortunately, discounts ALL occurences of non-alphanumerics
End If

strLengthd = CStr(chars)

strLengthdPLUS = " +" + strLengthd

End Function

Function replRegx(strB As Variant, replace As String, regEx As String) As String
    Dim localRegEx As RegExp
    Set localRegEx = New RegExp
    localRegEx.MultiLine = True
    localRegEx.Pattern = regEx
    localRegEx.Global = True
    
    replRegx = localRegEx.replace(strB, replace)
End Function

Function borderRemoval()
Dim cell As Range
Dim rngB As Range
Set rngB = ThisWorkbook.Sheets("Visual Memory Tablet").UsedRange
For Each cell In rngB
    If Not IsEmpty(cell) Then
        cell.Borders(xlEdgeRight).LineStyle = xlNone
        cell.Borders(xlEdgeLeft).LineStyle = xlNone
    End If
Next cell
End Function

Function nextLine(wordNum As Integer, rowPos As Integer) As Integer
        Debug.Print "Word Count: " + CStr(wordNum)
        If (wordNum Mod 6 = 0) Then
           rowPos = Round(wordNum / 6, 0)
           rowPos = rowPos + 1
        End If
    Debug.Print "Row Number: " + CStr(rowPos)
    Debug.Print "-"
    nextLine = rowPos
    
End Function

Function resetCols(wordN As Integer, colPos As Integer) As Integer 'use this function to fix J1 blank space
   Debug.Print "Word Count: " + CStr(wordN)
   If (wordN Mod 6 = 0) Then
    colPos = 1
   End If
   Debug.Print "Column Number: " + CStr(colPos)
   Debug.Print "*******"
   resetCols = colPos
End Function

'Public Function getComponentLevel() As Dictionary
'    Dim component As Dictionary
'    Set component = New Dictionary
'
'    component.Add "Blank Surface", "[Nulla]TS  "
'    component.Add "Letters", "[1]TS  "
'    component.Add "Words", "[2]TS  "
'    component.Add "Sentences", "[3]TS  "
'    component.Add "Paragraphs", "[4]TS  "
'    component.Add "Pages", "[5]TS  "
'    component.Add "Sections", "[6]TS  "
'    component.Add "Chapters", "[7]TS  "
'    component.Add "Book", "[8]TS  "
'    component.Add "Libraries", "[9]TS  "
'
'    Set getComponentLevel = component
'
'End Function

Function convBaseLtr(letter As String) As String
Dim zns As String
Dim wrkbk As Workbook
Set wrkbk = ThisWorkbook

If letter = "A" Then
zns = wrkbk.Sheets("Symbolization").Range("C2").Value
ElseIf letter = "B" Then
zns = wrkbk.Sheets("Symbolization").Range("C3").Value
ElseIf letter = "C" Then
zns = wrkbk.Sheets("Symbolization").Range("C4").Value
ElseIf letter = "D" Then
zns = wrkbk.Sheets("Symbolization").Range("C5").Value
ElseIf letter = "E" Then
zns = wrkbk.Sheets("Symbolization").Range("C6").Value
ElseIf letter = "F" Then
zns = wrkbk.Sheets("Symbolization").Range("C7").Value
ElseIf letter = "G" Then
zns = wrkbk.Sheets("Symbolization").Range("C8").Value
ElseIf letter = "H" Then
zns = wrkbk.Sheets("Symbolization").Range("C9").Value
ElseIf letter = "I" Then
zns = wrkbk.Sheets("Symbolization").Range("C10").Value

'Since we are using only 9 Numbers to represent 26 Letters of English, therefore we need to put a ceiling AFTER "I" so the user can comprehend which Letter we are referring to, as
'the Zeta Numeral System repeats itself over the course of the 26 Lettrs.
ElseIf letter = "J" Then
zns = wrkbk.Sheets("Symbolization").Range("C12").Value
ElseIf letter = "K" Then
zns = wrkbk.Sheets("Symbolization").Range("C13").Value
ElseIf letter = "L" Then
zns = wrkbk.Sheets("Symbolization").Range("C14").Value
ElseIf letter = "M" Then
zns = wrkbk.Sheets("Symbolization").Range("C15").Value
ElseIf letter = "N" Then
zns = wrkbk.Sheets("Symbolization").Range("C16").Value
ElseIf letter = "O" Then
zns = wrkbk.Sheets("Symbolization").Range("C17").Value
ElseIf letter = "P" Then
zns = wrkbk.Sheets("Symbolization").Range("C18").Value
ElseIf letter = "Q" Then
zns = wrkbk.Sheets("Symbolization").Range("C19").Value
ElseIf letter = "R" Then
zns = wrkbk.Sheets("Symbolization").Range("C20").Value

ElseIf letter = "S" Then
zns = wrkbk.Sheets("Symbolization").Range("C22").Value
ElseIf letter = "T" Then
zns = wrkbk.Sheets("Symbolization").Range("C23").Value
ElseIf letter = "U" Then
zns = wrkbk.Sheets("Symbolization").Range("C24").Value
ElseIf letter = "V" Then
zns = wrkbk.Sheets("Symbolization").Range("C25").Value
ElseIf letter = "W" Then
zns = wrkbk.Sheets("Symbolization").Range("C26").Value
ElseIf letter = "X" Then
zns = wrkbk.Sheets("Symbolization").Range("C27").Value
ElseIf letter = "Y" Then
zns = wrkbk.Sheets("Symbolization").Range("C28").Value
ElseIf letter = "Z" Then
zns = wrkbk.Sheets("Symbolization").Range("C29").Value

ElseIf letter = "1" Then
zns = wrkbk.Sheets("Symbolization").Range("C2").Value
ElseIf letter = "2" Then
zns = wrkbk.Sheets("Symbolization").Range("C3").Value
ElseIf letter = "3" Then
zns = wrkbk.Sheets("Symbolization").Range("C4").Value
ElseIf letter = "4" Then
zns = wrkbk.Sheets("Symbolization").Range("C5").Value
ElseIf letter = "5" Then
zns = wrkbk.Sheets("Symbolization").Range("C6").Value
ElseIf letter = "6" Then
zns = wrkbk.Sheets("Symbolization").Range("C7").Value
ElseIf letter = "7" Then
zns = wrkbk.Sheets("Symbolization").Range("C8").Value
ElseIf letter = "8" Then
zns = wrkbk.Sheets("Symbolization").Range("C9").Value
ElseIf letter = "9" Then
zns = wrkbk.Sheets("Symbolization").Range("C10").Value

Else                                                        'if first character is unknown
zns = wrkbk.Sheets("Symbolization").Range("C31").Value

End If

convBaseLtr = zns
End Function

