Attribute VB_Name = "clarify"
Public Function rmvSeparator(enter As String)

'if nothing is entered then prompt user
If enter = "" Then
MsgBox ("You forgot to enter anything.")
End If

'remove Separator if it is the only entry, e.g. "/"
If Left(enter, 1) = "/" Then
MsgBox ("/ is Separator, it cannot be the only entry.")
EntryFORM.TextBox1.Value = ""

'remove Separator if it is the last character withouth anything to the right, e.g. "s/"
ElseIf Right(enter, 1) = "/" Then
MsgBox ("/ is Separator, it cannot be the right most character; it is always put in between.")
EntryFORM.TextBox1.Value = ""

End If
End Function
