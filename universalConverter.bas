Attribute VB_Name = "universalConverter"
Public Function conv2VMN(entered As String) As String
Dim baseLetter As String
Dim Length As Integer
Dim strLength As String

baseLetter = UCase(Left(entered, 1))

conv2VMN = convBaseLtr(baseLetter) + strLengthdPLUS(entered, baseLetter)
End Function



