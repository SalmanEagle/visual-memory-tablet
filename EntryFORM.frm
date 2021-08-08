VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryFORM 
   Caption         =   "Visual Memory Tablet"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13830
   OleObjectBlob   =   "EntryFORM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EntryFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub CommandButton1_Click()
Dim enteredData As String

enteredData = Me.TextBox1.Value
Call vmtM

End Sub

Public Sub CommandButton2_Click()
Unload Me
End Sub



Public Sub Label1_Click()

End Sub



Public Sub TextBox1_Change()

End Sub
