VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formss 
   Caption         =   "form"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "formss.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
Dim window As LongPtr
Dim ll As Long
Dim l As Long
window = GetActiveWindow()
ll = Module1.converter(window)
Me.Caption = "the active windows is " & Str(window)
End Sub
