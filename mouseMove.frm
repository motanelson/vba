VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mouseMove 
   Caption         =   "mouse move"
   ClientHeight    =   5976
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13992
   OleObjectBlob   =   "mouseMove.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "mouseMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Click()

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.OptionButton1.Top = Y - (Me.OptionButton1.Height / 2)
Me.OptionButton1.Left = X - (Me.OptionButton1.Width / 2)
End Sub
