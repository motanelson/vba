VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bars 
   Caption         =   "percent"
   ClientHeight    =   5712
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13392
   OleObjectBlob   =   "bars.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "bars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Me.Label1.Height = Me.Label1.Height + 10
If Me.Label1.Height > Me.Label2.Height Then Me.Label2.Height = Me.Label1.Height
Me.Label1.Top = Me.Label2.Top + Me.Label2.Height - Me.Label1.Height
Me.Label2.Caption = Str(Int(Me.Label2.Height / 575 * Me.Label1.Height)) + "% "
End Sub

Private Sub UserForm_Activate()
Me.Label2.Caption = Str(Int(Me.Label2.Height / 575 * Me.Label1.Height)) + "% "
End Sub

Private Sub UserForm_Click()

End Sub
