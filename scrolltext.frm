VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} scrolltext 
   Caption         =   "scrolltext"
   ClientHeight    =   6132
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14184
   OleObjectBlob   =   "scrolltext.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "scrolltext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub UserForm_Activate()
Dim X As Integer
Dim Y As Integer
Dim a As Integer
Me.Repaint
DoEvents
X = -10
Y = 10

For a = 0 To 1000
    Me.OptionButton1.Left = Me.OptionButton1.Left + X
    If Me.OptionButton1.Left < -(Me.OptionButton1.Width) Then
        X = -10
        Me.OptionButton1.Left = Me.Width
        
    End If
    Me.Repaint
    DoEvents
    
    Sleep (100)
Next
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Terminate()
End
End Sub
