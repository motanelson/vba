VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ball 
   Caption         =   "jumping ball"
   ClientHeight    =   5868
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9552.001
   OleObjectBlob   =   "ball.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
Dim X As Integer
Dim Y As Integer
Dim a As Integer
Me.Repaint
DoEvents
X = 10
Y = 10

For a = 0 To 1000
    Me.OptionButton1.Top = Me.OptionButton1.Top + Y
    Me.OptionButton1.Left = Me.OptionButton1.Left + X
    If Me.OptionButton1.Top > Me.Height - Me.OptionButton1.Height - 30 Then
        Y = -10
        Me.OptionButton1.Top = Me.OptionButton1.Top + Y
    End If
    If Me.OptionButton1.Left > Me.Width - Me.OptionButton1.Width - 30 Then
        X = -10
        Me.OptionButton1.Left = Me.OptionButton1.Left + X
    End If
    If Me.OptionButton1.Top < 10 Then
        Y = 10
        Me.OptionButton1.Top = Me.OptionButton1.Top + Y
    End If
    If Me.OptionButton1.Left < 10 Then
        X = 10
        Me.OptionButton1.Left = Me.OptionButton1.Left + X
    End If
    Me.Repaint
    DoEvents
    
    Sleep (1000)
Next
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

End Sub

Private Sub UserForm_Terminate()
End
End Sub
