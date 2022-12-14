VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sizes 
   Caption         =   "sizes"
   ClientHeight    =   6012
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14496
   OleObjectBlob   =   "sizes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sizes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OptionButton1_Click()

End Sub

Private Sub UserForm_Activate()
Dim X As Integer
Dim Y As Integer
Dim a As Integer
Me.Repaint
DoEvents
X = 60
Y = 10
f = 3
Me.OptionButton1.AutoSize = False
For a = 0 To 1000
    Me.OptionButton1.Height = Me.OptionButton1.Height + Y
    Me.OptionButton1.Width = Me.OptionButton1.Width + X
    Me.OptionButton1.Font.Size = Me.OptionButton1.Font.Size + f
    Me.OptionButton1.Left = (Me.Width - Me.OptionButton1.Width) / 2
    Me.OptionButton1.Top = (Me.Height - Me.OptionButton1.Height) / 2
    
    If Me.OptionButton1.Height > Me.Height Then
        Y = -10
        X = -10
        f = -3
    End If
    If Me.OptionButton1.Height < 30 Then
        Y = 10
        X = 10
        f = 3
    End If
    
    Me.Repaint
    DoEvents
    
    Sleep (1000)
    
Next
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Terminate()
End

End Sub
