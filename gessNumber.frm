VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} gessNumber 
   Caption         =   "UserForm1"
   ClientHeight    =   5916
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11004
   OleObjectBlob   =   "gessNumber.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "gessNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim trays As Integer

Private Sub CommandButton1_Click()
Dim nn As Integer
If trays = 0 Then ListBox1.Clear
trays = trays + 1
On Error Resume Next
nn = 50
nn = Val(TextBox1.Text)
If nn > n Then
    ListBox1.AddItem ("You number is big " + Str(trays))
End If
If nn < n Then
    ListBox1.AddItem ("You number is less " + Str(trays))
End If
If nn = n Then
    ListBox1.AddItem ("You win in " + Str(trays))
    trays = 0
    n = Rnd() * 100
End If
End Sub

Private Sub UserForm_Activate()
Randomize Time()
n = Rnd(100)
trays = 0
ListBox1.AddItem ("gess number from 0 to 100")
End Sub

Private Sub UserForm_Click()

End Sub
