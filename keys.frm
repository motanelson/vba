VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} keys 
   Caption         =   "keys"
   ClientHeight    =   6648
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14952
   OleObjectBlob   =   "keys.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "keys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyLeft Then Image1.Left = Image1.Left - 10
If KeyCode = vbKeyRight Then Image1.Left = Image1.Left + 10
If Image1.Left < -(Image1.Width / 2) Then Image1.Left = -(Image1.Width / 2)
If Image1.Left > (Image1.Width / 2) Then Image1.Left = (Image1.Width / 2)
End Sub
