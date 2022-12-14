Attribute VB_Name = "Module1"
    Declare PtrSafe Function GetActiveWindow Lib "User32" () As LongPtr
    Declare PtrSafe Function Rectangle Lib "GDI32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    Public Function converter(l As LongPtr) As Long
        converter = Val(Str(l))
    End Function
    
