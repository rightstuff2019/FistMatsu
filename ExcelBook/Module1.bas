Attribute VB_Name = "Module1"
Declare Function CreateFile Lib "KERNEL32" Alias "CreateFileA" _
    (ByVal filename As String, ByVal rw As Long, ByVal d1 As Long, ByVal d2 As Long, ByVal d3 As Long, ByVal d4 As Long, ByVal d5 As Long) As Long
    
Declare Sub ReadFile Lib "KERNEL32" _
    (ByVal handle As Long, ByVal buf As String, ByVal bytes As Long, readbytes As Long, ByVal d1 As Long)
    
Declare Sub CloseHandle Lib "KERNEL32" (ByVal handle As Long)

Declare Sub WriteFile Lib "KERNEL32" _
    (ByVal handle As Long, ByVal buf As String, ByVal bytes As Long, writebytes As Long, ByVal d1 As Long)
    
Declare Sub SetCommTimeouts Lib "KERNEL32" (ByVal handle As Long, ct As COMMTIMEOUTS)

'CreateFile parameta-
Const GENERIC_READ = (&H8000000)
Const GENERIC_WRITE = (&H40000000)



Sub Arduino1()

End Sub
