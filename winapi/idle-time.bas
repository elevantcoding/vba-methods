'------------------------------------------------------------
' Windows Machine Idle Time (seconds)
'
' This routine uses Win32 API calls to retrieve the amount of
' time (in seconds) since the last user input at the OS level.
'
' Original source: Unknown (found and adopted in prior work).
' Retained and used extensively for session control and
' graceful application shutdown in production environments.
'
' If you recognize the original source, please let me know
' and I will happily add proper attribution.
'------------------------------------------------------------

Option Explicit

Private Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type

Private Declare PtrSafe Sub GetLastInputInfo Lib "user32" (ByRef plii As LASTINPUTINFO)
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Function IdleTime() As Single
    On Error GoTo Except

    Dim A As LASTINPUTINFO
    A.cbSize = LenB(A)
    GetLastInputInfo A
    IdleTime = (GetTickCount - A.dwTime) / 1000
  
Except:
  IdleTime = 0
  
End Function

Function PrintIdleTime()

  PrintIdleTime = 0
    
    If IdleTime > 0 Then
        PrintIdleTime = IdleTime
    End If
    
End Function
