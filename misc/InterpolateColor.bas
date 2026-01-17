' fade in a welcome or other message on a form:
' start with white (or adjust for form background color) and adjust from that color to the intended color in R, G, B
' requires form timer (see form module code below)

Public Function InterpolateColor(ByVal R As Long, ByVal G As Long, ByVal B As Long, ByRef stepIndex As Long, ByVal totalSteps As Long) As Long
    On Error GoTo Except

    ' r, g, b = color to arrive at
    Dim calcR As Long, calcG As Long, calcB As Long
    Const ProcName As String = "InterpolateColor"
    
    ' interpolate from white (255,255,255) to target color
    calcR = 255 - ((255 - R) * stepIndex / totalSteps)
    calcG = 255 - ((255 - G) * stepIndex / totalSteps)
    calcB = 255 - ((255 - B) * stepIndex / totalSteps)
    
    stepIndex = stepIndex + 1
    
    InterpolateColor = RGB(calcR, calcG, calcB)

Finally:
    Exit Function

Except:
    ReportExcept Err.Number, Err.Description, Erl, ProcName, ModName, Src
    Resume Finally

End Function

' form module / splash screen for calling InterpolateColor
Option Compare Database
Option Explicit
Dim stepIndex As Long
Private Sub Form_Load()
On Error GoTo Except

    Const ProcName As String = "Form_Load"
    ' set timer interval
    ' initialize stepIndex
    Me.TimerInterval = 100
    stepIndex = 0
    
Finally:
    Exit Sub

Except:
    ReportExcept Err.Number, Err.Description, Erl, ProcName, Me.Name, Src
    Resume Finally
    
End Sub
Private Sub Form_Timer()
    On Error GoTo Except
    
    Dim totalSteps As Long: totalSteps = 20
    Const ProcName As String = "Form_Timer"
    
    ' set color of label_Heading
    ' continue to interpolate color until arrive at 0, 114, 188
    Me.label_Heading.ForeColor = InterpolateColor(0, 114, 188, stepIndex, totalSteps)
    
    If stepIndex > totalSteps Then
        Me.TimerInterval = 0
        DoCmd.Close acForm, Me.Name
        DoCmd.OpenForm "Main_"
    End If
    
Finally:
    Exit Sub

Except:
    ReportExcept Err.Number, Err.Description, Erl, ProcName, Me.Name, Src
    Resume Finally
    
End Sub

