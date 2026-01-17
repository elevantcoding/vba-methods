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

' detect string similarity
Function CompareStrings(ByVal strFirst As String, ByVal strSecond As String) As Variant
    On Error GoTo Except

    'the result of this function will indicate % difference of two strings
    Const ProcName As String = "CompareStrings"
    Dim lFirstLen As Long
    Dim lSecondLen As Long
    
    Dim lDividend As Long
    Dim lDivisor As Long
             
    Dim lCharacterCount As Long
    Dim lCharacters As Long
    Dim lBeginCount As Long
    Dim lFoundCount As Long
    Dim lCount As Long
    Dim lSlideLen As Long
    
    Dim strCharacter As String
    Dim strCharactersFound As String
    Dim strEvaluate As String
    Dim strLookIn As String
    
    ' use text compare only
    'if strings are identical, result = 0, no difference, exit
    If StrComp(strFirst, strSecond, vbTextCompare) = 0 Then
        CompareStrings = 0
        Exit Function
    End If
        
    ' normalize case
    strFirst = LCase(strFirst)
    strSecond = LCase(strSecond)
    
    'replace distinct numeric values / replace St with Street, etc.
    strFirst = NormalizeString(strFirst)
    strSecond = NormalizeString(strSecond)
    
    'remove spaces and punctuation
    strFirst = CleanString(strFirst)
    strSecond = CleanString(strSecond)
    
    'get number characters in original string and new string
    lFirstLen = Len(strFirst)
    lSecondLen = Len(strSecond)
    
    'lCharacters is whichever string is shortest, if equal use lFirstLen
    Select Case True
        Case lFirstLen = lSecondLen, lFirstLen < lSecondLen
            lCharacters = lFirstLen
            strLookIn = strSecond
            strEvaluate = strFirst
            lDivisor = lSecondLen
            
        Case lFirstLen > lSecondLen
            lCharacters = lSecondLen
            strLookIn = strFirst
            strEvaluate = strSecond
            lDivisor = lFirstLen
            
        Case Else
            MsgBox "Unknown", vbInformation, "Compare Strings"
            Exit Function
    End Select
    
    'sliding length: review every lSlideLen characters for match
    'if 2 or less characters, slide by 1
    'if 6 or less characters, slide by 2
    'else slide by 4
    If lCharacters <= 2 Then
        lSlideLen = 1
    ElseIf lCharacters <= 6 Then
        lSlideLen = 2
    Else
        lSlideLen = 4
    End If
    
    lCount = 0
    lBeginCount = 0
    strCharactersFound = ""
    For lCharacterCount = 1 To lCharacters        
        strCharacter = Mid(strEvaluate, lCharacterCount, lSlideLen) 
        If Len(strCharacter) = lSlideLen Then 
            lFoundCount = InStr(1, strLookIn, strCharacter, vbTextCompare)
            
            If lFoundCount > 0 Then         
                If lBeginCount = 0 Then 
                    lBeginCount = lFoundCount
                    lCount = lSlideLen
                Else
                    lCount = lCount + 1
                End If
                strCharactersFound = Mid(strLookIn, lBeginCount, lCount)
            End If
        End If        
    Next
    
    If strCharactersFound <> "" Then
        lDividend = Len(strCharactersFound)
        CompareStrings = 1 - (lDividend / lDivisor)
        Exit Function
    Else
        CompareStrings = 1
    End If
          
Finally:
    Exit Function

Except:
    ReportExcept Err.Number, Err.Description, Erl, ProcName, Me.Name, Src
    Resume Finally

End Function
Function NormalizeString(ByVal strInput As String) As String
    On Error GoTo Except

    ' Add padding to help detect word boundaries at edges
    strInput = " " & strInput & " "

    ' Expand street suffixes (case-insensitive if needed)
    strInput = Replace(strInput, " St ", " Street ")
    strInput = Replace(strInput, " St.", " Street")
    strInput = Replace(strInput, " Ave ", " Avenue ")
    strInput = Replace(strInput, " Rd ", " Road ")
    strInput = Replace(strInput, " Blvd ", " Boulevard ")
    strInput = Replace(strInput, " Ln ", " Lane ")
    strInput = Replace(strInput, " Dr ", " Drive ")

    ' Normalize number words
    strInput = Replace(strInput, " 1 ", " One ")
    strInput = Replace(strInput, " 2 ", " Two ")
    strInput = Replace(strInput, " 3 ", " Three ")
    strInput = Replace(strInput, " 5 ", " Five ")
    strInput = Replace(strInput, " 6 ", " Six ")
    strInput = Replace(strInput, " 7 ", " Seven ")
    strInput = Replace(strInput, " 8 ", " Eight ")
    strInput = Replace(strInput, " 9 ", " Nine ")
    
    ' other numeric equivalents
    strInput = Replace(strInput, " 1st ", " First ")
    strInput = Replace(strInput, " 2nd ", " Second ")
    strInput = Replace(strInput, " 3rd ", " Third ")
    strInput = Replace(strInput, " 4th ", " Fourth ")
    strInput = Replace(strInput, " 5th ", " Fifth ")
    strInput = Replace(strInput, " 6th ", " Sixth ")
    strInput = Replace(strInput, " 7th ", " Seventh ")
    strInput = Replace(strInput, " 8th ", " Eighth ")
    strInput = Replace(strInput, " 9th ", " Ninth ")
    strInput = Replace(strInput, " 10th ", " Tenth ")
    strInput = Replace(strInput, " 11th ", " Eleventh ")
    strInput = Replace(strInput, " 12th ", " Twelfth ")

    ' Collapse double spaces and trim
    Do While InStr(strInput, "  ") > 0
        strInput = Replace(strInput, "  ", " ")
    Loop

    NormalizeString = Trim(strInput)

Finally:
    Exit Function

Except:
    NormalizeString = strInput
    ReportExcept Err.Number, Err.Description, Erl, ProcName, Me.Name, Src
    Resume Finally
End Function

Function CleanString(ByVal strInput As String) As String
    On Error GoTo Except
    
    Dim i As Long
    Dim char As String
    Dim Result As String

    Result = ""

    For i = 1 To Len(strInput)
        char = Mid(strInput, i, 1)
        If char Like "[A-Za-z0-9]" Then
            Result = Result & char
        End If
    Next i

    CleanString = Result
    
Finally:
    Exit Function

Except:
    ReportExcept Err.Number, Err.Description, Erl, ProcName, Me.Name, Src
    Resume Finally

End Function

