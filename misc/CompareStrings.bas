Public Function StringCompare(ByVal FirstStr As String, ByVal SecondStr As String, Optional Comp As VbCompareMethod = vbTextCompare) As Variant
    On Error GoTo Except

    Const ProcName As String = "StringCompare"
    
    ' dynamic sliding-window evaluation for two strings
    ' produces a % difference score, 0 is identical strings, 1 is completely different strings
    
    Dim FirstLen As Long, SecondLen As Long, CharCount As Long, LookIn As String, Evaluate As String, Divisor As Long, Dividend As Long, WindowSize As Long, i As Long
    Dim CharLen As Long, Characters As String, CharsFound As Long, Found As Boolean
    
    StringCompare = 0
    
    ' first, remove punctuation
    ' then, remove specified terms: construction, co, etc.
    ' then, normalize: example 1st to First
    ' then, clean: remove spaces
    
    FirstStr = StringClean(StringNormalize(StringExclude(StringPunctuation(FirstStr))))
    SecondStr = StringClean(StringNormalize(StringExclude(StringPunctuation(SecondStr))))
    
    ' if the strings match, no difference: exit
    If StrComp(FirstStr, SecondStr, Comp) = 0 Then Exit Function
    
    ' get length of each string
    FirstLen = Len(FirstStr)
    SecondLen = Len(SecondStr)
    
    ' LookIn the longer string for the shorter string
    Select Case True
        Case FirstLen <= SecondLen
            CharCount = FirstLen
            LookIn = SecondStr
            Evaluate = FirstStr
            Divisor = SecondLen
        Case Else
            CharCount = SecondLen
            LookIn = FirstStr
            Evaluate = SecondStr
            Divisor = FirstLen
    End Select

    ' granularity
    Select Case CharCount
        Case Is <= 4
            WindowSize = 1
        Case Is <= 10
            WindowSize = 2
        Case Else
            WindowSize = 3
    End Select
    
    Found = False
    CharsFound = 0
    
    ' traverse shorter string while advancing search through the longer string:
    ' for every n chars, while the character length equals the window length,
    ' when first full match is found, found is true: get number of chars found on first match
    ' after found is true and since advancing by one, for each following consecutive match, add 1 to chars found count
    ' if no longer found match, switch found to false to discontinue counting block / consecutive char matches until
    ' the next match is found in the same manner
    For i = 1 To CharCount
        Characters = Mid(Evaluate, i, WindowSize)
        CharLen = Len(Characters)
        If CharLen = WindowSize Then
            If InStr(i, LookIn, Characters, Comp) > 0 Then
                If Not Found Then
                    CharsFound = CharsFound + Len(Characters)
                    Found = True
                Else
                    CharsFound = CharsFound + 1
                End If
            Else
                Found = False
            End If
        End If
    Next
    
    StringCompare = 1
    If CharsFound > 0 Then
        Dividend = CharsFound
        StringCompare = CDec(1 - (Dividend / Divisor))
    End If

Finally:
    Exit Function
    
Except:
    ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Function
Private Function StringNormalize(ByVal strInput As String) As String
    
    ' normalize a string, return original val on err
    
    Dim ReturnValue As String
    
    On Error Resume Next

    ' Add padding to help detect word boundaries at edges
    ReturnValue = " " & strInput & " "
    
    ' Expand street suffixes (case-insensitive if needed)
    ReturnValue = Replace(ReturnValue, " St ", " Street ")
    ReturnValue = Replace(ReturnValue, " St.", " Street")
    ReturnValue = Replace(ReturnValue, " Ave ", " Avenue ")
    ReturnValue = Replace(ReturnValue, " Rd ", " Road ")
    ReturnValue = Replace(ReturnValue, " Blvd ", " Boulevard ")
    ReturnValue = Replace(ReturnValue, " Ln ", " Lane ")
    ReturnValue = Replace(ReturnValue, " Dr ", " Drive ")

    ' Normalize number words
    ReturnValue = Replace(ReturnValue, " 1 ", " One ")
    ReturnValue = Replace(ReturnValue, " 2 ", " Two ")
    ReturnValue = Replace(ReturnValue, " 3 ", " Three ")
    ReturnValue = Replace(ReturnValue, " 5 ", " Five ")
    ReturnValue = Replace(ReturnValue, " 6 ", " Six ")
    ReturnValue = Replace(ReturnValue, " 7 ", " Seven ")
    ReturnValue = Replace(ReturnValue, " 8 ", " Eight ")
    ReturnValue = Replace(ReturnValue, " 9 ", " Nine ")
    
    ' other numeric equivalents
    ReturnValue = Replace(ReturnValue, " 1st ", " First ")
    ReturnValue = Replace(ReturnValue, " 2nd ", " Second ")
    ReturnValue = Replace(ReturnValue, " 3rd ", " Third ")
    ReturnValue = Replace(ReturnValue, " 4th ", " Fourth ")
    ReturnValue = Replace(ReturnValue, " 5th ", " Fifth ")
    ReturnValue = Replace(ReturnValue, " 6th ", " Sixth ")
    ReturnValue = Replace(ReturnValue, " 7th ", " Seventh ")
    ReturnValue = Replace(ReturnValue, " 8th ", " Eighth ")
    ReturnValue = Replace(ReturnValue, " 9th ", " Ninth ")
    ReturnValue = Replace(ReturnValue, " 10th ", " Tenth ")
    ReturnValue = Replace(ReturnValue, " 11th ", " Eleventh ")
    ReturnValue = Replace(ReturnValue, " 12th ", " Twelfth ")
    
    ' directions
    ReturnValue = Replace(ReturnValue, " N ", " North ")
    ReturnValue = Replace(ReturnValue, " S ", " South ")
    ReturnValue = Replace(ReturnValue, " E ", " East ")
    ReturnValue = Replace(ReturnValue, " W ", " West ")
    
    ' misc
    ReturnValue = Replace(ReturnValue, " the ", " ")
    ReturnValue = Replace(ReturnValue, " and ", " ")
    ReturnValue = Replace(ReturnValue, " of ", " ")
    
    ' Collapse double spaces and trim
    Do While InStr(ReturnValue, "  ") > 0
        ReturnValue = Replace(ReturnValue, "  ", " ")
    Loop
    
    If Err.Number <> 0 Then
        StringNormalize = strInput
    Else
        StringNormalize = Trim(ReturnValue)
    End If
    
    On Error GoTo 0

End Function
Private Function StringExclude(ByVal s As String) As String
    
    Dim re As Object, terms As Variant, t As Variant
    Set re = CreateObject("VBScript.RegExp")
    
    re.IgnoreCase = True
    re.Global = True ' Replace all occurrences
    
    terms = Array("development", "developer", "construction", "builders", "contracting", _
        "group", "company", "co", "llc", "inc", "corp", "services", "solutions")
    
    For Each t In terms
        ' \b is the Regex "Word Boundary" marker. It handles start/end/spaces automatically.
        re.pattern = "\b" & t & "s?\b" ' The s? also catches plurals like "builder" vs "builders"
        s = re.Replace(s, "")
    Next
    
    ' Clean up commas/periods and double spaces left behind
    s = Replace(s, ",", "")
    s = Replace(s, ".", "")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    
    StringExclude = Trim(s)
End Function
Public Function StringPunctuation(ByVal strInput As String) As String

    Dim i As Long, Char As String, Result As String
    Result = ""
    
    For i = 1 To Len(strInput)
        Char = Mid(strInput, i, 1)
        Select Case Asc(Char)
            Case 33 To 47, 58 To 64, 91 To 96, 123 To 127
                Result = Result & " "
            Case Else
                Result = Result & Char
        End Select
    Next
    
    StringPunctuation = Result

End Function
Public Function StringClean(ByVal strInput As String) As String
    
    Dim i As Long
    Dim Char As String
    Dim Result As String

    Result = ""

    For i = 1 To Len(strInput)
        Char = Mid(strInput, i, 1)
        If Char Like "[A-Za-z0-9]" Then
            Result = Result & Char
        End If
    Next i

    StringClean = Result
    
End Function
