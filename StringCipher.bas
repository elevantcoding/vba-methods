Attribute VB_Name = "StringCipher"
Option Compare Database
Option Explicit
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Procedures in this module:
'---------------------------------------------------------------------------------------------------------------------------------------------------
' CipherString, DecipherString, GenerateCiph, GetAlterVals, GetRandChar, GetRandVal, NumCipher, ReplaceCharAtIndex, ValidateXorRange, ValidHexValues
'---------------------------------------------------------------------------------------------------------------------------------------------------
' CipherString performs a custom obfuscation using randoms, Xor and an on-the-fly, random numeric cipher.
' this obfuscation results in a 256-length string of hex values (128 chars).
' DecipherString reverses the obfuscation to the original string value.
'---------------------------------------------------------------------------------------------------------------------------------------------------
Const ModName As String = "StringCipher"
Function ValidHexValues(ByVal myhexstring As String) As Boolean
    
    'return true if every 2 chars can be evaluated as hex value
    Dim char As String, i As Long, isvalidhex As Boolean
    
    ValidHexValues = False
            
    If Len(myhexstring) = 0 Then Exit Function
    If Len(myhexstring) Mod 2 <> 0 Then Exit Function
    
    char = ""
    For i = 1 To Len(myhexstring) Step 2
        On Error Resume Next
        char = Chr(CLng("&H" & Mid(myhexstring, i, 2)))
        isvalidhex = (Err.Number = 0)
        If Not isvalidhex Then
            Err.Clear
            Exit Function
        End If
    Next
    ValidHexValues = True
    
End Function
Function GetRandChar(Optional ByVal l As Long, Optional ByVal u As Long) As String
    
    ' used for ascii values between 32 and 255, no negative values
    ' if passing optionals, optionals will only be used if within range
    Dim val As Long, upper As Long, lower As Long
    
    lower = 32: upper = 255
    l = Abs(l): u = Abs(u)
    
    If l > u Then
        val = l
        l = u
        u = val
    End If
    
    If l > 0 And l > lower Then lower = l
    If u > 0 And u < upper Then upper = u
    
    Do
        val = Int((upper - lower + 1) * Rnd + lower)
    Loop Until val <> 127
    
    GetRandChar = Chr(val)

End Function
Function GetRandVal(ByVal lower As Long, ByVal upper As Long) As Long
    
    ' return a random integer between lower and upper (inclusive), no negative values
    ' if lower > upper, the values are swapped.
    ' use Randomize in the calling procedure to ensure non-repeating sequences.
    Dim val As Long
    lower = Abs(lower): upper = Abs(upper)
    
    If lower > upper Then
        val = lower
        lower = upper
        upper = val
    End If
    
    GetRandVal = Int((upper - lower + 1) * Rnd + lower)
    
End Function
Function ReplaceCharAtIndex(ByVal origString As String, ByVal idx As Long, ByVal newChar As String) As String

    ' replaces the character at the specified 1-based index within a string.
    ' if the index is out of range or newChar is not exactly one character,
    ' the original string is returned unchanged.
    ReplaceCharAtIndex = origString
    If idx < 1 Or idx > Len(origString) Then Exit Function
    If Len(newChar) <> 1 Then Exit Function
    ReplaceCharAtIndex = Left(origString, idx - 1) & newChar & Mid(origString, idx + 1)

End Function
Function GetAlterVals(ByVal getvals As String, ByVal cipher As Boolean) As String
    
    ' applies a reversible alternating +/- 1 transformation to a numeric string.
    
    ' getvals:
    '   a numeric string (digits 0–9 only).
    ' cipher:
    '   True  ? apply +1 to odd positions, -1 to even positions
    '   False ? apply -1 to odd positions, +1 to even positions
    ' this creates a reversible numeric mutation used to rotate the XOR driver
    ' between cipher and decipher passes.
    ' example:
    '   getvals = "123", cipher = True  ? "214"
    '   getvals = "214", cipher = False ? "123"
    ' digits wrap safely using modulo 10 arithmetic.
    Dim v As Long, num As Long, stepVal As Long, returnvals As String
    
    GetAlterVals = getvals
    If Len(getvals) = 0 Or getvals Like "*[!0-9]*" Then Exit Function
    
    returnvals = ""
    For v = 1 To Len(getvals)
        num = CLng(Mid(getvals, v, 1))
        
        If (v Mod 2 <> 0) = cipher Then
            stepVal = 1
        Else
            stepVal = -1
        End If
        
        num = (num + stepVal + 10) Mod 10
        returnvals = returnvals & CStr(num)
    Next
    
    GetAlterVals = returnvals

End Function
Function GenerateCiph() As String

    ' generates a 10-character substitution cipher consisting of:
    '   no numeric characters
    '   unique characters only (no duplicates)
    '   subset ascii 58 - 126 only
    '   uses GetRandChar
    
    ' this cipher is used by NumCipher to reversibly map digits 0–9
    ' into symbolic characters for prefix obfuscation.
    ' use Randomize in the calling procedure to ensure non-repeating sequences.
    Dim ciph As String, i As Long, char As String
    
    ciph = ""
    For i = 1 To 10
        Do
            char = GetRandChar(58, 126)
        Loop Until InStr(1, ciph, char, vbBinaryCompare) = 0
        ciph = ciph & char
    Next
    
    GenerateCiph = ciph

End Function
Function NumCipher(ByVal chars As String, ByVal cipher As Boolean, ByVal ciph As String) As String
    
    ' uses a 10-character substitution cipher (from GenerateCiph) to transform numeric characters.
    ' encode (cipher = true):
    '   pass numeric chars with cipher = True and the 10-char cipher string.
    '   returns a non-numeric symbolic representation of the digits.
    ' decode (cipher = false):
    '   pass the ciphered symbolic chars with cipher = False and the original 10-char cipher.
    '   returns the original numeric digits.
    '
    ' example:
    '   chars = "101", cipher = True,  ciph = "ZNdfLMsrb~"  ? result = "NZN"
    '   chars = "NZN", cipher = False, ciph = "ZNdfLMsrb~"  ? result = "101"'
    ' assumptions:
    '   encoding only accepts digits 0–9.
    '   cipher string must be exactly 10 characters.
    Dim i As Long, k As Long, C As Long, char As String, Result As String
    
    NumCipher = chars
    If Len(chars) = 0 Then Exit Function
    If Len(ciph) <> 10 Then Exit Function
    If cipher = True And chars Like "*[!0-9]*" Then Exit Function
    
    Const key As String = "0123456789"
    
    Result = ""
    For i = 1 To Len(chars)
        char = Mid(chars, i, 1)
        If cipher Then
            C = InStr(1, key, char, vbBinaryCompare)
            Result = Result & Mid(ciph, C, 1)
        Else
            k = InStr(1, ciph, char, vbBinaryCompare)
            Result = Result & Mid(key, k, 1)
        End If
    Next
    
    NumCipher = Result

End Function
Function CipherString(ByVal myTextstring As String) As String
    
    Dim prefix As String, stringtoCipher As String, ciph As String, ciphprefix As String
    Dim vals As String, altervals As String, char As String, addchar As String, paddedString As String, hexvalue As String, hexstring As String
    
    Dim strLen As Long, getasc As Long, getval As Long, addasc As Long, loops As Long, randval As Long
    Dim loopcount As Long, i As Long, v As Long, p As Long, s As Long, H As Long
    Dim prefixlen As Long, availablelen As Long, spacing As Long
    
    ' max length of string for hex chars: 256 / 2 = 128 regular chars max string length
    Const maxstrlen As Long = 128
    
    ' default value
    CipherString = ""

    ' nothing to cipher
    If Len(myTextstring) = 0 Then Exit Function
    
    Randomize
    ' get a 6-digit random value, 0 to 999,999
    vals = Format(GetRandVal(0, 999999), "000000")
    
    ' get a 1-digit random value 2 to 6
    randval = GetRandVal(2, 6)
    
    strLen = Len(stringtoCipher)
    i = strLen + 1
    loops = strLen * randval
    v = 0
    altervals = vals
    ' build string:
    ' traverse backward on stringtoCipher, forward on altervals
    ' for each char in string, replace char as per ascii value Xor altervals value
    For loopcount = 1 To loops
        i = i - 1
        v = v + 1
    
        If i = 0 Then i = strLen
        
        If v > Len(altervals) Then
            v = 1
            altervals = GetAlterVals(altervals, True)
        End If

        char = Mid(stringtoCipher, i, 1)
        getasc = Asc(char)
        getval = CLng(Mid(altervals, v, 1))
        addasc = getasc Xor getval
        addchar = Chr(addasc)
        stringtoCipher = ReplaceCharAtIndex(stringtoCipher, i, addchar)
    Next
    
    'create prefix: altervals, randval, v, strlen(000)
    prefix = altervals & randval & CStr(v) & Format(strLen, "000")
    
    ' get random, numeric cipher
    ciph = GenerateCiph
    
    ' cipher all prefix chars
    ciphprefix = NumCipher(prefix, True, ciph)
    
    ' add the ciph to the prefix
    prefix = ciphprefix & ciph
    
    ' do not trim in case asc 32
    prefixlen = Len(prefix)
    
    ' exit function if it exceeds maximum defined length
    If strLen > (maxstrlen - prefixlen) Then
        MsgBox "Maximum string length of " & (maxstrlen - prefixlen) & " exceeded.", vbInformation, "CipherString"
        Exit Function
    End If
    
    ' determine spacing for padded string
    availablelen = maxstrlen - prefixlen
    spacing = 0
    If strLen < availablelen Then spacing = Int(availablelen / strLen)
    
    ' if no spacing, paddedstring is stringtoCipher
    If spacing = 0 Then
        paddedString = stringtoCipher
    Else
        ' if spacing has been calculated:
        ' in stringtoCipher only: follow each cipher char with randchars as per spacing
        paddedString = "": s = spacing: i = 1
        For p = 1 To availablelen
        
            If s = spacing Then
                s = 1
            Else
                s = s + 1
            End If
        
            If s = 1 And i <= strLen Then
                paddedString = paddedString & Mid(stringtoCipher, i, 1)
                i = i + 1
            Else
                paddedString = paddedString & GetRandChar
            End If
        
        Next
    End If
    
    ' add prefix to padded string
    paddedString = prefix & paddedString
    
    ' create hex value / count must be maxstrlen to ensure expected string length
    ' single hex digits with ascii values 0 to 31 have 0 chance of appearing
    ' but pad single hex value with 0 just in case
    hexstring = ""
    For H = 1 To maxstrlen
        hexvalue = Hex(Asc(Mid(paddedString, H, 1)))
        If Len(hexvalue) = 1 Then hexvalue = "0" & hexvalue
        hexstring = hexstring & hexvalue
    Next
    
    CipherString = hexstring
    
End Function
Function DecipherString(ByVal myCipherstring As String) As String

    Dim paddedString As String, stringtoDecipher As String, altervals As String, altervalsOrig As String
    Dim char As String, chars As String, addchar As String, prefix As String, numciph As String
    
    Dim i As Long, loops As Long, loopcount As Long, v As Long, randval As Long, strLen As Long, getasc As Long, getval As Long
    Dim addasc As Long, p As Long, s As Long, spacing As Long, availablelen As Long
    
    ' prefix of CipherString consists of information to be parsed as follows
    ' (vertical bars used for illustration only and are not in the prefix)
    ' -----------------------------------------------------------
    ' 000000|0|0|000|aaaaaaaaaa
    ' -----------------------------------------------------------
    ' altervals: 6-char numeric string (most recent iteration of original random vals used via altervals)
    ' randval: 1-char numeric string (randval)
    ' v: 1-char numeric string (last used index of altervals)
    ' original string length: 3-char numeric string, padded with zero if < 100
    ' numciph: 10 chars for the numeric cipher (GenerateCiph used in NumCipher)
    ' prefix length =  21 characters (no vertical bars)
    
    ' static values
    Const prefixlen As Long = 21
    Const maxstrlen As Long = 128

    ' default value
    DecipherString = ""
    
    ' trim
    myCipherstring = Trim(myCipherstring)
    
    ' nothing to decipher
    If Len(myCipherstring) = 0 Then Exit Function
    
    ' must be made of hex values
    If Not ValidHexValues(myCipherstring) Then
        MsgBox myCipherstring & " does not contain valid or all valid hex values to decipher.", vbInformation, "DecipherString"
        Exit Function
    End If
    
    ' get chars from hex
    stringtoDecipher = ""
    For i = 1 To Len(myCipherstring) Step 2
        stringtoDecipher = stringtoDecipher & Chr(CLng("&H" & Mid(myCipherstring, i, 2)))
    Next
    
    ' get prefix
    prefix = Left(stringtoDecipher, prefixlen)

' get the numeric cipher from the prefix and trim from prefix
    numciph = Right(prefix, 10)
    prefix = Left(prefix, Len(prefix) - 10)
    
    ' decipher using numcipher
    prefix = NumCipher(prefix, False, numciph)
    
    ' work from the right:
    ' get 3-digits for string length
    chars = Right(prefix, 3)
    If Not chars Like "###" Then
        MsgBox "Value of string length not found in prefix.", vbInformation, "DecipherString"
        Exit Function
    ElseIf CLng(chars) = 0 Then
        MsgBox "Value of string length is zero.", vbInformation, "DecipherString"
        Exit Function
    Else
        strLen = CLng(chars)
    End If
    
    ' trim string length from prefix
    prefix = Left(prefix, Len(prefix) - 3)
    
    ' get v
    char = Right(prefix, 1)
    If Not (char Like "#") Then
        MsgBox "Value of v not found in prefix.", vbInformation, "DecipherString"
        Exit Function
    Else
        v = CLng(char)
    End If
    
    ' trim v from prefix
    prefix = Left(prefix, Len(prefix) - 1)
        
    ' get randval from prefix
    char = Right(prefix, 1)
    If Not (char Like "#") Then
        MsgBox "Random value not found in prefix.", vbInformation, "DecipherString"
        Exit Function
    Else
        randval = CLng(char)
    End If
    
    ' trim randval from prefix
    prefix = Left(prefix, Len(prefix) - 1)
    
    ' get altervals from prefix, decipher using NumCipher, assign to altervals
    altervals = prefix
    If Len(altervals) <> 6 Then
        MsgBox "Random values for string alteration not found in prefix.", vbInformation, "DecipherString"
        Exit Function
    End If

    ' remove prefix from string from stringtoDecipher, leaving actual ciphered string chars
    paddedString = Right(stringtoDecipher, Len(stringtoDecipher) - prefixlen)
    
    ' get available length and spacing using strLen
    availablelen = maxstrlen - prefixlen
    spacing = Int((availablelen) / strLen)
    If spacing < 1 Then spacing = 1
    
    ' remove padding between cipher chars
    stringtoDecipher = "": i = 0: s = 1
    For p = 1 To availablelen
        If s = 1 Then
            stringtoDecipher = stringtoDecipher & Mid(paddedString, p, 1)
            i = i + 1
            If i = strLen Then Exit For
        End If
        
        If s = spacing Then
            s = 1
        Else
            s = s + 1
        End If
    Next
    
    i = 1
    loops = strLen * randval
    altervalsOrig = altervals
    altervals = Left(altervals, v)
    
    ' re-build string:
    ' traverse forward on stringtoDecipher, backward on altervals
    ' for each char in string, replace char as per ascii value Xor altervals value
    For loopcount = 1 To loops
        
        If i > strLen Then i = 1
        
        If v = 0 Then
            If Len(altervals) < 6 Then altervals = altervalsOrig
            v = Len(altervals)
            altervals = GetAlterVals(altervals, False)
        End If
        
        char = Mid(stringtoDecipher, i, 1)
        getasc = Asc(char)
        getval = CLng(Mid(altervals, v, 1))
        addasc = getasc Xor getval
        addchar = Chr(addasc)
        stringtoDecipher = ReplaceCharAtIndex(stringtoDecipher, i, addchar)
               
        i = i + 1
        v = v - 1

    Next
    
    DecipherString = stringtoDecipher

End Function
Sub ValidateXorRange()

    ' validation:
    ' confirms that Xor applied to any printable ascii character 32–255
    ' using numeric keys 0–9 never produces a value below 32 or above 255
    ' this guarantees XOR-based cipher output always remains printable
    Dim i As Long, k As Long, x As Long, C As Long
    
    C = 0
    For i = 32 To 255
        For k = 0 To 9
            x = i Xor k
            If x < 32 Or x > 255 Then
                C = C + 1
                Debug.Print i & " Xor " & k & " is " & x
            End If
        Next
    Next
    
    Debug.Print "finished: " & C & " results"
End Sub








