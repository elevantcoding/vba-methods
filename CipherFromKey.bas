Attribute VB_Name = "GitHub"
Option Compare Database
Option Explicit
Const ModName As String = "GitHub"
Function GetRandomizedCipherFromKey() As String

    ' create a randomized cipher that is relative to the key: get ascii value of the character in the key, cipher character must be one of the characters
    ' in the key, that has not yet been used and match the evenness or oddness of the key character value
    ' i.e. if a = ascii value = 97 (which is an odd number), the cipher character must also have an ascii value that is odd
    
    ' A to Z: 65 - 90
    ' a to z: 97 - 122
    ' 0 to 9: 48 - 57
    ' symbols: 33, 35 to 38, 40 to 47, 58 to 64
    
    ' for every character in the key, get the character and the ascii value of the character
    ' if the ascii value is even, loop for a random value based on the upper and lower value
    
    ' the condition of the loop is as follows:
    ' the result is an even number
    ' cipherchar is not already in rndcipher
    ' and cipherchar does exist in key
    
    ' do the same when the ascii value of the keychar in key is an odd number
    
    Dim i As Long
    Dim getval As Long
    Dim randval As Long
    Dim upperval As Long
    Dim lowerval As Long
    
    Dim rndcipher As String
    Dim keychar As String
    Dim cipherchar As String
    
    Const key As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!#$%&()*+,-./:;<=>?@"
    
    lowerval = 33
    upperval = 122
    
    rndcipher = ""
    Randomize
    
    For i = 1 To Len(key)
        keychar = Mid(key, i, 1)
        getval = Asc(keychar)
        
        If getval Mod 2 = 0 Then
            Do
                randval = Int((upperval - lowerval + 1) * Rnd + lowerval)
                cipherchar = Chr(randval)
            Loop Until randval Mod 2 = 0 And InStr(1, rndcipher, cipherchar, vbBinaryCompare) = 0 And InStr(1, key, cipherchar, vbBinaryCompare) > 0
            rndcipher = rndcipher & cipherchar
        Else
            Do
                randval = Int((upperval - lowerval + 1) * Rnd + lowerval)
                cipherchar = Chr(randval)
            Loop Until randval Mod 2 <> 0 And InStr(1, rndcipher, cipherchar, vbBinaryCompare) = 0 And InStr(1, key, cipherchar, vbBinaryCompare) > 0
            
            rndcipher = rndcipher & cipherchar
        End If
        
    Next
    
    GetRandomizedCipherFromKey = rndcipher

End Function


