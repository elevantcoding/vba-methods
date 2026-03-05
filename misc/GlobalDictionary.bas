' usage:
' UseDictionary(daNew)
' do stuff
' If UseDictionary(adExists, [KeyID]) Then
' do stuff
' UseDictionary(adAddEntry, [KeyID], [KeyValue])
' do stuff
' UseDictionary(daClose)

Public dict As Object

' options for UseDictionary
Public Enum DictActions
    daAddEntry
    daRemoveEntry
    daExists
    daInUse
    daNew
    daClose
    daRead
End Enum

Public Function UseDictionary(ByVal Selected As DictActions, Optional ByVal KeyID As Variant = Null, Optional ByRef KeyValue As Variant = Null) As Boolean
    
    ' store values for reference / retrieval
    ' use option daClose when finished to clear
    
    Dim blnResult As Boolean: blnResult = False
    Dim K As Variant
    
    ' make sure KeyId passed for specified actions
    Select Case Selected
        Case daAddEntry, daRemoveEntry, daExists
            If IsNull(KeyID) Or Len(KeyID & "") = 0 Then Exit Function
    End Select
    
    ' for single, selected action
    Select Case Selected
        Case daAddEntry ' add an entry to dict
            If dict Is Nothing Then
                Set dict = CreateObject("Scripting.Dictionary")
            End If
            
            If Not IsNull(KeyID) Then
                If dict.Exists(KeyID) Then
                    dict(KeyID) = KeyValue
                Else
                    dict.Add KeyID, KeyValue
                End If
                blnResult = True
            End If
        Case daRemoveEntry ' remove entry from dict
            If Not dict Is Nothing Then
                If Not IsNull(KeyID) Then
                    If dict.Exists(KeyID) Then
                        dict.Remove (KeyID)
                        blnResult = True
                    End If
                End If
            End If
            
        Case daExists ' check existence of dict
            If Not dict Is Nothing Then
                If Not IsNull(KeyID) Then
                    blnResult = dict.Exists(KeyID)
                End If
            End If
            
        Case daInUse 'is dict instantiated?
            blnResult = (Not dict Is Nothing)
            
        Case daClose, daNew ' remove entries / create new dict
            If Not dict Is Nothing Then
                dict.RemoveAll
                Set dict = Nothing
                If Selected = daClose Then
                    blnResult = True
                    Exit Function
                End If
            End If
            blnResult = True
            
        Case daRead ' read entry, return via ByRef KeyValue
            If Not dict Is Nothing Then
                If dict.Exists(KeyID) Then
                    KeyValue = dict(KeyID)
                    blnResult = True
                End If
            End If
    End Select
    
    UseDictionary = blnResult
    
End Function
