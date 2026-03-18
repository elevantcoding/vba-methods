' usage:
' UseDictionary(daNew) ' instantiate
' do stuff
' If UseDictionary(adExists, [KeyID]) Then 'eval if a key exists
' do stuff
' UseDictionary(adAddEntry, [KeyID], [KeyValue]) ' add an entry
' do stuff
' UseDictionary(daClose) ' clear and close dictionary

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

Public Function UseDictionary(ByRef dict As Object, ByVal SelectedAction As DictActions, Optional ByVal KeyID As Variant = Null, Optional ByRef KeyValue As Variant = Null) As Boolean
    On Error GoTo Except
    Const ProcName As String = "UseDictionary"
    
    ' store values for reference / retrieval
    ' use option daClose when finished to clear
    Dim blnResult As Boolean: blnResult = False
    
    ' make sure dict exists / KeyID passed for specified actions
    Select Case SelectedAction
        Case daAddEntry, daExists, daRemoveEntry, daRead, daCount
            If dict Is Nothing Then RaiseCustomMsg SysDictInit, ProcName
            If SelectedAction <> daCount Then
                If IsNull(KeyID) Or Len(KeyID & "") = 0 Then Exit Function
            End If
    End Select
    
    ' for single, selected action
    Select Case SelectedAction
        Case daNew ' create new dictionary
            If Not dict Is Nothing Then
                If dict.Count > 0 Then dict.RemoveAll
                Set dict = Nothing
            End If
            Set dict = CreateObject("Scripting.Dictionary")
            blnResult = (Not dict Is Nothing)
            
        Case daAddEntry ' add entry / update keyvalue in dict
            If Not IsNull(KeyID) Then
                If dict.Exists(KeyID) Then
                    dict(KeyID) = KeyValue
                Else
                    dict.Add KeyID, KeyValue
                End If
                blnResult = True
            End If
        
        Case daExists ' check existence of dict
            If Not IsNull(KeyID) Then
                blnResult = dict.Exists(KeyID)
            End If
            
        Case daRemoveEntry ' remove entry from dict
            If Not IsNull(KeyID) Then
                If dict.Exists(KeyID) Then
                    dict.Remove KeyID
                    blnResult = True
                End If
            End If
            
        Case daRead ' read entry, return via ByRef KeyValue
            If dict.Exists(KeyID) Then
                KeyValue = dict(KeyID)
                blnResult = True
            End If
            
        Case daCount ' has stored vals ?
            If dict.Count > 0 Then blnResult = True
            
        Case daInit ' is initialized ?
            blnResult = (Not dict Is Nothing)
            
        Case daClose ' remove entries
            If Not dict Is Nothing Then
                dict.RemoveAll
                Set dict = Nothing
            End If
            blnResult = True
    End Select
    
    UseDictionary = blnResult

Finally:
    Exit Function
Except:
    ReportExcept Erl, ProcName, ModName
    Resume Finally
End Function
