' performs a truncate operation in VBA
' batched transaction to perform both:
' delete all table records and reset counter (identity)
' standard error handling does not trap 3211 (when DDL action on table in use)
' so this function handles the error within the code module and
' gracefully performs rollback and exit
' best for staging tables / not known to work on tables with established relationships

Public Function DAOTruncate(ByVal tableName As String, ByVal columnName As String, Optional ByVal notifyUser As Boolean = False) As Boolean
    On Error GoTo Except

    ' similar to t-sql truncate:
    ' delete all records in a table and reset autonumber column
    ' works for tables without relationships (eg. staging tables)
    Const ProcName As String = "DAOTruncate"
    Dim db As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field, inTrans As Boolean
    Dim errNum As Long, errDesc As String, isDDLErr As Boolean
    
    DAOTruncate = False
    
    Set db = CurrentDb
    
    ' exit if named table does not exist
    If Not TableExists(tableName) Then
        MsgBox "Table " & tableName & " does not exist.", vbInformation, "DAOTruncate"
        Exit Function
    End If
    
    'determine if table is related to another table
    Set tdf = db.TableDefs(tableName)
    If TableIsRelated(db, tdf) Then
        MsgBox tableName & " has established relationships and cannot be truncated.", vbInformation, "DAOTruncate"
        Exit Function
    End If
    
    ' attempt to set named column for table
    On Error Resume Next
    Set fld = tdf.Fields(columnName)
    On Error GoTo 0
    
    ' if not set column, column does not exist
    If fld Is Nothing Then
        MsgBox "Column " & columnName & " does not exist in " & tableName & ".", vbInformation, "DAOTruncate"
        Exit Function
    End If
    
    ' make sure column is autonumber
    If Not IsAutoNumber(fld) Then
        MsgBox "Column " & columnName & " is not an auto-number.", vbInformation, "DAOTruncate"
        Exit Function
    End If
    
    ' confirm proceed
    If MsgBox("Truncate table " & tableName & "?", vbYesNo + vbQuestion, "DAOTruncate") = vbNo Then
        Exit Function
    End If
    
    ' add brackets if not exist
    tableName = AddBrackets(tableName)
    columnName = AddBrackets(columnName)
    
    ' perform ops as batch transaction
    DBEngine.BeginTrans
    inTrans = True
    
    ' delete records
    db.Execute "DELETE FROM " & tableName & ";", 128

    ' encapsulate ddl action in error handling
    On Error Resume Next
    db.Execute "ALTER TABLE " & tableName & " ALTER COLUMN " & columnName & " COUNTER(1,1);", 128
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    
    ' if err is found, is ddl error, ExitAttempt to rollback and notify
    If errNum <> 0 Then
        isDDLErr = True
        GoTo ExitAttempt
    End If
    
    ' commit the transactions
    DBEngine.CommitTrans
    inTrans = False
    
    ' success
    DAOTruncate = True
    
    ' show message if notifyUser
    If notifyUser Then MsgBox tableName & " successfully truncated.", vbInformation, "DAOTruncate"

ExitAttempt:
Finally:
    ' if transaction is open when reaching finally, rollback
    If inTrans Then DBEngine.Rollback
    
    ' if is ddl error, indicate
    If isDDLErr And notifyUser Then
        MsgBox "Error Number " & errNum & " occurred when attempting the DDL operation." & vbCrLf & vbCrLf & "Please " & _
            "make sure the table is not in use when truncating." & vbCrLf & vbCrLf & "Details: " & errDesc, vbInformation, "DAOTruncate"
    End If
    Exit Function

Except:
    ReportExcept Err.Number, Err.Description, Erl, ProcName, ModName, src
    Resume Finally
End Function

Public Function TableExists(ByVal tableName As String) As Boolean
    On Error Resume Next
    TableExists = Not Application.CurrentData.AllTables(tableName) Is Nothing
End Function

Private Function TableIsRelated(ByVal db As DAO.Database, ByVal tdf As DAO.TableDef) As Boolean
    On Error GoTo Except

    Const ProcName As String = "TableIsRelated"
    Dim rel As DAO.Relation
    For Each rel In db.Relations
        If rel.Table = tdf.Name Or rel.ForeignTable = tdf.Name Then
            TableIsRelated = True
            Exit Function
        End If
    Next

Finally:
    Exit Function
Except:
    ReportExcept Err.Number, Err.Description, Erl, ProcName, ModName, src
    Resume Finally
End Function

Private Function IsAutoNumber(ByVal fld As DAO.Field) As Boolean
    IsAutoNumber = ((fld.Attributes And dbAutoIncrField) <> 0)
End Function

Public Function AddBrackets(ByVal strInput As String) As String
    On Error GoTo Except
    
    Const ProcName As String = "AddBrackets"
    
    If Len(strInput) = 0 Then Exit Function

    strInput = Trim(strInput)
    
    If Right(strInput, 1) <> "]" Then strInput = strInput & "]"
    If Left(strInput, 1) <> "[" Then strInput = "[" & strInput    
    AddBrackets = strInput
    
Finally:
    Exit Function    
Except:
    ReportExcept Err.Number, Err.Description, Erl, ProcName, ModName, src
    Resume Finally    
End Function
