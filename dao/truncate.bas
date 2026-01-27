' similar to t-sql truncate table:
' batched transaction to perform both:
' delete all table records and reset counter (identity)
' gracefully performs rollback and exit if 3211 or other err
' best for staging tables / not known to work on tables with established relationships

Public Function DAOTruncate(ByVal tableName As String, ByVal columnName As String, Optional ByVal notifyUser As Boolean = False) As Boolean
    On Error GoTo Except

    ' delete all records in a table and reset autonumber column
    ' works for tables without relationships (eg. staging tables)
    Const ProcName As String = "DAOTruncate"
    Dim db As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field, inTrans As Boolean
    
    DAOTruncate = False
    
    Set db = CurrentDb
    
    ' exit if named table does not exist
    If Not TableExists(tableName) Then
        MsgBox "Table " & tableName & " does not exist.", vbInformation, ProcName
        Exit Function
    End If
    
    'determine if table is related to another table
    Set tdf = db.TableDefs(tableName)
    If TableIsRelated(db, tdf) Then
        MsgBox tableName & " has established relationships and cannot be truncated.", vbInformation, ProcName
        Exit Function
    End If

    ' suppress exception handling
    On Error Resume Next
    
    ' attempt to set named column for table    
    Set fld = tdf.Fields(columnName)
    
    ' resume exception handling
    On Error GoTo Except
    
    ' if not set column, column does not exist
    If fld Is Nothing Then
        MsgBox "Column " & columnName & " does not exist in " & tableName & ".", vbInformation, ProcName
        Exit Function
    End If
    
    ' make sure column is autonumber
    If Not IsAutoNumber(fld) Then
        MsgBox "Column " & columnName & " is not an auto-number.", vbInformation, ProcName
        Exit Function
    End If
    
    ' confirm proceed
    If MsgBox("Truncate table " & tableName & "?", vbYesNo + vbQuestion, ProcName) = vbNo Then
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
    ' ddl statement
    db.Execute "ALTER TABLE " & tableName & " ALTER COLUMN " & columnName & " COUNTER(1,1);", 128
    
    ' commit the transactions
    DBEngine.CommitTrans
    inTrans = False
    
    ' success
    DAOTruncate = True
    
    ' show message if notifyUser
    If notifyUser Then MsgBox tableName & " successfully truncated.", vbInformation, ProcName

ExitAttempt:
Finally:
    ' if transaction is open when reaching finally, rollback
    If inTrans Then DBEngine.Rollback    
    Exit Function

Except:
    Select Case Err.Number
        Case 3211
            MsgBox "Truncation cannot occur.  Table is in use.", vbInformation, ProcName
        Case Else
        ReportExcept Err.Number, Err.Description, Erl, ProcName, ModName, src
    End Select
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
