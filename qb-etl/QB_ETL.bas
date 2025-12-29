' =========================================================
' QuickBooks → SQL Server ETL Pipeline
'
' Requires:
'   - Reference: Microsoft ActiveX Data Objects x.x Library
'   - Reference: Microsoft DAO x.x Object Library
'
' External Dependencies:
'   - ADOConnect  As String  ' SQL Server connection string
'   - DisplayMsg             ' UI messaging helper (vba-sql-methods/Interaction.bas)
'   - MsgFrm                 ' Message form instance (vba-sql-methods/Interaction.bas)
'   - CloseMsgFrm            ' Closes message form (vba-sql-methods/Interaction.bas)
'
' Purpose:
'   High-performance ETL pipeline that transfers data from
'   ODBC-linked QuickBooks queries into SQL Server tables
'   using schema-driven SQL and parameterized ADO commands.
'   This ETL pipeline does not rely on saved Access queries.

' All source datasets are implemented as VBA functions that return SQL statements.  
' For date-based SQL, at runtime, the ETL engine resolves and executes the appropriate dataset function using `Eval()`.
' This removes the need for [Forms]![...] query dependencies.
' =========================================================

Public Sub UpdateQB(ByVal dFrom As Date)
    On Error GoTo Except
    
    Dim lCountConfirm As Long
    Dim inTransaction As Boolean
    Dim cSQL As ADODB.Connection
    
    'make sure connectivity exists
    If Not QBPreflight("UpdateQB") Then
        MsgBox "Not connected to QuickBooks. Please login to QuickBooks for the system to run the update.", vbInformation, "QuickBooks"
        Exit Sub
    End If
    
    Call DisplayMsg(, "Reviewing", "Just a Moment")
    DoCmd.RepaintObject acForm, MsgFrm
    DoEvents
    
    'make sure contract numbers in quickbooks inv and cm are in tblContracts in GBLDB
    If Not ConfirmMatchingContracts(dFrom) Then GoTo ExitProcessing
            
    If Not ReviewCC(dFrom) Then GoTo ExitProcessing
    
    Call DisplayMsg(, "Updating", "Just a Moment")
    DoCmd.RepaintObject acForm, MsgFrm
    DoEvents
    
    Set cSQL = New ADODB.Connection
    cSQL.open ADOConnect
    cSQL.BeginTrans 'batch operations within same transaction in case of error to rollback entire procedure
    inTransaction = True
    
    'if count of records in QBAccts, then update with QBAccts
    If GetQBRecordCount(QBAccts) > 0 Then
        Call DisplayMsg("UPDATING", "Accounts", "Just a Moment")
        DoCmd.RepaintObject acForm, MsgFrm
        DoEvents
        If Not WriteToSQLTables(cSQL, "tblQuickBooksAccounts", QBAccts) Then
            MsgBox "QuickBooks Accounts table was called to update but did not succeed.", vbInformation, "UpdateQB"
            GoTo ExitProcessing
        End If
    End If
    
    'if count of records in QBItem, then update with QBItem
    If GetQBRecordCount(QBItem) > 0 Then
        Call DisplayMsg("UPDATING", "Items", "Just a Moment")
        DoCmd.RepaintObject acForm, MsgFrm
        DoEvents
        If Not WriteToSQLTables(cSQL, "tblQuickBooksItem", QBItem) Then
            MsgBox "QuickBooks Item table was called to update but did not succeed.", vbInformation, "UpdateQB"
            GoTo ExitProcessing
        End If
    End If
    
    Call DisplayMsg("UPDATING", , "Just a Moment")
    DoCmd.RepaintObject acForm, MsgFrm
    DoEvents
    
    expectedCount = 0
    
    If Not UpdateQBTable(cSQL, "QBCC", "tblQuickBooksCreditCards", dFrom, "Credit Card Transactions") Then GoTo ExitProcessing
    
    If Not UpdateQBTable(cSQL, "QBInvoice", "tblQuickBooksRevenue", dFrom, "Revenue Transactions - Invoices") Then GoTo ExitProcessing
    
    If Not UpdateQBTable(cSQL, "QBCM", "tblQuickBooksRevenue", dFrom, "Revenue Transactions - Credits") Then GoTo ExitProcessing
    
    If Not UpdateQBTable(cSQL, "QBRecp", "tblQuickBooksReceipts", dFrom, "Revenue Transactions - Receipts") Then GoTo ExitProcessing
    
    If Not UpdateQBTable(cSQL, "QBPer", "tblQuickBooksPermits", dFrom, "Permits") Then GoTo ExitProcessing
    
    cSQL.CommitTrans
    inTransaction = False
    
    ' after commit, count transactions added to tables and compare with cumulative obtained in expectedCount
    lCountConfirm = CLng(Nz(DCount("*", "tblQuickBooksCreditCards", "ModDate>=#" & dFrom & "#"), 0))
    lCountConfirm = lCountConfirm + CLng(Nz(DCount("*", "tblQuickBooksRevenue", "ModDate>=#" & dFrom & "#"), 0))
    lCountConfirm = lCountConfirm + CLng(Nz(DCount("*", "tblQuickBooksReceipts", "ModDate>=#" & dFrom & "#"), 0))
    lCountConfirm = lCountConfirm + CLng(Nz(DCount("*", "tblQuickBooksPermits", "ModDate>=#" & dFrom & "#"), 0))
    
    If lCountConfirm <> expectedCount Then
        MsgBox expectedCount & " total records to be added does not match " & lCountConfirm & " records added.  Transaction has been committed.", vbOKOnly + vbInformation, "UpdateQB"
    Else
        MsgBox "QuickBooks Information has been successfully written to SQL Server.", vbOKOnly + vbInformation, "UpdateQB"
    End If
        
ExitProcessing:
Finally:
    If inTransaction Then cSQL.RollbackTrans
    Call CloseMsgFrm
    If Not cSQL Is Nothing Then
        If cSQL.State = 1 Then
            cSQL.Close
        End If
        Set cSQL = Nothing
    End If
    Exit Sub

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "UpdateQB", , ModName)
    Resume Finally
End Sub

Private Function UpdateQBTable(cSQL As ADODB.Connection, ByVal strSourceQuery As String, ByVal strDestinationTable As String, ByVal dFrom As Date, ByVal strType As String) As Boolean
    On Error GoTo Except

    Dim strFunction As String
    Dim strSQL As String
    
    Dim lCount As Long
    
    UpdateQBTable = False
    
    'join named query to from date to form function structure
    strFunction = strSourceQuery & "(#" & dFrom & "#)"
    
    'use Eval to parse as a function
    lCount = GetQBRecordCount(Eval(strFunction))
    
    If lCount > 0 Then
        expectedCount = expectedCount + lCount
    
        'display progress
        Call DisplayMsg("UPDATING", "From QB to SQL Server", strType)
        DoCmd.RepaintObject acForm, MsgFrm
        DoEvents
  
        'conditional sql
        Select Case strSourceQuery
            Case "QBCM"
                strSQL = "DELETE FROM " & strDestinationTable & " WHERE ModDate>='" & SQLDate(dFrom) & "' AND InvoiceType = 'Credit Memo'"
            Case "QBInvoice"
                strSQL = "DELETE FROM " & strDestinationTable & " WHERE ModDate>='" & SQLDate(dFrom) & "' AND InvoiceType = 'Invoice'"
            Case Else
                strSQL = "DELETE FROM " & strDestinationTable & " WHERE ModDate>='" & SQLDate(dFrom) & "'"

        End Select
  
        cSQL.Execute strSQL
  
        If Not WriteToSQLTables(cSQL, strDestinationTable, strSourceQuery, dFrom) Then GoTo ExitProcessing
    End If
    
    'once function is true, UpdateQB will continue, else it will be rolledback in UpdateQB
    UpdateQBTable = True

ExitProcessing:
Finally:
    Call CloseMsgFrm
    Exit Function

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "UpdateQBTable", , ModName)
    Resume Finally

End Function

Private Function WriteToSQLTables(cSQL As ADODB.Connection, ByVal strDestinationTable As String, ByVal strQuery As String, Optional ByVal dFrom As Date = 0) As Boolean
    On Error GoTo Except
    
    ' write from odbc-linked QuickBooks tables directly to SQL, avoiding linked SQL table write time
    Dim db As DAO.Database
    Dim recordsQB As DAO.Recordset ' record set for ODBC-linked table queries
    Dim fldQB As DAO.Field
    
    Dim cSQLCmd As ADODB.Command
    
    Dim adorsSQL As ADODB.Recordset ' record set for SQL Server tables
    Dim adorsFld As ADODB.Field
    Dim prm As ADODB.Parameter
    
    Dim lParamType As Long
    Dim lParamSize As Long
    Dim lParamDirection As Long
    
    Dim strSQL As String
    Dim strColumns As String
    Dim strFunction As String
    Dim fldName As String
    
    WriteToSQLTables = False
    
    ' dedicated sql connection
    Set adorsSQL = New ADODB.Recordset
    Set cSQLCmd = New ADODB.Command
    
    cSQLCmd.ActiveConnection = cSQL
    cSQLCmd.CommandType = adCmdText
    
    ' define function using strQuery and dFrom date to pass to the function
    ' use eval to parse strFunction as a function
    Set db = CurrentDb
    If dFrom = 0 Then
        strFunction = strQuery
        Set recordsQB = db.OpenRecordset(strFunction, 2, 512)
    Else
        strFunction = strQuery & "(#" & dFrom & "#)"
        Set recordsQB = db.OpenRecordset(Eval(strFunction), 2, 512)
    End If
    
    ' open SQL table to get column info
    adorsSQL.open "SELECT * FROM " & strDestinationTable & " WHERE 1=0", cSQL, adOpenStatic, adLockReadOnly
    
    ' build parameter info
    lParamDirection = 1 ' adParamInput
    
    ' build insert statement
    strSQL = "INSERT INTO " & strDestinationTable & " ("
    strColumns = "VALUES ("
    
    For Each adorsFld In adorsSQL.Fields
        If strDestinationTable <> "tblQuickBooksItem" Then
            If IsInsertableField(adorsFld) Then
                strSQL = strSQL & "[" & adorsFld.Name & "], "
                strColumns = strColumns & "?, "
            End If
        Else
            strSQL = strSQL & "[" & adorsFld.Name & "], "
            strColumns = strColumns & "?, "
        End If
    Next
    
    ' create qualified insert statement with values
    strSQL = Left(strSQL, Len(strSQL) - 2) & ") "
    strColumns = Left(strColumns, Len(strColumns) - 2) & ")"
    
    ' set forth the command text as qualified insert statement
    cSQLCmd.CommandText = strSQL & strColumns
    
    ' loop through all query fields to translate into ADO parameter column types
    ' create and append parameters
    For Each adorsFld In adorsSQL.Fields
        If IsInsertableField(adorsFld) Then
            Set fldQB = recordsQB.Fields(adorsFld.Name)
                
            lParamSize = adorsFld.DefinedSize
            lParamType = GetADOTypeFromDAO(fldQB.Type)
                
            If UsesADOSize(lParamType) Then
                cSQLCmd.Parameters.Append cSQLCmd.CreateParameter(adorsFld.Name, lParamType, lParamDirection, lParamSize)
            ElseIf adorsFld.Type = adNumeric Then ' if is decimal (numeric as per ADO), define precision and scale
                lParamType = adorsFld.Type
                Set prm = cSQLCmd.CreateParameter(adorsFld.Name, lParamType, lParamDirection)
                prm.Precision = 18
                prm.NumericScale = 2
                cSQLCmd.Parameters.Append prm
            Else
                cSQLCmd.Parameters.Append cSQLCmd.CreateParameter(adorsFld.Name, lParamType, lParamDirection)
            End If
        End If
    Next
    
    ' for each field in the DAO recordset, set the cSQLParameter value
    Do While Not recordsQB.EOF
        For Each adorsFld In adorsSQL.Fields
            If IsInsertableField(adorsFld) Then
                fldName = adorsFld.Name
                cSQLCmd.Parameters(fldName).Value = recordsQB.Fields(fldName).Value
            End If
        Next
        cSQLCmd.Execute
        recordsQB.MoveNext
    Loop
    WriteToSQLTables = True
    
Finally:
    If Not adorsSQL Is Nothing Then Call CloseADORS(adorsSQL)
    
    If Not recordsQB Is Nothing Then
        If IsDAORecordsetOpen(recordsQB) Then recordsQB.Close
        Set recordsQB = Nothing
    End If
    
    If Not cSQLCmd Is Nothing Then Set cSQLCmd = Nothing
    Exit Function
    
Except:
    Select Case Err.Number
        Case -2147217887
            MsgBox "Could not parse information from " & strQuery & " to write to destination table " & strDestinationTable & " due to invalid precision / scale of decimal for a column in the table.", vbOKOnly + vbInformation, "WriteToSQLTables"
        Case Else
            Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, "WriteToSQLTables", , ModName)
    End Select
    Resume Finally

End Function
