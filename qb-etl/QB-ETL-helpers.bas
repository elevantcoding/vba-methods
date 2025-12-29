' ==========================================================
' QB ETL Helper Functions
'
' Supporting utilities for the QuickBooks → SQL Server
' ETL pipeline.
'
' This module provides:
'   - Type translation helpers between DAO and ADO
'   - SQL formatting utilities
'   - QuickBooks connectivity preflight checks
'   - Field filtering logic for schema-driven inserts
'
' Public Functions:
'   GetADOTypeFromDAO()    ' Maps DAO field types to ADO types
'   UsesADOSize()          ' Determines whether ADO type requires size
'   SQLDate()              ' Formats dates safely for SQL Server
'   IsInsertableField()    ' Filters computed / non-insertable columns
'   QBPreflight()          ' Validates QuickBooks connectivity before ETL
'
' External Dependencies:
'   QBConnection()         ' Returns ODBC connection string for QuickBooks
'   Example: FILEDSN=\\QuickBooksServer\quickbooksfilename.qbw.dsn;UID=youruser;PWD=yourpassword

' Requires References:
'   Microsoft ActiveX Data Objects x.x Library
'   Microsoft DAO x.x Object Library
'
' ==========================================================
Function GetADOTypeFromDAO(ByVal lDAOType As Long) As Long
    
    Dim lType As Long

    'get ado field type value for dao field type value (numeric)
    Select Case lDAOType
            'boolean
        Case 1: lType = adBoolean ' 11
            'byte
        Case 2: lType = adBinary ' 128
            'integer, long integer
        Case 3, 4: lType = adInteger ' 3
            'currency
        Case 5: lType = adCurrency ' 6
            'single
        Case 6: lType = adSingle ' 4
            'double
        Case 7: lType = adDouble ' 5
            'Date/Time
        Case 8: lType = adDBDate ' 133
            'binary, varbinary
        Case 9, 17: lType = adVarBinary ' 204
            'text
        Case 10: lType = adVarChar ' 200
            'long binary ole
        Case 11: lType = adLongVarBinary ' 205
            'memo
        Case 12: lType = adLongVarChar ' 201, use -1 for length
            'guid
        Case 15: lType = adGUID ' 72
            'big integer
        Case 16: lType = adBigInt ' 20, 2019 > forward
            'character
        Case 18: lType = adChar ' 129
            'numeric
        Case 19: lType = adNumeric ' 131
            'decimal
        Case 20: lType = adDecimal ' 14
            'float
        Case 21: lType = adDouble ' 5
            'time
        Case 22: lType = adDBTime ' 134
            'timestamp
        Case 23: lType = adDBTimeStamp ' 135
        Case Else: lType = 0
    End Select
    GetADOTypeFromDAO = lType
    
End Function

Function UsesADOSize(ByVal lADOType As Long) As Boolean
    
UsesADOSize = False
    
    Select Case lADOType
        Case 200, 201, 202
            UsesADOSize = True
            Exit Function
    End Select

End Function

Function SQLDate(ByVal d As Date) As String
          
    SQLDate = Format(d, "yyyy-mm-dd")

End Function

Function QBPreflight(ByVal strOrigin As String, Optional ByVal timeout As Long = 10) As Boolean
    On Error GoTo Except

    ' prior to attempting build of QB links, determine if QB connectivity
    ' check file availability by attempting to open the connection
    ' and query a known table
    Dim cn As ADODB.Connection: Set cn = New ADODB.Connection
    Dim adors As ADODB.Recordset: Set adors = New ADODB.Recordset
    Const ProcName As String = "QBPreflight"
    
    cn.ConnectionTimeout = timeout
    cn.open QBConnection
    cn.Execute "SELECT 1 FROM QBReportAdminGroup_v_AccountType"
    QBPreflight = True

Finally:
    Call CloseADORS(adors)
    If Not cn Is Nothing Then
        If cn.State = 1 Then cn.Close
        Set cn = Nothing
    End If
    Exit Function

Except:
    Call SystemFunctionRpt(Err.Number, Erl, Err.Description, Err.Source, ProcName, , ModName, , strOrigin, False, , False)
    Resume Finally
End Function

Private Function IsInsertableField(fld As ADODB.Field) As Boolean

    ' 16 is primary key attribute, JobCostMonth, computed column
    IsInsertableField = (fld.Attributes <> 16 And fld.Name <> "JobCostMonth")

End Function

Function IsDAORecordsetOpen(rs As DAO.Recordset) As Boolean
    
    ' used to detect whether a DAO recordset is still open
    On Error Resume Next
    Dim t As Long
    t = rs.Type
    IsDAORecordsetOpen = (Err.Number = 0)
    Err.Clear
    
End Function

