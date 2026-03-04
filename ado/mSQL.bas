' publicly-available vars for SQL connection
Public SQLConnect As ADODB.Connection
Public SQLCmd As ADODB.Command
Public SQLPrm As ADODB.Parameter

' for use in SQLCmdGlobal
Public Enum CmdExecMethod
    emOrigin ' Execute SQL immediately within SQLCmdGlobal
    emCaller ' Setup Command and Params only; execution handled by the calling proc
End Enum

' global ADO Command function, returns true if successful
' --------------------------------------------------------------------------------
' if call SQLCmdGlobal directly, call params using ADOParam inside of Array
' (i.e., vParams as Variant: vParams = Array(ADOParam(name, type, direction, size, value), ADOParam(name, type, direction, size, value)), SQLCmdGlobal(SQL, Cmd, acCmdText, emOrigin, vParams)

' if call SQLCmdGlobal from wrapper, call params using ADOParam
' (i.e, SQLCmdGlobal(SQL, Cmd, acCmdText, emOrigin, ADOParam(name, type, direction, size, value))
' (i.e. SQLCmdGlobal(SQL, Cmd, acCmdStoredProc, emOrigin, ADOParam(name, type, direction, size, value))
' --------------------------------------------------------------------------------
' see helper function in this module: ADOParam
' parameter array for receiving ADO parameters
' p(0), p(1), p(2), p(3), p(4) represent args in .CreateParameter(name, type, direction, size, value)

Public Function SQLCmdGlobal(ByVal CmdText As String, ByRef Cmd As ADODB.Command, ByVal CmdType As ADODB.CommandTypeEnum, _
          CmdMethod As CmdExecMethod, ParamArray CmdParams() As Variant) As Boolean
    On Error GoTo Except

    Const ProcName As String = "SQLCmdGlobal"
    Dim i As Long, p As Variant, ResolveParams As Variant
    
    Dim MsgDetail As String: MsgDetail = "When calling CmdParams " & CmdText
    
    ' ByRef Cmd passes the specified command obj through the call stack: is modifiable by SQLCmdAsType as ByRef as well for use in this proc
    ' designed to be called using ADOParam if sp or cmd text has params
    ' CmdExecMethod: Public Enum in this module: emOrigin = execute from SQLCmdGlobal, emCaller = execute from calling procedure.
    ' default
    SQLCmdGlobal = False
    
    ' ensure stored proc name not zero length string
    CmdText = Trim(CmdText)
    If Len(CmdText) = 0 Then Exit Function
    
    ' set up the ADO Command
    SQLCmdAsType Cmd, CmdType
    
    With Cmd

        ' CmdText = name of stored proc OR is an SQL statement    
        .CommandText = CmdText

        ' if is from wrapper or direct call to SQLCmdGlobal
        ResolveParams = ADOParamResolve(CmdParams)
                   
        ' if has params
        If UBound(ResolveParams) >= 0 Then
            
            ' loop through param array
            For i = LBound(ResolveParams) To UBound(ResolveParams)
                
                ' loop each array
                p = ResolveParams(i)
                
                ' make sure is array
                If Not IsArray(p) Then RaiseCustomMsg SysArray, ProcName, MsgDetail
                
                ' expect between five and 7 params, lower bound 0
                If Not IsBetween(UBound(p), 4, 6) Then RaiseCustomMsg SysSQLSPParams, ProcName, MsgDetail
                
                ' if prm type is string and prm size is 0, raise err
                If IsIn(p(1), adVarChar, adVarWChar) And p(3) = 0 Then RaiseCustomMsg SysSQLSPParamsSize, ProcName, MsgDetail
                
                ' set parameters
                Set SQLPrm = .CreateParameter(p(0), p(1), p(2), p(3), p(4))
                
                ' if is decimal param, set precision and scale
                If p(1) = adDecimal Then
                    SQLPrm.Precision = p(5)
                    SQLPrm.NumericScale = p(6)
                End If
                
                ' append
                .Parameters.Append SQLPrm
            Next
        End If
        
        ' exec if method is to execute from this proc
        If CmdMethod = emOrigin Then .Execute
    
        ' function successfully exec
        ' results are available in cmd object .Parameters(varname)
        SQLCmdGlobal = True
    End With

Finally:
    Exit Function

Except:
    ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Function

' helper: returns a variant to be used in the ParamArray of SQLCmdGlobal
Public Function ADOParam(ByVal PrmName As String, ByVal PrmType As ADODB.DataTypeEnum, ByVal PrmDir As ADODB.ParameterDirectionEnum, _
          ByVal PrmSize As Long, ByVal PrmVal As Variant, Optional DecPrecision As Long = 0, Optional DecScale As Long = 0) As Variant
    On Error GoTo Except
    
    Const ProcName As String = "ADOParam"
    
    ' if decimal,
    ' precision must not be less than or equal to zero,
    ' scale must not be less than 0,
    ' scale must not be greater than precision
    If PrmType = adDecimal Then
        If DecPrecision <= 0 _
            Or DecScale < 0 _
            Or DecScale > DecPrecision Then
            RaiseCustomMsg SysSQLSPParamsDecimal, ProcName
        End If
        ADOParam = Array(PrmName, PrmType, PrmDir, PrmSize, PrmVal, DecPrecision, DecScale)
    Else
        ADOParam = Array(PrmName, PrmType, PrmDir, PrmSize, PrmVal)
    End If

Finally:
    Exit Function

Except:
    ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Function

' helper function: determines if params passed in paramarray to SQLCmdGlobal are an array or variant
Public Function ADOParamResolve(ByVal Params As Variant) As Variant
    On Error GoTo Except
    Const ProcName As String = "ADOParamResolve"
    
    ' if calling SQLCmdGlobal from a wrapper function, "flatten parameters":
    ' if the first element is an array, it's a nested ParamArray from a wrapper function
    If IsArray(Params(0)) Then
        ' return inner array
        ADOParamResolve = Params(0)
    Else
        ADOParamResolve = Params
    End If
    
Finally:
    Exit Function
Except:
    ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Function

' Cmd is ByRef so reference can be modified as needed
Sub SQLCmdAsType(Optional ByRef Cmd As ADODB.Command = Nothing, Optional ByVal CmdType As ADODB.CommandTypeEnum = adCmdStoredProc, Optional ByVal lTimeout As Long = 90)
    On Error GoTo Except
    
    Const ProcName As String = "SQLCmdAsType"

    ' open the global SQL connection if it is not open
    If Not ValidSQLConnect Then OpenSQL

    ' if no Cmd passed to proc, use global SQLCmd
    If Cmd Is Nothing Then
        If SQLCmd Is Nothing Then Set SQLCmd = New ADODB.Command
        Set Cmd = SQLCmd
    End If
    
    ' clear params (from any previous calls)
    With Cmd
        If .Parameters.Count > 0 Then
            Do Until .Parameters.Count = 0
                .Parameters.Delete 0
            Loop
        End If
    
        .ActiveConnection = SQLConnect ' global ADODB.Connection
        .CommandType = CmdType
        .CommandTimeout = lTimeout
        .NamedParameters = (CmdType = adCmdStoredProc) ' only named parameters if is stored proc
    End With
    
Finally:
    Exit Sub
Except:
    ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Sub

' wrapper that uses SQLCmdGlobal
' return a one-row result (scalar, for aggregates, one-row lookups, etc.)
Public Function ADOResult(ByVal SQL As String, ByVal CmdType As ADODB.CommandTypeEnum, _
          ParamArray Params() As Variant) As Variant
    On Error GoTo Except
    
    Const ProcName As String = "ADOResult"
    Const Sel As String = "SELECT "
    Dim adors As ADODB.Recordset

    ADOResult = Null
    
    SQL = Trim(SQL)
    
    If Left(SQL, 7) <> Sel Then RaiseCustomMsg SQLNoSelect, ProcName
    
    ' use SQLCmdGlobal to call the recordset
    If Not SQLCmdGlobal(SQL, SQLCmd, CmdType, emCaller, Params) Then RaiseCustomMsg SysSQLCmdGlobal, ProcName, SQL
        
    Set adors = SQLCmd.Execute ' SQLCmd is global
        
    If Not adors.EOF Then
        ADOResult = adors.Fields(0).Value ' first column
        adors.MoveNext ' see if more than one row
        If Not adors.EOF Then
            ADOResult = Null ' if more than one row, result is null, raise msg
            RaiseCustomMsg SQLMultipleResults, ProcName, SQL
        End If
    End If

Finally:
    If Not adors Is Nothing Then
        If adors.State = adStateOpen Then adors.Close
        Set adors = Nothing
    End If
    If Not SQLCmd Is Nothing Then
        Set SQLCmd = Nothing
    End If
    Exit Function
Except:
    ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Function

' open the global connection
Sub OpenSQL(Optional ByVal AttemptOnly As Boolean = False)
    On Error GoTo Except
    Const ProcName As String = "OpenSQL"
    
    ' connect to SQL
    
    ' if the connection exists, and is open, exit
    ' else re-instantiate
    If Not SQLConnect Is Nothing Then
        If SQLConnect.State = adStateOpen Then Exit Sub
        Set SQLConnect = Nothing
    End If
    
    Set SQLConnect = New ADODB.Connection
    SQLConnect.ConnectionTimeout = 20
    SQLConnect.Open ADOConnect ' connection string
 
Finally:
    Exit Sub

Except:
    If Not AttemptOnly Then ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Sub

' helper function: used in SQLCmdGlobal
Public Function IsBetween(ByVal evalNum As Double, ByVal valOne As Double, ByVal valTwo As Double) As Boolean
    Dim val As Double
    If valOne > valTwo Then
        val = valOne
        valOne = valTwo
        valTwo = val
    End If
        
    IsBetween = (evalNum >= valOne And evalNum <= valTwo)
End Function

-- helper function: used in SQLCmdGlobal to detect whether 
Public Function IsIn(ByVal ValComp As Variant, ParamArray Vals() As Variant) As Boolean
    Dim i As Long
    
    For i = LBound(Vals) To UBound(Vals)
        If Vals(i) = ValComp Then
            IsIn = True
            Exit Function
        End If
    Next
    
    IsIn = False    
End Function


