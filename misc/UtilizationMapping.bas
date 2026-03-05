-- traverse project for FindValue 
-- uses RegEx helper functions, bitwise enums
-- bitwise enums allows multiple enums to be passed in a single parameter
-- looks in tables, queries (excluding ~sq), forms, reports, modules
-- excludes ~sq queries because form traversal will report FindValue found in row-source queries

-- options:
-- scope (SearchScope)
-- create text file of results found
-- display notification during search (custom, not in this module)
-- exit upon first count of FindValue
-- search declaration lines only

' used in UtliizationV1 for identifying
' areas of the system to search for
' something
' bitwise
Public Enum SearchScope
    ssTables = 1
    ssQueries = 2
    ssForms = 4
    ssReports = 8
    ssModules = 16
End Enum

' search options for UtilizationV1
' bitwise
Public Enum SearchOptions
    soNone = 0
    soDecLinesOnly = 1
End Enum

' exe option for UtilizationV1
Public Enum ExecutionOptions
    eoNone = 0
    eoWriteResults = 1
    eoDisplayProgress = 2
    eoExitOnFirstFind = 4
End Enum

' bitwise sum of SearchScope options
Public Const ssAll = ssTables Or ssQueries Or ssForms Or ssReports Or ssModules


Function UtilizationV1(ByVal FindValue As String, Optional ByVal Scope As SearchScope = ssAll, Optional ByVal Exec As ExecutionOptions = eoNone, _
          Optional Options As SearchOptions = soNone) As Long
    On Error GoTo Except
    
    Const ProcName As String = "UtilizationV1"
    
    ' returns count of FindValue found in specified objects
    
    ' uses public enum SearchScope, SearchOptions, ExecutionOptions as bit values with Public Const ssAll:
    ' allows more than one search scope to be supplied to the Scope parameter using Or
    ' ssAll is not specified in the proc as an option becauase it is understood as
    ' the sum of all possibilities
    
    Dim db As DAO.Database, tdf As DAO.TableDef, qdf As DAO.QueryDef, fld As DAO.Field, prp As DAO.Property
    Dim obj As AccessObject, blnClose As Boolean, frm As Access.Form, Ctl As Control, rpt As Access.Report
    Dim vbc As VBIDE.VBComponent, cmod As VBIDE.CodeModule, CodeModLine As Long, CountLines As Long, Content As String, ModName As String, ReportedModName As String, ModProcName As String, ReportedModProcName As String
    Dim Stream As Object, FileName As String, FilePath As String, CollectionExists As Boolean, FindValueCount As Long
    Dim WriteResults As Boolean, DisplayProgress As Boolean, ExitOnFirstFind As Boolean
    
    FindValue = Trim(FindValue)
    
    If Len(FindValue) = 0 Then Exit Function
    
    ' if all, forms or reports, notify regarding nested objects
    If ((Scope And ssForms) = ssForms) Or ((Scope And ssReports) = ssReports) Then
        If MsgBox("For this utility to properly inspect form and report objects, please close forms / reports containing nested objects before running this utility. " & _
            "Do you want to continue?", vbYesNo + vbQuestion, "Continue?") = vbNo Then
            Exit Function
        End If
    End If
    
    ' default value count
    FindValueCount = 0
    
    ' execution options
    WriteResults = (Exec And eoWriteResults) = eoWriteResults
    DisplayProgress = (Exec And eoDisplayProgress) = eoDisplayProgress
    ExitOnFirstFind = (Exec And eoExitOnFirstFind) = eoExitOnFirstFind
    
    ' create text file only if WriteResults
    If WriteResults Then
        FilePath = Environ("USERPROFILE") & "\Desktop\"
        FileName = "FindValue.txt"
  
        ' FSO global file system object reference
        Set Stream = FSO.CreateTextFile(FilePath + FileName)
    
        Stream.WriteLine "Utilization Report for " & FindValue
        Stream.WriteLine "-----------------------------------"
    End If
    
    Set db = CurrentDb
    
    ' bitwise: for search scope tables (or all)
    If (Scope And ssTables) = ssTables Then
    
        If WriteResults Then
            Stream.WriteLine "Tables"
            Stream.WriteLine "-----------------------------------"
        End If
  
        If DisplayProgress Then
            DisplayMsg , , "Reviewing Tables"
            DoCmd.RepaintObject acForm, MsgFrm
            DoEvents
        End If
  
        ' find value as a table name
        CollectionExists = False
        For Each tdf In db.TableDefs
            CollectionExists = True
            If tdf.Name = FindValue Then
                If WriteResults Then Stream.WriteLine "Is a Named Table"
                FindValueCount = FindValueCount + 1
                If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
            End If
  
            ' find value as a table column name
            For Each fld In tdf.Fields
                If fld.Name = FindValue Then
                    If WriteResults Then Stream.WriteLine "Is a Named Column In " & tdf.Name
                    FindValueCount = FindValueCount + 1
                    If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                End If
          
                ' find value in a column row source
                For Each prp In fld.Properties
                    If prp.Name = "RowSource" Then
                        If Nz(prp.Value, "") <> "" Then
                            If IsWholeValue(prp.Value, FindValue) Then
                                If WriteResults Then Stream.WriteLine "RowSource of " & fld.Name & " in " & tdf.Name
                                FindValueCount = FindValueCount + 1
                                If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                            End If
                        End If
                    End If
              
                    ' find value as a column default value
                    If prp.Name = "DefaultValue" Then
                        If Nz(prp.Value, "") <> "" Then
                            If IsWholeValue(prp.Value, FindValue) Then
                                If WriteResults Then Stream.WriteLine "Default Value of " & fld.Name & " in " & tdf.Name
                                FindValueCount = FindValueCount + 1
                                If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                            End If
                        End If
                    End If
                Next
            Next
        Next
    
        If WriteResults Then
            If Not CollectionExists Then
                Stream.WriteLine "No Tables"
            End If
            Stream.WriteBlankLines 2
        End If
    End If
    
    ' bitwise: for search scope forms (or all)
    If (Scope And ssQueries) = ssQueries Then
    
        If WriteResults Then
            Stream.WriteLine "Queries"
            Stream.WriteLine "-----------------------------------"
        End If
    
        If DisplayProgress Then
            DisplayMsg , , "Reviewing Queries"
            DoCmd.RepaintObject acForm, MsgFrm
            DoEvents
        End If
  
        ' looping through rowsources of controls so no need to view ~sq_ queries
    
        CollectionExists = False
        For Each qdf In db.QueryDefs
            CollectionExists = True
            If Left(qdf.Name, 4) <> "~sq_" Then
          
                ' find value as query name
                If qdf.Name = FindValue Then
                    If WriteResults Then Stream.WriteLine "Is a Named Query"
                    FindValueCount = FindValueCount + 1
                    If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                End If
          
                ' find value as a query column name
                For Each fld In qdf.Fields
                    If fld.Name = FindValue Then
                        If WriteResults Then Stream.WriteLine "Is a Named Column in " & qdf.Name
                        FindValueCount = FindValueCount + 1
                        If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                    End If
                Next
  
                ' find value in query sql (column name, parameter, computed column)
                ' will return results for fld.name, as well
                If IsWholeValue(qdf.SQL, FindValue) Then
                    If WriteResults Then Stream.WriteLine "In SQL of " & qdf.Name
                    FindValueCount = FindValueCount + 1
                    If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                End If
            End If
        Next
    
    
        If WriteResults Then
            If Not CollectionExists Then
                Stream.WriteLine "No Queries"
            End If
            Stream.WriteBlankLines 2
        End If
    End If
    
    ' bitwise: for search scope forms (or all)
    If (Scope And ssForms) = ssForms Then

        If WriteResults Then
            Stream.WriteLine "Forms"
            Stream.WriteLine "-----------------------------------"
        End If
    
        If DisplayProgress Then
            DisplayMsg , , "Reviewing Forms"
            DoCmd.RepaintObject acForm, MsgFrm
            DoEvents
        End If
  
        ' find value as form name
        CollectionExists = False
        For Each obj In Application.CurrentProject.AllForms
            CollectionExists = True
            If obj.Name = FindValue Then
                If WriteResults Then Stream.WriteLine "Is a Named Form"
                FindValueCount = FindValueCount + 1
                If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
            End If
  
            ' open forms, if not already open
            blnClose = False
            If Not obj.IsLoaded Then
                DoCmd.OpenForm obj.Name, acDesign, , , , acHidden
                blnClose = True
            End If
  
            Set frm = Forms(obj.Name)
  
            'inspect form RecordSource, control's ControlSource and control's RowSource
            For Each prp In frm.Properties
                If prp.Name = "RecordSource" Then
                    If Nz(prp.Value, "") <> "" Then
                        If IsWholeValue(prp.Value, FindValue) Then
                            If WriteResults Then Stream.WriteLine "In RecordSource of " & frm.Name
                            FindValueCount = FindValueCount + 1
                            If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                        End If
                    End If
                End If
            Next
  
            For Each Ctl In frm.Controls
                Select Case Ctl.ControlType
                    Case acTextBox, acComboBox, acListBox
                        For Each prp In Ctl.Properties
                            If prp.Name = "ControlSource" Or prp.Name = "RowSource" Then
                                If Nz(prp.Value, "") <> "" Then
                                    If IsWholeValue(prp.Value, FindValue) Then
                                        If WriteResults Then Stream.WriteLine "In Control or RowSource of " & Ctl.Name & " of " & frm.Name
                                        FindValueCount = FindValueCount + 1
                                        If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                                    End If
                                End If
                            End If
                        Next
                    Case acSubform
                        If Ctl.Name = FindValue Then
                            If WriteResults Then Stream.WriteLine "Is a SubForm on " & frm.Name
                            FindValueCount = FindValueCount + 1
                            If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                        End If
                End Select
            Next
    
            ' close form if flagged
            If blnClose Then DoCmd.Close acForm, obj.Name
        Next
        If WriteResults Then
            If Not CollectionExists Then
                Stream.WriteLine "No Forms"
            End If
            Stream.WriteBlankLines 2
        End If
    End If
    
    ' bitwise: for search scope reports (or all)
    If (Scope And ssReports) = ssReports Then
    
        If WriteResults Then
            Stream.WriteLine "Reports"
            Stream.WriteLine "-----------------------------------"
        End If
  
        If DisplayProgress Then
            DisplayMsg , , "Reviewing Reports"
            DoCmd.RepaintObject acForm, MsgFrm
            DoEvents
        End If
    
        CollectionExists = False
        For Each obj In Application.CurrentProject.AllReports
            CollectionExists = True
    
            If obj.Name = FindValue Then
                If WriteResults Then Stream.WriteLine "Is a Named Report"
                FindValueCount = FindValueCount + 1
                If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
            End If
  
            ' open reports, if not already open
            blnClose = False
            If Not obj.IsLoaded Then
                DoCmd.OpenReport obj.Name, acDesign, , , , acHidden
                blnClose = True
            End If
  
            Set rpt = Reports(obj.Name)
  
            'inspect report RecordSource, control's ControlSource and control's RowSource
            For Each prp In rpt.Properties
                If prp.Name = "RecordSource" Then
                    If Nz(prp.Value, "") <> "" Then
                        If IsWholeValue(prp.Value, FindValue) Then
                            If WriteResults Then Stream.WriteLine "In RecordSource of " & rpt.Name
                            FindValueCount = FindValueCount + 1
                            If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                        End If
                    End If
                End If
            Next
  
            For Each Ctl In rpt.Controls
                Select Case Ctl.ControlType
                    Case acTextBox, acComboBox, acListBox
                        For Each prp In Ctl.Properties
                            If prp.Name = "ControlSource" Or prp.Name = "RowSource" Then
                                If Nz(prp.Value, "") <> "" Then
                                    If IsWholeValue(prp.Value, FindValue) Then
                                        If WriteResults Then Stream.WriteLine "In Control or RowSource of " & Ctl.Name & " of " & rpt.Name
                                        FindValueCount = FindValueCount + 1
                                        If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                                    End If
                                End If
                            End If
                        Next
                    Case acSubform
                        If Ctl.Name = FindValue Then
                            If WriteResults Then Stream.WriteLine "Is a SubReport on " & rpt.Name
                            FindValueCount = FindValueCount + 1
                            If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                        End If
                End Select
            Next
    
            ' close rpt if flagged
            If blnClose Then DoCmd.Close acReport, obj.Name
        Next
    
        If WriteResults Then
            If Not CollectionExists Then
                Stream.WriteLine "No Reports"
            End If
            Stream.WriteBlankLines 2
        End If
    End If
    
    ' bitwise
    If (Scope And ssModules) = ssModules Then
        If WriteResults Then
            Stream.WriteLine "Modules"
            Stream.WriteLine "-----------------------------------"
        End If
    
        If DisplayProgress Then
            DisplayMsg , , "Reviewing Modules"
            DoCmd.RepaintObject acForm, MsgFrm
            DoEvents
        End If
    
        ' look in modules
        ' no collection compare because there is a module count
        For Each vbc In Application.VBE.ActiveVBProject.VBComponents
            ReportedModName = ""
            Set cmod = vbc.CodeModule
      
            ' Options: bitwise
            If (Options And soDecLinesOnly) = soDecLinesOnly Then
                CountLines = CLng(Nz(cmod.CountOfDeclarationLines, 0)) ' only look in code declarations lines
            Else
                CountLines = CLng(Nz(cmod.CountOfLines, 0)) ' look in all lines
            End If
      
            If CountLines > 0 Then
                ReportedModProcName = ""
                For CodeModLine = 1 To CountLines
                    Content = cmod.Lines(CodeModLine, 1)
                    If Left(Trim(Content), 1) <> "'" Then
                        If IsWholeValue(Content, FindValue) Then
                            ModName = vbc.Name
                            ModProcName = cmod.ProcOfLine(CodeModLine, vbext_pk_Proc)
                                        
                            If ModName <> ReportedModName Then
                                If WriteResults Then
                                    Stream.WriteLine "------------------"
                                    Stream.WriteLine "Module " & ModName
                                    ReportedModName = ModName
                                End If
                            End If
                            
                            If ReportedModProcName <> ModProcName Then
                                If WriteResults Then Stream.WriteLine vbTab & "Procedure " & ModProcName
                                If WriteResults Then Stream.WriteLine vbTab & vbTab & Content
                                ReportedModProcName = ModProcName
                            End If
                            
                            FindValueCount = FindValueCount + 1
                            If ExitOnFirstFind Then: UtilizationV1 = FindValueCount: GoTo ExitProcessing
                        End If
                    End If
                Next
            End If
        Next
    End If
    
    ' return results
    If UtilizationV1 = 0 Then UtilizationV1 = FindValueCount
    
ExitProcessing:
Finally:
    ' close displayfrm
    If DisplayProgress Then CloseMsgFrm
    
    ' report completed, open
    If WriteResults Then
        If Not Stream Is Nothing Then
            On Error Resume Next
            Stream.Close
            Set Stream = Nothing
        End If
        CreateObject("WScript.Shell").Run """" & FilePath & FileName & """", 1, False
    End If
    Exit Function
    
Except:
    ReportExcept Erl, Err.Number, Err.Description, ProcName, ModName
    Resume Finally
End Function

Private Function EscRegex(ByVal s As String) As String

    ' escape chars in regex
    Static re As Object

    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = False
        re.pattern = "([\\\^\$\.\|\?\*\+\(\)\[\]\{\}])"
    End If

    EscRegex = re.Replace(s, "\$1")
    
End Function
Private Function IsWholeValue(ByVal SourceText As String, ByVal FindValue As String) As Boolean
    
    Static re As Object          ' cached RegExp
    Static lastPattern As String ' track pattern changes
    
    Dim pattern As String
    
    If Len(SourceText) = 0 Or Len(FindValue) = 0 Then Exit Function

    ' build the regex pattern
    pattern = "(^|[^A-Za-z0-9_])" & EscRegex(FindValue) & "([^A-Za-z0-9_]|$)"

    ' create or refresh regex only when needed
    If re Is Nothing Or lastPattern <> pattern Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = False
        re.IgnoreCase = True
        re.pattern = pattern
        lastPattern = pattern
    End If

    IsWholeValue = re.Test(SourceText)

End Function
