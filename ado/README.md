# mSQL — Global ADO Command Module for VBA/Access + SQL Server

A centralized ADO execution framework for Microsoft Access 
applications connected to SQL Server or Azure SQL.

## Problem It Solves
Every ADO call in a large Access/SQL application requires 
the same boilerplate: open connection, create command, set 
command type, append parameters, execute, clean up. Across 
hundreds of procedures this becomes a maintenance liability. 
mSQL centralizes all of this into a single reusable module.

## Core Functions

### SQLCmdGlobal
Global ADO Command executor. Accepts command text (SQL 
statement or stored procedure name), command type, execution 
method, and parameters via ParamArray using ADOParam helpers.
Supports two execution modes:
- emOrigin: execute immediately within SQLCmdGlobal
- emCaller: set up command and parameters only; 
  calling procedure handles execution

### ADOParam
Helper function that builds a parameter array for passing 
to SQLCmdGlobal. Supports all ADO data types including 
adDecimal with precision and scale validation.

### ADOResult  
Wrapper for single-value scalar lookups. Enforces SELECT-only 
queries, returns the first column of the first row, and raises 
an error if multiple rows are returned — preventing silent 
data integrity issues.

### SQLCmdAsType
Handles connection validation, command instantiation, 
parameter clearing, and connection assignment. Automatically 
opens the global connection if not already open.

### OpenSQL
Global connection manager. Checks connection state before 
re-instantiating to prevent redundant connections.

## Design Decisions
- ByRef Cmd parameter allows command object to be reused 
  or inspected by the calling procedure after execution
- ADOParamResolve handles both direct calls and wrapper 
  function calls transparently — flattening nested ParamArrays
- Parameter validation catches zero-length string types, 
  invalid decimal precision/scale, and malformed param arrays 
  before they reach SQL Server
- Named parameters enabled for stored procedures only

## Requirements
Microsoft ActiveX Data Objects (ADODB) reference
Global ADOConnect connection string variable
Global error handler (ReportExcept)
