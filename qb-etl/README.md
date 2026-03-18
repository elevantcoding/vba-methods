# QB_ETL — QuickBooks to SQL Server ETL Pipeline

A production-grade VBA/ADO ETL engine that synchronizes 
data from ODBC-linked QuickBooks Enterprise Desktop tables 
into SQL Server / Azure SQL, and generates QuickBooks-
importable IIF files from SQL Server data.

## Background
Built for a real construction company environment where 
QuickBooks Enterprise Desktop served as the accounting 
system of record and a custom Azure SQL ERP required 
synchronized financial data for job costing, payroll, 
backlog reporting, and executive dashboards.

## What It Does

### UpdateQB — Primary ETL Orchestrator
Synchronizes QuickBooks data into SQL Server within a 
single ADO transaction. If any table update fails, the 
entire batch rolls back — ensuring no partial state.

Tables synchronized:
- Accounts (tblQuickBooksAccounts)
- Items (tblQuickBooksItem)
- Credit Card Transactions (tblQuickBooksCreditCards)
- Revenue — Invoices and Credit Memos (tblQuickBooksRevenue)
- Receipts (tblQuickBooksReceipts)
- Permits (tblQuickBooksPermits)

### UpdateQBTable — Per-Table ETL Handler
Resolves the appropriate VBA dataset function at runtime 
using Eval(), deletes existing records from the target 
date forward, and writes fresh data. Handles Invoice and 
Credit Memo separation within the same destination table 
using conditional DELETE logic.

### WriteToSQLTables — Schema-Driven Insert Engine
Reads destination table schema at runtime to build 
parameterized INSERT statements dynamically. Maps DAO 
field types from QuickBooks ODBC to ADO parameter types. 
Handles decimal precision/scale separately. Uses dual 
recordsets — DAO for QuickBooks source, ADO for SQL 
Server destination — for maximum throughput.

Post-commit record count validation confirms expected 
versus actual rows written — raising an alert on mismatch 
without rolling back the committed transaction.

### CreateQBFile — IIF Export Generator
Exports allocated credit card transactions from SQL Server 
into a QuickBooks-compatible IIF file for batch import. 
Writes in configurable batch sizes (default 200 rows) 
to manage memory. Outputs correctly formatted IIF headers 
(HDR, TRNS, SPL, ENDTRNS) and renames the file from 
.txt to .iif for QuickBooks compatibility.

## Key Design Decisions
- No saved Access queries — all SQL is generated at 
  runtime from VBA functions, protecting business logic 
  in compiled .accde deployment
- Eval() resolves dataset function names dynamically 
  at runtime, eliminating [Forms]![...] dependencies
- Single ADO transaction wraps all table updates — 
  full rollback on any failure
- Schema introspection at runtime means the insert 
  engine adapts to table structure without hardcoding 
  column lists
- Post-commit count validation adds a reconciliation 
  layer after the transaction closes

## Requirements
- Microsoft ActiveX Data Objects (ADODB)
- Microsoft DAO Object Library
- ODBC connection to QuickBooks Enterprise Desktop
- ADOConnect — global SQL Server connection string
- DisplayMsg, MsgFrm, CloseMsgFrm — UI messaging helpers
- QBPreflight — QuickBooks connectivity check
- ConfirmMatchingContracts, ReviewCC — pre-flight validators

- # QB_ETL_helpers — Supporting Utilities for QB_ETL

Helper functions that support the QuickBooks → SQL Server 
ETL pipeline. Handles type translation, SQL formatting, 
connectivity validation, and field filtering.

## Functions

### GetADOTypeFromDAO
Maps DAO field type constants to their ADO equivalents.
Covers the full range of field types encountered in 
QuickBooks ODBC data — boolean, integer, currency, 
single, double, date, binary, text, memo, GUID, bigint, 
numeric, decimal, and timestamp. Returns 0 for unmapped 
types, allowing the caller to handle edge cases.

### UsesADOSize
Returns True if an ADO type requires a Size parameter 
when creating an ADO Parameter object (adVarChar, 
adLongVarChar, adVarWChar). Prevents runtime errors 
from size-dependent types receiving zero-length 
parameter definitions.

### SQLDate
Formats a VBA Date value as yyyy-mm-dd for safe 
inclusion in SQL Server WHERE clauses and DELETE 
statements — avoiding locale-dependent date format 
ambiguity.

### QBPreflight
Validates QuickBooks ODBC connectivity before the ETL 
pipeline begins. Opens a connection with a configurable 
timeout (default 10 seconds) and executes a lightweight 
probe query against a known QuickBooks table. Returns 
True only if both the connection and query succeed. 
Logs silently on failure without interrupting the 
caller — allowing UpdateQB to present a clean user 
message rather than an unhandled error.

### IsInsertableField
Filters out non-insertable columns from schema-driven 
INSERT generation. Excludes identity/primary key fields 
(Attributes = 16) and computed columns by name 
(JobCostMonth). Keeps the insert engine generic while 
accommodating table-specific exceptions.

### IsDAORecordsetOpen
Safely detects whether a DAO Recordset is still open 
by probing the Type property and checking for an error. 
Used in Finally blocks to conditionally close recordsets 
without raising errors on already-closed objects.

## Notes
- SQLDate uses yyyy-mm-dd format specifically to avoid 
  SQL Server ambiguity with mm/dd/yyyy in non-US locales
- QBPreflight logs silently (False on final param) so 
  the caller controls user-facing messaging
- IsInsertableField's hardcoded "JobCostMonth" exclusion 
  is intentional — it's a computed column in the 
  destination table that cannot receive an INSERT value
