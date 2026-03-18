# vba-methods

A collection of production-grade VBA modules for real-world systems:
ETL pipelines, SQL Server integration, security tooling, and advanced
automation built in Microsoft Access / VBA.

## Contents

### ADO Utilities
Centralized ADO command execution framework.

**SQLCmdGlobal** — Global ADO Command executor with 
parameterized stored procedure and SQL statement support. 
Enum-driven execution model (execute immediately or defer 
to caller). Handles parameter type validation, decimal 
precision/scale, and connection lifecycle management.

**ADOResult** — Scalar lookup wrapper enforcing SELECT-only 
queries with single-row validation.

Location: `/ado`

### DAO Utilities

**DAOTruncate** — Transactional Access table truncation 
with AutoNumber reseed. Validates relationships, column 
type, and user confirmation before executing DELETE and 
ALTER TABLE as an atomic operation.

Location: `/dao`

### QuickBooks → SQL Server ETL
High-performance ETL pipeline that:
- Reads ODBC-linked QuickBooks data
- Builds schema-driven INSERT statements
- Streams records using parameterized ADO commands
- Executes inside a single SQL transaction with rollback protection

Location: `/qb-etl`

###  SQL Server →  QuickBooks ETL
CreateQBFile
- VBA that generates QuickBooks-importable .IIF files from SQL data for batch transaction import

Location: `/qb-etl`

### Security & Encryption
Reversible string obfuscation system with cross-language
interoperability.

Location: `/security`

### Miscellaneous Utilities
**UseDictionary** — Enum-driven Dictionary lifecycle 
manager. Single function handles create, read, write, 
exists check, remove, count, and cleanup via ByRef 
pattern with Boolean success return.

**UtilizationV1** — Static analysis engine for Access 
applications. Searches tables, queries, forms, reports, 
and VBA modules for any value — including inside Dynamic 
SQL strings that standard dependency tools cannot detect. 
Bitwise scope and option parameters. Outputs to text file 
or returns count.

Location: `/misc`
### Windows API Utilities

Location: /winapi

Low-level Windows helpers implemented in VBA using Win32 API calls.

Includes:

Idle Time Detection
Retrieves machine-level idle time (seconds since last user input) using the Windows API.
Useful for session management, inactivity monitoring, and graceful application shutdown in long-running VBA applications.

## Design Principles
- Runtime schema introspection — no hardcoded column lists
- Eval()-based dynamic function resolution — eliminates 
  form reference dependencies in compiled deployments
- Transaction safety
- Schema-driven logic
- Explicit error boundaries
- Portable SQL generation
- Maintainable, testable VBA

This repository demonstrates how VBA can be used to build reliable,
enterprise-grade systems — not just macros.

