# vba-methods

A collection of production-grade VBA modules for real-world systems:
ETL pipelines, SQL Server integration, security tooling, and advanced
automation built in Microsoft Access / VBA.

## Contents

### QuickBooks → SQL Server ETL
High-performance ETL pipeline that:
- Reads ODBC-linked QuickBooks data
- Builds schema-driven INSERT statements
- Streams records using parameterized ADO commands
- Executes inside a single SQL transaction with rollback protection

Location: `/qb-etl`

###  SQL Server →  QuickBooks ETL
CreateQBFile

Location: `/qb-etl`

### Security & Encryption
Reversible string obfuscation and encryption system with cross-language
interoperability.

Location: `/security`

## Design Principles
- Transaction safety
- Schema-driven logic
- Explicit error boundaries
- Portable SQL generation
- Maintainable, testable VBA

This repository demonstrates how VBA can be used to build reliable,
enterprise-grade systems — not just macros.
