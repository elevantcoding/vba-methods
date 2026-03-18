# DAOTruncate

A VBA/Access utility that replicates T-SQL TRUNCATE TABLE 
behavior for Microsoft Access local tables — deleting all 
records AND resetting the AutoNumber counter in a single 
atomic transaction.

## Problem It Solves
Access has no native TRUNCATE TABLE equivalent. DELETE FROM 
removes records but leaves the AutoNumber counter at its last 
value. Resetting the counter requires a separate ALTER TABLE 
statement. If either operation succeeds without the other, 
the table is in an inconsistent state. DAOTruncate wraps 
both operations in a DAO transaction — they succeed together 
or roll back together.

## Safety Checks
Before executing, DAOTruncate validates:
- Table exists in the current database
- Table has no established relationships (referential 
  integrity would prevent deletion)
- Named column exists in the table
- Named column is an AutoNumber field
- User confirms the operation via prompt

## Error Handling
- Error 3211 (table locked / in use) is caught and reported 
  gracefully without raising an unhandled exception
- Any open transaction at exit is rolled back via the 
  Finally block — preventing partial state
- Brackets are added to table and column names automatically 
  to handle reserved words and spaces

## Best For
Staging tables, temporary data stores, and import/ETL 
tables that need to be cleared and reloaded with a 
fresh identity sequence.

## Not Suitable For
Tables with established relationships. DAOTruncate will 
detect and refuse these — use DELETE with WHERE clause 
instead.

## Helper Functions Included
- TableExists — checks AllTables collection
- TableIsRelated — iterates db.Relations for table 
  participation as primary or foreign
- IsAutoNumber — checks dbAutoIncrField attribute via 
  bitwise AND
- AddBrackets — safely wraps identifiers in square 
  brackets if not already present
