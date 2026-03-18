# miscellaneous vba
# GlobalDictionary

A VBA utility module that provides a single, unified interface 
for all Scripting.Dictionary operations using an enum-driven 
action pattern.

## Problem It Solves
Working with Dictionary objects in VBA requires repetitive 
boilerplate: declare the variable, create the object, check 
existence before adding, handle cleanup. Across a large 
project this scatters dictionary lifecycle management 
throughout the codebase. GlobalDictionary centralizes all 
dictionary operations into a single function call.

## Design
UseDictionary accepts a ByRef dictionary object and a 
DictActions enum value, performing one operation per call 
and returning True on success. The ByRef pattern means the 
dictionary is instantiated and managed through the function 
without requiring the caller to use Set or New directly.

## DictActions Enum

| Action       | Description                                      |
|--------------|--------------------------------------------------|
| daNew        | Create a new Dictionary, clearing any existing   |
| daInit       | Check whether dictionary is initialized          |
| daAddEntry   | Add key/value or update value if key exists      |
| daExists     | Check whether a key exists                       |
| daRead       | Read a value by key, returned via ByRef KeyValue |
| daRemoveEntry| Remove a key/value pair if it exists             |
| daCount      | Return True if dictionary has any entries        |
| daClose      | RemoveAll and Set to Nothing — full cleanup      |

## Usage Example
Dim myDict As Object
UseDictionary myDict, daNew
UseDictionary myDict, daAddEntry, "ProjectID", 12345
UseDictionary myDict, daRead, "ProjectID", myValue
UseDictionary myDict, daClose

## Key Behaviors
- daAddEntry safely handles both new keys and updates 
  to existing keys without requiring an Exists check 
  at the call site
- daRead returns the value via ByRef KeyValue parameter
- daClose performs full cleanup — RemoveAll then 
  Set Nothing — preventing memory leaks
- Guards against calling operational actions on an 
  uninitialized dictionary — raises custom message 
  rather than runtime error

## Requirements
Microsoft Scripting Runtime (early binding) or 
CreateObject late binding (as implemented)
Custom error handler (ReportExcept)
Custom message raising (RaiseCustomMsg)

InterpolateColor - steps through color hues from white to desired RGB, includes form module code example of use (used for splash screen)

CompareStrings - compare two string values and return a % difference to help avoid near-duplicates and for inference

# UtilizationMappings

A static analysis engine for Microsoft Access applications. 
Searches across all object types in an Access project to find 
every reference to a given value.

## Problem It Solves
In large Access applications, renaming a table, column, or 
procedure requires knowing every place it is referenced — 
including form RecordSources, control RowSources, report 
bindings, and VBA code. Standard dependency tools miss 
dynamic SQL references entirely. This engine finds them all.

## Search Scope (Bitwise)
- Tables — names, column names, row sources, default values
- Queries — names, column names, SQL definitions
- Forms — record sources, control sources, row sources
- Reports — record sources, control sources
- Modules — line-by-line traversal with procedure identification,
  comment skipping, and declaration-only or body-only options

## Options
- Write results to text file
- Display progress
- Exit on first find

## Requirements
Microsoft VBE Extensibility reference
Microsoft Access project (not ADP)
