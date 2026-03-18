# miscellaneous vba
# GlobalDictionary

A utility module for Microsoft Access / VBA that simplifies 
Dictionary object instantiation and lifecycle management.

## Problem It Solves
Declaring, instantiating, and cleaning up Scripting.Dictionary 
objects requires repetitive boilerplate across every procedure 
that uses one. This module provides a single global function 
that handles instantiation via ByRef, returning a ready-to-use 
Dictionary object and automatically managing cleanup.

## Usage
Pass an uninitialized Dictionary variable ByRef. The function 
sets it up with your preferred CompareMode and returns it 
ready to use. No Set/New boilerplate. No forgotten cleanup.

## Requirements
Microsoft Scripting Runtime reference (or late binding variant 
included)

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
