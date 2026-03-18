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

