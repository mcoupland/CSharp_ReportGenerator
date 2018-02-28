# CSharp_ReportGenerator
# Michael Coupland 2/28/2018

## Summary
* C# (WPF) application to display employees, managers, and divisions.  Provide search function as well as ability to easily click through from employee to senior manager (L4). Uses Excel interop, LINQ, File I/O, generics, delegates, event handling, and asynchronous UI updates via SynchronizationContext.

## Overview
* Setup
  * Read configuration text file and load settings
  * Read L4 source text file
* Load employee data
  * Read employee data from Excel file
  * Asyncronously import data and update UI
  * Raise custom progress and completed events
* Search
  * Perform simple "contains" search over employee list
  * Update UI with search results
  * Inline event handler delegate
  * Add grid elements from code
