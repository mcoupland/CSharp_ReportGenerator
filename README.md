# CSharp_ReportGenerator
# Michael Coupland 2/28/2018

## Overview
* C# (WPF) application to import Excel data, format it according to JSON configuration files, and create new, formatted, Excel reports.  Features drag and drop fields to allow user to include, exclude, and re-arrange report fields if user wants a slightly different report than that defined by the JSON file.

### Load report configuration
* Load all JSON files in configuration directory
* Deserialize using NewtonSoft JSON library
* Create custom "button" object for all reports

### Display report fields
* Update UI
* Create "button" object for all fields
* Add drag and drop behavior to buttons
* Load selected buttons into panel
* Load remaining buttons into seperate panel

### Create report
* Asyncronous cross thread progress updates
* Custom event handlers and event arg objects
* Format Excel report
  * Background color
  * Wrap text
  * Auto fit
  * Borders
* Excel printing options
  * Paper size
  * Orientation
  * Fit to width, height
  * Zoom
