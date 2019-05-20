# sheet2HTML
A tool to take a spreadsheet and a template and produce the relevant data from it

This is a browser-based tool to convert the data found in a spreadsheet into HTML.

So far it can:
* Accept a dropped XLSX file
* Preview the data
* Unmerge cells (copying the data into each cell)
* Reduce multiple header rows (user specified) into 1 row
* Generate data using mustache templates
* Save the modified file to the Downloads file

It needs to:
* Provide a way to filter and select the columns to display
* Rename column headers, file name and sheet name in the output
* Output in other formats
* Support functions | utilities for
    * Articles (use a | an) correctly
    * Stripping whitespace
    * Case transformations
* Be able to compare versions
    * Submit multiple files at the same time
    * Save old versions in the browser
* Join multiple files together
