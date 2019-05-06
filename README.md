# sheet2HTML
A tool to take a spreadsheet and a template and produce the relevant data from it

This is a browser-based tool to convert the data found in a spreadsheet into HTML.

So far it can:
* Accept a dropped XLSX file
* Preview the data
* Unmerge cells (copying the data into each cell)
* Reduce multiple header rows (user specified) into 1
* Generate data using mustache templates

It needs to:
* Output the results. This should be prettymuch boilerplate. Waiting on the issue below.
* Provide a nice way of filtering and selecting the data needed
* Rename headers in the output
