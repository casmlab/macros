macros
======

Excel macros, mostly for NodeXL and other data analysis

## CreateOneModeProjection 

Given a Two-Mode edgelist, create a One-Mode projection with edge weights.

### Setup and Usage ###

The macro expects (**requires**, really) a two-column spreadsheet with column headers where the first column is the node-type you want to keep and the second column is the other node-type. For instance, if you wanted a one-mode projection of actors where edges are based on their co-appearance in movies, you'd have a spreadsheet like so:

Actor | Movie
--------|---------
Jeremy Renner | American Hustle
Amy Adams | American Hustle
Amy Adams | Her

The easiest way to run the macro is to copy/paste the raw code from Github into the VisualBasic module screen in Excel. Here's Microsoft's explanation for creating macros: https://support.office.com/en-nz/article/Create-or-delete-a-macro-974ef220-f716-4e01-b015-3ea70e64937b

Once you have the macro in your workbook, run it: http://office.microsoft.com/en-us/excel-help/run-a-macro-HP010342865.aspx

After it runs, you can import the onemode_edgelist into NodeXL or whatever SNA package you're using.

### Output
Given a spreadsheet with two columns - VertexA and VertexB - the macro will create 3 sheets:

- two-mode matrix: a rectangular matrix with VertexA values as rows and VertexB values as columns
- onemode_matrix: a square matrix of VertexA values where the cells are edge weights
- onemode_edgelist: an edgelist of VertexA values, including an edgeweight column

