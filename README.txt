ExcelConverter
==============

In Digital Services, we worked a lot with excel files, but there were 
times in our workflow when we needed to convert an excel file to a 
different format.

The workflow for doing this consisted of opening 
the file in excel and using excel to convert the file to tab-delimited. This creates 
problems though, because when excel converts a file to a different file format, it 
takes any character within a cell that can be used as a delimiter (things like commas 
and quotations) andtreats those characters as delimiters, so everything in those cells 
gets put into quotations. It also puts quotations around cells that contain any sort of 
diacritic. What you get is a file full of un-wanted quotations mixed in with the good 
quotations. There is no easy workaround for this, so we would then have to open the 
converted file in Notepad ++, and use Notepad to delete all quotation marks. Finally, 
we would go back and add the quotations that were supposed to be in the file.

This process was time consuming and left a lot of room for human error. 
So I decided to create a script that makes this process efficient and 
removes the possibility of error. What I came up with was Excel Converter.

Excel Converter allows a user to convert any excel file to either a tab-delimited, unicode 
tab-delimited, comma-separated-values, or xml file. In our workflow, we mainly used the 
unicode tab-delimited option. That option also parses and removes only the unwanted quotes 
added by the Excel export.  All of the *good* quotes within a cell are retained.

Now looking under the hood, the script utilizes XlFileFormat Enumeration, which is a numerical 
value Microsoft uses to specify file type. For example, excel sees the number -4158 as a 
tab-delimited file. Knowing these values makes it easy to convert excel files to other file 
types. It would be easy to add more file types in the future as well. 
