# copyMacro
Purpose: Copy info from select files (specified template) and add to a master file.

Template is predefined and pre-filled. Macro will create a "Completed" folder. Open each file, copy specified information to master list, close file and move to compelted folder.

Sub Begin(): Runs through all files with .xls* extensions, excludes master file that macro is stored in. Checks for a "Completed" folder, creates one if none available. Opens each template file, calls Copy(). Closes file, does not save changes. Completes when there are no files left. If template file is protected, macro will close with no changes, does not move file.

Sub Copy(): Unmerges merged cells, stores values in variables. Adds to masterfile in appropriate columns, totals specified fields. Empties variables, ends sub.
