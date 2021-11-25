# Date Input
Code for automatically time stamping data entered into Excel

This automatically inserts the present date on an adjacent cell (immediate right) when any text is entered into a cell within the specified range (here, those cells are specified by the headers of their respective columns within a table--"Applicant" and "Current_Stage").

The code itself duplicates the script, because there are two table columns, each named by their headers, which this checks for any new data input. Those can be substituted with column references (A1:A, e.g.).

The core of what the script should look like is the following:
```
    If Target.Cells.Count > 1 Then Exit Sub

    If Not Intersect(Target, Range("[Targetted range the user wants timestamps for]")) Is Nothing Then

        With Target([relative coordinate references of where the time should go, with 1 meaning that it 
        stays in the same column or row, and 2 meaning that it moves either down a row or to the left column])

        .Value = Date

        End With
        
    End If
