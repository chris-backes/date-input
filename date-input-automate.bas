VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'This automatically inserts the present date on an adjacent cell (immediate right) when any text is entered into a cell within the specified range (here, those
'cells are specified by the headers of their respective columns within a table--"Applicant" and "Current_Stage".

Private Sub Worksheet_Change(ByVal Target As Range)

    If Target.Cells.Count > 1 Then Exit Sub

    If Not Intersect(Target, Range("Applicant")) Is Nothing Then

        With Target(1, 2)

        .Value = Date

        .EntireColumn.AutoFit

        End With

    End If
    
    If Target.Cells.Count > 1 Then Exit Sub

    If Not Intersect(Target, Range("Current_Stage")) Is Nothing Then

        With Target(1, 2)

        .Value = Date

        .EntireColumn.AutoFit

        End With
        
    End If
        
End Sub
