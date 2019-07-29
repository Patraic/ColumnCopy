Sub Get_Columns()
    Dim sPath As String
    Dim sFil As String
    Dim owb As Workbook
    Dim twb As Workbook

    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set twb = ThisWorkbook
    sPath = ThisWorkbook.Path & "\"
    sFil = Dir(sPath & "*.xlsx")

        Dim i As Integer: i = 1
Do While sFil <> "" And sFil <> twb.Name
    Set owb = Workbooks.Open(sPath & sFil)
    With owb.Sheets("Wyniki").Range("A1:A35")
    twb.Sheets("Podsumowanie").Cells(1, Columns.Count).End(xlToLeft).Offset(, 1).Resize(.Rows.Count, .Columns.Count).Value = .Value
    End With
        owb.Close False 'Close no save
        sFil = Dir
        
    Loop

        With Application
        .Calculation = xlAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub
