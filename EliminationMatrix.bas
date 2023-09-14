Attribute VB_Name = "EliminationMatrix"
Option Explicit


Sub Create_Elimination_Matrix()

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .AskToUpdateLinks = False
        .Calculation = xlCalculationManual
    End With
    
    'Variable Declaration
    Dim Comp_Type As String
    Dim FSO As Object
    Dim WrkBk_Input As Workbook: Dim WrkBk_Output As Workbook: Dim WrkBk_Repository As Workbook
    Dim Path As String: Dim Temp_Str As String: Dim Tab_Name As String: Dim Repo_TabName As String
    Dim Rng_PY As Range: Dim Cell As Range: Dim Rng As Range: Dim Fnd As Range: Dim Fnd_Rep As Range
    Dim LastRow As Long: Dim K As Integer: Dim Cnt As Integer
    Dim LastRow2 As Long
    
    'Open Input file
    MsgBox "Choose Input file First"
    Call File_Picker_Fun(Path, "Choose the input file")
    If Path <> "" Then
        Set WrkBk_Input = Workbooks.Open(Path)
    Else
        MsgBox "You did not choose the input file :Exiting Macro"
        GoTo ExitHere
    End If
    
    'Open Repository file
    Path = ""
    MsgBox "Choose the Repository file"
    Call File_Picker_Fun(Path, "Choose the Repository file")
    If Path <> "" Then
        Set WrkBk_Repository = Workbooks.Open(Path)
    Else
        MsgBox "You did not choose the Repository file :Exiting Macro"
        WrkBk_Input.Close
        GoTo ExitHere
    End If
    
    
    With ThisWorkbook.ActiveSheet
        If Not .Range("B7").Value = "Distribution" Or .Range("B7").Value = "Service" Then
            WrkBk_Input.Close
            WrkBk_Repository.Close
            MsgBox "Check Cell B7 Value, It Should be Distribution or Service: Existing Macro"
        End If
        
        Comp_Type = .Range("B7").Value
        Set Rng_PY = .Range("D6:D26") '.SpecialCells(xlCellValue)
        'MsgBox Rng_PY.Count
    End With
    
    'Choose input tab
    If Comp_Type = "Distribution" Then
        Tab_Name = "Distr_Dump"
        Repo_TabName = "Sample Distribution Set_EM"
    Else
        Tab_Name = "Services_Dump"
        Repo_TabName = "Sample Services Set_EM"
    End If
        
    'Check whether sheet is existing or not
    On Error Resume Next
    WrkBk_Input.Sheets(Tab_Name).Range("XX1").Value = WrkBk_Input.Sheets(Tab_Name).Range("XX1").Value
    WrkBk_Repository.Sheets(Repo_TabName).Range("XX1").Value = WrkBk_Repository.Sheets(Repo_TabName).Range("XX1").Value
    If Err.Number <> 0 Then
        MsgBox "1. Check Sheet name in the Input file" & vbNewLine & "If Type is Ditribution then sheet should be named as Distr_Dump else Services _Dump" & vbNewLine & _
                "2. Check Sheet name in the Output file" & vbNewLine & "If Type is Ditribution then sheet should be named as Sample Distribution Set_EM else Sample Services Set _EM" & vbNewLine & "Exiting Macro!!!"
        WrkBk_Input.Close
        WrkBk_Output.Close
        On Error GoTo 0
        GoTo ExitHere:
    End If
    On Error GoTo 0
    
    'Setup Output file
    Set WrkBk_Output = Workbooks.Add
    With WrkBk_Output
        .SaveAs (ThisWorkbook.Path & "\" & Comp_Type & " " & Date & ".xlsx")
        ThisWorkbook.Sheets("Sample").Copy Before:=.Sheets("Sheet1")
        .Sheets("Sample").Visible = True
        .Sheets("Sheet1").Name = "Reconcillation Table"
        .Sheets("Reconcillation Table").Range("A1").Value = "Company Name"
        .Sheets("Reconcillation Table").Range("B1").Value = "Status"
        .Sheets("Reconcillation Table").Range("C1").Value = "Comments"
        .Sheets("Sample").Name = Comp_Type
        .Sheets(Comp_Type).Range("AC1").Value = "Type of " & Comp_Type
    End With
    
    'Copy Data from input file
    With WrkBk_Input.Sheets(Tab_Name)
    
        'All
        .Activate
        If .FilterMode = True Then .ShowAllData
        '.AutoFilter.Sort.SortFields.Add2 Key:=Range("C2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("A2:AF" & LastRow).Sort Key1:=Range("C2"), Order1:=xlAscending, Header:=xlYes
        
        'Copy data
        .Range("C3:C" & LastRow).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("B5")
        .Range("D3:D" & LastRow).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("O5")
        .Range("E3:E" & LastRow).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("T5")
        .Range("G3:G" & LastRow).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("U5")
        .Range("H3:H" & LastRow).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("V5")
        .Range("I3:I" & LastRow).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("W5")
        .Range("J3:J" & LastRow).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("X5")
        .Range("K3:K" & LastRow).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("Y5")
        .Range("M3:M" & LastRow).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("C5")
        .Range("Z3:Z" & LastRow).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("Z5")
        .Range("AA3:AA" & LastRow).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("D5")
        LastRow2 = (LastRow - 2) + 4
        
        'Reject
        If .FilterMode = True Then .ShowAllData
        
        On Error Resume Next
        .Range("$A$2:$AE$" & LastRow).AutoFilter Field:=32, Criteria1:="Reject"
        If Err.Number <> 0 Then
            .Range("A2").AutoFilter
            .Range("$A$2:$AE$" & LastRow).AutoFilter Field:=32, Criteria1:="Reject"
        End If
        
        On Error GoTo 0
        Set Rng = .Range("C3:C" & LastRow).SpecialCells(xlCellTypeVisible)
        
        For Each Cell In Rng
            Set Fnd = WrkBk_Output.Sheets(Comp_Type).Range("B5:B" & LastRow2).Find(what:=Cell.Value)
            .Range("AB" & Cell.Row).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("E" & Fnd.Row)
            .Range("AC" & Cell.Row).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("J" & Fnd.Row)
            .Range("AD" & Cell.Row).Copy
            WrkBk_Output.Sheets(Comp_Type).Range("S" & Fnd.Row).PasteSpecial Paste:=xlPasteValues
            .Range("AE" & Cell.Row).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("P" & Fnd.Row)
            .Range("AF" & Cell.Row).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("N" & Fnd.Row)
        Next
        
        Set Rng = Nothing
        Set Cell = Nothing
        Set Fnd = Nothing
        
        '<>Reject
        If .FilterMode = True Then .ShowAllData
        .Range("$A$2:$AE$" & LastRow).AutoFilter Field:=32, Criteria1:="<>Reject"
        Set Rng = .Range("C3:C" & LastRow).SpecialCells(xlCellTypeVisible)
    End With
    
    'Copy data from Repository file
    With WrkBk_Repository.Sheets(Repo_TabName)
        .Activate
        If .FilterMode = True Then .ShowAllData
        For Each Cell In Rng
            Set Fnd_Rep = WrkBk_Output.Sheets(Comp_Type).Range("B5:B100000").Find(what:=Cell.Value)
            Set Fnd = .Range("B5:B100000").Find(what:=Cell.Value)
            If Not Fnd Is Nothing Then
                .Range("F" & Fnd.Row & ":" & "N" & Fnd.Row).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("F" & Fnd_Rep.Row)
                .Range("P" & Fnd.Row & ":" & "R" & Fnd.Row).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("P" & Fnd_Rep.Row)
                .Range("AA" & Fnd.Row & ":" & "AD" & Fnd.Row).Copy Destination:=WrkBk_Output.Sheets(Comp_Type).Range("AA" & Fnd_Rep.Row)
            Else
                WrkBk_Output.Sheets(Comp_Type).Range("AE" & Fnd_Rep.Row) = "New companies"
            End If
            Set Fnd = Nothing
            Set Fnd_Rep = Nothing
        Next Cell
    End With
    
    
    With WrkBk_Output.Sheets(Comp_Type)
        Set Rng = .Range("B5:B" & LastRow2)
        For Each Cell In Rng
            Set Fnd = WrkBk_Repository.Sheets(Repo_TabName).Range("B5:B100000").Find(what:=Cell.Value)
            If Not Fnd Is Nothing Then
                If .Range("O" & Cell.Row).Value <> WrkBk_Repository.Sheets(Repo_TabName).Range("O" & Fnd.Row).Value Then
                   If .Range("AE" & Cell.Row).Value = "" Then
                        .Range("AE" & Cell.Row).Value = "To review"
                    Else
                        .Range("AE" & Cell.Row).Value = .Range("AE" & Cell.Row).Value & ", " & "To review"
                    End If
                End If
            End If
        Next
        
        Set Rng = .Range("O5:O" & LastRow2)
        For Each Cell In Rng
        'Subsidiary
            Set Fnd = Cell.Find(what:="*Subsidiary*")
                If Not Fnd Is Nothing Then
                    If .Range("AE" & Cell.Row).Value = "" Then
                        .Range("AE" & Cell.Row).Value = "Subsidiary"
                    Else
                        .Range("AE" & Cell.Row).Value = .Range("AE" & Cell.Row).Value & ", " & "Subsidiary"
                    End If
                    Set Fnd = Nothing
                    GoTo NextCell
                End If
            
        'Subsidiaries
            Set Fnd = Cell.Find(what:="*Subsidiaries*")
                If Not Fnd Is Nothing Then
                    If .Range("AE" & Cell.Row).Value = "" Then
                        .Range("AE" & Cell.Row).Value = "Subsidiaries"
                    Else
                        .Range("AE" & Cell.Row).Value = .Range("AE" & Cell.Row).Value & ", " & "Subsidiaries"
                    End If
                    Set Fnd = Nothing
                    GoTo NextCell
                End If
                
                
        'Merger
            Set Fnd = Cell.Find(what:="*Merger*")
                If Not Fnd Is Nothing Then
                    If .Range("AE" & Cell.Row).Value = "" Then
                        .Range("AE" & Cell.Row).Value = "Merger"
                    Else
                        .Range("AE" & Cell.Row).Value = .Range("AE" & Cell.Row).Value & ", " & "Merger"
                    End If
                    Set Fnd = Nothing
                    GoTo NextCell
                End If
                
        'Jointly owned
            Set Fnd = Cell.Find(what:="*Jointly owned*")
                If Not Fnd Is Nothing Then
                    If .Range("AE" & Cell.Row).Value = "" Then
                        .Range("AE" & Cell.Row).Value = "Jointly owned"
                    Else
                        .Range("AE" & Cell.Row).Value = .Range("AE" & Cell.Row).Value & ", " & "Jointly owned"
                    End If
                    Set Fnd = Nothing
                    GoTo NextCell
                End If
NextCell:
        Next
    End With
    
    WrkBk_Repository.Close
    WrkBk_Input.Close
        
    'Work with Output file
    K = 1
    With WrkBk_Output
        
        'Main Sheet
        With .Sheets(Comp_Type)
            .Activate
            .Rows(LastRow2 + 1 & ":10000").EntireRow.Delete
            .Range("A5").Value = 1
            .Range("A5").AutoFill Destination:=Range("A5:A" & LastRow2), Type:=xlFillSeries
            .Range("A5:AD" & LastRow2).BorderAround xlContinuous
            .Range("A5:AD" & LastRow2).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Range("A5:AD" & LastRow2).Borders(xlInsideVertical).LineStyle = xlContinuous
            
            With .Range("B5:D" & LastRow2)
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With
            
            With .Range("O5:AD" & LastRow2)
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With
            
            With .Range("E5:N" & LastRow2)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            With .Range("Y5:Y" & LastRow2)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            .Cells.EntireColumn.AutoFit
            .Columns("A:A").ColumnWidth = 7
            .Columns("B:D").ColumnWidth = 28
            .Columns("G:N").ColumnWidth = 7
            .Columns("O:Q").ColumnWidth = 40
            .Columns("X:Y").ColumnWidth = 7
        End With
        
        'Reconcillation Table Sheet
        For Each Cell In Rng_PY
            If Not IsEmpty(Cell) Then
                .Sheets("Reconcillation Table").Range("A" & K + 1).Value = Cell.Value
                Cnt = WorksheetFunction.CountIf(.Sheets(Comp_Type).Range("B5:B" & LastRow2), Cell.Value)
                If Cnt >= 1 Then
                    .Sheets("Reconcillation Table").Range("B" & K + 1).Value = "Yes"
                Else
                    .Sheets("Reconcillation Table").Range("B" & K + 1).Value = "No"
                End If
                K = K + 1
            End If
            Cnt = 0
        Next Cell
        
        If K > 1 Then
            With .Sheets("Reconcillation Table").Range("A1:C" & K)
                .BorderAround xlContinuous
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
            End With
            With .Sheets("Reconcillation Table")
                .Cells.EntireColumn.AutoFit
            End With
        End If
    End With
    
    WrkBk_Output.Save
    MsgBox "Done"
ExitHere:

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .AskToUpdateLinks = True
        .Calculation = xlCalculationAutomatic
    End With
    
End Sub

'The Functionality of this Fun is to Select a file object
Function File_Picker_Fun(StrFile As String, Title_Str As String)

Dim FD As Office.FileDialog
Set FD = Application.FileDialog(msoFileDialogFilePicker)

    With FD
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx?", 1
        .Title = Title_Str
        .AllowMultiSelect = False
        .InitialFileName = "C:\VBA Folder"
        If .Show = True Then
            StrFile = .SelectedItems(1)
        End If
    End With

End Function
