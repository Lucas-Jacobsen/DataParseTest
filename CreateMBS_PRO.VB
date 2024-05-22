Sub CreateMBS_PRO_Dropdown()
    Dim wsMBS As Worksheet
    Dim wsPRO As Worksheet
    Dim wsDest As Worksheet
    Dim lastRowMBS As Long
    Dim lastRowPRO As Long
    Dim i As Long
    Dim j As Long
    Dim lgaSocket As String
    Dim processorList As String
    Dim partDescription As String
    Dim newRow As Long
    
    ' Set worksheets
    Set wsMBS = ThisWorkbook.Sheets("MBS")
    Set wsPRO = ThisWorkbook.Sheets("PRO")
    
    ' Find the last rows
    lastRowMBS = wsMBS.Cells(wsMBS.Rows.Count, "A").End(xlUp).Row
    lastRowPRO = wsPRO.Cells(wsPRO.Rows.Count, "A").End(xlUp).Row
    
    ' Create or clear the destination sheet
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("MBS+PRO")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDest.Name = "MBS+PRO"
    Else
        wsDest.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Add headers to the destination sheet
    wsDest.Cells(1, 1).Value = "Motherboard Part Number"
    wsDest.Cells(1, 2).Value = "Motherboard Description"
    wsDest.Cells(1, 3).Value = "Compatible Processors"
    
    ' Loop through each row in the MBS sheet
    newRow = 2
    For i = 2 To lastRowMBS
        partDescription = wsMBS.Cells(i, 3).Value
        
        ' Extract LGA socket from the description
        lgaSocket = ExtractLGASocket(partDescription)
        
        ' If LGA socket is found, find compatible processors
        If lgaSocket <> "" Then
            processorList = ""
            For j = 2 To lastRowPRO
                If InStr(1, wsPRO.Cells(j, 3).Value, lgaSocket, vbTextCompare) > 0 Then
                    processorList = processorList & wsPRO.Cells(j, 3).Value & vbCrLf
                End If
            Next j
            
            ' Populate the destination sheet
            If processorList <> "" Then
                wsDest.Cells(newRow, 1).Value = wsMBS.Cells(i, 1).Value
                wsDest.Cells(newRow, 2).Value = partDescription
                With wsDest.Cells(newRow, 3).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                    xlBetween, Formula1:=processorList
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = True
                End With
                wsDest.Cells(newRow, 3).Value = "Select Processor"
                newRow = newRow + 1
            End If
        End If
    Next i
    
    ' Autofit columns for better appearance
    wsDest.Columns.AutoFit
End Sub

Function ExtractLGASocket(description As String) As String
    Dim lgaPos As Long
    Dim lgaSocket As String
    
    lgaPos = InStr(1, description, "LGA", vbTextCompare)
    If lgaPos > 0 Then
        lgaSocket = Mid(description, lgaPos, 8) ' Assuming LGA socket format "LGAxxxx"
        ExtractLGASocket = lgaSocket
    Else
        ExtractLGASocket = ""
    End If
End Function
