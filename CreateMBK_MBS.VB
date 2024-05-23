Sub CreateMBK_MBS_Table()
    Dim wsMBK As Worksheet
    Dim wsMBS As Worksheet
    Dim lastRowMBK As Long
    Dim lastRowMBS As Long
    Dim i As Long
    Dim j As Long
    Dim mbkDescription As String
    Dim mbsDescription As String
    Dim mbkXU As String
    Dim newRow As Long
    
    ' Set worksheet
    Set wsMBK = ThisWorkbook.Sheets("MBK")
 
    lastRowMBK = wsMBK.Cells(wsMBK.Rows.Count, "A").End(xlUp).Row
    
    ' Create or clear destination sheet
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("MBK+MBS")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDest.Name = "MBK+MBS"
    Else
        wsDest.Cells.Clear
    End If
    On Error GoTo 0
    
    ' headers
    wsDest.Cells(1, 2).Value = "MBK Part Number"
    wsDest.Cells(1, 3).Value = "MBK Description"
    wsDest.Cells(1, 4).Value = "MBS Part Number"
    
    ' 1U
    newRow = newRow + 1
    wsDest.Cells(newRow, 1).Value = "1U"
    newRow = newRow + 1
    For i = 2 To lastRowMBK
        mbkDescription = wsMBK.Cells(i, 3).Value
        mbkXU = ExtractMotherboard(mbkDescription)
        
        If InStr(1, mbkDescription, "1U", vbTextCompare) > 0 Then
            ' Populate the MBK information
            wsDest.Cells(newRow, 2).Value = wsMBK.Cells(i, 1).Value
            wsDest.Cells(newRow, 3).Value = mbkDescription
            wsDest.Cells(newRow, 4).Value = mbkXU
            newRow = newRow + 1
        End If
    Next i
    
     ' 2U
    newRow = newRow + 1
    wsDest.Cells(newRow, 1).Value = "2U"
    newRow = newRow + 1
    For i = 2 To lastRowMBK
        mbkDescription = wsMBK.Cells(i, 3).Value
        mbkXU = ExtractMotherboard(mbkDescription)
        
        If InStr(1, mbkDescription, "2U", vbTextCompare) > 0 Then
            ' Populate the MBK information
            wsDest.Cells(newRow, 2).Value = wsMBK.Cells(i, 1).Value
            wsDest.Cells(newRow, 3).Value = mbkDescription
            wsDest.Cells(newRow, 4).Value = mbkXU
            newRow = newRow + 1
        End If
    Next i
    
     ' 3U
    newRow = newRow + 1
    wsDest.Cells(newRow, 1).Value = "3U"
    newRow = newRow + 1
    For i = 2 To lastRowMBK
        mbkDescription = wsMBK.Cells(i, 3).Value
        mbkXU = ExtractMotherboard(mbkDescription)
          
        
        If InStr(1, mbkDescription, "3U", vbTextCompare) > 0 Then
            ' Populate the MBK information
            wsDest.Cells(newRow, 2).Value = wsMBK.Cells(i, 1).Value
            wsDest.Cells(newRow, 3).Value = mbkDescription
            wsDest.Cells(newRow, 4).Value = mbkXU
            newRow = newRow + 1
        End If
    Next i
    
     ' 4U
    newRow = newRow + 1
    wsDest.Cells(newRow, 1).Value = "4U"
    newRow = newRow + 1
    For i = 2 To lastRowMBK
        mbkDescription = wsMBK.Cells(i, 3).Value
        mbkXU = ExtractMotherboard(mbkDescription)
        
        If InStr(1, mbkDescription, "4U", vbTextCompare) > 0 Then
            ' Populate the MBK information
            wsDest.Cells(newRow, 2).Value = wsMBK.Cells(i, 1).Value
            wsDest.Cells(newRow, 3).Value = mbkDescription
            wsDest.Cells(newRow, 4).Value = mbkXU
            newRow = newRow + 1
        End If
    Next i
    

    wsDest.Columns.AutoFit
End Sub

Function ExtractMotherboard(description As String) As String
    parts = Split(description, ",")
    
    If UBound(parts) >= 1 Then
        motherboardName = Trim(parts(1))
    Else
        Exit Function
    End If
    
    If Not motherboardName Like "*MOTHERBOARD*" Then
        Dim mbsSheet As Worksheet
        Set mbsSheet = ThisWorkbook.Sheets("MBS")
        
        Dim lastRow As Integer
        lastRow = mbsSheet.Cells(mbsSheet.Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To lastRow
            If InStr(1, mbsSheet.Cells(i, 3).Value, motherboardName, vbTextCompare) > 0 Then
                ExtractMotherboard = mbsSheet.Cells(i, 1).Value
                Exit Function
            End If
        Next i
        
        ExtractMotherboard = ""
    Else
        ExtractMotherboard = ""
    End If
End Function
