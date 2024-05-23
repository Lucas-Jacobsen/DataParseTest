Sub CreateMBS_PRO_Table()
    Dim wsMBS As Worksheet
    Dim wsPRO As Worksheet
    Dim wsDest As Worksheet
    Dim lastRowMBS As Long
    Dim lastRowPRO As Long
    Dim i As Long
    Dim j As Long
    Dim lgaSocket As String
    Dim partDescription As String
    Dim newRow As Long
    Dim processorFound As Boolean
    
    ' Set worksheets
    Set wsMBS = ThisWorkbook.Sheets("MBS")
    Set wsPRO = ThisWorkbook.Sheets("PRO")
    
    lastRowMBS = wsMBS.Cells(wsMBS.Rows.Count, "A").End(xlUp).Row
    lastRowPRO = wsPRO.Cells(wsPRO.Rows.Count, "A").End(xlUp).Row
    
    ' Create or clear destination sheet
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("MBS+PRO")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDest.Name = "MBS+PRO"
    Else
        wsDest.Cells.Clear
    End If
    On Error GoTo 0
    
    ' headers
    wsDest.Cells(1, 1).Value = "Motherboard Part Number"
    wsDest.Cells(1, 2).Value = "Motherboard Description"
    wsDest.Cells(1, 3).Value = "Processor Part Number"
    wsDest.Cells(1, 4).Value = "Processor Description"
    
    ' Loop through each row in the MBS sheet
    newRow = 2
    For i = 2 To lastRowMBS
        partDescription = wsMBS.Cells(i, 3).Value
        
        ' Extract LGA socket from the description
        lgaSocket = ExtractLGASocket(partDescription)
        
        ' If LGA socket is found, find compatible processors
        If lgaSocket <> "" Then
            processorFound = False
            
            ' Populate the motherboard information
            wsDest.Cells(newRow, 1).Value = wsMBS.Cells(i, 1).Value
            wsDest.Cells(newRow, 2).Value = partDescription
            newRow = newRow + 1
            
            ' Add processors that match the LGA socket
            For j = 2 To lastRowPRO
                If InStr(1, wsPRO.Cells(j, 3).Value, lgaSocket, vbTextCompare) > 0 Then
                    wsDest.Cells(newRow, 3).Value = wsPRO.Cells(j, 1).Value
                    wsDest.Cells(newRow, 4).Value = wsPRO.Cells(j, 3).Value
                    newRow = newRow + 1
                    processorFound = True
                End If
            Next j
            
            ' no processors found
            If Not processorFound Then
                wsDest.Cells(newRow, 3).Value = "No compatible processors found"
                newRow = newRow + 1
            End If
            
            newRow = newRow + 1
        End If
    Next i
    
    wsDest.Columns.AutoFit
End Sub

Function ExtractLGASocket(partDescription As String) As String
Dim lgaSocket As String
Dim lgaPattern As String
Dim regex As Object
Set regex = CreateObject("VBScript.RegExp")

lgaPattern = "LGA\s?\d{2,4}" ' Pattern to match LGA socket number

With regex
    .Global = True
    .Pattern = lgaPattern
    If .Test(partDescription) Then
        ExtractLGASocket = .Execute(partDescription)(0)
    Else
        ExtractLGASocket = ""
    End If
End With
End Function
