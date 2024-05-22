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
            
            ' If no processors found, indicate so
            If Not processorFound Then
                wsDest.Cells(newRow, 3).Value = "No compatible processors found"
                newRow = newRow + 1
            End If
            
            ' Add an empty row for separation
            newRow = newRow + 1
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
        lgaSocket = Mid(description, lgaPos, 8)
        ExtractLGASocket = lgaSocket
    Else
        ExtractLGASocket = ""
    End If
    Range("B276").Select
    ActiveWorkbook.Save
    Sheets("MMP").Select
    Columns("A:A").EntireColumn.AutoFit
    Sheets("MBS+PRO").Select
    Range("B266").Select
    ActiveWindow.SmallScroll Down:=0
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.ScrollRow = 576
    ActiveWindow.ScrollRow = 578
    ActiveWindow.ScrollRow = 581
    ActiveWindow.ScrollRow = 589
    ActiveWindow.ScrollRow = 592
    ActiveWindow.ScrollRow = 594
    ActiveWindow.ScrollRow = 597
    ActiveWindow.ScrollRow = 605
    ActiveWindow.ScrollRow = 607
    ActiveWindow.ScrollRow = 613
    ActiveWindow.ScrollRow = 616
    ActiveWindow.ScrollRow = 621
    ActiveWindow.ScrollRow = 628
    ActiveWindow.ScrollRow = 631
    ActiveWindow.ScrollRow = 632
    ActiveWindow.ScrollRow = 637
    ActiveWindow.ScrollRow = 640
    ActiveWindow.ScrollRow = 648
    ActiveWindow.ScrollRow = 653
    ActiveWindow.ScrollRow = 656
    ActiveWindow.ScrollRow = 658
    ActiveWindow.ScrollRow = 663
    ActiveWindow.ScrollRow = 666
    ActiveWindow.ScrollRow = 669
    ActiveWindow.ScrollRow = 672
    ActiveWindow.ScrollRow = 676
    ActiveWindow.ScrollRow = 679
    ActiveWindow.ScrollRow = 682
    ActiveWindow.ScrollRow = 684
    ActiveWindow.ScrollRow = 685
    ActiveWindow.ScrollRow = 688
    ActiveWindow.ScrollRow = 692
    ActiveWindow.ScrollRow = 693
    ActiveWindow.ScrollRow = 697
    ActiveWindow.ScrollRow = 698
    ActiveWindow.ScrollRow = 701
    ActiveWindow.ScrollRow = 703
    ActiveWindow.ScrollRow = 705
    ActiveWindow.ScrollRow = 706
    ActiveWindow.ScrollRow = 709
    ActiveWindow.ScrollRow = 711
    ActiveWindow.ScrollRow = 713
    ActiveWindow.ScrollRow = 716
    ActiveWindow.ScrollRow = 717
    ActiveWindow.ScrollRow = 719
    ActiveWindow.ScrollRow = 721
    ActiveWindow.ScrollRow = 722
    ActiveWindow.ScrollRow = 724
    ActiveWindow.ScrollRow = 725
    ActiveWindow.ScrollRow = 727
    ActiveWindow.ScrollRow = 729
    ActiveWindow.ScrollRow = 730
    ActiveWindow.ScrollRow = 732
    ActiveWindow.ScrollRow = 733
    ActiveWindow.ScrollRow = 735
    ActiveWindow.ScrollRow = 737
    ActiveWindow.ScrollRow = 738
    ActiveWindow.ScrollRow = 740
    ActiveWindow.ScrollRow = 741
    ActiveWindow.ScrollRow = 743
    ActiveWindow.ScrollRow = 745
    ActiveWindow.ScrollRow = 746
    ActiveWindow.ScrollRow = 748
    ActiveWindow.ScrollRow = 749
    ActiveWindow.ScrollRow = 751
    ActiveWindow.ScrollRow = 753
    ActiveWindow.SmallScroll ToRight:=-1
    Sheets("MBS+PRO").Select
    ActiveWindow.SmallScroll ToRight:=0
    Sheets("Sheet1").Select
    ActiveCell.FormulaR1C1 = "1"
    Cells.Find(What:="pro-00154", After:=ActiveCell, LookIn:=xlFormulas2, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveWindow.ScrollRow = 2611
    ActiveWindow.ScrollRow = 2843
    ActiveWindow.ScrollRow = 4165
    ActiveWindow.ScrollRow = 4992
    ActiveWindow.ScrollRow = 6050
    ActiveWindow.ScrollRow = 6183
    ActiveWindow.ScrollRow = 6282
    ActiveWindow.ScrollRow = 6315
    ActiveWindow.ScrollRow = 6381
    ActiveWindow.ScrollRow = 7307
    ActiveWindow.ScrollRow = 8332
    ActiveWindow.ScrollRow = 8498
    ActiveWindow.ScrollRow = 8861
    ActiveWindow.ScrollRow = 9622
    ActiveWindow.ScrollRow = 10019
    ActiveWindow.ScrollRow = 10482
    ActiveWindow.ScrollRow = 10812
    ActiveWindow.ScrollRow = 10879
    ActiveWindow.ScrollRow = 10978
    ActiveWindow.ScrollRow = 11507
    ActiveWindow.ScrollRow = 11738
    ActiveWindow.ScrollRow = 12135
    ActiveWindow.ScrollRow = 12367
    ActiveWindow.ScrollRow = 12664
    ActiveWindow.ScrollRow = 13260
    ActiveWindow.ScrollRow = 13491
    ActiveWindow.ScrollRow = 13789
    ActiveWindow.ScrollRow = 14152
    ActiveWindow.ScrollRow = 14285
    ActiveWindow.ScrollRow = 14549
    ActiveWindow.ScrollRow = 15111
    ActiveWindow.ScrollRow = 15244
    ActiveWindow.ScrollRow = 15376
    ActiveWindow.ScrollRow = 15607
    ActiveWindow.ScrollRow = 15574
    ActiveWindow.ScrollRow = 15475
    ActiveWindow.ScrollRow = 15343
    ActiveWindow.ScrollRow = 15178
    ActiveWindow.ScrollRow = 15111
    ActiveWindow.ScrollRow = 16765
    ActiveWindow.ScrollRow = 17426
    ActiveWindow.ScrollRow = 17823
    ActiveWindow.ScrollRow = 18286
    ActiveWindow.ScrollRow = 18551
    ActiveWindow.ScrollRow = 19080
    ActiveWindow.ScrollRow = 19576
    ActiveWindow.ScrollRow = 20568
    ActiveWindow.ScrollRow = 20832
    ActiveWindow.ScrollRow = 21031
    ActiveWindow.ScrollRow = 20832
    ActiveWindow.ScrollRow = 20634
    ActiveWindow.ScrollRow = 20568
    ActiveWindow.ScrollRow = 21196
    ActiveWindow.ScrollRow = 21560
    ActiveWindow.ScrollRow = 22817
    ActiveWindow.ScrollRow = 23346
    ActiveWindow.ScrollRow = 23875
    ActiveWindow.ScrollRow = 24569
    ActiveWindow.ScrollRow = 24801
    ActiveWindow.ScrollRow = 25363
    ActiveWindow.ScrollRow = 25495
    ActiveWindow.ScrollRow = 25859
    ActiveWindow.ScrollRow = 26024
    ActiveWindow.ScrollRow = 26322
    ActiveWindow.ScrollRow = 26454
    ActiveWindow.ScrollRow = 26752
    ActiveWindow.ScrollRow = 26851
    ActiveWindow.ScrollRow = 26983
    ActiveWindow.ScrollRow = 27149
    ActiveWindow.ScrollRow = 27347
    ActiveWindow.ScrollRow = 27479
    ActiveWindow.ScrollRow = 27579
    ActiveWindow.ScrollRow = 27843
    ActiveWindow.ScrollRow = 28042
    ActiveWindow.ScrollRow = 28571
    ActiveWindow.ScrollRow = 28736
    ActiveWindow.ScrollRow = 29199
    ActiveWindow.ScrollRow = 29298
    ActiveWindow.ScrollRow = 29464
    ActiveWindow.ScrollRow = 29695
    ActiveWindow.ScrollRow = 29728
    ActiveWindow.ScrollRow = 29761
    ActiveWindow.ScrollRow = 29728
    ActiveWindow.ScrollRow = 29596
    ActiveWindow.ScrollRow = 28571
    ActiveWindow.ScrollRow = 28372
    ActiveWindow.ScrollRow = 28174
    ActiveWindow.ScrollRow = 28108
    ActiveWindow.ScrollRow = 27975
    ActiveWindow.ScrollRow = 27942
    ActiveWindow.ScrollRow = 27876
    ActiveWindow.ScrollRow = 27810
    ActiveWindow.ScrollRow = 27777
    ActiveWindow.ScrollRow = 27744
    ActiveWindow.ScrollRow = 27711
    ActiveWindow.ScrollRow = 27678
    ActiveWindow.ScrollRow = 27645
    ActiveWindow.ScrollRow = 27612
    ActiveWindow.ScrollRow = 27579
    ActiveWindow.ScrollRow = 27546
    ActiveWindow.ScrollRow = 27446
    ActiveWindow.ScrollRow = 27281
    ActiveWindow.ScrollRow = 27116
    ActiveWindow.ScrollRow = 26950
    ActiveWindow.ScrollRow = 26851
    ActiveWindow.ScrollRow = 26785
    ActiveWindow.ScrollRow = 26587
    ActiveWindow.ScrollRow = 26520
    ActiveWindow.ScrollRow = 26388
    ActiveWindow.ScrollRow = 26223
    ActiveWindow.ScrollRow = 26190
    ActiveWindow.ScrollRow = 26124
    ActiveWindow.ScrollRow = 25991
    ActiveWindow.ScrollRow = 25958
    ActiveWindow.ScrollRow = 25925
    ActiveWindow.ScrollRow = 25892
    ActiveWindow.ScrollRow = 25859
    ActiveWindow.ScrollRow = 25826
    ActiveWindow.ScrollRow = 25793
    ActiveWindow.ScrollRow = 25760
    ActiveWindow.ScrollRow = 25727
    ActiveWindow.ScrollRow = 25694
    ActiveWindow.ScrollRow = 25661
    ActiveWindow.ScrollRow = 25628
    ActiveWindow.ScrollRow = 25528
    ActiveWindow.ScrollRow = 25495
    ActiveWindow.ScrollRow = 25462
    ActiveWindow.ScrollRow = 25429
    ActiveWindow.ScrollRow = 25396
    ActiveWindow.ScrollRow = 25363
    ActiveWindow.ScrollRow = 25330
    ActiveWindow.ScrollRow = 25297
    ActiveWindow.ScrollRow = 25264
    ActiveWindow.ScrollRow = 25231
    ActiveWindow.ScrollRow = 25198
    ActiveWindow.ScrollRow = 25165
    ActiveWindow.ScrollRow = 25131
    ActiveWindow.ScrollRow = 25098
    ActiveWindow.ScrollRow = 25032
    ActiveWindow.ScrollRow = 24966
    ActiveWindow.ScrollRow = 24933
    ActiveWindow.ScrollRow = 24900
    ActiveWindow.ScrollRow = 24867
    ActiveWindow.ScrollRow = 24834
    ActiveWindow.ScrollRow = 24801
    ActiveWindow.ScrollRow = 24768
    ActiveWindow.ScrollRow = 24735
    ActiveWindow.ScrollRow = 24702
    ActiveWindow.ScrollRow = 24669
    ActiveWindow.ScrollRow = 24635
    ActiveWindow.ScrollRow = 24602
    ActiveWindow.ScrollRow = 24569
    ActiveWindow.ScrollRow = 24536
    ActiveWindow.ScrollRow = 24503
    ActiveWindow.ScrollRow = 24470
    ActiveWindow.ScrollRow = 24437
    ActiveWindow.ScrollRow = 24404
    ActiveWindow.ScrollRow = 24371
    ActiveWindow.ScrollRow = 24338
    ActiveWindow.ScrollRow = 24305
    ActiveWindow.ScrollRow = 24272
    ActiveWindow.ScrollRow = 24239
    ActiveWindow.ScrollRow = 24206
    ActiveWindow.ScrollRow = 24172
    ActiveWindow.ScrollRow = 24106
    ActiveWindow.ScrollRow = 24073
    ActiveWindow.ScrollRow = 24040
    ActiveWindow.ScrollRow = 24007
    ActiveWindow.ScrollRow = 23974
    ActiveWindow.ScrollRow = 23941
    ActiveWindow.ScrollRow = 23908
    ActiveWindow.ScrollRow = 23875
    ActiveWindow.ScrollRow = 23842
    ActiveWindow.ScrollRow = 23809
    ActiveWindow.ScrollRow = 23776
    ActiveWindow.ScrollRow = 23743
    ActiveWindow.ScrollRow = 23709
    ActiveWindow.ScrollRow = 23676
    ActiveWindow.ScrollRow = 23643
    ActiveWindow.ScrollRow = 23610
    ActiveWindow.ScrollRow = 23577
    ActiveWindow.ScrollRow = 23544
    ActiveWindow.ScrollRow = 23511
    ActiveWindow.ScrollRow = 23478
    ActiveWindow.ScrollRow = 23445
    ActiveWindow.ScrollRow = 23412
    ActiveWindow.ScrollRow = 23379
    ActiveWindow.ScrollRow = 23346
    ActiveWindow.ScrollRow = 23313
    ActiveWindow.ScrollRow = 23280
    ActiveWindow.ScrollRow = 23247
    ActiveWindow.ScrollRow = 23213
    ActiveWindow.ScrollRow = 23180
    ActiveWindow.ScrollRow = 23147
    ActiveWindow.ScrollRow = 23114
    ActiveWindow.ScrollRow = 23081
    ActiveWindow.ScrollRow = 23048
    ActiveWindow.ScrollRow = 23015
    ActiveWindow.ScrollRow = 22982
    ActiveWindow.ScrollRow = 22949
    ActiveWindow.ScrollRow = 22916
    ActiveWindow.ScrollRow = 22883
    ActiveWindow.ScrollRow = 22850
    ActiveWindow.ScrollRow = 22817
    ActiveWindow.ScrollRow = 22784
    ActiveWindow.ScrollRow = 22750
    ActiveWindow.ScrollRow = 22717
    ActiveWindow.ScrollRow = 22684
    ActiveWindow.ScrollRow = 22651
    ActiveWindow.ScrollRow = 22618
    ActiveWindow.ScrollRow = 22585
    ActiveWindow.ScrollRow = 22552
    ActiveWindow.ScrollRow = 22519
    ActiveWindow.ScrollRow = 22453
    ActiveWindow.ScrollRow = 22420
    ActiveWindow.ScrollRow = 22387
    ActiveWindow.ScrollRow = 22354
    ActiveWindow.ScrollRow = 22321
    ActiveWindow.ScrollRow = 22288
    ActiveWindow.ScrollRow = 22254
    ActiveWindow.ScrollRow = 22221
    ActiveWindow.ScrollRow = 22188
    ActiveWindow.ScrollRow = 22155
    ActiveWindow.ScrollRow = 22122
    ActiveWindow.ScrollRow = 22089
    ActiveWindow.ScrollRow = 22056
    ActiveWindow.ScrollRow = 22023
    ActiveWindow.ScrollRow = 21990
    ActiveWindow.ScrollRow = 21924
    ActiveWindow.ScrollRow = 21858
    ActiveWindow.ScrollRow = 21725
    ActiveWindow.ScrollRow = 21692
    ActiveWindow.ScrollRow = 21659
    ActiveWindow.ScrollRow = 21626
    ActiveWindow.ScrollRow = 21593
    ActiveWindow.ScrollRow = 21560
    ActiveWindow.ScrollRow = 21527
    ActiveWindow.ScrollRow = 21494
    ActiveWindow.ScrollRow = 21428
    ActiveWindow.ScrollRow = 21362
    ActiveWindow.ScrollRow = 21262
    ActiveWindow.ScrollRow = 21196
    ActiveWindow.ScrollRow = 21130
    ActiveWindow.ScrollRow = 21097
    ActiveWindow.ScrollRow = 21064
    ActiveWindow.ScrollRow = 21031
    ActiveWindow.ScrollRow = 20998
    ActiveWindow.ScrollRow = 21031
    ActiveWindow.ScrollRow = 21064
    ActiveWindow.ScrollRow = 21097
    Sheets("Sheet11").Select
    ActiveWindow.SelectedSheets.Delete
    ActiveWorkbook.Save
    Sheets("Sheet1").Select
    ActiveWorkbook.RefreshAll
    Application.Run "DataParseTest.xlsm!getData"
    Sheets("PRO").Select
    ActiveWindow.SmallScroll Down:=-1
    Range("D10").Select
    ActiveCell.FormulaR1C1 = "pro-00154"
    ActiveCell.Next.Activate
    ActiveCell.Next.Activate
    ActiveCell.FormulaR1C1 = "1"
    ActiveCell.Next.Activate
    ActiveWorkbook.RefreshAll
    Range("D10").Select
    ActiveCell.FormulaR1C1 = "PRO-00154"
    ActiveCell.Next.Activate
    ActiveWindow.Close
    Sheets("MBS+PRO").Select
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.SmallScroll Down:=111
    ActiveWindow.ScrollRow = 980
    ActiveWindow.ScrollRow = 984
    ActiveWindow.ScrollRow = 987
    ActiveWindow.ScrollRow = 988
    ActiveWindow.ScrollRow = 996
    ActiveWindow.ScrollRow = 1001
    ActiveWindow.ScrollRow = 1004
    ActiveWindow.ScrollRow = 1009
    ActiveWindow.ScrollRow = 1012
    ActiveWindow.ScrollRow = 1016
    ActiveWindow.ScrollRow = 1024
    ActiveWindow.ScrollRow = 1028
    ActiveWindow.ScrollRow = 1030
    ActiveWindow.ScrollRow = 1033
    ActiveWindow.ScrollRow = 1036
    ActiveWindow.ScrollRow = 1040
    ActiveWindow.ScrollRow = 1043
    ActiveWindow.ScrollRow = 1046
    ActiveWindow.ScrollRow = 1049
    ActiveWindow.ScrollRow = 1057
    ActiveWindow.ScrollRow = 1061
    ActiveWindow.ScrollRow = 1064
    ActiveWindow.ScrollRow = 1065
    ActiveWindow.ScrollRow = 1069
    ActiveWindow.ScrollRow = 1072
    ActiveWindow.ScrollRow = 1075
    ActiveWindow.ScrollRow = 1078
    ActiveWindow.ScrollRow = 1081
    ActiveWindow.ScrollRow = 1085
    ActiveWindow.ScrollRow = 1088
    ActiveWindow.ScrollRow = 1091
    ActiveWindow.ScrollRow = 1094
    ActiveWindow.ScrollRow = 1097
    ActiveWindow.ScrollRow = 1099
    ActiveWindow.ScrollRow = 1102
    ActiveWindow.ScrollRow = 1104
    ActiveWindow.ScrollRow = 1105
    ActiveWindow.ScrollRow = 1107
    ActiveWindow.ScrollRow = 1109
    ActiveWindow.ScrollRow = 1110
    ActiveWindow.ScrollRow = 1112
    ActiveWindow.ScrollRow = 1113
    ActiveWindow.ScrollRow = 1115
    ActiveWindow.ScrollRow = 1117
    ActiveWindow.ScrollRow = 1120
    ActiveWindow.ScrollRow = 1121
    ActiveWindow.ScrollRow = 1123
    ActiveWindow.ScrollRow = 1125
    ActiveWindow.ScrollRow = 1126
    ActiveWindow.ScrollRow = 1128
    ActiveWindow.ScrollRow = 1129
    ActiveWindow.ScrollRow = 1133
    ActiveWindow.ScrollRow = 1136
    ActiveWindow.ScrollRow = 1138
    ActiveWindow.ScrollRow = 1139
    ActiveWindow.ScrollRow = 1142
    ActiveWindow.ScrollRow = 1144
    ActiveWindow.ScrollRow = 1146
    ActiveWindow.ScrollRow = 1147
    ActiveWindow.ScrollRow = 1149
    ActiveWindow.ScrollRow = 1152
    ActiveWindow.ScrollRow = 1154
    ActiveWindow.ScrollRow = 1155
    ActiveWindow.ScrollRow = 1157
    ActiveWindow.ScrollRow = 1158
    ActiveWindow.ScrollRow = 1160
    ActiveWindow.ScrollRow = 1162
    ActiveWindow.ScrollRow = 1163
    ActiveWindow.ScrollRow = 1165
    ActiveWindow.ScrollRow = 1166
    ActiveWindow.ScrollRow = 1170
    ActiveWindow.ScrollRow = 1171
    ActiveWindow.ScrollRow = 1173
    ActiveWindow.ScrollRow = 1178
    ActiveWindow.ScrollRow = 1179
    ActiveWindow.ScrollRow = 1181
    ActiveWindow.ScrollRow = 1182
    ActiveWindow.ScrollRow = 1184
    ActiveWindow.ScrollRow = 1186
    ActiveWindow.ScrollRow = 1189
    ActiveWindow.ScrollRow = 1192
    ActiveWindow.ScrollRow = 1194
    ActiveWindow.ScrollRow = 1195
    ActiveWindow.ScrollRow = 1197
    ActiveWindow.ScrollRow = 1198
    ActiveWindow.ScrollRow = 1200
    ActiveWindow.ScrollRow = 1202
    ActiveWindow.ScrollRow = 1203
    ActiveWindow.ScrollRow = 1205
    ActiveWindow.ScrollRow = 1206
    ActiveWindow.ScrollRow = 1208
    ActiveWindow.ScrollRow = 1210
    ActiveWindow.ScrollRow = 1211
    ActiveWindow.ScrollRow = 1213
    ActiveWindow.ScrollRow = 1214
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.SmallScroll Down:=174
    ActiveWindow.ScrollRow = 1386
    ActiveWindow.ScrollRow = 1372
    ActiveWindow.ScrollRow = 1352
    ActiveWindow.ScrollRow = 1335
    ActiveWindow.ScrollRow = 1325
    ActiveWindow.ScrollRow = 1296
    ActiveWindow.ScrollRow = 1176
    ActiveWindow.ScrollRow = 1128
    ActiveWindow.ScrollRow = 1099
    ActiveWindow.ScrollRow = 1035
    ActiveWindow.ScrollRow = 995
    ActiveWindow.ScrollRow = 963
    ActiveWindow.ScrollRow = 894
    ActiveWindow.ScrollRow = 871
    ActiveWindow.ScrollRow = 839
    ActiveWindow.ScrollRow = 788
    ActiveWindow.ScrollRow = 757
    ActiveWindow.ScrollRow = 735
    ActiveWindow.ScrollRow = 690
    ActiveWindow.ScrollRow = 653
    ActiveWindow.ScrollRow = 628
    ActiveWindow.ScrollRow = 568
    ActiveWindow.ScrollRow = 536
    ActiveWindow.ScrollRow = 507
    ActiveWindow.ScrollRow = 485
    ActiveWindow.ScrollRow = 453
    ActiveWindow.ScrollRow = 424
    ActiveWindow.ScrollRow = 368
    ActiveWindow.ScrollRow = 342
    ActiveWindow.ScrollRow = 291
    ActiveWindow.ScrollRow = 252
    ActiveWindow.ScrollRow = 227
    ActiveWindow.ScrollRow = 209
    ActiveWindow.ScrollRow = 193
    ActiveWindow.ScrollRow = 171
    ActiveWindow.ScrollRow = 159
    ActiveWindow.ScrollRow = 137
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 102
    ActiveWindow.ScrollRow = 92
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SmallScroll Down:=0
    ActiveWindow.ScrollColumn = 2
    Columns("B:B").ColumnWidth = 145.73
    ActiveWindow.SmallScroll Down:=48
    Range("B62").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A47:D60").Select
    Range("B55").Activate
    ActiveWindow.SmallScroll Down:=-9
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A32:D45").Select
    Range("B41").Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveWindow.SmallScroll Down:=12
    Range("B77").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    Range("B52").Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B64").Select
    Columns("B:B").ColumnWidth = 79.45
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SmallScroll Down:=-51
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SmallScroll Down:=48
    Range("B75").Select
    ActiveWindow.SmallScroll Down:=6
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 74
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 102
    ActiveWindow.ScrollRow = 107
    ActiveWindow.ScrollRow = 119
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 139
    ActiveWindow.ScrollRow = 144
    ActiveWindow.ScrollRow = 154
    ActiveWindow.ScrollRow = 160
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 167
    ActiveWindow.ScrollRow = 169
    ActiveWindow.ScrollRow = 171
    ActiveWindow.ScrollRow = 172
    ActiveWindow.ScrollRow = 174
    ActiveWindow.ScrollRow = 176
    ActiveWindow.ScrollRow = 179
    ActiveWindow.ScrollRow = 183
    ActiveWindow.ScrollRow = 185
    ActiveWindow.ScrollRow = 188
    ActiveWindow.ScrollRow = 191
    ActiveWindow.ScrollRow = 194
    ActiveWindow.ScrollRow = 195
    ActiveWindow.ScrollRow = 199
    ActiveWindow.ScrollRow = 204
    ActiveWindow.ScrollRow = 210
    ActiveWindow.ScrollRow = 220
    ActiveWindow.ScrollRow = 225
    ActiveWindow.ScrollRow = 227
    ActiveWindow.ScrollRow = 229
    ActiveWindow.ScrollRow = 231
    ActiveWindow.ScrollRow = 233
    ActiveWindow.ScrollRow = 234
    ActiveWindow.ScrollRow = 238
    ActiveWindow.ScrollRow = 241
    ActiveWindow.ScrollRow = 246
    ActiveWindow.ScrollRow = 256
    ActiveWindow.ScrollRow = 258
    ActiveWindow.ScrollRow = 259
    ActiveWindow.ScrollRow = 260
    ActiveWindow.ScrollRow = 262
    ActiveWindow.ScrollRow = 264
    ActiveWindow.ScrollRow = 265
    ActiveWindow.ScrollRow = 267
    ActiveWindow.ScrollRow = 269
    ActiveWindow.ScrollRow = 270
    ActiveWindow.ScrollRow = 272
    ActiveWindow.ScrollRow = 273
    ActiveWindow.ScrollRow = 274
    ActiveWindow.ScrollRow = 276
    ActiveWindow.ScrollRow = 278
    ActiveWindow.ScrollRow = 279
    ActiveWindow.ScrollRow = 281
    ActiveWindow.ScrollRow = 283
    ActiveWindow.ScrollRow = 284
    ActiveWindow.ScrollRow = 286
    Range("B308").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B305").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B299").Select
    Columns("B:B").EntireColumn.AutoFit
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    Range("A286").Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B295").Select
    Columns("B:B").EntireColumn.AutoFit
    ActiveWindow.SmallScroll Down:=9
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.SmallScroll Down:=99
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SmallScroll Down:=-21
    Sheets("Sheet1").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("CHK").Select
    Columns("C:C").EntireColumn.AutoFit
    ActiveWindow.SmallScroll Down:=12
    Range("A45").Select
    ActiveCell.FormulaR1C1 = "CHK-00059"
    Sheets("MBK").Select
    ActiveWindow.SmallScroll Down:=15
    Sheets("PRO").Select
    ActiveWindow.SmallScroll Down:=3
    ActiveWorkbook.Save
End Function
