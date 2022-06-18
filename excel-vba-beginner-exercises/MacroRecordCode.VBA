Option Explicit
Sub SortDeptLast()
'
' SortDeptLast Macro
'

'
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Sort & Filter").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sort & Filter").Sort.SortFields.Add2 Key:=Range( _
        "C4:C701"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Sort & Filter").Sort.SortFields.Add2 Key:=Range( _
        "A4:A701"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sort & Filter").Sort
        .SetRange Range("A3:H701")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub FilterOperationFulltime()
'
' FilterOperationFulltime Macro
'

'
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    ActiveWindow.SmallScroll Down:=-24
    ActiveWindow.ScrollRow = 656
    ActiveWindow.ScrollRow = 655
    ActiveWindow.ScrollRow = 648
    ActiveWindow.ScrollRow = 632
    ActiveWindow.ScrollRow = 625
    ActiveWindow.ScrollRow = 619
    ActiveWindow.ScrollRow = 607
    ActiveWindow.ScrollRow = 600
    ActiveWindow.ScrollRow = 587
    ActiveWindow.ScrollRow = 580
    ActiveWindow.ScrollRow = 574
    ActiveWindow.ScrollRow = 561
    ActiveWindow.ScrollRow = 557
    ActiveWindow.ScrollRow = 546
    ActiveWindow.ScrollRow = 543
    ActiveWindow.ScrollRow = 526
    ActiveWindow.ScrollRow = 519
    ActiveWindow.ScrollRow = 514
    ActiveWindow.ScrollRow = 491
    ActiveWindow.ScrollRow = 482
    ActiveWindow.ScrollRow = 468
    ActiveWindow.ScrollRow = 459
    ActiveWindow.ScrollRow = 451
    ActiveWindow.ScrollRow = 439
    ActiveWindow.ScrollRow = 310
    ActiveWindow.ScrollRow = 290
    ActiveWindow.ScrollRow = 283
    ActiveWindow.ScrollRow = 261
    ActiveWindow.ScrollRow = 255
    ActiveWindow.ScrollRow = 244
    ActiveWindow.ScrollRow = 240
    ActiveWindow.ScrollRow = 224
    ActiveWindow.ScrollRow = 219
    ActiveWindow.ScrollRow = 209
    ActiveWindow.ScrollRow = 206
    ActiveWindow.ScrollRow = 199
    ActiveWindow.ScrollRow = 193
    ActiveWindow.ScrollRow = 185
    ActiveWindow.ScrollRow = 180
    ActiveWindow.ScrollRow = 173
    ActiveWindow.ScrollRow = 167
    ActiveWindow.ScrollRow = 160
    ActiveWindow.ScrollRow = 158
    ActiveWindow.ScrollRow = 109
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 100
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 74
    ActiveWindow.ScrollRow = 71
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 55
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 1
    ActiveSheet.Range("$A$3:$H$701").AutoFilter Field:=3, Criteria1:= _
        "Operations"
    ActiveSheet.Range("$A$3:$H$701").AutoFilter Field:=7, Criteria1:= _
        "=Full Time", Operator:=xlOr, Criteria2:="=Full-Time"
End Sub
Sub ClearFilter()
'
' ClearFilter Macro
'

'
    ActiveSheet.ShowAllData
End Sub
