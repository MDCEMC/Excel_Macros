Attribute VB_Name = "Module18"
Sub SortRequestHiToLo()
Attribute SortRequestHiToLo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'
ThisWorkbook.Activate
'
    ActiveSheet.Unprotect
    ActiveWorkbook.Worksheets("Request DB").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Request DB").AutoFilter.Sort.SortFields.add Key:= _
        Range("A3:A582"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Request DB").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
End Sub
Sub SortRequestLoToHi()
'
' Macro2 Macro
'
ThisWorkbook.Activate

'
    ActiveSheet.Unprotect
    ActiveWorkbook.Worksheets("Request DB").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Request DB").AutoFilter.Sort.SortFields.add Key:= _
        Range("A3:A582"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Request DB").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
End Sub


