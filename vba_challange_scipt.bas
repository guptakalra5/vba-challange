Attribute VB_Name = "Module1"
Option Private Module
Sub getData()
Attribute getData.VB_ProcData.VB_Invoke_Func = " \n14"
Dim dic As Object, lr&, r As Range, cf As Range
With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
End With
Columns("i:m").Delete
Set dic = CreateObject("scripting.dictionary")
lr = Cells(Rows.Count, 1).End(xlUp).Row
For Each r In Range("a2:a" & lr)
    If Not dic.exists(r.Value) Then
        dic.Add r.Value, r.Offset(, 1).Value & ":" & r.Offset(, 1).Value & ":" & r.Offset(, 2).Value & ":" & r.Offset(, 5).Value & ":" & r.Offset(, 6).Value
    Else
        dic(r.Value) = WorksheetFunction.Min(r.Offset(, 1).Value, Split(dic(r.Value), ":")(0)) & ":" & WorksheetFunction.Max(r.Offset(, 1).Value, Split(dic(r.Value), ":")(1)) & ":" & Split(dic(r.Value), ":")(2) & ":" & r.Offset(, 5).Value & ":" & r.Offset(, 6).Value + Split(dic(r.Value), ":")(4)
    End If
Next r

Range("i2").Resize(dic.Count, 2).Value = Application.Transpose(Array(dic.keys, dic.items))
Set dic = Nothing
Columns("J:J").TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=":", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1)), TrailingMinusNumbers:=True
lr = Cells(Rows.Count, 10).End(xlUp).Row
Range("j2:j" & lr).Value = "=M2-L2"
Range("k2:k" & lr).Value = "=J2/L2"
Range("j2:k" & lr).Value = Range("j2:k" & lr).Value
Range("k2:k" & lr).NumberFormat = "0.00%"
Columns("L:M").Delete
Range("i1").Resize(, 4).Value = Array("Ticker", "Yearly Changed", "Percent Changed", "Total Volume")

Set cf = Range("j2:j" & lr)
cf.FormatConditions.Delete
cf.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
        Formula1:="=0"
cf.FormatConditions(1).Interior.Color = vbGreen
'Add second rule
cf.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
cf.FormatConditions(2).Interior.Color = vbRed

Set cf = Nothing
With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
End With
End Sub


