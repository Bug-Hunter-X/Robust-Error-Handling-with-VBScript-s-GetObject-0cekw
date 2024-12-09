Function GetObject(path)
  On Error Resume Next
  Set obj = GetObject(path)
  If Err.Number <> 0 Then
    Err.Clear
    Set obj = Nothing
  End If
  Set GetObject = obj
End Function

' Example usage
Set myExcel = GetObject("C:\\path\\to\\your\\excel.xls")
if myExcel is nothing then
    msgbox "Could not open excel file."
    WScript.Quit
end if

MsgBox myExcel.Worksheets(1).Cells(1,1).Value