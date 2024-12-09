Function GetObjectSafe(path)
  Dim obj, errNum
  On Error Resume Next
  Set obj = GetObject(path)
  errNum = Err.Number
  On Error GoTo 0
  If errNum <> 0 Then
    ' Handle errors appropriately, e.g., log the error or display a message.
    MsgBox "Error accessing file: " & path & ". Error number: " & errNum, vbCritical
    Set obj = Nothing
  End If
  Set GetObjectSafe = obj
End Function

' Example usage
Set myExcel = GetObjectSafe("C:\\path\\to\\your\\excel.xls")
if myExcel is nothing then
    msgbox "Could not open excel file."
    WScript.Quit
end if

MsgBox myExcel.Worksheets(1).Cells(1,1).Value