Attribute VB_Name = "Utils"
Option Explicit
'
' Some usefull functions
'

Public Function SheetExists(SheetName As String, Optional wb As EXCEL.Workbook) ' https://stackoverflow.com/questions/6688131/test-or-check-if-sheet-exists'
'
' Finds if sheet witch name: SheetName exits
'
   Dim s As EXCEL.Worksheet
   If wb Is Nothing Then Set wb = ThisWorkbook
   On Error Resume Next
   Set s = wb.Sheets(SheetName)
   On Error GoTo 0
   SheetExists = Not s Is Nothing
End Function

Public Function is_integer(user_input As Variant) As Boolean
'
' Check if user_input is reade for lossless conversion to integer
'
    is_integer = False
    If IsNumeric(user_input) Then
        If CStr(CLng(user_input)) = CStr(user_input) Then
          is_integer = True
        End If
    End If
End Function
