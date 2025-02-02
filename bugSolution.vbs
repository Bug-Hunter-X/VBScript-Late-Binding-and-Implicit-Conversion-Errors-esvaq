Option Explicit

' Function to demonstrate safe object handling
Function SafeGetObject(obj, methodName)
  On Error Resume Next
  Set SafeGetObject = obj.methodName
  If Err.Number <> 0 Then
    Err.Clear
    SafeGetObject = Null
    ' Handle the error appropriately (log, fallback, etc.)
  End If
  On Error GoTo 0
End Function

'Demonstrates explicit type checking
Sub CheckTypes()
  Dim myVar, myNum
  myVar = "123"
  If IsNumeric(myVar) Then
    myNum = CInt(myVar) 'Explicit conversion
    MsgBox "Number: " & myNum
  Else
    MsgBox "Not a Number"
  End If
End Sub

'Safe Object Handling Example
Sub SafeObjectExample()
  Dim objFSO, folderPath
  folderPath = "C:\MyFolder"

  Set objFSO = CreateObject("Scripting.FileSystemObject")

  'Instead of objFSO.GetFolder(folderPath) which might fail
  Set objFolder = SafeGetObject(objFSO, "GetFolder(folderPath)")
  If Not objFolder Is Nothing Then
    MsgBox "Folder exists"
  Else
    MsgBox "Folder does not exist"
  End If

  Set objFSO = Nothing
End Sub

CheckTypes()
SafeObjectExample()