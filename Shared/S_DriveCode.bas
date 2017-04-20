Attribute VB_Name = "S_DriveCode"

'============== GENERIC FUNCTIONS

Public Function DirExists(ByVal sDirName As String) As Boolean
On Error Resume Next
DirExists = (GetAttr(sDirName) And vbDirectory) = vbDirectory
Err.Clear
End Function


Public Function FileExists(ByVal sPathName As String) As Boolean
On Error Resume Next
FileExists = (GetAttr(sPathName) And vbNormal) = vbNormal
Err.Clear
End Function

