Attribute VB_Name = "FileFromPath"
' https://stackoverflow.com/a/28632739
Function GetFileNameFromPath(strFullPath As String) As String
  GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function

