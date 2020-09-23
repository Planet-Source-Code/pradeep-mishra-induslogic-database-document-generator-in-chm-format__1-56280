Attribute VB_Name = "Util"
Public Function MakeCompatibleFileName(strFilename As String) As String
    Dim tempFileName As String
    tempFileName = strFilename
    
    tempFileName = Replace(tempFileName, "\", "-")
    tempFileName = Replace(tempFileName, Chr(34), "-")
    tempFileName = Replace(tempFileName, "/", "-")
    tempFileName = Replace(tempFileName, "*", "-")
    tempFileName = Replace(tempFileName, "?", "-")
    tempFileName = Replace(tempFileName, "<", "-")
    tempFileName = Replace(tempFileName, ">", "-")
    tempFileName = Replace(tempFileName, "|", "-")

    MakeCompatibleFileName = tempFileName
End Function
