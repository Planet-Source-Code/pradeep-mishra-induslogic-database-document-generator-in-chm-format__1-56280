Attribute VB_Name = "HHProj"
Option Explicit

Public sHHHead As String
Public sHHFiles As String
Public HHFileName As String

Public Const HHHead = "[OPTIONS]" & vbCrLf & _
    "Auto Index = Yes" & vbCrLf & _
    "Compatibility = 1.1 Or later" & vbCrLf & _
    "Compiled File = <%COMPILEDFILE%>" & vbCrLf & _
    "Contents File = TOC.hhc" & vbCrLf & _
    "Default topic = db_details.htm" & vbCrLf & _
    "Display compile progress=No" & vbCrLf & _
    "Full-text search=Yes" & vbCrLf & _
    "Index File = Index.hhk" & vbCrLf & _
    "Language=0x409 English (United States)" & vbCrLf & _
    "Title=Database documentation for <%TITLE%>" & vbCrLf

Public Const HHFiles = "[FILES]" & vbCrLf


Public Const HTMLCompiler = "C:\Program Files\HTML Help Workshop\hhc.exe"

